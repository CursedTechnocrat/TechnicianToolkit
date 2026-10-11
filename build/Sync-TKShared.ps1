# Sync-TKShared.ps1 - Regenerates each tool's TK-SHARED block from the shared helper source.
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# SPDX-License-Identifier: GPL-3.0-or-later

<#
.SYNOPSIS
    Regenerates the TK-SHARED block in every tool from shared/TechnicianToolkit.Shared.ps1.

.DESCRIPTION
    The toolkit ships no shared module. Every tool carries, between
    '#region TK-SHARED' and '#endregion TK-SHARED', a generated copy of the
    shared helpers it calls -- plus the helpers those call, and any script-level
    state they read -- so a single .ps1 runs on its own with nothing beside it
    and nothing to download.

    The source of truth is shared/TechnicianToolkit.Shared.ps1. Edit it, then
    run this script to rewrite every block that no longer matches. Blocks are
    never edited by hand; the Pester suite dot-sources this script and fails on
    any block that differs from what it would write.

    Each tool keeps its own line endings, so a CRLF checkout stays CRLF.

.USAGE
    PS C:\> .\build\Sync-TKShared.ps1           # Rewrite every stale block
    PS C:\> .\build\Sync-TKShared.ps1 -Check    # List stale blocks; exit 1 if any

.NOTES
    Dot-sourcing the script defines its functions without running it, which is
    how the test suite uses it.
#>

[CmdletBinding()]
param(
    [switch]$Check
)

$TKSharedBegin = '#region TK-SHARED'
$TKSharedEnd   = '#endregion TK-SHARED'

function Get-TKSharedSource {
    <#
    .SYNOPSIS
        Parses the shared source into its top-level items, in source order.
    .DESCRIPTION
        Each item is a function or a top-level statement, with the comment lines
        directly above it. Functions record which commands they call so the
        closure of what a tool needs can be computed; '$script:' assignments
        record the variable so they travel only with functions that read it.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
    if ($errors.Count -gt 0) { throw "$Path does not parse: $($errors[0].Message)" }

    $lines = ([IO.File]::ReadAllText($Path) -replace "`r`n", "`n") -split "`n"

    foreach ($statement in $ast.EndBlock.Statements) {
        $first = $statement.Extent.StartLineNumber - 1
        while ($first -gt 0 -and $lines[$first - 1] -match '^\s*#' -and
               $lines[$first - 1] -notmatch '^\s*#(region|endregion)\b' -and
               $lines[$first - 1] -notmatch '#>\s*$') {
            $first--
        }
        $lead = if ($first -lt $statement.Extent.StartLineNumber - 1) {
            ($lines[$first..($statement.Extent.StartLineNumber - 2)] -join "`n") + "`n"
        } else { '' }

        $item = [PSCustomObject]@{
            Name      = $null
            Text      = $lead + ($statement.Extent.Text -replace "`r`n", "`n")
            Calls     = @()
            ScriptVar = $null
        }

        if ($statement -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
            $item.Name  = $statement.Name
            $item.Calls = @($statement.FindAll({
                param($node) $node -is [System.Management.Automation.Language.CommandAst]
            }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Select-Object -Unique)
        }
        elseif ($statement -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $statement.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $statement.Left.VariablePath.IsScript) {
            $item.ScriptVar = $statement.Left.VariablePath.UserPath -replace '^script:', ''
        }

        $item
    }
}

function Split-TKSharedText {
    <#
    .SYNOPSIS
        Splits LF-normalised tool text around its TK-SHARED block.
        Returns $null when the text has no block, which is how a script that is
        not a tool (a forwarding stub) is told apart.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $begin = [regex]::Match($Text, '(?m)^' + [regex]::Escape($TKSharedBegin) + '\b[^\n]*\n')
    if (-not $begin.Success) { return $null }
    $end = [regex]::Match($Text.Substring($begin.Index), '(?m)^' + [regex]::Escape($TKSharedEnd) + '\b[^\n]*(\n|$)')
    if (-not $end.Success) { throw "TK-SHARED block opens but never closes." }

    $stop = $begin.Index + $end.Index + $end.Length
    [PSCustomObject]@{
        Before = $Text.Substring(0, $begin.Index)
        Block  = $Text.Substring($begin.Index, $stop - $begin.Index)
        After  = $Text.Substring($stop)
    }
}

function Get-TKSharedItem {
    <#
    .SYNOPSIS
        Returns the shared items a tool needs, in source order: every function it
        calls outside its block, everything those call in turn, the script-level
        state they read, and every other top-level statement.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ToolCode,
        [Parameter(Mandatory)][object[]]$Source
    )

    $byName = @{}
    foreach ($item in $Source) { if ($item.Name) { $byName[$item.Name] = $item } }

    $toolAst = [System.Management.Automation.Language.Parser]::ParseInput($ToolCode, [ref]$null, [ref]$null)
    $pending = [System.Collections.Generic.Queue[string]]::new()
    $toolAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) |
        ForEach-Object { $_.GetCommandName() } |
        Where-Object { $_ -and $byName.ContainsKey($_) } |
        ForEach-Object { $pending.Enqueue($_) }

    $needed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    while ($pending.Count -gt 0) {
        $name = $pending.Dequeue()
        if (-not $needed.Add($name)) { continue }
        foreach ($call in $byName[$name].Calls) {
            if ($byName.ContainsKey($call)) { $pending.Enqueue($call) }
        }
    }

    $neededText = ($Source | Where-Object { $_.Name -and $needed.Contains($_.Name) } | ForEach-Object { $_.Text }) -join "`n"

    foreach ($item in $Source) {
        if ($item.Name) {
            if ($needed.Contains($item.Name)) { $item }
        }
        elseif ($item.ScriptVar) {
            if ($neededText -match ('\$script:' + [regex]::Escape($item.ScriptVar) + '\b')) { $item }
        }
        else { $item }
    }
}

function Update-TKSharedText {
    <#
    .SYNOPSIS
        Returns the tool's full text with its TK-SHARED block regenerated, in the
        tool's own line endings. Returns $null for a script with no block.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory)][object[]]$Source
    )

    $newline = if ($Text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $parts = Split-TKSharedText -Text ($Text -replace "`r`n", "`n")
    if ($null -eq $parts) { return $null }

    $items = @(Get-TKSharedItem -ToolCode ($parts.Before + $parts.After) -Source $Source)
    $block = @(
        $TKSharedBegin
        '# Generated from shared/TechnicianToolkit.Shared.ps1 by build/Sync-TKShared.ps1. Do not edit'
        '# this block: edit the source and re-run the script. The Pester suite fails on a stale block.'
        ''
        ($items | ForEach-Object { $_.Text }) -join "`n`n"
        $TKSharedEnd
    ) -join "`n"

    ($parts.Before + $block + "`n" + $parts.After) -replace "`n", $newline
}

if ($MyInvocation.InvocationName -ne '.') {
    $root   = Split-Path -Parent $PSScriptRoot
    $source = @(Get-TKSharedSource -Path (Join-Path (Join-Path $root 'shared') 'TechnicianToolkit.Shared.ps1'))

    $stale = @(foreach ($file in Get-ChildItem -Path $root -Filter '*.ps1' -File | Sort-Object Name) {
        $text     = [IO.File]::ReadAllText($file.FullName)
        $expected = Update-TKSharedText -Text $text -Source $source
        if ($null -eq $expected -or $expected -ceq $text) { continue }
        if (-not $Check) {
            [IO.File]::WriteAllText($file.FullName, $expected, [Text.UTF8Encoding]::new($true))
        }
        $file.Name
    })

    if ($stale.Count -eq 0) {
        Write-Host "  [+] Every TK-SHARED block matches shared/TechnicianToolkit.Shared.ps1." -ForegroundColor Green
    }
    elseif ($Check) {
        Write-Host "  [!!] $($stale.Count) tool(s) carry a stale TK-SHARED block - run build/Sync-TKShared.ps1:" -ForegroundColor Red
        $stale | ForEach-Object { Write-Host "       $_" -ForegroundColor Red }
        exit 1
    }
    else {
        Write-Host "  [+] Regenerated the TK-SHARED block in $($stale.Count) tool(s):" -ForegroundColor Green
        $stale | ForEach-Object { Write-Host "       $_" -ForegroundColor Gray }
    }
}
