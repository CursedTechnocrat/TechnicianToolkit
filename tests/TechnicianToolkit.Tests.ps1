# TechnicianToolkit.Tests.ps1 - Pester tests for the TechnicianToolkit shared module (TechnicianToolkit.psm1).
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

#Requires -Modules Pester
<#
.SYNOPSIS
    Pester tests for the TechnicianToolkit shared module (TechnicianToolkit.psm1).
    Tests cover the pure utility functions that do not require admin rights or
    live Windows APIs, so they can run in CI without elevated privileges.
#>

BeforeAll {
    $ModulePath = Join-Path $PSScriptRoot '..\TechnicianToolkit.psm1'
    Import-Module $ModulePath -Force

    # Tool helper extraction for the NECROPSY / RAVEN / WARD / GRIFFIN blocks.
    # Each tool launches its main flow on import, so its tables and pure
    # helpers are pulled out of the AST and evaluated on their own.
    function Get-ToolAst {
        param([string]$FileName)
        $errs = $null
        return [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path (Join-Path $PSScriptRoot '..') $FileName), [ref]$null, [ref]$errs)
    }

    function Get-ToolAssignmentValue {
        param($Ast, [string]$VarName)
        $assign = $Ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $n.Left.VariablePath.UserPath -eq $VarName
        }, $true) | Select-Object -First 1
        if (-not $assign) { return $null }
        return & ([scriptblock]::Create($assign.Right.Extent.Text))
    }

    function Get-ToolFunctionText {
        param($Ast, [string]$FuncName)
        $fn = $Ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FuncName
        }, $true) | Select-Object -First 1
        if (-not $fn) { return $null }
        return $fn.Extent.Text
    }
}

# Directories that live in the working tree but are not repository source.
# .git is version-control internals; .claude is local editor/agent tooling that
# happens to be written in PowerShell, so the recursive gates below would other-
# wise hold hook scripts to the toolkit's own header, BOM and naming rules and
# fail on every one. Defined at file scope because the file enumerations run
# during Pester's discovery phase, before any BeforeAll body executes.
$NonSourceDir = '{0}(\.git|\.claude){0}' -f [regex]::Escape([string][IO.Path]::DirectorySeparatorChar)

# Forwarding stubs left at the old filenames of the tools renamed in 6.0
# (old -> new). They carry no tool logic -- no param() block, so `@args` passes
# every argument through untouched -- so they are exempt from the tool-shape
# gates below; the 'Renamed-tool forwarding stubs' block checks them instead.
# Defined at file scope because the -ForEach case lists are built during
# discovery.
$RenamedToolStubs = [ordered]@{
    'restoration.ps1' = 'whetstone.ps1'
    'suture.ps1'      = 'solder.ps1'
    'threshold.ps1'   = 'fathom.ps1'
    'gargoyle.ps1'    = 'vigil.ps1'
    'pyre.ps1'        = 'hourglass.ps1'
    'cipher.ps1'      = 'wyrm.ps1'
    'sigil.ps1'       = 'basilisk.ps1'
    'citadel.ps1'     = 'sphinx.ps1'
    'artifact.ps1'    = 'phoenix.ps1'
    'paladin.ps1'     = 'griffin.ps1'
    'herald.ps1'      = 'argus.ps1'
    'catacomb.ps1'    = 'minotaur.ps1'
    'beacon.ps1'      = 'wisp.ps1'
    'shade.ps1'       = 'emissary.ps1'
    'oath.ps1'        = 'lodestar.ps1'
    'talisman.ps1'    = 'zenith.ps1'
    'reliquary.ps1'   = 'almanac.ps1'
    'golem.ps1'       = 'orbit.ps1'
    'wraith.ps1'      = 'eclipse.ps1'
    'conclave.ps1'    = 'asterism.ps1'
    'grove.ps1'       = 'cumulus.ps1'
    'tendril.ps1'     = 'orrery.ps1'
    'rampart.ps1'     = 'halo.ps1'
    'archive.ps1'     = 'embalm.ps1'
    'tether.ps1'      = 'phylactery.ps1'
}

# ─────────────────────────────────────────────────────────────────────────────
# EscHtml
# ─────────────────────────────────────────────────────────────────────────────
Describe 'EscHtml' {
    It 'escapes ampersands' {
        EscHtml 'a & b' | Should -Be 'a &amp; b'
    }
    It 'escapes less-than' {
        EscHtml '<script>' | Should -Be '&lt;script&gt;'
    }
    It 'escapes double quotes' {
        EscHtml '"hello"' | Should -Be '&quot;hello&quot;'
    }
    It 'returns empty string for null input' {
        EscHtml $null | Should -Be ''
    }
    It 'returns empty string for empty input' {
        EscHtml '' | Should -Be ''
    }
    It 'passes through plain text unchanged' {
        EscHtml 'hello world' | Should -Be 'hello world'
    }
    It 'handles multiple special chars in one string' {
        EscHtml '<b>me & you</b>' | Should -Be '&lt;b&gt;me &amp; you&lt;/b&gt;'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Format-Bytes
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Format-Bytes' {
    It 'returns bytes for values under 1 KB' {
        Format-Bytes 512 | Should -Be '512 B'
    }
    It 'returns KB for values under 1 MB' {
        Format-Bytes 2048 | Should -Be '2.00 KB'
    }
    It 'returns MB for values under 1 GB' {
        Format-Bytes (5 * 1MB) | Should -Be '5.00 MB'
    }
    It 'returns GB for values under 1 TB' {
        Format-Bytes (3 * 1GB) | Should -Be '3.00 GB'
    }
    It 'returns TB for values at or above 1 TB' {
        Format-Bytes (2 * 1TB) | Should -Be '2.00 TB'
    }
    It 'handles zero bytes' {
        Format-Bytes 0 | Should -Be '0 B'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Get-TKHtmlHead / Get-TKHtmlFoot — structural smoke tests
# ─────────────────────────────────────────────────────────────────────────────
Describe 'HTML report helpers' {
    Context 'Get-TKHtmlHead' {
        It 'returns a well-formed HTML preamble' {
            $html = Get-TKHtmlHead -Title 'Unit Test' -ScriptName 'T.E.S.T.'
            $html | Should -Match '^<!DOCTYPE html>'
            $html | Should -Match '<html lang="en">'
            $html | Should -Match '<title>Unit Test</title>'
            $html | Should -Match '<div class="tk-main">'
        }
        It 'embeds the shared CSS block' {
            $html = Get-TKHtmlHead -Title 'X' -ScriptName 'X'
            $html | Should -Match '--tk-bg:'
            $html | Should -Match 'class="tk-page-header"'
        }
        It 'HTML-escapes a title containing special characters' {
            $html = Get-TKHtmlHead -Title '<script>&"' -ScriptName 'X'
            $html | Should -Match '&lt;script&gt;&amp;&quot;'
        }
        It 'renders a meta bar when MetaItems are supplied' {
            $meta = [ordered]@{ Generated = '2026-04-22'; Host = 'UNIT01' }
            $html = Get-TKHtmlHead -Title 'X' -ScriptName 'X' -MetaItems $meta
            $html | Should -Match 'class=''tk-meta-bar'''
            $html | Should -Match '>Generated<'
            $html | Should -Match '>UNIT01<'
        }
        It 'renders a nav bar when NavItems are supplied' {
            $html = Get-TKHtmlHead -Title 'X' -ScriptName 'X' -NavItems @('Alpha','Beta')
            $html | Should -Match 'class=''tk-nav'''
            $html | Should -Match 'Alpha</a>'
            $html | Should -Match 'Beta</a>'
        }
    }

    Context 'Get-TKHtmlFoot' {
        It 'closes the document' {
            $html = Get-TKHtmlFoot -ScriptName 'T.E.S.T. v1'
            $html | Should -Match '</body>'
            $html | Should -Match '</html>'
        }
        It 'includes the script name in the footer' {
            $html = Get-TKHtmlFoot -ScriptName 'T.E.S.T. v1'
            $html | Should -Match 'T\.E\.S\.T\. v1'
        }
    }

    Context 'round-trip' {
        It 'head + body + foot produces balanced HTML' {
            $doc = (Get-TKHtmlHead -Title 'X' -ScriptName 'X') + '<p>body</p>' + (Get-TKHtmlFoot -ScriptName 'X')
            # Each opening tag should have exactly one closing tag.
            ($doc | Select-String -Pattern '<html' -AllMatches).Matches.Count   | Should -Be 1
            ($doc | Select-String -Pattern '</html>' -AllMatches).Matches.Count | Should -Be 1
            ($doc | Select-String -Pattern '<body'  -AllMatches).Matches.Count  | Should -Be 1
            ($doc | Select-String -Pattern '</body>' -AllMatches).Matches.Count | Should -Be 1
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Get-TKConfig
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Get-TKConfig' {
    BeforeAll {
        # Point the module's config path at a temp directory
        $script:TempDir = Join-Path $TestDrive 'TKConfig'
        New-Item -ItemType Directory -Path $script:TempDir -Force | Out-Null
    }

    Context 'when config.json does not exist' {
        BeforeAll {
            # Import module with a fresh PSScriptRoot pointing at TempDir (no config.json there)
            # We test via the exported function directly; the real config path uses $PSScriptRoot.
            # To isolate, we check that defaults have the expected shape.
            $cfg = Get-TKConfig
        }

        It 'returns an object' {
            $cfg | Should -Not -BeNullOrEmpty
        }
        It 'has OrgName property' {
            $cfg.PSObject.Properties.Name | Should -Contain 'OrgName'
        }
        It 'has LogDirectory property' {
            $cfg.PSObject.Properties.Name | Should -Contain 'LogDirectory'
        }
        It 'has TeamsWebhook property' {
            $cfg.PSObject.Properties.Name | Should -Contain 'TeamsWebhook'
        }
        It 'has Archive section' {
            $cfg.Archive | Should -Not -BeNullOrEmpty
        }
        It 'has Revenant section' {
            $cfg.Revenant | Should -Not -BeNullOrEmpty
        }
        It 'has Covenant section' {
            $cfg.Covenant | Should -Not -BeNullOrEmpty
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Set-TKConfig / Get-TKConfig round-trip
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Set-TKConfig and Get-TKConfig round-trip' {
    BeforeAll {
        # We cannot easily redirect $PSScriptRoot inside the module, so these
        # tests verify the JSON serialisation logic using a local temp file.
        $script:ConfigFile = Join-Path $TestDrive 'config.json'
    }

    It 'writes and reads a top-level key' {
        # Write a minimal config directly to simulate what Set-TKConfig would produce
        [PSCustomObject]@{ OrgName = 'Contoso' } | ConvertTo-Json | Set-Content $script:ConfigFile -Encoding UTF8
        $raw = Get-Content $script:ConfigFile | ConvertFrom-Json
        $raw.OrgName | Should -Be 'Contoso'
    }

    It 'produces valid JSON' {
        { Get-Content $script:ConfigFile | ConvertFrom-Json } | Should -Not -Throw
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Test-IsAdmin
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Test-IsAdmin' {
    It 'returns a boolean' {
        $result = Test-IsAdmin
        $result | Should -BeOfType [bool]
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Write-TKError — smoke tests (no real network/filesystem side effects in CI)
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Write-TKError' {
    It 'does not throw when LogDirectory is not configured' {
        { Write-TKError -ScriptName 'test' -Message 'unit test error' -Category 'Test' } |
            Should -Not -Throw
    }

    It 'does not throw with a blank Teams webhook' {
        { Write-TKError -ScriptName 'test' -Message 'webhook test' } |
            Should -Not -Throw
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Module exports
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Module exports' {
    # Pester 5 data binding: -ForEach hashtable keys become variables in the
    # test body. A plain `foreach { It { ... $fn } }` leaks the discovery-time
    # loop variable out of scope by run phase — use -ForEach so $fn resolves.
    It 'exports <fn>' -ForEach @(
        @{ fn = 'Write-Section' }
        @{ fn = 'Write-Step' }
        @{ fn = 'Write-Ok' }
        @{ fn = 'Write-Warn' }
        @{ fn = 'Write-Fail' }
        @{ fn = 'Write-Info' }
        @{ fn = 'Show-TKReportResult' }
        @{ fn = 'EscHtml' }
        @{ fn = 'Format-Bytes' }
        @{ fn = 'Get-TKHtmlCss' }
        @{ fn = 'Get-TKHtmlHead' }
        @{ fn = 'Get-TKHtmlFoot' }
        @{ fn = 'Test-IsAdmin' }
        @{ fn = 'Assert-AdminPrivilege' }
        @{ fn = 'Invoke-AdminElevation' }
        @{ fn = 'Get-TKConfig' }
        @{ fn = 'Set-TKConfig' }
        @{ fn = 'Resolve-LogDirectory' }
        @{ fn = 'Start-TKTranscript' }
        @{ fn = 'Stop-TKTranscript' }
        @{ fn = 'Write-TKError' }
        @{ fn = 'Add-TKNote' }
        @{ fn = 'Get-TKNote' }
        @{ fn = 'Clear-TKNote' }
        @{ fn = 'Export-TKNoteReport' }
    ) {
        Get-Command $fn -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Technician notes — Add-TKNote / Get-TKNote / Clear-TKNote / Export-TKNoteReport
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Technician notes' {
    BeforeEach { Clear-TKNote }

    It 'starts the session with no notes' {
        @(Get-TKNote).Count | Should -Be 0
    }
    It 'records a note' {
        Add-TKNote -Text 'first note'
        $n = @(Get-TKNote)
        $n.Count   | Should -Be 1
        $n[0].Text | Should -Be 'first note'
    }
    It 'defaults the category to Info' {
        Add-TKNote -Text 'x'
        (@(Get-TKNote))[0].Category | Should -Be 'Info'
    }
    It 'rejects an invalid category' {
        { Add-TKNote -Text 'x' -Category 'Bogus' } | Should -Throw
    }
    It 'preserves insertion order' {
        Add-TKNote -Text 'one'
        Add-TKNote -Text 'two'
        $n = @(Get-TKNote)
        $n[0].Text | Should -Be 'one'
        $n[1].Text | Should -Be 'two'
    }
    It 'Clear-TKNote empties the buffer' {
        Add-TKNote -Text 'x'
        Clear-TKNote
        @(Get-TKNote).Count | Should -Be 0
    }

    Context 'Export-TKNoteReport' {
        It 'writes an HTML file and returns its path' {
            Add-TKNote -Text 'did a thing' -Category Action
            $out = Join-Path $TestDrive 'notes.html'
            Export-TKNoteReport -Path $out -ScriptName 'T.E.S.T.' | Should -Be $out
            $out | Should -Exist
        }
        It 'produces a balanced HTML document' {
            Add-TKNote -Text 'note body'
            $out = Join-Path $TestDrive 'notes2.html'
            Export-TKNoteReport -Path $out
            $doc = Get-Content $out -Raw
            $doc | Should -Match '^<!DOCTYPE html>'
            $doc | Should -Match '</html>'
        }
        It 'escapes HTML-special characters in note text' {
            Add-TKNote -Text '<script>alert(1)</script>'
            $out = Join-Path $TestDrive 'notes3.html'
            Export-TKNoteReport -Path $out
            $doc = Get-Content $out -Raw
            $doc | Should -Not -Match '<script>alert'
            $doc | Should -Match '&lt;script&gt;'
        }
        It 'includes the ticket reference when supplied' {
            Add-TKNote -Text 'x'
            $out = Join-Path $TestDrive 'notes4.html'
            Export-TKNoteReport -Path $out -Ticket 'INC0099'
            (Get-Content $out -Raw) | Should -Match 'INC0099'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Script syntax validation — all .ps1 files must parse without errors
# ─────────────────────────────────────────────────────────────────────────────
Describe 'PowerShell syntax — all scripts' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> has no parse errors' -ForEach $scriptCases {
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$null, [ref]$errors
        )
        $errors.Count | Should -Be 0
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# UTF-8 BOM — every .ps1/.psm1 must start with a UTF-8 byte-order mark.
# Windows PowerShell 5.1 reads a BOM-less file as ANSI (Windows-1252), which
# mangles the Unicode box-drawing banners and menu glyphs at parse time. A BOM
# forces UTF-8 decoding.
#
# This used to justify itself by the launcher shelling out to powershell.exe 5.1.
# The launcher is gone and the desktop app hosts PowerShell 7, which assumes
# UTF-8 — but the gate stays, because the scripts remain runnable standalone
# under Windows PowerShell 5.1 and that is still the primary documented path.
# CI runs under pwsh (PS7) so it would not otherwise catch a missing BOM.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'UTF-8 BOM — all scripts' {
    $bomCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Include '*.ps1', '*.psm1' -File -Recurse |
        Where-Object { $_.FullName -notmatch $NonSourceDir } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> begins with a UTF-8 BOM' -ForEach $bomCases {
        $bytes = [System.IO.File]::ReadAllBytes($FullName)
        $bytes.Length | Should -BeGreaterThan 2
        $bytes[0..2] -join ',' | Should -Be '239,187,191' -Because "$Name must start with a UTF-8 BOM (EF BB BF) so Windows PowerShell 5.1 parses it as UTF-8"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Module bootstrap compliance — every tool script must use the shared-module
# bootstrap block so the module is auto-downloaded when missing and imports
# fail loudly (-ErrorAction Stop) rather than silently partially-executing.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Module bootstrap compliance — all tool scripts' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        Where-Object { -not $RenamedToolStubs.Contains($_.Name) } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> defines $TKModulePath next to $PSScriptRoot' -ForEach $scriptCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Match '\$TKModulePath\s*=\s*Join-Path\s+\$PSScriptRoot\s+''TechnicianToolkit\.psm1'''
    }

    It '<Name> imports via $TKModulePath with -ErrorAction Stop' -ForEach $scriptCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Match 'Import-Module\s+\$TKModulePath\s+-Force\s+-ErrorAction\s+Stop'
    }

    It '<Name> no longer uses the silent-fail import' -ForEach $scriptCases {
        $content = Get-Content $FullName -Raw
        # The old pattern (quoted path, no -ErrorAction) must be gone.
        $content | Should -Not -Match 'Import-Module\s+"\$PSScriptRoot\\TechnicianToolkit\.psm1"\s+-Force\s*$'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TK_DISABLE_DOWNLOAD — the switch an MSP sets (machine environment variable,
# pushed by GPO or RMM) so the toolkit never fetches its own code from GitHub at
# run time. Under SOC 2 change management, code that pulls the current `main`
# whenever a file is missing is unreviewed change; with the switch set, only the
# release the MSP deployed can run. Four paths fetch toolkit code — the module
# bootstrap, the forwarding stubs, GRIMOIRE's tool launcher and RITUAL's step
# resolver — and every one must check the switch before it downloads. The test
# is generic so a fifth path cannot be added without it.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'TK_DISABLE_DOWNLOAD — every self-fetch path honours it' {
    $fetchCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        Where-Object { (Get-Content $_.FullName -Raw) -match 'raw\.githubusercontent\.com' } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It 'finds the self-fetching scripts' {
        $fetchCases.Count | Should -BeGreaterThan 50
    }

    It '<Name> checks TK_DISABLE_DOWNLOAD before each download' -ForEach $fetchCases {
        $content   = Get-Content $FullName -Raw
        $downloads = [regex]::Matches($content, 'Invoke-RestMethod\s+-Uri\s+\$\w+\s+-OutFile')
        $downloads.Count | Should -BeGreaterThan 0 -Because "$Name names the GitHub raw URL, so it should download from it"
        foreach ($d in $downloads) {
            $start  = [Math]::Max(0, $d.Index - 1500)
            $before = $content.Substring($start, $d.Index - $start)
            $before | Should -Match '\$env:TK_DISABLE_DOWNLOAD\s+-in\s+@\(''1'',\s*''true''\)' -Because "the download at offset $($d.Index) in $Name must be gated on TK_DISABLE_DOWNLOAD"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Param block compliance — interactive tool scripts must declare -Unattended
# Excludes the two launcher-style scripts that don't have sensible defaults
# for their required inputs (grimoire needs a tool choice; shade needs a
# target machine + credentials).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Param block compliance — -Unattended switch' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        Where-Object { $_.Name -notin @('grimoire.ps1', 'emissary.ps1') -and -not $RenamedToolStubs.Contains($_.Name) } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> declares -Unattended' -ForEach $scriptCases {
        $errors = $null
        $ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$null, [ref]$errors
        )
        $params = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.ParameterAst]
        }, $true)
        $paramNames = $params | ForEach-Object { $_.Name.VariablePath.UserPath }
        $paramNames | Should -Contain 'Unattended'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GRIMOIRE registry integrity — every File entry must exist on disk
# ─────────────────────────────────────────────────────────────────────────────
Describe 'GRIMOIRE registry integrity' {
    BeforeAll {
        # Pester 5: BeforeAll-scoped variables are accessible to It bodies
        # without $script: prefix. Discovery-scope vars (below) are not.
        $grimoirePath = Join-Path $PSScriptRoot '..\grimoire.ps1'
    }

    It 'grimoire.ps1 exists' {
        $grimoirePath | Should -Exist
    }

    # -ForEach cases must be built at discovery time.
    $GrimoirePath = Join-Path $PSScriptRoot '..\grimoire.ps1'
    $ToolkitRoot  = Join-Path $PSScriptRoot '..'
    $registryContent = Get-Content $GrimoirePath -Raw
    $registryCases   = [regex]::Matches($registryContent, "File\s*=\s*'([^']+)'") |
        ForEach-Object { $_.Groups[1].Value } |
        Select-Object -Unique |
        ForEach-Object { @{ FileName = $_; FullPath = (Join-Path $ToolkitRoot $_) } }

    It "registered tool '<FileName>' exists on disk" -ForEach $registryCases {
        $FullPath | Should -Exist
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Version consistency — a tool's .NOTES Version and its GRIMOIRE registry entry
# are two hand-maintained copies of one fact, and they drifted: before 5.0 the
# suite carried 3.6, 3.6.2, 3.8.3, 4.2 and 1.0 at once, with wyrm.ps1 and
# whetstone.ps1 disagreeing with their own registry rows. Nothing detected it.
#
# The header pattern deliberately tolerates irregular spacing — zenith.ps1
# writes 'Version  : 5.0' with two spaces — because a stricter pattern would
# silently skip that file rather than fail, which is how it drifted unnoticed.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Version consistency — script header matches the GRIMOIRE registry' {
    $ToolkitRoot     = Join-Path $PSScriptRoot '..'
    $registryContent = Get-Content (Join-Path $ToolkitRoot 'grimoire.ps1') -Raw

    # Each registry entry is a hashtable literal; pair the File with the Version
    # declared in the same entry rather than matching the two lists positionally.
    $versionCases = [regex]::Matches(
            $registryContent, "File\s*=\s*'([^']+)'[\s\S]{0,400}?Version\s*=\s*'([^']+)'") |
        ForEach-Object {
            @{
                FileName        = $_.Groups[1].Value
                RegistryVersion = $_.Groups[2].Value
                FullPath        = (Join-Path $ToolkitRoot $_.Groups[1].Value)
            }
        }

    # Guards the regex itself: a silently empty case list would make every
    # -ForEach below vacuous and the whole gate would pass by doing nothing.
    # The count rides in as test data because $versionCases is a discovery-time
    # variable and an It body runs later, in a scope where it no longer exists.
    It 'the registry yielded version entries to check' -ForEach @{ CaseCount = $versionCases.Count } {
        $CaseCount | Should -BeGreaterThan 40
    }

    It "<FileName> header version matches its registry entry (<RegistryVersion>)" -ForEach $versionCases {
        $raw = Get-Content $FullPath -Raw
        $raw | Should -Match '(?m)^\s*(?:#\s*)?Version\s*:\s*[0-9][0-9.]*' -Because "$FileName must declare a .NOTES Version"

        $headerVersion = ([regex]::Match($raw, '(?m)^\s*(?:#\s*)?Version\s*:\s*([0-9][0-9.]*)')).Groups[1].Value
        $headerVersion | Should -Be $RegistryVersion -Because "$FileName's header and its GRIMOIRE registry entry are the same fact and must agree"
    }

    # The suite ships as one product; a split version line is what 5.0 set out
    # to end. This fails loudly on the next tool added at its own number.
    It 'every registered tool reports one single version across the suite' -ForEach @{
            Declared = @($versionCases | ForEach-Object { $_.RegistryVersion } | Select-Object -Unique) } {
        $Declared.Count | Should -Be 1 -Because "the registry declares: $($Declared -join ', ')"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Legacy-name regression — the v3.0 rename retired eight tool acronyms; the
# v3.1 cleanup deleted their forwarding stubs. New source or documentation
# must never reintroduce the retired names. CHANGELOG (which documents the
# rename) is exempt.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Legacy tool names must not reappear' {
    # Pester 5 scoping: the pattern list must be defined in BeforeAll so it
    # exists at Run time when each It block executes. Variables assigned in
    # the Describe body only exist at Discovery time.
    BeforeAll {
        $script:LegacyAcronyms = @(
            'O.R.A.C.L.E.', 'S.E.N.T.I.N.E.L.', 'B.A.S.T.I.O.N.',
            'V.A.U.L.T.',   'P.H.A.N.T.O.M.',   'S.P.E.C.T.E.R.',
            'A.E.G.I.S.',   'R.E.L.I.C.',
            # Renamed in 6.0 (per-category naming themes)
            'R.E.S.T.O.R.A.T.I.O.N.', 'S.U.T.U.R.E.', 'T.H.R.E.S.H.O.L.D.', 'G.A.R.G.O.Y.L.E.',
            'P.Y.R.E.', 'C.I.P.H.E.R.', 'S.I.G.I.L.', 'C.I.T.A.D.E.L.',
            'A.R.T.I.F.A.C.T.', 'P.A.L.A.D.I.N.', 'H.E.R.A.L.D.', 'C.A.T.A.C.O.M.B.',
            'B.E.A.C.O.N.', 'S.H.A.D.E.', 'O.A.T.H.', 'T.A.L.I.S.M.A.N.',
            'R.E.L.I.Q.U.A.R.Y.', 'G.O.L.E.M.', 'W.R.A.I.T.H.', 'C.O.N.C.L.A.V.E.',
            'G.R.O.V.E.', 'T.E.N.D.R.I.L.', 'R.A.M.P.A.R.T.', 'A.R.C.H.I.V.E.',
            'T.E.T.H.E.R.'
        )
    }

    # Files that are *about* the rename legitimately mention the retired names.
    $allowlist = @('CHANGELOG.md', 'TechnicianToolkit.Tests.ps1')

    $root  = Resolve-Path (Join-Path $PSScriptRoot '..')
    $files = Get-ChildItem -Path $root -Recurse -File -Include '*.ps1','*.md' |
        Where-Object {
            $_.FullName -notmatch $NonSourceDir -and
            $_.Name -notin $allowlist
        } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> contains no retired dotted acronyms' -ForEach $files {
        $hits = Select-String -Path $FullName -SimpleMatch -Pattern $script:LegacyAcronyms -ErrorAction SilentlyContinue
        $hits | Should -BeNullOrEmpty -Because "retired acronym found in $Name"
    }

    # Second form: bare underscore-prefixed filenames (e.g. `SPECTER_<MachineName>`,
    # `PHANTOM_MigrationLog_*.csv`). The v3.0 rename changed every tool's emitted
    # filename prefix, and the README logging table drifted without this catch
    # because the dotted form above didn't match the bare-prefix form.
    It '<Name> contains no retired filename prefixes' -ForEach $files {
        $prefixPatterns = @(
            'ORACLE_', 'SENTINEL_', 'BASTION_', 'VAULT_',
            'PHANTOM_', 'SPECTER_', 'AEGIS_', 'RELIC_',
            'RESTORATION_', 'SUTURE_', 'THRESHOLD_', 'GARGOYLE_', 'PYRE_', 'CIPHER_',
            'SIGIL_', 'CITADEL_', 'ARTIFACT_', 'PALADIN_', 'HERALD_', 'CATACOMB_',
            'BEACON_', 'SHADE_', 'OATH_', 'TALISMAN_', 'RELIQUARY_', 'GOLEM_',
            'WRAITH_', 'CONCLAVE_', 'GROVE_', 'TENDRIL_', 'RAMPART_', 'ARCHIVE_',
            'TETHER_'
        )
        $hits = Select-String -Path $FullName -SimpleMatch -Pattern $prefixPatterns -ErrorAction SilentlyContinue
        $hits | Should -BeNullOrEmpty -Because "retired filename prefix found in $Name"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# -WhatIf compliance — destructive tools must declare -WhatIf so that GRIMOIRE
# can pass it through in dry-run mode. See grimoire.ps1 Invoke-Tool.
# ─────────────────────────────────────────────────────────────────────────────
Describe '-WhatIf declared on destructive tools' {
    # Tools that make persistent, hard-to-reverse changes: file moves, registry
    # writes, domain joins, disk encryption toggles, AV policy changes, driver
    # and Windows Update installs, printer driver / network printer additions.
    $destructiveCases = @(
        'revenant.ps1','embalm.ps1','covenant.ps1','basilisk.ps1','cleanse.ps1','wyrm.ps1',
        'forge.ps1','whetstone.ps1','runepress.ps1','conjure.ps1','conduit.ps1',
        'lodestar.ps1','solder.ps1','chalice.ps1'
    ) | ForEach-Object {
        @{ Name = $_; FullName = (Join-Path $PSScriptRoot "..\$_") }
    }

    It '<Name> declares -WhatIf' -ForEach $destructiveCases {
        $errors = $null
        $ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$null, [ref]$errors
        )
        $paramNames = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.ParameterAst]
        }, $true) | ForEach-Object { $_.Name.VariablePath.UserPath }
        $paramNames | Should -Contain 'WhatIf'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Deprecation stubs removed — the eight v3.0 forwarding stubs were retired in
# v3.1. Their filenames must not reappear in the working tree (any reintroduction
# would resurrect a name we have explicitly deleted).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Deprecation stubs removed' {
    $retiredStubs = @(
        'oracle.ps1','sentinel.ps1','bastion.ps1','vault.ps1',
        'phantom.ps1','specter.ps1','aegis.ps1','relic.ps1'
    ) | ForEach-Object {
        @{ Name = $_; FullPath = (Join-Path $PSScriptRoot "..\$_") }
    }

    It '<Name> no longer exists' -ForEach $retiredStubs {
        $FullPath | Should -Not -Exist
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Renamed-tool forwarding stubs — 6.0 renamed 25 tools and left a stub at each
# old filename so pinned runbooks, custom RITUAL recipes and bookmarked
# quick-launch URLs keep working for a release or two. A stub must forward to
# the right file, the file it names must be a registered tool, and the old name
# must not be registered itself -- otherwise the stub shadows or points at
# nothing.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Renamed-tool forwarding stubs' {
    BeforeAll {
        $script:StubRegistry = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'grimoire.ps1') -Raw
    }

    $stubCases = foreach ($old in $RenamedToolStubs.Keys) {
        @{ Old = $old; New = $RenamedToolStubs[$old]; FullName = (Join-Path (Join-Path $PSScriptRoot '..') $old) }
    }

    It '<Old> forwards every argument to <New>' -ForEach $stubCases {
        $FullName | Should -Exist
        $content = Get-Content $FullName -Raw
        $content | Should -Match ([regex]::Escape("Join-Path `$PSScriptRoot '$New'"))
        $content | Should -Match '&\s+\$NewScript\s+@args'
        $content | Should -Match 'Write-Warning'
    }

    It '<Old> declares no param() block, so @args carries named parameters through' -ForEach $stubCases {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($FullName, [ref]$null, [ref]$null)
        $ast.ParamBlock | Should -BeNullOrEmpty
    }

    It '<New> exists and is registered, and <Old> is not' -ForEach $stubCases {
        Join-Path (Join-Path $PSScriptRoot '..') $New | Should -Exist
        $script:StubRegistry | Should -Match ("File\s*=\s*'" + [regex]::Escape($New) + "'")
        $script:StubRegistry | Should -Not -Match ("File\s*=\s*'" + [regex]::Escape($Old) + "'")
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Duplicated helpers removed — the local HtmlEncode / Format-Bytes definitions
# that were consolidated into the shared module must not reappear.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'No duplicated helper functions' {
    $toolCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> does not redefine HtmlEncode locally' -ForEach $toolCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Not -Match '(?m)^\s*function\s+HtmlEncode\b'
    }

    It '<Name> does not redefine Format-Bytes locally' -ForEach $toolCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Not -Match '(?m)^\s*function\s+Format-Bytes\b'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Generic collections built with New-Object — under PowerShell 7.4, @() over a
# List[object] created by New-Object throws "Argument types do not match"; the
# same list from ::new() does not. Windows PowerShell 5.1 is unaffected, so the
# primary path never showed it, but the desktop app hosts PowerShell 7 -- and
# ARGUS lost its report to it twice before the cause was known. ::new() works
# on 5.1 too, so the gate bans the New-Object form for every generic collection
# rather than tracking which variables later meet an @(). This test file is
# exempt: the ConvertTo-ArgusArray tests build the New-Object form on purpose.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'No generic collections built with New-Object' {
    $collectionCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Include '*.ps1', '*.psm1' -File -Recurse |
        Where-Object { $_.FullName -notmatch $NonSourceDir -and $_.Name -ne 'TechnicianToolkit.Tests.ps1' } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> builds generic collections with ::new(), not New-Object' -ForEach $collectionCases {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($FullName, [ref]$null, [ref]$null)
        $hits = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'New-Object' -and
            $n.Extent.Text -match 'System\.Collections\.Generic\.'
        }, $true) | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Extent.Text)" }
        $hits -join '; ' | Should -BeNullOrEmpty -Because 'use [System.Collections.Generic.List[object]]::new() -- @() over the New-Object form throws on PowerShell 7.4'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Tier-mapper data tables — the verdict logic in GRIFFIN / WISP / PORTAL
# leans on small reference hashtables (and one tiny helper for ASR action
# codes). This block extracts those tables via AST lookup and asserts on
# their contents, so a careless rename or removal fails CI loudly.
#
# We don't dot-source the tools whole because they have side-effects (they
# launch their main flow on import). Instead we walk the AST, find the
# specific assignment / function-definition node by name, and re-evaluate
# just that node in the test scope.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Tier-mapper data tables' {
    BeforeAll {
        $script:ToolkitRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
        $script:GriffinPath = Join-Path $script:ToolkitRoot 'griffin.ps1'
        $script:WispPath  = Join-Path $script:ToolkitRoot 'wisp.ps1'
        $script:PortalPath  = Join-Path $script:ToolkitRoot 'portal.ps1'
        $script:ConjurePath = Join-Path $script:ToolkitRoot 'conjure.ps1'
        $script:ArgusPath  = Join-Path $script:ToolkitRoot 'argus.ps1'
        $script:ConduitPath = Join-Path $script:ToolkitRoot 'conduit.ps1'

        function Import-ScriptHashtable {
            param([string]$ScriptPath, [string]$VarName)
            $errs = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $ScriptPath, [ref]$null, [ref]$errs
            )
            $assign = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $n.Left.VariablePath.UserPath -eq $VarName
            }, $true) | Select-Object -First 1
            if (-not $assign) { return $null }
            return & ([scriptblock]::Create($assign.Right.Extent.Text))
        }

        function Get-ScriptFunctionScriptBlock {
            param([string]$ScriptPath, [string]$FuncName)
            $errs = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $ScriptPath, [ref]$null, [ref]$errs
            )
            $fn = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FuncName
            }, $true) | Select-Object -First 1
            if (-not $fn) { return $null }
            # Returns a scriptblock that, when dot-sourced, defines the function in the caller's scope.
            return [scriptblock]::Create($fn.Extent.Text)
        }
    }

    Context 'GRIFFIN: $AsrRuleNames hashtable' {
        It 'covers at least 16 well-known ASR rule GUIDs' {
            $asr = Import-ScriptHashtable -ScriptPath $script:GriffinPath -VarName 'AsrRuleNames'
            $asr | Should -Not -BeNullOrEmpty
            $asr.Count | Should -BeGreaterOrEqual 16
        }
        It 'maps the abused-driver GUID to a recognisable name' {
            $asr = Import-ScriptHashtable -ScriptPath $script:GriffinPath -VarName 'AsrRuleNames'
            $asr['56a863a9-875e-4185-98a7-b882c64b5ce5'] | Should -Match 'vulnerable signed drivers'
        }
        It 'maps the LSASS-credential-theft GUID' {
            $asr = Import-ScriptHashtable -ScriptPath $script:GriffinPath -VarName 'AsrRuleNames'
            $asr['9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2'] | Should -Match 'LSASS'
        }
        It 'uses lowercase GUID keys (matches the case from Get-MpPreference output)' {
            $asr = Import-ScriptHashtable -ScriptPath $script:GriffinPath -VarName 'AsrRuleNames'
            foreach ($k in $asr.Keys) {
                $k | Should -Match '^[0-9a-f-]+$' -Because "ASR keys must be lowercase to match Get-MpPreference output (saw '$k')"
            }
        }
    }

    Context 'GRIFFIN: Get-AsrActionLabel function' {
        BeforeAll {
            # Extract the function definition AST from griffin.ps1 and dot-source
            # it so this Context can call the function directly.
            $sb = Get-ScriptFunctionScriptBlock -ScriptPath $script:GriffinPath -FuncName 'Get-AsrActionLabel'
            . $sb
        }
        It 'maps action 0 to Not Configured' {
            Get-AsrActionLabel -Action 0 | Should -Be 'Not Configured'
        }
        It 'maps action 1 to Block' {
            Get-AsrActionLabel -Action 1 | Should -Be 'Block'
        }
        It 'maps action 2 to Audit' {
            Get-AsrActionLabel -Action 2 | Should -Be 'Audit'
        }
        It 'maps action 6 to Warn' {
            Get-AsrActionLabel -Action 6 | Should -Be 'Warn'
        }
        It 'falls back to "Unknown (<n>)" for unmapped codes' {
            Get-AsrActionLabel -Action 99 | Should -Be 'Unknown (99)'
        }
    }

    Context 'CONDUIT: $ConduitFindings catalog' {
        It 'covers every finding code the tool can raise' {
            $f = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'ConduitFindings'
            $f | Should -Not -BeNullOrEmpty
            foreach ($code in @(
                'TimeServiceStopped','WinHttpProxySet','WsusUnreachable','BlockInternetWU',
                'WUAccessDisabled','AutoUpdateDisabled','MicrosoftEndpointsBlocked',
                'ServiceDisabled','WUServerUnparsable'
            )) {
                $f.ContainsKey($code) | Should -BeTrue -Because "Add-ConduitFinding raises '$code'"
            }
        }
        It 'gives every finding a Severity, Title, Summary and Remedy' {
            $f = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'ConduitFindings'
            foreach ($code in $f.Keys) {
                $f[$code].Severity | Should -Not -BeNullOrEmpty -Because "$code needs a severity"
                $f[$code].Title    | Should -Not -BeNullOrEmpty -Because "$code needs a title"
                $f[$code].Summary  | Should -Not -BeNullOrEmpty -Because "$code needs a summary"
                $f[$code].Remedy   | Should -Not -BeNullOrEmpty -Because "$code needs a remedy"
            }
        }
        It 'uses only severities the report can render as a badge class' {
            $f = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'ConduitFindings'
            foreach ($code in $f.Keys) {
                $f[$code].Severity | Should -BeIn @('Error','Warning','Info') -Because "Get-SeverityClass only maps these (saw '$($f[$code].Severity)' on $code)"
            }
        }
        It 'ranks an unreachable WSUS pointer and a policy internet block as Errors' {
            $f = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'ConduitFindings'
            $f['WsusUnreachable'].Severity  | Should -Be 'Error'
            $f['BlockInternetWU'].Severity  | Should -Be 'Error'
            $f['WUAccessDisabled'].Severity | Should -Be 'Error'
        }
        It 'ranks a stopped time service and a set proxy as Warnings, not Errors' {
            # Neither blocks the update service outright -- a proxy may be
            # intentional, and clock drift only breaks TLS past ~5 minutes.
            $f = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'ConduitFindings'
            $f['TimeServiceStopped'].Severity | Should -Be 'Warning'
            $f['WinHttpProxySet'].Severity    | Should -Be 'Warning'
        }
    }

    Context 'CONDUIT: $UpdateServiceDefaults table' {
        It 'covers the four services the update client depends on' {
            $d = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'UpdateServiceDefaults'
            $d.Keys | Should -Contain 'wuauserv'
            $d.Keys | Should -Contain 'bits'
            $d.Keys | Should -Contain 'cryptsvc'
            $d.Keys | Should -Contain 'usosvc'
        }
        It 'restores Windows defaults rather than forcing everything Automatic' {
            # wuauserv and bits are demand-started by design; setting them
            # Automatic would be a behaviour change, not a repair.
            $d = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'UpdateServiceDefaults'
            $d['wuauserv'] | Should -Be 'Manual'
            $d['bits']     | Should -Be 'Manual'
            $d['cryptsvc'] | Should -Be 'Automatic'
            $d['usosvc']   | Should -Be 'Automatic'
        }
        It 'names only start types Set-Service accepts' {
            $d = Import-ScriptHashtable -ScriptPath $script:ConduitPath -VarName 'UpdateServiceDefaults'
            foreach ($k in $d.Keys) {
                $d[$k] | Should -BeIn @('Automatic','Manual','Disabled','Boot','System')
            }
        }
    }

    Context 'WISP: $AuthStrength hashtable (Wi-Fi)' {
        It 'classifies open / shared / WEP as Insecure' {
            $auth = Import-ScriptHashtable -ScriptPath $script:WispPath -VarName 'AuthStrength'
            $auth['open']   | Should -Be 'Insecure'
            $auth['shared'] | Should -Be 'Insecure'
            $auth['WEP']    | Should -Be 'Insecure'
        }
        It 'classifies WPA1 (WPA / WPAPSK) as Weak' {
            $auth = Import-ScriptHashtable -ScriptPath $script:WispPath -VarName 'AuthStrength'
            $auth['WPA']    | Should -Be 'Weak'
            $auth['WPAPSK'] | Should -Be 'Weak'
        }
        It 'classifies WPA2 personal+enterprise / WPA3 / OWE as Strong' {
            $auth = Import-ScriptHashtable -ScriptPath $script:WispPath -VarName 'AuthStrength'
            $auth['WPA2']    | Should -Be 'Strong'
            $auth['WPA2PSK'] | Should -Be 'Strong'
            $auth['WPA3SAE'] | Should -Be 'Strong'
            $auth['WPA3ENT'] | Should -Be 'Strong'
            $auth['OWE']     | Should -Be 'Strong'
        }
    }

    Context 'WISP: $CipherStrength hashtable (Wi-Fi)' {
        It 'classifies cipher tiers correctly' {
            $c = Import-ScriptHashtable -ScriptPath $script:WispPath -VarName 'CipherStrength'
            $c['none'] | Should -Be 'Insecure'
            $c['WEP']  | Should -Be 'Insecure'
            $c['TKIP'] | Should -Be 'Weak'
            $c['AES']  | Should -Be 'Strong'
            $c['GCMP'] | Should -Be 'Strong'
        }
    }

    Context 'PORTAL: $AuthStrength hashtable (VPN)' {
        It 'classifies PAP as Insecure (cleartext credentials)' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['Pap'] | Should -Be 'Insecure'
        }
        It 'classifies CHAP as Weak' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['Chap'] | Should -Be 'Weak'
        }
        It 'classifies MS-CHAPv2 as Acceptable' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['MSChapv2'] | Should -Be 'Acceptable'
        }
        It 'classifies EAP and MachineCertificate as Strong' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['Eap']                | Should -Be 'Strong'
            $a['MachineCertificate'] | Should -Be 'Strong'
        }
    }

    Context 'PORTAL: $EncryptionStrength hashtable (VPN)' {
        It 'classifies encryption levels correctly' {
            $e = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'EncryptionStrength'
            $e['NoEncryption'] | Should -Be 'Insecure'
            $e['Optional']     | Should -Be 'Weak'
            $e['Required']     | Should -Be 'Strong'
            $e['Maximum']      | Should -Be 'Strong'
        }
    }

    Context 'CONJURE: $InstallExitInfo exit-code table' {
        It 'decodes the installer-hash-mismatch code (0x8A150011) as a failure' {
            $t = Import-ScriptHashtable -ScriptPath $script:ConjurePath -VarName 'InstallExitInfo'
            $t | Should -Not -BeNullOrEmpty
            $t['0x8A150011'].Class  | Should -Be 'Failed'
            $t['0x8A150011'].Reason | Should -Match 'hash'
        }
        It 'classifies already-installed / up-to-date winget codes as success no-ops' {
            $t = Import-ScriptHashtable -ScriptPath $script:ConjurePath -VarName 'InstallExitInfo'
            $t['0x8A150061'].Class | Should -Be 'Success'
            $t['0x8A15002B'].Class | Should -Be 'Success'
        }
        It 'classifies clean-success and reboot-pending codes as success' {
            $t = Import-ScriptHashtable -ScriptPath $script:ConjurePath -VarName 'InstallExitInfo'
            $t['0x00000000'].Class | Should -Be 'Success'  # 0
            $t['0x00000BC2'].Class | Should -Be 'Success'  # 3010 reboot required
        }
        It 'classifies other known winget failures as failures' {
            $t = Import-ScriptHashtable -ScriptPath $script:ConjurePath -VarName 'InstallExitInfo'
            $t['0x8A150010'].Class | Should -Be 'Failed'  # no applicable installer
            $t['0x8A150008'].Class | Should -Be 'Failed'  # download failed
        }
        It 'every entry carries a Class and a non-empty Reason' {
            $t = Import-ScriptHashtable -ScriptPath $script:ConjurePath -VarName 'InstallExitInfo'
            foreach ($k in $t.Keys) {
                $t[$k].Class  | Should -BeIn @('Success', 'Warning', 'Failed') -Because "code $k must be classified"
                $t[$k].Reason | Should -Not -BeNullOrEmpty -Because "code $k must have a human-readable reason"
            }
        }
        It 'normalises a signed $LASTEXITCODE back to the unsigned hex key the table uses' {
            # winget's 0x8A150011 is surfaced by $LASTEXITCODE as -1978335215;
            # Resolve-InstallExit relies on this conversion to find the entry.
            ('0x{0:X8}' -f (-1978335215 -band 4294967295)) | Should -Be '0x8A150011'
        }
    }

    Context 'ARGUS: $PrivilegedGroupTiers group table' {
        It 'covers the four forest/domain-wide administrative groups' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            $t | Should -Not -BeNullOrEmpty
            foreach ($g in 'Enterprise Admins', 'Schema Admins', 'Domain Admins', 'Administrators') {
                $t[$g].Role | Should -Be 'Domain Administrator' -Because "$g confers full domain control"
            }
        }
        It 'classifies the built-in operator groups as delegated administrators' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            foreach ($g in 'Account Operators', 'Server Operators', 'Backup Operators', 'Print Operators') {
                $t[$g].Role | Should -Be 'Delegated Administrator'
            }
        }
        It 'resolves domain-scoped groups by their well-known RID' {
            # Resolving by RID rather than name is what makes the audit survive a
            # renamed or localised "Domain Admins".
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            $t['Domain Admins'].Rid       | Should -Be 512
            $t['Domain Admins'].Scope     | Should -Be 'Domain'
            $t['Enterprise Admins'].Rid   | Should -Be 519
            $t['Schema Admins'].Rid       | Should -Be 518
        }
        It 'resolves BUILTIN groups against the S-1-5-32 authority' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            $t['Administrators'].Scope     | Should -Be 'Builtin'
            $t['Administrators'].Rid       | Should -Be 544
            $t['Backup Operators'].Scope   | Should -Be 'Builtin'
            $t['Backup Operators'].Rid     | Should -Be 551
        }
        It 'resolves DnsAdmins by name, since it has no fixed RID' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            $t['DnsAdmins'].Scope | Should -Be 'Name'
            $t['DnsAdmins'].Rid   | Should -BeNullOrEmpty
        }
        It 'every entry carries a known role and a non-empty reason' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            foreach ($k in $t.Keys) {
                $t[$k].Role   | Should -BeIn @('Domain Administrator', 'Delegated Administrator') -Because "group $k must map to a role ARGUS ranks"
                $t[$k].Reason | Should -Not -BeNullOrEmpty -Because "group $k must explain to the customer why it matters"
                $t[$k].Scope  | Should -BeIn @('Domain', 'Builtin', 'Name')
            }
        }
    }

    Context 'ARGUS: $PasswordPolicyBaseline table' {
        It 'covers the four settings the access questionnaire asks about' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PasswordPolicyBaseline'
            $t | Should -Not -BeNullOrEmpty
            foreach ($k in 'MinPasswordLength', 'ComplexityEnabled', 'PasswordHistoryCount',
                           'MaxPasswordAgeDays', 'LockoutThreshold') {
                $t.Contains($k) | Should -BeTrue -Because "the questionnaire asks about $k"
            }
        }
        It 'gives every setting a Kind the verdict logic can evaluate' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PasswordPolicyBaseline'
            foreach ($k in $t.Keys) {
                $t[$k].Kind  | Should -BeIn @('Number', 'Threshold', 'Age', 'Boolean', 'Duration')
                $t[$k].Label | Should -Not -BeNullOrEmpty
                $t[$k].Why   | Should -Not -BeNullOrEmpty -Because "$k is explained to the customer in the report"
            }
        }
        It 'treats lockout duration as a Duration, not a plain number' {
            # 0 minutes means "until an administrator unlocks", the strictest
            # setting available; scoring it as higher-is-better calls it Weak.
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PasswordPolicyBaseline'
            $t['LockoutDurationMinutes'].Kind | Should -Be 'Duration'
        }
        It 'flags reversible encryption as bad when enabled' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PasswordPolicyBaseline'
            $t['ReversibleEncryptionEnabled'].Kind | Should -Be 'Boolean'
            $t['ReversibleEncryptionEnabled'].Good | Should -BeFalse
        }
    }

    Context 'ARGUS: $RoleTiers rank table' {
        It 'ranks the roles from most to least privileged' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'RoleTiers'
            $t['Domain Administrator'].Rank    | Should -Be 1
            $t['Delegated Administrator'].Rank | Should -Be 2
            $t['Elevated (Custom Group)'].Rank | Should -Be 3
            $t['Standard User'].Rank           | Should -Be 4
        }
        It 'covers every role $PrivilegedGroupTiers can produce' {
            # Build-AccountRoster indexes $RoleTiers by the Role string a group
            # carries; a role with no tier entry would throw at classification time.
            $groups = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'PrivilegedGroupTiers'
            $roles  = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'RoleTiers'
            foreach ($k in $groups.Keys) {
                $roles.Contains($groups[$k].Role) | Should -BeTrue -Because "role '$($groups[$k].Role)' (from $k) must have a rank"
            }
        }
        It 'assigns every role a badge class the shared CSS defines' {
            $t = Import-ScriptHashtable -ScriptPath $script:ArgusPath -VarName 'RoleTiers'
            foreach ($k in $t.Keys) {
                $t[$k].Badge | Should -BeIn @('ok', 'warn', 'err', 'info')
                $t[$k].Blurb | Should -Not -BeNullOrEmpty -Because "role $k is explained to the customer in the report"
            }
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# NECROPSY — the finding catalog, the bugcheck catalog, and the pure parsers
# that turn event text, Kernel-Power fields and dump headers into incidents.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'NECROPSY crash analysis helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'necropsy.ps1'
        $NecropsyFindings = Get-ToolAssignmentValue -Ast $ast -VarName 'NecropsyFindings'
        $BugCheckCatalog  = Get-ToolAssignmentValue -Ast $ast -VarName 'BugCheckCatalog'
        $AreaGuidance     = Get-ToolAssignmentValue -Ast $ast -VarName 'AreaGuidance'
        foreach ($name in 'ConvertTo-BugCheckKey', 'ConvertFrom-BugCheckText', 'Get-KernelPowerCause',
                          'Get-DumpHeaderInfo', 'Get-BugCheckInfo', 'Get-NecropsyVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $necropsySource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'necropsy.ps1') -Raw

        function New-DumpHeader {
            param([string]$Signature, [int]$CodeOffset, [uint32]$Code)
            $bytes = New-Object byte[] 0x60
            [System.Text.Encoding]::ASCII.GetBytes($Signature).CopyTo($bytes, 0)
            [BitConverter]::GetBytes($Code).CopyTo($bytes, $CodeOffset)
            return , $bytes
        }
    }

    Context 'finding catalog' {
        It 'contains every code the tool raises' {
            $raised = [regex]::Matches($necropsySource, "Add-NecropsyFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
            @($raised).Count | Should -BeGreaterThan 5
            foreach ($code in $raised) {
                $NecropsyFindings.ContainsKey($code) | Should -BeTrue -Because "Add-NecropsyFinding raises '$code'"
            }
        }
        It 'gives every finding a Severity, Kind, Title, Summary and Remedy' {
            foreach ($code in $NecropsyFindings.Keys) {
                $f = $NecropsyFindings[$code]
                $f.Severity | Should -BeIn @('Error', 'Warning', 'Info') -Because "$code needs a renderable severity"
                $f.Kind     | Should -BeIn @('Crash', 'Instability', 'Readiness') -Because "the verdict reads Kind ($code)"
                $f.Title    | Should -Not -BeNullOrEmpty
                $f.Summary  | Should -Not -BeNullOrEmpty
                $f.Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'classes crash-dump configuration gaps as Readiness, not evidence' {
            $NecropsyFindings['DumpsDisabled'].Kind | Should -Be 'Readiness'
            $NecropsyFindings['NoPageFile'].Kind    | Should -Be 'Readiness'
            $NecropsyFindings['BugCheck'].Kind      | Should -Be 'Crash'
        }
    }

    Context 'bugcheck catalog' {
        It 'uses the canonical key form ConvertTo-BugCheckKey produces' {
            foreach ($k in $BugCheckCatalog.Keys) {
                $k | Should -MatchExactly '^0x[0-9A-F]+$' -Because "'$k' would never be looked up"
                ConvertTo-BugCheckKey -Code ([Convert]::ToUInt64($k.Substring(2), 16)) | Should -BeExactly $k
            }
        }
        It 'maps every entry to an area the guidance table covers' {
            foreach ($k in $BugCheckCatalog.Keys) {
                $AreaGuidance.Contains($BugCheckCatalog[$k].Area) | Should -BeTrue -Because "$k names area '$($BugCheckCatalog[$k].Area)'"
            }
        }
        It 'covers the common field bugchecks' {
            foreach ($k in '0xA', '0x1A', '0x3B', '0x50', '0x7E', '0x9F', '0xD1', '0xEF', '0x124', '0x133') {
                $BugCheckCatalog.ContainsKey($k) | Should -BeTrue -Because "$k is one of the most common stop codes"
            }
        }
        It 'falls back to Unknown for a code outside the catalog' {
            (Get-BugCheckInfo -Key '0xDEAD').Area | Should -Be 'Unknown'
        }
    }

    Context 'parsers' {
        It 'normalises padded, short and 32-bit-overflowing codes to one key' {
            ConvertTo-BugCheckKey -Code 0x0000009f | Should -BeExactly '0x9F'
            ConvertTo-BugCheckKey -Code 209        | Should -BeExactly '0xD1'
            ConvertTo-BugCheckKey -Code 3221226010 | Should -BeExactly '0xC000021A'
        }
        It 'reads the code and four parameters out of WER 1001 text' {
            $r = ConvertFrom-BugCheckText -Text '0x0000009f (0x0000000000000003, 0xffffe001d5c96060, 0xfffff80000000000, 0xffffe001d8c1f010)'
            $r.Key               | Should -BeExactly '0x9F'
            $r.Parameters.Count  | Should -Be 4
            $r.Parameters[0]     | Should -Be '0x0000000000000003'
        }
        It 'returns nothing for text with no hex code' {
            ConvertFrom-BugCheckText -Text 'no code here' | Should -BeNullOrEmpty
        }
        It 'classifies Kernel-Power 41 as bugcheck, power button or power loss' {
            Get-KernelPowerCause -BugcheckCode 209                              | Should -Be 'Bugcheck'
            Get-KernelPowerCause -BugcheckCode 0 -PowerButtonTimestamp 1324567  | Should -Be 'PowerButton'
            Get-KernelPowerCause -BugcheckCode 0 -PowerButtonTimestamp 0        | Should -Be 'PowerLoss'
        }
        It 'reads the bugcheck code from a 64-bit dump header' {
            $h = Get-DumpHeaderInfo -Bytes (New-DumpHeader -Signature 'PAGEDU64' -CodeOffset 0x38 -Code 0xD1)
            $h.Format | Should -Be '64-bit'
            $h.Key    | Should -BeExactly '0xD1'
        }
        It 'reads the bugcheck code from a 32-bit dump header' {
            $h = Get-DumpHeaderInfo -Bytes (New-DumpHeader -Signature 'PAGEDUMP' -CodeOffset 0x28 -Code 0x124)
            $h.Format | Should -Be '32-bit'
            $h.Key    | Should -BeExactly '0x124'
        }
        It 'rejects a file that is not a kernel dump, or is too short' {
            (Get-DumpHeaderInfo -Bytes (New-DumpHeader -Signature 'MDMP....' -CodeOffset 0x38 -Code 1)).Format | Should -Be 'Unknown'
            (Get-DumpHeaderInfo -Bytes (New-Object byte[] 8)).Format | Should -Be 'Unknown'
        }
    }

    Context 'verdict' {
        It 'reports Stable when the only findings are readiness gaps' {
            (Get-NecropsyVerdict -FindingList @([PSCustomObject]@{ Kind = 'Readiness' })).Verdict | Should -Be 'Stable'
            (Get-NecropsyVerdict -FindingList @()).Verdict | Should -Be 'Stable'
        }
        It 'reports Unstable on instability and Crashing on any crash' {
            (Get-NecropsyVerdict -FindingList @([PSCustomObject]@{ Kind = 'Instability' })).Verdict | Should -Be 'Unstable'
            (Get-NecropsyVerdict -FindingList @([PSCustomObject]@{ Kind = 'Instability' }, [PSCustomObject]@{ Kind = 'Crash' })).Verdict | Should -Be 'Crashing'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# RAVEN — the finding catalog and the pure scorers: recipient parsing, the
# internal/external test, inbox-rule risk, and the SPF / DMARC verdicts.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'RAVEN mailbox security helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'raven.ps1'
        $RavenFindings             = Get-ToolAssignmentValue -Ast $ast -VarName 'RavenFindings'
        $HiddenFolderPattern       = Get-ToolAssignmentValue -Ast $ast -VarName 'HiddenFolderPattern'
        $SensitiveKeywordPattern   = Get-ToolAssignmentValue -Ast $ast -VarName 'SensitiveKeywordPattern'
        $SuspiciousRuleNamePattern = Get-ToolAssignmentValue -Ast $ast -VarName 'SuspiciousRuleNamePattern'
        foreach ($name in 'Get-SmtpAddressFromRecipient', 'Test-ExternalAddress', 'Get-InboxRuleRisk',
                          'Get-SpfVerdict', 'Get-DmarcVerdict', 'Get-RavenVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $ravenSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'raven.ps1') -Raw
        $internal    = @('contoso.com', 'contoso.onmicrosoft.com')
    }

    Context 'finding catalog' {
        It 'contains every code the tool raises, directly or through a verdict helper' {
            # Codes reach Add-RavenFinding three ways: a literal -Code, a Code
            # field returned by the SPF / DMARC helpers, and the inbox-rule
            # scorer's $codes list. All three must land in the catalog.
            $patterns = @("-Code\s+'([^']+)'", "\bCode\s*=\s*'([A-Za-z]+)'", "\`$codes\.Add\('([^']+)'\)")
            $raised = foreach ($p in $patterns) { [regex]::Matches($ravenSource, $p) | ForEach-Object { $_.Groups[1].Value } }
            $raised = @($raised | Where-Object { $_ } | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 15
            foreach ($code in $raised) {
                $RavenFindings.ContainsKey($code) | Should -BeTrue -Because "raven.ps1 raises '$code'"
            }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $RavenFindings.Keys) {
                $RavenFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $RavenFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $RavenFindings[$code].Summary  | Should -Not -BeNullOrEmpty
                $RavenFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'ranks the compromise indicators as Errors' {
            $RavenFindings['ExternalForwarding'].Severity       | Should -Be 'Error'
            $RavenFindings['InboxRuleExternalForward'].Severity | Should -Be 'Error'
            $RavenFindings['InboxRuleHidesMail'].Severity       | Should -Be 'Error'
        }
    }

    Context 'recipient parsing' {
        It 'extracts the SMTP address from each form Exchange returns' {
            Get-SmtpAddressFromRecipient -Value '"Evil" [SMTP:Evil@Gmail.com]' | Should -BeExactly 'evil@gmail.com'
            Get-SmtpAddressFromRecipient -Value 'smtp:x@example.org'             | Should -BeExactly 'x@example.org'
            Get-SmtpAddressFromRecipient -Value 'plain@contoso.com'              | Should -BeExactly 'plain@contoso.com'
        }
        It 'returns nothing for an internal EX: recipient' {
            Get-SmtpAddressFromRecipient -Value '"Boss" [EX:/o=ExchangeLabs/ou=Exchange/cn=Recipients/cn=abc]' | Should -BeNullOrEmpty
        }
        It 'treats accepted domains and their subdomains as internal' {
            Test-ExternalAddress -Address 'a@contoso.com'    -InternalDomains $internal | Should -BeFalse
            Test-ExternalAddress -Address 'a@eu.contoso.com' -InternalDomains $internal | Should -BeFalse
            Test-ExternalAddress -Address ''                 -InternalDomains $internal | Should -BeFalse
        }
        It 'treats a look-alike domain as external' {
            Test-ExternalAddress -Address 'a@notcontoso.com' -InternalDomains $internal | Should -BeTrue
            Test-ExternalAddress -Address 'a@gmail.com'      -InternalDomains $internal | Should -BeTrue
        }
    }

    Context 'inbox rule risk' {
        It 'flags an external forward as an Error' {
            $rule = [PSCustomObject]@{ Name = 'Fwd'; ForwardTo = @('"X" [SMTP:x@gmail.com]') }
            $r = Get-InboxRuleRisk -Rule $rule -InternalDomains $internal
            $r.Severity | Should -Be 'Error'
            $r.Codes    | Should -Contain 'InboxRuleExternalForward'
        }
        It 'flags moving mail to RSS Feeds and marking it read as hiding mail' {
            $rule = [PSCustomObject]@{ Name = 'News'; MoveToFolder = 'jane:\RSS Feeds'; MarkAsRead = $true }
            (Get-InboxRuleRisk -Rule $rule -InternalDomains $internal).Codes | Should -Contain 'InboxRuleHidesMail'
        }
        It 'flags deleting mail about payments' {
            $rule = [PSCustomObject]@{ Name = 'cleanup'; DeleteMessage = $true; SubjectOrBodyContainsWords = @('invoice') }
            (Get-InboxRuleRisk -Rule $rule -InternalDomains $internal).Codes | Should -Contain 'InboxRuleHidesMail'
        }
        It 'flags a throwaway rule name as a Warning on its own' {
            $r = Get-InboxRuleRisk -Rule ([PSCustomObject]@{ Name = '..' }) -InternalDomains $internal
            $r.Severity | Should -Be 'Warning'
            $r.Codes    | Should -Contain 'InboxRuleSuspicious'
        }
        It 'leaves an ordinary filing rule unflagged' {
            $rule = [PSCustomObject]@{ Name = 'Newsletters'; MoveToFolder = 'Inbox\Newsletters'; ForwardTo = @('"Boss" [SMTP:boss@contoso.com]') }
            (Get-InboxRuleRisk -Rule $rule -InternalDomains $internal).Severity | Should -BeNullOrEmpty
        }
    }

    Context 'SPF verdict' {
        It 'reports a missing record' { (Get-SpfVerdict -Records @('MS=ms1234')).Code | Should -Be 'SpfMissing' }
        It 'treats two SPF records as invalid' { (Get-SpfVerdict -Records @('v=spf1 -all', 'v=spf1 ~all')).Code | Should -Be 'SpfInvalid' }
        It 'treats +all as invalid' { (Get-SpfVerdict -Records @('v=spf1 +all')).Code | Should -Be 'SpfInvalid' }
        It 'treats ?all and a missing all as weak' {
            (Get-SpfVerdict -Records @('v=spf1 mx ?all')).Code | Should -Be 'SpfWeak'
            (Get-SpfVerdict -Records @('v=spf1 mx')).Code      | Should -Be 'SpfWeak'
        }
        It 'accepts ~all, -all and a redirect' {
            (Get-SpfVerdict -Records @('v=spf1 include:spf.protection.outlook.com -all')).Code | Should -BeNullOrEmpty
            (Get-SpfVerdict -Records @('v=spf1 include:spf.protection.outlook.com ~all')).Code | Should -BeNullOrEmpty
            (Get-SpfVerdict -Records @('v=spf1 redirect=_spf.contoso.com')).Code              | Should -BeNullOrEmpty
        }
        It 'notices the Microsoft 365 include' {
            (Get-SpfVerdict -Records @('v=spf1 include:spf.protection.outlook.com -all')).IncludesM365 | Should -BeTrue
        }
    }

    Context 'DMARC verdict' {
        It 'reports a missing record' { (Get-DmarcVerdict -Records @()).Code | Should -Be 'DmarcMissing' }
        It 'treats a record without a valid policy as invalid' { (Get-DmarcVerdict -Records @('v=DMARC1; rua=mailto:a@b.com')).Code | Should -Be 'DmarcInvalid' }
        It 'treats p=none and pct below 100 as monitor-only' {
            (Get-DmarcVerdict -Records @('v=DMARC1; p=none')).Code                | Should -Be 'DmarcMonitorOnly'
            (Get-DmarcVerdict -Records @('v=DMARC1; p=quarantine; pct=50')).Code | Should -Be 'DmarcMonitorOnly'
        }
        It 'accepts an enforcing policy and does not mistake sp= for p=' {
            (Get-DmarcVerdict -Records @('v=DMARC1; p=reject; rua=mailto:a@b.com')).Code | Should -BeNullOrEmpty
            (Get-DmarcVerdict -Records @('v=DMARC1; sp=none; p=reject')).Policy         | Should -Be 'reject'
        }
    }

    Context 'verdict' {
        It 'maps the worst severity to At Risk / Review / Clean' {
            (Get-RavenVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict   | Should -Be 'At Risk'
            (Get-RavenVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' })).Verdict | Should -Be 'Review'
            (Get-RavenVerdict -FindingList @([PSCustomObject]@{ Severity = 'Info' })).Verdict    | Should -Be 'Clean'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# WARD — LAPS policy precedence. Windows LAPS reads the first policy root that
# sets BackupDirectory, so the order of $LapsPolicySources is behaviour.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'WARD LAPS policy resolution' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'ward.ps1'
        $LapsPolicySources = Get-ToolAssignmentValue -Ast $ast -VarName 'LapsPolicySources'
        $LapsBackupTargets = Get-ToolAssignmentValue -Ast $ast -VarName 'LapsBackupTargets'
        . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName 'Resolve-LapsPolicy')))
    }

    It 'lists the policy roots in Windows LAPS precedence order' {
        @($LapsPolicySources.Keys) | Should -Be @('CSP (Intune / MDM)', 'Group Policy', 'Local configuration')
    }
    It 'maps the three BackupDirectory values' {
        $LapsBackupTargets[0] | Should -Be 'Disabled'
        $LapsBackupTargets[1] | Should -Be 'Entra ID'
        $LapsBackupTargets[2] | Should -Be 'Active Directory'
    }
    It 'lets an Intune policy win over a Group Policy one' {
        $src = [ordered]@{ 'CSP (Intune / MDM)' = @{ BackupDirectory = 1; PasswordAgeDays = 14 }; 'Group Policy' = @{ BackupDirectory = 2 }; 'Local configuration' = $null }
        $p = Resolve-LapsPolicy -Sources $src -LegacyPolicy $null -LegacyCseInstalled $false
        $p.Mode            | Should -Be 'Windows LAPS'
        $p.Source          | Should -Be 'CSP (Intune / MDM)'
        $p.BackupDirectory | Should -Be 1
        $p.PasswordAgeDays | Should -Be 14
    }
    It 'skips a root that does not set BackupDirectory' {
        $src = [ordered]@{ 'CSP (Intune / MDM)' = @{ PasswordLength = 20 }; 'Group Policy' = @{ BackupDirectory = 2 }; 'Local configuration' = $null }
        (Resolve-LapsPolicy -Sources $src -LegacyPolicy $null -LegacyCseInstalled $false).Source | Should -Be 'Group Policy'
    }
    It 'reports BackupDirectory 0 as Disabled' {
        $src = [ordered]@{ 'Group Policy' = @{ BackupDirectory = 0 } }
        (Resolve-LapsPolicy -Sources $src -LegacyPolicy $null -LegacyCseInstalled $false).Mode | Should -Be 'Disabled'
    }
    It 'distinguishes legacy LAPS from Windows LAPS emulating it' {
        $none = [ordered]@{ 'Group Policy' = $null }
        (Resolve-LapsPolicy -Sources $none -LegacyPolicy @{ AdmPwdEnabled = 1 } -LegacyCseInstalled $true).Mode  | Should -Be 'Legacy Microsoft LAPS'
        (Resolve-LapsPolicy -Sources $none -LegacyPolicy @{ AdmPwdEnabled = 1 } -LegacyCseInstalled $false).Mode | Should -Be 'Windows LAPS (legacy emulation)'
    }
    It 'reports Not configured when nothing is set' {
        (Resolve-LapsPolicy -Sources ([ordered]@{ 'Group Policy' = $null }) -LegacyPolicy $null -LegacyCseInstalled $false).Mode | Should -Be 'Not configured'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GRIFFIN — platform protection scoring. Kept separate from the AV verdict, so
# these tests pin down its own Hardened / Partial / Not hardened boundaries.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'GRIFFIN platform protection' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'griffin.ps1'
        $DeviceGuardServiceNames = Get-ToolAssignmentValue -Ast $ast -VarName 'DeviceGuardServiceNames'
        $VbsStatusLabels         = Get-ToolAssignmentValue -Ast $ast -VarName 'VbsStatusLabels'
        . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName 'Get-PlatformProtectionVerdict')))

        function Get-TestPlatformVerdict {
            param([hashtable]$Override)
            $p = @{
                Available = $true; VbsStatus = 2; HvciRunning = $true; CredGuardRunning = $true
                CredGuardSupported = $true; LsaPplConfigured = $true; DriverBlocklist = $null
            }
            foreach ($k in $Override.Keys) { $p[$k] = $Override[$k] }
            return Get-PlatformProtectionVerdict @p
        }
    }

    It 'maps the Win32_DeviceGuard service codes for Credential Guard and HVCI' {
        $DeviceGuardServiceNames[1] | Should -Be 'Credential Guard'
        $DeviceGuardServiceNames[2] | Should -Match 'HVCI'
        $VbsStatusLabels[2]         | Should -Be 'Running'
    }
    It 'reports Hardened when every applicable protection runs' {
        (Get-TestPlatformVerdict @{}).Verdict | Should -Be 'Hardened'
    }
    It 'does not count Credential Guard against an edition that cannot run it' {
        $v = Get-TestPlatformVerdict @{ CredGuardSupported = $false; CredGuardRunning = $false }
        $v.Verdict | Should -Be 'Hardened'
        $v.Total   | Should -Be 2
    }
    It 'reports Partial when one protection is off' {
        (Get-TestPlatformVerdict @{ LsaPplConfigured = $false }).Verdict | Should -Be 'Partial'
    }
    It 'reports Not hardened when nothing runs' {
        (Get-TestPlatformVerdict @{ VbsStatus = 0; HvciRunning = $false; CredGuardRunning = $false; LsaPplConfigured = $false }).Verdict | Should -Be 'Not hardened'
    }
    It 'never reports Hardened when VBS state could not be read' {
        (Get-TestPlatformVerdict @{ Available = $false; VbsStatus = $null; HvciRunning = $false; CredGuardRunning = $false }).Verdict | Should -Be 'Partial'
    }
    It 'warns when the vulnerable driver blocklist is explicitly disabled' {
        $v = Get-TestPlatformVerdict @{ DriverBlocklist = 0 }
        $v.Verdict | Should -Be 'Partial'
        @($v.Findings | Where-Object { $_.Text -match 'blocklist' }).Count | Should -Be 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HALO — the finding catalog and the Conditional Access policy predicates.
# Policies are built as nested hashtables, the shape Invoke-MgGraphRequest
# returns, so the predicates are tested against what they will really read.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'HALO Conditional Access helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'halo.ps1'
        $HaloFindings         = Get-ToolAssignmentValue -Ast $ast -VarName 'HaloFindings'
        $PrivilegedRoleTemplates = Get-ToolAssignmentValue -Ast $ast -VarName 'PrivilegedRoleTemplates'
        $GlobalAdminTemplateId   = Get-ToolAssignmentValue -Ast $ast -VarName 'GlobalAdminTemplateId'
        foreach ($name in 'Get-CaList', 'Test-CaPolicyEnforced', 'Test-CaTargetsAllUsers', 'Test-CaTargetsAllApps',
                          'Test-CaRequiresMfa', 'Test-CaBlocks', 'Test-CaBlocksLegacyAuth', 'Test-CaCoversRole',
                          'Get-CaEmergencyExclusion', 'Test-BroadCidr', 'Get-HaloVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $rampartSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'halo.ps1') -Raw

        function New-CaPolicy {
            param([string]$State = 'enabled', [string[]]$Users = @('All'), [string[]]$Exclude = @(), [string[]]$Roles = @(),
                  [string[]]$Apps = @('All'), [string[]]$ClientApps = @('all'), [string[]]$Controls = @('mfa'),
                  [string]$Operator = 'OR', [switch]$Strength)
            $grant = @{ operator = $Operator; builtInControls = $Controls }
            if ($Strength) { $grant['authenticationStrength'] = @{ displayName = 'Phishing-resistant MFA' } }
            return @{
                displayName   = 'Test policy'
                state         = $State
                conditions    = @{
                    users          = @{ includeUsers = $Users; excludeUsers = $Exclude; includeRoles = $Roles }
                    applications   = @{ includeApplications = $Apps }
                    clientAppTypes = $ClientApps
                }
                grantControls = $grant
            }
        }
    }

    Context 'finding catalog' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($rampartSource, "Add-HaloFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value }) +
                      @([regex]::Matches($rampartSource, "_cover\s+'[^']+'\s+\`$\w+\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
            $raised = @($raised | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 10
            foreach ($code in $raised) { $HaloFindings.ContainsKey($code) | Should -BeTrue -Because "halo.ps1 raises '$code'" }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $HaloFindings.Keys) {
                $HaloFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $HaloFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $HaloFindings[$code].Summary  | Should -Not -BeNullOrEmpty
                $HaloFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'ranks the three baseline gaps as Errors' {
            $HaloFindings['NoMfaAllUsers'].Severity        | Should -Be 'Error'
            $HaloFindings['NoMfaAdmins'].Severity          | Should -Be 'Error'
            $HaloFindings['LegacyAuthNotBlocked'].Severity | Should -Be 'Error'
        }
        It 'keys the privileged roles by the Global Administrator template ID' {
            $PrivilegedRoleTemplates[$GlobalAdminTemplateId] | Should -Be 'Global Administrator'
        }
    }

    Context 'policy predicates' {
        It 'counts MFA and authentication strengths as MFA' {
            Test-CaRequiresMfa (New-CaPolicy) | Should -BeTrue
            Test-CaRequiresMfa (New-CaPolicy -Controls @() -Strength) | Should -BeTrue
        }
        It 'does not count "MFA OR compliant device" as MFA' {
            Test-CaRequiresMfa (New-CaPolicy -Controls @('mfa', 'compliantDevice') -Operator 'OR') | Should -BeFalse
            Test-CaRequiresMfa (New-CaPolicy -Controls @('mfa', 'compliantDevice') -Operator 'AND') | Should -BeTrue
        }
        It 'recognises a legacy-authentication block and nothing broader' {
            Test-CaBlocksLegacyAuth (New-CaPolicy -ClientApps @('exchangeActiveSync', 'other') -Controls @('block')) | Should -BeTrue
            Test-CaBlocksLegacyAuth (New-CaPolicy -ClientApps @('all') -Controls @('block')) | Should -BeFalse
            Test-CaBlocksLegacyAuth (New-CaPolicy -ClientApps @('exchangeActiveSync', 'other') -Controls @('mfa')) | Should -BeFalse
        }
        It 'treats report-only as not enforced' {
            Test-CaPolicyEnforced (New-CaPolicy -State 'enabledForReportingButNotEnforced') | Should -BeFalse
            Test-CaPolicyEnforced (New-CaPolicy) | Should -BeTrue
        }
        It 'covers Global Administrator through All users or the role, unless the role is excluded' {
            Test-CaCoversRole -Policy (New-CaPolicy) -RoleTemplateId $GlobalAdminTemplateId | Should -BeTrue
            Test-CaCoversRole -Policy (New-CaPolicy -Users @() -Roles @($GlobalAdminTemplateId)) -RoleTemplateId $GlobalAdminTemplateId | Should -BeTrue
            $p = New-CaPolicy; $p.conditions.users['excludeRoles'] = @($GlobalAdminTemplateId)
            Test-CaCoversRole -Policy $p -RoleTemplateId $GlobalAdminTemplateId | Should -BeFalse
            Test-CaCoversRole -Policy (New-CaPolicy -Apps @('00000002-0000-0ff1-ce00-000000000000')) -RoleTemplateId $GlobalAdminTemplateId | Should -BeFalse
        }
    }

    Context 'emergency access' {
        It 'returns the accounts excluded from every enforcing all-users gate' {
            $gates = @(
                (New-CaPolicy -Exclude @('bg1', 'bg2', 'svc')),
                (New-CaPolicy -ClientApps @('exchangeActiveSync', 'other') -Controls @('block') -Exclude @('bg1', 'bg2'))
            )
            $e = Get-CaEmergencyExclusion -Policies $gates
            $e.GateCount | Should -Be 2
            @($e.Users | Sort-Object) | Should -Be @('bg1', 'bg2')
        }
        It 'ignores report-only policies and non-gating grants' {
            $policies = @(
                (New-CaPolicy -Exclude @('bg1')),
                (New-CaPolicy -State 'enabledForReportingButNotEnforced' -Exclude @()),
                (New-CaPolicy -Controls @('mfa', 'compliantDevice') -Exclude @())
            )
            @((Get-CaEmergencyExclusion -Policies $policies).Users) | Should -Be @('bg1')
        }
        It 'reports an empty exclusion when a gate excludes no one' {
            $e = Get-CaEmergencyExclusion -Policies @((New-CaPolicy -Exclude @('bg1')), (New-CaPolicy -Exclude @()))
            $e.Users.Count | Should -Be 0
        }
        It 'returns nothing when there is no enforcing all-users gate' {
            Get-CaEmergencyExclusion -Policies @((New-CaPolicy -State 'disabled')) | Should -BeNullOrEmpty
        }
    }

    Context 'named locations and verdict' {
        It 'flags IPv4 ranges wider than /16 and IPv6 wider than /32' {
            Test-BroadCidr '10.0.0.0/8'     | Should -BeTrue
            Test-BroadCidr '0.0.0.0/0'      | Should -BeTrue
            Test-BroadCidr '203.0.113.0/24' | Should -BeFalse
            Test-BroadCidr '2001::/16'      | Should -BeTrue
            Test-BroadCidr '2001:db8::/48'  | Should -BeFalse
            Test-BroadCidr 'not-a-range'    | Should -BeFalse
        }
        It 'maps the worst severity to Exposed / Gaps / Enforced' {
            (Get-HaloVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict   | Should -Be 'Exposed'
            (Get-HaloVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' })).Verdict | Should -Be 'Gaps'
            (Get-HaloVerdict -FindingList @([PSCustomObject]@{ Severity = 'Info' })).Verdict    | Should -Be 'Enforced'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CARILLON — the finding catalog and the queue / routing helpers. The whole
# point of the tool is names instead of object IDs, so the name and target
# renderers are pinned here.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'CARILLON call queue helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'carillon.ps1'
        $CarillonFindings  = Get-ToolAssignmentValue -Ast $ast -VarName 'CarillonFindings'
        $DisconnectActions = Get-ToolAssignmentValue -Ast $ast -VarName 'DisconnectActions'
        foreach ($name in 'Get-CarillonList', 'Get-CarillonName', 'Format-CarillonNumber', 'Format-CarillonTarget',
                          'Format-CarillonAction', 'Get-CarillonQueueIssue', 'Get-CarillonAgentSource',
                          'Test-CarillonReachable', 'Get-CarillonVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $carillonSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'carillon.ps1') -Raw

        $jane  = 'bbbbbbbb-0000-0000-0000-000000000001'
        $group = 'cccccccc-0000-0000-0000-000000000001'
        $ra    = 'dddddddd-0000-0000-0000-000000000001'
        $Names = @{
            $jane  = [PSCustomObject]@{ DisplayName = 'Jane Doe';   Upn = 'jane@contoso.com' }
            $group = [PSCustomObject]@{ DisplayName = 'Sales Team'; Upn = '' }
            $ra    = [PSCustomObject]@{ DisplayName = 'RA Sales';   Upn = 'ra-sales@contoso.com' }
        }
    }

    Context 'finding catalog' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($carillonSource, "Add-CarillonFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value }) +
                      @([regex]::Matches($carillonSource, "\[void\]\`$codes\.Add\('([^']+)'\)") | ForEach-Object { $_.Groups[1].Value })
            $raised = @($raised | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 5
            foreach ($code in $raised) { $CarillonFindings.ContainsKey($code) | Should -BeTrue -Because "carillon.ps1 raises '$code'" }
        }
        It 'gives every entry a valid severity and full text' {
            foreach ($code in $CarillonFindings.Keys) {
                $CarillonFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $CarillonFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $CarillonFindings[$code].Summary  | Should -Not -BeNullOrEmpty
                $CarillonFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'treats an unanswerable queue as an error' {
            $CarillonFindings['QueueNoAgents'].Severity    | Should -Be 'Error'
            $CarillonFindings['QueueAllOptedOut'].Severity | Should -Be 'Error'
            $CarillonFindings['StaleReference'].Severity   | Should -Be 'Error'
        }
    }

    Context 'names, never object IDs' {
        It 'labels a resolved user with display name and UPN, case-insensitively' {
            Get-CarillonName -Id $jane.ToUpperInvariant() -Names $Names | Should -Be 'Jane Doe <jane@contoso.com>'
        }
        It 'labels a group by display name alone' {
            Get-CarillonName -Id $group -Names $Names | Should -Be 'Sales Team'
        }
        It 'marks an unresolved ID rather than passing it off as a name' {
            Get-CarillonName -Id 'ffffffff-0000-0000-0000-000000000000' -Names $Names | Should -Match '^\[unresolved '
        }
        It 'flattens null, single and [guid] values' {
            @(Get-CarillonList $null).Count | Should -Be 0
            Get-CarillonList ([guid]$jane) | Should -Be $jane
            @(Get-CarillonList @($jane, '', $group)).Count | Should -Be 2
        }
    }

    Context 'routing targets' {
        It 'strips tel: from external numbers' {
            Format-CarillonTarget -Target @{ Id = 'tel:+15551234567'; Type = 'ExternalPstn' } -Names $Names | Should -Be 'External number +15551234567'
        }
        It 'names the queue behind a resource account' {
            $text = Format-CarillonTarget -Target @{ Id = $ra; Type = 'ApplicationEndpoint' } -Names $Names -Endpoints @{ $ra = 'Call queue Sales' }
            $text | Should -Be 'Call queue Sales (via RA Sales <ra-sales@contoso.com>)'
        }
        It 'names a queue or attendant targeted by identity' {
            Format-CarillonTarget -Target @{ Id = 'AA-1'; Type = 'ConfigurationEndpoint' } -Names $Names -Configs @{ 'aa-1' = 'Auto attendant Main' } | Should -Be 'Auto attendant Main'
        }
        It 'returns nothing for an empty target' {
            Format-CarillonTarget -Target $null -Names $Names | Should -Be ''
        }
        It 'omits the target when the action ends the call' {
            Format-CarillonAction -Action 'DisconnectWithBusy' -TargetText 'User Jane Doe' | Should -Be 'DisconnectWithBusy'
            Format-CarillonAction -Action 'Forward' -TargetText 'User Jane Doe' | Should -Be 'Forward -> User Jane Doe'
        }
    }

    Context 'queue issues' {
        It 'flags a queue with no agents' {
            Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 0; OptedInCount = 0; OverflowThreshold = 50; OverflowAction = 'Voicemail'; TimeoutAction = 'Voicemail' }) |
                Should -Be @('QueueNoAgents')
        }
        It 'flags every agent opted out, and a single opted-in agent' {
            Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 3; OptedInCount = 0; OverflowThreshold = 50 }) | Should -Contain 'QueueAllOptedOut'
            Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 3; OptedInCount = 1; OverflowThreshold = 50 }) | Should -Contain 'QueueFewAgentsOptedIn'
        }
        It 'flags a zero overflow threshold unless the action keeps the call queued' {
            Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 2; OptedInCount = 2; OverflowThreshold = 0; OverflowAction = 'Voicemail' }) | Should -Contain 'OverflowImmediate'
            Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 2; OptedInCount = 2; OverflowThreshold = 0; OverflowAction = 'Queue' }) | Should -Not -Contain 'OverflowImmediate'
        }
        It 'flags queues that hang up' {
            Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 2; OptedInCount = 2; OverflowThreshold = 50; OverflowAction = 'Forward'; TimeoutAction = 'Disconnect' }) | Should -Contain 'CallsDisconnected'
        }
        It 'raises nothing for a healthy queue' {
            @(Get-CarillonQueueIssue ([PSCustomObject]@{ AgentCount = 4; OptedInCount = 3; OverflowThreshold = 50; OverflowAction = 'Voicemail'; TimeoutAction = 'Forward' })).Count | Should -Be 0
        }
    }

    Context 'agent source and reachability' {
        It 'names the group that brought an agent in' {
            $members = @{ $group = [System.Collections.Generic.HashSet[string]]::new([string[]]@($jane), [StringComparer]::OrdinalIgnoreCase) }
            Get-CarillonAgentSource -AgentId $jane -DirectUsers @() -Groups @($group) -GroupMembers $members -Names $Names -Channel $false | Should -Be 'Group: Sales Team'
        }
        It 'reports both a direct add and a group' {
            $members = @{ $group = [System.Collections.Generic.HashSet[string]]::new([string[]]@($jane), [StringComparer]::OrdinalIgnoreCase) }
            Get-CarillonAgentSource -AgentId $jane -DirectUsers @($jane) -Groups @($group) -GroupMembers $members -Names $Names -Channel $false | Should -Be 'Direct; Group: Sales Team'
        }
        It 'falls back to the channel when groups cannot be expanded' {
            Get-CarillonAgentSource -AgentId $jane -DirectUsers @() -Groups @($group) -GroupMembers @{} -Names $Names -Channel $true | Should -Be 'Teams channel'
        }
        It 'is reachable through a numbered resource account or an inbound route' {
            Test-CarillonReachable -ResourceAccountIds @($ra.ToUpperInvariant()) -NumberedAccounts @{ $ra = '+15551234567' } -ReachedFrom @() | Should -BeTrue
            Test-CarillonReachable -ResourceAccountIds @() -NumberedAccounts @{} -ReachedFrom @('Auto attendant Main') | Should -BeTrue
            Test-CarillonReachable -ResourceAccountIds @($ra) -NumberedAccounts @{} -ReachedFrom @() | Should -BeFalse
        }
    }

    Context 'verdict' {
        It 'ranks by the worst finding' {
            (Get-CarillonVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict   | Should -Be 'Broken'
            (Get-CarillonVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' })).Verdict | Should -Be 'Attention'
            (Get-CarillonVerdict -FindingList @([PSCustomObject]@{ Severity = 'Info' })).Verdict    | Should -Be 'Healthy'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# LODESTAR — parsers for nltest and w32tm output, captured from real runs, and the
# Netlogon status table that decides connectivity versus a broken trust.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'LODESTAR domain trust helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'lodestar.ps1'
        $LodestarFindings        = Get-ToolAssignmentValue -Ast $ast -VarName 'LodestarFindings'
        $NetlogonStatusCodes = Get-ToolAssignmentValue -Ast $ast -VarName 'NetlogonStatusCodes'
        foreach ($name in 'ConvertFrom-NltestDsGetDc', 'ConvertFrom-NltestSecureChannel', 'Get-NetlogonStatusInfo',
                          'ConvertFrom-W32tmStripchart', 'Test-PublicIpAddress', 'Get-LodestarVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $oathSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'lodestar.ps1') -Raw
    }

    Context 'finding catalog and status table' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($oathSource, "Add-LodestarFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 8
            foreach ($code in $raised) { $LodestarFindings.ContainsKey($code) | Should -BeTrue -Because "lodestar.ps1 raises '$code'" }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $LodestarFindings.Keys) {
                $LodestarFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $LodestarFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'classifies every Netlogon status as Ok, Connectivity or Trust' {
            foreach ($k in $NetlogonStatusCodes.Keys) {
                $NetlogonStatusCodes[$k].Kind | Should -BeIn @('Ok', 'Connectivity', 'Trust')
            }
            $NetlogonStatusCodes[1789].Kind | Should -Be 'Trust'
            $NetlogonStatusCodes[1311].Kind | Should -Be 'Connectivity'
        }
        It 'falls back to Unknown for a status it does not know' {
            (Get-NetlogonStatusInfo -Code 9999).Kind | Should -Be 'Unknown'
            (Get-NetlogonStatusInfo -Code $null).Kind | Should -Be 'Unknown'
        }
    }

    Context 'nltest parsing' {
        It 'reads the DC, address and site from /dsgetdc' {
            $r = ConvertFrom-NltestDsGetDc -Lines @('           DC: \\DC01.contoso.com', '      Address: \\10.0.0.10',
                ' Dc Site Name: HQ', 'Our Site Name: Branch', 'The command completed successfully')
            $r.Success | Should -BeTrue
            $r.Dc      | Should -Be 'DC01.contoso.com'
            $r.Address | Should -Be '10.0.0.10'
            $r.OurSite | Should -Be 'Branch'
        }
        It 'reads the failure status from /dsgetdc' {
            $r = ConvertFrom-NltestDsGetDc -Lines @('Getting DC name failed: Status = 1355 0x54b ERROR_NO_SUCH_DOMAIN')
            $r.Success    | Should -BeFalse
            $r.StatusCode | Should -Be 1355
        }
        It 'reports a healthy verified channel' {
            $r = ConvertFrom-NltestSecureChannel -Lines @('Trusted DC Name \\DC01.contoso.com',
                'Trusted DC Connection Status Status = 0 0x0 NERR_Success', 'Trust Verification Status = 0 0x0 NERR_Success')
            $r.StatusCode | Should -Be 0
            $r.Verified   | Should -BeTrue
            $r.TrustedDc  | Should -Be 'DC01.contoso.com'
        }
        It 'prefers the verification failure over a healthy connection' {
            $r = ConvertFrom-NltestSecureChannel -Lines @('Trusted DC Name \\DC01.contoso.com',
                'Trusted DC Connection Status Status = 0 0x0 NERR_Success',
                'Trust Verification Status = 1789 0x6fd ERROR_TRUSTED_RELATIONSHIP_FAILURE')
            $r.StatusCode | Should -Be 1789
        }
        It 'reads a connection failure and a bare failed line' {
            (ConvertFrom-NltestSecureChannel -Lines @('Trusted DC Connection Status Status = 1311 0x51f ERROR_NO_LOGON_SERVERS')).StatusCode | Should -Be 1311
            (ConvertFrom-NltestSecureChannel -Lines @('I_NetLogonControl failed: Status = 5 0x5 ERROR_ACCESS_DENIED')).StatusCode | Should -Be 5
        }
        It 'returns no status for output it cannot read' {
            (ConvertFrom-NltestSecureChannel -Lines @('something else entirely')).StatusCode | Should -BeNullOrEmpty
        }
    }

    Context 'time and DNS' {
        It 'reads the offset from w32tm /stripchart /dataonly' {
            ConvertFrom-W32tmStripchart -Lines @('Tracking dc01 [10.0.0.10:123].', '10:00:00, +00.0123456s') | Should -Be 0.0123456
            ConvertFrom-W32tmStripchart -Lines @('10:00:00, -412.5s') | Should -Be -412.5
            ConvertFrom-W32tmStripchart -Lines @('10:00:00, error: 0x800705B4') | Should -BeNullOrEmpty
        }
        It 'recognises public resolvers and leaves private ranges alone' {
            foreach ($a in '8.8.8.8', '1.1.1.1', '2606:4700:4700::1111') { Test-PublicIpAddress $a | Should -BeTrue -Because "$a is public" }
            foreach ($a in '10.0.0.1', '172.20.1.1', '192.168.1.1', '169.254.1.1', '100.64.0.1', '127.0.0.1', 'fd00::1', 'fe80::1', 'junk') {
                Test-PublicIpAddress $a | Should -BeFalse -Because "$a is not a public resolver"
            }
        }
    }

    Context 'verdict' {
        It 'reports Not joined, Broken, Degraded and Healthy' {
            (Get-LodestarVerdict -FindingList @([PSCustomObject]@{ Code = 'NotDomainJoined'; Severity = 'Info' })).Verdict | Should -Be 'Not joined'
            (Get-LodestarVerdict -FindingList @([PSCustomObject]@{ Code = 'TrustBroken'; Severity = 'Error' })).Verdict     | Should -Be 'Broken'
            (Get-LodestarVerdict -FindingList @([PSCustomObject]@{ Code = 'ClockDrift'; Severity = 'Warning' })).Verdict    | Should -Be 'Degraded'
            (Get-LodestarVerdict -FindingList @()).Verdict | Should -Be 'Healthy'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GARM — Security event XML in the shape the DCs log it, the Kerberos / NTLM
# failure-code tables, and the source grouping that turns three events about one
# machine (a 4740 caller name, a 4771 IP, a 4776 \\workstation) into one row.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'GARM lockout tracing helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'garm.ps1'
        $GarmFindings         = Get-ToolAssignmentValue -Ast $ast -VarName 'GarmFindings'
        $KerberosFailureCodes = Get-ToolAssignmentValue -Ast $ast -VarName 'KerberosFailureCodes'
        $NtlmStatusCodes      = Get-ToolAssignmentValue -Ast $ast -VarName 'NtlmStatusCodes'
        foreach ($name in 'ConvertTo-GarmStatusKey', 'Get-GarmStatusInfo', 'ConvertFrom-GarmFileTime', 'Get-GarmNameVariant',
                          'New-GarmEventXPath', 'ConvertFrom-GarmEventXml', 'ConvertTo-GarmSourceHost', 'ConvertTo-GarmAttempt',
                          'Get-GarmSourceKey', 'Get-GarmSourceRanking', 'Test-GarmRunAsMatch', 'ConvertFrom-QuserOutput', 'Get-GarmVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $garmSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'garm.ps1') -Raw

        $ns = 'http://schemas.microsoft.com/win/2004/08/events/event'
        $x4740 = "<Event xmlns='$ns'><System><EventID>4740</EventID><TimeCreated SystemTime='2026-10-08T14:03:11.1234567Z'/><Computer>DC01.contoso.com</Computer></System>" +
                 "<EventData><Data Name='TargetUserName'>jdoe</Data><Data Name='TargetDomainName'>PC-ACCT-07</Data><Data Name='TargetSid'>S-1-5-21-1-2-3-1105</Data>" +
                 "<Data Name='SubjectUserSid'>S-1-5-18</Data><Data Name='SubjectUserName'>DC01`$</Data><Data Name='SubjectDomainName'>CONTOSO</Data><Data Name='SubjectLogonId'>0x3e7</Data></EventData></Event>"
        $x4771 = "<Event xmlns='$ns'><System><EventID>4771</EventID><TimeCreated SystemTime='2026-10-08T14:02:59.0000000Z'/><Computer>DC02.contoso.com</Computer></System>" +
                 "<EventData><Data Name='TargetUserName'>JDoe</Data><Data Name='ServiceName'>krbtgt/CONTOSO</Data><Data Name='Status'>0x18</Data>" +
                 "<Data Name='IpAddress'>::ffff:10.0.4.27</Data><Data Name='IpPort'>51234</Data><Data Name='CertIssuerName'></Data></EventData></Event>"
        $x4776 = "<Event xmlns='$ns'><System><EventID>4776</EventID><TimeCreated SystemTime='2026-10-08T14:01:00.0000000Z'/><Computer>DC01.contoso.com</Computer></System>" +
                 "<EventData><Data Name='PackageName'>MICROSOFT_AUTHENTICATION_PACKAGE_V1_0</Data><Data Name='TargetUserName'>jdoe</Data>" +
                 "<Data Name='Workstation'>\\pc-acct-07</Data><Data Name='Status'>0xc000006a</Data></EventData></Event>"
    }

    Context 'finding catalog and code tables' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($garmSource, "Add-GarmFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 10
            foreach ($code in $raised) { $GarmFindings.ContainsKey($code) | Should -BeTrue -Because "garm.ps1 raises '$code'" }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $GarmFindings.Keys) {
                $GarmFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $GarmFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $GarmFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'marks exactly the wrong-password codes as counting toward a lockout' {
            @($KerberosFailureCodes.Keys | Where-Object { $KerberosFailureCodes[$_].BadPassword }) | Should -Be @('0x18')
            @($NtlmStatusCodes.Keys | Where-Object { $NtlmStatusCodes[$_].BadPassword })      | Should -Be @('0xc000006a')
        }
        It 'keys both tables in normalised form' {
            foreach ($k in @($KerberosFailureCodes.Keys) + @($NtlmStatusCodes.Keys)) { ConvertTo-GarmStatusKey -Status $k | Should -Be $k }
        }
    }

    Context 'status codes and FILETIME' {
        It 'normalises padded and upper-case codes' {
            ConvertTo-GarmStatusKey -Status '0x00000018' | Should -Be '0x18'
            ConvertTo-GarmStatusKey -Status '0XC000006A' | Should -Be '0xc000006a'
            ConvertTo-GarmStatusKey -Status '0x00000000' | Should -Be '0x0'
        }
        It 'decodes known codes and falls back for unknown ones' {
            (Get-GarmStatusInfo -Kind 'Kerberos' -Status '0x18').BadPassword    | Should -BeTrue
            (Get-GarmStatusInfo -Kind 'NTLM' -Status '0xC0000234').Name         | Should -Be 'STATUS_ACCOUNT_LOCKED_OUT'
            (Get-GarmStatusInfo -Kind 'NTLM' -Status '0xc00000ff').BadPassword  | Should -BeFalse
        }
        It 'treats 0 and the never value as not set' {
            ConvertFrom-GarmFileTime -Value 0                   | Should -BeNullOrEmpty
            ConvertFrom-GarmFileTime -Value 9223372036854775807 | Should -BeNullOrEmpty
            ConvertFrom-GarmFileTime -Value $null               | Should -BeNullOrEmpty
            (ConvertFrom-GarmFileTime -Value 134046000000000000).Year | Should -Be 2025
        }
    }

    Context 'event queries' {
        It 'filters by ID and window, and by every case variant of the name' {
            $xp = New-GarmEventXPath -EventId 4771, 4776 -Hours 24 -UserName (Get-GarmNameVariant -SamAccountName 'JDoe' -UserPrincipalName 'jdoe@contoso.com')
            $xp | Should -Match 'EventID=4771 or EventID=4776'
            $xp | Should -Match 'timediff\(@SystemTime\) <= 86400000'
            foreach ($n in 'JDoe', 'jdoe', 'JDOE', 'jdoe@contoso.com', 'JDOE@CONTOSO.COM') { $xp | Should -Match ([regex]::Escape("='$n'")) }
        }
        It 'quotes a name holding an apostrophe with double quotes' {
            New-GarmEventXPath -EventId 4740 -Hours 1 -UserName @("O'Brien") | Should -Match ([regex]::Escape("=`"O'Brien`""))
        }
        It 'does not overflow on the longest window' {
            New-GarmEventXPath -EventId 4740 -Hours 720 | Should -Match '<= 2592000000\]'
        }
    }

    Context 'event parsing' {
        It 'reads the ID, time, DC and data fields' {
            $e = ConvertFrom-GarmEventXml -Xml $x4740
            $e.EventId  | Should -Be 4740
            $e.Computer | Should -Be 'DC01.contoso.com'
            $e.Time.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') | Should -Be '2026-10-08 14:03:11'
            $e.Data['TargetUserName'] | Should -Be 'jdoe'
        }
        It 'takes the 4740 caller from TargetDomainName, not the DC''s own account' {
            $a = ConvertTo-GarmAttempt -Record (ConvertFrom-GarmEventXml -Xml $x4740)
            $a.Kind   | Should -Be 'Lockout'
            $a.Source | Should -Be 'PC-ACCT-07'
        }
        It 'strips the IPv4-mapped prefix from 4771 and decodes the code' {
            $a = ConvertTo-GarmAttempt -Record (ConvertFrom-GarmEventXml -Xml $x4771)
            $a.Source      | Should -Be '10.0.4.27'
            $a.StatusName  | Should -Be 'KDC_ERR_PREAUTH_FAILED'
            $a.BadPassword | Should -BeTrue
        }
        It 'strips the leading \\ from the 4776 workstation' {
            (ConvertTo-GarmAttempt -Record (ConvertFrom-GarmEventXml -Xml $x4776)).Source | Should -Be 'pc-acct-07'
        }
        It 'attributes a loopback source to the DC itself' {
            $e = ConvertFrom-GarmEventXml -Xml ($x4771 -replace '::ffff:10\.0\.4\.27', '::1')
            (ConvertTo-GarmAttempt -Record $e).Source | Should -Be 'DC02.contoso.com'
        }
        It 'maps the empty markers to no source' {
            ConvertTo-GarmSourceHost -Value '-' | Should -Be ''
            ConvertTo-GarmSourceHost -Value ''  | Should -Be ''
        }
    }

    Context 'source ranking' {
        It 'groups a machine''s name, FQDN and workstation forms into one source' {
            $a = @($x4740, $x4771, $x4776 | ForEach-Object { ConvertTo-GarmAttempt -Record (ConvertFrom-GarmEventXml -Xml $_) })
            $a[1].Source = 'pc-acct-07.contoso.com'   # as the collector leaves it after reverse DNS
            $r = @(Get-GarmSourceRanking -Attempts $a)
            $r.Count          | Should -Be 1
            $r[0].Source      | Should -Be 'PC-ACCT-07'
            $r[0].Lockouts    | Should -Be 1
            $r[0].Failures    | Should -Be 2
            $r[0].BadPasswords | Should -Be 2
            $r[0].Dcs.Count   | Should -Be 2
        }
        It 'ranks by lockouts first and keeps unknown sources as their own row' {
            $mk = { param($k, $s) [PSCustomObject]@{ Time = (Get-Date); Dc = 'DC01'; Kind = $k; User = 'jdoe'; Source = $s; Status = ''; StatusName = ''; Meaning = ''; BadPassword = ($k -ne 'Lockout') } }
            $r = @(Get-GarmSourceRanking -Attempts @((& $mk 'NTLM' 'PC1'), (& $mk 'NTLM' 'PC1'), (& $mk 'NTLM' 'PC1'), (& $mk 'Lockout' 'PC2'), (& $mk 'Lockout' '')))
            $r[0].Key | Should -BeIn @('PC2', '')
            $r[-1].Source | Should -Be 'PC1'
            @($r | Where-Object { $_.Source -eq '(unknown)' }).Count | Should -Be 1
        }
        It 'keeps an IP address whole rather than splitting on its dots' {
            Get-GarmSourceKey -Source '10.0.4.27'        | Should -Be '10.0.4.27'
            Get-GarmSourceKey -Source 'pc01.contoso.com' | Should -Be 'PC01'
        }
    }

    Context 'source machine inspection' {
        It 'matches the domain account in each run-as form' {
            $m = @{ SamAccountName = 'jdoe'; UserPrincipalName = 'john.doe@contoso.com'; NetBiosDomain = 'CONTOSO'; DnsDomain = 'contoso.com' }
            Test-GarmRunAsMatch -RunAs 'CONTOSO\jdoe' @m          | Should -BeTrue
            Test-GarmRunAsMatch -RunAs 'contoso.com\JDOE' @m      | Should -BeTrue
            Test-GarmRunAsMatch -RunAs 'john.doe@contoso.com' @m  | Should -BeTrue
            Test-GarmRunAsMatch -RunAs 'jdoe@contoso.com' @m      | Should -BeTrue
            Test-GarmRunAsMatch -RunAs 'jdoe' @m                  | Should -BeTrue
        }
        It 'does not match local, built-in or other-domain accounts' {
            $m = @{ SamAccountName = 'jdoe'; UserPrincipalName = 'jdoe@contoso.com'; NetBiosDomain = 'CONTOSO'; DnsDomain = 'contoso.com' }
            Test-GarmRunAsMatch -RunAs '.\jdoe' @m                       | Should -BeFalse
            Test-GarmRunAsMatch -RunAs 'FABRIKAM\jdoe' @m                | Should -BeFalse
            Test-GarmRunAsMatch -RunAs 'LocalSystem' @m                  | Should -BeFalse
            Test-GarmRunAsMatch -RunAs 'NT AUTHORITY\LocalService' @m    | Should -BeFalse
            Test-GarmRunAsMatch -RunAs '' @m                             | Should -BeFalse
        }
        It 'parses quser rows, including a disconnected session with no session name' {
            $rows = ConvertFrom-QuserOutput -Lines @(
                ' USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME',
                '>jdoe                  console             1  Active      none   10/8/2026 9:00 AM',
                ' jdoe                                      2  Disc         1:02  10/7/2026 4:12 PM',
                ' admin                 rdp-tcp#5           3  Active          .  10/8/2026 8:00 AM')
            $rows.Count    | Should -Be 3
            $rows[1].Id    | Should -Be 2
            $rows[1].State | Should -Be 'Disc'
            $rows[1].Session | Should -Be ''
            $rows[2].Session | Should -Be 'rdp-tcp#5'
        }
        It 'returns nothing for the no-sessions message' {
            @(ConvertFrom-QuserOutput -Lines @('No User exists for *')).Count | Should -Be 0
        }
    }

    Context 'verdict' {
        It 'puts a found culprit above a traced source above an untraced lockout' {
            (Get-GarmVerdict -FindingList @([PSCustomObject]@{ Code = 'TaskRunsAsUser'; Severity = 'Error' }, [PSCustomObject]@{ Code = 'StaleCredentialSource'; Severity = 'Warning' })).Verdict | Should -Be 'Culprit found'
            (Get-GarmVerdict -FindingList @([PSCustomObject]@{ Code = 'StaleCredentialSource'; Severity = 'Warning' })).Verdict | Should -Be 'Source traced'
            (Get-GarmVerdict -FindingList @([PSCustomObject]@{ Code = 'AccountLockedOut'; Severity = 'Error' })).Verdict       | Should -Be 'Untraced'
            (Get-GarmVerdict -FindingList @([PSCustomObject]@{ Code = 'AccountNotFound'; Severity = 'Error' })).Verdict        | Should -Be 'Not found'
            (Get-GarmVerdict -FindingList @()).Verdict | Should -Be 'Quiet'
        }
        It 'grades a sweep by severity' {
            (Get-GarmVerdict -Mode 'Sweep' -FindingList @([PSCustomObject]@{ Code = 'AccountLockedOut'; Severity = 'Error' })).Verdict  | Should -Be 'Accounts locked'
            (Get-GarmVerdict -Mode 'Sweep' -FindingList @([PSCustomObject]@{ Code = 'EventLogUnreadable'; Severity = 'Warning' })).Verdict | Should -Be 'Unreadable'
            (Get-GarmVerdict -Mode 'Sweep' -FindingList @()).Verdict | Should -Be 'Quiet'
        }
    }

    Context 'SPHINX lockout lookup' {
        It 'reads the 4740 caller computer from index 1 (TargetDomainName), not index 4' {
            $sphinx = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'sphinx.ps1') -Raw
            $sphinx | Should -Match '\$callerMachine\s*=\s*\$evt\.Properties\[1\]\.Value'
            $sphinx | Should -Not -Match '\$callerMachine\s*=\s*\$evt\.Properties\[4\]'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TORPOR — the usage arithmetic that turns two snapshots into per-process and
# per-disk load, the parsers, and the finding catalog.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'TORPOR slow-machine helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'torpor.ps1'
        $TorporFindings       = Get-ToolAssignmentValue -Ast $ast -VarName 'TorporFindings'
        $ProcessHints         = Get-ToolAssignmentValue -Ast $ast -VarName 'ProcessHints'
        $BootDegradationKinds = Get-ToolAssignmentValue -Ast $ast -VarName 'BootDegradationKinds'
        foreach ($name in 'Get-TorporAverage', 'Get-TorporLevel', 'Get-ProcessHint', 'Get-ProcessCpuUsage', 'Group-TorporProcess',
                          'Get-DiskRawDelta', 'ConvertFrom-PowercfgScheme', 'Get-PowerModeLabel', 'Test-StartupApprovedEnabled',
                          'Get-BootDegradationKind', 'Get-TorporVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $torporSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'torpor.ps1') -Raw

        function New-Snap { param($Id, $Name, $Cpu, $Io = 0, $Private = 1MB) [PSCustomObject]@{ Id = $Id; Name = $Name; CpuSeconds = $Cpu; WorkingSet = $Private; Private = $Private; IoBytes = $Io } }
    }

    Context 'finding catalog' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($torporSource, "Add-TorporFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 15
            foreach ($code in $raised) { $TorporFindings.ContainsKey($code) | Should -BeTrue -Because "torpor.ps1 raises '$code'" }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $TorporFindings.Keys) {
                $TorporFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $TorporFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $TorporFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
    }

    Context 'process usage' {
        It 'reports CPU as a share of the whole machine' {
            $rows = @(Get-ProcessCpuUsage -Before @(New-Snap 10 'app' 5) -After @(New-Snap 10 'app' 15) -ElapsedSeconds 10 -LogicalProcessors 4)
            $rows[0].CpuPercent | Should -Be 25
        }
        It 'credits a process that started in the window with all of its time' {
            $rows = @(Get-ProcessCpuUsage -Before @() -After @(New-Snap 20 'new' 2 -Io 1000) -ElapsedSeconds 10 -LogicalProcessors 1)
            $rows[0].CpuPercent    | Should -Be 20
            $rows[0].IoBytesPerSec | Should -Be 100
        }
        It 'does not credit a reused PID with the old process''s time' {
            $rows = @(Get-ProcessCpuUsage -Before @(New-Snap 30 'old' 100) -After @(New-Snap 30 'other' 1) -ElapsedSeconds 10 -LogicalProcessors 1)
            $rows[0].CpuPercent | Should -Be 10
        }
        It 'never reports negative usage and tolerates unreadable CPU time' {
            $rows = @(Get-ProcessCpuUsage -Before @(New-Snap 40 'p' 50 -Io 500) -After @(New-Snap 40 'p' 40 -Io 100), (New-Snap 41 'q' $null) -ElapsedSeconds 10 -LogicalProcessors 2)
            $rows[0].CpuPercent    | Should -Be 0
            $rows[0].IoBytesPerSec | Should -Be 0
            $rows[1].CpuPercent    | Should -Be 0
        }
        It 'groups processes by name and sums their usage' {
            $rows = @(
                [PSCustomObject]@{ Id = 1; Name = 'msedge'; CpuPercent = 10; WorkingSet = 100; Private = 200; IoBytesPerSec = 5 }
                [PSCustomObject]@{ Id = 2; Name = 'msedge'; CpuPercent = 5.5; WorkingSet = 100; Private = 300; IoBytesPerSec = 0 }
                [PSCustomObject]@{ Id = 3; Name = 'svchost'; CpuPercent = 1; WorkingSet = 10; Private = 10; IoBytesPerSec = 0 }
            )
            $edge = @(Group-TorporProcess -Rows $rows) | Where-Object { $_.Name -eq 'msedge' }
            $edge.Count      | Should -Be 2
            $edge.CpuPercent | Should -Be 15.5
            $edge.Private    | Should -Be 500
            $edge.Ids        | Should -Be @(1, 2)
        }
        It 'hints at well-known processes regardless of case or .exe' {
            Get-ProcessHint -Name 'MsMpEng.exe' | Should -Match 'Defender'
            Get-ProcessHint -Name 'SearchIndexer' | Should -Match 'Search'
            Get-ProcessHint -Name 'contoso-lob' | Should -BeNullOrEmpty
            foreach ($k in $ProcessHints.Keys) { $k | Should -BeExactly $k.ToLowerInvariant() -Because 'lookups are lower-case' }
        }
    }

    Context 'disk usage from raw counters' {
        BeforeAll {
            $before = [PSCustomObject]@{ Frequency_PerfTime = 1e7; Timestamp_PerfTime = 0; Timestamp_Sys100NS = 0; PercentIdleTime = 0
                                         AvgDisksecPerTransfer = 0; AvgDisksecPerTransfer_Base = 0; DiskTransfersPersec = 0; DiskBytesPersec = 0 }
        }
        It 'derives busy time, response time, IOPS and throughput' {
            # 10 s window, idle 40% of it, 100 transfers taking 0.2 s in total.
            $after = [PSCustomObject]@{ Frequency_PerfTime = 1e7; Timestamp_PerfTime = 1e8; Timestamp_Sys100NS = 1e8; PercentIdleTime = 4e7
                                        AvgDisksecPerTransfer = 2e6; AvgDisksecPerTransfer_Base = 100; DiskTransfersPersec = 1000; DiskBytesPersec = 1e7 }
            $d = Get-DiskRawDelta -Before $before -After $after
            $d.BusyPercent | Should -Be 60
            $d.LatencyMs   | Should -Be 2
            $d.Iops        | Should -Be 100
            $d.BytesPerSec | Should -Be 1e6
        }
        It 'clamps busy time to 0-100 and reports no latency without transfers' {
            $after = [PSCustomObject]@{ Frequency_PerfTime = 1e7; Timestamp_PerfTime = 1e8; Timestamp_Sys100NS = 1e8; PercentIdleTime = 1.2e8
                                        AvgDisksecPerTransfer = 0; AvgDisksecPerTransfer_Base = 0; DiskTransfersPersec = 0; DiskBytesPersec = 0 }
            $d = Get-DiskRawDelta -Before $before -After $after
            $d.BusyPercent | Should -Be 0
            $d.LatencyMs   | Should -Be 0
        }
    }

    Context 'parsers and levels' {
        It 'reads the active power scheme GUID and label' {
            $s = ConvertFrom-PowercfgScheme -Lines @('Power Scheme GUID: A1841308-3541-4FAB-BC81-F71556F20B4A  (Power saver)')
            $s.Guid | Should -Be 'a1841308-3541-4fab-bc81-f71556f20b4a'
            $s.Name | Should -Be 'Power saver'
            (ConvertFrom-PowercfgScheme -Lines @('nothing here')).Guid | Should -BeNullOrEmpty
        }
        It 'names the Windows power modes' {
            Get-PowerModeLabel -Guid '961cc777-2547-4f9d-8174-7d86181b8a7a' | Should -Be 'Best power efficiency'
            Get-PowerModeLabel -Guid '' | Should -BeNullOrEmpty
            Get-PowerModeLabel -Guid 'abc' | Should -Match 'Unrecognised'
        }
        It 'reads Task Manager startup state' {
            Test-StartupApprovedEnabled -Bytes ([byte[]](2, 0, 0)) | Should -BeTrue
            Test-StartupApprovedEnabled -Bytes ([byte[]](3, 0, 0)) | Should -BeFalse
            Test-StartupApprovedEnabled -Bytes ([byte[]](7, 0)) | Should -BeFalse
            Test-StartupApprovedEnabled -Bytes $null | Should -BeTrue
        }
        It 'classifies boot degradation events and falls back to Other' {
            Get-BootDegradationKind -EventId 102 | Should -Be 'Driver'
            Get-BootDegradationKind -EventId 999 | Should -Be 'Other'
            $BootDegradationKinds.ContainsKey(100) | Should -BeFalse -Because 'event 100 is the boot itself, not a component'
        }
        It 'grades values in both directions and averages only real samples' {
            Get-TorporLevel -Value 95 -Warning 70 -ErrorAt 90 | Should -Be 'Error'
            Get-TorporLevel -Value 75 -Warning 70 -ErrorAt 90 | Should -Be 'Warning'
            Get-TorporLevel -Value 12 -Warning 15 -ErrorAt 5 -LowerIsWorse | Should -Be 'Warning'
            Get-TorporLevel -Value 4 -Warning 15 -ErrorAt 5 -LowerIsWorse | Should -Be 'Error'
            Get-TorporLevel -Value $null -Warning 1 -ErrorAt 2 | Should -Be 'Ok'
            Get-TorporAverage -Values @(10, $null, 20) | Should -Be 15
            Get-TorporAverage -Values @() | Should -BeNullOrEmpty
        }
    }

    Context 'verdict' {
        It 'reports Struggling, Strained and Healthy' {
            (Get-TorporVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict | Should -Be 'Struggling'
            (Get-TorporVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' }, [PSCustomObject]@{ Severity = 'Info' })).Verdict | Should -Be 'Strained'
            (Get-TorporVerdict -FindingList @([PSCustomObject]@{ Severity = 'Info' })).Verdict | Should -Be 'Healthy'
            (Get-TorporVerdict -FindingList @()).Verdict | Should -Be 'Healthy'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SOLDER — the DISM / CBS.log / SFC parsers, the servicing error-code table
# that decides which tool fixes a failed update, and the finding catalog.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'SOLDER servicing helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'solder.ps1'
        $SolderFindings      = Get-ToolAssignmentValue -Ast $ast -VarName 'SolderFindings'
        $ServicingErrorCodes = Get-ToolAssignmentValue -Ast $ast -VarName 'ServicingErrorCodes'
        foreach ($name in 'ConvertTo-HResultString', 'Get-ServicingErrorInfo', 'ConvertFrom-DismHealth', 'ConvertFrom-DismAnalyze',
                          'Get-CbsQuotedPath', 'ConvertFrom-CbsSrLine', 'ConvertFrom-SfcOutput', 'Get-PendingRebootReason',
                          'Get-RepairSourceState', 'Get-SolderVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $sutureSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'solder.ps1') -Raw
    }

    Context 'finding catalog and error-code table' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($sutureSource, "Add-SolderFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 12
            foreach ($code in $raised) { $SolderFindings.ContainsKey($code) | Should -BeTrue -Because "solder.ps1 raises '$code'" }
            # The update-failure codes are chosen by expression, not a literal.
            foreach ($code in 'UpdateFailuresStore', 'UpdateFailuresClient') { $SolderFindings.ContainsKey($code) | Should -BeTrue }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $SolderFindings.Keys) {
                $SolderFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $SolderFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $SolderFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'keys the error-code table by normalised hex and classifies every entry' {
            foreach ($k in $ServicingErrorCodes.Keys) {
                $k | Should -Match '^0x[0-9a-f]{8}$'
                $ServicingErrorCodes[$k].Kind | Should -BeIn @('Store', 'Client', 'Reboot', 'Space', 'Access', 'Other')
            }
            $ServicingErrorCodes['0x800f081f'].Kind | Should -Be 'Store'
            $ServicingErrorCodes['0x8024402c'].Kind | Should -Be 'Client'
        }
    }

    Context 'error codes' {
        It 'normalises hex, bare hex and signed exit codes to one form' {
            ConvertTo-HResultString -Code '0x800F081F' | Should -Be '0x800f081f'
            ConvertTo-HResultString -Code '800f081f'   | Should -Be '0x800f081f'
            ConvertTo-HResultString -Code -2146498529  | Should -Be '0x800f081f'
            ConvertTo-HResultString -Code '0x5'        | Should -Be '0x00000005'
            ConvertTo-HResultString -Code 87           | Should -Be '0x00000057'
            ConvertTo-HResultString -Code 'nonsense'   | Should -BeNullOrEmpty
            ConvertTo-HResultString -Code $null        | Should -BeNullOrEmpty
        }
        It 'decodes a known code and falls back for an unknown one' {
            (Get-ServicingErrorInfo -Code -2146498529).Name | Should -Be 'CBS_E_SOURCE_MISSING'
            (Get-ServicingErrorInfo -Code '0x80070422').Kind | Should -Be 'Client'
            (Get-ServicingErrorInfo -Code '0x80001234').Kind | Should -Be 'Other'
            (Get-ServicingErrorInfo -Code $null).Name | Should -Be 'Unknown'
        }
    }

    Context 'DISM parsing' {
        It 'reads each component store state' {
            (ConvertFrom-DismHealth -Lines @('No component store corruption detected.', 'The operation completed successfully.')).State | Should -Be 'Healthy'
            (ConvertFrom-DismHealth -Lines @('The component store is repairable.')).State | Should -Be 'Repairable'
            (ConvertFrom-DismHealth -Lines @('The component store cannot be repaired.')).State | Should -Be 'NotRepairable'
            (ConvertFrom-DismHealth -Lines @('The restore operation completed successfully.')).State | Should -Be 'Repaired'
            (ConvertFrom-DismHealth -Lines @('something else')).State | Should -Be 'Unknown'
        }
        It 'captures the error code DISM prints' {
            (ConvertFrom-DismHealth -Lines @('Error: 0x800f081f', '', 'The source files could not be found.')).ErrorCode | Should -Be '0x800f081f'
            (ConvertFrom-DismHealth -Lines @('Error: 87')).ErrorCode | Should -Be '0x00000057'
        }
        It 'reads /AnalyzeComponentStore' {
            $a = ConvertFrom-DismAnalyze -Lines @('Actual Size of Component Store : 8.15 GB', 'Date of Last Cleanup : 2026-09-01 10:11:12',
                                                  'Number of Reclaimable Packages : 3', 'Component Store Cleanup Recommended : Yes')
            $a.ActualSize          | Should -Be '8.15 GB'
            $a.ReclaimablePackages | Should -Be 3
            $a.CleanupRecommended  | Should -BeTrue
            (ConvertFrom-DismAnalyze -Lines @('Component Store Cleanup Recommended : No')).CleanupRecommended | Should -BeFalse
            (ConvertFrom-DismAnalyze -Lines @()).ReclaimablePackages | Should -BeNullOrEmpty
        }
    }

    Context 'System File Checker' {
        It 'reads repaired and unrepairable files in both CBS path formats' {
            $r = ConvertFrom-CbsSrLine -Lines @(
                '2026-10-01 10:00:00, Info  CSI  00000001 [SR] Verifying 100 components'
                '2026-10-01 10:00:01, Info  CSI  00000002 [SR] Cannot repair member file [l:34]"msvcp_win.dll" of Microsoft-Windows-CoreSystem, version 10.0'
                '2026-10-01 10:00:02, Info  CSI  00000003 [SR] Repairing corrupted file [ml:520{260},l:66{33}]"\??\C:\WINDOWS\System32\drivers"\[l:22{11}]"netio.sys" from store'
                '2026-10-01 10:00:03, Info  CSI  00000004 [SR] Repairing corrupted file \??\C:\Windows\System32\foo.dll from store'
                '2026-10-01 10:00:04, Info  CSI  00000005 [SR] Could not reproject corrupted file [l:40]"\??\C:\Windows\System32"\[l:14]"bar.dll"; source file in store is also corrupted'
                '2026-10-01 10:00:05, Info  CSI  00000006 [SR] Cannot repair member file [l:34]"msvcp_win.dll" of Microsoft-Windows-CoreSystem, version 10.0'
                'an unrelated CBS line'
            )
            $r.Found        | Should -BeTrue
            $r.Repaired     | Should -Be @('C:\WINDOWS\System32\drivers\netio.sys', 'C:\Windows\System32\foo.dll')
            $r.CannotRepair | Should -Be @('msvcp_win.dll', 'C:\Windows\System32\bar.dll')
            $r.LastRun      | Should -Be '2026-10-01 10:00:05'
        }
        It 'reports no SFC run when the log has no [SR] lines' {
            $r = ConvertFrom-CbsSrLine -Lines @('2026-10-01 10:00:00, Info  CBS  Session started')
            $r.Found | Should -BeFalse
            @($r.Repaired).Count | Should -Be 0
        }
        It 'reads sfc console output through its UTF-16 NULs' {
            $nul = [char]0
            $spread = { param($s) ($s.ToCharArray() | ForEach-Object { "$_$nul" }) -join '' }
            ConvertFrom-SfcOutput -Lines @(& $spread 'Windows Resource Protection did not find any integrity violations.') | Should -Be 'Clean'
            ConvertFrom-SfcOutput -Lines @('Windows Resource Protection found corrupt files but was unable to fix some of them.') | Should -Be 'Unrepaired'
            ConvertFrom-SfcOutput -Lines @('Windows Resource Protection found corrupt files and successfully repaired them.') | Should -Be 'Repaired'
            ConvertFrom-SfcOutput -Lines @('') | Should -Be 'Unknown'
        }
    }

    Context 'pending restarts and repair source' {
        It 'separates servicing restarts from the rest' {
            $r = @(Get-PendingRebootReason -Signals @{ CbsRebootPending = $true; PendingXml = $false; FileRenames = $true; ComputerRename = $false })
            $r.Count | Should -Be 2
            ($r | Where-Object { $_.Signal -eq 'CbsRebootPending' }).Kind | Should -Be 'Servicing'
            ($r | Where-Object { $_.Signal -eq 'FileRenames' }).Kind | Should -Be 'Other'
            @(Get-PendingRebootReason -Signals @{}).Count | Should -Be 0
        }
        It 'flags a WSUS machine with no repair source' {
            (Get-RepairSourceState -WsusConfigured $true -RepairContentServerSource $null -LocalSourcePath '' -UseWindowsUpdate $null).AtRisk | Should -BeTrue
            (Get-RepairSourceState -WsusConfigured $true -RepairContentServerSource 2 -LocalSourcePath '' -UseWindowsUpdate $null).AtRisk | Should -BeFalse
            (Get-RepairSourceState -WsusConfigured $true -RepairContentServerSource $null -LocalSourcePath '\\srv\winsxs' -UseWindowsUpdate $null).AtRisk | Should -BeFalse
            (Get-RepairSourceState -WsusConfigured $false -RepairContentServerSource $null -LocalSourcePath '' -UseWindowsUpdate 2).AtRisk | Should -BeTrue
            (Get-RepairSourceState -WsusConfigured $false -RepairContentServerSource $null -LocalSourcePath '' -UseWindowsUpdate $null).AtRisk | Should -BeFalse
        }
    }

    Context 'verdict' {
        It 'reports Broken, Attention and Healthy' {
            (Get-SolderVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict | Should -Be 'Broken'
            (Get-SolderVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' })).Verdict | Should -Be 'Attention'
            (Get-SolderVerdict -FindingList @([PSCustomObject]@{ Severity = 'Info' })).Verdict | Should -Be 'Healthy'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CHALICE — product-family and support-date mapping, channel resolution, the
# dsregcmd / cmdkey parsers, identity classification, the finding catalog,
# and the per-user contract (no admin gate).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'CHALICE Microsoft 365 Apps helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'chalice.ps1'
        $ChaliceFindings    = Get-ToolAssignmentValue -Ast $ast -VarName 'ChaliceFindings'
        $ChannelGuids       = Get-ToolAssignmentValue -Ast $ast -VarName 'ChannelGuids'
        $UpdateBranchNames  = Get-ToolAssignmentValue -Ast $ast -VarName 'UpdateBranchNames'
        $OfficeEndOfSupport = Get-ToolAssignmentValue -Ast $ast -VarName 'OfficeEndOfSupport'
        foreach ($name in 'Get-OfficeProductFamily', 'Get-OfficeSupportState', 'Get-ChannelName', 'Test-UpdatesDisabled',
                          'ConvertFrom-DsregcmdStatus', 'ConvertFrom-DsregTime', 'Get-DeviceJoinState', 'ConvertFrom-CmdkeyList',
                          'ConvertTo-IdentityProvider', 'Get-IdentitySummary', 'Get-ChaliceVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $chaliceSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'chalice.ps1') -Raw
    }

    Context 'finding catalog and per-user contract' {
        It 'contains every code the tool raises' {
            $raised = @([regex]::Matches($chaliceSource, "Add-ChaliceFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            $raised.Count | Should -BeGreaterThan 15
            foreach ($code in $raised) { $ChaliceFindings.ContainsKey($code) | Should -BeTrue -Because "chalice.ps1 raises '$code'" }
        }
        It 'gives every finding a renderable Severity, Title, Summary and Remedy' {
            foreach ($code in $ChaliceFindings.Keys) {
                $ChaliceFindings[$code].Severity | Should -BeIn @('Error', 'Warning', 'Info')
                $ChaliceFindings[$code].Title    | Should -Not -BeNullOrEmpty
                $ChaliceFindings[$code].Remedy   | Should -Not -BeNullOrEmpty
            }
        }
        It 'never names an admin gate, so it runs as the signed-in user' {
            # ToolTraits in the desktop app detects the gate by substring, so
            # even a comment naming it would make the app elevate CHALICE.
            $chaliceSource | Should -Not -Match 'Invoke-AdminElevation|Assert-AdminPrivilege'
        }
    }

    Context 'products, support and channel' {
        It 'maps Click-to-Run release IDs to product families' {
            Get-OfficeProductFamily -ReleaseId 'O365ProPlusRetail'  | Should -Be 'Microsoft 365 Apps'
            Get-OfficeProductFamily -ReleaseId 'O365BusinessRetail' | Should -Be 'Microsoft 365 Apps'
            Get-OfficeProductFamily -ReleaseId 'ProPlus2024Volume'  | Should -Be 'Office 2024'
            Get-OfficeProductFamily -ReleaseId 'ProPlus2021Volume'  | Should -Be 'Office 2021'
            Get-OfficeProductFamily -ReleaseId 'VisioPro2019Retail' | Should -Be 'Office 2019'
            Get-OfficeProductFamily -ReleaseId 'ProjectProXVolume'  | Should -Be 'Office 2016'
            Get-OfficeProductFamily -ReleaseId 'ProPlusRetail'      | Should -Be 'Office 2016'
            Get-OfficeProductFamily -ReleaseId 'SomethingElse'      | Should -Be 'Other'
            Get-OfficeProductFamily -ReleaseId ''                   | Should -BeNullOrEmpty
        }
        It 'dates support from the table, with no end for the subscription' {
            (Get-OfficeSupportState -Family 'Office 2019' -Today ([datetime]'2026-01-01')).State | Should -Be 'Ended'
            (Get-OfficeSupportState -Family 'Office 2021' -Today ([datetime]'2026-10-08')).State | Should -Be 'EndingSoon'
            (Get-OfficeSupportState -Family 'Office 2021' -Today ([datetime]'2026-10-13')).State | Should -Be 'EndingSoon'
            (Get-OfficeSupportState -Family 'Office 2021' -Today ([datetime]'2026-10-14')).State | Should -Be 'Ended'
            (Get-OfficeSupportState -Family 'Office 2024' -Today ([datetime]'2026-10-08')).State | Should -Be 'Supported'
            (Get-OfficeSupportState -Family 'Microsoft 365 Apps' -Today ([datetime]'2040-01-01')).State | Should -Be 'Supported'
            foreach ($k in $OfficeEndOfSupport.Keys) { $OfficeEndOfSupport[$k] | Should -Match '^\d{4}-\d{2}-\d{2}$' }
        }
        It 'resolves the channel from the CDN URL, and policy wins' {
            Get-ChannelName -Url 'http://officecdn.microsoft.com/pr/492350F6-3A01-4F97-B9C0-C7C6DDF67D60' | Should -Be 'Current Channel'
            Get-ChannelName -Url 'http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6' -PolicyBranch 'Deferred' | Should -Be 'Semi-Annual Enterprise Channel (policy)'
            Get-ChannelName -Url 'http://officecdn.microsoft.com/pr/00000000-0000-0000-0000-000000000000' | Should -Match 'Unrecognised'
            Get-ChannelName -Url '\\server\office' | Should -Be 'Custom source'
            Get-ChannelName -Url '' | Should -Be 'Unknown'
            foreach ($k in $ChannelGuids.Keys) { $k | Should -BeExactly $k.ToLowerInvariant() }
            foreach ($k in $UpdateBranchNames.Keys) { $k | Should -BeExactly $k.ToLowerInvariant() }
        }
        It 'detects updates turned off locally or by policy' {
            Test-UpdatesDisabled -UpdatesEnabled 'False' -PolicyEnable $null | Should -BeTrue
            Test-UpdatesDisabled -UpdatesEnabled 'True'  -PolicyEnable 0     | Should -BeTrue
            Test-UpdatesDisabled -UpdatesEnabled 'True'  -PolicyEnable 1     | Should -BeFalse
            Test-UpdatesDisabled -UpdatesEnabled $null   -PolicyEnable $null | Should -BeFalse
        }
    }

    Context 'device, credentials and identities' {
        It 'reads join state and the PRT from dsregcmd /status' {
            $s = ConvertFrom-DsregcmdStatus -Lines @(
                '+----------------------------------------------------------------------+'
                '| Device State                                                         |'
                '             AzureAdJoined : YES'
                '              DomainJoined : NO'
                '                TenantName : Contoso'
                '                AzureAdPrt : NO'
                '      AzureAdPrtUpdateTime : 2026-10-08 13:10:12.000 UTC'
            )
            $j = Get-DeviceJoinState -Status $s
            $j.Label       | Should -Be 'Entra joined'
            $j.EntraJoined | Should -BeTrue
            $j.HasPrt      | Should -BeFalse
            $j.Tenant      | Should -Be 'Contoso'
            $j.PrtUpdated.ToUniversalTime().Hour | Should -Be 13
        }
        It 'labels hybrid, registered and unjoined devices' {
            (Get-DeviceJoinState -Status @{ AzureAdJoined = 'YES'; DomainJoined = 'YES' }).Label | Should -Be 'Hybrid Entra joined'
            (Get-DeviceJoinState -Status @{ WorkplaceJoined = 'YES' }).Label | Should -Be 'Entra registered'
            (Get-DeviceJoinState -Status @{ DomainJoined = 'YES' }).Label | Should -Be 'Domain joined only'
            (Get-DeviceJoinState -Status @{}).Label | Should -Be 'Not joined'
            ConvertFrom-DsregTime -Text 'garbage' | Should -BeNullOrEmpty
        }
        It 'finds Office credentials whatever the Target label is translated to' {
            $c = @(ConvertFrom-CmdkeyList -Lines @(
                '    Target: LegacyGeneric:target=MicrosoftOffice16_Data:SSPI:jane@contoso.com'
                '    Target: LegacyGeneric:target=OneDrive Cached Credential'
                '    Ziel: LegacyGeneric:target=MicrosoftOffice15_Data:ADAL:abc'
                '    Target: LegacyGeneric:target=MicrosoftOffice16_Data:SSPI:jane@contoso.com'
            ))
            $c | Should -Be @('LegacyGeneric:target=MicrosoftOffice16_Data:SSPI:jane@contoso.com', 'LegacyGeneric:target=MicrosoftOffice15_Data:ADAL:abc')
        }
        It 'classifies identities by ProviderId or key suffix' {
            ConvertTo-IdentityProvider -ProviderId 'AD' -KeyName 'x' | Should -Be 'AD'
            ConvertTo-IdentityProvider -ProviderId '' -KeyName 'abc_ADAL' | Should -Be 'AD'
            ConvertTo-IdentityProvider -ProviderId 'LiveId' -KeyName 'x' | Should -Be 'MSA'
            ConvertTo-IdentityProvider -ProviderId '' -KeyName '123_LiveId' | Should -Be 'MSA'
            ConvertTo-IdentityProvider -ProviderId 'Other' -KeyName 'x' | Should -Be 'Other'
        }
        It 'summarises personal-only, mixed and multi-tenant caches' {
            $work1 = [PSCustomObject]@{ Provider = 'AD'; TenantId = 'T1' }
            $work2 = [PSCustomObject]@{ Provider = 'AD'; TenantId = 'T2' }
            $msa   = [PSCustomObject]@{ Provider = 'MSA'; TenantId = '' }
            (Get-IdentitySummary -Identities @($msa)).PersonalOnly | Should -BeTrue
            (Get-IdentitySummary -Identities @($work1)).Mixed | Should -BeFalse
            (Get-IdentitySummary -Identities @($work1, $msa)).Mixed | Should -BeTrue
            $multi = Get-IdentitySummary -Identities @($work1, $work2)
            $multi.Tenants | Should -Be 2
            $multi.Mixed   | Should -BeTrue
            (Get-IdentitySummary -Identities @()).Total | Should -Be 0
        }
    }

    Context 'verdict' {
        It 'reports Broken, Attention and Healthy' {
            (Get-ChaliceVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict | Should -Be 'Broken'
            (Get-ChaliceVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' })).Verdict | Should -Be 'Attention'
            (Get-ChaliceVerdict -FindingList @()).Verdict | Should -Be 'Healthy'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MINOTAUR — the rights-mask collapse and SID categories that decide what
# counts as broad write access, and the finding catalog.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'MINOTAUR permission helpers' {
    BeforeAll {
        $ast = Get-ToolAst -FileName 'minotaur.ps1'
        $MinotaurFindings = Get-ToolAssignmentValue -Ast $ast -VarName 'MinotaurFindings'
        $BroadSids        = Get-ToolAssignmentValue -Ast $ast -VarName 'BroadSids'
        $BroadDomainRids  = Get-ToolAssignmentValue -Ast $ast -VarName 'BroadDomainRids'
        foreach ($name in 'Get-NtfsRightsLevel', 'Test-WriteLevel', 'Get-SidCategory', 'Get-MinotaurVerdict') {
            . ([scriptblock]::Create((Get-ToolFunctionText -Ast $ast -FuncName $name)))
        }
        $catacombSource = Get-Content (Join-Path (Join-Path $PSScriptRoot '..') 'minotaur.ps1') -Raw
    }

    It 'contains every code the tool raises' {
        $raised = @([regex]::Matches($catacombSource, "Add-MinotaurFinding\s+-Code\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
        $raised.Count | Should -BeGreaterThan 7
        foreach ($code in $raised) { $MinotaurFindings.ContainsKey($code) | Should -BeTrue -Because "minotaur.ps1 raises '$code'" }
    }
    It 'ranks broad write as an Error and broad read as a Warning' {
        $MinotaurFindings['BroadWriteAccess'].Severity | Should -Be 'Error'
        $MinotaurFindings['BroadReadAccess'].Severity  | Should -Be 'Warning'
    }
    It 'collapses the standard FileSystemRights masks' {
        Get-NtfsRightsLevel -Value 2032127 | Should -Be 'Full'
        Get-NtfsRightsLevel -Value 197055  | Should -Be 'Modify'
        Get-NtfsRightsLevel -Value 278     | Should -Be 'Write'
        Get-NtfsRightsLevel -Value 131241  | Should -Be 'Read'
        Get-NtfsRightsLevel -Value 65536   | Should -Be 'Special'
    }
    It 'maps the generic rights inheritable entries carry' {
        Get-NtfsRightsLevel -Value 268435456   | Should -Be 'Full'
        Get-NtfsRightsLevel -Value 1073741824  | Should -Be 'Write'
        Get-NtfsRightsLevel -Value -2147483648 | Should -Be 'Read'
    }
    It 'treats Full, Modify, Write and the share right Change as write' {
        foreach ($l in 'Full', 'Modify', 'Write', 'Change') { Test-WriteLevel $l | Should -BeTrue }
        foreach ($l in 'Read', 'Special') { Test-WriteLevel $l | Should -BeFalse }
    }
    It 'recognises everyone-type principals, including Domain Users by RID' {
        foreach ($s in 'S-1-1-0', 'S-1-5-11', 'S-1-5-32-545', 'S-1-5-21-1-2-3-513') { Get-SidCategory $s | Should -Be 'Broad' -Because "$s means everyone" }
        Get-SidCategory 'S-1-5-21-1-2-3-512'  | Should -Be 'Account'
        Get-SidCategory 'S-1-5-21-1-2-3-1105' | Should -Be 'Account'
        Get-SidCategory 'S-1-5-18'            | Should -Be 'WellKnown'
        Get-SidCategory 'S-1-5-32-544'        | Should -Be 'WellKnown'
    }
    It 'maps the worst severity to Exposed / Review / Tidy' {
        (Get-MinotaurVerdict -FindingList @([PSCustomObject]@{ Severity = 'Error' })).Verdict   | Should -Be 'Exposed'
        (Get-MinotaurVerdict -FindingList @([PSCustomObject]@{ Severity = 'Warning' })).Verdict | Should -Be 'Review'
        (Get-MinotaurVerdict -FindingList @([PSCustomObject]@{ Severity = 'Info' })).Verdict    | Should -Be 'Tidy'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ARGUS LDAP / DN helpers — ARGUS interpolates distinguished names straight
# into LDAP filter strings, so the escaping helper is the boundary between a
# correct query and one whose meaning a stray parenthesis has changed. The
# helpers are pure, so they are extracted by AST and exercised directly rather
# than dot-sourcing the tool (which launches its main flow on import).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'ARGUS LDAP and DN helpers' {
    BeforeAll {
        # Pester 5 restricts Should to It bodies, so the "did the helper load?"
        # check is recorded here and asserted in its own It below rather than
        # being asserted inline.
        $heraldPath = Join-Path $PSScriptRoot '..\argus.ps1'
        $errs = $null
        $ast  = [System.Management.Automation.Language.Parser]::ParseFile($heraldPath, [ref]$null, [ref]$errs)
        $script:ArgusHelpersLoaded = @()
        foreach ($name in 'ConvertTo-LdapFilterValue', 'Get-DnLeaf', 'Get-DnParent') {
            $fn = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
            }, $true) | Select-Object -First 1
            if ($fn) {
                . ([scriptblock]::Create($fn.Extent.Text))
                $script:ArgusHelpersLoaded += $name
            }
        }
    }

    It 'argus.ps1 defines the LDAP and DN helpers' {
        foreach ($name in 'ConvertTo-LdapFilterValue', 'Get-DnLeaf', 'Get-DnParent') {
            $script:ArgusHelpersLoaded | Should -Contain $name -Because "argus.ps1 must define $name"
        }
    }

    Context 'ConvertTo-LdapFilterValue' {
        It 'escapes the RFC 4515 reserved characters' {
            ConvertTo-LdapFilterValue 'a(b)c' | Should -Be 'a\28b\29c'
            ConvertTo-LdapFilterValue 'a*b'   | Should -Be 'a\2ab'
            ConvertTo-LdapFilterValue 'a\b'   | Should -Be 'a\5cb'
        }
        It 'escapes a distinguished name containing parentheses' {
            ConvertTo-LdapFilterValue 'CN=Admins (Legacy),DC=contoso,DC=com' |
                Should -Be 'CN=Admins \28Legacy\29,DC=contoso,DC=com'
        }
        It 'passes an ordinary distinguished name through unchanged' {
            $dn = 'CN=Domain Admins,CN=Users,DC=contoso,DC=com'
            ConvertTo-LdapFilterValue $dn | Should -Be $dn
        }
        It 'returns an empty string for null or empty input' {
            ConvertTo-LdapFilterValue $null | Should -Be ''
            ConvertTo-LdapFilterValue ''    | Should -Be ''
        }
    }

    Context 'Get-DnLeaf / Get-DnParent' {
        It 'reads the leaf value of a distinguished name' {
            Get-DnLeaf 'CN=Jane Doe,OU=Staff,DC=contoso,DC=com' | Should -Be 'Jane Doe'
        }
        It 'reads the container a distinguished name sits in' {
            Get-DnParent 'CN=Jane Doe,OU=Staff,DC=contoso,DC=com' | Should -Be 'OU=Staff,DC=contoso,DC=com'
        }
        It 'does not split on an escaped comma inside a CN' {
            # "Doe, Jane" is stored as CN=Doe\, Jane — splitting naively on every
            # comma would report the manager as "Doe" and the OU as " Jane,OU=...".
            Get-DnLeaf   'CN=Doe\, Jane,OU=Staff,DC=contoso,DC=com' | Should -Be 'Doe, Jane'
            Get-DnParent 'CN=Doe\, Jane,OU=Staff,DC=contoso,DC=com' | Should -Be 'OU=Staff,DC=contoso,DC=com'
        }
        It 'returns an empty string for null, empty, or single-component input' {
            Get-DnLeaf   $null              | Should -Be ''
            Get-DnParent ''                 | Should -Be ''
            Get-DnParent 'DC=com'           | Should -Be ''
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ARGUS report-section guard — ARGUS renders a whole domain's roster, so a
# single unrenderable row must not cost the technician the entire document. The
# guard replaces a failed section with a visible placeholder and reports where
# the fault came from, rather than losing the report to one exception.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'ARGUS report-section guard' {
    BeforeAll {
        $heraldPath = Join-Path $PSScriptRoot '..\argus.ps1'
        $errs = $null
        $ast  = [System.Management.Automation.Language.Parser]::ParseFile($heraldPath, [ref]$null, [ref]$errs)
        $script:ArgusGuardLoaded = @()
        foreach ($name in 'Invoke-ReportSection', 'Get-FaultLocation') {
            $fn = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
            }, $true) | Select-Object -First 1
            if ($fn) {
                . ([scriptblock]::Create($fn.Extent.Text))
                $script:ArgusGuardLoaded += $name
            }
        }
    }

    It 'argus.ps1 defines the section guard and the fault locator' {
        $script:ArgusGuardLoaded | Should -Contain 'Invoke-ReportSection'
        $script:ArgusGuardLoaded | Should -Contain 'Get-FaultLocation'
    }

    It 'passes a section that renders cleanly straight through' {
        Invoke-ReportSection -Name 'clean' -ColSpan 3 -Build { '<tr><td>ok</td></tr>' } |
            Should -Be '<tr><td>ok</td></tr>'
    }

    It 'swallows a failing section into a placeholder row rather than throwing' {
        $row = Invoke-ReportSection -Name 'boom' -ColSpan 7 -Build { throw 'kaboom' }
        $row | Should -Match 'could not be rendered'
    }

    It 'spans the placeholder across the real column count of the failed table' {
        # A placeholder that does not span the table renders as a broken row.
        Invoke-ReportSection -Name 'boom' -ColSpan 9 -Build { throw 'kaboom' } |
            Should -Match 'colspan="9"'
    }

    It 'lets the sections either side of a failure still render' {
        $before = Invoke-ReportSection -Name 'before' -ColSpan 2 -Build { '<tr><td>a</td></tr>' }
        $failed = Invoke-ReportSection -Name 'failed' -ColSpan 2 -Build { throw 'nope' }
        $after  = Invoke-ReportSection -Name 'after'  -ColSpan 2 -Build { '<tr><td>b</td></tr>' }
        ($before + $failed + $after) | Should -Match '<tr><td>a</td></tr>'
        ($before + $failed + $after) | Should -Match '<tr><td>b</td></tr>'
    }

    Context 'Get-FaultLocation' {
        It 'reports the originating file and line for a script-borne error' {
            try { throw 'boom' } catch { $rec = $_ }
            Get-FaultLocation $rec | Should -Match 'line \d+'
        }
        It 'degrades to a bare line number when the error carries no script origin' {
            $rec = [PSCustomObject]@{ InvocationInfo = [PSCustomObject]@{ ScriptLineNumber = 42; ScriptName = '' } }
            Get-FaultLocation $rec | Should -Be 'line 42'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ARGUS password-policy verdicts — the report answers an access questionnaire,
# so each setting is scored rather than merely printed. The zero cases carry the
# most meaning and the least intuition: zero lockout threshold disables lockout
# entirely, zero max age means passwords never expire, and zero lockout duration
# means locked until an administrator intervenes — the strictest option, not the
# weakest.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'ARGUS password-policy verdicts' {
    BeforeAll {
        $heraldPath = Join-Path $PSScriptRoot '..\argus.ps1'
        $errs = $null
        $ast  = [System.Management.Automation.Language.Parser]::ParseFile($heraldPath, [ref]$null, [ref]$errs)

        foreach ($name in 'Get-PolicyVerdict', 'Format-PolicyValue') {
            $fn = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
            }, $true) | Select-Object -First 1
            if ($fn) { . ([scriptblock]::Create($fn.Extent.Text)) }
        }

        # Mirror the plain-hashtable map herald builds at load, which the two
        # helpers read from.
        $assign = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $n.Left.VariablePath.UserPath -eq 'PasswordPolicyBaseline'
        }, $true) | Select-Object -First 1
        $ordered = & ([scriptblock]::Create($assign.Right.Extent.Text))
        $script:PolicyBaseline = @{}
        foreach ($k in @($ordered.Keys)) { $script:PolicyBaseline[[string]$k] = $ordered[[string]$k] }
        $PolicyBaseline = $script:PolicyBaseline
    }

    It 'loaded the baseline the helpers depend on' {
        # Without this the helpers silently score everything Strong off a null map.
        $script:PolicyBaseline | Should -Not -BeNullOrEmpty
        $script:PolicyBaseline.Count | Should -BeGreaterThan 5
    }

    Context 'password strength' {
        It 'scores minimum length against the baseline' {
            Get-PolicyVerdict -Key 'MinPasswordLength' -Value 14 | Should -Be 'Strong'
            Get-PolicyVerdict -Key 'MinPasswordLength' -Value 8  | Should -Be 'Acceptable'
            Get-PolicyVerdict -Key 'MinPasswordLength' -Value 6  | Should -Be 'Weak'
        }
        It 'requires complexity to be enabled' {
            Get-PolicyVerdict -Key 'ComplexityEnabled' -Value $true  | Should -Be 'Strong'
            Get-PolicyVerdict -Key 'ComplexityEnabled' -Value $false | Should -Be 'Weak'
        }
        It 'scores password history' {
            Get-PolicyVerdict -Key 'PasswordHistoryCount' -Value 24 | Should -Be 'Strong'
            Get-PolicyVerdict -Key 'PasswordHistoryCount' -Value 0  | Should -Be 'Weak'
        }
        It 'treats reversible encryption as a finding when enabled' {
            Get-PolicyVerdict -Key 'ReversibleEncryptionEnabled' -Value $true  | Should -Be 'Weak'
            Get-PolicyVerdict -Key 'ReversibleEncryptionEnabled' -Value $false | Should -Be 'Strong'
        }
    }

    Context 'the zero cases' {
        It 'calls a zero lockout threshold Weak, because lockout is then off' {
            Get-PolicyVerdict -Key 'LockoutThreshold' -Value 0  | Should -Be 'Weak'
            Get-PolicyVerdict -Key 'LockoutThreshold' -Value 5  | Should -Be 'Strong'
            Get-PolicyVerdict -Key 'LockoutThreshold' -Value 50 | Should -Be 'Weak'
        }
        It 'calls a zero lockout duration Strong, because it holds until an admin unlocks' {
            Get-PolicyVerdict -Key 'LockoutDurationMinutes' -Value 0 | Should -Be 'Strong'
            Get-PolicyVerdict -Key 'LockoutDurationMinutes' -Value 5 | Should -Be 'Weak'
        }
        It 'does not fail a no-expiry policy outright, but does not call it Strong either' {
            # NIST SP 800-63B advises against routine expiry, so this is reported
            # for the questionnaire rather than scored as a defect.
            Get-PolicyVerdict -Key 'MaxPasswordAgeDays' -Value 0   | Should -Be 'Acceptable'
            Get-PolicyVerdict -Key 'MaxPasswordAgeDays' -Value 90  | Should -Be 'Strong'
            Get-PolicyVerdict -Key 'MaxPasswordAgeDays' -Value 999 | Should -Be 'Weak'
        }
        It 'returns Unknown rather than guessing when the directory gave no value' {
            Get-PolicyVerdict -Key 'MinPasswordLength' -Value $null | Should -Be 'Unknown'
            Get-PolicyVerdict -Key 'NoSuchSetting' -Value 1         | Should -Be 'Unknown'
        }
    }

    Context 'Format-PolicyValue' {
        It 'spells out what each meaningful zero means' {
            Format-PolicyValue -Key 'MaxPasswordAgeDays'     -Value 0 | Should -Be 'Never expires'
            Format-PolicyValue -Key 'LockoutThreshold'       -Value 0 | Should -Be 'Never locks out'
            Format-PolicyValue -Key 'LockoutDurationMinutes' -Value 0 | Should -Be 'Until an administrator unlocks'
        }
        It 'renders booleans as Enabled / Disabled' {
            Format-PolicyValue -Key 'ComplexityEnabled' -Value $true  | Should -Be 'Enabled'
            Format-PolicyValue -Key 'ComplexityEnabled' -Value $false | Should -Be 'Disabled'
        }
        It 'appends the unit to a plain number' {
            Format-PolicyValue -Key 'MinPasswordLength' -Value 14 | Should -Be '14 characters'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ARGUS report parameters must stay untyped — a live 91-account domain on
# Windows PowerShell 5.1 threw System.ArgumentException "Argument types do not
# match" binding arguments into Build-ArgusReport, at the call statement and
# therefore before any section guard could catch it, losing the whole report. A
# type constraint is the only thing that can fail at a call site, so the
# parameters are deliberately unconstrained and the collections are normalised
# inside the function instead.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'ARGUS report parameters' {
    BeforeAll {
        $heraldPath = Join-Path $PSScriptRoot '..\argus.ps1'
        $errs = $null
        $ast  = [System.Management.Automation.Language.Parser]::ParseFile($heraldPath, [ref]$null, [ref]$errs)
        $script:BuildReportFn = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Build-ArgusReport'
        }, $true) | Select-Object -First 1
    }

    It 'defines Build-ArgusReport' {
        $script:BuildReportFn | Should -Not -BeNullOrEmpty
    }

    It 'declares every parameter without a type constraint' {
        $params = $script:BuildReportFn.Body.ParamBlock.Parameters
        $params.Count | Should -BeGreaterThan 0
        foreach ($p in $params) {
            $name = $p.Name.VariablePath.UserPath
            # Attributes on a ParameterAst include its type constraint, if any.
            $constraint = @($p.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            })
            $constraint.Count | Should -Be 0 -Because "binding a constraint on -$name is what lost the report on PowerShell 5.1"
        }
    }

    It 'normalises the two collections inside the body instead' {
        # This is what the removed [array] constraints were buying; without it
        # .Count on a single object or a null would misbehave. It deliberately
        # does not use @(), which is the construct suspected of raising
        # "Argument types do not match" over a List[object].
        $text = $script:BuildReportFn.Extent.Text
        $text | Should -Match '\$Roster\s*=\s*ConvertTo-ArgusArray\s+\$Roster'
        $text | Should -Match '\$GroupSummary\s*=\s*ConvertTo-ArgusArray\s+\$GroupSummary'
        $text | Should -Not -Match '@\(\$Roster\)'
        $text | Should -Not -Match '@\(\$GroupSummary\)'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ConvertTo-ArgusArray — the report was lost repeatedly to an ArgumentException
# ("Argument types do not match") raised while preparing its arguments. The one
# collection built as a System.Collections.Generic.List[object] is the group
# summary, and @() over such a list is the construct under suspicion. Rather
# than depend on @() behaving, the helper enumerates explicitly; these cases pin
# that it handles every shape the report is given.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'ConvertTo-ArgusArray' {
    BeforeAll {
        $heraldPath = Join-Path $PSScriptRoot '..\argus.ps1'
        $errs = $null
        $ast  = [System.Management.Automation.Language.Parser]::ParseFile($heraldPath, [ref]$null, [ref]$errs)
        $fn = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-ArgusArray'
        }, $true) | Select-Object -First 1
        $script:ConvertLoaded = $null -ne $fn
        if ($fn) { . ([scriptblock]::Create($fn.Extent.Text)) }
    }

    It 'is defined in argus.ps1' {
        $script:ConvertLoaded | Should -BeTrue
    }

    It 'converts a List[object] built with New-Object' {
        # The construction the group summary used before 5.1. ARGUS now builds
        # its lists with ::new(), but the helper must still accept this form.
        $list = New-Object System.Collections.Generic.List[object]
        $list.Add([PSCustomObject]@{ Name = 'G1' })
        $list.Add([PSCustomObject]@{ Name = 'G2' })
        $result = ConvertTo-ArgusArray $list
        $result -is [array] | Should -BeTrue
        $result.Count | Should -Be 2
    }

    It 'converts an empty List[object] to an empty array, not null' {
        $list = New-Object System.Collections.Generic.List[object]
        $result = ConvertTo-ArgusArray $list
        $result -is [array] | Should -BeTrue
        $result.Count | Should -Be 0
    }

    It 'passes an array through with its contents intact' {
        $result = ConvertTo-ArgusArray @(1, 2, 3)
        $result.Count | Should -Be 3
    }

    It 'returns an empty array for null rather than null' {
        $result = ConvertTo-ArgusArray $null
        $result -is [array] | Should -BeTrue
        $result.Count | Should -Be 0
    }

    It 'wraps a scalar as a single element' {
        (ConvertTo-ArgusArray ([PSCustomObject]@{ A = 1 })).Count | Should -Be 1
    }

    It 'does not split a string into characters' {
        $result = ConvertTo-ArgusArray 'hello'
        $result.Count | Should -Be 1
        $result[0] | Should -Be 'hello'
    }

    It 'keeps a hashtable as one element rather than enumerating its entries' {
        (ConvertTo-ArgusArray @{ a = 1; b = 2 }).Count | Should -Be 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# License header compliance — the toolkit is GPL-3.0-or-later and its whole
# distribution model is "copy one .ps1 onto the machine and run it". A lone
# script that travels without its notice cannot tell the next technician what
# it is or what rights they have, so every source file carries the notice in
# its own header. This test stops a new tool from being added without one.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'License header compliance — all source files' {
    $licenseCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Include '*.ps1', '*.psm1' -File -Recurse |
        Where-Object { $_.FullName -notmatch $NonSourceDir } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> carries the SPDX license identifier' -ForEach $licenseCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Match 'SPDX-License-Identifier:\s*GPL-3\.0-or-later' -Because "$Name must declare its license in its own header"
    }

    It '<Name> carries the GPL notice and a copyright line' -ForEach $licenseCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Match 'GNU General Public License'
        $content | Should -Match '(?m)^\s*#\s*Copyright \(C\) \d{4}'
    }

    It '<Name> keeps the notice at the top, above the comment-based help' -ForEach $licenseCases {
        # Comment-based help is only picked up when preceded solely by comments
        # and blank lines, so the notice must sit above it, not inside it.
        $content = Get-Content $FullName -Raw
        $spdxAt  = $content.IndexOf('SPDX-License-Identifier')
        $helpAt  = $content.IndexOf('<#')
        $spdxAt | Should -BeGreaterThan -1
        if ($helpAt -ge 0) {
            $spdxAt | Should -BeLessThan $helpAt -Because "$Name must declare its license before its help block"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# License header compliance — the desktop application sources.
#
# The application is a distributable unit of this GPL work just as the scripts
# are, so its sources carry the notice too. Two established shapes, and the gate
# follows the code rather than imposing a third: .cs files carry the full notice
# like the scripts do, while .xaml files carry a short comment with the
# copyright line and the SPDX tag — a XAML file is markup, and five of the seven
# already did it this way before this gate existed.
#
# bin/ and obj/ are build output, not source; generated assembly-info files
# there would otherwise be held to a header nobody wrote.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'License header compliance — desktop app sources' {
    $appRoot = Join-Path $PSScriptRoot '..\app'

    $appCases = if (Test-Path $appRoot) {
        Get-ChildItem -Path $appRoot -Include '*.cs', '*.xaml' -File -Recurse |
            Where-Object { $_.FullName -notmatch ('{0}(bin|obj){0}' -f [regex]::Escape([string][IO.Path]::DirectorySeparatorChar)) } |
            ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName; Ext = $_.Extension.ToLowerInvariant() } }
    } else { @() }

    # Without this an empty or moved app/ would make every case below vacuous
    # and the gate would pass while checking nothing. Carried as test data for
    # the same discovery-versus-run scoping reason as the version gate above.
    It 'the app directory yielded sources to check' -ForEach @{ CaseCount = $appCases.Count } {
        $CaseCount | Should -BeGreaterThan 20
    }

    It '<Name> carries the SPDX license identifier' -ForEach $appCases {
        (Get-Content $FullName -Raw) | Should -Match 'SPDX-License-Identifier:\s*GPL-3\.0-or-later' -Because "$Name must declare its license in its own header"
    }

    It '<Name> carries a copyright line' -ForEach $appCases {
        (Get-Content $FullName -Raw) | Should -Match 'Copyright \(C\) \d{4}' -Because "$Name must name the copyright holder and year"
    }

    It '<Name> carries the full GPL notice (.cs only)' -ForEach ($appCases | Where-Object { $_.Ext -eq '.cs' }) {
        (Get-Content $FullName -Raw) | Should -Match 'GNU General Public License'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# LICENSE file integrity — the GPL text is a legal document that must be
# distributed verbatim; a truncated or edited copy undermines the grant.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'LICENSE file' {
    BeforeAll {
        $script:LicensePath = Join-Path $PSScriptRoot '..\LICENSE'
    }

    It 'exists' {
        $script:LicensePath | Should -Exist
    }

    It 'is the GNU GPL version 3' {
        $text = Get-Content $script:LicensePath -Raw
        $text | Should -Match 'GNU GENERAL PUBLIC LICENSE'
        $text | Should -Match 'Version 3, 29 June 2007'
    }

    It 'is complete — carries the final "How to Apply" section' {
        $text = Get-Content $script:LicensePath -Raw
        $text | Should -Match 'How to Apply These Terms to Your New Programs'
        # The closing LGPL paragraph is the last thing in the document. It is
        # matched across a line break because the canonical FSF text wraps it.
        $text | Should -Match 'GNU Lesser General\s+Public License'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Color-map safety — guards two classes of "Cannot convert null to ConsoleColor"
# crashes that have bitten the toolkit:
#
#   1. CLEANSE: a tool referenced `$ColorSchema.Menu` six times but never
#      defined a `Menu` key, so `-ForegroundColor` received $null.
#   2. ANVIL / PORTAL / ORRERY: a `foreach ($c in ...)` loop variable collided
#      (PowerShell variables are case-insensitive) with the `$C` color map,
#      clobbering it so later `$C.Success` reads returned $null.
#
# A "color map" is detected structurally — any variable assigned a hashtable
# literal whose values are all valid ConsoleColor names — so the test stays
# name-agnostic (`$C`, `$ColorSchema`, anything).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Color-map safety — all scripts' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    BeforeAll {
        # Returns a hashtable: color-map variable name (lower-case) -> [string[]]
        # of defined keys (lower-case), for every hashtable literal in the AST
        # whose values are all ConsoleColor names.
        function Get-ColorMapKeys {
            param($Ast)
            $consoleColors = [enum]::GetNames([System.ConsoleColor])
            $maps = @{}
            $assignments = $Ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left -is [System.Management.Automation.Language.VariableExpressionAst]
            }, $true)
            foreach ($a in $assignments) {
                # The RHS must *be* a hashtable literal (not merely contain one).
                $ht = $a.Right.Find({
                    param($n) $n -is [System.Management.Automation.Language.HashtableAst]
                }, $false)
                if (-not $ht) { continue }
                if ($ht.Extent.Text.Trim() -ne $a.Right.Extent.Text.Trim()) { continue }
                if ($ht.KeyValuePairs.Count -eq 0) { continue }

                $keys      = @()
                $allColors = $true
                foreach ($pair in $ht.KeyValuePairs) {
                    $keys += $pair.Item1.Extent.Text.Trim().Trim("'`"")
                    $valText = $pair.Item2.Extent.Text.Trim().Trim("'`"")
                    if ($valText -notin $consoleColors) { $allColors = $false }
                }
                if ($allColors) {
                    $name = $a.Left.VariablePath.UserPath.ToLower()
                    if (-not $maps.ContainsKey($name)) { $maps[$name] = @() }
                    $maps[$name] += ($keys | ForEach-Object { $_.ToLower() })
                }
            }
            return $maps
        }
    }

    It '<Name>: every -Foreground/-BackgroundColor $map.Key uses a defined key' -ForEach $scriptCases {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($FullName, [ref]$null, [ref]$null)
        $maps = Get-ColorMapKeys -Ast $ast

        $bad = @()
        $commands = $ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.CommandAst]
        }, $true)
        foreach ($cmd in $commands) {
            $els = $cmd.CommandElements
            for ($i = 0; $i -lt $els.Count; $i++) {
                $el = $els[$i]
                if ($el -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
                if ($el.ParameterName -notin @('ForegroundColor', 'BackgroundColor')) { continue }
                $arg = if ($el.Argument) { $el.Argument } elseif ($i + 1 -lt $els.Count) { $els[$i + 1] } else { $null }
                if ($arg -is [System.Management.Automation.Language.MemberExpressionAst] -and
                    $arg.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and
                    $arg.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $mapName = $arg.Expression.VariablePath.UserPath.ToLower()
                    $member  = $arg.Member.Value.ToLower()
                    if ($maps.ContainsKey($mapName) -and $member -notin $maps[$mapName]) {
                        $bad += $arg.Extent.Text
                    }
                }
            }
        }
        $bad -join ', ' | Should -BeNullOrEmpty
    }

    It '<Name>: no foreach loop variable shadows a color map' -ForEach $scriptCases {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($FullName, [ref]$null, [ref]$null)
        $maps = Get-ColorMapKeys -Ast $ast

        $bad = @()
        $loops = $ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.ForEachStatementAst]
        }, $true)
        foreach ($loop in $loops) {
            $loopVar = $loop.Variable.VariablePath.UserPath.ToLower()
            if ($maps.ContainsKey($loopVar)) {
                $bad += "`$$($loop.Variable.VariablePath.UserPath)"
            }
        }
        $bad -join ', ' | Should -BeNullOrEmpty
    }
}
