# cleanse.ps1 - C.L.E.A.N.S.E. — Cleans Leftover, Ephemeral And Neglected System Entries
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
    C.L.E.A.N.S.E. — Cleans Leftover, Ephemeral And Neglected System Entries
    Disk Cleanup Tool for PowerShell 5.1+

.DESCRIPTION
    Cleans common junk accumulation points on Windows: user and system temp
    folders, Windows Update download cache, the Recycle Bin, and browser
    caches (Chrome, Edge, Firefox). Shows estimated space before each
    category and reports total freed space after cleanup.

.USAGE
    PS C:\> .\cleanse.ps1                    # Must be run as Administrator
    PS C:\> .\cleanse.ps1 -Unattended        # Silent mode — cleans all categories, no prompts
    PS C:\> .\cleanse.ps1 -WhatIf            # Preview what would be cleaned, without deleting anything

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# INITIALIZATION
# ─────────────────────────────────────────────────────────────────────────────

# ===========================
# SHARED HELPERS
# ===========================
#region TK-SHARED
# Generated from shared/TechnicianToolkit.Shared.ps1 by build/Sync-TKShared.ps1. Do not edit
# this block: edit the source and re-run the script. The Pester suite fails on a stale block.

# Forcing UTF-8 keeps the box-drawing banners intact when a console is attached.
# When one is not — a GUI host, a WinRM runspace, a scheduled task — the setter
# throws "The handle is invalid", so this must never be fatal. Every tool runs
# this as the first line of its generated block.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

function Format-Bytes {
    <#
    .SYNOPSIS
        Formats a byte count as a human-readable string (B/KB/MB/GB/TB).
    .EXAMPLE
        Format-Bytes 1572864   # -> '1.50 MB'
    #>
    param([Parameter(Mandatory)][long]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes B" }
    if ($Bytes -lt 1MB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -lt 1TB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    return "{0:N2} TB" -f ($Bytes / 1TB)
}

function Get-TKConfig {
    <#
    .SYNOPSIS
        Returns the toolkit configuration as a PSCustomObject.
        Reads config.json from the tool's own directory; missing keys are filled
        with empty-string defaults so callers never receive null.
    #>
    $configPath = Join-Path $PSScriptRoot 'config.json'

    $defaults = [PSCustomObject]@{
        OrgName       = ''
        LogDirectory  = ''
        TeamsWebhook  = ''
        Archive       = [PSCustomObject]@{ DefaultDestination   = '' }
        Revenant      = [PSCustomObject]@{ DefaultDestination   = '' }
        Covenant      = [PSCustomObject]@{ DefaultTimezone = ''; DefaultLocalAdminUser = '' }
        # DirectDownloads is an array of { Name; Url; Args; Sha256 } — installers
        # CONJURE fetches by URL because they have no winget/Chocolatey package.
        Conjure       = [PSCustomObject]@{ DirectDownloads = @() }
    }

    if (-not (Test-Path $configPath)) { return $defaults }

    try {
        $raw = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json

        # Legacy migration: v2.x used a 'Phantom' section; v3.0 renamed it to 'Revenant'.
        # Copy any populated 'Phantom' value forward so upgrading users don't lose their default destination.
        if ($raw.PSObject.Properties['Phantom'] -and
            $raw.Phantom -is [PSCustomObject] -and
            -not [string]::IsNullOrWhiteSpace($raw.Phantom.DefaultDestination)) {
            if ($raw.PSObject.Properties['Revenant'] -isnot [PSCustomObject] -or
                [string]::IsNullOrWhiteSpace($raw.Revenant.DefaultDestination)) {
                if (-not $raw.PSObject.Properties['Revenant']) {
                    $raw | Add-Member -NotePropertyName 'Revenant' -NotePropertyValue ([PSCustomObject]@{ DefaultDestination = '' }) -Force
                }
                $raw.Revenant.DefaultDestination = $raw.Phantom.DefaultDestination
            }
        }

        # Ensure top-level keys exist
        foreach ($key in ($defaults | Get-Member -MemberType NoteProperty).Name) {
            if ($null -eq $raw.$key) {
                $raw | Add-Member -NotePropertyName $key -NotePropertyValue $defaults.$key -Force
            }
        }
        # Ensure nested keys exist
        foreach ($section in @('Archive','Revenant','Covenant','Conjure')) {
            if ($raw.$section -isnot [PSCustomObject]) {
                $raw | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            } else {
                foreach ($key in ($defaults.$section | Get-Member -MemberType NoteProperty).Name) {
                    if ($null -eq $raw.$section.$key) {
                        # Carry the default's own type: filling an absent array key
                        # with '' would hand the caller a string to iterate.
                        $fallback = if ($defaults.$section.$key -is [Array]) { @() } else { '' }
                        $raw.$section | Add-Member -NotePropertyName $key -NotePropertyValue $fallback -Force
                    }
                }
            }
        }
        return $raw
    }
    catch { return $defaults }
}

function Start-TKTranscript {
    <#
    .SYNOPSIS
        Starts a timestamped PowerShell transcript in the configured log directory.
        Call once near the top of a script when -Transcript is active.
    #>
    param([Parameter(Mandatory)][string]$LogRoot)
    $path = Join-Path $LogRoot "TK_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    try {
        Start-Transcript -Path $path -Append -ErrorAction Stop | Out-Null
        Write-Host "  [*] Transcript: $path" -ForegroundColor Gray
    }
    catch {
        Write-Warn "Could not start transcript: $_"
    }
}

function Stop-TKTranscript {
    <#
    .SYNOPSIS
        Stops an active PowerShell transcript started by Start-TKTranscript.
    #>
    try { Stop-Transcript -ErrorAction Stop | Out-Null } catch {}
}

function Resolve-LogDirectory {
    <#
    .SYNOPSIS
        Returns the configured LogDirectory, or the supplied fallback path if
        LogDirectory is not set. Creates the directory if it does not exist.
    #>
    param([Parameter(Mandatory)][string]$FallbackPath)

    $cfg = Get-TKConfig
    if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory)) {
        if (-not (Test-Path $cfg.LogDirectory)) {
            $null = New-Item -ItemType Directory -Path $cfg.LogDirectory -Force -ErrorAction SilentlyContinue
        }
        return $cfg.LogDirectory
    }
    return $FallbackPath
}

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-AdminElevation {
    param([Parameter(Mandatory)][string]$ScriptFile)
    if (-not (Test-IsAdmin)) {
        Write-Host "  INFO: Restarting with administrator privileges..." -ForegroundColor Yellow
        $PSExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        Start-Process -FilePath $PSExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptFile`"" -Verb RunAs
        exit
    }
}
#endregion TK-SHARED
Invoke-AdminElevation -ScriptFile $PSCommandPath

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA
# ─────────────────────────────────────────────────────────────────────────────

$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
    Menu     = 'White'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-CleanseBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗██╗     ███████╗ █████╗ ███╗   ██╗███████╗███████╗
  ██╔════╝██║     ██╔════╝██╔══██╗████╗  ██║██╔════╝██╔════╝
  ██║     ██║     █████╗  ███████║██╔██╗ ██║███████╗█████╗
  ██║     ██║     ██╔══╝  ██╔══██║██║╚██╗██║╚════██║██╔══╝
  ╚██████╗███████╗███████╗██║  ██║██║ ╚████║███████║███████╗
   ╚═════╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝╚══════╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    C.L.E.A.N.S.E. — Cleans Leftover, Ephemeral And Neglected System Entries" -ForegroundColor Cyan
    Write-Host "    Disk Cleanup Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Get-FolderSize {
    param([string[]]$Paths)
    $total = 0
    foreach ($p in $Paths) {
        if (Test-Path $p) {
            try {
                $total += (Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue |
                    Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            } catch {}
        }
    }
    return $total
}

function Remove-FolderContents {
    param([string[]]$Paths, [string]$Label)
    $freed = 0
    foreach ($p in $Paths) {
        if (-not (Test-Path $p)) { continue }
        $items = Get-ChildItem -Path $p -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            try {
                $size = if ($item.PSIsContainer) {
                    (Get-ChildItem -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                } else { $item.Length }
                Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                $freed += $size
            } catch {}
        }
    }
    return $freed
}

# ─────────────────────────────────────────────────────────────────────────────
# CLEANUP CATEGORIES
# ─────────────────────────────────────────────────────────────────────────────

# Resolve all user profile temp paths
function Get-UserTempPaths {
    $paths = @($env:TEMP)
    $localAppData = $env:LOCALAPPDATA
    if ($localAppData) { $paths += Join-Path $localAppData 'Temp' }
    return $paths | Select-Object -Unique | Where-Object { $_ -and (Test-Path $_) }
}

function Get-BrowserCachePaths {
    $paths = [System.Collections.Generic.List[object]]::new()
    $profiles = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue

    foreach ($profile in $profiles) {
        $base = $profile.FullName
        # Chrome
        $chromeCacheBase = Join-Path $base 'AppData\Local\Google\Chrome\User Data'
        if (Test-Path $chromeCacheBase) {
            Get-ChildItem $chromeCacheBase -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^(Default|Profile)' } |
                ForEach-Object { $paths.Add((Join-Path $_.FullName 'Cache')); $paths.Add((Join-Path $_.FullName 'Code Cache')) }
        }
        # Edge
        $edgeCacheBase = Join-Path $base 'AppData\Local\Microsoft\Edge\User Data'
        if (Test-Path $edgeCacheBase) {
            Get-ChildItem $edgeCacheBase -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^(Default|Profile)' } |
                ForEach-Object { $paths.Add((Join-Path $_.FullName 'Cache')); $paths.Add((Join-Path $_.FullName 'Code Cache')) }
        }
        # Firefox
        $ffProfiles = Join-Path $base 'AppData\Local\Mozilla\Firefox\Profiles'
        if (Test-Path $ffProfiles) {
            Get-ChildItem $ffProfiles -Directory -ErrorAction SilentlyContinue |
                ForEach-Object { $paths.Add((Join-Path $_.FullName 'cache2')) }
        }
    }
    return $paths | Where-Object { Test-Path $_ }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-CleanseBanner }

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SCANNING CLEANUP TARGETS" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

Write-Host "  [*] Calculating sizes..." -ForegroundColor $ColorSchema.Progress

$userTempPaths    = Get-UserTempPaths
$sysTempPaths     = @('C:\Windows\Temp')
$wuCachePaths     = @('C:\Windows\SoftwareDistribution\Download')
$browserCachePaths = Get-BrowserCachePaths

$userTempSize    = Get-FolderSize -Paths $userTempPaths
$sysTempSize     = Get-FolderSize -Paths $sysTempPaths
$wuCacheSize     = Get-FolderSize -Paths $wuCachePaths
$browserCacheSize = Get-FolderSize -Paths $browserCachePaths

$recycleBinSize = 0
try {
    $shell = New-Object -ComObject Shell.Application
    $bin   = $shell.Namespace(0xA)
    if ($bin) {
        $recycleBinSize = ($bin.Items() | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum
    }
} catch {}

Write-Host ""
Write-Host "  Cleanup categories:" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host ("  [1]  User Temp Folders          {0,10}" -f (Format-Bytes $userTempSize)) -ForegroundColor $ColorSchema.Menu
Write-Host ("  [2]  System Temp (Windows\Temp) {0,10}" -f (Format-Bytes $sysTempSize)) -ForegroundColor $ColorSchema.Menu
Write-Host ("  [3]  Windows Update Cache       {0,10}" -f (Format-Bytes $wuCacheSize)) -ForegroundColor $ColorSchema.Menu
Write-Host ("  [4]  Recycle Bin               {0,10}" -f (Format-Bytes $recycleBinSize)) -ForegroundColor $ColorSchema.Menu
Write-Host ("  [5]  Browser Caches            {0,10}" -f (Format-Bytes $browserCacheSize)) -ForegroundColor $ColorSchema.Menu
$totalSize = $userTempSize + $sysTempSize + $wuCacheSize + $recycleBinSize + $browserCacheSize
Write-Host ("  [A]  All of the above          {0,10}" -f (Format-Bytes $totalSize)) -ForegroundColor $ColorSchema.Warning
Write-Host ""

if ($Unattended) {
    $selections = @('A')
    Write-Host "  Unattended mode — cleaning all categories." -ForegroundColor $ColorSchema.Info
} else {
    Write-Host -NoNewline "  Enter selection(s) separated by commas (e.g. 1,3,A): " -ForegroundColor $ColorSchema.Menu
    $raw        = Read-Host
    $selections = $raw.Split(',') | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne '' }
}

Write-Host ""

$runAll       = $selections -contains 'A'
$totalFreed   = 0L

function Run-Category {
    param([bool]$ShouldRun, [string]$Label, [scriptblock]$Action)
    if (-not $ShouldRun) { return }
    if ($WhatIf) {
        Write-Host ("  [~] WhatIf: Would clean {0}" -f $Label) -ForegroundColor $ColorSchema.Warning
        return
    }
    Write-Host "  [*] Cleaning $Label..." -ForegroundColor $ColorSchema.Progress
    $freed = & $Action
    $totalFreed += $freed
    Write-Host ("  [+] {0} — freed {1}" -f $Label, (Format-Bytes $freed)) -ForegroundColor $ColorSchema.Success
}

Run-Category -ShouldRun ($runAll -or $selections -contains '1') -Label "User Temp Folders" -Action {
    Remove-FolderContents -Paths $userTempPaths -Label "User Temp"
}

Run-Category -ShouldRun ($runAll -or $selections -contains '2') -Label "System Temp" -Action {
    Remove-FolderContents -Paths $sysTempPaths -Label "System Temp"
}

Run-Category -ShouldRun ($runAll -or $selections -contains '3') -Label "Windows Update Cache" -Action {
    # Stop Windows Update service before clearing download cache
    $wuSvc = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
    $wasRunning = $wuSvc -and $wuSvc.Status -eq 'Running'
    if ($wasRunning) { Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue }
    $freed = Remove-FolderContents -Paths $wuCachePaths -Label "WU Cache"
    if ($wasRunning) { Start-Service -Name wuauserv -ErrorAction SilentlyContinue }
    $freed
}

Run-Category -ShouldRun ($runAll -or $selections -contains '4') -Label "Recycle Bin" -Action {
    $before = $recycleBinSize
    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    } catch {}
    $before
}

Run-Category -ShouldRun ($runAll -or $selections -contains '5') -Label "Browser Caches" -Action {
    Remove-FolderContents -Paths $browserCachePaths -Label "Browser Cache"
}

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  CLEANSE COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($WhatIf) {
    Write-Host "  Mode              : DRY RUN — no files were deleted" -ForegroundColor $ColorSchema.Warning
} else {
    Write-Host "  Total Space Freed : $(Format-Bytes $totalFreed)" -ForegroundColor $ColorSchema.Success
}
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
