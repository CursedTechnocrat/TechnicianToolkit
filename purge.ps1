<#
.SYNOPSIS
    P.U.R.G.E. — Purges Unwanted Remnants, Garbage & Ephemeral data
    Disk Cleanup Tool for PowerShell 5.1+

.DESCRIPTION
    Cleans common junk accumulation points on Windows: user and system temp
    folders, Windows Update download cache, the Recycle Bin, and browser
    caches (Chrome, Edge, Firefox). Shows estimated space before each
    category and reports total freed space after cleanup.

.USAGE
    PS C:\> .\purge.ps1                    # Must be run as Administrator
    PS C:\> .\purge.ps1 -Unattended        # Silent mode — cleans all categories, no prompts
    PS C:\> .\purge.ps1 -WhatIf            # Preview what would be cleaned, without deleting anything

.NOTES
    Version : 1.1

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    D.W.A.R.F.             — Disk wear & health assessment, SMART status
    P.U.R.G.E.             — Disk cleanup — temp, update cache, browser caches

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param([switch]$Unattended, [switch]$WhatIf)

# ─────────────────────────────────────────────────────────────────────────────
# INITIALIZATION
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Invoke-AdminElevation -ScriptFile $PSCommandPath

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
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-PurgeBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██╗   ██╗██████╗  ██████╗ ███████╗
  ██╔══██╗██║   ██║██╔══██╗██╔════╝ ██╔════╝
  ██████╔╝██║   ██║██████╔╝██║  ███╗█████╗
  ██╔═══╝ ██║   ██║██╔══██╗██║   ██║██╔══╝
  ██║     ╚██████╔╝██║  ██║╚██████╔╝███████╗
  ╚═╝      ╚═════╝ ╚═╝  ╚═╝ ╚═════╝ ╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    P.U.R.G.E. — Purges Unwanted Remnants, Garbage & Ephemeral data" -ForegroundColor Cyan
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

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
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
    $paths = @()
    $profiles = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue

    foreach ($profile in $profiles) {
        $base = $profile.FullName
        # Chrome
        $chromeCacheBase = Join-Path $base 'AppData\Local\Google\Chrome\User Data'
        if (Test-Path $chromeCacheBase) {
            Get-ChildItem $chromeCacheBase -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^(Default|Profile)' } |
                ForEach-Object { $paths += Join-Path $_.FullName 'Cache'; $paths += Join-Path $_.FullName 'Code Cache' }
        }
        # Edge
        $edgeCacheBase = Join-Path $base 'AppData\Local\Microsoft\Edge\User Data'
        if (Test-Path $edgeCacheBase) {
            Get-ChildItem $edgeCacheBase -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^(Default|Profile)' } |
                ForEach-Object { $paths += Join-Path $_.FullName 'Cache'; $paths += Join-Path $_.FullName 'Code Cache' }
        }
        # Firefox
        $ffProfiles = Join-Path $base 'AppData\Local\Mozilla\Firefox\Profiles'
        if (Test-Path $ffProfiles) {
            Get-ChildItem $ffProfiles -Directory -ErrorAction SilentlyContinue |
                ForEach-Object { $paths += Join-Path $_.FullName 'cache2' }
        }
    }
    return $paths | Where-Object { Test-Path $_ }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-PurgeBanner }

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
Write-Host "  PURGE COMPLETE" -ForegroundColor $ColorSchema.Header
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
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
