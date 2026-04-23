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
    Version : 3.0

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
# SHARED MODULE BOOTSTRAP
# ===========================
$TKModulePath = Join-Path $PSScriptRoot 'TechnicianToolkit.psm1'
if (-not (Test-Path $TKModulePath)) {
    $TKModuleUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1'
    Write-Host "  [*] Shared module TechnicianToolkit.psm1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $TKModuleUrl -OutFile $TKModulePath -ErrorAction Stop
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($TKModulePath, [ref]$null, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            Remove-Item -Path $TKModulePath -Force -ErrorAction SilentlyContinue
            Write-Host "  [!!] Downloaded module failed syntax validation - file removed." -ForegroundColor Red
            Write-Host "       $($parseErrors[0].Message)" -ForegroundColor Red
            exit 1
        }
        Write-Host "  [+] Module downloaded and verified." -ForegroundColor Green
    } catch {
        Write-Host "  [!!] Could not download TechnicianToolkit.psm1:" -ForegroundColor Red
        Write-Host "       $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Place the module manually next to this script from:" -ForegroundColor Yellow
        Write-Host "       $TKModuleUrl" -ForegroundColor Yellow
        exit 1
    }
}
Import-Module $TKModulePath -Force -ErrorAction Stop
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
