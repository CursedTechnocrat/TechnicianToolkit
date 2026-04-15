<#
.SYNOPSIS
    P.H.A.N.T.O.M. — Portable Home Archive: Navigates & Transfers Objects to new Machine
    Profile Migration Tool for PowerShell 5.1+

.DESCRIPTION
    Migrates user profile data from a source profile or machine to a destination.
    Handles common user folders, Outlook data, browser bookmarks, and email signatures.
    Uses Robocopy for reliable folder transfers and generates a timestamped CSV log.

.USAGE
    PS C:\> .\phantom.ps1      # Must be run as Administrator

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    P.H.A.N.T.O.M.         — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT PATH RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

if ($PSScriptRoot) {
    $ScriptPath = $PSScriptRoot
} elseif ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
} else {
    $ScriptPath = (Get-Location).Path
}

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
# MIGRATION LOG
# ─────────────────────────────────────────────────────────────────────────────

$MigrationLog = New-Object System.Collections.ArrayList

function Add-MigrationRecord {
    param(
        [string]$Item,
        [string]$Status,
        [string]$Detail    = "",
        [string]$Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    )
    [void]$MigrationLog.Add([PSCustomObject]@{
        Timestamp = $Timestamp
        Item      = $Item
        Status    = $Status
        Detail    = $Detail
    })
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-PhantomBanner {
    Clear-Host
    Write-Host @"

  ██████╗ ██╗  ██╗ █████╗ ███╗   ██╗████████╗ ██████╗ ███╗   ███╗
  ██╔══██╗██║  ██║██╔══██╗████╗  ██║╚══██╔══╝██╔═══██╗████╗ ████║
  ██████╔╝███████║███████║██╔██╗ ██║   ██║   ██║   ██║██╔████╔██║
  ██╔═══╝ ██╔══██║██╔══██║██║╚██╗██║   ██║   ██║   ██║██║╚██╔╝██║
  ██║     ██║  ██║██║  ██║██║ ╚████║   ██║   ╚██████╔╝██║ ╚═╝ ██║
  ╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝   ╚═╝    ╚═════╝ ╚═╝     ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    P.H.A.N.T.O.M. — Portable Home Archive: Navigates & Transfers Objects to new Machine" -ForegroundColor Cyan
    Write-Host "    Profile Migration & Data Transfer Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Get-LocalProfiles {
    Get-CimInstance -ClassName Win32_UserProfile |
        Where-Object { -not $_.Special -and (Test-Path $_.LocalPath) } |
        Select-Object @{N='Username'; E={ Split-Path $_.LocalPath -Leaf }},
                      @{N='Path';     E={ $_.LocalPath }},
                      @{N='LastUse';  E={ $_.LastUseTime }} |
        Sort-Object LastUse -Descending
}

function Copy-ProfileFolder {
    param(
        [string]$SourcePath,
        [string]$DestPath,
        [string]$Label
    )

    if (-not (Test-Path $SourcePath)) {
        Write-Host "    [!!] $Label — source not found, skipping." -ForegroundColor $ColorSchema.Warning
        Add-MigrationRecord -Item $Label -Status "Skipped" -Detail "Source path not found"
        return
    }

    Write-Host "    [*] Copying $Label..." -ForegroundColor $ColorSchema.Progress

    try {
        $null = New-Item -ItemType Directory -Path $DestPath -Force -ErrorAction Stop
        $robocopyArgs = @($SourcePath, $DestPath, '/E', '/R:2', '/W:3', '/NP', '/NFL', '/NDL', '/NJH', '/NJS')
        & robocopy @robocopyArgs | Out-Null

        if ($LASTEXITCODE -le 7) {
            Write-Host "    [+] $Label — copied successfully." -ForegroundColor $ColorSchema.Success
            Add-MigrationRecord -Item $Label -Status "Copied"
        } else {
            Write-Host "    [!!] $Label — completed with warnings (exit $LASTEXITCODE)." -ForegroundColor $ColorSchema.Warning
            Add-MigrationRecord -Item $Label -Status "Partial" -Detail "Robocopy exit code $LASTEXITCODE"
        }
    }
    catch {
        Write-Host "    [-] $Label — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-MigrationRecord -Item $Label -Status "Failed" -Detail $_
    }
}

function Copy-ProfileFile {
    param(
        [string]$SourceFile,
        [string]$DestFolder,
        [string]$Label
    )

    if (-not (Test-Path $SourceFile)) {
        Write-Host "    [!!] $Label — not found, skipping." -ForegroundColor $ColorSchema.Warning
        Add-MigrationRecord -Item $Label -Status "Skipped" -Detail "Source file not found"
        return
    }

    try {
        $null = New-Item -ItemType Directory -Path $DestFolder -Force -ErrorAction Stop
        Copy-Item -Path $SourceFile -Destination $DestFolder -Force -ErrorAction Stop
        Write-Host "    [+] $Label — copied successfully." -ForegroundColor $ColorSchema.Success
        Add-MigrationRecord -Item $Label -Status "Copied"
    }
    catch {
        Write-Host "    [-] $Label — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-MigrationRecord -Item $Label -Status "Failed" -Detail $_
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-PhantomBanner

# ── SOURCE ────────────────────────────────────────────────────────────────────

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SOURCE PROFILE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  [1] Select from local profiles on this machine" -ForegroundColor $ColorSchema.Info
Write-Host "  [2] Enter a custom or remote path  (e.g. \\OldPC\C`$\Users\John)" -ForegroundColor $ColorSchema.Info
Write-Host "  [3] Restore from an A.R.C.H.I.V.E. ZIP" -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
$sourceChoice = (Read-Host).Trim()

$SourceRoot      = ""
$IsArchiveZip    = $false
$TempExtractDir  = ""

if ($sourceChoice -eq "1") {
    $profiles = Get-LocalProfiles

    if ($profiles.Count -eq 0) {
        Write-Host ""
        Write-Host "  [-] No user profiles found on this machine." -ForegroundColor $ColorSchema.Error
        exit 1
    }

    Write-Host ""
    Write-Host "  Available profiles:" -ForegroundColor $ColorSchema.Info
    Write-Host ""

    for ($i = 0; $i -lt $profiles.Count; $i++) {
        $lastUseStr = if ($profiles[$i].LastUse) { $profiles[$i].LastUse.ToString("yyyy-MM-dd") } else { "Never" }
        Write-Host ("  [{0,2}]  {1,-22}  {2,-45}  Last use: {3}" -f ($i + 1), $profiles[$i].Username, $profiles[$i].Path, $lastUseStr) -ForegroundColor $ColorSchema.Info
    }

    Write-Host ""
    Write-Host -NoNewline "  Select profile number: " -ForegroundColor $ColorSchema.Header
    $idx = (Read-Host).Trim()

    if ($idx -match '^\d+$' -and [int]$idx -ge 1 -and [int]$idx -le $profiles.Count) {
        $SourceRoot = $profiles[[int]$idx - 1].Path
    } else {
        Write-Host ""
        Write-Host "  [-] Invalid selection." -ForegroundColor $ColorSchema.Error
        exit 1
    }
}
elseif ($sourceChoice -eq "2") {
    Write-Host ""
    Write-Host -NoNewline "  Enter source path: " -ForegroundColor $ColorSchema.Header
    $SourceRoot = (Read-Host).Trim().TrimEnd('\')

    if (-not (Test-Path $SourceRoot)) {
        Write-Host ""
        Write-Host "  [-] Path not accessible: $SourceRoot" -ForegroundColor $ColorSchema.Error
        exit 1
    }
}
elseif ($sourceChoice -eq "3") {
    Write-Host ""
    Write-Host -NoNewline "  Enter path to A.R.C.H.I.V.E. ZIP file: " -ForegroundColor $ColorSchema.Header
    $zipSource = (Read-Host).Trim().Trim('"')

    if (-not (Test-Path $zipSource)) {
        Write-Host ""
        Write-Host "  [-] File not found: $zipSource" -ForegroundColor $ColorSchema.Error
        exit 1
    }
    if ([System.IO.Path]::GetExtension($zipSource) -ine ".zip") {
        Write-Host ""
        Write-Host "  [-] File does not appear to be a ZIP archive." -ForegroundColor $ColorSchema.Error
        exit 1
    }

    $TempExtractDir = Join-Path $env:TEMP "PHANTOM_Extract_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    try {
        Write-Host ""
        Write-Host "  [*] Extracting archive — this may take a moment..." -ForegroundColor $ColorSchema.Progress
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipSource, $TempExtractDir)
        $SourceRoot   = $TempExtractDir
        $IsArchiveZip = $true
        Write-Host "  [+] Archive extracted." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [-] Failed to extract archive: $_" -ForegroundColor $ColorSchema.Error
        exit 1
    }
}
else {
    Write-Host ""
    Write-Host "  [-] Invalid selection." -ForegroundColor $ColorSchema.Error
    exit 1
}

Write-Host ""
Write-Host "  [+] Source: $SourceRoot" -ForegroundColor $ColorSchema.Success

# ── DESTINATION ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  DESTINATION" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  [1] Current user's profile  ($env:USERPROFILE)" -ForegroundColor $ColorSchema.Info
Write-Host "  [2] Enter a custom path" -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
$destChoice = (Read-Host).Trim()

$DestRoot = ""

if ($destChoice -eq "1") {
    $DestRoot = $env:USERPROFILE
}
elseif ($destChoice -eq "2") {
    Write-Host ""
    Write-Host -NoNewline "  Enter destination path: " -ForegroundColor $ColorSchema.Header
    $DestRoot = (Read-Host).Trim().TrimEnd('\')

    if ([string]::IsNullOrWhiteSpace($DestRoot)) {
        Write-Host ""
        Write-Host "  [-] No path entered." -ForegroundColor $ColorSchema.Error
        exit 1
    }
}
else {
    Write-Host ""
    Write-Host "  [-] Invalid selection." -ForegroundColor $ColorSchema.Error
    exit 1
}

if ($SourceRoot -ieq $DestRoot) {
    Write-Host ""
    Write-Host "  [-] Source and destination cannot be the same path." -ForegroundColor $ColorSchema.Error
    exit 1
}

Write-Host ""
Write-Host "  [+] Destination: $DestRoot" -ForegroundColor $ColorSchema.Success

# ── ITEM SELECTION ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SELECT ITEMS TO MIGRATE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Enter numbers separated by commas, or A for all." -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host "  [1]  Desktop" -ForegroundColor $ColorSchema.Info
Write-Host "  [2]  Documents" -ForegroundColor $ColorSchema.Info
Write-Host "  [3]  Downloads" -ForegroundColor $ColorSchema.Info
Write-Host "  [4]  Pictures" -ForegroundColor $ColorSchema.Info
Write-Host "  [5]  Videos" -ForegroundColor $ColorSchema.Info
Write-Host "  [6]  Music" -ForegroundColor $ColorSchema.Info
Write-Host "  [7]  Outlook Profiles & Data Files" -ForegroundColor $ColorSchema.Info
Write-Host "  [8]  Email Signatures" -ForegroundColor $ColorSchema.Info
Write-Host "  [9]  Chrome Bookmarks" -ForegroundColor $ColorSchema.Info
Write-Host "  [10] Edge Bookmarks" -ForegroundColor $ColorSchema.Info
Write-Host "  [11] Firefox Profiles" -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
$rawInput = (Read-Host).Trim().ToUpper()

$selectedItems = @()

if ($rawInput -eq "A") {
    $selectedItems = 1..11
} else {
    $selectedItems = $rawInput -split ',' |
        ForEach-Object { $_.Trim() } |
        Where-Object   { $_ -match '^\d+$' } |
        ForEach-Object { [int]$_ } |
        Where-Object   { $_ -ge 1 -and $_ -le 11 } |
        Sort-Object -Unique
}

if ($selectedItems.Count -eq 0) {
    Write-Host ""
    Write-Host "  [-] No valid items selected." -ForegroundColor $ColorSchema.Error
    exit 1
}

# ── MIGRATE ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  MIGRATING" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$destMigration = Join-Path $DestRoot "PHANTOM_Migration"

# ARCHIVE ZIPs use a flat folder structure (Desktop, Outlook, Chrome, etc. at root).
# Live profiles use deep AppData paths. Build the item map accordingly.
if ($IsArchiveZip) {
    $itemMap = [ordered]@{
        1  = @{ Label = "Desktop";             Type = "Folder"; Src = (Join-Path $SourceRoot "Desktop");                         Dst = (Join-Path $DestRoot "Desktop") }
        2  = @{ Label = "Documents";           Type = "Folder"; Src = (Join-Path $SourceRoot "Documents");                       Dst = (Join-Path $DestRoot "Documents") }
        3  = @{ Label = "Downloads";           Type = "Folder"; Src = (Join-Path $SourceRoot "Downloads");                       Dst = (Join-Path $DestRoot "Downloads") }
        4  = @{ Label = "Pictures";            Type = "Folder"; Src = (Join-Path $SourceRoot "Pictures");                        Dst = (Join-Path $DestRoot "Pictures") }
        5  = @{ Label = "Videos";              Type = "Folder"; Src = (Join-Path $SourceRoot "Videos");                          Dst = (Join-Path $DestRoot "Videos") }
        6  = @{ Label = "Music";               Type = "Folder"; Src = (Join-Path $SourceRoot "Music");                           Dst = (Join-Path $DestRoot "Music") }
        7  = @{ Label = "Outlook Profiles";    Type = "Folder"; Src = (Join-Path $SourceRoot "Outlook");                         Dst = (Join-Path $destMigration "Outlook") }
        8  = @{ Label = "Email Signatures";    Type = "Folder"; Src = (Join-Path $SourceRoot "Signatures");                      Dst = (Join-Path $destMigration "Signatures") }
        9  = @{ Label = "Chrome Bookmarks";    Type = "File";   Src = (Join-Path $SourceRoot "Chrome\Bookmarks");                Dst = (Join-Path $destMigration "Chrome") }
        10 = @{ Label = "Edge Bookmarks";      Type = "File";   Src = (Join-Path $SourceRoot "Edge\Bookmarks");                  Dst = (Join-Path $destMigration "Edge") }
        11 = @{ Label = "Firefox Profiles";    Type = "Folder"; Src = (Join-Path $SourceRoot "Firefox");                         Dst = (Join-Path $destMigration "Firefox") }
    }
} else {
    $sourceAppData = Join-Path $SourceRoot "AppData\Roaming"
    $sourceLocal   = Join-Path $SourceRoot "AppData\Local"

    $itemMap = [ordered]@{
        1  = @{ Label = "Desktop";             Type = "Folder"; Src = (Join-Path $SourceRoot "Desktop");                                            Dst = (Join-Path $DestRoot "Desktop") }
        2  = @{ Label = "Documents";           Type = "Folder"; Src = (Join-Path $SourceRoot "Documents");                                          Dst = (Join-Path $DestRoot "Documents") }
        3  = @{ Label = "Downloads";           Type = "Folder"; Src = (Join-Path $SourceRoot "Downloads");                                          Dst = (Join-Path $DestRoot "Downloads") }
        4  = @{ Label = "Pictures";            Type = "Folder"; Src = (Join-Path $SourceRoot "Pictures");                                           Dst = (Join-Path $DestRoot "Pictures") }
        5  = @{ Label = "Videos";              Type = "Folder"; Src = (Join-Path $SourceRoot "Videos");                                             Dst = (Join-Path $DestRoot "Videos") }
        6  = @{ Label = "Music";               Type = "Folder"; Src = (Join-Path $SourceRoot "Music");                                              Dst = (Join-Path $DestRoot "Music") }
        7  = @{ Label = "Outlook Profiles";    Type = "Folder"; Src = (Join-Path $sourceAppData "Microsoft\Outlook");                               Dst = (Join-Path $destMigration "Outlook") }
        8  = @{ Label = "Email Signatures";    Type = "Folder"; Src = (Join-Path $sourceAppData "Microsoft\Signatures");                            Dst = (Join-Path $destMigration "Signatures") }
        9  = @{ Label = "Chrome Bookmarks";    Type = "File";   Src = (Join-Path $sourceLocal   "Google\Chrome\User Data\Default\Bookmarks");       Dst = (Join-Path $destMigration "Chrome") }
        10 = @{ Label = "Edge Bookmarks";      Type = "File";   Src = (Join-Path $sourceLocal   "Microsoft\Edge\User Data\Default\Bookmarks");      Dst = (Join-Path $destMigration "Edge") }
        11 = @{ Label = "Firefox Profiles";    Type = "Folder"; Src = (Join-Path $sourceAppData "Mozilla\Firefox\Profiles");                        Dst = (Join-Path $destMigration "Firefox") }
    }
}

foreach ($num in $selectedItems) {
    if (-not $itemMap.Contains($num)) { continue }
    $item = $itemMap[$num]

    if ($item.Type -eq "Folder") {
        Copy-ProfileFolder -SourcePath $item.Src -DestPath $item.Dst -Label $item.Label
    } else {
        Copy-ProfileFile -SourceFile $item.Src -DestFolder $item.Dst -Label $item.Label
    }
}

# ── CLEANUP ARCHIVE EXTRACT ───────────────────────────────────────────────────

if ($IsArchiveZip -and $TempExtractDir -and (Test-Path $TempExtractDir)) {
    Write-Host ""
    Write-Host "  [*] Removing temporary extract folder..." -ForegroundColor $ColorSchema.Progress
    try {
        Remove-Item -Path $TempExtractDir -Recurse -Force -ErrorAction Stop
        Write-Host "  [+] Temporary files cleaned up." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [!!] Could not remove temp folder — delete manually: $TempExtractDir" -ForegroundColor $ColorSchema.Warning
    }
}

# ── LOG ───────────────────────────────────────────────────────────────────────

$logFile = Join-Path $ScriptPath "PHANTOM_MigrationLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

try {
    $MigrationLog | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  [+] Log saved: $logFile" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host ""
    Write-Host "  [-] Could not save log: $_" -ForegroundColor $ColorSchema.Error
}

# ── SUMMARY ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  MIGRATION SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

foreach ($record in $MigrationLog) {
    $color = switch ($record.Status) {
        "Copied"  { $ColorSchema.Success }
        "Partial" { $ColorSchema.Warning }
        "Skipped" { $ColorSchema.Info    }
        default   { $ColorSchema.Error   }
    }
    $detail = if ($record.Detail) { " — $($record.Detail)" } else { "" }
    Write-Host ("  {0,-30} [{1}]{2}" -f $record.Item, $record.Status, $detail) -ForegroundColor $color
}

Write-Host ""
$copied  = ($MigrationLog | Where-Object { $_.Status -eq "Copied"  } | Measure-Object).Count
$partial = ($MigrationLog | Where-Object { $_.Status -eq "Partial" } | Measure-Object).Count
$skipped = ($MigrationLog | Where-Object { $_.Status -eq "Skipped" } | Measure-Object).Count
$failed  = ($MigrationLog | Where-Object { $_.Status -eq "Failed"  } | Measure-Object).Count

Write-Host "  Copied: $copied  |  Partial: $partial  |  Skipped: $skipped  |  Failed: $failed" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  P.H.A.N.T.O.M. COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

Read-Host "  Press Enter to exit"
