<#
.SYNOPSIS
    A.R.C.H.I.V.E. — Automated Repository Compressing & Housing Important Volume Exports
    Pre-Reimaging Profile Backup Tool for PowerShell 5.1+

.DESCRIPTION
    Creates a compressed ZIP backup of a selected user profile before a machine
    is reimaged or wiped. Supports local and UNC share destinations. Generates
    a manifest file listing every archived item and a timestamped CSV log.

.USAGE
    PS C:\> .\archive.ps1                                                           # Must be run as Administrator
    PS C:\> .\archive.ps1 -Unattended -Username "John"                              # Archive all items for user John
    PS C:\> .\archive.ps1 -Unattended -Username "John" -Items "1,2,3" -Destination "\\server\backup"
    PS C:\> .\archive.ps1 -WhatIf                                                   # Preview actions without staging or compressing

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    A.U.S.P.E.X.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    R.E.V.E.N.A.N.T.       — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    A.R.T.I.F.A.C.T.       — Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.           — Toolkit setup & configuration wizard

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [string]$Username    = "",
    [string]$Items       = "A",
    [string]$Destination = "",
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
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
Assert-AdminPrivilege

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

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

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
# ARCHIVE LOG
# ─────────────────────────────────────────────────────────────────────────────

$ArchiveLog = New-Object System.Collections.ArrayList

function Add-ArchiveRecord {
    param(
        [string]$Item,
        [string]$Status,
        [string]$Detail    = "",
        [string]$Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    )
    [void]$ArchiveLog.Add([PSCustomObject]@{
        Timestamp = $Timestamp
        Item      = $Item
        Status    = $Status
        Detail    = $Detail
    })
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-ArchiveBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   █████╗ ██████╗  ██████╗██╗  ██╗██╗██╗   ██╗███████╗
  ██╔══██╗██╔══██╗██╔════╝██║  ██║██║██║   ██║██╔════╝
  ███████║██████╔╝██║     ███████║██║██║   ██║█████╗
  ██╔══██║██╔══██╗██║     ██╔══██║██║╚██╗ ██╔╝██╔══╝
  ██║  ██║██║  ██║╚██████╗██║  ██║██║ ╚████╔╝ ███████╗
  ╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝╚═╝  ╚═╝╚═╝  ╚═══╝  ╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    A.R.C.H.I.V.E. — Automated Repository Compressing & Housing Important Volume Exports" -ForegroundColor Cyan
    Write-Host "    Pre-Reimaging Profile Backup Tool" -ForegroundColor Cyan
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

function Get-OneDriveBusinessPath {
    param([string]$ProfileRoot)

    # Registry is most reliable — only works when the source is the current user's profile
    if ($ProfileRoot -ieq $env:USERPROFILE) {
        foreach ($acct in @('Business1', 'Business2', 'Personal')) {
            $reg = "HKCU:\Software\Microsoft\OneDrive\Accounts\$acct"
            if (Test-Path $reg) {
                $folder = (Get-ItemProperty $reg -ErrorAction SilentlyContinue).UserFolder
                if ($folder -and (Test-Path $folder)) { return $folder }
            }
        }
    }

    # Fallback: scan profile root for any folder named OneDrive*
    $match = Get-ChildItem -Path $ProfileRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'OneDrive*' } |
        Select-Object -First 1

    return if ($match) { $match.FullName } else { $null }
}

function Test-KnownFolderMove {
    param([string]$ProfileRoot)
    $desktop   = Join-Path $ProfileRoot 'Desktop'
    $documents = Join-Path $ProfileRoot 'Documents'
    $oneDrive  = Get-OneDriveBusinessPath -ProfileRoot $ProfileRoot
    if (-not $oneDrive) { return $false }
    return ($desktop -like "$oneDrive*") -or ($documents -like "$oneDrive*")
}

function Get-FolderSizeMB {
    param([string]$FolderPath)
    try {
        $bytes = (Get-ChildItem -Path $FolderPath -Recurse -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        return [math]::Round($bytes / 1MB, 1)
    }
    catch { return 0 }
}

function Stage-Item {
    param(
        [string]$SourcePath,
        [string]$StageDest,
        [string]$Label
    )

    if (-not (Test-Path $SourcePath)) {
        Write-Host "    [!!] $Label — not found, skipping." -ForegroundColor $ColorSchema.Warning
        Add-ArchiveRecord -Item $Label -Status "Skipped" -Detail "Source not found"
        return
    }

    Write-Host "    [*] Staging $Label..." -ForegroundColor $ColorSchema.Progress

    if ($WhatIf) {
        Write-Host "    [~] Would stage $Label from $SourcePath" -ForegroundColor Cyan
        Add-ArchiveRecord -Item $Label -Status "WhatIf" -Detail "Would stage from: $SourcePath"
        return
    }

    try {
        $null = New-Item -ItemType Directory -Path $StageDest -Force -ErrorAction Stop

        if (Test-Path $SourcePath -PathType Container) {
            $robocopyArgs = @($SourcePath, $StageDest, '/E', '/R:1', '/W:2', '/NP', '/NFL', '/NDL', '/NJH', '/NJS')
            & robocopy @robocopyArgs | Out-Null
            $ok = $LASTEXITCODE -le 7
        } else {
            Copy-Item -Path $SourcePath -Destination $StageDest -Force -ErrorAction Stop
            $ok = $true
        }

        if ($ok) {
            Write-Host "    [+] $Label — staged." -ForegroundColor $ColorSchema.Success
            Add-ArchiveRecord -Item $Label -Status "Staged"
        } else {
            Write-Host "    [!!] $Label — staged with warnings." -ForegroundColor $ColorSchema.Warning
            Add-ArchiveRecord -Item $Label -Status "Partial" -Detail "Robocopy exit $LASTEXITCODE"
        }
    }
    catch {
        Write-Host "    [-] $Label — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-ArchiveRecord -Item $Label -Status "Failed" -Detail $_
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-ArchiveBanner }

if ($WhatIf) {
    Write-Host ""
    Write-Host "  *** DRY RUN MODE — No files will be staged or compressed ***" -ForegroundColor Cyan
    Write-Host ""
}

Write-Host "  [!!] Run this tool BEFORE reimaging or wiping the machine." -ForegroundColor $ColorSchema.Warning
Write-Host "       Ensure the destination has sufficient free space." -ForegroundColor $ColorSchema.Warning
Write-Host ""

# ── PROFILE SELECTION ─────────────────────────────────────────────────────────

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SELECT PROFILE TO ARCHIVE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$profiles = Get-LocalProfiles

if ($profiles.Count -eq 0) {
    Write-Host "  [-] No user profiles found." -ForegroundColor $ColorSchema.Error
    exit 1
}

$profileRoot     = ""
$profileUsername = ""

if ($Unattended) {
    if ([string]::IsNullOrWhiteSpace($Username)) {
        Write-Host "  [-] -Username is required in unattended mode." -ForegroundColor $ColorSchema.Error
        exit 1
    }
    $selectedProfile = $profiles | Where-Object { $_.Username -ieq $Username } | Select-Object -First 1
    if (-not $selectedProfile) {
        Write-Host "  [-] Profile not found for username: $Username" -ForegroundColor $ColorSchema.Error
        exit 1
    }
    $profileRoot     = $selectedProfile.Path
    $profileUsername = $selectedProfile.Username
    Write-Host "  [+] Profile: $profileRoot" -ForegroundColor $ColorSchema.Success
} else {
    for ($i = 0; $i -lt $profiles.Count; $i++) {
        $lastUseStr = if ($profiles[$i].LastUse) { $profiles[$i].LastUse.ToString("yyyy-MM-dd") } else { "Never" }
        Write-Host ("  [{0,2}]  {1,-22}  {2,-42}  Last use: {3}" -f ($i + 1), $profiles[$i].Username, $profiles[$i].Path, $lastUseStr) -ForegroundColor $ColorSchema.Info
    }

    Write-Host ""
    Write-Host -NoNewline "  Select profile number: " -ForegroundColor $ColorSchema.Header
    $idx = (Read-Host).Trim()

    if (-not ($idx -match '^\d+$' -and [int]$idx -ge 1 -and [int]$idx -le $profiles.Count)) {
        Write-Host ""
        Write-Host "  [-] Invalid selection." -ForegroundColor $ColorSchema.Error
        exit 1
    }

    $selectedProfile  = $profiles[[int]$idx - 1]
    $profileRoot      = $selectedProfile.Path
    $profileUsername  = $selectedProfile.Username

    Write-Host ""
    Write-Host "  [+] Profile: $profileRoot" -ForegroundColor $ColorSchema.Success
}

$oneDrivePath = Get-OneDriveBusinessPath -ProfileRoot $profileRoot
if ($oneDrivePath) {
    Write-Host "  [*] OneDrive for Business detected: $oneDrivePath" -ForegroundColor $ColorSchema.Info
    if (Test-KnownFolderMove -ProfileRoot $profileRoot) {
        Write-Host "  [!!] Known Folder Move is active — Desktop/Documents are already inside OneDrive." -ForegroundColor $ColorSchema.Warning
        Write-Host "       Selecting both [1]/[2] and [12] will duplicate those folders in the archive." -ForegroundColor $ColorSchema.Warning
    }
}

# ── ITEM SELECTION ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SELECT ITEMS TO ARCHIVE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if ($Unattended) {
    $rawInput = $Items.ToUpper()
} else {
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
    Write-Host "  [12] OneDrive for Business" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $rawInput = (Read-Host).Trim().ToUpper()
}

$selectedItems = @()

if ($rawInput -eq "A") {
    $selectedItems = 1..12
} else {
    $selectedItems = $rawInput -split ',' |
        ForEach-Object { $_.Trim() } |
        Where-Object   { $_ -match '^\d+$' } |
        ForEach-Object { [int]$_ } |
        Where-Object   { $_ -ge 1 -and $_ -le 12 } |
        Sort-Object -Unique
}

if ($selectedItems.Count -eq 0) {
    Write-Host ""
    Write-Host "  [-] No valid items selected." -ForegroundColor $ColorSchema.Error
    exit 1
}

# ── DESTINATION ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  BACKUP DESTINATION" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$DestRoot = ""

$_cfg = Get-TKConfig

if ($Unattended) {
    if ([string]::IsNullOrWhiteSpace($Destination)) {
        # Fall back to config default, then script directory
        if (-not [string]::IsNullOrWhiteSpace($_cfg.Archive.DefaultDestination)) {
            $Destination = $_cfg.Archive.DefaultDestination
            Write-Host "  [*] No -Destination provided — using config default: $Destination" -ForegroundColor $ColorSchema.Info
        }
    }
    if ([string]::IsNullOrWhiteSpace($Destination)) {
        $DestRoot = $ScriptPath
        Write-Host "  [*] No -Destination provided — using script directory." -ForegroundColor $ColorSchema.Info
    } else {
        $DestRoot = $Destination.TrimEnd('\')
        if (-not (Test-Path $DestRoot)) {
            try {
                $null = New-Item -ItemType Directory -Path $DestRoot -Force -ErrorAction Stop
                Write-Host "  [+] Destination created: $DestRoot" -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Could not create destination: $_" -ForegroundColor $ColorSchema.Error
                exit 1
            }
        }
    }
    Write-Host "  [+] Destination: $DestRoot" -ForegroundColor $ColorSchema.Success
} else {
    Write-Host "  [1] Script directory  ($ScriptPath)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Enter a custom path  (local or UNC share)" -ForegroundColor $ColorSchema.Info
    if (-not [string]::IsNullOrWhiteSpace($_cfg.Archive.DefaultDestination)) {
        Write-Host "  [3] Config default  ($($_cfg.Archive.DefaultDestination))" -ForegroundColor $ColorSchema.Info
    }
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $destChoice = (Read-Host).Trim()

    if ($destChoice -eq "1") {
        $DestRoot = $ScriptPath
    }
    elseif ($destChoice -eq "3" -and -not [string]::IsNullOrWhiteSpace($_cfg.Archive.DefaultDestination)) {
        $DestRoot = $_cfg.Archive.DefaultDestination.TrimEnd('\')
        if (-not (Test-Path $DestRoot)) {
            try { $null = New-Item -ItemType Directory -Path $DestRoot -Force -ErrorAction Stop }
            catch { Write-Host "  [-] Could not create destination: $_" -ForegroundColor $ColorSchema.Error; exit 1 }
        }
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

        if (-not (Test-Path $DestRoot)) {
            Write-Host ""
            Write-Host "  [*] Destination not found — attempting to create it..." -ForegroundColor $ColorSchema.Progress
            try {
                $null = New-Item -ItemType Directory -Path $DestRoot -Force -ErrorAction Stop
                Write-Host "  [+] Destination created." -ForegroundColor $ColorSchema.Success
            }
            catch {
                Write-Host "  [-] Could not create destination: $_" -ForegroundColor $ColorSchema.Error
                exit 1
            }
        }
    }
    else {
        Write-Host ""
        Write-Host "  [-] Invalid selection." -ForegroundColor $ColorSchema.Error
        exit 1
    }

    Write-Host ""
    Write-Host "  [+] Destination: $DestRoot" -ForegroundColor $ColorSchema.Success
}

# ── STAGE FILES ───────────────────────────────────────────────────────────────

$timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
$archiveName  = "ARCHIVE_$($env:COMPUTERNAME)_${profileUsername}_$timestamp"
$stagingDir   = Join-Path $env:TEMP $archiveName

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  STAGING FILES" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$sourceAppData = Join-Path $profileRoot "AppData\Roaming"
$sourceLocal   = Join-Path $profileRoot "AppData\Local"

$itemMap = [ordered]@{
    1  = @{ Label = "Desktop";          Src = (Join-Path $profileRoot "Desktop");                                            Dst = "Desktop" }
    2  = @{ Label = "Documents";        Src = (Join-Path $profileRoot "Documents");                                          Dst = "Documents" }
    3  = @{ Label = "Downloads";        Src = (Join-Path $profileRoot "Downloads");                                          Dst = "Downloads" }
    4  = @{ Label = "Pictures";         Src = (Join-Path $profileRoot "Pictures");                                           Dst = "Pictures" }
    5  = @{ Label = "Videos";           Src = (Join-Path $profileRoot "Videos");                                             Dst = "Videos" }
    6  = @{ Label = "Music";            Src = (Join-Path $profileRoot "Music");                                              Dst = "Music" }
    7  = @{ Label = "Outlook Profiles"; Src = (Join-Path $sourceAppData "Microsoft\Outlook");                                Dst = "Outlook" }
    8  = @{ Label = "Email Signatures"; Src = (Join-Path $sourceAppData "Microsoft\Signatures");                             Dst = "Signatures" }
    9  = @{ Label = "Chrome Bookmarks"; Src = (Join-Path $sourceLocal   "Google\Chrome\User Data\Default\Bookmarks");        Dst = "Chrome" }
    10 = @{ Label = "Edge Bookmarks";   Src = (Join-Path $sourceLocal   "Microsoft\Edge\User Data\Default\Bookmarks");       Dst = "Edge" }
    11 = @{ Label = "Firefox Profiles";       Src = (Join-Path $sourceAppData "Mozilla\Firefox\Profiles");                    Dst = "Firefox" }
    12 = @{ Label = "OneDrive for Business";   Src = $(if ($oneDrivePath) { $oneDrivePath } else { "" });                        Dst = "OneDrive" }
}

foreach ($num in $selectedItems) {
    if (-not $itemMap.Contains($num)) { continue }
    $item    = $itemMap[$num]
    $stageDst = Join-Path $stagingDir $item.Dst
    Stage-Item -SourcePath $item.Src -StageDest $stageDst -Label $item.Label
}

# ── MANIFEST ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  [*] Writing manifest..." -ForegroundColor $ColorSchema.Progress

$manifestPath = Join-Path $stagingDir "MANIFEST.txt"
$manifestLines = @(
    "A.R.C.H.I.V.E. Backup Manifest",
    "================================",
    "Machine   : $env:COMPUTERNAME",
    "Profile   : $profileRoot",
    "Username  : $profileUsername",
    "Timestamp : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Archive   : $archiveName.zip",
    "",
    "Items Archived",
    "--------------"
)
foreach ($record in $ArchiveLog) {
    $manifestLines += "{0,-30} [{1}]  {2}" -f $record.Item, $record.Status, $record.Detail
}

try {
    $manifestLines | Set-Content -Path $manifestPath -Encoding UTF8
    Write-Host "  [+] Manifest written." -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Could not write manifest: $_" -ForegroundColor $ColorSchema.Error
}

# ── COMPRESS ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  COMPRESSING ARCHIVE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$zipPath = Join-Path $DestRoot "$archiveName.zip"

# Pre-flight: count staged files and total size so the technician knows what they're waiting for
$stagedFiles   = Get-ChildItem -Recurse -File -Path $stagingDir -ErrorAction SilentlyContinue
$stagedCount   = $stagedFiles.Count
$stagedSizeMB  = [math]::Round(($stagedFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 1)

Write-Host "  [*] Creating ZIP: $zipPath" -ForegroundColor $ColorSchema.Progress
Write-Host ("  [*] Compressing {0} files ({1} MB) — this may take several minutes..." -f $stagedCount, $stagedSizeMB) -ForegroundColor $ColorSchema.Info
Write-Host ""

if ($WhatIf) {
    Write-Host "  [~] Would compress staged files into: $zipPath" -ForegroundColor Cyan
    Write-Host ""
    Add-ArchiveRecord -Item "ZIP Archive" -Status "WhatIf" -Detail "Would create: $zipPath ($stagedCount files, $stagedSizeMB MB)"
} else {

$compressionStart = Get-Date

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    [System.IO.Compression.ZipFile]::CreateFromDirectory($stagingDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

    $zipSizeMB      = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
    $compressionSec = [math]::Round(((Get-Date) - $compressionStart).TotalSeconds, 1)
    Write-Host ("  [+] Archive created: {0}  ({1} MB, {2}s)" -f $zipPath, $zipSizeMB, $compressionSec) -ForegroundColor $ColorSchema.Success
    Add-ArchiveRecord -Item "ZIP Archive" -Status "Created" -Detail "$zipSizeMB MB — $zipPath"
}
catch {
    Write-Host "  [-] Compression failed: $_" -ForegroundColor $ColorSchema.Error
    Add-ArchiveRecord -Item "ZIP Archive" -Status "Failed" -Detail $_
}
finally {
    Write-Host "  [*] Cleaning up staging folder..." -ForegroundColor $ColorSchema.Progress
    try {
        Remove-Item -Path $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  [+] Staging folder removed." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [!!] Could not remove staging folder: $stagingDir" -ForegroundColor $ColorSchema.Warning
    }
}
}  # end else (not WhatIf)

# ── LOG ───────────────────────────────────────────────────────────────────────

$logFile = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "ARCHIVE_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

try {
    $ArchiveLog | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
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
Write-Host "  ARCHIVE SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

foreach ($record in $ArchiveLog) {
    $color = switch ($record.Status) {
        "Staged"   { $ColorSchema.Success }
        "Created"  { $ColorSchema.Success }
        "Partial"  { $ColorSchema.Warning }
        "Skipped"  { $ColorSchema.Info    }
        "WhatIf"   { 'Cyan'               }
        default    { $ColorSchema.Error   }
    }
    $detail = if ($record.Detail) { " — $($record.Detail)" } else { "" }
    Write-Host ("  {0,-30} [{1}]{2}" -f $record.Item, $record.Status, $detail) -ForegroundColor $color
}

Write-Host ""
$staged  = ($ArchiveLog | Where-Object { $_.Status -eq "Staged"  } | Measure-Object).Count
$skipped = ($ArchiveLog | Where-Object { $_.Status -eq "Skipped" } | Measure-Object).Count
$failed  = ($ArchiveLog | Where-Object { $_.Status -eq "Failed"  } | Measure-Object).Count

Write-Host "  Staged: $staged  |  Skipped: $skipped  |  Failed: $failed" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  A.R.C.H.I.V.E. COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
