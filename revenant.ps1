<#
.SYNOPSIS
    R.E.V.E.N.A.N.T. — Relocates, Extracts, Validates Environments, Networks, Accounts 'N Transfers
    Profile Migration Tool for PowerShell 5.1+

.DESCRIPTION
    Migrates user profile data from a source profile or machine to a destination.
    Handles common user folders, Outlook data, browser bookmarks, and email signatures.
    Uses Robocopy for reliable folder transfers and generates a timestamped CSV log.
    Can restore directly from an ARCHIVE ZIP in both interactive and unattended modes.

.USAGE
    PS C:\> .\revenant.ps1                                         # Must be run as Administrator
    PS C:\> .\revenant.ps1 -Unattended -SourcePath "C:\Users\John" -DestPath "D:\Migration"
    PS C:\> .\revenant.ps1 -Unattended -SourcePath "\\OldPC\C$\Users\John" -DestPath "C:\Users\John" -Items "1,2,3"
    PS C:\> .\revenant.ps1 -Unattended -ArchiveZip "D:\Backup\John_20260101.zip" -DestPath "C:\Users\John"

.NOTES
    Version : 1.2

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
    [string]$SourcePath = "",
    [string]$DestPath   = "",
    [string]$Items      = "A",
    [string]$ArchiveZip = "",
    [switch]$Transcript,
    [switch]$WhatIf
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

function Show-RevenantBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ███████╗██╗   ██╗███████╗███╗   ██╗ █████╗ ███╗   ██╗████████╗
  ██╔══██╗██╔════╝██║   ██║██╔════╝████╗  ██║██╔══██╗████╗  ██║╚══██╔══╝
  ██████╔╝█████╗  ██║   ██║█████╗  ██╔██╗ ██║███████║██╔██╗ ██║   ██║
  ██╔══██╗██╔══╝  ╚██╗ ██╔╝██╔══╝  ██║╚██╗██║██╔══██║██║╚██╗██║   ██║
  ██║  ██║███████╗ ╚████╔╝ ███████╗██║ ╚████║██║  ██║██║ ╚████║   ██║
  ╚═╝  ╚═╝╚══════╝  ╚═══╝  ╚══════╝╚═╝  ╚═══╝╚═╝  ╚═╝╚═╝  ╚═══╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    R.E.V.E.N.A.N.T. — Relocates, Extracts, Validates Environments, Networks, Accounts 'N Transfers" -ForegroundColor Cyan
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

    # Pre-flight scan so the progress bar has a denominator
    Write-Host "    [*] Scanning $Label..." -ForegroundColor $ColorSchema.Progress
    $sourceFiles = Get-ChildItem -Recurse -File -Path $SourcePath -ErrorAction SilentlyContinue
    $totalFiles  = $sourceFiles.Count
    $totalSizeMB = [math]::Round(($sourceFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 1)

    if ($WhatIf) {
        Write-Host "    [~] WhatIf: Would copy $Label  ($totalFiles files, $totalSizeMB MB) => $DestPath" -ForegroundColor $ColorSchema.Warning
        Add-MigrationRecord -Item $Label -Status "WhatIf" -Detail "$totalFiles files, $totalSizeMB MB"
        return
    }

    Write-Host "    [*] Copying $Label  ($totalFiles files, $totalSizeMB MB)..." -ForegroundColor $ColorSchema.Progress

    try {
        $null = New-Item -ItemType Directory -Path $DestPath -Force -ErrorAction Stop

        # /NFL removed so robocopy emits per-file lines for progress tracking
        $robocopyArgs = @($SourcePath, $DestPath, '/E', '/R:2', '/W:3', '/NP', '/NDL', '/NJH', '/NJS')
        $copied    = 0
        $startTime = Get-Date

        & robocopy @robocopyArgs | ForEach-Object {
            # Match lines where robocopy reports a file it is copying
            if ($_ -match '^\s+(New File|Newer|Older|Changed)\s') {
                $copied++
                $pct = if ($totalFiles -gt 0) { [math]::Min(99, [math]::Floor($copied / $totalFiles * 100)) } else { 0 }
                Write-Progress -Activity "Copying $Label" `
                               -Status   "$copied / $totalFiles files" `
                               -PercentComplete $pct
            }
        }

        Write-Progress -Activity "Copying $Label" -Completed
        $elapsed = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)

        if ($LASTEXITCODE -le 7) {
            Write-Host "    [+] $Label — done  ($copied files copied, $totalSizeMB MB, ${elapsed}s)." -ForegroundColor $ColorSchema.Success
            Add-MigrationRecord -Item $Label -Status "Copied" -Detail "$totalFiles files, $totalSizeMB MB"
        } else {
            Write-Progress -Activity "Copying $Label" -Completed
            Write-Host "    [!!] $Label — completed with warnings (exit $LASTEXITCODE)." -ForegroundColor $ColorSchema.Warning
            Add-MigrationRecord -Item $Label -Status "Partial" -Detail "Robocopy exit $LASTEXITCODE, $totalFiles files, $totalSizeMB MB"
        }
    }
    catch {
        Write-Progress -Activity "Copying $Label" -Completed
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

    if ($WhatIf) {
        Write-Host "    [~] WhatIf: Would copy $Label => $DestFolder" -ForegroundColor $ColorSchema.Warning
        Add-MigrationRecord -Item $Label -Status "WhatIf"
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

    # Fallback: scan profile root for any folder named OneDrive* (covers other profiles and UNC sources)
    $match = Get-ChildItem -Path $ProfileRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'OneDrive*' } |
        Select-Object -First 1

    return if ($match) { $match.FullName } else { $null }
}

function Test-KnownFolderMove {
    param([string]$ProfileRoot)
    # KFM redirects Desktop/Documents into the OneDrive folder — detect by comparing paths
    $desktop   = Join-Path $ProfileRoot 'Desktop'
    $documents = Join-Path $ProfileRoot 'Documents'
    $oneDrive  = Get-OneDriveBusinessPath -ProfileRoot $ProfileRoot
    if (-not $oneDrive) { return $false }
    return ($desktop -like "$oneDrive*") -or ($documents -like "$oneDrive*")
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-RevenantBanner }

# ── SOURCE ────────────────────────────────────────────────────────────────────

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SOURCE PROFILE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$SourceRoot      = ""
$IsArchiveZip    = $false
$TempExtractDir  = ""

if ($Unattended) {
    if (-not [string]::IsNullOrWhiteSpace($ArchiveZip)) {
        if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
            Write-Host "  [!!] Both -ArchiveZip and -SourcePath were provided — -ArchiveZip takes precedence." -ForegroundColor $ColorSchema.Warning
        }
        if (-not (Test-Path $ArchiveZip)) {
            Write-Host "  [-] Archive file not found: $ArchiveZip" -ForegroundColor $ColorSchema.Error
            exit 1
        }
        if ([System.IO.Path]::GetExtension($ArchiveZip) -ine ".zip") {
            Write-Host "  [-] File does not appear to be a ZIP archive." -ForegroundColor $ColorSchema.Error
            exit 1
        }
        $TempExtractDir = Join-Path $env:TEMP "REVENANT_Extract_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        try {
            Write-Host "  [*] Extracting archive — this may take a moment..." -ForegroundColor $ColorSchema.Progress
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            [System.IO.Compression.ZipFile]::ExtractToDirectory($ArchiveZip, $TempExtractDir)
            $SourceRoot   = $TempExtractDir
            $IsArchiveZip = $true
            Write-Host "  [+] Archive extracted." -ForegroundColor $ColorSchema.Success
        }
        catch {
            Write-Host "  [-] Failed to extract archive: $_" -ForegroundColor $ColorSchema.Error
            exit 1
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
        $SourceRoot   = $SourcePath.TrimEnd('\')
        $IsArchiveZip = $false
        if (-not (Test-Path $SourceRoot)) {
            Write-Host "  [-] Source path not accessible: $SourceRoot" -ForegroundColor $ColorSchema.Error
            exit 1
        }
    } else {
        Write-Host "  [-] Either -SourcePath or -ArchiveZip is required in unattended mode." -ForegroundColor $ColorSchema.Error
        exit 1
    }
    Write-Host "  [+] Source: $SourceRoot" -ForegroundColor $ColorSchema.Success
} else {
    Write-Host "  [1] Select from local profiles on this machine" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Enter a custom or remote path  (e.g. \\OldPC\C`$\Users\John)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [3] Restore from an A.R.C.H.I.V.E. ZIP" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $sourceChoice = (Read-Host).Trim()

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

        $TempExtractDir = Join-Path $env:TEMP "REVENANT_Extract_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

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
}

if (-not $IsArchiveZip) {
    $oneDrivePath = Get-OneDriveBusinessPath -ProfileRoot $SourceRoot
    if ($oneDrivePath) {
        Write-Host "  [*] OneDrive for Business detected: $oneDrivePath" -ForegroundColor $ColorSchema.Info
        if (Test-KnownFolderMove -ProfileRoot $SourceRoot) {
            Write-Host "  [!!] Known Folder Move is active — Desktop/Documents are already inside OneDrive." -ForegroundColor $ColorSchema.Warning
            Write-Host "       Selecting both [1]/[2] and [12] will duplicate those folders." -ForegroundColor $ColorSchema.Warning
        }
    } else {
        $oneDrivePath = $null
    }
}

# ── DESTINATION ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  DESTINATION" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$DestRoot = ""

$_cfg = Get-TKConfig

if ($Unattended) {
    if ([string]::IsNullOrWhiteSpace($DestPath) -and -not [string]::IsNullOrWhiteSpace($_cfg.Revenant.DefaultDestination)) {
        $DestPath = $_cfg.Revenant.DefaultDestination
        Write-Host "  [*] No -DestPath provided — using config default: $DestPath" -ForegroundColor $ColorSchema.Info
    }
    if ([string]::IsNullOrWhiteSpace($DestPath)) {
        Write-Host "  [-] -DestPath is required in unattended mode (or set Revenant.DefaultDestination in config.json)." -ForegroundColor $ColorSchema.Error
        exit 1
    }
    $DestRoot = $DestPath.TrimEnd('\')
    if (-not (Test-Path $DestRoot)) {
        try {
            $null = New-Item -ItemType Directory -Path $DestRoot -Force -ErrorAction Stop
            Write-Host "  [+] Destination created: $DestRoot" -ForegroundColor $ColorSchema.Success
        }
        catch {
            Write-Host "  [-] Destination not accessible and could not be created: $_" -ForegroundColor $ColorSchema.Error
            exit 1
        }
    }
    Write-Host "  [+] Destination: $DestRoot" -ForegroundColor $ColorSchema.Success
} else {
    Write-Host "  [1] Current user's profile  ($env:USERPROFILE)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Enter a custom path" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $destChoice = (Read-Host).Trim()

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
}

if ($SourceRoot -ieq $DestRoot) {
    Write-Host ""
    Write-Host "  [-] Source and destination cannot be the same path." -ForegroundColor $ColorSchema.Error
    exit 1
}

# ── ITEM SELECTION ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SELECT ITEMS TO MIGRATE" -ForegroundColor $ColorSchema.Header
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

# ── MIGRATE ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  MIGRATING" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$destMigration = Join-Path $DestRoot "REVENANT_Migration"

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
        11 = @{ Label = "Firefox Profiles";       Type = "Folder"; Src = (Join-Path $SourceRoot "Firefox");                         Dst = (Join-Path $destMigration "Firefox") }
        12 = @{ Label = "OneDrive for Business";   Type = "Folder"; Src = (Join-Path $SourceRoot "OneDrive");                        Dst = (Join-Path $destMigration "OneDrive") }
    }
} else {
    $sourceAppData = Join-Path $SourceRoot "AppData\Roaming"
    $sourceLocal   = Join-Path $SourceRoot "AppData\Local"

    $itemMap = [ordered]@{
        1  = @{ Label = "Desktop";                 Type = "Folder"; Src = (Join-Path $SourceRoot "Desktop");                                            Dst = (Join-Path $DestRoot "Desktop") }
        2  = @{ Label = "Documents";               Type = "Folder"; Src = (Join-Path $SourceRoot "Documents");                                          Dst = (Join-Path $DestRoot "Documents") }
        3  = @{ Label = "Downloads";               Type = "Folder"; Src = (Join-Path $SourceRoot "Downloads");                                          Dst = (Join-Path $DestRoot "Downloads") }
        4  = @{ Label = "Pictures";                Type = "Folder"; Src = (Join-Path $SourceRoot "Pictures");                                           Dst = (Join-Path $DestRoot "Pictures") }
        5  = @{ Label = "Videos";                  Type = "Folder"; Src = (Join-Path $SourceRoot "Videos");                                             Dst = (Join-Path $DestRoot "Videos") }
        6  = @{ Label = "Music";                   Type = "Folder"; Src = (Join-Path $SourceRoot "Music");                                              Dst = (Join-Path $DestRoot "Music") }
        7  = @{ Label = "Outlook Profiles";        Type = "Folder"; Src = (Join-Path $sourceAppData "Microsoft\Outlook");                               Dst = (Join-Path $destMigration "Outlook") }
        8  = @{ Label = "Email Signatures";        Type = "Folder"; Src = (Join-Path $sourceAppData "Microsoft\Signatures");                            Dst = (Join-Path $destMigration "Signatures") }
        9  = @{ Label = "Chrome Bookmarks";        Type = "File";   Src = (Join-Path $sourceLocal   "Google\Chrome\User Data\Default\Bookmarks");       Dst = (Join-Path $destMigration "Chrome") }
        10 = @{ Label = "Edge Bookmarks";          Type = "File";   Src = (Join-Path $sourceLocal   "Microsoft\Edge\User Data\Default\Bookmarks");      Dst = (Join-Path $destMigration "Edge") }
        11 = @{ Label = "Firefox Profiles";        Type = "Folder"; Src = (Join-Path $sourceAppData "Mozilla\Firefox\Profiles");                        Dst = (Join-Path $destMigration "Firefox") }
        12 = @{ Label = "OneDrive for Business";   Type = "Folder"; Src = $(if ($oneDrivePath) { $oneDrivePath } else { Join-Path $SourceRoot "OneDrive - *" }); Dst = (Join-Path $destMigration "OneDrive") }
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

$logFile = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "REVENANT_MigrationLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

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
        "WhatIf"  { 'Cyan'              }
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
$whatif  = ($MigrationLog | Where-Object { $_.Status -eq "WhatIf"  } | Measure-Object).Count

$summaryLine = "  Copied: $copied  |  Partial: $partial  |  Skipped: $skipped  |  Failed: $failed"
if ($whatif -gt 0) { $summaryLine += "  |  WhatIf: $whatif" }
Write-Host $summaryLine -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  R.E.V.E.N.A.N.T. COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
