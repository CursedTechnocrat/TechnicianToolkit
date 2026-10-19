# embalm.ps1 - E.M.B.A.L.M. — Encapsulates My Belongings As Lasting Mementos
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
    E.M.B.A.L.M. — Encapsulates My Belongings As Lasting Mementos
    Pre-Reimaging Profile Backup Tool for PowerShell 5.1+

.DESCRIPTION
    Creates a compressed ZIP backup of a selected user profile before a machine
    is reimaged or wiped. Supports local and UNC share destinations. Generates
    a manifest file listing every archived item and a timestamped CSV log.

.USAGE
    PS C:\> .\embalm.ps1                                                           # Must be run as Administrator
    PS C:\> .\embalm.ps1 -Unattended -Username "John"                              # Archive all items for user John
    PS C:\> .\embalm.ps1 -Unattended -Username "John" -Items "1,2,3" -Destination "\\server\backup"
    PS C:\> .\embalm.ps1 -WhatIf                                                   # Preview actions without staging or compressing

.NOTES
    Version : 6.2

#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [string]$Username    = "",
    [ValidatePattern('^(?i:A)$|^\s*\d{1,2}(\s*,\s*\d{1,2})*\s*$')]
    [string]$Items       = "A",
    [string]$Destination = "",
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
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

function Write-TKError {
    <#
    .SYNOPSIS
        Appends a structured error to the toolkit's local error log,
        TK_Errors_<yyyyMM>.jsonl in the log directory. Nothing is sent off the
        machine: the toolkit keeps everything in the terminal or in files the
        technician can remove.
    .EXAMPLE
        Write-TKError -ScriptName 'basilisk' -Message $_.Exception.Message -Category 'Registry'
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][string]$Message,
        [string]$Category = 'General'
    )

    $entry = [PSCustomObject]@{
        Timestamp  = (Get-Date -Format 'o')
        Script     = $ScriptName
        Category   = $Category
        Message    = $Message
        Host       = $env:COMPUTERNAME
        User       = $env:USERNAME
    }

    # Append to monthly JSONL error log in the configured log directory
    try {
        $cfg     = Get-TKConfig
        $logRoot = if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory) -and (Test-Path $cfg.LogDirectory)) {
            $cfg.LogDirectory
        } else {
            $PSScriptRoot
        }
        $logFile = Join-Path $logRoot "TK_Errors_$(Get-Date -Format 'yyyyMM').jsonl"
        $line    = $entry | ConvertTo-Json -Compress
        Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch { <# never throw from a logging helper #> }
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

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
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
# EMBALM LOG
# ─────────────────────────────────────────────────────────────────────────────

$EmbalmLog = New-Object System.Collections.ArrayList

function Add-EmbalmRecord {
    param(
        [string]$Item,
        [string]$Status,
        [string]$Detail    = "",
        [string]$Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    )
    [void]$EmbalmLog.Add([PSCustomObject]@{
        Timestamp = $Timestamp
        Item      = $Item
        Status    = $Status
        Detail    = $Detail
    })
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-EmbalmBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███████╗███╗   ███╗██████╗  █████╗ ██╗     ███╗   ███╗
  ██╔════╝████╗ ████║██╔══██╗██╔══██╗██║     ████╗ ████║
  █████╗  ██╔████╔██║██████╔╝███████║██║     ██╔████╔██║
  ██╔══╝  ██║╚██╔╝██║██╔══██╗██╔══██║██║     ██║╚██╔╝██║
  ███████╗██║ ╚═╝ ██║██████╔╝██║  ██║███████╗██║ ╚═╝ ██║
  ╚══════╝╚═╝     ╚═╝╚═════╝ ╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    E.M.B.A.L.M. — Encapsulates My Belongings As Lasting Mementos" -ForegroundColor Cyan
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
        Add-EmbalmRecord -Item $Label -Status "Skipped" -Detail "Source not found"
        return
    }

    Write-Host "    [*] Staging $Label..." -ForegroundColor $ColorSchema.Progress

    if ($WhatIf) {
        Write-Host "    [~] Would stage $Label from $SourcePath" -ForegroundColor Cyan
        Add-EmbalmRecord -Item $Label -Status "WhatIf" -Detail "Would stage from: $SourcePath"
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
            Add-EmbalmRecord -Item $Label -Status "Staged"
        } else {
            Write-Host "    [!!] $Label — staged with warnings." -ForegroundColor $ColorSchema.Warning
            Add-EmbalmRecord -Item $Label -Status "Partial" -Detail "Robocopy exit $LASTEXITCODE"
        }
    }
    catch {
        Write-Host "    [-] $Label — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-EmbalmRecord -Item $Label -Status "Failed" -Detail $_
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-EmbalmBanner }

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
Write-Host "  SELECT PROFILE TO EMBALM" -ForegroundColor $ColorSchema.Header
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
Write-Host "  SELECT ITEMS TO EMBALM" -ForegroundColor $ColorSchema.Header
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
$archiveName  = "EMBALM_$($env:COMPUTERNAME)_${profileUsername}_$timestamp"
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
    "E.M.B.A.L.M. Backup Manifest",
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
foreach ($record in $EmbalmLog) {
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
Write-Host "  COMPRESSING EMBALM" -ForegroundColor $ColorSchema.Header
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
    Add-EmbalmRecord -Item "ZIP Archive" -Status "WhatIf" -Detail "Would create: $zipPath ($stagedCount files, $stagedSizeMB MB)"
} else {

$compressionStart = Get-Date

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    [System.IO.Compression.ZipFile]::CreateFromDirectory($stagingDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

    $zipSizeMB      = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
    $compressionSec = [math]::Round(((Get-Date) - $compressionStart).TotalSeconds, 1)
    Write-Host ("  [+] Archive created: {0}  ({1} MB, {2}s)" -f $zipPath, $zipSizeMB, $compressionSec) -ForegroundColor $ColorSchema.Success
    Add-EmbalmRecord -Item "ZIP Archive" -Status "Created" -Detail "$zipSizeMB MB — $zipPath"
}
catch {
    Write-Host "  [-] Compression failed: $_" -ForegroundColor $ColorSchema.Error
    Add-EmbalmRecord -Item "ZIP Archive" -Status "Failed" -Detail $_
    Write-TKError -ScriptName 'embalm' -Message "ZIP creation failed for '$zipPath': $($_.Exception.Message)" -Category 'Compression'
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

$logFile = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "EMBALM_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

try {
    $EmbalmLog | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
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
Write-Host "  EMBALM SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

foreach ($record in $EmbalmLog) {
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
$staged  = ($EmbalmLog | Where-Object { $_.Status -eq "Staged"  } | Measure-Object).Count
$skipped = ($EmbalmLog | Where-Object { $_.Status -eq "Skipped" } | Measure-Object).Count
$failed  = ($EmbalmLog | Where-Object { $_.Status -eq "Failed"  } | Measure-Object).Count

Write-Host "  Staged: $staged  |  Skipped: $skipped  |  Failed: $failed" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  E.M.B.A.L.M. COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
