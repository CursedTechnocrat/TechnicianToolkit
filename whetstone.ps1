# whetstone.ps1 - W.H.E.T.S.T.O.N.E. — Windows Hotfixes, Enforced Through Staged Testing Of Needed Enhancements
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
    W.H.E.T.S.T.O.N.E. — Windows Hotfixes, Enforced Through Staged Testing Of Needed Enhancements
    Windows Update & Maintenance Tool for PowerShell 5.1+

.DESCRIPTION
    Automates Windows Update management with minimal user intervention.
    Handles power settings, module installation, update deployment, and
    reboot detection. Disables sleep for the duration and restores settings
    on exit. PSWindowsUpdate module is auto-installed if missing.

.USAGE
    PS C:\> .\whetstone.ps1                          # Must be run as Administrator
    PS C:\> .\whetstone.ps1 -Unattended              # Silent mode — skip prompts and countdown
    PS C:\> .\whetstone.ps1 -Unattended -AutoReboot  # Silent mode — reboot automatically if needed

.NOTES
    Version : 6.2

#>

param(
    [switch]$Unattended,
    [switch]$AutoReboot,
    [switch]$Transcript,
    [switch]$WhatIf
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
# BANNER DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-WhetstoneBanner {
    Write-Host @"

  ██╗    ██╗██╗  ██╗███████╗████████╗███████╗████████╗ ██████╗ ███╗   ██╗███████╗
  ██║    ██║██║  ██║██╔════╝╚══██╔══╝██╔════╝╚══██╔══╝██╔═══██╗████╗  ██║██╔════╝
  ██║ █╗ ██║███████║█████╗     ██║   ███████╗   ██║   ██║   ██║██╔██╗ ██║█████╗
  ██║███╗██║██╔══██║██╔══╝     ██║   ╚════██║   ██║   ██║   ██║██║╚██╗██║██╔══╝
  ╚███╔███╔╝██║  ██║███████╗   ██║   ███████║   ██║   ╚██████╔╝██║ ╚████║███████╗
   ╚══╝╚══╝ ╚═╝  ╚═╝╚══════╝   ╚═╝   ╚══════╝   ╚═╝    ╚═════╝ ╚═╝  ╚═══╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    W.H.E.T.S.T.O.N.E. — Windows Hotfixes, Enforced Through Staged Testing Of Needed Enhancements" -ForegroundColor Cyan
    Write-Host "    Windows Update Deployment & Maintenance Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA DEFINITION
# ─────────────────────────────────────────────────────────────────────────────

$ColorSchema = @{
    Header       = 'Cyan'      # Section headers
    Success      = 'Green'     # Successful operations
    Warning      = 'Yellow'    # Warnings and cautions
    Error        = 'Red'       # Critical errors
    Info         = 'Gray'      # Information and details
    Progress     = 'Magenta'   # Progress indicators
    Accent       = 'Blue'      # Accent and highlights
}

# ─────────────────────────────────────────────────────────────────────────────
# TRANSCRIPT LOGGING
# ─────────────────────────────────────────────────────────────────────────────

$transcriptPath = $null
try {
    # Every run leaves a session log, so an unattended run can be read back.
    # It goes where the reports go -- the configured LogDirectory, else beside
    # the script -- never %TEMP%, where a technician clearing up would miss it.
    $transcriptRoot = Resolve-LogDirectory -FallbackPath $PSScriptRoot
    $transcriptPath = Join-Path $transcriptRoot "WHETSTONE_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Start-Transcript -Path $transcriptPath | Out-Null
}
catch {
    $transcriptPath = $null
}

# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY BANNER
# ─────────────────────────────────────────────────────────────────────────────

Show-WhetstoneBanner

Write-Host ""
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host "     WINDOWS UPDATE MANAGER" -ForegroundColor $ColorSchema.Header
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
if ($transcriptPath) {
    Write-Host "  Session log: $transcriptPath" -ForegroundColor $ColorSchema.Info
}
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: CONFIGURE POWER SETTINGS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[1/4] Configuring Power Settings..." -ForegroundColor $ColorSchema.Progress

# Initialize fallback values in case the query below fails
$script:originalMonitorAC = 10
$script:originalMonitorDC = 5

try {
    # Save current monitor timeout values before modifying
    $monitorQuery = powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 2>&1
    $acLine = $monitorQuery | Where-Object { $_ -match "Current AC Power Setting Index" }
    $dcLine = $monitorQuery | Where-Object { $_ -match "Current DC Power Setting Index" }
    if ($acLine) {
        $acHex = ($acLine -replace ".*:\s*0x", "").Trim()
        $script:originalMonitorAC = [math]::Round([convert]::ToInt32($acHex, 16) / 60)
    }
    if ($dcLine) {
        $dcHex = ($dcLine -replace ".*:\s*0x", "").Trim()
        $script:originalMonitorDC = [math]::Round([convert]::ToInt32($dcHex, 16) / 60)
    }
    Write-Host "    Monitor timeout saved (AC: $($script:originalMonitorAC)m, DC: $($script:originalMonitorDC)m)" -ForegroundColor $ColorSchema.Info

    powercfg /change standby-timeout-ac 0
    powercfg /change standby-timeout-dc 0
    powercfg /change monitor-timeout-ac 0
    powercfg /change monitor-timeout-dc 0
    Write-Host "[+] Power settings configured successfully" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error configuring power settings: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: INSTALL PSWINDOWSUPDATE MODULE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[2/4] Installing PSWindowsUpdate Module..." -ForegroundColor $ColorSchema.Progress
try {
    # Ensure NuGet package provider is present (required by Install-Module)
    Write-Host "    Checking NuGet package provider..." -ForegroundColor $ColorSchema.Info
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if ($null -eq $nuget -or $nuget.Version -lt [Version]"2.8.5.201") {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null
        Write-Host "    [+] NuGet provider installed" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    [+] NuGet provider is available" -ForegroundColor $ColorSchema.Success
    }

    $module = Get-Module -Name PSWindowsUpdate -ListAvailable
    if ($null -eq $module) {
        Write-Host "    Installing module (this may take a moment)..." -ForegroundColor $ColorSchema.Info
        Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
        Write-Host "[+] PSWindowsUpdate installed successfully" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    Checking for module updates..." -ForegroundColor $ColorSchema.Info
        Update-Module -Name PSWindowsUpdate -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "[+] PSWindowsUpdate is up to date" -ForegroundColor $ColorSchema.Success
    }
}
catch {
    Write-Host "[-] Error installing PSWindowsUpdate: $_" -ForegroundColor $ColorSchema.Error
    exit 1
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: IMPORT PSWINDOWSUPDATE MODULE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[3/4] Importing PSWindowsUpdate Module..." -ForegroundColor $ColorSchema.Progress
try {
    Import-Module -Name PSWindowsUpdate -Force
    Write-Host "[+] PSWindowsUpdate imported successfully" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error importing PSWindowsUpdate: $_" -ForegroundColor $ColorSchema.Error
    exit 1
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: INSTALL WINDOWS UPDATES
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[4/4] Installing Windows Updates..." -ForegroundColor $ColorSchema.Progress
Write-Host "    This may take several minutes..." -ForegroundColor $ColorSchema.Info
Write-Host ""

try {
    # Get updates without installing first to show what will be installed
    $updates = Get-WindowsUpdate -NotCategory "Drivers"

    if ($null -eq $updates -or $updates.Count -eq 0) {
        Write-Host "[+] No updates available. Your system is up to date!" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    Found $($updates.Count) update(s) to install:" -ForegroundColor $ColorSchema.Info
        $updates | ForEach-Object { Write-Host "    * $($_.Title)" -ForegroundColor $ColorSchema.Info }
        Write-Host ""

        if ($WhatIf) {
            Write-Host "[~] WhatIf: the $($updates.Count) update(s) above would be installed. No changes made." -ForegroundColor Cyan
        }
        else {
            # Install updates without reboot
            Install-WindowsUpdate -NotCategory "Drivers" -AutoReboot:$false -Confirm:$false

            Write-Host "[+] Windows Updates installed successfully" -ForegroundColor $ColorSchema.Success
        }
    }
}
catch {
    Write-Host "[-] Error installing updates: $_" -ForegroundColor $ColorSchema.Error
    Write-TKError -ScriptName 'whetstone' -Message "Windows Update install failed: $($_.Exception.Message)" -Category 'Windows Update'
}

Write-Host ""
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host "  UPDATE INSTALLATION COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# CHECK REBOOT STATUS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "Checking reboot status..." -ForegroundColor $ColorSchema.Progress
$rebootRequired = $false
try {
    # Get-WURebootStatus -Silent returns a plain [bool] ($true when a reboot is pending).
    $rebootRequired = [bool](Get-WURebootStatus -Silent -ErrorAction Stop)
}
catch {
    Write-Host "[-] Could not determine reboot status: $_" -ForegroundColor $ColorSchema.Warning
    Write-Host "    Proceeding without reboot check." -ForegroundColor $ColorSchema.Warning
}

Write-Host ""

if ($rebootRequired) {
    $lastBoot = try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime } catch { 'unknown' }
    Write-Host "  *** REBOOT REQUIRED ***" -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host "  Reboot Status Details:" -ForegroundColor $ColorSchema.Warning
    Write-Host "  | Reboot Required: $rebootRequired" -ForegroundColor $ColorSchema.Warning
    Write-Host "  | Last Boot Time: $lastBoot" -ForegroundColor $ColorSchema.Info
    Write-Host ""
}
else {
    Write-Host "[+] No reboot required at this time" -ForegroundColor $ColorSchema.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REBOOT DECISION - ONLY PROMPT IF REBOOT IS REQUIRED
# ─────────────────────────────────────────────────────────────────────────────

if ($WhatIf) {
    Write-Host "[~] WhatIf: reboot decision skipped — no updates were installed." -ForegroundColor Cyan
    Write-Host ""
}
elseif ($rebootRequired) {
    Write-Host "  *** REBOOT REQUIRED ***" -ForegroundColor $ColorSchema.Warning
    Write-Host ""

    if ($Unattended) {
        if ($AutoReboot) {
            Write-Host "  [*] Unattended mode: AutoReboot enabled — rebooting in 10 seconds..." -ForegroundColor $ColorSchema.Warning
            Write-Host ""
            Start-Sleep -Seconds 10
            Write-Host "Restoring monitor timeout before reboot..." -ForegroundColor $ColorSchema.Info
            powercfg /change monitor-timeout-ac $script:originalMonitorAC
            powercfg /change monitor-timeout-dc $script:originalMonitorDC
            if ($transcriptPath) { try { Stop-Transcript } catch {} }
            Restart-Computer -Force
        } else {
            Write-Host "  [*] Unattended mode: reboot required but -AutoReboot not set. Skipping." -ForegroundColor $ColorSchema.Warning
            Write-Host "  [!!] Reboot this machine when ready: Restart-Computer" -ForegroundColor $ColorSchema.Warning
            Write-Host ""
        }
    } else {
        Write-Host "The computer is ready to be rebooted." -ForegroundColor $ColorSchema.Warning
        Write-Host ""

        $rebootPrompt = Read-Host "Is it safe to reboot this computer now? (Y/N)"

        if ($rebootPrompt -eq 'Y' -or $rebootPrompt -eq 'y') {
            # [Console]::KeyAvailable throws when no console is attached — the GUI host, a
            # WinRM runspace, a scheduled task. Probe once and degrade to a plain wait.
            $canPollKeys = $false
            try { $null = [Console]::KeyAvailable; $canPollKeys = $true } catch { }

            Write-Host ""
            if ($canPollKeys) {
                Write-Host "Initiating reboot in 30 seconds. Press Escape to cancel..." -ForegroundColor $ColorSchema.Warning
            } else {
                Write-Host "Initiating reboot in 30 seconds..." -ForegroundColor $ColorSchema.Warning
            }
            Write-Host ""
            Write-Host "   30 [============================================]" -ForegroundColor $ColorSchema.Accent

            # 30-second countdown with Escape key detection
            $cancelled = $false
            for ($i = 30; $i -gt 0; $i--) {
                $progress = [math]::Floor((30 - $i) / 30 * 44)
                $bar = "=" * $progress
                $remaining = " " * (44 - $progress)
                Write-Host -NoNewline "`r   $i  [$bar$remaining]" -ForegroundColor $ColorSchema.Accent

                # Poll for Escape key in 100ms intervals
                for ($tick = 0; $tick -lt 10; $tick++) {
                    if ($canPollKeys -and [Console]::KeyAvailable) {
                        $key = [Console]::ReadKey($true)
                        if ($key.Key -eq [ConsoleKey]::Escape) {
                            $cancelled = $true
                            break
                        }
                    }
                    Start-Sleep -Milliseconds 100
                }
                if ($cancelled) { break }
            }

            Write-Host ""
            Write-Host ""

            if ($cancelled) {
                Write-Host "  Reboot cancelled." -ForegroundColor $ColorSchema.Warning
                Write-Host ""
                Write-Host "  !!! REBOOT SKIPPED !!!" -ForegroundColor $ColorSchema.Error
                Write-Host ""
                Write-Host "  IMPORTANT: You must reboot your computer to complete" -ForegroundColor $ColorSchema.Error
                Write-Host "  the updates!" -ForegroundColor $ColorSchema.Error
                Write-Host ""
                Write-Host "  When you are ready to reboot, use one of these methods:" -ForegroundColor $ColorSchema.Warning
                Write-Host "  | Command: Restart-Computer" -ForegroundColor $ColorSchema.Info
                Write-Host "  | Or manually restart through Settings > System > Power" -ForegroundColor $ColorSchema.Info
                Write-Host ""
            }
            else {
                Write-Host "Restoring monitor timeout before reboot..." -ForegroundColor $ColorSchema.Info
                powercfg /change monitor-timeout-ac $script:originalMonitorAC
                powercfg /change monitor-timeout-dc $script:originalMonitorDC
                Write-Host "Rebooting now..." -ForegroundColor $ColorSchema.Warning
                Write-Host ""
                if ($transcriptPath) { try { Stop-Transcript } catch {} }
                Restart-Computer -Force
            }
        }
        else {
            Write-Host ""
            Write-Host "  !!! REBOOT SKIPPED !!!" -ForegroundColor $ColorSchema.Error
            Write-Host ""
            Write-Host "  IMPORTANT: You must reboot your computer to complete" -ForegroundColor $ColorSchema.Error
            Write-Host "  the updates!" -ForegroundColor $ColorSchema.Error
            Write-Host ""
            Write-Host "  When you are ready to reboot, use one of these methods:" -ForegroundColor $ColorSchema.Warning
            Write-Host "  | Command: Restart-Computer" -ForegroundColor $ColorSchema.Info
            Write-Host "  | Or manually restart through Settings > System > Power" -ForegroundColor $ColorSchema.Info
            Write-Host ""
        }
    }
}
# Note: the "no reboot required" case is already reported by the status block above,
# so the reboot-decision block intentionally has no else branch here.

# ─────────────────────────────────────────────────────────────────────────────
# RESTORE MONITOR TIMEOUT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "Restoring monitor timeout settings..." -ForegroundColor $ColorSchema.Progress
try {
    powercfg /change monitor-timeout-ac $script:originalMonitorAC
    powercfg /change monitor-timeout-dc $script:originalMonitorDC
    Write-Host "[+] Monitor timeout restored (AC: $($script:originalMonitorAC)m, DC: $($script:originalMonitorDC)m)" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Could not restore monitor timeout: $_" -ForegroundColor $ColorSchema.Warning
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# COMPLETION MESSAGE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host "  SCRIPT EXECUTION COMPLETED" -ForegroundColor $ColorSchema.Header
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
if ($transcriptPath) {
    Write-Host ""
    Write-Host "  Session log saved to:" -ForegroundColor $ColorSchema.Info
    Write-Host "  $transcriptPath" -ForegroundColor $ColorSchema.Info
}
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STOP TRANSCRIPT
# ─────────────────────────────────────────────────────────────────────────────

if ($transcriptPath) {
    try { Stop-Transcript } catch {}
}
