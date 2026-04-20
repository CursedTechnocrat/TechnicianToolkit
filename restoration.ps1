<#
.SYNOPSIS
    R.E.S.T.O.R.A.T.I.O.N. — Renews Every System Through Orderly Rite — Automating The Installation Of New updates
    Windows Update & Maintenance Tool for PowerShell 5.1+

.DESCRIPTION
    Automates Windows Update management with minimal user intervention.
    Handles power settings, module installation, update deployment, and
    reboot detection. Disables sleep for the duration and restores settings
    on exit. PSWindowsUpdate module is auto-installed if missing.

.USAGE
    PS C:\> .\restoration.ps1                          # Must be run as Administrator
    PS C:\> .\restoration.ps1 -Unattended              # Silent mode — skip prompts and countdown
    PS C:\> .\restoration.ps1 -Unattended -AutoReboot  # Silent mode — reboot automatically if needed

.NOTES
    Version : 1.1

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    R.E.L.I.C.             — Certificate health & SSL expiry monitoring
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
    [switch]$AutoReboot
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Assert-AdminPrivilege

# ─────────────────────────────────────────────────────────────────────────────
# BANNER DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-RestorationBanner {
    Write-Host @"

  ██████╗ ███████╗███████╗████████╗ ██████╗ ██████╗  █████╗ ████████╗██╗ ██████╗ ███╗   ██╗
  ██╔══██╗██╔════╝██╔════╝╚══██╔══╝██╔═══██╗██╔══██╗██╔══██╗╚══██╔══╝██║██╔═══██╗████╗  ██║
  ██████╔╝█████╗  ███████╗   ██║   ██║   ██║██████╔╝███████║   ██║   ██║██║   ██║██╔██╗ ██║
  ██╔══██╗██╔══╝  ╚════██╗   ██║   ██║   ██║██╔══██╗██╔══██║   ██║   ██║██║   ██║██║╚██╗██║
  ██║  ██║███████╗███████║   ██║   ╚██████╔╝██║  ██║██║  ██║   ██║   ██║╚██████╔╝██║ ╚████║
  ╚═╝  ╚═╝╚══════╝╚══════╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    R.E.S.T.O.R.A.T.I.O.N. — Renews Every System Through Orderly Rite" -ForegroundColor Cyan
    Write-Host "    Automating The Installation Of New updates" -ForegroundColor Cyan
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
    $transcriptPath = "$env:TEMP\RESTORATION_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Start-Transcript -Path $transcriptPath | Out-Null
}
catch {
    $transcriptPath = $null
}

# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY BANNER
# ─────────────────────────────────────────────────────────────────────────────

Show-RestorationBanner

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

        # Install updates without reboot
        Install-WindowsUpdate -NotCategory "Drivers" -AutoReboot:$false -Confirm:$false

        Write-Host "[+] Windows Updates installed successfully" -ForegroundColor $ColorSchema.Success
    }
}
catch {
    Write-Host "[-] Error installing updates: $_" -ForegroundColor $ColorSchema.Error
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
    $rebootStatus = Get-WindowsUpdateRebootStatus
    $rebootRequired = $rebootStatus.RebootRequired
}
catch {
    Write-Host "[-] Could not determine reboot status: $_" -ForegroundColor $ColorSchema.Warning
    Write-Host "    Proceeding without reboot check." -ForegroundColor $ColorSchema.Warning
}

Write-Host ""

if ($rebootRequired) {
    Write-Host "  *** REBOOT REQUIRED ***" -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host "  Reboot Status Details:" -ForegroundColor $ColorSchema.Warning
    Write-Host "  | Reboot Required: $($rebootStatus.RebootRequired)" -ForegroundColor $ColorSchema.Warning
    Write-Host "  | Last Boot Time: $($rebootStatus.LastBootUpTime)" -ForegroundColor $ColorSchema.Info
    Write-Host ""
}
else {
    Write-Host "[+] No reboot required at this time" -ForegroundColor $ColorSchema.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REBOOT DECISION - ONLY PROMPT IF REBOOT IS REQUIRED
# ─────────────────────────────────────────────────────────────────────────────

if ($rebootRequired) {
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
            Write-Host ""
            Write-Host "Initiating reboot in 30 seconds. Press Escape to cancel..." -ForegroundColor $ColorSchema.Warning
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
                    if ([Console]::KeyAvailable) {
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
else {
    Write-Host "[+] No reboot required at this time" -ForegroundColor $ColorSchema.Success
    Write-Host ""
}

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
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
