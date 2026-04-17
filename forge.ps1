<#
.SYNOPSIS
    F.O.R.G.E. — Finds Outdated Resources & Generates Equipment-updates
    Driver Detection & Installation Tool for PowerShell 5.1+

.DESCRIPTION
    Audits the device tree for problem devices and outdated drivers.
    Scans Windows Update for available driver updates, installs drivers
    from a local folder (ZIP/INF/EXE/MSI), and generates a CSV report
    of all device driver versions.

.USAGE
    PS C:\> .\forge.ps1                                             # Interactive menu
    PS C:\> .\forge.ps1 -Unattended -Action Audit                  # List problem devices
    PS C:\> .\forge.ps1 -Unattended -Action WindowsUpdate          # Install WU driver updates
    PS C:\> .\forge.ps1 -Unattended -Action LocalInstall           # Install drivers from script folder

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
    S.I.G.I.L.             — Security baseline & policy enforcement
    S.P.E.C.T.E.R.         — Remote machine execution via WinRM
    L.E.Y.L.I.N.E.         — Network diagnostics & remediation
    F.O.R.G.E.             — Driver update detection & installation

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors / problem devices
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit','WindowsUpdate','LocalInstall','Report')]
    [string]$Action = 'Audit'
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptPath = (Get-Location).Path

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA
# ─────────────────────────────────────────────────────────────────────────────

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-ForgeBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███████╗ ██████╗ ██████╗  ██████╗ ███████╗
  ██╔════╝██╔═══██╗██╔══██╗██╔════╝ ██╔════╝
  █████╗  ██║   ██║██████╔╝██║  ███╗█████╗
  ██╔══╝  ██║   ██║██╔══██╗██║   ██║██╔══╝
  ██║     ╚██████╔╝██║  ██║╚██████╔╝███████╗
  ╚═╝      ╚═════╝ ╚═╝  ╚═╝ ╚═════╝ ╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    F.O.R.G.E. — Finds Outdated Resources & Generates Equipment-updates" -ForegroundColor Cyan
    Write-Host "    Driver Detection & Installation Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER
# ─────────────────────────────────────────────────────────────────────────────

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host "  $Title" -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# DEVICE AUDIT
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-DeviceAudit {
    Write-Section "DEVICE DRIVER AUDIT"

    Write-Host "  [*] Scanning device tree..." -ForegroundColor $C.Progress

    $allDevices = Get-WmiObject Win32_PnPEntity -ErrorAction SilentlyContinue |
                  Where-Object { $_.DeviceID -notmatch '^HTREE\\' } |
                  Sort-Object ConfigManagerErrorCode, Name

    $problemDevices = $allDevices | Where-Object { $_.ConfigManagerErrorCode -ne 0 }
    $okDevices      = $allDevices | Where-Object { $_.ConfigManagerErrorCode -eq 0 }

    Write-Host ("  [+] {0} total devices  |  {1} OK  |  {2} with errors" -f $allDevices.Count, $okDevices.Count, $problemDevices.Count) -ForegroundColor $(if ($problemDevices.Count -gt 0) { $C.Warning } else { $C.Success })
    Write-Host ""

    if ($problemDevices.Count -eq 0) {
        Write-Host "  [+] All devices functioning normally." -ForegroundColor $C.Success
    } else {
        Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Error
        Write-Host "  PROBLEM DEVICES" -ForegroundColor $C.Error
        Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Error
        Write-Host ""

        foreach ($dev in $problemDevices) {
            $errorDesc = switch ($dev.ConfigManagerErrorCode) {
                1  { "Device not configured correctly" }
                3  { "Driver missing or corrupted" }
                10 { "Device cannot start" }
                12 { "Insufficient resources" }
                18 { "Reinstall drivers required" }
                28 { "Driver not installed" }
                43 { "Device reported a problem" }
                45 { "Device not present" }
                default { "Error code $($dev.ConfigManagerErrorCode)" }
            }
            Write-Host ("  [!!] {0}" -f $dev.Name) -ForegroundColor $C.Error
            Write-Host ("       {0}" -f $errorDesc) -ForegroundColor $C.Warning
            Write-Host ("       DeviceID: {0}" -f ($dev.DeviceID -replace '\\', '\\')) -ForegroundColor $C.Info
            Write-Host ""
        }
    }

    return $problemDevices
}

# ─────────────────────────────────────────────────────────────────────────────
# WINDOWS UPDATE DRIVER SCAN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-WindowsUpdateDrivers {
    Write-Section "WINDOWS UPDATE — DRIVER UPDATES"

    Write-Host "  [*] Checking for driver updates via Windows Update..." -ForegroundColor $C.Progress
    Write-Host "      This may take a minute." -ForegroundColor $C.Info
    Write-Host ""

    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
            Write-Host "  [*] Installing PSWindowsUpdate module..." -ForegroundColor $C.Progress
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop

        $driverUpdates = Get-WUList -Category "Drivers" -ErrorAction Stop

        if ($driverUpdates.Count -eq 0) {
            Write-Host "  [+] No driver updates available via Windows Update." -ForegroundColor $C.Success
            return
        }

        Write-Host ("  [+] {0} driver update(s) available:" -f $driverUpdates.Count) -ForegroundColor $C.Warning
        Write-Host ""
        foreach ($u in $driverUpdates) {
            Write-Host ("  [ ] {0}" -f $u.Title) -ForegroundColor $C.Info
            Write-Host ("      Size: {0:N1} MB" -f ($u.Size / 1MB)) -ForegroundColor $C.Info
            Write-Host ""
        }

        if (-not $Unattended) {
            Write-Host -NoNewline "  Install all driver updates now? (Y/N): " -ForegroundColor $C.Warning
            $confirm = (Read-Host).Trim().ToUpper()
            if ($confirm -ne 'Y') {
                Write-Host "  [*] Skipped installation." -ForegroundColor $C.Info
                return
            }
        }

        Write-Host ""
        Write-Host "  [*] Installing driver updates..." -ForegroundColor $C.Progress
        Install-WindowsUpdate -Category "Drivers" -AcceptAll -IgnoreReboot -ErrorAction Stop |
            ForEach-Object { Write-Host ("  [*] {0}" -f $_.Title) -ForegroundColor $C.Progress }

        Write-Host ""
        Write-Host "  [+] Driver updates installed. A reboot may be required." -ForegroundColor $C.Success
    }
    catch {
        Write-Host "  [-] Windows Update driver scan failed: $_" -ForegroundColor $C.Error
        Write-Host "  [!!] Ensure the machine has internet access and PSWindowsUpdate can be installed." -ForegroundColor $C.Warning
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# LOCAL DRIVER INSTALL (ZIP / INF / EXE / MSI from script folder)
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-LocalDriverInstall {
    Write-Section "LOCAL DRIVER INSTALL"

    Write-Host "  Looks for driver packages in the current folder:" -ForegroundColor $C.Info
    Write-Host "  $ScriptPath" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host "  Supported formats: .zip (INF inside)  |  .inf  |  .exe  |  .msi" -ForegroundColor $C.Info
    Write-Host ""

    $driverFiles = Get-ChildItem -Path $ScriptPath -Include '*.zip','*.inf','*.exe','*.msi' -Recurse -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -ne (Split-Path $PSCommandPath -Leaf) }

    if (-not $driverFiles) {
        Write-Host "  [!!] No driver files found in $ScriptPath." -ForegroundColor $C.Warning
        Write-Host "       Drop .zip, .inf, .exe, or .msi driver packages here and re-run." -ForegroundColor $C.Info
        return
    }

    Write-Host ("  [+] Found {0} file(s):" -f $driverFiles.Count) -ForegroundColor $C.Success
    $driverFiles | ForEach-Object { Write-Host ("    - {0}" -f $_.Name) -ForegroundColor $C.Info }
    Write-Host ""

    if (-not $Unattended) {
        Write-Host -NoNewline "  Proceed with installation? (Y/N): " -ForegroundColor $C.Warning
        $confirm = (Read-Host).Trim().ToUpper()
        if ($confirm -ne 'Y') {
            Write-Host "  [*] Cancelled." -ForegroundColor $C.Info
            return
        }
    }

    $extractRoot = Join-Path $ScriptPath "FORGE_Extracted"

    foreach ($file in $driverFiles) {
        Write-Host ""
        Write-Host ("  [*] Processing: {0}" -f $file.Name) -ForegroundColor $C.Progress

        switch ($file.Extension.ToLower()) {
            '.zip' {
                $extractDir = Join-Path $extractRoot ($file.BaseName)
                try {
                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                    [IO.Compression.ZipFile]::ExtractToDirectory($file.FullName, $extractDir)
                    $infFiles = Get-ChildItem -Path $extractDir -Filter '*.inf' -Recurse
                    if ($infFiles) {
                        foreach ($inf in $infFiles) {
                            Write-Host ("    [*] Installing INF: {0}" -f $inf.Name) -ForegroundColor $C.Progress
                            $result = pnputil /add-driver $inf.FullName /install 2>&1
                            if ($LASTEXITCODE -eq 0) {
                                Write-Host "    [+] Installed successfully." -ForegroundColor $C.Success
                            } else {
                                Write-Host ("    [-] pnputil returned exit code {0}" -f $LASTEXITCODE) -ForegroundColor $C.Warning
                            }
                        }
                    } else {
                        Write-Host "    [!!] No .inf files found inside ZIP." -ForegroundColor $C.Warning
                    }
                } catch {
                    Write-Host ("    [-] Failed to extract/install: {0}" -f $_) -ForegroundColor $C.Error
                }
            }
            '.inf' {
                try {
                    $result = pnputil /add-driver $file.FullName /install 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "    [+] INF installed via pnputil." -ForegroundColor $C.Success
                    } else {
                        Write-Host ("    [-] pnputil exit code {0}" -f $LASTEXITCODE) -ForegroundColor $C.Warning
                    }
                } catch {
                    Write-Host ("    [-] {0}" -f $_) -ForegroundColor $C.Error
                }
            }
            '.exe' {
                try {
                    $proc = Start-Process -FilePath $file.FullName -ArgumentList '/s','/silent','/quiet' -Wait -PassThru
                    if ($proc.ExitCode -eq 0) {
                        Write-Host "    [+] EXE installer completed (exit 0)." -ForegroundColor $C.Success
                    } else {
                        Write-Host ("    [!!] EXE exit code: {0} — verify manually." -f $proc.ExitCode) -ForegroundColor $C.Warning
                    }
                } catch {
                    Write-Host ("    [-] {0}" -f $_) -ForegroundColor $C.Error
                }
            }
            '.msi' {
                try {
                    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$($file.FullName)`" /quiet /norestart" -Wait -PassThru
                    if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
                        Write-Host "    [+] MSI installed$(if ($proc.ExitCode -eq 3010) { ' — reboot required' })." -ForegroundColor $C.Success
                    } else {
                        Write-Host ("    [!!] MSI exit code: {0} — verify manually." -f $proc.ExitCode) -ForegroundColor $C.Warning
                    }
                } catch {
                    Write-Host ("    [-] {0}" -f $_) -ForegroundColor $C.Error
                }
            }
        }
    }

    if (Test-Path $extractRoot) {
        Remove-Item -Path $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "  [+] Local driver installation complete." -ForegroundColor $C.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# DRIVER REPORT (CSV)
# ─────────────────────────────────────────────────────────────────────────────

function Export-DriverReport {
    Write-Section "EXPORT DRIVER REPORT"

    Write-Host "  [*] Collecting driver information..." -ForegroundColor $C.Progress

    $logPath = Join-Path $ScriptPath ("FORGE_DriverReport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    try {
        $drivers = Get-WmiObject Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
                   Where-Object { $_.DeviceName } |
                   Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer,
                                 @{N='Status'; E={ (Get-WmiObject Win32_PnPEntity | Where-Object { $_.Name -eq $_.DeviceName } | Select-Object -First 1).Status }} |
                   Sort-Object DeviceName

        $drivers | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
        Write-Host ("  [+] Report saved: {0}" -f $logPath) -ForegroundColor $C.Success
        Write-Host ("      {0} drivers listed." -f $drivers.Count) -ForegroundColor $C.Info
    } catch {
        Write-Host ("  [-] Failed to generate report: {0}" -f $_) -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-ForgeBanner
    switch ($Action) {
        'Audit'         { Invoke-DeviceAudit }
        'WindowsUpdate' { Invoke-WindowsUpdateDrivers }
        'LocalInstall'  { Invoke-LocalDriverInstall }
        'Report'        { Export-DriverReport }
    }
} else {
    $choice = ''

    do {
        Show-ForgeBanner

        Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit device tree  (problem devices)" -ForegroundColor $C.Info
        Write-Host "  [2] Check & install driver updates via Windows Update" -ForegroundColor $C.Info
        Write-Host "  [3] Install drivers from current folder  (ZIP / INF / EXE / MSI)" -ForegroundColor $C.Info
        Write-Host "  [4] Export full driver report  (CSV)" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-DeviceAudit }
            '2' { Invoke-WindowsUpdateDrivers }
            '3' { Invoke-LocalDriverInstall }
            '4' { Export-DriverReport }
            'Q' {
                Write-Host ""
                Write-Host "  Closing F.O.R.G.E." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-4 or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @('Q')) {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
