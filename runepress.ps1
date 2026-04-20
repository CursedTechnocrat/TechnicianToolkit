<#
.SYNOPSIS
    R.U.N.E.P.R.E.S.S. — Remote Utility for Networked Equipment — Printer Registration, Extraction & Silent Setup
    Printer Driver Installation & Configuration Tool for PowerShell 5.1+

.DESCRIPTION
    Automates printer driver extraction, installation, and network printer
    configuration via a command-line interface. Supports ZIP, EXE, and MSI
    driver formats, INF-based installation via pnputil, and TCP/IP or UNC
    port configuration. Generates a timestamped CSV installation log.

.USAGE
    PS C:\> .\runepress.ps1                    # Must be run as Administrator
    PS C:\> .\runepress.ps1 -Unattended        # Silent mode — auto-selects first INF, skips printer config

.NOTES
    Version : 3.0

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
    [switch]$Transcript
)

# ===========================
# ADMIN PRIVILEGE CHECK
# ===========================
Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Invoke-AdminElevation -ScriptFile $PSCommandPath

# ===========================
# SCRIPT INITIALIZATION
# ===========================

# Resolve script execution path
if ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
}
elseif ($MyInvocation.MyCommand.Path) {
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $ScriptPath = (Get-Location).Path
}

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# Initialize global variables
$ExtractRoot     = Join-Path $ScriptPath "ExtractedDrivers"
$InstallationLog = @()

# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-Banner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██╗   ██╗███╗   ██╗███████╗██████╗ ██████╗ ███████╗███████╗
  ██╔══██╗██║   ██║████╗  ██║██╔════╝██╔══██╗██╔══██╗██╔════╝██╔════╝
  ██████╔╝██║   ██║██╔██╗ ██║█████╗  ██████╔╝██████╔╝█████╗  ███████╗
  ██╔══██╗██║   ██║██║╚██╗██║██╔══╝  ██╔═══╝ ██╔══██╗██╔══╝  ╚════██╗
  ██║  ██║╚██████╔╝██║ ╚████║███████╗██║     ██║  ██║███████╗███████║
  ╚═╝  ╚═╝ ╚═════╝ ╚═╝  ╚═══╝╚══════╝╚═╝     ╚═╝  ╚═╝╚══════╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    R.U.N.E.P.R.E.S.S. - Remote Utility for Networked Equipment" -ForegroundColor Cyan
    Write-Host "    Printer Registration, Extraction and Silent Setup" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Script Location: $ScriptPath" -ForegroundColor Gray
    Write-Host "    Execution Time:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host ""
}

# ===========================
# DISPLAY DRIVER PREP INSTRUCTIONS
# ===========================

function Show-DriverPrepInstructions {
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Step 1: Driver Preparation" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor Yellow
    Write-Host "  1. Download the printer driver from the manufacturer website" -ForegroundColor White
    Write-Host "  2. Save the file to this location:" -ForegroundColor White
    Write-Host "     $ScriptPath" -ForegroundColor Green
    Write-Host ""
    Write-Host "Supported formats:" -ForegroundColor Yellow
    Write-Host "  * ZIP archives (.zip)" -ForegroundColor White
    Write-Host "  * Executable installers (.exe)" -ForegroundColor White
    Write-Host "  * Windows Installer packages (.msi)" -ForegroundColor White
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
}

# ===========================
# PROMPT USER TO PLACE DRIVER
# ===========================

function Wait-ForDriverFile {
    while ($true) {
        Write-Host "Ready to proceed? (Y/N/Q)" -ForegroundColor Yellow
        Write-Host "  Y = Continue with installation" -ForegroundColor Gray
        Write-Host "  N = Go back and check folder" -ForegroundColor Gray
        Write-Host "  Q = Quit" -ForegroundColor Gray
        Write-Host ""

        $response = Read-Host "Enter choice"

        switch ($response.ToUpper()) {
            "Y" {
                return $true
            }
            "N" {
                Write-Host ""
                Show-DriverPrepInstructions
            }
            "Q" {
                Write-Host ""
                Write-Host "WARNING: Script terminated by user." -ForegroundColor Yellow
                exit 0
            }
            default {
                Write-Host "Invalid input. Please enter Y, N, or Q." -ForegroundColor Red
                Write-Host ""
            }
        }
    }
}

# ===========================
# LOCATE DRIVER FILES
# ===========================

function Find-DriverFiles {
    $DriverFiles = Get-ChildItem -Path $ScriptPath -File |
                   Where-Object { $_.Extension -match '\.(zip|exe|msi)$' } |
                   Where-Object { $_.Name -ne (Split-Path -Leaf $PSCommandPath) }

    return $DriverFiles
}

# ===========================
# INSTALL ZIP DRIVERS
# ===========================

function Install-ZipDriver {
    param(
        [System.IO.FileInfo]$ZipFile
    )

    $DriverName  = [System.IO.Path]::GetFileNameWithoutExtension($ZipFile.Name)
    $ExtractPath = Join-Path $ExtractRoot $DriverName

    Write-Host ""
    Write-Host "Processing ZIP: $($ZipFile.Name)" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

    # Clean previous extraction
    if (Test-Path $ExtractPath) {
        Write-Host "Removing previous extraction directory..." -ForegroundColor Yellow
        try {
            Remove-Item $ExtractPath -Recurse -Force -ErrorAction Stop
            Write-Host "OK: Previous extraction removed." -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Could not remove directory: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    # Extract ZIP archive
    Write-Host "Extracting: $($ZipFile.Name)..." -ForegroundColor Yellow
    try {
        Expand-Archive -Path $ZipFile.FullName -DestinationPath $ExtractPath -Force -ErrorAction Stop
        Write-Host "OK: Extraction complete." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Locate INF files in extracted content
    Write-Host "Searching for INF driver files..." -ForegroundColor Yellow
    $InfFiles = Get-ChildItem -Path $ExtractPath -Filter "*.inf" -Recurse

    if (-not $InfFiles) {
        Write-Host "ERROR: No INF files found in extracted content." -ForegroundColor Red
        return $false
    }

    Write-Host "Found $($InfFiles.Count) INF file(s)." -ForegroundColor Green

    # Select INF - prompt if multiple found
    if ($InfFiles.Count -eq 1) {
        $SelectedInf = $InfFiles[0]
    }
    elseif ($Unattended) {
        Write-Host "    [*] Multiple INFs found — auto-selecting first: $($InfFiles[0].Name)" -ForegroundColor Gray
        $SelectedInf = $InfFiles[0]
    }
    else {
        Write-Host ""
        Write-Host "Multiple INF files found. Select one to install:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $InfFiles.Count; $i++) {
            Write-Host "  [$($i + 1)] $($InfFiles[$i].FullName)" -ForegroundColor White
        }
        Write-Host ""
        do {
            $selection = Read-Host "Enter number (1-$($InfFiles.Count))"
        } while (-not ($selection -match '^\d+$') -or [int]$selection -lt 1 -or [int]$selection -gt $InfFiles.Count)
        $SelectedInf = $InfFiles[[int]$selection - 1]
    }

    Write-Host "Using INF: $($SelectedInf.FullName)" -ForegroundColor Cyan

    # Install driver package via pnputil
    Write-Host "Installing driver package..." -ForegroundColor Yellow
    try {
        $PnpResult = & pnputil /add-driver "$($SelectedInf.FullName)" /install 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: pnputil failed (exit code $LASTEXITCODE)." -ForegroundColor Red
            Write-Host ($PnpResult | Out-String) -ForegroundColor Red
            $script:InstallationLog += [PSCustomObject]@{
                File   = $ZipFile.Name
                Type   = "ZIP"
                INF    = $SelectedInf.Name
                Status = "Failed (pnputil exit $LASTEXITCODE)"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            return $false
        }
        Write-Host "OK: Driver package installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to run pnputil: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    $script:InstallationLog += [PSCustomObject]@{
        File   = $ZipFile.Name
        Type   = "ZIP"
        INF    = $SelectedInf.Name
        Status = "Success"
        Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    return $true
}

# ===========================
# INSTALL EXE DRIVERS
# ===========================

function Install-ExeDriver {
    param(
        [System.IO.FileInfo]$ExeFile
    )

    Write-Host ""
    Write-Host "Processing EXE: $($ExeFile.Name)" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Running silent installer..." -ForegroundColor Yellow

    try {
        $Process = Start-Process -FilePath $ExeFile.FullName `
            -ArgumentList "/S /silent /quiet /norestart" `
            -Wait -PassThru -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to launch installer: $($_.Exception.Message)" -ForegroundColor Red
        $script:InstallationLog += [PSCustomObject]@{
            File   = $ExeFile.Name
            Type   = "EXE"
            INF    = "N/A"
            Status = "Failed (launch error)"
            Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $false
    }

    # EXE exit codes vary by manufacturer; non-zero may still indicate success
    # (e.g. reboot required). Treat 0 as clean success, flag anything else as a warning.
    if ($Process.ExitCode -eq 0) {
        Write-Host "OK: Installer completed successfully." -ForegroundColor Green
        $status = "Success"
    }
    else {
        Write-Host "WARNING: Installer exited with code $($Process.ExitCode)." -ForegroundColor Yellow
        Write-Host "         Review manually - this may indicate a reboot requirement or vendor-specific code." -ForegroundColor Yellow
        $status = "Warning (exit $($Process.ExitCode))"
    }

    $script:InstallationLog += [PSCustomObject]@{
        File   = $ExeFile.Name
        Type   = "EXE"
        INF    = "N/A"
        Status = $status
        Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    return $true
}

# ===========================
# INSTALL MSI DRIVERS
# ===========================

function Install-MsiDriver {
    param(
        [System.IO.FileInfo]$MsiFile
    )

    Write-Host ""
    Write-Host "Processing MSI: $($MsiFile.Name)" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Running silent installer..." -ForegroundColor Yellow

    try {
        $Process = Start-Process -FilePath "msiexec.exe" `
            -ArgumentList "/i `"$($MsiFile.FullName)`" /qn /norestart" `
            -Wait -PassThru -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to launch msiexec: $($_.Exception.Message)" -ForegroundColor Red
        $script:InstallationLog += [PSCustomObject]@{
            File   = $MsiFile.Name
            Type   = "MSI"
            INF    = "N/A"
            Status = "Failed (launch error)"
            Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $false
    }

    switch ($Process.ExitCode) {
        0 {
            Write-Host "OK: MSI installed successfully." -ForegroundColor Green
            $status = "Success"
        }
        3010 {
            Write-Host "OK: MSI installed successfully. A system reboot is required." -ForegroundColor Yellow
            $status = "Success (reboot required)"
        }
        default {
            Write-Host "ERROR: msiexec failed with exit code $($Process.ExitCode)." -ForegroundColor Red
            $script:InstallationLog += [PSCustomObject]@{
                File   = $MsiFile.Name
                Type   = "MSI"
                INF    = "N/A"
                Status = "Failed (exit $($Process.ExitCode))"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            return $false
        }
    }

    $script:InstallationLog += [PSCustomObject]@{
        File   = $MsiFile.Name
        Type   = "MSI"
        INF    = "N/A"
        Status = $status
        Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    return $true
}

# ===========================
# SELECT INSTALLED PRINTER DRIVER
# ===========================

function Select-InstalledDriver {
    $Drivers = @(Get-PrinterDriver | Select-Object -ExpandProperty Name | Sort-Object)

    if ($Drivers.Count -eq 0) {
        Write-Host "ERROR: No printer drivers found on this system." -ForegroundColor Red
        return $null
    }

    Write-Host ""
    Write-Host "Available printer drivers:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Drivers.Count; $i++) {
        Write-Host "  [$($i + 1)] $($Drivers[$i])" -ForegroundColor White
    }
    Write-Host ""

    do {
        $selection = Read-Host "Select driver (1-$($Drivers.Count))"
    } while (-not ($selection -match '^\d+$') -or [int]$selection -lt 1 -or [int]$selection -gt $Drivers.Count)

    return $Drivers[[int]$selection - 1]
}

# ===========================
# NETWORK PRINTER CONFIGURATION
# ===========================

function Add-NetworkPrinter {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Step 3: Network Printer Configuration" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    while ($true) {
        $response = Read-Host "Add a network printer? (Y/N)"
        if ($response.ToUpper() -ne "Y") {
            Write-Host "Skipping network printer configuration." -ForegroundColor Yellow
            return
        }

        # Printer display name
        do {
            $PrinterName = Read-Host "Printer display name (e.g. Office Printer 1)"
            if (-not $PrinterName) {
                Write-Host "ERROR: Name cannot be empty." -ForegroundColor Red
            }
        } while (-not $PrinterName)

        # Connection type
        Write-Host ""
        Write-Host "Connection type:" -ForegroundColor Yellow
        Write-Host "  [1] IP Address (TCP/IP port)" -ForegroundColor White
        Write-Host "  [2] UNC path   (\\server\share)" -ForegroundColor White
        Write-Host ""

        do {
            $connType = Read-Host "Enter choice (1 or 2)"
        } while ($connType -ne "1" -and $connType -ne "2")

        if ($connType -eq "1") {
            # --- IP-based printer ---
            do {
                $IPAddress = Read-Host "Printer IP address"
                if ($IPAddress -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
                    Write-Host "ERROR: Invalid IP address format." -ForegroundColor Red
                    $IPAddress = $null
                }
            } while (-not $IPAddress)

            $PortName = "IP_$IPAddress"

            # Create TCP/IP port if it does not exist
            if (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue) {
                Write-Host "Port '$PortName' already exists, reusing." -ForegroundColor Gray
            }
            else {
                Write-Host "Creating TCP/IP port: $PortName..." -ForegroundColor Yellow
                try {
                    Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -ErrorAction Stop
                    Write-Host "OK: Port created." -ForegroundColor Green
                }
                catch {
                    Write-Host "ERROR: Could not create port: $($_.Exception.Message)" -ForegroundColor Red
                    continue
                }
            }

            # Driver selection
            $DriverName = Select-InstalledDriver
            if (-not $DriverName) { continue }

            # Add printer — try Add-Printer first; if it fails (e.g. device unreachable),
            # fall back to printui.dll which skips the reachability probe.
            Write-Host "Adding printer '$PrinterName'..." -ForegroundColor Yellow
            $printerAdded = $false
            try {
                Add-Printer -Name $PrinterName -PortName $PortName -DriverName $DriverName -ErrorAction Stop
                Write-Host "OK: Printer '$PrinterName' added successfully." -ForegroundColor Green
                $printerAdded = $true
            }
            catch {
                Write-Host "WARNING: Add-Printer failed ($($_.Exception.Message))" -ForegroundColor Yellow
                Write-Host "         Retrying via printui (offline-safe)..." -ForegroundColor Yellow
                try {
                    $printArgs = "/if /b `"$PrinterName`" /f `"$($DriverName)`" /r `"$PortName`" /m `"$DriverName`""
                    $p = Start-Process -FilePath "rundll32.exe" `
                        -ArgumentList "printui.dll,PrintUIEntry $printArgs" `
                        -Wait -PassThru -ErrorAction Stop
                    if ($p.ExitCode -eq 0) {
                        Write-Host "OK: Printer '$PrinterName' added via printui." -ForegroundColor Green
                        $printerAdded = $true
                    }
                    else {
                        Write-Host "ERROR: printui exited with code $($p.ExitCode)." -ForegroundColor Red
                    }
                }
                catch {
                    Write-Host "ERROR: printui fallback failed: $($_.Exception.Message)" -ForegroundColor Red
                }
            }

            if (-not $printerAdded) {
                $script:InstallationLog += [PSCustomObject]@{
                    File   = $PrinterName
                    Type   = "Network (IP)"
                    INF    = $DriverName
                    Status = "Failed"
                    Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                continue
            }

            $script:InstallationLog += [PSCustomObject]@{
                File   = $PrinterName
                Type   = "Network (IP)"
                INF    = $DriverName
                Status = "Added ($IPAddress)"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        }
        else {
            # --- UNC-based printer ---
            do {
                $UNCPath = Read-Host "UNC path (e.g. \\server\PrinterShare)"
                if (-not $UNCPath.StartsWith("\\")) {
                    Write-Host "ERROR: Path must start with \\." -ForegroundColor Red
                    $UNCPath = $null
                }
            } while (-not $UNCPath)

            Write-Host "Connecting to: $UNCPath..." -ForegroundColor Yellow
            try {
                Add-Printer -ConnectionName $UNCPath -ErrorAction Stop
                Write-Host "OK: Connected to printer at $UNCPath." -ForegroundColor Green
            }
            catch {
                Write-Host "ERROR: Could not connect to printer: $($_.Exception.Message)" -ForegroundColor Red
                $script:InstallationLog += [PSCustomObject]@{
                    File   = $PrinterName
                    Type   = "Network (UNC)"
                    INF    = "N/A"
                    Status = "Failed"
                    Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                continue
            }

            $script:InstallationLog += [PSCustomObject]@{
                File   = $PrinterName
                Type   = "Network (UNC)"
                INF    = "N/A"
                Status = "Added ($UNCPath)"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        }

        Write-Host ""
    }
}

# ===========================
# INSTALLATION SUMMARY
# ===========================

function Show-InstallationSummary {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Installation Summary" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $InstallationLog) {
        Write-Host "  No installations were performed." -ForegroundColor Yellow
        Write-Host ""
        return
    }

    foreach ($entry in $InstallationLog) {
        if ($entry.Status -like "Success*" -or $entry.Status -like "Added*") {
            $color = "Green"
        }
        elseif ($entry.Status -like "Warning*") {
            $color = "Yellow"
        }
        else {
            $color = "Red"
        }
        Write-Host "  [$($entry.Status)] $($entry.File) ($($entry.Type))" -ForegroundColor $color
    }

    Write-Host ""

    # Export log to CSV
    $LogPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "RUNEPRESS_InstallLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    try {
        $InstallationLog | Export-Csv -Path $LogPath -NoTypeInformation -ErrorAction Stop
        Write-Host "Log saved: $LogPath" -ForegroundColor Gray
    }
    catch {
        Write-Host "WARNING: Could not save log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host ""
}

# ===========================
# CLEANUP EXTRACTED FILES
# ===========================

function Invoke-CleanupPrompt {
    if (-not (Test-Path $ExtractRoot)) { return }

    if ($Unattended) {
        try {
            Remove-Item $ExtractRoot -Recurse -Force -ErrorAction Stop
            Write-Host "OK: Extracted files removed (unattended cleanup)." -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Could not remove extracted files: $($_.Exception.Message)" -ForegroundColor Red
        }
        return
    }

    Write-Host ""
    $response = Read-Host "Delete extracted driver files in '$ExtractRoot'? (Y/N)"
    if ($response.ToUpper() -ne "Y") { return }

    try {
        Remove-Item $ExtractRoot -Recurse -Force -ErrorAction Stop
        Write-Host "OK: Extracted files removed." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Could not remove extracted files: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ===========================
# MAIN
# ===========================

if (-not $Unattended) { Show-Banner }
if (-not $Unattended) {
    Show-DriverPrepInstructions
    Wait-ForDriverFile
}

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " Step 2: Driver Installation" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

$DriverFiles = Find-DriverFiles

if (-not $DriverFiles) {
    Write-Host ""
    Write-Host "ERROR: No driver files (.zip, .exe, .msi) found in:" -ForegroundColor Red
    Write-Host "  $ScriptPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Place a driver file in that directory and re-run the script." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "Found $(@($DriverFiles).Count) driver file(s) to process:" -ForegroundColor Green
foreach ($file in $DriverFiles) {
    Write-Host "  * $($file.Name)" -ForegroundColor White
}

foreach ($file in $DriverFiles) {
    switch ($file.Extension.ToLower()) {
        ".zip" { Install-ZipDriver -ZipFile $file }
        ".exe" { Install-ExeDriver -ExeFile $file }
        ".msi" { Install-MsiDriver -MsiFile $file }
    }
}

if (-not $Unattended) { Add-NetworkPrinter }

Show-InstallationSummary

Invoke-CleanupPrompt
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
