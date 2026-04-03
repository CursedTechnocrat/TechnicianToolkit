# ================================================================
#  M.A.G.I.C. - Machine Automated Graphical Ink Configurator
# ================================================================
#  Version 1.4
# ================================================================

# ===========================
# ADMIN CHECK (AUTO-ELEVATE)
# ===========================
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "INFO: Restarting script with administrator privileges..." -ForegroundColor Yellow
    $PSExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    Start-Process -FilePath $PSExe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

$ScriptPath  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExtractRoot = Join-Path $ScriptPath "ExtractedDrivers"
$InstalledManufacturers = @()

# ===========================
# BANNER
# ===========================
function Show-MAGICBanner {
    Write-Host @"
  ███╗   ███╗ █████╗  ██████╗ ██╗ ██████╗ 
  ████╗ ████║██╔══██╗██╔════╝ ██║██╔════╝ 
  ██╔████╔██║███████║██║  ███╗██║██║      
  ██║╚██╔╝██║██╔══██║██║   ██║██║██║      
  ██║ ╚═╝ ██║██║  ██║╚██████╔╝██║╚██████╗ 
  ╚═╝     ╚═╝╚═╝  ╚═╝ ╚═════╝ ╚═╝ ╚═════╝ 
"@ -ForegroundColor Cyan
    Write-Host "    Machine Automated Graphical Ink Configurator" -ForegroundColor Cyan
    Write-Host "    Printer Registration & Installation Network Tool" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Script Path: $ScriptPath" -ForegroundColor Cyan
    Write-Host ""
}

Show-MAGICBanner

# ===========================
# DRIVER PREPARATION PROMPT
# ===========================
function Driver-Preparation {
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host " Driver Preparation" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Download the printer driver from the manufacturer." -ForegroundColor White
    Write-Host "Supported formats: ZIP, EXE, MSI" -ForegroundColor White
    Write-Host ""
    Write-Host "Place the file in the same folder as this script." -ForegroundColor White
    Write-Host ""
    Write-Host "Press Enter to continue or type Q to quit." -ForegroundColor White

    $input = Read-Host
    if ($input -match '^[Qq]$') {
        Write-Host "WARNING: User exited script." -ForegroundColor Yellow
        exit
    }
}

# ===========================
# DRIVER INSTALLATION ENGINE
# ===========================
function Install-Drivers {

    $DriverFiles = Get-ChildItem -Path $ScriptPath |
                   Where-Object { $_.Extension -match '\.(zip|exe|msi)$' }

    if ($DriverFiles.Count -eq 0) {
        Write-Host "WARNING: No supported driver files found." -ForegroundColor Yellow
        return $false
    }

    if (-not (Test-Path $ExtractRoot)) {
        Write-Host "Processing: Creating extraction directory..." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $ExtractRoot | Out-Null
    }

    foreach ($File in $DriverFiles) {

        Write-Host ""
        Write-Host "================================================" -ForegroundColor Cyan
        Write-Host " Processing Driver File: $($File.Name)" -ForegroundColor Cyan
        Write-Host "================================================" -ForegroundColor Cyan

        switch ($File.Extension.ToLower()) {

            ".zip" {
                $DriverName  = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
                $ExtractPath = Join-Path $ExtractRoot $DriverName

                if (Test-Path $ExtractPath) {
                    Write-Host "Processing: Removing previous extraction..." -ForegroundColor Yellow
                    Remove-Item $ExtractPath -Recurse -Force
                }

                Write-Host "Processing: Extracting ZIP package..." -ForegroundColor Yellow
                Expand-Archive -Path $File.FullName -DestinationPath $ExtractPath -Force

                $InfFiles = Get-ChildItem -Path $ExtractPath -Recurse -Filter "*.inf"

                if ($InfFiles.Count -eq 0) {
                    Write-Host "WARNING: No INF files found in $DriverName." -ForegroundColor Yellow
                    break
                }

                foreach ($Inf in $InfFiles) {
                    Write-Host "Processing: Installing INF $($Inf.Name)..." -ForegroundColor Yellow
                    try {
                        pnputil /add-driver "`"$($Inf.FullName)`"" /install | Out-Null
                        Write-Host "OK: Installed INF $($Inf.Name)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "ERROR: Failed to install INF $($Inf.Name)" -ForegroundColor Red
                    }
                }

                $InstalledManufacturers += $DriverName
            }

            ".exe" {
                Write-Host "Processing: Running EXE installer silently..." -ForegroundColor Yellow
                try {
                    Start-Process -FilePath $File.FullName `
                        -ArgumentList "/s /quiet /norestart" `
                        -Wait -NoNewWindow
                    Write-Host "OK: EXE installer completed $($File.Name)" -ForegroundColor Green
                    $InstalledManufacturers += $File.BaseName
                }
                catch {
                    Write-Host "ERROR: EXE installer failed $($File.Name)" -ForegroundColor Red
                }
            }

            ".msi" {
                Write-Host "Processing: Running MSI installer silently..." -ForegroundColor Yellow
                try {
                    Start-Process "msiexec.exe" `
                        -ArgumentList "/i `"$($File.FullName)`" /qn /norestart" `
                        -Wait -NoNewWindow
                    Write-Host "OK: MSI installer completed $($File.Name)" -ForegroundColor Green
                    $InstalledManufacturers += $File.BaseName
                }
                catch {
                    Write-Host "ERROR: MSI installer failed $($File.Name)" -ForegroundColor Red
                }
            }
        }
    }

    return $true
}

# ===========================
# DRIVER DISCOVERY LOOP
# ===========================
do {
    Driver-Preparation
    $DriversInstalled = Install-Drivers
} until ($DriversInstalled)

# ===========================
# ADD NETWORK PRINTER
# ===========================
function Add-NetworkPrinter {

    if ($InstalledManufacturers.Count -eq 0) {
        Write-Host "ERROR: No installed drivers available." -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host " Add Network Printer" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Enter printer IP address:" -ForegroundColor White
    $IPAddress = Read-Host

    if ($IPAddress -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
        Write-Host "ERROR: Invalid IP address format." -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "Select printer manufacturer:" -ForegroundColor White
    for ($i = 0; $i -lt $InstalledManufacturers.Count; $i++) {
        Write-Host " $($i + 1). $($InstalledManufacturers[$i])" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "Enter selection number:" -ForegroundColor White
    $Selection = Read-Host

    if ($Selection -notmatch '^\d+$' -or
        $Selection -lt 1 -or
        $Selection -gt $InstalledManufacturers.Count) {
        Write-Host "ERROR: Invalid manufacturer selection." -ForegroundColor Red
        return
    }

    $Manufacturer = $InstalledManufacturers[$Selection - 1]

    Write-Host "Enter printer display name:" -ForegroundColor White
    $PrinterName = Read-Host

    if ([string]::IsNullOrWhiteSpace($PrinterName)) {
        Write-Host "ERROR: Printer name cannot be empty." -ForegroundColor Red
        return
    }

    $PortName = "IP_$IPAddress"

    try {
        if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
            Write-Host "Processing: Creating TCP/IP printer port..." -ForegroundColor Yellow
            Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress
        }

        $Driver = Get-PrinterDriver |
                  Where-Object { $_.Name -match $Manufacturer } |
                  Select-Object -First 1

        if (-not $Driver) {
            Write-Host "ERROR: No matching driver found for $Manufacturer." -ForegroundColor Red
            return
        }

        Write-Host "Processing: Using driver $($Driver.Name)" -ForegroundColor Cyan
        Write-Host "Processing: Adding network printer..." -ForegroundColor Yellow

        Add-Printer -Name $PrinterName `
                    -DriverName $Driver.Name `
                    -PortName $PortName

        Write-Host "OK: Printer added successfully - $PrinterName" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to add network printer." -ForegroundColor Red
    }
}

# ===========================
# USER PROMPT
# ===========================
Write-Host ""
Write-Host "Would you like to add a network printer now? (Y/N)" -ForegroundColor White
$AddPrinter = Read-Host

if ($AddPrinter -match '^[Yy]$') {
    Add-NetworkPrinter
}

# ===========================
# COMPLETION
# ===========================
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host " M.A.G.I.C. Completed" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Press Enter to exit..." -ForegroundColor White
Read-Host
