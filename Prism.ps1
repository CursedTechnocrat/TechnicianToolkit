# ╔════════════════════════════════════════════════════════════════╗
# ║                     P R I S M                                  ║
# ║           Printer Installation Script Manager                  ║
# ║                                                                ║
# ║  (Prints Remarkably Integrated System Management)             ║
# ╚════════════════════════════════════════════════════════════════╝

# Requires administrator privileges
# Place this script and driver zip files (HP.zip, Konica.zip, Xerox.zip) in the same folder

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsIdentity]::Administrator)) {
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║  PRISM - Administrator Rights Required!   ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Red
    exit
}

# Get the script's directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$printerDrivers = @{
    "HP"     = Join-Path $scriptPath "HP.zip"
    "Konica" = Join-Path $scriptPath "Konica.zip"
    "Xerox"  = Join-Path $scriptPath "Xerox.zip"
}

$extractPath = Join-Path $scriptPath "PrinterDrivers"

# Display banner
function Show-PRISMBanner {
    Write-Host "`n╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                                                           ║" -ForegroundColor Cyan
    Write-Host "║                    P R I S M v1.0                         ║" -ForegroundColor Cyan
    Write-Host "║          Printer Installation Script Manager              ║" -ForegroundColor Cyan
    Write-Host "║                                                           ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan
    Write-Host "Script location: $scriptPath" -ForegroundColor Green
    Write-Host "Looking for driver files in: $scriptPath`n" -ForegroundColor Green
}

# Function to display menu and get user selection
function Get-PrinterSelection {
    Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Printer Driver Installation" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════`n" -ForegroundColor Cyan
    Write-Host "Select printers to install (enter numbers separated by commas):`n"
    
    $printers = @("HP", "Konica", "Xerox")
    for ($i = 0; $i -lt $printers.Count; $i++) {
        Write-Host "  $($i + 1). $($printers[$i])"
    }
    
    $selection = Read-Host "`nEnter your selection (e.g., 1,2 or 1,3)"
    return $selection
}

# Function to extract and install drivers
function Install-PrinterDrivers {
    param([array]$selectedPrinters)
    
    # Create extraction directory if it doesn't exist
    if (-not (Test-Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
        Write-Host "✓ Created extraction directory: $extractPath`n" -ForegroundColor Green
    }
    
    foreach ($printer in $selectedPrinters) {
        $zipPath = $printerDrivers[$printer]
        
        if (-not (Test-Path $zipPath)) {
            Write-Host "✗ Warning: $zipPath not found. Skipping $printer...`n" -ForegroundColor Yellow
            continue
        }
        
        Write-Host "⧖ Extracting $printer drivers..." -ForegroundColor Yellow
        
        # Extract the zip file
        $printerExtractPath = Join-Path $extractPath $printer
        
        # Remove existing extraction folder
        if (Test-Path $printerExtractPath) {
            Remove-Item $printerExtractPath -Recurse -Force
        }
        
        # Extract zip file
        try {
            Expand-Archive -Path $zipPath -DestinationPath $printerExtractPath -Force
            Write-Host "✓ $printer drivers extracted successfully!`n" -ForegroundColor Green
        }
        catch {
            Write-Host "✗ Error extracting $printer drivers: $_`n" -ForegroundColor Red
        }
    }
}

# Function to add printer to system
function Add-NetworkPrinter {
    Write-Host "`n═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Add Network Printer" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════`n" -ForegroundColor Cyan
    
    $ipAddress = Read-Host "Enter printer IP address"
    
    # Validate IP address format
    if ($ipAddress -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        Write-Host "✗ Invalid IP address format!`n" -ForegroundColor Red
        return
    }
    
    Write-Host "`nSelect printer type:`n"
    Write-Host "  1. HP"
    Write-Host "  2. Konica"
    Write-Host "  3. Xerox"
    
    $typeSelection = Read-Host "`nEnter selection (1-3)"
    
    $printerTypes = @("HP", "Konica", "Xerox")
    
    if ($typeSelection -notmatch '^[1-3]$') {
        Write-Host "✗ Invalid selection!`n" -ForegroundColor Red
        return
    }
    
    $selectedType = $printerTypes[$typeSelection - 1]
    
    $printerName = Read-Host "Enter a name for this printer"
    
    if ([string]::IsNullOrWhiteSpace($printerName)) {
        Write-Host "✗ Printer name cannot be empty!`n" -ForegroundColor Red
        return
    }
    
    # Construct the printer port and add the printer
    $portName = "IP_$ipAddress"
    
    try {
        # Create the printer port
        Write-Host "`n⧖ Creating printer port..." -ForegroundColor Yellow
        $portExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
        
        if (-not $portExists) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $ipAddress -ErrorAction Stop
            Write-Host "✓ Port created successfully!" -ForegroundColor Green
        } else {
            Write-Host "✓ Port already exists." -ForegroundColor Green
        }
        
        # Add the printer
        Write-Host "⧖ Adding printer '$printerName'..." -ForegroundColor Yellow
        Add-Printer -Name $printerName -DriverName "$selectedType PCL" -PortName $portName -ErrorAction Stop
        
        Write-Host "✓ Printer '$printerName' added successfully!`n" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Error adding printer: $_`n" -ForegroundColor Red
        Write-Host "Note: Ensure the $selectedType PCL driver is installed on this system.`n" -ForegroundColor Yellow
    }
}

# Main script execution
try {
    Show-PRISMBanner
    
    # Get printer selection
    $selection = Get-PrinterSelection
    
    # Parse the selection
    $selectedNumbers = $selection -split ',' | ForEach-Object { $_.Trim() }
    $printerNames = @("HP", "Konica", "Xerox")
    $selectedPrinters = @()
    
    foreach ($num in $selectedNumbers) {
        if ($num -match '^\d+$' -and [int]$num -ge 1 -and [int]$num -le 3) {
            $selectedPrinters += $printerNames[[int]$num - 1]
        }
    }
    
    if ($selectedPrinters.Count -eq 0) {
        Write-Host "⚠ No valid printers selected.`n" -ForegroundColor Yellow
    } else {
        Write-Host "`nSelected printers: $($selectedPrinters -join ', ')`n" -ForegroundColor Cyan
        
        # Install selected drivers
        Install-PrinterDrivers -selectedPrinters $selectedPrinters
        
        Write-Host "✓ Driver installation complete!`n" -ForegroundColor Green
    }
    
    # Ask if user wants to add a printer
    $addPrinter = Read-Host "Would you like to add a network printer? (Y/N)"
    if ($addPrinter -eq "Y" -or $addPrinter -eq "y") {
        Add-NetworkPrinter
    }
    
    Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "✓ PRISM Script Completed!" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════`n" -ForegroundColor Cyan
}
catch {
    Write-Host "✗ An error occurred: $_`n" -ForegroundColor Red
}

# Pause before closing
Read-Host "Press Enter to exit"
