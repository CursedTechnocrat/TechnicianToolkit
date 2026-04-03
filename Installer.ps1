# PowerShell Script to Install Winget/Chocolatey and Multiple Software Packages

# Function to refresh environment variables
function Refresh-EnvironmentPath {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# Function to check and install Winget
function Initialize-Winget {
    try {
        $wingetVersion = winget --version
        Write-Host "Winget is already installed. Version: $wingetVersion" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Winget is not installed. Installing now..." -ForegroundColor Yellow
        try {
            $progressPreference = 'SilentlyContinue'
            irm https://aka.ms/getwinget | iex
            Start-Sleep -Seconds 2
            Refresh-EnvironmentPath
            $wingetVersion = winget --version
            Write-Host "Winget installed successfully. Version: $wingetVersion" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "Failed to install Winget: $_" -ForegroundColor Red
            return $false
        }
    }
}

# Function to check and install Chocolatey
function Initialize-Chocolatey {
    try {
        $chocoVersion = choco --version
        Write-Host "Chocolatey is already installed. Version: $chocoVersion" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Chocolatey not found in PATH. Checking if installed..." -ForegroundColor Yellow
        
        if (Test-Path "C:\ProgramData\chocolatey") {
            Write-Host "Chocolatey is installed but not in PATH. Adding to PATH..." -ForegroundColor Yellow
            Refresh-EnvironmentPath
            Start-Sleep -Seconds 1
            
            try {
                $chocoVersion = choco --version
                Write-Host "Chocolatey found. Version: $chocoVersion" -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "Chocolatey is installed but inaccessible. You may need to restart PowerShell." -ForegroundColor Yellow
                return $false
            }
        }
        else {
            Write-Host "Chocolatey is not installed. Installing now..." -ForegroundColor Yellow
            try {
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
                Start-Sleep -Seconds 2
                Refresh-EnvironmentPath
                Write-Host "Chocolatey installed successfully." -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "Failed to install Chocolatey: $_" -ForegroundColor Red
                return $false
            }
        }
    }
}

# Function to update package managers
function Update-PackageManagers {
    param(
        [bool]$UpdateWinget,
        [bool]$UpdateChocolatey
    )

    if ($UpdateWinget) {
        Write-Host "`nUpdating Winget..." -ForegroundColor Magenta
        try {
            $output = winget upgrade winget 2>&1
            if ($output -match "No package found matching input criteria" -or $output -match "No upgrades available") {
                Write-Host "Winget is already up to date." -ForegroundColor Green
            }
            else {
                Write-Host "Winget updated successfully." -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Could not update Winget (this may be normal): $_" -ForegroundColor Yellow
        }
    }

    if ($UpdateChocolatey) {
        Write-Host "`nUpdating Chocolatey..." -ForegroundColor Magenta
        try {
            $output = choco upgrade chocolatey -y 2>&1
            Write-Host "Chocolatey updated successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to update Chocolatey: $_" -ForegroundColor Red
        }
    }
}

# Array of core software (installed via Winget and Chocolatey)
$coreSoftware = @(
    @{ Winget = "Microsoft.Teams"; Chocolatey = "microsoft-teams" },
    @{ Winget = "Microsoft.Office"; Chocolatey = "office-deploy" },
    @{ Winget = "7zip.7zip"; Chocolatey = "7zip" },
    @{ Winget = "Google.Chrome"; Chocolatey = "googlechrome" },
    @{ Winget = "Adobe.Acrobat.Reader.64-bit"; Chocolatey = "adobereader" },
    @{ Winget = "Zoom.Zoom"; Chocolatey = "zoom" }
)

# Array of optional software with descriptions
$optionalSoftware = @(
    @{
        Name = "Zoom Outlook Plugin"
        Description = "Integrates Zoom meetings with Microsoft Outlook"
        Winget = "Zoom.ZoomOutlookPlugin"
    },
    @{
        Name = "DisplayLink Graphics Driver"
        Description = "Driver for DisplayLink USB graphics adapters"
        Winget = "DisplayLink.GraphicsDriver"
    },
    @{
        Name = "Dell Command Update"
        Description = "Manages Dell system updates and BIOS updates"
        Winget = "Dell.CommandUpdate"
    }
)

# Installation tracking array
$installationLog = @()

# Function to install software
function Install-Software {
    param(
        [string[]]$WingetPackages,
        [string[]]$ChocoPackages,
        [bool]$UseWinget,
        [bool]$UseChocolatey,
        [ref]$LogArray
    )

    if ($UseWinget -and $WingetPackages.Count -gt 0) {
        Write-Host "`n--- Installing via Winget ---" -ForegroundColor Magenta
        foreach ($item in $WingetPackages) {
            Write-Host "Installing $item (Winget)..." -ForegroundColor Yellow
            try {
                $output = winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1
                
                # Check for actual errors vs warnings
                if ($output -match "Successfully installed" -or $output -match "No newer package found") {
                    Write-Host "$item installed/already present." -ForegroundColor Green
                    $LogArray.Value += @{
                        Package = $item
                        Status = "Success"
                        Manager = "Winget"
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                elseif ($output -match "error|failed" -and $output -notmatch "InternetOpenUrl") {
                    Write-Host "Failed to install $item via Winget." -ForegroundColor Red
                    $LogArray.Value += @{
                        Package = $item
                        Status = "Failed"
                        Manager = "Winget"
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                else {
                    Write-Host "$item processed." -ForegroundColor Green
                    $LogArray.Value += @{
                        Package = $item
                        Status = "Success"
                        Manager = "Winget"
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
            }
            catch {
                Write-Host "Error installing $item via Winget: $_" -ForegroundColor Red
                $LogArray.Value += @{
                    Package = $item
                    Status = "Failed"
                    Manager = "Winget"
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
    }

    if ($UseChocolatey -and $ChocoPackages.Count -gt 0) {
        Write-Host "`n--- Installing via Chocolatey ---" -ForegroundColor Magenta
        foreach ($item in $ChocoPackages) {
            Write-Host "Installing $item (Chocolatey)..." -ForegroundColor Yellow
            try {
                $output = choco install $item -y 2>&1
                
                if ($output -match "installed successfully|already installed") {
                    Write-Host "$item installed/already present." -ForegroundColor Green
                    $LogArray.Value += @{
                        Package = $item
                        Status = "Success"
                        Manager = "Chocolatey"
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                else {
                    Write-Host "$item processed." -ForegroundColor Green
                    $LogArray.Value += @{
                        Package = $item
                        Status = "Success"
                        Manager = "Chocolatey"
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
            }
            catch {
                Write-Host "Error installing $item via Chocolatey: $_" -ForegroundColor Red
                $LogArray.Value += @{
                    Package = $item
                    Status = "Failed"
                    Manager = "Chocolatey"
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
    }
}

# Function to ask about optional software
function Get-OptionalSoftwareSelection {
    param([array]$OptionalList)
    
    $selected = @()
    
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "Optional Software" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    
    foreach ($software in $OptionalList) {
        Write-Host "`n$($software.Name)" -ForegroundColor Yellow
        Write-Host "Description: $($software.Description)" -ForegroundColor Gray
        
        $response = Read-Host "Install $($software.Name)? [Y/N]"
        
        if ($response -eq 'Y' -or $response -eq 'y') {
            $selected += $software.Winget
            Write-Host "$($software.Name) will be installed." -ForegroundColor Green
        }
        else {
            Write-Host "$($software.Name) skipped." -ForegroundColor Yellow
        }
    }
    
    return $selected
}

# Function to display S.P.A.R.K banner
function Show-SparkBanner {
    Write-Host @"
   _____ _____  ___  _____  _  __
  / ___//  __ \/ _ \|  _  || |/ /
  \___ \| |  \/ /_\ \ | | || ' / 
   ___) | |__ |  _  | | | || . \ 
  |____/|_| \_\|_| |_|_| |_||_|\_\
"@ -ForegroundColor Yellow
    
    Write-Host "    Software Package & Resource Kit" -ForegroundColor Yellow
    Write-Host "   Automated Package Manager Setup & Installation" -ForegroundColor Yellow
    Write-Host ""
}

# Function to display installation summary
function Show-InstallationSummary {
    param(
        [array]$InstallLog
    )
    
    $successCount = ($InstallLog | Where-Object { $_.Status -eq "Success" }).Count
    $failureCount = ($InstallLog | Where-Object { $_.Status -eq "Failed" }).Count
    $totalCount = $InstallLog.Count
    
    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "S.P.A.R.K Installation Summary Report" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    
    Write-Host "Execution Timestamp: $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')" -ForegroundColor Yellow
    Write-Host ""
    
    Write-Host "Installation Statistics:" -ForegroundColor Magenta
    Write-Host "├─ Total Packages: $totalCount"
    Write-Host "├─ Successful: $successCount" -ForegroundColor Green
    
    if ($failureCount -gt 0) {
        Write-Host "└─ Failed: $failureCount" -ForegroundColor Red
    }
    else {
        Write-Host "└─ Failed: $failureCount" -ForegroundColor Green
    }
    Write-Host ""
    
    if ($successCount -gt 0) {
        Write-Host "Successfully Installed Packages:" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        $InstallLog | Where-Object { $_.Status -eq "Success" } | ForEach-Object {
            Write-Host "  [+] $($_.Package)" -ForegroundColor Green
            Write-Host "      Manager: $($_.Manager) | Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    if ($failureCount -gt 0) {
        Write-Host "Failed Package Installations:" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        $InstallLog | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
            Write-Host "  [-] $($_.Package)" -ForegroundColor Red
            Write-Host "      Manager: $($_.Manager) | Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Installation process completed at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
}

# Main Script Execution
$scriptStartTime = Get-Date
Show-SparkBanner

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Initializing S.P.A.R.K" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

# Refresh PATH at start
Refresh-EnvironmentPath

# Initialize package managers
$wingetAvailable = Initialize-Winget
$chocoAvailable = Initialize-Chocolatey

if (!$wingetAvailable -and !$chocoAvailable) {
    Write-Host "`nError: Neither Winget nor Chocolatey could be initialized. Exiting." -ForegroundColor Red
    exit 1
}

# Update package managers
Update-PackageManagers -UpdateWinget $wingetAvailable -UpdateChocolatey $chocoAvailable

# Prepare software lists for installation
$wingetList = @()
$chocoList = @()

foreach ($software in $coreSoftware) {
    if ($wingetAvailable) { $wingetList += $software.Winget }
    if ($chocoAvailable) { $chocoList += $software.Chocolatey }
}

# Ask user about each optional software individually
$selectedOptional = @()
if ($wingetAvailable) {
    $selectedOptional = Get-OptionalSoftwareSelection -OptionalList $optionalSoftware
    $wingetList += $selectedOptional
}
else {
    Write-Host "`nWinget is not available. Optional software cannot be installed." -ForegroundColor Yellow
}

# Install software
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "S.P.A.R.K - Installation Phase" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Install-Software -WingetPackages $wingetList -ChocoPackages $chocoList -UseWinget $wingetAvailable -UseChocolatey $chocoAvailable -LogArray ([ref]$installationLog)

# Display installation summary
Show-InstallationSummary -InstallLog $installationLog

Write-Host "Note: Some installations may require a system restart to complete." -ForegroundColor Yellow
Write-Host ""
