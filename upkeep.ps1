<#
.SYNOPSIS
    U.P.K.E.E.P. - Update Package Keeping Everything Efficiently Prepared
    
    Automated Windows Update & Maintenance Tool for PowerShell 5.1+

.DESCRIPTION
    U.P.K.E.E.P. is a comprehensive Windows Update management script designed to 
    automate the process of checking, installing, and managing Windows updates with 
    minimal user intervention. The script handles power management, module installation, 
    update deployment, and intelligent reboot detection.

.FUNCTIONALITY
    The script performs the following operations in sequence:
    
    1. Sleep Configuration Management
       - Disables system sleep and monitor timeout on AC and DC power
       - Ensures the system remains active during update installation
    
    2. Module Preparation
       - Detects and installs the PSWindowsUpdate PowerShell module if needed
       - Imports the module for use in the update process
    
    3. Update Detection & Installation
       - Scans for available Windows updates (excluding drivers)
       - Displays a summary of updates to be installed
       - Installs updates without forcing an automatic reboot
    
    4. Reboot Status Detection
       - Uses Get-WindowsUpdateRebootStatus to check if reboot is required
       - Only prompts user if reboot is necessary
       - Provides 30-second countdown with cancel option before rebooting

.REQUIREMENTS
    - Windows PowerShell 5.1 or later
    - Administrator privileges (script will exit if not run as admin)
    - Internet connectivity (for module and update downloads)
    - PSWindowsUpdate PowerShell module (auto-installed if missing)

.PARAMETERS
    This script does not accept command-line parameters.
    All interaction is handled through user prompts during execution.

.EXAMPLES
    Run the script from PowerShell as Administrator:
    
    PS C:\> .\Upkeep.ps1
    
    The script will:
    - Display the U.P.K.E.E.P. banner
    - Execute all four installation steps
    - Prompt user if reboot is required
    - Handle reboot countdown with escape option

.NOTES
    Author:         [Your Name/Organization]
    Created:        [Date]
    Last Modified:  [Date]
    Version:        1.0
    
    Color Schema:
    - Cyan      : Headers and section dividers
    - Magenta   : Progress indicators and current step
    - Green     : Success messages and confirmations
    - Yellow    : Warnings and user cautions
    - Red       : Critical errors and important alerts
    - Gray      : Information, details, and supporting text
    - Blue      : Progress bars and accent highlights
    
    Toolbox Integration:
    This script is part of an automated toolbox alongside:
    - M.A.G.I.C. (Machine Automated Graphical Ink Configurator)
    - S.P.A.R.K. (Software Package & Resource Kit)

.TROUBLESHOOTING
    Administrator Check Failed:
    - Run PowerShell as Administrator
    - Use "Run as Administrator" context menu option
    
    PSWindowsUpdate Installation Fails:
    - Check internet connectivity
    - Verify PowerShell execution policy allows module installation
    - Try: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    
    Updates Installation Fails:
    - Ensure Windows Update service is running
    - Check available disk space (at least 2GB recommended)
    - Verify internet connectivity
    
    Reboot Detection Issues:
    - Restart PowerShell and re-run the script
    - Manually check reboot status: Get-WindowsUpdateRebootStatus

.DISCLAIMER
    This script modifies system power settings and may install updates that 
    require a system reboot. Ensure all unsaved work is backed up before running 
    this script. Use at your own risk.

#>

# U.P.K.E.E.P. - Update Package Keeping Everything Efficiently Prepared
# Run as Administrator check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsRoleLevel]::Administrator)) {
    Write-Host "This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

function Show-UpkeepBanner {
    Write-Host @"

  ██╗   ██╗██████╗ ██╗  ██╗███████╗███████╗██████╗ 
  ██║   ██║██╔══██╗██║ ██╔╝██╔════╝██╔════╝██╔══██╗
  ██║   ██║██████╔╝█████╔╝ █████╗  █████╗  ██████╔╝
  ██║   ██║██╔═══╝ ██╔═██╗ ██╔══╝  ██╔══╝  ██╔═══╝ 
  ╚██╗ ██╔╝██║     ██║  ██╗███████╗███████╗██║     
   ╚████╔╝ ╚═╝     ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝     
                                                    
"@ -ForegroundColor Cyan
    Write-Host "    Update Package Keeping Everything Efficiently Prepared" -ForegroundColor Cyan
    Write-Host "    Automated Windows Update & Maintenance Tool" -ForegroundColor Cyan
    Write-Host ""
}

# Color Schema Definition
$ColorSchema = @{
    Header       = 'Cyan'      # Section headers
    Success      = 'Green'     # Successful operations
    Warning      = 'Yellow'    # Warnings and cautions
    Error        = 'Red'       # Critical errors
    Info         = 'Gray'      # Information and details
    Progress     = 'Magenta'   # Progress indicators
    Accent       = 'Blue'      # Accent and highlights
}

Show-UpkeepBanner

Write-Host ""
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host "     WINDOWS UPDATE MANAGER" -ForegroundColor $ColorSchema.Header
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host ""

# Step 1: Disable Sleep
Write-Host "[1/4] Disabling Sleep Settings..." -ForegroundColor $ColorSchema.Progress
try {
    powercfg /change standby-timeout-ac 0
    powercfg /change standby-timeout-dc 0
    powercfg /change monitor-timeout-ac 0
    powercfg /change monitor-timeout-dc 0
    Write-Host "[+] Sleep settings disabled successfully" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error disabling sleep settings: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# Step 2: Install PSWindowsUpdate Module
Write-Host "[2/4] Installing PSWindowsUpdate Module..." -ForegroundColor $ColorSchema.Progress
try {
    $module = Get-Module -Name PSWindowsUpdate -ListAvailable
    if ($null -eq $module) {
        Write-Host "    Installing module (this may take a moment)..." -ForegroundColor $ColorSchema.Info
        Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
        Write-Host "[+] PSWindowsUpdate installed successfully" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "[+] PSWindowsUpdate is already installed" -ForegroundColor $ColorSchema.Success
    }
}
catch {
    Write-Host "[-] Error installing PSWindowsUpdate: $_" -ForegroundColor $ColorSchema.Error
    exit 1
}

Write-Host ""

# Step 3: Import PSWindowsUpdate Module
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

# Step 4: Install Windows Updates (without reboot)
Write-Host "[4/4] Installing Windows Updates..." -ForegroundColor $ColorSchema.Progress
Write-Host "    This may take several minutes..." -ForegroundColor $ColorSchema.Info
Write-Host ""

try {
    # Get updates without installing first to show what will be installed
    $updates = Get-WindowsUpdate
    
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

# Check Reboot Status
Write-Host "Checking reboot status..." -ForegroundColor $ColorSchema.Progress
$rebootStatus = Get-WindowsUpdateRebootStatus
$rebootRequired = $rebootStatus.RebootRequired

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

# Reboot Decision - Only prompt if reboot is required
if ($rebootRequired) {
    Write-Host "The computer is ready to be rebooted." -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    
    $rebootPrompt = Read-Host "Is it safe to reboot this computer now? (Y/N)"
    
    if ($rebootPrompt -eq 'Y' -or $rebootPrompt -eq 'y') {
        Write-Host ""
        Write-Host "Initiating reboot in 30 seconds. Press Ctrl+C to cancel..." -ForegroundColor $ColorSchema.Warning
        Write-Host ""
        Write-Host "   30 [============================================]" -ForegroundColor $ColorSchema.Accent
        
        # 30 second countdown with ASCII progress bar
        for ($i = 30; $i -gt 0; $i--) {
            $progress = [math]::Floor((30 - $i) / 30 * 44)
            $bar = "=" * $progress
            $remaining = " " * (44 - $progress)
            Write-Host -NoNewline "`r   $i  [$bar$remaining]" -ForegroundColor $ColorSchema.Accent
            Start-Sleep -Seconds 1
        }
        
        Write-Host ""
        Write-Host ""
        Write-Host "Rebooting now..." -ForegroundColor $ColorSchema.Warning
        Write-Host ""
        Restart-Computer -Force
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
else {
    Write-Host "[+] No reboot action required. System is ready to use." -ForegroundColor $ColorSchema.Success
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host "  SCRIPT EXECUTION COMPLETED" -ForegroundColor $ColorSchema.Header
Write-Host "========================================" -ForegroundColor $ColorSchema.Header
Write-Host ""
