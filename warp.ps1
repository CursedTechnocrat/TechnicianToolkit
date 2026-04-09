<#
.SYNOPSIS
    W.A.R.P. - Winget Application Rollout Platform
.DESCRIPTION
    Automated software installation and management with SOC compliance tracking
.PARAMETERS
    Mode                     Install, Update, or Both (default: Both)
    SkipOptional             Skip optional software selection prompt
.EXAMPLE
    .\warp.ps1
    .\warp.ps1 -Mode Install
    .\warp.ps1 -Mode Update -SkipOptional
.NOTES
    Requires Administrator privileges
    Requires Windows Package Manager (winget) to be installed
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Install", "Update", "Both")]
    [string]$Mode = "Both",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipOptional
)

# ─────────────────────────────────────────────────────────────────────────────
# W.A.R.P. - Winget Application Rollout Platform
# Automated Package Manager Setup & Installation
# ─────────────────────────────────────────────────────────────────────────────

# ================================================================
# CONFIGURATION - EASILY EDITABLE
# ================================================================

# Log directory - CHANGE THIS PATH TO YOUR PREFERRED LOCATION
$LogDirectory = "C:\The20dir"

# ================================================================
# ADMIN CHECK
# ================================================================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# ================================================================
# COLOR STANDARD
# ================================================================
$Colors = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ================================================================
# BANNER DISPLAY
# ================================================================

function Show-WarpBanner {
    Write-Host @"
  ██╗    ██╗     █████╗     ██████╗     ██████╗ 
  ██║    ██║    ██╔══██╗    ██╔══██╗    ██╔══██╗
  ██║ █╗ ██║    ███████║    ██████╔╝    ██████╔╝
  ██║███╗██║    ██╔══██║    ██╔══██╗    ██╔═══╝ 
  ╚███╔███╔╝    ██║  ██║    ██║  ██║    ██║     
   ╚══╝╚══╝     ╚═╝  ╚═╝    ╚═╝  ╚═╝    ╚═╝     
"@ -ForegroundColor $Colors.Accent
    Write-Host "    W.A.R.P. - Winget Application Rollout Platform" -ForegroundColor $Colors.Header
    Write-Host "    Automated Package Manager Setup & Installation" -ForegroundColor $Colors.Header
    Write-Host ""
}

# ================================================================
# SOFTWARE LISTS
# ================================================================

# Define software lists
$RequiredSoftware = @(
    "Microsoft.Teams",
    "Microsoft.Office",
    "7zip.7zip",
    "Google.Chrome",
    "Zoom.Zoom",
    "Adobe.Acrobat.Reader.32-bit"
)

$OptionalSoftware = @(
    "Zoom.ZoomOutlookPlugin",
    "Mozilla.Firefox",
    "Dell.CommandUpdate"
)

# ================================================================
# CREATE LOG DIRECTORY
# ================================================================

# Create log directory if it doesn't exist
if (-not (Test-Path $LogDirectory)) {
    try {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        Write-Host "Created log directory: $LogDirectory" -ForegroundColor $Colors.Success
    }
    catch {
        Write-Host "ERROR: Could not create log directory at $LogDirectory" -ForegroundColor $Colors.Error
        Write-Host "Please check the path and permissions, or modify the `$LogDirectory variable in the script." -ForegroundColor $Colors.Error
        exit 1
    }
}

# Create log files
$LogPath = "$LogDirectory\SoftwareInstall_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
$ComplianceLogPath = "$LogDirectory\Compliance_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
$Results = @()
$ComplianceResults = @()

# ─────────────────────────────────────────────────────────────────────────────
# SOC COMPLIANCE TRACKING
# ─────────────────────────────────────────────────────────────────────────────

function Write-ComplianceLog {
    param(
        [string]$Package,
        [string]$Status,
        [string]$Version = "N/A",
        [string]$Action = "Install"
    )
    
    $ComplianceEntry = @{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        Computer = $env:COMPUTERNAME
        Package = $Package
        Action = $Action
        Status = $Status
        Version = $Version
        Mode = $Mode
    }
    
    $ComplianceLog = "[$($ComplianceEntry.Timestamp)] | User: $($ComplianceEntry.User) | Computer: $($ComplianceEntry.Computer) | Package: $($ComplianceEntry.Package) | Action: $($ComplianceEntry.Action) | Status: $($ComplianceEntry.Status) | Version: $($ComplianceEntry.Version)"
    
    Add-Content -Path $ComplianceLogPath -Value $ComplianceLog
    $ComplianceResults += [PSCustomObject]$ComplianceEntry
}

function Get-InstalledSoftwareVersion {
    param([string]$PackageId)
    
    try {
        $Package = winget list --id $PackageId -q 2>$null | Select-String $PackageId
        if ($Package) {
            $VersionMatch = [regex]::Match($Package, '\d+\.\d+(\.\d+)*')
            if ($VersionMatch.Success) {
                return $VersionMatch.Value
            }
        }
    }
    catch {
        return "Unknown"
    }
    
    return "Unknown"
}

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    # Determine color based on level
    $ForegroundColor = switch ($Level) {
        "SUCCESS" { $Colors.Success }
        "WARNING" { $Colors.Warning }
        "ERROR"   { $Colors.Error }
        "INFO"    { $Colors.Info }
        default   { $Colors.Info }
    }
    
    Write-Host $LogMessage -ForegroundColor $ForegroundColor
    Add-Content -Path $LogPath -Value $LogMessage
}

function Select-OptionalSoftware {
    Write-Host "`n" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "SELECT OPTIONAL SOFTWARE" -ForegroundColor $Colors.Header
    Write-Host "========================================`n" -ForegroundColor $Colors.Header
    Write-Host "Enter the numbers of the software you want to install (comma-separated)" -ForegroundColor $Colors.Info
    Write-Host "Example: 1,3 (or leave blank to skip all)`n" -ForegroundColor $Colors.Info
    
    Write-Host "0. Install ALL optional software" -ForegroundColor $Colors.Success
    for ($i = 0; $i -lt $OptionalSoftware.Count; $i++) {
        Write-Host "$($i + 1). $($OptionalSoftware[$i])" -ForegroundColor $Colors.Info
    }
    Write-Host ""
    
    do {
        $Selection = Read-Host "Enter your choices"
        
        # If empty, skip all
        if ([string]::IsNullOrWhiteSpace($Selection)) {
            Write-Host "Skipping all optional software." -ForegroundColor $Colors.Success
            return @()
        }
        
        # If user enters 0, install all
        if ($Selection -eq "0") {
            Write-Host "Installing all optional software." -ForegroundColor $Colors.Success
            return $OptionalSoftware
        }
        
        # Parse input
        $Selections = $Selection -split ',' | ForEach-Object { $_.Trim() }
        $Valid = $true
        $SelectedIndices = @()
        
        foreach ($Sel in $Selections) {
            if ($Sel -notmatch '^\d+$' -or [int]$Sel -lt 1 -or [int]$Sel -gt $OptionalSoftware.Count) {
                Write-Host "Invalid input: $Sel. Please enter numbers between 1 and $($OptionalSoftware.Count), or 0 for all" -ForegroundColor $Colors.Error
                $Valid = $false
                break
            }
            $SelectedIndices += [int]$Sel - 1
        }
        
        if ($Valid) {
            # Remove duplicates
            $SelectedIndices = $SelectedIndices | Sort-Object -Unique
        }
    } while (-not $Valid)
    
    # Return selected software
    $Selected = @()
    foreach ($Index in $SelectedIndices) {
        $Selected += $OptionalSoftware[$Index]
    }
    
    return $Selected
}

function Install-Software {
    param(
        [string]$PackageId,
        [bool]$IsRequired = $true
    )
    
    $Type = if ($IsRequired) { "Required" } else { "Optional" }
    
    Write-Log "Processing $Type software: $PackageId" "INFO"
    
    try {
        # Check if already installed
        $Installed = winget list --id $PackageId -q 2>$null | Select-String $PackageId
        
        if ($Installed) {
            Write-Log "$PackageId is already installed. Skipping." "WARNING"
            $Version = Get-InstalledSoftwareVersion -PackageId $PackageId
            Write-ComplianceLog -Package $PackageId -Status "Already Installed" -Version $Version -Action "Skip"
            
            return @{
                Package = $PackageId
                Status = "Already Installed"
                Type = $Type
                Version = $Version
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        # Install the software
        Write-Log "Installing $PackageId..." "INFO"
        $InstallOutput = & winget install --id $PackageId --accept-package-agreements --accept-source-agreements -q 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "$PackageId installed successfully." "SUCCESS"
            $Version = Get-InstalledSoftwareVersion -PackageId $PackageId
            Write-ComplianceLog -Package $PackageId -Status "Installed" -Version $Version -Action "Install"
            
            return @{
                Package = $PackageId
                Status = "Installed"
                Type = $Type
                Version = $Version
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        else {
            Write-Log "$PackageId installation failed. Exit code: $LASTEXITCODE" "ERROR"
            Write-ComplianceLog -Package $PackageId -Status "Installation Failed" -Version "N/A" -Action "Install"
            
            return @{
                Package = $PackageId
                Status = "Installation Failed"
                Type = $Type
                Version = "N/A"
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    catch {
        Write-Log "Error processing $PackageId : $($_.Exception.Message)" "ERROR"
        Write-ComplianceLog -Package $PackageId -Status "Error" -Version "N/A" -Action "Install"
        
        return @{
            Package = $PackageId
            Status = "Error"
            Type = $Type
            Version = "N/A"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
}

function Update-Software {
    Write-Log "Checking for software updates..." "INFO"
    Write-ComplianceLog -Package "All Packages" -Status "Update Check Started" -Action "Update"
    
    try {
        $UpdateOutput = & winget upgrade --all --accept-package-agreements --accept-source-agreements 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "All software updated successfully." "SUCCESS"
            Write-Log "Update Output: $UpdateOutput" "INFO"
            Write-ComplianceLog -Package "All Packages" -Status "Updates Completed" -Action "Update"
            
            return @{
                Status = "Updates Complete"
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        else {
            Write-Log "Update process completed with exit code: $LASTEXITCODE" "WARNING"
            Write-ComplianceLog -Package "All Packages" -Status "Updates Completed with Warnings" -Action "Update"
            
            return @{
                Status = "Updates Completed (with warnings)"
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    catch {
        Write-Log "Error during update: $($_.Exception.Message)" "ERROR"
        Write-ComplianceLog -Package "All Packages" -Status "Update Error" -Action "Update"
        
        return @{
            Status = "Update Error"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

# Display Banner
Show-WarpBanner

Write-Log "========================================" "INFO"
Write-Log "W.A.R.P. Software Manager Started" "INFO"
Write-Log "Mode: $Mode | Skip Optional: $SkipOptional" "INFO"
Write-Log "Log Directory: $LogDirectory" "INFO"
Write-Log "Executed by: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" "INFO"
Write-Log "Computer: $env:COMPUTERNAME" "INFO"
Write-Log "========================================" "INFO"

Write-ComplianceLog -Package "Script Execution" -Status "Started" -Action "Initialize"

# Determine which optional software to install
$SelectedOptional = @()
if (-not $SkipOptional -and ($Mode -eq "Install" -or $Mode -eq "Both")) {
    $SelectedOptional = Select-OptionalSoftware
    if ($SelectedOptional.Count -gt 0) {
        Write-Log "Selected optional software: $($SelectedOptional -join ', ')" "INFO"
        Write-ComplianceLog -Package "Optional Selection" -Status "Selected: $($SelectedOptional -join ', ')" -Action "Selection"
    }
    else {
        Write-Log "No optional software selected." "INFO"
        Write-ComplianceLog -Package "Optional Selection" -Status "None Selected" -Action "Selection"
    }
}

# Process Required Software
if ($Mode -eq "Install" -or $Mode -eq "Both") {
    Write-Log "Installing Required Software..." "INFO"
    foreach ($Software in $RequiredSoftware) {
        $Result = Install-Software -PackageId $Software -IsRequired $true
        $Results += [PSCustomObject]$Result
    }
    
    # Install selected optional software
    if ($SelectedOptional.Count -gt 0) {
        Write-Log "Installing Selected Optional Software..." "INFO"
        foreach ($Software in $SelectedOptional) {
            $Result = Install-Software -PackageId $Software -IsRequired $false
            $Results += [PSCustomObject]$Result
        }
    }
}

# Process Updates
if ($Mode -eq "Update" -or $Mode -eq "Both") {
    Write-Log "Running Update Mode..." "INFO"
    $UpdateResult = Update-Software
    $Results += [PSCustomObject]$UpdateResult
}

# Summary Report
Write-Log "========================================" "INFO"
Write-Log "SUMMARY REPORT" "INFO"
Write-Log "========================================" "INFO"

$Results | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Log $_ "INFO" }

Write-Log "Script completed. Log saved to: $LogPath" "SUCCESS"
Write-Log "Compliance log saved to: $ComplianceLogPath" "SUCCESS"
Write-Log "========================================" "INFO"

# Display summary counts
$Successful = @($Results | Where-Object { $_.Status -like "*Installed*" -or $_.Status -like "*Updates*" }).Count
$Failed = @($Results | Where-Object { $_.Status -like "*Failed*" -or $_.Status -like "*Error*" }).Count

Write-Host "`nSummary: $Successful successful, $Failed failed operations" -ForegroundColor $Colors.Success

# Write Compliance Summary
Write-Host "`n========================================" -ForegroundColor $Colors.Header
Write-Host "COMPLIANCE LOG SUMMARY" -ForegroundColor $Colors.Header
Write-Host "========================================`n" -ForegroundColor $Colors.Header
$ComplianceResults | Format-Table -AutoSize @{Label="Timestamp";Expression={$_.Timestamp}}, @{Label="User";Expression={$_.User}}, @{Label="Computer";Expression={$_.Computer}}, @{Label="Package";Expression={$_.Package}}, @{Label="Action";Expression={$_.Action}}, @{Label="Status";Expression={$_.Status}}, @{Label="Version";Expression={$_.Version}} | Out-Host

Write-Host "Compliance records: $($ComplianceResults.Count)" -ForegroundColor $Colors.Success
Write-Host "Log directory: $LogDirectory" -ForegroundColor $Colors.Info
