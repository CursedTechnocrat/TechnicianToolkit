<#
.SYNOPSIS
    W.A.R.P. - Winget Application Rollout Platform
.DESCRIPTION
    Automated software installation and management with SOC compliance tracking
.PARAMETER Mode
    Install, Update, or Both (default: Both)
.PARAMETER SkipOptional
    Skip optional software selection prompt
.NOTES
    Requires Administrator privileges
    Requires Windows Package Manager (winget)
#>

param(
    [ValidateSet("Install", "Update", "Both")]
    [string]$Mode = "Both",

    [switch]$SkipOptional
)

# ===========================
# CONFIGURATION
# ===========================

$LogDirectory = "C:\The20dir"

$RequiredSoftware = @(
    "7zip.7zip",
    "Google.Chrome",
    "Mozilla.Firefox",
    "VLC.VLC",
    "Notepad++.Notepad++"
)

$OptionalSoftware = @(
    "Balena.Etcher",
    "PuTTY.PuTTY",
    "Gimp.Gimp"
)

# ===========================
# ADMIN CHECK
# ===========================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script requires Administrator privileges." -ForegroundColor Red
    exit 1
}

# ===========================
# COLORS
# ===========================

$Colors = @{
    Header  = 'Cyan'
    Success = 'Green'
    Warning = 'Yellow'
    Error   = 'Red'
    Info    = 'Gray'
    Accent  = 'Blue'
}

# ===========================
# BANNER
# ===========================

function Show-WarpBanner {
    Write-Host ""
    Write-Host "  W.A.R.P. - Winget Application Rollout Platform" -ForegroundColor $Colors.Accent
    Write-Host "  Automated Package Manager Setup & Installation" -ForegroundColor $Colors.Header
    Write-Host ""
}

# ===========================
# LOGGING SETUP
# ===========================

if (-not (Test-Path $LogDirectory)) {
    try {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        Write-Host "[OK] Created log directory: $LogDirectory" -ForegroundColor $Colors.Success
    }
    catch {
        Write-Host "[ERROR] Failed to create log directory" -ForegroundColor $Colors.Error
        exit 1
    }
}

$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogPath = "$LogDirectory\Install_$Timestamp.log"
$ComplianceLogPath = "$LogDirectory\Compliance_$Timestamp.log"

@"
===========================================
W.A.R.P. Execution Log
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Mode: $Mode
User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
Computer: $env:COMPUTERNAME
===========================================
"@ | Out-File $LogPath -Encoding UTF8

$Results = @()
$ComplianceResults = @()

# ===========================
# LOGGING FUNCTIONS
# ===========================

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $LogColor = switch ($Level) {
        "SUCCESS" { $Colors.Success }
        "WARNING" { $Colors.Warning }
        "ERROR"   { $Colors.Error }
        "INFO"    { $Colors.Info }
        default   { $Colors.Info }
    }
    
    $Symbol = switch ($Level) {
        "SUCCESS" { "[+]" }
        "WARNING" { "[!]" }
        "ERROR"   { "[X]" }
        "INFO"    { "[*]" }
        default   { "[*]" }
    }
    
    $Line = "$Symbol [$((Get-Date).ToString('HH:mm:ss'))] [$Level] $Message"
    Write-Host $Line -ForegroundColor $LogColor
    Add-Content -Path $LogPath -Value $Line -Encoding UTF8
}

function Write-Compliance {
    param(
        [string]$Package,
        [string]$Action,
        [string]$Status,
        [string]$Version = "N/A"
    )
    
    $Entry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        User      = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        Computer  = $env:COMPUTERNAME
        Package   = $Package
        Action    = $Action
        Status    = $Status
        Version   = $Version
        Mode      = $Mode
    }
    
    $ComplianceResults += $Entry
    
    $LineContent = "$($Entry.Timestamp)|$($Entry.User)|$($Entry.Computer)|$Package|$Action|$Status|$Version"
    Add-Content -Path $ComplianceLogPath -Value $LineContent -Encoding UTF8
}

# ===========================
# CORE FUNCTIONS
# ===========================

function Test-WingetAvailable {
    try {
        $output = & winget --version 2>$null
        return $LASTEXITCODE -eq 0
    }
    catch {
        return $false
    }
}

function Get-PackageVersion {
    param([string]$PackageId)
    
    try {
        $list = & winget list --id $PackageId --exact 2>$null
        if ($list) {
            $lines = $list -split "`n"
            foreach ($line in $lines) {
                if ($line -match $PackageId) {
                    $parts = $line -split '\s+' | Where-Object { $_ }
                    if ($parts.Count -ge 3) {
                        return $parts[-1]
                    }
                }
            }
        }
    }
    catch { }
    
    return "Unknown"
}

function Test-PackageInstalledAccurate {
    param([string]$PackageId)
    
    try {
        $list = & winget list --id $PackageId --exact 2>&1
        
        if ($list -and $list -match [regex]::Escape($PackageId)) {
            return $true
        }
    }
    catch { }
    
    return $false
}

function Install-PackageViaWinget {
    param(
        [string]$PackageId,
        [bool]$IsRequired = $true
    )
    
    $Type = if ($IsRequired) { "Required" } else { "Optional" }
    
    Write-Log "Installing $Type: $PackageId" "INFO"
    
    try {
        $installResult = & winget install `
            --id $PackageId `
            --accept-package-agreements `
            --accept-source-agreements `
            --silent `
            2>&1
        
        Start-Sleep -Seconds 2
        
        if (Test-PackageInstalledAccurate $PackageId) {
            $version = Get-PackageVersion $PackageId
            Write-Log "$PackageId installed successfully (v$version)" "SUCCESS"
            Write-Compliance $PackageId "Install" "Success" $version
            
            return @{
                Package = $PackageId
                Status = "INSTALLED"
                Version = $version
                Type = $Type
            }
        }
        else {
            Write-Log "$PackageId installation may have failed - could not verify" "ERROR"
            Write-Compliance $PackageId "Install" "Failed - Verification" "N/A"
            
            return @{
                Package = $PackageId
                Status = "FAILED"
                Version = "N/A"
                Type = $Type
            }
        }
    }
    catch {
        Write-Log "Error installing $PackageId : $($_.Exception.Message)" "ERROR"
        Write-Compliance $PackageId "Install" "Error" "N/A"
        
        return @{
            Package = $PackageId
            Status = "ERROR"
            Version = "N/A"
            Type = $Type
        }
    }
}

function Select-OptionalSoftware {
    if ($SkipOptional) {
        Write-Log "Skipping optional software selection" "WARNING"
        return @()
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "SELECT OPTIONAL SOFTWARE" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    
    for ($i = 0; $i -lt $OptionalSoftware.Count; $i++) {
        Write-Host "$($i + 1). $($OptionalSoftware[$i])" -ForegroundColor $Colors.Info
    }
    
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor $Colors.Header
    Write-Host "  - Enter comma-separated numbers: 1,2,3" -ForegroundColor $Colors.Info
    Write-Host "  - Leave blank to skip optional software" -ForegroundColor $Colors.Info
    Write-Host ""
    
    $input = Read-Host "Select packages"
    
    if ([string]::IsNullOrWhiteSpace($input)) {
        Write-Log "No optional software selected" "INFO"
        return @()
    }
    
    $selected = @()
    $input -split ',' | ForEach-Object {
        $index = [int]$_.Trim() - 1
        if ($index -ge 0 -and $index -lt $OptionalSoftware.Count) {
            $selected += $OptionalSoftware[$index]
        }
    }
    
    if ($selected.Count -gt 0) {
        Write-Log "Selected optional software: $($selected -join ', ')" "INFO"
    }
    
    return $selected
}

function Update-AllPackages {
    Write-Log "Checking for package updates..." "INFO"
    
    try {
        $upgradeResult = & winget upgrade --all `
            --accept-package-agreements `
            --accept-source-agreements `
            2>&1
        
        Write-Log "Package updates completed" "SUCCESS"
        Write-Compliance "All Packages" "Update" "Success"
        
        return @{
            Action = "Update"
            Status = "COMPLETED"
            Details = "Updates processed"
        }
    }
    catch {
        Write-Log "Error during update: $($_.Exception.Message)" "ERROR"
        Write-Compliance "All Packages" "Update" "Failed"
        
        return @{
            Action = "Update"
            Status = "FAILED"
            Details = $_.Exception.Message
        }
    }
}

# ===========================
# MAIN EXECUTION
# ===========================

Show-WarpBanner

Write-Log "========================================" "INFO"
Write-Log "W.A.R.P. Script Started" "INFO"
Write-Log "Mode: $Mode | Skip Optional: $SkipOptional" "INFO"
Write-Log "========================================" "INFO"

Write-Compliance "Script" "Start" "Initialized"

if (-not (Test-WingetAvailable)) {
    Write-Log "Windows Package Manager (winget) is not available" "ERROR"
    exit 1
}

Write-Log "Windows Package Manager is available" "SUCCESS"

if ($Mode -in @("Install", "Both")) {
    Write-Host ""
    Write-Log "======= INSTALLING REQUIRED SOFTWARE =======" "INFO"
    
    foreach ($package in $RequiredSoftware) {
        $result = Install-PackageViaWinget -PackageId $package -IsRequired $true
        $Results += [PSCustomObject]$result
    }
    
    Write-Host ""
    Write-Log "======= OPTIONAL SOFTWARE =======" "INFO"
    
    $optionalList = Select-OptionalSoftware
    
    if ($optionalList.Count -gt 0) {
        foreach ($package in $optionalList) {
            $result = Install-PackageViaWinget -PackageId $package -IsRequired $false
            $Results += [PSCustomObject]$result
        }
    }
}

if ($Mode -in @("Update", "Both")) {
    Write-Host ""
    Write-Log "======= RUNNING PACKAGE UPDATES =======" "INFO"
    
    $updateResult = Update-AllPackages
    $Results += [PSCustomObject]$updateResult
}

Write-Host ""
Write-Log "========================================" "INFO"
Write-Log "EXECUTION SUMMARY" "INFO"
Write-Log "========================================" "INFO"

Write-Host ""
$Results | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Host $_ }

Write-Host ""
$successCount = @($Results | Where-Object { $_.Status -eq "INSTALLED" }).Count
$failCount = @($Results | Where-Object { $_.Status -like "*FAILED*" -or $_.Status -like "*ERROR*" }).Count

Write-Log "Total operations: $($Results.Count) | Successful: $successCount | Failed: $failCount" "INFO"
Write-Log "Installation logs: $LogPath" "SUCCESS"
Write-Log "Compliance logs: $ComplianceLogPath" "SUCCESS"
Write-Log "========================================" "INFO"

Write-Host ""
Write-Host "Script completed! Check logs in: $LogDirectory" -ForegroundColor $Colors.Success
