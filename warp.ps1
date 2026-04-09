<#
.SYNOPSIS
    W.A.R.P. - Winget Application Rollout Platform
.DESCRIPTION
    Automated software installation and management
.NOTES
    Requires Administrator privileges
    Requires Windows Package Manager (winget)
#>

# ===========================
# CONFIGURATION
# ===========================

$ScriptPath = $PSScriptRoot
$ExecutionTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

$RequiredSoftware = @(
    "Microsoft.Teams",
    "Microsoft.Office",
    "7zip.7zip",
    "Google.Chrome",
    "Adobe.Acrobat.Reader.64-bit",
    "Zoom.Zoom"
)

$OptionalSoftware = @(
    "Zoom.ZoomOutlookPlugin",
    "Mozilla.Firefox",
    "Dell.CommandUpdate"
)

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

function Show-Banner {
    Clear-Host
    Write-Host @"
╔════════════════════════════════════════════════════════════╗
║                                                            ║
║       ██╗    ██╗ █████╗ ██████╗ ██████╗                 ║
║       ██║    ██║██╔══██╗██╔══██╗██╔══██╗                ║
║       ██║ █╗ ██║███████║██████╔╝██████╔╝                ║
║       ██║███╗██║██╔══██║██╔══██╗██╔═══╝                 ║
║       ╚███╔███╔╝██║  ██║██║  ██║██║                     ║
║        ╚══╝╚══╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝                     ║
║                                                            ║
║   W.A.R.P. - Winget Application Rollout Platform         ║
║   Automated software installation and management          ║
║                                                            ║
╚════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

    Write-Host ""
    Write-Host "Script Location: $ScriptPath" -ForegroundColor Gray
    Write-Host "Execution Time:  $ExecutionTime" -ForegroundColor Gray
    Write-Host ""
}

# ===========================
# INSTALLATION TRACKING
# ===========================

$InstallationLog = @()

function Add-InstallationRecord {
    param(
        [string]$Software,
        [string]$Status,
        [string]$Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    )
    
    $record = [PSCustomObject]@{
        Timestamp = $Timestamp
        Software  = $Software
        Status    = $Status
    }
    
    $InstallationLog += $record
}

# ===========================
# WINGET CHECK
# ===========================

function Test-WingetAvailable {
    try {
        $wingetVersion = winget --version
        Write-Host "✓ Winget is available. Version: $wingetVersion" -ForegroundColor $Colors.Success
        return $true
    }
    catch {
        Write-Host "✗ Winget is not installed. Installing now..." -ForegroundColor $Colors.Warning
        try {
            $progressPreference = 'SilentlyContinue'
            irm https://aka.ms/getwinget | iex
            Write-Host "✓ Winget installed successfully" -ForegroundColor $Colors.Success
            return $true
        }
        catch {
            Write-Host "✗ Failed to install Winget" -ForegroundColor $Colors.Error
            return $false
        }
    }
}

# ===========================
# INSTALLATION FUNCTIONS
# ===========================

function Install-Software {
    param(
        [string[]]$SoftwareList,
        [string]$Type = "Required"
    )

    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "Installing $Type Software" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""

    foreach ($item in $SoftwareList) {
        Write-Host "Installing: $item..." -ForegroundColor $Colors.Info
        
        try {
            $result = & winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ $item installed successfully at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor $Colors.Success
                Add-InstallationRecord -Software $item -Status "INSTALLED"
            }
            else {
                Write-Host "! $item installation completed with warnings at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor $Colors.Warning
                Add-InstallationRecord -Software $item -Status "INSTALLED (with warnings)"
            }
        }
        catch {
            Write-Host "✗ Error installing $item : $($_.Exception.Message)" -ForegroundColor $Colors.Error
            Add-InstallationRecord -Software $item -Status "FAILED"
        }
        
        Start-Sleep -Seconds 1
    }
}

function Update-AllSoftware {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "Running Package Updates" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""

    try {
        Write-Host "Running: winget upgrade --all" -ForegroundColor $Colors.Info
        $result = & winget upgrade --all --accept-source-agreements --accept-package-agreements 2>&1
        
        Write-Host "✓ Package update completed at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor $Colors.Success
        Add-InstallationRecord -Software "All Packages" -Status "UPDATED"
    }
    catch {
        Write-Host "✗ Error during update: $($_.Exception.Message)" -ForegroundColor $Colors.Error
        Add-InstallationRecord -Software "All Packages" -Status "UPDATE FAILED"
    }
}

function Select-OptionalSoftware {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "OPTIONAL SOFTWARE SELECTION" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    
    Write-Host "Available optional software:" -ForegroundColor $Colors.Header
    for ($i = 0; $i -lt $OptionalSoftware.Count; $i++) {
        Write-Host "  $($i + 1). $($OptionalSoftware[$i])" -ForegroundColor $Colors.Info
    }
    
    Write-Host ""
    Write-Host "Options:" -ForegroundColor $Colors.Header
    Write-Host "  [A] Install all optional software" -ForegroundColor $Colors.Info
    Write-Host "  [S] Select specific software (e.g., 1,2,3)" -ForegroundColor $Colors.Info
    Write-Host "  [N] Skip optional software" -ForegroundColor $Colors.Info
    Write-Host ""
    
    $choice = Read-Host "Enter your choice (A/S/N)"
    
    $selected = @()
    
    switch ($choice.ToUpper()) {
        "A" {
            $selected = $OptionalSoftware
            Write-Host "Selected all optional software" -ForegroundColor $Colors.Success
        }
        "S" {
            $input = Read-Host "Enter numbers (comma-separated, e.g., 1,2)"
            
            if (-not [string]::IsNullOrWhiteSpace($input)) {
                $input -split ',' | ForEach-Object {
                    $index = [int]$_.Trim() - 1
                    if ($index -ge 0 -and $index -lt $OptionalSoftware.Count) {
                        $selected += $OptionalSoftware[$index]
                    }
                }
            }
            
            if ($selected.Count -gt 0) {
                Write-Host "Selected: $($selected -join ', ')" -ForegroundColor $Colors.Success
            }
        }
        "N" {
            Write-Host "Skipping optional software" -ForegroundColor $Colors.Warning
        }
        default {
            Write-Host "Invalid choice. Skipping optional software" -ForegroundColor $Colors.Warning
        }
    }
    
    return $selected
}

function Show-InstallationSummary {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "INSTALLATION SUMMARY" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    
    if ($InstallationLog.Count -gt 0) {
        $InstallationLog | Format-Table -AutoSize -Property Timestamp, Software, Status | Out-String | ForEach-Object { Write-Host $_ }
    }
    else {
        Write-Host "No installations recorded" -ForegroundColor $Colors.Warning
    }
    
    Write-Host ""
    $successCount = @($InstallationLog | Where-Object { $_.Status -like "*INSTALLED*" }).Count
    $failCount = @($InstallationLog | Where-Object { $_.Status -like "*FAILED*" }).Count
    $updateCount = @($InstallationLog | Where-Object { $_.Status -like "*UPDATED*" }).Count
    
    Write-Host "Summary: Installed: $successCount | Failed: $failCount | Updated: $updateCount" -ForegroundColor $Colors.Header
    Write-Host ""
}

# ===========================
# MAIN EXECUTION
# ===========================

Show-Banner

# Check Winget
if (-not (Test-WingetAvailable)) {
    Write-Host "Cannot proceed without Winget" -ForegroundColor $Colors.Error
    exit 1
}

Write-Host ""
Write-Host "Choose operation:" -ForegroundColor $Colors.Header
Write-Host "  [1] Install software" -ForegroundColor $Colors.Info
Write-Host "  [2] Upgrade all software" -ForegroundColor $Colors.Info
Write-Host "  [3] Install and then Upgrade" -ForegroundColor $Colors.Info
Write-Host ""

$operation = Read-Host "Enter your choice (1/2/3)"

switch ($operation) {
    "1" {
        Write-Host "Selected: Install Mode" -ForegroundColor $Colors.Success
        
        # Install required software
        Install-Software -SoftwareList $RequiredSoftware -Type "Required"
        
        # Ask for optional software
        $optionalList = Select-OptionalSoftware
        
        if ($optionalList.Count -gt 0) {
            Install-Software -SoftwareList $optionalList -Type "Optional"
        }
    }
    "2" {
        Write-Host "Selected: Upgrade Mode" -ForegroundColor $Colors.Success
        Update-AllSoftware
    }
    "3" {
        Write-Host "Selected: Install and Upgrade Mode" -ForegroundColor $Colors.Success
        
        # Install required software
        Install-Software -SoftwareList $RequiredSoftware -Type "Required"
        
        # Ask for optional software
        $optionalList = Select-OptionalSoftware
        
        if ($optionalList.Count -gt 0) {
            Install-Software -SoftwareList $optionalList -Type "Optional"
        }
        
        # Then upgrade
        Update-AllSoftware
    }
    default {
        Write-Host "Invalid choice. Exiting." -ForegroundColor $Colors.Error
        exit 1
    }
}

# Show summary
Show-InstallationSummary

Write-Host "W.A.R.P. Script completed!" -ForegroundColor $Colors.Success
Write-Host ""
