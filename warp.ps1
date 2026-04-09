<#
.SYNOPSIS
    W.A.R.P. - Winget Application Rollout Platform
.DESCRIPTION
    Automated software installation and management
.NOTES
    Requires Administrator privileges
    Requires Windows Package Manager (winget)
    Compatible with PowerShell 5.1+
#>

# ===========================
# CONFIGURATION
# ===========================

$ScriptPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = Get-Location
}

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

$InstallationLog = New-Object System.Collections.ArrayList

function Add-InstallationRecord {
    param(
        [string]$Software,
        [string]$Status,
        [string]$Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    )
    
    $record = New-Object PSObject -Property @{
        Timestamp = $Timestamp
        Software  = $Software
        Status    = $Status
    }
    
    [void]$InstallationLog.Add($record)
}

# ===========================
# WINGET CHECK
# ===========================

function Test-WingetAvailable {
    try {
        $wingetVersion = & winget --version 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Winget is available. Version: $wingetVersion" -ForegroundColor $Colors.Success
            return $true
        }
    }
    catch {
        # Winget not found
    }
    
    Write-Host "[!!] Winget is not installed. Installing now..." -ForegroundColor $Colors.Warning
    
    try {
        $progressPreference = 'SilentlyContinue'
        
        # Download and execute Winget installer
        $wingetUrl = "https://aka.ms/getwinget"
        $tempFile = Join-Path $env:TEMP "GetWinget.ps1"
        
        (New-Object System.Net.WebClient).DownloadFile($wingetUrl, $tempFile)
        
        if (Test-Path $tempFile) {
            & $tempFile
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            
            Write-Host "[OK] Winget installed successfully" -ForegroundColor $Colors.Success
            return $true
        }
    }
    catch {
        Write-Host "[ERROR] Failed to install Winget: $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }
    
    return $false
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
        Write-Host "[*] Installing: $item..." -ForegroundColor $Colors.Info
        
        try {
            $startTime = Get-Date
            $output = & winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1
            $exitCode = $LASTEXITCODE
            $installTime = (Get-Date).ToString('HH:mm:ss')
            
            if ($exitCode -eq 0 -or $exitCode -eq 931 -or $exitCode -eq 3010) {
                Write-Host "[OK] $item installed successfully at $installTime" -ForegroundColor $Colors.Success
                Add-InstallationRecord -Software $item -Status "INSTALLED"
            }
            else {
                Write-Host "[!!] $item installation completed with status code $exitCode at $installTime" -ForegroundColor $Colors.Warning
                Add-InstallationRecord -Software $item -Status "INSTALLED (with warnings - Exit Code: $exitCode)"
            }
        }
        catch {
            Write-Host "[ERROR] Error installing $item : $($_.Exception.Message)" -ForegroundColor $Colors.Error
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
        Write-Host "[*] Running: winget upgrade --all" -ForegroundColor $Colors.Info
        $output = & winget upgrade --all --accept-source-agreements --accept-package-agreements 2>&1
        $exitCode = $LASTEXITCODE
        $updateTime = (Get-Date).ToString('HH:mm:ss')
        
        if ($exitCode -eq 0 -or $exitCode -eq 931) {
            Write-Host "[OK] Package update completed at $updateTime" -ForegroundColor $Colors.Success
            Add-InstallationRecord -Software "All Packages" -Status "UPDATED"
        }
        else {
            Write-Host "[!!] Package update completed with status code $exitCode at $updateTime" -ForegroundColor $Colors.Warning
            Add-InstallationRecord -Software "All Packages" -Status "UPDATED (with warnings)"
        }
    }
    catch {
        Write-Host "[ERROR] Error during update: $($_.Exception.Message)" -ForegroundColor $Colors.Error
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
    $i = 0
    foreach ($software in $OptionalSoftware) {
        $i++
        Write-Host "  $i. $software" -ForegroundColor $Colors.Info
    }
    
    Write-Host ""
    Write-Host "Options:" -ForegroundColor $Colors.Header
    Write-Host "  [A] Install all optional software" -ForegroundColor $Colors.Info
    Write-Host "  [S] Select specific software (e.g., 1,2,3)" -ForegroundColor $Colors.Info
    Write-Host "  [N] Skip optional software" -ForegroundColor $Colors.Info
    Write-Host ""
    
    $choice = Read-Host "Enter your choice (A/S/N)"
    
    $selected = New-Object System.Collections.ArrayList
    
    switch ($choice.ToUpper()) {
        "A" {
            foreach ($software in $OptionalSoftware) {
                [void]$selected.Add($software)
            }
            Write-Host "[OK] Selected all optional software" -ForegroundColor $Colors.Success
        }
        "S" {
            $input = Read-Host "Enter numbers (comma-separated, e.g., 1,2)"
            
            if (-not [string]::IsNullOrWhiteSpace($input)) {
                $numbers = $input -split ','
                foreach ($num in $numbers) {
                    $trimmed = $num.Trim()
                    if ([int]::TryParse($trimmed, [ref]$null)) {
                        $index = [int]$trimmed - 1
                        if ($index -ge 0 -and $index -lt $OptionalSoftware.Count) {
                            [void]$selected.Add($OptionalSoftware[$index])
                        }
                    }
                }
            }
            
            if ($selected.Count -gt 0) {
                $selectedList = $selected -join ', '
                Write-Host "[OK] Selected: $selectedList" -ForegroundColor $Colors.Success
            }
        }
        "N" {
            Write-Host "[!!] Skipping optional software" -ForegroundColor $Colors.Warning
        }
        default {
            Write-Host "[!!] Invalid choice. Skipping optional software" -ForegroundColor $Colors.Warning
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
        $InstallationLog | Select-Object Timestamp, Software, Status | Format-Table -AutoSize
    }
    else {
        Write-Host "[!!] No installations recorded" -ForegroundColor $Colors.Warning
    }
    
    Write-Host ""
    $successCount = ($InstallationLog | Where-Object { $_.Status -like "*INSTALLED*" } | Measure-Object).Count
    $failCount = ($InstallationLog | Where-Object { $_.Status -like "*FAILED*" } | Measure-Object).Count
    $updateCount = ($InstallationLog | Where-Object { $_.Status -like "*UPDATED*" } | Measure-Object).Count
    
    Write-Host "Summary: Installed: $successCount | Failed: $failCount | Updated: $updateCount" -ForegroundColor $Colors.Header
    Write-Host ""
}

# ===========================
# MAIN EXECUTION
# ===========================

Show-Banner

# Check Winget
if (-not (Test-WingetAvailable)) {
    Write-Host "[ERROR] Cannot proceed without Winget" -ForegroundColor $Colors.Error
    Read-Host "Press Enter to exit"
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
        Write-Host "[OK] Selected: Install Mode" -ForegroundColor $Colors.Success
        
        # Install required software
        Install-Software -SoftwareList $RequiredSoftware -Type "Required"
        
        # Ask for optional software
        $optionalList = Select-OptionalSoftware
        
        if ($optionalList.Count -gt 0) {
            Install-Software -SoftwareList $optionalList -Type "Optional"
        }
    }
    "2" {
        Write-Host "[OK] Selected: Upgrade Mode" -ForegroundColor $Colors.Success
        Update-AllSoftware
    }
    "3" {
        Write-Host "[OK] Selected: Install and Upgrade Mode" -ForegroundColor $Colors.Success
        
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
        Write-Host "[ERROR] Invalid choice. Exiting." -ForegroundColor $Colors.Error
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Show summary
Show-InstallationSummary

Write-Host "[OK] W.A.R.P. Script completed!" -ForegroundColor $Colors.Success
Write-Host ""
Read-Host "Press Enter to exit"
