<#
.SYNOPSIS
    C.O.N.J.U.R.E. — Centrally Orchestrates Network-Joined Updates, Rollouts & Executables
    Automated Software Deployment Tool for PowerShell 5.1+
.DESCRIPTION
    Automated software installation and management
.NOTES
    Requires Administrator privileges
    Requires Windows Package Manager (winget)
    Compatible with PowerShell 5.1+

    Part of the toolbox alongside:
    R.U.N.E.P.R.E.S.S.  — Remote Utility for Networked Equipment — Printer Registration, Extraction & Silent Setup
    R.E.S.T.O.R.A.T.I.O.N. — Renews Every System Through Orderly Rite — Automating The Installation Of New updates
    O.R.A.C.L.E.        — Observes, Reports & Audits Computer Logs & Environments
    C.O.V.E.N.A.N.T.    — Configures Onboarding Via Entra — Network, Accounts, Naming & Timezone
#>

# ===========================
# CONFIGURATION
# ===========================

$ScriptPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = Get-Location
}

$ExecutionTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

# Set console to UTF-8 so Unicode block characters render correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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

# Chocolatey package IDs (mapped to the same software as above)
$RequiredSoftwareChoco = @(
    "microsoft-teams",
    "microsoft365apps",
    "7zip",
    "googlechrome",
    "adobereader",
    "zoom"
)

$OptionalSoftwareChoco = @(
    "zoom-outlook",
    "firefox",
    "dell-command-update"
)

$PackageManager = "winget"

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

   ██████╗  ██████╗ ███╗   ██╗     ██╗██╗   ██╗██████╗ ███████╗
  ██╔════╝ ██╔═══██╗████╗  ██║     ██║██║   ██║██╔══██╗██╔════╝
  ██║      ██║   ██║██╔██╗ ██║     ██║██║   ██║██████╔╝█████╗
  ██║      ██║   ██║██║╚██╗██║██   ██║██║   ██║██╔══██╗██╔══╝
  ╚██████╗ ╚██████╔╝██║ ╚████║╚█████╔╝╚██████╔╝██║  ██║███████╗
   ╚═════╝  ╚═════╝ ╚═╝  ╚═══╝ ╚════╝  ╚═════╝ ╚═╝  ╚═╝╚══════╝

"@ -ForegroundColor Cyan

    Write-Host "    C.O.N.J.U.R.E. — Centrally Orchestrates Network-Joined Updates, Rollouts & Executables" -ForegroundColor Cyan
    Write-Host "    Automated Software Deployment Tool" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Script Location: $ScriptPath" -ForegroundColor Gray
    Write-Host "    Execution Time:  $ExecutionTime" -ForegroundColor Gray
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
# PACKAGE MANAGER CHECK
# ===========================

function Select-PackageManager {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "SELECT PACKAGE MANAGER" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    Write-Host "  [1] Winget (Windows Package Manager)" -ForegroundColor $Colors.Info
    Write-Host "  [2] Chocolatey" -ForegroundColor $Colors.Info
    Write-Host ""

    $pmChoice = Read-Host "Enter your choice (1/2)"

    switch ($pmChoice) {
        "1" {
            Write-Host "[OK] Using Winget" -ForegroundColor $Colors.Success
            $script:PackageManager = "winget"
        }
        "2" {
            Write-Host "[OK] Using Chocolatey" -ForegroundColor $Colors.Success
            $script:PackageManager = "chocolatey"
        }
        default {
            Write-Host "[!!] Invalid choice. Defaulting to Winget" -ForegroundColor $Colors.Warning
            $script:PackageManager = "winget"
        }
    }
}

function Test-ChocolateyAvailable {
    try {
        $chocoVersion = & choco --version 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Chocolatey is available. Version: $chocoVersion" -ForegroundColor $Colors.Success
            return $true
        }
    }
    catch {
        # Chocolatey not found
    }

    Write-Host "[!!] Chocolatey is not installed. Installing now..." -ForegroundColor $Colors.Warning

    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

        # Refresh PATH so choco is available in this session
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("Path", "User")

        $chocoVersion = & choco --version 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Chocolatey installed successfully. Version: $chocoVersion" -ForegroundColor $Colors.Success
            return $true
        }
    }
    catch {
        Write-Host "[ERROR] Failed to install Chocolatey: $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }

    return $false
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

            if ($script:PackageManager -eq "chocolatey") {
                $output = & choco install $item -y 2>&1
            }
            else {
                $output = & winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1
            }

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
        if ($script:PackageManager -eq "chocolatey") {
            Write-Host "[*] Running: choco upgrade all" -ForegroundColor $Colors.Info
            $output = & choco upgrade all -y 2>&1
        }
        else {
            Write-Host "[*] Running: winget upgrade --all" -ForegroundColor $Colors.Info
            $output = & winget upgrade --all --accept-source-agreements --accept-package-agreements 2>&1
        }
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
    param(
        [string[]]$SoftwareList = $OptionalSoftware
    )

    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "OPTIONAL SOFTWARE SELECTION" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""

    Write-Host "Available optional software:" -ForegroundColor $Colors.Header
    $i = 0
    foreach ($software in $SoftwareList) {
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
            foreach ($software in $SoftwareList) {
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
                        if ($index -ge 0 -and $index -lt $SoftwareList.Count) {
                            [void]$selected.Add($SoftwareList[$index])
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

# Select package manager
Select-PackageManager

# Check that the selected package manager is available
if ($PackageManager -eq "chocolatey") {
    if (-not (Test-ChocolateyAvailable)) {
        Write-Host "[ERROR] Cannot proceed without Chocolatey" -ForegroundColor $Colors.Error
        Read-Host "Press Enter to exit"
        exit 1
    }
}
else {
    if (-not (Test-WingetAvailable)) {
        Write-Host "[ERROR] Cannot proceed without Winget" -ForegroundColor $Colors.Error
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Resolve the correct package lists based on selected manager
$ActiveRequired = if ($PackageManager -eq "chocolatey") { $RequiredSoftwareChoco } else { $RequiredSoftware }
$ActiveOptional = if ($PackageManager -eq "chocolatey") { $OptionalSoftwareChoco } else { $OptionalSoftware }

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
        Install-Software -SoftwareList $ActiveRequired -Type "Required"

        # Ask for optional software
        $optionalList = Select-OptionalSoftware -SoftwareList $ActiveOptional

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
        Install-Software -SoftwareList $ActiveRequired -Type "Required"

        # Ask for optional software
        $optionalList = Select-OptionalSoftware -SoftwareList $ActiveOptional

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

Write-Host "[OK] C.O.N.J.U.R.E. Script completed!" -ForegroundColor $Colors.Success
Write-Host ""
Read-Host "Press Enter to exit"
