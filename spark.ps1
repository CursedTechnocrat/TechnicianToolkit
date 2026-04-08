# ─────────────────────────────────────────────────────────────────────────────
# S.P.A.R.K - Software Package Auto-Installer
# v6.5 - Parsing errors
# ─────────────────────────────────────────────────────────────────────────────

param(
    [string]$Mode = "",
    [switch]$InstallDellCommandUpdate,
    [switch]$InstallZoomOutlookPlugin,
    [switch]$InstallDellCommand
)

# ─────────────────────────────────────────────────────────────────────────────
# INITIALIZE SCRIPT VARIABLES
# ─────────────────────────────────────────────────────────────────────────────

$script:EventLogName = "S.P.A.R.K"
$script:EventLogSource = "S.P.A.R.K-Installer"
$script:LogPath = "$env:ProgramData\S.P.A.R.K\InstallLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$script:WingetAvailable = $false
$script:WingetInitialized = $false
$script:OperationMode = ""
$script:InstallLog = @()

$script:EventIds = @{
    ScriptStart              = 1000
    ScriptEnd                = 1001
    ModeSelected             = 1002
    WingetDetected           = 2001
    WingetInstallAttempt     = 2002
    WingetInstallSuccess     = 2003
    WingetInstallFailed      = 2004
    PackageInstallStart      = 3001
    PackageInstallSuccess    = 3002
    PackageInstallFailed     = 3003
    PackageSkipped           = 3004
    LogExportSuccess         = 4001
    LogExportFailed          = 4002
    SummaryReport            = 5001
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Show Banner
# ─────────────────────────────────────────────────────────────────────────────

function Show-SparkBanner {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║     S.P.A.R.K - Software Installer    ║" -ForegroundColor Magenta
    Write-Host "║   Streamlined Package Auto-Run Kit     ║" -ForegroundColor Magenta
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Initialize Event Log
# ─────────────────────────────────────────────────────────────────────────────

function Initialize-EventLog {
    try {
        if (-not (Get-EventLog -LogName $script:EventLogName -ErrorAction SilentlyContinue)) {
            New-EventLog -LogName $script:EventLogName -Source $script:EventLogSource -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host "⚠ Warning: Could not initialize Event Log" -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Write Event Log
# ─────────────────────────────────────────────────────────────────────────────

function Write-EventLog {
    param(
        [string]$Message,
        [ValidateSet("Information", "Warning", "Error")]
        [string]$EventType = "Information",
        [int]$EventId = 1000
    )

    try {
        [System.Diagnostics.EventLog]::WriteEntry($script:EventLogSource, $Message, $EventType, $EventId)
    }
    catch {
        # Silently continue if Event Log fails
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Update Environment Path
# ─────────────────────────────────────────────────────────────────────────────

function Update-EnvironmentPath {
    <#
    .SYNOPSIS
        Refreshes PATH from registry
    #>
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + 
                [System.Environment]::GetEnvironmentVariable("Path", "User")
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Initialize Winget
# ─────────────────────────────────────────────────────────────────────────────

function Initialize-Winget {
    <#
    .SYNOPSIS
        Detects Winget or installs from GitHub
    #>
    if ($script:WingetInitialized) {
        return $script:WingetAvailable
    }

    $wingetFound = $false
    try {
        $ver = winget --version 2>&1
        Write-Host "✓ Winget detected - Version: $ver" -ForegroundColor Green
        Write-EventLog -Message "Winget detected. Version: $ver" -EventType "Information" -EventId $script:EventIds.WingetDetected
        $wingetFound = $true
    }
    catch {
        Write-Host "⚠ Winget not found. Attempting installation from GitHub..." -ForegroundColor Yellow
        Write-EventLog -Message "Winget not found. Attempting installation." -EventType "Information" -EventId $script:EventIds.WingetInstallAttempt
    }

    if ($wingetFound) {
        $script:WingetAvailable = $true
        $script:WingetInitialized = $true
        return $true
    }

    # Try to install Winget from GitHub
    try {
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing
        $msixBundle = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1

        if (-not $msixBundle) {
            Write-Host "✗ Failed: Could not locate winget MSIX bundle" -ForegroundColor Red
            Write-EventLog -Message "Failed to locate winget MSIX bundle" -EventType "Error" -EventId $script:EventIds.WingetInstallFailed
            $script:WingetAvailable = $false
            $script:WingetInitialized = $true
            return $false
        }

        $tmpPath = Join-Path -Path $env:TEMP -ChildPath "winget_install.msixbundle"
        Write-Host "  Downloading: $($msixBundle.browser_download_url)" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $msixBundle.browser_download_url -OutFile $tmpPath -UseBasicParsing
        
        Write-Host "  Installing MSIX package..." -ForegroundColor Cyan
        Add-AppxPackage -Path $tmpPath
        Start-Sleep -Seconds 3
        Update-EnvironmentPath
        
        # Verify installation
        $ver = winget --version 2>&1
        Write-Host "✓ Winget installed successfully - Version: $ver" -ForegroundColor Green
        Write-EventLog -Message "Winget installed successfully. Version: $ver" -EventType "Information" -EventId $script:EventIds.WingetInstallSuccess
        $script:WingetAvailable = $true
        $script:WingetInitialized = $true
        return $true
    }
    catch {
        Write-Host "✗ Winget installation failed: $_" -ForegroundColor Red
        Write-EventLog -Message "Winget installation failed: $_" -EventType "Error" -EventId $script:EventIds.WingetInstallFailed
        $script:WingetAvailable = $false
        $script:WingetInitialized = $true
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Install Package
# ─────────────────────────────────────────────────────────────────────────────

function Install-Package {
    <#
    .SYNOPSIS
        Installs or upgrades a single package using Winget
    #>
    param(
        [string]$PackageName,
        [string]$WingetId,
        [string]$Action = "Install"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if (-not $script:WingetAvailable) {
        Write-Host "  [SKIP] $PackageName" -ForegroundColor Yellow
        $logEntry = @{
            Timestamp = $timestamp
            Package   = $PackageName
            Manager   = "Winget"
            Status    = "Skipped"
            ExitCode  = -1
        }
        $script:InstallLog += New-Object PSObject -Property $logEntry
        Write-EventLog -Message "Skipped $PackageName - Winget unavailable" -EventType "Information" -EventId $script:EventIds.PackageSkipped
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($WingetId)) {
        Write-Host "  [SKIP] $PackageName (no Winget ID)" -ForegroundColor Yellow
        $logEntry = @{
            Timestamp = $timestamp
            Package   = $PackageName
            Manager   = "Winget"
            Status    = "Skipped"
            ExitCode  = -1
        }
        $script:InstallLog += New-Object PSObject -Property $logEntry
        Write-EventLog -Message "Skipped $PackageName - no Winget ID" -EventType "Information" -EventId $script:EventIds.PackageSkipped
        return $false
    }

    try {
        Write-Host "  [$Action via Winget] $PackageName..." -ForegroundColor Cyan -NoNewline
        Write-EventLog -Message "Installing $PackageName with Winget" -EventType "Information" -EventId $script:EventIds.PackageInstallStart
        
        if ($Action -eq "Install") {
            & winget install --id $WingetId -e -h --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
        }
        else {
            & winget upgrade --id $WingetId -e -h --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
        }
        
        $exitCode = $LASTEXITCODE
        
        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
            Write-Host " ✓ Success" -ForegroundColor Green
            $logEntry = @{
                Timestamp = $timestamp
                Package   = $PackageName
                Manager   = "Winget"
                Status    = "Success"
                ExitCode  = $exitCode
            }
            $script:InstallLog += New-Object PSObject -Property $logEntry
            Write-EventLog -Message "$Action succeeded for $PackageName via Winget" -EventType "Information" -EventId $script:EventIds.PackageInstallSuccess
            return $true
        }
        else {
            Write-Host " ✗ Failed (Exit: $exitCode)" -ForegroundColor Red
            $logEntry = @{
                Timestamp = $timestamp
                Package   = $PackageName
                Manager   = "Winget"
                Status    = "Failed"
                ExitCode  = $exitCode
            }
            $script:InstallLog += New-Object PSObject -Property $logEntry
            Write-EventLog -Message "$Action failed for $PackageName. Exit Code: $exitCode" -EventType "Error" -EventId $script:EventIds.PackageInstallFailed
            return $false
        }
    }
    catch {
        Write-Host " ✗ Error: $_" -ForegroundColor Red
        $errorMsg = $_
        $logEntry = @{
            Timestamp = $timestamp
            Package   = $PackageName
            Manager   = "Winget"
            Status    = "Failed"
            ExitCode  = -1
        }
        $script:InstallLog += New-Object PSObject -Property $logEntry
        Write-EventLog -Message "$Action error for $PackageName`: $errorMsg" -EventType "Error" -EventId $script:EventIds.PackageInstallFailed
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Export Install Log
# ─────────────────────────────────────────────────────────────────────────────

function Export-InstallLog {
    <#
    .SYNOPSIS
        Exports installation log to CSV
    #>
    try {
        $logDir = Split-Path -Path $script:LogPath
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            Write-Host "✓ Created log directory: $logDir" -ForegroundColor Green
        }

        $script:InstallLog | Export-Csv -Path $script:LogPath -NoTypeInformation -Force
        Write-Host "✓ Log exported to: $($script:LogPath)" -ForegroundColor Green
        Write-EventLog -Message "Installation log exported: $($script:LogPath)" -EventType "Information" -EventId $script:EventIds.LogExportSuccess
        return $true
    }
    catch {
        Write-Host "✗ Failed to export log: $_" -ForegroundColor Red
        Write-EventLog -Message "Failed to export log: $_" -EventType "Error" -EventId $script:EventIds.LogExportFailed
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SOFTWARE DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

$coreSoftware = @(
    @{ Name = "Microsoft Edge";           Winget = "Microsoft.Edge" },
    @{ Name = "7-Zip";                    Winget = "7zip.7zip" },
    @{ Name = "Adobe Acrobat Reader";     Winget = "Adobe.Acrobat.Reader.64-bit" },
    @{ Name = "Zoom";                     Winget = "Zoom.Zoom" },
    @{ Name = "Microsoft Office 365";     Winget = "Microsoft.Office" }
)

$optionalSoftware = @(
    @{ Name = "Dell Command Update";      Winget = "Dell.CommandUpdate";                Param = "InstallDellCommandUpdate" },
    @{ Name = "Zoom Outlook Plugin";      Winget = "Zoom.ZoomOutlookPlugin";            Param = "InstallZoomOutlookPlugin" },
    @{ Name = "Dell Command Suite";       Winget = "Dell.CommandUpdate";                Param = "InstallDellCommand" }
)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

Show-SparkBanner

# Initialize logging
Initialize-EventLog
Write-EventLog -Message "S.P.A.R.K started. Mode: $Mode | Log Path: $($script:LogPath)" -EventType "Information" -EventId $script:EventIds.ScriptStart

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Initializing S.P.A.R.K..." -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

# Mode selection
if ([string]::IsNullOrWhiteSpace($Mode)) {
    Write-Host "Select operation mode:" -ForegroundColor Cyan
    Write-Host "1 = Install (new packages)"
    Write-Host "2 = Upgrade (existing packages)"
    Write-Host ""
    $choice = Read-Host "Enter choice (1 or 2)"
    
    if ($choice -eq "1") {
        $Mode = "Install"
    }
    elseif ($choice -eq "2") {
        $Mode = "Upgrade"
    }
    else {
        Write-Host "Invalid choice. Exiting." -ForegroundColor Red
        Write-EventLog -Message "Script terminated: Invalid mode selection" -EventType "Error" -EventId $script:EventIds.ScriptEnd
        exit 1
    }
}

$script:OperationMode = $Mode
Write-Host "✓ Mode selected: $Mode" -ForegroundColor Green
Write-EventLog -Message "Operation mode: $Mode" -EventType "Information" -EventId $script:EventIds.ModeSelected
Write-Host ""

# Initialize package managers
Write-Host "Initializing package managers..." -ForegroundColor Magenta
Initialize-Winget

if (-not $script:WingetAvailable) {
    Write-Host "`n✗ ERROR: Winget is not available!" -ForegroundColor Red
    Write-EventLog -Message "Script terminated: Winget unavailable" -EventType "Error" -EventId $script:EventIds.ScriptEnd
    exit 1
}

Write-Host ""
Write-Host "Starting $($Mode.ToLower())ation..." -ForegroundColor Magenta
Write-Host ""

# Install core software
foreach ($package in $coreSoftware) {
    Install-Package -PackageName $package.Name -WingetId $package.Winget -Action $Mode
    Start-Sleep -Milliseconds 500
}

Write-Host ""
Write-Host "Optional Software:" -ForegroundColor Magenta
Write-Host ""

# Install optional software
foreach ($package in $optionalSoftware) {
    $paramName = $package.Param
    if ((Get-Variable -Name $paramName -ValueOnly -ErrorAction SilentlyContinue) -eq $true) {
        Install-Package -PackageName $package.Name -WingetId $package.Winget -Action $Mode
        Start-Sleep -Milliseconds 500
    }
    else {
        Write-Host "  [SKIP] $($package.Name) (not selected)" -ForegroundColor Yellow
    }
}

# Generate summary report
Write-Host ""
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Installation Summary" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

$successCount = ($script:InstallLog | Where-Object { $_.Status -eq "Success" }).Count
$failureCount = ($script:InstallLog | Where-Object { $_.Status -eq "Failed" }).Count
$skipCount = ($script:InstallLog | Where-Object { $_.Status -eq "Skipped" }).Count
$totalCount = $script:InstallLog.Count

if ($successCount -gt 0) {
    Write-Host "Successful Installations:" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    $script:InstallLog | Where-Object { $_.Status -eq "Success" } | ForEach-Object {
        Write-Host "  [✓] $($_.Package)" -ForegroundColor Green
        Write-Host "      Manager: $($_.Manager) | Time: $($_.Timestamp)" -ForegroundColor Gray
    }
    Write-Host ""
}

if ($failureCount -gt 0) {
    Write-Host "Failed Installations:" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    $script:InstallLog | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
        Write-Host "  [✗] $($_.Package)" -ForegroundColor Red
        Write-Host "      Manager: $($_.Manager) | Exit Code: $($_.ExitCode) | Time: $($_.Timestamp)" -ForegroundColor Gray
    }
    Write-Host ""
}

if ($skipCount -gt 0) {
    Write-Host "Skipped Operations:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    $script:InstallLog | Where-Object { $_.Status -eq "Skipped" } | ForEach-Object {
        Write-Host "  [~] $($_.Package)" -ForegroundColor Yellow
        Write-Host "      Manager: $($_.Manager) | Time: $($_.Timestamp)" -ForegroundColor Gray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Completed at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Total: $totalCount | Success: $successCount | Failed: $failureCount | Skipped: $skipCount" -ForegroundColor Cyan
Write-Host ""

# Export log
Export-InstallLog

# Summary message and event log
$summaryMessage = "S.P.A.R.K $Mode completed. Total: $totalCount | Success: $successCount | Failed: $failureCount | Skipped: $skipCount"
if ($failureCount -gt 0) {
    $summaryEventType = "Warning"
}
else {
    $summaryEventType = "Information"
}

Write-EventLog -Message $summaryMessage -EventType $summaryEventType -EventId $script:EventIds.SummaryReport
Write-EventLog -Message "S.P.A.R.K script ended" -EventType "Information" -EventId $script:EventIds.ScriptEnd
