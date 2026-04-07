# ─────────────────────────────────────────────────────────────────────────────
# S.P.A.R.K - Software Package & Resource Kit
# Version 6.2 - Corrected and Complete
# Automated Package Manager (Winget/Chocolatey)
# ─────────────────────────────────────────────────────────────────────────────

<#
.SYNOPSIS
    S.P.A.R.K - Automated Package Manager (Winget/Chocolatey)
.DESCRIPTION
    Installs/upgrades software packages unattended via Winget or Chocolatey fallback.
    Logs to CSV and Windows Event Log for compliance and audit trails.
.PARAMETERS
    Mode                        - Install or Upgrade
    InstallDellCommandUpdate    - Enable Dell Command Update
    InstallZoomOutlookPlugin    - Enable Zoom Outlook Plugin
    InstallDellCommand          - Enable Dell Command Suite
    LogPath                     - Custom path for CSV log file
.EXAMPLES
    PS C:\> .\spark.ps1 -Mode Install
    PS C:\> .\spark.ps1 -Mode Upgrade -InstallDellCommand
#>

param(
    [ValidateSet("Install", "Upgrade")]
    [string]$Mode,
    [switch]$InstallDellCommandUpdate,
    [switch]$InstallZoomOutlookPlugin,
    [switch]$InstallDellCommand,
    [string]$LogPath
)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURABLE PATHS
# ─────────────────────────────────────────────────────────────────────────────

$SPARKLogRoot   = "C:\Logs"
$DefaultLogFile = "install_log.csv"

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path $SPARKLogRoot -ChildPath $DefaultLogFile
}

# ─────────────────────────────────────────────────────────────────────────────
# EVENT LOG CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

$script:EventLogSource = "S.P.A.R.K"
$script:EventLogName   = "Application"

$script:EventIds = @{
    ScriptStart           = 1000
    ScriptEnd             = 1001
    ModeSelected          = 1002
    WingetDetected        = 2000
    WingetInstallAttempt  = 2001
    WingetInstallSuccess  = 2002
    WingetInstallFailed   = 2003
    ChocoDetected         = 2100
    ChocoInstallAttempt   = 2101
    ChocoInstallSuccess   = 2102
    ChocoInstallFailed    = 2103
    PackageInstallStart   = 3000
    PackageInstallSuccess = 3001
    PackageInstallFailed  = 3002
    PackageSkipped        = 3003
    LogExportSuccess      = 4000
    LogExportFailed       = 4001
    SummaryReport         = 5000
}

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:ChocoInitialized  = $false
$script:OperationMode     = $Mode
$script:LogPath           = $LogPath
$script:LogRoot           = $SPARKLogRoot
$script:InstallLog        = @()
$script:WingetAvailable   = $false
$script:ChocoAvailable    = $false

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Show-SparkBanner {
    <#
    .SYNOPSIS
        Displays the S.P.A.R.K ASCII banner
    #>
    Write-Host @"

  ███████╗██████╗  █████╗ ██████╗ ██╗  ██╗
  ██╔════╝██╔══██╗██╔══██╗██╔══██╗██║ ██╔╝
  ███████╗██████╔╝███████║██████╔╝█████╔╝ 
  ╚════██║██╔═══╝ ██╔══██║██╔══██╗██╔═██╗ 
  ███████║██║     ██║  ██║██║  ██║██║  ██╗
  ╚══════╝╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝
"@ -ForegroundColor Yellow
    Write-Host "    Software Package & Resource Kit" -ForegroundColor Yellow
    Write-Host "    Automated Package Manager Setup & Installation" -ForegroundColor Yellow
    Write-Host ""
}

function Write-EventLog {
    <#
    .SYNOPSIS
        Writes an event to the Application event log
    #>
    param(
        [string]$Message,
        [string]$EventType = "Information",
        [int]$EventId = 0
    )
    
    try {
        $eventLog = New-Object System.Diagnostics.EventLog($script:EventLogName)
        $eventLog.Source = $script:EventLogSource
        
        switch ($EventType) {
            "Error"       { $entryType = [System.Diagnostics.EventLogEntryType]::Error }
            "Warning"     { $entryType = [System.Diagnostics.EventLogEntryType]::Warning }
            "Information" { $entryType = [System.Diagnostics.EventLogEntryType]::Information }
            default       { $entryType = [System.Diagnostics.EventLogEntryType]::Information }
        }
        
        $eventLog.WriteEntry($Message, $entryType, $EventId)
    }
    catch {
        Write-Host "Warning: Event log write failed - $_" -ForegroundColor Yellow
    }
}

function Initialize-EventLog {
    <#
    .SYNOPSIS
        Creates event log source if it doesn't exist
    #>
    try {
        $sourceExists = [System.Diagnostics.EventLog]::SourceExists($script:EventLogSource)
        
        if (-not $sourceExists) {
            [System.Diagnostics.EventLog]::CreateEventSource($script:EventLogSource, $script:EventLogName)
            Write-Host "✓ Event log source created: $($script:EventLogSource)" -ForegroundColor Green
        }
        else {
            Write-Host "✓ Event log source already exists: $($script:EventLogSource)" -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Host "⚠ Warning: Could not initialize event log - $_" -ForegroundColor Yellow
        return $false
    }
}

function Update-EnvironmentPath {
    <#
    .SYNOPSIS
        Refreshes PATH from registry
    #>
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + 
                [System.Environment]::GetEnvironmentVariable("Path", "User")
}

function Initialize-Winget {
    <#
    .SYNOPSIS
        Detects Winget or installs from GitHub
    #>
    try {
        $ver = winget --version 2>&1
        Write-Host "✓ Winget detected - Version: $ver" -ForegroundColor Green
        Write-EventLog -Message "Winget detected. Version: $ver" -EventType "Information" -EventId $script:EventIds.WingetDetected
        $script:WingetAvailable = $true
        return $true
    }
    catch {
        Write-Host "⚠ Winget not found. Attempting installation from GitHub..." -ForegroundColor Yellow
        Write-EventLog -Message "Winget not found. Attempting installation." -EventType "Information" -EventId $script:EventIds.WingetInstallAttempt
        
        try {
            $release = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing
            $msixBundle = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1

            if (-not $msixBundle) {
                Write-Host "✗ Failed: Could not locate winget MSIX bundle" -ForegroundColor Red
                Write-EventLog -Message "Failed to locate winget MSIX bundle" -EventType "Error" -EventId $script:EventIds.WingetInstallFailed
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
            return $true
        }
        catch {
            Write-Host "✗ Winget installation failed: $_" -ForegroundColor Red
            Write-EventLog -Message "Winget installation failed: $_" -EventType "Error" -EventId $script:EventIds.WingetInstallFailed
            return $false
        }
    }
}

function Initialize-Chocolatey {
    <#
    .SYNOPSIS
        Detects Chocolatey or installs it
    #>
    if ($script:ChocoInitialized) {
        return $script:ChocoAvailable
    }

    try {
        $version = choco --version 2>&1
        Write-Host "✓ Chocolatey detected - Version: $version" -ForegroundColor Green
        Write-EventLog -Message "Chocolatey detected. Version: $version" -EventType "Information" -EventId $script:EventIds.ChocoDetected
        $script:ChocoAvailable = $true
        $script:ChocoInitialized = $true
        return $true
    }
    catch {
        Write-Host "⚠ Chocolatey not found. Installing..." -ForegroundColor Yellow
        Write-EventLog -Message "Chocolatey not found. Attempting installation." -EventType "Information" -EventId $script:EventIds.ChocoInstallAttempt
        
        try {
            $installScript = Invoke-WebRequest -Uri "https://community.chocolatey.org/install.ps1" -UseBasicParsing | Select-Object -ExpandProperty Content
            Invoke-Expression $installScript
            Update-EnvironmentPath
            
            $version = choco --version 2>&1
            Write-Host "✓ Chocolatey installed - Version: $version" -ForegroundColor Green
            Write-EventLog -Message "Chocolatey installed successfully. Version: $version" -EventType "Information" -EventId $script:EventIds.ChocoInstallSuccess
            $script:ChocoAvailable = $true
            $script:ChocoInitialized = $true
            return $true
        }
        catch {
            Write-Host "✗ Chocolatey installation failed: $_" -ForegroundColor Red
            Write-EventLog -Message "Chocolatey installation failed: $_" -EventType "Error" -EventId $script:EventIds.ChocoInstallFailed
            $script:ChocoAvailable = $false
            $script:ChocoInitialized = $true
            return $false
        }
    }
}

function Install-Package {
    <#
    .SYNOPSIS
        Installs or upgrades a single package
    #>
    param(
        [string]$PackageName,
        [string]$WingetId,
        [string]$ChocoId,
        [string]$Action = "Install"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Try Winget first
    if ($script:WingetAvailable) {
        try {
            Write-Host "  [$Action via Winget] $PackageName..." -ForegroundColor Cyan -NoNewline
            Write-EventLog -Message "Starting package $Action via Winget: $PackageName" -EventType "Information" -EventId $script:EventIds.PackageInstallStart
            
            if ($Action -eq "Install") {
                & winget install --id $WingetId --exact --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
            }
            else {
                & winget upgrade --id $WingetId --exact --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
            }
            
            $exitCode = $LASTEXITCODE
            
            if ($exitCode -eq 0 -or $exitCode -eq 1641 -or $exitCode -eq 3010) {
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
                Write-Host " ⚠ Exit Code: $exitCode" -ForegroundColor Yellow
                # Fall through to Chocolatey
            }
        }
        catch {
            Write-Host " ⚠ Winget error: $_" -ForegroundColor Yellow
            # Fall through to Chocolatey
        }
    }

    # Fall back to Chocolatey
    if ($script:ChocoAvailable -and $ChocoId) {
        try {
            Write-Host "  [$Action via Chocolatey] $PackageName..." -ForegroundColor Cyan -NoNewline
            Write-EventLog -Message "Falling back to Chocolatey for $PackageName" -EventType "Information" -EventId $script:EventIds.PackageInstallStart
            
            if ($Action -eq "Install") {
                & choco install $ChocoId -y --no-progress 2>&1 | Out-Null
            }
            else {
                & choco upgrade $ChocoId -y --no-progress 2>&1 | Out-Null
            }
            
            $exitCode = $LASTEXITCODE
            
            if ($exitCode -eq 0) {
                Write-Host " ✓ Success" -ForegroundColor Green
                $logEntry = @{
                    Timestamp = $timestamp
                    Package   = $PackageName
                    Manager   = "Chocolatey"
                    Status    = "Success"
                    ExitCode  = $exitCode
                }
                $script:InstallLog += New-Object PSObject -Property $logEntry
                Write-EventLog -Message "$Action succeeded for $PackageName via Chocolatey" -EventType "Information" -EventId $script:EventIds.PackageInstallSuccess
                return $true
            }
            else {
                Write-Host " ✗ Failed (Exit: $exitCode)" -ForegroundColor Red
                $logEntry = @{
                    Timestamp = $timestamp
                    Package   = $PackageName
                    Manager   = "Chocolatey"
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
                Manager   = "Chocolatey"
                Status    = "Failed"
                ExitCode  = -1
            }
            $script:InstallLog += New-Object PSObject -Property $logEntry
            Write-EventLog -Message "$Action error for $PackageName`: $errorMsg" -EventType "Error" -EventId $script:EventIds.PackageInstallFailed
            return $false
        }
    }

    # No package manager available
    Write-Host " ⊘ Skipped (no manager)" -ForegroundColor Yellow
    $logEntry = @{
        Timestamp = $timestamp
        Package   = $PackageName
        Manager   = "None"
        Status    = "Skipped"
        ExitCode  = -1
    }
    $script:InstallLog += New-Object PSObject -Property $logEntry
    Write-EventLog -Message "Package skipped - no manager available: $PackageName" -EventType "Warning" -EventId $script:EventIds.PackageSkipped
    return $false
}

function Show-SummaryReport {
    <#
    .SYNOPSIS
        Displays the installation summary
    #>
    param(
        [string]$ActionVerb = "Installation"
    )

    $totalCount   = $script:InstallLog.Count
    $successCount = ($script:InstallLog | Where-Object { $_.Status -eq "Success" }).Count
    $failureCount = ($script:InstallLog | Where-Object { $_.Status -eq "Failed" }).Count
    $skipCount    = ($script:InstallLog | Where-Object { $_.Status -eq "Skipped" }).Count

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "S.P.A.R.K $ActionVerb Summary Report" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Mode: $($script:OperationMode)" -ForegroundColor Cyan
    Write-Host "Execution Timestamp: $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "$ActionVerb Statistics:" -ForegroundColor Magenta
    Write-Host "├─ Total Packages : $totalCount"
    Write-Host "├─ Successful     : $successCount" -ForegroundColor Green
    Write-Host "├─ Failed         : $failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })
    Write-Host "└─ Skipped        : $skipCount" -ForegroundColor Yellow
    Write-Host ""

    if ($successCount -gt 0) {
        Write-Host "Successfully $($script:OperationMode.ToLower())ed:" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        $script:InstallLog | Where-Object { $_.Status -eq "Success" } | ForEach-Object {
            Write-Host "  [+] $($_.Package)" -ForegroundColor Green
            Write-Host "      Manager: $($_.Manager) | Exit: $($_.ExitCode) | Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    if ($failureCount -gt 0) {
        Write-Host "Failed Operations:" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        $script:InstallLog | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
            Write-Host "  [-] $($_.Package)" -ForegroundColor Red
            Write-Host "      Manager: $($_.Manager) | Exit: $($_.ExitCode) | Time: $($_.Timestamp)" -ForegroundColor Gray
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
    Write-Host ""

    $summaryMessage = "S.P.A.R.K $ActionVerb completed. Total: $totalCount | Success: $successCount | Failed: $failureCount | Skipped: $skipCount"
    $summaryEventType = if ($failureCount -gt 0) { "Warning" } else { "Information" }
    Write-EventLog -Message $summaryMessage -EventType $summaryEventType -EventId $script:EventIds.SummaryReport
}

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
    @{ Name = "Microsoft Edge";           Winget = "Microsoft.Edge";                    Chocolatey = "microsoft-edge"        },
    @{ Name = "7-Zip";                    Winget = "7zip.7zip";                         Chocolatey = "7zip"                  },
    @{ Name = "Adobe Acrobat Reader";     Winget = "Adobe.Acrobat.Reader.64-bit";       Chocolatey = "adobereader"           },
    @{ Name = "Zoom";                     Winget = "Zoom.Zoom";                         Chocolatey = "zoom"                  },
    @{ Name = "Microsoft Office 365";     Winget = "Microsoft.Office";                  Chocolatey = "office365business"     }
)

$optionalSoftware = @(
    @{ Name = "Dell Command Update";      Winget = "Dell.CommandUpdate";                Chocolatey = "dell-command-update";   Param = "InstallDellCommandUpdate" },
    @{ Name = "Zoom Outlook Plugin";      Winget = "Zoom.ZoomOutlookPlugin";            Chocolatey = $null;                  Param = "InstallZoomOutlookPlugin" },
    @{ Name = "Dell Command Suite";       Winget = "Dell.Command";                      Chocolatey = "dell-command-suite";   Param = "InstallDellCommand" }
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
Initialize-Chocolatey

if (-not $script:WingetAvailable -and -not $script:ChocoAvailable) {
    Write-Host "`n✗ ERROR: Neither Winget nor Chocolatey is available!" -ForegroundColor Red
    Write-EventLog -Message "Script terminated: No package managers available" -EventType "Error" -EventId $script:EventIds.ScriptEnd
    exit 1
}

Write-Host ""
Write-Host "Starting $($Mode.ToLower())ation..." -ForegroundColor Magenta
Write-Host ""

# Install core software
foreach ($package in $coreSoftware) {
    Install-Package -PackageName $package.Name -WingetId $package.Winget -ChocoId $package.Chocolatey -Action $Mode
    Start-Sleep -Milliseconds 500
}

# Install optional software
foreach ($package in $optionalSoftware) {
    $paramName = $package.Param
    if ((Get-Variable -Name $paramName -ValueOnly -ErrorAction SilentlyContinue) -eq $true) {
        Install-Package -PackageName $package.Name -WingetId $package.Winget -ChocoId $package.Chocolatey -Action $Mode
        Start-Sleep -Milliseconds 500
    }
}

# Generate summary and export logs
Write-Host ""
Show-SummaryReport -ActionVerb $Mode
Export-InstallLog

Write-EventLog -Message "S.P.A.R.K $Mode completed successfully" -EventType "Information" -EventId $script:EventIds.ScriptEnd
Write-Host "`n✓ Script execution completed!" -ForegroundColor Green
