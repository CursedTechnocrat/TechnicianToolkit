<#
.VERSION
    6.0 Rebuilt to resolve errors
.SYNOPSIS
    S.P.A.R.K - Automated Package Manager (Winget/Chocolatey)
.DESCRIPTION
    Installs/upgrades software packages unattended via Winget or Chocolatey fallback.
    Logs to CSV and Windows Event Log for compliance and audit trails.
    Supports RMM/Kaseya/Task Scheduler deployment scenarios.
.PARAMETERS
    Mode                        - Install or Upgrade (default: interactive prompt)
    InstallDellCommandUpdate    - Enable Dell Command Update installation
    InstallZoomOutlookPlugin    - Enable Zoom Outlook Plugin installation
    InstallDellCommand          - Enable Dell Command Suite installation
    LogPath                     - Custom path for CSV log file
.EXAMPLES
    PS C:\> .\spark.ps1 -Mode Install
    PS C:\> .\spark.ps1 -Mode Upgrade -InstallDellCommand
    PS C:\> .\spark.ps1 -Mode Install -LogPath "C:\Custom\Logs\install.csv"
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
# CONFIGURABLE PATHS - EDIT HERE FOR CUSTOM LOCATIONS
# ─────────────────────────────────────────────────────────────────────────────

$SPARKLogRoot   = "C:\Logs"
$DefaultLogFile = "install_log.csv"

# Build the full log path if not provided via parameter
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path $SPARKLogRoot -ChildPath $DefaultLogFile
}

# ─────────────────────────────────────────────────────────────────────────────
# EVENT LOG CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

$script:EventLogSource = "S.P.A.R.K"
$script:EventLogName   = "Application"

# Event ID definitions for SOC compliance and audit trails
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

$script:ChocoInitialized = $false
$script:OperationMode    = $Mode
$script:LogPath          = $LogPath
$script:LogRoot          = $SPARKLogRoot
$installationLog         = @()

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
        Writes an event to the Application event log for S.P.A.R.K operations.
        Uses standard event types: Information, Warning, Error
    #>
    param(
        [string]$Message,
        [string]$EventType = "Information",
        [int]$EventId = 0
    )
    
    try {
        $eventLog = New-Object System.Diagnostics.EventLog($script:EventLogName)
        $eventLog.Source = $script:EventLogSource
        
        # Map event type strings to EventLogEntryType enum
        switch ($EventType) {
            "Error"       { $entryType = [System.Diagnostics.EventLogEntryType]::Error }
            "Warning"     { $entryType = [System.Diagnostics.EventLogEntryType]::Warning }
            "Information" { $entryType = [System.Diagnostics.EventLogEntryType]::Information }
            default       { $entryType = [System.Diagnostics.EventLogEntryType]::Information }
        }
        
        $eventLog.WriteEntry($Message, $entryType, $EventId)
    }
    catch {
        # Silent fail — don't interrupt script flow if event logging fails
    }
}

function Initialize-EventLog {
    <#
    .SYNOPSIS
        Initializes the event log source for S.P.A.R.K if it doesn't already exist.
        Logs to the standard Application event log to maintain SOC compliance.
    #>
    try {
        $sourceExists = [System.Diagnostics.EventLog]::SourceExists($script:EventLogSource)
        
        if (-not $sourceExists) {
            [System.Diagnostics.EventLog]::CreateEventSource($script:EventLogSource, $script:EventLogName)
            Write-Host "Event log source '$script:EventLogSource' created in Application log." -ForegroundColor Green
        }
        else {
            Write-Host "Event log source '$script:EventLogSource' already exists." -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Host "Warning: Could not initialize event log source: $_" -ForegroundColor Yellow
        return $false
    }
}

function Update-EnvironmentPath {
    <#
    .SYNOPSIS
        Refreshes the current session's PATH environment variable from system and user registry.
    #>
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
}

function Initialize-Winget {
    <#
    .SYNOPSIS
        Detects Winget or installs it from GitHub releases if unavailable.
    #>
    try {
        $ver = winget --version 2>&1
        Write-Host "Winget already installed. Version: $ver" -ForegroundColor Green
        Write-EventLog -Message "Winget detected. Version: $ver" -EventType "Information" -EventId $script:EventIds.WingetDetected
        return $true
    }
    catch {
        Write-Host "Winget not found. Attempting install via GitHub release..." -ForegroundColor Yellow
        Write-EventLog -Message "Winget not found. Attempting installation from GitHub." -EventType "Information" -EventId $script:EventIds.WingetInstallAttempt
        
        try {
            $release    = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing
            $msixBundle = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1

            if (-not $msixBundle) {
                Write-Host "Could not locate winget MSIX bundle in latest GitHub release." -ForegroundColor Red
                Write-EventLog -Message "Failed to locate winget MSIX bundle in GitHub release." -EventType "Error" -EventId $script:EventIds.WingetInstallFailed
                return $false
            }

            $tmpPath = "$env:TEMP\winget_install.msixbundle"
            Invoke-WebRequest -Uri $msixBundle.browser_download_url -OutFile $tmpPath -UseBasicParsing
            Add-AppxPackage -Path $tmpPath
            Start-Sleep -Seconds 3
            Update-EnvironmentPath

            $ver = winget --version 2>&1
            Write-Host "Winget installed successfully. Version: $ver" -ForegroundColor Green
            Write-EventLog -Message "Winget installed successfully. Version: $ver" -EventType "Information" -EventId $script:EventIds.WingetInstallSuccess
            return $true
        }
        catch {
            Write-Host "Failed to install Winget: $_" -ForegroundColor Red
            Write-EventLog -Message "Failed to install Winget. Error: $_" -EventType "Error" -EventId $script:EventIds.WingetInstallFailed
            return $false
        }
    }
}

function Initialize-Chocolatey {
    <#
    .SYNOPSIS
        Detects Chocolatey or installs it from the official install script if unavailable.
    #>
    if ($script:ChocoInitialized) {
        return $true
    }

    try {
        $ver = choco --version 2>&1
        Write-Host "Chocolatey already installed. Version: $ver" -ForegroundColor Green
        Write-EventLog -Message "Chocolatey detected. Version: $ver" -EventType "Information" -EventId $script:EventIds.ChocoDetected
        $script:ChocoInitialized = $true
        return $true
    }
    catch {
        if (Test-Path "C:\ProgramData\chocolatey") {
            Write-Host "Chocolatey found but not in PATH. Refreshing..." -ForegroundColor Yellow
            Update-EnvironmentPath
            Start-Sleep -Seconds 1
            try {
                $ver = choco --version 2>&1
                Write-Host "Chocolatey accessible. Version: $ver" -ForegroundColor Green
                Write-EventLog -Message "Chocolatey made accessible via PATH refresh. Version: $ver" -EventType "Information" -EventId $script:EventIds.ChocoDetected
                $script:ChocoInitialized = $true
                return $true
            }
            catch {
                Write-Host "Chocolatey present but inaccessible. A session restart may be required." -ForegroundColor Yellow
                Write-EventLog -Message "Chocolatey present but inaccessible. Session restart may be required." -EventType "Warning" -EventId $script:EventIds.ChocoInstallFailed
                return $false
            }
        }
        else {
            Write-Host "Chocolatey not found. Installing..." -ForegroundColor Yellow
            Write-EventLog -Message "Chocolatey not found. Attempting installation." -EventType "Information" -EventId $script:EventIds.ChocoInstallAttempt
            
            try {
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

                $chocoScript = (New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')
                Invoke-Expression $chocoScript

                Start-Sleep -Seconds 2
                Update-EnvironmentPath
                Write-Host "Chocolatey installed successfully." -ForegroundColor Green
                Write-EventLog -Message "Chocolatey installed successfully." -EventType "Information" -EventId $script:EventIds.ChocoInstallSuccess
                $script:ChocoInitialized = $true
                return $true
            }
            catch {
                Write-Host "Failed to install Chocolatey: $_" -ForegroundColor Red
                Write-EventLog -Message "Failed to install Chocolatey. Error: $_" -EventType "Error" -EventId $script:EventIds.ChocoInstallFailed
                return $false
            }
        }
    }
}

function Install-PackageViaWinget {
    <#
    .SYNOPSIS
        Attempts to install or upgrade a package using Winget.
        Falls back to Chocolatey if Winget fails.
    #>
    param(
        [string]$PackageId,
        [ref]$LogArray
    )
    
    $action = if ($script:OperationMode -eq "Upgrade") { "Upgrading" } else { "Installing" }
    Write-Host "  Attempting Winget: $action $PackageId..." -ForegroundColor Yellow
    Write-EventLog -Message "Attempting to $($action.ToLower()) $PackageId via Winget." -EventType "Information" -EventId $script:EventIds.PackageInstallStart
    
    try {
        if ($script:OperationMode -eq "Upgrade") {
            winget upgrade -e --id $PackageId --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
        }
        else {
            winget install -e --id $PackageId --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
        }
        
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0 -or $exitCode -eq -1978335189) {
            Write-Host "  [+] $PackageId $action successfully via Winget." -ForegroundColor Green
            Write-EventLog -Message "$action $PackageId successfully via Winget." -EventType "Information" -EventId $script:EventIds.PackageInstallSuccess
            $LogArray.Value += [PSCustomObject]@{
                Package   = $PackageId
                Status    = "Success"
                Manager   = "Winget"
                Mode      = $script:OperationMode
                ExitCode  = $exitCode
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            return $true
        }
        else {
            Write-Host "  [-] Winget failed for $PackageId (Exit: $exitCode). Will retry with Chocolatey..." -ForegroundColor Red
            Write-EventLog -Message "Winget failed for $PackageId (Exit Code: $exitCode). Retrying with Chocolatey." -EventType "Warning" -EventId $script:EventIds.PackageInstallFailed
            return $false
        }
    }
    catch {
        Write-Host "  [-] Winget error for $PackageId`: \$_. Will retry with Chocolatey..." -ForegroundColor Red
        Write-EventLog -Message "Winget error for \$PackageId`: $_" -EventType "Warning" -EventId $script:EventIds.PackageInstallFailed
        return $false
    }
}

function Install-PackageViaChocolatey {
    <#
    .SYNOPSIS
        Attempts to install or upgrade a package using Chocolatey.
    #>
    param(
        [string]$PackageId,
        [ref]$LogArray,
        [string]$OriginalPackageName
    )

    if (-not (Initialize-Chocolatey)) {
        Write-Host "  [-] Chocolatey unavailable. Skipping $PackageId." -ForegroundColor Red
        Write-EventLog -Message "Chocolatey unavailable. Skipping $PackageId." -EventType "Warning" -EventId $script:EventIds.PackageSkipped
        $LogArray.Value += [PSCustomObject]@{
            Package   = $OriginalPackageName
            Status    = "Failed"
            Manager   = "Chocolatey (Unavailable)"
            Mode      = $script:OperationMode
            ExitCode  = "N/A"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        return $false
    }

    $action = if ($script:OperationMode -eq "Upgrade") { "Upgrading" } else { "Installing" }
    Write-Host "  Attempting Chocolatey: $action $PackageId..." -ForegroundColor Yellow
    Write-EventLog -Message "Attempting to $($action.ToLower()) $PackageId via Chocolatey." -EventType "Information" -EventId $script:EventIds.PackageInstallStart
    
    try {
        if ($script:OperationMode -eq "Upgrade") {
            choco upgrade $PackageId -y 2>&1 | Out-Null
        }
        else {
            choco install $PackageId -y 2>&1 | Out-Null
        }
        
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            Write-Host "  [+] $PackageId $action successfully via Chocolatey." -ForegroundColor Green
            Write-EventLog -Message "$action $PackageId successfully via Chocolatey." -EventType "Information" -EventId $script:EventIds.PackageInstallSuccess
            $LogArray.Value += [PSCustomObject]@{
                Package   = $OriginalPackageName
                Status    = "Success"
                Manager   = "Chocolatey"
                Mode      = $script:OperationMode
                ExitCode  = $exitCode
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            return $true
        }
        else {
            Write-Host "  [-] Chocolatey failed for $PackageId (Exit: $exitCode)." -ForegroundColor Red
            Write-EventLog -Message "Chocolatey failed for $PackageId (Exit Code: $exitCode)." -EventType "Error" -EventId $script:EventIds.PackageInstallFailed
            $LogArray.Value += [PSCustomObject]@{
                Package   = $OriginalPackageName
                Status    = "Failed"
                Manager   = "Chocolatey"
                Mode      = $script:OperationMode
                ExitCode  = $exitCode
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            return $false
        }
    }
    catch {
        Write-Host "  [-] Chocolatey error for $PackageId`: \$_" -ForegroundColor Red
        Write-EventLog -Message "Chocolatey error for \$PackageId`: $_" -EventType "Error" -EventId $script:EventIds.PackageInstallFailed
        $LogArray.Value += [PSCustomObject]@{
            Package   = $OriginalPackageName
            Status    = "Failed"
            Manager   = "Chocolatey"
            Mode      = $script:OperationMode
            ExitCode  = "N/A"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        return $false
    }
}

function Install-Software {
    <#
    .SYNOPSIS
        Iterates through software list and installs via Winget (primary) or Chocolatey (fallback).
    #>
    param(
        [hashtable[]]$SoftwareList,
        [bool]$WingetAvailable,
        [ref]$LogArray
    )

    if ($SoftwareList.Count -eq 0) {
        Write-Host "No packages to $($script:OperationMode.ToLower())." -ForegroundColor Yellow
        Write-EventLog -Message "No packages to $($script:OperationMode.ToLower())." -EventType "Information" -EventId $script:EventIds.PackageInstallStart
        return
    }

    $actionVerb = if ($script:OperationMode -eq "Upgrade") { "Upgrade" } else { "Installation" }
    Write-Host "`n--- Starting \$actionVerb Phase ---" -ForegroundColor Magenta
    Write-EventLog -Message "Starting \$actionVerb phase for \$(\$SoftwareList.Count) package(s)." -EventType "Information" -EventId $script:EventIds.PackageInstallStart

    foreach ($software in \$SoftwareList) {
        \$wingetId  = \$software.Winget
        \$chocoId   = \$software.Chocolatey

        if (-not \$WingetAvailable) {
            Write-Host "Winget unavailable. Skipping \$wingetId." -ForegroundColor Yellow
            Write-EventLog -Message "Winget unavailable. Skipping \$wingetId." -EventType "Warning" -EventId \$script:EventIds.PackageSkipped
            \$LogArray.Value += [PSCustomObject]@{
                Package   = \$wingetId
                Status    = "Skipped"
                Manager   = "N/A"
                Mode      = \$script:OperationMode
                ExitCode  = "N/A"
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            continue
        }

        \$wingetSuccess = Install-PackageViaWinget -PackageId \$wingetId -LogArray \$LogArray

        if (-not \$wingetSuccess) {
            Install-PackageViaChocolatey -PackageId \$chocoId -LogArray \$LogArray -OriginalPackageName $wingetId | Out-Null
        }
    }
}

function Export-InstallLog {
    <#
    .SYNOPSIS
        Exports the installation log to CSV file for compliance and audit.
    #>
    param(
        [array]$InstallLog,
        [string]\$Path
    )
    try {
        # Verify the directory exists before attempting to write
        if (-not (Test-Path \$script:LogRoot)) {
            Write-Host "Warning: \$script:LogRoot does not exist. Creating directory..." -ForegroundColor Yellow
            New-Item -Path \$script:LogRoot -ItemType Directory -Force | Out-Null
        }

        \$InstallLog | Export-Csv -Path \$Path -NoTypeInformation -Append
        Write-Host "Log saved to: \$Path" -ForegroundColor Cyan
        Write-EventLog -Message "Installation log exported to \$Path" -EventType "Information" -EventId \$script:EventIds.LogExportSuccess
    }
    catch {
        Write-Host "Warning: Could not write log file: \$_" -ForegroundColor Yellow
        Write-EventLog -Message "Could not write CSV log file to \$Path. Error: \$_" -EventType "Warning" -EventId $script:EventIds.LogExportFailed
    }
}

function Show-InstallationSummary {
    <#
    .SYNOPSIS
        Displays a formatted installation summary report with statistics.
    #>
    param([array]$InstallLog)

    $successCount = ($InstallLog | Where-Object { \$_.Status -eq "Success" }).Count
    $failureCount  = ($InstallLog | Where-Object { \$_.Status -eq "Failed"  }).Count
    $skipCount     = ($InstallLog | Where-Object { \$_.Status -eq "Skipped" }).Count
    \$totalCount    = \$InstallLog.Count
    
    $actionVerb = if ($script:OperationMode -eq "Upgrade") { "Upgrade" } else { "Installation" }

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "S.P.A.R.K \$actionVerb Summary Report"   -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Mode: \$(\$script:OperationMode)" -ForegroundColor Cyan
    Write-Host "Execution Timestamp: $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "$actionVerb Statistics:" -ForegroundColor Magenta
    Write-Host "├─ Total Packages : \$totalCount"
    Write-Host "├─ Successful     : \$successCount" -ForegroundColor Green
    Write-Host "├─ Failed         : \$failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })
    Write-Host "└─ Skipped        : $skipCount" -ForegroundColor Yellow
    Write-Host ""

    if ($successCount -gt 0) {
        Write-Host "Successfully \$(\$script:OperationMode.ToLower())ed:" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        \$InstallLog | Where-Object { \$_.Status -eq "Success" } | ForEach-Object {
            Write-Host "  [+] \$(\$_.Package)" -ForegroundColor Green
            Write-Host "      Manager: \$(\$_.Manager) | Exit: \$(\$_.ExitCode) | Time: \$($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    if ($failureCount -gt 0) {
        Write-Host "Failed Operations:" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        \$InstallLog | Where-Object { \$_.Status -eq "Failed" } | ForEach-Object {
            Write-Host "  [-] \$(\$_.Package)" -ForegroundColor Red
            Write-Host "      Manager: \$(\$_.Manager) | Exit: \$(\$_.ExitCode) | Time: \$($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    if ($skipCount -gt 0) {
        Write-Host "Skipped Operations:" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        \$InstallLog | Where-Object { \$_.Status -eq "Skipped" } | ForEach-Object {
            Write-Host "  [~] \$(\$_.Package)" -ForegroundColor Yellow
            Write-Host "      Reason: \$(\$_.Manager) | Time: \$(\$_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Completed at \$(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""

    # Log the summary to event log
    \$summaryMessage = "S.P.A.R.K \$actionVerb completed. Total: \$totalCount | Success: \$successCount | Failed: \$failureCount | Skipped: \$skipCount"
    $summaryEventType = if ($failureCount -gt 0) { "Warning" } else { "Information" }
    Write-EventLog -Message \$summaryMessage -EventType \$summaryEventType -EventId \$script:EventIds.SummaryReport
}

# ─────────────────────────────────────────────────────────────────────────────
# SOFTWARE DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

# Core packages — required for all deployments
\$coreSoftware = @(
    @{ Winget = "Microsoft.Edge";                    Chocolatey = "microsoft-edge"        },
    @{ Winget = "7zip.7zip";                         Chocolatey = "7zip"                  },
    @{ Winget = "Adobe.Acrobat.Reader.64-bit";       Chocolatey = "adobereader"           },
    @{ Winget = "Zoom.Zoom";                         Chocolatey = "zoom"                  },
    @{ Winget = "Microsoft.Office";                  Chocolatey = "office365business"     }
)

# Optional packages — controlled via script parameters
\$optionalSoftware = @(
    @{ ParamName = "InstallDellCommandUpdate"; Winget = "Dell.CommandUpdate";        Chocolatey = "dell-command-update" },
    @{ ParamName = "InstallZoomOutlookPlugin"; Winget = "Zoom.ZoomOutlookPlugin";   Chocolatey = \$null                 },
    @{ ParamName = "InstallDellCommand";       Winget = "Dell.Command";              Chocolatey = "dell-command-suite"  }
)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

Show-SparkBanner

# Initialize event logging
Initialize-EventLog
Write-EventLog -Message "S.P.A.R.K started. Mode: \$Mode | Log Path: \$(\$script:LogPath) | Time: \$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -EventType "Information" -EventId $script:EventIds.ScriptStart
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Initializing S.P.A.R.K..." -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# MODE SELECTION
# ─────────────────────────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($Mode)) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║     S.P.A.R.K - Mode Selection         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Please choose an operation mode:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Install    - Install new software packages" -ForegroundColor Green
    Write-Host "  [2] Upgrade    - Upgrade existing software to latest versions" -ForegroundColor Cyan
    Write-Host ""
    
    do {
        $selection = Read-Host "Enter your choice (1 or 2)"
        
        if ($selection -eq "1") {
            $Mode = "Install"
            break
        }
        elseif ($selection -eq "2") {
            $Mode = "Upgrade"
            break
        }
        else {
            Write-Host "Invalid selection. Please enter 1 or 2." -ForegroundColor Red
        }
    } while ($true)
}

Write-Host ""
Write-Host "
