<#
.SYNOPSIS
    S.P.A.R.K - Software Package & Resource Kit
    Automated Winget-based Software Installation

.DESCRIPTION
    S.P.A.R.K is a PowerShell script that automates the installation of essential
    software packages using Windows Package Manager (Winget). It provides a clean,
    professional interface with real-time installation tracking and comprehensive
    summary reporting.

    The script installs core software automatically and supports optional packages
    via command-line parameters or interactive prompts. Perfect for deployment in 
    enterprise environments, RMM tools, and automated workflows.

.PARAMETER InstallZoomOutlookPlugin
    Install the Zoom Outlook Plugin for seamless meeting integration.

.PARAMETER InstallDellCommandUpdate
    Install Dell Command Update for Dell system management and updates.

.PARAMETER InstallDellCommandSuite
    Install Dell Command Suite for comprehensive Dell system management.

.PARAMETER SkipPrompts
    Skip interactive prompts and only install core packages.

.EXAMPLE
    .\SPARK.ps1
    Installs all core software packages and prompts for optional packages.

.EXAMPLE
    .\SPARK.ps1 -InstallZoomOutlookPlugin -InstallDellCommandUpdate
    Installs core software plus specified optional packages.

.EXAMPLE
    .\SPARK.ps1 -SkipPrompts
    Installs only core software without prompting for optional packages.

.NOTES
    Author: Your Organization Name
    Version: 1.1.0
    Requires: Windows 10+ with Administrator privileges
    Winget will be installed automatically if not present.

.LINK
    https://github.com/CursedTechnocrat/TechnicianToolkit

#>

param(
    [switch]$InstallZoomOutlookPlugin,
    [switch]$InstallDellCommandUpdate,
    [switch]$InstallDellCommandSuite,
    [switch]$SkipPrompts
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL VARIABLES
# ─────────────────────────────────────────────────────────────────────────────

$script:StartTime = Get-Date
$script:ErrorLog = @()

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Update-EnvironmentPath {
    <#
    .SYNOPSIS
        Refreshes the environment PATH variable to include newly installed tools.
    #>
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
}

function Initialize-Winget {
    <#
    .SYNOPSIS
        Checks for Winget installation and installs if missing.
    .DESCRIPTION
        Attempts to detect Winget. If not found, downloads and installs the latest
        release from the official Microsoft Winget CLI GitHub repository.
    #>
    try {
        $ver = winget --version 2>&1
        Write-Host "✓ Winget already installed. Version: $ver" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "⚠ Winget not found. Attempting install via GitHub release..." -ForegroundColor Yellow
        try {
            $release    = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing
            $msixBundle = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1

            if (-not $msixBundle) {
                Write-Host "✗ Could not locate winget MSIX bundle in latest GitHub release." -ForegroundColor Red
                return $false
            }

            $tmpPath = "$env:TEMP\winget_install.msixbundle"
            Write-Host "  Downloading from: $($msixBundle.browser_download_url)" -ForegroundColor Cyan
            Invoke-WebRequest -Uri $msixBundle.browser_download_url -OutFile $tmpPath -UseBasicParsing
            
            Write-Host "  Installing MSIX package..." -ForegroundColor Cyan
            Add-AppxPackage -Path $tmpPath -ErrorAction Stop
            Start-Sleep -Seconds 3
            Update-EnvironmentPath

            $ver = winget --version 2>&1
            Write-Host "✓ Winget installed successfully. Version: $ver" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "✗ Failed to install Winget: $_" -ForegroundColor Red
            $script:ErrorLog += "Winget installation failed: $_"
            return $false
        }
    }
}

function Update-Winget {
    <#
    .SYNOPSIS
        Updates Winget to the latest version.
    #>
    Write-Host "`nUpdating Winget..." -ForegroundColor Magenta
    try {
        winget upgrade winget --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
        Write-Host "✓ Winget update check complete." -ForegroundColor Green
    }
    catch {
        Write-Host "⚠ Could not update Winget (this may be normal): $_" -ForegroundColor Yellow
    }
}

function Install-Software {
    <#
    .SYNOPSIS
        Installs specified packages via Winget.
    .PARAMETER WingetPackages
        Array of Winget package IDs to install.
    .PARAMETER LogArray
        Reference to array for tracking installation results.
    #>
    param(
        [string[]]$WingetPackages,
        [ref]$LogArray
    )

    if ($WingetPackages.Count -gt 0) {
        Write-Host "`n--- Installing Packages ---" -ForegroundColor Magenta
        foreach ($item in $WingetPackages) {
            Write-Host "Installing $item..." -ForegroundColor Yellow -NoNewline
            try {
                winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
                $exitCode = $LASTEXITCODE

                if ($exitCode -eq 0 -or $exitCode -eq -1978335189 -or $exitCode -eq 3010) {
                    Write-Host " ✓" -ForegroundColor Green
                    $status = "Success"
                }
                else {
                    Write-Host " ✗ (Exit: $exitCode)" -ForegroundColor Red
                    $status = "Failed"
                    $script:ErrorLog += "$item failed with exit code $exitCode"
                }
            }
            catch {
                Write-Host " ✗" -ForegroundColor Red
                $exitCode = "Exception"
                $status   = "Failed"
                $script:ErrorLog += "Exception installing $item`: $_"
            }

            $LogArray.Value += [PSCustomObject]@{
                Package   = $item
                Status    = $status
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
}

function Show-SparkBanner {
    <#
    .SYNOPSIS
        Displays the S.P.A.R.K ASCII art banner and welcome message.
    #>
    Write-Host @"

  ███████╗██████╗  █████╗ ██████╗ ██╗  ██╗
  ██╔════╝██╔══██╗██╔══██╗██╔══██╗██║ ██╔╝
  ███████╗██████╔╝███████║██████╔╝█████╔╝ 
  ╚════██║██╔═══╝ ██╔══██║██╔══██╗██╔═██╗ 
  ███████║██║     ██║  ██║██║  ██║██║  ██╗
  ╚══════╝╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    Software Package & Resource Kit" -ForegroundColor Yellow
    Write-Host "    Automated Winget-based Software Installation" -ForegroundColor Yellow
    Write-Host ""
}

function Show-OptionalPackageMenu {
    <#
    .SYNOPSIS
        Displays interactive menu for optional packages.
    .DESCRIPTION
        Shows user a menu to select optional packages to install.
    .OUTPUT
        Returns hashtable with selected optional packages.
    #>
    Write-Host "`n" + ("=" * 50) -ForegroundColor Cyan
    Write-Host "S.P.A.R.K - Optional Packages" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select optional packages to install:" -ForegroundColor Yellow
    Write-Host ""
    
    $optionalChoices = @(
        @{ 
            Name = "Zoom Outlook Plugin"
            Description = "Seamless meeting integration with Outlook"
            WingetID = "Zoom.ZoomOutlookPlugin"
            Selected = $false
        },
        @{ 
            Name = "Dell Command Update"
            Description = "Dell system management & driver updates"
            WingetID = "Dell.CommandUpdate"
            Selected = $false
        },
        @{ 
            Name = "Dell Command Suite"
            Description = "Comprehensive Dell system management tools"
            WingetID = "Dell.CommandSuite"
            Selected = $false
        }
    )

    for ($i = 0; $i -lt $optionalChoices.Count; $i++) {
        Write-Host "  [$($i + 1)] $($optionalChoices[$i].Name)" -ForegroundColor Green
        Write-Host "      └─ $($optionalChoices[$i].Description)" -ForegroundColor Gray
    }
    
    Write-Host "  [0] Skip optional packages" -ForegroundColor Yellow
    Write-Host ""

    $selectedPackages = @()

    # Loop for multiple selections
    while ($true) {
        $choice = Read-Host "Enter selection (comma-separated for multiple, or 0 to skip)"
        
        if ($choice -eq "0") {
            Write-Host "✓ Skipping optional packages" -ForegroundColor Green
            break
        }

        $choices = $choice -split "," | ForEach-Object { $_.Trim() }
        $allValid = $true

        foreach ($c in $choices) {
            if ([int]$c -lt 1 -or [int]$c -gt $optionalChoices.Count) {
                Write-Host "✗ Invalid selection: $c" -ForegroundColor Red
                $allValid = $false
                break
            }
        }

        if ($allValid) {
            foreach ($c in $choices) {
                $selectedPackages += $optionalChoices[[int]$c - 1]
            }
            break
        }
    }

    return $selectedPackages
}

function Show-InstallationSummary {
    <#
    .SYNOPSIS
        Displays a comprehensive installation summary report.
    .PARAMETER InstallLog
        Array of installation results to summarize.
    #>
    param([array]$InstallLog)

    $successCount = ($InstallLog | Where-Object { $_.Status -eq "Success" }).Count
    $failureCount  = ($InstallLog | Where-Object { $_.Status -eq "Failed"  }).Count
    $totalCount    = $InstallLog.Count
    $elapsedTime   = (Get-Date) - $script:StartTime

    Write-Host "`n" + ("=" * 50) -ForegroundColor Magenta
    Write-Host "S.P.A.R.K Installation Summary Report" -ForegroundColor Magenta
    Write-Host ("=" * 50) -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Execution Timestamp: $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')" -ForegroundColor Yellow
    Write-Host "Total Execution Time: $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Installation Statistics:" -ForegroundColor Magenta
    Write-Host "├─ Total Packages : $totalCount"
    Write-Host "├─ Successful     : $successCount" -ForegroundColor Green

    if ($failureCount -gt 0) {
        Write-Host "└─ Failed         : $failureCount" -ForegroundColor Red
    }
    else {
        Write-Host "└─ Failed         : $failureCount" -ForegroundColor Green
    }
    Write-Host ""

    if ($successCount -gt 0) {
        Write-Host "Successfully Installed:" -ForegroundColor Green
        Write-Host ("=" * 50) -ForegroundColor Green
        $InstallLog | Where-Object { $_.Status -eq "Success" } | ForEach-Object {
            Write-Host "  [✓] $($_.Package)" -ForegroundColor Green
            Write-Host "      Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    if ($failureCount -gt 0) {
        Write-Host "Failed Installations:" -ForegroundColor Red
        Write-Host ("=" * 50) -ForegroundColor Red
        $InstallLog | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
            Write-Host "  [✗] $($_.Package)" -ForegroundColor Red
            Write-Host "      Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    Write-Host ("=" * 50) -ForegroundColor Magenta
    Write-Host "Completed at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
    Write-Host ("=" * 50) -ForegroundColor Magenta
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PACKAGE DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

# Core software packages (installed automatically)
$coreSoftware = @(
    "Microsoft.Teams",
    "Microsoft.Office",
    "7zip.7zip",
    "Adobe.Acrobat.Reader.64-bit",
    "Zoom.Zoom",
    "Google.Chrome"
)

# Optional software packages (installed via parameters or menu)
$optionalSoftware = @(
    @{ ParamName = "InstallZoomOutlookPlugin"; Winget = "Zoom.ZoomOutlookPlugin"   },
    @{ ParamName = "InstallDellCommandUpdate"; Winget = "Dell.CommandUpdate"       },
    @{ ParamName = "InstallDellCommandSuite";  Winget = "Dell.CommandSuite"        }
)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

$installationLog = @()

Show-SparkBanner

Write-Host ("=" * 50) -ForegroundColor Magenta
Write-Host "Initializing S.P.A.R.K" -ForegroundColor Magenta
Write-Host ("=" * 50) -ForegroundColor Magenta
Write-Host ""

Update-EnvironmentPath

$wingetAvailable = Initialize-Winget

if (-not $wingetAvailable) {
    Write-Host "`n✗ ERROR: Winget could not be initialized. Exiting." -ForegroundColor Red
    exit 1
}

Update-Winget

$wingetList = @()

# Add core packages
foreach ($software in $coreSoftware) {
    $wingetList += $software
}

# Handle optional packages - from parameters OR from interactive menu
$selectedOptional = @()

# Check if any optional parameters were passed
$hasOptionalParams = $PSBoundParameters.Keys | Where-Object { 
    $_ -match "^Install" 
} | Measure-Object | Select-Object -ExpandProperty Count

if ($hasOptionalParams -gt 0) {
    # Add optional packages from parameters
    foreach ($opt in $optionalSoftware) {
        $switchValue = (Get-Variable -Name $opt.ParamName -ValueOnly -ErrorAction SilentlyContinue)
        if ($switchValue -eq $true) {
            $wingetList += $opt.Winget
            $selectedOptional += $opt.Winget
        }
    }
}
elseif (-not $SkipPrompts) {
    # Show interactive menu if no parameters passed and not skipping prompts
    $menuSelection = Show-OptionalPackageMenu
    
    if ($menuSelection.Count -gt 0) {
        foreach ($item in $menuSelection) {
            $wingetList += $item.WingetID
            $selectedOptional += $item.Name
        }
    }
}

Write-Host "`n" + ("=" * 50) -ForegroundColor Magenta
Write-Host "S.P.A.R.K - Installation Phase" -ForegroundColor Magenta
Write-Host ("=" * 50) -ForegroundColor Magenta
Write-Host ""

if ($selectedOptional.Count -gt 0) {
    Write-Host "Optional Packages Selected:" -ForegroundColor Cyan
    foreach ($item in $selectedOptional) {
        Write-Host "  + $item" -ForegroundColor Cyan
    }
    Write-Host ""
}

Install-Software -WingetPackages $wingetList -LogArray ([ref]$installationLog)

Show-InstallationSummary -InstallLog $installationLog

Write-Host "ℹ Note: Some installations may require a system restart to complete." -ForegroundColor Yellow
Write-Host ""
