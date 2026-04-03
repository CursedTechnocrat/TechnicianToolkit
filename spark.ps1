<#
.SYNOPSIS
    S.P.A.R.K - Software Package & Resource Kit
    Automated Package Manager Setup & Installation

.DESCRIPTION
    Installs core and optional MSP software packages using Winget (primary)
    or Chocolatey (fallback). Designed for fully unattended use in RMM tools,
    Kaseya LiveConnect, and Task Scheduler. All optional package selections
    are parameter-driven — no interactive prompts.

.PARAMETER InstallZoomPlugin
    Install the Zoom Outlook Plugin.

.PARAMETER InstallDisplayLink
    Install the DisplayLink Graphics Driver.

.PARAMETER InstallDellCommandUpdate
    Install Dell Command Update.

.PARAMETER LogPath
    Path for the installation log CSV.
    Default: C:\ProgramData\SPARK\install_log.csv

.EXAMPLE
    .\SPARK.ps1
    Installs all core software using Winget (or Chocolatey fallback).

.EXAMPLE
    .\SPARK.ps1 -InstallZoomPlugin -InstallDellCommandUpdate
    Installs core software plus the Zoom plugin and Dell Command Update.

.EXAMPLE
    .\SPARK.ps1 -InstallDisplayLink -LogPath "D:\Logs\spark.csv"
    Installs core software plus DisplayLink, logging to a custom path.
#>

param(
    [switch]$InstallZoomPlugin,
    [switch]$InstallDisplayLink,
    [switch]$InstallDellCommandUpdate,
    [string]$LogPath = "C:\ProgramData\SPARK\install_log.csv"
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
# FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Update-EnvironmentPath {
    # Renamed from Refresh-EnvironmentPath to use an approved PS verb.
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
}

function Initialize-Winget {
    try {
        $ver = winget --version 2>&1
        Write-Host "Winget already installed. Version: $ver" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Winget not found. Attempting install via GitHub release..." -ForegroundColor Yellow
        try {
            # Pull the latest release from the official winget-cli repo.
            # This avoids the irm|iex anti-pattern and downloads a verified MSIX.
            $release    = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing
            $msixBundle = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1

            if (-not $msixBundle) {
                Write-Host "Could not locate winget MSIX bundle in latest GitHub release." -ForegroundColor Red
                return $false
            }

            $tmpPath = "$env:TEMP\winget_install.msixbundle"
            Invoke-WebRequest -Uri $msixBundle.browser_download_url -OutFile $tmpPath -UseBasicParsing
            Add-AppxPackage -Path $tmpPath
            Start-Sleep -Seconds 3
            Update-EnvironmentPath

            $ver = winget --version 2>&1
            Write-Host "Winget installed successfully. Version: $ver" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "Failed to install Winget: $_" -ForegroundColor Red
            return $false
        }
    }
}

function Initialize-Chocolatey {
    try {
        $ver = choco --version 2>&1
        Write-Host "Chocolatey already installed. Version: $ver" -ForegroundColor Green
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
                return $true
            }
            catch {
                Write-Host "Chocolatey present but inaccessible. A session restart may be required." -ForegroundColor Yellow
                return $false
            }
        }
        else {
            Write-Host "Chocolatey not found. Installing..." -ForegroundColor Yellow
            try {
                Set-ExecutionPolicy Bypass -Scope Process -Force
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

                # Using full cmdlet name (Invoke-Expression) instead of the iex alias,
                # and downloading the string separately so it can be inspected/logged.
                $chocoScript = (New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')
                Invoke-Expression $chocoScript

                Start-Sleep -Seconds 2
                Update-EnvironmentPath
                Write-Host "Chocolatey installed successfully." -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "Failed to install Chocolatey: $_" -ForegroundColor Red
                return $false
            }
        }
    }
}

function Update-PackageManagers {
    param(
        [bool]$UpdateWinget,
        [bool]$UpdateChocolatey
    )

    if ($UpdateWinget) {
        Write-Host "`nUpdating Winget..." -ForegroundColor Magenta
        try {
            winget upgrade winget --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
            Write-Host "Winget update check complete." -ForegroundColor Green
        }
        catch {
            Write-Host "Could not update Winget (this may be normal in some environments): $_" -ForegroundColor Yellow
        }
    }

    if ($UpdateChocolatey) {
        Write-Host "`nUpdating Chocolatey..." -ForegroundColor Magenta
        try {
            choco upgrade chocolatey -y 2>&1 | Out-Null
            Write-Host "Chocolatey updated." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to update Chocolatey: $_" -ForegroundColor Red
        }
    }
}

function Install-Software {
    param(
        [string[]]$WingetPackages,
        [string[]]$ChocoPackages,
        [bool]$UseWinget,
        [bool]$UseChocolatey,
        [ref]$LogArray
    )

    if ($UseWinget -and $WingetPackages.Count -gt 0) {
        Write-Host "`n--- Installing via Winget ---" -ForegroundColor Magenta
        foreach ($item in $WingetPackages) {
            Write-Host "Installing $item..." -ForegroundColor Yellow
            try {
                winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
                $exitCode = $LASTEXITCODE

                # Exit code 0           = success
                # Exit code -1978335189 = already installed (0x8A150011)
                if ($exitCode -eq 0 -or $exitCode -eq -1978335189) {
                    Write-Host "$item installed successfully." -ForegroundColor Green
                    $status = "Success"
                }
                else {
                    Write-Host "Failed to install $item. Exit code: $exitCode" -ForegroundColor Red
                    $status = "Failed"
                }
            }
            catch {
                Write-Host "Error installing $item`: $_" -ForegroundColor Red
                $exitCode = "N/A"
                $status   = "Failed"
            }

            $LogArray.Value += [PSCustomObject]@{
                Package   = $item
                Status    = $status
                Manager   = "Winget"
                ExitCode  = $exitCode
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }

    if ($UseChocolatey -and $ChocoPackages.Count -gt 0) {
        Write-Host "`n--- Installing via Chocolatey ---" -ForegroundColor Magenta
        foreach ($item in $ChocoPackages) {
            Write-Host "Installing $item..." -ForegroundColor Yellow
            try {
                choco install $item -y 2>&1 | Out-Null
                $exitCode = $LASTEXITCODE

                if ($exitCode -eq 0) {
                    Write-Host "$item installed successfully." -ForegroundColor Green
                    $status = "Success"
                }
                else {
                    Write-Host "Failed to install $item. Exit code: $exitCode" -ForegroundColor Red
                    $status = "Failed"
                }
            }
            catch {
                Write-Host "Error installing $item`: $_" -ForegroundColor Red
                $exitCode = "N/A"
                $status   = "Failed"
            }

            $LogArray.Value += [PSCustomObject]@{
                Package   = $item
                Status    = $status
                Manager   = "Chocolatey"
                ExitCode  = $exitCode
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
}

function Export-InstallLog {
    param(
        [array]$InstallLog,
        [string]$Path
    )
    try {
        $dir = Split-Path $Path -Parent
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $InstallLog | Export-Csv -Path $Path -NoTypeInformation -Append
        Write-Host "Log saved to: $Path" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Warning: Could not write log file: $_" -ForegroundColor Yellow
    }
}

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

function Show-InstallationSummary {
    param([array]$InstallLog)

    $successCount = ($InstallLog | Where-Object { $_.Status -eq "Success" }).Count
    $failureCount  = ($InstallLog | Where-Object { $_.Status -eq "Failed"  }).Count
    $totalCount    = $InstallLog.Count

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "S.P.A.R.K Installation Summary Report"   -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Execution Timestamp: $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')" -ForegroundColor Yellow
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
        Write-Host "========================================" -ForegroundColor Green
        $InstallLog | Where-Object { $_.Status -eq "Success" } | ForEach-Object {
            Write-Host "  [+] $($_.Package)" -ForegroundColor Green
            Write-Host "      Manager: $($_.Manager) | Exit: $($_.ExitCode) | Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    if ($failureCount -gt 0) {
        Write-Host "Failed Installations:" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        $InstallLog | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
            Write-Host "  [-] $($_.Package)" -ForegroundColor Red
            Write-Host "      Manager: $($_.Manager) | Exit: $($_.ExitCode) | Time: $($_.Timestamp)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Completed at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SOFTWARE DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

# Core packages — Winget is primary, Chocolatey IDs are the fallback.
$coreSoftware = @(
    @{ Winget = "Microsoft.Teams";             Chocolatey = "microsoft-teams" },
    @{ Winget = "Microsoft.Office";            Chocolatey = "office-deploy"   },
    @{ Winget = "7zip.7zip";                   Chocolatey = "7zip"            },
    @{ Winget = "Google.Chrome";               Chocolatey = "googlechrome"    },
    @{ Winget = "Adobe.Acrobat.Reader.64-bit"; Chocolatey = "adobereader"     },
    @{ Winget = "Zoom.Zoom";                   Chocolatey = "zoom"            }
)

# Optional packages — Winget only. Controlled via script parameters.
$optionalSoftware = @(
    @{ ParamName = "InstallZoomPlugin";        Winget = "Zoom.ZoomOutlookPlugin"     },
    @{ ParamName = "InstallDisplayLink";       Winget = "DisplayLink.GraphicsDriver" },
    @{ ParamName = "InstallDellCommandUpdate"; Winget = "Dell.CommandUpdate"         }
)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

$installationLog = @()

Show-SparkBanner

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Initializing S.P.A.R.K"                  -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Update-EnvironmentPath

$wingetAvailable = Initialize-Winget
$chocoAvailable  = Initialize-Chocolatey

if (-not $wingetAvailable -and -not $chocoAvailable) {
    Write-Host "`nERROR: Neither Winget nor Chocolatey could be initialized. Exiting." -ForegroundColor Red
    exit 1
}

Update-PackageManagers -UpdateWinget $wingetAvailable -UpdateChocolatey $chocoAvailable

# Build install lists.
# Winget is primary — Chocolatey IDs are only used if Winget is unavailable.
# This prevents the same package from being installed twice.
$wingetList = @()
$chocoList  = @()

foreach ($software in $coreSoftware) {
    if ($wingetAvailable) {
        $wingetList += $software.Winget
    }
    elseif ($chocoAvailable) {
        $chocoList += $software.Chocolatey
    }
}

# Add optional packages whose corresponding switch was passed at runtime.
foreach ($opt in $optionalSoftware) {
    $switchValue = (Get-Variable -Name $opt.ParamName -ValueOnly -ErrorAction SilentlyContinue)
    if ($switchValue -eq $true) {
        if ($wingetAvailable) {
            $wingetList += $opt.Winget
        }
        else {
            Write-Host "Skipping optional package $($opt.Winget) — requires Winget, which is unavailable." -ForegroundColor Yellow
        }
    }
}

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "S.P.A.R.K - Installation Phase"          -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Install-Software `
    -WingetPackages $wingetList `
    -ChocoPackages  $chocoList `
    -UseWinget      $wingetAvailable `
    -UseChocolatey  $chocoAvailable `
    -LogArray       ([ref]$installationLog)

Show-InstallationSummary -InstallLog $installationLog
Export-InstallLog        -InstallLog $installationLog -Path $LogPath

Write-Host "Note: Some installations may require a system restart to complete." -ForegroundColor Yellow
Write-Host ""
