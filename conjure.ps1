# conjure.ps1 - C.O.N.J.U.R.E. — Centrally Orchestrates Network-Joined Updates, Rollouts & Executables
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 CursedTechnocrat and the Technician Toolkit contributors
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# SPDX-License-Identifier: GPL-3.0-or-later

<#
.SYNOPSIS
    C.O.N.J.U.R.E. — Centrally Orchestrates Network-Joined Updates, Rollouts & Executables
    Automated Software Deployment Tool for PowerShell 5.1+

.DESCRIPTION
    Manages software deployment using the Windows Package Manager (winget)
    or Chocolatey. Installs required and optional software packages, supports
    an upgrade-all mode for keeping existing packages current, and tracks
    installation status per package.

    Software with no winget or Chocolatey package — RMM agents especially,
    which are usually a per-tenant link with a token in the URL — can be
    installed from a direct HTTPS download instead. Those installers are
    fetched, identified by their own magic bytes, optionally checked against
    a pinned SHA-256, and run silently. Recurring ones can be stored in
    config.json under Conjure.DirectDownloads so they deploy unattended.

.USAGE
    PS C:\> .\conjure.ps1                 # Interactive mode — Must be run as Administrator
    PS C:\> .\conjure.ps1 -Unattended     # Unattended mode — Required packages only, no prompts
    PS C:\> .\conjure.ps1 -DirectUrl 'https://vendor.example/agent.msi'
    PS C:\> .\conjure.ps1 -Unattended -WhatIf   # Preview everything; install nothing

.NOTES
    Version : 5.0

#>

param(
    [string[]]$DirectUrl,
    [switch]$Unattended,
    [switch]$Transcript,
    [switch]$WhatIf
)

# ===========================
# CONFIGURATION
# ===========================

$ScriptPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = Get-Location
}

$ExecutionTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

# Set console to UTF-8 so Unicode block characters render correctly
# ===========================
# SHARED MODULE BOOTSTRAP
# ===========================
$TKModulePath = Join-Path $PSScriptRoot 'TechnicianToolkit.psm1'
if (-not (Test-Path $TKModulePath)) {
    $TKModuleUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1'
    Write-Host "  [*] Shared module TechnicianToolkit.psm1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $TKModuleUrl -OutFile $TKModulePath -ErrorAction Stop
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($TKModulePath, [ref]$null, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            Remove-Item -Path $TKModulePath -Force -ErrorAction SilentlyContinue
            Write-Host "  [!!] Downloaded module failed syntax validation - file removed." -ForegroundColor Red
            Write-Host "       $($parseErrors[0].Message)" -ForegroundColor Red
            exit 1
        }
        Write-Host "  [+] Module downloaded and verified." -ForegroundColor Green
    } catch {
        Write-Host "  [!!] Could not download TechnicianToolkit.psm1:" -ForegroundColor Red
        Write-Host "       $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Place the module manually next to this script from:" -ForegroundColor Yellow
        Write-Host "       $TKModuleUrl" -ForegroundColor Yellow
        exit 1
    }
}
Import-Module $TKModulePath -Force -ErrorAction Stop

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# Single source of truth for the deployment catalog. Each entry carries a friendly
# Name plus the package ID for every supported manager, so the winget and Chocolatey
# IDs can never drift out of alignment. Add or remove a package in exactly one place;
# Get-PackageId resolves an entry to the ID for the manager selected at runtime.
$RequiredCatalog = @(
    [PSCustomObject]@{ Name = 'Microsoft Teams'; Winget = 'Microsoft.Teams';  Choco = 'microsoft-teams' }
    [PSCustomObject]@{ Name = 'Microsoft 365';   Winget = 'Microsoft.Office'; Choco = 'microsoft365apps' }
    [PSCustomObject]@{ Name = '7-Zip';           Winget = '7zip.7zip';        Choco = '7zip' }
    [PSCustomObject]@{ Name = 'Google Chrome';   Winget = 'Google.Chrome';    Choco = 'googlechrome' }
    [PSCustomObject]@{ Name = 'Zoom';            Winget = 'Zoom.Zoom';        Choco = 'zoom' }
)

$OptionalCatalog = @(
    [PSCustomObject]@{ Name = 'Zoom Outlook Plugin'; Winget = 'Zoom.ZoomOutlookPlugin'; Choco = 'zoom-outlook' }
    [PSCustomObject]@{ Name = 'Mozilla Firefox';     Winget = 'Mozilla.Firefox';        Choco = 'firefox' }
    [PSCustomObject]@{ Name = 'Dell Command Update'; Winget = 'Dell.CommandUpdate';      Choco = 'dell-command-update' }
    [PSCustomObject]@{ Name = 'Asana';               Winget = 'Asana.Asana';             Choco = 'asana' }
    [PSCustomObject]@{ Name = 'Google Earth Pro';    Winget = 'Google.EarthPro';         Choco = 'googleearthpro' }
)

# Adobe Acrobat is handled separately — the operator chooses Reader or Pro up front
# (see Select-AdobeEdition) and the chosen package is appended to the required list.
# Chocolatey's community repo has no Acrobat Pro package, so the Pro entry's Choco
# slot intentionally falls back to Reader (Select-AdobeEdition prints a note for it).
$AdobeReader = [PSCustomObject]@{ Name = 'Adobe Acrobat Reader'; Winget = 'Adobe.Acrobat.Reader.64-bit'; Choco = 'adobereader' }
$AdobePro    = [PSCustomObject]@{ Name = 'Adobe Acrobat Pro';    Winget = 'Adobe.Acrobat.Pro';           Choco = 'adobereader' }

$PackageManager = "winget"

# Resolve a catalog entry to the package ID for the manager selected at runtime.
function Get-PackageId {
    param([Parameter(Mandatory)][PSObject]$Entry)
    if ($script:PackageManager -eq "chocolatey") { return $Entry.Choco } else { return $Entry.Winget }
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

function Show-Banner {
    if (-not $Unattended) { Clear-Host }
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

# ---------------------------------------------------------------------------
# EXIT-CODE DECODING
# ---------------------------------------------------------------------------
# winget (and the installers it launches) report their result through the
# process exit code, which PowerShell surfaces as a signed 32-bit integer in
# $LASTEXITCODE. A winget result such as 0x8A150011 (installer hash mismatch)
# therefore arrives as the raw number -1978335215 — meaningless to a
# technician reading the run log, and previously logged verbatim. This table
# maps the codes we recognise to a plain-English reason and a class:
#
#   Success  installed, or a no-op because the package was already current
#   Failed   nothing was installed
#   (Warning is the fallback for a non-zero code we can't positively read)
#
# Keys are the unsigned hex form so they line up 1:1 with Microsoft's published
# winget return-code list; Resolve-InstallExit normalises the signed
# $LASTEXITCODE back to that form before looking it up. Codes and text are taken
# from the winget return-code documentation:
#   https://learn.microsoft.com/windows/package-manager/winget/returnCodes
$InstallExitInfo = @{
    # --- Clean success / reboot pending (Windows installer conventions) ---
    '0x00000000' = @{ Class = 'Success'; Reason = 'Installed successfully' }
    '0x00000BC2' = @{ Class = 'Success'; Reason = 'Installed successfully — a reboot is required to finish' }  # 3010 ERROR_SUCCESS_REBOOT_REQUIRED
    '0x00000669' = @{ Class = 'Success'; Reason = 'Installed successfully — a reboot has been initiated' }      # 1641 ERROR_SUCCESS_REBOOT_INITIATED
    '0x000003A3' = @{ Class = 'Success'; Reason = 'Installer reported success (vendor code 931)' }              # observed installer success code

    # --- Already current: winget no-op, not a failure ---
    '0x8A150061' = @{ Class = 'Success'; Reason = 'Already installed — an existing version of the package is present' }
    '0x8A15002B' = @{ Class = 'Success'; Reason = 'Already up to date — no applicable update found' }

    # --- Hard failures: nothing was installed ---
    '0x8A150011' = @{ Class = 'Failed';  Reason = "Installer hash mismatch — the downloaded installer's hash does not match the manifest, so winget rejected it" }
    '0x8A150010' = @{ Class = 'Failed';  Reason = 'No applicable installer — none of the installers fit this system' }
    '0x8A150008' = @{ Class = 'Failed';  Reason = 'Download failed — the installer could not be downloaded' }
    '0x8A150006' = @{ Class = 'Failed';  Reason = 'Installer failed — running the installer (ShellExecute) returned an error' }
    '0x8A150003' = @{ Class = 'Failed';  Reason = 'Command failed — winget could not complete the install' }
    '0x8A150056' = @{ Class = 'Failed';  Reason = 'Installer cannot be run from an administrator context' }
}

# Resolve a raw $LASTEXITCODE to a verdict object: { Class; Reason; Code; Hex }.
# Unrecognised non-zero codes are surfaced as 'Warning' — the run finished with
# a code we can't read, which is worth flagging but not treated as a clean
# install. Known-failure codes (e.g. a hash mismatch) return 'Failed' so the
# summary counts them honestly instead of as "installed with warnings".
function Resolve-InstallExit {
    param([int]$ExitCode)

    # 4294967295 (0xFFFFFFFF) masks the sign-extended int back to unsigned 32-bit;
    # written in decimal so it is unambiguously [long] on Windows PowerShell 5.1.
    $hex  = '0x{0:X8}' -f ($ExitCode -band 4294967295)
    $info = $script:InstallExitInfo[$hex]

    if ($info) {
        return [PSCustomObject]@{ Class = $info.Class; Reason = $info.Reason; Code = $ExitCode; Hex = $hex }
    }

    return [PSCustomObject]@{
        Class  = 'Warning'
        Reason = "Finished with an unrecognised exit code ($ExitCode / $hex) — verify the package manually"
        Code   = $ExitCode
        Hex    = $hex
    }
}

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

    $total = $SoftwareList.Count
    $i     = 0

    foreach ($item in $SoftwareList) {
        $i++
        Write-Progress -Activity "Installing $Type software" `
                       -Status        "[$i/$total] $item" `
                       -PercentComplete ([math]::Floor($i / $total * 100))
        Write-Host "[$i/$total] Installing: $item..." -ForegroundColor $Colors.Info

        if ($WhatIf) {
            $mgr = if ($script:PackageManager -eq "chocolatey") { "choco install $item -y" }
                   else { "winget install -e --id $item -h" }
            Write-Host "[~] WhatIf: would run  $mgr" -ForegroundColor $Colors.Accent
            Add-InstallationRecord -Software $item -Status "WHATIF - not installed"
            continue
        }

        try {
            $startTime = Get-Date

            if ($script:PackageManager -eq "chocolatey") {
                $output = & choco install $item -y 2>&1
            }
            else {
                $output = & winget install -e --id $item --accept-source-agreements --accept-package-agreements -h 2>&1
            }

            $exitCode    = $LASTEXITCODE
            $installTime = (Get-Date).ToString('HH:mm:ss')
            $elapsed     = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)

            $verdict = Resolve-InstallExit -ExitCode $exitCode

            switch ($verdict.Class) {
                'Success' {
                    Write-Host "[OK] $item — $($verdict.Reason) at $installTime (${elapsed}s)" -ForegroundColor $Colors.Success
                    Add-InstallationRecord -Software $item -Status "INSTALLED"
                }
                'Failed' {
                    Write-Host "[ERROR] $item failed — $($verdict.Reason) [exit code $exitCode] at $installTime" -ForegroundColor $Colors.Error
                    Add-InstallationRecord -Software $item -Status "FAILED - $($verdict.Reason)"
                    Write-TKError -ScriptName 'conjure' -Message "Package install failed ('$item' via $PackageManager): $($verdict.Reason) [exit code $exitCode]" -Category 'Package Install'
                }
                default {
                    Write-Host "[!!] $item — $($verdict.Reason) at $installTime" -ForegroundColor $Colors.Warning
                    Add-InstallationRecord -Software $item -Status "INSTALLED (with warnings - Exit Code: $exitCode)"
                }
            }
        }
        catch {
            Write-Host "[ERROR] Error installing $item : $($_.Exception.Message)" -ForegroundColor $Colors.Error
            Add-InstallationRecord -Software $item -Status "FAILED"
            Write-TKError -ScriptName 'conjure' -Message "Package install failed ('$item' via $PackageManager): $($_.Exception.Message)" -Category 'Package Install'
        }

        Start-Sleep -Seconds 1
    }

    Write-Progress -Activity "Installing $Type software" -Completed
}

# ---------------------------------------------------------------------------
# DIRECT DOWNLOAD INSTALLS
# ---------------------------------------------------------------------------
# Not everything a technician has to deploy exists in winget or Chocolatey.
# RMM agents in particular are almost always a per-tenant download link with a
# token in the URL, and there is no package ID to install by. These functions
# take a URL, fetch the installer, and run it silently.
#
# This path downloads an executable over the network and runs it elevated, so
# it verifies more than the package-manager path does: HTTPS is required, the
# file is identified by its own magic bytes rather than by a trusted URL, and
# a pinned SHA-256 (when supplied) is checked before anything is executed.

# Entries are { Name; Url; Args; Sha256 }. Validated once at collection time so
# PHASE 2 never has to stop and ask.
function New-DirectDownloadEntry {
    param(
        [Parameter(Mandatory)][string]$Url,
        [string]$Name,
        [string]$InstallArgs,
        [string]$Sha256
    )

    $trimmed = $Url.Trim()

    [Uri]$parsed = $null
    if (-not [Uri]::TryCreate($trimmed, [UriKind]::Absolute, [ref]$parsed)) {
        Write-Host "[ERROR] Not a valid URL, skipping: $trimmed" -ForegroundColor $Colors.Error
        return $null
    }

    # An installer fetched over plain HTTP can be swapped in transit and is then
    # run with Administrator rights. Refuse rather than warn.
    if ($parsed.Scheme -ne 'https') {
        Write-Host "[ERROR] Refusing a non-HTTPS installer URL: $trimmed" -ForegroundColor $Colors.Error
        Write-Host "        The file would be run elevated and cannot be trusted over $($parsed.Scheme)." -ForegroundColor $Colors.Error
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $leaf = [IO.Path]::GetFileName($parsed.AbsolutePath)
        $Name = if ([string]::IsNullOrWhiteSpace($leaf)) { $parsed.Host } else { $leaf }
    }

    if ($Sha256) {
        $Sha256 = ($Sha256 -replace '[^0-9A-Fa-f]', '')
        if ($Sha256.Length -ne 64) {
            Write-Host "[ERROR] SHA-256 for '$Name' is not 64 hex characters, skipping." -ForegroundColor $Colors.Error
            return $null
        }
    }

    return [PSCustomObject]@{
        Name   = $Name
        Url    = $trimmed
        Args   = $InstallArgs
        Sha256 = $Sha256
    }
}

# Direct downloads saved in config.json, so a per-tenant RMM agent is entered
# once and then deploys unattended on every machine afterwards.
function Get-ConfiguredDirectDownload {
    $entries = New-Object System.Collections.ArrayList

    try { $cfg = Get-TKConfig } catch { return $entries }
    if (-not $cfg.PSObject.Properties['Conjure']) { return $entries }
    if (-not $cfg.Conjure -or -not $cfg.Conjure.PSObject.Properties['DirectDownloads']) { return $entries }

    # Get-TKConfig fills absent nested keys with '', so coerce and drop blanks.
    foreach ($item in @($cfg.Conjure.DirectDownloads)) {
        if (-not $item -or -not $item.PSObject.Properties['Url'] -or [string]::IsNullOrWhiteSpace($item.Url)) { continue }
        $entry = New-DirectDownloadEntry -Url $item.Url `
                                         -Name   $(if ($item.PSObject.Properties['Name'])   { $item.Name }   else { '' }) `
                                         -InstallArgs $(if ($item.PSObject.Properties['Args'])   { $item.Args }   else { '' }) `
                                         -Sha256 $(if ($item.PSObject.Properties['Sha256']) { $item.Sha256 } else { '' })
        if ($entry) { [void]$entries.Add($entry) }
    }

    return $entries
}

function Read-DirectDownload {
    $entries = New-Object System.Collections.ArrayList

    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "DIRECT DOWNLOAD INSTALLS" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    Write-Host "For software with no winget or Chocolatey package — RMM agents," -ForegroundColor $Colors.Info
    Write-Host "vendor portals, per-tenant installers with a token in the link." -ForegroundColor $Colors.Info
    Write-Host "Paste a direct HTTPS link to an .exe or .msi." -ForegroundColor $Colors.Info
    Write-Host "Leave blank and press Enter to skip." -ForegroundColor $Colors.Info
    Write-Host ""

    while ($true) {
        $url = Read-Host "Installer URL"
        if ([string]::IsNullOrWhiteSpace($url)) { break }

        $name = Read-Host "  Display name for the log (blank = use the filename)"

        Write-Host "  Silent install arguments." -ForegroundColor $Colors.Info
        Write-Host "    MSI : /qn /norestart is applied automatically; add properties here (e.g. TOKEN=abc123)" -ForegroundColor $Colors.Info
        Write-Host "    EXE : varies by vendor — common ones are /S, /s, /silent, /quiet, /verysilent" -ForegroundColor $Colors.Info
        $instArgs = Read-Host "  Arguments (blank = none)"

        $sha = Read-Host "  Expected SHA-256 (blank = skip verification)"

        $entry = New-DirectDownloadEntry -Url $url -Name $name -InstallArgs $instArgs -Sha256 $sha
        if ($entry) {
            [void]$entries.Add($entry)
            Write-Host "[OK] Queued: $($entry.Name)" -ForegroundColor $Colors.Success
        }

        Write-Host ""
    }

    if ($entries.Count -eq 0) {
        Write-Host "[!!] No direct downloads queued" -ForegroundColor $Colors.Warning
    }

    return $entries
}

# Identify the payload from its own leading bytes rather than trusting the URL.
# RMM links routinely end in '?id=...' with no extension at all, and a link
# whose token has expired returns an HTML error page with HTTP 200 — which
# would otherwise be renamed .exe and executed.
function Get-InstallerKind {
    param([Parameter(Mandatory)][string]$Path)

    $buf = New-Object byte[] 8
    $read = 0
    $fs = [IO.File]::OpenRead($Path)
    try { $read = $fs.Read($buf, 0, 8) } finally { $fs.Dispose() }
    if ($read -lt 2) { return 'Unknown' }

    # MSI (and other OLE compound files): D0 CF 11 E0 A1 B1 1A E1
    $ole = @(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    if ($read -eq 8) {
        $match = $true
        for ($i = 0; $i -lt 8; $i++) { if ($buf[$i] -ne $ole[$i]) { $match = $false; break } }
        if ($match) { return 'MSI' }
    }

    # PE executable: 'MZ'
    if ($buf[0] -eq 0x4D -and $buf[1] -eq 0x5A) { return 'EXE' }

    return 'Unknown'
}

function Install-DirectDownload {
    param(
        [PSObject[]]$Entries
    )

    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "Installing Direct Downloads" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""

    $total = @($Entries).Count
    $i     = 0

    foreach ($entry in $Entries) {
        $i++
        Write-Progress -Activity "Installing direct downloads" `
                       -Status        "[$i/$total] $($entry.Name)" `
                       -PercentComplete ([math]::Floor($i / $total * 100))
        Write-Host "[$i/$total] $($entry.Name)" -ForegroundColor $Colors.Info
        Write-Host "        $($entry.Url)" -ForegroundColor $Colors.Info

        if ($WhatIf) {
            Write-Host "[~] WhatIf: would download and run it silently. Nothing was fetched." -ForegroundColor $Colors.Accent
            Add-InstallationRecord -Software $entry.Name -Status "WHATIF - not installed"
            continue
        }

        $workDir = Join-Path ([IO.Path]::GetTempPath()) ("conjure_" + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null

        try {
            # --- Fetch -------------------------------------------------------
            $payload = Join-Path $workDir 'payload.bin'
            Write-Host "        Downloading..." -ForegroundColor $Colors.Info
            try {
                [Net.ServicePointManager]::SecurityProtocol =
                    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

                # Invoke-WebRequest's own progress rendering slows large downloads
                # by orders of magnitude on 5.1; ours is enough.
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    Invoke-WebRequest -Uri $entry.Url -OutFile $payload -UseBasicParsing -ErrorAction Stop
                }
                finally { $ProgressPreference = $prevProgress }
            }
            catch {
                Write-Host "[ERROR] $($entry.Name) — download failed: $($_.Exception.Message)" -ForegroundColor $Colors.Error
                Add-InstallationRecord -Software $entry.Name -Status "FAILED - download failed"
                Write-TKError -ScriptName 'conjure' -Message "Direct download failed ('$($entry.Name)' from '$($entry.Url)'): $($_.Exception.Message)" -Category 'Package Install'
                continue
            }

            $sizeMb = [math]::Round((Get-Item $payload).Length / 1MB, 2)
            Write-Host "        Downloaded ${sizeMb} MB" -ForegroundColor $Colors.Info

            # --- Verify ------------------------------------------------------
            $actualHash = (Get-FileHash -Path $payload -Algorithm SHA256).Hash
            Write-Host "        SHA-256: $actualHash" -ForegroundColor $Colors.Info

            if ($entry.Sha256) {
                if ($actualHash -ne $entry.Sha256.ToUpperInvariant()) {
                    Write-Host "[ERROR] $($entry.Name) — SHA-256 mismatch. Nothing was run." -ForegroundColor $Colors.Error
                    Write-Host "        expected $($entry.Sha256.ToUpperInvariant())" -ForegroundColor $Colors.Error
                    Add-InstallationRecord -Software $entry.Name -Status "FAILED - SHA-256 mismatch"
                    Write-TKError -ScriptName 'conjure' -Message "Direct download hash mismatch ('$($entry.Name)' from '$($entry.Url)'): expected $($entry.Sha256), got $actualHash" -Category 'Package Install'
                    continue
                }
                Write-Host "        Hash verified against the pinned value." -ForegroundColor $Colors.Success
            }

            $kind = Get-InstallerKind -Path $payload
            if ($kind -eq 'Unknown') {
                Write-Host "[ERROR] $($entry.Name) — the download is neither an EXE nor an MSI." -ForegroundColor $Colors.Error
                Write-Host "        A link whose token has expired often returns an error page with HTTP 200." -ForegroundColor $Colors.Error
                Add-InstallationRecord -Software $entry.Name -Status "FAILED - not an installer"
                Write-TKError -ScriptName 'conjure' -Message "Direct download for '$($entry.Name)' is not a PE or MSI payload (from '$($entry.Url)')" -Category 'Package Install'
                continue
            }

            # --- Run ---------------------------------------------------------
            $installer = Join-Path $workDir ("installer." + $kind.ToLowerInvariant())
            Rename-Item -Path $payload -NewName (Split-Path -Leaf $installer) -Force

            $startTime = Get-Date
            if ($kind -eq 'MSI') {
                $argList = "/i `"$installer`" /qn /norestart"
                if ($entry.Args) { $argList = "$argList $($entry.Args)" }
                Write-Host "        Running: msiexec $argList" -ForegroundColor $Colors.Info
                $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $argList -Wait -PassThru -ErrorAction Stop
            }
            else {
                if ($entry.Args) {
                    Write-Host "        Running: installer.exe $($entry.Args)" -ForegroundColor $Colors.Info
                    $proc = Start-Process -FilePath $installer -ArgumentList $entry.Args -Wait -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host "[!!] No silent arguments given — this installer may open its window" -ForegroundColor $Colors.Warning
                    Write-Host "     and the run will wait until it is closed." -ForegroundColor $Colors.Warning
                    $proc = Start-Process -FilePath $installer -Wait -PassThru -ErrorAction Stop
                }
            }

            $elapsed = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
            $verdict = Resolve-InstallExit -ExitCode $proc.ExitCode

            switch ($verdict.Class) {
                'Success' {
                    Write-Host "[OK] $($entry.Name) — $($verdict.Reason) (${elapsed}s)" -ForegroundColor $Colors.Success
                    Add-InstallationRecord -Software $entry.Name -Status "INSTALLED (direct download, SHA-256 $actualHash)"
                }
                'Failed' {
                    Write-Host "[ERROR] $($entry.Name) failed — $($verdict.Reason) [exit code $($proc.ExitCode)]" -ForegroundColor $Colors.Error
                    Add-InstallationRecord -Software $entry.Name -Status "FAILED - $($verdict.Reason)"
                    Write-TKError -ScriptName 'conjure' -Message "Direct download install failed ('$($entry.Name)'): $($verdict.Reason) [exit code $($proc.ExitCode)]" -Category 'Package Install'
                }
                default {
                    Write-Host "[!!] $($entry.Name) — $($verdict.Reason)" -ForegroundColor $Colors.Warning
                    Add-InstallationRecord -Software $entry.Name -Status "INSTALLED (with warnings - Exit Code: $($proc.ExitCode))"
                }
            }
        }
        catch {
            Write-Host "[ERROR] $($entry.Name) — $($_.Exception.Message)" -ForegroundColor $Colors.Error
            Add-InstallationRecord -Software $entry.Name -Status "FAILED"
            Write-TKError -ScriptName 'conjure' -Message "Direct download install failed ('$($entry.Name)'): $($_.Exception.Message)" -Category 'Package Install'
        }
        finally {
            # The installer is a downloaded executable; do not leave it on disk.
            Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Progress -Activity "Installing direct downloads" -Completed
}

function Update-AllSoftware {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "Running Package Updates" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""

    if ($WhatIf) {
        $cmd = if ($script:PackageManager -eq "chocolatey") { "choco upgrade all -y" }
               else { "winget upgrade --all -h" }
        Write-Host "[~] WhatIf: would run  $cmd" -ForegroundColor $Colors.Accent
        Add-InstallationRecord -Software "All packages" -Status "WHATIF - not upgraded"
        return
    }

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
            $verdict = Resolve-InstallExit -ExitCode $exitCode
            Write-Host "[!!] Package update finished at $updateTime — $($verdict.Reason)" -ForegroundColor $Colors.Warning
            Add-InstallationRecord -Software "All Packages" -Status "UPDATED (with warnings - $($verdict.Reason))"
        }
    }
    catch {
        Write-Host "[ERROR] Error during update: $($_.Exception.Message)" -ForegroundColor $Colors.Error
        Add-InstallationRecord -Software "All Packages" -Status "UPDATE FAILED"
    }
}

function Select-AdobeEdition {
    # Prompts for the Adobe Acrobat edition and returns the package ID to install
    # for the currently selected package manager. Reader and Pro are both available
    # on winget; Chocolatey's community repo has no Acrobat Pro package, so a Pro
    # request there falls back to Reader with an explanatory note.
    $isChoco = ($script:PackageManager -eq "chocolatey")

    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "ADOBE ACROBAT EDITION" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    Write-Host "  [1] Adobe Acrobat Reader (free)" -ForegroundColor $Colors.Info
    Write-Host "  [2] Adobe Acrobat Pro (licensed)" -ForegroundColor $Colors.Info
    Write-Host ""

    $choice = Read-Host "Enter your choice (1/2)"

    switch ($choice) {
        "2" {
            if ($isChoco) {
                Write-Host "[!!] Chocolatey's community repo has no Adobe Acrobat Pro package." -ForegroundColor $Colors.Warning
                Write-Host "     Installing Adobe Acrobat Reader instead. If the user holds an Acrobat Pro" -ForegroundColor $Colors.Warning
                Write-Host "     license, have them sign in to Reader and run the in-app upgrade to activate Pro." -ForegroundColor $Colors.Warning
                return (Get-PackageId -Entry $script:AdobeReader)
            }
            Write-Host "[OK] Selected: Adobe Acrobat Pro" -ForegroundColor $Colors.Success
            return (Get-PackageId -Entry $script:AdobePro)
        }
        "1" {
            Write-Host "[OK] Selected: Adobe Acrobat Reader" -ForegroundColor $Colors.Success
            return (Get-PackageId -Entry $script:AdobeReader)
        }
        default {
            Write-Host "[!!] Invalid choice. Defaulting to Adobe Acrobat Reader" -ForegroundColor $Colors.Warning
            return (Get-PackageId -Entry $script:AdobeReader)
        }
    }
}

function Read-CustomPackages {
    # Lets the operator add their own package IDs on top of the curated optional
    # list. IDs are package-manager-specific, so whatever is typed is passed
    # through verbatim to the selected manager (winget -e --id <id> / choco install <id>).
    $custom = New-Object System.Collections.ArrayList

    $mgrName = if ($script:PackageManager -eq "chocolatey") { "Chocolatey" } else { "Winget" }
    $example = if ($script:PackageManager -eq "chocolatey") { "notepadplusplus, vlc" } else { "Notepad++.Notepad++, VideoLAN.VLC" }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host "CUSTOM PACKAGES" -ForegroundColor $Colors.Header
    Write-Host "========================================" -ForegroundColor $Colors.Header
    Write-Host ""
    Write-Host "Enter any additional $mgrName package IDs to install." -ForegroundColor $Colors.Info
    Write-Host "Separate multiple IDs with commas (e.g. $example)." -ForegroundColor $Colors.Info
    Write-Host "Leave blank and press Enter to skip." -ForegroundColor $Colors.Info
    Write-Host ""

    $userInput = Read-Host "Custom $mgrName package IDs"

    if (-not [string]::IsNullOrWhiteSpace($userInput)) {
        $ids = $userInput -split ','
        foreach ($id in $ids) {
            $trimmed = $id.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                [void]$custom.Add($trimmed)
            }
        }
    }

    if ($custom.Count -gt 0) {
        $customDisplay = $custom -join ', '
        Write-Host "[OK] Custom packages queued: $customDisplay" -ForegroundColor $Colors.Success
    }
    else {
        Write-Host "[!!] No custom packages entered" -ForegroundColor $Colors.Warning
    }

    return $custom
}

function Select-OptionalSoftware {
    param(
        [string[]]$SoftwareList = @()
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
            $userInput = Read-Host "Enter numbers (comma-separated, e.g., 1,2)"

            if (-not [string]::IsNullOrWhiteSpace($userInput)) {
                $numbers = $userInput -split ','
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
    # Preview rows are counted on their own. "WHATIF - not installed" contains
    # the word "installed", so a contains-match would report a dry run as a
    # successful deployment — the one number nobody can afford to misread.
    $whatIfCount = ($InstallationLog | Where-Object { $_.Status -like "WHATIF*" }    | Measure-Object).Count
    $live        =  $InstallationLog | Where-Object { $_.Status -notlike "WHATIF*" }

    $successCount = ($live | Where-Object { $_.Status -like "INSTALLED*" } | Measure-Object).Count
    $failCount    = ($live | Where-Object { $_.Status -like "*FAILED*" }   | Measure-Object).Count
    $updateCount  = ($live | Where-Object { $_.Status -like "*UPDATED*" }  | Measure-Object).Count

    Write-Host "Summary: Installed: $successCount | Failed: $failCount | Updated: $updateCount" -ForegroundColor $Colors.Header
    if ($whatIfCount -gt 0) {
        Write-Host "WhatIf: $whatIfCount previewed, nothing was installed or upgraded." -ForegroundColor $Colors.Accent
    }
    Write-Host ""
}

# ===========================
# MAIN EXECUTION
# ===========================

if (-not $Unattended) { Show-Banner }

# ---------------------------------------------------------------------------
# PHASE 1 — OPERATOR INPUT (front-loaded)
# All prompts are gathered here so the long-running install/upgrade work in
# PHASE 2 can run start-to-finish without anyone babysitting the machine.
# ---------------------------------------------------------------------------

# Select package manager
if (-not $Unattended) { Select-PackageManager }

# Choose operation
if ($Unattended) {
    Write-Host "[*] Unattended mode: installing required packages only" -ForegroundColor $Colors.Info
    $operation = "1"
}
else {
    Write-Host ""
    Write-Host "Choose operation:" -ForegroundColor $Colors.Header
    Write-Host "  [1] Install software" -ForegroundColor $Colors.Info
    Write-Host "  [2] Upgrade all software" -ForegroundColor $Colors.Info
    Write-Host "  [3] Install and then Upgrade" -ForegroundColor $Colors.Info
    Write-Host ""
    $operation = Read-Host "Enter your choice (1/2/3)"
}

if ($operation -notin @("1", "2", "3")) {
    Write-Host "[ERROR] Invalid choice. Exiting." -ForegroundColor $Colors.Error
    if (-not $Unattended) { Read-Host "Press Enter to exit" }
    exit 1
}

$doInstall = ($operation -eq "1" -or $operation -eq "3")
$doUpgrade = ($operation -eq "2" -or $operation -eq "3")

# Resolve the catalog to package IDs for the selected manager
$ActiveRequired = @($RequiredCatalog | ForEach-Object { Get-PackageId -Entry $_ })
$ActiveOptional = @($OptionalCatalog | ForEach-Object { Get-PackageId -Entry $_ })

$optionalList = New-Object System.Collections.ArrayList
$customList   = New-Object System.Collections.ArrayList

if ($doInstall) {
    Write-Host "[OK] Selected: Install Mode" -ForegroundColor $Colors.Success

    # Adobe Acrobat edition (Reader or Pro) — appended to the required list
    if ($Unattended) {
        $AdobePackage = Get-PackageId -Entry $AdobeReader
    }
    else {
        $AdobePackage = Select-AdobeEdition
    }
    $ActiveRequired += $AdobePackage

    # Optional software selection + operator-supplied custom IDs (interactive mode only)
    if (-not $Unattended) {
        $optionalList = Select-OptionalSoftware -SoftwareList $ActiveOptional
        $customList   = Read-CustomPackages
    }
}

# Direct downloads, gathered here in PHASE 1 like every other prompt. Sources
# stack: anything saved in config.json, plus -DirectUrl, plus whatever the
# operator pastes in. All of it is validated now so PHASE 2 never stops to ask.
$directList = New-Object System.Collections.ArrayList

if ($doInstall) {
    foreach ($entry in (Get-ConfiguredDirectDownload)) {
        [void]$directList.Add($entry)
        Write-Host "[OK] Direct download from config.json: $($entry.Name)" -ForegroundColor $Colors.Success
    }

    foreach ($url in @($DirectUrl)) {
        if ([string]::IsNullOrWhiteSpace($url)) { continue }
        $entry = New-DirectDownloadEntry -Url $url
        if ($entry) {
            [void]$directList.Add($entry)
            Write-Host "[OK] Direct download from -DirectUrl: $($entry.Name)" -ForegroundColor $Colors.Success
        }
    }

    if (-not $Unattended) {
        foreach ($entry in (Read-DirectDownload)) { [void]$directList.Add($entry) }
    }
}

if ($doUpgrade -and -not $doInstall) {
    Write-Host "[OK] Selected: Upgrade Mode" -ForegroundColor $Colors.Success
}
elseif ($doUpgrade) {
    Write-Host "[OK] Will upgrade all packages after install" -ForegroundColor $Colors.Success
}

Write-Host ""
Write-Host "[*] All selections captured — beginning unattended execution. No further input required." -ForegroundColor $Colors.Accent

# ---------------------------------------------------------------------------
# PHASE 2 — EXECUTION (no operator prompts)
# ---------------------------------------------------------------------------

# Check that the selected package manager is available. A missing manager is
# only fatal when there is manager work to do — direct downloads need neither
# winget nor Chocolatey, so an RMM agent still deploys on a machine where the
# package manager is broken or absent.
$mgrName = if ($PackageManager -eq "chocolatey") { "Chocolatey" } else { "Winget" }
$mgrAvailable = if ($PackageManager -eq "chocolatey") { Test-ChocolateyAvailable } else { Test-WingetAvailable }

if (-not $mgrAvailable) {
    if ($directList.Count -gt 0) {
        Write-Host "[!!] $mgrName is unavailable — skipping package installs and upgrades." -ForegroundColor $Colors.Warning
        Write-Host "     Direct downloads need no package manager and will still run." -ForegroundColor $Colors.Warning
        $doInstall = $false
        $doUpgrade = $false
    }
    else {
        Write-Host "[ERROR] Cannot proceed without $mgrName" -ForegroundColor $Colors.Error
        if (-not $Unattended) { Read-Host "Press Enter to exit" }
        exit 1
    }
}

if ($doInstall) {
    # Install required software
    Install-Software -SoftwareList $ActiveRequired -Type "Required"

    # Install any optional software selected up front
    if ($optionalList.Count -gt 0) {
        Install-Software -SoftwareList $optionalList -Type "Optional"
    }

    # Install any operator-supplied custom package IDs
    if ($customList.Count -gt 0) {
        Install-Software -SoftwareList $customList -Type "Custom"
    }
}

# Runs independently of the package manager, so it sits outside the $doInstall
# guard above — which is cleared when the manager is missing.
if ($directList.Count -gt 0) {
    Install-DirectDownload -Entries $directList.ToArray()
}

if ($doUpgrade) {
    Update-AllSoftware
}

# Show summary
Show-InstallationSummary

Write-Host "[OK] C.O.N.J.U.R.E. Script completed!" -ForegroundColor $Colors.Success
Write-Host ""
if (-not $Unattended) { Read-Host "Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
