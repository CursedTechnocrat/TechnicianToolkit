# runepress.ps1 - R.U.N.E.P.R.E.S.S. — Remote Utility for Networked Equipment — Printer Registration, Extraction & Silent Setup
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors
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
    R.U.N.E.P.R.E.S.S. — Remote Utility for Networked Equipment — Printer Registration, Extraction & Silent Setup
    Printer Driver Installation & Configuration Tool for PowerShell 5.1+

.DESCRIPTION
    Automates printer driver extraction, installation, and network printer
    configuration via a command-line interface. Supports ZIP, EXE and MSI
    driver packages as well as already-extracted folders containing a bare
    INF, plus TCP/IP or UNC port configuration. Generates a timestamped CSV
    installation log.

    INF-based installs are staged with pnputil and then registered with the
    print spooler via Add-PrinterDriver. Both steps are required: pnputil
    alone populates the DriverStore but leaves the driver invisible to
    Get-PrinterDriver and to the Add-Printer driver list.

.USAGE
    PS C:\> .\runepress.ps1                                  # Must be run as Administrator
    PS C:\> .\runepress.ps1 -DriverPath 'C:\Drivers\KM_UPD'  # Install from an extracted driver folder
    PS C:\> .\runepress.ps1 -DriverPath 'C:\Drivers\x.inf'   # Install one specific INF
    PS C:\> .\runepress.ps1 -Unattended                      # Silent mode — auto-selects first INF, skips printer config

.NOTES
    Version : 5.0

#>

param(
    [string]$DriverPath,
    [switch]$Unattended,
    [switch]$Transcript,
    [switch]$WhatIf
)

# ===========================
# ADMIN PRIVILEGE CHECK
# ===========================
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
Invoke-AdminElevation -ScriptFile $PSCommandPath

# ===========================
# SCRIPT INITIALIZATION
# ===========================

# Resolve script execution path
if ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
}
elseif ($MyInvocation.MyCommand.Path) {
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $ScriptPath = (Get-Location).Path
}

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# Driver source. -DriverPath lets a technician point RUNEPRESS at a download
# folder (or one specific file) instead of copying the package next to the
# script. Falls back to the script's own directory, the historical behaviour.
$SourceFile = $null
if ($DriverPath) {
    if (-not (Test-Path -LiteralPath $DriverPath)) {
        Write-Host "  [!!] -DriverPath not found: $DriverPath" -ForegroundColor Red
        exit 1
    }
    $resolved = Get-Item -LiteralPath $DriverPath
    if ($resolved.PSIsContainer) {
        $SourcePath = $resolved.FullName
    }
    else {
        $SourcePath = $resolved.DirectoryName
        $SourceFile = $resolved
    }
}
else {
    $SourcePath = $ScriptPath
}

# Architecture decoration used in INF [Manufacturer] lines for this machine.
$InfArchDecoration = switch ($env:PROCESSOR_ARCHITECTURE) {
    'AMD64' { 'NTamd64' }
    'ARM64' { 'NTARM64' }
    default { 'NTx86' }
}

# Initialize global variables
$ExtractRoot     = Join-Path $ScriptPath "ExtractedDrivers"
$InstallationLog = @()

# Driver names registered with the spooler during this run — offered first
# when the technician picks a driver for a new printer.
$InstalledDriverNames = @()

# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-Banner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██╗   ██╗███╗   ██╗███████╗██████╗ ██████╗ ███████╗███████╗
  ██╔══██╗██║   ██║████╗  ██║██╔════╝██╔══██╗██╔══██╗██╔════╝██╔════╝
  ██████╔╝██║   ██║██╔██╗ ██║█████╗  ██████╔╝██████╔╝█████╗  ███████╗
  ██╔══██╗██║   ██║██║╚██╗██║██╔══╝  ██╔═══╝ ██╔══██╗██╔══╝  ╚════██╗
  ██║  ██║╚██████╔╝██║ ╚████║███████╗██║     ██║  ██║███████╗███████║
  ╚═╝  ╚═╝ ╚═════╝ ╚═╝  ╚═══╝╚══════╝╚═╝     ╚═╝  ╚═╝╚══════╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    R.U.N.E.P.R.E.S.S. - Remote Utility for Networked Equipment" -ForegroundColor Cyan
    Write-Host "    Printer Registration, Extraction and Silent Setup" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Script Location: $ScriptPath" -ForegroundColor Gray
    Write-Host "    Execution Time:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host ""
}

# ===========================
# DISPLAY DRIVER PREP INSTRUCTIONS
# ===========================

function Show-DriverPrepInstructions {
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Step 1: Driver Preparation" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor Yellow
    Write-Host "  1. Download the printer driver from the manufacturer website" -ForegroundColor White
    Write-Host "  2. Save the file to this location:" -ForegroundColor White
    Write-Host "     $SourcePath" -ForegroundColor Green
    if (-not $DriverPath) {
        Write-Host "     (or re-run with -DriverPath to use a folder you already have)" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "Supported formats:" -ForegroundColor Yellow
    Write-Host "  * Extracted driver folders containing an INF (.inf)  <- preferred" -ForegroundColor White
    Write-Host "  * ZIP archives (.zip)" -ForegroundColor White
    Write-Host "  * Executable installers (.exe)" -ForegroundColor White
    Write-Host "  * Windows Installer packages (.msi)" -ForegroundColor White
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
}

# ===========================
# PROMPT USER TO PLACE DRIVER
# ===========================

function Wait-ForDriverFile {
    while ($true) {
        Write-Host "Ready to proceed? (Y/N/Q)" -ForegroundColor Yellow
        Write-Host "  Y = Continue with installation" -ForegroundColor Gray
        Write-Host "  N = Go back and check folder" -ForegroundColor Gray
        Write-Host "  Q = Quit" -ForegroundColor Gray
        Write-Host ""

        $response = Read-Host "Enter choice"

        switch ($response.ToUpper()) {
            "Y" {
                return $true
            }
            "N" {
                Write-Host ""
                Show-DriverPrepInstructions
            }
            "Q" {
                Write-Host ""
                Write-Host "WARNING: Script terminated by user." -ForegroundColor Yellow
                exit 0
            }
            default {
                Write-Host "Invalid input. Please enter Y, N, or Q." -ForegroundColor Red
                Write-Host ""
            }
        }
    }
}

# ===========================
# LOCATE DRIVER FILES
# ===========================

function Get-InfPrinterModel {
    <#
        Parses a printer INF and returns the model names it declares, plus
        whether it targets this machine's architecture.

        Windows needs the model name (the quoted left-hand side of a line in a
        [Manufacturer]-referenced models section) to register a staged driver
        with the spooler via Add-PrinterDriver. There is no API that recovers
        it from the DriverStore afterwards, so it has to come from the INF.
    #>
    param(
        [Parameter(Mandatory)][string]$InfFile
    )

    $result = [PSCustomObject]@{
        Models              = @()
        MatchesArchitecture = $false
    }

    try   { $rawLines = Get-Content -LiteralPath $InfFile -ErrorAction Stop }
    catch { return $result }

    # Split into sections, dropping blank lines and comments. A ';' only opens
    # a comment outside of a quoted string — model names may legitimately
    # contain one.
    $sections = @{}
    $current  = $null
    foreach ($raw in $rawLines) {
        $line    = $raw
        $inQuote = $false
        for ($i = 0; $i -lt $line.Length; $i++) {
            if     ($line[$i] -eq '"') { $inQuote = -not $inQuote }
            elseif ($line[$i] -eq ';' -and -not $inQuote) { $line = $line.Substring(0, $i); break }
        }
        $line = $line.Trim()
        if (-not $line) { continue }

        if ($line -match '^\[(.+?)\]$') {
            $current = $Matches[1].Trim().ToLowerInvariant()
            if (-not $sections.ContainsKey($current)) { $sections[$current] = @() }
            continue
        }
        if ($current) { $sections[$current] += $line }
    }

    if (-not $sections.ContainsKey('manufacturer')) { return $result }

    # [Strings] backs the %token% indirection used by most vendor INFs.
    $strings = @{}
    if ($sections.ContainsKey('strings')) {
        foreach ($line in $sections['strings']) {
            if ($line -match '^([^=]+?)\s*=\s*(.+)$') {
                $strings[$Matches[1].Trim().ToLowerInvariant()] = $Matches[2].Trim().Trim('"')
            }
        }
    }
    $resolve = {
        param($text)
        $t = $text.Trim().Trim('"')
        if ($t -match '^%(.+)%$') {
            $key = $Matches[1].ToLowerInvariant()
            if ($strings.ContainsKey($key)) { return $strings[$key] }
        }
        return $t
    }

    $models = New-Object System.Collections.Generic.List[string]
    $archOk = $false

    foreach ($line in $sections['manufacturer']) {
        if ($line -notmatch '=') { continue }

        $parts = ($line -split '=', 2)[1] -split ',' |
                 ForEach-Object { $_.Trim() } |
                 Where-Object   { $_ }
        if (-not $parts) { continue }

        $base  = & $resolve $parts[0]
        $decos = @($parts | Select-Object -Skip 1)

        # 'NTamd64' also covers OS-versioned forms such as 'NTamd64.6.0'.
        $archDecos = @($decos | Where-Object { $_ -like "$InfArchDecoration*" })
        if     ($archDecos) { $archOk = $true }
        elseif (-not $decos) { $archOk = $true }   # undecorated INF applies everywhere

        # Prefer the architecture-specific sections; fall back to whatever the
        # INF declares so a mismatched package still yields a name to report.
        $useDecos   = if ($archDecos) { $archDecos } else { $decos }
        $candidates = @($base)
        foreach ($d in $useDecos) { $candidates += "$base.$d" }

        foreach ($c in $candidates) {
            $key = $c.ToLowerInvariant()
            if (-not $sections.ContainsKey($key)) { continue }
            foreach ($entry in $sections[$key]) {
                if ($entry -notmatch '=') { continue }
                $name = & $resolve ($entry -split '=', 2)[0]
                if ($name -and -not $models.Contains($name)) { [void]$models.Add($name) }
            }
        }
    }

    $result.Models              = $models.ToArray()
    $result.MatchesArchitecture = $archOk
    return $result
}

function Find-DriverFiles {
    # An explicitly named file wins outright.
    if ($SourceFile) { return @($SourceFile) }

    # Prefer an already-extracted package. If the tree holds an INF targeting
    # this machine's architecture, install from that directly rather than
    # running a vendor bootstrapper that may ignore silent switches and leave
    # nothing registered.
    $InfFiles = @(
        Get-ChildItem -LiteralPath $SourcePath -Filter '*.inf' -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -notlike "$ExtractRoot*" }
    )
    $ArchInfs = @($InfFiles | Where-Object { (Get-InfPrinterModel -InfFile $_.FullName).MatchesArchitecture })
    if ($ArchInfs.Count -gt 0) { return $ArchInfs }

    $DriverFiles = @(
        Get-ChildItem -LiteralPath $SourcePath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '^\.(zip|exe|msi)$' } |
            Where-Object { $_.Name -ne (Split-Path -Leaf $PSCommandPath) }
    )

    return $DriverFiles
}

# ===========================
# INF INSTALL — STAGE + REGISTER
# ===========================

function Select-InfModel {
    param(
        [Parameter(Mandatory)][string[]]$Models
    )

    if (@($Models).Count -eq 1) { return $Models }

    if ($Unattended) {
        Write-Host "    [*] INF declares $(@($Models).Count) models - registering all (unattended)." -ForegroundColor Gray
        return $Models
    }

    Write-Host ""
    Write-Host "This INF declares $(@($Models).Count) printer models:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Models.Count; $i++) {
        Write-Host "  [$($i + 1)] $($Models[$i])" -ForegroundColor White
    }
    Write-Host "  [A] Register all" -ForegroundColor White
    Write-Host ""

    do {
        $sel = Read-Host "Select model to register (1-$($Models.Count), or A for all)"
        if ($sel -match '^[Aa]$') { return $Models }
    } while (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $Models.Count)

    return @($Models[[int]$sel - 1])
}

function Invoke-InfInstall {
    <#
        Installs a printer driver from an INF in the two steps Windows requires:

          1. pnputil /add-driver ... /install  — stages into the DriverStore
          2. Add-PrinterDriver -Name <model>   — registers with the print spooler

        Step 2 is not optional. pnputil reports success and exits 0 having done
        only step 1, at which point Get-PrinterDriver still does not list the
        driver and Add-Printer cannot reference it by name.
    #>
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$InfFile,
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string]$SourceType
    )

    $Models = @((Get-InfPrinterModel -InfFile $InfFile.FullName).Models)

    if ($WhatIf) {
        Write-Host "[~] WhatIf: would run  pnputil /add-driver `"$($InfFile.FullName)`" /install" -ForegroundColor Cyan
        if ($Models) {
            Write-Host "[~] WhatIf: would then register with the spooler (model names parsed from the INF):" -ForegroundColor Cyan
            foreach ($m in $Models) { Write-Host "      Add-PrinterDriver -Name `"$m`"" -ForegroundColor Cyan }
        }
        else {
            Write-Host "[~] WhatIf: no model name could be parsed - spooler registration would be skipped." -ForegroundColor Yellow
        }
        $script:InstallationLog += [PSCustomObject]@{
            File = $SourceName; Type = $SourceType; INF = $InfFile.Name
            Status = "WhatIf (not installed)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $true
    }

    # --- Step 1: stage into the DriverStore ---------------------------------
    Write-Host "Staging driver package (pnputil)..." -ForegroundColor Yellow
    try {
        $PnpResult = & pnputil /add-driver "$($InfFile.FullName)" /install 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: pnputil failed (exit code $LASTEXITCODE)." -ForegroundColor Red
            Write-Host ($PnpResult | Out-String) -ForegroundColor Red
            $script:InstallationLog += [PSCustomObject]@{
                File = $SourceName; Type = $SourceType; INF = $InfFile.Name
                Status = "Failed (pnputil exit $LASTEXITCODE)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            return $false
        }
        Write-Host "OK: Driver package staged in the DriverStore." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to run pnputil: $($_.Exception.Message)" -ForegroundColor Red
        Write-TKError -ScriptName 'runepress' -Message "pnputil failed for '$($InfFile.FullName)' (from '$SourceName'): $($_.Exception.Message)" -Category 'Printer Driver Install'
        $script:InstallationLog += [PSCustomObject]@{
            File = $SourceName; Type = $SourceType; INF = $InfFile.Name
            Status = "Failed (pnputil error)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $false
    }

    # --- Step 2: register with the print spooler ----------------------------
    if (-not $Models) {
        Write-Host "WARNING: No model name could be parsed from $($InfFile.Name)." -ForegroundColor Yellow
        Write-Host "         The driver is staged but NOT registered with the spooler," -ForegroundColor Yellow
        Write-Host "         so it will not appear in the printer driver list." -ForegroundColor Yellow
        Write-Host "         Register it manually with:  Add-PrinterDriver -Name '<model>'" -ForegroundColor Yellow
        $script:InstallationLog += [PSCustomObject]@{
            File = $SourceName; Type = $SourceType; INF = $InfFile.Name
            Status = "Staged only (no model name in INF)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $false
    }

    $Selected = @(Select-InfModel -Models $Models)

    Write-Host "Registering with the print spooler..." -ForegroundColor Yellow
    $registered = @()
    foreach ($model in $Selected) {
        if (Get-PrinterDriver -Name $model -ErrorAction SilentlyContinue) {
            Write-Host "  = $model (already registered)" -ForegroundColor Gray
            $registered += $model
            continue
        }
        try {
            Add-PrinterDriver -Name $model -ErrorAction Stop
            Write-Host "  + $model" -ForegroundColor Green
            $registered += $model
        }
        catch {
            Write-Host "  ! $model - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if (-not $registered) {
        Write-Host "ERROR: Driver staged, but no model could be registered with the spooler." -ForegroundColor Red
        Write-TKError -ScriptName 'runepress' -Message "Driver from '$SourceName' staged via pnputil but Add-PrinterDriver failed for every model in '$($InfFile.Name)'." -Category 'Printer Driver Install'
        $script:InstallationLog += [PSCustomObject]@{
            File = $SourceName; Type = $SourceType; INF = $InfFile.Name
            Status = "Failed (spooler registration)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $false
    }

    $script:InstalledDriverNames = @(@($script:InstalledDriverNames) + $registered | Select-Object -Unique)

    Write-Host "OK: Driver installed and registered with the spooler." -ForegroundColor Green
    $script:InstallationLog += [PSCustomObject]@{
        File = $SourceName; Type = $SourceType; INF = $InfFile.Name
        Status = "Success ($(@($registered).Count) model(s) registered)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
    return $true
}

function Install-InfDriver {
    param(
        [System.IO.FileInfo]$InfFile
    )

    Write-Host ""
    Write-Host "Processing INF: $($InfFile.Name)" -ForegroundColor Cyan
    Write-Host "  $($InfFile.FullName)" -ForegroundColor Gray
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

    return (Invoke-InfInstall -InfFile $InfFile -SourceName $InfFile.Name -SourceType 'INF')
}

# ===========================
# INSTALL ZIP DRIVERS
# ===========================

function Install-ZipDriver {
    param(
        [System.IO.FileInfo]$ZipFile
    )

    $DriverName  = [System.IO.Path]::GetFileNameWithoutExtension($ZipFile.Name)
    $ExtractPath = Join-Path $ExtractRoot $DriverName

    Write-Host ""
    Write-Host "Processing ZIP: $($ZipFile.Name)" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

    # Clean previous extraction
    if (Test-Path $ExtractPath) {
        Write-Host "Removing previous extraction directory..." -ForegroundColor Yellow
        try {
            Remove-Item $ExtractPath -Recurse -Force -ErrorAction Stop
            Write-Host "OK: Previous extraction removed." -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Could not remove directory: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    # Extract ZIP archive
    Write-Host "Extracting: $($ZipFile.Name)..." -ForegroundColor Yellow
    try {
        Expand-Archive -Path $ZipFile.FullName -DestinationPath $ExtractPath -Force -ErrorAction Stop
        Write-Host "OK: Extraction complete." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Locate INF files in extracted content
    Write-Host "Searching for INF driver files..." -ForegroundColor Yellow
    $InfFiles = Get-ChildItem -Path $ExtractPath -Filter "*.inf" -Recurse

    if (-not $InfFiles) {
        Write-Host "ERROR: No INF files found in extracted content." -ForegroundColor Red
        return $false
    }

    Write-Host "Found $($InfFiles.Count) INF file(s)." -ForegroundColor Green

    # Select INF - prompt if multiple found
    if ($InfFiles.Count -eq 1) {
        $SelectedInf = $InfFiles[0]
    }
    elseif ($Unattended) {
        Write-Host "    [*] Multiple INFs found — auto-selecting first: $($InfFiles[0].Name)" -ForegroundColor Gray
        $SelectedInf = $InfFiles[0]
    }
    else {
        Write-Host ""
        Write-Host "Multiple INF files found. Select one to install:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $InfFiles.Count; $i++) {
            Write-Host "  [$($i + 1)] $($InfFiles[$i].FullName)" -ForegroundColor White
        }
        Write-Host ""
        do {
            $selection = Read-Host "Enter number (1-$($InfFiles.Count))"
        } while (-not ($selection -match '^\d+$') -or [int]$selection -lt 1 -or [int]$selection -gt $InfFiles.Count)
        $SelectedInf = $InfFiles[[int]$selection - 1]
    }

    Write-Host "Using INF: $($SelectedInf.FullName)" -ForegroundColor Cyan

    return (Invoke-InfInstall -InfFile $SelectedInf -SourceName $ZipFile.Name -SourceType 'ZIP')
}

# ===========================
# INSTALL EXE DRIVERS
# ===========================

function Get-SpoolerDriverName {
    return @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
}

function Install-ExeDriver {
    param(
        [System.IO.FileInfo]$ExeFile
    )

    Write-Host ""
    Write-Host "Processing EXE: $($ExeFile.Name)" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

    if ($WhatIf) {
        Write-Host "[~] WhatIf: would run  $($ExeFile.FullName) /S /silent /quiet /norestart" -ForegroundColor Cyan
        $script:InstallationLog += [PSCustomObject]@{
            File = $ExeFile.Name; Type = "EXE"; INF = "N/A"
            Status = "WhatIf (not installed)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $true
    }

    Write-Host "Running silent installer..." -ForegroundColor Yellow
    $DriversBefore = Get-SpoolerDriverName

    try {
        $Process = Start-Process -FilePath $ExeFile.FullName `
            -ArgumentList "/S /silent /quiet /norestart" `
            -Wait -PassThru -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to launch installer: $($_.Exception.Message)" -ForegroundColor Red
        $script:InstallationLog += [PSCustomObject]@{
            File   = $ExeFile.Name
            Type   = "EXE"
            INF    = "N/A"
            Status = "Failed (launch error)"
            Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        Write-TKError -ScriptName 'runepress' -Message "EXE printer driver launch failed ('$($ExeFile.FullName)'): $($_.Exception.Message)" -Category 'Printer Driver Install'
        return $false
    }

    # EXE exit codes vary by manufacturer; non-zero may still indicate success
    # (e.g. reboot required). Treat 0 as clean success, flag anything else as a warning.
    if ($Process.ExitCode -eq 0) {
        Write-Host "OK: Installer completed (exit code 0)." -ForegroundColor Green
        $status = "Success"
    }
    else {
        Write-Host "WARNING: Installer exited with code $($Process.ExitCode)." -ForegroundColor Yellow
        Write-Host "         Review manually - this may indicate a reboot requirement or vendor-specific code." -ForegroundColor Yellow
        $status = "Warning (exit $($Process.ExitCode))"
    }

    # A clean exit code is not evidence that a driver landed. Vendor
    # bootstrappers routinely ignore silent switches and exit 0 having
    # installed nothing, so confirm against the spooler before claiming success.
    $added = @(Get-SpoolerDriverName | Where-Object { $DriversBefore -notcontains $_ })
    if ($added) {
        Write-Host "Driver(s) now registered with the spooler:" -ForegroundColor Green
        foreach ($d in $added) { Write-Host "  + $d" -ForegroundColor Green }
        $script:InstalledDriverNames = @(@($script:InstalledDriverNames) + $added | Select-Object -Unique)
    }
    else {
        Write-Host "WARNING: The installer registered no new printer driver." -ForegroundColor Yellow
        Write-Host "         It may require its GUI, or may only have extracted files to disk." -ForegroundColor Yellow
        Write-Host "         Re-run with -DriverPath pointed at the extracted driver folder so" -ForegroundColor Yellow
        Write-Host "         RUNEPRESS can install from the INF directly." -ForegroundColor Yellow
        $status = "Warning (no driver registered)"
    }

    $script:InstallationLog += [PSCustomObject]@{
        File   = $ExeFile.Name
        Type   = "EXE"
        INF    = "N/A"
        Status = $status
        Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    return $true
}

# ===========================
# INSTALL MSI DRIVERS
# ===========================

function Install-MsiDriver {
    param(
        [System.IO.FileInfo]$MsiFile
    )

    Write-Host ""
    Write-Host "Processing MSI: $($MsiFile.Name)" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

    if ($WhatIf) {
        Write-Host "[~] WhatIf: would run  msiexec /i `"$($MsiFile.FullName)`" /qn /norestart" -ForegroundColor Cyan
        $script:InstallationLog += [PSCustomObject]@{
            File = $MsiFile.Name; Type = "MSI"; INF = "N/A"
            Status = "WhatIf (not installed)"; Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        return $true
    }

    Write-Host "Running silent installer..." -ForegroundColor Yellow
    $DriversBefore = Get-SpoolerDriverName

    try {
        $Process = Start-Process -FilePath "msiexec.exe" `
            -ArgumentList "/i `"$($MsiFile.FullName)`" /qn /norestart" `
            -Wait -PassThru -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to launch msiexec: $($_.Exception.Message)" -ForegroundColor Red
        $script:InstallationLog += [PSCustomObject]@{
            File   = $MsiFile.Name
            Type   = "MSI"
            INF    = "N/A"
            Status = "Failed (launch error)"
            Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        Write-TKError -ScriptName 'runepress' -Message "msiexec printer driver launch failed ('$($MsiFile.FullName)'): $($_.Exception.Message)" -Category 'Printer Driver Install'
        return $false
    }

    switch ($Process.ExitCode) {
        0 {
            Write-Host "OK: MSI installed successfully." -ForegroundColor Green
            $status = "Success"
        }
        3010 {
            Write-Host "OK: MSI installed successfully. A system reboot is required." -ForegroundColor Yellow
            $status = "Success (reboot required)"
        }
        default {
            Write-Host "ERROR: msiexec failed with exit code $($Process.ExitCode)." -ForegroundColor Red
            $script:InstallationLog += [PSCustomObject]@{
                File   = $MsiFile.Name
                Type   = "MSI"
                INF    = "N/A"
                Status = "Failed (exit $($Process.ExitCode))"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            return $false
        }
    }

    # As with EXE packages, a clean exit code does not prove a driver was
    # registered — confirm against the spooler before reporting success.
    $added = @(Get-SpoolerDriverName | Where-Object { $DriversBefore -notcontains $_ })
    if ($added) {
        Write-Host "Driver(s) now registered with the spooler:" -ForegroundColor Green
        foreach ($d in $added) { Write-Host "  + $d" -ForegroundColor Green }
        $script:InstalledDriverNames = @(@($script:InstalledDriverNames) + $added | Select-Object -Unique)
    }
    else {
        Write-Host "WARNING: The installer registered no new printer driver." -ForegroundColor Yellow
        Write-Host "         Re-run with -DriverPath pointed at the extracted driver folder so" -ForegroundColor Yellow
        Write-Host "         RUNEPRESS can install from the INF directly." -ForegroundColor Yellow
        $status = "Warning (no driver registered)"
    }

    $script:InstallationLog += [PSCustomObject]@{
        File   = $MsiFile.Name
        Type   = "MSI"
        INF    = "N/A"
        Status = $status
        Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    return $true
}

# ===========================
# SELECT INSTALLED PRINTER DRIVER
# ===========================

function Select-InstalledDriver {
    # Returns the spooler entry (Name + InfPath), not just a name: the printui
    # fallback below needs the INF path.
    $All = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Sort-Object Name)

    if ($All.Count -eq 0) {
        Write-Host "ERROR: No printer drivers found on this system." -ForegroundColor Red
        return $null
    }

    # Anything registered during this run goes to the top — it is nearly always
    # what the technician wants, and its absence makes a failed install obvious.
    $Fresh   = @($All | Where-Object { $script:InstalledDriverNames -contains $_.Name })
    $Rest    = @($All | Where-Object { $script:InstalledDriverNames -notcontains $_.Name })
    $Drivers = @($Fresh + $Rest)

    Write-Host ""
    Write-Host "Available printer drivers:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Drivers.Count; $i++) {
        if ($script:InstalledDriverNames -contains $Drivers[$i].Name) {
            Write-Host "  [$($i + 1)] $($Drivers[$i].Name)  <- installed this run" -ForegroundColor Green
        }
        else {
            Write-Host "  [$($i + 1)] $($Drivers[$i].Name)" -ForegroundColor White
        }
    }
    Write-Host ""

    do {
        $selection = Read-Host "Select driver (1-$($Drivers.Count))"
    } while (-not ($selection -match '^\d+$') -or [int]$selection -lt 1 -or [int]$selection -gt $Drivers.Count)

    $chosen = $Drivers[[int]$selection - 1]
    return [PSCustomObject]@{
        Name    = $chosen.Name
        InfPath = $chosen.InfPath
    }
}

# ===========================
# NETWORK PRINTER CONFIGURATION
# ===========================

function Add-NetworkPrinter {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Step 3: Network Printer Configuration" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($WhatIf) {
        Write-Host "[~] WhatIf: network printer addition skipped. No port or printer created." -ForegroundColor Cyan
        return
    }

    while ($true) {
        $response = Read-Host "Add a network printer? (Y/N)"
        if ($response.ToUpper() -ne "Y") {
            Write-Host "Skipping network printer configuration." -ForegroundColor Yellow
            return
        }

        # Printer display name
        do {
            $PrinterName = Read-Host "Printer display name (e.g. Office Printer 1)"
            if (-not $PrinterName) {
                Write-Host "ERROR: Name cannot be empty." -ForegroundColor Red
            }
        } while (-not $PrinterName)

        # Connection type
        Write-Host ""
        Write-Host "Connection type:" -ForegroundColor Yellow
        Write-Host "  [1] IP Address (TCP/IP port)" -ForegroundColor White
        Write-Host "  [2] UNC path   (\\server\share)" -ForegroundColor White
        Write-Host ""

        do {
            $connType = Read-Host "Enter choice (1 or 2)"
        } while ($connType -ne "1" -and $connType -ne "2")

        if ($connType -eq "1") {
            # --- IP-based printer ---
            do {
                $IPAddress = Read-Host "Printer IP address"
                if ($IPAddress -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
                    Write-Host "ERROR: Invalid IP address format." -ForegroundColor Red
                    $IPAddress = $null
                }
            } while (-not $IPAddress)

            $PortName = "IP_$IPAddress"

            # Create TCP/IP port if it does not exist
            if (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue) {
                Write-Host "Port '$PortName' already exists, reusing." -ForegroundColor Gray
            }
            else {
                Write-Host "Creating TCP/IP port: $PortName..." -ForegroundColor Yellow
                try {
                    Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -ErrorAction Stop
                    Write-Host "OK: Port created." -ForegroundColor Green
                }
                catch {
                    Write-Host "ERROR: Could not create port: $($_.Exception.Message)" -ForegroundColor Red
                    continue
                }
            }

            # Driver selection
            $Driver = Select-InstalledDriver
            if (-not $Driver) { continue }
            $DriverName = $Driver.Name

            # Add printer — try Add-Printer first; if it fails (e.g. device unreachable),
            # fall back to printui.dll which skips the reachability probe.
            Write-Host "Adding printer '$PrinterName'..." -ForegroundColor Yellow
            $printerAdded = $false
            try {
                Add-Printer -Name $PrinterName -PortName $PortName -DriverName $DriverName -ErrorAction Stop
                Write-Host "OK: Printer '$PrinterName' added successfully." -ForegroundColor Green
                $printerAdded = $true
            }
            catch {
                Write-Host "WARNING: Add-Printer failed ($($_.Exception.Message))" -ForegroundColor Yellow
                Write-Host "         Retrying via printui (offline-safe)..." -ForegroundColor Yellow
                # printui's /f switch takes the driver's INF path, not its model
                # name — passing the name silently fails. The path comes from the
                # spooler entry, so it resolves for any driver, not just ones
                # RUNEPRESS installed.
                if (-not $Driver.InfPath) {
                    Write-Host "ERROR: No INF path recorded for '$DriverName' - printui fallback unavailable." -ForegroundColor Red
                }
                else {
                    try {
                        $printArgs = "/if /b `"$PrinterName`" /f `"$($Driver.InfPath)`" /r `"$PortName`" /m `"$DriverName`""
                        $p = Start-Process -FilePath "rundll32.exe" `
                            -ArgumentList "printui.dll,PrintUIEntry $printArgs" `
                            -Wait -PassThru -ErrorAction Stop
                        if ($p.ExitCode -eq 0) {
                            Write-Host "OK: Printer '$PrinterName' added via printui." -ForegroundColor Green
                            $printerAdded = $true
                        }
                        else {
                            Write-Host "ERROR: printui exited with code $($p.ExitCode)." -ForegroundColor Red
                        }
                    }
                    catch {
                        Write-Host "ERROR: printui fallback failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            if (-not $printerAdded) {
                $script:InstallationLog += [PSCustomObject]@{
                    File   = $PrinterName
                    Type   = "Network (IP)"
                    INF    = $DriverName
                    Status = "Failed"
                    Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                continue
            }

            $script:InstallationLog += [PSCustomObject]@{
                File   = $PrinterName
                Type   = "Network (IP)"
                INF    = $DriverName
                Status = "Added ($IPAddress)"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        }
        else {
            # --- UNC-based printer ---
            do {
                $UNCPath = Read-Host "UNC path (e.g. \\server\PrinterShare)"
                if (-not $UNCPath.StartsWith("\\")) {
                    Write-Host "ERROR: Path must start with \\." -ForegroundColor Red
                    $UNCPath = $null
                }
            } while (-not $UNCPath)

            Write-Host "Connecting to: $UNCPath..." -ForegroundColor Yellow
            try {
                Add-Printer -ConnectionName $UNCPath -ErrorAction Stop
                Write-Host "OK: Connected to printer at $UNCPath." -ForegroundColor Green
            }
            catch {
                Write-Host "ERROR: Could not connect to printer: $($_.Exception.Message)" -ForegroundColor Red
                $script:InstallationLog += [PSCustomObject]@{
                    File   = $PrinterName
                    Type   = "Network (UNC)"
                    INF    = "N/A"
                    Status = "Failed"
                    Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                continue
            }

            $script:InstallationLog += [PSCustomObject]@{
                File   = $PrinterName
                Type   = "Network (UNC)"
                INF    = "N/A"
                Status = "Added ($UNCPath)"
                Time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        }

        Write-Host ""
    }
}

# ===========================
# INSTALLATION SUMMARY
# ===========================

function Show-InstallationSummary {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Installation Summary" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not $InstallationLog) {
        Write-Host "  No installations were performed." -ForegroundColor Yellow
        Write-Host ""
        return
    }

    foreach ($entry in $InstallationLog) {
        if ($entry.Status -like "Success*" -or $entry.Status -like "Added*") {
            $color = "Green"
        }
        elseif ($entry.Status -like "Warning*") {
            $color = "Yellow"
        }
        else {
            $color = "Red"
        }
        Write-Host "  [$($entry.Status)] $($entry.File) ($($entry.Type))" -ForegroundColor $color
    }

    Write-Host ""

    # Export log to CSV
    $LogPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "RUNEPRESS_InstallLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    try {
        $InstallationLog | Export-Csv -Path $LogPath -NoTypeInformation -ErrorAction Stop
        Write-Host "Log saved: $LogPath" -ForegroundColor Gray
    }
    catch {
        Write-Host "WARNING: Could not save log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host ""
}

# ===========================
# CLEANUP EXTRACTED FILES
# ===========================

function Invoke-CleanupPrompt {
    if (-not (Test-Path $ExtractRoot)) { return }

    if ($Unattended) {
        try {
            Remove-Item $ExtractRoot -Recurse -Force -ErrorAction Stop
            Write-Host "OK: Extracted files removed (unattended cleanup)." -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Could not remove extracted files: $($_.Exception.Message)" -ForegroundColor Red
        }
        return
    }

    Write-Host ""
    $response = Read-Host "Delete extracted driver files in '$ExtractRoot'? (Y/N)"
    if ($response.ToUpper() -ne "Y") { return }

    try {
        Remove-Item $ExtractRoot -Recurse -Force -ErrorAction Stop
        Write-Host "OK: Extracted files removed." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Could not remove extracted files: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ===========================
# MAIN
# ===========================

if (-not $Unattended) { Show-Banner }
if (-not $Unattended) {
    Show-DriverPrepInstructions
    Wait-ForDriverFile
}

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " Step 2: Driver Installation" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

$DriverFiles = Find-DriverFiles

if (-not $DriverFiles) {
    Write-Host ""
    Write-Host "ERROR: No driver files (.inf, .zip, .exe, .msi) found in:" -ForegroundColor Red
    Write-Host "  $SourcePath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Place a driver file in that directory, or re-run with -DriverPath" -ForegroundColor Yellow
    Write-Host "pointed at the folder holding the driver, and try again." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "Found $(@($DriverFiles).Count) driver file(s) to process:" -ForegroundColor Green
foreach ($file in $DriverFiles) {
    Write-Host "  * $($file.Name)" -ForegroundColor White
}

# Return values feed the summary via $InstallationLog, not the pipeline —
# discard them so the installers' $true/$false does not print to the console.
foreach ($file in $DriverFiles) {
    switch ($file.Extension.ToLower()) {
        ".inf" { $null = Install-InfDriver -InfFile $file }
        ".zip" { $null = Install-ZipDriver -ZipFile $file }
        ".exe" { $null = Install-ExeDriver -ExeFile $file }
        ".msi" { $null = Install-MsiDriver -MsiFile $file }
    }
}

if (-not $Unattended) { Add-NetworkPrinter }

Show-InstallationSummary

Invoke-CleanupPrompt
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
