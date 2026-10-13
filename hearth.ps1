# hearth.ps1 - H.E.A.R.T.H. — Hub for Environment, Admin Runtime & Toolkit Hardening
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
    H.E.A.R.T.H. — Hub for Environment, Admin Runtime & Toolkit Hardening
    Toolkit Setup & Configuration Wizard for PowerShell 5.1+

.DESCRIPTION
    Interactive wizard for configuring the TechnicianToolkit. Sets org name,
    log directory, default paths for EMBALM and REVENANT, and default values
    for COVENANT. Runs first-run environment checks to confirm prerequisites
    (PowerShell version, RSAT, winget, required modules). All settings are
    persisted to config.json in the toolkit directory.

.USAGE
    PS C:\> .\hearth.ps1                    # Interactive wizard
    PS C:\> .\hearth.ps1 -Unattended        # Display current config and run environment checks silently

.NOTES
    Version : 6.1

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# INITIALIZATION
# ─────────────────────────────────────────────────────────────────────────────

# ===========================
# SHARED HELPERS
# ===========================
#region TK-SHARED
# Generated from shared/TechnicianToolkit.Shared.ps1 by build/Sync-TKShared.ps1. Do not edit
# this block: edit the source and re-run the script. The Pester suite fails on a stale block.

# Forcing UTF-8 keeps the box-drawing banners intact when a console is attached.
# When one is not — a GUI host, a WinRM runspace, a scheduled task — the setter
# throws "The handle is invalid", so this must never be fatal. Every tool runs
# this as the first line of its generated block.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

function Get-TKConfig {
    <#
    .SYNOPSIS
        Returns the toolkit configuration as a PSCustomObject.
        Reads config.json from the tool's own directory; missing keys are filled
        with empty-string defaults so callers never receive null.
    #>
    $configPath = Join-Path $PSScriptRoot 'config.json'

    $defaults = [PSCustomObject]@{
        OrgName       = ''
        LogDirectory  = ''
        Archive       = [PSCustomObject]@{ DefaultDestination   = '' }
        Revenant      = [PSCustomObject]@{ DefaultDestination   = '' }
        Covenant      = [PSCustomObject]@{ DefaultTimezone = ''; DefaultLocalAdminUser = '' }
        # DirectDownloads is an array of { Name; Url; Args; Sha256 } — installers
        # CONJURE fetches by URL because they have no winget/Chocolatey package.
        Conjure       = [PSCustomObject]@{ DirectDownloads = @() }
    }

    if (-not (Test-Path $configPath)) { return $defaults }

    try {
        $raw = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json

        # Legacy migration: v2.x used a 'Phantom' section; v3.0 renamed it to 'Revenant'.
        # Copy any populated 'Phantom' value forward so upgrading users don't lose their default destination.
        if ($raw.PSObject.Properties['Phantom'] -and
            $raw.Phantom -is [PSCustomObject] -and
            -not [string]::IsNullOrWhiteSpace($raw.Phantom.DefaultDestination)) {
            if ($raw.PSObject.Properties['Revenant'] -isnot [PSCustomObject] -or
                [string]::IsNullOrWhiteSpace($raw.Revenant.DefaultDestination)) {
                if (-not $raw.PSObject.Properties['Revenant']) {
                    $raw | Add-Member -NotePropertyName 'Revenant' -NotePropertyValue ([PSCustomObject]@{ DefaultDestination = '' }) -Force
                }
                $raw.Revenant.DefaultDestination = $raw.Phantom.DefaultDestination
            }
        }

        # Ensure top-level keys exist
        foreach ($key in ($defaults | Get-Member -MemberType NoteProperty).Name) {
            if ($null -eq $raw.$key) {
                $raw | Add-Member -NotePropertyName $key -NotePropertyValue $defaults.$key -Force
            }
        }
        # Ensure nested keys exist
        foreach ($section in @('Archive','Revenant','Covenant','Conjure')) {
            if ($raw.$section -isnot [PSCustomObject]) {
                $raw | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            } else {
                foreach ($key in ($defaults.$section | Get-Member -MemberType NoteProperty).Name) {
                    if ($null -eq $raw.$section.$key) {
                        # Carry the default's own type: filling an absent array key
                        # with '' would hand the caller a string to iterate.
                        $fallback = if ($defaults.$section.$key -is [Array]) { @() } else { '' }
                        $raw.$section | Add-Member -NotePropertyName $key -NotePropertyValue $fallback -Force
                    }
                }
            }
        }
        return $raw
    }
    catch { return $defaults }
}

function Set-TKConfig {
    <#
    .SYNOPSIS
        Writes a single value to config.json.
    .EXAMPLE
        Set-TKConfig -Key 'OrgName'             -Value 'Contoso'
        Set-TKConfig -Key 'DefaultDestination'  -Value '\\srv\backups' -Section 'Archive'
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Value,
        [string]$Section = ''
    )

    $configPath = Join-Path $PSScriptRoot 'config.json'
    $cfg = if (Test-Path $configPath) {
        Get-Content $configPath -Raw | ConvertFrom-Json
    } else {
        [PSCustomObject]@{}
    }

    if ($Section) {
        if ($null -eq $cfg.$Section) {
            $cfg | Add-Member -NotePropertyName $Section -NotePropertyValue ([PSCustomObject]@{}) -Force
        }
        $cfg.$Section | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
    } else {
        $cfg | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
    }

    $cfg | ConvertTo-Json -Depth 4 | Set-Content $configPath -Encoding UTF8
}

function Start-TKTranscript {
    <#
    .SYNOPSIS
        Starts a timestamped PowerShell transcript in the configured log directory.
        Call once near the top of a script when -Transcript is active.
    #>
    param([Parameter(Mandatory)][string]$LogRoot)
    $path = Join-Path $LogRoot "TK_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    try {
        Start-Transcript -Path $path -Append -ErrorAction Stop | Out-Null
        Write-Host "  [*] Transcript: $path" -ForegroundColor Gray
    }
    catch {
        Write-Warn "Could not start transcript: $_"
    }
}

function Stop-TKTranscript {
    <#
    .SYNOPSIS
        Stops an active PowerShell transcript started by Start-TKTranscript.
    #>
    try { Stop-Transcript -ErrorAction Stop | Out-Null } catch {}
}

function Resolve-LogDirectory {
    <#
    .SYNOPSIS
        Returns the configured LogDirectory, or the supplied fallback path if
        LogDirectory is not set. Creates the directory if it does not exist.
    #>
    param([Parameter(Mandatory)][string]$FallbackPath)

    $cfg = Get-TKConfig
    if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory)) {
        if (-not (Test-Path $cfg.LogDirectory)) {
            $null = New-Item -ItemType Directory -Path $cfg.LogDirectory -Force -ErrorAction SilentlyContinue
        }
        return $cfg.LogDirectory
    }
    return $FallbackPath
}

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
Assert-AdminPrivilege

if ($PSScriptRoot) {
    $ScriptPath = $PSScriptRoot
} elseif ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
} else {
    $ScriptPath = (Get-Location).Path
}

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ─────────────────────────────────────────────────────────────────────────────
# FIELD DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

$Fields = @(
    [PSCustomObject]@{
        DisplayName = 'Organization Name'
        Key         = 'OrgName'
        Section     = ''
        Description = 'Shown in the header of all HTML reports'
        Hint        = 'Contoso IT'
        IsPath      = $false
    },
    [PSCustomObject]@{
        DisplayName = 'Log & Report Directory'
        Key         = 'LogDirectory'
        Section     = ''
        Description = 'Where HTML reports and transcripts are saved'
        Hint        = 'C:\TKLogs'
        IsPath      = $true
    },
    [PSCustomObject]@{
        DisplayName = 'EMBALM Default Destination'
        Key         = 'DefaultDestination'
        Section     = 'Archive'
        Description = 'Default backup path for EMBALM (pre-reimaging profile backup)'
        Hint        = '\\fileserver\Backups\Archive'
        IsPath      = $true
    },
    [PSCustomObject]@{
        DisplayName = 'REVENANT Default Destination'
        Key         = 'DefaultDestination'
        Section     = 'Revenant'
        Description = 'Default migration target path for REVENANT'
        Hint        = '\\fileserver\Migrations'
        IsPath      = $true
    },
    [PSCustomObject]@{
        DisplayName = 'COVENANT Default Timezone'
        Key         = 'DefaultTimezone'
        Section     = 'Covenant'
        Description = 'Default Windows timezone ID for COVENANT onboarding'
        Hint        = 'Eastern Standard Time'
        IsPath      = $false
    },
    [PSCustomObject]@{
        DisplayName = 'COVENANT Default Local Admin'
        Key         = 'DefaultLocalAdminUser'
        Section     = 'Covenant'
        Description = 'Default local administrator account name for COVENANT'
        Hint        = 'LocalAdmin'
        IsPath      = $false
    }
)

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-HearthBanner {
    Clear-Host
    Write-Host ""
    Write-Host "  ██╗  ██╗███████╗ █████╗ ██████╗ ████████╗██╗  ██╗" -ForegroundColor $C.Header
    Write-Host "  ██║  ██║██╔====╝██╔==██╗██╔==██╗╚==██╔==╝██║  ██║" -ForegroundColor $C.Header
    Write-Host "  ███████║█████╗  ███████║██████╔╝   ██║   ███████║" -ForegroundColor $C.Header
    Write-Host "  ██╔==██║██╔==╝  ██╔==██║██╔==██╗   ██║   ██╔==██║" -ForegroundColor $C.Header
    Write-Host "  ██║  ██║███████╗██║  ██║██║  ██║   ██║   ██║  ██║" -ForegroundColor $C.Header
    Write-Host "  ╚=╝  ╚=╝╚======╝╚=╝  ╚=╝╚=╝  ╚=╝   ╚=╝   ╚=╝  ╚=╝" -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  H.E.A.R.T.H.  -  Hub for Environment, Admin Runtime & Toolkit Hardening" -ForegroundColor $C.Header
    Write-Host "  Toolkit Setup & Configuration Wizard" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  TechnicianToolkit  |  HEARTH v6.1  |  Run as Administrator" -ForegroundColor $C.Info
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Get current value for a field
# ─────────────────────────────────────────────────────────────────────────────

function Get-FieldValue {
    param([PSCustomObject]$Field)
    $cfg = Get-TKConfig
    if ($Field.Section) {
        return $cfg.($Field.Section).($Field.Key)
    } else {
        return $cfg.($Field.Key)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Validate and optionally create a path
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-PathValidation {
    param([string]$Path)
    if (Test-Path $Path) { return $true }

    Write-Host ""
    Write-Host "  [!] Path not found: $Path" -ForegroundColor $C.Warning
    Write-Host -NoNewline "      Path not found  -  create it? (Y/N): " -ForegroundColor $C.Warning
    $answer = (Read-Host).Trim().ToUpper()

    if ($answer -eq 'Y') {
        try {
            $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop
            Write-Host "  [+] Directory created: $Path" -ForegroundColor $C.Success
            return $true
        }
        catch {
            Write-Host "  [!] Could not create directory: $_" -ForegroundColor $C.Warning
            Write-Host "      Saving value anyway  -  verify the path manually." -ForegroundColor $C.Info
            return $false
        }
    } else {
        Write-Host "      Saving value anyway  -  verify the path manually." -ForegroundColor $C.Info
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Save a field value
# ─────────────────────────────────────────────────────────────────────────────

function Save-FieldValue {
    param(
        [PSCustomObject]$Field,
        [string]$Value
    )
    if ($Field.Section) {
        Set-TKConfig -Key $Field.Key -Value $Value -Section $Field.Section
    } else {
        Set-TKConfig -Key $Field.Key -Value $Value
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Show-CurrentConfig
# ─────────────────────────────────────────────────────────────────────────────

function Show-CurrentConfig {
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  CURRENT CONFIGURATION" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""

    foreach ($field in $Fields) {
        $value = Get-FieldValue -Field $field

        Write-Host -NoNewline ("  {0,-34} " -f ($field.DisplayName + ':')) -ForegroundColor $C.Info

        if ([string]::IsNullOrWhiteSpace($value)) {
            Write-Host "(not configured)" -ForegroundColor $C.Warning
        } elseif ($field.IsPath -and -not (Test-Path $value)) {
            Write-Host "$value  (path not found)" -ForegroundColor $C.Warning
        } else {
            Write-Host $value -ForegroundColor $C.Success
        }
    }

    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-SetupWizard
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-SetupWizard {
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  SETUP WIZARD" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  Step through each setting. Press Enter to keep the current value." -ForegroundColor $C.Info
    Write-Host ""

    $stepNum = 0
    foreach ($field in $Fields) {
        $stepNum++
        $currentValue = Get-FieldValue -Field $field

        Write-Host ("  " + ("-" * 40)) -ForegroundColor $C.Header
        Write-Host "  Step $stepNum of $($Fields.Count)  -  $($field.DisplayName)" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 40)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "      Description : $($field.Description)" -ForegroundColor $C.Info
        Write-Host "      Example     : $($field.Hint)" -ForegroundColor $C.Info

        if ([string]::IsNullOrWhiteSpace($currentValue)) {
            Write-Host "      Current     : (not configured)" -ForegroundColor $C.Warning
        } else {
            Write-Host "      Current     : $currentValue" -ForegroundColor $C.Success
        }

        Write-Host ""
        Write-Host -NoNewline "  Enter new value (or press Enter to keep current): " -ForegroundColor $C.Header
        $userInput = (Read-Host).Trim()

        if ([string]::IsNullOrWhiteSpace($userInput)) {
            $displayCurrent = if ([string]::IsNullOrWhiteSpace($currentValue)) { '(not configured)' } else { $currentValue }
            Write-Host "  [*] Kept: $displayCurrent" -ForegroundColor $C.Info
        } else {
            if ($field.IsPath) {
                $null = Invoke-PathValidation -Path $userInput
            }
            Save-FieldValue -Field $field -Value $userInput
            Write-Host "  [+] Saved." -ForegroundColor $C.Success
        }

        Write-Host ""
    }

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  Setup wizard complete. Run option [4] to verify your environment." -ForegroundColor $C.Success
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-EditField
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-EditField {
    do {
        Write-Host ""
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  EDIT A SINGLE SETTING" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""

        $index = 0
        foreach ($field in $Fields) {
            $index++
            $value = Get-FieldValue -Field $field

            Write-Host -NoNewline "  [$index] $($field.DisplayName)" -ForegroundColor $C.Info
            Write-Host -NoNewline "   -   " -ForegroundColor $C.Info

            if ([string]::IsNullOrWhiteSpace($value)) {
                Write-Host "(not configured)" -ForegroundColor $C.Warning
            } elseif ($field.IsPath -and -not (Test-Path $value)) {
                Write-Host "$value  (path not found)" -ForegroundColor $C.Warning
            } else {
                Write-Host $value -ForegroundColor $C.Success
            }
        }

        Write-Host ""
        Write-Host -NoNewline "  Select a field to edit (1-$($Fields.Count)): " -ForegroundColor $C.Header
        $selection = (Read-Host).Trim()

        $selIndex = 0
        if (-not ([int]::TryParse($selection, [ref]$selIndex)) -or $selIndex -lt 1 -or $selIndex -gt $Fields.Count) {
            Write-Host "  [!] Invalid selection." -ForegroundColor $C.Warning
            continue
        }

        $field = $Fields[$selIndex - 1]
        $currentValue = Get-FieldValue -Field $field

        Write-Host ""
        Write-Host ("  " + ("-" * 40)) -ForegroundColor $C.Header
        Write-Host "  $($field.DisplayName)" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 40)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "      Description : $($field.Description)" -ForegroundColor $C.Info
        Write-Host "      Example     : $($field.Hint)" -ForegroundColor $C.Info

        if ([string]::IsNullOrWhiteSpace($currentValue)) {
            Write-Host "      Current     : (not configured)" -ForegroundColor $C.Warning
        } else {
            Write-Host "      Current     : $currentValue" -ForegroundColor $C.Success
        }

        Write-Host ""
        Write-Host -NoNewline "  Enter new value (or press Enter to keep current): " -ForegroundColor $C.Header
        $newValue = (Read-Host).Trim()

        if ([string]::IsNullOrWhiteSpace($newValue)) {
            $displayCurrent = if ([string]::IsNullOrWhiteSpace($currentValue)) { '(not configured)' } else { $currentValue }
            Write-Host "  [*] Kept: $displayCurrent" -ForegroundColor $C.Info
        } else {
            if ($field.IsPath) {
                $null = Invoke-PathValidation -Path $newValue
            }
            Save-FieldValue -Field $field -Value $newValue
            Write-Host "  [+] Saved." -ForegroundColor $C.Success
        }

        Write-Host ""
        Write-Host -NoNewline "  Edit another field? (Y/N): " -ForegroundColor $C.Header
        $again = (Read-Host).Trim().ToUpper()

    } while ($again -eq 'Y')
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-EnvironmentCheck
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-EnvironmentCheck {
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  ENVIRONMENT CHECKS" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""

    # 1. PowerShell Version
    $psVersion = $PSVersionTable.PSVersion.Major
    if ($psVersion -ge 5) {
        Write-Host "  [+] PowerShell Version          -  $($PSVersionTable.PSVersion.ToString())" -ForegroundColor $C.Success
    } else {
        Write-Host "  [-] PowerShell Version          -  $($PSVersionTable.PSVersion.ToString()) (5.1+ required)" -ForegroundColor $C.Error
    }

    # 2. Running as Administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        Write-Host "  [+] Running as Administrator    -  Yes" -ForegroundColor $C.Success
    } else {
        Write-Host "  [-] Running as Administrator    -  No (required for most tools)" -ForegroundColor $C.Error
    }

    # 3. config.json exists (informational)
    if (Test-Path "$ScriptPath\config.json") {
        Write-Host "  [+] config.json                 -  Found" -ForegroundColor $C.Success
    } else {
        Write-Host "  [!] config.json                 -  Not found (will be created on first save)" -ForegroundColor $C.Warning
    }

    # 4. Log directory accessible
    $cfg = Get-TKConfig
    if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory)) {
        if (Test-Path $cfg.LogDirectory) {
            $testFile = Join-Path $cfg.LogDirectory "HEARTH_writetest_$([System.IO.Path]::GetRandomFileName()).tmp"
            try {
                [System.IO.File]::WriteAllText($testFile, 'test')
                Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
                Write-Host "  [+] Log directory accessible    -  $($cfg.LogDirectory)" -ForegroundColor $C.Success
            }
            catch {
                Write-Host "  [!] Log directory not writable  -  $($cfg.LogDirectory)" -ForegroundColor $C.Warning
            }
        } else {
            Write-Host "  [!] Log directory not found     -  $($cfg.LogDirectory)" -ForegroundColor $C.Warning
        }
    } else {
        Write-Host "  [!] Log directory               -  Not configured" -ForegroundColor $C.Warning
    }

    # 5. winget available (for CONJURE)
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host "  [+] winget                      -  Available ($($winget.Source))" -ForegroundColor $C.Success
    } else {
        Write-Host "  [!] winget                      -  Not found (required for CONJURE)" -ForegroundColor $C.Warning
    }

    # 6. Chocolatey available (for CONJURE)
    $choco = Get-Command choco -ErrorAction SilentlyContinue
    if ($choco) {
        Write-Host "  [+] Chocolatey (choco)          -  Available" -ForegroundColor $C.Success
    } else {
        Write-Host "  [!] Chocolatey (choco)          -  Not found (optional for CONJURE)" -ForegroundColor $C.Warning
    }

    # 7. RSAT ActiveDirectory module (for SPHINX)
    $adModule = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
    if ($adModule) {
        Write-Host "  [+] RSAT: ActiveDirectory       -  Available ($($adModule[0].Version))" -ForegroundColor $C.Success
    } else {
        Write-Host "  [!] RSAT: ActiveDirectory       -  Not found (required for SPHINX)" -ForegroundColor $C.Warning
    }

    # 8. Microsoft.Graph module (for ALMANAC)
    $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph -ErrorAction SilentlyContinue
    if ($graphModule) {
        Write-Host "  [+] Microsoft.Graph module      -  Available ($($graphModule[0].Version))" -ForegroundColor $C.Success
    } else {
        Write-Host "  [!] Microsoft.Graph module      -  Not found (required for ALMANAC)" -ForegroundColor $C.Warning
    }

    # 10. Az module (for ZENITH)
    $azModule = Get-Module -ListAvailable -Name Az -ErrorAction SilentlyContinue
    if ($azModule) {
        Write-Host "  [+] Az module                   -  Available ($($azModule[0].Version))" -ForegroundColor $C.Success
    } else {
        Write-Host "  [!] Az module                   -  Not found (required for ZENITH)" -ForegroundColor $C.Warning
    }

    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  Environment check complete." -ForegroundColor $C.Info
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-ResetConfig
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ResetConfig {
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  RESET CONFIGURATION" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  [!] This will delete config.json and clear all settings." -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host -NoNewline "  Are you sure? (YES to confirm): " -ForegroundColor $C.Warning
    $confirm = (Read-Host).Trim()

    if ($confirm -eq 'YES') {
        $configPath = Join-Path $ScriptPath 'config.json'
        try {
            if (Test-Path $configPath) {
                Remove-Item -Path $configPath -Force -ErrorAction Stop
            }
            Write-Host "  [+] Configuration reset." -ForegroundColor $C.Success
        }
        catch {
            Write-Host "  [-] Failed to delete config.json: $_" -ForegroundColor $C.Error
        }
    } else {
        Write-Host "  [*] Reset cancelled." -ForegroundColor $C.Info
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU HELPER
# ─────────────────────────────────────────────────────────────────────────────

function Show-HearthMenu {
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  MAIN MENU" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  [1] Run setup wizard           (step through all settings)" -ForegroundColor $C.Info
    Write-Host "  [2] View current configuration" -ForegroundColor $C.Info
    Write-Host "  [3] Edit a single setting" -ForegroundColor $C.Info
    Write-Host "  [4] Run environment checks" -ForegroundColor $C.Info
    Write-Host "  [5] Reset configuration        (clear all settings)" -ForegroundColor $C.Warning
    Write-Host "  [Q] Quit" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-CurrentConfig
    Invoke-EnvironmentCheck
} else {
    do {
        Show-HearthBanner
        Show-HearthMenu

        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Info
        $selection = (Read-Host).Trim().ToUpper()

        switch ($selection) {
            '1' {
                Invoke-SetupWizard
                Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
                $null = Read-Host
            }
            '2' {
                Show-CurrentConfig
                Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
                $null = Read-Host
            }
            '3' {
                Invoke-EditField
                Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
                $null = Read-Host
            }
            '4' {
                Invoke-EnvironmentCheck
                Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
                $null = Read-Host
            }
            '5' {
                Invoke-ResetConfig
                Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
                $null = Read-Host
            }
            'Q' {
                Clear-Host
                Write-Host ""
                Write-Host "  Closing H.E.A.R.T.H. Toolkit configured and ready." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!] Invalid selection. Enter 1-5 or Q to quit." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

    } while ($selection -ne 'Q')
}

# ─────────────────────────────────────────────────────────────────────────────
# CLEANUP
# ─────────────────────────────────────────────────────────────────────────────

if ($Transcript) { Stop-TKTranscript }
