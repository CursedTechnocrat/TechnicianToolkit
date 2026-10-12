# emissary.ps1 - E.M.I.S.S.A.R.Y. — Executes Modules In Sessions Sent Across Remote sYstems
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
    E.M.I.S.S.A.R.Y. — Executes Modules In Sessions Sent Across Remote sYstems
    Remote Machine Execution Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to a remote Windows machine via WinRM and runs Technician Toolkit
    scripts without needing to be physically at the target. Supports credential
    prompting, connectivity checks, remote execution of non-interactive tools,
    automatic retrieval of output files, and interactive remote sessions.

.USAGE
    PS C:\> .\emissary.ps1      # Must be run as Administrator
    Target machine must have WinRM enabled. Run on target:
        Enable-PSRemoting -Force

.NOTES
    Version : 6.0

    Remote-Compatible Tools
    ─────────────────────────────────────────────────────────────────
    A.U.S.P.E.X.   — Diagnostics report (HTML retrieved automatically)
    W.A.R.D.       — Account audit (HTML retrieved automatically)
    W.H.E.T.S.T.O.N.E. — Windows Updates (non-interactive)
    B.A.S.I.L.I.S.K.     — Baseline enforcement (auto-apply all categories)

#>

param(
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
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

# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT PATH RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

if ($PSScriptRoot) {
    $ScriptPath = $PSScriptRoot
} elseif ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
} else {
    $ScriptPath = (Get-Location).Path
}

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA
# ─────────────────────────────────────────────────────────────────────────────

$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-EmissaryBanner {
    Clear-Host
    Write-Host @"

  ███████╗███╗   ███╗██╗███████╗███████╗ █████╗ ██████╗ ██╗   ██╗
  ██╔════╝████╗ ████║██║██╔════╝██╔════╝██╔══██╗██╔══██╗╚██╗ ██╔╝
  █████╗  ██╔████╔██║██║███████╗███████╗███████║██████╔╝ ╚████╔╝
  ██╔══╝  ██║╚██╔╝██║██║╚════██║╚════██║██╔══██║██╔══██╗  ╚██╔╝
  ███████╗██║ ╚═╝ ██║██║███████║███████║██║  ██║██║  ██║   ██║
  ╚══════╝╚═╝     ╚═╝╚═╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    E.M.I.S.S.A.R.Y. — Executes Modules In Sessions Sent Across Remote sYstems" -ForegroundColor Cyan
    Write-Host "    Remote Machine Execution Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# WINRM HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-WinRMConnectivity {
    param([string]$ComputerName, [PSCredential]$Credential)

    Write-Host "  [*] Testing WinRM connectivity to $ComputerName..." -ForegroundColor $ColorSchema.Progress

    try {
        $wsmanParams = @{ ComputerName = $ComputerName; ErrorAction = "Stop" }
        if ($Credential) { $wsmanParams.Credential = $Credential }
        Test-WSMan @wsmanParams | Out-Null
        Write-Host "  [+] WinRM is reachable on $ComputerName." -ForegroundColor $ColorSchema.Success
        return $true
    }
    catch {
        Write-Host "  [-] WinRM connection failed: $_" -ForegroundColor $ColorSchema.Error
        return $false
    }
}

function Show-WinRMInstructions {
    param([string]$ComputerName)

    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host "  HOW TO ENABLE WINRM ON $($ComputerName.ToUpper())" -ForegroundColor $ColorSchema.Warning
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host "  Option 1 — Run on the target machine (as Administrator):" -ForegroundColor $ColorSchema.Info
    Write-Host "    Enable-PSRemoting -Force" -ForegroundColor $ColorSchema.Accent
    Write-Host ""
    Write-Host "  Option 2 — Push via Group Policy:" -ForegroundColor $ColorSchema.Info
    Write-Host "    Computer Config > Policies > Windows Settings > Scripts > Startup" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host "  Option 3 — Enable via remote registry (if admin share accessible):" -ForegroundColor $ColorSchema.Info
    Write-Host "    winrm /r:$ComputerName quickconfig" -ForegroundColor $ColorSchema.Accent
    Write-Host ""
    Write-Host "  Also verify:" -ForegroundColor $ColorSchema.Info
    Write-Host "    - Windows Firewall allows TCP 5985 (HTTP) or 5986 (HTTPS)" -ForegroundColor $ColorSchema.Info
    Write-Host "    - Target is reachable on the network (ping or Test-NetConnection)" -ForegroundColor $ColorSchema.Info
    Write-Host ""
}

function New-RemoteSession {
    param([string]$ComputerName, [PSCredential]$Credential)

    try {
        $sessionParams = @{ ComputerName = $ComputerName; ErrorAction = "Stop" }
        if ($Credential) { $sessionParams.Credential = $Credential }
        $session = New-PSSession @sessionParams
        return $session
    }
    catch {
        Write-Host "  [-] Failed to create remote session: $_" -ForegroundColor $ColorSchema.Error
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REMOTE TOOL EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RemoteTool {
    param(
        [System.Management.Automation.Runspaces.PSSession]$Session,
        [string]$ScriptFile,
        [string]$ComputerName,
        [string]$ToolName,
        [hashtable]$ScriptArgs = @{ Unattended = $true }
    )

    $localScript   = Join-Path $ScriptPath $ScriptFile
    $remoteTempDir = "C:\Temp\EmissaryToolkit"
    $remoteScript  = "$remoteTempDir\$ScriptFile"

    if (-not (Test-Path $localScript)) {
        Write-Host "  [-] Script not found locally: $localScript" -ForegroundColor $ColorSchema.Error
        Write-Host "  [!!] Ensure $ScriptFile is in the same folder as emissary.ps1." -ForegroundColor $ColorSchema.Warning
        return
    }

    try {
        # Create remote staging directory
        Write-Host "  [*] Preparing remote environment..." -ForegroundColor $ColorSchema.Progress
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            $null = New-Item -Path $dir -ItemType Directory -Force
        } -ArgumentList $remoteTempDir -ErrorAction Stop

        # Copy the tool to the remote machine. It carries its own shared helpers,
        # so nothing else needs to travel with it.
        Write-Host "  [*] Copying $ScriptFile to $ComputerName..." -ForegroundColor $ColorSchema.Progress
        Copy-Item -Path $localScript -Destination $remoteTempDir -ToSession $Session -ErrorAction Stop

        # Execute the script remotely
        Write-Host "  [*] Executing $ToolName on $ComputerName..." -ForegroundColor $ColorSchema.Progress
        Write-Host ""

        # A WinRM runspace has no console, so any Read-Host prompt would stall the
        # run. Every tool is invoked with -Unattended (plus any tool-specific args).
        Invoke-Command -Session $Session -ScriptBlock {
            param($scriptPath, $scriptArgs)
            & $scriptPath @scriptArgs
        } -ArgumentList $remoteScript, $ScriptArgs -ErrorAction Stop

        Write-Host ""

        # Retrieve any output files (HTML, CSV) back to local script directory
        $outputFiles = Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            Get-ChildItem -Path $dir -File |
                Where-Object { $_.Extension -in @('.html', '.csv') } |
                Select-Object -ExpandProperty Name
        } -ArgumentList $remoteTempDir

        if ($outputFiles) {
            $retrieveDir = Join-Path $ScriptPath "EMISSARY_$ComputerName"
            $null = New-Item -Path $retrieveDir -ItemType Directory -Force

            foreach ($file in $outputFiles) {
                $remoteFile = "$remoteTempDir\$file"
                $localDest  = Join-Path $retrieveDir $file
                Copy-Item -Path $remoteFile -Destination $localDest -FromSession $Session -ErrorAction SilentlyContinue
                Write-Host "  [+] Retrieved: $localDest" -ForegroundColor $ColorSchema.Success
            }
        }
    }
    catch {
        Write-Host "  [-] Remote execution failed: $_" -ForegroundColor $ColorSchema.Error
    }
    finally {
        # Clean up remote staging directory
        Write-Host "  [*] Cleaning up remote staging folder..." -ForegroundColor $ColorSchema.Progress
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
        } -ArgumentList $remoteTempDir
    }
}

function Invoke-RemoteSigil {
    param(
        [System.Management.Automation.Runspaces.PSSession]$Session,
        [string]$ComputerName
    )

    $localScript   = Join-Path $ScriptPath "basilisk.ps1"
    $remoteTempDir = "C:\Temp\EmissaryToolkit"
    $remoteScript  = "$remoteTempDir\basilisk.ps1"

    if (-not (Test-Path $localScript)) {
        Write-Host "  [-] basilisk.ps1 not found locally." -ForegroundColor $ColorSchema.Error
        return
    }

    try {
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir) $null = New-Item -Path $dir -ItemType Directory -Force
        } -ArgumentList $remoteTempDir -ErrorAction Stop

        Copy-Item -Path $localScript -Destination $remoteTempDir -ToSession $Session -ErrorAction Stop

        Write-Host "  [*] Executing B.A.S.I.L.I.S.K. baseline on $ComputerName (applying all categories)..." -ForegroundColor $ColorSchema.Progress
        Write-Host ""

        # -Categories A selects every baseline category; -Unattended keeps BASILISK off
        # Read-Host, which a WinRM runspace cannot answer.
        Invoke-Command -Session $Session -ScriptBlock {
            param($sigilScript)
            & $sigilScript -Unattended -Categories A
        } -ArgumentList $remoteScript

        Write-Host ""

        # Retrieve log
        $logFiles = Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            Get-ChildItem $dir -Filter "BASILISK_*.csv" | Select-Object -ExpandProperty Name
        } -ArgumentList $remoteTempDir

        if ($logFiles) {
            $retrieveDir = Join-Path $ScriptPath "EMISSARY_$ComputerName"
            $null = New-Item -Path $retrieveDir -ItemType Directory -Force
            foreach ($file in $logFiles) {
                Copy-Item -Path "$remoteTempDir\$file" -Destination (Join-Path $retrieveDir $file) -FromSession $Session -ErrorAction SilentlyContinue
                Write-Host "  [+] Retrieved: $(Join-Path $retrieveDir $file)" -ForegroundColor $ColorSchema.Success
            }
        }
    }
    catch {
        Write-Host "  [-] Remote B.A.S.I.L.I.S.K. failed: $_" -ForegroundColor $ColorSchema.Error
    }
    finally {
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir) Remove-Item $dir -Recurse -Force -EA SilentlyContinue
        } -ArgumentList $remoteTempDir
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-EmissaryBanner

Write-Host "  [!!] The target machine must have WinRM enabled." -ForegroundColor $ColorSchema.Warning
Write-Host "       On the target, run: Enable-PSRemoting -Force" -ForegroundColor $ColorSchema.Warning
Write-Host ""

# ── TARGET ────────────────────────────────────────────────────────────────────

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  TARGET MACHINE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host -NoNewline "  Enter hostname or IP address: " -ForegroundColor $ColorSchema.Header
$targetMachine = (Read-Host).Trim()

if ([string]::IsNullOrWhiteSpace($targetMachine)) {
    Write-Host ""
    Write-Host "  [-] No target entered." -ForegroundColor $ColorSchema.Error
    exit 1
}

# ── CREDENTIALS ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  CREDENTIALS" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  [1] Use current session credentials  (domain / Kerberos)" -ForegroundColor $ColorSchema.Info
Write-Host "  [2] Enter credentials manually" -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
$credChoice = (Read-Host).Trim()

$remoteCred = $null

if ($credChoice -eq "2") {
    Write-Host ""
    Write-Host "  Enter credentials for $targetMachine" -ForegroundColor $ColorSchema.Info
    try {
        $remoteCred = Get-Credential -ErrorAction Stop
    }
    catch {
        Write-Host "  [-] Credential entry cancelled." -ForegroundColor $ColorSchema.Error
        exit 1
    }
}

# ── CONNECTIVITY TEST ─────────────────────────────────────────────────────────

Write-Host ""
$connected = Test-WinRMConnectivity -ComputerName $targetMachine -Credential $remoteCred

if (-not $connected) {
    Show-WinRMInstructions -ComputerName $targetMachine
    Write-Host -NoNewline "  Retry connection? (Y/N): " -ForegroundColor $ColorSchema.Warning
    $retry = (Read-Host).Trim().ToUpper()
    if ($retry -eq "Y") {
        $connected = Test-WinRMConnectivity -ComputerName $targetMachine -Credential $remoteCred
    }
    if (-not $connected) {
        Write-Host ""
        Write-Host "  [-] Cannot connect to $targetMachine. Exiting." -ForegroundColor $ColorSchema.Error
        exit 1
    }
}

# ── MAIN MENU LOOP ────────────────────────────────────────────────────────────

$choice = ""

do {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  REMOTE OPERATIONS  —  Target: $targetMachine" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  [1] Run A.U.S.P.E.X.         — diagnostics & HTML report" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Run W.A.R.D.             — account audit & HTML report" -ForegroundColor $ColorSchema.Info
    Write-Host "  [3] Run W.H.E.T.S.T.O.N.E. — install Windows Updates" -ForegroundColor $ColorSchema.Info
    Write-Host "  [4] Run B.A.S.I.L.I.S.K.           — apply security baseline (all categories)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [5] Open interactive PS session" -ForegroundColor $ColorSchema.Info
    Write-Host "  [C] Check WinRM connectivity" -ForegroundColor $ColorSchema.Info
    Write-Host "  [Q] Quit" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $choice = (Read-Host).Trim().ToUpper()

    switch ($choice) {
        "1" {
            Write-Host ""
            $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
            if ($session) {
                Invoke-RemoteTool -Session $session -ScriptFile "auspex.ps1" -ComputerName $targetMachine -ToolName "A.U.S.P.E.X."
                Remove-PSSession $session
            }
        }
        "2" {
            Write-Host ""
            $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
            if ($session) {
                Invoke-RemoteTool -Session $session -ScriptFile "ward.ps1" -ComputerName $targetMachine -ToolName "W.A.R.D."
                Remove-PSSession $session
            }
        }
        "3" {
            Write-Host ""
            Write-Host "  [!!] WHETSTONE will install updates on the target machine." -ForegroundColor $ColorSchema.Warning
            Write-Host "       It runs without -AutoReboot, so a required reboot is reported, not performed." -ForegroundColor $ColorSchema.Warning
            Write-Host -NoNewline "  Continue? (Y/N): " -ForegroundColor $ColorSchema.Warning
            $confirm = (Read-Host).Trim().ToUpper()
            if ($confirm -eq "Y") {
                $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
                if ($session) {
                    Invoke-RemoteTool -Session $session -ScriptFile "whetstone.ps1" -ComputerName $targetMachine -ToolName "W.H.E.T.S.T.O.N.E."
                    Remove-PSSession $session -ErrorAction SilentlyContinue
                }
            } else {
                Write-Host "  [*] Operation cancelled." -ForegroundColor $ColorSchema.Info
            }
        }
        "4" {
            Write-Host ""
            $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
            if ($session) {
                Invoke-RemoteSigil -Session $session -ComputerName $targetMachine
                Remove-PSSession $session
            }
        }
        "5" {
            Write-Host ""
            Write-Host "  [*] Opening interactive session with $targetMachine..." -ForegroundColor $ColorSchema.Progress
            Write-Host "  [*] Type 'exit' to return to E.M.I.S.S.A.R.Y." -ForegroundColor $ColorSchema.Info
            Write-Host ""
            try {
                $enterParams = @{ ComputerName = $targetMachine }
                if ($remoteCred) { $enterParams.Credential = $remoteCred }
                Enter-PSSession @enterParams
            }
            catch {
                Write-Host "  [-] Could not open interactive session: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        "C" {
            Write-Host ""
            Test-WinRMConnectivity -ComputerName $targetMachine -Credential $remoteCred | Out-Null
        }
        "Q" {
            Write-Host ""
            Write-Host "  Closing E.M.I.S.S.A.R.Y." -ForegroundColor $ColorSchema.Header
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  [!!] Invalid selection. Enter 1-5, C, or Q." -ForegroundColor $ColorSchema.Warning
            Start-Sleep -Seconds 1
        }
    }

    if ($choice -notin @("Q", "C")) {
        Write-Host ""
        Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $ColorSchema.Info
        Read-Host | Out-Null
    }

} while ($choice -ne "Q")
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
