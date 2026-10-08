# suture.ps1 - S.U.T.U.R.E. — Stitches Up The Update & Repair Engine
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
    S.U.T.U.R.E. — Stitches Up The Update & Repair Engine
    Windows Servicing (Component Store & System File) Diagnosis and Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Diagnoses and repairs the Windows servicing stack -- the component store
    (WinSxS) that updates, features and system-file repair all depend on.
    Reach for it when updates fail with 0x800f081f / 0x80073712 / 0x800f0831,
    when Windows features will not install, or when system files are damaged.

    Audit (read-only):
      - Pending restarts, split into those that block servicing (CBS
        RebootPending, PackagesPending, pending.xml, Windows Update
        RebootRequired) and the rest (file renames, a pending computer rename)
      - Component store health (DISM /CheckHealth, or /ScanHealth with -Deep)
        and size / cleanup state (DISM /AnalyzeComponentStore)
      - The last System File Checker result, read from CBS.log: files repaired
        and files it could not repair
      - Windows Update install failures from the last 30 days, with each error
        code decoded into component-store, update-client, restart, space or
        access causes
      - The Windows Modules Installer service, free space on the system drive,
        and where DISM will look for repair files (WSUS machines often cannot
        get them)

    Repair runs the standard sequence: DISM /RestoreHealth (optionally from a
    -Source image), then sfc /scannow, then re-checks the store. -Cleanup adds
    DISM /StartComponentCleanup. It never runs /ResetBase (irreversible: it
    removes the ability to uninstall updates), never edits Windows Update
    policy, and never resets the update cache -- C.O.N.D.U.I.T. owns the
    update client's plumbing. -WhatIf previews every repair.

.USAGE
    PS C:\> .\suture.ps1                                    # Interactive menu
    PS C:\> .\suture.ps1 -Unattended                        # Read-only audit + HTML report
    PS C:\> .\suture.ps1 -Unattended -Deep                  # Audit with a full ScanHealth
    PS C:\> .\suture.ps1 -Unattended -Action Repair         # RestoreHealth + SFC + re-check
    PS C:\> .\suture.ps1 -Action Repair -Source 'WIM:D:\sources\install.wim:6'   # Repair from mounted media
    PS C:\> .\suture.ps1 -Action Repair -WhatIf             # Preview the repairs only

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit', 'Repair')]
    [string]$Action = 'Audit',
    [string]$Source = '',
    [switch]$Deep,
    [switch]$Cleanup,
    [switch]$Transcript,
    [switch]$WhatIf
)

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
$C = $ColorSchema

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

$CbsLogPath          = Join-Path $env:SystemRoot 'Logs\CBS\CBS.log'
$DismLogPath         = Join-Path $env:SystemRoot 'Logs\DISM\dism.log'
$CbsTailBytes        = 20MB       # CBS.log grows large; the last SFC run is near the end
$UpdateLookbackDays  = 30
$FreeSpaceWarningGB  = 10         # servicing stages payloads on the system drive
$FreeSpaceErrorGB    = 3

$CbsKey     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing'
$ServicingPolicyKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing'
$WuPolicyKey        = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Error codes that servicing and Windows Update report, split by what fixes
# them. Store = component-store damage (DISM / SFC); Client = the update
# client's plumbing (C.O.N.D.U.I.T.); Reboot, Space and Access are what they say.
$ServicingErrorCodes = @{
    '0x800f081f' = @{ Name = 'CBS_E_SOURCE_MISSING';                 Kind = 'Store';  Meaning = 'Repair or feature files could not be found. Give DISM a -Source, or let it reach Windows Update.' }
    '0x80073712' = @{ Name = 'ERROR_SXS_COMPONENT_STORE_CORRUPT';    Kind = 'Store';  Meaning = 'The component store is corrupt.' }
    '0x80073701' = @{ Name = 'ERROR_SXS_ASSEMBLY_MISSING';           Kind = 'Store';  Meaning = 'A component the update needs is missing from the store.' }
    '0x800f0831' = @{ Name = 'CBS_E_STORE_CORRUPTION';               Kind = 'Store';  Meaning = 'The store is missing a package the update builds on.' }
    '0x800f0982' = @{ Name = 'PSFX_E_MATCHING_COMPONENT_NOT_FOUND';  Kind = 'Store';  Meaning = 'A component the update patches is missing -- often a language pack or optional feature.' }
    '0x800f0900' = @{ Name = 'CBS_E_XML_PARSER_FAILURE';             Kind = 'Store';  Meaning = 'A servicing manifest could not be parsed.' }
    '0x80070570' = @{ Name = 'ERROR_FILE_CORRUPT';                   Kind = 'Store';  Meaning = 'A file or directory is corrupt and unreadable.' }
    '0x8007000d' = @{ Name = 'ERROR_INVALID_DATA';                   Kind = 'Store';  Meaning = 'Servicing data is invalid.' }
    '0x800f082f' = @{ Name = 'CBS_E_PENDING';                        Kind = 'Reboot'; Meaning = 'Servicing operations are waiting for a restart.' }
    '0x80070bc9' = @{ Name = 'ERROR_FAIL_REBOOT_REQUIRED';           Kind = 'Reboot'; Meaning = 'A restart is required before this can install.' }
    '0x80070070' = @{ Name = 'ERROR_DISK_FULL';                      Kind = 'Space';  Meaning = 'Not enough disk space.' }
    '0x80070005' = @{ Name = 'E_ACCESSDENIED';                       Kind = 'Access'; Meaning = 'Access denied -- often security software or a permissions change.' }
    '0x80070020' = @{ Name = 'ERROR_SHARING_VIOLATION';              Kind = 'Access'; Meaning = 'A file was in use -- often antivirus scanning it.' }
    '0x80070002' = @{ Name = 'ERROR_FILE_NOT_FOUND';                 Kind = 'Client'; Meaning = 'A file was not found -- usually a damaged download cache.' }
    '0x80070422' = @{ Name = 'ERROR_SERVICE_DISABLED';               Kind = 'Client'; Meaning = 'An update service is disabled.' }
    '0x8024402c' = @{ Name = 'WU_E_PT_WINHTTP_NAME_NOT_RESOLVED';    Kind = 'Client'; Meaning = 'The update server name did not resolve.' }
    '0x80244022' = @{ Name = 'WU_E_PT_HTTP_STATUS_SERVICE_UNAVAIL';  Kind = 'Client'; Meaning = 'The update server returned 503 Service Unavailable.' }
    '0x80240438' = @{ Name = 'WU_E_PT_ENDPOINT_UNREACHABLE';         Kind = 'Client'; Meaning = 'The update service endpoint could not be reached.' }
    '0x80072ee2' = @{ Name = 'ERROR_INTERNET_TIMEOUT';               Kind = 'Client'; Meaning = 'The connection to the update service timed out.' }
    '0x80072efd' = @{ Name = 'ERROR_INTERNET_CANNOT_CONNECT';        Kind = 'Client'; Meaning = 'The update service could not be connected to.' }
    '0x800f0922' = @{ Name = 'CBS_E_INSTALLERS_FAILED';              Kind = 'Other';  Meaning = 'An installer step failed -- commonly a full System Reserved partition, a VPN, or .NET setup.' }
    '0x800705b4' = @{ Name = 'ERROR_TIMEOUT';                        Kind = 'Other';  Meaning = 'The operation timed out.' }
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$SutureFindings = @{
    'StoreRepairable' = @{
        Severity = 'Error'
        Title    = 'Component store is corrupt (repairable)'
        Summary  = 'DISM found corruption in the component store. Updates, features and SFC all draw on the store, so they fail until it is repaired.'
        Remedy   = 'Repair runs DISM /RestoreHealth, then SFC.'
    }
    'StoreNotRepairable' = @{
        Severity = 'Error'
        Title    = 'Component store is corrupt and DISM says it cannot be repaired'
        Summary  = 'DISM reports damage it cannot fix from its current sources.'
        Remedy   = 'Repair with -Source pointing at install media of the same build and edition (WIM:D:\sources\install.wim:<index>). If that fails, an in-place upgrade (setup.exe from the same media, keeping files and apps) rebuilds the store.'
    }
    'StoreCheckFailed' = @{
        Severity = 'Warning'
        Title    = 'Component store health could not be read'
        Summary  = 'DISM did not run, or returned output S.U.T.U.R.E. could not read.'
        Remedy   = 'Run DISM /Online /Cleanup-Image /ScanHealth by hand and read %WINDIR%\Logs\DISM\dism.log.'
    }
    'CleanupRecommended' = @{
        Severity = 'Info'
        Title    = 'Component store cleanup recommended'
        Summary  = 'DISM reports superseded packages that a cleanup would remove, reclaiming space on the system drive.'
        Remedy   = 'Repair with -Cleanup runs DISM /StartComponentCleanup. S.U.T.U.R.E. never runs /ResetBase.'
    }
    'SfcUnrepairedFiles' = @{
        Severity = 'Error'
        Title    = 'System File Checker found files it could not repair'
        Summary  = 'The last sfc /scannow left damaged system files in place, usually because the component store was damaged too.'
        Remedy   = 'Repair runs RestoreHealth first so SFC has a healthy store to copy from, then SFC again.'
    }
    'SfcRepairedFiles' = @{
        Severity = 'Info'
        Title    = 'System File Checker repaired files'
        Summary  = 'The last sfc /scannow found and replaced damaged system files.'
        Remedy   = 'Nothing further unless the same files are damaged again -- recurring damage points at the disk (A.U.G.U.R.) or memory.'
    }
    'RebootPendingServicing' = @{
        Severity = 'Warning'
        Title    = 'A restart is pending for servicing'
        Summary  = 'Windows has servicing work waiting for a restart. Updates and DISM repairs fail or roll back until it happens (CBS_E_PENDING).'
        Remedy   = 'Restart first, then audit again. Repair stops before DISM unless the technician chooses to continue.'
    }
    'RebootPendingOther' = @{
        Severity = 'Info'
        Title    = 'A restart is pending (not servicing)'
        Summary  = 'File renames or a computer rename are waiting for a restart. These do not block servicing on their own.'
        Remedy   = 'Restart at a convenient point.'
    }
    'UpdateFailuresStore' = @{
        Severity = 'Error'
        Title    = 'Updates are failing because of the component store'
        Summary  = 'Recent Windows Update installs failed with codes that point at component-store damage.'
        Remedy   = 'Repair, then retry the update (R.E.S.T.O.R.A.T.I.O.N.). If RestoreHealth reports 0x800f081f, give it a -Source.'
    }
    'UpdateFailuresClient' = @{
        Severity = 'Warning'
        Title    = 'Updates are failing in the update client, not the store'
        Summary  = 'Recent Windows Update failures point at connectivity, disabled services or the download cache.'
        Remedy   = 'Run C.O.N.D.U.I.T. -- it repairs the update client. A store repair will not fix these.'
    }
    'UpdateFailuresOther' = @{
        Severity = 'Warning'
        Title    = 'Updates are failing for other reasons'
        Summary  = 'Recent Windows Update installs failed with codes that point at restarts, space, access or installer steps.'
        Remedy   = 'Read each code''s meaning in the update failures table and address it before retrying.'
    }
    'TrustedInstallerDisabled' = @{
        Severity = 'Error'
        Title    = 'Windows Modules Installer service is disabled'
        Summary  = 'TrustedInstaller installs every update and feature. Disabled, all servicing fails.'
        Remedy   = 'Repair sets it back to Manual, its default.'
    }
    'DiskSpaceCritical' = @{
        Severity = 'Error'
        Title    = 'System drive almost out of space'
        Summary  = 'Servicing stages its payloads on the system drive; with this little free space installs and repairs fail.'
        Remedy   = 'Free space first: C.L.E.A.N.S.E. clears temp and update caches.'
    }
    'DiskSpaceLow' = @{
        Severity = 'Warning'
        Title    = 'System drive low on space'
        Summary  = 'Feature updates and large cumulative updates need several GB of free space to stage.'
        Remedy   = 'Run C.L.E.A.N.S.E.; Repair with -Cleanup also reclaims superseded components.'
    }
    'WsusRepairSource' = @{
        Severity = 'Info'
        Title    = 'Repair files will be requested from WSUS'
        Summary  = 'This machine uses WSUS, which does not serve repair content, and policy does not send repairs to Windows Update instead. RestoreHealth then fails with 0x800f081f.'
        Remedy   = 'Repair with -Source pointing at matching install media, or enable "Specify settings for optional component installation and component repair" with "Download repair content directly from Windows Update".'
    }
    'SourceMissing' = @{
        Severity = 'Error'
        Title    = 'RestoreHealth could not find repair files'
        Summary  = 'DISM failed with 0x800f081f: neither Windows Update nor the given source had the files it needed.'
        Remedy   = 'Mount install media of the same build and edition and run Repair with -Source WIM:<drive>:\sources\install.wim:<index> (or ESD:...install.esd:<index>). Get-WindowsImage -ImagePath lists the indexes.'
    }
    'RepairFailed' = @{
        Severity = 'Error'
        Title    = 'Repair step failed'
        Summary  = 'DISM or SFC returned an error. The detail names the step and the code.'
        Remedy   = 'Read %WINDIR%\Logs\DISM\dism.log and CBS.log for the failing component. An in-place upgrade is the fallback when the store cannot be repaired.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[object]]::new()

function Add-SutureFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $SutureFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
    })
}

function Add-SutureAction {
    param([Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Status, [string]$Detail = '')
    [void]$Actions.Add([PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Step = $Step; Status = $Status; Detail = $Detail })
}

function Get-SeverityClass {
    param([string]$Severity)
    switch ($Severity) {
        'Error'   { return 'err'  }
        'Warning' { return 'warn' }
        default   { return 'info' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-SutureBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███████╗██╗   ██╗████████╗██╗   ██╗██████╗ ███████╗
  ██╔════╝██║   ██║╚══██╔══╝██║   ██║██╔══██╗██╔════╝
  ███████╗██║   ██║   ██║   ██║   ██║██████╔╝█████╗
  ╚════██║██║   ██║   ██║   ██║   ██║██╔══██╗██╔══╝
  ███████║╚██████╔╝   ██║   ╚██████╔╝██║  ██║███████╗
  ╚══════╝ ╚═════╝    ╚═╝    ╚═════╝ ╚═╝  ╚═╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    S.U.T.U.R.E. — Stitches Up The Update & Repair Engine" -ForegroundColor Cyan
    Write-Host "    Windows Servicing Diagnosis and Repair Tool" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host ""
        Write-Host "    *** DRY RUN — no changes will be applied ***" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS — parsers for tool output. No I/O; the Pester suite calls
# these with captured text.
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-HResultString {
    # Normalises an error code to '0x' + eight lower-case hex digits. Accepts
    # '0x800F081F', '800f081f', or the signed decimal a process exit code
    # carries (-2146498529).
    param($Code)
    if ($null -eq $Code) { return '' }
    $text = "$Code".Trim()
    if ($text -match '^(?i)(?:0x)?([0-9a-f]{8})$' -and $text -notmatch '^-?\d+$') { return '0x' + $Matches[1].ToLowerInvariant() }
    if ($text -match '^(?i)0x([0-9a-f]{1,8})$') { return '0x' + $Matches[1].ToLowerInvariant().PadLeft(8, '0') }
    $n = 0L
    if ([long]::TryParse($text, [ref]$n)) {
        if ($n -lt 0) { $n = $n + 4294967296L }
        if ($n -ge 0 -and $n -le 4294967295L) { return '0x{0:x8}' -f $n }
    }
    return ''
}

function Get-ServicingErrorInfo {
    param($Code)
    $hex = ConvertTo-HResultString -Code $Code
    if (-not $hex) { return [PSCustomObject]@{ Code = "$Code"; Name = 'Unknown'; Kind = 'Other'; Meaning = 'No code was reported.' } }
    $entry = $ServicingErrorCodes[$hex]
    if ($entry) { return [PSCustomObject]@{ Code = $hex; Name = $entry.Name; Kind = $entry.Kind; Meaning = $entry.Meaning } }
    return [PSCustomObject]@{ Code = $hex; Name = $hex; Kind = 'Other'; Meaning = 'Not a code S.U.T.U.R.E. recognises -- search the code with "Windows Update".' }
}

function ConvertFrom-DismHealth {
    # Reads DISM /CheckHealth, /ScanHealth or /RestoreHealth output, run with
    # /English so the phrases are stable across UI languages.
    param([string[]]$Lines)
    $text = ($Lines -join "`n")
    $state = 'Unknown'
    if ($text -match '(?i)cannot be repaired')                          { $state = 'NotRepairable' }
    elseif ($text -match '(?i)component store is repairable')           { $state = 'Repairable' }
    elseif ($text -match '(?i)restore operation completed successfully') { $state = 'Repaired' }
    elseif ($text -match '(?i)no component store corruption detected')  { $state = 'Healthy' }
    $err = [regex]::Match($text, '(?i)Error:\s*(0x[0-9a-f]{1,8}|\d+)')
    return [PSCustomObject]@{
        State     = $state
        ErrorCode = $(if ($err.Success) { ConvertTo-HResultString -Code $err.Groups[1].Value } else { '' })
    }
}

function ConvertFrom-DismAnalyze {
    # DISM /AnalyzeComponentStore: "Name : value" lines.
    param([string[]]$Lines)
    $get = {
        param($label)
        foreach ($l in $Lines) {
            $m = [regex]::Match("$l", '^\s*' + [regex]::Escape($label) + '\s*:\s*(.+?)\s*$')
            if ($m.Success) { return $m.Groups[1].Value }
        }
        return ''
    }
    $reclaimable = & $get 'Number of Reclaimable Packages'
    return [PSCustomObject]@{
        ActualSize          = & $get 'Actual Size of Component Store'
        SharedWithWindows   = & $get 'Shared with Windows'
        BackupsAndDisabled  = & $get 'Backups and Disabled Features'
        CacheAndTemp        = & $get 'Cache and Temporary Data'
        LastCleanup         = & $get 'Date of Last Cleanup'
        ReclaimablePackages = $(if ($reclaimable -match '^\d+$') { [int]$reclaimable } else { $null })
        CleanupRecommended  = ((& $get 'Component Store Cleanup Recommended') -match '(?i)^yes')
    }
}

function Get-CbsQuotedPath {
    # CBS.log writes a path either bare (\??\C:\Windows\...\foo.dll from store)
    # or as quoted, length-prefixed fragments ([l:23]"\??\C:\Windows\System32"\[l:10]"foo.dll").
    # Joins the fragments and drops the \??\ prefix.
    param([string]$Fragment)
    $text = ($Fragment -replace '(?i)\s+(?:from store|of \S.*|; source file.*)$', '').Trim()
    $parts = @([regex]::Matches($text, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value.TrimEnd('\') })
    $path = if ($parts.Count -gt 0) { $parts -join '\' } else { ($text -split '\s+')[0] }
    return ($path -replace '^\\\?\?\\', '')
}

function ConvertFrom-CbsSrLine {
    # System File Checker writes [SR] lines to CBS.log. Returns the files it
    # repaired and the files it could not, de-duplicated, plus the time of the
    # last [SR] line.
    param([string[]]$Lines)
    $repaired = [System.Collections.Generic.List[string]]::new()
    $failed   = [System.Collections.Generic.List[string]]::new()
    $last     = ''
    $sawSr    = $false
    foreach ($l in $Lines) {
        $line = "$l"
        if ($line -notmatch '\[SR\]') { continue }
        $sawSr = $true
        $ts = [regex]::Match($line, '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})')
        if ($ts.Success) { $last = $ts.Groups[1].Value }
        $cannot = [regex]::Match($line, '(?i)(?:Cannot repair member file|Could not reproject corrupted file)\s+(.+)$')
        if ($cannot.Success) {
            $name = Get-CbsQuotedPath -Fragment $cannot.Groups[1].Value
            if ($name -and -not $failed.Contains($name)) { [void]$failed.Add($name) }
            continue
        }
        $fixed = [regex]::Match($line, '(?i)Repairing corrupted file\s+(.+)$')
        if ($fixed.Success) {
            $name = Get-CbsQuotedPath -Fragment $fixed.Groups[1].Value
            if ($name -and -not $repaired.Contains($name)) { [void]$repaired.Add($name) }
        }
    }
    return [PSCustomObject]@{
        Found        = $sawSr
        Repaired     = $repaired.ToArray()
        CannotRepair = $failed.ToArray()
        LastRun      = $last
    }
}

function ConvertFrom-SfcOutput {
    # sfc writes UTF-16 to the console, which arrives with a NUL after every
    # character; strip them before matching. English text only -- CBS.log is
    # the language-neutral source, this is the fallback.
    param([string[]]$Lines)
    $text = (($Lines -join "`n") -replace "`0", '')
    if ($text -match '(?i)did not find any integrity violations')  { return 'Clean' }
    if ($text -match '(?i)unable to fix some of them')             { return 'Unrepaired' }
    if ($text -match '(?i)successfully repaired')                  { return 'Repaired' }
    if ($text -match '(?i)could not perform the requested operation') { return 'Failed' }
    return 'Unknown'
}

function Get-PendingRebootReason {
    # Takes the signal readings and returns one row per pending restart.
    # Servicing = blocks DISM and update installs; Other = does not.
    param([hashtable]$Signals)
    $rows = [System.Collections.Generic.List[object]]::new()
    $map = [ordered]@{
        CbsRebootPending     = @{ Kind = 'Servicing'; Text = 'Component Based Servicing: RebootPending' }
        CbsPackagesPending   = @{ Kind = 'Servicing'; Text = 'Component Based Servicing: PackagesPending' }
        PendingXml           = @{ Kind = 'Servicing'; Text = 'WinSxS\pending.xml present' }
        WuRebootRequired     = @{ Kind = 'Servicing'; Text = 'Windows Update: RebootRequired' }
        FileRenames          = @{ Kind = 'Other';     Text = 'Session Manager: PendingFileRenameOperations' }
        ComputerRename       = @{ Kind = 'Other';     Text = 'Computer rename waiting for a restart' }
    }
    foreach ($k in $map.Keys) {
        if ($Signals.ContainsKey($k) -and $Signals[$k]) { [void]$rows.Add([PSCustomObject]@{ Signal = $k; Kind = $map[$k].Kind; Text = $map[$k].Text }) }
    }
    return $rows.ToArray()
}

function Get-RepairSourceState {
    # Where DISM looks for repair files. A WSUS-managed machine asks WSUS,
    # which does not carry repair content, unless policy sends repairs to
    # Windows Update (RepairContentServerSource = 2) or names a local source.
    param([bool]$WsusConfigured, $RepairContentServerSource, [string]$LocalSourcePath, $UseWindowsUpdate)
    $toWu   = ("$RepairContentServerSource" -eq '2')
    $local  = -not [string]::IsNullOrWhiteSpace($LocalSourcePath)
    $neverWu = ("$UseWindowsUpdate" -eq '2')
    $text = if ($local) { "Policy source: $LocalSourcePath" }
            elseif ($WsusConfigured -and $toWu) { 'Windows Update (policy overrides WSUS for repairs)' }
            elseif ($WsusConfigured) { 'WSUS -- which does not serve repair content' }
            elseif ($neverWu) { 'None -- policy blocks Windows Update and names no source' }
            else { 'Windows Update' }
    return [PSCustomObject]@{
        Description = $text
        AtRisk      = (-not $local) -and (($WsusConfigured -and -not $toWu) -or $neverWu)
    }
}

function Get-SutureVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Broken';    Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Attention'; Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Native {
    # Runs a native tool, optionally echoing its output, and returns the lines
    # and exit code. A missing binary or a non-zero exit never aborts the run.
    param([string]$File, [string[]]$Arguments, [switch]$Echo)
    $lines = [System.Collections.Generic.List[string]]::new()
    $code  = $null
    try {
        & $File @Arguments 2>&1 | ForEach-Object {
            $l = ("$_" -replace "`0", '')
            [void]$lines.Add($l)
            if ($Echo -and -not [string]::IsNullOrWhiteSpace($l)) { Write-Host "      $($l.Trim())" -ForegroundColor $C.Info }
        }
        $code = $LASTEXITCODE
    } catch {
        [void]$lines.Add("failed: $($_.Exception.Message)")
        $code = -1
    }
    return [PSCustomObject]@{ Lines = $lines.ToArray(); ExitCode = $code }
}

function Get-EventsSafe {
    param([hashtable]$Filter, [int]$MaxEvents = 200)
    try { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $MaxEvents -ErrorAction Stop) } catch { return @() }
}

function Get-EventDataMap {
    param($Record)
    $map = @{}
    try {
        $xml = [xml]$Record.ToXml()
        foreach ($d in @($xml.Event.EventData.Data)) { if ($d -and $d.Name) { $map[$d.Name] = $d.'#text' } }
    } catch { return $map }
    return $map
}

function Read-LogTail {
    # The last $MaxBytes of a log another process may be writing, as lines.
    param([string]$Path, [long]$MaxBytes, [long]$FromOffset = -1)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $fs = $null; $reader = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
        $start = if ($FromOffset -ge 0) { [math]::Min($FromOffset, $fs.Length) } else { [math]::Max(0, $fs.Length - $MaxBytes) }
        [void]$fs.Seek($start, [System.IO.SeekOrigin]::Begin)
        $reader = [System.IO.StreamReader]::new($fs)
        $text = $reader.ReadToEnd()
        return @($text -split "`r?`n")
    } catch {
        return @()
    } finally {
        if ($reader) { $reader.Dispose() } elseif ($fs) { $fs.Dispose() }
    }
}

function Get-LogLength {
    param([string]$Path)
    try { return (Get-Item -LiteralPath $Path -ErrorAction Stop).Length } catch { return 0 }
}

function Get-SutureContext {
    $os = $null
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }
    $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $drive = $null
    try { $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'" -ErrorAction Stop } catch { $drive = $null }
    $build = if ($cv) { "$($cv.CurrentBuild).$($cv.UBR)" } elseif ($os) { "$($os.BuildNumber)" } else { 'Unknown' }
    return [PSCustomObject]@{
        Computer  = $env:COMPUTERNAME
        OS        = $(if ($os) { "$($os.Caption)" } else { 'Unknown' })
        Build     = $build
        Display   = $(if ($cv -and $cv.DisplayVersion) { "$($cv.DisplayVersion)" } else { '' })
        Edition   = $(if ($cv) { "$($cv.EditionID)" } else { '' })
        FreeBytes = $(if ($drive) { [double]$drive.FreeSpace } else { $null })
        SizeBytes = $(if ($drive) { [double]$drive.Size } else { $null })
    }
}

function Get-SutureRebootSignal {
    $sm = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction SilentlyContinue
    $renames = @()
    if ($sm) { $renames = @($sm.PendingFileRenameOperations) + @($sm.PendingFileRenameOperations2) | Where-Object { -not [string]::IsNullOrWhiteSpace("$_") } }
    $active  = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction SilentlyContinue).ComputerName
    $pending = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction SilentlyContinue).ComputerName
    return @{
        CbsRebootPending   = (Test-Path "$CbsKey\RebootPending")
        CbsPackagesPending = (Test-Path "$CbsKey\PackagesPending")
        PendingXml         = (Test-Path (Join-Path $env:SystemRoot 'WinSxS\pending.xml'))
        WuRebootRequired   = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
        FileRenames        = (@($renames).Count -gt 0)
        ComputerRename     = ($active -and $pending -and ($active -ne $pending))
    }
}

function Get-SutureUpdateFailure {
    $since = (Get-Date).AddDays(-$UpdateLookbackDays)
    $rows = foreach ($e in (Get-EventsSafe -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WindowsUpdateClient'; Id = 20; StartTime = $since } -MaxEvents 200)) {
        $d = Get-EventDataMap -Record $e
        $info = Get-ServicingErrorInfo -Code $d['errorCode']
        [PSCustomObject]@{
            Time    = $e.TimeCreated
            Update  = "$($d['updateTitle'])"
            Code    = $info.Code
            Name    = $info.Name
            Kind    = $info.Kind
            Meaning = $info.Meaning
        }
    }
    return @($rows | Sort-Object Time -Descending)
}

function Get-SutureRepairSource {
    $svc = Get-ItemProperty -Path $ServicingPolicyKey -ErrorAction SilentlyContinue
    $wu  = Get-ItemProperty -Path $WuPolicyKey -ErrorAction SilentlyContinue
    $au  = Get-ItemProperty -Path "$WuPolicyKey\AU" -ErrorAction SilentlyContinue
    $wsus = ($wu -and -not [string]::IsNullOrWhiteSpace("$($wu.WUServer)") -and $au -and $au.UseWUServer -eq 1)
    $state = Get-RepairSourceState -WsusConfigured $wsus `
        -RepairContentServerSource $(if ($svc) { $svc.RepairContentServerSource } else { $null }) `
        -LocalSourcePath $(if ($svc) { "$($svc.LocalSourcePath)" } else { '' }) `
        -UseWindowsUpdate $(if ($svc) { $svc.UseWindowsUpdate } else { $null })
    return [PSCustomObject]@{ Wsus = $wsus; WsusServer = $(if ($wu) { "$($wu.WUServer)" } else { '' }); Description = $state.Description; AtRisk = $state.AtRisk }
}

function Get-DismArgumentList {
    param([string]$Operation)
    $a = @('/English', '/Online', '/Cleanup-Image', "/$Operation")
    if ($Operation -eq 'RestoreHealth' -and -not [string]::IsNullOrWhiteSpace($Source)) { $a += @("/Source:$Source", '/LimitAccess') }
    return $a
}

function Invoke-SutureStoreCheck {
    param([switch]$Scan)
    $op = if ($Scan) { 'ScanHealth' } else { 'CheckHealth' }
    if ($Scan) { Write-Step "DISM /ScanHealth -- reads the whole store, usually 5-15 minutes..." } else { Write-Step "DISM /CheckHealth..." }
    $r = Invoke-Native -File 'dism.exe' -Arguments (Get-DismArgumentList -Operation $op) -Echo:$Scan
    $parsed = ConvertFrom-DismHealth -Lines $r.Lines
    return [PSCustomObject]@{ Operation = $op; State = $parsed.State; ErrorCode = $(if ($parsed.ErrorCode) { $parsed.ErrorCode } elseif ($r.ExitCode) { ConvertTo-HResultString -Code $r.ExitCode } else { '' }); ExitCode = $r.ExitCode }
}

function Invoke-SutureAudit {
    Write-Section "SYSTEM"
    $ctx = Get-SutureContext
    Write-Info "$($ctx.Computer)  |  $($ctx.OS) $($ctx.Display)  |  build $($ctx.Build)"
    if ($null -ne $ctx.FreeBytes) {
        $freeGB = [math]::Round($ctx.FreeBytes / 1GB, 1)
        Write-Info "$($env:SystemDrive) free: $(Format-Bytes $ctx.FreeBytes)"
        if ($freeGB -lt $FreeSpaceErrorGB)       { Add-SutureFinding -Code 'DiskSpaceCritical' -Detail "$freeGB GB free on $($env:SystemDrive)" }
        elseif ($freeGB -lt $FreeSpaceWarningGB) { Add-SutureFinding -Code 'DiskSpaceLow'      -Detail "$freeGB GB free on $($env:SystemDrive)" }
    }

    $ti = $null
    try { $ti = Get-CimInstance Win32_Service -Filter "Name='TrustedInstaller'" -ErrorAction Stop } catch { $ti = $null }
    if ($ti -and $ti.StartMode -eq 'Disabled') { Add-SutureFinding -Code 'TrustedInstallerDisabled' -Detail 'TrustedInstaller start type: Disabled' }
    Write-Info ("Windows Modules Installer: {0}" -f $(if ($ti) { "$($ti.StartMode), $($ti.State)" } else { 'not found' }))

    Write-Section "PENDING RESTART"
    $reboot = @(Get-PendingRebootReason -Signals (Get-SutureRebootSignal))
    if ($reboot.Count -eq 0) { Write-Ok "No restart pending." }
    foreach ($r in $reboot) { if ($r.Kind -eq 'Servicing') { Write-Warn $r.Text } else { Write-Info $r.Text } }
    $servicingReboot = @($reboot | Where-Object { $_.Kind -eq 'Servicing' })
    if ($servicingReboot.Count -gt 0) { Add-SutureFinding -Code 'RebootPendingServicing' -Detail (($servicingReboot | ForEach-Object { $_.Text }) -join '; ') }
    elseif ($reboot.Count -gt 0)      { Add-SutureFinding -Code 'RebootPendingOther'     -Detail (($reboot | ForEach-Object { $_.Text }) -join '; ') }

    Write-Section "COMPONENT STORE"
    $store = Invoke-SutureStoreCheck -Scan:$Deep
    switch ($store.State) {
        'Healthy'       { Write-Ok "No component store corruption detected ($($store.Operation))." }
        'Repairable'    { Write-Fail "Component store is corrupt and repairable."; Add-SutureFinding -Code 'StoreRepairable' -Detail "DISM /$($store.Operation)" }
        'NotRepairable' { Write-Fail "Component store is corrupt and cannot be repaired from current sources."; Add-SutureFinding -Code 'StoreNotRepairable' -Detail "DISM /$($store.Operation)" }
        default {
            Write-Warn "Could not read the store state."
            $info = Get-ServicingErrorInfo -Code $store.ErrorCode
            Add-SutureFinding -Code 'StoreCheckFailed' -Detail ("DISM /{0} exit {1}{2}" -f $store.Operation, $store.ExitCode, $(if ($store.ErrorCode) { " -- $($info.Code) $($info.Name)" } else { '' }))
        }
    }
    if (-not $Deep -and $store.State -eq 'Healthy') { Write-Info "CheckHealth reads the last recorded state; -Deep runs a full ScanHealth." }

    Write-Step "DISM /AnalyzeComponentStore..."
    $analyze = ConvertFrom-DismAnalyze -Lines (Invoke-Native -File 'dism.exe' -Arguments @('/English', '/Online', '/Cleanup-Image', '/AnalyzeComponentStore')).Lines
    if ($analyze.ActualSize) { Write-Info ("Store size {0}; reclaimable packages {1}; cleanup recommended: {2}" -f $analyze.ActualSize, $analyze.ReclaimablePackages, $(if ($analyze.CleanupRecommended) { 'yes' } else { 'no' })) }
    if ($analyze.CleanupRecommended) { Add-SutureFinding -Code 'CleanupRecommended' -Detail ("{0} reclaimable package(s); last cleanup {1}" -f $analyze.ReclaimablePackages, $analyze.LastCleanup) }

    $src = Get-SutureRepairSource
    Write-Info "Repair files come from: $($src.Description)"
    if ($src.AtRisk -and [string]::IsNullOrWhiteSpace($Source)) { Add-SutureFinding -Code 'WsusRepairSource' -Detail $src.Description }

    Write-Section "SYSTEM FILE CHECKER (last run)"
    $sfc = ConvertFrom-CbsSrLine -Lines (Read-LogTail -Path $CbsLogPath -MaxBytes $CbsTailBytes)
    if (-not $sfc.Found) { Write-Info "No SFC run found in the recent CBS.log." }
    else {
        Write-Info ("Last SFC activity {0}: {1} repaired, {2} could not be repaired" -f $sfc.LastRun, $sfc.Repaired.Count, $sfc.CannotRepair.Count)
        if ($sfc.CannotRepair.Count -gt 0) { Add-SutureFinding -Code 'SfcUnrepairedFiles' -Detail (($sfc.CannotRepair | Select-Object -First 5) -join ', ') }
        elseif ($sfc.Repaired.Count -gt 0) { Add-SutureFinding -Code 'SfcRepairedFiles' -Detail (($sfc.Repaired | Select-Object -First 5) -join ', ') }
    }

    Write-Section "WINDOWS UPDATE FAILURES ($UpdateLookbackDays days)"
    $failures = @(Get-SutureUpdateFailure)
    if ($failures.Count -eq 0) { Write-Ok "No failed update installs logged." }
    foreach ($f in ($failures | Select-Object -First 8)) { Write-Info ("{0}  {1} {2}  {3}" -f $f.Time.ToString('yyyy-MM-dd'), $f.Code, $f.Kind, $f.Update) }
    foreach ($kind in 'Store', 'Client') {
        $hit = @($failures | Where-Object { $_.Kind -eq $kind })
        if ($hit.Count -gt 0) {
            $codes = ($hit | Group-Object Code | ForEach-Object { "$($_.Name) $($_.Group[0].Name) x$($_.Count)" }) -join '; '
            Add-SutureFinding -Code $(if ($kind -eq 'Store') { 'UpdateFailuresStore' } else { 'UpdateFailuresClient' }) -Detail $codes
        }
    }
    $other = @($failures | Where-Object { $_.Kind -notin @('Store', 'Client') })
    if ($other.Count -gt 0) {
        Add-SutureFinding -Code 'UpdateFailuresOther' -Detail (($other | Group-Object Code | ForEach-Object { "$($_.Name) $($_.Group[0].Name) x$($_.Count)" }) -join '; ')
    }

    return [PSCustomObject]@{
        Context      = $ctx
        TrustedInstaller = $ti
        Reboot       = $reboot
        Store        = $store
        Analyze      = $analyze
        RepairSource = $src
        Sfc          = $sfc
        Failures     = $failures
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Repair-SutureTrustedInstaller {
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would set the Windows Modules Installer (TrustedInstaller) service to Manual."
        Add-SutureAction -Step 'Re-enable Windows Modules Installer' -Status 'WhatIf'
        return
    }
    try {
        Set-Service -Name 'TrustedInstaller' -StartupType Manual -ErrorAction Stop
        Write-Ok "Windows Modules Installer set to Manual."
        Add-SutureAction -Step 'Re-enable Windows Modules Installer' -Status 'Done' -Detail 'Start type: Manual'
        Add-TKNote -Text 'Set the Windows Modules Installer service back to Manual.' -Category 'Action' -ScriptName 'suture'
    } catch {
        Write-Fail "Could not change the service: $($_.Exception.Message)"
        Add-SutureAction -Step 'Re-enable Windows Modules Installer' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Repair-SutureRestoreHealth {
    $argList = Get-DismArgumentList -Operation 'RestoreHealth'
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: dism.exe $($argList -join ' ')"
        Add-SutureAction -Step 'DISM /RestoreHealth' -Status 'WhatIf' -Detail ($argList -join ' ')
        return $true
    }
    Write-Section "REPAIR — DISM /RestoreHealth"
    Write-Step "Repairing the component store$(if ($Source) { " from $Source" } else { ' from Windows Update' }). This usually takes 10-30 minutes; the percentage can sit still for a while."
    $r = Invoke-Native -File 'dism.exe' -Arguments $argList -Echo
    $parsed = ConvertFrom-DismHealth -Lines $r.Lines
    $code = if ($parsed.ErrorCode) { $parsed.ErrorCode } elseif ($r.ExitCode) { ConvertTo-HResultString -Code $r.ExitCode } else { '' }
    if ($r.ExitCode -eq 0 -and $parsed.State -ne 'NotRepairable') {
        Write-Ok "RestoreHealth completed."
        Add-SutureAction -Step 'DISM /RestoreHealth' -Status 'Done' -Detail $(if ($Source) { "Source: $Source" } else { 'Source: Windows Update' })
        Add-TKNote -Text "Ran DISM /RestoreHealth$(if ($Source) { " from $Source" }) -- completed." -Category 'Action' -ScriptName 'suture'
        return $true
    }
    $info = Get-ServicingErrorInfo -Code $code
    Write-Fail "RestoreHealth failed: $($info.Code) $($info.Name) -- $($info.Meaning)"
    if ($info.Code -eq '0x800f081f') { Add-SutureFinding -Code 'SourceMissing' -Detail $(if ($Source) { "Source tried: $Source" } else { 'Source tried: Windows Update' }) }
    else                              { Add-SutureFinding -Code 'RepairFailed' -Detail "DISM /RestoreHealth: $($info.Code) $($info.Name)" }
    Add-SutureAction -Step 'DISM /RestoreHealth' -Status 'Failed' -Detail "$($info.Code) $($info.Name) -- $($info.Meaning)"
    Write-TKError -ScriptName 'suture' -Message "DISM /RestoreHealth failed: $($info.Code) $($info.Name)" -Category 'Servicing'
    return $false
}

function Repair-SutureSfc {
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: sfc.exe /scannow"
        Add-SutureAction -Step 'sfc /scannow' -Status 'WhatIf'
        return $null
    }
    Write-Section "REPAIR — sfc /scannow"
    Write-Step "Verifying system files against the store. This usually takes 10-20 minutes."
    $offset = Get-LogLength -Path $CbsLogPath
    $r = Invoke-Native -File 'sfc.exe' -Arguments @('/scannow')
    $fromLog = ConvertFrom-CbsSrLine -Lines (Read-LogTail -Path $CbsLogPath -MaxBytes $CbsTailBytes -FromOffset $offset)
    $outcome = if ($fromLog.Found) {
        if ($fromLog.CannotRepair.Count -gt 0) { 'Unrepaired' } elseif ($fromLog.Repaired.Count -gt 0) { 'Repaired' } else { 'Clean' }
    } else { ConvertFrom-SfcOutput -Lines $r.Lines }

    switch ($outcome) {
        'Clean'      { Write-Ok "SFC found no integrity violations."; $status = 'Done'; $detail = 'No integrity violations' }
        'Repaired'   { Write-Ok "SFC repaired damaged files."; $status = 'Done'; $detail = "Repaired: $(($fromLog.Repaired | Select-Object -First 5) -join ', ')" }
        'Unrepaired' {
            Write-Fail "SFC could not repair some files."
            $status = 'Failed'; $detail = "Could not repair: $(($fromLog.CannotRepair | Select-Object -First 5) -join ', ')"
            Add-SutureFinding -Code 'RepairFailed' -Detail "sfc /scannow: $detail"
        }
        default {
            Write-Warn "Could not read the SFC result (exit $($r.ExitCode)); check CBS.log."
            $status = 'Failed'; $detail = "Result unreadable, exit $($r.ExitCode)"
        }
    }
    Add-SutureAction -Step 'sfc /scannow' -Status $status -Detail $detail
    Add-TKNote -Text "Ran sfc /scannow -- $detail." -Category 'Action' -ScriptName 'suture'
    return $outcome
}

function Repair-SutureCleanup {
    $argList = @('/English', '/Online', '/Cleanup-Image', '/StartComponentCleanup')
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: dism.exe $($argList -join ' ')  (never /ResetBase)"
        Add-SutureAction -Step 'DISM /StartComponentCleanup' -Status 'WhatIf'
        return
    }
    Write-Section "REPAIR — component cleanup"
    Write-Step "Removing superseded components. Can take 10+ minutes."
    $r = Invoke-Native -File 'dism.exe' -Arguments $argList -Echo
    if ($r.ExitCode -eq 0) {
        Write-Ok "Component cleanup completed."
        Add-SutureAction -Step 'DISM /StartComponentCleanup' -Status 'Done'
        Add-TKNote -Text 'Ran DISM /StartComponentCleanup.' -Category 'Action' -ScriptName 'suture'
    } else {
        $info = Get-ServicingErrorInfo -Code $r.ExitCode
        Write-Warn "Component cleanup failed: $($info.Code) $($info.Name)"
        Add-SutureAction -Step 'DISM /StartComponentCleanup' -Status 'Failed' -Detail "$($info.Code) $($info.Name)"
    }
}

function Invoke-SutureRepair {
    param([object]$Audit)
    $codes = @($Findings | ForEach-Object { $_.Code })

    if ($codes -contains 'TrustedInstallerDisabled') { Repair-SutureTrustedInstaller }

    if ($codes -contains 'RebootPendingServicing' -and -not $WhatIf) {
        $proceed = $false
        if (-not $Unattended) {
            Write-Warn "A restart is pending for servicing. DISM repairs usually fail (CBS_E_PENDING) until it happens."
            $ans = Read-Host "  Continue with the repair anyway? [y/N]"
            $proceed = ($ans -match '^[Yy]')
        }
        if (-not $proceed) {
            Write-Warn "Restart the machine, then run the repair again."
            Add-SutureAction -Step 'DISM /RestoreHealth and sfc /scannow' -Status 'Skipped' -Detail 'A servicing restart is pending -- restart first.'
            return
        }
    }

    if ($codes -contains 'DiskSpaceCritical') { Write-Warn "The system drive is almost full; DISM may fail for lack of space." }

    $restored = Repair-SutureRestoreHealth
    # SFC copies from the store, so it is only worth running once the store is
    # sound; after a failed RestoreHealth it would report the same damage.
    if ($restored) { [void](Repair-SutureSfc) }
    else           { Add-SutureAction -Step 'sfc /scannow' -Status 'Skipped' -Detail 'RestoreHealth failed; SFC would copy from a damaged store.' }

    $doCleanup = [bool]$Cleanup
    if (-not $doCleanup -and -not $Unattended -and -not $WhatIf -and $Audit.Analyze.CleanupRecommended -and $restored) {
        $ans = Read-Host "  DISM recommends a component cleanup. Run /StartComponentCleanup now? [y/N]"
        $doCleanup = ($ans -match '^[Yy]')
    }
    if ($doCleanup) { Repair-SutureCleanup }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-SutureReport {
    param([object]$Audit, [object]$Verdict, [object]$After)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $ctx        = $Audit.Context

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>No servicing issues found.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
            $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $storeBadge = switch ($Audit.Store.State) { 'Healthy' { 'ok' } 'Repaired' { 'ok' } 'Repairable' { 'err' } 'NotRepairable' { 'err' } default { 'warn' } }
    $afterText  = if ($After) { "$($After.State) (DISM /$($After.Operation) after repair)" } else { '' }

    $rRows = [System.Text.StringBuilder]::new()
    if ($Audit.Reboot.Count -eq 0) { [void]$rRows.Append("<tr><td colspan='2'>No restart pending.</td></tr>") }
    foreach ($r in $Audit.Reboot) {
        $badge = if ($r.Kind -eq 'Servicing') { "<span class='tk-badge-warn'>Blocks servicing</span>" } else { "<span class='tk-badge-info'>Other</span>" }
        [void]$rRows.Append("<tr><td>$(EscHtml $r.Text)</td><td>$badge</td></tr>")
    }

    $sfcRows = [System.Text.StringBuilder]::new()
    foreach ($file in $Audit.Sfc.CannotRepair) { [void]$sfcRows.Append("<tr><td class='tk-mono'>$(EscHtml $file)</td><td><span class='tk-badge-err'>Could not repair</span></td></tr>") }
    foreach ($file in $Audit.Sfc.Repaired)     { [void]$sfcRows.Append("<tr><td class='tk-mono'>$(EscHtml $file)</td><td><span class='tk-badge-ok'>Repaired</span></td></tr>") }
    if ($sfcRows.Length -eq 0) {
        [void]$sfcRows.Append("<tr><td colspan='2'>$(if ($Audit.Sfc.Found) { 'The last SFC run recorded no repaired or unrepairable files.' } else { 'No SFC run found in the recent CBS.log.' })</td></tr>")
    }

    $uRows = [System.Text.StringBuilder]::new()
    if ($Audit.Failures.Count -eq 0) { [void]$uRows.Append("<tr><td colspan='4'>No failed update installs in the last $UpdateLookbackDays days.</td></tr>") }
    foreach ($u in ($Audit.Failures | Select-Object -First 40)) {
        $kindBadge = switch ($u.Kind) { 'Store' { 'err' } 'Client' { 'warn' } default { 'info' } }
        [void]$uRows.Append(
            "<tr><td class='tk-mono'>$(EscHtml ($u.Time.ToString('yyyy-MM-dd HH:mm')))</td><td>$(EscHtml $u.Update)</td>" +
            "<td class='tk-mono'>$(EscHtml $u.Code)<br/>$(EscHtml $u.Name)</td>" +
            "<td><span class='tk-badge-$kindBadge'>$(EscHtml $u.Kind)</span> $(EscHtml $u.Meaning)</td></tr>")
    }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Actions.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>") }
    foreach ($a in $Actions) {
        $badge = switch ($a.Status) { 'Done' { 'ok' } 'WhatIf' { 'blue' } 'Failed' { 'err' } default { 'info' } }
        [void]$aRows.Append("<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td><td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>")
    }

    $an = $Audit.Analyze
    $freeText = if ($null -ne $ctx.FreeBytes) { "$(Format-Bytes $ctx.FreeBytes) of $(Format-Bytes $ctx.SizeBytes)" } else { 'unknown' }
    $tiText   = if ($Audit.TrustedInstaller) { "$($Audit.TrustedInstaller.StartMode), $($Audit.TrustedInstaller.State)" } else { 'not found' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'S.U.T.U.R.E. Servicing Report' `
        -ScriptName 'S.U.T.U.R.E.' `
        -Subtitle   "${orgPrefix}Windows Servicing -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Build'     = "$($ctx.Build)$(if ($ctx.Display) { " ($($ctx.Display))" })"
            'Generated' = $reportDate
            'Mode'      = $(if ($WhatIf) { "$Action (dry run)" } else { $Action })
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Component Store', 'Pending Restart', 'System Files', 'Update Failures', 'Actions Taken')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Servicing</div></div>
    <div class="tk-summary-card $storeBadge"><div class="tk-summary-num">$(EscHtml $Audit.Store.State)</div><div class="tk-summary-lbl">Component Store</div></div>
    <div class="tk-summary-card $(if (@($Audit.Reboot | Where-Object { $_.Kind -eq 'Servicing' }).Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$(if (@($Audit.Reboot | Where-Object { $_.Kind -eq 'Servicing' }).Count) { 'Yes' } else { 'No' })</div><div class="tk-summary-lbl">Servicing Restart Pending</div></div>
    <div class="tk-summary-card $(if ($Audit.Sfc.CannotRepair.Count) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($Audit.Sfc.CannotRepair.Count)</div><div class="tk-summary-lbl">Unrepaired Files</div></div>
    <div class="tk-summary-card $(if ($Audit.Failures.Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Audit.Failures.Count)</div><div class="tk-summary-lbl">Failed Updates ($UpdateLookbackDays d)</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Component Store</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">Health</span> $(EscHtml $Audit.Store.State) (DISM /$(EscHtml $Audit.Store.Operation))$(if ($afterText) { "<br/><span class='tk-info-label'>After repair</span> $(EscHtml $afterText)" })<br/>
      <span class="tk-info-label">Store size</span> $(EscHtml $(if ($an.ActualSize) { $an.ActualSize } else { 'not read' }))$(if ($an.BackupsAndDisabled) { " (backups and disabled features $(EscHtml $an.BackupsAndDisabled); cache and temp $(EscHtml $an.CacheAndTemp))" })<br/>
      <span class="tk-info-label">Cleanup</span> $(if ($an.CleanupRecommended) { 'Recommended' } else { 'Not recommended' })$(if ($null -ne $an.ReclaimablePackages) { " -- $($an.ReclaimablePackages) reclaimable package(s)" })$(if ($an.LastCleanup) { "; last cleanup $(EscHtml $an.LastCleanup)" })<br/>
      <span class="tk-info-label">Repair files from</span> $(EscHtml $(if ($Source) { "-Source $Source" } else { $Audit.RepairSource.Description }))<br/>
      <span class="tk-info-label">Windows Modules Installer</span> $(EscHtml $tiText)<br/>
      <span class="tk-info-label">System drive free</span> $(EscHtml $freeText)<br/>
      <span class="tk-info-label">Logs</span> <span class="tk-mono">$(EscHtml $DismLogPath)</span> · <span class="tk-mono">$(EscHtml $CbsLogPath)</span>
    </div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Pending Restart</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Signal</th><th>Effect</th></tr></thead>
      <tbody>$($rRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> System Files</div>
    <div class="tk-info-box"><span class="tk-info-label">Last SFC activity</span> $(EscHtml $(if ($Audit.Sfc.LastRun) { $Audit.Sfc.LastRun } else { 'none found' }))</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>File</th><th>Result</th></tr></thead>
      <tbody>$($sfcRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Update Failures</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>When</th><th>Update</th><th>Code</th><th>Cause</th></tr></thead>
      <tbody>$($uRows.ToString())</tbody></table></div>
    <div class="tk-info-box"><span class="tk-info-label">Reading the cause</span> Store = component-store damage, fixed here. Client = the update client's connectivity or services, fixed by C.O.N.D.U.I.T. Reboot, Space and Access are fixed by a restart, free space, and security-software exclusions.</div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Actions Taken</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'S.U.T.U.R.E. v5.1')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-SutureRun {
    param([string]$Mode)
    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-SutureAudit
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Component store, system files and servicing state all in order." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    $after = $null
    if ($Mode -eq 'Repair') {
        Invoke-SutureRepair -Audit $audit
        $ran = @($Actions | Where-Object { $_.Step -eq 'DISM /RestoreHealth' -and $_.Status -eq 'Done' }).Count -gt 0
        if ($ran) {
            Write-Section "VERIFY"
            $after = Invoke-SutureStoreCheck
            if ($after.State -eq 'Healthy') { Write-Ok "Component store reports healthy after repair." }
            else                            { Write-Warn "Component store still reports: $($after.State)" }
            Add-TKNote -Text "SUTURE repair on $($env:COMPUTERNAME): component store now $($after.State)." -Category $(if ($after.State -eq 'Healthy') { 'Resolution' } else { 'Issue' }) -ScriptName 'suture'
        }
    }

    $verdict = Get-SutureVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    if ($after -and $after.State -eq 'Healthy' -and $verdict.Class -ne 'ok') { Write-Info "Findings describe the state before the repair; the store now reports healthy. Audit again after a restart." }
    Add-TKNote -Text ("SUTURE {0} on {1}: verdict {2} ({3} finding(s))." -f $Mode, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'suture'

    Write-Step "Generating HTML report..."
    $html    = Build-SutureReport -Audit $audit -Verdict $verdict -After $after
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("SUTURE_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'suture' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-SutureBanner
    Invoke-SutureRun -Mode $Action
} else {
    $choice = ''
    do {
        Show-SutureBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only: store health (CheckHealth), last SFC result, update failures" -ForegroundColor $C.Info
        Write-Host "  [2] Deep audit  -  as [1] with a full DISM /ScanHealth (5-15 minutes)" -ForegroundColor $C.Info
        Write-Host "  [3] Repair  -  audit, then DISM /RestoreHealth, sfc /scannow, and re-check" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($Source) { Write-Host "  Repair source: $Source" -ForegroundColor $C.Info }
        if ($WhatIf) { Write-Host "  Dry run is active — option 3 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-SutureRun -Mode 'Audit' }
            '2' { $script:Deep = $true; Invoke-SutureRun -Mode 'Audit'; $script:Deep = $false }
            '3' { Invoke-SutureRun -Mode 'Repair' }
            'Q' { Write-Host ""; Write-Host "  Closing S.U.T.U.R.E." -ForegroundColor $C.Header; Write-Host "" }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1, 2, 3 or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }
        if ($choice -ne 'Q') {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
    } while ($choice -ne 'Q')
}

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
