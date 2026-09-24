# conduit.ps1 - C.O.N.D.U.I.T. — Checks Or Normalises Device Update Infrastructure Targeting
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
    C.O.N.D.U.I.T. — Checks Or Normalises Device Update Infrastructure Targeting
    Windows Update Connectivity Diagnostics & Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Diagnoses and repairs the "We couldn't connect to the update service" failure
    that stops Windows Update before any download begins. Audits the update
    client's plumbing rather than its content: Windows Time service state, the
    WinHTTP proxy, the WindowsUpdate policy key (WSUS pointer and blocking
    policies), whether that policy comes from domain or local Group Policy,
    reachability of the configured WSUS server and the Microsoft Update
    endpoints, and the start type of the four services the client depends on.

    Audit is read-only. Repair backs up the policy key first, then removes a
    stale WSUS pointer only when the WSUS server is unreachable AND the setting
    is not being pushed by domain or local Group Policy, restores the time
    service, and re-enables any disabled update service. ResetCache additionally
    renames SoftwareDistribution and catroot2 so the client rebuilds them.

    CONDUIT is the precondition check for R.E.S.T.O.R.A.T.I.O.N.: run CONDUIT
    when the update client cannot reach a service at all, and RESTORATION once
    it can, to actually deploy updates.

.USAGE
    PS C:\> .\conduit.ps1                                  # Interactive menu
    PS C:\> .\conduit.ps1 -Unattended                      # Read-only audit + HTML report
    PS C:\> .\conduit.ps1 -Unattended -Action Repair       # Audit, then apply safe repairs
    PS C:\> .\conduit.ps1 -Unattended -Action Repair -WhatIf      # Preview the repairs only
    PS C:\> .\conduit.ps1 -Unattended -Action Repair -Force       # Override the WSUS-removal blockers
    PS C:\> .\conduit.ps1 -Unattended -Action ResetCache          # Repair + rebuild the update cache

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit','Repair','ResetCache')]
    [string]$Action = 'Audit',
    [switch]$Force,
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
# REGISTRY PATHS
# ─────────────────────────────────────────────────────────────────────────────

$WUPolicyKey   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$AUPolicyKey   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
$WUPolicyRegEx = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$LocalPolFile  = Join-Path $env:SystemRoot 'System32\GroupPolicy\Machine\Registry.pol'

# The four services the update client cannot function without, mapped to the
# start type Windows ships them with. Used both to flag a disabled service and
# to restore the correct value -- setting them all to Automatic would be wrong,
# wuauserv and bits are demand-started by design.
$UpdateServiceDefaults = [ordered]@{
    'wuauserv' = 'Manual'
    'bits'     = 'Manual'
    'cryptsvc' = 'Automatic'
    'usosvc'   = 'Automatic'
}

# TCP endpoints the update client reaches when no WSUS server is configured.
# A block on any one of these is a network/filter problem, not a client one.
$MicrosoftUpdateEndpoints = @(
    'update.microsoft.com'
    'windowsupdate.microsoft.com'
    'dl.delivery.mp.microsoft.com'
)

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition CONDUIT can detect, keyed by a stable code. Severity drives
# both the console colour and the report badge class; Remedy is what the
# technician is told to do when CONDUIT will not do it automatically. Kept as
# one table so the Pester suite can extract and assert on it by AST lookup
# without executing the script.
# ─────────────────────────────────────────────────────────────────────────────

$ConduitFindings = @{
    'TimeServiceStopped' = @{
        Severity = 'Warning'
        Title    = 'Windows Time service not running'
        Summary  = 'Clock drift beyond five minutes breaks the TLS handshake to the update service.'
        Remedy   = 'Repair starts w32time, sets it to Automatic, and forces a resync.'
    }
    'WinHttpProxySet' = @{
        Severity = 'Warning'
        Title    = 'WinHTTP proxy configured'
        Summary  = 'The update client uses the WinHTTP proxy, not the per-user Internet Options proxy. A stale entry here blackholes every update request.'
        Remedy   = 'Confirm the proxy is intentional. Clear it manually with: netsh winhttp reset proxy'
    }
    'WsusUnreachable' = @{
        Severity = 'Error'
        Title    = 'WSUS server configured but unreachable'
        Summary  = 'The client is pointed at a WSUS server that does not resolve or does not answer on its port, so it never falls back to Microsoft Update.'
        Remedy   = 'Repair removes the pointer when no Group Policy is reapplying it; otherwise clear the policy at source.'
    }
    'BlockInternetWU' = @{
        Severity = 'Error'
        Title    = 'Policy blocks Microsoft Update endpoints'
        Summary  = 'DoNotConnectToWindowsUpdateInternetLocations = 1 forbids the client from contacting Microsoft Update even when WSUS is absent.'
        Remedy   = 'Clear the policy at its source (domain GPO or gpedit.msc). CONDUIT never edits this value automatically.'
    }
    'WUAccessDisabled' = @{
        Severity = 'Error'
        Title    = 'Windows Update access disabled by policy'
        Summary  = 'DisableWindowsUpdateAccess = 1 removes the update UI and blocks user-initiated scans.'
        Remedy   = 'Clear the policy at its source (domain GPO or gpedit.msc). CONDUIT never edits this value automatically.'
    }
    'AutoUpdateDisabled' = @{
        Severity = 'Warning'
        Title    = 'Automatic updates disabled by policy'
        Summary  = 'NoAutoUpdate = 1 stops scheduled scans. Manual checks still work, so this is context rather than the connection fault itself.'
        Remedy   = 'Expected on servers and maintenance-window builds. Clear at source if unintended.'
    }
    'MicrosoftEndpointsBlocked' = @{
        Severity = 'Error'
        Title    = 'Microsoft Update endpoints unreachable'
        Summary  = 'One or more Microsoft Update hosts refused a TCP connection on 443 -- a network, firewall, or content-filter block upstream of this machine.'
        Remedy   = 'Allow the endpoints through the filter. No local repair can fix an upstream block.'
    }
    'ServiceDisabled' = @{
        Severity = 'Error'
        Title    = 'Update service disabled'
        Summary  = 'One of wuauserv / bits / cryptsvc / usosvc has start type Disabled, so the update client cannot start it on demand.'
        Remedy   = 'Repair restores the service to its Windows default start type.'
    }
    'WUServerUnparsable' = @{
        Severity = 'Error'
        Title    = 'WSUS pointer is not a valid URL'
        Summary  = 'The WUServer policy value could not be parsed as a URI, so the client cannot build a request from it.'
        Remedy   = 'Correct or remove the WUServer value at its policy source.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = New-Object 'System.Collections.Generic.List[object]'
$Actions  = New-Object 'System.Collections.Generic.List[object]'

function Add-ConduitFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Detail = ''
    )
    $meta = $ConduitFindings[$Code]
    if (-not $meta) {
        $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' }
    }
    [void]$Findings.Add([PSCustomObject]@{
        Code     = $Code
        Severity = $meta.Severity
        Title    = $meta.Title
        Summary  = $meta.Summary
        Remedy   = $meta.Remedy
        Detail   = $Detail
    })
}

function Add-ConduitAction {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,
        [string]$Detail = ''
    )
    [void]$Actions.Add([PSCustomObject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Step      = $Step
        Status    = $Status
        Detail    = $Detail
    })
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

function Show-ConduitBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██████╗ ███╗   ██╗██████╗ ██╗   ██╗██╗████████╗
  ██╔════╝██╔═══██╗████╗  ██║██╔══██╗██║   ██║██║╚══██╔══╝
  ██║     ██║   ██║██╔██╗ ██║██║  ██║██║   ██║██║   ██║
  ██║     ██║   ██║██║╚██╗██║██║  ██║██║   ██║██║   ██║
  ╚██████╗╚██████╔╝██║ ╚████║██████╔╝╚██████╔╝██║   ██║
   ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝╚═════╝  ╚═════╝ ╚═╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    C.O.N.D.U.I.T. — Checks Or Normalises Device Update Infrastructure Targeting" -ForegroundColor Cyan
    Write-Host "    Windows Update Connectivity Diagnostics & Repair Tool" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host ""
        Write-Host "    *** DRY RUN — no changes will be applied ***" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-TCPPort {
    # Test-NetConnection takes ~20s to fail closed; a raw socket with an explicit
    # timeout keeps a six-endpoint sweep inside a few seconds.
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 3000)
    $tcp = $null
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok) { $tcp.EndConnect($ar) | Out-Null }
        return ($ok -and $tcp.Connected)
    } catch {
        return $false
    } finally {
        if ($tcp) { $tcp.Close() }
    }
}

function Get-ServiceSafe {
    param([string]$Name)
    try { return Get-Service -Name $Name -ErrorAction Stop }
    catch { return $null }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-ConduitDeviceContext {
    # Context only -- a CIM failure must not abort the diagnosis, since every
    # check that actually matters reads the registry and the service table.
    $cs = $null; $os = $null
    try { $cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop } catch { $cs = $null }
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }

    return [PSCustomObject]@{
        Computer     = $env:COMPUTERNAME
        OS           = if ($os) { "$($os.Caption) (build $($os.BuildNumber))" } else { 'Unknown' }
        Domain       = if ($cs) { $cs.Domain } else { 'Unknown' }
        PartOfDomain = if ($cs) { [bool]$cs.PartOfDomain } else { $false }
    }
}

function Get-ConduitTimeService {
    $svc = Get-ServiceSafe -Name 'w32time'
    if (-not $svc) {
        return [PSCustomObject]@{ Present = $false; Status = 'Absent'; StartType = 'Absent' }
    }
    if ($svc.Status -ne 'Running') { Add-ConduitFinding -Code 'TimeServiceStopped' -Detail "Current status: $($svc.Status)" }
    return [PSCustomObject]@{
        Present   = $true
        Status    = "$($svc.Status)"
        StartType = "$($svc.StartType)"
    }
}

function Get-ConduitProxy {
    $raw = ''
    try { $raw = (netsh winhttp show proxy) -join [Environment]::NewLine } catch { $raw = 'Could not query WinHTTP proxy.' }

    # "Direct access (no proxy server)" is the clean state. Anything else means a
    # proxy is in play for the update client specifically.
    $configured = ($raw -notmatch 'Direct access')
    if ($configured) { Add-ConduitFinding -Code 'WinHttpProxySet' -Detail (($raw -split "`r?`n" | Where-Object { $_.Trim() }) -join ' | ') }

    return [PSCustomObject]@{
        Configured = $configured
        Raw        = $raw.Trim()
    }
}

function Get-ConduitPolicy {
    $wu = Get-ItemProperty -Path $WUPolicyKey -ErrorAction SilentlyContinue
    $au = Get-ItemProperty -Path $AUPolicyKey -ErrorAction SilentlyContinue

    $wsusConfigured = (-not [string]::IsNullOrWhiteSpace($wu.WUServer)) -and ($au.UseWUServer -eq 1)

    if ($wu.DoNotConnectToWindowsUpdateInternetLocations -eq 1) { Add-ConduitFinding -Code 'BlockInternetWU' }
    if ($wu.DisableWindowsUpdateAccess -eq 1)                   { Add-ConduitFinding -Code 'WUAccessDisabled' }
    if ($au.NoAutoUpdate -eq 1)                                 { Add-ConduitFinding -Code 'AutoUpdateDisabled' }

    return [PSCustomObject]@{
        KeyPresent     = [bool](Test-Path $WUPolicyKey)
        WUServer       = "$($wu.WUServer)"
        WUStatusServer = "$($wu.WUStatusServer)"
        UseWUServer    = $au.UseWUServer
        BlockInternet  = $wu.DoNotConnectToWindowsUpdateInternetLocations
        AccessDisabled = $wu.DisableWindowsUpdateAccess
        NoAutoUpdate   = $au.NoAutoUpdate
        WsusConfigured = $wsusConfigured
    }
}

function Get-ConduitPolicySource {
    # Distinguishing "someone wrote the registry once" from "Group Policy pushes
    # this every 90 minutes" is the whole basis for deciding whether removing the
    # WSUS pointer will stick. Registry.pol is Unicode; Select-String needs telling.
    param([bool]$PartOfDomain)

    $localGpoSetsWsus = $false
    if (Test-Path $LocalPolFile) {
        try {
            $localGpoSetsWsus = [bool](Select-String -Path $LocalPolFile -Pattern 'WUServer' -Encoding Unicode -Quiet -ErrorAction Stop)
        } catch {
            $localGpoSetsWsus = $false
        }
    }

    return [PSCustomObject]@{
        LocalGpoSetsWsus = $localGpoSetsWsus
        DomainJoined     = $PartOfDomain
        LocalPolPresent  = (Test-Path $LocalPolFile)
    }
}

function Get-ConduitWsusReachability {
    param([object]$Policy)

    if (-not $Policy.WsusConfigured) {
        return [PSCustomObject]@{ Configured = $false; HostName = ''; Port = 0; DnsOk = $false; Reachable = $null }
    }

    $uri = $null
    try { $uri = [uri]$Policy.WUServer } catch { $uri = $null }
    if (-not $uri -or [string]::IsNullOrWhiteSpace($uri.Host)) {
        Add-ConduitFinding -Code 'WUServerUnparsable' -Detail "WUServer = $($Policy.WUServer)"
        return [PSCustomObject]@{ Configured = $true; HostName = $Policy.WUServer; Port = 0; DnsOk = $false; Reachable = $false }
    }

    $port  = if ($uri.Port -gt 0) { $uri.Port } else { 8530 }
    $dnsOk = $false
    try { $dnsOk = [bool](Resolve-DnsName -Name $uri.Host -ErrorAction Stop) } catch { $dnsOk = $false }

    $reachable = $dnsOk -and (Test-TCPPort -Hostname $uri.Host -Port $port)
    if (-not $reachable) {
        Add-ConduitFinding -Code 'WsusUnreachable' -Detail ("{0}:{1} (DNS={2})" -f $uri.Host, $port, $dnsOk)
    }

    return [PSCustomObject]@{
        Configured = $true
        HostName   = $uri.Host
        Port       = $port
        DnsOk      = $dnsOk
        Reachable  = $reachable
    }
}

function Get-ConduitEndpointReachability {
    $results = foreach ($ep in $MicrosoftUpdateEndpoints) {
        [PSCustomObject]@{
            Endpoint  = $ep
            Port      = 443
            Reachable = (Test-TCPPort -Hostname $ep -Port 443)
        }
    }
    $results = @($results)

    $blocked = @($results | Where-Object { -not $_.Reachable })
    if ($blocked.Count -gt 0) {
        Add-ConduitFinding -Code 'MicrosoftEndpointsBlocked' -Detail (($blocked | ForEach-Object { $_.Endpoint }) -join ', ')
    }
    return $results
}

function Get-ConduitUpdateServices {
    $results = foreach ($name in $UpdateServiceDefaults.Keys) {
        $svc = Get-ServiceSafe -Name $name
        if (-not $svc) {
            [PSCustomObject]@{ Name = $name; Present = $false; Status = 'Absent'; StartType = 'Absent'; Expected = $UpdateServiceDefaults[$name] }
            continue
        }
        if ("$($svc.StartType)" -eq 'Disabled') {
            Add-ConduitFinding -Code 'ServiceDisabled' -Detail "$name is Disabled (expected $($UpdateServiceDefaults[$name]))"
        }
        [PSCustomObject]@{
            Name      = $name
            Present   = $true
            Status    = "$($svc.Status)"
            StartType = "$($svc.StartType)"
            Expected  = $UpdateServiceDefaults[$name]
        }
    }
    return @($results)
}

# ─────────────────────────────────────────────────────────────────────────────
# AUDIT — runs every collector and prints the console summary
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ConduitAudit {
    Write-Section "DEVICE CONTEXT"
    $device = Get-ConduitDeviceContext
    Write-Info ("Computer      : {0}" -f $device.Computer)
    Write-Info ("OS            : {0}" -f $device.OS)
    Write-Info ("Domain        : {0}" -f $device.Domain)
    Write-Info ("Domain-joined : {0}" -f $device.PartOfDomain)

    Write-Section "WINDOWS TIME SERVICE"
    $time = Get-ConduitTimeService
    if (-not $time.Present) {
        Write-Fail "w32time service not present."
    } elseif ($time.Status -eq 'Running') {
        Write-Ok ("w32time is {0} (start type {1})." -f $time.Status, $time.StartType)
    } else {
        Write-Warn ("w32time is {0} (start type {1})." -f $time.Status, $time.StartType)
    }

    Write-Section "WINHTTP PROXY"
    $proxy = Get-ConduitProxy
    foreach ($line in ($proxy.Raw -split "`r?`n")) {
        if ($line.Trim()) { Write-Info $line.Trim() }
    }
    if ($proxy.Configured) {
        Write-Warn "A WinHTTP proxy is configured — confirm it is intentional."
    } else {
        Write-Ok "Direct access — no WinHTTP proxy."
    }

    Write-Section "WINDOWS UPDATE POLICY"
    $policy = Get-ConduitPolicy
    if (-not $policy.KeyPresent) {
        Write-Ok "No WindowsUpdate policy key — the client uses its defaults."
    } else {
        Write-Info ("WUServer                                     : {0}" -f $(if ($policy.WUServer) { $policy.WUServer } else { '(not set)' }))
        Write-Info ("WUStatusServer                               : {0}" -f $(if ($policy.WUStatusServer) { $policy.WUStatusServer } else { '(not set)' }))
        Write-Info ("UseWUServer                                  : {0}" -f $(if ($null -ne $policy.UseWUServer) { $policy.UseWUServer } else { '(not set)' }))
        Write-Info ("DoNotConnectToWindowsUpdateInternetLocations : {0}" -f $(if ($null -ne $policy.BlockInternet) { $policy.BlockInternet } else { '(not set)' }))
        Write-Info ("DisableWindowsUpdateAccess                   : {0}" -f $(if ($null -ne $policy.AccessDisabled) { $policy.AccessDisabled } else { '(not set)' }))
        Write-Info ("NoAutoUpdate                                 : {0}" -f $(if ($null -ne $policy.NoAutoUpdate) { $policy.NoAutoUpdate } else { '(not set)' }))
    }

    Write-Section "POLICY SOURCE"
    $source = Get-ConduitPolicySource -PartOfDomain $device.PartOfDomain
    Write-Info ("Local Group Policy sets WUServer : {0}" -f $source.LocalGpoSetsWsus)
    Write-Info ("Domain GPO possible              : {0}" -f $source.DomainJoined)
    if ($source.DomainJoined -and $policy.WsusConfigured) {
        Write-Info "Tip: gpresult /h C:\temp\gp.html — find the GPO that sets WUServer."
    }

    Write-Section "CONNECTIVITY"
    $wsus = Get-ConduitWsusReachability -Policy $policy
    if (-not $wsus.Configured) {
        Write-Ok "No WSUS server configured — the client targets Microsoft Update directly."
    } elseif ($wsus.Reachable) {
        Write-Ok ("WSUS {0}:{1} reachable (DNS={2})." -f $wsus.HostName, $wsus.Port, $wsus.DnsOk)
    } else {
        Write-Fail ("WSUS {0}:{1} unreachable (DNS={2})." -f $wsus.HostName, $wsus.Port, $wsus.DnsOk)
    }

    $endpoints = Get-ConduitEndpointReachability
    foreach ($ep in $endpoints) {
        if ($ep.Reachable) { Write-Ok   ("{0}:{1} reachable"   -f $ep.Endpoint, $ep.Port) }
        else               { Write-Fail ("{0}:{1} unreachable" -f $ep.Endpoint, $ep.Port) }
    }

    Write-Section "UPDATE SERVICES"
    $services = Get-ConduitUpdateServices
    foreach ($svc in $services) {
        $line = "{0,-10} {1,-10} start={2} (expected {3})" -f $svc.Name, $svc.Status, $svc.StartType, $svc.Expected
        if (-not $svc.Present)                { Write-Fail ("{0,-10} not present" -f $svc.Name) }
        elseif ($svc.StartType -eq 'Disabled'){ Write-Fail $line }
        else                                  { Write-Ok   $line }
    }

    return [PSCustomObject]@{
        Device    = $device
        Time      = $time
        Proxy     = $proxy
        Policy    = $policy
        Source    = $source
        Wsus      = $wsus
        Endpoints = $endpoints
        Services  = $services
    }
}

function Show-ConduitFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) {
        Write-Ok "No issues found — the update client's plumbing looks healthy."
        return
    }
    foreach ($f in $Findings) {
        switch ($f.Severity) {
            'Error'   { Write-Fail ("[{0}] {1}" -f $f.Code, $f.Title) }
            'Warning' { Write-Warn ("[{0}] {1}" -f $f.Code, $f.Title) }
            default   { Write-Info ("[{0}] {1}" -f $f.Code, $f.Title) }
        }
        if ($f.Detail) { Write-Info $f.Detail }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Backup-ConduitPolicy {
    Write-Section "REPAIR — POLICY BACKUP"

    if (-not (Test-Path $WUPolicyKey)) {
        Write-Info "No WindowsUpdate policy key to back up."
        Add-ConduitAction -Step 'Backup policy key' -Status 'Skipped' -Detail 'Key not present'
        return ''
    }

    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backup = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "CONDUIT_WUPolicy_$stamp.reg"

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would export $WUPolicyRegEx to $backup"
        Add-ConduitAction -Step 'Backup policy key' -Status 'WhatIf' -Detail $backup
        return ''
    }

    try {
        $null = reg export $WUPolicyRegEx $backup /y 2>&1
        if ($LASTEXITCODE -ne 0) { throw "reg export returned exit code $LASTEXITCODE" }
        Write-Ok "Policy key backed up."
        Write-Info $backup
        Write-Info "Restore with: reg import `"$backup`""
        Add-ConduitAction -Step 'Backup policy key' -Status 'Done' -Detail $backup
        return $backup
    } catch {
        Write-Fail "Could not back up the policy key: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "Policy key backup failed: $($_.Exception.Message)" -Category 'Registry'
        Add-ConduitAction -Step 'Backup policy key' -Status 'Failed' -Detail $_.Exception.Message
        return ''
    }
}

function Repair-ConduitWsusPointer {
    param([object]$Audit)

    Write-Section "REPAIR — WSUS POINTER"

    if (-not $Audit.Policy.WsusConfigured) {
        Write-Ok "No WSUS pointer present. Nothing to do."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Skipped' -Detail 'No pointer configured'
        return
    }

    # Three independent reasons not to touch the pointer. Each means removal is
    # either destructive (the server is live) or futile (policy reapplies it).
    $blockers = @()
    if ($Audit.Wsus.Reachable)          { $blockers += 'WSUS server is reachable (it may be in active use)' }
    if ($Audit.Device.PartOfDomain)     { $blockers += 'device is domain-joined (a domain GPO will reapply it)' }
    if ($Audit.Source.LocalGpoSetsWsus) { $blockers += 'local Group Policy sets it (clear it in gpedit.msc instead)' }

    if ($blockers.Count -gt 0 -and -not $Force) {
        Write-Warn ("Skipped WSUS removal: {0}" -f ($blockers -join '; '))
        Write-Info "Re-run with -Force to override."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Skipped' -Detail ($blockers -join '; ')
        return
    }

    if ($blockers.Count -gt 0) {
        Write-Warn ("-Force used despite: {0}" -f ($blockers -join '; '))
    }

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would remove UseWUServer, WUServer and WUStatusServer, then restart wuauserv."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'WhatIf' -Detail $Audit.Policy.WUServer
        return
    }

    try {
        Remove-ItemProperty -Path $AUPolicyKey -Name 'UseWUServer' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $WUPolicyKey -Name 'WUServer', 'WUStatusServer' -ErrorAction SilentlyContinue
        Restart-Service -Name 'wuauserv' -Force -ErrorAction Stop
        Write-Ok "Removed the WSUS pointer and restarted wuauserv."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Done' -Detail "Was: $($Audit.Policy.WUServer)"
        Add-TKNote -Text "Removed unreachable WSUS pointer $($Audit.Policy.WUServer) and restarted wuauserv." -Category 'Action'
    } catch {
        Write-Fail "Could not remove the WSUS pointer: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "WSUS pointer removal failed: $($_.Exception.Message)" -Category 'Registry'
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Repair-ConduitTimeService {
    Write-Section "REPAIR — WINDOWS TIME SERVICE"

    $svc = Get-ServiceSafe -Name 'w32time'
    if (-not $svc) {
        Write-Fail "w32time is not present on this machine."
        Add-ConduitAction -Step 'Repair time service' -Status 'Failed' -Detail 'Service absent'
        return
    }

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would set w32time to Automatic, start it, and run: w32tm /resync /force"
        Add-ConduitAction -Step 'Repair time service' -Status 'WhatIf' -Detail "Current: $($svc.Status) / $($svc.StartType)"
        return
    }

    try {
        Set-Service -Name 'w32time' -StartupType Automatic -ErrorAction Stop
        if ((Get-Service -Name 'w32time').Status -ne 'Running') {
            Start-Service -Name 'w32time' -ErrorAction Stop
            Start-Sleep -Seconds 3
        }
        $resync = & w32tm /resync /force 2>&1
        Write-Ok "w32time set to Automatic, running, and resynced."
        Write-Info (($resync | Out-String).Trim())
        Add-ConduitAction -Step 'Repair time service' -Status 'Done' -Detail (($resync | Out-String).Trim())
        Add-TKNote -Text 'Enabled and resynced the Windows Time service.' -Category 'Action'
    } catch {
        Write-Fail "Could not repair the time service: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "Time service repair failed: $($_.Exception.Message)" -Category 'Service'
        Add-ConduitAction -Step 'Repair time service' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Repair-ConduitServices {
    Write-Section "REPAIR — UPDATE SERVICES"

    $changed = $false
    foreach ($name in $UpdateServiceDefaults.Keys) {
        $svc = Get-ServiceSafe -Name $name
        if (-not $svc) {
            Write-Warn "$name is not present — skipping."
            continue
        }
        if ("$($svc.StartType)" -ne 'Disabled') { continue }

        $expected = $UpdateServiceDefaults[$name]
        $changed  = $true

        if ($WhatIf) {
            Write-Warn "[WhatIf] Would set $name start type from Disabled to $expected."
            Add-ConduitAction -Step "Re-enable $name" -Status 'WhatIf' -Detail "Disabled -> $expected"
            continue
        }

        try {
            Set-Service -Name $name -StartupType $expected -ErrorAction Stop
            Write-Ok "$name start type set to $expected."
            Add-ConduitAction -Step "Re-enable $name" -Status 'Done' -Detail "Disabled -> $expected"
            Add-TKNote -Text "Re-enabled the $name service (Disabled -> $expected)." -Category 'Action'
        } catch {
            Write-Fail "Could not set $name start type: $($_.Exception.Message)"
            Write-TKError -ScriptName 'conduit' -Message "Service start type change failed for ${name}: $($_.Exception.Message)" -Category 'Service'
            Add-ConduitAction -Step "Re-enable $name" -Status 'Failed' -Detail $_.Exception.Message
        }
    }

    if (-not $changed) {
        Write-Ok "No update service is disabled. Nothing to do."
        Add-ConduitAction -Step 'Re-enable update services' -Status 'Skipped' -Detail 'None disabled'
    }
}

function Reset-ConduitUpdateCache {
    Write-Section "REPAIR — UPDATE CACHE RESET"

    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $targets = @(
        (Join-Path $env:SystemRoot 'SoftwareDistribution')
        (Join-Path $env:SystemRoot 'System32\catroot2')
    )

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would stop wuauserv, bits, cryptsvc, usosvc."
        foreach ($t in $targets) {
            if (Test-Path $t) { Write-Warn "[WhatIf] Would rename $t to $(Split-Path $t -Leaf).old_$stamp" }
        }
        Write-Warn "[WhatIf] Would restart cryptsvc, bits, wuauserv, usosvc."
        Add-ConduitAction -Step 'Reset update cache' -Status 'WhatIf' -Detail ($targets -join '; ')
        return
    }

    Write-Step "Stopping update services..."
    Stop-Service -Name 'wuauserv', 'bits', 'cryptsvc', 'usosvc' -Force -ErrorAction SilentlyContinue

    foreach ($t in $targets) {
        if (-not (Test-Path $t)) {
            Write-Info "$t not present — skipping."
            continue
        }
        $newName = "$(Split-Path $t -Leaf).old_$stamp"
        try {
            Rename-Item -Path $t -NewName $newName -ErrorAction Stop
            Write-Ok "Renamed $t to $newName."
            Add-ConduitAction -Step 'Reset update cache' -Status 'Done' -Detail "$t -> $newName"
            Add-TKNote -Text "Renamed $t to $newName so the update client rebuilds it." -Category 'Action'
        } catch {
            Write-Fail "Could not rename ${t}: $($_.Exception.Message)"
            Write-TKError -ScriptName 'conduit' -Message "Cache rename failed for ${t}: $($_.Exception.Message)" -Category 'Filesystem'
            Add-ConduitAction -Step 'Reset update cache' -Status 'Failed' -Detail $_.Exception.Message
        }
    }

    Write-Step "Restarting update services..."
    Start-Service -Name 'cryptsvc', 'bits', 'wuauserv', 'usosvc' -ErrorAction SilentlyContinue
    Write-Ok "Update services restarted."
}

function Start-ConduitScan {
    # Nudges the client to re-scan so the Settings page reflects the repair
    # rather than showing the cached failure until the next scheduled sweep.
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would trigger UsoClient StartInteractiveScan."
        return
    }
    try {
        Start-Process -FilePath 'UsoClient.exe' -ArgumentList 'StartInteractiveScan' -WindowStyle Hidden -ErrorAction Stop
        Write-Ok "Triggered a Windows Update scan."
    } catch {
        Write-Info "Could not trigger a scan automatically — open Settings > Windows Update and check manually."
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Get-ConduitVerdict {
    $errors   = @($Findings | Where-Object { $_.Severity -eq 'Error' })
    $warnings = @($Findings | Where-Object { $_.Severity -eq 'Warning' })

    if ($errors.Count -gt 0)   { return [PSCustomObject]@{ Verdict = 'Blocked';   Class = 'err'  } }
    if ($warnings.Count -gt 0) { return [PSCustomObject]@{ Verdict = 'Degraded';  Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

function Build-ConduitReport {
    param([object]$Audit, [object]$Verdict)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $errorCount = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount  = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count

    # Findings rows
    $fRows = New-Object System.Text.StringBuilder
    if ($Findings.Count -eq 0) {
        [void]$fRows.Append("<tr><td colspan='4'>No issues found.</td></tr>")
    } else {
        foreach ($f in $Findings) {
            $badge = Get-SeverityClass -Severity $f.Severity
            [void]$fRows.Append(
                "<tr><td><span class='tk-badge-$badge'>$(EscHtml $f.Severity)</span></td>" +
                "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
                "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
                $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
                "<td>$(EscHtml $f.Remedy)</td></tr>"
            )
        }
    }

    # Policy rows
    $policyPairs = [ordered]@{
        'WUServer'                                     = $(if ($Audit.Policy.WUServer)       { $Audit.Policy.WUServer }       else { '(not set)' })
        'WUStatusServer'                               = $(if ($Audit.Policy.WUStatusServer) { $Audit.Policy.WUStatusServer } else { '(not set)' })
        'UseWUServer'                                  = $(if ($null -ne $Audit.Policy.UseWUServer)    { "$($Audit.Policy.UseWUServer)" }    else { '(not set)' })
        'DoNotConnectToWindowsUpdateInternetLocations' = $(if ($null -ne $Audit.Policy.BlockInternet)  { "$($Audit.Policy.BlockInternet)" }  else { '(not set)' })
        'DisableWindowsUpdateAccess'                   = $(if ($null -ne $Audit.Policy.AccessDisabled) { "$($Audit.Policy.AccessDisabled)" } else { '(not set)' })
        'NoAutoUpdate'                                 = $(if ($null -ne $Audit.Policy.NoAutoUpdate)   { "$($Audit.Policy.NoAutoUpdate)" }   else { '(not set)' })
    }
    $pRows = New-Object System.Text.StringBuilder
    foreach ($k in $policyPairs.Keys) {
        [void]$pRows.Append("<tr><td class='tk-mono'>$(EscHtml $k)</td><td>$(EscHtml $policyPairs[$k])</td></tr>")
    }

    # Endpoint rows
    $eRows = New-Object System.Text.StringBuilder
    if ($Audit.Wsus.Configured) {
        $wsusBadge = if ($Audit.Wsus.Reachable) { 'ok' } else { 'err' }
        $wsusText  = if ($Audit.Wsus.Reachable) { 'Reachable' } else { 'Unreachable' }
        [void]$eRows.Append("<tr><td class='tk-mono'>$(EscHtml $Audit.Wsus.HostName)</td><td>$($Audit.Wsus.Port)</td><td>WSUS</td><td>$(EscHtml "$($Audit.Wsus.DnsOk)")</td><td><span class='tk-badge-$wsusBadge'>$wsusText</span></td></tr>")
    }
    foreach ($ep in $Audit.Endpoints) {
        $badge = if ($ep.Reachable) { 'ok' } else { 'err' }
        $text  = if ($ep.Reachable) { 'Reachable' } else { 'Unreachable' }
        [void]$eRows.Append("<tr><td class='tk-mono'>$(EscHtml $ep.Endpoint)</td><td>$($ep.Port)</td><td>Microsoft Update</td><td>-</td><td><span class='tk-badge-$badge'>$text</span></td></tr>")
    }

    # Service rows
    $sRows = New-Object System.Text.StringBuilder
    foreach ($svc in $Audit.Services) {
        $badge = if (-not $svc.Present) { 'err' } elseif ($svc.StartType -eq 'Disabled') { 'err' } else { 'ok' }
        [void]$sRows.Append(
            "<tr><td class='tk-mono'>$(EscHtml $svc.Name)</td><td>$(EscHtml $svc.Status)</td>" +
            "<td><span class='tk-badge-$badge'>$(EscHtml $svc.StartType)</span></td><td>$(EscHtml $svc.Expected)</td></tr>"
        )
    }

    # Action rows
    $aRows = New-Object System.Text.StringBuilder
    if ($Actions.Count -eq 0) {
        [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>")
    } else {
        foreach ($a in $Actions) {
            $badge = switch ($a.Status) {
                'Done'   { 'ok'   }
                'WhatIf' { 'blue' }
                'Failed' { 'err'  }
                default  { 'info' }
            }
            [void]$aRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td>" +
                "<td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>"
            )
        }
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.O.N.D.U.I.T. Windows Update Connectivity Report' `
        -ScriptName 'C.O.N.D.U.I.T.' `
        -Subtitle   "${orgPrefix}Windows Update Connectivity -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'      = $machine
            'Generated'    = $reportDate
            'Mode'         = $(if ($WhatIf) { "$Action (dry run)" } else { $Action })
            'Verdict'      = $Verdict.Verdict
            'Domain'       = $Audit.Device.Domain
            'WSUS'         = $(if ($Audit.Policy.WsusConfigured) { $Audit.Policy.WUServer } else { 'Not configured' })
        }) `
        -NavItems   @('Findings', 'Device & Policy', 'Connectivity', 'Update Services', 'Actions Taken')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'C.O.N.D.U.I.T. v5.1'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Update Connectivity</div></div>
    <div class="tk-summary-card $(if ($errorCount -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errorCount</div><div class="tk-summary-lbl">Blocking Issues</div></div>
    <div class="tk-summary-card $(if ($warnCount -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(if ($Audit.Policy.WsusConfigured) { 'Yes' } else { 'No' })</div><div class="tk-summary-lbl">WSUS Configured</div></div>
    <div class="tk-summary-card $(if ($Audit.Proxy.Configured) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$(if ($Audit.Proxy.Configured) { 'Set' } else { 'Direct' })</div><div class="tk-summary-lbl">WinHTTP Proxy</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Actions.Count)</div><div class="tk-summary-lbl">Actions Recorded</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
        <tbody>$($fRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Device &amp; Policy</div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Computer</span> $(EscHtml $Audit.Device.Computer)<br/>
        <span class="tk-info-label">Operating system</span> $(EscHtml $Audit.Device.OS)<br/>
        <span class="tk-info-label">Domain</span> $(EscHtml $Audit.Device.Domain) (joined: $(EscHtml "$($Audit.Device.PartOfDomain)"))<br/>
        <span class="tk-info-label">Local GPO sets WUServer</span> $(EscHtml "$($Audit.Source.LocalGpoSetsWsus)")<br/>
        <span class="tk-info-label">Windows Time</span> $(EscHtml $Audit.Time.Status) (start type $(EscHtml $Audit.Time.StartType))<br/>
        <span class="tk-info-label">WinHTTP proxy</span> $(EscHtml $Audit.Proxy.Raw)
      </div>
      <table class="tk-table">
        <thead><tr><th>Policy value</th><th>Setting</th></tr></thead>
        <tbody>$($pRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Connectivity</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Endpoint</th><th>Port</th><th>Role</th><th>DNS</th><th>Result</th></tr></thead>
        <tbody>$($eRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Update Services</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Service</th><th>Status</th><th>Start type</th><th>Windows default</th></tr></thead>
        <tbody>$($sRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Actions Taken</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
        <tbody>$($aRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ConduitRepair {
    param([object]$Audit)

    Backup-ConduitPolicy | Out-Null
    Repair-ConduitWsusPointer -Audit $Audit
    Repair-ConduitTimeService
    Repair-ConduitServices
}

function Invoke-ConduitRun {
    param([string]$Mode)

    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-ConduitAudit
    Show-ConduitFindings

    if ($Mode -ne 'Audit') {
        Invoke-ConduitRepair -Audit $audit
        if ($Mode -eq 'ResetCache') { Reset-ConduitUpdateCache }
        Start-ConduitScan

        Write-Section "VERIFY"
        # Re-read the live state rather than trusting the repair's own return
        # values -- a policy that Group Policy reapplies will already be back.
        $post   = Get-ItemProperty -Path $WUPolicyKey -ErrorAction SilentlyContinue
        $postAu = Get-ItemProperty -Path $AUPolicyKey -ErrorAction SilentlyContinue
        $w32    = Get-ServiceSafe -Name 'w32time'
        $wua    = Get-ServiceSafe -Name 'wuauserv'

        Write-Info ("WUServer    : {0}" -f $(if ($post.WUServer) { $post.WUServer } else { '(not set)' }))
        Write-Info ("UseWUServer : {0}" -f $(if ($null -ne $postAu.UseWUServer) { $postAu.UseWUServer } else { '(not set)' }))
        Write-Info ("w32time     : {0}" -f $(if ($w32) { $w32.Status } else { 'absent' }))
        Write-Info ("wuauserv    : {0}" -f $(if ($wua) { $wua.Status } else { 'absent' }))

        $remaining = @()
        if ($post.WUServer -and $postAu.UseWUServer -eq 1 -and $audit.Wsus.Configured -and -not $audit.Wsus.Reachable) {
            $remaining += 'An unreachable WSUS pointer is still set'
        }
        if ($w32 -and $w32.Status -ne 'Running') { $remaining += 'The Windows Time service is not running' }
        if ($audit.Proxy.Configured)             { $remaining += 'A WinHTTP proxy is set — review it manually' }
        if (@($audit.Endpoints | Where-Object { -not $_.Reachable }).Count -gt 0) {
            $remaining += 'Microsoft Update endpoints are unreachable — this is an upstream network block'
        }

        Write-Host ""
        if ($remaining.Count -eq 0) {
            Write-Ok "Repair complete. Check Settings > Windows Update; reboot if the error persists."
            Add-TKNote -Text 'CONDUIT repair completed with no remaining blockers.' -Category 'Resolution' -ScriptName 'conduit'
        } else {
            foreach ($r in $remaining) { Write-Warn $r }
            Add-TKNote -Text ("CONDUIT repair left {0} item(s) outstanding: {1}" -f $remaining.Count, ($remaining -join '; ')) -Category 'Issue' -ScriptName 'conduit'
        }
    }

    $verdict = Get-ConduitVerdict
    Write-Section "VERDICT"
    switch ($verdict.Class) {
        'err'   { Write-Fail $verdict.Verdict }
        'warn'  { Write-Warn $verdict.Verdict }
        default { Write-Ok   $verdict.Verdict }
    }

    Add-TKNote -Text ("CONDUIT {0} run on {1}: verdict {2} ({3} finding(s))." -f $Mode, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'conduit'

    Write-Step "Generating HTML report..."
    $html      = Build-ConduitReport -Audit $audit -Verdict $verdict
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "CONDUIT_${timestamp}.html"

    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-ConduitBanner
    Invoke-ConduitRun -Mode $Action
} else {
    $choice = ''

    do {
        Show-ConduitBanner

        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only diagnosis + HTML report" -ForegroundColor $C.Info
        Write-Host "  [2] Repair  -  audit, then apply the safe repairs" -ForegroundColor $C.Info
        Write-Host "  [3] Repair + reset update cache  -  also rebuilds SoftwareDistribution / catroot2" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($WhatIf) { Write-Host "  Dry run is active — options 2 and 3 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-ConduitRun -Mode 'Audit' }
            '2' { Invoke-ConduitRun -Mode 'Repair' }
            '3' { Invoke-ConduitRun -Mode 'ResetCache' }
            'Q' {
                Write-Host ""
                Write-Host "  Closing C.O.N.D.U.I.T." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-3 or Q." -ForegroundColor $C.Warning
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
