# oath.ps1 - O.A.T.H. — Observes And Tends the Host's domain trust
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
    O.A.T.H. — Observes And Tends the Host's domain trust
    Domain Trust & Secure Channel Diagnostics and Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Diagnoses "The trust relationship between this workstation and the
    primary domain failed" and the quieter failures that precede it, on a
    domain-joined Windows machine:

      - Domain controller discovery (nltest /dsgetdc) and the DC locator SRV
        record in DNS
      - DNS servers that are public resolvers rather than domain DNS -- the
        most common cause of a member that cannot find its domain
      - Reachability of the DC on DNS, Kerberos, RPC, LDAP and SMB
      - Clock offset from the DC: Kerberos refuses tickets past five minutes
      - The secure channel itself (nltest /sc_verify), with the Netlogon
        status decoded into connectivity versus a broken machine password
      - Netlogon policy that stops the machine password from rotating

    Audit is read-only. Repair resynchronises the clock from the domain
    hierarchy, resets the secure channel to a DC (nltest /sc_reset), and --
    only when the machine password itself is rejected, and only
    interactively -- resets the computer account password with domain
    credentials the technician enters. It never changes DNS settings and never
    unjoins or rejoins the domain; those are reported with the remedy.
    -WhatIf previews every repair.

.USAGE
    PS C:\> .\oath.ps1                               # Interactive menu
    PS C:\> .\oath.ps1 -Unattended                   # Read-only audit + HTML report
    PS C:\> .\oath.ps1 -Unattended -Action Repair    # Audit, then the repairs that need no credentials
    PS C:\> .\oath.ps1 -Action Repair -WhatIf        # Preview the repairs only

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit', 'Repair')]
    [string]$Action = 'Audit',
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
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Netlogon status codes nltest reports, split by what fixes them. Connectivity
# failures are fixed on the network (DNS, firewall, the DC); Trust failures
# mean the machine account and its password no longer agree with the domain.
$NetlogonStatusCodes = @{
    0    = @{ Name = 'NERR_Success';                      Kind = 'Ok';           Meaning = 'The secure channel is healthy.' }
    5    = @{ Name = 'ERROR_ACCESS_DENIED';               Kind = 'Trust';        Meaning = 'The DC rejected the machine''s credentials, or the check ran without elevation.' }
    53   = @{ Name = 'ERROR_BAD_NETPATH';                 Kind = 'Connectivity'; Meaning = 'The DC''s network path could not be reached.' }
    1311 = @{ Name = 'ERROR_NO_LOGON_SERVERS';            Kind = 'Connectivity'; Meaning = 'No domain controller could be contacted.' }
    1355 = @{ Name = 'ERROR_NO_SUCH_DOMAIN';              Kind = 'Connectivity'; Meaning = 'The domain could not be found -- usually DNS.' }
    1722 = @{ Name = 'RPC_S_SERVER_UNAVAILABLE';          Kind = 'Connectivity'; Meaning = 'RPC to the DC failed -- firewall or the DC is down.' }
    1786 = @{ Name = 'ERROR_NO_TRUST_LSA_SECRET';         Kind = 'Trust';        Meaning = 'The local copy of the machine password is missing.' }
    1787 = @{ Name = 'ERROR_NO_TRUST_SAM_ACCOUNT';        Kind = 'Trust';        Meaning = 'The computer account is missing from the domain -- deleted, or reset by someone else.' }
    1789 = @{ Name = 'ERROR_TRUSTED_RELATIONSHIP_FAILURE'; Kind = 'Trust';       Meaning = 'The machine password does not match the domain''s copy.' }
    1790 = @{ Name = 'ERROR_TRUSTED_DOMAIN_FAILURE';      Kind = 'Trust';        Meaning = 'The trust between the domains failed.' }
}

# TCP ports a domain member needs to reach on its DC.
$DcPorts = [ordered]@{
    53  = 'DNS'
    88  = 'Kerberos'
    135 = 'RPC endpoint mapper'
    389 = 'LDAP'
    445 = 'SMB (SYSVOL / Netlogon)'
}

# Kerberos rejects tickets beyond five minutes of skew; drift past a minute is
# worth fixing before it gets there.
$ClockSkewErrorSeconds   = 300
$ClockSkewWarningSeconds = 60

$NetlogonParamsKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$OathFindings = @{
    'NotDomainJoined' = @{
        Severity = 'Info'
        Title    = 'Not joined to an Active Directory domain'
        Summary  = 'This machine is in a workgroup (or Entra-joined only), so there is no domain trust to check.'
        Remedy   = 'Nothing to do. For Entra-joined machines, check dsregcmd /status instead.'
    }
    'DomainNotFound' = @{
        Severity = 'Error'
        Title    = 'No domain controller could be located'
        Summary  = 'DC discovery (nltest /dsgetdc) failed, so nothing domain-related can work -- logons fall back to cached credentials.'
        Remedy   = 'Fix DNS first: the machine must use the domain''s DNS servers. Then check the network path to a DC.'
    }
    'DcSrvMissing' = @{
        Severity = 'Error'
        Title    = 'DC locator SRV record does not resolve'
        Summary  = '_ldap._tcp.dc._msdcs.<domain> did not resolve through this machine''s DNS servers, which is how Windows finds a domain controller.'
        Remedy   = 'Point the adapter at the domain DNS servers (usually the DCs), not the router or a public resolver.'
    }
    'PublicDnsServer' = @{
        Severity = 'Warning'
        Title    = 'Public DNS resolver configured on a domain member'
        Summary  = 'A public resolver cannot answer for the internal domain. Windows uses it whenever the first server is slow, so domain lookups fail intermittently.'
        Remedy   = 'Use only the domain DNS servers on domain members; let those forward to public resolvers. O.A.T.H. never changes DNS itself.'
    }
    'DcPortBlocked' = @{
        Severity = 'Error'
        Title    = 'Domain controller port unreachable'
        Summary  = 'A port the domain member needs on its DC did not answer -- a firewall, VPN split-tunnel or the DC itself.'
        Remedy   = 'Open the path to the DC for the ports listed. No local repair fixes an upstream block.'
    }
    'ClockSkew' = @{
        Severity = 'Error'
        Title    = 'Clock is more than five minutes off the DC'
        Summary  = 'Kerberos rejects tickets beyond five minutes of skew, so domain logons and access fail even with a healthy secure channel.'
        Remedy   = 'Repair resynchronises from the domain hierarchy (w32tm /resync).'
    }
    'ClockDrift' = @{
        Severity = 'Warning'
        Title    = 'Clock is drifting from the DC'
        Summary  = 'The offset is over a minute: not failing yet, but heading toward the five-minute Kerberos limit.'
        Remedy   = 'Repair resynchronises from the domain hierarchy. Check the Windows Time service if it keeps drifting.'
    }
    'SecureChannelBroken' = @{
        Severity = 'Error'
        Title    = 'Secure channel down -- the DC could not be reached'
        Summary  = 'Netlogon could not reach a DC to verify the channel. This is a connectivity fault, not a broken trust.'
        Remedy   = 'Fix the DNS and connectivity findings first. Repair then resets the secure channel to a reachable DC.'
    }
    'TrustBroken' = @{
        Severity = 'Error'
        Title    = 'Trust relationship failed'
        Summary  = 'The DC rejected the machine''s credentials: the machine password and the domain''s copy disagree, or the computer account is gone.'
        Remedy   = 'Repair (interactive) resets the machine password with domain credentials. If the computer account was deleted, the machine must be rejoined.'
    }
    'ComputerAccountMissing' = @{
        Severity = 'Error'
        Title    = 'Computer account missing from the domain'
        Summary  = 'The DC reports no account for this machine (ERROR_NO_TRUST_SAM_ACCOUNT). A password reset cannot fix this.'
        Remedy   = 'Rejoin the domain: sign in with a local admin, Remove-Computer to a workgroup, restart, then Add-Computer. O.A.T.H. never rejoins automatically.'
    }
    'SecureChannelUnknown' = @{
        Severity = 'Warning'
        Title    = 'Secure channel state could not be determined'
        Summary  = 'nltest returned output O.A.T.H. could not read, or was not available.'
        Remedy   = 'Run elevated, and check nltest /sc_verify:<domain> by hand.'
    }
    'PasswordChangeDisabled' = @{
        Severity = 'Warning'
        Title    = 'Machine password rotation disabled'
        Summary  = 'Netlogon DisablePasswordChange = 1. Usually set to "fix" snapshot restores, it leaves a long-lived machine credential instead.'
        Remedy   = 'Clear the setting at its source (GPO: Domain member: Disable machine account password changes) unless a documented reason exists.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[object]]::new()

function Add-OathFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $OathFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
    })
}

function Add-OathAction {
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

function Show-OathBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗  █████╗ ████████╗██╗  ██╗
  ██╔═══██╗██╔══██╗╚══██╔══╝██║  ██║
  ██║   ██║███████║   ██║   ███████║
  ██║   ██║██╔══██║   ██║   ██╔══██║
  ╚██████╔╝██║  ██║   ██║   ██║  ██║
   ╚═════╝ ╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    O.A.T.H. — Observes And Tends the Host's domain trust" -ForegroundColor Cyan
    Write-Host "    Domain Trust & Secure Channel Diagnostics and Repair Tool" -ForegroundColor Cyan
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

function ConvertFrom-NltestDsGetDc {
    param([string[]]$Lines)
    $text = $Lines -join "`n"
    $status = [regex]::Match($text, '(?i)Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')
    $dc     = [regex]::Match($text, '(?im)^\s*DC:\s*\\\\(\S+)')
    $addr   = [regex]::Match($text, '(?im)^\s*Address:\s*\\\\(\S+)')
    $site   = [regex]::Match($text, '(?im)^\s*Dc Site Name:\s*(.+?)\s*$')
    $ours   = [regex]::Match($text, '(?im)^\s*Our Site Name:\s*(.+?)\s*$')
    return [PSCustomObject]@{
        Success    = $dc.Success
        Dc         = if ($dc.Success) { $dc.Groups[1].Value } else { '' }
        Address    = if ($addr.Success) { $addr.Groups[1].Value } else { '' }
        DcSite     = if ($site.Success) { $site.Groups[1].Value } else { '' }
        OurSite    = if ($ours.Success) { $ours.Groups[1].Value } else { '' }
        StatusCode = if ($status.Success) { [int]$status.Groups[1].Value } elseif ($dc.Success) { 0 } else { $null }
        StatusName = if ($status.Success) { $status.Groups[2].Value } else { '' }
    }
}

function ConvertFrom-NltestSecureChannel {
    # Reads nltest /sc_query or /sc_verify. The overall code is the verify
    # status when present (it is the stronger check), else the connection
    # status, else a bare "failed: Status = n" line.
    param([string[]]$Lines)
    $text = $Lines -join "`n"
    $dc      = [regex]::Match($text, '(?i)Trusted DC Name\s+\\\\(\S+)')
    $conn    = [regex]::Match($text, '(?i)Trusted DC Connection Status\s+Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')
    $verify  = [regex]::Match($text, '(?i)Trust Verification Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')
    $failed  = [regex]::Match($text, '(?i)failed:\s*Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')

    $code = $null; $name = ''
    if ($verify.Success -and [int]$verify.Groups[1].Value -ne 0) { $code = [int]$verify.Groups[1].Value; $name = $verify.Groups[2].Value }
    elseif ($conn.Success -and [int]$conn.Groups[1].Value -ne 0) { $code = [int]$conn.Groups[1].Value;   $name = $conn.Groups[2].Value }
    elseif ($failed.Success)                                      { $code = [int]$failed.Groups[1].Value; $name = $failed.Groups[2].Value }
    elseif ($verify.Success -or $conn.Success)                    { $code = 0; $name = 'NERR_Success' }

    return [PSCustomObject]@{
        TrustedDc  = if ($dc.Success) { $dc.Groups[1].Value } else { '' }
        Verified   = $verify.Success
        StatusCode = $code
        StatusName = $name
    }
}

function Get-NetlogonStatusInfo {
    param($Code)
    if ($null -eq $Code) { return [PSCustomObject]@{ Name = 'Unknown'; Kind = 'Unknown'; Meaning = 'No status could be read.' } }
    $entry = $NetlogonStatusCodes[[int]$Code]
    if ($entry) { return [PSCustomObject]@{ Name = $entry.Name; Kind = $entry.Kind; Meaning = $entry.Meaning } }
    return [PSCustomObject]@{ Name = "Status $Code"; Kind = 'Unknown'; Meaning = 'Not a status O.A.T.H. recognises -- look it up with: net helpmsg ' + $Code }
}

function ConvertFrom-W32tmStripchart {
    # w32tm /stripchart /dataonly prints "hh:mm:ss, +00.0123456s" per sample.
    # Returns the offset in seconds of the first sample, or $null.
    param([string[]]$Lines)
    foreach ($l in $Lines) {
        $m = [regex]::Match("$l", '^\s*\d{1,2}:\d{2}:\d{2},\s*([+-]?\d+(?:\.\d+)?)s\s*$')
        if ($m.Success) { return [double]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) }
    }
    return $null
}

function Test-PublicIpAddress {
    # True for an address that is not private, loopback, link-local, CGNAT or
    # IPv6 unique-local -- i.e. a resolver on the internet.
    param([string]$Address)
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse("$Address", [ref]$ip)) { return $false }
    if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
        if ($ip.IsIPv6LinkLocal -or $ip.IsIPv6SiteLocal -or [System.Net.IPAddress]::IsLoopback($ip)) { return $false }
        $first = $ip.GetAddressBytes()[0]
        return -not (($first -band 0xFE) -eq 0xFC)
    }
    $b = $ip.GetAddressBytes()
    if ($b[0] -eq 10 -or $b[0] -eq 127 -or $b[0] -eq 0) { return $false }
    if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $false }
    if ($b[0] -eq 192 -and $b[1] -eq 168) { return $false }
    if ($b[0] -eq 169 -and $b[1] -eq 254) { return $false }
    if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $false }
    return $true
}

function Get-OathVerdict {
    param([object[]]$FindingList)
    $codes = @($FindingList | ForEach-Object { $_.Code })
    $sev   = @($FindingList | ForEach-Object { $_.Severity })
    if ($codes -contains 'NotDomainJoined') { return [PSCustomObject]@{ Verdict = 'Not joined'; Class = 'info' } }
    if ($sev -contains 'Error')             { return [PSCustomObject]@{ Verdict = 'Broken';     Class = 'err'  } }
    if ($sev -contains 'Warning')           { return [PSCustomObject]@{ Verdict = 'Degraded';   Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Native {
    # Captures a native tool's output without letting a missing binary or a
    # non-zero exit abort the audit.
    param([string]$File, [string[]]$Arguments)
    try {
        $out = & $File @Arguments 2>&1 | ForEach-Object { "$_" }
        return @($out)
    } catch {
        return @("failed: $($_.Exception.Message)")
    }
}

function Test-TcpPort {
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 2000)
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

function Get-OathContext {
    $cs = $null; $os = $null
    try { $cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop } catch { $cs = $null }
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }
    return [PSCustomObject]@{
        Computer     = $env:COMPUTERNAME
        PartOfDomain = if ($cs) { [bool]$cs.PartOfDomain } else { $false }
        Domain       = if ($cs) { "$($cs.Domain)" } else { '' }
        OS           = if ($os) { "$($os.Caption) (build $($os.BuildNumber))" } else { 'Unknown' }
        Edition      = "$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)"
    }
}

function Get-OathDnsServers {
    $rows = @()
    try {
        $up = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' } | ForEach-Object { $_.ifIndex })
        $rows = @(Get-DnsClientServerAddress -ErrorAction Stop | Where-Object { $up -contains $_.InterfaceIndex -and $_.ServerAddresses } | ForEach-Object {
            $alias = $_.InterfaceAlias
            foreach ($s in $_.ServerAddresses) { [PSCustomObject]@{ Interface = $alias; Server = "$s"; Public = (Test-PublicIpAddress -Address "$s") } }
        })
    } catch {
        $rows = @()
    }
    return $rows
}

function Invoke-OathAudit {
    Write-Section "DOMAIN MEMBERSHIP"
    $ctx = Get-OathContext
    Write-Info "$($ctx.Computer)  |  $($ctx.OS)"
    if (-not $ctx.PartOfDomain) {
        Write-Info "Not joined to an Active Directory domain."
        Add-OathFinding -Code 'NotDomainJoined' -Detail "Workgroup / domain field: $($ctx.Domain)"
        return [PSCustomObject]@{
            Context = $ctx; Joined = $false; Domain = ''; Dns = @(); SrvOk = $false; Ports = @(); Offset = $null; MaxPwdAge = $null
            DsGetDc = [PSCustomObject]@{ Success = $false; Dc = ''; Address = ''; DcSite = ''; OurSite = ''; StatusCode = $null; StatusName = '' }
            Channel = [PSCustomObject]@{ TrustedDc = ''; Verified = $false; StatusCode = $null; StatusName = '' }
            Status  = [PSCustomObject]@{ Name = 'n/a'; Kind = 'NotJoined'; Meaning = 'Not joined to a domain.' }
        }
    }
    $domain = $ctx.Domain
    Write-Info "Domain: $domain"

    Write-Section "DNS"
    $dns = Get-OathDnsServers
    foreach ($d in $dns) { Write-Info ("{0,-24} {1}{2}" -f $d.Interface, $d.Server, $(if ($d.Public) { '  (public resolver)' } else { '' })) }
    $public = @($dns | Where-Object { $_.Public })
    if ($public.Count -gt 0) { Add-OathFinding -Code 'PublicDnsServer' -Detail (($public | ForEach-Object { "$($_.Server) on $($_.Interface)" }) -join ', ') }

    $srvName = "_ldap._tcp.dc._msdcs.$domain"
    $srvOk = $false
    try { $srvOk = [bool](@(Resolve-DnsName -Name $srvName -Type SRV -DnsOnly -ErrorAction Stop | Where-Object { $_.Type -eq 'SRV' }).Count) } catch { $srvOk = $false }
    if (-not $srvOk) { Add-OathFinding -Code 'DcSrvMissing' -Detail $srvName }
    Write-Info ("DC locator SRV: {0}" -f $(if ($srvOk) { 'resolves' } else { 'DOES NOT RESOLVE' }))

    Write-Section "DOMAIN CONTROLLER"
    $dsget = ConvertFrom-NltestDsGetDc -Lines (Invoke-Native -File 'nltest.exe' -Arguments @("/dsgetdc:$domain"))
    if (-not $dsget.Success) {
        Add-OathFinding -Code 'DomainNotFound' -Detail ("nltest /dsgetdc: {0} {1}" -f $dsget.StatusCode, $dsget.StatusName)
        Write-Fail "No DC located ($($dsget.StatusName))."
    } else {
        Write-Info "DC: $($dsget.Dc)  ($($dsget.Address))  site $($dsget.DcSite)"
    }

    $ports = @()
    $offset = $null
    if ($dsget.Success) {
        $ports = foreach ($p in $DcPorts.Keys) {
            [PSCustomObject]@{ Port = [int]$p; Service = $DcPorts[$p]; Open = (Test-TcpPort -Hostname $dsget.Dc -Port ([int]$p)) }
        }
        $ports = @($ports)
        $closed = @($ports | Where-Object { -not $_.Open })
        if ($closed.Count -gt 0) { Add-OathFinding -Code 'DcPortBlocked' -Detail (($closed | ForEach-Object { "$($_.Port) $($_.Service)" }) -join ', ') }
        foreach ($p in $ports) { Write-Info ("{0,-5} {1,-26} {2}" -f $p.Port, $p.Service, $(if ($p.Open) { 'open' } else { 'BLOCKED' })) }

        Write-Section "TIME"
        $offset = ConvertFrom-W32tmStripchart -Lines (Invoke-Native -File 'w32tm.exe' -Arguments @('/stripchart', "/computer:$($dsget.Dc)", '/samples:1', '/dataonly'))
        if ($null -eq $offset) {
            Write-Warn "Could not measure the offset from $($dsget.Dc)."
        } else {
            Write-Info ("Offset from DC: {0:N2}s" -f $offset)
            if ([math]::Abs($offset) -gt $ClockSkewErrorSeconds)       { Add-OathFinding -Code 'ClockSkew'  -Detail ("{0:N0} seconds" -f $offset) }
            elseif ([math]::Abs($offset) -gt $ClockSkewWarningSeconds) { Add-OathFinding -Code 'ClockDrift' -Detail ("{0:N0} seconds" -f $offset) }
        }
    }

    Write-Section "SECURE CHANNEL"
    $channel = Get-OathSecureChannel -Domain $domain
    $status  = Get-NetlogonStatusInfo -Code $channel.StatusCode
    Write-Info ("Status: {0} ({1}) via {2}" -f $status.Name, $status.Meaning, $(if ($channel.TrustedDc) { $channel.TrustedDc } else { 'no DC' }))
    Add-OathChannelFinding -Channel $channel -Status $status

    $np = Get-ItemProperty -Path $NetlogonParamsKey -ErrorAction SilentlyContinue
    if ($np -and $np.DisablePasswordChange -eq 1) { Add-OathFinding -Code 'PasswordChangeDisabled' -Detail "$NetlogonParamsKey\DisablePasswordChange = 1" }

    return [PSCustomObject]@{
        Context   = $ctx
        Joined    = $true
        Domain    = $domain
        Dns       = @($dns)
        SrvOk     = $srvOk
        DsGetDc   = $dsget
        Ports     = @($ports)
        Offset    = $offset
        Channel   = $channel
        Status    = $status
        MaxPwdAge = if ($np -and $np.MaximumPasswordAge) { [int]$np.MaximumPasswordAge } else { 30 }
    }
}

function Get-OathSecureChannel {
    # /sc_verify checks the machine password with the DC; it needs elevation.
    # /sc_query only reports the last known state, so it is the fallback.
    param([string]$Domain)
    $channel = ConvertFrom-NltestSecureChannel -Lines (Invoke-Native -File 'nltest.exe' -Arguments @("/sc_verify:$Domain"))
    if ($null -eq $channel.StatusCode) {
        $channel = ConvertFrom-NltestSecureChannel -Lines (Invoke-Native -File 'nltest.exe' -Arguments @("/sc_query:$Domain"))
    }
    return $channel
}

function Add-OathChannelFinding {
    param([object]$Channel, [object]$Status)
    switch ($Status.Kind) {
        'Ok'           { }
        'Connectivity' { Add-OathFinding -Code 'SecureChannelBroken' -Detail "$($Status.Name): $($Status.Meaning)" }
        'Trust' {
            if ($Channel.StatusCode -eq 1787) { Add-OathFinding -Code 'ComputerAccountMissing' -Detail "$($Status.Name): $($Status.Meaning)" }
            else                              { Add-OathFinding -Code 'TrustBroken' -Detail "$($Status.Name): $($Status.Meaning)" }
        }
        default        { Add-OathFinding -Code 'SecureChannelUnknown' -Detail "$($Status.Name): $($Status.Meaning)" }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Repair-OathTime {
    Write-Section "REPAIR — TIME"
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would set w32time to sync from the domain hierarchy and force a resync."
        Add-OathAction -Step 'Resync time from the domain' -Status 'WhatIf'
        return
    }
    $cfg = Invoke-Native -File 'w32tm.exe' -Arguments @('/config', '/syncfromflags:domhier', '/update')
    $rs  = Invoke-Native -File 'w32tm.exe' -Arguments @('/resync', '/force')
    $ok  = ($rs -join ' ') -match '(?i)completed successfully'
    if ($ok) { Write-Ok "Clock resynchronised from the domain hierarchy." } else { Write-Warn "w32tm /resync: $($rs -join ' ')" }
    Add-OathAction -Step 'Resync time from the domain' -Status $(if ($ok) { 'Done' } else { 'Failed' }) -Detail (@($cfg + $rs) -join ' ')
    if ($ok) { Add-TKNote -Text 'Resynchronised the clock from the domain hierarchy.' -Category 'Action' -ScriptName 'oath' }
}

function Repair-OathSecureChannel {
    param([string]$Domain)
    Write-Section "REPAIR — SECURE CHANNEL"
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would reset the secure channel: nltest /sc_reset:$Domain"
        Add-OathAction -Step 'Reset secure channel' -Status 'WhatIf' -Detail "nltest /sc_reset:$Domain"
        return
    }
    $out    = Invoke-Native -File 'nltest.exe' -Arguments @("/sc_reset:$Domain")
    $parsed = ConvertFrom-NltestSecureChannel -Lines $out
    $ok     = $parsed.StatusCode -eq 0
    if ($ok) { Write-Ok "Secure channel reset to $($parsed.TrustedDc)." } else { Write-Warn "nltest /sc_reset: $($parsed.StatusName)" }
    Add-OathAction -Step 'Reset secure channel' -Status $(if ($ok) { 'Done' } else { 'Failed' }) -Detail ($out -join ' ')
    if ($ok) { Add-TKNote -Text "Reset the secure channel to $($parsed.TrustedDc)." -Category 'Action' -ScriptName 'oath' }
}

function Repair-OathMachinePassword {
    # Only for a rejected machine password. Needs domain credentials, so it is
    # interactive only. Reset-ComputerMachinePassword exists in Windows
    # PowerShell but not PowerShell 7, so under 7 it runs in powershell.exe.
    param([string]$Dc)
    Write-Section "REPAIR — MACHINE PASSWORD"
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would reset the computer account password against $Dc with domain credentials."
        Add-OathAction -Step 'Reset machine password' -Status 'WhatIf' -Detail $Dc
        return
    }
    if ($Unattended) {
        Write-Warn "Resetting the machine password needs domain credentials -- skipped in unattended mode."
        Add-OathAction -Step 'Reset machine password' -Status 'Skipped' -Detail 'Needs domain credentials; run interactively.'
        return
    }
    $ans = Read-Host "  Reset the computer account password against $Dc now? Needs a domain account with rights to reset it. [Y/N]"
    if ($ans -notmatch '^[Yy]') {
        Add-OathAction -Step 'Reset machine password' -Status 'Skipped' -Detail 'Declined by technician.'
        return
    }
    try {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            Reset-ComputerMachinePassword -Server $Dc -Credential (Get-Credential -Message 'Domain account to reset this computer''s password') -ErrorAction Stop
        } else {
            $command = "Reset-ComputerMachinePassword -Server '$Dc' -Credential (Get-Credential -Message 'Domain account to reset this computer''s password') -ErrorAction Stop"
            & powershell.exe -NoProfile -Command $command
            if ($LASTEXITCODE -ne 0) { throw "powershell.exe exited $LASTEXITCODE" }
        }
        Write-Ok "Machine password reset against $Dc."
        Add-OathAction -Step 'Reset machine password' -Status 'Done' -Detail $Dc
        Add-TKNote -Text "Reset the computer account password against $Dc." -Category 'Action' -ScriptName 'oath'
    } catch {
        Write-Fail "Machine password reset failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'oath' -Message "Machine password reset failed: $($_.Exception.Message)" -Category 'Domain Trust'
        Add-OathAction -Step 'Reset machine password' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Invoke-OathRepair {
    param([object]$Audit)
    $codes = @($Findings | ForEach-Object { $_.Code })

    if ($codes -contains 'ClockSkew' -or $codes -contains 'ClockDrift') { Repair-OathTime }

    if ($codes -contains 'SecureChannelBroken' -or $codes -contains 'TrustBroken' -or $codes -contains 'SecureChannelUnknown') {
        Repair-OathSecureChannel -Domain $Audit.Domain
        if ($codes -contains 'TrustBroken' -and -not $WhatIf) {
            # Re-verify: a stale DC choice can masquerade as a trust failure,
            # and sc_reset alone fixes that without touching the password.
            $again = Get-OathSecureChannel -Domain $Audit.Domain
            if ((Get-NetlogonStatusInfo -Code $again.StatusCode).Kind -eq 'Trust') {
                $dc = if ($Audit.DsGetDc.Dc) { $Audit.DsGetDc.Dc } else { $again.TrustedDc }
                Repair-OathMachinePassword -Dc $dc
            }
        } elseif ($codes -contains 'TrustBroken') {
            Repair-OathMachinePassword -Dc $Audit.DsGetDc.Dc
        }
    }

    if ($codes -contains 'ComputerAccountMissing') {
        Write-Warn "The computer account is missing from the domain. Rejoin the machine -- O.A.T.H. does not unjoin or rejoin."
        Add-OathAction -Step 'Rejoin domain' -Status 'Skipped' -Detail 'Computer account missing; rejoin manually.'
    }
    if ($codes -contains 'PublicDnsServer' -or $codes -contains 'DcSrvMissing') {
        Add-OathAction -Step 'Correct DNS servers' -Status 'Skipped' -Detail 'O.A.T.H. never changes DNS -- point the adapter at the domain DNS servers.'
    }
    if ($Actions.Count -eq 0) {
        Write-Ok "Nothing O.A.T.H. can repair automatically."
        Add-OathAction -Step 'Repair' -Status 'Skipped' -Detail 'No repairable finding.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-OathReport {
    param([object]$Audit, [object]$Verdict, [object]$After)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>No issues found.</td></tr>") }
    foreach ($f in $Findings) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
            $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $dRows = [System.Text.StringBuilder]::new()
    foreach ($d in $Audit.Dns) {
        $badge = if ($d.Public) { "<span class='tk-badge-warn'>Public</span>" } else { "<span class='tk-badge-ok'>Private</span>" }
        [void]$dRows.Append("<tr><td>$(EscHtml $d.Interface)</td><td class='tk-mono'>$(EscHtml $d.Server)</td><td>$badge</td></tr>")
    }
    if ($Audit.Dns.Count -eq 0) { [void]$dRows.Append("<tr><td colspan='3'>No DNS servers read.</td></tr>") }

    $pRows = [System.Text.StringBuilder]::new()
    foreach ($p in $Audit.Ports) {
        $badge = if ($p.Open) { "<span class='tk-badge-ok'>Open</span>" } else { "<span class='tk-badge-err'>Blocked</span>" }
        [void]$pRows.Append("<tr><td>$($p.Port)</td><td>$(EscHtml $p.Service)</td><td>$badge</td></tr>")
    }
    if ($Audit.Ports.Count -eq 0) { [void]$pRows.Append("<tr><td colspan='3'>No DC to test.</td></tr>") }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Actions.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>") }
    foreach ($a in $Actions) {
        $badge = switch ($a.Status) { 'Done' { 'ok' } 'WhatIf' { 'blue' } 'Failed' { 'err' } default { 'info' } }
        [void]$aRows.Append("<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td><td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>")
    }

    $channelBadge = switch ($Audit.Status.Kind) { 'Ok' { 'ok' } 'Connectivity' { 'err' } 'Trust' { 'err' } 'NotJoined' { 'info' } default { 'warn' } }
    $channelLabel = switch ($Audit.Status.Kind) { 'Ok' { 'Healthy' } 'Connectivity' { 'Unreachable' } 'Trust' { 'Trust failed' } 'NotJoined' { 'n/a' } default { 'Unknown' } }
    $dcBadge      = if (-not $Audit.Joined) { 'info' } elseif ($Audit.DsGetDc.Success) { 'ok' } else { 'err' }
    $afterText = if ($After) { "$($After.Name) -- $($After.Meaning)" } else { '' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'O.A.T.H. Domain Trust Report' `
        -ScriptName 'O.A.T.H.' `
        -Subtitle   "${orgPrefix}Domain Trust & Secure Channel -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Domain'    = $(if ($Audit.Joined) { $Audit.Domain } else { 'Not joined' })
            'Generated' = $reportDate
            'Mode'      = $(if ($WhatIf) { "$Action (dry run)" } else { $Action })
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Secure Channel', 'DNS', 'Domain Controller', 'Actions Taken')

    $offsetText = if ($null -ne $Audit.Offset) { '{0:N2} s' -f $Audit.Offset } else { 'not measured' }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Domain Trust</div></div>
    <div class="tk-summary-card $channelBadge"><div class="tk-summary-num">$channelLabel</div><div class="tk-summary-lbl">Secure Channel</div></div>
    <div class="tk-summary-card $dcBadge"><div class="tk-summary-num">$(if (-not $Audit.Joined) { 'n/a' } elseif ($Audit.DsGetDc.Success) { 'Found' } else { 'None' })</div><div class="tk-summary-lbl">Domain Controller</div></div>
    <div class="tk-summary-card $(if (@($Audit.Dns | Where-Object { $_.Public }).Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$(@($Audit.Dns | Where-Object { $_.Public }).Count)</div><div class="tk-summary-lbl">Public DNS Servers</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $offsetText)</div><div class="tk-summary-lbl">Clock Offset</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Secure Channel</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">Status</span> $(EscHtml $Audit.Status.Name) -- $(EscHtml $Audit.Status.Meaning)<br/>
      <span class="tk-info-label">Checked with</span> $(if ($Audit.Channel.Verified) { 'nltest /sc_verify (machine password checked with the DC)' } else { 'nltest /sc_query (last known state)' })<br/>
      <span class="tk-info-label">Trusted DC</span> $(EscHtml $(if ($Audit.Channel.TrustedDc) { $Audit.Channel.TrustedDc } else { '-' }))<br/>
      <span class="tk-info-label">Machine password max age</span> $(if ($Audit.MaxPwdAge) { "$($Audit.MaxPwdAge) day(s)" } else { '-' })$(if ($afterText) { "<br/><span class='tk-info-label'>After repair</span> $(EscHtml $afterText)" })
    </div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> DNS</div>
    <div class="tk-card">
      <div class="tk-info-box"><span class="tk-info-label">DC locator SRV</span> _ldap._tcp.dc._msdcs.$(EscHtml $Audit.Domain) -- $(if ($Audit.SrvOk) { 'resolves' } else { 'does not resolve' })</div>
      <table class="tk-table"><thead><tr><th>Interface</th><th>DNS server</th><th>Kind</th></tr></thead>
      <tbody>$($dRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Domain Controller</div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">DC</span> $(EscHtml $(if ($Audit.DsGetDc.Dc) { $Audit.DsGetDc.Dc } else { '-' })) $(EscHtml $Audit.DsGetDc.Address)<br/>
        <span class="tk-info-label">DC site / our site</span> $(EscHtml $Audit.DsGetDc.DcSite) / $(EscHtml $Audit.DsGetDc.OurSite)<br/>
        <span class="tk-info-label">Clock offset</span> $(EscHtml $offsetText)
      </div>
      <table class="tk-table"><thead><tr><th>Port</th><th>Service</th><th>Result</th></tr></thead>
      <tbody>$($pRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Actions Taken</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'O.A.T.H. v5.1')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-OathRun {
    param([string]$Mode)
    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-OathAudit
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Secure channel healthy; DC, DNS and time all in order." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    $after = $null
    if ($Mode -eq 'Repair' -and $audit.Joined) {
        Invoke-OathRepair -Audit $audit
        if (-not $WhatIf) {
            Write-Section "VERIFY"
            $post  = Get-OathSecureChannel -Domain $audit.Domain
            $after = Get-NetlogonStatusInfo -Code $post.StatusCode
            if ($after.Kind -eq 'Ok') { Write-Ok "Secure channel healthy after repair. Sign out and back in with a domain account to confirm." }
            else                      { Write-Warn "Secure channel still reports $($after.Name): $($after.Meaning)" }
            Add-TKNote -Text "OATH repair on $($env:COMPUTERNAME): secure channel now $($after.Name)." -Category $(if ($after.Kind -eq 'Ok') { 'Resolution' } else { 'Issue' }) -ScriptName 'oath'
        }
    }

    $verdict = Get-OathVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    Add-TKNote -Text ("OATH {0} on {1}: verdict {2} ({3} finding(s))." -f $Mode, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'oath'

    Write-Step "Generating HTML report..."
    $html    = Build-OathReport -Audit $audit -Verdict $verdict -After $after
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("OATH_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'oath' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-OathBanner
    Invoke-OathRun -Mode $Action
} else {
    $choice = ''
    do {
        Show-OathBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only diagnosis + HTML report" -ForegroundColor $C.Info
        Write-Host "  [2] Repair  -  audit, resync time, reset the secure channel, reset the machine password if rejected" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($WhatIf) { Write-Host "  Dry run is active — option 2 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-OathRun -Mode 'Audit' }
            '2' { Invoke-OathRun -Mode 'Repair' }
            'Q' { Write-Host ""; Write-Host "  Closing O.A.T.H." -ForegroundColor $C.Header; Write-Host "" }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1, 2 or Q." -ForegroundColor $C.Warning
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
