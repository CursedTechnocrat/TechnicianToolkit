# garm.ps1 - G.A.R.M. — Gets Account-lockout Root causes from Machines
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
    G.A.R.M. — Gets Account-lockout Root causes from Machines
    Active Directory Account Lockout Source Tracer for PowerShell 5.1+

.DESCRIPTION
    Answers "why does this account keep locking out, and from where?" --
    the ticket that unlocking alone never closes, because the stale
    credential that caused it locks the account again within minutes.

    Trace (one account):
      - Lockout state on every domain controller: bad-password count and
        time, lockout time, last logon -- the per-DC view LockoutStatus.exe
        gave, since badPwdCount is not replicated
      - Lockout events (4740) from the PDC emulator, with the caller
        computer each one names
      - Bad-password events from the DCs that saw them: Kerberos
        pre-authentication failures (4771, client IP) and NTLM credential
        validation failures (4776, workstation), with every failure code
        decoded
      - Sources ranked by lockouts and failures; IP addresses resolved
      - Each source machine inspected for what holds the old password:
        services and stored-password scheduled tasks running as the
        account, and the account's sessions (disconnected RDP included)
      - The lockout policy that applies to the account (fine-grained
        policy if one does), and whether the password changed recently

    Sweep (whole domain): every lockout the PDC logged in the window,
    grouped by account with sources, plus every account locked out now.

    Read-only. G.A.R.M. never unlocks an account, resets a password, stops
    a service or ends a session -- it names the source; fix that, then
    unlock with S.P.H.I.N.X. Reading DC Security logs needs Domain Admins
    or Event Log Readers on the DCs, and the Remote Event Log Management
    firewall rule.

.USAGE
    PS C:\> .\garm.ps1                                  # Interactive menu
    PS C:\> .\garm.ps1 -Unattended -Identity jdoe       # Trace one account + HTML report
    PS C:\> .\garm.ps1 -Unattended                      # Sweep every lockout in the last 24 hours
    PS C:\> .\garm.ps1 -Identity jdoe -Hours 72         # Look back three days
    PS C:\> .\garm.ps1 -Identity jdoe -SkipSourceScan   # Don't connect to the source machines

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [string]$Identity,
    [ValidateRange(1, 720)]
    [int]$Hours = 24,
    [string]$Server,
    [switch]$SkipSourceScan,
    [switch]$Transcript
)

# ===========================
# SHARED MODULE BOOTSTRAP
# ===========================
$TKModulePath = Join-Path $PSScriptRoot 'TechnicianToolkit.psm1'
if (-not (Test-Path $TKModulePath)) {
    $TKModuleUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1'
    if ($env:TK_DISABLE_DOWNLOAD -in @('1', 'true')) {
        Write-Host "  [!!] Shared module TechnicianToolkit.psm1 not found, and TK_DISABLE_DOWNLOAD forbids fetching it." -ForegroundColor Red
        Write-Host "       Deploy the module next to this script from an approved release." -ForegroundColor Yellow
        exit 1
    }
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

# Event 4771 (Kerberos pre-authentication failed) result codes. BadPassword
# marks the codes that increment badPwdCount and so count toward a lockout.
$KerberosFailureCodes = @{
    '0x6'  = @{ Name = 'KDC_ERR_C_PRINCIPAL_UNKNOWN'; BadPassword = $false; Meaning = 'The account name does not exist in the domain.' }
    '0x12' = @{ Name = 'KDC_ERR_CLIENT_REVOKED';      BadPassword = $false; Meaning = 'The account is disabled, expired or already locked out -- the source is still trying.' }
    '0x17' = @{ Name = 'KDC_ERR_KEY_EXPIRED';         BadPassword = $false; Meaning = 'The password has expired.' }
    '0x18' = @{ Name = 'KDC_ERR_PREAUTH_FAILED';      BadPassword = $true;  Meaning = 'Wrong password.' }
    '0x25' = @{ Name = 'KRB_AP_ERR_SKEW';             BadPassword = $false; Meaning = 'The client clock is too far from the DC.' }
}

# Event 4776 (NTLM credential validation) status codes.
$NtlmStatusCodes = @{
    '0xc000006a' = @{ Name = 'STATUS_WRONG_PASSWORD';            BadPassword = $true;  Meaning = 'Wrong password.' }
    '0xc0000234' = @{ Name = 'STATUS_ACCOUNT_LOCKED_OUT';        BadPassword = $false; Meaning = 'The account is already locked out -- the source is still trying.' }
    '0xc0000064' = @{ Name = 'STATUS_NO_SUCH_USER';              BadPassword = $false; Meaning = 'The account name does not exist in the domain.' }
    '0xc0000072' = @{ Name = 'STATUS_ACCOUNT_DISABLED';          BadPassword = $false; Meaning = 'The account is disabled.' }
    '0xc000006f' = @{ Name = 'STATUS_INVALID_LOGON_HOURS';       BadPassword = $false; Meaning = 'Sign-in attempted outside the account''s logon hours.' }
    '0xc0000070' = @{ Name = 'STATUS_INVALID_WORKSTATION';       BadPassword = $false; Meaning = 'The account may not sign in from this workstation.' }
    '0xc0000071' = @{ Name = 'STATUS_PASSWORD_EXPIRED';          BadPassword = $false; Meaning = 'The password has expired.' }
    '0xc0000193' = @{ Name = 'STATUS_ACCOUNT_EXPIRED';           BadPassword = $false; Meaning = 'The account has expired.' }
    '0xc0000224' = @{ Name = 'STATUS_PASSWORD_MUST_CHANGE';      BadPassword = $false; Meaning = 'The password must be changed at next sign-in.' }
    '0xc0000133' = @{ Name = 'STATUS_TIME_DIFFERENCE_AT_DC';     BadPassword = $false; Meaning = 'The client clock is too far from the DC.' }
}

# How far a password change still explains a lockout: devices keep retrying
# the old password until someone updates them.
$RecentPasswordChangeDays = 14

# Below this, typos alone lock accounts and a stale credential does it in
# seconds. Microsoft's security baseline uses 10.
$LowLockoutThreshold = 5

# The most source machines inspected per trace; each costs a CIM connection.
$MaxSourcesInspected = 3

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$GarmFindings = @{
    'AccountNotFound' = @{
        Severity = 'Error'
        Title    = 'Account not found'
        Summary  = 'No user in the domain matches the name given.'
        Remedy   = 'Check the spelling. G.A.R.M. accepts a sAMAccountName, DOMAIN\name or a UPN.'
    }
    'AccountLockedOut' = @{
        Severity = 'Error'
        Title    = 'Account is locked out now'
        Summary  = 'The domain reports the account as locked out.'
        Remedy   = 'Fix the source first, then unlock (S.P.H.I.N.X. or Unlock-ADAccount). Unlocking while the stale credential is still in use locks the account again.'
    }
    'RepeatedLockouts' = @{
        Severity = 'Warning'
        Title    = 'Locked out repeatedly'
        Summary  = 'Several lockouts in the window point at a device or service retrying a stale password, not at a user mistyping.'
        Remedy   = 'Work through the sources below. Unlocking alone will not hold.'
    }
    'StaleCredentialSource' = @{
        Severity = 'Warning'
        Title    = 'Bad passwords coming from a machine'
        Summary  = 'This machine is sending the account''s old or wrong password to the domain.'
        Remedy   = 'On that machine check, in order: services and scheduled tasks running as the account, Credential Manager (cmdkey /list), mapped drives with saved credentials, disconnected RDP sessions, and apps with a saved sign-in (Outlook, VPN client, line-of-business apps).'
    }
    'SourceUnknown' = @{
        Severity = 'Warning'
        Title    = 'Lockout with no caller computer'
        Summary  = 'The lockout event names no source. This is typical of a front end authenticating for a client: Exchange ActiveSync phones, OWA, ADFS, RADIUS / NPS (Wi-Fi, VPN) or an app doing LDAP binds.'
        Remedy   = 'Read the Kerberos and NTLM failures below for the IP behind it, then that server''s own logs (IIS on Exchange, ADFS event 411, NPS accounting). A phone with the old password is the most common cause.'
    }
    'SourceIsDomainController' = @{
        Severity = 'Warning'
        Title    = 'Bad passwords arrive through a domain controller'
        Summary  = 'The source is a DC, so it is authenticating on behalf of the real client -- NPS / RADIUS on the DC, an LDAP simple bind from an app, or a service on the DC itself.'
        Remedy   = 'Check NPS logs and LDAP-binding applications (copiers, scan-to-email, line-of-business apps) for the account, and services or tasks on that DC running as it.'
    }
    'ServiceRunsAsUser' = @{
        Severity = 'Error'
        Title    = 'A service runs as this account'
        Summary  = 'A Windows service on the source machine logs on as the account with a stored password -- every restart sends the old one.'
        Remedy   = 'Update the password on the service (services.msc > Log On), or better, move it to a group managed service account.'
    }
    'TaskRunsAsUser' = @{
        Severity = 'Error'
        Title    = 'A scheduled task stores this account''s password'
        Summary  = 'A scheduled task on the source machine runs as the account with a saved password -- every run sends the old one.'
        Remedy   = 'Re-enter the password on the task (Task Scheduler > Properties > Change User), or move it to a service account.'
    }
    'SessionOnSource' = @{
        Severity = 'Warning'
        Title    = 'The account has a session on the source machine'
        Summary  = 'A signed-in or disconnected session keeps the old credentials for mapped drives and apps after a password change.'
        Remedy   = 'Ask the user to sign out of that session, or log it off (logoff <id> /server:<machine>).'
    }
    'SourceUnreachable' = @{
        Severity = 'Info'
        Title    = 'Source machine could not be inspected'
        Summary  = 'G.A.R.M. could not reach the machine over CIM (WinRM or DCOM) to read its services and tasks.'
        Remedy   = 'Check it by hand: services.msc and Task Scheduler for the account, cmdkey /list, net use, and quser.'
    }
    'EventLogUnreadable' = @{
        Severity = 'Warning'
        Title    = 'Security log could not be read on a DC'
        Summary  = 'Without the DC''s Security log the lockout and bad-password events -- and so the source -- cannot be traced.'
        Remedy   = 'Run as a member of Domain Admins or Event Log Readers, and enable the Remote Event Log Management firewall rule on the DCs.'
    }
    'LockoutEventMissing' = @{
        Severity = 'Warning'
        Title    = 'Locked out, but no lockout event found'
        Summary  = 'The account is locked out, yet the PDC emulator logged no event 4740 for it in the window: the Security log has rolled over, or account-management auditing is off.'
        Remedy   = 'Re-run with a longer -Hours. Enable "Audit User Account Management", "Audit Kerberos Authentication Service" and "Audit Credential Validation" in the Default Domain Controllers Policy, and enlarge the DCs'' Security log.'
    }
    'LowLockoutThreshold' = @{
        Severity = 'Info'
        Title    = 'Lockout threshold is low'
        Summary  = 'With so few attempts allowed, a typo or a single stale device locks the account.'
        Remedy   = 'Microsoft''s security baseline uses 10 attempts. Raise the threshold in the domain policy or the fine-grained policy that applies.'
    }
    'PasswordRecentlyChanged' = @{
        Severity = 'Info'
        Title    = 'Password changed recently'
        Summary  = 'Lockouts that start right after a password change almost always come from a device still holding the old password.'
        Remedy   = 'Update the password on the user''s phone mail profile, laptop, and any saved credentials (Credential Manager, mapped drives, VPN client).'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$DnsCache = @{}

function Add-GarmFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $GarmFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
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

function Show-GarmBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗  █████╗ ██████╗ ███╗   ███╗
  ██╔════╝ ██╔══██╗██╔══██╗████╗ ████║
  ██║  ███╗███████║██████╔╝██╔████╔██║
  ██║   ██║██╔══██║██╔══██╗██║╚██╔╝██║
  ╚██████╔╝██║  ██║██║  ██║██║ ╚═╝ ██║
   ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝     ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    G.A.R.M. — Gets Account-lockout Root causes from Machines" -ForegroundColor Cyan
    Write-Host "    Active Directory Account Lockout Source Tracer" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS — no I/O; the Pester suite calls these with captured event XML
# and command output.
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-GarmStatusKey {
    # '0x18', '0x00000018' and '0XC000006A' all normalise to the table key.
    param([string]$Status)
    $s = "$Status".Trim().ToLowerInvariant()
    $m = [regex]::Match($s, '^0x0*([0-9a-f]+)$')
    if ($m.Success) { return '0x' + $m.Groups[1].Value }
    return $s
}

function Get-GarmStatusInfo {
    param([ValidateSet('Kerberos', 'NTLM')][string]$Kind, [string]$Status)
    $key   = ConvertTo-GarmStatusKey -Status $Status
    $table = if ($Kind -eq 'Kerberos') { $KerberosFailureCodes } else { $NtlmStatusCodes }
    $entry = $table[$key]
    if ($entry) { return [PSCustomObject]@{ Key = $key; Name = $entry.Name; BadPassword = [bool]$entry.BadPassword; Meaning = $entry.Meaning } }
    return [PSCustomObject]@{ Key = $key; Name = $key; BadPassword = $false; Meaning = 'Not a code G.A.R.M. recognises.' }
}

function ConvertFrom-GarmFileTime {
    # AD stores badPasswordTime, lockoutTime, lastLogon and pwdLastSet as
    # FILETIME; 0 and the never-expires maximum both mean "not set".
    param($Value)
    if ($null -eq $Value) { return $null }
    $n = $Value -as [long]
    if ($null -eq $n -or $n -le 0 -or $n -ge [DateTime]::MaxValue.ToFileTimeUtc()) { return $null }
    return [DateTime]::FromFileTime($n)
}

function Get-GarmNameVariant {
    # Event-log XPath string comparison is exact, and 4771 / 4776 record the
    # name as the client typed it. Matching the stored, lower and upper case
    # of the sAMAccountName and UPN covers what clients actually send.
    param([string]$SamAccountName, [string]$UserPrincipalName)
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($n in @($SamAccountName, $UserPrincipalName)) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        foreach ($v in @($n, $n.ToLowerInvariant(), $n.ToUpperInvariant())) {
            if (-not $names.Contains($v)) { $names.Add($v) }
        }
    }
    return @($names)
}

function New-GarmEventXPath {
    param([int[]]$EventId, [int]$Hours, [string[]]$UserName = @())
    $ms  = [long]$Hours * 3600000
    $ids = ($EventId | ForEach-Object { "EventID=$_" }) -join ' or '
    $xp  = "*[System[($ids) and TimeCreated[timediff(@SystemTime) <= $ms]]"
    $names = @($UserName | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($names.Count -gt 0) {
        # XPath 1.0 has no escape: quote with " when the name holds an
        # apostrophe (O'Brien); sAMAccountName cannot contain ".
        $terms = foreach ($n in $names) {
            if ($n.Contains("'")) { "Data[@Name='TargetUserName']=`"$n`"" } else { "Data[@Name='TargetUserName']='$n'" }
        }
        $xp += " and EventData[$($terms -join ' or ')]"
    }
    return $xp + ']'
}

function ConvertFrom-GarmEventXml {
    param([string]$Xml)
    $doc = [xml]$Xml
    $sys = $doc.Event.System
    $id  = $sys.EventID
    if ($id -is [System.Xml.XmlElement]) { $id = $id.'#text' }
    $time = $null
    $raw  = "$($sys.TimeCreated.SystemTime)"
    if ($raw) {
        $time = [datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture,
                                  [System.Globalization.DateTimeStyles]::RoundtripKind).ToLocalTime()
    }
    $data = @{}
    foreach ($d in @($doc.Event.EventData.Data)) {
        if ($null -eq $d) { continue }
        $name = "$($d.Name)"
        if (-not $name) { continue }
        $data[$name] = if ($d -is [System.Xml.XmlElement]) { "$($d.InnerText)" } else { "$d" }
    }
    return [PSCustomObject]@{ EventId = [int]$id; Time = $time; Computer = "$($sys.Computer)"; Data = $data }
}

function ConvertTo-GarmSourceHost {
    # Strips the decorations event fields carry: '-' for none, the IPv4-mapped
    # IPv6 prefix, and a leading \\ on workstation names.
    param([string]$Value)
    $v = "$Value".Trim()
    if ($v -eq '' -or $v -eq '-') { return '' }
    $v = $v -replace '^::ffff:', ''
    return $v.TrimStart('\')
}

function ConvertTo-GarmAttempt {
    # One lockout or failed sign-in, whatever event it came from. On 4740 the
    # caller computer is stored in TargetDomainName -- not a domain at all.
    param([object]$Record)
    if (-not $Record) { return $null }
    $d = $Record.Data
    $kind = ''; $source = ''; $status = ''; $statusName = ''; $meaning = ''; $bad = $false
    switch ($Record.EventId) {
        4740 {
            $kind = 'Lockout'; $source = ConvertTo-GarmSourceHost -Value $d['TargetDomainName']
            $statusName = 'Locked out'; $meaning = 'The account crossed the lockout threshold.'
        }
        4771 {
            $kind = 'Kerberos'; $source = ConvertTo-GarmSourceHost -Value $d['IpAddress']
            $info = Get-GarmStatusInfo -Kind 'Kerberos' -Status $d['Status']
            $status = $info.Key; $statusName = $info.Name; $meaning = $info.Meaning; $bad = $info.BadPassword
        }
        4776 {
            $kind = 'NTLM'; $source = ConvertTo-GarmSourceHost -Value $d['Workstation']
            $info = Get-GarmStatusInfo -Kind 'NTLM' -Status $d['Status']
            $status = $info.Key; $statusName = $info.Name; $meaning = $info.Meaning; $bad = $info.BadPassword
        }
        default { return $null }
    }
    # Loopback means the DC authenticated something running on itself.
    if ($source -eq '::1' -or $source -eq '127.0.0.1') { $source = $Record.Computer }
    return [PSCustomObject]@{
        Time = $Record.Time; Dc = $Record.Computer; Kind = $kind; User = "$($d['TargetUserName'])"
        Source = $source; Status = $status; StatusName = $statusName; Meaning = $meaning; BadPassword = $bad
    }
}

function Get-GarmSourceKey {
    # Groups the forms one machine appears under: PC01 in 4740 / 4776, its
    # FQDN once an IP from 4771 is resolved. IPs stay as they are.
    param([string]$Source)
    $s = "$Source".Trim()
    if (-not $s) { return '' }
    $ip = $null
    if ([System.Net.IPAddress]::TryParse($s, [ref]$ip)) { return $s }
    return $s.Split('.')[0].ToUpperInvariant()
}

function Get-GarmSourceRanking {
    param([object[]]$Attempts)
    $rows = foreach ($g in @($Attempts | Where-Object { $_ } | Group-Object { Get-GarmSourceKey -Source $_.Source })) {
        $items = @($g.Group)
        $names = @($items | ForEach-Object { $_.Source } | Where-Object { $_ } | Select-Object -Unique)
        $name  = @($names | Where-Object { $null -eq ($_ -as [System.Net.IPAddress]) } | Select-Object -First 1)
        if (-not $name) { $name = @($names | Select-Object -First 1) }
        $times = @($items | ForEach-Object { $_.Time } | Where-Object { $_ } | Sort-Object)
        [PSCustomObject]@{
            Key          = $g.Name
            Source       = if ($g.Name) { "$($name[0])" } else { '(unknown)' }
            Aliases      = @($names)
            Lockouts     = @($items | Where-Object { $_.Kind -eq 'Lockout' }).Count
            Failures     = @($items | Where-Object { $_.Kind -ne 'Lockout' }).Count
            BadPasswords = @($items | Where-Object { $_.BadPassword }).Count
            FirstSeen    = if ($times.Count) { $times[0] } else { $null }
            LastSeen     = if ($times.Count) { $times[-1] } else { $null }
            Dcs          = @($items | ForEach-Object { $_.Dc } | Where-Object { $_ } | Select-Object -Unique)
            Statuses     = @($items | Where-Object { $_.Kind -ne 'Lockout' } | ForEach-Object { $_.StatusName } | Select-Object -Unique)
        }
    }
    return @($rows | Sort-Object -Property @{ Expression = 'Lockouts'; Descending = $true },
                                           @{ Expression = 'Failures'; Descending = $true },
                                           @{ Expression = 'LastSeen'; Descending = $true })
}

function Test-GarmRunAsMatch {
    # True when a service or task's run-as names this domain account.
    # '.\name' and the NT AUTHORITY / NT SERVICE identities are local, and a
    # same-named account in another domain is someone else.
    param([string]$RunAs, [string]$SamAccountName, [string]$UserPrincipalName, [string]$NetBiosDomain, [string]$DnsDomain)
    $r = "$RunAs".Trim()
    if (-not $r -or -not $SamAccountName) { return $false }
    $m = [regex]::Match($r, '^(?<d>[^\\]+)\\(?<u>.+)$')
    if ($m.Success) {
        if ($m.Groups['u'].Value -ine $SamAccountName) { return $false }
        $d = $m.Groups['d'].Value
        if ($d -eq '.' -or $d -ieq 'NT AUTHORITY' -or $d -ieq 'NT SERVICE') { return $false }
        if (-not $NetBiosDomain -and -not $DnsDomain) { return $true }
        return ($d -ieq $NetBiosDomain -or $d -ieq $DnsDomain)
    }
    if ($r.Contains('@')) {
        if ($UserPrincipalName -and $r -ieq $UserPrincipalName) { return $true }
        $parts = $r.Split('@')
        return ($parts[0] -ieq $SamAccountName -and $DnsDomain -and $parts[1] -ieq $DnsDomain)
    }
    return ($r -ieq $SamAccountName)
}

function ConvertFrom-QuserOutput {
    # quser prints a header, then one row per session; a disconnected
    # session has no SESSIONNAME, so that column is optional.
    param([string[]]$Lines)
    $rows = foreach ($l in $Lines) {
        $m = [regex]::Match("$l", '^\s*>?(\S+)\s+(?:(\S+)\s+)?(\d+)\s+(\S+)\s+(\S+)\s+(.+?)\s*$')
        if (-not $m.Success) { continue }
        [PSCustomObject]@{
            User      = $m.Groups[1].Value
            Session   = $m.Groups[2].Value
            Id        = [int]$m.Groups[3].Value
            State     = $m.Groups[4].Value
            Idle      = $m.Groups[5].Value
            LogonTime = $m.Groups[6].Value
        }
    }
    return @($rows)
}

function Get-GarmVerdict {
    param([object[]]$FindingList, [ValidateSet('Trace', 'Sweep')][string]$Mode = 'Trace')
    $codes = @($FindingList | ForEach-Object { $_.Code })
    $sev   = @($FindingList | ForEach-Object { $_.Severity })
    if ($Mode -eq 'Sweep') {
        if ($codes -contains 'EventLogUnreadable') { return [PSCustomObject]@{ Verdict = 'Unreadable';      Class = 'warn' } }
        if ($sev -contains 'Error')                { return [PSCustomObject]@{ Verdict = 'Accounts locked'; Class = 'err'  } }
        if ($sev -contains 'Warning')              { return [PSCustomObject]@{ Verdict = 'Attention';       Class = 'warn' } }
        return [PSCustomObject]@{ Verdict = 'Quiet'; Class = 'ok' }
    }
    if ($codes -contains 'AccountNotFound') { return [PSCustomObject]@{ Verdict = 'Not found'; Class = 'err' } }
    foreach ($c in 'ServiceRunsAsUser', 'TaskRunsAsUser', 'SessionOnSource') {
        if ($codes -contains $c) { return [PSCustomObject]@{ Verdict = 'Culprit found'; Class = 'err' } }
    }
    if ($codes -contains 'StaleCredentialSource') { return [PSCustomObject]@{ Verdict = 'Source traced'; Class = 'warn' } }
    foreach ($c in 'AccountLockedOut', 'RepeatedLockouts', 'SourceUnknown', 'SourceIsDomainController', 'LockoutEventMissing', 'EventLogUnreadable') {
        if ($codes -contains $c) { return [PSCustomObject]@{ Verdict = 'Untraced'; Class = 'warn' } }
    }
    return [PSCustomObject]@{ Verdict = 'Quiet'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# ACTIVE DIRECTORY
# ─────────────────────────────────────────────────────────────────────────────

function Assert-GarmADModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            return $true
        } catch {
            Write-Fail "Failed to import the ActiveDirectory module: $($_.Exception.Message)"
            return $false
        }
    }

    Write-Host ""
    Write-Host "  ACTIVE DIRECTORY MODULE NOT FOUND" -ForegroundColor $C.Warning
    Write-Host "  The ActiveDirectory PowerShell module ships with RSAT and is required" -ForegroundColor $C.Info
    Write-Host "  to read lockout state from the domain controllers." -ForegroundColor $C.Info
    Write-Host ""

    if ($Unattended) {
        Write-Fail "RSAT ActiveDirectory tools are not installed. Install them and re-run:"
        Write-Info "Add-WindowsCapability -Online -Name RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
        return $false
    }

    Write-Host -NoNewline "  Install the RSAT ActiveDirectory tools now? (Y/N) " -ForegroundColor $C.Header
    $answer = Read-Host
    if ($answer -notmatch '^(y|yes)$') {
        Write-Info 'Cancelled.'
        return $false
    }

    Write-Step 'Installing RSAT ActiveDirectory tools — this may take several minutes...'
    try {
        $result = Add-WindowsCapability -Online -Name 'RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop
        if ($result.RestartNeeded) { Write-Warn 'A restart may be required to complete installation.' }
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Ok 'Module installed and imported.'
        return $true
    } catch {
        Write-Fail "Automatic installation failed: $($_.Exception.Message)"
        Write-Info 'Install manually: Settings > Optional Features > RSAT: Active Directory Domain Services.'
        return $false
    }
}

function Get-GarmDomain {
    $common = @{}
    if ($Server) { $common.Server = $Server }
    $domain = Get-ADDomain @common -ErrorAction Stop
    $dcs = @(Get-ADDomainController -Filter * @common -ErrorAction Stop | Sort-Object HostName)
    return [PSCustomObject]@{
        DnsRoot = "$($domain.DNSRoot)"
        NetBios = "$($domain.NetBIOSName)"
        Pdc     = "$($domain.PDCEmulator)"
        Dcs     = @($dcs | ForEach-Object { [PSCustomObject]@{ HostName = "$($_.HostName)"; Name = "$($_.Name)"; Site = "$($_.Site)" } })
    }
}

function Resolve-GarmUser {
    param([string]$Name, [string]$Pdc)
    $props = @('badPwdCount', 'badPasswordTime', 'lockoutTime', 'LockedOut', 'lastLogon', 'pwdLastSet', 'UserPrincipalName', 'DisplayName', 'Enabled')
    $n = "$Name".Trim()
    if ($n -match '^[^\\]+\\(.+)$') { $n = $Matches[1] }
    try {
        if ($n.Contains('@')) {
            $upn = $n.Replace("'", "''")
            $u = @(Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Server $Pdc -Properties $props -ErrorAction Stop)
            if ($u.Count -eq 1) { return $u[0] }
            $n = $n.Split('@')[0]
        }
        return Get-ADUser -Identity $n -Server $Pdc -Properties $props -ErrorAction Stop
    } catch {
        return $null
    }
}

function Get-GarmDcStatus {
    # badPwdCount and badPasswordTime are kept per DC and never replicated,
    # so the only way to see where the failures landed is to ask every DC.
    param([object]$User, [object]$Domain)
    $rows = foreach ($dc in $Domain.Dcs) {
        Write-Step "Reading lockout state from $($dc.HostName)..."
        try {
            $u = Get-ADUser -Identity $User.DistinguishedName -Server $dc.HostName -ErrorAction Stop `
                    -Properties 'badPwdCount', 'badPasswordTime', 'lockoutTime', 'LockedOut', 'lastLogon', 'pwdLastSet'
            [PSCustomObject]@{
                Dc = $dc.HostName; Site = $dc.Site; IsPdc = ($dc.HostName -ieq $Domain.Pdc); Reachable = $true
                BadPwdCount = [int]$u.badPwdCount
                LastBadPassword = ConvertFrom-GarmFileTime -Value $u.badPasswordTime
                LockoutTime = ConvertFrom-GarmFileTime -Value $u.lockoutTime
                LockedOut = [bool]$u.LockedOut
                LastLogon = ConvertFrom-GarmFileTime -Value $u.lastLogon
                PwdLastSet = ConvertFrom-GarmFileTime -Value $u.pwdLastSet
            }
        } catch {
            [PSCustomObject]@{
                Dc = $dc.HostName; Site = $dc.Site; IsPdc = ($dc.HostName -ieq $Domain.Pdc); Reachable = $false
                BadPwdCount = $null; LastBadPassword = $null; LockoutTime = $null; LockedOut = $false; LastLogon = $null; PwdLastSet = $null
            }
        }
    }
    return @($rows)
}

function Get-GarmLockoutPolicy {
    param([object]$User, [string]$Pdc)
    $policy = $null; $source = 'Default domain policy'
    if ($User) {
        try {
            $policy = Get-ADUserResultantPasswordPolicy -Identity $User.DistinguishedName -Server $Pdc -ErrorAction Stop
            if ($policy) { $source = "Fine-grained policy: $($policy.Name)" }
        } catch { $policy = $null }
    }
    if (-not $policy) {
        try { $policy = Get-ADDefaultDomainPasswordPolicy -Server $Pdc -ErrorAction Stop } catch { $policy = $null }
        $source = 'Default domain policy'
    }
    if (-not $policy) { return $null }
    return [PSCustomObject]@{
        Source             = $source
        Threshold          = [int]$policy.LockoutThreshold
        DurationMinutes    = [int]([timespan]$policy.LockoutDuration).TotalMinutes
        WindowMinutes      = [int]([timespan]$policy.LockoutObservationWindow).TotalMinutes
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# EVENT LOGS AND SOURCE MACHINES
# ─────────────────────────────────────────────────────────────────────────────

function Get-GarmEvent {
    # Returns parsed events, or $null when the log could not be read at all
    # (access denied, firewall, DC offline) -- distinct from an empty result.
    param([string]$Computer, [int[]]$EventId, [string[]]$UserName = @(), [int]$MaxEvents = 500)
    $xpath = New-GarmEventXPath -EventId $EventId -Hours $Hours -UserName $UserName
    try {
        $events = @(Get-WinEvent -ComputerName $Computer -LogName 'Security' -FilterXPath $xpath -MaxEvents $MaxEvents -ErrorAction Stop)
    } catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*' -or $_.Exception.Message -match '(?i)no events were found') { return @() }
        Add-GarmFinding -Code 'EventLogUnreadable' -Detail "$($Computer): $($_.Exception.Message)"
        return $null
    }
    return @($events | ForEach-Object { ConvertFrom-GarmEventXml -Xml $_.ToXml() })
}

function Resolve-GarmSourceName {
    # Reverse-resolves an IP source to a host name so it groups with the
    # name the same machine has in other events. Cached per run.
    param([string]$Source)
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse("$Source", [ref]$ip)) { return $Source }
    if ($DnsCache.ContainsKey($Source)) { return $DnsCache[$Source] }
    $name = $Source
    try { $name = [System.Net.Dns]::GetHostEntry($ip).HostName } catch { $name = $Source }
    $DnsCache[$Source] = $name
    return $name
}

function Invoke-GarmSourceInspection {
    # Read-only look at one source machine for what holds the account's
    # password. CIM over WinRM first, then DCOM for machines without it.
    param([string]$ComputerName, [object]$User, [object]$Domain)
    $result = [PSCustomObject]@{ Computer = $ComputerName; Reachable = $false; Services = @(); Tasks = @(); Sessions = @(); Error = '' }
    $session = $null
    try {
        $session = New-CimSession -ComputerName $ComputerName -OperationTimeoutSec 20 -ErrorAction Stop
    } catch {
        try {
            $session = New-CimSession -ComputerName $ComputerName -SessionOption (New-CimSessionOption -Protocol Dcom) -OperationTimeoutSec 20 -ErrorAction Stop
        } catch {
            $result.Error = $_.Exception.Message
            return $result
        }
    }
    $result.Reachable = $true
    $match = @{ SamAccountName = $User.SamAccountName; UserPrincipalName = $User.UserPrincipalName; NetBiosDomain = $Domain.NetBios; DnsDomain = $Domain.DnsRoot }
    try {
        $result.Services = @(Get-CimInstance -CimSession $session -ClassName Win32_Service -ErrorAction Stop |
            Where-Object { Test-GarmRunAsMatch -RunAs $_.StartName @match } |
            ForEach-Object { [PSCustomObject]@{ Name = "$($_.Name)"; DisplayName = "$($_.DisplayName)"; State = "$($_.State)"; StartMode = "$($_.StartMode)"; RunAs = "$($_.StartName)" } })
    } catch { $result.Error = "Services: $($_.Exception.Message)" }
    try {
        # Only a task that stores a password (LogonType Password) can send a
        # stale one; interactive and S4U tasks borrow a token instead.
        $result.Tasks = @(Get-ScheduledTask -CimSession $session -ErrorAction Stop |
            Where-Object { ("$($_.Principal.LogonType)" -eq 'Password' -or "$($_.Principal.LogonType)" -eq '1') -and (Test-GarmRunAsMatch -RunAs $_.Principal.UserId @match) } |
            ForEach-Object { [PSCustomObject]@{ Name = "$($_.TaskPath)$($_.TaskName)"; State = "$($_.State)"; RunAs = "$($_.Principal.UserId)" } })
    } catch {
        if (-not $result.Error) { $result.Error = "Tasks: $($_.Exception.Message)" }
    }
    Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue

    try {
        $out = & quser.exe "/server:$ComputerName" 2>&1 | ForEach-Object { "$_" }
        $result.Sessions = @(ConvertFrom-QuserOutput -Lines $out | Where-Object { $_.User -ieq $User.SamAccountName })
    } catch { $result.Sessions = @() }
    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# TRACE AND SWEEP
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-GarmTrace {
    param([string]$Name, [object]$Domain)

    Write-Section "ACCOUNT"
    $user = Resolve-GarmUser -Name $Name -Pdc $Domain.Pdc
    if (-not $user) {
        Add-GarmFinding -Code 'AccountNotFound' -Detail $Name
        Write-Fail "No account matches '$Name'."
        return [PSCustomObject]@{ Mode = 'Trace'; Name = $Name; User = $null; DcStatus = @(); Lockouts = @(); Failures = @(); Sources = @(); Inspections = @(); Policy = $null; LockedOut = $false }
    }
    Write-Info "$($user.SamAccountName)  ($($user.DisplayName))  $($user.UserPrincipalName)"

    Write-Section "LOCKOUT STATE BY DOMAIN CONTROLLER"
    $dcStatus  = Get-GarmDcStatus -User $user -Domain $Domain
    foreach ($r in $dcStatus) {
        if (-not $r.Reachable) { Write-Warn ("{0,-34} unreachable" -f $r.Dc); continue }
        Write-Info ("{0,-34} bad={1,-3} lastBad={2,-19} locked={3}" -f $r.Dc, $r.BadPwdCount,
            $(if ($r.LastBadPassword) { $r.LastBadPassword.ToString('yyyy-MM-dd HH:mm:ss') } else { '-' }), $r.LockedOut)
    }
    $lockedOut = [bool]$user.LockedOut -or [bool](@($dcStatus | Where-Object { $_.LockedOut }).Count)
    if ($lockedOut) {
        $when = @($dcStatus | ForEach-Object { $_.LockoutTime } | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1)
        Add-GarmFinding -Code 'AccountLockedOut' -Detail $(if ($when) { "Locked at $($when[0].ToString('yyyy-MM-dd HH:mm:ss'))" } else { '' })
        Write-Fail "Locked out."
    } else {
        Write-Ok "Not locked out."
    }

    $pwdSet = ConvertFrom-GarmFileTime -Value $user.pwdLastSet
    if ($pwdSet -and $pwdSet -gt (Get-Date).AddDays(-$RecentPasswordChangeDays)) {
        Add-GarmFinding -Code 'PasswordRecentlyChanged' -Detail "Password set $($pwdSet.ToString('yyyy-MM-dd HH:mm'))"
    }

    $policy = Get-GarmLockoutPolicy -User $user -Pdc $Domain.Pdc
    if ($policy -and $policy.Threshold -gt 0 -and $policy.Threshold -lt $LowLockoutThreshold) {
        Add-GarmFinding -Code 'LowLockoutThreshold' -Detail "$($policy.Threshold) attempts ($($policy.Source))"
    }

    Write-Section "LOCKOUT EVENTS (PDC $($Domain.Pdc))"
    $names = Get-GarmNameVariant -SamAccountName $user.SamAccountName -UserPrincipalName $user.UserPrincipalName
    $lockoutEvents = Get-GarmEvent -Computer $Domain.Pdc -EventId @(4740) -UserName $names
    $lockouts = @($lockoutEvents | ForEach-Object { ConvertTo-GarmAttempt -Record $_ } | Where-Object { $_ })
    Write-Info "$($lockouts.Count) lockout(s) in the last $Hours hour(s)."
    if ($null -ne $lockoutEvents -and $lockouts.Count -eq 0 -and $lockedOut) {
        Add-GarmFinding -Code 'LockoutEventMissing' -Detail "No event 4740 for $($user.SamAccountName) on $($Domain.Pdc) in the last $Hours hour(s)."
    }
    if ($lockouts.Count -ge 3) { Add-GarmFinding -Code 'RepeatedLockouts' -Detail "$($lockouts.Count) lockouts in the last $Hours hour(s)" }

    Write-Section "BAD-PASSWORD EVENTS"
    $cutoff = (Get-Date).AddHours(-$Hours)
    $dcsToRead = @($dcStatus | Where-Object { $_.Reachable -and ($_.IsPdc -or ($_.LastBadPassword -and $_.LastBadPassword -gt $cutoff)) } | ForEach-Object { $_.Dc })
    if ($dcsToRead.Count -eq 0) { $dcsToRead = @($Domain.Pdc) }
    $failures = @()
    foreach ($dc in $dcsToRead) {
        Write-Step "Reading Kerberos / NTLM failures on $dc..."
        $ev = Get-GarmEvent -Computer $dc -EventId @(4771, 4776) -UserName $names
        $failures += @($ev | ForEach-Object { ConvertTo-GarmAttempt -Record $_ } | Where-Object { $_ -and $_.Status -ne '0x0' })
    }
    Write-Info "$($failures.Count) failed sign-in(s) on $($dcsToRead.Count) DC(s)."

    foreach ($a in @($lockouts + $failures)) { if ($a.Source) { $a.Source = Resolve-GarmSourceName -Source $a.Source } }

    Write-Section "SOURCES"
    $sources  = Get-GarmSourceRanking -Attempts @($lockouts + $failures)
    $dcKeys   = @($Domain.Dcs | ForEach-Object { Get-GarmSourceKey -Source $_.HostName })
    foreach ($s in $sources) {
        Write-Info ("{0,-34} lockouts={1,-3} failures={2,-4} last={3}" -f $s.Source, $s.Lockouts, $s.Failures,
            $(if ($s.LastSeen) { $s.LastSeen.ToString('yyyy-MM-dd HH:mm:ss') } else { '-' }))
    }
    if (@($lockouts | Where-Object { -not $_.Source }).Count -gt 0) {
        Add-GarmFinding -Code 'SourceUnknown' -Detail "$(@($lockouts | Where-Object { -not $_.Source }).Count) lockout event(s) name no caller computer"
    }
    $known = @($sources | Where-Object { $_.Key })
    foreach ($s in @($known | Select-Object -First 5)) {
        $detail = "$($s.Source): $($s.Lockouts) lockout(s), $($s.Failures) failed sign-in(s)" + $(if ($s.Statuses.Count) { " ($($s.Statuses -join ', '))" } else { '' })
        if ($dcKeys -contains $s.Key) { Add-GarmFinding -Code 'SourceIsDomainController' -Detail $detail }
        else                          { Add-GarmFinding -Code 'StaleCredentialSource'   -Detail $detail }
    }

    $inspections = @()
    $targets = @($known | Where-Object { $dcKeys -notcontains $_.Key } | Select-Object -First $MaxSourcesInspected)
    if ($SkipSourceScan) {
        if ($targets.Count) { Write-Info "Source inspection skipped (-SkipSourceScan)." }
    } else {
        foreach ($t in $targets) {
            Write-Section "SOURCE — $($t.Source)"
            $insp = Invoke-GarmSourceInspection -ComputerName $t.Source -User $user -Domain $Domain
            $inspections += $insp
            if (-not $insp.Reachable) {
                Write-Warn "Could not connect: $($insp.Error)"
                Add-GarmFinding -Code 'SourceUnreachable' -Detail "$($t.Source): $($insp.Error)"
                continue
            }
            foreach ($svc in $insp.Services) { Add-GarmFinding -Code 'ServiceRunsAsUser' -Detail "$($t.Source): service $($svc.Name) ($($svc.DisplayName)), $($svc.State), runs as $($svc.RunAs)"; Write-Fail "Service $($svc.Name) runs as $($svc.RunAs)" }
            foreach ($tk in $insp.Tasks)     { Add-GarmFinding -Code 'TaskRunsAsUser'    -Detail "$($t.Source): task $($tk.Name), runs as $($tk.RunAs)"; Write-Fail "Task $($tk.Name) runs as $($tk.RunAs)" }
            foreach ($ss in $insp.Sessions)  { Add-GarmFinding -Code 'SessionOnSource'   -Detail "$($t.Source): session $($ss.Id) $($ss.State), logged on $($ss.LogonTime)"; Write-Warn "Session $($ss.Id) is $($ss.State)" }
            if (-not ($insp.Services.Count + $insp.Tasks.Count + $insp.Sessions.Count)) {
                Write-Info "No service, stored-password task or session for the account -- check Credential Manager, mapped drives and apps on this machine."
            }
        }
    }

    return [PSCustomObject]@{
        Mode = 'Trace'; Name = $Name; User = $user; DcStatus = @($dcStatus); Lockouts = @($lockouts); Failures = @($failures)
        Sources = @($sources); Inspections = @($inspections); Policy = $policy; LockedOut = $lockedOut
    }
}

function Invoke-GarmSweep {
    param([object]$Domain)

    Write-Section "ACCOUNTS LOCKED OUT NOW"
    $locked = @()
    try {
        $locked = @(Get-ADUser -LDAPFilter '(lockoutTime>=1)' -Server $Domain.Pdc -Properties 'LockedOut', 'lockoutTime', 'DisplayName' -ErrorAction Stop |
            Where-Object { $_.LockedOut } |
            ForEach-Object { [PSCustomObject]@{ Sam = "$($_.SamAccountName)"; Name = "$($_.DisplayName)"; LockoutTime = ConvertFrom-GarmFileTime -Value $_.lockoutTime } })
    } catch {
        Write-Warn "Could not list locked accounts: $($_.Exception.Message)"
    }
    Write-Info "$($locked.Count) account(s) locked out."

    Write-Section "LOCKOUT EVENTS (PDC $($Domain.Pdc))"
    $events   = Get-GarmEvent -Computer $Domain.Pdc -EventId @(4740) -MaxEvents 2000
    $lockouts = @($events | ForEach-Object { ConvertTo-GarmAttempt -Record $_ } | Where-Object { $_ })
    Write-Info "$($lockouts.Count) lockout(s) in the last $Hours hour(s)."

    $dcKeys = @($Domain.Dcs | ForEach-Object { Get-GarmSourceKey -Source $_.HostName })
    $byUser = foreach ($g in @($lockouts | Group-Object { $_.User.ToUpperInvariant() })) {
        $items   = @($g.Group)
        $sources = @($items | ForEach-Object { $_.Source } | Where-Object { $_ } | Sort-Object -Unique)
        $times   = @($items | ForEach-Object { $_.Time } | Sort-Object)
        [PSCustomObject]@{
            User        = "$($items[0].User)"
            Count       = $items.Count
            First       = $times[0]
            Last        = $times[-1]
            Sources     = $sources
            NoCaller    = @($items | Where-Object { -not $_.Source }).Count
            ViaDc       = @($sources | Where-Object { $dcKeys -contains (Get-GarmSourceKey -Source $_) })
            LockedNow   = [bool](@($locked | Where-Object { $_.Sam -ieq $items[0].User }).Count)
        }
    }
    $byUser = @($byUser | Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Last'; Descending = $true })
    foreach ($u in $byUser) {
        Write-Info ("{0,-24} {1,3} lockout(s)  last {2}  from {3}" -f $u.User, $u.Count, (Format-GarmTime $u.Last),
            $(if ($u.Sources.Count) { $u.Sources -join ', ' } else { '(no caller)' }))
    }

    foreach ($l in $locked) {
        $hit = @($byUser | Where-Object { $_.User -ieq $l.Sam } | Select-Object -First 1)
        $from = if ($hit -and $hit[0].Sources.Count) { "; source(s): $($hit[0].Sources -join ', ')" } else { '' }
        Add-GarmFinding -Code 'AccountLockedOut' -Detail ("{0}{1}{2}" -f $l.Sam, $(if ($l.LockoutTime) { " at $($l.LockoutTime.ToString('yyyy-MM-dd HH:mm'))" } else { '' }), $from)
    }
    foreach ($u in @($byUser | Where-Object { $_.Count -ge 3 })) {
        Add-GarmFinding -Code 'RepeatedLockouts' -Detail ("{0}: {1} lockouts; source(s): {2}" -f $u.User, $u.Count, $(if ($u.Sources.Count) { $u.Sources -join ', ' } else { 'none named' }))
    }
    $noCaller = @($byUser | Where-Object { $_.NoCaller -gt 0 } | ForEach-Object { $_.User })
    if ($noCaller.Count) { Add-GarmFinding -Code 'SourceUnknown' -Detail "Account(s): $($noCaller -join ', ')" }
    $viaDc = @($byUser | Where-Object { $_.ViaDc.Count } | ForEach-Object { "$($_.User) via $($_.ViaDc -join ', ')" })
    if ($viaDc.Count) { Add-GarmFinding -Code 'SourceIsDomainController' -Detail ($viaDc -join '; ') }

    return [PSCustomObject]@{ Mode = 'Sweep'; Locked = @($locked); Lockouts = @($lockouts); ByUser = @($byUser) }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Format-GarmTime {
    param($Value)
    if ($null -eq $Value) { return '-' }
    return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm:ss')
}

function Get-GarmAttemptRow {
    param([object[]]$Attempts, [int]$Columns)
    $sb = [System.Text.StringBuilder]::new()
    if (@($Attempts).Count -eq 0) { [void]$sb.Append("<tr><td colspan='$Columns'>None in the window.</td></tr>"); return $sb.ToString() }
    foreach ($a in @($Attempts | Sort-Object Time -Descending)) {
        $badge = if ($a.Kind -eq 'Lockout') { 'err' } elseif ($a.BadPassword) { 'warn' } else { 'info' }
        [void]$sb.Append("<tr><td class='tk-mono'>$(Format-GarmTime $a.Time)</td><td>$(EscHtml $a.Dc)</td><td>$(EscHtml $a.User)</td>" +
            "<td>$(EscHtml $(if ($a.Source) { $a.Source } else { '(none)' }))</td>" +
            "<td><span class='tk-badge-$badge'>$(EscHtml $a.Kind)</span></td>" +
            "<td><span class='tk-mono'>$(EscHtml $a.StatusName)</span><br/>$(EscHtml $a.Meaning)</td></tr>")
    }
    return $sb.ToString()
}

function Build-GarmReport {
    param([object]$Result, [object]$Domain, [object]$Verdict)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
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
    $findingsSection = @"
  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>
"@

    if ($Result.Mode -eq 'Trace') {
        $subject = if ($Result.User) { "$($Result.User.SamAccountName)" } else { $Result.Name }
        $nav = @('Findings', 'Sources', 'Domain Controllers', 'Lockout Events', 'Failed Sign-ins', 'Source Machines')

        $sRows = [System.Text.StringBuilder]::new()
        if ($Result.Sources.Count -eq 0) { [void]$sRows.Append("<tr><td colspan='6'>No lockout or failed sign-in events in the window.</td></tr>") }
        foreach ($s in $Result.Sources) {
            [void]$sRows.Append("<tr><td><strong>$(EscHtml $s.Source)</strong>" +
                $(if ($s.Aliases.Count -gt 1) { "<br/><span class='tk-mono'>$(EscHtml ($s.Aliases -join ', '))</span>" } else { '' }) + "</td>" +
                "<td>$($s.Lockouts)</td><td>$($s.Failures)</td><td>$(EscHtml ($s.Statuses -join ', '))</td>" +
                "<td class='tk-mono'>$(Format-GarmTime $s.FirstSeen)<br/>$(Format-GarmTime $s.LastSeen)</td><td>$(EscHtml ($s.Dcs -join ', '))</td></tr>")
        }

        $dRows = [System.Text.StringBuilder]::new()
        if ($Result.DcStatus.Count -eq 0) { [void]$dRows.Append("<tr><td colspan='7'>Not read.</td></tr>") }
        foreach ($r in $Result.DcStatus) {
            if (-not $r.Reachable) {
                [void]$dRows.Append("<tr><td>$(EscHtml $r.Dc)$(if ($r.IsPdc) { ' <span class=''tk-badge-blue''>PDC</span>' })</td><td>$(EscHtml $r.Site)</td><td colspan='5'><span class='tk-badge-warn'>Unreachable</span></td></tr>")
                continue
            }
            $lock = if ($r.LockedOut) { "<span class='tk-badge-err'>Locked</span>" } else { "<span class='tk-badge-ok'>No</span>" }
            [void]$dRows.Append("<tr><td>$(EscHtml $r.Dc)$(if ($r.IsPdc) { ' <span class=''tk-badge-blue''>PDC</span>' })</td><td>$(EscHtml $r.Site)</td>" +
                "<td>$($r.BadPwdCount)</td><td class='tk-mono'>$(Format-GarmTime $r.LastBadPassword)</td><td>$lock</td>" +
                "<td class='tk-mono'>$(Format-GarmTime $r.LockoutTime)</td><td class='tk-mono'>$(Format-GarmTime $r.LastLogon)</td></tr>")
        }

        $iRows = [System.Text.StringBuilder]::new()
        if ($Result.Inspections.Count -eq 0) {
            $why = if ($SkipSourceScan) { 'Skipped (-SkipSourceScan).' } else { 'No source machine to inspect.' }
            [void]$iRows.Append("<tr><td colspan='3'>$why</td></tr>")
        }
        foreach ($i in $Result.Inspections) {
            if (-not $i.Reachable) { [void]$iRows.Append("<tr><td>$(EscHtml $i.Computer)</td><td><span class='tk-badge-warn'>Unreachable</span></td><td>$(EscHtml $i.Error)</td></tr>"); continue }
            $items = @()
            $items += @($i.Services | ForEach-Object { "Service <strong>$(EscHtml $_.Name)</strong> ($(EscHtml $_.DisplayName)) -- $(EscHtml $_.State), runs as $(EscHtml $_.RunAs)" })
            $items += @($i.Tasks    | ForEach-Object { "Task <strong>$(EscHtml $_.Name)</strong> -- $(EscHtml $_.State), runs as $(EscHtml $_.RunAs)" })
            $items += @($i.Sessions | ForEach-Object { "Session $($_.Id) <strong>$(EscHtml $_.State)</strong> ($(EscHtml $_.Session)), logged on $(EscHtml $_.LogonTime)" })
            $badge = if ($items.Count) { "<span class='tk-badge-err'>$($items.Count) found</span>" } else { "<span class='tk-badge-ok'>Clean</span>" }
            $body  = if ($items.Count) { $items -join '<br/>' } else { 'No service, stored-password task or session for the account. Check Credential Manager, mapped drives and saved app sign-ins on this machine.' }
            [void]$iRows.Append("<tr><td>$(EscHtml $i.Computer)</td><td>$badge</td><td>$body</td></tr>")
        }

        $p = $Result.Policy
        $policyText = if ($p) { "$($p.Threshold) attempt(s) within $($p.WindowMinutes) min; locked for $(if ($p.DurationMinutes -eq 0) { 'until unlocked by an admin' } else { "$($p.DurationMinutes) min" }) -- $(EscHtml $p.Source)" } else { 'Not read' }
        $u = $Result.User

        $cards = @"
  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(if ($Result.LockedOut) { 'err' } else { 'ok' })"><div class="tk-summary-num">$(if ($Result.LockedOut) { 'Locked' } else { 'No' })</div><div class="tk-summary-lbl">Locked Out Now</div></div>
    <div class="tk-summary-card $(if ($Result.Lockouts.Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Result.Lockouts.Count)</div><div class="tk-summary-lbl">Lockouts ($Hours h)</div></div>
    <div class="tk-summary-card $(if ($Result.Failures.Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Result.Failures.Count)</div><div class="tk-summary-lbl">Failed Sign-ins</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Result.Sources | Where-Object { $_.Key }).Count)</div><div class="tk-summary-lbl">Sources</div></div>
  </div>
  <div class="tk-card"><div class="tk-info-box">
    <span class="tk-info-label">Account</span> $(if ($u) { "$(EscHtml $u.SamAccountName) -- $(EscHtml $u.DisplayName) -- $(EscHtml $u.UserPrincipalName)$(if (-not $u.Enabled) { ' (disabled)' })" } else { EscHtml $Result.Name })<br/>
    <span class="tk-info-label">Password last set</span> $(if ($u) { Format-GarmTime (ConvertFrom-GarmFileTime -Value $u.pwdLastSet) } else { '-' })<br/>
    <span class="tk-info-label">Lockout policy</span> $policyText
  </div></div>
"@
        $body = $findingsSection + @"
  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Sources</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Source</th><th>Lockouts</th><th>Failed sign-ins</th><th>Codes</th><th>First / last seen</th><th>Seen by DC</th></tr></thead>
      <tbody>$($sRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Domain Controllers</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>DC</th><th>Site</th><th>Bad pwd count</th><th>Last bad password</th><th>Locked</th><th>Lockout time</th><th>Last logon</th></tr></thead>
      <tbody>$($dRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Lockout Events</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>DC</th><th>Account</th><th>Caller computer</th><th>Kind</th><th>Result</th></tr></thead>
      <tbody>$(Get-GarmAttemptRow -Attempts $Result.Lockouts -Columns 6)</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Failed Sign-ins</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>DC</th><th>Name sent</th><th>Source</th><th>Kind</th><th>Result</th></tr></thead>
      <tbody>$(Get-GarmAttemptRow -Attempts $Result.Failures -Columns 6)</tbody></table></div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Source Machines</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Machine</th><th>Result</th><th>What holds the account's password</th></tr></thead>
      <tbody>$($iRows.ToString())</tbody></table></div>
  </div>
"@
    } else {
        $subject = 'All accounts'
        $nav = @('Findings', 'Locked Accounts', 'Lockouts by Account', 'Lockout Events')

        $lRows = [System.Text.StringBuilder]::new()
        if ($Result.Locked.Count -eq 0) { [void]$lRows.Append("<tr><td colspan='3'>No account is locked out.</td></tr>") }
        foreach ($l in $Result.Locked) { [void]$lRows.Append("<tr><td>$(EscHtml $l.Sam)</td><td>$(EscHtml $l.Name)</td><td class='tk-mono'>$(Format-GarmTime $l.LockoutTime)</td></tr>") }

        $uRows = [System.Text.StringBuilder]::new()
        if ($Result.ByUser.Count -eq 0) { [void]$uRows.Append("<tr><td colspan='5'>No lockouts in the window.</td></tr>") }
        foreach ($b in $Result.ByUser) {
            $now = if ($b.LockedNow) { "<span class='tk-badge-err'>Locked</span>" } else { "<span class='tk-badge-ok'>No</span>" }
            $src = if ($b.Sources.Count) { EscHtml ($b.Sources -join ', ') } else { '' }
            if ($b.NoCaller) { $src += $(if ($src) { '<br/>' } else { '' }) + "<span class='tk-badge-warn'>$($b.NoCaller) with no caller</span>" }
            [void]$uRows.Append("<tr><td><strong>$(EscHtml $b.User)</strong></td><td>$($b.Count)</td><td class='tk-mono'>$(Format-GarmTime $b.First)<br/>$(Format-GarmTime $b.Last)</td><td>$src</td><td>$now</td></tr>")
        }

        $cards = @"
  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(if ($Result.Locked.Count) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($Result.Locked.Count)</div><div class="tk-summary-lbl">Locked Out Now</div></div>
    <div class="tk-summary-card $(if ($Result.Lockouts.Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Result.Lockouts.Count)</div><div class="tk-summary-lbl">Lockouts ($Hours h)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Result.ByUser.Count)</div><div class="tk-summary-lbl">Accounts Affected</div></div>
  </div>
"@
        $body = $findingsSection + @"
  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Locked Accounts</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Account</th><th>Name</th><th>Locked at</th></tr></thead>
      <tbody>$($lRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Lockouts by Account</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Account</th><th>Lockouts</th><th>First / last</th><th>Caller computers</th><th>Locked now</th></tr></thead>
      <tbody>$($uRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Lockout Events</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>DC</th><th>Account</th><th>Caller computer</th><th>Kind</th><th>Result</th></tr></thead>
      <tbody>$(Get-GarmAttemptRow -Attempts $Result.Lockouts -Columns 6)</tbody></table></div>
  </div>
"@
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'G.A.R.M. Account Lockout Report' `
        -ScriptName 'G.A.R.M.' `
        -Subtitle   "${orgPrefix}Account Lockout Source -- $subject" `
        -MetaItems  ([ordered]@{
            'Domain'    = $Domain.DnsRoot
            'PDC'       = $Domain.Pdc
            'Mode'      = $Result.Mode
            'Window'    = "Last $Hours hour(s)"
            'Generated' = $reportDate
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   $nav

    return $htmlHead + $cards + $body + (Get-TKHtmlFoot -ScriptName 'G.A.R.M. v6.0')
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-GarmRun {
    param([ValidateSet('Trace', 'Sweep')][string]$Mode, [string]$Name)
    $Findings.Clear()
    $DnsCache.Clear()

    if (-not (Assert-GarmADModule)) { return }
    Write-Section "DOMAIN"
    try {
        $domain = Get-GarmDomain
    } catch {
        Write-Fail "Could not reach the domain: $($_.Exception.Message)"
        Write-TKError -ScriptName 'garm' -Message "Domain lookup failed: $($_.Exception.Message)" -Category 'Active Directory'
        return
    }
    Write-Info "$($domain.DnsRoot)  |  PDC emulator $($domain.Pdc)  |  $($domain.Dcs.Count) DC(s)  |  last $Hours hour(s)"

    $result = if ($Mode -eq 'Trace') { Invoke-GarmTrace -Name $Name -Domain $domain } else { Invoke-GarmSweep -Domain $domain }

    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "No lockouts or failed sign-ins in the window." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    $verdict = Get-GarmVerdict -FindingList $Findings.ToArray() -Mode $Mode
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    $subject = if ($Mode -eq 'Trace') { $Name } else { 'all accounts' }
    $top = @($Findings | Where-Object { $_.Code -in 'ServiceRunsAsUser', 'TaskRunsAsUser', 'SessionOnSource', 'StaleCredentialSource' } | Select-Object -First 1)
    Add-TKNote -Text ("GARM {0} for {1} (last {2} h): {3}{4}" -f $Mode.ToLower(), $subject, $Hours, $verdict.Verdict,
        $(if ($top) { " -- $($top[0].Detail)" } else { '' })) -Category $(if ($verdict.Class -eq 'ok') { 'Info' } else { 'Issue' }) -ScriptName 'garm'

    Write-Step "Generating HTML report..."
    $html    = Build-GarmReport -Result $result -Domain $domain -Verdict $verdict
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("GARM_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'garm' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-GarmBanner
    if ($Identity) { Invoke-GarmRun -Mode 'Trace' -Name $Identity } else { Invoke-GarmRun -Mode 'Sweep' }
} else {
    $choice = ''
    do {
        Show-GarmBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS  (looking back $Hours hour(s))" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Trace an account  -  where its lockouts and bad passwords come from" -ForegroundColor $C.Info
        Write-Host "  [2] Sweep  -  every lockout in the domain, grouped by account" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' {
                $prompt = if ($Identity) { "  Account (sAMAccountName, DOMAIN\name or UPN) [$Identity]" } else { "  Account (sAMAccountName, DOMAIN\name or UPN)" }
                $name = (Read-Host $prompt).Trim()
                if (-not $name) { $name = $Identity }
                if ($name) { Invoke-GarmRun -Mode 'Trace' -Name $name } else { Write-Warn "No account entered." }
            }
            '2' { Invoke-GarmRun -Mode 'Sweep' }
            'Q' { Write-Host ""; Write-Host "  Closing G.A.R.M." -ForegroundColor $C.Header; Write-Host "" }
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
