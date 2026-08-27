# herald.ps1 - H.E.R.A.L.D. — Hierarchy, Entitlements, Roles & Access-Level Directory
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
    H.E.R.A.L.D. — Hierarchy, Entitlements, Roles & Access-Level Directory
    Active Directory account roster & access-level report for PowerShell 5.1+

.DESCRIPTION
    Produces a customer-facing roster of on-premises Active Directory user accounts and
    the level of access each one holds — the "who has what" document an MSP hands to a
    client for review and cleanup.

    Every enabled user account is listed as Full Name / alias (SamAccountName) / Role,
    where Role is derived from *effective* (nested) security-group membership rather than
    direct membership alone. Membership is expanded server-side with the LDAP in-chain
    matching rule (1.2.840.113556.1.4.1941), so a user who is a Domain Admin three groups
    deep is still reported as a Domain Administrator. Primary-group membership — which
    nested-membership expansion does not cover — is resolved separately.

    Roles:
      Domain Administrator     Enterprise / Schema / Domain Admins, BUILTIN\Administrators,
                               Group Policy Creator Owners — full or near-full control.
      Delegated Administrator  Built-in operator groups (Account / Server / Print / Backup
                               Operators, DnsAdmins, Key Admins, Remote Management Users) —
                               scoped rights that still lead to domain compromise.
      Elevated (Custom Group)  Member of a customer-created group whose name matches
                               -AdminGroupPattern (e.g. "IT Admins", "Helpdesk Operators").
      Standard User            No privileged membership found.

    Read-only. HERALD queries the directory and writes a report; it never modifies AD.

.USAGE
    PS C:\> .\herald.ps1                             # Interactive — enabled accounts, HTML + CSV
    PS C:\> .\herald.ps1 -Unattended                 # Silent export, no prompts
    PS C:\> .\herald.ps1 -IncludeDisabled            # Include disabled accounts in the roster
    PS C:\> .\herald.ps1 -SearchBase 'OU=Staff,DC=contoso,DC=com'
    PS C:\> .\herald.ps1 -Server dc01.contoso.com -StaleDays 60

.NOTES
    Version : 3.7

#>

param(
    [switch]$Unattended,
    [switch]$IncludeDisabled,
    [ValidateRange(1, 3650)]
    [int]$StaleDays = 90,
    [string]$SearchBase,
    [string]$Server,
    [string]$AdminGroupPattern = '(?i)(admin|operator|helpdesk|privileg)',
    [switch]$SkipCustomGroupScan,
    [string]$OutputPath,
    [switch]$NoCsv,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

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
$C = $ColorSchema

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE DATA
#
# Privileged groups are resolved by RID, not by name, so a renamed or localised
# "Domain Admins" is still found. Domain-scope RIDs are appended to the domain
# SID; Builtin-scope RIDs hang off the well-known S-1-5-32 authority. DnsAdmins
# is created by the DNS Server role and has no fixed RID, so it resolves by name.
# ─────────────────────────────────────────────────────────────────────────────

$PrivilegedGroupTiers = [ordered]@{
    'Enterprise Admins'           = @{ Rid = 519;   Scope = 'Domain';  Role = 'Domain Administrator';    Reason = 'Full administrative control of every domain in the forest' }
    'Schema Admins'               = @{ Rid = 518;   Scope = 'Domain';  Role = 'Domain Administrator';    Reason = 'Can modify the Active Directory schema forest-wide' }
    'Domain Admins'               = @{ Rid = 512;   Scope = 'Domain';  Role = 'Domain Administrator';    Reason = 'Full administrative control of the domain and every domain-joined machine' }
    'Administrators'              = @{ Rid = 544;   Scope = 'Builtin'; Role = 'Domain Administrator';    Reason = 'Full administrative control of the domain controllers' }
    'Group Policy Creator Owners' = @{ Rid = 520;   Scope = 'Domain';  Role = 'Domain Administrator';    Reason = 'Can author Group Policy, which executes on every machine it is linked to' }
    'Account Operators'           = @{ Rid = 548;   Scope = 'Builtin'; Role = 'Delegated Administrator'; Reason = 'Can create, modify and delete most user and group accounts' }
    'Server Operators'            = @{ Rid = 549;   Scope = 'Builtin'; Role = 'Delegated Administrator'; Reason = 'Can sign in to, service and shut down domain controllers' }
    'Backup Operators'            = @{ Rid = 551;   Scope = 'Builtin'; Role = 'Delegated Administrator'; Reason = 'Can read and restore any file on a DC, bypassing file permissions' }
    'Print Operators'             = @{ Rid = 550;   Scope = 'Builtin'; Role = 'Delegated Administrator'; Reason = 'Can load printer drivers onto domain controllers' }
    'Remote Management Users'     = @{ Rid = 580;   Scope = 'Builtin'; Role = 'Delegated Administrator'; Reason = 'Can connect to machines over WinRM / PowerShell Remoting' }
    'Key Admins'                  = @{ Rid = 526;   Scope = 'Domain';  Role = 'Delegated Administrator'; Reason = 'Can write key credentials on domain objects' }
    'Enterprise Key Admins'       = @{ Rid = 527;   Scope = 'Domain';  Role = 'Delegated Administrator'; Reason = 'Can write key credentials forest-wide' }
    'Cert Publishers'             = @{ Rid = 517;   Scope = 'Domain';  Role = 'Delegated Administrator'; Reason = 'Can publish certificates to the directory' }
    'DnsAdmins'                   = @{ Rid = $null; Scope = 'Name';    Role = 'Delegated Administrator'; Reason = 'Can load an arbitrary DLL into the DNS service on a domain controller' }
}

# Role tiers, most privileged first. Rank drives sorting and "highest wins"
# classification; Badge maps to the shared CSS badge classes.
$RoleTiers = [ordered]@{
    'Domain Administrator'    = @{ Rank = 1; Badge = 'err';  Blurb = 'Unrestricted or near-unrestricted control of the domain. Every account here should be a named, justified administrator.' }
    'Delegated Administrator' = @{ Rank = 2; Badge = 'warn'; Blurb = 'Scoped administrative rights through a built-in operator group. Several of these paths lead to full domain control.' }
    'Elevated (Custom Group)' = @{ Rank = 3; Badge = 'warn'; Blurb = 'Member of a customer-created group whose name suggests elevated rights. Confirm what the group actually grants.' }
    'Standard User'           = @{ Rank = 4; Badge = 'info'; Blurb = 'No privileged group membership found. Ordinary day-to-day account.' }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-HeraldBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗  ██╗███████╗██████╗  █████╗ ██╗     ██████╗
  ██║  ██║██╔════╝██╔══██╗██╔══██╗██║     ██╔══██╗
  ███████║█████╗  ██████╔╝███████║██║     ██║  ██║
  ██╔══██║██╔══╝  ██╔══██╗██╔══██║██║     ██║  ██║
  ██║  ██║███████╗██║  ██║██║  ██║███████╗██████╔╝
  ╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚═════╝

"@ -ForegroundColor Cyan
    Write-Host "    H.E.R.A.L.D. — Hierarchy, Entitlements, Roles & Access-Level Directory" -ForegroundColor Cyan
    Write-Host "    Active Directory Account Roster & Access-Level Report" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SMALL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-LdapFilterValue {
    <#
        Escapes the characters RFC 4515 reserves inside an LDAP filter assertion.
        Distinguished names routinely contain parentheses and backslashes, and an
        unescaped one silently changes the meaning of the filter.
    #>
    param([string]$Value)

    if ([string]::IsNullOrEmpty($Value)) { return '' }

    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '\'     { [void]$sb.Append('\5c') }
            '*'     { [void]$sb.Append('\2a') }
            '('     { [void]$sb.Append('\28') }
            ')'     { [void]$sb.Append('\29') }
            "`0"    { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch)   }
        }
    }
    return $sb.ToString()
}

function Get-DnLeaf {
    <# Returns the first RDN value of a distinguished name (e.g. the manager's CN). #>
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return '' }
    $first = ($DistinguishedName -split '(?<!\\),')[0]
    return ($first -replace '^[A-Za-z]+=', '') -replace '\\(.)', '$1'
}

function Get-DnParent {
    <# Returns the container a distinguished name sits in — the OU path, for the report. #>
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return '' }
    $parts = $DistinguishedName -split '(?<!\\),'
    if ($parts.Count -le 1) { return '' }
    return ($parts[1..($parts.Count - 1)] -join ',')
}

function Format-HeraldDate {
    param($Value)
    if (-not $Value) { return 'Never' }
    return ([datetime]$Value).ToString('yyyy-MM-dd')
}

function Get-DaysSince {
    param($Value)
    if (-not $Value) { return $null }
    return [int]((Get-Date) - [datetime]$Value).TotalDays
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE & DOMAIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

function Assert-HeraldADModule {
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
    Write-Host "  to read the directory." -ForegroundColor $C.Info
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

function Test-DomainJoined {
    if ($Server) { return $true }
    if ($env:USERDNSDOMAIN) { return $true }
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        return ($cs.PartOfDomain -eq $true)
    } catch {
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GROUP RESOLUTION & MEMBERSHIP EXPANSION
# ─────────────────────────────────────────────────────────────────────────────

function Resolve-PrivilegedGroup {
    <#
        Resolves one entry of $PrivilegedGroupTiers to a real AD group object.
        Returns $null when the group does not exist in this domain — Enterprise
        and Schema Admins live only in the forest root, and Key Admins only on
        2016+ schemas, so a missing group is normal, not an error.
    #>
    param(
        [string]   $Label,
        [hashtable]$Spec,
        [string]   $DomainSid,
        [hashtable]$AdCommon
    )

    try {
        switch ($Spec.Scope) {
            'Builtin' { return Get-ADGroup -Identity ('S-1-5-32-{0}' -f $Spec.Rid) @AdCommon -ErrorAction Stop }
            'Domain'  { return Get-ADGroup -Identity ('{0}-{1}' -f $DomainSid, $Spec.Rid) @AdCommon -ErrorAction Stop }
            default   {
                $escaped = ConvertTo-LdapFilterValue $Label
                return Get-ADGroup -LDAPFilter "(sAMAccountName=$escaped)" @AdCommon -ErrorAction Stop |
                       Select-Object -First 1
            }
        }
    } catch {
        return $null
    }
}

function Get-EffectiveMemberSam {
    <#
        Returns the SamAccountNames of every user with an effective membership in
        $Group, nesting included. The 1.2.840.113556.1.4.1941 matching rule makes
        the DC walk the chain, which is both faster and more reliable than pulling
        members client-side and recursing by hand. Primary-group membership is not
        covered by the rule and is resolved separately by the caller.
    #>
    param($Group, [hashtable]$AdCommon)

    $dn     = ConvertTo-LdapFilterValue $Group.DistinguishedName
    $filter = "(&(objectCategory=person)(objectClass=user)(memberOf:1.2.840.113556.1.4.1941:=$dn))"

    try {
        return @(Get-ADUser -LDAPFilter $filter @AdCommon -ErrorAction Stop |
                 Select-Object -ExpandProperty SamAccountName)
    } catch {
        Write-Warn "Nested expansion failed for '$($Group.Name)' — falling back to direct members: $($_.Exception.Message)"
        try {
            return @(Get-ADGroupMember -Identity $Group.DistinguishedName @AdCommon -ErrorAction Stop |
                     Where-Object { $_.objectClass -eq 'user' } |
                     Select-Object -ExpandProperty SamAccountName)
        } catch {
            Write-Warn "Could not read members of '$($Group.Name)': $($_.Exception.Message)"
            return @()
        }
    }
}

function Get-CustomAdminGroup {
    <#
        Finds customer-created groups whose name matches -AdminGroupPattern. These
        are the "IT Admins" / "Helpdesk Operators" groups that carry real delegated
        rights but are invisible to a built-in-groups-only audit. Built-in groups
        already covered by $PrivilegedGroupTiers are excluded so a user is not
        reported twice for the same grant.
    #>
    param([string[]]$ExcludeDns, [hashtable]$AdCommon, [int]$Limit = 60)

    try {
        $all = @(Get-ADGroup -Filter * -Properties Name, DistinguishedName @AdCommon -ErrorAction Stop)
    } catch {
        Write-Warn "Could not enumerate groups for the custom-group scan: $($_.Exception.Message)"
        return @()
    }

    $matched = @($all | Where-Object {
        $_.Name -match $AdminGroupPattern -and $ExcludeDns -notcontains $_.DistinguishedName
    } | Sort-Object Name)

    if ($matched.Count -gt $Limit) {
        Write-Warn "$($matched.Count) group names matched -AdminGroupPattern; scanning the first $Limit only."
        Write-Info 'Narrow -AdminGroupPattern to bring the whole set into the report.'
        $matched = $matched[0..($Limit - 1)]
    }

    return $matched
}

# ─────────────────────────────────────────────────────────────────────────────
# ACCOUNT COLLECTION & CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Get-HeraldUser {
    param([hashtable]$AdCommon)

    $props = @(
        'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'UserPrincipalName',
        'EmailAddress', 'Enabled', 'Title', 'Department', 'Manager', 'LastLogonDate',
        'PasswordLastSet', 'PasswordNeverExpires', 'PasswordExpired', 'whenCreated',
        'DistinguishedName', 'ServicePrincipalName', 'PrimaryGroupID', 'adminCount',
        'Description', 'LockedOut', 'TrustedForDelegation'
    )

    $query = @{ Properties = $props } + $AdCommon
    if ($SearchBase) { $query['SearchBase'] = $SearchBase }

    if ($IncludeDisabled) {
        return @(Get-ADUser -Filter * @query -ErrorAction Stop)
    }
    return @(Get-ADUser -Filter 'Enabled -eq $true' @query -ErrorAction Stop)
}

function Get-AccountType {
    <#
        Separates the "is this a person?" question from the "what can they do?"
        question. Role reports privilege; Type reports what kind of object holds it.
    #>
    param($User)

    if ($User.SamAccountName -eq 'krbtgt')                    { return 'Built-in (KDC)'      }
    if ($User.SamAccountName -eq 'Guest')                      { return 'Built-in (Guest)'    }
    if (@($User.ServicePrincipalName).Count -gt 0)             { return 'Service Account'     }
    if ($User.SamAccountName -match '^(svc|service|sa)[-_.]')  { return 'Service Account'     }
    return 'User'
}

function Get-ReviewFlag {
    <#
        The cleanup half of the report: everything a customer should be asked about
        when they read the roster.
    #>
    param($User, [string]$Role, [int]$DaysInactive, [bool]$HasCurrentPrivilege)

    $flags = New-Object System.Collections.Generic.List[string]

    if (-not $User.Enabled) {
        $flags.Add('Disabled')
    } elseif ($null -eq $User.LastLogonDate) {
        $flags.Add('Never signed in')
    } elseif ($DaysInactive -ge $StaleDays) {
        $flags.Add("Inactive $DaysInactive days")
    }

    if ($User.LockedOut)            { $flags.Add('Locked out') }
    if ($User.PasswordExpired)      { $flags.Add('Password expired') }
    if ($User.PasswordNeverExpires) { $flags.Add('Password never expires') }
    if ($null -eq $User.PasswordLastSet) { $flags.Add('Password never set') }
    if ($User.TrustedForDelegation) { $flags.Add('Trusted for unconstrained delegation') }

    # adminCount is stamped by AdminSDHolder when an account joins a protected
    # group and is never cleared when it leaves. An account carrying the stamp
    # with no current privileged membership is a former administrator whose ACL
    # is still detached from its OU — worth raising during a cleanup.
    if ($User.adminCount -eq 1 -and -not $HasCurrentPrivilege) {
        $flags.Add('Former privileged account (adminCount set)')
    }

    if ($Role -eq 'Domain Administrator' -and $DaysInactive -ge $StaleDays) {
        $flags.Add('Unused administrator')
    }

    return $flags
}

function Build-AccountRoster {
    param(
        [array]    $Users,
        [hashtable]$GrantIndex,
        [hashtable]$GrantDetail
    )

    $roster = New-Object System.Collections.Generic.List[object]

    foreach ($u in $Users) {
        $key    = $u.SamAccountName.ToLowerInvariant()
        $grants = @()
        if ($GrantIndex.ContainsKey($key)) { $grants = @($GrantIndex[$key]) }

        # Highest privilege wins: the lowest rank among every group granting access.
        $role = 'Standard User'
        $rank = $RoleTiers['Standard User'].Rank
        foreach ($g in $grants) {
            $gRole = $GrantDetail[$g].Role
            $gRank = $RoleTiers[$gRole].Rank
            if ($gRank -lt $rank) {
                $rank = $gRank
                $role = $gRole
            }
        }

        $daysInactive = Get-DaysSince $u.LastLogonDate
        $daysValue    = if ($null -eq $daysInactive) { [int]::MaxValue } else { $daysInactive }

        $flags = Get-ReviewFlag -User $u -Role $role -DaysInactive $daysValue `
                                -HasCurrentPrivilege ($grants.Count -gt 0)

        $fullName = $u.DisplayName
        if ([string]::IsNullOrWhiteSpace($fullName)) {
            $fullName = (($u.GivenName, $u.Surname) | Where-Object { $_ }) -join ' '
        }
        if ([string]::IsNullOrWhiteSpace($fullName)) { $fullName = $u.Name }
        if ([string]::IsNullOrWhiteSpace($fullName)) { $fullName = $u.SamAccountName }

        $roster.Add([PSCustomObject]@{
            FullName             = $fullName
            Alias                = $u.SamAccountName
            Role                 = $role
            Rank                 = $rank
            AccountType          = Get-AccountType -User $u
            Enabled              = [bool]$u.Enabled
            UserPrincipalName    = $u.UserPrincipalName
            EmailAddress         = $u.EmailAddress
            Title                = $u.Title
            Department           = $u.Department
            Manager              = Get-DnLeaf $u.Manager
            GrantedBy            = ($grants -join '; ')
            LastLogon            = Format-HeraldDate $u.LastLogonDate
            DaysInactive         = $daysInactive
            PasswordLastSet      = Format-HeraldDate $u.PasswordLastSet
            PasswordNeverExpires = [bool]$u.PasswordNeverExpires
            AdminCount           = ($u.adminCount -eq 1)
            OrganizationalUnit   = Get-DnParent $u.DistinguishedName
            Description          = $u.Description
            ReviewFlags          = ($flags -join '; ')
        })
    }

    return $roster
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Get-RoleBadge {
    param([string]$Role)
    $badge = 'info'
    if ($RoleTiers.Contains($Role)) { $badge = $RoleTiers[$Role].Badge }
    return "<span class=`"tk-badge-$badge`">$(EscHtml $Role)</span>"
}

function Get-FlagCell {
    param([string]$Flags)
    if ([string]::IsNullOrWhiteSpace($Flags)) { return '<span class="tk-badge-ok">Clear</span>' }
    $cells = foreach ($f in ($Flags -split ';\s*')) {
        "<span class=`"tk-badge-warn`">$(EscHtml $f)</span>"
    }
    return ($cells -join ' ')
}

function Build-HeraldReport {
    param(
        [array]    $Roster,
        [array]    $GroupSummary,
        [string]   $DomainName,
        [string]   $ReportTimestamp,
        [hashtable]$Counts
    )

    $tkConfig  = Get-TKConfig
    $orgPrefix = ''
    if (-not [string]::IsNullOrWhiteSpace($tkConfig.OrgName)) { $orgPrefix = "$($tkConfig.OrgName) -- " }

    $scopeLabel = 'Enabled accounts'
    if ($IncludeDisabled) { $scopeLabel = 'Enabled and disabled accounts' }

    $metaItems = [ordered]@{
        'Domain'    = $DomainName
        'Generated' = $ReportTimestamp
        'Scope'     = $scopeLabel
        'Accounts'  = $Roster.Count
        'Inactive'  = "$StaleDays+ days"
    }

    $navItems = @(
        'Access Levels at a Glance',
        'Privileged Accounts',
        'Full Account Roster',
        'Privileged Group Membership',
        'Flagged for Review'
    )

    # ── 01 Access levels at a glance ─────────────────────────────────────────
    $tierRows = ''
    foreach ($tier in $RoleTiers.Keys) {
        $count = @($Roster | Where-Object { $_.Role -eq $tier }).Count
        $tierRows += @"
        <tr>
          <td>$(Get-RoleBadge $tier)</td>
          <td><strong>$count</strong></td>
          <td>$(EscHtml $RoleTiers[$tier].Blurb)</td>
        </tr>
"@
    }

    $typeRows = ''
    foreach ($grp in ($Roster | Group-Object AccountType | Sort-Object Name)) {
        $typeRows += @"
        <tr><td>$(EscHtml $grp.Name)</td><td><strong>$($grp.Count)</strong></td></tr>
"@
    }

    # ── 02 Privileged accounts ───────────────────────────────────────────────
    $privileged = @($Roster | Where-Object { $_.Rank -lt $RoleTiers['Standard User'].Rank } |
                    Sort-Object Rank, FullName)

    # A whole-domain roster can run to thousands of rows, and `+=` on a string
    # reallocates the entire buffer once per row, so the three account-scale
    # loops accumulate into a StringBuilder instead.
    $privSb = New-Object System.Text.StringBuilder
    foreach ($a in $privileged) {
        [void]$privSb.Append(@"
        <tr>
          <td><strong>$(EscHtml $a.FullName)</strong></td>
          <td class="tk-mono">$(EscHtml $a.Alias)</td>
          <td>$(Get-RoleBadge $a.Role)</td>
          <td>$(EscHtml $a.GrantedBy)</td>
          <td>$(EscHtml $a.AccountType)</td>
          <td>$(EscHtml $a.LastLogon)</td>
          <td>$(Get-FlagCell $a.ReviewFlags)</td>
        </tr>
"@)
    }
    $privRows = $privSb.ToString()
    if (-not $privRows) {
        $privRows = '<tr><td colspan="7">No accounts hold privileged group membership.</td></tr>'
    }

    # ── 03 Full roster ───────────────────────────────────────────────────────
    $rosterSb = New-Object System.Text.StringBuilder
    foreach ($a in ($Roster | Sort-Object Rank, FullName)) {
        $statusBadge = if ($a.Enabled) {
            '<span class="tk-badge-ok">Active</span>'
        } else {
            '<span class="tk-badge-warn">Disabled</span>'
        }
        [void]$rosterSb.Append(@"
        <tr>
          <td><strong>$(EscHtml $a.FullName)</strong></td>
          <td class="tk-mono">$(EscHtml $a.Alias)</td>
          <td>$(Get-RoleBadge $a.Role)</td>
          <td>$(EscHtml $a.AccountType)</td>
          <td>$statusBadge</td>
          <td>$(EscHtml $a.Title)</td>
          <td>$(EscHtml $a.Department)</td>
          <td>$(EscHtml $a.LastLogon)</td>
          <td>$(Get-FlagCell $a.ReviewFlags)</td>
        </tr>
"@)
    }
    $rosterRows = $rosterSb.ToString()
    if (-not $rosterRows) {
        $rosterRows = '<tr><td colspan="9">No accounts matched the report scope.</td></tr>'
    }

    # ── 04 Privileged group membership ───────────────────────────────────────
    $groupRows = ''
    foreach ($g in $GroupSummary) {
        $memberCell = if ($g.Members.Count -gt 0) { EscHtml (($g.Members | Sort-Object) -join ', ') } else { '<em>Empty</em>' }
        $countClass = if ($g.Members.Count -eq 0) { 'ok' } else { $RoleTiers[$g.Role].Badge }
        $groupRows += @"
        <tr>
          <td><strong>$(EscHtml $g.Name)</strong></td>
          <td>$(Get-RoleBadge $g.Role)</td>
          <td><span class="tk-badge-$countClass">$($g.Members.Count)</span></td>
          <td>$(EscHtml $g.Reason)</td>
          <td>$memberCell</td>
        </tr>
"@
    }
    if (-not $groupRows) {
        $groupRows = '<tr><td colspan="5">No privileged groups were resolved in this domain.</td></tr>'
    }

    # ── 05 Flagged for review ────────────────────────────────────────────────
    $flagged = @($Roster | Where-Object { $_.ReviewFlags } | Sort-Object Rank, FullName)
    $flagSb = New-Object System.Text.StringBuilder
    foreach ($a in $flagged) {
        [void]$flagSb.Append(@"
        <tr>
          <td><strong>$(EscHtml $a.FullName)</strong></td>
          <td class="tk-mono">$(EscHtml $a.Alias)</td>
          <td>$(Get-RoleBadge $a.Role)</td>
          <td>$(EscHtml $a.LastLogon)</td>
          <td>$(EscHtml $a.OrganizationalUnit)</td>
          <td>$(Get-FlagCell $a.ReviewFlags)</td>
        </tr>
"@)
    }
    $flagRows = $flagSb.ToString()
    if (-not $flagRows) {
        $flagRows = '<tr><td colspan="6">No accounts raised a review flag.</td></tr>'
    }

    $adminClass   = if ($Counts.DomainAdmin -gt 4) { 'err' } else { 'warn' }
    $flaggedClass = if ($flagged.Count -gt 0) { 'warn' } else { 'ok' }

    $html = (Get-TKHtmlHead `
        -Title      'Active Directory Account Roster & Access Levels' `
        -ScriptName 'H.E.R.A.L.D.' `
        -Subtitle   "$orgPrefix$DomainName" `
        -MetaItems  $metaItems `
        -NavItems   $navItems) + @"

<div class="tk-info-box">
  <span class="tk-info-label">How to read this report</span>
  Every account below is listed as <strong>Full Name</strong> / <strong>alias</strong> / <strong>Role</strong>.
  Role is the highest level of access the account holds through its <em>effective</em> security-group
  membership — nested groups included — not just the groups it is directly a member of.
  Hand sections 03 and 05 to the customer: section 03 is the roster to confirm, section 05 is the
  shortlist of accounts that look like cleanup candidates.
</div>

<div class="tk-summary-row">
  <div class="tk-summary-card info">
    <div class="tk-summary-num">$($Roster.Count)</div>
    <div class="tk-summary-lbl">Accounts</div>
  </div>
  <div class="tk-summary-card $adminClass">
    <div class="tk-summary-num">$($Counts.DomainAdmin)</div>
    <div class="tk-summary-lbl">Domain Administrators</div>
  </div>
  <div class="tk-summary-card warn">
    <div class="tk-summary-num">$($Counts.Delegated)</div>
    <div class="tk-summary-lbl">Delegated Admins</div>
  </div>
  <div class="tk-summary-card warn">
    <div class="tk-summary-num">$($Counts.Elevated)</div>
    <div class="tk-summary-lbl">Elevated (Custom)</div>
  </div>
  <div class="tk-summary-card ok">
    <div class="tk-summary-num">$($Counts.Standard)</div>
    <div class="tk-summary-lbl">Standard Users</div>
  </div>
  <div class="tk-summary-card $flaggedClass">
    <div class="tk-summary-num">$($flagged.Count)</div>
    <div class="tk-summary-lbl">Flagged for Review</div>
  </div>
</div>

<div class="tk-section" id="s01">
  <div class="tk-card-header">
    <span class="tk-section-title">Access Levels at a Glance</span>
    <span class="tk-section-num">Section 01</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Role</th><th>Accounts</th><th>What it means</th></tr></thead>
      <tbody>$tierRows</tbody>
    </table>
    <p class="tk-card-label" style="margin-top:22px">Account types</p>
    <table class="tk-table">
      <thead><tr><th>Type</th><th>Accounts</th></tr></thead>
      <tbody>$typeRows</tbody>
    </table>
  </div>
</div>

<div class="tk-section" id="s02">
  <div class="tk-card-header">
    <span class="tk-section-title">Privileged Accounts</span>
    <span class="tk-section-num">$($privileged.Count) account(s)</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead>
        <tr>
          <th>Full Name</th><th>Alias</th><th>Role</th><th>Granted By</th>
          <th>Type</th><th>Last Sign-in</th><th>Review Flags</th>
        </tr>
      </thead>
      <tbody>$privRows</tbody>
    </table>
    <div class="tk-info-box" style="margin-top:18px">
      <span class="tk-info-label">Note</span>
      "Granted By" names every privileged group the account reaches, directly or through nesting.
      A service account in this list is worth particular attention — its password rarely changes
      and it is rarely covered by MFA.
    </div>
  </div>
</div>

<div class="tk-section" id="s03">
  <div class="tk-card-header">
    <span class="tk-section-title">Full Account Roster</span>
    <span class="tk-section-num">$($Roster.Count) account(s)</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead>
        <tr>
          <th>Full Name</th><th>Alias</th><th>Role</th><th>Type</th><th>Status</th>
          <th>Title</th><th>Department</th><th>Last Sign-in</th><th>Review Flags</th>
        </tr>
      </thead>
      <tbody>$rosterRows</tbody>
    </table>
  </div>
</div>

<div class="tk-section" id="s04">
  <div class="tk-card-header">
    <span class="tk-section-title">Privileged Group Membership</span>
    <span class="tk-section-num">$($GroupSummary.Count) group(s)</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Group</th><th>Role Conferred</th><th>Members</th><th>Why it matters</th><th>Effective members</th></tr></thead>
      <tbody>$groupRows</tbody>
    </table>
  </div>
</div>

<div class="tk-section" id="s05">
  <div class="tk-card-header">
    <span class="tk-section-title">Flagged for Review</span>
    <span class="tk-section-num">$($flagged.Count) account(s)</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Full Name</th><th>Alias</th><th>Role</th><th>Last Sign-in</th><th>Organizational Unit</th><th>Review Flags</th></tr></thead>
      <tbody>$flagRows</tbody>
    </table>
    <div class="tk-info-box" style="margin-top:18px">
      <span class="tk-info-label">Scope &amp; caveats</span>
      HERALD reports privilege conferred by <em>security-group membership</em> in this domain.
      It does not evaluate rights delegated directly on an OU or object ACL, local administrator
      membership on individual workstations, Group Policy user-rights assignments, or group
      Managed Service Accounts (gMSAs), which are not user objects.
      Entra ID / Microsoft 365 directory roles are a separate surface — run W.R.A.I.T.H. for those.
      Local accounts on a single machine are covered by W.A.R.D.
    </div>
  </div>
</div>

"@ + (Get-TKHtmlFoot -ScriptName 'H.E.R.A.L.D. v3.7')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-HeraldBanner }

# A bad -AdminGroupPattern would otherwise surface as a confusing mid-scan
# exception, so it is validated before any directory work starts.
if (-not $SkipCustomGroupScan) {
    try {
        $null = 'probe' -match $AdminGroupPattern
    } catch {
        Write-Fail "-AdminGroupPattern is not a valid regular expression: $($_.Exception.Message)"
        if ($Transcript) { Stop-TKTranscript }
        exit 1
    }
}

if (-not (Test-DomainJoined)) {
    Write-Fail 'This machine is not joined to an Active Directory domain.'
    Write-Info 'Run HERALD from a domain-joined machine, or point it at a domain controller with -Server.'
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

if (-not (Assert-HeraldADModule)) {
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

$AdCommon = @{}
if ($Server) { $AdCommon['Server'] = $Server }

Write-Section 'DIRECTORY'

try {
    $domain = Get-ADDomain @AdCommon -ErrorAction Stop
} catch {
    Write-Fail "Could not contact the domain: $($_.Exception.Message)"
    Write-TKError -ScriptName 'herald' -Message $_.Exception.Message -Category 'Directory'
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

$domainName = $domain.DNSRoot
$domainSid  = $domain.DomainSID.Value
Write-Ok "Connected to $domainName"
if ($SearchBase) { Write-Info "Search base: $SearchBase" }

# ── Resolve privileged groups and expand their effective membership ──────────

Write-Section 'PRIVILEGED GROUPS'

$grantIndex   = @{}   # sam (lower-case) -> list of group labels granting access
$grantDetail  = @{}   # group label      -> role + reason
$groupSummary = New-Object System.Collections.Generic.List[object]
$resolvedDns  = New-Object System.Collections.Generic.List[string]
$primaryRid   = @{}   # primaryGroupID   -> group label

function Add-Grant {
    param([string]$Sam, [string]$Label)
    $key = $Sam.ToLowerInvariant()
    if (-not $grantIndex.ContainsKey($key)) {
        $grantIndex[$key] = New-Object System.Collections.Generic.List[string]
    }
    if (-not $grantIndex[$key].Contains($Label)) { $grantIndex[$key].Add($Label) }
}

foreach ($label in $PrivilegedGroupTiers.Keys) {
    $spec  = $PrivilegedGroupTiers[$label]
    $group = Resolve-PrivilegedGroup -Label $label -Spec $spec -DomainSid $domainSid -AdCommon $AdCommon

    if (-not $group) {
        Write-Info "$label — not present in this domain, skipped"
        continue
    }

    $members = @(Get-EffectiveMemberSam -Group $group -AdCommon $AdCommon)
    foreach ($m in $members) { Add-Grant -Sam $m -Label $label }

    $grantDetail[$label] = @{ Role = $spec.Role; Reason = $spec.Reason }
    $resolvedDns.Add($group.DistinguishedName)
    if ($spec.Scope -eq 'Domain') { $primaryRid[[int]$spec.Rid] = $label }

    $groupSummary.Add([PSCustomObject]@{
        Name    = $group.Name
        Label   = $label
        Role    = $spec.Role
        Reason  = $spec.Reason
        Members = $members
    })

    Write-Ok ("{0,-30} {1} effective member(s)" -f $group.Name, $members.Count)
}

# ── Customer-created groups that look administrative ─────────────────────────

if (-not $SkipCustomGroupScan) {
    Write-Section 'CUSTOM ADMIN-LIKE GROUPS'
    Write-Step "Scanning group names against pattern: $AdminGroupPattern"

    $customGroups = @(Get-CustomAdminGroup -ExcludeDns $resolvedDns -AdCommon $AdCommon)

    if ($customGroups.Count -eq 0) {
        Write-Info 'No customer-created group names matched.'
    }

    foreach ($cg in $customGroups) {
        $label = $cg.Name
        if ($grantDetail.ContainsKey($label)) { continue }

        $members = @(Get-EffectiveMemberSam -Group $cg -AdCommon $AdCommon)
        if ($members.Count -eq 0) { continue }

        foreach ($m in $members) { Add-Grant -Sam $m -Label $label }

        $grantDetail[$label] = @{
            Role   = 'Elevated (Custom Group)'
            Reason = 'Customer-created group whose name matches the administrative naming pattern'
        }
        $groupSummary.Add([PSCustomObject]@{
            Name    = $cg.Name
            Label   = $label
            Role    = 'Elevated (Custom Group)'
            Reason  = 'Customer-created group whose name matches the administrative naming pattern'
            Members = $members
        })

        Write-Ok ("{0,-30} {1} effective member(s)" -f $cg.Name, $members.Count)
    }
}

# ── Accounts ─────────────────────────────────────────────────────────────────

Write-Section 'ACCOUNTS'
Write-Step 'Reading user accounts from the directory...'

try {
    $users = Get-HeraldUser -AdCommon $AdCommon
} catch {
    Write-Fail "Could not read user accounts: $($_.Exception.Message)"
    Write-TKError -ScriptName 'herald' -Message $_.Exception.Message -Category 'Directory'
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

Write-Ok "$($users.Count) account(s) returned."

# Primary-group membership is stored on the user, not in the group's member
# attribute, so the in-chain expansion above cannot see it. An account whose
# primary group has been switched to Domain Admins is a real (and deliberately
# quiet) privilege grant, so it is folded in here.
$primaryGrants = 0
foreach ($u in $users) {
    if ($null -eq $u.PrimaryGroupID) { continue }
    $rid = [int]$u.PrimaryGroupID
    if ($primaryRid.ContainsKey($rid)) {
        Add-Grant -Sam $u.SamAccountName -Label $primaryRid[$rid]
        $primaryGrants++
    }
}
if ($primaryGrants -gt 0) {
    Write-Warn "$primaryGrants account(s) hold privilege through their primary group."
}

$roster = @(Build-AccountRoster -Users $users -GrantIndex $grantIndex -GrantDetail $grantDetail)

$counts = @{
    DomainAdmin = @($roster | Where-Object { $_.Role -eq 'Domain Administrator'    }).Count
    Delegated   = @($roster | Where-Object { $_.Role -eq 'Delegated Administrator' }).Count
    Elevated    = @($roster | Where-Object { $_.Role -eq 'Elevated (Custom Group)' }).Count
    Standard    = @($roster | Where-Object { $_.Role -eq 'Standard User'           }).Count
    Flagged     = @($roster | Where-Object { $_.ReviewFlags }).Count
}

# ── Console summary ──────────────────────────────────────────────────────────

Write-Section 'ACCESS LEVEL SUMMARY'

Write-Host ("  {0,-26} {1}" -f 'Domain Administrators',   $counts.DomainAdmin) -ForegroundColor $C.Error
Write-Host ("  {0,-26} {1}" -f 'Delegated Administrators', $counts.Delegated)  -ForegroundColor $C.Warning
Write-Host ("  {0,-26} {1}" -f 'Elevated (custom group)',  $counts.Elevated)   -ForegroundColor $C.Warning
Write-Host ("  {0,-26} {1}" -f 'Standard Users',           $counts.Standard)   -ForegroundColor $C.Success
Write-Host ("  {0,-26} {1}" -f 'Flagged for review',       $counts.Flagged)    -ForegroundColor $C.Info
Write-Host ""

$privilegedAccounts = @($roster | Where-Object { $_.Rank -lt $RoleTiers['Standard User'].Rank } |
                        Sort-Object Rank, FullName)

if ($privilegedAccounts.Count -gt 0) {
    Write-Section 'PRIVILEGED ACCOUNTS'
    Write-Host ("  {0,-30} {1,-22} {2}" -f 'Full Name', 'Alias', 'Role') -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 82)) -ForegroundColor $C.Header
    foreach ($a in $privilegedAccounts) {
        $line = "  {0,-30} {1,-22} {2}" -f $a.FullName, $a.Alias, $a.Role
        $tone = if ($a.Role -eq 'Domain Administrator') { $C.Error } else { $C.Warning }
        Write-Host $line -ForegroundColor $tone
        if ($a.ReviewFlags) {
            Write-Host ("  {0,-30} {1}" -f '', $a.ReviewFlags) -ForegroundColor $C.Info
        }
    }
    Write-Host ""
}

# ── Report output ────────────────────────────────────────────────────────────

Write-Section 'REPORT'

$outDir = $OutputPath
if ([string]::IsNullOrWhiteSpace($outDir)) {
    $outDir = Resolve-LogDirectory -FallbackPath $ScriptPath
}
if (-not (Test-Path $outDir)) {
    try   { $null = New-Item -Path $outDir -ItemType Directory -Force -ErrorAction Stop }
    catch { Write-Warn "Could not create $outDir — falling back to the script directory."; $outDir = $ScriptPath }
}

$stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportPath = Join-Path $outDir "HERALD_$stamp.html"
$csvPath    = Join-Path $outDir "HERALD_Roster_$stamp.csv"

try {
    $html = Build-HeraldReport -Roster $roster -GroupSummary @($groupSummary) `
                               -DomainName $domainName `
                               -ReportTimestamp (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') `
                               -Counts $counts
    [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
} catch {
    Write-Fail "Could not save the HTML report: $($_.Exception.Message)"
    Write-TKError -ScriptName 'herald' -Message $_.Exception.Message -Category 'Report'
}

if (-not $NoCsv) {
    # The CSV carries two deliberately empty columns so the customer can mark up
    # each account during review and hand the file straight back.
    try {
        $roster |
            Select-Object FullName, Alias, Role, AccountType, Enabled, UserPrincipalName,
                          EmailAddress, Title, Department, Manager, GrantedBy, LastLogon,
                          DaysInactive, PasswordLastSet, PasswordNeverExpires, AdminCount,
                          OrganizationalUnit, ReviewFlags,
                          @{ Name = 'Action (Keep/Disable/Delete)'; Expression = { '' } },
                          @{ Name = 'Customer Notes';               Expression = { '' } } |
            Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Ok 'Review CSV written (blank Action / Notes columns for the customer to fill in):'
        Write-Info $csvPath
    } catch {
        Write-Fail "Could not save the CSV roster: $($_.Exception.Message)"
    }
}

Show-TKReportResult -Path $reportPath -Unattended:$Unattended -Label 'Account roster report'

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $C.Header
Write-Host "  H.E.R.A.L.D. REPORT COMPLETE" -ForegroundColor $C.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $C.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
