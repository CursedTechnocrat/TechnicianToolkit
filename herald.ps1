# herald.ps1 - H.E.R.A.L.D. — Hierarchy, Entitlements, Roles & Access-Level Directory
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
    Version : 5.0

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

# $RoleTiers stays [ordered] because the report renders the tiers most-privileged
# first (and the Pester suite asserts on the literal). Lookups go through these
# plain copies instead: OrderedDictionary exposes both Item[Int32] and
# Item[Object], and an [object]-typed key reaching that indexer is a known way to
# get "Argument types do not match" out of Windows PowerShell 5.1.
$RoleTierOrder = [string[]]@($RoleTiers.Keys)
$RoleTierMap   = @{}
foreach ($roleKey in $RoleTierOrder) { $RoleTierMap[$roleKey] = $RoleTiers[[string]$roleKey] }

# Domain password / lockout policy baseline. The customer questionnaire asks what
# authentication mechanisms are in place, so each setting is reported with a
# verdict rather than a bare number — a technician answering an audit needs to
# know which values will draw a follow-up question.
#
# Kind drives how a value is judged:
#   Number    higher is better; Strong / Acceptable are the floors
#   Threshold lower is better, but 0 disables the control entirely (always Weak)
#   Age       lower is better, but 0 means "never expires" (called out, not failed)
#   Duration  higher is better, but 0 means "until an administrator unlocks",
#             which is the most restrictive setting rather than the weakest
#   Boolean   Good names the value that is not a finding
$PasswordPolicyBaseline = [ordered]@{
    'MinPasswordLength' = @{
        Label = 'Minimum password length'; Kind = 'Number'; Strong = 14; Acceptable = 8; Unit = 'characters'
        Why   = 'Length is the largest single factor in how long a stolen hash resists offline cracking.'
    }
    'ComplexityEnabled' = @{
        Label = 'Password complexity required'; Kind = 'Boolean'; Good = $true
        Why   = 'Requires three of five character classes and blocks the account name appearing in the password.'
    }
    'PasswordHistoryCount' = @{
        Label = 'Password history'; Kind = 'Number'; Strong = 24; Acceptable = 5; Unit = 'remembered'
        Why   = 'Stops a user returning straight to a known password at the next change.'
    }
    'MaxPasswordAgeDays' = @{
        Label = 'Maximum password age'; Kind = 'Age'; Strong = 90; Acceptable = 365; Unit = 'days'
        Why   = 'Auditors generally expect a bounded lifetime. Note that NIST SP 800-63B now advises against routine expiry where length and breach screening are strong, so a deliberate no-expiry policy may be defensible — say so rather than leaving it unexplained.'
    }
    'MinPasswordAgeDays' = @{
        Label = 'Minimum password age'; Kind = 'Number'; Strong = 1; Acceptable = 1; Unit = 'days'
        Why   = 'At zero a user can cycle through the entire history in one sitting and land back on the same password.'
    }
    'LockoutThreshold' = @{
        Label = 'Account lockout threshold'; Kind = 'Threshold'; Strong = 5; Acceptable = 10; Unit = 'failed attempts'
        Why   = 'At zero, online password guessing against every account in the domain is never interrupted.'
    }
    'LockoutDurationMinutes' = @{
        Label = 'Account lockout duration'; Kind = 'Duration'; Strong = 15; Acceptable = 15; Unit = 'minutes'
        Why   = 'A short lockout lets an attacker resume guessing almost immediately.'
    }
    'LockoutObservationMinutes' = @{
        Label = 'Lockout counter resets after'; Kind = 'Number'; Strong = 15; Acceptable = 15; Unit = 'minutes'
        Why   = 'Resetting the failed-attempt counter quickly widens the window for slow guessing.'
    }
    'ReversibleEncryptionEnabled' = @{
        Label = 'Reversible encryption'; Kind = 'Boolean'; Good = $false
        Why   = 'Stores passwords in a recoverable form - equivalent to plaintext for anyone who can read the directory.'
    }
}

$PolicyKeyOrder = [string[]]@($PasswordPolicyBaseline.Keys)
$PolicyBaseline = @{}
foreach ($policyKey in $PolicyKeyOrder) { $PolicyBaseline[$policyKey] = $PasswordPolicyBaseline[[string]$policyKey] }

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

function ConvertTo-HeraldArray {
    <#
        Normalises any value to [object[]] by explicit enumeration.

        The report was repeatedly lost to System.ArgumentException "Argument
        types do not match" raised while preparing the arguments for
        Build-HeraldReport. The group summary is the one collection built as a
        System.Collections.Generic.List[object], and @() over that list is the
        construct under suspicion - but sandbox testing gave contradictory
        results, so this does not rely on @() behaving. Enumerating item by item
        into an ArrayList cannot raise a conversion error whatever the input is.
    #>
    param($Value)

    $acc = New-Object System.Collections.ArrayList

    if ($null -ne $Value) {
        $isEnumerable = ($Value -is [System.Collections.IEnumerable]) -and
                        ($Value -isnot [string]) -and
                        ($Value -isnot [System.Collections.IDictionary])
        if ($isEnumerable) {
            foreach ($item in $Value) { [void]$acc.Add($item) }
        } else {
            [void]$acc.Add($Value)
        }
    }

    # The unary comma stops PowerShell unrolling the result on output, so an
    # empty collection comes back as an empty array rather than $null.
    return , [object[]]$acc.ToArray()
}

function Get-FaultLocation {
    <#
        Renders "<file> line <n>" for an error record. The file matters because a
        fault can surface from the shared module rather than from this script, and
        naming the wrong one sends the next reader to the wrong place.
    #>
    param($ErrorRecord)

    $line = $ErrorRecord.InvocationInfo.ScriptLineNumber
    $file = $ErrorRecord.InvocationInfo.ScriptName
    if ([string]::IsNullOrWhiteSpace($file)) { return "line $line" }
    return ("{0} line {1}" -f (Split-Path -Leaf $file), $line)
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

function Get-PolicyVerdict {
    <#
        Judges one policy value against its baseline entry. Returns Strong /
        Acceptable / Weak, or Unknown when the directory did not supply a value.
    #>
    param([string]$Key, $Value)

    if (-not $PolicyBaseline.ContainsKey($Key)) { return 'Unknown' }
    if ($null -eq $Value)                        { return 'Unknown' }

    $spec = $PolicyBaseline[$Key]

    switch ($spec.Kind) {
        'Boolean' {
            if ([bool]$Value -eq [bool]$spec.Good) { return 'Strong' }
            return 'Weak'
        }
        'Threshold' {
            # 0 disables lockout altogether, which is worse than any large value.
            if ([int]$Value -eq 0)                  { return 'Weak' }
            if ([int]$Value -le [int]$spec.Strong)     { return 'Strong' }
            if ([int]$Value -le [int]$spec.Acceptable) { return 'Acceptable' }
            return 'Weak'
        }
        'Duration' {
            # 0 keeps the account locked until an administrator intervenes, which
            # is the most restrictive option available, not the weakest.
            if ([int]$Value -eq 0)                     { return 'Strong' }
            if ([int]$Value -ge [int]$spec.Strong)     { return 'Strong' }
            if ([int]$Value -ge [int]$spec.Acceptable) { return 'Acceptable' }
            return 'Weak'
        }
        'Age' {
            # 0 means passwords never expire. Flagged for the questionnaire, but
            # deliberately not called Weak - see the Why text on this entry.
            if ([int]$Value -eq 0)                     { return 'Acceptable' }
            if ([int]$Value -le [int]$spec.Strong)     { return 'Strong' }
            if ([int]$Value -le [int]$spec.Acceptable) { return 'Acceptable' }
            return 'Weak'
        }
        default {
            if ([int]$Value -ge [int]$spec.Strong)     { return 'Strong' }
            if ([int]$Value -ge [int]$spec.Acceptable) { return 'Acceptable' }
            return 'Weak'
        }
    }
}

function Format-PolicyValue {
    <# Renders a policy value the way a person reading an audit response expects. #>
    param([string]$Key, $Value)

    if ($null -eq $Value) { return 'Unknown' }

    $spec = $null
    if ($PolicyBaseline.ContainsKey($Key)) { $spec = $PolicyBaseline[$Key] }

    if ($spec -and $spec.Kind -eq 'Boolean') {
        if ([bool]$Value) { return 'Enabled' }
        return 'Disabled'
    }
    if ($Key -eq 'MaxPasswordAgeDays'   -and [int]$Value -eq 0) { return 'Never expires' }
    if ($Key -eq 'LockoutThreshold'     -and [int]$Value -eq 0) { return 'Never locks out' }
    if ($Key -eq 'LockoutDurationMinutes' -and [int]$Value -eq 0) { return 'Until an administrator unlocks' }

    if ($spec -and $spec.Unit) { return "$Value $($spec.Unit)" }
    return "$Value"
}

function Get-AuthenticationPolicy {
    <#
        Reads the default domain password and lockout policy, plus any
        fine-grained password policies (PSOs). PSOs matter because they override
        the default for the principals they target - answering an access-review
        questionnaire from the default policy alone can be flatly wrong when one
        exists over, say, the admin group.
    #>
    param([hashtable]$AdCommon)

    $result = [PSCustomObject]@{
        Available = $false
        Error     = ''
        Values    = @{}
        Pso       = @()
        PsoError  = ''
    }

    try {
        $p = Get-ADDefaultDomainPasswordPolicy @AdCommon -ErrorAction Stop
    } catch {
        $result.Error = $_.Exception.Message
        return $result
    }

    $result.Available = $true
    $result.Values = @{
        MinPasswordLength           = [int]$p.MinPasswordLength
        ComplexityEnabled           = [bool]$p.ComplexityEnabled
        PasswordHistoryCount        = [int]$p.PasswordHistoryCount
        MaxPasswordAgeDays          = [int]$p.MaxPasswordAge.TotalDays
        MinPasswordAgeDays          = [int]$p.MinPasswordAge.TotalDays
        LockoutThreshold            = [int]$p.LockoutThreshold
        LockoutDurationMinutes      = [int]$p.LockoutDuration.TotalMinutes
        LockoutObservationMinutes   = [int]$p.LockoutObservationWindow.TotalMinutes
        ReversibleEncryptionEnabled = [bool]$p.ReversibleEncryptionEnabled
    }

    try {
        $result.Pso = @(Get-ADFineGrainedPasswordPolicy -Filter * @AdCommon -ErrorAction Stop |
            ForEach-Object {
                [PSCustomObject]@{
                    Name                 = $_.Name
                    Precedence           = $_.Precedence
                    MinPasswordLength    = $_.MinPasswordLength
                    ComplexityEnabled    = $_.ComplexityEnabled
                    PasswordHistoryCount = $_.PasswordHistoryCount
                    MaxPasswordAgeDays   = [int]$_.MaxPasswordAge.TotalDays
                    LockoutThreshold     = $_.LockoutThreshold
                    AppliesTo            = (@($_.AppliesTo | ForEach-Object { Get-DnLeaf $_ }) -join '; ')
                }
            })
    } catch {
        $result.PsoError = $_.Exception.Message
    }

    return $result
}

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
        $rank = $RoleTierMap['Standard User'].Rank
        foreach ($g in $grants) {
            $gRole = $GrantDetail[$g].Role
            $gRank = $RoleTierMap[[string]$gRole].Rank
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
    if ($RoleTierMap.ContainsKey($Role)) { $badge = $RoleTierMap[$Role].Badge }
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

function Invoke-ReportSection {
    <#
        Renders one section's rows inside its own guard. A section that throws is
        reported with its name, exception type and originating line, and is
        replaced in the document by a visible note — one bad row must not cost
        the technician the entire report, and the failure has to say where to look
        rather than just that something went wrong.
    #>
    param(
        [Parameter(Mandatory)][string]      $Name,
        [Parameter(Mandatory)][int]         $ColSpan,
        [Parameter(Mandatory)][scriptblock] $Build
    )

    try {
        return (& $Build)
    } catch {
        $where = Get-FaultLocation $_
        Write-Fail "Report section '$Name' failed to render: $($_.Exception.Message)"
        Write-Info "$($_.Exception.GetType().FullName) at $where"
        Write-Info $_.InvocationInfo.Line.Trim()
        foreach ($frame in ($_.ScriptStackTrace -split "`r?`n")) {
            if ($frame.Trim()) { Write-Info $frame.Trim() }
        }
        Write-TKError -ScriptName 'herald' -Category 'Report' `
            -Message "Section '$Name': $($_.Exception.GetType().FullName): $($_.Exception.Message) [$where]"
        return "<tr><td colspan=""$ColSpan""><strong>This section could not be rendered.</strong> See the console output for the failure detail.</td></tr>"
    }
}

function Build-HeraldReport {
    <#
        Parameters are deliberately untyped.

        A live 91-account domain on Windows PowerShell 5.1 threw
        System.ArgumentException "Argument types do not match" at the *call site*
        of this function - before any section rendered, so none of the section
        guards could catch it and no report was written. A type constraint makes
        the binder convert every argument on the way in, and that conversion is
        the only thing that can fail at a call statement.

        Nothing here needs the coercion: the body enumerates and indexes, and the
        two collections are normalised with @() below, which is what the [array]
        constraint was buying. Removing the constraints removes the failure mode
        outright rather than guessing which of them was at fault.
    #>
    param(
        $Roster,
        $GroupSummary,
        $DomainName,
        $ReportTimestamp,
        $Counts,
        $AuthPolicy
    )

    # Normalise here instead of at the binder, so .Count is still safe on a
    # single-element or empty result.
    $Roster       = ConvertTo-HeraldArray $Roster
    $GroupSummary = ConvertTo-HeraldArray $GroupSummary

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
        'Authentication Policy',
        'Access Levels at a Glance',
        'Privileged Accounts',
        'Full Account Roster',
        'Privileged Group Membership',
        'Flagged for Review'
    )

    # The two account subsets are needed by the summary cards as well as by their
    # own sections, so they are computed once, ahead of rendering.
    $privileged = @($Roster | Where-Object { $_.Rank -lt $RoleTierMap['Standard User'].Rank } |
                    Sort-Object Rank, FullName)
    $flagged    = @($Roster | Where-Object { $_.ReviewFlags } | Sort-Object Rank, FullName)

    # ── 01 Access levels at a glance ─────────────────────────────────────────
    $tierRows = Invoke-ReportSection -Name 'Access levels at a glance' -ColSpan 3 -Build {
        $sb = New-Object System.Text.StringBuilder
        foreach ($tier in $RoleTierOrder) {
            $count = @($Roster | Where-Object { $_.Role -eq $tier }).Count
            [void]$sb.Append(@"
        <tr>
          <td>$(Get-RoleBadge $tier)</td>
          <td><strong>$count</strong></td>
          <td>$(EscHtml $RoleTierMap[$tier].Blurb)</td>
        </tr>
"@)
        }
        $sb.ToString()
    }

    $typeRows = Invoke-ReportSection -Name 'Account types' -ColSpan 2 -Build {
        $sb = New-Object System.Text.StringBuilder
        foreach ($grp in ($Roster | Group-Object AccountType | Sort-Object Name)) {
            [void]$sb.Append(@"
        <tr><td>$(EscHtml $grp.Name)</td><td><strong>$($grp.Count)</strong></td></tr>
"@)
        }
        $sb.ToString()
    }

    # ── 02 Privileged accounts ───────────────────────────────────────────────
    # A whole-domain roster can run to thousands of rows, and `+=` on a string
    # reallocates the entire buffer once per row, so the account-scale loops
    # accumulate into a StringBuilder instead.
    $privRows = Invoke-ReportSection -Name 'Privileged accounts' -ColSpan 7 -Build {
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
        $privSb.ToString()
    }
    if (-not $privRows) {
        $privRows = '<tr><td colspan="7">No accounts hold privileged group membership.</td></tr>'
    }

    # ── 03 Full roster ───────────────────────────────────────────────────────
    $rosterRows = Invoke-ReportSection -Name 'Full account roster' -ColSpan 9 -Build {
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
        $rosterSb.ToString()
    }
    if (-not $rosterRows) {
        $rosterRows = '<tr><td colspan="9">No accounts matched the report scope.</td></tr>'
    }

    # ── 04 Privileged group membership ───────────────────────────────────────
    $groupRows = Invoke-ReportSection -Name 'Privileged group membership' -ColSpan 5 -Build {
        $groupSb = New-Object System.Text.StringBuilder
        foreach ($g in $GroupSummary) {
            $memberNames = @($g.Members)
            $memberCell  = if ($memberNames.Count -gt 0) { EscHtml (($memberNames | Sort-Object) -join ', ') } else { '<em>Empty</em>' }
            $countClass  = if ($memberNames.Count -eq 0) { 'ok' } else { $RoleTierMap[[string]$g.Role].Badge }
            [void]$groupSb.Append(@"
        <tr>
          <td><strong>$(EscHtml $g.Name)</strong></td>
          <td>$(Get-RoleBadge $g.Role)</td>
          <td><span class="tk-badge-$countClass">$($memberNames.Count)</span></td>
          <td>$(EscHtml $g.Reason)</td>
          <td>$memberCell</td>
        </tr>
"@)
        }
        $groupSb.ToString()
    }
    if (-not $groupRows) {
        $groupRows = '<tr><td colspan="5">No privileged groups were resolved in this domain.</td></tr>'
    }

    # ── 05 Flagged for review ────────────────────────────────────────────────
    $flagRows = Invoke-ReportSection -Name 'Flagged for review' -ColSpan 6 -Build {
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
        $flagSb.ToString()
    }
    if (-not $flagRows) {
        $flagRows = '<tr><td colspan="6">No accounts raised a review flag.</td></tr>'
    }

    # ── 01 Authentication policy ─────────────────────────────────────────────
    $policyBadge = @{ Strong = 'ok'; Acceptable = 'warn'; Weak = 'err'; Unknown = 'info' }

    $policyRows = Invoke-ReportSection -Name 'Authentication policy' -ColSpan 4 -Build {
        if (-not $AuthPolicy.Available) {
            return "<tr><td colspan=""4"">Domain password policy could not be read: $(EscHtml $AuthPolicy.Error)</td></tr>"
        }
        $sb = New-Object System.Text.StringBuilder
        foreach ($key in $PolicyKeyOrder) {
            $spec    = $PolicyBaseline[$key]
            $value   = $AuthPolicy.Values[$key]
            $verdict = Get-PolicyVerdict -Key $key -Value $value
            $badge   = $policyBadge[$verdict]
            [void]$sb.Append(@"
        <tr>
          <td><strong>$(EscHtml $spec.Label)</strong></td>
          <td class="tk-mono">$(EscHtml (Format-PolicyValue -Key $key -Value $value))</td>
          <td style="white-space:nowrap"><span class="tk-badge-$badge">$(EscHtml $verdict)</span></td>
          <td>$(EscHtml $spec.Why)</td>
        </tr>
"@)
        }
        $sb.ToString()
    }

    # The questionnaire asks four specific things. Answering them in prose here
    # means the technician can paste the response rather than re-derive it from
    # the table above.
    $answerHtml = '<p>The domain password policy could not be read, so these answers must be gathered by hand.</p>'
    if ($AuthPolicy.Available) {
        $v = $AuthPolicy.Values
        $complexity = if ($v.ComplexityEnabled) { 'complexity enforced (three of five character classes, and the account name may not appear in the password)' }
                      else { '<strong>complexity not enforced</strong>' }
        # Each branch is written as a finished sentence. An earlier version
        # capitalised the first character afterwards, which silently did nothing
        # whenever the sentence opened with a <strong> tag.
        $expiry = if ($v.MaxPasswordAgeDays -eq 0) {
            '<strong>Passwords never expire.</strong>'
        } else {
            "Passwords expire every $($v.MaxPasswordAgeDays) days, and cannot be changed again for $($v.MinPasswordAgeDays) day(s)."
        }
        $lock = if ($v.LockoutThreshold -eq 0) {
            '<strong>Accounts never lock out</strong>, so online password guessing against the domain is never interrupted.'
        } else {
            $dur = if ($v.LockoutDurationMinutes -eq 0) { 'and stay locked until an administrator unlocks them' }
                   else { "and stay locked for $($v.LockoutDurationMinutes) minutes" }
            "Accounts lock after $($v.LockoutThreshold) failed attempts $dur, with the failed-attempt counter resetting after $($v.LockoutObservationMinutes) minutes."
        }
        $answerHtml = @"
<p><strong>Password length and complexity.</strong> Minimum $($v.MinPasswordLength) characters, with $complexity.</p>
<p><strong>Password history.</strong> The last $($v.PasswordHistoryCount) password(s) are remembered and cannot be reused.</p>
<p><strong>Password expiration.</strong> $expiry</p>
<p><strong>Lockout for failed attempts.</strong> $lock</p>
"@
    }

    $psoHtml = ''
    if ($AuthPolicy.Available) {
        if ($AuthPolicy.PsoError) {
            $psoHtml = "<div class=""tk-info-box""><span class=""tk-info-label"">Fine-grained policies</span> Could not be read: $(EscHtml $AuthPolicy.PsoError). Check by hand before answering — a policy here overrides everything above for the accounts it targets.</div>"
        } elseif (@($AuthPolicy.Pso).Count -eq 0) {
            $psoHtml = '<div class="tk-info-box"><span class="tk-info-label">Fine-grained policies</span> None defined, so the settings above apply to every account in the domain.</div>'
        } else {
            $psoSb = New-Object System.Text.StringBuilder
            foreach ($q in $AuthPolicy.Pso) {
                $qExpiry = if ([int]$q.MaxPasswordAgeDays -eq 0) { 'Never' } else { "$($q.MaxPasswordAgeDays) days" }
                [void]$psoSb.Append(@"
        <tr>
          <td><strong>$(EscHtml $q.Name)</strong></td>
          <td>$(EscHtml $q.Precedence)</td>
          <td>$(EscHtml $q.MinPasswordLength)</td>
          <td>$(EscHtml $q.ComplexityEnabled)</td>
          <td>$(EscHtml $q.PasswordHistoryCount)</td>
          <td>$(EscHtml $qExpiry)</td>
          <td>$(EscHtml $q.LockoutThreshold)</td>
          <td>$(EscHtml $q.AppliesTo)</td>
        </tr>
"@)
            }
            $psoHtml = @"
<div class="tk-info-box">
  <span class="tk-info-label">Fine-grained policies</span>
  <strong>$(@($AuthPolicy.Pso).Count) policy(ies) override the defaults above</strong> for the principals they target.
  Answer the questionnaire from these as well, not from the domain default alone.
</div>
<table class="tk-table" style="margin-top:14px">
  <thead><tr><th>Policy</th><th>Precedence</th><th>Min length</th><th>Complexity</th><th>History</th><th>Max age</th><th>Lockout</th><th>Applies to</th></tr></thead>
  <tbody>$($psoSb.ToString())</tbody>
</table>
"@
        }
    }

    $weakCount = 0
    if ($AuthPolicy.Available) {
        foreach ($key in $PolicyKeyOrder) {
            if ((Get-PolicyVerdict -Key $key -Value $AuthPolicy.Values[$key]) -eq 'Weak') { $weakCount++ }
        }
    }
    $policyClass = if (-not $AuthPolicy.Available) { 'info' } elseif ($weakCount -gt 0) { 'err' } else { 'ok' }
    $policyNum   = if (-not $AuthPolicy.Available) { 'n/a' } else { "$weakCount" }

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
  <div class="tk-summary-card $policyClass">
    <div class="tk-summary-num">$policyNum</div>
    <div class="tk-summary-lbl">Weak Policy Settings</div>
  </div>
</div>

<div class="tk-section" id="s01">
  <div class="tk-card-header">
    <span class="tk-section-title">Authentication Policy</span>
    <span class="tk-section-num">Section 01</span>
  </div>
  <div class="tk-card">
    <div class="tk-info-box">
      <span class="tk-info-label">Questionnaire answer</span>
      $answerHtml
    </div>
    <table class="tk-table" style="margin-top:18px">
      <thead><tr><th>Setting</th><th>Configured</th><th>Verdict</th><th>Why it matters</th></tr></thead>
      <tbody>$policyRows</tbody>
    </table>
    $psoHtml
  </div>
</div>

<div class="tk-section" id="s02">
  <div class="tk-card-header">
    <span class="tk-section-title">Access Levels at a Glance</span>
    <span class="tk-section-num">Section 02</span>
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

<div class="tk-section" id="s03">
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

<div class="tk-section" id="s04">
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

<div class="tk-section" id="s05">
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

<div class="tk-section" id="s06">
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

"@ + (Get-TKHtmlFoot -ScriptName 'H.E.R.A.L.D. v3.8.3')

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

# ── Authentication policy ────────────────────────────────────────────────────

Write-Section 'AUTHENTICATION POLICY'

$authPolicy = Get-AuthenticationPolicy -AdCommon $AdCommon

if (-not $authPolicy.Available) {
    Write-Warn "Could not read the domain password policy: $($authPolicy.Error)"
    Write-Info 'The account roster below is unaffected; the policy section of the report will say so.'
} else {
    $policyWeak = 0
    foreach ($key in $PolicyKeyOrder) {
        $value   = $authPolicy.Values[$key]
        $verdict = Get-PolicyVerdict -Key $key -Value $value
        if ($verdict -eq 'Weak') { $policyWeak++ }

        $tone = switch ($verdict) {
            'Strong'     { $C.Success }
            'Acceptable' { $C.Info    }
            'Weak'       { $C.Error   }
            default      { $C.Info    }
        }
        Write-Host ("  {0,-30} {1,-32} {2}" -f `
            $PolicyBaseline[$key].Label, (Format-PolicyValue -Key $key -Value $value), $verdict) -ForegroundColor $tone
    }

    if ($policyWeak -gt 0) {
        Write-Warn "$policyWeak setting(s) fall short of the baseline - see section 01 of the report."
    } else {
        Write-Ok 'All password and lockout settings meet the baseline.'
    }

    if ($authPolicy.PsoError) {
        Write-Warn "Fine-grained password policies could not be read: $($authPolicy.PsoError)"
    } elseif (@($authPolicy.Pso).Count -gt 0) {
        Write-Warn "$(@($authPolicy.Pso).Count) fine-grained password policy(ies) override the defaults above."
        foreach ($q in $authPolicy.Pso) {
            Write-Info ("{0} (precedence {1}) applies to: {2}" -f $q.Name, $q.Precedence, $q.AppliesTo)
        }
    } else {
        Write-Info 'No fine-grained password policies - the defaults apply domain-wide.'
    }
}

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

$privilegedAccounts = @($roster | Where-Object { $_.Rank -lt $RoleTierMap['Standard User'].Rank } |
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

# Every argument is prepared on its own line and the call is splatted.
#
# The report was lost twice to a System.ArgumentException raised at the call
# statement with no Build-HeraldReport frame in the stack trace, which places the
# fault in evaluating or binding the arguments rather than inside the function.
# The type constraints that could have explained binding are already gone
# (3.8.1), leaving the two composite argument expressions -- and both are now
# removed. The @() that wrapped $groupSummary is redundant because
# Build-HeraldReport normalises both collections itself, and the inline Get-Date
# becomes its own statement.
#
# One value per line means a failure names the argument by line number rather
# than implicating the whole call, and splatting binds by name from a plain
# hashtable.
$reportArgs = @{}

try {
    $reportArgs['Roster']          = $roster
    $reportArgs['GroupSummary']    = $groupSummary
    $reportArgs['DomainName']      = $domainName
    $reportArgs['ReportTimestamp'] = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $reportArgs['Counts']          = $counts
    $reportArgs['AuthPolicy']      = $authPolicy

    $html = Build-HeraldReport @reportArgs
    [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
} catch {
    # Name the exception type and the originating line: "could not save the
    # report" on its own gives a technician in the field nothing to act on and
    # nothing to report back.
    $where = Get-FaultLocation $_
    Write-Fail "Could not save the HTML report: $($_.Exception.Message)"
    Write-Info "$($_.Exception.GetType().FullName) at $where"
    Write-Info $_.InvocationInfo.Line.Trim()
    # The stack trace distinguishes a failure at the call statement (parameter
    # binding) from one inside the callee. Without it, a call-site line number
    # is ambiguous between the two.
    foreach ($frame in ($_.ScriptStackTrace -split "`r?`n")) {
        if ($frame.Trim()) { Write-Info $frame.Trim() }
    }

    # The runtime type of each argument is the one thing the previous rounds of
    # diagnostics could not supply. Reported defensively so that a fault here
    # cannot mask the fault being reported.
    try {
        $shapes = foreach ($argName in ($reportArgs.Keys | Sort-Object)) {
            $argValue = $reportArgs[$argName]
            if ($null -eq $argValue) { "$argName=<null>" }
            else { "{0}={1}" -f $argName, $argValue.GetType().FullName }
        }
        if ($shapes) { Write-Info ("argument types: " + ($shapes -join ', ')) }
    } catch {
        Write-Info 'argument types could not be read.'
    }

    Write-TKError -ScriptName 'herald' -Category 'Report' `
        -Message "$($_.Exception.GetType().FullName): $($_.Exception.Message) [$where]"
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
