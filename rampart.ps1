# rampart.ps1 - R.A.M.P.A.R.T. — Reviews Access Management Policies And Rule Targeting
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
    R.A.M.P.A.R.T. — Reviews Access Management Policies And Rule Targeting
    Entra ID Conditional Access Posture Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Scores an Entra ID tenant's Conditional Access policies against the
    baseline every tenant should have, then reviews the policies themselves:

      - Baseline coverage: MFA for all users, MFA for privileged roles,
        legacy authentication blocked, and -- where licensed -- sign-in and
        user risk policies and a device-compliance requirement. Security
        defaults are taken into account when no Conditional Access policy
        exists.
      - Hygiene: policies left in report-only or disabled, broad exclusion
        lists on enforcing policies, policies that reference deleted users
        or groups, and trusted named locations wide enough to exempt most
        of the internet.
      - Emergency access: the accounts or groups excluded from every
        enforcing all-users policy -- the break-glass path. None at all
        means a misconfigured policy can lock every administrator out.

    Read-only -- nothing in the tenant is changed. Requires only the
    Microsoft.Graph.Authentication module (offered for install if missing)
    and the Policy.Read.All and Directory.Read.All delegated scopes.

.USAGE
    PS C:\> .\rampart.ps1                    # Interactive: sign in, audit, HTML report
    PS C:\> .\rampart.ps1 -Unattended        # Silent: sign in, audit, export HTML

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
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
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

$GraphScopes = @('Policy.Read.All', 'Directory.Read.All')

# Directory role template IDs are the same in every tenant. These are the roles
# that can take over the tenant or its security configuration, so each should
# sit behind an enforcing MFA policy.
$PrivilegedRoleTemplates = [ordered]@{
    '62e90394-69f5-4237-9190-012177145e10' = 'Global Administrator'
    'e8611ab8-c189-46e8-94e1-60213ab1f814' = 'Privileged Role Administrator'
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' = 'Privileged Authentication Administrator'
    '194ae4cb-b126-40b2-bd5b-6091b380977d' = 'Security Administrator'
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' = 'Conditional Access Administrator'
    '29232cdf-9323-42fd-ade2-1d097af3e4de' = 'Exchange Administrator'
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' = 'SharePoint Administrator'
    'fe930be7-5e62-47db-91af-98c3a49a38b1' = 'User Administrator'
    'c4e39bd9-1100-46d3-8c65-fb160da0071f' = 'Authentication Administrator'
    '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' = 'Application Administrator'
    '158c047a-c907-4556-b7ef-446551a6b5f7' = 'Cloud Application Administrator'
    '729827e3-9c14-49f7-bb1b-9608f156bbb8' = 'Helpdesk Administrator'
    '966707d0-3269-4727-9be2-8c3a10f19b9d' = 'Password Administrator'
    '3a2c62db-5318-420d-8d74-23affee5d9d5' = 'Intune Administrator'
}
$GlobalAdminTemplateId = '62e90394-69f5-4237-9190-012177145e10'

# An enforcing all-users policy with more direct exclusions than this is worth
# a second look -- each exclusion is a user the baseline does not reach.
$ExclusionThreshold = 5

# Include/exclude values that are keywords rather than object IDs.
$TargetKeywords = @('All', 'None', 'GuestsOrExternalUsers')

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition R.A.M.P.A.R.T. can report, keyed by a stable code. The Pester
# suite extracts it by AST lookup and checks it against the codes the script
# raises.
# ─────────────────────────────────────────────────────────────────────────────

$RampartFindings = @{
    'NoBaseline' = @{
        Severity = 'Error'
        Title    = 'No Conditional Access policy and security defaults off'
        Summary  = 'Nothing requires MFA or blocks legacy authentication anywhere in the tenant.'
        Remedy   = 'Turn security defaults on today, or build the baseline policies: MFA for all users, MFA for admins, block legacy authentication.'
    }
    'SecurityDefaultsOnly' = @{
        Severity = 'Info'
        Title    = 'Protected by security defaults, not Conditional Access'
        Summary  = 'Security defaults give every tenant MFA and a legacy-auth block, but allow no exclusions, locations or device conditions.'
        Remedy   = 'Fine for small tenants. Move to Conditional Access (Entra ID P1) when you need exceptions or device-based rules.'
    }
    'NoMfaAllUsers' = @{
        Severity = 'Error'
        Title    = 'No enforced MFA policy for all users'
        Summary  = 'No enabled policy requires MFA (or an authentication strength) for all users on all cloud apps, so a stolen password alone signs in.'
        Remedy   = 'Create a policy: All users, All cloud apps, Grant: require MFA. Exclude only the emergency-access accounts.'
    }
    'NoMfaAdmins' = @{
        Severity = 'Error'
        Title    = 'Global Administrator is not behind enforced MFA'
        Summary  = 'No enabled policy requires MFA for the Global Administrator role on all cloud apps.'
        Remedy   = 'Create a policy targeting the privileged directory roles, All cloud apps, requiring MFA or a phishing-resistant authentication strength.'
    }
    'AdminRolesNotCovered' = @{
        Severity = 'Warning'
        Title    = 'Privileged roles outside enforced MFA'
        Summary  = 'These privileged directory roles are not covered by any enabled MFA policy for all cloud apps.'
        Remedy   = 'Add them to the admin MFA policy, or cover them with an all-users MFA policy that does not exclude them.'
    }
    'LegacyAuthNotBlocked' = @{
        Severity = 'Error'
        Title    = 'Legacy authentication is not blocked'
        Summary  = 'No enabled policy blocks Exchange ActiveSync and other legacy clients for all users. Legacy protocols cannot do MFA, so they bypass it.'
        Remedy   = 'Create a policy: All users, All cloud apps, Client apps: Exchange ActiveSync + Other clients, Grant: block.'
    }
    'NoRiskPolicies' = @{
        Severity = 'Info'
        Title    = 'No sign-in or user risk policy'
        Summary  = 'No enabled policy reacts to Entra ID Protection risk. Needs Entra ID P2.'
        Remedy   = 'If licensed: require MFA at medium sign-in risk and a secure password change at high user risk.'
    }
    'NoDevicePolicy' = @{
        Severity = 'Info'
        Title    = 'No device-based policy'
        Summary  = 'No enabled policy requires a compliant or hybrid-joined device, so any device with valid credentials and MFA is trusted.'
        Remedy   = 'Consider requiring a compliant device for admin portals or sensitive apps once Intune enrolment is in place.'
    }
    'ReportOnlyPolicy' = @{
        Severity = 'Warning'
        Title    = 'Policy left in report-only mode'
        Summary  = 'Report-only policies are evaluated and logged but never enforced.'
        Remedy   = 'Review its sign-in log impact, then switch it On -- or delete it if it is no longer wanted.'
    }
    'DisabledPolicy' = @{
        Severity = 'Info'
        Title    = 'Disabled policies'
        Summary  = 'Disabled policies do nothing. Old ones make the policy set harder to reason about.'
        Remedy   = 'Delete the ones that are no longer needed.'
    }
    'BroadExclusion' = @{
        Severity = 'Warning'
        Title    = 'Enforcing all-users policy with a long exclusion list'
        Summary  = 'Every excluded user or group skips the control. Long lists usually hold one-off exceptions nobody removed.'
        Remedy   = 'Trim the exclusions to the emergency-access accounts and documented service accounts.'
    }
    'NoEmergencyAccess' = @{
        Severity = 'Warning'
        Title    = 'No emergency-access exclusion'
        Summary  = 'No user or group is excluded from every enforcing all-users policy. A mistake in one policy, or an MFA outage, can lock every administrator out.'
        Remedy   = 'Create two cloud-only emergency-access Global Administrators with long random passwords (or FIDO2 keys), exclude them from all policies, and alert on their sign-ins.'
    }
    'EmergencyAccessPresent' = @{
        Severity = 'Info'
        Title    = 'Emergency-access exclusion in place'
        Summary  = 'These users or groups are excluded from every enforcing all-users policy -- the break-glass path.'
        Remedy   = 'Confirm they are the intended accounts, cloud-only, and that their sign-ins raise an alert.'
    }
    'StaleObjectReference' = @{
        Severity = 'Warning'
        Title    = 'Policy references deleted users or groups'
        Summary  = 'An include or exclude list names an object that no longer exists. A deleted include silently narrows the policy.'
        Remedy   = 'Edit the policy and remove or replace the missing objects.'
    }
    'BroadTrustedLocation' = @{
        Severity = 'Warning'
        Title    = 'Trusted named location is very broad'
        Summary  = 'A trusted IP range this wide covers far more than an office network. Policies that skip MFA from trusted locations skip it for everyone inside it.'
        Remedy   = 'Narrow the range to the organisation''s actual egress addresses.'
    }
    'GraphQueryFailed' = @{
        Severity = 'Warning'
        Title    = 'Part of the tenant could not be read'
        Summary  = 'A Microsoft Graph query failed, so the results below may be incomplete.'
        Remedy   = 'Sign in with an account holding Security Reader, Global Reader or Conditional Access Administrator, and consent to Policy.Read.All.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-RampartFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Subject = '',
        [string]$Detail = ''
    )
    $meta = $RampartFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary
        Remedy = $meta.Remedy; Subject = $Subject; Detail = $Detail
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

function Show-RampartBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗  █████╗ ███╗   ███╗██████╗  █████╗ ██████╗ ████████╗
  ██╔══██╗██╔══██╗████╗ ████║██╔══██╗██╔══██╗██╔══██╗╚══██╔══╝
  ██████╔╝███████║██╔████╔██║██████╔╝███████║██████╔╝   ██║
  ██╔══██╗██╔══██║██║╚██╔╝██║██╔═══╝ ██╔══██║██╔══██╗   ██║
  ██║  ██║██║  ██║██║ ╚═╝ ██║██║     ██║  ██║██║  ██║   ██║
  ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝     ╚═╝╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    R.A.M.P.A.R.T. — Reviews Access Management Policies And Rule Targeting" -ForegroundColor Cyan
    Write-Host "    Entra ID Conditional Access Posture Audit Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
#
# Policies arrive from Graph as nested hashtables; the tests pass hashtables or
# PSCustomObjects. Member access reads both the same way, so nothing below
# cares which. No I/O until the COLLECTORS section.
# ─────────────────────────────────────────────────────────────────────────────

function Get-CaList {
    # Graph returns missing collections as $null and single values unwrapped.
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value | ForEach-Object { "$_" } | Where-Object { $_ })
}

function Test-CaPolicyEnforced {
    param($Policy)
    return "$($Policy.state)" -eq 'enabled'
}

function Test-CaTargetsAllUsers {
    param($Policy)
    return (Get-CaList $Policy.conditions.users.includeUsers) -contains 'All'
}

function Test-CaTargetsAllApps {
    param($Policy)
    return (Get-CaList $Policy.conditions.applications.includeApplications) -contains 'All'
}

function Test-CaRequiresMfa {
    # True when every way through the grant control includes MFA: an
    # authentication strength, or 'mfa' under AND, or an OR list offering
    # nothing but MFA. "MFA OR compliant device" lets a compliant device in
    # without MFA, so it does not count.
    param($Policy)
    $grant = $Policy.grantControls
    if ($null -eq $grant) { return $false }
    if ($null -ne $grant.authenticationStrength) { return $true }
    $controls = Get-CaList $grant.builtInControls
    if ($controls -notcontains 'mfa') { return $false }
    if ("$($grant.operator)" -eq 'OR' -and @($controls | Where-Object { $_ -ne 'mfa' }).Count -gt 0) { return $false }
    return $true
}

function Test-CaBlocks {
    param($Policy)
    return (Get-CaList $Policy.grantControls.builtInControls) -contains 'block'
}

function Test-CaBlocksLegacyAuth {
    # A block for all users whose client-app condition includes both legacy
    # client types. 'all' client apps would block everything, which is not a
    # legacy-auth policy and would lock the tenant, so it does not count.
    param($Policy)
    if (-not (Test-CaBlocks $Policy) -or -not (Test-CaTargetsAllUsers $Policy)) { return $false }
    $apps = Get-CaList $Policy.conditions.clientAppTypes
    return ($apps -contains 'exchangeActiveSync') -and ($apps -contains 'other')
}

function Test-CaCoversRole {
    # Does this policy put the given role behind its control on all apps?
    param($Policy, [string]$RoleTemplateId)
    if (-not (Test-CaTargetsAllApps $Policy)) { return $false }
    $users = $Policy.conditions.users
    if ((Get-CaList $users.excludeRoles) -contains $RoleTemplateId) { return $false }
    if ((Get-CaList $users.includeUsers) -contains 'All') { return $true }
    return (Get-CaList $users.includeRoles) -contains $RoleTemplateId
}

function Test-CaUsesRisk {
    param($Policy)
    return (@(Get-CaList $Policy.conditions.signInRiskLevels).Count + @(Get-CaList $Policy.conditions.userRiskLevels).Count) -gt 0
}

function Test-CaUsesDevice {
    param($Policy)
    $controls = Get-CaList $Policy.grantControls.builtInControls
    return ($controls -contains 'compliantDevice') -or ($controls -contains 'domainJoinedDevice')
}

function Get-CaEmergencyExclusion {
    # The users and groups excluded from EVERY enforcing all-users policy that
    # blocks or requires MFA -- anyone outside that set is caught by at least
    # one of them. Returns $null when there is no such policy to be excluded from.
    param([object[]]$Policies)
    $gates = @($Policies | Where-Object {
        (Test-CaPolicyEnforced $_) -and (Test-CaTargetsAllUsers $_) -and ((Test-CaBlocks $_) -or (Test-CaRequiresMfa $_))
    })
    if ($gates.Count -eq 0) { return $null }

    $users  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in (Get-CaList $gates[0].conditions.users.excludeUsers))  { [void]$users.Add($id) }
    foreach ($id in (Get-CaList $gates[0].conditions.users.excludeGroups)) { [void]$groups.Add($id) }
    foreach ($p in ($gates | Select-Object -Skip 1)) {
        $users.IntersectWith([string[]]@(Get-CaList $p.conditions.users.excludeUsers))
        $groups.IntersectWith([string[]]@(Get-CaList $p.conditions.users.excludeGroups))
    }
    return [PSCustomObject]@{
        GateCount = $gates.Count
        Users     = @($users)
        Groups    = @($groups)
    }
}

function Test-BroadCidr {
    # IPv4 wider than /16, or IPv6 wider than /32, is more than any one
    # organisation's egress. An unparsable range is not judged.
    param([string]$Cidr)
    $parts = "$Cidr".Split('/')
    if ($parts.Count -ne 2) { return $false }
    $prefix = 0
    if (-not [int]::TryParse($parts[1], [ref]$prefix)) { return $false }
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse($parts[0], [ref]$ip)) { return $false }
    if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) { return $prefix -lt 32 }
    return $prefix -lt 16
}

function Get-RampartVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Exposed';  Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Gaps';     Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Enforced'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE + CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

function Install-RampartModule {
    Write-Section "MODULE CHECK"
    $name = 'Microsoft.Graph.Authentication'
    if (Get-Module -ListAvailable -Name $name) {
        Write-Ok "$name — installed"
    } else {
        Write-Warn "$name — NOT found"
        $doInstall = $Unattended
        if (-not $Unattended) {
            $ans = Read-Host "  Install $name for current user? [Y/N]"
            $doInstall = $ans -match '^[Yy]'
        }
        if (-not $doInstall) { Write-Fail "$name is required."; return $false }
        try {
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Ok "$name installed."
        } catch {
            Write-Fail "Install failed: $($_.Exception.Message)"
            Write-TKError -ScriptName 'rampart' -Message "$name install failed: $($_.Exception.Message)" -Category 'Module Install'
            return $false
        }
    }
    try { Import-Module $name -ErrorAction Stop; return $true }
    catch { Write-Fail "Could not import ${name}: $($_.Exception.Message)"; return $false }
}

function Connect-RampartGraph {
    Write-Section "CONNECT TO MICROSOFT GRAPH"
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account -and -not @($GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })) {
            Write-Ok "Reusing existing session: $($ctx.Account)"
            return $ctx.Account
        }
    } catch {
        Write-Verbose "No cached Graph context: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in..."
    Write-Info "Scopes: $($GraphScopes -join ', ')"
    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        Write-Ok "Connected as: $($ctx.Account)"
        return $ctx.Account
    } catch {
        Write-Fail "Connect-MgGraph failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'rampart' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RampartGraph {
    # GET with paging. Returns every item of 'value', or the single object for
    # endpoints that are not collections.
    param([string]$Uri, [string]$Label)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    try {
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
            if ($resp.ContainsKey('value')) {
                foreach ($v in @($resp['value'])) { [void]$items.Add($v) }
                $next = $resp['@odata.nextLink']
            } else {
                return $resp
            }
        }
    } catch {
        Add-RampartFinding -Code 'GraphQueryFailed' -Subject $Label -Detail $_.Exception.Message
        Write-Warn "$Label query failed: $($_.Exception.Message)"
        return $null
    }
    return , $items.ToArray()
}

function Resolve-RampartObjects {
    # directoryObjects/getByIds resolves users and groups in one call per 1000
    # IDs. Anything asked for and not returned no longer exists.
    param([string[]]$Ids)
    $names = @{}
    $ids   = @($Ids | Where-Object { $_ -and $_ -notin $TargetKeywords } | Select-Object -Unique)
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = @($ids[$i..([math]::Min($i + 999, $ids.Count - 1))])
        try {
            $body = @{ ids = $chunk; types = @('user', 'group') } | ConvertTo-Json -Depth 3
            $resp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/directoryObjects/getByIds' -Body $body -ContentType 'application/json' -ErrorAction Stop
            foreach ($o in @($resp['value'])) {
                $label = if ($o['userPrincipalName']) { $o['userPrincipalName'] } else { $o['displayName'] }
                $names["$($o['id'])"] = "$label"
            }
        } catch {
            Add-RampartFinding -Code 'GraphQueryFailed' -Subject 'directoryObjects/getByIds' -Detail $_.Exception.Message
            return $null
        }
    }
    return $names
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RampartAudit {
    Write-Section "CONDITIONAL ACCESS"
    Write-Step "Reading policies, named locations and security defaults..."
    $policies  = Invoke-RampartGraph -Uri 'v1.0/identity/conditionalAccess/policies' -Label 'Conditional Access policies'
    $locations = Invoke-RampartGraph -Uri 'v1.0/identity/conditionalAccess/namedLocations' -Label 'Named locations'
    $defaults  = Invoke-RampartGraph -Uri 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -Label 'Security defaults'

    $policies  = @($policies | Where-Object { $_ })
    $locations = @($locations | Where-Object { $_ })
    $defaultsOn = [bool]($defaults -and $defaults['isEnabled'])
    $enforced  = @($policies | Where-Object { Test-CaPolicyEnforced $_ })

    Write-Info ("Policies: {0} ({1} enforced)  |  Named locations: {2}  |  Security defaults: {3}" -f `
        $policies.Count, $enforced.Count, $locations.Count, $(if ($defaultsOn) { 'on' } else { 'off' }))

    # Resolve every user / group ID the policies mention, once.
    $ids = foreach ($p in $policies) {
        $u = $p.conditions.users
        Get-CaList $u.includeUsers; Get-CaList $u.excludeUsers; Get-CaList $u.includeGroups; Get-CaList $u.excludeGroups
    }
    $names = Resolve-RampartObjects -Ids @($ids)
    function _name { param([string]$Id) if ($names -and $names.ContainsKey($Id)) { $names[$Id] } else { $Id } }

    # ── Baseline coverage ──
    $coverage = [System.Collections.Generic.List[object]]::new()
    function _cover { param([string]$Check, [object[]]$By, [string]$Missing)
        [void]$coverage.Add([PSCustomObject]@{ Check = $Check; Met = ($By.Count -gt 0); By = (@($By | ForEach-Object { $_.displayName }) -join '; '); Missing = $Missing })
    }

    if ($policies.Count -eq 0 -or $enforced.Count -eq 0) {
        if ($defaultsOn) { Add-RampartFinding -Code 'SecurityDefaultsOnly' -Subject 'Tenant' }
        else             { Add-RampartFinding -Code 'NoBaseline' -Subject 'Tenant' }
    }

    $mfaAll = @($enforced | Where-Object { (Test-CaTargetsAllUsers $_) -and (Test-CaTargetsAllApps $_) -and (Test-CaRequiresMfa $_) })
    $legacy = @($enforced | Where-Object { Test-CaBlocksLegacyAuth $_ })
    $risk   = @($enforced | Where-Object { Test-CaUsesRisk $_ })
    $device = @($enforced | Where-Object { Test-CaUsesDevice $_ })
    $mfaGa  = @($enforced | Where-Object { (Test-CaRequiresMfa $_) -and (Test-CaCoversRole -Policy $_ -RoleTemplateId $GlobalAdminTemplateId) })

    _cover 'MFA for all users, all apps'       $mfaAll 'NoMfaAllUsers'
    _cover 'MFA for Global Administrator'      $mfaGa  'NoMfaAdmins'
    _cover 'Legacy authentication blocked'     $legacy 'LegacyAuthNotBlocked'
    _cover 'Sign-in / user risk policy'        $risk   'NoRiskPolicies'
    _cover 'Device-based requirement'          $device 'NoDevicePolicy'

    # Security defaults already enforce MFA and block legacy auth, so with
    # them on (and no CA to replace them) those two baseline gaps are closed.
    if (-not $defaultsOn -or $enforced.Count -gt 0) {
        if ($mfaAll.Count -eq 0) { Add-RampartFinding -Code 'NoMfaAllUsers' -Subject 'Tenant' }
        if ($mfaGa.Count  -eq 0) { Add-RampartFinding -Code 'NoMfaAdmins'  -Subject 'Global Administrator' }
        if ($legacy.Count -eq 0) { Add-RampartFinding -Code 'LegacyAuthNotBlocked' -Subject 'Tenant' }
    }
    if ($enforced.Count -gt 0) {
        if ($risk.Count   -eq 0) { Add-RampartFinding -Code 'NoRiskPolicies' -Subject 'Tenant' }
        if ($device.Count -eq 0) { Add-RampartFinding -Code 'NoDevicePolicy' -Subject 'Tenant' }
    }

    if ($mfaGa.Count -gt 0) {
        $uncovered = @($PrivilegedRoleTemplates.Keys | Where-Object {
            $role = $_
            @($enforced | Where-Object { (Test-CaRequiresMfa $_) -and (Test-CaCoversRole -Policy $_ -RoleTemplateId $role) }).Count -eq 0
        } | ForEach-Object { $PrivilegedRoleTemplates[$_] })
        if ($uncovered.Count -gt 0) { Add-RampartFinding -Code 'AdminRolesNotCovered' -Subject "$($uncovered.Count) role(s)" -Detail ($uncovered -join ', ') }
    }

    # ── Hygiene ──
    foreach ($p in ($policies | Where-Object { "$($_.state)" -eq 'enabledForReportingButNotEnforced' })) {
        Add-RampartFinding -Code 'ReportOnlyPolicy' -Subject "$($p.displayName)"
    }
    $disabled = @($policies | Where-Object { "$($_.state)" -eq 'disabled' })
    if ($disabled.Count -gt 0) {
        Add-RampartFinding -Code 'DisabledPolicy' -Subject "$($disabled.Count) policy(ies)" -Detail (@($disabled | ForEach-Object { $_.displayName }) -join '; ')
    }
    foreach ($p in ($enforced | Where-Object { (Test-CaTargetsAllUsers $_) -and ((Test-CaBlocks $_) -or (Test-CaRequiresMfa $_)) })) {
        $ex = @(Get-CaList $p.conditions.users.excludeUsers) + @(Get-CaList $p.conditions.users.excludeGroups)
        if ($ex.Count -gt $ExclusionThreshold) {
            Add-RampartFinding -Code 'BroadExclusion' -Subject "$($p.displayName)" -Detail ("{0} exclusions: {1}" -f $ex.Count, (@($ex | Select-Object -First 15 | ForEach-Object { _name $_ }) -join ', '))
        }
    }

    if ($null -ne $names) {
        $stale = [System.Collections.Generic.List[string]]::new()
        foreach ($p in $policies) {
            $u = $p.conditions.users
            $missing = @(@(Get-CaList $u.includeUsers) + @(Get-CaList $u.excludeUsers) + @(Get-CaList $u.includeGroups) + @(Get-CaList $u.excludeGroups) |
                Where-Object { $_ -notin $TargetKeywords -and -not $names.ContainsKey($_) })
            if ($missing.Count -gt 0) { [void]$stale.Add("$($p.displayName): $($missing -join ', ')") }
        }
        if ($stale.Count -gt 0) { Add-RampartFinding -Code 'StaleObjectReference' -Subject "$($stale.Count) policy(ies)" -Detail ($stale -join ' | ') }
    }

    $emergency = Get-CaEmergencyExclusion -Policies $policies
    if ($null -ne $emergency) {
        $who = @(@($emergency.Users | ForEach-Object { "user $(_name $_)" }) + @($emergency.Groups | ForEach-Object { "group $(_name $_)" }))
        if ($who.Count -eq 0) { Add-RampartFinding -Code 'NoEmergencyAccess' -Subject "$($emergency.GateCount) enforcing all-users policy(ies)" }
        else                  { Add-RampartFinding -Code 'EmergencyAccessPresent' -Subject "$($who.Count) exclusion(s)" -Detail ($who -join ', ') }
    }

    $locationRows = foreach ($l in $locations) {
        $ranges = @($l.ipRanges | ForEach-Object { "$($_.cidrAddress)" } | Where-Object { $_ })
        $broad  = @($ranges | Where-Object { Test-BroadCidr $_ })
        if ([bool]$l.isTrusted -and $broad.Count -gt 0) {
            Add-RampartFinding -Code 'BroadTrustedLocation' -Subject "$($l.displayName)" -Detail ($broad -join ', ')
        }
        [PSCustomObject]@{
            Name    = "$($l.displayName)"
            Type    = if ($ranges.Count -gt 0) { 'IP ranges' } elseif ($l.countriesAndRegions) { 'Countries' } else { 'Other' }
            Trusted = [bool]$l.isTrusted
            Detail  = if ($ranges.Count -gt 0) { $ranges -join ', ' } else { (@($l.countriesAndRegions) -join ', ') }
            Broad   = $broad.Count -gt 0
        }
    }

    # ── Inventory rows ──
    $policyRows = foreach ($p in ($policies | Sort-Object { "$($_.state)" -ne 'enabled' }, { "$($_.displayName)" })) {
        $u = $p.conditions.users
        $inc = @(@(Get-CaList $u.includeUsers | ForEach-Object { if ($_ -in $TargetKeywords) { $_ } else { _name $_ } }) +
                 @(Get-CaList $u.includeGroups | ForEach-Object { "group $(_name $_)" }) +
                 @(Get-CaList $u.includeRoles  | ForEach-Object { if ($PrivilegedRoleTemplates.Contains($_)) { "role $($PrivilegedRoleTemplates[$_])" } else { "role $_" } }))
        $exc = @(@(Get-CaList $u.excludeUsers | ForEach-Object { _name $_ }) + @(Get-CaList $u.excludeGroups | ForEach-Object { "group $(_name $_)" }) + @(Get-CaList $u.excludeRoles | ForEach-Object { "role $_" }))
        $grant = @(Get-CaList $p.grantControls.builtInControls)
        if ($p.grantControls -and $p.grantControls.authenticationStrength) { $grant += "strength: $($p.grantControls.authenticationStrength.displayName)" }
        [PSCustomObject]@{
            Name       = "$($p.displayName)"
            State      = "$($p.state)"
            Users      = $inc -join ', '
            Excluded   = $exc -join ', '
            Apps       = (Get-CaList $p.conditions.applications.includeApplications) -join ', '
            ClientApps = (Get-CaList $p.conditions.clientAppTypes) -join ', '
            Grant      = if ($grant.Count) { "$(if ($p.grantControls.operator) { "$($p.grantControls.operator): " })$($grant -join ', ')" } else { '(session controls only)' }
        }
    }

    return [PSCustomObject]@{
        Policies     = @($policyRows)
        PolicyCount  = $policies.Count
        Enforced     = $enforced.Count
        DefaultsOn   = $defaultsOn
        Coverage     = @($coverage)
        Locations    = @($locationRows)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE + HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Show-RampartFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Nothing to report."; return }
    $order = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        $line = "$($f.Title)" + $(if ($f.Subject) { " -- $($f.Subject)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }
}

function Build-RampartReport {
    param([object]$Audit, [object]$Verdict, [string]$ConnectedAs)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant     = if ($ConnectedAs -match '@(.+)$') { $Matches[1] } else { 'Tenant' }
    $order      = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>Nothing to report.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)</td>" +
            "<td>$(EscHtml $f.Subject)" + $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $cRows = [System.Text.StringBuilder]::new()
    foreach ($c in $Audit.Coverage) {
        $sev   = $RampartFindings[$c.Missing].Severity
        $badge = if ($c.Met) { "<span class='tk-badge-ok'>Met</span>" } else { "<span class='tk-badge-$(Get-SeverityClass $sev)'>Missing</span>" }
        [void]$cRows.Append("<tr><td>$(EscHtml $c.Check)</td><td>$badge</td><td>$(EscHtml $(if ($c.By) { $c.By } else { '-' }))</td></tr>")
    }
    if ($Audit.DefaultsOn) {
        [void]$cRows.Append("<tr><td>Security defaults</td><td><span class='tk-badge-info'>On</span></td><td>Tenant-wide MFA registration and legacy-auth block</td></tr>")
    }

    $pRows = [System.Text.StringBuilder]::new()
    if ($Audit.Policies.Count -eq 0) { [void]$pRows.Append("<tr><td colspan='6'>No Conditional Access policies.</td></tr>") }
    foreach ($p in $Audit.Policies) {
        $stateBadge = switch ($p.State) {
            'enabled'                          { "<span class='tk-badge-ok'>On</span>" }
            'enabledForReportingButNotEnforced' { "<span class='tk-badge-warn'>Report-only</span>" }
            default                            { "<span class='tk-badge-info'>Off</span>" }
        }
        [void]$pRows.Append(
            "<tr><td><strong>$(EscHtml $p.Name)</strong></td><td>$stateBadge</td>" +
            "<td>$(EscHtml $p.Users)" + $(if ($p.Excluded) { "<br/><span class='tk-mono'>excl: $(EscHtml $p.Excluded)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $p.Apps)</td><td>$(EscHtml $p.ClientApps)</td><td>$(EscHtml $p.Grant)</td></tr>")
    }

    $lRows = [System.Text.StringBuilder]::new()
    if ($Audit.Locations.Count -eq 0) { [void]$lRows.Append("<tr><td colspan='4'>No named locations.</td></tr>") }
    foreach ($l in $Audit.Locations) {
        $trust = if ($l.Trusted -and $l.Broad) { "<span class='tk-badge-warn'>Trusted (broad)</span>" } elseif ($l.Trusted) { "<span class='tk-badge-ok'>Trusted</span>" } else { 'No' }
        [void]$lRows.Append("<tr><td>$(EscHtml $l.Name)</td><td>$(EscHtml $l.Type)</td><td>$trust</td><td class='tk-mono'>$(EscHtml $l.Detail)</td></tr>")
    }

    $metCount = @($Audit.Coverage | Where-Object { $_.Met }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'R.A.M.P.A.R.T. Conditional Access Posture Report' `
        -ScriptName 'R.A.M.P.A.R.T.' `
        -Subtitle   "${orgPrefix}Conditional Access Posture -- $tenant" `
        -MetaItems  ([ordered]@{
            'Tenant'       = $tenant
            'Signed in as' = $ConnectedAs
            'Generated'    = $reportDate
            'Verdict'      = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Baseline Coverage', 'Policies', 'Named Locations')

    $errCount  = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(if ($errCount) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errCount</div><div class="tk-summary-lbl">Baseline Gaps</div></div>
    <div class="tk-summary-card $(if ($warnCount) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$metCount / $($Audit.Coverage.Count)</div><div class="tk-summary-lbl">Baseline Checks Met</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Enforced) / $($Audit.PolicyCount)</div><div class="tk-summary-lbl">Policies Enforced</div></div>
    <div class="tk-summary-card $(if ($Audit.DefaultsOn) { 'ok' } else { 'info' })"><div class="tk-summary-num">$(if ($Audit.DefaultsOn) { 'On' } else { 'Off' })</div><div class="tk-summary-lbl">Security Defaults</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Baseline Coverage</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Check</th><th>Status</th><th>Satisfied by</th></tr></thead>
      <tbody>$($cRows.ToString())</tbody></table>
      <div class="tk-info-box"><span class="tk-info-label">Counting rule</span> Only enabled policies count. "MFA OR compliant device" does not count as MFA, because a compliant device gets in without it.</div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Policies</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Policy</th><th>State</th><th>Users</th><th>Apps</th><th>Client apps</th><th>Grant</th></tr></thead>
      <tbody>$($pRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Named Locations</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Name</th><th>Type</th><th>Trusted</th><th>Ranges / countries</th></tr></thead>
      <tbody>$($lRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'R.A.M.P.A.R.T. v5.1')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RampartRun {
    $Findings.Clear()

    if (-not (Install-RampartModule)) { if ($Unattended) { exit 1 }; return }
    $who = Connect-RampartGraph
    if (-not $who) { if ($Unattended) { exit 1 }; return }

    $audit = Invoke-RampartAudit
    Show-RampartFindings

    $verdict = Get-RampartVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }

    Add-TKNote -Text ("RAMPART Conditional Access audit: verdict {0}; {1} of {2} policies enforced; {3} finding(s)." -f $verdict.Verdict, $audit.Enforced, $audit.PolicyCount, $Findings.Count) -Category 'Info' -ScriptName 'rampart'

    Write-Step "Generating HTML report..."
    $html    = Build-RampartReport -Audit $audit -Verdict $verdict -ConnectedAs "$who"
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("RAMPART_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'rampart' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

Show-RampartBanner
Invoke-RampartRun

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
