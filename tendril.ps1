<#
.SYNOPSIS
    T.E.N.D.R.I.L. — Traces Entitlements, Nested Dependencies, Roles, Integrations & Licenses
    Entra ID Group Dependency Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph (and optionally Az, Exchange Online, and
    SharePoint Online via PnP) and enumerates every tenant-level dependency
    on a single Entra ID security or Microsoft 365 group. The report
    answers the question "what breaks if we delete this group?" before
    the deletion happens.

    Audited dependencies:
      - Group properties (type, dynamic rule, mail-enabled, etc.)
      - Members and owners
      - Group-based licensing
      - Conditional Access policies (Include and Exclude conditions)
      - Enterprise Application role assignments
      - Entra directory role assignments (active + PIM-eligible)
      - Nested memberships (group as a child of another group)
      - Administrative Unit membership
      - Intune compliance / configuration / mobile app assignments
      - SharePoint Online site permissions (optional, via PnP.PowerShell)
      - Exchange recipient policies (optional, via ExchangeOnlineManagement):
        transport rules, send-as / send-on-behalf, role groups,
        distribution group nesting
      - Azure subscription RBAC (optional, via Az.Resources): every
        role assignment for the group ID across every subscription
        visible to the signed-in principal

    Results are written to a dark-themed HTML report so the group can be
    retired (or its membership reshaped) without breaking access.

.USAGE
    PS C:\> .\tendril.ps1
    PS C:\> .\tendril.ps1 -GroupName 'All_CNP_Users'
    PS C:\> .\tendril.ps1 -GroupId 'a1b2c3d4-e5f6-7890-abcd-ef0123456789'
    PS C:\> .\tendril.ps1 -GroupName 'All_CNP_Users' -IncludeAzureRbac
    PS C:\> .\tendril.ps1 -GroupName 'All_CNP_Users' -IncludeExchange -IncludeAzureRbac
    PS C:\> .\tendril.ps1 -GroupName 'All_CNP_Users' -IncludeSharePoint -SharePointAdminUrl 'https://contoso-admin.sharepoint.com'
    PS C:\> .\tendril.ps1 -GroupName 'All_CNP_Users' -Unattended

.NOTES
    Version : 1.0
    Read-only — never modifies the target group or any dependency.

#>

param(
    [switch]$Unattended,
    [string]$GroupName       = '',
    [ValidatePattern('^$|^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$GroupId         = '',
    [switch]$IncludeSharePoint,
    [switch]$IncludeExchange,
    [switch]$IncludeAzureRbac,
    [string]$SharePointAdminUrl    = '',
    [int]   $SharePointSiteLimit   = 200,
    [string]$OutputPath            = '',
    [switch]$NoOpen,
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

# -----------------------------------------------------------------------------
# COLOR SCHEMA
# -----------------------------------------------------------------------------

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# -----------------------------------------------------------------------------
# BANNER
# -----------------------------------------------------------------------------

if (-not $Unattended) { Clear-Host }
Write-Host ""
Write-Host "  T.E.N.D.R.I.L.  -  Traces Entitlements, Nested Dependencies, Roles, Integrations & Licenses" -ForegroundColor Cyan
Write-Host "  Entra ID Group Dependency Audit  v1.0" -ForegroundColor Cyan
Write-Host ""

# -----------------------------------------------------------------------------
# MODULE INSTALL
# -----------------------------------------------------------------------------

function Install-OptionalModule {
    param(
        [string[]]$ModuleNames,
        [string]$BundleName,
        [switch]$Required
    )

    $missing = $ModuleNames | Where-Object { -not (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue) }
    if (-not $missing) {
        foreach ($m in $ModuleNames) {
            try { Import-Module $m -ErrorAction Stop } catch { Write-Warn "Could not import ${m}: $($_.Exception.Message)" }
        }
        return $true
    }

    Write-Warn "Missing module(s): $($missing -join ', ')"
    $installTarget = if ($BundleName) { $BundleName } else { $ModuleNames[0] }

    $doInstall = $Unattended
    if (-not $Unattended) {
        $ans = Read-Host "  Install $installTarget for current user? [Y/N]"
        $doInstall = $ans -match '^[Yy]'
    }

    if (-not $doInstall) {
        if ($Required) {
            Write-Fail "$installTarget is required. Exiting."
            exit 1
        }
        Write-Warn "Skipping $installTarget  -  related checks will be omitted."
        return $false
    }

    try {
        Install-Module -Name $installTarget -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
        Write-Ok "$installTarget installed"
    } catch {
        Write-Fail "Install failed for ${installTarget}: $($_.Exception.Message)"
        Write-TKError -ScriptName 'tendril' -Message "Install $installTarget failed: $($_.Exception.Message)" -Category 'Module Install'
        if ($Required) { exit 1 }
        return $false
    }

    foreach ($m in $ModuleNames) {
        try { Import-Module $m -ErrorAction Stop } catch { Write-Warn "Could not import ${m}: $($_.Exception.Message)" }
    }
    return $true
}

Write-Section "MODULE CHECK & INSTALL"

$GraphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Identity.Governance',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.Devices.CorporateManagement'
)
[void](Install-OptionalModule -ModuleNames $GraphModules -BundleName 'Microsoft.Graph' -Required)

$AzAvailable  = $false
$ExoAvailable = $false
$PnpAvailable = $false

if ($IncludeAzureRbac) {
    $AzAvailable  = Install-OptionalModule -ModuleNames @('Az.Accounts','Az.Resources') -BundleName 'Az.Accounts'
    if ($AzAvailable) { $null = Install-OptionalModule -ModuleNames @('Az.Resources') -BundleName 'Az.Resources' }
}
if ($IncludeExchange) {
    $ExoAvailable = Install-OptionalModule -ModuleNames @('ExchangeOnlineManagement') -BundleName 'ExchangeOnlineManagement'
}
if ($IncludeSharePoint) {
    $PnpAvailable = Install-OptionalModule -ModuleNames @('PnP.PowerShell') -BundleName 'PnP.PowerShell'
    if ($PnpAvailable -and [string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
        if ($Unattended) {
            Write-Warn "-IncludeSharePoint set but -SharePointAdminUrl is empty. SharePoint scan will be skipped."
            $PnpAvailable = $false
        } else {
            Write-Host ""
            $SharePointAdminUrl = (Read-Host "  Enter SharePoint admin URL (e.g. https://contoso-admin.sharepoint.com)").Trim()
            if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
                Write-Warn "No admin URL supplied  -  SharePoint scan will be skipped."
                $PnpAvailable = $false
            }
        }
    }
}

Write-Ok "Module check complete"
Write-Host ""

# -----------------------------------------------------------------------------
# GRAPH AUTH
# -----------------------------------------------------------------------------

Write-Section "AUTHENTICATION"

$GraphScopes = @(
    'Group.Read.All',
    'Directory.Read.All',
    'Policy.Read.All',
    'Application.Read.All',
    'RoleManagement.Read.Directory',
    'DeviceManagementConfiguration.Read.All',
    'DeviceManagementApps.Read.All',
    'AdministrativeUnit.Read.All',
    'User.Read.All'
)

try {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Step "Connecting to Microsoft Graph..."
        Write-Info "Scopes: $($GraphScopes -join ', ')"
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext -ErrorAction Stop
    } else {
        Write-Ok "Reusing existing Graph session: $($ctx.Account)"
    }
} catch {
    Write-Fail "Graph authentication failed: $($_.Exception.Message)"
    Write-TKError -ScriptName 'tendril' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
    exit 1
}

$TenantDomain = ''
try {
    $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($org -and $org.VerifiedDomains) {
        $primary = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
        if ($primary) { $TenantDomain = $primary.Name }
    }
} catch {
    # Best-effort  -  tenant display falls back to the signed-in account.
}

Write-Ok "Signed in as: $($ctx.Account)"
if ($TenantDomain) { Write-Ok "Tenant: $TenantDomain" }
Write-Host ""

# -----------------------------------------------------------------------------
# RESOLVE GROUP
# -----------------------------------------------------------------------------

Write-Section "RESOLVING TARGET GROUP"

if (-not $GroupId -and -not $GroupName) {
    if ($Unattended) {
        Write-Fail "Unattended mode requires -GroupId or -GroupName."
        exit 1
    }
    Write-Host ""
    $GroupName = (Read-Host "  Enter the display name of the group to audit").Trim()
    if ([string]::IsNullOrWhiteSpace($GroupName)) {
        Write-Fail "No group supplied. Exiting."
        exit 1
    }
}

$Group = $null
try {
    if ($GroupId) {
        $Group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
    } else {
        $candidates = @(Get-MgGroup -Filter "displayName eq '$($GroupName.Replace("'","''"))'" -ConsistencyLevel eventual -CountVariable groupCount -All -ErrorAction Stop)
        if ($candidates.Count -eq 0) {
            Write-Fail "No group named '$GroupName' found in this tenant."
            exit 1
        }
        if ($candidates.Count -gt 1 -and -not $Unattended) {
            Write-Warn "$($candidates.Count) groups match  -  select one:"
            for ($i = 0; $i -lt $candidates.Count; $i++) {
                Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $candidates[$i].DisplayName, $candidates[$i].Id) -ForegroundColor $C.Info
            }
            $sel = [int](Read-Host "  Select group [1-$($candidates.Count)]") - 1
            if ($sel -lt 0 -or $sel -ge $candidates.Count) {
                Write-Fail "Invalid selection. Exiting."
                exit 1
            }
            $Group = $candidates[$sel]
        } else {
            $Group = $candidates[0]
            if ($candidates.Count -gt 1) {
                Write-Warn "$($candidates.Count) groups match  -  using first ($($Group.Id)). Pass -GroupId to disambiguate."
            }
        }
    }
} catch {
    Write-Fail "Failed to resolve group: $($_.Exception.Message)"
    Write-TKError -ScriptName 'tendril' -Message "Get-MgGroup failed: $($_.Exception.Message)" -Category 'Graph Query'
    exit 1
}

$GroupId       = $Group.Id
$GroupName     = $Group.DisplayName
$GroupMail     = $Group.Mail
$GroupSmtp     = $Group.Mail
$IsDynamic     = ($Group.GroupTypes -contains 'DynamicMembership')
$IsUnified     = ($Group.GroupTypes -contains 'Unified')
$IsMailEnabled = [bool]$Group.MailEnabled
$IsSecurity    = [bool]$Group.SecurityEnabled
$Membership    = if ($IsDynamic) { 'Dynamic' } else { 'Assigned' }
$GroupKind     = if ($IsUnified) {
    'Microsoft 365 Group'
} elseif ($IsMailEnabled -and $IsSecurity) {
    'Mail-enabled Security Group'
} elseif ($IsMailEnabled) {
    'Distribution Group'
} elseif ($IsSecurity) {
    'Security Group'
} else {
    'Other'
}

Write-Ok "Group: $GroupName"
Write-Host "      Id:           $GroupId"            -ForegroundColor $C.Info
Write-Host "      Kind:         $GroupKind"          -ForegroundColor $C.Info
Write-Host "      Membership:   $Membership"         -ForegroundColor $C.Info
Write-Host "      Mail:         $($GroupMail)"       -ForegroundColor $C.Info
if ($Group.MembershipRule) {
    Write-Host "      Rule:         $($Group.MembershipRule)" -ForegroundColor $C.Info
}
Write-Host ""

# -----------------------------------------------------------------------------
# DATA COLLECTION
# -----------------------------------------------------------------------------

Write-Section "COLLECTING DEPENDENCY DATA"

# ---- 1. Members & owners ----------------------------------------------------
Write-Step "Members and owners..."
$Members = @()
$Owners  = @()
try {
    $Members = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop)
} catch {
    Write-Warn "Member enumeration failed: $($_.Exception.Message)"
}
try {
    $Owners = @(Get-MgGroupOwner -GroupId $GroupId -All -ErrorAction Stop)
} catch {
    Write-Warn "Owner enumeration failed: $($_.Exception.Message)"
}
Write-Ok "Members: $($Members.Count) | Owners: $($Owners.Count)"

$MemberRows = foreach ($m in $Members) {
    $p = $m.AdditionalProperties
    [PSCustomObject]@{
        DisplayName = $p['displayName']
        UPN         = $p['userPrincipalName']
        Mail        = $p['mail']
        Enabled     = $p['accountEnabled']
        UserType    = $p['userType']
        ObjectType  = ($p['@odata.type'] -replace '#microsoft\.graph\.', '')
        Id          = $m.Id
    }
}
$OwnerRows = foreach ($o in $Owners) {
    $p = $o.AdditionalProperties
    [PSCustomObject]@{
        DisplayName = $p['displayName']
        UPN         = $p['userPrincipalName']
        ObjectType  = ($p['@odata.type'] -replace '#microsoft\.graph\.', '')
        Id          = $o.Id
    }
}

# ---- 2. Group-based licensing -----------------------------------------------
Write-Step "Group-based licensing..."
$LicenseRows = @()
if ($Group.AssignedLicenses -and $Group.AssignedLicenses.Count -gt 0) {
    try {
        $AllSkus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
    } catch {
        $AllSkus = @()
        Write-Warn "Could not enumerate tenant SKUs: $($_.Exception.Message)"
    }
    $LicenseRows = foreach ($lic in $Group.AssignedLicenses) {
        $sku = $AllSkus | Where-Object { $_.SkuId -eq $lic.SkuId } | Select-Object -First 1
        [PSCustomObject]@{
            SkuId         = $lic.SkuId
            SkuPartNumber = if ($sku) { $sku.SkuPartNumber } else { '(unknown)' }
            DisabledPlans = if ($lic.DisabledPlans) { ($lic.DisabledPlans -join ', ') } else { '' }
        }
    }
}
Write-Ok "Group-based licenses: $($LicenseRows.Count)"

# ---- 3. Conditional Access --------------------------------------------------
Write-Step "Conditional Access policies..."
$CAHits = [System.Collections.Generic.List[object]]::new()
try {
    $CAPolicies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
    foreach ($p in $CAPolicies) {
        $inc = $p.Conditions.Users.IncludeGroups -contains $GroupId
        $exc = $p.Conditions.Users.ExcludeGroups -contains $GroupId
        if ($inc -or $exc) {
            $CAHits.Add([PSCustomObject]@{
                PolicyName = $p.DisplayName
                PolicyId   = $p.Id
                State      = $p.State
                Role       = if ($inc -and $exc) { 'Include + Exclude' } elseif ($inc) { 'Include' } else { 'Exclude' }
            })
        }
    }
} catch {
    Write-Warn "Conditional Access enumeration failed: $($_.Exception.Message)"
}
Write-Ok "Conditional Access references: $($CAHits.Count)"

# ---- 4. Enterprise app role assignments -------------------------------------
Write-Step "Enterprise application assignments..."
$AppHits = @()
try {
    $appAssignments = @(Get-MgGroupAppRoleAssignment -GroupId $GroupId -All -ErrorAction Stop)
    $AppHits = foreach ($a in $appAssignments) {
        [PSCustomObject]@{
            App          = $a.ResourceDisplayName
            ResourceId   = $a.ResourceId
            AssignmentId = $a.Id
            Created      = if ($a.CreatedDateTime) { ([datetime]$a.CreatedDateTime).ToString('yyyy-MM-dd') } else { '' }
        }
    }
} catch {
    Write-Warn "Enterprise app enumeration failed: $($_.Exception.Message)"
}
Write-Ok "Enterprise app assignments: $($AppHits.Count)"

# ---- 5. Directory role assignments ------------------------------------------
Write-Step "Entra directory role assignments..."
$RoleHits = [System.Collections.Generic.List[object]]::new()
try {
    $activeRoles = @(Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$GroupId'" -All -ErrorAction Stop)
    foreach ($r in $activeRoles) {
        $def = $null
        try { $def = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $r.RoleDefinitionId -ErrorAction Stop } catch {
            # Definition lookup may 403 on highly-restricted roles  -  fall back to the ID.
        }
        $RoleHits.Add([PSCustomObject]@{
            RoleName       = if ($def) { $def.DisplayName } else { $r.RoleDefinitionId }
            Scope          = $r.DirectoryScopeId
            AssignmentKind = 'Active'
        })
    }
} catch {
    Write-Warn "Active role assignment enumeration failed: $($_.Exception.Message)"
}
try {
    $eligible = @(Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "principalId eq '$GroupId'" -All -ErrorAction Stop)
    foreach ($r in $eligible) {
        $def = $null
        try { $def = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $r.RoleDefinitionId -ErrorAction Stop } catch {
            # PIM eligibility schedules can reference roles we can't fully resolve.
        }
        $RoleHits.Add([PSCustomObject]@{
            RoleName       = if ($def) { $def.DisplayName } else { $r.RoleDefinitionId }
            Scope          = $r.DirectoryScopeId
            AssignmentKind = 'PIM-Eligible'
        })
    }
} catch {
    # PIM endpoints require the tenant to have AAD P2; ignore on P1-only tenants.
}
Write-Ok "Directory role assignments: $($RoleHits.Count)"

# ---- 6. Nested memberships --------------------------------------------------
Write-Step "Nested group memberships (this group as child)..."
$NestedRows = @()
try {
    $memberOf = @(Get-MgGroupMemberOf -GroupId $GroupId -All -ErrorAction Stop)
    $NestedRows = foreach ($parent in $memberOf) {
        $p = $parent.AdditionalProperties
        [PSCustomObject]@{
            ParentGroup = $p['displayName']
            ParentId    = $parent.Id
            ParentType  = ($p['@odata.type'] -replace '#microsoft\.graph\.', '')
        }
    }
} catch {
    Write-Warn "Nested membership enumeration failed: $($_.Exception.Message)"
}
Write-Ok "Nested under: $($NestedRows.Count) parent group(s)"

# ---- 7. Administrative units ------------------------------------------------
Write-Step "Administrative Unit membership..."
$AURows = [System.Collections.Generic.List[object]]::new()
try {
    $aus = @(Get-MgDirectoryAdministrativeUnit -All -ErrorAction Stop)
    foreach ($au in $aus) {
        try {
            $auMembers = @(Get-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $au.Id -All -ErrorAction Stop)
            if ($auMembers.Id -contains $GroupId) {
                $AURows.Add([PSCustomObject]@{ AdminUnit = $au.DisplayName; AdminUnitId = $au.Id })
            }
        } catch {
            # AUs can be scoped so the signed-in principal can't list members  -  skip silently.
        }
    }
} catch {
    Write-Warn "Administrative Unit enumeration failed: $($_.Exception.Message)"
}
Write-Ok "Administrative units containing this group: $($AURows.Count)"

# ---- 8. Intune assignments --------------------------------------------------
Write-Step "Intune compliance / configuration / mobile app assignments..."
$IntuneHits = [System.Collections.Generic.List[object]]::new()

function Test-IntuneAssignmentMatch {
    param($Item, [string]$GroupIdToMatch)
    if (-not $Item.Assignments) { return $false }
    foreach ($asn in $Item.Assignments) {
        $tgt = $asn.Target
        if (-not $tgt) { continue }
        # Modern objects expose GroupId directly; older objects buried it under AdditionalProperties.
        $gid = $null
        if ($tgt.PSObject.Properties['GroupId'] -and $tgt.GroupId) {
            $gid = $tgt.GroupId
        } elseif ($tgt.AdditionalProperties -and $tgt.AdditionalProperties['groupId']) {
            $gid = $tgt.AdditionalProperties['groupId']
        }
        if ($gid -eq $GroupIdToMatch) { return $true }
    }
    return $false
}

try {
    $compliance = @(Get-MgDeviceManagementDeviceCompliancePolicy -All -ExpandProperty Assignments -ErrorAction Stop)
    foreach ($c in $compliance) {
        if (Test-IntuneAssignmentMatch -Item $c -GroupIdToMatch $GroupId) {
            $IntuneHits.Add([PSCustomObject]@{ Type='Compliance Policy';   Name=$c.DisplayName; Id=$c.Id })
        }
    }
} catch {
    Write-Warn "Compliance policy enumeration failed: $($_.Exception.Message)"
}
try {
    $configs = @(Get-MgDeviceManagementDeviceConfiguration -All -ExpandProperty Assignments -ErrorAction Stop)
    foreach ($c in $configs) {
        if (Test-IntuneAssignmentMatch -Item $c -GroupIdToMatch $GroupId) {
            $IntuneHits.Add([PSCustomObject]@{ Type='Configuration Profile'; Name=$c.DisplayName; Id=$c.Id })
        }
    }
} catch {
    Write-Warn "Configuration profile enumeration failed: $($_.Exception.Message)"
}
try {
    $apps = @(Get-MgDeviceAppManagementMobileApp -All -ExpandProperty Assignments -ErrorAction Stop)
    foreach ($a in $apps) {
        if (Test-IntuneAssignmentMatch -Item $a -GroupIdToMatch $GroupId) {
            $IntuneHits.Add([PSCustomObject]@{ Type='Mobile App'; Name=$a.DisplayName; Id=$a.Id })
        }
    }
} catch {
    Write-Warn "Mobile app enumeration failed: $($_.Exception.Message)"
}
Write-Ok "Intune assignments referencing this group: $($IntuneHits.Count)"

# ---- 9. SharePoint Online (optional) ----------------------------------------
$SharePointRows = [System.Collections.Generic.List[object]]::new()
$SharePointSiteScanned = 0
$SharePointSiteSkipped = 0
$SharePointConnectedSite = ''
if ($IsUnified) {
    try {
        $rootSite = Get-MgGroupSite -GroupId $GroupId -SiteId 'root' -ErrorAction Stop
        if ($rootSite -and $rootSite.WebUrl) {
            $SharePointConnectedSite = $rootSite.WebUrl
        }
    } catch {
        # M365 group with no provisioned site (rare)  -  ignore.
    }
}
if ($IncludeSharePoint -and $PnpAvailable) {
    Write-Step "SharePoint Online site permissions (capped at $SharePointSiteLimit sites)..."
    try {
        Connect-PnPOnline -Url $SharePointAdminUrl -Interactive -ErrorAction Stop
        $sites = @(Get-PnPTenantSite -ErrorAction Stop)
        if ($SharePointSiteLimit -gt 0 -and $sites.Count -gt $SharePointSiteLimit) {
            $SharePointSiteSkipped = $sites.Count - $SharePointSiteLimit
            $sites = $sites | Select-Object -First $SharePointSiteLimit
        }
        $sitesProcessed = 0
        foreach ($site in $sites) {
            $sitesProcessed++
            if (($sitesProcessed % 25) -eq 0 -or $sitesProcessed -eq $sites.Count) {
                Write-Host ("`r      Scanning site $sitesProcessed/$($sites.Count)...") -NoNewline -ForegroundColor $C.Progress
            }
            try {
                Connect-PnPOnline -Url $site.Url -Interactive -ErrorAction Stop
                $spGroups = @(Get-PnPGroup -ErrorAction SilentlyContinue)
                foreach ($spg in $spGroups) {
                    try {
                        $gm = @(Get-PnPGroupMember -Identity $spg -ErrorAction SilentlyContinue)
                        foreach ($member in $gm) {
                            $login = $member.LoginName
                            if ($login -and ($login -match [regex]::Escape($GroupId) -or $login -match "\|$GroupId$")) {
                                $SharePointRows.Add([PSCustomObject]@{
                                    SiteUrl       = $site.Url
                                    SharePointGroup = $spg.Title
                                    Role          = ($spg.Title -replace '.*?(Owners|Members|Visitors).*','$1')
                                })
                            }
                        }
                    } catch {
                        # Get-PnPGroupMember can 401 on app-only / no-access groups.
                    }
                }
            } catch {
                # Site-level connect can fail for locked / no-access sites; skip and continue.
            }
        }
        $SharePointSiteScanned = $sitesProcessed
        Write-Host ""
        Write-Ok "SharePoint sites scanned: $SharePointSiteScanned (skipped over limit: $SharePointSiteSkipped). Hits: $($SharePointRows.Count)"
    } catch {
        Write-Warn "SharePoint enumeration failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'tendril' -Message "PnP scan failed: $($_.Exception.Message)" -Category 'PnP'
    }
}

# ---- 10. Exchange Online (optional) -----------------------------------------
$EXOTransportHits   = [System.Collections.Generic.List[object]]::new()
$EXOSendAsHits      = [System.Collections.Generic.List[object]]::new()
$EXOSendOnBehalf    = [System.Collections.Generic.List[object]]::new()
$EXORoleGroupHits   = [System.Collections.Generic.List[object]]::new()
$EXOParentDLs       = [System.Collections.Generic.List[object]]::new()
if ($IncludeExchange -and $ExoAvailable) {
    Write-Step "Exchange Online recipient policies..."
    try {
        if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' -and $_.ConnectionUri -like '*outlook.office*' })) {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    } catch {
        Write-Warn "Connect-ExchangeOnline failed: $($_.Exception.Message)"
    }

    # Transport rules referencing the group's SMTP address.
    if ($GroupSmtp) {
        try {
            $rules = @(Get-TransportRule -ResultSize Unlimited -ErrorAction Stop)
            foreach ($r in $rules) {
                $blob = (
                    ($r.SentTo            -join ';'),
                    ($r.SentToMemberOf    -join ';'),
                    ($r.From              -join ';'),
                    ($r.FromMemberOf      -join ';'),
                    ($r.ManagerAddresses  -join ';'),
                    ($r.RedirectMessageTo -join ';')
                ) -join '|'
                if ($blob -match [regex]::Escape($GroupSmtp)) {
                    $EXOTransportHits.Add([PSCustomObject]@{
                        RuleName = $r.Name
                        State    = $r.State
                        Priority = $r.Priority
                    })
                }
            }
        } catch {
            Write-Warn "Transport rule enumeration failed: $($_.Exception.Message)"
        }
    }

    # Send-as / send-on-behalf delegations to the group (only mail-enabled groups).
    if ($IsMailEnabled) {
        try {
            $recipients = @(Get-RecipientPermission -Trustee $GroupSmtp -ErrorAction SilentlyContinue)
            foreach ($p in $recipients) {
                $EXOSendAsHits.Add([PSCustomObject]@{
                    Identity = $p.Identity
                    Access   = ($p.AccessRights -join ', ')
                })
            }
        } catch {
            # Tenant may not permit trustee-based queries.
        }
        try {
            $sobMailboxes = @(Get-Mailbox -ResultSize Unlimited -Filter "GrantSendOnBehalfTo -ne `$null" -ErrorAction Stop)
            foreach ($mbx in $sobMailboxes) {
                $matched = $false
                foreach ($id in $mbx.GrantSendOnBehalfTo) {
                    if ($id -and ($id -ieq $GroupName -or $id -match [regex]::Escape($GroupSmtp))) { $matched = $true; break }
                }
                if ($matched) {
                    $EXOSendOnBehalf.Add([PSCustomObject]@{
                        Mailbox      = $mbx.DisplayName
                        SmtpAddress  = $mbx.PrimarySmtpAddress
                    })
                }
            }
        } catch {
            Write-Warn "Send-on-behalf enumeration failed: $($_.Exception.Message)"
        }
    }

    # Exchange admin role groups that include this group as a member.
    try {
        $rgs = @(Get-RoleGroup -ErrorAction Stop)
        foreach ($rg in $rgs) {
            try {
                $rgMembers = @(Get-RoleGroupMember -Identity $rg.Identity -ErrorAction Stop)
                $matched = $rgMembers | Where-Object {
                    $_.ExternalDirectoryObjectId -eq $GroupId -or
                    $_.Guid -eq $GroupId -or
                    $_.DistinguishedName -match [regex]::Escape($GroupName)
                }
                if ($matched) {
                    $EXORoleGroupHits.Add([PSCustomObject]@{
                        RoleGroup = $rg.DisplayName
                        Roles     = ($rg.Roles -join ', ')
                    })
                }
            } catch {
                # Some role groups are linked to AAD universal groups and 403 on read.
            }
        }
    } catch {
        Write-Warn "Role group enumeration failed: $($_.Exception.Message)"
    }

    # Distribution groups that have this group nested inside them.
    if ($IsMailEnabled) {
        try {
            $allDLs = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
            foreach ($dl in $allDLs) {
                try {
                    $dlm = @(Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited -ErrorAction Stop)
                    if ($dlm | Where-Object { $_.ExternalDirectoryObjectId -eq $GroupId }) {
                        $EXOParentDLs.Add([PSCustomObject]@{
                            DistributionList = $dl.DisplayName
                            SmtpAddress      = $dl.PrimarySmtpAddress
                        })
                    }
                } catch {
                    # Large DLs may exceed the WinRM message size; skip rather than fail the audit.
                }
            }
        } catch {
            Write-Warn "Distribution group enumeration failed: $($_.Exception.Message)"
        }
    }
    Write-Ok ("Exchange hits  -  rules:{0}, send-as:{1}, send-on-behalf:{2}, role groups:{3}, parent DLs:{4}" -f `
        $EXOTransportHits.Count, $EXOSendAsHits.Count, $EXOSendOnBehalf.Count, $EXORoleGroupHits.Count, $EXOParentDLs.Count)
}

# ---- 11. Azure subscription RBAC (optional) ---------------------------------
$AzureRbacHits = [System.Collections.Generic.List[object]]::new()
$AzSubsScanned = 0
if ($IncludeAzureRbac -and $AzAvailable) {
    Write-Step "Azure subscription RBAC..."
    try {
        $azCtx = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $azCtx) {
            Connect-AzAccount -ErrorAction Stop | Out-Null
        }
        $subs = @(Get-AzSubscription -ErrorAction Stop)
        foreach ($sub in $subs) {
            $AzSubsScanned++
            try {
                Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
                $assignments = @(Get-AzRoleAssignment -ObjectId $GroupId -ErrorAction SilentlyContinue)
                foreach ($a in $assignments) {
                    $AzureRbacHits.Add([PSCustomObject]@{
                        Subscription = $sub.Name
                        SubscriptionId = $sub.Id
                        Role         = $a.RoleDefinitionName
                        Scope        = $a.Scope
                    })
                }
            } catch {
                Write-Warn "Could not scan subscription '$($sub.Name)': $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Warn "Azure RBAC enumeration failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'tendril' -Message "Az RBAC scan failed: $($_.Exception.Message)" -Category 'Az'
    }
    Write-Ok "Subscriptions scanned: $AzSubsScanned. RBAC assignments: $($AzureRbacHits.Count)"
}

Write-Host ""

# -----------------------------------------------------------------------------
# BUILD HTML
# -----------------------------------------------------------------------------

Write-Section "BUILDING REPORT"

$reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
$tenantDisplay = if ($TenantDomain) { $TenantDomain } else { $ctx.Account }
$connectedAs   = $ctx.Account
$totalHits     = $LicenseRows.Count + $CAHits.Count + $AppHits.Count + $RoleHits.Count + $NestedRows.Count + $AURows.Count + $IntuneHits.Count + $SharePointRows.Count + $EXOTransportHits.Count + $EXOSendAsHits.Count + $EXOSendOnBehalf.Count + $EXORoleGroupHits.Count + $EXOParentDLs.Count + $AzureRbacHits.Count

function Format-PropertyRow {
    param([string]$Label, [string]$Value, [string]$Badge = '')
    $valHtml = if ($Value) { EscHtml $Value } else { '<span class="tk-badge-info">(none)</span>' }
    if ($Badge) {
        $valHtml = "<span class='$Badge'>$valHtml</span>"
    }
    return "<tr><td style='width:30%'><strong>$(EscHtml $Label)</strong></td><td>$valHtml</td></tr>"
}

# ---- Group properties table -------------------------------------------------
$propRows = [System.Text.StringBuilder]::new()
[void]$propRows.Append((Format-PropertyRow 'Display Name'   $GroupName))
[void]$propRows.Append((Format-PropertyRow 'Object ID'      $GroupId))
[void]$propRows.Append((Format-PropertyRow 'Kind'           $GroupKind))
[void]$propRows.Append((Format-PropertyRow 'Membership'     $Membership))
[void]$propRows.Append((Format-PropertyRow 'Mail-Enabled'   $(if ($IsMailEnabled) {'Yes'} else {'No'})))
[void]$propRows.Append((Format-PropertyRow 'Security-Enabled' $(if ($IsSecurity)  {'Yes'} else {'No'})))
[void]$propRows.Append((Format-PropertyRow 'Mail / SMTP'    $GroupMail))
[void]$propRows.Append((Format-PropertyRow 'Visibility'     $Group.Visibility))
[void]$propRows.Append((Format-PropertyRow 'Created'        $(if ($Group.CreatedDateTime) { ([datetime]$Group.CreatedDateTime).ToString('yyyy-MM-dd') } else { '' })))
if ($IsDynamic -and $Group.MembershipRule) {
    [void]$propRows.Append("<tr><td><strong>Dynamic Rule</strong></td><td><code>$(EscHtml $Group.MembershipRule)</code></td></tr>")
}
if ($SharePointConnectedSite) {
    [void]$propRows.Append((Format-PropertyRow 'Connected SharePoint Site' $SharePointConnectedSite))
}

# ---- Members table ----------------------------------------------------------
$memberRowsHtml = [System.Text.StringBuilder]::new()
if ($MemberRows.Count -eq 0) {
    [void]$memberRowsHtml.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;padding:14px'>No direct members (group may be dynamic, empty, or membership unreadable).</td></tr>")
} else {
    foreach ($m in ($MemberRows | Sort-Object DisplayName | Select-Object -First 250)) {
        $enBadge = if ($m.Enabled) { "<span class='tk-badge-ok'>Yes</span>" } elseif ($null -eq $m.Enabled) { '-' } else { "<span class='tk-badge-warn'>No</span>" }
        [void]$memberRowsHtml.Append("<tr><td>$(EscHtml $m.DisplayName)</td><td>$(EscHtml $m.UPN)</td><td>$(EscHtml $m.UserType)</td><td>$(EscHtml $m.ObjectType)</td><td>$enBadge</td></tr>")
    }
    if ($MemberRows.Count -gt 250) {
        [void]$memberRowsHtml.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;padding:10px'>... and $($MemberRows.Count - 250) more members (truncated).</td></tr>")
    }
}

# ---- Owners table -----------------------------------------------------------
$ownerRowsHtml = [System.Text.StringBuilder]::new()
if ($OwnerRows.Count -eq 0) {
    [void]$ownerRowsHtml.Append("<tr><td colspan='3' class='tk-badge-warn' style='text-align:center;padding:14px'>No owners assigned  -  group is ownerless.</td></tr>")
} else {
    foreach ($o in ($OwnerRows | Sort-Object DisplayName)) {
        [void]$ownerRowsHtml.Append("<tr><td>$(EscHtml $o.DisplayName)</td><td>$(EscHtml $o.UPN)</td><td>$(EscHtml $o.ObjectType)</td></tr>")
    }
}

# ---- License table ----------------------------------------------------------
$licRowsHtml = [System.Text.StringBuilder]::new()
if ($LicenseRows.Count -eq 0) {
    [void]$licRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No group-based licenses assigned.</td></tr>")
} else {
    foreach ($l in $LicenseRows) {
        $disabled = if ($l.DisabledPlans) { EscHtml $l.DisabledPlans } else { '-' }
        [void]$licRowsHtml.Append("<tr><td><strong>$(EscHtml $l.SkuPartNumber)</strong></td><td><code>$(EscHtml $l.SkuId)</code></td><td>$disabled</td></tr>")
    }
}

# ---- CA table ---------------------------------------------------------------
$caRowsHtml = [System.Text.StringBuilder]::new()
if ($CAHits.Count -eq 0) {
    [void]$caRowsHtml.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;padding:14px'>No Conditional Access policies reference this group.</td></tr>")
} else {
    foreach ($p in ($CAHits | Sort-Object Role, PolicyName)) {
        $stateBadge = switch ($p.State) {
            'enabled'           { "<span class='tk-badge-ok'>Enabled</span>" }
            'enabledForReportingButNotEnforced' { "<span class='tk-badge-info'>Report-Only</span>" }
            'disabled'          { "<span class='tk-badge-info'>Disabled</span>" }
            default             { EscHtml $p.State }
        }
        $roleBadge = if ($p.Role -match 'Exclude') { "<span class='tk-badge-warn'>$(EscHtml $p.Role)</span>" } else { "<span class='tk-badge-blue'>$(EscHtml $p.Role)</span>" }
        [void]$caRowsHtml.Append("<tr><td>$(EscHtml $p.PolicyName)</td><td>$roleBadge</td><td>$stateBadge</td><td><code>$(EscHtml $p.PolicyId)</code></td></tr>")
    }
}

# ---- Enterprise apps table --------------------------------------------------
$appRowsHtml = [System.Text.StringBuilder]::new()
if ($AppHits.Count -eq 0) {
    [void]$appRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No enterprise application role assignments.</td></tr>")
} else {
    foreach ($a in ($AppHits | Sort-Object App)) {
        [void]$appRowsHtml.Append("<tr><td><strong>$(EscHtml $a.App)</strong></td><td><code>$(EscHtml $a.ResourceId)</code></td><td>$(EscHtml $a.Created)</td></tr>")
    }
}

# ---- Directory roles table --------------------------------------------------
$roleRowsHtml = [System.Text.StringBuilder]::new()
if ($RoleHits.Count -eq 0) {
    [void]$roleRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No directory role assignments (active or PIM-eligible).</td></tr>")
} else {
    foreach ($r in ($RoleHits | Sort-Object AssignmentKind, RoleName)) {
        $kindBadge = if ($r.AssignmentKind -eq 'Active') { "<span class='tk-badge-err'>Active</span>" } else { "<span class='tk-badge-warn'>PIM-Eligible</span>" }
        [void]$roleRowsHtml.Append("<tr><td><strong>$(EscHtml $r.RoleName)</strong></td><td>$kindBadge</td><td><code>$(EscHtml $r.Scope)</code></td></tr>")
    }
}

# ---- Nested table -----------------------------------------------------------
$nestRowsHtml = [System.Text.StringBuilder]::new()
if ($NestedRows.Count -eq 0) {
    [void]$nestRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>This group is not nested under any other group.</td></tr>")
} else {
    foreach ($n in ($NestedRows | Sort-Object ParentGroup)) {
        [void]$nestRowsHtml.Append("<tr><td><strong>$(EscHtml $n.ParentGroup)</strong></td><td>$(EscHtml $n.ParentType)</td><td><code>$(EscHtml $n.ParentId)</code></td></tr>")
    }
}

# ---- AU table ---------------------------------------------------------------
$auRowsHtml = [System.Text.StringBuilder]::new()
if ($AURows.Count -eq 0) {
    [void]$auRowsHtml.Append("<tr><td colspan='2' class='tk-badge-ok' style='text-align:center;padding:14px'>Not a member of any administrative unit.</td></tr>")
} else {
    foreach ($au in ($AURows | Sort-Object AdminUnit)) {
        [void]$auRowsHtml.Append("<tr><td><strong>$(EscHtml $au.AdminUnit)</strong></td><td><code>$(EscHtml $au.AdminUnitId)</code></td></tr>")
    }
}

# ---- Intune table -----------------------------------------------------------
$intuneRowsHtml = [System.Text.StringBuilder]::new()
if ($IntuneHits.Count -eq 0) {
    [void]$intuneRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No Intune assignments target this group.</td></tr>")
} else {
    foreach ($i in ($IntuneHits | Sort-Object Type, Name)) {
        $typeBadge = switch ($i.Type) {
            'Compliance Policy'     { "<span class='tk-badge-err'>Compliance</span>" }
            'Configuration Profile' { "<span class='tk-badge-warn'>Config</span>" }
            'Mobile App'            { "<span class='tk-badge-blue'>App</span>" }
            default                 { EscHtml $i.Type }
        }
        [void]$intuneRowsHtml.Append("<tr><td>$typeBadge</td><td><strong>$(EscHtml $i.Name)</strong></td><td><code>$(EscHtml $i.Id)</code></td></tr>")
    }
}

# ---- SharePoint section -----------------------------------------------------
$spSection = ''
if ($IncludeSharePoint) {
    $spRowsHtml = [System.Text.StringBuilder]::new()
    if (-not $PnpAvailable) {
        [void]$spRowsHtml.Append("<tr><td colspan='3' class='tk-badge-warn' style='text-align:center;padding:14px'>PnP.PowerShell not available or admin URL missing  -  SharePoint scan skipped.</td></tr>")
    } elseif ($SharePointRows.Count -eq 0) {
        [void]$spRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No SharePoint site groups reference this Entra group across $SharePointSiteScanned scanned site(s).</td></tr>")
    } else {
        foreach ($s in ($SharePointRows | Sort-Object SiteUrl, SharePointGroup)) {
            [void]$spRowsHtml.Append("<tr><td><strong>$(EscHtml $s.SiteUrl)</strong></td><td>$(EscHtml $s.SharePointGroup)</td><td>$(EscHtml $s.Role)</td></tr>")
        }
    }
    $spFootnote = ''
    if ($SharePointSiteSkipped -gt 0) {
        $spFootnote = "<div class='tk-info-box'><span class='tk-badge-warn'>Truncated</span> $SharePointSiteSkipped site(s) skipped over the $SharePointSiteLimit scan limit  -  pass -SharePointSiteLimit 0 to scan everything.</div>"
    }
    $spSection = @"
<div class="tk-section" id="s09">
  <div class="tk-card-header">
    <span class="tk-section-title">SharePoint Online Site Permissions</span>
    <span class="tk-section-num">$($SharePointRows.Count) hit(s) / $SharePointSiteScanned site(s)</span>
  </div>
  <div class="tk-card">
    $spFootnote
    <table class="tk-table">
      <thead><tr><th>Site URL</th><th>SharePoint Group</th><th>Role</th></tr></thead>
      <tbody>$($spRowsHtml.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# ---- Exchange section -------------------------------------------------------
$exoSection = ''
if ($IncludeExchange) {
    $exoTransHtml = [System.Text.StringBuilder]::new()
    if (-not $ExoAvailable) {
        [void]$exoTransHtml.Append("<tr><td colspan='3' class='tk-badge-warn' style='text-align:center;padding:14px'>ExchangeOnlineManagement not available  -  Exchange scan skipped.</td></tr>")
    } elseif ($EXOTransportHits.Count -eq 0) {
        [void]$exoTransHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No transport rules reference this group's SMTP address.</td></tr>")
    } else {
        foreach ($r in ($EXOTransportHits | Sort-Object Priority)) {
            $stateBadge = if ($r.State -eq 'Enabled') { "<span class='tk-badge-err'>Enabled</span>" } else { "<span class='tk-badge-info'>$(EscHtml $r.State)</span>" }
            [void]$exoTransHtml.Append("<tr><td><strong>$(EscHtml $r.RuleName)</strong></td><td>$stateBadge</td><td>$($r.Priority)</td></tr>")
        }
    }

    $exoDelegHtml = [System.Text.StringBuilder]::new()
    $delegRows = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $EXOSendAsHits)   { $delegRows.Add([PSCustomObject]@{ Type='Send-As';        Target=$s.Identity; Detail=$s.Access }) }
    foreach ($s in $EXOSendOnBehalf) { $delegRows.Add([PSCustomObject]@{ Type='Send-On-Behalf'; Target=$s.Mailbox;  Detail=$s.SmtpAddress }) }
    if ($delegRows.Count -eq 0) {
        [void]$exoDelegHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No mailbox delegations grant this group send rights.</td></tr>")
    } else {
        foreach ($d in $delegRows) {
            $typeBadge = if ($d.Type -eq 'Send-As') { "<span class='tk-badge-err'>$(EscHtml $d.Type)</span>" } else { "<span class='tk-badge-warn'>$(EscHtml $d.Type)</span>" }
            [void]$exoDelegHtml.Append("<tr><td>$typeBadge</td><td><strong>$(EscHtml $d.Target)</strong></td><td>$(EscHtml $d.Detail)</td></tr>")
        }
    }

    $exoRoleHtml = [System.Text.StringBuilder]::new()
    if ($EXORoleGroupHits.Count -eq 0) {
        [void]$exoRoleHtml.Append("<tr><td colspan='2' class='tk-badge-ok' style='text-align:center;padding:14px'>Not a member of any Exchange admin role group.</td></tr>")
    } else {
        foreach ($rg in ($EXORoleGroupHits | Sort-Object RoleGroup)) {
            [void]$exoRoleHtml.Append("<tr><td><strong>$(EscHtml $rg.RoleGroup)</strong></td><td>$(EscHtml $rg.Roles)</td></tr>")
        }
    }

    $exoDlHtml = [System.Text.StringBuilder]::new()
    if ($EXOParentDLs.Count -eq 0) {
        [void]$exoDlHtml.Append("<tr><td colspan='2' class='tk-badge-ok' style='text-align:center;padding:14px'>Not nested inside any distribution group.</td></tr>")
    } else {
        foreach ($dl in ($EXOParentDLs | Sort-Object DistributionList)) {
            [void]$exoDlHtml.Append("<tr><td><strong>$(EscHtml $dl.DistributionList)</strong></td><td>$(EscHtml $dl.SmtpAddress)</td></tr>")
        }
    }

    $exoSection = @"
<div class="tk-section" id="s10">
  <div class="tk-card-header">
    <span class="tk-section-title">Exchange Online Recipient Policies</span>
    <span class="tk-section-num">$($EXOTransportHits.Count + $delegRows.Count + $EXORoleGroupHits.Count + $EXOParentDLs.Count) hit(s)</span>
  </div>
  <div class="tk-card">
    <p class="tk-card-label">Transport Rules Referencing Group SMTP</p>
    <table class="tk-table" style="margin-bottom:24px">
      <thead><tr><th>Rule Name</th><th>State</th><th>Priority</th></tr></thead>
      <tbody>$($exoTransHtml.ToString())</tbody>
    </table>

    <p class="tk-card-label">Mailbox Delegations (Send-As / Send-On-Behalf)</p>
    <table class="tk-table" style="margin-bottom:24px">
      <thead><tr><th>Type</th><th>Target Mailbox</th><th>Detail</th></tr></thead>
      <tbody>$($exoDelegHtml.ToString())</tbody>
    </table>

    <p class="tk-card-label">Exchange Admin Role Groups</p>
    <table class="tk-table" style="margin-bottom:24px">
      <thead><tr><th>Role Group</th><th>Roles</th></tr></thead>
      <tbody>$($exoRoleHtml.ToString())</tbody>
    </table>

    <p class="tk-card-label">Distribution Groups With This Group Nested</p>
    <table class="tk-table">
      <thead><tr><th>Distribution Group</th><th>SMTP</th></tr></thead>
      <tbody>$($exoDlHtml.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# ---- Azure RBAC section -----------------------------------------------------
$azSection = ''
if ($IncludeAzureRbac) {
    $azRowsHtml = [System.Text.StringBuilder]::new()
    if (-not $AzAvailable) {
        [void]$azRowsHtml.Append("<tr><td colspan='3' class='tk-badge-warn' style='text-align:center;padding:14px'>Az modules not available  -  Azure RBAC scan skipped.</td></tr>")
    } elseif ($AzureRbacHits.Count -eq 0) {
        [void]$azRowsHtml.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:14px'>No RBAC assignments across $AzSubsScanned subscription(s).</td></tr>")
    } else {
        foreach ($a in ($AzureRbacHits | Sort-Object Subscription, Role, Scope)) {
            $roleBadge = if ($a.Role -in @('Owner','User Access Administrator','Contributor')) { "<span class='tk-badge-err'>$(EscHtml $a.Role)</span>" } else { "<span class='tk-badge-warn'>$(EscHtml $a.Role)</span>" }
            [void]$azRowsHtml.Append("<tr><td><strong>$(EscHtml $a.Subscription)</strong></td><td>$roleBadge</td><td><code>$(EscHtml $a.Scope)</code></td></tr>")
        }
    }
    $azSection = @"
<div class="tk-section" id="s11">
  <div class="tk-card-header">
    <span class="tk-section-title">Azure Subscription RBAC</span>
    <span class="tk-section-num">$($AzureRbacHits.Count) assignment(s) across $AzSubsScanned sub(s)</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Subscription</th><th>Role</th><th>Scope</th></tr></thead>
      <tbody>$($azRowsHtml.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# ---- Cleanup recommendations ------------------------------------------------
$cleanup = [System.Collections.Generic.List[hashtable]]::new()

if ($LicenseRows.Count -gt 0) {
    $skus = (($LicenseRows | ForEach-Object { $_.SkuPartNumber }) -join ', ')
    $cleanup.Add(@{ Sev='high'; Title="Group-based licensing is active  -  $($LicenseRows.Count) SKU(s)"
        Body="Removing this group will revoke the following license assignments from every direct or indirect member: $(EscHtml $skus). Provision replacement licensing before deletion or document accepted downgrade." })
}
if ($CAHits.Count -gt 0) {
    $exc = @($CAHits | Where-Object { $_.Role -match 'Exclude' })
    if ($exc.Count -gt 0) {
        $names = ($exc | ForEach-Object { $_.PolicyName }) -join ', '
        $cleanup.Add(@{ Sev='high'; Title="Group is used as a Conditional Access EXCLUSION  -  $($exc.Count) policy(ies)"
            Body="Removing this group will re-include members in: $(EscHtml $names). Any user who relied on this exclusion (break-glass accounts, service automation, geo exemptions) will be subject to the policy on next sign-in." })
    }
    $inc = @($CAHits | Where-Object { $_.Role -eq 'Include' })
    if ($inc.Count -gt 0) {
        $names = ($inc | ForEach-Object { $_.PolicyName }) -join ', '
        $cleanup.Add(@{ Sev='med'; Title="Group is used as a Conditional Access INCLUSION  -  $($inc.Count) policy(ies)"
            Body="Removing this group will drop members out of: $(EscHtml $names). Reverify that the intended population is still covered by another include condition." })
    }
}
if ($RoleHits.Count -gt 0) {
    $cleanup.Add(@{ Sev='high'; Title="Directory roles are assigned to this group  -  $($RoleHits.Count) assignment(s)"
        Body="Group-based role assignment is in play. Deleting the group removes the role from every direct member. Reassign the role to individuals or to a replacement group first." })
}
if ($AzureRbacHits.Count -gt 0) {
    $cleanup.Add(@{ Sev='high'; Title="Azure RBAC depends on this group  -  $($AzureRbacHits.Count) assignment(s) across $AzSubsScanned sub(s)"
        Body="Members lose Azure resource access when this group is deleted  -  including any Owner / Contributor / User Access Administrator grants. Reassign at the appropriate scope before deletion." })
}
if ($AppHits.Count -gt 0) {
    $cleanup.Add(@{ Sev='med'; Title="Enterprise application access depends on this group  -  $($AppHits.Count) assignment(s)"
        Body="Apps that require group-based assignment (rather than 'Assignment required = No') will deny access to members when this group is deleted. Reassign or set the app's assignment requirement before removal." })
}
if ($IntuneHits.Count -gt 0) {
    $cleanup.Add(@{ Sev='med'; Title="Intune policies target this group  -  $($IntuneHits.Count) assignment(s)"
        Body="Compliance, configuration, or app-deployment assignments will stop applying to members when this group is deleted. Devices may drift to default (or non-compliant) state." })
}
if ($NestedRows.Count -gt 0) {
    $cleanup.Add(@{ Sev='med'; Title="This group is nested inside $($NestedRows.Count) parent group(s)"
        Body="Deleting this group will remove its membership from each parent group, indirectly removing access from every nested user." })
}
if ($EXOTransportHits.Count -gt 0) {
    $cleanup.Add(@{ Sev='med'; Title="Transport rules reference this group's SMTP  -  $($EXOTransportHits.Count) rule(s)"
        Body="Mail flow rules that route, redirect, or stamp messages based on this group will silently fail-open or fail-closed when the address goes away. Edit or disable the rules before deletion." })
}
if ($EXORoleGroupHits.Count -gt 0) {
    $cleanup.Add(@{ Sev='med'; Title="Exchange admin role groups include this group  -  $($EXORoleGroupHits.Count) role group(s)"
        Body="Members lose Exchange admin rights conferred via the listed role groups. Reassign admin permissions to individuals or another security group first." })
}
if ($SharePointRows.Count -gt 0) {
    $cleanup.Add(@{ Sev='med'; Title="SharePoint sites reference this group  -  $($SharePointRows.Count) site-group entry(ies)"
        Body="SharePoint Owners/Members/Visitors group memberships granted via this AAD group will be revoked on deletion. Reassign at the site level or via a replacement AAD group." })
}
if ($AURows.Count -gt 0) {
    $cleanup.Add(@{ Sev='low'; Title="This group is a member of $($AURows.Count) administrative unit(s)"
        Body="Deletion removes the group from each AU. AU-scoped admins lose the ability to manage this group's membership (which is moot post-deletion) but verify no AU-scoped report depends on its inclusion." })
}
if ($OwnerRows.Count -eq 0) {
    $cleanup.Add(@{ Sev='low'; Title='Group has no owners'
        Body="An ownerless group has no clear party responsible for cleanup approval or member maintenance. Assign at least one owner before retiring or transferring the group." })
}
if ($cleanup.Count -eq 0) {
    $cleanup.Add(@{ Sev='low'; Title='No tenant-level dependencies detected'
        Body="No dependencies were found in the scanned scopes. Verify manually any out-of-scope dependencies before deletion: third-party SaaS provisioning, on-prem AD-synced membership, and any application that resolves groups via Graph at runtime." })
}

$cleanupHtml = [System.Text.StringBuilder]::new()
foreach ($issue in $cleanup) {
    $sevLabel = switch ($issue.Sev) { 'high' { 'High' } 'med' { 'Med' } default { 'Low' } }
    $sevBadge = switch ($issue.Sev) { 'high' { 'tk-badge-err' } 'med' { 'tk-badge-warn' } default { 'tk-badge-ok' } }
    [void]$cleanupHtml.Append("<div class='tk-info-box'><span class='$sevBadge'>$sevLabel</span> <strong>$(EscHtml $issue.Title)</strong><p>$($issue.Body)</p></div>")
}

# ---- Out-of-scope manual checks --------------------------------------------
$manualNotes = [System.Collections.Generic.List[string]]::new()
if (-not $IncludeSharePoint -or -not $PnpAvailable) {
    [void]$manualNotes.Add('SharePoint / OneDrive site permissions  -  rerun with -IncludeSharePoint -SharePointAdminUrl, or enumerate manually via PnP.PowerShell.')
}
if (-not $IncludeExchange -or -not $ExoAvailable) {
    [void]$manualNotes.Add('Exchange Online recipient policies, transport rules, role groups, and DL nesting  -  rerun with -IncludeExchange, or use Get-TransportRule / Get-RoleGroup / Get-DistributionGroupMember.')
}
if (-not $IncludeAzureRbac -or -not $AzAvailable) {
    [void]$manualNotes.Add('Azure subscription RBAC  -  rerun with -IncludeAzureRbac, or check Get-AzRoleAssignment -ObjectId in each subscription.')
}
[void]$manualNotes.Add('On-prem AD sync  -  if this group is sourced from on-prem AD, deletion in Entra ID is reversed on next sync cycle.')
[void]$manualNotes.Add('Third-party SaaS provisioning (Okta, JumpCloud, ServiceNow SCIM)  -  enumerate via each provider`s admin portal.')
[void]$manualNotes.Add('Application code that queries Graph for this group at runtime  -  search source repos for the group ID or display name.')

$manualHtml = [System.Text.StringBuilder]::new()
foreach ($n in $manualNotes) {
    [void]$manualHtml.Append("<li>$(EscHtml $n)</li>")
}

# ---- Assemble document ------------------------------------------------------
$tkCfg     = Get-TKConfig
$orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

$navItems = @(
    'Group Properties',
    'Members & Owners',
    'Group-Based Licensing',
    'Conditional Access',
    'Enterprise Apps',
    'Directory Roles',
    'Nested Memberships',
    'Administrative Units',
    'Intune Assignments'
)
if ($IncludeSharePoint) { $navItems += 'SharePoint Sites' }
if ($IncludeExchange)   { $navItems += 'Exchange Policies' }
if ($IncludeAzureRbac)  { $navItems += 'Azure RBAC' }
$navItems += @('Cleanup Checklist', 'Manual Verification')

$htmlHead = Get-TKHtmlHead `
    -Title      "Group Dependency Audit -- $GroupName" `
    -ScriptName 'T.E.N.D.R.I.L.' `
    -Subtitle   "${orgPrefix}Entra ID Group Dependency Audit -- $tenantDisplay" `
    -MetaItems  ([ordered]@{
        'Generated'    = $reportDate
        'Group'        = $GroupName
        'Object ID'    = $GroupId
        'Kind'         = $GroupKind
        'Connected As' = $connectedAs
        'Tenant'       = $tenantDisplay
    }) `
    -NavItems   $navItems

$htmlFoot = Get-TKHtmlFoot -ScriptName 'T.E.N.D.R.I.L. v1.0'

# Summary cards (colors depend on count)
function Get-SummaryClass { param([int]$Count, [string]$WarnSev = 'warn') if ($Count -gt 0) { $WarnSev } else { 'ok' } }
$licCardClass    = Get-SummaryClass $LicenseRows.Count    'err'
$caCardClass     = Get-SummaryClass $CAHits.Count         'warn'
$roleCardClass   = Get-SummaryClass $RoleHits.Count       'err'
$appCardClass    = Get-SummaryClass $AppHits.Count        'warn'
$intuneCardClass = Get-SummaryClass $IntuneHits.Count     'warn'
$nestedCardClass = Get-SummaryClass $NestedRows.Count     'warn'
$rbacCardClass   = Get-SummaryClass $AzureRbacHits.Count  'err'

$rbacCard   = if ($IncludeAzureRbac)  { "<div class='tk-summary-card $rbacCardClass'><div class='tk-summary-num'>$($AzureRbacHits.Count)</div><div class='tk-summary-lbl'>Azure RBAC Hits</div></div>" }   else { '' }
$spCard     = if ($IncludeSharePoint) { "<div class='tk-summary-card info'><div class='tk-summary-num'>$($SharePointRows.Count)</div><div class='tk-summary-lbl'>SharePoint Hits</div></div>" } else { '' }
$exoCard    = if ($IncludeExchange)   { "<div class='tk-summary-card info'><div class='tk-summary-num'>$($EXOTransportHits.Count + $EXOSendAsHits.Count + $EXOSendOnBehalf.Count + $EXORoleGroupHits.Count + $EXOParentDLs.Count)</div><div class='tk-summary-lbl'>Exchange Hits</div></div>" } else { '' }

$html = $htmlHead + @"

  <!-- Executive Summary -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Executive Summary</span>
      <span class="tk-section-num">$totalHits total dependency hit(s)</span>
    </div>
    <div class="tk-card">
      <div class="tk-summary-row">
        <div class="tk-summary-card info"><div class="tk-summary-num">$($Members.Count)</div><div class="tk-summary-lbl">Members</div></div>
        <div class="tk-summary-card $(if ($OwnerRows.Count -eq 0) { 'warn' } else { 'info' })"><div class="tk-summary-num">$($OwnerRows.Count)</div><div class="tk-summary-lbl">Owners</div></div>
        <div class="tk-summary-card $licCardClass"><div class="tk-summary-num">$($LicenseRows.Count)</div><div class="tk-summary-lbl">Group-Based Licenses</div></div>
        <div class="tk-summary-card $caCardClass"><div class="tk-summary-num">$($CAHits.Count)</div><div class="tk-summary-lbl">Conditional Access Policies</div></div>
        <div class="tk-summary-card $appCardClass"><div class="tk-summary-num">$($AppHits.Count)</div><div class="tk-summary-lbl">Enterprise App Assignments</div></div>
        <div class="tk-summary-card $roleCardClass"><div class="tk-summary-num">$($RoleHits.Count)</div><div class="tk-summary-lbl">Directory Roles</div></div>
        <div class="tk-summary-card $nestedCardClass"><div class="tk-summary-num">$($NestedRows.Count)</div><div class="tk-summary-lbl">Nested Under</div></div>
        <div class="tk-summary-card $intuneCardClass"><div class="tk-summary-num">$($IntuneHits.Count)</div><div class="tk-summary-lbl">Intune Assignments</div></div>
        $spCard
        $exoCard
        $rbacCard
      </div>
      <div class="tk-info-box" style="margin-top:18px">
        <strong>$(EscHtml $GroupName)</strong> ($(EscHtml $GroupKind), $(EscHtml $Membership) membership) has
        <strong>$totalHits</strong> tenant-level dependency hit(s) across the scanned scopes. Review the
        Cleanup Checklist before deleting or refactoring this group.
      </div>
    </div>
  </div>

  <!-- 01 Group properties -->
  <div class="tk-section" id="s01">
    <div class="tk-card-header">
      <span class="tk-section-title">Group Properties</span>
      <span class="tk-section-num">Section 01</span>
    </div>
    <div class="tk-card">
      <table class="tk-table"><tbody>$($propRows.ToString())</tbody></table>
    </div>
  </div>

  <!-- 02 Members & owners -->
  <div class="tk-section" id="s02">
    <div class="tk-card-header">
      <span class="tk-section-title">Members &amp; Owners</span>
      <span class="tk-section-num">$($Members.Count) member(s) / $($OwnerRows.Count) owner(s)</span>
    </div>
    <div class="tk-card">
      <p class="tk-card-label">Owners</p>
      <table class="tk-table" style="margin-bottom:24px">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Type</th></tr></thead>
        <tbody>$($ownerRowsHtml.ToString())</tbody>
      </table>
      <p class="tk-card-label">Members (up to first 250 shown)</p>
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>User Type</th><th>Object Type</th><th>Enabled</th></tr></thead>
        <tbody>$($memberRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 03 Licensing -->
  <div class="tk-section" id="s03">
    <div class="tk-card-header">
      <span class="tk-section-title">Group-Based Licensing</span>
      <span class="tk-section-num">$($LicenseRows.Count) SKU(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>SKU</th><th>SKU ID</th><th>Disabled Service Plans</th></tr></thead>
        <tbody>$($licRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 04 Conditional Access -->
  <div class="tk-section" id="s04">
    <div class="tk-card-header">
      <span class="tk-section-title">Conditional Access Policies</span>
      <span class="tk-section-num">$($CAHits.Count) reference(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Policy</th><th>Role</th><th>State</th><th>Policy ID</th></tr></thead>
        <tbody>$($caRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 05 Enterprise apps -->
  <div class="tk-section" id="s05">
    <div class="tk-card-header">
      <span class="tk-section-title">Enterprise Application Assignments</span>
      <span class="tk-section-num">$($AppHits.Count) assignment(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Application</th><th>Resource ID</th><th>Created</th></tr></thead>
        <tbody>$($appRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 06 Directory roles -->
  <div class="tk-section" id="s06">
    <div class="tk-card-header">
      <span class="tk-section-title">Entra Directory Role Assignments</span>
      <span class="tk-section-num">$($RoleHits.Count) assignment(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Role</th><th>Kind</th><th>Scope</th></tr></thead>
        <tbody>$($roleRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 07 Nested -->
  <div class="tk-section" id="s07">
    <div class="tk-card-header">
      <span class="tk-section-title">Nested Memberships</span>
      <span class="tk-section-num">$($NestedRows.Count) parent group(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Parent Group</th><th>Type</th><th>Parent ID</th></tr></thead>
        <tbody>$($nestRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 08 AUs -->
  <div class="tk-section" id="s08">
    <div class="tk-card-header">
      <span class="tk-section-title">Administrative Unit Membership</span>
      <span class="tk-section-num">$($AURows.Count) AU(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Administrative Unit</th><th>AU ID</th></tr></thead>
        <tbody>$($auRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  <!-- 09 Intune -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Intune Assignments</span>
      <span class="tk-section-num">$($IntuneHits.Count) assignment(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Type</th><th>Name</th><th>Id</th></tr></thead>
        <tbody>$($intuneRowsHtml.ToString())</tbody>
      </table>
    </div>
  </div>

  $spSection
  $exoSection
  $azSection

  <!-- Cleanup -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Cleanup Checklist</span>
      <span class="tk-section-num">$($cleanup.Count) item(s)</span>
    </div>
    <div class="tk-card">
      $($cleanupHtml.ToString())
    </div>
  </div>

  <!-- Manual -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Out-of-Scope -- Verify Manually</span>
      <span class="tk-section-num">Reminders</span>
    </div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Manual checks</span>
        <ul style="margin:8px 0 0 16px;line-height:1.6">
          $($manualHtml.ToString())
        </ul>
      </div>
    </div>
  </div>

"@ + $htmlFoot

# -----------------------------------------------------------------------------
# OUTPUT
# -----------------------------------------------------------------------------

if (-not $OutputPath) {
    $safe      = ($GroupName -replace '[^A-Za-z0-9_.-]', '_')
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "TENDRIL_${safe}_${timestamp}.html"
}

try {
    $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
    Write-TKError -ScriptName 'tendril' -Message "Save report failed: $($_.Exception.Message)" -Category 'Report Output'
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

Show-TKReportResult -Path $OutputPath -Unattended:($NoOpen -or $Unattended)

Write-Host ""
Write-Ok "Audit complete. $totalHits dependency hit(s) recorded."
Write-Host ""

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
