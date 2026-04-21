<#
.SYNOPSIS
    T.A.L.I.S.M.A.N.  -  Tenant Assessment, Logging, Infrastructure, Security, Monitoring & Access Navigator
    Azure subscription assessment and HTML report generator for PowerShell 5.1+

.DESCRIPTION
    Connects to Azure, installs any missing Az modules automatically, enumerates
    all resources, and produces a styled HTML assessment report covering:
    services in use, security posture (NSG exposure, public storage, SQL firewall,
    HTTPS enforcement), access & governance (RBAC, resource locks, policy compliance),
    backup coverage, VM inventory, SQL database hygiene, orphaned resources,
    tag coverage, Azure Advisor alerts, Defender secure score, and prioritized
    remediation recommendations.

.USAGE
    PS C:\> .\talisman.ps1
    PS C:\> .\talisman.ps1 -SubscriptionId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    PS C:\> .\talisman.ps1 -OutputPath "C:\Reports\azure.html" -NoOpen

.NOTES
    Version  : 2.2
    All required Az modules are installed automatically on first run.

    Tools Available
    -----------------------------------------------------------------
    G.R.I.M.O.I.R.E.        -  Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.      -  Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N.  -  Windows Update management
    C.O.N.J.U.R.E.          -  Software deployment via winget / Chocolatey
    A.U.S.P.E.X.            -  System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.        -  Machine onboarding & Entra ID domain join
    R.E.V.E.N.A.N.T.        -  Profile migration & data transfer
    C.I.P.H.E.R.            -  BitLocker drive encryption management
    W.A.R.D.                -  User account & local security audit
    A.R.C.H.I.V.E.          -  Pre-reimaging profile backup
    S.I.G.I.L.              -  Security baseline & policy enforcement
    S.H.A.D.E.              -  Remote machine execution via WinRM
    L.E.Y.L.I.N.E.          -  Network diagnostics & remediation
    F.O.R.G.E.              -  Driver update detection & installation
    T.A.L.I.S.M.A.N.        -  Azure environment & governance inspection
    A.R.T.I.F.A.C.T.        -  Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.            -  Toolkit setup & configuration wizard

    Color Schema
    -----------------------------------------
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success / complete
    Yellow   Warnings / degraded
    Red      Critical errors
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [string]$TenantId        = '',
    [string]$SubscriptionId  = '',
    [string]$OutputPath      = "$env:TEMP\azure-assessment-$(Get-Date -Format 'yyyyMMdd-HHmmss').html",
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

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

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
}

# -----------------------------------------------------------------------------
# BANNER
# -----------------------------------------------------------------------------

Clear-Host
Write-Host @"

  ████████╗ █████╗ ██╗     ██╗███████╗███╗   ███╗ █████╗ ███╗   ██╗
  ╚══██╔══╝██╔══██╗██║     ██║██╔════╝████╗ ████║██╔══██╗████╗  ██║
     ██║   ███████║██║     ██║███████╗██╔████╔██║███████║██╔██╗ ██║
     ██║   ██╔══██║██║     ██║╚════██║██║╚██╔╝██║██╔══██║██║╚██╗██║
     ██║   ██║  ██║███████╗██║███████║██║ ╚═╝ ██║██║  ██║██║ ╚████║
     ╚═╝   ╚═╝  ╚═╝╚══════╝╚═╝╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
Write-Host "  T.A.L.I.S.M.A.N.  -  Tenant Assessment, Logging, Infrastructure, Security, Monitoring & Access Navigator" -ForegroundColor Cyan
Write-Host "  Azure Subscription Assessment & Report Generator  v2.2" -ForegroundColor Cyan
Write-Host ""

# -----------------------------------------------------------------------------
# MODULE AUTO-INSTALL
# -----------------------------------------------------------------------------

Write-Section "MODULE CHECK & INSTALL"

$requiredModules = @(
    'Az.Accounts',
    'Az.Compute',
    'Az.Websites',
    'Az.Sql',
    'Az.Storage',
    'Az.RecoveryServices',
    'Az.Network',
    'Az.Resources',
    'Az.Advisor',
    'Az.Security',
    'Az.PolicyInsights'
)

$needsInstall = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue) }

if ($needsInstall) {
    Write-Warn "Missing modules: $($needsInstall -join ', ')"
    Write-Step "Installing missing modules from PSGallery (CurrentUser scope)..."
    Write-Host "      This only runs once  -  modules are cached after first install." -ForegroundColor $C.Info
    Write-Host ""

    foreach ($mod in $needsInstall) {
        Write-Step "Installing $mod ..."
        try {
            Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Ok "$mod installed"
        } catch {
            Write-Warn "Could not install $mod  -  related checks will be skipped: $($_.Exception.Message)"
        }
    }
    Write-Host ""
}

foreach ($mod in $requiredModules) {
    Import-Module $mod -ErrorAction SilentlyContinue
}
Write-Ok "All modules ready"

# -----------------------------------------------------------------------------
# AUTHENTICATION
# -----------------------------------------------------------------------------

Write-Section "AUTHENTICATION"

$ctx = Get-AzContext -ErrorAction SilentlyContinue

if ($Unattended) {
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        Write-Step "Connecting to tenant: $TenantId"
        try {
            Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
            Connect-AzAccount -TenantId $TenantId -ErrorAction Stop | Out-Null
            $ctx = Get-AzContext -ErrorAction Stop
        } catch {
            Write-Fail "Authentication failed: $_"
            exit 1
        }
    } elseif ($ctx) {
        Write-Step "Unattended  -  using existing session (pass -TenantId to force a specific tenant)"
    } else {
        Write-Step "No active session  -  launching sign-in..."
        try {
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $ctx = Get-AzContext -ErrorAction Stop
        } catch {
            Write-Fail "Authentication failed: $_"
            exit 1
        }
    }
} else {
    # Interactive: always show who is currently signed in and ask before proceeding
    if ($ctx) {
        Write-Host ""
        Write-Host "  Currently signed in:" -ForegroundColor $C.Header
        Write-Host "  Account : $($ctx.Account.Id)" -ForegroundColor $C.Info
        Write-Host "  Tenant  : $($ctx.Tenant.Id)" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Continue with this account? [Y] or sign in to a different tenant [N]: " -ForegroundColor $C.Header
        $useExisting = (Read-Host).Trim().ToUpper()
        if ($useExisting -ne 'Y') {
            $ctx = $null
        }
    }

    if (-not $ctx) {
        Write-Host ""
        Write-Host -NoNewline "  Enter tenant domain or ID (e.g. contoso.com) or leave blank for default: " -ForegroundColor $C.Header
        $tenantInput = (Read-Host).Trim()
        Write-Step "Launching browser sign-in..."
        try {
            Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
            if ([string]::IsNullOrWhiteSpace($tenantInput)) {
                Connect-AzAccount -ErrorAction Stop | Out-Null
            } else {
                Connect-AzAccount -TenantId $tenantInput -ErrorAction Stop | Out-Null
            }
            $ctx = Get-AzContext -ErrorAction Stop
        } catch {
            Write-Fail "Authentication failed: $_"
            exit 1
        }
    }
}

Write-Ok "Signed in as: $($ctx.Account.Id)"

if ($SubscriptionId) {
    try {
        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
    } catch {
        Write-Fail "Could not set subscription '$SubscriptionId': $_"
        exit 1
    }
} else {
    $subs = @(Get-AzSubscription -ErrorAction SilentlyContinue)
    if ($subs.Count -gt 1) {
        Write-Host ""
        Write-Host "  Available Subscriptions:" -ForegroundColor $C.Header
        for ($i = 0; $i -lt $subs.Count; $i++) {
            Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $subs[$i].Name, $subs[$i].Id) -ForegroundColor $C.Info
        }
        Write-Host ""
        Write-Host -NoNewline "  Select subscription [1-$($subs.Count)]: " -ForegroundColor $C.Header
        $sel = [int](Read-Host).Trim() - 1
        if ($sel -ge 0 -and $sel -lt $subs.Count) {
            Set-AzContext -SubscriptionId $subs[$sel].Id -ErrorAction Stop | Out-Null
        }
    }
}

$ctx     = Get-AzContext
$subName = $ctx.Subscription.Name
$subId   = $ctx.Subscription.Id
Write-Ok "Subscription: $subName  ($subId)"
Write-Host ""

# -----------------------------------------------------------------------------
# DATA COLLECTION
# -----------------------------------------------------------------------------

Write-Section "COLLECTING RESOURCE DATA"

Write-Step "All resources..."
$allResources = @(Get-AzResource -ErrorAction SilentlyContinue)
Write-Ok "Found $($allResources.Count) resources"

Write-Step "Resource groups..."
$resourceGroups = @(Get-AzResourceGroup -ErrorAction SilentlyContinue)
Write-Ok "Found $($resourceGroups.Count) resource groups"

Write-Step "Virtual machines..."
$vms = @(Get-AzVM -Status -ErrorAction SilentlyContinue)
Write-Ok "Found $($vms.Count) VMs"

Write-Step "Web apps & function apps..."
$webApps = @(Get-AzWebApp -ErrorAction SilentlyContinue)
Write-Ok "Found $($webApps.Count) sites"

Write-Step "App service plans..."
$appServicePlans = @(Get-AzAppServicePlan -ErrorAction SilentlyContinue)
Write-Ok "Found $($appServicePlans.Count) plans"

Write-Step "SQL servers & databases..."
$sqlServers   = @(Get-AzSqlServer -ErrorAction SilentlyContinue)
$sqlDatabases = @{}
$sqlFirewallIssues = @()
$totalDbCount = 0
foreach ($srv in $sqlServers) {
    $dbs = @(Get-AzSqlDatabase -ServerName $srv.ServerName `
                -ResourceGroupName $srv.ResourceGroupName `
                -ErrorAction SilentlyContinue |
             Where-Object { $_.DatabaseName -ne 'master' })
    $sqlDatabases[$srv.ServerName] = $dbs
    $totalDbCount += $dbs.Count

    $fwRules = @(Get-AzSqlServerFirewallRule -ServerName $srv.ServerName `
                    -ResourceGroupName $srv.ResourceGroupName `
                    -ErrorAction SilentlyContinue)
    foreach ($rule in $fwRules) {
        if ($rule.StartIpAddress -eq '0.0.0.0' -and $rule.EndIpAddress -in @('0.0.0.0','255.255.255.255')) {
            # 0.0.0.0/0.0.0.0 = Azure services rule (acceptable); 0.0.0.0/255.255.255.255 = allow all (not acceptable)
            if ($rule.EndIpAddress -eq '255.255.255.255') {
                $sqlFirewallIssues += [pscustomobject]@{ Server = $srv.ServerName; Rule = $rule.FirewallRuleName; Range = "$($rule.StartIpAddress) - $($rule.EndIpAddress)" }
            }
        } elseif ($rule.StartIpAddress -ne '0.0.0.0' -and $rule.EndIpAddress -eq '255.255.255.255') {
            $sqlFirewallIssues += [pscustomobject]@{ Server = $srv.ServerName; Rule = $rule.FirewallRuleName; Range = "$($rule.StartIpAddress) - $($rule.EndIpAddress)" }
        }
    }
}
Write-Ok "Found $($sqlServers.Count) SQL servers, $totalDbCount databases, $($sqlFirewallIssues.Count) permissive firewall rules"

Write-Step "Storage accounts..."
$storageAccounts = @(Get-AzStorageAccount -ErrorAction SilentlyContinue)
$publicStorageAccounts = @($storageAccounts | Where-Object { $_.AllowBlobPublicAccess -ne $false })
Write-Ok "Found $($storageAccounts.Count) accounts ($($publicStorageAccounts.Count) with public blob access not disabled)"

Write-Step "Recovery services vaults & backup items..."
$recoveryVaults  = @(Get-AzRecoveryServicesVault -ErrorAction SilentlyContinue)
$backedUpVmNames = @()
$backupRows      = [System.Collections.Generic.List[pscustomobject]]::new()
foreach ($vault in $recoveryVaults) {
    try {
        Set-AzRecoveryServicesVaultContext -Vault $vault -ErrorAction Stop
        $items = @(Get-AzRecoveryServicesBackupItem -WorkloadType AzureVM -BackupManagementType AzureVM -ErrorAction SilentlyContinue)
        foreach ($item in $items) {
            $vmFriendly = $item.FriendlyName
            if ($vmFriendly) {
                $backedUpVmNames += $vmFriendly
                $backupRows.Add([pscustomobject]@{
                    VMName        = $vmFriendly
                    Vault         = $vault.Name
                    Status        = $item.ProtectionStatus
                    LastBackup    = if ($item.LastBackupTime) { $item.LastBackupTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
                    PolicyName    = $item.ProtectionPolicyName
                })
            }
        }
    } catch {}
}
$unbackedVms = @($vms | Where-Object { $backedUpVmNames -notcontains $_.Name } | Select-Object -ExpandProperty Name)
Write-Ok "Found $($recoveryVaults.Count) vaults, $($backedUpVmNames.Count) protected VMs, $($unbackedVms.Count) unprotected"

Write-Step "Managed disks..."
$allDisks      = @(Get-AzDisk -ErrorAction SilentlyContinue)
$orphanedDisks = @($allDisks | Where-Object { $_.DiskState -eq 'Unattached' })
Write-Ok "Found $($allDisks.Count) disks ($($orphanedDisks.Count) unattached)"

Write-Step "Public IP addresses..."
$publicIPs       = @(Get-AzPublicIpAddress -ErrorAction SilentlyContinue)
$unassociatedIPs = @($publicIPs | Where-Object { -not $_.IpConfiguration })
Write-Ok "Found $($publicIPs.Count) IPs ($($unassociatedIPs.Count) unassociated)"

Write-Step "Network security groups..."
$nsgs            = @(Get-AzNetworkSecurityGroup -ErrorAction SilentlyContinue)
$exposedNsgRules = @()
$dangerPorts     = @('*', '3389', '22', '5985', '5986', '23', '3306', '1433', '5432')
foreach ($nsg in $nsgs) {
    $dangerous = $nsg.SecurityRules | Where-Object {
        $_.Direction -eq 'Inbound' -and $_.Access -eq 'Allow' -and
        ($_.SourceAddressPrefix -in @('*','Internet','0.0.0.0/0') -or ($_.SourceAddressPrefixes -and $_.SourceAddressPrefixes -contains '*')) -and
        (
            $_.DestinationPortRange -in $dangerPorts -or
            ($_.DestinationPortRanges | Where-Object { $_ -in $dangerPorts })
        )
    }
    foreach ($rule in $dangerous) {
        $port = if ($rule.DestinationPortRange) { $rule.DestinationPortRange } else { ($rule.DestinationPortRanges -join ', ') }
        $exposedNsgRules += [pscustomobject]@{
            NSG      = EscHtml $nsg.Name
            RG       = EscHtml $nsg.ResourceGroupName
            Rule     = EscHtml $rule.Name
            Port     = EscHtml $port
            Priority = $rule.Priority
        }
    }
}
Write-Ok "Found $($nsgs.Count) NSGs ($($exposedNsgRules.Count) rules exposing sensitive ports to the internet)"

Write-Step "RBAC role assignments..."
$allRbac         = @(Get-AzRoleAssignment -ErrorAction SilentlyContinue)
$subOwners       = @($allRbac | Where-Object { $_.RoleDefinitionName -eq 'Owner'       -and $_.Scope -eq "/subscriptions/$subId" })
$subContributors = @($allRbac | Where-Object { $_.RoleDefinitionName -eq 'Contributor' -and $_.Scope -eq "/subscriptions/$subId" })
$subUAAs         = @($allRbac | Where-Object { $_.RoleDefinitionName -eq 'User Access Administrator' -and $_.Scope -eq "/subscriptions/$subId" })
Write-Ok "Subscription-level: $($subOwners.Count) Owners, $($subContributors.Count) Contributors, $($subUAAs.Count) User Access Admins"

Write-Step "Resource locks..."
$resourceLocks = @(Get-AzResourceLock -ErrorAction SilentlyContinue)
$criticalTypes = @('Microsoft.Compute/virtualMachines','Microsoft.Sql/servers','Microsoft.KeyVault/vaults','Microsoft.Storage/storageAccounts','Microsoft.RecoveryServices/vaults')
$criticalResources = @($allResources | Where-Object { $_.ResourceType -in $criticalTypes })
$lockedScopes = $resourceLocks | ForEach-Object {
    $_.ResourceId -replace '/providers/Microsoft.Authorization/locks/.*',''
}
$unlockedCritical = @($criticalResources | Where-Object {
    $resId = $_.ResourceId
    $rgScope = "/subscriptions/$subId/resourceGroups/$($_.ResourceGroupName)"
    -not ($lockedScopes | Where-Object { $_ -ieq $resId -or $_ -ieq $rgScope })
})
Write-Ok "Found $($resourceLocks.Count) locks; $($unlockedCritical.Count) critical resources without a lock"

Write-Step "Azure Policy compliance..."
$nonCompliantPolicies = @()
$totalPolicyStates    = 0
try {
    $nonCompliantPolicies = @(Get-AzPolicyState -Filter "complianceState eq 'NonCompliant'" -Top 500 -ErrorAction Stop |
        Select-Object PolicyDefinitionName, ResourceId, PolicyAssignmentName -Unique)
    $totalPolicyStates = $nonCompliantPolicies.Count
} catch {
    Write-Warn "Policy compliance data unavailable (requires Az.PolicyInsights and Reader+ access)"
}
Write-Ok "Found $totalPolicyStates non-compliant policy states"

Write-Step "Defender for Cloud secure score..."
$defenderScore    = $null
$defenderScoreTxt = 'N/A'
$defenderPct      = $null
try {
    $scores = @(Get-AzSecurityScore -ErrorAction Stop)
    $defenderScore = $scores | Where-Object { $_.Name -eq 'ascScore' } | Select-Object -First 1
    if (-not $defenderScore) { $defenderScore = $scores | Select-Object -First 1 }
    if ($defenderScore -and $defenderScore.MaxScore -gt 0) {
        $defenderPct      = [math]::Round($defenderScore.CurrentScore / $defenderScore.MaxScore * 100)
        $defenderScoreTxt = "$([math]::Round($defenderScore.CurrentScore, 1)) / $($defenderScore.MaxScore)  ($defenderPct%)"
    }
} catch {
    Write-Warn "Defender score unavailable (requires Az.Security and Security Reader access)"
}
Write-Ok "Defender secure score: $defenderScoreTxt"

Write-Step "Azure Advisor recommendations..."
$advisorRecs = @()
try {
    $advisorRecs = @(Get-AzAdvisorRecommendation -ErrorAction Stop)
    Write-Ok "Found $($advisorRecs.Count) Advisor recommendations"
} catch {
    Write-Warn "Advisor data unavailable"
}

Write-Host ""

# -----------------------------------------------------------------------------
# ANALYSIS
# -----------------------------------------------------------------------------

Write-Section "ANALYZING"

# Tag coverage
$taggedCount = ($allResources | Where-Object { $_.Tags -and $_.Tags.Count -gt 0 }).Count
$tagCoverage = if ($allResources.Count -gt 0) { [math]::Round($taggedCount / $allResources.Count * 100) } else { 0 }
$untaggedPct = 100 - $tagCoverage
Write-Step "Tag coverage: $tagCoverage%"

# Advisor counts
$highAdvisor = ($advisorRecs | Where-Object { $_.Impact -eq 'High'   }).Count
$medAdvisor  = ($advisorRecs | Where-Object { $_.Impact -eq 'Medium' }).Count
Write-Step "Advisor: $highAdvisor High, $medAdvisor Medium"

# Ad-hoc databases
$adHocPattern = '\d{8}|\d{4}[-_]\d{2}[-_]\d{2}|backup|copy|_old|_temp|restore'
$adHocDbs = @()
foreach ($srv in $sqlServers) {
    $adHocDbs += @($sqlDatabases[$srv.ServerName] | Where-Object { $_.DatabaseName -imatch $adHocPattern })
}
Write-Step "Ad-hoc databases: $($adHocDbs.Count)"

# Regions
$regions = @($allResources | Select-Object -ExpandProperty Location -Unique | Where-Object { $_ } | Sort-Object)
Write-Step "Regions: $($regions -join ', ')"

# VMs without extensions
$extResources  = $allResources | Where-Object { $_.ResourceType -eq 'Microsoft.Compute/virtualMachines/extensions' }
$vmsWithExts   = $extResources | ForEach-Object { ($_.ResourceId -split '/')[8] } | Select-Object -Unique
$vmsWithoutMon = @($vms | Where-Object { $vmsWithExts -notcontains $_.Name } | Select-Object -ExpandProperty Name)
Write-Step "VMs without extensions: $($vmsWithoutMon.Count)"

# Apps not enforcing HTTPS
$httpApps = @($webApps | Where-Object { -not $_.HttpsOnly })
Write-Step "Apps not enforcing HTTPS: $($httpApps.Count)"

# Suspect app service plans
$suspectPlans = @($appServicePlans | Where-Object { $_.Name -match 'ASP-[A-Za-z0-9]+-[a-f0-9]{4,}$|Plan\d{12,}$|\d{14}' })

Write-Ok "Analysis complete"
Write-Host ""

# -----------------------------------------------------------------------------
# BUILD HTML SECTIONS
# -----------------------------------------------------------------------------

Write-Step "Generating HTML report..."

# -- Services ------------------------------------------------------------------

$serviceItems = [System.Text.StringBuilder]::new()

if ($vms.Count -gt 0) {
    $n = (($vms | Select-Object -ExpandProperty Name) | ForEach-Object { EscHtml $_ }) -join ', '
    [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>Virtual Machines ($($vms.Count))</span><code>Microsoft.Compute/virtualMachines</code><br>$n</div>`n")
}
if ($webApps.Count -gt 0) {
    $n = (($webApps | Select-Object -ExpandProperty Name) | ForEach-Object { EscHtml $_ }) -join ', '
    [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>App Service / Functions ($($webApps.Count) sites)</span><code>Microsoft.Web/sites</code><br>$n</div>`n")
}
if ($sqlServers.Count -gt 0) {
    $n = (($sqlServers | Select-Object -ExpandProperty ServerName) | ForEach-Object { EscHtml $_ }) -join ', '
    [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>Azure SQL ($($sqlServers.Count) servers, $totalDbCount databases)</span><code>Microsoft.Sql/servers</code><br>$n</div>`n")
}
if ($storageAccounts.Count -gt 0) {
    $n = (($storageAccounts | Select-Object -ExpandProperty StorageAccountName) | ForEach-Object { EscHtml $_ }) -join ', '
    [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>Storage Accounts ($($storageAccounts.Count))</span><code>Microsoft.Storage/storageAccounts</code><br>$n</div>`n")
}
if ($recoveryVaults.Count -gt 0) {
    $n = (($recoveryVaults | Select-Object -ExpandProperty Name) | ForEach-Object { EscHtml $_ }) -join ', '
    [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>Recovery Services Vaults ($($recoveryVaults.Count))</span><code>Microsoft.RecoveryServices/vaults</code><br>$n</div>`n")
}
if ($appServicePlans.Count -gt 0) {
    $n = (($appServicePlans | Select-Object -ExpandProperty Name) | ForEach-Object { EscHtml $_ }) -join ', '
    [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>App Service Plans ($($appServicePlans.Count))</span><code>Microsoft.Web/serverfarms</code><br>$n</div>`n")
}

$extraTypes = [ordered]@{
    'Microsoft.DataFactory/factories'                = 'Azure Data Factory'
    'Microsoft.Logic/workflows'                      = 'Logic Apps'
    'Microsoft.KeyVault/vaults'                      = 'Azure Key Vault'
    'Microsoft.Network/virtualNetworkGateways'       = 'VPN / ExpressRoute Gateways'
    'Microsoft.Network/bastionHosts'                 = 'Azure Bastion'
    'Microsoft.AppConfiguration/configurationStores' = 'App Configuration'
    'Microsoft.Web/staticSites'                      = 'Static Web Apps'
    'Microsoft.Network/privateEndpoints'             = 'Private Endpoints'
    'Microsoft.ContainerRegistry/registries'         = 'Container Registry'
    'Microsoft.ContainerService/managedClusters'     = 'AKS Clusters'
    'Microsoft.ApiManagement/service'                = 'API Management'
    'Microsoft.Cdn/profiles'                         = 'Azure CDN'
    'Microsoft.SaaS/resources'                       = 'SaaS Resources'
}
foreach ($type in $extraTypes.Keys) {
    $matched = @($allResources | Where-Object { $_.ResourceType -ieq $type })
    if ($matched.Count -gt 0) {
        $n = ($matched | ForEach-Object { EscHtml $_.Name }) -join ', '
        [void]$serviceItems.Append("<div class='tk-info-box'><span class='tk-info-label'>$($extraTypes[$type]) ($($matched.Count))</span><code>$type</code><br>$n</div>`n")
    }
}

# -- Security Posture ----------------------------------------------------------

$secRows = [System.Text.StringBuilder]::new()

# NSG rules
foreach ($rule in ($exposedNsgRules | Sort-Object Priority)) {
    $portLabel = switch ($rule.Port) {
        '3389' { 'RDP (3389)' }  '22'   { 'SSH (22)' }
        '5985' { 'WinRM (5985)' } '5986' { 'WinRM SSL (5986)' }
        '1433' { 'SQL (1433)' }   '3306' { 'MySQL (3306)' }
        '5432' { 'Postgres (5432)' } '23' { 'Telnet (23)' }
        '*'    { 'All Ports (*)' } default { $rule.Port }
    }
    [void]$secRows.Append("<tr><td><span class='tk-badge-err'>NSG</span></td><td><strong>$($rule.NSG)</strong> / $($rule.Rule)</td><td>$portLabel open to Internet</td><td>$($rule.RG)</td></tr>`n")
}

# Public storage
foreach ($sa in $publicStorageAccounts) {
    $n = EscHtml $sa.StorageAccountName
    $rg = EscHtml $sa.ResourceGroupName
    [void]$secRows.Append("<tr><td><span class='tk-badge-warn'>Storage</span></td><td><strong>$n</strong></td><td>Public blob access not disabled</td><td>$rg</td></tr>`n")
}

# SQL allow-all firewall
foreach ($fw in $sqlFirewallIssues) {
    $n = EscHtml $fw.Server; $r = EscHtml $fw.Rule; $range = EscHtml $fw.Range
    [void]$secRows.Append("<tr><td><span class='tk-badge-err'>SQL FW</span></td><td><strong>$n</strong> / $r</td><td>Permissive firewall rule: $range</td><td> - </td></tr>`n")
}

# Apps without HTTPS
foreach ($app in $httpApps) {
    $n = EscHtml $app.Name; $rg = EscHtml $app.ResourceGroup
    [void]$secRows.Append("<tr><td><span class='tk-badge-warn'>App Svc</span></td><td><strong>$n</strong></td><td>HTTPS-only not enforced</td><td>$rg</td></tr>`n")
}

$securityIssueCount = $exposedNsgRules.Count + $publicStorageAccounts.Count + $sqlFirewallIssues.Count + $httpApps.Count

$securitySection = ''
if ($securityIssueCount -gt 0 -or $true) {
    $noIssueRow = if ($secRows.Length -eq 0) { "<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;padding:20px'>No security issues detected in scanned areas.</td></tr>" } else { '' }
    $securitySection = @"
<div class="tk-section">
  <div class="tk-card-header">
    <span class="tk-section-title">Security Posture</span>
    <span class="tk-section-num">Section 3</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Type</th><th>Resource</th><th>Finding</th><th>Resource Group</th></tr></thead>
      <tbody>$($secRows.ToString())$noIssueRow</tbody>
    </table>
  </div>
</div>
"@
}

# -- Access & Governance -------------------------------------------------------

$rbacRows = [System.Text.StringBuilder]::new()
foreach ($r in ($subOwners + $subContributors + $subUAAs | Sort-Object RoleDefinitionName, DisplayName)) {
    $roleClass = if ($r.RoleDefinitionName -eq 'Owner') { 'tk-badge-err' } elseif ($r.RoleDefinitionName -eq 'User Access Administrator') { 'tk-badge-err' } else { 'tk-badge-warn' }
    $who  = EscHtml ($r.DisplayName)
    $type = EscHtml ($r.ObjectType)
    $role = EscHtml ($r.RoleDefinitionName)
    [void]$rbacRows.Append("<tr><td>$who</td><td>$type</td><td><span class='$roleClass'>$role</span></td><td>Subscription</td></tr>`n")
}
if ($rbacRows.Length -eq 0) {
    [void]$rbacRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;padding:16px'>No subscription-level Owner / Contributor assignments found (or insufficient permissions to list).</td></tr>")
}

$lockRows = [System.Text.StringBuilder]::new()
foreach ($res in ($unlockedCritical | Sort-Object ResourceType, Name)) {
    $n  = EscHtml $res.Name
    $t  = EscHtml ($res.ResourceType -replace 'Microsoft\.\w+/','')
    $rg = EscHtml $res.ResourceGroupName
    [void]$lockRows.Append("<tr><td><strong>$n</strong></td><td>$t</td><td>$rg</td><td><span class='tk-badge-err'>No Lock</span></td></tr>`n")
}
if ($lockRows.Length -eq 0) {
    [void]$lockRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;padding:16px'>All critical resources have resource locks.</td></tr>")
}

$policyNote  = if ($totalPolicyStates -eq 0) { 'No non-compliant states found (or no policies assigned / insufficient access).' } else { "$totalPolicyStates non-compliant policy state(s) detected." }
$policyBadge = if ($totalPolicyStates -gt 0) { 'tk-badge-err' } else { 'tk-badge-ok' }

$policyRows = [System.Text.StringBuilder]::new()
if ($nonCompliantPolicies.Count -gt 0) {
    foreach ($p in ($nonCompliantPolicies | Select-Object -First 25)) {
        $pname = EscHtml $p.PolicyDefinitionName
        $res   = EscHtml (($p.ResourceId -split '/')[-1])
        $asgn  = EscHtml $p.PolicyAssignmentName
        [void]$policyRows.Append("<tr><td>$pname</td><td>$asgn</td><td>$res</td></tr>`n")
    }
    if ($nonCompliantPolicies.Count -gt 25) {
        [void]$policyRows.Append("<tr><td colspan='3' class='tk-badge-warn' style='padding:10px 14px'>... and $($nonCompliantPolicies.Count - 25) more  -  review in Azure Portal under Policy &gt; Compliance.</td></tr>")
    }
} else {
    [void]$policyRows.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;padding:16px'>$policyNote</td></tr>")
}

$governanceSection = @"
<div class="tk-section">
  <div class="tk-card-header">
    <span class="tk-section-title">Access &amp; Governance</span>
    <span class="tk-section-num">Section 4</span>
  </div>
  <div class="tk-card">

    <p class="tk-card-label">Subscription-Level Privileged Role Assignments</p>
    <table class="tk-table" style="margin-bottom:28px">
      <thead><tr><th>Principal</th><th>Type</th><th>Role</th><th>Scope</th></tr></thead>
      <tbody>$($rbacRows.ToString())</tbody>
    </table>

    <p class="tk-card-label">Critical Resources Without Delete Locks</p>
    <table class="tk-table" style="margin-bottom:28px">
      <thead><tr><th>Resource Name</th><th>Type</th><th>Resource Group</th><th>Lock Status</th></tr></thead>
      <tbody>$($lockRows.ToString())</tbody>
    </table>

    <p class="tk-card-label">Azure Policy Non-Compliance &nbsp;<span class="$policyBadge">$policyNote</span></p>
    <table class="tk-table">
      <thead><tr><th>Policy Definition</th><th>Assignment</th><th>Non-Compliant Resource</th></tr></thead>
      <tbody>$($policyRows.ToString())</tbody>
    </table>

  </div>
</div>
"@

# -- VM Inventory --------------------------------------------------------------

$vmRows = [System.Text.StringBuilder]::new()
foreach ($vm in $vms) {
    $power   = if ($vm.PowerState) { EscHtml ($vm.PowerState -replace 'VM ','') } else { 'Unknown' }
    $monPill = if ($vmsWithoutMon -contains $vm.Name) { "<span class='tk-badge-err'>None</span>" } else { "<span class='tk-badge-ok'>Deployed</span>" }
    [void]$vmRows.Append("<tr><td><strong>$(EscHtml $vm.Name)</strong></td><td>$(EscHtml $vm.ResourceGroupName)</td><td>$(EscHtml $vm.Location)</td><td>$(EscHtml $vm.HardwareProfile.VmSize)</td><td>$power</td><td>$monPill</td></tr>`n")
}
$vmSection = ''
if ($vms.Count -gt 0) {
    $vmSection = @"
<div class="tk-section">
  <div class="tk-card-header">
    <span class="tk-section-title">Virtual Machine Inventory</span>
    <span class="tk-section-num">Section 5</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>VM Name</th><th>Resource Group</th><th>Location</th><th>Size</th><th>Power State</th><th>Extensions</th></tr></thead>
      <tbody>$($vmRows.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# -- SQL Database Inventory ----------------------------------------------------

$dbRows = [System.Text.StringBuilder]::new()
foreach ($srv in $sqlServers) {
    $dbs = $sqlDatabases[$srv.ServerName]
    if (-not $dbs) { continue }
    foreach ($db in ($dbs | Sort-Object DatabaseName)) {
        $pill = if ($db.DatabaseName -imatch $adHocPattern) { "<span class='tk-badge-warn'>Ad-hoc / Backup</span>" } else { "<span class='tk-badge-ok'>Active</span>" }
        [void]$dbRows.Append("<tr><td>$(EscHtml $db.DatabaseName)</td><td>$(EscHtml $srv.ServerName)</td><td>$(EscHtml $db.SkuName)</td><td>$pill</td></tr>`n")
    }
}
$dbSection = ''
if ($totalDbCount -gt 0) {
    $dbSection = @"
<div class="tk-section">
  <div class="tk-card-header">
    <span class="tk-section-title">SQL Database Inventory</span>
    <span class="tk-section-num">Section 6</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Database Name</th><th>Server</th><th>SKU / Tier</th><th>Classification</th></tr></thead>
      <tbody>$($dbRows.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# -- Backup Coverage -----------------------------------------------------------

$bkpRows = [System.Text.StringBuilder]::new()
foreach ($vm in $vms) {
    $item = $backupRows | Where-Object { $_.VMName -ieq $vm.Name } | Select-Object -First 1
    if ($item) {
        $statusClass = if ($item.Status -eq 'Healthy') { 'tk-badge-ok' } elseif ($item.Status) { 'tk-badge-warn' } else { 'tk-badge-info' }
        [void]$bkpRows.Append("<tr><td><strong>$(EscHtml $vm.Name)</strong></td><td><span class='$statusClass'>$(EscHtml $item.Status)</span></td><td>$(EscHtml $item.Vault)</td><td>$(EscHtml $item.LastBackup)</td><td>$(EscHtml $item.PolicyName)</td></tr>`n")
    } else {
        [void]$bkpRows.Append("<tr><td><strong>$(EscHtml $vm.Name)</strong></td><td><span class='tk-badge-err'>Not Protected</span></td><td> - </td><td> - </td><td> - </td></tr>`n")
    }
}
$backupSection = ''
if ($vms.Count -gt 0) {
    $backupSection = @"
<div class="tk-section">
  <div class="tk-card-header">
    <span class="tk-section-title">VM Backup Coverage</span>
    <span class="tk-section-num">Section 7</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>VM Name</th><th>Protection Status</th><th>Vault</th><th>Last Backup</th><th>Policy</th></tr></thead>
      <tbody>$($bkpRows.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# -- Advisor -------------------------------------------------------------------

# Schema-agnostic field extraction for Az.Advisor recommendations.
# Az.Advisor 1.x used nested ShortDescription.Problem / ResourceId;
# Az.Advisor 2.x flattened to ShortDescriptionProblem / ResourceMetadataResourceId;
# some builds expose ResourceMetadata.ResourceId or ImpactedValue instead.
function Get-AdvisorProblemText {
    param($Rec)
    foreach ($name in 'ShortDescriptionProblem','Description','Problem') {
        $p = $Rec.PSObject.Properties[$name]
        if ($p -and $p.Value -and "$($p.Value)".Trim()) { return "$($p.Value)" }
    }
    # Nested: ShortDescription.Problem
    $sd = $Rec.PSObject.Properties['ShortDescription']
    if ($sd -and $sd.Value -and $sd.Value.PSObject.Properties['Problem'] -and $sd.Value.Problem) {
        return "$($sd.Value.Problem)"
    }
    return ''
}
function Get-AdvisorResourceId {
    param($Rec)
    foreach ($name in 'ResourceMetadataResourceId','ResourceId','ImpactedValue') {
        $p = $Rec.PSObject.Properties[$name]
        if ($p -and $p.Value -and "$($p.Value)".Trim()) { return "$($p.Value)" }
    }
    # Nested: ResourceMetadata.ResourceId
    $rm = $Rec.PSObject.Properties['ResourceMetadata']
    if ($rm -and $rm.Value -and $rm.Value.PSObject.Properties['ResourceId'] -and $rm.Value.ResourceId) {
        return "$($rm.Value.ResourceId)"
    }
    return ''
}

$advisorSection = ''
if ($advisorRecs.Count -gt 0) {
    # One-shot diagnostic: if the first record gives us neither description nor
    # resource id, the schema is unknown - surface the actual property names so
    # we can extend the extractor rather than silently rendering blank cells.
    $sample = $advisorRecs[0]
    if (-not (Get-AdvisorProblemText $sample) -and -not (Get-AdvisorResourceId $sample)) {
        $propList = ($sample.PSObject.Properties | ForEach-Object { $_.Name }) -join ', '
        Write-Warn "Advisor recommendation schema not recognized. Properties on first record:"
        Write-Info "  $propList"
    }

    $advisorRows = [System.Text.StringBuilder]::new()
    foreach ($rec in ($advisorRecs | Sort-Object -Property @{E={switch($_.Impact){'High'{0}'Medium'{1}default{2}}}},Category | Select-Object -First 50)) {
        $impactClass = switch ($rec.Impact) { 'High' { 'tk-badge-err' } 'Medium' { 'tk-badge-warn' } default { 'tk-badge-ok' } }

        $problemText    = Get-AdvisorProblemText $rec
        $resourceIdText = Get-AdvisorResourceId  $rec

        $problem = if ($problemText)    { EscHtml $problemText }                         else { '<span class="tk-badge-info">(no description)</span>' }
        $resName = if ($resourceIdText) { EscHtml (($resourceIdText -split '/')[-1]) }   else { '-' }
        $cat     = EscHtml $rec.Category
        [void]$advisorRows.Append("<tr><td>$problem</td><td>$cat</td><td><span class='$impactClass'>$(EscHtml $rec.Impact)</span></td><td>$resName</td></tr>`n")
    }
    $advisorSection = @"
<div class="tk-section">
  <div class="tk-card-header">
    <span class="tk-section-title">Azure Advisor Recommendations</span>
    <span class="tk-section-num">Section 8</span>
  </div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Recommendation</th><th>Category</th><th>Impact</th><th>Resource</th></tr></thead>
      <tbody>$($advisorRows.ToString())</tbody>
    </table>
  </div>
</div>
"@
}

# -----------------------------------------------------------------------------
# ISSUES
# -----------------------------------------------------------------------------

$issues = [System.Collections.Generic.List[hashtable]]::new()

# Security issues (highest priority)
if ($exposedNsgRules.Count -gt 0) {
    $ruleList = ($exposedNsgRules | ForEach-Object { "$($_.NSG)/$($_.Rule) (port $($_.Port))" }) -join '; '
    $issues.Add(@{ Sev='high'; Title="$($exposedNsgRules.Count) NSG Rule(s) Exposing Sensitive Ports to the Internet"
        Body="The following NSG rules allow inbound access from any source on sensitive ports: $ruleList. This is a direct attack surface for brute-force, ransomware, and lateral movement. Use Azure Bastion for RDP/SSH instead." })
}
if ($sqlFirewallIssues.Count -gt 0) {
    $fwList = ($sqlFirewallIssues | ForEach-Object { "$(EscHtml $_.Server) (rule: $(EscHtml $_.Rule))" }) -join ', '
    $issues.Add(@{ Sev='high'; Title="$($sqlFirewallIssues.Count) SQL Server(s) With Permissive Firewall Rules"
        Body="The following SQL servers have firewall rules that allow connections from any IP address: $fwList. This exposes SQL authentication to the public internet. Restrict firewall rules to known IP ranges or use private endpoints." })
}
if ($publicStorageAccounts.Count -gt 0) {
    $saList = ($publicStorageAccounts | ForEach-Object { EscHtml $_.StorageAccountName }) -join ', '
    $issues.Add(@{ Sev='high'; Title="$($publicStorageAccounts.Count) Storage Account(s) With Public Blob Access Not Disabled"
        Body="The following storage accounts have not explicitly disabled public blob access: $saList. If any container is accidentally set to public, its data is accessible without authentication. Set AllowBlobPublicAccess=false on all accounts unless public hosting is intentional." })
}
if ($httpApps.Count -gt 0) {
    $appList = ($httpApps | ForEach-Object { EscHtml $_.Name }) -join ', '
    $issues.Add(@{ Sev='high'; Title="$($httpApps.Count) App Service Site(s) Not Enforcing HTTPS"
        Body="The following sites do not have HTTPS-only enforced: $appList. Credentials, session tokens, and sensitive data can be transmitted in plaintext. Enable the HTTPS Only setting on each App Service." })
}
if ($unlockedCritical.Count -gt 0) {
    $resList = ($unlockedCritical | Select-Object -First 8 | ForEach-Object { EscHtml $_.Name }) -join ', '
    $issues.Add(@{ Sev='high'; Title="$($unlockedCritical.Count) Critical Resource(s) Have No Delete Lock"
        Body="The following production-class resources have no CanNotDelete or ReadOnly lock: $resList$(if($unlockedCritical.Count -gt 8){', and more'}). A single accidental deletion or malicious action could permanently destroy data. Apply CanNotDelete locks to all VMs, SQL servers, key vaults, and storage accounts." })
}

# Governance issues
if ($highAdvisor -gt 0) {
    $issues.Add(@{ Sev='high'; Title="$highAdvisor Unresolved High-Impact Azure Advisor Alerts"
        Body="Azure Advisor is flagging $highAdvisor High-severity and $medAdvisor Medium issues from live telemetry  -  Microsoft's own assessment of availability, security, and cost risk in this subscription." })
}
if ($untaggedPct -gt 50) {
    $issues.Add(@{ Sev='high'; Title="~$untaggedPct% of Resources Have No Tags"
        Body="Only $tagCoverage% of $($allResources.Count) resources are tagged. Without Environment, Owner, Application, and CostCenter tags there is no cost allocation, ownership tracking, or lifecycle enforcement." })
}
if ($adHocDbs.Count -gt 3) {
    $sample = (($adHocDbs | Select-Object -First 5 | Select-Object -ExpandProperty DatabaseName) | ForEach-Object { EscHtml $_ }) -join ', '
    $issues.Add(@{ Sev='high'; Title="$($adHocDbs.Count) Ad-Hoc / Date-Stamped Databases Detected"
        Body="Databases matching backup/copy/date patterns: $sample$(if($adHocDbs.Count -gt 5){', and more'}). Manual copies are risky and costly  -  use Azure SQL point-in-time restore instead." })
}
if ($subOwners.Count -gt 3) {
    $ownerList = ($subOwners | ForEach-Object { EscHtml $_.DisplayName }) -join ', '
    $issues.Add(@{ Sev='high'; Title="$($subOwners.Count) Principals Hold Owner at Subscription Scope"
        Body="Subscription-level Owners: $ownerList. Owner is the highest privilege level  -  any of these accounts being compromised grants full control over every resource in the subscription. Reduce to the minimum required and prefer Contributor for day-to-day work." })
}

# Operational issues
if ($vmsWithoutMon.Count -gt 0) {
    $vmList = ($vmsWithoutMon | ForEach-Object { EscHtml $_ }) -join ', '
    $issues.Add(@{ Sev='med'; Title="$($vmsWithoutMon.Count) VMs Have No Extensions Deployed"
        Body="VMs without any extensions: $vmList. No diagnostics, no Log Analytics  -  if these VMs fail or degrade, there is no telemetry to diagnose the cause." })
}
if ($unbackedVms.Count -gt 0) {
    $vmList = ($unbackedVms | ForEach-Object { EscHtml $_ }) -join ', '
    $issues.Add(@{ Sev='med'; Title="$($unbackedVms.Count) VM(s) Have No Confirmed Backup Protection"
        Body="The following VMs were not found in any Recovery Services vault backup policy: $vmList. If these VMs are lost or corrupted, no recovery point exists." })
}
if ($totalPolicyStates -gt 0) {
    $issues.Add(@{ Sev='med'; Title="$totalPolicyStates Non-Compliant Azure Policy State(s) Detected"
        Body="Resources in this subscription are failing assigned Azure Policy rules. Review the Access &amp; Governance section for details and resolve non-compliant states through the Azure Portal under Policy &gt; Compliance." })
}
if ($regions.Count -gt 2) {
    $issues.Add(@{ Sev='med'; Title="Resources Spread Across $($regions.Count) Regions"
        Body="Regions in use: $(($regions | ForEach-Object { EscHtml $_ }) -join ', '). Multi-region spread without a deliberate geo-redundancy design adds egress cost and management complexity." })
}
if ($appServicePlans.Count -gt 2) {
    $issues.Add(@{ Sev='med'; Title="$($appServicePlans.Count) App Service Plans  -  Verify All Are Active"
        Body="Plans: $(( $appServicePlans | Select-Object -ExpandProperty Name | ForEach-Object { EscHtml $_ }) -join ', '). Idle plans incur cost. Auto-generated or timestamp-named plans are likely from abandoned deployments." })
}
if ($recoveryVaults.Count -gt 1) {
    $issues.Add(@{ Sev='med'; Title="$($recoveryVaults.Count) Recovery Vaults  -  Ownership Unclear"
        Body="Multiple vaults: $(($recoveryVaults | Select-Object -ExpandProperty Name | ForEach-Object { EscHtml $_ }) -join ', '). Verify which vault protects which VMs and consolidate if possible." })
}
if ($orphanedDisks.Count -gt 0) {
    $diskList = ($orphanedDisks | ForEach-Object { "$(EscHtml $_.Name) ($($_.DiskSizeGB) GB)" }) -join ', '
    $issues.Add(@{ Sev='low'; Title="$($orphanedDisks.Count) Unattached Managed Disk(s)"
        Body="Billed at idle: $diskList. Snapshot if needed, then delete." })
}
if ($unassociatedIPs.Count -gt 0) {
    $ipList = (($unassociatedIPs | Select-Object -ExpandProperty Name) | ForEach-Object { EscHtml $_ }) -join ', '
    $issues.Add(@{ Sev='low'; Title="$($unassociatedIPs.Count) Unassociated Public IP Address(es)"
        Body="IPs not assigned to any resource: $ipList. Static public IPs are charged when idle." })
}
if ($issues.Count -eq 0) {
    $issues.Add(@{ Sev='low'; Title="No Major Issues Detected"
        Body="No significant governance or security issues were automatically detected. Continue monitoring Azure Advisor and Defender for Cloud." })
}

$issueHtml = [System.Text.StringBuilder]::new()
foreach ($issue in $issues) {
    $sevLabel = switch ($issue.Sev) { 'high' { 'High' } 'med' { 'Med' } default { 'Low' } }
    $sevBadge = switch ($issue.Sev) { 'high' { 'tk-badge-err' } 'med' { 'tk-badge-warn' } default { 'tk-badge-ok' } }
    [void]$issueHtml.Append("<div class='tk-info-box'><span class='$sevBadge'>$sevLabel</span> <strong>$(EscHtml $issue.Title)</strong><p>$($issue.Body)</p></div>`n")
}

# -----------------------------------------------------------------------------
# RECOMMENDATIONS
# -----------------------------------------------------------------------------

$recs = [System.Collections.Generic.List[hashtable]]::new()

if ($exposedNsgRules.Count -gt 0) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Remove Internet-Exposed NSG Rules  -  Use Azure Bastion'
        Body="Replace inbound NSG rules that allow RDP/SSH from any source with Azure Bastion for secure browser-based access. Azure Bastion is already deployed in this subscription  -  use it and remove the open NSG rules." })
}
if ($sqlFirewallIssues.Count -gt 0) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Restrict SQL Server Firewall Rules'
        Body="Remove allow-all (0.0.0.0-255.255.255.255) SQL firewall rules. Replace with specific IP ranges or, better, use private endpoints with network-level isolation to remove SQL from the public internet entirely." })
}
if ($publicStorageAccounts.Count -gt 0) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Disable Public Blob Access on All Storage Accounts'
        Body="Set AllowBlobPublicAccess=false on all storage accounts unless public static website hosting is explicitly required. This prevents accidental exposure if a container is ever misconfigured to public." })
}
if ($httpApps.Count -gt 0) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Enable HTTPS-Only on All App Service Sites'
        Body="Toggle the HTTPS Only setting to On for: $(($httpApps | Select-Object -ExpandProperty Name | ForEach-Object { EscHtml $_ }) -join ', '). This takes seconds per app and eliminates plaintext credential transmission." })
}
if ($unlockedCritical.Count -gt 0) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Apply CanNotDelete Locks to Critical Resources'
        Body="Add a CanNotDelete resource lock to all VMs, SQL servers, key vaults, and storage accounts. This takes minutes to implement and prevents accidental or malicious deletion without significantly impacting operations." })
}
if ($untaggedPct -gt 20) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Implement a Tagging Policy'
        Body="Apply mandatory tags  -  at minimum <em>Environment</em>, <em>Application</em>, <em>Owner</em>, and <em>CostCenter</em>  -  to all resources. Use Azure Policy with a Deny effect to enforce tags on new resources. Only $tagCoverage% of resources are currently tagged." })
}
if ($highAdvisor -gt 0) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title="Resolve $highAdvisor High-Impact Advisor Alerts"
        Body="Work through High-severity Advisor items first. These are Microsoft's direct assessment of availability and security risk based on live telemetry." })
}
if ($adHocDbs.Count -gt 3) {
    $recs.Add(@{ P='immediate'; L='Immediate'; Title='Replace Ad-Hoc Database Copies with Point-in-Time Restore'
        Body="Audit the $($adHocDbs.Count) date-stamped databases, archive required data, and delete the rest. Azure SQL automated backups with point-in-time restore are already included in the service cost and are a safer approach." })
}
if ($subOwners.Count -gt 3) {
    $recs.Add(@{ P='short'; L='Short-Term'; Title='Reduce Subscription-Level Owner Assignments'
        Body="$($subOwners.Count) principals currently hold Owner at subscription scope. Audit each assignment  -  downgrade to Contributor where full Owner is not required, and prefer just-in-time access (PIM) for break-glass scenarios." })
}
if ($unbackedVms.Count -gt 0) {
    $recs.Add(@{ P='short'; L='Short-Term'; Title='Enroll All VMs in a Backup Policy'
        Body="Add the following VMs to a Recovery Services vault backup policy: $(($unbackedVms | ForEach-Object { EscHtml $_ }) -join ', '). Daily backup with 30-day retention is a reasonable baseline for production VMs." })
}
if ($appServicePlans.Count -gt 2) {
    $recs.Add(@{ P='short'; L='Short-Term'; Title='Consolidate App Service Plans'
        Body="Identify which plan hosts production workloads, migrate all apps onto one or two appropriately-sized plans, and delete the rest. Auto-generated and timestamp-named plans are safe to remove after confirming no apps are assigned to them." })
}
if ($vmsWithoutMon.Count -gt 0) {
    $recs.Add(@{ P='short'; L='Short-Term'; Title='Deploy Azure Monitor Agent to All VMs'
        Body="Deploy monitoring to: $(($vmsWithoutMon | ForEach-Object { EscHtml $_ }) -join ', '). Connect to a central Log Analytics workspace and configure CPU, disk, and memory alerts." })
}
if ($orphanedDisks.Count -gt 0 -or $unassociatedIPs.Count -gt 0) {
    $recs.Add(@{ P='short'; L='Short-Term'; Title='Remove Orphaned Resources'
        Body="Delete $($orphanedDisks.Count) unattached disk(s) and $($unassociatedIPs.Count) unassociated public IP(s)  -  ongoing cost with no operational value." })
}
if ($recoveryVaults.Count -gt 1) {
    $recs.Add(@{ P='short'; L='Short-Term'; Title='Consolidate Backup Vaults'
        Body="Verify which vault is actively protecting which resources, consolidate to one where possible, and test at least one full restore to confirm recoverability." })
}
if ($regions.Count -gt 2) {
    $recs.Add(@{ P='medium'; L='Medium-Term'; Title='Evaluate Region Consolidation'
        Body="$(($regions | ForEach-Object { EscHtml $_ }) -join ', ') are all in use. Consolidating to a primary region eliminates cross-region egress cost and simplifies network topology unless geo-redundancy is intentional." })
}
if ($totalPolicyStates -gt 0) {
    $recs.Add(@{ P='medium'; L='Medium-Term'; Title='Resolve Azure Policy Non-Compliance'
        Body="$totalPolicyStates resources are failing policy assignments. Review each item in Azure Portal under Policy &gt; Compliance and remediate or document accepted risk for each." })
}
$recs.Add(@{ P='long'; L='Long-Term'; Title='Formalize Dev / Test / Prod Environment Boundaries'
    Body='Establish separate resource groups or subscriptions with naming conventions and Azure Policy enforcement per environment tier. This prevents test resources from landing in production groups and enables clean cost reporting per environment.' })

$recHtml = [System.Text.StringBuilder]::new()
foreach ($rec in $recs) {
    $recBadge = switch ($rec.P) { 'immediate' { 'tk-badge-err' } 'short' { 'tk-badge-warn' } 'medium' { 'tk-badge-blue' } default { 'tk-badge-info' } }
    [void]$recHtml.Append("<div class='tk-info-box'><span class='$recBadge'>$(EscHtml $rec.L)</span> <strong>$(EscHtml $rec.Title)</strong><p>$($rec.Body)</p></div>`n")
}

# -----------------------------------------------------------------------------
# ASSEMBLE HTML
# -----------------------------------------------------------------------------

$reportDate = Get-Date -Format "MMMM d, yyyy"
$orgName    = EscHtml $subName
$subIdEsc   = EscHtml $subId
$regionDisp = ($regions | ForEach-Object { EscHtml $_ }) -join ', '

$defenderStatBox = if ($defenderPct -ne $null) {
    $dClass = if ($defenderPct -ge 70) { 'ok' } elseif ($defenderPct -ge 40) { 'warn' } else { 'err' }
    "<div class='tk-summary-card $dClass'><div class='tk-summary-num'>$defenderPct%</div><div class='tk-summary-lbl'>Defender Secure Score</div></div>"
} else {
    "<div class='tk-summary-card info'><div class='tk-summary-num'>N/A</div><div class='tk-summary-lbl'>Defender Secure Score</div></div>"
}

$htmlHead = Get-TKHtmlHead `
    -Title      "Azure Environment Assessment -- $orgName" `
    -ScriptName 'T.A.L.I.S.M.A.N.' `
    -Subtitle   "$orgName -- Cloud Infrastructure Review" `
    -MetaItems  ([ordered]@{
        'Report Date'     = $reportDate
        'Subscription'    = $orgName
        'Subscription ID' = $subIdEsc
        'Total Resources' = "$($allResources.Count)"
        'Region(s)'       = $regionDisp
    }) `
    -NavItems   @(
        'Executive Summary',
        'Services in Use',
        'Issues',
        'Security Posture',
        'Access & Governance',
        'Virtual Machines',
        'SQL Databases',
        'Backup Coverage',
        'Advisor',
        'Recommendations'
    )

$htmlFoot = Get-TKHtmlFoot -ScriptName 'T.A.L.I.S.M.A.N. v2.2'

$html = $htmlHead + @"

  <!-- Executive Summary -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Executive Summary</span>
      <span class="tk-section-num">Overview</span>
    </div>
    <div class="tk-card">
      <div class="tk-summary-row">
        <div class="tk-summary-card info"><div class="tk-summary-num">$($vms.Count)</div><div class="tk-summary-lbl">Virtual Machines</div></div>
        <div class="tk-summary-card info"><div class="tk-summary-num">$($webApps.Count)</div><div class="tk-summary-lbl">Web Apps &amp; Functions</div></div>
        <div class="tk-summary-card info"><div class="tk-summary-num">$($sqlServers.Count)</div><div class="tk-summary-lbl">SQL Servers</div></div>
        <div class="tk-summary-card warn"><div class="tk-summary-num">$totalDbCount</div><div class="tk-summary-lbl">SQL Databases (incl. copies)</div></div>
        <div class="tk-summary-card $(if($exposedNsgRules.Count -gt 0){'err'}else{'ok'})"><div class="tk-summary-num">$($exposedNsgRules.Count)</div><div class="tk-summary-lbl">Internet-Exposed NSG Rules</div></div>
        <div class="tk-summary-card $(if($unbackedVms.Count -gt 0){'err'}else{'ok'})"><div class="tk-summary-num">$($unbackedVms.Count)</div><div class="tk-summary-lbl">Unprotected VMs (Backup)</div></div>
        <div class="tk-summary-card err"><div class="tk-summary-num">$highAdvisor</div><div class="tk-summary-lbl">High-Impact Advisor Alerts</div></div>
        <div class="tk-summary-card $(if($untaggedPct -gt 50){'err'}elseif($untaggedPct -gt 20){'warn'}else{'ok'})"><div class="tk-summary-num">~$untaggedPct%</div><div class="tk-summary-lbl">Resources Without Tags</div></div>
        <div class="tk-summary-card $(if($unlockedCritical.Count -gt 0){'err'}else{'ok'})"><div class="tk-summary-num">$($unlockedCritical.Count)</div><div class="tk-summary-lbl">Critical Resources Unlocked</div></div>
        <div class="tk-summary-card $(if($orphanedDisks.Count -gt 0){'warn'}else{'ok'})"><div class="tk-summary-num">$($orphanedDisks.Count)</div><div class="tk-summary-lbl">Unattached Disks</div></div>
        <div class="tk-summary-card info"><div class="tk-summary-num">$($subOwners.Count)</div><div class="tk-summary-lbl">Subscription Owners</div></div>
        $defenderStatBox
      </div>
      <div class="tk-info-box" style="margin-top:20px">
        This assessment covers <strong>$orgName</strong> -- $($allResources.Count) resources across
        $($resourceGroups.Count) resource groups and $($regions.Count) region(s).
        $(if($exposedNsgRules.Count -gt 0){"<span class='tk-badge-err'>$($exposedNsgRules.Count) NSG rule(s) expose sensitive ports to the internet</span> and require immediate attention. "})$(if($unlockedCritical.Count -gt 0){"$($unlockedCritical.Count) critical resources have no delete lock. "})The environment has $highAdvisor high-impact Advisor alerts, $($adHocDbs.Count) ad-hoc database copies, and approximately $untaggedPct% of resources without tags.
      </div>
    </div>
  </div>

  <!-- Section 1: Services in Use -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Azure Services Currently in Use</span>
      <span class="tk-section-num">Section 1</span>
    </div>
    <div class="tk-card">
      $($serviceItems.ToString())
    </div>
  </div>

  <!-- Section 2: Issues -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Operational Limitations &amp; Issues</span>
      <span class="tk-section-num">Section 2</span>
    </div>
    <div class="tk-card">
      $($issueHtml.ToString())
    </div>
  </div>

  $securitySection

  $governanceSection

  $vmSection

  $dbSection

  $backupSection

  $advisorSection

  <!-- Recommendations -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Recommended Improvements</span>
      <span class="tk-section-num">Recommendations</span>
    </div>
    <div class="tk-card">
      $($recHtml.ToString())
    </div>
  </div>

"@ + $htmlFoot

# -----------------------------------------------------------------------------
# OUTPUT
# -----------------------------------------------------------------------------

$html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force

Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
Write-Host ""
Write-Ok "Report saved: $OutputPath"

if (-not $NoOpen) {
    Write-Step "Opening in default browser..."
    Start-Process $OutputPath
}

Write-Host ""
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
