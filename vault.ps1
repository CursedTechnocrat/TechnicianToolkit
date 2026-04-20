<#
.SYNOPSIS
    V.A.U.L.T. — Validates Assets & User License Tracking
    Microsoft 365 License & Mailbox Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft 365 via the Microsoft Graph API to audit license
    assignments, identify unlicensed and inactive users, check MFA registration
    status, and audit shared mailboxes. The required Microsoft.Graph module is
    installed automatically if missing. Results can be reviewed interactively
    in the console or exported as a combined dark-themed HTML report.

.USAGE
    PS C:\> .\vault.ps1                    # Interactive menu
    PS C:\> .\vault.ps1 -Unattended        # Auto-connect and export full audit report

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    P.H.A.N.T.O.M.         — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    S.I.G.I.L.             — Security baseline & policy enforcement
    S.P.E.C.T.E.R.         — Remote machine execution via WinRM
    L.E.Y.L.I.N.E.         — Network diagnostics & remediation
    F.O.R.G.E.             — Driver update detection & installation
    A.E.G.I.S.             — Azure environment assessment & reporting
    B.A.S.T.I.O.N.         — Active Directory & identity management
    L.A.N.T.E.R.N.         — Network discovery & asset inventory
    T.H.R.E.S.H.O.L.D.     — Disk & storage health monitoring
    V.A.U.L.T.             — M365 license & mailbox auditing
    S.E.N.T.I.N.E.L.       — Service & scheduled task monitoring
    R.E.L.I.C.             — Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.           — Toolkit setup & configuration wizard

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success / compliant
    Yellow   Warnings / attention needed
    Red      Critical issues
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ENCODING & SCRIPT PATH
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force

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

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
}

# ─────────────────────────────────────────────────────────────────────────────
# CONNECTION STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:Connected      = $false
$script:ConnectedAs    = ''
$script:TenantDomain   = ''

# ─────────────────────────────────────────────────────────────────────────────
# SKU FRIENDLY NAME MAP
# ─────────────────────────────────────────────────────────────────────────────

$SkuNames = @{
    'SPE_E3'                          = 'Microsoft 365 E3'
    'SPE_E5'                          = 'Microsoft 365 E5'
    'SPE_F1'                          = 'Microsoft 365 F3'
    'O365_BUSINESS_PREMIUM'           = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS'        = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                   = 'Microsoft 365 Apps for Business'
    'ENTERPRISEPACK'                  = 'Office 365 E3'
    'ENTERPRISEPREMIUM'               = 'Office 365 E5'
    'STANDARDPACK'                    = 'Office 365 E1'
    'DESKLESSPACK'                    = 'Office 365 F3'
    'AAD_PREMIUM'                     = 'Azure AD Premium P1'
    'AAD_PREMIUM_P2'                  = 'Azure AD Premium P2'
    'EMS'                             = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                      = 'Enterprise Mobility + Security E5'
    'POWER_BI_PRO'                    = 'Power BI Pro'
    'POWER_BI_PREMIUM_PER_USER'       = 'Power BI Premium Per User'
    'FLOW_FREE'                       = 'Power Automate Free'
    'POWERFLOW_P1'                    = 'Power Automate Plan 1'
    'INTUNE_A'                        = 'Microsoft Intune'
    'INTUNE_A_D'                      = 'Microsoft Intune Device'
    'PROJECTPREMIUM'                  = 'Project Plan 5'
    'PROJECTPROFESSIONAL'             = 'Project Plan 3'
    'VISIOCLIENT'                     = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'               = 'Visio Plan 1'
    'TEAMS_EXPLORATORY'               = 'Microsoft Teams Exploratory'
    'TEAMS_FREE'                      = 'Microsoft Teams Free'
    'MCOSTANDARD'                     = 'Skype for Business Online Plan 2'
    'EXCHANGESTANDARD'                = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'              = 'Exchange Online Plan 2'
    'EXCHANGE_S_DESKLESS'             = 'Exchange Online Kiosk'
    'DEFENDER_ENDPOINT_P1'            = 'Microsoft Defender for Endpoint P1'
    'WIN_DEF_ATP'                     = 'Microsoft Defender for Endpoint P2'
    'RIGHTSMANAGEMENT'                = 'Azure Information Protection P1'
    'MCOMEETADV'                      = 'Microsoft Teams Audio Conferencing'
    'PHONESYSTEM_VIRTUALUSER'         = 'Microsoft Teams Phone Resource Account'
    'MCOPSTN1'                        = 'Microsoft Teams Domestic Calling Plan'
    'M365_LIGHTHOUSE_PARTNER_PLAN1'   = 'Microsoft 365 Lighthouse'
    'WINDOWS_STORE'                   = 'Windows Store for Business'
    'DEVELOPERPACK_E3'                = 'Microsoft 365 E3 Developer'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-VaultBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗   ██╗ █████╗ ██╗   ██╗██╗  ████████╗
  ██║   ██║██╔══██╗██║   ██║██║  ╚══██╔══╝
  ██║   ██║███████║██║   ██║██║     ██║
  ╚██╗ ██╔╝██╔══██║██║   ██║██║     ██║
   ╚████╔╝ ██║  ██║╚██████╔╝███████╗██║
    ╚═══╝  ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝

"@ -ForegroundColor Cyan
    Write-Host "  V.A.U.L.T. — Validates Assets & User License Tracking" -ForegroundColor Cyan
    Write-Host "  Microsoft 365 License & Mailbox Audit Tool  v1.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE CHECK & INSTALL
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Reports'
    )

    $needsInstall = $false
    foreach ($mod in $requiredModules) {
        if (Get-Module -ListAvailable -Name $mod) {
            Write-Ok "$mod — installed"
        } else {
            Write-Warn "$mod — NOT found"
            $needsInstall = $true
        }
    }

    if ($needsInstall) {
        Write-Host ""
        if ($Unattended) {
            Write-Step "Installing Microsoft.Graph module (unattended)..."
            try {
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Ok "Microsoft.Graph installed successfully."
            } catch {
                Write-Fail "Failed to install Microsoft.Graph: $_"
                exit 1
            }
        } else {
            Write-Warn "One or more Microsoft.Graph sub-modules are not installed."
            Write-Info "The Microsoft.Graph module suite is required for V.A.U.L.T. to operate."
            Write-Host ""
            $ans = Read-Host "  Install Microsoft.Graph for current user? [Y/N]"
            if ($ans -match '^[Yy]') {
                Write-Step "Installing Microsoft.Graph — this may take a minute..."
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Ok "Microsoft.Graph installed successfully."
                } catch {
                    Write-Fail "Installation failed: $_"
                    Write-Info "Run manually: Install-Module Microsoft.Graph -Scope CurrentUser -Force"
                    exit 1
                }
            } else {
                Write-Fail "Microsoft.Graph is required. Exiting."
                exit 1
            }
        }
    } else {
        Write-Ok "All required Graph modules are available."
    }

    # Import the sub-modules
    foreach ($mod in $requiredModules) {
        try {
            Import-Module $mod -ErrorAction Stop
        } catch {
            Write-Warn "Could not import ${mod}: $_"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

$GraphScopes = @(
    'User.Read.All',
    'Directory.Read.All',
    'AuditLog.Read.All',
    'Reports.Read.All'
)

function Test-GraphConnection {
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account) {
            $script:Connected   = $true
            $script:ConnectedAs = $ctx.Account
            if ($ctx.TenantId) {
                # Try to get the verified domain for display
                try {
                    $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($org -and $org.VerifiedDomains) {
                        $primary = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
                        if ($primary) { $script:TenantDomain = $primary.Name }
                    }
                } catch { }
            }
            return $true
        }
    } catch { }
    $script:Connected   = $false
    $script:ConnectedAs = ''
    return $false
}

function Invoke-Connect {
    Write-Section "CONNECT TO MICROSOFT 365"
    Write-Step "Initiating interactive authentication..."
    Write-Info "Requested scopes: $($GraphScopes -join ', ')"
    Write-Host ""

    try {
        if ($Unattended) {
            # Attempt silent connect using existing token cache
            Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        } else {
            Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        }

        if (Test-GraphConnection) {
            Write-Ok "Connected as: $($script:ConnectedAs)"
            if ($script:TenantDomain) {
                Write-Ok "Tenant domain: $($script:TenantDomain)"
            }
        } else {
            Write-Fail "Connection appeared to succeed but context could not be verified."
        }
    } catch {
        $errMsg = $_.Exception.Message
        if ($Unattended) {
            Write-Fail "Silent connect failed. An existing auth token is required for unattended mode."
            Write-Info "Run the script interactively first to cache credentials, or use a service principal."
            Write-Info "Error: $errMsg"
            exit 1
        } else {
            Write-Fail "Authentication failed: $errMsg"
        }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# AUDIT FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Get-LicenseInventory {
    Write-Section "LICENSE SKU INVENTORY"
    Write-Step "Querying subscribed SKUs..."

    try {
        $skus = Get-MgSubscribedSku -All -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve license data: $_"
        return @()
    }

    if (-not $skus -or $skus.Count -eq 0) {
        Write-Warn "No SKUs found. The tenant may have no paid subscriptions."
        return @()
    }

    $inventory = foreach ($sku in $skus) {
        $partNumber  = $sku.SkuPartNumber
        $friendlyName = if ($SkuNames.ContainsKey($partNumber)) { $SkuNames[$partNumber] } else { $partNumber }
        $total     = $sku.PrepaidUnits.Enabled
        $assigned  = $sku.ConsumedUnits
        $available = $total - $assigned

        [PSCustomObject]@{
            SkuPartNumber = $partNumber
            FriendlyName  = $friendlyName
            Assigned      = $assigned
            Total         = $total
            Available     = $available
            CapabilityStatus = $sku.CapabilityStatus
        }
    }

    # Console display
    $colWidth = @(38, 10, 10, 12)
    Write-Host ("  {0,-38} {1,10} {2,10} {3,12}" -f "SKU Name", "Assigned", "Total", "Available") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 74)) -ForegroundColor $C.Info

    foreach ($sku in ($inventory | Sort-Object FriendlyName)) {
        $availColor = if ($sku.Available -lt 0) { $C.Error }
                      elseif ($sku.Available -le 5) { $C.Warning }
                      else { $C.Success }
        $statusTag  = if ($sku.CapabilityStatus -ne 'Enabled') { " [$($sku.CapabilityStatus)]" } else { '' }
        Write-Host ("  {0,-38} {1,10} {2,10} {3,12}" -f `
            ($sku.FriendlyName + $statusTag), $sku.Assigned, $sku.Total, $sku.Available) -ForegroundColor $availColor
    }

    Write-Host ""
    Write-Ok "$($inventory.Count) SKU(s) retrieved."
    return @($inventory)
}

function Get-UnlicensedUsers {
    Write-Section "UNLICENSED USERS"
    Write-Step "Querying enabled users with no license assigned..."
    Write-Info "This uses ConsistencyLevel=eventual — results may be slightly cached."
    Write-Host ""

    try {
        $users = Get-MgUser `
            -Filter "assignedLicenses/`$count eq 0 and accountEnabled eq true" `
            -CountVariable licCount `
            -ConsistencyLevel eventual `
            -All `
            -Property "displayName,userPrincipalName,department,accountEnabled,userType" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Query failed: $_"
        return @()
    }

    # Exclude guests unless they show up as members
    $members = @($users | Where-Object { $_.UserType -ne 'Guest' })

    if ($members.Count -eq 0) {
        Write-Ok "No unlicensed enabled member accounts found."
        return @()
    }

    $result = foreach ($u in ($members | Sort-Object DisplayName)) {
        [PSCustomObject]@{
            DisplayName    = $u.DisplayName
            UPN            = $u.UserPrincipalName
            Department     = if ($u.Department) { $u.Department } else { '—' }
            AccountEnabled = $u.AccountEnabled
        }
    }

    # Console display
    Write-Host ("  {0,-32} {1,-40} {2,-22}" -f "Display Name", "UPN", "Department") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 96)) -ForegroundColor $C.Info
    foreach ($u in $result) {
        Write-Host ("  {0,-32} {1,-40} {2,-22}" -f $u.DisplayName, $u.UPN, $u.Department) -ForegroundColor $C.Warning
    }

    Write-Host ""
    Write-Warn "$($result.Count) unlicensed enabled user(s) found."
    return @($result)
}

function Get-InactiveUsers {
    Write-Section "INACTIVE USERS (90+ DAYS)"
    Write-Step "Retrieving all users with sign-in activity..."
    Write-Info "This may take a moment depending on tenant size. Requires AuditLog.Read.All."
    Write-Host ""

    $thresholdDate = (Get-Date).AddDays(-90)
    $allUsers      = $null

    try {
        $allUsers = Get-MgUser -All `
            -Property "displayName,userPrincipalName,department,accountEnabled,assignedLicenses,signInActivity,userType" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        Write-Info "Ensure AuditLog.Read.All is consented and the account has sufficient permissions."
        return @()
    }

    $members = @($allUsers | Where-Object { $_.UserType -ne 'Guest' -and $_.AccountEnabled -eq $true })

    $inactive = foreach ($u in $members) {
        $lastSignIn = $null
        $daysInactive = $null

        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn   = [datetime]$u.SignInActivity.LastSignInDateTime
            $daysInactive = [math]::Round(((Get-Date) - $lastSignIn).TotalDays, 0)
        }

        $isInactive = (-not $lastSignIn) -or ($lastSignIn -lt $thresholdDate)

        if ($isInactive) {
            [PSCustomObject]@{
                DisplayName   = $u.DisplayName
                UPN           = $u.UserPrincipalName
                Department    = if ($u.Department) { $u.Department } else { '—' }
                LastSignIn    = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
                DaysInactive  = if ($daysInactive -ne $null) { $daysInactive } else { 'N/A' }
                IsLicensed    = ($u.AssignedLicenses.Count -gt 0)
            }
        }
    }

    $inactive = @($inactive)

    if ($inactive.Count -eq 0) {
        Write-Ok "No inactive users found (all active within 90 days)."
        return @()
    }

    $sorted = $inactive | Sort-Object @{ Expression={ if ($_.DaysInactive -eq 'N/A') { 99999 } else { [int]$_.DaysInactive } }; Descending=$true }

    # Console display
    Write-Host ("  {0,-28} {1,-36} {2,-14} {3,-12}" -f "Display Name", "UPN", "Last Sign-In", "Days Inactive") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 92)) -ForegroundColor $C.Info
    foreach ($u in $sorted) {
        $color = if ($u.LastSignIn -eq 'Never') { $C.Error } else { $C.Warning }
        Write-Host ("  {0,-28} {1,-36} {2,-14} {3,-12}" -f $u.DisplayName, $u.UPN, $u.LastSignIn, $u.DaysInactive) -ForegroundColor $color
    }

    Write-Host ""
    Write-Warn "$($inactive.Count) inactive user(s) found (no sign-in within 90 days or never)."
    return @($inactive)
}

function Get-MfaStatus {
    Write-Section "MFA REGISTRATION STATUS"
    Write-Warn "This check calls the API once per user — it may be slow on large tenants."
    Write-Step "Retrieving enabled member users..."
    Write-Host ""

    try {
        $allUsers = Get-MgUser -All `
            -Filter "accountEnabled eq true" `
            -Property "displayName,userPrincipalName,department,userType,id" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        return @()
    }

    $members = @($allUsers | Where-Object { $_.UserType -ne 'Guest' })

    if ($members.Count -eq 0) {
        Write-Warn "No enabled member users found."
        return @()
    }

    Write-Step "Checking authentication methods for $($members.Count) user(s)..."
    Write-Info "Progress is shown below. Guests and service accounts are skipped on error."
    Write-Host ""

    $noMfa    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total    = $members.Count
    $current  = 0
    $skipped  = 0

    foreach ($u in $members) {
        $current++

        # Progress every 10 users or at key milestones
        if ($current -eq 1 -or ($current % 10) -eq 0 -or $current -eq $total) {
            $pct = [math]::Round(($current / $total) * 100, 0)
            Write-Host ("`r  [*] Checking user {0}/{1} ({2}%)..." -f $current, $total, $pct) -ForegroundColor $C.Progress -NoNewline
        }

        try {
            $methods = Get-MgUserAuthenticationMethod -UserId $u.Id -ErrorAction Stop
            # Filter out the default "password" method — only count real MFA factors
            $mfaMethods = @($methods | Where-Object {
                $_.AdditionalProperties['@odata.type'] -notmatch 'passwordAuthentication'
            })

            if ($mfaMethods.Count -eq 0) {
                $noMfa.Add([PSCustomObject]@{
                    DisplayName = $u.DisplayName
                    UPN         = $u.UserPrincipalName
                    Department  = if ($u.Department) { $u.Department } else { '—' }
                })
            }
        } catch {
            $skipped++
        }
    }

    Write-Host ""  # end the progress line
    Write-Host ""

    if ($skipped -gt 0) {
        Write-Warn "$skipped user(s) skipped (guests, service accounts, or permission errors)."
    }

    if ($noMfa.Count -eq 0) {
        Write-Ok "All checked users have at least one MFA method registered."
        return @()
    }

    # Console display
    Write-Host ("  {0,-32} {1,-42} {2,-22}" -f "Display Name", "UPN", "Department") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 98)) -ForegroundColor $C.Info
    foreach ($u in ($noMfa | Sort-Object DisplayName)) {
        Write-Host ("  {0,-32} {1,-42} {2,-22}" -f $u.DisplayName, $u.UPN, $u.Department) -ForegroundColor $C.Error
    }

    Write-Host ""
    Write-Fail "$($noMfa.Count) user(s) have NO MFA method registered."
    return @($noMfa)
}

function Show-SharedMailboxGuidance {
    Write-Section "SHARED MAILBOX AUDIT"
    Write-Warn "Shared mailbox auditing requires the Exchange Online Management module."
    Write-Host ""
    Write-Info "V.A.U.L.T. connects to Microsoft Graph, which does not expose mailbox"
    Write-Info "type details (Shared vs. User) in the same way as Exchange Online."
    Write-Host ""
    Write-Info "To audit shared mailboxes, run the following in a separate session:"
    Write-Host ""
    Write-Host "    # Install module if needed:" -ForegroundColor $C.Info
    Write-Host "    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # Connect:" -ForegroundColor $C.Info
    Write-Host "    Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # List shared mailboxes:" -ForegroundColor $C.Info
    Write-Host "    Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress,ProhibitSendQuota" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # Find shared mailboxes with a license assigned (unnecessary cost):" -ForegroundColor $C.Info
    Write-Host "    Get-Mailbox -RecipientTypeDetails SharedMailbox | Where-Object { `$_.SKUAssigned -eq `$true }" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # Check full-access permissions on all shared mailboxes:" -ForegroundColor $C.Info
    Write-Host "    Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-MailboxPermission | Where-Object { `$_.User -notlike '*SELF*' }" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Info "Shared mailboxes under 50 GB do not require a license. Larger mailboxes"
    Write-Info "require at least an Exchange Online Plan 2 or M365 E3 license."
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$LicenseData,
        [array]$UnlicensedData,
        [array]$InactiveData,
        [array]$NoMfaData
    )

    $reportDate    = Get-Date -Format "MMMM d, yyyy HH:mm"
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $connectedAs   = EscHtml $script:ConnectedAs

    $totalLicensed   = ($UnlicensedData.Count)  # unlicensed count
    $totalUsers      = 0
    try {
        $totalUsers = (Get-MgUser -Filter "accountEnabled eq true and userType eq 'Member'" -CountVariable x -ConsistencyLevel eventual -All -Property "id" -ErrorAction SilentlyContinue | Measure-Object).Count
    } catch {
        Write-Warn "Could not retrieve total user count — licensed count will show as '—'."
    }
    $licensedCount   = if ($totalUsers -gt 0) { $totalUsers - $UnlicensedData.Count } else { '—' }
    $inactiveCount   = $InactiveData.Count
    $noMfaCount      = $NoMfaData.Count

    # ── License rows ─────────────────────────────────────────────────────────
    $licRows = [System.Text.StringBuilder]::new()
    foreach ($sku in ($LicenseData | Sort-Object FriendlyName)) {
        $availClass = if ($sku.Available -lt 0) { 'badge-crit' }
                      elseif ($sku.Available -le 5) { 'badge-warn' }
                      else { 'badge-ok' }
        $statusCell = if ($sku.CapabilityStatus -ne 'Enabled') {
            "<span class='badge badge-warn'>$(EscHtml $sku.CapabilityStatus)</span>"
        } else {
            "<span class='badge badge-ok'>Active</span>"
        }
        [void]$licRows.Append("<tr><td>$(EscHtml $sku.FriendlyName)</td><td><code>$(EscHtml $sku.SkuPartNumber)</code></td><td>$($sku.Assigned)</td><td>$($sku.Total)</td><td><span class='badge $availClass'>$($sku.Available)</span></td><td>$statusCell</td></tr>`n")
    }

    # ── Unlicensed rows ───────────────────────────────────────────────────────
    $unLicRows = [System.Text.StringBuilder]::new()
    if ($UnlicensedData.Count -eq 0) {
        [void]$unLicRows.Append("<tr><td colspan='4' style='color:#2ecc71;text-align:center;'>No unlicensed enabled users found.</td></tr>")
    } else {
        foreach ($u in ($UnlicensedData | Sort-Object DisplayName)) {
            [void]$unLicRows.Append("<tr><td>$(EscHtml $u.DisplayName)</td><td>$(EscHtml $u.UPN)</td><td>$(EscHtml $u.Department)</td><td><span class='badge badge-ok'>Enabled</span></td></tr>`n")
        }
    }

    # ── Inactive rows ─────────────────────────────────────────────────────────
    $inactRows = [System.Text.StringBuilder]::new()
    if ($InactiveData.Count -eq 0) {
        [void]$inactRows.Append("<tr><td colspan='4' style='color:#2ecc71;text-align:center;'>No inactive users found.</td></tr>")
    } else {
        $sorted = $InactiveData | Sort-Object @{ Expression={ if ($_.DaysInactive -eq 'N/A') { 99999 } else { [int]$_.DaysInactive } }; Descending=$true }
        foreach ($u in $sorted) {
            $badgeClass = if ($u.LastSignIn -eq 'Never') { 'badge-crit' } else { 'badge-warn' }
            [void]$inactRows.Append("<tr><td>$(EscHtml $u.DisplayName)</td><td>$(EscHtml $u.UPN)</td><td><span class='badge $badgeClass'>$(EscHtml $u.LastSignIn)</span></td><td>$($u.DaysInactive)</td></tr>`n")
        }
    }

    # ── No-MFA rows ───────────────────────────────────────────────────────────
    $mfaRows = [System.Text.StringBuilder]::new()
    if ($NoMfaData.Count -eq 0) {
        [void]$mfaRows.Append("<tr><td colspan='3' style='color:#2ecc71;text-align:center;'>All users have MFA registered.</td></tr>")
    } else {
        foreach ($u in ($NoMfaData | Sort-Object DisplayName)) {
            [void]$mfaRows.Append("<tr><td>$(EscHtml $u.DisplayName)</td><td>$(EscHtml $u.UPN)</td><td>$(EscHtml $u.Department)</td></tr>`n")
        }
    }

    # ── Summary card colors ───────────────────────────────────────────────────
    $unlicClass  = if ($UnlicensedData.Count -gt 0) { 'warn' } else { 'ok' }
    $inactClass  = if ($InactiveData.Count   -gt 0) { 'warn' } else { 'ok' }
    $mfaClass    = if ($NoMfaData.Count      -gt 0) { 'crit' } else { 'ok' }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>V.A.U.L.T. Audit Report — $tenantDisplay</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI', Consolas, monospace; font-size: 14px; }
    header { background: linear-gradient(135deg, #0f3460 0%, #1a1a2e 100%); padding: 40px 48px 36px; border-bottom: 2px solid #00d4ff; }
    header .label { font-size: 11px; letter-spacing: 3px; text-transform: uppercase; color: #00d4ff; margin-bottom: 10px; }
    header h1 { font-size: 28px; font-weight: 700; color: #00d4ff; margin-bottom: 6px; }
    header .subtitle { font-size: 14px; color: #a0b0c8; }
    .meta { margin-top: 18px; display: flex; gap: 28px; font-size: 12px; color: #6b8cae; flex-wrap: wrap; }
    .meta span strong { color: #e0e0e0; }
    main { max-width: 1100px; margin: 0 auto; padding: 40px 32px; display: flex; flex-direction: column; gap: 36px; }
    /* Summary cards */
    .summary { display: flex; gap: 16px; flex-wrap: wrap; }
    .card-stat { background: #16213e; border: 1px solid #0f3460; border-radius: 8px; padding: 18px 24px; min-width: 150px; text-align: center; border-top: 3px solid #00d4ff; }
    .card-stat .val { font-size: 32px; font-weight: 800; color: #00d4ff; }
    .card-stat .lbl { font-size: 11px; color: #6b8cae; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
    .card-stat.ok   .val { color: #2ecc71; }
    .card-stat.warn .val { color: #f39c12; }
    .card-stat.crit .val { color: #e74c3c; }
    .card-stat.ok   { border-top-color: #2ecc71; }
    .card-stat.warn { border-top-color: #f39c12; }
    .card-stat.crit { border-top-color: #e74c3c; }
    /* Section cards */
    .section { background: #16213e; border: 1px solid #0f3460; border-radius: 8px; overflow: hidden; }
    .section-header { display: flex; align-items: center; gap: 12px; padding: 18px 24px; background: #0f2040; border-bottom: 1px solid #0f3460; }
    .section-header h2 { font-size: 15px; font-weight: 700; color: #00d4ff; }
    .section-header .count { margin-left: auto; font-size: 12px; font-weight: 700; color: #6b8cae; background: #1a1a2e; padding: 3px 10px; border-radius: 10px; }
    .section-body { padding: 20px 24px; }
    /* Tables */
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th { background: #0a1628; color: #00d4ff; padding: 10px 14px; text-align: left; font-size: 11px; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase; }
    td { padding: 9px 14px; border-bottom: 1px solid #1e2d4d; vertical-align: top; color: #d0d8e8; }
    tr:last-child td { border-bottom: none; }
    tr:hover td { background: #1e2d4d; }
    code { font-family: Consolas, monospace; font-size: 12px; color: #a0b8d0; background: #0a1628; padding: 1px 5px; border-radius: 3px; }
    /* Badges */
    .badge { display: inline-block; padding: 2px 9px; border-radius: 10px; font-size: 11px; font-weight: 700; }
    .badge-ok   { background: #1a4a2e; color: #2ecc71; }
    .badge-warn { background: #4a3a10; color: #f39c12; }
    .badge-crit { background: #4a1a1a; color: #e74c3c; }
    .badge-info { background: #1a2a4a; color: #00d4ff; }
    /* Guidance box */
    .guidance { background: #0a1628; border: 1px solid #00d4ff33; border-radius: 6px; padding: 16px 20px; color: #a0b8d0; font-size: 13px; line-height: 1.7; }
    .guidance code { color: #f39c12; }
    footer { background: #0a1628; border-top: 1px solid #0f3460; color: #3a5070; text-align: center; padding: 20px; font-size: 11px; margin-top: 8px; }
  </style>
</head>
<body>

<header>
  <div class="label">Confidential — Internal Use Only</div>
  <h1>V.A.U.L.T. Audit Report</h1>
  <div class="subtitle">Microsoft 365 License &amp; User Security Audit — $tenantDisplay</div>
  <div class="meta">
    <span><strong>Generated:</strong> $reportDate</span>
    <span><strong>Connected As:</strong> $connectedAs</span>
    <span><strong>Tenant:</strong> $tenantDisplay</span>
  </div>
</header>

<main>

  <!-- Summary -->
  <div class="summary">
    <div class="card-stat"><div class="val">$totalUsers</div><div class="lbl">Total Enabled Users</div></div>
    <div class="card-stat ok"><div class="val">$licensedCount</div><div class="lbl">Licensed Users</div></div>
    <div class="card-stat $unlicClass"><div class="val">$($UnlicensedData.Count)</div><div class="lbl">Unlicensed Users</div></div>
    <div class="card-stat $inactClass"><div class="val">$($InactiveData.Count)</div><div class="lbl">Inactive (90d+)</div></div>
    <div class="card-stat $mfaClass"><div class="val">$($NoMfaData.Count)</div><div class="lbl">No MFA Registered</div></div>
    <div class="card-stat"><div class="val">$($LicenseData.Count)</div><div class="lbl">License SKUs</div></div>
  </div>

  <!-- License Inventory -->
  <div class="section">
    <div class="section-header">
      <h2>License SKU Inventory</h2>
      <span class="count">$($LicenseData.Count) SKU(s)</span>
    </div>
    <div class="section-body">
      <table>
        <thead>
          <tr><th>SKU Name</th><th>Part Number</th><th>Assigned</th><th>Total</th><th>Available</th><th>Status</th></tr>
        </thead>
        <tbody>
          $($licRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- Unlicensed Users -->
  <div class="section">
    <div class="section-header">
      <h2>Unlicensed Enabled Users</h2>
      <span class="count">$($UnlicensedData.Count) user(s)</span>
    </div>
    <div class="section-body">
      <table>
        <thead>
          <tr><th>Display Name</th><th>UPN</th><th>Department</th><th>Account Status</th></tr>
        </thead>
        <tbody>
          $($unLicRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- Inactive Users -->
  <div class="section">
    <div class="section-header">
      <h2>Inactive Users — No Sign-In Within 90 Days</h2>
      <span class="count">$($InactiveData.Count) user(s)</span>
    </div>
    <div class="section-body">
      <table>
        <thead>
          <tr><th>Display Name</th><th>UPN</th><th>Last Sign-In</th><th>Days Inactive</th></tr>
        </thead>
        <tbody>
          $($inactRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- MFA Status -->
  <div class="section">
    <div class="section-header">
      <h2>Users Without MFA Registered</h2>
      <span class="count">$($NoMfaData.Count) user(s)</span>
    </div>
    <div class="section-body">
      <table>
        <thead>
          <tr><th>Display Name</th><th>UPN</th><th>Department</th></tr>
        </thead>
        <tbody>
          $($mfaRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- Shared Mailbox Note -->
  <div class="section">
    <div class="section-header">
      <h2>Shared Mailbox Audit</h2>
      <span class="count">Requires Exchange Online</span>
    </div>
    <div class="section-body">
      <div class="guidance">
        <strong style="color:#f39c12;">Shared mailbox auditing requires the Exchange Online Management module.</strong><br><br>
        Run the following commands in a separate PowerShell session:<br><br>
        <code>Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force</code><br>
        <code>Connect-ExchangeOnline -UserPrincipalName admin@$tenantDisplay</code><br>
        <code>Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress,ProhibitSendQuota</code><br><br>
        <strong>Key checks:</strong> Shared mailboxes under 50 GB do not require a license. Verify
        that no shared mailboxes have unnecessary licenses assigned, and audit full-access
        permissions to ensure only authorised users have access.
      </div>
    </div>
  </div>

</main>

<footer>
  V.A.U.L.T. — Validates Assets &amp; User License Tracking &nbsp;|&nbsp; $tenantDisplay &nbsp;|&nbsp; $reportDate &nbsp;|&nbsp; Confidential
</footer>

</body>
</html>
"@

    return $html
}

function Export-HtmlReport {
    param(
        [array]$LicenseData,
        [array]$UnlicensedData,
        [array]$InactiveData,
        [array]$NoMfaData
    )

    Write-Section "EXPORTING HTML REPORT"

    # Run any missing collections
    if ($null -eq $LicenseData)   { $LicenseData   = Get-LicenseInventory }
    if ($null -eq $UnlicensedData){ $UnlicensedData = Get-UnlicensedUsers  }
    if ($null -eq $InactiveData)  { $InactiveData   = Get-InactiveUsers    }
    if ($null -eq $NoMfaData)     { $NoMfaData      = Get-MfaStatus        }

    Write-Step "Building HTML report..."

    $html      = Build-HtmlReport -LicenseData $LicenseData -UnlicensedData $UnlicensedData -InactiveData $InactiveData -NoMfaData $NoMfaData
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "VAULT_${timestamp}.html"

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Write-Ok "Report saved: $outPath"
        Write-Step "Opening in default browser..."
        Start-Process $outPath
    } catch {
        Write-Fail "Failed to save report: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-VaultBanner

    $connStatus = if ($script:Connected) {
        "  Connected as : $($script:ConnectedAs)"
    } else {
        "  Not Connected — select option 1 to authenticate"
    }
    $connColor = if ($script:Connected) { $C.Success } else { $C.Warning }

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  License inventory — SKUs, assigned vs. total" -ForegroundColor $C.Info
    Write-Host "  3  Unlicensed users — enabled accounts with no license" -ForegroundColor $C.Info
    Write-Host "  4  Inactive users — no sign-in in 90+ days" -ForegroundColor $C.Info
    Write-Host "  5  MFA status — users with no MFA method registered" -ForegroundColor $C.Warning
    Write-Host "  6  Shared mailbox audit guidance (Exchange Online)" -ForegroundColor $C.Info
    Write-Host "  7  Export full HTML report (all findings)" -ForegroundColor $C.Success
    Write-Host "  Q  Quit" -ForegroundColor $C.Info
    Write-Host ""
}

function Assert-Connected {
    if (-not $script:Connected) {
        Write-Host ""
        Write-Warn "Not connected to Microsoft 365. Please select option 1 first."
        Write-Host ""
        Read-Host "  Press Enter to return to menu"
        return $false
    }
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

Show-VaultBanner
Install-GraphModule

# Check if already connected (e.g. token still cached from prior session)
if (Test-GraphConnection) {
    Write-Ok "Already connected as $($script:ConnectedAs)"
    Write-Host ""
}

# Unattended mode: connect then run full report and exit
if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    Write-Step "Attempting to connect using cached credentials..."

    if (-not $script:Connected) {
        Invoke-Connect
    }

    if (-not $script:Connected) {
        Write-Fail "Could not establish a Graph connection in unattended mode. Exiting."
        exit 1
    }

    $licData       = Get-LicenseInventory
    $unlicData     = Get-UnlicensedUsers
    $inactiveData  = Get-InactiveUsers
    $noMfaData     = Get-MfaStatus

    Export-HtmlReport `
        -LicenseData   $licData `
        -UnlicensedData $unlicData `
        -InactiveData  $inactiveData `
        -NoMfaData     $noMfaData

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE MENU LOOP
# ─────────────────────────────────────────────────────────────────────────────

# Cached audit data — reset per session so menu re-runs pull fresh data
$cachedLicense    = $null
$cachedUnlicensed = $null
$cachedInactive   = $null
$cachedNoMfa      = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {

        '1' {
            Invoke-Connect
            # Reset cache on reconnect
            $cachedLicense    = $null
            $cachedUnlicensed = $null
            $cachedInactive   = $null
            $cachedNoMfa      = $null
            Read-Host "  Press Enter to return to menu"
        }

        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedLicense = Get-LicenseInventory
            Read-Host "  Press Enter to return to menu"
        }

        '3' {
            if (-not (Assert-Connected)) { break }
            $cachedUnlicensed = Get-UnlicensedUsers
            Read-Host "  Press Enter to return to menu"
        }

        '4' {
            if (-not (Assert-Connected)) { break }
            $cachedInactive = Get-InactiveUsers
            Read-Host "  Press Enter to return to menu"
        }

        '5' {
            if (-not (Assert-Connected)) { break }
            $cachedNoMfa = Get-MfaStatus
            Read-Host "  Press Enter to return to menu"
        }

        '6' {
            Show-SharedMailboxGuidance
            Read-Host "  Press Enter to return to menu"
        }

        '7' {
            if (-not (Assert-Connected)) { break }
            Export-HtmlReport `
                -LicenseData    $cachedLicense `
                -UnlicensedData $cachedUnlicensed `
                -InactiveData   $cachedInactive `
                -NoMfaData      $cachedNoMfa
            Read-Host "  Press Enter to return to menu"
        }

        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            Write-Host "  Goodbye." -ForegroundColor $C.Header
            Write-Host ""
        }

        default {
            Write-Host ""
            Write-Warn "Invalid option. Please choose 1–7 or Q."
            Write-Host ""
            Start-Sleep -Milliseconds 800
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
