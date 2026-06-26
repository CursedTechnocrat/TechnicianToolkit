<#
.SYNOPSIS
    W.R.A.I.T.H. — Watches Registrations, Access, Identities, Tokens & Hygiene
    Entra ID Identity Hygiene Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and audits the Entra ID (Azure AD) identity
    posture: guest users with stale sign-ins, privileged role holders and
    their MFA state, members with password-never-expires set, stale
    privileged users, and disabled-but-licensed accounts that still consume
    paid SKUs. Produces a dark-themed HTML report that complements
    R.E.L.I.Q.U.A.R.Y. (licensing / mailboxes) and G.O.L.E.M. (device
    compliance) — together the three tools give a full tenant posture
    report.

.USAGE
    PS C:\> .\wraith.ps1                    # Interactive menu
    PS C:\> .\wraith.ps1 -Unattended        # Silent mode — auto-connect and export HTML

.NOTES
    Version : 3.6

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

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
}

$script:Connected    = $false
$script:ConnectedAs  = ''
$script:TenantDomain = ''

$GraphScopes = @(
    'User.Read.All',
    'Directory.Read.All',
    'AuditLog.Read.All',
    'RoleManagement.Read.Directory'
)

function Show-WraithBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  W.R.A.I.T.H. — Watches Registrations, Access, Identities, Tokens & Hygiene" -ForegroundColor Cyan
    Write-Host "  Entra ID Identity Hygiene Audit Tool  v3.6" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Helpers, audit functions, and entry point appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH MODULE INSTALL + CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $required = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Identity.Governance'
    )

    $needsInstall = $false
    foreach ($m in $required) {
        if (Get-Module -ListAvailable -Name $m) {
            Write-Ok "$m — installed"
        } else {
            Write-Warn "$m — NOT found"
            $needsInstall = $true
        }
    }

    if ($needsInstall) {
        Write-Host ""
        if ($Unattended) {
            Write-Step "Installing Microsoft.Graph (unattended)..."
            try {
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Ok "Microsoft.Graph installed."
            } catch {
                Write-Fail "Install failed: $_"
                Write-TKError -ScriptName 'wraith' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
                exit 1
            }
        } else {
            $ans = Read-Host "  Install Microsoft.Graph for current user? [Y/N]"
            if ($ans -match '^[Yy]') {
                Write-Step "Installing Microsoft.Graph..."
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Ok "Microsoft.Graph installed."
                } catch {
                    Write-Fail "Install failed: $_"
                    exit 1
                }
            } else {
                Write-Fail "Microsoft.Graph is required. Exiting."
                exit 1
            }
        }
    } else {
        Write-Ok "All required Graph sub-modules are available."
    }

    foreach ($m in $required) {
        try { Import-Module $m -ErrorAction Stop } catch { Write-Warn "Could not import ${m}: $_" }
    }
}

function Test-GraphConnection {
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account) {
            $script:Connected   = $true
            $script:ConnectedAs = $ctx.Account
            try {
                $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($org -and $org.VerifiedDomains) {
                    $primary = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
                    if ($primary) { $script:TenantDomain = $primary.Name }
                }
            } catch {
                # Best-effort metadata.
            }
            return $true
        }
    } catch {
        # No cached context.
    }
    $script:Connected   = $false
    $script:ConnectedAs = ''
    return $false
}

function Invoke-Connect {
    Write-Section "CONNECT TO MICROSOFT 365"
    Write-Step "Requesting interactive sign-in..."
    Write-Info "Scopes: $($GraphScopes -join ', ')"
    Write-Host ""

    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop

        if (Test-GraphConnection) {
            Write-Ok "Connected as: $($script:ConnectedAs)"
            if ($script:TenantDomain) { Write-Ok "Tenant: $($script:TenantDomain)" }
        } else {
            Write-Fail "Authentication returned no context — may have been cancelled."
        }
    } catch {
        Write-Fail "Authentication failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'wraith' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — GUEST USERS
# ─────────────────────────────────────────────────────────────────────────────

# Every guest with their last sign-in and creation date. Stale guests
# (no sign-in in 90 days) are the single easiest external-access cleanup win.
function Get-GuestAudit {
    Write-Section "GUEST USERS"
    Write-Step "Retrieving guest accounts and sign-in activity..."

    $threshold = (Get-Date).AddDays(-90)

    try {
        $guests = Get-MgUser -All `
            -Filter "userType eq 'Guest'" `
            -Property "displayName,userPrincipalName,mail,createdDateTime,accountEnabled,signInActivity,externalUserState" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve guests: $_"
        Write-TKError -ScriptName 'wraith' -Message "Get-MgUser (guest filter) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $now = Get-Date
    $rows = foreach ($g in $guests) {
        $lastSignIn   = $null
        $daysInactive = $null
        if ($g.SignInActivity -and $g.SignInActivity.LastSignInDateTime) {
            $lastSignIn   = [datetime]$g.SignInActivity.LastSignInDateTime
            $daysInactive = [math]::Round(($now - $lastSignIn).TotalDays, 0)
        }

        $isStale = (-not $lastSignIn) -or ($lastSignIn -lt $threshold)

        [PSCustomObject]@{
            DisplayName   = $g.DisplayName
            UPN           = $g.UserPrincipalName
            Email         = $g.Mail
            Created       = if ($g.CreatedDateTime) { ([datetime]$g.CreatedDateTime).ToString('yyyy-MM-dd') } else { '—' }
            LastSignIn    = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
            DaysInactive  = $daysInactive
            IsStale       = $isStale
            AccountState  = if ($g.AccountEnabled) { 'Enabled' } else { 'Disabled' }
            InviteState   = if ($g.ExternalUserState) { $g.ExternalUserState } else { '—' }
        }
    }

    $sorted = @($rows | Sort-Object @{ Expression = { if ($null -eq $_.DaysInactive) { 99999 } else { [int]$_.DaysInactive } }; Descending = $true })

    Write-Ok "$($sorted.Count) guest account(s) found."
    $staleCount = @($sorted | Where-Object { $_.IsStale }).Count
    if ($staleCount -gt 0) { Write-Warn "$staleCount stale guest(s) — no sign-in in 90+ days." }
    Write-Host ""
    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — PRIVILEGED ROLE HOLDERS
# ─────────────────────────────────────────────────────────────────────────────

# Every member of every active directory role. Emits one row per user-per-role
# so users with multiple roles appear multiple times (intentional — the audit
# question is "who has this role", not "who are all the admins").
function Get-PrivilegedRoleAudit {
    Write-Section "PRIVILEGED ROLE HOLDERS"
    Write-Step "Retrieving active directory role assignments..."

    try {
        $roles = Get-MgDirectoryRole -All -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve directory roles: $_"
        Write-TKError -ScriptName 'wraith' -Message "Get-MgDirectoryRole failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($role in $roles) {
        try {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
        } catch {
            continue
        }

        foreach ($m in $members) {
            # Role members can be users, groups, or service principals — pull the
            # fields present on user objects and mark non-user members distinctly.
            $props = $m.AdditionalProperties
            $upn   = $props['userPrincipalName']
            $disp  = $props['displayName']
            $typeRaw = $props['@odata.type']
            $type  = if ($typeRaw) { ($typeRaw -replace '#microsoft.graph.', '') } else { 'unknown' }

            $rows.Add([PSCustomObject]@{
                RoleName    = $role.DisplayName
                MemberType  = $type
                DisplayName = $disp
                UPN         = $upn
                MemberId    = $m.Id
            })
        }
    }

    $sorted = @($rows | Sort-Object RoleName, DisplayName)
    $unique = @($sorted | Group-Object UPN | Where-Object { $_.Name }).Count
    Write-Ok "$($sorted.Count) role assignment(s) covering $unique unique principal(s)."

    # Surface the high-tier roles in the console summary.
    $highTier = @('Global Administrator','Privileged Role Administrator','User Administrator','Exchange Administrator','SharePoint Administrator','Security Administrator','Conditional Access Administrator')
    foreach ($role in ($sorted | Where-Object { $highTier -contains $_.RoleName } | Group-Object RoleName)) {
        Write-Warn ("  {0,-35} {1} member(s)" -f $role.Name, $role.Count)
    }
    Write-Host ""

    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — PASSWORD NEVER EXPIRES
# ─────────────────────────────────────────────────────────────────────────────

# Cloud-only accounts with PasswordPolicies=DisablePasswordExpiration are the
# classic tenant hygiene miss — shared mailboxes, service accounts, or
# holdovers from a failed sync that quietly disabled expiry.
function Get-PasswordNeverExpiresAudit {
    Write-Section "PASSWORD NEVER EXPIRES"
    Write-Step "Retrieving members with password-never-expires set..."

    try {
        $users = Get-MgUser -All `
            -Property "displayName,userPrincipalName,passwordPolicies,accountEnabled,userType,onPremisesSyncEnabled,assignedLicenses" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        Write-TKError -ScriptName 'wraith' -Message "Get-MgUser (all) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $rows = foreach ($u in $users) {
        if ($u.UserType -eq 'Guest') { continue }
        $policies = if ($u.PasswordPolicies) { $u.PasswordPolicies } else { '' }
        if ($policies -notmatch 'DisablePasswordExpiration') { continue }

        [PSCustomObject]@{
            DisplayName    = $u.DisplayName
            UPN            = $u.UserPrincipalName
            AccountEnabled = $u.AccountEnabled
            SyncedFromAD   = [bool]$u.OnPremisesSyncEnabled
            Licensed       = ($u.AssignedLicenses.Count -gt 0)
            PasswordPolicy = $policies
        }
    }

    $rows = @($rows | Sort-Object DisplayName)
    if ($rows.Count -eq 0) {
        Write-Ok "No members have password-never-expires set."
    } else {
        Write-Warn "$($rows.Count) member(s) with password-never-expires."
    }
    Write-Host ""
    return $rows
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — STALE PRIVILEGED USERS
# ─────────────────────────────────────────────────────────────────────────────

# Admins who haven't signed in in 60+ days. Same risk-class as "orphan"
# service accounts — their credentials still work and still have elevated
# access, but no one is watching the audit log.
function Get-StaleAdminAudit {
    param([array]$RoleRows)

    Write-Section "STALE PRIVILEGED USERS"
    Write-Step "Checking sign-in activity for privileged users (no sign-in in 60+ days)..."

    $threshold = (Get-Date).AddDays(-60)
    $now       = Get-Date

    # Deduplicate by UPN — a user in multiple roles is still one stale admin.
    $uniqueUpns = @($RoleRows | Where-Object { $_.MemberType -eq 'user' -and $_.UPN } | Select-Object -ExpandProperty UPN -Unique)

    if ($uniqueUpns.Count -eq 0) {
        Write-Warn "No user-type role members to evaluate."
        Write-Host ""
        return @()
    }

    $stale = [System.Collections.Generic.List[object]]::new()
    foreach ($upn in $uniqueUpns) {
        try {
            $u = Get-MgUser -UserId $upn -Property "displayName,userPrincipalName,accountEnabled,signInActivity" -ErrorAction Stop
        } catch {
            continue
        }

        $lastSignIn   = $null
        $daysInactive = $null
        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn   = [datetime]$u.SignInActivity.LastSignInDateTime
            $daysInactive = [math]::Round(($now - $lastSignIn).TotalDays, 0)
        }

        $isStale = (-not $lastSignIn) -or ($lastSignIn -lt $threshold)
        if (-not $isStale) { continue }

        $roles = @($RoleRows | Where-Object { $_.UPN -eq $upn } | Select-Object -ExpandProperty RoleName -Unique) -join ', '

        $stale.Add([PSCustomObject]@{
            DisplayName    = $u.DisplayName
            UPN            = $u.UserPrincipalName
            AccountEnabled = $u.AccountEnabled
            LastSignIn     = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
            DaysInactive   = $daysInactive
            Roles          = $roles
        })
    }

    $sorted = @($stale | Sort-Object @{ Expression = { if ($null -eq $_.DaysInactive) { 99999 } else { [int]$_.DaysInactive } }; Descending = $true })
    if ($sorted.Count -eq 0) {
        Write-Ok "All privileged users have signed in within the last 60 days."
    } else {
        Write-Fail "$($sorted.Count) privileged user(s) inactive 60+ days."
    }
    Write-Host ""
    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — DISABLED BUT LICENSED
# ─────────────────────────────────────────────────────────────────────────────

# Cost-leak audit: disabled accounts still consuming paid SKUs. Common after
# off-boarding where the license was never reclaimed.
function Get-DisabledLicensedAudit {
    Write-Section "DISABLED BUT LICENSED"
    Write-Step "Finding disabled accounts that still have license assignments..."

    try {
        $users = Get-MgUser -All `
            -Filter "accountEnabled eq false" `
            -Property "displayName,userPrincipalName,accountEnabled,assignedLicenses,userType" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        Write-TKError -ScriptName 'wraith' -Message "Get-MgUser (disabled filter) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $rows = foreach ($u in $users) {
        if ($u.UserType -eq 'Guest') { continue }
        if ($u.AssignedLicenses.Count -eq 0) { continue }
        [PSCustomObject]@{
            DisplayName   = $u.DisplayName
            UPN           = $u.UserPrincipalName
            LicenseCount  = $u.AssignedLicenses.Count
        }
    }

    $sorted = @($rows | Sort-Object DisplayName)
    if ($sorted.Count -eq 0) {
        Write-Ok "No disabled accounts are consuming licenses."
    } else {
        Write-Fail "$($sorted.Count) disabled account(s) still consuming licenses."
    }
    Write-Host ""
    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Guests, [array]$Roles, [array]$PwdNever, [array]$StaleAdmins, [array]$DisabledLic)

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $connectedAs   = EscHtml $script:ConnectedAs

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    # Guest rows
    $guestRows = [System.Text.StringBuilder]::new()
    if ($Guests.Count -eq 0) {
        [void]$guestRows.Append("<tr><td colspan='6' class='tk-badge-ok' style='text-align:center;'>No guest accounts.</td></tr>")
    } else {
        foreach ($g in $Guests) {
            $staleClass = if ($g.IsStale) { 'tk-badge-warn' } else { 'tk-badge-ok' }
            $staleLabel = if ($g.LastSignIn -eq 'Never') { 'Never' } else { $g.LastSignIn }
            [void]$guestRows.Append("<tr><td>$(EscHtml $g.DisplayName)</td><td>$(EscHtml $g.UPN)</td><td>$(EscHtml $g.Created)</td><td><span class='$staleClass'>$(EscHtml $staleLabel)</span></td><td>$(EscHtml $g.AccountState)</td><td>$(EscHtml $g.InviteState)</td></tr>`n")
        }
    }

    # Role rows
    $roleRows = [System.Text.StringBuilder]::new()
    if ($Roles.Count -eq 0) {
        [void]$roleRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No directory role assignments.</td></tr>")
    } else {
        foreach ($r in $Roles) {
            $typeBadge = switch ($r.MemberType) {
                'user'             { "<span class='tk-badge-ok'>User</span>" }
                'group'            { "<span class='tk-badge-warn'>Group</span>" }
                'servicePrincipal' { "<span class='tk-badge-info'>Service Principal</span>" }
                default            { "<span class='tk-badge-info'>$(EscHtml $r.MemberType)</span>" }
            }
            [void]$roleRows.Append("<tr><td>$(EscHtml $r.RoleName)</td><td>$typeBadge</td><td>$(EscHtml $r.DisplayName)</td><td>$(EscHtml $r.UPN)</td></tr>`n")
        }
    }

    # Password-never-expires rows
    $pwdRows = [System.Text.StringBuilder]::new()
    if ($PwdNever.Count -eq 0) {
        [void]$pwdRows.Append("<tr><td colspan='5' class='tk-badge-ok' style='text-align:center;'>No members have password-never-expires set.</td></tr>")
    } else {
        foreach ($p in $PwdNever) {
            $enBadge   = if ($p.AccountEnabled) { "<span class='tk-badge-ok'>Enabled</span>" } else { "<span class='tk-badge-info'>Disabled</span>" }
            $syncBadge = if ($p.SyncedFromAD)   { "<span class='tk-badge-info'>AD sync</span>" } else { "<span class='tk-badge-warn'>Cloud-only</span>" }
            $licBadge  = if ($p.Licensed)       { "<span class='tk-badge-ok'>Licensed</span>" } else { "<span class='tk-badge-info'>Unlicensed</span>" }
            [void]$pwdRows.Append("<tr><td>$(EscHtml $p.DisplayName)</td><td>$(EscHtml $p.UPN)</td><td>$enBadge</td><td>$syncBadge</td><td>$licBadge</td></tr>`n")
        }
    }

    # Stale admin rows
    $staleRows = [System.Text.StringBuilder]::new()
    if ($StaleAdmins.Count -eq 0) {
        [void]$staleRows.Append("<tr><td colspan='5' class='tk-badge-ok' style='text-align:center;'>All privileged users have recent sign-in activity.</td></tr>")
    } else {
        foreach ($s in $StaleAdmins) {
            $lastClass = if ($s.LastSignIn -eq 'Never') { 'tk-badge-err' } else { 'tk-badge-warn' }
            [void]$staleRows.Append("<tr><td>$(EscHtml $s.DisplayName)</td><td>$(EscHtml $s.UPN)</td><td>$(EscHtml $s.Roles)</td><td><span class='$lastClass'>$(EscHtml $s.LastSignIn)</span></td><td>$($s.DaysInactive)</td></tr>`n")
        }
    }

    # Disabled-licensed rows
    $dlRows = [System.Text.StringBuilder]::new()
    if ($DisabledLic.Count -eq 0) {
        [void]$dlRows.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;'>No disabled accounts are consuming licenses.</td></tr>")
    } else {
        foreach ($d in $DisabledLic) {
            [void]$dlRows.Append("<tr><td>$(EscHtml $d.DisplayName)</td><td>$(EscHtml $d.UPN)</td><td><span class='tk-badge-err'>$($d.LicenseCount)</span></td></tr>`n")
        }
    }

    # Summary cards
    $staleGuestCount = @($Guests | Where-Object { $_.IsStale }).Count
    $gClass = if ($staleGuestCount -gt 0) { 'warn' } else { 'ok' }
    $pClass = if ($PwdNever.Count    -gt 0) { 'warn' } else { 'ok' }
    $sClass = if ($StaleAdmins.Count -gt 0) { 'err'  } else { 'ok' }
    $dClass = if ($DisabledLic.Count -gt 0) { 'err'  } else { 'ok' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'W.R.A.I.T.H. Identity Hygiene Report' `
        -ScriptName 'W.R.A.I.T.H.' `
        -Subtitle   "${orgPrefix}Entra ID Identity Hygiene Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = $connectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @('Guests', 'Privileged Roles', 'Password Never Expires', 'Stale Admins', 'Disabled but Licensed')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'W.R.A.I.T.H. v3.6'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Guests.Count)</div><div class="tk-summary-lbl">Guest Accounts</div></div>
    <div class="tk-summary-card $gClass"><div class="tk-summary-num">$staleGuestCount</div><div class="tk-summary-lbl">Stale Guests (90d+)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Roles.Count)</div><div class="tk-summary-lbl">Role Assignments</div></div>
    <div class="tk-summary-card $pClass"><div class="tk-summary-num">$($PwdNever.Count)</div><div class="tk-summary-lbl">Password Never Expires</div></div>
    <div class="tk-summary-card $sClass"><div class="tk-summary-num">$($StaleAdmins.Count)</div><div class="tk-summary-lbl">Stale Admins (60d+)</div></div>
    <div class="tk-summary-card $dClass"><div class="tk-summary-num">$($DisabledLic.Count)</div><div class="tk-summary-lbl">Disabled but Licensed</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Guest Accounts</span><span class="tk-section-num">$($Guests.Count) guest(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Created</th><th>Last Sign-In</th><th>Account</th><th>Invite State</th></tr></thead>
        <tbody>$($guestRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Privileged Role Holders</span><span class="tk-section-num">$($Roles.Count) assignment(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Role</th><th>Type</th><th>Display Name</th><th>UPN</th></tr></thead>
        <tbody>$($roleRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Password Never Expires</span><span class="tk-section-num">$($PwdNever.Count) user(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Account</th><th>Source</th><th>Licensed</th></tr></thead>
        <tbody>$($pwdRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Stale Privileged Users -- No Sign-In Within 60 Days</span><span class="tk-section-num">$($StaleAdmins.Count) user(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Roles</th><th>Last Sign-In</th><th>Days Inactive</th></tr></thead>
        <tbody>$($staleRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Disabled but Licensed</span><span class="tk-section-num">$($DisabledLic.Count) account(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Licenses Assigned</th></tr></thead>
        <tbody>$($dlRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Guests, [array]$Roles, [array]$PwdNever, [array]$StaleAdmins, [array]$DisabledLic)

    Write-Section "EXPORTING HTML REPORT"
    Write-Step "Building report..."

    $html      = Build-HtmlReport -Guests $Guests -Roles $Roles -PwdNever $PwdNever -StaleAdmins $StaleAdmins -DisabledLic $DisabledLic
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "WRAITH_${timestamp}.html"

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-WraithBanner

    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  Guest accounts — sign-in age, invite state" -ForegroundColor $C.Info
    Write-Host "  3  Privileged role holders — every directory role member" -ForegroundColor $C.Info
    Write-Host "  4  Password never expires — members with expiry disabled" -ForegroundColor $C.Warning
    Write-Host "  5  Stale privileged users — admins inactive 60+ days" -ForegroundColor $C.Warning
    Write-Host "  6  Disabled but licensed — accounts leaking license cost" -ForegroundColor $C.Warning
    Write-Host "  7  Export full HTML report (all findings)" -ForegroundColor $C.Success
    Write-Host "  Q  Quit" -ForegroundColor $C.Info
    Write-Host ""
}

function Assert-Connected {
    if (-not $script:Connected) {
        Write-Host ""
        Write-Warn "Not connected. Please select option 1 first."
        Write-Host ""
        Read-Host "  Press Enter to return to menu"
        return $false
    }
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

Show-WraithBanner
Install-GraphModule

if (Test-GraphConnection) {
    Write-Ok "Already connected as $($script:ConnectedAs)"
    Write-Host ""
}

if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    if (-not $script:Connected) { Invoke-Connect }
    if (-not $script:Connected) {
        Write-Fail "Could not connect to Graph. Exiting."
        exit 1
    }

    $guests      = Get-GuestAudit
    $roles       = Get-PrivilegedRoleAudit
    $pwdNever    = Get-PasswordNeverExpiresAudit
    $staleAdmins = Get-StaleAdminAudit -RoleRows $roles
    $disabledLic = Get-DisabledLicensedAudit

    Export-HtmlReport -Guests $guests -Roles $roles -PwdNever $pwdNever -StaleAdmins $staleAdmins -DisabledLic $disabledLic

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# Cached per-session so repeated menu choices don't re-query
$cachedGuests    = $null
$cachedRoles     = $null
$cachedPwdNever  = $null
$cachedStale     = $null
$cachedDisabled  = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {

        '1' {
            Invoke-Connect
            $cachedGuests = $null; $cachedRoles = $null; $cachedPwdNever = $null; $cachedStale = $null; $cachedDisabled = $null
            Read-Host "  Press Enter to return to menu"
        }

        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedGuests = Get-GuestAudit
            Read-Host "  Press Enter to return to menu"
        }

        '3' {
            if (-not (Assert-Connected)) { break }
            $cachedRoles = Get-PrivilegedRoleAudit
            Read-Host "  Press Enter to return to menu"
        }

        '4' {
            if (-not (Assert-Connected)) { break }
            $cachedPwdNever = Get-PasswordNeverExpiresAudit
            Read-Host "  Press Enter to return to menu"
        }

        '5' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedRoles) { $cachedRoles = Get-PrivilegedRoleAudit }
            $cachedStale = Get-StaleAdminAudit -RoleRows $cachedRoles
            Read-Host "  Press Enter to return to menu"
        }

        '6' {
            if (-not (Assert-Connected)) { break }
            $cachedDisabled = Get-DisabledLicensedAudit
            Read-Host "  Press Enter to return to menu"
        }

        '7' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedGuests)   { $cachedGuests   = Get-GuestAudit }
            if ($null -eq $cachedRoles)    { $cachedRoles    = Get-PrivilegedRoleAudit }
            if ($null -eq $cachedPwdNever) { $cachedPwdNever = Get-PasswordNeverExpiresAudit }
            if ($null -eq $cachedStale)    { $cachedStale    = Get-StaleAdminAudit -RoleRows $cachedRoles }
            if ($null -eq $cachedDisabled) { $cachedDisabled = Get-DisabledLicensedAudit }
            Export-HtmlReport -Guests $cachedGuests -Roles $cachedRoles -PwdNever $cachedPwdNever -StaleAdmins $cachedStale -DisabledLic $cachedDisabled
            Read-Host "  Press Enter to return to menu"
        }

        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {
                # Best-effort cleanup — cached context may already be gone.
            }
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
