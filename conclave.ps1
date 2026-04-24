<#
.SYNOPSIS
    C.O.N.C.L.A.V.E. — Consolidates Organisational Networks, Chats, Licenses, Access, Visibility & Entitlements
    Microsoft Teams Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and audits the M365 Teams estate:
    team inventory with visibility and size, orphan teams (no owner or
    all owners disabled), public teams (join without approval), guest
    membership, large teams, and stale teams with no recent activity.
    Produces a dark-themed HTML report that complements R.E.L.I.Q.U.A.R.Y.
    (licensing), G.O.L.E.M. (devices), and W.R.A.I.T.H. (identity) by
    covering the collaboration layer.

.USAGE
    PS C:\> .\conclave.ps1                    # Interactive menu
    PS C:\> .\conclave.ps1 -Unattended        # Silent: auto-connect + export HTML

.NOTES
    Version : 3.0

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
    'Team.ReadBasic.All',
    'TeamMember.Read.All',
    'Group.Read.All',
    'User.Read.All',
    'Directory.Read.All'
)

function Show-ConclaveBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  C.O.N.C.L.A.V.E. — Consolidates Organisational Networks, Chats, Licenses, Access, Visibility & Entitlements" -ForegroundColor Cyan
    Write-Host "  Microsoft Teams Audit Tool  v3.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Connect, audits, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH MODULE INSTALL + CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $required = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Teams',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    $needsInstall = $false
    foreach ($m in $required) {
        if (Get-Module -ListAvailable -Name $m) { Write-Ok "$m — installed" }
        else { Write-Warn "$m — NOT found"; $needsInstall = $true }
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
                Write-TKError -ScriptName 'conclave' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
                exit 1
            }
        } else {
            $ans = Read-Host "  Install Microsoft.Graph for current user? [Y/N]"
            if ($ans -match '^[Yy]') {
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Ok "Microsoft.Graph installed."
                } catch { Write-Fail "Install failed: $_"; exit 1 }
            } else { Write-Fail "Microsoft.Graph required. Exiting."; exit 1 }
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
                # Organization lookup is best-effort metadata for the report header.
            }
            return $true
        }
    } catch {
        # No cached context — caller will invoke interactive auth.
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
        Write-TKError -ScriptName 'conclave' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — TEAMS INVENTORY (enriched)
# ─────────────────────────────────────────────────────────────────────────────

# Enumerate teams via the M365 groups that have resourceProvisioningOptions=Team,
# then for each team pull owner/member counts and the archive/visibility state.
# This is chattier than Get-MgTeam but gives us the owner list we need for the
# orphan detector in section 2.
function Get-TeamsInventory {
    Write-Section "TEAMS INVENTORY"
    Write-Step "Enumerating teams (this can take a while on large tenants)..."

    try {
        $groups = Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
            -Property "id,displayName,description,visibility,mailNickname,createdDateTime,renewedDateTime,assignedLabels" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to list team-backed groups: $_"
        Write-TKError -ScriptName 'conclave' -Message "Get-MgGroup (team filter) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    if (-not $groups -or $groups.Count -eq 0) {
        Write-Warn "No teams found in this tenant."
        return @()
    }

    Write-Ok "$($groups.Count) team(s) found. Enriching owner/member counts..."

    $total   = $groups.Count
    $idx     = 0
    $records = foreach ($g in $groups) {
        $idx++
        if ($idx -eq 1 -or ($idx % 25) -eq 0 -or $idx -eq $total) {
            $pct = [math]::Round(($idx / $total) * 100, 0)
            Write-Host ("`r  [*] Enriching team {0}/{1} ({2}%)..." -f $idx, $total, $pct) -ForegroundColor $C.Progress -NoNewline
        }

        $ownerList  = @()
        $memberList = @()
        try { $ownerList  = @(Get-MgGroupOwner  -GroupId $g.Id -All -ErrorAction SilentlyContinue) } catch {
            # Owner read failures on a single team don't justify aborting the audit.
        }
        try { $memberList = @(Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue) } catch {
            # Same for member read failures.
        }

        $guestCount = 0
        foreach ($m in $memberList) {
            $type = $m.AdditionalProperties['@odata.type']
            $userType = $m.AdditionalProperties['userType']
            if ($type -match 'user' -and $userType -eq 'Guest') { $guestCount++ }
        }

        $enabledOwners = 0
        foreach ($o in $ownerList) {
            $enabled = $o.AdditionalProperties['accountEnabled']
            if ($enabled -ne $false) { $enabledOwners++ }
        }

        [PSCustomObject]@{
            Id             = $g.Id
            DisplayName    = $g.DisplayName
            Description    = $g.Description
            Visibility     = $g.Visibility
            Created        = if ($g.CreatedDateTime) { ([datetime]$g.CreatedDateTime).ToString('yyyy-MM-dd') } else { '' }
            LastRenewed    = if ($g.RenewedDateTime) { ([datetime]$g.RenewedDateTime).ToString('yyyy-MM-dd') } else { '' }
            OwnerCount     = $ownerList.Count
            EnabledOwners  = $enabledOwners
            MemberCount    = $memberList.Count
            GuestCount     = $guestCount
            Labels         = if ($g.AssignedLabels) { ($g.AssignedLabels | ForEach-Object { $_.DisplayName }) -join '; ' } else { '' }
        }
    }

    Write-Host ""  # close the progress line
    Write-Ok "Collected $($records.Count) team record(s)."
    Write-Host ""
    return @($records)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTIONS 2-5 — DERIVED FINDINGS FROM THE INVENTORY
# ─────────────────────────────────────────────────────────────────────────────

# Thresholds editable here; section 2+ functions are thin filters over the
# inventory so the audit stays deterministic and cheap.
$script:LargeTeamThreshold = 250
$script:StaleRenewalDays   = 365

function Get-OrphanTeams {
    param([array]$Teams)
    # Orphan = zero owners at all, OR every owner is disabled.
    @($Teams | Where-Object {
        $_.OwnerCount -eq 0 -or ($_.OwnerCount -gt 0 -and $_.EnabledOwners -eq 0)
    })
}

function Get-PublicTeams {
    param([array]$Teams)
    @($Teams | Where-Object { $_.Visibility -eq 'Public' })
}

function Get-GuestHeavyTeams {
    param([array]$Teams)
    # Any team with at least one guest member; the flagging is done in the UI.
    @($Teams | Where-Object { $_.GuestCount -gt 0 } | Sort-Object GuestCount -Descending)
}

function Get-LargeTeams {
    param([array]$Teams)
    @($Teams | Where-Object { $_.MemberCount -ge $script:LargeTeamThreshold } | Sort-Object MemberCount -Descending)
}

function Get-StaleTeams {
    param([array]$Teams)
    # M365 groups can expire if a group-expiration policy is set; RenewedDateTime
    # is the last time the owner confirmed the group is still in use. A long gap
    # suggests the team has lapsed out of active management.
    $threshold = (Get-Date).AddDays(-$script:StaleRenewalDays)
    @($Teams | Where-Object {
        $_.LastRenewed -and ([datetime]$_.LastRenewed) -lt $threshold
    } | Sort-Object LastRenewed)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Teams, [array]$Orphans, [array]$Public, [array]$Guests, [array]$Large, [array]$Stale)

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $tkCfg         = Get-TKConfig
    $orgPrefix     = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _inv-row {
        param($t)
        $vizBadge = switch ($t.Visibility) {
            'Public'  { "<span class='tk-badge-warn'>Public</span>" }
            'Private' { "<span class='tk-badge-ok'>Private</span>" }
            default   { "<span class='tk-badge-info'>$(EscHtml $t.Visibility)</span>" }
        }
        $ownerBadge = if ($t.OwnerCount -eq 0) { "<span class='tk-badge-err'>$($t.OwnerCount)</span>" }
                      elseif ($t.EnabledOwners -eq 0) { "<span class='tk-badge-err'>$($t.OwnerCount) (all disabled)</span>" }
                      else { "<span class='tk-badge-ok'>$($t.OwnerCount) ($($t.EnabledOwners) enabled)</span>" }
        $guestBadge = if ($t.GuestCount -eq 0) { "<span class='tk-badge-ok'>0</span>" } else { "<span class='tk-badge-warn'>$($t.GuestCount)</span>" }
        $memBadge   = if ($t.MemberCount -ge $script:LargeTeamThreshold) { "<span class='tk-badge-warn'>$($t.MemberCount)</span>" } else { "<span class='tk-badge-info'>$($t.MemberCount)</span>" }
        return "<tr><td>$(EscHtml $t.DisplayName)</td><td>$vizBadge</td><td>$ownerBadge</td><td>$memBadge</td><td>$guestBadge</td><td>$(EscHtml $t.Created)</td><td>$(EscHtml $t.LastRenewed)</td><td>$(EscHtml $t.Labels)</td></tr>"
    }

    $invRows = [System.Text.StringBuilder]::new()
    if ($Teams.Count -eq 0) {
        [void]$invRows.Append("<tr><td colspan='8' class='tk-badge-info' style='text-align:center;'>No teams found.</td></tr>")
    } else {
        foreach ($t in ($Teams | Sort-Object DisplayName)) { [void]$invRows.Append((_inv-row $t) + "`n") }
    }

    function _category-rows {
        param([array]$Rows, [int]$Cols, [string]$EmptyMessage)
        $sb = [System.Text.StringBuilder]::new()
        if ($Rows.Count -eq 0) {
            [void]$sb.Append("<tr><td colspan='$Cols' class='tk-badge-ok' style='text-align:center;'>$EmptyMessage</td></tr>")
        } else {
            foreach ($t in $Rows) { [void]$sb.Append((_inv-row $t) + "`n") }
        }
        return $sb.ToString()
    }

    $orphanRows = _category-rows -Rows $Orphans -Cols 8 -EmptyMessage 'No orphan teams.'
    $publicRows = _category-rows -Rows $Public  -Cols 8 -EmptyMessage 'No public teams.'
    $guestRows  = _category-rows -Rows $Guests  -Cols 8 -EmptyMessage 'No teams contain guest members.'
    $largeRows  = _category-rows -Rows $Large   -Cols 8 -EmptyMessage "No teams exceed $script:LargeTeamThreshold members."
    $staleRows  = _category-rows -Rows $Stale   -Cols 8 -EmptyMessage "No teams have gone unrenewed for $script:StaleRenewalDays+ days."

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.O.N.C.L.A.V.E. Teams Audit Report' `
        -ScriptName 'C.O.N.C.L.A.V.E.' `
        -Subtitle   "${orgPrefix}Microsoft Teams Estate Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = EscHtml $script:ConnectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @('Teams Inventory', 'Orphan Teams', 'Public Teams', 'Guest Members', 'Large Teams', 'Stale Teams')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'C.O.N.C.L.A.V.E. v3.0'

    $orphClass  = if ($Orphans.Count -gt 0) { 'err' } else { 'ok' }
    $pubClass   = if ($Public.Count  -gt 0) { 'warn' } else { 'ok' }
    $guestClass = if ($Guests.Count  -gt 0) { 'warn' } else { 'ok' }
    $staleClass = if ($Stale.Count   -gt 0) { 'warn' } else { 'ok' }

    $invHeader = "<tr><th>Display Name</th><th>Visibility</th><th>Owners</th><th>Members</th><th>Guests</th><th>Created</th><th>Renewed</th><th>Sensitivity Labels</th></tr>"

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Teams.Count)</div><div class="tk-summary-lbl">Teams Total</div></div>
    <div class="tk-summary-card $orphClass"><div class="tk-summary-num">$($Orphans.Count)</div><div class="tk-summary-lbl">Orphan Teams</div></div>
    <div class="tk-summary-card $pubClass"><div class="tk-summary-num">$($Public.Count)</div><div class="tk-summary-lbl">Public Teams</div></div>
    <div class="tk-summary-card $guestClass"><div class="tk-summary-num">$($Guests.Count)</div><div class="tk-summary-lbl">Teams With Guests</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Large.Count)</div><div class="tk-summary-lbl">Large Teams (&ge; $script:LargeTeamThreshold)</div></div>
    <div class="tk-summary-card $staleClass"><div class="tk-summary-num">$($Stale.Count)</div><div class="tk-summary-lbl">Stale (&gt; $script:StaleRenewalDays d unrenewed)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Teams Inventory</span><span class="tk-section-num">$($Teams.Count) team(s)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$($invRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Orphan Teams -- No Owners (or All Owners Disabled)</span><span class="tk-section-num">$($Orphans.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$orphanRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Public Teams -- Anyone in Tenant Can Join</span><span class="tk-section-num">$($Public.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$publicRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Teams With Guest Members</span><span class="tk-section-num">$($Guests.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$guestRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Large Teams -- &ge; $script:LargeTeamThreshold Members</span><span class="tk-section-num">$($Large.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$largeRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Stale Teams -- Unrenewed &gt; $script:StaleRenewalDays Days</span><span class="tk-section-num">$($Stale.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$staleRows</tbody></table></div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Teams, [array]$Orphans, [array]$Public, [array]$Guests, [array]$Large, [array]$Stale)

    Write-Section "EXPORTING HTML REPORT"
    $html      = Build-HtmlReport -Teams $Teams -Orphans $Orphans -Public $Public -Guests $Guests -Large $Large -Stale $Stale
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "CONCLAVE_${timestamp}.html"

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Write-Ok "Report saved: $outPath"
        if (-not $Unattended) {
            Write-Step "Opening in default browser..."
            Start-Process $outPath
        }
    } catch {
        Write-Fail "Could not save report: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-ConclaveBanner
    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  Teams inventory (refresh underlying data)" -ForegroundColor $C.Info
    Write-Host "  3  Export full HTML report (orphans, public, guests, large, stale)" -ForegroundColor $C.Success
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

Show-ConclaveBanner
Install-GraphModule

if (Test-GraphConnection) {
    Write-Ok "Already connected as $($script:ConnectedAs)"
    Write-Host ""
}

if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    if (-not $script:Connected) { Invoke-Connect }
    if (-not $script:Connected) { Write-Fail "Could not connect to Graph. Exiting."; exit 1 }

    $teams   = Get-TeamsInventory
    $orphans = Get-OrphanTeams    -Teams $teams
    $public  = Get-PublicTeams    -Teams $teams
    $guests  = Get-GuestHeavyTeams -Teams $teams
    $large   = Get-LargeTeams     -Teams $teams
    $stale   = Get-StaleTeams     -Teams $teams

    Export-HtmlReport -Teams $teams -Orphans $orphans -Public $public -Guests $guests -Large $large -Stale $stale

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

$cachedTeams = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {
        '1' {
            Invoke-Connect
            $cachedTeams = $null
            Read-Host "  Press Enter to return to menu"
        }
        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedTeams = Get-TeamsInventory
            Read-Host "  Press Enter to return to menu"
        }
        '3' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedTeams) { $cachedTeams = Get-TeamsInventory }
            $orphans = Get-OrphanTeams     -Teams $cachedTeams
            $public  = Get-PublicTeams     -Teams $cachedTeams
            $guests  = Get-GuestHeavyTeams -Teams $cachedTeams
            $large   = Get-LargeTeams      -Teams $cachedTeams
            $stale   = Get-StaleTeams      -Teams $cachedTeams
            Export-HtmlReport -Teams $cachedTeams -Orphans $orphans -Public $public -Guests $guests -Large $large -Stale $stale
            Read-Host "  Press Enter to return to menu"
        }
        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {
                # Best-effort cleanup.
            }
            Write-Host "  Goodbye." -ForegroundColor $C.Header
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Warn "Invalid option. Please choose 1, 2, 3, or Q."
            Write-Host ""
            Start-Sleep -Milliseconds 800
        }
    }
} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
