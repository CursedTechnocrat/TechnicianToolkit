<#
.SYNOPSIS
    G.R.O.V.E. — Gathers, Reports On, & Verifies Estates (SharePoint)
    SharePoint Online Site Inventory Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and inventories the SharePoint Online
    estate: every site collection with storage used, file count, last
    activity date, external-sharing state, and owner metadata. Uses the
    Graph usage-report endpoint (getSharePointSiteUsageDetail) so the
    audit is a single bulk read rather than N per-site queries. Flags
    storage-heavy sites, externally shared sites, owner-less sites, and
    stale sites with no recent activity.

.USAGE
    PS C:\> .\grove.ps1                    # Interactive menu
    PS C:\> .\grove.ps1 -Unattended        # Silent: auto-connect + export HTML

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
    'Sites.Read.All',
    'Reports.Read.All',
    'Directory.Read.All'
)

function Show-GroveBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  G.R.O.V.E. — Gathers, Reports On, & Verifies Estates" -ForegroundColor Cyan
    Write-Host "  SharePoint Online Site Inventory Tool  v3.0" -ForegroundColor Cyan
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
        'Microsoft.Graph.Sites',
        'Microsoft.Graph.Reports',
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
                Write-TKError -ScriptName 'grove' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
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
                # Org lookup is best-effort metadata.
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
        Write-TKError -ScriptName 'grove' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SITE INVENTORY VIA USAGE REPORT
# ─────────────────────────────────────────────────────────────────────────────

# The getSharePointSiteUsageDetail report returns one CSV row per site with
# storage, file count, last activity date, owner, and external-share flag --
# everything we need for a fleet-level audit in a single bulk read. Beats the
# alternative of iterating Get-MgSite per URL, which rate-limits on large
# tenants.
function Get-SharePointInventory {
    Write-Section "SHAREPOINT SITE INVENTORY"
    Write-Step "Pulling SharePoint usage report (period=D30)..."

    $uri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')"
    $csv = $null
    try {
        # Invoke-MgGraphRequest returns the CSV as a string for report endpoints
        # when no output file is specified. Explicit -OutputType is required to
        # avoid the SDK auto-wrapping.
        $csv = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType HttpResponseMessage -ErrorAction Stop
        if ($csv -is [System.Net.Http.HttpResponseMessage]) {
            $csv = $csv.Content.ReadAsStringAsync().Result
        }
    } catch {
        Write-Fail "Usage report fetch failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'grove' -Message "getSharePointSiteUsageDetail failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    if (-not $csv) {
        Write-Warn "Usage report returned empty content."
        return @()
    }

    # The first row is the header. Parse with ConvertFrom-Csv directly on the
    # multi-line string; the endpoint emits a UTF-8 CSV with a fixed column set.
    $rows = $null
    try {
        $rows = $csv | ConvertFrom-Csv
    } catch {
        Write-Fail "Could not parse CSV response: $($_.Exception.Message)"
        return @()
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warn "Usage report has no rows — tenant may have no SharePoint sites, or report data is still aggregating."
        return @()
    }

    $records = foreach ($r in $rows) {
        # Column names include "Storage Used (Byte)" etc. with spaces and units,
        # so we reach for them by name and coerce defensively.
        $storageBytes = 0
        $files        = 0
        $activeFiles  = 0
        try { $storageBytes = [int64]$r.'Storage Used (Byte)' } catch { }
        try { $files        = [int]$r.'File Count' } catch { }
        try { $activeFiles  = [int]$r.'Active File Count' } catch { }

        $lastActivity = $null
        if ($r.'Last Activity Date') {
            try { $lastActivity = [datetime]$r.'Last Activity Date' } catch { }
        }

        [PSCustomObject]@{
            SiteId            = $r.'Site Id'
            SiteUrl           = $r.'Site URL'
            OwnerDisplayName  = $r.'Owner Display Name'
            OwnerPrincipal    = $r.'Owner Principal Name'
            SiteType          = $r.'Root Web Template'
            StorageBytes      = $storageBytes
            StorageSize       = Format-Bytes $storageBytes
            FileCount         = $files
            ActiveFileCount   = $activeFiles
            LastActivity      = $lastActivity
            LastActivityStr   = if ($lastActivity) { $lastActivity.ToString('yyyy-MM-dd') } else { 'Never' }
            ExternalSharing   = ($r.'External Sharing' -eq 'True')
        }
    }

    $records = @($records)
    Write-Ok "Collected $($records.Count) site(s) from the usage report."
    Write-Host ""
    return $records
}

# Tenant-wide SharePoint sharing policy. Emits a single record with the three
# top-level toggles that govern every site in the tenant.
function Get-TenantSharingPolicy {
    Write-Step "Reading tenant sharing policy (admin/sharepoint/settings)..."
    $uri = 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'

    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        return [PSCustomObject]@{
            SharingCapability             = $resp.sharingCapability
            SharingDomainRestrictionMode  = $resp.sharingDomainRestrictionMode
            DefaultSharingLinkType        = $resp.defaultSharingLinkType
            DefaultLinkPermission         = $resp.defaultLinkPermission
            AvailableManagedPathsForSiteCreation = $resp.availableManagedPathsForSiteCreation
            CollectorError                = $null
        }
    } catch {
        return [PSCustomObject]@{
            SharingCapability            = 'unknown'
            SharingDomainRestrictionMode = 'unknown'
            DefaultSharingLinkType       = 'unknown'
            DefaultLinkPermission        = 'unknown'
            CollectorError               = $_.Exception.Message
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# DERIVED FINDINGS
# ─────────────────────────────────────────────────────────────────────────────

$script:LargeSiteThresholdBytes = 100GB
$script:StaleActivityDays       = 180

function Get-LargeSites   { param([array]$S) @($S | Where-Object { $_.StorageBytes -ge $script:LargeSiteThresholdBytes } | Sort-Object StorageBytes -Descending) }
function Get-SharedSites  { param([array]$S) @($S | Where-Object { $_.ExternalSharing } | Sort-Object SiteUrl) }
function Get-OwnerlessSites { param([array]$S) @($S | Where-Object { -not $_.OwnerDisplayName -or $_.OwnerDisplayName -eq '' } | Sort-Object SiteUrl) }
function Get-StaleSites {
    param([array]$S)
    $threshold = (Get-Date).AddDays(-$script:StaleActivityDays)
    @($S | Where-Object { -not $_.LastActivity -or $_.LastActivity -lt $threshold } | Sort-Object LastActivityStr)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Sites, $Policy, [array]$Large, [array]$Shared, [array]$Ownerless, [array]$Stale)

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $tkCfg         = Get-TKConfig
    $orgPrefix     = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalBytes = ($Sites | Measure-Object -Property StorageBytes -Sum).Sum
    $totalFiles = ($Sites | Measure-Object -Property FileCount    -Sum).Sum

    function _site-row {
        param($s)
        $sizeClass = if ($s.StorageBytes -ge $script:LargeSiteThresholdBytes) { 'tk-badge-warn' }
                     elseif ($s.StorageBytes -ge 10GB) { 'tk-badge-info' }
                     else { 'tk-badge-ok' }
        $shareBadge = if ($s.ExternalSharing) { "<span class='tk-badge-warn'>Enabled</span>" } else { "<span class='tk-badge-ok'>Disabled</span>" }
        $ownerCell  = if ($s.OwnerDisplayName) { EscHtml $s.OwnerDisplayName } else { "<span class='tk-badge-err'>(no owner)</span>" }
        $activeClass = if (-not $s.LastActivity) { 'tk-badge-warn' }
                       elseif ($s.LastActivity -lt (Get-Date).AddDays(-$script:StaleActivityDays)) { 'tk-badge-warn' }
                       else { 'tk-badge-ok' }
        return "<tr><td><code>$(EscHtml $s.SiteUrl)</code></td><td>$(EscHtml $s.SiteType)</td><td>$ownerCell</td><td><span class='$sizeClass'>$(EscHtml $s.StorageSize)</span></td><td>$($s.FileCount)</td><td><span class='$activeClass'>$(EscHtml $s.LastActivityStr)</span></td><td>$shareBadge</td></tr>"
    }

    $invRows = [System.Text.StringBuilder]::new()
    if ($Sites.Count -eq 0) {
        [void]$invRows.Append("<tr><td colspan='7' class='tk-badge-info' style='text-align:center;'>No SharePoint sites returned by the usage report.</td></tr>")
    } else {
        foreach ($s in ($Sites | Sort-Object StorageBytes -Descending)) { [void]$invRows.Append((_site-row $s) + "`n") }
    }

    function _category-rows {
        param([array]$Rows, [int]$Cols, [string]$EmptyMessage)
        $sb = [System.Text.StringBuilder]::new()
        if ($Rows.Count -eq 0) {
            [void]$sb.Append("<tr><td colspan='$Cols' class='tk-badge-ok' style='text-align:center;'>$EmptyMessage</td></tr>")
        } else {
            foreach ($s in $Rows) { [void]$sb.Append((_site-row $s) + "`n") }
        }
        return $sb.ToString()
    }

    $largeRows     = _category-rows -Rows $Large     -Cols 7 -EmptyMessage "No sites exceed $(Format-Bytes $script:LargeSiteThresholdBytes)."
    $sharedRows    = _category-rows -Rows $Shared    -Cols 7 -EmptyMessage 'No sites have external sharing enabled.'
    $ownerlessRows = _category-rows -Rows $Ownerless -Cols 7 -EmptyMessage 'No sites without an owner.'
    $staleRows     = _category-rows -Rows $Stale     -Cols 7 -EmptyMessage "No sites with activity older than $script:StaleActivityDays days."

    # Tenant sharing policy table
    $policyClass = switch ($Policy.SharingCapability) {
        'disabled'                     { 'tk-badge-ok' }
        'externalUserSharingOnly'      { 'tk-badge-warn' }
        'externalUserAndGuestSharing'  { 'tk-badge-warn' }
        'existingExternalUserSharingOnly' { 'tk-badge-warn' }
        default                        { 'tk-badge-info' }
    }

    $policyErr = if ($Policy.CollectorError) {
        "<tr><th>Policy Read Error</th><td><span class='tk-badge-warn'>$(EscHtml $Policy.CollectorError)</span></td></tr>"
    } else { '' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'G.R.O.V.E. SharePoint Inventory Report' `
        -ScriptName 'G.R.O.V.E.' `
        -Subtitle   "${orgPrefix}SharePoint Online Estate Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = EscHtml $script:ConnectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @('Sharing Policy', 'Site Inventory', 'Large Sites', 'External Sharing', 'Ownerless Sites', 'Stale Sites')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'G.R.O.V.E. v3.0'

    $largeClass     = if ($Large.Count -gt 0)     { 'warn' } else { 'ok' }
    $sharedClass    = if ($Shared.Count -gt 0)    { 'warn' } else { 'ok' }
    $ownerlessClass = if ($Ownerless.Count -gt 0) { 'err'  } else { 'ok' }
    $staleClass     = if ($Stale.Count -gt 0)     { 'warn' } else { 'ok' }

    $invHeader = "<tr><th>Site URL</th><th>Template</th><th>Owner</th><th>Storage</th><th>Files</th><th>Last Activity</th><th>External Sharing</th></tr>"

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Sites.Count)</div><div class="tk-summary-lbl">Sites Total</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(Format-Bytes $totalBytes)</div><div class="tk-summary-lbl">Total Storage</div></div>
    <div class="tk-summary-card $largeClass"><div class="tk-summary-num">$($Large.Count)</div><div class="tk-summary-lbl">Large Sites (&ge; $(Format-Bytes $script:LargeSiteThresholdBytes))</div></div>
    <div class="tk-summary-card $sharedClass"><div class="tk-summary-num">$($Shared.Count)</div><div class="tk-summary-lbl">External Sharing On</div></div>
    <div class="tk-summary-card $ownerlessClass"><div class="tk-summary-num">$($Ownerless.Count)</div><div class="tk-summary-lbl">Ownerless</div></div>
    <div class="tk-summary-card $staleClass"><div class="tk-summary-num">$($Stale.Count)</div><div class="tk-summary-lbl">Stale (&gt; $script:StaleActivityDays d)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Tenant Sharing Policy</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Sharing Capability</th><td><span class='$policyClass'>$(EscHtml $Policy.SharingCapability)</span></td></tr>
          <tr><th>Domain Restriction Mode</th><td>$(EscHtml $Policy.SharingDomainRestrictionMode)</td></tr>
          <tr><th>Default Link Type</th><td>$(EscHtml $Policy.DefaultSharingLinkType)</td></tr>
          <tr><th>Default Link Permission</th><td>$(EscHtml $Policy.DefaultLinkPermission)</td></tr>
          $policyErr
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Site Inventory</span><span class="tk-section-num">$($Sites.Count) site(s)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$($invRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Large Sites -- &ge; $(Format-Bytes $script:LargeSiteThresholdBytes)</span><span class="tk-section-num">$($Large.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$largeRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Sites With External Sharing Enabled</span><span class="tk-section-num">$($Shared.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$sharedRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Ownerless Sites</span><span class="tk-section-num">$($Ownerless.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$ownerlessRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Stale Sites -- No Activity &gt; $script:StaleActivityDays Days</span><span class="tk-section-num">$($Stale.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$staleRows</tbody></table></div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Sites, $Policy, [array]$Large, [array]$Shared, [array]$Ownerless, [array]$Stale)

    Write-Section "EXPORTING HTML REPORT"
    $html      = Build-HtmlReport -Sites $Sites -Policy $Policy -Large $Large -Shared $Shared -Ownerless $Ownerless -Stale $Stale
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "GROVE_${timestamp}.html"

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
    Show-GroveBanner
    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  SharePoint inventory (refresh underlying data)" -ForegroundColor $C.Info
    Write-Host "  3  Export full HTML report (all findings)" -ForegroundColor $C.Success
    Write-Host "  Q  Quit" -ForegroundColor $C.Info
    Write-Host ""
}

function Assert-Connected {
    if (-not $script:Connected) {
        Write-Host ""; Write-Warn "Not connected. Please select option 1 first."; Write-Host ""
        Read-Host "  Press Enter to return to menu"
        return $false
    }
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

Show-GroveBanner
Install-GraphModule

if (Test-GraphConnection) { Write-Ok "Already connected as $($script:ConnectedAs)"; Write-Host "" }

if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    if (-not $script:Connected) { Invoke-Connect }
    if (-not $script:Connected) { Write-Fail "Could not connect to Graph. Exiting."; exit 1 }

    $sites  = Get-SharePointInventory
    $policy = Get-TenantSharingPolicy
    $large  = Get-LargeSites      -S $sites
    $shared = Get-SharedSites     -S $sites
    $owner  = Get-OwnerlessSites  -S $sites
    $stale  = Get-StaleSites      -S $sites

    Export-HtmlReport -Sites $sites -Policy $policy -Large $large -Shared $shared -Ownerless $owner -Stale $stale

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

$cachedSites  = $null
$cachedPolicy = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {
        '1' {
            Invoke-Connect
            $cachedSites = $null; $cachedPolicy = $null
            Read-Host "  Press Enter to return to menu"
        }
        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedSites  = Get-SharePointInventory
            $cachedPolicy = Get-TenantSharingPolicy
            Read-Host "  Press Enter to return to menu"
        }
        '3' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedSites)  { $cachedSites  = Get-SharePointInventory }
            if ($null -eq $cachedPolicy) { $cachedPolicy = Get-TenantSharingPolicy }
            $large  = Get-LargeSites      -S $cachedSites
            $shared = Get-SharedSites     -S $cachedSites
            $owner  = Get-OwnerlessSites  -S $cachedSites
            $stale  = Get-StaleSites      -S $cachedSites
            Export-HtmlReport -Sites $cachedSites -Policy $cachedPolicy -Large $large -Shared $shared -Ownerless $owner -Stale $stale
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
