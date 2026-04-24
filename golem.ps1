<#
.SYNOPSIS
    G.O.L.E.M. — Governs & Observes Licensed Endpoint Management
    Microsoft Intune / MDM Compliance Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and audits the Intune-managed device estate.
    Reports enrollment breakdown (ownership, OS, compliance state), stale
    devices that have not checked in recently, and configuration-profile
    assignment coverage. Produces a dark-themed HTML report and a console
    summary for each section.

.USAGE
    PS C:\> .\golem.ps1                  # Interactive menu
    PS C:\> .\golem.ps1 -Unattended      # Auto-connect and export full audit report

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

$script:Connected    = $false
$script:ConnectedAs  = ''
$script:TenantDomain = ''

$GraphScopes = @(
    'DeviceManagementManagedDevices.Read.All',
    'DeviceManagementConfiguration.Read.All',
    'Directory.Read.All'
)

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-GolemBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗  ██████╗ ██╗     ███████╗███╗   ███╗
  ██╔════╝ ██╔═══██╗██║     ██╔════╝████╗ ████║
  ██║  ███╗██║   ██║██║     █████╗  ██╔████╔██║
  ██║   ██║██║   ██║██║     ██╔══╝  ██║╚██╔╝██║
  ╚██████╔╝╚██████╔╝███████╗███████╗██║ ╚═╝ ██║
   ╚═════╝  ╚═════╝ ╚══════╝╚══════╝╚═╝     ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "  G.O.L.E.M. — Governs & Observes Licensed Endpoint Management" -ForegroundColor Cyan
    Write-Host "  Microsoft Intune / MDM Compliance Audit Tool  v3.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH MODULE INSTALL + CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $required = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.Identity.DirectoryManagement'
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
                Write-TKError -ScriptName 'golem' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
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
            Write-Fail "Connection returned no context — authentication may have been cancelled."
        }
    } catch {
        Write-Fail "Authentication failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'golem' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION
# ─────────────────────────────────────────────────────────────────────────────

function Get-DeviceInventory {
    Write-Section "DEVICE INVENTORY"
    Write-Step "Retrieving managed devices from Intune..."

    try {
        $devices = Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve devices: $($_.Exception.Message)"
        Write-TKError -ScriptName 'golem' -Message "Get-MgDeviceManagementManagedDevice failed: $($_.Exception.Message)" -Category 'Intune Query'
        return @()
    }

    if (-not $devices -or $devices.Count -eq 0) {
        Write-Warn "No managed devices returned. The tenant may have no Intune enrollments."
        return @()
    }

    $now = Get-Date
    $enriched = foreach ($d in $devices) {
        $lastSync = if ($d.LastSyncDateTime) { [datetime]$d.LastSyncDateTime } else { $null }
        $daysSinceSync = if ($lastSync) { [math]::Round(($now - $lastSync).TotalDays, 0) } else { $null }

        [PSCustomObject]@{
            DeviceName      = $d.DeviceName
            UserPrincipal   = $d.UserPrincipalName
            OS              = $d.OperatingSystem
            OSVersion       = $d.OSVersion
            Ownership       = $d.ManagedDeviceOwnerType
            EnrolledDate    = if ($d.EnrolledDateTime) { ([datetime]$d.EnrolledDateTime).ToString('yyyy-MM-dd') } else { '—' }
            LastSync        = if ($lastSync) { $lastSync.ToString('yyyy-MM-dd') } else { 'Never' }
            DaysSinceSync   = $daysSinceSync
            ComplianceState = if ($d.ComplianceState) { $d.ComplianceState } else { 'unknown' }
            JoinType        = if ($d.JoinType) { $d.JoinType } else { '—' }
            Model           = if ($d.Model) { $d.Model } else { '—' }
            Manufacturer    = if ($d.Manufacturer) { $d.Manufacturer } else { '—' }
        }
    }

    Write-Ok "$($enriched.Count) managed device(s) retrieved."
    Write-Host ""

    # Console breakdown by OS
    $byOs = $enriched | Group-Object OS | Sort-Object Count -Descending
    Write-Host ("  {0,-22} {1,8}" -f "Operating System", "Count") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 32)) -ForegroundColor $C.Info
    foreach ($g in $byOs) {
        $osLabel = if ($g.Name) { $g.Name } else { '(unknown)' }
        Write-Host ("  {0,-22} {1,8}" -f $osLabel, $g.Count) -ForegroundColor $C.Info
    }
    Write-Host ""

    return @($enriched)
}

function Get-ComplianceSummary {
    param([array]$Devices)

    Write-Section "COMPLIANCE STATE"

    if ($Devices.Count -eq 0) {
        Write-Warn "No devices to summarise."
        return @{}
    }

    $counts = [ordered]@{
        compliant    = 0
        noncompliant = 0
        inGracePeriod = 0
        configManager = 0
        error        = 0
        unknown      = 0
    }

    foreach ($d in $Devices) {
        $state = if ($d.ComplianceState) { $d.ComplianceState.ToString() } else { 'unknown' }
        if ($counts.Contains($state)) { $counts[$state]++ } else { $counts['unknown']++ }
    }

    Write-Host ("  {0,-18} {1,8}" -f "State", "Count") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 28)) -ForegroundColor $C.Info
    foreach ($k in $counts.Keys) {
        $color = switch ($k) {
            'compliant'    { $C.Success }
            'noncompliant' { $C.Error }
            'error'        { $C.Error }
            default        { $C.Warning }
        }
        Write-Host ("  {0,-18} {1,8}" -f $k, $counts[$k]) -ForegroundColor $color
    }
    Write-Host ""

    return $counts
}

function Get-StaleDevices {
    param([array]$Devices)

    Write-Section "STALE DEVICES"

    if ($Devices.Count -eq 0) {
        Write-Warn "No devices to evaluate."
        return @()
    }

    # Buckets: 30+ days, 60+ days, 90+ days (not mutually exclusive — each device lands in its deepest bucket)
    $stale = foreach ($d in $Devices) {
        if ($null -ne $d.DaysSinceSync -and $d.DaysSinceSync -ge 30) { $d }
    }
    $stale = @($stale | Sort-Object DaysSinceSync -Descending)

    if ($stale.Count -eq 0) {
        Write-Ok "No devices have been silent for 30+ days."
        return @()
    }

    $b30 = @($stale | Where-Object { $_.DaysSinceSync -lt 60 }).Count
    $b60 = @($stale | Where-Object { $_.DaysSinceSync -ge 60 -and $_.DaysSinceSync -lt 90 }).Count
    $b90 = @($stale | Where-Object { $_.DaysSinceSync -ge 90 }).Count

    Write-Host ("  30-59 days:  {0}" -f $b30) -ForegroundColor $C.Warning
    Write-Host ("  60-89 days:  {0}" -f $b60) -ForegroundColor $C.Warning
    Write-Host ("  90+  days:  {0}" -f $b90) -ForegroundColor $C.Error
    Write-Host ""
    Write-Warn "$($stale.Count) stale device(s) total."
    return $stale
}

function Get-ConfigurationProfileAssignments {
    Write-Section "CONFIGURATION PROFILES"
    Write-Step "Retrieving configuration profiles and assignments..."

    try {
        $profiles = Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve configuration profiles: $($_.Exception.Message)"
        return @()
    }

    if (-not $profiles -or $profiles.Count -eq 0) {
        Write-Warn "No device configuration profiles found."
        return @()
    }

    $rows = foreach ($p in $profiles) {
        $assignmentCount = 0
        try {
            $assignments = Get-MgDeviceManagementDeviceConfigurationAssignment `
                -DeviceConfigurationId $p.Id -ErrorAction Stop
            $assignmentCount = @($assignments).Count
        } catch {
            # Assignment read failure is non-fatal; profile row still rendered with Count=0.
        }

        [PSCustomObject]@{
            DisplayName     = $p.DisplayName
            Id              = $p.Id
            AssignmentCount = $assignmentCount
            LastModified    = if ($p.LastModifiedDateTime) { ([datetime]$p.LastModifiedDateTime).ToString('yyyy-MM-dd') } else { '—' }
        }
    }

    $unassigned = @($rows | Where-Object { $_.AssignmentCount -eq 0 })

    Write-Ok "$($rows.Count) configuration profile(s)."
    if ($unassigned.Count -gt 0) {
        Write-Warn "$($unassigned.Count) profile(s) have no assignments."
    } else {
        Write-Ok "All profiles are assigned."
    }
    Write-Host ""

    return @($rows | Sort-Object DisplayName)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$Devices,
        [hashtable]$Compliance,
        [array]$Stale,
        [array]$Profiles
    )

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalDevices = $Devices.Count
    $compCount    = [int]($Compliance['compliant'])
    $noncompCount = [int]($Compliance['noncompliant']) + [int]($Compliance['error'])
    $graceCount   = [int]($Compliance['inGracePeriod'])
    $unknownCount = [int]($Compliance['unknown'])
    $unassigned   = @($Profiles | Where-Object { $_.AssignmentCount -eq 0 }).Count

    # ── Device rows (limit to newest 200 to keep HTML manageable) ──
    $deviceRows = [System.Text.StringBuilder]::new()
    foreach ($d in ($Devices | Sort-Object DaysSinceSync)) {
        $compClass = switch ($d.ComplianceState) {
            'compliant'    { 'tk-badge-ok' }
            'noncompliant' { 'tk-badge-err' }
            'error'        { 'tk-badge-err' }
            'inGracePeriod'{ 'tk-badge-warn' }
            default        { 'tk-badge-info' }
        }
        $daysLabel = if ($null -eq $d.DaysSinceSync) { 'n/a' } else { "$($d.DaysSinceSync)d" }
        $syncClass = if ($null -eq $d.DaysSinceSync) { 'tk-badge-warn' }
                     elseif ($d.DaysSinceSync -ge 90) { 'tk-badge-err' }
                     elseif ($d.DaysSinceSync -ge 30) { 'tk-badge-warn' }
                     else { 'tk-badge-ok' }
        [void]$deviceRows.Append("<tr><td>$(EscHtml $d.DeviceName)</td><td>$(EscHtml $d.UserPrincipal)</td><td>$(EscHtml $d.OS) $(EscHtml $d.OSVersion)</td><td>$(EscHtml $d.Ownership)</td><td><span class='$compClass'>$(EscHtml $d.ComplianceState)</span></td><td>$(EscHtml $d.LastSync)</td><td><span class='$syncClass'>$daysLabel</span></td></tr>`n")
    }

    # ── Stale rows ──
    $staleRows = [System.Text.StringBuilder]::new()
    if ($Stale.Count -eq 0) {
        [void]$staleRows.Append("<tr><td colspan='5' class='tk-badge-ok' style='text-align:center;'>No stale devices.</td></tr>")
    } else {
        foreach ($d in $Stale) {
            $bucketClass = if ($d.DaysSinceSync -ge 90) { 'tk-badge-err' }
                           elseif ($d.DaysSinceSync -ge 60) { 'tk-badge-warn' }
                           else { 'tk-badge-warn' }
            [void]$staleRows.Append("<tr><td>$(EscHtml $d.DeviceName)</td><td>$(EscHtml $d.UserPrincipal)</td><td>$(EscHtml $d.OS)</td><td>$(EscHtml $d.LastSync)</td><td><span class='$bucketClass'>$($d.DaysSinceSync)d</span></td></tr>`n")
        }
    }

    # ── Profile rows ──
    $profileRows = [System.Text.StringBuilder]::new()
    if ($Profiles.Count -eq 0) {
        [void]$profileRows.Append("<tr><td colspan='3' class='tk-badge-info' style='text-align:center;'>No configuration profiles found.</td></tr>")
    } else {
        foreach ($p in $Profiles) {
            $aClass = if ($p.AssignmentCount -eq 0) { 'tk-badge-warn' } else { 'tk-badge-ok' }
            [void]$profileRows.Append("<tr><td>$(EscHtml $p.DisplayName)</td><td><span class='$aClass'>$($p.AssignmentCount)</span></td><td>$(EscHtml $p.LastModified)</td></tr>`n")
        }
    }

    $compClass       = if ($noncompCount -gt 0) { 'err' } else { 'ok' }
    $staleClass      = if ($Stale.Count -gt 0) { 'warn' } else { 'ok' }
    $unassignedClass = if ($unassigned -gt 0) { 'warn' } else { 'ok' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'G.O.L.E.M. Intune Compliance Report' `
        -ScriptName 'G.O.L.E.M.' `
        -Subtitle   "${orgPrefix}Microsoft Intune / MDM Compliance Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Tenant'       = $tenantDisplay
            'Connected As' = EscHtml $script:ConnectedAs
        }) `
        -NavItems   @('Devices', 'Stale Devices', 'Configuration Profiles')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'G.O.L.E.M. v3.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalDevices</div><div class="tk-summary-lbl">Managed Devices</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$compCount</div><div class="tk-summary-lbl">Compliant</div></div>
    <div class="tk-summary-card $compClass"><div class="tk-summary-num">$noncompCount</div><div class="tk-summary-lbl">Non-compliant / Error</div></div>
    <div class="tk-summary-card warn"><div class="tk-summary-num">$graceCount</div><div class="tk-summary-lbl">In Grace Period</div></div>
    <div class="tk-summary-card $staleClass"><div class="tk-summary-num">$($Stale.Count)</div><div class="tk-summary-lbl">Stale (30d+)</div></div>
    <div class="tk-summary-card $unassignedClass"><div class="tk-summary-num">$unassigned</div><div class="tk-summary-lbl">Unassigned Profiles</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Managed Devices</span>
      <span class="tk-section-num">$totalDevices device(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Device Name</th><th>User</th><th>OS</th><th>Ownership</th><th>Compliance</th><th>Last Sync</th><th>Days Since</th></tr></thead>
        <tbody>
          $($deviceRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Stale Devices -- No Sync Within 30 Days</span>
      <span class="tk-section-num">$($Stale.Count) device(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Device Name</th><th>User</th><th>OS</th><th>Last Sync</th><th>Days Since</th></tr></thead>
        <tbody>
          $($staleRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Configuration Profiles</span>
      <span class="tk-section-num">$($Profiles.Count) profile(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>Assignments</th><th>Last Modified</th></tr></thead>
        <tbody>
          $($profileRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Devices, [hashtable]$Compliance, [array]$Stale, [array]$Profiles)

    Write-Section "EXPORTING HTML REPORT"
    Write-Step "Building report..."

    $html      = Build-HtmlReport -Devices $Devices -Compliance $Compliance -Stale $Stale -Profiles $Profiles
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "GOLEM_${timestamp}.html"

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Write-Ok "Report saved: $outPath"
        Write-Step "Opening in default browser..."
        Start-Process $outPath
    } catch {
        Write-Fail "Could not save report: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-GolemBanner

    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  Device inventory — managed devices grouped by OS" -ForegroundColor $C.Info
    Write-Host "  3  Compliance state — per-state counts" -ForegroundColor $C.Info
    Write-Host "  4  Stale devices — no sync in 30+ / 60+ / 90+ days" -ForegroundColor $C.Warning
    Write-Host "  5  Configuration profiles — assignment coverage" -ForegroundColor $C.Info
    Write-Host "  6  Export full HTML report (all findings)" -ForegroundColor $C.Success
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

Show-GolemBanner
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

    $devices    = Get-DeviceInventory
    $compliance = Get-ComplianceSummary -Devices $devices
    $stale      = Get-StaleDevices       -Devices $devices
    $profiles   = Get-ConfigurationProfileAssignments

    Export-HtmlReport -Devices $devices -Compliance $compliance -Stale $stale -Profiles $profiles

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# Cached per-session so repeated menu choices don't re-query
$cachedDevices    = $null
$cachedCompliance = $null
$cachedStale      = $null
$cachedProfiles   = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {

        '1' {
            Invoke-Connect
            $cachedDevices = $null; $cachedCompliance = $null; $cachedStale = $null; $cachedProfiles = $null
            Read-Host "  Press Enter to return to menu"
        }

        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedDevices = Get-DeviceInventory
            Read-Host "  Press Enter to return to menu"
        }

        '3' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedDevices) { $cachedDevices = Get-DeviceInventory }
            $cachedCompliance = Get-ComplianceSummary -Devices $cachedDevices
            Read-Host "  Press Enter to return to menu"
        }

        '4' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedDevices) { $cachedDevices = Get-DeviceInventory }
            $cachedStale = Get-StaleDevices -Devices $cachedDevices
            Read-Host "  Press Enter to return to menu"
        }

        '5' {
            if (-not (Assert-Connected)) { break }
            $cachedProfiles = Get-ConfigurationProfileAssignments
            Read-Host "  Press Enter to return to menu"
        }

        '6' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedDevices)    { $cachedDevices    = Get-DeviceInventory }
            if ($null -eq $cachedCompliance) { $cachedCompliance = Get-ComplianceSummary -Devices $cachedDevices }
            if ($null -eq $cachedStale)      { $cachedStale      = Get-StaleDevices -Devices $cachedDevices }
            if ($null -eq $cachedProfiles)   { $cachedProfiles   = Get-ConfigurationProfileAssignments }
            Export-HtmlReport -Devices $cachedDevices -Compliance $cachedCompliance -Stale $cachedStale -Profiles $cachedProfiles
            Read-Host "  Press Enter to return to menu"
        }

        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {
                # Disconnect is best-effort — cached context may already be gone.
            }
            Write-Host "  Goodbye." -ForegroundColor $C.Header
            Write-Host ""
        }

        default {
            Write-Host ""
            Write-Warn "Invalid option. Please choose 1–6 or Q."
            Write-Host ""
            Start-Sleep -Milliseconds 800
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
