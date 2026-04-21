<#
.SYNOPSIS
    W.A.R.D. — Watches Accounts, Reviews Roles & Detects anomalies
    User Account & Security Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits all local user accounts on the machine. Reports account status,
    last logon time, password configuration, and group memberships. Flags
    potentially risky accounts and exports a dark-themed HTML report to the
    script directory.

.USAGE
    PS C:\> .\ward.ps1                    # Must be run as Administrator
    PS C:\> .\ward.ps1 -Unattended        # Silent mode — no prompts, no banner

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
    R.E.L.I.C.             — Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.           — Toolkit setup & configuration wizard

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
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

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-WardBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗    ██╗ █████╗ ██████╗ ██████╗
  ██║    ██║██╔══██╗██╔══██╗██╔══██╗
  ██║ █╗ ██║███████║██████╔╝██║  ██║
  ██║███╗██║██╔══██║██╔══██╗██║  ██║
  ╚███╔███╔╝██║  ██║██║  ██║██████╔╝
   ╚══╝╚══╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝

"@ -ForegroundColor Cyan
    Write-Host "    W.A.R.D. — Watches Accounts, Reviews Roles & Detects anomalies" -ForegroundColor Cyan
    Write-Host "    User Account & Local Security Audit Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION
# ─────────────────────────────────────────────────────────────────────────────

function Get-AdminMembers {
    try {
        $admins = Get-LocalGroupMember -Group "Administrators" -ErrorAction Stop
        return $admins | ForEach-Object { ($_.Name -split '\\')[-1] }
    }
    catch {
        return @()
    }
}

function Get-AccountData {
    param([string[]]$AdminNames)

    $staleDays  = 90
    $staleDate  = (Get-Date).AddDays(-$staleDays)
    $accounts   = @()

    $localUsers = Get-LocalUser -ErrorAction SilentlyContinue

    foreach ($user in $localUsers) {
        $isAdmin   = $AdminNames -contains $user.Name
        $lastLogon = if ($user.LastLogon) { $user.LastLogon } else { $null }

        $flags = @()

        if ($user.Enabled -and -not $user.PasswordRequired) {
            $flags += "No password required"
        }
        if ($user.Enabled -and -not $user.PasswordLastSet) {
            $flags += "Password never set"
        }
        if ($user.Enabled -and (-not $lastLogon -or $lastLogon -lt $staleDate)) {
            $flags += "Stale (>$staleDays days)"
        }
        if (-not $user.Enabled) {
            $flags += "Disabled"
        }

        $accounts += [PSCustomObject]@{
            Name              = $user.Name
            FullName          = $user.FullName
            Enabled           = $user.Enabled
            IsAdmin           = $isAdmin
            LastLogon         = if ($lastLogon) { $lastLogon.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
            PasswordLastSet   = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
            PasswordExpires   = if ($user.PasswordExpires) { $user.PasswordExpires.ToString("yyyy-MM-dd") } else { "Never / No Expiry" }
            PasswordRequired  = $user.PasswordRequired
            Description       = $user.Description
            Flags             = if ($flags.Count -gt 0) { $flags -join '; ' } else { "" }
        }
    }

    return $accounts
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────

function HtmlEncode([string]$s) {
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Build-HtmlReport {
    param([array]$Accounts, [string]$MachineName, [string]$ReportTimestamp)

    $totalAccounts   = $Accounts.Count
    $enabledCount    = ($Accounts | Where-Object { $_.Enabled  } | Measure-Object).Count
    $disabledCount   = ($Accounts | Where-Object { -not $_.Enabled } | Measure-Object).Count
    $adminCount      = ($Accounts | Where-Object { $_.IsAdmin  } | Measure-Object).Count
    $flaggedCount    = ($Accounts | Where-Object { $_.Flags    } | Measure-Object).Count

    # Build account rows
    $rows = ""
    foreach ($acct in ($Accounts | Sort-Object IsAdmin -Descending)) {
        $enabledBadge = if ($acct.Enabled) {
            "<span class='tk-badge-ok'>Enabled</span>"
        } else {
            "<span class='tk-badge-warn'>Disabled</span>"
        }
        $adminBadge = if ($acct.IsAdmin) {
            "<span class='tk-badge-err'>Admin</span>"
        } else {
            "<span class='tk-badge-info'>Standard</span>"
        }
        $flagCell = if ($acct.Flags) {
            "<span class='tk-badge-warn'>$(HtmlEncode($acct.Flags))</span>"
        } else { "" }

        $rows += @"
            <tr>
                <td><strong>$(HtmlEncode($acct.Name))</strong></td>
                <td>$(HtmlEncode($acct.FullName))</td>
                <td>$enabledBadge</td>
                <td>$adminBadge</td>
                <td>$(HtmlEncode($acct.LastLogon))</td>
                <td>$(HtmlEncode($acct.PasswordLastSet))</td>
                <td>$(HtmlEncode($acct.PasswordExpires))</td>
                <td>$flagCell</td>
            </tr>
"@
    }

    $tkConfig  = Get-TKConfig
    $tkOrgName = if (-not [string]::IsNullOrWhiteSpace($tkConfig.OrgName)) { EscHtml $tkConfig.OrgName } else { $null }
    $tkSubtitle = if ($tkOrgName) { "$tkOrgName -- $MachineName" } else { $MachineName }

    $tkMetaItems = [ordered]@{
        'Machine'   = $MachineName
        'Generated' = $ReportTimestamp
        'Accounts'  = $totalAccounts
        'Flagged'   = $flaggedCount
    }

    $tkNavItems = @('Local User Accounts')

    $flaggedClass    = if ($flaggedCount -gt 0) { "err" } else { "ok" }
    $adminClass      = if ($adminCount -gt 1)   { "warn" } else { "info" }

    $summaryCards = @"
<div class="tk-summary-row">
  <div class="tk-summary-card info">
    <div class="tk-summary-num">$totalAccounts</div>
    <div class="tk-summary-lbl">Total Accounts</div>
  </div>
  <div class="tk-summary-card ok">
    <div class="tk-summary-num">$enabledCount</div>
    <div class="tk-summary-lbl">Enabled</div>
  </div>
  <div class="tk-summary-card">
    <div class="tk-summary-num">$disabledCount</div>
    <div class="tk-summary-lbl">Disabled</div>
  </div>
  <div class="tk-summary-card $adminClass">
    <div class="tk-summary-num">$adminCount</div>
    <div class="tk-summary-lbl">Administrators</div>
  </div>
  <div class="tk-summary-card $flaggedClass">
    <div class="tk-summary-num">$flaggedCount</div>
    <div class="tk-summary-lbl">Flagged</div>
  </div>
</div>
"@

    $html = (Get-TKHtmlHead `
        -Title      'Account Audit Report' `
        -ScriptName 'W.A.R.D.' `
        -Subtitle   $tkSubtitle `
        -MetaItems  $tkMetaItems `
        -NavItems   $tkNavItems) + @"

  $summaryCards

  <div class="tk-section" id="local-user-accounts">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 1</span>
        <h2 class="tk-section-title">Local User Accounts</h2>
      </div>
      <div style="padding:20px;">
        <table class="tk-table">
          <thead>
            <tr>
              <th>Username</th>
              <th>Full Name</th>
              <th>Status</th>
              <th>Role</th>
              <th>Last Logon</th>
              <th>Password Set</th>
              <th>Password Expires</th>
              <th>Flags</th>
            </tr>
          </thead>
          <tbody>
            $rows
          </tbody>
        </table>
        <div class="tk-info-box" style="margin-top:18px;">
          <span class="tk-info-label">Note</span> Stale threshold: 90 days without logon
        </div>
      </div>
    </div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'W.A.R.D. v1.0')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-WardBanner }

$machineName      = $env:COMPUTERNAME
$reportTimestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  COLLECTING ACCOUNT DATA" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

Write-Host "  [*] Resolving Administrators group members..." -ForegroundColor $ColorSchema.Progress
$adminNames = Get-AdminMembers

Write-Host "  [*] Enumerating local user accounts..." -ForegroundColor $ColorSchema.Progress
$accounts = Get-AccountData -AdminNames $adminNames

Write-Host "  [+] Found $($accounts.Count) local user account(s)." -ForegroundColor $ColorSchema.Success
Write-Host ""

# Console summary
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  ACCOUNT OVERVIEW" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

foreach ($acct in ($accounts | Sort-Object IsAdmin -Descending)) {
    $statusColor = if ($acct.Enabled) { $ColorSchema.Success } else { $ColorSchema.Info }
    $roleLabel   = if ($acct.IsAdmin) { " [ADMIN]" } else { "" }
    $flagLabel   = if ($acct.Flags)   { "  [!!] $($acct.Flags)" } else { "" }

    Write-Host ("  {0,-22} Enabled: {1,-6} Last Logon: {2}{3}" -f `
        ($acct.Name + $roleLabel), $acct.Enabled, $acct.LastLogon, "") -ForegroundColor $statusColor

    if ($acct.Flags) {
        Write-Host ("  {0,-22} {1}" -f "", $flagLabel.Trim()) -ForegroundColor $ColorSchema.Warning
    }
}

# Flagged accounts callout
$flagged = $accounts | Where-Object { $_.Flags }
if ($flagged.Count -gt 0) {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host "  FLAGGED ACCOUNTS ($($flagged.Count))" -ForegroundColor $ColorSchema.Warning
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    foreach ($acct in $flagged) {
        Write-Host "  $($acct.Name)" -ForegroundColor $ColorSchema.Warning
        Write-Host "    $($acct.Flags)" -ForegroundColor $ColorSchema.Info
    }
}

# HTML report
Write-Host ""
Write-Host "  [*] Generating HTML report..." -ForegroundColor $ColorSchema.Progress

$reportFilename = "WARD_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath     = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportFilename

try {
    $htmlContent = Build-HtmlReport -Accounts $accounts -MachineName $machineName -ReportTimestamp $reportTimestamp
    [System.IO.File]::WriteAllText($reportPath, $htmlContent, [System.Text.Encoding]::UTF8)
    Write-Host "  [+] Report saved: $reportPath" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Could not save report: $_" -ForegroundColor $ColorSchema.Error
}

# Summary
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  AUDIT SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$enabledCount  = ($accounts | Where-Object {  $_.Enabled } | Measure-Object).Count
$disabledCount = ($accounts | Where-Object { -not $_.Enabled } | Measure-Object).Count
$adminCount    = ($accounts | Where-Object {  $_.IsAdmin } | Measure-Object).Count
$flaggedCount  = ($accounts | Where-Object {  $_.Flags   } | Measure-Object).Count

Write-Host "  Total Accounts : $($accounts.Count)" -ForegroundColor $ColorSchema.Info
Write-Host "  Enabled        : $enabledCount" -ForegroundColor $ColorSchema.Success
Write-Host "  Disabled       : $disabledCount" -ForegroundColor $ColorSchema.Info
Write-Host "  Administrators : $adminCount" -ForegroundColor $ColorSchema.Warning
Write-Host "  Flagged        : $flaggedCount" -ForegroundColor $(if ($flaggedCount -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Success })
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  W.A.R.D. AUDIT COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
