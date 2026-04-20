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
            "<span class='badge badge-ok'>Enabled</span>"
        } else {
            "<span class='badge badge-warn'>Disabled</span>"
        }
        $adminBadge = if ($acct.IsAdmin) {
            "<span class='badge badge-crit'>Admin</span>"
        } else {
            "<span class='badge badge-neutral'>Standard</span>"
        }
        $flagCell = if ($acct.Flags) {
            "<span class='flag'>$(HtmlEncode($acct.Flags))</span>"
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

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>W.A.R.D. Account Audit — $MachineName</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI', Consolas, monospace; font-size: 14px; padding: 24px; }
  h1 { color: #00d4ff; font-size: 22px; margin-bottom: 4px; }
  .subtitle { color: #888; font-size: 13px; margin-bottom: 24px; }
  .summary { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 28px; }
  .card { background: #16213e; border: 1px solid #0f3460; border-radius: 8px; padding: 16px 24px; min-width: 120px; text-align: center; }
  .card .val { font-size: 28px; font-weight: bold; color: #00d4ff; }
  .card .lbl { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
  .card.warn .val { color: #f39c12; }
  .card.crit .val { color: #e74c3c; }
  .card.ok   .val { color: #2ecc71; }
  table { width: 100%; border-collapse: collapse; margin-top: 8px; }
  th { background: #0f3460; color: #00d4ff; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
  td { padding: 9px 12px; border-bottom: 1px solid #1e2d4d; vertical-align: top; }
  tr:hover td { background: #1e2d4d; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; }
  .badge-ok      { background: #1a4a2e; color: #2ecc71; }
  .badge-warn    { background: #4a3a10; color: #f39c12; }
  .badge-crit    { background: #4a1a1a; color: #e74c3c; }
  .badge-neutral { background: #2a2a3e; color: #aaa; }
  .flag { color: #f39c12; font-size: 12px; }
  .section-title { color: #00d4ff; font-size: 15px; margin: 28px 0 10px; border-bottom: 1px solid #0f3460; padding-bottom: 6px; }
  .footer { margin-top: 32px; color: #555; font-size: 11px; }
</style>
</head>
<body>
<h1>W.A.R.D. — Account Audit Report</h1>
<div class="subtitle">$(if (-not [string]::IsNullOrWhiteSpace((Get-TKConfig).OrgName)) { "$(EscHtml (Get-TKConfig).OrgName) &nbsp;|&nbsp; " })Machine: <strong>$MachineName</strong> &nbsp;|&nbsp; Generated: $ReportTimestamp</div>

<div class="summary">
  <div class="card"><div class="val">$totalAccounts</div><div class="lbl">Total Accounts</div></div>
  <div class="card ok"><div class="val">$enabledCount</div><div class="lbl">Enabled</div></div>
  <div class="card"><div class="val">$disabledCount</div><div class="lbl">Disabled</div></div>
  <div class="card warn"><div class="val">$adminCount</div><div class="lbl">Administrators</div></div>
  <div class="card crit"><div class="val">$flaggedCount</div><div class="lbl">Flagged</div></div>
</div>

<div class="section-title">Local User Accounts</div>
<table>
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

<div class="footer">
  Generated by W.A.R.D. — Technician Toolkit &nbsp;|&nbsp; Stale threshold: 90 days without logon
</div>
</body>
</html>
"@

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
