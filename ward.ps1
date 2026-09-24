# ward.ps1 - W.A.R.D. — Watches Accounts, Reviews Roles & Detects anomalies
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
    W.A.R.D. — Watches Accounts, Reviews Roles & Detects anomalies
    User Account & Security Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits all local user accounts on the machine. Reports account status,
    last logon time, password configuration, and group memberships. Flags
    potentially risky accounts and exports a dark-themed HTML report to the
    script directory.

    Also reports whether the local administrator password is managed by LAPS:
    which implementation (Windows LAPS or legacy Microsoft LAPS) and policy
    source applies, where the password is backed up, whether the device is
    joined to that directory, which account is managed, and whether its
    password is actually being rotated.

.USAGE
    PS C:\> .\ward.ps1                    # Must be run as Administrator
    PS C:\> .\ward.ps1 -Unattended        # Silent mode — no prompts, no banner

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

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
    $accounts   = [System.Collections.Generic.List[object]]::new()

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

        $accounts.Add([PSCustomObject]@{
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
        })
    }

    return $accounts
}

# ─────────────────────────────────────────────────────────────────────────────
# LAPS — LOCAL ADMINISTRATOR PASSWORD SOLUTION
# ─────────────────────────────────────────────────────────────────────────────

# Windows LAPS policy roots, highest precedence first. Windows LAPS reads the
# first root that sets BackupDirectory and ignores the rest, so a stale GPO
# under an Intune policy is inert -- which is why the report names the source.
$LapsPolicySources = [ordered]@{
    'CSP (Intune / MDM)'  = 'HKLM:\SOFTWARE\Microsoft\Policies\LAPS'
    'Group Policy'        = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS'
    'Local configuration' = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\LAPS\Config'
}
$LegacyLapsPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft Services\AdmPwd'
$LegacyLapsCsePath   = Join-Path $env:ProgramFiles 'LAPS\CSE\AdmPwd.dll'

# BackupDirectory -> where Windows LAPS escrows the password.
$LapsBackupTargets = @{
    0 = 'Disabled'
    1 = 'Entra ID'
    2 = 'Active Directory'
}

# Days of grace past PasswordAgeDays before a rotation counts as overdue. The
# LAPS processing cycle runs hourly, but a machine that was off over the
# expiry date rotates on its next cycle, not on the day.
$LapsRotationGraceDays = 3

function Resolve-LapsPolicy {
    # Pure: decides which LAPS implementation governs this machine from the raw
    # policy values, so the precedence rules can be tested without a registry.
    #   $Sources      ordered name -> hashtable of values (or $null if absent)
    #   $LegacyPolicy hashtable of AdmPwd values (or $null)
    param(
        [System.Collections.IDictionary]$Sources,
        [hashtable]$LegacyPolicy,
        [bool]$LegacyCseInstalled
    )

    foreach ($name in $Sources.Keys) {
        $v = $Sources[$name]
        if ($null -eq $v -or $null -eq $v['BackupDirectory']) { continue }
        $dir = [int]$v['BackupDirectory']
        return [PSCustomObject]@{
            Mode            = if ($dir -eq 0) { 'Disabled' } else { 'Windows LAPS' }
            Source          = $name
            BackupDirectory = $dir
            AccountName     = "$($v['AdministratorAccountName'])"
            AutoAccount     = ($v['AutomaticAccountManagementEnabled'] -eq 1)
            PasswordAgeDays = if ($v['PasswordAgeDays']) { [int]$v['PasswordAgeDays'] } else { 30 }
        }
    }

    if ($LegacyPolicy -and $LegacyPolicy['AdmPwdEnabled'] -eq 1) {
        # With the legacy CSE installed, legacy LAPS owns the account. Without it,
        # Windows LAPS honours the legacy policy in emulation mode (AD only).
        return [PSCustomObject]@{
            Mode            = if ($LegacyCseInstalled) { 'Legacy Microsoft LAPS' } else { 'Windows LAPS (legacy emulation)' }
            Source          = 'Legacy LAPS policy (AdmPwd)'
            BackupDirectory = 2
            AccountName     = "$($LegacyPolicy['AdminAccountName'])"
            AutoAccount     = $false
            PasswordAgeDays = if ($LegacyPolicy['PasswordAgeDays']) { [int]$LegacyPolicy['PasswordAgeDays'] } else { 30 }
        }
    }

    return [PSCustomObject]@{
        Mode = 'Not configured'; Source = ''; BackupDirectory = $null; AccountName = ''
        AutoAccount = $false; PasswordAgeDays = $null
    }
}

function Get-RegistryValueTable {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    $item  = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue
    if (-not $item) { return $null }
    $table = @{}
    foreach ($p in $item.PSObject.Properties) {
        if ($p.Name -notlike 'PS*') { $table[$p.Name] = $p.Value }
    }
    return $table
}

function Get-LapsStatus {
    param([array]$Accounts)

    $issues = [System.Collections.Generic.List[object]]::new()
    function _issue { param([string]$Severity, [string]$Text) $issues.Add([PSCustomObject]@{ Severity = $Severity; Text = $Text }) }

    $sources = [ordered]@{}
    foreach ($name in $LapsPolicySources.Keys) { $sources[$name] = Get-RegistryValueTable -Path $LapsPolicySources[$name] }
    $legacy = Get-RegistryValueTable -Path $LegacyLapsPolicyKey
    $policy = Resolve-LapsPolicy -Sources $sources -LegacyPolicy $legacy -LegacyCseInstalled (Test-Path $LegacyLapsCsePath)

    # Join state decides which backup targets can work at all.
    $domainJoined = $false
    try { $domainJoined = [bool](Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { $domainJoined = $false }
    $entraJoined = $false
    try { $entraJoined = [bool]((dsregcmd /status 2>$null) -match 'AzureAdJoined\s*:\s*YES') } catch { $entraJoined = $false }
    $joinLabel = @(@(if ($domainJoined) { 'Active Directory' }) + @(if ($entraJoined) { 'Entra ID' }))
    $joinLabel = if ($joinLabel.Count -gt 0) { $joinLabel -join ' + ' } else { 'Workgroup (not joined)' }

    # The managed account: named by policy, else the built-in Administrator (RID 500).
    $builtin = $null
    try { $builtin = Get-LocalUser -ErrorAction Stop | Where-Object { "$($_.SID)" -like 'S-1-5-*-500' } | Select-Object -First 1 } catch { $builtin = $null }
    $accountName = if ($policy.AccountName) { $policy.AccountName } elseif ($builtin) { $builtin.Name } else { '' }
    $managed = if ($accountName) { $Accounts | Where-Object { $_.Name -eq $accountName } | Select-Object -First 1 } else { $null }
    $localUser = if ($accountName) { Get-LocalUser -Name $accountName -ErrorAction SilentlyContinue } else { $null }

    $lastRotation = if ($localUser -and $localUser.PasswordLastSet) { $localUser.PasswordLastSet } else { $null }
    $rotationAge  = if ($lastRotation) { [int]((Get-Date) - $lastRotation).TotalDays } else { $null }

    # Recent errors from the Windows LAPS operational log, whatever their ID.
    $lastError = $null
    try {
        $lastError = Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-LAPS/Operational'; Level = 2; StartTime = (Get-Date).AddDays(-7) } -MaxEvents 1 -ErrorAction Stop
    } catch {
        $lastError = $null
    }

    $enabledAdmins = @($Accounts | Where-Object { $_.Enabled -and $_.IsAdmin })

    switch -Wildcard ($policy.Mode) {
        'Not configured' {
            if (-not $domainJoined -and -not $entraJoined) {
                _issue 'Info' 'No LAPS policy. This machine is not joined to Active Directory or Entra ID, and Windows LAPS needs one of them to store the password.'
            } elseif ($enabledAdmins.Count -gt 0) {
                _issue 'Warning' "No LAPS policy: the password on $($enabledAdmins.Count) enabled local administrator account(s) is not rotated or escrowed. A password shared across machines lets one compromise spread to all of them."
            } else {
                _issue 'Info' 'No LAPS policy, but no local administrator account is enabled.'
            }
        }
        'Disabled' {
            _issue 'Warning' "A LAPS policy exists ($($policy.Source)) but BackupDirectory is 0 -- LAPS is switched off."
        }
        'Legacy Microsoft LAPS' {
            _issue 'Info' 'Legacy Microsoft LAPS (AdmPwd CSE) manages this machine. It is deprecated; migrate the policy to Windows LAPS.'
        }
    }

    if ($policy.BackupDirectory -eq 2 -and -not $domainJoined) {
        _issue 'Error' 'LAPS backs up to Active Directory, but the machine is not domain-joined -- the password cannot be escrowed.'
    }
    if ($policy.BackupDirectory -eq 1 -and -not $entraJoined) {
        _issue 'Error' 'LAPS backs up to Entra ID, but the machine is not Entra-joined -- the password cannot be escrowed.'
    }

    if ($policy.BackupDirectory -in 1, 2) {
        if (-not $accountName) {
            _issue 'Warning' 'Could not determine which local account LAPS manages.'
        } elseif (-not $localUser -and -not $policy.AutoAccount) {
            _issue 'Error' "LAPS is configured to manage '$accountName', but no local account has that name. Nothing is being rotated."
        } else {
            if ($localUser -and -not $localUser.Enabled) {
                _issue 'Info' "The managed account '$accountName' is disabled. LAPS still rotates it; enable it only when it is needed."
            }
            if ($null -ne $rotationAge -and $policy.PasswordAgeDays -and $rotationAge -gt ($policy.PasswordAgeDays + $LapsRotationGraceDays)) {
                _issue 'Warning' "The managed account's password was last set $rotationAge day(s) ago, past the $($policy.PasswordAgeDays)-day policy -- rotation looks stalled."
            }
        }
        if ($lastError) {
            _issue 'Warning' ("Windows LAPS logged an error on {0} (event {1}): {2}" -f $lastError.TimeCreated.ToString('yyyy-MM-dd HH:mm'), $lastError.Id, (($lastError.Message -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1))
        }
    }

    if ($managed) { $managed | Add-Member -NotePropertyName LapsManaged -NotePropertyValue $true -Force }

    return [PSCustomObject]@{
        Mode            = $policy.Mode
        Source          = $policy.Source
        BackupTarget    = if ($null -ne $policy.BackupDirectory -and $LapsBackupTargets.ContainsKey($policy.BackupDirectory)) { $LapsBackupTargets[$policy.BackupDirectory] } else { '-' }
        JoinState       = $joinLabel
        AccountName     = if ($accountName) { $accountName } else { '-' }
        AccountExists   = [bool]$localUser
        PasswordAgeDays = $policy.PasswordAgeDays
        LastRotation    = if ($lastRotation) { $lastRotation.ToString('yyyy-MM-dd HH:mm') } else { '-' }
        RotationAgeDays = $rotationAge
        Issues          = @($issues)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Accounts, [object]$Laps, [string]$MachineName, [string]$ReportTimestamp)

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
        if ($acct.LapsManaged) { $adminBadge += " <span class='tk-badge-ok'>LAPS</span>" }
        $flagCell = if ($acct.Flags) {
            "<span class='tk-badge-warn'>$(EscHtml($acct.Flags))</span>"
        } else { "" }

        $rows += @"
            <tr>
                <td><strong>$(EscHtml($acct.Name))</strong></td>
                <td>$(EscHtml($acct.FullName))</td>
                <td>$enabledBadge</td>
                <td>$adminBadge</td>
                <td>$(EscHtml($acct.LastLogon))</td>
                <td>$(EscHtml($acct.PasswordLastSet))</td>
                <td>$(EscHtml($acct.PasswordExpires))</td>
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
        'LAPS'      = $Laps.Mode
    }

    $tkNavItems = @('Local User Accounts', 'Local Administrator Password (LAPS)')

    $lapsWorst = @($Laps.Issues | ForEach-Object { $_.Severity })
    $lapsClass = if ($lapsWorst -contains 'Error') { 'err' }
                 elseif ($lapsWorst -contains 'Warning') { 'warn' }
                 elseif ($Laps.Mode -like 'Windows LAPS*' -or $Laps.Mode -eq 'Legacy Microsoft LAPS') { 'ok' }
                 else { 'info' }

    $lapsIssueRows = ''
    if ($Laps.Issues.Count -eq 0) {
        $lapsIssueRows = "<tr><td><span class='tk-badge-ok'>OK</span></td><td>LAPS is configured and rotating the managed account.</td></tr>"
    } else {
        foreach ($i in $Laps.Issues) {
            $cls = switch ($i.Severity) { 'Error' { 'err' } 'Warning' { 'warn' } default { 'info' } }
            $lapsIssueRows += "<tr><td><span class='tk-badge-$cls'>$(EscHtml $i.Severity)</span></td><td>$(EscHtml $i.Text)</td></tr>"
        }
    }

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
  <div class="tk-summary-card $lapsClass">
    <div class="tk-summary-num">$(EscHtml $Laps.BackupTarget)</div>
    <div class="tk-summary-lbl">LAPS Backup</div>
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

  <div class="tk-section" id="local-administrator-password-laps">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 2</span>
        <h2 class="tk-section-title">Local Administrator Password (LAPS)</h2>
      </div>
      <div style="padding:20px;">
        <div class="tk-info-box">
          <span class="tk-info-label">Mode</span> $(EscHtml $Laps.Mode)<br/>
          <span class="tk-info-label">Policy source</span> $(if ($Laps.Source) { EscHtml $Laps.Source } else { '-' })<br/>
          <span class="tk-info-label">Password backup</span> $(EscHtml $Laps.BackupTarget)<br/>
          <span class="tk-info-label">Device joined to</span> $(EscHtml $Laps.JoinState)<br/>
          <span class="tk-info-label">Managed account</span> $(EscHtml $Laps.AccountName)$(if ($Laps.AccountName -ne '-' -and -not $Laps.AccountExists) { ' (not found)' } else { '' })<br/>
          <span class="tk-info-label">Password age policy</span> $(if ($Laps.PasswordAgeDays) { "$($Laps.PasswordAgeDays) day(s)" } else { '-' })<br/>
          <span class="tk-info-label">Password last set</span> $(EscHtml $Laps.LastRotation)$(if ($null -ne $Laps.RotationAgeDays) { " ($($Laps.RotationAgeDays) day(s) ago)" } else { '' })
        </div>
        <table class="tk-table" style="margin-top:18px;">
          <thead><tr><th>Status</th><th>Finding</th></tr></thead>
          <tbody>
            $lapsIssueRows
          </tbody>
        </table>
      </div>
    </div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'W.A.R.D. v3.6')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-WardBanner }

$machineName      = $env:COMPUTERNAME
$reportTimestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  COLLECTING ACCOUNT DATA" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

Write-Host "  [*] Resolving Administrators group members..." -ForegroundColor $ColorSchema.Progress
$adminNames = Get-AdminMembers

Write-Host "  [*] Enumerating local user accounts..." -ForegroundColor $ColorSchema.Progress
$accounts = Get-AccountData -AdminNames $adminNames

Write-Host "  [+] Found $($accounts.Count) local user account(s)." -ForegroundColor $ColorSchema.Success
Write-Host ""

Write-Host "  [*] Checking LAPS policy and the managed account..." -ForegroundColor $ColorSchema.Progress
$laps = Get-LapsStatus -Accounts $accounts
Write-Host ""

# Console summary
Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  ACCOUNT OVERVIEW" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
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
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host "  FLAGGED ACCOUNTS ($($flagged.Count))" -ForegroundColor $ColorSchema.Warning
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    foreach ($acct in $flagged) {
        Write-Host "  $($acct.Name)" -ForegroundColor $ColorSchema.Warning
        Write-Host "    $($acct.Flags)" -ForegroundColor $ColorSchema.Info
    }
}

# LAPS
Write-Host ""
Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  LOCAL ADMINISTRATOR PASSWORD (LAPS)" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Mode             : $($laps.Mode)$(if ($laps.Source) { " ($($laps.Source))" })" -ForegroundColor $ColorSchema.Info
Write-Host "  Password backup  : $($laps.BackupTarget)" -ForegroundColor $ColorSchema.Info
Write-Host "  Joined to        : $($laps.JoinState)" -ForegroundColor $ColorSchema.Info
Write-Host "  Managed account  : $($laps.AccountName)  (password last set $($laps.LastRotation))" -ForegroundColor $ColorSchema.Info
foreach ($i in $laps.Issues) {
    $color = switch ($i.Severity) { 'Error' { $ColorSchema.Error } 'Warning' { $ColorSchema.Warning } default { $ColorSchema.Info } }
    Write-Host "  [$($i.Severity.ToUpper())] $($i.Text)" -ForegroundColor $color
}
if ($laps.Issues.Count -eq 0) {
    Write-Host "  [+] LAPS is configured and rotating the managed account." -ForegroundColor $ColorSchema.Success
}

# HTML report
Write-Host ""
Write-Host "  [*] Generating HTML report..." -ForegroundColor $ColorSchema.Progress

$reportFilename = "WARD_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath     = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportFilename

try {
    $htmlContent = Build-HtmlReport -Accounts $accounts -Laps $laps -MachineName $machineName -ReportTimestamp $reportTimestamp
    [System.IO.File]::WriteAllText($reportPath, $htmlContent, [System.Text.Encoding]::UTF8)
}
catch {
    Write-Host "  [-] Could not save report: $_" -ForegroundColor $ColorSchema.Error
}

# Summary
Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  AUDIT SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
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
Write-Host "  LAPS           : $($laps.Mode)" -ForegroundColor $(if (@($laps.Issues | Where-Object { $_.Severity -ne 'Info' }).Count -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Success })
Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  W.A.R.D. AUDIT COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

Show-TKReportResult -Path $reportPath -Unattended:$Unattended

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
