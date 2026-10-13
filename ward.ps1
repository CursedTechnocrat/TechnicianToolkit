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
    Version : 6.1

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

# ===========================
# SHARED HELPERS
# ===========================
#region TK-SHARED
# Generated from shared/TechnicianToolkit.Shared.ps1 by build/Sync-TKShared.ps1. Do not edit
# this block: edit the source and re-run the script. The Pester suite fails on a stale block.

# Forcing UTF-8 keeps the box-drawing banners intact when a console is attached.
# When one is not — a GUI host, a WinRM runspace, a scheduled task — the setter
# throws "The handle is invalid", so this must never be fatal. Every tool runs
# this as the first line of its generated block.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Ok   { param([string]$Msg) Write-Host ("  [+] {0}" -f $Msg) -ForegroundColor Green   }

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

function Write-Info { param([string]$Msg) Write-Host ("      {0}" -f $Msg) -ForegroundColor Gray    }

function Show-TKReportResult {
    <#
    .SYNOPSIS
        Prints a consistent "report saved" line and offers to open the file.
    .DESCRIPTION
        Standardises the end-of-run experience for tools that emit a report.
        Prints the resolved path once, and — unless -Unattended is set — prompts
        to open it in the default handler. Safe to call when the file is missing:
        it prints a warning instead of prompting.
    .PARAMETER Path        Path to the generated report file.
    .PARAMETER Unattended  Suppress the interactive "open now?" prompt.
    .PARAMETER Label       Noun shown in the saved line (default 'Report').
    .EXAMPLE
        Show-TKReportResult -Path $reportPath -Unattended:$Unattended
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Unattended,
        [string]$Label = 'Report'
    )

    if (-not (Test-Path $Path)) {
        Write-Warn "$Label was not created: $Path"
        return
    }

    $full = (Resolve-Path $Path -ErrorAction SilentlyContinue).Path
    if (-not $full) { $full = $Path }

    Write-Ok "$Label saved:"
    Write-Info $full

    if ($Unattended) { return }

    Write-Host ""
    Write-Host -NoNewline ("  Open the {0} now? (Y/N) " -f $Label.ToLower()) -ForegroundColor Cyan
    $answer = Read-Host
    if ($answer -match '^(y|yes)$') {
        try   { Invoke-Item -Path $full -ErrorAction Stop }
        catch { Write-Warn "Could not open the file: $($_.Exception.Message)" }
    }
}

function EscHtml {
    param([string]$s)
    if (-not $s) { return '' }
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Get-TKHtmlCss {
    return @'
<style>
:root {
  --tk-bg:          #0a0e14;
  --tk-surface:     #111820;
  --tk-surface2:    #162030;
  --tk-border:      #1e2d3d;
  --tk-cyan:        #00e5cc;
  --tk-cyan-dim:    rgba(0,229,204,0.12);
  --tk-text:        #c8d4e0;
  --tk-text-dim:    #637587;
  --tk-green:       #3fb950;
  --tk-green-dim:   rgba(63,185,80,0.12);
  --tk-yellow:      #e3b341;
  --tk-yellow-dim:  rgba(227,179,65,0.12);
  --tk-red:         #f85149;
  --tk-red-dim:     rgba(248,81,73,0.12);
  --tk-blue:        #58a6ff;
  --tk-blue-dim:    rgba(88,166,255,0.12);
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--tk-bg);color:var(--tk-text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;font-size:14px;line-height:1.6}

/* nav */
.tk-nav{background:#0d1219;border-bottom:1px solid var(--tk-border);padding:0 32px;height:44px;display:flex;align-items:center;overflow-x:auto;white-space:nowrap;font-family:'Consolas','Courier New',monospace;font-size:11px;letter-spacing:.05em;color:var(--tk-text-dim);gap:0}
.tk-nav a{padding:0 16px;height:44px;display:inline-flex;align-items:center;text-decoration:none;color:var(--tk-text-dim);border-bottom:2px solid transparent;gap:6px}
.tk-nav a:hover{color:var(--tk-text)}
.tk-nav-num{color:var(--tk-cyan)}

/* page header */
.tk-page-header{background:linear-gradient(180deg,#0e1520 0%,var(--tk-bg) 100%);border-bottom:1px solid var(--tk-border);padding:36px 48px 32px}
.tk-report-label{font-family:'Consolas','Courier New',monospace;font-size:11px;letter-spacing:.15em;text-transform:uppercase;color:var(--tk-cyan);margin-bottom:10px}
.tk-page-title{font-size:28px;font-weight:600;color:#e8f0f8;line-height:1.2;margin-bottom:6px}
.tk-page-subtitle{font-size:13px;color:var(--tk-text-dim);margin-bottom:20px}
.tk-meta-bar{display:flex;gap:32px;flex-wrap:wrap;margin-top:18px}
.tk-meta-label{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-text-dim);margin-bottom:3px}
.tk-meta-value{font-size:14px;font-weight:600;color:var(--tk-text)}

/* main */
.tk-main{padding:40px 48px;max-width:1280px}

/* section */
.tk-section{margin-bottom:48px}
.tk-section-tag{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--tk-cyan);margin-bottom:8px}
.tk-section-title{font-size:20px;font-weight:600;color:#e0eaf4;margin-bottom:4px;display:flex;align-items:baseline;gap:10px}
.tk-section-num{font-family:'Consolas','Courier New',monospace;font-size:12px;color:var(--tk-cyan)}
.tk-section-subtitle{font-size:13px;color:var(--tk-text-dim);margin-bottom:16px}
.tk-divider{border:none;border-top:1px solid var(--tk-border);margin:12px 0 20px}

/* card */
.tk-card{background:var(--tk-surface);border:1px solid var(--tk-border);border-radius:8px;padding:20px 24px;margin-bottom:16px}
.tk-card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}
.tk-card-label{font-family:'Consolas','Courier New',monospace;font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-cyan)}

/* summary row */
.tk-summary-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:32px}
.tk-summary-card{background:var(--tk-surface);border:1px solid var(--tk-border);border-radius:8px;padding:18px 24px;min-width:130px;flex:1}
.tk-summary-num{font-size:28px;font-weight:700;color:var(--tk-text);line-height:1;margin-bottom:6px}
.tk-summary-lbl{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--tk-text-dim)}
.tk-summary-card.ok   .tk-summary-num{color:var(--tk-green)}
.tk-summary-card.warn .tk-summary-num{color:var(--tk-yellow)}
.tk-summary-card.err  .tk-summary-num{color:var(--tk-red)}
.tk-summary-card.info .tk-summary-num{color:var(--tk-cyan)}

/* table */
.tk-table-wrap{overflow-x:auto}
table.tk-table{width:100%;border-collapse:collapse;font-size:13px}
table.tk-table th{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-cyan);text-align:left;padding:10px 12px;border-bottom:1px solid var(--tk-border);font-weight:normal;white-space:nowrap}
table.tk-table td{padding:11px 12px;border-bottom:1px solid #162030;color:var(--tk-text);vertical-align:middle}
table.tk-table tr:last-child td{border-bottom:none}
table.tk-table tr:hover td{background:rgba(255,255,255,.02)}

/* badges */
/* The pill shape belongs to every colour class too: most reports write class='tk-badge-ok' alone. */
.tk-badge,.tk-badge-ok,.tk-badge-warn,.tk-badge-err,.tk-badge-info,.tk-badge-blue{display:inline-block;padding:2px 10px;border-radius:20px;font-family:'Consolas','Courier New',monospace;font-size:11px;font-weight:600;letter-spacing:.03em;white-space:nowrap}
.tk-badge-ok   {background:var(--tk-green-dim); color:var(--tk-green); border:1px solid rgba(63,185,80,.25)}
.tk-badge-warn {background:var(--tk-yellow-dim);color:var(--tk-yellow);border:1px solid rgba(227,179,65,.25)}
.tk-badge-err  {background:var(--tk-red-dim);   color:var(--tk-red);   border:1px solid rgba(248,81,73,.25)}
.tk-badge-info {background:var(--tk-cyan-dim);  color:var(--tk-cyan);  border:1px solid rgba(0,229,204,.25)}
.tk-badge-blue {background:var(--tk-blue-dim);  color:var(--tk-blue);  border:1px solid rgba(88,166,255,.25)}

/* info box */
.tk-info-box{background:var(--tk-surface);border-left:3px solid var(--tk-cyan);border-radius:0 6px 6px 0;padding:14px 18px;margin-top:12px;font-size:13px}
.tk-info-label{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-cyan);margin-bottom:4px}

/* progress */
.tk-progress-wrap{background:#162030;border-radius:4px;height:6px;overflow:hidden;width:120px;display:inline-block;vertical-align:middle;margin-left:8px}
.tk-progress-bar{height:100%;border-radius:4px}
.tk-progress-bar.ok  {background:var(--tk-green)}
.tk-progress-bar.warn{background:var(--tk-yellow)}
.tk-progress-bar.err {background:var(--tk-red)}

/* mono / code */
code,.tk-mono{font-family:'Consolas','Courier New',monospace;font-size:12px;background:var(--tk-surface2);padding:1px 5px;border-radius:3px}

/* footer */
.tk-footer{border-top:1px solid var(--tk-border);padding:20px 48px;font-family:'Consolas','Courier New',monospace;font-size:11px;color:var(--tk-text-dim);letter-spacing:.05em;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px}

/* responsive */
img,svg,video{max-width:100%;height:auto}
.tk-main,.tk-card,.tk-info-box,.tk-summary-card{max-width:100%}
.tk-card{overflow-x:auto}
table.tk-table{max-width:100%}
table.tk-table td,table.tk-table th{overflow-wrap:anywhere;word-break:break-word}
.tk-info-box,.tk-meta-value,.tk-page-title,.tk-page-subtitle,.tk-section-title,.tk-section-subtitle{overflow-wrap:anywhere}

@media (max-width:900px){
  .tk-page-header{padding:24px 20px 20px}
  .tk-page-title{font-size:22px}
  .tk-meta-bar{gap:18px}
  .tk-nav{padding:0 16px}
  .tk-nav a{padding:0 12px}
  .tk-main{padding:24px 20px}
  .tk-section{margin-bottom:32px}
  .tk-card{padding:16px 18px}
  .tk-footer{padding:16px 20px}
  table.tk-table th,table.tk-table td{padding:9px 8px}
}
@media (max-width:560px){
  body{font-size:13px}
  .tk-page-header{padding:18px 14px 16px}
  .tk-page-title{font-size:19px}
  .tk-page-subtitle{font-size:12px}
  .tk-meta-bar{gap:14px}
  .tk-nav{padding:0 10px;font-size:10px}
  .tk-nav a{padding:0 10px}
  .tk-main{padding:18px 14px}
  .tk-section-title{font-size:17px}
  .tk-card{padding:14px}
  .tk-summary-row{gap:10px}
  .tk-summary-card{min-width:0;flex:1 1 calc(50% - 5px);padding:14px 16px}
  .tk-summary-num{font-size:22px}
  .tk-card-header{flex-direction:column;align-items:flex-start;gap:8px}
  .tk-footer{padding:14px;flex-direction:column;align-items:flex-start;gap:4px}
  table.tk-table{font-size:12px}
  table.tk-table th,table.tk-table td{padding:8px 6px}
  .tk-progress-wrap{width:80px}
}
</style>
'@
}

function Get-TKHtmlHead {
    <#
    .SYNOPSIS
        Returns the opening HTML, head, and page-header markup for a toolkit report.
    .PARAMETER Title      Report title shown as the page heading.
    .PARAMETER ScriptName Acronym label (e.g. 'A.U.S.P.E.X.') used in the report-label tag.
    .PARAMETER Subtitle   Optional subtitle / machine name line beneath the title.
    .PARAMETER MetaItems  Ordered hashtable of label->value pairs shown in the metadata bar.
    .PARAMETER NavItems   Array of section-label strings for the sticky nav bar.
    #>
    param(
        [string]   $Title      = 'Technician Toolkit Report',
        [string]   $ScriptName = 'T.K.',
        [string]   $Subtitle   = '',
        [hashtable]$MetaItems  = @{},
        [string[]] $NavItems   = @()
    )

    $css = Get-TKHtmlCss

    $metaHtml = ''
    if ($MetaItems.Count -gt 0) {
        $parts = foreach ($k in $MetaItems.Keys) {
            "<div class='tk-meta-item'><div class='tk-meta-label'>$(EscHtml $k)</div><div class='tk-meta-value'>$(EscHtml $MetaItems[$k])</div></div>"
        }
        $metaHtml = "<div class='tk-meta-bar'>$($parts -join '')</div>"
    }

    $subtitleHtml = if ($Subtitle) { "<div class='tk-page-subtitle'>$(EscHtml $Subtitle)</div>" } else { '' }

    $navHtml = ''
    if ($NavItems.Count -gt 0) {
        $links = for ($i = 0; $i -lt $NavItems.Count; $i++) {
            $n = '{0:D2}' -f ($i + 1)
            "<a href='#s$n'><span class='tk-nav-num'>$n</span> &middot; $(EscHtml $NavItems[$i])</a>"
        }
        $navHtml = "<nav class='tk-nav'>$($links -join '')</nav>"
    }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>$Title</title>
$css
</head>
<body>
<div class="tk-page-header">
  <div class="tk-report-label">$(EscHtml $ScriptName) REPORT</div>
  <div class="tk-page-title">$(EscHtml $Title)</div>
  $subtitleHtml
  $metaHtml
</div>
$navHtml
<div class="tk-main">
"@
}

function Get-TKHtmlFoot {
    <#
    .SYNOPSIS
        Returns the closing HTML markup for a toolkit report.
    .PARAMETER ScriptName  Shown in footer right (e.g. 'A.U.S.P.E.X. v1.0').
    #>
    param([string]$ScriptName = 'TechnicianToolkit')
    $ts       = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $hostName = $env:COMPUTERNAME
    return @"
</div>
<div class="tk-footer">
  <span>Generated $ts on $hostName</span>
  <span>$(EscHtml $ScriptName)</span>
</div>
</body>
</html>
"@
}

function Get-TKConfig {
    <#
    .SYNOPSIS
        Returns the toolkit configuration as a PSCustomObject.
        Reads config.json from the tool's own directory; missing keys are filled
        with empty-string defaults so callers never receive null.
    #>
    $configPath = Join-Path $PSScriptRoot 'config.json'

    $defaults = [PSCustomObject]@{
        OrgName       = ''
        LogDirectory  = ''
        Archive       = [PSCustomObject]@{ DefaultDestination   = '' }
        Revenant      = [PSCustomObject]@{ DefaultDestination   = '' }
        Covenant      = [PSCustomObject]@{ DefaultTimezone = ''; DefaultLocalAdminUser = '' }
        # DirectDownloads is an array of { Name; Url; Args; Sha256 } — installers
        # CONJURE fetches by URL because they have no winget/Chocolatey package.
        Conjure       = [PSCustomObject]@{ DirectDownloads = @() }
    }

    if (-not (Test-Path $configPath)) { return $defaults }

    try {
        $raw = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json

        # Legacy migration: v2.x used a 'Phantom' section; v3.0 renamed it to 'Revenant'.
        # Copy any populated 'Phantom' value forward so upgrading users don't lose their default destination.
        if ($raw.PSObject.Properties['Phantom'] -and
            $raw.Phantom -is [PSCustomObject] -and
            -not [string]::IsNullOrWhiteSpace($raw.Phantom.DefaultDestination)) {
            if ($raw.PSObject.Properties['Revenant'] -isnot [PSCustomObject] -or
                [string]::IsNullOrWhiteSpace($raw.Revenant.DefaultDestination)) {
                if (-not $raw.PSObject.Properties['Revenant']) {
                    $raw | Add-Member -NotePropertyName 'Revenant' -NotePropertyValue ([PSCustomObject]@{ DefaultDestination = '' }) -Force
                }
                $raw.Revenant.DefaultDestination = $raw.Phantom.DefaultDestination
            }
        }

        # Ensure top-level keys exist
        foreach ($key in ($defaults | Get-Member -MemberType NoteProperty).Name) {
            if ($null -eq $raw.$key) {
                $raw | Add-Member -NotePropertyName $key -NotePropertyValue $defaults.$key -Force
            }
        }
        # Ensure nested keys exist
        foreach ($section in @('Archive','Revenant','Covenant','Conjure')) {
            if ($raw.$section -isnot [PSCustomObject]) {
                $raw | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            } else {
                foreach ($key in ($defaults.$section | Get-Member -MemberType NoteProperty).Name) {
                    if ($null -eq $raw.$section.$key) {
                        # Carry the default's own type: filling an absent array key
                        # with '' would hand the caller a string to iterate.
                        $fallback = if ($defaults.$section.$key -is [Array]) { @() } else { '' }
                        $raw.$section | Add-Member -NotePropertyName $key -NotePropertyValue $fallback -Force
                    }
                }
            }
        }
        return $raw
    }
    catch { return $defaults }
}

function Start-TKTranscript {
    <#
    .SYNOPSIS
        Starts a timestamped PowerShell transcript in the configured log directory.
        Call once near the top of a script when -Transcript is active.
    #>
    param([Parameter(Mandatory)][string]$LogRoot)
    $path = Join-Path $LogRoot "TK_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    try {
        Start-Transcript -Path $path -Append -ErrorAction Stop | Out-Null
        Write-Host "  [*] Transcript: $path" -ForegroundColor Gray
    }
    catch {
        Write-Warn "Could not start transcript: $_"
    }
}

function Stop-TKTranscript {
    <#
    .SYNOPSIS
        Stops an active PowerShell transcript started by Start-TKTranscript.
    #>
    try { Stop-Transcript -ErrorAction Stop | Out-Null } catch {}
}

function Resolve-LogDirectory {
    <#
    .SYNOPSIS
        Returns the configured LogDirectory, or the supplied fallback path if
        LogDirectory is not set. Creates the directory if it does not exist.
    #>
    param([Parameter(Mandatory)][string]$FallbackPath)

    $cfg = Get-TKConfig
    if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory)) {
        if (-not (Test-Path $cfg.LogDirectory)) {
            $null = New-Item -ItemType Directory -Path $cfg.LogDirectory -Force -ErrorAction SilentlyContinue
        }
        return $cfg.LogDirectory
    }
    return $FallbackPath
}

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
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

    # Mark the account only when a policy actually rotates it. With no policy (or a Disabled one) the
    # built-in Administrator is just the account LAPS would manage, and a LAPS badge would imply
    # protection that is not there.
    if ($managed -and $policy.Mode -notin 'Not configured', 'Disabled') {
        $managed | Add-Member -NotePropertyName LapsManaged -NotePropertyValue $true -Force
    }

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

  <div class="tk-section" id="s01">
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

  <div class="tk-section" id="s02">
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

"@ + (Get-TKHtmlFoot -ScriptName 'W.A.R.D. v6.1')

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
