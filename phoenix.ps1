# phoenix.ps1 - P.H.O.E.N.I.X. — Pinpoints Hosts' Outdated & Expiring Notarised Identity X.509s
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
    P.H.O.E.N.I.X. — Pinpoints Hosts' Outdated & Expiring Notarised Identity X.509s
    Certificate Health Monitor for PowerShell 5.1+

.DESCRIPTION
    Audits the local Windows certificate stores (Personal, CA, Trusted Root) for
    expired and expiring-soon certificates. Optionally checks SSL/TLS certificate
    expiry on remote hosts by connecting and reading their presented certificate.
    Exports a dark-themed HTML report with color-coded expiry indicators.

.USAGE
    PS C:\> .\phoenix.ps1                                          # Interactive menu
    PS C:\> .\phoenix.ps1 -Unattended                              # Full audit (local + HTML report)
    PS C:\> .\phoenix.ps1 -Unattended -Targets "srv1.contoso.com,srv2.contoso.com:8443"

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [string]$Targets = '',
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

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ("  " + ("-" * 62)) -ForegroundColor Cyan
    Write-Host ""
}

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

$C = @{
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

function Show-PhoenixBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██╗  ██╗ ██████╗ ███████╗███╗   ██╗██╗██╗  ██╗
  ██╔══██╗██║  ██║██╔═══██╗██╔════╝████╗  ██║██║╚██╗██╔╝
  ██████╔╝███████║██║   ██║█████╗  ██╔██╗ ██║██║ ╚███╔╝
  ██╔═══╝ ██╔══██║██║   ██║██╔══╝  ██║╚██╗██║██║ ██╔██╗
  ██║     ██║  ██║╚██████╔╝███████╗██║ ╚████║██║██╔╝ ██╗
  ╚═╝     ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═══╝╚═╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    P.H.O.E.N.I.X. — Pinpoints Hosts' Outdated & Expiring Notarised Identity X.509s" -ForegroundColor Cyan
    Write-Host "    Certificate Health & SSL Expiry Monitor" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# CORE DATA FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Get-LocalCertHealth {
    param([string[]]$StoreNames = @('My', 'CA', 'Root', 'TrustedPublisher'))
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($storeName in $StoreNames) {
        try {
            $store = [Security.Cryptography.X509Certificates.X509Store]::new($storeName, 'LocalMachine')
            $store.Open('ReadOnly')
            foreach ($cert in $store.Certificates) {
                $daysLeft = [int]($cert.NotAfter - (Get-Date)).TotalDays
                $status   = if ($daysLeft -lt 0)   { 'Expired'  }
                            elseif ($daysLeft -le 30)  { 'Critical' }
                            elseif ($daysLeft -le 90)  { 'Warning'  }
                            else                       { 'Healthy'  }
                $results.Add([PSCustomObject]@{
                    Store      = $storeName
                    Subject    = $cert.Subject
                    Issuer     = $cert.Issuer
                    Thumbprint = $cert.Thumbprint.Substring(0, [Math]::Min(16, $cert.Thumbprint.Length)) + '...'
                    Expiry     = $cert.NotAfter
                    DaysLeft   = $daysLeft
                    Status     = $status
                })
            }
            $store.Close()
        } catch {}
    }
    return $results
}

function Get-SslCertExpiry {
    param([string]$Hostname, [int]$Port = 443, [int]$TimeoutMs = 6000)

    $result = [PSCustomObject]@{
        Hostname = $Hostname
        Port     = $Port
        Subject  = 'N/A'
        Issuer   = 'N/A'
        Expiry   = $null
        DaysLeft = $null
        Status   = 'Unknown'
        Error    = $null
    }

    try {
        $tcp = New-Object Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $wait) { throw "Connection timed out" }
        $tcp.EndConnect($connect)
        $ssl = New-Object Net.Security.SslStream($tcp.GetStream(), $false, { $true })
        $ssl.AuthenticateAsClient($Hostname)
        $cert  = $ssl.RemoteCertificate
        $cert2 = New-Object Security.Cryptography.X509Certificates.X509Certificate2($cert)
        $result.Subject  = $cert2.Subject
        $result.Issuer   = $cert2.Issuer
        $result.Expiry   = $cert2.NotAfter
        $result.DaysLeft = [int]($cert2.NotAfter - (Get-Date)).TotalDays
        $result.Status   = if ($result.DaysLeft -lt 0)   { 'Expired'  }
                           elseif ($result.DaysLeft -le 30)  { 'Critical' }
                           elseif ($result.DaysLeft -le 90)  { 'Warning'  }
                           else                              { 'Healthy'  }
        try { $ssl.Dispose() } catch {}
        try { $tcp.Dispose() } catch {}
    } catch {
        $result.Status = 'Error'
        $result.Error  = $_.Exception.Message
    }

    return $result
}

function ConvertTo-TargetList {
    param([string]$TargetString)

    $entries = [System.Collections.Generic.List[object]]::new()

    if ([string]::IsNullOrWhiteSpace($TargetString)) {
        return $entries
    }

    # If it looks like a file path and the file exists, read it line by line
    if (Test-Path $TargetString -PathType Leaf) {
        $lines = Get-Content $TargetString -ErrorAction SilentlyContinue |
                 Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.TrimStart().StartsWith('#') }
    } else {
        $lines = $TargetString -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }

    foreach ($line in $lines) {
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Check for hostname:port format
        if ($line -match '^(.+):(\d+)$') {
            $entries.Add([PSCustomObject]@{
                Hostname = $Matches[1].Trim()
                Port     = [int]$Matches[2]
            })
        } else {
            $entries.Add([PSCustomObject]@{
                Hostname = $line
                Port     = 443
            })
        }
    }

    return $entries
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE DISPLAY FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Show-LocalCertSummary {
    param([array]$Certs)

    Write-Section "LOCAL CERTIFICATE STORE AUDIT"

    if (-not $Certs -or $Certs.Count -eq 0) {
        Write-Host "  No certificates found." -ForegroundColor $C.Warning
        return
    }

    $expired  = @($Certs | Where-Object { $_.Status -eq 'Expired'  })
    $critical = @($Certs | Where-Object { $_.Status -eq 'Critical' })
    $warning  = @($Certs | Where-Object { $_.Status -eq 'Warning'  })
    $healthy  = @($Certs | Where-Object { $_.Status -eq 'Healthy'  })

    # Summary counts
    Write-Host ("  Total: {0}   Expired: {1}   Critical: {2}   Warning: {3}   Healthy: {4}" -f `
        $Certs.Count, $expired.Count, $critical.Count, $warning.Count, $healthy.Count) -ForegroundColor $C.Info
    Write-Host ""

    $nonHealthy = @($Certs | Where-Object { $_.Status -ne 'Healthy' } | Sort-Object DaysLeft)

    if ($nonHealthy.Count -eq 0) {
        Write-Host "  [+] All certificates are healthy." -ForegroundColor $C.Success
        return
    }

    # Table header
    $col = @(10, 52, 22, 10, 10)
    Write-Host ("  {0,-$($col[0])} {1,-$($col[1])} {2,-$($col[2])} {3,-$($col[3])} {4}" -f `
        'Store', 'Subject', 'Expiry', 'Days Left', 'Status') -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 110)) -ForegroundColor $C.Header

    foreach ($cert in $nonHealthy) {
        $rowColor = switch ($cert.Status) {
            'Expired'  { $C.Error   }
            'Critical' { $C.Error   }
            'Warning'  { $C.Warning }
            default    { $C.Info    }
        }

        $subjectShort = if ($cert.Subject.Length -gt 50) { $cert.Subject.Substring(0,47) + '...' } else { $cert.Subject }
        $expiryStr    = $cert.Expiry.ToString('yyyy-MM-dd HH:mm')

        Write-Host ("  {0,-$($col[0])} {1,-$($col[1])} {2,-$($col[2])} {3,-$($col[3])} {4}" -f `
            $cert.Store, $subjectShort, $expiryStr, $cert.DaysLeft, $cert.Status) -ForegroundColor $rowColor
    }

    Write-Host ""
}

function Show-SslSummary {
    param([array]$SslResults)

    Write-Section "SSL/TLS REMOTE CERTIFICATE CHECKS"

    if (-not $SslResults -or $SslResults.Count -eq 0) {
        Write-Host "  No remote hosts checked." -ForegroundColor $C.Warning
        return
    }

    # Table header
    Write-Host ("  {0,-35} {1,-45} {2,-22} {3,-10} {4}" -f `
        'Host:Port', 'Subject', 'Expiry', 'Days Left', 'Status') -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 120)) -ForegroundColor $C.Header

    foreach ($r in $SslResults) {
        $hostPort = "$($r.Hostname):$($r.Port)"
        if ($hostPort.Length -gt 34) { $hostPort = $hostPort.Substring(0,31) + '...' }

        $rowColor = switch ($r.Status) {
            'Expired'  { $C.Error   }
            'Critical' { $C.Error   }
            'Warning'  { $C.Warning }
            'Error'    { $C.Error   }
            'Healthy'  { $C.Success }
            default    { $C.Info    }
        }

        if ($r.Status -eq 'Error') {
            $subjectShort = if ($r.Error -and $r.Error.Length -gt 44) { $r.Error.Substring(0,41) + '...' } else { $r.Error }
            Write-Host ("  {0,-35} {1,-45} {2,-22} {3,-10} {4}" -f `
                $hostPort, $subjectShort, 'N/A', 'N/A', 'Error') -ForegroundColor $rowColor
        } else {
            $subjectShort = if ($r.Subject.Length -gt 44) { $r.Subject.Substring(0,41) + '...' } else { $r.Subject }
            $expiryStr    = if ($r.Expiry) { $r.Expiry.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
            $daysStr      = if ($null -ne $r.DaysLeft) { "$($r.DaysLeft)" } else { 'N/A' }

            Write-Host ("  {0,-35} {1,-45} {2,-22} {3,-10} {4}" -f `
                $hostPort, $subjectShort, $expiryStr, $daysStr, $r.Status) -ForegroundColor $rowColor
        }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$LocalCerts,
        [array]$SslResults,
        [string]$MachineName,
        [string]$ReportTimestamp
    )

    # Summary counts
    $totalCerts   = if ($LocalCerts) { $LocalCerts.Count } else { 0 }
    $expiredCount = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Expired'  } | Measure-Object).Count } else { 0 }
    $critCount    = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Critical' } | Measure-Object).Count } else { 0 }
    $warnCount    = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Warning'  } | Measure-Object).Count } else { 0 }
    $healthyCount = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Healthy'  } | Measure-Object).Count } else { 0 }

    # Build local cert rows — ALL certs included in HTML, sorted by DaysLeft ascending
    $localRows = ''
    if ($LocalCerts) {
        foreach ($cert in ($LocalCerts | Sort-Object DaysLeft)) {
            $badgeClass = switch ($cert.Status) {
                'Expired'  { 'tk-badge-err'  }
                'Critical' { 'tk-badge-err'  }
                'Warning'  { 'tk-badge-warn' }
                default    { 'tk-badge-ok'   }
            }
            $expiryStr = $cert.Expiry.ToString('yyyy-MM-dd HH:mm')
            $daysStr   = if ($cert.DaysLeft -lt 0) { "$($cert.DaysLeft)" } else { "$($cert.DaysLeft)" }

            $localRows += "            <tr>
                <td><code>$(EscHtml $cert.Store)</code></td>
                <td class='tk-mono'>$(EscHtml $cert.Subject)</td>
                <td class='tk-mono'>$(EscHtml $cert.Issuer)</td>
                <td>$expiryStr</td>
                <td>$daysStr</td>
                <td><span class='$badgeClass'>$(EscHtml $cert.Status)</span></td>
            </tr>`n"
        }
    }

    # Build SSL rows
    $sslSection = ''
    if ($SslResults -and $SslResults.Count -gt 0) {
        $sslRows = ''
        foreach ($r in $SslResults) {
            $badgeClass = switch ($r.Status) {
                'Expired'  { 'tk-badge-err'  }
                'Critical' { 'tk-badge-err'  }
                'Warning'  { 'tk-badge-warn' }
                'Error'    { 'tk-badge-err'  }
                default    { 'tk-badge-ok'   }
            }

            if ($r.Status -eq 'Error') {
                $sslRows += "            <tr>
                <td class='tk-mono'><strong>$(EscHtml $r.Hostname)</strong>:$($r.Port)</td>
                <td class='tk-mono'>$(EscHtml $r.Error)</td>
                <td>N/A</td>
                <td>N/A</td>
                <td><span class='tk-badge-err'>Error</span></td>
            </tr>`n"
            } else {
                $expiryStr = if ($r.Expiry) { $r.Expiry.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
                $daysStr   = if ($null -ne $r.DaysLeft) { "$($r.DaysLeft)" } else { 'N/A' }
                $sslRows += "            <tr>
                <td class='tk-mono'><strong>$(EscHtml $r.Hostname)</strong>:$($r.Port)</td>
                <td class='tk-mono'>$(EscHtml $r.Subject)</td>
                <td>$expiryStr</td>
                <td>$daysStr</td>
                <td><span class='$badgeClass'>$(EscHtml $r.Status)</span></td>
            </tr>`n"
            }
        }

        $sslSection = @"

  <div class="tk-divider"></div>

  <div class="tk-card">
    <div class="tk-card-header">
      <span class="tk-card-label">SSL/TLS Remote Certificate Checks</span>
    </div>
    <table class="tk-table">
      <thead>
        <tr>
          <th>Host:Port</th>
          <th>Subject</th>
          <th>Expiry</th>
          <th>Days Left</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
$sslRows
      </tbody>
    </table>
  </div>
"@
    }

    $cfg = Get-TKConfig
    $orgSubtitle = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) {
        "$(EscHtml $cfg.OrgName)  -  $MachineName"
    } else {
        $MachineName
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'Certificate Audit Report' `
        -ScriptName 'P.H.O.E.N.I.X.' `
        -Subtitle    $orgSubtitle `
        -MetaItems  ([ordered]@{
            'Machine'   = $MachineName
            'Generated' = $ReportTimestamp
            'Stores'    = 'My, CA, Root, TrustedPublisher'
        }) `
        -NavItems   @('Local Certificates', 'SSL/TLS Checks')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'P.H.O.E.N.I.X. v6.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalCerts</div><div class="tk-summary-lbl">Total Certs</div></div>
    <div class="tk-summary-card err"><div class="tk-summary-num">$expiredCount</div><div class="tk-summary-lbl">Expired</div></div>
    <div class="tk-summary-card err"><div class="tk-summary-num">$critCount</div><div class="tk-summary-lbl">Critical (&lt;30d)</div></div>
    <div class="tk-summary-card warn"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warning (&lt;90d)</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$healthyCount</div><div class="tk-summary-lbl">Healthy</div></div>
  </div>

  <div class="tk-card">
    <div class="tk-card-header">
      <span class="tk-card-label">Local Certificate Stores</span>
    </div>
    <table class="tk-table">
      <thead>
        <tr>
          <th>Store</th>
          <th>Subject</th>
          <th>Issuer</th>
          <th>Expiry</th>
          <th>Days Left</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
$localRows
      </tbody>
    </table>
  </div>
$sslSection
"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT HELPER
# ─────────────────────────────────────────────────────────────────────────────

function Save-HtmlReport {
    param(
        [array]$LocalCerts,
        [array]$SslResults
    )

    $machineName     = $env:COMPUTERNAME
    $reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    Write-Host "  [*] Building HTML report..." -ForegroundColor $C.Progress

    $html = Build-HtmlReport `
        -LocalCerts      $LocalCerts `
        -SslResults      $SslResults `
        -MachineName     $machineName `
        -ReportTimestamp $reportTimestamp

    $reportFilename = "PHOENIX_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $reportPath     = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportFilename

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
    } catch {
        Write-Host "  [-] Could not save report: $_" -ForegroundColor $C.Error
        $reportPath = $null
    }

    return $reportPath
}

# ─────────────────────────────────────────────────────────────────────────────
# UNATTENDED MODE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-PhoenixBanner

    $machineName = $env:COMPUTERNAME
    Write-Host "  [*] Unattended mode  -  Machine: $machineName" -ForegroundColor $C.Progress
    Write-Host ""

    # Local cert audit
    Write-Host "  [*] Auditing local certificate stores..." -ForegroundColor $C.Progress
    $localCerts = Get-LocalCertHealth

    $expiredCount = ($localCerts | Where-Object { $_.Status -eq 'Expired'  } | Measure-Object).Count
    $critCount    = ($localCerts | Where-Object { $_.Status -eq 'Critical' } | Measure-Object).Count
    $warnCount    = ($localCerts | Where-Object { $_.Status -eq 'Warning'  } | Measure-Object).Count
    $healthyCount = ($localCerts | Where-Object { $_.Status -eq 'Healthy'  } | Measure-Object).Count

    Write-Host ("  [+] Found {0} certificate(s): Expired={1}, Critical={2}, Warning={3}, Healthy={4}" -f `
        $localCerts.Count, $expiredCount, $critCount, $warnCount, $healthyCount) -ForegroundColor $C.Success

    # SSL checks
    $sslResults = [System.Collections.Generic.List[object]]::new()
    if (-not [string]::IsNullOrWhiteSpace($Targets)) {
        $targetList = ConvertTo-TargetList -TargetString $Targets
        if ($targetList.Count -gt 0) {
            Write-Host ""
            Write-Host "  [*] Checking SSL/TLS certificates on $($targetList.Count) remote host(s)..." -ForegroundColor $C.Progress
            foreach ($t in $targetList) {
                Write-Host "  [*] Checking $($t.Hostname):$($t.Port)..." -ForegroundColor $C.Progress
                $sslResults.AddRange([object[]]@(Get-SslCertExpiry -Hostname $t.Hostname -Port $t.Port))
            }
        }
    }

    # Generate HTML report
    Write-Host ""
    $reportPath = Save-HtmlReport -LocalCerts $localCerts -SslResults $sslResults
    if ($reportPath) { Show-TKReportResult -Path $reportPath -Unattended:$Unattended }

    # Console summary
    Write-Host ""
    Write-Host ("  " + ("=" * 62)) -ForegroundColor $C.Header
    Write-Host "  P.H.O.E.N.I.X. AUDIT SUMMARY" -ForegroundColor $C.Header
    Write-Host ("  " + ("=" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  Total Certificates : $($localCerts.Count)" -ForegroundColor $C.Info
    Write-Host "  Expired            : $expiredCount" -ForegroundColor $(if ($expiredCount -gt 0) { $C.Error } else { $C.Info })
    Write-Host "  Critical (<30d)    : $critCount"   -ForegroundColor $(if ($critCount    -gt 0) { $C.Error } else { $C.Info })
    Write-Host "  Warning  (<90d)    : $warnCount"   -ForegroundColor $(if ($warnCount    -gt 0) { $C.Warning } else { $C.Info })
    Write-Host "  Healthy            : $healthyCount" -ForegroundColor $C.Success
    if ($sslResults.Count -gt 0) {
        $sslErrors  = ($sslResults | Where-Object { $_.Status -eq 'Error'    } | Measure-Object).Count
        $sslExpired = ($sslResults | Where-Object { $_.Status -eq 'Expired'  } | Measure-Object).Count
        $sslCrit    = ($sslResults | Where-Object { $_.Status -eq 'Critical' } | Measure-Object).Count
        Write-Host ""
        Write-Host "  SSL Hosts Checked  : $($sslResults.Count)" -ForegroundColor $C.Info
        Write-Host "  SSL Errors         : $sslErrors"  -ForegroundColor $(if ($sslErrors  -gt 0) { $C.Error   } else { $C.Info })
        Write-Host "  SSL Expired        : $sslExpired" -ForegroundColor $(if ($sslExpired -gt 0) { $C.Error   } else { $C.Info })
        Write-Host "  SSL Critical       : $sslCrit"    -ForegroundColor $(if ($sslCrit    -gt 0) { $C.Error   } else { $C.Info })
    }
    Write-Host ""
    Write-Host ("  " + ("=" * 62)) -ForegroundColor $C.Header
    Write-Host "  P.H.O.E.N.I.X. UNATTENDED RUN COMPLETE" -ForegroundColor $C.Header
    Write-Host ("  " + ("=" * 62)) -ForegroundColor $C.Header
    Write-Host ""

    if ($Transcript) { Stop-TKTranscript }
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE MENU
# ─────────────────────────────────────────────────────────────────────────────

Show-PhoenixBanner

$choice = ''

do {
    Show-PhoenixBanner

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  MAIN MENU" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  [1] Audit local certificate stores" -ForegroundColor $C.Info
    Write-Host "  [2] Check SSL certificate expiry on remote hosts" -ForegroundColor $C.Info
    Write-Host "  [3] Full audit  -  local stores + SSL + HTML report" -ForegroundColor $C.Info
    Write-Host "  [Q] Quit" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $choice = (Read-Host).Trim().ToUpper()

    switch ($choice) {

        # ── Option 1: Local cert audit ──────────────────────────────────────
        '1' {
            Show-PhoenixBanner
            Write-Host "  [*] Auditing local certificate stores (My, CA, Root, TrustedPublisher)..." -ForegroundColor $C.Progress
            $localCerts = Get-LocalCertHealth
            Write-Host "  [+] Found $($localCerts.Count) certificate(s)." -ForegroundColor $C.Success

            Show-LocalCertSummary -Certs $localCerts

            Write-Host -NoNewline "  Export HTML report? (Y/N): " -ForegroundColor $C.Header
            $exportChoice = (Read-Host).Trim().ToUpper()
            if ($exportChoice -eq 'Y') {
                $reportPath = Save-HtmlReport -LocalCerts $localCerts -SslResults @()
                if ($reportPath) {
                    Show-TKReportResult -Path $reportPath -Unattended:$Unattended
                }
            }

            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

        # ── Option 2: SSL remote checks ─────────────────────────────────────
        '2' {
            Show-PhoenixBanner
            Write-Section "SSL/TLS REMOTE CERTIFICATE CHECK"

            Write-Host "  Enter target hosts to check." -ForegroundColor $C.Info
            Write-Host "  Format: hostname or hostname:port (comma-separated, or a file path)" -ForegroundColor $C.Info
            Write-Host "  Examples: server1.contoso.com, mail.contoso.com:443, 10.0.0.5:8443" -ForegroundColor $C.Info
            Write-Host ""
            Write-Host -NoNewline "  Targets: " -ForegroundColor $C.Header
            $rawTargets = (Read-Host).Trim()

            if ([string]::IsNullOrWhiteSpace($rawTargets)) {
                Write-Host "  [!!] No targets entered." -ForegroundColor $C.Warning
            } else {
                $targetList = ConvertTo-TargetList -TargetString $rawTargets
                if ($targetList.Count -eq 0) {
                    Write-Host "  [!!] No valid targets parsed." -ForegroundColor $C.Warning
                } else {
                    $sslResults = [System.Collections.Generic.List[object]]::new()
                    Write-Host ""
                    foreach ($t in $targetList) {
                        Write-Host "  [*] Checking $($t.Hostname):$($t.Port)..." -ForegroundColor $C.Progress
                        $sslResults.AddRange([object[]]@(Get-SslCertExpiry -Hostname $t.Hostname -Port $t.Port))
                    }

                    Show-SslSummary -SslResults $sslResults

                    Write-Host -NoNewline "  Export HTML report? (Y/N): " -ForegroundColor $C.Header
                    $exportChoice = (Read-Host).Trim().ToUpper()
                    if ($exportChoice -eq 'Y') {
                        $reportPath = Save-HtmlReport -LocalCerts @() -SslResults $sslResults
                        if ($reportPath) {
                            Show-TKReportResult -Path $reportPath -Unattended:$Unattended
                        }
                    }
                }
            }

            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

        # ── Option 3: Full audit ────────────────────────────────────────────
        '3' {
            Show-PhoenixBanner

            # Local cert audit
            Write-Host "  [*] Auditing local certificate stores..." -ForegroundColor $C.Progress
            $localCerts = Get-LocalCertHealth
            Write-Host "  [+] Found $($localCerts.Count) certificate(s)." -ForegroundColor $C.Success

            # SSL targets
            Write-Host ""
            Write-Host "  Enter SSL/TLS target hosts (optional  -  press Enter to skip)." -ForegroundColor $C.Info
            Write-Host "  Format: hostname or hostname:port (comma-separated, or a file path)" -ForegroundColor $C.Info
            Write-Host ""
            Write-Host -NoNewline "  Targets: " -ForegroundColor $C.Header
            $rawTargets = (Read-Host).Trim()

            $sslResults = [System.Collections.Generic.List[object]]::new()
            if (-not [string]::IsNullOrWhiteSpace($rawTargets)) {
                $targetList = ConvertTo-TargetList -TargetString $rawTargets
                if ($targetList.Count -gt 0) {
                    Write-Host ""
                    foreach ($t in $targetList) {
                        Write-Host "  [*] Checking $($t.Hostname):$($t.Port)..." -ForegroundColor $C.Progress
                        $sslResults.AddRange([object[]]@(Get-SslCertExpiry -Hostname $t.Hostname -Port $t.Port))
                    }
                }
            }

            # Show console summaries
            Show-LocalCertSummary -Certs $localCerts
            if ($sslResults.Count -gt 0) {
                Show-SslSummary -SslResults $sslResults
            }

            # Generate report automatically
            $reportPath = Save-HtmlReport -LocalCerts $localCerts -SslResults $sslResults

            if ($reportPath) {
                Show-TKReportResult -Path $reportPath -Unattended:$Unattended
            }

            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

        # ── Quit ─────────────────────────────────────────────────────────────
        'Q' {
            Write-Host ""
            Write-Host "  Closing P.H.O.E.N.I.X." -ForegroundColor $C.Header
            Write-Host ""
        }

        default {
            Write-Host ""
            Write-Host "  [!!] Invalid selection. Enter 1, 2, 3, or Q." -ForegroundColor $C.Warning
            Start-Sleep -Seconds 1
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
