# exhume.ps1 - E.X.H.U.M.E. — Enumerates, eXposes & Hunts Unmigrated Mail Entries
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
    E.X.H.U.M.E. — Enumerates, eXposes & Hunts Unmigrated Mail Entries
    Outlook PST / OST Discovery Tool for PowerShell 5.1+

.DESCRIPTION
    Scans the machine for Outlook data files and attached profiles before a
    mail migration. Inventories every .pst (and optionally .ost) on local
    drives, cross-references them against configured Outlook profiles,
    flags orphaned and stale archives, and highlights large files that
    exceed Exchange Online import limits. Produces a dark-themed HTML
    report with a migration readiness summary.

.USAGE
    PS C:\> .\exhume.ps1                                 # Interactive: scans all local drives, skips .ost
    PS C:\> .\exhume.ps1 -Unattended                    # Silent mode: exports HTML and exits
    PS C:\> .\exhume.ps1 -ScanDrives C:,D: -IncludeOst  # Custom drive list and include .ost

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [string[]]$ScanDrives = @(),
    [switch]$IncludeOst
)

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

function Write-Step { param([string]$Msg) Write-Host ("  [*] {0}" -f $Msg) -ForegroundColor Magenta }

function Write-Ok   { param([string]$Msg) Write-Host ("  [+] {0}" -f $Msg) -ForegroundColor Green   }

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

function Write-Fail { param([string]$Msg) Write-Host ("  [-] {0}" -f $Msg) -ForegroundColor Red     }

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

function Format-Bytes {
    <#
    .SYNOPSIS
        Formats a byte count as a human-readable string (B/KB/MB/GB/TB).
    .EXAMPLE
        Format-Bytes 1572864   # -> '1.50 MB'
    #>
    param([Parameter(Mandatory)][long]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes B" }
    if ($Bytes -lt 1MB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -lt 1TB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    return "{0:N2} TB" -f ($Bytes / 1TB)
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
#endregion TK-SHARED

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

function Show-ExhumeBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  E.X.H.U.M.E. — Enumerates, eXposes & Hunts Unmigrated Mail Entries" -ForegroundColor Cyan
    Write-Host "  Outlook PST / OST Discovery Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Functions and main entry point appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — OUTLOOK PROFILE DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────

# Walks every Outlook profile registered under the current user and extracts
# the paths of any PST / OST stores each profile has mounted. Works across
# Outlook 2016 / 2019 / 2021 / 365 (all use the 16.0 hive).
function Get-OutlookProfiles {
    $roots = @(
        'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles',
        'HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles',
        'HKCU:\Software\Microsoft\Office\14.0\Outlook\Profiles'
    )

    $defaultProfile = $null
    try {
        $defaultProfile = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
    } catch {
        # DefaultProfile key missing is normal on machines without Outlook configured.
    }

    $profiles = foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $version = if     ($root -match '16\.0') { 'Outlook 2016+' }
                   elseif ($root -match '15\.0') { 'Outlook 2013' }
                   elseif ($root -match '14\.0') { 'Outlook 2010' }
                   else { 'Unknown' }

        foreach ($p in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $stores = [System.Collections.Generic.List[string]]::new()
            # Each profile has N subkeys for each account / store. PST path lives in
            # the '001f6700' binary value (UTF-16 encoded) on store entries.
            try {
                $all = Get-ChildItem $p.PSPath -Recurse -ErrorAction SilentlyContinue
                foreach ($sub in $all) {
                    $props = Get-ItemProperty -Path $sub.PSPath -ErrorAction SilentlyContinue
                    if (-not $props) { continue }
                    foreach ($pn in $props.PSObject.Properties.Name) {
                        if ($pn -eq '001f6700') {
                            $raw  = $props.$pn
                            if ($raw -is [byte[]]) {
                                $str = [System.Text.Encoding]::Unicode.GetString($raw) -replace "`0",''
                                if ($str -match '\.(pst|ost)$') {
                                    $stores.Add($str)
                                }
                            }
                        }
                    }
                }
            } catch {
                # Some sub-keys are unreadable under constrained token ACLs. Move on
                # and surface whatever stores we managed to collect.
            }

            [PSCustomObject]@{
                Name        = $p.PSChildName
                IsDefault   = ($p.PSChildName -eq $defaultProfile)
                Version     = $version
                Stores      = @($stores | Select-Object -Unique)
            }
        }
    }

    return @($profiles)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — DATA FILE INVENTORY (DRIVE SCAN)
# ─────────────────────────────────────────────────────────────────────────────

# Default exclusions: system dirs and OneDrive sync cache (PSTs should not live
# in OneDrive — if one is there it's a migration landmine, but the scan is
# already slow and these caches are noisy).
$script:ScanExclusions = @(
    'Windows', 'Program Files', 'Program Files (x86)', 'ProgramData',
    '$Recycle.Bin', 'System Volume Information', 'Recovery'
)

function Get-DriveRoots {
    param([string[]]$Override)

    if ($Override -and $Override.Count -gt 0) {
        return @($Override | ForEach-Object {
            $d = $_.TrimEnd(':','\','/')
            "${d}:\"
        })
    }

    # Fixed local drives only — skip CD-ROM, network, and removable by default.
    $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^[A-Z]$' -and $null -ne $_.Used }
    return @($drives | ForEach-Object { "$($_.Name):\" })
}

function Find-DataFiles {
    param(
        [string[]]$Roots,
        [bool]$IncludeOstFlag
    )

    $patterns = if ($IncludeOstFlag) { @('*.pst', '*.ost') } else { @('*.pst') }
    $found    = [System.Collections.Generic.List[object]]::new()

    foreach ($root in $Roots) {
        if (-not (Test-Path $root)) { continue }
        Write-Step "Scanning $root for Outlook data files..."

        # Enumerate top-level children so we can skip the standard exclusions without
        # paying for a full recursion into them.
        $topLevel = Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue |
                    Where-Object { $script:ScanExclusions -notcontains $_.Name }

        # Also hit loose files in the drive root (rare but possible).
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Filter $pat -File -Force -ErrorAction SilentlyContinue |
                ForEach-Object { $found.Add($_) }
        }

        foreach ($dir in $topLevel) {
            foreach ($pat in $patterns) {
                try {
                    Get-ChildItem -LiteralPath $dir.FullName -Filter $pat -Recurse -File -Force -ErrorAction SilentlyContinue |
                        ForEach-Object { $found.Add($_) }
                } catch {
                    # Access-denied on a sub-tree is expected for other-user profiles
                    # when not running as admin; continue with the next directory.
                }
            }
        }
    }

    $now = Get-Date
    $rows = foreach ($f in $found) {
        $daysSinceModified = [math]::Round(($now - $f.LastWriteTime).TotalDays, 0)
        $daysSinceAccessed = $null
        try {
            if ($f.LastAccessTime) { $daysSinceAccessed = [math]::Round(($now - $f.LastAccessTime).TotalDays, 0) }
        } catch {
            # LastAccessTime can be $null or raise on some FS configs; optional field.
        }

        [PSCustomObject]@{
            FullPath          = $f.FullName
            Name              = $f.Name
            Extension         = $f.Extension.ToLower()
            Bytes             = $f.Length
            Size              = Format-Bytes $f.Length
            LastModified      = $f.LastWriteTime.ToString('yyyy-MM-dd')
            DaysSinceModified = $daysSinceModified
            DaysSinceAccessed = $daysSinceAccessed
        }
    }

    return @($rows | Sort-Object Bytes -Descending)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — PROFILE/FILE CORRELATION
# ─────────────────────────────────────────────────────────────────────────────

# Annotates the disk inventory with "attached to which profile" and marks any
# file that no profile references as orphaned.
function Add-ProfileAttachment {
    param([array]$Files, [array]$Profiles)

    # Build a case-insensitive lookup of store-path -> profile-names.
    $map = @{}
    foreach ($p in $Profiles) {
        foreach ($s in $p.Stores) {
            $key = $s.ToLower()
            if (-not $map.ContainsKey($key)) { $map[$key] = [System.Collections.Generic.List[string]]::new() }
            [void]$map[$key].Add($p.Name)
        }
    }

    foreach ($f in $Files) {
        $attached = @()
        if ($map.ContainsKey($f.FullPath.ToLower())) {
            $attached = $map[$f.FullPath.ToLower()]
        }
        $f | Add-Member -NotePropertyName 'AttachedTo'  -NotePropertyValue (@($attached) -join ', ') -Force
        $f | Add-Member -NotePropertyName 'IsOrphaned'  -NotePropertyValue ($attached.Count -eq 0) -Force
    }
    return $Files
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-Verdict {
    param([array]$Files, [array]$Profiles)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    $psts       = @($Files | Where-Object { $_.Extension -eq '.pst' })
    $orphaned   = @($Files | Where-Object { $_.IsOrphaned -and $_.Extension -eq '.pst' })
    $oversize   = @($Files | Where-Object { $_.Extension -eq '.pst' -and $_.Bytes -ge 50GB })
    $large      = @($Files | Where-Object { $_.Extension -eq '.pst' -and $_.Bytes -ge 10GB -and $_.Bytes -lt 50GB })
    $stale      = @($Files | Where-Object { $_.Extension -eq '.pst' -and $_.DaysSinceAccessed -and $_.DaysSinceAccessed -ge 365 })

    if ($orphaned.Count -gt 0) {
        $warns.Add("$($orphaned.Count) PST(s) on disk are not attached to any Outlook profile — review whether each should be ingested or deleted.")
    }
    foreach ($f in $oversize) {
        $issues.Add("PST '$($f.FullPath)' is $($f.Size) — Exchange Online Import Service has a 50 GB hard limit per file; split before ingest.")
    }
    if ($large.Count -gt 0) {
        $warns.Add("$($large.Count) PST(s) are 10-50 GB — migrations will be slow; consider splitting or ingesting overnight.")
    }
    if ($stale.Count -gt 0) {
        $warns.Add("$($stale.Count) PST(s) have not been accessed in 365+ days — candidates for archive-on-ingest rather than primary-mailbox import.")
    }
    if ($psts.Count -eq 0) {
        $warns.Add('No PSTs found on disk — either already migrated or never existed on this machine.')
    }

    $verdict = if ($issues.Count -gt 0) { 'ACTION REQUIRED' }
               elseif ($warns.Count -gt 0) { 'REVIEW BEFORE MIGRATION' }
               else { 'READY TO MIGRATE' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict   = $verdict
        Class     = $class
        Issues    = @($issues)
        Warns     = @($warns)
        PstCount  = $psts.Count
        Orphans   = $orphaned.Count
        Oversize  = $oversize.Count
        Large     = $large.Count
        Stale     = $stale.Count
        TotalPstBytes = (($psts | Measure-Object -Property Bytes -Sum).Sum)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Files, [array]$Profiles, $Verdict, [array]$Roots, [bool]$IncludeOstFlag)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $runUser    = "$env:USERDOMAIN\$env:USERNAME"

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    # Profile table
    $profileRows = [System.Text.StringBuilder]::new()
    if ($Profiles.Count -eq 0) {
        [void]$profileRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No Outlook profiles configured for the current user.</td></tr>")
    } else {
        foreach ($p in $Profiles) {
            $defBadge = if ($p.IsDefault) { "<span class='tk-badge-ok'>Default</span>" } else { "<span class='tk-badge-info'>Secondary</span>" }
            $storeList = if ($p.Stores.Count -gt 0) {
                ($p.Stores | ForEach-Object { "<code>$(EscHtml $_)</code>" }) -join '<br>'
            } else { "<span class='tk-badge-info'>(none)</span>" }
            [void]$profileRows.Append("<tr><td>$(EscHtml $p.Name)</td><td>$defBadge</td><td>$(EscHtml $p.Version)</td><td>$storeList</td></tr>`n")
        }
    }

    # File inventory table
    $fileRows = [System.Text.StringBuilder]::new()
    if ($Files.Count -eq 0) {
        [void]$fileRows.Append("<tr><td colspan='6' class='tk-badge-ok' style='text-align:center;'>No Outlook data files found on the scanned drives.</td></tr>")
    } else {
        foreach ($f in $Files) {
            $sizeClass = if ($f.Bytes -ge 50GB) { 'tk-badge-err' }
                         elseif ($f.Bytes -ge 10GB) { 'tk-badge-warn' }
                         elseif ($f.Bytes -ge 1GB)  { 'tk-badge-info' }
                         else { 'tk-badge-ok' }
            $orphanCell = if ($f.IsOrphaned) { "<span class='tk-badge-warn'>Orphaned</span>" } else { "<span class='tk-badge-ok'>$(EscHtml $f.AttachedTo)</span>" }
            $extCell    = if ($f.Extension -eq '.ost') { "<span class='tk-badge-info'>OST</span>" } else { "<span class='tk-badge-ok'>PST</span>" }
            [void]$fileRows.Append("<tr><td><code>$(EscHtml $f.FullPath)</code></td><td>$extCell</td><td><span class='$sizeClass'>$(EscHtml $f.Size)</span></td><td>$(EscHtml $f.LastModified)</td><td>$orphanCell</td><td>$($f.DaysSinceAccessed)</td></tr>`n")
        }
    }

    # Verdict block
    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>All pre-migration checks passed.</li>")
    }

    $scopeLabel = "$($Roots -join ', ') ($(if ($IncludeOstFlag) { 'PST + OST' } else { 'PST only' }))"

    $htmlHead = Get-TKHtmlHead `
        -Title      'E.X.H.U.M.E. PST Discovery Report' `
        -ScriptName 'E.X.H.U.M.E.' `
        -Subtitle   "${orgPrefix}Outlook PST / OST Pre-Migration Discovery -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Run As'    = $runUser
            'Generated' = $reportDate
            'Scope'     = $scopeLabel
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'Outlook Profiles', 'Data Files')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'E.X.H.U.M.E. v6.0'

    $totalSize = Format-Bytes $Verdict.TotalPstBytes

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Migration Readiness</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Verdict.PstCount)</div><div class="tk-summary-lbl">PST Files Found</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalSize</div><div class="tk-summary-lbl">Total PST Size</div></div>
    <div class="tk-summary-card $(if ($Verdict.Orphans -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Verdict.Orphans)</div><div class="tk-summary-lbl">Orphaned (not in any profile)</div></div>
    <div class="tk-summary-card $(if ($Verdict.Oversize -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($Verdict.Oversize)</div><div class="tk-summary-lbl">Over 50 GB (import blocker)</div></div>
    <div class="tk-summary-card $(if ($Verdict.Stale -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Verdict.Stale)</div><div class="tk-summary-lbl">Stale (&gt;365 days)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Verdict &amp; Findings</span>
      <span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span>
    </div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Outlook Profiles</span><span class="tk-section-num">$($Profiles.Count) profile(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Default</th><th>Version</th><th>Attached Stores</th></tr></thead>
        <tbody>$($profileRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Data Files (.pst$(if ($IncludeOstFlag) { ' + .ost' }))</span><span class="tk-section-num">$($Files.Count) file(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Path</th><th>Type</th><th>Size</th><th>Last Modified</th><th>Attached To</th><th>Days Since Last Access</th></tr></thead>
        <tbody>$($fileRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

function Write-ConsoleSummary {
    param([array]$Files, [array]$Profiles, $Verdict, [array]$Roots, [bool]$IncludeOstFlag)

    Write-Section "OUTLOOK PROFILES"
    if ($Profiles.Count -eq 0) {
        Write-Warn "No Outlook profiles configured for the current user."
    } else {
        foreach ($p in $Profiles) {
            $defLabel = if ($p.IsDefault) { ' [DEFAULT]' } else { '' }
            Write-Host "  $($p.Name)$defLabel  ($($p.Version))" -ForegroundColor $C.Header
            if ($p.Stores.Count -eq 0) {
                Write-Host "    (no attached stores)" -ForegroundColor $C.Info
            } else {
                foreach ($s in $p.Stores) { Write-Host "    * $s" -ForegroundColor $C.Info }
            }
        }
    }

    Write-Section "DATA FILE INVENTORY"
    $scopeLabel = "$($Roots -join ', ') ($(if ($IncludeOstFlag) { 'PST + OST' } else { 'PST only' }))"
    Write-Info "Scanned: $scopeLabel"
    if ($Files.Count -eq 0) {
        Write-Ok "No Outlook data files found."
    } else {
        foreach ($f in $Files | Select-Object -First 30) {
            $color = if ($f.Bytes -ge 50GB) { $C.Error }
                     elseif ($f.Bytes -ge 10GB) { $C.Warning }
                     elseif ($f.IsOrphaned) { $C.Warning }
                     else { $C.Info }
            $orphanTag = if ($f.IsOrphaned) { '  [ORPHAN]' } else { "  [$($f.AttachedTo)]" }
            Write-Host ("  {0,10}  {1}{2}" -f $f.Size, $f.FullPath, $orphanTag) -ForegroundColor $color
        }
        if ($Files.Count -gt 30) {
            Write-Info "... and $($Files.Count - 30) more (see HTML report for full list)."
        }
    }

    Write-Section "MIGRATION READINESS VERDICT"
    $vColor = switch ($Verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
    Write-Host "  $($Verdict.Verdict)" -ForegroundColor $vColor
    foreach ($i in $Verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
    foreach ($w in $Verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-ExhumeBanner

$roots = Get-DriveRoots -Override $ScanDrives
if ($roots.Count -eq 0) {
    Write-Fail "No local drives detected to scan."
    exit 1
}

Write-Section "OUTLOOK DATA FILE DISCOVERY"
Write-Info "Drives: $($roots -join ', ')"
Write-Info "Include OST: $IncludeOst"
Write-Host ""

$profiles = Get-OutlookProfiles
$files    = Find-DataFiles -Roots $roots -IncludeOstFlag:$IncludeOst
$files    = Add-ProfileAttachment -Files $files -Profiles $profiles
$verdict  = Get-Verdict -Files $files -Profiles $profiles

Write-ConsoleSummary -Files $files -Profiles $profiles -Verdict $verdict -Roots $roots -IncludeOstFlag:$IncludeOst

Write-Host ""
Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -Files $files -Profiles $profiles -Verdict $verdict -Roots $roots -IncludeOstFlag:$IncludeOst
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "EXHUME_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
