# augur.ps1 - A.U.G.U.R. — Analyzes, Uncovers & Gauges Unit Reliability
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
    A.U.G.U.R. — Analyzes, Uncovers & Gauges Unit Reliability
    Disk Health & SMART Assessment Tool for PowerShell 5.1+

.DESCRIPTION
    Inspects every physical disk in the system: health status, operational
    state, SMART failure prediction, volume integrity, and bus/media type.
    Flags disks that report degraded health or predicted failures and exports
    a dark-themed HTML report to the script directory.

.USAGE
    PS C:\> .\augur.ps1                    # Must be run as Administrator
    PS C:\> .\augur.ps1 -Unattended        # Silent mode — no prompts, no banner

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# INITIALIZATION
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

function Invoke-AdminElevation {
    param([Parameter(Mandatory)][string]$ScriptFile)
    if (-not (Test-IsAdmin)) {
        Write-Host "  INFO: Restarting with administrator privileges..." -ForegroundColor Yellow
        $PSExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        Start-Process -FilePath $PSExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptFile`"" -Verb RunAs
        exit
    }
}
#endregion TK-SHARED
Invoke-AdminElevation -ScriptFile $PSCommandPath

$ScriptPath = $PSScriptRoot

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

function Show-AugurBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   █████╗ ██╗   ██╗ ██████╗ ██╗   ██╗██████╗
  ██╔══██╗██║   ██║██╔════╝ ██║   ██║██╔══██╗
  ███████║██║   ██║██║  ███╗██║   ██║██████╔╝
  ██╔══██║██║   ██║██║   ██║██║   ██║██╔══██╗
  ██║  ██║╚██████╔╝╚██████╔╝╚██████╔╝██║  ██║
  ╚═╝  ╚═╝ ╚═════╝  ╚═════╝  ╚═════╝ ╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    A.U.G.U.R. — Analyzes, Uncovers & Gauges Unit Reliability" -ForegroundColor Cyan
    Write-Host "    Disk Health & SMART Assessment Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-AugurBanner }

$reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$reportFilename  = "AUGUR_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath      = Join-Path $ScriptPath $reportFilename

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  A.U.G.U.R. DISK HEALTH ASSESSMENT" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Timestamp : $reportTimestamp" -ForegroundColor $ColorSchema.Info
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: PHYSICAL DISKS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "  [1/3] Reading physical disk health..." -ForegroundColor $ColorSchema.Progress

$diskReport = [System.Collections.Generic.List[object]]::new()

try {
    $physDisks = Get-PhysicalDisk -ErrorAction Stop

    # SMART failure prediction via WMI
    $smartData = @{}
    try {
        $smartRaw = Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue
        foreach ($s in $smartRaw) {
            # InstanceName typically ends in _0, _1, etc. — use it as a loose key
            $key = ($s.InstanceName -split '\\' | Select-Object -Last 1) -replace '_\d+$',''
            $smartData[$key] = $s
        }
    } catch {}

    # Win32_DiskDrive for serial/firmware
    $wmiDisks = @{}
    try {
        Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
            $wmiDisks[$_.Index] = $_
        }
    } catch {}

    foreach ($pd in $physDisks) {
        $sizeGB = if ($pd.Size -gt 0) { [math]::Round($pd.Size / 1GB, 1) } else { 0 }

        # Match SMART entry by device index or friendly name fragment
        $smartEntry  = $smartData.Values | Where-Object { $_.InstanceName -match [regex]::Escape($pd.DeviceId) } | Select-Object -First 1
        $smartFail   = if ($smartEntry) { $smartEntry.PredictFailure } else { $null }
        $smartReason = if ($smartEntry -and $smartEntry.Reason) { "0x{0:X8}" -f $smartEntry.Reason } else { 'N/A' }

        $wmiDisk   = $wmiDisks[$pd.DeviceId]
        $serial    = if ($wmiDisk -and $wmiDisk.SerialNumber) { $wmiDisk.SerialNumber.Trim() } else { 'N/A' }
        $firmware  = if ($wmiDisk -and $wmiDisk.FirmwareRevision) { $wmiDisk.FirmwareRevision.Trim() } else { 'N/A' }

        $healthColor = switch ($pd.HealthStatus) {
            'Healthy' { $ColorSchema.Success }
            'Warning' { $ColorSchema.Warning }
            default   { $ColorSchema.Error   }
        }

        $smartLabel = if ($null -eq $smartFail)   { 'N/A' }
                      elseif ($smartFail -eq $false) { 'OK' }
                      else                           { 'FAILING' }

        $displayLine = "  [{0}] {1} | {2} {3} | {4} GB | Health: {5} | SMART: {6}" -f `
            $pd.DeviceId, $pd.FriendlyName, $pd.MediaType, $pd.BusType, $sizeGB, $pd.HealthStatus, $smartLabel

        $lineColor = if ($smartFail -eq $true -or $pd.HealthStatus -ne 'Healthy') {
            $ColorSchema.Error
        } elseif ($pd.HealthStatus -eq 'Warning') {
            $ColorSchema.Warning
        } else {
            $ColorSchema.Success
        }

        Write-Host $displayLine -ForegroundColor $lineColor

        $diskReport.Add([PSCustomObject]@{
            ID                = $pd.DeviceId
            Name              = $pd.FriendlyName
            Serial            = $serial
            Firmware          = $firmware
            MediaType         = $pd.MediaType
            BusType           = $pd.BusType
            SizeGB            = $sizeGB
            HealthStatus      = $pd.HealthStatus
            OperationalStatus = $pd.OperationalStatus
            SMARTPrediction   = $smartLabel
            SMARTReason       = $smartReason
        })
    }

    Write-Host ""
    Write-Host "  [+] $($diskReport.Count) physical disk(s) assessed" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Error reading physical disks: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: VOLUME HEALTH
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "  [2/3] Reading volume health..." -ForegroundColor $ColorSchema.Progress

$volumeReport = [System.Collections.Generic.List[object]]::new()

try {
    $volumes = Get-Volume -ErrorAction Stop | Where-Object { $_.DriveLetter -or $_.FileSystemLabel }

    foreach ($vol in $volumes) {
        $totalGB = if ($vol.Size -gt 0) { [math]::Round($vol.Size / 1GB, 1) } else { 0 }
        $freeGB  = if ($vol.SizeRemaining -gt 0) { [math]::Round($vol.SizeRemaining / 1GB, 1) } else { 0 }
        $pct     = if ($vol.Size -gt 0) { [math]::Round(($vol.Size - $vol.SizeRemaining) / $vol.Size * 100, 1) } else { 0 }

        $healthColor = switch ($vol.HealthStatus) {
            'Healthy' { $ColorSchema.Success }
            'Warning' { $ColorSchema.Warning }
            default   { $ColorSchema.Error   }
        }

        $label  = if ($vol.DriveLetter)      { "$($vol.DriveLetter):" } else { "(no letter)" }
        $fsLabel = if ($vol.FileSystemLabel) { $vol.FileSystemLabel }   else { "" }

        Write-Host ("  {0,-6} {1,-20} {2,7} GB total  {3,7} GB free  {4,5}%  [{5}]" -f `
            $label, $fsLabel, $totalGB, $freeGB, $pct, $vol.HealthStatus) -ForegroundColor $healthColor

        $volumeReport.Add([PSCustomObject]@{
            Drive       = $label
            Label       = $fsLabel
            FileSystem  = $vol.FileSystem
            TotalGB     = $totalGB
            FreeGB      = $freeGB
            PctUsed     = $pct
            Health      = $vol.HealthStatus
            DriveType   = $vol.DriveType
        })
    }

    Write-Host ""
    Write-Host "  [+] $($volumeReport.Count) volume(s) assessed" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Error reading volumes: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: GENERATE HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "  [3/3] Generating HTML report..." -ForegroundColor $ColorSchema.Progress

$healthyCount   = ($diskReport | Where-Object { $_.HealthStatus -eq 'Healthy' -and $_.SMARTPrediction -ne 'FAILING' }).Count
$warningCount   = ($diskReport | Where-Object { $_.HealthStatus -eq 'Warning' }).Count
$criticalCount  = ($diskReport | Where-Object { $_.HealthStatus -ne 'Healthy' -and $_.HealthStatus -ne 'Warning' }).Count
$smartFailCount = ($diskReport | Where-Object { $_.SMARTPrediction -eq 'FAILING' }).Count

# Physical disk rows
$diskRows = ""
foreach ($d in $diskReport) {
    $hBadge = switch ($d.HealthStatus) {
        'Healthy' { 'tk-badge-ok' } 'Warning' { 'tk-badge-warn' } default { 'tk-badge-err' }
    }
    $sBadge = if ($d.SMARTPrediction -eq 'FAILING') { 'tk-badge-err' } elseif ($d.SMARTPrediction -eq 'OK') { 'tk-badge-ok' } else { 'tk-badge-info' }
    $diskRows += @"
    <tr>
      <td>$(EscHtml $d.ID)</td>
      <td>$(EscHtml $d.Name)</td>
      <td>$(EscHtml $d.Serial)</td>
      <td>$(EscHtml $d.MediaType)</td>
      <td>$(EscHtml $d.BusType)</td>
      <td>$($d.SizeGB) GB</td>
      <td><span class="$hBadge">$(EscHtml $d.HealthStatus)</span></td>
      <td>$(EscHtml $d.OperationalStatus)</td>
      <td><span class="$sBadge">$(EscHtml $d.SMARTPrediction)</span></td>
      <td><code>$(EscHtml $d.SMARTReason)</code></td>
      <td>$(EscHtml $d.Firmware)</td>
    </tr>
"@
}

# Volume rows
$volumeRows = ""
foreach ($v in $volumeReport) {
    $barBadge  = if ($v.PctUsed -ge 90) { 'tk-badge-err' } elseif ($v.PctUsed -ge 75) { 'tk-badge-warn' } else { 'tk-badge-ok' }
    $vhBadge   = switch ($v.Health) { 'Healthy' { 'tk-badge-ok' } 'Warning' { 'tk-badge-warn' } default { 'tk-badge-err' } }
    $volumeRows += @"
    <tr>
      <td><code>$(EscHtml $v.Drive)</code></td>
      <td>$(EscHtml $v.Label)</td>
      <td>$(EscHtml $v.FileSystem)</td>
      <td>$($v.TotalGB) GB</td>
      <td>$($v.FreeGB) GB</td>
      <td><span class="$barBadge">$($v.PctUsed)%</span></td>
      <td><span class="$vhBadge">$(EscHtml $v.Health)</span></td>
    </tr>
"@
}

$overallBadge = if ($criticalCount -gt 0 -or $smartFailCount -gt 0) {
    "<span class='tk-badge-err'>$($criticalCount + $smartFailCount) critical issue(s)</span>"
} elseif ($warningCount -gt 0) {
    "<span class='tk-badge-warn'>$warningCount warning(s)</span>"
} else {
    "<span class='tk-badge-ok'>All Healthy</span>"
}

$tkCfg     = Get-TKConfig
$orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

$htmlHead = Get-TKHtmlHead `
    -Title     'Disk Health Assessment' `
    -ScriptName 'A.U.G.U.R.' `
    -Subtitle   "$orgPrefix$env:COMPUTERNAME" `
    -MetaItems  ([ordered]@{
        'Machine'   = $env:COMPUTERNAME
        'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
        'Generated' = $reportTimestamp
        'Overall'   = $overallBadge
    }) `
    -NavItems   @('Physical Disks', 'Volumes')

$htmlFoot = Get-TKHtmlFoot -ScriptName 'A.U.G.U.R. v6.0'

$htmlReport = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card ok"><div class="tk-summary-num">$healthyCount</div><div class="tk-summary-lbl">Healthy</div></div>
    <div class="tk-summary-card warn"><div class="tk-summary-num">$warningCount</div><div class="tk-summary-lbl">Warning</div></div>
    <div class="tk-summary-card err"><div class="tk-summary-num">$criticalCount</div><div class="tk-summary-lbl">Critical</div></div>
    <div class="tk-summary-card err"><div class="tk-summary-num">$smartFailCount</div><div class="tk-summary-lbl">SMART Fail</div></div>
  </div>

  <!-- PHYSICAL DISKS -->
  <div class="tk-card">
    <div class="tk-card-header">
      <span class="tk-card-label">Physical Disks ($($diskReport.Count))</span>
    </div>
    $(if ($diskRows) {
      "<table class='tk-table'><thead><tr><th>ID</th><th>Name</th><th>Serial</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Status</th><th>SMART</th><th>SMART Reason</th><th>Firmware</th></tr></thead><tbody>$diskRows</tbody></table>"
    } else {
      "<p class='tk-info-box'>No physical disk data available.</p>"
    })
  </div>

  <div class="tk-divider"></div>

  <!-- VOLUMES -->
  <div class="tk-card">
    <div class="tk-card-header">
      <span class="tk-card-label">Volumes ($($volumeReport.Count))</span>
    </div>
    $(if ($volumeRows) {
      "<table class='tk-table'><thead><tr><th>Drive</th><th>Label</th><th>File System</th><th>Total</th><th>Free</th><th>Usage</th><th>Health</th></tr></thead><tbody>$volumeRows</tbody></table>"
    } else {
      "<p class='tk-info-box'>No volume data available.</p>"
    })
  </div>

"@ + $htmlFoot

try {
    [System.IO.File]::WriteAllText($reportPath, $htmlReport, [System.Text.Encoding]::UTF8)
    Write-Host "  [+] Report saved: $reportPath" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Could not save report: $_" -ForegroundColor $ColorSchema.Error
}

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  A.U.G.U.R. ASSESSMENT COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Physical Disks : $($diskReport.Count) assessed" -ForegroundColor $ColorSchema.Info

if ($criticalCount -gt 0) {
    Write-Host "  Critical       : $criticalCount disk(s) in critical state!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  Critical       : None" -ForegroundColor $ColorSchema.Success
}

if ($smartFailCount -gt 0) {
    Write-Host "  SMART Failure  : $smartFailCount disk(s) predicting failure!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  SMART Failure  : None predicted" -ForegroundColor $ColorSchema.Success
}

if ($warningCount -gt 0) {
    Write-Host "  Warnings       : $warningCount disk(s) in warning state" -ForegroundColor $ColorSchema.Warning
} else {
    Write-Host "  Healthy        : All disks healthy" -ForegroundColor $ColorSchema.Success
}

Write-Host ""
Show-TKReportResult -Path $reportPath -Unattended:$Unattended

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
