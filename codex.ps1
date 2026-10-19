# codex.ps1 - C.O.D.E.X. — Compiles Output Documents into an EXhibit
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
    C.O.D.E.X. — Compiles Output Documents into an EXhibit
    Toolkit Report Index Builder for PowerShell 5.1+

.DESCRIPTION
    Walks the configured log directory, finds every TechnicianToolkit-
    generated HTML report (filename ending in _YYYYMMDD_HHMMSS.html),
    groups them by tool prefix, and emits a single dark-themed HTML
    rollup with relative links to each child report. Where R.I.T.U.A.L.
    composes a fresh recipe run, CODEX answers "what reports already
    exist on disk?" -- useful when a technician has run AUSPEX, HOURGLASS,
    AUGUR, etc. ad-hoc throughout the week and wants one bound volume
    to attach to a ticket.

    Reports are sorted newest-first within each tool group. The summary
    row tallies report count, distinct tool count, total disk size, and
    the date range covered. Files that don't match the timestamp naming
    convention are skipped so random HTMLs in the log directory don't
    pollute the index.

.USAGE
    PS C:\> .\codex.ps1                         # Index every report in the log dir
    PS C:\> .\codex.ps1 -DaysBack 7             # Only the last week of reports
    PS C:\> .\codex.ps1 -LogDir 'D:\Reports'    # Override the scan target
    PS C:\> .\codex.ps1 -Unattended             # Silent: write rollup and exit

.NOTES
    Version : 6.2

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [int]$DaysBack = 0,
    [string]$LogDir
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

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-AdminElevation {
    <#
    .SYNOPSIS
        Stops the tool, with instructions, when the session is not elevated.
    .DESCRIPTION
        This used to relaunch the script elevated with the execution policy
        switched off. A script that re-launches itself as Administrator that
        way is exactly what malware does, and endpoint protection treats it
        so -- and the relaunch dropped every parameter the technician had
        passed. Elevation is now the technician's
        own, visible act: open PowerShell with "Run as administrator", then run
        the tool. The name is kept because the tools call it and the desktop
        app reads it as one of the two admin gates.
    #>
    param([Parameter(Mandatory)][string]$ScriptFile)
    if (-not (Test-IsAdmin)) {
        Write-Host "  $(Split-Path -Leaf $ScriptFile) needs Administrator rights." -ForegroundColor Yellow
        Write-Host "  Open PowerShell with 'Run as administrator' and run it again." -ForegroundColor Yellow
        exit 1
    }
}
#endregion TK-SHARED
Invoke-AdminElevation -ScriptFile $PSCommandPath

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
    Accent   = 'Blue'
}

function Show-CodexBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  C.O.D.E.X. — Compiles Output Documents into an EXhibit" -ForegroundColor Cyan
    Write-Host "  Toolkit Report Index Builder  v6.2" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REPORT DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────

# Reports written before the 6.0 renames carry the old tool name as their
# prefix. Group them under the new name so a machine's history stays in one
# section instead of splitting across two.
$RenamedToolPrefixes = @{
    'RESTORATION'   = 'WHETSTONE'
    'SUTURE'        = 'SOLDER'
    'THRESHOLD'     = 'FATHOM'
    'GARGOYLE'      = 'VIGIL'
    'PYRE'          = 'HOURGLASS'
    'CIPHER'        = 'WYRM'
    'SIGIL'         = 'BASILISK'
    'CITADEL'       = 'SPHINX'
    'ARTIFACT'      = 'PHOENIX'
    'PALADIN'       = 'GRIFFIN'
    'HERALD'        = 'ARGUS'
    'CATACOMB'      = 'MINOTAUR'
    'BEACON'        = 'WISP'
    'SHADE'         = 'EMISSARY'
    'OATH'          = 'LODESTAR'
    'TALISMAN'      = 'ZENITH'
    'RELIQUARY'     = 'ALMANAC'
    'GOLEM'         = 'ORBIT'
    'WRAITH'        = 'ECLIPSE'
    'CONCLAVE'      = 'ASTERISM'
    'GROVE'         = 'CUMULUS'
    'TENDRIL'       = 'ORRERY'
    'RAMPART'       = 'HALO'
    'ARCHIVE'       = 'EMBALM'
    'TETHER'        = 'PHYLACTERY'
}

# Walks the log directory for HTML files whose names end in the toolkit's
# canonical _YYYYMMDD_HHMMSS timestamp. Files outside that pattern are
# skipped on purpose so that browser-saved pages or hand-renamed copies
# don't pollute the rollup. CODEX's own output is excluded so successive
# runs don't index each other recursively.
function Get-ToolkitReportFiles {
    param(
        [Parameter(Mandatory)] [string]$LogDir,
        [int]$DaysBack = 0
    )

    if (-not (Test-Path $LogDir)) { return @() }

    # Trailing-slash normalization keeps the Substring math below correct
    # whether the caller passed `C:\Logs` or `C:\Logs\`.
    $logRoot = $LogDir.TrimEnd('\','/')
    $cutoff  = if ($DaysBack -gt 0) { (Get-Date).AddDays(-$DaysBack) } else { $null }
    $files   = @(Get-ChildItem -Path $logRoot -Filter '*.html' -File -Recurse -ErrorAction SilentlyContinue)

    $rows = foreach ($f in $files) {
        if ($f.Name -match '^CODEX_') { continue }

        # The toolkit-wide convention is `<label>_YYYYMMDD_HHMMSS.html`.
        # The label is the tool acronym, optionally followed by a variant
        # tag (HOURGLASS_battery_report, SPHINX_StaleAccounts, etc.). Capture
        # all three groups up front -- $matches gets clobbered by any
        # later -match call below, so we can't rely on it surviving.
        if ($f.BaseName -notmatch '^(?<label>.+)_(?<date>\d{8})_(?<time>\d{6})$') { continue }
        $label   = $matches['label']
        $dateStr = $matches['date']
        $timeStr = $matches['time']

        $tool    = ($label -split '_', 2)[0].ToUpper()
        if ($RenamedToolPrefixes.ContainsKey($tool)) { $tool = $RenamedToolPrefixes[$tool] }
        $variant = if ($label -match '^[^_]+_(?<v>.+)$') { $matches['v'] } else { '' }

        $ts = $null
        try {
            $ts = [datetime]::ParseExact("${dateStr}_${timeStr}", 'yyyyMMdd_HHmmss', $null)
        } catch {
            $ts = $f.LastWriteTime
        }

        if ($cutoff -and $ts -lt $cutoff) { continue }

        # Relative path from the log directory keeps the rollup portable
        # when the whole folder is zipped or moved to a ticket attachment.
        $rel = if ($f.FullName.StartsWith($logRoot, [StringComparison]::OrdinalIgnoreCase)) {
            $f.FullName.Substring($logRoot.Length).TrimStart('\','/')
        } else {
            $f.FullName
        }

        [PSCustomObject]@{
            FullPath  = $f.FullName
            RelPath   = $rel
            Name      = $f.Name
            Tool      = $tool
            Variant   = $variant
            Timestamp = $ts
            SizeBytes = $f.Length
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-CodexHtml {
    param(
        [Parameter(Mandatory)] [array]$Reports,
        [Parameter(Mandatory)] [string]$LogDir,
        [int]$DaysBack = 0
    )

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalCount = $Reports.Count
    $totalSize  = ($Reports | Measure-Object -Property SizeBytes -Sum).Sum
    if ($null -eq $totalSize) { $totalSize = 0 }

    $tools = @($Reports | Group-Object Tool | Sort-Object Name)
    $toolCount = $tools.Count

    $oldest = if ($totalCount -gt 0) { ($Reports | Measure-Object -Property Timestamp -Minimum).Minimum } else { $null }
    $newest = if ($totalCount -gt 0) { ($Reports | Measure-Object -Property Timestamp -Maximum).Maximum } else { $null }

    $weekCutoff = (Get-Date).AddDays(-7)
    $lastWeek   = @($Reports | Where-Object { $_.Timestamp -ge $weekCutoff }).Count

    $rangeText = if ($DaysBack -gt 0) { "Last $DaysBack day(s)" } else { 'All time' }

    # Per-tool sections
    $toolSections = [System.Text.StringBuilder]::new()
    if ($tools.Count -eq 0) {
        [void]$toolSections.Append(@"

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">No reports found</span></div>
    <div class="tk-card"><div class="tk-info-box">Log directory: <span class="tk-mono">$(EscHtml $LogDir)</span><br>Either no reports exist in that path, or none match the <span class="tk-mono">_YYYYMMDD_HHMMSS.html</span> naming convention.$(if ($DaysBack -gt 0) { " The <span class=""tk-mono"">-DaysBack $DaysBack</span> filter may also be excluding older reports." } else { '' })</div></div>
  </div>
"@)
    }

    $navItems = @('Overview')

    foreach ($g in $tools) {
        $navItems += $g.Name
        $entries  = @($g.Group | Sort-Object Timestamp -Descending)
        $newestEntry = ($entries | Select-Object -First 1).Timestamp
        $oldestEntry = ($entries | Select-Object -Last 1).Timestamp
        $groupSize   = ($entries | Measure-Object -Property SizeBytes -Sum).Sum
        if ($null -eq $groupSize) { $groupSize = 0 }

        $rowsSb = [System.Text.StringBuilder]::new()
        foreach ($r in $entries) {
            $variant = if ([string]::IsNullOrWhiteSpace($r.Variant)) { '<span class="tk-badge-info">main</span>' } else { "<span class=""tk-badge-blue"">$(EscHtml $r.Variant)</span>" }
            $href    = EscHtml ($r.RelPath -replace '\\','/')
            $size    = Format-Bytes $r.SizeBytes
            $when    = $r.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
            [void]$rowsSb.Append("<tr><td class=""tk-mono"">$when</td><td>$variant</td><td><a href=""$href"">$(EscHtml $r.Name)</a></td><td>$size</td></tr>`n")
        }

        $latestStr = $newestEntry.ToString('yyyy-MM-dd HH:mm')
        $earliestStr = $oldestEntry.ToString('yyyy-MM-dd HH:mm')

        [void]$toolSections.Append(@"

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">$(EscHtml $g.Name)</span><span class="tk-section-num">$($entries.Count) report(s)</span></div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Latest</span> $latestStr&nbsp;&nbsp;
        <span class="tk-info-label">Earliest</span> $earliestStr&nbsp;&nbsp;
        <span class="tk-info-label">Total size</span> $(Format-Bytes $groupSize)
      </div>
      <table class="tk-table">
        <thead><tr><th>Generated</th><th>Variant</th><th>File</th><th>Size</th></tr></thead>
        <tbody>$($rowsSb.ToString())</tbody>
      </table>
    </div>
  </div>
"@)
    }

    $newestCard = if ($newest) { $newest.ToString('yyyy-MM-dd') } else { 'n/a' }
    $oldestCard = if ($oldest) { $oldest.ToString('yyyy-MM-dd') } else { 'n/a' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.O.D.E.X. Toolkit Report Index' `
        -ScriptName 'C.O.D.E.X.' `
        -Subtitle   "${orgPrefix}Toolkit Report Index -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'    = $machine
            'Log Dir'    = $LogDir
            'Generated'  = $reportDate
            'Range'      = $rangeText
            'Reports'    = $totalCount
            'Tools'      = $toolCount
        }) `
        -NavItems   $navItems

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'C.O.D.E.X. v6.2'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalCount</div><div class="tk-summary-lbl">Reports</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$toolCount</div><div class="tk-summary-lbl">Tools</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$lastWeek</div><div class="tk-summary-lbl">Last 7 Days</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(Format-Bytes $totalSize)</div><div class="tk-summary-lbl">Total Size</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$newestCard</div><div class="tk-summary-lbl">Newest</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$oldestCard</div><div class="tk-summary-lbl">Oldest</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Overview</span><span class="tk-section-num">$rangeText</span></div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Log directory</span> <span class="tk-mono">$(EscHtml $LogDir)</span><br>
        <span class="tk-info-label">Filter</span> $(EscHtml $rangeText)$(if ($DaysBack -gt 0) { " (cutoff $((Get-Date).AddDays(-$DaysBack).ToString('yyyy-MM-dd HH:mm')))" } else { '' })<br>
        <span class="tk-info-label">Naming pattern</span> <span class="tk-mono">&lt;TOOL&gt;_YYYYMMDD_HHMMSS.html</span> (files outside this pattern are skipped)
      </div>
    </div>
  </div>
$($toolSections.ToString())
"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-CodexBanner

$resolvedLogDir = if (-not [string]::IsNullOrWhiteSpace($LogDir)) {
    $LogDir
} else {
    Resolve-LogDirectory -FallbackPath $ScriptPath
}

Write-Section "REPORT DISCOVERY"
Write-Step "Scanning: $resolvedLogDir"
if ($DaysBack -gt 0) { Write-Info "  Filter: last $DaysBack day(s)" }

if (-not (Test-Path $resolvedLogDir)) {
    Write-Fail "Log directory does not exist: $resolvedLogDir"
    if (-not $Unattended) { Read-Host "  Press Enter to exit" }
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

$reports = Get-ToolkitReportFiles -LogDir $resolvedLogDir -DaysBack $DaysBack

if ($reports.Count -eq 0) {
    Write-Warn "No matching reports found."
} else {
    $byTool = $reports | Group-Object Tool | Sort-Object Name
    Write-Ok ("Found {0} report(s) across {1} tool(s)." -f $reports.Count, $byTool.Count)
    foreach ($g in $byTool) {
        $latest = ($g.Group | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp.ToString('yyyy-MM-dd HH:mm')
        Write-Host ("    {0,-22} {1,3} report(s)   newest {2}" -f $g.Name, $g.Count, $latest) -ForegroundColor $C.Info
    }
}
Write-Host ""

Write-Step "Generating rollup HTML..."
$html      = Build-CodexHtml -Reports $reports -LogDir $resolvedLogDir -DaysBack $DaysBack
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path $resolvedLogDir "CODEX_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended -Label 'Index'
} catch {
    Write-Fail "Could not save index: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
