# auspex.ps1 - A.U.S.P.E.X. — Audits, Uncovers, Surveys Performance, Events & eXceptions
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
    A.U.S.P.E.X. — Audits, Uncovers, Surveys Performance, Events & eXceptions
    System Diagnostic & Health Assessment Tool for PowerShell 5.1+

.DESCRIPTION
    Audits and reports on the current state of a Windows machine. Collects
    hardware specs, OS info, network config, uptime, pending Windows Updates,
    installed software, and recent event log errors — then exports a
    dark-themed HTML report with color-coded indicators to the Desktop.

.USAGE
    PS C:\> .\auspex.ps1                    # Must be run as Administrator
    PS C:\> .\auspex.ps1 -Unattended        # Silent mode — no prompts, no banner

.NOTES
    Version : 6.0

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
.tk-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-family:'Consolas','Courier New',monospace;font-size:11px;font-weight:600;letter-spacing:.03em;white-space:nowrap}
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
        TeamsWebhook  = ''
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
# BANNER DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-AuspexBanner {
    Write-Host @"

   █████╗ ██╗   ██╗███████╗██████╗ ███████╗██╗  ██╗
  ██╔══██╗██║   ██║██╔════╝██╔══██╗██╔════╝╚██╗██╔╝
  ███████║██║   ██║███████╗██████╔╝█████╗   ╚███╔╝
  ██╔══██║██║   ██║╚════██║██╔═══╝ ██╔══╝   ██╔██╗
  ██║  ██║╚██████╔╝███████║██║     ███████╗██╔╝ ██╗
  ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝     ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    A.U.S.P.E.X. — Audits, Uncovers, Surveys Performance, Events & eXceptions" -ForegroundColor Cyan
    Write-Host "    System Diagnostic & Health Assessment Tool" -ForegroundColor Cyan
    Write-Host ""
}

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
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

# Folder where the HTML report will be saved.
# Override via config.json LogDirectory, or set $ReportOutputPath manually below.
$ReportOutputPath = Resolve-LogDirectory -FallbackPath $ScriptPath

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA DEFINITION
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
# DISPLAY BANNER
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-AuspexBanner }

$reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$reportFilename  = "AUSPEX_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath      = "$ReportOutputPath\$reportFilename"

Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "     SYSTEM DIAGNOSTIC REPORT" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Run As    : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Timestamp : $reportTimestamp" -ForegroundColor $ColorSchema.Info
Write-Host "  Report    : $reportPath" -ForegroundColor $ColorSchema.Info
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION STORAGE
# ─────────────────────────────────────────────────────────────────────────────

$reportData = [ordered]@{}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: HARDWARE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[1/10] Collecting Hardware Info..." -ForegroundColor $ColorSchema.Progress

try {
    $cs        = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios      = Get-CimInstance -ClassName Win32_BIOS
    $cpu       = Get-CimInstance -ClassName Win32_Processor
    $ramBytes  = ($cs.TotalPhysicalMemory)
    $ramGB     = [math]::Round($ramBytes / 1GB, 2)
    $disks     = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"

    Write-Host "    Manufacturer : $($cs.Manufacturer)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Model        : $($cs.Model)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Serial       : $($bios.SerialNumber)" -ForegroundColor $ColorSchema.Info
    Write-Host "    CPU          : $($cpu.Name)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Cores/Threads: $($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) threads" -ForegroundColor $ColorSchema.Info
    Write-Host "    RAM          : $ramGB GB" -ForegroundColor $ColorSchema.Info

    $diskSummary = [System.Collections.Generic.List[object]]::new()
    foreach ($disk in $disks) {
        $totalGB = [math]::Round($disk.Size / 1GB, 1)
        $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 1)
        $usedGB  = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 1)
        $pct     = if ($disk.Size -gt 0) { [math]::Round(($disk.Size - $disk.FreeSpace) / $disk.Size * 100, 1) } else { 0 }

        $color = if ($pct -ge 90) { $ColorSchema.Error } elseif ($pct -ge 75) { $ColorSchema.Warning } else { $ColorSchema.Info }
        Write-Host "    Disk $($disk.DeviceID)      : $usedGB GB used / $totalGB GB total ($pct% full)" -ForegroundColor $color

        $diskSummary.Add([PSCustomObject]@{
            Drive   = $disk.DeviceID
            Label   = $disk.VolumeName
            TotalGB = $totalGB
            UsedGB  = $usedGB
            FreeGB  = $freeGB
            PctUsed = $pct
        })
    }

    $reportData['Hardware'] = [PSCustomObject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $cs.Model
        Serial       = $bios.SerialNumber
        CPU          = $cpu.Name
        Cores        = $cpu.NumberOfCores
        Threads      = $cpu.NumberOfLogicalProcessors
        RAMGB        = $ramGB
        Disks        = $diskSummary
    }

    Write-Host "[+] Hardware info collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting hardware info: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: OPERATING SYSTEM
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[2/10] Collecting OS Info..." -ForegroundColor $ColorSchema.Progress

try {
    $os          = Get-CimInstance -ClassName Win32_OperatingSystem
    $osVersion   = $os.Caption
    $osBuild     = $os.BuildNumber
    $osArch      = $os.OSArchitecture
    $osInstall   = $os.InstallDate

    # Activation status via slmgr
    $licenseStatus = "Unknown"
    try {
        $slmgr = & cscript //nologo "$env:SystemRoot\System32\slmgr.vbs" /dli 2>&1
        $licenseLine = $slmgr | Where-Object { $_ -match "License Status" }
        if ($licenseLine) {
            $licenseStatus = ($licenseLine -replace "License Status:\s*", "").Trim()
        }
    }
    catch {
        $licenseStatus = "Could not query"
    }

    Write-Host "    OS       : $osVersion" -ForegroundColor $ColorSchema.Info
    Write-Host "    Build    : $osBuild  ($osArch)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Installed: $(Get-Date $osInstall -Format 'yyyy-MM-dd')" -ForegroundColor $ColorSchema.Info

    $activationColor = if ($licenseStatus -match "Licensed") { $ColorSchema.Success } else { $ColorSchema.Warning }
    Write-Host "    Activation: $licenseStatus" -ForegroundColor $activationColor

    $reportData['OS'] = [PSCustomObject]@{
        Caption      = $osVersion
        Build        = $osBuild
        Architecture = $osArch
        InstallDate  = Get-Date $osInstall -Format 'yyyy-MM-dd'
        Activation   = $licenseStatus
    }

    Write-Host "[+] OS info collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting OS info: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: NETWORK CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[3/10] Collecting Network Configuration..." -ForegroundColor $ColorSchema.Progress

try {
    $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
    $netSummary = [System.Collections.Generic.List[object]]::new()

    foreach ($adapter in $adapters) {
        $ip      = ($adapter.IPAddress      | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ', '
        $dns     = ($adapter.DNSServerSearchOrder) -join ', '
        $gateway = ($adapter.DefaultIPGateway) -join ', '

        Write-Host "    Adapter  : $($adapter.Description)" -ForegroundColor $ColorSchema.Info
        Write-Host "    IP       : $ip" -ForegroundColor $ColorSchema.Info
        Write-Host "    MAC      : $($adapter.MACAddress)" -ForegroundColor $ColorSchema.Info
        Write-Host "    Gateway  : $gateway" -ForegroundColor $ColorSchema.Info
        Write-Host "    DNS      : $dns" -ForegroundColor $ColorSchema.Info
        Write-Host ""

        $netSummary.Add([PSCustomObject]@{
            Adapter = $adapter.Description
            IP      = $ip
            MAC     = $adapter.MACAddress
            Gateway = $gateway
            DNS     = $dns
        })
    }

    $reportData['Network'] = $netSummary
    Write-Host "[+] Network config collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting network config: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: SYSTEM HEALTH & UPTIME
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[4/10] Collecting System Health..." -ForegroundColor $ColorSchema.Progress

try {
    $os        = Get-CimInstance -ClassName Win32_OperatingSystem
    $lastBoot  = $os.LastBootUpTime
    $uptime    = (Get-Date) - $lastBoot
    $uptimeStr = "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes

    Write-Host "    Last Boot : $(Get-Date $lastBoot -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $ColorSchema.Info
    Write-Host "    Uptime    : $uptimeStr" -ForegroundColor $ColorSchema.Info

    # Battery (laptops only)
    $batteryInfo = $null
    $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    if ($battery) {
        $charge = $battery.EstimatedChargeRemaining
        $status = switch ($battery.BatteryStatus) {
            1 { "Discharging" } 2 { "AC - Fully Charged" } 3 { "Fully Charged" }
            4 { "Low" } 5 { "Critical" } 6 { "Charging" } 7 { "Charging/High" }
            8 { "Charging/Low" } 9 { "Charging/Critical" } default { "Unknown" }
        }
        $battColor = if ($charge -lt 20) { $ColorSchema.Error } elseif ($charge -lt 40) { $ColorSchema.Warning } else { $ColorSchema.Success }
        Write-Host "    Battery   : $charge% ($status)" -ForegroundColor $battColor
        $batteryInfo = [PSCustomObject]@{ Charge = $charge; Status = $status }
    }
    else {
        Write-Host "    Battery   : N/A (desktop or not detected)" -ForegroundColor $ColorSchema.Info
    }

    $reportData['Health'] = [PSCustomObject]@{
        LastBoot = Get-Date $lastBoot -Format 'yyyy-MM-dd HH:mm:ss'
        Uptime   = $uptimeStr
        Battery  = $batteryInfo
    }

    Write-Host "[+] System health collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting system health: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5: STORAGE & RAID HEALTH
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[5/10] Collecting Storage & RAID Health..." -ForegroundColor $ColorSchema.Progress

$physicalDiskSummary = [System.Collections.Generic.List[object]]::new()
$virtualDiskSummary  = [System.Collections.Generic.List[object]]::new()

try {
    $physicalDisks = Get-PhysicalDisk -ErrorAction Stop

    foreach ($pd in $physicalDisks) {
        $sizeGB = if ($pd.Size -gt 0) { [math]::Round($pd.Size / 1GB, 1) } else { 0 }
        $healthColor = switch ($pd.HealthStatus) {
            'Healthy' { $ColorSchema.Success }
            'Warning' { $ColorSchema.Warning }
            default   { $ColorSchema.Error   }
        }
        Write-Host "    [$($pd.DeviceId)] $($pd.FriendlyName) | $($pd.MediaType) | $sizeGB GB | Health: $($pd.HealthStatus) | Status: $($pd.OperationalStatus)" -ForegroundColor $healthColor

        $physicalDiskSummary.Add([PSCustomObject]@{
            ID                = $pd.DeviceId
            Name              = $pd.FriendlyName
            MediaType         = $pd.MediaType
            BusType           = $pd.BusType
            SizeGB            = $sizeGB
            HealthStatus      = $pd.HealthStatus
            OperationalStatus = $pd.OperationalStatus
        })
    }

    # Storage Spaces virtual disks (software RAID)
    $virtualDisks = Get-VirtualDisk -ErrorAction SilentlyContinue
    if ($virtualDisks) {
        Write-Host "    Storage Spaces (virtual disks):" -ForegroundColor $ColorSchema.Info
        foreach ($vd in $virtualDisks) {
            $vHealthColor = switch ($vd.HealthStatus) {
                'Healthy' { $ColorSchema.Success }
                'Warning' { $ColorSchema.Warning }
                default   { $ColorSchema.Error   }
            }
            $vSizeGB = if ($vd.Size -gt 0) { [math]::Round($vd.Size / 1GB, 1) } else { 0 }
            Write-Host "      VDisk: $($vd.FriendlyName) | $($vd.ResiliencySettingName) | $vSizeGB GB | Health: $($vd.HealthStatus) / Op: $($vd.OperationalStatus)" -ForegroundColor $vHealthColor

            $virtualDiskSummary.Add([PSCustomObject]@{
                Name              = $vd.FriendlyName
                ResiliencyType    = $vd.ResiliencySettingName
                SizeGB            = $vSizeGB
                HealthStatus      = $vd.HealthStatus
                OperationalStatus = $vd.OperationalStatus
            })
        }
    }
    else {
        Write-Host "    No Storage Spaces virtual disks detected" -ForegroundColor $ColorSchema.Info
    }

    # Hardware RAID controller detection via WMI
    $raidControllers = [System.Collections.Generic.List[object]]::new()
    try {
        $scsiControllers = Get-CimInstance -ClassName Win32_SCSIController -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'RAID|MegaRAID|PERC|Smart Array|LSI|Adaptec|Avago|Broadcom|RST|Intel.*RAID|Areca|3ware' }

        foreach ($ctrl in $scsiControllers) {
            $ctrlColor = if ($ctrl.Status -eq 'OK') { $ColorSchema.Success } else { $ColorSchema.Warning }
            Write-Host "    RAID Controller : $($ctrl.Name) [$($ctrl.Status)]" -ForegroundColor $ctrlColor
            $raidControllers.Add([PSCustomObject]@{
                Name         = $ctrl.Name
                Manufacturer = if ($ctrl.Manufacturer) { $ctrl.Manufacturer } else { 'N/A' }
                Status       = $ctrl.Status
                DriverName   = if ($ctrl.DriverName) { $ctrl.DriverName } else { 'N/A' }
            })
        }

        if ($raidControllers.Count -eq 0) {
            Write-Host "    No dedicated hardware RAID controllers detected" -ForegroundColor $ColorSchema.Info
        }
    }
    catch {
        Write-Host "    Could not query RAID controllers: $_" -ForegroundColor $ColorSchema.Warning
    }

    # Vendor CLI tools — capture detailed RAID config if available
    $raidVendorOutput = [System.Collections.Generic.List[object]]::new()
    $cliCandidates = @(
        @{ Name = 'StorCLI';   Exe = 'StorCLI64.exe'; Dirs = @('C:\Windows\System32','C:\Program Files\MegaRAID\StorCLI','C:\Program Files (x86)\MegaRAID\StorCLI');         Args = @('/call','show') },
        @{ Name = 'PERCCLI';   Exe = 'perccli64.exe'; Dirs = @('C:\Windows\System32','C:\Program Files\PERCCLI','C:\Program Files (x86)\PERCCLI');                            Args = @('/call','show') },
        @{ Name = 'HP SSACLI'; Exe = 'ssacli.exe';    Dirs = @('C:\Program Files\Smart Storage Administrator\ssacli\bin','C:\Program Files (x86)\Compaq\Hpacucli\Bin');       Args = @('ctrl','all','show','config') }
    )

    foreach ($cli in $cliCandidates) {
        $exePath = $cli.Dirs | ForEach-Object { Join-Path $_ $cli.Exe } | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($exePath) {
            try {
                Write-Host "    Found $($cli.Name): $exePath" -ForegroundColor $ColorSchema.Info
                $rawOutput = & "$exePath" @($cli.Args) 2>&1 | Out-String
                $raidVendorOutput.Add([PSCustomObject]@{ Tool = $cli.Name; Output = $rawOutput.Trim() })
                Write-Host "    [+] $($cli.Name) output captured" -ForegroundColor $ColorSchema.Success
            }
            catch {
                Write-Host "    $($cli.Name) found but failed to run: $_" -ForegroundColor $ColorSchema.Warning
            }
        }
    }

    $reportData['Storage'] = [PSCustomObject]@{
        PhysicalDisks    = $physicalDiskSummary
        VirtualDisks     = $virtualDiskSummary
        RaidControllers  = $raidControllers
        RaidVendorOutput = $raidVendorOutput
    }

    Write-Host "[+] Storage health collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting storage health: $_" -ForegroundColor $ColorSchema.Error
    $reportData['Storage'] = [PSCustomObject]@{
        PhysicalDisks    = @()
        VirtualDisks     = @()
        RaidControllers  = @()
        RaidVendorOutput = @()
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: PENDING WINDOWS UPDATES
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[6/10] Scanning for Pending Updates..." -ForegroundColor $ColorSchema.Progress
Write-Host "    This may take a moment..." -ForegroundColor $ColorSchema.Info

$pendingUpdates = [System.Collections.Generic.List[object]]::new()
try {
    $updateSession  = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult   = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

    if ($searchResult.Updates.Count -eq 0) {
        Write-Host "    [+] No pending updates found" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    [!!] $($searchResult.Updates.Count) pending update(s) found:" -ForegroundColor $ColorSchema.Warning
        foreach ($update in $searchResult.Updates) {
            Write-Host "      * $($update.Title)" -ForegroundColor $ColorSchema.Warning
            $pendingUpdates.Add([PSCustomObject]@{
                Title    = $update.Title
                Severity = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { "N/A" }
                KB       = ($update.KBArticleIDs -join ', ')
            })
        }
    }

    $reportData['Updates'] = $pendingUpdates
    Write-Host "[+] Update scan complete" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error scanning for updates: $_" -ForegroundColor $ColorSchema.Error
    $reportData['Updates'] = @()
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: INSTALLED SOFTWARE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[7/10] Collecting Installed Software..." -ForegroundColor $ColorSchema.Progress

$installedApps = [System.Collections.Generic.List[object]]::new()
try {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $regPaths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -ne '' } |
            ForEach-Object {
                $installedApps.Add([PSCustomObject]@{
                    Name      = $_.DisplayName
                    Version   = if ($_.DisplayVersion) { $_.DisplayVersion } else { "N/A" }
                    Publisher = if ($_.Publisher)       { $_.Publisher }       else { "N/A" }
                    InstallDate = if ($_.InstallDate)   { $_.InstallDate }     else { "N/A" }
                })
            }
    }

    # Deduplicate by name and sort
    $installedApps = $installedApps | Sort-Object Name -Unique

    Write-Host "    Found $($installedApps.Count) installed application(s)" -ForegroundColor $ColorSchema.Info
    $reportData['Software'] = $installedApps
    Write-Host "[+] Software list collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting software list: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7: RECENT EVENT LOG ERRORS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[8/10] Scanning Event Logs (last 24 hours)..." -ForegroundColor $ColorSchema.Progress

$eventSummary = @()
try {
    $since  = (Get-Date).AddHours(-24)
    $levels = @(1, 2)  # 1 = Critical, 2 = Error

    $events = Get-WinEvent -FilterHashtable @{
        LogName   = 'System', 'Application'
        Level     = $levels
        StartTime = $since
    } -ErrorAction SilentlyContinue -MaxEvents 50

    if ($null -eq $events -or $events.Count -eq 0) {
        Write-Host "    [+] No critical/error events in the last 24 hours" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    [!!] $($events.Count) error/critical event(s) found in the last 24 hours" -ForegroundColor $ColorSchema.Warning
        $events | Select-Object -First 10 | ForEach-Object {
            $lvl = if ($_.Level -eq 1) { "CRITICAL" } else { "ERROR" }
            Write-Host "      [$lvl] $(Get-Date $_.TimeCreated -Format 'HH:mm:ss') | $($_.ProviderName) — $($_.Message.Split([Environment]::NewLine)[0])" -ForegroundColor $ColorSchema.Warning
        }
        if ($events.Count -gt 10) {
            Write-Host "      ... and $($events.Count - 10) more (see full report)" -ForegroundColor $ColorSchema.Info
        }

        $eventSummary = $events | ForEach-Object {
            [PSCustomObject]@{
                Time     = Get-Date $_.TimeCreated -Format 'yyyy-MM-dd HH:mm:ss'
                Level    = if ($_.Level -eq 1) { "Critical" } else { "Error" }
                Source   = $_.ProviderName
                Log      = $_.LogName
                Message  = $_.Message.Split([Environment]::NewLine)[0]
                EventID  = $_.Id
            }
        }
    }

    $reportData['Events'] = $eventSummary
    Write-Host "[+] Event log scan complete" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error reading event logs: $_" -ForegroundColor $ColorSchema.Error
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 9: SCHEDULED TASKS (NON-MICROSOFT)
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[9/10] Collecting Scheduled Tasks..." -ForegroundColor $ColorSchema.Progress

$scheduledTaskSummary = [System.Collections.Generic.List[object]]::new()
try {
    $tasks = Get-ScheduledTask -ErrorAction Stop |
        Where-Object { $_.TaskPath -notmatch '^\\Microsoft\\' }

    foreach ($task in $tasks) {
        $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue
        $actions = ($task.Actions | ForEach-Object {
            if ($_.Execute) { "$($_.Execute) $($_.Arguments)".Trim() }
        }) -join '; '

        $lastRun = if ($info -and $info.LastRunTime -and $info.LastRunTime -gt [DateTime]::MinValue) {
            $info.LastRunTime.ToString('yyyy-MM-dd HH:mm')
        } else { 'Never' }

        $nextRun = if ($info -and $info.NextRunTime -and $info.NextRunTime -gt [DateTime]::MinValue) {
            $info.NextRunTime.ToString('yyyy-MM-dd HH:mm')
        } else { 'N/A' }

        $scheduledTaskSummary.Add([PSCustomObject]@{
            Name       = $task.TaskName
            Path       = $task.TaskPath
            State      = $task.State
            LastRun    = $lastRun
            NextRun    = $nextRun
            LastResult = if ($info) { "0x{0:X8}" -f $info.LastTaskResult } else { 'N/A' }
            Actions    = $actions
        })
    }

    Write-Host "    Found $($scheduledTaskSummary.Count) non-Microsoft scheduled task(s)" -ForegroundColor $ColorSchema.Info
    $reportData['ScheduledTasks'] = $scheduledTaskSummary
    Write-Host "[+] Scheduled tasks collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting scheduled tasks: $_" -ForegroundColor $ColorSchema.Error
    $reportData['ScheduledTasks'] = @()
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 10: SECURITY / AV STATUS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[10/10] Collecting Security & AV Status..." -ForegroundColor $ColorSchema.Progress

$avProducts = [System.Collections.Generic.List[object]]::new()
try {
    $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defender) {
        $defAgeDays = if ($defender.AntivirusSignatureLastUpdated) {
            [int](New-TimeSpan -Start $defender.AntivirusSignatureLastUpdated -End (Get-Date)).TotalDays
        } else { $null }

        $lastScan = if ($defender.QuickScanEndTime -and $defender.QuickScanEndTime -ne [datetime]::MinValue) {
            Get-Date $defender.QuickScanEndTime -Format 'yyyy-MM-dd HH:mm'
        } else { 'Never' }

        $avProducts.Add([PSCustomObject]@{
            Product            = 'Windows Defender'
            RealTimeProtection = if ($defender.RealTimeProtectionEnabled) { 'On' } else { 'Off' }
            DefinitionAge      = if ($null -ne $defAgeDays) { "$defAgeDays day(s)" } else { 'Unknown' }
            LastQuickScan      = $lastScan
            ServiceEnabled     = if ($defender.AMServiceEnabled) { 'Yes' } else { 'No' }
        })
        Write-Host "[+] Windows Defender status collected" -ForegroundColor $ColorSchema.Success
    }
}
catch {
    Write-Host "[-] Could not read Windows Defender status: $_" -ForegroundColor $ColorSchema.Warning
}

try {
    $thirdParty = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction SilentlyContinue |
                  Where-Object { $_.displayName -notmatch 'Windows Defender|Microsoft Defender' }
    foreach ($av in $thirdParty) {
        $avProducts.Add([PSCustomObject]@{
            Product            = $av.displayName
            RealTimeProtection = 'N/A'
            DefinitionAge      = 'N/A'
            LastQuickScan      = 'N/A'
            ServiceEnabled     = 'Registered'
        })
    }
}
catch {
    # SecurityCenter2 namespace is absent on Server SKUs and some Core editions.
}

$reportData['Security'] = $avProducts

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# GENERATE HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "Generating HTML report..." -ForegroundColor $ColorSchema.Progress

function ConvertTo-HtmlTable {
    param([array]$Objects, [string]$EmptyMessage = "No data available.")
    if (-not $Objects -or $Objects.Count -eq 0) {
        return "<p class='tk-info-box'>$EmptyMessage</p>"
    }
    $headers = $Objects[0].PSObject.Properties.Name
    $html  = "<table class='tk-table'><thead><tr>"
    $html += ($headers | ForEach-Object { "<th>$_</th>" }) -join ""
    $html += "</tr></thead><tbody>"
    foreach ($row in $Objects) {
        $html += "<tr>"
        foreach ($h in $headers) {
            $val = $row.$h
            $escaped = "$val" -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
            $html += "<td>$escaped</td>"
        }
        $html += "</tr>"
    }
    $html += "</tbody></table>"
    return $html
}

# Storage health badge
$storageUnhealthy  = ($reportData['Storage'].PhysicalDisks   | Where-Object { $_.HealthStatus -ne 'Healthy' }).Count
$raidCtrlUnhealthy = ($reportData['Storage'].RaidControllers  | Where-Object { $_.Status       -ne 'OK'      }).Count
$storageBadge = if ($storageUnhealthy -eq 0 -and $raidCtrlUnhealthy -eq 0) {
    "<span class='tk-badge-ok'>Healthy</span>"
} elseif ($raidCtrlUnhealthy -gt 0) {
    "<span class='tk-badge-err'>$raidCtrlUnhealthy controller issue(s)</span>"
} else {
    "<span class='tk-badge-err'>$storageUnhealthy degraded</span>"
}

# Physical disk rows
$physDiskRows = ""
foreach ($pd in $reportData['Storage'].PhysicalDisks) {
    $hBadge = switch ($pd.HealthStatus) {
        'Healthy' { "tk-badge-ok" } 'Warning' { "tk-badge-warn" } default { "tk-badge-err" }
    }
    $physDiskRows += "<tr><td>$($pd.ID)</td><td>$($pd.Name)</td><td>$($pd.MediaType)</td><td>$($pd.BusType)</td><td>$($pd.SizeGB) GB</td>"
    $physDiskRows += "<td><span class='$hBadge'>$($pd.HealthStatus)</span></td><td>$($pd.OperationalStatus)</td></tr>"
}

# Virtual disk rows
$virtDiskRows = ""
foreach ($vd in $reportData['Storage'].VirtualDisks) {
    $hBadge = switch ($vd.HealthStatus) {
        'Healthy' { "tk-badge-ok" } 'Warning' { "tk-badge-warn" } default { "tk-badge-err" }
    }
    $virtDiskRows += "<tr><td>$($vd.Name)</td><td>$($vd.ResiliencyType)</td><td>$($vd.SizeGB) GB</td>"
    $virtDiskRows += "<td><span class='$hBadge'>$($vd.HealthStatus)</span></td><td>$($vd.OperationalStatus)</td></tr>"
}

# RAID controller rows
$raidCtrlRows = ""
foreach ($rc in $reportData['Storage'].RaidControllers) {
    $rcBadge = if ($rc.Status -eq 'OK') { "tk-badge-ok" } else { "tk-badge-warn" }
    $raidCtrlRows += "<tr><td>$($rc.Name)</td><td>$($rc.Manufacturer)</td>"
    $raidCtrlRows += "<td><span class='$rcBadge'>$($rc.Status)</span></td><td>$($rc.DriverName)</td></tr>"
}

# Vendor CLI output blocks
$raidVendorHtml = ""
foreach ($rv in $reportData['Storage'].RaidVendorOutput) {
    $escaped = $rv.Output -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    $raidVendorHtml += "<div class='tk-info-box' style='margin:14px 0 6px;'><span class='tk-info-label'>$($rv.Tool) Output</span></div>"
    $raidVendorHtml += "<pre class='tk-mono' style='padding:12px;font-size:0.78em;overflow-x:auto;white-space:pre-wrap;'>$escaped</pre>"
}

# Disk rows for hardware section
$diskRows = ""
foreach ($d in $reportData['Hardware'].Disks) {
    $barClass = if ($d.PctUsed -ge 90) { "err" } elseif ($d.PctUsed -ge 75) { "warn" } else { "ok" }
    $diskRows += @"
        <tr>
            <td>$($d.Drive)</td>
            <td>$($d.Label)</td>
            <td>$($d.TotalGB) GB</td>
            <td>$($d.UsedGB) GB</td>
            <td>$($d.FreeGB) GB</td>
            <td>
                <div class='tk-progress-wrap'>
                    <div class='tk-progress-bar $barClass' style='width:$($d.PctUsed)%;'></div>
                </div>
                $($d.PctUsed)%
            </td>
        </tr>
"@
}

# Network rows
$netRows = ""
foreach ($n in $reportData['Network']) {
    $netRows += "<tr><td>$($n.Adapter)</td><td>$($n.IP)</td><td>$($n.MAC)</td><td>$($n.Gateway)</td><td>$($n.DNS)</td></tr>"
}

# Battery
$hw = $reportData['Hardware']
$os = $reportData['OS']
$health = $reportData['Health']
# Updates badge
$updateCount = $reportData['Updates'].Count
$updateBadge = if ($updateCount -eq 0) {
    "<span class='tk-badge-ok'>Up to date</span>"
} else {
    "<span class='tk-badge-warn'>$updateCount pending</span>"
}
$updatesTable = if ($updateCount -gt 0) {
    ConvertTo-HtmlTable -Objects $reportData['Updates'] -EmptyMessage "No pending updates."
} else {
    "<p class='empty'>System is fully up to date.</p>"
}

# Events badge
$eventCount = $reportData['Events'].Count
$eventBadge = if ($eventCount -eq 0) {
    "<span class='tk-badge-ok'>Clean</span>"
} else {
    "<span class='tk-badge-warn'>$eventCount events</span>"
}

$softwareTable = ConvertTo-HtmlTable -Objects $reportData['Software']        -EmptyMessage "No software found."
$eventsTable   = ConvertTo-HtmlTable -Objects $reportData['Events']          -EmptyMessage "No critical/error events in the last 24 hours."
$tasksTable    = ConvertTo-HtmlTable -Objects $reportData['ScheduledTasks']  -EmptyMessage "No non-Microsoft scheduled tasks found."

$taskCount = $reportData['ScheduledTasks'].Count
$taskBadge = if ($taskCount -eq 0) {
    "<span class='tk-badge-ok'>None</span>"
} else {
    "<span class='tk-badge-info'>$taskCount task(s)</span>"
}

# AV / Security badge
$avUnprotected = ($reportData['Security'] | Where-Object { $_.RealTimeProtection -eq 'Off' }).Count
$avBadge = if ($reportData['Security'].Count -eq 0) {
    "<span class='tk-badge-warn'>No AV detected</span>"
} elseif ($avUnprotected -gt 0) {
    "<span class='tk-badge-err'>$avUnprotected unprotected</span>"
} else {
    "<span class='tk-badge-ok'>Protected</span>"
}
$securityTable = ConvertTo-HtmlTable -Objects $reportData['Security'] -EmptyMessage "No AV products detected."

$tkConfig   = Get-TKConfig
$tkOrgName  = if (-not [string]::IsNullOrWhiteSpace($tkConfig.OrgName)) { EscHtml $tkConfig.OrgName } else { $null }
$tkSubtitle = if ($tkOrgName) { "$tkOrgName  -  $env:COMPUTERNAME" } else { $env:COMPUTERNAME }

$tkMetaItems = [ordered]@{
    'Machine'    = $env:COMPUTERNAME
    'Run As'     = "$env:USERDOMAIN\$env:USERNAME"
    'Generated'  = $reportTimestamp
    'Storage'    = $storageBadge
    'Updates'    = $updateBadge
}

$tkNavItems = @(
    'Hardware',
    'Operating System',
    'Network Configuration',
    'System Health',
    'Storage & RAID Health',
    'Pending Updates',
    'Installed Software',
    'Event Log',
    'Scheduled Tasks',
    'Security & AV'
)

$htmlReport = (Get-TKHtmlHead `
    -Title      'System Diagnostic Report' `
    -ScriptName 'A.U.S.P.E.X.' `
    -Subtitle   $tkSubtitle `
    -MetaItems  $tkMetaItems `
    -NavItems   $tkNavItems) + @"

  <!-- HARDWARE -->
  <div class="tk-section" id="hardware">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 1</span>
        <h2 class="tk-section-title">Hardware</h2>
      </div>
      <div style="padding:20px;">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:20px;">
          <div class="tk-info-box">
            <div><span class="tk-info-label">Manufacturer</span> $($hw.Manufacturer)</div>
            <div><span class="tk-info-label">Model</span> $($hw.Model)</div>
            <div><span class="tk-info-label">Serial Number</span> $($hw.Serial)</div>
          </div>
          <div class="tk-info-box">
            <div><span class="tk-info-label">CPU</span> $($hw.CPU)</div>
            <div><span class="tk-info-label">Cores / Threads</span> $($hw.Cores) / $($hw.Threads)</div>
            <div><span class="tk-info-label">RAM</span> $($hw.RAMGB) GB</div>
          </div>
        </div>
        <table class="tk-table">
          <thead><tr><th>Drive</th><th>Label</th><th>Total</th><th>Used</th><th>Free</th><th>Usage</th></tr></thead>
          <tbody>$diskRows</tbody>
        </table>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- OPERATING SYSTEM -->
  <div class="tk-section" id="operating-system">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 2</span>
        <h2 class="tk-section-title">Operating System</h2>
      </div>
      <div style="padding:20px;">
        <div class="tk-info-box">
          <div><span class="tk-info-label">OS</span> $($os.Caption)</div>
          <div><span class="tk-info-label">Build</span> $($os.Build)</div>
          <div><span class="tk-info-label">Architecture</span> $($os.Architecture)</div>
          <div><span class="tk-info-label">Install Date</span> $($os.InstallDate)</div>
          <div><span class="tk-info-label">Activation</span> $($os.Activation)</div>
        </div>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- NETWORK -->
  <div class="tk-section" id="network-configuration">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 3</span>
        <h2 class="tk-section-title">Network Configuration</h2>
      </div>
      <div style="padding:20px;">
        <table class="tk-table">
          <thead><tr><th>Adapter</th><th>IP Address</th><th>MAC</th><th>Gateway</th><th>DNS</th></tr></thead>
          <tbody>$netRows</tbody>
        </table>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- HEALTH -->
  <div class="tk-section" id="system-health">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 4</span>
        <h2 class="tk-section-title">System Health</h2>
      </div>
      <div style="padding:20px;">
        <div class="tk-info-box">
          <div><span class="tk-info-label">Last Boot</span> $($health.LastBoot)</div>
          <div><span class="tk-info-label">Uptime</span> $($health.Uptime)</div>
          <div><span class="tk-info-label">Battery</span> $(if ($health.Battery) { "$($health.Battery.Charge)% - $($health.Battery.Status)" } else { "N/A" })</div>
        </div>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- STORAGE & RAID HEALTH -->
  <div class="tk-section" id="storage-raid-health">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 5</span>
        <h2 class="tk-section-title">Storage &amp; RAID Health</h2>
        <span style="margin-left:auto;">$storageBadge</span>
      </div>
      <div style="padding:20px;">
        <p class="tk-section-subtitle" style="margin-bottom:10px;">Physical Disks</p>
        $(if ($physDiskRows) {
          "<table class='tk-table'><thead><tr><th>ID</th><th>Name</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Status</th></tr></thead><tbody>$physDiskRows</tbody></table>"
        } else {
          "<div class='tk-info-box'>No physical disk data available.</div>"
        })
        $(if ($virtDiskRows) {
          "<p class='tk-section-subtitle' style='margin:16px 0 10px;'>Storage Spaces (Virtual Disks)</p><table class='tk-table'><thead><tr><th>Name</th><th>Resiliency</th><th>Size</th><th>Health</th><th>Status</th></tr></thead><tbody>$virtDiskRows</tbody></table>"
        } else {
          "<div class='tk-info-box' style='margin-top:12px;'>No Storage Spaces virtual disks detected.</div>"
        })
        <p class="tk-section-subtitle" style="margin:16px 0 10px;">Hardware RAID Controllers</p>
        $(if ($raidCtrlRows) {
          "<table class='tk-table'><thead><tr><th>Name</th><th>Manufacturer</th><th>Status</th><th>Driver</th></tr></thead><tbody>$raidCtrlRows</tbody></table>"
        } else {
          "<div class='tk-info-box'>No dedicated hardware RAID controllers detected.</div>"
        })
        $(if ($raidVendorHtml) {
          "<p class='tk-section-subtitle' style='margin:16px 0 6px;'>Vendor CLI Detail</p>$raidVendorHtml"
        })
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- PENDING UPDATES -->
  <div class="tk-section" id="pending-updates">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 6</span>
        <h2 class="tk-section-title">Pending Windows Updates</h2>
        <span style="margin-left:auto;">$updateBadge</span>
      </div>
      <div style="padding:20px;">$updatesTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- INSTALLED SOFTWARE -->
  <div class="tk-section" id="installed-software">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 7</span>
        <h2 class="tk-section-title">Installed Software</h2>
        <span class="tk-badge-info" style="margin-left:auto;">$($reportData['Software'].Count) apps</span>
      </div>
      <div style="padding:20px;">$softwareTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- EVENT LOG -->
  <div class="tk-section" id="event-log">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 8</span>
        <h2 class="tk-section-title">Event Log  -  Errors &amp; Critical (Last 24h)</h2>
        <span style="margin-left:auto;">$eventBadge</span>
      </div>
      <div style="padding:20px;">$eventsTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- SCHEDULED TASKS -->
  <div class="tk-section" id="scheduled-tasks">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 9</span>
        <h2 class="tk-section-title">Non-Microsoft Scheduled Tasks</h2>
        <span style="margin-left:auto;">$taskBadge</span>
      </div>
      <div style="padding:20px;">$tasksTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- SECURITY / AV STATUS -->
  <div class="tk-section" id="security-av">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 10</span>
        <h2 class="tk-section-title">Security &amp; Antivirus Status</h2>
        <span style="margin-left:auto;">$avBadge</span>
      </div>
      <div style="padding:20px;">$securityTable</div>
    </div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'A.U.S.P.E.X. v6.0')

try {
    $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8 -Force
    Write-Host "[+] Report saved: $reportPath" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error saving report: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host "     PROBE COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Machine    : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  OS         : $($reportData['OS'].Caption)" -ForegroundColor $ColorSchema.Info
Write-Host "  RAM        : $($reportData['Hardware'].RAMGB) GB" -ForegroundColor $ColorSchema.Info
Write-Host "  Uptime     : $($reportData['Health'].Uptime)" -ForegroundColor $ColorSchema.Info

$physDiskCount  = $reportData['Storage'].PhysicalDisks.Count
$raidCtrlCount  = $reportData['Storage'].RaidControllers.Count
if ($storageUnhealthy -gt 0) {
    Write-Host "  Storage    : $storageUnhealthy degraded disk(s) detected!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  Storage    : $physDiskCount disk(s)  -  all healthy" -ForegroundColor $ColorSchema.Success
}
if ($raidCtrlCount -gt 0) {
    if ($raidCtrlUnhealthy -gt 0) {
        Write-Host "  RAID       : $raidCtrlUnhealthy controller(s) reporting issues!" -ForegroundColor $ColorSchema.Error
    } else {
        Write-Host "  RAID       : $raidCtrlCount controller(s)  -  all OK" -ForegroundColor $ColorSchema.Success
    }
} else {
    Write-Host "  RAID       : No hardware RAID controllers detected" -ForegroundColor $ColorSchema.Info
}

if ($updateCount -gt 0) {
    Write-Host "  Updates    : $updateCount pending" -ForegroundColor $ColorSchema.Warning
}
else {
    Write-Host "  Updates    : Up to date" -ForegroundColor $ColorSchema.Success
}

if ($eventCount -gt 0) {
    Write-Host "  Events     : $eventCount error(s) in last 24h" -ForegroundColor $ColorSchema.Warning
}
else {
    Write-Host "  Events     : Clean" -ForegroundColor $ColorSchema.Success
}

Write-Host "  Ext. Tasks : $taskCount non-Microsoft scheduled task(s)" -ForegroundColor $ColorSchema.Info

if ($avUnprotected -gt 0) {
    Write-Host "  Security   : $avUnprotected AV product(s) with protection OFF" -ForegroundColor $ColorSchema.Error
} elseif ($reportData['Security'].Count -eq 0) {
    Write-Host "  Security   : No AV detected" -ForegroundColor $ColorSchema.Warning
} else {
    Write-Host "  Security   : Protected" -ForegroundColor $ColorSchema.Success
}

Write-Host ""
Show-TKReportResult -Path $reportPath -Unattended:$Unattended
Write-Host ""
Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host "  SCRIPT EXECUTION COMPLETED" -ForegroundColor $ColorSchema.Header
Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
