# fathom.ps1 - F.A.T.H.O.M. — Free-space Analysis: Tallies Hogs, Old profiles & Mess
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
    F.A.T.H.O.M. — Free-space Analysis: Tallies Hogs, Old profiles & Mess
    Disk & Storage Health Monitor for PowerShell 5.1+

.DESCRIPTION
    F.A.T.H.O.M. is a comprehensive disk and storage health monitoring tool that provides
    physical disk health checks via CIM/WMI, volume space summaries, disk cleanup operations,
    old profile detection, and HTML report generation. It uses Get-PhysicalDisk, Get-Disk,
    Get-Volume, and Win32_DiskDrive to surface health status, operational state, and space usage.
    Volumes below 15% free are flagged as Warning; below 5% as Critical. Full SMART attribute
    reading requires third-party tools such as CrystalDiskInfo — this tool uses HealthStatus
    as a proxy.

.USAGE
    PS C:\> .\fathom.ps1                         # Interactive menu
    PS C:\> .\fathom.ps1 -Unattended             # Run health check and export HTML report silently

.NOTES
    Version : 6.1

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

#region ── Bootstrap ────────────────────────────────────────────────────────────

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

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
Assert-AdminPrivilege

# Script path resolution
$ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# Color hashtable
$C = @{
    Cyan    = 'Cyan'
    Magenta = 'Magenta'
    Green   = 'Green'
    Yellow  = 'Yellow'
    Red     = 'Red'
    Gray    = 'Gray'
    White   = 'White'
}

#endregion

#region ── Banner ───────────────────────────────────────────────────────────────

function Show-Banner {
    Clear-Host
    $banner = @"

  ███████╗ █████╗ ████████╗██╗  ██╗ ██████╗ ███╗   ███╗
  ██╔════╝██╔══██╗╚══██╔══╝██║  ██║██╔═══██╗████╗ ████║
  █████╗  ███████║   ██║   ███████║██║   ██║██╔████╔██║
  ██╔══╝  ██╔══██║   ██║   ██╔══██║██║   ██║██║╚██╔╝██║
  ██║     ██║  ██║   ██║   ██║  ██║╚██████╔╝██║ ╚═╝ ██║
  ╚═╝     ╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝ ╚═════╝ ╚═╝     ╚═╝
"@
    Write-Host $banner -ForegroundColor $C.Cyan
    Write-Host "  Free-space Analysis: Tallies Hogs, Old profiles & Mess" -ForegroundColor $C.Gray
    Write-Host "  Disk & Storage Health Monitor  |  v6.1" -ForegroundColor $C.Gray
    Write-Host ""
}

#endregion

#region ── Helpers ──────────────────────────────────────────────────────────────

function Get-HealthColor {
    param([string]$Status)
    switch ($Status) {
        'Healthy'   { return $C.Green }
        'Warning'   { return $C.Yellow }
        'Unhealthy' { return $C.Red }
        default     { return $C.Gray }
    }
}

#endregion

#region ── Data Collection ──────────────────────────────────────────────────────

function Get-PhysicalDiskInfo {
    $physDisks   = Get-PhysicalDisk -ErrorAction SilentlyContinue
    $diskObjects = Get-Disk        -ErrorAction SilentlyContinue
    $cimDisks    = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($pd in $physDisks) {
        # Try to match by serial or DeviceId
        $matchedDisk = $diskObjects | Where-Object { $_.Number -eq ($pd.DeviceId -replace '\D','') } | Select-Object -First 1
        if (-not $matchedDisk) {
            # Fallback: match by size
            $matchedDisk = $diskObjects | Where-Object { $_.Size -eq $pd.Size } | Select-Object -First 1
        }

        $cimMatch = $cimDisks | Where-Object { $_.Size -eq $pd.Size } | Select-Object -First 1

        $results.Add([PSCustomObject]@{
            FriendlyName      = $pd.FriendlyName
            MediaType         = $pd.MediaType
            Size              = $pd.Size
            SizeFormatted     = Format-Bytes -Bytes ([long]$pd.Size)
            HealthStatus      = if ($pd.HealthStatus) { $pd.HealthStatus } else { 'Unknown' }
            OperationalStatus = if ($pd.OperationalStatus) { $pd.OperationalStatus } else {
                                    if ($cimMatch) { $cimMatch.Status } else { 'Unknown' }
                                }
            BusType           = $pd.BusType
            DiskNumber        = if ($matchedDisk) { $matchedDisk.Number } else { 'N/A' }
            PartitionStyle    = if ($matchedDisk) { $matchedDisk.PartitionStyle } else { 'N/A' }
            IsSystem          = if ($matchedDisk) { $matchedDisk.IsSystem } else { $false }
            IsBoot            = if ($matchedDisk) { $matchedDisk.IsBoot } else { $false }
        })
    }
    return $results
}

function Get-VolumeInfo {
    $volumes = Get-Volume -ErrorAction SilentlyContinue | Where-Object {
        $_.DriveType -ne 'CD-ROM' -and $_.DriveLetter -and $_.Size -gt 0
    }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($vol in $volumes) {
        $pctFree = if ($vol.Size -gt 0) { [Math]::Round(($vol.SizeRemaining / $vol.Size) * 100, 1) } else { 0 }
        $spaceStatus = if ($pctFree -lt 5) { 'Critical' } elseif ($pctFree -lt 15) { 'Warning' } else { 'OK' }

        $results.Add([PSCustomObject]@{
            DriveLetter     = $vol.DriveLetter
            Label           = if ($vol.FileSystemLabel) { $vol.FileSystemLabel } else { '(No Label)' }
            FileSystem      = $vol.FileSystem
            Size            = $vol.Size
            SizeFormatted   = Format-Bytes -Bytes ([long]$vol.Size)
            SizeRemaining   = $vol.SizeRemaining
            FreeFormatted   = Format-Bytes -Bytes ([long]$vol.SizeRemaining)
            PercentFree     = $pctFree
            SpaceStatus     = $spaceStatus
            DriveType       = $vol.DriveType
            HealthStatus    = if ($vol.HealthStatus) { $vol.HealthStatus } else { 'Unknown' }
        })
    }
    return $results
}

#endregion

#region ── Display Functions ────────────────────────────────────────────────────

function Show-DiskSummary {
    Write-Section "PHYSICAL DISK SUMMARY"

    $diskInfo = Get-PhysicalDiskInfo
    if (-not $diskInfo -or $diskInfo.Count -eq 0) {
        Write-Host "  No physical disks detected." -ForegroundColor $C.Yellow
    } else {
        Write-Host ("  {0,-30} {1,-6} {2,-10} {3,-12} {4,-12} {5}" -f `
            "Name", "Num", "Type", "Size", "Health", "OpStatus") -ForegroundColor $C.Gray
        Write-Host ("  " + ("-" * 80)) -ForegroundColor $C.Gray

        foreach ($d in $diskInfo) {
            $hColor = Get-HealthColor -Status $d.HealthStatus
            $flag = if ($d.IsBoot) { "[BOOT]" } elseif ($d.IsSystem) { "[SYS]" } else { "" }
            Write-Host ("  {0,-30} {1,-6} {2,-10} {3,-12}" -f `
                ($d.FriendlyName.Substring(0, [Math]::Min(29, $d.FriendlyName.Length))),
                $d.DiskNumber, $d.MediaType, $d.SizeFormatted) -NoNewline -ForegroundColor $C.White
            Write-Host (" {0,-12}" -f $d.HealthStatus) -NoNewline -ForegroundColor $hColor
            Write-Host (" {0} {1}" -f $d.OperationalStatus, $flag) -ForegroundColor $C.Gray
        }
    }

    Write-Section "VOLUME SPACE SUMMARY"

    $volInfo = Get-VolumeInfo
    if (-not $volInfo -or $volInfo.Count -eq 0) {
        Write-Host "  No volumes detected." -ForegroundColor $C.Yellow
    } else {
        Write-Host ("  {0,-5} {1,-20} {2,-8} {3,-10} {4,-10} {5,-8} {6}" -f `
            "Drive", "Label", "FS", "Total", "Free", "% Free", "Status") -ForegroundColor $C.Gray
        Write-Host ("  " + ("-" * 75)) -ForegroundColor $C.Gray

        foreach ($v in $volInfo) {
            $sColor = switch ($v.SpaceStatus) {
                'Critical' { $C.Red }
                'Warning'  { $C.Yellow }
                default    { $C.Green }
            }
            $bar = if ($v.PercentFree -ge 100) { 100 } else { [int]$v.PercentFree }
            $barStr = "[" + ("#" * [Math]::Round($bar / 5)) + ("." * (20 - [Math]::Round($bar / 5))) + "]"

            Write-Host ("  {0,-5} {1,-20} {2,-8} {3,-10} {4,-10}" -f `
                "$($v.DriveLetter):", $v.Label.Substring(0, [Math]::Min(19, $v.Label.Length)),
                $v.FileSystem, $v.SizeFormatted, $v.FreeFormatted) -NoNewline -ForegroundColor $C.White
            Write-Host (" {0,-6}% {1}" -f $v.PercentFree, $barStr) -ForegroundColor $sColor
        }
    }

    # Totals
    if ($volInfo) {
        $totalSize = ($volInfo | Measure-Object -Property Size -Sum).Sum
        $totalFree = ($volInfo | Measure-Object -Property SizeRemaining -Sum).Sum
        Write-Host ""
        Write-Host ("  Total Storage : {0}" -f (Format-Bytes -Bytes ([long]$totalSize))) -ForegroundColor $C.Cyan
        Write-Host ("  Total Free    : {0}" -f (Format-Bytes -Bytes ([long]$totalFree))) -ForegroundColor $C.Cyan
    }
}

function Show-DiskHealth {
    Write-Section "DISK HEALTH STATUS"

    $diskInfo = Get-PhysicalDiskInfo
    $cimDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue

    Write-Host "  Physical Disk Health (via Get-PhysicalDisk):" -ForegroundColor $C.Cyan
    Write-Host ""

    foreach ($d in $diskInfo) {
        $hColor = Get-HealthColor -Status $d.HealthStatus
        Write-Host "  Disk $($d.DiskNumber): $($d.FriendlyName)" -ForegroundColor $C.White
        Write-Host ("    Media Type        : {0}" -f $d.MediaType) -ForegroundColor $C.Gray
        Write-Host ("    Size              : {0}" -f $d.SizeFormatted) -ForegroundColor $C.Gray
        Write-Host ("    Bus Type          : {0}" -f $d.BusType) -ForegroundColor $C.Gray
        Write-Host ("    Partition Style   : {0}" -f $d.PartitionStyle) -ForegroundColor $C.Gray
        Write-Host "    Health Status     : " -NoNewline -ForegroundColor $C.Gray
        Write-Host $d.HealthStatus -ForegroundColor $hColor
        Write-Host ("    Operational Status: {0}" -f $d.OperationalStatus) -ForegroundColor $C.Gray
        Write-Host ""
    }

    Write-Host "  WMI Win32_DiskDrive Operational Status:" -ForegroundColor $C.Cyan
    Write-Host ""
    foreach ($cd in $cimDisks) {
        $statusColor = if ($cd.Status -eq 'OK') { $C.Green } else { $C.Yellow }
        Write-Host ("  {0}" -f $cd.Caption) -ForegroundColor $C.White
        Write-Host ("    Status       : {0}" -f $cd.Status) -NoNewline -ForegroundColor $C.Gray
        Write-Host "" -ForegroundColor $statusColor
        Write-Host ("    Availability : {0}" -f $cd.Availability) -ForegroundColor $C.Gray
        Write-Host ("    Interface    : {0}" -f $cd.InterfaceType) -ForegroundColor $C.Gray
        Write-Host ("    Sectors/Track: {0}" -f $cd.SectorsPerTrack) -ForegroundColor $C.Gray
        Write-Host ""
    }

    Write-Host "  NOTE: Full SMART attribute data requires a third-party tool such as" -ForegroundColor $C.Yellow
    Write-Host "        CrystalDiskInfo. HealthStatus above is a proxy indicator only." -ForegroundColor $C.Yellow
}

#endregion

#region ── Cleanup Functions ────────────────────────────────────────────────────

function Invoke-TempCleanup {
    $tempPaths = @($env:TEMP, "C:\Windows\Temp")
    $totalDeleted = 0
    $totalBytes   = 0L

    foreach ($path in $tempPaths) {
        if (-not (Test-Path $path)) { continue }

        Write-Host "  Scanning: $path" -ForegroundColor $C.Magenta
        $files = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue |
                 Where-Object { -not $_.PSIsContainer }

        $beforeBytes = ($files | Measure-Object -Property Length -Sum).Sum
        Write-Host ("    Found {0} files ({1})" -f $files.Count, (Format-Bytes -Bytes ([long]$beforeBytes))) -ForegroundColor $C.Gray

        $deleted = 0
        $bytesFreed = 0L
        foreach ($f in $files) {
            try {
                $bytesFreed += $f.Length
                Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                $deleted++
            } catch {
                # Skip locked/protected files
            }
        }

        # Also remove empty subdirectories
        Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.PSIsContainer } |
            Sort-Object FullName -Descending |
            ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }

        Write-Host ("    Deleted {0} files, freed {1}" -f $deleted, (Format-Bytes -Bytes $bytesFreed)) -ForegroundColor $C.Green
        $totalDeleted += $deleted
        $totalBytes   += $bytesFreed
    }

    return [PSCustomObject]@{
        FilesDeleted = $totalDeleted
        BytesFreed   = $totalBytes
    }
}

function Invoke-WUCacheCleanup {
    $wuPath = "C:\Windows\SoftwareDistribution\Download"

    if (-not (Test-Path $wuPath)) {
        Write-Host "  WU cache path not found: $wuPath" -ForegroundColor $C.Yellow
        return
    }

    $beforeSize = (Get-ChildItem -Path $wuPath -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object -Property Length -Sum).Sum

    Write-Host ("  Before: {0}" -f (Format-Bytes -Bytes ([long]$beforeSize))) -ForegroundColor $C.Gray
    Write-Host "  Stopping Windows Update service (wuauserv)..." -ForegroundColor $C.Magenta

    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Host "  Removing cached downloads..." -ForegroundColor $C.Magenta
    Get-ChildItem -Path $wuPath -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

    $afterSize = (Get-ChildItem -Path $wuPath -Recurse -Force -ErrorAction SilentlyContinue |
                  Measure-Object -Property Length -Sum).Sum

    Write-Host "  Starting Windows Update service..." -ForegroundColor $C.Magenta
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue

    $freed = [long]$beforeSize - [long]$afterSize
    Write-Host ("  After : {0}  |  Freed: {1}" -f (Format-Bytes -Bytes ([long]$afterSize)), (Format-Bytes -Bytes $freed)) -ForegroundColor $C.Green
}

function Invoke-RecycleBinCleanup {
    Write-Host "  Emptying Recycle Bin..." -ForegroundColor $C.Magenta
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Host "  Recycle Bin emptied successfully." -ForegroundColor $C.Green
    } catch {
        Write-Host ("  Warning: {0}" -f $_.Exception.Message) -ForegroundColor $C.Yellow
    }
}

function Show-CleanupMenu {
    Write-Section "DISK CLEANUP"

    Write-Host ""
    Write-Host "  Select cleanup operation:" -ForegroundColor $C.White
    Write-Host "    A.  Clear Temp Files (%TEMP% and C:\Windows\Temp)" -ForegroundColor $C.Gray
    Write-Host "    B.  Clear Windows Update Cache" -ForegroundColor $C.Gray
    Write-Host "    C.  Empty Recycle Bin" -ForegroundColor $C.Gray
    Write-Host "    D.  All of the above" -ForegroundColor $C.Cyan
    Write-Host "    X.  Back to main menu" -ForegroundColor $C.Gray
    Write-Host ""

    $sub = Read-Host "  Choice"
    switch ($sub.ToUpper()) {
        'A' {
            $confirm = Read-Host "  Delete temp files? This cannot be undone. [Y/N]"
            if ($confirm -match '^[Yy]') {
                $result = Invoke-TempCleanup
                Write-Host ""
                Write-Host ("  Total: {0} files deleted, {1} freed." -f $result.FilesDeleted, (Format-Bytes -Bytes $result.BytesFreed)) -ForegroundColor $C.Green
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'B' {
            $confirm = Read-Host "  Clear Windows Update cache? [Y/N]"
            if ($confirm -match '^[Yy]') {
                Invoke-WUCacheCleanup
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'C' {
            $confirm = Read-Host "  Empty Recycle Bin? [Y/N]"
            if ($confirm -match '^[Yy]') {
                Invoke-RecycleBinCleanup
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'D' {
            $confirm = Read-Host "  Run ALL cleanup operations? [Y/N]"
            if ($confirm -match '^[Yy]') {
                Write-Host ""
                Write-Host "  [1/3] Temp Files" -ForegroundColor $C.Cyan
                $result = Invoke-TempCleanup
                Write-Host ("  Total temp: {0} files, {1} freed." -f $result.FilesDeleted, (Format-Bytes -Bytes $result.BytesFreed)) -ForegroundColor $C.Green

                Write-Host ""
                Write-Host "  [2/3] Windows Update Cache" -ForegroundColor $C.Cyan
                Invoke-WUCacheCleanup

                Write-Host ""
                Write-Host "  [3/3] Recycle Bin" -ForegroundColor $C.Cyan
                Invoke-RecycleBinCleanup
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'X' { return }
        default { Write-Host "  Invalid selection." -ForegroundColor $C.Yellow }
    }
}

#endregion

#region ── Old Profile Detection ────────────────────────────────────────────────

function Show-OldProfiles {
    Write-Section "OLD USER PROFILES (> 90 DAYS)"

    $usersRoot = "C:\Users"
    $cutoff    = (Get-Date).AddDays(-90)

    if (-not (Test-Path $usersRoot)) {
        Write-Host "  C:\Users not found." -ForegroundColor $C.Yellow
        return
    }

    $excluded = @('Public', 'Default', 'Default User', 'All Users', 'defaultuser0')
    $profiles = Get-ChildItem -Path $usersRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notin $excluded -and $_.LastWriteTime -lt $cutoff }

    if (-not $profiles -or $profiles.Count -eq 0) {
        Write-Host "  No user profiles older than 90 days found." -ForegroundColor $C.Green
        return
    }

    Write-Host ("  {0,-20} {1,-22} {2,-12} {3}" -f "Profile", "Last Modified", "Est. Size", "Path") -ForegroundColor $C.Gray
    Write-Host ("  " + ("-" * 75)) -ForegroundColor $C.Gray

    foreach ($prof in $profiles) {
        $sizeStr = "Calculating..."
        try {
            $sizeBytes = (Get-ChildItem -Path $prof.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                          Measure-Object -Property Length -Sum).Sum
            $sizeStr = Format-Bytes -Bytes ([long]$sizeBytes)
        } catch {
            $sizeStr = "Access Denied"
        }

        $age = [Math]::Round(((Get-Date) - $prof.LastWriteTime).TotalDays)
        $nameStr = $prof.Name.Substring(0, [Math]::Min(19, $prof.Name.Length))

        Write-Host ("  {0,-20} {1,-22} {2,-12} {3}" -f `
            $nameStr,
            ($prof.LastWriteTime.ToString("yyyy-MM-dd") + " ($age days ago)"),
            $sizeStr,
            $prof.FullName) -ForegroundColor $C.Yellow
    }

    Write-Host ""
    Write-Host ("  {0} profile(s) found older than 90 days." -f $profiles.Count) -ForegroundColor $C.Cyan
    Write-Host "  Review before deletion  -  use EMBALM or REVENANT for profile migration." -ForegroundColor $C.Gray
}

#endregion

#region ── HTML Report ──────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$DiskData,
        [array]$VolumeData
    )

    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $hostname   = $env:COMPUTERNAME
    $filename   = "FATHOM_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $outputPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $filename

    $totalDrives   = $DiskData.Count
    $healthyCount  = ($DiskData | Where-Object { $_.HealthStatus -eq 'Healthy' }).Count
    $warnCount     = ($DiskData | Where-Object { $_.HealthStatus -notin @('Healthy', 'Unknown') }).Count
    $totalStorage  = ($VolumeData | Measure-Object -Property Size -Sum).Sum
    $totalFree     = ($VolumeData | Measure-Object -Property SizeRemaining -Sum).Sum

    # Build disk rows
    $diskRows = ""
    foreach ($d in $DiskData) {
        $badgeClass = switch ($d.HealthStatus) {
            'Healthy'   { 'tk-badge-ok'   }
            'Warning'   { 'tk-badge-warn' }
            'Unhealthy' { 'tk-badge-err'  }
            default     { 'tk-badge-info' }
        }
        $bootTag = if ($d.IsBoot) { ' <span class="tk-badge-blue">BOOT</span>' } else { '' }
        $diskRows += @"
            <tr>
                <td>$($d.FriendlyName)$bootTag</td>
                <td>$($d.MediaType)</td>
                <td>$($d.SizeFormatted)</td>
                <td><span class="$badgeClass">$($d.HealthStatus)</span></td>
                <td>$($d.OperationalStatus)</td>
                <td>$($d.BusType)</td>
                <td>$($d.PartitionStyle)</td>
            </tr>
"@
    }

    # Build volume rows
    $volRows = ""
    foreach ($v in $VolumeData) {
        $barClass = switch ($v.SpaceStatus) {
            'Critical' { 'err'  }
            'Warning'  { 'warn' }
            default    { 'ok'   }
        }
        $statusBadgeClass = switch ($v.SpaceStatus) {
            'Critical' { 'tk-badge-err'  }
            'Warning'  { 'tk-badge-warn' }
            default    { 'tk-badge-ok'   }
        }
        $barWidth = [Math]::Min(100, [Math]::Max(1, [int]$v.PercentFree))

        $volRows += @"
            <tr>
                <td><strong>$($v.DriveLetter):</strong></td>
                <td>$($v.Label)</td>
                <td>$($v.FileSystem)</td>
                <td>$($v.SizeFormatted)</td>
                <td>$($v.FreeFormatted)</td>
                <td>
                    <div class="tk-progress-wrap">
                        <div class="tk-progress-bar $barClass" style="width:$($barWidth)%"></div>
                    </div>
                    <span class="tk-mono" style="font-size:11px;">$($v.PercentFree)%</span>
                </td>
                <td><span class="$statusBadgeClass">$($v.SpaceStatus)</span></td>
                <td>$($v.HealthStatus)</td>
            </tr>
"@
    }

    # Recommendations
    $recommendations = ""
    $critVols = $VolumeData | Where-Object { $_.SpaceStatus -eq 'Critical' }
    $warnVols = $VolumeData | Where-Object { $_.SpaceStatus -eq 'Warning' }
    $badDisks = $DiskData   | Where-Object { $_.HealthStatus -notin @('Healthy', 'Unknown') }

    if ($critVols) {
        foreach ($v in $critVols) {
            $recommendations += "<li class='tk-info-box' style='list-style:none;margin-bottom:8px;'><span class='tk-info-label tk-badge-err'>CRITICAL</span> Drive <strong>$($v.DriveLetter):</strong> ($($v.Label)) is critically low on space ($($v.PercentFree)% free). Immediate action required.</li>`n"
        }
    }
    if ($warnVols) {
        foreach ($v in $warnVols) {
            $recommendations += "<li class='tk-info-box' style='list-style:none;margin-bottom:8px;'><span class='tk-info-label tk-badge-warn'>WARNING</span> Drive <strong>$($v.DriveLetter):</strong> ($($v.Label)) has low free space ($($v.PercentFree)% free). Consider cleanup.</li>`n"
        }
    }
    if ($badDisks) {
        foreach ($d in $badDisks) {
            $recommendations += "<li class='tk-info-box' style='list-style:none;margin-bottom:8px;'><span class='tk-info-label tk-badge-err'>DISK ALERT</span> Physical disk <strong>$($d.FriendlyName)</strong> reports Health Status: <strong>$($d.HealthStatus)</strong>. Back up data immediately.</li>`n"
        }
    }
    if (-not $recommendations) {
        $recommendations = "<li class='tk-info-box' style='list-style:none;margin-bottom:8px;'><span class='tk-info-label tk-badge-ok'>OK</span> All monitored disks and volumes are within healthy thresholds.</li>"
    }

    $warnCountClass = if ($warnCount -gt 0) { 'warn' } else { 'ok' }

    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $html  = Get-TKHtmlHead `
        -Title      'Disk & Storage Health Report' `
        -ScriptName 'F.A.T.H.O.M.' `
        -Subtitle   "$orgPrefix$hostname" `
        -MetaItems  ([ordered]@{
            'Host'      = $hostname
            'Generated' = $timestamp
            'Tool'      = 'Disk & Storage Health Monitor v3.6'
        }) `
        -NavItems   @('Physical Disks', 'Volume Space', 'Recommendations')

    $html += @"

<div class="tk-summary-row">
    <div class="tk-summary-card info">
        <div class="tk-summary-num">$totalDrives</div>
        <div class="tk-summary-lbl">Total Drives</div>
    </div>
    <div class="tk-summary-card ok">
        <div class="tk-summary-num">$healthyCount</div>
        <div class="tk-summary-lbl">Healthy</div>
    </div>
    <div class="tk-summary-card $warnCountClass">
        <div class="tk-summary-num">$warnCount</div>
        <div class="tk-summary-lbl">Warning / Critical</div>
    </div>
    <div class="tk-summary-card info">
        <div class="tk-summary-num">$(Format-Bytes -Bytes ([long]$totalStorage))</div>
        <div class="tk-summary-lbl">Total Storage</div>
    </div>
    <div class="tk-summary-card info">
        <div class="tk-summary-num">$(Format-Bytes -Bytes ([long]$totalFree))</div>
        <div class="tk-summary-lbl">Free Storage</div>
    </div>
</div>

<div class="tk-section" id="physical-disks">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Physical Disks</span>
        </div>
        <table class="tk-table">
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Type</th>
                    <th>Size</th>
                    <th>Health Status</th>
                    <th>Operational Status</th>
                    <th>Bus Type</th>
                    <th>Partition Style</th>
                </tr>
            </thead>
            <tbody>
$diskRows
            </tbody>
        </table>
        <div class="tk-info-box" style="margin-top:12px">
            <span class="tk-info-label">NOTE</span>
            Health Status is sourced from <code>Get-PhysicalDisk</code> and serves as a proxy for SMART data.
            For full SMART attribute analysis (reallocated sectors, spin retries, uncorrectable errors, etc.),
            use a dedicated tool such as <strong>CrystalDiskInfo</strong> or <strong>smartctl</strong>.
        </div>
    </div>
</div>

<div class="tk-section" id="volume-space">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Volume Space</span>
        </div>
        <table class="tk-table">
            <thead>
                <tr>
                    <th>Drive</th>
                    <th>Label</th>
                    <th>File System</th>
                    <th>Total Size</th>
                    <th>Free Space</th>
                    <th>% Free</th>
                    <th>Space Status</th>
                    <th>Health</th>
                </tr>
            </thead>
            <tbody>
$volRows
            </tbody>
        </table>
    </div>
</div>

<div class="tk-section" id="recommendations">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Recommendations</span>
        </div>
        <ul class="rec-list" style="list-style:none;padding:0;margin:0;">
$recommendations
        </ul>
    </div>
</div>

"@

    $html += Get-TKHtmlFoot -ScriptName 'F.A.T.H.O.M. v6.1'

    $html | Out-File -FilePath $outputPath -Encoding UTF8 -Force
    return $outputPath
}

#endregion

#region ── Main Loop ────────────────────────────────────────────────────────────

function Show-Menu {
    Show-Banner
    Write-Host "  MAIN MENU" -ForegroundColor $C.Cyan
    Write-Host ""
    Write-Host "    1.  Show Disk & Volume Summary" -ForegroundColor $C.White
    Write-Host "    2.  Check Disk Health Status" -ForegroundColor $C.White
    Write-Host "    3.  Run Disk Cleanup" -ForegroundColor $C.White
    Write-Host "    4.  Detect Large Old Profiles (> 90 days)" -ForegroundColor $C.White
    Write-Host "    5.  Export HTML Health Report" -ForegroundColor $C.White
    Write-Host ""
    Write-Host "    Q.  Quit" -ForegroundColor $C.Gray
    Write-Host ""
}

if ($Unattended) {
    Show-Banner
    Write-Host "  Running in Unattended mode  -  collecting data and generating report..." -ForegroundColor $C.Magenta
    Write-Host ""

    Write-Host "  Collecting physical disk data..." -ForegroundColor $C.Gray
    $diskData = Get-PhysicalDiskInfo

    Write-Host "  Collecting volume data..." -ForegroundColor $C.Gray
    $volData  = Get-VolumeInfo

    Write-Host "  Building HTML report..." -ForegroundColor $C.Gray
    $reportPath = Build-HtmlReport -DiskData $diskData -VolumeData $volData

    Write-Host ""
    Write-Host ("  Report saved: {0}" -f $reportPath) -ForegroundColor $C.Green
    Write-Host "  Unattended run complete." -ForegroundColor $C.Cyan
    exit 0
}

# Interactive loop
do {
    Show-Menu
    $choice = Read-Host "  Select an option"

    switch ($choice.ToUpper()) {
        '1' {
            Show-DiskSummary
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '2' {
            Show-DiskHealth
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '3' {
            Show-CleanupMenu
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '4' {
            Show-OldProfiles
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '5' {
            Write-Section "EXPORTING HTML REPORT"
            Write-Host "  Collecting disk data..." -ForegroundColor $C.Magenta
            $diskData = Get-PhysicalDiskInfo

            Write-Host "  Collecting volume data..." -ForegroundColor $C.Magenta
            $volData  = Get-VolumeInfo

            Write-Host "  Building report..." -ForegroundColor $C.Magenta
            $reportPath = Build-HtmlReport -DiskData $diskData -VolumeData $volData

            Write-Host ""
            Show-TKReportResult -Path $reportPath -Unattended:$Unattended
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        'Q' {
            Show-Banner
            Write-Host "  Exiting F.A.T.H.O.M.  -  Disk & Storage Health Monitor." -ForegroundColor $C.Cyan
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  Invalid selection. Please choose 1-5 or Q." -ForegroundColor $C.Yellow
            Start-Sleep -Seconds 1
        }
    }
} while ($choice.ToUpper() -ne 'Q')

if ($Transcript) { Stop-TKTranscript }

#endregion
