# scryer.ps1 - S.C.R.Y.E.R. -- System Consolidated Report Yielding Exhaustive Results
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
    S.C.R.Y.E.R. -- System Consolidated Report Yielding Exhaustive Results
    Unified Diagnostic Report for PowerShell 5.1+

.DESCRIPTION
    Runs all major diagnostic queries in a single pass and produces one
    comprehensive HTML report suitable for machine handoffs or audit records.
    Covers system information, user accounts, disk space, disk health (SMART),
    and service/scheduled task status.

.USAGE
    PS C:\> .\scryer.ps1                          # Must be run as Administrator
    PS C:\> .\scryer.ps1 -Unattended              # Silent mode, no prompts
    PS C:\> .\scryer.ps1 -OutputPath "D:\Reports" # Write report to a specific directory

.NOTES
    Version : 6.1

#>

param(
    [switch]$Unattended,
    [ValidateScript({
        [string]::IsNullOrWhiteSpace($_) -or
        (Test-Path -LiteralPath $_) -or
        (Test-Path -LiteralPath (Split-Path -Parent $_))
    })]
    [string]$OutputPath = "",
    [switch]$Transcript
)

# ---------------------------------------------------------------------------
# INITIALIZATION
# ---------------------------------------------------------------------------

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

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

# ---------------------------------------------------------------------------
# COLOR SCHEMA
# ---------------------------------------------------------------------------

$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ---------------------------------------------------------------------------
# BANNER
# ---------------------------------------------------------------------------

function Show-ScryerBanner {
    Clear-Host
    Write-Host @"

  ███████╗ ██████╗██████╗ ██╗   ██╗███████╗██████╗
  ██╔════╝██╔════╝██╔══██╗╚██╗ ██╔╝██╔════╝██╔══██╗
  ███████╗██║     ██████╔╝ ╚████╔╝ █████╗  ██████╔╝
  ╚════██║██║     ██╔══██╗  ╚██╔╝  ██╔══╝  ██╔══██╗
  ███████║╚██████╗██║  ██║   ██║   ███████╗██║  ██║
  ╚══════╝ ╚═════╝╚═╝  ╚═╝   ╚═╝   ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    S.C.R.Y.E.R. -- System Consolidated Report Yielding Exhaustive Results" -ForegroundColor Cyan
    Write-Host "    Unified Diagnostic Report Tool" -ForegroundColor Cyan
    Write-Host ""
}

if (-not $Unattended) { Show-ScryerBanner }

Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  UNIFIED DIAGNOSTIC REPORT" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Run As    : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Started   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $ColorSchema.Info
Write-Host ""

# ---------------------------------------------------------------------------
# SECTION 1 -- SYSTEM OVERVIEW
# ---------------------------------------------------------------------------

Write-Host "  [1/5] Collecting system information..." -ForegroundColor $ColorSchema.Progress

$sysOS   = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$sysCS   = Get-CimInstance -ClassName Win32_ComputerSystem  -ErrorAction SilentlyContinue
$sysCPU  = Get-CimInstance -ClassName Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1

$osCaption    = if ($sysOS)  { $sysOS.Caption }       else { "Unknown" }
$osBuild      = if ($sysOS)  { $sysOS.BuildNumber }   else { "" }
$manufacturer = if ($sysCS)  { $sysCS.Manufacturer }  else { "" }
$model        = if ($sysCS)  { $sysCS.Model }          else { "" }
$cpuName      = if ($sysCPU) { $sysCPU.Name.Trim() }  else { "Unknown" }
$cpuCores     = if ($sysCS)  { $sysCS.NumberOfLogicalProcessors } else { "" }
$totalRAM_GB  = if ($sysCS)  { [math]::Round($sysCS.TotalPhysicalMemory / 1GB, 1) } else { 0 }
$freeRAM_GB   = if ($sysOS)  { [math]::Round($sysOS.FreePhysicalMemory / 1MB, 1) }  else { 0 }
$lastBoot     = if ($sysOS)  { $sysOS.LastBootUpTime.ToString("yyyy-MM-dd HH:mm") }  else { "" }
$uptimeSpan   = if ($sysOS)  { (Get-Date) - $sysOS.LastBootUpTime } else { $null }
$uptimeStr    = if ($uptimeSpan) { "$([int]$uptimeSpan.TotalDays)d $($uptimeSpan.Hours)h" } else { "" }
$psVersion    = $PSVersionTable.PSVersion.ToString()

Write-Host "  [+] System info collected." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 2 -- USER ACCOUNTS
# ---------------------------------------------------------------------------

Write-Host "  [2/5] Auditing local user accounts..." -ForegroundColor $ColorSchema.Progress

$localUsers   = Get-LocalUser -ErrorAction SilentlyContinue
$adminMembers = @()
try { $adminMembers = (Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue).Name }
catch {}

$userRows = foreach ($u in ($localUsers | Sort-Object { -not ($adminMembers -contains "$env:COMPUTERNAME\$($u.Name)") }, { $_.LastLogon } -Descending)) {
    $isAdmin = ($adminMembers -contains "$env:COMPUTERNAME\$($u.Name)") -or ($adminMembers -contains $u.Name)
    [PSCustomObject]@{
        Name        = $u.Name
        FullName    = $u.FullName
        Enabled     = $u.Enabled
        LastLogon   = if ($u.LastLogon) { $u.LastLogon.ToString("yyyy-MM-dd") } else { "Never" }
        IsAdmin     = $isAdmin
        PwdRequired = $u.PasswordRequired
    }
}

Write-Host "  [+] User accounts audited ($($userRows.Count) users)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 3 -- DISK SPACE
# ---------------------------------------------------------------------------

Write-Host "  [3/5] Checking disk space..." -ForegroundColor $ColorSchema.Progress

$volumes = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
$volumeRows = foreach ($v in $volumes) {
    $totalGB = [math]::Round($v.Size / 1GB, 1)
    $freeGB  = [math]::Round($v.FreeSpace / 1GB, 1)
    $usedGB  = [math]::Round(($v.Size - $v.FreeSpace) / 1GB, 1)
    $pctUsed = if ($v.Size -gt 0) { [math]::Round(($v.Size - $v.FreeSpace) / $v.Size * 100, 0) } else { 0 }
    $health  = if ($pctUsed -gt 95) { "err" } elseif ($pctUsed -gt 85) { "warn" } else { "ok" }
    [PSCustomObject]@{
        Letter  = $v.DeviceID
        Label   = $v.VolumeName
        TotalGB = $totalGB
        UsedGB  = $usedGB
        FreeGB  = $freeGB
        PctUsed = $pctUsed
        Health  = $health
    }
}

$volWarnCount = ($volumeRows | Where-Object { $_.Health -ne "ok" }).Count
Write-Host "  [+] Disk space checked ($($volumeRows.Count) volumes, $volWarnCount warnings)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 4 -- DISK HEALTH
# ---------------------------------------------------------------------------

Write-Host "  [4/5] Assessing disk health..." -ForegroundColor $ColorSchema.Progress

$diskRows      = @()
$smartAvailable = $false
try {
    $physDisks      = Get-PhysicalDisk -ErrorAction Stop
    $smartAvailable = $true
    $diskRows = foreach ($d in $physDisks) {
        $sizeGB = [math]::Round($d.Size / 1GB, 0)
        $wear   = $null
        $temp   = $null
        try {
            $rel = Get-StorageReliabilityCounter -PhysicalDisk $d -ErrorAction SilentlyContinue
            if ($rel) {
                $wear = $rel.Wear
                $temp = $rel.Temperature
            }
        } catch {}
        $healthClass = switch -Regex ($d.HealthStatus) {
            'Healthy' { "ok"   }
            'Warning' { "warn" }
            default   { "err"  }
        }
        [PSCustomObject]@{
            Name        = $d.FriendlyName
            MediaType   = $d.MediaType
            SizeGB      = $sizeGB
            Health      = $d.HealthStatus
            HealthClass = $healthClass
            Temp        = $temp
            Wear        = $wear
        }
    }
} catch {
    $smartAvailable = $false
}

$diskWarnCount = ($diskRows | Where-Object { $_.HealthClass -ne "ok" }).Count
Write-Host "  [+] Disk health assessed ($($diskRows.Count) disks)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 5 -- SERVICES & SCHEDULED TASKS
# ---------------------------------------------------------------------------

Write-Host "  [5/5] Checking services and scheduled tasks..." -ForegroundColor $ColorSchema.Progress

$triggerExclusions = @('gupdate','gupdatem','edgeupdate','edgeupdatem','MapsBroker',
    'RemoteRegistry','SharedAccess','TabletInputService','WbioSrvc','lfsvc',
    'SCardSvr','SensrSvc','WSearch','wuauserv','BITS','DoSvc','UsoSvc','WerSvc',
    'AppReadiness','tiledatamodelsvc','CDPSvc','OneSyncSvc','PimIndexMaintenanceSvc',
    'MessagingService','cbdhsvc','DevicesFlowUserSvc')

$stoppedSvcs = Get-Service -ErrorAction SilentlyContinue |
    Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -eq 'Stopped' -and
                   $_.Name -notin $triggerExclusions } |
    Select-Object Name, DisplayName, Status, StartType |
    Sort-Object DisplayName

$failedTasks = @()
try {
    $failedTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.State -ne 'Disabled' -and $_.TaskPath -notlike '\Microsoft\*' } |
        ForEach-Object {
            $info = $_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
            if ($info -and $info.LastTaskResult -ne 0 -and $info.LastRunTime -gt [datetime]::MinValue) {
                [PSCustomObject]@{
                    TaskName    = $_.TaskName
                    TaskPath    = $_.TaskPath
                    LastRunTime = $info.LastRunTime.ToString("yyyy-MM-dd HH:mm")
                    LastResult  = "0x{0:X8}" -f $info.LastTaskResult
                }
            }
        } | Where-Object { $_ -ne $null } | Select-Object -First 20
} catch {}

Write-Host "  [+] Services/tasks checked ($($stoppedSvcs.Count) svc issues, $($failedTasks.Count) task failures)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# HTML REPORT BUILD
# ---------------------------------------------------------------------------

$_cfg = Get-TKConfig
$reportDir = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath }
             elseif (-not [string]::IsNullOrWhiteSpace($_cfg.LogDirectory)) { $_cfg.LogDirectory }
             else { $PSScriptRoot }
if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
$reportFile = Join-Path $reportDir "SCRYER_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

Write-Host ""
Write-Host "  [*] Building HTML report..." -ForegroundColor $ColorSchema.Progress

$tkCfg     = Get-TKConfig
$orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

$html = Get-TKHtmlHead -Title "SCRYER Unified Report" `
    -ScriptName "S.C.R.Y.E.R." `
    -Subtitle "$orgPrefix$env:COMPUTERNAME" `
    -MetaItems ([ordered]@{
        'Generated' = (Get-Date -Format 'yyyy-MM-dd HH:mm')
        'OS'        = "$osCaption (Build $osBuild)"
        'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
    }) `
    -NavItems @('System Overview','User Accounts','Disk Space','Disk Health','Services and Tasks')

# ------------------------------------------------------------------
# Section 01 -- System Overview
# ------------------------------------------------------------------

$html += @"

<div class="tk-section" id="system-overview">
  <div class="tk-section-title"><span class="tk-section-num">01</span> System Overview</div>
  <div class="tk-card">
    <table class="tk-table">
      <tbody>
        <tr><td class="tk-card-label">Hostname</td><td>$(EscHtml $env:COMPUTERNAME)</td></tr>
        <tr><td class="tk-card-label">Operating System</td><td>$(EscHtml $osCaption) (Build $osBuild)</td></tr>
        <tr><td class="tk-card-label">Manufacturer / Model</td><td>$(EscHtml $manufacturer) $(EscHtml $model)</td></tr>
        <tr><td class="tk-card-label">CPU</td><td>$(EscHtml $cpuName) ($cpuCores logical cores)</td></tr>
        <tr><td class="tk-card-label">RAM</td><td>$totalRAM_GB GB total / $freeRAM_GB GB free</td></tr>
        <tr><td class="tk-card-label">Last Boot</td><td>$(EscHtml $lastBoot)</td></tr>
        <tr><td class="tk-card-label">Uptime</td><td>$(EscHtml $uptimeStr)</td></tr>
        <tr><td class="tk-card-label">PowerShell Version</td><td>$(EscHtml $psVersion)</td></tr>
      </tbody>
    </table>
  </div>
</div>

"@

# ------------------------------------------------------------------
# Section 02 -- User Accounts
# ------------------------------------------------------------------

$userTableRows = ""
foreach ($u in $userRows) {
    $statusBadge = if ($u.Enabled) {
        "<span class='tk-badge-ok'>Enabled</span>"
    } else {
        "<span class='tk-badge-err'>Disabled</span>"
    }
    $adminCell = if ($u.IsAdmin) { "<span class='tk-badge-err'>Yes</span>" } else { "" }
    $pwdCell   = if ($u.PwdRequired) { "Yes" } else { "No" }
    $userTableRows += @"
        <tr>
          <td>$(EscHtml $u.Name)</td>
          <td>$(EscHtml $u.FullName)</td>
          <td>$statusBadge</td>
          <td>$(EscHtml $u.LastLogon)</td>
          <td>$adminCell</td>
          <td>$pwdCell</td>
        </tr>
"@
}

$html += @"

<div class="tk-section" id="user-accounts">
  <div class="tk-section-title"><span class="tk-section-num">02</span> User Accounts</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead>
        <tr>
          <th>User</th>
          <th>Full Name</th>
          <th>Status</th>
          <th>Last Logon</th>
          <th>Admin</th>
          <th>Pwd Required</th>
        </tr>
      </thead>
      <tbody>
$userTableRows
      </tbody>
    </table>
  </div>
</div>

"@

# ------------------------------------------------------------------
# Section 03 -- Disk Space
# ------------------------------------------------------------------

$diskSpaceCards = ""
foreach ($vol in $volumeRows) {
    $letterLabel = if ($vol.Label) { "$(EscHtml $vol.Letter) -- $(EscHtml $vol.Label)" } else { "$(EscHtml $vol.Letter)" }
    $diskSpaceCards += @"
  <div class="tk-card">
    <div class="tk-card-header">$letterLabel</div>
    <div class="tk-progress-wrap">
      <div class="tk-progress-bar $($vol.Health)" style="width: $($vol.PctUsed)%"></div>
    </div>
    <div style="font-size:0.85em; color:var(--tk-text-dim); margin-top:4px">
      $($vol.PctUsed)% used &nbsp;|&nbsp; $($vol.UsedGB) GB used / $($vol.TotalGB) GB total &nbsp;|&nbsp; $($vol.FreeGB) GB free
    </div>
  </div>
"@
}

if (-not $diskSpaceCards) {
    $diskSpaceCards = "<div class='tk-card'><div class='tk-info-box'>No fixed volumes found.</div></div>"
}

$html += @"

<div class="tk-section" id="disk-space">
  <div class="tk-section-title"><span class="tk-section-num">03</span> Disk Space</div>
$diskSpaceCards
</div>

"@

# ------------------------------------------------------------------
# Section 04 -- Disk Health
# ------------------------------------------------------------------

if (-not $smartAvailable) {
    $diskHealthBody = @"
  <div class="tk-card">
    <div class="tk-info-box">SMART data unavailable -- Storage module not accessible on this system.</div>
  </div>
"@
} else {
    $diskHealthRows = ""
    foreach ($d in $diskRows) {
        $healthBadge = "<span class='tk-badge-$($d.HealthClass)'>$(EscHtml $d.Health)</span>"
        $tempVal     = if ($null -ne $d.Temp) { $d.Temp } else { "--" }
        $wearVal     = if ($null -ne $d.Wear) { "$($d.Wear)%" } else { "--" }
        $diskHealthRows += @"
        <tr>
          <td>$(EscHtml $d.Name)</td>
          <td>$(EscHtml $d.MediaType)</td>
          <td>$($d.SizeGB)</td>
          <td>$healthBadge</td>
          <td>$tempVal</td>
          <td>$wearVal</td>
        </tr>
"@
    }
    $diskHealthBody = @"
  <div class="tk-card">
    <table class="tk-table">
      <thead>
        <tr>
          <th>Drive</th>
          <th>Type</th>
          <th>Size (GB)</th>
          <th>Health</th>
          <th>Temp (C)</th>
          <th>Wear (%)</th>
        </tr>
      </thead>
      <tbody>
$diskHealthRows
      </tbody>
    </table>
  </div>
"@
}

$html += @"

<div class="tk-section" id="disk-health">
  <div class="tk-section-title"><span class="tk-section-num">04</span> Disk Health</div>
$diskHealthBody
</div>

"@

# ------------------------------------------------------------------
# Section 05 -- Services and Tasks
# ------------------------------------------------------------------

# Stopped automatic services card
if (-not $stoppedSvcs -or $stoppedSvcs.Count -eq 0) {
    $svcCardBody = "    <div class='tk-info-box'><span class='tk-badge-ok'>OK</span> No stopped automatic services found.</div>"
} else {
    $svcRows = ""
    foreach ($s in $stoppedSvcs) {
        $svcRows += @"
        <tr>
          <td>$(EscHtml $s.Name)</td>
          <td>$(EscHtml $s.DisplayName)</td>
          <td><span class='tk-badge-warn'>Stopped</span></td>
          <td>$(EscHtml $s.StartType)</td>
        </tr>
"@
    }
    $svcCardBody = @"
    <table class="tk-table">
      <thead>
        <tr>
          <th>Name</th>
          <th>Display Name</th>
          <th>Status</th>
          <th>Start Type</th>
        </tr>
      </thead>
      <tbody>
$svcRows
      </tbody>
    </table>
"@
}

# Failed scheduled tasks card
if (-not $failedTasks -or $failedTasks.Count -eq 0) {
    $taskCardBody = "    <div class='tk-info-box'><span class='tk-badge-ok'>OK</span> No failed scheduled tasks found.</div>"
} else {
    $taskRows = ""
    foreach ($t in $failedTasks) {
        $taskRows += @"
        <tr>
          <td>$(EscHtml $t.TaskName)</td>
          <td>$(EscHtml $t.TaskPath)</td>
          <td>$(EscHtml $t.LastRunTime)</td>
          <td><span class='tk-badge-warn'>$(EscHtml $t.LastResult)</span></td>
        </tr>
"@
    }
    $taskCardBody = @"
    <table class="tk-table">
      <thead>
        <tr>
          <th>Task Name</th>
          <th>Path</th>
          <th>Last Run</th>
          <th>Last Result</th>
        </tr>
      </thead>
      <tbody>
$taskRows
      </tbody>
    </table>
"@
}

$html += @"

<div class="tk-section" id="services-and-tasks">
  <div class="tk-section-title"><span class="tk-section-num">05</span> Services and Tasks</div>
  <div class="tk-card">
    <div class="tk-card-header">Stopped Automatic Services</div>
$svcCardBody
  </div>
  <div class="tk-card">
    <div class="tk-card-header">Failed Scheduled Tasks</div>
$taskCardBody
  </div>
</div>

"@

# ------------------------------------------------------------------
# Footer
# ------------------------------------------------------------------

$html += Get-TKHtmlFoot -ScriptName "S.C.R.Y.E.R. v6.1"

# ------------------------------------------------------------------
# Write report and open
# ------------------------------------------------------------------

$html | Out-File -FilePath $reportFile -Encoding UTF8
Write-Host "  [+] Report saved: $reportFile" -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# CONSOLE SUMMARY
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SCRYER COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Report    : $reportFile" -ForegroundColor $ColorSchema.Info
Write-Host "  Users     : $($userRows.Count)" -ForegroundColor $ColorSchema.Info
Write-Host "  Volumes   : $($volumeRows.Count) ($volWarnCount with warnings)" -ForegroundColor $(if ($volWarnCount -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host "  Disks     : $($diskRows.Count) ($diskWarnCount with issues)" -ForegroundColor $(if ($diskWarnCount -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host "  Svc Issues: $($stoppedSvcs.Count)" -ForegroundColor $(if ($stoppedSvcs.Count -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host "  Task Fails: $($failedTasks.Count)" -ForegroundColor $(if ($failedTasks.Count -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

Show-TKReportResult -Path $reportFile -Unattended:$Unattended

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
