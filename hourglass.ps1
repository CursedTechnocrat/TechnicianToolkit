# hourglass.ps1 - H.O.U.R.G.L.A.S.S. — Health Of Unit's Rechargeable Gauge: Life, Ageing, State & Service
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
    H.O.U.R.G.L.A.S.S. — Health Of Unit's Rechargeable Gauge: Life, Ageing, State & Service
    Laptop Battery Health Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Surfaces the three numbers that matter for a laptop battery's
    retirement decision: design capacity, current full-charge capacity,
    and cycle count. Pulls values from the ROOT\WMI battery classes --
    BatteryStaticData, BatteryFullChargedCapacity, BatteryCycleCount,
    BatteryStatus -- and correlates against Win32_Battery for the user-
    friendly device name. Reports health as a percentage of design,
    applies industry thresholds (80 / 60 for capacity, 300 / 500 for
    cycles), and emits a dark-themed HTML report with a red / yellow /
    green replacement recommendation.

    Also runs `powercfg /batteryreport` to capture data the live WMI
    classes do not expose: per-battery serial number and manufacture
    date, the full-charge capacity history (degradation trend), and
    Windows' own runtime estimates for active use and connected standby
    at both current full charge and original design capacity. The full
    Microsoft-formatted HTML report is saved alongside HOURGLASS's report and
    linked from it.

.USAGE
    PS C:\> .\hourglass.ps1                    # Interactive run
    PS C:\> .\hourglass.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
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

function Show-HourglassBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  H.O.U.R.G.L.A.S.S. — Health Of Unit's Rechargeable Gauge: Life, Ageing, State & Service" -ForegroundColor Cyan
    Write-Host "  Laptop Battery Health Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Collectors, verdict, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# BATTERY DATA COLLECTION
# ─────────────────────────────────────────────────────────────────────────────

# ROOT\WMI battery classes key off InstanceName so multiple batteries (rare
# but possible on gaming laptops) get the same joining key. Win32_Battery
# provides the device caption that a technician recognises; the ROOT\WMI
# classes provide the energy-accounting figures powercfg uses internally.
function Get-BatteryHealthRecords {
    $win32 = @()
    try { $win32 = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue) } catch {
        # Desktops without a battery will naturally return nothing; surface that as "no battery" later.
    }

    $static = @()
    $full   = @()
    $cycle  = @()
    $status = @()
    try { $static = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryStaticData           -ErrorAction SilentlyContinue) } catch { }
    try { $full   = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryFullChargedCapacity  -ErrorAction SilentlyContinue) } catch { }
    try { $cycle  = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryCycleCount           -ErrorAction SilentlyContinue) } catch { }
    try { $status = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryStatus               -ErrorAction SilentlyContinue) } catch { }

    if ($win32.Count -eq 0 -and $static.Count -eq 0) {
        return @()
    }

    # The union of all instance names tells us how many battery records to emit.
    $instances = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($set in @($static, $full, $cycle, $status)) {
        foreach ($item in $set) {
            if ($item.InstanceName) { [void]$instances.Add($item.InstanceName) }
        }
    }

    # Fallback: no ROOT\WMI data at all. Emit the Win32_Battery record(s) with
    # nulls for the detail fields so the tool still produces a report.
    if ($instances.Count -eq 0) {
        return @($win32 | ForEach-Object {
            [PSCustomObject]@{
                InstanceName   = $_.DeviceID
                Name           = $_.Name
                Caption        = $_.Caption
                Chemistry      = Convert-BatteryChemistry $_.Chemistry
                DesignCapacity = $null
                FullCapacity   = $null
                CycleCount     = $null
                HealthPercent  = $null
                ChargeRate     = $null
                DischargeRate  = $null
                Voltage        = $null
                Charging       = $null
                Discharging    = $null
                StatusText     = Convert-BatteryStatus $_.BatteryStatus
            }
        })
    }

    $rows = foreach ($inst in $instances) {
        $s = $static | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
        $f = $full   | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
        $c = $cycle  | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
        $t = $status | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1

        # Match the instance to a Win32_Battery device for the friendly caption.
        # Instance names look like "PNP0C0A\1_0"; Win32_Battery.DeviceID looks like
        # "1_0" on most machines. Use a suffix match.
        $wi = $win32 | Where-Object { $inst -like "*$($_.DeviceID)*" } | Select-Object -First 1

        $design = if ($s) { [int64]$s.DesignedCapacity } else { $null }
        $fullCap = if ($f) { [int64]$f.FullChargedCapacity } else { $null }
        $cycles = if ($c) { [int]$c.CycleCount } else { $null }

        $health = $null
        if ($design -and $fullCap -and $design -gt 0) {
            $health = [math]::Round(($fullCap / $design) * 100, 1)
        }

        [PSCustomObject]@{
            InstanceName   = $inst
            Name           = if ($wi) { $wi.Name } else { $inst }
            Caption        = if ($wi) { $wi.Caption } else { '' }
            Chemistry      = if ($wi) { Convert-BatteryChemistry $wi.Chemistry } else { 'Unknown' }
            DesignCapacity = $design
            FullCapacity   = $fullCap
            CycleCount     = $cycles
            HealthPercent  = $health
            ChargeRate     = if ($t) { [int]$t.ChargeRate } else { $null }
            DischargeRate  = if ($t) { [int]$t.DischargeRate } else { $null }
            Voltage        = if ($t) { [int]$t.Voltage } else { $null }
            Charging       = if ($t) { [bool]$t.Charging } else { $null }
            Discharging    = if ($t) { [bool]$t.Discharging } else { $null }
            StatusText     = if ($wi) { Convert-BatteryStatus $wi.BatteryStatus } else { 'Unknown' }
        }
    }

    return @($rows)
}

function Convert-BatteryChemistry {
    param($code)
    switch ([int]$code) {
        1 { 'Other' }       2 { 'Unknown' }  3 { 'Lead Acid' }  4 { 'Nickel Cadmium' }
        5 { 'Nickel Metal Hydride' }  6 { 'Lithium-ion' }  7 { 'Zinc air' }  8 { 'Lithium Polymer' }
        default { 'Unknown' }
    }
}

function Convert-BatteryStatus {
    param($code)
    switch ([int]$code) {
        1 { 'Discharging' }       2 { 'AC Power (charging/charged)' }  3 { 'Fully Charged' }
        4 { 'Low' }               5 { 'Critical' }                     6 { 'Charging' }
        7 { 'Charging + High' }   8 { 'Charging + Low' }               9 { 'Charging + Critical' }
        10 { 'Undefined' }        11 { 'Partially Charged' }
        default { 'Unknown' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# POWERCFG /BATTERYREPORT ENRICHMENT
# ─────────────────────────────────────────────────────────────────────────────

# powercfg expresses durations as ISO-8601 (PT4H22M3S). Parse to a minute
# count so totals can sum and format consistently. Returns 0 for null/empty
# input so callers don't have to null-check before adding to a running total.
function Convert-IsoDurationToMinutes {
    param([string]$Iso)
    if ([string]::IsNullOrWhiteSpace($Iso)) { return 0 }
    try {
        $ts = [System.Xml.XmlConvert]::ToTimeSpan($Iso)
        return [int]$ts.TotalMinutes
    } catch {
        return 0
    }
}

# Friendly time formatter for minute counts. Days appear once over 24h so a
# week of standby shows as "7d 2h 15m" instead of "170h 15m".
function Format-Minutes {
    param([int]$Minutes)
    if ($Minutes -le 0) { return '0 min' }
    $d = [int][math]::Floor($Minutes / 1440)
    $h = [int][math]::Floor(($Minutes - $d * 1440) / 60)
    $m = $Minutes - $d * 1440 - $h * 60
    if ($d -gt 0) { return ('{0}d {1}h {2}m' -f $d, $h, $m) }
    if ($h -gt 0) { return ('{0}h {1}m' -f $h, $m) }
    return ('{0}m' -f $m)
}

# Runs powercfg /batteryreport twice -- once as XML for parsing, once as HTML
# so the operator can drill into the full Microsoft-formatted report. The XML
# carries data the ROOT\WMI classes do not expose: per-battery serial numbers,
# manufacture dates, full-charge capacity history (degradation over time),
# and runtime estimates at both current full charge and original design.
function Invoke-PowerCfgBatteryReport {
    param(
        [Parameter(Mandatory)] [string]$LogDir,
        [Parameter(Mandatory)] [string]$Timestamp
    )

    $xmlPath  = Join-Path $LogDir "HOURGLASS_battery_report_${Timestamp}.xml"
    $htmlPath = Join-Path $LogDir "HOURGLASS_battery_report_${Timestamp}.html"

    $result = [PSCustomObject]@{
        XmlPath         = $xmlPath
        HtmlPath        = $htmlPath
        HtmlAvailable   = $false
        Available       = $false
        Error           = $null
        Batteries       = @()
        Estimates       = $null
        CapacityHistory = @()
        UsageSummary    = $null
    }

    try {
        $null = & powercfg.exe /batteryreport /xml /output $xmlPath 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $xmlPath)) {
            $result.Error = "powercfg /batteryreport /xml exit code $LASTEXITCODE"
            return $result
        }
    } catch {
        $result.Error = "powercfg invocation failed: $($_.Exception.Message)"
        return $result
    }

    # The HTML output is a nice-to-have for technicians who want the full
    # Microsoft report. Don't fail the function if it errors -- the XML
    # already gave us everything we need to enrich our own report.
    try {
        $null = & powercfg.exe /batteryreport /output $htmlPath 2>&1
        if ($LASTEXITCODE -eq 0 -and (Test-Path $htmlPath)) {
            $result.HtmlAvailable = $true
        }
    } catch { }

    try {
        $xml = [xml](Get-Content $xmlPath -Raw -ErrorAction Stop)
    } catch {
        $result.Error = "could not read powercfg XML: $($_.Exception.Message)"
        return $result
    }

    $report = $xml.BatteryReport
    if (-not $report) {
        $result.Error = 'powercfg XML had no BatteryReport root element.'
        return $result
    }

    # Per-battery static info -- Win32_Battery doesn't expose serial number
    # or manufacture date, so operators currently have to crack the case to
    # read the label. The XML report has both.
    $batteries = [System.Collections.Generic.List[object]]::new()
    if ($report.Batteries -and $report.Batteries.Battery) {
        foreach ($b in @($report.Batteries.Battery)) {
            $batteries.Add([PSCustomObject]@{
                Id                 = [string]$b.Id
                Manufacturer       = [string]$b.Manufacturer
                SerialNumber       = [string]$b.SerialNumber
                ManufactureDate    = [string]$b.ManufactureDate
                Chemistry          = [string]$b.Chemistry
                DesignCapacity     = if ($b.DesignCapacity)     { [int64]$b.DesignCapacity }     else { $null }
                FullChargeCapacity = if ($b.FullChargeCapacity) { [int64]$b.FullChargeCapacity } else { $null }
            })
        }
    }
    $result.Batteries = @($batteries)

    # Runtime estimates -- powercfg's predicted Active and ConnectedStandby
    # runtime, calculated against both current full charge and the original
    # design capacity. The gap between the two is the runtime the user has
    # already lost to wear, expressed in minutes a technician can quote.
    $rt = $report.RuntimeEstimates
    if ($rt) {
        $atFull   = $rt.FullChargeCapacity
        $atDesign = $rt.DesignCapacity
        $result.Estimates = [PSCustomObject]@{
            ActiveAtFull    = if ($atFull   -and $atFull.ActiveRuntime)      { Convert-IsoDurationToMinutes $atFull.ActiveRuntime }      else { $null }
            StandbyAtFull   = if ($atFull   -and $atFull.ConnectedStandby)   { Convert-IsoDurationToMinutes $atFull.ConnectedStandby }   else { $null }
            ActiveAtDesign  = if ($atDesign -and $atDesign.ActiveRuntime)    { Convert-IsoDurationToMinutes $atDesign.ActiveRuntime }    else { $null }
            StandbyAtDesign = if ($atDesign -and $atDesign.ConnectedStandby) { Convert-IsoDurationToMinutes $atDesign.ConnectedStandby } else { $null }
        }
    }

    # Battery capacity history -- one entry per measurement period showing
    # how full-charge capacity has drifted relative to design. Useful to see
    # whether degradation is accelerating before deciding to replace.
    $history = [System.Collections.Generic.List[object]]::new()
    if ($report.HistoryEntries -and $report.HistoryEntries.HistoryEntry) {
        foreach ($h in @($report.HistoryEntries.HistoryEntry)) {
            $history.Add([PSCustomObject]@{
                StartDate          = [string]$h.StartDate
                EndDate            = [string]$h.EndDate
                DesignCapacity     = if ($h.DesignCapacity)     { [int64]$h.DesignCapacity }     else { $null }
                FullChargeCapacity = if ($h.FullChargeCapacity) { [int64]$h.FullChargeCapacity } else { $null }
                CycleCount         = if ($h.CycleCount)         { [int]$h.CycleCount }            else { $null }
            })
        }
    }
    # Most recent entries first; cap at 12 so the table stays readable on a
    # machine that's been reporting daily for years.
    $result.CapacityHistory = @($history | Sort-Object EndDate -Descending | Select-Object -First 12)

    # Aggregate active/standby/AC/DC time across the runtime history so a
    # technician can answer "is this laptop actually run on battery much?"
    # at a glance, without having to load the full powercfg HTML.
    if ($report.RuntimeHistory -and $report.RuntimeHistory.RuntimeEntry) {
        $entries = @($report.RuntimeHistory.RuntimeEntry)
        $totActDc = 0; $totActAc = 0; $totCsDc = 0; $totCsAc = 0
        foreach ($e in $entries) {
            $totActDc += (Convert-IsoDurationToMinutes $e.ActiveDcTime)
            $totActAc += (Convert-IsoDurationToMinutes $e.ActiveAcTime)
            $totCsDc  += (Convert-IsoDurationToMinutes $e.CsDcTime)
            $totCsAc  += (Convert-IsoDurationToMinutes $e.CsAcTime)
        }
        $result.UsageSummary = [PSCustomObject]@{
            EntryCount   = $entries.Count
            ActiveDcMin  = $totActDc
            ActiveAcMin  = $totActAc
            StandbyDcMin = $totCsDc
            StandbyAcMin = $totCsAc
        }
    }

    $result.Available = $true
    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

# Industry-consensus thresholds. Manufacturers warrant ~80% capacity at 300
# cycles; below 60% the laptop's unplugged runtime is noticeably short.
$script:CapOkPct   = 80.0
$script:CapWarnPct = 60.0
$script:CycOk      = 300
$script:CycWarn    = 500

function Get-BatteryClass {
    param([PSCustomObject]$B)
    if (-not $B.HealthPercent) { return 'info' }
    if ($B.HealthPercent -ge $script:CapOkPct)    { return 'ok' }
    elseif ($B.HealthPercent -ge $script:CapWarnPct) { return 'warn' }
    else { return 'err' }
}

function Get-CycleClass {
    param([PSCustomObject]$B)
    if ($null -eq $B.CycleCount) { return 'info' }
    if ($B.CycleCount -lt $script:CycOk)   { return 'ok' }
    elseif ($B.CycleCount -lt $script:CycWarn) { return 'warn' }
    else { return 'err' }
}

function Get-Verdict {
    param([array]$Batteries)

    if ($Batteries.Count -eq 0) {
        return [PSCustomObject]@{
            Verdict = 'NO BATTERY'
            Class   = 'info'
            Issues  = @()
            Warns   = @('No battery hardware detected — this is expected on desktops and servers.')
            Worst   = $null
        }
    }

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()
    $worstClass = 'ok'

    foreach ($b in $Batteries) {
        $capClass = Get-BatteryClass -B $b
        $cycClass = Get-CycleClass   -B $b

        if ($capClass -eq 'err') {
            $issues.Add("Battery '$($b.Name)' is at $($b.HealthPercent)% of design capacity — below the $script:CapWarnPct% replacement threshold.")
            $worstClass = 'err'
        } elseif ($capClass -eq 'warn' -and $worstClass -ne 'err') {
            $warns.Add("Battery '$($b.Name)' is at $($b.HealthPercent)% of design capacity — plan replacement within the next upgrade cycle.")
            $worstClass = 'warn'
        }

        if ($cycClass -eq 'err') {
            $issues.Add("Battery '$($b.Name)' has $($b.CycleCount) charge cycles — past the $script:CycWarn-cycle comfort limit.")
            $worstClass = 'err'
        } elseif ($cycClass -eq 'warn' -and $worstClass -ne 'err') {
            $warns.Add("Battery '$($b.Name)' has $($b.CycleCount) charge cycles — approaching the $script:CycWarn-cycle limit.")
            if ($worstClass -ne 'err') { $worstClass = 'warn' }
        }

        if ($null -eq $b.HealthPercent -and $null -eq $b.CycleCount) {
            $warns.Add("Battery '$($b.Name)' reports no design/cycle data — the firmware exposes only basic status. Replacement decisions require powercfg /batteryreport.")
        }
    }

    $verdict = switch ($worstClass) {
        'err'  { 'REPLACE NOW' }
        'warn' { 'REPLACEMENT SOON' }
        'info' { 'DATA INCOMPLETE' }
        default { 'HEALTHY' }
    }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $worstClass
        Issues  = @($issues)
        Warns   = @($warns)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Batteries, $Verdict, $PowerCfg)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $bestHealth = if ($Batteries.Count -gt 0) {
        ($Batteries | Where-Object { $_.HealthPercent } | Measure-Object -Property HealthPercent -Maximum).Maximum
    } else { $null }

    $worstHealth = if ($Batteries.Count -gt 0) {
        ($Batteries | Where-Object { $_.HealthPercent } | Measure-Object -Property HealthPercent -Minimum).Minimum
    } else { $null }

    $maxCycles = if ($Batteries.Count -gt 0) {
        ($Batteries | Where-Object { $null -ne $_.CycleCount } | Measure-Object -Property CycleCount -Maximum).Maximum
    } else { $null }

    # Battery detail rows
    $rows = [System.Text.StringBuilder]::new()
    if ($Batteries.Count -eq 0) {
        [void]$rows.Append("<tr><td colspan='8' class='tk-badge-info' style='text-align:center;'>No battery hardware detected.</td></tr>")
    } else {
        foreach ($b in $Batteries) {
            $capClass = Get-BatteryClass -B $b
            $cycClass = Get-CycleClass   -B $b
            $capBadge = if ($null -eq $b.HealthPercent) { "<span class='tk-badge-info'>n/a</span>" } else { "<span class='tk-badge-$capClass'>$($b.HealthPercent)%</span>" }
            $cycBadge = if ($null -eq $b.CycleCount) { "<span class='tk-badge-info'>n/a</span>" } else { "<span class='tk-badge-$cycClass'>$($b.CycleCount)</span>" }
            $design   = if ($b.DesignCapacity) { "$([math]::Round($b.DesignCapacity / 1000, 1)) Wh" } else { 'n/a' }
            $full     = if ($b.FullCapacity)   { "$([math]::Round($b.FullCapacity / 1000, 1)) Wh" } else { 'n/a' }
            $voltage  = if ($b.Voltage)        { "$([math]::Round($b.Voltage / 1000, 2)) V" } else { 'n/a' }
            [void]$rows.Append("<tr><td>$(EscHtml $b.Name)</td><td>$(EscHtml $b.Chemistry)</td><td>$design</td><td>$full</td><td>$capBadge</td><td>$cycBadge</td><td>$voltage</td><td>$(EscHtml $b.StatusText)</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>All batteries are within healthy thresholds.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'H.O.U.R.G.L.A.S.S. Battery Health Report' `
        -ScriptName 'H.O.U.R.G.L.A.S.S.' `
        -Subtitle   "${orgPrefix}Laptop Battery Health -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Generated' = $reportDate
            'Verdict'   = $Verdict.Verdict
            'Batteries' = $Batteries.Count
        }) `
        -NavItems   @('Verdict', 'Batteries', 'powercfg Report')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'H.O.U.R.G.L.A.S.S. v6.0'

    $bestCard  = if ($null -ne $bestHealth)  { "$bestHealth%"  } else { 'n/a' }
    $worstCard = if ($null -ne $worstHealth) { "$worstHealth%" } else { 'n/a' }
    $cycCard   = if ($null -ne $maxCycles)   { "$maxCycles" } else { 'n/a' }

    # ── powercfg /batteryreport enrichment section ──
    $pcSection = [System.Text.StringBuilder]::new()
    [void]$pcSection.Append("`n  <div class=""tk-section"">`n")
    [void]$pcSection.Append("    <div class=""tk-card-header""><span class=""tk-section-title"">powercfg Battery Report</span><span class=""tk-section-num"">historical data</span></div>`n")

    if (-not $PowerCfg -or -not $PowerCfg.Available) {
        $msg = if ($PowerCfg -and $PowerCfg.Error) { $PowerCfg.Error } else { 'powercfg /batteryreport did not produce a parseable XML report.' }
        [void]$pcSection.Append("    <div class=""tk-card""><div class=""tk-info-box"">")
        [void]$pcSection.Append("<span class=""tk-info-label"">powercfg unavailable</span> $(EscHtml $msg)")
        [void]$pcSection.Append("</div></div>`n")
    } else {
        # Per-battery serial / manufacture-date table.
        if ($PowerCfg.Batteries.Count -gt 0) {
            $bRows = [System.Text.StringBuilder]::new()
            foreach ($pb in $PowerCfg.Batteries) {
                $design = if ($pb.DesignCapacity)     { "$([math]::Round($pb.DesignCapacity / 1000, 1)) Wh" } else { 'n/a' }
                $full   = if ($pb.FullChargeCapacity) { "$([math]::Round($pb.FullChargeCapacity / 1000, 1)) Wh" } else { 'n/a' }
                $mfg    = if ([string]::IsNullOrWhiteSpace($pb.ManufactureDate)) { 'n/a' } else { EscHtml $pb.ManufactureDate }
                $serial = if ([string]::IsNullOrWhiteSpace($pb.SerialNumber))    { 'n/a' } else { EscHtml $pb.SerialNumber }
                [void]$bRows.Append("<tr><td>$(EscHtml $pb.Id)</td><td>$(EscHtml $pb.Manufacturer)</td><td class=""tk-mono"">$serial</td><td>$mfg</td><td>$(EscHtml $pb.Chemistry)</td><td>$design</td><td>$full</td></tr>`n")
            }
            [void]$pcSection.Append("    <div class=""tk-card"">`n")
            [void]$pcSection.Append("      <div class=""tk-card-label"">Battery identification</div>`n")
            [void]$pcSection.Append("      <table class=""tk-table"">`n")
            [void]$pcSection.Append("        <thead><tr><th>Battery ID</th><th>Manufacturer</th><th>Serial</th><th>Manufactured</th><th>Chemistry</th><th>Design</th><th>Full Charge</th></tr></thead>`n")
            [void]$pcSection.Append("        <tbody>$($bRows.ToString())</tbody>`n")
            [void]$pcSection.Append("      </table>`n")
            [void]$pcSection.Append("    </div>`n")
        }

        # Runtime estimates -- the headline number for "how much battery have you lost?"
        if ($PowerCfg.Estimates) {
            $e = $PowerCfg.Estimates
            $activeFull   = if ($null -ne $e.ActiveAtFull)    { Format-Minutes $e.ActiveAtFull }    else { 'n/a' }
            $activeDes    = if ($null -ne $e.ActiveAtDesign)  { Format-Minutes $e.ActiveAtDesign }  else { 'n/a' }
            $standbyFull  = if ($null -ne $e.StandbyAtFull)   { Format-Minutes $e.StandbyAtFull }   else { 'n/a' }
            $standbyDes   = if ($null -ne $e.StandbyAtDesign) { Format-Minutes $e.StandbyAtDesign } else { 'n/a' }
            $lostActive   = if ($null -ne $e.ActiveAtFull -and $null -ne $e.ActiveAtDesign) { Format-Minutes ([math]::Max(0, $e.ActiveAtDesign - $e.ActiveAtFull)) } else { 'n/a' }

            [void]$pcSection.Append("    <div class=""tk-card"">`n")
            [void]$pcSection.Append("      <div class=""tk-card-label"">Runtime estimates (Windows-modeled)</div>`n")
            [void]$pcSection.Append("      <table class=""tk-table"">`n")
            [void]$pcSection.Append("        <thead><tr><th>Scenario</th><th>At full charge</th><th>At design capacity</th></tr></thead>`n")
            [void]$pcSection.Append("        <tbody>`n")
            [void]$pcSection.Append("          <tr><td>Active use</td><td>$activeFull</td><td>$activeDes</td></tr>`n")
            [void]$pcSection.Append("          <tr><td>Connected standby</td><td>$standbyFull</td><td>$standbyDes</td></tr>`n")
            [void]$pcSection.Append("        </tbody>`n")
            [void]$pcSection.Append("      </table>`n")
            [void]$pcSection.Append("      <div class=""tk-info-box"" style=""margin-top:14px;""><span class=""tk-info-label"">Active runtime lost to wear</span>$lostActive (full charge vs design)</div>`n")
            [void]$pcSection.Append("    </div>`n")
        }

        # Capacity history -- the degradation trend.
        if ($PowerCfg.CapacityHistory.Count -gt 0) {
            $hRows = [System.Text.StringBuilder]::new()
            foreach ($h in $PowerCfg.CapacityHistory) {
                $period = "$(EscHtml $h.StartDate) &rarr; $(EscHtml $h.EndDate)"
                $design = if ($h.DesignCapacity)     { "$([math]::Round($h.DesignCapacity / 1000, 1)) Wh" } else { 'n/a' }
                $full   = if ($h.FullChargeCapacity) { "$([math]::Round($h.FullChargeCapacity / 1000, 1)) Wh" } else { 'n/a' }
                $pct    = if ($h.DesignCapacity -and $h.FullChargeCapacity -and $h.DesignCapacity -gt 0) {
                    $p = [math]::Round(($h.FullChargeCapacity / $h.DesignCapacity) * 100, 1)
                    $cls = if ($p -ge $script:CapOkPct) { 'ok' } elseif ($p -ge $script:CapWarnPct) { 'warn' } else { 'err' }
                    "<span class=""tk-badge-$cls"">$p%</span>"
                } else { "<span class=""tk-badge-info"">n/a</span>" }
                $cyc    = if ($null -ne $h.CycleCount) { [string]$h.CycleCount } else { 'n/a' }
                [void]$hRows.Append("<tr><td>$period</td><td>$design</td><td>$full</td><td>$pct</td><td>$cyc</td></tr>`n")
            }
            [void]$pcSection.Append("    <div class=""tk-card"">`n")
            [void]$pcSection.Append("      <div class=""tk-card-label"">Full-charge capacity history (most recent $($PowerCfg.CapacityHistory.Count) entries)</div>`n")
            [void]$pcSection.Append("      <table class=""tk-table"">`n")
            [void]$pcSection.Append("        <thead><tr><th>Period</th><th>Design</th><th>Full Charge</th><th>Health %</th><th>Cycles</th></tr></thead>`n")
            [void]$pcSection.Append("        <tbody>$($hRows.ToString())</tbody>`n")
            [void]$pcSection.Append("      </table>`n")
            [void]$pcSection.Append("    </div>`n")
        }

        # AC vs DC time totals -- "is this thing actually run on battery?"
        if ($PowerCfg.UsageSummary) {
            $u = $PowerCfg.UsageSummary
            $totalDc = $u.ActiveDcMin + $u.StandbyDcMin
            $totalAc = $u.ActiveAcMin + $u.StandbyAcMin
            $totalAll = $totalDc + $totalAc
            $dcPct = if ($totalAll -gt 0) { [math]::Round(($totalDc / $totalAll) * 100, 1) } else { 0 }
            [void]$pcSection.Append("    <div class=""tk-card"">`n")
            [void]$pcSection.Append("      <div class=""tk-card-label"">Recent usage ($($u.EntryCount) periods recorded)</div>`n")
            [void]$pcSection.Append("      <table class=""tk-table"">`n")
            [void]$pcSection.Append("        <thead><tr><th>Mode</th><th>On battery (DC)</th><th>Plugged in (AC)</th></tr></thead>`n")
            [void]$pcSection.Append("        <tbody>`n")
            [void]$pcSection.Append("          <tr><td>Active</td><td>$(Format-Minutes $u.ActiveDcMin)</td><td>$(Format-Minutes $u.ActiveAcMin)</td></tr>`n")
            [void]$pcSection.Append("          <tr><td>Connected standby</td><td>$(Format-Minutes $u.StandbyDcMin)</td><td>$(Format-Minutes $u.StandbyAcMin)</td></tr>`n")
            [void]$pcSection.Append("        </tbody>`n")
            [void]$pcSection.Append("      </table>`n")
            [void]$pcSection.Append("      <div class=""tk-info-box"" style=""margin-top:14px;""><span class=""tk-info-label"">Time on battery</span>$dcPct% of recorded runtime</div>`n")
            [void]$pcSection.Append("    </div>`n")
        }

        # Link out to the full Microsoft-formatted HTML report.
        if ($PowerCfg.HtmlAvailable) {
            $href = EscHtml ([System.IO.Path]::GetFileName($PowerCfg.HtmlPath))
            [void]$pcSection.Append("    <div class=""tk-card""><div class=""tk-info-box""><span class=""tk-info-label"">Full Microsoft report</span><a class=""tk-mono"" href=""$href"">$href</a> (saved alongside this report)</div></div>`n")
        }
    }
    [void]$pcSection.Append("  </div>`n")
    $powerCfgBlock = $pcSection.ToString()

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Battery Verdict</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Batteries.Count)</div><div class="tk-summary-lbl">Batteries</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$bestCard</div><div class="tk-summary-lbl">Best Health</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$worstCard</div><div class="tk-summary-lbl">Worst Health</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$cycCard</div><div class="tk-summary-lbl">Max Cycle Count</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$script:CapOkPct% / $script:CycOk</div><div class="tk-summary-lbl">OK Thresholds</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Battery Details</span><span class="tk-section-num">$($Batteries.Count) battery/batteries</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Name</th><th>Chemistry</th><th>Design</th><th>Full Charge</th><th>Health %</th><th>Cycles</th><th>Voltage</th><th>Status</th></tr></thead>
        <tbody>$($rows.ToString())</tbody>
      </table>
      <div class="tk-info-box" style="margin-top:18px;">
        <span class="tk-info-label">Thresholds</span>
        Capacity: &ge; $script:CapOkPct% healthy, $script:CapWarnPct-$script:CapOkPct% plan replacement, &lt; $script:CapWarnPct% replace now.&nbsp;&nbsp;
        Cycles: &lt; $script:CycOk healthy, $script:CycOk-$script:CycWarn plan replacement, &gt;= $script:CycWarn replace now.
      </div>
    </div>
  </div>
$powerCfgBlock
"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-HourglassBanner

Write-Section "BATTERY INVENTORY"
$batteries = Get-BatteryHealthRecords

if ($batteries.Count -eq 0) {
    Write-Warn "No battery hardware detected (desktop / server / VM)."
} else {
    foreach ($b in $batteries) {
        Write-Host "  $($b.Name)  ($($b.Chemistry))" -ForegroundColor $C.Header
        $design = if ($b.DesignCapacity) { "$([math]::Round($b.DesignCapacity / 1000, 1)) Wh" } else { 'n/a' }
        $full   = if ($b.FullCapacity)   { "$([math]::Round($b.FullCapacity / 1000, 1)) Wh" } else { 'n/a' }
        Write-Host ("    Design     : {0}" -f $design) -ForegroundColor $C.Info
        Write-Host ("    Full charge: {0}" -f $full)   -ForegroundColor $C.Info
        if ($null -ne $b.HealthPercent) {
            $cls = Get-BatteryClass -B $b
            $col = switch ($cls) { 'ok' { $C.Success } 'warn' { $C.Warning } 'err' { $C.Error } default { $C.Info } }
            Write-Host ("    Health     : {0}%" -f $b.HealthPercent) -ForegroundColor $col
        }
        if ($null -ne $b.CycleCount) {
            $cls = Get-CycleClass -B $b
            $col = switch ($cls) { 'ok' { $C.Success } 'warn' { $C.Warning } 'err' { $C.Error } default { $C.Info } }
            Write-Host ("    Cycles     : {0}" -f $b.CycleCount) -ForegroundColor $col
        }
        Write-Host ("    Status     : {0}" -f $b.StatusText) -ForegroundColor $C.Info
    }
}
Write-Host ""

$verdict = Get-Verdict -Batteries $batteries

Write-Section "BATTERY VERDICT"
$vColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } 'err' { $C.Error } default { $C.Info } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $vColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All batteries within healthy thresholds." -ForegroundColor $C.Success
}
Write-Host ""

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir    = Resolve-LogDirectory -FallbackPath $ScriptPath

Write-Section "POWERCFG BATTERY REPORT"
if ($batteries.Count -eq 0) {
    Write-Info "  Skipping powercfg /batteryreport (no battery hardware)."
    $powerCfg = $null
} else {
    Write-Step "Running powercfg /batteryreport (XML + HTML)..."
    $powerCfg = Invoke-PowerCfgBatteryReport -LogDir $logDir -Timestamp $timestamp
    if ($powerCfg.Available) {
        Write-Ok "powercfg XML parsed: $($powerCfg.XmlPath)"
        if ($powerCfg.HtmlAvailable) { Write-Ok "powercfg HTML saved: $($powerCfg.HtmlPath)" }
        if ($powerCfg.Estimates) {
            $e = $powerCfg.Estimates
            if ($null -ne $e.ActiveAtFull)   { Write-Host ("    Active runtime (full charge)  : {0}" -f (Format-Minutes $e.ActiveAtFull))   -ForegroundColor $C.Info }
            if ($null -ne $e.ActiveAtDesign) { Write-Host ("    Active runtime (design)       : {0}" -f (Format-Minutes $e.ActiveAtDesign)) -ForegroundColor $C.Info }
        }
        Write-Host ("    Capacity history entries      : {0}" -f $powerCfg.CapacityHistory.Count) -ForegroundColor $C.Info
    } else {
        Write-Warn "  powercfg /batteryreport unavailable: $($powerCfg.Error)"
    }
}
Write-Host ""

Write-Step "Generating HTML report..."
$html    = Build-HtmlReport -Batteries $batteries -Verdict $verdict -PowerCfg $powerCfg
$outPath = Join-Path $logDir "HOURGLASS_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
