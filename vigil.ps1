# vigil.ps1 - V.I.G.I.L. — Verifies Integrity of services, Guards & Inspects Logs
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
    V.I.G.I.L. — Verifies Integrity of services, Guards & Inspects Logs
    Service, Task & Event Log Monitor for PowerShell 5.1+

.DESCRIPTION
    Audits critical Windows services, scheduled tasks, and recent event log errors on the local
    machine or optionally a remote machine via WinRM. Generates a dark-themed HTML health report,
    displays interactive console summaries, and can restart stopped critical services with
    per-service confirmation prompts.

.USAGE
    PS C:\> .\vigil.ps1                              # Interactive menu (local machine)
    PS C:\> .\vigil.ps1 -Unattended                  # Export health report silently
    PS C:\> .\vigil.ps1 -Unattended -Target HOSTNAME  # Remote machine report

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [ValidatePattern('^$|^[A-Za-z0-9][A-Za-z0-9._-]{0,252}[A-Za-z0-9]$|^[A-Za-z0-9]$')]
    [string]$Target = '',
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
# REMOTE TARGET STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:RemoteTarget = $Target.Trim()

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-VigilBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗   ██╗██╗ ██████╗ ██╗██╗
  ██║   ██║██║██╔════╝ ██║██║
  ██║   ██║██║██║  ███╗██║██║
  ╚██╗ ██╔╝██║██║   ██║██║██║
   ╚████╔╝ ██║╚██████╔╝██║███████╗
    ╚═══╝  ╚═╝ ╚═════╝ ╚═╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    V.I.G.I.L. — Verifies Integrity of services, Guards & Inspects Logs" -ForegroundColor Cyan
    Write-Host "    Service, Task & Event Log Monitor" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Show-Spinner {
    param([string]$Message, [scriptblock]$Work)
    $frames = @('|','/','-','\')
    $i = 0
    $job = Start-Job -ScriptBlock $Work
    Write-Host -NoNewline "  $Message " -ForegroundColor $C.Progress
    while ($job.State -eq 'Running') {
        Write-Host -NoNewline "`b$($frames[$i % 4])" -ForegroundColor $C.Progress
        $i++
        Start-Sleep -Milliseconds 120
    }
    Write-Host -NoNewline "`b" -ForegroundColor $C.Progress
    $result = Receive-Job $job
    Remove-Job $job
    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# CRITICAL SERVICES LIST
# ─────────────────────────────────────────────────────────────────────────────

$CriticalServices = @(
    @{ Name='wuauserv';          Display='Windows Update'                   },
    @{ Name='WinDefend';         Display='Windows Defender Antivirus'       },
    @{ Name='EventLog';          Display='Windows Event Log'                },
    @{ Name='Schedule';          Display='Task Scheduler'                   },
    @{ Name='Dnscache';          Display='DNS Client'                       },
    @{ Name='LanmanWorkstation'; Display='Workstation (SMB Client)'         },
    @{ Name='W32Time';           Display='Windows Time'                     },
    @{ Name='SamSs';             Display='Security Accounts Manager'        },
    @{ Name='RpcSs';             Display='Remote Procedure Call'            },
    @{ Name='BITS';              Display='Background Intelligent Transfer'  },
    @{ Name='cryptsvc';          Display='Cryptographic Services'           },
    @{ Name='MpsSvc';            Display='Windows Firewall'                 },
    @{ Name='Spooler';           Display='Print Spooler'                    },
    @{ Name='lmhosts';           Display='TCP/IP NetBIOS Helper'            },
    @{ Name='WerSvc';            Display='Windows Error Reporting'          }
)

# ─────────────────────────────────────────────────────────────────────────────
# GET-SERVICEHEALTH
# ─────────────────────────────────────────────────────────────────────────────

function Get-ServiceHealth {
    $serviceNames = $CriticalServices | ForEach-Object { $_.Name }
    $displayMap   = @{}
    foreach ($svc in $CriticalServices) { $displayMap[$svc.Name] = $svc.Display }

    $scriptBlock = {
        param($names)
        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($name in $names) {
            try {
                $svc = Get-Service -Name $name -ErrorAction Stop
                $results.Add([PSCustomObject]@{
                    Name      = $svc.Name
                    Status    = $svc.Status.ToString()
                    StartType = $svc.StartType.ToString()
                })
            }
            catch {
                $results.Add([PSCustomObject]@{
                    Name      = $name
                    Status    = 'NotFound'
                    StartType = 'Unknown'
                })
            }
        }
        return $results
    }

    if ($script:RemoteTarget) {
        try {
            $raw = Invoke-Command -ComputerName $script:RemoteTarget -ScriptBlock $scriptBlock -ArgumentList (,$serviceNames) -ErrorAction Stop
        }
        catch {
            Write-Host "  [-] Remote service query failed: $_" -ForegroundColor $C.Error
            $raw = @()
        }
    } else {
        $raw = & $scriptBlock -names $serviceNames
    }

    $output = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $raw) {
        $concern = ($r.Status -eq 'Stopped' -and $r.StartType -eq 'Automatic')
        $output.Add([PSCustomObject]@{
            Name        = $r.Name
            DisplayName = if ($displayMap.ContainsKey($r.Name)) { $displayMap[$r.Name] } else { $r.Name }
            Status      = $r.Status
            StartType   = $r.StartType
            Concern     = $concern
        })
    }
    return $output
}

# ─────────────────────────────────────────────────────────────────────────────
# GET-TASKAUDIT
# ─────────────────────────────────────────────────────────────────────────────

function Get-TaskAudit {
    $scriptBlock = {
        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue
        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($task in $tasks) {
            try {
                $info = $task | Get-ScheduledTaskInfo -ErrorAction Stop
                $results.Add([PSCustomObject]@{
                    TaskName       = $task.TaskName
                    TaskPath       = $task.TaskPath
                    State          = $task.State.ToString()
                    LastRunTime    = $info.LastRunTime
                    LastTaskResult = $info.LastTaskResult
                    NextRunTime    = $info.NextRunTime
                })
            }
            catch {
                $results.Add([PSCustomObject]@{
                    TaskName       = $task.TaskName
                    TaskPath       = $task.TaskPath
                    State          = $task.State.ToString()
                    LastRunTime    = $null
                    LastTaskResult = -1
                    NextRunTime    = $null
                })
            }
        }
        return $results
    }

    if ($script:RemoteTarget) {
        try {
            $raw = Invoke-Command -ComputerName $script:RemoteTarget -ScriptBlock $scriptBlock -ErrorAction Stop
        }
        catch {
            Write-Host "  [-] Remote task query failed: $_" -ForegroundColor $C.Error
            return @()
        }
    } else {
        $raw = & $scriptBlock
    }

    $cutoff = (Get-Date).AddDays(-7)
    # LastTaskResult codes: 0=success, 267009=running, 267011=never run
    $successCodes = @(0, 267009, 267011)

    $output = [System.Collections.Generic.List[object]]::new()
    foreach ($t in $raw) {
        $isMicrosoft  = $t.TaskPath -like '\Microsoft\*'
        $lastRunStale = ($null -ne $t.LastRunTime -and $t.LastRunTime -lt $cutoff -and $t.LastRunTime -gt [datetime]'1900-01-01')
        $hasFailed    = ($t.LastTaskResult -notin $successCodes)
        $isDisabled   = ($t.State -eq 'Disabled')

        # Flag logic
        $flagReason = ''
        if ($hasFailed -and $lastRunStale -and -not $isMicrosoft) {
            $flagReason = 'Failed+Stale'
        } elseif ($hasFailed -and -not $isMicrosoft) {
            $flagReason = 'Failed'
        } elseif ($isDisabled -and -not $isMicrosoft) {
            $flagReason = 'Disabled'
        } elseif ($hasFailed -and $isMicrosoft) {
            $flagReason = 'MSFailed'
        }

        $output.Add([PSCustomObject]@{
            TaskName       = $t.TaskName
            TaskPath       = $t.TaskPath
            State          = $t.State
            LastRunTime    = $t.LastRunTime
            LastTaskResult = $t.LastTaskResult
            NextRunTime    = $t.NextRunTime
            IsMicrosoft    = $isMicrosoft
            FlagReason     = $flagReason
        })
    }
    return $output
}

# ─────────────────────────────────────────────────────────────────────────────
# GET-EVENTERRORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-EventErrors {
    $since = (Get-Date).AddHours(-24)

    if ($script:RemoteTarget) {
        $target = $script:RemoteTarget
        $sysErrors = try {
            Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 2; StartTime = $since } -MaxEvents 50 -ComputerName $target -ErrorAction Stop
        } catch { @() }

        $appErrors = try {
            Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 2; StartTime = $since } -MaxEvents 50 -ComputerName $target -ErrorAction Stop
        } catch { @() }
    } else {
        $sysErrors = try {
            Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 2; StartTime = $since } -MaxEvents 50 -ErrorAction Stop
        } catch { @() }

        $appErrors = try {
            Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 2; StartTime = $since } -MaxEvents 50 -ErrorAction Stop
        } catch { @() }
    }

    $allErrors = [System.Collections.Generic.List[object]]::new()
    foreach ($e in $sysErrors) {
        $allErrors.Add([PSCustomObject]@{
            Log           = 'System'
            TimeGenerated = $e.TimeCreated
            EntryType     = $e.LevelDisplayName
            Source        = $e.ProviderName
            EventID       = $e.Id
            Message       = if ($e.Message -and $e.Message.Length -gt 120) { $e.Message.Substring(0,120) + '...' } else { $e.Message }
        })
    }
    foreach ($e in $appErrors) {
        $allErrors.Add([PSCustomObject]@{
            Log           = 'Application'
            TimeGenerated = $e.TimeCreated
            EntryType     = $e.LevelDisplayName
            Source        = $e.ProviderName
            EventID       = $e.Id
            Message       = if ($e.Message -and $e.Message.Length -gt 120) { $e.Message.Substring(0,120) + '...' } else { $e.Message }
        })
    }

    # Build source summary (top 10)
    $sourceSummary = $allErrors |
        Group-Object Source |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        ForEach-Object { [PSCustomObject]@{ Source = $_.Name; Count = $_.Count } }

    return [PSCustomObject]@{
        Events        = ($allErrors | Sort-Object TimeGenerated -Descending)
        SourceSummary = $sourceSummary
        TotalCount    = $allErrors.Count
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# INVOKE-RESTARTSERVICE
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RestartService {
    param([array]$ServiceData)

    $stopped = $ServiceData | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' }

    if (-not $stopped -or $stopped.Count -eq 0) {
        Write-Host "  [+] No stopped Automatic-start services to restart." -ForegroundColor $C.Success
        return
    }

    Write-Host "  Found $($stopped.Count) stopped Automatic-start service(s):" -ForegroundColor $C.Warning
    Write-Host ""

    $restartAll = $false
    $skipAll    = $false

    foreach ($svc in $stopped) {
        if ($skipAll) { break }

        $label = "$($svc.DisplayName) ($($svc.Name))"

        if ($restartAll) {
            Write-Host "  [*] Restarting: $label" -ForegroundColor $C.Progress
        } else {
            Write-Host -NoNewline "  Restart $label ? (Y/N/A=All/S=Skip all): " -ForegroundColor $C.Warning
            $answer = (Read-Host).Trim().ToUpper()
            switch ($answer) {
                'A' { $restartAll = $true }
                'S' { $skipAll    = $true; Write-Host "  [*] Skipping remaining." -ForegroundColor $C.Info; break }
                'N' { Write-Host "  [*] Skipped." -ForegroundColor $C.Info; continue }
            }
        }

        if (-not $skipAll) {
            try {
                if ($script:RemoteTarget) {
                    Invoke-Command -ComputerName $script:RemoteTarget -ScriptBlock {
                        param($n) Start-Service -Name $n -ErrorAction Stop
                    } -ArgumentList $svc.Name -ErrorAction Stop
                } else {
                    Start-Service -Name $svc.Name -ErrorAction Stop
                }
                Write-Host "  [+] Started: $label" -ForegroundColor $C.Success
            }
            catch {
                Write-Host "  [-] Failed to start $label`: $_" -ForegroundColor $C.Error
            }
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW-SERVICEHEALTH (console)
# ─────────────────────────────────────────────────────────────────────────────

function Show-ServiceHealth {
    param([array]$Data)

    Write-Section "SERVICE HEALTH — CRITICAL SERVICES"

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Host "  No service data available." -ForegroundColor $C.Warning
        return
    }

    $colW = @(26, 34, 10, 12, 8)
    $hdr  = "  {0,-$($colW[0])}{1,-$($colW[1])}{2,-$($colW[2])}{3,-$($colW[3])}{4}" -f 'Service Name','Display Name','Status','Start Type','Concern'
    Write-Host $hdr -ForegroundColor $C.Header

    foreach ($svc in $Data) {
        $statusColor = switch ($svc.Status) {
            'Running' { $C.Success }
            'Stopped' { $C.Error   }
            default   { $C.Warning }
        }
        $concernStr  = if ($svc.Concern) { '!! CONCERN' } else { '' }
        $concernColor = if ($svc.Concern) { $C.Error } else { $C.Info }

        $line = "  {0,-$($colW[0])}{1,-$($colW[1])}{2,-$($colW[2])}{3,-$($colW[3])}" -f `
            $svc.Name, $svc.DisplayName, $svc.Status, $svc.StartType
        Write-Host -NoNewline $line
        Write-Host $concernStr -ForegroundColor $concernColor
    }

    $ok      = ($Data | Where-Object { -not $_.Concern }).Count
    $concern = ($Data | Where-Object {    $_.Concern }).Count
    Write-Host ""
    Write-Host "  Services OK: $ok   Concerns: $concern" -ForegroundColor $(if ($concern -gt 0) { $C.Warning } else { $C.Success })
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW-TASKAUDIT (console)
# ─────────────────────────────────────────────────────────────────────────────

function Show-TaskAudit {
    param([array]$Data)

    Write-Section "SCHEDULED TASK AUDIT"

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Host "  No task data available." -ForegroundColor $C.Warning
        return
    }

    $flagged = $Data | Where-Object { $_.FlagReason -and $_.FlagReason -ne 'MSFailed' } | Sort-Object FlagReason
    $msFailed = $Data | Where-Object { $_.FlagReason -eq 'MSFailed' }

    if ($flagged.Count -eq 0) {
        Write-Host "  [+] No problematic non-Microsoft tasks found." -ForegroundColor $C.Success
    } else {
        Write-Host "  Non-Microsoft flagged tasks ($($flagged.Count)):" -ForegroundColor $C.Warning
        Write-Host ""
        $hdr = "  {0,-32}{1,-28}{2,-10}{3,-12}{4}" -f 'Task Name','Path','State','Last Result','Flag'
        Write-Host $hdr -ForegroundColor $C.Header

        foreach ($t in $flagged) {
            $color = switch -Wildcard ($t.FlagReason) {
                'Failed*' { $C.Error   }
                'Disabled' { $C.Warning }
                default    { $C.Info   }
            }
            $shortPath = if ($t.TaskPath.Length -gt 27) { $t.TaskPath.Substring(0,24) + '...' } else { $t.TaskPath }
            $shortName = if ($t.TaskName.Length -gt 31) { $t.TaskName.Substring(0,28) + '...' } else { $t.TaskName }
            $line = "  {0,-32}{1,-28}{2,-10}{3,-12}{4}" -f $shortName, $shortPath, $t.State, $t.LastTaskResult, $t.FlagReason
            Write-Host $line -ForegroundColor $color
        }
    }

    if ($msFailed.Count -gt 0) {
        Write-Host ""
        Write-Host "  Microsoft tasks with errors ($($msFailed.Count)):" -ForegroundColor $C.Warning
        foreach ($t in $msFailed | Select-Object -First 10) {
            Write-Host "    $($t.TaskPath)$($t.TaskName)  [Result: $($t.LastTaskResult)]" -ForegroundColor $C.Info
        }
    }

    Write-Host ""
    $disabledNonMs = ($Data | Where-Object { $_.FlagReason -eq 'Disabled' }).Count
    $failedNonMs   = ($Data | Where-Object { $_.FlagReason -like 'Failed*' }).Count
    Write-Host "  Failed (non-MS): $failedNonMs   Disabled (non-MS): $disabledNonMs   MS errors: $($msFailed.Count)" -ForegroundColor $C.Info
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW-EVENTERRORS (console)
# ─────────────────────────────────────────────────────────────────────────────

function Show-EventErrors {
    param([object]$Data)

    Write-Section "RECENT EVENT LOG ERRORS (Last 24 Hours)"

    if (-not $Data -or $Data.TotalCount -eq 0) {
        Write-Host "  [+] No errors found in System or Application logs." -ForegroundColor $C.Success
        return
    }

    Write-Host "  Total errors: $($Data.TotalCount)" -ForegroundColor $C.Warning
    Write-Host ""

    # Source summary
    Write-Host "  Top Error Sources:" -ForegroundColor $C.Header
    $hdr = "  {0,-40}{1}" -f 'Source','Count'
    Write-Host $hdr -ForegroundColor $C.Header
    foreach ($src in $Data.SourceSummary) {
        $color = if ($src.Count -ge 10) { $C.Error } elseif ($src.Count -ge 3) { $C.Warning } else { $C.Info }
        Write-Host ("  {0,-40}{1}" -f $src.Source, $src.Count) -ForegroundColor $color
    }

    Write-Host ""
    Write-Host "  Recent Events (newest first):" -ForegroundColor $C.Header
    $hdr2 = "  {0,-22}{1,-6}{2,-26}{3,-8}{4}" -f 'Time','Log','Source','EventID','Message'
    Write-Host $hdr2 -ForegroundColor $C.Header

    foreach ($evt in ($Data.Events | Select-Object -First 20)) {
        $timeStr   = $evt.TimeGenerated.ToString('MM/dd HH:mm:ss')
        $shortSrc  = if ($evt.Source.Length -gt 25) { $evt.Source.Substring(0,22) + '...' } else { $evt.Source }
        $shortMsg  = if ($evt.Message.Length -gt 50) { $evt.Message.Substring(0,47) + '...' } else { $evt.Message }
        $logAbbr   = if ($evt.Log -eq 'System') { 'SYS' } else { 'APP' }
        $line = "  {0,-22}{1,-6}{2,-26}{3,-8}{4}" -f $timeStr, $logAbbr, $shortSrc, $evt.EventID, $shortMsg
        Write-Host $line -ForegroundColor $C.Error
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BUILD-HTMLREPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$Services,
        [array]$Tasks,
        [object]$Events,
        [string]$MachineName,
        [string]$ReportTimestamp
    )

    # Summary counts
    $svcOk       = ($Services | Where-Object { -not $_.Concern }).Count
    $svcConcern  = ($Services | Where-Object {    $_.Concern }).Count
    $taskFailed  = ($Tasks    | Where-Object { $_.FlagReason -like 'Failed*' }).Count
    $taskDisabled = ($Tasks   | Where-Object { $_.FlagReason -eq 'Disabled' }).Count
    $eventTotal  = if ($Events) { $Events.TotalCount } else { 0 }

    # Service rows
    $svcRows = ''
    foreach ($svc in $Services) {
        $statusBadge = if ($svc.Status -eq 'Running') {
            "<span class='tk-badge-ok'>Running</span>"
        } elseif ($svc.Status -eq 'Stopped') {
            "<span class='tk-badge-err'>Stopped</span>"
        } else {
            "<span class='tk-badge-warn'>$(EscHtml($svc.Status))</span>"
        }
        $concernCell = if ($svc.Concern) {
            "<span class='tk-badge-err'>!! Concern</span>"
        } else {
            "<span style='color:#555'>-</span>"
        }
        $svcRows += "<tr>
            <td>$(EscHtml($svc.Name))</td>
            <td>$(EscHtml($svc.DisplayName))</td>
            <td>$statusBadge</td>
            <td>$(EscHtml($svc.StartType))</td>
            <td>$concernCell</td>
        </tr>`n"
    }

    # Task rows — non-Microsoft flagged
    $taskRowsFlagged = ''
    $flaggedTasks = $Tasks | Where-Object { $_.FlagReason -and $_.FlagReason -ne 'MSFailed' } | Sort-Object FlagReason
    foreach ($t in $flaggedTasks) {
        $resultBadge = if ($t.LastTaskResult -in @(0, 267009, 267011)) {
            "<span class='tk-badge-ok'>$($t.LastTaskResult)</span>"
        } else {
            "<span class='tk-badge-err'>$($t.LastTaskResult)</span>"
        }
        $stateBadge = if ($t.State -eq 'Ready') {
            "<span class='tk-badge-ok'>Ready</span>"
        } elseif ($t.State -eq 'Disabled') {
            "<span class='tk-badge-warn'>Disabled</span>"
        } elseif ($t.State -eq 'Running') {
            "<span class='tk-badge-ok'>Running</span>"
        } else {
            "<span class='tk-badge-info'>$(EscHtml($t.State))</span>"
        }
        $lastRun  = if ($t.LastRunTime -and $t.LastRunTime -gt [datetime]'1900-01-01') { $t.LastRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
        $nextRun  = if ($t.NextRunTime -and $t.NextRunTime -gt [datetime]'1900-01-01') { $t.NextRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
        $taskRowsFlagged += "<tr>
            <td>$(EscHtml($t.TaskName))</td>
            <td class='tk-mono' style='font-size:11px'>$(EscHtml($t.TaskPath))</td>
            <td>$stateBadge</td>
            <td>$lastRun</td>
            <td>$resultBadge</td>
            <td>$nextRun</td>
        </tr>`n"
    }

    # Task rows — MS failed
    $taskRowsMs = ''
    $msFailed = $Tasks | Where-Object { $_.FlagReason -eq 'MSFailed' }
    foreach ($t in $msFailed) {
        $lastRun = if ($t.LastRunTime -and $t.LastRunTime -gt [datetime]'1900-01-01') { $t.LastRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
        $taskRowsMs += "<tr>
            <td>$(EscHtml($t.TaskName))</td>
            <td class='tk-mono' style='font-size:11px'>$(EscHtml($t.TaskPath))</td>
            <td><span class='tk-badge-info'>$(EscHtml($t.State))</span></td>
            <td>$lastRun</td>
            <td><span class='tk-badge-warn'>$($t.LastTaskResult)</span></td>
            <td>N/A</td>
        </tr>`n"
    }

    # Event source summary rows
    $srcRows = ''
    if ($Events -and $Events.SourceSummary) {
        foreach ($src in $Events.SourceSummary) {
            $cntBadge = if ($src.Count -ge 10) { 'tk-badge-err' } elseif ($src.Count -ge 3) { 'tk-badge-warn' } else { 'tk-badge-info' }
            $srcRows += "<tr><td>$(EscHtml($src.Source))</td><td><span class='$cntBadge'>$($src.Count)</span></td></tr>`n"
        }
    }

    # Event detail rows
    $evtRows = ''
    if ($Events -and $Events.Events) {
        foreach ($evt in ($Events.Events | Select-Object -First 50)) {
            $timeStr  = $evt.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss')
            $logBadge = if ($evt.Log -eq 'System') {
                "<span class='tk-badge-warn'>SYS</span>"
            } else {
                "<span class='tk-badge-info'>APP</span>"
            }
            $evtRows += "<tr>
                <td class='tk-mono'>$timeStr</td>
                <td>$logBadge</td>
                <td>$(EscHtml($evt.Source))</td>
                <td class='tk-mono'>$($evt.EventID)</td>
                <td style='font-size:12px'>$(EscHtml($evt.Message))</td>
            </tr>`n"
        }
    }

    $noFlagged  = if ($flaggedTasks.Count -eq 0) { "<div class='tk-info-box'><span class='tk-info-label'>INFO</span> No non-Microsoft tasks flagged.</div>" } else { '' }
    $noMsFailed = if ($msFailed.Count -eq 0)     { "<div class='tk-info-box'><span class='tk-info-label'>INFO</span> No Microsoft task failures.</div>"      } else { '' }
    $noEvents   = if ($eventTotal -eq 0)          { "<div class='tk-info-box'><span class='tk-info-label tk-badge-ok'>OK</span> No errors in System or Application logs in the last 24 hours.</div>" } else { '' }

    $svcOkClass      = if ($svcConcern -gt 0) { 'warn' } else { 'ok' }
    $svcConcernClass = if ($svcConcern -eq 0) { 'ok' }   else { 'err' }
    $taskFailClass   = if ($taskFailed -eq 0)  { 'ok' }   else { 'err' }
    $taskDisClass    = if ($taskDisabled -eq 0) { 'ok' }  else { 'warn' }
    $evtClass        = if ($eventTotal -eq 0)  { 'ok' }   elseif ($eventTotal -lt 10) { 'warn' } else { 'err' }

    $showFlagTable = if ($flaggedTasks.Count -eq 0) { ' style="display:none"' } else { '' }
    $showMsTable   = if ($msFailed.Count -eq 0)     { ' style="display:none"' } else { '' }

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $html  = Get-TKHtmlHead `
        -Title      'Service & Task Health Report' `
        -ScriptName 'V.I.G.I.L.' `
        -Subtitle   "$orgPrefix$MachineName" `
        -MetaItems  ([ordered]@{
            'Machine'   = $MachineName
            'Generated' = $ReportTimestamp
            'Tool'      = 'Service, Task & Event Log Monitor v3.6'
        }) `
        -NavItems   @('Critical Services', 'Scheduled Tasks', 'Event Log Errors')

    $html += @"

<div class="tk-summary-row">
    <div class="tk-summary-card $svcOkClass">
        <div class="tk-summary-num">$svcOk</div>
        <div class="tk-summary-lbl">Services OK</div>
    </div>
    <div class="tk-summary-card $svcConcernClass">
        <div class="tk-summary-num">$svcConcern</div>
        <div class="tk-summary-lbl">Services Concern</div>
    </div>
    <div class="tk-summary-card $taskFailClass">
        <div class="tk-summary-num">$taskFailed</div>
        <div class="tk-summary-lbl">Tasks Failed</div>
    </div>
    <div class="tk-summary-card $taskDisClass">
        <div class="tk-summary-num">$taskDisabled</div>
        <div class="tk-summary-lbl">Tasks Disabled</div>
    </div>
    <div class="tk-summary-card $evtClass">
        <div class="tk-summary-num">$eventTotal</div>
        <div class="tk-summary-lbl">Event Errors (24h)</div>
    </div>
</div>

<div class="tk-section" id="critical-services">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Critical Services</span>
        </div>
        <table class="tk-table">
            <thead><tr>
                <th>Service Name</th>
                <th>Display Name</th>
                <th>Status</th>
                <th>Start Type</th>
                <th>Concern</th>
            </tr></thead>
            <tbody>
                $svcRows
            </tbody>
        </table>
    </div>
</div>

<div class="tk-section" id="scheduled-tasks">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Scheduled Tasks</span>
        </div>
        <div class="tk-section-subtitle" style="padding:0 0 10px 0">Non-Microsoft Flagged Tasks</div>
        $noFlagged
        <table class="tk-table"$showFlagTable>
            <thead><tr>
                <th>Task Name</th>
                <th>Path</th>
                <th>State</th>
                <th>Last Run</th>
                <th>Last Result</th>
                <th>Next Run</th>
            </tr></thead>
            <tbody>
                $taskRowsFlagged
            </tbody>
        </table>
        <div class="tk-divider"></div>
        <div class="tk-section-subtitle" style="padding:10px 0">Microsoft Tasks with Errors</div>
        $noMsFailed
        <table class="tk-table"$showMsTable>
            <thead><tr>
                <th>Task Name</th>
                <th>Path</th>
                <th>State</th>
                <th>Last Run</th>
                <th>Last Result</th>
                <th>Next Run</th>
            </tr></thead>
            <tbody>
                $taskRowsMs
            </tbody>
        </table>
    </div>
</div>

<div class="tk-section" id="event-log-errors">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Event Log Errors (Last 24 Hours)</span>
        </div>
        $noEvents
        <div class="tk-section-subtitle" style="padding:0 0 10px 0">Top Error Sources</div>
        <table class="tk-table" style="max-width:480px">
            <thead><tr><th>Source</th><th>Error Count</th></tr></thead>
            <tbody>$srcRows</tbody>
        </table>
        <div class="tk-divider"></div>
        <div class="tk-section-subtitle" style="padding:10px 0">Recent Error Events (newest first, up to 50)</div>
        <table class="tk-table">
            <thead><tr>
                <th>Time</th>
                <th>Log</th>
                <th>Source</th>
                <th>Event ID</th>
                <th>Message</th>
            </tr></thead>
            <tbody>
                $evtRows
            </tbody>
        </table>
    </div>
</div>

"@

    $html += Get-TKHtmlFoot -ScriptName 'V.I.G.I.L. v6.0'
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# REMOTE TARGET CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Connect-RemoteTarget {
    Write-Host ""
    Write-Host -NoNewline "  Enter hostname or IP (leave blank to revert to LOCAL): " -ForegroundColor $C.Header
    $newTarget = (Read-Host).Trim()

    if ([string]::IsNullOrWhiteSpace($newTarget)) {
        $script:RemoteTarget = ''
        Write-Host "  [*] Reverted to LOCAL machine." -ForegroundColor $C.Success
        return
    }

    Write-Host "  [*] Testing WinRM connectivity to $newTarget..." -ForegroundColor $C.Progress
    try {
        Test-WSMan -ComputerName $newTarget -ErrorAction Stop | Out-Null
        $script:RemoteTarget = $newTarget
        Write-Host "  [+] Connected. Target is now: $newTarget" -ForegroundColor $C.Success
    }
    catch {
        Write-Host "  [-] WinRM connection failed: $_" -ForegroundColor $C.Error
        Write-Host "      Ensure WinRM is enabled on target: Enable-PSRemoting -Force" -ForegroundColor $C.Warning
        Write-Host "  [*] Keeping current target: $(if ($script:RemoteTarget) { $script:RemoteTarget } else { 'LOCAL' })" -ForegroundColor $C.Info
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Export-HtmlReport {
    param(
        [array]$Services,
        [array]$Tasks,
        [object]$Events
    )

    $machineName = if ($script:RemoteTarget) { $script:RemoteTarget } else { $env:COMPUTERNAME }
    $timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportName  = "VIGIL_${timestamp}.html"
    $reportPath  = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportName

    Write-Host "  [*] Building HTML report..." -ForegroundColor $C.Progress

    $html = Build-HtmlReport `
        -Services  $Services `
        -Tasks     $Tasks `
        -Events    $Events `
        -MachineName      $machineName `
        -ReportTimestamp  (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $reportPath -Unattended:$Unattended
    }
    catch {
        Write-Host "  [-] Failed to save report: $_" -ForegroundColor $C.Error
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# UNATTENDED MODE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-VigilBanner
    $machineName = if ($script:RemoteTarget) { $script:RemoteTarget } else { $env:COMPUTERNAME }
    Write-Host "  [*] Unattended mode  -  Target: $machineName" -ForegroundColor $C.Progress
    Write-Host ""

    Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
    $services = Get-ServiceHealth

    Write-Host "  [*] Collecting scheduled task audit..." -ForegroundColor $C.Progress
    $tasks = Get-TaskAudit

    Write-Host "  [*] Collecting event log errors (this may take a moment)..." -ForegroundColor $C.Progress
    $events = Get-EventErrors

    Export-HtmlReport -Services $services -Tasks $tasks -Events $events
    Write-Host ""
    Write-Host "  [+] V.I.G.I.L. unattended run complete." -ForegroundColor $C.Success
    Write-Host ""
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE MENU
# ─────────────────────────────────────────────────────────────────────────────

# Pre-load data
$cachedServices = $null
$cachedTasks    = $null
$cachedEvents   = $null

function Invoke-RefreshAll {
    $machineName = if ($script:RemoteTarget) { $script:RemoteTarget } else { $env:COMPUTERNAME }
    Write-Host "  [*] Refreshing all data from: $machineName" -ForegroundColor $C.Progress

    Write-Host "  [*] Querying services..." -ForegroundColor $C.Progress
    $script:cachedServices = Get-ServiceHealth

    Write-Host "  [*] Querying scheduled tasks..." -ForegroundColor $C.Progress
    $script:cachedTasks = Get-TaskAudit

    Write-Host "  [*] Querying event logs (last 24h, up to 50 each log)..." -ForegroundColor $C.Progress
    $script:cachedEvents = Get-EventErrors

    Write-Host "  [+] Refresh complete." -ForegroundColor $C.Success
}

# Initial data load
Show-VigilBanner
Invoke-RefreshAll

$choice = ''

do {
    Show-VigilBanner

    $targetLabel = if ($script:RemoteTarget) { $script:RemoteTarget } else { 'LOCAL' }
    $svcConcern  = if ($cachedServices) { ($cachedServices | Where-Object { $_.Concern }).Count } else { '?' }
    $taskFailed  = if ($cachedTasks)    { ($cachedTasks | Where-Object { $_.FlagReason -like 'Failed*' }).Count } else { '?' }
    $evtCount    = if ($cachedEvents)   { $cachedEvents.TotalCount } else { '?' }

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  V.I.G.I.L. MENU   -   Target: $targetLabel" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""

    $svcColor = if ($svcConcern -gt 0) { $C.Error } else { $C.Success }
    $tskColor = if ($taskFailed -gt 0) { $C.Error } else { $C.Success }
    $evtColor = if ($evtCount -gt 0)   { $C.Warning } else { $C.Success }

    Write-Host -NoNewline "  [1] " -ForegroundColor $C.Header
    Write-Host -NoNewline "Show service health          " -ForegroundColor $C.Info
    Write-Host "Concerns: $svcConcern" -ForegroundColor $svcColor

    Write-Host -NoNewline "  [2] " -ForegroundColor $C.Header
    Write-Host -NoNewline "Show scheduled task audit    " -ForegroundColor $C.Info
    Write-Host "Failed:   $taskFailed" -ForegroundColor $tskColor

    Write-Host -NoNewline "  [3] " -ForegroundColor $C.Header
    Write-Host -NoNewline "Show recent event log errors " -ForegroundColor $C.Info
    Write-Host "Errors:   $evtCount" -ForegroundColor $evtColor

    Write-Host "  [4] Restart stopped critical services (interactive)" -ForegroundColor $C.Warning
    Write-Host "  [5] Export HTML report" -ForegroundColor $C.Info
    Write-Host "  [6] Connect to remote machine" -ForegroundColor $C.Info
    Write-Host "  [R] Refresh all" -ForegroundColor $C.Progress
    Write-Host "  [Q] Quit" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $choice = (Read-Host).Trim().ToUpper()

    switch ($choice) {
        '1' {
            Show-VigilBanner
            if (-not $cachedServices) {
                Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
                $cachedServices = Get-ServiceHealth
            }
            Show-ServiceHealth -Data $cachedServices
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '2' {
            Show-VigilBanner
            if (-not $cachedTasks) {
                Write-Host "  [*] Collecting task audit..." -ForegroundColor $C.Progress
                $cachedTasks = Get-TaskAudit
            }
            Show-TaskAudit -Data $cachedTasks
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '3' {
            Show-VigilBanner
            if (-not $cachedEvents) {
                Write-Host "  [*] Collecting event log errors (this may take a moment)..." -ForegroundColor $C.Progress
                $cachedEvents = Get-EventErrors
            }
            Show-EventErrors -Data $cachedEvents
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '4' {
            Show-VigilBanner
            Write-Section "RESTART STOPPED CRITICAL SERVICES"
            if (-not $cachedServices) {
                Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
                $cachedServices = Get-ServiceHealth
            }
            Invoke-RestartService -ServiceData $cachedServices
            Write-Host ""
            Write-Host "  [*] Refreshing service data after restart..." -ForegroundColor $C.Progress
            $cachedServices = Get-ServiceHealth
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '5' {
            Show-VigilBanner
            Write-Section "EXPORT HTML REPORT"
            if (-not $cachedServices) {
                Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
                $cachedServices = Get-ServiceHealth
            }
            if (-not $cachedTasks) {
                Write-Host "  [*] Collecting task audit..." -ForegroundColor $C.Progress
                $cachedTasks = Get-TaskAudit
            }
            if (-not $cachedEvents) {
                Write-Host "  [*] Collecting event log errors (this may take a moment)..." -ForegroundColor $C.Progress
                $cachedEvents = Get-EventErrors
            }
            Export-HtmlReport -Services $cachedServices -Tasks $cachedTasks -Events $cachedEvents
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '6' {
            Show-VigilBanner
            Connect-RemoteTarget
            # Clear cached data so next action re-fetches from new target
            $cachedServices = $null
            $cachedTasks    = $null
            $cachedEvents   = $null
            if ($script:RemoteTarget) {
                Write-Host ""
                Write-Host "  [*] Loading data from $($script:RemoteTarget)..." -ForegroundColor $C.Progress
                Invoke-RefreshAll
            }
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        'R' {
            Show-VigilBanner
            Write-Section "REFRESH ALL DATA"
            Invoke-RefreshAll
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        'Q' {
            Write-Host ""
            Write-Host "  Closing V.I.G.I.L." -ForegroundColor $C.Header
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  [!!] Invalid selection. Enter 1-6, R, or Q." -ForegroundColor $C.Warning
            Start-Sleep -Seconds 1
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
