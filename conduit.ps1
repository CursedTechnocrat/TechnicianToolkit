# conduit.ps1 - C.O.N.D.U.I.T. — Checks Or Normalises Device Update Infrastructure Targeting
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
    C.O.N.D.U.I.T. — Checks Or Normalises Device Update Infrastructure Targeting
    Windows Update Connectivity Diagnostics & Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Diagnoses and repairs the "We couldn't connect to the update service" failure
    that stops Windows Update before any download begins. Audits the update
    client's plumbing rather than its content: Windows Time service state, the
    WinHTTP proxy, the WindowsUpdate policy key (WSUS pointer and blocking
    policies), whether that policy comes from domain or local Group Policy,
    reachability of the configured WSUS server and the Microsoft Update
    endpoints, and the start type of the four services the client depends on.

    Audit is read-only. Repair backs up the policy key first, then removes a
    stale WSUS pointer only when the WSUS server is unreachable AND the setting
    is not being pushed by domain or local Group Policy, restores the time
    service, and re-enables any disabled update service. ResetCache additionally
    renames SoftwareDistribution and catroot2 so the client rebuilds them.

    CONDUIT is the precondition check for W.H.E.T.S.T.O.N.E.: run CONDUIT
    when the update client cannot reach a service at all, and WHETSTONE once
    it can, to actually deploy updates.

.USAGE
    PS C:\> .\conduit.ps1                                  # Interactive menu
    PS C:\> .\conduit.ps1 -Unattended                      # Read-only audit + HTML report
    PS C:\> .\conduit.ps1 -Unattended -Action Repair       # Audit, then apply safe repairs
    PS C:\> .\conduit.ps1 -Unattended -Action Repair -WhatIf      # Preview the repairs only
    PS C:\> .\conduit.ps1 -Unattended -Action Repair -Force       # Override the WSUS-removal blockers
    PS C:\> .\conduit.ps1 -Unattended -Action ResetCache          # Repair + rebuild the update cache

.NOTES
    Version : 6.1

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit','Repair','ResetCache')]
    [string]$Action = 'Audit',
    [switch]$Force,
    [switch]$Transcript,
    [switch]$WhatIf
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

function Write-TKError {
    <#
    .SYNOPSIS
        Appends a structured error to the toolkit's local error log,
        TK_Errors_<yyyyMM>.jsonl in the log directory. Nothing is sent off the
        machine: the toolkit keeps everything in the terminal or in files the
        technician can remove.
    .EXAMPLE
        Write-TKError -ScriptName 'basilisk' -Message $_.Exception.Message -Category 'Registry'
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][string]$Message,
        [string]$Category = 'General'
    )

    $entry = [PSCustomObject]@{
        Timestamp  = (Get-Date -Format 'o')
        Script     = $ScriptName
        Category   = $Category
        Message    = $Message
        Host       = $env:COMPUTERNAME
        User       = $env:USERNAME
    }

    # Append to monthly JSONL error log in the configured log directory
    try {
        $cfg     = Get-TKConfig
        $logRoot = if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory) -and (Test-Path $cfg.LogDirectory)) {
            $cfg.LogDirectory
        } else {
            $PSScriptRoot
        }
        $logFile = Join-Path $logRoot "TK_Errors_$(Get-Date -Format 'yyyyMM').jsonl"
        $line    = $entry | ConvertTo-Json -Compress
        Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch { <# never throw from a logging helper #> }
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

# Script-scoped note buffer. Lives for the duration of one tool run; the sync
# script copies it only into tools that use the note functions.
$script:TKNotes = New-Object System.Collections.ArrayList

function Add-TKNote {
    <#
    .SYNOPSIS
        Records a timestamped technician note for the current session.
    .DESCRIPTION
        Notes accumulate in script-scoped state for the life of one tool run.
        The tool saves them with Save-TKNoteReport at the end of the run.
    .PARAMETER Text        The note text.
    .PARAMETER Category    One of Info, Action, Warning, Issue, Resolution.
    .PARAMETER ScriptName  Optional originating tool name, shown in the report.
    .EXAMPLE
        Add-TKNote -Text 'Renamed computer to DESKTOP-01' -Category Action -ScriptName 'covenant'
    #>
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Info','Action','Warning','Issue','Resolution')]
        [string]$Category   = 'Info',
        [string]$ScriptName = ''
    )
    [void]$script:TKNotes.Add([PSCustomObject]@{
        Timestamp = Get-Date
        Category  = $Category
        Script    = $ScriptName
        Text      = $Text
    })
}

function Export-TKNoteReport {
    <#
    .SYNOPSIS
        Writes the session's technician notes to a ticket-ready HTML report.
    .DESCRIPTION
        Renders the notes recorded with Add-TKNote into an HTML file built on
        the shared report template. The report includes a plain-text block
        sized for pasting straight into a ticketing system's comment field.
    .PARAMETER Path        Output .html file path.
    .PARAMETER Title       Report heading.
    .PARAMETER ScriptName  Acronym label for the report header/footer.
    .PARAMETER Ticket      Optional ticket / case reference shown in the metadata.
    .PARAMETER Technician  Name of the technician (defaults to the current user).
    .PARAMETER Summary     Optional free-text summary shown above the notes.
    .EXAMPLE
        Export-TKNoteReport -Path 'C:\Reports\notes.html' -Ticket 'INC012345'
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Title      = 'Technician Notes',
        [string]$ScriptName = 'T.K.',
        [string]$Ticket     = '',
        [string]$Technician = $env:USERNAME,
        [string]$Summary    = ''
    )

    $notes     = @($script:TKNotes)
    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $hostName  = $env:COMPUTERNAME

    $badgeClass = @{
        Info       = 'tk-badge-info'
        Action     = 'tk-badge-blue'
        Warning    = 'tk-badge-warn'
        Issue      = 'tk-badge-err'
        Resolution = 'tk-badge-ok'
    }

    $countAction     = @($notes | Where-Object { $_.Category -eq 'Action' }).Count
    $countIssue      = @($notes | Where-Object { $_.Category -eq 'Issue' }).Count
    $countResolution = @($notes | Where-Object { $_.Category -eq 'Resolution' }).Count

    $meta = [ordered]@{
        'Generated'  = $generated
        'Machine'    = $hostName
        'Technician' = $Technician
    }
    if (-not [string]::IsNullOrWhiteSpace($Ticket)) { $meta['Ticket'] = $Ticket }

    $html = Get-TKHtmlHead -Title $Title -ScriptName $ScriptName -Subtitle $hostName `
                -MetaItems $meta -NavItems @('Notes', 'Ticket Summary')

    $html += @"
<div class="tk-summary-row">
  <div class="tk-summary-card info"><div class="tk-summary-num">$($notes.Count)</div><div class="tk-summary-lbl">Total Notes</div></div>
  <div class="tk-summary-card"><div class="tk-summary-num">$countAction</div><div class="tk-summary-lbl">Actions</div></div>
  <div class="tk-summary-card err"><div class="tk-summary-num">$countIssue</div><div class="tk-summary-lbl">Issues</div></div>
  <div class="tk-summary-card ok"><div class="tk-summary-num">$countResolution</div><div class="tk-summary-lbl">Resolutions</div></div>
</div>
"@

    if (-not [string]::IsNullOrWhiteSpace($Summary)) {
        $html += "<div class=`"tk-info-box`"><div class=`"tk-info-label`">Summary</div>$(EscHtml $Summary)</div>"
    }

    $rows = ''
    if ($notes.Count -eq 0) {
        $rows = "<tr><td colspan='4'>No notes were recorded this session.</td></tr>"
    }
    else {
        foreach ($n in $notes) {
            $cls = $badgeClass[[string]$n.Category]
            if (-not $cls) { $cls = 'tk-badge-info' }
            $src = if ($n.Script) { EscHtml $n.Script } else { '&mdash;' }
            $rows += "<tr><td class='tk-mono'>$($n.Timestamp.ToString('HH:mm:ss'))</td>" +
                     "<td><span class='tk-badge $cls'>$(EscHtml ([string]$n.Category))</span></td>" +
                     "<td>$src</td><td>$(EscHtml $n.Text)</td></tr>"
        }
    }

    $html += @"
<div class="tk-section" id="s01">
  <div class="tk-section-title"><span class="tk-section-num">01</span> Notes</div>
  <div class="tk-card">
    <div class="tk-table-wrap">
    <table class="tk-table"><thead><tr><th>Time</th><th>Category</th><th>Source</th><th>Note</th></tr></thead>
    <tbody>$rows</tbody></table>
    </div>
  </div>
</div>
"@

    # Plain-text block — copy/paste straight into a ticket comment field.
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("TECHNICIAN NOTES - $Title")
    [void]$sb.AppendLine("Machine    : $hostName")
    [void]$sb.AppendLine("Technician : $Technician")
    [void]$sb.AppendLine("Generated  : $generated")
    if (-not [string]::IsNullOrWhiteSpace($Ticket)) { [void]$sb.AppendLine("Ticket     : $Ticket") }
    if (-not [string]::IsNullOrWhiteSpace($Summary)) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Summary: $Summary")
    }
    [void]$sb.AppendLine("")
    if ($notes.Count -eq 0) {
        [void]$sb.AppendLine("(no notes recorded)")
    }
    else {
        foreach ($n in $notes) {
            [void]$sb.AppendLine(("[{0}] [{1,-10}] {2}" -f `
                $n.Timestamp.ToString('HH:mm:ss'), ([string]$n.Category).ToUpper(), $n.Text))
        }
    }

    $html += @"
<div class="tk-section" id="s02">
  <div class="tk-section-title"><span class="tk-section-num">02</span> Ticket Summary</div>
  <div class="tk-section-subtitle">Copy the block below straight into the ticket comment field.</div>
  <div class="tk-card">
    <pre class="tk-mono" style="white-space:pre-wrap;display:block;padding:14px;line-height:1.5">$(EscHtml $sb.ToString())</pre>
  </div>
</div>
"@

    $html += Get-TKHtmlFoot -ScriptName $ScriptName

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) {
        $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue
    }
    Set-Content -Path $Path -Value $html -Encoding UTF8
    return $Path
}

function Save-TKNoteReport {
    <#
    .SYNOPSIS
        Saves the run's technician notes beside the tool's reports, if it has any.
    .DESCRIPTION
        The note buffer lives only as long as the tool run, so every tool that
        records notes calls this once, at the end of the run, after its menu
        loop -- one file per run, holding every note the run recorded. Writes
        <Prefix>_Notes_<yyyyMMdd_HHmmss>.html, the name CODEX groups under the
        tool. Returns the path, or $null when there were no notes or the write
        failed; a failure is reported, never thrown, so it cannot end a run that
        has already done its work.
    .PARAMETER Prefix      Tool name used for the file label, e.g. 'CONDUIT'.
    .PARAMETER Directory   Folder to write into -- normally Resolve-LogDirectory.
    .PARAMETER ScriptName  Header label; defaults to the dotted form of Prefix.
    .EXAMPLE
        $null = Save-TKNoteReport -Prefix 'CONDUIT' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)
    #>
    param(
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z]+$')][string]$Prefix,
        [Parameter(Mandatory)][string]$Directory,
        [string]$ScriptName = ''
    )

    if (@($script:TKNotes).Count -eq 0) { return $null }
    if (-not $ScriptName) { $ScriptName = ($Prefix.ToUpper().ToCharArray() -join '.') + '.' }

    $path = Join-Path $Directory ('{0}_Notes_{1}.html' -f $Prefix.ToUpper(), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        $null = Export-TKNoteReport -Path $path -ScriptName $ScriptName -ErrorAction Stop
        Write-Ok "Technician notes saved:"
        Write-Info $path
        return $path
    }
    catch {
        Write-Warn "Could not save technician notes: $($_.Exception.Message)"
        return $null
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
$C = $ColorSchema

# ─────────────────────────────────────────────────────────────────────────────
# REGISTRY PATHS
# ─────────────────────────────────────────────────────────────────────────────

$WUPolicyKey   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$AUPolicyKey   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
$WUPolicyRegEx = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$LocalPolFile  = Join-Path $env:SystemRoot 'System32\GroupPolicy\Machine\Registry.pol'

# The four services the update client cannot function without, mapped to the
# start type Windows ships them with. Used both to flag a disabled service and
# to restore the correct value -- setting them all to Automatic would be wrong,
# wuauserv and bits are demand-started by design.
$UpdateServiceDefaults = [ordered]@{
    'wuauserv' = 'Manual'
    'bits'     = 'Manual'
    'cryptsvc' = 'Automatic'
    'usosvc'   = 'Automatic'
}

# TCP endpoints the update client reaches when no WSUS server is configured.
# A block on any one of these is a network/filter problem, not a client one.
$MicrosoftUpdateEndpoints = @(
    'update.microsoft.com'
    'windowsupdate.microsoft.com'
    'dl.delivery.mp.microsoft.com'
)

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition CONDUIT can detect, keyed by a stable code. Severity drives
# both the console colour and the report badge class; Remedy is what the
# technician is told to do when CONDUIT will not do it automatically. Kept as
# one table so the Pester suite can extract and assert on it by AST lookup
# without executing the script.
# ─────────────────────────────────────────────────────────────────────────────

$ConduitFindings = @{
    'TimeServiceStopped' = @{
        Severity = 'Warning'
        Title    = 'Windows Time service not running'
        Summary  = 'Clock drift beyond five minutes breaks the TLS handshake to the update service.'
        Remedy   = 'Repair starts w32time, sets it to Automatic, and forces a resync.'
    }
    'WinHttpProxySet' = @{
        Severity = 'Warning'
        Title    = 'WinHTTP proxy configured'
        Summary  = 'The update client uses the WinHTTP proxy, not the per-user Internet Options proxy. A stale entry here blackholes every update request.'
        Remedy   = 'Confirm the proxy is intentional. Clear it manually with: netsh winhttp reset proxy'
    }
    'WsusUnreachable' = @{
        Severity = 'Error'
        Title    = 'WSUS server configured but unreachable'
        Summary  = 'The client is pointed at a WSUS server that does not resolve or does not answer on its port, so it never falls back to Microsoft Update.'
        Remedy   = 'Repair removes the pointer when no Group Policy is reapplying it; otherwise clear the policy at source.'
    }
    'BlockInternetWU' = @{
        Severity = 'Error'
        Title    = 'Policy blocks Microsoft Update endpoints'
        Summary  = 'DoNotConnectToWindowsUpdateInternetLocations = 1 forbids the client from contacting Microsoft Update even when WSUS is absent.'
        Remedy   = 'Clear the policy at its source (domain GPO or gpedit.msc). CONDUIT never edits this value automatically.'
    }
    'WUAccessDisabled' = @{
        Severity = 'Error'
        Title    = 'Windows Update access disabled by policy'
        Summary  = 'DisableWindowsUpdateAccess = 1 removes the update UI and blocks user-initiated scans.'
        Remedy   = 'Clear the policy at its source (domain GPO or gpedit.msc). CONDUIT never edits this value automatically.'
    }
    'AutoUpdateDisabled' = @{
        Severity = 'Warning'
        Title    = 'Automatic updates disabled by policy'
        Summary  = 'NoAutoUpdate = 1 stops scheduled scans. Manual checks still work, so this is context rather than the connection fault itself.'
        Remedy   = 'Expected on servers and maintenance-window builds. Clear at source if unintended.'
    }
    'MicrosoftEndpointsBlocked' = @{
        Severity = 'Error'
        Title    = 'Microsoft Update endpoints unreachable'
        Summary  = 'One or more Microsoft Update hosts refused a TCP connection on 443 -- a network, firewall, or content-filter block upstream of this machine.'
        Remedy   = 'Allow the endpoints through the filter. No local repair can fix an upstream block.'
    }
    'ServiceDisabled' = @{
        Severity = 'Error'
        Title    = 'Update service disabled'
        Summary  = 'One of wuauserv / bits / cryptsvc / usosvc has start type Disabled, so the update client cannot start it on demand.'
        Remedy   = 'Repair restores the service to its Windows default start type.'
    }
    'WUServerUnparsable' = @{
        Severity = 'Error'
        Title    = 'WSUS pointer is not a valid URL'
        Summary  = 'The WUServer policy value could not be parsed as a URI, so the client cannot build a request from it.'
        Remedy   = 'Correct or remove the WUServer value at its policy source.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[object]]::new()

function Add-ConduitFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Detail = ''
    )
    $meta = $ConduitFindings[$Code]
    if (-not $meta) {
        $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' }
    }
    [void]$Findings.Add([PSCustomObject]@{
        Code     = $Code
        Severity = $meta.Severity
        Title    = $meta.Title
        Summary  = $meta.Summary
        Remedy   = $meta.Remedy
        Detail   = $Detail
    })
}

function Add-ConduitAction {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,
        [string]$Detail = ''
    )
    [void]$Actions.Add([PSCustomObject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Step      = $Step
        Status    = $Status
        Detail    = $Detail
    })
}

function Get-SeverityClass {
    param([string]$Severity)
    switch ($Severity) {
        'Error'   { return 'err'  }
        'Warning' { return 'warn' }
        default   { return 'info' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-ConduitBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██████╗ ███╗   ██╗██████╗ ██╗   ██╗██╗████████╗
  ██╔════╝██╔═══██╗████╗  ██║██╔══██╗██║   ██║██║╚══██╔══╝
  ██║     ██║   ██║██╔██╗ ██║██║  ██║██║   ██║██║   ██║
  ██║     ██║   ██║██║╚██╗██║██║  ██║██║   ██║██║   ██║
  ╚██████╗╚██████╔╝██║ ╚████║██████╔╝╚██████╔╝██║   ██║
   ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝╚═════╝  ╚═════╝ ╚═╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    C.O.N.D.U.I.T. — Checks Or Normalises Device Update Infrastructure Targeting" -ForegroundColor Cyan
    Write-Host "    Windows Update Connectivity Diagnostics & Repair Tool" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host ""
        Write-Host "    *** DRY RUN — no changes will be applied ***" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-TCPPort {
    # Test-NetConnection takes ~20s to fail closed; a raw socket with an explicit
    # timeout keeps a six-endpoint sweep inside a few seconds.
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 3000)
    $tcp = $null
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok) { $tcp.EndConnect($ar) | Out-Null }
        return ($ok -and $tcp.Connected)
    } catch {
        return $false
    } finally {
        if ($tcp) { $tcp.Close() }
    }
}

function Get-ServiceSafe {
    param([string]$Name)
    try { return Get-Service -Name $Name -ErrorAction Stop }
    catch { return $null }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-ConduitDeviceContext {
    # Context only -- a CIM failure must not abort the diagnosis, since every
    # check that actually matters reads the registry and the service table.
    $cs = $null; $os = $null
    try { $cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop } catch { $cs = $null }
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }

    return [PSCustomObject]@{
        Computer     = $env:COMPUTERNAME
        OS           = if ($os) { "$($os.Caption) (build $($os.BuildNumber))" } else { 'Unknown' }
        Domain       = if ($cs) { $cs.Domain } else { 'Unknown' }
        PartOfDomain = if ($cs) { [bool]$cs.PartOfDomain } else { $false }
    }
}

function Get-ConduitTimeService {
    $svc = Get-ServiceSafe -Name 'w32time'
    if (-not $svc) {
        return [PSCustomObject]@{ Present = $false; Status = 'Absent'; StartType = 'Absent' }
    }
    if ($svc.Status -ne 'Running') { Add-ConduitFinding -Code 'TimeServiceStopped' -Detail "Current status: $($svc.Status)" }
    return [PSCustomObject]@{
        Present   = $true
        Status    = "$($svc.Status)"
        StartType = "$($svc.StartType)"
    }
}

function Get-ConduitProxy {
    $raw = ''
    try { $raw = (netsh winhttp show proxy) -join [Environment]::NewLine } catch { $raw = 'Could not query WinHTTP proxy.' }

    # "Direct access (no proxy server)" is the clean state. Anything else means a
    # proxy is in play for the update client specifically.
    $configured = ($raw -notmatch 'Direct access')
    if ($configured) { Add-ConduitFinding -Code 'WinHttpProxySet' -Detail (($raw -split "`r?`n" | Where-Object { $_.Trim() }) -join ' | ') }

    return [PSCustomObject]@{
        Configured = $configured
        Raw        = $raw.Trim()
    }
}

function Get-ConduitPolicy {
    $wu = Get-ItemProperty -Path $WUPolicyKey -ErrorAction SilentlyContinue
    $au = Get-ItemProperty -Path $AUPolicyKey -ErrorAction SilentlyContinue

    $wsusConfigured = (-not [string]::IsNullOrWhiteSpace($wu.WUServer)) -and ($au.UseWUServer -eq 1)

    if ($wu.DoNotConnectToWindowsUpdateInternetLocations -eq 1) { Add-ConduitFinding -Code 'BlockInternetWU' }
    if ($wu.DisableWindowsUpdateAccess -eq 1)                   { Add-ConduitFinding -Code 'WUAccessDisabled' }
    if ($au.NoAutoUpdate -eq 1)                                 { Add-ConduitFinding -Code 'AutoUpdateDisabled' }

    return [PSCustomObject]@{
        KeyPresent     = [bool](Test-Path $WUPolicyKey)
        WUServer       = "$($wu.WUServer)"
        WUStatusServer = "$($wu.WUStatusServer)"
        UseWUServer    = $au.UseWUServer
        BlockInternet  = $wu.DoNotConnectToWindowsUpdateInternetLocations
        AccessDisabled = $wu.DisableWindowsUpdateAccess
        NoAutoUpdate   = $au.NoAutoUpdate
        WsusConfigured = $wsusConfigured
    }
}

function Get-ConduitPolicySource {
    # Distinguishing "someone wrote the registry once" from "Group Policy pushes
    # this every 90 minutes" is the whole basis for deciding whether removing the
    # WSUS pointer will stick. Registry.pol is Unicode; Select-String needs telling.
    param([bool]$PartOfDomain)

    $localGpoSetsWsus = $false
    if (Test-Path $LocalPolFile) {
        try {
            $localGpoSetsWsus = [bool](Select-String -Path $LocalPolFile -Pattern 'WUServer' -Encoding Unicode -Quiet -ErrorAction Stop)
        } catch {
            $localGpoSetsWsus = $false
        }
    }

    return [PSCustomObject]@{
        LocalGpoSetsWsus = $localGpoSetsWsus
        DomainJoined     = $PartOfDomain
        LocalPolPresent  = (Test-Path $LocalPolFile)
    }
}

function Get-ConduitWsusReachability {
    param([object]$Policy)

    if (-not $Policy.WsusConfigured) {
        return [PSCustomObject]@{ Configured = $false; HostName = ''; Port = 0; DnsOk = $false; Reachable = $null }
    }

    $uri = $null
    try { $uri = [uri]$Policy.WUServer } catch { $uri = $null }
    if (-not $uri -or [string]::IsNullOrWhiteSpace($uri.Host)) {
        Add-ConduitFinding -Code 'WUServerUnparsable' -Detail "WUServer = $($Policy.WUServer)"
        return [PSCustomObject]@{ Configured = $true; HostName = $Policy.WUServer; Port = 0; DnsOk = $false; Reachable = $false }
    }

    $port  = if ($uri.Port -gt 0) { $uri.Port } else { 8530 }
    $dnsOk = $false
    try { $dnsOk = [bool](Resolve-DnsName -Name $uri.Host -ErrorAction Stop) } catch { $dnsOk = $false }

    $reachable = $dnsOk -and (Test-TCPPort -Hostname $uri.Host -Port $port)
    if (-not $reachable) {
        Add-ConduitFinding -Code 'WsusUnreachable' -Detail ("{0}:{1} (DNS={2})" -f $uri.Host, $port, $dnsOk)
    }

    return [PSCustomObject]@{
        Configured = $true
        HostName   = $uri.Host
        Port       = $port
        DnsOk      = $dnsOk
        Reachable  = $reachable
    }
}

function Get-ConduitEndpointReachability {
    $results = foreach ($ep in $MicrosoftUpdateEndpoints) {
        [PSCustomObject]@{
            Endpoint  = $ep
            Port      = 443
            Reachable = (Test-TCPPort -Hostname $ep -Port 443)
        }
    }
    $results = @($results)

    $blocked = @($results | Where-Object { -not $_.Reachable })
    if ($blocked.Count -gt 0) {
        Add-ConduitFinding -Code 'MicrosoftEndpointsBlocked' -Detail (($blocked | ForEach-Object { $_.Endpoint }) -join ', ')
    }
    return $results
}

function Get-ConduitUpdateServices {
    $results = foreach ($name in $UpdateServiceDefaults.Keys) {
        $svc = Get-ServiceSafe -Name $name
        if (-not $svc) {
            [PSCustomObject]@{ Name = $name; Present = $false; Status = 'Absent'; StartType = 'Absent'; Expected = $UpdateServiceDefaults[$name] }
            continue
        }
        if ("$($svc.StartType)" -eq 'Disabled') {
            Add-ConduitFinding -Code 'ServiceDisabled' -Detail "$name is Disabled (expected $($UpdateServiceDefaults[$name]))"
        }
        [PSCustomObject]@{
            Name      = $name
            Present   = $true
            Status    = "$($svc.Status)"
            StartType = "$($svc.StartType)"
            Expected  = $UpdateServiceDefaults[$name]
        }
    }
    return @($results)
}

# ─────────────────────────────────────────────────────────────────────────────
# AUDIT — runs every collector and prints the console summary
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ConduitAudit {
    Write-Section "DEVICE CONTEXT"
    $device = Get-ConduitDeviceContext
    Write-Info ("Computer      : {0}" -f $device.Computer)
    Write-Info ("OS            : {0}" -f $device.OS)
    Write-Info ("Domain        : {0}" -f $device.Domain)
    Write-Info ("Domain-joined : {0}" -f $device.PartOfDomain)

    Write-Section "WINDOWS TIME SERVICE"
    $time = Get-ConduitTimeService
    if (-not $time.Present) {
        Write-Fail "w32time service not present."
    } elseif ($time.Status -eq 'Running') {
        Write-Ok ("w32time is {0} (start type {1})." -f $time.Status, $time.StartType)
    } else {
        Write-Warn ("w32time is {0} (start type {1})." -f $time.Status, $time.StartType)
    }

    Write-Section "WINHTTP PROXY"
    $proxy = Get-ConduitProxy
    foreach ($line in ($proxy.Raw -split "`r?`n")) {
        if ($line.Trim()) { Write-Info $line.Trim() }
    }
    if ($proxy.Configured) {
        Write-Warn "A WinHTTP proxy is configured — confirm it is intentional."
    } else {
        Write-Ok "Direct access — no WinHTTP proxy."
    }

    Write-Section "WINDOWS UPDATE POLICY"
    $policy = Get-ConduitPolicy
    if (-not $policy.KeyPresent) {
        Write-Ok "No WindowsUpdate policy key — the client uses its defaults."
    } else {
        Write-Info ("WUServer                                     : {0}" -f $(if ($policy.WUServer) { $policy.WUServer } else { '(not set)' }))
        Write-Info ("WUStatusServer                               : {0}" -f $(if ($policy.WUStatusServer) { $policy.WUStatusServer } else { '(not set)' }))
        Write-Info ("UseWUServer                                  : {0}" -f $(if ($null -ne $policy.UseWUServer) { $policy.UseWUServer } else { '(not set)' }))
        Write-Info ("DoNotConnectToWindowsUpdateInternetLocations : {0}" -f $(if ($null -ne $policy.BlockInternet) { $policy.BlockInternet } else { '(not set)' }))
        Write-Info ("DisableWindowsUpdateAccess                   : {0}" -f $(if ($null -ne $policy.AccessDisabled) { $policy.AccessDisabled } else { '(not set)' }))
        Write-Info ("NoAutoUpdate                                 : {0}" -f $(if ($null -ne $policy.NoAutoUpdate) { $policy.NoAutoUpdate } else { '(not set)' }))
    }

    Write-Section "POLICY SOURCE"
    $source = Get-ConduitPolicySource -PartOfDomain $device.PartOfDomain
    Write-Info ("Local Group Policy sets WUServer : {0}" -f $source.LocalGpoSetsWsus)
    Write-Info ("Domain GPO possible              : {0}" -f $source.DomainJoined)
    if ($source.DomainJoined -and $policy.WsusConfigured) {
        Write-Info "Tip: gpresult /h C:\temp\gp.html — find the GPO that sets WUServer."
    }

    Write-Section "CONNECTIVITY"
    $wsus = Get-ConduitWsusReachability -Policy $policy
    if (-not $wsus.Configured) {
        Write-Ok "No WSUS server configured — the client targets Microsoft Update directly."
    } elseif ($wsus.Reachable) {
        Write-Ok ("WSUS {0}:{1} reachable (DNS={2})." -f $wsus.HostName, $wsus.Port, $wsus.DnsOk)
    } else {
        Write-Fail ("WSUS {0}:{1} unreachable (DNS={2})." -f $wsus.HostName, $wsus.Port, $wsus.DnsOk)
    }

    $endpoints = Get-ConduitEndpointReachability
    foreach ($ep in $endpoints) {
        if ($ep.Reachable) { Write-Ok   ("{0}:{1} reachable"   -f $ep.Endpoint, $ep.Port) }
        else               { Write-Fail ("{0}:{1} unreachable" -f $ep.Endpoint, $ep.Port) }
    }

    Write-Section "UPDATE SERVICES"
    $services = Get-ConduitUpdateServices
    foreach ($svc in $services) {
        $line = "{0,-10} {1,-10} start={2} (expected {3})" -f $svc.Name, $svc.Status, $svc.StartType, $svc.Expected
        if (-not $svc.Present)                { Write-Fail ("{0,-10} not present" -f $svc.Name) }
        elseif ($svc.StartType -eq 'Disabled'){ Write-Fail $line }
        else                                  { Write-Ok   $line }
    }

    return [PSCustomObject]@{
        Device    = $device
        Time      = $time
        Proxy     = $proxy
        Policy    = $policy
        Source    = $source
        Wsus      = $wsus
        Endpoints = $endpoints
        Services  = $services
    }
}

function Show-ConduitFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) {
        Write-Ok "No issues found — the update client's plumbing looks healthy."
        return
    }
    foreach ($f in $Findings) {
        switch ($f.Severity) {
            'Error'   { Write-Fail ("[{0}] {1}" -f $f.Code, $f.Title) }
            'Warning' { Write-Warn ("[{0}] {1}" -f $f.Code, $f.Title) }
            default   { Write-Info ("[{0}] {1}" -f $f.Code, $f.Title) }
        }
        if ($f.Detail) { Write-Info $f.Detail }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Backup-ConduitPolicy {
    Write-Section "REPAIR — POLICY BACKUP"

    if (-not (Test-Path $WUPolicyKey)) {
        Write-Info "No WindowsUpdate policy key to back up."
        Add-ConduitAction -Step 'Backup policy key' -Status 'Skipped' -Detail 'Key not present'
        return ''
    }

    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backup = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "CONDUIT_WUPolicy_$stamp.reg"

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would export $WUPolicyRegEx to $backup"
        Add-ConduitAction -Step 'Backup policy key' -Status 'WhatIf' -Detail $backup
        return ''
    }

    try {
        $null = reg export $WUPolicyRegEx $backup /y 2>&1
        if ($LASTEXITCODE -ne 0) { throw "reg export returned exit code $LASTEXITCODE" }
        Write-Ok "Policy key backed up."
        Write-Info $backup
        Write-Info "Restore with: reg import `"$backup`""
        Add-ConduitAction -Step 'Backup policy key' -Status 'Done' -Detail $backup
        return $backup
    } catch {
        Write-Fail "Could not back up the policy key: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "Policy key backup failed: $($_.Exception.Message)" -Category 'Registry'
        Add-ConduitAction -Step 'Backup policy key' -Status 'Failed' -Detail $_.Exception.Message
        return ''
    }
}

function Repair-ConduitWsusPointer {
    param([object]$Audit)

    Write-Section "REPAIR — WSUS POINTER"

    if (-not $Audit.Policy.WsusConfigured) {
        Write-Ok "No WSUS pointer present. Nothing to do."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Skipped' -Detail 'No pointer configured'
        return
    }

    # Three independent reasons not to touch the pointer. Each means removal is
    # either destructive (the server is live) or futile (policy reapplies it).
    $blockers = @()
    if ($Audit.Wsus.Reachable)          { $blockers += 'WSUS server is reachable (it may be in active use)' }
    if ($Audit.Device.PartOfDomain)     { $blockers += 'device is domain-joined (a domain GPO will reapply it)' }
    if ($Audit.Source.LocalGpoSetsWsus) { $blockers += 'local Group Policy sets it (clear it in gpedit.msc instead)' }

    if ($blockers.Count -gt 0 -and -not $Force) {
        Write-Warn ("Skipped WSUS removal: {0}" -f ($blockers -join '; '))
        Write-Info "Re-run with -Force to override."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Skipped' -Detail ($blockers -join '; ')
        return
    }

    if ($blockers.Count -gt 0) {
        Write-Warn ("-Force used despite: {0}" -f ($blockers -join '; '))
    }

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would remove UseWUServer, WUServer and WUStatusServer, then restart wuauserv."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'WhatIf' -Detail $Audit.Policy.WUServer
        return
    }

    try {
        Remove-ItemProperty -Path $AUPolicyKey -Name 'UseWUServer' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $WUPolicyKey -Name 'WUServer', 'WUStatusServer' -ErrorAction SilentlyContinue
        Restart-Service -Name 'wuauserv' -Force -ErrorAction Stop
        Write-Ok "Removed the WSUS pointer and restarted wuauserv."
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Done' -Detail "Was: $($Audit.Policy.WUServer)"
        Add-TKNote -Text "Removed unreachable WSUS pointer $($Audit.Policy.WUServer) and restarted wuauserv." -Category 'Action'
    } catch {
        Write-Fail "Could not remove the WSUS pointer: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "WSUS pointer removal failed: $($_.Exception.Message)" -Category 'Registry'
        Add-ConduitAction -Step 'Remove WSUS pointer' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Repair-ConduitTimeService {
    Write-Section "REPAIR — WINDOWS TIME SERVICE"

    $svc = Get-ServiceSafe -Name 'w32time'
    if (-not $svc) {
        Write-Fail "w32time is not present on this machine."
        Add-ConduitAction -Step 'Repair time service' -Status 'Failed' -Detail 'Service absent'
        return
    }

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would set w32time to Automatic, start it, and run: w32tm /resync /force"
        Add-ConduitAction -Step 'Repair time service' -Status 'WhatIf' -Detail "Current: $($svc.Status) / $($svc.StartType)"
        return
    }

    try {
        Set-Service -Name 'w32time' -StartupType Automatic -ErrorAction Stop
        if ((Get-Service -Name 'w32time').Status -ne 'Running') {
            Start-Service -Name 'w32time' -ErrorAction Stop
            Start-Sleep -Seconds 3
        }
        $resync = & w32tm /resync /force 2>&1
        Write-Ok "w32time set to Automatic, running, and resynced."
        Write-Info (($resync | Out-String).Trim())
        Add-ConduitAction -Step 'Repair time service' -Status 'Done' -Detail (($resync | Out-String).Trim())
        Add-TKNote -Text 'Enabled and resynced the Windows Time service.' -Category 'Action'
    } catch {
        Write-Fail "Could not repair the time service: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "Time service repair failed: $($_.Exception.Message)" -Category 'Service'
        Add-ConduitAction -Step 'Repair time service' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Repair-ConduitServices {
    Write-Section "REPAIR — UPDATE SERVICES"

    $changed = $false
    foreach ($name in $UpdateServiceDefaults.Keys) {
        $svc = Get-ServiceSafe -Name $name
        if (-not $svc) {
            Write-Warn "$name is not present — skipping."
            continue
        }
        if ("$($svc.StartType)" -ne 'Disabled') { continue }

        $expected = $UpdateServiceDefaults[$name]
        $changed  = $true

        if ($WhatIf) {
            Write-Warn "[WhatIf] Would set $name start type from Disabled to $expected."
            Add-ConduitAction -Step "Re-enable $name" -Status 'WhatIf' -Detail "Disabled -> $expected"
            continue
        }

        try {
            Set-Service -Name $name -StartupType $expected -ErrorAction Stop
            Write-Ok "$name start type set to $expected."
            Add-ConduitAction -Step "Re-enable $name" -Status 'Done' -Detail "Disabled -> $expected"
            Add-TKNote -Text "Re-enabled the $name service (Disabled -> $expected)." -Category 'Action'
        } catch {
            Write-Fail "Could not set $name start type: $($_.Exception.Message)"
            Write-TKError -ScriptName 'conduit' -Message "Service start type change failed for ${name}: $($_.Exception.Message)" -Category 'Service'
            Add-ConduitAction -Step "Re-enable $name" -Status 'Failed' -Detail $_.Exception.Message
        }
    }

    if (-not $changed) {
        Write-Ok "No update service is disabled. Nothing to do."
        Add-ConduitAction -Step 'Re-enable update services' -Status 'Skipped' -Detail 'None disabled'
    }
}

function Reset-ConduitUpdateCache {
    Write-Section "REPAIR — UPDATE CACHE RESET"

    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $targets = @(
        (Join-Path $env:SystemRoot 'SoftwareDistribution')
        (Join-Path $env:SystemRoot 'System32\catroot2')
    )

    if ($WhatIf) {
        Write-Warn "[WhatIf] Would stop wuauserv, bits, cryptsvc, usosvc."
        foreach ($t in $targets) {
            if (Test-Path $t) { Write-Warn "[WhatIf] Would rename $t to $(Split-Path $t -Leaf).old_$stamp" }
        }
        Write-Warn "[WhatIf] Would restart cryptsvc, bits, wuauserv, usosvc."
        Add-ConduitAction -Step 'Reset update cache' -Status 'WhatIf' -Detail ($targets -join '; ')
        return
    }

    Write-Step "Stopping update services..."
    Stop-Service -Name 'wuauserv', 'bits', 'cryptsvc', 'usosvc' -Force -ErrorAction SilentlyContinue

    foreach ($t in $targets) {
        if (-not (Test-Path $t)) {
            Write-Info "$t not present — skipping."
            continue
        }
        $newName = "$(Split-Path $t -Leaf).old_$stamp"
        try {
            Rename-Item -Path $t -NewName $newName -ErrorAction Stop
            Write-Ok "Renamed $t to $newName."
            Add-ConduitAction -Step 'Reset update cache' -Status 'Done' -Detail "$t -> $newName"
            Add-TKNote -Text "Renamed $t to $newName so the update client rebuilds it." -Category 'Action'
        } catch {
            Write-Fail "Could not rename ${t}: $($_.Exception.Message)"
            Write-TKError -ScriptName 'conduit' -Message "Cache rename failed for ${t}: $($_.Exception.Message)" -Category 'Filesystem'
            Add-ConduitAction -Step 'Reset update cache' -Status 'Failed' -Detail $_.Exception.Message
        }
    }

    Write-Step "Restarting update services..."
    Start-Service -Name 'cryptsvc', 'bits', 'wuauserv', 'usosvc' -ErrorAction SilentlyContinue
    Write-Ok "Update services restarted."
}

function Start-ConduitScan {
    # Nudges the client to re-scan so the Settings page reflects the repair
    # rather than showing the cached failure until the next scheduled sweep.
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would trigger UsoClient StartInteractiveScan."
        return
    }
    try {
        Start-Process -FilePath 'UsoClient.exe' -ArgumentList 'StartInteractiveScan' -WindowStyle Hidden -ErrorAction Stop
        Write-Ok "Triggered a Windows Update scan."
    } catch {
        Write-Info "Could not trigger a scan automatically — open Settings > Windows Update and check manually."
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Get-ConduitVerdict {
    $errors   = @($Findings | Where-Object { $_.Severity -eq 'Error' })
    $warnings = @($Findings | Where-Object { $_.Severity -eq 'Warning' })

    if ($errors.Count -gt 0)   { return [PSCustomObject]@{ Verdict = 'Blocked';   Class = 'err'  } }
    if ($warnings.Count -gt 0) { return [PSCustomObject]@{ Verdict = 'Degraded';  Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

function Build-ConduitReport {
    param([object]$Audit, [object]$Verdict)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $errorCount = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount  = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count

    # Findings rows
    $fRows = New-Object System.Text.StringBuilder
    if ($Findings.Count -eq 0) {
        [void]$fRows.Append("<tr><td colspan='4'>No issues found.</td></tr>")
    } else {
        foreach ($f in $Findings) {
            $badge = Get-SeverityClass -Severity $f.Severity
            [void]$fRows.Append(
                "<tr><td><span class='tk-badge-$badge'>$(EscHtml $f.Severity)</span></td>" +
                "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
                "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
                $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
                "<td>$(EscHtml $f.Remedy)</td></tr>"
            )
        }
    }

    # Policy rows
    $policyPairs = [ordered]@{
        'WUServer'                                     = $(if ($Audit.Policy.WUServer)       { $Audit.Policy.WUServer }       else { '(not set)' })
        'WUStatusServer'                               = $(if ($Audit.Policy.WUStatusServer) { $Audit.Policy.WUStatusServer } else { '(not set)' })
        'UseWUServer'                                  = $(if ($null -ne $Audit.Policy.UseWUServer)    { "$($Audit.Policy.UseWUServer)" }    else { '(not set)' })
        'DoNotConnectToWindowsUpdateInternetLocations' = $(if ($null -ne $Audit.Policy.BlockInternet)  { "$($Audit.Policy.BlockInternet)" }  else { '(not set)' })
        'DisableWindowsUpdateAccess'                   = $(if ($null -ne $Audit.Policy.AccessDisabled) { "$($Audit.Policy.AccessDisabled)" } else { '(not set)' })
        'NoAutoUpdate'                                 = $(if ($null -ne $Audit.Policy.NoAutoUpdate)   { "$($Audit.Policy.NoAutoUpdate)" }   else { '(not set)' })
    }
    $pRows = New-Object System.Text.StringBuilder
    foreach ($k in $policyPairs.Keys) {
        [void]$pRows.Append("<tr><td class='tk-mono'>$(EscHtml $k)</td><td>$(EscHtml $policyPairs[$k])</td></tr>")
    }

    # Endpoint rows
    $eRows = New-Object System.Text.StringBuilder
    if ($Audit.Wsus.Configured) {
        $wsusBadge = if ($Audit.Wsus.Reachable) { 'ok' } else { 'err' }
        $wsusText  = if ($Audit.Wsus.Reachable) { 'Reachable' } else { 'Unreachable' }
        [void]$eRows.Append("<tr><td class='tk-mono'>$(EscHtml $Audit.Wsus.HostName)</td><td>$($Audit.Wsus.Port)</td><td>WSUS</td><td>$(EscHtml "$($Audit.Wsus.DnsOk)")</td><td><span class='tk-badge-$wsusBadge'>$wsusText</span></td></tr>")
    }
    foreach ($ep in $Audit.Endpoints) {
        $badge = if ($ep.Reachable) { 'ok' } else { 'err' }
        $text  = if ($ep.Reachable) { 'Reachable' } else { 'Unreachable' }
        [void]$eRows.Append("<tr><td class='tk-mono'>$(EscHtml $ep.Endpoint)</td><td>$($ep.Port)</td><td>Microsoft Update</td><td>-</td><td><span class='tk-badge-$badge'>$text</span></td></tr>")
    }

    # Service rows
    $sRows = New-Object System.Text.StringBuilder
    foreach ($svc in $Audit.Services) {
        $badge = if (-not $svc.Present) { 'err' } elseif ($svc.StartType -eq 'Disabled') { 'err' } else { 'ok' }
        [void]$sRows.Append(
            "<tr><td class='tk-mono'>$(EscHtml $svc.Name)</td><td>$(EscHtml $svc.Status)</td>" +
            "<td><span class='tk-badge-$badge'>$(EscHtml $svc.StartType)</span></td><td>$(EscHtml $svc.Expected)</td></tr>"
        )
    }

    # Action rows
    $aRows = New-Object System.Text.StringBuilder
    if ($Actions.Count -eq 0) {
        [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>")
    } else {
        foreach ($a in $Actions) {
            $badge = switch ($a.Status) {
                'Done'   { 'ok'   }
                'WhatIf' { 'blue' }
                'Failed' { 'err'  }
                default  { 'info' }
            }
            [void]$aRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td>" +
                "<td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>"
            )
        }
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.O.N.D.U.I.T. Windows Update Connectivity Report' `
        -ScriptName 'C.O.N.D.U.I.T.' `
        -Subtitle   "${orgPrefix}Windows Update Connectivity -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'      = $machine
            'Generated'    = $reportDate
            'Mode'         = $(if ($WhatIf) { "$Action (dry run)" } else { $Action })
            'Verdict'      = $Verdict.Verdict
            'Domain'       = $Audit.Device.Domain
            'WSUS'         = $(if ($Audit.Policy.WsusConfigured) { $Audit.Policy.WUServer } else { 'Not configured' })
        }) `
        -NavItems   @('Findings', 'Device & Policy', 'Connectivity', 'Update Services', 'Actions Taken')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'C.O.N.D.U.I.T. v6.1'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Update Connectivity</div></div>
    <div class="tk-summary-card $(if ($errorCount -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errorCount</div><div class="tk-summary-lbl">Blocking Issues</div></div>
    <div class="tk-summary-card $(if ($warnCount -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(if ($Audit.Policy.WsusConfigured) { 'Yes' } else { 'No' })</div><div class="tk-summary-lbl">WSUS Configured</div></div>
    <div class="tk-summary-card $(if ($Audit.Proxy.Configured) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$(if ($Audit.Proxy.Configured) { 'Set' } else { 'Direct' })</div><div class="tk-summary-lbl">WinHTTP Proxy</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Actions.Count)</div><div class="tk-summary-lbl">Actions Recorded</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
        <tbody>$($fRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Device &amp; Policy</div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Computer</span> $(EscHtml $Audit.Device.Computer)<br/>
        <span class="tk-info-label">Operating system</span> $(EscHtml $Audit.Device.OS)<br/>
        <span class="tk-info-label">Domain</span> $(EscHtml $Audit.Device.Domain) (joined: $(EscHtml "$($Audit.Device.PartOfDomain)"))<br/>
        <span class="tk-info-label">Local GPO sets WUServer</span> $(EscHtml "$($Audit.Source.LocalGpoSetsWsus)")<br/>
        <span class="tk-info-label">Windows Time</span> $(EscHtml $Audit.Time.Status) (start type $(EscHtml $Audit.Time.StartType))<br/>
        <span class="tk-info-label">WinHTTP proxy</span> $(EscHtml $Audit.Proxy.Raw)
      </div>
      <table class="tk-table">
        <thead><tr><th>Policy value</th><th>Setting</th></tr></thead>
        <tbody>$($pRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Connectivity</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Endpoint</th><th>Port</th><th>Role</th><th>DNS</th><th>Result</th></tr></thead>
        <tbody>$($eRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Update Services</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Service</th><th>Status</th><th>Start type</th><th>Windows default</th></tr></thead>
        <tbody>$($sRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Actions Taken</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
        <tbody>$($aRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ConduitRepair {
    param([object]$Audit)

    Backup-ConduitPolicy | Out-Null
    Repair-ConduitWsusPointer -Audit $Audit
    Repair-ConduitTimeService
    Repair-ConduitServices
}

function Invoke-ConduitRun {
    param([string]$Mode)

    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-ConduitAudit
    Show-ConduitFindings

    if ($Mode -ne 'Audit') {
        Invoke-ConduitRepair -Audit $audit
        if ($Mode -eq 'ResetCache') { Reset-ConduitUpdateCache }
        Start-ConduitScan

        Write-Section "VERIFY"
        # Re-read the live state rather than trusting the repair's own return
        # values -- a policy that Group Policy reapplies will already be back.
        $post   = Get-ItemProperty -Path $WUPolicyKey -ErrorAction SilentlyContinue
        $postAu = Get-ItemProperty -Path $AUPolicyKey -ErrorAction SilentlyContinue
        $w32    = Get-ServiceSafe -Name 'w32time'
        $wua    = Get-ServiceSafe -Name 'wuauserv'

        Write-Info ("WUServer    : {0}" -f $(if ($post.WUServer) { $post.WUServer } else { '(not set)' }))
        Write-Info ("UseWUServer : {0}" -f $(if ($null -ne $postAu.UseWUServer) { $postAu.UseWUServer } else { '(not set)' }))
        Write-Info ("w32time     : {0}" -f $(if ($w32) { $w32.Status } else { 'absent' }))
        Write-Info ("wuauserv    : {0}" -f $(if ($wua) { $wua.Status } else { 'absent' }))

        $remaining = @()
        if ($post.WUServer -and $postAu.UseWUServer -eq 1 -and $audit.Wsus.Configured -and -not $audit.Wsus.Reachable) {
            $remaining += 'An unreachable WSUS pointer is still set'
        }
        if ($w32 -and $w32.Status -ne 'Running') { $remaining += 'The Windows Time service is not running' }
        if ($audit.Proxy.Configured)             { $remaining += 'A WinHTTP proxy is set — review it manually' }
        if (@($audit.Endpoints | Where-Object { -not $_.Reachable }).Count -gt 0) {
            $remaining += 'Microsoft Update endpoints are unreachable — this is an upstream network block'
        }

        Write-Host ""
        if ($remaining.Count -eq 0) {
            Write-Ok "Repair complete. Check Settings > Windows Update; reboot if the error persists."
            Add-TKNote -Text 'CONDUIT repair completed with no remaining blockers.' -Category 'Resolution' -ScriptName 'conduit'
        } else {
            foreach ($r in $remaining) { Write-Warn $r }
            Add-TKNote -Text ("CONDUIT repair left {0} item(s) outstanding: {1}" -f $remaining.Count, ($remaining -join '; ')) -Category 'Issue' -ScriptName 'conduit'
        }
    }

    $verdict = Get-ConduitVerdict
    Write-Section "VERDICT"
    switch ($verdict.Class) {
        'err'   { Write-Fail $verdict.Verdict }
        'warn'  { Write-Warn $verdict.Verdict }
        default { Write-Ok   $verdict.Verdict }
    }

    Add-TKNote -Text ("CONDUIT {0} run on {1}: verdict {2} ({3} finding(s))." -f $Mode, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'conduit'

    Write-Step "Generating HTML report..."
    $html      = Build-ConduitReport -Audit $audit -Verdict $verdict
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "CONDUIT_${timestamp}.html"

    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'conduit' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-ConduitBanner
    Invoke-ConduitRun -Mode $Action
} else {
    $choice = ''

    do {
        Show-ConduitBanner

        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only diagnosis + HTML report" -ForegroundColor $C.Info
        Write-Host "  [2] Repair  -  audit, then apply the safe repairs" -ForegroundColor $C.Info
        Write-Host "  [3] Repair + reset update cache  -  also rebuilds SoftwareDistribution / catroot2" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($WhatIf) { Write-Host "  Dry run is active — options 2 and 3 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-ConduitRun -Mode 'Audit' }
            '2' { Invoke-ConduitRun -Mode 'Repair' }
            '3' { Invoke-ConduitRun -Mode 'ResetCache' }
            'Q' {
                Write-Host ""
                Write-Host "  Closing C.O.N.D.U.I.T." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-3 or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -ne 'Q') {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}

# The note buffer ends with this run, so save it now -- one file per run.
$null = Save-TKNoteReport -Prefix 'CONDUIT' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if ($Transcript) { Stop-TKTranscript }
