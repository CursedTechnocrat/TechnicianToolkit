# lodestar.ps1 - L.O.D.E.S.T.A.R. — Locates Our Domain controller, Establishes Secure Trust And Repairs
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
    L.O.D.E.S.T.A.R. — Locates Our Domain controller, Establishes Secure Trust And Repairs
    Domain Trust & Secure Channel Diagnostics and Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Diagnoses "The trust relationship between this workstation and the
    primary domain failed" and the quieter failures that precede it, on a
    domain-joined Windows machine:

      - Domain controller discovery (nltest /dsgetdc) and the DC locator SRV
        record in DNS
      - DNS servers that are public resolvers rather than domain DNS -- the
        most common cause of a member that cannot find its domain
      - Reachability of the DC on DNS, Kerberos, RPC, LDAP and SMB
      - Clock offset from the DC: Kerberos refuses tickets past five minutes
      - The secure channel itself (nltest /sc_verify), with the Netlogon
        status decoded into connectivity versus a broken machine password
      - Netlogon policy that stops the machine password from rotating

    Audit is read-only. Repair resynchronises the clock from the domain
    hierarchy, resets the secure channel to a DC (nltest /sc_reset), and --
    only when the machine password itself is rejected, and only
    interactively -- resets the computer account password with domain
    credentials the technician enters. It never changes DNS settings and never
    unjoins or rejoins the domain; those are reported with the remedy.
    -WhatIf previews every repair.

.USAGE
    PS C:\> .\lodestar.ps1                               # Interactive menu
    PS C:\> .\lodestar.ps1 -Unattended                   # Read-only audit + HTML report
    PS C:\> .\lodestar.ps1 -Unattended -Action Repair    # Audit, then the repairs that need no credentials
    PS C:\> .\lodestar.ps1 -Action Repair -WhatIf        # Preview the repairs only

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit', 'Repair')]
    [string]$Action = 'Audit',
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
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Netlogon status codes nltest reports, split by what fixes them. Connectivity
# failures are fixed on the network (DNS, firewall, the DC); Trust failures
# mean the machine account and its password no longer agree with the domain.
$NetlogonStatusCodes = @{
    0    = @{ Name = 'NERR_Success';                      Kind = 'Ok';           Meaning = 'The secure channel is healthy.' }
    5    = @{ Name = 'ERROR_ACCESS_DENIED';               Kind = 'Trust';        Meaning = 'The DC rejected the machine''s credentials, or the check ran without elevation.' }
    53   = @{ Name = 'ERROR_BAD_NETPATH';                 Kind = 'Connectivity'; Meaning = 'The DC''s network path could not be reached.' }
    1311 = @{ Name = 'ERROR_NO_LOGON_SERVERS';            Kind = 'Connectivity'; Meaning = 'No domain controller could be contacted.' }
    1355 = @{ Name = 'ERROR_NO_SUCH_DOMAIN';              Kind = 'Connectivity'; Meaning = 'The domain could not be found -- usually DNS.' }
    1722 = @{ Name = 'RPC_S_SERVER_UNAVAILABLE';          Kind = 'Connectivity'; Meaning = 'RPC to the DC failed -- firewall or the DC is down.' }
    1786 = @{ Name = 'ERROR_NO_TRUST_LSA_SECRET';         Kind = 'Trust';        Meaning = 'The local copy of the machine password is missing.' }
    1787 = @{ Name = 'ERROR_NO_TRUST_SAM_ACCOUNT';        Kind = 'Trust';        Meaning = 'The computer account is missing from the domain -- deleted, or reset by someone else.' }
    1789 = @{ Name = 'ERROR_TRUSTED_RELATIONSHIP_FAILURE'; Kind = 'Trust';       Meaning = 'The machine password does not match the domain''s copy.' }
    1790 = @{ Name = 'ERROR_TRUSTED_DOMAIN_FAILURE';      Kind = 'Trust';        Meaning = 'The trust between the domains failed.' }
}

# TCP ports a domain member needs to reach on its DC.
$DcPorts = [ordered]@{
    53  = 'DNS'
    88  = 'Kerberos'
    135 = 'RPC endpoint mapper'
    389 = 'LDAP'
    445 = 'SMB (SYSVOL / Netlogon)'
}

# Kerberos rejects tickets beyond five minutes of skew; drift past a minute is
# worth fixing before it gets there.
$ClockSkewErrorSeconds   = 300
$ClockSkewWarningSeconds = 60

$NetlogonParamsKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$LodestarFindings = @{
    'NotDomainJoined' = @{
        Severity = 'Info'
        Title    = 'Not joined to an Active Directory domain'
        Summary  = 'This machine is in a workgroup (or Entra-joined only), so there is no domain trust to check.'
        Remedy   = 'Nothing to do. For Entra-joined machines, check dsregcmd /status instead.'
    }
    'DomainNotFound' = @{
        Severity = 'Error'
        Title    = 'No domain controller could be located'
        Summary  = 'DC discovery (nltest /dsgetdc) failed, so nothing domain-related can work -- logons fall back to cached credentials.'
        Remedy   = 'Fix DNS first: the machine must use the domain''s DNS servers. Then check the network path to a DC.'
    }
    'DcSrvMissing' = @{
        Severity = 'Error'
        Title    = 'DC locator SRV record does not resolve'
        Summary  = '_ldap._tcp.dc._msdcs.<domain> did not resolve through this machine''s DNS servers, which is how Windows finds a domain controller.'
        Remedy   = 'Point the adapter at the domain DNS servers (usually the DCs), not the router or a public resolver.'
    }
    'PublicDnsServer' = @{
        Severity = 'Warning'
        Title    = 'Public DNS resolver configured on a domain member'
        Summary  = 'A public resolver cannot answer for the internal domain. Windows uses it whenever the first server is slow, so domain lookups fail intermittently.'
        Remedy   = 'Use only the domain DNS servers on domain members; let those forward to public resolvers. L.O.D.E.S.T.A.R. never changes DNS itself.'
    }
    'DcPortBlocked' = @{
        Severity = 'Error'
        Title    = 'Domain controller port unreachable'
        Summary  = 'A port the domain member needs on its DC did not answer -- a firewall, VPN split-tunnel or the DC itself.'
        Remedy   = 'Open the path to the DC for the ports listed. No local repair fixes an upstream block.'
    }
    'ClockSkew' = @{
        Severity = 'Error'
        Title    = 'Clock is more than five minutes off the DC'
        Summary  = 'Kerberos rejects tickets beyond five minutes of skew, so domain logons and access fail even with a healthy secure channel.'
        Remedy   = 'Repair resynchronises from the domain hierarchy (w32tm /resync).'
    }
    'ClockDrift' = @{
        Severity = 'Warning'
        Title    = 'Clock is drifting from the DC'
        Summary  = 'The offset is over a minute: not failing yet, but heading toward the five-minute Kerberos limit.'
        Remedy   = 'Repair resynchronises from the domain hierarchy. Check the Windows Time service if it keeps drifting.'
    }
    'SecureChannelBroken' = @{
        Severity = 'Error'
        Title    = 'Secure channel down -- the DC could not be reached'
        Summary  = 'Netlogon could not reach a DC to verify the channel. This is a connectivity fault, not a broken trust.'
        Remedy   = 'Fix the DNS and connectivity findings first. Repair then resets the secure channel to a reachable DC.'
    }
    'TrustBroken' = @{
        Severity = 'Error'
        Title    = 'Trust relationship failed'
        Summary  = 'The DC rejected the machine''s credentials: the machine password and the domain''s copy disagree, or the computer account is gone.'
        Remedy   = 'Repair (interactive) resets the machine password with domain credentials. If the computer account was deleted, the machine must be rejoined.'
    }
    'ComputerAccountMissing' = @{
        Severity = 'Error'
        Title    = 'Computer account missing from the domain'
        Summary  = 'The DC reports no account for this machine (ERROR_NO_TRUST_SAM_ACCOUNT). A password reset cannot fix this.'
        Remedy   = 'Rejoin the domain: sign in with a local admin, Remove-Computer to a workgroup, restart, then Add-Computer. L.O.D.E.S.T.A.R. never rejoins automatically.'
    }
    'SecureChannelUnknown' = @{
        Severity = 'Warning'
        Title    = 'Secure channel state could not be determined'
        Summary  = 'nltest returned output L.O.D.E.S.T.A.R. could not read, or was not available.'
        Remedy   = 'Run elevated, and check nltest /sc_verify:<domain> by hand.'
    }
    'PasswordChangeDisabled' = @{
        Severity = 'Warning'
        Title    = 'Machine password rotation disabled'
        Summary  = 'Netlogon DisablePasswordChange = 1. Usually set to "fix" snapshot restores, it leaves a long-lived machine credential instead.'
        Remedy   = 'Clear the setting at its source (GPO: Domain member: Disable machine account password changes) unless a documented reason exists.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[object]]::new()

function Add-LodestarFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $LodestarFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
    })
}

function Add-LodestarAction {
    param([Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Status, [string]$Detail = '')
    [void]$Actions.Add([PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Step = $Step; Status = $Status; Detail = $Detail })
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

function Show-LodestarBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗      ██████╗ ██████╗ ███████╗███████╗████████╗ █████╗ ██████╗
  ██║     ██╔═══██╗██╔══██╗██╔════╝██╔════╝╚══██╔══╝██╔══██╗██╔══██╗
  ██║     ██║   ██║██║  ██║█████╗  ███████╗   ██║   ███████║██████╔╝
  ██║     ██║   ██║██║  ██║██╔══╝  ╚════██║   ██║   ██╔══██║██╔══██╗
  ███████╗╚██████╔╝██████╔╝███████╗███████║   ██║   ██║  ██║██║  ██║
  ╚══════╝ ╚═════╝ ╚═════╝ ╚══════╝╚══════╝   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    L.O.D.E.S.T.A.R. — Locates Our Domain controller, Establishes Secure Trust And Repairs" -ForegroundColor Cyan
    Write-Host "    Domain Trust & Secure Channel Diagnostics and Repair Tool" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host ""
        Write-Host "    *** DRY RUN — no changes will be applied ***" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS — parsers for tool output. No I/O; the Pester suite calls
# these with captured text.
# ─────────────────────────────────────────────────────────────────────────────

function ConvertFrom-NltestDsGetDc {
    param([string[]]$Lines)
    $text = $Lines -join "`n"
    $status = [regex]::Match($text, '(?i)Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')
    $dc     = [regex]::Match($text, '(?im)^\s*DC:\s*\\\\(\S+)')
    $addr   = [regex]::Match($text, '(?im)^\s*Address:\s*\\\\(\S+)')
    $site   = [regex]::Match($text, '(?im)^\s*Dc Site Name:\s*(.+?)\s*$')
    $ours   = [regex]::Match($text, '(?im)^\s*Our Site Name:\s*(.+?)\s*$')
    return [PSCustomObject]@{
        Success    = $dc.Success
        Dc         = if ($dc.Success) { $dc.Groups[1].Value } else { '' }
        Address    = if ($addr.Success) { $addr.Groups[1].Value } else { '' }
        DcSite     = if ($site.Success) { $site.Groups[1].Value } else { '' }
        OurSite    = if ($ours.Success) { $ours.Groups[1].Value } else { '' }
        StatusCode = if ($status.Success) { [int]$status.Groups[1].Value } elseif ($dc.Success) { 0 } else { $null }
        StatusName = if ($status.Success) { $status.Groups[2].Value } else { '' }
    }
}

function ConvertFrom-NltestSecureChannel {
    # Reads nltest /sc_query or /sc_verify. The overall code is the verify
    # status when present (it is the stronger check), else the connection
    # status, else a bare "failed: Status = n" line.
    param([string[]]$Lines)
    $text = $Lines -join "`n"
    $dc      = [regex]::Match($text, '(?i)Trusted DC Name\s+\\\\(\S+)')
    $conn    = [regex]::Match($text, '(?i)Trusted DC Connection Status\s+Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')
    $verify  = [regex]::Match($text, '(?i)Trust Verification Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')
    $failed  = [regex]::Match($text, '(?i)failed:\s*Status\s*=\s*(\d+)\s+0x[0-9a-f]+\s+(\S+)')

    $code = $null; $name = ''
    if ($verify.Success -and [int]$verify.Groups[1].Value -ne 0) { $code = [int]$verify.Groups[1].Value; $name = $verify.Groups[2].Value }
    elseif ($conn.Success -and [int]$conn.Groups[1].Value -ne 0) { $code = [int]$conn.Groups[1].Value;   $name = $conn.Groups[2].Value }
    elseif ($failed.Success)                                      { $code = [int]$failed.Groups[1].Value; $name = $failed.Groups[2].Value }
    elseif ($verify.Success -or $conn.Success)                    { $code = 0; $name = 'NERR_Success' }

    return [PSCustomObject]@{
        TrustedDc  = if ($dc.Success) { $dc.Groups[1].Value } else { '' }
        Verified   = $verify.Success
        StatusCode = $code
        StatusName = $name
    }
}

function Get-NetlogonStatusInfo {
    param($Code)
    if ($null -eq $Code) { return [PSCustomObject]@{ Name = 'Unknown'; Kind = 'Unknown'; Meaning = 'No status could be read.' } }
    $entry = $NetlogonStatusCodes[[int]$Code]
    if ($entry) { return [PSCustomObject]@{ Name = $entry.Name; Kind = $entry.Kind; Meaning = $entry.Meaning } }
    return [PSCustomObject]@{ Name = "Status $Code"; Kind = 'Unknown'; Meaning = 'Not a status L.O.D.E.S.T.A.R. recognises -- look it up with: net helpmsg ' + $Code }
}

function ConvertFrom-W32tmStripchart {
    # w32tm /stripchart /dataonly prints "hh:mm:ss, +00.0123456s" per sample.
    # Returns the offset in seconds of the first sample, or $null.
    param([string[]]$Lines)
    foreach ($l in $Lines) {
        $m = [regex]::Match("$l", '^\s*\d{1,2}:\d{2}:\d{2},\s*([+-]?\d+(?:\.\d+)?)s\s*$')
        if ($m.Success) { return [double]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) }
    }
    return $null
}

function Test-PublicIpAddress {
    # True for an address that is not private, loopback, link-local, CGNAT or
    # IPv6 unique-local -- i.e. a resolver on the internet.
    param([string]$Address)
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse("$Address", [ref]$ip)) { return $false }
    if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
        if ($ip.IsIPv6LinkLocal -or $ip.IsIPv6SiteLocal -or [System.Net.IPAddress]::IsLoopback($ip)) { return $false }
        $first = $ip.GetAddressBytes()[0]
        return -not (($first -band 0xFE) -eq 0xFC)
    }
    $b = $ip.GetAddressBytes()
    if ($b[0] -eq 10 -or $b[0] -eq 127 -or $b[0] -eq 0) { return $false }
    if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $false }
    if ($b[0] -eq 192 -and $b[1] -eq 168) { return $false }
    if ($b[0] -eq 169 -and $b[1] -eq 254) { return $false }
    if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $false }
    return $true
}

function Get-LodestarVerdict {
    param([object[]]$FindingList)
    $codes = @($FindingList | ForEach-Object { $_.Code })
    $sev   = @($FindingList | ForEach-Object { $_.Severity })
    if ($codes -contains 'NotDomainJoined') { return [PSCustomObject]@{ Verdict = 'Not joined'; Class = 'info' } }
    if ($sev -contains 'Error')             { return [PSCustomObject]@{ Verdict = 'Broken';     Class = 'err'  } }
    if ($sev -contains 'Warning')           { return [PSCustomObject]@{ Verdict = 'Degraded';   Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Native {
    # Captures a native tool's output without letting a missing binary or a
    # non-zero exit abort the audit.
    param([string]$File, [string[]]$Arguments)
    try {
        $out = & $File @Arguments 2>&1 | ForEach-Object { "$_" }
        return @($out)
    } catch {
        return @("failed: $($_.Exception.Message)")
    }
}

function Test-TcpPort {
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 2000)
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

function Get-LodestarContext {
    $cs = $null; $os = $null
    try { $cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop } catch { $cs = $null }
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }
    return [PSCustomObject]@{
        Computer     = $env:COMPUTERNAME
        PartOfDomain = if ($cs) { [bool]$cs.PartOfDomain } else { $false }
        Domain       = if ($cs) { "$($cs.Domain)" } else { '' }
        OS           = if ($os) { "$($os.Caption) (build $($os.BuildNumber))" } else { 'Unknown' }
        Edition      = "$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)"
    }
}

function Get-LodestarDnsServers {
    $rows = @()
    try {
        $up = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' } | ForEach-Object { $_.ifIndex })
        $rows = @(Get-DnsClientServerAddress -ErrorAction Stop | Where-Object { $up -contains $_.InterfaceIndex -and $_.ServerAddresses } | ForEach-Object {
            $alias = $_.InterfaceAlias
            foreach ($s in $_.ServerAddresses) { [PSCustomObject]@{ Interface = $alias; Server = "$s"; Public = (Test-PublicIpAddress -Address "$s") } }
        })
    } catch {
        $rows = @()
    }
    return $rows
}

function Invoke-LodestarAudit {
    Write-Section "DOMAIN MEMBERSHIP"
    $ctx = Get-LodestarContext
    Write-Info "$($ctx.Computer)  |  $($ctx.OS)"
    if (-not $ctx.PartOfDomain) {
        Write-Info "Not joined to an Active Directory domain."
        Add-LodestarFinding -Code 'NotDomainJoined' -Detail "Workgroup / domain field: $($ctx.Domain)"
        return [PSCustomObject]@{
            Context = $ctx; Joined = $false; Domain = ''; Dns = @(); SrvOk = $false; Ports = @(); Offset = $null; MaxPwdAge = $null
            DsGetDc = [PSCustomObject]@{ Success = $false; Dc = ''; Address = ''; DcSite = ''; OurSite = ''; StatusCode = $null; StatusName = '' }
            Channel = [PSCustomObject]@{ TrustedDc = ''; Verified = $false; StatusCode = $null; StatusName = '' }
            Status  = [PSCustomObject]@{ Name = 'n/a'; Kind = 'NotJoined'; Meaning = 'Not joined to a domain.' }
        }
    }
    $domain = $ctx.Domain
    Write-Info "Domain: $domain"

    Write-Section "DNS"
    $dns = Get-LodestarDnsServers
    foreach ($d in $dns) { Write-Info ("{0,-24} {1}{2}" -f $d.Interface, $d.Server, $(if ($d.Public) { '  (public resolver)' } else { '' })) }
    $public = @($dns | Where-Object { $_.Public })
    if ($public.Count -gt 0) { Add-LodestarFinding -Code 'PublicDnsServer' -Detail (($public | ForEach-Object { "$($_.Server) on $($_.Interface)" }) -join ', ') }

    $srvName = "_ldap._tcp.dc._msdcs.$domain"
    $srvOk = $false
    try { $srvOk = [bool](@(Resolve-DnsName -Name $srvName -Type SRV -DnsOnly -ErrorAction Stop | Where-Object { $_.Type -eq 'SRV' }).Count) } catch { $srvOk = $false }
    if (-not $srvOk) { Add-LodestarFinding -Code 'DcSrvMissing' -Detail $srvName }
    Write-Info ("DC locator SRV: {0}" -f $(if ($srvOk) { 'resolves' } else { 'DOES NOT RESOLVE' }))

    Write-Section "DOMAIN CONTROLLER"
    $dsget = ConvertFrom-NltestDsGetDc -Lines (Invoke-Native -File 'nltest.exe' -Arguments @("/dsgetdc:$domain"))
    if (-not $dsget.Success) {
        Add-LodestarFinding -Code 'DomainNotFound' -Detail ("nltest /dsgetdc: {0} {1}" -f $dsget.StatusCode, $dsget.StatusName)
        Write-Fail "No DC located ($($dsget.StatusName))."
    } else {
        Write-Info "DC: $($dsget.Dc)  ($($dsget.Address))  site $($dsget.DcSite)"
    }

    $ports = @()
    $offset = $null
    if ($dsget.Success) {
        $ports = foreach ($p in $DcPorts.Keys) {
            [PSCustomObject]@{ Port = [int]$p; Service = $DcPorts[$p]; Open = (Test-TcpPort -Hostname $dsget.Dc -Port ([int]$p)) }
        }
        $ports = @($ports)
        $closed = @($ports | Where-Object { -not $_.Open })
        if ($closed.Count -gt 0) { Add-LodestarFinding -Code 'DcPortBlocked' -Detail (($closed | ForEach-Object { "$($_.Port) $($_.Service)" }) -join ', ') }
        foreach ($p in $ports) { Write-Info ("{0,-5} {1,-26} {2}" -f $p.Port, $p.Service, $(if ($p.Open) { 'open' } else { 'BLOCKED' })) }

        Write-Section "TIME"
        $offset = ConvertFrom-W32tmStripchart -Lines (Invoke-Native -File 'w32tm.exe' -Arguments @('/stripchart', "/computer:$($dsget.Dc)", '/samples:1', '/dataonly'))
        if ($null -eq $offset) {
            Write-Warn "Could not measure the offset from $($dsget.Dc)."
        } else {
            Write-Info ("Offset from DC: {0:N2}s" -f $offset)
            if ([math]::Abs($offset) -gt $ClockSkewErrorSeconds)       { Add-LodestarFinding -Code 'ClockSkew'  -Detail ("{0:N0} seconds" -f $offset) }
            elseif ([math]::Abs($offset) -gt $ClockSkewWarningSeconds) { Add-LodestarFinding -Code 'ClockDrift' -Detail ("{0:N0} seconds" -f $offset) }
        }
    }

    Write-Section "SECURE CHANNEL"
    $channel = Get-LodestarSecureChannel -Domain $domain
    $status  = Get-NetlogonStatusInfo -Code $channel.StatusCode
    Write-Info ("Status: {0} ({1}) via {2}" -f $status.Name, $status.Meaning, $(if ($channel.TrustedDc) { $channel.TrustedDc } else { 'no DC' }))
    Add-LodestarChannelFinding -Channel $channel -Status $status

    $np = Get-ItemProperty -Path $NetlogonParamsKey -ErrorAction SilentlyContinue
    if ($np -and $np.DisablePasswordChange -eq 1) { Add-LodestarFinding -Code 'PasswordChangeDisabled' -Detail "$NetlogonParamsKey\DisablePasswordChange = 1" }

    return [PSCustomObject]@{
        Context   = $ctx
        Joined    = $true
        Domain    = $domain
        Dns       = @($dns)
        SrvOk     = $srvOk
        DsGetDc   = $dsget
        Ports     = @($ports)
        Offset    = $offset
        Channel   = $channel
        Status    = $status
        MaxPwdAge = if ($np -and $np.MaximumPasswordAge) { [int]$np.MaximumPasswordAge } else { 30 }
    }
}

function Get-LodestarSecureChannel {
    # /sc_verify checks the machine password with the DC; it needs elevation.
    # /sc_query only reports the last known state, so it is the fallback.
    param([string]$Domain)
    $channel = ConvertFrom-NltestSecureChannel -Lines (Invoke-Native -File 'nltest.exe' -Arguments @("/sc_verify:$Domain"))
    if ($null -eq $channel.StatusCode) {
        $channel = ConvertFrom-NltestSecureChannel -Lines (Invoke-Native -File 'nltest.exe' -Arguments @("/sc_query:$Domain"))
    }
    return $channel
}

function Add-LodestarChannelFinding {
    param([object]$Channel, [object]$Status)
    switch ($Status.Kind) {
        'Ok'           { }
        'Connectivity' { Add-LodestarFinding -Code 'SecureChannelBroken' -Detail "$($Status.Name): $($Status.Meaning)" }
        'Trust' {
            if ($Channel.StatusCode -eq 1787) { Add-LodestarFinding -Code 'ComputerAccountMissing' -Detail "$($Status.Name): $($Status.Meaning)" }
            else                              { Add-LodestarFinding -Code 'TrustBroken' -Detail "$($Status.Name): $($Status.Meaning)" }
        }
        default        { Add-LodestarFinding -Code 'SecureChannelUnknown' -Detail "$($Status.Name): $($Status.Meaning)" }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Repair-LodestarTime {
    Write-Section "REPAIR — TIME"
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would set w32time to sync from the domain hierarchy and force a resync."
        Add-LodestarAction -Step 'Resync time from the domain' -Status 'WhatIf'
        return
    }
    $cfg = Invoke-Native -File 'w32tm.exe' -Arguments @('/config', '/syncfromflags:domhier', '/update')
    $rs  = Invoke-Native -File 'w32tm.exe' -Arguments @('/resync', '/force')
    $ok  = ($rs -join ' ') -match '(?i)completed successfully'
    if ($ok) { Write-Ok "Clock resynchronised from the domain hierarchy." } else { Write-Warn "w32tm /resync: $($rs -join ' ')" }
    Add-LodestarAction -Step 'Resync time from the domain' -Status $(if ($ok) { 'Done' } else { 'Failed' }) -Detail (@($cfg + $rs) -join ' ')
    if ($ok) { Add-TKNote -Text 'Resynchronised the clock from the domain hierarchy.' -Category 'Action' -ScriptName 'lodestar' }
}

function Repair-LodestarSecureChannel {
    param([string]$Domain)
    Write-Section "REPAIR — SECURE CHANNEL"
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would reset the secure channel: nltest /sc_reset:$Domain"
        Add-LodestarAction -Step 'Reset secure channel' -Status 'WhatIf' -Detail "nltest /sc_reset:$Domain"
        return
    }
    $out    = Invoke-Native -File 'nltest.exe' -Arguments @("/sc_reset:$Domain")
    $parsed = ConvertFrom-NltestSecureChannel -Lines $out
    $ok     = $parsed.StatusCode -eq 0
    if ($ok) { Write-Ok "Secure channel reset to $($parsed.TrustedDc)." } else { Write-Warn "nltest /sc_reset: $($parsed.StatusName)" }
    Add-LodestarAction -Step 'Reset secure channel' -Status $(if ($ok) { 'Done' } else { 'Failed' }) -Detail ($out -join ' ')
    if ($ok) { Add-TKNote -Text "Reset the secure channel to $($parsed.TrustedDc)." -Category 'Action' -ScriptName 'lodestar' }
}

function Repair-LodestarMachinePassword {
    # Only for a rejected machine password. Needs domain credentials, so it is
    # interactive only. Reset-ComputerMachinePassword exists in Windows
    # PowerShell but not PowerShell 7, so under 7 it runs in powershell.exe.
    param([string]$Dc)
    Write-Section "REPAIR — MACHINE PASSWORD"
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would reset the computer account password against $Dc with domain credentials."
        Add-LodestarAction -Step 'Reset machine password' -Status 'WhatIf' -Detail $Dc
        return
    }
    if ($Unattended) {
        Write-Warn "Resetting the machine password needs domain credentials -- skipped in unattended mode."
        Add-LodestarAction -Step 'Reset machine password' -Status 'Skipped' -Detail 'Needs domain credentials; run interactively.'
        return
    }
    $ans = Read-Host "  Reset the computer account password against $Dc now? Needs a domain account with rights to reset it. [Y/N]"
    if ($ans -notmatch '^[Yy]') {
        Add-LodestarAction -Step 'Reset machine password' -Status 'Skipped' -Detail 'Declined by technician.'
        return
    }
    try {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            Reset-ComputerMachinePassword -Server $Dc -Credential (Get-Credential -Message 'Domain account to reset this computer''s password') -ErrorAction Stop
        } else {
            $command = "Reset-ComputerMachinePassword -Server '$Dc' -Credential (Get-Credential -Message 'Domain account to reset this computer''s password') -ErrorAction Stop"
            & powershell.exe -NoProfile -Command $command
            if ($LASTEXITCODE -ne 0) { throw "powershell.exe exited $LASTEXITCODE" }
        }
        Write-Ok "Machine password reset against $Dc."
        Add-LodestarAction -Step 'Reset machine password' -Status 'Done' -Detail $Dc
        Add-TKNote -Text "Reset the computer account password against $Dc." -Category 'Action' -ScriptName 'lodestar'
    } catch {
        Write-Fail "Machine password reset failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'lodestar' -Message "Machine password reset failed: $($_.Exception.Message)" -Category 'Domain Trust'
        Add-LodestarAction -Step 'Reset machine password' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Invoke-LodestarRepair {
    param([object]$Audit)
    $codes = @($Findings | ForEach-Object { $_.Code })

    if ($codes -contains 'ClockSkew' -or $codes -contains 'ClockDrift') { Repair-LodestarTime }

    if ($codes -contains 'SecureChannelBroken' -or $codes -contains 'TrustBroken' -or $codes -contains 'SecureChannelUnknown') {
        Repair-LodestarSecureChannel -Domain $Audit.Domain
        if ($codes -contains 'TrustBroken' -and -not $WhatIf) {
            # Re-verify: a stale DC choice can masquerade as a trust failure,
            # and sc_reset alone fixes that without touching the password.
            $again = Get-LodestarSecureChannel -Domain $Audit.Domain
            if ((Get-NetlogonStatusInfo -Code $again.StatusCode).Kind -eq 'Trust') {
                $dc = if ($Audit.DsGetDc.Dc) { $Audit.DsGetDc.Dc } else { $again.TrustedDc }
                Repair-LodestarMachinePassword -Dc $dc
            }
        } elseif ($codes -contains 'TrustBroken') {
            Repair-LodestarMachinePassword -Dc $Audit.DsGetDc.Dc
        }
    }

    if ($codes -contains 'ComputerAccountMissing') {
        Write-Warn "The computer account is missing from the domain. Rejoin the machine -- L.O.D.E.S.T.A.R. does not unjoin or rejoin."
        Add-LodestarAction -Step 'Rejoin domain' -Status 'Skipped' -Detail 'Computer account missing; rejoin manually.'
    }
    if ($codes -contains 'PublicDnsServer' -or $codes -contains 'DcSrvMissing') {
        Add-LodestarAction -Step 'Correct DNS servers' -Status 'Skipped' -Detail 'L.O.D.E.S.T.A.R. never changes DNS -- point the adapter at the domain DNS servers.'
    }
    if ($Actions.Count -eq 0) {
        Write-Ok "Nothing L.O.D.E.S.T.A.R. can repair automatically."
        Add-LodestarAction -Step 'Repair' -Status 'Skipped' -Detail 'No repairable finding.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-LodestarReport {
    param([object]$Audit, [object]$Verdict, [object]$After)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>No issues found.</td></tr>") }
    foreach ($f in $Findings) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
            $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $dRows = [System.Text.StringBuilder]::new()
    foreach ($d in $Audit.Dns) {
        $badge = if ($d.Public) { "<span class='tk-badge-warn'>Public</span>" } else { "<span class='tk-badge-ok'>Private</span>" }
        [void]$dRows.Append("<tr><td>$(EscHtml $d.Interface)</td><td class='tk-mono'>$(EscHtml $d.Server)</td><td>$badge</td></tr>")
    }
    if ($Audit.Dns.Count -eq 0) { [void]$dRows.Append("<tr><td colspan='3'>No DNS servers read.</td></tr>") }

    $pRows = [System.Text.StringBuilder]::new()
    foreach ($p in $Audit.Ports) {
        $badge = if ($p.Open) { "<span class='tk-badge-ok'>Open</span>" } else { "<span class='tk-badge-err'>Blocked</span>" }
        [void]$pRows.Append("<tr><td>$($p.Port)</td><td>$(EscHtml $p.Service)</td><td>$badge</td></tr>")
    }
    if ($Audit.Ports.Count -eq 0) { [void]$pRows.Append("<tr><td colspan='3'>No DC to test.</td></tr>") }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Actions.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>") }
    foreach ($a in $Actions) {
        $badge = switch ($a.Status) { 'Done' { 'ok' } 'WhatIf' { 'blue' } 'Failed' { 'err' } default { 'info' } }
        [void]$aRows.Append("<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td><td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>")
    }

    $channelBadge = switch ($Audit.Status.Kind) { 'Ok' { 'ok' } 'Connectivity' { 'err' } 'Trust' { 'err' } 'NotJoined' { 'info' } default { 'warn' } }
    $channelLabel = switch ($Audit.Status.Kind) { 'Ok' { 'Healthy' } 'Connectivity' { 'Unreachable' } 'Trust' { 'Trust failed' } 'NotJoined' { 'n/a' } default { 'Unknown' } }
    $dcBadge      = if (-not $Audit.Joined) { 'info' } elseif ($Audit.DsGetDc.Success) { 'ok' } else { 'err' }
    $afterText = if ($After) { "$($After.Name) -- $($After.Meaning)" } else { '' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'L.O.D.E.S.T.A.R. Domain Trust Report' `
        -ScriptName 'L.O.D.E.S.T.A.R.' `
        -Subtitle   "${orgPrefix}Domain Trust & Secure Channel -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Domain'    = $(if ($Audit.Joined) { $Audit.Domain } else { 'Not joined' })
            'Generated' = $reportDate
            'Mode'      = $(if ($WhatIf) { "$Action (dry run)" } else { $Action })
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Secure Channel', 'DNS', 'Domain Controller', 'Actions Taken')

    $offsetText = if ($null -ne $Audit.Offset) { '{0:N2} s' -f $Audit.Offset } else { 'not measured' }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Domain Trust</div></div>
    <div class="tk-summary-card $channelBadge"><div class="tk-summary-num">$channelLabel</div><div class="tk-summary-lbl">Secure Channel</div></div>
    <div class="tk-summary-card $dcBadge"><div class="tk-summary-num">$(if (-not $Audit.Joined) { 'n/a' } elseif ($Audit.DsGetDc.Success) { 'Found' } else { 'None' })</div><div class="tk-summary-lbl">Domain Controller</div></div>
    <div class="tk-summary-card $(if (@($Audit.Dns | Where-Object { $_.Public }).Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$(@($Audit.Dns | Where-Object { $_.Public }).Count)</div><div class="tk-summary-lbl">Public DNS Servers</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $offsetText)</div><div class="tk-summary-lbl">Clock Offset</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Secure Channel</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">Status</span> $(EscHtml $Audit.Status.Name) -- $(EscHtml $Audit.Status.Meaning)<br/>
      <span class="tk-info-label">Checked with</span> $(if ($Audit.Channel.Verified) { 'nltest /sc_verify (machine password checked with the DC)' } else { 'nltest /sc_query (last known state)' })<br/>
      <span class="tk-info-label">Trusted DC</span> $(EscHtml $(if ($Audit.Channel.TrustedDc) { $Audit.Channel.TrustedDc } else { '-' }))<br/>
      <span class="tk-info-label">Machine password max age</span> $(if ($Audit.MaxPwdAge) { "$($Audit.MaxPwdAge) day(s)" } else { '-' })$(if ($afterText) { "<br/><span class='tk-info-label'>After repair</span> $(EscHtml $afterText)" })
    </div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> DNS</div>
    <div class="tk-card">
      <div class="tk-info-box"><span class="tk-info-label">DC locator SRV</span> _ldap._tcp.dc._msdcs.$(EscHtml $Audit.Domain) -- $(if ($Audit.SrvOk) { 'resolves' } else { 'does not resolve' })</div>
      <table class="tk-table"><thead><tr><th>Interface</th><th>DNS server</th><th>Kind</th></tr></thead>
      <tbody>$($dRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Domain Controller</div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">DC</span> $(EscHtml $(if ($Audit.DsGetDc.Dc) { $Audit.DsGetDc.Dc } else { '-' })) $(EscHtml $Audit.DsGetDc.Address)<br/>
        <span class="tk-info-label">DC site / our site</span> $(EscHtml $Audit.DsGetDc.DcSite) / $(EscHtml $Audit.DsGetDc.OurSite)<br/>
        <span class="tk-info-label">Clock offset</span> $(EscHtml $offsetText)
      </div>
      <table class="tk-table"><thead><tr><th>Port</th><th>Service</th><th>Result</th></tr></thead>
      <tbody>$($pRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Actions Taken</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'L.O.D.E.S.T.A.R. v6.0')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-LodestarRun {
    param([string]$Mode)
    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-LodestarAudit
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Secure channel healthy; DC, DNS and time all in order." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    $after = $null
    if ($Mode -eq 'Repair' -and $audit.Joined) {
        Invoke-LodestarRepair -Audit $audit
        if (-not $WhatIf) {
            Write-Section "VERIFY"
            $post  = Get-LodestarSecureChannel -Domain $audit.Domain
            $after = Get-NetlogonStatusInfo -Code $post.StatusCode
            if ($after.Kind -eq 'Ok') { Write-Ok "Secure channel healthy after repair. Sign out and back in with a domain account to confirm." }
            else                      { Write-Warn "Secure channel still reports $($after.Name): $($after.Meaning)" }
            Add-TKNote -Text "LODESTAR repair on $($env:COMPUTERNAME): secure channel now $($after.Name)." -Category $(if ($after.Kind -eq 'Ok') { 'Resolution' } else { 'Issue' }) -ScriptName 'lodestar'
        }
    }

    $verdict = Get-LodestarVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    Add-TKNote -Text ("LODESTAR {0} on {1}: verdict {2} ({3} finding(s))." -f $Mode, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'lodestar'

    Write-Step "Generating HTML report..."
    $html    = Build-LodestarReport -Audit $audit -Verdict $verdict -After $after
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("LODESTAR_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'lodestar' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-LodestarBanner
    Invoke-LodestarRun -Mode $Action
} else {
    $choice = ''
    do {
        Show-LodestarBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only diagnosis + HTML report" -ForegroundColor $C.Info
        Write-Host "  [2] Repair  -  audit, resync time, reset the secure channel, reset the machine password if rejected" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($WhatIf) { Write-Host "  Dry run is active — option 2 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-LodestarRun -Mode 'Audit' }
            '2' { Invoke-LodestarRun -Mode 'Repair' }
            'Q' { Write-Host ""; Write-Host "  Closing L.O.D.E.S.T.A.R." -ForegroundColor $C.Header; Write-Host "" }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1, 2 or Q." -ForegroundColor $C.Warning
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
$null = Save-TKNoteReport -Prefix 'LODESTAR' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if ($Transcript) { Stop-TKTranscript }
