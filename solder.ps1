# solder.ps1 - S.O.L.D.E.R. — Servicing-stack Overhaul: Locates Damage, Enacts Repair
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
    S.O.L.D.E.R. — Servicing-stack Overhaul: Locates Damage, Enacts Repair
    Windows Servicing (Component Store & System File) Diagnosis and Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Diagnoses and repairs the Windows servicing stack -- the component store
    (WinSxS) that updates, features and system-file repair all depend on.
    Reach for it when updates fail with 0x800f081f / 0x80073712 / 0x800f0831,
    when Windows features will not install, or when system files are damaged.

    Audit (read-only):
      - Pending restarts, split into those that block servicing (CBS
        RebootPending, PackagesPending, pending.xml, Windows Update
        RebootRequired) and the rest (file renames, a pending computer rename)
      - Component store health (DISM /CheckHealth, or /ScanHealth with -Deep)
        and size / cleanup state (DISM /AnalyzeComponentStore)
      - The last System File Checker result, read from CBS.log: files repaired
        and files it could not repair
      - Windows Update install failures from the last 30 days, with each error
        code decoded into component-store, update-client, restart, space or
        access causes
      - The Windows Modules Installer service, free space on the system drive,
        and where DISM will look for repair files (WSUS machines often cannot
        get them)

    Repair runs the standard sequence: DISM /RestoreHealth (optionally from a
    -Source image), then sfc /scannow, then re-checks the store. -Cleanup adds
    DISM /StartComponentCleanup. It never runs /ResetBase (irreversible: it
    removes the ability to uninstall updates), never edits Windows Update
    policy, and never resets the update cache -- C.O.N.D.U.I.T. owns the
    update client's plumbing. -WhatIf previews every repair.

.USAGE
    PS C:\> .\solder.ps1                                    # Interactive menu
    PS C:\> .\solder.ps1 -Unattended                        # Read-only audit + HTML report
    PS C:\> .\solder.ps1 -Unattended -Deep                  # Audit with a full ScanHealth
    PS C:\> .\solder.ps1 -Unattended -Action Repair         # RestoreHealth + SFC + re-check
    PS C:\> .\solder.ps1 -Action Repair -Source 'WIM:D:\sources\install.wim:6'   # Repair from mounted media
    PS C:\> .\solder.ps1 -Action Repair -WhatIf             # Preview the repairs only

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit', 'Repair')]
    [string]$Action = 'Audit',
    [string]$Source = '',
    [switch]$Deep,
    [switch]$Cleanup,
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
    param([Parameter(Mandatory)][string]$ScriptFile)
    if (-not (Test-IsAdmin)) {
        Write-Host "  INFO: Restarting with administrator privileges..." -ForegroundColor Yellow
        $PSExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        Start-Process -FilePath $PSExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptFile`"" -Verb RunAs
        exit
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
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

$CbsLogPath          = Join-Path $env:SystemRoot 'Logs\CBS\CBS.log'
$DismLogPath         = Join-Path $env:SystemRoot 'Logs\DISM\dism.log'
$CbsTailBytes        = 20MB       # CBS.log grows large; the last SFC run is near the end
$UpdateLookbackDays  = 30
$FreeSpaceWarningGB  = 10         # servicing stages payloads on the system drive
$FreeSpaceErrorGB    = 3

$CbsKey     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing'
$ServicingPolicyKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing'
$WuPolicyKey        = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Error codes that servicing and Windows Update report, split by what fixes
# them. Store = component-store damage (DISM / SFC); Client = the update
# client's plumbing (C.O.N.D.U.I.T.); Reboot, Space and Access are what they say.
$ServicingErrorCodes = @{
    '0x800f081f' = @{ Name = 'CBS_E_SOURCE_MISSING';                 Kind = 'Store';  Meaning = 'Repair or feature files could not be found. Give DISM a -Source, or let it reach Windows Update.' }
    '0x80073712' = @{ Name = 'ERROR_SXS_COMPONENT_STORE_CORRUPT';    Kind = 'Store';  Meaning = 'The component store is corrupt.' }
    '0x80073701' = @{ Name = 'ERROR_SXS_ASSEMBLY_MISSING';           Kind = 'Store';  Meaning = 'A component the update needs is missing from the store.' }
    '0x800f0831' = @{ Name = 'CBS_E_STORE_CORRUPTION';               Kind = 'Store';  Meaning = 'The store is missing a package the update builds on.' }
    '0x800f0982' = @{ Name = 'PSFX_E_MATCHING_COMPONENT_NOT_FOUND';  Kind = 'Store';  Meaning = 'A component the update patches is missing -- often a language pack or optional feature.' }
    '0x800f0900' = @{ Name = 'CBS_E_XML_PARSER_FAILURE';             Kind = 'Store';  Meaning = 'A servicing manifest could not be parsed.' }
    '0x80070570' = @{ Name = 'ERROR_FILE_CORRUPT';                   Kind = 'Store';  Meaning = 'A file or directory is corrupt and unreadable.' }
    '0x8007000d' = @{ Name = 'ERROR_INVALID_DATA';                   Kind = 'Store';  Meaning = 'Servicing data is invalid.' }
    '0x800f082f' = @{ Name = 'CBS_E_PENDING';                        Kind = 'Reboot'; Meaning = 'Servicing operations are waiting for a restart.' }
    '0x80070bc9' = @{ Name = 'ERROR_FAIL_REBOOT_REQUIRED';           Kind = 'Reboot'; Meaning = 'A restart is required before this can install.' }
    '0x80070070' = @{ Name = 'ERROR_DISK_FULL';                      Kind = 'Space';  Meaning = 'Not enough disk space.' }
    '0x80070005' = @{ Name = 'E_ACCESSDENIED';                       Kind = 'Access'; Meaning = 'Access denied -- often security software or a permissions change.' }
    '0x80070020' = @{ Name = 'ERROR_SHARING_VIOLATION';              Kind = 'Access'; Meaning = 'A file was in use -- often antivirus scanning it.' }
    '0x80070002' = @{ Name = 'ERROR_FILE_NOT_FOUND';                 Kind = 'Client'; Meaning = 'A file was not found -- usually a damaged download cache.' }
    '0x80070422' = @{ Name = 'ERROR_SERVICE_DISABLED';               Kind = 'Client'; Meaning = 'An update service is disabled.' }
    '0x8024402c' = @{ Name = 'WU_E_PT_WINHTTP_NAME_NOT_RESOLVED';    Kind = 'Client'; Meaning = 'The update server name did not resolve.' }
    '0x80244022' = @{ Name = 'WU_E_PT_HTTP_STATUS_SERVICE_UNAVAIL';  Kind = 'Client'; Meaning = 'The update server returned 503 Service Unavailable.' }
    '0x80240438' = @{ Name = 'WU_E_PT_ENDPOINT_UNREACHABLE';         Kind = 'Client'; Meaning = 'The update service endpoint could not be reached.' }
    '0x80072ee2' = @{ Name = 'ERROR_INTERNET_TIMEOUT';               Kind = 'Client'; Meaning = 'The connection to the update service timed out.' }
    '0x80072efd' = @{ Name = 'ERROR_INTERNET_CANNOT_CONNECT';        Kind = 'Client'; Meaning = 'The update service could not be connected to.' }
    '0x800f0922' = @{ Name = 'CBS_E_INSTALLERS_FAILED';              Kind = 'Other';  Meaning = 'An installer step failed -- commonly a full System Reserved partition, a VPN, or .NET setup.' }
    '0x800705b4' = @{ Name = 'ERROR_TIMEOUT';                        Kind = 'Other';  Meaning = 'The operation timed out.' }
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$SolderFindings = @{
    'StoreRepairable' = @{
        Severity = 'Error'
        Title    = 'Component store is corrupt (repairable)'
        Summary  = 'DISM found corruption in the component store. Updates, features and SFC all draw on the store, so they fail until it is repaired.'
        Remedy   = 'Repair runs DISM /RestoreHealth, then SFC.'
    }
    'StoreNotRepairable' = @{
        Severity = 'Error'
        Title    = 'Component store is corrupt and DISM says it cannot be repaired'
        Summary  = 'DISM reports damage it cannot fix from its current sources.'
        Remedy   = 'Repair with -Source pointing at install media of the same build and edition (WIM:D:\sources\install.wim:<index>). If that fails, an in-place upgrade (setup.exe from the same media, keeping files and apps) rebuilds the store.'
    }
    'StoreCheckFailed' = @{
        Severity = 'Warning'
        Title    = 'Component store health could not be read'
        Summary  = 'DISM did not run, or returned output S.O.L.D.E.R. could not read.'
        Remedy   = 'Run DISM /Online /Cleanup-Image /ScanHealth by hand and read %WINDIR%\Logs\DISM\dism.log.'
    }
    'CleanupRecommended' = @{
        Severity = 'Info'
        Title    = 'Component store cleanup recommended'
        Summary  = 'DISM reports superseded packages that a cleanup would remove, reclaiming space on the system drive.'
        Remedy   = 'Repair with -Cleanup runs DISM /StartComponentCleanup. S.O.L.D.E.R. never runs /ResetBase.'
    }
    'SfcUnrepairedFiles' = @{
        Severity = 'Error'
        Title    = 'System File Checker found files it could not repair'
        Summary  = 'The last sfc /scannow left damaged system files in place, usually because the component store was damaged too.'
        Remedy   = 'Repair runs RestoreHealth first so SFC has a healthy store to copy from, then SFC again.'
    }
    'SfcRepairedFiles' = @{
        Severity = 'Info'
        Title    = 'System File Checker repaired files'
        Summary  = 'The last sfc /scannow found and replaced damaged system files.'
        Remedy   = 'Nothing further unless the same files are damaged again -- recurring damage points at the disk (A.U.G.U.R.) or memory.'
    }
    'RebootPendingServicing' = @{
        Severity = 'Warning'
        Title    = 'A restart is pending for servicing'
        Summary  = 'Windows has servicing work waiting for a restart. Updates and DISM repairs fail or roll back until it happens (CBS_E_PENDING).'
        Remedy   = 'Restart first, then audit again. Repair stops before DISM unless the technician chooses to continue.'
    }
    'RebootPendingOther' = @{
        Severity = 'Info'
        Title    = 'A restart is pending (not servicing)'
        Summary  = 'File renames or a computer rename are waiting for a restart. These do not block servicing on their own.'
        Remedy   = 'Restart at a convenient point.'
    }
    'UpdateFailuresStore' = @{
        Severity = 'Error'
        Title    = 'Updates are failing because of the component store'
        Summary  = 'Recent Windows Update installs failed with codes that point at component-store damage.'
        Remedy   = 'Repair, then retry the update (W.H.E.T.S.T.O.N.E.). If RestoreHealth reports 0x800f081f, give it a -Source.'
    }
    'UpdateFailuresClient' = @{
        Severity = 'Warning'
        Title    = 'Updates are failing in the update client, not the store'
        Summary  = 'Recent Windows Update failures point at connectivity, disabled services or the download cache.'
        Remedy   = 'Run C.O.N.D.U.I.T. -- it repairs the update client. A store repair will not fix these.'
    }
    'UpdateFailuresOther' = @{
        Severity = 'Warning'
        Title    = 'Updates are failing for other reasons'
        Summary  = 'Recent Windows Update installs failed with codes that point at restarts, space, access or installer steps.'
        Remedy   = 'Read each code''s meaning in the update failures table and address it before retrying.'
    }
    'TrustedInstallerDisabled' = @{
        Severity = 'Error'
        Title    = 'Windows Modules Installer service is disabled'
        Summary  = 'TrustedInstaller installs every update and feature. Disabled, all servicing fails.'
        Remedy   = 'Repair sets it back to Manual, its default.'
    }
    'DiskSpaceCritical' = @{
        Severity = 'Error'
        Title    = 'System drive almost out of space'
        Summary  = 'Servicing stages its payloads on the system drive; with this little free space installs and repairs fail.'
        Remedy   = 'Free space first: C.L.E.A.N.S.E. clears temp and update caches.'
    }
    'DiskSpaceLow' = @{
        Severity = 'Warning'
        Title    = 'System drive low on space'
        Summary  = 'Feature updates and large cumulative updates need several GB of free space to stage.'
        Remedy   = 'Run C.L.E.A.N.S.E.; Repair with -Cleanup also reclaims superseded components.'
    }
    'WsusRepairSource' = @{
        Severity = 'Info'
        Title    = 'Repair files will be requested from WSUS'
        Summary  = 'This machine uses WSUS, which does not serve repair content, and policy does not send repairs to Windows Update instead. RestoreHealth then fails with 0x800f081f.'
        Remedy   = 'Repair with -Source pointing at matching install media, or enable "Specify settings for optional component installation and component repair" with "Download repair content directly from Windows Update".'
    }
    'SourceMissing' = @{
        Severity = 'Error'
        Title    = 'RestoreHealth could not find repair files'
        Summary  = 'DISM failed with 0x800f081f: neither Windows Update nor the given source had the files it needed.'
        Remedy   = 'Mount install media of the same build and edition and run Repair with -Source WIM:<drive>:\sources\install.wim:<index> (or ESD:...install.esd:<index>). Get-WindowsImage -ImagePath lists the indexes.'
    }
    'RepairFailed' = @{
        Severity = 'Error'
        Title    = 'Repair step failed'
        Summary  = 'DISM or SFC returned an error. The detail names the step and the code.'
        Remedy   = 'Read %WINDIR%\Logs\DISM\dism.log and CBS.log for the failing component. An in-place upgrade is the fallback when the store cannot be repaired.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[object]]::new()

function Add-SolderFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $SolderFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
    })
}

function Add-SolderAction {
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

function Show-SolderBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███████╗ ██████╗ ██╗     ██████╗ ███████╗██████╗
  ██╔════╝██╔═══██╗██║     ██╔══██╗██╔════╝██╔══██╗
  ███████╗██║   ██║██║     ██║  ██║█████╗  ██████╔╝
  ╚════██║██║   ██║██║     ██║  ██║██╔══╝  ██╔══██╗
  ███████║╚██████╔╝███████╗██████╔╝███████╗██║  ██║
  ╚══════╝ ╚═════╝ ╚══════╝╚═════╝ ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    S.O.L.D.E.R. — Servicing-stack Overhaul: Locates Damage, Enacts Repair" -ForegroundColor Cyan
    Write-Host "    Windows Servicing Diagnosis and Repair Tool" -ForegroundColor Cyan
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

function ConvertTo-HResultString {
    # Normalises an error code to '0x' + eight lower-case hex digits. Accepts
    # '0x800F081F', '800f081f', or the signed decimal a process exit code
    # carries (-2146498529).
    param($Code)
    if ($null -eq $Code) { return '' }
    $text = "$Code".Trim()
    if ($text -match '^(?i)(?:0x)?([0-9a-f]{8})$' -and $text -notmatch '^-?\d+$') { return '0x' + $Matches[1].ToLowerInvariant() }
    if ($text -match '^(?i)0x([0-9a-f]{1,8})$') { return '0x' + $Matches[1].ToLowerInvariant().PadLeft(8, '0') }
    $n = 0L
    if ([long]::TryParse($text, [ref]$n)) {
        if ($n -lt 0) { $n = $n + 4294967296L }
        if ($n -ge 0 -and $n -le 4294967295L) { return '0x{0:x8}' -f $n }
    }
    return ''
}

function Get-ServicingErrorInfo {
    param($Code)
    $hex = ConvertTo-HResultString -Code $Code
    if (-not $hex) { return [PSCustomObject]@{ Code = "$Code"; Name = 'Unknown'; Kind = 'Other'; Meaning = 'No code was reported.' } }
    $entry = $ServicingErrorCodes[$hex]
    if ($entry) { return [PSCustomObject]@{ Code = $hex; Name = $entry.Name; Kind = $entry.Kind; Meaning = $entry.Meaning } }
    return [PSCustomObject]@{ Code = $hex; Name = $hex; Kind = 'Other'; Meaning = 'Not a code S.O.L.D.E.R. recognises -- search the code with "Windows Update".' }
}

function ConvertFrom-DismHealth {
    # Reads DISM /CheckHealth, /ScanHealth or /RestoreHealth output, run with
    # /English so the phrases are stable across UI languages.
    param([string[]]$Lines)
    $text = ($Lines -join "`n")
    $state = 'Unknown'
    if ($text -match '(?i)cannot be repaired')                          { $state = 'NotRepairable' }
    elseif ($text -match '(?i)component store is repairable')           { $state = 'Repairable' }
    elseif ($text -match '(?i)restore operation completed successfully') { $state = 'Repaired' }
    elseif ($text -match '(?i)no component store corruption detected')  { $state = 'Healthy' }
    $err = [regex]::Match($text, '(?i)Error:\s*(0x[0-9a-f]{1,8}|\d+)')
    return [PSCustomObject]@{
        State     = $state
        ErrorCode = $(if ($err.Success) { ConvertTo-HResultString -Code $err.Groups[1].Value } else { '' })
    }
}

function ConvertFrom-DismAnalyze {
    # DISM /AnalyzeComponentStore: "Name : value" lines.
    param([string[]]$Lines)
    $get = {
        param($label)
        foreach ($l in $Lines) {
            $m = [regex]::Match("$l", '^\s*' + [regex]::Escape($label) + '\s*:\s*(.+?)\s*$')
            if ($m.Success) { return $m.Groups[1].Value }
        }
        return ''
    }
    $reclaimable = & $get 'Number of Reclaimable Packages'
    return [PSCustomObject]@{
        ActualSize          = & $get 'Actual Size of Component Store'
        SharedWithWindows   = & $get 'Shared with Windows'
        BackupsAndDisabled  = & $get 'Backups and Disabled Features'
        CacheAndTemp        = & $get 'Cache and Temporary Data'
        LastCleanup         = & $get 'Date of Last Cleanup'
        ReclaimablePackages = $(if ($reclaimable -match '^\d+$') { [int]$reclaimable } else { $null })
        CleanupRecommended  = ((& $get 'Component Store Cleanup Recommended') -match '(?i)^yes')
    }
}

function Get-CbsQuotedPath {
    # CBS.log writes a path either bare (\??\C:\Windows\...\foo.dll from store)
    # or as quoted, length-prefixed fragments ([l:23]"\??\C:\Windows\System32"\[l:10]"foo.dll").
    # Joins the fragments and drops the \??\ prefix.
    param([string]$Fragment)
    $text = ($Fragment -replace '(?i)\s+(?:from store|of \S.*|; source file.*)$', '').Trim()
    $parts = @([regex]::Matches($text, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value.TrimEnd('\') })
    $path = if ($parts.Count -gt 0) { $parts -join '\' } else { ($text -split '\s+')[0] }
    return ($path -replace '^\\\?\?\\', '')
}

function ConvertFrom-CbsSrLine {
    # System File Checker writes [SR] lines to CBS.log. Returns the files it
    # repaired and the files it could not, de-duplicated, plus the time of the
    # last [SR] line.
    param([string[]]$Lines)
    $repaired = [System.Collections.Generic.List[string]]::new()
    $failed   = [System.Collections.Generic.List[string]]::new()
    $last     = ''
    $sawSr    = $false
    foreach ($l in $Lines) {
        $line = "$l"
        if ($line -notmatch '\[SR\]') { continue }
        $sawSr = $true
        $ts = [regex]::Match($line, '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})')
        if ($ts.Success) { $last = $ts.Groups[1].Value }
        $cannot = [regex]::Match($line, '(?i)(?:Cannot repair member file|Could not reproject corrupted file)\s+(.+)$')
        if ($cannot.Success) {
            $name = Get-CbsQuotedPath -Fragment $cannot.Groups[1].Value
            if ($name -and -not $failed.Contains($name)) { [void]$failed.Add($name) }
            continue
        }
        $fixed = [regex]::Match($line, '(?i)Repairing corrupted file\s+(.+)$')
        if ($fixed.Success) {
            $name = Get-CbsQuotedPath -Fragment $fixed.Groups[1].Value
            if ($name -and -not $repaired.Contains($name)) { [void]$repaired.Add($name) }
        }
    }
    return [PSCustomObject]@{
        Found        = $sawSr
        Repaired     = $repaired.ToArray()
        CannotRepair = $failed.ToArray()
        LastRun      = $last
    }
}

function ConvertFrom-SfcOutput {
    # sfc writes UTF-16 to the console, which arrives with a NUL after every
    # character; strip them before matching. English text only -- CBS.log is
    # the language-neutral source, this is the fallback.
    param([string[]]$Lines)
    $text = (($Lines -join "`n") -replace "`0", '')
    if ($text -match '(?i)did not find any integrity violations')  { return 'Clean' }
    if ($text -match '(?i)unable to fix some of them')             { return 'Unrepaired' }
    if ($text -match '(?i)successfully repaired')                  { return 'Repaired' }
    if ($text -match '(?i)could not perform the requested operation') { return 'Failed' }
    return 'Unknown'
}

function Get-PendingRebootReason {
    # Takes the signal readings and returns one row per pending restart.
    # Servicing = blocks DISM and update installs; Other = does not.
    param([hashtable]$Signals)
    $rows = [System.Collections.Generic.List[object]]::new()
    $map = [ordered]@{
        CbsRebootPending     = @{ Kind = 'Servicing'; Text = 'Component Based Servicing: RebootPending' }
        CbsPackagesPending   = @{ Kind = 'Servicing'; Text = 'Component Based Servicing: PackagesPending' }
        PendingXml           = @{ Kind = 'Servicing'; Text = 'WinSxS\pending.xml present' }
        WuRebootRequired     = @{ Kind = 'Servicing'; Text = 'Windows Update: RebootRequired' }
        FileRenames          = @{ Kind = 'Other';     Text = 'Session Manager: PendingFileRenameOperations' }
        ComputerRename       = @{ Kind = 'Other';     Text = 'Computer rename waiting for a restart' }
    }
    foreach ($k in $map.Keys) {
        if ($Signals.ContainsKey($k) -and $Signals[$k]) { [void]$rows.Add([PSCustomObject]@{ Signal = $k; Kind = $map[$k].Kind; Text = $map[$k].Text }) }
    }
    return $rows.ToArray()
}

function Get-RepairSourceState {
    # Where DISM looks for repair files. A WSUS-managed machine asks WSUS,
    # which does not carry repair content, unless policy sends repairs to
    # Windows Update (RepairContentServerSource = 2) or names a local source.
    param([bool]$WsusConfigured, $RepairContentServerSource, [string]$LocalSourcePath, $UseWindowsUpdate)
    $toWu   = ("$RepairContentServerSource" -eq '2')
    $local  = -not [string]::IsNullOrWhiteSpace($LocalSourcePath)
    $neverWu = ("$UseWindowsUpdate" -eq '2')
    $text = if ($local) { "Policy source: $LocalSourcePath" }
            elseif ($WsusConfigured -and $toWu) { 'Windows Update (policy overrides WSUS for repairs)' }
            elseif ($WsusConfigured) { 'WSUS -- which does not serve repair content' }
            elseif ($neverWu) { 'None -- policy blocks Windows Update and names no source' }
            else { 'Windows Update' }
    return [PSCustomObject]@{
        Description = $text
        AtRisk      = (-not $local) -and (($WsusConfigured -and -not $toWu) -or $neverWu)
    }
}

function Get-SolderVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Broken';    Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Attention'; Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Native {
    # Runs a native tool, optionally echoing its output, and returns the lines
    # and exit code. A missing binary or a non-zero exit never aborts the run.
    param([string]$File, [string[]]$Arguments, [switch]$Echo)
    $lines = [System.Collections.Generic.List[string]]::new()
    $code  = $null
    try {
        & $File @Arguments 2>&1 | ForEach-Object {
            $l = ("$_" -replace "`0", '')
            [void]$lines.Add($l)
            if ($Echo -and -not [string]::IsNullOrWhiteSpace($l)) { Write-Host "      $($l.Trim())" -ForegroundColor $C.Info }
        }
        $code = $LASTEXITCODE
    } catch {
        [void]$lines.Add("failed: $($_.Exception.Message)")
        $code = -1
    }
    return [PSCustomObject]@{ Lines = $lines.ToArray(); ExitCode = $code }
}

function Get-EventsSafe {
    param([hashtable]$Filter, [int]$MaxEvents = 200)
    try { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $MaxEvents -ErrorAction Stop) } catch { return @() }
}

function Get-EventDataMap {
    param($Record)
    $map = @{}
    try {
        $xml = [xml]$Record.ToXml()
        foreach ($d in @($xml.Event.EventData.Data)) { if ($d -and $d.Name) { $map[$d.Name] = $d.'#text' } }
    } catch { return $map }
    return $map
}

function Read-LogTail {
    # The last $MaxBytes of a log another process may be writing, as lines.
    param([string]$Path, [long]$MaxBytes, [long]$FromOffset = -1)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $fs = $null; $reader = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
        $start = if ($FromOffset -ge 0) { [math]::Min($FromOffset, $fs.Length) } else { [math]::Max(0, $fs.Length - $MaxBytes) }
        [void]$fs.Seek($start, [System.IO.SeekOrigin]::Begin)
        $reader = [System.IO.StreamReader]::new($fs)
        $text = $reader.ReadToEnd()
        return @($text -split "`r?`n")
    } catch {
        return @()
    } finally {
        if ($reader) { $reader.Dispose() } elseif ($fs) { $fs.Dispose() }
    }
}

function Get-LogLength {
    param([string]$Path)
    try { return (Get-Item -LiteralPath $Path -ErrorAction Stop).Length } catch { return 0 }
}

function Get-SolderContext {
    $os = $null
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }
    $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $drive = $null
    try { $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'" -ErrorAction Stop } catch { $drive = $null }
    $build = if ($cv) { "$($cv.CurrentBuild).$($cv.UBR)" } elseif ($os) { "$($os.BuildNumber)" } else { 'Unknown' }
    return [PSCustomObject]@{
        Computer  = $env:COMPUTERNAME
        OS        = $(if ($os) { "$($os.Caption)" } else { 'Unknown' })
        Build     = $build
        Display   = $(if ($cv -and $cv.DisplayVersion) { "$($cv.DisplayVersion)" } else { '' })
        Edition   = $(if ($cv) { "$($cv.EditionID)" } else { '' })
        FreeBytes = $(if ($drive) { [double]$drive.FreeSpace } else { $null })
        SizeBytes = $(if ($drive) { [double]$drive.Size } else { $null })
    }
}

function Get-SolderRebootSignal {
    $sm = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction SilentlyContinue
    $renames = @()
    if ($sm) { $renames = @($sm.PendingFileRenameOperations) + @($sm.PendingFileRenameOperations2) | Where-Object { -not [string]::IsNullOrWhiteSpace("$_") } }
    $active  = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction SilentlyContinue).ComputerName
    $pending = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction SilentlyContinue).ComputerName
    return @{
        CbsRebootPending   = (Test-Path "$CbsKey\RebootPending")
        CbsPackagesPending = (Test-Path "$CbsKey\PackagesPending")
        PendingXml         = (Test-Path (Join-Path $env:SystemRoot 'WinSxS\pending.xml'))
        WuRebootRequired   = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
        FileRenames        = (@($renames).Count -gt 0)
        ComputerRename     = ($active -and $pending -and ($active -ne $pending))
    }
}

function Get-SolderUpdateFailure {
    $since = (Get-Date).AddDays(-$UpdateLookbackDays)
    $rows = foreach ($e in (Get-EventsSafe -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WindowsUpdateClient'; Id = 20; StartTime = $since } -MaxEvents 200)) {
        $d = Get-EventDataMap -Record $e
        $info = Get-ServicingErrorInfo -Code $d['errorCode']
        [PSCustomObject]@{
            Time    = $e.TimeCreated
            Update  = "$($d['updateTitle'])"
            Code    = $info.Code
            Name    = $info.Name
            Kind    = $info.Kind
            Meaning = $info.Meaning
        }
    }
    return @($rows | Sort-Object Time -Descending)
}

function Get-SolderRepairSource {
    $svc = Get-ItemProperty -Path $ServicingPolicyKey -ErrorAction SilentlyContinue
    $wu  = Get-ItemProperty -Path $WuPolicyKey -ErrorAction SilentlyContinue
    $au  = Get-ItemProperty -Path "$WuPolicyKey\AU" -ErrorAction SilentlyContinue
    $wsus = ($wu -and -not [string]::IsNullOrWhiteSpace("$($wu.WUServer)") -and $au -and $au.UseWUServer -eq 1)
    $state = Get-RepairSourceState -WsusConfigured $wsus `
        -RepairContentServerSource $(if ($svc) { $svc.RepairContentServerSource } else { $null }) `
        -LocalSourcePath $(if ($svc) { "$($svc.LocalSourcePath)" } else { '' }) `
        -UseWindowsUpdate $(if ($svc) { $svc.UseWindowsUpdate } else { $null })
    return [PSCustomObject]@{ Wsus = $wsus; WsusServer = $(if ($wu) { "$($wu.WUServer)" } else { '' }); Description = $state.Description; AtRisk = $state.AtRisk }
}

function Get-DismArgumentList {
    param([string]$Operation)
    $a = @('/English', '/Online', '/Cleanup-Image', "/$Operation")
    if ($Operation -eq 'RestoreHealth' -and -not [string]::IsNullOrWhiteSpace($Source)) { $a += @("/Source:$Source", '/LimitAccess') }
    return $a
}

function Invoke-SolderStoreCheck {
    param([switch]$Scan)
    $op = if ($Scan) { 'ScanHealth' } else { 'CheckHealth' }
    if ($Scan) { Write-Step "DISM /ScanHealth -- reads the whole store, usually 5-15 minutes..." } else { Write-Step "DISM /CheckHealth..." }
    $r = Invoke-Native -File 'dism.exe' -Arguments (Get-DismArgumentList -Operation $op) -Echo:$Scan
    $parsed = ConvertFrom-DismHealth -Lines $r.Lines
    return [PSCustomObject]@{ Operation = $op; State = $parsed.State; ErrorCode = $(if ($parsed.ErrorCode) { $parsed.ErrorCode } elseif ($r.ExitCode) { ConvertTo-HResultString -Code $r.ExitCode } else { '' }); ExitCode = $r.ExitCode }
}

function Invoke-SolderAudit {
    Write-Section "SYSTEM"
    $ctx = Get-SolderContext
    Write-Info "$($ctx.Computer)  |  $($ctx.OS) $($ctx.Display)  |  build $($ctx.Build)"
    if ($null -ne $ctx.FreeBytes) {
        $freeGB = [math]::Round($ctx.FreeBytes / 1GB, 1)
        Write-Info "$($env:SystemDrive) free: $(Format-Bytes $ctx.FreeBytes)"
        if ($freeGB -lt $FreeSpaceErrorGB)       { Add-SolderFinding -Code 'DiskSpaceCritical' -Detail "$freeGB GB free on $($env:SystemDrive)" }
        elseif ($freeGB -lt $FreeSpaceWarningGB) { Add-SolderFinding -Code 'DiskSpaceLow'      -Detail "$freeGB GB free on $($env:SystemDrive)" }
    }

    $ti = $null
    try { $ti = Get-CimInstance Win32_Service -Filter "Name='TrustedInstaller'" -ErrorAction Stop } catch { $ti = $null }
    if ($ti -and $ti.StartMode -eq 'Disabled') { Add-SolderFinding -Code 'TrustedInstallerDisabled' -Detail 'TrustedInstaller start type: Disabled' }
    Write-Info ("Windows Modules Installer: {0}" -f $(if ($ti) { "$($ti.StartMode), $($ti.State)" } else { 'not found' }))

    Write-Section "PENDING RESTART"
    $reboot = @(Get-PendingRebootReason -Signals (Get-SolderRebootSignal))
    if ($reboot.Count -eq 0) { Write-Ok "No restart pending." }
    foreach ($r in $reboot) { if ($r.Kind -eq 'Servicing') { Write-Warn $r.Text } else { Write-Info $r.Text } }
    $servicingReboot = @($reboot | Where-Object { $_.Kind -eq 'Servicing' })
    if ($servicingReboot.Count -gt 0) { Add-SolderFinding -Code 'RebootPendingServicing' -Detail (($servicingReboot | ForEach-Object { $_.Text }) -join '; ') }
    elseif ($reboot.Count -gt 0)      { Add-SolderFinding -Code 'RebootPendingOther'     -Detail (($reboot | ForEach-Object { $_.Text }) -join '; ') }

    Write-Section "COMPONENT STORE"
    $store = Invoke-SolderStoreCheck -Scan:$Deep
    switch ($store.State) {
        'Healthy'       { Write-Ok "No component store corruption detected ($($store.Operation))." }
        'Repairable'    { Write-Fail "Component store is corrupt and repairable."; Add-SolderFinding -Code 'StoreRepairable' -Detail "DISM /$($store.Operation)" }
        'NotRepairable' { Write-Fail "Component store is corrupt and cannot be repaired from current sources."; Add-SolderFinding -Code 'StoreNotRepairable' -Detail "DISM /$($store.Operation)" }
        default {
            Write-Warn "Could not read the store state."
            $info = Get-ServicingErrorInfo -Code $store.ErrorCode
            Add-SolderFinding -Code 'StoreCheckFailed' -Detail ("DISM /{0} exit {1}{2}" -f $store.Operation, $store.ExitCode, $(if ($store.ErrorCode) { " -- $($info.Code) $($info.Name)" } else { '' }))
        }
    }
    if (-not $Deep -and $store.State -eq 'Healthy') { Write-Info "CheckHealth reads the last recorded state; -Deep runs a full ScanHealth." }

    Write-Step "DISM /AnalyzeComponentStore..."
    $analyze = ConvertFrom-DismAnalyze -Lines (Invoke-Native -File 'dism.exe' -Arguments @('/English', '/Online', '/Cleanup-Image', '/AnalyzeComponentStore')).Lines
    if ($analyze.ActualSize) { Write-Info ("Store size {0}; reclaimable packages {1}; cleanup recommended: {2}" -f $analyze.ActualSize, $analyze.ReclaimablePackages, $(if ($analyze.CleanupRecommended) { 'yes' } else { 'no' })) }
    if ($analyze.CleanupRecommended) { Add-SolderFinding -Code 'CleanupRecommended' -Detail ("{0} reclaimable package(s); last cleanup {1}" -f $analyze.ReclaimablePackages, $analyze.LastCleanup) }

    $src = Get-SolderRepairSource
    Write-Info "Repair files come from: $($src.Description)"
    if ($src.AtRisk -and [string]::IsNullOrWhiteSpace($Source)) { Add-SolderFinding -Code 'WsusRepairSource' -Detail $src.Description }

    Write-Section "SYSTEM FILE CHECKER (last run)"
    $sfc = ConvertFrom-CbsSrLine -Lines (Read-LogTail -Path $CbsLogPath -MaxBytes $CbsTailBytes)
    if (-not $sfc.Found) { Write-Info "No SFC run found in the recent CBS.log." }
    else {
        Write-Info ("Last SFC activity {0}: {1} repaired, {2} could not be repaired" -f $sfc.LastRun, $sfc.Repaired.Count, $sfc.CannotRepair.Count)
        if ($sfc.CannotRepair.Count -gt 0) { Add-SolderFinding -Code 'SfcUnrepairedFiles' -Detail (($sfc.CannotRepair | Select-Object -First 5) -join ', ') }
        elseif ($sfc.Repaired.Count -gt 0) { Add-SolderFinding -Code 'SfcRepairedFiles' -Detail (($sfc.Repaired | Select-Object -First 5) -join ', ') }
    }

    Write-Section "WINDOWS UPDATE FAILURES ($UpdateLookbackDays days)"
    $failures = @(Get-SolderUpdateFailure)
    if ($failures.Count -eq 0) { Write-Ok "No failed update installs logged." }
    foreach ($f in ($failures | Select-Object -First 8)) { Write-Info ("{0}  {1} {2}  {3}" -f $f.Time.ToString('yyyy-MM-dd'), $f.Code, $f.Kind, $f.Update) }
    foreach ($kind in 'Store', 'Client') {
        $hit = @($failures | Where-Object { $_.Kind -eq $kind })
        if ($hit.Count -gt 0) {
            $codes = ($hit | Group-Object Code | ForEach-Object { "$($_.Name) $($_.Group[0].Name) x$($_.Count)" }) -join '; '
            Add-SolderFinding -Code $(if ($kind -eq 'Store') { 'UpdateFailuresStore' } else { 'UpdateFailuresClient' }) -Detail $codes
        }
    }
    $other = @($failures | Where-Object { $_.Kind -notin @('Store', 'Client') })
    if ($other.Count -gt 0) {
        Add-SolderFinding -Code 'UpdateFailuresOther' -Detail (($other | Group-Object Code | ForEach-Object { "$($_.Name) $($_.Group[0].Name) x$($_.Count)" }) -join '; ')
    }

    return [PSCustomObject]@{
        Context      = $ctx
        TrustedInstaller = $ti
        Reboot       = $reboot
        Store        = $store
        Analyze      = $analyze
        RepairSource = $src
        Sfc          = $sfc
        Failures     = $failures
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Repair-SolderTrustedInstaller {
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would set the Windows Modules Installer (TrustedInstaller) service to Manual."
        Add-SolderAction -Step 'Re-enable Windows Modules Installer' -Status 'WhatIf'
        return
    }
    try {
        Set-Service -Name 'TrustedInstaller' -StartupType Manual -ErrorAction Stop
        Write-Ok "Windows Modules Installer set to Manual."
        Add-SolderAction -Step 'Re-enable Windows Modules Installer' -Status 'Done' -Detail 'Start type: Manual'
        Add-TKNote -Text 'Set the Windows Modules Installer service back to Manual.' -Category 'Action' -ScriptName 'solder'
    } catch {
        Write-Fail "Could not change the service: $($_.Exception.Message)"
        Add-SolderAction -Step 'Re-enable Windows Modules Installer' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Repair-SolderRestoreHealth {
    $argList = Get-DismArgumentList -Operation 'RestoreHealth'
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: dism.exe $($argList -join ' ')"
        Add-SolderAction -Step 'DISM /RestoreHealth' -Status 'WhatIf' -Detail ($argList -join ' ')
        return $true
    }
    Write-Section "REPAIR — DISM /RestoreHealth"
    Write-Step "Repairing the component store$(if ($Source) { " from $Source" } else { ' from Windows Update' }). This usually takes 10-30 minutes; the percentage can sit still for a while."
    $r = Invoke-Native -File 'dism.exe' -Arguments $argList -Echo
    $parsed = ConvertFrom-DismHealth -Lines $r.Lines
    $code = if ($parsed.ErrorCode) { $parsed.ErrorCode } elseif ($r.ExitCode) { ConvertTo-HResultString -Code $r.ExitCode } else { '' }
    if ($r.ExitCode -eq 0 -and $parsed.State -ne 'NotRepairable') {
        Write-Ok "RestoreHealth completed."
        Add-SolderAction -Step 'DISM /RestoreHealth' -Status 'Done' -Detail $(if ($Source) { "Source: $Source" } else { 'Source: Windows Update' })
        Add-TKNote -Text "Ran DISM /RestoreHealth$(if ($Source) { " from $Source" }) -- completed." -Category 'Action' -ScriptName 'solder'
        return $true
    }
    $info = Get-ServicingErrorInfo -Code $code
    Write-Fail "RestoreHealth failed: $($info.Code) $($info.Name) -- $($info.Meaning)"
    if ($info.Code -eq '0x800f081f') { Add-SolderFinding -Code 'SourceMissing' -Detail $(if ($Source) { "Source tried: $Source" } else { 'Source tried: Windows Update' }) }
    else                              { Add-SolderFinding -Code 'RepairFailed' -Detail "DISM /RestoreHealth: $($info.Code) $($info.Name)" }
    Add-SolderAction -Step 'DISM /RestoreHealth' -Status 'Failed' -Detail "$($info.Code) $($info.Name) -- $($info.Meaning)"
    Write-TKError -ScriptName 'solder' -Message "DISM /RestoreHealth failed: $($info.Code) $($info.Name)" -Category 'Servicing'
    return $false
}

function Repair-SolderSfc {
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: sfc.exe /scannow"
        Add-SolderAction -Step 'sfc /scannow' -Status 'WhatIf'
        return $null
    }
    Write-Section "REPAIR — sfc /scannow"
    Write-Step "Verifying system files against the store. This usually takes 10-20 minutes."
    $offset = Get-LogLength -Path $CbsLogPath
    $r = Invoke-Native -File 'sfc.exe' -Arguments @('/scannow')
    $fromLog = ConvertFrom-CbsSrLine -Lines (Read-LogTail -Path $CbsLogPath -MaxBytes $CbsTailBytes -FromOffset $offset)
    $outcome = if ($fromLog.Found) {
        if ($fromLog.CannotRepair.Count -gt 0) { 'Unrepaired' } elseif ($fromLog.Repaired.Count -gt 0) { 'Repaired' } else { 'Clean' }
    } else { ConvertFrom-SfcOutput -Lines $r.Lines }

    switch ($outcome) {
        'Clean'      { Write-Ok "SFC found no integrity violations."; $status = 'Done'; $detail = 'No integrity violations' }
        'Repaired'   { Write-Ok "SFC repaired damaged files."; $status = 'Done'; $detail = "Repaired: $(($fromLog.Repaired | Select-Object -First 5) -join ', ')" }
        'Unrepaired' {
            Write-Fail "SFC could not repair some files."
            $status = 'Failed'; $detail = "Could not repair: $(($fromLog.CannotRepair | Select-Object -First 5) -join ', ')"
            Add-SolderFinding -Code 'RepairFailed' -Detail "sfc /scannow: $detail"
        }
        default {
            Write-Warn "Could not read the SFC result (exit $($r.ExitCode)); check CBS.log."
            $status = 'Failed'; $detail = "Result unreadable, exit $($r.ExitCode)"
        }
    }
    Add-SolderAction -Step 'sfc /scannow' -Status $status -Detail $detail
    Add-TKNote -Text "Ran sfc /scannow -- $detail." -Category 'Action' -ScriptName 'solder'
    return $outcome
}

function Repair-SolderCleanup {
    $argList = @('/English', '/Online', '/Cleanup-Image', '/StartComponentCleanup')
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: dism.exe $($argList -join ' ')  (never /ResetBase)"
        Add-SolderAction -Step 'DISM /StartComponentCleanup' -Status 'WhatIf'
        return
    }
    Write-Section "REPAIR — component cleanup"
    Write-Step "Removing superseded components. Can take 10+ minutes."
    $r = Invoke-Native -File 'dism.exe' -Arguments $argList -Echo
    if ($r.ExitCode -eq 0) {
        Write-Ok "Component cleanup completed."
        Add-SolderAction -Step 'DISM /StartComponentCleanup' -Status 'Done'
        Add-TKNote -Text 'Ran DISM /StartComponentCleanup.' -Category 'Action' -ScriptName 'solder'
    } else {
        $info = Get-ServicingErrorInfo -Code $r.ExitCode
        Write-Warn "Component cleanup failed: $($info.Code) $($info.Name)"
        Add-SolderAction -Step 'DISM /StartComponentCleanup' -Status 'Failed' -Detail "$($info.Code) $($info.Name)"
    }
}

function Invoke-SolderRepair {
    param([object]$Audit)
    $codes = @($Findings | ForEach-Object { $_.Code })

    if ($codes -contains 'TrustedInstallerDisabled') { Repair-SolderTrustedInstaller }

    if ($codes -contains 'RebootPendingServicing' -and -not $WhatIf) {
        $proceed = $false
        if (-not $Unattended) {
            Write-Warn "A restart is pending for servicing. DISM repairs usually fail (CBS_E_PENDING) until it happens."
            $ans = Read-Host "  Continue with the repair anyway? [y/N]"
            $proceed = ($ans -match '^[Yy]')
        }
        if (-not $proceed) {
            Write-Warn "Restart the machine, then run the repair again."
            Add-SolderAction -Step 'DISM /RestoreHealth and sfc /scannow' -Status 'Skipped' -Detail 'A servicing restart is pending -- restart first.'
            return
        }
    }

    if ($codes -contains 'DiskSpaceCritical') { Write-Warn "The system drive is almost full; DISM may fail for lack of space." }

    $restored = Repair-SolderRestoreHealth
    # SFC copies from the store, so it is only worth running once the store is
    # sound; after a failed RestoreHealth it would report the same damage.
    if ($restored) { [void](Repair-SolderSfc) }
    else           { Add-SolderAction -Step 'sfc /scannow' -Status 'Skipped' -Detail 'RestoreHealth failed; SFC would copy from a damaged store.' }

    $doCleanup = [bool]$Cleanup
    if (-not $doCleanup -and -not $Unattended -and -not $WhatIf -and $Audit.Analyze.CleanupRecommended -and $restored) {
        $ans = Read-Host "  DISM recommends a component cleanup. Run /StartComponentCleanup now? [y/N]"
        $doCleanup = ($ans -match '^[Yy]')
    }
    if ($doCleanup) { Repair-SolderCleanup }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-SolderReport {
    param([object]$Audit, [object]$Verdict, [object]$After)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $ctx        = $Audit.Context

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>No servicing issues found.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
            $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $storeBadge = switch ($Audit.Store.State) { 'Healthy' { 'ok' } 'Repaired' { 'ok' } 'Repairable' { 'err' } 'NotRepairable' { 'err' } default { 'warn' } }
    $afterText  = if ($After) { "$($After.State) (DISM /$($After.Operation) after repair)" } else { '' }

    $rRows = [System.Text.StringBuilder]::new()
    if ($Audit.Reboot.Count -eq 0) { [void]$rRows.Append("<tr><td colspan='2'>No restart pending.</td></tr>") }
    foreach ($r in $Audit.Reboot) {
        $badge = if ($r.Kind -eq 'Servicing') { "<span class='tk-badge-warn'>Blocks servicing</span>" } else { "<span class='tk-badge-info'>Other</span>" }
        [void]$rRows.Append("<tr><td>$(EscHtml $r.Text)</td><td>$badge</td></tr>")
    }

    $sfcRows = [System.Text.StringBuilder]::new()
    foreach ($file in $Audit.Sfc.CannotRepair) { [void]$sfcRows.Append("<tr><td class='tk-mono'>$(EscHtml $file)</td><td><span class='tk-badge-err'>Could not repair</span></td></tr>") }
    foreach ($file in $Audit.Sfc.Repaired)     { [void]$sfcRows.Append("<tr><td class='tk-mono'>$(EscHtml $file)</td><td><span class='tk-badge-ok'>Repaired</span></td></tr>") }
    if ($sfcRows.Length -eq 0) {
        [void]$sfcRows.Append("<tr><td colspan='2'>$(if ($Audit.Sfc.Found) { 'The last SFC run recorded no repaired or unrepairable files.' } else { 'No SFC run found in the recent CBS.log.' })</td></tr>")
    }

    $uRows = [System.Text.StringBuilder]::new()
    if ($Audit.Failures.Count -eq 0) { [void]$uRows.Append("<tr><td colspan='4'>No failed update installs in the last $UpdateLookbackDays days.</td></tr>") }
    foreach ($u in ($Audit.Failures | Select-Object -First 40)) {
        $kindBadge = switch ($u.Kind) { 'Store' { 'err' } 'Client' { 'warn' } default { 'info' } }
        [void]$uRows.Append(
            "<tr><td class='tk-mono'>$(EscHtml ($u.Time.ToString('yyyy-MM-dd HH:mm')))</td><td>$(EscHtml $u.Update)</td>" +
            "<td class='tk-mono'>$(EscHtml $u.Code)<br/>$(EscHtml $u.Name)</td>" +
            "<td><span class='tk-badge-$kindBadge'>$(EscHtml $u.Kind)</span> $(EscHtml $u.Meaning)</td></tr>")
    }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Actions.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>") }
    foreach ($a in $Actions) {
        $badge = switch ($a.Status) { 'Done' { 'ok' } 'WhatIf' { 'blue' } 'Failed' { 'err' } default { 'info' } }
        [void]$aRows.Append("<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td><td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>")
    }

    $an = $Audit.Analyze
    $freeText = if ($null -ne $ctx.FreeBytes) { "$(Format-Bytes $ctx.FreeBytes) of $(Format-Bytes $ctx.SizeBytes)" } else { 'unknown' }
    $tiText   = if ($Audit.TrustedInstaller) { "$($Audit.TrustedInstaller.StartMode), $($Audit.TrustedInstaller.State)" } else { 'not found' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'S.O.L.D.E.R. Servicing Report' `
        -ScriptName 'S.O.L.D.E.R.' `
        -Subtitle   "${orgPrefix}Windows Servicing -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Build'     = "$($ctx.Build)$(if ($ctx.Display) { " ($($ctx.Display))" })"
            'Generated' = $reportDate
            'Mode'      = $(if ($WhatIf) { "$Action (dry run)" } else { $Action })
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Component Store', 'Pending Restart', 'System Files', 'Update Failures', 'Actions Taken')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Servicing</div></div>
    <div class="tk-summary-card $storeBadge"><div class="tk-summary-num">$(EscHtml $Audit.Store.State)</div><div class="tk-summary-lbl">Component Store</div></div>
    <div class="tk-summary-card $(if (@($Audit.Reboot | Where-Object { $_.Kind -eq 'Servicing' }).Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$(if (@($Audit.Reboot | Where-Object { $_.Kind -eq 'Servicing' }).Count) { 'Yes' } else { 'No' })</div><div class="tk-summary-lbl">Servicing Restart Pending</div></div>
    <div class="tk-summary-card $(if ($Audit.Sfc.CannotRepair.Count) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($Audit.Sfc.CannotRepair.Count)</div><div class="tk-summary-lbl">Unrepaired Files</div></div>
    <div class="tk-summary-card $(if ($Audit.Failures.Count) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Audit.Failures.Count)</div><div class="tk-summary-lbl">Failed Updates ($UpdateLookbackDays d)</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Component Store</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">Health</span> $(EscHtml $Audit.Store.State) (DISM /$(EscHtml $Audit.Store.Operation))$(if ($afterText) { "<br/><span class='tk-info-label'>After repair</span> $(EscHtml $afterText)" })<br/>
      <span class="tk-info-label">Store size</span> $(EscHtml $(if ($an.ActualSize) { $an.ActualSize } else { 'not read' }))$(if ($an.BackupsAndDisabled) { " (backups and disabled features $(EscHtml $an.BackupsAndDisabled); cache and temp $(EscHtml $an.CacheAndTemp))" })<br/>
      <span class="tk-info-label">Cleanup</span> $(if ($an.CleanupRecommended) { 'Recommended' } else { 'Not recommended' })$(if ($null -ne $an.ReclaimablePackages) { " -- $($an.ReclaimablePackages) reclaimable package(s)" })$(if ($an.LastCleanup) { "; last cleanup $(EscHtml $an.LastCleanup)" })<br/>
      <span class="tk-info-label">Repair files from</span> $(EscHtml $(if ($Source) { "-Source $Source" } else { $Audit.RepairSource.Description }))<br/>
      <span class="tk-info-label">Windows Modules Installer</span> $(EscHtml $tiText)<br/>
      <span class="tk-info-label">System drive free</span> $(EscHtml $freeText)<br/>
      <span class="tk-info-label">Logs</span> <span class="tk-mono">$(EscHtml $DismLogPath)</span> · <span class="tk-mono">$(EscHtml $CbsLogPath)</span>
    </div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Pending Restart</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Signal</th><th>Effect</th></tr></thead>
      <tbody>$($rRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> System Files</div>
    <div class="tk-info-box"><span class="tk-info-label">Last SFC activity</span> $(EscHtml $(if ($Audit.Sfc.LastRun) { $Audit.Sfc.LastRun } else { 'none found' }))</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>File</th><th>Result</th></tr></thead>
      <tbody>$($sfcRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Update Failures</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>When</th><th>Update</th><th>Code</th><th>Cause</th></tr></thead>
      <tbody>$($uRows.ToString())</tbody></table></div>
    <div class="tk-info-box"><span class="tk-info-label">Reading the cause</span> Store = component-store damage, fixed here. Client = the update client's connectivity or services, fixed by C.O.N.D.U.I.T. Reboot, Space and Access are fixed by a restart, free space, and security-software exclusions.</div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Actions Taken</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'S.O.L.D.E.R. v6.0')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-SolderRun {
    param([string]$Mode)
    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-SolderAudit
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Component store, system files and servicing state all in order." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    $after = $null
    if ($Mode -eq 'Repair') {
        Invoke-SolderRepair -Audit $audit
        $ran = @($Actions | Where-Object { $_.Step -eq 'DISM /RestoreHealth' -and $_.Status -eq 'Done' }).Count -gt 0
        if ($ran) {
            Write-Section "VERIFY"
            $after = Invoke-SolderStoreCheck
            if ($after.State -eq 'Healthy') { Write-Ok "Component store reports healthy after repair." }
            else                            { Write-Warn "Component store still reports: $($after.State)" }
            Add-TKNote -Text "SOLDER repair on $($env:COMPUTERNAME): component store now $($after.State)." -Category $(if ($after.State -eq 'Healthy') { 'Resolution' } else { 'Issue' }) -ScriptName 'solder'
        }
    }

    $verdict = Get-SolderVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    if ($after -and $after.State -eq 'Healthy' -and $verdict.Class -ne 'ok') { Write-Info "Findings describe the state before the repair; the store now reports healthy. Audit again after a restart." }
    Add-TKNote -Text ("SOLDER {0} on {1}: verdict {2} ({3} finding(s))." -f $Mode, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'solder'

    Write-Step "Generating HTML report..."
    $html    = Build-SolderReport -Audit $audit -Verdict $verdict -After $after
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("SOLDER_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'solder' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-SolderBanner
    Invoke-SolderRun -Mode $Action
} else {
    $choice = ''
    do {
        Show-SolderBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only: store health (CheckHealth), last SFC result, update failures" -ForegroundColor $C.Info
        Write-Host "  [2] Deep audit  -  as [1] with a full DISM /ScanHealth (5-15 minutes)" -ForegroundColor $C.Info
        Write-Host "  [3] Repair  -  audit, then DISM /RestoreHealth, sfc /scannow, and re-check" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($Source) { Write-Host "  Repair source: $Source" -ForegroundColor $C.Info }
        if ($WhatIf) { Write-Host "  Dry run is active — option 3 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-SolderRun -Mode 'Audit' }
            '2' { $script:Deep = $true; Invoke-SolderRun -Mode 'Audit'; $script:Deep = $false }
            '3' { Invoke-SolderRun -Mode 'Repair' }
            'Q' { Write-Host ""; Write-Host "  Closing S.O.L.D.E.R." -ForegroundColor $C.Header; Write-Host "" }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1, 2, 3 or Q." -ForegroundColor $C.Warning
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
$null = Save-TKNoteReport -Prefix 'SOLDER' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
