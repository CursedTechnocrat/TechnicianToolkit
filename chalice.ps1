# chalice.ps1 - C.H.A.L.I.C.E. — Checks Health, Activation, Logins & Identity of Click-to-run Editions
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
    C.H.A.L.I.C.E. — Checks Health, Activation, Logins & Identity of Click-to-run Editions
    Microsoft 365 Apps Client Diagnosis and Repair Tool for PowerShell 5.1+

.DESCRIPTION
    Answers the desk-side Microsoft 365 Apps tickets -- "Outlook keeps asking
    for my password", "Word says Unlicensed Product", "Teams won't sign in" --
    on the machine, for the signed-in user:

      - The Office install: Click-to-Run version, products and their support
        dates, update channel (and any policy override), whether updates are
        enabled and their source reachable, and when Office last updated
      - Activation: shared computer activation, and whether this user holds a
        Microsoft 365 license token
      - Sign-in: the accounts Office has cached (work vs personal, how many
        tenants), the modern-auth / Web Account Manager settings that cause
        password loops when switched off, the AAD token broker, cached Office
        credentials, and the device's Entra join state and Primary Refresh
        Token (dsregcmd /status)
      - Teams: new vs classic client and the size of its cache

    Repairs, each on request and each previewed by -WhatIf:
      ResetSignIn  -- clears Office's cached identities, license tokens and
                      stored Office credentials (backed up to .reg first) and
                      re-registers a missing token broker. The user signs in
                      to Office again; this is Microsoft's documented reset
                      for activation and sign-in loops.
      ResetTeams   -- closes Teams and clears its cache.
      QuickRepair  -- Office Quick Repair (offline; needs elevation, which it
                      requests).
      Update       -- starts a Click-to-Run update now.

    Runs as the signed-in user rather than elevating: Office's identities,
    licenses and caches live in that user's profile, and an elevated run as a
    different account would read and reset the wrong one. Server-side mailbox
    settings are R.A.V.E.N.'s; Outlook data files are E.X.H.U.M.E.'s.

.USAGE
    PS C:\> .\chalice.ps1                                 # Interactive menu
    PS C:\> .\chalice.ps1 -Unattended                     # Read-only audit + HTML report
    PS C:\> .\chalice.ps1 -Unattended -Action ResetSignIn # Audit, then reset Office sign-in and activation
    PS C:\> .\chalice.ps1 -Action ResetTeams -WhatIf      # Preview the Teams cache reset

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [ValidateSet('Audit', 'ResetSignIn', 'ResetTeams', 'QuickRepair', 'Update')]
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
        Logs a structured error to the central toolkit error log and optionally
        posts to a Teams incoming webhook configured as TeamsWebhook in config.json.
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

    # Optional Teams webhook notification
    try {
        $cfg = Get-TKConfig
        $webhook = $cfg.TeamsWebhook
        if (-not [string]::IsNullOrWhiteSpace($webhook)) {
            $card = @{
                '@type'      = 'MessageCard'
                '@context'   = 'http://schema.org/extensions'
                themeColor   = 'FF0000'
                summary      = "TechnicianToolkit error in $ScriptName"
                sections     = @(@{
                    activityTitle    = "TechnicianToolkit - $ScriptName [$Category]"
                    activitySubtitle = "$($entry.Host) / $($entry.User)  |  $($entry.Timestamp)"
                    activityText     = $Message
                })
            } | ConvertTo-Json -Depth 5

            Invoke-RestMethod -Uri $webhook -Method Post -Body $card `
                -ContentType 'application/json' -ErrorAction Stop | Out-Null
        }
    }
    catch { <# webhook failures are silent - never interrupt the caller #> }
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

$C2RConfigKey      = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$OfficeUpdatePolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'
$IdentityKey       = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity'
$IdentityPolicyKey = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Identity'
$LicensingKey      = 'HKCU:\Software\Microsoft\Office\16.0\Common\Licensing'
$LicenseTokenDir   = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\Licenses'
$ScaTokenDir       = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Licensing'
$NewTeamsCacheDir  = Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams'
$ClassicTeamsDir   = Join-Path $env:APPDATA 'Microsoft\Teams'
$ClassicTeamsExe   = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\current\Teams.exe'
$C2RClientDir      = Join-Path $env:CommonProgramFiles 'microsoft shared\ClickToRun'
$BrokerManifest    = Join-Path $env:SystemRoot 'SystemApps\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\Appxmanifest.xml'
$BrokerPackageDir  = Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy'
$NewTeamsPackageDir = Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe'

$StaleBuildDays      = 90     # Office binaries untouched this long = updates are not landing
$SupportWarningDays  = 180
$PrtStaleDays        = 7
$TeamsCacheInfoBytes = 1GB

# Office apps that hold the identity and license caches open.
$OfficeProcessNames = @('winword', 'excel', 'powerpnt', 'outlook', 'onenote', 'msaccess', 'mspub', 'visio', 'winproj', 'lync')

# Classic Teams cache folders Microsoft documents clearing.
$ClassicTeamsCacheFolders = @('application cache', 'blob_storage', 'Cache', 'Code Cache', 'databases', 'GPUCache', 'IndexedDB', 'Local Storage', 'tmp')

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Click-to-Run update channels, by the GUID at the end of the CDN URL.
$ChannelGuids = @{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current Channel'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current Channel (Preview)'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise Channel'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise Channel'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise Channel (Preview)'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta Channel'
    'f2e724c1-748f-4b47-8fb8-8e0d210e9208' = 'Office 2019 Perpetual Enterprise'
    '5030841d-c919-4594-8d2d-84ae4f96e58e' = 'Office LTSC 2021 Perpetual Enterprise'
}

# The officeupdate policy's updatebranch values.
$UpdateBranchNames = @{
    'current'              = 'Current Channel'
    'firstreleasecurrent'  = 'Current Channel (Preview)'
    'monthlyenterprise'    = 'Monthly Enterprise Channel'
    'deferred'             = 'Semi-Annual Enterprise Channel'
    'firstreleasedeferred' = 'Semi-Annual Enterprise Channel (Preview)'
    'insiderfast'          = 'Beta Channel'
}

# End of support by product family. Microsoft 365 Apps is a subscription and
# stays supported on a supported channel, so it has no date.
$OfficeEndOfSupport = @{
    'Office 2016' = '2025-10-14'
    'Office 2019' = '2025-10-14'
    'Office 2021' = '2026-10-13'
    'Office 2024' = '2029-10-09'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$ChaliceFindings = @{
    'NotClickToRun' = @{
        Severity = 'Info'
        Title    = 'No Click-to-Run Office found'
        Summary  = 'Microsoft 365 Apps (Click-to-Run) is not installed, so the install and update checks were skipped.'
        Remedy   = 'Install Microsoft 365 Apps (C.O.N.J.U.R.E. can deploy it) if this user needs it.'
    }
    'MsiOffice' = @{
        Severity = 'Error'
        Title    = 'MSI-based Office installed'
        Summary  = 'An MSI (Windows Installer) edition of Office 2016 or earlier is installed. These are out of support and do not activate with a Microsoft 365 subscription.'
        Remedy   = 'Remove it and install Microsoft 365 Apps.'
    }
    'OutOfSupport' = @{
        Severity = 'Error'
        Title    = 'Office product past end of support'
        Summary  = 'A perpetual Office product installed here no longer receives security updates.'
        Remedy   = 'Move the user to Microsoft 365 Apps or a supported Office LTSC release.'
    }
    'SupportEndingSoon' = @{
        Severity = 'Warning'
        Title    = 'Office product support ends soon'
        Summary  = 'A perpetual Office product installed here reaches end of support within six months.'
        Remedy   = 'Plan the move to Microsoft 365 Apps or a supported Office LTSC release.'
    }
    'UpdatesDisabled' = @{
        Severity = 'Warning'
        Title    = 'Office updates are turned off'
        Summary  = 'Click-to-Run updates are disabled locally or by policy, so the install falls behind on security and fixes -- and older builds fail to sign in as services retire old protocols.'
        Remedy   = 'Re-enable automatic updates (File > Account > Update Options), or fix the policy (enableautomaticupdates) at its source.'
    }
    'UpdateSourceUnreachable' = @{
        Severity = 'Warning'
        Title    = 'Office update source unreachable'
        Summary  = 'Office is set to update from a network share or path this machine cannot reach, so it never updates.'
        Remedy   = 'Restore the share, or clear the update path (policy updatepath / ClickToRun UpdateUrl) so Office updates from the Microsoft CDN.'
    }
    'StaleBuild' = @{
        Severity = 'Warning'
        Title    = 'Office has not updated in months'
        Summary  = 'The Office program files have not changed in over 90 days. Every supported channel ships at least monthly fixes, so updates are not landing.'
        Remedy   = 'Run -Action Update; if it does not move, check the update channel, the update source and the ClickToRunSvc service.'
    }
    'NoLicenseToken' = @{
        Severity = 'Warning'
        Title    = 'No Microsoft 365 license token for this user'
        Summary  = 'Microsoft 365 Apps is installed but this user has no license token, so Office shows Unlicensed Product or asks to activate.'
        Remedy   = 'Sign in to Office with the licensed work account (File > Account). If it still says Unlicensed, run -Action ResetSignIn and sign in again; confirm the license in A.L.M.A.N.A.C.'
    }
    'PersonalAccountOnly' = @{
        Severity = 'Warning'
        Title    = 'Office is signed in with a personal account only'
        Summary  = 'Every account Office has cached is a personal Microsoft account, but a business Microsoft 365 product is installed -- it will not activate against a personal account.'
        Remedy   = 'Sign out of the personal account in File > Account and sign in with the work account.'
    }
    'MultipleAccounts' = @{
        Severity = 'Info'
        Title    = 'Office has several accounts cached'
        Summary  = 'Office holds more than one account, or work accounts from more than one tenant. Office can pick the wrong one for licensing or for opening files.'
        Remedy   = 'Remove the accounts the user does not need in File > Account, or run -Action ResetSignIn.'
    }
    'ModernAuthDisabled' = @{
        Severity = 'Error'
        Title    = 'Modern authentication is turned off for Office'
        Summary  = 'EnableADAL = 0 forces legacy authentication, which Microsoft 365 no longer accepts -- the classic cause of endless password prompts.'
        Remedy   = 'Remove EnableADAL (or set it to 1) under Office\16.0\Common\Identity, and fix the policy that set it.'
    }
    'WamDisabled' = @{
        Severity = 'Warning'
        Title    = 'Office is told not to use Web Account Manager'
        Summary  = 'DisableAADWAM or DisableADALatopWAMOverride = 1 keeps Office off the Windows token broker. Microsoft no longer supports this; it causes repeated prompts and breaks single sign-on.'
        Remedy   = 'Remove the value under Office\16.0\Common\Identity, and fix the policy that set it.'
    }
    'BrokerPluginMissing' = @{
        Severity = 'Error'
        Title    = 'AAD token broker missing for this user'
        Summary  = 'The Microsoft.AAD.BrokerPlugin package that signs Office and Teams in to work accounts is not registered for this user. Sign-in fails or loops.'
        Remedy   = '-Action ResetSignIn re-registers it from the system copy.'
    }
    'NoPrimaryRefreshToken' = @{
        Severity = 'Error'
        Title    = 'No Primary Refresh Token on an Entra-joined device'
        Summary  = 'The device is Entra (or hybrid) joined but this user has no PRT, so single sign-on to Office and Teams fails and apps prompt for credentials.'
        Remedy   = 'Sign out of Windows and back in with the work account. If that does not issue a PRT, check the device in Entra (enabled, not deleted) and the time on the machine.'
    }
    'StalePrimaryRefreshToken' = @{
        Severity = 'Warning'
        Title    = 'Primary Refresh Token not renewed recently'
        Summary  = 'The PRT is normally renewed every few hours in use; this one has not been renewed in over a week.'
        Remedy   = 'Lock and unlock, or sign out and in, while connected to the internet. Persistent failures point at the device object in Entra.'
    }
    'NotDeviceRegistered' = @{
        Severity = 'Info'
        Title    = 'Device not joined or registered to Entra ID'
        Summary  = 'Office signs in with its own tokens rather than device single sign-on. Not a fault on its own.'
        Remedy   = 'Nothing, unless the organisation expects Entra join (C.O.V.E.N.A.N.T.).'
    }
    'ClassicTeams' = @{
        Severity = 'Warning'
        Title    = 'Classic Teams is still installed'
        Summary  = 'Microsoft retired classic Teams; it no longer signs in reliably and receives no fixes.'
        Remedy   = 'Remove classic Teams and use the new Teams client.'
    }
    'TeamsCacheLarge' = @{
        Severity = 'Info'
        Title    = 'Teams cache is large'
        Summary  = 'The Teams cache has grown past 1 GB. A bloated or damaged cache is a common cause of Teams starting slowly, showing stale data or failing to sign in.'
        Remedy   = '-Action ResetTeams closes Teams and clears it.'
    }
    'SharedComputerActivation' = @{
        Severity = 'Info'
        Title    = 'Shared computer activation is on'
        Summary  = 'Office activates per user session with a short-lived token, as on Remote Desktop hosts. Each user needs a license that allows shared activation.'
        Remedy   = 'Nothing, unless this is a single-user machine -- then switch it off so the license token is cached normally.'
    }
    'DifferentUser' = @{
        Severity = 'Warning'
        Title    = 'Running as a different account than the signed-in user'
        Summary  = 'Office''s accounts, licenses and caches are per user. This run describes the account it runs as, not the person signed in at the console.'
        Remedy   = 'Run C.H.A.L.I.C.E. in the affected user''s own session, without "Run as administrator" under another account.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[object]]::new()

function Add-ChaliceFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $ChaliceFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
    })
}

function Add-ChaliceAction {
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

function Show-ChaliceBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗██╗  ██╗ █████╗ ██╗     ██╗ ██████╗███████╗
  ██╔════╝██║  ██║██╔══██╗██║     ██║██╔════╝██╔════╝
  ██║     ███████║███████║██║     ██║██║     █████╗
  ██║     ██╔══██║██╔══██║██║     ██║██║     ██╔══╝
  ╚██████╗██║  ██║██║  ██║███████╗██║╚██████╗███████╗
   ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚═╝ ╚═════╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    C.H.A.L.I.C.E. — Checks Health, Activation, Logins & Identity of Click-to-run Editions" -ForegroundColor Cyan
    Write-Host "    Microsoft 365 Apps Client Diagnosis and Repair Tool" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host ""
        Write-Host "    *** DRY RUN — no changes will be applied ***" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS — no I/O; the Pester suite calls these with captured data.
# ─────────────────────────────────────────────────────────────────────────────

function Get-OfficeProductFamily {
    # Click-to-Run ProductReleaseIds: O365ProPlusRetail, ProPlus2021Volume,
    # VisioPro2019Retail, ProjectProXVolume (2016 volume), ProPlusRetail (2016).
    param([string]$ReleaseId)
    $id = "$ReleaseId".Trim()
    if (-not $id) { return '' }
    if ($id -match '(?i)^O365')   { return 'Microsoft 365 Apps' }
    if ($id -match '2024')        { return 'Office 2024' }
    if ($id -match '2021')        { return 'Office 2021' }
    if ($id -match '2019')        { return 'Office 2019' }
    if ($id -match '(?i)(Retail|Volume)$') { return 'Office 2016' }
    return 'Other'
}

function Get-OfficeSupportState {
    # 'Supported', 'EndingSoon' or 'Ended' for a product family on a date.
    param([string]$Family, [datetime]$Today, [int]$WarningDays = 180)
    if (-not $OfficeEndOfSupport.ContainsKey($Family)) { return [PSCustomObject]@{ State = 'Supported'; EndDate = $null } }
    $end = [datetime]::ParseExact($OfficeEndOfSupport[$Family], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $state = if ($Today.Date -gt $end) { 'Ended' } elseif (($end - $Today.Date).TotalDays -le $WarningDays) { 'EndingSoon' } else { 'Supported' }
    return [PSCustomObject]@{ State = $state; EndDate = $end }
}

function Get-ChannelName {
    # Resolves a channel from a CDN URL (or bare GUID); policy wins over the
    # local setting when present.
    param([string]$Url, [string]$PolicyBranch = '')
    if (-not [string]::IsNullOrWhiteSpace($PolicyBranch)) {
        $key = $PolicyBranch.Trim().ToLowerInvariant()
        if ($UpdateBranchNames.ContainsKey($key)) { return "$($UpdateBranchNames[$key]) (policy)" }
        return "$PolicyBranch (policy)"
    }
    $m = [regex]::Match("$Url", '(?i)([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})')
    if (-not $m.Success) { return $(if ($Url) { 'Custom source' } else { 'Unknown' }) }
    $guid = $m.Groups[1].Value.ToLowerInvariant()
    if ($ChannelGuids.ContainsKey($guid)) { return $ChannelGuids[$guid] }
    return "Unrecognised channel ($guid)"
}

function Test-UpdatesDisabled {
    # ClickToRun UpdatesEnabled is the string 'False' when turned off in the
    # UI; the policy enableautomaticupdates is DWORD 0.
    param($UpdatesEnabled, $PolicyEnable)
    if ("$PolicyEnable" -eq '0') { return $true }
    if ("$UpdatesEnabled" -match '(?i)^false$') { return $true }
    return $false
}

function ConvertFrom-DsregcmdStatus {
    # dsregcmd /status prints "Name : VALUE" lines under section banners.
    # Returns a case-insensitive hashtable of every pair.
    param([string[]]$Lines)
    $map = @{}
    foreach ($l in $Lines) {
        $m = [regex]::Match("$l", '^\s*([A-Za-z][A-Za-z0-9 ]*?)\s+:\s+(.*?)\s*$')
        if ($m.Success -and -not $map.ContainsKey($m.Groups[1].Value)) { $map[$m.Groups[1].Value] = $m.Groups[2].Value }
    }
    return $map
}

function ConvertFrom-DsregTime {
    # "2026-10-08 13:10:12.000 UTC" -> UTC DateTime, or $null.
    param([string]$Text)
    $m = [regex]::Match("$Text", '(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})')
    if (-not $m.Success) { return $null }
    return [datetime]::ParseExact($m.Groups[1].Value, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
}

function Get-DeviceJoinState {
    param([hashtable]$Status)
    $aad    = ("$($Status['AzureAdJoined'])" -eq 'YES')
    $domain = ("$($Status['DomainJoined'])" -eq 'YES')
    $wpj    = ("$($Status['WorkplaceJoined'])" -eq 'YES')
    $label = if ($aad -and $domain) { 'Hybrid Entra joined' } elseif ($aad) { 'Entra joined' } elseif ($wpj) { 'Entra registered' } elseif ($domain) { 'Domain joined only' } else { 'Not joined' }
    return [PSCustomObject]@{
        Label      = $label
        EntraJoined = $aad
        Registered = $wpj
        HasPrt     = ("$($Status['AzureAdPrt'])" -eq 'YES')
        PrtUpdated = ConvertFrom-DsregTime -Text "$($Status['AzureAdPrtUpdateTime'])"
        Tenant     = "$($Status['TenantName'])"
    }
}

function ConvertFrom-CmdkeyList {
    # The Office credentials in cmdkey /list output. The "Target:" label is
    # translated; the target names are not.
    param([string[]]$Lines)
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($l in $Lines) {
        $m = [regex]::Match("$l", '((?:\w+:target=)?MicrosoftOffice1[56]_Data:\S+)')
        if ($m.Success -and -not $out.Contains($m.Groups[1].Value)) { [void]$out.Add($m.Groups[1].Value) }
    }
    return $out.ToArray()
}

function ConvertTo-IdentityProvider {
    # Office identity keys are named <id>_ADAL / <id>_OrgId for work or school
    # accounts and <id>_LiveId for personal ones, with a ProviderId value to
    # match. Returns 'AD', 'MSA' or the raw value.
    param([string]$ProviderId, [string]$KeyName)
    if ("$ProviderId" -match '(?i)^(AD|ADAL|OrgId)$' -or "$KeyName" -match '(?i)_(ADAL|OrgId)$') { return 'AD' }
    if ("$ProviderId" -match '(?i)^(MSA|LiveId)$' -or "$KeyName" -match '(?i)_LiveId$') { return 'MSA' }
    return "$ProviderId"
}

function Get-IdentitySummary {
    # Classifies Office's cached identities. ProviderId 'AD' = work or school,
    # 'MSA' = personal Microsoft account.
    param([object[]]$Identities)
    $work     = @($Identities | Where-Object { $_.Provider -eq 'AD' })
    $personal = @($Identities | Where-Object { $_.Provider -eq 'MSA' })
    $tenants  = @($work | ForEach-Object { "$($_.TenantId)".ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
    return [PSCustomObject]@{
        Total        = @($Identities).Count
        Work         = $work.Count
        Personal     = $personal.Count
        Tenants      = $tenants.Count
        PersonalOnly = ($personal.Count -gt 0 -and $work.Count -eq 0)
        Mixed        = (@($Identities).Count -gt 1) -and (($personal.Count -gt 0 -and $work.Count -gt 0) -or $tenants.Count -gt 1)
    }
}

function Get-ChaliceVerdict {
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
    param([string]$File, [string[]]$Arguments)
    try {
        $out = & $File @Arguments 2>&1 | ForEach-Object { "$_" }
        return @($out)
    } catch {
        return @("failed: $($_.Exception.Message)")
    }
}

function Get-FolderSize {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return 0 }
    try {
        return [double](@(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue) | Measure-Object -Property Length -Sum).Sum
    } catch { return 0 }
}

function Get-ChaliceContext {
    $console = ''
    try { $console = "$((Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName)" } catch { $console = '' }
    $runAs = "$env:USERDOMAIN\$env:USERNAME"
    return [PSCustomObject]@{
        Computer    = $env:COMPUTERNAME
        RunAs       = $runAs
        ConsoleUser = $console
        Elevated    = Test-IsAdmin
        Mismatch    = (-not [string]::IsNullOrWhiteSpace($console)) -and ($console -ne $runAs)
    }
}

function Get-ChaliceInstall {
    $cfg = Get-ItemProperty -Path $C2RConfigKey -ErrorAction SilentlyContinue
    $pol = Get-ItemProperty -Path $OfficeUpdatePolicy -ErrorAction SilentlyContinue
    $msi = @()
    foreach ($v in '16.0', '15.0', '14.0') {
        $root = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\$v\Common\InstallRoot" -ErrorAction SilentlyContinue
        if ($root -and $root.Path -and -not $cfg) { $msi += "Office $v at $($root.Path)" }
    }
    if (-not $cfg) {
        return [PSCustomObject]@{ ClickToRun = $false; Msi = $msi; Products = @(); Version = ''; Platform = ''; Culture = ''; Channel = ''; UpdatesEnabled = $null
                                  PolicyEnable = $null; UpdatePath = ''; SharedComputer = $false; InstallPath = ''; LastBinaryWrite = $null }
    }

    $products = foreach ($id in ("$($cfg.ProductReleaseIds)" -split ',')) {
        $id = $id.Trim()
        if (-not $id -or $id -match '(?i)LanguagePack|Proofing') { continue }
        $family = Get-OfficeProductFamily -ReleaseId $id
        $support = Get-OfficeSupportState -Family $family -Today (Get-Date) -WarningDays $SupportWarningDays
        [PSCustomObject]@{ Id = $id; Family = $family; Support = $support.State; EndDate = $support.EndDate }
    }

    $installPath = "$($cfg.InstallationPath)"
    $lastWrite = $null
    foreach ($exe in 'WINWORD.EXE', 'EXCEL.EXE', 'OUTLOOK.EXE') {
        $p = Join-Path $installPath "root\Office16\$exe"
        if (Test-Path -LiteralPath $p) { $t = (Get-Item -LiteralPath $p).LastWriteTime; if (-not $lastWrite -or $t -gt $lastWrite) { $lastWrite = $t } }
    }

    $url = if ($cfg.UpdateChannel) { "$($cfg.UpdateChannel)" } else { "$($cfg.CDNBaseUrl)" }
    $updatePath = if ($pol -and $pol.updatepath) { "$($pol.updatepath)" } elseif ($cfg.UpdateUrl) { "$($cfg.UpdateUrl)" } else { '' }
    return [PSCustomObject]@{
        ClickToRun      = $true
        Msi             = $msi
        Products        = @($products)
        Version         = "$($cfg.VersionToReport)"
        Platform        = "$($cfg.Platform)"
        Culture         = "$($cfg.ClientCulture)"
        Channel         = Get-ChannelName -Url $url -PolicyBranch $(if ($pol) { "$($pol.updatebranch)" } else { '' })
        UpdatesEnabled  = $cfg.UpdatesEnabled
        PolicyEnable    = $(if ($pol) { $pol.enableautomaticupdates } else { $null })
        UpdatePath      = $updatePath
        SharedComputer  = ("$($cfg.SharedComputerLicensing)" -eq '1')
        InstallPath     = $installPath
        LastBinaryWrite = $lastWrite
    }
}

function Get-ChaliceIdentity {
    $rows = foreach ($k in @(Get-ChildItem -Path "$IdentityKey\Identities" -ErrorAction SilentlyContinue)) {
        $p = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
        if (-not $p) { continue }
        $email = "$($p.EmailAddress)"
        if (-not $email) { $email = "$($p.SignInName)" }
        [PSCustomObject]@{
            Email    = $email
            Name     = "$($p.FriendlyName)"
            Provider = ConvertTo-IdentityProvider -ProviderId "$($p.ProviderId)" -KeyName $k.PSChildName
            TenantId = "$($p.TenantId)"
        }
    }
    $local  = Get-ItemProperty -Path $IdentityKey -ErrorAction SilentlyContinue
    $policy = Get-ItemProperty -Path $IdentityPolicyKey -ErrorAction SilentlyContinue
    $value = {
        param($name)
        if ($policy -and $null -ne $policy.$name) { return [PSCustomObject]@{ Value = $policy.$name; Source = 'policy' } }
        if ($local -and $null -ne $local.$name)   { return [PSCustomObject]@{ Value = $local.$name;  Source = 'local' } }
        return $null
    }
    return [PSCustomObject]@{
        Identities     = @($rows)
        EnableADAL     = & $value 'EnableADAL'
        DisableAADWAM  = & $value 'DisableAADWAM'
        DisableWamOverride = & $value 'DisableADALatopWAMOverride'
    }
}

function Get-UserPackage {
    # Get-AppxPackage is not available in every PowerShell 7 session; the
    # per-user package folder is the fallback signal when it throws.
    param([string]$Name, [string]$FolderFallback)
    try {
        $pkg = Get-AppxPackage -Name $Name -ErrorAction Stop | Select-Object -First 1
        return [PSCustomObject]@{ Present = [bool]$pkg; Version = $(if ($pkg) { "$($pkg.Version)" } else { '' }) }
    } catch {
        return [PSCustomObject]@{ Present = (Test-Path -LiteralPath $FolderFallback); Version = 'unknown' }
    }
}

function Get-ChaliceBroker {
    $pkg = Get-UserPackage -Name 'Microsoft.AAD.BrokerPlugin' -FolderFallback $BrokerPackageDir
    return [PSCustomObject]@{
        Registered = $pkg.Present
        Version    = $pkg.Version
        CanRepair  = (Test-Path -LiteralPath $BrokerManifest)
    }
}

function Get-ChaliceTeamsClient {
    $newPkg = Get-UserPackage -Name 'MSTeams' -FolderFallback $NewTeamsPackageDir
    $classic = Test-Path -LiteralPath $ClassicTeamsExe
    $newSize = Get-FolderSize -Path $NewTeamsCacheDir
    $classicSize = 0
    if (Test-Path -LiteralPath $ClassicTeamsDir) {
        foreach ($f in $ClassicTeamsCacheFolders) { $classicSize += Get-FolderSize -Path (Join-Path $ClassicTeamsDir $f) }
    }
    return [PSCustomObject]@{
        NewInstalled     = $newPkg.Present
        NewVersion       = $newPkg.Version
        ClassicInstalled = $classic
        NewCacheBytes    = $newSize
        ClassicCacheBytes = $classicSize
    }
}

function Get-OfficeRunningProcess {
    return @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $OfficeProcessNames -contains $_.ProcessName.ToLowerInvariant() })
}

function Invoke-ChaliceAudit {
    Write-Section "SESSION"
    $ctx = Get-ChaliceContext
    Write-Info "$($ctx.Computer)  |  running as $($ctx.RunAs)$(if ($ctx.Elevated) { ' (elevated)' })  |  console user $(if ($ctx.ConsoleUser) { $ctx.ConsoleUser } else { 'none' })"
    if ($ctx.Mismatch) { Add-ChaliceFinding -Code 'DifferentUser' -Detail "Running as $($ctx.RunAs); signed in at the console: $($ctx.ConsoleUser)" }

    Write-Section "OFFICE INSTALL"
    $install = Get-ChaliceInstall
    foreach ($m in $install.Msi) { Add-ChaliceFinding -Code 'MsiOffice' -Detail $m }
    if (-not $install.ClickToRun) {
        Write-Info "No Click-to-Run Office installed."
        Add-ChaliceFinding -Code 'NotClickToRun'
    } else {
        Write-Info "Version $($install.Version) ($($install.Platform))  |  $($install.Channel)"
        foreach ($p in $install.Products) {
            Write-Info ("{0,-28} {1}{2}" -f $p.Id, $p.Family, $(if ($p.EndDate) { "  -- support ends $($p.EndDate.ToString('yyyy-MM-dd'))" } else { '' }))
            if ($p.Support -eq 'Ended')      { Add-ChaliceFinding -Code 'OutOfSupport'      -Detail "$($p.Id) ($($p.Family)) -- ended $($p.EndDate.ToString('yyyy-MM-dd'))" }
            if ($p.Support -eq 'EndingSoon') { Add-ChaliceFinding -Code 'SupportEndingSoon' -Detail "$($p.Id) ($($p.Family)) -- ends $($p.EndDate.ToString('yyyy-MM-dd'))" }
        }
        if (Test-UpdatesDisabled -UpdatesEnabled $install.UpdatesEnabled -PolicyEnable $install.PolicyEnable) {
            Add-ChaliceFinding -Code 'UpdatesDisabled' -Detail $(if ("$($install.PolicyEnable)" -eq '0') { 'Policy: enableautomaticupdates = 0' } else { 'ClickToRun UpdatesEnabled = False' })
        }
        if ($install.UpdatePath -and $install.UpdatePath -match '^(\\\\|[A-Za-z]:\\)' -and -not (Test-Path -LiteralPath $install.UpdatePath)) {
            Add-ChaliceFinding -Code 'UpdateSourceUnreachable' -Detail $install.UpdatePath
        }
        if ($install.LastBinaryWrite) {
            $age = ((Get-Date) - $install.LastBinaryWrite).TotalDays
            Write-Info ("Office files last updated {0} ({1:N0} days ago)" -f $install.LastBinaryWrite.ToString('yyyy-MM-dd'), $age)
            if ($age -gt $StaleBuildDays) { Add-ChaliceFinding -Code 'StaleBuild' -Detail ("Last updated {0}, {1:N0} days ago" -f $install.LastBinaryWrite.ToString('yyyy-MM-dd'), $age) }
        }
        if ($install.SharedComputer) { Add-ChaliceFinding -Code 'SharedComputerActivation' }
    }

    Write-Section "ACTIVATION"
    $licenseFiles = @()
    if (Test-Path -LiteralPath $LicenseTokenDir) { $licenseFiles = @(Get-ChildItem -LiteralPath $LicenseTokenDir -Recurse -File -Force -ErrorAction SilentlyContinue) }
    $scaFiles = @()
    if (Test-Path -LiteralPath $ScaTokenDir) { $scaFiles = @(Get-ChildItem -LiteralPath $ScaTokenDir -Recurse -File -Force -ErrorAction SilentlyContinue) }
    $hasSubscription = @($install.Products | Where-Object { $_.Family -eq 'Microsoft 365 Apps' }).Count -gt 0
    Write-Info ("License tokens: {0} file(s){1}" -f $licenseFiles.Count, $(if ($install.SharedComputer) { "; shared-computer tokens: $($scaFiles.Count)" } else { '' }))
    if ($hasSubscription -and -not $install.SharedComputer -and $licenseFiles.Count -eq 0) { Add-ChaliceFinding -Code 'NoLicenseToken' -Detail "Nothing under $LicenseTokenDir" }
    if ($hasSubscription -and $install.SharedComputer -and $scaFiles.Count -eq 0) { Add-ChaliceFinding -Code 'NoLicenseToken' -Detail "Shared computer activation on; nothing under $ScaTokenDir" }

    Write-Section "SIGN-IN"
    $identity = Get-ChaliceIdentity
    $summary  = Get-IdentitySummary -Identities $identity.Identities
    Write-Info ("Office accounts: {0} work / school, {1} personal, {2} tenant(s)" -f $summary.Work, $summary.Personal, $summary.Tenants)
    foreach ($i in $identity.Identities) { Write-Info ("  {0}  [{1}]" -f $i.Email, $(if ($i.Provider -eq 'MSA') { 'personal' } elseif ($i.Provider -eq 'AD') { 'work' } else { $i.Provider })) }
    if ($summary.PersonalOnly -and $hasSubscription -and @($install.Products | Where-Object { $_.Id -match '(?i)Business|ProPlus|E3|E5' }).Count) {
        Add-ChaliceFinding -Code 'PersonalAccountOnly' -Detail (($identity.Identities | ForEach-Object { $_.Email }) -join ', ')
    }
    if ($summary.Mixed) { Add-ChaliceFinding -Code 'MultipleAccounts' -Detail (($identity.Identities | ForEach-Object { $_.Email }) -join ', ') }
    if ($identity.EnableADAL -and "$($identity.EnableADAL.Value)" -eq '0') { Add-ChaliceFinding -Code 'ModernAuthDisabled' -Detail "EnableADAL = 0 ($($identity.EnableADAL.Source))" }
    foreach ($w in @(@{ N = 'DisableAADWAM'; V = $identity.DisableAADWAM }, @{ N = 'DisableADALatopWAMOverride'; V = $identity.DisableWamOverride })) {
        if ($w.V -and "$($w.V.Value)" -eq '1') { Add-ChaliceFinding -Code 'WamDisabled' -Detail "$($w.N) = 1 ($($w.V.Source))" }
    }

    $broker = Get-ChaliceBroker
    Write-Info ("Token broker (AAD BrokerPlugin): {0}" -f $(if ($broker.Registered) { "registered $($broker.Version)" } else { 'NOT REGISTERED' }))
    if (-not $broker.Registered) { Add-ChaliceFinding -Code 'BrokerPluginMissing' -Detail $(if ($broker.CanRepair) { 'System copy present; ResetSignIn can re-register it.' } else { "System copy not found at $BrokerManifest" }) }

    $creds = @(ConvertFrom-CmdkeyList -Lines (Invoke-Native -File 'cmdkey.exe' -Arguments @('/list')))
    Write-Info "Cached Office credentials: $($creds.Count)"

    $dsreg = Get-DeviceJoinState -Status (ConvertFrom-DsregcmdStatus -Lines (Invoke-Native -File 'dsregcmd.exe' -Arguments @('/status')))
    Write-Info ("Device: {0}{1}  |  PRT: {2}" -f $dsreg.Label, $(if ($dsreg.Tenant) { " ($($dsreg.Tenant))" } else { '' }), $(if ($dsreg.HasPrt) { 'yes' } else { 'no' }))
    if ($dsreg.EntraJoined -and -not $dsreg.HasPrt) { Add-ChaliceFinding -Code 'NoPrimaryRefreshToken' -Detail "$($dsreg.Label); AzureAdPrt = NO" }
    elseif ($dsreg.EntraJoined -and $dsreg.PrtUpdated -and ((Get-Date).ToUniversalTime() - $dsreg.PrtUpdated).TotalDays -gt $PrtStaleDays) {
        Add-ChaliceFinding -Code 'StalePrimaryRefreshToken' -Detail ("Last renewed {0:yyyy-MM-dd HH:mm} UTC" -f $dsreg.PrtUpdated)
    }
    if (-not $dsreg.EntraJoined -and -not $dsreg.Registered) { Add-ChaliceFinding -Code 'NotDeviceRegistered' -Detail $dsreg.Label }

    Write-Section "TEAMS"
    $teams = Get-ChaliceTeamsClient
    Write-Info ("New Teams: {0}  |  classic Teams: {1}  |  cache {2}" -f $(if ($teams.NewInstalled) { $teams.NewVersion } else { 'not installed' }), $(if ($teams.ClassicInstalled) { 'INSTALLED' } else { 'no' }), (Format-Bytes ($teams.NewCacheBytes + $teams.ClassicCacheBytes)))
    if ($teams.ClassicInstalled) { Add-ChaliceFinding -Code 'ClassicTeams' -Detail $ClassicTeamsExe }
    if (($teams.NewCacheBytes + $teams.ClassicCacheBytes) -gt $TeamsCacheInfoBytes) { Add-ChaliceFinding -Code 'TeamsCacheLarge' -Detail (Format-Bytes ($teams.NewCacheBytes + $teams.ClassicCacheBytes)) }

    return [PSCustomObject]@{
        Context      = $ctx
        Install      = $install
        LicenseFiles = $licenseFiles.Count
        ScaFiles     = $scaFiles.Count
        Identity     = $identity
        Summary      = $summary
        Broker       = $broker
        Credentials  = $creds
        Device       = $dsreg
        Teams        = $teams
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REPAIR
# ─────────────────────────────────────────────────────────────────────────────

function Confirm-OfficeClosed {
    # Office holds its identity and license caches open. Returns $true when
    # no Office app is running (closing them after asking, interactively).
    param([string]$Purpose)
    $running = @(Get-OfficeRunningProcess)
    if ($running.Count -eq 0) { return $true }
    $names = ($running | ForEach-Object { $_.ProcessName } | Select-Object -Unique) -join ', '
    if ($WhatIf) { Write-Warn "[WhatIf] Would close: $names"; return $true }
    if ($Unattended) {
        Write-Warn "Office apps are open ($names); $Purpose needs them closed. Skipped in unattended mode."
        return $false
    }
    $ans = Read-Host "  Close $names now? Unsaved work in them will be lost. [y/N]"
    if ($ans -notmatch '^[Yy]') { return $false }
    $running | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    return (@(Get-OfficeRunningProcess).Count -eq 0)
}

function Export-RegistryBackup {
    param([string]$Key, [string]$Label)
    $native = $Key -replace '^HKCU:\\', 'HKCU\' -replace '^HKLM:\\', 'HKLM\'
    $out = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("CHALICE_{0}_{1}.reg" -f $Label, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $null = Invoke-Native -File 'reg.exe' -Arguments @('export', $native, $out, '/y')
    if (Test-Path -LiteralPath $out) { return $out }
    return ''
}

function Invoke-ChaliceResetSignIn {
    param([object]$Audit)
    Write-Section "RESET — OFFICE SIGN-IN AND ACTIVATION"
    if (-not (Confirm-OfficeClosed -Purpose 'resetting sign-in')) {
        Add-ChaliceAction -Step 'Reset Office sign-in' -Status 'Skipped' -Detail 'Office apps were open.'
        return
    }

    foreach ($k in @(@{ Key = "$IdentityKey\Identities"; Label = 'Identities' }, @{ Key = $LicensingKey; Label = 'Licensing' })) {
        if (-not (Test-Path $k.Key)) { continue }
        if ($WhatIf) {
            Write-Warn "[WhatIf] Would back up and remove $($k.Key)"
            Add-ChaliceAction -Step "Clear $($k.Label) cache" -Status 'WhatIf' -Detail $k.Key
            continue
        }
        $backup = Export-RegistryBackup -Key $k.Key -Label $k.Label
        try {
            Remove-Item -Path $k.Key -Recurse -Force -ErrorAction Stop
            Write-Ok "Cleared $($k.Key)$(if ($backup) { " (backup: $backup)" })"
            Add-ChaliceAction -Step "Clear $($k.Label) cache" -Status 'Done' -Detail $(if ($backup) { "Backup: $backup" } else { 'No backup written' })
        } catch {
            Write-Fail "Could not clear $($k.Key): $($_.Exception.Message)"
            Add-ChaliceAction -Step "Clear $($k.Label) cache" -Status 'Failed' -Detail $_.Exception.Message
        }
    }

    foreach ($dir in @($LicenseTokenDir, $ScaTokenDir)) {
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        if ($WhatIf) {
            Write-Warn "[WhatIf] Would delete the license tokens in $dir"
            Add-ChaliceAction -Step 'Delete license tokens' -Status 'WhatIf' -Detail $dir
            continue
        }
        try {
            Get-ChildItem -LiteralPath $dir -Force -ErrorAction Stop | Remove-Item -Recurse -Force -ErrorAction Stop
            Write-Ok "Deleted license tokens in $dir"
            Add-ChaliceAction -Step 'Delete license tokens' -Status 'Done' -Detail $dir
        } catch {
            Write-Fail "Could not delete tokens in ${dir}: $($_.Exception.Message)"
            Add-ChaliceAction -Step 'Delete license tokens' -Status 'Failed' -Detail $_.Exception.Message
        }
    }

    foreach ($target in $Audit.Credentials) {
        if ($WhatIf) {
            Write-Warn "[WhatIf] Would delete credential $target"
            Add-ChaliceAction -Step 'Delete cached Office credential' -Status 'WhatIf' -Detail $target
            continue
        }
        $out = Invoke-Native -File 'cmdkey.exe' -Arguments @("/delete:$target")
        $gone = -not ((ConvertFrom-CmdkeyList -Lines (Invoke-Native -File 'cmdkey.exe' -Arguments @('/list'))) -contains $target)
        Add-ChaliceAction -Step 'Delete cached Office credential' -Status $(if ($gone) { 'Done' } else { 'Failed' }) -Detail $(if ($gone) { $target } else { "$target -- $($out -join ' ')" })
    }

    if (-not $Audit.Broker.Registered -and $Audit.Broker.CanRepair) {
        if ($WhatIf) {
            Write-Warn "[WhatIf] Would re-register the AAD token broker from $BrokerManifest"
            Add-ChaliceAction -Step 'Re-register AAD token broker' -Status 'WhatIf'
        } else {
            try {
                # The Appx module is Windows PowerShell's; under PowerShell 7 hand it over.
                if ($PSVersionTable.PSEdition -eq 'Desktop') {
                    Add-AppxPackage -Register $BrokerManifest -DisableDevelopmentMode -ForceApplicationShutdown -ErrorAction Stop
                } else {
                    & powershell.exe -NoProfile -Command "Add-AppxPackage -Register '$BrokerManifest' -DisableDevelopmentMode -ForceApplicationShutdown -ErrorAction Stop"
                    if ($LASTEXITCODE -ne 0) { throw "powershell.exe exited $LASTEXITCODE" }
                }
                Write-Ok "Re-registered the AAD token broker."
                Add-ChaliceAction -Step 'Re-register AAD token broker' -Status 'Done' -Detail $BrokerManifest
            } catch {
                Write-Fail "Could not re-register the token broker: $($_.Exception.Message)"
                Add-ChaliceAction -Step 'Re-register AAD token broker' -Status 'Failed' -Detail $_.Exception.Message
            }
        }
    }

    if (-not $WhatIf) {
        Write-Info "Open Word or Outlook and sign in with the work account when asked. If Settings > Accounts > Access work or school lists a wrong account, remove it there too."
        Add-TKNote -Text "Reset Office sign-in and activation for $($env:USERNAME) (identities, license tokens, Office credentials cleared)." -Category 'Action' -ScriptName 'chalice'
    }
}

function Invoke-ChaliceTeamsCacheReset {
    Write-Section "RESET — TEAMS CACHE"
    $procs = @(Get-Process -Name 'ms-teams', 'Teams' -ErrorAction SilentlyContinue)
    $targets = @()
    if (Test-Path -LiteralPath $NewTeamsCacheDir) { $targets += $NewTeamsCacheDir }
    foreach ($f in $ClassicTeamsCacheFolders) {
        $p = Join-Path $ClassicTeamsDir $f
        if (Test-Path -LiteralPath $p) { $targets += $p }
    }
    if ($targets.Count -eq 0) {
        Write-Info "No Teams cache found."
        Add-ChaliceAction -Step 'Clear Teams cache' -Status 'Skipped' -Detail 'No cache folders present.'
        return
    }
    if ($WhatIf) {
        if ($procs.Count) { Write-Warn "[WhatIf] Would close Teams." }
        foreach ($t in $targets) { Write-Warn "[WhatIf] Would clear $t" }
        Add-ChaliceAction -Step 'Clear Teams cache' -Status 'WhatIf' -Detail ($targets -join '; ')
        return
    }
    if ($procs.Count) {
        if (-not $Unattended) {
            $ans = Read-Host "  Teams is running and will be closed. Continue? [y/N]"
            if ($ans -notmatch '^[Yy]') { Add-ChaliceAction -Step 'Clear Teams cache' -Status 'Skipped' -Detail 'Declined -- Teams was running.'; return }
        }
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
    $failed = @()
    foreach ($t in $targets) {
        try { Get-ChildItem -LiteralPath $t -Force -ErrorAction Stop | Remove-Item -Recurse -Force -ErrorAction Stop }
        catch { $failed += "${t}: $($_.Exception.Message)" }
    }
    if ($failed.Count -eq 0) {
        Write-Ok "Teams cache cleared. Start Teams again; the first launch rebuilds it and may ask the user to sign in."
        Add-ChaliceAction -Step 'Clear Teams cache' -Status 'Done' -Detail ($targets -join '; ')
        Add-TKNote -Text "Cleared the Teams cache for $($env:USERNAME)." -Category 'Action' -ScriptName 'chalice'
    } else {
        Write-Warn "Some cache files could not be removed (still in use?)."
        Add-ChaliceAction -Step 'Clear Teams cache' -Status 'Failed' -Detail ($failed -join '; ')
    }
}

function Invoke-ChaliceQuickRepair {
    param([object]$Audit)
    Write-Section "REPAIR — OFFICE QUICK REPAIR"
    $exe = Join-Path $C2RClientDir 'OfficeClickToRun.exe'
    if (-not $Audit.Install.ClickToRun -or -not (Test-Path -LiteralPath $exe)) {
        Write-Warn "Quick Repair needs Click-to-Run Office."
        Add-ChaliceAction -Step 'Office Quick Repair' -Status 'Skipped' -Detail 'Click-to-Run Office not found.'
        return
    }
    $platform = if ($Audit.Install.Platform) { $Audit.Install.Platform } else { 'x64' }
    $culture  = if ($Audit.Install.Culture) { $Audit.Install.Culture } else { 'en-us' }
    $argList = @('scenario=Repair', "platform=$platform", "culture=$culture", 'RepairType=QuickRepair', 'forceappshutdown=True', 'DisplayLevel=False')
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: $exe $($argList -join ' ')"
        Add-ChaliceAction -Step 'Office Quick Repair' -Status 'WhatIf' -Detail ($argList -join ' ')
        return
    }
    if (-not (Test-IsAdmin) -and $Unattended) {
        Write-Warn "Quick Repair needs elevation; skipped in unattended mode."
        Add-ChaliceAction -Step 'Office Quick Repair' -Status 'Skipped' -Detail 'Needs elevation; run interactively.'
        return
    }
    if (-not $Unattended) {
        $ans = Read-Host "  Quick Repair closes every Office app and takes a few minutes. Continue? [y/N]"
        if ($ans -notmatch '^[Yy]') { Add-ChaliceAction -Step 'Office Quick Repair' -Status 'Skipped' -Detail 'Declined by technician.'; return }
    }
    try {
        $sp = @{ FilePath = $exe; ArgumentList = ($argList -join ' '); Wait = $true; PassThru = $true; ErrorAction = 'Stop' }
        if (-not (Test-IsAdmin)) { $sp['Verb'] = 'RunAs' }
        $proc = Start-Process @sp
        $ok = ($proc.ExitCode -eq 0)
        if ($ok) { Write-Ok "Quick Repair finished." } else { Write-Warn "Quick Repair exited with $($proc.ExitCode)." }
        Add-ChaliceAction -Step 'Office Quick Repair' -Status $(if ($ok) { 'Done' } else { 'Failed' }) -Detail "Exit $($proc.ExitCode)"
        if ($ok) { Add-TKNote -Text 'Ran Office Quick Repair.' -Category 'Action' -ScriptName 'chalice' }
    } catch {
        Write-Fail "Quick Repair did not start: $($_.Exception.Message)"
        Add-ChaliceAction -Step 'Office Quick Repair' -Status 'Failed' -Detail $_.Exception.Message
    }
}

function Invoke-ChaliceUpdate {
    param([object]$Audit)
    Write-Section "UPDATE — CLICK-TO-RUN"
    $exe = Join-Path $C2RClientDir 'OfficeC2RClient.exe'
    if (-not $Audit.Install.ClickToRun -or -not (Test-Path -LiteralPath $exe)) {
        Write-Warn "Updating needs Click-to-Run Office."
        Add-ChaliceAction -Step 'Start Office update' -Status 'Skipped' -Detail 'Click-to-Run Office not found.'
        return
    }
    $argList = @('/update', 'user', 'updatepromptuser=false', 'forceappshutdown=false', 'displaylevel=true')
    if ($WhatIf) {
        Write-Warn "[WhatIf] Would run: $exe $($argList -join ' ')"
        Add-ChaliceAction -Step 'Start Office update' -Status 'WhatIf' -Detail ($argList -join ' ')
        return
    }
    try {
        Start-Process -FilePath $exe -ArgumentList ($argList -join ' ') -ErrorAction Stop
        Write-Ok "Office update started; its own window shows progress. Apps that are open update when they next close."
        Add-ChaliceAction -Step 'Start Office update' -Status 'Done' -Detail "From $($Audit.Install.Version) on $($Audit.Install.Channel)"
        Add-TKNote -Text "Started an Office update from $($Audit.Install.Version)." -Category 'Action' -ScriptName 'chalice'
    } catch {
        Write-Fail "Could not start the update: $($_.Exception.Message)"
        Add-ChaliceAction -Step 'Start Office update' -Status 'Failed' -Detail $_.Exception.Message
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-ChaliceReport {
    param([object]$Audit, [object]$Verdict, [string]$Mode)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $in         = $Audit.Install

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>No Microsoft 365 Apps issues found.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
            $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $pRows = [System.Text.StringBuilder]::new()
    if (@($in.Products).Count -eq 0) { [void]$pRows.Append("<tr><td colspan='3'>$(if ($in.ClickToRun) { 'No products listed.' } else { 'Click-to-Run Office is not installed.' })</td></tr>") }
    foreach ($p in $in.Products) {
        $badge = switch ($p.Support) { 'Ended' { "<span class='tk-badge-err'>Ended $($p.EndDate.ToString('yyyy-MM-dd'))</span>" } 'EndingSoon' { "<span class='tk-badge-warn'>Ends $($p.EndDate.ToString('yyyy-MM-dd'))</span>" } default { $(if ($p.EndDate) { "<span class='tk-badge-ok'>Until $($p.EndDate.ToString('yyyy-MM-dd'))</span>" } else { "<span class='tk-badge-ok'>Subscription</span>" }) } }
        [void]$pRows.Append("<tr><td class='tk-mono'>$(EscHtml $p.Id)</td><td>$(EscHtml $p.Family)</td><td>$badge</td></tr>")
    }
    foreach ($m in $in.Msi) { [void]$pRows.Append("<tr><td class='tk-mono'>$(EscHtml $m)</td><td>MSI Office</td><td><span class='tk-badge-err'>Ended</span></td></tr>") }

    $iRows = [System.Text.StringBuilder]::new()
    if ($Audit.Identity.Identities.Count -eq 0) { [void]$iRows.Append("<tr><td colspan='3'>Office has no cached accounts for this user.</td></tr>") }
    foreach ($i in $Audit.Identity.Identities) {
        $kind = switch ($i.Provider) { 'AD' { "<span class='tk-badge-blue'>Work / school</span>" } 'MSA' { "<span class='tk-badge-info'>Personal</span>" } default { "<span class='tk-badge-info'>$(EscHtml $i.Provider)</span>" } }
        [void]$iRows.Append("<tr><td>$(EscHtml $i.Email)</td><td>$kind</td><td class='tk-mono'>$(EscHtml $i.TenantId)</td></tr>")
    }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Actions.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='4'>Read-only audit — no changes were attempted.</td></tr>") }
    foreach ($a in $Actions) {
        $badge = switch ($a.Status) { 'Done' { 'ok' } 'WhatIf' { 'blue' } 'Failed' { 'err' } default { 'info' } }
        [void]$aRows.Append("<tr><td class='tk-mono'>$(EscHtml $a.Timestamp)</td><td>$(EscHtml $a.Step)</td><td><span class='tk-badge-$badge'>$(EscHtml $a.Status)</span></td><td>$(EscHtml $a.Detail)</td></tr>")
    }

    $policyText = {
        param($v)
        if (-not $v) { return 'not set' }
        return "$($v.Value) ($($v.Source))"
    }
    $dev = $Audit.Device
    $prtText = if ($dev.HasPrt) { "Yes$(if ($dev.PrtUpdated) { ', renewed ' + $dev.PrtUpdated.ToString('yyyy-MM-dd HH:mm') + ' UTC' })" } else { 'No' }
    $lastUpdate = if ($in.LastBinaryWrite) { $in.LastBinaryWrite.ToString('yyyy-MM-dd') } else { 'unknown' }
    $teams = $Audit.Teams
    $teamsCache = $teams.NewCacheBytes + $teams.ClassicCacheBytes

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.H.A.L.I.C.E. Microsoft 365 Apps Report' `
        -ScriptName 'C.H.A.L.I.C.E.' `
        -Subtitle   "${orgPrefix}Microsoft 365 Apps -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'User'      = $Audit.Context.RunAs
            'Generated' = $reportDate
            'Mode'      = $(if ($WhatIf) { "$Mode (dry run)" } else { $Mode })
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Install', 'Activation', 'Sign-in', 'Teams', 'Actions Taken')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Microsoft 365 Apps</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $(if ($in.Version) { $in.Version -replace '^16\.0\.', '' } else { 'n/a' }))</div><div class="tk-summary-lbl">Build</div></div>
    <div class="tk-summary-card $(if ($Audit.LicenseFiles -gt 0 -or $Audit.ScaFiles -gt 0) { 'ok' } else { 'warn' })"><div class="tk-summary-num">$($Audit.LicenseFiles + $Audit.ScaFiles)</div><div class="tk-summary-lbl">License Tokens</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Summary.Total)</div><div class="tk-summary-lbl">Office Accounts</div></div>
    <div class="tk-summary-card $(if ($dev.EntraJoined -and -not $dev.HasPrt) { 'err' } else { 'info' })"><div class="tk-summary-num">$(EscHtml $dev.Label)</div><div class="tk-summary-lbl">Device</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Install</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">Version</span> $(EscHtml $(if ($in.Version) { "$($in.Version) ($($in.Platform))" } else { 'not installed' }))<br/>
      <span class="tk-info-label">Channel</span> $(EscHtml $(if ($in.Channel) { $in.Channel } else { '-' }))<br/>
      <span class="tk-info-label">Updates</span> $(if (Test-UpdatesDisabled -UpdatesEnabled $in.UpdatesEnabled -PolicyEnable $in.PolicyEnable) { 'Disabled' } elseif ($in.ClickToRun) { 'Enabled' } else { '-' })$(if ($in.UpdatePath) { " from <span class='tk-mono'>$(EscHtml $in.UpdatePath)</span>" })<br/>
      <span class="tk-info-label">Files last updated</span> $(EscHtml $lastUpdate)<br/>
      <span class="tk-info-label">Install path</span> <span class="tk-mono">$(EscHtml $(if ($in.InstallPath) { $in.InstallPath } else { '-' }))</span>
    </div>
    <table class="tk-table"><thead><tr><th>Product</th><th>Family</th><th>Support</th></tr></thead>
      <tbody>$($pRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Activation</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">Shared computer activation</span> $(if ($in.SharedComputer) { 'On' } else { 'Off' })<br/>
      <span class="tk-info-label">License tokens</span> $($Audit.LicenseFiles) in <span class="tk-mono">$(EscHtml $LicenseTokenDir)</span>$(if ($in.SharedComputer) { "; $($Audit.ScaFiles) shared-computer token(s)" })
    </div></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Sign-in</div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">Accounts Office has cached</span></div><table class="tk-table">
      <thead><tr><th>Account</th><th>Kind</th><th>Tenant</th></tr></thead>
      <tbody>$($iRows.ToString())</tbody></table></div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">EnableADAL</span> $(EscHtml (& $policyText $Audit.Identity.EnableADAL))<br/>
      <span class="tk-info-label">DisableAADWAM</span> $(EscHtml (& $policyText $Audit.Identity.DisableAADWAM))<br/>
      <span class="tk-info-label">DisableADALatopWAMOverride</span> $(EscHtml (& $policyText $Audit.Identity.DisableWamOverride))<br/>
      <span class="tk-info-label">AAD token broker</span> $(if ($Audit.Broker.Registered) { "Registered $(EscHtml $Audit.Broker.Version)" } else { 'Not registered' })<br/>
      <span class="tk-info-label">Cached Office credentials</span> $(@($Audit.Credentials).Count)<br/>
      <span class="tk-info-label">Device</span> $(EscHtml $dev.Label)$(if ($dev.Tenant) { " -- $(EscHtml $dev.Tenant)" })<br/>
      <span class="tk-info-label">Primary Refresh Token</span> $(EscHtml $prtText)<br/>
      <span class="tk-info-label">Checked as</span> $(EscHtml $Audit.Context.RunAs)$(if ($Audit.Context.Mismatch) { " -- console user is $(EscHtml $Audit.Context.ConsoleUser)" })
    </div></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Teams</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">New Teams</span> $(if ($teams.NewInstalled) { EscHtml $teams.NewVersion } else { 'Not installed' })<br/>
      <span class="tk-info-label">Classic Teams</span> $(if ($teams.ClassicInstalled) { 'Installed (retired)' } else { 'Not installed' })<br/>
      <span class="tk-info-label">Cache size</span> $(EscHtml (Format-Bytes $teamsCache))
    </div></div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Actions Taken</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Time</th><th>Step</th><th>Status</th><th>Detail</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'C.H.A.L.I.C.E. v6.0')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ChaliceRun {
    param([string]$Mode)
    $Findings.Clear()
    $Actions.Clear()

    $audit = Invoke-ChaliceAudit
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Office install, activation, sign-in and Teams all in order." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    switch ($Mode) {
        'ResetSignIn' { Invoke-ChaliceResetSignIn -Audit $audit }
        'ResetTeams'  { Invoke-ChaliceTeamsCacheReset }
        'QuickRepair' { Invoke-ChaliceQuickRepair -Audit $audit }
        'Update'      { Invoke-ChaliceUpdate -Audit $audit }
    }

    $verdict = Get-ChaliceVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    Add-TKNote -Text ("CHALICE {0} for {1} on {2}: verdict {3} ({4} finding(s))." -f $Mode, $audit.Context.RunAs, $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'chalice'

    Write-Step "Generating HTML report..."
    $html    = Build-ChaliceReport -Audit $audit -Verdict $verdict -Mode $Mode
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("CHALICE_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'chalice' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-ChaliceBanner
    Invoke-ChaliceRun -Mode $Action
} else {
    $choice = ''
    do {
        Show-ChaliceBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Audit  -  read-only: install, activation, sign-in, Teams + HTML report" -ForegroundColor $C.Info
        Write-Host "  [2] Reset sign-in  -  clear Office accounts, license tokens and credentials (user signs in again)" -ForegroundColor $C.Info
        Write-Host "  [3] Reset Teams  -  close Teams and clear its cache" -ForegroundColor $C.Info
        Write-Host "  [4] Quick Repair  -  Office's offline repair (asks for elevation)" -ForegroundColor $C.Info
        Write-Host "  [5] Update  -  start a Click-to-Run update now" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        if ($WhatIf) { Write-Host "  Dry run is active — options 2-5 will preview only." -ForegroundColor $C.Warning }
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-ChaliceRun -Mode 'Audit' }
            '2' { Invoke-ChaliceRun -Mode 'ResetSignIn' }
            '3' { Invoke-ChaliceRun -Mode 'ResetTeams' }
            '4' { Invoke-ChaliceRun -Mode 'QuickRepair' }
            '5' { Invoke-ChaliceRun -Mode 'Update' }
            'Q' { Write-Host ""; Write-Host "  Closing C.H.A.L.I.C.E." -ForegroundColor $C.Header; Write-Host "" }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-5 or Q." -ForegroundColor $C.Warning
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
$null = Save-TKNoteReport -Prefix 'CHALICE' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
