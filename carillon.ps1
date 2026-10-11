# carillon.ps1 - C.A.R.I.L.L.O.N. — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers
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
    C.A.R.I.L.L.O.N. — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers
    Teams Phone Call Queue & Auto Attendant Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Inventories every Teams Phone call queue and auto attendant in the tenant
    and shows who is actually connected to them -- by display name, never by
    object ID:

      - Call queues: phone numbers, routing method, presence-based routing,
        conference mode, overflow / timeout / no-agent handling and where
        those calls go, and every agent with their opt-in state, whether
        they are voice-enabled, and how they got there (added directly,
        through a group, or through a Teams channel).
      - Auto attendants: phone numbers, language and time zone, operator,
        the business-hours menu (each key and where it transfers), and the
        after-hours and holiday call flows with their schedules.
      - Resource accounts: every account, its number, and the queue or
        attendant it fronts.
      - Routing: which attendants and queues hand calls to each queue or
        attendant, so an unreachable one stands out.

    Flags queues nobody can answer (no agents, or every agent opted out),
    agents who are not voice-enabled or are disabled, routing targets that
    point at deleted users or groups, queues and attendants with no number
    and nothing routing to them, and resource accounts assigned to nothing.

    Read-only -- nothing in the tenant is changed. Needs the MicrosoftTeams
    module for the queues and attendants, and Microsoft.Graph.Authentication
    (Directory.Read.All) to put names on groups and to show which group
    brought each agent in. Both are offered for install if missing.
    -SkipGraph avoids the second sign-in; users are still named through the
    Teams module, but group names are not resolved.

    Writes an HTML report and a CSV of every queue / agent pairing to the
    log directory.

.USAGE
    PS C:\> .\carillon.ps1                     # Interactive: sign in, audit, HTML + CSV
    PS C:\> .\carillon.ps1 -Name 'Sales*'      # Only queues / attendants whose name matches
    PS C:\> .\carillon.ps1 -SkipGraph          # Teams sign-in only; group names left unresolved
    PS C:\> .\carillon.ps1 -Unattended         # Silent: sign in, audit, export

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [string]$Name,
    [switch]$SkipGraph
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

# Script-scoped note buffer. Lives for the duration of one tool run; the sync
# script copies it only into tools that use the note functions.
$script:TKNotes = New-Object System.Collections.ArrayList

function Add-TKNote {
    <#
    .SYNOPSIS
        Records a timestamped technician note for the current session.
    .DESCRIPTION
        Notes accumulate in script-scoped state for the life of one tool run.
        Export them to a ticket-ready HTML report with Export-TKNoteReport
        before the tool exits.
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

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

$GraphScopes = @('Directory.Read.All')

# The application IDs a resource account carries say what it can front. They
# are the same in every tenant.
$ResourceAccountApps = @{
    '11cd3e2e-fccb-42ad-ad00-878b93575e07' = 'Call queue'
    'ce933385-9390-45d1-9512-c8d228074e07' = 'Auto attendant'
}

# Queue actions that end the call rather than send it somewhere.
$DisconnectActions = @('Disconnect', 'DisconnectWithBusy')

# Readable text for the routing enums the Teams module returns.
$RoutingMethodNames = @{
    'Attendant'   = 'Attendant (ring all)'
    'Serial'      = 'Serial'
    'RoundRobin'  = 'Round robin'
    'LongestIdle' = 'Longest idle'
}
$MenuActionNames = @{
    'DisconnectCall'         = 'Disconnect'
    'TransferCallToOperator' = 'Transfer to operator'
    'TransferCallToTarget'   = 'Transfer'
    'Announcement'           = 'Play announcement'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition C.A.R.I.L.L.O.N. can report, keyed by a stable code. The
# Pester suite extracts it by AST lookup and checks it against the codes the
# script raises.
# ─────────────────────────────────────────────────────────────────────────────

$CarillonFindings = @{
    'QueueNoAgents' = @{
        Severity = 'Error'
        Title    = 'Call queue has no agents'
        Summary  = 'Nobody is connected to this queue, so every call waits out the timeout or goes straight to the no-agent action.'
        Remedy   = 'Add users, a group, or a Teams channel as the queue''s agents in the Teams admin center (Voice > Call queues > Agents).'
    }
    'QueueAllOptedOut' = @{
        Severity = 'Error'
        Title    = 'Every agent has opted out of the queue'
        Summary  = 'The queue has agents, but none is taking calls. Calls are not offered to opted-out agents.'
        Remedy   = 'Have at least one agent opt back in (Teams > Settings > Calls > Call queues), or turn off "Agents can opt out" on the queue.'
    }
    'QueueFewAgentsOptedIn' = @{
        Severity = 'Warning'
        Title    = 'Only one agent is taking calls'
        Summary  = 'A single opted-in agent means the queue stops being answered the moment that person is away.'
        Remedy   = 'Confirm the coverage is intended, or add or opt in a second agent.'
    }
    'AgentNotVoiceEnabled' = @{
        Severity = 'Warning'
        Title    = 'Agent is not voice-enabled'
        Summary  = 'The agent has no Enterprise Voice, which usually means no Teams Phone licence. Queue calls from the phone network may not reach them.'
        Remedy   = 'Assign a Teams Phone licence to the agent, or remove them from the queue.'
    }
    'AgentDisabled' = @{
        Severity = 'Warning'
        Title    = 'Agent account is disabled'
        Summary  = 'A disabled account is still listed as an agent but will never answer.'
        Remedy   = 'Remove the account from the queue, or from the group that adds it.'
    }
    'StaleReference' = @{
        Severity = 'Error'
        Title    = 'Routing points at a deleted user or group'
        Summary  = 'An agent list, overflow / timeout target, menu option or operator names an object that no longer exists. Calls sent there fail.'
        Remedy   = 'Edit the queue or attendant and replace the missing target.'
    }
    'OverflowImmediate' = @{
        Severity = 'Warning'
        Title    = 'Queue overflows on the first call'
        Summary  = 'The overflow threshold is 0, so every call takes the overflow action instead of ringing an agent.'
        Remedy   = 'Raise the overflow threshold unless the queue is meant to be bypassed.'
    }
    'CallsDisconnected' = @{
        Severity = 'Info'
        Title    = 'Queue hangs up on overflow or timeout'
        Summary  = 'Callers who hit the limit are disconnected rather than sent to voicemail or another target.'
        Remedy   = 'Consider shared voicemail or a forward so those calls are not lost.'
    }
    'Unreachable' = @{
        Severity = 'Warning'
        Title    = 'Nothing routes calls here'
        Summary  = 'No resource account with a phone number fronts it, and no auto attendant or queue transfers to it.'
        Remedy   = 'Give its resource account a number, or add it as a target in an attendant menu or queue overflow -- or delete it if it is no longer used.'
    }
    'NoAfterHours' = @{
        Severity = 'Info'
        Title    = 'Auto attendant has no after-hours handling'
        Summary  = 'Every call gets the business-hours flow, day and night.'
        Remedy   = 'Fine for a 24/7 line. Otherwise add business hours and an after-hours call flow.'
    }
    'ResourceAccountUnassigned' = @{
        Severity = 'Info'
        Title    = 'Resource account not assigned to anything'
        Summary  = 'The account fronts no call queue or auto attendant. Any number or licence on it is doing nothing.'
        Remedy   = 'Assign it to the queue or attendant it was made for, or delete it and reclaim the number and licence.'
    }
    'GroupsNotResolved' = @{
        Severity = 'Info'
        Title    = 'Group names were not resolved'
        Summary  = 'Microsoft Graph was skipped or unavailable, so groups appear by object ID and agents are not traced to the group that added them.'
        Remedy   = 'Run again without -SkipGraph and consent to Directory.Read.All.'
    }
    'QueryFailed' = @{
        Severity = 'Warning'
        Title    = 'Part of the tenant could not be read'
        Summary  = 'A Teams or Microsoft Graph query failed, so the results below may be incomplete.'
        Remedy   = 'Sign in with Teams Administrator (or Global Reader plus Teams Communications Support Engineer) and retry.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-CarillonFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Subject = '',
        [string]$Detail = ''
    )
    $meta = $CarillonFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary
        Remedy = $meta.Remedy; Subject = $Subject; Detail = $Detail
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

function Show-CarillonBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗ █████╗ ██████╗ ██╗██╗     ██╗      ██████╗ ███╗   ██╗
  ██╔════╝██╔══██╗██╔══██╗██║██║     ██║     ██╔═══██╗████╗  ██║
  ██║     ███████║██████╔╝██║██║     ██║     ██║   ██║██╔██╗ ██║
  ██║     ██╔══██║██╔══██╗██║██║     ██║     ██║   ██║██║╚██╗██║
  ╚██████╗██║  ██║██║  ██║██║███████╗███████╗╚██████╔╝██║ ╚████║
   ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝╚══════╝╚══════╝ ╚═════╝ ╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    C.A.R.I.L.L.O.N. — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers" -ForegroundColor Cyan
    Write-Host "    Teams Phone Call Queue & Auto Attendant Audit Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
#
# Queue and attendant objects come from the Teams module as typed objects; the
# tests pass PSCustomObjects or hashtables. Member access reads all of them the
# same way. No I/O until the COLLECTORS section.
# ─────────────────────────────────────────────────────────────────────────────

function Get-CarillonList {
    # The Teams module returns missing collections as $null and single values
    # unwrapped; GUIDs arrive as [guid]. Normalise to a flat string array.
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
}

function Get-CarillonName {
    # Display label for a directory object ID -- never the bare GUID when the
    # object resolved. $Names maps ID -> @{ DisplayName; Upn }.
    param([string]$Id, [hashtable]$Names)
    if ($Names -and $Id -and $Names.ContainsKey($Id.ToLowerInvariant())) {
        $n = $Names[$Id.ToLowerInvariant()]
        if ($n.Upn -and $n.Upn -ne $n.DisplayName) { return "$($n.DisplayName) <$($n.Upn)>" }
        return "$($n.DisplayName)"
    }
    return "[unresolved $Id]"
}

function Format-CarillonNumber {
    param([string]$Uri)
    return ("$Uri" -replace '^(?i)tel:', '').Trim()
}

function Format-CarillonTarget {
    # Renders a callable entity -- a queue overflow / timeout target, an
    # attendant menu option or operator -- as readable text. $Endpoints maps a
    # resource-account ID to what it fronts; $Configs maps a queue or attendant
    # identity to its label.
    param($Target, [hashtable]$Names, [hashtable]$Endpoints = @{}, [hashtable]$Configs = @{})
    if ($null -eq $Target -or -not "$($Target.Id)") { return '' }
    $id  = "$($Target.Id)".Trim()
    $key = $id.ToLowerInvariant()
    switch ("$($Target.Type)") {
        'ExternalPstn'    { return "External number $(Format-CarillonNumber $id)" }
        'User'            { return "User $(Get-CarillonName -Id $id -Names $Names)" }
        'SharedVoicemail' { return "Shared voicemail $(Get-CarillonName -Id $id -Names $Names)" }
        'ApplicationEndpoint' {
            if ($Endpoints.ContainsKey($key)) { return "$($Endpoints[$key]) (via $(Get-CarillonName -Id $id -Names $Names))" }
            return "Resource account $(Get-CarillonName -Id $id -Names $Names)"
        }
        'ConfigurationEndpoint' {
            if ($Configs.ContainsKey($key)) { return "$($Configs[$key])" }
            return "[unresolved queue / attendant $id]"
        }
        default { return "$($Target.Type) $(Get-CarillonName -Id $id -Names $Names)" }
    }
}

function Format-CarillonAction {
    # "Forward -> User Jane Doe" / "Disconnect". Queue actions name a target
    # only when they send the call somewhere.
    param([string]$Action, [string]$TargetText)
    if (-not $Action) { return '' }
    if ($TargetText -and $Action -notin $DisconnectActions) { return "$Action -> $TargetText" }
    return $Action
}

function Get-CarillonQueueIssue {
    # The finding codes a queue earns from its own settings and agent list.
    # $Queue carries AgentCount, OptedInCount, OverflowThreshold,
    # OverflowAction and TimeoutAction.
    param($Queue)
    $codes = [System.Collections.Generic.List[string]]::new()
    if ([int]$Queue.AgentCount -eq 0) {
        [void]$codes.Add('QueueNoAgents')
    } elseif ([int]$Queue.OptedInCount -eq 0) {
        [void]$codes.Add('QueueAllOptedOut')
    } elseif ([int]$Queue.OptedInCount -eq 1) {
        [void]$codes.Add('QueueFewAgentsOptedIn')
    }
    if ($null -ne $Queue.OverflowThreshold -and "$($Queue.OverflowThreshold)" -ne '' -and
        [int]$Queue.OverflowThreshold -eq 0 -and "$($Queue.OverflowAction)" -ne 'Queue') {
        [void]$codes.Add('OverflowImmediate')
    }
    if ("$($Queue.OverflowAction)" -in $DisconnectActions -or "$($Queue.TimeoutAction)" -in $DisconnectActions) {
        [void]$codes.Add('CallsDisconnected')
    }
    return $codes.ToArray()
}

function Get-CarillonAgentSource {
    # How an agent came to be on the queue. $GroupMembers maps a group ID to
    # the set of user IDs in it (empty when groups could not be expanded).
    param([string]$AgentId, [string[]]$DirectUsers, [string[]]$Groups, [hashtable]$GroupMembers, [hashtable]$Names, [bool]$Channel)
    $src = [System.Collections.Generic.List[string]]::new()
    if (@($DirectUsers) -contains $AgentId) { [void]$src.Add('Direct') }
    foreach ($g in @($Groups)) {
        $key = "$g".ToLowerInvariant()
        if ($GroupMembers -and $GroupMembers.ContainsKey($key) -and $GroupMembers[$key].Contains($AgentId.ToLowerInvariant())) {
            [void]$src.Add("Group: $(Get-CarillonName -Id $g -Names $Names)")
        }
    }
    if ($src.Count -eq 0) {
        if ($Channel)                   { [void]$src.Add('Teams channel') }
        elseif (@($Groups).Count -gt 0) { [void]$src.Add('Group') }
        else                            { [void]$src.Add('Direct') }
    }
    return ($src -join '; ')
}

function Test-CarillonReachable {
    # A queue or attendant is reachable when one of its resource accounts has a
    # number, or another attendant / queue routes to it.
    param([string[]]$ResourceAccountIds, [hashtable]$NumberedAccounts, [string[]]$ReachedFrom)
    if (@($ReachedFrom | Where-Object { $_ }).Count -gt 0) { return $true }
    foreach ($id in @($ResourceAccountIds)) {
        if ($NumberedAccounts -and $NumberedAccounts.ContainsKey("$id".ToLowerInvariant())) { return $true }
    }
    return $false
}

function Get-CarillonVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Broken';    Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Attention'; Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULES + CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

function Install-CarillonModule {
    param([string]$ModuleName)
    if (Get-Module -ListAvailable -Name $ModuleName) {
        Write-Ok "$ModuleName — installed"
    } else {
        Write-Warn "$ModuleName — NOT found"
        $doInstall = $Unattended
        if (-not $Unattended) {
            $ans = Read-Host "  Install $ModuleName for current user? [Y/N]"
            $doInstall = $ans -match '^[Yy]'
        }
        if (-not $doInstall) { Write-Fail "$ModuleName was not installed."; return $false }
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Ok "$ModuleName installed."
        } catch {
            Write-Fail "Install failed: $($_.Exception.Message)"
            Write-TKError -ScriptName 'carillon' -Message "$ModuleName install failed: $($_.Exception.Message)" -Category 'Module Install'
            return $false
        }
    }
    try { Import-Module $ModuleName -ErrorAction Stop; return $true }
    catch { Write-Fail "Could not import ${ModuleName}: $($_.Exception.Message)"; return $false }
}

function Connect-CarillonTeams {
    Write-Section "CONNECT TO MICROSOFT TEAMS"
    try {
        $tenant = Get-CsTenant -ErrorAction Stop
        if ($tenant) {
            Write-Ok "Reusing existing Teams session: $($tenant.DisplayName)"
            return [PSCustomObject]@{ Account = ''; Tenant = "$($tenant.DisplayName)" }
        }
    } catch {
        Write-Verbose "No cached Teams session: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in..."
    try {
        $conn = Connect-MicrosoftTeams -ErrorAction Stop
        $tenantName = ''
        try { $tenantName = "$((Get-CsTenant -ErrorAction Stop).DisplayName)" } catch { Write-Verbose "Get-CsTenant failed: $($_.Exception.Message)" }
        Write-Ok "Connected as: $($conn.Account)"
        return [PSCustomObject]@{ Account = "$($conn.Account)"; Tenant = $tenantName }
    } catch {
        Write-Fail "Connect-MicrosoftTeams failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "Connect-MicrosoftTeams failed: $($_.Exception.Message)" -Category 'Teams Auth'
        return $null
    }
}

function Connect-CarillonGraph {
    Write-Section "CONNECT TO MICROSOFT GRAPH"
    if (-not (Install-CarillonModule -ModuleName 'Microsoft.Graph.Authentication')) { return $false }
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account -and -not @($GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })) {
            Write-Ok "Reusing existing session: $($ctx.Account)"
            return $true
        }
    } catch {
        Write-Verbose "No cached Graph context: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in (used only to name groups)..."
    Write-Info "Scopes: $($GraphScopes -join ', ')"
    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        Write-Ok "Connected as: $((Get-MgContext).Account)"
        return $true
    } catch {
        Write-Warn "Connect-MgGraph failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-CarillonPaged {
    # Get-CsCallQueue and Get-CsAutoAttendant return at most 100 objects per
    # call; page with -First / -Skip until a short page comes back.
    param([string]$Command, [string]$Label)
    $items = [System.Collections.Generic.List[object]]::new()
    $skip  = 0
    $page  = 100
    try {
        do {
            $batch = @(& $Command -First $page -Skip $skip -ErrorAction Stop -WarningAction SilentlyContinue)
            foreach ($b in $batch) { [void]$items.Add($b) }
            $skip += $page
        } while ($batch.Count -eq $page)
    } catch {
        Add-CarillonFinding -Code 'QueryFailed' -Subject $Label -Detail $_.Exception.Message
        Write-Warn "$Label query failed: $($_.Exception.Message)"
        return $null
    }
    return , $items.ToArray()
}

function Resolve-CarillonGraphName {
    # directoryObjects/getByIds names users and groups in one call per 1000
    # IDs. Returns the set of IDs Graph was asked about, so the caller can tell
    # "does not exist" from "never asked".
    param([string[]]$Ids, [hashtable]$Names, [hashtable]$Kinds)
    $asked = @{}
    $ids   = @($Ids | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique)
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = @($ids[$i..([math]::Min($i + 999, $ids.Count - 1))])
        try {
            $body = @{ ids = $chunk; types = @('user', 'group') } | ConvertTo-Json -Depth 3
            $resp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/directoryObjects/getByIds' -Body $body -ContentType 'application/json' -ErrorAction Stop
            foreach ($o in @($resp['value'])) {
                $key = "$($o['id'])".ToLowerInvariant()
                $Names[$key] = [PSCustomObject]@{ DisplayName = "$($o['displayName'])"; Upn = "$($o['userPrincipalName'])" }
                $Kinds[$key] = if ("$($o['@odata.type'])" -match 'group$') { 'Group' } else { 'User' }
            }
            foreach ($c in $chunk) { $asked[$c] = $true }
        } catch {
            Add-CarillonFinding -Code 'QueryFailed' -Subject 'Graph directoryObjects/getByIds' -Detail $_.Exception.Message
            Write-Warn "Graph name lookup failed: $($_.Exception.Message)"
            return $asked
        }
    }
    return $asked
}

function Get-CarillonGroupMember {
    # Transitive user members of a group, as a lower-case ID set; $null when
    # the query fails.
    param([string]$GroupId)
    $set  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $next = "v1.0/groups/$GroupId/transitiveMembers/microsoft.graph.user" + '?$select=id&$top=999'
    try {
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
            foreach ($m in @($resp['value'])) { [void]$set.Add("$($m['id'])") }
            $next = $resp['@odata.nextLink']
        }
    } catch {
        Write-Verbose "Group $GroupId expansion failed: $($_.Exception.Message)"
        return $null
    }
    return , $set
}

function Get-CarillonUser {
    # Get-CsOnlineUser for one ID. Status is Found, Missing (the object does
    # not exist) or Failed (the lookup itself broke -- not evidence of anything).
    param([string]$Id)
    try {
        $u = Get-CsOnlineUser -Identity $Id -ErrorAction Stop -WarningAction SilentlyContinue
        return [PSCustomObject]@{
            Status        = 'Found'
            DisplayName   = "$($u.DisplayName)"
            Upn           = "$($u.UserPrincipalName)"
            VoiceEnabled  = [bool]$u.EnterpriseVoiceEnabled
            AccountEnabled = if ($null -eq $u.AccountEnabled) { $true } else { [bool]$u.AccountEnabled }
        }
    } catch {
        $status = if ("$($_.Exception.Message)" -match 'not found|NotFound|does not exist|couldn''t find|could not find') { 'Missing' } else { 'Failed' }
        return [PSCustomObject]@{ Status = $status; DisplayName = ''; Upn = ''; VoiceEnabled = $null; AccountEnabled = $null }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Get-CarillonMenuTarget {
    # Every callable entity an attendant can send a call to: its operator and
    # each menu option of each call flow.
    param($Attendant)
    $targets = [System.Collections.Generic.List[object]]::new()
    if ($Attendant.Operator -and "$($Attendant.Operator.Id)") { [void]$targets.Add($Attendant.Operator) }
    foreach ($flow in @(@($Attendant.DefaultCallFlow) + @($Attendant.CallFlows) | Where-Object { $_ })) {
        foreach ($opt in @($flow.Menu.MenuOptions | Where-Object { $_ })) {
            if ($opt.CallTarget -and "$($opt.CallTarget.Id)") { [void]$targets.Add($opt.CallTarget) }
        }
    }
    return , $targets.ToArray()
}

function Get-CarillonQueueTarget {
    param($Queue)
    return , @(@($Queue.OverflowActionTarget, $Queue.TimeoutActionTarget, $Queue.NoAgentActionTarget) |
        Where-Object { $_ -and "$($_.Id)" })
}

function Format-CarillonMenu {
    # One line per menu option: "1: Transfer -> Call queue Sales".
    param($Flow, [hashtable]$Names, [hashtable]$Endpoints, [hashtable]$Configs)
    if (-not $Flow -or -not $Flow.Menu) { return @() }
    $lines = foreach ($opt in @($Flow.Menu.MenuOptions | Where-Object { $_ })) {
        $key = ("$($opt.DtmfResponse)" -replace '^Tone', '') -replace '^Star$', '*' -replace '^Pound$', '#'
        if ($key -eq 'Automatic') { $key = 'Auto' }
        $action = if ($MenuActionNames.ContainsKey("$($opt.Action)")) { $MenuActionNames["$($opt.Action)"] } else { "$($opt.Action)" }
        $target = Format-CarillonTarget -Target $opt.CallTarget -Names $Names -Endpoints $Endpoints -Configs $Configs
        if ($target -and "$($opt.Action)" -eq 'TransferCallToTarget') { "${key}: $action -> $target" } else { "${key}: $action" }
    }
    $lines = @($lines | Where-Object { $_ })
    if ($Flow.Menu.DialByNameEnabled) { $lines += 'Dial by name enabled' }
    return $lines
}

function Invoke-CarillonAudit {
    param([bool]$UseGraph)

    Write-Section "CALL QUEUES, AUTO ATTENDANTS AND RESOURCE ACCOUNTS"
    Write-Step "Reading call queues..."
    $queues = Get-CarillonPaged -Command 'Get-CsCallQueue' -Label 'Call queues'
    $queues = @($queues | Where-Object { $_ })
    Write-Step "Reading auto attendants..."
    $attendants = Get-CarillonPaged -Command 'Get-CsAutoAttendant' -Label 'Auto attendants'
    $attendants = @($attendants | Where-Object { $_ })
    Write-Step "Reading resource accounts..."
    $accounts = @()
    try {
        $accounts = @(Get-CsOnlineApplicationInstance -ErrorAction Stop -WarningAction SilentlyContinue | Where-Object { $_ })
    } catch {
        Add-CarillonFinding -Code 'QueryFailed' -Subject 'Resource accounts' -Detail $_.Exception.Message
        Write-Warn "Resource account query failed: $($_.Exception.Message)"
    }
    Write-Info ("Call queues: {0}  |  Auto attendants: {1}  |  Resource accounts: {2}" -f $queues.Count, $attendants.Count, $accounts.Count)

    # ── Names ──
    # Resource accounts come back named; everything else is resolved below.
    $names = @{}
    $kinds = @{}
    $numbered = @{}
    foreach ($ra in $accounts) {
        $key = "$($ra.ObjectId)".ToLowerInvariant()
        $names[$key] = [PSCustomObject]@{ DisplayName = "$($ra.DisplayName)"; Upn = "$($ra.UserPrincipalName)" }
        $kinds[$key] = 'ResourceAccount'
        if ("$($ra.PhoneNumber)") { $numbered[$key] = Format-CarillonNumber "$($ra.PhoneNumber)" }
    }

    $agentIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $objectIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($q in $queues) {
        foreach ($a in @($q.Agents | Where-Object { $_ })) { [void]$agentIds.Add("$($a.ObjectId)") }
        foreach ($u in (Get-CarillonList $q.Users)) { [void]$agentIds.Add($u) }
        foreach ($g in (Get-CarillonList $q.DistributionLists)) { [void]$groupIds.Add($g) }
        foreach ($t in (Get-CarillonQueueTarget $q)) {
            if ("$($t.Type)" -in @('User', 'SharedVoicemail', 'ApplicationEndpoint')) { [void]$objectIds.Add("$($t.Id)") }
        }
    }
    foreach ($aa in $attendants) {
        foreach ($t in (Get-CarillonMenuTarget $aa)) {
            if ("$($t.Type)" -in @('User', 'SharedVoicemail', 'ApplicationEndpoint')) { [void]$objectIds.Add("$($t.Id)") }
        }
    }
    foreach ($id in $agentIds) { [void]$objectIds.Add($id) }
    foreach ($id in $groupIds) { [void]$objectIds.Add($id) }

    # $missing holds IDs confirmed not to exist; an ID that simply could not be
    # looked up is never counted as deleted.
    $missing = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groupMembers = @{}
    if ($UseGraph -and $objectIds.Count -gt 0) {
        Write-Step "Resolving names through Microsoft Graph..."
        $asked = Resolve-CarillonGraphName -Ids @($objectIds) -Names $names -Kinds $kinds
        foreach ($id in $objectIds) {
            $key = $id.ToLowerInvariant()
            if ($asked.ContainsKey($key) -and -not $names.ContainsKey($key)) { [void]$missing.Add($key) }
        }
        foreach ($g in $groupIds) {
            $members = Get-CarillonGroupMember -GroupId $g
            if ($null -ne $members) { $groupMembers[$g.ToLowerInvariant()] = $members }
        }
    }

    # Every agent through Get-CsOnlineUser: names without Graph, and the voice
    # and account state Graph does not report.
    $agentInfo = @{}
    $i = 0
    foreach ($id in $agentIds) {
        $i++
        Write-Progress -Activity 'Reading agents' -Status "$i of $($agentIds.Count)" -PercentComplete ([math]::Min(100, 100 * $i / [math]::Max(1, $agentIds.Count)))
        $info = Get-CarillonUser -Id $id
        $key  = $id.ToLowerInvariant()
        $agentInfo[$key] = $info
        if ($info.Status -eq 'Found') {
            if (-not $names.ContainsKey($key)) { $names[$key] = [PSCustomObject]@{ DisplayName = $info.DisplayName; Upn = $info.Upn } }
            [void]$missing.Remove($key)
        } elseif ($info.Status -eq 'Missing' -and -not $names.ContainsKey($key)) {
            [void]$missing.Add($key)
        }
    }
    Write-Progress -Activity 'Reading agents' -Completed

    # Without Graph, the non-agent user targets still get a name.
    if (-not $UseGraph) {
        foreach ($id in $objectIds) {
            $key = $id.ToLowerInvariant()
            if ($names.ContainsKey($key) -or $agentInfo.ContainsKey($key) -or $groupIds.Contains($id)) { continue }
            $info = Get-CarillonUser -Id $id
            if ($info.Status -eq 'Found')       { $names[$key] = [PSCustomObject]@{ DisplayName = $info.DisplayName; Upn = $info.Upn } }
            elseif ($info.Status -eq 'Missing') { [void]$missing.Add($key) }
        }
    }

    $unresolvedGroups = @($groupIds | Where-Object { -not $names.ContainsKey($_.ToLowerInvariant()) -and -not $missing.Contains($_) })
    if ($unresolvedGroups.Count -gt 0) {
        Add-CarillonFinding -Code 'GroupsNotResolved' -Subject "$($unresolvedGroups.Count) group(s)" -Detail ($unresolvedGroups -join ', ')
    }

    # ── Who fronts what ──
    $endpoints = @{}   # resource account ID -> "Call queue Sales"
    $configs   = @{}   # queue / attendant identity -> "Call queue Sales"
    $raOwner   = @{}   # resource account ID -> list of labels
    foreach ($q in $queues) {
        $label = "Call queue $($q.Name)"
        $configs["$($q.Identity)".ToLowerInvariant()] = $label
        foreach ($ra in (Get-CarillonList $q.ApplicationInstances)) {
            $endpoints[$ra.ToLowerInvariant()] = $label
            if (-not $raOwner.ContainsKey($ra.ToLowerInvariant())) { $raOwner[$ra.ToLowerInvariant()] = [System.Collections.Generic.List[string]]::new() }
            [void]$raOwner[$ra.ToLowerInvariant()].Add($label)
        }
    }
    foreach ($aa in $attendants) {
        $label = "Auto attendant $($aa.Name)"
        $configs["$($aa.Identity)".ToLowerInvariant()] = $label
        foreach ($ra in (Get-CarillonList $aa.ApplicationInstances)) {
            $endpoints[$ra.ToLowerInvariant()] = $label
            if (-not $raOwner.ContainsKey($ra.ToLowerInvariant())) { $raOwner[$ra.ToLowerInvariant()] = [System.Collections.Generic.List[string]]::new() }
            [void]$raOwner[$ra.ToLowerInvariant()].Add($label)
        }
    }

    # ── Who routes to what ──
    $reachedFrom = @{}   # queue / attendant identity -> list of source labels
    function _link { param([string]$ToIdentity, [string]$From)
        $k = $ToIdentity.ToLowerInvariant()
        if (-not $reachedFrom.ContainsKey($k)) { $reachedFrom[$k] = [System.Collections.Generic.List[string]]::new() }
        if (-not $reachedFrom[$k].Contains($From)) { [void]$reachedFrom[$k].Add($From) }
    }
    $sources = @(@($queues | ForEach-Object { [PSCustomObject]@{ Label = "Call queue $($_.Name)"; Targets = (Get-CarillonQueueTarget $_) } }) +
                 @($attendants | ForEach-Object { [PSCustomObject]@{ Label = "Auto attendant $($_.Name)"; Targets = (Get-CarillonMenuTarget $_) } }))
    foreach ($s in $sources) {
        foreach ($t in @($s.Targets)) {
            $tid = "$($t.Id)".ToLowerInvariant()
            if ("$($t.Type)" -eq 'ConfigurationEndpoint') { _link $tid $s.Label; continue }
            if ("$($t.Type)" -ne 'ApplicationEndpoint') { continue }
            foreach ($q in $queues)     { if ((Get-CarillonList $q.ApplicationInstances)  -contains $tid) { _link "$($q.Identity)"  $s.Label } }
            foreach ($aa in $attendants) { if ((Get-CarillonList $aa.ApplicationInstances) -contains $tid) { _link "$($aa.Identity)" $s.Label } }
        }
    }
    function _from { param([string]$Identity) $k = $Identity.ToLowerInvariant(); if ($reachedFrom.ContainsKey($k)) { @($reachedFrom[$k]) } else { @() } }
    function _stale { param([object[]]$Targets)
        @($Targets | Where-Object { "$($_.Type)" -in @('User', 'SharedVoicemail', 'ApplicationEndpoint') -and $missing.Contains("$($_.Id)") } |
            ForEach-Object { "$($_.Type) $($_.Id)" })
    }

    $inScope = { param([string]$ItemName) (-not $Name) -or ($ItemName -like $Name) }

    # ── Call queues ──
    $queueRows = [System.Collections.Generic.List[object]]::new()
    $agentRows = [System.Collections.Generic.List[object]]::new()
    foreach ($q in ($queues | Sort-Object Name)) {
        if (-not (& $inScope "$($q.Name)")) { continue }
        $qName   = "$($q.Name)"
        $direct  = @(Get-CarillonList $q.Users | ForEach-Object { $_.ToLowerInvariant() })
        $groups  = @(Get-CarillonList $q.DistributionLists)
        $channel = [bool]"$($q.ChannelId)"

        # Agents is the effective list (direct users plus expanded groups and
        # channel members) with opt-in state. Older module builds lack it.
        $agentList = @($q.Agents | Where-Object { $_ } | ForEach-Object { [PSCustomObject]@{ Id = "$($_.ObjectId)"; OptIn = $_.OptIn } })
        if ($agentList.Count -eq 0 -and $direct.Count -gt 0) {
            $agentList = @($direct | ForEach-Object { [PSCustomObject]@{ Id = $_; OptIn = $null } })
        }

        $rows = foreach ($a in $agentList) {
            $key  = $a.Id.ToLowerInvariant()
            $info = $agentInfo[$key]
            [PSCustomObject]@{
                Queue          = $qName
                Agent          = if ($names.ContainsKey($key)) { $names[$key].DisplayName } elseif ($missing.Contains($key)) { "[deleted user $($a.Id)]" } else { "[unresolved $($a.Id)]" }
                Upn            = if ($names.ContainsKey($key)) { $names[$key].Upn } else { '' }
                OptedIn        = if ($null -eq $a.OptIn) { 'Unknown' } elseif ($a.OptIn) { 'Yes' } else { 'No' }
                Source         = Get-CarillonAgentSource -AgentId $key -DirectUsers $direct -Groups $groups -GroupMembers $groupMembers -Names $names -Channel $channel
                VoiceEnabled   = if ($info -and $info.Status -eq 'Found') { if ($info.VoiceEnabled) { 'Yes' } else { 'No' } } else { 'Unknown' }
                AccountEnabled = if ($info -and $info.Status -eq 'Found') { if ($info.AccountEnabled) { 'Yes' } else { 'No' } } else { 'Unknown' }
                Missing        = $missing.Contains($key)
                ObjectId       = $a.Id
            }
        }
        $rows = @($rows | Sort-Object Missing, Agent)
        foreach ($r in $rows) { [void]$agentRows.Add($r) }

        $optedIn = @($rows | Where-Object { $_.OptedIn -ne 'No' -and -not $_.Missing }).Count
        $raIds   = @(Get-CarillonList $q.ApplicationInstances)
        $numbers = @($raIds | Where-Object { $numbered.ContainsKey($_.ToLowerInvariant()) } | ForEach-Object { $numbered[$_.ToLowerInvariant()] })
        $from    = @(_from "$($q.Identity)")

        $row = [PSCustomObject]@{
            Name               = $qName
            Numbers            = $numbers -join ', '
            ResourceAccounts   = (@($raIds | ForEach-Object { Get-CarillonName -Id $_ -Names $names }) -join ', ')
            Routing            = if ($RoutingMethodNames.ContainsKey("$($q.RoutingMethod)")) { $RoutingMethodNames["$($q.RoutingMethod)"] } else { "$($q.RoutingMethod)" }
            PresenceBased      = [bool]$q.PresenceBasedRouting
            ConferenceMode     = [bool]$q.ConferenceMode
            AllowOptOut        = [bool]$q.AllowOptOut
            AlertTime          = "$($q.AgentAlertTime)"
            AgentSource        = @(@(if ($direct.Count) { "$($direct.Count) direct" }) + @($groups | ForEach-Object { "group $(Get-CarillonName -Id $_ -Names $names)" }) + @(if ($channel) { 'Teams channel' })) -join ', '
            AgentCount         = $rows.Count
            OptedInCount       = $optedIn
            OverflowThreshold  = $q.OverflowThreshold
            OverflowAction     = "$($q.OverflowAction)"
            Overflow           = Format-CarillonAction -Action "$($q.OverflowAction)" -TargetText (Format-CarillonTarget -Target $q.OverflowActionTarget -Names $names -Endpoints $endpoints -Configs $configs)
            TimeoutThreshold   = "$($q.TimeoutThreshold)"
            TimeoutAction      = "$($q.TimeoutAction)"
            Timeout            = Format-CarillonAction -Action "$($q.TimeoutAction)" -TargetText (Format-CarillonTarget -Target $q.TimeoutActionTarget -Names $names -Endpoints $endpoints -Configs $configs)
            NoAgent            = Format-CarillonAction -Action "$($q.NoAgentAction)" -TargetText (Format-CarillonTarget -Target $q.NoAgentActionTarget -Names $names -Endpoints $endpoints -Configs $configs)
            ReachedFrom        = $from -join ', '
            Agents             = $rows
        }
        [void]$queueRows.Add($row)

        foreach ($code in (Get-CarillonQueueIssue $row)) {
            switch ($code) {
                'QueueNoAgents'         { Add-CarillonFinding -Code 'QueueNoAgents' -Subject $qName }
                'QueueAllOptedOut'      { Add-CarillonFinding -Code 'QueueAllOptedOut' -Subject $qName -Detail "$($rows.Count) agent(s), all opted out" }
                'QueueFewAgentsOptedIn' { Add-CarillonFinding -Code 'QueueFewAgentsOptedIn' -Subject $qName -Detail (@($rows | Where-Object { $_.OptedIn -ne 'No' -and -not $_.Missing } | ForEach-Object { $_.Agent }) -join ', ') }
                'OverflowImmediate'     { Add-CarillonFinding -Code 'OverflowImmediate' -Subject $qName -Detail $row.Overflow }
                'CallsDisconnected'     { Add-CarillonFinding -Code 'CallsDisconnected' -Subject $qName -Detail ("Overflow: {0}; Timeout: {1}" -f $row.Overflow, $row.Timeout) }
            }
        }
        $noVoice = @($rows | Where-Object { $_.VoiceEnabled -eq 'No' })
        if ($noVoice.Count) { Add-CarillonFinding -Code 'AgentNotVoiceEnabled' -Subject $qName -Detail (@($noVoice | ForEach-Object { $_.Agent }) -join ', ') }
        $disabled = @($rows | Where-Object { $_.AccountEnabled -eq 'No' })
        if ($disabled.Count) { Add-CarillonFinding -Code 'AgentDisabled' -Subject $qName -Detail (@($disabled | ForEach-Object { $_.Agent }) -join ', ') }
        $stale = @(@($rows | Where-Object { $_.Missing } | ForEach-Object { "agent $($_.ObjectId)" }) +
                   @($groups | Where-Object { $missing.Contains($_) } | ForEach-Object { "group $_" }) +
                   @(_stale (Get-CarillonQueueTarget $q)))
        if ($stale.Count) { Add-CarillonFinding -Code 'StaleReference' -Subject "Call queue $qName" -Detail ($stale -join ', ') }
        if (-not (Test-CarillonReachable -ResourceAccountIds $raIds -NumberedAccounts $numbered -ReachedFrom $from)) {
            Add-CarillonFinding -Code 'Unreachable' -Subject "Call queue $qName" -Detail $(if ($raIds.Count) { 'Resource account has no phone number' } else { 'No resource account assigned' })
        }
    }

    # ── Auto attendants ──
    $attendantRows = [System.Collections.Generic.List[object]]::new()
    foreach ($aa in ($attendants | Sort-Object Name)) {
        if (-not (& $inScope "$($aa.Name)")) { continue }
        $aName   = "$($aa.Name)"
        $raIds   = @(Get-CarillonList $aa.ApplicationInstances)
        $numbers = @($raIds | Where-Object { $numbered.ContainsKey($_.ToLowerInvariant()) } | ForEach-Object { $numbered[$_.ToLowerInvariant()] })
        $from    = @(_from "$($aa.Identity)")

        # CallHandlingAssociations tie each non-default call flow to a schedule.
        $schedules = @{}
        foreach ($s in @($aa.Schedules | Where-Object { $_ })) { $schedules["$($s.Id)"] = "$($s.Name)" }
        $flows = @{}
        foreach ($f in @($aa.CallFlows | Where-Object { $_ })) { $flows["$($f.Id)"] = $f }
        $otherFlows = foreach ($assoc in @($aa.CallHandlingAssociations | Where-Object { $_ })) {
            $flow  = $flows["$($assoc.CallFlowId)"]
            $sched = if ($schedules.ContainsKey("$($assoc.ScheduleId)")) { $schedules["$($assoc.ScheduleId)"] } else { "$($assoc.ScheduleId)" }
            [PSCustomObject]@{
                Type     = "$($assoc.Type)"
                Schedule = $sched
                Enabled  = if ($null -eq $assoc.Enabled) { $true } else { [bool]$assoc.Enabled }
                Menu     = @(Format-CarillonMenu -Flow $flow -Names $names -Endpoints $endpoints -Configs $configs)
            }
        }
        $otherFlows = @($otherFlows | Where-Object { $_ })

        $row = [PSCustomObject]@{
            Name             = $aName
            Numbers          = $numbers -join ', '
            ResourceAccounts = (@($raIds | ForEach-Object { Get-CarillonName -Id $_ -Names $names }) -join ', ')
            Language         = "$($aa.LanguageId)"
            TimeZone         = "$($aa.TimeZoneId)"
            Operator         = Format-CarillonTarget -Target $aa.Operator -Names $names -Endpoints $endpoints -Configs $configs
            VoiceInput       = [bool]$aa.VoiceResponseEnabled
            BusinessHours    = @(Format-CarillonMenu -Flow $aa.DefaultCallFlow -Names $names -Endpoints $endpoints -Configs $configs)
            OtherFlows       = $otherFlows
            ReachedFrom      = $from -join ', '
        }
        [void]$attendantRows.Add($row)

        if (@($otherFlows | Where-Object { $_.Type -eq 'AfterHours' }).Count -eq 0) { Add-CarillonFinding -Code 'NoAfterHours' -Subject $aName }
        $stale = @(_stale (Get-CarillonMenuTarget $aa))
        if ($stale.Count) { Add-CarillonFinding -Code 'StaleReference' -Subject "Auto attendant $aName" -Detail ($stale -join ', ') }
        if (-not (Test-CarillonReachable -ResourceAccountIds $raIds -NumberedAccounts $numbered -ReachedFrom $from)) {
            Add-CarillonFinding -Code 'Unreachable' -Subject "Auto attendant $aName" -Detail $(if ($raIds.Count) { 'Resource account has no phone number' } else { 'No resource account assigned' })
        }
    }

    # ── Resource accounts ──
    $accountRows = foreach ($ra in ($accounts | Sort-Object DisplayName)) {
        $key   = "$($ra.ObjectId)".ToLowerInvariant()
        $labels = if ($raOwner.ContainsKey($key)) { @($raOwner[$key]) } else { @() }
        if ($Name) {
            # Filtered run: only the accounts fronting a queue or attendant in scope.
            if (@($labels | Where-Object { ($_ -replace '^(Call queue|Auto attendant) ', '') -like $Name }).Count -eq 0) { continue }
        } elseif ($labels.Count -eq 0) {
            Add-CarillonFinding -Code 'ResourceAccountUnassigned' -Subject "$($ra.DisplayName)" -Detail $(if ("$($ra.PhoneNumber)") { "Holds $(Format-CarillonNumber "$($ra.PhoneNumber)")" } else { '' })
        }
        [PSCustomObject]@{
            Name       = "$($ra.DisplayName)"
            Upn        = "$($ra.UserPrincipalName)"
            Number     = Format-CarillonNumber "$($ra.PhoneNumber)"
            Type       = if ($ResourceAccountApps.ContainsKey("$($ra.ApplicationId)".ToLowerInvariant())) { $ResourceAccountApps["$($ra.ApplicationId)".ToLowerInvariant()] } else { 'Other' }
            AssignedTo = $labels -join ', '
        }
    }

    return [PSCustomObject]@{
        Queues           = $queueRows.ToArray()
        Agents           = $agentRows.ToArray()
        Attendants       = $attendantRows.ToArray()
        Accounts         = @($accountRows | Where-Object { $_ })
        QueueTotal       = $queues.Count
        AttendantTotal   = $attendants.Count
        AccountTotal     = $accounts.Count
        GraphUsed        = $UseGraph
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE + HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Show-CarillonInventory {
    param([object]$Audit)

    Write-Section "CALL QUEUES"
    if ($Audit.Queues.Count -eq 0) { Write-Info "No call queues$(if ($Name) { " matching '$Name'" })." }
    foreach ($q in $Audit.Queues) {
        Write-Host ""
        Write-Host ("  {0}" -f $q.Name) -ForegroundColor $ColorSchema.Header -NoNewline
        Write-Host ("   {0}  |  {1}  |  {2} of {3} agent(s) opted in" -f $(if ($q.Numbers) { $q.Numbers } else { 'no number' }), $q.Routing, $q.OptedInCount, $q.AgentCount) -ForegroundColor $ColorSchema.Info
        if ($q.AgentCount -eq 0) { Write-Fail "No agents"; continue }
        foreach ($a in $q.Agents) {
            $color = if ($a.Missing -or $a.AccountEnabled -eq 'No') { $ColorSchema.Error } elseif ($a.OptedIn -eq 'No' -or $a.VoiceEnabled -eq 'No') { $ColorSchema.Warning } else { $ColorSchema.Success }
            $state = if ($a.OptedIn -eq 'No') { 'opted out' } elseif ($a.OptedIn -eq 'Yes') { 'opted in' } else { 'opt-in unknown' }
            Write-Host ("      {0,-34} {1,-40} {2,-15} {3}" -f $a.Agent, $a.Upn, $state, $a.Source) -ForegroundColor $color
        }
    }

    Write-Section "AUTO ATTENDANTS"
    if ($Audit.Attendants.Count -eq 0) { Write-Info "No auto attendants$(if ($Name) { " matching '$Name'" })." }
    foreach ($aa in $Audit.Attendants) {
        Write-Host ""
        Write-Host ("  {0}" -f $aa.Name) -ForegroundColor $ColorSchema.Header -NoNewline
        Write-Host ("   {0}  |  {1}" -f $(if ($aa.Numbers) { $aa.Numbers } else { 'no number' }), $aa.TimeZone) -ForegroundColor $ColorSchema.Info
        if ($aa.Operator) { Write-Host "      Operator: $($aa.Operator)" -ForegroundColor $ColorSchema.Info }
        foreach ($line in $aa.BusinessHours) { Write-Host "      $line" -ForegroundColor $ColorSchema.Info }
        foreach ($f in $aa.OtherFlows) {
            Write-Host ("      [{0}: {1}]" -f $f.Type, $f.Schedule) -ForegroundColor $ColorSchema.Accent
            foreach ($line in $f.Menu) { Write-Host "        $line" -ForegroundColor $ColorSchema.Info }
        }
    }
}

function Show-CarillonFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Nothing to report."; return }
    $order = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        $line = "$($f.Title)" + $(if ($f.Subject) { " -- $($f.Subject)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }
}

function Build-CarillonReport {
    param([object]$Audit, [object]$Verdict, [object]$Connection)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant     = if ($Connection.Tenant) { $Connection.Tenant } elseif ($Connection.Account -match '@(.+)$') { $Matches[1] } else { 'Tenant' }
    $order      = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    function _yn { param([bool]$Value) if ($Value) { 'Yes' } else { 'No' } }
    function _cell { param([string]$Text) if ($Text) { EscHtml $Text } else { '-' } }

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>Nothing to report.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)</td>" +
            "<td>$(EscHtml $f.Subject)" + $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $qRows = [System.Text.StringBuilder]::new()
    if ($Audit.Queues.Count -eq 0) { [void]$qRows.Append("<tr><td colspan='7'>No call queues.</td></tr>") }
    foreach ($q in $Audit.Queues) {
        $cls = if ($q.AgentCount -eq 0 -or $q.OptedInCount -eq 0) { 'err' } elseif ($q.OptedInCount -eq 1) { 'warn' } else { 'ok' }
        $opts = @(@(if ($q.PresenceBased) { 'presence-based' }) + @(if ($q.ConferenceMode) { 'conference mode' }) + @(if ($q.AllowOptOut) { 'opt-out allowed' })) -join ', '
        [void]$qRows.Append(
            "<tr><td><strong>$(EscHtml $q.Name)</strong><br/><span class='tk-mono'>$(_cell $q.ResourceAccounts)</span></td>" +
            "<td class='tk-mono'>$(_cell $q.Numbers)</td>" +
            "<td>$(EscHtml $q.Routing)$(if ($opts) { "<br/><span class='tk-mono'>$(EscHtml $opts)</span>" })<br/><span class='tk-mono'>ring $(EscHtml $q.AlertTime) s</span></td>" +
            "<td><span class='tk-badge-$cls'>$($q.OptedInCount) / $($q.AgentCount)</span><br/><span class='tk-mono'>$(_cell $q.AgentSource)</span></td>" +
            "<td>after $(EscHtml "$($q.OverflowThreshold)") call(s): $(_cell $q.Overflow)</td>" +
            "<td>after $(EscHtml $q.TimeoutThreshold) s: $(_cell $q.Timeout)$(if ($q.NoAgent) { "<br/><span class='tk-mono'>no agents: $(EscHtml $q.NoAgent)</span>" })</td>" +
            "<td>$(_cell $q.ReachedFrom)</td></tr>")
    }

    $agentCards = [System.Text.StringBuilder]::new()
    if ($Audit.Queues.Count -eq 0) { [void]$agentCards.Append("<div class='tk-card'>No call queues.</div>") }
    foreach ($q in $Audit.Queues) {
        [void]$agentCards.Append("<div class='tk-card'><div class='tk-card-header'><span class='tk-card-label'>$(EscHtml $q.Name)</span><span class='tk-mono'>$($q.OptedInCount) of $($q.AgentCount) opted in</span></div>")
        if ($q.AgentCount -eq 0) {
            [void]$agentCards.Append("<div class='tk-info-box'><span class='tk-info-label'>No agents</span> Nobody is connected to this queue.</div></div>")
            continue
        }
        [void]$agentCards.Append("<table class='tk-table'><thead><tr><th>Agent</th><th>Sign-in name</th><th>Opted in</th><th>Voice enabled</th><th>Account</th><th>Added through</th></tr></thead><tbody>")
        foreach ($a in $q.Agents) {
            $opt   = switch ($a.OptedIn) { 'Yes' { "<span class='tk-badge-ok'>Yes</span>" } 'No' { "<span class='tk-badge-warn'>No</span>" } default { "<span class='tk-badge-info'>Unknown</span>" } }
            $voice = switch ($a.VoiceEnabled) { 'Yes' { "<span class='tk-badge-ok'>Yes</span>" } 'No' { "<span class='tk-badge-warn'>No</span>" } default { "<span class='tk-badge-info'>Unknown</span>" } }
            $acct  = if ($a.Missing) { "<span class='tk-badge-err'>Deleted</span>" } else { switch ($a.AccountEnabled) { 'Yes' { "<span class='tk-badge-ok'>Enabled</span>" } 'No' { "<span class='tk-badge-err'>Disabled</span>" } default { "<span class='tk-badge-info'>Unknown</span>" } } }
            [void]$agentCards.Append("<tr><td><strong>$(EscHtml $a.Agent)</strong></td><td class='tk-mono'>$(_cell $a.Upn)</td><td>$opt</td><td>$voice</td><td>$acct</td><td>$(EscHtml $a.Source)</td></tr>")
        }
        [void]$agentCards.Append("</tbody></table></div>")
    }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Audit.Attendants.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='5'>No auto attendants.</td></tr>") }
    foreach ($aa in $Audit.Attendants) {
        $bh = if ($aa.BusinessHours.Count) { (@($aa.BusinessHours | ForEach-Object { EscHtml $_ }) -join '<br/>') } else { '-' }
        $other = if ($aa.OtherFlows.Count) {
            (@($aa.OtherFlows | ForEach-Object {
                "<strong>$(EscHtml $_.Type)</strong> <span class='tk-mono'>($(EscHtml $_.Schedule)$(if (-not $_.Enabled) { ', disabled' }))</span><br/>" +
                (@($_.Menu | ForEach-Object { EscHtml $_ }) -join '<br/>')
            }) -join '<br/><br/>')
        } else { "<span class='tk-badge-info'>None -- business hours 24/7</span>" }
        [void]$aRows.Append(
            "<tr><td><strong>$(EscHtml $aa.Name)</strong><br/><span class='tk-mono'>$(_cell $aa.ResourceAccounts)</span><br/><span class='tk-mono'>$(EscHtml $aa.Language) | $(EscHtml $aa.TimeZone)$(if ($aa.VoiceInput) { ' | voice input' })</span></td>" +
            "<td class='tk-mono'>$(_cell $aa.Numbers)</td>" +
            "<td>$bh$(if ($aa.Operator) { "<br/><span class='tk-mono'>Operator: $(EscHtml $aa.Operator)</span>" })</td>" +
            "<td>$other</td><td>$(_cell $aa.ReachedFrom)</td></tr>")
    }

    $rRows = [System.Text.StringBuilder]::new()
    if (@($Audit.Accounts).Count -eq 0) { [void]$rRows.Append("<tr><td colspan='5'>No resource accounts.</td></tr>") }
    foreach ($r in $Audit.Accounts) {
        $assigned = if ($r.AssignedTo) { EscHtml $r.AssignedTo } else { "<span class='tk-badge-info'>Unassigned</span>" }
        [void]$rRows.Append("<tr><td><strong>$(EscHtml $r.Name)</strong></td><td class='tk-mono'>$(_cell $r.Upn)</td><td class='tk-mono'>$(_cell $r.Number)</td><td>$(EscHtml $r.Type)</td><td>$assigned</td></tr>")
    }

    $errCount   = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount  = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count
    $agentTotal = @($Audit.Agents | Select-Object -ExpandProperty ObjectId -Unique).Count

    $meta = [ordered]@{
        'Tenant'    = $tenant
        'Generated' = $reportDate
        'Verdict'   = $Verdict.Verdict
    }
    if ($Connection.Account) { $meta['Signed in as'] = $Connection.Account }
    if ($Name) { $meta['Name filter'] = $Name }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.A.R.I.L.L.O.N. Call Queue & Auto Attendant Report' `
        -ScriptName 'C.A.R.I.L.L.O.N.' `
        -Subtitle   "${orgPrefix}Teams Phone Call Queues & Auto Attendants -- $tenant" `
        -MetaItems  $meta `
        -NavItems   @('Findings', 'Call Queues', 'Queue Agents', 'Auto Attendants', 'Resource Accounts')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Queues.Count)</div><div class="tk-summary-lbl">Call Queues</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$agentTotal</div><div class="tk-summary-lbl">Distinct Agents</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Attendants.Count)</div><div class="tk-summary-lbl">Auto Attendants</div></div>
    <div class="tk-summary-card $(if ($errCount) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errCount</div><div class="tk-summary-lbl">Errors</div></div>
    <div class="tk-summary-card $(if ($warnCount) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Call Queues</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Queue</th><th>Number</th><th>Routing</th><th>Agents (opted in / total)</th><th>Overflow</th><th>Timeout</th><th>Reached from</th></tr></thead>
      <tbody>$($qRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Queue Agents</div>
    $($agentCards.ToString())
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Auto Attendants</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Auto attendant</th><th>Number</th><th>Business hours</th><th>After hours / holidays</th><th>Reached from</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Resource Accounts</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Name</th><th>Sign-in name</th><th>Number</th><th>Type</th><th>Assigned to</th></tr></thead>
      <tbody>$($rRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'C.A.R.I.L.L.O.N. v6.0')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-CarillonRun {
    $Findings.Clear()

    Write-Section "MODULE CHECK"
    if (-not (Install-CarillonModule -ModuleName 'MicrosoftTeams')) { if ($Unattended) { exit 1 }; return }
    $conn = Connect-CarillonTeams
    if (-not $conn) { if ($Unattended) { exit 1 }; return }

    $useGraph = $false
    if (-not $SkipGraph) {
        $useGraph = Connect-CarillonGraph
        if (-not $useGraph) { Write-Warn "Continuing without Graph -- users are named through Teams, groups are not." }
    }

    $audit = Invoke-CarillonAudit -UseGraph $useGraph
    Show-CarillonInventory -Audit $audit
    Show-CarillonFindings

    $verdict = Get-CarillonVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }

    Add-TKNote -Text ("CARILLON Teams Phone audit: verdict {0}; {1} call queue(s), {2} auto attendant(s); {3} finding(s)." -f $verdict.Verdict, $audit.Queues.Count, $audit.Attendants.Count, $Findings.Count) -Category 'Info' -ScriptName 'carillon'

    $logDir = Resolve-LogDirectory -FallbackPath $ScriptPath
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'

    Write-Step "Writing queue agent CSV..."
    try {
        $csvPath = Join-Path $logDir "CARILLON_Agents_$stamp.csv"
        @($audit.Agents | Select-Object Queue, Agent, @{ n = 'UserPrincipalName'; e = { $_.Upn } }, OptedIn, VoiceEnabled, AccountEnabled, Source, ObjectId) |
            Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-Ok "Agent CSV saved: $csvPath"
    } catch {
        Write-Fail "Could not save CSV: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "CSV save failed: $($_.Exception.Message)" -Category 'Report'
    }

    Write-Step "Generating HTML report..."
    $html    = Build-CarillonReport -Audit $audit -Verdict $verdict -Connection $conn
    $outPath = Join-Path $logDir "CARILLON_$stamp.html"
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

Show-CarillonBanner
Invoke-CarillonRun

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
