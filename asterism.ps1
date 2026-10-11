# asterism.ps1 - A.S.T.E.R.I.S.M. — Audits Settings of Teams: Exposure, Roles, Inactivity, Sprawl & Membership
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
    A.S.T.E.R.I.S.M. — Audits Settings of Teams: Exposure, Roles, Inactivity, Sprawl & Membership
    Microsoft Teams Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and audits the M365 Teams estate:
    team inventory with visibility and size, orphan teams (no owner or
    all owners disabled), public teams (join without approval), guest
    membership, large teams, and stale teams with no recent activity.
    Produces a dark-themed HTML report that complements A.L.M.A.N.A.C.
    (licensing), O.R.B.I.T. (devices), and E.C.L.I.P.S.E. (identity) by
    covering the collaboration layer.

.USAGE
    PS C:\> .\asterism.ps1                    # Interactive menu
    PS C:\> .\asterism.ps1 -Unattended        # Silent: auto-connect + export HTML

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
#endregion TK-SHARED

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

$script:Connected    = $false
$script:ConnectedAs  = ''
$script:TenantDomain = ''

$GraphScopes = @(
    'Team.ReadBasic.All',
    'TeamMember.Read.All',
    'Group.Read.All',
    'User.Read.All',
    'Directory.Read.All'
)

function Show-AsterismBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  A.S.T.E.R.I.S.M. — Audits Settings of Teams: Exposure, Roles, Inactivity, Sprawl & Membership" -ForegroundColor Cyan
    Write-Host "  Microsoft Teams Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Connect, audits, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH MODULE INSTALL + CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $required = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Teams',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    $needsInstall = $false
    foreach ($m in $required) {
        if (Get-Module -ListAvailable -Name $m) { Write-Ok "$m — installed" }
        else { Write-Warn "$m — NOT found"; $needsInstall = $true }
    }

    if ($needsInstall) {
        Write-Host ""
        if ($Unattended) {
            Write-Step "Installing Microsoft.Graph (unattended)..."
            try {
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Ok "Microsoft.Graph installed."
            } catch {
                Write-Fail "Install failed: $_"
                Write-TKError -ScriptName 'asterism' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
                exit 1
            }
        } else {
            $ans = Read-Host "  Install Microsoft.Graph for current user? [Y/N]"
            if ($ans -match '^[Yy]') {
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Ok "Microsoft.Graph installed."
                } catch { Write-Fail "Install failed: $_"; exit 1 }
            } else { Write-Fail "Microsoft.Graph required. Exiting."; exit 1 }
        }
    } else {
        Write-Ok "All required Graph sub-modules are available."
    }

    foreach ($m in $required) {
        try { Import-Module $m -ErrorAction Stop } catch { Write-Warn "Could not import ${m}: $_" }
    }
}

function Test-GraphConnection {
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account) {
            $script:Connected   = $true
            $script:ConnectedAs = $ctx.Account
            try {
                $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($org -and $org.VerifiedDomains) {
                    $primary = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
                    if ($primary) { $script:TenantDomain = $primary.Name }
                }
            } catch {
                # Organization lookup is best-effort metadata for the report header.
            }
            return $true
        }
    } catch {
        # No cached context — caller will invoke interactive auth.
    }
    $script:Connected   = $false
    $script:ConnectedAs = ''
    return $false
}

function Invoke-Connect {
    Write-Section "CONNECT TO MICROSOFT 365"
    Write-Step "Requesting interactive sign-in..."
    Write-Info "Scopes: $($GraphScopes -join ', ')"
    Write-Host ""

    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        if (Test-GraphConnection) {
            Write-Ok "Connected as: $($script:ConnectedAs)"
            if ($script:TenantDomain) { Write-Ok "Tenant: $($script:TenantDomain)" }
        } else {
            Write-Fail "Authentication returned no context — may have been cancelled."
        }
    } catch {
        Write-Fail "Authentication failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'asterism' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — TEAMS INVENTORY (enriched)
# ─────────────────────────────────────────────────────────────────────────────

# Enumerate teams via the M365 groups that have resourceProvisioningOptions=Team,
# then for each team pull owner/member counts and the archive/visibility state.
# This is chattier than Get-MgTeam but gives us the owner list we need for the
# orphan detector in section 2.
function Get-TeamsInventory {
    Write-Section "TEAMS INVENTORY"
    Write-Step "Enumerating teams (this can take a while on large tenants)..."

    try {
        $groups = Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
            -Property "id,displayName,description,visibility,mailNickname,createdDateTime,renewedDateTime,assignedLabels" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to list team-backed groups: $_"
        Write-TKError -ScriptName 'asterism' -Message "Get-MgGroup (team filter) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    if (-not $groups -or $groups.Count -eq 0) {
        Write-Warn "No teams found in this tenant."
        return @()
    }

    Write-Ok "$($groups.Count) team(s) found. Enriching owner/member counts..."

    $total   = $groups.Count
    $idx     = 0
    $records = foreach ($g in $groups) {
        $idx++
        if ($idx -eq 1 -or ($idx % 25) -eq 0 -or $idx -eq $total) {
            $pct = [math]::Round(($idx / $total) * 100, 0)
            Write-Host ("`r  [*] Enriching team {0}/{1} ({2}%)..." -f $idx, $total, $pct) -ForegroundColor $C.Progress -NoNewline
        }

        $ownerList  = @()
        $memberList = @()
        try { $ownerList  = @(Get-MgGroupOwner  -GroupId $g.Id -All -ErrorAction SilentlyContinue) } catch {
            # Owner read failures on a single team don't justify aborting the audit.
        }
        try { $memberList = @(Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue) } catch {
            # Same for member read failures.
        }

        $guestCount = 0
        foreach ($m in $memberList) {
            $type = $m.AdditionalProperties['@odata.type']
            $userType = $m.AdditionalProperties['userType']
            if ($type -match 'user' -and $userType -eq 'Guest') { $guestCount++ }
        }

        $enabledOwners = 0
        foreach ($o in $ownerList) {
            $enabled = $o.AdditionalProperties['accountEnabled']
            if ($enabled -ne $false) { $enabledOwners++ }
        }

        [PSCustomObject]@{
            Id             = $g.Id
            DisplayName    = $g.DisplayName
            Description    = $g.Description
            Visibility     = $g.Visibility
            Created        = if ($g.CreatedDateTime) { ([datetime]$g.CreatedDateTime).ToString('yyyy-MM-dd') } else { '' }
            LastRenewed    = if ($g.RenewedDateTime) { ([datetime]$g.RenewedDateTime).ToString('yyyy-MM-dd') } else { '' }
            OwnerCount     = $ownerList.Count
            EnabledOwners  = $enabledOwners
            MemberCount    = $memberList.Count
            GuestCount     = $guestCount
            Labels         = if ($g.AssignedLabels) { ($g.AssignedLabels | ForEach-Object { $_.DisplayName }) -join '; ' } else { '' }
        }
    }

    Write-Host ""  # close the progress line
    Write-Ok "Collected $($records.Count) team record(s)."
    Write-Host ""
    return @($records)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTIONS 2-5 — DERIVED FINDINGS FROM THE INVENTORY
# ─────────────────────────────────────────────────────────────────────────────

# Thresholds editable here; section 2+ functions are thin filters over the
# inventory so the audit stays deterministic and cheap.
$script:LargeTeamThreshold = 250
$script:StaleRenewalDays   = 365

function Get-OrphanTeams {
    param([array]$Teams)
    # Orphan = zero owners at all, OR every owner is disabled.
    @($Teams | Where-Object {
        $_.OwnerCount -eq 0 -or ($_.OwnerCount -gt 0 -and $_.EnabledOwners -eq 0)
    })
}

function Get-PublicTeams {
    param([array]$Teams)
    @($Teams | Where-Object { $_.Visibility -eq 'Public' })
}

function Get-GuestHeavyTeams {
    param([array]$Teams)
    # Any team with at least one guest member; the flagging is done in the UI.
    @($Teams | Where-Object { $_.GuestCount -gt 0 } | Sort-Object GuestCount -Descending)
}

function Get-LargeTeams {
    param([array]$Teams)
    @($Teams | Where-Object { $_.MemberCount -ge $script:LargeTeamThreshold } | Sort-Object MemberCount -Descending)
}

function Get-StaleTeams {
    param([array]$Teams)
    # M365 groups can expire if a group-expiration policy is set; RenewedDateTime
    # is the last time the owner confirmed the group is still in use. A long gap
    # suggests the team has lapsed out of active management.
    $threshold = (Get-Date).AddDays(-$script:StaleRenewalDays)
    @($Teams | Where-Object {
        $_.LastRenewed -and ([datetime]$_.LastRenewed) -lt $threshold
    } | Sort-Object LastRenewed)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Teams, [array]$Orphans, [array]$Public, [array]$Guests, [array]$Large, [array]$Stale)

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $tkCfg         = Get-TKConfig
    $orgPrefix     = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _inv-row {
        param($t)
        $vizBadge = switch ($t.Visibility) {
            'Public'  { "<span class='tk-badge-warn'>Public</span>" }
            'Private' { "<span class='tk-badge-ok'>Private</span>" }
            default   { "<span class='tk-badge-info'>$(EscHtml $t.Visibility)</span>" }
        }
        $ownerBadge = if ($t.OwnerCount -eq 0) { "<span class='tk-badge-err'>$($t.OwnerCount)</span>" }
                      elseif ($t.EnabledOwners -eq 0) { "<span class='tk-badge-err'>$($t.OwnerCount) (all disabled)</span>" }
                      else { "<span class='tk-badge-ok'>$($t.OwnerCount) ($($t.EnabledOwners) enabled)</span>" }
        $guestBadge = if ($t.GuestCount -eq 0) { "<span class='tk-badge-ok'>0</span>" } else { "<span class='tk-badge-warn'>$($t.GuestCount)</span>" }
        $memBadge   = if ($t.MemberCount -ge $script:LargeTeamThreshold) { "<span class='tk-badge-warn'>$($t.MemberCount)</span>" } else { "<span class='tk-badge-info'>$($t.MemberCount)</span>" }
        return "<tr><td>$(EscHtml $t.DisplayName)</td><td>$vizBadge</td><td>$ownerBadge</td><td>$memBadge</td><td>$guestBadge</td><td>$(EscHtml $t.Created)</td><td>$(EscHtml $t.LastRenewed)</td><td>$(EscHtml $t.Labels)</td></tr>"
    }

    $invRows = [System.Text.StringBuilder]::new()
    if ($Teams.Count -eq 0) {
        [void]$invRows.Append("<tr><td colspan='8' class='tk-badge-info' style='text-align:center;'>No teams found.</td></tr>")
    } else {
        foreach ($t in ($Teams | Sort-Object DisplayName)) { [void]$invRows.Append((_inv-row $t) + "`n") }
    }

    function _category-rows {
        param([array]$Rows, [int]$Cols, [string]$EmptyMessage)
        $sb = [System.Text.StringBuilder]::new()
        if ($Rows.Count -eq 0) {
            [void]$sb.Append("<tr><td colspan='$Cols' class='tk-badge-ok' style='text-align:center;'>$EmptyMessage</td></tr>")
        } else {
            foreach ($t in $Rows) { [void]$sb.Append((_inv-row $t) + "`n") }
        }
        return $sb.ToString()
    }

    $orphanRows = _category-rows -Rows $Orphans -Cols 8 -EmptyMessage 'No orphan teams.'
    $publicRows = _category-rows -Rows $Public  -Cols 8 -EmptyMessage 'No public teams.'
    $guestRows  = _category-rows -Rows $Guests  -Cols 8 -EmptyMessage 'No teams contain guest members.'
    $largeRows  = _category-rows -Rows $Large   -Cols 8 -EmptyMessage "No teams exceed $script:LargeTeamThreshold members."
    $staleRows  = _category-rows -Rows $Stale   -Cols 8 -EmptyMessage "No teams have gone unrenewed for $script:StaleRenewalDays+ days."

    $htmlHead = Get-TKHtmlHead `
        -Title      'A.S.T.E.R.I.S.M. Teams Audit Report' `
        -ScriptName 'A.S.T.E.R.I.S.M.' `
        -Subtitle   "${orgPrefix}Microsoft Teams Estate Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = EscHtml $script:ConnectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @('Teams Inventory', 'Orphan Teams', 'Public Teams', 'Guest Members', 'Large Teams', 'Stale Teams')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'A.S.T.E.R.I.S.M. v6.0'

    $orphClass  = if ($Orphans.Count -gt 0) { 'err' } else { 'ok' }
    $pubClass   = if ($Public.Count  -gt 0) { 'warn' } else { 'ok' }
    $guestClass = if ($Guests.Count  -gt 0) { 'warn' } else { 'ok' }
    $staleClass = if ($Stale.Count   -gt 0) { 'warn' } else { 'ok' }

    $invHeader = "<tr><th>Display Name</th><th>Visibility</th><th>Owners</th><th>Members</th><th>Guests</th><th>Created</th><th>Renewed</th><th>Sensitivity Labels</th></tr>"

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Teams.Count)</div><div class="tk-summary-lbl">Teams Total</div></div>
    <div class="tk-summary-card $orphClass"><div class="tk-summary-num">$($Orphans.Count)</div><div class="tk-summary-lbl">Orphan Teams</div></div>
    <div class="tk-summary-card $pubClass"><div class="tk-summary-num">$($Public.Count)</div><div class="tk-summary-lbl">Public Teams</div></div>
    <div class="tk-summary-card $guestClass"><div class="tk-summary-num">$($Guests.Count)</div><div class="tk-summary-lbl">Teams With Guests</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Large.Count)</div><div class="tk-summary-lbl">Large Teams (&ge; $script:LargeTeamThreshold)</div></div>
    <div class="tk-summary-card $staleClass"><div class="tk-summary-num">$($Stale.Count)</div><div class="tk-summary-lbl">Stale (&gt; $script:StaleRenewalDays d unrenewed)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Teams Inventory</span><span class="tk-section-num">$($Teams.Count) team(s)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$($invRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Orphan Teams -- No Owners (or All Owners Disabled)</span><span class="tk-section-num">$($Orphans.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$orphanRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Public Teams -- Anyone in Tenant Can Join</span><span class="tk-section-num">$($Public.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$publicRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Teams With Guest Members</span><span class="tk-section-num">$($Guests.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$guestRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Large Teams -- &ge; $script:LargeTeamThreshold Members</span><span class="tk-section-num">$($Large.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$largeRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Stale Teams -- Unrenewed &gt; $script:StaleRenewalDays Days</span><span class="tk-section-num">$($Stale.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$staleRows</tbody></table></div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Teams, [array]$Orphans, [array]$Public, [array]$Guests, [array]$Large, [array]$Stale)

    Write-Section "EXPORTING HTML REPORT"
    $html      = Build-HtmlReport -Teams $Teams -Orphans $Orphans -Public $Public -Guests $Guests -Large $Large -Stale $Stale
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "ASTERISM_${timestamp}.html"

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-AsterismBanner
    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  Teams inventory (refresh underlying data)" -ForegroundColor $C.Info
    Write-Host "  3  Export full HTML report (orphans, public, guests, large, stale)" -ForegroundColor $C.Success
    Write-Host "  Q  Quit" -ForegroundColor $C.Info
    Write-Host ""
}

function Assert-Connected {
    if (-not $script:Connected) {
        Write-Host ""
        Write-Warn "Not connected. Please select option 1 first."
        Write-Host ""
        Read-Host "  Press Enter to return to menu"
        return $false
    }
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

Show-AsterismBanner
Install-GraphModule

if (Test-GraphConnection) {
    Write-Ok "Already connected as $($script:ConnectedAs)"
    Write-Host ""
}

if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    if (-not $script:Connected) { Invoke-Connect }
    if (-not $script:Connected) { Write-Fail "Could not connect to Graph. Exiting."; exit 1 }

    $teams   = Get-TeamsInventory
    $orphans = Get-OrphanTeams    -Teams $teams
    $public  = Get-PublicTeams    -Teams $teams
    $guests  = Get-GuestHeavyTeams -Teams $teams
    $large   = Get-LargeTeams     -Teams $teams
    $stale   = Get-StaleTeams     -Teams $teams

    Export-HtmlReport -Teams $teams -Orphans $orphans -Public $public -Guests $guests -Large $large -Stale $stale

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

$cachedTeams = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {
        '1' {
            Invoke-Connect
            $cachedTeams = $null
            Read-Host "  Press Enter to return to menu"
        }
        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedTeams = Get-TeamsInventory
            Read-Host "  Press Enter to return to menu"
        }
        '3' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedTeams) { $cachedTeams = Get-TeamsInventory }
            $orphans = Get-OrphanTeams     -Teams $cachedTeams
            $public  = Get-PublicTeams     -Teams $cachedTeams
            $guests  = Get-GuestHeavyTeams -Teams $cachedTeams
            $large   = Get-LargeTeams      -Teams $cachedTeams
            $stale   = Get-StaleTeams      -Teams $cachedTeams
            Export-HtmlReport -Teams $cachedTeams -Orphans $orphans -Public $public -Guests $guests -Large $large -Stale $stale
            Read-Host "  Press Enter to return to menu"
        }
        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {
                # Best-effort cleanup.
            }
            Write-Host "  Goodbye." -ForegroundColor $C.Header
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Warn "Invalid option. Please choose 1, 2, 3, or Q."
            Write-Host ""
            Start-Sleep -Milliseconds 800
        }
    }
} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
