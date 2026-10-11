# eclipse.ps1 - E.C.L.I.P.S.E. — Entra Credentials: Lapsed, Idle, Privileged, Stale & External
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
    E.C.L.I.P.S.E. — Entra Credentials: Lapsed, Idle, Privileged, Stale & External
    Entra ID Identity Hygiene Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and audits the Entra ID (Azure AD) identity
    posture: guest users with stale sign-ins, privileged role holders and
    their MFA state, members with password-never-expires set, stale
    privileged users, and disabled-but-licensed accounts that still consume
    paid SKUs. Produces a dark-themed HTML report that complements
    A.L.M.A.N.A.C. (licensing / mailboxes) and O.R.B.I.T. (device
    compliance) — together the three tools give a full tenant posture
    report.

.USAGE
    PS C:\> .\eclipse.ps1                    # Interactive menu
    PS C:\> .\eclipse.ps1 -Unattended        # Silent mode — auto-connect and export HTML

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
    'User.Read.All',
    'Directory.Read.All',
    'AuditLog.Read.All',
    'RoleManagement.Read.Directory'
)

function Show-EclipseBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  E.C.L.I.P.S.E. — Entra Credentials: Lapsed, Idle, Privileged, Stale & External" -ForegroundColor Cyan
    Write-Host "  Entra ID Identity Hygiene Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Helpers, audit functions, and entry point appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH MODULE INSTALL + CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $required = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Identity.Governance'
    )

    $needsInstall = $false
    foreach ($m in $required) {
        if (Get-Module -ListAvailable -Name $m) {
            Write-Ok "$m — installed"
        } else {
            Write-Warn "$m — NOT found"
            $needsInstall = $true
        }
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
                Write-TKError -ScriptName 'eclipse' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
                exit 1
            }
        } else {
            $ans = Read-Host "  Install Microsoft.Graph for current user? [Y/N]"
            if ($ans -match '^[Yy]') {
                Write-Step "Installing Microsoft.Graph..."
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Ok "Microsoft.Graph installed."
                } catch {
                    Write-Fail "Install failed: $_"
                    exit 1
                }
            } else {
                Write-Fail "Microsoft.Graph is required. Exiting."
                exit 1
            }
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
                # Best-effort metadata.
            }
            return $true
        }
    } catch {
        # No cached context.
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
        Write-TKError -ScriptName 'eclipse' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — GUEST USERS
# ─────────────────────────────────────────────────────────────────────────────

# Every guest with their last sign-in and creation date. Stale guests
# (no sign-in in 90 days) are the single easiest external-access cleanup win.
function Get-GuestAudit {
    Write-Section "GUEST USERS"
    Write-Step "Retrieving guest accounts and sign-in activity..."

    $threshold = (Get-Date).AddDays(-90)

    try {
        $guests = Get-MgUser -All `
            -Filter "userType eq 'Guest'" `
            -Property "displayName,userPrincipalName,mail,createdDateTime,accountEnabled,signInActivity,externalUserState" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve guests: $_"
        Write-TKError -ScriptName 'eclipse' -Message "Get-MgUser (guest filter) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $now = Get-Date
    $rows = foreach ($g in $guests) {
        $lastSignIn   = $null
        $daysInactive = $null
        if ($g.SignInActivity -and $g.SignInActivity.LastSignInDateTime) {
            $lastSignIn   = [datetime]$g.SignInActivity.LastSignInDateTime
            $daysInactive = [math]::Round(($now - $lastSignIn).TotalDays, 0)
        }

        $isStale = (-not $lastSignIn) -or ($lastSignIn -lt $threshold)

        [PSCustomObject]@{
            DisplayName   = $g.DisplayName
            UPN           = $g.UserPrincipalName
            Email         = $g.Mail
            Created       = if ($g.CreatedDateTime) { ([datetime]$g.CreatedDateTime).ToString('yyyy-MM-dd') } else { '—' }
            LastSignIn    = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
            DaysInactive  = $daysInactive
            IsStale       = $isStale
            AccountState  = if ($g.AccountEnabled) { 'Enabled' } else { 'Disabled' }
            InviteState   = if ($g.ExternalUserState) { $g.ExternalUserState } else { '—' }
        }
    }

    $sorted = @($rows | Sort-Object @{ Expression = { if ($null -eq $_.DaysInactive) { 99999 } else { [int]$_.DaysInactive } }; Descending = $true })

    Write-Ok "$($sorted.Count) guest account(s) found."
    $staleCount = @($sorted | Where-Object { $_.IsStale }).Count
    if ($staleCount -gt 0) { Write-Warn "$staleCount stale guest(s) — no sign-in in 90+ days." }
    Write-Host ""
    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — PRIVILEGED ROLE HOLDERS
# ─────────────────────────────────────────────────────────────────────────────

# Every member of every active directory role. Emits one row per user-per-role
# so users with multiple roles appear multiple times (intentional — the audit
# question is "who has this role", not "who are all the admins").
function Get-PrivilegedRoleAudit {
    Write-Section "PRIVILEGED ROLE HOLDERS"
    Write-Step "Retrieving active directory role assignments..."

    try {
        $roles = Get-MgDirectoryRole -All -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve directory roles: $_"
        Write-TKError -ScriptName 'eclipse' -Message "Get-MgDirectoryRole failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($role in $roles) {
        try {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
        } catch {
            continue
        }

        foreach ($m in $members) {
            # Role members can be users, groups, or service principals — pull the
            # fields present on user objects and mark non-user members distinctly.
            $props = $m.AdditionalProperties
            $upn   = $props['userPrincipalName']
            $disp  = $props['displayName']
            $typeRaw = $props['@odata.type']
            $type  = if ($typeRaw) { ($typeRaw -replace '#microsoft.graph.', '') } else { 'unknown' }

            $rows.Add([PSCustomObject]@{
                RoleName    = $role.DisplayName
                MemberType  = $type
                DisplayName = $disp
                UPN         = $upn
                MemberId    = $m.Id
            })
        }
    }

    $sorted = @($rows | Sort-Object RoleName, DisplayName)
    $unique = @($sorted | Group-Object UPN | Where-Object { $_.Name }).Count
    Write-Ok "$($sorted.Count) role assignment(s) covering $unique unique principal(s)."

    # Surface the high-tier roles in the console summary.
    $highTier = @('Global Administrator','Privileged Role Administrator','User Administrator','Exchange Administrator','SharePoint Administrator','Security Administrator','Conditional Access Administrator')
    foreach ($role in ($sorted | Where-Object { $highTier -contains $_.RoleName } | Group-Object RoleName)) {
        Write-Warn ("  {0,-35} {1} member(s)" -f $role.Name, $role.Count)
    }
    Write-Host ""

    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — PASSWORD NEVER EXPIRES
# ─────────────────────────────────────────────────────────────────────────────

# Cloud-only accounts with PasswordPolicies=DisablePasswordExpiration are the
# classic tenant hygiene miss — shared mailboxes, service accounts, or
# holdovers from a failed sync that quietly disabled expiry.
function Get-PasswordNeverExpiresAudit {
    Write-Section "PASSWORD NEVER EXPIRES"
    Write-Step "Retrieving members with password-never-expires set..."

    try {
        $users = Get-MgUser -All `
            -Property "displayName,userPrincipalName,passwordPolicies,accountEnabled,userType,onPremisesSyncEnabled,assignedLicenses" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        Write-TKError -ScriptName 'eclipse' -Message "Get-MgUser (all) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $rows = foreach ($u in $users) {
        if ($u.UserType -eq 'Guest') { continue }
        $policies = if ($u.PasswordPolicies) { $u.PasswordPolicies } else { '' }
        if ($policies -notmatch 'DisablePasswordExpiration') { continue }

        [PSCustomObject]@{
            DisplayName    = $u.DisplayName
            UPN            = $u.UserPrincipalName
            AccountEnabled = $u.AccountEnabled
            SyncedFromAD   = [bool]$u.OnPremisesSyncEnabled
            Licensed       = ($u.AssignedLicenses.Count -gt 0)
            PasswordPolicy = $policies
        }
    }

    $rows = @($rows | Sort-Object DisplayName)
    if ($rows.Count -eq 0) {
        Write-Ok "No members have password-never-expires set."
    } else {
        Write-Warn "$($rows.Count) member(s) with password-never-expires."
    }
    Write-Host ""
    return $rows
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — STALE PRIVILEGED USERS
# ─────────────────────────────────────────────────────────────────────────────

# Admins who haven't signed in in 60+ days. Same risk-class as "orphan"
# service accounts — their credentials still work and still have elevated
# access, but no one is watching the audit log.
function Get-StaleAdminAudit {
    param([array]$RoleRows)

    Write-Section "STALE PRIVILEGED USERS"
    Write-Step "Checking sign-in activity for privileged users (no sign-in in 60+ days)..."

    $threshold = (Get-Date).AddDays(-60)
    $now       = Get-Date

    # Deduplicate by UPN — a user in multiple roles is still one stale admin.
    $uniqueUpns = @($RoleRows | Where-Object { $_.MemberType -eq 'user' -and $_.UPN } | Select-Object -ExpandProperty UPN -Unique)

    if ($uniqueUpns.Count -eq 0) {
        Write-Warn "No user-type role members to evaluate."
        Write-Host ""
        return @()
    }

    $stale = [System.Collections.Generic.List[object]]::new()
    foreach ($upn in $uniqueUpns) {
        try {
            $u = Get-MgUser -UserId $upn -Property "displayName,userPrincipalName,accountEnabled,signInActivity" -ErrorAction Stop
        } catch {
            continue
        }

        $lastSignIn   = $null
        $daysInactive = $null
        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn   = [datetime]$u.SignInActivity.LastSignInDateTime
            $daysInactive = [math]::Round(($now - $lastSignIn).TotalDays, 0)
        }

        $isStale = (-not $lastSignIn) -or ($lastSignIn -lt $threshold)
        if (-not $isStale) { continue }

        $roles = @($RoleRows | Where-Object { $_.UPN -eq $upn } | Select-Object -ExpandProperty RoleName -Unique) -join ', '

        $stale.Add([PSCustomObject]@{
            DisplayName    = $u.DisplayName
            UPN            = $u.UserPrincipalName
            AccountEnabled = $u.AccountEnabled
            LastSignIn     = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
            DaysInactive   = $daysInactive
            Roles          = $roles
        })
    }

    $sorted = @($stale | Sort-Object @{ Expression = { if ($null -eq $_.DaysInactive) { 99999 } else { [int]$_.DaysInactive } }; Descending = $true })
    if ($sorted.Count -eq 0) {
        Write-Ok "All privileged users have signed in within the last 60 days."
    } else {
        Write-Fail "$($sorted.Count) privileged user(s) inactive 60+ days."
    }
    Write-Host ""
    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — DISABLED BUT LICENSED
# ─────────────────────────────────────────────────────────────────────────────

# Cost-leak audit: disabled accounts still consuming paid SKUs. Common after
# off-boarding where the license was never reclaimed.
function Get-DisabledLicensedAudit {
    Write-Section "DISABLED BUT LICENSED"
    Write-Step "Finding disabled accounts that still have license assignments..."

    try {
        $users = Get-MgUser -All `
            -Filter "accountEnabled eq false" `
            -Property "displayName,userPrincipalName,accountEnabled,assignedLicenses,userType" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        Write-TKError -ScriptName 'eclipse' -Message "Get-MgUser (disabled filter) failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    $rows = foreach ($u in $users) {
        if ($u.UserType -eq 'Guest') { continue }
        if ($u.AssignedLicenses.Count -eq 0) { continue }
        [PSCustomObject]@{
            DisplayName   = $u.DisplayName
            UPN           = $u.UserPrincipalName
            LicenseCount  = $u.AssignedLicenses.Count
        }
    }

    $sorted = @($rows | Sort-Object DisplayName)
    if ($sorted.Count -eq 0) {
        Write-Ok "No disabled accounts are consuming licenses."
    } else {
        Write-Fail "$($sorted.Count) disabled account(s) still consuming licenses."
    }
    Write-Host ""
    return $sorted
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Guests, [array]$Roles, [array]$PwdNever, [array]$StaleAdmins, [array]$DisabledLic)

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $connectedAs   = EscHtml $script:ConnectedAs

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    # Guest rows
    $guestRows = [System.Text.StringBuilder]::new()
    if ($Guests.Count -eq 0) {
        [void]$guestRows.Append("<tr><td colspan='6' class='tk-badge-ok' style='text-align:center;'>No guest accounts.</td></tr>")
    } else {
        foreach ($g in $Guests) {
            $staleClass = if ($g.IsStale) { 'tk-badge-warn' } else { 'tk-badge-ok' }
            $staleLabel = if ($g.LastSignIn -eq 'Never') { 'Never' } else { $g.LastSignIn }
            [void]$guestRows.Append("<tr><td>$(EscHtml $g.DisplayName)</td><td>$(EscHtml $g.UPN)</td><td>$(EscHtml $g.Created)</td><td><span class='$staleClass'>$(EscHtml $staleLabel)</span></td><td>$(EscHtml $g.AccountState)</td><td>$(EscHtml $g.InviteState)</td></tr>`n")
        }
    }

    # Role rows
    $roleRows = [System.Text.StringBuilder]::new()
    if ($Roles.Count -eq 0) {
        [void]$roleRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No directory role assignments.</td></tr>")
    } else {
        foreach ($r in $Roles) {
            $typeBadge = switch ($r.MemberType) {
                'user'             { "<span class='tk-badge-ok'>User</span>" }
                'group'            { "<span class='tk-badge-warn'>Group</span>" }
                'servicePrincipal' { "<span class='tk-badge-info'>Service Principal</span>" }
                default            { "<span class='tk-badge-info'>$(EscHtml $r.MemberType)</span>" }
            }
            [void]$roleRows.Append("<tr><td>$(EscHtml $r.RoleName)</td><td>$typeBadge</td><td>$(EscHtml $r.DisplayName)</td><td>$(EscHtml $r.UPN)</td></tr>`n")
        }
    }

    # Password-never-expires rows
    $pwdRows = [System.Text.StringBuilder]::new()
    if ($PwdNever.Count -eq 0) {
        [void]$pwdRows.Append("<tr><td colspan='5' class='tk-badge-ok' style='text-align:center;'>No members have password-never-expires set.</td></tr>")
    } else {
        foreach ($p in $PwdNever) {
            $enBadge   = if ($p.AccountEnabled) { "<span class='tk-badge-ok'>Enabled</span>" } else { "<span class='tk-badge-info'>Disabled</span>" }
            $syncBadge = if ($p.SyncedFromAD)   { "<span class='tk-badge-info'>AD sync</span>" } else { "<span class='tk-badge-warn'>Cloud-only</span>" }
            $licBadge  = if ($p.Licensed)       { "<span class='tk-badge-ok'>Licensed</span>" } else { "<span class='tk-badge-info'>Unlicensed</span>" }
            [void]$pwdRows.Append("<tr><td>$(EscHtml $p.DisplayName)</td><td>$(EscHtml $p.UPN)</td><td>$enBadge</td><td>$syncBadge</td><td>$licBadge</td></tr>`n")
        }
    }

    # Stale admin rows
    $staleRows = [System.Text.StringBuilder]::new()
    if ($StaleAdmins.Count -eq 0) {
        [void]$staleRows.Append("<tr><td colspan='5' class='tk-badge-ok' style='text-align:center;'>All privileged users have recent sign-in activity.</td></tr>")
    } else {
        foreach ($s in $StaleAdmins) {
            $lastClass = if ($s.LastSignIn -eq 'Never') { 'tk-badge-err' } else { 'tk-badge-warn' }
            [void]$staleRows.Append("<tr><td>$(EscHtml $s.DisplayName)</td><td>$(EscHtml $s.UPN)</td><td>$(EscHtml $s.Roles)</td><td><span class='$lastClass'>$(EscHtml $s.LastSignIn)</span></td><td>$($s.DaysInactive)</td></tr>`n")
        }
    }

    # Disabled-licensed rows
    $dlRows = [System.Text.StringBuilder]::new()
    if ($DisabledLic.Count -eq 0) {
        [void]$dlRows.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;'>No disabled accounts are consuming licenses.</td></tr>")
    } else {
        foreach ($d in $DisabledLic) {
            [void]$dlRows.Append("<tr><td>$(EscHtml $d.DisplayName)</td><td>$(EscHtml $d.UPN)</td><td><span class='tk-badge-err'>$($d.LicenseCount)</span></td></tr>`n")
        }
    }

    # Summary cards
    $staleGuestCount = @($Guests | Where-Object { $_.IsStale }).Count
    $gClass = if ($staleGuestCount -gt 0) { 'warn' } else { 'ok' }
    $pClass = if ($PwdNever.Count    -gt 0) { 'warn' } else { 'ok' }
    $sClass = if ($StaleAdmins.Count -gt 0) { 'err'  } else { 'ok' }
    $dClass = if ($DisabledLic.Count -gt 0) { 'err'  } else { 'ok' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'E.C.L.I.P.S.E. Identity Hygiene Report' `
        -ScriptName 'E.C.L.I.P.S.E.' `
        -Subtitle   "${orgPrefix}Entra ID Identity Hygiene Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = $connectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @('Guests', 'Privileged Roles', 'Password Never Expires', 'Stale Admins', 'Disabled but Licensed')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'E.C.L.I.P.S.E. v6.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Guests.Count)</div><div class="tk-summary-lbl">Guest Accounts</div></div>
    <div class="tk-summary-card $gClass"><div class="tk-summary-num">$staleGuestCount</div><div class="tk-summary-lbl">Stale Guests (90d+)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Roles.Count)</div><div class="tk-summary-lbl">Role Assignments</div></div>
    <div class="tk-summary-card $pClass"><div class="tk-summary-num">$($PwdNever.Count)</div><div class="tk-summary-lbl">Password Never Expires</div></div>
    <div class="tk-summary-card $sClass"><div class="tk-summary-num">$($StaleAdmins.Count)</div><div class="tk-summary-lbl">Stale Admins (60d+)</div></div>
    <div class="tk-summary-card $dClass"><div class="tk-summary-num">$($DisabledLic.Count)</div><div class="tk-summary-lbl">Disabled but Licensed</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Guest Accounts</span><span class="tk-section-num">$($Guests.Count) guest(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Created</th><th>Last Sign-In</th><th>Account</th><th>Invite State</th></tr></thead>
        <tbody>$($guestRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Privileged Role Holders</span><span class="tk-section-num">$($Roles.Count) assignment(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Role</th><th>Type</th><th>Display Name</th><th>UPN</th></tr></thead>
        <tbody>$($roleRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Password Never Expires</span><span class="tk-section-num">$($PwdNever.Count) user(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Account</th><th>Source</th><th>Licensed</th></tr></thead>
        <tbody>$($pwdRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Stale Privileged Users -- No Sign-In Within 60 Days</span><span class="tk-section-num">$($StaleAdmins.Count) user(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Roles</th><th>Last Sign-In</th><th>Days Inactive</th></tr></thead>
        <tbody>$($staleRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Disabled but Licensed</span><span class="tk-section-num">$($DisabledLic.Count) account(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Display Name</th><th>UPN</th><th>Licenses Assigned</th></tr></thead>
        <tbody>$($dlRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Guests, [array]$Roles, [array]$PwdNever, [array]$StaleAdmins, [array]$DisabledLic)

    Write-Section "EXPORTING HTML REPORT"
    Write-Step "Building report..."

    $html      = Build-HtmlReport -Guests $Guests -Roles $Roles -PwdNever $PwdNever -StaleAdmins $StaleAdmins -DisabledLic $DisabledLic
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "ECLIPSE_${timestamp}.html"

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
    Show-EclipseBanner

    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  Guest accounts — sign-in age, invite state" -ForegroundColor $C.Info
    Write-Host "  3  Privileged role holders — every directory role member" -ForegroundColor $C.Info
    Write-Host "  4  Password never expires — members with expiry disabled" -ForegroundColor $C.Warning
    Write-Host "  5  Stale privileged users — admins inactive 60+ days" -ForegroundColor $C.Warning
    Write-Host "  6  Disabled but licensed — accounts leaking license cost" -ForegroundColor $C.Warning
    Write-Host "  7  Export full HTML report (all findings)" -ForegroundColor $C.Success
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

Show-EclipseBanner
Install-GraphModule

if (Test-GraphConnection) {
    Write-Ok "Already connected as $($script:ConnectedAs)"
    Write-Host ""
}

if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    if (-not $script:Connected) { Invoke-Connect }
    if (-not $script:Connected) {
        Write-Fail "Could not connect to Graph. Exiting."
        exit 1
    }

    $guests      = Get-GuestAudit
    $roles       = Get-PrivilegedRoleAudit
    $pwdNever    = Get-PasswordNeverExpiresAudit
    $staleAdmins = Get-StaleAdminAudit -RoleRows $roles
    $disabledLic = Get-DisabledLicensedAudit

    Export-HtmlReport -Guests $guests -Roles $roles -PwdNever $pwdNever -StaleAdmins $staleAdmins -DisabledLic $disabledLic

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# Cached per-session so repeated menu choices don't re-query
$cachedGuests    = $null
$cachedRoles     = $null
$cachedPwdNever  = $null
$cachedStale     = $null
$cachedDisabled  = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {

        '1' {
            Invoke-Connect
            $cachedGuests = $null; $cachedRoles = $null; $cachedPwdNever = $null; $cachedStale = $null; $cachedDisabled = $null
            Read-Host "  Press Enter to return to menu"
        }

        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedGuests = Get-GuestAudit
            Read-Host "  Press Enter to return to menu"
        }

        '3' {
            if (-not (Assert-Connected)) { break }
            $cachedRoles = Get-PrivilegedRoleAudit
            Read-Host "  Press Enter to return to menu"
        }

        '4' {
            if (-not (Assert-Connected)) { break }
            $cachedPwdNever = Get-PasswordNeverExpiresAudit
            Read-Host "  Press Enter to return to menu"
        }

        '5' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedRoles) { $cachedRoles = Get-PrivilegedRoleAudit }
            $cachedStale = Get-StaleAdminAudit -RoleRows $cachedRoles
            Read-Host "  Press Enter to return to menu"
        }

        '6' {
            if (-not (Assert-Connected)) { break }
            $cachedDisabled = Get-DisabledLicensedAudit
            Read-Host "  Press Enter to return to menu"
        }

        '7' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedGuests)   { $cachedGuests   = Get-GuestAudit }
            if ($null -eq $cachedRoles)    { $cachedRoles    = Get-PrivilegedRoleAudit }
            if ($null -eq $cachedPwdNever) { $cachedPwdNever = Get-PasswordNeverExpiresAudit }
            if ($null -eq $cachedStale)    { $cachedStale    = Get-StaleAdminAudit -RoleRows $cachedRoles }
            if ($null -eq $cachedDisabled) { $cachedDisabled = Get-DisabledLicensedAudit }
            Export-HtmlReport -Guests $cachedGuests -Roles $cachedRoles -PwdNever $cachedPwdNever -StaleAdmins $cachedStale -DisabledLic $cachedDisabled
            Read-Host "  Press Enter to return to menu"
        }

        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {
                # Best-effort cleanup — cached context may already be gone.
            }
            Write-Host "  Goodbye." -ForegroundColor $C.Header
            Write-Host ""
        }

        default {
            Write-Host ""
            Write-Warn "Invalid option. Please choose 1–7 or Q."
            Write-Host ""
            Start-Sleep -Milliseconds 800
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
