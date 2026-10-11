# wisp.ps1 - W.I.S.P. — Wireless Inventory & Security Profiler
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
    W.I.S.P. — Wireless Inventory & Security Profiler
    Wi-Fi Profile Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits saved Wi-Fi (WLAN) profiles on a Windows endpoint. Exports
    each profile to XML via `netsh wlan export profile ... key=clear`,
    parses the locale-stable XML schema, and produces a dark-themed
    HTML report covering authentication / encryption / connection mode
    / auto-switch / hidden-SSID / MAC-randomization for every saved
    profile, plus filtered tables for open, weak, and auto-connecting
    networks.

    Read-only audit. Key material is masked by default and only
    rendered in cleartext when `-IncludeKey` is explicitly passed.

.USAGE
    PS C:\> .\wisp.ps1                    # Interactive run; key material masked
    PS C:\> .\wisp.ps1 -Unattended        # Silent: export HTML and exit
    PS C:\> .\wisp.ps1 -IncludeKey        # Render cleartext PSKs in the report (technician-managed audit only)

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [switch]$IncludeKey
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

function Show-WispBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  W.I.S.P. -- Wireless Inventory & Security Profiler" -ForegroundColor Yellow
    Write-Host "  Wi-Fi Profile Audit  v6.0" -ForegroundColor Yellow
    if ($IncludeKey) {
        Write-Host "  *** Key material WILL be rendered in cleartext (-IncludeKey) ***" -ForegroundColor Magenta
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Authentication strength tier. Drives badge colour and verdict.
# Values come from the WLAN profile XML schema:
# learn.microsoft.com/en-us/windows/win32/nativewifi/wlan-profileschema-authentication-authencryption-element
$AuthStrength = @{
    'open'       = 'Insecure'
    'shared'     = 'Insecure'   # Shared-key WEP
    'WEP'        = 'Insecure'
    'WPA'        = 'Weak'
    'WPAPSK'     = 'Weak'
    'WPA2'       = 'Strong'     # WPA2 enterprise
    'WPA2PSK'    = 'Strong'
    'WPA3'       = 'Strong'
    'WPA3SAE'    = 'Strong'
    'WPA3ENT'    = 'Strong'
    'WPA3ENT192' = 'Strong'
    'OWE'        = 'Strong'     # Open with opportunistic encryption
}

# Cipher tier. TKIP and WEP are deprecated; AES (CCMP) and GCMP are current.
$CipherStrength = @{
    'none' = 'Insecure'
    'WEP'  = 'Insecure'
    'TKIP' = 'Weak'
    'AES'  = 'Strong'   # CCMP
    'GCMP' = 'Strong'
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-WifiAdapters {
    # PhysicalMediaType 'Native 802.11' is the canonical filter for Wi-Fi NICs.
    try {
        $rows = Get-NetAdapter -Physical -ErrorAction Stop | Where-Object {
            $_.PhysicalMediaType -eq 'Native 802.11' -or $_.MediaType -eq '802.11'
        } | ForEach-Object {
            [PSCustomObject]@{
                Name              = $_.Name
                InterfaceDescription = $_.InterfaceDescription
                Status            = $_.Status
                LinkSpeed         = $_.LinkSpeed
                MacAddress        = $_.MacAddress
                DriverVersion     = $_.DriverVersion
                DriverDate        = $_.DriverDate
            }
        }
        return @($rows)
    } catch {
        return @()
    }
}

function Get-WlanProfileList {
    # `netsh wlan show profiles` output is locale-dependent for headers but
    # the per-profile lines all have the form `    <something> : <profileName>`.
    # We extract the trailing token (the profile name) from any line containing
    # a colon and indented two-plus spaces. Empty / non-matching lines are skipped.
    $raw = & netsh wlan show profiles 2>$null
    if (-not $raw) { return @() }

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $raw) {
        if ($line -match '^\s+[^:]+:\s+(.+?)\s*$') {
            $candidate = $matches[1].Trim()
            # Skip anything that's clearly a header (contains keywords we never see in real names).
            if ($candidate -and $candidate -notmatch '^(Hosted network|Group policy|User|All User|Granted Permission|Restrictions)') {
                $names.Add($candidate)
            }
        }
    }
    # Names can repeat across User / All-User scopes. De-duplicate but keep order.
    return @($names | Select-Object -Unique)
}

function Export-WlanProfileXml {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$Folder
    )

    # `key=clear` requires elevation. If we're elevated (Invoke-AdminElevation
    # already ensured this) we get cleartext keys; otherwise we get DPAPI blobs
    # we can't decode. Suppress the netsh stdout / stderr noise -- callers only
    # care about the XML file landing on disk.
    $null = & netsh wlan export profile name="$Name" folder="$Folder" key=clear 2>&1
    # netsh emits the file name from the SSID name with characters sanitised:
    # spaces -> hyphens, etc. The simplest reliable approach is to glob the
    # folder for any new XML file. Caller passes a dedicated empty folder per
    # profile, so we just take the first XML there.
    $xml = Get-ChildItem -Path $Folder -Filter '*.xml' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $xml) { return $null }
    return $xml.FullName
}

function ConvertFrom-WlanProfileXml {
    param([Parameter(Mandatory)] [string]$XmlPath)

    try {
        [xml]$doc = Get-Content -Path $XmlPath -Raw -ErrorAction Stop
    } catch {
        return $null
    }

    # The XML uses a default namespace; XPath queries must declare it. We pluck
    # values by walking the DOM directly to avoid an XmlNamespaceManager dance.
    $wlanProfile = $doc.WLANProfile
    if (-not $wlanProfile) { return $null }

    $name           = $wlanProfile.name
    $ssidName       = $null
    $ssidHex        = $null
    $nonBroadcast   = $false
    if ($wlanProfile.SSIDConfig -and $wlanProfile.SSIDConfig.SSID) {
        $ssidName = $wlanProfile.SSIDConfig.SSID.name
        $ssidHex  = $wlanProfile.SSIDConfig.SSID.hex
    }
    if ($wlanProfile.SSIDConfig -and $wlanProfile.SSIDConfig.nonBroadcast) {
        $nonBroadcast = ($wlanProfile.SSIDConfig.nonBroadcast -eq 'true')
    }

    $connectionType = $wlanProfile.connectionType
    $connectionMode = $wlanProfile.connectionMode
    $autoSwitch     = ($wlanProfile.autoSwitch -eq 'true')

    $auth         = $null
    $encryption   = $null
    $useOneX      = $false
    $keyType      = $null
    $keyProtected = $null
    $keyMaterial  = $null
    if ($wlanProfile.MSM -and $wlanProfile.MSM.security) {
        $sec = $wlanProfile.MSM.security
        if ($sec.authEncryption) {
            $auth       = $sec.authEncryption.authentication
            $encryption = $sec.authEncryption.encryption
            $useOneX    = ($sec.authEncryption.useOneX -eq 'true')
        }
        if ($sec.sharedKey) {
            $keyType      = $sec.sharedKey.keyType
            $keyProtected = ($sec.sharedKey.protected -eq 'true')
            $keyMaterial  = $sec.sharedKey.keyMaterial
        }
    }

    $macRand = $null
    # MacRandomization sits in a v3 child namespace. .NET's XmlElement returns
    # the property if present regardless of namespace, but if it's missing the
    # accessor returns $null in PS5.1 (no exception).
    if ($wlanProfile.MacRandomization) {
        $macRand = ($wlanProfile.MacRandomization.enableRandomization -eq 'true')
    }

    return [PSCustomObject]@{
        Name             = $name
        SsidName         = $ssidName
        SsidHex          = $ssidHex
        NonBroadcast     = $nonBroadcast
        ConnectionType   = $connectionType
        ConnectionMode   = $connectionMode
        AutoSwitch       = $autoSwitch
        Authentication   = $auth
        Encryption       = $encryption
        UseOneX          = $useOneX
        KeyType          = $keyType
        KeyProtected     = $keyProtected
        KeyMaterial      = $keyMaterial
        MacRandomization = $macRand
    }
}

function Get-WlanProfiles {
    $names = Get-WlanProfileList
    if (-not $names -or $names.Count -eq 0) { return @() }

    $rootTemp = Join-Path ([System.IO.Path]::GetTempPath()) ("TK-WISP-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $rootTemp -Force -ErrorAction SilentlyContinue

    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($n in $names) {
            $sub = Join-Path $rootTemp ([guid]::NewGuid().ToString('N'))
            $null = New-Item -ItemType Directory -Path $sub -Force -ErrorAction SilentlyContinue
            $xmlPath = Export-WlanProfileXml -Name $n -Folder $sub
            if ($null -eq $xmlPath) {
                Write-TKError -ScriptName 'wisp.ps1' -Message "Could not export profile '$n' (netsh produced no XML)." -Category 'Wi-Fi'
                continue
            }
            $row = ConvertFrom-WlanProfileXml -XmlPath $xmlPath
            if ($row) { $rows.Add($row) }
        }
    } finally {
        Remove-Item -Path $rootTemp -Recurse -Force -ErrorAction SilentlyContinue
    }
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Get-AuthTier {
    param([string]$Auth)
    if ([string]::IsNullOrWhiteSpace($Auth)) { return 'Unknown' }
    if ($AuthStrength.ContainsKey($Auth)) { return $AuthStrength[$Auth] }
    return 'Unknown'
}

function Get-CipherTier {
    param([string]$Cipher)
    if ([string]::IsNullOrWhiteSpace($Cipher)) { return 'Unknown' }
    if ($CipherStrength.ContainsKey($Cipher)) { return $CipherStrength[$Cipher] }
    return 'Unknown'
}

function Test-IsOpenProfile {
    param($P)
    $a = "$($P.Authentication)".ToLower()
    return ($a -eq 'open' -and ("$($P.Encryption)".ToLower() -in @('none', '')))
}

function Test-IsWeakProfile {
    param($P)
    if (Test-IsOpenProfile $P) { return $true }
    $auth = "$($P.Authentication)"
    $cipher = "$($P.Encryption)"
    if ((Get-AuthTier $auth) -eq 'Insecure') { return $true }
    if ((Get-CipherTier $cipher) -eq 'Insecure') { return $true }
    if ((Get-AuthTier $auth) -eq 'Weak')   { return $true }
    if ((Get-CipherTier $cipher) -eq 'Weak') { return $true }
    return $false
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-WispVerdict {
    param([array]$Adapters, [array]$Profiles)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if ($Adapters.Count -eq 0) {
        # Desktops without Wi-Fi are a normal state, not a finding.
        $warns.Add('No Wi-Fi adapter present on this machine -- desktop / wired-only configuration.')
    }

    foreach ($p in $Profiles) {
        $auth   = "$($p.Authentication)"
        $cipher = "$($p.Encryption)"
        $auto   = ($p.ConnectionMode -eq 'auto')

        if (Test-IsOpenProfile $p) {
            if ($auto) {
                $issues.Add("Open profile '$($p.Name)' is set to AUTO-CONNECT -- this machine will silently associate to any AP advertising that SSID.")
            } else {
                $warns.Add("Open profile '$($p.Name)' is configured (manual-connect).")
            }
        }
        if ((Get-CipherTier $cipher) -eq 'Insecure' -and $cipher -ne 'none') {
            $issues.Add("Profile '$($p.Name)' uses deprecated $cipher encryption -- WEP is broken; remove the profile.")
        }
        if ((Get-AuthTier $auth) -eq 'Weak') {
            $warns.Add("Profile '$($p.Name)' uses legacy $auth -- upgrade the network to WPA2-PSK / WPA3-SAE.")
        }
        if ((Get-CipherTier $cipher) -eq 'Weak') {
            $warns.Add("Profile '$($p.Name)' uses TKIP cipher -- upgrade to AES (CCMP) or GCMP.")
        }
        if ($p.AutoSwitch -and $auto) {
            $warns.Add("Profile '$($p.Name)' has autoSwitch enabled -- machine may roam between this and other known networks without prompt.")
        }
        if ($p.NonBroadcast -and $auto) {
            $warns.Add("Hidden-SSID profile '$($p.Name)' is auto-connecting -- the client probes for it constantly, leaking the SSID.")
        }
        if ($p.MacRandomization -eq $false) {
            $warns.Add("Profile '$($p.Name)' has MAC randomization disabled.")
        }
    }

    if ($Profiles.Count -gt 25) {
        $warns.Add("$($Profiles.Count) saved Wi-Fi profiles -- a large profile estate increases roaming surprises and probe-request leakage. Consider periodic cleanup.")
    }

    $verdict = if ($issues.Count -gt 0) { 'AT RISK' }
               elseif ($warns.Count -gt 0) { 'ATTENTION NEEDED' }
               else { 'CLEAN' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $class
        Issues  = @($issues)
        Warns   = @($warns)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-WispReport {
    param([array]$Adapters, [array]$Profiles, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _tierBadge { param([string]$Tier)
        switch ($Tier) {
            'Strong'   { "<span class='tk-badge-ok'>$(EscHtml $Tier)</span>" }
            'Weak'     { "<span class='tk-badge-warn'>$(EscHtml $Tier)</span>" }
            'Insecure' { "<span class='tk-badge-err'>$(EscHtml $Tier)</span>" }
            default    { "<span class='tk-badge-info'>$(EscHtml $Tier)</span>" }
        }
    }
    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-info'>No</span>" } }
    function _ynWarn { param($b) if ($b) { "<span class='tk-badge-warn'>Yes</span>" } else { "<span class='tk-badge-ok'>No</span>" } }

    function _keyCell { param($P)
        if ([string]::IsNullOrEmpty($P.KeyMaterial)) { return "<span class='tk-badge-info'>None / enterprise</span>" }
        if ($IncludeKey) { return "<code class='tk-mono'>$(EscHtml $P.KeyMaterial)</code>" }
        return "<span class='tk-badge-warn'>masked (use -IncludeKey to show)</span>"
    }

    # Adapters table
    $adRows = [System.Text.StringBuilder]::new()
    if ($Adapters.Count -eq 0) {
        [void]$adRows.Append("<tr><td colspan='6' class='tk-badge-info' style='text-align:center;'>No Wi-Fi adapters detected.</td></tr>")
    } else {
        foreach ($a in $Adapters) {
            $statBadge = if ($a.Status -eq 'Up') { "<span class='tk-badge-ok'>Up</span>" } else { "<span class='tk-badge-warn'>$(EscHtml $a.Status)</span>" }
            [void]$adRows.Append(
                "<tr><td>$(EscHtml $a.Name)</td><td>$(EscHtml $a.InterfaceDescription)</td>" +
                "<td>$statBadge</td><td>$(EscHtml $a.LinkSpeed)</td>" +
                "<td class='tk-mono'>$(EscHtml $a.MacAddress)</td>" +
                "<td>$(EscHtml $a.DriverVersion) ($(EscHtml $a.DriverDate))</td></tr>"
            )
        }
    }

    # Profile table
    $pRows = [System.Text.StringBuilder]::new()
    if ($Profiles.Count -eq 0) {
        [void]$pRows.Append("<tr><td colspan='9' class='tk-badge-info' style='text-align:center;'>No saved Wi-Fi profiles found.</td></tr>")
    } else {
        foreach ($p in ($Profiles | Sort-Object Name)) {
            $authBadge   = _tierBadge (Get-AuthTier $p.Authentication)
            $cipherBadge = _tierBadge (Get-CipherTier $p.Encryption)
            $modeBadge   = if ($p.ConnectionMode -eq 'auto') { "<span class='tk-badge-warn'>auto</span>" } else { "<span class='tk-badge-info'>$(EscHtml $p.ConnectionMode)</span>" }
            [void]$pRows.Append(
                "<tr><td>$(EscHtml $p.Name)</td>" +
                "<td>$(EscHtml $p.Authentication) $authBadge</td>" +
                "<td>$(EscHtml $p.Encryption) $cipherBadge</td>" +
                "<td>$modeBadge</td>" +
                "<td>$(_ynWarn $p.AutoSwitch)</td>" +
                "<td>$(_ynWarn $p.NonBroadcast)</td>" +
                "<td>$(if ($null -eq $p.MacRandomization) { '<span class=''tk-badge-info''>n/a</span>' } else { _yn $p.MacRandomization })</td>" +
                "<td>$(_yn $p.UseOneX)</td>" +
                "<td>$(_keyCell $p)</td></tr>"
            )
        }
    }

    # Open / weak profiles
    $weak = @($Profiles | Where-Object { Test-IsWeakProfile $_ })
    $weakRows = [System.Text.StringBuilder]::new()
    if ($weak.Count -eq 0) {
        [void]$weakRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No open or weak-cipher profiles detected.</td></tr>")
    } else {
        foreach ($p in $weak | Sort-Object Name) {
            $reason = @()
            if (Test-IsOpenProfile $p)             { $reason += 'open authentication' }
            if ((Get-AuthTier $p.Authentication) -eq 'Insecure' -and "$($p.Authentication)".ToLower() -ne 'open') { $reason += "insecure auth $($p.Authentication)" }
            if ((Get-AuthTier $p.Authentication) -eq 'Weak')      { $reason += "weak auth $($p.Authentication)" }
            if ((Get-CipherTier $p.Encryption) -eq 'Insecure')    { $reason += "insecure cipher $($p.Encryption)" }
            if ((Get-CipherTier $p.Encryption) -eq 'Weak')        { $reason += "weak cipher $($p.Encryption)" }
            $reasonStr = ($reason -join ', ')
            [void]$weakRows.Append(
                "<tr><td>$(EscHtml $p.Name)</td><td>$(EscHtml $p.Authentication)</td>" +
                "<td>$(EscHtml $p.Encryption)</td><td>$(EscHtml $reasonStr)</td></tr>"
            )
        }
    }

    # Auto-connect profiles
    $auto = @($Profiles | Where-Object { $_.ConnectionMode -eq 'auto' })
    $autoRows = [System.Text.StringBuilder]::new()
    if ($auto.Count -eq 0) {
        [void]$autoRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No auto-connect profiles configured.</td></tr>")
    } else {
        foreach ($p in $auto | Sort-Object Name) {
            $rowClass = if (Test-IsOpenProfile $p) { 'tk-badge-err' }
                        elseif (Test-IsWeakProfile $p) { 'tk-badge-warn' }
                        else { 'tk-badge-info' }
            [void]$autoRows.Append(
                "<tr><td>$(EscHtml $p.Name)</td><td>$(EscHtml $p.Authentication)</td>" +
                "<td>$(EscHtml $p.Encryption)</td><td><span class='$rowClass'>$(if (Test-IsOpenProfile $p) { 'open auto-connect' } elseif (Test-IsWeakProfile $p) { 'weak auto-connect' } else { 'auto-connect' })</span></td></tr>"
            )
        }
    }

    $findingsList = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$findingsList.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>") }
    foreach ($w in $Verdict.Warns)  { [void]$findingsList.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$findingsList.Append("<li class='tk-badge-ok'>No insecure or surprising Wi-Fi profiles found.</li>")
    }

    $strongCount = @($Profiles | Where-Object { (Get-AuthTier $_.Authentication) -eq 'Strong' -and (Get-CipherTier $_.Encryption) -eq 'Strong' }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'W.I.S.P. Wi-Fi Profile Audit' `
        -ScriptName 'W.I.S.P.' `
        -Subtitle   "${orgPrefix}Wi-Fi Profile Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'        = $machine
            'Generated'      = $reportDate
            'Verdict'        = $Verdict.Verdict
            'Saved Profiles' = $Profiles.Count
            'Adapters'       = $Adapters.Count
        }) `
        -NavItems   @('Verdict', 'Adapters', 'Saved Profiles', 'Open / Weak', 'Auto-connect')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'W.I.S.P. v6.0'

    $keyNoteBadge = if ($IncludeKey) {
        "<span class='tk-badge-warn'>Key material rendered in cleartext (-IncludeKey)</span>"
    } else {
        "<span class='tk-badge-ok'>Key material masked (default)</span>"
    }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Wi-Fi Posture</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Adapters.Count)</div><div class="tk-summary-lbl">Wi-Fi Adapter(s)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Profiles.Count)</div><div class="tk-summary-lbl">Saved Profiles</div></div>
    <div class="tk-summary-card $(if ($weak.Count -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($weak.Count)</div><div class="tk-summary-lbl">Open / Weak</div></div>
    <div class="tk-summary-card $(if ($auto.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($auto.Count)</div><div class="tk-summary-lbl">Auto-connect</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$strongCount</div><div class="tk-summary-lbl">Strong (WPA2/WPA3 + AES)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card">
      <ul class="tk-info-box" style="list-style:none;padding-left:0;">$($findingsList.ToString())</ul>
      <div class="tk-info-label">Privacy</div>
      <div class="tk-info-box">$keyNoteBadge</div>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Wi-Fi Adapters</span><span class="tk-section-num">$($Adapters.Count) detected</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Name</th><th>Description</th><th>Status</th><th>Link</th><th>MAC</th><th>Driver</th></tr></thead>
        <tbody>$($adRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Saved Profiles</span><span class="tk-section-num">$($Profiles.Count) total</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Authentication</th><th>Cipher</th><th>Connection</th><th>AutoSwitch</th><th>Hidden</th><th>MAC random</th><th>802.1X</th><th>Key</th></tr></thead>
        <tbody>$($pRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Open / Weak Profiles</span><span class="tk-section-num">$($weak.Count) flagged</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Authentication</th><th>Cipher</th><th>Reason</th></tr></thead>
        <tbody>$($weakRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Auto-connecting Profiles</span><span class="tk-section-num">$($auto.Count) auto-connect</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Authentication</th><th>Cipher</th><th>Risk Tier</th></tr></thead>
        <tbody>$($autoRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-WispBanner

Write-Section "WI-FI ADAPTERS"
$adapters = Get-WifiAdapters
if ($adapters.Count -eq 0) {
    Write-Info "No Wi-Fi adapters detected (desktop or wired-only configuration)."
} else {
    foreach ($a in $adapters) {
        $color = if ($a.Status -eq 'Up') { $C.Success } else { $C.Warning }
        Write-Host ("  {0,-12} {1,-8} {2}" -f $a.Name, $a.Status, $a.InterfaceDescription) -ForegroundColor $color
    }
}
Write-Host ""

Write-Section "SAVED PROFILES"
Write-Step "Exporting WLAN profiles to XML and parsing..."
$profiles = Get-WlanProfiles
Write-Host ("  Profiles parsed      : {0}" -f $profiles.Count) -ForegroundColor $C.Info
$openCount = @($profiles | Where-Object { Test-IsOpenProfile $_ }).Count
$weakCount = @($profiles | Where-Object { Test-IsWeakProfile $_ }).Count
Write-Host ("  Open profiles        : {0}" -f $openCount) -ForegroundColor $(if ($openCount -gt 0) { $C.Error } else { $C.Success })
Write-Host ("  Open / weak profiles : {0}" -f $weakCount) -ForegroundColor $(if ($weakCount -gt 0) { $C.Warning } else { $C.Success })
$autoCount = @($profiles | Where-Object { $_.ConnectionMode -eq 'auto' }).Count
Write-Host ("  Auto-connect         : {0}" -f $autoCount) -ForegroundColor $C.Info
Write-Host ""

$verdict = Get-WispVerdict -Adapters $adapters -Profiles $profiles

Write-Section "WI-FI VERDICT"
$verdictColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $verdictColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-WispReport -Adapters $adapters -Profiles $profiles -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "WISP_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
