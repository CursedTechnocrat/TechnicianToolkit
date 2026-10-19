# cumulus.ps1 - C.U.M.U.L.U.S. — Catalogs Usage, Members, Unowned, Links, Untouched & Shared sites (SharePoint)
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
    C.U.M.U.L.U.S. — Catalogs Usage, Members, Unowned, Links, Untouched & Shared sites (SharePoint)
    SharePoint Online Site Inventory Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft Graph and inventories the SharePoint Online
    estate: every site collection with storage used, file count, last
    activity date, external-sharing state, and owner metadata. Uses the
    Graph usage-report endpoint (getSharePointSiteUsageDetail) so the
    audit is a single bulk read rather than N per-site queries. Flags
    storage-heavy sites, externally shared sites, owner-less sites, and
    stale sites with no recent activity.

.USAGE
    PS C:\> .\cumulus.ps1                    # Interactive menu
    PS C:\> .\cumulus.ps1 -Unattended        # Silent: auto-connect + export HTML

.NOTES
    Version : 6.2

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
    'Sites.Read.All',
    'Reports.Read.All',
    'Directory.Read.All'
)

function Show-CumulusBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  C.U.M.U.L.U.S. — Catalogs Usage, Members, Unowned, Links, Untouched & Shared sites" -ForegroundColor Cyan
    Write-Host "  SharePoint Online Site Inventory Tool  v6.2" -ForegroundColor Cyan
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
        'Microsoft.Graph.Sites',
        'Microsoft.Graph.Reports',
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
                Write-TKError -ScriptName 'cumulus' -Message "Microsoft.Graph install failed: $($_.Exception.Message)" -Category 'Module Install'
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
                # Org lookup is best-effort metadata.
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
        Write-TKError -ScriptName 'cumulus' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        if ($Unattended) { exit 1 }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SITE INVENTORY VIA USAGE REPORT
# ─────────────────────────────────────────────────────────────────────────────

# The getSharePointSiteUsageDetail report returns one CSV row per site with
# storage, file count, last activity date, owner, and external-share flag --
# everything we need for a fleet-level audit in a single bulk read. Beats the
# alternative of iterating Get-MgSite per URL, which rate-limits on large
# tenants.
function Get-SharePointInventory {
    Write-Section "SHAREPOINT SITE INVENTORY"
    Write-Step "Pulling SharePoint usage report (period=D30)..."

    $uri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')"
    $csv = $null
    try {
        # Invoke-MgGraphRequest returns the CSV as a string for report endpoints
        # when no output file is specified. Explicit -OutputType is required to
        # avoid the SDK auto-wrapping.
        $csv = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType HttpResponseMessage -ErrorAction Stop
        if ($csv -is [System.Net.Http.HttpResponseMessage]) {
            $csv = $csv.Content.ReadAsStringAsync().Result
        }
    } catch {
        Write-Fail "Usage report fetch failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'cumulus' -Message "getSharePointSiteUsageDetail failed: $($_.Exception.Message)" -Category 'Graph Query'
        return @()
    }

    if (-not $csv) {
        Write-Warn "Usage report returned empty content."
        return @()
    }

    # The first row is the header. Parse with ConvertFrom-Csv directly on the
    # multi-line string; the endpoint emits a UTF-8 CSV with a fixed column set.
    $rows = $null
    try {
        $rows = $csv | ConvertFrom-Csv
    } catch {
        Write-Fail "Could not parse CSV response: $($_.Exception.Message)"
        return @()
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warn "Usage report has no rows — tenant may have no SharePoint sites, or report data is still aggregating."
        return @()
    }

    $records = foreach ($r in $rows) {
        # Column names include "Storage Used (Byte)" etc. with spaces and units,
        # so we reach for them by name and coerce defensively.
        $storageBytes = 0
        $files        = 0
        $activeFiles  = 0
        try { $storageBytes = [int64]$r.'Storage Used (Byte)' } catch { }
        try { $files        = [int]$r.'File Count' } catch { }
        try { $activeFiles  = [int]$r.'Active File Count' } catch { }

        $lastActivity = $null
        if ($r.'Last Activity Date') {
            try { $lastActivity = [datetime]$r.'Last Activity Date' } catch { }
        }

        [PSCustomObject]@{
            SiteId            = $r.'Site Id'
            SiteUrl           = $r.'Site URL'
            OwnerDisplayName  = $r.'Owner Display Name'
            OwnerPrincipal    = $r.'Owner Principal Name'
            SiteType          = $r.'Root Web Template'
            StorageBytes      = $storageBytes
            StorageSize       = Format-Bytes $storageBytes
            FileCount         = $files
            ActiveFileCount   = $activeFiles
            LastActivity      = $lastActivity
            LastActivityStr   = if ($lastActivity) { $lastActivity.ToString('yyyy-MM-dd') } else { 'Never' }
            ExternalSharing   = ($r.'External Sharing' -eq 'True')
        }
    }

    $records = @($records)
    Write-Ok "Collected $($records.Count) site(s) from the usage report."
    Write-Host ""
    return $records
}

# Tenant-wide SharePoint sharing policy. Emits a single record with the three
# top-level toggles that govern every site in the tenant.
function Get-TenantSharingPolicy {
    Write-Step "Reading tenant sharing policy (admin/sharepoint/settings)..."
    $uri = 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'

    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        return [PSCustomObject]@{
            SharingCapability             = $resp.sharingCapability
            SharingDomainRestrictionMode  = $resp.sharingDomainRestrictionMode
            DefaultSharingLinkType        = $resp.defaultSharingLinkType
            DefaultLinkPermission         = $resp.defaultLinkPermission
            AvailableManagedPathsForSiteCreation = $resp.availableManagedPathsForSiteCreation
            CollectorError                = $null
        }
    } catch {
        return [PSCustomObject]@{
            SharingCapability            = 'unknown'
            SharingDomainRestrictionMode = 'unknown'
            DefaultSharingLinkType       = 'unknown'
            DefaultLinkPermission        = 'unknown'
            CollectorError               = $_.Exception.Message
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# DERIVED FINDINGS
# ─────────────────────────────────────────────────────────────────────────────

$script:LargeSiteThresholdBytes = 100GB
$script:StaleActivityDays       = 180

function Get-LargeSites   { param([array]$S) @($S | Where-Object { $_.StorageBytes -ge $script:LargeSiteThresholdBytes } | Sort-Object StorageBytes -Descending) }
function Get-SharedSites  { param([array]$S) @($S | Where-Object { $_.ExternalSharing } | Sort-Object SiteUrl) }
function Get-OwnerlessSites { param([array]$S) @($S | Where-Object { -not $_.OwnerDisplayName -or $_.OwnerDisplayName -eq '' } | Sort-Object SiteUrl) }
function Get-StaleSites {
    param([array]$S)
    $threshold = (Get-Date).AddDays(-$script:StaleActivityDays)
    @($S | Where-Object { -not $_.LastActivity -or $_.LastActivity -lt $threshold } | Sort-Object LastActivityStr)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Sites, $Policy, [array]$Large, [array]$Shared, [array]$Ownerless, [array]$Stale)

    $reportDate    = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $tkCfg         = Get-TKConfig
    $orgPrefix     = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalBytes = ($Sites | Measure-Object -Property StorageBytes -Sum).Sum
    $totalFiles = ($Sites | Measure-Object -Property FileCount    -Sum).Sum

    function _site-row {
        param($s)
        $sizeClass = if ($s.StorageBytes -ge $script:LargeSiteThresholdBytes) { 'tk-badge-warn' }
                     elseif ($s.StorageBytes -ge 10GB) { 'tk-badge-info' }
                     else { 'tk-badge-ok' }
        $shareBadge = if ($s.ExternalSharing) { "<span class='tk-badge-warn'>Enabled</span>" } else { "<span class='tk-badge-ok'>Disabled</span>" }
        $ownerCell  = if ($s.OwnerDisplayName) { EscHtml $s.OwnerDisplayName } else { "<span class='tk-badge-err'>(no owner)</span>" }
        $activeClass = if (-not $s.LastActivity) { 'tk-badge-warn' }
                       elseif ($s.LastActivity -lt (Get-Date).AddDays(-$script:StaleActivityDays)) { 'tk-badge-warn' }
                       else { 'tk-badge-ok' }
        return "<tr><td><code>$(EscHtml $s.SiteUrl)</code></td><td>$(EscHtml $s.SiteType)</td><td>$ownerCell</td><td><span class='$sizeClass'>$(EscHtml $s.StorageSize)</span></td><td>$($s.FileCount)</td><td><span class='$activeClass'>$(EscHtml $s.LastActivityStr)</span></td><td>$shareBadge</td></tr>"
    }

    $invRows = [System.Text.StringBuilder]::new()
    if ($Sites.Count -eq 0) {
        [void]$invRows.Append("<tr><td colspan='7' class='tk-badge-info' style='text-align:center;'>No SharePoint sites returned by the usage report.</td></tr>")
    } else {
        foreach ($s in ($Sites | Sort-Object StorageBytes -Descending)) { [void]$invRows.Append((_site-row $s) + "`n") }
    }

    function _category-rows {
        param([array]$Rows, [int]$Cols, [string]$EmptyMessage)
        $sb = [System.Text.StringBuilder]::new()
        if ($Rows.Count -eq 0) {
            [void]$sb.Append("<tr><td colspan='$Cols' class='tk-badge-ok' style='text-align:center;'>$EmptyMessage</td></tr>")
        } else {
            foreach ($s in $Rows) { [void]$sb.Append((_site-row $s) + "`n") }
        }
        return $sb.ToString()
    }

    $largeRows     = _category-rows -Rows $Large     -Cols 7 -EmptyMessage "No sites exceed $(Format-Bytes $script:LargeSiteThresholdBytes)."
    $sharedRows    = _category-rows -Rows $Shared    -Cols 7 -EmptyMessage 'No sites have external sharing enabled.'
    $ownerlessRows = _category-rows -Rows $Ownerless -Cols 7 -EmptyMessage 'No sites without an owner.'
    $staleRows     = _category-rows -Rows $Stale     -Cols 7 -EmptyMessage "No sites with activity older than $script:StaleActivityDays days."

    # Tenant sharing policy table
    $policyClass = switch ($Policy.SharingCapability) {
        'disabled'                     { 'tk-badge-ok' }
        'externalUserSharingOnly'      { 'tk-badge-warn' }
        'externalUserAndGuestSharing'  { 'tk-badge-warn' }
        'existingExternalUserSharingOnly' { 'tk-badge-warn' }
        default                        { 'tk-badge-info' }
    }

    $policyErr = if ($Policy.CollectorError) {
        "<tr><th>Policy Read Error</th><td><span class='tk-badge-warn'>$(EscHtml $Policy.CollectorError)</span></td></tr>"
    } else { '' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.U.M.U.L.U.S. SharePoint Inventory Report' `
        -ScriptName 'C.U.M.U.L.U.S.' `
        -Subtitle   "${orgPrefix}SharePoint Online Estate Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = EscHtml $script:ConnectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @('Sharing Policy', 'Site Inventory', 'Large Sites', 'External Sharing', 'Ownerless Sites', 'Stale Sites')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'C.U.M.U.L.U.S. v6.2'

    $largeClass     = if ($Large.Count -gt 0)     { 'warn' } else { 'ok' }
    $sharedClass    = if ($Shared.Count -gt 0)    { 'warn' } else { 'ok' }
    $ownerlessClass = if ($Ownerless.Count -gt 0) { 'err'  } else { 'ok' }
    $staleClass     = if ($Stale.Count -gt 0)     { 'warn' } else { 'ok' }

    $invHeader = "<tr><th>Site URL</th><th>Template</th><th>Owner</th><th>Storage</th><th>Files</th><th>Last Activity</th><th>External Sharing</th></tr>"

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Sites.Count)</div><div class="tk-summary-lbl">Sites Total</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(Format-Bytes $totalBytes)</div><div class="tk-summary-lbl">Total Storage</div></div>
    <div class="tk-summary-card $largeClass"><div class="tk-summary-num">$($Large.Count)</div><div class="tk-summary-lbl">Large Sites (&ge; $(Format-Bytes $script:LargeSiteThresholdBytes))</div></div>
    <div class="tk-summary-card $sharedClass"><div class="tk-summary-num">$($Shared.Count)</div><div class="tk-summary-lbl">External Sharing On</div></div>
    <div class="tk-summary-card $ownerlessClass"><div class="tk-summary-num">$($Ownerless.Count)</div><div class="tk-summary-lbl">Ownerless</div></div>
    <div class="tk-summary-card $staleClass"><div class="tk-summary-num">$($Stale.Count)</div><div class="tk-summary-lbl">Stale (&gt; $script:StaleActivityDays d)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Tenant Sharing Policy</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Sharing Capability</th><td><span class='$policyClass'>$(EscHtml $Policy.SharingCapability)</span></td></tr>
          <tr><th>Domain Restriction Mode</th><td>$(EscHtml $Policy.SharingDomainRestrictionMode)</td></tr>
          <tr><th>Default Link Type</th><td>$(EscHtml $Policy.DefaultSharingLinkType)</td></tr>
          <tr><th>Default Link Permission</th><td>$(EscHtml $Policy.DefaultLinkPermission)</td></tr>
          $policyErr
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Site Inventory</span><span class="tk-section-num">$($Sites.Count) site(s)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$($invRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Large Sites -- &ge; $(Format-Bytes $script:LargeSiteThresholdBytes)</span><span class="tk-section-num">$($Large.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$largeRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Sites With External Sharing Enabled</span><span class="tk-section-num">$($Shared.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$sharedRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Ownerless Sites</span><span class="tk-section-num">$($Ownerless.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$ownerlessRows</tbody></table></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Stale Sites -- No Activity &gt; $script:StaleActivityDays Days</span><span class="tk-section-num">$($Stale.Count)</span></div>
    <div class="tk-card"><table class="tk-table"><thead>$invHeader</thead><tbody>$staleRows</tbody></table></div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param([array]$Sites, $Policy, [array]$Large, [array]$Shared, [array]$Ownerless, [array]$Stale)

    Write-Section "EXPORTING HTML REPORT"
    $html      = Build-HtmlReport -Sites $Sites -Policy $Policy -Large $Large -Shared $Shared -Ownerless $Ownerless -Stale $Stale
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "CUMULUS_${timestamp}.html"

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
    Show-CumulusBanner
    $connStatus = if ($script:Connected) { "  Connected as : $($script:ConnectedAs)" } else { "  Not Connected — select option 1 to authenticate" }
    $connColor  = if ($script:Connected) { $C.Success } else { $C.Warning }
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  SharePoint inventory (refresh underlying data)" -ForegroundColor $C.Info
    Write-Host "  3  Export full HTML report (all findings)" -ForegroundColor $C.Success
    Write-Host "  Q  Quit" -ForegroundColor $C.Info
    Write-Host ""
}

function Assert-Connected {
    if (-not $script:Connected) {
        Write-Host ""; Write-Warn "Not connected. Please select option 1 first."; Write-Host ""
        Read-Host "  Press Enter to return to menu"
        return $false
    }
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

Show-CumulusBanner
Install-GraphModule

if (Test-GraphConnection) { Write-Ok "Already connected as $($script:ConnectedAs)"; Write-Host "" }

if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    if (-not $script:Connected) { Invoke-Connect }
    if (-not $script:Connected) { Write-Fail "Could not connect to Graph. Exiting."; exit 1 }

    $sites  = Get-SharePointInventory
    $policy = Get-TenantSharingPolicy
    $large  = Get-LargeSites      -S $sites
    $shared = Get-SharedSites     -S $sites
    $owner  = Get-OwnerlessSites  -S $sites
    $stale  = Get-StaleSites      -S $sites

    Export-HtmlReport -Sites $sites -Policy $policy -Large $large -Shared $shared -Ownerless $owner -Stale $stale

    Write-Host ""
    Write-Ok "Unattended audit complete."
    if ($Transcript) { Stop-TKTranscript }
    exit 0
}

$cachedSites  = $null
$cachedPolicy = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {
        '1' {
            Invoke-Connect
            $cachedSites = $null; $cachedPolicy = $null
            Read-Host "  Press Enter to return to menu"
        }
        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedSites  = Get-SharePointInventory
            $cachedPolicy = Get-TenantSharingPolicy
            Read-Host "  Press Enter to return to menu"
        }
        '3' {
            if (-not (Assert-Connected)) { break }
            if ($null -eq $cachedSites)  { $cachedSites  = Get-SharePointInventory }
            if ($null -eq $cachedPolicy) { $cachedPolicy = Get-TenantSharingPolicy }
            $large  = Get-LargeSites      -S $cachedSites
            $shared = Get-SharedSites     -S $cachedSites
            $owner  = Get-OwnerlessSites  -S $cachedSites
            $stale  = Get-StaleSites      -S $cachedSites
            Export-HtmlReport -Sites $cachedSites -Policy $cachedPolicy -Large $large -Shared $shared -Ownerless $owner -Stale $stale
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
