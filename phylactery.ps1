# phylactery.ps1 - P.H.Y.L.A.C.T.E.R.Y. — Pre-migration Health of Your Libraries: Accounts, Client, Tethering, Errors, Readiness & Yield
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
    P.H.Y.L.A.C.T.E.R.Y. — Pre-migration Health of Your Libraries: Accounts, Client, Tethering, Errors, Readiness & Yield
    OneDrive Known-Folder-Move Pre-Migration Validator for PowerShell 5.1+

.DESCRIPTION
    Local machine audit that answers the pre-migration question: "Is this
    user's data actually going to be in the cloud when we hand them a new
    laptop?" Checks OneDrive client state, signed-in accounts, Known
    Folder Move engagement for Desktop / Documents / Pictures, content
    volumes, and recent sync errors. Produces a dark-themed HTML report
    with a red / yellow / green readiness verdict.

.USAGE
    PS C:\> .\phylactery.ps1                    # Interactive run against current user
    PS C:\> .\phylactery.ps1 -Unattended        # Silent mode, export HTML and exit

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

function Show-PhylacteryBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  P.H.Y.L.A.C.T.E.R.Y. — Pre-migration Health of Your Libraries: Accounts, Client, Tethering, Errors, Readiness & Yield" -ForegroundColor Cyan
    Write-Host "  OneDrive Known-Folder-Move Pre-Migration Validator  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Data collection and report-building functions are appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — ONEDRIVE CLIENT STATE
# ─────────────────────────────────────────────────────────────────────────────

function Get-OneDriveClient {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        'C:\Program Files\Microsoft OneDrive\OneDrive.exe',
        'C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe'
    )

    $exe = $null
    foreach ($p in $candidates) {
        if (Test-Path $p) { $exe = $p; break }
    }

    $version = $null
    if ($exe) {
        try {
            $version = (Get-Item $exe).VersionInfo.ProductVersion
        } catch {
            $version = 'unknown'
        }
    }

    $running = $false
    try {
        $running = [bool](Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue)
    } catch {
        # Process enumeration may fail under constrained contexts — keep default $false.
    }

    return [PSCustomObject]@{
        Installed = [bool]$exe
        Path      = $exe
        Version   = $version
        Running   = $running
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — SIGNED-IN ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────

function Get-OneDriveAccounts {
    $root = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (-not (Test-Path $root)) { return @() }

    $accounts = foreach ($child in Get-ChildItem $root -ErrorAction SilentlyContinue) {
        $key = $child.PSPath
        $props = $null
        try {
            $props = Get-ItemProperty -Path $key -ErrorAction Stop
        } catch {
            continue
        }

        $accountType = if ($child.PSChildName -like 'Business*') { 'Business' }
                       elseif ($child.PSChildName -eq 'Personal') { 'Personal' }
                       else { 'Unknown' }

        [PSCustomObject]@{
            AccountKey  = $child.PSChildName
            AccountType = $accountType
            UserEmail   = $props.UserEmail
            UserFolder  = $props.UserFolder
            TenantId    = $props.ConfiguredTenantId
            DisplayName = $props.DisplayName
        }
    }

    return @($accounts)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — KNOWN FOLDER MOVE STATUS
# ─────────────────────────────────────────────────────────────────────────────

# Shell folder GUIDs (User Shell Folders registry uses named values, but the
# KFM check is the simpler "does the resolved path contain 'OneDrive'" rule).
$script:KfmFolders = @(
    @{ Label = 'Desktop';   ShellName = 'Desktop' },
    @{ Label = 'Documents'; ShellName = 'Personal' },
    @{ Label = 'Pictures';  ShellName = 'My Pictures' }
)

function Get-KfmStatus {
    $shellKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $results = foreach ($folder in $script:KfmFolders) {
        $path = $null
        try {
            $raw = (Get-ItemProperty -Path $shellKey -Name $folder.ShellName -ErrorAction Stop).($folder.ShellName)
            $path = [Environment]::ExpandEnvironmentVariables($raw)
        } catch {
            $path = $null
        }

        $inOneDrive = $false
        if ($path) {
            $inOneDrive = ($path -match '(?i)\\OneDrive') -or
                          ($path -like "$env:OneDrive*") -or
                          ($path -like "$env:OneDriveCommercial*")
        }

        [PSCustomObject]@{
            Label      = $folder.Label
            Path       = if ($path) { $path } else { '(not set)' }
            Redirected = $inOneDrive
        }
    }

    return @($results)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — FOLDER CONTENT VOLUME
# ─────────────────────────────────────────────────────────────────────────────

function Get-FolderVolume {
    param([array]$KfmStatus)

    $results = foreach ($f in $KfmStatus) {
        $count = 0
        $bytes = 0
        $errMsg = $null

        if ($f.Path -and (Test-Path -LiteralPath $f.Path)) {
            try {
                $items = Get-ChildItem -LiteralPath $f.Path -Recurse -File -Force -ErrorAction SilentlyContinue
                $count = @($items).Count
                $bytes = ($items | Measure-Object -Property Length -Sum).Sum
                if (-not $bytes) { $bytes = 0 }
            } catch {
                $errMsg = $_.Exception.Message
            }
        } else {
            $errMsg = 'Path does not exist'
        }

        [PSCustomObject]@{
            Label     = $f.Label
            Path      = $f.Path
            FileCount = $count
            Bytes     = $bytes
            Size      = Format-Bytes $bytes
            Error     = $errMsg
        }
    }

    return @($results)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — RECENT SYNC ERRORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-SyncErrors {
    $since = (Get-Date).AddDays(-7)
    $events = @()

    # OneDrive errors historically land in the Application log with provider
    # 'Microsoft-Windows-User Profiles Service' or the generic 'OneDrive' source.
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = 'Application'
            ProviderName = @('OneDrive', 'Microsoft-Windows-User Profiles Service')
            Level     = 1, 2, 3  # Critical, Error, Warning
            StartTime = $since
        } -ErrorAction SilentlyContinue
    } catch {
        # Filtering on non-existent providers throws — fall back to a broader sweep.
        try {
            $events = Get-WinEvent -LogName Application -MaxEvents 500 -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.TimeCreated -ge $since -and
                    $_.LevelDisplayName -in @('Error','Warning','Critical') -and
                    ($_.ProviderName -like '*OneDrive*' -or $_.Message -match 'OneDrive')
                }
        } catch {
            $events = @()
        }
    }

    $rows = foreach ($e in $events) {
        [PSCustomObject]@{
            Time     = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm')
            Level    = $e.LevelDisplayName
            Provider = $e.ProviderName
            EventId  = $e.Id
            Message  = ($e.Message -split "`n" | Select-Object -First 1).Trim()
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-ReadinessVerdict {
    param($Client, [array]$Accounts, [array]$Kfm, [array]$Volume, [array]$Errors)

    $issues  = [System.Collections.Generic.List[string]]::new()
    $warns   = [System.Collections.Generic.List[string]]::new()

    if (-not $Client.Installed) { $issues.Add('OneDrive client is not installed.') }
    elseif (-not $Client.Running) { $warns.Add('OneDrive client is installed but not currently running.') }

    $businessAccounts = @($Accounts | Where-Object { $_.AccountType -eq 'Business' -and $_.UserEmail })
    if ($businessAccounts.Count -eq 0) {
        $issues.Add('No Business/Work OneDrive account is signed in.')
    }

    $notRedirected = @($Kfm | Where-Object { -not $_.Redirected })
    foreach ($r in $notRedirected) {
        $issues.Add("Known Folder '$($r.Label)' is not redirected to OneDrive.")
    }

    foreach ($v in $Volume) {
        if ($v.Bytes -gt 25GB) {
            $warns.Add("Folder '$($v.Label)' is $($v.Size) — large uploads may take time to complete.")
        }
    }

    $errCount = @($Errors).Count
    if ($errCount -gt 0) {
        $warns.Add("$errCount OneDrive Application-log errors/warnings in the last 7 days.")
    }

    $verdict = if ($issues.Count -gt 0) { 'NOT READY' }
               elseif ($warns.Count -gt 0) { 'READY WITH WARNINGS' }
               else { 'READY' }
    $class = if ($issues.Count -gt 0) { 'err' }
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

function Build-HtmlReport {
    param($Client, [array]$Accounts, [array]$Kfm, [array]$Volume, [array]$Errors, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $runUser    = "$env:USERDOMAIN\$env:USERNAME"

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $clientRow = if ($Client.Installed) {
        $runBadge = if ($Client.Running) { "<span class='tk-badge-ok'>Running</span>" } else { "<span class='tk-badge-warn'>Not running</span>" }
        "<tr><td><span class='tk-badge-ok'>Installed</span></td><td>$(EscHtml $Client.Version)</td><td>$runBadge</td><td><code>$(EscHtml $Client.Path)</code></td></tr>"
    } else {
        "<tr><td colspan='4' class='tk-badge-err' style='text-align:center;'>OneDrive client is NOT installed on this machine.</td></tr>"
    }

    $acctRows = [System.Text.StringBuilder]::new()
    if ($Accounts.Count -eq 0) {
        [void]$acctRows.Append("<tr><td colspan='4' class='tk-badge-err' style='text-align:center;'>No OneDrive accounts signed in.</td></tr>")
    } else {
        foreach ($a in $Accounts) {
            $typeBadge = if ($a.AccountType -eq 'Business') { "<span class='tk-badge-ok'>Business</span>" }
                         elseif ($a.AccountType -eq 'Personal') { "<span class='tk-badge-warn'>Personal</span>" }
                         else { "<span class='tk-badge-info'>Unknown</span>" }
            [void]$acctRows.Append("<tr><td>$typeBadge</td><td>$(EscHtml $a.UserEmail)</td><td><code>$(EscHtml $a.UserFolder)</code></td><td>$(EscHtml $a.TenantId)</td></tr>`n")
        }
    }

    $kfmRows = [System.Text.StringBuilder]::new()
    foreach ($k in $Kfm) {
        $badge = if ($k.Redirected) { "<span class='tk-badge-ok'>Redirected</span>" } else { "<span class='tk-badge-err'>Local only</span>" }
        [void]$kfmRows.Append("<tr><td>$(EscHtml $k.Label)</td><td>$badge</td><td><code>$(EscHtml $k.Path)</code></td></tr>`n")
    }

    $volRows = [System.Text.StringBuilder]::new()
    foreach ($v in $Volume) {
        $sizeClass = if ($v.Bytes -ge 25GB) { 'tk-badge-err' }
                     elseif ($v.Bytes -ge 5GB) { 'tk-badge-warn' }
                     else { 'tk-badge-ok' }
        $errCell = if ($v.Error) { "<span class='tk-badge-warn'>$(EscHtml $v.Error)</span>" } else { '' }
        [void]$volRows.Append("<tr><td>$(EscHtml $v.Label)</td><td>$($v.FileCount)</td><td><span class='$sizeClass'>$(EscHtml $v.Size)</span></td><td>$errCell</td></tr>`n")
    }

    $errRows = [System.Text.StringBuilder]::new()
    if ($Errors.Count -eq 0) {
        [void]$errRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No OneDrive errors in the last 7 days.</td></tr>")
    } else {
        foreach ($e in $Errors | Select-Object -First 50) {
            $levelClass = if ($e.Level -eq 'Error' -or $e.Level -eq 'Critical') { 'tk-badge-err' } else { 'tk-badge-warn' }
            [void]$errRows.Append("<tr><td>$(EscHtml $e.Time)</td><td><span class='$levelClass'>$(EscHtml $e.Level)</span></td><td>$($e.EventId)</td><td>$(EscHtml $e.Message)</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>All pre-migration checks passed.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'P.H.Y.L.A.C.T.E.R.Y. OneDrive Pre-Migration Report' `
        -ScriptName 'P.H.Y.L.A.C.T.E.R.Y.' `
        -Subtitle   "${orgPrefix}OneDrive KFM Readiness -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Run As'    = $runUser
            'Generated' = $reportDate
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'Client', 'Accounts', 'Known Folders', 'Content Volume', 'Sync Errors')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'P.H.Y.L.A.C.T.E.R.Y. v6.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Migration Readiness</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Accounts.Count)</div><div class="tk-summary-lbl">OneDrive Accounts</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Kfm | Where-Object { $_.Redirected }).Count) / $($Kfm.Count)</div><div class="tk-summary-lbl">Folders Redirected</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(Format-Bytes (($Volume | Measure-Object -Property Bytes -Sum).Sum))</div><div class="tk-summary-lbl">Total Content</div></div>
    <div class="tk-summary-card $(if ($Errors.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Errors.Count)</div><div class="tk-summary-lbl">Sync Errors (7d)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Verdict &amp; Findings</span>
      <span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span>
    </div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">OneDrive Client</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Installed</th><th>Version</th><th>Process</th><th>Path</th></tr></thead>
        <tbody>$clientRow</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Signed-In Accounts</span><span class="tk-section-num">$($Accounts.Count) account(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Type</th><th>Email</th><th>Sync Root</th><th>Tenant ID</th></tr></thead>
        <tbody>$($acctRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Known Folder Redirection</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Folder</th><th>Status</th><th>Resolved Path</th></tr></thead>
        <tbody>$($kfmRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Content Volume</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Folder</th><th>File Count</th><th>Size</th><th>Error</th></tr></thead>
        <tbody>$($volRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Sync Errors (last 7 days)</span><span class="tk-section-num">$($Errors.Count) event(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Level</th><th>Event ID</th><th>Message</th></tr></thead>
        <tbody>$($errRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

function Write-ConsoleSummary {
    param($Client, [array]$Accounts, [array]$Kfm, [array]$Volume, [array]$Errors, $Verdict)

    Write-Section "ONEDRIVE CLIENT"
    if ($Client.Installed) {
        Write-Ok "Installed: $($Client.Path) (v$($Client.Version))"
        if ($Client.Running) { Write-Ok "Process is running." } else { Write-Warn "Process is not currently running." }
    } else {
        Write-Fail "OneDrive client is NOT installed."
    }

    Write-Section "SIGNED-IN ACCOUNTS"
    if ($Accounts.Count -eq 0) {
        Write-Fail "No OneDrive accounts are signed in."
    } else {
        foreach ($a in $Accounts) {
            $label = "[$($a.AccountType)] $($a.UserEmail)"
            Write-Host "  $label" -ForegroundColor $(if ($a.AccountType -eq 'Business') { $C.Success } else { $C.Warning })
            Write-Host "    Sync root : $($a.UserFolder)" -ForegroundColor $C.Info
        }
    }

    Write-Section "KNOWN FOLDER REDIRECTION"
    foreach ($k in $Kfm) {
        if ($k.Redirected) {
            Write-Ok "$($k.Label) -> redirected to OneDrive"
        } else {
            Write-Fail "$($k.Label) -> local only ($($k.Path))"
        }
    }

    Write-Section "CONTENT VOLUME"
    foreach ($v in $Volume) {
        $color = if ($v.Bytes -ge 25GB) { $C.Error } elseif ($v.Bytes -ge 5GB) { $C.Warning } else { $C.Success }
        Write-Host ("  {0,-12} {1,8} file(s)  {2}" -f $v.Label, $v.FileCount, $v.Size) -ForegroundColor $color
    }

    Write-Section "SYNC ERRORS (LAST 7 DAYS)"
    if ($Errors.Count -eq 0) {
        Write-Ok "No OneDrive errors in the last 7 days."
    } else {
        Write-Warn "$($Errors.Count) event(s) in the last 7 days."
        foreach ($e in $Errors | Select-Object -First 10) {
            $color = if ($e.Level -in 'Error','Critical') { $C.Error } else { $C.Warning }
            Write-Host ("  [{0}] {1} (id={2}): {3}" -f $e.Time, $e.Level, $e.EventId, $e.Message) -ForegroundColor $color
        }
    }

    Write-Section "MIGRATION READINESS VERDICT"
    $vColor = switch ($Verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
    Write-Host "  $($Verdict.Verdict)" -ForegroundColor $vColor
    foreach ($i in $Verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
    foreach ($w in $Verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-PhylacteryBanner

$client   = Get-OneDriveClient
$accounts = Get-OneDriveAccounts
$kfm      = Get-KfmStatus
$volume   = Get-FolderVolume -KfmStatus $kfm
$errors   = Get-SyncErrors
$verdict  = Get-ReadinessVerdict -Client $client -Accounts $accounts -Kfm $kfm -Volume $volume -Errors $errors

Write-ConsoleSummary -Client $client -Accounts $accounts -Kfm $kfm -Volume $volume -Errors $errors -Verdict $verdict

Write-Host ""
Write-Step "Generating HTML report..."
$html = Build-HtmlReport -Client $client -Accounts $accounts -Kfm $kfm -Volume $volume -Errors $errors -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "PHYLACTERY_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
