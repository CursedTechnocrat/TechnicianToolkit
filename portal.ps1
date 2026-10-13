# portal.ps1 - P.O.R.T.A.L. — Profiles, Observes & Reports Tunnels, Authentication & Links
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
    P.O.R.T.A.L. — Profiles, Observes & Reports Tunnels, Authentication & Links
    VPN / Always-On VPN Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits the VPN posture of a Windows endpoint. Enumerates built-in
    Windows VPN connections at user and all-user scope via
    `Get-VpnConnection`, surfaces Always-On VPN app triggers via
    `Get-VpnConnectionTrigger`, dumps the Name Resolution Policy Table
    (NRPT) via `Get-DnsClientNrptPolicy`, lists VPN tunnel interfaces,
    and detects installed third-party VPN clients via the service
    table (Cisco AnyConnect / Cisco Secure Client, Palo Alto
    GlobalProtect, Ivanti / Pulse, OpenVPN, WireGuard, Tailscale,
    ZeroTier, Cloudflare WARP, NordVPN, ProtonVPN).

    Read-only audit -- no state-changing actions are performed.

.USAGE
    PS C:\> .\portal.ps1                    # Interactive run
    PS C:\> .\portal.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 6.1

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

function Show-PortalBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  P.O.R.T.A.L. -- Profiles, Observes & Reports Tunnels, Authentication & Links" -ForegroundColor Green
    Write-Host "  VPN / Always-On VPN Audit  v6.1" -ForegroundColor Green
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Authentication-method strength tier. PAP transmits credentials in cleartext;
# CHAP and MS-CHAPv1 are deprecated. MS-CHAPv2 is acceptable for non-EAP VPNs.
# EAP and MachineCertificate are the strong tier.
$AuthStrength = @{
    'Pap'                = 'Insecure'
    'Chap'               = 'Weak'
    'MSChapv2'           = 'Acceptable'
    'Eap'                = 'Strong'
    'MachineCertificate' = 'Strong'
}

# Encryption level strength. NoEncryption and Optional are red flags; Required
# and Maximum are the safe tiers.
$EncryptionStrength = @{
    'NoEncryption' = 'Insecure'
    'Optional'     = 'Weak'
    'Required'     = 'Strong'
    'Maximum'      = 'Strong'
}

# Curated catalog of third-party VPN client signals. Each entry is matched by
# service-name pattern (Get-Service -Name) -- absent services are silently
# skipped, present services are surfaced in the report. Patterns ending in '*'
# are wildcarded (e.g. WireGuardTunnel$<name>). The friendly label is what
# appears in the report; vendor lets the technician spot which org owns the
# client when the service name is opaque.
$ThirdPartyClients = @(
    @{ Pattern = 'vpnagent';                Friendly = 'Cisco AnyConnect / Secure Client'; Vendor = 'Cisco' }
    @{ Pattern = 'csc_vpnagent';            Friendly = 'Cisco Secure Client (newer)';      Vendor = 'Cisco' }
    @{ Pattern = 'PanGPS';                  Friendly = 'GlobalProtect (service)';          Vendor = 'Palo Alto' }
    @{ Pattern = 'PanGPA';                  Friendly = 'GlobalProtect (agent)';            Vendor = 'Palo Alto' }
    @{ Pattern = 'JuniperNetworksTunnelService'; Friendly = 'Juniper / Pulse Tunnel';      Vendor = 'Ivanti' }
    @{ Pattern = 'PulseService';            Friendly = 'Pulse Secure Service';             Vendor = 'Ivanti' }
    @{ Pattern = 'OpenVPNService';          Friendly = 'OpenVPN (interactive service)';    Vendor = 'OpenVPN' }
    @{ Pattern = 'OpenVPNServiceLegacy';    Friendly = 'OpenVPN (legacy)';                 Vendor = 'OpenVPN' }
    @{ Pattern = 'OpenVPNServiceInteractive'; Friendly = 'OpenVPN GUI';                    Vendor = 'OpenVPN' }
    @{ Pattern = 'WireGuardManager';        Friendly = 'WireGuard manager';                Vendor = 'WireGuard' }
    @{ Pattern = 'WireGuardTunnel*';        Friendly = 'WireGuard tunnel (per-config)';    Vendor = 'WireGuard' }
    @{ Pattern = 'Tailscale';               Friendly = 'Tailscale';                        Vendor = 'Tailscale' }
    @{ Pattern = 'ZeroTierOneService';      Friendly = 'ZeroTier One';                     Vendor = 'ZeroTier' }
    @{ Pattern = 'CloudflareWARP';          Friendly = 'Cloudflare WARP';                  Vendor = 'Cloudflare' }
    @{ Pattern = 'nordvpn-service';         Friendly = 'NordVPN';                          Vendor = 'Nord Security' }
    @{ Pattern = 'ProtonVPNService';        Friendly = 'Proton VPN';                       Vendor = 'Proton AG' }
    @{ Pattern = 'F5 BIG-IP Edge Client';   Friendly = 'F5 BIG-IP Edge Client';            Vendor = 'F5' }
    @{ Pattern = 'CitrixWorkspaceUpdater';  Friendly = 'Citrix Workspace (heuristic)';     Vendor = 'Citrix' }
)

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function _vpnRowFromCmdlet {
    param($V, [string]$Scope)

    # AuthenticationMethod is an array; encode as a comma-separated string for
    # display while keeping the original for verdict logic.
    $authArr = @()
    if ($V.AuthenticationMethod) { $authArr = @($V.AuthenticationMethod | ForEach-Object { "$_" }) }

    return [PSCustomObject]@{
        Name                  = $V.Name
        Scope                 = $Scope
        ServerAddress         = $V.ServerAddress
        TunnelType            = "$($V.TunnelType)"
        AuthMethods           = $authArr
        AuthMethodsDisplay    = ($authArr -join ', ')
        EncryptionLevel       = "$($V.EncryptionLevel)"
        L2tpIPsecAuth         = "$($V.L2tpIPsecAuth)"
        UseWinlogonCredential = [bool]$V.UseWinlogonCredential
        ConnectionStatus      = "$($V.ConnectionStatus)"
        RememberCredential    = [bool]$V.RememberCredential
        SplitTunneling        = [bool]$V.SplitTunneling
        DnsSuffix             = $V.DnsSuffix
        IdleDisconnectSeconds = $V.IdleDisconnectSeconds
        ProfileType           = "$($V.ProfileType)"
    }
}

function Get-VpnConnections {
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $user = @(Get-VpnConnection -ErrorAction Stop)
        foreach ($v in $user) { $rows.Add((_vpnRowFromCmdlet -V $v -Scope 'User')) }
    } catch {
        Write-TKError -ScriptName 'portal.ps1' -Message "Get-VpnConnection (user scope) failed: $($_.Exception.Message)" -Category 'VPN'
    }
    try {
        $all = @(Get-VpnConnection -AllUserConnection -ErrorAction Stop)
        foreach ($v in $all) { $rows.Add((_vpnRowFromCmdlet -V $v -Scope 'AllUser')) }
    } catch {
        Write-TKError -ScriptName 'portal.ps1' -Message "Get-VpnConnection (all-user scope) failed: $($_.Exception.Message)" -Category 'VPN'
    }
    return @($rows)
}

function Get-AlwaysOnTriggers {
    param([array]$Connections)

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($conn in $Connections) {
        try {
            $params = @{ ConnectionName = $conn.Name; ErrorAction = 'Stop' }
            if ($conn.Scope -eq 'AllUser') { $params.AllUserConnection = $true }
            $trigger = Get-VpnConnectionTrigger @params
            if (-not $trigger) { continue }

            $apps = @()
            if ($trigger.ApplicationID) {
                $apps = @($trigger.ApplicationID | Where-Object { $_ })
            }
            $dns = @()
            if ($trigger.DnsConfiguration) {
                $dns = @($trigger.DnsConfiguration | ForEach-Object {
                    "$($_.DnsSuffix) -> $(($_.DnsServers -join ', '))"
                })
            }
            if ($apps.Count -gt 0 -or $dns.Count -gt 0) {
                $rows.Add([PSCustomObject]@{
                    ConnectionName = $conn.Name
                    Scope          = $conn.Scope
                    Apps           = $apps
                    DnsRules       = $dns
                })
            }
        } catch {
            # Connection has no trigger -- expected for plain manual VPNs.
            continue
        }
    }
    return @($rows)
}

function Get-NrptEntries {
    try {
        $rules = @(Get-DnsClientNrptPolicy -ErrorAction Stop)
    } catch {
        return @()
    }
    return @($rules | ForEach-Object {
        [PSCustomObject]@{
            Namespace          = ($_.Namespace -join ', ')
            DnsServers         = ($_.NameServers -join ', ')
            DirectAccessServers= (@($_.DirectAccessDnsServers) -join ', ')
            IpsecRequired      = [bool]$_.IPsecRequired
            EncryptionType     = "$($_.DnsSecValidationRequired)"
            Comment            = $_.Comment
        }
    })
}

function Get-VpnTunnelInterfaces {
    try {
        # Tunnel interfaces show up with InterfaceAlias starting with the VPN
        # connection name OR with PPP* / WireGuardTunnel*. Filter by both
        # ConnectionState (Connected) and a heuristic name match so we don't
        # pollute the table with every IPv6 tunnel pseudo-interface.
        $ifaces = @(Get-NetIPInterface -ErrorAction Stop | Where-Object {
            $_.ConnectionState -eq 'Connected' -and (
                $_.InterfaceAlias -match 'VPN|Wintun|WireGuard|Tailscale|GlobalProtect|AnyConnect|PPP'
            )
        })
    } catch {
        return @()
    }
    return @($ifaces | ForEach-Object {
        [PSCustomObject]@{
            InterfaceAlias  = $_.InterfaceAlias
            AddressFamily   = "$($_.AddressFamily)"
            ConnectionState = "$($_.ConnectionState)"
            Forwarding      = "$($_.Forwarding)"
            Dhcp            = "$($_.Dhcp)"
            InterfaceMetric = $_.InterfaceMetric
        }
    })
}

function Get-ThirdPartyVpnClients {
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $ThirdPartyClients) {
        try {
            $svcs = @(Get-Service -Name $entry.Pattern -ErrorAction SilentlyContinue)
        } catch { $svcs = @() }
        foreach ($s in $svcs) {
            $rows.Add([PSCustomObject]@{
                Name      = $s.Name
                Friendly  = $entry.Friendly
                Vendor    = $entry.Vendor
                Status    = "$($s.Status)"
                StartType = "$($s.StartType)"
            })
        }
    }
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Get-VpnAuthTier {
    param([array]$Methods)
    if (-not $Methods -or $Methods.Count -eq 0) { return 'Unknown' }
    # The weakest method in the configured list governs the tier -- a VPN that
    # accepts PAP as one of multiple methods is still PAP-vulnerable.
    $tier = 'Strong'
    foreach ($m in $Methods) {
        $t = if ($AuthStrength.ContainsKey("$m")) { $AuthStrength["$m"] } else { 'Unknown' }
        switch ($t) {
            'Insecure'   { return 'Insecure' }
            'Weak'       { if ($tier -ne 'Insecure') { $tier = 'Weak' } }
            'Acceptable' { if ($tier -in @('Strong')) { $tier = 'Acceptable' } }
            'Unknown'    { if ($tier -eq 'Strong') { $tier = 'Unknown' } }
        }
    }
    return $tier
}

function Get-VpnEncryptionTier {
    param([string]$Level)
    if ($EncryptionStrength.ContainsKey($Level)) { return $EncryptionStrength[$Level] }
    return 'Unknown'
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-PortalVerdict {
    param([array]$Vpns, [array]$Triggers, [array]$Nrpt, [array]$ThirdParty)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if ($Vpns.Count -eq 0 -and $ThirdParty.Count -eq 0) {
        return [PSCustomObject]@{
            Verdict = 'NONE CONFIGURED'
            Class   = 'info'
            Issues  = @()
            Warns   = @('No built-in Windows VPN connections and no third-party VPN clients detected. This is a posture statement, not a finding.')
        }
    }

    foreach ($v in $Vpns) {
        $authTier = Get-VpnAuthTier -Methods $v.AuthMethods
        $encTier  = Get-VpnEncryptionTier -Level $v.EncryptionLevel

        if ($authTier -eq 'Insecure') {
            $msg = "VPN '$($v.Name)' accepts PAP authentication -- credentials transmitted in cleartext."
            $issues.Add($msg)
            Write-TKError -ScriptName 'portal.ps1' -Message $msg -Category 'VPN'
        } elseif ($authTier -eq 'Weak') {
            $warns.Add("VPN '$($v.Name)' accepts CHAP -- deprecated; require MS-CHAPv2 or EAP.")
        }

        if ($encTier -eq 'Insecure') {
            $msg = "VPN '$($v.Name)' uses NoEncryption -- traffic flows in the clear over the tunnel."
            $issues.Add($msg)
            Write-TKError -ScriptName 'portal.ps1' -Message $msg -Category 'VPN'
        } elseif ($encTier -eq 'Weak') {
            $warns.Add("VPN '$($v.Name)' has Optional encryption -- raise to Required or Maximum.")
        }

        if ($v.SplitTunneling) {
            $hasNrptForThis = @($Nrpt | Where-Object { $_.Namespace -and $v.DnsSuffix -and ($_.Namespace -match [regex]::Escape($v.DnsSuffix)) }).Count -gt 0
            if (-not $hasNrptForThis) {
                $warns.Add("VPN '$($v.Name)' is split-tunnel with no NRPT entry covering its DNS suffix '$($v.DnsSuffix)' -- internal-name resolution may leak to the public resolver.")
            }
        }

        if ($v.Scope -eq 'AllUser' -and $v.ConnectionStatus -eq 'Disconnected') {
            $warns.Add("All-user VPN '$($v.Name)' is currently Disconnected -- expected if Always-On has not triggered yet, but worth confirming.")
        }
    }

    # Concurrent third-party clients -> routing-conflict risk.
    $vendors = @($ThirdParty | Select-Object -ExpandProperty Vendor -Unique)
    if ($vendors.Count -gt 1) {
        $warns.Add("$($vendors.Count) different VPN vendors detected on this machine ($([string]::Join(', ', $vendors))) -- competing route tables and DNS resolvers can produce intermittent connectivity.")
    }

    $verdict = if ($issues.Count -gt 0) { 'AT RISK' }
               elseif ($warns.Count -gt 0) { 'ATTENTION NEEDED' }
               else { 'OK' }
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

function Build-PortalReport {
    param([array]$Vpns, [array]$Triggers, [array]$Nrpt, [array]$Tunnels, [array]$ThirdParty, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _tierBadge { param([string]$Tier)
        switch ($Tier) {
            'Strong'     { "<span class='tk-badge-ok'>$(EscHtml $Tier)</span>" }
            'Acceptable' { "<span class='tk-badge-info'>$(EscHtml $Tier)</span>" }
            'Weak'       { "<span class='tk-badge-warn'>$(EscHtml $Tier)</span>" }
            'Insecure'   { "<span class='tk-badge-err'>$(EscHtml $Tier)</span>" }
            default      { "<span class='tk-badge-info'>$(EscHtml $Tier)</span>" }
        }
    }
    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-info'>No</span>" } }
    function _ynWarn { param($b) if ($b) { "<span class='tk-badge-warn'>Yes</span>" } else { "<span class='tk-badge-ok'>No</span>" } }

    # VPN connections table
    $vRows = [System.Text.StringBuilder]::new()
    if ($Vpns.Count -eq 0) {
        [void]$vRows.Append("<tr><td colspan='9' class='tk-badge-info' style='text-align:center;'>No built-in Windows VPN connections configured.</td></tr>")
    } else {
        foreach ($v in $Vpns | Sort-Object Scope, Name) {
            $authTier = Get-VpnAuthTier -Methods $v.AuthMethods
            $encTier  = Get-VpnEncryptionTier -Level $v.EncryptionLevel
            $stateBadge = switch ($v.ConnectionStatus) {
                'Connected'    { "<span class='tk-badge-ok'>Connected</span>" }
                'Disconnected' { "<span class='tk-badge-info'>Disconnected</span>" }
                default        { "<span class='tk-badge-warn'>$(EscHtml $v.ConnectionStatus)</span>" }
            }
            [void]$vRows.Append(
                "<tr><td>$(EscHtml $v.Name)</td><td>$(EscHtml $v.Scope)</td>" +
                "<td class='tk-mono'>$(EscHtml $v.ServerAddress)</td>" +
                "<td>$(EscHtml $v.TunnelType)</td>" +
                "<td>$(EscHtml $v.AuthMethodsDisplay) $(_tierBadge $authTier)</td>" +
                "<td>$(EscHtml $v.EncryptionLevel) $(_tierBadge $encTier)</td>" +
                "<td>$(_ynWarn $v.SplitTunneling)</td>" +
                "<td>$stateBadge</td>" +
                "<td>$(EscHtml $v.ProfileType)</td></tr>"
            )
        }
    }

    # Always-On triggers
    $tRows = [System.Text.StringBuilder]::new()
    if ($Triggers.Count -eq 0) {
        [void]$tRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No app-triggered Always-On entries configured.</td></tr>")
    } else {
        foreach ($t in $Triggers) {
            $appList = if ($t.Apps.Count -gt 0) { ($t.Apps | ForEach-Object { "<code class='tk-mono'>$(EscHtml $_)</code>" }) -join '<br/>' } else { '<span class=''tk-badge-info''>none</span>' }
            $dnsList = if ($t.DnsRules.Count -gt 0) { ($t.DnsRules | ForEach-Object { "<code class='tk-mono'>$(EscHtml $_)</code>" }) -join '<br/>' } else { '<span class=''tk-badge-info''>none</span>' }
            [void]$tRows.Append(
                "<tr><td>$(EscHtml $t.ConnectionName)</td><td>$(EscHtml $t.Scope)</td>" +
                "<td>$appList</td><td>$dnsList</td></tr>"
            )
        }
    }

    # NRPT
    $nRows = [System.Text.StringBuilder]::new()
    if ($Nrpt.Count -eq 0) {
        [void]$nRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No NRPT rules configured.</td></tr>")
    } else {
        foreach ($n in $Nrpt) {
            [void]$nRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $n.Namespace)</td>" +
                "<td class='tk-mono'>$(EscHtml $n.DnsServers)</td>" +
                "<td class='tk-mono'>$(EscHtml $n.DirectAccessServers)</td>" +
                "<td>$(_ynWarn $n.IpsecRequired)</td>" +
                "<td>$(EscHtml $n.Comment)</td></tr>"
            )
        }
    }

    # Tunnel interfaces
    $iRows = [System.Text.StringBuilder]::new()
    if ($Tunnels.Count -eq 0) {
        [void]$iRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No active VPN tunnel interfaces detected.</td></tr>")
    } else {
        foreach ($i in $Tunnels) {
            [void]$iRows.Append(
                "<tr><td>$(EscHtml $i.InterfaceAlias)</td><td>$(EscHtml $i.AddressFamily)</td>" +
                "<td>$(EscHtml $i.ConnectionState)</td><td>$(EscHtml $i.Dhcp)</td>" +
                "<td>$(EscHtml $i.InterfaceMetric)</td></tr>"
            )
        }
    }

    # Third-party clients
    $cRows = [System.Text.StringBuilder]::new()
    if ($ThirdParty.Count -eq 0) {
        [void]$cRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No third-party VPN clients detected.</td></tr>")
    } else {
        foreach ($tp in $ThirdParty) {
            $statusBadge = if ($tp.Status -eq 'Running') { "<span class='tk-badge-ok'>Running</span>" } else { "<span class='tk-badge-info'>$(EscHtml $tp.Status)</span>" }
            [void]$cRows.Append(
                "<tr><td>$(EscHtml $tp.Friendly)</td><td>$(EscHtml $tp.Vendor)</td>" +
                "<td class='tk-mono'>$(EscHtml $tp.Name)</td><td>$statusBadge</td>" +
                "<td>$(EscHtml $tp.StartType)</td></tr>"
            )
        }
    }

    $findingsList = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$findingsList.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>") }
    foreach ($w in $Verdict.Warns)  { [void]$findingsList.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$findingsList.Append("<li class='tk-badge-ok'>No insecure VPN configurations detected.</li>")
    }

    $allUserCount = @($Vpns | Where-Object { $_.Scope -eq 'AllUser' }).Count
    $userCount    = @($Vpns | Where-Object { $_.Scope -eq 'User' }).Count
    $insecureCount = @($Vpns | Where-Object {
        (Get-VpnAuthTier -Methods $_.AuthMethods) -eq 'Insecure' -or
        (Get-VpnEncryptionTier -Level $_.EncryptionLevel) -eq 'Insecure'
    }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'P.O.R.T.A.L. VPN / Always-On VPN Audit' `
        -ScriptName 'P.O.R.T.A.L.' `
        -Subtitle   "${orgPrefix}VPN / Always-On VPN Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'             = $machine
            'Generated'           = $reportDate
            'Verdict'             = $Verdict.Verdict
            'Built-in VPNs'       = $Vpns.Count
            'Third-Party Clients' = $ThirdParty.Count
        }) `
        -NavItems   @('Verdict', 'Built-in VPNs', 'Always-On Triggers', 'NRPT', 'Tunnel Interfaces', 'Third-Party Clients')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'P.O.R.T.A.L. v6.1'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">VPN Posture</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Vpns.Count)</div><div class="tk-summary-lbl">Built-in VPNs</div></div>
    <div class="tk-summary-card $(if ($insecureCount -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$insecureCount</div><div class="tk-summary-lbl">Insecure Auth / Encryption</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$allUserCount</div><div class="tk-summary-lbl">All-user (Always-On candidate)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Triggers.Count)</div><div class="tk-summary-lbl">App Triggers</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Nrpt.Count)</div><div class="tk-summary-lbl">NRPT Entries</div></div>
    <div class="tk-summary-card $(if ($ThirdParty.Count -gt 0) { 'warn' } else { 'info' })"><div class="tk-summary-num">$($ThirdParty.Count)</div><div class="tk-summary-lbl">Third-Party Clients</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($findingsList.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Built-in Windows VPN Connections</span><span class="tk-section-num">$userCount user / $allUserCount all-user</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Name</th><th>Scope</th><th>Server</th><th>Tunnel</th><th>Auth</th><th>Encryption</th><th>Split-Tunnel</th><th>State</th><th>Profile</th></tr></thead>
        <tbody>$($vRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Always-On VPN App Triggers</span><span class="tk-section-num">$($Triggers.Count) configured</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Connection</th><th>Scope</th><th>Triggering Apps</th><th>DNS Triggers</th></tr></thead>
        <tbody>$($tRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">NRPT Entries</span><span class="tk-section-num">$($Nrpt.Count) rules</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Namespace</th><th>DNS Servers</th><th>DirectAccess Servers</th><th>IPsec required</th><th>Comment</th></tr></thead>
        <tbody>$($nRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">VPN Tunnel Interfaces</span><span class="tk-section-num">$($Tunnels.Count) connected</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Interface</th><th>Family</th><th>State</th><th>DHCP</th><th>Metric</th></tr></thead>
        <tbody>$($iRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Third-Party VPN Clients</span><span class="tk-section-num">$($ThirdParty.Count) detected</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Product</th><th>Vendor</th><th>Service</th><th>Status</th><th>Start Type</th></tr></thead>
        <tbody>$($cRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-PortalBanner

Write-Section "BUILT-IN VPN CONNECTIONS"
$vpns = Get-VpnConnections
Write-Host ("  Built-in VPN connections : {0}" -f $vpns.Count) -ForegroundColor $C.Info
foreach ($v in $vpns) {
    $authTier = Get-VpnAuthTier -Methods $v.AuthMethods
    $encTier  = Get-VpnEncryptionTier -Level $v.EncryptionLevel
    $color = if ($authTier -eq 'Insecure' -or $encTier -eq 'Insecure') { $C.Error }
             elseif ($authTier -in @('Weak','Acceptable') -or $encTier -eq 'Weak') { $C.Warning }
             else { $C.Success }
    Write-Host ("  - {0,-25} [{1}] auth={2} enc={3} state={4}" -f $v.Name, $v.Scope, $v.AuthMethodsDisplay, $v.EncryptionLevel, $v.ConnectionStatus) -ForegroundColor $color
}
Write-Host ""

Write-Section "ALWAYS-ON TRIGGERS"
$triggers = Get-AlwaysOnTriggers -Connections $vpns
Write-Host ("  App-trigger / DNS-trigger entries : {0}" -f $triggers.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "NRPT ENTRIES"
$nrpt = Get-NrptEntries
Write-Host ("  NRPT rules                : {0}" -f $nrpt.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "ACTIVE TUNNEL INTERFACES"
$tunnels = Get-VpnTunnelInterfaces
Write-Host ("  Connected tunnel ifaces   : {0}" -f $tunnels.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "THIRD-PARTY VPN CLIENTS"
$thirdParty = Get-ThirdPartyVpnClients
Write-Host ("  Detected client services  : {0}" -f $thirdParty.Count) -ForegroundColor $C.Info
foreach ($tp in $thirdParty) {
    $color = if ($tp.Status -eq 'Running') { $C.Success } else { $C.Info }
    Write-Host ("    - {0,-30} ({1})  service: {2,-30} {3}" -f $tp.Friendly, $tp.Vendor, $tp.Name, $tp.Status) -ForegroundColor $color
}
Write-Host ""

$verdict = Get-PortalVerdict -Vpns $vpns -Triggers $triggers -Nrpt $nrpt -ThirdParty $thirdParty

Write-Section "VPN VERDICT"
$verdictColor = switch ($verdict.Class) {
    'ok'   { $C.Success }
    'warn' { $C.Warning }
    'err'  { $C.Error }
    default { $C.Info }
}
Write-Host "  $($verdict.Verdict)" -ForegroundColor $verdictColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Class -eq 'ok' -and $verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-PortalReport -Vpns $vpns -Triggers $triggers -Nrpt $nrpt -Tunnels $tunnels -ThirdParty $thirdParty -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "PORTAL_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
