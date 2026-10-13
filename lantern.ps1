# lantern.ps1 - L.A.N.T.E.R.N. — Locates & Audits Network Topology, Enumerating Resources & Nodes
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
    L.A.N.T.E.R.N. — Locates & Audits Network Topology, Enumerating Resources & Nodes
    Network Discovery & Asset Inventory Tool for PowerShell 5.1+

.DESCRIPTION
    Scans the local /24 subnet to discover all live hosts via parallel ping sweep,
    resolves hostnames through DNS reverse lookup, reads MAC addresses from the ARP
    neighbor table, optionally performs TCP port scans against common service ports,
    and produces a dark-themed HTML inventory report with color-coded port badges
    and summary cards. Results can also be exported to CSV.

.USAGE
    PS C:\> .\lantern.ps1                           # Interactive menu
    PS C:\> .\lantern.ps1 -Unattended -Action Sweep # Run sweep and export HTML silently

.NOTES
    Version : 6.1

#>

param(
    [switch]$Unattended,
    [ValidateSet('Sweep')]
    [string]$Action = 'Sweep',
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

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

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

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

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
Assert-AdminPrivilege

# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT PATH RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

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

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
}

# ─────────────────────────────────────────────────────────────────────────────
# SHARED STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:DiscoveredHosts = @()

# Common ports to scan
$script:ScanPorts = @(21, 22, 23, 80, 443, 445, 3389, 5985, 8080, 8443)

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-LanternBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗      █████╗ ███╗   ██╗████████╗███████╗██████╗ ███╗   ██╗
  ██║     ██╔══██╗████╗  ██║╚══██╔══╝██╔════╝██╔══██╗████╗  ██║
  ██║     ███████║██╔██╗ ██║   ██║   █████╗  ██████╔╝██╔██╗ ██║
  ██║     ██╔══██║██║╚██╗██║   ██║   ██╔══╝  ██╔══██╗██║╚██╗██║
  ███████╗██║  ██║██║ ╚████║   ██║   ███████╗██║  ██║██║ ╚████║
  ╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝   ╚═╝   ╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    L.A.N.T.E.R.N. — Locates & Audits Network Topology, Enumerating Resources & Nodes" -ForegroundColor Cyan
    Write-Host "    Network Discovery & Asset Inventory Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-TCPPort {
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 1500)
    try {
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $ar  = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $tcp.Close()
        return $ok
    } catch { return $false }
}

function Get-LocalSubnetPrefix {
    $ip = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
          Where-Object {
              $_.IPAddress -notmatch '^127\.' -and
              $_.IPAddress -notmatch '^169\.254\.' -and
              $_.PrefixOrigin -ne 'WellKnown'
          } |
          Sort-Object -Property PrefixLength -Descending |
          Select-Object -First 1

    if (-not $ip) { return $null }
    $parts = $ip.IPAddress -split '\.'
    return ($parts[0..2] -join '.')
}

# ─────────────────────────────────────────────────────────────────────────────
# SUBNET SWEEP
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-SubnetSweep {
    Write-Section "SUBNET DISCOVERY SWEEP"

    $prefix = Get-LocalSubnetPrefix
    if (-not $prefix) {
        Write-Host "  [-] Could not determine local IP address. Ensure a network adapter is connected." -ForegroundColor $C.Error
        Write-Host ""
        return
    }

    Write-Host ("  [*] Local subnet detected: {0}.0/24" -f $prefix) -ForegroundColor $C.Info
    Write-Host ("  [*] Sweeping {0}.1 – {0}.254 using parallel ping jobs..." -f $prefix) -ForegroundColor $C.Progress
    Write-Host ""

    # Build runspace pool for parallel pings
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 50)
    $pool.Open()

    $jobs = [System.Collections.Generic.List[hashtable]]::new()

    $pingScript = {
        param([string]$IP)
        $result = @{ IP = $IP; Alive = $false; ResponseTimeMs = -1 }
        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($IP, 1000)
            if ($reply.Status -eq 'Success') {
                $result.Alive = $true
                $result.ResponseTimeMs = [int]$reply.RoundtripTime
            }
        } catch {}
        return $result
    }

    # Submit all 254 jobs
    for ($i = 1; $i -le 254; $i++) {
        $ip  = "$prefix.$i"
        $ps  = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($pingScript).AddArgument($ip)
        $handle = $ps.BeginInvoke()
        $jobs.Add(@{ PS = $ps; Handle = $handle; Index = $i })
    }

    # Wait and show progress
    $timeout   = (Get-Date).AddSeconds(15)
    $completed = 0
    $alive     = [System.Collections.Generic.List[hashtable]]::new()

    while ($completed -lt $jobs.Count -and (Get-Date) -lt $timeout) {
        foreach ($job in $jobs) {
            if (-not $job.ContainsKey('Done') -and $job.Handle.IsCompleted) {
                $res = $job.PS.EndInvoke($job.Handle)
                $job.PS.Dispose()
                $job['Done'] = $true
                $completed++
                if ($res -and $res.Alive) {
                    $alive.Add($res)
                }
            }
        }
        $pct = [int](($completed / 254) * 100)
        Write-Progress -Activity "Sweeping $prefix.x" `
                       -Status ("Checked {0}/254 hosts — {1} responding" -f $completed, $alive.Count) `
                       -PercentComplete $pct
        Start-Sleep -Milliseconds 50
    }

    Write-Progress -Activity "Sweeping $prefix.x" -Completed

    # Force-close any incomplete jobs
    foreach ($job in $jobs) {
        if (-not $job.ContainsKey('Done')) {
            try { $job.PS.Stop() } catch {}
            $job.PS.Dispose()
        }
    }
    $pool.Close()
    $pool.Dispose()

    Write-Host ("  [+] Sweep complete. Found {0} responding host(s)." -f $alive.Count) -ForegroundColor $C.Success
    Write-Host ""

    if ($alive.Count -eq 0) {
        Write-Host "  [!!] No hosts responded. Check your network connection." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    # Read ARP table into a lookup dictionary
    Write-Host "  [*] Reading ARP neighbor table..." -ForegroundColor $C.Progress
    $arpLookup = @{}
    Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -ne '00-00-00-00-00-00' } |
        ForEach-Object { $arpLookup[$_.IPAddress] = $_.LinkLayerAddress }

    # Build discovered hosts list
    $script:DiscoveredHosts = @()
    $total  = $alive.Count
    $idx    = 0

    Write-Host "  [*] Resolving hostnames..." -ForegroundColor $C.Progress
    Write-Host ""

    foreach ($r in ($alive | Sort-Object { [version]$_.IP })) {
        $idx++
        Write-Progress -Activity "Resolving hostnames" `
                       -Status ("Resolving {0} ({1}/{2})" -f $r.IP, $idx, $total) `
                       -PercentComplete ([int](($idx / $total) * 100))

        $hostname = ''
        try {
            $entry    = [System.Net.Dns]::GetHostEntry($r.IP)
            $hostname = $entry.HostName
        } catch {}

        $mac = if ($arpLookup.ContainsKey($r.IP)) { $arpLookup[$r.IP] } else { '' }

        $script:DiscoveredHosts += [PSCustomObject]@{
            IP            = $r.IP
            Hostname      = $hostname
            MAC           = $mac
            Vendor        = ''
            OpenPorts     = @()
            ResponseTimeMs = $r.ResponseTimeMs
            LastSeen      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
    }

    Write-Progress -Activity "Resolving hostnames" -Completed

    Write-Host ("  [+] Discovery complete. {0} host(s) inventoried." -f $script:DiscoveredHosts.Count) -ForegroundColor $C.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PORT SCAN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-PortScan {
    Write-Section "TCP PORT SCAN"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $portList  = $script:ScanPorts
    $hostCount = $script:DiscoveredHosts.Count
    $totalOps  = $hostCount * $portList.Count
    $done      = 0

    Write-Host ("  [*] Scanning {0} host(s) across {1} ports ({2} total checks)..." -f $hostCount, $portList.Count, $totalOps) -ForegroundColor $C.Progress
    Write-Host ("      Ports: {0}" -f ($portList -join ', ')) -ForegroundColor $C.Info
    Write-Host ""

    foreach ($h in $script:DiscoveredHosts) {
        $h.OpenPorts = @()
        foreach ($port in $portList) {
            $done++
            $pct = [int](($done / $totalOps) * 100)
            Write-Progress -Activity "Port scanning" `
                           -Status ("Testing {0}:{1} ({2}/{3})" -f $h.IP, $port, $done, $totalOps) `
                           -PercentComplete $pct

            if (Test-TCPPort -Hostname $h.IP -Port $port -TimeoutMs 1500) {
                $h.OpenPorts += $port
            }
        }

        $portDisplay = if ($h.OpenPorts.Count -gt 0) {
            $h.OpenPorts -join ', '
        } else {
            'none'
        }
        $portColor = if ($h.OpenPorts.Count -gt 0) { $C.Warning } else { $C.Info }
        Write-Host ("  {0,-18} {1}" -f $h.IP, $portDisplay) -ForegroundColor $portColor
    }

    Write-Progress -Activity "Port scanning" -Completed
    Write-Host ""
    Write-Host "  [+] Port scan complete." -ForegroundColor $C.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW HOST TABLE
# ─────────────────────────────────────────────────────────────────────────────

function Show-HostTable {
    Write-Section "DISCOVERED HOSTS"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $hdr = "  {0,-18} {1,-30} {2,-20} {3,-8} {4}" -f "IP Address", "Hostname", "MAC Address", "RTT(ms)", "Open Ports"
    Write-Host $hdr -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 90)) -ForegroundColor $C.Header

    foreach ($h in $script:DiscoveredHosts | Sort-Object { [version]$_.IP }) {
        $ports    = if ($h.OpenPorts.Count -gt 0) { $h.OpenPorts -join ', ' } else { '—' }
        $hostname = if ($h.Hostname) { $h.Hostname } else { '—' }
        $mac      = if ($h.MAC)      { $h.MAC      } else { '—' }
        $rtt      = if ($h.ResponseTimeMs -ge 0) { $h.ResponseTimeMs.ToString() } else { '—' }

        $rowColor = if ($h.OpenPorts -contains 3389 -or $h.OpenPorts -contains 445) {
            $C.Warning
        } elseif ($h.OpenPorts.Count -gt 0) {
            $C.Success
        } else {
            $C.Info
        }

        Write-Host ("  {0,-18} {1,-30} {2,-20} {3,-8} {4}" -f $h.IP, $hostname, $mac, $rtt, $ports) -ForegroundColor $rowColor
    }

    Write-Host ""
    Write-Host ("  Total hosts: {0}" -f $script:DiscoveredHosts.Count) -ForegroundColor $C.Info

    $rdpExposed = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 3389 }).Count
    $smbExposed = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 445  }).Count

    if ($rdpExposed -gt 0) {
        Write-Host ("  [!!] RDP exposed (port 3389): {0} host(s)" -f $rdpExposed) -ForegroundColor $C.Warning
    }
    if ($smbExposed -gt 0) {
        Write-Host ("  [!!] SMB exposed (port 445): {0} host(s)"  -f $smbExposed) -ForegroundColor $C.Warning
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    Write-Section "EXPORT HTML REPORT"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportName   = "LANTERN_$timestamp.html"
    $reportPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportName
    $generated    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $totalHosts   = $script:DiscoveredHosts.Count
    $rdpExposed   = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 3389 }).Count
    $smbExposed   = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 445  }).Count

    Write-Host "  [*] Building HTML report..." -ForegroundColor $C.Progress

    # Port badge helper
    function Get-PortBadgeHtml {
        param([int[]]$Ports)
        if (-not $Ports -or $Ports.Count -eq 0) { return '<span class="tk-mono" style="color:#666;">none</span>' }

        $badges = foreach ($p in $Ports) {
            $label = switch ($p) {
                21   { 'FTP'      }
                22   { 'SSH'      }
                23   { 'Telnet'   }
                80   { 'HTTP'     }
                443  { 'HTTPS'    }
                445  { 'SMB'      }
                3389 { 'RDP'      }
                5985 { 'WinRM'    }
                8080 { 'HTTP-Alt' }
                8443 { 'HTTPS-Alt'}
                default { "$p"   }
            }

            $badgeClass = switch ($p) {
                { $_ -in @(3389, 445, 23) } { 'tk-badge-err'  }
                { $_ -in @(80, 8080) }       { 'tk-badge-warn' }
                { $_ -in @(443, 8443) }      { 'tk-badge-ok'   }
                default                      { 'tk-badge-info' }
            }

            "<span class=`"$badgeClass`">$label</span>"
        }
        return $badges -join ' '
    }

    # Build table rows
    $rows = foreach ($h in $script:DiscoveredHosts | Sort-Object { [version]$_.IP }) {
        $hostname   = if ($h.Hostname) { [System.Net.WebUtility]::HtmlEncode($h.Hostname) } else { '<span style="color:#555;">-</span>' }
        $mac        = if ($h.MAC)      { $h.MAC } else { '<span style="color:#555;">-</span>' }
        $rtt        = if ($h.ResponseTimeMs -ge 0) { "$($h.ResponseTimeMs) ms" } else { '-' }
        $portBadges = Get-PortBadgeHtml -Ports $h.OpenPorts

        $rowRisk = if ($h.OpenPorts -contains 3389 -or $h.OpenPorts -contains 445 -or $h.OpenPorts -contains 23) {
            'err'
        } elseif ($h.OpenPorts.Count -gt 0) {
            'ok'
        } else {
            ''
        }

        @"
            <tr>
                <td>$($h.IP)</td>
                <td>$hostname</td>
                <td class="tk-mono">$mac</td>
                <td>$portBadges</td>
                <td>$rtt</td>
                $(if ($rowRisk) { "<td><span class=`"tk-badge-$rowRisk`">$(if ($rowRisk -eq 'err') {'Risk'} else {'Active'})</span></td>" } else { '<td></td>' })
            </tr>
"@
    }

    $tableRows = $rows -join "`n"

    $rdpClass = if ($rdpExposed -gt 0) { 'err' } else { 'ok' }
    $smbClass = if ($smbExposed -gt 0) { 'warn' } else { 'ok' }

    $subnet = Get-LocalSubnetPrefix
    $subnetLabel = if ($subnet) { "$subnet.0/24" } else { 'Unknown' }

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $htmlReport = (Get-TKHtmlHead `
        -Title      'L.A.N.T.E.R.N. Network Discovery Report' `
        -ScriptName 'L.A.N.T.E.R.N.' `
        -Subtitle   "${orgPrefix}Network Discovery & Asset Inventory" `
        -MetaItems  ([ordered]@{
            'Generated' = $generated
            'Subnet'    = $subnetLabel
            'Host'      = $env:COMPUTERNAME
        }) `
        -NavItems   @('Host Inventory')) + @"

<div class="tk-info-box">
  <span class="tk-info-label">Note</span>
  Port scan covers: 21 (FTP), 22 (SSH), 23 (Telnet), 80 (HTTP), 443 (HTTPS), 445 (SMB), 3389 (RDP), 5985 (WinRM), 8080, 8443.
  An empty port column means no port scan has been performed yet for that host.
</div>

<div class="tk-summary-row">
  <div class="tk-summary-card ok">
    <div class="tk-summary-num">$totalHosts</div>
    <div class="tk-summary-lbl">Total Hosts</div>
  </div>
  <div class="tk-summary-card ok">
    <div class="tk-summary-num">$totalHosts</div>
    <div class="tk-summary-lbl">Responding</div>
  </div>
  <div class="tk-summary-card $rdpClass">
    <div class="tk-summary-num">$rdpExposed</div>
    <div class="tk-summary-lbl">RDP Exposed (3389)</div>
  </div>
  <div class="tk-summary-card $smbClass">
    <div class="tk-summary-num">$smbExposed</div>
    <div class="tk-summary-lbl">SMB Exposed (445)</div>
  </div>
</div>

<div class="tk-section">
  <div class="tk-section-title">Host Inventory</div>
  <table class="tk-table">
    <thead>
      <tr>
        <th>IP Address</th>
        <th>Hostname</th>
        <th>MAC Address</th>
        <th>Open Ports</th>
        <th>Response Time</th>
        <th>Risk</th>
      </tr>
    </thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

"@ + (Get-TKHtmlFoot -ScriptName 'L.A.N.T.E.R.N. v6.1')

    try {
        [System.IO.File]::WriteAllText($reportPath, $htmlReport, [System.Text.Encoding]::UTF8)
        Write-Host ("  [+] HTML report saved: {0}" -f $reportPath) -ForegroundColor $C.Success
    } catch {
        Write-Host ("  [-] Failed to write report: {0}" -f $_) -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# CSV EXPORT
# ─────────────────────────────────────────────────────────────────────────────

function Export-CsvReport {
    Write-Section "EXPORT CSV"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "LANTERN_$timestamp.csv"

    try {
        $script:DiscoveredHosts |
            Sort-Object { [version]$_.IP } |
            Select-Object IP, Hostname, MAC,
                @{ N='OpenPorts'; E={ $_.OpenPorts -join ',' } },
                ResponseTimeMs, LastSeen |
            Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        Write-Host ("  [+] CSV saved: {0}" -f $csvPath) -ForegroundColor $C.Success
    } catch {
        Write-Host ("  [-] Failed to write CSV: {0}" -f $_) -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-LanternBanner
    switch ($Action) {
        'Sweep' {
            Invoke-SubnetSweep
            Build-HtmlReport
        }
    }
} else {
    $choice = ''

    do {
        Show-LanternBanner

        # Status line if hosts are already known
        if ($script:DiscoveredHosts.Count -gt 0) {
            Write-Host ("  Inventory: {0} host(s) in memory" -f $script:DiscoveredHosts.Count) -ForegroundColor $C.Info

            $lastSeen = $script:DiscoveredHosts | Select-Object -Last 1 -ExpandProperty LastSeen
            Write-Host ("  Last sweep: {0}" -f $lastSeen) -ForegroundColor $C.Info
            Write-Host ""
        }

        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Run discovery          (ping sweep /24 subnet)" -ForegroundColor $C.Info
        Write-Host "  [2] Show discovered hosts  (console table)" -ForegroundColor $C.Info
        Write-Host "  [3] Port scan hosts        (TCP connect, 10 common ports)" -ForegroundColor $C.Info
        Write-Host "  [4] Export HTML report     (dark-themed inventory)" -ForegroundColor $C.Info
        Write-Host "  [5] Export CSV" -ForegroundColor $C.Info
        Write-Host "  [R] Re-run discovery" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-SubnetSweep }
            '2' { Show-HostTable }
            '3' { Invoke-PortScan }
            '4' { Build-HtmlReport }
            '5' { Export-CsvReport }
            'R' { Invoke-SubnetSweep }
            'Q' {
                Write-Host ""
                Write-Host "  Closing L.A.N.T.E.R.N." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-5, R, or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @('Q', 'R', '1')) {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}

if ($Transcript) { Stop-TKTranscript }
