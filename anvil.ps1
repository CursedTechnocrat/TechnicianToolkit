# anvil.ps1 - A.N.V.I.L. — Audits & Notates Vendor Inventory & Lifecycle
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
    A.N.V.I.L. — Audits & Notates Vendor Inventory & Lifecycle
    BIOS / UEFI / Firmware Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Inventories system identity, BIOS / UEFI state, Secure Boot posture,
    and vendor-specific firmware-update channel availability (Dell Command,
    HP CMSL, Lenovo Vantage / System Update, Microsoft Surface). Scans
    Windows Update for pending firmware and driver updates. Produces a
    dark-themed HTML report with a readiness verdict so a technician can
    confirm a machine's firmware is supportable before shipping it.

.USAGE
    PS C:\> .\anvil.ps1                    # Interactive run
    PS C:\> .\anvil.ps1 -Unattended        # Silent: export HTML and exit

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

function Show-AnvilBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  A.N.V.I.L. — Audits & Notates Vendor Inventory & Lifecycle" -ForegroundColor Cyan
    Write-Host "  BIOS / UEFI / Firmware Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Data collection and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — SYSTEM IDENTITY + BIOS
# ─────────────────────────────────────────────────────────────────────────────

function Get-SystemInfo {
    $cs   = Get-CimInstance -ClassName Win32_ComputerSystem       -ErrorAction SilentlyContinue
    $bios = Get-CimInstance -ClassName Win32_BIOS                 -ErrorAction SilentlyContinue
    $sys  = Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue

    $releaseDate = $null
    if ($bios -and $bios.ReleaseDate) {
        try {
            $releaseDate = [Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
        } catch {
            # Date parse failures happen on some virtualised BIOSes; leave $null.
        }
    }

    $biosAgeDays = if ($releaseDate) { [math]::Round(((Get-Date) - $releaseDate).TotalDays, 0) } else { $null }

    # Manufacturer is useful upstream for vendor-channel detection.
    $mfr = if ($cs -and $cs.Manufacturer) { $cs.Manufacturer } else { '' }
    $vendor = switch -Wildcard ($mfr.ToLower()) {
        '*dell*'    { 'Dell' }
        '*hp*'      { 'HP' }
        '*hewlett*' { 'HP' }
        '*lenovo*'  { 'Lenovo' }
        '*microsoft*' { 'Microsoft' }
        default     { 'Other' }
    }

    return [PSCustomObject]@{
        Manufacturer = $mfr
        Vendor       = $vendor
        Model        = if ($cs) { $cs.Model } else { '' }
        SystemSKU    = if ($sys) { $sys.IdentifyingNumber } else { '' }
        UUID         = if ($sys) { $sys.UUID } else { '' }
        SerialNumber = if ($bios) { $bios.SerialNumber } else { '' }
        BIOSVendor   = if ($bios) { $bios.Manufacturer } else { '' }
        BIOSVersion  = if ($bios) { $bios.SMBIOSBIOSVersion } else { '' }
        BIOSReleaseDate = $releaseDate
        BIOSAgeDays  = $biosAgeDays
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — UEFI / SECURE BOOT POSTURE
# ─────────────────────────────────────────────────────────────────────────────

function Get-UefiPosture {
    $bootMode = 'Unknown'
    try {
        # PartitionStyle GPT on the system disk is a decent proxy for UEFI boot.
        $sys = Get-Disk -Number ([int](Get-CimInstance -ClassName Win32_OperatingSystem).SystemDevice.Replace('\\.\PHYSICALDRIVE','')) -ErrorAction SilentlyContinue
        if (-not $sys) { $sys = Get-Disk | Where-Object { $_.IsSystem } | Select-Object -First 1 }
        if ($sys) {
            $bootMode = if ($sys.PartitionStyle -eq 'GPT') { 'UEFI' } else { 'Legacy/BIOS' }
        }
    } catch {
        # Fall through to $bootMode = 'Unknown'.
    }

    # Firmware type from the registry is the authoritative check but may be absent.
    $firmwareType = 'Unknown'
    try {
        $env = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'PEFirmwareType' -ErrorAction SilentlyContinue
        if ($env -and $env.PEFirmwareType) {
            $firmwareType = switch ([int]$env.PEFirmwareType) {
                1 { 'BIOS (legacy)' }
                2 { 'UEFI' }
                default { "Unknown ($($env.PEFirmwareType))" }
            }
        }
    } catch {
        # PEFirmwareType may not exist on older builds.
    }

    $secureBoot = $null
    try {
        $secureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
    } catch {
        # Confirm-SecureBootUEFI requires UEFI firmware; on legacy BIOS it throws.
        $secureBoot = $null
    }

    return [PSCustomObject]@{
        BootMode         = $bootMode
        FirmwareType     = $firmwareType
        SecureBootState  = if ($null -eq $secureBoot) { 'Unsupported' } elseif ($secureBoot) { 'Enabled' } else { 'Disabled' }
        SecureBoot       = $secureBoot
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — VENDOR-SPECIFIC UPDATE CHANNELS
# ─────────────────────────────────────────────────────────────────────────────

# Known install locations for per-vendor firmware-update tooling. We test for
# the tool binary on disk — presence is what matters to a technician ("can I
# check firmware updates on this machine?"). Running the tool is out of scope
# for an audit.
$script:VendorChannels = @(
    @{ Vendor = 'Dell';      Name = 'Dell Command | Update (CLI)';       Paths = @('C:\Program Files\Dell\CommandUpdate\dcu-cli.exe', 'C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe') }
    @{ Vendor = 'Dell';      Name = 'Dell Command | Update (GUI)';       Paths = @('C:\Program Files\Dell\CommandUpdate\DellCommandUpdate.exe', 'C:\Program Files (x86)\Dell\CommandUpdate\DellCommandUpdate.exe') }
    @{ Vendor = 'HP';        Name = 'HP Image Assistant (HPIA)';         Paths = @('C:\Program Files\HP\HPIA\HPImageAssistant.exe', 'C:\SWSetup\HPIA\HPImageAssistant.exe') }
    @{ Vendor = 'HP';        Name = 'HP Support Assistant';              Paths = @('C:\Program Files (x86)\Hewlett-Packard\HP Support Solutions\HPSF.exe') }
    @{ Vendor = 'Lenovo';    Name = 'Lenovo System Update (Tvsu)';       Paths = @('C:\Program Files (x86)\Lenovo\System Update\Tvsu.exe') }
    @{ Vendor = 'Lenovo';    Name = 'Lenovo Vantage';                    Paths = @("$env:LOCALAPPDATA\Packages\E046963F.LenovoSettingsforEnterprise_k1h2ywk1493x8") }
    @{ Vendor = 'Microsoft'; Name = 'Surface UEFI Configurator (SEMM)';  Paths = @('C:\Program Files (x86)\Microsoft\Surface\UefiConfigurator\UEFIConfigurator.exe') }
)

function Get-VendorChannelStatus {
    param([string]$Vendor)

    $applicable = $script:VendorChannels | Where-Object { $_.Vendor -eq $Vendor -or $Vendor -eq 'Other' }
    if (-not $applicable) {
        return @([PSCustomObject]@{
            Name      = "(no known firmware-update channel for '$Vendor')"
            Vendor    = $Vendor
            Found     = $false
            Path      = ''
        })
    }

    $rows = foreach ($ch in $applicable) {
        $foundPath = $null
        foreach ($p in $ch.Paths) {
            if (Test-Path $p) { $foundPath = $p; break }
        }
        [PSCustomObject]@{
            Name   = $ch.Name
            Vendor = $ch.Vendor
            Found  = [bool]$foundPath
            Path   = if ($foundPath) { $foundPath } else { '(not installed)' }
        }
    }

    return @($rows)
}

# Also check PSWindowsUpdate availability since the next section uses it.
function Test-PSWindowsUpdate {
    return [bool](Get-Module -ListAvailable -Name PSWindowsUpdate)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — WINDOWS UPDATE PENDING FIRMWARE / DRIVER SCAN
# ─────────────────────────────────────────────────────────────────────────────

function Get-PendingFirmwareUpdates {
    Write-Step "Scanning Windows Update for pending driver/firmware updates..."

    if (-not (Test-PSWindowsUpdate)) {
        Write-Warn "PSWindowsUpdate module not installed — skipping WU scan. Install with:"
        Write-Info "  Install-Module PSWindowsUpdate -Scope CurrentUser -Force"
        return @()
    }

    try {
        Import-Module PSWindowsUpdate -ErrorAction Stop
        # -Category 'Drivers' covers firmware as well on most manufacturer distributions.
        $updates = Get-WindowsUpdate -Category 'Drivers' -ErrorAction Stop
    } catch {
        Write-Fail "Windows Update scan failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'anvil' -Message "Get-WindowsUpdate driver scan failed: $($_.Exception.Message)" -Category 'Windows Update'
        return @()
    }

    if (-not $updates -or $updates.Count -eq 0) {
        Write-Ok "No pending driver/firmware updates found."
        return @()
    }

    $rows = foreach ($u in $updates) {
        [PSCustomObject]@{
            Title    = $u.Title
            KB       = if ($u.KB) { $u.KB } else { '—' }
            Size     = if ($u.Size) { Format-Bytes $u.Size } else { '—' }
            Severity = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { 'Optional' }
        }
    }

    Write-Warn "$($rows.Count) pending driver/firmware update(s)."
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-Verdict {
    param($System, $Uefi, [array]$Channels, [array]$Pending)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if ($Uefi.FirmwareType -like 'BIOS*') {
        $issues.Add('Machine is booting in legacy BIOS mode — Secure Boot and modern OS requirements (Windows 11, BitLocker PCR-7) need UEFI.')
    }
    if ($Uefi.SecureBootState -eq 'Disabled') {
        $issues.Add('Secure Boot is disabled — required for Windows 11 and many security baselines.')
    } elseif ($Uefi.SecureBootState -eq 'Unsupported') {
        $warns.Add('Secure Boot unsupported — the firmware does not expose UEFI variables; likely legacy BIOS.')
    }

    if ($System.BIOSAgeDays -and $System.BIOSAgeDays -ge 730) {
        $warns.Add("BIOS is $([math]::Round($System.BIOSAgeDays / 365, 1)) years old ($($System.BIOSReleaseDate.ToString('yyyy-MM-dd'))) — review the vendor release channel for a newer version.")
    }

    $installedChannels = @($Channels | Where-Object { $_.Found })
    if ($System.Vendor -ne 'Other' -and $installedChannels.Count -eq 0) {
        $warns.Add("No $($System.Vendor) firmware-update tooling detected — deploying DCU / HPIA / Lenovo System Update / Vantage makes firmware maintenance trackable.")
    }

    if ($Pending.Count -gt 0) {
        $warns.Add("$($Pending.Count) pending driver/firmware update(s) are available via Windows Update.")
    }

    $verdict = if ($issues.Count -gt 0) { 'ATTENTION REQUIRED' }
               elseif ($warns.Count -gt 0) { 'REVIEW RECOMMENDED' }
               else { 'READY' }
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

function Build-HtmlReport {
    param($System, $Uefi, [array]$Channels, [array]$Pending, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $biosAgeStr = if ($System.BIOSAgeDays) { "$($System.BIOSAgeDays) days" } else { 'Unknown' }
    $biosDateStr = if ($System.BIOSReleaseDate) { $System.BIOSReleaseDate.ToString('yyyy-MM-dd') } else { 'Unknown' }

    $sbClass = switch ($Uefi.SecureBootState) {
        'Enabled'     { 'tk-badge-ok' }
        'Disabled'    { 'tk-badge-err' }
        'Unsupported' { 'tk-badge-warn' }
        default       { 'tk-badge-info' }
    }
    $fwClass = if ($Uefi.FirmwareType -eq 'UEFI') { 'tk-badge-ok' } elseif ($Uefi.FirmwareType -like 'BIOS*') { 'tk-badge-err' } else { 'tk-badge-warn' }

    $chRows = [System.Text.StringBuilder]::new()
    foreach ($ch in $Channels) {
        $badge = if ($ch.Found) { "<span class='tk-badge-ok'>Installed</span>" } else { "<span class='tk-badge-warn'>Not installed</span>" }
        [void]$chRows.Append("<tr><td>$(EscHtml $ch.Vendor)</td><td>$(EscHtml $ch.Name)</td><td>$badge</td><td><code>$(EscHtml $ch.Path)</code></td></tr>`n")
    }

    $pendingRows = [System.Text.StringBuilder]::new()
    if ($Pending.Count -eq 0) {
        [void]$pendingRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No pending driver/firmware updates.</td></tr>")
    } else {
        foreach ($p in $Pending) {
            [void]$pendingRows.Append("<tr><td>$(EscHtml $p.Title)</td><td>$(EscHtml $p.KB)</td><td>$(EscHtml $p.Size)</td><td>$(EscHtml $p.Severity)</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>Firmware posture is clean.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'A.N.V.I.L. Firmware / BIOS Audit' `
        -ScriptName 'A.N.V.I.L.' `
        -Subtitle   "${orgPrefix}BIOS / UEFI / Firmware Posture -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'      = $machine
            'Manufacturer' = $System.Manufacturer
            'Model'        = $System.Model
            'Serial'       = $System.SerialNumber
            'Generated'    = $reportDate
            'Verdict'      = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'System Identity', 'UEFI / Secure Boot', 'Vendor Channels', 'Pending Updates')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'A.N.V.I.L. v6.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Firmware Readiness</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $System.Vendor)</div><div class="tk-summary-lbl">Vendor</div></div>
    <div class="tk-summary-card $fwClass"><div class="tk-summary-num">$(EscHtml $Uefi.FirmwareType)</div><div class="tk-summary-lbl">Firmware Type</div></div>
    <div class="tk-summary-card $sbClass"><div class="tk-summary-num">$(EscHtml $Uefi.SecureBootState)</div><div class="tk-summary-lbl">Secure Boot</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$biosAgeStr</div><div class="tk-summary-lbl">BIOS Age</div></div>
    <div class="tk-summary-card $(if ($Pending.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Pending.Count)</div><div class="tk-summary-lbl">Pending WU Updates</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">System Identity</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Manufacturer</th><td>$(EscHtml $System.Manufacturer)</td></tr>
          <tr><th>Model</th><td>$(EscHtml $System.Model)</td></tr>
          <tr><th>System SKU / Service Tag</th><td><code>$(EscHtml $System.SystemSKU)</code></td></tr>
          <tr><th>Serial Number</th><td><code>$(EscHtml $System.SerialNumber)</code></td></tr>
          <tr><th>UUID</th><td><code>$(EscHtml $System.UUID)</code></td></tr>
          <tr><th>BIOS Vendor</th><td>$(EscHtml $System.BIOSVendor)</td></tr>
          <tr><th>BIOS Version</th><td><code>$(EscHtml $System.BIOSVersion)</code></td></tr>
          <tr><th>BIOS Release Date</th><td>$biosDateStr</td></tr>
          <tr><th>BIOS Age</th><td>$biosAgeStr</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">UEFI / Secure Boot</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Firmware Type</th><td><span class='$fwClass'>$(EscHtml $Uefi.FirmwareType)</span></td></tr>
          <tr><th>Boot Mode (from system disk)</th><td>$(EscHtml $Uefi.BootMode)</td></tr>
          <tr><th>Secure Boot State</th><td><span class='$sbClass'>$(EscHtml $Uefi.SecureBootState)</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Vendor Firmware-Update Channels</span><span class="tk-section-num">$($Channels.Count) channel(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Vendor</th><th>Tool</th><th>Status</th><th>Path</th></tr></thead>
        <tbody>$($chRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Pending Driver / Firmware Updates (Windows Update)</span><span class="tk-section-num">$($Pending.Count)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Title</th><th>KB</th><th>Size</th><th>Severity</th></tr></thead>
        <tbody>$($pendingRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-AnvilBanner

Write-Section "SYSTEM IDENTITY"
$system = Get-SystemInfo
Write-Ok ("Manufacturer: {0} ({1})" -f $system.Manufacturer, $system.Vendor)
Write-Ok ("Model       : {0}" -f $system.Model)
Write-Ok ("Serial      : {0}" -f $system.SerialNumber)
Write-Ok ("BIOS        : {0} v{1}" -f $system.BIOSVendor, $system.BIOSVersion)
if ($system.BIOSReleaseDate) {
    $age = [math]::Round($system.BIOSAgeDays / 365, 1)
    $color = if ($system.BIOSAgeDays -ge 730) { $C.Warning } else { $C.Success }
    Write-Host ("  BIOS released {0} ({1} yr old)" -f $system.BIOSReleaseDate.ToString('yyyy-MM-dd'), $age) -ForegroundColor $color
}
Write-Host ""

Write-Section "UEFI / SECURE BOOT"
$uefi = Get-UefiPosture
$fwColor = if ($uefi.FirmwareType -eq 'UEFI') { $C.Success } elseif ($uefi.FirmwareType -like 'BIOS*') { $C.Error } else { $C.Warning }
Write-Host ("  Firmware Type : {0}" -f $uefi.FirmwareType) -ForegroundColor $fwColor
Write-Host ("  Boot Mode     : {0}" -f $uefi.BootMode) -ForegroundColor $C.Info
$sbColor = if ($uefi.SecureBootState -eq 'Enabled') { $C.Success } elseif ($uefi.SecureBootState -eq 'Disabled') { $C.Error } else { $C.Warning }
Write-Host ("  Secure Boot   : {0}" -f $uefi.SecureBootState) -ForegroundColor $sbColor
Write-Host ""

Write-Section "VENDOR FIRMWARE CHANNELS"
$channels = Get-VendorChannelStatus -Vendor $system.Vendor
foreach ($ch in $channels) {
    $color = if ($ch.Found) { $C.Success } else { $C.Info }
    Write-Host ("  [{0,-10}] {1,-40} {2}" -f $ch.Vendor, $ch.Name, $(if ($ch.Found) { 'Installed' } else { 'Not installed' })) -ForegroundColor $color
}
Write-Host ""

Write-Section "WINDOWS UPDATE FIRMWARE / DRIVER SCAN"
$pending = Get-PendingFirmwareUpdates
if ($pending.Count -gt 0) {
    foreach ($p in $pending) {
        Write-Host ("  [{0}] {1}" -f $p.Severity, $p.Title) -ForegroundColor $C.Warning
    }
}
Write-Host ""

Write-Section "FIRMWARE READINESS VERDICT"
$verdict = Get-Verdict -System $system -Uefi $uefi -Channels $channels -Pending $pending
$vColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $vColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -System $system -Uefi $uefi -Channels $channels -Pending $pending -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "ANVIL_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
