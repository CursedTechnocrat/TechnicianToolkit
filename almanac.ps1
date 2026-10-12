# almanac.ps1 - A.L.M.A.N.A.C. — Assigned Licenses, Mailboxes, Authentication & Notable Account Counts
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
    A.L.M.A.N.A.C. — Assigned Licenses, Mailboxes, Authentication & Notable Account Counts
    Microsoft 365 License & Mailbox Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to Microsoft 365 via the Microsoft Graph API to audit license
    assignments, identify unlicensed and inactive users, check MFA registration
    status, and audit shared mailboxes. The required Microsoft.Graph module is
    installed automatically if missing. Results can be reviewed interactively
    in the console or exported as a combined dark-themed HTML report.

.USAGE
    PS C:\> .\almanac.ps1                    # Interactive menu
    PS C:\> .\almanac.ps1 -Unattended        # Auto-connect and export full audit report

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ENCODING & SCRIPT PATH
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
# CONNECTION STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:Connected      = $false
$script:ConnectedAs    = ''
$script:TenantDomain   = ''

# ─────────────────────────────────────────────────────────────────────────────
# SKU FRIENDLY NAME MAP
# ─────────────────────────────────────────────────────────────────────────────

$SkuNames = @{
    'SPE_E3'                          = 'Microsoft 365 E3'
    'SPE_E5'                          = 'Microsoft 365 E5'
    'SPE_F1'                          = 'Microsoft 365 F3'
    'O365_BUSINESS_PREMIUM'           = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS'        = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                   = 'Microsoft 365 Apps for Business'
    'ENTERPRISEPACK'                  = 'Office 365 E3'
    'ENTERPRISEPREMIUM'               = 'Office 365 E5'
    'STANDARDPACK'                    = 'Office 365 E1'
    'DESKLESSPACK'                    = 'Office 365 F3'
    'AAD_PREMIUM'                     = 'Azure AD Premium P1'
    'AAD_PREMIUM_P2'                  = 'Azure AD Premium P2'
    'EMS'                             = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                      = 'Enterprise Mobility + Security E5'
    'POWER_BI_PRO'                    = 'Power BI Pro'
    'POWER_BI_PREMIUM_PER_USER'       = 'Power BI Premium Per User'
    'FLOW_FREE'                       = 'Power Automate Free'
    'POWERFLOW_P1'                    = 'Power Automate Plan 1'
    'INTUNE_A'                        = 'Microsoft Intune'
    'INTUNE_A_D'                      = 'Microsoft Intune Device'
    'PROJECTPREMIUM'                  = 'Project Plan 5'
    'PROJECTPROFESSIONAL'             = 'Project Plan 3'
    'VISIOCLIENT'                     = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'               = 'Visio Plan 1'
    'TEAMS_EXPLORATORY'               = 'Microsoft Teams Exploratory'
    'TEAMS_FREE'                      = 'Microsoft Teams Free'
    'MCOSTANDARD'                     = 'Skype for Business Online Plan 2'
    'EXCHANGESTANDARD'                = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'              = 'Exchange Online Plan 2'
    'EXCHANGE_S_DESKLESS'             = 'Exchange Online Kiosk'
    'DEFENDER_ENDPOINT_P1'            = 'Microsoft Defender for Endpoint P1'
    'WIN_DEF_ATP'                     = 'Microsoft Defender for Endpoint P2'
    'RIGHTSMANAGEMENT'                = 'Azure Information Protection P1'
    'MCOMEETADV'                      = 'Microsoft Teams Audio Conferencing'
    'PHONESYSTEM_VIRTUALUSER'         = 'Microsoft Teams Phone Resource Account'
    'MCOPSTN1'                        = 'Microsoft Teams Domestic Calling Plan'
    'M365_LIGHTHOUSE_PARTNER_PLAN1'   = 'Microsoft 365 Lighthouse'
    'WINDOWS_STORE'                   = 'Windows Store for Business'
    'DEVELOPERPACK_E3'                = 'Microsoft 365 E3 Developer'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-AlmanacBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   █████╗ ██╗     ███╗   ███╗ █████╗ ███╗   ██╗ █████╗  ██████╗
  ██╔══██╗██║     ████╗ ████║██╔══██╗████╗  ██║██╔══██╗██╔════╝
  ███████║██║     ██╔████╔██║███████║██╔██╗ ██║███████║██║
  ██╔══██║██║     ██║╚██╔╝██║██╔══██║██║╚██╗██║██╔══██║██║
  ██║  ██║███████╗██║ ╚═╝ ██║██║  ██║██║ ╚████║██║  ██║╚██████╗
  ╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝╚═╝  ╚═╝ ╚═════╝

"@ -ForegroundColor Cyan
    Write-Host "  A.L.M.A.N.A.C. — Assigned Licenses, Mailboxes, Authentication & Notable Account Counts" -ForegroundColor Cyan
    Write-Host "  Microsoft 365 License & Mailbox Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE CHECK & INSTALL
# ─────────────────────────────────────────────────────────────────────────────

function Install-GraphModule {
    Write-Section "MODULE CHECK"

    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Reports'
    )

    $needsInstall = $false
    foreach ($mod in $requiredModules) {
        if (Get-Module -ListAvailable -Name $mod) {
            Write-Ok "$mod — installed"
        } else {
            Write-Warn "$mod — NOT found"
            $needsInstall = $true
        }
    }

    if ($needsInstall) {
        Write-Host ""
        if ($Unattended) {
            Write-Step "Installing Microsoft.Graph module (unattended)..."
            try {
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Ok "Microsoft.Graph installed successfully."
            } catch {
                Write-Fail "Failed to install Microsoft.Graph: $_"
                exit 1
            }
        } else {
            Write-Warn "One or more Microsoft.Graph sub-modules are not installed."
            Write-Info "The Microsoft.Graph module suite is required for A.L.M.A.N.A.C. to operate."
            Write-Host ""
            $ans = Read-Host "  Install Microsoft.Graph for current user? [Y/N]"
            if ($ans -match '^[Yy]') {
                Write-Step "Installing Microsoft.Graph — this may take a minute..."
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Ok "Microsoft.Graph installed successfully."
                } catch {
                    Write-Fail "Installation failed: $_"
                    Write-Info "Run manually: Install-Module Microsoft.Graph -Scope CurrentUser -Force"
                    exit 1
                }
            } else {
                Write-Fail "Microsoft.Graph is required. Exiting."
                exit 1
            }
        }
    } else {
        Write-Ok "All required Graph modules are available."
    }

    # Import the sub-modules
    foreach ($mod in $requiredModules) {
        try {
            Import-Module $mod -ErrorAction Stop
        } catch {
            Write-Warn "Could not import ${mod}: $_"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

$GraphScopes = @(
    'User.Read.All',
    'Directory.Read.All',
    'AuditLog.Read.All',
    'Reports.Read.All'
)

function Test-GraphConnection {
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account) {
            $script:Connected   = $true
            $script:ConnectedAs = $ctx.Account
            if ($ctx.TenantId) {
                # Try to get the verified domain for display
                try {
                    $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($org -and $org.VerifiedDomains) {
                        $primary = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
                        if ($primary) { $script:TenantDomain = $primary.Name }
                    }
                } catch { }
            }
            return $true
        }
    } catch { }
    $script:Connected   = $false
    $script:ConnectedAs = ''
    return $false
}

function Invoke-Connect {
    Write-Section "CONNECT TO MICROSOFT 365"
    Write-Step "Initiating interactive authentication..."
    Write-Info "Requested scopes: $($GraphScopes -join ', ')"
    Write-Host ""

    try {
        if ($Unattended) {
            # Attempt silent connect using existing token cache
            Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        } else {
            Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        }

        if (Test-GraphConnection) {
            Write-Ok "Connected as: $($script:ConnectedAs)"
            if ($script:TenantDomain) {
                Write-Ok "Tenant domain: $($script:TenantDomain)"
            }
        } else {
            Write-Fail "Connection appeared to succeed but context could not be verified."
        }
    } catch {
        $errMsg = $_.Exception.Message
        if ($Unattended) {
            Write-Fail "Silent connect failed. An existing auth token is required for unattended mode."
            Write-Info "Run the script interactively first to cache credentials, or use a service principal."
            Write-Info "Error: $errMsg"
            exit 1
        } else {
            Write-Fail "Authentication failed: $errMsg"
        }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# AUDIT FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Get-LicenseInventory {
    Write-Section "LICENSE SKU INVENTORY"
    Write-Step "Querying subscribed SKUs..."

    try {
        $skus = Get-MgSubscribedSku -All -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve license data: $_"
        return @()
    }

    if (-not $skus -or $skus.Count -eq 0) {
        Write-Warn "No SKUs found. The tenant may have no paid subscriptions."
        return @()
    }

    $inventory = foreach ($sku in $skus) {
        $partNumber  = $sku.SkuPartNumber
        $friendlyName = if ($SkuNames.ContainsKey($partNumber)) { $SkuNames[$partNumber] } else { $partNumber }
        $total     = $sku.PrepaidUnits.Enabled
        $assigned  = $sku.ConsumedUnits
        $available = $total - $assigned

        [PSCustomObject]@{
            SkuPartNumber = $partNumber
            FriendlyName  = $friendlyName
            Assigned      = $assigned
            Total         = $total
            Available     = $available
            CapabilityStatus = $sku.CapabilityStatus
        }
    }

    # Console display
    $colWidth = @(38, 10, 10, 12)
    Write-Host ("  {0,-38} {1,10} {2,10} {3,12}" -f "SKU Name", "Assigned", "Total", "Available") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 74)) -ForegroundColor $C.Info

    foreach ($sku in ($inventory | Sort-Object FriendlyName)) {
        $availColor = if ($sku.Available -lt 0) { $C.Error }
                      elseif ($sku.Available -le 5) { $C.Warning }
                      else { $C.Success }
        $statusTag  = if ($sku.CapabilityStatus -ne 'Enabled') { " [$($sku.CapabilityStatus)]" } else { '' }
        Write-Host ("  {0,-38} {1,10} {2,10} {3,12}" -f `
            ($sku.FriendlyName + $statusTag), $sku.Assigned, $sku.Total, $sku.Available) -ForegroundColor $availColor
    }

    Write-Host ""
    Write-Ok "$($inventory.Count) SKU(s) retrieved."
    return @($inventory)
}

function Get-UnlicensedUsers {
    Write-Section "UNLICENSED USERS"
    Write-Step "Querying enabled users with no license assigned..."
    Write-Info "This uses ConsistencyLevel=eventual — results may be slightly cached."
    Write-Host ""

    try {
        $users = Get-MgUser `
            -Filter "assignedLicenses/`$count eq 0 and accountEnabled eq true" `
            -CountVariable licCount `
            -ConsistencyLevel eventual `
            -All `
            -Property "displayName,userPrincipalName,department,accountEnabled,userType" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Query failed: $_"
        return @()
    }

    # Exclude guests unless they show up as members
    $members = @($users | Where-Object { $_.UserType -ne 'Guest' })

    if ($members.Count -eq 0) {
        Write-Ok "No unlicensed enabled member accounts found."
        return @()
    }

    $result = foreach ($u in ($members | Sort-Object DisplayName)) {
        [PSCustomObject]@{
            DisplayName    = $u.DisplayName
            UPN            = $u.UserPrincipalName
            Department     = if ($u.Department) { $u.Department } else { '—' }
            AccountEnabled = $u.AccountEnabled
        }
    }

    # Console display
    Write-Host ("  {0,-32} {1,-40} {2,-22}" -f "Display Name", "UPN", "Department") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 96)) -ForegroundColor $C.Info
    foreach ($u in $result) {
        Write-Host ("  {0,-32} {1,-40} {2,-22}" -f $u.DisplayName, $u.UPN, $u.Department) -ForegroundColor $C.Warning
    }

    Write-Host ""
    Write-Warn "$($result.Count) unlicensed enabled user(s) found."
    return @($result)
}

function Get-InactiveUsers {
    Write-Section "INACTIVE USERS (90+ DAYS)"
    Write-Step "Retrieving all users with sign-in activity..."
    Write-Info "This may take a moment depending on tenant size. Requires AuditLog.Read.All."
    Write-Host ""

    $thresholdDate = (Get-Date).AddDays(-90)
    $allUsers      = $null

    try {
        $allUsers = Get-MgUser -All `
            -Property "displayName,userPrincipalName,department,accountEnabled,assignedLicenses,signInActivity,userType" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        Write-Info "Ensure AuditLog.Read.All is consented and the account has sufficient permissions."
        return @()
    }

    $members = @($allUsers | Where-Object { $_.UserType -ne 'Guest' -and $_.AccountEnabled -eq $true })

    $inactive = foreach ($u in $members) {
        $lastSignIn = $null
        $daysInactive = $null

        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn   = [datetime]$u.SignInActivity.LastSignInDateTime
            $daysInactive = [math]::Round(((Get-Date) - $lastSignIn).TotalDays, 0)
        }

        $isInactive = (-not $lastSignIn) -or ($lastSignIn -lt $thresholdDate)

        if ($isInactive) {
            [PSCustomObject]@{
                DisplayName   = $u.DisplayName
                UPN           = $u.UserPrincipalName
                Department    = if ($u.Department) { $u.Department } else { '—' }
                LastSignIn    = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
                DaysInactive  = if ($null -ne $daysInactive) { $daysInactive } else { 'N/A' }
                IsLicensed    = ($u.AssignedLicenses.Count -gt 0)
            }
        }
    }

    $inactive = @($inactive)

    if ($inactive.Count -eq 0) {
        Write-Ok "No inactive users found (all active within 90 days)."
        return @()
    }

    $sorted = $inactive | Sort-Object @{ Expression={ if ($_.DaysInactive -eq 'N/A') { 99999 } else { [int]$_.DaysInactive } }; Descending=$true }

    # Console display
    Write-Host ("  {0,-28} {1,-36} {2,-14} {3,-12}" -f "Display Name", "UPN", "Last Sign-In", "Days Inactive") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 92)) -ForegroundColor $C.Info
    foreach ($u in $sorted) {
        $color = if ($u.LastSignIn -eq 'Never') { $C.Error } else { $C.Warning }
        Write-Host ("  {0,-28} {1,-36} {2,-14} {3,-12}" -f $u.DisplayName, $u.UPN, $u.LastSignIn, $u.DaysInactive) -ForegroundColor $color
    }

    Write-Host ""
    Write-Warn "$($inactive.Count) inactive user(s) found (no sign-in within 90 days or never)."
    return @($inactive)
}

function Get-MfaStatus {
    Write-Section "MFA REGISTRATION STATUS"
    Write-Warn "This check calls the API once per user — it may be slow on large tenants."
    Write-Step "Retrieving enabled member users..."
    Write-Host ""

    try {
        $allUsers = Get-MgUser -All `
            -Filter "accountEnabled eq true" `
            -Property "displayName,userPrincipalName,department,userType,id" `
            -ErrorAction Stop
    } catch {
        Write-Fail "Failed to retrieve users: $_"
        return @()
    }

    $members = @($allUsers | Where-Object { $_.UserType -ne 'Guest' })

    if ($members.Count -eq 0) {
        Write-Warn "No enabled member users found."
        return @()
    }

    Write-Step "Checking authentication methods for $($members.Count) user(s)..."
    Write-Info "Progress is shown below. Guests and service accounts are skipped on error."
    Write-Host ""

    $noMfa    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total    = $members.Count
    $current  = 0
    $skipped  = 0

    foreach ($u in $members) {
        $current++

        # Progress every 10 users or at key milestones
        if ($current -eq 1 -or ($current % 10) -eq 0 -or $current -eq $total) {
            $pct = [math]::Round(($current / $total) * 100, 0)
            Write-Host ("`r  [*] Checking user {0}/{1} ({2}%)..." -f $current, $total, $pct) -ForegroundColor $C.Progress -NoNewline
        }

        try {
            $methods = Get-MgUserAuthenticationMethod -UserId $u.Id -ErrorAction Stop
            # Filter out the default "password" method — only count real MFA factors
            $mfaMethods = @($methods | Where-Object {
                $_.AdditionalProperties['@odata.type'] -notmatch 'passwordAuthentication'
            })

            if ($mfaMethods.Count -eq 0) {
                $noMfa.Add([PSCustomObject]@{
                    DisplayName = $u.DisplayName
                    UPN         = $u.UserPrincipalName
                    Department  = if ($u.Department) { $u.Department } else { '—' }
                })
            }
        } catch {
            $skipped++
        }
    }

    Write-Host ""  # end the progress line
    Write-Host ""

    if ($skipped -gt 0) {
        Write-Warn "$skipped user(s) skipped (guests, service accounts, or permission errors)."
    }

    if ($noMfa.Count -eq 0) {
        Write-Ok "All checked users have at least one MFA method registered."
        return @()
    }

    # Console display
    Write-Host ("  {0,-32} {1,-42} {2,-22}" -f "Display Name", "UPN", "Department") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 98)) -ForegroundColor $C.Info
    foreach ($u in ($noMfa | Sort-Object DisplayName)) {
        Write-Host ("  {0,-32} {1,-42} {2,-22}" -f $u.DisplayName, $u.UPN, $u.Department) -ForegroundColor $C.Error
    }

    Write-Host ""
    Write-Fail "$($noMfa.Count) user(s) have NO MFA method registered."
    return @($noMfa)
}

function Show-SharedMailboxGuidance {
    Write-Section "SHARED MAILBOX AUDIT"
    Write-Warn "Shared mailbox auditing requires the Exchange Online Management module."
    Write-Host ""
    Write-Info "A.L.M.A.N.A.C. connects to Microsoft Graph, which does not expose mailbox"
    Write-Info "type details (Shared vs. User) in the same way as Exchange Online."
    Write-Host ""
    Write-Info "To audit shared mailboxes, run the following in a separate session:"
    Write-Host ""
    Write-Host "    # Install module if needed:" -ForegroundColor $C.Info
    Write-Host "    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # Connect:" -ForegroundColor $C.Info
    Write-Host "    Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # List shared mailboxes:" -ForegroundColor $C.Info
    Write-Host "    Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress,ProhibitSendQuota" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # Find shared mailboxes with a license assigned (unnecessary cost):" -ForegroundColor $C.Info
    Write-Host "    Get-Mailbox -RecipientTypeDetails SharedMailbox | Where-Object { `$_.SKUAssigned -eq `$true }" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "    # Check full-access permissions on all shared mailboxes:" -ForegroundColor $C.Info
    Write-Host "    Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-MailboxPermission | Where-Object { `$_.User -notlike '*SELF*' }" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Info "Shared mailboxes under 50 GB do not require a license. Larger mailboxes"
    Write-Info "require at least an Exchange Online Plan 2 or M365 E3 license."
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$LicenseData,
        [array]$UnlicensedData,
        [array]$InactiveData,
        [array]$NoMfaData
    )

    $reportDate    = Get-Date -Format "MMMM d, yyyy HH:mm"
    $tenantDisplay = if ($script:TenantDomain) { EscHtml $script:TenantDomain } else { EscHtml $script:ConnectedAs }
    $connectedAs   = EscHtml $script:ConnectedAs

    $totalUsers      = 0
    try {
        $totalUsers = (Get-MgUser -Filter "accountEnabled eq true and userType eq 'Member'" -CountVariable x -ConsistencyLevel eventual -All -Property "id" -ErrorAction SilentlyContinue | Measure-Object).Count
    } catch {
        Write-Warn "Could not retrieve total user count — licensed count will show as '—'."
    }
    $licensedCount   = if ($totalUsers -gt 0) { $totalUsers - $UnlicensedData.Count } else { '—' }

    # ── License rows ─────────────────────────────────────────────────────────
    $licRows = [System.Text.StringBuilder]::new()
    foreach ($sku in ($LicenseData | Sort-Object FriendlyName)) {
        $availClass = if ($sku.Available -lt 0) { 'tk-badge-err' }
                      elseif ($sku.Available -le 5) { 'tk-badge-warn' }
                      else { 'tk-badge-ok' }
        $statusCell = if ($sku.CapabilityStatus -ne 'Enabled') {
            "<span class='tk-badge-warn'>$(EscHtml $sku.CapabilityStatus)</span>"
        } else {
            "<span class='tk-badge-ok'>Active</span>"
        }
        [void]$licRows.Append("<tr><td>$(EscHtml $sku.FriendlyName)</td><td><code>$(EscHtml $sku.SkuPartNumber)</code></td><td>$($sku.Assigned)</td><td>$($sku.Total)</td><td><span class='$availClass'>$($sku.Available)</span></td><td>$statusCell</td></tr>`n")
    }

    # ── Unlicensed rows ───────────────────────────────────────────────────────
    $unLicRows = [System.Text.StringBuilder]::new()
    if ($UnlicensedData.Count -eq 0) {
        [void]$unLicRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No unlicensed enabled users found.</td></tr>")
    } else {
        foreach ($u in ($UnlicensedData | Sort-Object DisplayName)) {
            [void]$unLicRows.Append("<tr><td>$(EscHtml $u.DisplayName)</td><td>$(EscHtml $u.UPN)</td><td>$(EscHtml $u.Department)</td><td><span class='tk-badge-ok'>Enabled</span></td></tr>`n")
        }
    }

    # ── Inactive rows ─────────────────────────────────────────────────────────
    $inactRows = [System.Text.StringBuilder]::new()
    if ($InactiveData.Count -eq 0) {
        [void]$inactRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No inactive users found.</td></tr>")
    } else {
        $sorted = $InactiveData | Sort-Object @{ Expression={ if ($_.DaysInactive -eq 'N/A') { 99999 } else { [int]$_.DaysInactive } }; Descending=$true }
        foreach ($u in $sorted) {
            $badgeClass = if ($u.LastSignIn -eq 'Never') { 'tk-badge-err' } else { 'tk-badge-warn' }
            [void]$inactRows.Append("<tr><td>$(EscHtml $u.DisplayName)</td><td>$(EscHtml $u.UPN)</td><td><span class='$badgeClass'>$(EscHtml $u.LastSignIn)</span></td><td>$($u.DaysInactive)</td></tr>`n")
        }
    }

    # ── No-MFA rows ───────────────────────────────────────────────────────────
    $mfaRows = [System.Text.StringBuilder]::new()
    if ($NoMfaData.Count -eq 0) {
        [void]$mfaRows.Append("<tr><td colspan='3' class='tk-badge-ok' style='text-align:center;'>All users have MFA registered.</td></tr>")
    } else {
        foreach ($u in ($NoMfaData | Sort-Object DisplayName)) {
            [void]$mfaRows.Append("<tr><td>$(EscHtml $u.DisplayName)</td><td>$(EscHtml $u.UPN)</td><td>$(EscHtml $u.Department)</td></tr>`n")
        }
    }

    # ── Summary card colors ───────────────────────────────────────────────────
    $unlicClass  = if ($UnlicensedData.Count -gt 0) { 'warn' } else { 'ok' }
    $inactClass  = if ($InactiveData.Count   -gt 0) { 'warn' } else { 'ok' }
    $mfaClass    = if ($NoMfaData.Count      -gt 0) { 'err'  } else { 'ok' }

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'A.L.M.A.N.A.C. Audit Report' `
        -ScriptName 'A.L.M.A.N.A.C.' `
        -Subtitle   "${orgPrefix}Microsoft 365 License &amp; User Security Audit -- $tenantDisplay" `
        -MetaItems  ([ordered]@{
            'Generated'    = $reportDate
            'Connected As' = $connectedAs
            'Tenant'       = $tenantDisplay
        }) `
        -NavItems   @(
            'License Inventory',
            'Unlicensed Users',
            'Inactive Users',
            'MFA Status',
            'Shared Mailbox'
        )

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'A.L.M.A.N.A.C. v6.0'

    $html = $htmlHead + @"

  <!-- Summary -->
  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalUsers</div><div class="tk-summary-lbl">Total Enabled Users</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$licensedCount</div><div class="tk-summary-lbl">Licensed Users</div></div>
    <div class="tk-summary-card $unlicClass"><div class="tk-summary-num">$($UnlicensedData.Count)</div><div class="tk-summary-lbl">Unlicensed Users</div></div>
    <div class="tk-summary-card $inactClass"><div class="tk-summary-num">$($InactiveData.Count)</div><div class="tk-summary-lbl">Inactive (90d+)</div></div>
    <div class="tk-summary-card $mfaClass"><div class="tk-summary-num">$($NoMfaData.Count)</div><div class="tk-summary-lbl">No MFA Registered</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($LicenseData.Count)</div><div class="tk-summary-lbl">License SKUs</div></div>
  </div>

  <!-- License Inventory -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">License SKU Inventory</span>
      <span class="tk-section-num">$($LicenseData.Count) SKU(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead>
          <tr><th>SKU Name</th><th>Part Number</th><th>Assigned</th><th>Total</th><th>Available</th><th>Status</th></tr>
        </thead>
        <tbody>
          $($licRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- Unlicensed Users -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Unlicensed Enabled Users</span>
      <span class="tk-section-num">$($UnlicensedData.Count) user(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead>
          <tr><th>Display Name</th><th>UPN</th><th>Department</th><th>Account Status</th></tr>
        </thead>
        <tbody>
          $($unLicRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- Inactive Users -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Inactive Users -- No Sign-In Within 90 Days</span>
      <span class="tk-section-num">$($InactiveData.Count) user(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead>
          <tr><th>Display Name</th><th>UPN</th><th>Last Sign-In</th><th>Days Inactive</th></tr>
        </thead>
        <tbody>
          $($inactRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- MFA Status -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Users Without MFA Registered</span>
      <span class="tk-section-num">$($NoMfaData.Count) user(s)</span>
    </div>
    <div class="tk-card">
      <table class="tk-table">
        <thead>
          <tr><th>Display Name</th><th>UPN</th><th>Department</th></tr>
        </thead>
        <tbody>
          $($mfaRows.ToString())
        </tbody>
      </table>
    </div>
  </div>

  <!-- Shared Mailbox Note -->
  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Shared Mailbox Audit</span>
      <span class="tk-section-num">Requires Exchange Online</span>
    </div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Note</span>
        <strong>Shared mailbox auditing requires the Exchange Online Management module.</strong><br><br>
        Run the following commands in a separate PowerShell session:<br><br>
        <code>Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force</code><br>
        <code>Connect-ExchangeOnline -UserPrincipalName admin@$tenantDisplay</code><br>
        <code>Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress,ProhibitSendQuota</code><br><br>
        <strong>Key checks:</strong> Shared mailboxes under 50 GB do not require a license. Verify
        that no shared mailboxes have unnecessary licenses assigned, and audit full-access
        permissions to ensure only authorised users have access.
      </div>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-HtmlReport {
    param(
        [array]$LicenseData,
        [array]$UnlicensedData,
        [array]$InactiveData,
        [array]$NoMfaData
    )

    Write-Section "EXPORTING HTML REPORT"

    # Run any missing collections
    if ($null -eq $LicenseData)   { $LicenseData   = Get-LicenseInventory }
    if ($null -eq $UnlicensedData){ $UnlicensedData = Get-UnlicensedUsers  }
    if ($null -eq $InactiveData)  { $InactiveData   = Get-InactiveUsers    }
    if ($null -eq $NoMfaData)     { $NoMfaData      = Get-MfaStatus        }

    Write-Step "Building HTML report..."

    $html      = Build-HtmlReport -LicenseData $LicenseData -UnlicensedData $UnlicensedData -InactiveData $InactiveData -NoMfaData $NoMfaData
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "ALMANAC_${timestamp}.html"

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Failed to save report: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-AlmanacBanner

    $connStatus = if ($script:Connected) {
        "  Connected as : $($script:ConnectedAs)"
    } else {
        "  Not Connected — select option 1 to authenticate"
    }
    $connColor = if ($script:Connected) { $C.Success } else { $C.Warning }

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host $connStatus -ForegroundColor $connColor
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  1  Connect to Microsoft 365 (Graph authentication)" -ForegroundColor $C.Info
    Write-Host "  2  License inventory — SKUs, assigned vs. total" -ForegroundColor $C.Info
    Write-Host "  3  Unlicensed users — enabled accounts with no license" -ForegroundColor $C.Info
    Write-Host "  4  Inactive users — no sign-in in 90+ days" -ForegroundColor $C.Info
    Write-Host "  5  MFA status — users with no MFA method registered" -ForegroundColor $C.Warning
    Write-Host "  6  Shared mailbox audit guidance (Exchange Online)" -ForegroundColor $C.Info
    Write-Host "  7  Export full HTML report (all findings)" -ForegroundColor $C.Success
    Write-Host "  Q  Quit" -ForegroundColor $C.Info
    Write-Host ""
}

function Assert-Connected {
    if (-not $script:Connected) {
        Write-Host ""
        Write-Warn "Not connected to Microsoft 365. Please select option 1 first."
        Write-Host ""
        Read-Host "  Press Enter to return to menu"
        return $false
    }
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

Show-AlmanacBanner
Install-GraphModule

# Check if already connected (e.g. token still cached from prior session)
if (Test-GraphConnection) {
    Write-Ok "Already connected as $($script:ConnectedAs)"
    Write-Host ""
}

# Unattended mode: connect then run full report and exit
if ($Unattended) {
    Write-Section "UNATTENDED MODE"
    Write-Step "Attempting to connect using cached credentials..."

    if (-not $script:Connected) {
        Invoke-Connect
    }

    if (-not $script:Connected) {
        Write-Fail "Could not establish a Graph connection in unattended mode. Exiting."
        exit 1
    }

    $licData       = Get-LicenseInventory
    $unlicData     = Get-UnlicensedUsers
    $inactiveData  = Get-InactiveUsers
    $noMfaData     = Get-MfaStatus

    Export-HtmlReport `
        -LicenseData   $licData `
        -UnlicensedData $unlicData `
        -InactiveData  $inactiveData `
        -NoMfaData     $noMfaData

    Write-Host ""
    Write-Ok "Unattended audit complete."
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE MENU LOOP
# ─────────────────────────────────────────────────────────────────────────────

# Cached audit data — reset per session so menu re-runs pull fresh data
$cachedLicense    = $null
$cachedUnlicensed = $null
$cachedInactive   = $null
$cachedNoMfa      = $null

do {
    Show-Menu
    $choice = (Read-Host "  Select option").Trim().ToUpper()

    switch ($choice) {

        '1' {
            Invoke-Connect
            # Reset cache on reconnect
            $cachedLicense    = $null
            $cachedUnlicensed = $null
            $cachedInactive   = $null
            $cachedNoMfa      = $null
            Read-Host "  Press Enter to return to menu"
        }

        '2' {
            if (-not (Assert-Connected)) { break }
            $cachedLicense = Get-LicenseInventory
            Read-Host "  Press Enter to return to menu"
        }

        '3' {
            if (-not (Assert-Connected)) { break }
            $cachedUnlicensed = Get-UnlicensedUsers
            Read-Host "  Press Enter to return to menu"
        }

        '4' {
            if (-not (Assert-Connected)) { break }
            $cachedInactive = Get-InactiveUsers
            Read-Host "  Press Enter to return to menu"
        }

        '5' {
            if (-not (Assert-Connected)) { break }
            $cachedNoMfa = Get-MfaStatus
            Read-Host "  Press Enter to return to menu"
        }

        '6' {
            Show-SharedMailboxGuidance
            Read-Host "  Press Enter to return to menu"
        }

        '7' {
            if (-not (Assert-Connected)) { break }
            Export-HtmlReport `
                -LicenseData    $cachedLicense `
                -UnlicensedData $cachedUnlicensed `
                -InactiveData   $cachedInactive `
                -NoMfaData      $cachedNoMfa
            Read-Host "  Press Enter to return to menu"
        }

        'Q' {
            Write-Host ""
            Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor $C.Progress
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
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
