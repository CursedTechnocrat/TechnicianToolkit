# totem.ps1 - T.O.T.E.M. — Trusted Observer of Transparent Execution Modules
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
    T.O.T.E.M. — Trusted Observer of Transparent Execution Modules
    TPM Health Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits the state of the Trusted Platform Module -- presence,
    specification version, manufacturer and firmware, ownership, clear
    state, attestation readiness -- and cross-references against
    BitLocker to confirm which volumes actually depend on the TPM's
    protector chain. Produces a dark-themed HTML report with a red /
    yellow / green readiness verdict suitable for a Windows 11 /
    Autopilot / BitLocker readiness gate.

.USAGE
    PS C:\> .\totem.ps1                    # Interactive run
    PS C:\> .\totem.ps1 -Unattended        # Silent: export HTML and exit

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

function Show-TotemBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  T.O.T.E.M. — Trusted Observer of Transparent Execution Modules" -ForegroundColor Cyan
    Write-Host "  TPM Health Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Collectors, verdict, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — TPM STATUS
# ─────────────────────────────────────────────────────────────────────────────

function Get-TpmStatus {
    # Get-Tpm is the canonical way to read TPM state on Windows 10/11. It's in
    # the TrustedPlatformModule module which is present in-box.
    $tpm = $null
    try {
        $tpm = Get-Tpm -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            Present          = $false
            Ready            = $false
            Enabled          = $false
            Activated        = $false
            Owned            = $false
            ManufacturerId   = $null
            ManufacturerName = $null
            ManufacturerVersion = $null
            SpecVersion      = $null
            PhysicalPresence = $null
            AutoProvisioning = $null
            RestartPending   = $false
            CollectorError   = $_.Exception.Message
        }
    }

    return [PSCustomObject]@{
        Present             = [bool]$tpm.TpmPresent
        Ready               = [bool]$tpm.TpmReady
        Enabled             = [bool]$tpm.TpmEnabled
        Activated           = [bool]$tpm.TpmActivated
        Owned               = [bool]$tpm.TpmOwned
        ManufacturerId      = $tpm.ManufacturerId
        ManufacturerName    = $tpm.ManufacturerIdTxt
        ManufacturerVersion = $tpm.ManufacturerVersion
        SpecVersion         = $tpm.SpecVersion
        PhysicalPresence    = $tpm.PhysicalPresenceVersionInfo
        AutoProvisioning    = if ($null -ne $tpm.AutoProvisioning) { $tpm.AutoProvisioning.ToString() } else { 'Unknown' }
        RestartPending      = [bool]$tpm.RestartPending
        CollectorError      = $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — SPEC VERSION + MANUFACTURER PARSING
# ─────────────────────────────────────────────────────────────────────────────

function Get-TpmVersionSummary {
    param($Tpm)

    # SpecVersion is a comma-separated string like "2.0, 0, 1.38". First token
    # is the TPM spec. Second is the family level (rarely useful). Third is
    # the rev of the TPM library spec.
    $spec = 'Unknown'
    if ($Tpm.SpecVersion) {
        $first = ($Tpm.SpecVersion -split ',')[0].Trim()
        if ($first) { $spec = $first }
    }

    $label = switch ($spec) {
        '2.0' { 'TPM 2.0' }
        '1.2' { 'TPM 1.2' }
        default { "TPM $spec" }
    }

    return [PSCustomObject]@{
        SpecVersion   = $spec
        Label         = $label
        IsWin11Ready  = ($spec -eq '2.0')
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — BITLOCKER DEPENDENCY
# ─────────────────────────────────────────────────────────────────────────────

# Answers "if this TPM goes bad, which drives can't unlock?" -- any volume with
# a Tpm, TpmPin, TpmPinStartupKey, or RecoveryPassword protector chain hangs
# on TPM measurements.
function Get-BitLockerTpmDependency {
    $result = [System.Collections.Generic.List[object]]::new()

    try {
        $vols = Get-BitLockerVolume -ErrorAction Stop
    } catch {
        return @($result)
    }

    foreach ($v in $vols) {
        $tpmProtectors = @($v.KeyProtector | Where-Object {
            $_.KeyProtectorType -in @('Tpm','TpmPin','TpmPinStartupKey','TpmStartupKey')
        })
        $result.Add([PSCustomObject]@{
            MountPoint        = $v.MountPoint
            VolumeType        = $v.VolumeType
            ProtectionStatus  = $v.ProtectionStatus
            EncryptionMethod  = $v.EncryptionMethod
            Protectors        = (@($v.KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -join ', ')
            TpmProtectorCount = $tpmProtectors.Count
            DependsOnTpm      = ($tpmProtectors.Count -gt 0)
        })
    }

    return @($result)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — ATTESTATION / EK INFO
# ─────────────────────────────────────────────────────────────────────────────

function Get-TpmAttestationInfo {
    $result = [PSCustomObject]@{
        EkPresent      = $null
        EkAlgorithms   = $null
        CollectorError = $null
    }

    try {
        $ek = Get-TpmEndorsementKeyInfo -ErrorAction Stop
        if ($ek) {
            $result.EkPresent    = $true
            # EK has manufacturer certificates plus RSA / ECC algorithm slots;
            # caller cares about presence, not the cert chain itself.
            $algos = @()
            if ($ek.IsPresent)      { $algos += 'EK present' }
            if ($ek.ManufacturerCertificates -and $ek.ManufacturerCertificates.Count -gt 0) {
                $algos += "$($ek.ManufacturerCertificates.Count) manufacturer cert(s)"
            }
            $result.EkAlgorithms = ($algos -join '; ')
        } else {
            $result.EkPresent = $false
        }
    } catch {
        $result.EkPresent      = $null
        $result.CollectorError = $_.Exception.Message
    }

    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-Verdict {
    param($Tpm, $Version, [array]$BitLockerDependencies, $Attestation)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if (-not $Tpm.Present) {
        $issues.Add('No TPM is present on this machine — Windows 11 upgrade and BitLocker key protection are blocked.')
    } else {
        if (-not $Tpm.Enabled)   { $issues.Add('TPM is present but DISABLED in firmware — enable in BIOS/UEFI Security settings.') }
        if (-not $Tpm.Activated) { $issues.Add('TPM is DEACTIVATED — activate in BIOS/UEFI and take ownership.') }
        if (-not $Tpm.Ready)     { $warns.Add('TPM reports NOT READY — may need Initialize-Tpm to provision, or a clear-and-reprovision cycle.') }
        if (-not $Tpm.Owned)     { $warns.Add('TPM has no owner — BitLocker provisioning will handle this on first encrypt, but document the state before proceeding.') }
        if ($Tpm.RestartPending) { $warns.Add('TPM reports a pending restart — commit outstanding state changes before audit-time conclusions.') }
        if (-not $Version.IsWin11Ready) {
            $warns.Add("This machine has $($Version.Label), not TPM 2.0 — Windows 11 upgrade is blocked until firmware offers a 2.0-capable chip (vendor BIOS update or dTPM->fTPM toggle may help).")
        }
    }

    # BitLocker cross-check
    $dependent = @($BitLockerDependencies | Where-Object { $_.DependsOnTpm })
    $anyBitLocker = @($BitLockerDependencies | Where-Object { $_.ProtectionStatus -eq 'On' }).Count
    if ($Tpm.Present -and $Tpm.Ready -and $dependent.Count -eq 0 -and $anyBitLocker -gt 0) {
        $warns.Add('BitLocker is enabled but none of the protected volumes use a TPM-based key protector — consider adding TPM for seamless unlock.')
    }

    if ($null -eq $Attestation.EkPresent -and $Tpm.Present) {
        $warns.Add('Endorsement Key info could not be read — Autopilot device pre-registration and attested boot will fail until the EK is retrievable.')
    }

    $verdict = if ($issues.Count -gt 0) { 'NOT READY' }
               elseif ($warns.Count -gt 0) { 'READY WITH WARNINGS' }
               else { 'READY' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $class
        Issues  = @($issues)
        Warns   = @($warns)
        BitLockerVolumesDependingOnTpm = $dependent.Count
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param($Tpm, $Version, [array]$BitLocker, $Attestation, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    # TPM status badges
    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-err'>No</span>" } }

    $specClass = if ($Version.IsWin11Ready) { 'tk-badge-ok' } else { 'tk-badge-warn' }

    $blRows = [System.Text.StringBuilder]::new()
    if ($BitLocker.Count -eq 0) {
        [void]$blRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No BitLocker-managed volumes found.</td></tr>")
    } else {
        foreach ($v in $BitLocker) {
            $protState = switch ($v.ProtectionStatus) {
                'On'  { "<span class='tk-badge-ok'>On</span>" }
                'Off' { "<span class='tk-badge-warn'>Off</span>" }
                default { "<span class='tk-badge-info'>$(EscHtml $v.ProtectionStatus)</span>" }
            }
            $tpmDep = if ($v.DependsOnTpm) { "<span class='tk-badge-ok'>Yes ($($v.TpmProtectorCount))</span>" } else { "<span class='tk-badge-warn'>No</span>" }
            [void]$blRows.Append("<tr><td>$(EscHtml $v.MountPoint)</td><td>$(EscHtml $v.VolumeType)</td><td>$protState</td><td>$(EscHtml $v.Protectors)</td><td>$tpmDep</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>TPM posture is clean for Windows 11, BitLocker, and Autopilot.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'T.O.T.E.M. TPM Health Report' `
        -ScriptName 'T.O.T.E.M.' `
        -Subtitle   "${orgPrefix}TPM Health Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'     = $machine
            'Generated'   = $reportDate
            'TPM Spec'    = $Version.Label
            'Verdict'     = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'TPM Status', 'BitLocker Dependency', 'Attestation')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'T.O.T.E.M. v6.0'

    $collectorErrCell = if ($Tpm.CollectorError) {
        "<tr><th>Collector Error</th><td><span class='tk-badge-err'>$(EscHtml $Tpm.CollectorError)</span></td></tr>"
    } else { '' }

    $attErrCell = if ($Attestation.CollectorError) {
        "<tr><th>EK Read Error</th><td><span class='tk-badge-warn'>$(EscHtml $Attestation.CollectorError)</span></td></tr>"
    } else { '' }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">TPM Readiness</div></div>
    <div class="tk-summary-card $(if ($Tpm.Present) { 'ok' } else { 'err' })"><div class="tk-summary-num">$(if ($Tpm.Present) { 'Present' } else { 'Absent' })</div><div class="tk-summary-lbl">TPM Hardware</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $Version.Label)</div><div class="tk-summary-lbl">Specification</div></div>
    <div class="tk-summary-card $(if ($Tpm.Ready) { 'ok' } else { 'warn' })"><div class="tk-summary-num">$(if ($Tpm.Ready) { 'Yes' } else { 'No' })</div><div class="tk-summary-lbl">TPM Ready</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Verdict.BitLockerVolumesDependingOnTpm)</div><div class="tk-summary-lbl">BitLocker Volumes Using TPM</div></div>
    <div class="tk-summary-card $(if ($Attestation.EkPresent) { 'ok' } elseif ($null -eq $Attestation.EkPresent) { 'warn' } else { 'err' })"><div class="tk-summary-num">$(if ($Attestation.EkPresent) { 'Yes' } elseif ($null -eq $Attestation.EkPresent) { 'Unknown' } else { 'No' })</div><div class="tk-summary-lbl">Endorsement Key</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">TPM Status</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Present</th><td>$(_yn $Tpm.Present)</td></tr>
          <tr><th>Enabled (firmware)</th><td>$(_yn $Tpm.Enabled)</td></tr>
          <tr><th>Activated</th><td>$(_yn $Tpm.Activated)</td></tr>
          <tr><th>Ready for Use</th><td>$(_yn $Tpm.Ready)</td></tr>
          <tr><th>Owned</th><td>$(_yn $Tpm.Owned)</td></tr>
          <tr><th>Specification</th><td><span class='$specClass'>$(EscHtml $Version.Label)</span></td></tr>
          <tr><th>Manufacturer (raw ID)</th><td><code>$(EscHtml $Tpm.ManufacturerId)</code></td></tr>
          <tr><th>Manufacturer (text)</th><td>$(EscHtml $Tpm.ManufacturerName)</td></tr>
          <tr><th>Manufacturer Version</th><td><code>$(EscHtml $Tpm.ManufacturerVersion)</code></td></tr>
          <tr><th>Physical Presence Version</th><td>$(EscHtml $Tpm.PhysicalPresence)</td></tr>
          <tr><th>Auto-Provisioning</th><td>$(EscHtml $Tpm.AutoProvisioning)</td></tr>
          <tr><th>Restart Pending</th><td>$(_yn $Tpm.RestartPending)</td></tr>
          $collectorErrCell
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">BitLocker Dependency on TPM</span><span class="tk-section-num">$($BitLocker.Count) volume(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Mount</th><th>Volume Type</th><th>Protection</th><th>Protectors</th><th>Uses TPM?</th></tr></thead>
        <tbody>$($blRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Attestation / Endorsement Key</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>EK Present</th><td>$(if ($Attestation.EkPresent) { "<span class='tk-badge-ok'>Yes</span>" } elseif ($null -eq $Attestation.EkPresent) { "<span class='tk-badge-warn'>Unknown</span>" } else { "<span class='tk-badge-err'>No</span>" })</td></tr>
          <tr><th>EK Details</th><td>$(EscHtml $Attestation.EkAlgorithms)</td></tr>
          $attErrCell
        </tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-TotemBanner

Write-Section "TPM STATUS"
$tpm = Get-TpmStatus
if ($tpm.CollectorError) {
    Write-Fail "Get-Tpm failed: $($tpm.CollectorError)"
} else {
    $pColor = if ($tpm.Present) { $C.Success } else { $C.Error }
    Write-Host ("  Present   : {0}" -f $tpm.Present) -ForegroundColor $pColor
    if ($tpm.Present) {
        Write-Host ("  Enabled   : {0}" -f $tpm.Enabled) -ForegroundColor $(if ($tpm.Enabled) { $C.Success } else { $C.Error })
        Write-Host ("  Activated : {0}" -f $tpm.Activated) -ForegroundColor $(if ($tpm.Activated) { $C.Success } else { $C.Error })
        Write-Host ("  Ready     : {0}" -f $tpm.Ready) -ForegroundColor $(if ($tpm.Ready) { $C.Success } else { $C.Warning })
        Write-Host ("  Owned     : {0}" -f $tpm.Owned) -ForegroundColor $(if ($tpm.Owned) { $C.Success } else { $C.Warning })
        Write-Host ("  Mfr       : {0}" -f $tpm.ManufacturerName) -ForegroundColor $C.Info
    }
}
Write-Host ""

$version = Get-TpmVersionSummary -Tpm $tpm
Write-Section "SPECIFICATION"
$vColor = if ($version.IsWin11Ready) { $C.Success } else { $C.Warning }
Write-Host ("  {0}  (raw: {1})" -f $version.Label, $version.SpecVersion) -ForegroundColor $vColor
Write-Host ""

Write-Section "BITLOCKER DEPENDENCY"
$bitLocker = Get-BitLockerTpmDependency
if ($bitLocker.Count -eq 0) {
    Write-Warn "No BitLocker volumes found (or BitLocker module unavailable)."
} else {
    foreach ($v in $bitLocker) {
        $depLabel = if ($v.DependsOnTpm) { "uses TPM ($($v.TpmProtectorCount))" } else { 'no TPM protector' }
        $color = if ($v.ProtectionStatus -eq 'On' -and $v.DependsOnTpm) { $C.Success }
                 elseif ($v.ProtectionStatus -eq 'On') { $C.Warning }
                 else { $C.Info }
        Write-Host ("  {0,-4} {1,-16} Protection: {2,-3}  Protectors: {3}  ({4})" -f $v.MountPoint, $v.VolumeType, $v.ProtectionStatus, $v.Protectors, $depLabel) -ForegroundColor $color
    }
}
Write-Host ""

Write-Section "ATTESTATION / ENDORSEMENT KEY"
$attestation = Get-TpmAttestationInfo
if ($attestation.EkPresent) {
    Write-Ok "Endorsement Key info readable: $($attestation.EkAlgorithms)"
} elseif ($null -eq $attestation.EkPresent) {
    Write-Warn "Could not read EK info: $($attestation.CollectorError)"
} else {
    Write-Fail "No Endorsement Key data available from TPM."
}
Write-Host ""

$verdict = Get-Verdict -Tpm $tpm -Version $version -BitLockerDependencies $bitLocker -Attestation $attestation

Write-Section "TPM READINESS VERDICT"
$vColorOut = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $vColorOut
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -Tpm $tpm -Version $version -BitLocker $bitLocker -Attestation $attestation -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "TOTEM_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
