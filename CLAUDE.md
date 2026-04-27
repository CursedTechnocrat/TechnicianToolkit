# TechnicianToolkit — Developer Guide

## Project Overview

A collection of PowerShell 5.1+ scripts for IT technicians. Each script is a self-contained tool
with a themed acronym name (GRIMOIRE, AUSPEX, REVENANT, etc.). All tools share a common module
(`TechnicianToolkit.psm1`) that provides logging, privilege checks, HTML helpers, and config I/O.

## Repository Layout

```
TechnicianToolkit/
├── TechnicianToolkit.psm1   # Shared module — imported by every tool
├── grimoire.ps1             # Hub launcher — interactive menu for all tools
├── config.json              # Optional runtime config (org name, log dir, webhooks, defaults)
├── hearth.ps1               # Setup wizard — writes config.json
├── <tool>.ps1               # Individual tool scripts
└── tests/
    └── TechnicianToolkit.Tests.ps1   # Pester 5 test suite
```

## Architecture: Shared Module Pattern

Every tool script must follow this initialization pattern at the top (after the param block):

```powershell
# ===========================
# SHARED MODULE BOOTSTRAP
# ===========================
$TKModulePath = Join-Path $PSScriptRoot 'TechnicianToolkit.psm1'
if (-not (Test-Path $TKModulePath)) {
    $TKModuleUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1'
    Write-Host "  [*] Shared module TechnicianToolkit.psm1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $TKModuleUrl -OutFile $TKModulePath -ErrorAction Stop
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($TKModulePath, [ref]$null, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            Remove-Item -Path $TKModulePath -Force -ErrorAction SilentlyContinue
            Write-Host "  [!!] Downloaded module failed syntax validation - file removed." -ForegroundColor Red
            Write-Host "       $($parseErrors[0].Message)" -ForegroundColor Red
            exit 1
        }
        Write-Host "  [+] Module downloaded and verified." -ForegroundColor Green
    } catch {
        Write-Host "  [!!] Could not download TechnicianToolkit.psm1:" -ForegroundColor Red
        Write-Host "       $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Place the module manually next to this script from:" -ForegroundColor Yellow
        Write-Host "       $TKModuleUrl" -ForegroundColor Yellow
        exit 1
    }
}
Import-Module $TKModulePath -Force -ErrorAction Stop
Invoke-AdminElevation -ScriptFile $PSCommandPath
```

The bootstrap ensures a single-file distribution works — drop any tool .ps1 on a
machine and it will pull `TechnicianToolkit.psm1` from GitHub on first run. TLS 1.2
is forced for older Windows builds. `-ErrorAction Stop` on the final `Import-Module`
prevents the silent-partial-execution failure mode (where a missing module used to
let the script continue until it hit an undefined function like `Get-TKHtmlHead`).

`Invoke-AdminElevation` re-launches the script as Administrator if not already elevated.
Scripts that use `Assert-AdminPrivilege` instead will error-exit if not elevated rather
than auto-relaunching — this is appropriate for scripts called programmatically (REVENANT,
HEARTH, ARCHIVE).

## Module Exports

Key functions exported by `TechnicianToolkit.psm1`:

| Function | Purpose |
|----------|---------|
| `Invoke-AdminElevation` | Re-launch as admin if needed (for hub-launched tools) |
| `Assert-AdminPrivilege` | Error-exit if not admin (for directly-called tools) |
| `Test-IsAdmin` | Returns `[bool]` |
| `Get-TKConfig` | Read `config.json`; returns object with defaults if file missing |
| `Set-TKConfig` | Write a key/value into `config.json` (section-aware) |
| `Resolve-LogDirectory` | Return configured log dir or fallback path |
| `Start-TKTranscript` / `Stop-TKTranscript` | PowerShell transcript wrappers |
| `Write-TKError` | Log error to file and optionally POST to Teams webhook |
| `EscHtml` | HTML-escape a string for use in report templates |
| `Get-TKHtmlCss` | Returns the shared `<style>` block — rarely called directly |
| `Get-TKHtmlHead` | Returns `<!DOCTYPE html>…<div class="tk-main">` with shared CSS, page header, and nav bar |
| `Get-TKHtmlFoot` | Returns `</div><footer>…</body></html>` |
| `Write-Section`, `Write-Step`, `Write-Ok`, `Write-Warn`, `Write-Fail`, `Write-Info` | Formatted console output helpers |

### HTML Report Pattern

All tools that produce HTML reports use the shared template helpers:

```powershell
$html  = Get-TKHtmlHead -Title 'Report Title' -ScriptName 'T.O.O.L.' `
             -Subtitle $env:COMPUTERNAME `
             -MetaItems ([ordered]@{ 'Generated' = (Get-Date -Format 'yyyy-MM-dd HH:mm') }) `
             -NavItems @('Section One', 'Section Two')
$html += @"
<div class="tk-section">
  <div class="tk-section-title"><span class="tk-section-num">01</span> Section One</div>
  <div class="tk-card">
    <table class="tk-table"><thead><tr><th>Column</th></tr></thead>
    <tbody><tr><td>Data</td></tr></tbody></table>
  </div>
</div>
"@
$html += Get-TKHtmlFoot -ScriptName 'T.O.O.L. v1.0'
```

Key CSS classes: `.tk-card`, `.tk-card-header`, `.tk-card-label`, `.tk-summary-row`,
`.tk-summary-card` (+ modifier `ok`/`warn`/`err`/`info`), `.tk-section`, `.tk-section-title`,
`.tk-section-num`, `.tk-table`, `.tk-badge-ok/warn/err/info/blue`, `.tk-info-box`, `.tk-info-label`,
`.tk-progress-wrap` + `.tk-progress-bar.ok/warn/err`, `.tk-mono`.

## Running Tests

```powershell
# Install Pester 5 if needed
Install-Module -Name Pester -MinimumVersion 5.0 -Force -SkipPublisherCheck

# Run the suite
Invoke-Pester -Path .\tests\TechnicianToolkit.Tests.ps1 -Output Detailed
```

Tests run without Administrator privileges and without Windows-only APIs, so they work in CI.
The suite covers: `EscHtml`, `Get-TKConfig`/`Set-TKConfig`, `Test-IsAdmin`, `Write-TKError`,
module exports, PowerShell syntax validation on all `.ps1` files, module-import compliance,
param block compliance (`-Unattended`), and GRIMOIRE registry integrity.

## Key Conventions

### Color Schema

Every script defines a local `$ColorSchema` hashtable:

```powershell
$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}
```

### Parameter Conventions

- All interactive tools expose `[switch]$Unattended` — skips prompts, runs defaults.
- Destructive or state-changing tools also expose `[switch]$WhatIf` — previews actions without
  executing them. The current set is REVENANT, ARCHIVE, COVENANT, SIGIL, CLEANSE, CIPHER, FORGE,
  RESTORATION, and RUNEPRESS. GRIMOIRE auto-detects and passes `-WhatIf` to any tool that declares
  it, and the Pester suite (`'-WhatIf declared on destructive tools'`) enforces the list.
- Tools that write logs expose `[switch]$Transcript`.

### Script Header Block

Every script carries a `.SYNOPSIS / .DESCRIPTION / .USAGE / .NOTES` comment block. The
`.NOTES` section holds only the `Version : X.Y` line (bump when changing script behaviour).
Earlier versions embedded a cross-reference `Tools Available` list and a `Color Schema`
legend in every header; those were removed in v3.0 because they drifted out of sync on
every rename. The canonical tool list lives in `grimoire.ps1`'s `$Tools` registry.

### config.json Shape

```json
{
  "OrgName": "",
  "LogDirectory": "",
  "TeamsWebhook": "",
  "Archive": { "DefaultDestination": "" },
  "Revenant": { "DefaultDestination": "" },
  "Covenant": { "DefaultTimezone": "", "DefaultLocalAdminUser": "" }
}
```

`Get-TKConfig` returns these defaults if `config.json` is absent; `Set-TKConfig` creates or
updates the file.

### Adding a New Tool

1. Copy the header block from an existing tool and update acronym, synopsis, version.
2. Add the shared-module bootstrap block (see the initialization pattern above) and the
   appropriate admin check (`Invoke-AdminElevation` or `Assert-AdminPrivilege`). Copy the
   block verbatim from an existing tool — the Pester suite enforces the exact shape.
3. Register the tool in `grimoire.ps1`'s `$Tools` array with a unique numeric `Key`.
4. Add the script's filename to the Quick Launch and Usage sections in `README.md`.
5. The syntax-validation and module-bootstrap compliance Pester tests will cover it automatically.

## Tool Distinctions

### THRESHOLD vs AUGUR

Both tools deal with disk health but cover different layers:

| Tool | Focus |
|------|-------|
| **T.H.R.E.S.H.O.L.D.** | Volume space monitoring — used/free space, low-space alerts, temp cleanup, old profile detection |
| **A.U.G.U.R.** | Physical hardware health — SMART status, wear prediction, failure forecasting, bus/media type |

Run THRESHOLD for "is this drive running out of space?"; run AUGUR for "is this drive about to die?".

### SCRYER vs the single-domain diagnostic tools

S.C.R.Y.E.R. (`scryer.ps1`) is a one-shot consolidated report that rolls five diagnostic passes (system overview, local users, disk space, SMART health, services & tasks) into a single HTML file. It exists for ticket attachments and machine handoffs where one snapshot is more useful than five separate reports.

| Question | Reach for |
|----------|-----------|
| "Give me one file summarising this machine." | **SCRYER** |
| Deep dive on any one of: system health, users, free space, disk reliability, services | AUSPEX / WARD / THRESHOLD / AUGUR / GARGOYLE respectively |

SCRYER's per-section depth is intentionally shallower than the dedicated tools — it samples each domain rather than reproducing the full report.

### RITUAL vs CODEX

Both tools produce a rollup HTML that links out to other tool reports — they answer different questions.

| Question | Reach for |
|----------|-----------|
| "Run an ordered sequence of tools and give me one rollup of the run." | **RITUAL** (executes a recipe, captures status / duration / artifacts per step) |
| "I've already run a bunch of tools ad-hoc — give me one index of what's on disk." | **CODEX** (filesystem scan only, no execution; relative links so the rollup stays clickable when zipped) |

RITUAL produces a record *of an execution* — step status, durations, errors. CODEX produces a record *of a directory* — what reports exist, when, and how big they are. Use RITUAL when you control the run; use CODEX when the reports already exist.
