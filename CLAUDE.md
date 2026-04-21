# TechnicianToolkit — Developer Guide

## Project Overview

A collection of PowerShell 5.1+ scripts for IT technicians. Each script is a self-contained tool
with a themed acronym name (GRIMOIRE, ORACLE, PHANTOM, etc.). All tools share a common module
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
Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Invoke-AdminElevation -ScriptFile $PSCommandPath
```

`Invoke-AdminElevation` re-launches the script as Administrator if not already elevated.
Scripts that use `Assert-AdminPrivilege` instead will error-exit if not elevated rather
than auto-relaunching — this is appropriate for scripts called programmatically (PHANTOM,
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
| `Write-Section`, `Write-Step`, `Write-Ok`, `Write-Warn`, `Write-Fail`, `Write-Info` | Formatted console output helpers |

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
- Destructive tools (PHANTOM, PURGE) also expose `[switch]$WhatIf` — previews actions without
  executing them. GRIMOIRE auto-detects and passes `-WhatIf` to any tool that declares it.
- Tools that write logs expose `[switch]$Transcript`.

### Script Header Block

Every script carries a `.SYNOPSIS / .DESCRIPTION / .USAGE / .NOTES` comment block with:
- `Version : X.Y` — bump minor version when changing script behaviour
- `Tools Available` section — full toolkit list for quick reference
- `Color Schema` section — documents the color conventions above

### config.json Shape

```json
{
  "OrgName": "",
  "LogDirectory": "",
  "TeamsWebhook": "",
  "Archive": { "DefaultDestination": "" },
  "Phantom": { "DefaultDestination": "" },
  "Covenant": { "DefaultTimezone": "", "DefaultLocalAdminUser": "" }
}
```

`Get-TKConfig` returns these defaults if `config.json` is absent; `Set-TKConfig` creates or
updates the file.

### Adding a New Tool

1. Copy the header block from an existing tool and update acronym, synopsis, version.
2. Add `Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force` and the appropriate
   admin check (`Invoke-AdminElevation` or `Assert-AdminPrivilege`).
3. Register the tool in `grimoire.ps1`'s `$Tools` array with a unique numeric `Key`.
4. Add the script's filename to the Quick Launch and Usage sections in `README.md`.
5. The syntax-validation and module-import compliance Pester tests will cover it automatically.

## Tool Distinctions: THRESHOLD vs DWARF

Both tools deal with disk health but cover different layers:

| Tool | Focus |
|------|-------|
| **T.H.R.E.S.H.O.L.D.** | Volume space monitoring — used/free space, low-space alerts, temp cleanup, old profile detection |
| **D.W.A.R.F.** | Physical hardware health — SMART status, wear prediction, failure forecasting, bus/media type |

Run THRESHOLD for "is this drive running out of space?"; run DWARF for "is this drive about to die?".
