# TechnicianToolkit — Developer Guide

## Project Overview

A collection of PowerShell 5.1+ scripts for IT technicians. Each script is a self-contained tool
with a themed acronym name (GRIMOIRE, AUSPEX, REVENANT, etc.). All tools share a common module
(`TechnicianToolkit.psm1`) that provides logging, privilege checks, HTML helpers, and config I/O.

## Two ways to run the suite

Since 5.0 the toolkit ships both as standalone scripts and as one portable
desktop application. **The scripts are the engine; the app drives them.** No tool
logic lives in C#, and porting any there is explicitly out of scope — see
`docs/desktop-port.md`.

The practical consequence when editing: a change to a tool script is a change to
the application too, because the app embeds the scripts verbatim at build time.
The scripts must stay independently runnable under Windows PowerShell 5.1, which
remains the primary documented path — that is why the UTF-8 BOM gate still exists
even though the app hosts PowerShell 7.

## Repository Layout

```
TechnicianToolkit/
├── TechnicianToolkit.psm1   # Shared module — imported by every tool
├── grimoire.ps1             # Hub launcher — interactive menu for all tools
├── config.json              # Optional runtime config (org name, log dir, webhooks, defaults)
├── hearth.ps1               # Setup wizard — writes config.json
├── <tool>.ps1               # Individual tool scripts
├── tests/
│   └── TechnicianToolkit.Tests.ps1   # Pester 5 test suite — guards the scripts
├── app/                     # The desktop application (.NET 8, WPF)
│   ├── TechnicianToolkit.Engine/       # Headless: embeds the suite, hosts PS7,
│   │                                   #   AST readers, runner
│   ├── TechnicianToolkit.Engine.Tests/ # xUnit — guards the C# that reads the scripts
│   ├── TechnicianToolkit.Harness/      # Console front end; the CI gate runs this
│   ├── TechnicianToolkit.App/          # The WPF window
│   └── spike/                          # Phase 00 proof of concept, kept for reference
├── packaging/winget/        # winget manifest source
├── RELEASING.md             # The manual half of a release, including signing
└── docs/desktop-port.md     # The port's plan, decisions and open risks
```

Two test suites, and neither sees the other's regressions. Pester guards the
PowerShell; xUnit guards the C# that parses it. A `param()` block the form builder
misreads is still valid PowerShell, so nothing on the script side would notice.

```powershell
Invoke-Pester -Path .\tests\TechnicianToolkit.Tests.ps1 -Output Detailed
dotnet test app/TechnicianToolkit.Engine.Tests
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
| `Add-TKNote` | Record a timestamped technician note (category: Info/Action/Warning/Issue/Resolution) |
| `Get-TKNote` / `Clear-TKNote` | Read / reset the session note buffer |
| `Export-TKNoteReport` | Write the session's notes to a ticket-ready HTML report (with a plain-text paste block) |
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
The suite covers: `EscHtml`, `Format-Bytes`, `Get-TKConfig`/`Set-TKConfig`, `Test-IsAdmin`,
`Write-TKError`, the technician-note helpers, HTML report helpers, and module exports; plus
repo-wide gates — PowerShell syntax validation and UTF-8 BOM on every script, module-bootstrap
compliance, param block compliance (`-Unattended`, and `-WhatIf` on the destructive set),
GRIMOIRE registry integrity, license-header compliance (GPL notice and SPDX tag present and
correctly positioned in every source file), LICENSE integrity, retired tool names and filename
prefixes, removed deprecation stubs, no locally redefined shared helpers, and the PALADIN /
BEACON / PORTAL / CONJURE tier-mapper data tables (extracted by AST lookup rather than
dot-sourcing, since the tools launch their main flow on import).

### Verifying on Linux / in an agent sandbox

CI runs on `windows-latest`, but most of the suite is platform-agnostic and the linter runs
anywhere. In a sandbox where PowerShell is not installed, note that **PSGallery is often
blocked by network policy** — `Install-Module` then fails with *"No repository with the name
'PSGallery' was found"*, and registering it by hand does not help. Fetch from GitHub releases
instead:

```bash
# PowerShell 7 (tarball) and PSScriptAnalyzer (nupkg is a zip; extract onto PSModulePath)
curl -sSL -o pwsh.tar.gz https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/powershell-7.4.6-linux-x64.tar.gz
curl -sSL -o psa.zip    https://github.com/PowerShell/PSScriptAnalyzer/releases/download/1.22.0/PSScriptAnalyzer.1.22.0.nupkg
```

Pester cannot be obtained this way. Its GitHub releases carry **source**, and building it needs
the .NET SDK for its compiled assembly. Note the distinction: PSGallery ships Pester **prebuilt,
assembly included**, so `Install-Module Pester` is the route that works — it is only unavailable
when PSGallery itself is blocked. If you can get the gallery allowed through the sandbox's egress
policy, most of the suite then runs on Linux pwsh; `Describe 'Test-IsAdmin'` still fails there,
since `[Security.Principal.WindowsPrincipal]` does not exist off Windows. Otherwise the suite
stays CI-only.

What *is* reachable offline, and worth running before pushing:

- `Invoke-ScriptAnalyzer -Path . -Recurse -Settings .github/PSScriptAnalyzerSettings.psd1 -ExcludeRule PSAvoidUsingWriteHost`
  (CI fails on `Error` severity only; warnings are advisory).
- `[System.Management.Automation.Language.Parser]::ParseFile()` over every `.ps1`/`.psm1` — the
  same check the syntax tests make.
- The repo-wide gates above are all plain string/AST assertions and are cheap to replicate
  directly against the working tree.
- `Import-Module ./TechnicianToolkit.psm1` works on Linux pwsh, so the pure helpers
  (`EscHtml`, `Format-Bytes`, the HTML builders) can be exercised without Pester.

Two analyzer rules produce **false positives** throughout this repo — check before "fixing" a
hit: `PSReviewUnusedParameter` misses parameters used only inside nested function scopes (this
is why `-WhatIf` and `-Unattended` appear unused), and `PSUseUsingScopeModifierInNewRunspaces`
flags `Invoke-Command` script blocks that correctly declare their own `param()` and receive
values through `-ArgumentList`.

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
  RESTORATION, RUNEPRESS, and CONJURE. GRIMOIRE auto-detects and passes `-WhatIf` to any tool that
  declares it, and the Pester suite (`'-WhatIf declared on destructive tools'`) enforces the list.
- Tools that write logs expose `[switch]$Transcript`.

### License Notice Block

The toolkit is **GPL-3.0-or-later**. Every `.ps1` and `.psm1` opens with the GPL notice
header, above the comment-based help block:

```powershell
# <filename> - <A.C.R.O.N.Y.M.> — <one-line description>
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors
#
# This program is free software: you can redistribute it and/or modify
# ... (standard GPLv3 notice, copied verbatim from any existing script)
#
# SPDX-License-Identifier: GPL-3.0-or-later
```

Position matters: comment-based help is only picked up when preceded solely by comments and
blank lines, so the notice goes *above* the `<# .SYNOPSIS #>` block, never inside it. The
notice is per-file because a single tool script is a valid unit of distribution here — a
technician copying one `.ps1` onto a machine should still receive the license with it. The
Pester suite (`'License header compliance — all source files'`) enforces presence and position.

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
  "Covenant": { "DefaultTimezone": "", "DefaultLocalAdminUser": "" },
  "Conjure": {
    "DirectDownloads": [
      { "Name": "Acme RMM Agent", "Url": "https://...", "Args": "/S", "Sha256": "" }
    ]
  }
}
```

`Conjure.DirectDownloads` is the one array in the shape. `Get-TKConfig`'s nested-key
fill matches the default's type, so an absent array key comes back as `@()` rather
than `''` — a caller iterating it would otherwise get a single empty string.

`Get-TKConfig` returns these defaults if `config.json` is absent; `Set-TKConfig` creates or
updates the file.

### Adding a New Tool

1. Copy the GPL notice block and the header block from an existing tool; update the filename,
   acronym, synopsis, and version. The notice must stay above the `<# .SYNOPSIS #>` block.
2. Add the shared-module bootstrap block (see the initialization pattern above) and the
   appropriate admin check (`Invoke-AdminElevation` or `Assert-AdminPrivilege`). Copy the
   block verbatim from an existing tool — the Pester suite enforces the exact shape.
3. Register the tool in `grimoire.ps1`'s `$Tools` array with a unique numeric `Key`.
4. Add the script's filename to the Quick Launch and Usage sections in `README.md`.
5. The syntax-validation, module-bootstrap, and license-header compliance Pester tests will
   cover it automatically.
6. Nothing needs doing for the desktop application. The `.csproj` glob embeds every root
   `.ps1`, and the app reads `grimoire.ps1` at runtime — so a correctly registered tool
   appears in the window with a generated form and no C# change. That is the point of
   reading the registry rather than duplicating it.

### What the application reads out of a tool

The app never hardcodes anything about a tool. Three readers in
`app/TechnicianToolkit.Engine/` parse the scripts, which is what keeps the two halves from
drifting — and which means these conventions are load-bearing, not cosmetic:

| Reader | Reads | Breaks if |
|---|---|---|
| `ToolCatalog` | `$Tools` and `$CategoryOrder` in `grimoire.ps1` | The registry stops being an array of hashtable literals with a `File` key |
| `ToolParameters` | The **top-level** `param()` block | A tool takes input some other way; a nested function's params are correctly ignored |
| `ToolTraits` | `-WhatIf` / `-Unattended`, and the admin-gate call | A tool invents its own name for either switch |

`ToolParameters` turns type and validation attributes into form controls, so the attributes
are worth writing precisely: `[switch]` → checkbox, `[securestring]` → masked field,
`[ValidateSet]` → dropdown, `[ValidateScript]` → path picker, everything else → text box.
A `[ValidatePattern]` is carried through for the form to enforce.

The xUnit suite covers all three plus the extractor. Run it after touching anything in
`app/`, and after any change to the registry's shape.

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

### PALADIN vs SIGIL

Both touch Microsoft Defender, but they sit on opposite sides of the audit/enforce divide.

| Question | Reach for |
|----------|-----------|
| "Show me the current AV state, signatures, threats, exclusions, and ASR posture — I just want to read it." | **PALADIN** (read-only audit; never writes) |
| "Bring this machine into line with our security baseline — set the registry, enable the firewall rules, configure audit policy." | **SIGIL** (state-changing enforcement; supports `-WhatIf`) |

PALADIN is the diagnostic tool you run to decide whether enforcement is needed; SIGIL is the tool that does the enforcing. They compose: SIGIL hardens the machine, PALADIN later confirms the AV side held.

### CITADEL vs HERALD

Both are Active Directory tools; they sit on opposite sides of the act/report divide.

| Question | Reach for |
|----------|-----------|
| "Unlock this account / reset this password / add them to a group." | **CITADEL** (interactive AD management — it changes the directory) |
| "Give the customer a list of every account and what each one can do." | **HERALD** (read-only roster: full name / alias / access level, plus a review CSV) |
| "What are the password and lockout rules on this domain?" | **HERALD** (reads and scores the default domain policy and any fine-grained policies; SIGIL *sets* local policy but never reports domain policy) |

CITADEL's reports (stale accounts, password expiry) answer *account hygiene* questions about the
directory. HERALD answers an *access review* question — who holds privilege, and through which
groups. HERALD never writes to AD.

Where they overlap on inactivity: CITADEL's stale report is the standalone "who hasn't logged in"
export; HERALD folds the same signal in as one review flag among several, in the context of the
account's access level (an inactive Domain Admin ranks differently from an inactive standard user).

### HERALD vs WARD

Both produce an account roster with a Role column, at different scopes.

| Question | Reach for |
|----------|-----------|
| "Who can administer *this machine*?" | **WARD** (local SAM accounts + local Administrators group, one machine) |
| "Who can administer *the domain*?" | **HERALD** (AD user objects + privileged domain groups, whole directory) |

WARD runs on any Windows machine, domain-joined or not. HERALD requires a domain and RSAT.
Neither subsumes the other: a local administrator on a workstation does not appear in HERALD,
and a Domain Admin does not appear in WARD unless they also hold a local account.

### BEACON vs LANTERN

Both tools live in Network & Remote, but cover different layers of the network stack.

| Question | Reach for |
|----------|-----------|
| "What Wi-Fi networks does this machine remember, and which of them auto-connect?" | **BEACON** (saved WLAN profile inventory: SSID, auth, cipher, autoSwitch, key material) |
| "What hosts are alive on the LAN this machine is currently sitting on?" | **LANTERN** (subnet ping sweep + DNS / MAC / port scan of discovered hosts) |

BEACON looks inward at the wireless config baked into the machine; LANTERN looks outward at the LAN segment the machine is attached to. BEACON runs the same regardless of where the machine is plugged in; LANTERN's output is wholly dependent on the network it sits on at audit time.

### PORTAL vs LEYLINE

Both touch network connectivity but answer fundamentally different questions.

| Question | Reach for |
|----------|-----------|
| "What tunnels can leave this machine, and are any of them configured to leak credentials?" | **PORTAL** (built-in VPNs, Always-On triggers, NRPT, third-party VPN clients — inventory + auth/encryption tier verdict) |
| "Why can't this machine reach `<host>` right now?" | **LEYLINE** (live diagnostics: adapter state, ping, DNS, port test, IP renew, stack reset) |

PORTAL is a static configuration audit (read-only, identifies risky settings before they bite). LEYLINE is a live troubleshooting tool (can trigger remediation actions like `ipconfig /renew` and `netsh winsock reset`). Run PORTAL during onboarding and quarterly review; run LEYLINE when something is broken right now.
