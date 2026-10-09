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
    if ($env:TK_DISABLE_DOWNLOAD -in @('1', 'true')) {
        Write-Host "  [!!] Shared module TechnicianToolkit.psm1 not found, and TK_DISABLE_DOWNLOAD forbids fetching it." -ForegroundColor Red
        Write-Host "       Deploy the module next to this script from an approved release." -ForegroundColor Yellow
        exit 1
    }
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

The `TK_DISABLE_DOWNLOAD` check is the opt-out for MSPs running the kit under SOC 2:
with the machine environment variable set, nothing fetches toolkit code at run time,
so only the release they deployed can run (`docs/soc2.md`). Every path that downloads
toolkit code (this bootstrap, the forwarding stubs, GRIMOIRE's launcher, RITUAL's step
resolver) must check it first; `'TK_DISABLE_DOWNLOAD — every self-fetch path honours it'`
fails on any download of a GitHub raw URL that is not gated.

`Invoke-AdminElevation` re-launches the script as Administrator if not already elevated.
Scripts that use `Assert-AdminPrivilege` instead will error-exit if not elevated rather
than auto-relaunching — this is appropriate for scripts called programmatically (REVENANT,
HEARTH, EMBALM).

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
prefixes, removed deprecation stubs, the 6.0 renamed-tool forwarding stubs, no locally redefined shared helpers, and the GRIFFIN /
WISP / PORTAL / CONJURE tier-mapper data tables (extracted by AST lookup rather than
dot-sourcing, since the tools launch their main flow on import). The same AST extraction covers
the pure helpers of NECROPSY (bugcheck and dump-header parsing), RAVEN (inbox-rule, SPF and
DMARC scoring), WARD (LAPS policy precedence), GRIFFIN (platform-protection verdict), HALO
(Conditional Access predicates and emergency-access exclusion), LODESTAR (`nltest` / `w32tm`
parsers, Netlogon status table) and MINOTAUR (rights-mask and SID classification); the
root `BeforeAll` provides `Get-ToolAst` / `Get-ToolAssignmentValue` / `Get-ToolFunctionText` for
that. CARILLON's queue-issue, name and routing-target helpers, TORPOR's process / disk
usage arithmetic and power-plan parsing, SOLDER's DISM / CBS.log / SFC parsers and
servicing error-code table, CHALICE's product-family, support-date, channel,
`dsregcmd` and `cmdkey` parsers, and GARM's security-event XML, failure-code, source-ranking,
run-as and `quser` helpers, are covered the same way.
A finding catalog (CONDUIT, SOLDER, NECROPSY, TORPOR, RAVEN, HALO, CARILLON, CHALICE, LODESTAR, MINOTAUR, GARM) is tested against the codes its script
actually raises, so a new `Add-*Finding -Code` without a catalog entry fails CI.

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
  executing them. The current set is REVENANT, EMBALM, COVENANT, BASILISK, CLEANSE, WYRM, FORGE,
  WHETSTONE, RUNEPRESS, CONJURE, CONDUIT, SOLDER, LODESTAR, and CHALICE. GRIMOIRE auto-detects and passes `-WhatIf` to any tool that
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
`.NOTES` section holds only the `Version : X.Y` line.
Earlier versions embedded a cross-reference `Tools Available` list and a `Color Schema`
legend in every header; those were removed in v3.0 because they drifted out of sync on
every rename. The canonical tool list lives in `grimoire.ps1`'s `$Tools` registry.

**The suite has one version.** Every tool's `Version` line and its `Version` field in
`grimoire.ps1`'s `$Tools` registry carry the same suite-wide number. Do not
bump a single tool when you change it; record the change in `CHANGELOG.md` under
`[Unreleased]` instead. The version moves only at a release, when every header and every
registry entry move together. The Pester suite (`'Version consistency — script header matches
the GRIMOIRE registry'`) fails if any header disagrees with its registry row, or if the registry
holds more than one version.

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

### Naming

Each category has a theme, and since 6.0 its tools take their names from it. The category's
registry name carries the theme and the plain label together, and GRIMOIRE's menu letter follows
the plain label:

| Menu | Category | Theme | Keys |
|------|----------|-------|------|
| D | The Workshop — Deployment & Onboarding | the artificer: crafting & enchanting (FORGE, SOLDER, WHETSTONE) | 1–19 |
| R | The Observatory — Diagnostics & Reporting | scrying & magical sight (AUSPEX, AUGUR, SCRYER) | 20–39 |
| S | The Bestiary — Security | guardian beasts (WYRM, GRIFFIN, SPHINX) | 40–59 |
| N | The Crossroads — Network & Remote | arcane paths & lights (LEYLINE, LANTERN, PORTAL) | 60–69 |
| C | The Firmament — Cloud & Identity | the sky & the celestial (ZENITH, ORBIT, HALO) | 70–89 |
| M | The Necropolis — Data & Migration | the undead & ghosts (REVENANT, EXHUME, EMBALM) | 90–99 |

Rules for a new name, in order:

1. **Say what the tool does if it can.** A name that hints at the job beats perfect theme fit —
   NECROPSY and TORPOR keep off-theme names because they describe crash analysis and slowness
   exactly.
2. **Fit the category's theme** otherwise.
3. **Collide with nothing a technician or a SOC would misread** — no built-in Windows command
   (CIPHER clashed with `cipher.exe`) and no well-known attacker tool or malware family (BEACON,
   CITADEL, SHADE; Hydra and Cerberus were rejected for the same reason).
4. **Never reuse a retired name.** The v3.0 names (ORACLE, SENTINEL, BASTION, VAULT, PHANTOM,
   SPECTER, AEGIS, RELIC) and the 25 renamed in 6.0 are listed in the
   `'Legacy tool names must not reappear'` test, in dotted and report-prefix form.

The tools renamed in 6.0 each left a forwarding stub at the old filename (`$RenamedToolStubs` in
the test file lists them; `'Renamed-tool forwarding stubs'` checks them), and CODEX folds reports
saved under an old prefix into the new name (`$RenamedToolPrefixes` in `codex.ps1`). A future
rename follows the same pattern. Two names were deliberately left alone: the `config.json`
section `Archive` (EMBALM's settings — an ordinary noun, and renaming it would break existing
configs and the app's settings screen) and REVENANT's `-ArchiveZip` parameter.

### Adding a New Tool

1. Name it from its category's theme (see **Naming**), then copy the GPL notice block and the
   header block from an existing tool; update the filename, acronym, and synopsis. Keep the version at the current suite version. The notice must stay
   above the `<# .SYNOPSIS #>` block.
2. Add the shared-module bootstrap block (see the initialization pattern above) and the
   appropriate admin check (`Invoke-AdminElevation` or `Assert-AdminPrivilege`). Copy the
   block verbatim from an existing tool — the Pester suite enforces the exact shape.
3. Register the tool in `grimoire.ps1`'s `$Tools` array with the category's full name, the next
   free `Key` in the category's block, and the suite `Version`.
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

### CONDUIT vs WHETSTONE

Both are Windows Update tools; they sit on opposite sides of the connect/deploy divide.

| Question | Reach for |
|----------|-----------|
| "Windows Update says it couldn't connect to the update service." | **CONDUIT** (repairs the client's plumbing: WSUS pointer, WinHTTP proxy, blocking policies, time service, update services) |
| "The client works — go install the updates." | **WHETSTONE** (deploys updates via PSWindowsUpdate, handles power settings and reboots) |

CONDUIT is WHETSTONE's precondition. WHETSTONE assumes the update client can reach *a*
service and fails opaquely when it cannot; CONDUIT answers why. They compose in that order —
CONDUIT until its verdict is Healthy, then WHETSTONE.

The split matters for scope: CONDUIT never installs an update and never touches
`PSWindowsUpdate`, and WHETSTONE never edits the WindowsUpdate policy key. Neither tool
should grow into the other's half.

CONDUIT also declines to auto-fix two policy values it reports —
`DoNotConnectToWindowsUpdateInternetLocations` and `DisableWindowsUpdateAccess`. Both are
deliberate administrative decisions often pushed by domain GPO, so silently clearing them
would fight Group Policy and mask the real configuration. They are reported with the remedy
and left to the technician.

### SOLDER — the third Windows Update tool

| Question | Reach for |
|----------|-----------|
| "Updates fail with 0x800f081f / 0x80073712, a feature won't install, or system files are damaged." | **SOLDER** (component store and system files: DISM CheckHealth / RestoreHealth, SFC, CBS.log, failure codes decoded) |

The three compose as CONDUIT (can the client connect?) → SOLDER (is the store it installs into
sound?) → WHETSTONE (install). SOLDER classes each recent update failure code as *Store* (its
own), *Client* (CONDUIT's) or something else, so a technician knows which tool to reach for.

SOLDER never edits the WindowsUpdate policy key, never renames SoftwareDistribution / catroot2
(CONDUIT's `ResetCache`), and never runs `DISM /ResetBase` — that one is irreversible, removing
the ability to uninstall updates, and is left to a deliberate manual decision.

### NECROPSY vs VIGIL

Both read the event log, but for different questions.

| Question | Reach for |
|----------|-----------|
| "Why does this machine keep crashing, blue-screening or rebooting?" | **NECROPSY** (bugchecks, Kernel-Power 41, WHEA, TDR, dump files — one incident per unplanned stop) |
| "Are the services and scheduled tasks healthy, and what errors is the log throwing?" | **VIGIL** (service / task state and recent event-log errors in general) |

NECROPSY is read-only and stops at the bugcheck code and parameters. Walking the stack to name
the faulting driver is a debugger's job (WinDbg `!analyze -v`), and it should not grow a parser
for that.

### CHALICE vs ALMANAC vs RAVEN

| Question | Reach for |
|----------|-----------|
| "Office says Unlicensed Product / keeps asking for my password / Teams won't sign in — on this machine." | **CHALICE** (the Microsoft 365 Apps client: install, activation tokens, cached accounts, WAM, PRT, Teams cache) |
| "Is this user actually licensed?" | **ALMANAC** (tenant side) |
| "Is the mailbox itself compromised or misconfigured?" | **RAVEN** |

CHALICE is the one tool that deliberately runs **without elevating**: Office's identities, license
tokens and caches live in the signed-in user's profile, and an elevated run as another account
reads and resets the wrong one. It flags a run whose account differs from the console user. Its
resets touch only that user's Office state — never the WAM accounts in Settings, never the tenant.
Because the desktop app reads the admin gate by substring, chalice.ps1 must not mention either
gate function, even in a comment.

### TORPOR vs AUSPEX vs TALON

| Question | Reach for |
|----------|-----------|
| "Why is this PC slow right now?" | **TORPOR** (samples CPU / memory / disk load and names the processes, limits and boot delays responsible) |
| "Give me a general health snapshot of this machine." | **AUSPEX** |
| "Is anything persisting on this machine that shouldn't be?" | **TALON** |

TORPOR lists startup programs for their *cost*, not their safety, and is read-only — it names
what to stop or upgrade but never ends a process or disables a startup entry.

### RAVEN vs ALMANAC vs ORRERY

All three touch Exchange Online or Microsoft 365, at different layers.

| Question | Reach for |
|----------|-----------|
| "Is anyone's mailbox forwarding out, hiding mail, or otherwise showing signs of compromise? Is SPF / DKIM / DMARC right?" | **RAVEN** (Exchange Online, read-only) |
| "Who is licensed, who is unlicensed, who has MFA registered?" | **ALMANAC** (Microsoft Graph) |
| "What breaks if I delete this group?" — including its Exchange transport rules and delegations | **ORRERY** |

### LODESTAR vs LEYLINE vs SPHINX

| Question | Reach for |
|----------|-----------|
| "This machine can't reach `<host>`." | **LEYLINE** (general network diagnostics and stack resets) |
| "The trust relationship with the domain failed" / domain logons fail on one machine | **LODESTAR** (DC discovery, domain DNS, Kerberos time, secure channel; repairs the trust from the client side) |
| "Reset this user's password / unlock this account." | **SPHINX** (changes the directory) |

LODESTAR repairs from the member machine and never edits Active Directory, never changes DNS
settings, and never unjoins or rejoins — a deleted computer account is reported with the
rejoin steps. The machine-password reset is interactive only, because it needs domain
credentials typed by the technician.

### GARM vs SPHINX vs ARGUS

| Question | Reach for |
|----------|-----------|
| "Why does this account keep locking out, and from where?" | **GARM** (per-DC lockout state, 4740 / 4771 / 4776 from the DCs, sources ranked, the source machine checked for services, tasks and sessions holding the old password) |
| "Unlock it / reset the password." | **SPHINX** (its lockout option is the quick 4740 lookup; GARM is the full trace) |
| "Who holds privilege in the domain, and what is the lockout policy?" | **ARGUS** |

GARM is read-only: it never unlocks, resets a password, stops a service or ends a session, and
its source scan stops at reading services, scheduled tasks and `quser`. Fix the source, then
unlock with SPHINX — unlocking first just locks the account again. On event 4740 the caller
computer lives in `TargetDomainName`, not a field named for it; anything reading 4740 by
position must use index 1 (index 4 is the DC's own machine account).

### MINOTAUR vs ARGUS vs WARD

All three are access reviews; they differ in what is being accessed.

| Question | Reach for |
|----------|-----------|
| "Who can get into this file share, and can everyone write to it?" | **MINOTAUR** (SMB share + NTFS permissions, HTML + CSV) |
| "Who holds privilege in the domain?" | **ARGUS** |
| "Who can administer this machine?" | **WARD** |

### CARILLON vs ASTERISM

| Question | Reach for |
|----------|-----------|
| "Who answers the Sales line? Why does the main number ring nobody after hours?" | **CARILLON** (Teams Phone call queues, agents by name, auto attendant menus, resource accounts) |
| "Which teams are orphaned, public, or full of guests?" | **ASTERISM** (the Teams / M365 group estate) |

CARILLON is read-only. It reports agent membership and routing but never adds or removes agents
or edits a queue — changing who takes calls is a Teams admin center task.

### HALO vs ECLIPSE

| Question | Reach for |
|----------|-----------|
| "Which identities are risky — guests, stale admins, password-never-expires?" | **ECLIPSE** |
| "Do the Conditional Access policies actually enforce MFA and block legacy auth, and is there a break-glass path?" | **HALO** |

### FATHOM vs AUGUR

Both tools deal with disk health but cover different layers:

| Tool | Focus |
|------|-------|
| **F.A.T.H.O.M.** | Volume space monitoring — used/free space, low-space alerts, temp cleanup, old profile detection |
| **A.U.G.U.R.** | Physical hardware health — SMART status, wear prediction, failure forecasting, bus/media type |

Run FATHOM for "is this drive running out of space?"; run AUGUR for "is this drive about to die?".

### SCRYER vs the single-domain diagnostic tools

S.C.R.Y.E.R. (`scryer.ps1`) is a one-shot consolidated report that rolls five diagnostic passes (system overview, local users, disk space, SMART health, services & tasks) into a single HTML file. It exists for ticket attachments and machine handoffs where one snapshot is more useful than five separate reports.

| Question | Reach for |
|----------|-----------|
| "Give me one file summarising this machine." | **SCRYER** |
| Deep dive on any one of: system health, users, free space, disk reliability, services | AUSPEX / WARD / FATHOM / AUGUR / VIGIL respectively |

SCRYER's per-section depth is intentionally shallower than the dedicated tools — it samples each domain rather than reproducing the full report.

### RITUAL vs CODEX

Both tools produce a rollup HTML that links out to other tool reports — they answer different questions.

| Question | Reach for |
|----------|-----------|
| "Run an ordered sequence of tools and give me one rollup of the run." | **RITUAL** (executes a recipe, captures status / duration / artifacts per step) |
| "I've already run a bunch of tools ad-hoc — give me one index of what's on disk." | **CODEX** (filesystem scan only, no execution; relative links so the rollup stays clickable when zipped) |

RITUAL produces a record *of an execution* — step status, durations, errors. CODEX produces a record *of a directory* — what reports exist, when, and how big they are. Use RITUAL when you control the run; use CODEX when the reports already exist.

### GRIFFIN vs BASILISK

Both touch Microsoft Defender, but they sit on opposite sides of the audit/enforce divide.

| Question | Reach for |
|----------|-----------|
| "Show me the current AV state, signatures, threats, exclusions, and ASR posture — I just want to read it." | **GRIFFIN** (read-only audit; never writes) |
| "Bring this machine into line with our security baseline — set the registry, enable the firewall rules, configure audit policy." | **BASILISK** (state-changing enforcement; supports `-WhatIf`) |

GRIFFIN is the diagnostic tool you run to decide whether enforcement is needed; BASILISK is the tool that does the enforcing. They compose: BASILISK hardens the machine, GRIFFIN later confirms the AV side held.

### SPHINX vs ARGUS

Both are Active Directory tools; they sit on opposite sides of the act/report divide.

| Question | Reach for |
|----------|-----------|
| "Unlock this account / reset this password / add them to a group." | **SPHINX** (interactive AD management — it changes the directory) |
| "Give the customer a list of every account and what each one can do." | **ARGUS** (read-only roster: full name / alias / access level, plus a review CSV) |
| "What are the password and lockout rules on this domain?" | **ARGUS** (reads and scores the default domain policy and any fine-grained policies; BASILISK *sets* local policy but never reports domain policy) |

SPHINX's reports (stale accounts, password expiry) answer *account hygiene* questions about the
directory. ARGUS answers an *access review* question — who holds privilege, and through which
groups. ARGUS never writes to AD.

Where they overlap on inactivity: SPHINX's stale report is the standalone "who hasn't logged in"
export; ARGUS folds the same signal in as one review flag among several, in the context of the
account's access level (an inactive Domain Admin ranks differently from an inactive standard user).

### ARGUS vs WARD

Both produce an account roster with a Role column, at different scopes.

| Question | Reach for |
|----------|-----------|
| "Who can administer *this machine*?" | **WARD** (local SAM accounts + local Administrators group, one machine) |
| "Who can administer *the domain*?" | **ARGUS** (AD user objects + privileged domain groups, whole directory) |

WARD runs on any Windows machine, domain-joined or not. ARGUS requires a domain and RSAT.
Neither subsumes the other: a local administrator on a workstation does not appear in ARGUS,
and a Domain Admin does not appear in WARD unless they also hold a local account.

### WISP vs LANTERN

Both tools live in The Crossroads (Network & Remote), but cover different layers of the network stack.

| Question | Reach for |
|----------|-----------|
| "What Wi-Fi networks does this machine remember, and which of them auto-connect?" | **WISP** (saved WLAN profile inventory: SSID, auth, cipher, autoSwitch, key material) |
| "What hosts are alive on the LAN this machine is currently sitting on?" | **LANTERN** (subnet ping sweep + DNS / MAC / port scan of discovered hosts) |

WISP looks inward at the wireless config baked into the machine; LANTERN looks outward at the LAN segment the machine is attached to. WISP runs the same regardless of where the machine is plugged in; LANTERN's output is wholly dependent on the network it sits on at audit time.

### PORTAL vs LEYLINE

Both touch network connectivity but answer fundamentally different questions.

| Question | Reach for |
|----------|-----------|
| "What tunnels can leave this machine, and are any of them configured to leak credentials?" | **PORTAL** (built-in VPNs, Always-On triggers, NRPT, third-party VPN clients — inventory + auth/encryption tier verdict) |
| "Why can't this machine reach `<host>` right now?" | **LEYLINE** (live diagnostics: adapter state, ping, DNS, port test, IP renew, stack reset) |

PORTAL is a static configuration audit (read-only, identifies risky settings before they bite). LEYLINE is a live troubleshooting tool (can trigger remediation actions like `ipconfig /renew` and `netsh winsock reset`). Run PORTAL during onboarding and quarterly review; run LEYLINE when something is broken right now.
