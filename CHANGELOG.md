# Changelog

All notable changes to TechnicianToolkit are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [Unreleased]

### Added
- **`Format-Bytes` helper** in `TechnicianToolkit.psm1`. Unifies the byte-to-human-readable formatter that was previously duplicated in `cleanse.ps1` and `threshold.ps1`. Supports B/KB/MB/GB/TB with two-decimal precision.
- **Pester coverage for HTML report helpers.** New `Describe 'HTML report helpers'` block exercises `Get-TKHtmlHead` / `Get-TKHtmlFoot` — document structure, CSS embedding, meta-bar / nav-bar rendering, HTML-escape of special characters in the title, and round-trip tag balance.
- **Pester coverage for `Format-Bytes`** — B, KB, MB, GB, TB, and zero-byte cases.
- **`-WhatIf` compliance test** — asserts that each destructive tool (`revenant`, `archive`, `covenant`, `sigil`, `cleanse`, `cipher`) declares the `-WhatIf` switch so GRIMOIRE's dry-run passthrough has something to bind.
- **Deprecation stub forwarding test** — asserts each v3.0 legacy-name stub (`oracle`, `sentinel`, `bastion`, `vault`, `phantom`, `specter`, `aegis`, `relic`) forwards to the correct renamed target, emits a `Write-Warning`, and captures remaining arguments via `ValueFromRemainingArguments`.
- **No-duplicated-helpers test** — guards against re-introducing local `HtmlEncode` / `Format-Bytes` definitions in any tool script.
- **`Write-TKError` wired into `covenant.ps1`** at the three domain-join failure paths (unattended AD, interactive AD, Entra ID). Failures now emit a Teams webhook notification when one is configured, not just a console line.
- **`Write-TKError` wired into four more data-safety tools** so silent failures in high-impact paths now emit Teams telemetry instead of only a console line:
  - `archive.ps1` — ZIP creation failure (pre-reimage backup lost).
  - `revenant.ps1` — Copy-ProfileFolder and Copy-ProfileFile failures (migration data loss).
  - `cipher.ps1` — Enable-BitLocker, Disable-BitLocker, recovery-key backup, and Resume-BitLocker failures.
  - `sigil.ps1` — central `Set-BaselineReg` helper, so every security-baseline write failure now telemeters with the category it belongs to.
- **SCRYER documentation added across README and CLAUDE.md.** The 579-line unified diagnostic tool (`scryer.ps1`, GRIMOIRE key 16) was registered in the hub but previously absent from the overview table, description section, Quick Launch, Usage, Configuration, and Logging tables. Each of those sections now lists SCRYER. CLAUDE.md's Tool Distinctions section gained a SCRYER-vs-single-domain-tools entry.
- **`-Transcript` switch added to `augur.ps1` and `cleanse.ps1`**, matching the toolkit-wide convention. Both scripts now start a TKTranscript when the flag is passed and stop it on exit.
- **`-Transcript` switch added to `restoration.ps1`** and its pre-existing always-on transcript routed through the configured log directory when the flag is set, keeping the historical `%TEMP%` default when it is not.
- **Pester `Legacy tool names must not reappear` gained a second pattern list** covering bare-underscore filename prefixes (`ORACLE_`, `SENTINEL_`, `BASTION_`, `VAULT_`, `PHANTOM_`, `SPECTER_`, `AEGIS_`, `RELIC_`). The original test only caught dotted-acronym form, which is why the README logging table drifted through v2→v3 without failing CI.
- **`-WhatIf` added to three more state-changing tools** and included in the Pester `-WhatIf compliance` test list:
  - `forge.ps1` — lists pending Windows Update driver updates and previews the extension-specific handler for each local file (extract / pnputil / silent EXE / msiexec) without running any.
  - `restoration.ps1` — lists every pending update that would be installed and short-circuits the reboot decision so a dry run never reboots the machine.
  - `runepress.ps1` — previews each driver install at the call site it would hit (pnputil / EXE silent / msiexec) and skips the network-printer stage entirely, leaving no ports, printers, or extracted staging folders behind.
- **`OrgName` now appears in the subtitle of seven HTML reports that previously ignored it** — `threshold`, `augur`, `gargoyle`, `lantern`, `reliquary`, `scryer`, and `citadel` (both the stale and password-expiry reports). Each tool now calls `Get-TKConfig`, HTML-escapes `OrgName`, and prepends it to its existing subtitle string when configured. Unconfigured deployments see no change. `talisman.ps1` continues to display the Azure subscription name in place of an org name, which matches the Azure-focused scope of that tool.
- **`Write-TKError` wired into five more tools' highest-impact failure paths** (12 new call sites). Combined with earlier rounds the toolkit now surfaces telemetry on every incident-worthy failure rather than only the domain-join and data-safety paths:
  - `citadel.ps1` — `Unlock-ADAccount`, `Set-ADAccountPassword`, `Enable-ADAccount`, `Disable-ADAccount`, `Add-ADGroupMember`, `Remove-ADGroupMember`.
  - `conjure.ps1` — per-package install failure inside the winget/Chocolatey loop.
  - `forge.ps1` — Windows Update driver scan/install failure.
  - `runepress.ps1` — pnputil driver install (ZIP path), EXE launch failure, and msiexec launch failure.
  - `restoration.ps1` — `Install-WindowsUpdate` failure.
- **New tool: `G.O.L.E.M.`** (`golem.ps1`, Cloud & Identity key 42) — Governs & Observes Licensed Endpoint Management. Connects to Microsoft Graph and audits the Intune-managed device estate: device inventory with OS / ownership / compliance / join-type breakdown, compliance state summary, stale-device detection at 30 / 60 / 90-day buckets, and configuration-profile assignment coverage. Dark-themed HTML report with `OrgName` prefix support. Telemetry wired on Graph auth failure and Intune query failure. Complements `R.E.L.I.Q.U.A.R.Y.` (M365 licensing / MFA / shared mailboxes) by covering the device-management side of the same tenant.
- **New tool: `T.E.T.H.E.R.`** (`tether.ps1`, Data & Migration key 52) — Tests Endpoint Tethering: Hosted Environment Readiness. Pre-migration validator that answers "is this user's data actually going to be in the cloud when we hand them a new laptop?" Checks OneDrive client state (install, version, process), signed-in accounts from `HKCU:\Software\Microsoft\OneDrive\Accounts`, Known Folder Move redirection for Desktop / Documents / Pictures via `User Shell Folders`, per-folder file count and size, and OneDrive-related Application-log events from the last 7 days. Produces a dark-themed HTML report with a red / yellow / green readiness verdict summarising what will and will not follow the user to the new machine.
- **New tool: `E.X.H.U.M.E.`** (`exhume.ps1`, Data & Migration key 53) — Enumerates, eXposes & Hunts Unmigrated Mail Entries. Outlook data-file discovery that runs before a mail migration. Walks the Outlook profile hives (Office 14 / 15 / 16), recursively scans every local fixed drive for `*.pst` (and optionally `*.ost` via `-IncludeOst`), correlates disk files against configured profiles to flag orphans, and applies three heuristics: ≥ 50 GB (Exchange Online Import Service hard limit — red), 10–50 GB (slow-migration warning — yellow), unaccessed 365+ days (archive-on-ingest candidate — yellow). Produces a dark-themed HTML report with a red / yellow / green verdict and the full file inventory.
- **New tool: `W.R.A.I.T.H.`** (`wraith.ps1`, Cloud & Identity key 43) — Watches Registrations, Access, Identities, Tokens & Hygiene. Entra ID identity-hygiene audit that covers the security-and-cost questions RELIQUARY (licensing) and GOLEM (devices) do not: guest accounts with sign-in age and invite state, every member of every active directory role (one row per user-per-role, high-tier roles called out in console), members with `DisablePasswordExpiration` set, privileged users inactive 60+ days (deduplicated from the role audit and annotated with all their role assignments), and disabled-but-still-licensed accounts that leak license cost. Dark-themed HTML report with `OrgName` prefix and six summary cards. Telemetry via `Write-TKError` on Graph auth failure and each Graph query failure. Fills slot 43 — the originally-proposed slot 42 was taken by GOLEM when WRAITH was skipped the first time.
- **New tool: `R.I.T.U.A.L.`** (`ritual.ps1`, Deployment & Onboarding key 7) — Runs Integrated Tool Usage in Automation Loops. Workflow orchestrator that runs an ordered sequence of toolkit scripts as a single recipe and rolls the results up into one HTML report with per-step status, duration, and clickable links to each child report. Ships four built-in recipes (`Onboard`, `Retire`, `HealthCheck`, `TenantSweep`) and accepts custom PSD1 recipe files. Per-step log-directory snapshotting attributes each newly-produced file back to the step that produced it. Default behaviour aborts on first failure with `-ContinueOnError` as opt-out; individual steps can override with `StopOnError = $true` in the recipe. First meta-tool in the toolkit — the other 28 tools now compose.
- **New tool: `A.N.V.I.L.`** (`anvil.ps1`, Diagnostics & Reporting key 17) — Audits & Notates Vendor Inventory & Lifecycle. BIOS / UEFI / firmware audit. Collects system identity (manufacturer, model, SKU, serial, UUID), BIOS vendor / version / release date / age, firmware type (UEFI vs legacy BIOS with GPT cross-check), and Secure Boot state. Detects presence of vendor firmware-update tooling — Dell Command, HP Image Assistant / Support Assistant, Lenovo System Update / Vantage, Microsoft Surface UEFI Configurator — against the auto-detected manufacturer. Scans Windows Update via PSWindowsUpdate for pending driver / firmware updates. Produces a dark-themed HTML report with a red / yellow / green readiness verdict.
- **New tool: `T.A.L.O.N.`** (`talon.ps1`, Security key 24) — Tracks Anomalies & Locates Otherwise-silent Nastiness. Persistence / autoruns audit. Sweeps seven persistence surfaces: Run/RunOnce keys across HKCU, HKLM, and WOW6432Node; per-user and All-Users Startup folders (with `.lnk` target resolution); non-Microsoft auto-start services; non-Microsoft scheduled tasks (one row per action); WMI event subscriptions under `root\subscription`; IFEO Debugger hijacks; Winlogon Shell / Userinit / AppInit_DLLs. Every entry is enriched with target-exists check and Authenticode signature status (Microsoft-signed, third-party signed, unsigned, tampered). Does not judge malice — the tool's job is visibility. Dark-themed HTML report with six summary cards and one detail table per surface.
- **New tool: `T.O.T.E.M.`** (`totem.ps1`, Security key 25) — Trusted Observer of Transparent Execution Modules. TPM health audit. Reads TPM state via `Get-Tpm` (present / enabled / activated / ready / owned, manufacturer, spec version, auto-provisioning, restart-pending), parses the spec into a Win11-readiness flag, and cross-references `Get-BitLockerVolume` to enumerate which BitLocker volumes actually have TPM-based key protectors so a technician can answer "if this TPM goes bad, which drives stop unlocking?". Surfaces endorsement-key presence via `Get-TpmEndorsementKeyInfo` for Autopilot attestation readiness. Red / yellow / green verdict with specific remediation hints (enable in firmware, provision, clear-and-reprovision, vendor BIOS update for dTPM→fTPM). Dark-themed HTML report with six summary cards.
- **New tool: `P.Y.R.E.`** (`pyre.ps1`, Diagnostics & Reporting key 18) — Power-Yield Reliability Evaluator. Laptop battery health audit. Pulls design capacity, current full-charge capacity, cycle count, and live charge/discharge state from the `ROOT\WMI` battery classes (`BatteryStaticData`, `BatteryFullChargedCapacity`, `BatteryCycleCount`, `BatteryStatus`) keyed by `InstanceName`, joins to `Win32_Battery` for the user-friendly device name and chemistry. Health % = full / design. Applies industry thresholds (≥ 80% / 60-80% / < 60% capacity; < 300 / 300-500 / ≥ 500 cycles) to generate a HEALTHY / REPLACEMENT SOON / REPLACE NOW verdict per battery; worst value across all batteries drives the overall report class. Dark-themed HTML report with Wh values, colour-coded badges, and explicit NO BATTERY handling for desktops / servers / VMs.
- **New tool: `C.O.N.C.L.A.V.E.`** (`conclave.ps1`, Cloud & Identity key 44) — Consolidates Organisational Networks, Chats, Licenses, Access, Visibility & Entitlements. Microsoft Teams audit. Enumerates every team-backed M365 group (`resourceProvisioningOptions/Any(x:x eq 'Team')`) and enriches with owner count / enabled-owner count / member count / guest count / visibility / sensitivity labels / renewed date. Derives five findings: orphan teams (no owners or all owners disabled), public teams, teams with guest members (sorted by guest count), large teams (≥ 250 members, editable), and stale teams (`RenewedDateTime` older than 365 days, editable). Dark-themed HTML report with six summary cards and six detail tables. Completes the four-tool tenant-posture sweep (TALISMAN + RELIQUARY + GOLEM + WRAITH + CONCLAVE in one sign-in via R.I.T.U.A.L. TenantSweep recipe).
- **Parameter validation added to nine string inputs across eight tools** so obvious mis-inputs fail at bind-time instead of deeper in the script:
  - `cipher.ps1` `-Drive` — drive-letter pattern `^[A-Za-z]:?$`.
  - `gargoyle.ps1` `-Target` and `leyline.ps1` `-Target` — hostname / hostname:port pattern with empty-string default preserved.
  - `talisman.ps1` `-TenantId` and `-SubscriptionId` — GUID pattern.
  - `archive.ps1` `-Items`, `revenant.ps1` `-Items`, `sigil.ps1` `-Categories` — `"A"` or a comma-separated list of one/two-digit numbers.
  - `covenant.ps1` `-NewComputerName` — Windows NetBIOS hostname rules (1–15 chars, alphanumerics and hyphens, no leading/trailing hyphen).
  - `scryer.ps1` `-OutputPath` — `ValidateScript` that accepts empty string, an existing path, or a path whose parent exists (so the tool can create a new output directory).

### Changed
- **README logging table corrected for seven tools** whose emitted filename prefixes drifted at the v2→v3 rename and were never updated. `auspex.ps1`, `gargoyle.ps1`, `citadel.ps1`, `artifact.ps1`, `shade.ps1`, `reliquary.ps1`, and `revenant.ps1` now show the prefixes their source actually produces (`AUSPEX_`, `GARGOYLE_`, `CITADEL_Stale_`/`CITADEL_PwdExpiry_`, `ARTIFACT_`, `SHADE_`, `RELIQUARY_`, `REVENANT_MigrationLog_`).
- **README "no LiveConnect counterpart" row expanded** to include the six tools that were missing from it: WARD, AUSPEX, SCRYER, CONJURE, RESTORATION, SIGIL.
- **`README.md` SHADE description** — retrieved-output folder reference corrected from `SPECTER_<MachineName>\` to `SHADE_<MachineName>\`.
- **Empty `catch {}` blocks annotated in three tools** where the swallow-on-failure intent was not obvious: `auspex.ps1` (SecurityCenter2 WMI namespace absent on Server/Core), `sigil.ps1` (Disable-NetFirewallRule group missing on Home/Core), and `talisman.ps1` (Recovery Services vault contents requiring Backup Reader RBAC). Other empty catches in the toolkit cover dead-host sweeps, disposal chains, and other best-effort paths where the intent is already clear from context.
- **`HtmlEncode` helper removed from five tools** (`augur.ps1`, `artifact.ps1`, `citadel.ps1`, `gargoyle.ps1`, `ward.ps1`) — each reimplemented the same escape logic. Callers now use the module's `EscHtml`, which also handles null input.
- **Local `Format-Bytes` definitions removed from `cleanse.ps1` and `threshold.ps1`** in favour of the module export. Threshold's output format changes slightly — KB values now show two decimals instead of one for consistency with the other units.
- **`revenant.ps1` source-path parameters validated** — `-SourcePath` and `-ArchiveZip` now carry `[ValidateScript]` attributes that reject non-existent paths at parameter-bind time rather than failing deeper in the script.
- **`grimoire.ps1` self-delete guarded against git checkouts.** The cleanup that removes `$PSCommandPath` after a one-shot hub session now skips when a `.git` directory is present next to the script, so running GRIMOIRE from a cloned working tree no longer silently deletes the file.

---

## [3.0.0] - 2026-04-21

### Added
- **`Phantom` -> `Revenant` config migration shim** in `Get-TKConfig` (`TechnicianToolkit.psm1`). Existing `config.json` files that still carry a populated `Phantom.DefaultDestination` from v2.x are transparently copied into the new `Revenant.DefaultDestination` on first read, so upgrading deployments do not lose the configured migration target.
- **Legacy-name regression Pester test** (`tests/TechnicianToolkit.Tests.ps1`). Scans every `.ps1` and `.md` file in the repo (excluding the eight deprecation stubs, the test file itself, and CHANGELOG) and fails if any of the retired dotted acronyms (`O.R.A.C.L.E.`, `S.E.N.T.I.N.E.L.`, `B.A.S.T.I.O.N.`, `V.A.U.L.T.`, `P.H.A.N.T.O.M.`, `S.P.E.C.T.E.R.`, `A.E.G.I.S.`, `R.E.L.I.C.`) reappears in new code.
- **Deprecation stubs for eight retired filenames** — `oracle.ps1`, `sentinel.ps1`, `bastion.ps1`, `vault.ps1`, `phantom.ps1`, `specter.ps1`, `aegis.ps1`, `relic.ps1`. Each forwards every argument to the renamed script (downloading the new script from GitHub if missing), prints a one-line `Write-Warning` so pinned runbooks and old quick-launch snippets surface the rename, and will be removed in a future release.

### Changed
- **Toolkit-wide version bump to 3.0.** Every tool's `.NOTES Version`, banner v-tag, `Get-TKHtmlFoot -ScriptName` footer, and GRIMOIRE registry `Version` field now reads `3.0`. The GRIMOIRE hub banner now reads `Hub v3.0`.
- **Header comment blocks trimmed.** The `Tools Available` cross-reference list and `Color Schema` legend have been removed from every script's `<# ... #>` header block. They duplicated information already available in `grimoire.ps1`'s registry and README, and drifted out of sync every time a tool was added or renamed. Scripts that carried a tool-specific list (e.g. `shade.ps1`'s `Remote-Compatible Tools`) keep it.

---

## [2.0.0] - 2026-04-21

### Changed
- **Thematic rename of eight tools** to avoid name collisions with commercial products and strengthen the arcane theme. Filenames, banners, log/report prefixes, config section, and all cross-references updated toolkit-wide.
  - `oracle.ps1` -> `auspex.ps1` — O.R.A.C.L.E. -> A.U.S.P.E.X. (Audits, Uncovers, Surveys Performance, Events & eXceptions)
  - `sentinel.ps1` -> `gargoyle.ps1` — S.E.N.T.I.N.E.L. -> G.A.R.G.O.Y.L.E. (Guards Against Runtime Glitches On Your Log Events)
  - `bastion.ps1` -> `citadel.ps1` — B.A.S.T.I.O.N. -> C.I.T.A.D.E.L. (Centralizes Identity, Tasks, Accounts, Directories, Entitlements & Logons)
  - `vault.ps1` -> `reliquary.ps1` — V.A.U.L.T. -> R.E.L.I.Q.U.A.R.Y. (Reports, Evaluates Licenses, Inventories, Quotas, Users, Access & Registration Yields)
  - `phantom.ps1` -> `revenant.ps1` — P.H.A.N.T.O.M. -> R.E.V.E.N.A.N.T. (Relocates, Extracts, Validates Environments, Networks, Accounts 'N Transfers)
  - `specter.ps1` -> `shade.ps1` — S.P.E.C.T.E.R. -> S.H.A.D.E. (Summons Hosts for Administrative Deployment & Execution)
  - `aegis.ps1` -> `talisman.ps1` — A.E.G.I.S. -> T.A.L.I.S.M.A.N. (Tenant Assessment, Logging, Infrastructure, Security, Monitoring & Access Navigator)
  - `relic.ps1` -> `artifact.ps1` — R.E.L.I.C. -> A.R.T.I.F.A.C.T. (Audits, Reports Trust, Identity, Fingerprints, Authority, Certificates & TLS)
- **config.json section `Phantom` renamed to `Revenant`** — the `DefaultDestination` key remains; `Get-TKConfig` defaults and the HEARTH wizard field were updated to match. Existing configs with the old `Phantom` section will lose that value after re-save and must be re-entered via HEARTH.
- **GRIMOIRE registry** — eight `Name`/`File` entries updated to the new tool names and filenames.
- **Pester tests** — `has Phantom section` check updated to `has Revenant section`; param-compliance exclusion list swapped `specter.ps1` for `shade.ps1`.

### Added
- **DWARF and PURGE in GRIMOIRE** — both tools are now registered in the GRIMOIRE hub (keys 22 and 23) under the "Diagnostics & Reporting" category and are accessible from the central launcher.
- **DWARF and PURGE module integration** — both scripts now import `TechnicianToolkit.psm1` and use `Invoke-AdminElevation`, matching every other tool and enabling centralized error telemetry and config access.
- **Module import compliance test** — new Pester test block verifies that every `.ps1` tool script in the root imports `TechnicianToolkit.psm1`, catching regressions before CI.
- **`TeamsWebhook` in `config.json` template** — the on-disk template now includes the `TeamsWebhook` key so technicians can see and populate it without inspecting the module source.
- **SIGIL `-WhatIf` mode** — preview every registry, firewall, account policy, audit policy, and protocol change before applying. All functions updated: `Set-BaselineReg`, `Apply-Firewall`, `Apply-GuestAccount`, `Apply-PasswordPolicy`, `Apply-RemoteDesktop`, `Apply-AuditPolicy`, `Apply-LegacyProtocols`. Summary output reflects WhatIf status counts.
- **Centralized error telemetry** (`Write-TKError` in `TechnicianToolkit.psm1`) — appends structured JSON-lines to a monthly error log in the configured `LogDirectory`. Optionally posts to a Teams incoming webhook via the new `TeamsWebhook` config key.
- **GRIMOIRE download integrity** — downloaded scripts are passed through the PowerShell parser before execution. Corrupt or syntactically invalid files are removed and the launch is aborted rather than executing unknown code.
- **Script version display** — each tool in the GRIMOIRE registry carries a `Version` field; the interactive menu now shows the version alongside every tool name.
- **GitHub Actions CI** (`.github/workflows/ci.yml`) — runs PSScriptAnalyzer on every push and pull request to `main`. A second job runs the Pester test suite and publishes results as a workflow artifact.
- **PSScriptAnalyzer settings** (`.github/PSScriptAnalyzerSettings.psd1`) — suppresses intentional patterns (non-exported verb names, `Write-Host` in a console tool).
- **Pester test suite** (`tests/TechnicianToolkit.Tests.ps1`) — covers `EscHtml`, `Test-IsAdmin`, `Write-TKError`, module exports, and a syntax-validation sweep of all `.ps1` files.

- **`-WhatIf` for PHANTOM** — preview what profile data would be copied (file count, size, destination) without performing any transfers. Compatible with all item types (folder and file). WhatIf items appear in the migration summary with a distinct Cyan color and count.
- **`-WhatIf` for PURGE** — preview which cleanup categories would be cleaned without deleting any files. Category sizes are shown in the selection menu; WhatIf mode replaces the "Total Space Freed" summary line with a "DRY RUN" notice.
- **HEARTH TeamsWebhook field** — the setup wizard now prompts for and saves the Teams incoming webhook URL as step 3 of the wizard, making it configurable without hand-editing `config.json`.
- **CLAUDE.md** — developer guide covering architecture, module pattern, test commands, key conventions, and the DWARF/THRESHOLD distinction.
- **GRIMOIRE registry validation test** — new Pester test block parses every `File = '...'` entry from `grimoire.ps1` and verifies the corresponding `.ps1` file exists on disk.
- **Param block compliance test** — new Pester test block verifies that every tool script (excluding `grimoire.ps1`) declares a `-Unattended` parameter.
- **Version bumps** — `hearth.ps1`, `phantom.ps1`, `purge.ps1`, and `grimoire.ps1` bumped from 1.0 to 1.1.
- **HTML theme system** (`TechnicianToolkit.psm1`) — three new exported functions: `Get-TKHtmlCss`, `Get-TKHtmlHead`, `Get-TKHtmlFoot`. Provide a shared dark-theme `<style>` block, a full page header with title/subtitle/meta bar/nav anchors, and a footer. All HTML-generating scripts now call these instead of embedding bespoke `<style>` blocks and boilerplate HTML.
- **Unified dark HTML theme** — all 10 report-generating scripts (ORACLE, WARD, THRESHOLD, SENTINEL, AUGUR, RELIC, BASTION, LANTERN, VAULT, AEGIS) migrated to the `Get-TKHtmlHead`/`Get-TKHtmlFoot` pattern. Reports share a dark cyan-accented design via CSS custom properties (`--tk-bg`, `--tk-surface`, `--tk-cyan`, etc.) and standardized class names (`.tk-card`, `.tk-table`, `.tk-badge-ok/warn/err/info`, `.tk-section`, `.tk-summary-card`, etc.).
- **HTML Report Pattern in CLAUDE.md** — documents the `Get-TKHtmlHead`/`Get-TKHtmlFoot` usage pattern and full CSS class reference so developers adding new HTML-generating tools can use the shared theme correctly.
- **Pester module-export tests updated** — `Get-TKHtmlCss`, `Get-TKHtmlHead`, and `Get-TKHtmlFoot` added to the `$expectedFunctions` list in `TechnicianToolkit.Tests.ps1`.

### Changed
- **GRIMOIRE THRESHOLD description** — updated from "Disk & storage health — physical disk status, volume space, cleanup, old profiles" to "Disk space monitor — volume usage, low-space alerts, temp cleanup, old profile detection" to clearly distinguish it from DWARF's SMART/hardware focus.
- **GRIMOIRE HEARTH description** — updated to mention Teams webhook configuration.
- **README Diagnostics & Reporting table** — THRESHOLD description updated to match GRIMOIRE; DWARF and PURGE rows added (keys 22 and 23).
- **README THRESHOLD description** — clarified as "Disk space monitor" to distinguish from DWARF.
- **README HEARTH section** — wizard now described as covering seven fields including Teams webhook URL.
- **README PHANTOM section** — updated to document OneDrive KFM awareness, ARCHIVE ZIP restore, and `-WhatIf` support.
- **README Quick Launch** — DWARF and PURGE one-liners added.
- **README Usage** — DWARF and PURGE direct-run entries added with disambiguating comments.
- **README Configuration** — added config key reference table (OrgName, LogDirectory, TeamsWebhook, Archive, Phantom, Covenant); DWARF and PURGE rows added; HEARTH row updated.
- **README Logging** — DWARF and PURGE log output rows added.
- **README License** — replaced placeholder with MIT reference.
- **DWARF renamed to AUGUR** (`dwarf.ps1` -> `augur.ps1`) — full rebrand to A.U.G.U.R. (Automated Universal Gauge for Understanding Resources); banner, ASCII art, internal references, `.NOTES` tool list, and GRIMOIRE registry entry all updated.
- **PURGE renamed to CLEANSE** (`purge.ps1` -> `cleanse.ps1`) — full rebrand to C.L.E.A.N.S.E. (Cleans Leftover, Ephemeral And Neglected System Entries); new ASCII art banner (`Show-CleanseBanner`), internal references, `.NOTES` tool list, and GRIMOIRE registry entry all updated. Version bumped to 1.2.
- **GRIMOIRE key scheme** — keys reorganised into category ranges to support future growth without displacing existing tools: 1-9 Deployment, 10-19 Diagnostics & Reporting, 20-29 Security & Compliance, 30-39 Network & Remote, 40-49 Cloud & M365, 50-59 Data & Profiles. Final mapping: ORACLE 10, WARD 11, THRESHOLD 12, SENTINEL 13, AUGUR 14, CLEANSE 15, CIPHER 20, SIGIL 21, BASTION 22, RELIC 23, LEYLINE 30, SPECTER 31, LANTERN 32, AEGIS 40, VAULT 41, PHANTOM 50, ARCHIVE 51.
- **GRIMOIRE scroll buffer** — `[Console]::Clear()` replaces `Clear-Host` in `Show-Banner` and the Q-exit branch, eliminating stale menu renders when scrolling up through terminal history.
- **Bulk ASCII encoding fix** — all 24 `.ps1` files and `TechnicianToolkit.psm1` had multi-byte Unicode characters (`-`, `--`, `-`, `*`) in functional code lines (string literals, region markers, format strings) replaced with plain ASCII equivalents. Banner heredocs and block comments were left intact. Prevents Windows-1252 mojibake when PowerShell 5.1 reads files cloned without a UTF-8 BOM.
- **CLAUDE.md** — Module Exports table updated to include `Get-TKHtmlCss`, `Get-TKHtmlHead`, `Get-TKHtmlFoot`; new HTML Report Pattern section with code example and CSS class reference; Tool Distinctions section updated to reference AUGUR instead of DWARF.
- **README** — AUGUR/CLEANSE renames reflected throughout; GRIMOIRE key numbers updated in all tables; Quick Launch and Usage sections updated with AUGUR/CLEANSE one-liners.
- **Version bumps** — `grimoire.ps1` 1.1 -> 1.2; `augur.ps1` introduced at 1.1; `cleanse.ps1` introduced at 1.2.

### Fixed
- **SIGIL unattended mode bug** — `-Unattended -Categories` was silently a no-op because `$selectedKeys` was never populated in the unattended branch. Categories are now correctly parsed and applied.
- **`Get-TKConfig` defaults** — added `TeamsWebhook` to the defaults object so callers never receive null for the new key.
- **`purge.ps1` .NOTES tool list** — incorrectly listed `S.E.N.T.I.N.E.L.` as "Disk health assessment & SMART status"; corrected to `D.W.A.R.F.`.

---

## [1.0.0] — 2025-12-01 (initial public release)

### Added
- **GRIMOIRE** — central interactive hub launcher for all 21 tools; supports `-WhatIf` pass-through to tools that accept it.
- **COVENANT** — machine onboarding, Entra ID domain join, computer rename, timezone, network drives, local admin creation. Supports `-WhatIf` and `-Unattended`.
- **CONJURE** — software deployment via winget / Chocolatey with package list editor.
- **RUNEPRESS** — printer driver installation and network printer configuration.
- **FORGE** — driver detection and installation (problem devices, Windows Update, local packages).
- **RESTORATION** — automated Windows Update management via PSWindowsUpdate.
- **HEARTH** — interactive configuration wizard for org name, log directory, and tool defaults.
- **ORACLE** — system diagnostics with dark-themed HTML report generation.
- **WARD** — user account audit (roles, last logon, flags) with HTML report.
- **THRESHOLD** — disk and storage health monitoring with cleanup and old profile detection.
- **SENTINEL** — service and scheduled task monitor with event log error surfacing.
- **CIPHER** — BitLocker management (enable, disable, suspend, resume, key backup to AD/Entra). Supports `-WhatIf` and `-Unattended`.
- **SIGIL** — security baseline enforcement (telemetry, screensaver lock, UAC, autorun, firewall, guest account, password policy, RDP, audit policy, Windows Update policy, SMBv1/LLMNR/NetBIOS, credential protection). CSV action log.
- **BASTION** — Active Directory user and group management with lockout forensics.
- **RELIC** — certificate health monitor (local stores, SSL/TLS expiry) with HTML report.
- **LEYLINE** — network diagnostics and remediation (adapters, ping, DNS, port tests).
- **SPECTER** — remote execution via WinRM; runs toolkit tools on remote machines.
- **LANTERN** — network discovery and asset inventory (subnet sweep, DNS, MAC, port scan).
- **AEGIS** — Azure environment assessment (security posture, RBAC, backup coverage) with HTML report.
- **VAULT** — Microsoft 365 license and mailbox audit with MFA status.
- **PHANTOM** — user profile migration and data transfer.
- **ARCHIVE** — pre-reimaging profile backup to ZIP on local or network share.
- **TechnicianToolkit.psm1** — shared module: logging helpers, HTML utilities, config management (`Get-TKConfig` / `Set-TKConfig`), transcript helpers, privilege management.
- **config.json** — central optional configuration file; all tools function without it.

### Notes
- All tools auto-elevate to Administrator via `Invoke-AdminElevation`.
- All diagnostic/audit tools support `-Unattended` for silent/scheduled execution.
- All HTML reports use a consistent dark-themed layout with colour-coded severity.
- Module dependencies (PSWindowsUpdate, Az.*, Microsoft.Graph, RSAT) are auto-installed on first use.
