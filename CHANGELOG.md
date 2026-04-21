# Changelog

All notable changes to TechnicianToolkit are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [Unreleased]

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
