# Technician Toolkit

> A PowerShell-based toolkit for IT technicians to automate common system administration tasks — forged in the arcane arts of automation.

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3%20or%20later-blue.svg)](LICENSE)

**Free software for technicians, by technicians.** Use it, change it, share it — see [License](#license) and [Contributing](CONTRIBUTING.md).

Written and maintained by **John Joseph Bejarana** ([@CursedTechnocrat](https://github.com/CursedTechnocrat)).

---

## Two ways to run it

As of 5.0 the suite ships in two forms. They run the same 42 tools from the same
source — the application drives the scripts, it does not replace them.

| | **The application** | **The scripts** |
|---|---|---|
| What it is | One portable `.exe` with a window | 42 `.ps1` files plus the shared module |
| Needs on the target machine | Nothing | Windows PowerShell 5.1 (built into Windows) |
| PowerShell | 7, hosted inside the executable | 5.1, the one already there |
| Best for | Carrying on a USB stick; walking up to a machine | Remote sessions, scripted runs, dropping a single tool onto a box |
| Install | `winget install CursedTechnocrat.TechnicianToolkit`, or download from [Releases](https://github.com/CursedTechnocrat/TechnicianToolkit/releases) | Clone the repo, or fetch one script — see [Quick Launch](#quick-launch) |

Neither is deprecated. The scripts remain the primary documented path and stay
independently runnable; if you only need one tool on one machine, a single `.ps1`
is still the smallest thing that works.

### The application

A single self-contained executable. PowerShell 7 is hosted inside it, so the
machine it runs on needs no runtime, no modules, and no internet connection. It
extracts the suite beside itself, runs tools through a real window with live
output, and watches for the HTML reports they produce.

It requests **Administrator at launch**, once, because most of what it does needs
it — including read-only tools that do not. That is a deliberate simplification
for 5.0 and is called out here rather than buried.

> **5.0 binaries are not code-signed.** The certificate is in validation.
> SmartScreen will warn on first run, and some antivirus products may flag a
> single-file executable that unpacks scripts and runs them elevated — that is
> structurally what a dropper looks like, and a signature is what normally
> offsets it. Check the SHA-256 published in the release notes against your
> download. Signed builds follow in 5.0.1.

> **The ARM64 build has never run on real hardware.** It is built and published
> because withholding it helps nobody, but nothing has verified it beyond
> compiling and linking. If you have a Snapdragon X, Surface Pro, or any
> Windows-on-ARM machine, **[telling us what happened](https://github.com/CursedTechnocrat/TechnicianToolkit/issues)**
> is one of the most useful things you can contribute — a report that it simply
> worked is as valuable as a bug.

Building it yourself needs the .NET 8 SDK:

```powershell
dotnet publish app/TechnicianToolkit.App -c Release -r win-x64 -o publish
```

---

## LiveConnect Suite

> **Deploying remotely via Kaseya VSA LiveConnect?** Use the companion repository instead:
> ### [TechnicianToolkit-LiveConnect →](https://github.com/CursedTechnocrat/TechnicianToolkit-LiveConnect)

This toolkit is built around interactive menus, guided prompts, and real-time feedback — it is designed for technicians who are **present at the machine**, whether physically or via a full interactive remote session (RDP, Enter-PSSession, etc.).

If you are running scripts through **Kaseya VSA LiveConnect**, that shell cannot handle `Read-Host`, `ReadKey`, `Clear-Host`, or multi-step menu navigation. Those calls cause the session to hang or error immediately. The LiveConnect Suite is a separate set of scripts written from the ground up to run entirely from parameters, with no interactive calls of any kind.

| Situation | Use |
|-----------|-----|
| Sitting at the machine or in a full RDP session | **This repo** — TechnicianToolkit |
| Running through Kaseya VSA LiveConnect | **[TechnicianToolkit-LiveConnect](https://github.com/CursedTechnocrat/TechnicianToolkit-LiveConnect)** |
| Need a guided, menu-driven workflow | **This repo** — full prompts and confirmations at every step |
| Need fire-and-forget with parameter-only input | **[TechnicianToolkit-LiveConnect](https://github.com/CursedTechnocrat/TechnicianToolkit-LiveConnect)** |
| Need tools with no LiveConnect counterpart (COVENANT, CONJURE, REVENANT, WYRM, EMBALM, EMISSARY, RUNEPRESS, LEYLINE, FORGE, ZENITH, SPHINX, LANTERN, FATHOM, AUGUR, CLEANSE, ALMANAC, ORBIT, ECLIPSE, ASTERISM, CUMULUS, ORRERY, PHYLACTERY, EXHUME, VIGIL, PHOENIX, HEARTH, RITUAL, AUSPEX, WARD, SCRYER, WHETSTONE, BASILISK, ANVIL, TALON, TOTEM, HOURGLASS, GRIFFIN, WISP, PORTAL, NECROPSY, TORPOR, SOLDER, RAVEN, HALO, CARILLON, CHALICE, LODESTAR, MINOTAUR) | **This repo** — these tools are interactive by nature or require auth flows incompatible with LiveConnect |

---

## Table of Contents

- [Two ways to run it](#two-ways-to-run-it)
- [Tools Overview](#tools-overview)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Launch](#quick-launch)
- [Usage](#usage)
- [Configuration](#configuration)
- [Logging](#logging)
- [Contributing](#contributing)
- [Disclaimer](#disclaimer)
- [License](#license)

---

## Tools Overview

### Hub

| Script | Acronym | Purpose |
|--------|---------|---------|
| **grimoire.ps1** | **G.R.I.M.O.I.R.E.** — General Repository for Integrated Management and Orchestration of IT Resources & Executables | Central hub launcher — run all tools from one interactive menu |

### Deployment & Onboarding

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 1 | **covenant.ps1** | **C.O.V.E.N.A.N.T.** — Configures Onboarding Via Entra — Network, Accounts, Naming & Timezone | Machine onboarding, Entra ID domain join, and new device setup |
| 2 | **conjure.ps1** | **C.O.N.J.U.R.E.** — Centrally Orchestrates Network-Joined Updates, Rollouts & Executables | Software deployment via Windows Package Manager or Chocolatey |
| 3 | **runepress.ps1** | **R.U.N.E.P.R.E.S.S.** — Remote Utility for Networked Equipment — Printer Registration, Extraction & Silent Setup | Printer driver installation and network printer configuration |
| 4 | **forge.ps1** | **F.O.R.G.E.** — Finds Outdated Resources & Generates Equipment-updates | Driver detection & installation — problem devices, Windows Update drivers, local packages |
| 5 | **whetstone.ps1** | **W.H.E.T.S.T.O.N.E.** — Windows Hotfixes, Enforced Through Staged Testing Of Needed Enhancements | Automated Windows Update management and maintenance |
| 6 | **hearth.ps1** | **H.E.A.R.T.H.** — Hub for Environment, Admin Runtime & Toolkit Hardening | Toolkit setup wizard — configure org name, log paths, and default values |
| 7 | **ritual.ps1** | **R.I.T.U.A.L.** — Runs Integrated Tool Usage in Automation Loops | Workflow orchestrator — runs named recipes (Onboard, Retire, HealthCheck, SecuritySweep, NetworkSweep, TenantSweep) or custom PSD1 files, with rollup HTML report |
| 8 | **conduit.ps1** | **C.O.N.D.U.I.T.** — Checks Or Normalises Device Update Infrastructure Targeting | Windows Update connectivity diagnosis & repair — WSUS pointer, WinHTTP proxy, policy source, update services |
| 9 | **solder.ps1** | **S.O.L.D.E.R.** — Servicing-stack Overhaul: Locates Damage, Enacts Repair | Windows servicing diagnosis & repair — component store health and size, pending restarts, last SFC result from CBS.log, Windows Update failure codes decoded, DISM RestoreHealth + SFC |

### Diagnostics & Reporting

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 10 | **auspex.ps1** | **A.U.S.P.E.X.** — Audits, Uncovers, Surveys Performance, Events & eXceptions | System diagnostics, health assessment, and HTML report generation |
| 11 | **ward.ps1** | **W.A.R.D.** — Watches Accounts, Reviews Roles & Detects anomalies | Local user account audit with role, last logon, flags, and HTML report |
| 12 | **fathom.ps1** | **F.A.T.H.O.M.** — Free-space Analysis: Tallies Hogs, Old profiles & Mess | Disk space monitor — volume usage, low-space alerts, temp cleanup, old profile detection, HTML report |
| 13 | **vigil.ps1** | **V.I.G.I.L.** — Verifies Integrity of services, Guards & Inspects Logs | Service, task & event log monitor — health check local or remote machine, HTML report |
| 14 | **augur.ps1** | **A.U.G.U.R.** — Analyzes, Uncovers & Gauges Unit Reliability | Physical disk health — SMART status, wear prediction, failure forecast, hardware reliability, HTML report |
| 15 | **cleanse.ps1** | **C.L.E.A.N.S.E.** — Cleans Leftover, Ephemeral And Neglected System Entries | Disk cleanup — user & system temp, Windows Update cache, browser caches, Recycle Bin |
| 16 | **scryer.ps1** | **S.C.R.Y.E.R.** — System Consolidated Report Yielding Exhaustive Results | Unified diagnostic report — system info, users, disks, SMART, services in one HTML |
| 17 | **anvil.ps1** | **A.N.V.I.L.** — Audits & Notates Vendor Inventory & Lifecycle | BIOS / UEFI / firmware audit — system identity, Secure Boot posture, vendor update channels, pending Windows Update firmware, HTML report |
| 18 | **hourglass.ps1** | **H.O.U.R.G.L.A.S.S.** — Health Of Unit's Rechargeable Gauge: Life, Ageing, State & Service | Laptop battery health audit — design vs current capacity, cycle count, red/yellow/green replacement verdict, `powercfg /batteryreport` enrichment, HTML report |
| 19 | **codex.ps1** | **C.O.D.E.X.** — Compiles Output Documents into an EXhibit | Toolkit report index — scans the log directory for existing HTML reports, groups by tool, emits one rollup with relative links |
| 60 | **necropsy.ps1** | **N.E.C.R.O.P.S.Y.** — Names Each Crash, Reboot & Outage — Post-mortem Summary Yield | Crash & unexpected-reboot analysis — bugchecks, Kernel-Power 41, WHEA hardware errors, display resets, dump files, change timeline, HTML report |
| 61 | **torpor.ps1** | **T.O.R.P.O.R.** — Traces Overload, Resource Pressure & Offending Routines | Slow-machine triage — CPU, memory and disk load over a sample window, top processes grouped by name with hints, CPU power / thermal limits, boot delays Windows blamed, startup programs, app hangs, HTML report |

Diagnostics outgrew keys 10–19, so it continues at 60 rather than renumbering the keys technicians already know.

### Security

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 20 | **wyrm.ps1** | **W.Y.R.M.** — Whole-drive encrYption & Recovery Manager | BitLocker drive encryption management — enable, disable, key backup, report export |
| 21 | **basilisk.ps1** | **B.A.S.I.L.I.S.K.** — Baseline Applier: Sets Integrity, Locks In Security Knobs | Security baseline enforcement — telemetry, UAC, firewall, audit policy, password policy |
| 22 | **sphinx.ps1** | **S.P.H.I.N.X.** — Service Portal for Help-desk Identity Needs & eXpiry | Active Directory user & group management — unlock, reset, lockout forensics, stale & expiry reports |
| 23 | **phoenix.ps1** | **P.H.O.E.N.I.X.** — Pinpoints Hosts' Outdated & Expiring Notarised Identity X.509s | Certificate health monitor — local cert stores, SSL/TLS expiry, HTML report |
| 24 | **talon.ps1** | **T.A.L.O.N.** — Tracks Anomalies & Locates Otherwise-silent Nastiness | Persistence / autoruns audit — Run keys, startup folders, services, tasks, WMI subscriptions, IFEO hijacks, Winlogon, HTML report |
| 25 | **totem.ps1** | **T.O.T.E.M.** — Trusted Observer of Transparent Execution Modules | TPM health audit — presence, spec version, ownership, readiness, BitLocker dependency, endorsement key, HTML report |
| 26 | **griffin.ps1** | **G.R.I.F.F.I.N.** — Gauges Real-time protection, Inspects Findings, Freshness, Intrusions & Notifications | AV / Microsoft Defender health audit — core state, real-time / cloud / sample, signature freshness, scan history, threats, exclusions, ASR rules, third-party AV, service health, recent events, HTML report |
| 27 | **argus.ps1** | **A.R.G.U.S.** — Access Roster: Groups, Users & Scopes | Active Directory authentication & access review — domain password/lockout policy with verdicts, full name / alias / access level per account, nested group expansion, privileged group membership, review CSV, HTML report |
| 28 | **minotaur.ps1** | **M.I.N.O.T.A.U.R.** — Maps Inheritance, NTFS Owners, Trustees & Access on UNC Roots | File share & NTFS permissions review — share and NTFS access, everyone-type write, direct user grants, orphaned SIDs, broken inheritance, HTML report + CSV of every entry |

### Network & Remote

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 30 | **leyline.ps1** | **L.E.Y.L.I.N.E.** — Locates, Examines & Yields Latency, Infrastructure, Network & Endpoints | Network diagnostics & remediation — adapters, ping, DNS, port tests, IP renew, stack reset |
| 31 | **emissary.ps1** | **E.M.I.S.S.A.R.Y.** — Executes Modules In Sessions Sent Across Remote sYstems | Remote machine execution via WinRM — run toolkit tools without physical access |
| 32 | **lantern.ps1** | **L.A.N.T.E.R.N.** — Locates & Audits Network Topology, Enumerating Resources & Nodes | Network discovery — subnet ping sweep, DNS lookup, MAC addresses, port scan, HTML report |
| 33 | **wisp.ps1** | **W.I.S.P.** — Wireless Inventory & Security Profiler | Wi-Fi profile audit — saved profiles via XML export, authentication/cipher tier, auto-connect risk, hidden SSID, MAC randomisation, optional key cleartext, HTML report |
| 34 | **portal.ps1** | **P.O.R.T.A.L.** — Profiles, Observes & Reports Tunnels, Authentication & Links | VPN / Always-On VPN audit — built-in user and all-user connections, auth/encryption tier, app triggers, NRPT, tunnel interfaces, third-party clients (Cisco / Palo Alto / Pulse / OpenVPN / WireGuard / Tailscale / WARP / etc.), HTML report |
| 35 | **lodestar.ps1** | **L.O.D.E.S.T.A.R.** — Locates Our Domain controller, Establishes Secure Trust And Repairs | Domain trust & secure channel diagnosis and repair — DC discovery, DNS, DC ports, clock skew, `nltest` secure channel, machine password reset |

### Cloud & Identity

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 40 | **zenith.ps1** | **Z.E.N.I.T.H.** — Zone-wide Evaluation of Networks, Identity, Tenancy & Hardening | Azure subscription assessment — security posture, RBAC, backup coverage, Advisor alerts, HTML report |
| 41 | **almanac.ps1** | **A.L.M.A.N.A.C.** — Assigned Licenses, Mailboxes, Authentication & Notable Account Counts | Microsoft 365 license & mailbox audit — license assignments, MFA status, shared mailboxes, HTML report |
| 42 | **orbit.ps1** | **O.R.B.I.T.** — Observes Registered devices, Baselines, Inventory & Timeliness | Intune / MDM compliance audit — managed devices, compliance state, stale devices, configuration profiles, HTML report |
| 43 | **eclipse.ps1** | **E.C.L.I.P.S.E.** — Entra Credentials: Lapsed, Idle, Privileged, Stale & External | Entra ID identity hygiene audit — guests, privileged roles, password-never-expires, stale admins, disabled-but-licensed, HTML report |
| 44 | **asterism.ps1** | **A.S.T.E.R.I.S.M.** — Audits Settings of Teams: Exposure, Roles, Inactivity, Sprawl & Membership | Microsoft Teams audit — orphan teams, public teams, guest membership, large teams, stale teams, HTML report |
| 45 | **cumulus.ps1** | **C.U.M.U.L.U.S.** — Catalogs Usage, Members, Unowned, Links, Untouched & Shared sites | SharePoint Online audit — site inventory, storage, external sharing, ownerless sites, stale sites, HTML report |
| 46 | **orrery.ps1** | **O.R.R.E.R.Y.** — Outlines Reliances: Roles, Entitlements, Rules & whY things break | Entra ID group dependency audit — "what breaks if we delete this group?" — group-based licensing, Conditional Access, enterprise apps, directory roles, nested membership, AUs, Intune, SharePoint, Exchange, Azure RBAC, HTML report |
| 47 | **raven.ps1** | **R.A.V.E.N.** — Reviews Auto-forwarding, Vulnerable Exchange settings & Nefarious rules | Exchange Online mailbox security audit — external forwarding, suspicious inbox rules, auto-forward policy, SMTP AUTH, auditing, delegation, SPF / DKIM / DMARC, HTML report |
| 48 | **halo.ps1** | **H.A.L.O.** — Holistic Access-policy Logic Overview | Entra ID Conditional Access posture — MFA and legacy-auth baseline, admin role coverage, report-only and disabled policies, exclusions, emergency access, named locations, HTML report |
| 49 | **carillon.ps1** | **C.A.R.I.L.L.O.N.** — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers | Teams Phone call queue & auto attendant audit — every queue's agents by display name with opt-in, voice and account state, overflow / timeout / no-agent routing, attendant menus and after-hours flows, resource accounts, unreachable queues, HTML report + agent CSV |
| 70 | **chalice.ps1** | **C.H.A.L.I.C.E.** — Checks Health, Activation, Logins & Identity of Click-to-run Editions | Microsoft 365 Apps client diagnosis & repair — version, channel and support dates, license tokens, cached Office accounts, modern auth / WAM settings, AAD token broker, device PRT, Teams cache; resets sign-in and activation, clears the Teams cache, Quick Repair, update now, HTML report |

Cloud & Identity outgrew keys 40–49, so it continues at 70.

### Data & Migration

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 50 | **revenant.ps1** | **R.E.V.E.N.A.N.T.** — Relocates, Extracts, Validates Environments, Networks, Accounts 'N Transfers | Profile migration and data transfer between machines or profiles |
| 51 | **embalm.ps1** | **E.M.B.A.L.M.** — Encapsulates My Belongings As Lasting Mementos | Pre-reimaging profile backup — ZIP to local path or network share |
| 52 | **phylactery.ps1** | **P.H.Y.L.A.C.T.E.R.Y.** — Pre-migration Health of Your Libraries: Accounts, Client, Tethering, Errors, Readiness & Yield | OneDrive Known-Folder-Move pre-migration validator — client, accounts, KFM status, content volume, recent sync errors, HTML report |
| 53 | **exhume.ps1** | **E.X.H.U.M.E.** — Enumerates, eXposes & Hunts Unmigrated Mail Entries | Outlook PST / OST discovery — profile walk, drive-wide file scan, orphan / oversize / stale detection, HTML report |

---

## G.R.I.M.O.I.R.E.

The central hub for the Technician Toolkit. Presents a categorized, interactive menu to launch any tool without navigating the file system. After a tool completes, control returns to the GRIMOIRE menu automatically.

- Auto-elevates to Administrator on first launch if not already elevated
- Validates that each script file exists before attempting to launch it
- Downloads missing scripts from GitHub automatically on first use
- Returns to the hub menu after each tool finishes or errors out
- All tools remain independently runnable without the hub

---

## Deployment & Onboarding

### C.O.V.E.N.A.N.T.

Guides a technician through the full setup of a new Windows machine.

- Pre-flight check of current domain and Entra ID join status
- Optional computer rename with hostname validation
- Entra ID (Azure AD) domain join — UPN and password entered securely in the terminal
- Network drive mapping — repeatable, supports per-share credentials and persistent mapping
- Local administrator account creation (or password reset if account exists)
- Timezone configuration with common presets or manual entry
- Action summary with 30-second reboot countdown and Escape to cancel

---

### C.O.N.J.U.R.E.

Manages software deployment using the Windows Package Manager (winget) or Chocolatey.

- Supports both winget and Chocolatey package managers (user selectable at runtime)
- Installs required and optional software packages defined at the top of the script
- **All operator prompts are front-loaded** — package manager, operation, Adobe edition, optional-software selection, and any custom package IDs are gathered up front, then the install/upgrade work runs start-to-finish with no further input (no babysitting the machine)
- **Custom packages** — after the curated optional list, the operator can type their own comma-separated winget/Chocolatey package IDs to install alongside the built-ins (passed through verbatim to the selected manager)
- **Adobe Acrobat edition prompt** — choose Reader (`Adobe.Acrobat.Reader.64-bit`) or Pro (`Adobe.Acrobat.Pro`) at runtime; under Chocolatey, which has no Acrobat Pro package, a Pro request installs Reader with a note to sign in and run the in-app upgrade against a Pro license
- **Direct download installs** — for software with no winget or Chocolatey package, most often an RMM agent whose installer is a per-tenant link with a token in the URL. Paste an HTTPS link (or pass `-DirectUrl`) and the installer is fetched and run silently. Because this downloads an executable and runs it elevated, it is deliberately stricter than the package-manager path: plain HTTP is refused outright, the payload is identified by its own magic bytes rather than by the URL — so a link whose token has expired and returns an HTML error page with HTTP 200 is caught instead of executed — and a pinned SHA-256, when given, is verified *before* anything runs. The computed hash is always printed and recorded so it can be pinned next time. The installer is deleted afterwards.
- **Repeat deployments** — direct downloads can be stored in `config.json` under `Conjure.DirectDownloads` (`Name` / `Url` / `Args` / `Sha256`), so a tenant's RMM agent is entered once and then deploys unattended on every machine after that. A missing or broken package manager no longer aborts the run when direct downloads are queued — they need neither winget nor Chocolatey
- Upgrade-all mode for keeping existing packages current
- Tracks and displays installation status per package
- `-WhatIf` previews every install, upgrade, and direct download without fetching or running anything

**Default required packages:** Microsoft Teams, Microsoft 365, 7-Zip, Google Chrome, Zoom, plus Adobe Acrobat (Reader or Pro — prompted at runtime)

**Default optional packages:** Zoom Outlook Plugin, Mozilla Firefox, Dell Command Update, Asana, Google Earth Pro

---

### R.U.N.E.P.R.E.S.S.

Automates printer driver extraction, installation, and network printer configuration via a command-line interface.

- Supports already-extracted driver folders (bare INF) plus ZIP, EXE, and MSI packages
- `-DriverPath` points the tool at any folder or file, so the driver need not be copied next to the script
- Installs an INF in the two steps Windows requires: `pnputil` to stage it in the DriverStore, then `Add-PrinterDriver` to register it with the print spooler — staging alone leaves the driver invisible to `Get-PrinterDriver` and unusable by `Add-Printer`
- Picks the INF matching the machine's architecture and reads the model name out of it, so x86 INFs in a combined package are skipped
- Verifies EXE and MSI installs against the spooler rather than trusting the exit code, and warns when a vendor bootstrapper exits cleanly having registered nothing
- Configures network printers via IP (TCP/IP port) or UNC path post-install, with a printui fallback for devices that are offline at setup time
- Generates a timestamped installation log (CSV) in the script directory
- `-WhatIf` previews each driver install (pnputil + spooler registration / EXE silent / msiexec) and skips the network-printer stage, leaving no files or printers behind

---

### F.O.R.G.E.

Audits the device tree for driver problems and automates driver installation from multiple sources.

- Scans all devices for errors with human-readable error descriptions (missing, corrupted, cannot start, etc.)
- Checks Windows Update for available driver updates via PSWindowsUpdate (auto-installed if missing)
- Installs drivers from the current folder: ZIP (extracts and runs pnputil on INF), bare INF, EXE (silent), MSI (quiet)
- Exports a full driver inventory CSV with device name, driver version, date, and manufacturer
- Cleans up extracted driver staging folders automatically
- `-WhatIf` lists pending Windows Update drivers and previews the extension-specific handler for each local file (extract / pnputil / silent EXE / msiexec) without running any of them

---

### W.H.E.T.S.T.O.N.E.

Automates Windows Update detection, installation, and reboot handling with minimal user intervention.

- Disables sleep and display timeout for the duration of the run; restores settings on exit
- Ensures NuGet provider and PSWindowsUpdate module are installed and current
- Installs available updates (drivers excluded) with no forced reboot
- Checks reboot status and prompts only when required
- 30-second reboot countdown with Escape key cancel
- `-WhatIf` lists every pending update that would be installed and skips both the install and the reboot decision

---

### H.E.A.R.T.H.

Interactive setup wizard for the Technician Toolkit — configure all settings without hand-editing JSON.

- Step-by-step wizard covers all seven configuration fields with descriptions, hints, and live validation
- Configures: organization name, log/report directory, Teams webhook URL, EMBALM default destination, REVENANT default destination, COVENANT default timezone and local admin username
- Path fields validate on entry — prompts to create missing directories automatically
- View current configuration with color-coded status: green = configured, yellow = empty or path not found
- Edit individual fields without re-running the full wizard
- **Environment checks**: PowerShell version, admin status, module presence, winget, Chocolatey, RSAT, Microsoft.Graph, Az, and log directory write access
- Configuration reset with `YES` confirmation guard
- All settings persisted to `config.json` in the toolkit directory
- `-Unattended` displays current config and runs environment checks silently

---

### R.I.T.U.A.L.

Workflow orchestrator. Runs an ordered sequence of toolkit scripts as a single named recipe and rolls the results up into one HTML report with per-step status, duration, and clickable links to each child report.

- **Built-in recipes** (pass via `-Recipe <Name>`):
  - `Onboard` — new machine bring-up: COVENANT → BASILISK → CONJURE → WYRM → AUSPEX → PHOENIX
  - `Retire` — pre-reimage workflow: PHYLACTERY → EXHUME → EMBALM → CLEANSE
  - `HealthCheck` — quarterly machine review (read-only): AUSPEX → WARD → FATHOM → AUGUR → VIGIL → PHOENIX → GRIFFIN → NECROPSY → TORPOR → SOLDER (audit only; TORPOR adds a 30-second load sample)
  - `SecuritySweep` — endpoint security posture (read-only): BASILISK → TALON → TOTEM → GRIFFIN → PHOENIX
  - `NetworkSweep` — endpoint network posture (read-only): LEYLINE → LANTERN → WISP → PORTAL → LODESTAR (audit only)
  - `TenantSweep` — cloud tenant posture: ZENITH → ALMANAC → ORBIT → ECLIPSE → ASTERISM → CUMULUS → RAVEN → HALO → CARILLON (RAVEN signs in to Exchange Online, and CARILLON to Microsoft Teams, separately from the Graph tools)
- **Custom recipes** via `-RecipeFile path\to\recipe.psd1` — a hashtable with `Name`, `Description`, and an ordered `Steps` array (each step specifies `Tool`, `Args`, `StopOnError`, `Label`)
- Per-step log-directory snapshot — any new files produced during a step are attributed to that step and linked from the rollup report
- Default behaviour: abort on first failure. Pass `-ContinueOnError` to run every step regardless, unless the step's own `StopOnError = $true` overrides
- Interactive menu lists every built-in recipe with its step summary; `-Recipe` or `-RecipeFile` triggers a headless run and exits with `0` on full success, `1` on any failure
- Rollup HTML saved as `RITUAL_<timestamp>.html` with summary cards (outcome, total / succeeded / failed / skipped / duration) and step-by-step results
- Auto-elevates to Administrator (inherits the requirements of the tools it orchestrates)

---

### S.O.L.D.E.R.

Diagnoses and repairs the Windows servicing stack — the component store (WinSxS) that updates, optional features and system-file repair all draw on. Reach for it when updates fail with `0x800f081f`, `0x80073712` or `0x800f0831`, when a feature will not install, or when system files are damaged.

- **Audit** (read-only, default):
  - **Pending restarts**, split into those that block servicing (CBS `RebootPending` / `PackagesPending`, `pending.xml`, Windows Update `RebootRequired`) and those that do not (file renames, a computer rename)
  - **Component store health** via `DISM /CheckHealth` (or a full `/ScanHealth` with `-Deep`), and size / reclaimable packages via `/AnalyzeComponentStore`. DISM runs with `/English`, so the output parses on any UI language
  - **The last System File Checker result**, read from `CBS.log`: files repaired and files SFC could not repair
  - **Windows Update install failures** from the last 30 days, each code decoded and classed as *Store* (fixed here), *Client* (fixed by CONDUIT), *Reboot*, *Space*, *Access* or *Other*
  - The Windows Modules Installer service, free space on the system drive, and **where DISM will look for repair files** — a WSUS-managed machine asks WSUS, which serves no repair content, so `RestoreHealth` fails with `0x800f081f` unless policy sends repairs to Windows Update or `-Source` is given
- **Repair** (`-Action Repair`): `DISM /RestoreHealth` → `sfc /scannow` → re-check the store. SFC is skipped when RestoreHealth fails, since it would copy from a damaged store. A pending servicing restart stops the repair unless the technician chooses to continue. A disabled Windows Modules Installer is set back to Manual. `-Cleanup` (or a prompt, when DISM recommends it) adds `/StartComponentCleanup`
- `-Source 'WIM:D:\sources\install.wim:6'` (or `ESD:…install.esd:<index>`) repairs from mounted install media of the same build and edition, with `/LimitAccess`
- Never runs `/ResetBase` (irreversible — updates can no longer be uninstalled), never edits Windows Update policy, never resets the update cache. `-WhatIf` previews every repair
- Verdict: Broken / Attention / Healthy

> **CONDUIT vs SOLDER vs WHETSTONE:** CONDUIT fixes the update client's *connection*; SOLDER fixes the *component store* the update installs into; WHETSTONE *installs* the updates. A failure code SOLDER classes as *Client* is CONDUIT's to fix.

---

## Diagnostics & Reporting

### A.U.S.P.E.X.

Audits the current state of a Windows machine and exports a formatted HTML report to the script directory.

- Hardware inventory: CPU, RAM, disk usage with visual bar charts, model and serial number
- OS details: version, build, architecture, install date, activation status
- Network configuration: all active adapters with IP, MAC, gateway, and DNS
- System health: uptime, last reboot time, battery status (laptops)
- Pending Windows Update scan (read-only, no installation)
- Installed software list sourced from registry
- Recent event log errors and critical events (last 24 hours)
- **Security & AV status**: Windows Defender real-time protection, definition age, last scan; third-party AV products via SecurityCenter2
- Dark-themed HTML report with color-coded indicators and status badges

---

### W.A.R.D.

Audits all local user accounts and exports a dark-themed HTML report to the script directory.

- Lists all local accounts: enabled/disabled status, last logon, password info
- Identifies group memberships — flags all Administrator accounts
- Flags potentially risky accounts: no password required, password never set, stale (no logon in 90+ days)
- **LAPS status**: which implementation governs the local administrator password (Windows LAPS, Windows LAPS in legacy emulation, or legacy Microsoft LAPS) and from which policy source (Intune / CSP, Group Policy, local configuration — highest precedence wins); where the password is backed up (Entra ID / Active Directory) and whether the device is actually joined there; which account is managed and whether it exists; whether its password has been rotated within the policy's age; recent errors from the Windows LAPS operational log. A machine with enabled local admins and no LAPS policy is flagged
- Console summary with highlighted flagged accounts
- HTML report with color-coded badges and summary cards
- Report saved to script directory as `WARD_<timestamp>.html`

---

### F.A.T.H.O.M.

Audits physical disk and volume health, flags space problems, and performs optional cleanup.

- Physical disk health status via `Get-PhysicalDisk` — Healthy / Warning / Unhealthy
- Disk operational status and media type (SSD / HDD / Unspecified) via `Get-Disk`
- Volume space summary for all lettered drives: used, free, total, percentage
- **Warning** flagged at < 15% free space; **Critical** at < 5% free space
- Disk cleanup: Windows Temp, user Temp, Recycle Bin, Windows Update cache
- Old profile detection: user profile folders not accessed in 90+ days
- Dark-themed HTML report with color-coded status badges
- `-Unattended` for silent health check and HTML export

---

### V.I.G.I.L.

Audits Windows services, scheduled tasks, and recent event log errors — locally or against a remote machine.

- Critical service audit: checks a predefined set of essential services (WinDefend, Spooler, BITS, WMI, W32Time, and more)
- Flags stopped or non-automatic services; offers one-at-a-time restart with confirmation
- Scheduled task audit: lists all active non-Microsoft tasks with last/next run and status
- Event log sweep: Warning and Error events from System and Application logs in the last 24 hours
- Supports remote execution via WinRM with `-Target HOSTNAME`
- Dark-themed HTML health report with color-coded service status badges
- `-Unattended` for silent report export; `-Unattended -Target HOSTNAME` for remote

---

### A.U.G.U.R.

Inspects every physical disk in the system for hardware-level reliability issues and SMART failure prediction.

- Physical disk health status — Healthy, Warning, Unhealthy — sourced from SMART and WMI
- Disk operational status and media type (SSD / HDD / Unspecified) via `Get-PhysicalDisk`
- SMART failure prediction flag — surfaces any disk with a predicted imminent failure
- Bus type and model details for every physical disk
- Volume integrity status across all lettered volumes
- Dark-themed HTML report with color-coded disk status badges
- `-Unattended` for silent scan and HTML export

> **AUGUR vs FATHOM:** AUGUR answers "is this drive about to fail?" (SMART/hardware).
> FATHOM answers "is this drive running out of space?" (volume usage/cleanup).

---

### A.N.V.I.L.

BIOS / UEFI / firmware audit that answers the technician's question "is this machine's firmware in a supportable state?" before a deployment, reimage, or OS upgrade.

- System identity: manufacturer, model, system SKU / service tag, UUID, serial number, BIOS vendor and version, BIOS release date, BIOS age in years
- UEFI / Secure Boot posture: firmware type (UEFI vs legacy BIOS), system-disk partition style as a cross-check, Secure Boot state (Enabled / Disabled / Unsupported)
- Vendor firmware-update channel detection — checks known install paths for Dell Command | Update, HP Image Assistant, HP Support Assistant, Lenovo System Update / Vantage, and Microsoft Surface UEFI Configurator. Vendor is auto-detected from `Win32_ComputerSystem.Manufacturer`.
- Windows Update driver / firmware scan via PSWindowsUpdate (installed on demand); lists every pending update with KB ID, size, and severity
- Readiness verdict: red / yellow / green based on firmware type, Secure Boot state, BIOS age (≥ 2 years flagged), vendor tooling presence, and WU backlog
- Dark-themed HTML report with OrgName prefix and six summary cards
- Auto-elevates and telemeters WU failures via `Write-TKError`

---

### H.O.U.R.G.L.A.S.S.

Laptop battery health audit. Surfaces the three numbers that matter for a retirement decision (design capacity, current full-charge capacity, cycle count) and applies industry-consensus thresholds so a technician can answer "is this battery still worth it, or are we buying a replacement?"

- **Data source**: `ROOT\WMI` battery classes — `BatteryStaticData` (design capacity), `BatteryFullChargedCapacity` (current full charge), `BatteryCycleCount` (cycles), `BatteryStatus` (instantaneous voltage, charge/discharge rate). Merges on `InstanceName` so multi-battery laptops emit one row per cell. Joins to `Win32_Battery` for the user-friendly name and chemistry.
- **Thresholds**: capacity health ≥ 80% green / 60-80% yellow / < 60% red; cycle count < 300 green / 300-500 yellow / ≥ 500 red. Worst value across both dimensions drives the verdict.
- **Verdict**: HEALTHY / REPLACEMENT SOON / REPLACE NOW / DATA INCOMPLETE / NO BATTERY (explicit desktop/VM case).
- **Dark HTML report** with six summary cards (verdict, battery count, best/worst health, max cycles, thresholds reference) and a full per-battery detail table with Wh values and colour-coded badges.
- **`powercfg /batteryreport` enrichment**: also runs `powercfg /batteryreport /xml` and parses the result to surface data the live `ROOT\WMI` classes do not expose — per-battery serial number and manufacture date, full-charge capacity history (degradation trend over the lifetime of the machine), Windows runtime estimates at both current full charge and original design capacity (so the technician can quote the runtime lost to wear), and aggregated AC vs DC time totals across the runtime history. The full Microsoft-formatted HTML is saved alongside `HOURGLASS_*.html` and linked from the new section.
- Auto-elevates; read-only.

---

### C.O.D.E.X.

Toolkit report index builder. Walks the configured log directory, finds every TechnicianToolkit-generated HTML report (filename ending in `_YYYYMMDD_HHMMSS.html`), groups them by tool prefix, and emits a single dark-themed HTML rollup with relative links to each child report.

- **Data source**: filesystem only — `Get-ChildItem` over the configured log directory (or a `-LogDir` override), filtered to files matching `<TOOL>_YYYYMMDD_HHMMSS.html`. Files outside that pattern are skipped on purpose so browser-saved pages or hand-renamed copies don't pollute the index. CODEX excludes its own outputs from the scan.
- **Grouping**: by tool prefix (first underscore-delimited segment) so HOURGLASS, AUSPEX, AUGUR, etc. each get their own section. Variants like `HOURGLASS_battery_report_*` and `SPHINX_StaleAccounts_*` are surfaced as a separate badge inside the parent tool's section.
- **Filtering**: optional `-DaysBack <int>` limits the index to reports younger than N days.
- **Dark HTML report** with six summary cards (total reports, distinct tools, last-7-days count, total disk size, newest, oldest) and one section per tool with a per-report table (timestamp, variant, file link, size). Links are relative to the log directory so the rollup stays clickable when the folder is zipped or moved to a ticket attachment.
- Distinct from `R.I.T.U.A.L.` — RITUAL composes a fresh recipe run and produces a rollup of *what it just ran*; CODEX answers "what reports already exist on disk?" for ad-hoc work that didn't go through a recipe.

---

### C.L.E.A.N.S.E.

Frees disk space by cleaning common junk accumulation points across the system.

- User temp folders (`%TEMP%`, `%LOCALAPPDATA%\Temp`)
- System temp folder (`C:\Windows\Temp`)
- Windows Update download cache (`SoftwareDistribution\Download`) — stops and restarts the service safely
- Recycle Bin — all users
- Browser caches — Chrome, Edge, and Firefox across all user profiles
- Shows estimated space for each category before cleaning; reports total freed space at the end
- `-Unattended` cleans all categories silently
- `-WhatIf` previews what would be cleaned without deleting anything

---

### S.C.R.Y.E.R.

Produces a single consolidated HTML report covering the most commonly requested diagnostic checks — designed for machine handoffs, ticket attachments, and audit records.

- Section 1 — **System overview**: OS caption, build, install date, CPU, memory, uptime
- Section 2 — **User accounts**: local user inventory with enabled/disabled status and admin flagging
- Section 3 — **Disk space**: volume usage with warning/critical thresholds on free space
- Section 4 — **Disk health**: SMART status and physical disk reliability via `Get-PhysicalDisk`
- Section 5 — **Services & scheduled tasks**: critical service state and non-Microsoft tasks
- Single dark-themed HTML report with summary cards and nav anchors for each section
- `-Unattended` for silent run; `-OutputPath <dir>` to redirect the report destination

> **SCRYER vs individual diagnostic tools:** SCRYER is a one-shot snapshot that rolls five checks into one file. Reach for AUSPEX, WARD, FATHOM, AUGUR, or VIGIL when you want a deeper single-domain report.

---

### N.E.C.R.O.P.S.Y.

Answers "why does this machine keep crashing or rebooting?" from the evidence Windows leaves behind after an unplanned stop. Read-only.

- **Bugchecks** from WER event 1001, with the code mapped to its name and the area it usually implicates (driver, memory, storage, graphics, hardware, power, system)
- **Kernel-Power 41** classified three ways: a blue screen, a power-button press (the user forced off a hung machine), or neither — a sudden power loss or a freeze too hard to crash
- **Dump files** in the minidump folder and `MEMORY.DMP`, with the bugcheck code read straight out of each dump's header — so a crash whose event has rolled out of the log is still counted
- **WHEA hardware errors**, fatal and corrected, with the failing component; **display driver resets** (TDR, event 4101)
- **Timeline** interleaving crashes with driver installs and Windows updates in the same window, so "it started after the update" is visible at a glance
- **Crash dump readiness**: dump type, page file, automatic restart — dumps disabled or no page file means the next crash leaves nothing to analyse
- Reliability Monitor stability index and boot count for the window
- **Next steps** per implicated area, pointing at the toolkit tool that goes deeper (AUGUR for storage, FORGE for drivers, ANVIL for firmware, HOURGLASS for batteries)
- Verdict: Crashing / Unstable / Stable — a readiness gap such as disabled dumps never downgrades a machine that has not crashed
- `-Days <1-365>` sets the look-back window (default 30)

NECROPSY reads the bugcheck code and parameters, not the stack. Naming the faulting driver needs a debugger: open a dump it lists in WinDbg and run `!analyze -v`.

---

### T.O.R.P.O.R.

Answers "why is this PC slow?" by watching the machine for a sample window (30 seconds by default) while the slowness is happening, then naming what is using it. Read-only.

- **CPU**: average and peak load, processor queue length per logical processor, and the `% Performance Limit` counter — a CPU held back by the power plan, battery saver or thermal throttling shows up even when the load looks moderate
- **Memory**: lowest available RAM, commit charge against the commit limit, hard page faults, the page file, and whether the installed RAM is simply too little
- **Disks**: busy time, response time, IOPS and throughput per physical disk from raw counters (the formatted class rounds response time to whole seconds), media type, and a spinning system disk; free space on the system drive
- **Top processes** by CPU, memory and disk I/O, **grouped by name** so thirty browser processes read as one browser; CPU is the share of the whole machine, so the rows add up to the total. Well-known offenders (Defender, Search, servicing, OneDrive, WMI, browsers, Teams, Outlook…) carry a hint, and a busy `svchost` lists the services inside it
- **Context**: uptime (with Fast Startup on, *Shut down* does not reset it), the power plan and Windows 11 power mode, enabled startup programs, boot times and the apps / drivers / services / Group Policy Windows blamed for slow boots (Diagnostics-Performance log), and applications that keep hanging (Application Hang 1002)
- Counters come from WMI performance classes rather than `Get-Counter`, whose counter paths are translated on non-English Windows
- Verdict: Struggling / Strained / Healthy
- `-SampleSeconds <5-300>` sets the window (default 30); the interactive menu also offers a 120-second sample for slowness that comes and goes

> **TORPOR vs AUSPEX / TALON:** AUSPEX is a general health snapshot and TALON audits startup entries as persistence. TORPOR measures load while the machine is slow and lists startup programs for their cost, not their safety.

---

## Security

### W.Y.R.M.

Manages BitLocker drive encryption across all volumes, from a menu or unattended with `-Action`.
All changes go through `manage-bde` and all reads through the `Win32_EncryptableVolume` CIM class,
so it does not depend on the BitLocker PowerShell module and runs the same under PowerShell 5.1 and 7.

- Displays current encryption status for all drives on launch
- Enable BitLocker: TPM + recovery password on the OS drive; recovery password + auto-unlock on other drives. Starts used-space-only XTS-AES 256 and retries full-volume if the disk rejects it
- On an already-encrypted drive with protection off (suspended, or an OEM clear-key volume), adds any missing protector and turns protection on rather than re-encrypting
- Recovery key displayed once encryption starts
- Disable BitLocker (full decryption) with confirmation prompt
- Back up recovery keys to Active Directory or Entra ID
- View recovery key ID and password for any encrypted drive
- Suspend BitLocker for BIOS/firmware updates (auto-resumes after one reboot)
- Resume suspended BitLocker protection
- Export a drive-status + recovery-key HTML report. Unattended: `-Action Export [-OutputPath <dir>]`

---

### B.A.S.I.L.I.S.K.

Applies a standardized security and configuration baseline to a Windows machine. Pairs naturally with C.O.V.E.N.A.N.T. as a post-onboarding hardening step.

- Select individual categories or apply all at once
- **Telemetry & Privacy** — minimize Windows telemetry, disable advertising ID and ink personalization
- **Screensaver & Display Lock** — 10-minute lock timeout, password required on resume, machine-level inactivity policy
- **UAC** — enable UAC, set to Always Notify, prompt on secure desktop
- **Autorun & Autoplay** — disable for all drive types (machine and user scope)
- **Windows Firewall** — enable all profiles, block inbound on Public profile
- **Guest Account** — disable if present
- **Password Policy** — minimum length 8, max age 90 days, lockout after 5 attempts
- **Remote Desktop** — enable (with NLA) or disable with firewall rule update
- **Audit Policy** — enable logon, logoff, lockout, policy change, and account management auditing
- **Windows Update Behavior** — exclude driver updates, no auto-reboot with logged-on users
- **SMBv1, LLMNR, NetBIOS, LSA PPL, NoLMHash, RDP Restricted Admin** — additional hardening controls
- Domain Group Policy takes precedence over local settings where applicable
- Changes logged to `BASILISK_BaselineLog_<timestamp>.csv` in the script directory

---

### S.P.H.I.N.X.

Interactive Active Directory user and group management tool. Requires RSAT (auto-installed if missing).

- Search and view AD users by name, UPN, or SAM account name
- Unlock locked-out accounts
- Reset user passwords with force-change-on-next-logon option
- Enable and disable user accounts
- View and modify group memberships — add or remove from security/distribution groups
- **Account lockout forensics** — queries the PDC Emulator Security log (Event ID 4740) to identify the source machine behind each lockout, with built-in remediation guidance
- **Password expiry report** — console view of users with passwords expiring within a configurable threshold
- **Password expiry HTML export** — dark-themed report with Expired / Critical / Warning summary cards
- **Stale account report** — identifies accounts inactive for 90+ days, exports dark-themed HTML report
- `-Unattended -Action StaleReport` for silent stale account HTML export
- `-Unattended -Action PasswordExpiryReport` for silent password expiry HTML export

---

### A.R.G.U.S.

Answers the two questions a customer security review asks: *what authentication controls are in place*, and *who has an account here and what can each of them do?* Produces the roster an MSP hands to a client for sign-off and cleanup — every account as **Full Name / alias / Role** — plus the evidence behind each role. Read-only; ARGUS never modifies the directory.

- **Authentication policy, answered in prose** — reads the default domain password and lockout policy and scores each setting **Strong / Acceptable / Weak** against a stated baseline, then writes the four items an access questionnaire asks for (password length and complexity, password history, password expiration, lockout for failed attempts) as finished sentences the technician can paste into the response
- **Fine-grained password policies (PSOs) enumerated** — a PSO overrides the domain default for the principals it targets, so answering from the default alone can be flatly wrong; any that exist are listed with their precedence and targets
- The meaningful zeros are interpreted rather than printed: lockout threshold `0` means lockout is **off** (Weak), max password age `0` means passwords **never expire** (reported, not failed — NIST SP 800-63B advises against routine expiry), and lockout duration `0` means **locked until an administrator unlocks**, the strictest setting rather than the weakest
- **Effective, not direct, membership** — group membership is expanded server-side with the LDAP in-chain matching rule (`1.2.840.113556.1.4.1941`), so an account that reaches Domain Admins three nested groups deep is still reported as a Domain Administrator
- **Primary-group membership resolved separately** — an account whose *primary* group has been switched to Domain Admins does not appear in that group's member list at all, and is folded in explicitly
- **Groups resolved by well-known RID, not by name** — a renamed or localised `Domain Admins` is still found
- **Four role tiers**:
  - **Domain Administrator** — Enterprise / Schema / Domain Admins, `BUILTIN\Administrators`, Group Policy Creator Owners
  - **Delegated Administrator** — Account / Server / Backup / Print Operators, DnsAdmins, Key Admins, Enterprise Key Admins, Cert Publishers, Remote Management Users
  - **Elevated (Custom Group)** — member of a customer-created group whose name matches `-AdminGroupPattern` (the "IT Admins" / "Helpdesk Operators" groups a built-ins-only audit misses)
  - **Standard User** — no privileged membership found
- **Account type** is reported separately from role — `User`, `Service Account` (SPN present or `svc_`-style naming), `Built-in (KDC)`, `Built-in (Guest)`
- **Review flags for the cleanup conversation** — never signed in, inactive beyond `-StaleDays`, locked out, password expired / never expires / never set, trusted for unconstrained delegation, unused administrator, and *former privileged account* (an `adminCount` stamp with no current privileged membership — an ex-administrator whose ACL is still detached from its OU)
- **Review CSV** alongside the HTML — the same roster with two deliberately empty columns (`Action (Keep/Disable/Delete)`, `Customer Notes`) so the customer can mark it up and hand it straight back
- Report sections: access levels at a glance, privileged accounts, full roster, privileged group membership (with effective members per group), and accounts flagged for review
- `-IncludeDisabled` widens the roster past enabled accounts; `-SearchBase` scopes to one OU; `-Server` targets a specific domain controller; `-StaleDays` sets the inactivity threshold (default 90); `-SkipCustomGroupScan` limits the audit to built-in groups; `-NoCsv` suppresses the CSV
- Requires the RSAT ActiveDirectory module (offered for install if missing)

---

### M.I.N.O.T.A.U.R.

File share and NTFS permissions review — answers "who has access to this share?" the way ARGUS answers it for the domain and WARD for the local machine. Read-only.

- Every non-administrative SMB share with its **share permissions** and the **NTFS permissions** at its root
- A walk down each share (`-Depth`, default 2 folder levels; capped at 5,000 folders) recording every folder with **explicit permissions or broken inheritance** — folders that only inherit are not repeated
- **Broad write**: Everyone, Authenticated Users, Users or Domain Users granted write in NTFS *and* let through by the share permissions — the combination that lets any account, or ransomware running as one, change the data. Broad read is reported separately, and expected on `NETLOGON` / `SYSVOL`
- **Direct user grants** (access given to a person instead of a group), **orphaned SIDs** left by deleted accounts, **Deny** entries, and folders the elevated session still could not read
- `-Path <folder>` reviews one folder tree instead of the shares
- HTML report plus `MINOTAUR_<timestamp>.csv` holding every access-control entry scanned (share, folder, identity, kind, level, allow/deny, inherited) — ready for a customer access review
- Auto-elevates

> **MINOTAUR vs ARGUS vs WARD:** all three are access reviews, at different scopes — MINOTAUR for file shares, ARGUS for Active Directory, WARD for one machine's local accounts.

---

### P.H.O.E.N.I.X.

Monitors certificate health across the local machine and remote hosts — surfaces expiring and expired certificates before they cause outages.

- Audits local Windows certificate stores: Personal (My), Intermediate CA, Trusted Root, Trusted Publisher
- Classifies every certificate: **Expired** (red), **Critical** < 30 days (red), **Warning** < 90 days (yellow), **Healthy** (green)
- **SSL/TLS remote check** — connects to any `hostname` or `hostname:port` via TCP + SslStream and reads the presented certificate
- `-Targets` accepts a comma-separated list of hosts or a path to a text file (one host per line)
- Dark-themed HTML report with summary cards (Total / Expired / Critical / Warning / Healthy) and full cert inventory tables
- Console summary shows only non-Healthy certs; HTML report includes the complete inventory
- `-Unattended` runs all local stores + SSL checks silently and exports the report

---

### T.A.L.O.N.

Sweeps the standard Windows persistence surfaces and inventories every entry so a technician can answer "what runs on this machine without me asking it to?" Every entry is enriched with signature status (Microsoft-signed / 3rd-party signed / unsigned / tampered) and target-on-disk check (present / missing / n-a). Does not attempt to judge malice — visibility is the job.

- **Surfaces covered**:
  - Run / RunOnce keys (HKCU, HKLM, WOW6432Node)
  - Startup folders (per-user + All Users), with `.lnk` target resolution
  - Non-Microsoft services with auto-start
  - Scheduled Tasks outside the `\Microsoft\` subtree, one row per action
  - WMI event subscriptions (`__EventFilter`, `__EventConsumer`, `__FilterToConsumerBinding` under `root\subscription`) — high-signal persistence surface
  - Image File Execution Options (IFEO) `Debugger` hijacks
  - Winlogon `Shell`, `Userinit`, and `AppInit_DLLs` hijack points
- Per-entry enrichment: resolves the binary path (handles quoted and unquoted command lines with environment-variable expansion), runs `Get-AuthenticodeSignature`, and flags Microsoft-signed entries distinctly from third-party signed
- Dark-themed HTML report with six summary cards (total / MS-signed / 3rd-party / unsigned / missing-targets) and one detail table per surface
- Nav bar auto-generated from the categories found so large machines paginate naturally
- Auto-elevates; read-only audit with no state changes

---

### T.O.T.E.M.

TPM health audit gating the four questions that matter for modern Windows management: "Can this machine run Windows 11?" "Is BitLocker going to unlock cleanly?" "Is this machine eligible for Autopilot attestation?" "Is the TPM provisioned and owned?"

- **TPM status** via `Get-Tpm`: present / enabled / activated / ready / owned, manufacturer ID and text, manufacturer version, physical-presence interface version, auto-provisioning state, restart-pending flag
- **Specification summary**: parses the raw `SpecVersion` string, labels as TPM 1.2 / TPM 2.0, flags Windows-11-readiness
- **BitLocker dependency** via `Get-BitLockerVolume`: for each volume, lists every key protector and flags whether any is TPM-based (`Tpm`, `TpmPin`, `TpmPinStartupKey`, `TpmStartupKey`) so a technician can answer "if this TPM goes bad, which drives stop unlocking?"
- **Attestation / Endorsement Key** via `Get-TpmEndorsementKeyInfo`: surfaces EK presence and manufacturer-certificate count — required for Windows Autopilot pre-provisioning and attested boot
- **Red / yellow / green verdict** with specific remediation hints (enable in firmware, provision, clear-and-reprovision, vendor BIOS update for dTPM->fTPM, etc.)
- Dark-themed HTML report with six summary cards including readiness, spec, BitLocker-volumes-depending-on-TPM, and EK presence
- Auto-elevates; read-only audit

---

### G.R.I.F.F.I.N.

Antivirus and Microsoft Defender health audit. Answers the four questions that decide whether a Windows endpoint is actually protected: "Is real-time protection on?" "Are signatures fresh?" "Are there unresolved threats hiding in history?" "Is the policy surface (exclusions, ASR rules, cloud submission) configured to block, audit, or just shrug?"

- **Defender core state** via `Get-MpComputerStatus`: antivirus / antispyware / AM service enabled, AM running mode (Normal / Passive / EDR Block Mode), engine and product versions, real-time / behavior / IOAV / on-access / NIS toggles, tamper protection
- **Cloud and sample posture** via `Get-MpPreference`: MAPS reporting tier, sample submission consent, cloud block level, cloud extended timeout, PUA protection mode
- **Signatures**: AV / antispyware / NIS signature versions, last-updated timestamps, age in days; configurable yellow / red thresholds via `-SignatureMaxAgeDays`
- **Scan history**: last quick scan and last full scan with start / end / age; flags machines that have never had a full scan
- **Threat history** via `Get-MpThreat`: every threat the engine has ever logged with severity badge, active vs resolved state, detection count, affected resources; unresolved high / severe entries drive the verdict
- **Recent detections** via `Get-MpThreatDetection`: last 50 detections with process name, user, resource, cleanup-success flag
- **Exclusions inventory**: path, extension, process, and IP exclusions surfaced as separate tables so a technician can spot over-broad scope
- **Attack Surface Reduction rules**: every configured ASR GUID resolved to its friendly name, with mode (Block / Audit / Warn / Not Configured); audit-only rules drive a yellow finding
- **Third-party AV products** via the `root\SecurityCenter2` namespace: each registered AV with its packed `productState` decoded into real-time and up-to-date flags; flags concurrent third-party real-time alongside Defender
- **Service health** for `WinDefend`, `WdNisSvc`, `Sense`, `WdFilter`, `SecurityHealthService`, and `mpssvc`; critical services not running drive a red finding
- **Recent Defender events** from `Microsoft-Windows-Windows Defender/Operational` (configurable lookback via `-EventDays`)
- **Platform protection**: virtualization-based security, memory integrity (HVCI), Credential Guard (Enterprise / Education / Server editions), LSA protection (`RunAsPPL`, confirmed against the Wininit boot event), the vulnerable driver blocklist, and Smart App Control. Scored with a verdict of its own — Hardened / Partial / Not hardened — that does not change the AV verdict, so older hardware without HVCI still reads correctly on the AV side
- **Red / yellow / green verdict** with explicit remediation hints. `Write-TKError` telemetry fires on `Get-MpComputerStatus` failure and on each unresolved high / severe threat
- Dark-themed HTML report with seven summary cards (posture, real-time, tamper, signature age, threats, AM mode, platform protection)
- Auto-elevates; read-only audit

---

## Network & Remote

### L.E.Y.L.I.N.E.

Tests and diagnoses network connectivity at every layer with one-click remediation options.

- Displays all network adapters with status, IPv4 address, and MAC
- Ping tests: default gateway, Google DNS (8.8.8.8), Cloudflare (1.1.1.1), and DNS resolution
- Color-coded latency indicators (green < 50ms, yellow < 150ms, red ≥ 150ms)
- DNS server listing per adapter
- TCP port test — enter any host:port to check reachability
- Traceroute to any destination
- **Remediation**: flush DNS cache, DHCP release & renew, full network stack reset (Winsock + TCP/IP + firewall)

---

### E.M.I.S.S.A.R.Y.

Connects to a remote Windows machine via WinRM and runs Technician Toolkit scripts without needing physical access.

- Enter target hostname or IP; supports current credentials (domain/Kerberos) or manual entry
- WinRM connectivity test with step-by-step enable instructions if unreachable
- **Run A.U.S.P.E.X.** — copies script to remote, executes, retrieves HTML report locally
- **Run W.A.R.D.** — copies script to remote, executes, retrieves HTML report locally
- **Run W.H.E.T.S.T.O.N.E.** — installs Windows Updates on target (reboot warning shown)
- **Run B.A.S.I.L.I.S.K.** — applies full security baseline on target, retrieves CSV log
- **Interactive session** — opens a full `Enter-PSSession` shell on the target
- All output files retrieved to `EMISSARY_<MachineName>\` in the script directory
- Remote staging folder cleaned up automatically after each operation
- Target machine prerequisite: `Enable-PSRemoting -Force` (run as Administrator)

---

### L.A.N.T.E.R.N.

Discovers all live hosts on the local /24 subnet and produces a network asset inventory.

- Parallel ICMP ping sweep across all 254 host addresses
- DNS reverse lookup for hostname resolution on each live host
- MAC address retrieval from the ARP neighbor table
- Optional TCP port scan against common service ports (21, 22, 23, 80, 443, 445, 3389, 5985, 8080, 8443)
- Color-coded port badges in the HTML report (open vs. closed)
- Summary cards: total hosts discovered, ports scanned, unreachable
- CSV export of the full host inventory
- Dark-themed HTML report saved to the script directory
- `-Unattended -Action Sweep` for silent sweep and report export

---

### W.I.S.P.

Wi-Fi profile audit. Answers "what wireless networks does this machine know, and which of them silently auto-connect to surfaces an attacker could spoof?"

- **Profile inventory** via `netsh wlan export profile ... key=clear` to a per-run temp folder, parsed from the locale-stable WLAN profile XML schema (avoids the brittle locale-dependent text output of `netsh wlan show profile`)
- **Per profile**: SSID, hidden flag, connection type (ESS / IBSS), connection mode (auto / manual), autoSwitch, authentication (Open / WEP / WPA / WPA2-PSK / WPA2-Enterprise / WPA3-SAE / OWE / etc.), encryption (none / WEP / TKIP / AES / GCMP), 802.1X usage, MAC randomization, key material (masked by default; cleartext only when `-IncludeKey` is supplied)
- **Adapter inventory** via `Get-NetAdapter -Physical` filtered to `Native 802.11`: name, status, link speed, MAC, driver version and date
- **Filtered tables**: open / weak profiles (broken WEP, deprecated TKIP, open-auth) and auto-connecting profiles (with risk tier badges)
- **Red / yellow / green verdict**: red on open auto-connect, WEP, deprecated cipher; yellow on TKIP, autoSwitch, hidden+auto, MAC randomization off, > 25 saved profiles
- Privacy posture banner in the HTML report makes it explicit whether keys were rendered in cleartext
- Dark-themed HTML report with six summary cards
- Auto-elevates; read-only audit; temp folder is wiped after parsing

---

### P.O.R.T.A.L.

VPN and Always-On VPN audit. Answers "what tunnels can leave this machine, are they configured to leak credentials, and is split-tunnel routing covered by NRPT?"

- **Built-in Windows VPNs** via `Get-VpnConnection` (user scope) and `Get-VpnConnection -AllUserConnection` (all-user / Always-On candidate scope): name, server, tunnel type (PPTP / L2TP / SSTP / IKEv2 / Automatic), authentication methods (PAP / CHAP / MS-CHAPv2 / EAP / MachineCertificate), encryption level (NoEncryption / Optional / Required / Maximum), split-tunnel flag, connection state, profile type
- **Always-On VPN app triggers** via `Get-VpnConnectionTrigger`: which applications and DNS suffixes auto-launch each VPN
- **Name Resolution Policy Table** via `Get-DnsClientNrptPolicy`: namespace -> DNS server mappings, IPsec-required flag, DirectAccess servers
- **Active VPN tunnel interfaces** via `Get-NetIPInterface` filtered to interfaces matching `VPN`/`Wintun`/`WireGuard`/`Tailscale`/`GlobalProtect`/`AnyConnect`/`PPP`
- **Third-party VPN clients** via service-table probing: Cisco AnyConnect / Cisco Secure Client (`vpnagent`, `csc_vpnagent`), Palo Alto GlobalProtect (`PanGPS`, `PanGPA`), Ivanti / Pulse (`JuniperNetworksTunnelService`, `PulseService`), OpenVPN (interactive, legacy, GUI), WireGuard (manager + per-config tunnels), Tailscale, ZeroTier, Cloudflare WARP, NordVPN, Proton VPN, F5 BIG-IP Edge Client
- **Red / yellow / green verdict**: red on PAP authentication or NoEncryption (Write-TKError telemetry fires for both -- credentials in cleartext / traffic in cleartext are immediate findings); yellow on CHAP, Optional encryption, split-tunnel without matching NRPT, multiple competing third-party VPN vendors
- Special-case `NONE CONFIGURED` posture statement when zero built-in VPNs and zero third-party clients are detected -- not a finding, just an inventory result
- Dark-themed HTML report with six summary cards and six detail sections
- Auto-elevates; read-only audit

---

### L.O.D.E.S.T.A.R.

Diagnoses and repairs "The trust relationship between this workstation and the primary domain failed" — and the quieter faults that come before it.

- **Domain controller discovery** (`nltest /dsgetdc`) and the DC locator SRV record in DNS
- **DNS servers** on each active adapter, flagging public resolvers on a domain member — the most common reason a machine intermittently cannot find its domain
- **DC reachability** on DNS, Kerberos, RPC, LDAP and SMB
- **Clock offset** from the DC (`w32tm /stripchart`): past five minutes Kerberos refuses tickets; past one minute it is flagged as drift
- **Secure channel** via `nltest /sc_verify`, with the Netlogon status decoded into *connectivity* (fix the network) versus *trust* (the machine password or computer account is wrong). Uses `nltest` rather than `Test-ComputerSecureChannel`, which does not exist in PowerShell 7
- Netlogon policy that disables machine password rotation
- `-Action Audit` (default) is read-only. `-Action Repair` resyncs time from the domain hierarchy, resets the secure channel (`nltest /sc_reset`), and — only when the DC still rejects the machine password, and only interactively — resets the computer account password with domain credentials the technician enters (`Reset-ComputerMachinePassword`, run through Windows PowerShell when hosted in PowerShell 7). It never changes DNS and never unjoins or rejoins; a deleted computer account is reported with the rejoin steps
- `-WhatIf` previews every repair

> **LODESTAR vs LEYLINE:** LEYLINE diagnoses general network reachability. LODESTAR diagnoses the machine's membership of its domain — the DC, the secure channel, Kerberos time — and repairs the trust.

---

## Cloud & Identity

### Z.E.N.I.T.H.

Connects to an Azure subscription and generates a comprehensive HTML assessment report for the environment.

- Auto-installs required `Az.*` modules if missing
- Security posture: NSG inbound exposure, publicly accessible storage accounts, SQL firewall rules, HTTPS enforcement on web apps
- Access & governance: RBAC role assignments, resource locks, Azure Policy compliance
- Backup coverage: Recovery Services Vaults, protected items, storage redundancy
- VM inventory: OS, size, region, power state, NIC and disk details
- SQL hygiene: database tier, size, backup retention, geo-redundancy
- Orphaned resources: unattached disks, unused public IPs, empty NICs
- Tag coverage: resources and resource groups missing tags
- Azure Advisor alerts and Defender for Cloud secure score
- Prioritized remediation recommendations section
- Parameters: `-SubscriptionId` to target a specific subscription; `-OutputPath` to set report destination; `-NoOpen` to suppress auto-open

---

### A.L.M.A.N.A.C.

Connects to Microsoft 365 via the Microsoft Graph API and audits the tenant's license and mailbox state.

- Auto-installs required `Microsoft.Graph` modules if missing
- License assignment audit: per-user SKU name, assigned licenses, consumed vs. purchased units
- Unlicensed user identification — active accounts with no M365 license assigned
- Inactive user report — accounts with no sign-in activity in 90+ days
- MFA registration status per user (registered / not registered)
- Shared mailbox audit: display name, primary SMTP address, size, last activity
- Dark-themed HTML report combining all sections with summary cards
- `-Unattended` to auto-connect and export the full report without prompts

---

### O.R.B.I.T.

Connects to Microsoft Graph and audits the Intune-managed device estate.

- Auto-installs required `Microsoft.Graph` modules if missing
- Device inventory: managed devices grouped by OS, ownership, compliance state, join type, last sync
- Compliance state summary: compliant, non-compliant, in grace period, error, unknown
- Stale device detection with 30 / 60 / 90 day buckets so silent devices surface before they lose management
- Configuration profile inventory with assignment coverage — flags profiles that have no assignments
- Dark-themed HTML report with summary cards and nav anchors for each section
- OrgName from `config.json` is prepended to the report subtitle when configured
- Telemetry via `Write-TKError` for Graph authentication failure and Intune query failure
- `-Unattended` to auto-connect and export the full report without prompts

---

### E.C.L.I.P.S.E.

Connects to Microsoft Graph and audits Entra ID identity hygiene — the security-and-cost questions that ALMANAC (licensing) and ORBIT (devices) don't cover.

- Guest users: every guest in the tenant with creation date, last sign-in, and invite state; flags guests inactive 90+ days for external-access cleanup
- Privileged role holders: every member of every active directory role, one row per user-per-role; high-tier roles (Global Admin, Privileged Role Admin, User Admin, Exchange / SharePoint / Security / Conditional Access Admin) called out in the console summary
- Password never expires: cloud-only members with `DisablePasswordExpiration` set — the classic shared-mailbox / service-account hygiene miss
- Stale privileged users: deduplicated set of admins from the role audit who have not signed in in 60+ days, annotated with all their role assignments
- Disabled but licensed: disabled accounts still consuming paid SKUs (cost-leak after off-boarding)
- Dark-themed HTML report with OrgName prefix from `config.json` and six summary cards, one per audit section
- Telemetry via `Write-TKError` on Graph auth failure and each Graph query failure
- `-Unattended` to auto-connect and export the full report without prompts

---

### A.S.T.E.R.I.S.M.

Connects to Microsoft Graph and audits the M365 Teams estate. Complements ALMANAC (licensing), ORBIT (devices), and ECLIPSE (identity) with the collaboration layer — orphan teams, external-access blind spots, and governance stragglers that survive every tenant cleanup pass.

- Enumerates every team-backed M365 group (`resourceProvisioningOptions/Any(x:x eq 'Team')`) with owner list, member list, guest count (filtered on `userType eq 'Guest'`), visibility, sensitivity labels, creation and last-renewed dates
- **Orphan teams**: zero owners OR every owner disabled — the #1 governance issue in long-running tenants
- **Public teams**: visibility = `Public`, meaning any tenant user can join without approval; flagged regardless of content
- **Teams with guest members**: sorted by guest count descending so external-access blast radius is front-of-page
- **Large teams**: ≥ 250 members by default (editable constant); governance candidates (should these be channels of a smaller team, or split?)
- **Stale teams**: M365-group `RenewedDateTime` older than 365 days (editable) — indicates the owner has not confirmed ongoing use if a group-expiration policy is set
- Dark-themed HTML report with OrgName prefix, six summary cards, and six per-category tables
- Telemetry via `Write-TKError` on auth and group-query failures
- `-Unattended` auto-connect + export

---

### C.U.M.U.L.U.S.

Connects to Microsoft Graph and inventories the SharePoint Online estate in a single bulk read. Uses the `getSharePointSiteUsageDetail(period='D30')` report endpoint rather than per-site queries, so the audit completes in one round trip regardless of tenant size.

- **Tenant sharing policy**: `sharingCapability`, `sharingDomainRestrictionMode`, `defaultSharingLinkType`, `defaultLinkPermission` — one card showing the tenant-wide posture that governs every site below
- **Site inventory**: every site with URL, template, owner, storage (Wh-formatted), file count, last activity date, external-sharing flag
- **Large sites**: storage ≥ 100 GB (editable constant)
- **External-sharing sites**: sites with `External Sharing = True` in the usage report
- **Ownerless sites**: no owner display name — typical after owner off-boarding without transfer
- **Stale sites**: no activity in 180+ days (editable)
- Dark-themed HTML report with OrgName prefix, six summary cards, and six detail tables (sharing policy, full inventory, plus one per derived finding)
- Telemetry via `Write-TKError` on auth failure and usage-report fetch failure
- `-Unattended` auto-connect + export

---

### R.A.V.E.N.

Exchange Online mailbox security audit — looks for the signs and preconditions of business email compromise, where an attacker signs in, quietly forwards or hides mail, and waits. Read-only.

- **Mailbox forwarding** (`ForwardingSmtpAddress` / `ForwardingAddress`) classified external or internal against the tenant's accepted domains, noting whether a copy is kept
- **Inbox rules** in every user and shared mailbox, flagged when they forward or redirect externally, move mail into rarely opened folders (RSS Feeds, Conversation History, Archive…) and mark it read, delete messages about payments or security, or carry a throwaway name like `.` — each flagged rule lists why
- **Tenant settings**: outbound spam policies that allow automatic external forwarding, transport rules that redirect or copy mail outside the organisation, SMTP AUTH enabled org-wide or per mailbox, mailbox auditing disabled
- **Delegation**: Full Access and Send As grants, for access review (`-SkipDelegation` to skip the per-mailbox permission sweep on large tenants)
- **Email authentication** for each domain: SPF (missing, multiple, `+all`, `?all`), DMARC (missing, `p=none`, `pct` below 100), and DKIM signing state
- **DNS-only mode**: `-DnsOnly -Domain contoso.com` runs the SPF / DKIM / DMARC checks with no sign-in and no module — useful before a tenant is taken on
- Verdict: At Risk / Review / Clean
- Requires `ExchangeOnlineManagement` (offered for install if missing) and a role that can read recipient and transport configuration — Global Reader or View-Only Organization Management is enough

> **RAVEN vs ALMANAC:** ALMANAC audits *licensing and MFA registration* through Microsoft Graph. RAVEN audits *what the mailboxes are doing* through Exchange Online — forwarding, rules, and the settings that let a compromise go unnoticed.

---

### H.A.L.O.

Entra ID Conditional Access posture audit. Scores the tenant's policies against the baseline every tenant should have, then reviews the policies themselves. Read-only.

- **Baseline coverage**: MFA for all users on all apps, MFA for Global Administrator, legacy authentication blocked — plus, where licensed, sign-in / user risk policies and a device-compliance requirement. Only *enabled* policies count, and "MFA **or** compliant device" does not count as MFA, because a compliant device gets in without it. Security defaults are recognised when no Conditional Access policy exists
- **Privileged roles**: each of 14 admin roles (Global, Privileged Role, Security, Conditional Access, Exchange, SharePoint, User, Intune…) checked for an enforcing MFA policy that does not exclude it
- **Hygiene**: policies left in report-only or disabled, enforcing all-users policies with long exclusion lists, and policies that reference deleted users or groups
- **Emergency access**: the users and groups excluded from *every* enforcing all-users policy. None at all is flagged, because one bad policy or an MFA outage can then lock every administrator out
- **Named locations**: trusted IP ranges wider than /16 (IPv4) or /32 (IPv6)
- Needs only `Microsoft.Graph.Authentication` (offered for install if missing) and the `Policy.Read.All` + `Directory.Read.All` delegated scopes — all reads go through `Invoke-MgGraphRequest`
- Verdict: Exposed / Gaps / Enforced

> **HALO vs ECLIPSE:** ECLIPSE audits the *identities* — guests, privileged role holders, stale admins. HALO audits the *policies* that decide how those identities may sign in.

### C.A.R.I.L.L.O.N.

Teams Phone call queue and auto attendant audit. Answers "who is in this queue?" by name — never by object ID — and where every call can end up. Read-only.

- **Call queues**: number, routing method, presence-based routing, conference mode, alert time, and the overflow / timeout / no-agent actions with their targets resolved to names ("Forward -> User Jane Doe", "Auto attendant Main Line")
- **Agents**: every agent on every queue with display name, sign-in name, opt-in state, whether they are voice-enabled, whether the account is enabled, and how they got there — directly, through a named group, or through a Teams channel
- **Auto attendants**: number, language, time zone, operator, the business-hours menu key by key, and each after-hours / holiday call flow with its schedule
- **Resource accounts**: number, type, and the queue or attendant each one fronts
- **Routing map**: which attendants and queues send calls to each queue or attendant, so an unreachable one stands out
- **Flags**: queues with no agents, every agent opted out, or a single agent opted in; agents not voice-enabled or disabled; targets pointing at deleted users or groups; an overflow threshold of 0; queues that hang up on overflow / timeout; attendants with no after-hours flow; unreachable queues and attendants; resource accounts assigned to nothing
- Needs `MicrosoftTeams` (offered for install if missing). Group names and "which group added this agent" come from Microsoft Graph (`Microsoft.Graph.Authentication`, `Directory.Read.All`); `-SkipGraph` avoids that second sign-in and still names users through the Teams module
- `-Name 'Sales*'` limits the output to matching queues and attendants
- Writes `CARILLON_<timestamp>.html` and `CARILLON_Agents_<timestamp>.csv` (one row per queue / agent pairing)
- Verdict: Broken / Attention / Healthy

### C.H.A.L.I.C.E.

Microsoft 365 Apps client diagnosis and repair, on the machine, for the signed-in user — "Outlook keeps asking for my password", "Word says Unlicensed Product", "Teams won't sign in". The tenant-side tools (ALMANAC, RAVEN, HALO) look at the service; CHALICE looks at the client.

- **Install**: Click-to-Run version and platform, products with their **support dates** (Office 2016 / 2019 ended 2025-10-14, Office 2021 ends 2026-10-13), update channel with any policy override, whether updates are enabled and their source reachable, and when the Office files last changed (90+ days = updates are not landing). MSI-based Office is flagged
- **Activation**: shared computer activation, and whether this user holds a Microsoft 365 license token
- **Sign-in**: the accounts Office has cached (work vs personal, how many tenants), `EnableADAL = 0` and `DisableAADWAM` / `DisableADALatopWAMOverride` (classic password-loop causes, local or by policy), the AAD token broker package, cached Office credentials, and the device's Entra join state and Primary Refresh Token from `dsregcmd /status`
- **Teams**: new vs classic (retired) client, and the cache size
- **Actions** (`-Action`), each previewed by `-WhatIf`:
  - `ResetSignIn` — Microsoft's documented activation reset: backs up and clears Office's cached identities and licensing key, deletes the license tokens and cached Office credentials, and re-registers a missing token broker. The user signs in to Office again
  - `ResetTeams` — closes Teams and clears the new (and classic) Teams cache
  - `QuickRepair` — Office's offline Quick Repair (requests elevation)
  - `Update` — starts a Click-to-Run update now
- **Runs as the signed-in user, not elevated**: Office's accounts, licenses and caches are per user, and an elevated run under another account would read and reset the wrong profile. A run whose account differs from the console user is flagged
- Verdict: Broken / Attention / Healthy

---

## Data & Migration

### R.E.V.E.N.A.N.T.

Migrates user profile data from a source machine or profile to a destination using Robocopy for reliable folder transfers.

- Select source from local profiles or enter a custom/UNC path
- Select destination profile or custom path
- Choose individual items or migrate all at once
- Migrates: Desktop, Documents, Downloads, Pictures, Videos, Music
- Migrates: Outlook profiles & data files, email signatures
- Migrates: Chrome bookmarks, Edge bookmarks, Firefox profiles
- OneDrive for Business detection with Known Folder Move awareness
- Restore from an E.M.B.A.L.M. ZIP as the source
- `-WhatIf` previews what would be copied (with file count and size) without performing any transfers
- Generates a timestamped CSV migration log in the script directory

---

### E.M.B.A.L.M.

Creates a compressed ZIP backup of a selected user profile before a machine is reimaged or wiped.

- Select profile from detected local user profiles
- Choose individual items or archive all at once
- Archives: Desktop, Documents, Downloads, Pictures, Videos, Music
- Archives: Outlook data, email signatures, Chrome/Edge/Firefox bookmarks
- Destination can be a local path or UNC network share
- Stages files to `%TEMP%` via Robocopy before compressing
- Creates ZIP using .NET `System.IO.Compression.ZipFile` (no 2 GB file limit)
- Writes a plain-text manifest inside the ZIP listing every archived item
- Cleans up staging folder automatically on completion
- Generates a timestamped CSV log in the script directory

---

### P.H.Y.L.A.C.T.E.R.Y.

Pre-migration validator that answers the single question a technician cares about before a laptop swap or reimage: "Is this user's data actually going to be in the cloud when we hand them a new machine?"

- OneDrive client state: installed path, file version, process running
- Signed-in accounts: lists every Business / Personal account in `HKCU:\Software\Microsoft\OneDrive\Accounts` with email, sync-root path, and tenant ID
- Known Folder Move: Desktop, Documents, and Pictures — for each, resolves the `User Shell Folders` registry entry and flags whether it redirects into OneDrive or sits locally
- Content volume: file count and total size per known folder with `Format-Bytes` colour coding (>5 GB yellow, >25 GB red) so a technician can anticipate upload time
- Sync errors: surfaces OneDrive-related Application log events from the last 7 days
- Readiness verdict: red / yellow / green based on client state + account + KFM coverage, with specific issues and warnings enumerated
- Dark-themed HTML report with OrgName prefix from `config.json` and a summary-card row including the verdict
- `-Unattended` silent run that exports the HTML without opening the browser

---

### E.X.H.U.M.E.

Outlook data-file discovery that inventories every PST (and optionally OST) on the machine before a mail migration, flags the files that will be a problem, and cross-references them against configured Outlook profiles.

- Profile walk: iterates `HKCU:\Software\Microsoft\Office\{14,15,16}.0\Outlook\Profiles` and extracts every mounted store path per profile, marks the default profile
- Drive-wide file scan: recursively enumerates every local fixed drive for `*.pst` (and `*.ost` with `-IncludeOst`), skipping system folders to keep the scan fast
- Orphan detection: any PST on disk that no configured profile references is flagged — often lost archives in old user-data folders after a domain migration
- Oversize detection: PSTs ≥ 50 GB are flagged red as Exchange Online Import Service blockers; 10–50 GB flagged yellow as slow-import warnings
- Stale detection: PSTs not accessed in 365+ days are flagged as archive-on-ingest candidates rather than primary-mailbox imports
- Readiness verdict: red / yellow / green based on the worst finding, with specific issues and recommendations enumerated
- Dark-themed HTML report with OrgName prefix from `config.json`, six summary cards (verdict, file count, total size, orphans, oversize, stale), and the full file inventory with per-file colour-coded size badges
- `-ScanDrives C:,D:` to narrow the scan; `-IncludeOst` to add `.ost` caches (off by default since OSTs regenerate on the next machine); `-Unattended` for silent export

---

## Requirements

**Running the application:** Windows 10 1809 or later, x64 or ARM64. Nothing else
— PowerShell, the shared module, and every tool ship inside the executable. The
per-tool requirements below still apply to what each tool *does* (RSAT for the AD
tools, a BitLocker-capable edition for WYRM, and so on), but nothing in the
first four rows is needed.

**Running the scripts:** everything below.

| Requirement | Notes |
|-------------|-------|
| Windows PowerShell 5.1+ | All scripts |
| **TechnicianToolkit.psm1 in the same folder** | All scripts — shared module providing logging, HTML, and privilege helpers |
| Administrator privileges | All scripts (auto-elevation on `grimoire.ps1`) |
| Internet connectivity | All scripts |
| Windows Package Manager (winget) | `conjure.ps1` (Chocolatey supported as alternative) |
| PSWindowsUpdate module | `whetstone.ps1`, `forge.ps1` (auto-installed if missing) |
| *(none — built-in cmdlets only)* | `conduit.ps1`, `solder.ps1`, `necropsy.ps1`, `torpor.ps1`, `lodestar.ps1`, `minotaur.ps1`, `chalice.ps1` (runs as the signed-in user, not elevated) |
| Entra ID account with device join permissions | `covenant.ps1` |
| Robocopy (built into Windows) | `revenant.ps1`, `embalm.ps1` |
| BitLocker-capable Windows edition (Pro/Enterprise) | `wyrm.ps1` |
| WinRM enabled on target machine | `emissary.ps1`, `vigil.ps1` (remote mode) |
| RSAT ActiveDirectory module | `sphinx.ps1`, `argus.ps1` (auto-installed if missing) |
| Az PowerShell modules | `zenith.ps1`, `orrery.ps1` (optional, auto-installed if -IncludeAzureRbac) |
| Microsoft.Graph modules | `almanac.ps1`, `orbit.ps1`, `eclipse.ps1`, `asterism.ps1`, `cumulus.ps1`, `orrery.ps1` (auto-installed if missing) |
| ExchangeOnlineManagement module | `raven.ps1` (offered for install if missing; not needed for `-DnsOnly`), `orrery.ps1` (optional, auto-installed if -IncludeExchange) |
| PnP.PowerShell module | `orrery.ps1` (optional, auto-installed if -IncludeSharePoint) |
| Azure subscription + appropriate RBAC | `zenith.ps1` |
| Microsoft 365 tenant + Global Reader or equivalent | `almanac.ps1`, `orbit.ps1`, `eclipse.ps1`, `asterism.ps1`, `cumulus.ps1`, `orrery.ps1`, `raven.ps1` |
| Microsoft.Graph.Authentication module + Policy.Read.All / Directory.Read.All scopes | `halo.ps1` (offered for install if missing) |
| MicrosoftTeams module + Teams Administrator (or Global Reader) | `carillon.ps1` (offered for install if missing; Microsoft.Graph.Authentication + Directory.Read.All also used to name groups unless `-SkipGraph`) |
| Microsoft Intune licence + DeviceManagement Graph permissions | `orbit.ps1`, `orrery.ps1` |
| RoleManagement.Read.Directory + AuditLog.Read.All Graph scopes | `eclipse.ps1`, `orrery.ps1` |
| On-premises Active Directory domain membership | `sphinx.ps1`, `argus.ps1`, `lodestar.ps1` |

---

## Installation

### The application

```powershell
winget install CursedTechnocrat.TechnicianToolkit
```

Or download `TechnicianToolkit.exe` for your architecture from
[Releases](https://github.com/CursedTechnocrat/TechnicianToolkit/releases) and
run it. There is no installer and nothing to uninstall — it writes its working
files into a `TechnicianToolkit` folder beside itself, or under
`%LOCALAPPDATA%` when the medium is read-only, and deleting the `.exe` removes
it. Copying it to a USB stick is a supported way to deploy it.

Verify the SHA-256 against the release notes before running it — see the note on
[unsigned binaries](#the-application) above.

### The scripts

1. Clone or download this repository
2. Extract **all files** (`.ps1` and `TechnicianToolkit.psm1`) into the same folder — the module must be co-located with the scripts
3. Open PowerShell as Administrator
4. Navigate to the toolkit directory

```powershell
cd C:\Path\To\Toolkit
```

Any single tool also bootstraps itself: drop one `.ps1` onto a machine and it
fetches the shared module from GitHub on first run. See [Quick Launch](#quick-launch)
for one-liners that do exactly that.

---

## Quick Launch

Run any script directly from GitHub without cloning — scripts download into whatever directory your shell is currently in. `cd` to your working folder first, then paste the command.

> **Note:** All scripts depend on `TechnicianToolkit.psm1`. The module is downloaded automatically by the GRIMOIRE command below. If running an individual script without GRIMOIRE, download the module first — see the first command in the block.

```powershell
# Example: navigate to your working folder first
cd C:\Technicians\JobSite42\

# Now run the quick launch — grimoire.ps1 (and any tools it downloads) will land here
```

```powershell
# TechnicianToolkit.psm1 — Shared module (download once per working folder; required by all scripts)
Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1 -OutFile "$(Get-Location)\TechnicianToolkit.psm1"

# G.R.I.M.O.I.R.E. — Hub launcher (recommended starting point; downloads module automatically)
Set-ExecutionPolicy Bypass -Scope Process -Force; $d="$(Get-Location)"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1 -OutFile "$d\TechnicianToolkit.psm1"; $f="$d\grimoire.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/grimoire.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Deployment & Onboarding ──────────────────────────────────────────────────

# C.O.V.E.N.A.N.T. — Machine onboarding & Entra ID domain join
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\covenant.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/covenant.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.O.N.J.U.R.E. — Software deployment via winget or Chocolatey
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\conjure.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/conjure.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# R.U.N.E.P.R.E.S.S. — Printer driver installation and configuration
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\runepress.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/runepress.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# F.O.R.G.E. — Driver detection & installation
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\forge.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/forge.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# W.H.E.T.S.T.O.N.E. — Windows Update management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\whetstone.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/whetstone.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# H.E.A.R.T.H. — Toolkit setup wizard
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\hearth.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/hearth.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# R.I.T.U.A.L. — Workflow orchestrator (runs recipes of other tools)
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\ritual.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/ritual.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.O.N.D.U.I.T. — Windows Update connectivity diagnosis & repair
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\conduit.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/conduit.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.O.L.D.E.R. — Windows servicing (component store & SFC) repair
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\solder.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/solder.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Diagnostics & Reporting ──────────────────────────────────────────────────

# A.U.S.P.E.X. — System diagnostics and HTML report
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\auspex.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/auspex.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# W.A.R.D. — Local user account audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\ward.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/ward.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# F.A.T.H.O.M. — Disk & storage health monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\fathom.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/fathom.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# V.I.G.I.L. — Service, task & event log monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\vigil.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/vigil.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.U.G.U.R. — Physical disk health & SMART status
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\augur.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/augur.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.L.E.A.N.S.E. — Disk cleanup (temp, update cache, browser caches, Recycle Bin)
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\cleanse.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/cleanse.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.C.R.Y.E.R. — Unified diagnostic HTML report
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\scryer.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/scryer.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.N.V.I.L. — BIOS / UEFI / firmware audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\anvil.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/anvil.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# H.O.U.R.G.L.A.S.S. — Laptop battery health audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\hourglass.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/hourglass.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.O.D.E.X. — Toolkit report index (rolls up existing HTML reports)
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\codex.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/codex.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# N.E.C.R.O.P.S.Y. — Crash & unexpected-reboot analysis
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\necropsy.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/necropsy.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# T.O.R.P.O.R. — Slow-machine triage
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\torpor.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/torpor.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Security ─────────────────────────────────────────────────────────────────

# W.Y.R.M. — BitLocker encryption management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\wyrm.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/wyrm.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# B.A.S.I.L.I.S.K. — Security baseline enforcement
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\basilisk.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/basilisk.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.P.H.I.N.X. — Active Directory management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\sphinx.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/sphinx.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.R.G.U.S. — AD account roster & access levels
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\argus.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/argus.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# M.I.N.O.T.A.U.R. — File share & NTFS permissions review
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\minotaur.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/minotaur.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# P.H.O.E.N.I.X. — Certificate health monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\phoenix.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/phoenix.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# T.A.L.O.N. — Persistence / autoruns audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\talon.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/talon.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# T.O.T.E.M. — TPM health audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\totem.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/totem.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# G.R.I.F.F.I.N. — AV / Defender health audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\griffin.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/griffin.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Network & Remote ─────────────────────────────────────────────────────────

# L.E.Y.L.I.N.E. — Network diagnostics & remediation
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\leyline.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/leyline.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# E.M.I.S.S.A.R.Y. — Remote execution via WinRM
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\emissary.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/emissary.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# L.A.N.T.E.R.N. — Network discovery & asset inventory
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\lantern.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/lantern.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# W.I.S.P. — Wi-Fi profile audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\wisp.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/wisp.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# P.O.R.T.A.L. — VPN / Always-On VPN audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\portal.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/portal.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# L.O.D.E.S.T.A.R. — Domain trust & secure channel diagnosis and repair
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\lodestar.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/lodestar.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Cloud & Identity ─────────────────────────────────────────────────────────

# Z.E.N.I.T.H. — Azure environment assessment
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\zenith.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/zenith.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.L.M.A.N.A.C. — Microsoft 365 license & mailbox audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\almanac.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/almanac.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# O.R.B.I.T. — Intune / MDM compliance audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\orbit.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/orbit.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# E.C.L.I.P.S.E. — Entra ID identity hygiene audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\eclipse.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/eclipse.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.S.T.E.R.I.S.M. — Microsoft Teams audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\asterism.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/asterism.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.U.M.U.L.U.S. — SharePoint Online audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\cumulus.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/cumulus.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# O.R.R.E.R.Y. — Entra ID group dependency audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\orrery.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/orrery.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# R.A.V.E.N. — Exchange Online mailbox security audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\raven.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/raven.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# H.A.L.O. — Entra ID Conditional Access posture
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\halo.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/halo.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.A.R.I.L.L.O.N. — Teams Phone call queue & auto attendant audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\carillon.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/carillon.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.H.A.L.I.C.E. — Microsoft 365 Apps client diagnosis & repair (run as the affected user)
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\chalice.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/chalice.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Data & Migration ─────────────────────────────────────────────────────────

# R.E.V.E.N.A.N.T. — Profile migration
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\revenant.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/revenant.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# E.M.B.A.L.M. — Pre-reimaging profile backup
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\embalm.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/embalm.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# P.H.Y.L.A.C.T.E.R.Y. — OneDrive Known-Folder-Move pre-migration validator
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\phylactery.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/phylactery.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# E.X.H.U.M.E. — Outlook PST / OST discovery
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\exhume.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/exhume.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f
```

> All scripts require an Administrator PowerShell session. The `-Scope Process` flag limits the execution policy bypass to the current session only — it does not permanently change system policy.

---

## Usage

### Recommended: Launch via GRIMOIRE (hub)

```powershell
.\grimoire.ps1
```

Select a tool by number. Control returns to the menu when the tool finishes.

### Or run tools directly

```powershell
# Deployment & Onboarding
.\covenant.ps1      # New machine onboarding and Entra ID domain join
.\conjure.ps1       # Software deployment via winget or Chocolatey
.\runepress.ps1     # Printer driver installation and configuration
.\forge.ps1         # Driver detection and installation
.\whetstone.ps1   # Windows Update management
.\hearth.ps1        # Toolkit setup wizard
.\ritual.ps1        # Workflow orchestrator — runs sequences of other tools
.\conduit.ps1       # Windows Update connectivity diagnosis and repair
.\solder.ps1        # Windows servicing repair — component store, SFC, update failure codes

# Diagnostics & Reporting
.\auspex.ps1        # System diagnostics and HTML health report
.\ward.ps1          # User account audit and HTML report
.\fathom.ps1     # Disk space monitor — volume usage, low-space alerts, cleanup
.\vigil.ps1      # Service, task, and event log monitor
.\augur.ps1         # Physical disk health — SMART status, wear prediction, failure forecast
.\cleanse.ps1       # Disk cleanup — temp files, update cache, browser caches, Recycle Bin
.\scryer.ps1        # Unified diagnostic report — system, users, disks, SMART, services in one HTML
.\anvil.ps1         # BIOS / UEFI / firmware audit and HTML report
.\hourglass.ps1          # Laptop battery health audit
.\codex.ps1         # Toolkit report index — rolls up existing HTML reports into one bound exhibit
.\necropsy.ps1      # Crash & unexpected-reboot analysis and HTML report
.\torpor.ps1        # Slow-machine triage — what is using the CPU, memory and disk right now

# Security
.\wyrm.ps1        # BitLocker drive encryption management
.\basilisk.ps1         # Security baseline enforcement
.\sphinx.ps1       # Active Directory user and group management
.\argus.ps1        # Active Directory account roster and access-level report
.\minotaur.ps1      # File share & NTFS permissions review, HTML + CSV
.\phoenix.ps1         # Certificate health and SSL expiry monitor
.\talon.ps1          # Persistence / autoruns audit
.\totem.ps1          # TPM health audit
.\griffin.ps1        # AV / Microsoft Defender health audit

# Network & Remote
.\leyline.ps1       # Network diagnostics and remediation
.\emissary.ps1       # Remote execution via WinRM
.\lantern.ps1       # Network discovery and asset inventory
.\wisp.ps1        # Wi-Fi profile audit
.\portal.ps1        # VPN / Always-On VPN audit
.\lodestar.ps1          # Domain trust & secure channel diagnosis and repair

# Cloud & Identity
.\zenith.ps1      # Azure environment assessment and HTML report
.\almanac.ps1     # Microsoft 365 license and mailbox audit
.\orbit.ps1         # Intune / MDM compliance audit and HTML report
.\eclipse.ps1        # Entra ID identity hygiene audit and HTML report
.\asterism.ps1      # Microsoft Teams audit and HTML report
.\cumulus.ps1         # SharePoint Online audit and HTML report
.\orrery.ps1       # Entra ID group dependency audit and HTML report
.\raven.ps1         # Exchange Online mailbox security audit and HTML report
.\halo.ps1       # Entra ID Conditional Access posture audit
.\carillon.ps1      # Teams Phone call queues & auto attendants — agents by name
.\chalice.ps1       # Microsoft 365 Apps client — activation, sign-in loops, Teams cache, Quick Repair

# Data & Migration
.\revenant.ps1       # Profile migration and data transfer
.\embalm.ps1        # Pre-reimaging profile backup to ZIP
.\phylactery.ps1         # OneDrive Known-Folder-Move pre-migration validator
.\exhume.ps1         # Outlook PST / OST discovery before mail migration
```

All scripts must be run as Administrator.

---

## Configuration

The toolkit uses an optional `config.json` file in the toolkit directory. All scripts function without it — it only pre-fills common values to reduce prompts. Use **H.E.A.R.T.H.** (`hearth.ps1`) to configure settings interactively.

| Key | Description |
|-----|-------------|
| `OrgName` | Organization name shown in HTML report headers |
| `LogDirectory` | Directory where HTML reports and transcripts are saved |
| `TeamsWebhook` | Incoming webhook URL for Teams error notifications (used by `Write-TKError`) |
| `Archive.DefaultDestination` | Default backup path for EMBALM |
| `Revenant.DefaultDestination` | Default migration destination for REVENANT |
| `Covenant.DefaultTimezone` | Default Windows timezone ID for COVENANT |
| `Covenant.DefaultLocalAdminUser` | Default local administrator account name for COVENANT |

| Script | Configurable Variables |
|--------|------------------------|
| **grimoire.ps1** | None — tool list is defined in the `$Tools` array in the script |
| **covenant.ps1** | `config.json` — `Covenant.DefaultTimezone`, `Covenant.DefaultLocalAdminUser` |
| **conjure.ps1** | `$RequiredCatalog` / `$OptionalCatalog` — single-source package catalog (each entry has `Name` / `Winget` / `Choco`); `$AdobeReader` / `$AdobePro` — Adobe edition entries; `$PackageManager` — default manager (`winget` or `choco`); `$InstallExitInfo` — winget/installer exit-code → human-readable reason + class (`Success`/`Failed`) map used by `Resolve-InstallExit`; `config.json` — `Conjure.DirectDownloads` (array of `Name` / `Url` / `Args` / `Sha256`); `-DirectUrl` for one-off direct downloads; `-WhatIf` for dry run |
| **runepress.ps1** | `-DriverPath` — folder or file to install the driver from (defaults to the script directory); `$ExtractRoot` — driver extraction staging folder (defaults to `.\ExtractedDrivers`); `-WhatIf` for dry run |
| **forge.ps1** | None — driver sources scanned from current folder automatically; `-WhatIf` previews Windows Update and local driver installs |
| **whetstone.ps1** | None — power settings are detected and restored automatically; `-WhatIf` lists available updates without installing |
| **hearth.ps1** | None — all settings entered via the interactive wizard; `config.json` is the output (see config key table above) |
| **ritual.ps1** | `-Recipe {Onboard\|Retire\|HealthCheck\|SecuritySweep\|NetworkSweep\|TenantSweep}` — named recipe to run; `-RecipeFile <path.psd1>` — custom recipe file; `-ContinueOnError` — tolerate per-step failures |
| **conduit.ps1** | `-Action {Audit\|Repair\|ResetCache}` — Audit is read-only (default), Repair applies the safe fixes, ResetCache also rebuilds SoftwareDistribution / catroot2; `-Force` — remove the WSUS pointer even when the server is reachable, the device is domain-joined, or local Group Policy sets it; `-WhatIf` — preview every change without applying it |
| **solder.ps1** | `-Action {Audit\|Repair}` — Audit is read-only (default), Repair runs DISM /RestoreHealth then sfc /scannow; `-Deep` — full DISM /ScanHealth in the audit; `-Source <WIM:path:index>` — repair from install media (`/LimitAccess`); `-Cleanup` — also run DISM /StartComponentCleanup; `-WhatIf` — preview every repair |
| **auspex.ps1** | `$ReportOutputPath` — folder where the HTML report is saved (defaults to script directory; accepts any local or UNC path) |
| **ward.ps1** | None — audit runs automatically; stale threshold is 90 days (editable in script); LAPS rotation is flagged overdue 3 days past the policy's `PasswordAgeDays` (`$LapsRotationGraceDays`) |
| **fathom.ps1** | None — thresholds are Warning < 15% free, Critical < 5% free (editable in script); old profile threshold is 90 days |
| **vigil.ps1** | None — critical service list editable in script; `-Target` accepts any WinRM-reachable hostname |
| **augur.ps1** | None — scans all physical disks automatically; `-Unattended` for silent HTML export |
| **cleanse.ps1** | None — categories selected interactively or all cleaned with `-Unattended`; `-WhatIf` for dry run |
| **scryer.ps1** | `-OutputPath` — directory to write `SCRYER_Report_<timestamp>.html` (defaults to configured log directory) |
| **anvil.ps1** | None — system identity, UEFI state, vendor channels, and Windows Update pending firmware are all auto-detected |
| **hourglass.ps1** | None — ROOT\WMI battery classes and Win32_Battery are queried unconditionally; thresholds (80/60 pct, 300/500 cycles) are editable constants in the script |
| **codex.ps1** | `LogDirectory` (read) — defines which directory CODEX scans for existing HTML reports; CLI overrides via `-LogDir`. Optional `-DaysBack <int>` filter and pattern-strict file matching are constants in the script |
| **necropsy.ps1** | `-Days <int>` — look-back window (default 30, range 1-365); read-only otherwise |
| **torpor.ps1** | `-SampleSeconds <int>` — sample window (default 30, range 5-300); thresholds are constants at the top of the script |
| **wyrm.ps1** | `LogDirectory` (read) — Export action writes the HTML report there unless `-OutputPath` overrides it. `OrgName` is shown in the report header. Drive and action are selected interactively, or with `-Drive` / `-Action` |
| **basilisk.ps1** | None — categories selected interactively; screensaver timeout editable in script (default 600 s) |
| **sphinx.ps1** | None — user search and action selected interactively; stale threshold is 90 days (editable in script) |
| **argus.ps1** | `LogDirectory` (read) — HTML and CSV are written there unless `-OutputPath` overrides it; `OrgName` is shown in the report header. `-StaleDays <int>` inactivity threshold (default 90), `-SearchBase <dn>` to scope to one OU, `-Server <dc>` to target a domain controller, `-IncludeDisabled`, `-AdminGroupPattern <regex>` for customer-created admin groups (default `(?i)(admin\|operator\|helpdesk\|privileg)`), `-SkipCustomGroupScan`, `-NoCsv` |
| **minotaur.ps1** | `-Path <folder>` — review one folder tree instead of the shares; `-Depth <0-10>` — folder levels to walk into each share (default 2); the 5,000-folder cap and the expected-broad-read shares (`NETLOGON`, `SYSVOL`) are constants in the script |
| **phoenix.ps1** | None — stores and targets selected interactively or via `-Targets` parameter |
| **talon.ps1** | None — every persistence surface is enumerated unconditionally |
| **totem.ps1** | None — reads TPM state, BitLocker protectors, and endorsement key info unconditionally |
| **griffin.ps1** | `-EventDays <int>` — Defender event-log lookback window (default 7, range 1-90); `-SignatureMaxAgeDays <int>` — yellow / red threshold for signature age (default 7, doubles for the red tier); read-only otherwise |
| **leyline.ps1** | None — all tests run interactively; no persistent config |
| **emissary.ps1** | None — target, credentials, and operation selected interactively at runtime |
| **lantern.ps1** | `$script:ScanPorts` — list of TCP ports checked during scan (editable in script) |
| **wisp.ps1** | `-IncludeKey` — opt-in switch to render WLAN profile pre-shared keys in cleartext (default: masked); read-only otherwise |
| **portal.ps1** | None — enumerates VPN connections, NRPT, tunnel interfaces, and third-party VPN client services unconditionally; the third-party client catalog is an editable `$ThirdPartyClients` array in the script |
| **lodestar.ps1** | `-Action {Audit\|Repair}` — Audit is read-only (default); Repair resyncs time, resets the secure channel, and (interactive only) resets the machine password when the DC rejects it; `-WhatIf` — preview every repair |
| **zenith.ps1** | `-SubscriptionId` — target a specific Azure subscription; `-OutputPath` — HTML report destination; `-NoOpen` — suppress auto-open after export |
| **almanac.ps1** | None — tenant and report scope selected interactively at runtime |
| **orbit.ps1** | None — tenant, device scope, and report scope selected interactively at runtime |
| **eclipse.ps1** | None — tenant and audit scope selected interactively at runtime |
| **asterism.ps1** | None — tenant selected interactively; large-team and stale thresholds (250 members, 365 days) are editable constants in the script |
| **cumulus.ps1** | None — tenant selected interactively; large-site and stale thresholds (100 GB, 180 days) are editable constants in the script |
| **orrery.ps1** | `-GroupName` / `-GroupId` — target group (one required, prompted otherwise); `-IncludeSharePoint` + `-SharePointAdminUrl` — scan tenant SP sites via PnP (capped by `-SharePointSiteLimit`, default 200); `-IncludeExchange` — scan EXO transport rules, delegations, role groups, DL nesting; `-IncludeAzureRbac` — scan every visible subscription via Az; `-OutputPath` — HTML report destination; `-NoOpen` — suppress auto-open |
| **raven.ps1** | `-DnsOnly` + `-Domain <name[,name]>` — SPF / DKIM / DMARC only, no sign-in; `-Domain` in a full audit limits the DNS checks to those domains (default: every accepted domain except `*.onmicrosoft.com`); `-SkipDelegation` — skip the per-mailbox Full Access / Send As sweep |
| **halo.ps1** | None — tenant chosen at sign-in; the exclusion threshold (5) and the privileged-role table are constants in the script |
| **carillon.ps1** | `-Name <wildcard>` — only queues / attendants whose name matches; `-SkipGraph` — Teams sign-in only (group names left unresolved) |
| **chalice.ps1** | `-Action {Audit\|ResetSignIn\|ResetTeams\|QuickRepair\|Update}` — Audit is read-only (default); ResetSignIn clears Office's accounts, license tokens and credentials; ResetTeams clears the Teams cache; QuickRepair runs Office's offline repair; Update starts a Click-to-Run update; `-WhatIf` — preview every change |
| **revenant.ps1** | `config.json` — `Revenant.DefaultDestination`; source, items, and destination also selectable interactively |
| **embalm.ps1** | `config.json` — `Archive.DefaultDestination`; profile, items, and destination also selectable interactively |
| **phylactery.ps1** | None — reads HKCU OneDrive and User Shell Folders registry and enumerates Desktop / Documents / Pictures for the currently logged-on user |
| **exhume.ps1** | `-ScanDrives` — comma-separated drive list to restrict the scan (defaults to all fixed local drives); `-IncludeOst` — include `.ost` caches in the scan (off by default) |

---

## Logging

All HTML reports and transcripts are saved to the configured `LogDirectory` from `config.json` if set, otherwise to the script's own directory.

| Script | Log Output |
|--------|------------|
| **grimoire.ps1** | No log file — hub activity is visible on-screen only |
| **covenant.ps1** | Console — action summary printed at completion |
| **conjure.ps1** | Console — per-package status table printed at completion |
| **runepress.ps1** | Script directory — `RUNEPRESS_InstallLog_<timestamp>.csv` |
| **forge.ps1** | Script directory — `FORGE_DriverReport_<timestamp>.csv` |
| **whetstone.ps1** | `WHETSTONE_<timestamp>.log` — PowerShell transcript of the full session. Default path is `%TEMP%`; with `-Transcript` it is written to the configured log directory instead. |
| **hearth.ps1** | Console only — settings persisted to `config.json` |
| **ritual.ps1** | Log directory — `RITUAL_<timestamp>.html` (rollup report with per-step status, duration, and links to each child report) |
| **conduit.ps1** | Log directory — `CONDUIT_<timestamp>.html` (findings, policy values, endpoint reachability, service state, and every action taken). `Repair` also writes `CONDUIT_WUPolicy_<timestamp>.reg`, a backup of the WindowsUpdate policy key, to the same directory. |
| **solder.ps1** | Log directory — `SOLDER_<timestamp>.html` (store health, pending restarts, SFC result, decoded update failures, every action taken). DISM and SFC keep their own logs in `%WINDIR%\Logs\DISM\dism.log` and `%WINDIR%\Logs\CBS\CBS.log` |
| **auspex.ps1** | Log directory — `AUSPEX_<timestamp>.html` (dark-themed HTML report) |
| **ward.ps1** | Log directory — `WARD_<timestamp>.html` (dark-themed HTML report) |
| **fathom.ps1** | Log directory — `FATHOM_<timestamp>.html` (dark-themed HTML report) |
| **vigil.ps1** | Log directory — `VIGIL_<timestamp>.html` (dark-themed HTML health report) |
| **augur.ps1** | Log directory — `AUGUR_<timestamp>.html` (dark-themed HTML report) |
| **cleanse.ps1** | Console only — cleanup summary printed at completion; no log file |
| **scryer.ps1** | `-OutputPath` (defaults to log directory) — `SCRYER_Report_<timestamp>.html` (unified diagnostic report) |
| **anvil.ps1** | Log directory — `ANVIL_<timestamp>.html` (BIOS / UEFI / firmware audit report) |
| **hourglass.ps1** | Log directory — `HOURGLASS_<timestamp>.html` (laptop battery health audit), `HOURGLASS_battery_report_<timestamp>.xml` (parsed `powercfg` data), `HOURGLASS_battery_report_<timestamp>.html` (full Microsoft `powercfg /batteryreport` HTML) |
| **necropsy.ps1** | Log directory — `NECROPSY_<timestamp>.html` (crash & unexpected-reboot analysis) |
| **torpor.ps1** | Log directory — `TORPOR_<timestamp>.html` (slow-machine triage) |
| **codex.ps1** | Log directory — `CODEX_<timestamp>.html` (rollup index of every other report in the log directory; CODEX excludes its own outputs from the index) |
| **wyrm.ps1** | Console only by default; the Export action writes `WYRM_Report_<timestamp>.html` (status + recovery keys) to `-OutputPath` or the log directory |
| **basilisk.ps1** | Log directory — `BASILISK_BaselineLog_<timestamp>.csv` |
| **sphinx.ps1** | Log directory — `SPHINX_Stale_<timestamp>.html`; `SPHINX_PwdExpiry_<timestamp>.html` |
| **argus.ps1** | Log directory (or `-OutputPath`) — `ARGUS_<timestamp>.html` (authentication policy + account roster & access levels), `ARGUS_Roster_<timestamp>.csv` (same roster with blank Action / Notes columns for customer review) |
| **minotaur.ps1** | Log directory — `MINOTAUR_<timestamp>.html` (share permissions review) and `MINOTAUR_<timestamp>.csv` (every access-control entry scanned) |
| **phoenix.ps1** | Log directory — `PHOENIX_<timestamp>.html` (cert inventory & SSL results) |
| **talon.ps1** | Log directory — `TALON_<timestamp>.html` (persistence / autoruns audit) |
| **totem.ps1** | Log directory — `TOTEM_<timestamp>.html` (TPM health audit) |
| **griffin.ps1** | Log directory — `GRIFFIN_<timestamp>.html` (AV / Defender health audit) |
| **leyline.ps1** | Console only — no log file |
| **emissary.ps1** | Script directory — `EMISSARY_<MachineName>\` folder containing retrieved output files |
| **lantern.ps1** | Log directory — `LANTERN_<timestamp>.html` and `LANTERN_<timestamp>.csv` |
| **wisp.ps1** | Log directory — `WISP_<timestamp>.html` (Wi-Fi profile audit). Per-profile XMLs are exported to a temp folder and deleted after parsing |
| **portal.ps1** | Log directory — `PORTAL_<timestamp>.html` (VPN / Always-On VPN audit) |
| **lodestar.ps1** | Log directory — `LODESTAR_<timestamp>.html` (findings, secure channel, DNS, DC ports, clock offset, and every repair action) |
| **zenith.ps1** | `-OutputPath` (default `%TEMP%`) — `azure-assessment-<timestamp>.html`; auto-opens in browser |
| **almanac.ps1** | Log directory — `ALMANAC_<timestamp>.html` (combined license & mailbox report) |
| **orbit.ps1** | Log directory — `ORBIT_<timestamp>.html` (Intune / MDM compliance report) |
| **eclipse.ps1** | Log directory — `ECLIPSE_<timestamp>.html` (Entra ID identity hygiene report) |
| **asterism.ps1** | Log directory — `ASTERISM_<timestamp>.html` (Teams estate audit report) |
| **cumulus.ps1** | Log directory — `CUMULUS_<timestamp>.html` (SharePoint Online estate audit) |
| **orrery.ps1** | Log directory — `ORRERY_<GroupName>_<timestamp>.html` (Entra ID group dependency audit) |
| **raven.ps1** | Log directory — `RAVEN_<timestamp>.html` (Exchange Online mailbox security audit) |
| **halo.ps1** | Log directory — `HALO_<timestamp>.html` (Conditional Access posture) |
| **carillon.ps1** | Log directory — `CARILLON_<timestamp>.html` (call queues & auto attendants) and `CARILLON_Agents_<timestamp>.csv` (queue / agent roster) |
| **chalice.ps1** | Log directory — `CHALICE_<timestamp>.html` (install, activation, sign-in, Teams, every action taken). `ResetSignIn` also writes `CHALICE_Identities_<timestamp>.reg` and `CHALICE_Licensing_<timestamp>.reg`, backups of the registry keys it clears |
| **revenant.ps1** | Log directory — `REVENANT_MigrationLog_<timestamp>.csv` |
| **embalm.ps1** | Script directory — `EMBALM_Log_<timestamp>.csv`; manifest inside ZIP |
| **phylactery.ps1** | Log directory — `PHYLACTERY_<timestamp>.html` (OneDrive KFM readiness report) |
| **exhume.ps1** | Log directory — `EXHUME_<timestamp>.html` (Outlook PST / OST discovery report) |

---

## Contributing

This toolkit is built for working technicians, and it gets better when the people using it in the field push their fixes back. Bug reports, new tools, and corrections from real tickets are all welcome — see [CONTRIBUTING.md](CONTRIBUTING.md) for the full guide.

Please ensure all additions maintain:

- Consistent formatting and naming conventions
- The GPL notice header block at the top of every new file (copy it from any existing script)
- The standard `<# .SYNOPSIS / .DESCRIPTION / .USAGE / .NOTES #>` header block
- Comprehensive error handling
- Detailed logging and user feedback
- Administrator privilege checks

Contributions are accepted under the same license as the project, **GPL-3.0-or-later**. You keep the copyright in what you write; you are simply licensing it under the project's terms so it can ship with everything else.

---

## Disclaimer

These scripts modify system settings and may install software, updates, or change domain membership in ways that require a reboot. Save all work before running. Use at your own risk.

---

## License

**GNU General Public License v3.0 or later (GPL-3.0-or-later)** — see [LICENSE](LICENSE) for the full text.

This toolkit is free software. You may use it, study it, change it, and pass it on. The one condition is that it stays free: if you distribute a modified version, your recipients get the same source and the same rights you had.

What that means day to day:

| You want to… | GPL says |
|---|---|
| Run these tools on client machines, at any scale, commercially | **Go ahead.** Running the software is unrestricted — the GPL only attaches obligations when you *distribute* it. |
| Modify a script for your own shop's workflow and keep it in-house | **Go ahead.** Internal use is not distribution. No obligation to publish anything. |
| Share your modified version with other technicians, or ship it to clients as a tool | Fine — include the source and license it GPL-3.0-or-later too. |
| Fold these scripts into a closed-source commercial RMM product | Not permitted. That is exactly what the copyleft is here to prevent. |

Every script carries its own copyright and license notice in its header, so a single `.ps1` copied onto a technician's USB stick still tells the next person what it is and where it came from.

Copyright © 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors.

Prior releases were published under the MIT License. That grant is not revoked — anyone who obtained a copy under MIT keeps those terms for that copy. Everything from this point forward is GPL-3.0-or-later.
