# Technician Toolkit

> A PowerShell-based toolkit for IT technicians to automate common system administration tasks — forged in the arcane arts of automation.

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
| Need tools with no LiveConnect counterpart (COVENANT, REVENANT, CIPHER, ARCHIVE, SHADE, RUNEPRESS, LEYLINE, FORGE, TALISMAN, CITADEL, LANTERN, THRESHOLD, AUGUR, CLEANSE, RELIQUARY, GARGOYLE, ARTIFACT, HEARTH) | **This repo** — these tools are interactive by nature or require auth flows incompatible with LiveConnect |

---

## Table of Contents

- [Tools Overview](#tools-overview)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Launch](#quick-launch)
- [Usage](#usage)
- [Configuration](#configuration)
- [Logging](#logging)
- [Contributing](#contributing)
- [Disclaimer](#disclaimer)

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
| 5 | **restoration.ps1** | **R.E.S.T.O.R.A.T.I.O.N.** — Renews Every System Through Orderly Rite — Automating The Installation Of New updates | Automated Windows Update management and maintenance |
| 6 | **hearth.ps1** | **H.E.A.R.T.H.** — Hub for Environment, Admin Runtime & Toolkit Hardening | Toolkit setup wizard — configure org name, log paths, and default values |

### Diagnostics & Reporting

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 10 | **auspex.ps1** | **A.U.S.P.E.X.** — Audits, Uncovers, Surveys Performance, Events & eXceptions | System diagnostics, health assessment, and HTML report generation |
| 11 | **ward.ps1** | **W.A.R.D.** — Watches Accounts, Reviews Roles & Detects anomalies | Local user account audit with role, last logon, flags, and HTML report |
| 12 | **threshold.ps1** | **T.H.R.E.S.H.O.L.D.** — Tests Hardware Reliability, Evaluates Storage Health, & Optimizes/Logs Disk data | Disk space monitor — volume usage, low-space alerts, temp cleanup, old profile detection, HTML report |
| 13 | **gargoyle.ps1** | **G.A.R.G.O.Y.L.E.** — Guards Against Runtime Glitches On Your Log Events | Service, task & event log monitor — health check local or remote machine, HTML report |
| 14 | **augur.ps1** | **A.U.G.U.R.** — Analyzes, Uncovers & Gauges Unit Reliability | Physical disk health — SMART status, wear prediction, failure forecast, hardware reliability, HTML report |
| 15 | **cleanse.ps1** | **C.L.E.A.N.S.E.** — Cleans Leftover, Ephemeral And Neglected System Entries | Disk cleanup — user & system temp, Windows Update cache, browser caches, Recycle Bin |

### Security

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 20 | **cipher.ps1** | **C.I.P.H.E.R.** — Configures & Implements Policy-based Hardware Encryption & Recovery | BitLocker drive encryption management — enable, disable, key backup |
| 21 | **sigil.ps1** | **S.I.G.I.L.** — Secures Infrastructure: Governs via Integrated Lockdown | Security baseline enforcement — telemetry, UAC, firewall, audit policy, password policy |
| 22 | **citadel.ps1** | **C.I.T.A.D.E.L.** — Centralizes Identity, Tasks, Accounts, Directories, Entitlements & Logons | Active Directory user & group management — unlock, reset, lockout forensics, stale & expiry reports |
| 23 | **artifact.ps1** | **A.R.T.I.F.A.C.T.** — Audits, Reports Trust, Identity, Fingerprints, Authority, Certificates & TLS | Certificate health monitor — local cert stores, SSL/TLS expiry, HTML report |

### Network & Remote

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 30 | **leyline.ps1** | **L.E.Y.L.I.N.E.** — Locates, Examines & Yields Latency, Infrastructure, Network & Endpoints | Network diagnostics & remediation — adapters, ping, DNS, port tests, IP renew, stack reset |
| 31 | **shade.ps1** | **S.H.A.D.E.** — Summons Hosts for Administrative Deployment & Execution | Remote machine execution via WinRM — run toolkit tools without physical access |
| 32 | **lantern.ps1** | **L.A.N.T.E.R.N.** — Locates & Audits Network Topology, Enumerating Resources & Nodes | Network discovery — subnet ping sweep, DNS lookup, MAC addresses, port scan, HTML report |

### Cloud & Identity

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 40 | **talisman.ps1** | **T.A.L.I.S.M.A.N.** — Tenant Assessment, Logging, Infrastructure, Security, Monitoring & Access Navigator | Azure subscription assessment — security posture, RBAC, backup coverage, Advisor alerts, HTML report |
| 41 | **reliquary.ps1** | **R.E.L.I.Q.U.A.R.Y.** — Reports, Evaluates Licenses, Inventories, Quotas, Users, Access & Registration Yields | Microsoft 365 license & mailbox audit — license assignments, MFA status, shared mailboxes, HTML report |

### Data & Migration

| # | Script | Acronym | Purpose |
|---|--------|---------|---------|
| 50 | **revenant.ps1** | **R.E.V.E.N.A.N.T.** — Relocates, Extracts, Validates Environments, Networks, Accounts 'N Transfers | Profile migration and data transfer between machines or profiles |
| 51 | **archive.ps1** | **A.R.C.H.I.V.E.** — Automated Repository Compressing & Housing Important Volume Exports | Pre-reimaging profile backup — ZIP to local path or network share |

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
- Upgrade-all mode for keeping existing packages current
- Tracks and displays installation status per package

**Default required packages:** Microsoft Teams, Microsoft 365, 7-Zip, Google Chrome, Adobe Acrobat Reader, Zoom

**Default optional packages:** Zoom Outlook Plugin, Mozilla Firefox, Dell Command Update

---

### R.U.N.E.P.R.E.S.S.

Automates printer driver extraction, installation, and network printer configuration via a command-line interface.

- Supports ZIP, EXE, and MSI driver formats
- Handles automatic driver extraction and INF-based installation via pnputil
- Configures network printers via IP (TCP/IP port) or UNC path post-install
- Generates a timestamped installation log (CSV) in the script directory

---

### F.O.R.G.E.

Audits the device tree for driver problems and automates driver installation from multiple sources.

- Scans all devices for errors with human-readable error descriptions (missing, corrupted, cannot start, etc.)
- Checks Windows Update for available driver updates via PSWindowsUpdate (auto-installed if missing)
- Installs drivers from the current folder: ZIP (extracts and runs pnputil on INF), bare INF, EXE (silent), MSI (quiet)
- Exports a full driver inventory CSV with device name, driver version, date, and manufacturer
- Cleans up extracted driver staging folders automatically

---

### R.E.S.T.O.R.A.T.I.O.N.

Automates Windows Update detection, installation, and reboot handling with minimal user intervention.

- Disables sleep and display timeout for the duration of the run; restores settings on exit
- Ensures NuGet provider and PSWindowsUpdate module are installed and current
- Installs available updates (drivers excluded) with no forced reboot
- Checks reboot status and prompts only when required
- 30-second reboot countdown with Escape key cancel

---

### H.E.A.R.T.H.

Interactive setup wizard for the Technician Toolkit — configure all settings without hand-editing JSON.

- Step-by-step wizard covers all seven configuration fields with descriptions, hints, and live validation
- Configures: organization name, log/report directory, Teams webhook URL, ARCHIVE default destination, REVENANT default destination, COVENANT default timezone and local admin username
- Path fields validate on entry — prompts to create missing directories automatically
- View current configuration with color-coded status: green = configured, yellow = empty or path not found
- Edit individual fields without re-running the full wizard
- **Environment checks**: PowerShell version, admin status, module presence, winget, Chocolatey, RSAT, Microsoft.Graph, Az, and log directory write access
- Configuration reset with `YES` confirmation guard
- All settings persisted to `config.json` in the toolkit directory
- `-Unattended` displays current config and runs environment checks silently

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
- Console summary with highlighted flagged accounts
- HTML report with color-coded badges and summary cards
- Report saved to script directory as `WARD_<timestamp>.html`

---

### T.H.R.E.S.H.O.L.D.

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

### G.A.R.G.O.Y.L.E.

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

> **AUGUR vs THRESHOLD:** AUGUR answers "is this drive about to fail?" (SMART/hardware).
> THRESHOLD answers "is this drive running out of space?" (volume usage/cleanup).

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

## Security

### C.I.P.H.E.R.

Manages BitLocker drive encryption across all volumes with an interactive menu-driven interface.

- Displays current encryption status for all drives on launch
- Enable BitLocker with TPM, TPM + PIN, or recovery password only
- Recovery key displayed and confirmed before encryption begins
- Disable BitLocker (full decryption) with confirmation prompt
- Back up recovery key to Active Directory or Entra ID (Azure AD)
- View recovery key ID and password for any encrypted drive
- Suspend BitLocker for BIOS/firmware updates (auto-resumes after reboot)
- Resume suspended BitLocker protection

---

### S.I.G.I.L.

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
- Changes logged to `SIGIL_BaselineLog_<timestamp>.csv` in the script directory

---

### C.I.T.A.D.E.L.

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

### A.R.T.I.F.A.C.T.

Monitors certificate health across the local machine and remote hosts — surfaces expiring and expired certificates before they cause outages.

- Audits local Windows certificate stores: Personal (My), Intermediate CA, Trusted Root, Trusted Publisher
- Classifies every certificate: **Expired** (red), **Critical** < 30 days (red), **Warning** < 90 days (yellow), **Healthy** (green)
- **SSL/TLS remote check** — connects to any `hostname` or `hostname:port` via TCP + SslStream and reads the presented certificate
- `-Targets` accepts a comma-separated list of hosts or a path to a text file (one host per line)
- Dark-themed HTML report with summary cards (Total / Expired / Critical / Warning / Healthy) and full cert inventory tables
- Console summary shows only non-Healthy certs; HTML report includes the complete inventory
- `-Unattended` runs all local stores + SSL checks silently and exports the report

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

### S.H.A.D.E.

Connects to a remote Windows machine via WinRM and runs Technician Toolkit scripts without needing physical access.

- Enter target hostname or IP; supports current credentials (domain/Kerberos) or manual entry
- WinRM connectivity test with step-by-step enable instructions if unreachable
- **Run A.U.S.P.E.X.** — copies script to remote, executes, retrieves HTML report locally
- **Run W.A.R.D.** — copies script to remote, executes, retrieves HTML report locally
- **Run R.E.S.T.O.R.A.T.I.O.N.** — installs Windows Updates on target (reboot warning shown)
- **Run S.I.G.I.L.** — applies full security baseline on target, retrieves CSV log
- **Interactive session** — opens a full `Enter-PSSession` shell on the target
- All output files retrieved to `SPECTER_<MachineName>\` in the script directory
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

## Cloud & Identity

### T.A.L.I.S.M.A.N.

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

### R.E.L.I.Q.U.A.R.Y.

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
- Restore from an A.R.C.H.I.V.E. ZIP as the source
- `-WhatIf` previews what would be copied (with file count and size) without performing any transfers
- Generates a timestamped CSV migration log in the script directory

---

### A.R.C.H.I.V.E.

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

## Requirements

| Requirement | Notes |
|-------------|-------|
| Windows PowerShell 5.1+ | All scripts |
| **TechnicianToolkit.psm1 in the same folder** | All scripts — shared module providing logging, HTML, and privilege helpers |
| Administrator privileges | All scripts (auto-elevation on `grimoire.ps1`) |
| Internet connectivity | All scripts |
| Windows Package Manager (winget) | `conjure.ps1` (Chocolatey supported as alternative) |
| PSWindowsUpdate module | `restoration.ps1`, `forge.ps1` (auto-installed if missing) |
| Entra ID account with device join permissions | `covenant.ps1` |
| Robocopy (built into Windows) | `revenant.ps1`, `archive.ps1` |
| BitLocker-capable Windows edition (Pro/Enterprise) | `cipher.ps1` |
| WinRM enabled on target machine | `shade.ps1`, `gargoyle.ps1` (remote mode) |
| RSAT ActiveDirectory module | `citadel.ps1` (auto-installed if missing) |
| Az PowerShell modules | `talisman.ps1` (auto-installed if missing) |
| Microsoft.Graph modules | `reliquary.ps1` (auto-installed if missing) |
| Azure subscription + appropriate RBAC | `talisman.ps1` |
| Microsoft 365 tenant + Global Reader or equivalent | `reliquary.ps1` |
| On-premises Active Directory domain membership | `citadel.ps1` |

---

## Installation

1. Clone or download this repository
2. Extract **all files** (`.ps1` and `TechnicianToolkit.psm1`) into the same folder — the module must be co-located with the scripts
3. Open PowerShell as Administrator
4. Navigate to the toolkit directory

```powershell
cd C:\Path\To\Toolkit
```

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

# R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\restoration.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/restoration.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# H.E.A.R.T.H. — Toolkit setup wizard
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\hearth.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/hearth.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Diagnostics & Reporting ──────────────────────────────────────────────────

# A.U.S.P.E.X. — System diagnostics and HTML report
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\auspex.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/auspex.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# W.A.R.D. — Local user account audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\ward.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/ward.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# T.H.R.E.S.H.O.L.D. — Disk & storage health monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\threshold.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/threshold.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# G.A.R.G.O.Y.L.E. — Service, task & event log monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\gargoyle.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/gargoyle.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.U.G.U.R. — Physical disk health & SMART status
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\augur.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/augur.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.L.E.A.N.S.E. — Disk cleanup (temp, update cache, browser caches, Recycle Bin)
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\cleanse.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/cleanse.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Security ─────────────────────────────────────────────────────────────────

# C.I.P.H.E.R. — BitLocker encryption management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\cipher.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/cipher.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.I.G.I.L. — Security baseline enforcement
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\sigil.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/sigil.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.I.T.A.D.E.L. — Active Directory management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\citadel.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/citadel.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.R.T.I.F.A.C.T. — Certificate health monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\artifact.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/artifact.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Network & Remote ─────────────────────────────────────────────────────────

# L.E.Y.L.I.N.E. — Network diagnostics & remediation
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\leyline.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/leyline.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.H.A.D.E. — Remote execution via WinRM
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\shade.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/shade.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# L.A.N.T.E.R.N. — Network discovery & asset inventory
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\lantern.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/lantern.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Cloud & Identity ─────────────────────────────────────────────────────────

# T.A.L.I.S.M.A.N. — Azure environment assessment
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\talisman.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/talisman.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# R.E.L.I.Q.U.A.R.Y. — Microsoft 365 license & mailbox audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\reliquary.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/reliquary.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# ── Data & Migration ─────────────────────────────────────────────────────────

# R.E.V.E.N.A.N.T. — Profile migration
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\revenant.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/revenant.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.R.C.H.I.V.E. — Pre-reimaging profile backup
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\archive.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/archive.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f
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
.\restoration.ps1   # Windows Update management
.\hearth.ps1        # Toolkit setup wizard

# Diagnostics & Reporting
.\auspex.ps1        # System diagnostics and HTML health report
.\ward.ps1          # User account audit and HTML report
.\threshold.ps1     # Disk space monitor — volume usage, low-space alerts, cleanup
.\gargoyle.ps1      # Service, task, and event log monitor
.\augur.ps1         # Physical disk health — SMART status, wear prediction, failure forecast
.\cleanse.ps1       # Disk cleanup — temp files, update cache, browser caches, Recycle Bin

# Security
.\cipher.ps1        # BitLocker drive encryption management
.\sigil.ps1         # Security baseline enforcement
.\citadel.ps1       # Active Directory user and group management
.\artifact.ps1         # Certificate health and SSL expiry monitor

# Network & Remote
.\leyline.ps1       # Network diagnostics and remediation
.\shade.ps1       # Remote execution via WinRM
.\lantern.ps1       # Network discovery and asset inventory

# Cloud & Identity
.\talisman.ps1         # Azure environment assessment and HTML report
.\reliquary.ps1         # Microsoft 365 license and mailbox audit

# Data & Migration
.\revenant.ps1       # Profile migration and data transfer
.\archive.ps1       # Pre-reimaging profile backup to ZIP
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
| `Archive.DefaultDestination` | Default backup path for ARCHIVE |
| `Revenant.DefaultDestination` | Default migration destination for REVENANT |
| `Covenant.DefaultTimezone` | Default Windows timezone ID for COVENANT |
| `Covenant.DefaultLocalAdminUser` | Default local administrator account name for COVENANT |

| Script | Configurable Variables |
|--------|------------------------|
| **grimoire.ps1** | None — tool list is defined in the `$Tools` array in the script |
| **covenant.ps1** | `config.json` — `Covenant.DefaultTimezone`, `Covenant.DefaultLocalAdminUser` |
| **conjure.ps1** | `$RequiredSoftware` / `$RequiredSoftwareChoco` — required package IDs; `$OptionalSoftware` / `$OptionalSoftwareChoco` — optional package IDs; `$PackageManager` — default manager (`winget` or `choco`) |
| **runepress.ps1** | `$ExtractRoot` — driver extraction staging folder (defaults to `.\ExtractedDrivers`) |
| **forge.ps1** | None — driver sources scanned from current folder automatically |
| **restoration.ps1** | None — power settings are detected and restored automatically |
| **hearth.ps1** | None — all settings entered via the interactive wizard; `config.json` is the output (see config key table above) |
| **auspex.ps1** | `$ReportOutputPath` — folder where the HTML report is saved (defaults to script directory; accepts any local or UNC path) |
| **ward.ps1** | None — audit runs automatically; stale threshold is 90 days (editable in script) |
| **threshold.ps1** | None — thresholds are Warning < 15% free, Critical < 5% free (editable in script); old profile threshold is 90 days |
| **gargoyle.ps1** | None — critical service list editable in script; `-Target` accepts any WinRM-reachable hostname |
| **augur.ps1** | None — scans all physical disks automatically; `-Unattended` for silent HTML export |
| **cleanse.ps1** | None — categories selected interactively or all cleaned with `-Unattended`; `-WhatIf` for dry run |
| **cipher.ps1** | None — drive and action selected interactively at runtime |
| **sigil.ps1** | None — categories selected interactively; screensaver timeout editable in script (default 600 s) |
| **citadel.ps1** | None — user search and action selected interactively; stale threshold is 90 days (editable in script) |
| **artifact.ps1** | None — stores and targets selected interactively or via `-Targets` parameter |
| **leyline.ps1** | None — all tests run interactively; no persistent config |
| **shade.ps1** | None — target, credentials, and operation selected interactively at runtime |
| **lantern.ps1** | `$script:ScanPorts` — list of TCP ports checked during scan (editable in script) |
| **talisman.ps1** | `-SubscriptionId` — target a specific Azure subscription; `-OutputPath` — HTML report destination; `-NoOpen` — suppress auto-open after export |
| **reliquary.ps1** | None — tenant and report scope selected interactively at runtime |
| **revenant.ps1** | `config.json` — `Revenant.DefaultDestination`; source, items, and destination also selectable interactively |
| **archive.ps1** | `config.json` — `Archive.DefaultDestination`; profile, items, and destination also selectable interactively |

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
| **restoration.ps1** | `%TEMP%\RESTORATION_<timestamp>.log` (PowerShell transcript of the full session) |
| **hearth.ps1** | Console only — settings persisted to `config.json` |
| **auspex.ps1** | `$ReportOutputPath` — `ORACLE_<timestamp>.html` (defaults to Desktop; configurable) |
| **ward.ps1** | Log directory — `WARD_<timestamp>.html` (dark-themed HTML report) |
| **threshold.ps1** | Log directory — `THRESHOLD_<timestamp>.html` (dark-themed HTML report) |
| **gargoyle.ps1** | Log directory — `SENTINEL_<timestamp>.html` (dark-themed HTML health report) |
| **cipher.ps1** | Console only — no log file |
| **sigil.ps1** | Log directory — `SIGIL_BaselineLog_<timestamp>.csv` |
| **citadel.ps1** | Log directory — `BASTION_Stale_<timestamp>.html`; `BASTION_PwdExpiry_<timestamp>.html` |
| **artifact.ps1** | Log directory — `RELIC_<timestamp>.html` (cert inventory & SSL results) |
| **leyline.ps1** | Console only — no log file |
| **shade.ps1** | Script directory — `SPECTER_<MachineName>\` folder containing retrieved output files |
| **lantern.ps1** | Log directory — `LANTERN_<timestamp>.html` and `LANTERN_<timestamp>.csv` |
| **talisman.ps1** | `-OutputPath` (default `%TEMP%`) — `azure-assessment-<timestamp>.html`; auto-opens in browser |
| **reliquary.ps1** | Log directory — `VAULT_<timestamp>.html` (combined license & mailbox report) |
| **revenant.ps1** | Script directory — `PHANTOM_MigrationLog_<timestamp>.csv` |
| **archive.ps1** | Script directory — `ARCHIVE_Log_<timestamp>.csv`; manifest inside ZIP |
| **augur.ps1** | Log directory — `AUGUR_<timestamp>.html` (dark-themed HTML report) |
| **cleanse.ps1** | Console only — cleanup summary printed at completion; no log file |

---

## Contributing

Contributions are welcome. Please ensure all additions maintain:

- Consistent formatting and naming conventions
- The standard `<# .SYNOPSIS / .DESCRIPTION / .USAGE / .NOTES #>` header block
- Comprehensive error handling
- Detailed logging and user feedback
- Administrator privilege checks

---

## Disclaimer

These scripts modify system settings and may install software, updates, or change domain membership in ways that require a reboot. Save all work before running. Use at your own risk.

---

## License

MIT License — see [LICENSE](LICENSE) for details.
