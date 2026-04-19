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
| Need tools with no LiveConnect counterpart (COVENANT, PHANTOM, CIPHER, ARCHIVE, SPECTER, RUNEPRESS, LEYLINE, FORGE, AEGIS, BASTION, LANTERN, THRESHOLD, VAULT, SENTINEL) | **This repo** — these tools are interactive by nature or require auth flows incompatible with LiveConnect |

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

| Script | Acronym | Purpose |
|--------|---------|---------|
| **grimoire.ps1** | **G.R.I.M.O.I.R.E.** — General Repository for Integrated Management and Orchestration of IT Resources & Executables | Central hub launcher — run all tools from one interactive menu |
| **runepress.ps1** | **R.U.N.E.P.R.E.S.S.** — Remote Utility for Networked Equipment — Printer Registration, Extraction & Silent Setup | Printer driver installation and network printer configuration |
| **restoration.ps1** | **R.E.S.T.O.R.A.T.I.O.N.** — Renews Every System Through Orderly Rite — Automating The Installation Of New updates | Automated Windows Update management and maintenance |
| **conjure.ps1** | **C.O.N.J.U.R.E.** — Centrally Orchestrates Network-Joined Updates, Rollouts & Executables | Software deployment via Windows Package Manager or Chocolatey |
| **oracle.ps1** | **O.R.A.C.L.E.** — Observes, Reports & Audits Computer Logs & Environments | System diagnostics, health assessment, and HTML report generation |
| **covenant.ps1** | **C.O.V.E.N.A.N.T.** — Configures Onboarding Via Entra — Network, Accounts, Naming & Timezone | Machine onboarding, Entra ID domain join, and new device setup |
| **phantom.ps1** | **P.H.A.N.T.O.M.** — Portable Home Archive: Navigates & Transfers Objects to new Machine | Profile migration and data transfer between machines or profiles |
| **cipher.ps1** | **C.I.P.H.E.R.** — Configures & Implements Policy-based Hardware Encryption & Recovery | BitLocker drive encryption management — enable, disable, key backup |
| **ward.ps1** | **W.A.R.D.** — Watches Accounts, Reviews Roles & Detects anomalies | Local user account audit with role, last logon, flags, and HTML report |
| **archive.ps1** | **A.R.C.H.I.V.E.** — Automated Repository Compressing & Housing Important Volume Exports | Pre-reimaging profile backup — ZIP to local path or network share |
| **sigil.ps1** | **S.I.G.I.L.** — Secures Infrastructure: Governs via Integrated Lockdown | Security baseline enforcement — telemetry, UAC, firewall, audit policy, password policy |
| **specter.ps1** | **S.P.E.C.T.E.R.** — Sends PowerShell Execution Commands To External Remotes | Remote machine execution via WinRM — run toolkit tools without physical access |
| **leyline.ps1** | **L.E.Y.L.I.N.E.** — Locates, Examines & Yields Latency, Infrastructure, Network & Endpoints | Network diagnostics & remediation — adapters, ping, DNS, port tests, IP renew, stack reset |
| **forge.ps1** | **F.O.R.G.E.** — Finds Outdated Resources & Generates Equipment-updates | Driver detection & installation — problem devices, Windows Update drivers, local packages |
| **aegis.ps1** | **A.E.G.I.S.** — Azure Environment & Governance Inspection System | Azure subscription assessment — security posture, RBAC, backup coverage, Advisor alerts, HTML report |
| **bastion.ps1** | **B.A.S.T.I.O.N.** — Bulk Active-directory Stewardship: Tasks, Identity, Operations & Namespacing | Active Directory user & group management — unlock, reset, enable/disable, stale account report |
| **lantern.ps1** | **L.A.N.T.E.R.N.** — Locates & Audits Network Topology, Enumerating Resources & Nodes | Network discovery — subnet ping sweep, DNS lookup, MAC addresses, port scan, HTML report |
| **threshold.ps1** | **T.H.R.E.S.H.O.L.D.** — Tests Hardware Reliability, Evaluates Storage Health, & Optimizes/Logs Disk data | Disk & storage health — physical disk status, volume space, cleanup, old profiles, HTML report |
| **vault.ps1** | **V.A.U.L.T.** — Validates Assets & User License Tracking | Microsoft 365 license & mailbox audit — license assignments, MFA status, shared mailboxes, HTML report |
| **sentinel.ps1** | **S.E.N.T.I.N.E.L.** — Scans & Evaluates services, Networks, Tasks, Infrastructure, Node Events & Logs | Service, task & event log monitor — health check local or remote machine, HTML report |

---

### G.R.I.M.O.I.R.E.

The central hub for the Technician Toolkit. Presents a categorized, interactive menu to launch any tool without navigating the file system. After a tool completes, control returns to the GRIMOIRE menu automatically.

- Auto-elevates to Administrator on first launch if not already elevated
- Validates that each script file exists before attempting to launch it
- Returns to the hub menu after each tool finishes or errors out
- All tools remain independently runnable without the hub

---

### R.U.N.E.P.R.E.S.S.

Automates printer driver extraction, installation, and network printer configuration via a command-line interface.

- Supports ZIP, EXE, and MSI driver formats
- Handles automatic driver extraction and INF-based installation via pnputil
- Configures network printers via IP (TCP/IP port) or UNC path post-install
- Generates a timestamped installation log (CSV) in the script directory

---

### R.E.S.T.O.R.A.T.I.O.N.

Automates Windows Update detection, installation, and reboot handling with minimal user intervention.

- Disables sleep and display timeout for the duration of the run; restores settings on exit
- Ensures NuGet provider and PSWindowsUpdate module are installed and current
- Installs available updates (drivers excluded) with no forced reboot
- Checks reboot status and prompts only when required
- 30-second reboot countdown with Escape key cancel

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

### O.R.A.C.L.E.

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

### P.H.A.N.T.O.M.

Migrates user profile data from a source machine or profile to a destination using Robocopy for reliable folder transfers.

- Select source from local profiles or enter a custom/UNC path
- Select destination profile or custom path
- Choose individual items or migrate all at once
- Migrates: Desktop, Documents, Downloads, Pictures, Videos, Music
- Migrates: Outlook profiles & data files, email signatures
- Migrates: Chrome bookmarks, Edge bookmarks, Firefox profiles
- Generates a timestamped CSV migration log in the script directory

---

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

### W.A.R.D.

Audits all local user accounts and exports a dark-themed HTML report to the script directory.

- Lists all local accounts: enabled/disabled status, last logon, password info
- Identifies group memberships — flags all Administrator accounts
- Flags potentially risky accounts: no password required, password never set, stale (no logon in 90+ days)
- Console summary with highlighted flagged accounts
- HTML report with color-coded badges and summary cards
- Report saved to script directory as `WARD_<timestamp>.html`

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
- Domain Group Policy takes precedence over local settings where applicable
- Changes logged to `SIGIL_BaselineLog_<timestamp>.csv` in the script directory

---

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

### F.O.R.G.E.

Audits the device tree for driver problems and automates driver installation from multiple sources.

- Scans all devices for errors with human-readable error descriptions (missing, corrupted, cannot start, etc.)
- Checks Windows Update for available driver updates via PSWindowsUpdate (auto-installed if missing)
- Installs drivers from the current folder: ZIP (extracts and runs pnputil on INF), bare INF, EXE (silent), MSI (quiet)
- Exports a full driver inventory CSV with device name, driver version, date, and manufacturer
- Cleans up extracted driver staging folders automatically

---

### S.P.E.C.T.E.R.

Connects to a remote Windows machine via WinRM and runs Technician Toolkit scripts without needing physical access.

- Enter target hostname or IP; supports current credentials (domain/Kerberos) or manual entry
- WinRM connectivity test with step-by-step enable instructions if unreachable
- **Run O.R.A.C.L.E.** — copies script to remote, executes, retrieves HTML report locally
- **Run W.A.R.D.** — copies script to remote, executes, retrieves HTML report locally
- **Run R.E.S.T.O.R.A.T.I.O.N.** — installs Windows Updates on target (reboot warning shown)
- **Run S.I.G.I.L.** — applies full security baseline on target, retrieves CSV log
- **Interactive session** — opens a full `Enter-PSSession` shell on the target
- All output files retrieved to `SPECTER_<MachineName>\` in the script directory
- Remote staging folder cleaned up automatically after each operation
- Target machine prerequisite: `Enable-PSRemoting -Force` (run as Administrator)

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

### A.E.G.I.S.

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

### B.A.S.T.I.O.N.

Interactive Active Directory user and group management tool. Requires RSAT (auto-installed if missing).

- Search and view AD users by name, UPN, or SAM account name
- Unlock locked-out accounts
- Reset user passwords with force-change-on-next-logon option
- Enable and disable user accounts
- View and modify group memberships — add or remove from security/distribution groups
- Stale account report: identifies accounts inactive for 90+ days, exports dark-themed HTML report
- `-Unattended -Action StaleReport` for silent stale account HTML export

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

### V.A.U.L.T.

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

### S.E.N.T.I.N.E.L.

Audits Windows services, scheduled tasks, and recent event log errors — locally or against a remote machine.

- Critical service audit: checks a predefined set of essential services (WinDefend, Spooler, BITS, WMI, W32Time, and more)
- Flags stopped or non-automatic services; offers one-at-a-time restart with confirmation
- Scheduled task audit: lists all active non-Microsoft tasks with last/next run and status
- Event log sweep: Warning and Error events from System and Application logs in the last 24 hours
- Supports remote execution via WinRM with `-Target HOSTNAME`
- Dark-themed HTML health report with color-coded service status badges
- `-Unattended` for silent report export; `-Unattended -Target HOSTNAME` for remote

---

## Requirements

| Requirement | Notes |
|-------------|-------|
| Windows PowerShell 5.1+ | All scripts |
| **TechnicianToolkit.psm1 in the same folder** | All scripts — shared module providing logging, HTML, and privilege helpers |
| Administrator privileges | All scripts (auto-elevation on `grimoire.ps1` and `runepress.ps1`) |
| Internet connectivity | All scripts |
| Windows Package Manager (winget) | `conjure.ps1` (Chocolatey supported as alternative) |
| PSWindowsUpdate module | `restoration.ps1`, `forge.ps1` (auto-installed if missing) |
| Entra ID account with device join permissions | `covenant.ps1` |
| Robocopy (built into Windows) | `phantom.ps1`, `archive.ps1` |
| BitLocker-capable Windows edition (Pro/Enterprise) | `cipher.ps1` |
| WinRM enabled on target machine | `specter.ps1`, `sentinel.ps1` (remote mode) |
| RSAT ActiveDirectory module | `bastion.ps1` (auto-installed if missing) |
| Az PowerShell modules | `aegis.ps1` (auto-installed if missing) |
| Microsoft.Graph modules | `vault.ps1` (auto-installed if missing) |
| Azure subscription + appropriate RBAC | `aegis.ps1` |
| Microsoft 365 tenant + Global Reader or equivalent | `vault.ps1` |
| On-premises Active Directory domain membership | `bastion.ps1` |

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

# A.R.C.H.I.V.E. — Pre-reimaging profile backup
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\archive.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/archive.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.I.P.H.E.R. — BitLocker encryption management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\cipher.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/cipher.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.O.N.J.U.R.E. — Software deployment
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\conjure.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/conjure.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# C.O.V.E.N.A.N.T. — Machine onboarding
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\covenant.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/covenant.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# O.R.A.C.L.E. — System diagnostics and HTML report
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\oracle.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/oracle.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# P.H.A.N.T.O.M. — Profile migration
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\phantom.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/phantom.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\restoration.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/restoration.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# R.U.N.E.P.R.E.S.S. — Printer driver installation
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\runepress.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/runepress.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.I.G.I.L. — Security baseline enforcement
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\sigil.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/sigil.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.P.E.C.T.E.R. — Remote execution via WinRM
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\specter.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/specter.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# W.A.R.D. — Local user account audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\ward.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/ward.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# L.E.Y.L.I.N.E. — Network diagnostics & remediation
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\leyline.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/leyline.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# F.O.R.G.E. — Driver detection & installation
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\forge.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/forge.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# A.E.G.I.S. — Azure environment assessment
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\aegis.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/aegis.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# B.A.S.T.I.O.N. — Active Directory management
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\bastion.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/bastion.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# L.A.N.T.E.R.N. — Network discovery & asset inventory
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\lantern.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/lantern.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# T.H.R.E.S.H.O.L.D. — Disk & storage health monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\threshold.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/threshold.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# V.A.U.L.T. — Microsoft 365 license & mailbox audit
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\vault.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/vault.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f

# S.E.N.T.I.N.E.L. — Service, task & event log monitor
Set-ExecutionPolicy Bypass -Scope Process -Force; $f="$(Get-Location)\sentinel.ps1"; irm https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/sentinel.ps1 -OutFile $f; [IO.File]::WriteAllText($f,[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true)); & $f
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
.\runepress.ps1     # Printer driver installation and configuration
.\restoration.ps1   # Windows Update management
.\conjure.ps1       # Software deployment via winget or Chocolatey
.\oracle.ps1        # System diagnostics and HTML health report
.\covenant.ps1      # New machine onboarding and Entra ID domain join
.\phantom.ps1       # Profile migration and data transfer
.\cipher.ps1        # BitLocker drive encryption management
.\ward.ps1          # User account audit and HTML report
.\archive.ps1       # Pre-reimaging profile backup to ZIP
.\sigil.ps1         # Security baseline enforcement
.\specter.ps1       # Remote execution via WinRM
.\leyline.ps1       # Network diagnostics and remediation
.\forge.ps1         # Driver detection and installation
.\aegis.ps1         # Azure environment assessment and HTML report
.\bastion.ps1       # Active Directory user and group management
.\lantern.ps1       # Network discovery and asset inventory
.\threshold.ps1     # Disk and storage health monitor
.\vault.ps1         # Microsoft 365 license and mailbox audit
.\sentinel.ps1      # Service, task, and event log monitor
```

All scripts must be run as Administrator.

---

## Configuration

Only `conjure.ps1` exposes configurable variables at the top of the file. All other scripts collect their inputs interactively at runtime.

| Script | Configurable Variables |
|--------|------------------------|
| **grimoire.ps1** | None — tool list is defined in the `$Tools` array in the script |
| **runepress.ps1** | `$ExtractRoot` — driver extraction staging folder (defaults to `.\ExtractedDrivers`) |
| **restoration.ps1** | None — power settings are detected and restored automatically |
| **conjure.ps1** | `$RequiredSoftware` / `$RequiredSoftwareChoco` — required package IDs; `$OptionalSoftware` / `$OptionalSoftwareChoco` — optional package IDs; `$PackageManager` — default manager (`winget` or `choco`) |
| **oracle.ps1** | `$ReportOutputPath` — folder where the HTML report is saved (defaults to script directory; accepts any local or UNC path) |
| **covenant.ps1** | None — all settings entered interactively at each step |
| **phantom.ps1** | None — source, destination, and items selected interactively at runtime |
| **cipher.ps1** | None — drive and action selected interactively at runtime |
| **ward.ps1** | None — audit runs automatically; stale threshold is 90 days (editable in script) |
| **archive.ps1** | None — profile, items, and destination selected interactively at runtime |
| **sigil.ps1** | None — categories selected interactively; screensaver timeout editable in script (default 600 s) |
| **specter.ps1** | None — target, credentials, and operation selected interactively at runtime |
| **leyline.ps1** | None — all tests run interactively; no persistent config |
| **forge.ps1** | None — driver sources scanned from current folder automatically |
| **aegis.ps1** | `-SubscriptionId` — target a specific Azure subscription; `-OutputPath` — HTML report destination; `-NoOpen` — suppress auto-open after export |
| **bastion.ps1** | None — user search and action selected interactively; stale threshold is 90 days (editable in script) |
| **lantern.ps1** | `$script:ScanPorts` — list of TCP ports checked during scan (editable in script) |
| **threshold.ps1** | None — thresholds are Warning < 15% free, Critical < 5% free (editable in script); old profile threshold is 90 days |
| **vault.ps1** | None — tenant and report scope selected interactively at runtime |
| **sentinel.ps1** | None — critical service list editable in script; `-Target` accepts any WinRM-reachable hostname |

---

## Logging

| Script | Log Output |
|--------|------------|
| **grimoire.ps1** | No log file — hub activity is visible on-screen only |
| **runepress.ps1** | Script directory — `RUNEPRESS_InstallLog_<timestamp>.csv` |
| **restoration.ps1** | `%TEMP%\RESTORATION_<timestamp>.log` (PowerShell transcript of the full session) |
| **conjure.ps1** | Console — per-package status table printed at completion |
| **oracle.ps1** | `$ReportOutputPath` — `ORACLE_<timestamp>.html` (defaults to Desktop; configurable) |
| **covenant.ps1** | Console — action summary printed at completion |
| **phantom.ps1** | Script directory — `PHANTOM_MigrationLog_<timestamp>.csv` |
| **cipher.ps1** | Console only — no log file |
| **ward.ps1** | Script directory — `WARD_<timestamp>.html` (dark-themed HTML report) |
| **archive.ps1** | Script directory — `ARCHIVE_Log_<timestamp>.csv`; manifest inside ZIP |
| **sigil.ps1** | Script directory — `SIGIL_BaselineLog_<timestamp>.csv` |
| **specter.ps1** | Script directory — `SPECTER_<MachineName>\` folder containing retrieved output files |
| **leyline.ps1** | Console only — no log file |
| **forge.ps1** | Script directory — `FORGE_DriverReport_<timestamp>.csv` |
| **aegis.ps1** | `-OutputPath` (default `%TEMP%`) — `azure-assessment-<timestamp>.html`; auto-opens in browser |
| **bastion.ps1** | Script directory — `BASTION_StaleAccounts_<timestamp>.html` (stale account report mode only) |
| **lantern.ps1** | Script directory — `LANTERN_<timestamp>.html` and `LANTERN_<timestamp>.csv` |
| **threshold.ps1** | Script directory — `THRESHOLD_<timestamp>.html` (dark-themed HTML report) |
| **vault.ps1** | Script directory — `VAULT_<timestamp>.html` (combined license & mailbox report) |
| **sentinel.ps1** | Script directory — `SENTINEL_<timestamp>.html` (dark-themed HTML health report) |

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

[Add license information here]
