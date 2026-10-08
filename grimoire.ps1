# grimoire.ps1 - G.R.I.M.O.I.R.E. — General Repository for Integrated Management and Orchestration of IT Resources & Executables
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
    G.R.I.M.O.I.R.E. — General Repository for Integrated Management and Orchestration of IT Resources & Executables
    Technician Toolkit Hub for PowerShell 5.1+

.DESCRIPTION
    Central launcher for the Technician Toolkit. Presents an interactive menu
    to select and run any of the available tools, then returns to the hub on
    completion.

.USAGE
    PS C:\> .\grimoire.ps1           # Must be run as Administrator
    PS C:\> .\grimoire.ps1 -WhatIf   # Launch tools in dry-run mode (passed through to each tool that supports it)

.NOTES
    Version : 5.1

#>

param(
    [switch]$WhatIf
)

# ===========================
# ADMIN PRIVILEGE CHECK
# ===========================
# ===========================
# INITIALIZATION
# ===========================

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

$ScriptPath      = (Get-Location).Path
$DownloadedFiles = [System.Collections.Generic.List[string]]::new()

# Counts non-CODEX tool runs in the current Grimoire session. Drives the
# "roll up reports" hint on the main category menu so the [X] shortcut
# becomes visually prominent after the user has run multiple tools.
$script:ToolRunCount = 0

$ColorSchema = @{
    Header  = 'Cyan'
    Accent  = 'Magenta'
    Success = 'Green'
    Warning = 'Yellow'
    Error   = 'Red'
    Info    = 'Gray'
    Menu    = 'White'
}

# ===========================
# TOOL REGISTRY
# ===========================

$CategoryOrder = @(
    'The Workshop — Deployment & Onboarding'
    'The Observatory — Diagnostics & Reporting'
    'The Bestiary — Security'
    'The Crossroads — Network & Remote'
    'The Firmament — Cloud & Identity'
    'The Necropolis — Data & Migration'
)

# Menu letters follow the plain half of each name, so they stay the ones
# technicians already know. Each category's tools take their names from its
# theme (see CLAUDE.md, 'Naming').
$CategoryKeys = [ordered]@{
    'D' = 'The Workshop — Deployment & Onboarding'
    'R' = 'The Observatory — Diagnostics & Reporting'
    'S' = 'The Bestiary — Security'
    'N' = 'The Crossroads — Network & Remote'
    'C' = 'The Firmament — Cloud & Identity'
    'M' = 'The Necropolis — Data & Migration'
}

$Tools = @(
    # ── The Workshop (1–19) — the artificer: crafting & enchanting ──────────
    [PSCustomObject]@{
        Key         = '1'
        Name        = 'C.O.V.E.N.A.N.T.'
        File        = 'covenant.ps1'
        Version     = '5.1'
        Description = 'Machine onboarding, Entra ID domain join, and new device setup'
        Color       = 'Blue'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '2'
        Name        = 'C.O.N.J.U.R.E.'
        File        = 'conjure.ps1'
        Version     = '5.1'
        Description = 'Software deployment via Windows Package Manager or Chocolatey'
        Color       = 'Magenta'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '3'
        Name        = 'R.U.N.E.P.R.E.S.S.'
        File        = 'runepress.ps1'
        Version     = '5.1'
        Description = 'Printer driver installation and network printer configuration'
        Color       = 'Cyan'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '4'
        Name        = 'F.O.R.G.E.'
        File        = 'forge.ps1'
        Version     = '5.1'
        Description = 'Driver detection & installation  -  problem devices, Windows Update, local packages'
        Color       = 'Yellow'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '5'
        Name        = 'W.H.E.T.S.T.O.N.E.'
        File        = 'whetstone.ps1'
        Version     = '5.1'
        Description = 'Automated Windows Update management and maintenance'
        Color       = 'Green'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '6'
        Name        = 'H.E.A.R.T.H.'
        File        = 'hearth.ps1'
        Version     = '5.1'
        Description = 'Toolkit setup wizard  -  org name, log path, Teams webhook, and tool defaults'
        Color       = 'White'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '7'
        Name        = 'R.I.T.U.A.L.'
        File        = 'ritual.ps1'
        Version     = '5.1'
        Description = 'Workflow orchestrator  -  runs named recipes (Onboard, Retire, HealthCheck, SecuritySweep, NetworkSweep, TenantSweep) or custom PSD1 files'
        Color       = 'Magenta'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '8'
        Name        = 'C.O.N.D.U.I.T.'
        File        = 'conduit.ps1'
        Version     = '5.1'
        Description = 'Windows Update connectivity diagnosis & repair  -  WSUS pointer, WinHTTP proxy, policy source, update services'
        Color       = 'Cyan'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '9'
        Name        = 'S.O.L.D.E.R.'
        File        = 'solder.ps1'
        Version     = '5.1'
        Description = 'Windows servicing repair  -  component store health, pending restarts, last SFC result, update failure codes, DISM RestoreHealth + SFC'
        Color       = 'Green'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '10'
        Name        = 'C.L.E.A.N.S.E.'
        File        = 'cleanse.ps1'
        Version     = '5.1'
        Description = 'Disk cleanup  -  temp files, Windows Update cache, browser caches, Recycle Bin'
        Color       = 'Magenta'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '11'
        Name        = 'A.N.V.I.L.'
        File        = 'anvil.ps1'
        Version     = '5.1'
        Description = 'BIOS / UEFI / firmware audit  -  system identity, Secure Boot, vendor channels, pending WU updates'
        Color       = 'Yellow'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '12'
        Name        = 'C.H.A.L.I.C.E.'
        File        = 'chalice.ps1'
        Version     = '5.1'
        Description = 'Microsoft 365 Apps client repair  -  version, channel, support dates, license tokens, cached accounts, WAM / token broker, PRT, Teams cache; reset sign-in, Teams, Quick Repair'
        Color       = 'Yellow'
        Category    = 'The Workshop — Deployment & Onboarding'
    },
    # ── The Observatory (20–39) — scrying & magical sight ───────────────────
    [PSCustomObject]@{
        Key         = '20'
        Name        = 'A.U.S.P.E.X.'
        File        = 'auspex.ps1'
        Version     = '5.1'
        Description = 'System diagnostics, health assessment, and HTML report generation'
        Color       = 'Yellow'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '21'
        Name        = 'S.C.R.Y.E.R.'
        File        = 'scryer.ps1'
        Version     = '5.1'
        Description = 'Unified diagnostic report  -  system info, users, disks, SMART, services in one HTML'
        Color       = 'Cyan'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '22'
        Name        = 'A.U.G.U.R.'
        File        = 'augur.ps1'
        Version     = '5.1'
        Description = 'Disk wear & health  -  SMART status, physical disk reliability, HTML report'
        Color       = 'Yellow'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '23'
        Name        = 'F.A.T.H.O.M.'
        File        = 'fathom.ps1'
        Version     = '5.1'
        Description = 'Disk space monitor  -  volume usage, low-space alerts, temp cleanup, old profile detection'
        Color       = 'Yellow'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '24'
        Name        = 'V.I.G.I.L.'
        File        = 'vigil.ps1'
        Version     = '5.1'
        Description = 'Service & task monitor  -  critical services, scheduled tasks, event log errors'
        Color       = 'Red'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '25'
        Name        = 'H.O.U.R.G.L.A.S.S.'
        File        = 'hourglass.ps1'
        Version     = '5.1'
        Description = 'Laptop battery health audit  -  design vs current capacity, cycle count, replacement verdict, powercfg /batteryreport enrichment'
        Color       = 'Red'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '26'
        Name        = 'N.E.C.R.O.P.S.Y.'
        File        = 'necropsy.ps1'
        Version     = '5.1'
        Description = 'Crash & unexpected-reboot analysis  -  bugchecks, Kernel-Power 41, WHEA hardware errors, display resets, dump files, change timeline'
        Color       = 'Red'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '27'
        Name        = 'T.O.R.P.O.R.'
        File        = 'torpor.ps1'
        Version     = '5.1'
        Description = 'Slow-machine triage  -  CPU, memory and disk load over a sample, top processes by name, CPU limits, boot delays, startup programs, app hangs'
        Color       = 'Yellow'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '28'
        Name        = 'C.O.D.E.X.'
        File        = 'codex.ps1'
        Version     = '5.1'
        Description = 'Toolkit report index  -  scans log directory for existing HTML reports, groups by tool, emits a single rollup with relative links'
        Color       = 'Blue'
        Category    = 'The Observatory — Diagnostics & Reporting'
    },
    # ── The Bestiary (40–59) — guardian beasts ──────────────────────────────
    [PSCustomObject]@{
        Key         = '40'
        Name        = 'W.Y.R.M.'
        File        = 'wyrm.ps1'
        Version     = '5.1'
        Description = 'BitLocker drive encryption  -  enable, disable, backup keys'
        Color       = 'Green'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '41'
        Name        = 'B.A.S.I.L.I.S.K.'
        File        = 'basilisk.ps1'
        Version     = '5.1'
        Description = 'Security baseline enforcement  -  telemetry, UAC, firewall, audit policy'
        Color       = 'Red'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '42'
        Name        = 'S.P.H.I.N.X.'
        File        = 'sphinx.ps1'
        Version     = '5.1'
        Description = 'Active Directory management  -  search, unlock, reset passwords, group membership'
        Color       = 'Blue'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '43'
        Name        = 'P.H.O.E.N.I.X.'
        File        = 'phoenix.ps1'
        Version     = '5.1'
        Description = 'Certificate health monitor  -  local cert stores, SSL/TLS expiry, HTML report'
        Color       = 'Yellow'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '44'
        Name        = 'T.A.L.O.N.'
        File        = 'talon.ps1'
        Version     = '5.1'
        Description = 'Persistence / autoruns audit  -  Run keys, startup folders, services, tasks, WMI, IFEO, Winlogon'
        Color       = 'Red'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '45'
        Name        = 'T.O.T.E.M.'
        File        = 'totem.ps1'
        Version     = '5.1'
        Description = 'TPM health audit  -  presence, spec, ownership, readiness, BitLocker dependency, attestation'
        Color       = 'Cyan'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '46'
        Name        = 'G.R.I.F.F.I.N.'
        File        = 'griffin.ps1'
        Version     = '5.1'
        Description = 'AV / Microsoft Defender health audit  -  state, signatures, scans, threats, exclusions, ASR rules, services'
        Color       = 'Magenta'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '47'
        Name        = 'A.R.G.U.S.'
        File        = 'argus.ps1'
        Version     = '5.1'
        Description = 'Active Directory account roster  -  domain password & lockout policy verdicts, full name / alias / access level per account, nested group expansion, privileged group membership, HTML + review CSV'
        Color       = 'Cyan'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '48'
        Name        = 'M.I.N.O.T.A.U.R.'
        File        = 'minotaur.ps1'
        Version     = '5.1'
        Description = 'File share & NTFS permissions review  -  share and NTFS access, everyone-type write, direct user grants, orphaned SIDs, broken inheritance, HTML + CSV'
        Color       = 'Yellow'
        Category    = 'The Bestiary — Security'
    },
    [PSCustomObject]@{
        Key         = '49'
        Name        = 'W.A.R.D.'
        File        = 'ward.ps1'
        Version     = '5.1'
        Description = 'User account audit  -  roles, last logon, flags, HTML report'
        Color       = 'Yellow'
        Category    = 'The Bestiary — Security'
    },
    # ── The Crossroads (60–69) — arcane paths & lights ──────────────────────
    [PSCustomObject]@{
        Key         = '60'
        Name        = 'L.E.Y.L.I.N.E.'
        File        = 'leyline.ps1'
        Version     = '5.1'
        Description = 'Network diagnostics & remediation  -  adapters, ping, DNS, port tests'
        Color       = 'Cyan'
        Category    = 'The Crossroads — Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '61'
        Name        = 'E.M.I.S.S.A.R.Y.'
        File        = 'emissary.ps1'
        Version     = '5.1'
        Description = 'Remote execution via WinRM  -  run toolkit tools on a remote machine'
        Color       = 'White'
        Category    = 'The Crossroads — Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '62'
        Name        = 'L.A.N.T.E.R.N.'
        File        = 'lantern.ps1'
        Version     = '5.1'
        Description = 'Network discovery & asset inventory  -  subnet sweep, DNS, MAC, port scan'
        Color       = 'Cyan'
        Category    = 'The Crossroads — Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '63'
        Name        = 'W.I.S.P.'
        File        = 'wisp.ps1'
        Version     = '5.1'
        Description = 'Wi-Fi profile audit  -  saved profiles, auth/cipher tier, auto-connect, hidden SSID, MAC randomisation, key material'
        Color       = 'Yellow'
        Category    = 'The Crossroads — Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '64'
        Name        = 'P.O.R.T.A.L.'
        File        = 'portal.ps1'
        Version     = '5.1'
        Description = 'VPN / Always-On VPN audit  -  built-in connections, app triggers, NRPT, tunnel interfaces, third-party clients'
        Color       = 'Green'
        Category    = 'The Crossroads — Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '65'
        Name        = 'L.O.D.E.S.T.A.R.'
        File        = 'lodestar.ps1'
        Version     = '5.1'
        Description = 'Domain trust & secure channel diagnosis and repair  -  DC discovery, DNS, DC ports, clock skew, nltest secure channel, machine password reset'
        Color       = 'Green'
        Category    = 'The Crossroads — Network & Remote'
    },
    # ── The Firmament (70–89) — the sky & the celestial ─────────────────────
    [PSCustomObject]@{
        Key         = '70'
        Name        = 'Z.E.N.I.T.H.'
        File        = 'zenith.ps1'
        Version     = '5.1'
        Description = 'Azure environment assessment  -  security posture, RBAC, backup coverage, HTML report'
        Color       = 'Cyan'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '71'
        Name        = 'A.L.M.A.N.A.C.'
        File        = 'almanac.ps1'
        Version     = '5.1'
        Description = 'M365 license & mailbox audit  -  SKU inventory, unlicensed users, MFA status'
        Color       = 'Green'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '72'
        Name        = 'O.R.B.I.T.'
        File        = 'orbit.ps1'
        Version     = '5.1'
        Description = 'Intune / MDM compliance audit  -  managed devices, compliance state, stale devices, config profiles'
        Color       = 'Yellow'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '73'
        Name        = 'E.C.L.I.P.S.E.'
        File        = 'eclipse.ps1'
        Version     = '5.1'
        Description = 'Entra ID identity hygiene audit  -  guests, privileged roles, password-never-expires, stale admins, disabled-but-licensed'
        Color       = 'Red'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '74'
        Name        = 'A.S.T.E.R.I.S.M.'
        File        = 'asterism.ps1'
        Version     = '5.1'
        Description = 'Microsoft Teams audit  -  orphan teams, public teams, guest membership, large teams, stale teams'
        Color       = 'Magenta'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '75'
        Name        = 'C.U.M.U.L.U.S.'
        File        = 'cumulus.ps1'
        Version     = '5.1'
        Description = 'SharePoint Online audit  -  site inventory, storage, external sharing, ownerless and stale sites'
        Color       = 'Green'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '76'
        Name        = 'O.R.R.E.R.Y.'
        File        = 'orrery.ps1'
        Version     = '5.1'
        Description = 'Entra ID group dependency audit  -  what breaks if we delete this group? Licensing, CA, apps, roles, AUs, Intune, SP, EXO, Azure RBAC'
        Color       = 'Blue'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '77'
        Name        = 'R.A.V.E.N.'
        File        = 'raven.ps1'
        Version     = '5.1'
        Description = 'Exchange Online mailbox security audit  -  external forwarding, suspicious inbox rules, SMTP AUTH, auditing, delegation, SPF / DKIM / DMARC'
        Color       = 'Magenta'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '78'
        Name        = 'H.A.L.O.'
        File        = 'halo.ps1'
        Version     = '5.1'
        Description = 'Entra ID Conditional Access posture  -  MFA and legacy-auth baseline, admin coverage, report-only and disabled policies, exclusions, break-glass, named locations'
        Color       = 'Cyan'
        Category    = 'The Firmament — Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '79'
        Name        = 'C.A.R.I.L.L.O.N.'
        File        = 'carillon.ps1'
        Version     = '5.1'
        Description = 'Teams Phone call queues & auto attendants  -  agents by display name, opt-in and voice state, overflow / timeout routing, menus, resource accounts'
        Color       = 'Magenta'
        Category    = 'The Firmament — Cloud & Identity'
    },
    # ── The Necropolis (90–99) — the undead & ghosts ────────────────────────
    [PSCustomObject]@{
        Key         = '90'
        Name        = 'R.E.V.E.N.A.N.T.'
        File        = 'revenant.ps1'
        Version     = '5.1'
        Description = 'Profile migration and data transfer to a new machine'
        Color       = 'Cyan'
        Category    = 'The Necropolis — Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '91'
        Name        = 'E.M.B.A.L.M.'
        File        = 'embalm.ps1'
        Version     = '5.1'
        Description = 'Pre-reimaging profile backup  -  ZIP to local or network share'
        Color       = 'Magenta'
        Category    = 'The Necropolis — Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '92'
        Name        = 'P.H.Y.L.A.C.T.E.R.Y.'
        File        = 'phylactery.ps1'
        Version     = '5.1'
        Description = 'OneDrive Known-Folder-Move pre-migration validator  -  client, accounts, KFM, volume, sync errors, HTML report'
        Color       = 'Cyan'
        Category    = 'The Necropolis — Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '93'
        Name        = 'E.X.H.U.M.E.'
        File        = 'exhume.ps1'
        Version     = '5.1'
        Description = 'Outlook PST / OST discovery  -  profiles, data files, orphans, oversize, stale archives, HTML report'
        Color       = 'Yellow'
        Category    = 'The Necropolis — Data & Migration'
    }
)

# ===========================
# BANNER
# ===========================

function Show-Banner {
    [Console]::Clear()
    Write-Host ""
    Write-Host "  ██████╗ ██████╗ ██╗███╗   ███╗ ██████╗ ██╗██████╗ ███████╗" -ForegroundColor $ColorSchema.Header
    Write-Host " ██╔════╝ ██╔══██╗██║████╗ ████║██╔═══██╗██║██╔══██╗██╔════╝" -ForegroundColor $ColorSchema.Header
    Write-Host " ██║  ███╗██████╔╝██║██╔████╔██║██║   ██║██║██████╔╝█████╗  " -ForegroundColor $ColorSchema.Header
    Write-Host " ██║   ██║██╔══██╗██║██║╚██╔╝██║██║   ██║██║██╔══██╗██╔══╝  " -ForegroundColor $ColorSchema.Header
    Write-Host " ╚██████╔╝██║  ██║██║██║ ╚═╝ ██║╚██████╔╝██║██║  ██║███████╗" -ForegroundColor $ColorSchema.Header
    Write-Host "  ╚═════╝ ╚═╝  ╚═╝╚═╝╚═╝     ╚═╝ ╚═════╝ ╚═╝╚═╝  ╚═╝╚══════╝" -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  General Repository for Integrated Management and" -ForegroundColor $ColorSchema.Info
    Write-Host "  Orchestration of IT Resources & Executables" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    $toolCount = $Tools.Count
    Write-Host "  Technician Toolkit  |  Hub v3.6  |  $toolCount tools  |  Run as Administrator" -ForegroundColor $ColorSchema.Info
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    # GPLv3 Appropriate Legal Notices. The toolkit is free software and is meant
    # to be shared, forked, and improved by the technicians who use it -- say so
    # where every operator will actually see it.
    Write-Host "  Free software under the GNU GPL v3 or later, with NO WARRANTY." -ForegroundColor $ColorSchema.Info
    Write-Host "  Source, licence, and contributions: github.com/CursedTechnocrat/TechnicianToolkit" -ForegroundColor $ColorSchema.Info
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    if ($WhatIf) {
        Write-Host ""
        Write-Host "  *** DRY RUN MODE  -  tools that support -WhatIf will preview actions only ***" -ForegroundColor Cyan
    }
    Write-Host ""
}

# ===========================
# MENU
# ===========================

function Show-Menu {
    Write-Host "  Select a category:" -ForegroundColor $ColorSchema.Header
    Write-Host ""

    foreach ($key in $CategoryKeys.Keys) {
        $cat   = $CategoryKeys[$key]
        $count = ($Tools | Where-Object { $_.Category -eq $cat }).Count
        Write-Host "  [$key]  $cat" -NoNewline -ForegroundColor $ColorSchema.Menu
        Write-Host "  ($count tools)" -ForegroundColor $ColorSchema.Info
    }

    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header

    # CODEX is registered in The Observatory (key 28) for completeness,
    # but is also surfaced here as a top-level shortcut so the rollup is one
    # keystroke away after a multi-tool session. The hint and colour shift
    # once the session has produced something worth indexing (>= 2 runs).
    if ($script:ToolRunCount -ge 2) {
        Write-Host "  [X]  Roll up reports (C.O.D.E.X.)" -NoNewline -ForegroundColor $ColorSchema.Accent
        Write-Host "  -  $($script:ToolRunCount) tool run(s) this session" -ForegroundColor $ColorSchema.Info
    } else {
        Write-Host "  [X]  Roll up reports (C.O.D.E.X.)" -ForegroundColor $ColorSchema.Menu
    }
    Write-Host "  [F]  Find a tool by name or keyword" -ForegroundColor $ColorSchema.Menu
    Write-Host "  [Q]  Exit GRIMOIRE" -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host "  Tip: type a tool name or number (e.g. 'hourglass' or '25') to jump straight to it." -ForegroundColor $ColorSchema.Info
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
}

function Show-CategoryMenu {
    param([string]$Category)

    [Console]::Clear()
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  GRIMOIRE  /  $Category" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""

    foreach ($tool in ($Tools | Where-Object { $_.Category -eq $Category })) {
        Write-Host "  [$($tool.Key)]  $($tool.Name)  " -NoNewline -ForegroundColor $tool.Color
        Write-Host "v$($tool.Version)" -ForegroundColor $ColorSchema.Info
        Write-Host "       $($tool.Description)" -ForegroundColor $ColorSchema.Info
        Write-Host ""
    }

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  [B]  Back to categories" -ForegroundColor $ColorSchema.Warning
    Write-Host "  [Q]  Exit GRIMOIRE" -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
}

# ===========================
# TOOL LAUNCHER
# ===========================

$BaseUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main'

function Invoke-Tool {
    param([PSCustomObject]$Tool)

    $ToolPath = Join-Path $ScriptPath $Tool.File

    if (-not (Test-Path $ToolPath)) {
        $DownloadUrl = "$BaseUrl/$($Tool.File)"
        Write-Host ""
        Write-Host "  Downloading $($Tool.File) from GitHub..." -ForegroundColor $ColorSchema.Accent
        try {
            Invoke-RestMethod -Uri $DownloadUrl -OutFile $ToolPath -ErrorAction Stop
            [IO.File]::WriteAllText($ToolPath, [IO.File]::ReadAllText($ToolPath, [Text.Encoding]::UTF8), [Text.UTF8Encoding]::new($true))

            # Validate the downloaded file parses as valid PowerShell before executing it
            $parseErrors = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($ToolPath, [ref]$null, [ref]$parseErrors)
            if ($parseErrors.Count -gt 0) {
                Remove-Item -Path $ToolPath -Force -ErrorAction SilentlyContinue
                Write-Host ""
                Write-Host "  [!!] $($Tool.File) failed syntax validation after download  -  file removed." -ForegroundColor $ColorSchema.Error
                Write-Host "       $($parseErrors[0].Message)" -ForegroundColor $ColorSchema.Error
                Write-Host ""
                Pause-ForKey
                return
            }

            $DownloadedFiles.Add($ToolPath)
            Write-Host "  Downloaded and verified successfully." -ForegroundColor $ColorSchema.Success
        }
        catch {
            Write-Host ""
            Write-Host "  [!!] Could not download $($Tool.File):" -ForegroundColor $ColorSchema.Error
            Write-Host "       $($_.Exception.Message)" -ForegroundColor $ColorSchema.Error
            Write-Host ""
            Pause-ForKey
            return
        }
    }

    Write-Host ""
    Write-Host "  Launching $($Tool.Name)..." -ForegroundColor $ColorSchema.Accent
    Write-Host ""
    Start-Sleep -Milliseconds 600

    # Build argument list — pass -WhatIf only if the target script accepts it
    $toolArgs = @{}
    if ($WhatIf) {
        $toolCmd = Get-Command $ToolPath -ErrorAction SilentlyContinue
        if ($toolCmd -and $toolCmd.Parameters.ContainsKey('WhatIf')) {
            $toolArgs['WhatIf'] = $true
        }
    }

    try {
        & $ToolPath @toolArgs
    }
    catch {
        Write-Host ""
        Write-Host "  [!!] $($Tool.Name) exited with an error:" -ForegroundColor $ColorSchema.Error
        Write-Host "       $($_.Exception.Message)" -ForegroundColor $ColorSchema.Error
    }

    # Track non-CODEX tool runs so the main-menu rollup hint can highlight
    # once the session has produced multiple reports. CODEX itself doesn't
    # contribute -- it's the indexer, not an indexable diagnostic.
    if ($Tool.Name -ne 'C.O.D.E.X.') { $script:ToolRunCount++ }

    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  $($Tool.Name) has finished. Returning to GRIMOIRE..." -ForegroundColor $ColorSchema.Accent
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Pause-ForKey
}

function Pause-ForKey {
    Write-Host ""
    Write-Host "  Press any key to continue..." -ForegroundColor $ColorSchema.Info
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ===========================
# TOOL SEARCH
# ===========================

function Find-Tool {
    # Resolve a free-text query to one or more tools. Matching order:
    #   1. Exact tool key       (e.g. '25')
    #   2. Exact acronym/file   (e.g. 'hourglass', 'H.O.U.R.G.L.A.S.S.', 'hourglass.ps1')
    #   3. Substring on acronym/filename, plus a keyword scan of descriptions
    # Acronyms are normalised by dropping dots/spaces/hyphens so 'H.O.U.R.G.L.A.S.S.',
    # 'hourglass' and 'HOURGLASS' all collapse to the same token.
    param([string]$Query)

    $q = $Query.Trim()
    if ([string]::IsNullOrWhiteSpace($q)) { return @() }

    $byKey = @($Tools | Where-Object { $_.Key -eq $q })
    if ($byKey.Count -gt 0) { return $byKey }

    $norm = ($q -replace '[\s\.\-]', '').ToLower()
    $ql   = $q.ToLower()
    if (-not $norm) { return @() }

    $exact = @($Tools | Where-Object {
        ((($_.Name -replace '[\s\.\-]', '').ToLower()) -eq $norm) -or
        ((($_.File -replace '\.ps1$', '').ToLower())    -eq $norm)
    })
    if ($exact.Count -gt 0) { return $exact }

    $byName = @($Tools | Where-Object {
        ((($_.Name -replace '[\s\.\-]', '').ToLower()) -like "*$norm*") -or
        ((($_.File -replace '\.ps1$', '').ToLower())    -like "*$norm*")
    })
    # Description keyword search only for queries of 3+ chars, so a stray single
    # letter doesn't dump the whole catalogue.
    $byDesc = @()
    if ($ql.Length -ge 3) {
        $byDesc = @($Tools | Where-Object { $_.Description.ToLower() -like "*$ql*" })
    }

    $seen   = @{}
    $result = foreach ($t in @($byName + $byDesc)) {
        if (-not $seen.ContainsKey($t.Key)) { $seen[$t.Key] = $true; $t }
    }
    # Filter guards the no-match case: an empty foreach yields $null, and
    # @($null) would otherwise report a phantom count of 1.
    return @($result | Where-Object { $_ })
}

function Select-FromMatches {
    # Given the result of Find-Tool, return the single tool to run. One match
    # launches straight away; several render a pick-list; [B]/blank cancels.
    param([object[]]$Found)

    if ($Found.Count -eq 1) { return $Found[0] }

    [Console]::Clear()
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  GRIMOIRE  /  Search results ($($Found.Count))" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""

    foreach ($tool in $Found) {
        Write-Host "  [$($tool.Key)]  $($tool.Name)  " -NoNewline -ForegroundColor $tool.Color
        Write-Host "v$($tool.Version)" -ForegroundColor $ColorSchema.Info
        Write-Host "       $($tool.Description)" -ForegroundColor $ColorSchema.Info
        Write-Host "       $($tool.Category)" -ForegroundColor $ColorSchema.Accent
        Write-Host ""
    }

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  Enter a tool number to launch, or [B] to go back." -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Menu
    $pick = (Read-Host).Trim().ToUpper()
    if ($pick -eq 'B' -or $pick -eq '') { return $null }
    return ($Found | Where-Object { $_.Key -eq $pick } | Select-Object -First 1)
}

# ===========================
# MAIN LOOP
# ===========================

$exitGrimoire = $false

do {
    Show-Banner
    Show-Menu

    Write-Host -NoNewline "  Enter category: " -ForegroundColor $ColorSchema.Menu
    $CatSelection = (Read-Host).Trim().ToUpper()

    if ($CatSelection -eq 'Q') {
        $exitGrimoire = $true
        break
    }

    # Top-level shortcut: roll up every existing HTML report in the log
    # directory via CODEX without drilling into the Diagnostics category.
    if ($CatSelection -eq 'X') {
        $codexTool = $Tools | Where-Object { $_.File -eq 'codex.ps1' } | Select-Object -First 1
        if ($codexTool) {
            Invoke-Tool -Tool $codexTool
        } else {
            Write-Host ""
            Write-Host "  [!!] CODEX is not registered in this Grimoire build." -ForegroundColor $ColorSchema.Warning
            Start-Sleep -Seconds 1
        }
        continue
    }

    # Explicit search prompt.
    if ($CatSelection -eq 'F') {
        Write-Host -NoNewline "  Search tools: " -ForegroundColor $ColorSchema.Menu
        $query = (Read-Host).Trim()
        if ($query) {
            $found = Find-Tool -Query $query
            if ($found.Count -gt 0) {
                $chosen = Select-FromMatches -Found $found
                if ($chosen) { Invoke-Tool -Tool $chosen }
            } else {
                Write-Host ""
                Write-Host "  [!!] No tool matched '$query'." -ForegroundColor $ColorSchema.Warning
                Start-Sleep -Seconds 1
            }
        }
        continue
    }

    $SelectedCategory = $CategoryKeys[$CatSelection]

    if (-not $SelectedCategory) {
        # Not a category letter — treat the input as a tool name/number search
        # so 'hourglass' or '25' jumps straight to the tool from the main menu.
        $found = Find-Tool -Query $CatSelection
        if ($found.Count -gt 0) {
            $chosen = Select-FromMatches -Found $found
            if ($chosen) { Invoke-Tool -Tool $chosen }
            continue
        }
        Write-Host ""
        Write-Host "  [!!] No category or tool matched '$CatSelection'. Enter a category letter, a tool name/number, [F] to search, [X] to roll up reports, or [Q] to quit." -ForegroundColor $ColorSchema.Warning
        Start-Sleep -Seconds 1
        continue
    }

    $backToMain = $false
    do {
        Show-CategoryMenu -Category $SelectedCategory

        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Menu
        $Selection = (Read-Host).Trim().ToUpper()

        if ($Selection -eq 'Q') {
            $exitGrimoire = $true
            $backToMain   = $true
        }
        elseif ($Selection -eq 'B') {
            $backToMain = $true
        }
        else {
            $MatchedTool = $Tools | Where-Object { $_.Key -eq $Selection -and $_.Category -eq $SelectedCategory }
            if ($MatchedTool) {
                Invoke-Tool -Tool $MatchedTool
            }
            else {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter a tool number, [B] to go back, or [Q] to quit." -ForegroundColor $ColorSchema.Warning
                Start-Sleep -Seconds 1
            }
        }

    } while (-not $backToMain)

} while (-not $exitGrimoire)

[Console]::Clear()
Write-Host ""
Write-Host "  Closing GRIMOIRE. Stay arcane." -ForegroundColor $ColorSchema.Header
Write-Host ""

# ===========================
# CLEANUP
# ===========================

foreach ($f in $DownloadedFiles) {
    Remove-Item -Path $f -Force -ErrorAction SilentlyContinue
}

# Self-delete grimoire.ps1 after a one-shot bootstrapped session so the host
# stays clean when run from `irm … | iex`-style quick-launch snippets.
# Skip the self-delete when the script lives inside a git checkout — that's a
# clone, not a throwaway download, and silently removing it would lose work.
if ($PSCommandPath) {
    $isRepoCheckout = Test-Path (Join-Path $PSScriptRoot '.git')
    if (-not $isRepoCheckout) {
        Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue
    }
}
