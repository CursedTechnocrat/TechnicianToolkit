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
    Version : 3.0

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
    'Deployment & Onboarding'
    'Diagnostics & Reporting'
    'Security'
    'Network & Remote'
    'Cloud & Identity'
    'Data & Migration'
)

$CategoryKeys = [ordered]@{
    'D' = 'Deployment & Onboarding'
    'R' = 'Diagnostics & Reporting'
    'S' = 'Security'
    'N' = 'Network & Remote'
    'C' = 'Cloud & Identity'
    'M' = 'Data & Migration'
}

$Tools = @(
    # ── Deployment & Onboarding (1–9) ───────────────────────────────
    [PSCustomObject]@{
        Key         = '1'
        Name        = 'C.O.V.E.N.A.N.T.'
        File        = 'covenant.ps1'
        Version     = '3.0'
        Description = 'Machine onboarding, Entra ID domain join, and new device setup'
        Color       = 'Blue'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '2'
        Name        = 'C.O.N.J.U.R.E.'
        File        = 'conjure.ps1'
        Version     = '3.0'
        Description = 'Software deployment via Windows Package Manager or Chocolatey'
        Color       = 'Magenta'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '3'
        Name        = 'R.U.N.E.P.R.E.S.S.'
        File        = 'runepress.ps1'
        Version     = '3.0'
        Description = 'Printer driver installation and network printer configuration'
        Color       = 'Cyan'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '4'
        Name        = 'F.O.R.G.E.'
        File        = 'forge.ps1'
        Version     = '3.0'
        Description = 'Driver detection & installation  -  problem devices, Windows Update, local packages'
        Color       = 'Yellow'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '5'
        Name        = 'R.E.S.T.O.R.A.T.I.O.N.'
        File        = 'restoration.ps1'
        Version     = '3.0'
        Description = 'Automated Windows Update management and maintenance'
        Color       = 'Green'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '6'
        Name        = 'H.E.A.R.T.H.'
        File        = 'hearth.ps1'
        Version     = '3.0'
        Description = 'Toolkit setup wizard  -  org name, log path, Teams webhook, and tool defaults'
        Color       = 'White'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '7'
        Name        = 'R.I.T.U.A.L.'
        File        = 'ritual.ps1'
        Version     = '3.0'
        Description = 'Workflow orchestrator  -  runs named recipes (Onboard, Retire, HealthCheck, TenantSweep) or custom PSD1 files'
        Color       = 'Magenta'
        Category    = 'Deployment & Onboarding'
    },
    # ── Diagnostics & Reporting (10–19) ─────────────────────────────
    [PSCustomObject]@{
        Key         = '10'
        Name        = 'A.U.S.P.E.X.'
        File        = 'auspex.ps1'
        Version     = '3.0'
        Description = 'System diagnostics, health assessment, and HTML report generation'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '11'
        Name        = 'W.A.R.D.'
        File        = 'ward.ps1'
        Version     = '3.0'
        Description = 'User account audit  -  roles, last logon, flags, HTML report'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '12'
        Name        = 'T.H.R.E.S.H.O.L.D.'
        File        = 'threshold.ps1'
        Version     = '3.0'
        Description = 'Disk space monitor  -  volume usage, low-space alerts, temp cleanup, old profile detection'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '13'
        Name        = 'G.A.R.G.O.Y.L.E.'
        File        = 'gargoyle.ps1'
        Version     = '3.0'
        Description = 'Service & task monitor  -  critical services, scheduled tasks, event log errors'
        Color       = 'Red'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '14'
        Name        = 'A.U.G.U.R.'
        File        = 'augur.ps1'
        Version     = '3.0'
        Description = 'Disk wear & health  -  SMART status, physical disk reliability, HTML report'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '15'
        Name        = 'C.L.E.A.N.S.E.'
        File        = 'cleanse.ps1'
        Version     = '3.0'
        Description = 'Disk cleanup  -  temp files, Windows Update cache, browser caches, Recycle Bin'
        Color       = 'Magenta'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '16'
        Name        = 'S.C.R.Y.E.R.'
        File        = 'scryer.ps1'
        Version     = '3.0'
        Description = 'Unified diagnostic report  -  system info, users, disks, SMART, services in one HTML'
        Color       = 'Cyan'
        Category    = 'Diagnostics & Reporting'
    },
    # ── Security (20–29) ─────────────────────────────────────────────
    [PSCustomObject]@{
        Key         = '20'
        Name        = 'C.I.P.H.E.R.'
        File        = 'cipher.ps1'
        Version     = '3.0'
        Description = 'BitLocker drive encryption  -  enable, disable, backup keys'
        Color       = 'Green'
        Category    = 'Security'
    },
    [PSCustomObject]@{
        Key         = '21'
        Name        = 'S.I.G.I.L.'
        File        = 'sigil.ps1'
        Version     = '3.0'
        Description = 'Security baseline enforcement  -  telemetry, UAC, firewall, audit policy'
        Color       = 'Red'
        Category    = 'Security'
    },
    [PSCustomObject]@{
        Key         = '22'
        Name        = 'C.I.T.A.D.E.L.'
        File        = 'citadel.ps1'
        Version     = '3.0'
        Description = 'Active Directory management  -  search, unlock, reset passwords, group membership'
        Color       = 'Blue'
        Category    = 'Security'
    },
    [PSCustomObject]@{
        Key         = '23'
        Name        = 'A.R.T.I.F.A.C.T.'
        File        = 'artifact.ps1'
        Version     = '3.0'
        Description = 'Certificate health monitor  -  local cert stores, SSL/TLS expiry, HTML report'
        Color       = 'Yellow'
        Category    = 'Security'
    },
    # ── Network & Remote (30–39) ─────────────────────────────────────
    [PSCustomObject]@{
        Key         = '30'
        Name        = 'L.E.Y.L.I.N.E.'
        File        = 'leyline.ps1'
        Version     = '3.0'
        Description = 'Network diagnostics & remediation  -  adapters, ping, DNS, port tests'
        Color       = 'Cyan'
        Category    = 'Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '31'
        Name        = 'S.H.A.D.E.'
        File        = 'shade.ps1'
        Version     = '3.0'
        Description = 'Remote execution via WinRM  -  run toolkit tools on a remote machine'
        Color       = 'White'
        Category    = 'Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '32'
        Name        = 'L.A.N.T.E.R.N.'
        File        = 'lantern.ps1'
        Version     = '3.0'
        Description = 'Network discovery & asset inventory  -  subnet sweep, DNS, MAC, port scan'
        Color       = 'Cyan'
        Category    = 'Network & Remote'
    },
    # ── Cloud & Identity (40–49) ─────────────────────────────────────
    [PSCustomObject]@{
        Key         = '40'
        Name        = 'T.A.L.I.S.M.A.N.'
        File        = 'talisman.ps1'
        Version     = '3.0'
        Description = 'Azure environment assessment  -  security posture, RBAC, backup coverage, HTML report'
        Color       = 'Cyan'
        Category    = 'Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '41'
        Name        = 'R.E.L.I.Q.U.A.R.Y.'
        File        = 'reliquary.ps1'
        Version     = '3.0'
        Description = 'M365 license & mailbox audit  -  SKU inventory, unlicensed users, MFA status'
        Color       = 'Green'
        Category    = 'Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '42'
        Name        = 'G.O.L.E.M.'
        File        = 'golem.ps1'
        Version     = '3.0'
        Description = 'Intune / MDM compliance audit  -  managed devices, compliance state, stale devices, config profiles'
        Color       = 'Yellow'
        Category    = 'Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '43'
        Name        = 'W.R.A.I.T.H.'
        File        = 'wraith.ps1'
        Version     = '3.0'
        Description = 'Entra ID identity hygiene audit  -  guests, privileged roles, password-never-expires, stale admins, disabled-but-licensed'
        Color       = 'Red'
        Category    = 'Cloud & Identity'
    },
    # ── Data & Migration (50–59) ─────────────────────────────────────
    [PSCustomObject]@{
        Key         = '50'
        Name        = 'R.E.V.E.N.A.N.T.'
        File        = 'revenant.ps1'
        Version     = '3.0'
        Description = 'Profile migration and data transfer to a new machine'
        Color       = 'Cyan'
        Category    = 'Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '51'
        Name        = 'A.R.C.H.I.V.E.'
        File        = 'archive.ps1'
        Version     = '3.0'
        Description = 'Pre-reimaging profile backup  -  ZIP to local or network share'
        Color       = 'Magenta'
        Category    = 'Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '52'
        Name        = 'T.E.T.H.E.R.'
        File        = 'tether.ps1'
        Version     = '3.0'
        Description = 'OneDrive Known-Folder-Move pre-migration validator  -  client, accounts, KFM, volume, sync errors, HTML report'
        Color       = 'Cyan'
        Category    = 'Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '53'
        Name        = 'E.X.H.U.M.E.'
        File        = 'exhume.ps1'
        Version     = '3.0'
        Description = 'Outlook PST / OST discovery  -  profiles, data files, orphans, oversize, stale archives, HTML report'
        Color       = 'Yellow'
        Category    = 'Data & Migration'
    }
)

# ===========================
# BANNER
# ===========================

function Show-Banner {
    [Console]::Clear()
    Write-Host ""
    Write-Host "  ██████╗ ██████╗ ██╗███╗   ███╗ ██████╗ ██╗██████╗ ███████╗" -ForegroundColor $ColorSchema.Header
    Write-Host " ██╔====╝ ██╔==██╗██║████╗ ████║██╔===██╗██║██╔==██╗██╔====╝" -ForegroundColor $ColorSchema.Header
    Write-Host " ██║  ███╗██████╔╝██║██╔████╔██║██║   ██║██║██████╔╝█████╗  " -ForegroundColor $ColorSchema.Header
    Write-Host " ██║   ██║██╔==██╗██║██║╚██╔╝██║██║   ██║██║██╔==██╗██╔==╝  " -ForegroundColor $ColorSchema.Header
    Write-Host " ╚██████╔╝██║  ██║██║██║ ╚=╝ ██║╚██████╔╝██║██║  ██║███████╗" -ForegroundColor $ColorSchema.Header
    Write-Host "  ╚=====╝ ╚=╝  ╚=╝╚=╝╚=╝     ╚=╝ ╚=====╝ ╚=╝╚=╝  ╚=╝╚======╝" -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  General Repository for Integrated Management and" -ForegroundColor $ColorSchema.Info
    Write-Host "  Orchestration of IT Resources & Executables" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $ColorSchema.Header
    $toolCount = $Tools.Count
    Write-Host "  Technician Toolkit  |  Hub v3.0  |  $toolCount tools  |  Run as Administrator" -ForegroundColor $ColorSchema.Info
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
    Write-Host "  [Q]  Exit GRIMOIRE" -ForegroundColor $ColorSchema.Warning
    Write-Host ""
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

    $SelectedCategory = $CategoryKeys[$CatSelection]

    if (-not $SelectedCategory) {
        Write-Host ""
        Write-Host "  [!!] Invalid selection. Enter a category letter or Q to quit." -ForegroundColor $ColorSchema.Warning
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
