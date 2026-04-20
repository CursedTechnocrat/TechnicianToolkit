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
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    P.H.A.N.T.O.M.         — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    S.I.G.I.L.             — Security baseline & policy enforcement
    S.P.E.C.T.E.R.         — Remote machine execution via WinRM
    L.E.Y.L.I.N.E.         — Network diagnostics & remediation
    F.O.R.G.E.             — Driver update detection & installation
    A.E.G.I.S.             — Azure environment assessment & reporting
    B.A.S.T.I.O.N.         — Active Directory & identity management
    L.A.N.T.E.R.N.         — Network discovery & asset inventory
    T.H.R.E.S.H.O.L.D.     — Disk & storage health monitoring
    V.A.U.L.T.             — M365 license & mailbox auditing
    S.E.N.T.I.N.E.L.       — Service & scheduled task monitoring
    R.E.L.I.C.             — Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.           — Toolkit setup & configuration wizard

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
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

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
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

$Tools = @(
    # ── Deployment & Onboarding ──────────────────────────────────────
    [PSCustomObject]@{
        Key         = '1'
        Name        = 'C.O.V.E.N.A.N.T.'
        File        = 'covenant.ps1'
        Version     = '1.0'
        Description = 'Machine onboarding, Entra ID domain join, and new device setup'
        Color       = 'Blue'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '2'
        Name        = 'C.O.N.J.U.R.E.'
        File        = 'conjure.ps1'
        Version     = '1.0'
        Description = 'Software deployment via Windows Package Manager or Chocolatey'
        Color       = 'Magenta'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '3'
        Name        = 'R.U.N.E.P.R.E.S.S.'
        File        = 'runepress.ps1'
        Version     = '1.0'
        Description = 'Printer driver installation and network printer configuration'
        Color       = 'Cyan'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '4'
        Name        = 'F.O.R.G.E.'
        File        = 'forge.ps1'
        Version     = '1.0'
        Description = 'Driver detection & installation — problem devices, Windows Update, local packages'
        Color       = 'Yellow'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '5'
        Name        = 'R.E.S.T.O.R.A.T.I.O.N.'
        File        = 'restoration.ps1'
        Version     = '1.0'
        Description = 'Automated Windows Update management and maintenance'
        Color       = 'Green'
        Category    = 'Deployment & Onboarding'
    },
    [PSCustomObject]@{
        Key         = '6'
        Name        = 'H.E.A.R.T.H.'
        File        = 'hearth.ps1'
        Version     = '1.0'
        Description = 'Toolkit setup wizard — configure org name, log paths, and default values'
        Color       = 'White'
        Category    = 'Deployment & Onboarding'
    },
    # ── Diagnostics & Reporting ──────────────────────────────────────
    [PSCustomObject]@{
        Key         = '7'
        Name        = 'O.R.A.C.L.E.'
        File        = 'oracle.ps1'
        Version     = '1.0'
        Description = 'System diagnostics, health assessment, and HTML report generation'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '8'
        Name        = 'W.A.R.D.'
        File        = 'ward.ps1'
        Version     = '1.0'
        Description = 'User account audit — roles, last logon, flags, HTML report'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '9'
        Name        = 'T.H.R.E.S.H.O.L.D.'
        File        = 'threshold.ps1'
        Version     = '1.0'
        Description = 'Disk & storage health — physical disk status, volume space, cleanup, old profiles'
        Color       = 'Yellow'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '10'
        Name        = 'S.E.N.T.I.N.E.L.'
        File        = 'sentinel.ps1'
        Version     = '1.0'
        Description = 'Service & task monitor — critical services, scheduled tasks, event log errors'
        Color       = 'Red'
        Category    = 'Diagnostics & Reporting'
    },
    # ── Security ─────────────────────────────────────────────────────
    [PSCustomObject]@{
        Key         = '11'
        Name        = 'C.I.P.H.E.R.'
        File        = 'cipher.ps1'
        Version     = '1.0'
        Description = 'BitLocker drive encryption — enable, disable, backup keys'
        Color       = 'Green'
        Category    = 'Security'
    },
    [PSCustomObject]@{
        Key         = '12'
        Name        = 'S.I.G.I.L.'
        File        = 'sigil.ps1'
        Version     = '1.0'
        Description = 'Security baseline enforcement — telemetry, UAC, firewall, audit policy'
        Color       = 'Red'
        Category    = 'Security'
    },
    [PSCustomObject]@{
        Key         = '13'
        Name        = 'B.A.S.T.I.O.N.'
        File        = 'bastion.ps1'
        Version     = '1.0'
        Description = 'Active Directory management — search, unlock, reset passwords, group membership'
        Color       = 'Blue'
        Category    = 'Security'
    },
    [PSCustomObject]@{
        Key         = '14'
        Name        = 'R.E.L.I.C.'
        File        = 'relic.ps1'
        Version     = '1.0'
        Description = 'Certificate health monitor — local cert stores, SSL/TLS expiry, HTML report'
        Color       = 'Yellow'
        Category    = 'Security'
    },
    # ── Network & Remote ─────────────────────────────────────────────
    [PSCustomObject]@{
        Key         = '15'
        Name        = 'L.E.Y.L.I.N.E.'
        File        = 'leyline.ps1'
        Version     = '1.0'
        Description = 'Network diagnostics & remediation — adapters, ping, DNS, port tests'
        Color       = 'Cyan'
        Category    = 'Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '16'
        Name        = 'S.P.E.C.T.E.R.'
        File        = 'specter.ps1'
        Version     = '1.0'
        Description = 'Remote execution via WinRM — run toolkit tools on a remote machine'
        Color       = 'White'
        Category    = 'Network & Remote'
    },
    [PSCustomObject]@{
        Key         = '17'
        Name        = 'L.A.N.T.E.R.N.'
        File        = 'lantern.ps1'
        Version     = '1.0'
        Description = 'Network discovery & asset inventory — subnet sweep, DNS, MAC, port scan'
        Color       = 'Cyan'
        Category    = 'Network & Remote'
    },
    # ── Cloud & Identity ─────────────────────────────────────────────
    [PSCustomObject]@{
        Key         = '18'
        Name        = 'A.E.G.I.S.'
        File        = 'aegis.ps1'
        Version     = '1.0'
        Description = 'Azure environment assessment — security posture, RBAC, backup coverage, HTML report'
        Color       = 'Cyan'
        Category    = 'Cloud & Identity'
    },
    [PSCustomObject]@{
        Key         = '19'
        Name        = 'V.A.U.L.T.'
        File        = 'vault.ps1'
        Version     = '1.0'
        Description = 'M365 license & mailbox audit — SKU inventory, unlicensed users, MFA status'
        Color       = 'Green'
        Category    = 'Cloud & Identity'
    },
    # ── Data & Migration ─────────────────────────────────────────────
    [PSCustomObject]@{
        Key         = '20'
        Name        = 'P.H.A.N.T.O.M.'
        File        = 'phantom.ps1'
        Version     = '1.0'
        Description = 'Profile migration and data transfer to a new machine'
        Color       = 'Cyan'
        Category    = 'Data & Migration'
    },
    [PSCustomObject]@{
        Key         = '21'
        Name        = 'A.R.C.H.I.V.E.'
        File        = 'archive.ps1'
        Version     = '1.0'
        Description = 'Pre-reimaging profile backup — ZIP to local or network share'
        Color       = 'Magenta'
        Category    = 'Data & Migration'
    }
)

# ===========================
# BANNER
# ===========================

function Show-Banner {
    Clear-Host
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
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    $toolCount = $Tools.Count
    Write-Host "  Technician Toolkit  |  Hub v1.0  |  $toolCount tools  |  Run as Administrator" -ForegroundColor $ColorSchema.Info
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    if ($WhatIf) {
        Write-Host ""
        Write-Host "  *** DRY RUN MODE — tools that support -WhatIf will preview actions only ***" -ForegroundColor Cyan
    }
    Write-Host ""
}

# ===========================
# MENU
# ===========================

function Show-Menu {
    Write-Host "  Select a tool to launch:" -ForegroundColor $ColorSchema.Header
    Write-Host ""

    foreach ($category in $CategoryOrder) {
        $categoryTools = $Tools | Where-Object { $_.Category -eq $category }
        if (-not $categoryTools) { continue }

        Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
        Write-Host "  $category" -ForegroundColor $ColorSchema.Header
        Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
        Write-Host ""

        foreach ($tool in $categoryTools) {
            Write-Host "  [$($tool.Key)]  $($tool.Name)  " -NoNewline -ForegroundColor $tool.Color
            Write-Host "v$($tool.Version)" -ForegroundColor $ColorSchema.Info
            Write-Host "       $($tool.Description)" -ForegroundColor $ColorSchema.Info
            Write-Host ""
        }
    }

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  [Q]  Exit GRIMOIRE" -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
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
                Write-Host "  [!!] $($Tool.File) failed syntax validation after download — file removed." -ForegroundColor $ColorSchema.Error
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
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  $($Tool.Name) has finished. Returning to GRIMOIRE..." -ForegroundColor $ColorSchema.Accent
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
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

do {
    Show-Banner
    Show-Menu

    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Menu
    $Selection = (Read-Host).Trim().ToUpper()

    $MatchedTool = $Tools | Where-Object { $_.Key -eq $Selection }

    if ($MatchedTool) {
        Invoke-Tool -Tool $MatchedTool
    }
    elseif ($Selection -eq 'Q') {
        Clear-Host
        Write-Host ""
        Write-Host "  Closing GRIMOIRE. Stay arcane." -ForegroundColor $ColorSchema.Header
        Write-Host ""
        break
    }
    else {
        Write-Host ""
        Write-Host "  [!!] Invalid selection. Enter a tool number or Q to quit." -ForegroundColor $ColorSchema.Warning
        Start-Sleep -Seconds 1
    }

} while ($true)

# ===========================
# CLEANUP
# ===========================

foreach ($f in $DownloadedFiles) {
    Remove-Item -Path $f -Force -ErrorAction SilentlyContinue
}
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
