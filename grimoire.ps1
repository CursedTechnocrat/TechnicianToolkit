<#
.SYNOPSIS
    G.R.I.M.O.I.R.E. — General Repository for Integrated Management and Orchestration of IT Resources & Executables
    Technician Toolkit Hub for PowerShell 5.1+

.DESCRIPTION
    Central launcher for the Technician Toolkit. Presents an interactive menu
    to select and run any of the available tools, then returns to the hub on
    completion.

.USAGE
    PS C:\> .\grimoire.ps1      # Must be run as Administrator

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

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

# ===========================
# ADMIN PRIVILEGE CHECK
# ===========================
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "  Restarting with administrator privileges..." -ForegroundColor Yellow
    $PSExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    Start-Process -FilePath $PSExe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

# ===========================
# INITIALIZATION
# ===========================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

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

$Tools = @(
    [PSCustomObject]@{
        Key         = '1'
        Name        = 'R.U.N.E.P.R.E.S.S.'
        File        = 'runepress.ps1'
        Description = 'Printer driver installation and network printer configuration'
        Color       = 'Cyan'
    },
    [PSCustomObject]@{
        Key         = '2'
        Name        = 'R.E.S.T.O.R.A.T.I.O.N.'
        File        = 'restoration.ps1'
        Description = 'Automated Windows Update management and maintenance'
        Color       = 'Green'
    },
    [PSCustomObject]@{
        Key         = '3'
        Name        = 'C.O.N.J.U.R.E.'
        File        = 'conjure.ps1'
        Description = 'Software deployment via Windows Package Manager or Chocolatey'
        Color       = 'Magenta'
    },
    [PSCustomObject]@{
        Key         = '4'
        Name        = 'O.R.A.C.L.E.'
        File        = 'oracle.ps1'
        Description = 'System diagnostics, health assessment, and HTML report generation'
        Color       = 'Yellow'
    },
    [PSCustomObject]@{
        Key         = '5'
        Name        = 'C.O.V.E.N.A.N.T.'
        File        = 'covenant.ps1'
        Description = 'Machine onboarding, Entra ID domain join, and new device setup'
        Color       = 'Blue'
    },
    [PSCustomObject]@{
        Key         = '6'
        Name        = 'P.H.A.N.T.O.M.'
        File        = 'phantom.ps1'
        Description = 'Profile migration and data transfer to a new machine'
        Color       = 'Cyan'
    },
    [PSCustomObject]@{
        Key         = '7'
        Name        = 'C.I.P.H.E.R.'
        File        = 'cipher.ps1'
        Description = 'BitLocker drive encryption — enable, disable, backup keys'
        Color       = 'Green'
    },
    [PSCustomObject]@{
        Key         = '8'
        Name        = 'W.A.R.D.'
        File        = 'ward.ps1'
        Description = 'User account audit — roles, last logon, flags, HTML report'
        Color       = 'Yellow'
    },
    [PSCustomObject]@{
        Key         = '9'
        Name        = 'A.R.C.H.I.V.E.'
        File        = 'archive.ps1'
        Description = 'Pre-reimaging profile backup — ZIP to local or network share'
        Color       = 'Magenta'
    },
    [PSCustomObject]@{
        Key         = '10'
        Name        = 'S.I.G.I.L.'
        File        = 'sigil.ps1'
        Description = 'Security baseline enforcement — telemetry, UAC, firewall, audit policy'
        Color       = 'Red'
    },
    [PSCustomObject]@{
        Key         = '11'
        Name        = 'S.P.E.C.T.E.R.'
        File        = 'specter.ps1'
        Description = 'Remote execution via WinRM — run toolkit tools on a remote machine'
        Color       = 'White'
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
    Write-Host "  Technician Toolkit  |  Hub v1.0  |  Run as Administrator" -ForegroundColor $ColorSchema.Info
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
}

# ===========================
# MENU
# ===========================

function Show-Menu {
    Write-Host "  Select a tool to launch:" -ForegroundColor $ColorSchema.Header
    Write-Host ""

    foreach ($tool in $Tools) {
        $label = "  [$($tool.Key)]  $($tool.Name)"
        Write-Host $label -ForegroundColor $tool.Color -NoNewline
        Write-Host ""
        Write-Host "       $($tool.Description)" -ForegroundColor $ColorSchema.Info
        Write-Host ""
    }

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
            Write-Host "  Downloaded successfully." -ForegroundColor $ColorSchema.Success
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

    try {
        & $ToolPath
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
        Write-Host "  [!!] Invalid selection. Enter 1-11 or Q to quit." -ForegroundColor $ColorSchema.Warning
        Start-Sleep -Seconds 1
    }

} while ($true)
