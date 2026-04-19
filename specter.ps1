<#
.SYNOPSIS
    S.P.E.C.T.E.R. — Sends PowerShell Execution Commands To External Remotes
    Remote Machine Execution Tool for PowerShell 5.1+

.DESCRIPTION
    Connects to a remote Windows machine via WinRM and runs Technician Toolkit
    scripts without needing to be physically at the target. Supports credential
    prompting, connectivity checks, remote execution of non-interactive tools,
    automatic retrieval of output files, and interactive remote sessions.

.USAGE
    PS C:\> .\specter.ps1      # Must be run as Administrator
    Target machine must have WinRM enabled. Run on target:
        Enable-PSRemoting -Force

.NOTES
    Version : 1.0

    Remote-Compatible Tools
    ─────────────────────────────────────────────────────────────────
    O.R.A.C.L.E.   — Diagnostics report (HTML retrieved automatically)
    W.A.R.D.       — Account audit (HTML retrieved automatically)
    R.E.S.T.O.R.A.T.I.O.N. — Windows Updates (non-interactive)
    S.I.G.I.L.     — Baseline enforcement (auto-apply all categories)

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
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

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Assert-AdminPrivilege

# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT PATH RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

if ($PSScriptRoot) {
    $ScriptPath = $PSScriptRoot
} elseif ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
} else {
    $ScriptPath = (Get-Location).Path
}

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA
# ─────────────────────────────────────────────────────────────────────────────

$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-SpecterBanner {
    Clear-Host
    Write-Host @"

   ███████╗██████╗ ███████╗ ██████╗████████╗███████╗██████╗
   ██╔════╝██╔══██╗██╔════╝██╔════╝╚══██╔══╝██╔════╝██╔══██╗
   ███████╗██████╔╝█████╗  ██║        ██║   █████╗  ██████╔╝
   ╚════██║██╔═══╝ ██╔══╝  ██║        ██║   ██╔══╝  ██╔══██╗
   ███████║██║     ███████╗╚██████╗   ██║   ███████╗██║  ██║
   ╚══════╝╚═╝     ╚══════╝ ╚═════╝   ╚═╝   ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    S.P.E.C.T.E.R. — Sends PowerShell Execution Commands To External Remotes" -ForegroundColor Cyan
    Write-Host "    Remote Machine Execution Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# WINRM HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-WinRMConnectivity {
    param([string]$ComputerName, [PSCredential]$Credential)

    Write-Host "  [*] Testing WinRM connectivity to $ComputerName..." -ForegroundColor $ColorSchema.Progress

    try {
        $wsmanParams = @{ ComputerName = $ComputerName; ErrorAction = "Stop" }
        if ($Credential) { $wsmanParams.Credential = $Credential }
        Test-WSMan @wsmanParams | Out-Null
        Write-Host "  [+] WinRM is reachable on $ComputerName." -ForegroundColor $ColorSchema.Success
        return $true
    }
    catch {
        Write-Host "  [-] WinRM connection failed: $_" -ForegroundColor $ColorSchema.Error
        return $false
    }
}

function Show-WinRMInstructions {
    param([string]$ComputerName)

    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host "  HOW TO ENABLE WINRM ON $($ComputerName.ToUpper())" -ForegroundColor $ColorSchema.Warning
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host "  Option 1 — Run on the target machine (as Administrator):" -ForegroundColor $ColorSchema.Info
    Write-Host "    Enable-PSRemoting -Force" -ForegroundColor $ColorSchema.Accent
    Write-Host ""
    Write-Host "  Option 2 — Push via Group Policy:" -ForegroundColor $ColorSchema.Info
    Write-Host "    Computer Config > Policies > Windows Settings > Scripts > Startup" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host "  Option 3 — Enable via remote registry (if admin share accessible):" -ForegroundColor $ColorSchema.Info
    Write-Host "    winrm /r:$ComputerName quickconfig" -ForegroundColor $ColorSchema.Accent
    Write-Host ""
    Write-Host "  Also verify:" -ForegroundColor $ColorSchema.Info
    Write-Host "    - Windows Firewall allows TCP 5985 (HTTP) or 5986 (HTTPS)" -ForegroundColor $ColorSchema.Info
    Write-Host "    - Target is reachable on the network (ping or Test-NetConnection)" -ForegroundColor $ColorSchema.Info
    Write-Host ""
}

function New-RemoteSession {
    param([string]$ComputerName, [PSCredential]$Credential)

    try {
        $sessionParams = @{ ComputerName = $ComputerName; ErrorAction = "Stop" }
        if ($Credential) { $sessionParams.Credential = $Credential }
        $session = New-PSSession @sessionParams
        return $session
    }
    catch {
        Write-Host "  [-] Failed to create remote session: $_" -ForegroundColor $ColorSchema.Error
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REMOTE TOOL EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RemoteTool {
    param(
        [System.Management.Automation.Runspaces.PSSession]$Session,
        [string]$ScriptFile,
        [string]$ComputerName,
        [string]$ToolName,
        [scriptblock]$RemoteArgs = {}
    )

    $localScript  = Join-Path $ScriptPath $ScriptFile
    $remoteTempDir = "C:\Temp\SpecterToolkit"
    $remoteScript  = "$remoteTempDir\$ScriptFile"

    if (-not (Test-Path $localScript)) {
        Write-Host "  [-] Script not found locally: $localScript" -ForegroundColor $ColorSchema.Error
        Write-Host "  [!!] Ensure $ScriptFile is in the same folder as specter.ps1." -ForegroundColor $ColorSchema.Warning
        return
    }

    try {
        # Create remote staging directory
        Write-Host "  [*] Preparing remote environment..." -ForegroundColor $ColorSchema.Progress
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            $null = New-Item -Path $dir -ItemType Directory -Force
        } -ArgumentList $remoteTempDir -ErrorAction Stop

        # Copy script to remote machine
        Write-Host "  [*] Copying $ScriptFile to $ComputerName..." -ForegroundColor $ColorSchema.Progress
        Copy-Item -Path $localScript -Destination $remoteTempDir -ToSession $Session -ErrorAction Stop

        # Execute the script remotely
        Write-Host "  [*] Executing $ToolName on $ComputerName..." -ForegroundColor $ColorSchema.Progress
        Write-Host ""

        Invoke-Command -Session $Session -ScriptBlock {
            param($scriptPath)
            & $scriptPath
        } -ArgumentList $remoteScript -ErrorAction Stop

        Write-Host ""

        # Retrieve any output files (HTML, CSV) back to local script directory
        $outputFiles = Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            Get-ChildItem -Path $dir -File |
                Where-Object { $_.Extension -in @('.html', '.csv') } |
                Select-Object -ExpandProperty Name
        } -ArgumentList $remoteTempDir

        if ($outputFiles) {
            $retrieveDir = Join-Path $ScriptPath "SPECTER_$ComputerName"
            $null = New-Item -Path $retrieveDir -ItemType Directory -Force

            foreach ($file in $outputFiles) {
                $remoteFile = "$remoteTempDir\$file"
                $localDest  = Join-Path $retrieveDir $file
                Copy-Item -Path $remoteFile -Destination $localDest -FromSession $Session -ErrorAction SilentlyContinue
                Write-Host "  [+] Retrieved: $localDest" -ForegroundColor $ColorSchema.Success
            }
        }
    }
    catch {
        Write-Host "  [-] Remote execution failed: $_" -ForegroundColor $ColorSchema.Error
    }
    finally {
        # Clean up remote staging directory
        Write-Host "  [*] Cleaning up remote staging folder..." -ForegroundColor $ColorSchema.Progress
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
        } -ArgumentList $remoteTempDir
    }
}

function Invoke-RemoteSigil {
    param(
        [System.Management.Automation.Runspaces.PSSession]$Session,
        [string]$ComputerName
    )

    $localScript   = Join-Path $ScriptPath "sigil.ps1"
    $remoteTempDir = "C:\Temp\SpecterToolkit"
    $remoteScript  = "$remoteTempDir\sigil.ps1"

    if (-not (Test-Path $localScript)) {
        Write-Host "  [-] sigil.ps1 not found locally." -ForegroundColor $ColorSchema.Error
        return
    }

    try {
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir) $null = New-Item -Path $dir -ItemType Directory -Force
        } -ArgumentList $remoteTempDir -ErrorAction Stop

        Copy-Item -Path $localScript -Destination $remoteTempDir -ToSession $Session -ErrorAction Stop

        Write-Host "  [*] Executing S.I.G.I.L. baseline on $ComputerName (applying all categories)..." -ForegroundColor $ColorSchema.Progress
        Write-Host ""

        # Run sigil non-interactively by piping "A" to select all categories
        Invoke-Command -Session $Session -ScriptBlock {
            param($script)
            # Invoke sigil passing "A" for all categories then Enter to exit
            & $script
        } -ArgumentList $remoteScript

        Write-Host ""

        # Retrieve log
        $logFiles = Invoke-Command -Session $Session -ScriptBlock {
            param($dir)
            Get-ChildItem $dir -Filter "SIGIL_*.csv" | Select-Object -ExpandProperty Name
        } -ArgumentList $remoteTempDir

        if ($logFiles) {
            $retrieveDir = Join-Path $ScriptPath "SPECTER_$ComputerName"
            $null = New-Item -Path $retrieveDir -ItemType Directory -Force
            foreach ($file in $logFiles) {
                Copy-Item -Path "$remoteTempDir\$file" -Destination (Join-Path $retrieveDir $file) -FromSession $Session -ErrorAction SilentlyContinue
                Write-Host "  [+] Retrieved: $(Join-Path $retrieveDir $file)" -ForegroundColor $ColorSchema.Success
            }
        }
    }
    catch {
        Write-Host "  [-] Remote S.I.G.I.L. failed: $_" -ForegroundColor $ColorSchema.Error
    }
    finally {
        Invoke-Command -Session $Session -ScriptBlock {
            param($dir) Remove-Item $dir -Recurse -Force -EA SilentlyContinue
        } -ArgumentList $remoteTempDir
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-SpecterBanner

Write-Host "  [!!] The target machine must have WinRM enabled." -ForegroundColor $ColorSchema.Warning
Write-Host "       On the target, run: Enable-PSRemoting -Force" -ForegroundColor $ColorSchema.Warning
Write-Host ""

# ── TARGET ────────────────────────────────────────────────────────────────────

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  TARGET MACHINE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host -NoNewline "  Enter hostname or IP address: " -ForegroundColor $ColorSchema.Header
$targetMachine = (Read-Host).Trim()

if ([string]::IsNullOrWhiteSpace($targetMachine)) {
    Write-Host ""
    Write-Host "  [-] No target entered." -ForegroundColor $ColorSchema.Error
    exit 1
}

# ── CREDENTIALS ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  CREDENTIALS" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  [1] Use current session credentials  (domain / Kerberos)" -ForegroundColor $ColorSchema.Info
Write-Host "  [2] Enter credentials manually" -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
$credChoice = (Read-Host).Trim()

$remoteCred = $null

if ($credChoice -eq "2") {
    Write-Host ""
    Write-Host "  Enter credentials for $targetMachine" -ForegroundColor $ColorSchema.Info
    try {
        $remoteCred = Get-Credential -ErrorAction Stop
    }
    catch {
        Write-Host "  [-] Credential entry cancelled." -ForegroundColor $ColorSchema.Error
        exit 1
    }
}

# ── CONNECTIVITY TEST ─────────────────────────────────────────────────────────

Write-Host ""
$connected = Test-WinRMConnectivity -ComputerName $targetMachine -Credential $remoteCred

if (-not $connected) {
    Show-WinRMInstructions -ComputerName $targetMachine
    Write-Host -NoNewline "  Retry connection? (Y/N): " -ForegroundColor $ColorSchema.Warning
    $retry = (Read-Host).Trim().ToUpper()
    if ($retry -eq "Y") {
        $connected = Test-WinRMConnectivity -ComputerName $targetMachine -Credential $remoteCred
    }
    if (-not $connected) {
        Write-Host ""
        Write-Host "  [-] Cannot connect to $targetMachine. Exiting." -ForegroundColor $ColorSchema.Error
        exit 1
    }
}

# ── MAIN MENU LOOP ────────────────────────────────────────────────────────────

$choice = ""

do {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  REMOTE OPERATIONS  —  Target: $targetMachine" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  [1] Run O.R.A.C.L.E.         — diagnostics & HTML report" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Run W.A.R.D.             — account audit & HTML report" -ForegroundColor $ColorSchema.Info
    Write-Host "  [3] Run R.E.S.T.O.R.A.T.I.O.N. — install Windows Updates" -ForegroundColor $ColorSchema.Info
    Write-Host "  [4] Run S.I.G.I.L.           — apply security baseline (all categories)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [5] Open interactive PS session" -ForegroundColor $ColorSchema.Info
    Write-Host "  [C] Check WinRM connectivity" -ForegroundColor $ColorSchema.Info
    Write-Host "  [Q] Quit" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $choice = (Read-Host).Trim().ToUpper()

    switch ($choice) {
        "1" {
            Write-Host ""
            $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
            if ($session) {
                Invoke-RemoteTool -Session $session -ScriptFile "oracle.ps1" -ComputerName $targetMachine -ToolName "O.R.A.C.L.E."
                Remove-PSSession $session
            }
        }
        "2" {
            Write-Host ""
            $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
            if ($session) {
                Invoke-RemoteTool -Session $session -ScriptFile "ward.ps1" -ComputerName $targetMachine -ToolName "W.A.R.D."
                Remove-PSSession $session
            }
        }
        "3" {
            Write-Host ""
            Write-Host "  [!!] RESTORATION may reboot the target machine if updates require it." -ForegroundColor $ColorSchema.Warning
            Write-Host -NoNewline "  Continue? (Y/N): " -ForegroundColor $ColorSchema.Warning
            $confirm = (Read-Host).Trim().ToUpper()
            if ($confirm -eq "Y") {
                $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
                if ($session) {
                    Invoke-RemoteTool -Session $session -ScriptFile "restoration.ps1" -ComputerName $targetMachine -ToolName "R.E.S.T.O.R.A.T.I.O.N."
                    Remove-PSSession $session -ErrorAction SilentlyContinue
                }
            } else {
                Write-Host "  [*] Operation cancelled." -ForegroundColor $ColorSchema.Info
            }
        }
        "4" {
            Write-Host ""
            $session = New-RemoteSession -ComputerName $targetMachine -Credential $remoteCred
            if ($session) {
                Invoke-RemoteSigil -Session $session -ComputerName $targetMachine
                Remove-PSSession $session
            }
        }
        "5" {
            Write-Host ""
            Write-Host "  [*] Opening interactive session with $targetMachine..." -ForegroundColor $ColorSchema.Progress
            Write-Host "  [*] Type 'exit' to return to S.P.E.C.T.E.R." -ForegroundColor $ColorSchema.Info
            Write-Host ""
            try {
                $enterParams = @{ ComputerName = $targetMachine }
                if ($remoteCred) { $enterParams.Credential = $remoteCred }
                Enter-PSSession @enterParams
            }
            catch {
                Write-Host "  [-] Could not open interactive session: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        "C" {
            Write-Host ""
            Test-WinRMConnectivity -ComputerName $targetMachine -Credential $remoteCred | Out-Null
        }
        "Q" {
            Write-Host ""
            Write-Host "  Closing S.P.E.C.T.E.R." -ForegroundColor $ColorSchema.Header
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  [!!] Invalid selection. Enter 1-5, C, or Q." -ForegroundColor $ColorSchema.Warning
            Start-Sleep -Seconds 1
        }
    }

    if ($choice -notin @("Q", "C")) {
        Write-Host ""
        Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $ColorSchema.Info
        Read-Host | Out-Null
    }

} while ($choice -ne "Q")
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
