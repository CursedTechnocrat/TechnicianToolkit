<#
.SYNOPSIS
    S.I.G.I.L. — Secures Infrastructure: Governs via Integrated Lockdown
    Security Baseline Enforcement Tool for PowerShell 5.1+

.DESCRIPTION
    Applies a standardized security and configuration baseline to a Windows machine.
    Covers telemetry, screensaver lock, UAC, autorun, firewall, account policy,
    password policy, Remote Desktop, audit policy, legacy protocol hardening
    (SMBv1, LLMNR, NetBIOS), and credential protection (LSA PPL, NoLMHash,
    RDP Restricted Admin). Pairs naturally with C.O.V.E.N.A.N.T. as a
    post-onboarding hardening step. Changes are logged to a timestamped CSV.

.USAGE
    PS C:\> .\sigil.ps1                              # Must be run as Administrator
    PS C:\> .\sigil.ps1 -Unattended                  # Apply all categories silently (default)
    PS C:\> .\sigil.ps1 -Unattended -Categories "1,3,5"  # Apply specific categories silently

.NOTES
    Version : 1.0

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

param(
    [switch]$Unattended,
    [string]$Categories = "A"
)

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
# ACTION LOG
# ─────────────────────────────────────────────────────────────────────────────

$ActionLog = New-Object System.Collections.ArrayList

function Add-ActionRecord {
    param(
        [string]$Category,
        [string]$Setting,
        [string]$Status,
        [string]$Detail    = "",
        [string]$Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    )
    [void]$ActionLog.Add([PSCustomObject]@{
        Timestamp = $Timestamp
        Category  = $Category
        Setting   = $Setting
        Status    = $Status
        Detail    = $Detail
    })
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-SigilBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ███████╗██╗ ██████╗ ██╗██╗
   ██╔════╝██║██╔════╝ ██║██║
   ███████╗██║██║  ███╗██║██║
   ╚════██║██║██║   ██║██║██║
   ███████║██║╚██████╔╝██║███████╗
   ╚══════╝╚═╝ ╚═════╝ ╚═╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    S.I.G.I.L. — Secures Infrastructure: Governs via Integrated Lockdown" -ForegroundColor Cyan
    Write-Host "    Security Baseline & Policy Enforcement Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# BASELINE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Set-BaselineReg {
    param(
        [string]$Path,
        [string]$Name,
        $Value,
        [string]$Type     = "DWord",
        [string]$Category,
        [string]$Label
    )

    try {
        $current = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue).$Name

        if ($null -ne $current -and $current -eq $Value) {
            Write-Host "    [OK] $Label — already set ($Value)." -ForegroundColor $ColorSchema.Info
            Add-ActionRecord -Category $Category -Setting $Label -Status "Already Set" -Detail "Value: $Value"
            return
        }

        if (-not (Test-Path $Path)) {
            $null = New-Item -Path $Path -Force -ErrorAction Stop
        }

        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
        $prev = if ($null -ne $current) { $current } else { "(not set)" }
        Write-Host "    [+] $Label — applied.  ($prev → $Value)" -ForegroundColor $ColorSchema.Success
        Add-ActionRecord -Category $Category -Setting $Label -Status "Applied" -Detail "$prev → $Value"
    }
    catch {
        Write-Host "    [-] $Label — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Category $Category -Setting $Label -Status "Failed" -Detail $_
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BASELINE CATEGORIES
# ─────────────────────────────────────────────────────────────────────────────

function Apply-Telemetry {
    Write-Host "  [*] Applying telemetry & privacy settings..." -ForegroundColor $ColorSchema.Progress

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" `
        -Name     "AllowTelemetry" `
        -Value    1 `
        -Category "Telemetry" `
        -Label    "Windows Telemetry (set to Security/minimal)"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" `
        -Name     "DisableEnterpriseAuthProxy" `
        -Value    1 `
        -Category "Telemetry" `
        -Label    "Disable enterprise auth proxy for telemetry"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" `
        -Name     "DisabledByGroupPolicy" `
        -Value    1 `
        -Category "Telemetry" `
        -Label    "Disable advertising ID"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization" `
        -Name     "RestrictImplicitInkCollection" `
        -Value    1 `
        -Category "Telemetry" `
        -Label    "Restrict ink & typing personalization"

    Write-Host ""
}

function Apply-ScreensaverLock {
    param([int]$TimeoutSeconds = 600)

    Write-Host "  [*] Applying screensaver & display lock settings..." -ForegroundColor $ColorSchema.Progress
    Write-Host "  [!!] Screensaver settings apply to the currently logged-on user profile." -ForegroundColor $ColorSchema.Warning

    Set-BaselineReg `
        -Path     "HKCU:\Control Panel\Desktop" `
        -Name     "ScreenSaveActive" `
        -Value    "1" `
        -Type     "String" `
        -Category "Screensaver" `
        -Label    "Enable screensaver"

    Set-BaselineReg `
        -Path     "HKCU:\Control Panel\Desktop" `
        -Name     "ScreenSaverIsSecure" `
        -Value    "1" `
        -Type     "String" `
        -Category "Screensaver" `
        -Label    "Require password on screensaver resume"

    Set-BaselineReg `
        -Path     "HKCU:\Control Panel\Desktop" `
        -Name     "ScreenSaveTimeOut" `
        -Value    "$TimeoutSeconds" `
        -Type     "String" `
        -Category "Screensaver" `
        -Label    "Screensaver timeout ($($TimeoutSeconds / 60) min)"

    # Also enforce lock via machine policy (applies system-wide)
    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
        -Name     "InactivityTimeoutSecs" `
        -Value    $TimeoutSeconds `
        -Category "Screensaver" `
        -Label    "Machine inactivity lock timeout ($($TimeoutSeconds / 60) min)"

    Write-Host ""
}

function Apply-UAC {
    Write-Host "  [*] Applying UAC settings..." -ForegroundColor $ColorSchema.Progress

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
        -Name     "EnableLUA" `
        -Value    1 `
        -Category "UAC" `
        -Label    "Enable UAC (User Account Control)"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
        -Name     "ConsentPromptBehaviorAdmin" `
        -Value    2 `
        -Category "UAC" `
        -Label    "UAC prompt for admins (Always Notify)"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
        -Name     "ConsentPromptBehaviorUser" `
        -Value    3 `
        -Category "UAC" `
        -Label    "UAC prompt for standard users (require credentials)"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
        -Name     "PromptOnSecureDesktop" `
        -Value    1 `
        -Category "UAC" `
        -Label    "Show UAC prompt on secure desktop"

    Write-Host ""
}

function Apply-Autorun {
    Write-Host "  [*] Applying autorun & autoplay settings..." -ForegroundColor $ColorSchema.Progress

    # 255 = disable for all drive types
    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" `
        -Name     "NoDriveTypeAutoRun" `
        -Value    255 `
        -Category "Autorun" `
        -Label    "Disable AutoRun for all drives (machine)"

    Set-BaselineReg `
        -Path     "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" `
        -Name     "NoDriveTypeAutoRun" `
        -Value    255 `
        -Category "Autorun" `
        -Label    "Disable AutoRun for all drives (user)"

    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" `
        -Name     "NoAutoplayfornonVolume" `
        -Value    1 `
        -Category "Autorun" `
        -Label    "Disable AutoPlay for non-volume devices"

    Write-Host ""
}

function Apply-Firewall {
    Write-Host "  [*] Applying Windows Firewall settings..." -ForegroundColor $ColorSchema.Progress

    try {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True -ErrorAction Stop
        Write-Host "    [+] Windows Firewall enabled on all profiles." -ForegroundColor $ColorSchema.Success
        Add-ActionRecord -Category "Firewall" -Setting "Enable all firewall profiles" -Status "Applied"
    }
    catch {
        Write-Host "    [-] Failed to configure firewall: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Category "Firewall" -Setting "Enable all firewall profiles" -Status "Failed" -Detail $_
    }

    # Block inbound by default on Public profile
    try {
        Set-NetFirewallProfile -Profile Public -DefaultInboundAction Block -ErrorAction Stop
        Write-Host "    [+] Public profile — inbound connections blocked by default." -ForegroundColor $ColorSchema.Success
        Add-ActionRecord -Category "Firewall" -Setting "Block inbound on Public profile" -Status "Applied"
    }
    catch {
        Write-Host "    [-] Failed to set Public profile inbound policy: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Category "Firewall" -Setting "Block inbound on Public profile" -Status "Failed" -Detail $_
    }

    Write-Host ""
}

function Apply-GuestAccount {
    Write-Host "  [*] Applying guest account settings..." -ForegroundColor $ColorSchema.Progress

    try {
        $guest = Get-LocalUser -Name "Guest" -ErrorAction SilentlyContinue
        if ($guest) {
            if ($guest.Enabled) {
                Disable-LocalUser -Name "Guest" -ErrorAction Stop
                Write-Host "    [+] Guest account disabled." -ForegroundColor $ColorSchema.Success
                Add-ActionRecord -Category "Accounts" -Setting "Disable Guest account" -Status "Applied"
            } else {
                Write-Host "    [OK] Guest account already disabled." -ForegroundColor $ColorSchema.Info
                Add-ActionRecord -Category "Accounts" -Setting "Disable Guest account" -Status "Already Set"
            }
        } else {
            Write-Host "    [OK] Guest account not present." -ForegroundColor $ColorSchema.Info
            Add-ActionRecord -Category "Accounts" -Setting "Disable Guest account" -Status "Not Present"
        }
    }
    catch {
        Write-Host "    [-] Failed to disable Guest account: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Category "Accounts" -Setting "Disable Guest account" -Status "Failed" -Detail $_
    }

    Write-Host ""
}

function Apply-PasswordPolicy {
    Write-Host "  [*] Applying local password policy..." -ForegroundColor $ColorSchema.Progress
    Write-Host "  [!!] These settings apply to local accounts only. Domain policy takes precedence." -ForegroundColor $ColorSchema.Warning

    $policies = @(
        @{ Args = "/minpwlen:8";      Label = "Minimum password length (8 characters)" },
        @{ Args = "/maxpwage:90";     Label = "Maximum password age (90 days)" },
        @{ Args = "/minpwage:1";      Label = "Minimum password age (1 day)" },
        @{ Args = "/uniquepw:5";      Label = "Password history (remember 5)" },
        @{ Args = "/lockoutthreshold:5"; Label = "Account lockout threshold (5 attempts)" },
        @{ Args = "/lockoutduration:30"; Label = "Account lockout duration (30 minutes)" }
    )

    foreach ($policy in $policies) {
        try {
            & net accounts $policy.Args.Split(' ') 2>&1 | Out-Null
            Write-Host "    [+] $($policy.Label) — applied." -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Category "Password Policy" -Setting $policy.Label -Status "Applied"
        }
        catch {
            Write-Host "    [-] $($policy.Label) — failed: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Category "Password Policy" -Setting $policy.Label -Status "Failed" -Detail $_
        }
    }

    Write-Host ""
}

function Apply-RemoteDesktop {
    Write-Host "  [*] Configuring Remote Desktop..." -ForegroundColor $ColorSchema.Progress
    Write-Host ""
    Write-Host "  [1] Enable RDP (for remote tech access)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Disable RDP (recommended for end-user workstations)" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    if (-not $Unattended) {
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
        $rdpChoice = (Read-Host).Trim()
    } else {
        $rdpChoice = "2"
    }

    $rdpValue = if ($rdpChoice -eq "1") { 0 } else { 1 }
    $rdpLabel = if ($rdpChoice -eq "1") { "Enable" } else { "Disable" }

    Set-BaselineReg `
        -Path     "HKLM:\System\CurrentControlSet\Control\Terminal Server" `
        -Name     "fDenyTSConnections" `
        -Value    $rdpValue `
        -Category "Remote Desktop" `
        -Label    "$rdpLabel Remote Desktop"

    if ($rdpChoice -eq "1") {
        # Require Network Level Authentication
        Set-BaselineReg `
            -Path     "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" `
            -Name     "UserAuthentication" `
            -Value    1 `
            -Category "Remote Desktop" `
            -Label    "Require Network Level Authentication (NLA)"

        try {
            Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction Stop
            Write-Host "    [+] Remote Desktop firewall rules enabled." -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Category "Remote Desktop" -Setting "Enable RDP firewall rules" -Status "Applied"
        }
        catch {
            Write-Host "    [!!] Could not update RDP firewall rules: $_" -ForegroundColor $ColorSchema.Warning
        }
    } else {
        try {
            Disable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue
            Add-ActionRecord -Category "Remote Desktop" -Setting "Disable RDP firewall rules" -Status "Applied"
        }
        catch { }
    }

    Write-Host ""
}

function Apply-AuditPolicy {
    Write-Host "  [*] Applying audit policy..." -ForegroundColor $ColorSchema.Progress

    $auditSettings = @(
        @{ Sub = "Logon";             Args = '/subcategory:"Logon" /success:enable /failure:enable';                     Label = "Logon events" },
        @{ Sub = "Logoff";            Args = '/subcategory:"Logoff" /success:enable';                                    Label = "Logoff events" },
        @{ Sub = "Account Lockout";   Args = '/subcategory:"Account Lockout" /failure:enable';                           Label = "Account lockout failures" },
        @{ Sub = "Audit Policy Change"; Args = '/subcategory:"Audit Policy Change" /success:enable /failure:enable';     Label = "Audit policy changes" },
        @{ Sub = "User Account Management"; Args = '/subcategory:"User Account Management" /success:enable /failure:enable'; Label = "User account management" }
    )

    foreach ($audit in $auditSettings) {
        try {
            $result = & auditpol $audit.Args.Split(' ') 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "    [+] $($audit.Label) — enabled." -ForegroundColor $ColorSchema.Success
                Add-ActionRecord -Category "Audit Policy" -Setting $audit.Label -Status "Applied"
            } else {
                Write-Host "    [!!] $($audit.Label) — may require domain policy override." -ForegroundColor $ColorSchema.Warning
                Add-ActionRecord -Category "Audit Policy" -Setting $audit.Label -Status "Skipped" -Detail "May be overridden by domain policy"
            }
        }
        catch {
            Write-Host "    [-] $($audit.Label) — failed: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Category "Audit Policy" -Setting $audit.Label -Status "Failed" -Detail $_
        }
    }

    Write-Host ""
}

function Apply-WindowsUpdatePolicy {
    Write-Host "  [*] Applying Windows Update behavior..." -ForegroundColor $ColorSchema.Progress

    # Exclude drivers from Windows Update (managed separately)
    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" `
        -Name     "ExcludeWUDriversInQualityUpdate" `
        -Value    1 `
        -Category "Windows Update" `
        -Label    "Exclude driver updates from Windows Update"

    # Disable automatic restart without user consent
    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" `
        -Name     "NoAutoRebootWithLoggedOnUsers" `
        -Value    1 `
        -Category "Windows Update" `
        -Label    "No auto-reboot while users are logged on"

    Write-Host ""
}

function Apply-LegacyProtocols {
    Write-Host "  [*] Disabling legacy network protocols..." -ForegroundColor $ColorSchema.Progress

    # ── SMBv1 ─────────────────────────────────────────────────────────────────
    # Exploited by WannaCry / EternalBlue ransomware; no legitimate modern use.
    try {
        $smb1Enabled = (Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol
        if (-not $smb1Enabled) {
            Write-Host "    [OK] SMBv1 — already disabled." -ForegroundColor $ColorSchema.Info
            Add-ActionRecord -Category "Legacy Protocols" -Setting "Disable SMBv1" -Status "Already Set" -Detail "EnableSMB1Protocol = False"
        } else {
            Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force -ErrorAction Stop
            Write-Host "    [+] SMBv1 — disabled." -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Category "Legacy Protocols" -Setting "Disable SMBv1" -Status "Applied" -Detail "True → False"
        }
    }
    catch {
        Write-Host "    [-] SMBv1 — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Category "Legacy Protocols" -Setting "Disable SMBv1" -Status "Failed" -Detail $_
    }

    # ── LLMNR ─────────────────────────────────────────────────────────────────
    # Link-Local Multicast Name Resolution — abused by Responder to capture hashes.
    Set-BaselineReg `
        -Path     "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient" `
        -Name     "EnableMulticast" `
        -Value    0 `
        -Category "Legacy Protocols" `
        -Label    "Disable LLMNR (Responder attack mitigation)"

    # ── NetBIOS over TCP/IP ────────────────────────────────────────────────────
    # Legacy protocol; disabling reduces attack surface for NBNS poisoning.
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop |
                    Where-Object { $_.IPEnabled }

        $toDisable = $adapters | Where-Object { $_.TcpipNetbiosOptions -ne 2 }

        if ($toDisable.Count -eq 0) {
            Write-Host "    [OK] NetBIOS over TCP/IP — already disabled on all adapters." -ForegroundColor $ColorSchema.Info
            Add-ActionRecord -Category "Legacy Protocols" -Setting "Disable NetBIOS over TCP/IP" -Status "Already Set"
        } else {
            $changed = 0
            foreach ($adapter in $toDisable) {
                $result = $adapter.SetTcpipNetbios(2)   # 2 = Disable NetBIOS over TCP/IP
                if ($result.ReturnValue -eq 0) { $changed++ }
            }
            Write-Host "    [+] NetBIOS over TCP/IP — disabled on $changed adapter(s)." -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Category "Legacy Protocols" -Setting "Disable NetBIOS over TCP/IP" -Status "Applied" -Detail "$changed adapter(s) updated"
        }
    }
    catch {
        Write-Host "    [-] NetBIOS disable — failed: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Category "Legacy Protocols" -Setting "Disable NetBIOS over TCP/IP" -Status "Failed" -Detail $_
    }

    Write-Host ""
}

function Apply-CredentialProtection {
    Write-Host "  [*] Applying credential protection settings..." -ForegroundColor $ColorSchema.Progress

    # ── LSA Protected Process (RunAsPPL) ──────────────────────────────────────
    # Prevents lsass.exe memory dumps by tools like Mimikatz. Requires reboot.
    Set-BaselineReg `
        -Path     "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" `
        -Name     "RunAsPPL" `
        -Value    1 `
        -Category "Credential Protection" `
        -Label    "LSA Protected Process Light (RunAsPPL) — blocks lsass memory dumps"

    # ── No LM Hash ────────────────────────────────────────────────────────────
    # Stops Windows storing weak LAN Manager password hashes alongside NTLM.
    Set-BaselineReg `
        -Path     "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" `
        -Name     "NoLMHash" `
        -Value    1 `
        -Category "Credential Protection" `
        -Label    "No LM Hash storage — prevents weak NTLM hash capture"

    # ── RDP Restricted Admin mode ─────────────────────────────────────────────
    # Prevents credential forwarding when connecting via Remote Desktop.
    # DisableRestrictedAdmin = 0 means Restricted Admin IS enabled (double-negative).
    Set-BaselineReg `
        -Path     "HKLM:\System\CurrentControlSet\Control\Lsa" `
        -Name     "DisableRestrictedAdmin" `
        -Value    0 `
        -Category "Credential Protection" `
        -Label    "RDP Restricted Admin mode — prevents credential exposure over RDP"

    Write-Host "  [!!] NOTE: RunAsPPL requires a reboot to take full effect." -ForegroundColor $ColorSchema.Warning
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-SigilBanner }

Write-Host "  [!!] This script modifies registry keys and local security policy." -ForegroundColor $ColorSchema.Warning
Write-Host "       Domain Group Policy will override local settings where applicable." -ForegroundColor $ColorSchema.Warning
Write-Host ""

$categories = [ordered]@{
    "1"  = @{ Label = "Telemetry & Privacy";        Fn = { Apply-Telemetry } }
    "2"  = @{ Label = "Screensaver & Display Lock"; Fn = { Apply-ScreensaverLock } }
    "3"  = @{ Label = "UAC (User Account Control)"; Fn = { Apply-UAC } }
    "4"  = @{ Label = "Autorun & Autoplay";         Fn = { Apply-Autorun } }
    "5"  = @{ Label = "Windows Firewall";           Fn = { Apply-Firewall } }
    "6"  = @{ Label = "Guest Account";              Fn = { Apply-GuestAccount } }
    "7"  = @{ Label = "Password Policy";            Fn = { Apply-PasswordPolicy } }
    "8"  = @{ Label = "Remote Desktop";             Fn = { Apply-RemoteDesktop } }
    "9"  = @{ Label = "Audit Policy";               Fn = { Apply-AuditPolicy } }
    "10" = @{ Label = "Windows Update Behavior";    Fn = { Apply-WindowsUpdatePolicy } }
    "11" = @{ Label = "Legacy Protocol Hardening";  Fn = { Apply-LegacyProtocols } }
    "12" = @{ Label = "Credential Protection";      Fn = { Apply-CredentialProtection } }
}

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SELECT BASELINE CATEGORIES" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if ($Unattended) {
    Write-Host "  [*] Unattended mode: applying categories — $Categories" -ForegroundColor $ColorSchema.Info
    $rawInput = $Categories.ToUpper()
} else {
    Write-Host "  Enter numbers separated by commas, or A for all." -ForegroundColor $ColorSchema.Info
    Write-Host ""

    foreach ($key in $categories.Keys) {
        Write-Host ("  [{0,2}] {1}" -f $key, $categories[$key].Label) -ForegroundColor $ColorSchema.Info
    }

    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $rawInput = (Read-Host).Trim().ToUpper()
}

$selectedKeys = @()

if ($rawInput -eq "A") {
    $selectedKeys = $categories.Keys
} else {
    $selectedKeys = $rawInput -split ',' |
        ForEach-Object { $_.Trim() } |
        Where-Object   { $categories.ContainsKey($_) }
}

if ($selectedKeys.Count -eq 0) {
    Write-Host ""
    Write-Host "  [-] No valid categories selected." -ForegroundColor $ColorSchema.Error
    exit 1
}

Write-Host ""
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  APPLYING BASELINE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

foreach ($key in $selectedKeys) {
    $cat = $categories[$key]
    Write-Host ("  " + ("─" * 40)) -ForegroundColor $ColorSchema.Header
    Write-Host "  $($cat.Label)" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 40)) -ForegroundColor $ColorSchema.Header
    & $cat.Fn
}

# ── LOG ───────────────────────────────────────────────────────────────────────

$logFile = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "SIGIL_BaselineLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

try {
    $ActionLog | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
    Write-Host "  [+] Log saved: $logFile" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Could not save log: $_" -ForegroundColor $ColorSchema.Error
}

# ── SUMMARY ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  BASELINE SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

$byCategory = $ActionLog | Group-Object -Property Category

foreach ($group in $byCategory) {
    Write-Host "  $($group.Name)" -ForegroundColor $ColorSchema.Header
    foreach ($record in $group.Group) {
        $color = switch ($record.Status) {
            "Applied"     { $ColorSchema.Success }
            "Already Set" { $ColorSchema.Info    }
            "Not Present" { $ColorSchema.Info    }
            "Skipped"     { $ColorSchema.Warning }
            default       { $ColorSchema.Error   }
        }
        $detail = if ($record.Detail) { "  ($($record.Detail))" } else { "" }
        Write-Host ("    {0,-44} [{1}]{2}" -f $record.Setting, $record.Status, $detail) -ForegroundColor $color
    }
    Write-Host ""
}

$applied    = ($ActionLog | Where-Object { $_.Status -eq "Applied"     } | Measure-Object).Count
$alreadySet = ($ActionLog | Where-Object { $_.Status -eq "Already Set" } | Measure-Object).Count
$failed     = ($ActionLog | Where-Object { $_.Status -eq "Failed"      } | Measure-Object).Count

Write-Host "  Applied: $applied  |  Already Set: $alreadySet  |  Failed: $failed" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  S.I.G.I.L. BASELINE COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
