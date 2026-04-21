<#
.SYNOPSIS
    B.A.S.T.I.O.N. — Bulk Active-directory Stewardship: Tasks, Identity, Operations & Namespacing
    Active Directory & Identity Management Tool for PowerShell 5.1+

.DESCRIPTION
    Interactive tool for IT technicians to manage on-premises Active Directory users and groups.
    Uses the ActiveDirectory PowerShell module (RSAT). Supports user search, account unlock,
    password reset, enable/disable, group membership management, and stale account reporting.
    If RSAT is not installed, the script will detect this and offer to install it automatically.

.USAGE
    PS C:\> .\bastion.ps1                         # Interactive menu — must be run as Administrator
    PS C:\> .\bastion.ps1 -Unattended -Action StaleReport   # Export stale accounts HTML report silently

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
    [switch]$Unattended,
    [ValidateSet('StaleReport','PasswordExpiryReport')]
    [string]$Action = 'StaleReport',
    [switch]$Transcript
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

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA
# ─────────────────────────────────────────────────────────────────────────────

$C = @{
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

function Show-BastionBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗  █████╗ ███████╗████████╗██╗ ██████╗ ███╗   ██╗
  ██╔══██╗██╔══██╗██╔════╝╚══██╔══╝██║██╔═══██╗████╗  ██║
  ██████╔╝███████║███████╗   ██║   ██║██║   ██║██╔██╗ ██║
  ██╔══██╗██╔══██║╚════██║   ██║   ██║██║   ██║██║╚██╗██║
  ██████╔╝██║  ██║███████║   ██║   ██║╚██████╔╝██║ ╚████║
  ╚═════╝ ╚═╝  ╚═╝╚══════╝   ╚═╝   ╚═╝ ╚═════╝ ╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    B.A.S.T.I.O.N. — Bulk Active-directory Stewardship: Tasks, Identity, Operations & Namespacing" -ForegroundColor Cyan
    Write-Host "    Active Directory & Identity Management Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Format-LastLogon {
    param($Date)
    if (-not $Date -or $Date -eq [DateTime]::MinValue) { return "Never" }
    return $Date.ToString("yyyy-MM-dd HH:mm")
}

function Get-DaysInactive {
    param($Date)
    if (-not $Date -or $Date -eq [DateTime]::MinValue) { return "N/A" }
    return [int]((Get-Date) - $Date).TotalDays
}

function HtmlEncode {
    param([string]$s)
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE & DOMAIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

function Test-DomainJoined {
    if ($env:USERDNSDOMAIN) {
        return $true
    }
    try {
        $cs = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
        return ($cs.PartOfDomain -eq $true)
    } catch {
        return $false
    }
}

function Assert-ADModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            return $true
        } catch {
            Write-Host "  [-] Failed to import the ActiveDirectory module: $_" -ForegroundColor $C.Error
            return $false
        }
    }

    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Warning
    Write-Host "  ACTIVE DIRECTORY MODULE NOT FOUND" -ForegroundColor $C.Warning
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host "  The ActiveDirectory PowerShell module is part of RSAT" -ForegroundColor $C.Info
    Write-Host "  (Remote Server Administration Tools). It is required for" -ForegroundColor $C.Info
    Write-Host "  all B.A.S.T.I.O.N. functions." -ForegroundColor $C.Info
    Write-Host ""
    Write-Host "  RSAT is supported on Windows 10 (1809+) and Windows 11." -ForegroundColor $C.Info
    Write-Host ""
    Write-Host "  [1] Install RSAT ActiveDirectory tools automatically (recommended)" -ForegroundColor $C.Info
    Write-Host "  [2] Show manual installation instructions" -ForegroundColor $C.Info
    Write-Host "  [Q] Cancel and quit" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $sel = (Read-Host).Trim().ToUpper()

    switch ($sel) {
        "1" {
            Write-Host ""
            Write-Host "  [*] Installing RSAT ActiveDirectory tools..." -ForegroundColor $C.Progress
            Write-Host "  [*] This may take several minutes. Please wait." -ForegroundColor $C.Progress
            Write-Host ""
            try {
                $result = Add-WindowsCapability -Online -Name "RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                if ($result.RestartNeeded) {
                    Write-Host "  [!!] A restart may be required to complete installation." -ForegroundColor $C.Warning
                }
                Write-Host "  [+] RSAT ActiveDirectory tools installed successfully." -ForegroundColor $C.Success
                Write-Host "  [*] Importing module..." -ForegroundColor $C.Progress
                Import-Module ActiveDirectory -ErrorAction Stop
                Write-Host "  [+] Module imported successfully." -ForegroundColor $C.Success
                return $true
            } catch {
                Write-Host "  [-] Automatic installation failed: $_" -ForegroundColor $C.Error
                Write-Host "  [!!] Try running: Add-WindowsCapability -Online -Name RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor $C.Warning
                Write-Host "  [!!] Or install RSAT manually from Windows Settings > Optional Features." -ForegroundColor $C.Warning
                return $false
            }
        }
        "2" {
            Write-Host ""
            Write-Host "  MANUAL INSTALLATION INSTRUCTIONS" -ForegroundColor $C.Header
            Write-Host ""
            Write-Host "  Option A — PowerShell (run as Administrator):" -ForegroundColor $C.Info
            Write-Host "    Add-WindowsCapability -Online -Name RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor $C.Warning
            Write-Host ""
            Write-Host "  Option B — Windows Settings:" -ForegroundColor $C.Info
            Write-Host "    Settings > System > Optional Features > Add a Feature" -ForegroundColor $C.Warning
            Write-Host "    Search for 'RSAT: Active Directory Domain Services'" -ForegroundColor $C.Warning
            Write-Host "    and install it." -ForegroundColor $C.Warning
            Write-Host ""
            Write-Host "  Option C — Windows Server / RSAT MSI:" -ForegroundColor $C.Info
            Write-Host "    Download RSAT from the Microsoft Download Center and run" -ForegroundColor $C.Warning
            Write-Host "    the installer, then enable the AD DS Tools feature." -ForegroundColor $C.Warning
            Write-Host ""
            Write-Host "  After installation, re-run bastion.ps1." -ForegroundColor $C.Info
            Write-Host ""
            Read-Host "  Press Enter to exit"
            return $false
        }
        default {
            Write-Host ""
            Write-Host "  Operation cancelled." -ForegroundColor $C.Info
            return $false
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# USER SEARCH
# ─────────────────────────────────────────────────────────────────────────────

function Search-ADUsers {
    Write-Section "SEARCH USERS"
    Write-Host "  Search by name, username (SamAccountName), or email address." -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter search term: " -ForegroundColor $C.Header
    $query = (Read-Host).Trim()

    if ([string]::IsNullOrWhiteSpace($query)) {
        Write-Host "  [!!] No search term entered." -ForegroundColor $C.Warning
        return $null
    }

    Write-Host ""
    Write-Host "  [*] Searching Active Directory..." -ForegroundColor $C.Progress

    try {
        $filter = "Name -like '*$query*' -or SamAccountName -like '*$query*' -or EmailAddress -like '*$query*' -or DisplayName -like '*$query*'"
        $users = Get-ADUser -Filter $filter -Properties DisplayName, EmailAddress, Department, Title, Enabled -ErrorAction Stop |
                 Sort-Object SamAccountName

        if ($users.Count -eq 0) {
            Write-Host "  [-] No users found matching '$query'." -ForegroundColor $C.Warning
            return $null
        }

        Write-Host "  [+] Found $($users.Count) user(s):" -ForegroundColor $C.Success
        Write-Host ""

        $i = 1
        foreach ($u in $users) {
            $statusColor = if ($u.Enabled) { $C.Success } else { $C.Info }
            $status      = if ($u.Enabled) { "Enabled" } else { "Disabled" }
            Write-Host ("  [{0,2}] {1,-22} {2,-30} {3}" -f $i, $u.SamAccountName, $u.DisplayName, $status) -ForegroundColor $statusColor
            $i++
        }

        Write-Host ""
        Write-Host -NoNewline "  Select user number (or Enter to cancel): " -ForegroundColor $C.Header
        $pick = (Read-Host).Trim()

        if ([string]::IsNullOrWhiteSpace($pick)) { return $null }

        $idx = 0
        if ([int]::TryParse($pick, [ref]$idx) -and $idx -ge 1 -and $idx -le $users.Count) {
            return $users[$idx - 1]
        } else {
            Write-Host "  [!!] Invalid selection." -ForegroundColor $C.Warning
            return $null
        }
    } catch {
        Write-Host "  [-] Search failed: $_" -ForegroundColor $C.Error
        return $null
    }
}

function Pick-ADUser {
    param([string]$Prompt = "Enter username or search term to find a user")
    Write-Host ""
    Write-Host "  $Prompt" -ForegroundColor $C.Info
    Write-Host -NoNewline "  Username (SamAccountName) or search term: " -ForegroundColor $C.Header
    $input = (Read-Host).Trim()
    if ([string]::IsNullOrWhiteSpace($input)) { return $null }

    try {
        # Try exact match first
        $user = Get-ADUser -Identity $input -Properties DisplayName, EmailAddress, Department, Title, Enabled -ErrorAction Stop
        return $user
    } catch {
        # Fall back to search
        Write-Host "  [*] Exact match not found — searching..." -ForegroundColor $C.Progress
        try {
            $filter = "Name -like '*$input*' -or SamAccountName -like '*$input*' -or DisplayName -like '*$input*'"
            $users  = Get-ADUser -Filter $filter -Properties DisplayName, EmailAddress, Department, Title, Enabled -ErrorAction Stop |
                      Sort-Object SamAccountName

            if ($users.Count -eq 0) {
                Write-Host "  [-] No users found matching '$input'." -ForegroundColor $C.Warning
                return $null
            }

            if ($users.Count -eq 1) {
                Write-Host "  [+] Found: $($users[0].SamAccountName) — $($users[0].DisplayName)" -ForegroundColor $C.Success
                return $users[0]
            }

            Write-Host "  [+] Found $($users.Count) matching users:" -ForegroundColor $C.Success
            Write-Host ""
            $i = 1
            foreach ($u in $users) {
                $statusColor = if ($u.Enabled) { $C.Success } else { $C.Info }
                Write-Host ("  [{0,2}] {1,-22} {2}" -f $i, $u.SamAccountName, $u.DisplayName) -ForegroundColor $statusColor
                $i++
            }
            Write-Host ""
            Write-Host -NoNewline "  Select user number (or Enter to cancel): " -ForegroundColor $C.Header
            $pick = (Read-Host).Trim()
            if ([string]::IsNullOrWhiteSpace($pick)) { return $null }

            $idx = 0
            if ([int]::TryParse($pick, [ref]$idx) -and $idx -ge 1 -and $idx -le $users.Count) {
                return $users[$idx - 1]
            } else {
                Write-Host "  [!!] Invalid selection." -ForegroundColor $C.Warning
                return $null
            }
        } catch {
            Write-Host "  [-] Search failed: $_" -ForegroundColor $C.Error
            return $null
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# VIEW USER DETAILS
# ─────────────────────────────────────────────────────────────────────────────

function Show-UserDetails {
    param($User)

    $props = @(
        'DisplayName','GivenName','Surname','SamAccountName','UserPrincipalName',
        'EmailAddress','Department','Title','Manager','Office','TelephoneNumber',
        'Enabled','LockedOut','PasswordExpired','PasswordNeverExpires',
        'PasswordLastSet','LastLogonDate','Created','Modified',
        'DistinguishedName','MemberOf'
    )

    try {
        $u = Get-ADUser -Identity $User.SamAccountName -Properties $props -ErrorAction Stop
    } catch {
        Write-Host "  [-] Could not retrieve details: $_" -ForegroundColor $C.Error
        return
    }

    # Resolve manager display name
    $managerName = "N/A"
    if ($u.Manager) {
        try {
            $mgr = Get-ADUser -Identity $u.Manager -Properties DisplayName -ErrorAction SilentlyContinue
            $managerName = if ($mgr.DisplayName) { $mgr.DisplayName } else { $mgr.SamAccountName }
        } catch { $managerName = $u.Manager }
    }

    # Resolve group memberships
    $groups = @()
    try {
        $groups = (Get-ADPrincipalGroupMembership -Identity $u.SamAccountName -ErrorAction Stop |
                   Sort-Object Name | ForEach-Object { $_.Name })
    } catch {
        Write-Host "    [!] Could not resolve group memberships: $_" -ForegroundColor $C.Warning
    }

    Write-Section "USER DETAILS — $($u.SamAccountName)"

    $enabledColor  = if ($u.Enabled)    { $C.Success } else { $C.Warning }
    $lockedColor   = if ($u.LockedOut)  { $C.Error   } else { $C.Success }
    $expiredColor  = if ($u.PasswordExpired) { $C.Warning } else { $C.Success }

    Write-Host "  Display Name       : $($u.DisplayName)"           -ForegroundColor $C.Info
    Write-Host "  First / Last       : $($u.GivenName) $($u.Surname)" -ForegroundColor $C.Info
    Write-Host "  Username           : $($u.SamAccountName)"        -ForegroundColor $C.Info
    Write-Host "  UPN                : $($u.UserPrincipalName)"     -ForegroundColor $C.Info
    Write-Host "  Email              : $(if ($u.EmailAddress) { $u.EmailAddress } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host "  Department         : $(if ($u.Department)   { $u.Department   } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host "  Title              : $(if ($u.Title)        { $u.Title        } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host "  Manager            : $managerName"               -ForegroundColor $C.Info
    Write-Host "  Office             : $(if ($u.Office) { $u.Office } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host "  Phone              : $(if ($u.TelephoneNumber) { $u.TelephoneNumber } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Account Enabled    : $($u.Enabled)  " -ForegroundColor $enabledColor
    Write-Host ""
    Write-Host -NoNewline "  Locked Out         : $($u.LockedOut)  " -ForegroundColor $lockedColor
    Write-Host ""
    Write-Host -NoNewline "  Password Expired   : $($u.PasswordExpired)  " -ForegroundColor $expiredColor
    Write-Host ""
    Write-Host "  Pwd Never Expires  : $($u.PasswordNeverExpires)" -ForegroundColor $C.Info
    Write-Host "  Password Last Set  : $(if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('yyyy-MM-dd HH:mm') } else { 'Never' })" -ForegroundColor $C.Info
    Write-Host "  Last Logon         : $(Format-LastLogon $u.LastLogonDate)" -ForegroundColor $C.Info
    Write-Host "  Account Created    : $(if ($u.Created) { $u.Created.ToString('yyyy-MM-dd') } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host "  Last Modified      : $(if ($u.Modified) { $u.Modified.ToString('yyyy-MM-dd') } else { 'N/A' })" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host "  Distinguished Name : $($u.DistinguishedName)" -ForegroundColor $C.Info
    Write-Host ""

    if ($groups.Count -gt 0) {
        Write-Host "  Group Memberships ($($groups.Count)):" -ForegroundColor $C.Header
        foreach ($g in $groups) {
            Write-Host "    - $g" -ForegroundColor $C.Info
        }
    } else {
        Write-Host "  Group Memberships  : None found" -ForegroundColor $C.Info
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# UNLOCK ACCOUNT
# ─────────────────────────────────────────────────────────────────────────────

function Unlock-UserAccount {
    Write-Section "UNLOCK ACCOUNT"
    $user = Pick-ADUser -Prompt "Enter the username of the account to unlock."
    if (-not $user) { return }

    try {
        $u = Get-ADUser -Identity $user.SamAccountName -Properties LockedOut -ErrorAction Stop

        if (-not $u.LockedOut) {
            Write-Host "  [!!] Account '$($u.SamAccountName)' is not currently locked." -ForegroundColor $C.Warning
            return
        }

        Write-Host "  [*] Unlocking account '$($u.SamAccountName)'..." -ForegroundColor $C.Progress
        Unlock-ADAccount -Identity $u.SamAccountName -ErrorAction Stop
        Write-Host "  [+] Account '$($u.SamAccountName)' has been unlocked successfully." -ForegroundColor $C.Success
    } catch {
        Write-Host "  [-] Failed to unlock account: $_" -ForegroundColor $C.Error
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# RESET PASSWORD
# ─────────────────────────────────────────────────────────────────────────────

function Reset-UserPassword {
    Write-Section "RESET PASSWORD"
    $user = Pick-ADUser -Prompt "Enter the username of the account to reset."
    if (-not $user) { return }

    Write-Host ""
    Write-Host "  Resetting password for: $($user.SamAccountName)" -ForegroundColor $C.Warning
    Write-Host "  The user will be required to change their password at next logon." -ForegroundColor $C.Info
    Write-Host ""

    $match = $false
    $newPwd = $null

    while (-not $match) {
        Write-Host -NoNewline "  Enter new password: " -ForegroundColor $C.Header
        $pwd1 = Read-Host -AsSecureString
        Write-Host -NoNewline "  Confirm new password: " -ForegroundColor $C.Header
        $pwd2 = Read-Host -AsSecureString

        # Compare secure strings by converting to plain text temporarily
        $plain1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                      [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd1))
        $plain2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                      [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd2))

        if ($plain1 -eq $plain2) {
            $match  = $true
            $newPwd = $pwd1
        } else {
            Write-Host ""
            Write-Host "  [!!] Passwords do not match. Please try again." -ForegroundColor $C.Warning
            Write-Host ""
        }

        # Zero out plain text strings from memory
        $plain1 = $null
        $plain2 = $null
    }

    try {
        Write-Host ""
        Write-Host "  [*] Setting new password..." -ForegroundColor $C.Progress
        Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword $newPwd -Reset -ErrorAction Stop
        Set-ADUser -Identity $user.SamAccountName -ChangePasswordAtLogon $true -ErrorAction Stop
        Write-Host "  [+] Password reset successfully for '$($user.SamAccountName)'." -ForegroundColor $C.Success
        Write-Host "  [+] User will be prompted to change password at next logon." -ForegroundColor $C.Success
    } catch {
        Write-Host "  [-] Password reset failed: $_" -ForegroundColor $C.Error
        Write-Host "  [!!] Ensure the password meets your domain's complexity requirements." -ForegroundColor $C.Warning
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# ENABLE / DISABLE ACCOUNT
# ─────────────────────────────────────────────────────────────────────────────

function Set-AccountState {
    Write-Section "ENABLE / DISABLE ACCOUNT"
    $user = Pick-ADUser -Prompt "Enter the username of the account to enable or disable."
    if (-not $user) { return }

    try {
        $u = Get-ADUser -Identity $user.SamAccountName -Properties Enabled -ErrorAction Stop
    } catch {
        Write-Host "  [-] Could not retrieve account: $_" -ForegroundColor $C.Error
        return
    }

    $currentState = if ($u.Enabled) { "Enabled" } else { "Disabled" }
    $stateColor   = if ($u.Enabled) { $C.Success } else { $C.Warning }

    Write-Host "  Account          : $($u.SamAccountName)" -ForegroundColor $C.Info
    Write-Host -NoNewline "  Current State    : " -ForegroundColor $C.Info
    Write-Host $currentState -ForegroundColor $stateColor
    Write-Host ""
    Write-Host "  [1] Enable account" -ForegroundColor $C.Info
    Write-Host "  [2] Disable account" -ForegroundColor $C.Info
    Write-Host "  [C] Cancel" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $sel = (Read-Host).Trim().ToUpper()

    switch ($sel) {
        "1" {
            try {
                Write-Host "  [*] Enabling account '$($u.SamAccountName)'..." -ForegroundColor $C.Progress
                Enable-ADAccount -Identity $u.SamAccountName -ErrorAction Stop
                Write-Host "  [+] Account enabled successfully." -ForegroundColor $C.Success
            } catch {
                Write-Host "  [-] Failed to enable account: $_" -ForegroundColor $C.Error
            }
        }
        "2" {
            try {
                Write-Host "  [*] Disabling account '$($u.SamAccountName)'..." -ForegroundColor $C.Progress
                Disable-ADAccount -Identity $u.SamAccountName -ErrorAction Stop
                Write-Host "  [+] Account disabled successfully." -ForegroundColor $C.Success
            } catch {
                Write-Host "  [-] Failed to disable account: $_" -ForegroundColor $C.Error
            }
        }
        default {
            Write-Host "  [*] Operation cancelled." -ForegroundColor $C.Info
        }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# VIEW GROUP MEMBERSHIP
# ─────────────────────────────────────────────────────────────────────────────

function Show-GroupMembership {
    Write-Section "VIEW GROUP MEMBERSHIP"
    $user = Pick-ADUser -Prompt "Enter the username to view group memberships."
    if (-not $user) { return }

    Write-Host "  [*] Retrieving group memberships for '$($user.SamAccountName)'..." -ForegroundColor $C.Progress
    Write-Host ""

    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName -ErrorAction Stop |
                  Sort-Object Name

        if ($groups.Count -eq 0) {
            Write-Host "  [-] No group memberships found for '$($user.SamAccountName)'." -ForegroundColor $C.Warning
        } else {
            Write-Host "  Group memberships for $($user.SamAccountName) ($($groups.Count) groups):" -ForegroundColor $C.Success
            Write-Host ""
            foreach ($g in $groups) {
                Write-Host ("  {0,-40} {1}" -f $g.Name, $g.GroupScope) -ForegroundColor $C.Info
            }
        }
    } catch {
        Write-Host "  [-] Failed to retrieve group memberships: $_" -ForegroundColor $C.Error
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# ADD / REMOVE USER FROM GROUP
# ─────────────────────────────────────────────────────────────────────────────

function Manage-GroupMembership {
    Write-Section "ADD / REMOVE USER FROM GROUP"
    $user = Pick-ADUser -Prompt "Enter the username to add or remove from a group."
    if (-not $user) { return }

    Write-Host ""
    Write-Host -NoNewline "  Enter group name (or partial name to search): " -ForegroundColor $C.Header
    $groupQuery = (Read-Host).Trim()

    if ([string]::IsNullOrWhiteSpace($groupQuery)) {
        Write-Host "  [!!] No group name entered." -ForegroundColor $C.Warning
        return
    }

    # Find the group
    $targetGroup = $null
    try {
        # Try exact match first
        $targetGroup = Get-ADGroup -Identity $groupQuery -ErrorAction Stop
    } catch {
        # Fall back to search
        Write-Host "  [*] Searching for groups matching '$groupQuery'..." -ForegroundColor $C.Progress
        try {
            $groups = Get-ADGroup -Filter "Name -like '*$groupQuery*'" -ErrorAction Stop | Sort-Object Name

            if ($groups.Count -eq 0) {
                Write-Host "  [-] No groups found matching '$groupQuery'." -ForegroundColor $C.Warning
                return
            }

            if ($groups.Count -eq 1) {
                $targetGroup = $groups[0]
                Write-Host "  [+] Found group: $($targetGroup.Name)" -ForegroundColor $C.Success
            } else {
                Write-Host "  [+] Found $($groups.Count) matching groups:" -ForegroundColor $C.Success
                Write-Host ""
                $i = 1
                foreach ($g in $groups) {
                    Write-Host ("  [{0,2}] {1,-40} {2}" -f $i, $g.Name, $g.GroupScope) -ForegroundColor $C.Info
                    $i++
                }
                Write-Host ""
                Write-Host -NoNewline "  Select group number (or Enter to cancel): " -ForegroundColor $C.Header
                $pick = (Read-Host).Trim()
                if ([string]::IsNullOrWhiteSpace($pick)) { return }

                $idx = 0
                if ([int]::TryParse($pick, [ref]$idx) -and $idx -ge 1 -and $idx -le $groups.Count) {
                    $targetGroup = $groups[$idx - 1]
                } else {
                    Write-Host "  [!!] Invalid selection." -ForegroundColor $C.Warning
                    return
                }
            }
        } catch {
            Write-Host "  [-] Group search failed: $_" -ForegroundColor $C.Error
            return
        }
    }

    Write-Host ""
    Write-Host "  User  : $($user.SamAccountName)" -ForegroundColor $C.Info
    Write-Host "  Group : $($targetGroup.Name)" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host "  [1] Add user to group" -ForegroundColor $C.Info
    Write-Host "  [2] Remove user from group" -ForegroundColor $C.Info
    Write-Host "  [C] Cancel" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $sel = (Read-Host).Trim().ToUpper()

    switch ($sel) {
        "1" {
            try {
                Write-Host "  [*] Adding '$($user.SamAccountName)' to '$($targetGroup.Name)'..." -ForegroundColor $C.Progress
                Add-ADGroupMember -Identity $targetGroup.SamAccountName -Members $user.SamAccountName -ErrorAction Stop
                Write-Host "  [+] User added to group successfully." -ForegroundColor $C.Success
            } catch {
                Write-Host "  [-] Failed to add user to group: $_" -ForegroundColor $C.Error
                Write-Host "  [!!] The user may already be a member of this group." -ForegroundColor $C.Warning
            }
        }
        "2" {
            try {
                Write-Host "  [*] Removing '$($user.SamAccountName)' from '$($targetGroup.Name)'..." -ForegroundColor $C.Progress
                Remove-ADGroupMember -Identity $targetGroup.SamAccountName -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
                Write-Host "  [+] User removed from group successfully." -ForegroundColor $C.Success
            } catch {
                Write-Host "  [-] Failed to remove user from group: $_" -ForegroundColor $C.Error
                Write-Host "  [!!] The user may not be a member of this group." -ForegroundColor $C.Warning
            }
        }
        default {
            Write-Host "  [*] Operation cancelled." -ForegroundColor $C.Info
        }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# FIND STALE ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────

function Get-StaleAccounts {
    Write-Host "  [*] Searching for stale accounts (no logon in 90+ days, enabled only)..." -ForegroundColor $C.Progress

    $staleDate = (Get-Date).AddDays(-90)

    try {
        $stale = Get-ADUser -Filter { Enabled -eq $true -and LastLogonDate -lt $staleDate } `
                            -Properties DisplayName, Department, LastLogonDate, Enabled, SamAccountName `
                            -ErrorAction Stop |
                 Sort-Object LastLogonDate

        # Also include accounts that have never logged on (null LastLogonDate)
        $neverLogon = Get-ADUser -Filter { Enabled -eq $true -and LastLogonDate -notlike "*" } `
                                 -Properties DisplayName, Department, LastLogonDate, Enabled, SamAccountName `
                                 -ErrorAction SilentlyContinue |
                      Sort-Object SamAccountName

        $combined = @()
        if ($stale)      { $combined += $stale      }
        if ($neverLogon) { $combined += $neverLogon }

        # Deduplicate
        $combined = $combined | Sort-Object SamAccountName -Unique

        return $combined
    } catch {
        Write-Host "  [-] Failed to query stale accounts: $_" -ForegroundColor $C.Error
        return @()
    }
}

function Show-StaleAccounts {
    Write-Section "FIND STALE ACCOUNTS"

    $stale = Get-StaleAccounts

    if ($stale.Count -eq 0) {
        Write-Host "  [+] No stale accounts found. All enabled accounts have logged in within 90 days." -ForegroundColor $C.Success
        Write-Host ""
        return
    }

    Write-Host "  [!!] Found $($stale.Count) stale account(s):" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host ("  {0,-22} {1,-30} {2,-20} {3}" -f "Username", "Display Name", "Last Logon", "Days Inactive") -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 82)) -ForegroundColor $C.Header

    foreach ($u in $stale) {
        $lastLogon = Format-LastLogon $u.LastLogonDate
        $days      = Get-DaysInactive $u.LastLogonDate

        Write-Host ("  {0,-22} {1,-30} {2,-20} {3}" -f `
            $u.SamAccountName, `
            (if ($u.DisplayName) { $u.DisplayName } else { "N/A" }), `
            $lastLogon, `
            $days) -ForegroundColor $C.Warning
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT — STALE ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────

function Export-StaleReport {
    Write-Section "EXPORT STALE ACCOUNTS REPORT"

    $stale = Get-StaleAccounts

    if ($stale.Count -eq 0) {
        Write-Host "  [+] No stale accounts found — nothing to export." -ForegroundColor $C.Success
        Write-Host ""
        return
    }

    Write-Host "  [+] Found $($stale.Count) stale account(s). Building report..." -ForegroundColor $C.Success

    $reportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $domain          = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { $env:COMPUTERNAME }

    # Find oldest last logon
    $oldestDate   = $stale | Where-Object { $_.LastLogonDate } | Sort-Object LastLogonDate | Select-Object -First 1
    $oldestLogon  = if ($oldestDate) { $oldestDate.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
    $oldestUser   = if ($oldestDate) { $oldestDate.SamAccountName } else { "N/A" }

    # Build rows
    $rows = ""
    foreach ($u in $stale) {
        $lastLogon = Format-LastLogon $u.LastLogonDate
        $days      = Get-DaysInactive $u.LastLogonDate
        $dept      = if ($u.Department) { HtmlEncode $u.Department } else { "N/A" }
        $dispName  = if ($u.DisplayName) { HtmlEncode $u.DisplayName } else { "N/A" }

        $daysClass = if ($days -eq "N/A" -or [int]$days -gt 365) { "err" } elseif ([int]$days -gt 180) { "warn" } else { "warn" }

        $rows += @"
            <tr>
                <td><strong>$(HtmlEncode $u.SamAccountName)</strong></td>
                <td>$dispName</td>
                <td>$dept</td>
                <td>$lastLogon</td>
                <td><span class="tk-badge-$daysClass">$days</span></td>
                <td><span class="tk-badge-ok">Enabled</span></td>
            </tr>
"@
    }

    $html = (Get-TKHtmlHead `
        -Title     'B.A.S.T.I.O.N. Stale Accounts Report' `
        -ScriptName 'B.A.S.T.I.O.N.' `
        -Subtitle  "Domain: $domain" `
        -MetaItems ([ordered]@{ 'Generated' = $reportTimestamp; 'Stale Threshold' = '90 days' }) `
        -NavItems  @('Stale User Accounts')) + @"

<div class="tk-info-box">
  <span class="tk-info-label">Note</span>
  Stale threshold: accounts enabled but with no logon activity in the past 90 days, or accounts that have never logged on.
  Review these accounts and disable or remove those that are no longer needed.
</div>

<div class="tk-summary-row">
  <div class="tk-summary-card warn">
    <div class="tk-summary-num">$($stale.Count)</div>
    <div class="tk-summary-lbl">Total Stale</div>
  </div>
  <div class="tk-summary-card err">
    <div class="tk-summary-num">$oldestLogon</div>
    <div class="tk-summary-lbl">Oldest Last Logon</div>
  </div>
  <div class="tk-summary-card info">
    <div class="tk-summary-num">$oldestUser</div>
    <div class="tk-summary-lbl">Oldest Account</div>
  </div>
</div>

<div class="tk-section">
  <div class="tk-section-title">Stale User Accounts</div>
  <table class="tk-table">
    <thead>
      <tr>
        <th>Username</th>
        <th>Display Name</th>
        <th>Department</th>
        <th>Last Logon</th>
        <th>Days Inactive</th>
        <th>Enabled</th>
      </tr>
    </thead>
    <tbody>
      $rows
    </tbody>
  </table>
</div>

"@ + (Get-TKHtmlFoot -ScriptName 'B.A.S.T.I.O.N. v1.0')

    $reportFilename = "BASTION_Stale_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $reportPath     = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportFilename

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Host "  [+] Stale accounts report saved:" -ForegroundColor $C.Success
        Write-Host "      $reportPath" -ForegroundColor $C.Success

        if (-not $Unattended) {
            Write-Host ""
            Write-Host -NoNewline "  Open report in browser? (Y/N): " -ForegroundColor $C.Header
            $open = (Read-Host).Trim().ToUpper()
            if ($open -eq "Y") {
                Start-Process $reportPath
            }
        }
    } catch {
        Write-Host "  [-] Failed to save report: $_" -ForegroundColor $C.Error
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# ACCOUNT LOCKOUT FORENSICS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-LockoutForensics {
    Write-Section "ACCOUNT LOCKOUT INVESTIGATION"
    Write-Host "  This tool queries the PDC Emulator's Security event log for lockout" -ForegroundColor $C.Info
    Write-Host "  events (Event ID 4740) to identify the source of an account lockout." -ForegroundColor $C.Info
    Write-Host ""

    $user = Pick-ADUser -Prompt "Enter the username to investigate."
    if (-not $user) { return }

    Write-Host "  [*] Locating PDC Emulator..." -ForegroundColor $C.Progress
    try {
        $pdc = (Get-ADDomain -ErrorAction Stop).PDCEmulator
        Write-Host "  [+] PDC Emulator : $pdc" -ForegroundColor $C.Success
    } catch {
        Write-Host "  [-] Could not determine PDC Emulator: $_" -ForegroundColor $C.Error
        Write-Host "  [!!] Ensure this machine is joined to the domain and can contact a DC." -ForegroundColor $C.Warning
        return
    }

    Write-Host ""
    Write-Host "  [*] Querying Security event log on $pdc for Event ID 4740 (last 7 days)..." -ForegroundColor $C.Progress
    Write-Host "  [*] This may take a moment..." -ForegroundColor $C.Progress
    Write-Host ""

    try {
        $allEvents = Get-WinEvent -ComputerName $pdc -FilterHashtable @{
            LogName   = 'Security'
            Id        = 4740
            StartTime = (Get-Date).AddDays(-7)
        } -ErrorAction Stop

        $events = $allEvents | Where-Object {
            $_.Properties[0].Value -ieq $user.SamAccountName
        } | Sort-Object TimeCreated -Descending

        if ($events.Count -eq 0) {
            Write-Host "  [+] No lockout events found for '$($user.SamAccountName)' in the last 7 days." -ForegroundColor $C.Success
            Write-Host ""
            return
        }

        Write-Host "  [!!] Found $($events.Count) lockout event(s) for '$($user.SamAccountName)':" -ForegroundColor $C.Warning
        Write-Host ""
        Write-Host ("  {0,-22} {1,-32} {2}" -f "Timestamp", "Source Machine (Caller)", "Domain Controller") -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 82)) -ForegroundColor $C.Header

        foreach ($evt in $events) {
            $timestamp     = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            $callerMachine = $evt.Properties[4].Value
            $dc            = $evt.MachineName
            if ([string]::IsNullOrWhiteSpace($callerMachine)) { $callerMachine = "(unknown)" }

            Write-Host ("  {0,-22} {1,-32} {2}" -f $timestamp, $callerMachine, $dc) -ForegroundColor $C.Warning
        }

        Write-Host ""
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  REMEDIATION TIPS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  The 'Source Machine' column shows where the bad credentials originated." -ForegroundColor $C.Info
        Write-Host "  Common causes:" -ForegroundColor $C.Info
        Write-Host "    - Saved credentials in Credential Manager referencing old password" -ForegroundColor $C.Info
        Write-Host "    - Mapped network drives or printers with cached credentials" -ForegroundColor $C.Info
        Write-Host "    - Scheduled tasks running under the user account" -ForegroundColor $C.Info
        Write-Host "    - Mobile device or Outlook profile with old password" -ForegroundColor $C.Info
        Write-Host "    - Applications with embedded credentials" -ForegroundColor $C.Info
        Write-Host ""

    } catch {
        Write-Host "  [-] Could not query event log on $pdc : $_" -ForegroundColor $C.Error
        Write-Host "  [!!] Ensure you have remote event log access (Event Log Readers group) on the PDC." -ForegroundColor $C.Warning
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PASSWORD EXPIRY REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Get-ExpiringPasswords {
    param([int]$ThresholdDays)

    Write-Host "  [*] Retrieving domain password policy..." -ForegroundColor $C.Progress
    try {
        $maxAge = (Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop).MaxPasswordAge.Days
        if ($maxAge -eq 0) {
            Write-Host "  [!] Domain default password policy has no maximum age (passwords never expire)." -ForegroundColor $C.Warning
            Write-Host "      Fine-Grained Password Policies (PSOs) may still apply to specific users." -ForegroundColor $C.Info
            return @()
        }
        Write-Host "  [+] Domain max password age : $maxAge days" -ForegroundColor $C.Info
    } catch {
        Write-Host "  [-] Could not retrieve domain password policy: $_" -ForegroundColor $C.Error
        return @()
    }

    Write-Host "  [*] Querying users with passwords expiring within $ThresholdDays days..." -ForegroundColor $C.Progress

    try {
        $cutoff = (Get-Date).AddDays($ThresholdDays)
        $users  = Get-ADUser -Filter {
            Enabled -eq $true -and PasswordNeverExpires -eq $false
        } -Properties DisplayName, EmailAddress, Department, PasswordLastSet, PasswordNeverExpires -ErrorAction Stop |
            Where-Object {
                $_.PasswordLastSet -ne $null -and
                $_.PasswordLastSet.AddDays($maxAge) -le $cutoff
            } |
            ForEach-Object {
                $expiry   = $_.PasswordLastSet.AddDays($maxAge)
                $daysLeft = [int]($expiry - (Get-Date)).TotalDays
                [PSCustomObject]@{
                    SamAccountName = $_.SamAccountName
                    DisplayName    = $_.DisplayName
                    Department     = $_.Department
                    EmailAddress   = $_.EmailAddress
                    Expiry         = $expiry
                    DaysLeft       = $daysLeft
                    Status         = if ($daysLeft -lt 0) { 'Expired' } elseif ($daysLeft -le 7) { 'Critical' } elseif ($daysLeft -le 14) { 'Warning' } else { 'Expiring' }
                }
            } |
            Sort-Object Expiry

        return $users
    } catch {
        Write-Host "  [-] Query failed: $_" -ForegroundColor $C.Error
        return @()
    }
}

function Show-PasswordExpiryReport {
    Write-Section "PASSWORD EXPIRY REPORT"

    Write-Host -NoNewline "  Report users expiring within how many days? (default 30): " -ForegroundColor $C.Header
    $raw = (Read-Host).Trim()
    $thresholdDays = 30
    if (-not [string]::IsNullOrWhiteSpace($raw)) {
        $parsed = 0
        if ([int]::TryParse($raw, [ref]$parsed) -and $parsed -gt 0) {
            $thresholdDays = $parsed
        } else {
            Write-Host "  [!] Invalid input  -  using default of 30 days." -ForegroundColor $C.Warning
        }
    }

    Write-Host ""
    $users = Get-ExpiringPasswords -ThresholdDays $thresholdDays
    if ($users.Count -eq 0) {
        Write-Host "  [+] No users have passwords expiring within $thresholdDays days." -ForegroundColor $C.Success
        Write-Host ""
        return
    }

    Write-Host ""
    Write-Host "  [!!] $($users.Count) user(s) with passwords expiring within $thresholdDays days:" -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host ("  {0,-22} {1,-28} {2,-14} {3}" -f "Username", "Display Name", "Expires On", "Days Left") -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 78)) -ForegroundColor $C.Header

    foreach ($u in $users) {
        $color    = switch ($u.Status) {
            'Expired'  { $C.Error   }
            'Critical' { $C.Error   }
            'Warning'  { $C.Warning }
            default    { $C.Info    }
        }
        $daysStr = if ($u.DaysLeft -lt 0) { "EXPIRED ($([Math]::Abs($u.DaysLeft))d ago)" } else { "$($u.DaysLeft) days" }
        Write-Host ("  {0,-22} {1,-28} {2,-14} {3}" -f `
            $u.SamAccountName, `
            (if ($u.DisplayName) { $u.DisplayName } else { "N/A" }), `
            $u.Expiry.ToString("yyyy-MM-dd"), `
            $daysStr) -ForegroundColor $color
    }

    Write-Host ""
}

function Export-PasswordExpiryReport {
    Write-Section "EXPORT PASSWORD EXPIRY REPORT"

    $thresholdDays = 30
    if (-not $Unattended) {
        Write-Host -NoNewline "  Report users expiring within how many days? (default 30): " -ForegroundColor $C.Header
        $raw = (Read-Host).Trim()
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $parsed = 0
            if ([int]::TryParse($raw, [ref]$parsed) -and $parsed -gt 0) { $thresholdDays = $parsed }
        }
    }

    Write-Host ""
    $users = Get-ExpiringPasswords -ThresholdDays $thresholdDays
    if ($users.Count -eq 0) {
        Write-Host "  [+] No users with passwords expiring within $thresholdDays days  -  nothing to export." -ForegroundColor $C.Success
        Write-Host ""
        return
    }

    Write-Host "  [+] Building report for $($users.Count) user(s)..." -ForegroundColor $C.Success

    $reportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $domain          = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { $env:COMPUTERNAME }

    $expiredCount  = ($users | Where-Object { $_.DaysLeft -lt 0 }).Count
    $criticalCount = ($users | Where-Object { $_.DaysLeft -ge 0 -and $_.DaysLeft -le 7 }).Count
    $warningCount  = ($users | Where-Object { $_.DaysLeft -gt 7 -and $_.DaysLeft -le 14 }).Count
    $expiringCount = ($users | Where-Object { $_.DaysLeft -gt 14 }).Count

    $rows = ""
    foreach ($u in $users) {
        $badgeClass = switch ($u.Status) {
            'Expired'  { 'err' }
            'Critical' { 'err' }
            'Warning'  { 'warn' }
            default    { 'info' }
        }
        $daysStr = if ($u.DaysLeft -lt 0) { "EXPIRED" } else { "$($u.DaysLeft)d" }
        $dept    = if ($u.Department)   { HtmlEncode $u.Department }   else { "N/A" }
        $disp    = if ($u.DisplayName)  { HtmlEncode $u.DisplayName }  else { "N/A" }
        $email   = if ($u.EmailAddress) { HtmlEncode $u.EmailAddress } else { "N/A" }

        $rows += @"
            <tr>
                <td><strong>$(HtmlEncode $u.SamAccountName)</strong></td>
                <td>$disp</td>
                <td>$dept</td>
                <td>$email</td>
                <td>$($u.Expiry.ToString("yyyy-MM-dd"))</td>
                <td><span class="tk-badge-$badgeClass">$daysStr</span></td>
            </tr>
"@
    }

    $html = (Get-TKHtmlHead `
        -Title     'B.A.S.T.I.O.N. Password Expiry Report' `
        -ScriptName 'B.A.S.T.I.O.N.' `
        -Subtitle  "Domain: $domain" `
        -MetaItems ([ordered]@{ 'Generated' = $reportTimestamp; 'Threshold' = "$thresholdDays days" }) `
        -NavItems  @('Users with Expiring Passwords')) + @"

<div class="tk-info-box">
  <span class="tk-info-label">Note</span>
  Shows all enabled users (without PasswordNeverExpires) whose passwords expire within $thresholdDays days.
  Contact these users to prompt a password change, or reset passwords as appropriate.
</div>

<div class="tk-summary-row">
  <div class="tk-summary-card err">
    <div class="tk-summary-num">$expiredCount</div>
    <div class="tk-summary-lbl">Expired</div>
  </div>
  <div class="tk-summary-card err">
    <div class="tk-summary-num">$criticalCount</div>
    <div class="tk-summary-lbl">Critical (&lt;7d)</div>
  </div>
  <div class="tk-summary-card warn">
    <div class="tk-summary-num">$warningCount</div>
    <div class="tk-summary-lbl">Warning (&lt;14d)</div>
  </div>
  <div class="tk-summary-card info">
    <div class="tk-summary-num">$expiringCount</div>
    <div class="tk-summary-lbl">Expiring Soon</div>
  </div>
</div>

<div class="tk-section">
  <div class="tk-section-title">Users with Expiring Passwords</div>
  <table class="tk-table">
    <thead>
      <tr>
        <th>Username</th><th>Display Name</th><th>Department</th><th>Email</th><th>Expires On</th><th>Days Left</th>
      </tr>
    </thead>
    <tbody>
      $rows
    </tbody>
  </table>
</div>

"@ + (Get-TKHtmlFoot -ScriptName 'B.A.S.T.I.O.N. v1.0')

    $reportFilename = "BASTION_PwdExpiry_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $reportPath     = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportFilename

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Host "  [+] Password expiry report saved:" -ForegroundColor $C.Success
        Write-Host "      $reportPath" -ForegroundColor $C.Success

        if (-not $Unattended) {
            Write-Host ""
            Write-Host -NoNewline "  Open report in browser? (Y/N): " -ForegroundColor $C.Header
            $open = (Read-Host).Trim().ToUpper()
            if ($open -eq "Y") { Start-Process $reportPath }
        }
    } catch {
        Write-Host "  [-] Failed to save report: $_" -ForegroundColor $C.Error
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-BastionBanner

    # Domain-join warning
    if (-not (Test-DomainJoined)) {
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Warning
        Write-Host "  WARNING: This machine does not appear to be domain-joined." -ForegroundColor $C.Warning
        Write-Host "  AD operations require a domain-joined machine or a domain" -ForegroundColor $C.Warning
        Write-Host "  controller accessible on the network." -ForegroundColor $C.Warning
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Warning
        Write-Host ""
    } else {
        Write-Host "  Domain : $env:USERDNSDOMAIN" -ForegroundColor $C.Info
        Write-Host ""
    }

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  MAIN MENU" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  [1] Search users" -ForegroundColor $C.Info
    Write-Host "  [2] View user details" -ForegroundColor $C.Info
    Write-Host "  [3] Unlock account" -ForegroundColor $C.Info
    Write-Host "  [4] Reset password" -ForegroundColor $C.Info
    Write-Host "  [5] Enable / Disable account" -ForegroundColor $C.Info
    Write-Host "  [6] View group membership" -ForegroundColor $C.Info
    Write-Host "  [7] Add / Remove user from group" -ForegroundColor $C.Info
    Write-Host "  [8] Find stale accounts  (90+ days inactive)" -ForegroundColor $C.Info
    Write-Host "  [9] Export stale accounts report  (HTML)" -ForegroundColor $C.Info
    Write-Host "  [10] Investigate account lockout  (lockout source forensics)" -ForegroundColor $C.Info
    Write-Host "  [11] Password expiry report" -ForegroundColor $C.Info
    Write-Host "  [12] Export password expiry report  (HTML)" -ForegroundColor $C.Info
    Write-Host "  [Q] Quit" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

# Module check
if (-not (Assert-ADModule)) {
    exit 1
}

if ($Unattended) {
    switch ($Action) {
        'StaleReport' {
            Write-Host "[*] B.A.S.T.I.O.N.  -  Running unattended stale accounts report..." -ForegroundColor $C.Progress
            Export-StaleReport
        }
        'PasswordExpiryReport' {
            Write-Host "[*] B.A.S.T.I.O.N.  -  Running unattended password expiry report..." -ForegroundColor $C.Progress
            Export-PasswordExpiryReport
        }
    }
} else {
    $choice = ""

    do {
        Show-Menu
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            "1" {
                $selectedUser = Search-ADUsers
                if ($selectedUser) {
                    Write-Host ""
                    Write-Host -NoNewline "  View full details for '$($selectedUser.SamAccountName)'? (Y/N): " -ForegroundColor $C.Header
                    $viewDetails = (Read-Host).Trim().ToUpper()
                    if ($viewDetails -eq "Y") {
                        Show-UserDetails -User $selectedUser
                    }
                }
            }
            "2" {
                $user = Pick-ADUser -Prompt "Enter the username to view details."
                if ($user) { Show-UserDetails -User $user }
            }
            "3" { Unlock-UserAccount }
            "4" { Reset-UserPassword }
            "5" { Set-AccountState }
            "6" { Show-GroupMembership }
            "7" { Manage-GroupMembership }
            "8" { Show-StaleAccounts }
            "9" { Export-StaleReport }
            "10" { Invoke-LockoutForensics }
            "11" { Show-PasswordExpiryReport }
            "12" { Export-PasswordExpiryReport }
            "Q" {
                Write-Host ""
                Write-Host "  Closing B.A.S.T.I.O.N." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-12 or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @("Q")) {
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne "Q")
}

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
