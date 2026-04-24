<#
.SYNOPSIS
    C.O.V.E.N.A.N.T. — Configures Onboarding Via Enrollment — Network, Accounts, Naming & Timezone
    Machine Setup & Domain Enrollment Tool for PowerShell 5.1+

.DESCRIPTION
    Guides a technician through the full setup of a new Windows machine:
    computer rename, Entra ID or on-premises Active Directory domain join with
    interactive credential entry, network drive mapping, local admin account
    creation, timezone configuration, and Group Policy synchronisation. Ends
    with an action summary and optional reboot countdown.

.USAGE
    PS C:\> .\covenant.ps1 -WhatIf                # Preview all onboarding steps without making changes
    PS C:\> .\covenant.ps1                                                     # Must be run as Administrator
    PS C:\> .\covenant.ps1 -Unattended -NewComputerName "DESKTOP-01"           # Rename only
    PS C:\> .\covenant.ps1 -Unattended -Timezone "Eastern Standard Time"       # Set timezone only
    PS C:\> .\covenant.ps1 -Unattended -LocalAdminUser "admin" -LocalAdminPassword (ConvertTo-SecureString "Pass" -AsPlainText -Force)
    PS C:\> .\covenant.ps1 -Unattended -DomainJoinType "AD" -ADDomain "corp.contoso.com"
    PS C:\> .\covenant.ps1 -Unattended -DomainJoinType "AD" -ADDomain "corp.contoso.com" -ADOUPath "OU=Workstations,DC=corp,DC=contoso,DC=com"

.NOTES
    Version : 3.0

#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [ValidatePattern('^$|^[A-Za-z0-9]$|^[A-Za-z0-9][A-Za-z0-9-]{0,13}[A-Za-z0-9]$')]
    [string]$NewComputerName = "",
    [string]$Timezone        = "",
    [string]$LocalAdminUser  = "",
    [securestring]$LocalAdminPassword = $null,
    [switch]$Transcript,
    [string]$DomainJoinType  = "",      # "Entra", "AD", or "" (prompt interactively)
    [string]$ADDomain        = "",      # e.g. "corp.contoso.com"
    [string]$ADOUPath        = "",      # optional, e.g. "OU=Workstations,DC=corp,DC=contoso,DC=com"
    [switch]$SkipGPUpdate               # skip the new gpupdate step
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

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
Assert-AdminPrivilege

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

# ─────────────────────────────────────────────────────────────────────────────
# BANNER DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-CovenantBanner {
    Write-Host @"

   ██████╗  ██████╗ ██╗   ██╗███████╗███╗   ██╗ █████╗ ███╗   ██╗████████╗
 ██╔════╝ ██╔═══██╗██║   ██║██╔════╝████╗  ██║██╔══██╗████╗  ██║╚══██╔══╝
 ██║      ██║   ██║██║   ██║█████╗  ██╔██╗ ██║███████║██╔██╗ ██║   ██║
 ██║      ██║   ██║╚██╗ ██╔╝██╔══╝  ██║╚██╗██║██╔══██║██║╚██╗██║   ██║
 ╚██████╗ ╚██████╔╝ ╚████╔╝ ███████╗██║ ╚████║██║  ██║██║ ╚████║   ██║
  ╚═════╝  ╚═════╝   ╚═══╝  ╚══════╝╚═╝  ╚═══╝╚═╝  ╚═╝╚═╝  ╚═══╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    C.O.V.E.N.A.N.T. — Configures Onboarding Via Enrollment — Network, Accounts, Naming & Timezone" -ForegroundColor Cyan
    Write-Host "    Machine Setup & Domain Enrollment Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA DEFINITION
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
    param([string]$Step, [string]$Status, [string]$Detail = "")
    [void]$ActionLog.Add([PSCustomObject]@{
        Step   = $Step
        Status = $Status
        Detail = $Detail
    })
}

# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY BANNER & HEADER
# ─────────────────────────────────────────────────────────────────────────────

Show-CovenantBanner

if ($WhatIf) {
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "  [~] DRY RUN MODE — No changes will be made to this system." -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host ""
}

# Apply config defaults for parameters not supplied on the command line
$_cfg = Get-TKConfig
if ([string]::IsNullOrWhiteSpace($Timezone)       -and -not [string]::IsNullOrWhiteSpace($_cfg.Covenant.DefaultTimezone))       { $Timezone      = $_cfg.Covenant.DefaultTimezone }
if ([string]::IsNullOrWhiteSpace($LocalAdminUser) -and -not [string]::IsNullOrWhiteSpace($_cfg.Covenant.DefaultLocalAdminUser)) { $LocalAdminUser = $_cfg.Covenant.DefaultLocalAdminUser }

$executionTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$rebootRequired = $false

Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "     MACHINE ONBOARDING WIZARD" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Run As    : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Timestamp : $executionTime" -ForegroundColor $ColorSchema.Info
Write-Host ""
Write-Host "  Each step below is optional. Press Enter to skip any step." -ForegroundColor $ColorSchema.Warning
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: PRE-FLIGHT CHECKS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[1/7] Running Pre-Flight Checks..." -ForegroundColor $ColorSchema.Progress

# OS Version
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    Write-Host "    OS       : $($os.Caption) (Build $($os.BuildNumber))" -ForegroundColor $ColorSchema.Info
}
catch {
    Write-Host "    [!!] Could not retrieve OS info: $_" -ForegroundColor $ColorSchema.Warning
}

# Current join status
Write-Host "    Checking current domain/Entra join status..." -ForegroundColor $ColorSchema.Info
try {
    $dsregOutput = & dsregcmd /status 2>&1
    $azureAdJoined  = ($dsregOutput | Where-Object { $_ -match "AzureAdJoined\s*:\s*YES" })    -ne $null
    $domainJoined   = ($dsregOutput | Where-Object { $_ -match "DomainJoined\s*:\s*YES" })     -ne $null
    $workplaceJoined = ($dsregOutput | Where-Object { $_ -match "WorkplaceJoined\s*:\s*YES" }) -ne $null

    if ($azureAdJoined) {
        Write-Host "    [!!] This machine is already Entra ID (Azure AD) joined." -ForegroundColor $ColorSchema.Warning
        $tenantName = ($dsregOutput | Select-String "TenantName\s*:\s*(.+)" | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() })
        if ($tenantName) { Write-Host "         Tenant: $tenantName" -ForegroundColor $ColorSchema.Warning }
    }
    elseif ($domainJoined) {
        Write-Host "    [!!] This machine is joined to an on-premises Active Directory domain." -ForegroundColor $ColorSchema.Warning
        Write-Host "         Hybrid Entra ID join may be available, but a clean Entra join requires unjoin first." -ForegroundColor $ColorSchema.Warning
    }
    else {
        Write-Host "    [+] Machine is not domain-joined. Ready for Entra ID enrollment." -ForegroundColor $ColorSchema.Success
    }

    if ($workplaceJoined) {
        Write-Host "    [!!] A workplace (registered) account is present. This may conflict with a full join." -ForegroundColor $ColorSchema.Warning
    }
}
catch {
    Write-Host "    [-] Could not determine join status: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: COMPUTER RENAME
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[2/7] Computer Rename..." -ForegroundColor $ColorSchema.Progress
Write-Host "    Current name: $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host ""

if ($Unattended) {
    $newName = $NewComputerName
} else {
    $newName = Read-Host "  Enter new computer name (leave blank to skip)"
}

if (-not [string]::IsNullOrWhiteSpace($newName)) {
    $newName = $newName.Trim()

    # Validate: 1–15 chars, letters/digits/hyphens, no leading or trailing hyphen
    if ($newName -match '^[a-zA-Z0-9][a-zA-Z0-9\-]{0,13}[a-zA-Z0-9]$' -or $newName -match '^[a-zA-Z0-9]$') {
        if ($WhatIf) {
            Write-Host "    [~] Would rename computer from '$env:COMPUTERNAME' to '$newName'" -ForegroundColor Cyan
            Add-ActionRecord -Step "Computer Rename" -Status "WhatIf" -Detail "Would rename to: $newName"
            $rebootRequired = $true
        } else {
        try {
            Rename-Computer -NewName $newName -Force -ErrorAction Stop
            Write-Host "    [+] Computer will be renamed to '$newName' after reboot." -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Step "Computer Rename" -Status "Pending Reboot" -Detail "New name: $newName"
            $rebootRequired = $true
        }
        catch {
            Write-Host "    [-] Rename failed: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Step "Computer Rename" -Status "Failed" -Detail $_
        }
        }
    }
    else {
        Write-Host "    [!!] Invalid name. Must be 1–15 characters, letters/digits/hyphens, no leading/trailing hyphen." -ForegroundColor $ColorSchema.Warning
        Write-Host "         Rename skipped." -ForegroundColor $ColorSchema.Warning
        Add-ActionRecord -Step "Computer Rename" -Status "Skipped" -Detail "Invalid name entered: $newName"
    }
}
else {
    Write-Host "    Skipped." -ForegroundColor $ColorSchema.Info
    Add-ActionRecord -Step "Computer Rename" -Status "Skipped" -Detail "No name entered"
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: DOMAIN JOIN (ENTRA ID OR LOCAL ACTIVE DIRECTORY)
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[3/7] Domain Join..." -ForegroundColor $ColorSchema.Progress
Write-Host ""
Write-Host "  This step can join this machine to Entra ID (Azure AD) or a local Active Directory domain." -ForegroundColor $ColorSchema.Info
Write-Host ""

if ($Unattended) {
    if ($DomainJoinType -eq "Entra") {
        Write-Host "    [*] Unattended mode: Entra ID join requires interactive credentials — skipped." -ForegroundColor $ColorSchema.Warning
        Add-ActionRecord -Step "Entra ID Join" -Status "Skipped" -Detail "Not supported in unattended mode"
    }
    elseif ($DomainJoinType -eq "AD") {
        if ([string]::IsNullOrWhiteSpace($ADDomain)) {
            Write-Host "    [*] Unattended AD join: -ADDomain not provided — skipped." -ForegroundColor $ColorSchema.Warning
            Add-ActionRecord -Step "AD Join" -Status "Skipped" -Detail "-ADDomain parameter not provided"
        }
        else {
            if ($WhatIf) {
                Write-Host "    [~] Would join domain '$ADDomain'" -ForegroundColor Cyan
                Add-ActionRecord -Step "AD Join" -Status "WhatIf" -Detail "Would join domain: $ADDomain"
                $rebootRequired = $true
            }
            else {
                Write-Host "    [*] Joining Active Directory domain '$ADDomain'..." -ForegroundColor $ColorSchema.Progress
                try {
                    if (-not [string]::IsNullOrWhiteSpace($ADOUPath)) {
                        Add-Computer -DomainName $ADDomain -OUPath $ADOUPath -ErrorAction Stop
                    }
                    else {
                        Add-Computer -DomainName $ADDomain -ErrorAction Stop
                    }
                    Write-Host "    [+] Successfully joined to domain '$ADDomain'." -ForegroundColor $ColorSchema.Success
                    Add-ActionRecord -Step "AD Join" -Status "Joined" -Detail "Domain: $ADDomain"
                    $rebootRequired = $true
                }
                catch {
                    Write-Host "    [-] AD join failed: $_" -ForegroundColor $ColorSchema.Error
                    Add-ActionRecord -Step "AD Join" -Status "Failed" -Detail $_
                    Write-TKError -ScriptName 'covenant' -Message "AD join to '$ADDomain' failed: $($_.Exception.Message)" -Category 'AD Join'
                }
            }
        }
    }
    else {
        Write-Host "    [*] No -DomainJoinType specified — domain join skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Domain Join" -Status "Skipped" -Detail "DomainJoinType not specified"
    }
} else {
    Write-Host "  Select join type:" -ForegroundColor $ColorSchema.Info
    Write-Host "    [1] Entra ID join (Azure AD)" -ForegroundColor $ColorSchema.Info
    Write-Host "    [2] Local Active Directory domain join" -ForegroundColor $ColorSchema.Info
    Write-Host "    [3] Skip" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    $joinTypeChoice = Read-Host "  Enter choice (1/2/3)"

    # ── Entra ID branch ──────────────────────────────────────────────────────
    if ($joinTypeChoice -eq '1') {

        Write-Host ""
        Write-Host "  This will join this machine to your organization's Entra ID (Azure AD) tenant." -ForegroundColor $ColorSchema.Info
        Write-Host "  You will need an account that has permission to join devices." -ForegroundColor $ColorSchema.Info
        Write-Host ""

        if ($WhatIf) {
            Write-Host ""
            Write-Host "  [~] Would collect Entra ID credentials (UPN + password)" -ForegroundColor Cyan
            Write-Host "  [~] Would run: dsregcmd /join to join this device to your Entra ID tenant" -ForegroundColor Cyan
            Write-Host ""
            Add-ActionRecord -Step "Entra ID Join" -Status "WhatIf" -Detail "Would join device to Entra ID"
            $rebootRequired = $true
        } else {

        Write-Host ""
        Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
        Write-Host "   ENTRA ID CREDENTIALS" -ForegroundColor $ColorSchema.Header
        Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
        Write-Host ""

        # Collect UPN
        do {
            $upn = Read-Host "  Enter Entra ID username (UPN, e.g. user@company.com)"
            $upn = $upn.Trim()
            if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                Write-Host "    [!!] That doesn't look like a valid UPN. Please try again." -ForegroundColor $ColorSchema.Warning
                $upn = ""
            }
        } while ([string]::IsNullOrWhiteSpace($upn))

        # Collect password securely
        $securePassword = Read-Host "  Enter password" -AsSecureString

        Write-Host ""
        Write-Host "    Preparing credentials and initiating Entra ID join..." -ForegroundColor $ColorSchema.Progress
        Write-Host "    (Note: if your account requires MFA or Conditional Access, a browser" -ForegroundColor $ColorSchema.Info
        Write-Host "     prompt may appear. Complete it to continue the join.)" -ForegroundColor $ColorSchema.Info
        Write-Host ""

        # Convert SecureString to plain text briefly for credential storage
        $credPtr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $plainPassword = $null
        $joinSuccess = $false

        try {
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($credPtr)

            # Store credentials in Windows Credential Manager so dsregcmd can pick them up
            & cmdkey /add:login.microsoftonline.com /user:$upn /pass:$plainPassword | Out-Null

            # Wipe the plain-text string from memory immediately
            $plainPassword = $null
            [System.GC]::Collect()

            Write-Host "    Running dsregcmd /join ..." -ForegroundColor $ColorSchema.Progress

            # Execute the join and capture output
            $joinOutput = & dsregcmd /join 2>&1

            # Remove the stored credential right away
            & cmdkey /delete:login.microsoftonline.com | Out-Null

            # Verify result
            $verifyOutput = & dsregcmd /status 2>&1
            $nowJoined = ($verifyOutput | Where-Object { $_ -match "AzureAdJoined\s*:\s*YES" }) -ne $null

            if ($nowJoined) {
                Write-Host "    [+] Successfully joined to Entra ID!" -ForegroundColor $ColorSchema.Success
                $tenantLine = $verifyOutput | Select-String "TenantName\s*:\s*(.+)" | Select-Object -First 1
                if ($tenantLine) {
                    $tn = $tenantLine.Matches[0].Groups[1].Value.Trim()
                    Write-Host "    Tenant : $tn" -ForegroundColor $ColorSchema.Success
                }
                Add-ActionRecord -Step "Entra ID Join" -Status "Joined" -Detail "UPN: $upn"
                $rebootRequired = $true
                $joinSuccess = $true
            }
            else {
                # Check if join output contains useful error info
                $errorLine = $joinOutput | Where-Object { $_ -match "error|fail|0x" } | Select-Object -First 1
                Write-Host "    [-] Join did not complete successfully." -ForegroundColor $ColorSchema.Error
                if ($errorLine) {
                    Write-Host "    Detail: $errorLine" -ForegroundColor $ColorSchema.Error
                }
                Write-Host ""
                Write-Host "    [!!] Possible reasons:" -ForegroundColor $ColorSchema.Warning
                Write-Host "         - Incorrect UPN or password" -ForegroundColor $ColorSchema.Warning
                Write-Host "         - Account requires MFA (complete the browser prompt if it appeared)" -ForegroundColor $ColorSchema.Warning
                Write-Host "         - Account lacks 'Join devices' permission in Entra ID" -ForegroundColor $ColorSchema.Warning
                Write-Host "         - Device join limit reached for this user (check Azure AD > Devices)" -ForegroundColor $ColorSchema.Warning
                Add-ActionRecord -Step "Entra ID Join" -Status "Failed" -Detail "UPN: $upn — dsregcmd did not confirm join"
            }
        }
        catch {
            & cmdkey /delete:login.microsoftonline.com 2>$null | Out-Null
            Write-Host "    [-] Unexpected error during join: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Step "Entra ID Join" -Status "Failed" -Detail $_
            Write-TKError -ScriptName 'covenant' -Message "Entra ID join failed: $($_.Exception.Message)" -Category 'Entra Join'
        }
        finally {
            # Always zero out the BSTR pointer
            if ($credPtr -ne [IntPtr]::Zero) {
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($credPtr)
            }
            $plainPassword = $null
            [System.GC]::Collect()
        }
        }  # end else (not WhatIf)
    }

    # ── Local Active Directory branch ────────────────────────────────────────
    elseif ($joinTypeChoice -eq '2') {

        Write-Host ""
        Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
        Write-Host "   ACTIVE DIRECTORY DOMAIN JOIN" -ForegroundColor $ColorSchema.Header
        Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
        Write-Host ""

        # Domain name
        $adDomainInput = ""
        do {
            $adDomainInput = (Read-Host "  Enter domain name (e.g. corp.contoso.com)").Trim()
            if ([string]::IsNullOrWhiteSpace($adDomainInput)) {
                Write-Host "    [!!] Domain name cannot be blank." -ForegroundColor $ColorSchema.Warning
            }
        } while ([string]::IsNullOrWhiteSpace($adDomainInput))

        # Optional OU path
        $adOUInput = (Read-Host "  Enter OU path (leave blank for default OU, e.g. OU=Workstations,DC=corp,DC=contoso,DC=com)").Trim()

        # Credentials
        $adCred = Get-Credential -Message "Enter credentials to join '$adDomainInput' (e.g. CORP\JoinAccount)"

        Write-Host ""
        Write-Host "    Joining Active Directory domain '$adDomainInput'..." -ForegroundColor $ColorSchema.Progress

        if ($WhatIf) {
            Write-Host ""
            Write-Host "    [~] Would join domain '$adDomainInput'" -ForegroundColor Cyan
            Write-Host ""
            Add-ActionRecord -Step "AD Join" -Status "WhatIf" -Detail "Would join domain: $adDomainInput"
            $rebootRequired = $true
        }
        else {
            try {
                if (-not [string]::IsNullOrWhiteSpace($adOUInput)) {
                    Add-Computer -DomainName $adDomainInput -Credential $adCred -OUPath $adOUInput -ErrorAction Stop
                }
                else {
                    Add-Computer -DomainName $adDomainInput -Credential $adCred -ErrorAction Stop
                }
                Write-Host "    [+] Successfully joined to domain '$adDomainInput'." -ForegroundColor $ColorSchema.Success
                Add-ActionRecord -Step "AD Join" -Status "Joined" -Detail "Domain: $adDomainInput"
                $rebootRequired = $true
            }
            catch {
                Write-Host "    [-] AD join failed: $_" -ForegroundColor $ColorSchema.Error
                Add-ActionRecord -Step "AD Join" -Status "Failed" -Detail $_
                Write-TKError -ScriptName 'covenant' -Message "AD join to '$adDomainInput' failed: $($_.Exception.Message)" -Category 'AD Join'
            }
        }
    }

    # ── Skip ─────────────────────────────────────────────────────────────────
    else {
        Write-Host "    Skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Domain Join" -Status "Skipped" -Detail "User chose to skip"
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: NETWORK DRIVE MAPPING
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[4/7] Network Drive Mapping..." -ForegroundColor $ColorSchema.Progress
Write-Host ""

if ($Unattended) {
    Write-Host "    [*] Unattended mode: drive mapping requires interactive input — skipped." -ForegroundColor $ColorSchema.Warning
    Add-ActionRecord -Step "Drive Mapping" -Status "Skipped" -Detail "Not supported in unattended mode"
} else {
    $driveChoice = Read-Host "  Map network drives? (Y/N)"

    if ($driveChoice -eq 'Y' -or $driveChoice -eq 'y') {

        $mapAnother = $true

        while ($mapAnother) {

            Write-Host ""
            Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
            Write-Host "   ADD NETWORK DRIVE" -ForegroundColor $ColorSchema.Header
            Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
            Write-Host ""

            # Drive letter
            $driveLetter = ""
            do {
                $driveLetter = Read-Host "  Drive letter (e.g. Z, Y, H)"
                $driveLetter = $driveLetter.Trim().TrimEnd(':').ToUpper()
                if ($driveLetter -notmatch '^[D-Z]$') {
                    Write-Host "    [!!] Enter a single letter between D and Z." -ForegroundColor $ColorSchema.Warning
                    $driveLetter = ""
                }
            } while ([string]::IsNullOrWhiteSpace($driveLetter))

            # UNC path
            $uncPath = ""
            do {
                $uncPath = Read-Host "  UNC path (e.g. \\server\share)"
                $uncPath = $uncPath.Trim()
                if ($uncPath -notmatch '^\\\\[^\\]+\\[^\\]+') {
                    Write-Host "    [!!] Path must be in \\server\share format." -ForegroundColor $ColorSchema.Warning
                    $uncPath = ""
                }
            } while ([string]::IsNullOrWhiteSpace($uncPath))

            # Optional: credentials for the share
            Write-Host ""
            $useShareCreds = Read-Host "  Use specific credentials for this share? (Y/N)"
            $shareCred = $null

            if ($useShareCreds -eq 'Y' -or $useShareCreds -eq 'y') {
                $shareUser = Read-Host "  Share username (e.g. DOMAIN\user or user@domain.com)"
                $sharePass = Read-Host "  Share password" -AsSecureString
                $shareCred = New-Object System.Management.Automation.PSCredential($shareUser.Trim(), $sharePass)
            }

            # Persistent mapping
            $persistChoice = Read-Host "  Make this mapping persistent across reboots? (Y/N)"
            $persist = ($persistChoice -eq 'Y' -or $persistChoice -eq 'y')

            Write-Host ""
            Write-Host "    Mapping $driveLetter`: to $uncPath ..." -ForegroundColor $ColorSchema.Progress

            if ($WhatIf) {
                $persistLabel = if ($persist) { "persistent" } else { "non-persistent" }
                Write-Host ""
                Write-Host "  [~] Would map $driveLetter`: to $uncPath ($persistLabel)" -ForegroundColor Cyan
                Write-Host ""
                Add-ActionRecord -Step "Drive Mapping" -Status "WhatIf" -Detail "$driveLetter`: → $uncPath (Persistent: $persist)"
            } else {

            try {
                # Remove existing mapping on that letter if present
                if (Test-Path "$driveLetter`:") {
                    Remove-PSDrive -Name $driveLetter -Force -ErrorAction SilentlyContinue
                    & net use "$driveLetter`:" /delete /yes 2>$null | Out-Null
                }

                if ($shareCred) {
                    $credPtr2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($shareCred.Password)
                    try {
                        $sharePlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($credPtr2)
                        if ($persist) {
                            & net use "$driveLetter`:" $uncPath $sharePlain /user:$($shareCred.UserName) /persistent:yes 2>&1 | Out-Null
                        } else {
                            & net use "$driveLetter`:" $uncPath $sharePlain /user:$($shareCred.UserName) /persistent:no  2>&1 | Out-Null
                        }
                        $sharePlain = $null
                    }
                    finally {
                        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($credPtr2)
                    }
                }
                else {
                    if ($persist) {
                        & net use "$driveLetter`:" $uncPath /persistent:yes 2>&1 | Out-Null
                    } else {
                        & net use "$driveLetter`:" $uncPath /persistent:no  2>&1 | Out-Null
                    }
                }

                if ($LASTEXITCODE -eq 0) {
                    Write-Host "    [+] $driveLetter`: mapped to $uncPath" -ForegroundColor $ColorSchema.Success
                    Add-ActionRecord -Step "Drive Mapping" -Status "Mapped" -Detail "$driveLetter`: → $uncPath (Persistent: $persist)"
                }
                else {
                    Write-Host "    [-] Mapping failed (exit code $LASTEXITCODE). Check path and credentials." -ForegroundColor $ColorSchema.Error
                    Add-ActionRecord -Step "Drive Mapping" -Status "Failed" -Detail "$driveLetter`: → $uncPath"
                }
            }
            catch {
                Write-Host "    [-] Error mapping drive: $_" -ForegroundColor $ColorSchema.Error
                Add-ActionRecord -Step "Drive Mapping" -Status "Failed" -Detail $_
            }
            }  # end else (not WhatIf)

            Write-Host ""
            $anotherChoice = Read-Host "  Map another drive? (Y/N)"
            $mapAnother = ($anotherChoice -eq 'Y' -or $anotherChoice -eq 'y')
        }
    }
    else {
        Write-Host "    Skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Drive Mapping" -Status "Skipped" -Detail "User chose to skip"
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5: LOCAL ADMIN ACCOUNT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[5/7] Local Administrator Account..." -ForegroundColor $ColorSchema.Progress
Write-Host ""

if ($Unattended) {
    if (-not [string]::IsNullOrWhiteSpace($LocalAdminUser) -and $LocalAdminPassword -ne $null) {
        Write-Host "    [*] Creating local admin account '$LocalAdminUser'..." -ForegroundColor $ColorSchema.Progress
        try {
            $existing = Get-LocalUser -Name $LocalAdminUser -ErrorAction SilentlyContinue
            if ($existing) {
                Set-LocalUser -Name $LocalAdminUser -Password $LocalAdminPassword -ErrorAction Stop
                Add-ActionRecord -Step "Local Admin" -Status "Updated" -Detail "Password reset for: $LocalAdminUser"
            } else {
                New-LocalUser -Name $LocalAdminUser -Password $LocalAdminPassword -PasswordNeverExpires $true -AccountNeverExpires $true -ErrorAction Stop | Out-Null
                Add-LocalGroupMember -Group "Administrators" -Member $LocalAdminUser -ErrorAction Stop
                Write-Host "    [+] Account '$LocalAdminUser' created and added to Administrators." -ForegroundColor $ColorSchema.Success
                Add-ActionRecord -Step "Local Admin" -Status "Created" -Detail "Username: $LocalAdminUser"
            }
        } catch {
            Write-Host "    [-] Error: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Step "Local Admin" -Status "Failed" -Detail $_
        }
    } else {
        Write-Host "    [*] Unattended mode: no -LocalAdminUser or -LocalAdminPassword provided — skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Local Admin" -Status "Skipped" -Detail "Parameters not provided"
    }
} else {
    $adminChoice = Read-Host "  Create a local administrator account? (Y/N)"

    if ($adminChoice -eq 'Y' -or $adminChoice -eq 'y') {

        Write-Host ""
        Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
        Write-Host "   NEW LOCAL ADMIN ACCOUNT" -ForegroundColor $ColorSchema.Header
        Write-Host "  ─────────────────────────────────────────" -ForegroundColor $ColorSchema.Header
        Write-Host ""

        # Username
        $localUser = ""
        do {
            $localUser = Read-Host "  Username"
            $localUser = $localUser.Trim()
            if ([string]::IsNullOrWhiteSpace($localUser)) {
                Write-Host "    [!!] Username cannot be blank." -ForegroundColor $ColorSchema.Warning
            }
            elseif ($localUser.Length -gt 20) {
                Write-Host "    [!!] Username must be 20 characters or fewer." -ForegroundColor $ColorSchema.Warning
                $localUser = ""
            }
        } while ([string]::IsNullOrWhiteSpace($localUser))

        # Password
        $localPass = Read-Host "  Password" -AsSecureString

        # Full name (optional label)
        $fullName = Read-Host "  Full name (optional, press Enter to skip)"

        Write-Host ""
        Write-Host "    Creating account '$localUser' ..." -ForegroundColor $ColorSchema.Progress

        if ($WhatIf) {
            Write-Host ""
            Write-Host "  [~] Would create local admin account '$localUser'" -ForegroundColor Cyan
            Write-Host "  [~] Would add '$localUser' to the Administrators group" -ForegroundColor Cyan
            Write-Host ""
            Add-ActionRecord -Step "Local Admin" -Status "WhatIf" -Detail "Would create: $localUser"
        } else {

        try {
            $existing = Get-LocalUser -Name $localUser -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Host "    [!!] User '$localUser' already exists. Updating password instead." -ForegroundColor $ColorSchema.Warning
                Set-LocalUser -Name $localUser -Password $localPass -ErrorAction Stop
                Add-ActionRecord -Step "Local Admin" -Status "Updated" -Detail "Password reset for existing user: $localUser"
            }
            else {
                $newUserParams = @{
                    Name                 = $localUser
                    Password             = $localPass
                    PasswordNeverExpires = $true
                    AccountNeverExpires  = $true
                }
                if (-not [string]::IsNullOrWhiteSpace($fullName)) {
                    $newUserParams['FullName'] = $fullName.Trim()
                }

                New-LocalUser @newUserParams -ErrorAction Stop | Out-Null
                Add-LocalGroupMember -Group "Administrators" -Member $localUser -ErrorAction Stop

                Write-Host "    [+] Account '$localUser' created and added to Administrators." -ForegroundColor $ColorSchema.Success
                Add-ActionRecord -Step "Local Admin" -Status "Created" -Detail "Username: $localUser"
            }
        }
        catch {
            Write-Host "    [-] Error creating account: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Step "Local Admin" -Status "Failed" -Detail $_
        }
        }  # end else (not WhatIf)
    }
    else {
        Write-Host "    Skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Local Admin" -Status "Skipped" -Detail "User chose to skip"
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: TIMEZONE CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
#   (was Step 6/6, now Step 6/7)
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[6/7] Timezone Configuration..." -ForegroundColor $ColorSchema.Progress
Write-Host "    Current timezone: $((Get-TimeZone).DisplayName)" -ForegroundColor $ColorSchema.Info
Write-Host ""

if ($Unattended) {
    if (-not [string]::IsNullOrWhiteSpace($Timezone)) {
        try {
            Set-TimeZone -Id $Timezone -ErrorAction Stop
            Write-Host "    [+] Timezone set to: $((Get-TimeZone).DisplayName)" -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Step "Timezone" -Status "Set" -Detail $Timezone
        } catch {
            Write-Host "    [-] Failed to set timezone '$Timezone': $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Step "Timezone" -Status "Failed" -Detail $_
        }
    } else {
        Write-Host "    [*] No -Timezone provided — skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Timezone" -Status "Skipped" -Detail "Parameter not provided"
    }
} else {
    $tzChoice = Read-Host "  Change timezone? (Y/N)"

    if ($tzChoice -eq 'Y' -or $tzChoice -eq 'y') {

        Write-Host ""
        Write-Host "  Common timezones:" -ForegroundColor $ColorSchema.Header
        $commonZones = @(
            @{ Id = "Eastern Standard Time";   Label = " [1] Eastern (ET)   — New York, Atlanta" }
            @{ Id = "Central Standard Time";   Label = " [2] Central (CT)   — Chicago, Dallas" }
            @{ Id = "Mountain Standard Time";  Label = " [3] Mountain (MT)  — Denver, Phoenix" }
            @{ Id = "Pacific Standard Time";   Label = " [4] Pacific (PT)   — Los Angeles, Seattle" }
            @{ Id = "UTC";                     Label = " [5] UTC" }
            @{ Id = "GMT Standard Time";       Label = " [6] GMT             — London (no DST)" }
            @{ Id = "Central Europe Standard Time"; Label = " [7] CET             — Berlin, Paris" }
            @{ Id = "Tokyo Standard Time";     Label = " [8] JST             — Tokyo" }
            @{ Id = "";                        Label = " [0] Enter manually" }
        )

        foreach ($z in $commonZones) { Write-Host "  $($z.Label)" -ForegroundColor $ColorSchema.Info }

        Write-Host ""
        $tzSel = Read-Host "  Enter choice (1–8 or 0 for manual)"

        $selectedTzId = ""
        if ($tzSel -match '^[1-8]$') {
            $idx = [int]$tzSel - 1
            $selectedTzId = $commonZones[$idx].Id
        }
        elseif ($tzSel -eq '0') {
            Write-Host "  Tip: Run 'Get-TimeZone -ListAvailable | Select-Object Id, DisplayName' for all IDs." -ForegroundColor $ColorSchema.Info
            $selectedTzId = (Read-Host "  Enter timezone ID exactly").Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($selectedTzId)) {
            try {
                Set-TimeZone -Id $selectedTzId -ErrorAction Stop
                Write-Host "    [+] Timezone set to: $((Get-TimeZone).DisplayName)" -ForegroundColor $ColorSchema.Success
                Add-ActionRecord -Step "Timezone" -Status "Set" -Detail $selectedTzId
            }
            catch {
                Write-Host "    [-] Failed to set timezone '$selectedTzId': $_" -ForegroundColor $ColorSchema.Error
                Add-ActionRecord -Step "Timezone" -Status "Failed" -Detail $_
            }
        }
        else {
            Write-Host "    [!!] No timezone selected. Skipped." -ForegroundColor $ColorSchema.Warning
            Add-ActionRecord -Step "Timezone" -Status "Skipped" -Detail "No valid selection"
        }
    }
    else {
        Write-Host "    Skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "Timezone" -Status "Skipped" -Detail "User chose to skip"
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7: GROUP POLICY SYNCHRONISATION
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[7/7] Group Policy Synchronisation..." -ForegroundColor $ColorSchema.Progress
Write-Host ""

if ($SkipGPUpdate) {
    Write-Host "    [*] -SkipGPUpdate specified — skipped." -ForegroundColor $ColorSchema.Info
    Add-ActionRecord -Step "GP Sync" -Status "Skipped" -Detail "-SkipGPUpdate switch set"
}
elseif ($WhatIf) {
    Write-Host "    [~] Would run: gpupdate /force" -ForegroundColor Cyan
    Add-ActionRecord -Step "GP Sync" -Status "WhatIf" -Detail "Would run gpupdate /force"
}
elseif ($Unattended) {
    Write-Host "    [*] Running gpupdate /force..." -ForegroundColor $ColorSchema.Progress
    try {
        $gpOutput = & gpupdate /force 2>&1
        Write-Host "    [+] Group Policy update completed." -ForegroundColor $ColorSchema.Success
        Add-ActionRecord -Step "GP Sync" -Status "Completed" -Detail "gpupdate /force ran successfully"
    }
    catch {
        Write-Host "    [-] gpupdate failed: $_" -ForegroundColor $ColorSchema.Error
        Add-ActionRecord -Step "GP Sync" -Status "Failed" -Detail $_
    }
}
else {
    $gpChoice = Read-Host "  Run gpupdate /force now? (Y/N)"

    if ($gpChoice -eq 'Y' -or $gpChoice -eq 'y') {
        Write-Host "    [*] Running gpupdate /force..." -ForegroundColor $ColorSchema.Progress
        try {
            $gpOutput = & gpupdate /force 2>&1
            Write-Host "    [+] Group Policy update completed." -ForegroundColor $ColorSchema.Success
            Add-ActionRecord -Step "GP Sync" -Status "Completed" -Detail "gpupdate /force ran successfully"
        }
        catch {
            Write-Host "    [-] gpupdate failed: $_" -ForegroundColor $ColorSchema.Error
            Add-ActionRecord -Step "GP Sync" -Status "Failed" -Detail $_
        }
    }
    else {
        Write-Host "    Skipped." -ForegroundColor $ColorSchema.Info
        Add-ActionRecord -Step "GP Sync" -Status "Skipped" -Detail "User chose to skip"
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "     ONBOARDING SUMMARY" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host ""

foreach ($record in $ActionLog) {
    $color = switch -Regex ($record.Status) {
        'Joined|Created|Mapped|Set|Updated' { $ColorSchema.Success }
        'Skipped'                           { $ColorSchema.Info    }
        'Pending Reboot'                    { $ColorSchema.Warning }
        'WhatIf'                            { 'Cyan'               }
        default                             { $ColorSchema.Error   }
    }
    $detail = if ($record.Detail) { " — $($record.Detail)" } else { "" }
    Write-Host ("  {0,-20} [{1}]{2}" -f $record.Step, $record.Status, $detail) -ForegroundColor $color
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# REBOOT DECISION
# ─────────────────────────────────────────────────────────────────────────────

if ($rebootRequired) {
    Write-Host "  *** A REBOOT IS REQUIRED to apply changes (rename, Entra join). ***" -ForegroundColor $ColorSchema.Warning
    Write-Host ""

    if ($Unattended) {
        Write-Host "  [*] Unattended mode: reboot required. Reboot this machine when ready." -ForegroundColor $ColorSchema.Warning
        Write-Host "       Command: Restart-Computer" -ForegroundColor $ColorSchema.Info
    } else {
        $rebootPrompt = Read-Host "  Is it safe to reboot this computer now? (Y/N)"

        if ($rebootPrompt -eq 'Y' -or $rebootPrompt -eq 'y') {
            if ($WhatIf) {
                Write-Host ""
                Write-Host "  [~] Would restart this computer now" -ForegroundColor Cyan
                Write-Host ""
            } else {
            Write-Host ""
            Write-Host "  Rebooting in 30 seconds. Press Escape to cancel..." -ForegroundColor $ColorSchema.Warning
            Write-Host ""
            Write-Host "   30 [============================================]" -ForegroundColor $ColorSchema.Accent

            $cancelled = $false
            for ($i = 30; $i -gt 0; $i--) {
                $progress  = [math]::Floor((30 - $i) / 30 * 44)
                $bar       = "=" * $progress
                $remaining = " " * (44 - $progress)
                Write-Host -NoNewline "`r   $i  [$bar$remaining]" -ForegroundColor $ColorSchema.Accent

                for ($tick = 0; $tick -lt 10; $tick++) {
                    if ([Console]::KeyAvailable) {
                        $key = [Console]::ReadKey($true)
                        if ($key.Key -eq [ConsoleKey]::Escape) { $cancelled = $true; break }
                    }
                    Start-Sleep -Milliseconds 100
                }
                if ($cancelled) { break }
            }

            Write-Host ""
            Write-Host ""

            if ($cancelled) {
                Write-Host "  Reboot cancelled." -ForegroundColor $ColorSchema.Warning
                Write-Host "  [!!] Reboot when ready: Start-Menu → Power → Restart" -ForegroundColor $ColorSchema.Warning
            }
            else {
                Restart-Computer -Force
            }
            }  # end else (not WhatIf)
        }
        else {
            Write-Host ""
            Write-Host "  [!!] Remember to reboot before using this machine on the domain." -ForegroundColor $ColorSchema.Warning
            Write-Host "       Command: Restart-Computer" -ForegroundColor $ColorSchema.Info
        }
    }
}
else {
    Write-Host "  [+] No reboot required. All changes are active immediately." -ForegroundColor $ColorSchema.Success
}

Write-Host ""
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "  SCRIPT EXECUTION COMPLETED" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
