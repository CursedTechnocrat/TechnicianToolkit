# catacomb.ps1 - C.A.T.A.C.O.M.B. — Catalogs Access To All Content: Owners, Members & Breadth
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
    C.A.T.A.C.O.M.B. — Catalogs Access To All Content: Owners, Members & Breadth
    File Share & NTFS Permissions Review Tool for PowerShell 5.1+

.DESCRIPTION
    Answers "who has access to this share?" on a file server, the way HERALD
    answers it for the domain and WARD for the local machine:

      - Every non-administrative SMB share with its share permissions and the
        NTFS permissions on its root
      - A walk down each share (default two folder levels) recording every
        folder with explicit permissions or broken inheritance
      - Everyone / Authenticated Users / Users / Domain Users granted write
        where the share permissions let it through -- the combination that
        actually lets anyone in the organisation change or encrypt the data
      - Permissions granted straight to user accounts instead of groups,
        entries for deleted accounts (bare SIDs), and Deny entries

    Produces an HTML report and a CSV of every access-control entry scanned,
    ready for a customer access review. Read-only -- no permission is changed.

.USAGE
    PS C:\> .\catacomb.ps1                         # Interactive: all shares, two levels deep
    PS C:\> .\catacomb.ps1 -Unattended             # Silent: all shares + HTML + CSV
    PS C:\> .\catacomb.ps1 -Depth 4                # Walk four folder levels into each share
    PS C:\> .\catacomb.ps1 -Path 'D:\Data'         # Review one folder tree instead of the shares

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [ValidateScript({ [string]::IsNullOrWhiteSpace($_) -or (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]$Path,
    [ValidateRange(0, 10)]
    [int]$Depth = 2,
    [switch]$Transcript
)

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
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Principals that mean "effectively everyone". Domain Users is matched by its
# relative ID (-513) because the domain part of the SID differs per domain.
$BroadSids = [ordered]@{
    'S-1-1-0'      = 'Everyone'
    'S-1-5-7'      = 'Anonymous Logon'
    'S-1-5-11'     = 'Authenticated Users'
    'S-1-5-32-545' = 'BUILTIN\Users'
    'S-1-5-32-546' = 'BUILTIN\Guests'
}
$BroadDomainRids = @{ '513' = 'Domain Users'; '514' = 'Domain Guests' }

# Shares whose broad read access is by design.
$ExpectedBroadReadShares = @('NETLOGON', 'SYSVOL')

# A walk past this many folders stops, so one enormous share cannot run for hours.
$MaxFolders = 5000

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$CatacombFindings = @{
    'BroadWriteAccess' = @{
        Severity = 'Error'
        Title    = 'Everyone-type group can write'
        Summary  = 'Both the share and NTFS permissions let Everyone, Authenticated Users, Users or Domain Users change files -- so any account in the organisation, including a compromised one running ransomware, can modify or encrypt them.'
        Remedy   = 'Grant write to the specific groups that need it and reduce the broad group to read (or remove it).'
    }
    'BroadReadAccess' = @{
        Severity = 'Warning'
        Title    = 'Everyone-type group can read'
        Summary  = 'Any account in the organisation can read this share. Fine for public material; not for HR, finance or client data.'
        Remedy   = 'Confirm the content is meant for everyone. If not, replace the broad group with the teams that need it.'
    }
    'DirectUserAce' = @{
        Severity = 'Warning'
        Title    = 'Permissions granted directly to user accounts'
        Summary  = 'Access given to named users instead of groups is invisible in group membership reviews and is left behind when people change roles.'
        Remedy   = 'Move each user into a group that holds the access, then remove the user entry.'
    }
    'OrphanedSid' = @{
        Severity = 'Warning'
        Title    = 'Permissions for deleted accounts'
        Summary  = 'These entries name a SID that no longer resolves -- the account or group was deleted but its access stayed.'
        Remedy   = 'Remove the orphaned entries (they grant nothing today, but hide what the ACL really says).'
    }
    'InheritanceBroken' = @{
        Severity = 'Info'
        Title    = 'Folders with inheritance disabled'
        Summary  = 'These folders do not take permissions from their parent, so a change at the share root will not reach them.'
        Remedy   = 'Expected for deliberately restricted folders. Review any you cannot explain.'
    }
    'DenyAce' = @{
        Severity = 'Info'
        Title    = 'Deny entries present'
        Summary  = 'Deny entries override allows and are a frequent cause of "I can''t open it" tickets.'
        Remedy   = 'Prefer removing the allow over adding a deny; keep denies documented.'
    }
    'FolderUnreadable' = @{
        Severity = 'Info'
        Title    = 'Folders C.A.T.A.C.O.M.B. could not read'
        Summary  = 'Access was denied even to the elevated session, so these folders were not reviewed.'
        Remedy   = 'Take ownership only if you are authorised to; otherwise review them with their owner.'
    }
    'ScanTruncated' = @{
        Severity = 'Warning'
        Title    = 'Folder walk stopped early'
        Summary  = 'The walk reached its folder limit ($MaxFolders), so deeper folders were not reviewed.'
        Remedy   = 'Rerun against the share with -Path and a smaller -Depth, or split the review by top-level folder.'
    }
    'NoShares' = @{
        Severity = 'Info'
        Title    = 'No file shares on this machine'
        Summary  = 'Only the administrative shares (C$, ADMIN$, IPC$) exist.'
        Remedy   = 'Run on the file server, or pass -Path to review a folder tree.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-CatacombFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Subject = '', [string]$Detail = '')
    $meta = $CatacombFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Subject = $Subject; Detail = $Detail
    })
}

function Get-SeverityClass {
    param([string]$Severity)
    switch ($Severity) {
        'Error'   { return 'err'  }
        'Warning' { return 'warn' }
        default   { return 'info' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-CatacombBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗ █████╗ ████████╗ █████╗  ██████╗ ██████╗ ███╗   ███╗██████╗
  ██╔════╝██╔══██╗╚══██╔══╝██╔══██╗██╔════╝██╔═══██╗████╗ ████║██╔══██╗
  ██║     ███████║   ██║   ███████║██║     ██║   ██║██╔████╔██║██████╔╝
  ██║     ██╔══██║   ██║   ██╔══██║██║     ██║   ██║██║╚██╔╝██║██╔══██╗
  ╚██████╗██║  ██║   ██║   ██║  ██║╚██████╗╚██████╔╝██║ ╚═╝ ██║██████╔╝
   ╚═════╝╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝ ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═════╝

"@ -ForegroundColor Cyan
    Write-Host "    C.A.T.A.C.O.M.B. — Catalogs Access To All Content: Owners, Members & Breadth" -ForegroundColor Cyan
    Write-Host "    File Share & NTFS Permissions Review Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Get-NtfsRightsLevel {
    # Collapses a FileSystemRights mask to Full / Modify / Write / Read /
    # Special. Generic rights (GENERIC_ALL etc.) appear on inheritable ACEs
    # and are mapped to what they grant.
    param([long]$Value)
    $v = $Value
    if ($v -lt 0) { $v = $v + 4294967296 }
    if (($v -band 0x10000000) -ne 0) { return 'Full' }
    if (($v -band 0x1F01FF) -eq 0x1F01FF) { return 'Full' }
    if (($v -band 0x301BF) -eq 0x301BF) { return 'Modify' }
    if (($v -band 0x40000000) -ne 0 -or ($v -band 0x6) -ne 0) { return 'Write' }
    if (($v -band 0x80000000) -ne 0 -or ($v -band 0x20000000) -ne 0 -or ($v -band 0x1) -ne 0) { return 'Read' }
    return 'Special'
}

function Test-WriteLevel {
    param([string]$Level)
    return $Level -in @('Full', 'Modify', 'Write', 'Change')
}

function Get-SidCategory {
    # 'Broad' for everyone-type principals, 'WellKnown' for the other built-in
    # and service SIDs, 'Account' for anything a directory or SAM issued.
    param([string]$Sid)
    if ($BroadSids.Contains($Sid)) { return 'Broad' }
    if ($Sid -match '^S-1-5-21-\d+-\d+-\d+-(\d+)$' -and $BroadDomainRids.ContainsKey($Matches[1])) { return 'Broad' }
    if ($Sid -match '^S-1-5-21-') { return 'Account' }
    return 'WellKnown'
}

function Get-CatacombVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Exposed'; Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Review';  Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Tidy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL RESOLUTION (cached)
# ─────────────────────────────────────────────────────────────────────────────

$PrincipalCache = @{}

function Resolve-Principal {
    # Name and kind for a SID: Broad / WellKnown / User / Group / Computer /
    # Orphaned / Account (resolved, class unknown). Domain classes come from
    # an LDAP bind by SID; local ones from the SAM.
    param([string]$Sid)
    if ($PrincipalCache.ContainsKey($Sid)) { return $PrincipalCache[$Sid] }

    $category = Get-SidCategory -Sid $Sid
    $name = $Sid
    try { $name = ([System.Security.Principal.SecurityIdentifier]$Sid).Translate([System.Security.Principal.NTAccount]).Value } catch { $name = $null }

    $kind = $category
    if ($category -eq 'Broad' -and $BroadSids.Contains($Sid)) { $name = $BroadSids[$Sid] }
    if ($category -eq 'Account') {
        if (-not $name) {
            $kind = 'Orphaned'
        } else {
            $kind = 'Account'
            if (Get-LocalUser -SID $Sid -ErrorAction SilentlyContinue)  { $kind = 'User' }
            elseif (Get-LocalGroup -SID $Sid -ErrorAction SilentlyContinue) { $kind = 'Group' }
            else {
                try {
                    $entry   = [adsi]"LDAP://<SID=$Sid>"
                    $classes = @($entry.Properties['objectClass'] | ForEach-Object { "$_" })
                    if ($classes -contains 'computer')   { $kind = 'Computer' }
                    elseif ($classes -contains 'group')  { $kind = 'Group' }
                    elseif ($classes -contains 'user')   { $kind = 'User' }
                } catch {
                    $kind = 'Account'
                }
            }
        }
    }
    $result = [PSCustomObject]@{ Sid = $Sid; Name = if ($name) { $name } else { $Sid }; Kind = $kind }
    $PrincipalCache[$Sid] = $result
    return $result
}

function ConvertTo-SidString {
    param($Identity)
    try {
        if ($Identity -is [System.Security.Principal.SecurityIdentifier]) { return $Identity.Value }
        return ([System.Security.Principal.NTAccount]"$Identity").Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch {
        # A bare SID string is how Windows shows an entry it cannot resolve.
        if ("$Identity" -match '^S-1-[\d-]+$') { return "$Identity" }
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-CatacombTargets {
    if ($Path) {
        return @([PSCustomObject]@{ Name = $Path; Path = $Path; ShareAccess = $null; Description = 'Folder tree (-Path)' })
    }
    $shares = @()
    try {
        $shares = @(Get-SmbShare -ErrorAction Stop | Where-Object { -not $_.Special -and "$($_.ShareType)" -eq 'FileSystemDirectory' -and $_.Path })
    } catch {
        Write-Fail "Get-SmbShare failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'catacomb' -Message "Get-SmbShare failed: $($_.Exception.Message)" -Category 'Shares'
    }
    $targets = foreach ($s in $shares) {
        $access = @()
        try {
            $access = @(Get-SmbShareAccess -Name $s.Name -ErrorAction Stop | ForEach-Object {
                $sid = ConvertTo-SidString -Identity $_.AccountName
                [PSCustomObject]@{
                    Account = "$($_.AccountName)"
                    Sid     = $sid
                    Broad   = [bool]($sid -and (Get-SidCategory -Sid $sid) -eq 'Broad')
                    Right   = "$($_.AccessRight)"
                    Type    = "$($_.AccessControlType)"
                }
            })
        } catch {
            $access = @()
        }
        [PSCustomObject]@{ Name = $s.Name; Path = $s.Path; ShareAccess = $access; Description = "$($s.Description)" }
    }
    return @($targets)
}

function Get-FolderAces {
    # Returns the ACL's rules as rows, or $null when the folder cannot be read.
    param([string]$Folder)
    try {
        $acl = Get-Acl -LiteralPath $Folder -ErrorAction Stop
    } catch {
        return $null
    }
    $rows = foreach ($r in $acl.Access) {
        $sid = ConvertTo-SidString -Identity $r.IdentityReference
        if (-not $sid) { continue }
        $p = Resolve-Principal -Sid $sid
        [PSCustomObject]@{
            Folder    = $Folder
            Identity  = $p.Name
            Sid       = $sid
            Kind      = $p.Kind
            Rights    = "$($r.FileSystemRights)"
            Level     = Get-NtfsRightsLevel -Value ([long]$r.FileSystemRights.value__)
            Type      = "$($r.AccessControlType)"
            Inherited = [bool]$r.IsInherited
        }
    }
    return [PSCustomObject]@{ Protected = [bool]$acl.AreAccessRulesProtected; Owner = "$($acl.Owner)"; Aces = @($rows) }
}

function Invoke-CatacombWalk {
    # Breadth-first to -Depth, recording the root and every folder that has
    # explicit entries or broken inheritance. Inherited-only folders repeat
    # their parent and are not stored.
    param([object]$Target, [ref]$Budget)

    $interesting = [System.Collections.Generic.List[object]]::new()
    $unreadable  = [System.Collections.Generic.List[string]]::new()
    $queue = [System.Collections.Generic.Queue[object]]::new()
    $queue.Enqueue(@($Target.Path, 0))
    $truncated = $false

    while ($queue.Count -gt 0) {
        $item   = $queue.Dequeue()
        $folder = $item[0]; $level = $item[1]
        if ($Budget.Value -le 0) { $truncated = $true; break }
        $Budget.Value = $Budget.Value - 1

        $acl = Get-FolderAces -Folder $folder
        if ($null -eq $acl) { [void]$unreadable.Add($folder); continue }

        $explicit = @($acl.Aces | Where-Object { -not $_.Inherited })
        if ($level -eq 0 -or $acl.Protected -or $explicit.Count -gt 0) {
            [void]$interesting.Add([PSCustomObject]@{
                Folder = $folder; Level = $level; Protected = $acl.Protected; Owner = $acl.Owner
                Aces = $acl.Aces; Explicit = $explicit
            })
        }

        if ($level -lt $Depth) {
            try {
                foreach ($child in @(Get-ChildItem -LiteralPath $folder -Directory -Force -ErrorAction Stop)) {
                    if (($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    $queue.Enqueue(@($child.FullName, ($level + 1)))
                }
            } catch {
                [void]$unreadable.Add($folder)
            }
        }
    }
    return [PSCustomObject]@{ Folders = @($interesting); Unreadable = @($unreadable | Select-Object -Unique); Truncated = $truncated }
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-CatacombAudit {
    Write-Section "SHARES"
    $targets = Get-CatacombTargets
    if ($targets.Count -eq 0) {
        Add-CatacombFinding -Code 'NoShares' -Subject $env:COMPUTERNAME
        Write-Info "No file shares to review."
        return [PSCustomObject]@{ Shares = @(); AceRows = @() }
    }
    foreach ($t in $targets) { Write-Info ("{0,-24} {1}" -f $t.Name, $t.Path) }

    $budget  = $MaxFolders
    $results = [System.Collections.Generic.List[object]]::new()
    $aceRows = [System.Collections.Generic.List[object]]::new()

    foreach ($t in $targets) {
        Write-Section "SHARE: $($t.Name)"
        Write-Step "Walking $($t.Path) ($Depth level(s))..."
        $walk = Invoke-CatacombWalk -Target $t -Budget ([ref]$budget)
        Write-Info ("{0} folder(s) with their own permissions, {1} unreadable" -f $walk.Folders.Count, $walk.Unreadable.Count)

        foreach ($f in $walk.Folders) {
            foreach ($a in $f.Aces) {
                [void]$aceRows.Add([PSCustomObject]@{
                    Share = $t.Name; Folder = $f.Folder; Identity = $a.Identity; Kind = $a.Kind; Level = $a.Level
                    Rights = $a.Rights; Type = $a.Type; Inherited = $a.Inherited; InheritanceDisabled = $f.Protected
                })
            }
        }

        # Share permissions gate NTFS: a broad write only lands if the share
        # also lets a broad principal change. With -Path there is no share.
        $shareBroadWrite = ($null -eq $t.ShareAccess) -or @($t.ShareAccess | Where-Object { $_.Broad -and $_.Type -eq 'Allow' -and (Test-WriteLevel $_.Right) }).Count -gt 0
        $shareBroadRead  = ($null -eq $t.ShareAccess) -or @($t.ShareAccess | Where-Object { $_.Broad -and $_.Type -eq 'Allow' }).Count -gt 0

        $root = $walk.Folders | Where-Object { $_.Level -eq 0 } | Select-Object -First 1
        $broadWrite = @($walk.Folders | ForEach-Object { $folder = $_.Folder; $_.Aces | Where-Object {
            $_.Kind -eq 'Broad' -and $_.Type -eq 'Allow' -and (Test-WriteLevel $_.Level) } | ForEach-Object { "$($_.Identity) $($_.Level) on $folder" } } | Select-Object -Unique)
        $broadRead = @()
        if ($root) {
            $broadRead = @($root.Aces | Where-Object { $_.Kind -eq 'Broad' -and $_.Type -eq 'Allow' -and -not (Test-WriteLevel $_.Level) } | ForEach-Object { "$($_.Identity) $($_.Level)" } | Select-Object -Unique)
        }

        if ($shareBroadWrite -and $broadWrite.Count -gt 0) {
            Add-CatacombFinding -Code 'BroadWriteAccess' -Subject $t.Name -Detail (($broadWrite | Select-Object -First 10) -join '; ')
        }
        if ($shareBroadRead -and $broadRead.Count -gt 0 -and $t.Name -notin $ExpectedBroadReadShares) {
            Add-CatacombFinding -Code 'BroadReadAccess' -Subject $t.Name -Detail ($broadRead -join '; ')
        }

        $direct = @($walk.Folders | ForEach-Object { $_.Explicit } | Where-Object { $_.Kind -eq 'User' -and $_.Type -eq 'Allow' })
        if ($direct.Count -gt 0) {
            Add-CatacombFinding -Code 'DirectUserAce' -Subject $t.Name -Detail ("{0} entr(ies): {1}" -f $direct.Count, (@($direct | Select-Object -First 10 | ForEach-Object { "$($_.Identity) on $($_.Folder)" }) -join '; '))
        }
        $orphans = @($walk.Folders | ForEach-Object { $_.Aces } | Where-Object { $_.Kind -eq 'Orphaned' } | ForEach-Object { $_.Sid } | Select-Object -Unique)
        if ($orphans.Count -gt 0) { Add-CatacombFinding -Code 'OrphanedSid' -Subject $t.Name -Detail ($orphans -join ', ') }
        $protected = @($walk.Folders | Where-Object { $_.Protected -and $_.Level -gt 0 })
        if ($protected.Count -gt 0) {
            Add-CatacombFinding -Code 'InheritanceBroken' -Subject $t.Name -Detail ("{0} folder(s): {1}" -f $protected.Count, (@($protected | Select-Object -First 10 | ForEach-Object { $_.Folder }) -join '; '))
        }
        $denies = @($walk.Folders | ForEach-Object { $_.Explicit } | Where-Object { $_.Type -eq 'Deny' })
        if ($denies.Count -gt 0) {
            Add-CatacombFinding -Code 'DenyAce' -Subject $t.Name -Detail (@($denies | Select-Object -First 10 | ForEach-Object { "$($_.Identity) on $($_.Folder)" }) -join '; ')
        }
        if ($walk.Unreadable.Count -gt 0) {
            Add-CatacombFinding -Code 'FolderUnreadable' -Subject $t.Name -Detail ("{0} folder(s): {1}" -f $walk.Unreadable.Count, (@($walk.Unreadable | Select-Object -First 10) -join '; '))
        }
        if ($walk.Truncated) { Add-CatacombFinding -Code 'ScanTruncated' -Subject $t.Name }

        [void]$results.Add([PSCustomObject]@{
            Name        = $t.Name
            Path        = $t.Path
            Description = $t.Description
            ShareAccess = $t.ShareAccess
            Root        = $root
            Folders     = $walk.Folders
            Unreadable  = $walk.Unreadable.Count
            BroadWrite  = ($shareBroadWrite -and $broadWrite.Count -gt 0)
            BroadRead   = ($shareBroadRead -and $broadRead.Count -gt 0)
        })
        if ($walk.Truncated) { break }
    }

    return [PSCustomObject]@{ Shares = @($results); AceRows = @($aceRows) }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Format-AceSummary {
    param([object[]]$Aces)
    return (@($Aces | Where-Object { $_.Type -eq 'Allow' } | Sort-Object Identity | ForEach-Object { "$($_.Identity): $($_.Level)" } | Select-Object -Unique) -join '; ')
}

function Build-CatacombReport {
    param([object]$Audit, [object]$Verdict, [string]$CsvName)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $order      = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>Nothing to report.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)</td>" +
            "<td>$(EscHtml $f.Subject)" + $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $sRows = [System.Text.StringBuilder]::new()
    if ($Audit.Shares.Count -eq 0) { [void]$sRows.Append("<tr><td colspan='5'>No shares reviewed.</td></tr>") }
    foreach ($s in $Audit.Shares) {
        $share = if ($null -eq $s.ShareAccess) { '(folder tree, no share)' } else { (@($s.ShareAccess | ForEach-Object { "$($_.Account): $($_.Right)$(if ($_.Type -ne 'Allow') { " ($($_.Type))" })" }) -join '; ') }
        $ntfs  = if ($s.Root) { Format-AceSummary -Aces $s.Root.Aces } else { '(unreadable)' }
        $flag  = if ($s.BroadWrite) { "<span class='tk-badge-err'>Broad write</span>" } elseif ($s.BroadRead) { "<span class='tk-badge-warn'>Broad read</span>" } else { "<span class='tk-badge-ok'>Scoped</span>" }
        [void]$sRows.Append(
            "<tr><td><strong>$(EscHtml $s.Name)</strong><br/><span class='tk-mono'>$(EscHtml $s.Path)</span></td>" +
            "<td>$flag</td><td>$(EscHtml $share)</td><td>$(EscHtml $ntfs)</td><td>$(@($s.Folders).Count)</td></tr>")
    }

    $dRows = [System.Text.StringBuilder]::new()
    $explicitFolders = @($Audit.Shares | ForEach-Object { $share = $_.Name; $_.Folders | Where-Object { $_.Level -gt 0 } | ForEach-Object { $_ | Add-Member -NotePropertyName Share -NotePropertyValue $share -Force -PassThru } })
    if ($explicitFolders.Count -eq 0) { [void]$dRows.Append("<tr><td colspan='4'>Every folder below the share roots inherits its permissions.</td></tr>") }
    foreach ($f in ($explicitFolders | Select-Object -First 500)) {
        $explicit = Format-AceSummary -Aces $f.Explicit
        [void]$dRows.Append(
            "<tr><td>$(EscHtml $f.Share)</td><td class='tk-mono'>$(EscHtml $f.Folder)</td>" +
            "<td>$(if ($f.Protected) { "<span class='tk-badge-info'>Disabled</span>" } else { 'On' })</td><td>$(EscHtml $explicit)</td></tr>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.A.T.A.C.O.M.B. Share Permissions Report' `
        -ScriptName 'C.A.T.A.C.O.M.B.' `
        -Subtitle   "${orgPrefix}File Share & NTFS Permissions -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Generated' = $reportDate
            'Scope'     = $(if ($Path) { $Path } else { "$(@($Audit.Shares).Count) share(s)" })
            'Depth'     = "$Depth level(s)"
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Shares', 'Folders With Their Own Permissions')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Audit.Shares).Count)</div><div class="tk-summary-lbl">Shares Reviewed</div></div>
    <div class="tk-summary-card $(if (@($Audit.Shares | Where-Object { $_.BroadWrite }).Count) { 'err' } else { 'ok' })"><div class="tk-summary-num">$(@($Audit.Shares | Where-Object { $_.BroadWrite }).Count)</div><div class="tk-summary-lbl">Broad Write</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($explicitFolders.Count)</div><div class="tk-summary-lbl">Folders With Own Permissions</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Audit.AceRows).Count)</div><div class="tk-summary-lbl">Entries in CSV</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Shares</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Share</th><th>Exposure</th><th>Share permissions</th><th>NTFS at the root (allow)</th><th>Folders recorded</th></tr></thead>
      <tbody>$($sRows.ToString())</tbody></table>
      <div class="tk-info-box"><span class="tk-info-label">Effective access</span> is the more restrictive of the share and NTFS permissions. Broad write is reported only where both let an everyone-type group change files.</div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Folders With Their Own Permissions</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Share</th><th>Folder</th><th>Inheritance</th><th>Explicit allow entries</th></tr></thead>
      <tbody>$($dRows.ToString())</tbody></table>
      <div class="tk-info-box"><span class="tk-info-label">Full detail</span> Every access-control entry scanned is in <span class="tk-mono">$(EscHtml $CsvName)</span> next to this report.</div></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'C.A.T.A.C.O.M.B. v5.1')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-CatacombBanner

if ($Path -and -not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Fail "-Path '$Path' is not a folder."
    exit 1
}

$audit = Invoke-CatacombAudit

Write-Section "FINDINGS"
if ($Findings.Count -eq 0) { Write-Ok "No broad, direct or orphaned access found." }
foreach ($f in $Findings) {
    $line = "$($f.Title) -- $($f.Subject)"
    switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
}

$verdict = Get-CatacombVerdict -FindingList $Findings.ToArray()
Write-Section "VERDICT"
switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
Add-TKNote -Text ("CATACOMB on {0}: verdict {1}; {2} share(s), {3} finding(s)." -f $env:COMPUTERNAME, $verdict.Verdict, @($audit.Shares).Count, $Findings.Count) -Category 'Info' -ScriptName 'catacomb'

$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir  = Resolve-LogDirectory -FallbackPath $ScriptPath
$csvName = "CATACOMB_$stamp.csv"
try {
    $audit.AceRows | Export-Csv -LiteralPath (Join-Path $logDir $csvName) -NoTypeInformation -Encoding UTF8
    Write-Ok "Access-control entries exported: $csvName"
} catch {
    Write-Warn "Could not write the CSV: $($_.Exception.Message)"
}

Write-Step "Generating HTML report..."
$outPath = Join-Path $logDir "CATACOMB_$stamp.html"
try {
    [System.IO.File]::WriteAllText($outPath, (Build-CatacombReport -Audit $audit -Verdict $verdict -CsvName $csvName), [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
    Write-TKError -ScriptName 'catacomb' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
