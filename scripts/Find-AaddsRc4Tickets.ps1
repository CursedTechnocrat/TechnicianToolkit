<#
.SYNOPSIS
    Lists Kerberos service tickets issued with RC4 (etype 0x17) by a Microsoft
    Entra Domain Services managed domain, by querying the Log Analytics
    workspace that receives the AADDS Security Audit stream.

.DESCRIPTION
    Wraps the AADDomainServicesAccountLogon table KQL query as a PowerShell
    call. Returns one row per (ServiceName, AccountName, IpAddress) with how
    often that combination asked the KDC for an RC4 ticket. Use this to triage
    AADDS123 "Kerberos RC4 encryption is enabled" alerts -- the rows tell you
    which SPNs and clients are still AES-incompatible.

    Prerequisite: AADDS Security audits must be enabled and streaming to the
    target Log Analytics workspace. Enable via:
        Microsoft Entra Domain Services -> [domain] -> Security audits

.PARAMETER WorkspaceId
    Log Analytics workspace ID (GUID). Portal -> workspace -> Overview ->
    "Workspace ID".

.PARAMETER WorkspaceName
    Alternative to -WorkspaceId. Resolved against -ResourceGroup.

.PARAMETER ResourceGroup
    Resource group containing -WorkspaceName.

.PARAMETER Days
    Lookback window in days. Default 7.

.PARAMETER Etype
    Kerberos encryption type to filter on. Default '0x17' (RC4-HMAC).
    '0x18' = RC4-HMAC-EXP, '0x1'/'0x3' = legacy DES.

.PARAMETER OutCsv
    Optional CSV output path.

.EXAMPLE
    .\Find-AaddsRc4Tickets.ps1 -WorkspaceId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    .\Find-AaddsRc4Tickets.ps1 -WorkspaceName aadds-logs -ResourceGroup rg-identity -Days 14 -OutCsv .\rc4.csv
#>
[CmdletBinding(DefaultParameterSetName = 'ById')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [string]$WorkspaceId,

    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [string]$WorkspaceName,

    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [string]$ResourceGroup,

    [int]$Days = 7,

    [string]$Etype = '0x17',

    [string]$OutCsv
)

# --- prerequisites ---------------------------------------------------------
foreach ($mod in 'Az.Accounts', 'Az.OperationalInsights') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "  [*] Installing $mod..." -ForegroundColor Magenta
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module $mod -ErrorAction Stop
}

# --- auth ------------------------------------------------------------------
if (-not (Get-AzContext)) {
    Write-Host "  [*] Connecting to Azure..." -ForegroundColor Magenta
    Connect-AzAccount -ErrorAction Stop | Out-Null
}

# --- resolve workspace -----------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'ByName') {
    $ws = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroup `
                                             -Name $WorkspaceName -ErrorAction Stop
    $WorkspaceId = $ws.CustomerId.Guid
}

# --- query -----------------------------------------------------------------
$kql = @"
AADDomainServicesAccountLogon
| where TimeGenerated > ago(${Days}d)
| where OperationName == "Kerberos Service Ticket Operations" or EventID == 4769
| where TicketEncryptionType == "$Etype"
| summarize Tickets = count(), LastSeen = max(TimeGenerated)
        by ServiceName, AccountName, IpAddress
| order by Tickets desc
"@

Write-Host "  [*] Querying workspace $WorkspaceId for last $Days day(s) (etype $Etype)..." -ForegroundColor Magenta
$result = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $kql -ErrorAction Stop

if (-not $result.Results -or $result.Results.Count -eq 0) {
    Write-Host "  [+] No $Etype tickets in the last $Days day(s)." -ForegroundColor Green
    return
}

$rows = $result.Results |
    Select-Object ServiceName, AccountName, IpAddress, Tickets, LastSeen

$rows | Format-Table -AutoSize

if ($OutCsv) {
    $rows | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  [+] CSV: $OutCsv" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Next: AES-only the offenders found above:" -ForegroundColor Cyan
Write-Host "    Set-ADUser     <Sam> -Replace @{ 'msDS-SupportedEncryptionTypes' = 0x18 }" -ForegroundColor Gray
Write-Host "    Set-ADComputer <Sam> -Replace @{ 'msDS-SupportedEncryptionTypes' = 0x18 }" -ForegroundColor Gray
Write-Host "  And in the portal: AADDS -> Security settings -> Disable Kerberos RC4 = Enabled" -ForegroundColor Gray
