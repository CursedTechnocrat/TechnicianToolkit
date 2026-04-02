<#
.SYNOPSIS
  Install software preferring Chocolatey with Winget fallback, then remove non-English parts of Microsoft 365.
  Prompts or accepts switches for whether DisplayLink is needed, whether this is a Dell machine, and whether to install Zoom Outlook plugin.

.PARAMETER ForceWinget
  Skip Chocolatey and use Winget only.

.PARAMETER NoInstallProviders
  Don't attempt to install Chocolatey or Winget if missing.

.PARAMETER DisplayLinkNeeded
  If present, treat DisplayLink as required (non-interactive). If omitted, the script will ask.

.PARAMETER IsDell
  If present, treat this as a Dell machine (non-interactive). If omitted, the script will ask.

.PARAMETER InstallZoomOutlookPlugin
  If present, attempt to install the Zoom Outlook plugin automatically when Zoom is installed. If omitted, the script will ask.
#>

param(
    [switch] $ForceWinget,
    [switch] $NoInstallProviders,
    [switch] $DisplayLinkNeeded,
    [switch] $IsDell,
    [switch] $InstallZoomOutlookPlugin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $current = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $current.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    Write-Host "Not running as Administrator. Relaunching elevated..."
    $argsList = @()
    if ($ForceWinget) { $argsList += "-ForceWinget" }
    if ($NoInstallProviders) { $argsList += "-NoInstallProviders" }
    if ($DisplayLinkNeeded) { $argsList += "-DisplayLinkNeeded" }
    if ($IsDell) { $argsList += "-IsDell" }
    if ($InstallZoomOutlookPlugin) { $argsList += "-InstallZoomOutlookPlugin" }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwsh) { $psi.FileName = $pwsh.Source } else { $psi.FileName = (Get-Command powershell).Source }
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $($argsList -join ' ')"
    $psi.Verb = "runas"
    try {
        [System.Diagnostics.Process]::Start($psi) | Out-Null
        exit
    } catch {
        Write-Error "Failed to relaunch elevated: $($_.Exception.Message)"
        exit 1
    }
}

function Is-ChocoInstalled {
    try { choco --version > $null 2>&1; return $true } catch { return $false }
}
function Install-Chocolatey {
    if ($NoInstallProviders) { Write-Host "Skipping Chocolatey installation due to -NoInstallProviders."; return $false }
    Write-Host "Installing Chocolatey..."
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-Host "Chocolatey installed."
        return $true
    } catch {
        Write-Warning "Chocolatey install failed: $($_.Exception.Message)"
        return $false
    }
}

function Is-WingetInstalled {
    try { winget --version > $null 2>&1; return $true } catch { return $false }
}
function Install-Winget {
    if ($NoInstallProviders) { Write-Host "Skipping Winget installation due to -NoInstallProviders."; return $false }
    Write-Host "Installing WinGet (Windows Package Manager)..."
    try {
        $progressPreference = 'SilentlyContinue'
        irm https://aka.ms/getwinget | iex
        Start-Sleep -Seconds 3
        return Is-WingetInstalled
    } catch {
        Write-Warning "WinGet install failed: $($_.Exception.Message)"
        return $false
    }
}

function Fix-WingetSources {
    if (-not (Is-WingetInstalled)) { return }
    try {
        Write-Host "Fixing Winget sources..."
        winget source remove msstore -h -e 2>$null
    } catch { }
    try {
        winget source add --name winget https://cdn.winget.microsoft.com/cache -h
        Write-Host "Sources fixed successfully."
    } catch {
        Write-Warning "Error fixing Winget sources: $($_.Exception.Message)"
    }
}

function Update-Winget {
    if (-not (Is-WingetInstalled)) { return }
    try {
        Write-Host "Checking for Winget updates..."
        winget upgrade --id Microsoft.WinGet -h -e 2>$null
    } catch {
        Write-Verbose "No Winget update or error: $($_.Exception.Message)"
    }
}

# === Package mapping with optional metadata ===
$PackageMap = @{
    "Microsoft Teams" = @{ choco = "microsoft-teams"; winget = "Microsoft.Teams" }
    "Microsoft Office" = @{ choco = $null;                  winget = "Microsoft.Office" }
    "7-Zip"            = @{ choco = "7zip.install";         winget = "7zip.7zip" }
    "Google Chrome"    = @{ choco = "googlechrome";         winget = "Google.Chrome" }
    "DisplayLink Driver" = @{
        choco = $null;
        winget = "DisplayLink.GraphicsDriver";
        Optional = @{ Param = "DisplayLinkNeeded"; Prompt = "Is DisplayLink needed on this device? (Y/N)" }
    }
    "Zoom"             = @{ choco = "zoom";                 winget = "Zoom.Zoom" }
    "Adobe Reader"     = @{ choco = "adobereader";          winget = "Adobe.Acrobat.Reader.64-bit" }
    "Dell Command Update" = @{
        choco = "dell-command-update";
        winget = "Dell.CommandUpdate";
        Optional = @{ Param = "IsDell"; Prompt = "Is this computer a Dell? (Y/N)" }
    }
}

# Resolve user choices for optional entries
$UserChoices = @{}
function Ask-YesNo {
    param([string]$Question, [bool]$DefaultYes = $true)
    while ($true) {
        $def = if ($DefaultYes) { "Y" } else { "N" }
        $answer = Read-Host "$Question [$def]"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $DefaultYes }
        switch ($answer.ToUpper()) {
            "Y" { return $true }
            "YES" { return $true }
            "N" { return $false }
            "NO" { return $false }
            default { Write-Host "Please answer Y or N." }
        }
    }
}

# Populate choices from switches or prompt if missing
foreach ($entry in $PackageMap.GetEnumerator()) {
    if ($entry.Value.ContainsKey('Optional')) {
        $opt = $entry.Value.Optional
        $paramName = $opt.Param
        if ($PSBoundParameters.ContainsKey($paramName)) {
            $UserChoices[$paramName] = [bool]$PSBoundParameters[$paramName]
        } else {
            $prompt = $opt.Prompt
            $choice = Ask-YesNo -Question $prompt -DefaultYes $false
            $UserChoices[$paramName] = $choice
        }
    }
}

# Provider availability
$chocoAvailable = -not $ForceWinget -and (Is-ChocoInstalled)
if (-not $chocoAvailable -and -not $ForceWinget) {
    Write-Host "Chocolatey not found."
    $chocoInstalledNow = Install-Chocolatey
    $chocoAvailable = $chocoInstalledNow -or (Is-ChocoInstalled)
}

$wingetAvailable = Is-WingetInstalled
if (-not $wingetAvailable) {
    Write-Host "Winget not found."
    $wingetInstalledNow = Install-Winget
    $wingetAvailable = $wingetInstalledNow -or (Is-WingetInstalled)
}

if ($wingetAvailable) {
    Fix-WingetSources
    Update-Winget
}

$installResults = @()

function Install-PackageWithFallback {
    param([string]$FriendlyName, [string]$ChocoId, [string]$WingetId)
    $status = [ordered]@{ Name = $FriendlyName; ChocoAttempted = $false; ChocoSuccess = $false; WingetAttempted = $false; WingetSuccess = $false; Error = $null }

    if (-not $ForceWinget -and $chocoAvailable -and $ChocoId) {
        $status.ChocoAttempted = $true
        Write-Host "Attempting Chocolatey install for $FriendlyName (id: $ChocoId)..."
        try {
            & choco install $ChocoId -y --no-progress
            $status.ChocoSuccess = $true
            Write-Host "Chocolatey: $FriendlyName installed (or already present)."
        } catch {
            $status.Error = "Chocolatey error: $($_.Exception.Message)"
            Write-Warning "Chocolatey install failed for $FriendlyName: $($_.Exception.Message)"
        }
    }

    if ($status.ChocoSuccess) { return $status }

    if ($wingetAvailable -and $WingetId) {
        $status.WingetAttempted = $true
        Write-Host "Attempting Winget install for $FriendlyName (id: $WingetId)..."
        try {
            winget install -e --id $WingetId --accept-source-agreements --accept-package-agreements --silent
            $status.WingetSuccess = $true
            Write-Host "Winget: $FriendlyName installed (or already present)."
        } catch {
            $status.Error = ($status.Error + " ; Winget error: $($_.Exception.Message)").Trim(' ', ';')
            Write-Warning "Winget install failed for $FriendlyName: $($_.Exception.Message)"
        }
    } elseif (-not $wingetAvailable) {
        $status.Error = ($status.Error + " ; Winget not available").Trim(' ', ';')
        Write-Warning "Winget is not available to install $FriendlyName."
    } else {
        $status.Error = ($status.Error + " ; No mapping for Winget").Trim(' ', ';')
        Write-Warning "No Winget mapping to attempt for $FriendlyName."
    }

    return $status
}

# ---- Zoom Outlook plugin helper ----
function Get-ZoomOutlookPluginInstalled {
    # Search uninstall registry entries for typical Zoom Outlook plugin display names
    $searchPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($p in $searchPaths) {
        try {
            $items = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -and ($_.DisplayName -match "Zoom" -and ($_.DisplayName -match "Outlook" -or $_.DisplayName -match "Plugin" -or $_.DisplayName -match "Add-in"))
            }
            if ($items) { return $true }
        } catch {}
    }
    return $false
}

function Install-ZoomOutlookPlugin {
    Write-Host "Preparing to install Zoom Outlook plugin..."

    if (Get-ZoomOutlookPluginInstalled) {
        Write-Host "Zoom Outlook plugin is already installed. Skipping."
        return $true
    }

    $installed = $false
    # Try Chocolatey candidates first if available
    if ($chocoAvailable) {
        $chocoCandidates = @("zoomoutlookplugin","zoom-outlook-plugin","zoomplugin","zoomoutlook")
        foreach ($c in $chocoCandidates) {
            try {
                Write-Host "Attempting Chocolatey install candidate: $c"
                & choco install $c -y --no-progress
                Start-Sleep -Seconds 2
                if (Get-ZoomOutlookPluginInstalled) { $installed = $true; break }
            } catch {
                Write-Verbose "Chocolatey candidate $c failed: $($_.Exception.Message)"
            }
        }
    }

    # If not installed, try Winget candidates
    if (-not $installed -and $wingetAvailable) {
        $wingetCandidates = @("Zoom.OutlookPlugin","Zoom.ZoomOutlook","Zoom.ZoomOutlookPlugin","Zoom.ZoomAddinOutlook")
        foreach ($id in $wingetCandidates) {
            try {
                Write-Host "Attempting Winget install candidate: $id"
                winget install -e --id $id --accept-source-agreements --accept-package-agreements --silent
                Start-Sleep -Seconds 2
                if (Get-ZoomOutlookPluginInstalled) { $installed = $true; break }
            } catch {
                Write-Verbose "Winget candidate $id failed: $($_.Exception.Message)"
            }
        }
    }

    if ($installed) {
        Write-Host "Zoom Outlook plugin installed successfully (detected)."
        return $true
    } else {
        Write-Warning "Could not install Zoom Outlook plugin automatically. Possible reasons: no package exists in Chocolatey/Winget with expected names, or the plugin requires interactive installer. Manual installation instructions:"
        Write-Host " - Visit https://support.zoom.us and search 'Zoom Plugin for Microsoft Outlook' or 'Zoom Outlook add-in' to download the correct installer for your environment."
        return $false
    }
}

# ---- M365 language removal helpers ----
function Reconfigure-OfficeToEnUS {
    Write-Host "Attempting to reconfigure Office to en-US (best-effort)..."
    $odtSetup = Join-Path $PSScriptRoot "ODT"
    $setupExe = Join-Path $odtSetup "setup.exe"
    if (Test-Path $setupExe) {
        $configXml = Join-Path $odtSetup "en-us-only-config.xml"
        if (-not (Test-Path $configXml)) {
            Write-Warning "ODT setup.exe found but en-us-only-config.xml missing in $odtSetup. Please supply a config.xml and run setup.exe /configure config.xml"
            return $false
        }
        try {
            Start-Process -FilePath $setupExe -ArgumentList "/configure `"$configXml`"" -Wait -NoNewWindow
            Write-Host "ODT configure attempted."
            return $true
        } catch {
            Write-Warning "Failed to run ODT: $($_.Exception.Message)"
            return $false
        }
    } else {
        Write-Warning "Office Deployment Tool (ODT) not found in script folder. To reliably remove languages from Click-to-Run Office you should use ODT with a config.xml that specifies only en-US. See: https://learn.microsoft.com/deployoffice/overview-office-deployment-tool"
        return $false
    }
}

function Remove-M365 {
    param([string]$DisplayName)

    Write-Host "Remove-M365: searching for installed items matching '$DisplayName'..."

    $searchPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $matches = @()
    foreach ($p in $searchPaths) {
        try {
            $items = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and ($_.DisplayName -like "*$DisplayName*") }
            if ($items) { $matches += $items }
        } catch {}
    }

    if ($matches.Count -gt 0) {
        foreach ($m in $matches) {
            $display = $m.DisplayName
            $uninstall = $m.UninstallString
            Write-Host "Found: $display"

            if (-not $uninstall) {
                Write-Warning "No UninstallString for $display; skipping registry uninstall attempt."
                continue
            }

            $cmd = $uninstall.Trim()
            if ($cmd -match 'msiexec' -or $cmd -match '/X' ) {
                if ($cmd -notmatch '/qn' -and $cmd -notmatch '/quiet') { $cmd += " /qn /norestart" }
                Write-Host "Running MSI uninstall for $display: $cmd"
                try {
                    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -NoNewWindow -Wait
                    Write-Host "Uninstall invoked for $display"
                } catch {
                    Write-Warning "Failed to run uninstall for $display: $($_.Exception.Message)"
                }
                continue
            }

            $exe = $null; $args = $null
            if ($cmd -match '^"([^"]+)"\s*(.*)$') {
                $exe = $matches[1]; $args = $matches[2]
            } else {
                $parts = $cmd.Split(" ",2)
                $exe = $parts[0]
                $args = if ($parts.Count -gt 1) { $parts[1] } else { "" }
            }

            if ($exe -match 'ClickToRun' -or $exe -match 'OfficeC2RClient' -or $display -match 'Microsoft 365') {
                Write-Warning "Detected Click-to-Run / Microsoft 365 entry ($display). Registry uninstall may not remove language packs. Attempting ODT reconfiguration as fallback."
                $ok = Reconfigure-OfficeToEnUS
                if (-not $ok) {
                    Write-Warning "ODT reconfiguration did not run. For a reliable language pack removal please download the Office Deployment Tool and run a configuration that keeps only en-US. Example config and docs: https://learn.microsoft.com/deployoffice/overview-office-deployment-tool"
                }
                continue
            }

            if ($args -notmatch '/quiet|/q|/S|/s|/qn') {
                $args = ($args + " /quiet").Trim()
            }

            Write-Host "Executing uninstall: $exe $args"
            try {
                Start-Process -FilePath $exe -ArgumentList $args -Wait -NoNewWindow
                Write-Host "Uninstall called for $display"
            } catch {
                Write-Warning "Failed to run uninstall for $display: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Warning "No installed entries found that match '$DisplayName'. Attempting ODT reconfiguration fallback..."
        $ok = Reconfigure-OfficeToEnUS
        if (-not $ok) {
            Write-Warning "ODT reconfiguration did not run. Manual intervention may be required to remove language components."
        }
    }
}

# ---- Main install loop ----
foreach ($entry in $PackageMap.GetEnumerator()) {
    $name = $entry.Key
    $meta = $entry.Value

    # Skip optional package if user chose no
    if ($meta.ContainsKey('Optional')) {
        $optParam = $meta.Optional.Param
        if ($UserChoices.ContainsKey($optParam) -and -not $UserChoices[$optParam]) {
            Write-Host "Skipping optional package '$name' based on user choice."
            continue
        }
    }

    $chocoId = $meta.choco
    $wingetId = $meta.winget

    $result = Install-PackageWithFallback -FriendlyName $name -ChocoId $chocoId -WingetId $wingetId
    $installResults += $result

    # If we just installed Zoom, optionally attempt to install the Zoom Outlook plugin
    if ($name -eq "Zoom" -and ($result.ChocoSuccess -or $result.WingetSuccess)) {
        $doPlugin = $false
        if ($PSBoundParameters.ContainsKey("InstallZoomOutlookPlugin")) {
            $doPlugin = [bool]$InstallZoomOutlookPlugin
        } else {
            $doPlugin = Ask-YesNo -Question "Zoom was installed. Install Zoom Outlook plugin? (Y/N)" -DefaultYes $true
        }
        if ($doPlugin) {
            Install-ZoomOutlookPlugin | Out-Null
        } else {
            Write-Host "Skipping Zoom Outlook plugin installation by user choice."
        }
    }
}

# After installing Office products, remove non-English language components you requested.
Remove-M365 "Microsoft 365 - fr-fr"
Remove-M365 "Microsoft 365 Apps for business - fr-fr"
Remove-M365 "Microsoft 365 - es-es"
Remove-M365 "Microsoft 365 Apps for business - es-es"
Remove-M365 "Aplicaciones de Microsoft 365 para negocios - es-es"
Remove-M365 "Microsoft 365 - pt-br"
Remove-M365 "Microsoft 365 Apps for business - pt-br"
Remove-M365 "Microsoft OneNote - fr-fr"
Remove-M365 "Microsoft OneNote - es-es"
Remove-M365 "Microsoft OneNote - pt-br"

# Results summary
Write-Host ""
Write-Host "=== Installation summary ==="
foreach ($r in $installResults) {
    $line = "{0} => ChocoAttempted:{1} ChocoSuccess:{2} WingetAttempted:{3} WingetSuccess:{4}" -f $r.Name, $r.ChocoAttempted, $r.ChocoSuccess, $r.WingetAttempted, $r.WingetSuccess
    Write-Host $line
    if ($r.Error) { Write-Host "  Error: $($r.Error)" -ForegroundColor Yellow }
}

Write-Host "Software install, optional Zoom Outlook plugin attempt, and M365 language removal steps completed (see warnings above if any)."
