# ================================
# Package Manager Setup
# ================================

function Ensure-Winget {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "Winget detected. Updating sources..."
        winget source update
    }
    else {
        Write-Host "Winget not found. Installing..."
        $progressPreference = 'SilentlyContinue'
        irm https://aka.ms/getwinget | iex
    }
}

function Ensure-Chocolatey {
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Host "Chocolatey detected. Updating..."
        choco upgrade chocolatey -y
    }
    else {
        Write-Host "Chocolatey not found. Installing..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    }
}

Ensure-Winget
Ensure-Chocolatey

# ================================
# Software Lists
# ================================

$RequiredSoftware = @(
    "Microsoft.Teams",
    "Microsoft.Office",
    "7zip.7zip",
    "Google.Chrome",
    "Adobe.Acrobat.Reader.64-bit",
    "Zoom.Zoom"
)

$OptionalSoftware = @(
    "Zoom.ZoomOutlookPlugin",
    "DisplayLink.GraphicsDriver",
    "Dell.CommandUpdate"
)

# ================================
# Install Function
# ================================

function Install-Software {
    param (
        [string]$PackageId
    )

    Write-Host "`nInstalling $PackageId..."

    try {
        winget install -e --id $PackageId `
            --accept-source-agreements `
            --accept-package-agreements `
            -h
    }
    catch {
        Write-Host "Winget failed. Trying Chocolatey..."

        switch ($PackageId) {
            "Microsoft.Teams" { choco install microsoft-teams -y }
            "Microsoft.Office" { choco install microsoft-office-deployment -y }
            "7zip.7zip" { choco install 7zip -y }
            "Google.Chrome" { choco install googlechrome -y }
            "Adobe.Acrobat.Reader.64-bit" { choco install adobereader -y }
            "Zoom.Zoom" { choco install zoom -y }
            "Zoom.ZoomOutlookPlugin" { choco install zoom-outlook -y }
            "DisplayLink.GraphicsDriver" { choco install displaylink -y }
            "Dell.CommandUpdate" { choco install dellcommandupdate -y }
            default { Write-Host "No Chocolatey fallback defined for $PackageId" }
        }
    }
}

# ================================
# Install Required Software
# ================================

Write-Host "`n=== Installing Required Software ==="
foreach ($app in $RequiredSoftware) {
    Install-Software -PackageId $app
}

# ================================
# Optional Software Prompt
# ================================

Write-Host "`n=== Optional Software ==="
foreach ($optional in $OptionalSoftware) {
    $choice = Read-Host "Install $optional? (Y/N)"
    if ($choice -match "^[Yy]") {
        Install-Software -PackageId $optional
    }
}

# ================================
# Completion
# ================================

Write-Host "`n✅ Software installation process completed."
