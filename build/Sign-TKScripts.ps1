# Sign-TKScripts.ps1 - Signs the release copies of the toolkit scripts and writes the release assets.
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
    Signs the release copies of every toolkit script and writes the release assets.

.DESCRIPTION
    Run on the signing machine during a release (RELEASING.md, "Sign the
    scripts"), with SimplySign Desktop connected. It works on the
    TechnicianToolkit-scripts-unsigned artifact the release workflow uploads,
    or on a checkout of the release tag.

    1. Copies every top-level .ps1, LICENSE and config.json into -OutDir.
    2. Signs each script with signtool, by thumbprint, with an RFC 3161
       timestamp so the signature outlives the short-lived certificate.
    3. Verifies each one with the same check GRIMOIRE and RITUAL apply to a
       download (Test-TKScriptSignature): valid, and from the toolkit's
       publisher. It stops if any script fails, or carries no timestamp.
    4. Writes SHA256SUMS-scripts.txt and TechnicianToolkit-scripts-<version>.zip.

    Attach every signed .ps1 to the release individually -- GRIMOIRE and RITUAL
    fetch a missing tool from releases/latest/download/<tool>.ps1 -- plus the
    zip and the checksum file.

    The repository's own copies are never signed. A signature breaks the moment
    a file changes, and build/Sync-TKShared.ps1 changes them routinely; only
    what ships is signed.

.USAGE
    PS C:\> .\build\Sign-TKScripts.ps1 -Thumbprint F14F545CB084EDCB18630CA0B520ABC3C46B2516
    PS C:\> .\scripts-unsigned\build\Sign-TKScripts.ps1 -Source .\scripts-unsigned -Thumbprint <thumbprint>

.NOTES
    Windows only: signtool and Get-AuthenticodeSignature are Windows tools.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9A-Fa-f]{40}$')][string]$Thumbprint,
    [string]$Source       = (Split-Path -Parent $PSScriptRoot),
    [string]$OutDir       = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist') 'scripts'),
    [string]$SignTool     = '',
    [string]$TimestampUrl = 'http://time.certum.pl'
)

$ErrorActionPreference = 'Stop'

# Test-TKScriptSignature and the publisher name come from the shared helpers, so
# this checks exactly what a download is checked against.
. (Join-Path (Join-Path $Source 'shared') 'TechnicianToolkit.Shared.ps1')

if (-not $SignTool) {
    $SignTool = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe' -ErrorAction SilentlyContinue |
        Sort-Object FullName | Select-Object -Last 1 -ExpandProperty FullName
}
if (-not $SignTool -or -not (Test-Path $SignTool)) {
    throw "signtool.exe not found. Install the Windows SDK, or pass -SignTool <path>."
}

$version = ([regex]::Match((Get-Content (Join-Path $Source 'grimoire.ps1') -Raw), '(?m)^\s*Version\s*:\s*([0-9][0-9.]*)')).Groups[1].Value
if (-not $version) { throw "Could not read the suite version from grimoire.ps1." }

$scripts = @(Get-ChildItem -Path $Source -Filter '*.ps1' -File | Sort-Object Name)
if ($scripts.Count -eq 0) { throw "No scripts found in $Source." }

if (Test-Path $OutDir) { Remove-Item -LiteralPath $OutDir -Recurse -Force }
$null = New-Item -ItemType Directory -Path $OutDir -Force
$scripts | Copy-Item -Destination $OutDir
Copy-Item -LiteralPath (Join-Path $Source 'LICENSE'), (Join-Path $Source 'config.json') -Destination $OutDir

$signed = @(Get-ChildItem -Path $OutDir -Filter '*.ps1' -File | Sort-Object Name)
Write-Host "  [*] Signing $($signed.Count) scripts with certificate $Thumbprint..." -ForegroundColor Magenta
& $SignTool sign /sha1 $Thumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 @($signed.FullName)
if ($LASTEXITCODE -ne 0) { throw "signtool failed (exit $LASTEXITCODE)." }

$failed = @(foreach ($f in $signed) {
    $sig = Get-AuthenticodeSignature -FilePath $f.FullName
    if (-not (Test-TKScriptSignature -Path $f.FullName)) { "$($f.Name): $($sig.Status) $($sig.StatusMessage)" }
    elseif (-not $sig.TimeStamperCertificate)           { "$($f.Name): signed but not timestamped" }
})
if ($failed.Count -gt 0) {
    $failed | ForEach-Object { Write-Host "  [!!] $_" -ForegroundColor Red }
    throw "$($failed.Count) script(s) failed verification; do not publish."
}
Write-Host "  [+] Every script is signed, timestamped and verifies as $script:TKPublisher." -ForegroundColor Green

$sums = Join-Path $OutDir 'SHA256SUMS-scripts.txt'
Get-ChildItem -Path $OutDir -File | Where-Object Name -ne 'SHA256SUMS-scripts.txt' | Sort-Object Name |
    ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLower(), $_.Name } |
    Set-Content -Path $sums -Encoding ascii

$zip = Join-Path (Split-Path -Parent $OutDir) "TechnicianToolkit-scripts-$version.zip"
Compress-Archive -Path (Join-Path $OutDir '*') -DestinationPath $zip -Force

Write-Host ""
Write-Host "  Attach to the release:" -ForegroundColor Cyan
Write-Host "    every .ps1 in $OutDir (individually -- GRIMOIRE and RITUAL fetch them by name)" -ForegroundColor Gray
Write-Host "    $sums" -ForegroundColor Gray
Write-Host "    $zip" -ForegroundColor Gray
