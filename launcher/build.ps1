# build.ps1 - Build the self-contained, portable TechnicianToolkit launcher (.exe).
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 CursedTechnocrat and the Technician Toolkit contributors
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

#requires -Version 5.1
<#
.SYNOPSIS
    Build the self-contained, portable TechnicianToolkit launcher (.exe).

.DESCRIPTION
    Publishes launcher\TechnicianToolkit.Launcher.csproj as a single-file,
    self-contained executable. Every toolkit script is embedded inside the .exe,
    so the resulting file is the whole suite in one portable artifact: drop it on
    a USB stick and run it on any Windows machine, fully offline, with no update
    checks.

    Requires the .NET SDK 8.0+ (https://dotnet.microsoft.com/download). The .exe
    bundles the .NET runtime, so target machines do NOT need .NET installed — they
    only need Windows PowerShell, which ships with Windows.

.USAGE
    PS> .\launcher\build.ps1
    PS> .\launcher\build.ps1 -Runtime win-x64 -Output .\dist
    PS> .\launcher\build.ps1 -Runtime win-arm64

.NOTES
    Version : 1.0
#>
[CmdletBinding()]
param(
    [string]$Runtime       = 'win-x64',
    [string]$Configuration = 'Release',
    [string]$Output        = (Join-Path $PSScriptRoot 'dist')
)

$ErrorActionPreference = 'Stop'

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
    Write-Host "[!!] The .NET SDK (dotnet) was not found on PATH." -ForegroundColor Red
    Write-Host "     Install .NET SDK 8.0+ from https://dotnet.microsoft.com/download and retry." -ForegroundColor Yellow
    exit 1
}

$project = Join-Path $PSScriptRoot 'TechnicianToolkit.Launcher.csproj'

Write-Host ""
Write-Host "  Building TechnicianToolkit launcher" -ForegroundColor Cyan
Write-Host "    runtime : $Runtime"
Write-Host "    config  : $Configuration"
Write-Host "    output  : $Output"
Write-Host ""

dotnet publish $project `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -o $Output

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[!!] Build failed (dotnet exit code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}

$exe = Get-ChildItem -Path $Output -Filter 'TechnicianToolkit*.exe' -File -ErrorAction SilentlyContinue |
    Select-Object -First 1

Write-Host ""
Write-Host "  Build complete." -ForegroundColor Green
if ($exe) {
    $sizeMb = [math]::Round($exe.Length / 1MB, 1)
    Write-Host "    $($exe.FullName)  ($sizeMb MB)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Copy that single file to a USB stick and run it on any Windows box." -ForegroundColor Gray
}
Write-Host ""
