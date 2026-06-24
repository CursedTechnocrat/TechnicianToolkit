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
