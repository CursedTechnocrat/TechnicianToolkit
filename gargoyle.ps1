# gargoyle.ps1 - forwarding stub: GARGOYLE is now VIGIL (vigil.ps1)
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
    Forwarding stub. GARGOYLE was renamed VIGIL; this script passes every argument to vigil.ps1.

.DESCRIPTION
    GARGOYLE was renamed VIGIL in 6.0 to fit The Observatory's theme, scrying & magical sight.
    This stub keeps pinned runbooks, custom RITUAL recipe files and quick-launch snippets working:
    it prints a warning, downloads vigil.ps1 next to itself if it is missing, and runs it with
    the same arguments. It will be removed in a future release -- update references to vigil.ps1.

.USAGE
    PS C:\> .\gargoyle.ps1 [arguments]          # Same as .\vigil.ps1 [arguments]

.NOTES
    Version : 6.0

#>

$NewScript = Join-Path $PSScriptRoot 'vigil.ps1'
if (-not (Test-Path $NewScript)) {
    $NewScriptUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/vigil.ps1'
    Write-Host "  [*] vigil.ps1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $NewScriptUrl -OutFile $NewScript -ErrorAction Stop
        # Re-save with a BOM so Windows PowerShell 5.1 reads the script as UTF-8.
        [IO.File]::WriteAllText($NewScript, [IO.File]::ReadAllText($NewScript, [Text.Encoding]::UTF8), [Text.UTF8Encoding]::new($true))
    } catch {
        Write-Host "  [!!] Could not download vigil.ps1: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Get it from $NewScriptUrl" -ForegroundColor Yellow
        exit 1
    }
}

Write-Warning "gargoyle.ps1 has been renamed to vigil.ps1 (VIGIL). This forwarding stub will be removed in a future release; update the reference."
& $NewScript @args
exit $LASTEXITCODE
