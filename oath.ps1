# oath.ps1 - forwarding stub: OATH is now LODESTAR (lodestar.ps1)
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
    Forwarding stub. OATH was renamed LODESTAR; this script passes every argument to lodestar.ps1.

.DESCRIPTION
    OATH was renamed LODESTAR in 6.0 to fit The Crossroads' theme, arcane paths & lights.
    This stub keeps pinned runbooks, custom RITUAL recipe files and quick-launch snippets working:
    it prints a warning, downloads lodestar.ps1 next to itself if it is missing, and runs it with
    the same arguments. It will be removed in a future release -- update references to lodestar.ps1.

.USAGE
    PS C:\> .\oath.ps1 [arguments]          # Same as .\lodestar.ps1 [arguments]

.NOTES
    Version : 6.0

#>

$NewScript = Join-Path $PSScriptRoot 'lodestar.ps1'
if (-not (Test-Path $NewScript)) {
    $NewScriptUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/lodestar.ps1'
    Write-Host "  [*] lodestar.ps1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $NewScriptUrl -OutFile $NewScript -ErrorAction Stop
        # Re-save with a BOM so Windows PowerShell 5.1 reads the script as UTF-8.
        [IO.File]::WriteAllText($NewScript, [IO.File]::ReadAllText($NewScript, [Text.Encoding]::UTF8), [Text.UTF8Encoding]::new($true))
    } catch {
        Write-Host "  [!!] Could not download lodestar.ps1: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Get it from $NewScriptUrl" -ForegroundColor Yellow
        exit 1
    }
}

Write-Warning "oath.ps1 has been renamed to lodestar.ps1 (LODESTAR). This forwarding stub will be removed in a future release; update the reference."
& $NewScript @args
exit $LASTEXITCODE
