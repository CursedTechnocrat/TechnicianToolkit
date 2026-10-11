# tether.ps1 - forwarding stub: TETHER is now PHYLACTERY (phylactery.ps1)
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
    Forwarding stub. TETHER was renamed PHYLACTERY; this script passes every argument to phylactery.ps1.

.DESCRIPTION
    TETHER was renamed PHYLACTERY in 6.0 to fit The Necropolis's theme, the undead & ghosts.
    This stub keeps pinned runbooks, custom RITUAL recipe files and quick-launch snippets working:
    it prints a warning, downloads phylactery.ps1 next to itself if it is missing, and runs it with
    the same arguments. It will be removed in a future release -- update references to phylactery.ps1.

.USAGE
    PS C:\> .\tether.ps1 [arguments]          # Same as .\phylactery.ps1 [arguments]

.NOTES
    Version : 6.0

#>

$NewScript = Join-Path $PSScriptRoot 'phylactery.ps1'
if (-not (Test-Path $NewScript)) {
    $NewScriptUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/phylactery.ps1'
    if ($env:TK_DISABLE_DOWNLOAD -in @('1', 'true')) {
        Write-Host "  [!!] phylactery.ps1 not found, and TK_DISABLE_DOWNLOAD forbids fetching it." -ForegroundColor Red
        Write-Host "       Deploy phylactery.ps1 next to this script from an approved release." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "  [*] phylactery.ps1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $NewScriptUrl -OutFile $NewScript -ErrorAction Stop
        # Re-save with a BOM so Windows PowerShell 5.1 reads the script as UTF-8.
        [IO.File]::WriteAllText($NewScript, [IO.File]::ReadAllText($NewScript, [Text.Encoding]::UTF8), [Text.UTF8Encoding]::new($true))
    } catch {
        Write-Host "  [!!] Could not download phylactery.ps1: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Get it from $NewScriptUrl" -ForegroundColor Yellow
        exit 1
    }
}

Write-Warning "tether.ps1 has been renamed to phylactery.ps1 (PHYLACTERY). This forwarding stub will be removed in a future release; update the reference."
& $NewScript @args
exit $LASTEXITCODE
