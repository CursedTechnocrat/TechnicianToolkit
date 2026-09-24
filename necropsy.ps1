# necropsy.ps1 - N.E.C.R.O.P.S.Y. — Names Each Crash, Reboot & Outage — Post-mortem Summary Yield
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
    N.E.C.R.O.P.S.Y. — Names Each Crash, Reboot & Outage — Post-mortem Summary Yield
    Crash & Unexpected-Reboot Analysis Tool for PowerShell 5.1+

.DESCRIPTION
    Answers "why does this machine keep crashing or rebooting?" by correlating
    the evidence Windows leaves behind after an unplanned stop: bugcheck
    (blue-screen) events, Kernel-Power 41 unexpected-shutdown records, WHEA
    hardware-error reports, display-driver resets (TDR), the dump files on disk
    and their headers, and the Reliability Monitor stability index.

    Each crash is classified -- a bugcheck, a sudden power loss or hard hang, or
    a forced power-off from the power button -- and every bugcheck code is
    mapped to its name and the area it usually implicates (driver, memory,
    storage, graphics, hardware, power). Driver installs and Windows updates in
    the same window are laid on the timeline, so "it started after the update"
    is visible at a glance. The crash-dump configuration is checked too: a
    machine with dumps disabled or no page file leaves no evidence next time.

    Read-only -- nothing on the machine is changed. N.E.C.R.O.P.S.Y. reads the
    bugcheck code and parameters, not the stack: naming the faulting driver
    needs a debugger (WinDbg, !analyze -v) against the dump files it lists.

.USAGE
    PS C:\> .\necropsy.ps1                    # Interactive menu
    PS C:\> .\necropsy.ps1 -Unattended        # Analyse the last 30 days + HTML report
    PS C:\> .\necropsy.ps1 -Unattended -Days 90   # Look back 90 days instead

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [ValidateRange(1, 365)]
    [int]$Days = 30,
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
$C = $ColorSchema

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

$CrashControlKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl'
$MemoryMgmtKey   = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'

# Events closer together than this are treated as one incident. A crash leaves
# a Kernel-Power 41, an EventLog 6008 and (for a bugcheck) a WER 1001, all
# written during the next boot within a minute or two of each other.
$IncidentWindowMinutes = 15

# Corrected hardware errors are routine in small numbers (a single PCIe
# correction is not news); this many in the window is worth a warning.
$CorrectedWheaThreshold = 5

# CrashDumpEnabled -> what Windows writes on a bugcheck.
$DumpTypeLabels = @{
    0 = 'None'
    1 = 'Complete memory dump'
    2 = 'Kernel memory dump'
    3 = 'Small memory dump (256 KB)'
    7 = 'Automatic memory dump'
}

# Bugcheck code -> name, the area it usually implicates, and a one-line hint.
# Keys are the canonical form produced by ConvertTo-BugCheckKey: '0x', then
# upper-case hex with no leading zeros. Codes outside this table still report
# by number; the table only adds the name and the triage hint.
$BugCheckCatalog = @{
    '0xA'      = @{ Name = 'IRQL_NOT_LESS_OR_EQUAL';             Area = 'Driver';   Hint = 'A kernel driver touched memory it should not have at a raised IRQL.' }
    '0x19'     = @{ Name = 'BAD_POOL_HEADER';                    Area = 'Driver';   Hint = 'Kernel pool corruption, almost always from a driver writing past its allocation.' }
    '0x1A'     = @{ Name = 'MEMORY_MANAGEMENT';                  Area = 'Memory';   Hint = 'Memory manager found corrupted state -- faulty RAM first, then a misbehaving driver.' }
    '0x1E'     = @{ Name = 'KMODE_EXCEPTION_NOT_HANDLED';        Area = 'Driver';   Hint = 'A kernel-mode driver raised an exception nothing handled.' }
    '0x24'     = @{ Name = 'NTFS_FILE_SYSTEM';                   Area = 'Storage';  Hint = 'NTFS hit an inconsistency -- check the disk and run chkdsk.' }
    '0x3B'     = @{ Name = 'SYSTEM_SERVICE_EXCEPTION';           Area = 'Driver';   Hint = 'An exception during a system call, typically a driver or security product.' }
    '0x50'     = @{ Name = 'PAGE_FAULT_IN_NONPAGED_AREA';        Area = 'Memory';   Hint = 'Invalid memory referenced -- faulty RAM, or a driver using freed memory.' }
    '0x77'     = @{ Name = 'KERNEL_STACK_INPAGE_ERROR';          Area = 'Storage';  Hint = 'Kernel stack could not be read back from the page file -- disk or controller.' }
    '0x7A'     = @{ Name = 'KERNEL_DATA_INPAGE_ERROR';           Area = 'Storage';  Hint = 'Kernel data could not be paged in from disk -- failing drive, cable or controller.' }
    '0x7B'     = @{ Name = 'INACCESSIBLE_BOOT_DEVICE';           Area = 'Storage';  Hint = 'Windows lost the boot volume -- storage driver, controller mode (AHCI/RAID) or disk.' }
    '0x7E'     = @{ Name = 'SYSTEM_THREAD_EXCEPTION_NOT_HANDLED'; Area = 'Driver';  Hint = 'A system thread raised an unhandled exception, usually inside a driver.' }
    '0x7F'     = @{ Name = 'UNEXPECTED_KERNEL_MODE_TRAP';        Area = 'Hardware'; Hint = 'CPU trap the kernel did not expect -- hardware fault, overclock, or stack overflow.' }
    '0x9C'     = @{ Name = 'MACHINE_CHECK_EXCEPTION';            Area = 'Hardware'; Hint = 'The CPU reported a fatal machine check -- processor, memory controller or overheating.' }
    '0x9F'     = @{ Name = 'DRIVER_POWER_STATE_FAILURE';         Area = 'Power';    Hint = 'A driver failed a sleep / wake / shutdown power transition in time.' }
    '0xA0'     = @{ Name = 'INTERNAL_POWER_ERROR';               Area = 'Power';    Hint = 'Power policy manager hit a fatal error, often during hibernate.' }
    '0xBE'     = @{ Name = 'ATTEMPTED_WRITE_TO_READONLY_MEMORY'; Area = 'Driver';   Hint = 'A driver wrote to read-only memory.' }
    '0xC2'     = @{ Name = 'BAD_POOL_CALLER';                    Area = 'Driver';   Hint = 'A driver made an invalid pool request.' }
    '0xC5'     = @{ Name = 'DRIVER_CORRUPTED_EXPOOL';            Area = 'Driver';   Hint = 'A driver corrupted the kernel pool.' }
    '0xD1'     = @{ Name = 'DRIVER_IRQL_NOT_LESS_OR_EQUAL';      Area = 'Driver';   Hint = 'A driver accessed pageable memory at a raised IRQL -- network, storage and AV drivers are common culprits.' }
    '0xE2'     = @{ Name = 'MANUALLY_INITIATED_CRASH';           Area = 'Manual';   Hint = 'Deliberately triggered (keyboard crash or NMI). Not a fault.' }
    '0xEF'     = @{ Name = 'CRITICAL_PROCESS_DIED';              Area = 'System';   Hint = 'A process Windows cannot run without exited -- system file corruption or storage.' }
    '0xF4'     = @{ Name = 'CRITICAL_OBJECT_TERMINATION';        Area = 'Storage';  Hint = 'A critical process or thread terminated, most often because the system disk stopped responding.' }
    '0xFC'     = @{ Name = 'ATTEMPTED_EXECUTE_OF_NOEXECUTE_MEMORY'; Area = 'Driver'; Hint = 'Code ran from a non-executable page -- driver bug or memory corruption.' }
    '0x101'    = @{ Name = 'CLOCK_WATCHDOG_TIMEOUT';             Area = 'Hardware'; Hint = 'A processor stopped responding to clock interrupts -- CPU, firmware or overclock.' }
    '0x109'    = @{ Name = 'CRITICAL_STRUCTURE_CORRUPTION';      Area = 'Driver';   Hint = 'Kernel code or data was modified -- a driver patching the kernel, or bad RAM.' }
    '0x10D'    = @{ Name = 'WDF_VIOLATION';                      Area = 'Driver';   Hint = 'A driver built on the Windows Driver Framework broke its rules.' }
    '0x10E'    = @{ Name = 'VIDEO_MEMORY_MANAGEMENT_INTERNAL';   Area = 'Graphics'; Hint = 'The graphics memory manager hit a fatal error.' }
    '0x113'    = @{ Name = 'VIDEO_DXGKRNL_FATAL_ERROR';          Area = 'Graphics'; Hint = 'The DirectX graphics kernel failed.' }
    '0x116'    = @{ Name = 'VIDEO_TDR_FAILURE';                  Area = 'Graphics'; Hint = 'The GPU hung and could not be reset -- graphics driver, GPU or its power delivery.' }
    '0x117'    = @{ Name = 'VIDEO_TDR_TIMEOUT_DETECTED';         Area = 'Graphics'; Hint = 'The display driver failed to respond in time.' }
    '0x119'    = @{ Name = 'VIDEO_SCHEDULER_INTERNAL_ERROR';     Area = 'Graphics'; Hint = 'The GPU scheduler detected a fatal violation.' }
    '0x124'    = @{ Name = 'WHEA_UNCORRECTABLE_ERROR';           Area = 'Hardware'; Hint = 'The hardware reported an uncorrectable error -- CPU, RAM, PCIe device, overheating or unstable overclock.' }
    '0x12B'    = @{ Name = 'FAULTY_HARDWARE_CORRUPTED_PAGE';     Area = 'Memory';   Hint = 'A single-bit error was found in a memory page -- faulty RAM.' }
    '0x133'    = @{ Name = 'DPC_WATCHDOG_VIOLATION';             Area = 'Storage';  Hint = 'A DPC ran too long -- storage controller drivers and SSD firmware are the usual cause.' }
    '0x139'    = @{ Name = 'KERNEL_SECURITY_CHECK_FAILURE';      Area = 'Driver';   Hint = 'The kernel detected corruption of a critical data structure.' }
    '0x13A'    = @{ Name = 'KERNEL_MODE_HEAP_CORRUPTION';        Area = 'Driver';   Hint = 'A driver corrupted the kernel heap.' }
    '0x154'    = @{ Name = 'UNEXPECTED_STORE_EXCEPTION';         Area = 'Storage';  Hint = 'The memory-compression store hit an error -- often a failing disk.' }
    '0x1CA'    = @{ Name = 'SYNTHETIC_WATCHDOG_TIMEOUT';         Area = 'Hardware'; Hint = 'A system-wide watchdog expired -- the machine stopped making progress.' }
    '0xC000021A' = @{ Name = 'STATUS_SYSTEM_PROCESS_TERMINATED'; Area = 'System';   Hint = 'Winlogon or CSRSS died -- system file corruption, a bad update, or a broken security product.' }
}

# Area -> what to do next, with the toolkit tool that goes deeper where one exists.
$AreaGuidance = [ordered]@{
    'Driver'   = 'Identify the driver with WinDbg (!analyze -v) against the dump. Update or roll back recently changed drivers -- F.O.R.G.E. lists problem devices and pending driver updates.'
    'Memory'   = 'Test the RAM: run mdsched.exe (Windows Memory Diagnostic) or MemTest86 overnight. Reseat or swap modules; remove any XMP / overclock profile.'
    'Storage'  = 'Check the disk: A.U.G.U.R. reads SMART health and wear. Update SSD firmware and the storage controller driver; run chkdsk /scan.'
    'Graphics' = 'Clean-install the current graphics driver, or roll back one version if the crashes started after an update. Check GPU temperature and power connectors.'
    'Hardware' = 'Suspect the hardware itself: check temperatures, remove overclocks, update BIOS / UEFI (A.N.V.I.L. reports firmware state), and review the WHEA errors below for the failing component.'
    'Power'    = 'A driver is failing a sleep or wake transition. Update chipset, network and storage drivers; test with hibernate / fast startup disabled. On laptops, P.Y.R.E. checks the battery.'
    'System'   = 'Repair system files: DISM /Online /Cleanup-Image /RestoreHealth, then sfc /scannow. Uninstall the most recent update if the crashes began right after it.'
    'Power loss' = 'The machine lost power or froze too hard to blue-screen. Check the power supply, battery, power cabling and any UPS first, then temperatures and firmware.'
    'Hang'     = 'The machine stopped responding and was forced off. Hangs share causes with crashes: check the disk with A.U.G.U.R., recent driver changes, and temperatures.'
    'Manual'   = 'Someone triggered this crash deliberately. No repair needed -- confirm who and why.'
    'Unknown'  = 'The code is not in the catalog. Search the code at learn.microsoft.com (Bug Check Code Reference) and analyse the dump in WinDbg.'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition N.E.C.R.O.P.S.Y. can report, keyed by a stable code. Kind
# separates evidence of a crash ('Crash'), evidence of instability short of a
# crash ('Instability'), and a configuration gap that would hide the next crash
# ('Readiness'). The verdict reads Kind, so a machine with dumps disabled but
# no crashes is reported as Stable with a readiness warning -- not Unstable.
# ─────────────────────────────────────────────────────────────────────────────

$NecropsyFindings = @{
    'BugCheck' = @{
        Severity = 'Error'
        Kind     = 'Crash'
        Title    = 'Blue-screen crash (bugcheck) recorded'
        Summary  = 'Windows stopped with a bugcheck. The code names the kind of failure and the area it usually implicates.'
        Remedy   = 'Follow the guidance for the area below. To name the faulting driver, open the dump in WinDbg and run !analyze -v.'
    }
    'UnexpectedPowerLoss' = @{
        Severity = 'Error'
        Kind     = 'Crash'
        Title    = 'Sudden power loss or hard hang'
        Summary  = 'Kernel-Power 41 with no bugcheck and no power-button press: the machine lost power or froze so hard it could not blue-screen.'
        Remedy   = 'Check the power supply, battery and power cabling first, then temperatures. A hard freeze with no dump points at hardware or firmware rather than a driver.'
    }
    'ForcedPowerOff' = @{
        Severity = 'Warning'
        Kind     = 'Instability'
        Title    = 'Machine forced off with the power button'
        Summary  = 'Kernel-Power 41 recorded a power-button press: someone held the button, usually because the machine had stopped responding.'
        Remedy   = 'Ask the user what the machine was doing. Repeated hangs share causes with crashes -- drivers, storage and heat.'
    }
    'FatalHardwareError' = @{
        Severity = 'Error'
        Kind     = 'Crash'
        Title    = 'Fatal hardware error (WHEA)'
        Summary  = 'The Windows Hardware Error Architecture logged an uncorrectable error. The component field names the failing part.'
        Remedy   = 'Treat as a hardware fault: check temperatures, remove overclocks, update firmware, and test or replace the named component.'
    }
    'CorrectedHardwareErrors' = @{
        Severity = 'Warning'
        Kind     = 'Instability'
        Title    = 'Repeated corrected hardware errors (WHEA)'
        Summary  = 'The hardware is correcting errors often enough to be worth attention. Corrected errors are a common precursor to uncorrectable ones.'
        Remedy   = 'Note the component. PCIe errors: update chipset drivers and firmware, reseat the device. Memory or cache errors: test the RAM and CPU.'
    }
    'DisplayDriverReset' = @{
        Severity = 'Warning'
        Kind     = 'Instability'
        Title    = 'Display driver stopped responding and recovered (TDR)'
        Summary  = 'The GPU hung and Windows reset the display driver. Frequent resets often escalate to a VIDEO_TDR_FAILURE bugcheck.'
        Remedy   = 'Clean-install the graphics driver; check GPU temperature and power. Roll back if the resets started after a driver update.'
    }
    'DumpsDisabled' = @{
        Severity = 'Warning'
        Kind     = 'Readiness'
        Title    = 'Crash dumps are disabled'
        Summary  = 'CrashDumpEnabled is 0, so the next bugcheck will leave no dump to analyse.'
        Remedy   = 'System Properties > Advanced > Startup and Recovery: set "Write debugging information" to Automatic memory dump.'
    }
    'NoPageFile' = @{
        Severity = 'Warning'
        Kind     = 'Readiness'
        Title    = 'No page file for crash dumps'
        Summary  = 'Windows writes a crash dump through the page file on the system drive. With no page file and no dedicated dump file, the dump is lost.'
        Remedy   = 'Re-enable a system-managed page file on the system drive, or configure a DedicatedDumpFile.'
    }
    'DumpsMissing' = @{
        Severity = 'Info'
        Kind     = 'Readiness'
        Title    = 'Bugchecks recorded but no dump files on disk'
        Summary  = 'The event log shows bugchecks, but the dump files are gone -- usually removed by Disk Cleanup or a cleanup tool.'
        Remedy   = 'Nothing to analyse from past crashes. Leave dumps in place after the next one and copy them off before any cleanup.'
    }
    'LowStability' = @{
        Severity = 'Warning'
        Kind     = 'Instability'
        Title    = 'Low Reliability Monitor stability index'
        Summary  = 'Windows rates this machine below 5 out of 10, from the application, driver and system failures it has recorded.'
        Remedy   = 'Open Reliability Monitor (perfmon /rel) to see the failures behind the score.'
    }
    'EventLogUnreadable' = @{
        Severity = 'Warning'
        Kind     = 'Readiness'
        Title    = 'System event log could not be read'
        Summary  = 'Part of the crash history could not be queried, so the picture below may be incomplete.'
        Remedy   = 'Run elevated, and check that the Windows Event Log service is running.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-NecropsyFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Detail = ''
    )
    $meta = $NecropsyFindings[$Code]
    if (-not $meta) {
        $meta = @{ Severity = 'Warning'; Kind = 'Instability'; Title = $Code; Summary = ''; Remedy = '' }
    }
    [void]$Findings.Add([PSCustomObject]@{
        Code     = $Code
        Severity = $meta.Severity
        Kind     = $meta.Kind
        Title    = $meta.Title
        Summary  = $meta.Summary
        Remedy   = $meta.Remedy
        Detail   = $Detail
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

function Show-NecropsyBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███╗   ██╗███████╗ ██████╗██████╗  ██████╗ ██████╗ ███████╗██╗   ██╗
  ████╗  ██║██╔════╝██╔════╝██╔══██╗██╔═══██╗██╔══██╗██╔════╝╚██╗ ██╔╝
  ██╔██╗ ██║█████╗  ██║     ██████╔╝██║   ██║██████╔╝███████╗ ╚████╔╝
  ██║╚██╗██║██╔══╝  ██║     ██╔══██╗██║   ██║██╔═══╝ ╚════██║  ╚██╔╝
  ██║ ╚████║███████╗╚██████╗██║  ██║╚██████╔╝██║     ███████║   ██║
  ╚═╝  ╚═══╝╚══════╝ ╚═════╝╚═╝  ╚═╝ ╚═════╝ ╚═╝     ╚══════╝   ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    N.E.C.R.O.P.S.Y. — Names Each Crash, Reboot & Outage — Post-mortem Summary Yield" -ForegroundColor Cyan
    Write-Host "    Crash & Unexpected-Reboot Analysis Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
#
# No I/O below this line until COLLECTORS -- the Pester suite extracts these by
# AST lookup and calls them directly with synthetic input.
# ─────────────────────────────────────────────────────────────────────────────

function ConvertTo-BugCheckKey {
    # 0x0000009f, 159 and 0x9F must all land on the same catalog key.
    param([Parameter(Mandatory)][uint64]$Code)
    return ('0x{0:X}' -f $Code)
}

function ConvertFrom-BugCheckText {
    # WER 1001 carries the bugcheck as "0x0000009f (0x0000000000000003, ...)".
    # The first hex token is the code; the four in parentheses are parameters.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $tokens = @([regex]::Matches($Text, '0x[0-9A-Fa-f]+') | ForEach-Object { $_.Value })
    if ($tokens.Count -eq 0) { return $null }
    $code = [Convert]::ToUInt64($tokens[0].Substring(2), 16)
    return [PSCustomObject]@{
        Code       = $code
        Key        = (ConvertTo-BugCheckKey -Code $code)
        Parameters = @($tokens | Select-Object -Skip 1 -First 4)
    }
}

function Get-KernelPowerCause {
    # Kernel-Power 41 is written on the boot after any unclean stop. Its fields
    # say which kind: a bugcheck code means a blue screen; a power-button
    # timestamp means the user held the button; neither means the power went
    # or the machine froze too hard to crash.
    param(
        [uint64]$BugcheckCode = 0,
        [uint64]$PowerButtonTimestamp = 0
    )
    if ($BugcheckCode -ne 0)         { return 'Bugcheck' }
    if ($PowerButtonTimestamp -ne 0) { return 'PowerButton' }
    return 'PowerLoss'
}

function Get-DumpHeaderInfo {
    # A kernel dump (complete, kernel or small) opens with a DUMP_HEADER.
    # 64-bit: 'PAGE' 'DU64', BugCheckCode at 0x38, parameters at 0x40.
    # 32-bit: 'PAGE' 'DUMP', BugCheckCode at 0x28, parameters at 0x2C.
    param([byte[]]$Bytes)

    $unknown = [PSCustomObject]@{ Format = 'Unknown'; Code = $null; Key = ''; Parameters = @() }
    if (-not $Bytes -or $Bytes.Length -lt 0x60) { return $unknown }

    $sig   = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, 4)
    $valid = [System.Text.Encoding]::ASCII.GetString($Bytes, 4, 4)
    if ($sig -ne 'PAGE') { return $unknown }

    if ($valid -eq 'DU64') {
        $code   = [uint64][BitConverter]::ToUInt32($Bytes, 0x38)
        $params = foreach ($off in 0x40, 0x48, 0x50, 0x58) { '0x{0:X}' -f [BitConverter]::ToUInt64($Bytes, $off) }
        return [PSCustomObject]@{ Format = '64-bit'; Code = $code; Key = (ConvertTo-BugCheckKey -Code $code); Parameters = @($params) }
    }
    if ($valid -eq 'DUMP') {
        $code   = [uint64][BitConverter]::ToUInt32($Bytes, 0x28)
        $params = foreach ($off in 0x2C, 0x30, 0x34, 0x38) { '0x{0:X}' -f [BitConverter]::ToUInt32($Bytes, $off) }
        return [PSCustomObject]@{ Format = '32-bit'; Code = $code; Key = (ConvertTo-BugCheckKey -Code $code); Parameters = @($params) }
    }
    return $unknown
}

function Get-BugCheckInfo {
    param([string]$Key)
    $entry = $BugCheckCatalog[$Key]
    if ($entry) {
        return [PSCustomObject]@{ Key = $Key; Name = $entry.Name; Area = $entry.Area; Hint = $entry.Hint }
    }
    return [PSCustomObject]@{ Key = $Key; Name = 'Unrecognised bugcheck'; Area = 'Unknown'; Hint = 'Not in the N.E.C.R.O.P.S.Y. catalog.' }
}

function Get-NecropsyVerdict {
    # Reads the Kind of each finding, not its severity: a readiness gap such as
    # disabled dumps must not make a machine that never crashed look unstable.
    param([object[]]$FindingList)
    $kinds = @($FindingList | ForEach-Object { $_.Kind })
    if ($kinds -contains 'Crash')       { return [PSCustomObject]@{ Verdict = 'Crashing'; Class = 'err'  } }
    if ($kinds -contains 'Instability') { return [PSCustomObject]@{ Verdict = 'Unstable'; Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Stable'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-EventsSafe {
    # Get-WinEvent throws when a filter matches nothing. That is an empty result,
    # not a failure, and only a real failure should raise EventLogUnreadable.
    param([hashtable]$Filter)
    try {
        return [PSCustomObject]@{ Events = @(Get-WinEvent -FilterHashtable $Filter -ErrorAction Stop); Failed = $false; Error = '' }
    } catch {
        if ($_.FullyQualifiedErrorId -match 'NoMatchingEventsFound' -or $_.Exception.Message -match 'No events were found') {
            return [PSCustomObject]@{ Events = @(); Failed = $false; Error = '' }
        }
        return [PSCustomObject]@{ Events = @(); Failed = $true; Error = $_.Exception.Message }
    }
}

function Get-EventDataValue {
    param($Record, [string]$Name)
    try {
        $xml  = [xml]$Record.ToXml()
        $node = @($xml.Event.EventData.Data | Where-Object { $_.Name -eq $Name }) | Select-Object -First 1
        if ($node) { return $node.'#text' }
    } catch {
        return $null
    }
    return $null
}

function Get-FirstLine {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return (($Text -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
}

function Get-NecropsyDeviceContext {
    $cs = $null; $os = $null
    try { $cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop } catch { $cs = $null }
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop } catch { $os = $null }

    $lastBoot = if ($os) { $os.LastBootUpTime } else { $null }
    $uptime   = if ($lastBoot) { (Get-Date) - $lastBoot } else { $null }

    return [PSCustomObject]@{
        Computer = $env:COMPUTERNAME
        Model    = if ($cs) { "$($cs.Manufacturer) $($cs.Model)".Trim() } else { 'Unknown' }
        OS       = if ($os) { "$($os.Caption) (build $($os.BuildNumber))" } else { 'Unknown' }
        LastBoot = if ($lastBoot) { $lastBoot.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }
        Uptime   = if ($uptime) { '{0}d {1}h {2}m' -f $uptime.Days, $uptime.Hours, $uptime.Minutes } else { 'Unknown' }
    }
}

function Get-NecropsyDumpConfig {
    $cc = Get-ItemProperty -Path $CrashControlKey -ErrorAction SilentlyContinue
    $mm = Get-ItemProperty -Path $MemoryMgmtKey   -ErrorAction SilentlyContinue

    $type = if ($null -ne $cc.CrashDumpEnabled) { [int]$cc.CrashDumpEnabled } else { 7 }
    $label = if ($DumpTypeLabels.ContainsKey($type)) { $DumpTypeLabels[$type] } else { "Code $type" }
    # CrashDumpEnabled 1 with FilterPages 1 is the "Active memory dump" option.
    if ($type -eq 1 -and $cc.FilterPages -eq 1) { $label = 'Active memory dump' }

    $dumpFile    = if ($cc.DumpFile)     { [Environment]::ExpandEnvironmentVariables($cc.DumpFile) }     else { Join-Path $env:SystemRoot 'MEMORY.DMP' }
    $minidumpDir = if ($cc.MinidumpDir)  { [Environment]::ExpandEnvironmentVariables($cc.MinidumpDir) }  else { Join-Path $env:SystemRoot 'Minidump' }
    $dedicated   = if ($cc.DedicatedDumpFile) { [Environment]::ExpandEnvironmentVariables($cc.DedicatedDumpFile) } else { '' }

    # PagingFiles is a multi-string; an empty one (or the lone "") means no page file.
    $pageFiles = @($mm.PagingFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $hasPageFile = $pageFiles.Count -gt 0

    if ($type -eq 0) { Add-NecropsyFinding -Code 'DumpsDisabled' }
    if ($type -ne 0 -and -not $hasPageFile -and -not $dedicated) { Add-NecropsyFinding -Code 'NoPageFile' }

    return [PSCustomObject]@{
        DumpType     = $type
        DumpLabel    = $label
        AutoReboot   = if ($null -ne $cc.AutoReboot) { [bool]$cc.AutoReboot } else { $true }
        DumpFile     = $dumpFile
        MinidumpDir  = $minidumpDir
        Dedicated    = $dedicated
        PageFiles    = if ($hasPageFile) { $pageFiles -join '; ' } else { '(none)' }
        HasPageFile  = $hasPageFile
    }
}

function Get-NecropsyDumpFiles {
    param([object]$Config)

    $files = @()
    if (Test-Path $Config.MinidumpDir) {
        $files += @(Get-ChildItem -Path $Config.MinidumpDir -Filter '*.dmp' -File -ErrorAction SilentlyContinue)
    }
    if (Test-Path $Config.DumpFile) {
        $files += @(Get-Item -Path $Config.DumpFile -ErrorAction SilentlyContinue)
    }

    $rows = foreach ($f in ($files | Sort-Object LastWriteTime -Descending)) {
        $header = [PSCustomObject]@{ Format = 'Unreadable'; Code = $null; Key = ''; Parameters = @() }
        $fs = $null
        try {
            # Read the header only -- a complete memory dump can be tens of GB.
            $fs  = [System.IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
            $buf = New-Object byte[] 0x60
            $read = $fs.Read($buf, 0, $buf.Length)
            if ($read -eq $buf.Length) { $header = Get-DumpHeaderInfo -Bytes $buf }
        } catch {
            $header = [PSCustomObject]@{ Format = 'Unreadable'; Code = $null; Key = ''; Parameters = @() }
        } finally {
            if ($fs) { $fs.Dispose() }
        }

        [PSCustomObject]@{
            Name      = $f.Name
            Path      = $f.FullName
            SizeBytes = $f.Length
            Written   = $f.LastWriteTime
            Format    = $header.Format
            Key       = $header.Key
            Info      = if ($header.Key) { Get-BugCheckInfo -Key $header.Key } else { $null }
        }
    }
    return @($rows)
}

function Get-NecropsyEvents {
    param([datetime]$Since)

    $failures = [System.Collections.Generic.List[string]]::new()
    function _q {
        param([hashtable]$Filter, [string]$Label)
        $r = Get-EventsSafe -Filter $Filter
        if ($r.Failed) { $failures.Add("${Label}: $($r.Error)") }
        return $r.Events
    }

    $bugchecks = _q -Label 'BugCheck (WER 1001)'      -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WER-SystemErrorReporting'; Id = 1001; StartTime = $Since }
    $power41   = _q -Label 'Kernel-Power 41'          -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 41; StartTime = $Since }
    $unexp6008 = _q -Label 'EventLog 6008'            -Filter @{ LogName = 'System'; ProviderName = 'EventLog'; Id = 6008; StartTime = $Since }
    $planned   = _q -Label 'User32 1074'              -Filter @{ LogName = 'System'; ProviderName = 'User32'; Id = 1074; StartTime = $Since }
    $tdr       = _q -Label 'Display 4101'             -Filter @{ LogName = 'System'; ProviderName = 'Display'; Id = 4101; StartTime = $Since }
    $whea      = _q -Label 'WHEA-Logger'              -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WHEA-Logger'; StartTime = $Since }
    $boots     = _q -Label 'Kernel-General 12'        -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-General'; Id = 12; StartTime = $Since }
    $drivers   = _q -Label 'UserPnp 20001'            -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-UserPnp'; Id = 20001; StartTime = $Since }
    $updates   = _q -Label 'WindowsUpdateClient 19'   -Filter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WindowsUpdateClient'; Id = 19; StartTime = $Since }

    if ($failures.Count -gt 0) { Add-NecropsyFinding -Code 'EventLogUnreadable' -Detail ($failures -join ' | ') }

    return [PSCustomObject]@{
        BugChecks = @($bugchecks)
        Power41   = @($power41)
        Unexpected= @($unexp6008)
        Planned   = @($planned)
        Tdr       = @($tdr)
        Whea      = @($whea)
        Boots     = @($boots)
        Drivers   = @($drivers)
        Updates   = @($updates)
    }
}

function Get-NecropsyStability {
    try {
        $m = Get-CimInstance -ClassName Win32_ReliabilityStabilityMetrics -ErrorAction Stop |
            Sort-Object TimeGenerated -Descending | Select-Object -First 1
        if ($m) { return [math]::Round([double]$m.SystemStabilityIndex, 1) }
    } catch {
        return $null
    }
    return $null
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Test-NearIncident {
    param([datetime]$Time, [object[]]$Incidents)
    foreach ($i in $Incidents) {
        if ([math]::Abs(($i.Time - $Time).TotalMinutes) -le $IncidentWindowMinutes) { return $true }
    }
    return $false
}

function Get-NecropsyIncidents {
    # One row per unplanned stop. A bugcheck is reported once even though it
    # leaves a WER 1001, a Kernel-Power 41 and a dump file behind.
    param([object]$Events, [object[]]$Dumps, [datetime]$Since)

    $incidents = [System.Collections.Generic.List[object]]::new()

    foreach ($e in $Events.BugChecks) {
        $text   = if ($e.Properties.Count -gt 0) { "$($e.Properties[0].Value)" } else { $e.Message }
        $parsed = ConvertFrom-BugCheckText -Text $text
        if (-not $parsed) { $parsed = ConvertFrom-BugCheckText -Text $e.Message }
        $dump   = if ($e.Properties.Count -gt 1) { "$($e.Properties[1].Value)" } else { '' }
        $info   = if ($parsed) { Get-BugCheckInfo -Key $parsed.Key } else { Get-BugCheckInfo -Key '' }
        [void]$incidents.Add([PSCustomObject]@{
            Time       = $e.TimeCreated
            Type       = 'Bugcheck'
            Key        = if ($parsed) { $parsed.Key } else { '?' }
            Name       = $info.Name
            Area       = $info.Area
            Parameters = if ($parsed) { $parsed.Parameters -join ', ' } else { '' }
            Source     = 'WER 1001'
            Dump       = $dump
        })
    }

    foreach ($e in $Events.Power41) {
        $bc    = Get-EventDataValue -Record $e -Name 'BugcheckCode'
        $pb    = Get-EventDataValue -Record $e -Name 'PowerButtonTimestamp'
        $sleep = Get-EventDataValue -Record $e -Name 'SleepInProgress'
        [uint64]$bcNum = 0; [uint64]$pbNum = 0
        [void][uint64]::TryParse("$bc", [ref]$bcNum)
        [void][uint64]::TryParse("$pb", [ref]$pbNum)
        $cause = Get-KernelPowerCause -BugcheckCode $bcNum -PowerButtonTimestamp $pbNum

        if ($cause -eq 'Bugcheck') {
            # Normally already counted from WER 1001. Kept only when 1001 is
            # missing -- the dump failed to write, so WER had nothing to report.
            if (Test-NearIncident -Time $e.TimeCreated -Incidents @($incidents | Where-Object { $_.Type -eq 'Bugcheck' })) { continue }
            $key  = ConvertTo-BugCheckKey -Code $bcNum
            $info = Get-BugCheckInfo -Key $key
            [void]$incidents.Add([PSCustomObject]@{
                Time = $e.TimeCreated; Type = 'Bugcheck'; Key = $key; Name = $info.Name; Area = $info.Area
                Parameters = ''; Source = 'Kernel-Power 41 (no dump written)'; Dump = ''
            })
            continue
        }

        $sleepNote = if ("$sleep" -and "$sleep" -ne '0') { ' during a sleep transition' } else { '' }
        [void]$incidents.Add([PSCustomObject]@{
            Time       = $e.TimeCreated
            Type       = if ($cause -eq 'PowerButton') { 'Forced power-off' } else { 'Power loss / hard hang' }
            Key        = ''
            Name       = if ($cause -eq 'PowerButton') { "Power button held$sleepNote" } else { "No bugcheck, no button press$sleepNote" }
            Area       = if ($cause -eq 'PowerButton') { 'Hang' } else { 'Power loss' }
            Parameters = ''
            Source     = 'Kernel-Power 41'
            Dump       = ''
        })
    }

    # A dump whose event has rolled out of the log is still evidence.
    foreach ($d in $Dumps) {
        if ($d.Written -lt $Since -or -not $d.Key) { continue }
        if (Test-NearIncident -Time $d.Written -Incidents @($incidents | Where-Object { $_.Type -eq 'Bugcheck' })) { continue }
        [void]$incidents.Add([PSCustomObject]@{
            Time = $d.Written; Type = 'Bugcheck'; Key = $d.Key; Name = $d.Info.Name; Area = $d.Info.Area
            Parameters = ''; Source = 'Dump file header'; Dump = $d.Path
        })
    }

    return @($incidents | Sort-Object Time -Descending)
}

function Get-NecropsyWhea {
    param([object[]]$Events)
    # Levels 1-2 are uncorrectable, 3 is corrected; informational WHEA records are not errors.
    $rows = foreach ($e in @($Events | Where-Object { $_.Level -in 1, 2, 3 })) {
        $component = ''
        if ($e.Message -match 'Component:\s*(.+)') { $component = $Matches[1].Trim() }
        [PSCustomObject]@{
            Time      = $e.TimeCreated
            Id        = $e.Id
            Fatal     = ($e.Level -in 1, 2)
            Level     = "$($e.LevelDisplayName)"
            Component = $component
            Message   = Get-FirstLine -Text $e.Message
        }
    }
    return @($rows)
}

function Invoke-NecropsyAnalysis {
    param([int]$LookbackDays)

    $since = (Get-Date).AddDays(-$LookbackDays)

    Write-Section "DEVICE"
    $device = Get-NecropsyDeviceContext
    Write-Info "$($device.Computer)  |  $($device.Model)"
    Write-Info "$($device.OS)"
    Write-Info "Last boot $($device.LastBoot)  (up $($device.Uptime))"

    Write-Section "CRASH DUMP CONFIGURATION"
    $config = Get-NecropsyDumpConfig
    Write-Info "Dump type   : $($config.DumpLabel)"
    Write-Info "Page file   : $($config.PageFiles)"
    Write-Info "Minidumps   : $($config.MinidumpDir)"

    Write-Section "EVENT LOG (last $LookbackDays day(s))"
    Write-Step "Reading System log..."
    $events = Get-NecropsyEvents -Since $since
    Write-Info ("Bugchecks {0}  |  Kernel-Power 41 {1}  |  WHEA {2}  |  TDR {3}  |  Boots {4}" -f `
        $events.BugChecks.Count, $events.Power41.Count, $events.Whea.Count, $events.Tdr.Count, $events.Boots.Count)

    Write-Section "DUMP FILES"
    $dumps = Get-NecropsyDumpFiles -Config $config
    if ($dumps.Count -eq 0) {
        Write-Info "No dump files on disk."
    } else {
        foreach ($d in $dumps) {
            $label = if ($d.Info) { "$($d.Key) $($d.Info.Name)" } else { $d.Format }
            Write-Info ("{0}  {1}  {2}" -f $d.Written.ToString('yyyy-MM-dd HH:mm'), $d.Name, $label)
        }
    }

    $incidents = Get-NecropsyIncidents -Events $events -Dumps $dumps -Since $since
    $whea      = Get-NecropsyWhea -Events $events.Whea
    $stability = Get-NecropsyStability

    # ── Findings from the evidence ──
    $bugchecks = @($incidents | Where-Object { $_.Type -eq 'Bugcheck' })
    foreach ($g in ($bugchecks | Group-Object Key | Sort-Object Count -Descending)) {
        $first = $g.Group | Sort-Object Time -Descending | Select-Object -First 1
        Add-NecropsyFinding -Code 'BugCheck' -Detail ("{0} {1} -- {2} time(s), last {3}. Area: {4}." -f `
            $g.Name, $first.Name, $g.Count, $first.Time.ToString('yyyy-MM-dd HH:mm'), $first.Area)
    }

    $powerLoss = @($incidents | Where-Object { $_.Type -eq 'Power loss / hard hang' })
    if ($powerLoss.Count -gt 0) {
        Add-NecropsyFinding -Code 'UnexpectedPowerLoss' -Detail ("{0} time(s), last {1}." -f $powerLoss.Count, $powerLoss[0].Time.ToString('yyyy-MM-dd HH:mm'))
    }
    $forced = @($incidents | Where-Object { $_.Type -eq 'Forced power-off' })
    if ($forced.Count -gt 0) {
        Add-NecropsyFinding -Code 'ForcedPowerOff' -Detail ("{0} time(s), last {1}." -f $forced.Count, $forced[0].Time.ToString('yyyy-MM-dd HH:mm'))
    }

    $fatalWhea     = @($whea | Where-Object { $_.Fatal })
    $correctedWhea = @($whea | Where-Object { -not $_.Fatal })
    if ($fatalWhea.Count -gt 0) {
        $components = @($fatalWhea | Where-Object { $_.Component } | ForEach-Object { $_.Component } | Select-Object -Unique)
        Add-NecropsyFinding -Code 'FatalHardwareError' -Detail ("{0} event(s){1}" -f $fatalWhea.Count, $(if ($components) { ". Component: $($components -join ', ')" } else { '' }))
    }
    if ($correctedWhea.Count -ge $CorrectedWheaThreshold) {
        $components = @($correctedWhea | Where-Object { $_.Component } | ForEach-Object { $_.Component } | Select-Object -Unique)
        Add-NecropsyFinding -Code 'CorrectedHardwareErrors' -Detail ("{0} event(s){1}" -f $correctedWhea.Count, $(if ($components) { ". Component: $($components -join ', ')" } else { '' }))
    }

    if ($events.Tdr.Count -gt 0) {
        Add-NecropsyFinding -Code 'DisplayDriverReset' -Detail ("{0} reset(s), last {1}." -f $events.Tdr.Count, ($events.Tdr | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated.ToString('yyyy-MM-dd HH:mm'))
    }

    if ($bugchecks.Count -gt 0 -and $dumps.Count -eq 0) {
        Add-NecropsyFinding -Code 'DumpsMissing' -Detail "$($bugchecks.Count) bugcheck(s) in the window, 0 dump files in $($config.MinidumpDir) or $($config.DumpFile)."
    }

    if ($null -ne $stability -and $stability -lt 5) {
        Add-NecropsyFinding -Code 'LowStability' -Detail "Current index: $stability / 10"
    }

    # ── Timeline: incidents plus the changes that might explain them ──
    $timeline = [System.Collections.Generic.List[object]]::new()
    foreach ($i in $incidents) {
        $what = if ($i.Key) { "$($i.Key) $($i.Name)" } else { $i.Name }
        [void]$timeline.Add([PSCustomObject]@{ Time = $i.Time; Kind = $i.Type; Class = 'err'; Detail = $what })
    }
    foreach ($e in $events.Tdr)     { [void]$timeline.Add([PSCustomObject]@{ Time = $e.TimeCreated; Kind = 'Display driver reset'; Class = 'warn'; Detail = (Get-FirstLine -Text $e.Message) }) }
    foreach ($w in $fatalWhea)      { [void]$timeline.Add([PSCustomObject]@{ Time = $w.Time; Kind = 'Fatal hardware error'; Class = 'err'; Detail = $(if ($w.Component) { $w.Component } else { $w.Message }) }) }
    foreach ($e in $events.Drivers) { [void]$timeline.Add([PSCustomObject]@{ Time = $e.TimeCreated; Kind = 'Driver installed'; Class = 'blue'; Detail = (Get-FirstLine -Text $e.Message) }) }
    foreach ($e in $events.Updates) { [void]$timeline.Add([PSCustomObject]@{ Time = $e.TimeCreated; Kind = 'Update installed'; Class = 'blue'; Detail = (Get-FirstLine -Text $e.Message) }) }
    foreach ($e in $events.Planned) { [void]$timeline.Add([PSCustomObject]@{ Time = $e.TimeCreated; Kind = 'Planned restart'; Class = 'info'; Detail = (Get-FirstLine -Text $e.Message) }) }

    # Areas seen across bugchecks and power events, most frequent first.
    $areas = @($incidents | Where-Object { $_.Area } | Group-Object Area | Sort-Object Count -Descending | ForEach-Object { $_.Name })
    if ($fatalWhea.Count -gt 0 -and $areas -notcontains 'Hardware') { $areas += 'Hardware' }

    return [PSCustomObject]@{
        LookbackDays = $LookbackDays
        Since        = $since
        Device       = $device
        Config       = $config
        Events       = $events
        Dumps        = $dumps
        Incidents    = $incidents
        Whea         = $whea
        Stability    = $stability
        Timeline     = @($timeline | Sort-Object Time -Descending)
        Areas        = $areas
    }
}

function Show-NecropsyFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) {
        Write-Ok "No crashes, hangs or hardware errors in the window, and crash dumps are configured."
        return
    }
    foreach ($f in ($Findings | Sort-Object @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } })) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) {
            'Error'   { Write-Fail $line }
            'Warning' { Write-Warn $line }
            default   { Write-Info $line }
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-NecropsyReport {
    param([object]$Analysis, [object]$Verdict)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $bugcheckCount = @($Analysis.Incidents | Where-Object { $_.Type -eq 'Bugcheck' }).Count
    $powerCount    = @($Analysis.Incidents | Where-Object { $_.Type -ne 'Bugcheck' }).Count
    $fatalWhea     = @($Analysis.Whea | Where-Object { $_.Fatal }).Count
    $corrWhea      = @($Analysis.Whea | Where-Object { -not $_.Fatal }).Count

    # Findings
    $fRows = New-Object System.Text.StringBuilder
    if ($Findings.Count -eq 0) {
        [void]$fRows.Append("<tr><td colspan='4'>No crashes, hangs or hardware errors in the window, and crash dumps are configured.</td></tr>")
    } else {
        foreach ($f in $Findings) {
            $badge = Get-SeverityClass -Severity $f.Severity
            [void]$fRows.Append(
                "<tr><td><span class='tk-badge-$badge'>$(EscHtml $f.Severity)</span></td>" +
                "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
                "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
                $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
                "<td>$(EscHtml $f.Remedy)</td></tr>"
            )
        }
    }

    # Next steps by area
    $gRows = New-Object System.Text.StringBuilder
    if ($Analysis.Areas.Count -eq 0) {
        [void]$gRows.Append("<tr><td colspan='2'>No crash evidence to act on.</td></tr>")
    } else {
        foreach ($a in $Analysis.Areas) {
            $text = if ($AreaGuidance.Contains($a)) { $AreaGuidance[$a] } else { $AreaGuidance['Unknown'] }
            [void]$gRows.Append("<tr><td><strong>$(EscHtml $a)</strong></td><td>$(EscHtml $text)</td></tr>")
        }
    }

    # Incidents
    $iRows = New-Object System.Text.StringBuilder
    if ($Analysis.Incidents.Count -eq 0) {
        [void]$iRows.Append("<tr><td colspan='6'>No unplanned stops in the last $($Analysis.LookbackDays) day(s).</td></tr>")
    } else {
        foreach ($i in $Analysis.Incidents) {
            $badge = if ($i.Type -eq 'Forced power-off') { 'warn' } else { 'err' }
            $code  = if ($i.Key) { "<span class='tk-mono'>$(EscHtml $i.Key)</span><br/>" } else { '' }
            [void]$iRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $i.Time.ToString('yyyy-MM-dd HH:mm'))</td>" +
                "<td><span class='tk-badge-$badge'>$(EscHtml $i.Type)</span></td>" +
                "<td>$code$(EscHtml $i.Name)" + $(if ($i.Parameters) { "<br/><span class='tk-mono'>$(EscHtml $i.Parameters)</span>" } else { '' }) + "</td>" +
                "<td>$(EscHtml $i.Area)</td><td>$(EscHtml $i.Source)</td><td class='tk-mono'>$(EscHtml $i.Dump)</td></tr>"
            )
        }
    }

    # Timeline
    $tRows = New-Object System.Text.StringBuilder
    if ($Analysis.Timeline.Count -eq 0) {
        [void]$tRows.Append("<tr><td colspan='3'>Nothing recorded in the window.</td></tr>")
    } else {
        foreach ($t in ($Analysis.Timeline | Select-Object -First 200)) {
            [void]$tRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $t.Time.ToString('yyyy-MM-dd HH:mm'))</td>" +
                "<td><span class='tk-badge-$($t.Class)'>$(EscHtml $t.Kind)</span></td><td>$(EscHtml $t.Detail)</td></tr>"
            )
        }
    }

    # Dump files
    $dRows = New-Object System.Text.StringBuilder
    if ($Analysis.Dumps.Count -eq 0) {
        [void]$dRows.Append("<tr><td colspan='5'>No dump files on disk.</td></tr>")
    } else {
        foreach ($d in $Analysis.Dumps) {
            $bc = if ($d.Info) { "<span class='tk-mono'>$(EscHtml $d.Key)</span> $(EscHtml $d.Info.Name)" } else { EscHtml $d.Format }
            [void]$dRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $d.Path)</td><td>$(EscHtml $d.Written.ToString('yyyy-MM-dd HH:mm'))</td>" +
                "<td>$(EscHtml (Format-Bytes $d.SizeBytes))</td><td>$(EscHtml $d.Format)</td><td>$bc</td></tr>"
            )
        }
    }

    # WHEA
    $wRows = New-Object System.Text.StringBuilder
    if ($Analysis.Whea.Count -eq 0) {
        [void]$wRows.Append("<tr><td colspan='4'>No hardware errors reported.</td></tr>")
    } else {
        foreach ($w in ($Analysis.Whea | Select-Object -First 100)) {
            $badge = if ($w.Fatal) { 'err' } else { 'warn' }
            $label = if ($w.Fatal) { 'Fatal' } else { 'Corrected' }
            [void]$wRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $w.Time.ToString('yyyy-MM-dd HH:mm'))</td><td><span class='tk-badge-$badge'>$label</span> $($w.Id)</td>" +
                "<td>$(EscHtml $w.Component)</td><td>$(EscHtml $w.Message)</td></tr>"
            )
        }
    }

    $stabilityText  = if ($null -ne $Analysis.Stability) { "$($Analysis.Stability)" } else { 'n/a' }
    $stabilityClass = if ($null -eq $Analysis.Stability) { 'info' } elseif ($Analysis.Stability -lt 5) { 'warn' } else { 'ok' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'N.E.C.R.O.P.S.Y. Crash Analysis Report' `
        -ScriptName 'N.E.C.R.O.P.S.Y.' `
        -Subtitle   "${orgPrefix}Crash & Unexpected-Reboot Analysis -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Model'     = $Analysis.Device.Model
            'Generated' = $reportDate
            'Window'    = "Last $($Analysis.LookbackDays) day(s)"
            'Verdict'   = $Verdict.Verdict
            'Last boot' = $Analysis.Device.LastBoot
        }) `
        -NavItems   @('Findings', 'Next Steps', 'Incidents', 'Timeline', 'Dump Files', 'Hardware Errors', 'Configuration')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'N.E.C.R.O.P.S.Y. v5.1'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(if ($bugcheckCount -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$bugcheckCount</div><div class="tk-summary-lbl">Bugchecks</div></div>
    <div class="tk-summary-card $(if ($powerCount -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$powerCount</div><div class="tk-summary-lbl">Power Losses / Forced Off</div></div>
    <div class="tk-summary-card $(if ($fatalWhea -gt 0) { 'err' } elseif ($corrWhea -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($fatalWhea + $corrWhea)</div><div class="tk-summary-lbl">Hardware Errors</div></div>
    <div class="tk-summary-card $(if ($Analysis.Events.Tdr.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Analysis.Events.Tdr.Count)</div><div class="tk-summary-lbl">Display Resets</div></div>
    <div class="tk-summary-card $stabilityClass"><div class="tk-summary-num">$stabilityText</div><div class="tk-summary-lbl">Stability Index</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Analysis.Events.Boots.Count)</div><div class="tk-summary-lbl">Boots in Window</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
        <tbody>$($fRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Next Steps</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Area</th><th>What to do</th></tr></thead>
        <tbody>$($gRows.ToString())</tbody>
      </table>
      <div class="tk-info-box">
        <span class="tk-info-label">Scope</span> N.E.C.R.O.P.S.Y. reads the bugcheck code and parameters, not the stack. To name the faulting driver, open a dump from section 05 in WinDbg and run <span class="tk-mono">!analyze -v</span>.
      </div>
    </div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Incidents</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Type</th><th>Bugcheck / cause</th><th>Area</th><th>Source</th><th>Dump</th></tr></thead>
        <tbody>$($iRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Timeline</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Event</th><th>Detail</th></tr></thead>
        <tbody>$($tRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Dump Files</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Path</th><th>Written</th><th>Size</th><th>Format</th><th>Bugcheck (from header)</th></tr></thead>
        <tbody>$($dRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Hardware Errors</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Type / ID</th><th>Component</th><th>Message</th></tr></thead>
        <tbody>$($wRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s07">
    <div class="tk-section-title"><span class="tk-section-num">07</span> Configuration</div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Operating system</span> $(EscHtml $Analysis.Device.OS)<br/>
        <span class="tk-info-label">Last boot</span> $(EscHtml $Analysis.Device.LastBoot) (up $(EscHtml $Analysis.Device.Uptime))<br/>
        <span class="tk-info-label">Dump type</span> $(EscHtml $Analysis.Config.DumpLabel)<br/>
        <span class="tk-info-label">Automatic restart</span> $(if ($Analysis.Config.AutoReboot) { 'On' } else { 'Off (stays on the blue screen)' })<br/>
        <span class="tk-info-label">Dump file</span> <span class="tk-mono">$(EscHtml $Analysis.Config.DumpFile)</span><br/>
        <span class="tk-info-label">Minidump folder</span> <span class="tk-mono">$(EscHtml $Analysis.Config.MinidumpDir)</span><br/>
        <span class="tk-info-label">Dedicated dump file</span> $(if ($Analysis.Config.Dedicated) { "<span class='tk-mono'>$(EscHtml $Analysis.Config.Dedicated)</span>" } else { '(none)' })<br/>
        <span class="tk-info-label">Page file</span> <span class="tk-mono">$(EscHtml $Analysis.Config.PageFiles)</span>
      </div>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-NecropsyRun {
    param([int]$LookbackDays)

    $Findings.Clear()

    $analysis = Invoke-NecropsyAnalysis -LookbackDays $LookbackDays
    Show-NecropsyFindings

    $verdict = Get-NecropsyVerdict -FindingList @($Findings)
    Write-Section "VERDICT"
    switch ($verdict.Class) {
        'err'   { Write-Fail $verdict.Verdict }
        'warn'  { Write-Warn $verdict.Verdict }
        default { Write-Ok   $verdict.Verdict }
    }
    if ($analysis.Areas.Count -gt 0) {
        Write-Info "Most implicated: $($analysis.Areas[0])"
        Write-Info $AreaGuidance[$(if ($AreaGuidance.Contains($analysis.Areas[0])) { $analysis.Areas[0] } else { 'Unknown' })]
    }

    Add-TKNote -Text ("NECROPSY on {0} (last {1} day(s)): verdict {2}; {3} incident(s), {4} finding(s)." -f `
        $env:COMPUTERNAME, $LookbackDays, $verdict.Verdict, $analysis.Incidents.Count, $Findings.Count) -Category 'Info' -ScriptName 'necropsy'

    Write-Step "Generating HTML report..."
    $html      = Build-NecropsyReport -Analysis $analysis -Verdict $verdict
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "NECROPSY_${timestamp}.html"

    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'necropsy' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-NecropsyBanner
    Invoke-NecropsyRun -LookbackDays $Days
} else {
    $choice = ''

    do {
        Show-NecropsyBanner

        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Analyse the last $Days day(s)  -  read-only + HTML report" -ForegroundColor $C.Info
        Write-Host "  [2] Analyse a different window  -  enter the number of days" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-NecropsyRun -LookbackDays $Days }
            '2' {
                Write-Host -NoNewline "  Days to look back (1-365): " -ForegroundColor $C.Header
                $n = 0
                if ([int]::TryParse((Read-Host).Trim(), [ref]$n) -and $n -ge 1 -and $n -le 365) {
                    Invoke-NecropsyRun -LookbackDays $n
                } else {
                    Write-Warn "Enter a whole number from 1 to 365."
                }
            }
            'Q' {
                Write-Host ""
                Write-Host "  Closing N.E.C.R.O.P.S.Y." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1, 2 or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -ne 'Q') {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
