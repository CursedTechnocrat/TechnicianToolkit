# Phase 00 spike — hosting PowerShell 7 in a single-file WPF app

**Throwaway code.** This project exists to answer one question before any of the
real application is written:

> Can a WPF app that hosts PowerShell 7 in-process be published as a
> self-contained single-file `.exe` and actually run a real toolkit script?

Microsoft does not officially support `Microsoft.PowerShell.SDK` together with
`PublishSingleFile`, which is why this was built and measured first.

**Answer: yes — but only with two non-obvious build settings.** Neither is
discoverable from the error messages. Both are recorded below.

## Result

| Check | Result |
|---|---|
| 1 — Engine loads, `$PSHOME` resolves, CIM works | **PASS** — PowerShell 7.4.6 (Core), `Get-CimInstance` returned the OS caption |
| 2 — A real toolkit script runs (WARD) | **Partial** — runs, imports the module, reaches its admin gate with zero error records; a full elevated run is still outstanding |
| 3 — The host carries Write-Host, prompts, streams, Clear-Host | **PASS** — 6 colors, `Clear-Host`, `Read-Host`, 2 progress records, warning + verbose streams |
| Single file, both architectures | **PASS** — 83 MB `win-x64`, 80 MB `win-arm64`, one `.exe` and nothing beside it |

Check 3 is the one that matters most, and it is deliberately independent of
Administrator rights — see below.

---

## Building

```bash
dotnet publish TechnicianToolkit.Spike.csproj -c Release -r win-x64 \
    -o ./dist -p:IncludeAllContentForSelfExtract=true
```

`win-arm64` builds the same way. The result is one `.exe` and nothing else.

To verify a build without clicking anything:

```bash
./dist/TechnicianToolkit.Spike.exe --probe report.txt   # exit 0 = engine works
```

The probe runs both checks headlessly and writes a report. A WinExe has no
console to write to, which is why the results go to a file — this is also what
makes the spike runnable from CI on a clean runner.

Running the `.exe` with no arguments opens the window instead: two buttons, a
colored output pane, and a progress line.

---

## What had to be discovered

### 1. Plain single-file publish fails before any script runs

Without `IncludeAllContentForSelfExtract`, constructing `PSHostUserInterface`
throws:

```
System.TypeInitializationException: The type initializer for
'System.Management.Automation.Configuration.PowerShellConfig' threw an exception.
   at System.Management.Automation.Host.PSHostUserInterface..ctor()
```

`PowerShellConfig` looks for `powershell.config.json` relative to `$PSHOME`,
which PowerShell derives from `Assembly.Location` — and that is an empty string
in single-file mode. The failure happens in the host constructor, so it is not
something a script-level `try/catch` can reach.

`IncludeAllContentForSelfExtract=true` makes the bundle extract to a real
directory on disk, `$PSHOME` resolves to it, and the engine loads.

### 2. The built-in modules land in the wrong place

Fixing the first problem exposes a second:

```
The 'Get-CimInstance' command was found in the module 'CimCmdlets', but the
module could not be loaded ... Cannot find the built-in module 'CimCmdlets'
that is compatible with the 'Core' edition.
```

The SDK ships PowerShell's built-in modules (CimCmdlets, Utility, Management,
Security, Diagnostics, WSMan, PSDiagnostics) as **content** files under a
RID-relative path:

```
contentFiles/any/any/runtimes/win/lib/net8.0/Modules/**
```

Assemblies get flattened into the bundle root, but content keeps its relative
path — so the modules extract to `runtimes\win\lib\net8.0\Modules\` while
PowerShell only ever looks in `$PSHOME\Modules\`. Every built-in cmdlet fails to
resolve, `Get-CimInstance` included, which 35 sites across the suite depend on.

The `RemapPowerShellModules` target in the `.csproj` re-links them to
`Modules\**`. `DropRidPathModules` then removes the stranded RID-path copies,
including the `unix` set, which is dead weight in a Windows-only app.

This is the fix that actually makes the approach viable, and nothing in the
error message points at it.

---

## Findings that change the plan

### `[Console]::OutputEncoding` is noisy, not fatal

`TechnicianToolkit.psm1:32` throws `Exception setting "OutputEncoding": "The
handle is invalid."` in a GUI host — as predicted. But it surfaces as a
**non-terminating error record**, not a crash: the module still imports and the
script continues. The plan rated this "crashes app"; the accurate severity is
"one error record at the top of every single run".

It still needs the fix, because it is the first thing every tool does and it
would put a spurious error at the head of all 41 tools' output. Fixed in this
branch by wrapping the assignment in `try/catch`. Verified afterwards that
Windows PowerShell 5.1 still sets UTF-8 when a console *is* attached — the
`catch` only swallows the headless case, so the standalone script path does not
regress.

### Script `exit` never reaches `SetShouldExit`

The plan assumed `Assert-AdminPrivilege`'s `exit 1` would arrive at
`PSHost.SetShouldExit` and need absorbing. It does not. In a hosted runspace,
`exit` inside a script simply ends the pipeline — the probe recorded
`exit calls: 0` while WARD's admin message printed normally.

`SetShouldExit` is still implemented (a host must), but this removes a class of
worry: a script calling `exit` cannot take the application down with it.

### `Clear-Host` arrives via `SetBufferContents`, not `Clear`

PowerShell 7 implements `Clear-Host` as a call to
`RawUI.SetBufferContents(rectangle, fill)` with `Top`/`Bottom`/`Left`/`Right`
all `-1`. There is no `RawUI.Clear()` to override. The plan's host-mapping table
names the wrong member; `SpikeRawUi.SetBufferContents` shows the real one.

### Size is roughly half the estimate

| Target | Single-file `.exe` |
|---|---|
| `win-x64` | 83 MB |
| `win-arm64` | 80 MB |

The plan budgeted ~150 MB. Compression is on; trimming stays off, since it
strips the reflection metadata the engine resolves cmdlets through.

### Minor

`ControlKeyStates` has no `None` member — use `default(ControlKeyStates)`.

---

## Two manifests

`app.manifest` requests `requireAdministrator`, which is the shipping
configuration: it makes `Test-IsAdmin` return true so `Assert-AdminPrivilege`
and `Invoke-AdminElevation` never take their failure branches.

`app.asinvoker.manifest` is test-only. It exists so the probe can run from an
unelevated shell without a UAC prompt, which is what makes automated
verification possible. Select it with
`-p:ApplicationManifest=app.asinvoker.manifest`.

---

## The host surface check (check 3)

The plan's central claim is that all 3,512 `Write-Host` calls and 220 prompts
route through the host, which is why no tool script needs rewriting to gain a
GUI. Check 3 verifies exactly that, driving the **real module's** console
helpers rather than a synthetic stand-in:

```
colors seen   : 6 (Cyan, Gray, Green, Magenta, Red, Yellow)
Clear-Host    : 1
prompt calls  : 1
progress calls: 2
VERDICT       : PASS
```

`Write-Ok` / `Write-Warn` / `Write-Fail` / `Write-Info` / `Write-Step` /
`Write-Section` all arrive with their colors intact, `Clear-Host` reaches
`SetBufferContents`, `Read-Host` reaches `ReadLine`, `Write-Progress` reaches
`WriteProgress`, and the warning and verbose streams come through their own
handlers.

This check needs no Administrator rights, which is deliberate: it means the
load-bearing component can be regression-tested in CI, where nothing is
elevated.

---

## What this spike does *not* prove

- **WARD has not run end to end.** It reaches its `Assert-AdminPrivilege` gate
  cleanly, with zero error records, and prints the refusal in red through the
  host — but the probe build is `asInvoker`, so the audit itself never executes.
  The elevated run was attempted and the UAC prompt was declined, so this is
  still open. To close it: run `dist-ship\TechnicianToolkit.Spike.exe --probe
  report.txt` and accept the prompt, or run the `asInvoker` build from an
  already-elevated shell.
- It has only run on the development machine, which has PowerShell 7 installed
  independently. The claim that the `.exe` carries its own engine needs a clean
  VM with no PowerShell 7 to be confirmed properly.
- `win-arm64` compiles and bundles, but has not been executed on ARM hardware.
- Only WARD has been targeted. It is a read-only tool with one
  `-Unattended`-guarded prompt; the prompt-heavy and destructive tools are
  untested here.
