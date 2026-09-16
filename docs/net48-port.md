# net48 Port — moving the app onto what Windows already has

**Status:** planning · **Target release:** 5.1 — *the app stops carrying a runtime*

The 5.0 application is a self-contained .NET 8 single-file executable with
PowerShell 7 hosted inside it. That was the right call to prove the port, and it
works. This plan moves it to **.NET Framework 4.8 hosting the inbox Windows
PowerShell 5.1** — both of which are already on every Windows 11 machine.

The motivation is not size, though size improves considerably. It is that the
application currently runs the suite on an engine the suite is not documented or
tested against, and pays a compatibility debt for the privilege — a debt that
still has one unpaid item blocking a shipping feature.

> **Nothing about the scripts changes.** They stay UTF-8 BOM, 5.1-targeted, and
> independently runnable. The already-completed PS7 compatibility fixes stay too
> — see *No reverts* below. This is a change to `app/` only.

---

## What is actually inbox on Windows 11

| Present on a clean install | Not present |
|---|---|
| .NET Framework 4.8 (21H2) / 4.8.1 (22H2+, native ARM64) | Any .NET 5/6/7/8/9/10 runtime |
| Windows PowerShell 5.1 | PowerShell 7 |
| WebView2 Evergreen Runtime (Win11 only) | Windows App SDK / WinUI 3 |

.NET Framework is serviced by Windows Update and supported for the life of the
OS. This is why framework-dependent .NET 8 was never an option: the Desktop
Runtime is a prerequisite download on every clean machine, which breaks the
promise the winget listing makes twice — *installs nothing on the machine it
runs on*.

---

## Decisions to lock

| Decision | Proposed | Confidence |
|---|---|---|
| Target framework | `net48` across Engine, App, Harness, Tests | settled |
| PowerShell | Inbox Windows PowerShell 5.1, hosted in-process | settled |
| Compile-time PS reference | `Microsoft.PowerShell.5.ReferenceAssemblies` (compile-only, ships nothing) | settled |
| Bitness | AnyCPU with **`Prefer32Bit=false`** — see *The bitness trap* | settled, and load-bearing |
| Architectures | One binary. The `win-x64`/`win-arm64` matrix collapses | settled |
| JSON | Newtonsoft.Json (one assembly) over System.Text.Json (six + binding redirects) | **decide** — alternative below |
| Single-file packaging | Costura.Fody embeds the dependency into the exe | **decide** — depends on the JSON choice |
| Version | App and package to 5.1.0; script `Version :` lines stay 5.0 | **decide** — breaks 5.0 lockstep |
| Rename to SANCTUM | Deferred, but see *Sequencing with the rename* | open |

---

## What this deletes

**The entire compatibility debt.** `docs/desktop-port.md` carries a table of 11
call sites that PowerShell 7 broke. Ten are fixed; one is not:

> `covenant.ps1:304,307,529,532` — **`Add-Computer` was removed in PowerShell 7** —
> 4 sites, all in the AD domain-join path. *Not yet fixed.*
>
> "This is the one decision in the port that a domain is required to validate."

On 5.1, `Add-Computer` is inbox and the AD-join path works unmodified. The
blocker does not get solved, it stops existing. No domain needed to validate a
rewrite that no longer happens.

**The divergence between the two halves.** `CLAUDE.md` states that 5.1 "remains
the primary documented path." Today the app runs PS7 and the documented path runs
5.1 — two engines, two test suites, and neither suite sees the other's
regressions. After this, the app runs what the scripts are documented against.

**Three mandatory build hacks**, all of which exist only to make PowerShell 7
survive single-file publish:

- `IncludeAllContentForSelfExtract` — the `$PSHOME`-from-`Assembly.Location`
  `TypeInitializationException` that no script-level catch can reach.
- The `RemapPowerShellModules` target — relocating built-in modules the SDK ships
  as RID-relative content, without which every inbox cmdlet fails to resolve.
- `DropRidPathModules` — cleaning up after the above.

Also gone: `PublishTrimmed=false`, `InvariantGlobalization=false`, the `IL3000`
and `NETSDK1179` suppressions, and the `ResolveRoot` workaround for
`AppContext.BaseDirectory` pointing into a hashed TEMP bundle — .NET Framework
has no bundle, so `Assembly.Location` is simply correct.

**The untested-ARM64 risk.** A single AnyCPU binary runs native on ARM64 under
4.8.1. The CI matrix, the second SHA-256 in the release notes, and the "help
wanted, never run on real hardware" caveat all collapse to one artifact.

**Most of the binary.** Roughly 150 MB to low single-digit MB. Worth measuring
rather than trusting: publish the current app and compare. The secondary effect
matters as much — a small conventional executable is far less AV-provocative than
a 150 MB self-extracting bundle that unpacks scripts and runs them elevated,
which is the exact shape flagged in the README. It does not remove the need for
signing.

---

## What it costs

**You inherit the machine's PowerShell.** Today the executable carries its own
engine and cannot be affected by host policy. After this, WDAC or AppLocker
forcing Constrained Language Mode cannot be bypassed by
`InitialSessionState.ExecutionPolicy` the way execution policy can. For the
stated use case — freshly imaged machines, minimal hardening — this is
acceptable, and the standalone script path already carries the same exposure.
It should be documented in the README rather than discovered in the field.

**.NET Framework WPF is in maintenance.** No new features, but fully supported
and no upgrade treadmill. Note the counterpoint: **.NET 8 LTS support ends
10 November 2026**, so staying put is not a no-op — it means moving to .NET 10
instead. Verify the date against Microsoft's current policy page before quoting it.

**Cloud module drift is the real long-term risk.** See *Risks*.

---

## The port surface

~6,000 lines of C# outside `app/spike/`. A scan for .NET 6+ APIs and C# 9+
language features found only two things:

| Site | Issue | Fix |
|---|---|---|
| `ScriptExtractor.cs:228` | `Environment.ProcessPath` is .NET 6+ | `Assembly.GetEntryAssembly().Location`. The long comment above it about single-file bundle extraction becomes moot and should be rewritten, not just retargeted |
| `Output.cs:39-40`, `MainWindow.xaml.cs:49-51,68-70`, `ToolkitLayout`, `ReportArtifact`, `RunRecord`, `ToolRunResult` | `{ get; init; }` needs C# 9 + a runtime type net48 lacks | `LangVersion 9`+ and a five-line internal `IsExternalInit` shim in the Engine |
| `ScriptExtractor.cs:179-183`, `ToolkitConfig.cs`, `RunHistory.cs` | `System.Text.Json` / `JsonNode` / `JsonStringEnumConverter` | See *The JSON decision* |

Everything else compiles as-is: `is not` patterns, switch expressions,
`using var`, target-typed `new`, `async`/`await`,
`Task.Factory.FromAsync(BeginInvoke, EndInvoke)`, `#nullable enable`. Expect more
nullable warnings, since the .NET Framework BCL is not annotated.

**`ToolkitHost` and `ToolRunner` port close to verbatim.** `PSHost`,
`PSHostUserInterface`, `PSHostRawUserInterface`, `Runspace`,
`InitialSessionState.CreateDefault2()` and `PowerShell.Stop()` are the same APIs
on 5.1. This is the part that looked expensive and is not.

---

## The bitness trap

**This is the one that will silently corrupt output if missed.**

.NET Framework executables default to `Prefer32Bit=true` under AnyCPU. A 32-bit
process on 64-bit Windows gets WOW64 registry redirection (`Wow6432Node`) and
File System redirection (`SysWOW64`). The suite would then read a different
machine than the one in front of it — SIGIL writes baseline registry values,
TALON reads autorun keys, ANVIL and AUGUR read firmware and disk data. Nothing
would error. The reports would just be wrong.

```xml
<PlatformTarget>AnyCPU</PlatformTarget>
<Prefer32Bit>false</Prefer32Bit>
```

That combination gives a 64-bit process on x64 and a native ARM64 process on
ARM64 under 4.8.1 — one binary, correct bitness everywhere, and a hosted
PowerShell that matches what `powershell.exe` gives by default.

Add a test asserting `Environment.Is64BitProcess` on an x64 CI runner so this
cannot regress silently.

---

## The JSON decision

`System.Text.Json` supports net48 via netstandard2.0, but drags in six
transitive assemblies (`System.Memory`, `System.Buffers`,
`System.Runtime.CompilerServices.Unsafe`, `System.Text.Encodings.Web`,
`System.Threading.Tasks.Extensions`, `System.Numerics.Vectors`) and the binding
redirects on net48 are a known source of runtime `FileLoadException`s.

Three options, in the order I would try them:

1. **Newtonsoft.Json** — one assembly, net48-native, no redirect pain,
   Costura-friendly. `JsonNode`/`JsonObject` become `JObject`;
   `JsonStringEnumConverter` becomes `StringEnumConverter`. Three files touched.
   *Recommended.*
2. **Inbox serialisation only** — `JavaScriptSerializer`
   (`System.Web.Extensions`, inbox) or `DataContractJsonSerializer`. Clunkier
   code, but the app ends with **zero** NuGet runtime dependencies, Costura
   becomes unnecessary, and the exe is a genuinely single dependency-free file.
   Most elegant end state; most rewriting.
3. **Keep System.Text.Json** and let Costura absorb the transitive set. Least
   code change, most runtime risk.

The JSON choice determines whether Costura is needed at all, which is why the two
decisions are linked in the table above.

---

## No reverts

The compatibility fixes already made for PS7 are **all valid on 5.1** and must
not be reverted:

- `Get-CimInstance` / `Invoke-CimMethod` — PowerShell 3.0+
- `Get-WinEvent -FilterHashtable` — PowerShell 2.0+
- `Clear-Host` in place of `[Console]::Clear()` in `scryer.ps1`
- The `[Console]::OutputEncoding` and `[Console]::KeyAvailable` guards

Only `Add-Computer` changes status, from *blocked* to *works*. The
`TechnicianToolkit.psm1` `#Requires -Version 5.1` line is already correct.

**Verify, do not assume:** `Clear-Host` on 5.1 is a function that calls
`SetBufferContents` with an all-`-1` rectangle, the same signal the current
`ToolkitRawUi` override keys on — so all 37 `Clear-Host` sites are expected to
work unchanged. It also sets `CursorPosition` first, which the current
implementation exposes as a plain auto-property. Prove it in the spike.

`Start-Transcript` throws on a custom host under 5.1 exactly as under 7, and
`Start-TKTranscript` already wraps it in `try/catch`. No change.

---

## Build and packaging changes

**`app/TechnicianToolkit.App.csproj`** — delete `SelfContained`,
`PublishSingleFile`, `IncludeNativeLibrariesForSelfExtract`,
`EnableCompressionInSingleFile`, `IncludeAllContentForSelfExtract`,
`PublishTrimmed`, `InvariantGlobalization`, both remap targets, and the
`Microsoft.PowerShell.SDK` reference. Add `Prefer32Bit=false` and `LangVersion`.
Keep `UseWPF`, both manifests, and the `Elevate=false` switch — the manifest
mechanism is unchanged, including the trap that a double hyphen inside an XML
comment makes it invalid and kills startup with *"the side-by-side configuration
is incorrect."*

**`app.manifest`** — bump `assemblyIdentity version`. Consider adding a
`<supportedRuntime>`/`<startup>` element so a machine somehow lacking 4.8 fails
with a clear message rather than a silent nothing.

**Publishing** — `dotnet publish` for net48 produces an exe plus any dependency
DLLs; there is no `PublishSingleFile`. Costura.Fody (with `FodyWeavers.xml`)
folds them in and is the option that works reliably with WPF; ILRepack/ILMerge
fight BAML and XAML resource URIs and should be avoided here.

**winget** — `InstallerType: portable` and `PackageIdentifier` are unchanged.
The two architecture entries collapse to one. Decide between
`Architecture: neutral` (accurate for AnyCPU, verify winget validation accepts
it) and `x64` (safe, still runs on ARM64). One new `InstallerSha256`, one new
`InstallerUrl`.

**Local build note** — SDK-style net48 projects build only on Windows, with the
.NET Framework targeting pack. `windows-latest` has it. The Linux/sandbox
verification guidance in `CLAUDE.md` is unaffected, since `net8.0-windows` +
`UseWPF` could not be built off Windows either.

---

## CI changes

`.github/workflows/ci.yml`, `Desktop app` job:

- Keep `setup-dotnet` — the .NET SDK still drives SDK-style net48 builds.
- Drop `-r win-x64` and the single-file properties from both publish steps.
- The `probe`, catalog-read, config round-trip and `--screenshot` render steps
  are unchanged in intent. Probe is still the gate that proves a real runspace
  opens and executes.

`.github/workflows/release-app.yml`: the `[win-x64, win-arm64]` matrix collapses
to a single build, and the "ARM64 built but never run" caveat comes out of the
release notes, `CONTRIBUTING.md`, `README.md` and `docs/desktop-port.md`.

`TechnicianToolkit.Engine.Tests` moves to net48 — xunit 2.9.x and
`Microsoft.NET.Test.Sdk` 17.x both support it, and `dotnet test` works on
Windows. Add the `Is64BitProcess` assertion here.

---

## Phases

Deliberately mirrors the phase 00 structure that worked for the original port:
prove the runtime before touching anything else.

### Phase A — Spike: prove 5.1 hosting · ~1 day

A throwaway net48 WPF project that hosts a 5.1 runspace with the existing
`ToolkitHost` copied in, and runs `ward.ps1` end to end.

**Exit criteria** — matching what phase 00 demanded:
- WARD audits local accounts and writes its HTML report, zero error records.
- Colored `Write-Host` from `Write-Ok`/`Write-Warn`/`Write-Fail`/`Write-Info`/
  `Write-Step`/`Write-Section` arrives with colors intact.
- `Clear-Host` clears the pane (the `SetBufferContents` question above).
- `Read-Host`, `Read-Host -AsSecureString`, `Write-Progress`, and the warning and
  verbose streams all arrive.
- `Environment.Is64BitProcess` is true.
- **`Add-Computer` resolves** — `Get-Command Add-Computer` in the hosted runspace.
  This is the whole reason for the exercise; check it explicitly rather than
  inferring it from the version number.

### Phase B — Retarget the Engine · 1–2 days

Engine and Harness to net48. `IsExternalInit` shim, `Environment.ProcessPath`
replacement, the JSON decision applied. Tests to net48 and green.

**Exit criteria:** `TechnicianToolkit.Harness.exe probe`, `list`, and
`run WARD -Unattended` all behave as they do today, and the xUnit suite passes.

### Phase C — Retarget the App · 2–3 days

App to net48, build properties stripped, Costura wired if needed. Walk the panes:
catalog, generated forms, run, cancel, queue, history, reports, settings.

**Exit criteria:** the `--screenshot` render produces a correct window, and a
manual pass over the four run outcomes (`Succeeded`, `CompletedWithErrors`,
`Cancelled`, `RefusedNeedsAdmin`) records correctly in `history.json`.

### Phase D — Packaging, CI, docs · 1–2 days

Workflows, winget manifests, README, `docs/desktop-port.md` amended to record
that the PS7 debt is retired rather than paid.

**Exit criteria:** a tagged build produces one signed-shaped artifact, and
`winget install` from a local manifest works.

### Phase E — The COVENANT dividend · separate change

With `Add-Computer` available, close out the AD domain-join path and remove the
"unavailable in the app" caveat. Keep this a separate PR from the retarget so the
runtime change stays reviewable on its own — and it still wants a real domain to
test against, which is now a normal feature test rather than a port blocker.

---

## Risks

| Risk | Assessment |
|---|---|
| **Cloud module ecosystem drifting to PS7-only** | The real long-term threat, and the strongest argument against this plan. Seven Cloud & Identity tools depend on Az, Microsoft.Graph, ExchangeOnlineManagement and Teams modules. All currently support 5.1, but the direction of travel is PS7. Check each module's current minimum before committing, and accept that this may eventually force a split: 5.1 for the on-machine tools, something else for the cloud ones |
| Constrained Language Mode on hardened machines | Accepted for the stated use case. Document it |
| Module resolution changes | Under 5.1 hosting, modules a technician installed with Windows PowerShell resolve natively — likely an improvement over a hosted PS7 whose `$PSHOME` sits in an extraction directory, but verify rather than assume |
| `Clear-Host` behaving differently | Low, and phase A settles it |
| Costura + WPF | Low; well-trodden. Avoid ILMerge |
| Binding redirects | Avoided entirely by JSON option 1 or 2 |
| .NET Framework in maintenance | Accepted. No feature needs are outstanding, and the alternative carries a support deadline |

---

## Sequencing with the rename

The SANCTUM rename is deferred but not dropped. If both happen, **do them in the
same release**. Both change the shipped filename, and `UpgradeBehavior: install`
plus a changed filename means an upgrading user is left with the old
`TechnicianToolkit\` folder holding their reports and history. One migration —
`ResolveRoot` adopting a legacy folder when the new one is absent — covers both
changes if they land together, and has to be written twice if they do not.

---

## Deliberately not in scope

- **Porting tool logic to C#.** Unchanged from `docs/desktop-port.md`.
- **Retiring the script path.** Unchanged.
- **An in-app report viewer.** WebView2 is now inbox on Windows 11, which
  reopens the question — but `MinimumOSVersion` is Win10 1809, so it needs a
  browser fallback. Revisit separately.
- **Splitting the elevation model.** One UAC prompt at launch stays.
