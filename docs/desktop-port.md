# Desktop Port — turning the toolkit into one portable application

**Status:** phase 00 complete · **Target release:** 5.0 — *bringing it together to be portable*

> **Phase 00 outcome — the approach is viable.** A single-file self-contained WPF
> app hosting PowerShell 7 loads the engine, resolves `$PSHOME`, runs CIM, and
> carries the toolkit's console output, prompts, and streams through a custom
> host. It needs two non-obvious build settings that nothing in the error
> messages points at. See [`app/spike/README.md`](../app/spike/README.md) for the
> build recipe and the full findings; the corrections it forced are folded into
> the sections below.

The suite ships today as 41 standalone PowerShell scripts plus a prototype `.exe`
launcher. Version 5.0 replaces that arrangement with a single signed Windows
application: a real GUI over the existing engine, with PowerShell 7 hosted inside
the executable so nothing needs installing on the machine it runs on.

The scripts themselves are not going anywhere. They remain runnable standalone,
and they remain the engine — the application drives them, it does not replace them.

---

## Decisions locked

| Decision | Choice |
|----------|--------|
| Shape | WPF desktop application on .NET 8 |
| Engine | PowerShell 7 hosted in-process via `Microsoft.PowerShell.SDK` |
| Prototype launcher | **Removed** — `launcher/` is superseded by the app |
| Distribution | Portable single `.exe`, plus a winget package |
| Signing | **SignPath Foundation** free open-source tier |
| Architectures | `win-x64` **and** `win-arm64`, both built and signed in CI |
| Versioning | Everything to **5.0** |

Nothing is left open. The rationale for each choice is in the sections below.

---

## Where the repo stands

[`launcher/`](../launcher/) (added in #22) already produces a self-contained
single-file `.exe` that embeds every script and shells out to `powershell.exe` to
run the GRIMOIRE console menu. It was the proof that a one-file distribution
works, and it did its job. The app supersedes it, so it goes — but its
resource-embedding and extraction logic is lifted into the app first, not
rewritten.

What changes is everything above and below the extractor: the console hub becomes
a window, and the child `powershell.exe` process becomes a runspace living inside
the app.

| | |
|---|---|
| 41 | tool scripts, ~31k lines of PowerShell — all of it kept |
| 3,512 | `Write-Host` calls that must land in a GUI output pane |
| 220 | `Read-Host` prompts that must become dialogs |
| 39 / 41 | tools that already accept `-Unattended` |
| 28 | tools emitting an HTML report to open afterwards |
| 11 | cmdlet call sites blocking PowerShell 7 — the entire compatibility debt |

---

## Architecture

### WPF on .NET 8, not WinUI 3 or Avalonia

The toolkit is Windows-only by nature — WMI, the registry, Defender, `netsh`.
Cross-platform buys nothing. WPF publishes cleanly as self-contained single-file;
WinUI 3 drags in Windows App SDK bootstrapping that fights the
one-file-on-a-USB-stick requirement.

*Cost:* WPF's default look is dated. Budget real time for styling, or the app will
look older than the console it replaces.

### PowerShell 7 hosted in-process

A runspace inside the app gives live stream capture, real objects instead of
scraped text, cancellation through `PowerShell.Stop()`, and — via a custom
`PSHost` — the ability to answer `Read-Host` with a dialog. Shelling out to a child
process gives none of that. It also removes the last external dependency: the
target machine needs no PowerShell at all.

*Cost:* 11 cmdlet call sites to modernize, and an 83 MB executable — about half
the 150 MB originally budgeted.

**Two build settings are mandatory** (phase 00 established both; neither is
hinted at by the error you get without it):

1. `IncludeAllContentForSelfExtract=true`. Without it, constructing
   `PSHostUserInterface` throws `TypeInitializationException` on
   `PowerShellConfig` before a single script runs — PowerShell derives `$PSHOME`
   from `Assembly.Location`, which is empty under single-file. It fails inside
   the host constructor, so no script-level `try/catch` can reach it.
2. **Re-link PowerShell's built-in modules to `Modules\**`.** The SDK ships
   CimCmdlets, Utility, Management, Security, Diagnostics and WSMan as *content*
   under `contentFiles/any/any/runtimes/win/lib/net8.0/Modules/`. Assemblies
   flatten into the bundle root but content keeps its RID-relative path, so the
   modules extract somewhere PowerShell never looks and every built-in cmdlet
   fails to resolve — `Get-CimInstance` included, which 35 sites depend on. See
   the `RemapPowerShellModules` target in the spike's project file.

### Scripts extract to disk and run by path — never from a string

Every tool's bootstrap block depends on `$PSScriptRoot` and `$PSCommandPath`
resolving to a real location. Extracting first also means the module and every tool
are already present, so the download-from-GitHub-if-missing paths never fire — the
offline guarantee holds exactly as it does today.

### The manifest requests Administrator up front

`Invoke-AdminElevation` calls `Start-Process -Verb RunAs` and then `exit`. Inside a
hosted runspace that would spawn a stray elevated console window and abandon the
run. Elevating the whole app makes `Test-IsAdmin` return true, so that branch is
never reached and `Assert-AdminPrivilege` never error-exits.

*Cost:* one UAC prompt at launch, every launch, for every tool — including
read-only ones. Document it plainly rather than trying to be clever about it.

---

## How the console becomes a window

The custom `PSHost` is the load-bearing component of the whole port. All 3,512
`Write-Host` calls and all 220 prompts route through it — which is precisely why no
tool script needs rewriting to gain a GUI. **Verified in phase 00:** the module's
own `Write-Ok` / `Write-Warn` / `Write-Fail` / `Write-Info` / `Write-Step` /
`Write-Section` helpers all arrived with their colors intact, alongside
`Clear-Host`, `Read-Host`, `Write-Progress`, and the warning and verbose streams.

Two rows below were wrong before the spike and are corrected here: PowerShell 7
has no `RawUI.Clear()` — `Clear-Host` calls `SetBufferContents` with a rectangle
of all `-1` — and a script's `exit` never reaches `SetShouldExit` at all.

| What the script does | Where it lands in the app | Implemented by |
|---|---|---|
| `Write-Host -ForegroundColor` | Colored line in the run output pane | `PSHostUserInterface.Write` |
| `Read-Host` | Modal input dialog; the run pauses | `UI.ReadLine` |
| `Read-Host -AsSecureString` | Masked password dialog | `UI.ReadLineAsSecureString` |
| Yes / No / Cancel prompts | Button row beneath the question | `UI.PromptForChoice` |
| `Write-Progress` | Determinate progress bar on the run | `UI.WriteProgress` |
| `Clear-Host` (37 sites) | Clears the output pane | `RawUI.SetBufferContents` |
| Warning / Error / Verbose streams | Severity-tinted lines; errors collected for the run summary | `PSDataCollection` handlers |
| `Start-Transcript` | Degrades to a no-op; the app logs the run itself | already wrapped in `try/catch` |

---

## Forms build themselves from the param blocks

The tools are already richly typed. Parsing each `param()` block with the
PowerShell AST yields a form for free, and one that can never drift from the script
it describes. The same trick supplies the catalog: `grimoire.ps1`'s `$Tools` array
stays the single source of truth for names, versions, and categories, read by AST
rather than duplicated in C#. The Pester suite already extracts data tables this
way, so the technique has precedent here.

| Declared as | Renders as | Seen in |
|---|---|---|
| `[switch]$Unattended` | Checkbox, ticked by default — the form supplies what the prompts would have asked | 39 tools |
| `[switch]$WhatIf` | Checkbox, given prominence on the destructive set | 9 tools |
| `[ValidateSet('Status','Enable',…)]` | Dropdown of exactly the valid values | `cipher.ps1`, `ritual.ps1` |
| `[ValidatePattern('^[A-Za-z]:?$')]` | Text field validating live against the regex | `covenant.ps1`, `cipher.ps1` |
| `[ValidateScript({ Test-Path … })]` | Path field with a browse button | `revenant.ps1`, `ritual.ps1` |
| `[securestring]$LocalAdminPassword` | Masked field; never held as plain text | `covenant.ps1` |

---

## The compatibility debt, in full

Moving from Windows PowerShell 5.1 to PowerShell 7 costs far less than it sounds.
This is the complete list — every call site, found by scanning the tree. Nothing
else in the suite blocks the move.

| Site | Problem | Fix | Severity |
|---|---|---|---|
| `augur.ps1:154`, `auspex.ps1:690`, `citadel.ps1:156`, `forge.ps1:136,380,383`, `sigil.ps1:608` | `Get-WmiObject` was removed in PowerShell 7 — 7 sites across 5 files | Replace with `Get-CimInstance`; `-Namespace` carries over unchanged | must fix |
| `gargoyle.ps1:332–344` | `Get-EventLog` was removed in PowerShell 7 — 4 sites | Replace with `Get-WinEvent -FilterHashtable`, which keeps `-ComputerName` | must fix |
| `TechnicianToolkit.psm1:32` | `[Console]::OutputEncoding` throws `"The handle is invalid"` when no console is attached | Wrap in `try/catch` | ~~crashes app~~ **DONE** |
| `covenant.ps1:984–985`, `restoration.ps1:360–361` | `[Console]::KeyAvailable` / `ReadKey` in press-a-key-to-skip loops | Guard on a host-capability check; skip the poll entirely when hosted | must fix |
| `scryer.ps1:108` | `[Console]::Clear()` bypasses the host | Use `Clear-Host`, which the custom host implements | must fix |
| `grimoire.ps1` ×4 | Same `[Console]::Clear()`, but in the hub the app replaces | Leave alone — console mode still uses it | no action |

**Severity corrected by the spike.** `[Console]::OutputEncoding` was rated
"crashes app". It does not: it surfaces as a *non-terminating error record* and
the module still imports. It was fixed anyway — the module runs it at import
time, so without the fix all 41 tools open with a spurious error. Windows
PowerShell 5.1 was re-checked afterwards and still sets UTF-8 when a console is
attached, so the standalone script path does not regress.

The `[Console]::KeyAvailable` sites are **not** yet verified either way; they sit
behind interactive branches the spike never reached. Treat their severity as
unknown until a prompt-heavy tool is actually run.

### Found while scanning

`forge.ps1:383` reads `Where-Object { $_.Name -eq $_.DeviceName }` — inside that
block `$_` is the inner `Win32_PnPEntity`, so it compares an object to itself
instead of to the outer driver. It is a pre-existing bug unrelated to this port,
and worth a separate fix so the two changes stay reviewable apart.

---

## Removing the launcher

Smaller than expected: `launcher/` was never documented in `README.md`, so removal
touches almost nothing outside its own directory.

| Touches | Action |
|---|---|
| `launcher/` (4 files) | Delete, after lifting `Program.cs`'s extraction logic into the app |
| `.github/workflows/release-launcher.yml` | Replace with `release-app.yml` (signing + ARM64 + winget, below) |
| `tests/…Tests.ps1:355–361` | Update the BOM test's comment — see the note below |
| `README.md` | Nothing to remove; add the app instead |

**Keep the UTF-8 BOM gate.** Its comment currently justifies itself by saying "the
launcher invokes powershell.exe 5.1, so it hits this every run." That reason
disappears with the launcher — the app hosts PS7, which assumes UTF-8 — but the
*gate* must stay: the scripts remain runnable standalone under Windows PowerShell
5.1, which is still the primary documented path and still reads a BOM-less file as
ANSI. Rewrite the comment, keep the test.

---

## The 5.0 bump

Version 5.0 marks the release where the suite stops being a folder of scripts and
becomes one portable application. Everything moves to 5.0 in the same commit so
there is a single coherent line.

**Why 5.0 and not 4.0.** `cipher.ps1` is already at 4.2 — it drifted ahead of the
rest of the suite. Unifying on 4.0 would have versioned one script *backwards*,
which is the one thing a version number must never do. Starting at 5.0 clears every
existing header, so all 41 scripts move forward and no reader has to be told about
an exception. The toolkit simply never has a 4.x line, and the CHANGELOG jumps
`[3.6.4] → [5.0.0]`; say so in one line in that entry so the gap reads as
deliberate rather than as a missing release.

Convention: script headers and the registry use two-part (`5.0`); the CHANGELOG and
the `.csproj` use three-part (`5.0.0`), matching what each already does.

### What needs touching

| Where | Now | To |
|---|---|---|
| 37 script headers, `.NOTES Version : 3.6` | `3.6` | `5.0` |
| `grimoire.ps1` `$Tools` registry — 40 entries | 39 at `'3.6'`, 1 at `'1.0'` | all `'5.0'` |
| `tendril.ps1:64` + its registry entry (Key 46) | `1.0` | `5.0` |
| `restoration.ps1:38` | `3.6.2` | `5.0` |
| `cipher.ps1:43` | `4.2` | `5.0` |
| New `.csproj` `<Version>` | — | `5.0.0` |
| `CHANGELOG.md` | `[3.6.4]` | new `## [5.0.0]` entry |

### Two traps in the mechanical edit

- **`talisman.ps1:41` reads `Version  : 3.6` with two spaces.** A naive
  `sed 's/Version : 3.6/Version : 5.0/'` silently skips it. Match on
  `Version\s*:` instead, and verify all 41 files afterwards.
- **`restoration.ps1` is three-part (`3.6.2`)** where every other header is
  two-part, so a pattern anchored to `3.6` alone will leave a stray `.2`.

### Add a gate so this cannot drift again

The current spread — 3.6, 3.6.2, 4.2, 1.0, and one with irregular spacing — is
exactly what a test prevents. Add a Pester `Describe` asserting that each tool's
`.NOTES Version` matches its `Version` field in the GRIMOIRE registry. Both are
already read by AST elsewhere in the suite, so it is cheap.

---

## Signing — SignPath Foundation

The free open-source tier. GPL-3.0-or-later qualifies, the repository is public,
and the build runs on GitHub Actions, which is the CI SignPath Foundation requires
so that the signed artifact is reproducible from public source.

Setup, in order:

1. Apply to the SignPath Foundation open-source programme. **Do this first** — it
   is an application with a review, and it is the only item in this plan with
   external lead time.
2. Once approved, record the organization ID, project slug, and signing-policy slug,
   and store the API token as a repository secret.
3. In `release-app.yml`, after the build and before the release attach, submit each
   artifact with `signpath/github-action-submit-signing-request` and download the
   signed binary back.
4. Sign both architectures. The winget manifest points at the signed artifacts.

Note that SignPath Foundation signs with *its* certificate, not one attributed to
this project — that is the trade for it being free, and it is the normal
arrangement for open-source tooling.

---

## Phases

### 00 — Spike: prove the runtime · ~~2–3 days~~ **DONE**

Everything downstream assumed something Microsoft does not officially support.
Settled in [`app/spike/`](../app/spike/).

- ✅ Single-file self-contained WPF + `Microsoft.PowerShell.SDK` publishes and
  starts — once the two build settings above are applied
- ✅ Engine loads: PowerShell 7.4.6 (Core), `$PSHOME` resolves, `Get-CimInstance`
  works
- ✅ The host carries the toolkit's colored output, prompts, progress, and
  streams — driven through the real module's own helpers
- ✅ 83 MB `win-x64` / 80 MB `win-arm64`, one `.exe` and nothing beside it
- ⬜ **WARD end to end.** It reaches its admin gate cleanly with zero error
  records, but the probe build is `asInvoker`, so the audit never runs. Needs one
  elevated run.
- ⬜ **Clean VM.** This machine has PowerShell 7 installed independently, so
  "the `.exe` carries its own engine" is not yet honestly proven.
- ⬜ **ARM64 on real hardware.** It bundles; it has never been executed.
- ⬜ **SignPath Foundation application** — still to submit, still the longest
  lead time in the plan.

The three open boxes are cheap and should be closed before phase 02 commits to
the architecture, but none of them blocks starting phase 01.

### 01 — Engine · 1–2 weeks

Headless and testable, with no window yet.

- Lift resource extraction out of `launcher/Program.cs` into the app, then delete
  `launcher/`
- Full `PSHost` / `PSHostUserInterface` / `PSHostRawUserInterface` implementation,
  with async events for every stream
- AST readers for the `$Tools` catalog and for each tool's `param()` block
- Apply the compatibility fixes above

**Exit:** a console harness runs any tool by name with parameters, streams its
output live, and cancels it mid-run.

### 02 — The window · 2–3 weeks

- Catalog pane: categories from the registry, search across name and description
- Detail pane: generated parameter form, Run and Cancel, admin-required and
  destructive badges
- Output pane preserving console colors, with copy and save
- Prompt dialogs wired to the host's `ReadLine` and `PromptForChoice`
- Manifest requesting Administrator; app icon; version metadata

**Exit:** every tool in the catalog is runnable from the GUI. This is the first
build worth showing anyone.

### 03 — What makes it a program rather than a launcher · 1–2 weeks

- Report handling: watch the configured log directory, surface new HTML reports,
  open on click
- Run history: what ran, when, exit status, and which artifacts it produced
- Settings screen over `config.json` using `Get-TKConfig` / `Set-TKConfig` —
  `hearth.ps1` stays for console users
- Queue several tools in sequence, mirroring what `ritual.ps1` does for recipes

**Exit:** a technician can run a machine's full workup and collect every report
without leaving the window.

### 04 — Ship 5.0 · 1 week

- `release-app.yml`: build `win-x64` and `win-arm64`, sign both through SignPath,
  attach both to the tagged release
- winget manifest as `CursedTechnocrat.TechnicianToolkit`, installer type
  `portable`, both architectures, generated with `wingetcreate`
- The 5.0 version bump across all 41 scripts, the registry, and the `.csproj`
- Extend the Pester license gate to `.cs` and `.xaml`; add the version-consistency
  gate; add xUnit tests for the AST readers and the extractor
- `README.md` and `CLAUDE.md` rewritten around two ways to run the suite: the app,
  and the scripts standalone
- `CHANGELOG.md` entry for `[5.0.0]`

**Exit:** a tagged release publishes signed `win-x64` and `win-arm64` binaries, and
`winget install` works.

---

## Risks

| Risk | Why it bites | Response |
|---|---|---|
| ~~**high**~~ **retired** — PowerShell SDK with single-file publish | Both failure modes found and fixed in phase 00; neither was discoverable from its error message | Locked in via two build settings. Keep the spike's `--probe` check in CI so a dependency bump cannot silently reintroduce either |
| **low** — the clean-VM claim is unverified | The dev box has PowerShell 7 installed independently, so a stray dependency on it would not show up here | One run on a VM with no PowerShell 7 before phase 02 |
| **med** — SignPath application lead time | It is an application with a review, not a purchase | Submitted during phase 00. Unsigned builds trip SmartScreen, which is fatal for a tool run on client machines |
| **med** — Antivirus false positives | A single-file executable that unpacks scripts to disk and runs them elevated is, structurally, what a dropper looks like | Signing helps most. Submit to Microsoft and the major vendors for whitelisting ahead of the release |
| **med** — Prompt-heavy tools | `covenant.ps1` has 26 `Read-Host` calls and `citadel.ps1` has 17; a separate dialog for each is a miserable experience | Prefer `-Unattended` driven by the generated form. Treat modal prompts as the fallback path, not the primary one |
| **med** — ARM64 has no signed precedent here | The prototype workflow built ARM64 but nothing ever verified it on real hardware | Test the ARM64 build on an actual ARM device before tagging, not just that it compiles |
| **low** — Everything runs elevated | Read-only tools like `ward.ps1` get Administrator they do not need | Accept for 5.0 and say so in the README. A split-process design is a later refinement |
| **low** — Binary size | Roughly 150 MB against about 40 MB today | Compression stays on. Trimming stays off — it breaks the reflection the SDK depends on |

---

## Deliberately not in scope

- **Porting tool logic to C#.** The existing PowerShell is tested and works.
  Rewriting 31,000 lines trades a working suite for an unproven one and discards the
  Pester coverage that guards it.
- **Retiring the script path.** `grimoire.ps1` and the individual scripts keep
  working standalone. The app is a second way to run the suite, not a replacement.
- **An in-app report viewer.** WebView2 means another runtime dependency on a tool
  that prizes having none. Reports open in the default browser; revisit later.
- **MSI or MSIX packaging.** Portable plus winget covers the field-technician case.
  Worth adding if managed deployment is ever asked for.

---

## Settled

- **Version line starts at 5.0**, not 4.0, so that `cipher.ps1`'s existing 4.2 moves
  forward like everything else. See *The 5.0 bump* above.
- **The app takes the top-level `app/` directory**, with no `src/` restructure —
  a question that only existed while `launcher/` was staying.

Nothing is blocking phase 00.
