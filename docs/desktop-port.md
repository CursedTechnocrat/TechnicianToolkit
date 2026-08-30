# Desktop Port — turning the toolkit into one portable application

**Status:** phase 03 landed · **Target release:** 5.0 — *bringing it together to be portable*

> **Phase 00 outcome — the approach works.** A single-file self-contained WPF app
> hosting PowerShell 7 ran `ward.ps1` end to end with zero error records: it
> audited the machine's local accounts and wrote its HTML report, with the
> toolkit's colored output, prompts, progress, and streams all arriving through a
> custom host. It needs two non-obvious build settings that nothing in the error
> messages points at. See [`app/spike/README.md`](../app/spike/README.md) for the
> build recipe and the full findings; the corrections it forced are folded into
> the sections below.

The suite ships today as 42 standalone PowerShell scripts plus a prototype `.exe`
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
| Signing | **Certum Open Source Code Signing** on SimplySign — certificate already held |
| Architectures | `win-x64` **and** `win-arm64`, both built in CI; both signed locally |
| Versioning | Everything to **5.0** |

Nothing is left open. The rationale for each choice is in the sections below.

---

## Where the repo stands

`launcher/` (added in #22) produced a self-contained single-file `.exe` that
embedded every script and shelled out to `powershell.exe` to run the GRIMOIRE
console menu. It was the proof that a one-file distribution works, and it did its
job. Its resource-embedding and extraction logic was lifted into the app as
`ScriptExtractor`, and the directory has since been **deleted**.

What changes is everything above and below the extractor: the console hub becomes
a window, and the child `powershell.exe` process becomes a runspace living inside
the app.

| | |
|---|---|
| 42 | tool scripts, ~31k lines of PowerShell — all of it kept |
| 3,546 | `Write-Host` calls that must land in a GUI output pane |
| 222 | `Read-Host` prompts that must become dialogs |
| 40 / 42 | tools that already accept `-Unattended` |
| 29 | tools emitting an HTML report to open afterwards |
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

**Paid, in phase 02.** `app/TechnicianToolkit.App/Theme.xaml` templates every
control a technician touches -- scrollbars, buttons, text fields, dropdowns,
checkboxes, list rows -- and takes its palette verbatim from the tk custom
properties the HTML reports already use. The app and the reports it produces
read as one product, and nothing is left at a WPF default.

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

One trap for anyone editing `app.manifest` or `app.asinvoker.manifest`: Windows
parses them as XML, so a double hyphen inside an XML comment makes the manifest
invalid and the application dies at startup with *"the side-by-side configuration
is incorrect"* — which names neither the file nor the reason. The build does not
catch it.

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

This list was **not** complete as first written: it covered the WMI and EventLog
families but not the rest of what left `Microsoft.PowerShell.Management` in the
same release, and it prescribed a cmdlet swap at two sites where a swap alone
would have failed. It has since been re-swept against the removed-cmdlet set
(`*-WmiObject`, `*-EventLog`, `*-Computer`, `*-ComputerRestore*`,
`New-WebServiceProxy`) and the `[WMI]` / `[WMICLASS]` / `[WMISEARCHER]` type
accelerators. `Add-Computer` is the only blocker that sweep found still standing.

| Site | Problem | Fix | Severity |
|---|---|---|---|
| `augur.ps1:154`, `auspex.ps1:690`, `citadel.ps1:156`, `forge.ps1:136,380,383`, `sigil.ps1:608` | `Get-WmiObject` was removed in PowerShell 7 — 7 sites across 5 files | Replaced with `Get-CimInstance`; `-Namespace` carries over unchanged | ~~must fix~~ **DONE** |
| `sigil.ps1:622` | `$adapter.SetTcpipNetbios(2)` — a WMI *instance method*. A CIM instance carries no methods, so swapping the cmdlet alone would have left a call that throws at runtime | Rewritten as `Invoke-CimMethod -InputObject $adapter -MethodName SetTcpipNetbios` | ~~missed by the original scan~~ **DONE** |
| `gargoyle.ps1:332–344` | `Get-EventLog` was removed in PowerShell 7 — 4 sites | Replaced with `Get-WinEvent -FilterHashtable`, using `Level = 2` for exact parity with `-EntryType Error`, plus the property remap below | ~~must fix~~ **DONE** |
| `gargoyle.ps1:352–366` | `Get-WinEvent` returns `TimeCreated` / `LevelDisplayName` / `ProviderName` / `Id`, not `TimeGenerated` / `EntryType` / `Source` / `EventID`. A straight cmdlet swap would have emptied the console table and the HTML report without erroring | Remapped at the two consumer loops, keeping the emitted shape identical; `Message` is now null-guarded because `Get-WinEvent` leaves it null when a provider resource DLL is missing | ~~missed by the original scan~~ **DONE** |
| `covenant.ps1:304,307,529,532` | **`Add-Computer` was removed in PowerShell 7** — 4 sites, all in the AD domain-join path | Not yet fixed. See *The site the first scan missed* below | must fix |
| `TechnicianToolkit.psm1:32` | `[Console]::OutputEncoding` throws `"The handle is invalid"` when no console is attached | Wrap in `try/catch` | ~~crashes app~~ **DONE** |
| `covenant.ps1:984–985`, `restoration.ps1:360–361` | `[Console]::KeyAvailable` / `ReadKey` in press-a-key-to-skip loops | Each countdown now probes `[Console]::KeyAvailable` once in a `try/catch` and degrades to a plain wait. The banner is conditional too, so a console-less host no longer promises an Escape key that cannot be pressed | ~~must fix~~ **DONE** |
| `scryer.ps1:108` | `[Console]::Clear()` bypasses the host | Now `Clear-Host`, which the custom host implements | ~~must fix~~ **DONE** |
| `grimoire.ps1` ×4 | Same `[Console]::Clear()`, but in the hub the app replaces | Leave alone — console mode still uses it | no action |

**Severity corrected by the spike.** `[Console]::OutputEncoding` was rated
"crashes app". It does not: it surfaces as a *non-terminating error record* and
the module still imports. It was fixed anyway — the module runs it at import
time, so without the fix all 42 tools open with a spurious error. Windows
PowerShell 5.1 was re-checked afterwards and still sets UTF-8 when a console is
attached, so the standalone script path does not regress.

The `[Console]::KeyAvailable` sites are guarded but still **not exercised**: they
sit behind reboot countdowns that only a full COVENANT or RESTORATION run reaches.
The guard is written to be safe either way — it probes the call it is about to
make rather than inferring from the host name — but it has not been watched
running.

### The site the first scan missed

`Add-Computer` is removed in PowerShell 7, and `covenant.ps1` calls it at four
sites to join a machine to an Active Directory domain. The original sweep looked
for the WMI and EventLog families and did not cover the `*-Computer` cmdlets that
left `Microsoft.PowerShell.Management` in the same release.

It is deliberately **not** fixed yet, because none of the options is a mechanical
swap and all of them change a credentialed, destructive operation:

- **`Invoke-CimMethod` on `Win32_ComputerSystem.JoinDomainOrWorkgroup`.** Native
  and dependency-free, but it takes the password as plain text and the join
  behaviour is driven by an `FJoinOptions` bitmask rather than named switches.
  It needs testing against a real domain before it can be trusted.
- **`Import-Module Microsoft.PowerShell.Management -UseWindowsPowerShell`.** Keeps
  `Add-Computer` verbatim, but runs it in a Windows PowerShell 5.1 compatibility
  session — a separate runspace, so its output and prompts do not travel through
  the custom host. It also quietly reintroduces a dependency on the machine own
  PowerShell, which is the thing hosting the engine was meant to remove.
- **Leave the AD-join path console-only for 5.0** and have the app surface it as
  unavailable, deferring the rewrite.

This is the one decision in the port that a domain is required to validate.

### Found while scanning

`forge.ps1:383` reads `Where-Object { $_.Name -eq $_.DeviceName }` — inside that
block `$_` is the inner `Win32_PnPEntity`, so it compares an object to itself
instead of to the outer driver. It is a pre-existing bug unrelated to this port,
and worth a separate fix so the two changes stay reviewable apart.

---

## Removing the launcher

**Done.** Smaller than expected, as predicted: every `launcher` match in
`README.md` and `CLAUDE.md` turned out to refer to `grimoire.ps1` as the *hub
launcher*, not to the directory, so removal touched nothing outside its own
directory but the workflow and one test comment.

| Touches | Action |
|---|---|
| `launcher/` (4 files) | ✅ Deleted, after lifting `Program.cs`'s extraction logic into `ScriptExtractor` |
| `.github/workflows/release-launcher.yml` | ✅ Deleted in the same commit. Its **replacement is deliberately not written yet** — see below |
| `tests/…Tests.ps1:355–361` | ✅ BOM gate kept, comment rewritten around the 5.1 standalone path |
| `README.md` | Nothing to remove; add the app instead |

**There is no release workflow on `main` right now, and that is intentional.**
`release-app.yml` is phase 04 work and its whole point is signed artifacts.
Shipping an unsigned single-file `.exe` that unpacks scripts and runs them
elevated is precisely the SmartScreen and antivirus problem this plan calls
fatal, so a stopgap workflow would be worse than none. Until phase 04, tagging
`v*` publishes nothing, which is correct — phase 03 landed the window and the
run history, but the 5.0 bump, the winget package and the release path itself
are all still ahead.

The reason has changed shape since this was written, and it is now permanent
rather than a wait. There are no CI signing credentials to write against and
there never will be: the certificate exists, but its key lives in SimplySign and
cannot be reached from a hosted runner at all (see *Signing* below). So
`release-app.yml` is not blocked on paperwork arriving — it is waiting on phase
04 to define a workflow whose output is deliberately **unsigned**, with signing
as a documented manual step afterwards.

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
existing header, so all 42 scripts move forward and no reader has to be told about
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
  `Version\s*:` instead, and verify all 42 files afterwards.
- **`restoration.ps1` is three-part (`3.6.2`)** where every other header is
  two-part, so a pattern anchored to `3.6` alone will leave a stray `.2`.

### Add a gate so this cannot drift again

The current spread — 3.6, 3.6.2, 4.2, 1.0, and one with irregular spacing — is
exactly what a test prevents. Add a Pester `Describe` asserting that each tool's
`.NOTES Version` matches its `Version` field in the GRIMOIRE registry. Both are
already read by AST elsewhere in the suite, so it is cheap.

---

## Signing — Certum Open Source Code Signing, on SimplySign

**The certificate already exists**, which removes the longest external lead time
this plan had. It is Certum's Open Source Code Signing certificate with the key
held in **SimplySign**, Certum's cloud HSM (FIPS 140-2 Level 3 / CC EAL 4+). The
subject is a natural person prefixed `Open Source Developer`, so the publisher
string on the UAC prompt is a developer name rather than an organization. Say so
in the README — a personal name where users expect a company reads as suspicious
only when it is left unexplained.

**This costs the plan its automated signing step.** The key is non-exportable by
design — that is the point of the CA/B hardware requirement — so it cannot be
imported into a signing service, and reaching it needs the SimplySign Desktop
client to mount a virtual smart card against a session authenticated with the
SimplySign mobile app. Ephemeral GitHub-hosted runners cannot do that. Certum has
said CI/CD support is planned, with no date attached.

Three ways out, in the order they were considered:

| Approach | Verdict |
|---|---|
| **Build in CI, sign locally, attach the signed binaries** | **Chosen.** One manual step per release |
| Self-hosted runner holding a SimplySign session open | Rejected for 5.0 — an always-on Windows box, with sessions that expire, to save one step on a release that happens a few times a year |
| Script the OTP by extracting the TOTP secret | Rejected. Fragile, and it defeats the second factor the certificate's assurance rests on |

So the release is **built in public CI and signed on the maintainer's machine**:

1. `release-app.yml` builds `win-x64` and `win-arm64` and uploads both as workflow
   artifacts — unsigned, and *not* attached to the release yet.
2. Download both. Start SimplySign Desktop and authenticate with the mobile app;
   the certificate appears in the Windows certificate store through the virtual
   card.
3. Sign each with **timestamping** — not optional. The certificate is short-lived,
   and a timestamped signature stays valid after it expires where an
   un-timestamped one dies with it:

   ```
   signtool sign /n "Open Source Developer" /fd SHA256 ^
       /tr http://time.certum.pl /td SHA256 TechnicianToolkit.exe
   ```

4. `signtool verify /pa /v` each binary, then attach both to the tagged release.
   The winget manifest carries the artifact hashes, so it is generated *after*
   signing — a hash taken from the unsigned artifact will not match.

Signing appends to the PE rather than rebuilding it, so the published binary is
still the CI-built one plus a signature. The public-CI provenance survives the
manual step.

**ARM64 needs no separate arrangement.** `signtool` signs any PE regardless of its
target architecture, so both binaries sign from the same x64 machine in the same
session. That retires ARM64 as a *signing* problem; it remains an *execution*
problem (see Risks).

**It is not an EV certificate.** Certum sells EV separately, and this product's own
description is that it "supports building" SmartScreen reputation — which is what
a non-EV certificate does. EV grants reputation immediately; OV accrues it over
downloads and time. Early 5.0 downloads may therefore still meet a SmartScreen
warning even though the binary is correctly signed. Signing remains worth it: it
makes the publisher identifiable, it removes the "Unknown publisher" prompt, and
it is the precondition for reputation ever accruing. But do not promise
technicians a clean first-run experience on day one.

*Practical limit:* 5,000 signatures per month. A release signs two files.

---

## Groundwork for phase 03

Four things were fixed before starting phase 03, because its features would
have been built directly on top of them.

### A run that did nothing reported success

`ToolRunResult.Succeeded` returned true for WARD refusing to start without
Administrator: no error records, no exit code, nothing to distinguish it from a
clean run. Phase 03's run history is specified as *what ran, when, exit status,
and which artifacts it produced* — built on that, it would have recorded a
success for a tool that did no work.

Two changes, because there are two separate problems hiding in one symptom:

- **Refusal is now determined before the run.** A tool that asserts
  Administrator while the process is not elevated is reported as
  `RefusedNeedsAdmin`. It has to be decided up front, because a refusal leaves
  nothing behind to read afterwards.
- **The exit code is now actually captured.** The tool is invoked through a
  wrapper that reads `$LASTEXITCODE` in a `finally` block. A script's `exit`
  never reaches `SetShouldExit`, but it does set `$LASTEXITCODE`, and the
  `finally` runs even when the tool exits.

The exit code is deliberately **not** part of `Succeeded`. Five tools end on
`robocopy`, `manage-bde` or `netsh`, and robocopy returns 1 for *files were
copied* — treating non-zero as failure would invent failures for ARCHIVE and
REVENANT. It is recorded as information for the run history, not as a verdict.

### Reports were written in among the extracted scripts

`LogDirectory` is empty in the shipped `config.json`, so `Resolve-LogDirectory`
fell back to each tool's own directory — which is where the suite gets
extracted. Phase 03 wants to *watch* that directory for new reports, and
watching a folder where 45 files reappear on every launch is the wrong
foundation. It was also the cause of the extractor miscounting its own output.

The layout is now explicit: `TechnicianToolkit\suite` for the extracted tools,
`TechnicianToolkit\reports` for what they produce, and the app writes the
report path into the extracted `config.json` so the toolkit's own
`Resolve-LogDirectory` honours it. An existing non-empty `LogDirectory` is left
alone, because a technician who set one through HEARTH meant it.

`config.json` is also no longer overwritten on launch — it is seeded when
missing and preserved thereafter. It was being replaced from the embedded copy
every time, which would have quietly wiped whatever the phase 03 settings
screen had just written.

### The working directory was resolved from the wrong place

Found while dry-running the CI gate. `AppContext.BaseDirectory` and the
executable's own directory are the same thing in a normal build and *different*
in the configuration that ships: `IncludeAllContentForSelfExtract` is mandatory
for the PowerShell SDK under single-file, and it makes `BaseDirectory` point at
the bundle's extraction folder under `%TEMP%` — a hashed path that changes with
every build and gets cleaned up.

So the published application was extracting the suite, and would have written
every report, into a temporary directory no technician could find. That defeats
both the USB-stick promise and the report-watching phase 03 is built around.
The root is now resolved from `Environment.ProcessPath`, with
`LocalApplicationData` rather than `%TEMP%` as the fallback, so reports survive
even when the medium is read-only.

### Output was dropped silently

The output pane keeps the last 20,000 lines. For one chatty tool that is fine;
for the queue of them phase 03 introduces it is not, because the first tool's
output would vanish and *Save* would then write an incomplete log with nothing
to say so. Trimming now leaves a visible notice line carrying the count.

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
- ✅ **WARD end to end, elevated, zero error records** — audited 6 local
  accounts and wrote a valid 13 KB HTML report. This also confirms the
  requestedExecutionLevel decision (`Assert-AdminPrivilege` passes rather than
  exiting), that `-Unattended` suppresses the prompt (`prompt calls: 0`), and
  that the shared HTML report path works untouched
- ⬜ **Clean VM.** Weaker than recorded, but still open. This machine turns out
  to have **no PowerShell 7 installed at all** — no `pwsh` on `PATH`, nothing under
  `Program Files\PowerShell` — and the phase 00 probe reported `$PSHOME` resolving
  into the single-file extraction directory, not a system install. Phase 01 then
  ran two whole tools on it. That is good evidence the engine travels in the
  bundle; a genuinely clean VM is still the honest proof.
- ⬜ **ARM64 on real hardware.** It bundles; it has never been executed.
- ✅ **Signing certificate obtained** — Certum Open Source Code Signing on
  SimplySign. This was the longest-lead item in the plan and it is closed. It
  arrives with a constraint rather than for free: no unattended CI signing.
  See *Signing* below.

Both remaining boxes are *verification* — clean VM, ARM64 hardware. They are
cheap and should be closed before phase 02 commits to the architecture, but
neither blocks starting phase 01.

### 01 — Engine · 1–2 weeks · **exit criteria met**

Headless and testable, with no window yet. Lives in
[`app/TechnicianToolkit.Engine`](../app/TechnicianToolkit.Engine/) with the console
harness in [`app/TechnicianToolkit.Harness`](../app/TechnicianToolkit.Harness/).

- ✅ Resource extraction lifted out of `launcher/Program.cs` into `ScriptExtractor`.
  It embeds and writes out 45 files — 42 tools, the module, `config.json` and the
  licence — byte for byte, so the UTF-8 BOM every script carries survives
- ✅ Full `PSHost` / `PSHostUserInterface` / `PSHostRawUserInterface`, with the
  streams handled asynchronously and `[securestring]` fields answered without ever
  materialising a managed string
- ✅ AST readers for the `$Tools` catalog and for each tool's `param()` block.
  The catalog reads all 41 registry entries out of `grimoire.ps1`; the parameter
  reader turns `ValidateSet` into a dropdown, `ValidateScript` into a path picker
  and `[securestring]` into a masked field, with no per-tool knowledge in C#
- ✅ The compatibility fixes above, except the newly found `Add-Computer` sites
- ✅ `launcher/` deleted, together with `release-launcher.yml`, which was the only
  thing that built it. No release workflow replaces it until phase 04 can sign

**Exit:** a console harness runs any tool by name with parameters, streams its
output live, and cancels it mid-run. **Met**, and verified on this machine:

| Check | Result |
|---|---|
| Extraction | 45 files written, every BOM intact |
| Catalog by AST | 41 tools read from the `grimoire.ps1` registry, grouped by its own `$CategoryOrder` |
| Parameters by AST | CIPHER renders its 8-value `ValidateSet` as a dropdown and its `ValidatePattern` verbatim |
| A tool end to end | EXHUME ran unelevated through the hosted engine, colours and sections intact, and wrote its HTML report |
| Cancel mid-run | `PowerShell.Stop()` interrupted EXHUME mid-scan at 3.2s; harness exit 130 |
| Elevation refusal | WARD prints its refusal and stops. `exit` inside a module function never reaches `SetShouldExit`, so the refusal is detected before the run — see *Groundwork for phase 03* |

### 02 — The window · 2–3 weeks · **exit criteria met**

In [`app/TechnicianToolkit.App`](../app/TechnicianToolkit.App/). It owns no toolkit
logic: the catalog, the forms, the host and the runner all come from the engine.

- ✅ Catalog pane: all 41 tools grouped by the hub's own `$CategoryOrder`, with
  search across name and description
- ✅ Detail pane: the form generates itself from each tool's `param()` block, with
  Run and Cancel and REQUIRES ADMIN / CHANGES THIS MACHINE / READ ONLY badges read
  from the script rather than maintained by hand
- ✅ Output pane preserving console colours, with copy, save and clear
- ✅ Prompt dialogs wired to `ReadLine`, `ReadLineAsSecureString`,
  `PromptForCredential` and `PromptForChoice`
- ✅ Manifest requesting Administrator, app icon, version metadata

**Exit:** every tool in the catalog is runnable from the GUI.

| Check | Result |
|---|---|
| Catalog and badges | 41 tools, correctly grouped and badged, from the registry and the scripts |
| Generated form | CIPHER renders its `ValidateSet` as a dropdown, its `ValidatePattern` live-validated, defaults pre-filled |
| A tool end to end in the GUI | EXHUME ran through the hosted engine and its output arrived with full colour fidelity |
| Published single file | 84 MB `win-x64` one-file `.exe` ran the same tool end to end, so the two mandatory build settings carry over |
| Elevation state | Shown in the title bar, because a tool refusing for want of Administrator is otherwise silent |

**Honestly still unproven.** Only EXHUME has actually been executed from the
window. The path is uniform -- every tool goes through the same catalog entry,
the same generated form and the same runner -- but most of the suite needs
Administrator or a cloud sign-in, so "every tool is runnable" is an argument from
uniformity, not 41 observations. The prompt dialogs are wired and compile but no
live `Read-Host` has yet opened one, since `-Unattended` is ticked by default and
suppresses exactly that. Auto-scrolling the output to the tail is likewise
unverified: it does not take effect in the headless render, which may be a
limitation of rendering a window that was never shown rather than a defect.

### Seeing the window without a person in the loop

`TechnicianToolkit.exe --screenshot <path> [--tool NAME] [--run]` lays the window
out and renders it to a PNG without ever showing it. It follows the precedent
phase 00 set with `--probe`: a GUI has no console to report into, so the result
goes to a file.

With `--run` it executes the named tool through the real engine first, so the
capture is of genuine output rather than staged text. This is how the layout and
theme were reviewed, and it is what CI should use to catch a visual regression.
It found two defects that a build cannot: the detail pane painting no background
of its own, and every output line rendering empty because the segments were
surfaced through a dependency property but mutated in place, so WPF compared the
same list reference against itself and concluded nothing had changed.

Phase 03 extended it: `--pane <output|reports|history|queue|settings>` chooses
what to capture, and `--queue A,B` stages a queue that `--run` then executes for
real. `--pane settings` renders the settings dialog, which switching panes
cannot reach. That is how the settings screen was found unable to read its own
configuration — it built its script from a string, where `$PSScriptRoot` is
null, so `Join-Path` failed with a binding error naming neither the file nor the
variable.

Build with `-p:Elevate=false` to get the asInvoker manifest and run any of this
from an unelevated shell.

### 03 — What makes it a program rather than a launcher · 1–2 weeks · **exit criteria met**

The window's lower half became four panes: OUTPUT, REPORTS, HISTORY, QUEUE.

- ✅ **Report handling.** The report directory is listed and watched, and an
  artifact opens in the default application on double-click. The watcher only
  says *something changed*; the list itself always comes from a directory
  listing, because a watcher can coalesce or miss events under load and a report
  that never appeared would be worse than one that appeared slowly
- ✅ **Run history.** Tool, time, duration, outcome, exit code, parameters and the
  artifacts produced, persisted to `history.json` beside the suite. Artifacts are
  attributed by diffing the report directory across the run rather than by
  watching it during: a tool can write several files, none, or rewrite one it
  produced earlier, and a diff answers all three
- ✅ **Settings.** A form over `config.json`, reading through `Get-TKConfig` and
  writing through `Set-TKConfig` rather than touching the file. `hearth.ps1` is
  untouched and remains the console way to do the same thing
- ✅ **Queue.** Tools are queued with the parameters captured at the moment they
  were added, so the same tool can be queued twice with different arguments.
  Cancelling stops the queue rather than only the tool in flight — a technician
  who cancels a workup means the workup

**Exit:** a technician can run a machine's full workup and collect every report
without leaving the window. **Met**, and verified end to end:

| Check | Result |
|---|---|
| A real queue | EXHUME then WARD queued and run in sequence through the window |
| History | EXHUME recorded `OK` with its HTML report attached, 14.2s; WARD recorded `REFUSED`, no artifacts, exit 1 |
| Reports | Both runs' artifacts listed with tool, size and timestamp, newest first |
| Queue | Three tools queued, each carrying its own captured parameters — CIPHER kept `-Action Status -Drive C` |
| Settings read | The form is populated from `Get-TKConfig`, report directory included |
| Settings write | A top-level key and a sectioned key both round-trip through `Set-TKConfig` and survive re-extraction |

That WARD row is the pre-phase-03 work paying off. Before it, the same run would
have been filed as a success.

**Still unproven.** The same two things as phase 02, and for the same reason: no
live `Read-Host` has opened a prompt dialog, and output auto-scroll has never
been watched working. Both need a person at the keyboard. Nothing in phase 03
changed either.

### 04 — Ship 5.0 · 1 week

- `release-app.yml`: build `win-x64` and `win-arm64` and publish both as unsigned
  workflow artifacts. Signing is the documented manual step that follows; the
  release attach and the winget manifest both run after it, on the signed files
- A `RELEASING.md` checklist for that manual step, so it is not carried in one
  person's head: SimplySign session, `signtool` with timestamp, `signtool verify`,
  attach, then `wingetcreate` against the signed hashes
- winget manifest as `CursedTechnocrat.TechnicianToolkit`, installer type
  `portable`, both architectures, generated with `wingetcreate`
- The 5.0 version bump across all 42 scripts, the registry, and the `.csproj`
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
| ~~**high**~~ **retired** — PowerShell SDK with single-file publish | Both failure modes found and fixed in phase 00; neither was discoverable from its error message | Locked in via two build settings, and **now gated in CI**: the `Desktop app` job publishes single-file and runs `TechnicianToolkit.Harness.exe probe`, which opens a real runspace and resolves `Get-CimInstance`. Both failures live in exactly that path, so an SDK bump or a dropped `.csproj` property fails the build rather than the field |
| **low** — the clean-VM claim is unverified | Recorded as *the dev box has PowerShell 7 installed independently*. It does not — see phase 00 above — so the risk is smaller than it was written, but a machine that never had it is still the only real proof | One run on a VM with no PowerShell 7 before phase 02 |
| ~~**med**~~ **retired** — signing certificate lead time | Was the only item in the plan with external lead time | Certum Open Source Code Signing obtained. Superseded by the two rows below |
| **med** — signing cannot run in CI | SimplySign needs an interactive session with phone 2FA, which GitHub-hosted runners cannot hold. A release therefore carries a manual step, and a manual step can be skipped or botched | `RELEASING.md` checklist, and `signtool verify /pa` on both binaries before the release is published. Revisit if Certum ships CI/CD support |
| **med** — SmartScreen on early downloads | The certificate is OV, not EV, so reputation accrues rather than being granted. A correctly signed 5.0 can still warn on first run | Sign and timestamp from the first release so reputation starts accruing. Set the expectation in the README rather than treating the warning as a bug |
| **med** — Antivirus false positives | A single-file executable that unpacks scripts to disk and runs them elevated is, structurally, what a dropper looks like | Signing helps most. Submit to Microsoft and the major vendors for whitelisting ahead of the release |
| **med** — Prompt-heavy tools | `covenant.ps1` has 26 `Read-Host` calls and `citadel.ps1` has 17; a separate dialog for each is a miserable experience | Prefer `-Unattended` driven by the generated form. Treat modal prompts as the fallback path, not the primary one |
| **med** — ARM64 unverified on hardware | The prototype workflow built ARM64 but nothing ever ran it on a real device. Signing is *not* the gap — `signtool` signs any PE from the same x64 session | Run the ARM64 build on an actual ARM device before tagging, not merely confirm it compiles |
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
