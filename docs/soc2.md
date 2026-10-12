# Running the toolkit inside SOC 2 scope

This guide is for an MSP or internal IT team whose SOC 2 examination covers the
tools its technicians run on client machines. It says what the toolkit does that
an auditor will ask about, and which settings and procedures put it under your
controls.

## What SOC 2 means for a tool like this

SOC 2 is an attestation about an **organization's** controls, not a certificate a
piece of software holds. The toolkit cannot be "SOC 2 compliant" by itself, and
nobody can certify it as such. What an auditor examines is how *your* team gets,
changes, runs and protects the output of the software it uses. Treat the toolkit
like any other administrative tool in scope: a component your controls govern.

It is open-source software under GPL-3.0-or-later with no vendor, no SLA and no
warranty (LICENSE §§15–16). For vendor management (CC9.2), record it as
open-source software you maintain an approved copy of, not as a supplier.

The sections below follow the Trust Services Criteria an auditor is most likely
to test against the toolkit.

## 1. Change management (CC8.1) — run only the release you approved

**The risk.** Every tool is self-contained: it carries its own copy of the shared
helpers and fetches nothing to run. But the scripts that launch *other* tools
fetch a missing one from GitHub at run time, so a run can execute code nobody on
your team reviewed. Three paths do it, and each downloads from the repository's
`main` branch, not from a release:

| Path | Fetches |
|---|---|
| GRIMOIRE's tool launcher | A tool script, when it is not beside `grimoire.ps1` |
| RITUAL's step resolver | A recipe step's tool script, when it is not beside `ritual.ps1` |
| The 6.0 forwarding stubs (`archive.ps1`, `cipher.ps1`, …) | The renamed tool's script |

**The control.** Set the machine environment variable **`TK_DISABLE_DOWNLOAD=1`**
on every machine your technicians run the toolkit from, pushed by GPO or your RMM.
With it set, all three paths refuse to download and exit with a message saying
which file is missing. Only files you deployed can run. The Pester suite checks
that every fetch path honours the variable, so a new one cannot be added without
it.

```powershell
# Machine-wide; new processes pick it up
[Environment]::SetEnvironmentVariable('TK_DISABLE_DOWNLOAD', '1', 'Machine')
```

**The procedure.**

1. **Get a release, not a branch.** Download a tagged release from the GitHub
   Releases page. Never deploy a copy of `main`, a copy someone emailed, or the
   `irm … | iex` quick-launch snippet.
2. **Verify it.** The desktop application is Authenticode-signed from 6.0.0 on.
   Check it with `Get-AuthenticodeSignature` and compare it with the release's
   `SHA256SUMS.txt`. The loose `.ps1` files are not signed upstream. If
   your policy needs signed scripts, sign the approved copy with your own
   code-signing certificate (`Set-AuthenticodeSignature`) and enforce
   `AllSigned` execution policy on technician machines.
3. **Review the change.** Read the release's `CHANGELOG.md` section before
   approving an upgrade, and record the approval like any other change.
4. **Deploy the whole set together** to a location technicians can read but not
   write, such as a share with read-only ACLs for technicians or the desktop
   application. A single tool runs on its own, but GRIMOIRE and RITUAL launch
   the others from beside themselves; with the download switch on, a missing
   tool stops the run instead of being fetched.
5. **Keep current.** Only the latest release gets security fixes (see
   `SECURITY.md`). Watch the repository's releases and security advisories, and
   remove old copies (USB sticks, jump boxes) when you upgrade.

The desktop application embeds every script at build time, so its tools never
use the self-fetch paths. It is the simplest way to run one verified, signed set.

## 2. Third-party code the tools install (CC8.1, CC9.2)

Separate from the toolkit's own code, some tools install software as part of
their job. `TK_DISABLE_DOWNLOAD` does not cover these; they need their own
controls.

| What | Tools | Control |
|---|---|---|
| PowerShell Gallery modules: `Microsoft.Graph`, `ExchangeOnlineManagement`, `PSWindowsUpdate`, `Microsoft.WinGet.Client` | ALMANAC, ASTERISM, CARILLON, CUMULUS, ECLIPSE, HALO, ORBIT, ORRERY, RAVEN, ZENITH, FORGE, WHETSTONE, CONJURE | Pre-install approved versions on technician machines, or from an internal repository, so the tools find them instead of installing the latest from the Gallery |
| CONJURE `DirectDownloads` | CONJURE | Fill in `Sha256` for **every** entry in `config.json`. CONJURE refuses to run a payload whose hash does not match. An entry without a hash runs whatever the URL serves |
| winget packages and the App Installer fallback | CONJURE | winget verifies package hashes against its manifests, and the App Installer bundle is Microsoft-signed. Restrict which packages are approved for deployment in your own procedure |

## 3. Logical access (CC6.1, CC6.3)

- Most tools need Administrator and relaunch elevated. CHALICE deliberately runs
  as the signed-in user. Elevation belongs to the technician's admin account, so
  your existing privileged-access controls apply. The toolkit adds no accounts,
  no service and no listener.
- Restrict who can **write** to the deployed copy. Anyone who can edit the
  scripts can change what runs elevated on client machines.
- `config.json` sits beside the scripts and holds settings that matter for
  control, including `LogDirectory` and CONJURE's download list
  and hashes. Protect it with the same ACLs as the scripts.

## 4. Confidentiality of output (C1.1, C1.2)

Reports and CSVs are written to the configured `LogDirectory`, or to a per-tool
fallback path when it is unset. Several contain data a client would treat as
confidential:

| Tool | Sensitive content |
|---|---|
| WISP | Wi-Fi keys in cleartext, **only** with `-IncludeKey` (masked by default) |
| WYRM | BitLocker recovery information |
| ARGUS, WARD, SPHINX, ECLIPSE, ALMANAC | Account rosters, privilege, MFA and licensing state |
| MINOTAUR | Share and NTFS permissions |
| RAVEN, HALO | Mailbox rules and forwarding, Conditional Access gaps |
| Transcripts (`-Transcript`) | Everything the tool printed to the console |

Controls to put around them:

- Set `LogDirectory` (HEARTH writes it) to a protected location, not a desktop or
  a shared temp folder.
- Define retention: move reports into the ticket or the client's evidence store,
  then delete the local copy. Clear the directory between client sites.
- Forbid `-IncludeKey` on WISP unless a ticket calls for it, and record why.
- Send reports over your approved channels (PSA attachment, encrypted share), not
  ad-hoc email.

**Where the data goes.** Nowhere but the log directory. No tool transmits what it
collects: there is no telemetry, webhook, mail or upload, and the Pester suite fails
on any code that would add one. Every report, CSV, transcript, notes file and error
log lands in `LogDirectory` (or beside the script), so clearing that directory
removes everything the toolkit wrote. Keep `LogDirectory` on the local disk; a
network path would put client data on that share.

**Cloud tools only read.** The tools in The Firmament (ZENITH, ALMANAC, ORBIT,
ECLIPSE, ASTERISM, CUMULUS, ORRERY, RAVEN, HALO, CARILLON) request Microsoft Graph
`Read` permissions only and call no command that changes a tenant; the Pester suite
enforces both. Exchange, Teams, Azure and SharePoint sign-ins carry whatever the
signed-in account's role allows, so run these tools under a read-only role — Global
Reader, Azure Reader, Exchange View-Only Organization Management — and the session
itself cannot write, whatever the script does.

## 5. Secrets (CC6.1)

- COVENANT takes the local administrator password as a `SecureString`. Supply it
  interactively or from your vault at run time, never in a saved command line or
  script.

## 6. Monitoring and evidence (CC7.2, CC7.3, CC8.1)

The toolkit produces evidence your change and incident records can cite:

| Evidence | Source |
|---|---|
| What a run did, step by step | `-Transcript` on tools that write logs; RITUAL's rollup records each step's status, duration and artifacts |
| What a change *would* do, before approval | `-WhatIf` on every state-changing tool (REVENANT, EMBALM, COVENANT, BASILISK, CLEANSE, WYRM, FORGE, WHETSTONE, RUNEPRESS, CONJURE, CONDUIT, SOLDER, LODESTAR, CHALICE) |
| Technician actions and decisions | Every tool that records notes (the repair tools CONDUIT, SOLDER, LODESTAR and CHALICE, and the audit verdicts of the rest) saves `<TOOL>_Notes_<timestamp>.html` at the end of the run: timestamped, categorised, with a plain-text block for the ticket |
| Tool errors | `Write-TKError` appends to `TK_Errors_<yyyyMM>.jsonl` in `LogDirectory` (host, user, tool, message) and can alert a Teams channel |
| Point-in-time state | Every audit tool's HTML report; CODEX indexes a directory of them |

A procedure that holds up: run the `-WhatIf` preview and attach it to the change
ticket, run the change with `-Transcript`, and attach the transcript, the report and the
notes file.

Set `LogDirectory`. If it is unset, the error log falls back to the folder the
tool lives in, which is the deployed copy technicians should not be writing to.

## Checklist

- [ ] Toolkit deployed from a verified, tagged release to a read-only location
- [ ] `TK_DISABLE_DOWNLOAD=1` set machine-wide on every technician machine
- [ ] Upgrades go through change approval with the CHANGELOG reviewed
- [ ] Scripts signed with your certificate and `AllSigned` enforced (if policy needs it)
- [ ] Gallery modules pre-installed at approved versions
- [ ] Every CONJURE `DirectDownloads` entry has a `Sha256`
- [ ] `config.json` ACL-protected
- [ ] Cloud tools run under a read-only admin role
- [ ] `LogDirectory` set to a protected location, with a retention rule
- [ ] WISP `-IncludeKey` use restricted and recorded
- [ ] `-WhatIf`, transcript and note export attached to change tickets
