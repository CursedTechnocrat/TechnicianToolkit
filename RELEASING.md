# Releasing the Technician Toolkit

The build is automated. The release is not, and cannot be — see *Why signing is
manual* below. This checklist exists so the manual half is not carried in one
person's head.

**Every release is signed.** 6.0.0 was the first; 5.0.0 shipped unsigned while
the certificate was in validation. A release is tagged first and **published only
once it is signed**, and a release that cannot be signed waits for it rather than
shipping unsigned with a promise.

**If signing falls through** and a release has to ship unsigned anyway, the signed
build becomes the next patch version. It changes no code, but it has to be a new
version because the signed files have different hashes, and winget cannot carry
two hashes for one version. Only the three `.csproj` files move; the script
headers and the GRIMOIRE registry stay where they are. Then tag the patch and run
this checklist again. Follow the *While unsigned* line in step 7 for the unsigned
release.

---

## Before you tag

- [ ] `main` is green in CI — PSScriptAnalyzer, Pester, and the Desktop app job.
- [ ] `CHANGELOG.md` has a dated section for the release (`[6.0.0]`), not `[Unreleased]`.
- [ ] Every script header and every GRIMOIRE registry row reads the release
      version. The `Version consistency` Pester gate enforces this, so a
      mismatch fails CI rather than reaching here.
- [ ] The three `.csproj` files carry the matching three-part version
      (`6.0.0`).

Then tag and push:

```powershell
git tag -a v6.0.0 -m 'Technician Toolkit 6.0.0'
git push origin v6.0.0
```

That fires `release-app.yml`, which builds `win-x64` and `win-arm64` and uploads
each as an **unsigned** workflow artifact. It does not create the release.

---

## 1. Collect the artifacts

Download both `TechnicianToolkit-<rid>-unsigned` artifacts from the workflow run.
Each contains `TechnicianToolkit.exe` and a `SHA256SUMS.txt` recording what CI
built.

- [ ] Both artifacts downloaded and unzipped, **each into its own folder** —
      both hold a file named `TechnicianToolkit.exe`. The `.exe` sits two levels
      down, at `publish\<rid>\TechnicianToolkit.exe`.
- [ ] The hash of each `.exe` matches its `SHA256SUMS.txt`:

```powershell
Expand-Archive .\TechnicianToolkit-win-x64-unsigned.zip   -DestinationPath .\tk-x64
Expand-Archive .\TechnicianToolkit-win-arm64-unsigned.zip -DestinationPath .\tk-arm64
Get-ChildItem .\tk-x64, .\tk-arm64 -Recurse -Filter TechnicianToolkit.exe |
    Get-FileHash -Algorithm SHA256 | Format-Table Hash, Path -AutoSize
Get-Content .\tk-x64\SHA256SUMS.txt, .\tk-arm64\SHA256SUMS.txt
```

## 2. Check the ARM64 build

- [ ] If an ARM device is available, run the ARM64 build on it and note the
      result in the release notes.
- [ ] If not — which is the current situation — the release notes must say the
      ARM64 binary is **built but untested on hardware**, and ask for a report.
      See the ARM64 row in [`docs/desktop-port.md`](docs/desktop-port.md) and the
      request in [`CONTRIBUTING.md`](CONTRIBUTING.md). Do not quietly ship it as
      though it were verified.

## 3. Smoke-test the x64 build

On a machine that has never had PowerShell 7 installed, if one can be found —
that is the claim the whole port rests on.

- [ ] The application launches and shows the window.
- [ ] A read-only tool (WARD or SCRYER) runs to completion and writes its report.
- [ ] The report opens in the browser.
- [ ] Settings persist across a relaunch.

## 4. Sign

On a Windows machine with the Windows SDK installed, which provides `signtool`.
Run everything in **PowerShell**, not the Developer Command Prompt, which is
`cmd` and rejects the pipelines below. `signtool` is usually not on PATH; it lives
under the SDK's versioned `bin` folder:

```powershell
$signtool = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe' |
    Sort-Object FullName | Select-Object -Last 1 -ExpandProperty FullName
& $signtool /?
```

- [ ] Start **SimplySign Desktop** and authenticate with the SimplySign mobile
      app. The certificate appears in the Windows certificate store through a
      virtual smart card.
- [ ] Confirm that exactly one code-signing certificate matches. Sign by
      thumbprint (`/sha1`) rather than subject (`/n`): if two certificates ever
      match a subject, for example an expired one left behind, `signtool` stops
      and asks. The current certificate is
      `CN=Open Source Developer John Bejarana`, thumbprint
      `F14F545CB084EDCB18630CA0B520ABC3C46B2516`, valid to 2027-10-11. When it
      is renewed, the thumbprint changes; update it here.

```powershell
Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object Subject -like '*Open Source Developer*' |
    Format-List Subject, Thumbprint, NotAfter
```

- [ ] Sign each binary **with timestamping**. Timestamping is not optional: the
      certificate is short-lived, and a timestamped signature stays valid after
      it expires where an un-timestamped one dies with it.

```powershell
$x64   = '.\tk-x64\publish\win-x64\TechnicianToolkit.exe'
$arm64 = '.\tk-arm64\publish\win-arm64\TechnicianToolkit.exe'
& $signtool sign /sha1 F14F545CB084EDCB18630CA0B520ABC3C46B2516 /fd SHA256 `
    /tr http://time.certum.pl /td SHA256 $x64 $arm64
```

## 5. Verify the signature

- [ ] `signtool verify /pa /v TechnicianToolkit.exe` passes on **both** binaries.
      `signtool` signs and verifies any PE from the same x64 session, so ARM64
      needs no separate arrangement.
- [ ] The verify output shows a **timestamp**. A signature without one passes
      today and dies when the certificate expires.
- [ ] The *Hash of file* line in the verify output is the Authenticode digest,
      which excludes the signature. It is **not** the file's SHA-256; do not put
      it in the release notes (step 6 computes the right one).
- [ ] Check what a user will see. The UAC prompt should name the publisher
      **Open Source Developer John Bejarana**, not *Unknown publisher*. Launch
      it from **Explorer**: a run from an elevated PowerShell window inherits its
      elevation and shows no prompt at all. Without launching, the file's
      **Properties → Digital Signatures** tab shows the same signer name.

## 5b. Sign the scripts

Same machine, same SimplySign session. The release workflow uploads a third
artifact, `TechnicianToolkit-scripts-unsigned`: every top-level script, the
licence, `config.json`, and the `shared` and `build` folders the signing script
needs. Only these release copies are signed — never the repository's, where
`build/Sync-TKShared.ps1` rewrites files and would break a signature.

Signed scripts are what GRIMOIRE and RITUAL accept: they fetch a missing tool
from `releases/latest/download/<tool>.ps1` and discard it unrun unless its
signature is valid and from **Open Source Developer John Bejarana**. A release
without signed scripts attached leaves those downloads failing (safely), and
leaves the README's quick-launch commands with nothing to fetch.

- [ ] Download and unzip `TechnicianToolkit-scripts-unsigned`.
- [ ] **Unblock the unzipped files.** A zip downloaded through a browser carries
      the *Mark of the Web*, and every file extracted from it inherits it. Under
      the usual `RemoteSigned` execution policy Windows then refuses to run the
      still-unsigned `Sign-TKScripts.ps1` (*"is not digitally signed... cannot be
      loaded"*). These are files you just downloaded from your own CI run, so
      clearing the mark is safe. Do it before signing — the signed copies are
      new files and do not need it.
- [ ] Sign, timestamp and verify every script in one go:

```powershell
Expand-Archive .\TechnicianToolkit-scripts-unsigned.zip -DestinationPath .\scripts-unsigned
Get-ChildItem .\scripts-unsigned -Recurse -File | Unblock-File
.\scripts-unsigned\build\Sign-TKScripts.ps1 -Source .\scripts-unsigned `
    -Thumbprint F14F545CB084EDCB18630CA0B520ABC3C46B2516 -SignTool $signtool
```

- [ ] It ends with *Every script is signed, timestamped and verifies*. It stops,
      and nothing may be published, if any script fails verification or lacks a
      timestamp.
- [ ] It writes the signed scripts to `scripts-unsigned\dist\scripts\` (with
      `LICENSE`, `config.json` and `SHA256SUMS-scripts.txt`) and
      `scripts-unsigned\dist\TechnicianToolkit-scripts-<version>.zip`.

## 6. Re-hash what you are actually shipping

Signing changes the file, so the hashes from step 1 no longer apply.

- [ ] Copy each binary to its release name (renaming does not touch the
      signature) and recompute the SHA-256 of that final file:

```powershell
Copy-Item $x64   .\TechnicianToolkit-6.0.0-win-x64.exe
Copy-Item $arm64 .\TechnicianToolkit-6.0.0-win-arm64.exe
Get-FileHash .\TechnicianToolkit-6.0.0-win-*.exe -Algorithm SHA256 | Format-Table Hash, Path -AutoSize
```

- [ ] Both hashes differ from step 1's. One that does not was never signed.
- [ ] Those are the hashes that go in the release notes and, later, the winget
      manifest. A hash taken from the unsigned artifact will not match a signed
      download.

## 7. Publish the release

- [ ] Create the GitHub release against the tag. Keep it a **draft** until every
      item below is done: a published release with no binaries, or with
      placeholder hashes, is what users see in the meantime.
- [ ] Attach both binaries, named so the architecture is unambiguous:
      `TechnicianToolkit-6.0.0-win-x64.exe`,
      `TechnicianToolkit-6.0.0-win-arm64.exe`.
- [ ] Attach **every signed `.ps1` individually**, under its own name — GRIMOIRE,
      RITUAL and the README fetch them by name from the latest release — plus
      `TechnicianToolkit-scripts-<version>.zip` and `SHA256SUMS-scripts.txt`.
      Dragging the whole `dist\scripts` folder's `.ps1` files into the release
      page works; so does `gh release upload v6.0.0 .\scripts-unsigned\dist\scripts\*.ps1`.
- [ ] Paste the `CHANGELOG.md` section for this version.
- [ ] Include the SHA-256 of both attached files.
- [ ] **When signed**, say so, and say what the publisher name will look like
      (see *The publisher name will look like a person* below). A new signature
      builds SmartScreen reputation over time, so the first downloads can still
      see a warning. Say that too rather than promising none.
- [ ] **When signed for the first time**, replace the README's "binaries are not
      code-signed" callout with the publisher-name explanation.
- [ ] **While unsigned** (only in the fallback above), state it plainly and set
      the expectation:

  > These binaries are not code-signed. SmartScreen will warn on first run, and
  > some antivirus products may flag a single-file executable that unpacks
  > scripts and runs them elevated. Verify the SHA-256 above against your
  > download. A signed build will follow as the next patch release.

- [ ] Note that the ARM64 build is untested on hardware, and ask for reports.
- [ ] After publishing, reload the release page and check it: both binaries are
      listed under *Assets* (an upload still in progress when you save is
      dropped), and the notes run to the end rather than stopping part-way.

## 8. Generate the winget manifest — last

Only after the release is published and the files are final, because the
manifest carries their hashes.

- [ ] `winget install Microsoft.WingetCreate` if `wingetcreate` is missing.
- [ ] `winget show CursedTechnocrat.TechnicianToolkit` finds the package? Then:

```powershell
wingetcreate update CursedTechnocrat.TechnicianToolkit --version 6.0.0 `
    --urls <x64-url> <arm64-url> `
    --release-notes-url https://github.com/CursedTechnocrat/TechnicianToolkit/releases/tag/v6.0.0 `
    --submit
```

  and copy the generated manifests back into `packaging/winget/manifests/`.
  If it reports *No package found*, `update` cannot run; the package's first
  submission is `wingetcreate submit packaging\winget\manifests`, after a PR
  that brings those files to the release (see the winget README).

- [ ] Manifest validates (`winget validate --manifest packaging/winget/manifests`).
- [ ] Installs cleanly from the submitted manifest on a clean machine.

See [`packaging/winget/README.md`](packaging/winget/README.md) for the manifest
layout and the fields that are easy to get wrong.

---

## Why signing is manual

The certificate is Certum's Open Source Code Signing, with the key held in
**SimplySign**, their cloud HSM. The key is non-exportable by design — that is
the point of the CA/B hardware requirement — so it cannot be handed to a signing
service, and reaching it needs the SimplySign Desktop client to mount a virtual
smart card against a session authenticated with a phone.

GitHub-hosted runners cannot hold such a session. This is not a missing secret
that could be added later; there is no credential that would work. Certum has
said CI/CD support is planned, with no date attached — revisit then.

Two alternatives were considered and rejected for 5.0:

| Approach | Verdict |
|---|---|
| Self-hosted runner holding a SimplySign session open | An always-on Windows box, with sessions that expire, to save one manual step on a release that happens a few times a year |
| Script the OTP by extracting the TOTP secret | Fragile, and it defeats the second factor the certificate's assurance rests on |

## The publisher name will look like a person

The certificate subject is a natural person prefixed `Open Source Developer`, so
the UAC prompt names a developer rather than an organization. That is normal for
this certificate type, and it reads as suspicious only when left unexplained —
so explain it in the README rather than treating it as a problem to hide.

## Validation asked for the name in public

Certum's validation asks for a GitHub repository, website or similar where the
applicant's full name is publicly readable, so it can be tied to the identity
documents on file. That is why the copyright notice in every source file, the
four `.csproj` files, the winget locale manifest and the README all name
**John Joseph Bejarana (CursedTechnocrat)** rather than the handle alone.

It is deliberate. The certificate subject itself carries the shorter
**John Bejarana** (`CN=Open Source Developer John Bejarana`), so that is the name
the UAC prompt shows and the README and release notes quote. Do not tidy the
public record back to the pseudonym: it is what tied the certificate to the
person, and the handle is kept beside the name precisely so the GitHub account
and the person resolve to each other.
