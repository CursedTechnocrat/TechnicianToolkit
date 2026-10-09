# Releasing the Technician Toolkit

The build is automated. The release is not, and cannot be — see *Why signing is
manual* below. This checklist exists so the manual half is not carried in one
person's head.

**Status for 6.0:** 6.0.0 is the first release to be **signed**. 5.0.0 shipped
unsigned while the certificate was in validation, and the signed 5.0.1 it promised
never went out (nor did 5.1.0, which is folded into 6.0.0). So 6.0.0 is tagged
first and **published only once it is signed**. Steps 4 and 5 run for the first
time, and the release waits for them rather than going out unsigned.

**If signing falls through** and 6.0.0 has to ship unsigned, the signed build
becomes **6.0.1**. It is a patch release because it changes no code. It has to be
a new version because the signed files have different hashes, and winget cannot
carry two hashes for one version. Only the three `.csproj` files move, to
`6.0.1`; the script headers and the GRIMOIRE registry stay at `6.0`. Then tag
`v6.0.1` and run this checklist again. Follow the *While unsigned* line in step 7
for the unsigned release.

**From then on, every release is signed.** A release that cannot be signed waits
for it, rather than shipping unsigned with a promise.

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

- [ ] Both artifacts downloaded and unzipped.
- [ ] The hash of each `.exe` matches its `SHA256SUMS.txt`:

```powershell
Get-FileHash .\TechnicianToolkit.exe -Algorithm SHA256
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

- [ ] Start **SimplySign Desktop** and authenticate with the SimplySign mobile
      app. The certificate appears in the Windows certificate store through a
      virtual smart card.
- [ ] **First time only:** confirm that exactly one code-signing certificate
      matches the subject `/n` selects. If two match, for example an expired one
      left behind, `signtool` stops and asks. Pin that certificate with
      `/sha1 <thumbprint>` instead of `/n`.

```powershell
Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object Subject -like '*Open Source Developer*' |
    Format-List Subject, Thumbprint, NotAfter
```

- [ ] Sign each binary **with timestamping**. Timestamping is not optional: the
      certificate is short-lived, and a timestamped signature stays valid after
      it expires where an un-timestamped one dies with it.

```powershell
signtool sign /n "Open Source Developer" /fd SHA256 `
    /tr http://time.certum.pl /td SHA256 TechnicianToolkit.exe
```

## 5. Verify the signature

- [ ] `signtool verify /pa /v TechnicianToolkit.exe` passes on **both** binaries.
      `signtool` signs and verifies any PE from the same x64 session, so ARM64
      needs no separate arrangement.
- [ ] The verify output shows a **timestamp**. A signature without one passes
      today and dies when the certificate expires.
- [ ] Launch the signed x64 build once. The UAC prompt should name the publisher
      (`Open Source Developer, John Joseph Bejarana`), not *Unknown publisher*.

## 6. Re-hash what you are actually shipping

Signing changes the file, so the hashes from step 1 no longer apply.

- [ ] Recompute the SHA-256 of each binary in its final form.
- [ ] Those are the hashes that go in the release notes and, later, the winget
      manifest. A hash taken from the unsigned artifact will not match a signed
      download.

## 7. Publish the release

- [ ] Create the GitHub release against the tag.
- [ ] Attach both binaries, named so the architecture is unambiguous:
      `TechnicianToolkit-6.0.0-win-x64.exe`,
      `TechnicianToolkit-6.0.0-win-arm64.exe`.
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

## 8. Generate the winget manifest — last

Only after the release is published and the files are final, because the
manifest carries their hashes.

```powershell
wingetcreate update CursedTechnocrat.TechnicianToolkit --version 6.0.0 `
    --urls <x64-url> <arm64-url> --submit
```

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

It is deliberate, and it is the name the certificate subject will carry. Do not
tidy it back to the pseudonym — the public record and the certificate have to
agree, and the handle is kept beside the name precisely so the GitHub account
and the person resolve to each other.
