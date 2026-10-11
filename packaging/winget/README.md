# winget packaging

The manifest for `CursedTechnocrat.TechnicianToolkit`, kept in the repository so
it is reviewable and versioned. The copy that actually ships lives in
[microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs); these three
files are the source they are generated from.

A submission is always three files, all carrying the same `PackageIdentifier`
and `PackageVersion`. They live in `manifests/`:

| File | Holds |
|---|---|
| `CursedTechnocrat.TechnicianToolkit.yaml` | The version manifest — the tie between the other two |
| `…installer.yaml` | Architectures, URLs, hashes, installer type |
| `…locale.en-US.yaml` | Everything `winget show` displays |

**Nothing else may go in `manifests/`.** `winget validate` parses *every* file in
the directory it is given, as YAML — this README lived there at first and failed
validation with `[YAML:Scanner] mapping values are not allowed in this context`,
pointing at a line of prose that happened to end in a colon. The error names a
line number and no file, so it is a genuinely confusing few minutes. Keep the
directory to the three manifests.

## First submission vs. later releases

`wingetcreate update` only works on a package that already exists in
winget-pkgs. Check first:

```powershell
winget show CursedTechnocrat.TechnicianToolkit
```

**Not found — first submission.** Update these three files in a PR (version,
both URLs and hashes, `ReleaseDate`, `ReleaseNotesUrl`), merge it, then submit
them as they are:

```powershell
winget validate --manifest packaging\winget\manifests
wingetcreate submit packaging\winget\manifests
```

The hashes then come from step 6 of [`RELEASING.md`](../../RELEASING.md), and
each must equal GitHub's own digest of the release asset (the release API's
`digest` field) — check that before submitting.

**Found — every later release.** Let `wingetcreate` do it from the published
URLs, then copy what it generated back into this folder:

```powershell
wingetcreate update CursedTechnocrat.TechnicianToolkit --version <version> `
    --urls <x64-url> <arm64-url> `
    --release-notes-url https://github.com/CursedTechnocrat/TechnicianToolkit/releases/tag/v<version> `
    --submit
```

`wingetcreate` downloads each URL and computes the hash itself, which removes
the most common way a submission fails: a hash taken from the wrong copy of the
file.

`wingetcreate` is not part of winget itself: `winget install Microsoft.WingetCreate`.
The first `submit` or `update --submit` asks you to sign in to GitHub in the
browser, because it opens the pull request on microsoft/winget-pkgs as you.

**The hash must come from the file after signing.** Signing appends to the PE, so
a hash computed from the unsigned CI artifact will not match what a user
downloads. This is the ordering constraint behind "generate the manifest last".

## Why `portable`

`InstallerType: portable` says the payload is a single executable that installs
nothing. winget places it in its links directory, adds the
`PortableCommandAlias`, and uninstalls by deleting it — which is exactly the
shape of the application.

Setting this to `exe` instead would tell winget the file is a setup program and
run it with installer switches it does not have. The application would launch
its window in the middle of `winget install`, and the install would appear to
hang. This field is the easiest one to get wrong and the most annoying to
diagnose.

## Validating locally

```powershell
winget validate --manifest packaging/winget/manifests
```

That checks schema and required fields. It does not check that the URLs resolve
or that the hashes are right — only a real submission does.

## Before submitting

- The GitHub release exists and both binaries are attached.
- `ReleaseDate` is the real release date, not the placeholder.
- Hashes were computed from the attached files, after signing.
- `winget validate` passes.
- Ideally, the package installs on a clean machine from the submitted manifest.

## Signing

The binaries are signed from 6.0.0 on. winget does not require a signature for
a portable package, but a new certificate builds SmartScreen reputation over
time, so the first installs can still see a warning. Do not treat that as a
packaging bug.
