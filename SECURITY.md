# Security policy

## Reporting a vulnerability

Please report a vulnerability privately, not in a public issue. Use
[Report a vulnerability](https://github.com/WilliamSmithEdward/pyOfficeEditor/security/advisories/new)
on the repository's Security tab. It opens a draft advisory that only
you and the maintainer can see.

A useful report names the pyOfficeEditor version, the Python version
and the operating system, and includes a file or a short script that
shows the problem. A crafted file is the most direct proof, so attach a
minimal one if you can.

## Supported versions

Only the latest release on PyPI is supported. A fix ships in a new
release, and earlier versions do not get one.

## What to report

pyOfficeEditor reads and writes Office files that may come from anyone,
so the input to worry about is a file. For example:

- a file that makes the library hang, run out of memory or crash the
  interpreter;
- a file that gets past the XML parser's limits: a `<!DOCTYPE`, an
  entity other than the five predefined ones, or nesting deeper than 256
  elements, any of which should be refused;
- a file whose reading writes anything, or a save that touches a file
  other than the one it was asked to write.

The formula engine calculates a workbook's formulas in Python. It never
starts a process, opens a network connection or reads another file. A
formula that reads another workbook reads the values the file itself
caches for that link, as Excel does while the linked workbook is closed.

## How the code is checked

The package has no runtime dependencies. The Security and Malware scan
workflows run on every push to main, every pull request, every day, and
before every release:

- Security: CodeQL and Semgrep scan the package, the workflows that
  build and publish it, and the scripts that check the scans.
- Malware scan: ClamAV and YARA-X scan every tracked file, the Excel
  test fixtures among them, and the built wheel and sdist. ClamAV
  fetches the current official signatures on every run and reports any
  file that holds a VBA project. YARA-X runs YARA Forge's full rule set,
  which gathers the public YARA rule collections into one.

A finding fails the scan unless
[.github/security/accepted.toml](.github/security/accepted.toml), or
[.github/security/malware-accepted.toml](.github/security/malware-accepted.toml)
for a malware match, lists it with the reason it is accepted, and an
entry there that no longer matches fails it too. So does a warning a
scanner raises about its own run.

A release is published only after its commit passes both scans, and it
carries the reports as `pyofficeeditor-<version>-security-report.md` and
`pyofficeeditor-<version>-malware-report.md`, beside the SARIF they were
made from.

Everything the workflows run is pinned. Each action is pinned to a
commit, the Semgrep and ClamAV images to a digest, and the YARA-X engine
and the YARA Forge rules to a release and its SHA-256. The build tools
are hash-locked, and each runner is a named OS release. Dependabot
proposes updates to the actions, the images and the locked tools, each a
week after its release. A weekly workflow moves the YARA pins in
.github/security/yara.json: YARA Forge's newest release at once, since the
pull request it opens is scanned before it is merged, and a YARA-X release
once it is a week old.
