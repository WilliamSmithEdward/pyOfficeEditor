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

The package has no runtime dependencies. CodeQL and Semgrep scan it,
with the workflows that build and publish it, on every push to main,
every pull request, every week, and before every release. A finding
fails the scan unless
[.github/security/accepted.toml](.github/security/accepted.toml) lists it
with the reason it is accepted, and an entry there that no longer
matches fails it too.

A release is published only after its commit passes the scan, and it
carries the report as `pyofficeeditor-<version>-security-report.md`,
beside the SARIF the report was made from. Every action the workflows
use is pinned to a commit, and Dependabot proposes updates to the
actions and to the pinned Semgrep.
