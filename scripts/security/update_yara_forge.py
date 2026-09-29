"""Pin the newest YARA Forge release that is at least a week old.

The Security workflow's YARA-X scan runs the rule set named in
``.github/security/yara-forge.json``: a dated release, the package in it
(``full``, the widest of YARA Forge's three), the archive, and the
archive's SHA-256.  This rewrites that pin to the newest stable release
published at least ``COOLDOWN`` ago, a week, long enough for most broken
or tampered releases to be caught and pulled first.  The digest comes
from the one GitHub records for the release asset, and the download has
to match it before anything is written.

The weekly YARA Forge workflow runs this, opens a pull request when the
pin moves, and starts the Security workflow on that pull request's
branch, so the new rules scan the repository before anyone merges them.

Usage::

    python scripts/security/update_yara_forge.py

The exit status is 0 whether or not the pin moved, and 1 on anything
that cannot be verified.  ``GITHUB_TOKEN``, when set, authenticates the
API request.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import os
import re
import sys
import urllib.request
import zipfile
from pathlib import Path
from typing import Any

PIN = Path(__file__).parents[2] / ".github" / "security" / "yara-forge.json"
RELEASES = "https://api.github.com/repos/YARAHQ/yara-forge/releases?per_page=30"
DOWNLOAD = "https://github.com/YARAHQ/yara-forge/releases/download/{release}/{asset}"
COOLDOWN = dt.timedelta(days=7)


def fetch(url: str) -> bytes:
    headers = {"Accept": "application/vnd.github+json", "User-Agent": "pyOfficeEditor-security"}
    token = os.environ.get("GITHUB_TOKEN")
    if token and url.startswith("https://api.github.com/"):
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(request, timeout=120) as response:
        return response.read()


def choose(releases: list[dict[str, Any]], asset: str, now: dt.datetime) -> tuple[str, str]:
    """The newest dated, stable release a week old or more, and the SHA-256
    GitHub records for its `asset`."""
    for release in sorted(releases, key=lambda r: r["tag_name"], reverse=True):
        tag = release["tag_name"]
        if not re.fullmatch(r"\d{8}", tag) or release["draft"] or release["prerelease"]:
            continue
        published = dt.datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
        if now - published < COOLDOWN:
            continue
        matching = [a for a in release["assets"] if a["name"] == asset]
        if len(matching) != 1:
            raise ValueError(f"YARA Forge {tag} has no single {asset}")
        if matching[0]["browser_download_url"] != DOWNLOAD.format(release=tag, asset=asset):
            raise ValueError(f"YARA Forge {tag}'s {asset} is not served from its release")
        digest = matching[0].get("digest") or ""
        if not re.fullmatch(r"sha256:[0-9a-f]{64}", digest):
            raise ValueError(f"YARA Forge {tag}'s {asset} has no SHA-256 digest")
        return tag, digest.removeprefix("sha256:")
    raise ValueError("no YARA Forge release is a week old")


def verify(archive: bytes, sha256: str, package: str) -> None:
    """The archive matches its digest and holds the package's rule file."""
    if hashlib.sha256(archive).hexdigest() != sha256:
        raise ValueError("the downloaded archive does not match its SHA-256")
    with zipfile.ZipFile(io.BytesIO(archive)) as zipped:
        wanted = f"packages/{package}/yara-rules-{package}.yar"
        if not any(name.endswith(wanted) for name in zipped.namelist()):
            raise ValueError(f"the archive holds no {wanted}")


def main(now: dt.datetime | None = None) -> int:
    pin: dict[str, str] = json.loads(PIN.read_text(encoding="utf-8"))
    releases: list[dict[str, Any]] = json.loads(fetch(RELEASES))
    release, sha256 = choose(releases, pin["asset"], now or dt.datetime.now(dt.timezone.utc))
    if release == pin["release"]:
        if sha256 != pin["sha256"]:
            raise ValueError(f"the digest of pinned YARA Forge {release} changed")
        print(f"YARA Forge {release} is already pinned")
        return 0
    if release < pin["release"]:
        raise ValueError(f"the newest week-old release, {release}, predates the pinned {pin['release']}")
    verify(fetch(DOWNLOAD.format(release=release, asset=pin["asset"])), sha256, pin["package"])
    PIN.write_text(json.dumps({**pin, "release": release, "sha256": sha256}, indent=2) + "\n", encoding="utf-8")
    print(f"Pinned YARA Forge {release}, SHA-256 {sha256}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ValueError as error:
        print(f"error: {error}", file=sys.stderr)
        sys.exit(1)
