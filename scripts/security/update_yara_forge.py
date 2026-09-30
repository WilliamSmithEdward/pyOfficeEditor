"""Pin the newest YARA Forge release that is at least a week old.

The Malware scan workflow's YARA-X job runs the rule set named in
``.github/security/yara-forge.json``: a dated release, the package in it
(``full``, the widest of YARA Forge's three), the archive, and the
archive's SHA-256.  The weekly Update YARA rules workflow moves that pin in two
steps, and fetches everything itself, with ``gh api`` and ``curl``, so
this script never touches the network:

``choose RELEASES_JSON``
    Reads YARA Forge's release list, as ``gh api`` returns it, and prints
    the download URL, the release and the SHA-256 GitHub records for the
    archive of the newest stable release published at least ``COOLDOWN``
    ago, a week, long enough for most broken or tampered releases to be
    caught and pulled first.  It prints nothing when that release is the
    one already pinned.

``pin RELEASE SHA256 ARCHIVE``
    Checks the downloaded archive against that SHA-256 and for the
    package's rule file, and only then rewrites the pin.

The workflow then opens a pull request and starts the Malware scan workflow
on its branch, so the new rules scan the repository before anyone merges
them.  The exit status is 1 on anything that cannot be verified.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

PIN = Path(__file__).parents[2] / ".github" / "security" / "yara-forge.json"
DOWNLOAD = "https://github.com/YARAHQ/yara-forge/releases/download/{release}/{asset}"
COOLDOWN = dt.timedelta(days=7)


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


def candidate(releases: list[dict[str, Any]], pin: dict[str, str], now: dt.datetime) -> tuple[str, str] | None:
    """The release and digest to move the pin to, or None to stay."""
    release, sha256 = choose(releases, pin["asset"], now)
    if release == pin["release"]:
        if sha256 != pin["sha256"]:
            raise ValueError(f"the digest of pinned YARA Forge {release} changed")
        return None
    if release < pin["release"]:
        raise ValueError(f"the newest week-old release, {release}, predates the pinned {pin['release']}")
    return release, sha256


def verify(archive: Path, sha256: str, package: str) -> None:
    """The archive matches its digest and holds the package's rule file."""
    if hashlib.sha256(archive.read_bytes()).hexdigest() != sha256:
        raise ValueError("the downloaded archive does not match its SHA-256")
    with zipfile.ZipFile(archive) as zipped:
        wanted = f"packages/{package}/yara-rules-{package}.yar"
        if not any(name.endswith(wanted) for name in zipped.namelist()):
            raise ValueError(f"the archive holds no {wanted}")


def main(argv: list[str], now: dt.datetime | None = None) -> int:
    pin: dict[str, str] = json.loads(PIN.read_text(encoding="utf-8"))
    if len(argv) == 2 and argv[0] == "choose":
        releases: list[dict[str, Any]] = json.loads(Path(argv[1]).read_text(encoding="utf-8"))
        moved = candidate(releases, pin, now or dt.datetime.now(dt.timezone.utc))
        if moved is None:
            print(f"YARA Forge {pin['release']} is already pinned", file=sys.stderr)
        else:
            release, sha256 = moved
            print(DOWNLOAD.format(release=release, asset=pin["asset"]), release, sha256)
        return 0
    if len(argv) == 4 and argv[0] == "pin":
        release, sha256, archive = argv[1], argv[2], Path(argv[3])
        if not re.fullmatch(r"\d{8}", release) or not re.fullmatch(r"[0-9a-f]{64}", sha256):
            raise ValueError("pin takes a dated release and a SHA-256")
        verify(archive, sha256, pin["package"])
        PIN.write_text(json.dumps({**pin, "release": release, "sha256": sha256}, indent=2) + "\n", encoding="utf-8")
        print(f"Pinned YARA Forge {release}, SHA-256 {sha256}")
        return 0
    raise ValueError("usage: update_yara_forge.py choose RELEASES_JSON | pin RELEASE SHA256 ARCHIVE")


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except ValueError as error:
        print(f"error: {error}", file=sys.stderr)
        sys.exit(1)
