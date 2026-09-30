"""scripts/security/update_yara_forge.py, which chooses the YARA Forge
release the malware scan runs and checks its archive before pinning it."""

from __future__ import annotations

import datetime as dt
import hashlib
import importlib.util
import sys
import zipfile
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

SCRIPT = Path(__file__).parents[1] / "scripts" / "security" / "update_yara_forge.py"
NOW = dt.datetime(2026, 10, 7, tzinfo=dt.timezone.utc)
ASSET = "yara-forge-rules-full.zip"
PIN = {"release": "20260920", "package": "full", "asset": ASSET, "sha256": "a" * 64}


def load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("update_yara_forge", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def release(tag: str, days_old: int, digest: str = "b" * 64, **extra: Any) -> dict[str, Any]:
    url = f"https://github.com/YARAHQ/yara-forge/releases/download/{tag}/{ASSET}"
    return {
        "tag_name": tag,
        "draft": False,
        "prerelease": False,
        "published_at": (NOW - dt.timedelta(days=days_old)).isoformat().replace("+00:00", "Z"),
        "assets": [{"name": ASSET, "browser_download_url": url, "digest": f"sha256:{digest}"}],
        **extra,
    }


def test_the_newest_release_a_week_old_is_chosen() -> None:
    releases = [release("20260920", 17), release("20260927", 10), release("20261004", 3)]
    assert getattr(load(), "candidate")(releases, PIN, NOW) == ("20260927", "b" * 64)


def test_drafts_prereleases_and_undated_tags_are_passed_over() -> None:
    releases = [
        release("20260927", 10, draft=True),
        release("20260928", 9, prerelease=True),
        release("latest", 9),
        release("20260920", 17, digest="a" * 64),
    ]
    assert getattr(load(), "candidate")(releases, PIN, NOW) is None


def test_a_pinned_release_whose_digest_changed_is_refused() -> None:
    with pytest.raises(ValueError, match="changed"):
        getattr(load(), "candidate")([release("20260920", 17, digest="c" * 64)], PIN, NOW)


def test_an_asset_served_from_elsewhere_is_refused() -> None:
    moved = release("20260927", 10)
    moved["assets"][0]["browser_download_url"] = "https://example.com/rules.zip"
    with pytest.raises(ValueError, match="not served"):
        getattr(load(), "candidate")([moved], PIN, NOW)


def test_an_archive_is_checked_before_it_is_pinned(tmp_path: Path) -> None:
    archive = tmp_path / "rules.zip"
    with zipfile.ZipFile(archive, "w") as zipped:
        zipped.writestr("packages/full/yara-rules-full.yar", "rule r { condition: false }")
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    verify = getattr(load(), "verify")
    verify(archive, digest, "full")
    with pytest.raises(ValueError, match="SHA-256"):
        verify(archive, "0" * 64, "full")
    with pytest.raises(ValueError, match="holds no"):
        verify(archive, digest, "core")
