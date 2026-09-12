"""The round trip over samples/ is clean, CR-001 section 9.

The contract after Phase B: every part of every sample document parses,
serialises, and comes back canonically identical, with nothing skipped. The
only accepted difference is the spelling of an `xsd:boolean` -- Word writes
`1`/`0`, some producers write `true`/`false`, and a serialiser has to pick one
(`SerializerConfig(bool_format="numeric")` picks Word's).

This runs the real harness, `scripts/roundtrip.py`, in a subprocess, so what is
tested is the command the CR quotes its numbers from and not a reimplementation
of it.

    .venv-fork/bin/python -m pytest tests/test_roundtrip.py -q
    .venv-fork/bin/python -m pytest tests/ -q -m "not slow"     # skips it
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

pytestmark = pytest.mark.slow


@pytest.fixture(scope="module")
def harness(tmp_path_factory) -> subprocess.CompletedProcess:
    out = tmp_path_factory.mktemp("roundtrip")
    return subprocess.run(
        [
            sys.executable,
            str(ROOT / "scripts" / "roundtrip.py"),
            "--models-module",
            "docx4j_py.wml",
            "--runtime",
            "docx4j_xsdata",
            "--out",
            str(out),
            "--repeat",
            "1",
        ],
        capture_output=True,
        text=True,
    )


def verdict(harness, label: str) -> int:
    match = re.search(rf"^\s*{re.escape(label)}\s+(\d+)", harness.stdout, re.M)
    assert match, f"{label!r} not in the harness output:\n{harness.stdout[-3000:]}"
    return int(match.group(1))


def test_the_harness_succeeds(harness) -> None:
    assert harness.returncode == 0, harness.stdout[-4000:] + harness.stderr[-2000:]
    assert "ROUND TRIP CLEAN" in harness.stdout


def test_every_part_parses_and_serialises(harness) -> None:
    assert verdict(harness, "parts") == 50
    assert verdict(harness, "parse/serialise failures") == 0


def test_every_part_is_canonically_identical(harness) -> None:
    assert verdict(harness, "canonically identical") == 50


def test_there_are_no_differences_in_any_category(harness) -> None:
    assert verdict(harness, "differences") == 0
    for category in (
        "element-dropped",
        "element-added",
        "element-renamed",
        "child-order",
        "text",
        "attribute-added",
        "attribute-dropped",
        "attribute-value",
    ):
        assert f"  {category}" not in harness.stdout, category


def test_nothing_was_skipped(harness) -> None:
    assert verdict(harness, "skipped content") == 0
    assert "none, over 50 parts" in harness.stdout


def test_the_only_respellings_are_booleans(harness) -> None:
    # 34 of them, all in one styles.xml written with true/false rather than
    # Word's 1/0; with bool_format="words" the same corpus gives 3,034 the
    # other way round, which is why numeric is the default (CR-001 q. 5).
    assert verdict(harness, "benign respellings") == 34
    assert "['boolean-spelling']" in harness.stdout
