"""Fixtures for CR-010's tests: the revisions engine against Word's saves.

The Word-made fixtures of docx4j CR-043's agreement run, and Word's own Accept
All and Reject All saves of them, are in ``tests/fixtures/revisions/`` (copied
from docx4j-core-tests at ``00614748e``; ``tests/README.md`` has the
provenance). Everything else under docx4j-core-tests'
``src/test/resources/revisions/`` is read from the sibling checkout, and the
tests that need it skip without it (CR-010 section 9, Q4).
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
# scripts/canon.py, the prefix-insensitive normal form
sys.path.insert(0, str(ROOT / "scripts"))

#: The copied fixtures.
FIXTURES = ROOT / "tests" / "fixtures" / "revisions"

#: docx4j-core-tests' revisions fixtures, in the sibling checkout.
SIBLING = ROOT.parent / "docx4j" / "docx4j-core-tests" / "src" / "test" / "resources" / "revisions"


@pytest.fixture
def zone(monkeypatch):
    """Set the process's time zone for one test: ``zone("Australia/Sydney")``; restored after.

    CR-010 section 12.9: ``w:date`` is written and read on the process's local
    wall clock, so a test of it fixes the zone (``TZ`` and ``time.tzset()``)
    rather than depend on the machine's.
    """
    import time

    def set_zone(name: str) -> None:
        monkeypatch.setenv("TZ", name)
        time.tzset()

    yield set_zone
    monkeypatch.undo()
    time.tzset()
