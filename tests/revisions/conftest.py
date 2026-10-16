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

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
# scripts/canon.py, the prefix-insensitive normal form
sys.path.insert(0, str(ROOT / "scripts"))

#: The copied fixtures.
FIXTURES = ROOT / "tests" / "fixtures" / "revisions"

#: docx4j-core-tests' revisions fixtures, in the sibling checkout.
SIBLING = ROOT.parent / "docx4j" / "docx4j-core-tests" / "src" / "test" / "resources" / "revisions"
