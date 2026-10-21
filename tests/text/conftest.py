"""Fixtures for CR-011's tests: docx4j's text-extraction suite and its harness.

The suite is copied from docx4j at ``198d015a3`` into
``tests/fixtures/text-extraction/`` (``tests/README.md`` has the provenance);
``conformance.py`` is the port of docx4j's harness and ``adapter.py`` its
engine over this package's readers. docx4j's own revisions and text-extraction
fixtures are read from the sibling checkout, and the tests that need them skip
without it.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
