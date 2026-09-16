"""Fixtures for the content API's tests. CR-003 Phase B, section 7.

Two helpers carry most of the weight. :func:`reloaded` is the rule CR-003
section 7 sets for every phase --- "every mutation followed by save, reload and
read back" --- and :func:`part_bytes` is the other one, "untouched parts still
byte-identical through the content API".
"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import WordprocessingMLPackage, create_package, load

#: The sample documents this phase reads.
SAMPLES = ROOT / "samples"


@pytest.fixture
def new_package() -> WordprocessingMLPackage:
    """A fresh document, as ``create_package()`` makes it."""
    package = create_package()
    package.id_seed = 20260916  # CR-003 section 3.4: the same calls, the same bytes
    return package


@pytest.fixture
def body(new_package: WordprocessingMLPackage):
    """The new document's body."""
    return new_package.body


def sample(name: str) -> WordprocessingMLPackage:
    """Load one of ``samples/``."""
    return load(SAMPLES / name)


def reloaded(package: WordprocessingMLPackage) -> WordprocessingMLPackage:
    """Save the package to bytes and load it again: what Word would see."""
    return WordprocessingMLPackage.load(package.save())


def part_bytes(data: bytes, *, relationships: bool = False) -> dict[str, bytes]:
    """Every entry of a saved package, by name.

    The relationship parts are left out by default: CR-002's own round-trip
    test makes the byte-for-byte promise for "every non-relationship entry",
    because a relationship part is rebuilt from the model on every save (its
    attribute order is the serialiser's, not the source's).
    """
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {
            entry.filename: archive.read(entry.filename)
            for entry in archive.infolist()
            if relationships or not entry.filename.endswith(".rels")
        }
