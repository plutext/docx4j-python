"""CR-021 phase C: the four edges nobody had measured, held to Word 365's saves.

The fixtures are ``tests/fixtures/integers-edges/`` (``README.md``): the raw edge files of
CR-021's Word check and Word's own saves of them, from the user's run of 2026-10-26.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx4j_py import load
from docx4j_py.model.values import integer

EDGES = Path(__file__).resolve().parents[1] / "fixtures" / "integers-edges"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: case: (the element and attribute it sets, the value as written)
CASES = {
    "E1": ("spacing", "before", "239.99999999999999"),
    "E2": ("spacing", "before", "."),
    "E3": ("spacing", "before", "+"),
    "E4": ("ind", "left", "-"),
}


def test_the_fixtures_are_as_copied() -> None:
    """Every file ``SHA1SUMS`` lists, byte for byte, and nothing else beside it."""
    listed = dict(reversed(line.split("  ", 1)) for line in (EDGES / "SHA1SUMS").read_text(encoding="utf-8").splitlines())
    listed = {name: digest for name, digest in listed.items()}
    for name, digest in listed.items():
        assert hashlib.sha1((EDGES / name).read_bytes()).hexdigest() == digest, name
    assert sorted(p.name for p in EDGES.iterdir() if p.name not in ("SHA1SUMS", "README.md")) == sorted(listed)


def _value(path: Path, case: str) -> str | None:
    """The case's attribute in the file's document part, or None where the file has none."""
    element, attribute = CASES[case][:2]
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    for p in root.iter(W + "p"):
        if "".join(t.text or "" for t in p.iter(W + "t")).startswith(case):
            found = p.find(f"{W}pPr/{W}{element}")
            return None if found is None else found.get(W + attribute)
    raise AssertionError(f"no paragraph labelled {case} in {path.name}")


@pytest.mark.parametrize("case", sorted(CASES))
def test_word_opened_each_edge_and_values_integer_reads_it_as_word(case: str) -> None:
    """E1 truncates the string (239, where a double reads 240); E2 to E4 are accepted, and read 0.

    Word's saves of E2 to E4 hold no attribute: Word leaves out direct formatting equal to the
    style's, and the style's value is 0 here, so 0 is Word's reading or indistinguishable from it.
    """
    source = EDGES / f"cr021-edge-{case}-raw.docx"
    word = EDGES / f"cr021-edge-{case}-raw-word365.docx"
    written = CASES[case][2]
    assert _value(source, case) == written
    saved = _value(word, case)
    if case == "E1":
        assert saved == "239"
        assert integer(written) == 239
    else:
        assert saved is None
        assert integer(written) == 0


def test_the_four_together_read_the_same() -> None:
    """``cr021-edges-raw.docx``, all four in one file, opened too, and Word read each as alone."""
    for case in CASES:
        assert _value(EDGES / "cr021-edges-raw-word365.docx", case) == _value(EDGES / f"cr021-edge-{case}-raw-word365.docx", case)


@pytest.mark.parametrize("case", sorted(CASES))
def test_a_normalised_save_writes_what_word_reads(case: str) -> None:
    """Our save after an edit writes 239 for E1, and 0 for E2 to E4 (Word's save leaves the 0 out)."""
    package = load(EDGES / f"cr021-edge-{case}-raw.docx")
    package.body.paragraphs[0].insert_text(" edited", location="End")
    data = package.save()
    path = io.BytesIO(data)
    element, attribute = CASES[case][:2]
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    values = [e.get(W + attribute) for e in root.iter(W + element) if e.get(W + attribute) is not None]
    assert values == (["239"] if case == "E1" else ["0"])
