"""Fixtures for the agent surface's tests. CR-003 Phase D, section 7.

The scenarios here are *tool shaped*: every step is something an MCP server
would expose as one tool call, with JSON-ready arguments and a JSON-ready
result, so that what these tests pin is the surface a server is written over
rather than the Python idiom.
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

#: A sample that carries ``w14:paraId`` on its paragraphs, and one that does not.
WITH_PARA_IDS = "DrawingML_GraphicData_wps.docx"
WITHOUT_PARA_IDS = "2010-sample1.docx"

#: The seed every test that compares bytes fixes, so that the ids are the same.
SEED = 20260916


def sample(name: str) -> WordprocessingMLPackage:
    """Load one of ``samples/``."""
    package = load(SAMPLES / name)
    package.id_seed = SEED
    return package


def reloaded(package: WordprocessingMLPackage) -> WordprocessingMLPackage:
    """Save the package to bytes and load it again: what Word would see."""
    return WordprocessingMLPackage.load(package.save())


def part_bytes(data: bytes, *, relationships: bool = False) -> dict[str, bytes]:
    """Every entry of a saved package, by name; the ``.rels`` left out."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {
            entry.filename: archive.read(entry.filename)
            for entry in archive.infolist()
            if relationships or not entry.filename.endswith(".rels")
        }


@pytest.fixture
def new_package() -> WordprocessingMLPackage:
    """A fresh document, seeded so that the paragraph ids are reproducible."""
    package = create_package()
    package.id_seed = SEED
    return package


@pytest.fixture
def report(new_package: WordprocessingMLPackage) -> WordprocessingMLPackage:
    """A small document with headings, body text and a table: enough to outline."""
    body = new_package.body
    body.insert_paragraph("Quarterly Report", style="Heading 1")
    body.insert_paragraph("The quick brown fox jumps over the lazy dog.")
    body.insert_paragraph("Revenue", style="Heading 2")
    body.insert_paragraph("Revenue rose in every region except the lazy one.")
    body.insert_xml(
        "<w:tbl><w:tblPr/><w:tblGrid/>"
        "<w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Total</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>120</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )
    body.insert_paragraph("Outlook", style="Heading 2")
    body.insert_paragraph("The end.")
    new_package.changes.clear()
    return new_package


def big_document(*, paragraphs: int = 2000, tables: int = 20, heading_every: int = 40):
    """A 200-page document: CR-003 section 7's token-budget fixture.

    Two thousand paragraphs (about ten to a page), a heading every fortieth, and
    twenty three-by-three tables spread through them, built with the tree-layer
    builders of Phase A so that building it costs a fraction of what parsing one
    would.
    """
    from docx4j_py.wml import p, tbl

    package = create_package()
    package.id_seed = SEED
    body = package.body
    content = body.content
    owner = body._owner()
    every = max(1, paragraphs // tables)
    for index in range(paragraphs):
        if index % heading_every == 0:
            content.append(p(f"Section {index // heading_every + 1}", style="Heading1"))
            continue
        content.append(
            p(
                f"Paragraph {index} of the quarterly report, which says something "
                f"about the quick brown fox and the lazy dog."
            )
        )
        if index % every == every - 1 and len(package.body) < paragraphs + tables:
            content.append(
                tbl(
                    [
                        ["Region", "Quarter", "Total"],
                        ["North", "Q1", str(index)],
                        ["South", "Q2", str(index * 2)],
                    ]
                )
            )
    for item in content:
        item.parent = owner
    package.changes.clear()
    return package
