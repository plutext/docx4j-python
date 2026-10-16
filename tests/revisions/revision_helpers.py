"""Helpers for CR-010's tests: loading, round-tripping, and handwritten bodies."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from description import describe

from docx4j_py import WordprocessingMLPackage, create_package, load
from docx4j_py.model import revisions
from docx4j_py.model.revisions.lister import Lister

ROOT = Path(__file__).resolve().parent.parent.parent
FIXTURES = ROOT / "tests" / "fixtures" / "revisions"
SIBLING = ROOT.parent / "docx4j" / "docx4j-core-tests" / "src" / "test" / "resources" / "revisions"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (
    f'xmlns:w="{W}" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"'
)


def fixture(name: str) -> WordprocessingMLPackage:
    """One of the copied fixtures."""
    return load(FIXTURES / name)


def round_trip(package: WordprocessingMLPackage) -> WordprocessingMLPackage:
    """Saved and loaded again: what is compared is what this package writes."""
    return WordprocessingMLPackage.load(package.save())


def unresolved(package: Any) -> list[Any]:
    """What is left that this release resolves, in any story or branch."""
    return [change for change in Lister(package, False).list() if change.resolvable]


def assert_like_word(ours: Any, word: Any, *, records: bool, label: str) -> None:
    """Two packages described alike; on a difference, the lines that differ."""
    mine, theirs = describe(round_trip(ours), records), describe(word, records)
    if mine != theirs:
        import difflib

        diff = "\n".join(difflib.unified_diff(theirs.splitlines(), mine.splitlines(), "word", "ours", lineterm=""))
        raise AssertionError(f"{label}: differs from Word's save\n{diff}")


def with_body(body_xml: str) -> WordprocessingMLPackage:
    """A package whose body is `body_xml` (the content of ``w:body``)."""
    package = create_package()
    package.main_document_part.set_xml(f"<w:document {NS}><w:body>{body_xml}</w:body></w:document>")
    return package


def ins(identifier: int, author: str, text: str, extra: str = "") -> str:
    return (
        f'<w:ins w:id="{identifier}" w:author="{author}"{extra}><w:r>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:ins>'
    )


def dele(identifier: int, author: str, text: str) -> str:
    return (
        f'<w:del w:id="{identifier}" w:author="{author}"><w:r>'
        f'<w:delText xml:space="preserve">{text}</w:delText></w:r></w:del>'
    )


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def paragraphs(package: Any) -> list[Any]:
    """The body's ``w:p``\\ s, top level."""
    from docx4j_py.wml import P

    return [item for item in package.main_document_part.contents.body.content if isinstance(item, P)]


def text_of(p: Any) -> str:
    """A paragraph's accepted text, as the engine reads it."""
    from docx4j_py.model.revisions import trees

    return trees.paragraph_text(p, trees.View.ACCEPTED)


def body_xml(package: Any) -> str:
    return package.main_document_part.get_xml()


__all__ = [
    "FIXTURES",
    "NS",
    "SIBLING",
    "assert_like_word",
    "body_xml",
    "dele",
    "describe",
    "fixture",
    "ins",
    "paragraphs",
    "revisions",
    "round_trip",
    "run",
    "text_of",
    "unresolved",
    "with_body",
]
