"""A private reader leaves its tree as parsed: CR-002 section 16.3, test 1.

``XmlPart.read_contents()`` parses a part into a private slot, which a save
ignores, and section 16 has ``contents`` *adopt* that tree rather than parse
the bytes again. So the rule every private reader keeps --- it never
writes into the tree it reads --- is no longer a convention whose breach costs
nothing: a breach would be saved, as an edit nobody made. This file holds the
rule.

**The readers.** Every caller of ``read_contents`` in ``docx4j_py``
(``grep -rn read_contents docx4j_py``), by the part it reads and the reader
below that reaches it. :func:`test_every_read_contents_caller_is_listed` keeps
the list whole: a new caller fails it until it is added here and a reader
reaches it, and :func:`test_every_listed_caller_is_reached` proves each one is
reached over the documents.

- ``model/property_resolver.py``, ``PropertyResolver._read_source``: the
  styles. The resolver, and every reader that builds one.
- ``openpackaging/packages/wordprocessingml_package.py``,
  ``WordprocessingMLPackage.get_run_font_selector``: the theme and the
  settings; ``WordprocessingMLPackage.set_font_mapper``: the font table.
  The fonts.
- ``openpackaging/parts/wml/__init__.py``, ``MainDocumentPart.fonts_in_use``:
  the numbering. The fonts.
- ``model/text/trees.py``, ``_read``: the main document, the headers and
  footers, the notes, the comments and ``commentsExtended``, the settings.
  The text model, Markdown and the outline.
- ``model/text/extractor.py``, ``_SmartArt.emit``: the diagram data parts
  (``smartart``); ``Extractor._glossary``: the glossary document
  (``include: glossary``). The text model.
- ``model/chart/model.py``, ``ChartModel.read``: the charts (``charts``). The
  text model.
- ``model/text/metadata.py``, ``_contents``: the core, extended and custom
  properties (``metadata``). The text model.
- ``model/text/field_values.py``, ``FieldValues._keywords``: the core
  properties; ``model/fields/resolvers.py``, ``_contents``: the three
  properties parts and the settings; ``model/fields/formatting.py``,
  ``_decimal_symbol``: the settings (all three ``fields: update``). The text
  model.
- ``model/content/placeholders.py``, ``GlossaryEntries.load``: the glossary
  document, read when a body is first read; ``has_placeholder_styles``: a
  styles part, the glossary's before a save completes it (CR-019 Phase B).
  The placeholders, and every reader that reads a body.

The numbering emulator reads ``/word/numbering.xml`` and the styles' numbering
with lxml from ``get_xml()``, and ``describe()`` reads every part with lxml
from ``part.xml``: neither calls ``read_contents`` on its own account, and both
are run here regardless, the emulator through the property resolver it uses.

**The check.** After each reader, every typed part holding a tree --- private,
or live where the reader unmarshalled it (``fonts_in_use`` and ``outline()`` do,
through ``contents``) --- must marshal (``part._marshal(tree)``) to the bytes a
fresh parse of the same source marshals to, and its parent pointers must be
``link_parents``' (``tests/test_child.py``'s ``assert_parents_consistent``). A
deliberate write in a test reader, of a field or of a parent pointer, makes it
fail (:func:`test_a_reader_that_writes_is_caught`).

**The documents.** ``samples/``, ``tests/fixtures/text-extraction/`` and CR-010's
``tests/fixtures/revisions/``, every reader on every one: the ``slow`` run.
The fast suite runs every reader over :data:`FAST`, chosen for reach: the probe
has the charts, SmartArt, the glossary, the fields and the properties, and
:data:`WITH_FIELDS` the two fields no file of the corpus has.
"""

from __future__ import annotations

import ast
import dataclasses
import io
import sys
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import docx4j_py
from docx4j_py import load
from docx4j_py.child import Child, ChildList
from docx4j_py.model import style_util
from docx4j_py.model.listnumbering import NumberingStates, emulator_of
from docx4j_py.model.property_resolver import (
    PropertyResolver,
    heading_level_by_name,
)
from docx4j_py.model.text import (
    extract,
    markdown,
    markdown_blocks,
    plain_text,
)
from docx4j_py.openpackaging import XmlPart
from docx4j_py.traversal import iter_nodes, text_of
from docx4j_py.wml import P, PPr, R, RPr, Style, Tbl, TblPr

#: The package imported: the checkout's, or the installed wheel's under scripts/release_check.sh.
PACKAGE = Path(docx4j_py.__file__).parent
SAMPLES = ROOT / "samples"
FIXTURES = ROOT / "tests" / "fixtures"
TEXT_EXTRACTION = FIXTURES / "text-extraction"
REVISIONS = FIXTURES / "revisions"

#: The whole corpus of section 16.3: the samples' Word documents, the
#: text-extraction suite and CR-010's revisions fixtures.
CORPUS = sorted(
    [*SAMPLES.glob("*.docx"), *SAMPLES.glob("*.dotm")]
    + sorted(TEXT_EXTRACTION.rglob("*.doc[xm]"))
    + sorted(REVISIONS.rglob("*.docx"))
)

#: A document no file of the corpus is: ``2010-sample1.docx`` with a
#: ``KEYWORDS`` field and a ``DOCPROPERTY`` with a numeric picture, which are
#: what reach ``FieldValues._keywords`` and ``fields.formatting._decimal_symbol``
#: under ``fields: update``. Built by :func:`_with_fields`.
WITH_FIELDS = "2010-sample1-with-fields.docx"

#: The fast suite's documents: the probe (charts, SmartArt, a glossary, fields,
#: properties, text boxes, notes, comments, revisions), a revisions fixture with
#: lists, a sample with headers, and :data:`WITH_FIELDS`.
FAST: list[Path | str] = [
    TEXT_EXTRACTION / "phase0" / "extraction-probe.docx",
    REVISIONS / "core-ts" / "check25" / "25a-bullets.docx",
    SAMPLES / "Headers.docx",
    WITH_FIELDS,
]

_FIELDS = (
    b'<w:p><w:fldSimple w:instr=" KEYWORDS "><w:r><w:t>keywords</w:t></w:r></w:fldSimple></w:p>'
    b'<w:p><w:fldSimple w:instr=" DOCPROPERTY Words \\# &quot;0.0&quot; "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>'
)
_WITH_FIELDS_BYTES: list[bytes] = []


def _with_fields() -> bytes:
    """:data:`WITH_FIELDS`: the two fields put before the body's ``w:sectPr``."""
    if not _WITH_FIELDS_BYTES:
        out = io.BytesIO()
        with zipfile.ZipFile(SAMPLES / "2010-sample1.docx") as source, zipfile.ZipFile(out, "w") as target:
            for info in source.infolist():
                data = source.read(info)
                if info.filename == "word/document.xml":
                    assert data.count(b"<w:sectPr") == 1
                    data = data.replace(b"<w:sectPr", _FIELDS + b"<w:sectPr")
                target.writestr(info, data)
        _WITH_FIELDS_BYTES.append(out.getvalue())
    return _WITH_FIELDS_BYTES[0]


def _open(source: Path | str) -> Any:
    return load(_with_fields()) if source == WITH_FIELDS else load(source)


def _id(source: Path | str) -> str:
    if not isinstance(source, Path):
        return source
    return str(source.relative_to(FIXTURES if FIXTURES in source.parents else ROOT))


# ---------------------------------------------------------------------------
# what a reader reads
# ---------------------------------------------------------------------------


def _story_parts(pkg: Any) -> list[Any]:
    main = pkg.main_document_part
    parts = [main, *main.header_parts(), *main.footer_parts()]
    for name in ("footnotes_part", "endnotes_part", "comments_part", "glossary_document_part"):
        part = getattr(main, name, None)
        if part is not None:
            parts.append(part)
    return parts


def _stories(pkg: Any) -> list[tuple[Any, Any]]:
    """Each story part with its tree, read privately (as the text model reads them)."""
    return [(part, part.read_contents()) for part in _story_parts(pkg)]


def _paragraphs(root: Any) -> list[Any]:
    return [node for node in iter_nodes(root, mce="all") if isinstance(node, P)]


def _runs(paragraph: Any) -> list[Any]:
    return [node for node in iter_nodes(paragraph, mce="all") if isinstance(node, R)]


def _tables(root: Any) -> list[Any]:
    return [node for node in iter_nodes(root, mce="all") if isinstance(node, Tbl)]


# ---------------------------------------------------------------------------
# the readers
# ---------------------------------------------------------------------------


def read_resolver(pkg: Any) -> None:
    """``PropertyResolver`` and ``StyleUtil``, over every style and every paragraph, run and table."""
    pkg.get_property_resolver()
    resolver = PropertyResolver(pkg)
    resolver.get_document_default_ppr()
    resolver.get_document_default_rpr()
    resolver.get_resolved_default_paragraph_style()
    styles = resolver.styles_tree
    for style in (styles.style if styles is not None else None) or ():
        style_id = style.style_id
        resolver.get_effective_ppr(style_id)
        resolver.get_effective_rpr(style_id)
        resolver.get_chain_ppr(style_id)
        resolver.get_chain_rpr(style_id)
        resolver.ancestry(style_id)
        if style_id is not None:
            resolver.get_lvl_from_heading_style(style_id)
        heading_level_by_name(style)
        style_util.is_empty_style(style)
        style_util.apply_style(style, Style())
    for _part, root in _stories(pkg):
        for paragraph in _paragraphs(root):
            p_pr = paragraph.p_pr
            resolver.get_effective_ppr(p_pr)
            resolver.get_effective_paragraph_mark_rpr(p_pr)
            if p_pr is not None:
                style_util.is_empty_ppr(p_pr)
                style_util.has_direct_formatting_ppr(p_pr)
                style_util.apply_ppr(p_pr, PPr())
            for run in _runs(paragraph):
                resolver.get_effective_rpr(run.r_pr, p_pr)
                resolver.has_direct_rpr_formatting(run.r_pr)
                if run.r_pr is not None:
                    style_util.is_empty_rpr(run.r_pr)
                    style_util.apply_rpr(run.r_pr, RPr())
        for table in _tables(root):
            resolver.get_effective_table_style(table.tbl_pr)
            resolver.reaches_default_table_style(table.tbl_pr)
            if table.tbl_pr is not None:
                style_util.is_empty_tbl_pr(table.tbl_pr)
                style_util.apply_tbl_pr(table.tbl_pr, TblPr())


def read_numbering(pkg: Any) -> None:
    """The numbering emulator: every paragraph of every story numbered, every level's indent read."""
    emulator = emulator_of(pkg)
    if emulator is None:
        return
    part = pkg.main_document_part.numbering_definitions_part
    if part is not None:
        part.get_emulator()
    emulator.abstract_list_definitions  # noqa: B018 - read for its effect
    instances = emulator.instance_list_definitions
    emulator.styles  # noqa: B018
    states = NumberingStates()
    for story_part, root in _stories(pkg):
        state = states.for_part(story_part)
        for paragraph in _paragraphs(root):
            p_pr = paragraph.p_pr
            emulator.peek_number(p_pr, state)
            result = emulator.number(p_pr, state)
            if result is not None:
                result.to_dict()
            num_pr = None if p_pr is None else p_pr.num_pr
            if num_pr is not None:
                emulator.get_ind(num_pr)
    for num_id in list(instances):
        for ilvl in range(9):
            emulator.ind_of(num_id, str(ilvl))
            emulator.indent_for(num_id, ilvl)


def read_fonts(pkg: Any) -> None:
    """The font selector over every run, ``fonts_in_use``, ``styles_in_use`` and the font mapper."""
    main = pkg.main_document_part
    selector = pkg.get_run_font_selector()
    selector.default_font  # noqa: B018
    resolver = pkg.get_property_resolver()
    for _part, root in _stories(pkg):
        for paragraph in _paragraphs(root):
            p_pr = paragraph.p_pr
            for run in _runs(paragraph):
                text = text_of(run)
                effective = resolver.get_effective_rpr(run.r_pr, p_pr)
                selector.spans(p_pr, effective, text, r_pr_is_effective=True)
                selector.spans(p_pr, run.r_pr, text)
    fonts = main.fonts_in_use()
    main.styles_in_use()
    mapper = pkg.get_font_mapper()
    for font in sorted(fonts):
        mapper.get_decision(font)


#: Every option of WD section 10 the adapter can turn on, at the value that
#: reaches the most code, and the same with the other values of the choices.
ALL_ON: dict[str, Any] = {
    "include": ["headers", "footers", "comments", "glossary"],
    "text_boxes": True,
    "notes": True,
    "note_references": True,
    "comments": True,
    "hidden": True,
    "field_codes": True,
    "fields": "update",
    "toc": True,
    "placeholders": True,
    "bindings": "surface",
    "list_labels": True,
    "charts": True,
    "smartart": True,
    "embedded": True,
    "altchunks": True,
    "alt_text": True,
    "metadata": True,
    "tables": "grid",
    "markup_detail": "author_date",
    "page_markers": True,
    "layout": "indent",
    "warnings": True,
}
ALL_ON_OTHERWISE: dict[str, Any] = {
    **ALL_ON,
    "field_codes": False,
    "bindings": "bound",
    "tables": "tsv",
    "markup_detail": "author",
    "layout": "none",
}
OPTION_SETS: tuple[dict[str, Any], ...] = ({}, ALL_ON, ALL_ON_OTHERWISE)


def read_text(view: str) -> Callable[[Any], None]:
    """The text model in one view: ``extract`` under each option set, written as plain text and Markdown."""

    def reader(pkg: Any) -> None:
        for options in OPTION_SETS:
            result = extract(pkg, view=view, **options)
            plain_text(result)
            markdown(result, pkg)
            markdown_blocks(result, pkg)
        main = pkg.main_document_part
        for part in (main, *main.header_parts(), *main.footer_parts()):
            plain_text(extract(part, view=view, **ALL_ON))

    return reader


def read_markdown(pkg: Any) -> None:
    """``pkg.to_markdown()`` in each view, with and without addresses, and ``markdown_budget``."""
    for view in ("accepted", "original", "markup"):
        pkg.to_markdown(view=view)
        pkg.to_markdown(view=view, addresses=True)
    pkg.markdown_budget(200, addresses=True)


def read_outline(pkg: Any) -> None:
    """``outline()``, whole and headings only."""
    pkg.outline()
    pkg.outline(headings_only=True, depth=1)


def read_describe(pkg: Any) -> None:
    """``describe()``."""
    pkg.describe()


def read_placeholders(pkg: Any) -> None:
    """Every control's placeholder, read from the glossary (CR-019 Phase B), and the placeholder styles' check."""
    from docx4j_py.model.content.placeholders import has_placeholder_styles

    for body in pkg.bodies():
        for control in body.content_controls:
            control.placeholder_text  # noqa: B018 - read for its effect
            control.shows_words_form_placeholder  # noqa: B018
    main = pkg.main_document_part
    glossary = main.glossary_document_part
    for part in (main.style_definitions_part, getattr(glossary, "style_definitions_part", None)):
        if part is not None:
            has_placeholder_styles(part)


READERS: dict[str, Callable[[Any], None]] = {
    "resolver": read_resolver,
    "numbering": read_numbering,
    "fonts": read_fonts,
    "text-accepted": read_text("accepted"),
    "text-original": read_text("original"),
    "text-markup": read_text("markup"),
    "markdown": read_markdown,
    "outline": read_outline,
    "describe": read_describe,
    "placeholders": read_placeholders,
}


# ---------------------------------------------------------------------------
# the check
# ---------------------------------------------------------------------------

_FRESH: dict[tuple[str, str], bytes | None] = {}


def _fresh(source: Path | str, part_name: str) -> bytes | None:
    """What a fresh parse of the part marshals to: a new load of `source`, the part read, marshalled."""
    key = (str(source), part_name)
    if key not in _FRESH:
        part = _open(source).get_part(part_name)
        _FRESH[key] = None if part is None else part._marshal(part.read_contents())
    return _FRESH[key]


def _parent_problem(root: Any) -> str | None:
    """Where the tree breaks ``link_parents``' invariants, or None. ``assert_parents_consistent``'s walk."""
    stack = [(root, root.get_parent())]
    seen = {id(root)}
    while stack:
        obj, want = stack.pop()
        if obj.get_parent() is not want:
            return f"a {type(obj).__name__} has the wrong parent"
        for field in dataclasses.fields(obj):
            value = getattr(obj, field.name, None)
            if isinstance(value, ChildList) and value.owner is not obj:
                return f"{type(obj).__name__}.{field.name} is bound to another owner"
            items = value if isinstance(value, (list, tuple)) else (value,)
            for item in items:
                if isinstance(item, Child) and id(item) not in seen:
                    seen.add(id(item))
                    stack.append((item, obj))
    return None


def changed_trees(pkg: Any, source: Path | str) -> list[str]:
    """Every typed part whose tree is no longer as parsed: what a reader wrote. Empty is the rule held."""
    out: list[str] = []
    for part in pkg.parts.parts():
        if not isinstance(part, XmlPart) or part.model_class is None:
            continue
        for slot, kind in (("_private_contents", "private"), ("_contents", "live")):
            tree = getattr(part, slot)
            if tree is None:
                continue
            name = str(part.part_name)
            expected = _fresh(source, name)
            if expected is None:
                out.append(f"{name} ({kind}): a part the source does not have")
            elif part._marshal(tree) != expected:
                out.append(f"{name} ({kind}): marshals differently from a fresh parse")
            problem = _parent_problem(tree)
            if problem is not None:
                out.append(f"{name} ({kind}): {problem}")
    return out


def _private_count(pkg: Any) -> int:
    return sum(
        1
        for part in pkg.parts.parts()
        if isinstance(part, XmlPart) and (part._private_contents is not None or part._contents is not None)
    )


def _check(source: Path | str, reader: str) -> int:
    pkg = _open(source)
    READERS[reader](pkg)
    assert changed_trees(pkg, source) == [], f"{_id(source)}: {reader} wrote into a tree it read"
    return _private_count(pkg)


# ---------------------------------------------------------------------------
# the tests
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("reader", list(READERS))
@pytest.mark.parametrize("source", FAST, ids=_id)
def test_a_private_reader_leaves_its_trees_as_parsed(source: Path | str, reader: str) -> None:
    """Section 16.3, test 1, over the fast documents."""
    trees = _check(source, reader)
    if reader != "describe":
        assert trees > 0, f"{reader} read no part of {_id(source)}"


@pytest.mark.slow
@pytest.mark.parametrize("reader", list(READERS))
@pytest.mark.parametrize("path", CORPUS, ids=_id)
def test_a_private_reader_leaves_its_trees_as_parsed_over_the_corpus(path: Path, reader: str) -> None:
    """Section 16.3, test 1, over ``samples/``, the text-extraction suite and the revisions fixtures."""
    _check(path, reader)


def test_the_corpus_is_the_one_section_16_names() -> None:
    assert len([p for p in CORPUS if SAMPLES in p.parents]) >= 13
    assert len([p for p in CORPUS if TEXT_EXTRACTION in p.parents]) >= 10
    assert len([p for p in CORPUS if REVISIONS in p.parents]) >= 40
    assert all(source in CORPUS for source in FAST if isinstance(source, Path))


def _write_a_field(pkg: Any) -> None:
    styles = pkg.main_document_part.style_definitions_part.read_contents()
    styles.style[0].style_id = f"{styles.style[0].style_id}Written"


def _write_into_the_document(pkg: Any) -> None:
    document = pkg.main_document_part.read_contents()
    paragraph = _paragraphs(document)[0]
    paragraph.p_pr = PPr() if paragraph.p_pr is None else None


def _reparent(pkg: Any) -> None:
    document = pkg.main_document_part.read_contents()
    _paragraphs(document)[0].set_parent(None)


@pytest.mark.parametrize("writer", [_write_a_field, _write_into_the_document, _reparent])
def test_a_reader_that_writes_is_caught(writer: Callable[[Any], None]) -> None:
    """The check is not vacuous: a test reader that writes into a private tree fails it."""
    path = SAMPLES / "2010-sample1.docx"
    pkg = load(path)
    read_resolver(pkg)
    assert changed_trees(pkg, path) == []
    writer(pkg)
    found = changed_trees(pkg, path)
    assert len(found) == 1, found
    assert "(private)" in found[0]


# ---------------------------------------------------------------------------
# the list of callers, kept whole
# ---------------------------------------------------------------------------

#: Every caller of ``read_contents`` under ``docx4j_py``, as ``path::qualname``:
#: the table of the module docstring.
CALLERS: frozenset[str] = frozenset(
    {
        "model/property_resolver.py::PropertyResolver._read_source",
        "openpackaging/packages/wordprocessingml_package.py::WordprocessingMLPackage.get_run_font_selector",
        "openpackaging/packages/wordprocessingml_package.py::WordprocessingMLPackage.set_font_mapper",
        "openpackaging/parts/wml/__init__.py::MainDocumentPart.fonts_in_use",
        "model/text/trees.py::_read",
        "model/text/extractor.py::_SmartArt.emit",
        "model/text/extractor.py::Extractor._glossary",
        "model/chart/model.py::ChartModel.read",
        "model/text/metadata.py::_contents",
        "model/text/field_values.py::FieldValues._keywords",
        "model/fields/resolvers.py::_contents",
        "model/fields/formatting.py::_decimal_symbol",
        "model/content/placeholders.py::GlossaryEntries.load",
        "model/content/placeholders.py::has_placeholder_styles",
    }
)

#: Not callers: the definitions themselves.
_DEFINITIONS = frozenset(
    {
        "openpackaging/parts/xml_part.py::XmlPart.read_contents",
        "model/content/trial.py::TrialPart.read_contents",
    }
)


def _callers_in_source() -> set[str]:
    package = PACKAGE
    found: set[str] = set()
    for path in sorted(package.rglob("*.py")):
        source = path.read_text(encoding="utf-8")
        if "read_contents" not in source:
            continue
        tree = ast.parse(source)
        relative = path.relative_to(package).as_posix()

        def visit(node: ast.AST, scope: list[str], relative: str = relative) -> None:
            for child in ast.iter_child_nodes(node):
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                    visit(child, [*scope, child.name])
                    continue
                if (
                    isinstance(child, ast.Call)
                    and isinstance(child.func, ast.Attribute)
                    and child.func.attr == "read_contents"
                ):
                    found.add(f"{relative}::{'.'.join(scope)}")
                visit(child, scope)

        visit(tree, [])
    return found - _DEFINITIONS


def test_every_read_contents_caller_is_listed() -> None:
    """A new private reader fails this until it is in the table, and reached by a reader here."""
    assert _callers_in_source() == set(CALLERS)


def _reached(sources: list[Path | str]) -> set[str]:
    """The callers the readers reach over `sources`, by the frame that called ``read_contents``."""
    package = PACKAGE
    reached: set[str] = set()
    original = XmlPart.read_contents

    def recording(self: Any) -> Any:
        frame = sys._getframe(1)
        filename = Path(frame.f_code.co_filename)
        if package in filename.parents:
            relative = filename.relative_to(package).as_posix()
            reached.add(f"{relative}::{frame.f_code.co_qualname}")
        return original(self)

    XmlPart.read_contents = recording  # type: ignore[method-assign]
    try:
        for source in sources:
            for reader in READERS.values():
                reader(_open(source))
    finally:
        XmlPart.read_contents = original  # type: ignore[method-assign]
    return reached


def test_every_listed_caller_is_reached() -> None:
    """Each caller of the table runs, over the fast documents, under one reader or another."""
    reached = _reached(FAST)
    assert set(CALLERS) - reached == set()
