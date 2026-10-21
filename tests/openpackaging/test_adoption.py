"""``contents`` adopts the tree ``read_contents`` parsed: CR-002 section 16, tests 2 to 5.

``read_contents()`` parses a part into a private slot, which a save ignores.
Since section 16, ``contents`` finding that tree promotes it to the live slot
instead of parsing the bytes again, when it was parsed under the load options a
parse now would use. Test 1, that no private reader writes into its tree, is
``test_private_readers.py``. Here:

* the unit: the very same object, one parse, the slots as an unmarshal leaves
  them, and ``set_bytes`` and ``set_contents`` still winning;
* **test 2**, an adopted tree is a parsed one: its parent pointers are
  ``link_parents``', its ``skipped`` is an ordinary unmarshal's, and
  ``strict=True`` still raises;
* **test 3**, a save after adoption writes the bytes a save after a fresh parse
  writes: CR-011 section 14.4's read-then-edit-by-address workflow, a resolver
  read followed by a style edit, and every sample read by every private reader
  and then unmarshalled;
* **test 4**, changed load options block adoption;
* **test 5**, the thread rule unchanged: one package per thread, the same bytes
  as a sequential run.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_adoption.py -q
"""

from __future__ import annotations

import io
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from test_child import assert_parents_consistent

from docx4j_py import load
from docx4j_py.model.markdown import ADDRESS_COMMENT
from docx4j_py.openpackaging import LoadOptions, XmlPart
from docx4j_py.openpackaging.exceptions import Docx4JException

SAMPLES = ROOT / "samples"
FIXTURES = ROOT / "tests" / "fixtures"
SAMPLE = SAMPLES / "2010-sample1.docx"
#: A sample with ``mc:AlternateContent`` in its main part, so that
#: ``mce_preprocess`` gives a different tree.
WITH_MCE = SAMPLES / "2010-glow-then-AlternateContent.docx"
#: The seed every byte comparison fixes, so that new paraIds are the same.
SEED = 20260916


def entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def with_document(source: Path, document: bytes) -> bytes:
    """`source` with its ``word/document.xml`` replaced."""
    out = io.BytesIO()
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info)
            zout.writestr(info, document if info.filename == "word/document.xml" else data)
    return out.getvalue()


def with_unknown_content() -> bytes:
    """A document whose main part has an element and an attribute no schema declares."""
    return with_document(SAMPLE, (FIXTURES / "unknown_content.xml").read_bytes())


class Parses:
    """Counts the parses ``XmlPart._unmarshal`` makes of one part, private and live.

    By part name, since loading a package parses its relationships parts.
    """

    def __init__(self, monkeypatch: pytest.MonkeyPatch, part_name: str) -> None:
        self.live = 0
        self.private = 0
        original = XmlPart._unmarshal

        def counting(part: XmlPart, private: bool = False) -> None:
            if str(part.part_name) == part_name:
                if private:
                    self.private += 1
                else:
                    self.live += 1
            original(part, private)

        monkeypatch.setattr(XmlPart, "_unmarshal", counting)

    @property
    def counts(self) -> tuple[int, int]:
        """(private, live)."""
        return self.private, self.live


# ---------------------------------------------------------------------------
# the unit
# ---------------------------------------------------------------------------


def test_contents_is_the_very_tree_read_contents_parsed(monkeypatch: pytest.MonkeyPatch) -> None:
    parses = Parses(monkeypatch, "/word/document.xml")
    pkg = load(SAMPLE)
    part = pkg.main_document_part
    source = part.xml

    private = part.read_contents()
    assert parses.counts == (1, 0)
    assert not part.is_unmarshalled
    assert not part.is_loaded
    assert part.xml == source, "a private read leaves the part as it came in"

    assert part.contents is private
    assert parses.counts == (1, 0), "adopted, not parsed again"
    assert part.is_unmarshalled
    assert part.is_loaded
    assert part._private_contents is None
    assert part._private_settings is None
    assert part._bytes is None

    assert part.read_contents() is private, "read_contents answers the live tree"
    assert part.contents is private
    assert parses.counts == (1, 0)


def test_without_a_private_read_contents_parses_as_before(monkeypatch: pytest.MonkeyPatch) -> None:
    parses = Parses(monkeypatch, "/word/document.xml")
    part = load(SAMPLE).main_document_part
    live = part.contents
    assert parses.counts == (0, 1)
    assert part.read_contents() is live
    assert parses.counts == (0, 1)


def test_the_resolver_and_the_styles_part_share_one_parse(monkeypatch: pytest.MonkeyPatch) -> None:
    """Section 16.1's other case: a resolver read, then an edit of the styles."""
    parses = Parses(monkeypatch, "/word/styles.xml")
    pkg = load(SAMPLE)
    resolver = pkg.get_property_resolver()
    styles_part = pkg.style_definitions_part
    assert not styles_part.is_unmarshalled
    assert parses.counts == (1, 0)
    assert styles_part.contents is resolver.styles_tree
    assert parses.counts == (1, 0)


def test_set_bytes_after_a_private_read_is_parsed_not_adopted() -> None:
    part = load(SAMPLE).main_document_part
    private = part.read_contents()
    data = part.xml.replace(b"<w:body>", b"<w:body><w:p/>", 1)
    part.set_bytes(data)
    assert part._private_contents is None
    live = part.contents
    assert live is not private
    assert len(live.body.content) == len(private.body.content) + 1


def test_set_contents_after_a_private_read_wins() -> None:
    from docx4j_py.wml import Body, Document

    part = load(SAMPLE).main_document_part
    part.read_contents()
    document = Document(body=Body())
    part.set_contents(document)
    assert part.contents is document
    assert part.read_contents() is document
    assert part._private_contents is None


def test_a_private_read_of_set_bytes_is_adopted_and_the_bytes_cleared() -> None:
    part = load(SAMPLE).main_document_part
    data = part.xml.replace(b"<w:body>", b"<w:body><w:p/>", 1)
    part.set_bytes(data)
    private = part.read_contents()
    assert part._bytes is not None
    assert part.contents is private
    assert part._bytes is None


# ---------------------------------------------------------------------------
# test 2: an adopted tree is a parsed one
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", ["/word/document.xml", "/word/styles.xml", "/word/settings.xml"])
def test_an_adopted_tree_has_its_parent_pointers(name: str) -> None:
    adopted = load(SAMPLES / "DrawingML_GraphicData_wps.docx").get_part(name)
    adopted.read_contents()
    fresh = load(SAMPLES / "DrawingML_GraphicData_wps.docx").get_part(name)
    count = assert_parents_consistent(adopted.contents)
    assert count == assert_parents_consistent(fresh.contents) > 10
    assert adopted.contents.get_parent() is None


def test_an_adopted_tree_reports_what_an_unmarshal_skips() -> None:
    data = with_unknown_content()
    fresh = load(data)
    fresh_part = fresh.main_document_part
    fresh_part.contents  # noqa: B018 - read for its effect
    assert len(fresh_part.skipped) == 2, "the fixture's two intruders"

    adopted = load(data)
    part = adopted.main_document_part
    tree = part.read_contents()
    assert part.contents is tree
    assert [repr(item) for item in part.skipped] == [repr(item) for item in fresh_part.skipped]
    assert [(str(n), repr(i)) for n, i in adopted.skipped] == [(str(n), repr(i)) for n, i in fresh.skipped]
    assert part.source_ns_map == fresh_part.source_ns_map
    assert part.xml == fresh_part.xml, "and it marshals as the unmarshalled one does"


def test_strict_raises_at_the_private_read_and_adopts_nothing() -> None:
    pkg = load(with_unknown_content(), options=LoadOptions(strict=True))
    part = pkg.main_document_part
    with pytest.raises(Docx4JException, match="strict=True"):
        part.read_contents()
    assert part._private_contents is None
    with pytest.raises(Docx4JException, match="strict=True"):
        part.contents  # noqa: B018 - raises
    assert not part.is_unmarshalled


def test_strict_raises_at_contents_after_a_lenient_private_read() -> None:
    """A tree that skipped content is never adopted under strict: the options differ, so the parse runs and raises."""
    pkg = load(with_unknown_content())
    part = pkg.main_document_part
    part.read_contents()
    assert len(part.skipped) == 2
    pkg.load_options.strict = True
    with pytest.raises(Docx4JException, match="strict=True"):
        part.contents  # noqa: B018 - raises
    assert not part.is_unmarshalled


def test_strict_with_nothing_skipped_adopts() -> None:
    pkg = load(SAMPLE, options=LoadOptions(strict=True))
    part = pkg.main_document_part
    tree = part.read_contents()
    assert part.contents is tree
    assert part.skipped == []


# ---------------------------------------------------------------------------
# test 3: a save after adoption writes what a save after a fresh parse writes
# ---------------------------------------------------------------------------

#: CR-011 section 14.4's documents (ordinal addresses), one with paraIds, one
#: with notes, and the probe.
WORKFLOW = [
    SAMPLE,
    SAMPLES / "DrawingML_GraphicData_wps.docx",
    FIXTURES / "footnotes.docx",
    FIXTURES / "text-extraction" / "phase0" / "extraction-probe.docx",
]


def _edit_by_address(pkg: Any, markdown: str) -> int:
    """CR-011 section 14.4's edit: the first two paragraphs the Markdown addressed, edited by address."""
    from docx4j_py.model.content import AddressError, Paragraph

    edited = 0
    for address in ADDRESS_COMMENT.findall(markdown):
        try:
            paragraph = pkg.paragraph_at(address)
        except AddressError:
            continue
        if not isinstance(paragraph, Paragraph):
            continue
        paragraph.insert_text(" Edited.", location="End")
        edited += 1
        if edited == 2:
            break
    return edited


@pytest.mark.parametrize("path", WORKFLOW, ids=lambda p: p.name)
def test_markdown_then_an_edit_by_address_saves_as_after_a_fresh_parse(path: Path) -> None:
    adopted = load(path)
    adopted.id_seed = SEED
    markdown = adopted.to_markdown(addresses=True)
    main = adopted.main_document_part
    private = main._private_contents
    assert private is not None and not main.is_unmarshalled
    assert _edit_by_address(adopted, markdown) >= 1
    assert main.contents is private, "the body adopted the tree the Markdown read"

    fresh = load(path)
    fresh.id_seed = SEED
    fresh.main_document_part.contents  # noqa: B018 - parsed first: nothing to adopt
    assert fresh.to_markdown(addresses=True) == markdown
    _edit_by_address(fresh, markdown)

    assert entries(adopted.save()) == entries(fresh.save())


def test_markdown_then_the_body_then_a_save_with_no_edit_is_the_fresh_parse_save() -> None:
    adopted = load(SAMPLE)
    adopted.to_markdown()
    adopted.body  # noqa: B018 - unmarshals the main part, by adoption
    fresh = load(SAMPLE)
    fresh.body  # noqa: B018
    assert entries(adopted.save()) == entries(fresh.save())


def _style_edit(pkg: Any) -> None:
    from docx4j_py.wml import RPr, el

    styles = pkg.style_definitions_part.contents
    normal = next(style for style in styles.style if style.style_id == "Normal")
    if normal.r_pr is None:
        normal.r_pr = RPr()
    normal.r_pr.b = el.b()
    pkg.get_property_resolver().refresh()


def test_a_resolver_read_then_a_style_edit_saves_as_after_a_fresh_parse() -> None:
    adopted = load(SAMPLE)
    resolver = adopted.get_property_resolver()
    before = resolver.get_effective_rpr("Normal")
    assert before.b is None
    private = adopted.style_definitions_part._private_contents
    assert private is not None
    _style_edit(adopted)
    assert adopted.style_definitions_part.contents is private
    assert resolver.get_effective_rpr("Normal").b is not None, "refresh() sees the edit"

    fresh = load(SAMPLE)
    fresh.style_definitions_part.contents  # noqa: B018 - parsed first
    fresh.get_property_resolver()
    _style_edit(fresh)

    assert entries(adopted.save()) == entries(fresh.save())


def test_a_resolver_read_then_a_content_api_style_saves_as_after_a_fresh_parse() -> None:
    """``Paragraph.style`` adds a style the document lacks, through ``ensure_style``."""

    def edit(pkg: Any) -> None:
        pkg.body.paragraphs[0].style = "Heading 2"

    adopted = load(SAMPLE)
    adopted.id_seed = SEED
    adopted.to_markdown()
    assert adopted.style_definitions_part._private_contents is not None
    edit(adopted)
    assert adopted.style_definitions_part.is_unmarshalled

    fresh = load(SAMPLE)
    fresh.id_seed = SEED
    fresh.style_definitions_part.contents  # noqa: B018 - parsed first
    fresh.main_document_part.contents  # noqa: B018
    edit(fresh)

    assert entries(adopted.save()) == entries(fresh.save())


def _every_private_reader(pkg: Any) -> None:
    pkg.to_markdown(addresses=True)
    pkg.get_property_resolver()
    pkg.get_run_font_selector()
    pkg.get_font_mapper()


def _private_then_unmarshal_all(path: Path) -> dict[str, bytes]:
    pkg = load(path)
    _every_private_reader(pkg)
    pkg.unmarshal_all()
    return entries(pkg.save())


def _unmarshal_all(path: Path) -> dict[str, bytes]:
    pkg = load(path)
    pkg.unmarshal_all()
    return entries(pkg.save())


@pytest.mark.parametrize("path", [SAMPLE, SAMPLES / "Headers.docx"], ids=lambda p: p.name)
def test_every_part_adopted_saves_as_every_part_parsed(path: Path) -> None:
    assert _private_then_unmarshal_all(path) == _unmarshal_all(path)


@pytest.mark.slow
@pytest.mark.parametrize("path", sorted(SAMPLES.glob("*.docx")), ids=lambda p: p.name)
def test_every_part_adopted_saves_as_every_part_parsed_over_the_samples(path: Path) -> None:
    assert _private_then_unmarshal_all(path) == _unmarshal_all(path)


# ---------------------------------------------------------------------------
# test 4: changed load options block adoption
# ---------------------------------------------------------------------------


def _mce_tree(path: Path) -> bytes:
    part = load(path, options=LoadOptions(mce_preprocess=True)).main_document_part
    return part._marshal(part.contents)


def test_mce_preprocess_turned_on_after_a_private_read_parses_afresh(monkeypatch: pytest.MonkeyPatch) -> None:
    pkg = load(WITH_MCE)
    part = pkg.main_document_part
    private = part.read_contents()
    pkg.load_options = LoadOptions(mce_preprocess=True)
    parses = Parses(monkeypatch, "/word/document.xml")
    live = part.contents
    assert live is not private
    assert parses.counts == (0, 1)
    assert part._private_contents is None
    assert part._marshal(live) == _mce_tree(WITH_MCE)
    assert part._marshal(live) != part._marshal(private), "the two options give two trees here"


def test_an_option_changed_in_place_blocks_adoption() -> None:
    pkg = load(WITH_MCE)
    part = pkg.main_document_part
    private = part.read_contents()
    pkg.load_options.mce_preprocess = True
    assert part.contents is not private


def test_mce_preprocess_turned_off_after_a_private_read_parses_afresh() -> None:
    pkg = load(WITH_MCE, options=LoadOptions(mce_preprocess=True))
    part = pkg.main_document_part
    private = part.read_contents()
    pkg.load_options = LoadOptions()
    live = part.contents
    assert live is not private
    fresh = load(WITH_MCE).main_document_part
    assert part._marshal(live) == fresh._marshal(fresh.contents)


def test_options_replaced_by_equal_ones_still_adopt() -> None:
    pkg = load(WITH_MCE, options=LoadOptions(mce_preprocess=True))
    part = pkg.main_document_part
    private = part.read_contents()
    pkg.load_options = LoadOptions(mce_preprocess=True)
    assert part.contents is private


def test_a_part_read_with_no_package_is_not_adopted_under_a_package_that_differs() -> None:
    from docx4j_py.openpackaging.parts.wml import MainDocumentPart

    part = MainDocumentPart()
    part.set_bytes(load(SAMPLE).main_document_part.xml)
    private = part.read_contents()

    class Strict:
        load_options = LoadOptions(strict=True)

    part.package = Strict()  # type: ignore[assignment]
    assert part.contents is not private


# ---------------------------------------------------------------------------
# test 5: the thread rule is unchanged
# ---------------------------------------------------------------------------


def _job(path: Path) -> dict[str, bytes]:
    pkg = load(path)
    pkg.id_seed = SEED
    markdown = pkg.to_markdown(addresses=True)
    pkg.get_property_resolver()
    _edit_by_address(pkg, markdown)
    pkg.unmarshal_all()
    return entries(pkg.save())


def test_one_package_per_thread_gives_the_bytes_of_a_sequential_run() -> None:
    """Section 16.3 test 5: read, adopt, edit and save on eight threads, against a sequential run."""
    paths = [SAMPLE, SAMPLES / "Headers.docx", SAMPLES / "DrawingML_GraphicData_wps.docx", FIXTURES / "footnotes.docx"]
    expected = {path.name: _job(path) for path in paths}
    with ThreadPoolExecutor(max_workers=8) as pool:
        results = list(pool.map(lambda path: (path.name, _job(path)), paths * 4))
    for name, result in results:
        assert result == expected[name], name
