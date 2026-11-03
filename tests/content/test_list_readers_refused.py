"""The content API's list readers refuse a numbering or styles part that is not well formed. CR-031 section 4.6, Phase A2.

Section 4.6 as amended on 2026-11-03 (the user: "make the list readers refuse"; section 10.8 item 1). The list
emulator read both parts with lxml and read one that would not parse as absent, so that ``is_list_item`` was False
for a numbered paragraph with nothing said. Now every reader over the emulator raises the part's own refusal: the
``Docx4JException`` its ``contents`` raises ("Problem with part /word/numbering.xml: <lxml's message>"), lxml's
``XMLSyntaxError`` the deepest cause, kept on the part as its ``unreadable`` (D8), so that a later read raises it
again without a parse.

- **The styles part** raises at the emulator's first read of the styles, so at every list reader.
- **The numbering part** raises only where a definition is needed: over a paragraph that names a list (a
  ``w:numId`` other than 0, its own or its style's), the paragraphs where the text model raises ``part.not_read``.
  Over a paragraph that names none, nothing raises.
- **The intact and absent twins** read as before, and a well-formed part pays nothing: no parse, no re-marshal.
- **The text model is unchanged** (``tests/text/test_part_not_read.py`` holds it to docx4j's fixture): it hands the
  emulator neither part, the emulator built without the styles and a paragraph resolved without the numbering.
- **The property resolver's readers** over a paragraph that names a list raise the numbering part's refusal too,
  through the level indent its paragraph merge asks the emulator for (section 11).

The fixture is docx4j's ``part-not-read`` (``tests/fixtures/part-not-read/``): ten paragraphs, two ``Heading1``
numbered by their style ("Article I"), two numbered directly ("1."), six that name no list.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from docx4j_py import load
from docx4j_py.model.content import List, ListItem
from docx4j_py.model.listnumbering.emulator import Emulator
from docx4j_py.model.text import extract, plain_text
from docx4j_py.model.text.trees import without_numbering
from docx4j_py.openpackaging import XmlPart
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.parts.xml_part import syntax_error_of

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "part-not-read"

#: The paragraphs that name a list: two headings by their style, two directly.
NAMED = (0, 1, 2, 3)
#: The paragraphs that name none.
UNNAMED = (4, 5, 6, 7, 8, 9)
#: The paragraphs numbered by a ``w:numPr`` of their own.
DIRECT = (2, 3)

#: The labels each fixture reads, by paragraph: what the intact and absent twins give, before and after.
LABELS = {
    "intact": ["Article I", "Article II", "1.", "2.", None, None, None, None, None, None],
    "styles-absent": [None, None, "1.", "2.", None, None, None, None, None, None],
    "numbering-absent": [None] * 10,
    "settings": ["Article I", "Article II", "1.", "2.", None, None, None, None, None, None],
}

#: The two broken fixtures: the part not well formed, and lxml's message for it.
BROKEN = {
    "styles": ("/word/styles.xml", "AttValue: ' expected, line 2, column 625"),
    "numbering": ("/word/numbering.xml", "Specification mandates value for attribute w, line 2, column 395"),
}


def _load(name: str) -> Any:
    return load(FIXTURES / f"part-not-read-{name}.docx")


def _broken(pkg: Any, name: str) -> Any:
    return pkg.style_definitions_part if name == "styles" else pkg.numbering_definitions_part


def _paragraph(pkg: Any, index: int) -> Any:
    return pkg.body.paragraphs[index]


#: Every reader over one paragraph, each through the emulator.
PARAGRAPH_READERS: dict[str, Callable[[Any], Any]] = {
    "is_list_item": lambda p: p.is_list_item,
    "list_item": lambda p: p.list_item,
    "list": lambda p: p.list,
    "to_dict": lambda p: p.to_dict().get("list_item"),
    "ListItem.list_string": lambda p: ListItem(p).list_string,
    "ListItem.level": lambda p: ListItem(p).level,
    "ListItem.sibling_index": lambda p: ListItem(p).sibling_index,
    "ListItem.get_ancestor": lambda p: ListItem(p).get_ancestor(),
    "ListItem.get_descendants": lambda p: ListItem(p).get_descendants(),
    "ListItem.to_dict": lambda p: ListItem(p).to_dict(),
    "Emulator.get_number": lambda p: Emulator.get_number(p.parent_body.package, p.element.p_pr),
    "Emulator.peek": lambda p: Emulator.peek(p.parent_body.package, p.element.p_pr),
}

#: Every reader over the body or a list, each needing a definition.
BODY_READERS: dict[str, Callable[[Any], Any]] = {
    "body.lists": lambda pkg: pkg.body.lists,
    "body.to_api_script": lambda pkg: pkg.body.to_api_script(),
    "List.definition": lambda pkg: List(1, pkg.body).definition,
    "List.level_types": lambda pkg: List(1, pkg.body).level_types,
    "List.paragraphs": lambda pkg: List(1, pkg.body).paragraphs,
    "List.level_exists": lambda pkg: List(1, pkg.body).level_exists(0),
    "List.get_level_paragraphs": lambda pkg: List(1, pkg.body).get_level_paragraphs(0),
    "List.get_level_string": lambda pkg: List(1, pkg.body).get_level_string(0),
    "List.to_dict": lambda pkg: List(1, pkg.body).to_dict(),
    "get_emulator().instance_list_definitions": lambda pkg: (
        pkg.numbering_definitions_part.get_emulator().instance_list_definitions
    ),
    "get_emulator().abstract_list_definitions": lambda pkg: (
        pkg.numbering_definitions_part.get_emulator().abstract_list_definitions
    ),
    "attach_to_list": lambda pkg: _paragraph(pkg, 4).attach_to_list(1),
    "detach_from_list": lambda pkg: _paragraph(pkg, 2).detach_from_list(),
}


def _refused(name: str, read: Callable[[], Any], pkg: Any) -> Docx4JException:
    """`read` raises the broken part's own refusal: its record, naming it, lxml's message, lxml's error the deepest cause."""
    part_name, message = BROKEN[name]
    with pytest.raises(Docx4JException) as raised:
        read()
    part = _broken(pkg, name)
    assert raised.value is part.unreadable
    assert str(raised.value) == f"Problem with part {part_name}: {message}"
    assert syntax_error_of(raised.value) is not None
    return raised.value


# ---------------------------------------------------------------------------
# the styles part: every list reader
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("index", range(10))
@pytest.mark.parametrize("reader", list(PARAGRAPH_READERS))
def test_a_styles_part_not_well_formed_refuses_every_paragraph_reader(reader: str, index: int) -> None:
    pkg = _load("styles")
    paragraph = _paragraph(pkg, index)
    _refused("styles", lambda: PARAGRAPH_READERS[reader](paragraph), pkg)


@pytest.mark.parametrize("name", ["styles", "numbering"])
def test_to_api_script_refuses_a_paragraph_numbered_by_its_own_numpr(name: str) -> None:
    """A script expresses a ``w:numPr`` of the paragraph's own through the list verbs, so it reads the list there;
    a heading numbered by its style it expresses by the style, and every other paragraph as before."""
    absent = _load(f"{name}-absent")
    for index in range(10):
        pkg = _load(name)
        if index in DIRECT:
            _refused(name, lambda pkg=pkg, index=index: _paragraph(pkg, index).to_api_script(), pkg)
        else:
            assert _paragraph(pkg, index).to_api_script() == _paragraph(absent, index).to_api_script()


@pytest.mark.parametrize("reader", list(BODY_READERS))
def test_a_styles_part_not_well_formed_refuses_every_body_and_list_reader(reader: str) -> None:
    pkg = _load("styles")
    if reader == "body.to_api_script":
        with pytest.raises(Docx4JException):
            BODY_READERS[reader](pkg)
        return
    _refused("styles", lambda: BODY_READERS[reader](pkg), pkg)


def test_the_styles_part_is_refused_at_the_emulator_s_first_read_of_the_styles() -> None:
    """``Emulator.styles`` raises, and a second read raises the same record without a parse."""
    pkg = _load("styles")
    emulator = pkg.numbering_definitions_part.get_emulator()
    first = _refused("styles", lambda: emulator.styles, pkg)
    calls = _count_parses(pkg.style_definitions_part)
    with pytest.raises(Docx4JException) as again:
        emulator.styles  # noqa: B018
    assert again.value is first
    assert calls() == 0


# ---------------------------------------------------------------------------
# the numbering part: where a definition is needed, and nowhere else
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("index", NAMED)
@pytest.mark.parametrize("reader", list(PARAGRAPH_READERS))
def test_a_numbering_part_not_well_formed_refuses_a_reader_over_a_paragraph_that_names_a_list(
    reader: str, index: int
) -> None:
    pkg = _load("numbering")
    paragraph = _paragraph(pkg, index)
    _refused("numbering", lambda: PARAGRAPH_READERS[reader](paragraph), pkg)


@pytest.mark.parametrize("index", UNNAMED)
@pytest.mark.parametrize("reader", [r for r in PARAGRAPH_READERS if r != "to_api_script"])
def test_a_numbering_part_not_well_formed_refuses_nothing_over_a_paragraph_that_names_none(
    reader: str, index: int
) -> None:
    """No definition is needed, so nothing is read and nothing raises: the reader reads as over the absent part."""
    broken = _load("numbering")
    absent = _load("numbering-absent")
    value = PARAGRAPH_READERS[reader](_paragraph(broken, index))
    assert _plain(value) == _plain(PARAGRAPH_READERS[reader](_paragraph(absent, index)))
    part = broken.numbering_definitions_part
    assert part.unreadable is None
    assert not part.is_unmarshalled


@pytest.mark.parametrize("reader", list(BODY_READERS))
def test_a_numbering_part_not_well_formed_refuses_every_body_and_list_reader(reader: str) -> None:
    pkg = _load("numbering")
    _refused("numbering", lambda: BODY_READERS[reader](pkg), pkg)


def test_a_numbering_part_no_paragraph_needs_refuses_no_reader() -> None:
    """``part-not-read-numbering-unneeded.docx``: two plain paragraphs; every paragraph and body reader reads."""
    pkg = _load("numbering-unneeded")
    for paragraph in pkg.body.paragraphs:
        assert paragraph.is_list_item is False
        assert paragraph.list_item is None
        assert paragraph.list is None
        assert "list_item" not in paragraph.to_dict()
        assert ListItem(paragraph).list_string == ""
    assert pkg.body.to_api_script()
    assert pkg.body.lists == []  # no paragraph names a list, so no definition is needed
    part = pkg.numbering_definitions_part
    assert part.unreadable is None
    assert not part.is_unmarshalled
    # the definitions themselves are refused
    with pytest.raises(Docx4JException) as raised:
        part.get_emulator().instance_list_definitions  # noqa: B018
    assert raised.value is part.unreadable


def test_a_numbering_part_s_refusal_is_raised_again_without_a_parse() -> None:
    pkg = _load("numbering")
    first = _refused("numbering", lambda: _paragraph(pkg, 2).is_list_item, pkg)
    calls = _count_parses(pkg.numbering_definitions_part)
    for index in NAMED:
        with pytest.raises(Docx4JException) as again:
            _paragraph(pkg, index).list_item  # noqa: B018
        assert again.value is first
    with pytest.raises(Docx4JException) as again:
        pkg.body.lists  # noqa: B018
    assert again.value is first
    assert calls() == 0
    assert pkg.numbering_definitions_part.is_content_unreadable


def test_the_record_made_by_another_read_is_the_one_the_list_readers_raise() -> None:
    """A read of ``contents`` first records the refusal; the emulator raises that record, parsing nothing."""
    pkg = _load("numbering")
    part = pkg.numbering_definitions_part
    with pytest.raises(Docx4JException) as raised:
        part.contents  # noqa: B018
    calls = _count_parses(part)
    with pytest.raises(Docx4JException) as again:
        _paragraph(pkg, 2).is_list_item  # noqa: B018
    assert again.value is raised.value
    assert calls() == 0


def test_a_replacement_reads_again() -> None:
    """``set_bytes`` clears the record: the emulator, refreshed, reads the replacement."""
    pkg = _load("numbering")
    _refused("numbering", lambda: _paragraph(pkg, 2).is_list_item, pkg)
    intact = _load("intact").numbering_definitions_part.xml
    pkg.numbering_definitions_part.set_bytes(intact)
    pkg.numbering_definitions_part.get_emulator(reset=True)
    assert [p.list_item.list_string if p.list_item else None for p in pkg.body.paragraphs] == LABELS["intact"]


# ---------------------------------------------------------------------------
# the property resolver's readers over a paragraph that names a list (section 11)
# ---------------------------------------------------------------------------

RESOLVER_READERS: dict[str, Callable[[Any], Any]] = {
    "effective_p_pr": lambda p: p.effective_p_pr,
    "left_indent": lambda p: p.left_indent,
    "first_line_indent": lambda p: p.first_line_indent,
    "outline_level": lambda p: p.outline_level,
    "alignment": lambda p: p.alignment,
}


@pytest.mark.parametrize("reader", list(RESOLVER_READERS))
def test_the_resolver_s_readers_refuse_a_paragraph_that_names_a_list(reader: str) -> None:
    """The paragraph merge asks the emulator for the level's indent, so the numbering part's refusal reaches them."""
    for index in NAMED:
        pkg = _load("numbering")
        _refused("numbering", lambda pkg=pkg, index=index: RESOLVER_READERS[reader](_paragraph(pkg, index)), pkg)
    pkg = _load("numbering")
    absent = _load("numbering-absent")
    for index in UNNAMED:
        assert _plain(RESOLVER_READERS[reader](_paragraph(pkg, index))) == _plain(
            RESOLVER_READERS[reader](_paragraph(absent, index))
        )


def test_paragraph_text_reads_with_a_numbering_part_not_well_formed() -> None:
    """``text`` resolves run properties alone, which never ask the numbering (Phase A's case 16 too)."""
    pkg = _load("numbering")
    assert pkg.body.text == _load("numbering-absent").body.text


# ---------------------------------------------------------------------------
# the twins, and what a well-formed part pays
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", list(LABELS))
def test_the_intact_and_absent_twins_read_as_before(name: str) -> None:
    pkg = _load(name)
    labels = LABELS[name]
    paragraphs = pkg.body.paragraphs
    assert [p.is_list_item for p in paragraphs] == [label is not None for label in labels]
    assert [None if p.list_item is None else p.list_item.list_string for p in paragraphs] == labels
    assert [None if p.list is None else p.list.id for p in paragraphs] == [
        None if label is None else (2 if label.startswith("Article") else 1) for label in labels
    ]
    expected = sorted({2 if label.startswith("Article") else 1 for label in labels if label is not None})
    assert sorted(item.id for item in pkg.body.lists) == expected


@pytest.mark.parametrize("name", ["intact", "styles-absent", "settings"])
def test_a_well_formed_numbering_part_is_read_once_with_lxml_and_never_typed(name: str) -> None:
    """No typed parse of the numbering part, no re-marshal: the list readers read its bytes, as before."""
    pkg = _load(name)
    part = pkg.numbering_definitions_part
    calls = _count_parses(part)
    for paragraph in pkg.body.paragraphs:
        paragraph.is_list_item  # noqa: B018
        paragraph.list_item  # noqa: B018
    pkg.body.lists  # noqa: B018
    assert calls() == 0
    assert not part.is_unmarshalled
    assert part._private_contents is None
    assert part.unreadable is None


# ---------------------------------------------------------------------------
# the text model reads around both parts, in either order
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", ["styles", "numbering"])
def test_the_text_model_reads_as_before_after_and_before_a_list_reader_s_refusal(name: str) -> None:
    """The refusal recorded by a list reader is the text model's ``part.not_read``; the reading is its absent twin's."""
    absent = plain_text(extract(_load(f"{name}-absent"), list_labels=True))
    pkg = _load(name)
    with pytest.raises(Docx4JException) as raised:
        pkg.body.lists  # noqa: B018
    result = extract(pkg, list_labels=True)
    assert plain_text(result) == absent
    assert [str(w) for w in result.warnings] == [f"part.not_read ({BROKEN[name][0]})"]
    pkg = _load(name)
    result = extract(pkg, list_labels=True)
    assert plain_text(result) == absent
    with pytest.raises(Docx4JException) as again:
        pkg.body.lists  # noqa: B018
    assert again.value is _broken(pkg, name).unreadable
    assert str(again.value) == str(raised.value)


def test_the_text_model_s_emulator_reads_no_styles_part_it_could_not_read() -> None:
    """``Emulator(package, read_styles=False)``: no style map, no resolver, the styles part never read."""
    pkg = _load("styles")
    emulator = Emulator(pkg, read_styles=False)
    assert emulator.styles == {}
    assert emulator.resolver is None
    assert emulator.default_paragraph_style_id is None
    assert sorted(emulator.instance_list_definitions) == ["1", "2"]
    assert pkg.style_definitions_part.unreadable is None


def test_without_numbering_answers_only_the_numbering_part_s_own_refusal() -> None:
    pkg = _load("numbering")
    resolver = pkg.get_property_resolver()
    error = _refused("numbering", lambda: _paragraph(pkg, 0).outline_level, pkg)
    fallback = without_numbering(resolver, error)
    assert fallback is not None and fallback is not resolver
    assert fallback.numbering_indents is None
    assert _paragraph(pkg, 0).element.p_pr is not None
    effective = fallback.get_effective_ppr(_paragraph(pkg, 0).element.p_pr)
    assert effective.outline_lvl.val == 0  # the heading's own, read as without the numbering
    assert without_numbering(resolver, Docx4JException(str(error))) is None
    assert without_numbering(resolver, ValueError("other")) is None
    intact = _load("intact")
    assert without_numbering(intact.get_property_resolver(), error) is None


def test_a_dry_run_reads_a_numbering_part_not_well_formed_as_the_package_does() -> None:
    """A trial cannot copy the part, so its resolver's numbering seam raises the part's refusal; resolved as without
    the part, the trial reads as the package. Before Phase A2 its headings numbered by their style lost their level."""
    pkg = _load("numbering")
    with pkg.dry_run() as trial:
        markdown_of_trial = trial.to_markdown()
    assert markdown_of_trial == _load("numbering").to_markdown() == _load("numbering-absent").to_markdown()
    assert markdown_of_trial.startswith("# Definitions\n\n# Interpretation\n")


def test_the_text_model_reads_a_numbered_quote_style_as_without_the_numbering() -> None:
    """The Markdown renderer's quote depth asks the resolver for the quote style's own indent (CR-031 section 4.6)."""
    from docx4j_py.model.text import markdown

    def quoted(numbering_cut: bool) -> str:
        pkg = _load("numbering" if numbering_cut else "numbering-absent")
        styles = pkg.style_definitions_part.contents
        from docx4j_py.fragments import wml

        styles.style.append(
            wml(
                '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>'
                '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="864"/></w:pPr>'
                "</w:style>"
            )
        )
        pkg.refresh_property_resolver()
        paragraph = pkg.main_document_part.contents.body.content[4]
        paragraph.p_pr = wml(
            '<w:pPr><w:pStyle w:val="Quote"/><w:numPr><w:numId w:val="0"/></w:numPr><w:ind w:left="1728"/></w:pPr>'
        )
        paragraph.p_pr.parent = paragraph
        return markdown(extract(pkg), pkg)

    assert quoted(True) == quoted(False)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _count_parses(part: Any) -> Callable[[], int]:
    """How many times `part` is parsed from here on (``XmlPart._parse``)."""
    calls: list[int] = []
    original = XmlPart._parse

    def counting(self: Any, *args: Any, **kwargs: Any) -> Any:
        if self is part:
            calls.append(1)
        return original(self, *args, **kwargs)

    XmlPart._parse = counting  # type: ignore[method-assign]
    _RESTORE.append(lambda: setattr(XmlPart, "_parse", original))
    return lambda: len(calls)


_RESTORE: list[Callable[[], None]] = []


@pytest.fixture(autouse=True)
def _restore_parse() -> Any:
    yield
    while _RESTORE:
        _RESTORE.pop()()


def _plain(value: Any) -> Any:
    """A reader's value, comparable across two packages: a view by what it shows."""
    if isinstance(value, ListItem):
        return ("ListItem", value.to_dict())
    if isinstance(value, List):
        return ("List", value.id)
    if hasattr(value, "text") and hasattr(value, "element"):
        return ("Paragraph", value.text)
    if isinstance(value, list):
        return [_plain(item) for item in value]
    if value is not None and hasattr(value, "__dataclass_fields__"):
        from docx4j_py.fragments import to_xml

        return to_xml(value)
    return value
