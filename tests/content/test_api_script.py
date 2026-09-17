"""CR-003 Phase I, section 3.11: ``to_api_script``, the reveal-codes generator.

The promise the section makes is that the script *reproduces* what it was shown:
the verbs where they exist, ``insert_xml`` of the marshalled fragment where they
do not, and the decision taken per block. So the test of section 7 is the one
that matters --- **the script executed against a fresh body, compared to the
source** --- and the rest of this file pins the individual rules.

What the comparison ignores, and why (CR-003 section 19.4). None of these is a
difference Word paints; every one is something the verbs do not write and the
normal form of ``scripts/canon.py`` would otherwise report:

* the revision-save ids (``w:rsid*``), ``w14:paraId`` and ``w14:textId``: ids of
  the *source* document, which section 3.11's determinism rule keeps out of the
  script;
* ``w:lang`` and ``w:noProof``: language and proofing hints, dropped by the
  generator as ``w:proofErr`` is;
* ``w:bCs``, ``w:iCs`` and ``w:szCs``: the complex-script twins that CR-003
  section 3.12's one mapping always writes beside their Latin twins;
* an empty ``w:rPr`` / ``w:pPr`` / ``w:tcPr``, and a run left with only a
  ``w:rPr``: no direct formatting, and nothing painted;
* ``w:numPr``: the ``w:numId`` is the *target* document's own allocation
  (section 18.7), and the labels are compared instead;
* a redundant ``xml:space="preserve"``;
* two adjacent runs with the same ``w:rPr`` written as one, which is what
  ``insert_text`` does and what Word paints either way.
"""

from __future__ import annotations

import ast
import sys

import canon
import pytest
from conftest import FIXTURES, SAMPLES, fixture, sample, threaded_package
from lxml import etree

from docx4j_py import create_package, load
from docx4j_py.fragments import to_xml
from docx4j_py.model.content.api_script import (
    PLACEHOLDER_PNG,
    script_blocks,
    to_api_script,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

#: Elements the comparison drops; see the module docstring.
_DROPPED = {
    W + "proofErr",
    W + "bookmarkStart",
    W + "bookmarkEnd",
    W + "lastRenderedPageBreak",
    W + "commentRangeStart",
    W + "commentRangeEnd",
    W + "commentReference",
    W + "annotationRef",
    W + "bCs",
    W + "iCs",
    W + "szCs",
    W + "numPr",
    W + "lang",
    W + "noProof",
}
#: Property elements that say nothing once they are empty.
_EMPTY_OK = {W + "rPr", W + "pPr", W + "tcPr", W + "trPr", W + "tblPr"}

#: The documents section 7 names for the executed round trip.
ROUND_TRIP = [
    "2010-sample1.docx",
    "sample-docx.docx",
    "lists.docx",
    "comments-modern.docx",
    "hyperlink.docx",
    "nested-table.docx",
]


def _document(name: str):
    """One of ``samples/`` or ``tests/fixtures/``, seeded."""
    package = sample(name) if (SAMPLES / name).exists() else fixture(name)
    package.id_seed = 20260917
    return package


# ---------------------------------------------------------------------------
# the normal form the executed round trip compares with
# ---------------------------------------------------------------------------


def _strip(element) -> None:
    for name in list(element.attrib):
        local = name.rpartition("}")[2]
        if local.startswith("rsid") or name in (W14 + "paraId", W14 + "textId"):
            del element.attrib[name]
    for child in list(element):
        if not isinstance(child.tag, str) or child.tag in _DROPPED:
            element.remove(child)
            continue
        _strip(child)
    for child in list(element):
        if child.tag in _EMPTY_OK and len(child) == 0 and not child.attrib:
            element.remove(child)


def _r_pr(run) -> bytes:
    found = run.find(W + "rPr")
    return b"" if found is None else etree.tostring(found)


def _only_text(run) -> bool:
    return all(child.tag in (W + "t", W + "rPr") for child in run)


def _merge(element) -> None:
    for parent in element.iter():
        children = list(parent)
        index = 0
        while index + 1 < len(children):
            one, two = children[index], children[index + 1]
            if (
                one.tag == W + "r"
                and two.tag == W + "r"
                and _only_text(one)
                and _only_text(two)
                and _r_pr(one) == _r_pr(two)
            ):
                for item in list(two):
                    if item.tag == W + "t":
                        one.append(item)
                parent.remove(two)
                del children[index + 1]
                continue
            index += 1
    for run in element.iter(W + "r"):
        texts = [child for child in run if child.tag == W + "t"]
        if len(texts) > 1:
            texts[0].text = "".join(text.text or "" for text in texts)
            for extra in texts[1:]:
                run.remove(extra)
    for parent in element.iter():
        for run in list(parent):
            if run.tag == W + "r" and all(child.tag == W + "rPr" for child in run):
                parent.remove(run)


def _space(element) -> None:
    for text in element.iter(W + "t"):
        value = text.text or ""
        if value == value.strip() and "  " not in value:
            text.attrib.pop(XML_SPACE, None)
        else:
            text.set(XML_SPACE, "preserve")


def normal_form(xml: str) -> bytes:
    """``scripts/canon.py``'s normal form, less what the module docstring lists."""
    root = etree.fromstring(xml.encode("utf-8"))
    _strip(root)
    _merge(root)
    _space(root)
    return canon.normalize(etree.tostring(root))


def executed(package, **options) -> tuple:
    """The script of a package's body, and the body it builds in a fresh one."""
    script = package.body.to_api_script(**options)
    target = create_package()
    target.id_seed = 20260917
    exec(compile(script, "<to_api_script>", "exec"), {"body": target.body})  # noqa: S102
    return script, target


# ---------------------------------------------------------------------------
# the executed round trip (CR-003 section 7)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", ROUND_TRIP)
def test_the_script_reproduces_every_block_it_claims(name: str) -> None:
    """Every block the verbs expressed is the source block, canonically."""
    package = _document(name)
    blocks = [block for block in script_blocks(package.body) if block.kind != "note"]
    _script, target = executed(package)
    produced = list(target.body.content)
    assert len(produced) == len(blocks), "one produced block per source block"
    for block, made in zip(blocks, produced, strict=True):
        if not block.exact:
            continue
        assert normal_form(to_xml(block.element)) == normal_form(to_xml(made)), (
            f"{name} {block.address} ({block.kind}): {block.lines}"
        )


@pytest.mark.parametrize("name", ROUND_TRIP)
def test_the_whole_document_script_keeps_the_text_and_the_count(name: str) -> None:
    """A whole-document script gives the same text and the same paragraph count."""
    package = _document(name)
    _script, target = executed(package)
    assert target.body.get_text() == package.body.get_text()
    assert len(target.body.paragraphs) == len(package.body.paragraphs)


def test_an_insert_xml_block_is_the_marshalled_element() -> None:
    """The fallback carries the element itself, so it comes back canonically."""
    package = _document("nested-table.docx")
    blocks = [block for block in script_blocks(package.body) if block.kind == "xml"]
    assert blocks
    _script, target = executed(package)
    for block, made in zip(blocks, target.body.content, strict=False):
        assert normal_form(to_xml(block.element)) == normal_form(to_xml(made))


# ---------------------------------------------------------------------------
# the paragraph rules
# ---------------------------------------------------------------------------


def test_a_plain_paragraph_is_one_line(body) -> None:
    """No properties, one unformatted run: one statement and no variable."""
    body.insert_paragraph("Hello World")
    assert to_api_script(body.paragraphs[0]) == 'body.insert_paragraph("Hello World")'


def test_a_styled_paragraph_sets_its_members_in_order(body) -> None:
    """The style first, then alignment, the indents and the spacing, in points."""
    paragraph = body.insert_paragraph("Chapter 1", style="Heading 1")
    paragraph.alignment = "Centered"
    paragraph.left_indent = 18
    paragraph.space_after = 6
    paragraph.outline_level = 2
    assert to_api_script(paragraph).splitlines() == [
        'p1 = body.insert_paragraph("Chapter 1")',
        'p1.style_built_in = "Heading1"',
        'p1.alignment = "Centered"',
        "p1.left_indent = 18",
        "p1.space_after = 6",
        "p1.outline_level = 2",
    ]


def test_a_style_that_is_not_built_in_is_written_by_its_id(body) -> None:
    """``style`` takes a display name the target may not define; ``style_id`` does not."""
    paragraph = body.insert_paragraph("Note")
    paragraph.style_id = "MyOwnStyle"
    assert 'p1.style_id = "MyOwnStyle"' in to_api_script(paragraph)


def test_each_run_emits_the_font_that_differs_from_the_previous(body) -> None:
    """One ``insert_text`` per run, and only the ``Font`` members that changed."""
    body.insert_xml(
        "<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>bold</w:t></w:r>"
        '<w:r><w:t xml:space="preserve"> plain </w:t></w:r>'
        '<w:r><w:rPr><w:i/><w:iCs/><w:color w:val="FF0000"/></w:rPr><w:t>red</w:t></w:r></w:p>'
    )
    assert to_api_script(body.paragraphs[0]).splitlines() == [
        'p1 = body.insert_paragraph("")',
        'r1 = p1.insert_text("bold")',
        "r1.font.bold = True",
        'r2 = p1.insert_text(" plain ")',
        "r2.font.bold = False",
        'r3 = p1.insert_text("red")',
        "r3.font.italic = True",
        'r3.font.color = "#FF0000"',
    ]


def test_adjacent_runs_of_the_same_formatting_are_one_call(body) -> None:
    """Word paints them as one run, so the script writes one ``insert_text``."""
    body.insert_xml(
        '<w:p><w:r w:rsidRPr="0034500C"><w:rPr><w:lang w:val="en-AU"/></w:rPr>'
        '<w:t xml:space="preserve">one </w:t></w:r>'
        '<w:r w:rsidRPr="00AD6B9B"><w:rPr><w:lang w:val="en-AU"/></w:rPr><w:t>two</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t> three</w:t></w:r></w:p>'
    )
    assert to_api_script(body.paragraphs[0]).splitlines() == [
        'p1 = body.insert_paragraph("one two")',
        'r1 = p1.insert_text(" three")',
        "r1.font.bold = True",
    ]


def test_the_lists_fixture_item_is_one_call_and_still_exact() -> None:
    """The ten alternating one-word calls that used to come out of this are one.

    Its runs differ only by ``w:rsidRPr`` and all carry the same
    ``w:lang``, which the generator drops (CR-003 section 19.3 items 3 and 16).
    """
    package = _document("lists.docx")
    paragraph = package.body.paragraphs[2]
    block = script_blocks(paragraph)[0]
    assert block.exact
    assert [line for line in block.lines if "insert_text" in line] == []
    assert block.lines[0].startswith('p1 = body.insert_paragraph("None   Some content, .  ')
    assert block.lines[0].count("Some content,") == 8


def test_a_tab_run_is_not_coalesced_into_its_neighbour(body) -> None:
    """A ``w:tab`` makes its run more than text, and two such runs stay two."""
    body.insert_xml("<w:p><w:r><w:t>a</w:t></w:r><w:r><w:tab/><w:t>b</w:t></w:r></w:p>")
    lines = to_api_script(body.paragraphs[0]).splitlines()
    assert lines[-1] == 'p1.insert_text("\\tb")'


def test_a_run_style_is_a_font_member_here(body) -> None:
    """``Font.style`` is ``w:rStyle``; the TypeScript engine has no such member."""
    body.insert_xml(
        '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>hm</w:t></w:r></w:p>'
    )
    assert 'font.style = "Emphasis"' in to_api_script(body.paragraphs[0])


def test_a_break_of_its_own_is_insert_break(body) -> None:
    """A ``w:br`` alone in a run is ``insert_break``; the type is Office JS's."""
    body.insert_xml('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>')
    lines = to_api_script(body.paragraphs[0]).splitlines()
    assert lines[-1] == 'p1.insert_break("Page")'


def test_a_tab_in_the_leading_run_is_a_tab_character(body) -> None:
    """``insert_paragraph`` goes through the ``r`` builder, which writes ``w:tab``."""
    body.insert_xml("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>")
    assert to_api_script(body.paragraphs[0]) == 'body.insert_paragraph("a\\tb")'


def test_bookmarks_and_proofing_marks_do_not_force_a_fallback(body) -> None:
    """Dropped, as docx4j-core-ts drops them: they are not content."""
    body.insert_xml(
        '<w:p><w:proofErr w:type="spellStart"/><w:bookmarkStart w:id="0" w:name="x"/>'
        "<w:r><w:t>plain</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>"
    )
    assert to_api_script(body.paragraphs[0]) == 'body.insert_paragraph("plain")'


def test_a_hyperlink_paragraph_is_one_insert_xml() -> None:
    """A run holder no verb makes: the whole paragraph falls back, with the reason."""
    package = _document("hyperlink.docx")
    blocks = script_blocks(package.body)
    found = [block for block in blocks if block.kind == "xml"]
    assert found and found[0].reason == "w:hyperlink in the paragraph"
    assert found[0].lines[0] == "# w:hyperlink in the paragraph: as XML"
    assert found[0].lines[1].startswith("body.insert_xml(")


def test_a_paragraph_of_plain_runs_never_falls_back(body) -> None:
    """The verbs cover it, so no ``insert_xml`` appears."""
    body.insert_paragraph("one")
    body.insert_paragraph("two")
    assert "insert_xml" not in to_api_script(body)


def test_a_tracked_change_is_one_insert_xml(new_package) -> None:
    """A revision is reproduced, ids and all, rather than silently accepted.

    Emitting the accepted text would drop the revision without saying so, and
    ``change_tracking_mode`` is never written by the script (CR-003 section 19.3).
    """
    body = new_package.body
    body.insert_xml(
        '<w:p><w:ins w:id="7" w:author="Ada" w:date="2026-09-17T09:00:00Z">'
        "<w:r><w:t>new</w:t></w:r></w:ins></w:p>"
    )
    block = script_blocks(body)[0]
    assert block.kind == "xml"
    assert block.reason == "w:ins in the paragraph"
    assert "change_tracking_mode" not in "\n".join(block.lines)


def test_the_paragraph_marks_own_run_properties_fall_back(body) -> None:
    """``Paragraph.font`` is over the runs, so the mark's ``w:rPr`` has no verb."""
    body.insert_xml("<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>")
    block = script_blocks(body)[0]
    assert block.kind == "xml"
    assert block.reason == "the paragraph mark's w:rPr (b)"


def test_a_language_hint_on_the_mark_does_not(body) -> None:
    """``w:lang`` and ``w:noProof`` are dropped wherever they are."""
    body.insert_xml(
        '<w:p><w:pPr><w:rPr><w:lang w:val="en-GB"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:noProof/></w:rPr><w:t>x</w:t></w:r></w:p>'
    )
    assert to_api_script(body) == 'body.insert_paragraph("x")'


def test_auto_line_spacing_falls_back(body) -> None:
    """``line_spacing`` writes ``w:lineRule="exact"``, so only an exact rule round-trips."""
    body.insert_xml(
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
        "<w:r><w:t>x</w:t></w:r></w:p>"
    )
    assert script_blocks(body)[0].reason == 'w:spacing w:lineRule="auto"'


# ---------------------------------------------------------------------------
# tables
# ---------------------------------------------------------------------------


def test_a_plain_grid_is_insert_table(body) -> None:
    """``insert_table`` with the values, then the style and the header rows."""
    table = body.insert_table(2, 2, values=[["Region", "Total"], ["North", "120"]])
    table.style_built_in = "TableGrid"
    table.header_row_count = 1
    lines = to_api_script(table).splitlines()
    assert lines[0] == (
        't1 = body.insert_table(2, 2, values=[["Region", "Total"], ["North", "120"]])'
    )
    assert lines[1:] == ['t1.style_built_in = "TableGrid"', "t1.header_row_count = 1"]


def test_a_table_that_is_more_than_a_grid_falls_back() -> None:
    """A nested table is not a grid of text: the whole ``w:tbl`` is XML."""
    package = _document("nested-table.docx")
    kinds = {block.kind for block in script_blocks(package.body)}
    assert kinds == {"xml"}


# ---------------------------------------------------------------------------
# lists (CR-003 section 18.7)
# ---------------------------------------------------------------------------


def _two_lists(package):
    """A numbered list with a non-default level 0, and a bullet list after it."""
    body = package.body
    first = body.insert_paragraph("First")
    numbered = first.start_new_list(kind="Number")
    numbered.set_level_numbering(0, "LowerLetter", "%1)")
    body.insert_paragraph("Second").attach_to_list(numbered.id, 0)
    body.insert_paragraph("Nested").attach_to_list(numbered.id, 1)
    bullet = body.insert_paragraph("Bullet one").start_new_list(kind="Bullet")
    body.insert_paragraph("Bullet two").attach_to_list(bullet.id, 0)
    return body


def test_a_list_is_start_new_list_then_attach_to_list(new_package) -> None:
    """One ``start_new_list`` per ``w:numId``, then the script's own ids."""
    body = _two_lists(new_package)
    script = to_api_script(body)
    assert script.count("start_new_list(") == 2
    assert "list1 = p1.start_new_list(kind=\"Number\")" in script
    assert "p2.attach_to_list(list1.id, 0)" in script
    assert "p3.attach_to_list(list1.id, 1)" in script
    assert "list2 = p4.start_new_list(kind=\"Bullet\")" in script
    assert "p5.attach_to_list(list2.id, 0)" in script


def test_a_definition_that_is_not_a_default_emits_the_level_setters(new_package) -> None:
    """``set_level_numbering`` for the level that differs from docx4j's default."""
    body = _two_lists(new_package)
    assert 'list1.set_level_numbering(0, "LowerLetter", "%1)")' in to_api_script(body)


def test_the_executed_scripts_labels_are_the_sources(new_package) -> None:
    """What Word paints in front of each item is what matters, not the ``w:numId``."""
    body = _two_lists(new_package)
    _script, target = executed(new_package)
    labels = [
        item.list_item.list_string if item.is_list_item else None
        for item in target.body.paragraphs
    ]
    assert labels == [
        item.list_item.list_string if item.is_list_item else None for item in body.paragraphs
    ]
    assert labels == ["a)", "b)", "a.", "•", "•"]


def test_a_list_item_with_no_style_of_its_own_keeps_none(new_package) -> None:
    """``start_new_list`` applies *List Paragraph*; the script takes it off again."""
    body = new_package.body
    made = body.insert_paragraph("item")
    made.start_new_list()
    made.style_id = ""  # Word's List Paragraph, taken off again
    script = to_api_script(made)
    assert script.splitlines() == [
        'p1 = body.insert_paragraph("item")',
        'list1 = p1.start_new_list(kind="Number")',
        'p1.style_id = ""',
    ]
    _script, target = executed(new_package)
    assert target.body.paragraphs[0].style_id == "Normal"
    assert target.body.paragraphs[0].is_list_item


# ---------------------------------------------------------------------------
# pictures
# ---------------------------------------------------------------------------


def test_a_placeholder_picture_executes_and_keeps_its_size_and_alt_text() -> None:
    """The default mode: a 1x1 PNG constant, with the part name in a comment."""
    package = _document("Images.docx")
    script, target = executed(package)
    assert f'PLACEHOLDER_PNG = "{PLACEHOLDER_PNG}"' in script
    assert "# /word/media/image2.jpeg, 16095 bytes" in script
    made = target.body.inline_pictures
    assert [picture.width for picture in made] == [
        picture.width for picture in package.body.inline_pictures
    ]
    assert [picture.alt_text_description for picture in made] == [
        picture.alt_text_description for picture in package.body.inline_pictures
    ]


def test_an_inline_picture_reproduces_the_bytes() -> None:
    """``pictures="inline"`` writes the real base64."""
    package = _document("Images.docx")
    _script, target = executed(package, pictures="inline")
    assert [len(picture.get_bytes()) for picture in target.body.inline_pictures] == [
        len(picture.get_bytes()) for picture in package.body.inline_pictures
    ]


# ---------------------------------------------------------------------------
# comments (CR-003 Phase G)
# ---------------------------------------------------------------------------


def test_a_thread_with_a_reply_and_a_resolved_thread_reproduce() -> None:
    """The anchor, the author, the reply nesting and ``resolved``."""
    package = threaded_package()
    script, target = executed(package)
    assert "from docx4j_py.model.content import Author" in script
    assert 'body.package.author = Author("Ada Lovelace", initials="AL"' in script

    def summary(pkg):
        return [
            (
                comment.author_name,
                comment.content,
                comment.resolved,
                [reply.content for reply in comment.replies],
            )
            for comment in pkg.body.get_comments()
        ]

    assert summary(target) == summary(package)


# ---------------------------------------------------------------------------
# the options
# ---------------------------------------------------------------------------


def test_limit_stops_and_says_where(body) -> None:
    """A budget: the blocks, then a comment naming how many are left."""
    for index in range(5):
        body.insert_paragraph(f"para {index}")
    script = to_api_script(body, limit=2)
    assert script.count("insert_paragraph") == 2
    assert script.splitlines()[-1] == (
        "# ... 3 more blocks; call to_api_script(body, limit=None) "
        "or pass a Paragraph/Table for one"
    )
    assert to_api_script(body, limit=4).splitlines()[-1].startswith("# ... 1 more block;")


def test_addresses_are_the_ordinals(body) -> None:
    """``addresses=True`` names the block, in the ordinal form of section 3.4."""
    body.insert_paragraph("one")
    body.insert_paragraph("two")
    assert to_api_script(body, addresses=True).splitlines() == [
        "# body/0",
        'body.insert_paragraph("one")',
        "",
        "# body/1",
        'body.insert_paragraph("two")',
    ]


def test_the_variable_is_the_callers(body) -> None:
    """The name the script inserts into."""
    body.insert_paragraph("in a cell")
    assert to_api_script(body, variable="cell") == 'cell.insert_paragraph("in a cell")'


def test_the_location_is_passed_to_every_block_insert(body) -> None:
    """``location="Start"`` shows in the call, and ``"End"`` is left implicit."""
    body.insert_paragraph("one")
    assert to_api_script(body, location="Start") == (
        'body.insert_paragraph("one", location="Start")'
    )


def test_a_range_emits_its_runs_only(body) -> None:
    """The TypeScript rule: no paragraph properties, against a ``Paragraph`` name."""
    body.insert_xml(
        "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>plain </w:t></w:r>"
        '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>red</w:t></w:r></w:p>'
    )
    span = body.paragraphs[0].search("red")[0]
    assert span.to_api_script(variable="p1").splitlines() == [
        'r1 = p1.insert_text("red")',
        "r1.font.italic = True",
    ]


# ---------------------------------------------------------------------------
# determinism, and the script itself (CR-003 section 12.6)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", ROUND_TRIP)
def test_the_same_document_gives_the_same_script(name: str) -> None:
    """Byte for byte, twice, from two loads."""
    assert to_api_script(_document(name).body) == to_api_script(_document(name).body)


@pytest.mark.parametrize("name", ROUND_TRIP)
def test_every_script_is_python(name: str) -> None:
    """``ast.parse`` on the whole script, which is what ``exec`` will compile."""
    ast.parse(to_api_script(_document(name).body))


def test_the_script_passes_ruff() -> None:
    """The emitted source is code someone reads: it is clean by the repository's rules."""
    import shutil
    import subprocess

    ruff = shutil.which("ruff", path=str(FIXTURES.parent.parent / ".venv-fork" / "bin"))
    if ruff is None:  # pragma: no cover - a checkout without the virtualenv
        pytest.skip("ruff is not installed")
    # `body` is what `exec` puts in scope, so the check is given one too
    script = "body = None\n" + to_api_script(_document("sample-docx.docx").body)
    done = subprocess.run(
        [ruff, "check", "--stdin-filename", "script.py", "-"],
        input=script.encode(),
        capture_output=True,
        check=False,
    )
    assert done.returncode == 0, done.stdout.decode() + done.stderr.decode()


@pytest.mark.slow
@pytest.mark.parametrize("name", sorted(path.name for path in SAMPLES.glob("*.docx")))
def test_every_sample_executes(name: str) -> None:
    """The whole corpus: the script runs against a fresh body, and keeps the text."""
    package = load(SAMPLES / name)
    package.id_seed = 20260917
    _script, target = executed(package)
    assert target.body.get_text() == package.body.get_text()


def test_scripts_directory_is_on_the_path() -> None:
    """``canon`` comes from ``scripts/``, as the other tests take it."""
    assert any(path.endswith("scripts") for path in sys.path)
