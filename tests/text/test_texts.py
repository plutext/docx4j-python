"""The text model, ``docx4j_py.model.text``. CR-011 Phase B (section 12).

docx4j's ``TextsTest``, ``TextsTocTest`` and ``LevelOwnerTest`` where their
statements are Phase B's, after core-ts CR-012's ``text-model.test.mjs``,
with this package's own: the understood list (REQ-306), the options
(REQ-1001), the package unchanged (REQ-204), the resolution by namespace URI
(REQ-303), the text boxes' addresses (REQ-519), the notes, headers, comments
and table of contents, the runs' ``office_offset`` against the content API's
``get_text`` (REQ-202's exemption, CR-011 Q4), and the probe's goldens as
Phase B can hold them.
"""

from __future__ import annotations

import json

import pytest
from conformance import PHASE0, ROOT, SUITE

from docx4j_py import load
from docx4j_py.model.listnumbering.definitions import NumberingState
from docx4j_py.model.listnumbering.emulator import Emulator, emulator_of
from docx4j_py.model.text import (
    OPTION_NAMES,
    TextOptionError,
    TextOptions,
    extract,
    plain_text,
)

PROBE = PHASE0 / "extraction-probe.docx"
STRUCTURE = PHASE0 / "extraction-probe-structure.docx"
LEVEL_OWNER = SUITE / "level-owner"
SIBLING_LIST = ROOT.parent / "docx4j" / "docx4j-core" / "src" / "main" / "resources" / "org" / "docx4j" / "jaxb" / "understood-namespaces.txt"


@pytest.fixture(scope="module")
def probe():
    return load(PROBE)


# ---------------------------------------------------------------------------
# REQ-306: the understood list
# ---------------------------------------------------------------------------


def test_the_understood_list_is_docx4js_verbatim():
    """The package's copy is the suite's, which is docx4j's at 198d015a3, and the sibling's where it is checked out."""
    resource = ROOT / "docx4j_py" / "resources" / "understood-namespaces.txt"
    assert resource.read_bytes() == (SUITE / "understood-namespaces.txt").read_bytes()
    if SIBLING_LIST.is_file():
        assert resource.read_bytes() == SIBLING_LIST.read_bytes()
    from docx4j_py.child import understood_namespaces

    assert len(understood_namespaces()) == 118
    assert "urn:schemas-microsoft-com:vml" in understood_namespaces()


# ---------------------------------------------------------------------------
# REQ-1001: the options
# ---------------------------------------------------------------------------


def test_the_options_are_section_10s_names():
    assert len(OPTION_NAMES) == 26
    assert list(TextOptions().to_dict()) == list(OPTION_NAMES)
    defaults = TextOptions().to_dict()
    assert defaults["view"] == "accepted" and defaults["text_boxes"] and defaults["notes"]
    assert not defaults["comments"] and not defaults["hidden"] and not defaults["toc"] and not defaults["list_labels"]


def test_an_option_is_set_by_its_name_and_a_strings_value():
    o = TextOptions.of({"view": "current", "hidden": "true", "include": "headers, footers", "tables": "TSV"})
    assert o.view == "accepted" and o.hidden is True and o.tables == "tsv"
    assert o.includes == {"headers", "footers"}
    assert TextOptions(comments=True).includes == {"comments"}


@pytest.mark.parametrize(
    ("name", "value", "reason", "statement"),
    [
        ("colour", True, "unknown", "REQ-1001"),
        ("hidden", "yes", "invalid", "REQ-507"),
        ("view", "final", "invalid", "REQ-401"),
        ("include", "footnotes", "invalid", "REQ-501"),
    ],
)
def test_an_option_section_10_lacks_is_refused(name, value, reason, statement):
    with pytest.raises(TextOptionError) as raised:
        TextOptions().set(name, value)
    assert raised.value.reason == reason and raised.value.statement == statement


@pytest.mark.parametrize(("name", "value", "statement"), [("math", "unicodemath", "REQ-604"), ("symbols", "unicode", "REQ-601")])
def test_what_docx4j_refuses_is_refused(probe, name, value, statement):
    with pytest.raises(TextOptionError) as raised:
        extract(probe, **{name: value})
    assert raised.value.reason == "not_implemented" and raised.value.statement == statement


def test_every_value_is_acted_on_since_phase_c():
    from docx4j_py.model.text import NOT_YET

    assert NOT_YET == {}  # CR-011 section 13
    everything = TextOptions(fields="update", embedded=True, alt_text=True, metadata=True, tables="grid")
    assert everything.not_yet() == []


# ---------------------------------------------------------------------------
# REQ-204: the package is never changed
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("path", [PROBE, STRUCTURE, PHASE0 / "extraction-probe-external.docm"])
def test_a_reading_changes_nothing(path):
    package = load(path)
    before = package.save()
    for view in ("accepted", "original", "markup"):
        for options in ({}, {"include": "headers, footers", "comments": True, "hidden": True, "field_codes": True, "toc": True, "list_labels": True}):
            extract(package, view=view, **options)
    assert not package.main_document_part.is_unmarshalled
    assert package.save() == before


# ---------------------------------------------------------------------------
# REQ-301 to REQ-307: resolve
# ---------------------------------------------------------------------------


def test_a_choice_is_selected_by_namespace_uri(probe):
    """Case 6 binds the wps namespace to shape2010: its Choice is read (REQ-303, CR-011 Q3)."""
    text = plain_text(extract(probe))
    assert "MCE6_REBOUND_CHOICE" in text and "MCE6_FALLBACK" not in text
    assert "MCE7_FALLBACK" in text  # wps zz9: one unknown, the Fallback (REQ-302)
    assert "CHOICE_ONLY_789" in text  # Choices and no Fallback


def test_branches_that_differ_are_warned(probe):
    warnings = [w for w in extract(probe).warnings if w.kind == "mce.branches_differ"]
    assert ("body/3", 'mc:Choice Requires="wps"') in {(w.address, w.branch) for w in warnings}
    assert ("body/13", 'mc:Choice Requires="shape2010"') in {(w.address, w.branch) for w in warnings}
    assert not any("CHOICE_ALPHA" in str(w.to_dict()) for w in warnings)  # never the text (REQ-802)
    assert not [w for w in extract(probe, warnings=False).warnings if w.kind == "mce.branches_differ"]


# ---------------------------------------------------------------------------
# scope
# ---------------------------------------------------------------------------


def test_a_text_box_follows_its_anchor_with_a_textbox_address(probe):
    blocks = extract(probe).blocks
    at = {b.address: b for b in blocks}
    box = at["body/3/textbox:0/0"]
    assert box.story == "TEXT_BOX" and box.textbox and box.text == "CHOICE_ALPHA_123"
    assert blocks[blocks.index(box) - 1].address == "body/3"
    cell_box = at["body/68/0/0/0/textbox:0/0"]
    assert blocks[blocks.index(cell_box) - 1].kind == "TABLE"  # after the table its anchor is in
    assert not [b for b in extract(probe, text_boxes=False).blocks if b.textbox]


def test_notes_are_numbered_in_reference_order_and_appended(probe):
    result = extract(probe)
    text = plain_text(result)
    assert "FN27_REF_HERE [2]" in text and "EN27_REF_HERE [1]" in text
    assert text.endswith("[1] TC17_NOTE_BASE \n[2] FN27_NOTE_TEXT\n[1] EN27_NOTE_TEXT\n")
    assert "FN27_DELETED_REF_NOTE" not in text  # its reference is deleted (REQ-402)
    assert "FN27_DELETED_REF_NOTE" in plain_text(extract(probe, view="original"))
    assert "[1]" not in plain_text(extract(probe, notes=False))


def test_headers_and_footers_by_section(probe):
    text = plain_text(extract(probe, include="headers, footers"))
    lines = text.split("\n")
    assert lines[:3] == ["HDR26_FIRST", "HDR26_DEFAULT ", "TBX25_IN_HEADER"]
    assert lines.index("FTR26_DEFAULT") < lines.index("HDR26_SECTION2") < lines.index("The first paragraph of section 2.")
    assert "HDR26_ORPHAN" not in text
    blocks = extract(probe, include="headers").blocks
    assert blocks[0].address.startswith("header:rId") and blocks[0].story == "HEADER"


def test_comments_in_markup_with_their_replies(probe):
    text = plain_text(extract(probe, view="markup"))
    assert "CMT28_ANCHOR{>>CMT28_COMMENT_TEXT<<}{>>CMT28_REPLY<<}" in text
    assert "CMT28_COMMENT_TEXT" not in plain_text(extract(probe))
    detailed = plain_text(extract(probe, view="markup", markup_detail="author"))
    assert "{>>" in detailed and ": CMT28_COMMENT_TEXT<<}" in detailed


def test_a_table_of_contents_only_under_toc(probe):
    structure = load(STRUCTURE)
    assert "TOC55_ENTRY" not in plain_text(extract(structure))
    with_toc = plain_text(extract(structure, toc=True))
    assert with_toc.count("HEAD55_ONE") == 2


def test_hidden_text_through_the_styles(probe):
    text = plain_text(extract(probe))
    assert "HID18_VANISH" not in text and "HID18_STYLE_VANISH" not in text
    assert "HID18_WHITE_TEXT" in text
    hidden = extract(probe, hidden=True)
    assert "HID18_STYLE_VANISH" in plain_text(hidden)
    assert any(r.hidden for b in hidden.blocks for r in b.runs if "HID18_VANISH" in r.text)


def test_field_codes_under_field_codes(probe):
    assert "FLD19_INSTR_ONLY" not in plain_text(extract(probe))
    with_codes = extract(probe, field_codes=True)
    assert "FLD19_INSTR_ONLY" in plain_text(with_codes)
    assert any(r.kind == "FIELD_CODE" for b in with_codes.blocks for r in b.runs)


def test_the_views(probe):
    accepted, original, markup = (plain_text(extract(probe, view=v)) for v in ("accepted", "original", "markup"))
    assert "TC11_FIRST TC11_SECOND" in accepted  # a deleted mark joins (REQ-402)
    assert "TC12_ONE TC12_TWO" in original  # an inserted mark rejected joins
    assert "TC13_DELETED_ROW" not in accepted and "TC13_INSERTED_ROW" not in original
    assert "TC9_BASE {++TC9_INSERTED ++}{--TC9_DELETED--}" in markup
    assert "{++TC14_INS_KEPT ++}{--TC14_DEL_IN_INS--}" in markup  # flat, never nested


# ---------------------------------------------------------------------------
# the plain-text goldens, exact since Phase C (CR-011 section 13)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("fixture", ["extraction-probe", "extraction-probe-structure", "extraction-probe-external"])
@pytest.mark.parametrize("view", ["accepted", "original", "markup"])
def test_the_plain_text_goldens_are_exact(fixture, view):
    path = next(PHASE0.glob(f"{fixture}.doc*"))
    golden = (PHASE0 / f"{fixture}.{view}.txt").read_text(encoding="utf-8")
    # the suite compares a golden after REQ-1201's normalisation; this, byte for byte but for the final line feeds
    assert plain_text(extract(load(path), view=view)).rstrip("\n") == golden.rstrip("\n")


# ---------------------------------------------------------------------------
# REQ-202's exemption: each run's place in Office JS's text (CR-011 Q4)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("view", ["accepted", "original"])
def test_a_runs_office_offset_is_where_its_text_is_in_get_text(probe, view):
    from docx4j_py.model.content.text_model import text_of_view

    checked = 0
    for block in extract(probe, view=view).blocks:
        for run in block.runs:
            if run.office_offset is None:
                continue
            office = text_of_view(run.source.paragraph, view=view)
            assert office[run.office_offset : run.office_offset + len(run.text)] == run.text, block.address
            checked += 1
    assert checked > 100


def test_hidden_text_is_the_gap_between_offsets():
    """Text Paragraph.text holds and the model leaves out lies between two runs' offsets."""
    from docx4j_py.model.content.text_model import text_of_view
    from docx4j_py.wml import wml

    p = wml(
        "<w:p><w:r><w:t>shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>gone </w:t></w:r>"
        "<w:r><w:t>again</w:t></w:r></w:p>"
    )
    runs = extract(p).blocks[0].runs
    assert [(r.text, r.office_offset) for r in runs] == [("shown ", 0), ("again", 11)]
    office = text_of_view(p)
    assert office[runs[0].office_offset + len(runs[0].text) : runs[1].office_offset] == "gone "


# ---------------------------------------------------------------------------
# REQ-510: the level's owner (docx4j LevelOwnerTest)
# ---------------------------------------------------------------------------

LABELS = {
    "numbering-linked-level-other-style": "- - A0-1 - A0-2 - A0-3 A1-1 A0-4 A0-5 A0-6 - A0-7 - A0-8 - A0-9 A0-10 A0-11 A0-12 A0-13 C0-1 - C0-2 A0-14",
    "numbering-unlinked-level-claim": "- - D0-1 - D0-2 - E0-1 - E0-2 F0-1 - F0-2 - - - H0-1 H0-2 H0-3 H0-4 I0-1 - I0-2 J0-1 - J1-1 - J1-2",
    "numbering-shared-abstract-styles": "- K0-1 - K0-2 - L0-1 - L0-2 -",
}


@pytest.mark.parametrize("name", list(LABELS))
def test_the_emulator_numbers_as_word_does(name):
    package = load(LEVEL_OWNER / f"{name}.docx")
    state = NumberingState()
    from docx4j_py.wml import P

    labels = []
    for item in package.main_document_part.contents.body.content:
        if isinstance(item, P):
            result = Emulator.get_number(package, item.p_pr, state)
            labels.append("-" if result is None or result.num_string is None else result.num_string)
    assert " ".join(labels) == LABELS[name]


@pytest.mark.parametrize("name", list(LABELS))
def test_the_text_model_numbers_as_word_does(name):
    package = load(LEVEL_OWNER / f"{name}.docx")
    emulator = emulator_of(package)
    labels = []
    for block in extract(package, list_labels=True).blocks:
        if block.kind != "PARAGRAPH":
            continue
        label = "-" if block.list_label is None else block.list_label.strip()
        assert (label == "-") == (block.numbering is None), block.address
        if block.numbering is not None:
            assert block.numbering.value == int(label.split("-")[1])
            definition = emulator.instance_list_definitions[block.numbering.num_id]
            assert block.numbering.abstract_num_id == definition.abstract_definition.id
        labels.append(label)
    assert " ".join(labels) == LABELS[name]


def test_the_owners():
    emulator = emulator_of(load(LEVEL_OWNER / "numbering-unlinked-level-claim.docx"))
    owners = {
        ("110", "0"): "DZebra",
        ("111", "0"): "EAlpha",
        ("112", "0"): "FZebra",
        ("113", "0"): "GAlphaUnused",
        ("114", "0"): "HBase",
        ("115", "0"): "IOwner",
        ("116", "0"): "JLinked",
        ("116", "1"): "OtherIlvl1",
        ("116", "2"): None,
    }
    assert {key: emulator.level_style_id(*key) for key in owners} == owners


def test_the_style_that_lost_the_level_keeps_its_indent():
    from docx4j_py.wml import PPr, el

    package = load(LEVEL_OWNER / "numbering-unlinked-level-claim.docx")
    ind = package.get_property_resolver().get_effective_ppr(PPr(p_style=el.pStyle(val="IOther"))).ind
    assert (int(ind.left), int(ind.hanging)) == (720, 720)


# ---------------------------------------------------------------------------
# the block model and the overloads
# ---------------------------------------------------------------------------


def test_the_block_model_is_json(probe):
    result = extract(probe, view="markup")
    data = json.loads(json.dumps(result.to_dict()))
    assert data["options"]["view"] == "markup"
    first = data["blocks"][0]
    assert first["address"] == "body/0" and first["kind"] == "PARAGRAPH" and "runs" in first
    assert any(b.get("numbering") for b in data["blocks"])
    assert data["comments"]


def test_content_is_read_in_the_packages_view(probe):
    nine = next(p for p in probe.body.paragraphs if p.text.startswith("TC9_BASE"))
    result = extract(nine)
    assert result.blocks[0].address == "content/0"
    assert plain_text(result) == "TC9_BASE TC9_INSERTED \n"
    assert plain_text(extract(nine, view="markup")) == "TC9_BASE {++TC9_INSERTED ++}{--TC9_DELETED--}\n"


def test_content_with_no_package_is_read_as_it_stands():
    from docx4j_py.wml import wml

    p = wml("<w:p><w:r><w:t>a</w:t></w:r><w:ins><w:r><w:t>b</w:t></w:r></w:ins><w:del><w:r><w:delText>c</w:delText></w:r></w:del></w:p>")
    assert plain_text(extract(p)) == "ab\n"
    assert plain_text(extract(p, view="original")) == "ac\n"
    assert plain_text(extract(p, view="markup")) == "a{++b++}{--c--}\n"
