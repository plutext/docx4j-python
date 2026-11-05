"""``profile`` (REQ-1002): docx4j ``TextOptionsProfileTest``. CR-027 Phase D, sections 2.1 and 4.1.

docx4j's ``TextOptionsProfileTest`` at ``docx4j-17.3.5``, its four tests, the
``list_labels`` value at ``87c421731`` (WD 0.84, section 4.1's amendment):

- ``wordTextIsWordsPlainTextAsTheOptionsReachIt``: the preset's values;
- ``theCallersOptionOverridesTheProfileWhicheverCameFirst``: both orders, through
  every route a caller has here (a keyword, :meth:`TextOptions.set`,
  :meth:`TextOptions.of`, a mapping in either order, :meth:`TextOptions.include`
  and an attribute);
- ``anUnknownProfileIsRefused``;
- ``theMemoUnderThePreset``: the fields-fun memo's stored results.

Then this port's own (section 4.1): the defaults and :meth:`TextOptions.copy`
recording nothing, the name's forms, ``to_dict`` and a round trip by name, the
preset's two values that are CR-029's left out, a list label under the preset,
and the docstring's list of what the preset does not reproduce (REQ-1002's
"MUST document"). ``dpp/``'s six Word saves are checked by ``test_conformance.py``.
"""

from __future__ import annotations

import hashlib
import subprocess
from pathlib import Path

import pytest
from mini_docx import MiniDocx

from docx4j_py import load
from docx4j_py.model.text import OPTION_NAMES, TextOptionError, TextOptions, extract, plain_text

ROOT = Path(__file__).resolve().parent.parent.parent
MEMO = ROOT / "tests" / "fixtures" / "fields-fun-cr067" / "one-file-many-readings.docx"
DOCX4J = ROOT.parent / "docx4j"

#: The preset's values by the specification's names: docx4j's ``WORD_TEXT_VALUES`` less ``drop_caps`` and
#: ``small_caps``, which come with their options in CR-029 (section 4.1).
WORD_TEXT = {
    "view": "accepted",
    "include": ["headers", "footers", "comments"],
    "text_boxes": False,
    "notes": True,
    "note_references": True,
    "comments": True,
    "hidden": True,
    "field_codes": False,
    "fields": "cached",
    "toc": True,
    "placeholders": True,
    "bindings": "bound",
    "list_labels": True,
    "charts": False,
    "smartart": False,
    "embedded": False,
    "altchunks": True,
    "alt_text": False,
    "metadata": False,
    "symbols": "asis",
    "math": "latex",
    "tables": "tsv",
    "page_markers": False,
    "layout": "none",
}


def test_word_text_is_words_plain_text_as_the_options_reach_it():
    options = TextOptions(profile="word-text")
    assert options.view == "accepted"
    assert options.hidden
    assert {"headers", "footers"} <= options.includes
    assert options.comments
    assert not options.text_boxes
    assert options.notes and options.note_references
    assert options.list_labels, "Word's Save As Plain Text writes list labels (WD 0.75; label-separator probe)"
    assert options.fields == "cached"
    assert not options.field_codes
    assert options.toc
    assert options.placeholders
    assert options.bindings == "bound"
    assert options.altchunks
    assert options.symbols == "asis"
    assert options.tables == "tsv"
    assert not options.page_markers
    assert options.layout == "none"
    assert not (options.alt_text or options.metadata or options.charts or options.smartart or options.embedded)
    assert options.math == "latex"  # Word's linear math is not produced (unicodemath not implemented)
    assert options.to_dict()["profile"] == "word-text"


def test_the_preset_is_docx4js_24_values():
    """Section 2.1's list, each value, and nothing else moved from the defaults."""
    assert len(WORD_TEXT) == 24
    got = TextOptions(profile="word-text").to_dict()
    assert {name: got[name] for name in WORD_TEXT} == WORD_TEXT
    defaults = TextOptions().to_dict()
    others = set(OPTION_NAMES) - WORD_TEXT.keys() - {"profile"}
    assert {name: got[name] for name in others} == {name: defaults[name] for name in others}
    assert others == {"markup_detail", "warnings", "duplicate_entries"}


def test_the_callers_option_overrides_the_profile_whichever_came_first():
    after = TextOptions().set("profile", "word-text").set("fields", "update").set("tables", "grid")
    assert (after.fields, after.tables) == ("update", "grid")
    before = TextOptions().set("fields", "update").set("tables", "grid").set("profile", "word-text")
    assert (before.fields, before.tables) == ("update", "grid")
    assert before.hidden  # the rest of the preset applied
    # by name, in either map order, as docx4j-mcp passes them
    from_map = TextOptions.of({"fields": "update", "profile": "word-text"})
    assert from_map.fields == "update"
    assert from_map.hidden and from_map.toc
    other_order = TextOptions.of({"profile": "WORD_TEXT", "toc": False})
    assert not other_order.toc
    assert other_order.hidden


@pytest.mark.parametrize(
    "made",
    [
        lambda: TextOptions(fields="update", profile="word-text"),
        lambda: TextOptions(profile="word-text", fields="update"),
        lambda: TextOptions.of({"fields": "update"}, profile="word-text"),
        lambda: TextOptions.of({"profile": "word-text"}, fields="update"),
        lambda: TextOptions.of(TextOptions(fields="update"), profile="word-text"),
        lambda: TextOptions.of(TextOptions(profile="word-text"), fields="update"),
    ],
    ids=["keywords", "keywords after", "of", "of after", "of options", "of options after"],
)
def test_either_order_by_every_route(made):
    options = made()
    assert options.fields == "update"
    assert options.hidden


def test_attribute_assignment_is_recorded():
    """Q5: without the ``__setattr__``, ``opts.fields = "update"`` then the profile would give ``cached``."""
    before = TextOptions()
    before.fields = "update"
    before.comments = False
    before.set("profile", "word-text")
    assert before.fields == "update"
    assert before.hidden
    after = TextOptions(profile="word-text")
    after.fields = "update"
    assert after.fields == "update"


def test_include_is_recorded():
    """docx4j's builder ``include(Include...)``: the caller's stories stand."""
    options = TextOptions()
    options.include("glossary")
    options.set("profile", "word-text")
    assert options.includes == {"glossary", "comments"}  # comments: true is the preset's


def test_a_preset_value_is_not_the_callers():
    """A preset's value does not record: a second preset would fill it, and a caller's own still overrides it."""
    options = TextOptions(profile="word-text")
    assert "hidden" not in options._explicit
    options.set("hidden", False)
    options.set("profile", "word-text")
    assert not options.hidden


def test_the_defaults_and_a_copy_record_nothing():
    """Section 4.1: ``__init__`` and ``copy()`` past the recording, or the preset would fill nothing."""
    assert TextOptions()._explicit == set()
    assert TextOptions().copy()._explicit == set()
    assert TextOptions.of(TextOptions(), profile="word-text").hidden
    assert TextOptions().copy().set("profile", "word-text").hidden


def test_a_copy_carries_what_was_recorded_and_the_profile():
    options = TextOptions(fields="update").copy()
    assert options.set("profile", "word-text").fields == "update"
    profiled = TextOptions(profile="word-text").copy()
    assert profiled.profile == "word-text"
    assert profiled.hidden
    copied = TextOptions(fields="update")
    twin = copied.copy()
    twin.set("hidden", True)
    assert "hidden" not in copied._explicit


@pytest.mark.parametrize("name", ["word-text", "word_text", "WORD-TEXT", "WORD_TEXT", "Word-Text", " word-text "])
def test_the_name_is_matched_without_case_with_a_hyphen_or_an_underscore(name):
    assert TextOptions(profile=name).profile == "word-text"
    assert TextOptions().set("profile", name).hidden
    assert TextOptions.of({"profile": name}).hidden


@pytest.mark.parametrize("name", ["libreoffice", "wordtext", "", "word text", 7])
def test_an_unknown_profile_is_refused(name):
    with pytest.raises(TextOptionError) as caught:
        TextOptions.of({"profile": name})
    assert "REQ-1002" in str(caught.value)
    assert (caught.value.option, caught.value.statement, caught.value.reason) == ("profile", "REQ-1002", "invalid")


def test_a_profile_attribute_applies_the_preset():
    options = TextOptions()
    options.profile = "word-text"
    assert options.profile == "word-text"
    assert options.hidden
    with pytest.raises(TextOptionError):
        options.profile = "libreoffice"


def test_no_profile():
    """``profile`` is None until one is set; None sets it back, and takes no value back: docx4j's ``profile(null)``."""
    assert TextOptions().profile is None
    options = TextOptions(profile="word-text").set("profile", None)
    assert options.profile is None
    assert options.hidden


def test_to_dict_carries_profile_last():
    assert list(TextOptions().to_dict())[-1] == "profile"
    assert TextOptions().to_dict()["profile"] is None
    assert list(TextOptions(profile="word-text").to_dict())[-1] == "profile"


def test_options_round_trip_by_name():
    options = TextOptions(profile="word-text", fields="update", tables="grid")
    assert TextOptions.of(options.to_dict()).to_dict() == options.to_dict()


def test_repr_names_the_profile():
    assert "profile='word-text'" in repr(TextOptions(profile="word-text"))
    assert repr(TextOptions()) == "TextOptions()"


def test_the_preset_is_implemented():
    """``math: latex``, not the WD's ``unicodemath``, so a reading under the preset is not refused (Q5)."""
    TextOptions(profile="word-text").check_implemented()


def test_the_values_that_are_cr_029s_are_not_options_yet():
    """``drop_caps`` and ``small_caps`` come with their options in the CR-069 port (section 4.1)."""
    with pytest.raises(TextOptionError):
        TextOptions(small_caps="upper")
    with pytest.raises(TextOptionError):
        TextOptions(drop_caps="asis")


def test_what_the_preset_does_not_reproduce_is_documented():
    """REQ-1002: "MUST document", at least the NOTE's list, with ``math`` among it (Q5) and the space after a label."""
    doc = " ".join((TextOptions.__doc__ or "").split())
    for said in (
        "the notes' and comments' forms",
        "CRLF lines and CR cells",
        "one space",
        "folded hyphens",
        "linear math",
        "``math: unicodemath``",
        "ANSI code page",
        "DATE, TIME or SAVEDATE",
        "SYMBOL",
        "stale PAGE",
    ):
        assert said in doc, said


def test_the_memo_is_the_copy_sha1sums_lists():
    digest = hashlib.sha1(MEMO.read_bytes()).hexdigest()
    assert (MEMO.parent / "SHA1SUMS").read_text(encoding="utf-8").split() == [digest, MEMO.name]
    assert sorted(p.name for p in MEMO.parent.iterdir()) == ["README.md", "SHA1SUMS", MEMO.name]


@pytest.mark.skipif(not (DOCX4J / ".git").exists(), reason="needs the sibling checkout ../docx4j")
def test_the_memo_is_docx4js_at_the_tag():
    path = f"docx4j-17.3.5:docx4j-core-tests/src/test/resources/fields/cr067-fields-fun/{MEMO.name}"
    blob = subprocess.run(["git", "-C", str(DOCX4J), "rev-parse", path], capture_output=True, check=True, text=True)
    data = MEMO.read_bytes()
    assert hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest() == blob.stdout.strip()


def test_the_memo_under_the_preset():
    """Under the preset a document reads as Word's plain text where the two agree (the memo's stored results)."""
    text = plain_text(extract(load(MEMO), profile="word-text"))
    assert "Payment terms: Net 90 days." in text
    assert "Discount: 10%." in text
    assert "Reference: Purchase approval." in text


def _numbered(suff: str) -> MiniDocx:
    numbering = (
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>{suff}'
        '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    )
    return MiniDocx.body(
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>'
    ).part(
        "numbering.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "rIdN",
        numbering,
    )


@pytest.mark.parametrize(
    ("suff", "written"),
    [("", "1.\titem\n"), ('<w:suff w:val="space"/>', "1. item\n"), ('<w:suff w:val="nothing"/>', "1.item\n")],
    ids=["tab", "space", "nothing"],
)
def test_a_label_under_the_preset_is_followed_by_the_levels_suffix(suff, written):
    """WD 0.84: ``list_labels: true``, the label with the level's ``w:suff``, where Word's save writes one space.

    That space is among what the preset does not reproduce (section 4.1), so
    a Word save of a list is not this text; ``dpp/``'s 22 hold no list.
    """
    package = _numbered(suff).load()
    assert plain_text(extract(package, profile="word-text")) == written
    assert plain_text(extract(package, profile="word-text", list_labels=False)) == "item\n"
