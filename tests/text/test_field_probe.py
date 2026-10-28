"""docx4j CR-067 phase 1's Word probe, replayed through the text model. CR-025 Phase B.

docx4j's ``FieldProbeWordTest`` at ``c3e9920e5``, through ``fields: update``
instead of ``FieldUpdater``: the probe (``tests/fixtures/field-probe-cr067/``,
copied from docx4j) gives, case for case, what Word 365 gave on F9 as
``results.txt`` records it, in Sydney's zone, where the probe ran. Where Word
writes an error text (``INFO`` with no name or an unknown one, ``REF`` and
``NOTEREF`` to a missing bookmark), the field keeps its cached ``OLD``
(docx4j's D5). The user fields need a user, as docx4j's ``FieldContext.user``
does: supplied (``TextOptions(field_user=...)``), Word's; not, their cached
``OLDUSER``, which is also what Word showed before F9.

What only a write-back shows (a ``REF``'s result in the bookmark's bold,
``NOTEREF \\f``'s ``FootnoteReference`` style, ``\\h``'s link) is pinned as
not extraction's: the text model's runs carry no such formatting, docx4j's
Markdown of the same reading shows none, and Word's save, read cached, does.

And docx4j's own extraction of the probe, run against its build of
``c3e9920e5`` (the fixture's ``README.md``): byte for byte.
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

import pytest

from docx4j_py import load
from docx4j_py.model.fields import FieldUser
from docx4j_py.model.text import extract, markdown, plain_text

PROBE = Path(__file__).parent.parent / "fixtures" / "field-probe-cr067"
WORD_USER = FieldUser("Author A", "AA")
SYDNEY = "Australia/Sydney"


def word_results() -> dict[str, str]:
    """Word's results by token, from ``results.txt``'s ``Results:`` section, as ``FieldProbeWordTest`` reads it."""
    text = (PROBE / "results.txt").read_text(encoding="utf-8")
    section = text[text.index("Results:") :]
    return {m.group(1): m.group(2).strip() for m in re.finditer(r"^(P\d\d_[A-Z0-9_]+) = (.*)$", section, re.MULTILINE)}


def results(pkg, **options) -> dict[str, str]:
    """Each token's result in the plain text: the line ``P01_CREATEDATE = <result>``."""
    out = {}
    for line in plain_text(extract(pkg, **options)).splitlines():
        found = re.match(r"^(P\d\d_[A-Z0-9_]+) = (.*)$", line)
        if found:
            out[found.group(1)] = found.group(2).strip()
    return out


def test_the_probe_files_are_docx4js():
    sums = dict(reversed(line.split()) for line in (PROBE / "SHA1SUMS").read_text().splitlines())
    assert len(sums) == 10
    for name, digest in sums.items():
        assert hashlib.sha1((PROBE / name).read_bytes()).hexdigest() == digest, name


def test_results_txt_holds_the_25_cases():
    word = word_results()
    assert len(word) == 25
    assert word["P01_CREATEDATE"] == "2/3/2024 3:05:00 PM"
    assert word["P12_QUOTE_NOARG"] == ""
    assert sum(value.startswith("Error!") for value in word.values()) == 4


def _replay(got: dict[str, str]) -> int:
    """Word's result in every case, the cached ``OLD`` where Word writes an error text; how many were compared."""
    compared = 0
    for token, expected in word_results().items():
        if expected.startswith("Error!"):
            assert got[token] == "OLD", token  # D5: the cached result is kept where Word writes its error text
            continue
        assert got[token] == expected, token
        compared += 1
    return compared


def test_the_text_model_gives_words_results_in_the_processs_zone(zone):
    """``FieldProbeWordTest.fieldUpdaterGivesWordsResults``: the process in Sydney's zone, no package zone."""
    zone(SYDNEY)
    got = results(load(PROBE / "field-probe-r1.docx"), fields="update", field_user=WORD_USER)
    assert _replay(got) == 21


@pytest.mark.parametrize("place", ["America/New_York", "Asia/Kolkata", "UTC"])
def test_the_text_model_gives_words_results_in_the_packages_zone(zone, place):
    """The client's zone on the package (CR-025 D1): the process's plays no part."""
    zone(place)
    pkg = load(PROBE / "field-probe-r1.docx")
    pkg.date_time_zone = SYDNEY
    assert _replay(results(pkg, fields="update", field_user=WORD_USER)) == 21


def test_without_a_user_the_user_fields_keep_their_cached_result(zone):
    """docx4j's extraction has no user to give: ``OLDUSER``, what Word showed before F9 (``results.txt`` step 2)."""
    zone(SYDNEY)
    got = results(load(PROBE / "field-probe-r1.docx"), fields="update")
    assert [got[t] for t in ("P20_USERNAME", "P21_USERINITIALS", "P22_USERADDRESS")] == ["OLDUSER"] * 3
    word = word_results()
    assert {t: got[t] for t in word if not t.startswith(("P20", "P21", "P22"))} == {
        t: ("OLD" if v.startswith("Error!") else v) for t, v in word.items() if not t.startswith(("P20", "P21", "P22"))
    }


def test_a_user_given_by_name_or_mapping(zone):
    zone(SYDNEY)
    pkg = load(PROBE / "field-probe-r1.docx")
    by_name = results(pkg, fields="update", field_user="Jo Bloggs")
    assert [by_name[t] for t in ("P20_USERNAME", "P21_USERINITIALS", "P22_USERADDRESS")] == ["Jo Bloggs", "", ""]
    mapping = results(pkg, fields="update", field_user={"name": "Jo", "initials": "JB", "address": "1 Street"})
    assert [mapping[t] for t in ("P20_USERNAME", "P21_USERINITIALS", "P22_USERADDRESS")] == ["Jo", "JB", "1 Street"]


def test_cached_the_probe_says_old_everywhere(zone):
    zone(SYDNEY)
    got = results(load(PROBE / "field-probe-r1.docx"))
    assert set(got.values()) == {"OLD", "OLDUSER"}


def test_words_save_read_cached_is_results_txt():
    """The oracle file and the record agree: Word's save holds what ``results.txt`` says Word showed."""
    assert results(load(PROBE / "field-probe-r1-word365.docx")) == word_results()


@pytest.mark.parametrize(
    ("source", "out", "options", "package_zone"),
    [
        ("field-probe-r1.docx", "docx4j-update.txt", {"fields": "update"}, None),
        ("field-probe-r1.docx", "docx4j-update-tokyo.txt", {"fields": "update"}, "Asia/Tokyo"),
        ("field-probe-r1.docx", "docx4j-cached.txt", {"fields": "cached"}, None),
        ("field-probe-r1-word365.docx", "docx4j-word365-cached.txt", {"fields": "cached"}, None),
    ],
)
def test_docx4js_own_extraction_byte_for_byte(zone, source, out, options, package_zone):
    """docx4j's ``Texts.plainText`` at ``c3e9920e5``, the JVM in Sydney's zone, against this one in the same."""
    zone(SYDNEY)
    pkg = load(PROBE / source)
    if package_zone is not None:
        pkg.date_time_zone = package_zone
    assert plain_text(extract(pkg, **options)) == (PROBE / out).read_text(encoding="utf-8")


def test_docx4js_markdown_byte_for_byte(zone):
    zone(SYDNEY)
    pkg = load(PROBE / "field-probe-r1.docx")
    assert markdown(extract(pkg, fields="update")) == (PROBE / "docx4j-update.md").read_text(encoding="utf-8")


def test_the_write_back_only_results_are_not_extractions(zone):
    """A ``REF``'s bold (case 13), ``NOTEREF \\f``'s style (18): Word's save has them, an extraction does not.

    ``FieldUpdater`` writes them (docx4j's ``setResultRuns`` and the
    ``FootnoteReference`` run style); the text model gives text, and this
    repository has no ``FieldUpdater``.
    """
    zone(SYDNEY)
    updated = markdown(extract(load(PROBE / "field-probe-r1.docx"), fields="update"))
    assert "P13\\_REF = Bookmarked text\n" in updated, "the computed REF result is plain"
    assert "A bookmark: **Bookmarked text**." in updated, "the bookmark itself is bold"
    word = load(PROBE / "field-probe-r1-word365.docx")
    words = markdown(extract(word))
    assert "P13\\_REF = **Bookmarked text**" in words, "Word's result carries the bookmark's bold"
    # NOTEREF \f: Word's result run is in the FootnoteReference style; the text model's is the field's text
    from docx4j_py.model.text import extract as read

    block = next(b for b in read(word).blocks if b.text.startswith("P18_NOTEREF_F"))
    assert block.fields[0].result == "1"
    xml = word.main_document_part.get_xml()
    assert re.search(r'P18_NOTEREF_F = .*?<w:rStyle w:val="FootnoteReference"/>', xml, re.DOTALL)
    assert "P16_REF_H = Bookmarked text" in plain_text(read(word)), "\\h changes nothing in the result's text"
