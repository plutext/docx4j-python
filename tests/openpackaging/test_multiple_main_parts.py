"""A package with more than one ``officeDocument`` relationship. WD REQ-206's statement, ruled 2026-11-07; CR-032 section 11.

And an ``External`` one beside an internal one (REQ-206's NOTE, WD 0.102; CR-032 section 11.5), refused at every
``repair`` but ``all``, which reads the internal target and drops the ``External`` one; with ``LoadOptions(repair=)``.

Refused by default, naming the targets, as a duplicated entry is and for the same reason: Word reads such a package only
through a repair. Under the same switch, ``LoadOptions(duplicate_entries="first")``, the **last** relationship's target is
read, as a fresh Word reads it, and ``package.multiple_main_parts`` raised once.

Word's own saves aren't the oracle: its choice alternates with each repaired open in one session (the WD's ``open-order``
round), so the `pw-two-main-rels-*` saves read the first or the last by where they fell. Every zip is built here.
"""

from __future__ import annotations

import io
import logging
import zipfile
from typing import Any

import pytest

from docx4j_py import create_package, load
from docx4j_py.model.text import TextWarning, extract, plain_text
from docx4j_py.openpackaging import (
    ExternalMainPart,
    ExternalMainPartException,
    LoadOptions,
    MultipleMainParts,
    MultipleMainPartsException,
)
from docx4j_py.openpackaging.exceptions import InvalidFormatException

MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
FIRST = LoadOptions(duplicate_entries="first")


def _package(*targets: str, types: tuple[str, ...] = (), orphan: str | None = None) -> bytes:
    """A package whose main part is each of `targets`, in that order, holding its own name's stem as its words."""
    pkg = create_package()
    pkg.body.insert_paragraph("WORDS")
    out = io.BytesIO()
    pkg.save(out)
    made = zipfile.ZipFile(io.BytesIO(out.getvalue()))
    document = made.read("word/document.xml")
    entries = {name: made.read(name) for name in made.namelist() if name != "word/document.xml"}
    rels = entries["_rels/.rels"].decode()
    head, _, rest = rels.partition("<Relationships")
    opening, _, body = rest.partition(">")
    kept = [rel for rel in body.replace("</Relationships>", "").split("/>") if rel.strip() and RELATIONSHIP not in rel]
    made_rels = [
        f'<Relationship Id="rIdMain{i}" Type="{(types[i] if i < len(types) else RELATIONSHIP)}" Target="{target}"'
        for i, target in enumerate(targets)
    ]
    entries["_rels/.rels"] = (
        head + "<Relationships" + opening + ">" + "".join(f"{rel}/>" for rel in [*kept, *made_rels]) + "</Relationships>"
    ).encode()
    named = [*targets, *([orphan] if orphan else [])]
    types_part = entries["[Content_Types].xml"].decode()
    overrides = "".join(f'<Override PartName="/{name}" ContentType="{MAIN}"/>' for name in named if name != "word/document.xml")
    entries["[Content_Types].xml"] = types_part.replace("</Types>", overrides + "</Types>").encode()
    for name in named:
        entries[name] = document.replace(b"WORDS", name.split("/")[-1].split(".")[0].upper().encode())
    if "word/document.xml" not in named:
        entries["[Content_Types].xml"] = entries["[Content_Types].xml"].replace(
            f'<Override PartName="/word/document.xml" ContentType="{MAIN}"/>'.encode(), b""
        )
    rels_of_main = entries.pop("word/_rels/document.xml.rels", None)
    if rels_of_main is not None:
        for name in named:
            folder, _, file = name.rpartition("/")
            entries[f"{folder}/_rels/{file}.rels"] = rels_of_main
    built = io.BytesIO()
    with zipfile.ZipFile(built, "w") as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    return built.getvalue()


def _saved(pkg: Any) -> zipfile.ZipFile:
    out = io.BytesIO()
    pkg.save(out)
    return zipfile.ZipFile(io.BytesIO(out.getvalue()))


TWO = ("word/document.xml", "word/other.xml")


def test_two_relationships_are_refused_naming_both() -> None:
    with pytest.raises(MultipleMainPartsException) as caught:
        load(_package(*TWO))
    assert isinstance(caught.value, InvalidFormatException)
    assert caught.value.code == "package.multiple_main_parts"
    assert caught.value.targets == ("/word/document.xml", "/word/other.xml")
    assert "/word/document.xml, /word/other.xml" in str(caught.value) and "REQ-206" in str(caught.value)


def test_a_strict_and_a_transitional_relationship_are_two() -> None:
    with pytest.raises(MultipleMainPartsException):
        load(_package(*TWO, types=(RELATIONSHIP, STRICT)))


def test_strict_refuses_as_the_default_does() -> None:
    with pytest.raises(MultipleMainPartsException):
        load(_package(*TWO), options=LoadOptions(strict=True))


@pytest.mark.parametrize(("targets", "word"), [(TWO, "OTHER"), (TWO[::-1], "DOCUMENT")])
def test_under_the_switch_the_last_relationship_s_target_is_read(targets: tuple[str, ...], word: str) -> None:
    assert plain_text(extract(load(_package(*targets), options=FIRST))).strip() == word


def test_the_order_of_three_decides_and_they_are_counted() -> None:
    pkg = load(_package("word/alpha.xml", "word/zeta.xml", "word/beta.xml"), options=FIRST)
    assert plain_text(extract(pkg)).strip() == "BETA"
    assert pkg.warnings == [MultipleMainParts("/word/beta.xml", 3)]


def test_every_route_to_the_main_part_agrees() -> None:
    pkg = load(_package(*TWO), options=FIRST)
    assert str(pkg.main_document_part.part_name) == "/word/other.xml"
    assert pkg.get_main_part() is pkg.main_document_part
    assert "/word/document.xml" not in {str(part.part_name) for part in pkg.parts.parts()}


def test_the_warning_comes_once_and_is_logged(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        pkg = load(_package(*TWO), options=FIRST)
    [record] = pkg.warnings
    assert record.code == "package.multiple_main_parts"
    assert record.to_dict() == {"code": record.code, "part_name": "/word/other.xml", "count": 2, "message": record.message}
    assert [item.getMessage() for item in caplog.records if "officeDocument" in item.getMessage()] == [record.message]


def test_the_text_model_raises_it_once_with_the_part_and_the_count() -> None:
    [warning] = extract(load(_package(*TWO), options=FIRST)).warnings
    assert warning.kind == TextWarning.PACKAGE_MULTIPLE_MAIN_PARTS == "package.multiple_main_parts"
    assert (warning.address, warning.selector, warning.detail) == (None, "/word/other.xml", 2)


def test_strict_doesn_t_refuse_under_the_switch() -> None:
    assert load(_package(*TWO), options=LoadOptions(strict=True, duplicate_entries="first")).warnings


def test_a_save_holds_one_main_part_and_loads_with_nothing_said() -> None:
    saved = _saved(load(_package(*TWO), options=FIRST))
    assert "word/other.xml" in saved.namelist() and "word/document.xml" not in saved.namelist()
    assert saved.read("_rels/.rels").count(b"/officeDocument\"") == 1
    assert b"/word/document.xml" not in saved.read("[Content_Types].xml")
    out = io.BytesIO()
    load(_package(*TWO), options=FIRST).save(out)
    again = load(out.getvalue())
    assert plain_text(extract(again)).strip() == "OTHER" and not again.warnings


def test_an_unreferenced_main_part_is_an_orphan_and_no_case() -> None:
    """One relationship and an Override for a part nothing names: it's never read, and a save drops it, as Word's does."""
    data = _package("word/document.xml", orphan="word/orphan.xml")
    pkg = load(data)
    assert plain_text(extract(pkg)).strip() == "DOCUMENT" and not pkg.warnings
    saved = _saved(pkg)
    assert "word/orphan.xml" not in saved.namelist()


def test_two_relationships_to_one_target_are_still_two() -> None:
    """The statement counts relationships. Under the switch the one part is read, and its content type stays."""
    data = _package(*TWO)
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as zout:
        for name in zin.namelist():
            zout.writestr(name, zin.read(name).replace(b'Target="word/other.xml"', b'Target="word/document.xml"'))
    with pytest.raises(MultipleMainPartsException) as caught:
        load(out.getvalue())
    assert caught.value.targets == ("/word/document.xml", "/word/document.xml")
    pkg = load(out.getvalue(), options=FIRST)
    assert plain_text(extract(pkg)).strip() == "DOCUMENT"
    assert b"/word/document.xml" in _saved(pkg).read("[Content_Types].xml")


def _with_external(*, first: bool, targets: tuple[str, ...] = ("word/document.xml",)) -> bytes:
    """A package whose relationships also hold an ``External`` ``officeDocument`` one, first or last."""
    data = _package(*targets)
    external = f'<Relationship Id="rIdOut" Type="{RELATIONSHIP}" Target="http://example.org/x.docx" TargetMode="External"/>'.encode()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as zout:
        for name in zin.namelist():
            held = zin.read(name)
            if name == "_rels/.rels":
                at = held.index(b"<Relationship ") if first else held.index(b"</Relationships>")
                held = held[:at] + external + held[at:]
            zout.writestr(name, held)
    return out.getvalue()


URL = "http://example.org/x.docx"
ALL = LoadOptions(repair="all")


@pytest.mark.parametrize("first", [True, False], ids=["first", "last"])
@pytest.mark.parametrize(
    "options",
    [None, FIRST, LoadOptions(repair="word"), LoadOptions(strict=True)],
    ids=["default", "duplicate_entries=first", "repair=word", "strict"],
)
def test_an_external_relationship_beside_an_internal_one_is_refused_naming_the_url(first: bool, options: Any) -> None:
    """Word refuses it, offering a repair and then failing (x-ext-*, 8 of 8): REQ-206's NOTE, WD 0.102."""
    with pytest.raises(ExternalMainPartException) as caught:
        load(_with_external(first=first), options=options)
    assert caught.value.code == "package.external_main_part"
    assert (caught.value.urls, caught.value.targets) == ((URL,), ("/word/document.xml",))
    assert URL in str(caught.value) and 'LoadOptions(repair="all")' in str(caught.value)
    assert isinstance(caught.value, InvalidFormatException)


@pytest.mark.parametrize("first", [True, False], ids=["first", "last"])
def test_under_repair_all_the_internal_target_is_read_and_the_external_one_dropped(first: bool) -> None:
    pkg = load(_with_external(first=first), options=ALL)
    assert plain_text(extract(pkg)).strip() == "DOCUMENT"
    assert pkg.warnings == [ExternalMainPart("/word/document.xml", (URL,))]
    assert pkg.get_main_part() is pkg.main_document_part
    assert [r.target for r in pkg.relationships_part.list if r.type_value == RELATIONSHIP] == ["word/document.xml"]


def test_strict_doesn_t_refuse_under_repair_all() -> None:
    """The option is the caller's choice, as for ``duplicate_entries="first"`` (CR-032 Q4)."""
    assert load(_with_external(first=True), options=LoadOptions(strict=True, repair="all")).warnings


def test_the_record_comes_once_and_is_logged(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        pkg = load(_with_external(first=True), options=ALL)
    [record] = pkg.warnings
    assert record.code == "package.external_main_part" and record.detail == URL
    assert record.to_dict() == {"code": record.code, "part_name": "/word/document.xml", "urls": [URL], "message": record.message}
    assert [item.getMessage() for item in caplog.records if "External" in item.getMessage()] == [record.message]


def test_the_text_model_raises_it_once_and_the_result_says_the_repair_is_its_own() -> None:
    result = extract(load(_with_external(first=False), options=ALL))
    [warning] = result.warnings
    assert warning.kind == TextWarning.PACKAGE_EXTERNAL_MAIN_PART == "package.external_main_part"
    assert (warning.address, warning.selector, warning.detail) == (None, "/word/document.xml", URL)
    assert result.metadata["repaired"] == "own"


def test_the_url_is_never_fetched(monkeypatch: pytest.MonkeyPatch) -> None:
    """REQ-206's NOTE: MUST NOT fetch the URL. No socket is opened, at any setting."""
    import socket

    def refuse(*args: Any, **kwargs: Any) -> None:
        raise AssertionError("a socket was opened")

    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.setattr(socket, "create_connection", refuse)
    monkeypatch.setattr(socket, "getaddrinfo", refuse)
    with pytest.raises(ExternalMainPartException):
        load(_with_external(first=True))
    pkg = load(_with_external(first=True), options=ALL)
    assert plain_text(extract(pkg)).strip() == "DOCUMENT"
    _saved(pkg)


@pytest.mark.parametrize("first", [True, False], ids=["first", "last"])
def test_a_save_under_repair_all_holds_no_external_relationship_and_loads_with_nothing_said(first: bool) -> None:
    saved = _saved(load(_with_external(first=first), options=ALL))
    rels = saved.read("_rels/.rels")
    assert URL.encode() not in rels and b"External" not in rels
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for name in saved.namelist():
            zout.writestr(name, saved.read(name))
    again = load(out.getvalue())
    assert plain_text(extract(again)).strip() == "DOCUMENT" and not again.warnings


def test_every_external_one_goes_and_the_detail_is_the_first_as_docx4j_s_is() -> None:
    second = f'<Relationship Id="rIdOut2" Type="{RELATIONSHIP}" Target="https://example.org/y.docx" TargetMode="External"/>'
    data = _with_external(first=True)
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as zout:
        for name in zin.namelist():
            held = zin.read(name)
            if name == "_rels/.rels":
                held = held.replace(b"</Relationships>", second.encode() + b"</Relationships>")
            zout.writestr(name, held)
    with pytest.raises(ExternalMainPartException) as caught:
        load(out.getvalue())
    assert caught.value.urls == (URL, "https://example.org/y.docx")
    pkg = load(out.getvalue(), options=ALL)
    [record] = pkg.warnings
    assert record.urls == (URL, "https://example.org/y.docx") and record.detail == URL
    assert extract(pkg).warnings[0].detail == URL
    assert b"External" not in _saved(pkg).read("_rels/.rels")


def test_two_internal_ones_beside_it_are_refused_for_the_external_one_first() -> None:
    """Unmeasured in Word; the External rule is asked first, as docx4j's ``Load3.oneMainPart`` asks it."""
    for options in (None, FIRST, LoadOptions(repair="word")):
        with pytest.raises(ExternalMainPartException) as caught:
            load(_with_external(first=False, targets=TWO), options=options)
        assert caught.value.targets == ("/word/document.xml", "/word/other.xml")
    pkg = load(_with_external(first=False, targets=TWO), options=ALL)
    assert pkg.warnings == [ExternalMainPart("/word/other.xml", (URL,)), MultipleMainParts("/word/other.xml", 2)]
    assert plain_text(extract(pkg)).strip() == "OTHER"
    assert extract(pkg).metadata["repaired"] == "own"


@pytest.mark.parametrize("options", [None, ALL], ids=["default", "repair=all"])
def test_only_an_external_relationship_is_no_main_part_at_all(options: Any) -> None:
    """With no internal one there is no main part, which REQ-321 refuses at every setting."""
    data = _with_external(first=True)
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as zout:
        for name in zin.namelist():
            zout.writestr(name, zin.read(name).replace(b'Id="rIdMain0" Type="' + RELATIONSHIP.encode() + b'"', b'Id="rIdMain0" Type="urn:other"'))
    with pytest.raises(InvalidFormatException, match="No relationship of type officeDocument"):
        load(out.getvalue(), options=options)


# ---------------------------------------------------------------------------
# repair (WD REQ-207, 0.102)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("repair", ["word", "all"])
def test_repair_word_and_all_imply_duplicate_entries_first(repair: str) -> None:
    pkg = load(_package(*TWO), options=LoadOptions(repair=repair))
    assert pkg.warnings == [MultipleMainParts("/word/other.xml", 2)]
    result = extract(pkg)
    assert plain_text(result).strip() == "OTHER" and result.metadata["repaired"] == "word"


def test_nothing_repaired_says_none() -> None:
    assert extract(load(_package("word/document.xml"))).metadata["repaired"] == "none"
    assert extract(load(_package("word/document.xml"), options=ALL)).metadata["repaired"] == "none"


@pytest.mark.parametrize("value", ["none", "first", "", "ALL"])
def test_another_repair_value_is_refused(value: str) -> None:
    with pytest.raises(ValueError, match="REQ-207"):
        LoadOptions(repair=value)


def test_one_relationship_says_nothing() -> None:
    pkg = load(_package("word/document.xml"))
    assert plain_text(extract(pkg)).strip() == "DOCUMENT" and not pkg.warnings
