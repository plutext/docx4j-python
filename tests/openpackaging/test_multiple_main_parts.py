"""A package with more than one ``officeDocument`` relationship. WD REQ-206's statement, ruled 2026-11-07; CR-032 section 11.

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
from docx4j_py.openpackaging import LoadOptions, MultipleMainParts, MultipleMainPartsException
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


def test_one_relationship_says_nothing() -> None:
    pkg = load(_package("word/document.xml"))
    assert plain_text(extract(pkg)).strip() == "DOCUMENT" and not pkg.warnings
