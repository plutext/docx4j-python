"""``pkg.write_times_on_save``: the save's time in the core properties, off by default. CR-025 Phase A, D4.

The portfolio's date rule's point 6 (2026-10-28): docx4j's ``docx4j.dcterms.write``
(``DocPropsCorePart.writeTimes()``, 17.3.3, its CR-036) is the shape every port
offers, off by default, under docx4j-core-ts's name (``writeTimesOnSave``). On,
``dcterms:modified`` is the save's time, and ``dcterms:created`` too where the
core part has none and was not read from the source; UTC with a ``Z``, to the
second, as docx4j and core-ts write them (Q3). A package with no core part gets
none. Off, a save of a loaded file writes its core part byte for byte.
"""

from __future__ import annotations

import datetime
import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

import docx4j_py.openpackaging.save as save_module
from docx4j_py import create_package, load

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "samples"
FIXTURES = ROOT / "tests" / "fixtures"
CORE = "docProps/core.xml"
DCTERMS = "{http://purl.org/dc/terms/}"
XSI_TYPE = "{http://www.w3.org/2001/XMLSchema-instance}type"
SAVED_AT = datetime.datetime(2026, 10, 28, 3, 12, 45, 678901, tzinfo=datetime.UTC)


@pytest.fixture
def clock(monkeypatch):
    """The save's time fixed at :data:`SAVED_AT`."""
    monkeypatch.setattr(save_module, "_now", lambda: SAVED_AT)


def _core(data: bytes) -> bytes:
    return zipfile.ZipFile(io.BytesIO(data)).read(CORE)


def _dates(core: bytes) -> dict[str, tuple[str | None, str | None]]:
    root = etree.fromstring(core)
    out = {}
    for name in ("created", "modified"):
        found = root.find(DCTERMS + name)
        out[name] = (found.text, found.get(XSI_TYPE)) if found is not None else (None, None)
    return out


def _without_dates(core: bytes) -> list[tuple[str, str | None]]:
    root = etree.fromstring(core)
    return sorted((child.tag, child.text) for child in root if child.tag not in (DCTERMS + "created", DCTERMS + "modified"))


def test_off_by_default_a_loaded_core_part_is_saved_byte_for_byte(clock):
    path = SAMPLES / "2010-sample1.docx"
    source = zipfile.ZipFile(path).read(CORE)
    pkg = load(path)
    assert pkg.write_times_on_save is False
    assert _core(pkg.save()) == source
    assert pkg.last_save.times is None
    assert not pkg.doc_props_core_part.is_unmarshalled


def test_on_a_save_sets_modified_to_its_time_and_keeps_created_and_the_rest(clock):
    path = SAMPLES / "2010-sample1.docx"
    source = zipfile.ZipFile(path).read(CORE)
    pkg = load(path)
    pkg.write_times_on_save = True
    saved = _core(pkg.save())
    assert _dates(saved) == {
        "created": _dates(source)["created"],
        "modified": ("2026-10-28T03:12:45Z", "dcterms:W3CDTF"),
    }
    assert _without_dates(saved) == _without_dates(source), "nothing else in the part changes"
    assert pkg.last_save.times == "2026-10-28T03:12:45Z"


def test_on_every_save_writes_its_own_time(monkeypatch):
    pkg = create_package()
    pkg.write_times_on_save = True
    for when in ("2026-10-28T03:12:45Z", "2026-10-29T00:00:01Z"):
        monkeypatch.setattr(save_module, "_now", lambda when=when: datetime.datetime.fromisoformat(when))
        assert _dates(_core(pkg.save()))["modified"][0] == when
        assert pkg.last_save.times == when


def test_a_created_package_keeps_create_packages_created(clock):
    pkg = create_package()
    created = pkg.doc_props_core_part.contents.created
    pkg.write_times_on_save = True
    dates = _dates(_core(pkg.save()))
    assert dates["created"][0] == str(created.value)
    assert dates["modified"] == ("2026-10-28T03:12:45Z", "dcterms:W3CDTF")


def test_a_core_part_made_here_with_no_created_gets_one(clock):
    """docx4j: ``created`` too "where the package is new and has none"; here, where the part was not read."""
    pkg = create_package()
    pkg.doc_props_core_part.contents.created = None
    pkg.write_times_on_save = True
    assert _dates(_core(pkg.save())) == {
        "created": ("2026-10-28T03:12:45Z", "dcterms:W3CDTF"),
        "modified": ("2026-10-28T03:12:45Z", "dcterms:W3CDTF"),
    }


def test_a_loaded_core_part_with_no_created_gets_none(clock):
    """``2010-glow-then-AlternateContent.docx`` has a core part with no ``dcterms:created``: it stays without."""
    pkg = load(SAMPLES / "2010-glow-then-AlternateContent.docx")
    pkg.write_times_on_save = True
    dates = _dates(_core(pkg.save()))
    assert dates["created"] == (None, None)
    assert dates["modified"][0] == "2026-10-28T03:12:45Z"


def test_a_package_with_no_core_part_gets_none(clock):
    """``comments.docx`` has no ``docProps/core.xml``; docx4j writes the times only into a part it has."""
    pkg = load(FIXTURES / "comments.docx")
    pkg.write_times_on_save = True
    data = pkg.save()
    assert CORE not in zipfile.ZipFile(io.BytesIO(data)).namelist()
    assert pkg.last_save.times is None


def test_the_opt_in_reloads_as_the_dates_written(clock):
    pkg = load(SAMPLES / "2010-sample1.docx")
    pkg.write_times_on_save = True
    again = load(io.BytesIO(pkg.save()))
    from docx4j_py.model.text import extract

    assert extract(again).metadata["core"]["modified"] == "2026-10-28T03:12:45Z"
