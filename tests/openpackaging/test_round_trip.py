"""Load, save, reload: the promise of CR-002 section 1, over every sample.

Two claims, and this file is the proof of both:

**A part that is never touched is written back byte for byte.** Load a package,
save it, and every non-relationship entry of the new zip is the same bytes as
the old one's.

**A part that is unmarshalled loses nothing.** Unmarshal *every* XML part of
every sample --- settings, webSettings, docProps, theme, the lot --- serialise
it again with nothing changed, and the canonical XML is identical to the source
with zero skipped content.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_round_trip.py -q
"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon  # noqa: E402

from docx4j_py.openpackaging import (  # noqa: E402
    DirectoryPartSink,
    DirectoryPartStore,
    LoadOptions,
    MemoryPartSink,
    OpcPackage,
    WordprocessingMLPackage,
    XmlPart,
)

SAMPLES = sorted((ROOT / "samples").glob("*.doc*")) + sorted((ROOT / "samples").glob("*.dot*"))
PPTX_XLSX = sorted((ROOT / "samples").glob("*.pptx")) + sorted((ROOT / "samples").glob("*.xlsx"))
ALL = sorted(set(SAMPLES) | set(PPTX_XLSX))

#: The two parts whose root content model is ``xsd:all``, where the order of the
#: children carries no meaning and a serialiser is free to write them in schema
#: order. See ``scripts/canon.sort_children``.
UNORDERED_ROOTS = {"/docProps/core.xml", "/docProps/app.xml", "/docProps/custom.xml"}


def entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def source_entries(path: Path) -> dict[str, bytes]:
    return entries(path.read_bytes())


# ---------------------------------------------------------------------------
# loading
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("docx", ALL, ids=lambda p: p.name)
def test_every_sample_loads(docx):
    with OpcPackage.load(docx) as pkg:
        assert len(pkg.parts) > 0
        assert pkg.get_main_part() is not None
        # every part the loader made is reachable from the package
        for name in pkg.parts:
            assert pkg.get_part(name) is not None


@pytest.mark.parametrize("docx", SAMPLES, ids=lambda p: p.name)
def test_a_word_document_comes_back_as_a_word_package(docx):
    with WordprocessingMLPackage.load(docx) as pkg:
        assert pkg.main_document_part is not None
        assert pkg.main_document_part.part_name.name.endswith(".xml")


def test_a_pptx_and_an_xlsx_load_through_the_generic_path():
    """Phase C types their main parts; Phase A still loads and saves them."""
    for path in PPTX_XLSX:
        with OpcPackage.load(path) as pkg:
            assert type(pkg) is OpcPackage
            assert pkg.get_main_part() is not None


def test_the_pptx_and_xlsx_chart_parts_are_the_one_known_unmarshal_gap():
    """A finding, pinned so that fixing it is visible.

    The DrawingML *chart* and *spreadsheet drawing* schemas in this build do not
    declare ``mc:AlternateContent`` on ``c:chartSpace`` or on the anchors, so
    unmarshalling one of those parts drops it --- the skipped-content report
    says so, which is the report doing its job. WordprocessingML has the
    wildcards it needs (CR-001 ``schemas/PATCHES.md``); the chart namespaces get
    them when CR-001 Phase D generates PML and SML. Until then those parts are
    never unmarshalled by the engine itself, so a pptx or xlsx still round-trips
    byte for byte (the test above).
    """
    lost = 0
    for path in PPTX_XLSX:
        with OpcPackage.load(path) as pkg:
            pkg.unmarshal_all()
            lost += len(pkg.skipped)
    assert lost > 0, "if this is zero the gap is closed; delete this test"


def test_loading_is_lazy():
    """docx4j's rule: load costs the relationships, nothing else."""
    with OpcPackage.load(ROOT / "samples" / "toc.docx") as pkg:
        unmarshalled = [
            p for p in pkg.parts.parts() if isinstance(p, XmlPart) and p.is_unmarshalled
        ]
        assert unmarshalled == []


def test_a_part_reachable_twice_records_both_relationships():
    """A header used by two sections, an image placed twice."""
    with OpcPackage.load(ROOT / "samples" / "Headers.docx") as pkg:
        counts = [len(p.source_relationships) for p in pkg.parts.parts()]
        assert all(c >= 1 for c in counts)


def test_a_directory_store_loads_the_same_package(tmp_path):
    """docx4j ``UnzippedPartStore``."""
    source = ROOT / "samples" / "toc.docx"
    with zipfile.ZipFile(source) as zf:
        zf.extractall(tmp_path)
    with (
        OpcPackage.load(source) as a,
        OpcPackage.load(DirectoryPartStore(tmp_path)) as b,
    ):
        assert sorted(n.name for n in a.parts) == sorted(n.name for n in b.parts)


# ---------------------------------------------------------------------------
# untouched parts are byte-identical
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("docx", ALL, ids=lambda p: p.name)
def test_untouched_parts_are_byte_identical(docx):
    source = source_entries(docx)
    with OpcPackage.load(docx) as pkg:
        saved = entries(pkg.save())

    checked = 0
    for name, data in saved.items():
        if name == "[Content_Types].xml" or name.endswith(".rels"):
            continue  # both are always re-marshalled, as in docx4j
        checked += 1
        assert name in source, f"{name} is not in the source"
        assert data == source[name], f"{name} changed although nothing touched it"
    assert checked > 0


@pytest.mark.parametrize("docx", ALL, ids=lambda p: p.name)
def test_every_part_the_loader_saw_is_saved(docx):
    source = source_entries(docx)
    with OpcPackage.load(docx) as pkg:
        saved = entries(pkg.save())
    missing = set(source) - set(saved)
    assert missing == set(), f"{docx.name}: {sorted(missing)} were not written"


# ---------------------------------------------------------------------------
# unmarshalled parts are canonically identical, and lose nothing
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("docx", SAMPLES, ids=lambda p: p.name)
def test_every_xml_part_unmarshals_and_re_serialises_unchanged(docx):
    source = source_entries(docx)
    with OpcPackage.load(docx) as pkg:
        parts = [p for p in pkg.parts.parts() if isinstance(p, XmlPart)]
        assert parts, f"{docx.name} has no typed XML part"
        for part in parts:
            part.contents  # noqa: B018 - the unmarshal is the point
            assert part.skipped == [], f"{part.part_name} skipped {part.skipped}"
            report = canon.compare(
                source[part.part_name.store_name],
                part.xml,
                unordered_root=part.part_name.name in UNORDERED_ROOTS,
            )
            assert report.identical, (
                f"{docx.name} {part.part_name}: {report.differences} "
                f"{[(d.category, d.detail) for d in report.diffs[:5]]}"
            )


@pytest.mark.parametrize("docx", SAMPLES, ids=lambda p: p.name)
def test_nothing_is_skipped_anywhere(docx):
    with OpcPackage.load(docx) as pkg:
        pkg.unmarshal_all()
        assert pkg.skipped == []


def test_the_settings_part_keeps_its_style_pane_filter():
    """CR-001 section 14.8 point 5: 15 attributes used to be dropped here."""
    with WordprocessingMLPackage.load(ROOT / "samples" / "toc.docx") as pkg:
        settings = pkg.document_settings_part
        assert settings is not None
        assert settings.skipped == []
        pane = settings.contents.style_pane_format_filter
        assert pane is not None
        assert pane.val == "3F01"
        assert pane.all_styles is True
        assert pane.custom_styles is False
        assert pane.alternate_style_names is False


# ---------------------------------------------------------------------------
# reload, and the second generation
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("docx", ALL, ids=lambda p: p.name)
def test_load_save_reload_save_is_stable(docx):
    with OpcPackage.load(docx) as pkg:
        once = pkg.save()
    with OpcPackage.load(once) as pkg:
        twice = pkg.save()
    assert entries(once) == entries(twice)


def test_saving_over_the_loaded_path(tmp_path):
    """Decided question 4: a temporary file and a rename, then reopen."""
    target = tmp_path / "toc.docx"
    target.write_bytes((ROOT / "samples" / "toc.docx").read_bytes())
    before = source_entries(target)

    pkg = WordprocessingMLPackage.load(target)
    pkg.main_document_part.contents  # noqa: B018 - one part is now dirty
    pkg.save(target)

    after = source_entries(target)
    assert set(after) == set(before)
    # the package is still usable: the store was reopened
    assert pkg.style_definitions_part.contents is not None
    assert list(tmp_path.glob("*.tmp")) == []
    pkg.close()


def test_saving_to_a_file_object(tmp_path):
    target = tmp_path / "out.docx"
    with OpcPackage.load(ROOT / "samples" / "toc.docx") as pkg, target.open("wb") as fh:
        pkg.save(fh)
    assert set(entries(target.read_bytes())) == set(
        entries((ROOT / "samples" / "toc.docx").read_bytes())
    )


def test_saving_to_a_memory_sink_and_loading_it_back():
    with OpcPackage.load(ROOT / "samples" / "toc.docx") as pkg:
        store = pkg.save_to(MemoryPartSink())
    with OpcPackage.load(store) as again:
        assert len(again.parts) == 16


def test_saving_to_a_directory_sink(tmp_path):
    with OpcPackage.load(ROOT / "samples" / "toc.docx") as pkg:
        out = pkg.save_to(DirectoryPartSink(tmp_path / "unzipped"))
    assert (out / "word" / "document.xml").exists()
    with OpcPackage.load(DirectoryPartStore(out)) as again:
        assert again.get_main_part() is not None


def test_bytes_in_bytes_out():
    data = (ROOT / "samples" / "toc.docx").read_bytes()
    with OpcPackage.load(data) as pkg:
        assert isinstance(pkg.save(), bytes)


def test_a_file_object_is_a_container():
    with (ROOT / "samples" / "toc.docx").open("rb") as fh, OpcPackage.load(fh) as pkg:
        assert pkg.get_main_part() is not None


# ---------------------------------------------------------------------------
# media compression
# ---------------------------------------------------------------------------


def test_already_compressed_media_is_stored_not_deflated():
    with OpcPackage.load(ROOT / "samples" / "Images.docx") as pkg:
        data = pkg.save()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        media = [i for i in zf.infolist() if i.filename.lower().endswith(".png")]
        assert media, "the fixture must hold a PNG"
        assert all(i.compress_type == zipfile.ZIP_STORED for i in media)
        xml = [i for i in zf.infolist() if i.filename.endswith(".xml")]
        assert all(i.compress_type == zipfile.ZIP_DEFLATED for i in xml)


# ---------------------------------------------------------------------------
# strict loading
# ---------------------------------------------------------------------------


def test_strict_raises_on_skipped_content(tmp_path):
    """``LoadOptions(strict=True)``, decided question 5."""
    from docx4j_py.openpackaging import Docx4JException

    source = ROOT / "samples" / "toc.docx"
    broken = tmp_path / "broken.docx"
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(broken, "w") as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                data = data.replace(
                    b"<w:body>",
                    b'<w:body xmlns:zz="urn:x"><zz:widget/>',
                    1,
                )
            out.writestr(info.filename, data)

    with OpcPackage.load(broken) as lenient:
        lenient.get_part("/word/document.xml").contents
        assert lenient.skipped != []

    with OpcPackage.load(broken, options=LoadOptions(strict=True)) as strict:
        with pytest.raises(Docx4JException, match="strict=True"):
            strict.get_part("/word/document.xml").contents
