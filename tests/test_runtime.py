"""``warm_up``, the prefix table and the shared runtime, CR-001 sections 7 and 8.

.venv-fork/bin/python -m pytest tests/test_runtime.py -q
"""

from __future__ import annotations

import subprocess
import sys
from importlib import resources
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import namespaces as ns  # noqa: E402
from docx4j_py.runtime import (  # noqa: E402
    WARM_UP_PARTS,
    context,
    needs_preserve,
    parser,
    parser_config,
    serializer_config,
    warm_up,
)

# ---------------------------------------------------------------------------
# the prefix table
# ---------------------------------------------------------------------------


def test_the_prefixes_docx4j_names():
    """The ones CR-001 section 7 lists by name."""
    for prefix in (
        "w",
        "r",
        "wp",
        "a",
        "pic",
        "v",
        "o",
        "w10",
        "w14",
        "w15",
        "mc",
        "m",
    ):
        assert prefix in ns.PREFIXES
    for prefix in ("w16se", "w16cid", "w16", "wps", "wpg", "c", "dgm", "xml"):
        assert prefix in ns.PREFIXES
    assert ns.PREFIXES["w"] == ns.WML_NS
    assert ns.W_NAMESPACE_DECLARATION == f'xmlns:w="{ns.WML_NS}"'


def test_the_table_is_a_bijection():
    assert len(ns.NAMESPACE_TO_PREFIX) == len(ns.PREFIXES)
    for prefix, uri in ns.PREFIXES.items():
        assert ns.NAMESPACE_TO_PREFIX[uri] == prefix


def test_ns_map_leaves_out_xml():
    """`xml` is bound by definition; declaring it makes a document invalid."""
    assert "xml" not in ns.ns_map()
    assert ns.ns_map(default=ns.WML_NS)[None] == ns.WML_NS
    assert "xmlns:xml=" not in ns.declarations()


def test_mc_ignorable_prefixes_are_all_declared():
    assert ns.MC_IGNORABLE_PREFIXES <= set(ns.PREFIXES)
    # what Word writes over the samples/ corpus
    assert {"w14", "w15", "w16se", "w16cid", "wp14"} <= ns.MC_IGNORABLE_PREFIXES


def test_every_generated_namespace_has_a_prefix():
    """A part the engine writes must never need an invented prefix.

    Three of the generated namespaces are not in docx4j's own mapper --- MathML,
    InkML and `sharedTypes` --- so `docx4j_py.namespaces` adds them, and this is
    what keeps that list complete as Phase D generates more.
    """
    from docx4j_py.el_index import EL_MODULES

    missing = sorted(uri for uri in EL_MODULES if uri not in ns.NAMESPACE_TO_PREFIX)
    assert missing == []
    assert ns.NAMESPACE_TO_PREFIX["http://www.w3.org/1998/Math/MathML"] == "mml"
    assert ns.NAMESPACE_TO_PREFIX["http://www.w3.org/2003/InkML"] == "inkml"


def test_qname_helpers():
    assert ns.qname("w", "p") == f"{{{ns.WML_NS}}}p"
    assert ns.split_qname(f"{{{ns.WML_NS}}}p") == (ns.WML_NS, "p")
    assert ns.split_qname("p") == (None, "p")


# ---------------------------------------------------------------------------
# xml:space
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("", False),
        ("x", False),
        ("one space here", False),
        (" x", True),
        ("x ", True),
        ("a  b", True),
        ("\tx", True),
        ("x\n", True),
        ("a \t b", True),
    ],
)
def test_needs_preserve(text, expected):
    assert needs_preserve(text) is expected


# ---------------------------------------------------------------------------
# configuration
# ---------------------------------------------------------------------------


def test_parser_config_is_lenient_but_never_silent():
    config = parser_config()
    assert config.fail_on_unknown_properties is False
    assert config.fail_on_unknown_attributes is False
    assert config.skipped is not None
    strict = parser_config(lenient=False, skipped_report=False)
    assert strict.fail_on_unknown_properties is True
    assert strict.skipped is None


def test_a_parser_is_never_shared_but_the_context_is():
    assert parser() is not parser()
    assert parser().context is context() is parser().context


def test_booleans_are_numeric_as_word_writes_them():
    assert serializer_config().bool_format == "numeric"
    assert serializer_config().xml_declaration is False
    assert serializer_config(pretty=True).indent == "  "


# ---------------------------------------------------------------------------
# warm_up
# ---------------------------------------------------------------------------


def test_the_warm_up_parts_are_representative():
    """They have to hold what a real document holds, or they warm nothing."""
    document = (
        resources.files("docx4j_py.resources") / "warmup_document.xml"
    ).read_text()
    for marker in (
        "<w:tbl>",
        "<w:drawing>",
        "mc:AlternateContent",
        "w14:paraId",
        "<w:sdt>",
        "<w:hyperlink",
        "<w:ins ",
        "<w:del ",
        "<w:sym ",
        "<w:tab/>",
        "m:oMath",
    ):
        assert marker in document, marker
    styles = (resources.files("docx4j_py.resources") / "warmup_styles.xml").read_text()
    assert "<w:latentStyles" in styles and "<w:docDefaults>" in styles


def test_the_warm_up_parts_parse_with_nothing_skipped():
    """CR-001's promise applies to the embedded parts too."""
    for name, target in WARM_UP_PARTS:
        module, _, class_name = target.rpartition(".")
        from importlib import import_module

        clazz = getattr(import_module(module), class_name)
        data = (resources.files("docx4j_py.resources") / name).read_bytes()
        p = parser()
        p.from_bytes(data, clazz)
        assert list(p.skipped) == [], (name, list(p.skipped))


def test_warm_up_builds_the_metadata_and_is_idempotent():
    before = len(context().cache)
    spent = warm_up(force=True)
    after = len(context().cache)
    assert spent > 0
    assert after >= before
    assert after > 100, "the warm-up parts should build a hundred classes or so"
    assert warm_up() == 0.0  # already warm


@pytest.mark.slow
def test_warm_up_shortens_the_first_parse():
    """The measurement CR-001 section 14 records, in a fresh process each time.

    Cold: importing the model builds no class metadata at all, so the first
    document pays for every class it touches. Warm: `warm_up()` has paid for
    most of them already.
    """
    script = """
import sys, time, zipfile
sys.path.insert(0, {root!r})
import docx4j_py.wml as W
from docx4j_py.runtime import context, parser, warm_up
if {warm}:
    warm_up()
data = zipfile.ZipFile({root!r} + "/samples/tables.docx").read("word/document.xml")
built = len(context().cache)
start = time.perf_counter()
parser().from_bytes(data, W.Document)
print(f"{{time.perf_counter() - start:.4f}} {{built}} {{len(context().cache)}}")
"""
    results = {}
    for warm in (False, True):
        out = subprocess.run(
            [sys.executable, "-c", script.format(root=str(ROOT), warm=warm)],
            capture_output=True,
            text=True,
            check=True,
        )
        seconds, before, after = out.stdout.split()
        results[warm] = (float(seconds), int(before), int(after))

    cold_seconds, cold_before, cold_after = results[False]
    warm_seconds, warm_before, warm_after = results[True]

    assert cold_before == 0, "importing the model must build no metadata"
    assert warm_before > 100, "warm_up must build the metadata up front"
    # the first parse then has almost nothing left to build
    assert warm_after - warm_before < (cold_after - cold_before) / 4
    assert warm_seconds < cold_seconds
