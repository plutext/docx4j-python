"""The property resolver against docx4j's own answers. CR-002 §8, §15.1.

Every golden of ``../docx4j-core-ts/test/golden/`` whose fixture is here, over

* ``styles``: the default paragraph style id, the document defaults
  (``w:pPr`` and ``w:rPr``, and the ``w:docDefaults`` element as the part states
  it), and per style id ``effectivePPr(styleId)`` and ``effectiveRPr(styleId)``;
* per paragraph of every story: ``effectivePPr`` and ``paragraphMarkRPr``;
* per run: ``effectiveRPr``;
* ``tables``: the effective table style and whether the chain reaches the
  default table style.

Everything is compared **as an object tree**: docx4j marshalled its answers
with its own prefixes, so both sides go through one unmarshal and ours through
a marshal first, and neither prefixes nor attribute order nor the class a
member happens to be is part of the contract (CR-002 §12.12).

A comparison named in ``goldens.PINS`` is expected to differ until the goldens
are regenerated (CR-009 Phase A: two, where this package already merges colours
as docx4j 17.3.3 does); its (test, golden) pair is ``xfail(strict=True)`` and
every other comparison of the golden is still held to zero differences.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import goldens
import pytest
from goldens import (
    W_PPR,
    W_RPR,
    W_STYLE,
    Comparisons,
    children_of,
    each_golden,
    fixture_for,
    read,
    runs_of,
    story,
    walk,
)

from docx4j_py import load
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.wml import Tbl

pytestmark = pytest.mark.skipif(
    not goldens.GOLDENS.is_dir(),
    reason=f"the docx4j parity goldens are not checked out at {goldens.GOLDENS}",
)


def _resolver(golden: Path) -> tuple[Any, Any, dict[str, Any]]:
    recorded = read(golden)
    package = load(fixture_for(golden), options=LoadOptions(mce_preprocess=True))
    return package, package.get_property_resolver(), recorded


@each_golden
def test_the_styles_resolve_as_docx4j_resolves_them(
    golden: Path, request: pytest.FixtureRequest
) -> None:
    """The document defaults and every style id in the part."""
    _package, resolver, recorded = _resolver(golden)
    check = Comparisons(request.node.originalname, golden)
    styles = recorded["styles"]

    assert resolver.default_paragraph_style_id == styles["defaultParagraphStyleId"], (
        f"{golden.stem}: default paragraph style"
    )
    check.same(
        resolver.get_document_default_ppr(),
        styles["documentDefaults"]["pPr"],
        f"{golden.stem}: document default w:pPr",
        name=W_PPR,
    )
    check.same(
        resolver.get_document_default_rpr(),
        styles["documentDefaults"]["rPr"],
        f"{golden.stem}: document default w:rPr",
        name=W_RPR,
    )

    for style_id, want in styles["byId"].items():
        style = resolver.get_style(style_id)
        assert style is not None, f"{golden.stem}: no style {style_id!r}"
        assert str(getattr(style.type_value, "value", style.type_value)) == want["type"]
        assert getattr(style.based_on, "val", None) == want["basedOn"]
        assert bool(style.default) == want["isDefault"]
        check.same(
            resolver.get_effective_ppr(style_id),
            want["effectivePPr"],
            f"{golden.stem}: effectivePPr({style_id!r})",
            name=W_PPR,
        )
        check.same(
            resolver.get_effective_rpr(style_id),
            want["effectiveRPr"],
            f"{golden.stem}: effectiveRPr({style_id!r})",
            name=W_RPR,
        )
    check.finish()


@each_golden
def test_every_paragraph_and_run_resolves_as_docx4j_resolves_it(
    golden: Path, request: pytest.FixtureRequest
) -> None:
    """``effectivePPr``, ``paragraphMarkRPr`` and every run's ``effectiveRPr``."""
    package, resolver, recorded = _resolver(golden)
    check = Comparisons(request.node.originalname, golden)
    compared = 0

    for key, recorded_story in recorded["stories"].items():
        root, _part = story(package, key)
        found: list[Any] = []
        walk(children_of(root), None, found)
        expected = recorded_story["paragraphs"]
        assert len(found) == len(expected), (
            f"{golden.stem}/{key}: {len(found)} paragraphs, docx4j walked {len(expected)}"
        )

        for (paragraph, _state), want in zip(found, expected, strict=True):
            where = f"{golden.stem}/{key}[{want['index']}] {want['text']!r}"
            p_pr = paragraph.p_pr
            check.same(
                resolver.get_effective_ppr(p_pr),
                want["effectivePPr"],
                f"{where}: effectivePPr",
                name=W_PPR,
            )
            check.same(
                resolver.get_effective_paragraph_mark_rpr(p_pr),
                want["paragraphMarkRPr"],
                f"{where}: paragraphMarkRPr",
                name=W_RPR,
            )
            compared += 1

            runs = runs_of(paragraph)
            assert len(runs) == len(want["runs"]), (
                f"{where}: {len(runs)} runs, docx4j walked {len(want['runs'])}"
            )
            for index, (run, run_want) in enumerate(zip(runs, want["runs"], strict=True)):
                check.same(
                    resolver.get_effective_rpr(run.r_pr, p_pr),
                    run_want["effectiveRPr"],
                    f"{where} run {index}: effectiveRPr",
                    name=W_RPR,
                )
                compared += 1

    assert compared > 0
    check.finish()


@each_golden
def test_every_table_resolves_as_docx4j_resolves_it(
    golden: Path, request: pytest.FixtureRequest
) -> None:
    """The effective table style of every ``w:tbl`` of the main story."""
    package, resolver, recorded = _resolver(golden)
    check = Comparisons(request.node.originalname, golden)
    tables = _tables_of(package.main_document_part.contents.body)
    assert len(tables) == len(recorded["tables"]), (
        f"{golden.stem}: {len(tables)} tables, docx4j found {len(recorded['tables'])}"
    )
    for table, want in zip(tables, recorded["tables"], strict=True):
        where = f"{golden.stem}: table {want['index']}"
        tbl_pr = table.tbl_pr
        assert getattr(getattr(tbl_pr, "tbl_style", None), "val", None) == want["tblStyle"]
        check.same(
            resolver.get_effective_table_style(tbl_pr),
            want["effectiveTableStyle"],
            f"{where}: effectiveTableStyle",
            name=W_STYLE,
        )
        assert resolver.reaches_default_table_style(tbl_pr) == want[
            "reachesDefaultTableStyle"
        ], f"{where}: reachesDefaultTableStyle"
    check.finish()


def _tables_of(root: Any) -> list[Any]:
    """Every ``w:tbl`` of the main story in document order, nested ones included."""
    from docx4j_py.child import iter_children

    out: list[Any] = []

    def collect(obj: Any, depth: int) -> None:
        if depth > 60:  # pragma: no cover
            return
        for _name, child in iter_children(obj):
            if isinstance(child, Tbl):
                out.append(child)
            if goldens.is_text_box(child):
                continue
            collect(child, depth + 1)

    collect(root, 0)
    return out
