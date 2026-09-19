"""The font selector and the mapper against docx4j's own answers. CR-002 §8, §15.3.

Every golden of ``../docx4j-core-ts/test/golden/`` whose fixture is here, over

* ``fonts.themePart`` and ``fonts.defaultFont``: which theme part the package has
  and the face its document defaults resolve to;
* ``fonts.fontsInUse`` and ``fonts.stylesInUse`` as **sorted lists**: docx4j's
  walk for names, over every story;
* ``fonts.mapping``: the ``IdentityPlusMapper`` decision per document font ---
  which pass answered, how, and the physical font;
* every run's ``fontSpans``: the text folded into maximal stretches of one
  document font, with the run's effective ``w:b``, ``w:i``, ``w:cs`` and
  ``w:rtl``.

The font environment is docx4j's four font jars and nothing else
(:func:`~docx4j_py.model.fonts.registry.default_font_registry`), which is what
the goldens were made in, so these answers do not depend on what this machine has
installed.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import goldens
import pytest
from goldens import (
    all_goldens,
    children_of,
    fixture_for,
    read,
    run_text,
    runs_of,
    story,
    walk,
)

from docx4j_py import load
from docx4j_py.openpackaging.load import LoadOptions

pytestmark = pytest.mark.skipif(
    not goldens.GOLDENS.is_dir(),
    reason=f"the docx4j parity goldens are not checked out at {goldens.GOLDENS}",
)


def _package(golden: Path) -> tuple[Any, dict[str, Any]]:
    recorded = read(golden)
    package = load(fixture_for(golden), options=LoadOptions(mce_preprocess=True))
    return package, recorded


@pytest.mark.parametrize("golden", all_goldens(), ids=lambda path: path.stem)
def test_the_fonts_a_document_uses_are_the_fonts_docx4j_finds(golden: Path) -> None:
    """``themePart``, ``defaultFont``, ``fontsInUse``, ``stylesInUse`` and the mapping."""
    package, recorded = _package(golden)
    want = recorded["fonts"]
    main = package.main_document_part

    theme_part = package.theme_part
    name = None if theme_part is None else str(theme_part.part_name)
    assert name == want["themePart"], f"{golden.stem}: theme part"

    selector = package.get_run_font_selector()
    assert selector.default_font == want["defaultFont"], f"{golden.stem}: default font"

    fonts = sorted(main.fonts_in_use())
    assert fonts == sorted(want["fontsInUse"]), f"{golden.stem}: fontsInUse"
    styles = sorted(main.styles_in_use())
    assert styles == sorted(want["stylesInUse"]), f"{golden.stem}: stylesInUse"

    mapper = package.get_font_mapper()
    ours = {}
    for font in fonts:
        decision = mapper.get_decision(font)
        ours[font] = None if decision is None else decision.to_dict()
    assert ours == want["mapping"], f"{golden.stem}: the IdentityPlusMapper decisions"


@pytest.mark.parametrize("golden", all_goldens(), ids=lambda path: path.stem)
def test_every_run_is_split_into_the_font_spans_docx4j_splits_it_into(golden: Path) -> None:
    """``fontSpans`` of every run of every story."""
    package, recorded = _package(golden)
    resolver = package.get_property_resolver()
    selector = package.get_run_font_selector()
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
            p_pr = paragraph.p_pr
            runs = runs_of(paragraph)
            assert len(runs) == len(want["runs"])
            for index, (run, run_want) in enumerate(zip(runs, want["runs"], strict=True)):
                where = f"{golden.stem}/{key}[{want['index']}] run {index}"
                effective = resolver.get_effective_rpr(run.r_pr, p_pr)
                spans = selector.spans(
                    p_pr, effective, run_text(run), r_pr_is_effective=True
                )
                ours = [span.to_dict() for span in spans]
                assert ours == run_want["fontSpans"], f"{where}: fontSpans"
                compared += len(ours)

    assert compared >= 0
