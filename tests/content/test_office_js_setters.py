"""CR-024: the content API's spacing, line, list indent and font size setters, held to Word 365's own.

Core-ts's check 71b (its CR-002 section 65.9, ``b7c2680``) set each of Office JS's
``spaceBefore``, ``spaceAfter``, ``lineSpacing``, ``List.setLevelIndents``' text indent and
``font.size`` at and around where Word 365's own setter stops, in Word 365 16.0.20326, and
recorded what Word refused (``InvalidArgument``), what it read back and what it wrote.
Its run, ``office_js/71b-setter-limits-word365-run.json``, is copied unedited
(``SHA1SUMS``). Each of its 51 cases runs here through the setter of the same name: refused
(``MeasureRangeError``, nothing changed) exactly where Word refused, and otherwise the same
value read back and the same attributes written. One difference, as core-ts lists it: Word
writes ``spaceBefore = 0`` as nothing where the style already gives 0, and this setter writes
``w:before="0"``.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from docx4j_py import create_package
from docx4j_py.model.content import MeasureRangeError
from docx4j_py.wml import to_xml

RUN = Path(__file__).resolve().parents[1] / "office_js" / "71b-setter-limits-word365-run.json"
CASES = json.loads(RUN.read_text(encoding="utf-8"))["cases"]

#: Word's attribute names, as its fragments spell them, and the model's fields.
SPACING = {"before": "before", "after": "after", "line": "line", "lineRule": "line_rule"}

#: The cases where this setter writes what Word's does not, by case number.
DIFFERENCES = {1: "Word writes spaceBefore 0 as nothing where the style gives 0; here w:before=\"0\""}


def _attributes(fragment: str) -> dict[str, str]:
    return dict(re.findall(r'w:(\w+)="([^"]*)"', fragment))


def _set(case: dict[str, Any]) -> tuple[Any, Any, Any, Any]:
    """Run `case`'s setter on a fresh paragraph: the package, the paragraph, the list (or None), the error (or None)."""
    package = create_package()
    paragraph = package.body.insert_paragraph("x")
    lst = paragraph.start_new_list() if case["setter"] == "listIndent" else None
    before = to_xml(package.main_document_part.contents)
    try:
        if case["setter"] == "spaceBefore":
            paragraph.space_before = case["value"]
        elif case["setter"] == "spaceAfter":
            paragraph.space_after = case["value"]
        elif case["setter"] == "lineSpacing":
            paragraph.line_spacing = case["value"]
        elif case["setter"] == "fontSize":
            paragraph.font.size = case["value"]
        elif case["setter"] == "listIndent":
            lst.set_level_indents(0, case["value"], 0)
        else:  # pragma: no cover - a new setter in the run
            raise AssertionError(case["setter"])
    except MeasureRangeError as error:
        assert to_xml(package.main_document_part.contents) == before, "a refusal changes nothing"
        return package, paragraph, lst, error
    return package, paragraph, lst, None


def _level_ind(package: Any, lst: Any) -> Any:
    """The list's level 0 ``w:ind``, where ``set_level_indents`` writes it."""
    part = package.main_document_part.numbering_definitions_part
    num = next(n for n in part.contents.num if n.num_id == lst.id)
    abstract = next(a for a in part.contents.abstract_num if a.abstract_num_id == num.abstract_num_id.val)
    return abstract.lvl[0].p_pr.ind


def test_the_run_is_check_71bs_51_cases() -> None:
    assert len(CASES) == 51
    assert {case["setter"] for case in CASES} == {"spaceBefore", "spaceAfter", "lineSpacing", "listIndent", "fontSize"}
    assert sum("error" in case for case in CASES) == 31, "Word refused 31, the other 20 set"


@pytest.mark.parametrize("case", CASES, ids=[f"{c['n']}-{c['setter']}-{c['value']}" for c in CASES])
def test_each_setter_refuses_where_word_refuses_and_writes_what_word_writes(case: dict[str, Any]) -> None:
    package, paragraph, lst, error = _set(case)
    word_refused = "error" in case
    assert (error is not None) == word_refused, (case, error)
    if word_refused:
        assert error.code == "value.out_of_range"
        return
    read = case["after"]["read"]
    written = case["after"]["ooxml"]["document"]
    setter = case["setter"]
    if setter == "listIndent":
        ind = _level_ind(package, lst)
        assert (ind.left, ind.first_line, ind.hanging) == (round(case["value"] * 20), 0, None)
        assert read == {"leftIndent": pytest.approx(case["value"], abs=1e-3), "firstLineIndent": 0}
        return
    if setter == "fontSize":
        assert {_attributes(fragment)["val"] for fragment in written} == {str(paragraph.runs[0].r_pr.sz.val)}
        assert paragraph.font.size == pytest.approx(read, abs=1e-3)
        return
    spacing = paragraph.element.p_pr.spacing
    getter = {"spaceBefore": "space_before", "spaceAfter": "space_after", "lineSpacing": "line_spacing"}[setter]
    assert getattr(paragraph, getter) == pytest.approx(read, abs=1e-3)
    if case["n"] in DIFFERENCES:
        assert written == [] and spacing.before == 0, DIFFERENCES[case["n"]]
        return
    (fragment,) = written
    ours = {}
    for name, field in SPACING.items():
        value = getattr(spacing, field)
        if value is not None:
            ours[name] = str(getattr(value, "value", value))
    assert ours == _attributes(fragment), (case["n"], fragment, to_xml(spacing))


def test_only_the_listed_case_differs() -> None:
    assert set(DIFFERENCES) == {1}
    assert CASES[0]["setter"] == "spaceBefore" and CASES[0]["value"] == 0
