"""CR-022 D1: the limits table, held to Word 365's saves and to the model.

The table is ``docx4j_py/resources/integer-limits.json``, hand-written from docx4j CR-066's
phase 0, rounds 5 and 6 of the shared probe (CR-022 sections 3.1 and 3.3). The fixtures are
``tests/fixtures/integer-limits/`` (``README.md``, ``SHA1SUMS``): every group file of both
rounds, the D1 and D2 singles, and Word's save of each. What is held here:

* every case of both rounds is cited by exactly one row, or by the table's ``unmeasured``
  list, but for the controls (``Z-1``) and CR-021's digitless forms (round 6's group ``E``);
* each row agrees with Word's saves: **a case inside the row's range is kept, one outside
  is changed** (wrapped, clamped, its rule flipped, or dropped), boundaries included;
* each row's element binds, in the model, only classes that type its attribute an integer
  (the way CR-021's canon table, :func:`docx4j_py.openpackaging.normalise._integer_tables`,
  reads the model), and the holder a row names is needed;
* no 32- or 64-bit attribute has a row: the ids, ``w:numId``, DrawingML's EMUs. Only
  ``w:pgNumType/@w:start`` is 32 bits (section 3.1).
"""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from functools import cache
from importlib.resources import files
from pathlib import Path

import pytest
from lxml import etree

from docx4j_py.model.values import (
    IntegerLimit,
    integer,
    integer_limit,
    integer_limits,
    unmeasured_attributes,
)

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "integer-limits"
ROUNDS = ("round5", "round6")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

#: The sibling attributes whose change means Word read the value otherwise (a rule flipped).
RULES = ("lineRule", "hRule", "type")
#: Attributes Word reads in a transitional document and writes back under another name.
WRITTEN_BACK = {("w:ind", "w:start"): "w:left", ("w:ind", "w:end"): "w:right"}


def _q(name: str) -> str:
    return f"{{{W}}}{name.removeprefix('w:')}"


@cache
def _resource() -> dict:
    """The table as written: its rows before they are expanded, and its unmeasured list."""
    return json.loads(files("docx4j_py.resources").joinpath("integer-limits.json").read_text("utf-8"))


# ---------------------------------------------------------------------------
# the fixtures
# ---------------------------------------------------------------------------


def test_the_fixtures_are_as_copied() -> None:
    """Every file ``SHA1SUMS`` lists, byte for byte, and nothing else beside it."""
    listed = {}
    for line in (FIXTURES / "SHA1SUMS").read_text(encoding="utf-8").splitlines():
        digest, name = line.split("  ", 1)
        listed[name] = digest
    for name, digest in listed.items():
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name
    present = sorted(
        p.relative_to(FIXTURES).as_posix() for p in FIXTURES.rglob("*") if p.is_file()
    )
    assert present == sorted([*listed, "README.md", "SHA1SUMS"])
    assert len(listed) == 72


@cache
def cases(round_name: str) -> dict[str, tuple[str, str, str | None, str]]:
    """A round's cases by id: (group, attribute as cases.txt spells it, place or None, value)."""
    out = {}
    lines = (FIXTURES / round_name / "cases.txt").read_text(encoding="utf-8").splitlines()
    for line in lines[1:]:
        columns = line.split("\t")
        if round_name == "round5":
            case, group, attribute, value = columns[:4]
            place = None
        else:
            case, group, attribute, place, value = columns[:5]
        out[case] = (group, attribute, place, value)
    return out


def test_the_rounds_hold_the_cases_cr_022_counts() -> None:
    """Round 5 is 191 cases (190 and the control); round 6, 489 (481, and CR-021's eight forms)."""
    assert len(cases("round5")) == 191
    assert len(cases("round6")) == 489


def _parse(attribute: str) -> tuple[list[str], str, str | None]:
    """``w:pBdr/w:top/@w:sz`` as (["w:pBdr", "w:top"], "w:sz", None); ``...[exact]``'s context too."""
    found = re.fullmatch(r"(.*)/@(w:\w+)(?:\[(\w+)\])?", attribute)
    assert found is not None, attribute
    path, name, context = found.groups()
    return path.split("/"), name, context


# ---------------------------------------------------------------------------
# reading a case in a file
# ---------------------------------------------------------------------------


@cache
def _xml(path: Path, part: str) -> etree._Element:
    return etree.fromstring(zipfile.ZipFile(path).read(part))


def _files(round_name: str, group: str, case: str) -> tuple[Path, Path]:
    """The case's source file and Word's save of it."""
    if round_name == "round5":
        stem = f"cr066-{group}"
    elif group == "D":
        stem = f"cr066-r6-{case}"
    else:
        stem = f"cr066-r6-{group}"
    return FIXTURES / round_name / f"{stem}.docx", FIXTURES / round_name / f"{stem}-word365.docx"


def _label(body: etree._Element, case: str) -> tuple[list[etree._Element], int]:
    children = list(body)
    for index, child in enumerate(children):
        if child.tag == _q("p") and "".join(child.itertext()).startswith(f"{case} "):
            return children, index
    raise AssertionError(f"no paragraph labelled {case}")


def _scope(path: Path, round_name: str, case: str) -> etree._Element:
    """The element a case's value is under: a paragraph's ``w:pPr``, a run's ``w:rPr``, a table, a section."""
    group, attribute, place, _value = cases(round_name)[case]
    if place == "settings":
        return _xml(path, "word/settings.xml")
    children, index = _label(_xml(path, "word/document.xml").find(_q("body")), case)
    label = children[index]
    steps = _parse(attribute)[0]
    if place is None:  # round 5: by group, and the N group by attribute
        kind = {"P": "pPr", "B": "pPr", "R": "rPr", "T": "tbl", "S": "sectPr"}.get(group)
        if kind is None:
            kind = (
                "tbl" if steps[0] == "w:trHeight"
                else "sectPr" if steps[0] in ("w:pgMar", "w:pgSz")
                else "rPr" if steps[0] in ("w:sz", "w:w")
                else "pPr"
            )
    else:
        kind = place
    if kind == "pPr":
        return label.find(_q("pPr"))
    if kind == "rPr":
        for run in label.iter(_q("r")):
            r_pr = run.find(_q("rPr"))
            if r_pr is not None and r_pr.find(_q("rStyle")) is not None:
                return r_pr
        raise AssertionError(f"no ProbeRun run in {case}")
    if kind == "tbl":
        return next(child for child in children[index + 1 :] if child.tag == _q("tbl"))
    if kind == "sectPr":
        for child in children[index:]:
            found = child.find(f"{_q('pPr')}/{_q('sectPr')}")
            if found is not None:
                return found
        raise AssertionError(f"no section after {case}")
    raise AssertionError(f"no scope for {case} ({kind})")


def _element(path: Path, round_name: str, case: str) -> etree._Element | None:
    steps = _parse(cases(round_name)[case][1])[0]
    found = _scope(path, round_name, case).xpath(".//" + "/".join(steps), namespaces=NS)
    return found[0] if found else None


def _attribute(element: etree._Element | None, name: str) -> str | None:
    return None if element is None else element.get(_q(name))


def _reading(row: IntegerLimit, round_name: str, case: str) -> tuple[int | str | None, tuple[str | None, ...]]:
    """What Word's save holds for the case, as Word read it: the value, and the rules beside it."""
    group, _spelled, _place, _value = cases(round_name)[case]
    _source, save = _files(round_name, group, case)
    element = _element(save, round_name, case)
    name = WRITTEN_BACK.get((row.element, row.name), row.name)
    text = _attribute(element, name)
    rules = tuple(_attribute(element, rule) for rule in RULES)
    if text is None:
        return None, rules
    if row.sentinel is not None and integer(text) == row.sentinel:
        return "sentinel", rules
    for stored, spelling in row.word_writes:
        if text == spelling:
            return stored, rules
    return integer(text), rules


def _cited() -> list[tuple[IntegerLimit, str, str]]:
    return [(row, name, case) for row in integer_limits() for name, ids in row.cases for case in ids]


# ---------------------------------------------------------------------------
# the table against the probe
# ---------------------------------------------------------------------------


def test_every_case_is_cited_once() -> None:
    """Each case of both rounds is one row's, or an unmeasured attribute's; but the controls and group E."""
    cited: dict[tuple[str, str], int] = {}
    for _row, name, case in _cited():
        cited[(name, case)] = cited.get((name, case), 0) + 1
    unmeasured = set()
    for entries in unmeasured_attributes().values():
        assert entries
    for entry in _resource()["unmeasured"]:
        for name, ids in entry["cases"].items():
            unmeasured.update((name, case) for case in ids)
    # a row naming several attributes cites its cases once, under its first attribute's row
    expected = {
        (name, case)
        for name in ROUNDS
        for case, (group, *_rest) in cases(name).items()
        if group not in ("Z", "E")
    }
    assert set(cited) | unmeasured == expected
    assert not set(cited) & unmeasured
    # each case once: a row of several attributes repeats its cases on each, and only there
    for (name, case), count in cited.items():
        rows = [row for row, n, c in _cited() if (n, c) == (name, case)]
        assert len({row.attribute for row in rows}) == count, (name, case)
        assert len({(row.minimum, row.maximum, row.storage) for row in rows}) == 1, (name, case)


@pytest.mark.parametrize(("row", "round_name", "case"), _cited(), ids=lambda v: getattr(v, "attribute", v))
def test_a_cited_case_is_of_its_row(row: IntegerLimit, round_name: str, case: str) -> None:
    """The case sets the row's element and attribute, under the row's context."""
    _group, attribute, _place, value = cases(round_name)[case]
    steps, name, context = _parse(attribute)
    assert name == row.name
    if row.element not in ("w:top", "w:bdr"):  # the border rows name the edge; the case its family
        assert steps[-1] == row.element
    else:
        assert steps[-1] in ("w:top", "w:bdr")
    source, _save = _files(round_name, *cases(round_name)[case][:1], case)
    element = _element(source, round_name, case)
    assert _attribute(element, name) == value, "the source carries the case's value"
    if row.context is not None:
        # the rule the source states (round 6's cases.txt brackets only the rules it varies)
        stated = _attribute(element, row.context[0])
        assert stated in row.context[1], f"{case}: {row.context[0]}={stated!r}"
        assert context is None or context == stated
    else:
        assert context is None


def _row_cases() -> list[tuple[IntegerLimit, list[tuple[str, str]]]]:
    seen, out = set(), []
    for row in integer_limits():
        key = (row.minimum, row.maximum, row.cases)
        if not row.cases or key in seen:
            continue
        seen.add(key)
        out.append((row, [(name, case) for name, ids in row.cases for case in ids]))
    return out


@pytest.mark.parametrize(("row", "cited"), _row_cases(), ids=lambda v: getattr(v, "attribute", None) or "")
def test_words_saves_keep_each_value_inside_the_range_and_change_each_outside(
    row: IntegerLimit, cited: list[tuple[str, str]]
) -> None:
    """Word's reading of each cited case: kept where the row admits the value, changed where it does not.

    Kept means the save holds the value written (as Word spells it: 65535 as ``"-1"`` where
    ``word_writes`` says so) under the same rule; changed, anything else: wrapped,
    clamped, the rule flipped (``exact`` to ``atLeast``, ``w:firstLine`` to
    ``w:hanging``), Word's sentinel, or dropped. A value ``word_writes`` names is the
    one exception, by choice (CR-022 section 10, decision 3): Word keeps it, but writes
    it back as another number (65535 as ``"-1"``), so the row leaves it out.
    """
    inside = outside = 0
    for round_name, case in cited:
        group, _spelled, _place, value = cases(round_name)[case]
        written = int(value)
        source, _save = _files(round_name, group, case)
        before_rules = tuple(_attribute(_element(source, round_name, case), rule) for rule in RULES)
        reading, rules = _reading(row, round_name, case)
        kept = reading == written and rules == before_rules
        if written in dict(row.word_writes):
            assert kept, f"{case}: Word keeps {written}, written back as {dict(row.word_writes)[written]!r}"
            assert not row.admits(written), f"{case}: the row leaves {written} out, by choice"
            continue
        assert kept == row.admits(written), (
            f"{round_name} {case}: {written} read as {reading!r} {rules}, "
            f"where the row is {row.minimum} to {row.maximum}"
        )
        inside += kept
        outside += not kept
        if written in (row.minimum, row.maximum):
            assert kept, f"{case}: a boundary is kept"
        if written in (row.minimum - 1, row.maximum + 1):
            assert not kept, f"{case}: one past a boundary is changed"
    assert inside and outside, "each row rests on a value kept and a value changed"


def test_the_boundaries_the_probe_measured_are_the_rows() -> None:
    """Both ends kept and one past each changed, where the probe set those values (the 16-bit edges)."""
    measured = {
        ("w:ind/@w:left", None): ((-32768, 32767), (-32769, 32768)),
        ("w:spacing/@w:before", None): ((32768,), (-1, 65536)),
        ("w:spacing/@w:line", "exact"): ((32767, 32768), (65535, -1)),
        ("w:trHeight/@w:val", None): ((32767,), (32768,)),
        ("w:trHeight/@w:val", "exact"): ((32768,), (65535,)),
        ("w:rPr/w:sz/@w:val", None): ((65535,), (65536, -1)),
        ("w:pgSz/@w:w", None): ((31681,), (32767, 65536)),
        ("w:tab/@w:pos", None): ((-1,), (32767, -32768)),
        ("w:top/@w:sz", None): ((255,), (256, -1)),
        ("w:top/@w:space", None): ((31,), (32, -1)),
        ("w:ind/@w:firstLine", None): ((32767,), (32768, -1)),
        ("w:ind/@w:hanging", None): ((32767, 32768), (65535, -1)),
        ("w:pgNumType/@w:start", None): ((2147483647, -1), (2147483648,)),
    }
    for (attribute, context), (kept, changed) in measured.items():
        row = integer_limit(attribute, context)
        assert row is not None
        assert all(row.admits(v) for v in kept), attribute
        assert not any(row.admits(v) for v in changed), attribute


def test_where_word_writes_65535_as_minus_1_the_rows_stop_at_65534() -> None:
    """``w:before``, ``w:after`` and the five unsigned margins: Word keeps 0xFFFF but saves it as ``"-1"``.

    So 65,535 is refused there by choice (CR-022 section 10, decision 3, as docx4j-core-ts
    stops), and the rows end at 65,534; the unsigned attributes Word writes 65535 back as
    written (``w:sz``, ``w:szCs``, ``w:kern``, the text scale, ``w:defaultTabStop``,
    ``w:hyphenationZone``) keep 0 to 65,535.
    """
    spelled = sorted(row.attribute for row in integer_limits() if row.word_writes)
    assert spelled == sorted(
        ["w:spacing/@w:before", "w:spacing/@w:after"]
        + [f"w:pgMar/@w:{edge}" for edge in ("left", "right", "header", "footer", "gutter")]
    )
    for row in integer_limits():
        if row.word_writes:
            assert dict(row.word_writes) == {65535: "-1"}
            assert (row.minimum, row.maximum) == (0, 65534), row.attribute
    for attribute in (
        "w:rPr/w:sz/@w:val",
        "w:szCs/@w:val",
        "w:kern/@w:val",
        "w:w/@w:val",
        "w:defaultTabStop/@w:val",
        "w:hyphenationZone/@w:val",
    ):
        row = integer_limit(attribute)
        assert row is not None and (row.minimum, row.maximum) == (0, 65535), attribute
    # and Word's saves agree: 65535 written back as "-1" in the one family, as written in the other
    for name, case, saved in (("round5", "P1-3", "-1"), ("round6", "A3-3", "-1"), ("round6", "S1-3", "-1"),
                              ("round5", "R1-5", "65535"), ("round6", "R2-3", "65535"), ("round6", "D1-3", "65535")):
        group, attribute, _place, value = cases(name)[case]
        assert value == "65535"
        _source, save = _files(name, group, case)
        assert _attribute(_element(save, name, case), _parse(attribute)[1]) == saved, case


def test_the_unmeasured_attributes_are_what_word_did_with_them() -> None:
    """``hRule="auto"`` heights, ``w:wBefore`` and ``w:wAfter`` dropped; ``w:gridCol`` recomputed; ``w:fitText`` kept."""
    by_attribute = {entry["attribute"]: entry for entry in _resource()["unmeasured"]}
    assert set(by_attribute) == {
        "w:trHeight/@w:val",
        "w:gridCol/@w:w",
        "w:wBefore/@w:w",
        "w:wAfter/@w:w",
        "w:legacy/@w:legacySpace",
        "w:legacy/@w:legacyIndent",
        "w:fitText/@w:val",
    }
    for attribute, entry in by_attribute.items():
        if attribute.startswith("w:legacy"):
            numbering = _xml(FIXTURES / "round6" / "cr066-r6-N-word365.docx", "word/numbering.xml")
            assert numbering.find(f".//{_q('legacy')}") is None, "Word 365 drops w:legacy"
            continue
        for name, ids in entry["cases"].items():
            for case in ids:
                group, case_attribute, _place, value = cases(name)[case]
                _source, save = _files(name, group, case)
                saved = _attribute(_element(save, name, case), _parse(case_attribute)[1])
                if attribute == "w:fitText/@w:val":
                    assert saved == value, case
                elif attribute == "w:gridCol/@w:w":
                    assert saved != value, case
                else:
                    assert saved is None, case
    assert integer_limit("w:trHeight/@w:val", "auto") is None
    assert integer_limit("w:gridCol/@w:w") is None


# ---------------------------------------------------------------------------
# the table against the model
# ---------------------------------------------------------------------------


def _qname(prefixed: str) -> str:
    return f"{{{W}}}{prefixed.removeprefix('w:')}"


@pytest.mark.parametrize("row", integer_limits(), ids=lambda row: f"{row.attribute}{row.context or ''}")
def test_each_row_names_an_integer_attribute_of_every_class_its_element_binds(row: IntegerLimit) -> None:
    """Every class the element name binds (under the row's holder, where it names one) that declares the
    attribute types it an integer, and declares the row's context attribute; at least one declares it.

    As CR-021's canon table reads the model (Q5): a class that does not declare the attribute
    (``w:top`` in ``w:tcMar`` for a border's ``w:sz``) does not count. A row that names a holder needs
    it: without the holder, some class bound to the element types the attribute otherwise.
    """
    from docx4j_py.openpackaging.normalise import _bound, integer_attributes
    from docx4j_py.runtime import context

    bound, holders, broken = _bound()
    element, attribute = _qname(row.element), _qname(row.name)
    classes = {cls for cls in bound.get(element, ()) if cls not in broken}
    assert classes, f"{row.element} binds no class"

    def under(holder: str | None) -> set[type]:
        if holder is None:
            return classes
        names_of: dict[type, set[str]] = {}
        for name, bound_classes in bound.items():
            for cls in bound_classes:
                names_of.setdefault(cls, set()).add(name)
        return {
            cls
            for cls in classes
            if any(_qname(holder) in names_of.get(owner, ()) for owner in holders.get((element, cls), ()))
        }

    def declaring(of: set[type]) -> dict[type, object]:
        out = {}
        for cls in of:
            for var in context().build(cls).get_attribute_vars():
                if var.qname == attribute:
                    out[cls] = var
        return out

    found = declaring(under(row.holder))
    assert found, f"no class bound to {row.element} declares {row.name}"
    for cls, var in found.items():
        integers = {v.qname for v, word in integer_attributes(cls) if word}
        assert var.qname in integers, f"{cls.__name__}.{var.name} is not an integer attribute of w"
        if row.context is not None:
            names = {v.qname for v in context().build(cls).get_attribute_vars()}
            assert _qname(row.context[0]) in names, f"{cls.__name__} has no {row.context[0]}"
    if row.holder is not None:
        # a holder is there because the element alone binds a class typing the attribute otherwise
        # (w:frame/w:sz is text), or because the probe measured the attribute under that holder only
        # (a cell margin's edge, Q3), which the row's note says
        everywhere = declaring(classes)
        typed_otherwise = any(
            v.qname not in {w.qname for w, word in integer_attributes(cls) if word} for cls, v in everywhere.items()
        )
        assert typed_otherwise or "only" in (row.note or ""), f"{row.attribute}: the holder {row.holder} is not needed"


def test_the_border_rows_name_only_border_edges() -> None:
    """``w:sz`` and ``w:space`` as attributes are a border's alone, wherever the edge's name binds."""
    from docx4j_py.openpackaging.normalise import _bound

    bound, _holders, _broken = _bound()
    for row in integer_limits():
        if row.name not in ("w:sz", "w:space") or row.element in ("w:cols", "w:col"):
            continue
        for cls in bound[_qname(row.element)]:
            if any(f.name == row.name.removeprefix("w:") for f in cls.__dataclass_fields__.values()):
                assert "Border" in cls.__name__, (row.attribute, cls.__name__)


#: The wide integer attributes a 16-bit row would refuse legitimate values of (CR-022 section 3).
WIDE = (
    "w:comment/@w:id",
    "w:commentRangeStart/@w:id",
    "w:commentReference/@w:id",
    "w:bookmarkStart/@w:id",
    "w:bookmarkEnd/@w:id",
    "w:ins/@w:id",
    "w:del/@w:id",
    "w:numId/@w:val",
    "w:num/@w:numId",
    "w:abstractNum/@w:abstractNumId",
    "w:abstractNumId/@w:val",
    "w:footnoteReference/@w:id",
    "w:id/@w:val",
)


def test_no_32_or_64_bit_attribute_has_a_row() -> None:
    """The ids, ``w:numId`` and DrawingML's EMUs have no row; every row fits 16 bits but ``w:pgNumType/@w:start``."""
    attributes = {row.attribute for row in integer_limits()}
    for wide in WIDE:
        assert wide not in attributes
        with pytest.raises(LookupError):
            integer_limit(wide)
    for row in integer_limits():
        assert all(step.startswith("w:") for step in (*row.path, row.name)), "WordprocessingML only: no EMU"
        if row.attribute == "w:pgNumType/@w:start":
            assert (row.minimum, row.maximum) == (-(2**31), 2**31 - 1)
            continue
        assert -32768 <= row.minimum <= row.maximum <= 65535, row.attribute
    for emu in ("wp:extent/@cx", "a:ext/@cx", "wp:positionH/wp:posOffset", "a:off/@x"):
        with pytest.raises(LookupError):
            integer_limit(emu)


def test_the_table_is_what_cr_022_counts() -> None:
    """60 rows in the resource, 82 attribute paths and contexts once expanded; seven unmeasured."""
    data = _resource()
    assert len(data["rows"]) == 60
    assert len(integer_limits()) == 82
    assert len(data["unmeasured"]) == 7
    keys = [(row.attribute, row.context) for row in integer_limits()]
    assert len(keys) == len(set(keys)), "one row per attribute and context"
    by_attribute: dict[str, list[IntegerLimit]] = {}
    for row in integer_limits():
        by_attribute.setdefault(row.attribute, []).append(row)
    for attribute, rows in by_attribute.items():
        if len(rows) > 1:  # a range by context: the contexts do not overlap
            contexts = [value for row in rows for value in row.context[1]]  # type: ignore[index]
            assert len(contexts) == len(set(contexts)), attribute
