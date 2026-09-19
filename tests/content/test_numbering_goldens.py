"""The numbering emulator against docx4j's own answers. CR-003 §18.10, CR-002 §12.12.

The goldens are produced by the **Java** parity harness that lives in
``../docx4j-core-ts/test/java/``: it loads each fixture with docx4j itself and
writes down what docx4j answers, one JSON per ``.docx`` under
``../docx4j-core-ts/test/golden/``. CR-002 section 12.12 settled that the
harness lives there and that this repository **consumes** its committed
goldens, so the JSON is not copied here; the sibling checkout is the source and
these tests skip when it is absent.

What is compared, per paragraph of per story, in document order:

* ``numString`` --- the label Word paints;
* ``isBullet``, ``ilvl`` and ``numId``;
* ``numRef`` --- docx4j's own resolution (``Emulator.numRefFor``): the
  ``numId``, the ``ilvl``, whether the ``w:numPr`` was direct and whether the
  paragraph is numbered at all, against :meth:`Emulator.resolve`;
* ``stateAfter`` --- the story's counters once the paragraph has taken its
  number, keyed ``"<abstractNumId>/<ilvl>"`` with ``value``,
  ``encounteredAlready`` and ``resetPending``, and the ``w:num`` start
  overrides spent so far, keyed ``"<numId>/<ilvl>"``. All three counter fields
  are compared: this emulator has all three.

What is **not** compared, because it is CR-002 Phase B's (the property
resolver, the font selector) and this emulator does not claim it:
``effectivePPr``, ``paragraphMarkRPr``, ``runs`` and their ``fontSpans``, and
numbering's ``ind``, ``indResolved``, ``lvl``, ``labelRPr`` and ``numFont``.
``ind`` and ``lvl`` in particular come out here as an :class:`Indent` and a
:class:`ListLevel` rather than as typed ``w:ind`` and ``w:lvl`` elements
(CR-003 section 18.2 item 2), so comparing them is a Phase B job.

**The story rule** is docx4j's ``NumberingStates.forPart``, which
:class:`~docx4j_py.model.listnumbering.NumberingStates` already implements: the
main document is one story; **every header and footer of the document share
one**; the footnotes, endnotes and comments parts have one each; and a text box
is a fresh story whose paragraphs stay in the containing part's list, in
document order. The stories are walked in the order the golden lists them,
which is the harness's own order (main, then headers by relationship id, then
footers, then footnotes, endnotes, comments) --- it matters, because the
headers and footers share a state.

The walk here mirrors the harness's ``Walk``: descend into tables, content
controls, hyperlinks and revisions, and start a fresh state at a text box.
``mce_preprocess=True`` is what makes a text box's paragraphs typed at all
(CLAUDE.md), and it takes the branch the harness takes.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from conftest import FIXTURES, ROOT

from docx4j_py import load
from docx4j_py.child import iter_children
from docx4j_py.model.listnumbering import NumberingState, NumberingStates, emulator_of
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.wml import DelText, P, Text

#: The committed goldens of the sibling checkout. Not copied here: CR-002 §12.12.
GOLDENS = ROOT.parent / "docx4j-core-ts" / "test" / "golden"

#: Where a golden's fixture is under another name here. ``lists.docx`` is
#: ``numbering_indentation.docx`` byte for byte (tests/README.md), so the map
#: is the whole of the duplication this test needed.
FIXTURE_NAMES = {"numbering_indentation": "lists"}

#: The node whose children begin a story of their own. docx4j's ``isTextBox``
#: reaches this through ``wp:inline``/``wp:anchor`` → ``a:graphicData`` →
#: ``wps:wsp``; starting the new story at the ``wps:wsp`` itself gives every
#: paragraph below it the same fresh state. ``v:textbox`` is named for
#: completeness: VML has no typed classes in this object model, so a VML text
#: box's paragraphs are a wildcard tree and no fixture here has one.
TEXT_BOXES = frozenset({"Wsp", "Textbox", "CTTextbox"})

pytestmark = pytest.mark.skipif(
    not GOLDENS.is_dir(),
    reason=f"the docx4j parity goldens are not checked out at {GOLDENS}",
)


# ---------------------------------------------------------------------------
# which goldens this repository can answer
# ---------------------------------------------------------------------------


def _fixture_for(golden: Path) -> Path:
    return FIXTURES / f"{FIXTURE_NAMES.get(golden.stem, golden.stem)}.docx"


def _goldens() -> list[Path]:
    """Every golden whose fixture is in ``tests/fixtures/``."""
    if not GOLDENS.is_dir():
        return []
    return [path for path in sorted(GOLDENS.glob("*.json")) if _fixture_for(path).exists()]


# ---------------------------------------------------------------------------
# the walk: docx4j's Harness.Walk
# ---------------------------------------------------------------------------


def _is_text_box(obj: Any) -> bool:
    return type(obj).__name__ in TEXT_BOXES


def _walk(children: Any, state: NumberingState, out: list[Any], depth: int = 0) -> None:
    """Every ``w:p`` in document order, with the state its story counts in."""
    if depth > 60:  # pragma: no cover - the harness's own guard
        return
    for _name, child in children:
        if isinstance(child, P):
            out.append((child, state))
            _walk(iter_children(child), state, out, depth + 1)
        elif _is_text_box(child):
            _walk(iter_children(child), NumberingState(), out, depth + 1)
        else:
            _walk(iter_children(child), state, out, depth + 1)


def _story(package: Any, key: str) -> tuple[Any, Any]:
    """The content a golden's story key names, and the part it belongs to."""
    main = package.main_document_part
    if key == "main":
        return main.contents.body, main
    if key.startswith(("header:", "footer:")):
        part = main.relationships_part.get_part(key.split(":", 1)[1])
        return part.contents, part
    if key == "footnotes":
        return main.footnotes_part.contents.footnote, main.footnotes_part
    if key == "endnotes":
        return main.endnotes_part.contents.endnote, main.endnotes_part
    if key == "comments":
        return main.comments_part.contents.comment, main.comments_part
    raise AssertionError(f"unknown story {key!r}")  # pragma: no cover


def _children_of(root: Any) -> list[Any]:
    """A story root's children: a list of notes or comments is their contents."""
    if isinstance(root, list):
        out: list[Any] = []
        for item in root:
            out.extend(iter_children(item))
        return out
    return list(iter_children(root))


def _text_of(paragraph: Any) -> str:
    """The harness's ``textOf``: every ``w:t`` and ``w:delText``, no text box."""
    out: list[str] = []

    def collect(obj: Any, depth: int) -> None:
        if depth > 60:  # pragma: no cover - the harness's own guard
            return
        for _name, child in iter_children(obj):
            if isinstance(child, (Text, DelText)):
                out.append(child.value or "")
            elif _is_text_box(child) or isinstance(child, P):
                continue
            else:
                collect(child, depth + 1)

    collect(paragraph, 0)
    return "".join(out)


# ---------------------------------------------------------------------------
# our answers
# ---------------------------------------------------------------------------


def _value(obj: Any) -> Any:
    val = getattr(obj, "val", None)
    return getattr(val, "value", val)


def _num_ref(emulator: Any, paragraph: Any) -> Any:
    """:meth:`Emulator.resolve` over a paragraph, as ``lists._resolve`` calls it."""
    p_pr = getattr(paragraph, "p_pr", None)
    p_style = _value(getattr(p_pr, "p_style", None)) if p_pr is not None else None
    num_pr = getattr(p_pr, "num_pr", None) if p_pr is not None else None
    num_id = _value(getattr(num_pr, "num_id", None)) if num_pr is not None else None
    ilvl = _value(getattr(num_pr, "ilvl", None)) if num_pr is not None else None
    num_id = None if num_id is None else str(num_id)
    ilvl = None if ilvl is None else str(ilvl)
    return emulator.resolve(p_style, num_id, ilvl, bool(num_id))


def _check_num_ref(ours: Any, want: dict[str, Any], where: str) -> None:
    # a direct or style-contributed ``w:numId`` 0 no longer reaches here: since
    # docx4j 01d661547 (2026-09-19) it resolves as not numbered on both sides
    # and the golden carries ``numbering: null`` for the paragraph
    assert (ours.num_id, ours.ilvl, ours.direct, ours.not_numbered) == (
        want["numId"],
        want["ilvl"],
        want["direct"],
        want["notNumbered"],
    ), f"{where}: numRef {ours} != {want}"


def _check_state(state: NumberingState, want: dict[str, Any], where: str) -> None:
    counters = state.counters
    for key, expected in want["counters"].items():
        counter = counters.get(key)
        assert counter is not None, f"{where}: no counter {key}, docx4j has {expected}"
        assert str(counter.value) == expected["value"], (
            f"{where}: counter {key} value {counter.value} != {expected['value']}"
        )
        assert counter.encountered_already == expected["encounteredAlready"], (
            f"{where}: counter {key} encounteredAlready {counter.encountered_already} "
            f"!= {expected['encounteredAlready']}"
        )
        assert counter.reset_pending == expected["resetPending"], (
            f"{where}: counter {key} resetPending {counter.reset_pending} "
            f"!= {expected['resetPending']}"
        )
    extra = sorted(set(counters) - set(want["counters"]))
    assert not extra, f"{where}: counters docx4j does not have: {extra}"
    # docx4j records a ``w:num``'s first use here even where it carries no
    # ``w:startOverride``; so does ``ListLevel.increment_counter``.
    assert sorted(state.start_overrides_applied) == sorted(want["startOverridesApplied"]), (
        f"{where}: startOverridesApplied {sorted(state.start_overrides_applied)} "
        f"!= {want['startOverridesApplied']}"
    )


# ---------------------------------------------------------------------------
# the test
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("golden", _goldens(), ids=lambda path: path.stem)
def test_the_emulator_answers_what_docx4j_answers(golden: Path) -> None:
    """Every numbered paragraph of every story, against docx4j's own answer."""
    recorded = json.loads(golden.read_text(encoding="utf-8"))
    package = load(_fixture_for(golden), options=LoadOptions(mce_preprocess=True))
    emulator = emulator_of(package)
    assert emulator is not None
    states = NumberingStates()
    compared = 0

    for key, story in recorded["stories"].items():
        root, part = _story(package, key)
        state = states.for_part(part)
        found: list[Any] = []
        _walk(_children_of(root), state, found)
        expected = story["paragraphs"]
        assert len(found) == len(expected), (
            f"{golden.stem}/{key}: {len(found)} paragraphs, docx4j walked {len(expected)}"
        )

        for (paragraph, story_state), want in zip(found, expected, strict=True):
            where = f"{golden.stem}/{key}[{want['index']}] {want['text']!r}"
            assert _text_of(paragraph)[:80] == (want["text"] or ""), (
                f"{where}: the walks are out of step"
            )
            result = emulator.number(getattr(paragraph, "p_pr", None), story_state)
            numbering = want["numbering"]
            compared += 1

            if numbering is None:
                assert result is None, (
                    f"{where}: this numbers it {result.num_string!r}, docx4j does not"
                )
                continue

            if numbering["numString"] is None:
                # CR-003 section 18.2 item 4: where the ``w:numId`` names no
                # list docx4j returns an empty ``ResultTriple`` and logs; here
                # ``number`` answers None, which is what "not a list item"
                # means to every caller in this package. No label either way,
                # and the counters below must be untouched.
                assert result is None, (
                    f"{where}: this numbers it {result.num_string!r}, "
                    "docx4j returns an empty result"
                )
            else:
                assert result is not None, (
                    f"{where}: this does not number it, docx4j says "
                    f"{numbering['numString']!r}"
                )
                assert result.num_string == numbering["numString"], (
                    f"{where}: numString {result.num_string!r} != "
                    f"{numbering['numString']!r}"
                )
                assert result.is_bullet == numbering["isBullet"], (
                    f"{where}: isBullet {result.is_bullet} != {numbering['isBullet']}"
                )
                assert str(result.ilvl) == numbering["ilvl"], (
                    f"{where}: ilvl {result.ilvl} != {numbering['ilvl']}"
                )
                assert result.num_id == numbering["numId"], (
                    f"{where}: numId {result.num_id!r} != {numbering['numId']!r}"
                )

            _check_num_ref(_num_ref(emulator, paragraph), numbering["numRef"], where)
            _check_state(story_state, numbering["stateAfter"], where)

    assert compared == sum(len(s["paragraphs"]) for s in recorded["stories"].values())


def test_every_numbering_golden_has_its_fixture() -> None:
    """Every golden whose subject is numbering has its fixture here, and ran."""
    numbering = {
        path.stem
        for path in GOLDENS.glob("*.json")
        if any(
            paragraph["numbering"]
            for story in json.loads(path.read_text(encoding="utf-8"))["stories"].values()
            for paragraph in story["paragraphs"]
        )
    }
    have = {path.stem for path in _goldens()}
    # CR-002 Phase B copied the rest of the parity set, so every golden that
    # numbers anything --- ``fonts-modesOfApplication`` and ``tracked-changes``
    # among them --- now has its fixture here and runs.
    assert numbering - have == set()
