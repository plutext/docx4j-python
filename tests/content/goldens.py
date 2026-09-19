"""The parity goldens, and the walk that reads them. CR-002 §12.12, §15.

The goldens are produced by the **Java** parity harness that lives in
``../docx4j-core-ts/test/java/``: it loads each fixture with docx4j itself and
writes down what docx4j answers, one JSON per ``.docx`` under
``../docx4j-core-ts/test/golden/``. CR-002 section 12.12 settled that the
harness lives there and that this repository **consumes** its committed
goldens, so the JSON is not copied here; the sibling checkout is the source and
every test over them skips when it is absent.

This module is what the three golden tests share: where the goldens are, which
fixture each belongs to, the harness's ``Walk`` over a story, its ``textOf``,
and the comparison of a property value --- an XML string docx4j marshalled with
**its** prefixes --- as an **object tree**, never as text.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from conftest import FIXTURES, ROOT

from docx4j_py.child import iter_children
from docx4j_py.fragments import to_xml, wml
from docx4j_py.wml import DelText, P, Text

#: The committed goldens of the sibling checkout. Not copied here: CR-002 §12.12.
GOLDENS = ROOT.parent / "docx4j-core-ts" / "test" / "golden"

#: Where a golden's fixture is under another name here. ``lists.docx`` is
#: ``numbering_indentation.docx`` byte for byte (tests/README.md), so the map
#: is the whole of the duplication the fixtures needed.
FIXTURE_NAMES = {"numbering_indentation": "lists"}

#: The node whose children begin a story of their own. docx4j's ``isTextBox``
#: reaches this through ``wp:inline``/``wp:anchor`` → ``a:graphicData`` →
#: ``wps:wsp``; starting the new story at the ``wps:wsp`` itself gives every
#: paragraph below it the same fresh state. ``v:textbox`` is named for
#: completeness: VML has no typed classes in this object model, so a VML text
#: box's paragraphs are a wildcard tree and no fixture here has one.
TEXT_BOXES = frozenset({"Wsp", "Textbox", "CTTextbox"})

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: The root element names the golden's property values carry, in the qualified
#: form :func:`docx4j_py.fragments.to_xml` takes.
W_PPR = f"{_W}pPr"
W_RPR = f"{_W}rPr"
W_STYLE = f"{_W}style"
W_IND = f"{_W}ind"
W_LVL = f"{_W}lvl"


def fixture_for(golden: Path) -> Path:
    """The ``.docx`` a golden was made from, under the name this repository gives it."""
    return FIXTURES / f"{FIXTURE_NAMES.get(golden.stem, golden.stem)}.docx"


def all_goldens() -> list[Path]:
    """Every golden whose fixture is in ``tests/fixtures/``."""
    if not GOLDENS.is_dir():
        return []
    return [path for path in sorted(GOLDENS.glob("*.json")) if fixture_for(path).exists()]


def read(golden: Path) -> dict[str, Any]:
    """A golden, parsed."""
    return json.loads(golden.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# the walk: the harness's Walk
# ---------------------------------------------------------------------------


def is_text_box(obj: Any) -> bool:
    """Whether this node begins a story of its own."""
    return type(obj).__name__ in TEXT_BOXES


def walk(children: Any, state: Any, out: list[Any], depth: int = 0) -> None:
    """Every ``w:p`` in document order, with the state its story counts in.

    Descends into tables, content controls, hyperlinks and revisions, and starts
    a fresh state at a text box, as the harness does. `state` may be anything
    the caller wants carried down --- a ``NumberingState``, or None where the
    walk is only after the paragraphs.
    """
    from docx4j_py.model.listnumbering import NumberingState

    if depth > 60:  # pragma: no cover - the harness's own guard
        return
    for _name, child in children:
        if isinstance(child, P):
            out.append((child, state))
            walk(iter_children(child), state, out, depth + 1)
        elif is_text_box(child):
            walk(iter_children(child), NumberingState(), out, depth + 1)
        else:
            walk(iter_children(child), state, out, depth + 1)


def story(package: Any, key: str) -> tuple[Any, Any]:
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


def children_of(root: Any) -> list[Any]:
    """A story root's children: a list of notes or comments is their contents."""
    if isinstance(root, list):
        out: list[Any] = []
        for item in root:
            out.extend(iter_children(item))
        return out
    return list(iter_children(root))


def text_of(paragraph: Any) -> str:
    """The harness's ``textOf``: every ``w:t`` and ``w:delText``, no text box."""
    out: list[str] = []

    def collect(obj: Any, depth: int) -> None:
        if depth > 60:  # pragma: no cover - the harness's own guard
            return
        for _name, child in iter_children(obj):
            if isinstance(child, (Text, DelText)):
                out.append(child.value or "")
            elif is_text_box(child) or isinstance(child, P):
                continue
            else:
                collect(child, depth + 1)

    collect(paragraph, 0)
    return "".join(out)


def runs_of(paragraph: Any) -> list[Any]:
    """Every ``w:r`` of a paragraph in document order, the harness's order.

    Runs inside a nested ``w:p`` --- a text box's --- belong to that paragraph,
    not to this one, so the descent stops there.
    """
    from docx4j_py.wml import R

    out: list[Any] = []

    def collect(obj: Any, depth: int) -> None:
        if depth > 60:  # pragma: no cover
            return
        for _name, child in iter_children(obj):
            if isinstance(child, R):
                out.append(child)
            elif is_text_box(child) or isinstance(child, P):
                continue
            else:
                collect(child, depth + 1)

    collect(paragraph, 0)
    return out


def run_text(run: Any) -> str:
    """A run's text: every ``w:t`` and ``w:delText`` under it."""
    out: list[str] = []
    for _name, child in iter_children(run):
        if isinstance(child, (Text, DelText)):
            out.append(child.value or "")
    return "".join(out)


# ---------------------------------------------------------------------------
# comparing a property value
# ---------------------------------------------------------------------------


def as_tree(xml: str | None, wrapper: type | str | None = None) -> Any:
    """A golden's XML string as an object tree, or None.

    docx4j marshalled it with **its** prefixes, so it is never compared as text.
    """
    return None if xml is None else wml(xml, wrapper=wrapper, lenient=True)


def normalised(obj: Any, name: str | None = None, wrapper: type | str | None = None) -> Any:
    """One of our objects through a marshal and an unmarshal.

    So that what is compared is what each library would *write*, not how it
    happens to hold it: a value only one side keeps in memory (an empty list, a
    default nobody marshals) cannot make a difference, and neither can a class
    that stands for the same element.
    """
    if obj is None:
        return None
    return wml(to_xml(obj, name=name), wrapper=wrapper, lenient=True)


def same(ours: Any, theirs: Any, where: str, name: str | None = None,
         wrapper: type | str | None = None) -> None:
    """Assert that our object and a golden's XML string are the same tree."""
    mine = normalised(ours, name=name, wrapper=wrapper)
    yours = as_tree(theirs, wrapper=wrapper)
    assert mine == yours, (
        f"{where}:\n  ours   {None if ours is None else to_xml(ours, name=name)}"
        f"\n  docx4j {theirs}"
    )
