"""A document described for comparison with Word's saves. docx4j-core-tests ``DocumentDescription``.

CR-010 section 3.6; docx4j CR-043 section 3.5. The port of docx4j's
``org.docx4j.model.revisions.DocumentDescription`` (541 lines, at ``00614748e``),
which its Word tests and the ports' agreement run (``PortsAgreement``) compare
by: each story's paragraphs --- style, alignment, indent, list, the mark's
revision --- with their text in runs of one formatting and one revision
(adjacent runs merged, since Word resplits runs on save); tables with their
table, row and cell properties; notes, comments, text boxes (every branch's),
equations, bookmarks; and each section's properties and effective headers and
footers.

What Word's save normalises does not show:

* ids, dates and rsids;
* field codes --- a field is its result, so an accepted ``HYPERLINK`` field and
  Word's ``w:hyperlink`` read alike (C16);
* whether a section's header reference is explicit or inherited;
* the order of comments, and list numbers (named by first appearance);
* the order of comment range ends and references standing together: ends
  first, Word 365 saving a reply's reference after its own range's end
  (docx4j ``d18693760``);
* the empty headers, footers and notes parts Word's save adds;
* the defaults Word writes out: a table's ``w:tblLook``, a cell's ``w:tcW``
  auto 0, a section's ``w:cols w:space="720"``, and a row's ``w:tblPrEx``
  repeating the table or holding zero top and bottom cell margins (C23, C25).

``describe(pkg)`` shows the table, row, cell and section properties (canonical
XML) and the records' presence, as docx4j's phase 2 made it (CR-010 Phase B);
``describe(pkg, records=False)`` leaves them out, as its phase 1 did.
``differing(a, b)`` is ``PortsAgreement.differing``: how many lines of one
description the other lacks, both ways.
"""

from __future__ import annotations

import dataclasses
from typing import Any

from lxml import etree

from docx4j_py.child import is_any_element, iter_children
from docx4j_py.model.revisions import tables, trees

__all__ = ["describe", "differing"]

_W = trees.W_NS

_TABLE_SKIP = frozenset({"tblPrChange", "tblLook"})
_ROW_SKIP = frozenset({"ins", "del", "trPrChange"})
_EX_SKIP = frozenset({"tblPrExChange"})
_CELL_SKIP = frozenset({"tcPrChange"})
_SECTION_SKIP = frozenset({"headerReference", "footerReference", "sectPrChange"})


def describe(package: Any, records: bool = True) -> str:
    """The description of a package."""
    description = _Description(records)
    description.describe(package)
    return "".join(description.out)


def differing(a: str, b: str) -> int:
    """How many lines of one description the other lacks, both ways (``PortsAgreement.differing``)."""
    x, y = a.split("\n"), b.split("\n")
    remaining = list(y)
    left = []
    for line in x:
        if line in remaining:
            remaining.remove(line)
        else:
            left.append(line)
    return len(left) + len(remaining)


@dataclasses.dataclass(slots=True)
class _Seg:
    """A piece of text (or a marker) with its formatting and the revisions around it."""

    text: str
    format: str
    context: str
    marker: bool


def _render(segs: list[_Seg]) -> str:
    segs = _ends_before_references(segs)
    merged: list[_Seg] = []
    for seg in segs:
        last = merged[-1] if merged else None
        if (
            last is not None
            and not seg.marker
            and not last.marker
            and last.format == seg.format
            and last.context == seg.context
        ):
            last.text += seg.text
        else:
            merged.append(_Seg(seg.text, seg.format, seg.context, seg.marker))
    out = []
    for seg in merged:
        if seg.marker:
            out.append(seg.text)
        else:
            text = seg.text.replace("\t", "\\t").replace("\n", "\\n")
            out.append(f"\u00ab{text}\u00bb{seg.format}" + (f"[{seg.context}]" if seg.context else ""))
    return "".join(out)


def _ends_before_references(segs: list[_Seg]) -> list[_Seg]:
    """Comment range ends and references standing together, ends first.

    docx4j ``endsBeforeReferences`` (``d18693760``): Word 365 saves each reply's
    reference after its own range's end (end, reference, end, reference) where
    the source and Word 2010 have the ends together (CR-043 section 15,
    ``review-for-word-word365-rejected``), which says nothing about the comments.
    """
    out = list(segs)
    i = 0
    while i < len(out):
        j = i
        while j < len(out) and out[j].marker and out[j].text in ("c\u2046", "\u2045cref\u2046"):
            j += 1
        if j - i > 1:
            ends = [x for x in out[i:j] if x.text == "c\u2046"]
            refs = [x for x in out[i:j] if x.text != "c\u2046"]
            out[i:j] = ends + refs
        i = max(j, i + 1)
    return out


def _value(found: Any) -> Any:
    """An attribute value as it reads: an enum's value, else the value itself."""
    import enum

    return found.value if isinstance(found, enum.Enum) else found


class _Description:
    def __init__(self, records: bool) -> None:
        self.out: list[str] = []
        self.records = records
        self.num_ids: dict[Any, int] = {}

    def line(self, depth: int, text: str) -> None:
        self.out.append("  " * depth + text + "\n")

    # -- the document ----------------------------------------------------------

    def describe(self, package: Any) -> None:
        from docx4j_py.openpackaging.parts.binary_part import BinaryPart
        from docx4j_py.openpackaging.parts.wml import FooterPart, HeaderPart

        main = package.main_document_part
        body = main.contents.body
        self.line(0, "== main")
        self.blocks(body.content, 0)
        for title, attr, field in (("footnotes", "footnotes_part", "footnote"), ("endnotes", "endnotes_part", "endnote")):
            part = getattr(main, attr, None)
            notes = getattr(part.contents, field, None) if part is not None else None
            if notes and any(_normal(note) for note in notes):
                self.line(0, f"== {title}")
                for note in notes:
                    if not _normal(note):
                        continue
                    self.line(1, "-- note")
                    self.blocks(note.content, 1)
        if main.comments_part is not None:
            self.line(0, "== comments")
            # Word writes the comments in the order of their anchors: compared as a sorted set
            described = []
            for comment in main.comments_part.contents.comment or ():
                inner = _Description(self.records)
                inner.num_ids = self.num_ids
                inner.line(1, f"-- comment by {comment.author}")
                inner.blocks(comment.content, 1)
                described.append("".join(inner.out))
            for text in sorted(described):
                self.out.append(text)
        self.headers_and_footers(package, main)
        parts = list(package.parts.values())
        images = sum(
            1
            for part in parts
            if isinstance(part, BinaryPart) and (part.content_type or "").startswith("image/")
        )
        self.line(0, f"== image parts: {images}")
        # a dropped section's headers and footers go with it (C06, D7); the empty ones Word's save adds do not count
        headers = sum(1 for part in parts if isinstance(part, HeaderPart) and not _empty(part.contents.content))
        footers = sum(1 for part in parts if isinstance(part, FooterPart) and not _empty(part.contents.content))
        self.line(0, f"== header parts: {headers}, footer parts: {footers}")

    def headers_and_footers(self, package: Any, main: Any) -> None:
        """Each section's effective header and footer of each type, inherited references resolved as Word does."""
        body = main.contents.body
        sections = []
        p_class = trees.cls("P")
        for item in body.content:
            if isinstance(item, p_class) and item.p_pr is not None and item.p_pr.sect_pr is not None:
                sections.append(item.p_pr.sect_pr)
        if body.sect_pr is not None:
            sections.append(body.sect_pr)
        rels = main.relationships_part
        effective: dict[str, Any] = {}
        for n, sect_pr in enumerate(sections, start=1):
            for ref in sect_pr.header_reference_or_footer_reference or ():
                kind = ("header " if trees.name(ref) == "headerReference" else "footer ") + str(
                    _value(getattr(ref, "type_value", None))
                )
                effective[kind] = rels.get_part(ref.id) if rels is not None and ref.id else None
            if self.records:
                changed = " sectPrChange" if sect_pr.sect_pr_change is not None else ""
                described = _props(sect_pr, "sectPr", _SECTION_SKIP).replace('<w:cols w:space="720">', "")
                self.line(0, f"== section {n}{changed} [{described}]")
            for kind, part in effective.items():
                content = part.contents.content if part is not None else []
                if _empty(content):
                    continue
                self.line(0, f"== section {n} {kind.lower()}")
                self.blocks(content, 1)

    # -- blocks ----------------------------------------------------------------

    def blocks(self, items: list, depth: int) -> None:
        p_class, tbl_class = trees.cls("P"), trees.cls("Tbl")
        from docx4j_py.wml import SdtBlock

        for item in items:
            name = trees.name(item)
            if isinstance(item, p_class):
                self.paragraph(item, depth)
            elif isinstance(item, tbl_class):
                self.table(item, depth)
            elif isinstance(item, SdtBlock):
                self.line(depth, "SDT")
                self.blocks(item.sdt_content.content if item.sdt_content is not None else [], depth + 1)
            elif (
                trees.revision_of(item) is not None
                or name.endswith(("RangeStart", "RangeEnd"))
                or name in ("bookmarkStart", "bookmarkEnd")
            ):
                segs: list[_Seg] = []
                self.run_level([item], "", segs)
                self.line(depth, "BLOCK-LEVEL " + _render(segs))
            elif trees.block_items(item) is not None:
                self.line(depth, type(item).__name__)
                self.blocks(trees.block_items(item), depth + 1)

    def table(self, tbl: Any, depth: int) -> None:
        head = "TABLE"
        if self.records:
            if tables.has_records(tbl):
                head += " records"
            grid = " ".join(str(col.w) for col in (tbl.tbl_grid.grid_col if tbl.tbl_grid is not None else ()))
            head += f" tblPr[{_props(tbl.tbl_pr, 'tblPr', _TABLE_SKIP)}] grid[{grid}]"
        self.line(depth, head)
        for tr, _owner in tables.rows_of(tbl):
            rev = ""
            tr_pr = tr.tr_pr
            if tr_pr is not None and tr_pr.ins is not None:
                rev += f" ins({tr_pr.ins.author})"
            if tr_pr is not None and tr_pr.del_value is not None:
                rev += f" del({tr_pr.del_value.author})"
            if self.records:
                if _in_change(tr):
                    rev += " records"
                rev += f" trPr[{_props(tr_pr, 'trPr', _ROW_SKIP)}]"
                ex = _exceptions(tr, tbl)
                if ex:
                    rev += f" tblPrEx[{ex}]"
            self.line(depth + 1, "ROW" + rev)
            for tc, _cell_owner in tables.cells_of(tr):
                described = ""
                if self.records:
                    # Word's save writes w:tcW auto 0 where a cell has no width
                    described = " tcPr[" + _props(tc.tc_pr, "tcPr", _CELL_SKIP).replace(
                        '<w:tcW w:type="auto" w:w="0">', ""
                    ) + "]"
                self.line(depth + 2, "CELL" + described)
                self.blocks(tc.content, depth + 3)

    def paragraph(self, p: Any, depth: int) -> None:
        out = ["P"]
        p_pr = p.p_pr
        if p_pr is not None:
            if p_pr.p_style is not None:
                out.append(f" style={p_pr.p_style.val}")
            if p_pr.jc is not None:
                out.append(f" jc={_value(p_pr.jc.val)}")
            ind = p_pr.ind
            if ind is not None:
                out.append(f" ind={ind.left}/{ind.right}/{ind.hanging}/{ind.first_line}")
            num_pr = p_pr.num_pr
            if num_pr is not None:
                num_id = self.num_id(num_pr.num_id.val if num_pr.num_id is not None else None)
                ilvl = num_pr.ilvl.val if num_pr.ilvl is not None else None
                out.append(f" num={num_id}/{ilvl}")
            if p_pr.p_pr_change is not None:
                out.append(f" pPrChange({p_pr.p_pr_change.author})")
            if p_pr.sect_pr is not None:
                out.append(" sect")
            r_pr = p_pr.r_pr
            if r_pr is not None:
                for what, marker in (
                    ("ins", r_pr.ins),
                    ("del", r_pr.del_value),
                    ("moveFrom", r_pr.move_from),
                    ("moveTo", r_pr.move_to),
                ):
                    if marker is not None:
                        out.append(f" mark-{what}({marker.author})")
                if r_pr.r_pr_change is not None:
                    out.append(f" mark-rPrChange({r_pr.r_pr_change.author})")
        segs: list[_Seg] = []
        self.run_level(p.content, "", segs)
        self.line(depth, "".join(out) + " | " + _render(segs))

    def num_id(self, value: Any) -> str:
        """Word renumbers lists on save: a list is named by the order it first appears in."""
        if value is None or str(value) == "0":
            return str(value)
        found = self.num_ids.get(value)
        if found is None:
            found = self.num_ids[value] = len(self.num_ids) + 1
        return f"#{found}"

    # -- run level -------------------------------------------------------------

    def run_level(self, items: list, context: str, out: list[_Seg]) -> None:
        from docx4j_py.wml import SdtRun

        r_class = trees.cls("R")
        for item in items:
            name = trees.name(item)
            revision = trees.revision_of(item)
            if revision is not None:
                label = {"ins": "ins", "del": "del", "moveFrom": "move_from", "moveTo": "move_to"}[revision.value]
                inner = (context + ">" if context else "") + f"{label}({item.author})"
                self.run_level(trees.run_items(item) or [], inner, out)
                continue
            if isinstance(item, r_class):
                self.run(item, context, out)
                continue
            if name == "commentRangeStart":
                out.append(_Seg("\u2045c", "", "", True))
                continue
            if name == "commentRangeEnd":
                out.append(_Seg("c\u2046", "", "", True))
                continue
            if name == "bookmarkStart":
                mark = getattr(item, "name", None)
                if mark is not None and not mark.startswith("_"):
                    out.append(_Seg(f"\u2045b:{mark}", "", "", True))
                continue
            if name == "bookmarkEnd":
                continue  # ids only: a start without its end shows as the start alone
            if name.startswith(("moveFromRange", "moveToRange", "customXml")):
                out.append(_Seg(f"\u2045{name}\u2046", "", "", True))
                continue
            if isinstance(item, SdtRun):
                out.append(_Seg("\u2045sdt", "", "", True))
                self.run_level(item.sdt_content.content if item.sdt_content is not None else [], context, out)
                out.append(_Seg("sdt\u2046", "", "", True))
                continue
            if trees.is_math(item):
                out.append(_Seg(f"\u2045math:{_math(item)}\u2046", "", "", True))
                continue
            nested = trees.run_items(item)
            if nested is not None:
                self.run_level(nested, context, out)

    def run(self, run: Any, context: str, out: list[_Seg]) -> None:
        described = _format(run.r_pr)
        for item in run.content:
            name = trees.name(item)
            text = trees.item_text(item)
            if text is not None:
                out.append(_Seg(text, described, context, False))
            elif name in ("footnoteReference", "endnoteReference"):
                out.append(_Seg("\u2045" + ("fn" if name == "footnoteReference" else "en") + "\u2046", "", context, True))
            elif name == "commentReference":
                out.append(_Seg("\u2045cref\u2046", "", context, True))
            elif name in ("drawing", "pict", "AlternateContent", "object"):
                out.append(_Seg(f"\u2045{self.drawing(item)}\u2046", "", context, True))

    def drawing(self, node: Any) -> str:
        """A drawing: a picture, or the text boxes in it (every branch's)."""
        boxes: list[Any] = []
        picture = [False]

        def enter(found: Any) -> bool:
            name = trees.name(found)
            if name == "txbxContent":
                boxes.append(found)
                return False
            if name in ("blip", "imagedata"):
                picture[0] = True
            return True

        trees.walk(node, enter)
        if not boxes:
            return "img" if picture[0] else "drawing"
        out = []
        for box in boxes:
            typed = _typed_box(box)
            inner = _Description(self.records)
            inner.blocks(typed.content, 0)
            out.append("txbx{" + "".join(inner.out).strip().replace("\n", " / ") + "}")
        return "".join(out)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _normal(note: Any) -> bool:
    kind = getattr(getattr(note, "type_value", None), "value", None)
    return kind is None or kind == "normal"


def _empty(items: list) -> bool:
    """True when a header or footer says nothing: Word's save writes empty ones for every section."""
    p_class = trees.cls("P")
    for item in items or ():
        if not isinstance(item, p_class):
            return False
        if item.content:
            segs: list[_Seg] = []
            _Description(False).run_level(item.content, "", segs)
            if segs:
                return False
    return True


def _typed_box(box: Any) -> Any:
    """A text box's content, typed: itself, or its wildcard tree parsed (as an island is)."""
    if not is_any_element(box):
        return box
    from docx4j_py.model.revisions.islands import Island

    return Island(box, []).typed


def _format(r_pr: Any) -> str:
    if r_pr is None:
        return ""
    out = []
    if r_pr.r_style is not None:
        out.append(f"rStyle={r_pr.r_style.val}")
    for name, field in (
        ("b", "b"),
        ("i", "i"),
        ("strike", "strike"),
        ("dstrike", "dstrike"),
        ("caps", "caps"),
        ("smallCaps", "small_caps"),
        ("vanish", "vanish"),
    ):
        found = getattr(r_pr, field, None)
        if found is not None and getattr(found, "val", None) in (None, True, 1, "1", "true", "on"):
            out.append(name)
    if r_pr.u is not None and r_pr.u.val is not None:
        out.append(f"u={_value(r_pr.u.val)}")
    if r_pr.color is not None:
        out.append(f"color={r_pr.color.val}")
    if r_pr.sz is not None:
        out.append(f"sz={r_pr.sz.val}")
    if r_pr.highlight is not None:
        out.append(f"highlight={_value(r_pr.highlight.val)}")
    if r_pr.vert_align is not None:
        out.append(f"vertAlign={_value(r_pr.vert_align.val)}")
    if r_pr.r_pr_change is not None:
        out.append(f"rPrChange({r_pr.r_pr_change.author})")
    return "{" + " ".join(out) + "}" if out else ""


def _math(root: Any) -> str:
    """An equation: its text, structures by name, and the revisions in it."""
    ctr, ctrl_pr = trees.cls("CTR"), trees.cls("CTCtrlPr")
    track = trees.cls("CTMathRunTrackChange")
    out: list[str] = []

    def visit(node: Any) -> None:
        for qname, child in iter_children(node, mce="all"):
            local = (qname or "").rpartition("}")[2]
            if isinstance(child, ctr):
                for item in child.content:
                    if isinstance(item, track):
                        out.append(f"[{trees.name(item)}({item.author}):")
                        for inner in item.content:
                            text = trees.item_text(inner)
                            if text is not None:
                                out.append(text)
                        out.append("]")
                    else:
                        text = trees.item_text(item)
                        if text is not None:
                            out.append(text)
                continue
            if isinstance(child, ctrl_pr):
                marker = child.r_pr_or_ins_or_del
                name = trees.name(marker) if marker is not None else ""
                if name == "ins":
                    out.append("[struct-ins]")
                if name == "del":
                    out.append("[struct-del]")
                continue
            if trees.is_math(child) and local and local not in ("r", "oMath") and not local.endswith("Pr"):
                out.append(local + ":")
            visit(child)

    visit(root)
    return "".join(out)


def _in_change(tr: Any) -> bool:
    """A row in the change: one with a ``w:trPrChange`` or a ``w:tblPrExChange``, or a cell with a ``w:tcPrChange``."""
    if tr.tr_pr is not None and tr.tr_pr.tr_pr_change is not None:
        return True
    if tr.tbl_pr_ex is not None and tr.tbl_pr_ex.tbl_pr_ex_change is not None:
        return True
    return any(tc.tc_pr is not None and tc.tc_pr.tc_pr_change is not None for tc, _o in tables.cells_of(tr))


def _element(value: Any, local: str) -> Any:
    """A properties object as an lxml element named ``w:<local>``."""
    if value is None:
        return None
    from docx4j_py.fragments import to_xml

    return etree.fromstring(to_xml(value, name=f"{{{_W}}}{local}").encode("utf-8"))


def _inner(element: Any, skip: frozenset[str] | None, out: list[str]) -> None:
    """An element's attributes and children, canonical (docx4j ``Props.inner``); `skip` applies to its own children."""
    attributes = {}
    for key, value in element.attrib.items():
        uri, _, local = key[1:].partition("}") if key.startswith("{") else ("", "", key)
        if local.startswith("rsid"):
            continue
        attributes["{" + (uri or "null") + "}" + local] = value
    for key in sorted(attributes):
        out.append(f' {key}="{attributes[key]}"')
    if element.text and element.text.strip():
        out.append(element.text)
    for child in element:
        if not isinstance(child.tag, str):
            continue
        uri, _, local = child.tag[1:].partition("}")
        if skip is not None and local in skip:
            continue
        out.append(f"<{{{uri}}}{local}")
        _inner(child, None, out)
        out.append(">")
        if child.tail and child.tail.strip():
            out.append(child.tail)


def _props(value: Any, local: str, skip: frozenset[str]) -> str:
    """Properties as canonical XML, the WordprocessingML namespace shortened."""
    element = _element(value, local)
    if element is None:
        return ""
    out: list[str] = []
    _inner(element, skip, out)
    return "".join(out).replace("{" + _W + "}", "w:")


def _exceptions(tr: Any, tbl: Any) -> str:
    """A row's ``w:tblPrEx``, without what Word's save drops (C23, C25)."""
    if tr.tbl_pr_ex is None:
        return ""
    ex = _element(tr.tbl_pr_ex, "tblPrEx")
    tbl_pr = _element(tbl.tbl_pr, "tblPr")
    out = []
    for child in ex:
        if not isinstance(child.tag, str):
            continue
        local = child.tag.rpartition("}")[2]
        if local in _EX_SKIP:
            continue
        mine: list[str] = []
        _inner(child, None, mine)
        same = False
        for theirs in tbl_pr if tbl_pr is not None else ():
            if not isinstance(theirs.tag, str) or theirs.tag.rpartition("}")[2] != local:
                continue
            other: list[str] = []
            _inner(theirs, None, other)
            same = "".join(mine) == "".join(other)
        text = "".join(mine).replace("{" + _W + "}", "w:")
        if same or (local == "tblCellMar" and text == '<w:top w:type="dxa" w:w="0"><w:bottom w:type="dxa" w:w="0">'):
            continue
        out.append(f"<w:{local}{text}>")
    return "".join(out)
