"""Canonicalisation + structural diff helpers for OOXML round-trip comparison.

Two levels are provided:

* ``c14n(xml_bytes)`` - lxml Exclusive-C14N.  This already normalises
  attribute order, namespace *declaration* order and redundant declarations,
  and it discards the XML declaration.  It does NOT normalise prefixes
  (C14N is prefix-preserving) and it does NOT touch whitespace-only text.

* ``normalize(xml_bytes)`` - a prefix-insensitive normal form: every element
  and attribute is rendered as ``{namespace}localname``, attributes are
  sorted, and whitespace-only text nodes are dropped *unless* the element
  (or an ancestor) carries ``xml:space="preserve"``.  This is the form used
  for the "identical?" verdict, because two documents that differ only in
  which prefix is bound to a namespace are semantically the same OOXML.

``diff_trees`` walks two normalised trees in parallel and classifies the
first N differences into categories.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Any

from lxml import etree

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def parse(xml_bytes: bytes) -> etree._ElementTree:
    p = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
    return etree.parse(io.BytesIO(xml_bytes), p)


def c14n(xml_bytes: bytes) -> bytes:
    tree = parse(xml_bytes)
    buf = io.BytesIO()
    tree.write_c14n(buf, exclusive=True, with_comments=False)
    return buf.getvalue()


def _norm_elem(el, out, preserve: bool) -> None:
    if isinstance(el.tag, str):
        preserve = {"preserve": True, "default": False}.get(
            el.get(XML_SPACE), preserve
        )
    out.append(el)


def normalize_tree(xml_bytes: bytes) -> etree._Element:
    """Return a prefix-free, whitespace-normalised copy of the tree."""
    root = parse(xml_bytes).getroot()
    return _copy(root, preserve=False)


def _copy(el, preserve: bool):
    if not isinstance(el.tag, str):  # comment / PI
        return None
    sp = el.get(XML_SPACE)
    if sp == "preserve":
        preserve = True
    elif sp == "default":
        preserve = False

    new = etree.Element(el.tag)
    for k, v in sorted(el.attrib.items()):
        new.set(k, v)

    text = el.text
    if text is not None and not preserve and not text.strip():
        text = None
    new.text = text

    for child in el:
        c = _copy(child, preserve)
        if c is not None:
            tail = child.tail
            if tail is not None and not preserve and not tail.strip():
                tail = None
            c.tail = tail
            new.append(c)
    return new


def normalize(xml_bytes: bytes) -> bytes:
    return etree.tostring(normalize_tree(xml_bytes))


# --------------------------------------------------------------------------
# structural diff
# --------------------------------------------------------------------------


@dataclass
class Diff:
    category: str
    path: str
    detail: str


# ``xsd:boolean`` has four lexical forms and OOXML uses all of them: Word
# writes ``1``/``0``, other producers write ``true``/``false``, and ST_OnOff
# accepts both for the same value. A serialiser has to pick one
# (``SerializerConfig(bool_format=...)`` in the fork), so a corpus written by
# more than one producer can never match on the spelling alone. Re-spelling a
# boolean is not a round-trip loss and CR-001 section 9 says so; it is counted
# on its own so that it is visible rather than hidden.
BOOLEAN_SPELLINGS = ({"1", "true"}, {"0", "false"})


def same_boolean(a: str, b: str) -> bool:
    """Return whether two attribute values are the same xsd:boolean."""
    return any(a in forms and b in forms for forms in BOOLEAN_SPELLINGS)


@dataclass
class DiffReport:
    diffs: list[Diff] = field(default_factory=list)
    counts: dict[str, int] = field(default_factory=dict)

    #: categories that are a different spelling of the same value, not a
    #: difference; they are counted and reported but ``identical`` ignores them
    BENIGN = ("boolean-spelling",)

    def add(self, category: str, path: str, detail: str, limit: int = 400) -> None:
        self.counts[category] = self.counts.get(category, 0) + 1
        if len(self.diffs) < limit:
            self.diffs.append(Diff(category, path, detail))

    @property
    def identical(self) -> bool:
        return not self.differences

    @property
    def differences(self) -> dict[str, int]:
        """Return the counts that are real differences."""
        return {k: v for k, v in self.counts.items() if k not in self.BENIGN}


URI_PREFIX = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
    "http://schemas.openxmlformats.org/officeDocument/2006/math": "m",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
    "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing": "wp",
    "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "c",
    "http://schemas.openxmlformats.org/markup-compatibility/2006": "mc",
    "http://schemas.microsoft.com/office/word/2010/wordml": "w14",
    "http://schemas.microsoft.com/office/word/2012/wordml": "w15",
    "http://schemas.microsoft.com/office/word/2006/wordml": "wne",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing": "wp14",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape": "wps",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup": "wpg",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas": "wpc",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk": "wpi",
    "http://schemas.microsoft.com/office/drawing/2010/main": "a14",
    "http://schemas.microsoft.com/office/drawing/2014/main": "a16",
    "http://schemas.microsoft.com/office/drawing/2016/SVG/main": "asvg",
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:office": "o",
    "urn:schemas-microsoft-com:office:word": "w10",
    "http://www.w3.org/XML/1998/namespace": "xml",
}


def _q(tag: str) -> str:
    """Render {uri}local as prefix:local when the uri is known."""
    if tag.startswith("{"):
        uri, local = tag[1:].split("}", 1)
        return f"{URI_PREFIX.get(uri, uri)}:{local}"
    return tag


def _path(el, parent_path: str, idx: int) -> str:
    return f"{parent_path}/{el.tag}[{idx}]"


def _ns(tag: str) -> str:
    return tag[1:].split("}")[0] if tag.startswith("{") else ""


def diff_trees(a: etree._Element, b: etree._Element, rep: DiffReport,
               path: str = "") -> None:
    if a.tag != b.tag:
        rep.add("element-renamed", path, f"{_q(a.tag)} -> {_q(b.tag)}")
        return

    here = path + "/" + _q(a.tag)

    # attributes
    ka, kb = set(a.attrib), set(b.attrib)
    for k in sorted(ka - kb):
        rep.add("attribute-dropped", here, f"{_q(k)}={a.attrib[k]!r}")
    for k in sorted(kb - ka):
        rep.add("attribute-added", here, f"{_q(k)}={b.attrib[k]!r}")
    for k in sorted(ka & kb):
        va, vb = a.attrib[k], b.attrib[k]
        if va != vb:
            category = ("boolean-spelling" if same_boolean(va, vb)
                        else "attribute-value")
            rep.add(category, here, f"{_q(k)}: {va!r} -> {vb!r}")

    # text
    ta = a.text or ""
    tb = b.text or ""
    if ta != tb:
        rep.add("text", here, f"{ta[:60]!r} -> {tb[:60]!r}")

    # children (order-sensitive)
    ca, cb = list(a), list(b)
    if len(ca) != len(cb):
        sa = [c.tag for c in ca]
        sb = [c.tag for c in cb]
        missing = [t for t in sa if sa.count(t) > sb.count(t)]
        extra = [t for t in sb if sb.count(t) > sa.count(t)]
        if missing:
            for t in sorted(set(missing)):
                rep.add("element-dropped", here, f"{_q(t)} x{missing.count(t)}")
        if extra:
            for t in sorted(set(extra)):
                rep.add("element-added", here, f"{_q(t)} x{extra.count(t)}")
        # still compare the common prefix pairwise by tag matching
        return

    for i, (x, y) in enumerate(zip(ca, cb)):
        if x.tag != y.tag:
            rep.add("child-order", here, f"pos {i}: {_q(x.tag)} -> {_q(y.tag)}")
        diff_trees(x, y, rep, here)
        # tails
        tx = x.tail or ""
        ty = y.tail or ""
        if tx != ty:
            rep.add("tail-text", here, f"after {_q(x.tag)}: {tx[:40]!r} -> {ty[:40]!r}")


def sort_children(el: etree._Element) -> None:
    """Sort an element's direct children by tag, in place.

    For a part whose root is an ``xsd:all`` --- ``docProps/core.xml`` and
    ``docProps/app.xml`` are the two in OOXML --- the order of the children is
    *defined* to carry no meaning, so a serialiser is free to write them in
    schema order where Word wrote them in another. Sorting both sides before
    the diff is how the round trip asks the only question that matters there.
    """
    children = sorted(el, key=lambda c: (c.tag if isinstance(c.tag, str) else ""))
    for child in children:
        el.append(child)


def compare(orig: bytes, new: bytes, *, unordered_root: bool = False) -> DiffReport:
    """Diff two parts.

    Args:
        orig: the source bytes.
        new: what the round trip produced.
        unordered_root: sort the root's children on both sides first, for a
            part whose root content model is ``xsd:all`` (see
            :func:`sort_children`).
    """
    rep = DiffReport()
    a, b = normalize_tree(orig), normalize_tree(new)
    if unordered_root:
        sort_children(a)
        sort_children(b)
    diff_trees(a, b, rep)
    return rep


def prefix_map(xml_bytes: bytes) -> dict[str, str]:
    """All prefix->uri bindings declared anywhere in the document."""
    out: dict[str, str] = {}
    for _, el in etree.iterwalk(parse(xml_bytes), events=("start",)):
        for p, u in el.nsmap.items():
            out[p or ""] = u
    return out
