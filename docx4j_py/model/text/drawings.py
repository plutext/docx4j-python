"""DrawingML's text for the text model: a paragraph's (``a:p``), and SmartArt's in reading order (WD REQ-511).

docx4j's ``org.docx4j.model.text.Drawings`` at ``198d015a3``. CR-011
section 6.
"""

from __future__ import annotations

from typing import Any

__all__ = ["lines", "smart_art", "text"]


def text(p: Any) -> str:
    """A DrawingML paragraph's text: its runs' and fields' ``a:t``, an ``a:br`` a line break."""
    out: list[str] = []
    for item in getattr(p, "content", None) or ():
        name = type(item).__name__
        if name in ("CTRegularTextRun", "CTTextField"):
            value = getattr(item, "t", None)
            if value is not None:
                out.append(value)
        elif name == "CTTextLineBreak":
            out.append("\n")
    return "".join(out)


def lines(body: Any) -> list[str]:
    """A text body's paragraphs, a line each; none when every one is empty."""
    if body is None:
        return []
    out = [text(p) for p in getattr(body, "p", None) or ()]
    return out if any(out) else []


def _value(item: Any, default: str) -> str:
    value = getattr(item, "type_value", None)
    value = getattr(value, "value", value)
    return default if value is None else str(value)


def smart_art(model: Any) -> list[str]:
    """SmartArt's text from its data part (REQ-511), in reading order.

    From the ``doc`` point, depth-first along the ``parOf`` connections, each
    point's children by ``@srcOrd``; the text of the points of type ``node``
    (the default) and ``asst``, each ``a:p`` a line. A point whose paragraphs
    are all empty gives nothing. Storage order (of ``dgm:ptLst`` or
    ``dgm:cxnLst``) is not reading order, and the drawing part's cached shapes
    are not read.
    """
    out: list[str] = []
    pt_lst = getattr(model, "pt_lst", None) if model is not None else None
    if pt_lst is None:
        return out
    points: dict[str, Any] = {}
    doc = None
    for pt in pt_lst.pt or ():
        if pt.model_id is None:
            continue
        points[str(pt.model_id)] = pt
        if doc is None and _value(pt, "node") == "doc":
            doc = pt
    if doc is None:
        return out
    children: dict[str, list[Any]] = {}
    cxn_lst = getattr(model, "cxn_lst", None)
    for cxn in (cxn_lst.cxn if cxn_lst is not None else None) or ():
        if _value(cxn, "parOf") != "parOf" or cxn.src_id is None:
            continue
        children.setdefault(str(cxn.src_id), []).append(cxn)
    for found in children.values():
        found.sort(key=lambda c: int(c.src_ord or 0))
    seen: set[str] = set()

    def visit(pt: Any) -> None:
        key = str(pt.model_id)
        if key in seen:
            return  # a cycle
        seen.add(key)
        if _value(pt, "node") in ("node", "asst"):
            out.extend(lines(pt.t))
        for cxn in children.get(key, ()):
            child = points.get(str(cxn.dest_id))
            if child is not None:
                visit(child)

    visit(doc)
    return out
