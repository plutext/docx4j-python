"""docx4j-core-ts's Office JS checks 52 to 64, ported to the content API: CR-013's, CR-016's and CR-019's oracles.

CR-013 section 6.1. Word 365 ran each check's Script Lab snippet on the file in
this directory and wrote JSON (``52-ranges-word365-run.json``,
``52-ranges-word365-accepted.json``, ``53-reads-word365-simple.json`` and
``-all.json``). core-ts runs the same snippets through its Office JS shim; this
package has none (CR-013 section 4.9 item 6), so the snippets are ported here,
step for step, onto the content API, and give the same JSON shape with the same
``vis()``, so that Word's JSON compares field for field.

* :func:`run_check_52` is ``CHECK52.runCheck`` (``check52-script-lab.js`` at
  ``ef57729``, the run's own at ``8fdccf8`` with the ``reads`` mode added) in
  its ``reads``, ``report`` and (CR-013 Phase C) ``run`` modes: the run applies
  each case's write in its region, with change tracking as the case says, then
  reports the region (``8fdccf8``'s ``apply`` and ``searches``).
* :func:`run_check_53` is ``CHECK53.runCheck`` (``check53-script-lab.js`` at
  ``ef57729``).
* :func:`run_check_54` is ``CHECK54.runCheck`` (``check54-script-lab.js`` at
  ``87bbb38``; CR-013 Phase C) in its ``run`` and ``report`` modes.
* :func:`run_check_57` is ``CHECK57.runCheck`` (``check57-script-lab.js`` at
  ``906d5e3``; CR-013 section 13.7) in its ``run`` and ``report`` modes.
* :func:`run_check_58` is ``CHECK58.runCheck`` (``check58-script-lab.js`` at
  ``222cc20``; CR-016 Phase B) in its ``read`` and ``edit`` modes, with its
  ``QUERIES`` (:data:`QUERIES_58`).
* :func:`run_check_59` is ``CHECK59.runCheck`` (``check59-script-lab.js`` at
  ``222cc20``, unchanged at ``3ba1e68``; CR-016 Phase C), read only, with its
  ``REGION_QUERIES`` and ``PART_QUERIES`` (:data:`REGION_QUERIES_59`,
  :data:`PART_QUERIES_59`).
* :func:`run_check_63` and :func:`run_check_64` are ``CHECK63.runCheck`` and
  ``CHECK64.runCheck`` (``check63-script-lab.js`` and
  ``check64-script-lab.js`` at ``020c039``; CR-019 Phase A) in their ``run``
  and ``report`` modes: each case's one write into a content control showing
  its placeholder, in the bare form (63) and in Word's (64), then the region.

Each read is attempted alone, as the snippets sync each read alone, so that
what fails fails alone, as ``'error: ' + (code or message)``. A member this
package does not offer (``Range.footnotes``, ``Paragraph.get_text_ranges``,
``Body.footnotes``) is ``error: NotImplemented: ...``, as core-ts's shim
reports one. ``ContentControl.showing_placeholder_text`` is offered since CR-003
section 24 (core-ts ``222cc20``).

Office JS's names map to the content API's as the subset does (CR-003 section
3.1): ``getRange`` is ``get_range``, ``parentTableCell`` is
``parent_table_cell``, a collection is a list, ``getFirst()`` is ``[0]``,
``getByTag(tag).getFirst()`` the first control with that tag.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Any

__all__ = [
    "PARAGRAPH_OOXML",
    "TABLE_OOXML",
    "parse_region_52",
    "parse_region_53",
    "parse_region_54",
    "parse_region_57",
    "parse_region_58",
    "parse_region_59",
    "parse_region_63",
    "parse_region_64",
    "run_check_52",
    "run_check_53",
    "run_check_54",
    "run_check_57",
    "run_check_58",
    "run_check_59",
    "run_check_63",
    "run_check_64",
    "vis",
]


def vis(value: Any) -> Any:
    """Every character outside printable ASCII as ``\\uXXXX`` (a surrogate pair as two), as the snippets' ``vis``."""
    if value is None:
        return None
    out = []
    for char in str(value):
        code = ord(char)
        if 0x20 <= code <= 0x7E:
            out.append(char)
        elif code > 0xFFFF:
            code -= 0x10000
            out.append(f"\\u{0xD800 + (code >> 10):04x}\\u{0xDC00 + (code & 0x3FF):04x}")
        else:
            out.append(f"\\u{code:04x}")
    return "".join(out)


class NotImplementedMember(Exception):
    """A member of Office JS's this package does not offer: core-ts's shim's ``NotImplemented``."""

    code = "NotImplemented"


def failed(error: BaseException) -> str:
    """``'error: ' + (code or message)``, the snippets' ``failed``."""
    if isinstance(error, NotImplementedMember):
        return f"error: NotImplemented: {error}"
    code = getattr(error, "code", None)
    return "error: " + (str(code) if code else str(error) or type(error).__name__)


def attempt(read: Callable[[], Any]) -> Any:
    """One read alone: its answer, or :func:`failed`'s string."""
    try:
        return read()
    except Exception as error:  # noqa: BLE001 - every failure is the read's answer, as in the snippet
        return failed(error)


def member(owner: Any, name: str, office_js: str) -> Any:
    """``owner.name``, or :class:`NotImplementedMember` naming Office JS's member where this package has none."""
    if not hasattr(type(owner), name) and not hasattr(owner, name):
        raise NotImplementedMember(f"{type(owner).__name__}.{office_js}")
    return getattr(owner, name)


TOKEN = {"match_case": True, "match_whole_word": True}


def find(body: Any, token: str) -> Any:
    """``body.search(token, TOKEN).getFirstOrNullObject()``, refusing a token that is not there."""
    hits = body.search(token, **TOKEN)
    if not hits:
        raise LookupError(f"not found: {token}")
    return hits[0]


def control_by_tag(body: Any, tag: str) -> Any:
    """``body.contentControls.getByTag(tag).getFirst()``."""
    for control in body.content_controls:
        if control.tag == tag:
            return control
    raise LookupError(f"no content control tagged {tag}")


def font_of(range_: Any, names: tuple[str, ...]) -> dict[str, Any]:
    """A font's properties, Office JS's names to the content API's."""
    font = range_.font
    python = {
        "bold": "bold",
        "italic": "italic",
        "underline": "underline",
        "name": "name",
        "size": "size",
        "color": "color",
        "highlightColor": "highlight_color",
    }
    return {name: getattr(font, python[name]) for name in names}


def first_paragraph_text(range_: Any) -> Any:
    """``range.paragraphs.getFirstOrNullObject()``'s text, or None."""
    paragraphs = range_.paragraphs
    return vis(paragraphs[0].text) if paragraphs else None


# ---------------------------------------------------------------------------
# check 52 (core-ts check52-script-lab.js)
# ---------------------------------------------------------------------------

_BEGIN_52 = re.compile(r"^BEGIN(\d+) (\w+) ([\w-]+) (untracked|tracked)(?: (.*))?$")

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: What ``ooxml-replace`` puts in a cell: a flat OPC package whose body is a table of one cell, NEWOOXML, and nothing after it.
TABLE_OOXML = (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    "</pkg:xmlData></pkg:part>"
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    f'<w:document xmlns:w="{_W}"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>NEWOOXML</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'
    "</pkg:xmlData></pkg:part></pkg:package>"
)


def _query_text(query: str) -> str:
    """A query as written in the marker, its tokens upper-cased and Word's codes (``^p``) kept."""
    parts = re.split(r"(\^.)", query)
    return "".join(part if index % 2 else part.upper() for index, part in enumerate(parts))


def parse_region_52(text: str) -> dict[str, Any] | None:
    """A BEGIN paragraph's text as a case, or None (the snippet's ``parseRegion``)."""
    match = _BEGIN_52.match(text.strip())
    if match is None:
        return None
    spec = match.group(5) or ""
    region: dict[str, Any] = {
        "n": int(match.group(1)),
        "kind": match.group(2),
        "op": match.group(3),
        "tracked": match.group(4) == "tracked",
        "spec": spec,
    }
    patterns = [
        (r"^from (\S+) to (\S+)$", lambda m: {"from": m.group(1).upper(), "to": m.group(2).upper()}),
        (r"^from (\S+) to-mark (\S+)$", lambda m: {"from": m.group(1).upper(), "toMark": m.group(2).upper()}),
        (r"^whole (\S+)$", lambda m: {"whole": m.group(1).upper()}),
        (r"^control (\S+)$", lambda m: {"control": m.group(1)}),
        (r"^cell (\S+)$", lambda m: {"cell": m.group(1).upper()}),
        (r"^paragraph (\S+)$", lambda m: {"paragraph": m.group(1).upper()}),
    ]
    for pattern, make in patterns:
        found = re.match(pattern, spec)
        if found is not None:
            region["range"] = make(found)
            return region
    found = re.match(r"^queries (.*)$", spec)
    if found is not None:
        region["queries"] = [_query_text(q) for q in found.group(1).split(";")]
        return region
    if spec:
        raise ValueError(f'BEGIN{region["n"]}: no range in "{spec}"')
    return region


def _paragraph_of_52(body: Any, token: str) -> Any:
    """The whole of the paragraph holding a token, its mark included."""
    return find(body, token).paragraphs[0].get_range("Whole")


def _cell_of(body: Any, token: str) -> Any:
    cell = find(body, token).parent_table_cell
    if cell is None:
        raise LookupError(f"{token} is in no table cell")
    return cell


def range_of_52(body: Any, region: dict[str, Any]) -> Any:
    """The case's range, made afresh (the snippet's ``rangeOf``)."""
    r = region["range"]
    if "control" in r:
        return control_by_tag(body, r["control"]).get_range("Content")
    if "whole" in r:
        return _paragraph_of_52(body, r["whole"])
    if "cell" in r:
        return _cell_of(body, r["cell"]).body.get_range("Whole")
    if "paragraph" in r:
        return _paragraph_of_52(body, r["paragraph"])
    start = find(body, r["from"])
    return start.expand_to(_paragraph_of_52(body, r["toMark"]) if "toMark" in r else find(body, r["to"]))


def region_of_52(body: Any, n: int) -> Any:
    """The region from its BEGIN paragraph to its END paragraph, both whole."""
    return _paragraph_of_52(body, f"BEGIN{n}").expand_to(_paragraph_of_52(body, f"END{n}"))


def reads_52(body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """Everything a range reads, before anything is written (the snippet's ``reads``)."""
    out: dict[str, Any] = {}
    range_ = range_of_52(body, region)
    out["range"] = attempt(
        lambda: {
            "text": vis(range_.text),
            "isEmpty": range_.is_empty,
            "style": range_.style,
            "styleBuiltIn": range_.style_built_in,
        }
    )
    out["font"] = attempt(lambda: font_of(range_, ("bold", "italic", "underline", "name", "size")))
    for location in ("Whole", "Start", "End", "Content", "After"):

        def part_reads(location: str = location) -> dict[str, Any]:
            part = range_.get_range(location)
            return {"text": vis(part.text), "isEmpty": part.is_empty, "paragraph": first_paragraph_text(part)}

        out[location] = attempt(part_reads)
    out["paragraphs"] = attempt(lambda: [vis(p.text) for p in range_.paragraphs])
    out["tables"] = attempt(lambda: [table.row_count for table in range_.tables])
    out["marks"] = attempt(lambda: [vis(hit.text) for hit in range_.search("^p", match_case=True)])
    if "to" in region["range"]:
        r = region["range"]
        out["swapped"] = attempt(lambda: vis(find(body, r["to"]).expand_to(find(body, r["from"])).text))
    return out


def _list_of(paragraph: Any) -> dict[str, Any] | None:
    item = paragraph.list_item
    if item is None:
        return None
    return {"listString": vis(item.list_string), "level": item.level}


def report_52(body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """The region as it stands (the snippet's ``report``)."""
    out: dict[str, Any] = {}
    whole = region_of_52(body, region["n"])
    out["text"] = attempt(lambda: vis(whole.text))

    def paragraphs() -> list[dict[str, Any]]:
        views = whole.paragraphs
        words = []
        if region["op"] == "bold":
            words = [member(p, "get_text_ranges", "getTextRanges") for p in views]
        entries = []
        for index, p in enumerate(views):
            font = p.font
            entry: dict[str, Any] = {
                "text": vis(p.text),
                "style": p.style,
                "styleBuiltIn": p.style_built_in,
                "alignment": p.alignment,
                "tableNestingLevel": p.table_nesting_level,
                "bold": font.bold,
                "italic": font.italic,
                "underline": font.underline,
            }
            listed = _list_of(p)
            if listed is not None:
                entry["list"] = listed
            if words:
                entry["words"] = [{"text": vis(w.text), "bold": w.font.bold} for w in words[index]([" "], False)]
            entries.append(entry)
        return entries

    out["paragraphs"] = attempt(paragraphs)
    out["tables"] = attempt(
        lambda: [
            {
                "rowCount": table.row_count,
                "values": [[vis(value) for value in row] for row in table.values],
                "whole": vis(table.get_range("Whole").text),
            }
            for table in whole.tables
        ]
    )
    out["trackedChanges"] = attempt(
        lambda: [
            {"type": change.type, "text": vis(change.text), "range": vis(change.get_range().text)}
            for change in whole.get_tracked_changes()
        ]
    )
    out["comments"] = attempt(
        lambda: [
            {"content": vis(comment.content), "range": vis(comment.get_range().text)}
            for comment in whole.get_comments()
        ]
    )
    out["contentControls"] = attempt(
        lambda: [{"tag": control.tag, "text": vis(control.text)} for control in whole.content_controls]
    )

    def footnotes() -> list[Any]:
        notes = member(whole, "footnotes", "footnotes")
        return [vis(note.body.text) for note in notes]

    out["footnotes"] = attempt(footnotes)
    return out


def _returned(range_: Any) -> dict[str, Any]:
    """What a write returned, read back: ``{returned: text}`` (the snippets' ``returned``)."""
    return {"returned": vis(range_.text)}


def apply_52(package: Any, body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """The case's operation, on its range made afresh; what the operation returns (the snippet's ``apply``)."""
    op = region["op"]
    r = region.get("range", {})
    # CR-002 section 55's: a cell's body, its paragraph or its control
    if op == "table-end":
        table = _cell_of(body, r["cell"]).body.insert_table(1, 1, location="End", values=[["NEWCELL"]])
        return {"returned": [[vis(value) for value in row] for row in table.values]}
    if op == "paragraph-delete":
        find(body, r["paragraph"]).paragraphs[0].delete()
        return {}
    if op == "clear":
        _cell_of(body, r["cell"]).body.clear()
        return {}
    if op == "ooxml-replace":
        views = _cell_of(body, r["cell"]).body.insert_ooxml(TABLE_OOXML, location="Replace")
        return {"returned": vis(views[0].text) if views else ""}
    if op == "control-delete":
        control_by_tag(body, r["control"]).delete(keep_content=False)
        return {}
    range_ = range_of_52(body, region)
    if op == "delete":
        range_.delete()
        return {}
    if op == "replace":
        return _returned(range_.insert_text("NEWTEXT", location="Replace"))
    if op == "bold":
        range_.font.bold = True
        return {}
    if op == "paragraph-before":
        return _returned(range_.insert_paragraph("NEWPARA", location="Before"))
    if op == "paragraph-after":
        return _returned(range_.insert_paragraph("NEWPARA", location="After"))
    if op == "text-start":
        return _returned(range_.insert_text("NEWSTART", location="Start"))
    if op == "text-end":
        return _returned(range_.insert_text("NEWEND", location="End"))
    if op == "comment":
        return _returned(range_.insert_comment("NEWCOMMENT").get_range())
    raise ValueError(f"no operation {op}")


def searches_52(body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """Each query's hits in the region: their text and the paragraphs each touches (the snippet's ``searches``)."""
    out: dict[str, Any] = {}
    for query in region["queries"]:

        def hits(query: str = query) -> list[dict[str, Any]]:
            found = region_of_52(body, region["n"]).search(query, match_case=True)
            return [{"text": vis(hit.text), "paragraphs": [vis(p.text) for p in hit.paragraphs]} for hit in found]

        out[query] = attempt(hits)
    return out


def run_check_52(package: Any, mode: str) -> dict[str, Any]:
    """``CHECK52.runCheck`` in its ``reads``, ``report`` or ``run`` mode, on a loaded package.

    ``reads``: each kind's reads alone, nothing written and no region reported,
    what core-ts's phase 1 test compares with Word's run's ``reads``.
    ``report``: every region reported, nothing changed (for Word's save after
    Accept All, or after the run). ``run`` (CR-013 Phase C): each case's
    write, with change tracking on where the case says (``TrackAll``), then
    its region reported; tracking off at the end.
    """
    if mode not in ("reads", "report", "run"):
        raise ValueError(f"check 52 has no {mode!r} mode")
    body = package.body
    result: dict[str, Any] = {"check": 52, "mode": mode, "host": None, "reads": {}, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_52(p.text)) is not None]
    result["caseCount"] = len(regions)
    for region in regions:
        entry = {key: region[key] for key in ("n", "kind", "op", "tracked", "spec")}
        if mode == "reads":
            if "range" in region and region["kind"] not in result["reads"]:
                result["reads"][region["kind"]] = attempt(lambda region=region: reads_52(body, region))
            continue
        try:
            if mode == "run" and region["op"] != "read":
                package.change_tracking_mode = "TrackAll" if region["tracked"] else "Off"
                if "queries" in region:
                    entry["searches"] = searches_52(body, region)
                else:
                    if region["kind"] not in result["reads"]:
                        result["reads"][region["kind"]] = attempt(lambda region=region: reads_52(body, region))
                    entry["range"] = attempt(lambda region=region: vis(range_of_52(body, region).text))
                    entry["result"] = apply_52(package, body, region)
        except Exception as error:  # noqa: BLE001 - what Office JS refuses fails alone, as in the snippet
            entry["error"] = failed(error)
        entry["region"] = attempt(lambda region=region: report_52(body, region))
        result["cases"].append(entry)
    if mode == "run":
        package.change_tracking_mode = "Off"
    return result


# ---------------------------------------------------------------------------
# check 53 (core-ts check53-script-lab.js)
# ---------------------------------------------------------------------------

_BEGIN_53 = re.compile(r"^BEGIN(\d+) (\w+)(?: (.*))?$")


def parse_region_53(text: str) -> dict[str, Any] | None:
    """A BEGIN paragraph's text as a case, or None; the arguments upper-cased where they are words of the content."""
    match = _BEGIN_53.match(text.strip())
    if match is None:
        return None
    op = match.group(2)
    args = [a for a in (match.group(3) or "").split(" ") if a]
    return {"n": int(match.group(1)), "op": op, "args": args if op in ("control", "placeholder") else [a.upper() for a in args]}


def _paragraph_of_53(body: Any, token: str) -> Any:
    return find(body, token).paragraphs[0]


def _region_of_53(body: Any, n: int) -> Any:
    return _paragraph_of_53(body, f"BEGIN{n}").get_range("Whole").expand_to(
        _paragraph_of_53(body, f"END{n}").get_range("Whole")
    )


def _where(range_: Any) -> dict[str, Any]:
    """A collapsed or short range: its text and the paragraph it is in."""
    return {"text": vis(range_.text), "isEmpty": range_.is_empty, "paragraph": first_paragraph_text(range_)}


def range_reads_53(range_: Any) -> dict[str, Any]:
    """Every read a range has, each alone (the snippet's ``rangeReads``)."""
    out: dict[str, Any] = {}
    out["range"] = attempt(
        lambda: {
            "text": vis(range_.text),
            "isEmpty": range_.is_empty,
            "style": range_.style,
            "styleBuiltIn": range_.style_built_in,
        }
    )
    out["font"] = attempt(
        lambda: font_of(range_, ("bold", "italic", "underline", "name", "size", "color", "highlightColor"))
    )
    for location in ("Whole", "Content", "Start", "End", "After"):
        out[location] = attempt(lambda location=location: _where(range_.get_range(location)))
    out["paragraphs"] = attempt(
        lambda: [{"text": vis(p.text), "style": p.style, "styleBuiltIn": p.style_built_in} for p in range_.paragraphs]
    )
    out["tables"] = attempt(
        lambda: [
            {"rowCount": table.row_count, "values": [[vis(v) for v in row] for row in table.values]}
            for table in range_.tables
        ]
    )
    out["contentControls"] = attempt(
        lambda: [{"tag": control.tag, "text": vis(control.text)} for control in range_.content_controls]
    )
    return out


def changes_of_53(changes: list[Any]) -> list[dict[str, Any]]:
    """Tracked changes as type, text, and their range's text and paragraph count."""
    out = []
    for change in changes:
        found = change.get_range()
        out.append(
            {
                "type": change.type,
                "text": vis(change.text),
                "range": vis(found.text),
                "rangeParagraphs": len(found.paragraphs),
            }
        )
    return out


def _op_range(body: Any, args: list[str], n: int) -> Any:
    return range_reads_53(find(body, args[0]).expand_to(find(body, args[1])))


def _op_para(body: Any, args: list[str], n: int) -> Any:
    p = _paragraph_of_53(body, args[0])
    out: dict[str, Any] = {}
    out["paragraph"] = attempt(lambda: {"text": vis(p.text), "style": p.style, "styleBuiltIn": p.style_built_in})

    def whole() -> dict[str, Any]:
        w = p.get_range("Whole")
        return {"text": vis(w.text), "style": w.style, "styleBuiltIn": w.style_built_in}

    out["whole"] = attempt(whole)
    out["trackedChanges"] = attempt(lambda: changes_of_53(p.get_tracked_changes()))
    out["wholeTrackedChanges"] = attempt(lambda: changes_of_53(p.get_range("Whole").get_tracked_changes()))
    return out


def _op_control(body: Any, args: list[str], n: int) -> Any:
    control = control_by_tag(body, args[0])
    out: dict[str, Any] = {}
    out["control"] = attempt(
        lambda: {
            "text": vis(control.text),
            "type": control.type,
            "paragraphs": [vis(p.text) for p in control.paragraphs],
        }
    )
    for location in ("Whole", "Content", "Start", "End"):

        def at(location: str = location) -> dict[str, Any]:
            r = control.get_range(location)
            return {"text": vis(r.text), "paragraphs": [vis(p.text) for p in r.paragraphs]}

        out[location] = attempt(at)
    return out


def _op_table(body: Any, args: list[str], n: int) -> Any:
    table = find(body, args[0]).parent_table
    return {location: attempt(lambda location=location: _where(table.get_range(location))) for location in ("Whole", "Start", "End", "After")}


def _op_cellbody(body: Any, args: list[str], n: int) -> Any:
    cell_body = find(body, args[0]).parent_table_cell.body
    out: dict[str, Any] = {}
    for location in ("Whole", "Content", "Start", "End"):

        def at(location: str = location) -> dict[str, Any]:
            reads = range_reads_53(cell_body.get_range(location))
            return {
                "text": reads["range"]["text"],
                "paragraphs": reads["paragraphs"],
                "tables": reads["tables"],
                "contentControls": reads["contentControls"],
            }

        out[location] = attempt(at)
    return out


def _op_changes(body: Any, args: list[str], n: int) -> Any:
    region = _region_of_53(body, n)
    out: dict[str, Any] = {}
    out["region"] = attempt(lambda: changes_of_53(region.get_tracked_changes()))

    def paragraphs() -> list[dict[str, Any]]:
        return [
            {
                "text": vis(p.text),
                "changes": attempt(lambda p=p: changes_of_53(p.get_tracked_changes())),
                "wholeChanges": attempt(lambda p=p: changes_of_53(p.get_range("Whole").get_tracked_changes())),
            }
            for p in region.paragraphs
        ]

    out["paragraphs"] = attempt(paragraphs)
    return out


def _op_stories(body: Any, args: list[str], n: int) -> Any:
    def across() -> dict[str, Any]:
        notes = member(body, "footnotes", "footnotes")
        return {"text": vis(find(body, args[0]).expand_to(notes[0].body.get_range("Whole")).text)}

    return attempt(across)


def _op_body(body: Any, args: list[str], n: int) -> Any:
    out: dict[str, Any] = {}
    out["text"] = attempt(lambda: {"length": len(body.text), "tail": vis(body.text[-16:])})
    for location in ("Whole", "Content", "Start", "End"):

        def at(location: str = location) -> dict[str, Any]:
            r = body.get_range(location)
            return {"length": len(r.text), "tail": vis(r.text[-16:]), "firstParagraph": vis(r.paragraphs[0].text)}

        out[location] = attempt(at)
    return out


def _op_placeholder(body: Any, args: list[str], n: int) -> Any:
    control = control_by_tag(body, args[0])
    out: dict[str, Any] = {}
    out["control"] = attempt(
        lambda: {
            "text": vis(control.text),
            "placeholderText": vis(control.placeholder_text),
            "paragraph": vis(control.paragraphs[0].text),
            "whole": vis(control.get_range("Whole").text),
        }
    )
    out["showingPlaceholderText"] = attempt(
        lambda: member(control, "showing_placeholder_text", "showingPlaceholderText")
    )
    return out


OPS_53: dict[str, Callable[[Any, list[str], int], Any]] = {
    "range": _op_range,
    "para": _op_para,
    "control": _op_control,
    "table": _op_table,
    "cellbody": _op_cellbody,
    "changes": _op_changes,
    "stories": _op_stories,
    "body": _op_body,
    "placeholder": _op_placeholder,
}


def run_check_53(package: Any, label: str) -> dict[str, Any]:
    """``CHECK53.runCheck`` on a loaded package; `label` names the view, as the snippet's."""
    body = package.body
    result: dict[str, Any] = {"check": 53, "label": label, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_53(p.text)) is not None]
    for region in regions:
        op = OPS_53.get(region["op"])
        entry = {"n": region["n"], "op": region["op"], "args": region["args"]}
        if op is None:
            entry["reads"] = f"error: no op {region['op']}"
        else:
            entry["reads"] = attempt(lambda op=op, region=region: op(body, region["args"], region["n"]))
        result["cases"].append(entry)
    return result


# ---------------------------------------------------------------------------
# check 54 (core-ts check54-script-lab.js at 87bbb38; CR-013 Phase C)
# ---------------------------------------------------------------------------

_BEGIN_54 = re.compile(r"^BEGIN(\d+) (\w+) ([\w-]+) (untracked|tracked) (.*)$")

#: What the ``ooxml`` operation puts in place of the range: a flat OPC package whose body is one paragraph, NEWOOXML.
PARAGRAPH_OOXML = (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    "</pkg:xmlData></pkg:part>"
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    f'<w:document xmlns:w="{_W}"><w:body><w:p><w:r><w:t>NEWOOXML</w:t></w:r></w:p></w:body></w:document>'
    "</pkg:xmlData></pkg:part></pkg:package>"
)


def parse_region_54(text: str) -> dict[str, Any] | None:
    """A BEGIN paragraph's text as a case, or None (the snippet's ``parseRegion``)."""
    match = _BEGIN_54.match(text.strip())
    if match is None:
        return None
    region: dict[str, Any] = {
        "n": int(match.group(1)),
        "kind": match.group(2),
        "op": match.group(3),
        "tracked": match.group(4) == "tracked",
        "spec": match.group(5),
    }
    spec = region["spec"]
    patterns = [
        (r"^from (\S+) to (\S+)$", lambda m: {"from": m.group(1).upper(), "to": m.group(2).upper()}),
        (r"^from (\S+) to-mark (\S+)$", lambda m: {"from": m.group(1).upper(), "toMark": m.group(2).upper()}),
        (r"^from (\S+) to-table (\S+)$", lambda m: {"from": m.group(1).upper(), "toTable": m.group(2).upper()}),
        (r"^start (\S+) to-mark (\S+)$", lambda m: {"start": m.group(1).upper(), "toMark": m.group(2).upper()}),
        (r"^whole (\S+)$", lambda m: {"whole": m.group(1).upper()}),
        (r"^word (\S+)$", lambda m: {"word": m.group(1).upper()}),
    ]
    for pattern, make in patterns:
        found = re.match(pattern, spec)
        if found is not None:
            region["range"] = make(found)
            return region
    raise ValueError(f'BEGIN{region["n"]}: no range in "{spec}"')


def _paragraph_of_54(body: Any, token: str) -> Any:
    """The paragraph holding a token."""
    return find(body, token).paragraphs[0]


def range_of_54(body: Any, region: dict[str, Any]) -> Any:
    """The case's range, made afresh (the snippet's ``rangeOf``)."""
    r = region["range"]
    if "whole" in r:
        return _paragraph_of_54(body, r["whole"]).get_range("Whole")
    if "word" in r:
        return find(body, r["word"])
    if "toMark" in r:
        to = _paragraph_of_54(body, r["toMark"]).get_range("Whole")
    elif "toTable" in r:
        to = find(body, r["toTable"]).parent_table.get_range("Whole")
    else:
        to = find(body, r["to"])
    start = _paragraph_of_54(body, r["start"]).get_range("Start") if "start" in r else find(body, r["from"])
    return start.expand_to(to)


def region_of_54(body: Any, n: int) -> Any:
    """The region from its BEGIN paragraph to its END paragraph, both whole."""
    return _paragraph_of_54(body, f"BEGIN{n}").get_range("Whole").expand_to(
        _paragraph_of_54(body, f"END{n}").get_range("Whole")
    )


def apply_54(package: Any, body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """The case's operation, on its range made afresh; what the operation returns (the snippet's ``apply``)."""
    op = region["op"]
    n = region["n"]
    if op == "own-delete":
        # a paragraph inserted under tracking by whoever runs this, then a range from the heading into it deleted
        _paragraph_of_54(body, f"HEAD{n}A").insert_paragraph(f"OWN{n}A OWN{n}B", location="After")
        range_of_54(body, region).delete()
        return {}
    range_ = range_of_54(body, region)
    if op == "delete":
        range_.delete()
        return {}
    if op == "replace":
        return _returned(range_.insert_text("NEWTEXT", location="Replace"))
    if op == "paragraph-before":
        return _returned(range_.insert_paragraph("NEWPARA", location="Before"))
    if op == "paragraph-after":
        return _returned(range_.insert_paragraph("NEWPARA", location="After"))
    if op == "italic":
        range_.font.italic = True
        return {}
    if op == "size14":
        range_.font.size = 14
        return {}
    if op == "underline":
        range_.font.underline = "Single"
        return {}
    if op == "style":
        range_.style_built_in = "Heading2"
        return {}
    if op == "hyperlink":
        range_.hyperlink = "https://example.com/check54"
        return {}
    if op == "ooxml":
        views = range_.insert_ooxml(PARAGRAPH_OOXML, location="Replace")
        return {"returned": vis(views[0].text) if views else ""}
    if op == "control":
        from docx4j_py.wml import el

        control = range_.insert_content_control()
        # control.tag = ...: this package's tag is read-only and put_property writes it (CR-003 section 17.9 item 3)
        control.put_property(el.tag(val=f"new54-{n}"))
        return {"returned": vis(control.text), "type": control.type}
    raise ValueError(f"no operation {op}")


def report_54(body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """The region as it stands: text, paragraphs, tables, tracked changes, comments, controls, bookmarks, link."""
    out: dict[str, Any] = {}
    whole = region_of_54(body, region["n"])
    out["text"] = attempt(lambda: vis(whole.text))

    def paragraphs() -> list[dict[str, Any]]:
        entries = []
        for p in whole.paragraphs:
            entry: dict[str, Any] = {
                "text": vis(p.text),
                "style": p.style,
                "styleBuiltIn": p.style_built_in,
                "alignment": p.alignment,
                "tableNestingLevel": p.table_nesting_level,
            }
            listed = _list_of(p)
            if listed is not None:
                entry["list"] = listed
            entries.append(entry)
        return entries

    out["paragraphs"] = attempt(paragraphs)
    out["tables"] = attempt(
        lambda: [
            {"rowCount": table.row_count, "values": [[vis(value) for value in row] for row in table.values]}
            for table in whole.tables
        ]
    )
    out["trackedChanges"] = attempt(
        lambda: [
            {"type": change.type, "text": vis(change.text), "range": vis(change.get_range().text)}
            for change in whole.get_tracked_changes()
        ]
    )
    out["comments"] = attempt(
        lambda: [
            {"content": vis(comment.content), "range": vis(comment.get_range().text)}
            for comment in whole.get_comments()
        ]
    )
    out["contentControls"] = attempt(
        lambda: [{"tag": control.tag, "type": control.type, "text": vis(control.text)} for control in whole.content_controls]
    )
    out["bookmarks"] = attempt(lambda: member(whole, "get_bookmarks", "getBookmarks")(True, True))
    out["hyperlink"] = attempt(lambda: whole.hyperlink)
    return out


def run_check_54(package: Any, mode: str) -> dict[str, Any]:
    """``CHECK54.runCheck`` in its ``run`` or ``report`` mode, on a loaded package.

    ``run``: each case's write, with change tracking on where the case says,
    then its region reported; ``report``: every region reported, nothing
    changed (for Word's save after Accept All).
    """
    if mode not in ("run", "report"):
        raise ValueError(f"check 54 has no {mode!r} mode")
    body = package.body
    result: dict[str, Any] = {"check": 54, "mode": mode, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_54(p.text)) is not None]
    result["caseCount"] = len(regions)
    for region in regions:
        entry = {key: region[key] for key in ("n", "kind", "op", "tracked", "spec")}
        try:
            if mode == "run":
                package.change_tracking_mode = "TrackAll" if region["tracked"] else "Off"
                if region["op"] != "own-delete":
                    entry["range"] = attempt(lambda region=region: vis(range_of_54(body, region).text))
                entry["result"] = apply_54(package, body, region)
        except Exception as error:  # noqa: BLE001 - what Office JS refuses fails alone, as in the snippet
            entry["error"] = failed(error)
        entry["region"] = attempt(lambda region=region: report_54(body, region))
        result["cases"].append(entry)
    if mode == "run":
        package.change_tracking_mode = "Off"
    return result


# ---------------------------------------------------------------------------
# check 57 (core-ts check57-script-lab.js at 906d5e3; CR-013 section 13.7)
# ---------------------------------------------------------------------------

_BEGIN_57 = re.compile(r"^BEGIN(\d+) ([\w-]+) ([\w-]+)$")


def failed_57(error: BaseException) -> str:
    """Check 57's ``failed``: ``'error: ' + (code or 'no code') + (': ' + message)``."""
    if isinstance(error, NotImplementedMember):
        return f"error: NotImplemented: {error}"
    code = getattr(error, "code", None) or "no code"
    message = str(error)
    return f"error: {code}" + (f": {message}" if message else "")


def attempt_57(read: Callable[[], Any]) -> Any:
    try:
        return read()
    except Exception as error:  # noqa: BLE001 - every failure is the read's answer, as in the snippet
        return failed_57(error)


def parse_region_57(text: str) -> dict[str, Any] | None:
    """A BEGIN paragraph's text as a case, or None (the snippet's ``parseRegion``)."""
    match = _BEGIN_57.match(text.strip())
    return {"n": int(match.group(1)), "op": match.group(2), "shape": match.group(3)} if match else None


def region_of_57(body: Any, n: int) -> Any:
    """The region from its BEGIN paragraph to its END paragraph, both whole."""
    return find(body, f"BEGIN{n}").paragraphs[0].get_range("Whole").expand_to(
        find(body, f"END{n}").paragraphs[0].get_range("Whole")
    )


def apply_57(package: Any, body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """The case's write, the equations' under TrackAll; what it returns (the snippet's ``apply``)."""
    tracked = region["shape"].endswith("-equation")
    package.change_tracking_mode = "TrackAll" if tracked else "Off"
    out: dict[str, Any] = {"changeTrackingMode": package.change_tracking_mode}
    n = region["n"]
    try:
        op = region["op"]
        if op == "insertText-after":
            out["returned"] = vis(find(body, f"FA{n}").insert_text("x\ny", location="After").text)
        elif op == "insertParagraph-after":
            out["returned"] = vis(find(body, f"FA{n}").insert_paragraph(f"NEWPARA{n}", location="After").text)
        elif op == "paragraph-delete":
            region_of_57(body, n).paragraphs[1].delete()
        elif op == "range-delete":
            if region["shape"] == "inline-equation":
                find(body, f"L{n}").expand_to(find(body, f"R{n}")).delete()
            else:
                region_of_57(body, n).paragraphs[1].get_range("Content").delete()
        else:
            raise ValueError(f"no operation {op}")
        return out
    finally:
        package.change_tracking_mode = "Off"


def _tidy_57(xml: str) -> str:
    return re.sub(r' w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"| xmlns:\w+="[^"]*"', "", xml)


def body_of_57(xml: str) -> Any:
    """A range's OOXML: the body's content without its section properties (or the range's own XML, as given)."""
    if "pkg:package" not in xml:
        return vis(_tidy_57(xml))
    at = xml.find('pkg:name="/word/document.xml"')
    start = xml.find("<pkg:xmlData>", at)
    end = xml.find("</pkg:xmlData>", start)
    found = re.search(r"<w:body>([\s\S]*?)</w:body>", "" if at < 0 or start < 0 or end < 0 else xml[start:end])
    return vis(_tidy_57(re.sub(r"<w:sectPr\b[\s\S]*</w:sectPr>", "", found.group(1)))) if found else None


def report_57(body: Any, region: dict[str, Any]) -> dict[str, Any]:
    """The region as it stands: its text, paragraphs, OOXML, tracked changes and fields."""
    out: dict[str, Any] = {}
    whole = region_of_57(body, region["n"])
    out["text"] = attempt_57(lambda: vis(whole.text))
    out["paragraphs"] = attempt_57(lambda: [vis(p.text) for p in whole.paragraphs])
    out["ooxml"] = attempt_57(lambda: body_of_57(whole.get_xml()))
    out["trackedChanges"] = attempt_57(
        lambda: [
            {"type": change.type, "text": vis(change.text), "author": change.author}
            for change in whole.get_tracked_changes()
        ]
    )

    def fields() -> list[dict[str, Any]]:
        found = member(whole, "fields", "fields")
        return [{"code": vis(field.code), "result": vis(field.result.text)} for field in found]

    out["fields"] = attempt_57(fields)
    return out


def run_check_57(package: Any, mode: str) -> dict[str, Any]:
    """``CHECK57.runCheck`` in its ``run`` or ``report`` mode, on a loaded package."""
    if mode not in ("run", "report"):
        raise ValueError(f"check 57 has no {mode!r} mode")
    body = package.body
    result: dict[str, Any] = {"check": 57, "mode": mode, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_57(p.text)) is not None]
    result["caseCount"] = len(regions)
    for region in regions:
        entry: dict[str, Any] = {"n": region["n"], "op": region["op"], "shape": region["shape"]}
        try:
            if mode == "run" and region["op"] != "open":
                entry["result"] = apply_57(package, body, region)
        except Exception as error:  # noqa: BLE001 - what Office JS refuses fails alone, as in the snippet
            entry["error"] = failed_57(error)
        entry["region"] = attempt_57(lambda region=region: report_57(body, region))
        result["cases"].append(entry)
    return result


# ---------------------------------------------------------------------------
# check 58 (core-ts check58-script-lab.js at 222cc20; CR-016 Phase B)
# ---------------------------------------------------------------------------

_BEGIN_58 = re.compile(r"^BEGIN(\d+) (\w+)(?: (.*))?$")

WILD = {"match_wildcards": True}

#: The queries a ``search`` region is searched with, paragraph by paragraph: ``(query, options)``.
QUERIES_58: dict[str, Callable[[int], list[tuple[str, dict[str, Any]]]]] = {
    # a control showing its placeholder: Word's own text, a word of it, and a hit from plain text into it
    "placeholder": lambda n: [("Click or tap here to enter text.", {}), ("tap", TOKEN), (f"PLAIN{n}A Click", {})],
    # a surrogate pair (U+1F600) and an Arabic word with its fathas (combining marks)
    "wildcards": lambda n: [("?", WILD), ("[!a-z ]", WILD), ("a?b", WILD), ("x?y", WILD), ("[!a-z ]@", WILD)],
}

#: The ops the Edit run runs; the Read run reads every case and changes nothing.
EDITS_58 = frozenset({"accept", "reject", "fill"})


def parse_region_58(text: str) -> dict[str, Any] | None:
    """A BEGIN paragraph's text as a case, or None; a ``range`` op's word arguments upper-cased, as the content has them."""
    match = _BEGIN_58.match(text.strip())
    if match is None:
        return None
    op = match.group(2)
    args = [a for a in (match.group(3) or "").split(" ") if a]
    return {"n": int(match.group(1)), "op": op, "args": [a.upper() for a in args] if op == "range" else args}


def _paragraph_of_58(body: Any, token: str) -> Any:
    return find(body, token).paragraphs[0]


def region_of_58(body: Any, n: int) -> Any:
    """The region's range: its BEGIN paragraph's whole range expanded to its END paragraph's."""
    return _paragraph_of_58(body, f"BEGIN{n}").get_range("Whole").expand_to(
        _paragraph_of_58(body, f"END{n}").get_range("Whole")
    )


def _content_paragraphs_58(body: Any, n: int) -> list[Any]:
    """The region's paragraphs between its BEGIN and END paragraphs."""
    return region_of_58(body, n).paragraphs[1:-1]


def changes_of_58(changes: list[Any]) -> list[dict[str, Any]]:
    """Tracked changes as type, text, and their range's text and paragraph count (the snippet's ``changesOf``)."""
    out = []
    for change in changes:
        found = change.get_range()
        out.append(
            {
                "type": change.type,
                "text": vis(change.text),
                "range": vis(found.text),
                "rangeParagraphs": len(found.paragraphs),
            }
        )
    return out


def _changes_58(body: Any, n: int) -> dict[str, Any]:
    """The region's changes, and each paragraph's, its whole range's and its content range's."""
    region = region_of_58(body, n)
    out: dict[str, Any] = {}
    out["region"] = attempt(lambda: changes_of_58(region.get_tracked_changes()))

    def paragraphs() -> list[dict[str, Any]]:
        return [
            {
                "text": vis(p.text),
                "changes": attempt(lambda p=p: changes_of_58(p.get_tracked_changes())),
                "wholeChanges": attempt(lambda p=p: changes_of_58(p.get_range("Whole").get_tracked_changes())),
                "contentChanges": attempt(lambda p=p: changes_of_58(p.get_range("Content").get_tracked_changes())),
            }
            for p in region.paragraphs
        ]

    out["paragraphs"] = attempt(paragraphs)
    return out


def _placeholder_58(body: Any, tag: str) -> dict[str, Any]:
    """A control showing its placeholder: its text and placeholder text, its paragraph's text and ranges, its own ranges."""
    out: dict[str, Any] = {}

    def control() -> dict[str, Any]:
        found = control_by_tag(body, tag)
        p = found.get_range("Whole").paragraphs[0]
        return {
            "text": vis(found.text),
            "placeholderText": vis(found.placeholder_text),
            "paragraph": vis(p.text),
            "paragraphWhole": vis(p.get_range("Whole").text),
            "whole": vis(found.get_range("Whole").text),
            "content": vis(found.get_range("Content").text),
        }

    out["control"] = attempt(control)
    out["showingPlaceholderText"] = attempt(
        lambda: member(control_by_tag(body, tag), "showing_placeholder_text", "showingPlaceholderText")
    )
    return out


def _region_after_58(body: Any, n: int) -> dict[str, Any]:
    """After an edit: the region's paragraphs and its changes."""
    return {
        "paragraphs": attempt(lambda: [vis(p.text) for p in _content_paragraphs_58(body, n)]),
        "changes": attempt(lambda: changes_of_58(region_of_58(body, n).get_tracked_changes())),
    }


def _resolve_part_58(body: Any, n: int, i: int, verb: str) -> dict[str, Any]:
    """The first change paragraph `i` of region `n` lists (its BEGIN paragraph is 0), accepted or rejected; the region after."""

    def resolved() -> Any:
        listed = region_of_58(body, n).paragraphs[i].get_tracked_changes()
        if not listed:
            return "none listed"
        change = listed[0]
        what = {"type": change.type, "text": vis(change.text), "listed": len(listed)}
        getattr(change, verb)()
        return what

    out: dict[str, Any] = {"resolved": attempt(resolved)}
    out["after"] = _region_after_58(body, n)
    return out


def _op_range_58(body: Any, args: list[str], n: int, mode: str) -> Any:
    found = find(body, args[0]).expand_to(find(body, args[1]))
    out: dict[str, Any] = {}
    out["range"] = attempt(lambda: {"text": vis(found.text), "isEmpty": found.is_empty})
    out["paragraphs"] = attempt(lambda: [vis(p.text) for p in found.paragraphs])
    out["whole"] = attempt(lambda: vis(found.get_range("Whole").text))
    return out


def _op_cell_58(body: Any, args: list[str], n: int, mode: str) -> Any:
    table = region_of_58(body, n).tables[0]
    out: dict[str, Any] = {}
    out["table"] = attempt(
        lambda: {"values": [[vis(v) for v in row] for row in table.values], "whole": vis(table.get_range("Whole").text)}
    )

    def cell() -> dict[str, Any]:
        found = table.cell(0, 0)
        return {
            "value": vis(found.value),
            "bodyText": vis(found.body.text),
            "whole": vis(found.body.get_range("Whole").text),
            "paragraphs": [
                {"text": vis(p.text), "whole": vis(p.get_range("Whole").text)} for p in found.body.paragraphs
            ],
        }

    out["cell"] = attempt(cell)
    return out


def _op_search_58(body: Any, args: list[str], n: int, mode: str) -> Any:
    paragraphs = _content_paragraphs_58(body, n)
    out: dict[str, Any] = {}
    for query, options in QUERIES_58[args[0]](n):
        out[vis(query)] = attempt(
            lambda query=query, options=options: [
                [vis(hit.text) for hit in p.search(query, **options)] for p in paragraphs
            ]
        )
    return out


def _op_resolve_58(verb: str) -> Callable[[Any, list[str], int, str], Any]:
    def op(body: Any, args: list[str], n: int, mode: str) -> Any:
        if mode != "edit":
            return _changes_58(body, n)
        return _resolve_part_58(body, n, int(args[0]), verb)

    return op


def _op_fill_58(body: Any, args: list[str], n: int, mode: str) -> Any:
    tag, how = args
    if mode != "edit":
        return _placeholder_58(body, tag)
    out: dict[str, Any] = {}
    out["inserted"] = attempt(
        lambda: vis(
            control_by_tag(body, tag).insert_text(f"FILL{n}", location="Replace" if how == "replace" else "End").text
        )
    )
    out["after"] = _placeholder_58(body, tag)
    return out


OPS_58: dict[str, Callable[[Any, list[str], int, str], Any]] = {
    "changes": lambda body, args, n, mode: _changes_58(body, n),
    "range": _op_range_58,
    "cell": _op_cell_58,
    "placeholder": lambda body, args, n, mode: _placeholder_58(body, args[0]),
    "search": _op_search_58,
    "accept": _op_resolve_58("accept"),
    "reject": _op_resolve_58("reject"),
    "fill": _op_fill_58,
}


def run_check_58(package: Any, label: str, mode: str = "read") -> dict[str, Any]:
    """``CHECK58.runCheck`` in its ``read`` or ``edit`` mode, on a loaded package; `label` names the view.

    ``read`` reads every case and changes nothing; ``edit`` accepts, rejects or
    fills in its edit cases only, and reads each region after (the snippet at
    core-ts ``222cc20``, byte for byte the share's ``check58-script-lab.js``).
    """
    if mode not in ("read", "edit"):
        raise ValueError(f"check 58 has no {mode!r} mode")
    body = package.body
    result: dict[str, Any] = {"check": 58, "label": label, "mode": mode, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_58(p.text)) is not None]
    for region in regions:
        if mode == "edit" and region["op"] not in EDITS_58:
            continue
        op = OPS_58.get(region["op"])
        entry: dict[str, Any] = {"n": region["n"], "op": region["op"], "args": region["args"]}
        if op is None:
            entry["reads"] = f"error: no op {region['op']}"
        else:
            entry["reads"] = attempt(lambda op=op, region=region: op(body, region["args"], region["n"], mode))
        result["cases"].append(entry)
    return result


# ---------------------------------------------------------------------------
# check 59 (core-ts check59-script-lab.js at 222cc20 and 3ba1e68; CR-016 Phase C)
# ---------------------------------------------------------------------------

_BEGIN_59 = re.compile(r"^BEGIN(\d+) ([\w-]+)$")

#: Across the region's content: a mark between two letters, Word's paragraph mark code alone.
REGION_QUERIES_59 = ("Q?Z", "Q^13Z", "Q*Z", "^13")

#: Within a paragraph, a cell's body, a control: every character, every one but a capital, the mark code.
PART_QUERIES_59 = ("?", "[!A-Z]", "^13")


def parse_region_59(text: str) -> dict[str, Any] | None:
    """A BEGIN paragraph's text as a case, ``{n, shape}``, or None (the snippet's ``parseRegion``)."""
    match = _BEGIN_59.match(text.strip())
    if match is None:
        return None
    return {"n": int(match.group(1)), "shape": match.group(2)}


def region_of_59(body: Any, n: int) -> Any:
    """The region's range: its BEGIN paragraph's whole range expanded to its END paragraph's (the snippet's ``regionOf``)."""
    return region_of_58(body, n)


def searched_59(scope: Any, queries: tuple[str, ...] | list[str]) -> dict[str, Any]:
    """Each query's hits in a scope, as their text, each searched alone with wildcards (the snippet's ``searched``)."""
    return {
        vis(query): attempt(lambda query=query: [vis(hit.text) for hit in scope.search(query, **WILD)])
        for query in queries
    }


def _read_59(body: Any, n: int) -> dict[str, Any]:
    """One region read: its range's searches, each content paragraph's, the first table's first cell's body and whole range's, the first control's."""
    region = region_of_59(body, n)
    out: dict[str, Any] = {"region": searched_59(region, REGION_QUERIES_59)}
    out["paragraphs"] = attempt(
        lambda: [{"text": vis(p.text), "hits": searched_59(p, PART_QUERIES_59)} for p in region.paragraphs[1:-1]]
    )
    tables = region.tables
    controls = region.content_controls
    if tables:
        out["cell"] = attempt(lambda: searched_59(tables[0].cell(0, 0).body, PART_QUERIES_59))
        out["table"] = attempt(lambda: searched_59(tables[0].get_range("Whole"), [*PART_QUERIES_59, "Q?Z", "Q*Z"]))
    if controls:
        out["control"] = attempt(lambda: searched_59(controls[0], [*PART_QUERIES_59, "Q?Z", "Q*Z"]))
    return out


def run_check_59(package: Any, label: str) -> dict[str, Any]:
    """``CHECK59.runCheck`` on a loaded package; `label` names the view. Read only: nothing changes.

    Each case is a region ``BEGIN<n> <shape>`` ... ``END<n>`` searched with
    wildcards in four scopes: the region's range, each content paragraph, the
    first table's first cell's body and the table's ``"Whole"`` range, and
    the first content control (the snippet at core-ts ``222cc20``, byte for
    byte the share's ``check59-script-lab.js``, ``abd3bfa5...``).
    """
    body = package.body
    result: dict[str, Any] = {"check": 59, "label": label, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_59(p.text)) is not None]
    for region in regions:
        result["cases"].append(
            {
                "n": region["n"],
                "shape": region["shape"],
                "reads": attempt(lambda region=region: _read_59(body, region["n"])),
            }
        )
    return result


# ---------------------------------------------------------------------------
# checks 63 and 64 (core-ts check63-script-lab.js and check64-script-lab.js at 020c039; CR-019 Phase A)
# ---------------------------------------------------------------------------

_BEGIN_63 = re.compile(r"^BEGIN(\d+) ([\w-]+) ([\w-]+)(?: (untracked|tracked))?$")
_BEGIN_64 = re.compile(r"^BEGIN(\d+) ([\w-]+) ([\w-]+) (untracked|tracked)$")

#: Each check's custom XML part's namespace and its controls' tags (the snippets' ``DATA_NS`` and ``tagOf``).
DATA_NS_63 = "urn:check63"
DATA_NS_64 = "urn:check64"


def parse_region_63(text: str) -> dict[str, Any] | None:
    """A check 63 BEGIN paragraph's text as a case, ``{n, op, shape[, mode]}``, or None (the snippet's ``parseRegion``)."""
    match = _BEGIN_63.match(text.strip())
    if match is None:
        return None
    region: dict[str, Any] = {"n": int(match.group(1)), "op": match.group(2), "shape": match.group(3)}
    if match.group(4):
        region["mode"] = match.group(4)
    return region


def parse_region_64(text: str) -> dict[str, Any] | None:
    """A check 64 BEGIN paragraph's text as a case, ``{n, op, shape, mode}``, or None (the snippet's ``parseRegion``)."""
    match = _BEGIN_64.match(text.strip())
    if match is None:
        return None
    return {"n": int(match.group(1)), "op": match.group(2), "shape": match.group(3), "mode": match.group(4)}


def _paragraph_of_63(body: Any, token: str) -> Any:
    """The paragraph holding a token (the snippets' ``paragraphOf``)."""
    return find(body, token).paragraphs[0]


def region_of_63(body: Any, n: int) -> Any:
    """The region from its BEGIN paragraph to its END paragraph, both whole (the snippets' ``regionOf``)."""
    return (
        _paragraph_of_63(body, f"BEGIN{n}")
        .get_range("Whole")
        .expand_to(_paragraph_of_63(body, f"END{n}").get_range("Whole"))
    )


def _content_paragraphs_63(body: Any, n: int) -> list[Any]:
    """The region's paragraphs between its BEGIN and END paragraphs (the snippets' ``contentParagraphs``)."""
    return region_of_63(body, n).paragraphs[1:-1]


def _control_of_63(body: Any, tag: str) -> dict[str, Any] | None:
    """The case's control, ``{control, via}``, through ``body.getContentControls()``, the snippets' first way; or None.

    Their second way, ``body.contentControls``, is this package's
    :attr:`~docx4j_py.model.content.Body.content_controls`, which lists every
    kind (CR-013 section 11, G1), so the first way finds whatever it would.
    """
    mine = [control for control in body.get_content_controls() if control.tag == tag]
    return {"control": mine[0], "via": "body.getContentControls()"} if mine else None


def _paragraph_for_63(body: Any, region: dict[str, Any], control: Any, *, block: bool) -> Any:
    """The paragraph a case's paragraph op writes: the control's first (block) or the one holding it (run)."""
    if block:
        return control.paragraphs[0]
    return _paragraph_of_63(body, f"PLAIN{region['n']}A")


def _apply_63(body: Any, region: dict[str, Any], tag: str, *, block: bool) -> dict[str, Any]:
    """The case's write; what its result reads (the snippets' ``apply``)."""
    found = _control_of_63(body, tag)
    if found is None:
        raise LookupError(f"not found: {tag}")
    c = found["control"]
    via = found["via"]
    text = f"NEW{region['n']}"

    def read(r: Any) -> dict[str, Any]:
        return {"via": via, "returned": vis(r.text)}

    op = region["op"]
    if op == "control-text-replace":
        return read(c.insert_text(text, location="Replace"))
    if op == "control-text-start":
        return read(c.insert_text(text, location="Start"))
    if op == "control-text-end":
        return read(c.insert_text(text, location="End"))
    if op == "control-paragraph-end":
        return read(c.insert_paragraph(text, location="End"))
    if op == "content-text-replace":
        return read(c.get_range("Content").insert_text(text, location="Replace"))
    if op == "content-text-end":
        return read(c.get_range("Content").insert_text(text, location="End"))
    if op == "content-delete":
        c.get_range("Content").delete()
        return {"via": via}
    if op == "whole-text-replace":
        return read(c.get_range("Whole").insert_text(text, location="Replace"))
    paragraph = _paragraph_for_63(body, region, c, block=block) if op.startswith("paragraph-") else None
    if op == "paragraph-text-replace":
        return read(paragraph.insert_text(text, location="Replace"))
    if op == "paragraph-text-start":
        return read(paragraph.insert_text(text, location="Start"))
    if op == "paragraph-text-end":
        return read(paragraph.insert_text(text, location="End"))
    if op == "paragraph-after":
        return read(paragraph.insert_paragraph(text, location="After"))
    if op == "paragraph-before":
        return read(paragraph.insert_paragraph(text, location="Before"))
    if op == "paragraph-content-delete":
        paragraph.get_range("Content").delete()
        return {"via": via}
    raise ValueError(f"no operation {op}")


def node_of_63(xml: str, n: int) -> Any:
    """``value<n>``'s content in the custom XML part's XML, as it stands there (not parsed); None if it isn't there."""
    name = rf"(?:[\w.-]+:)?value{n}"
    match = re.search(rf"<{name}(?:\s[^>]*)?(?:/>|>([\s\S]*?)</{name}>)", xml)
    if match is None:
        return None
    return vis("" if match.group(1) is None else match.group(1))


def _report_63(
    package: Any, body: Any, region: dict[str, Any], tag: str, data_ns: str, *, node: bool
) -> dict[str, Any]:
    """The region as it stands: its paragraphs' text, its OOXML, the control's reads, and a bound control's node."""
    n = region["n"]
    out: dict[str, Any] = {}
    out["paragraphs"] = attempt_57(lambda: [vis(p.text) for p in _content_paragraphs_63(body, n)])
    out["ooxml"] = attempt_57(lambda: body_of_57(region_of_63(body, n).get_xml()))
    try:
        found = _control_of_63(body, tag)
    except Exception:  # noqa: BLE001 - the snippet's controlOf(...).catch(() => null)
        found = None
    if found is None:
        out["control"] = None
    else:

        def control() -> dict[str, Any]:
            c = found["control"]
            return {
                "via": found["via"],
                "type": c.type,
                "text": vis(c.text),
                "placeholderText": vis(c.placeholder_text),
                "paragraphs": [vis(p.text) for p in c.paragraphs],
            }

        out["control"] = attempt_57(control)
        out["showingPlaceholderText"] = attempt_57(lambda: found["control"].showing_placeholder_text)
    if node:

        def read_node() -> Any:
            parts = package.custom_xml_parts.get_by_namespace(data_ns)
            return node_of_63(parts[0].get_xml(), n)

        out["node"] = attempt_57(read_node)
    return out


def _track_63(package: Any, on: bool) -> None:
    """Change tracking on or off, as a case's mode says (the snippets' ``track``)."""
    package.change_tracking_mode = "TrackAll" if on else "Off"


def run_check_63(package: Any, mode: str) -> dict[str, Any]:
    """``CHECK63.runCheck`` in its ``run`` or ``report`` mode, on a loaded package (the snippet at core-ts ``020c039``).

    ``report`` reads every region and changes nothing; ``run`` gives each case
    its one write, ``NEW<n>``, with change tracking on for a ``tracked`` case and
    off otherwise, then reads the region: its paragraphs' text, its OOXML (this
    package's ``Range.get_xml``, where Office JS's ``getOoxml`` is a flat OPC
    package: the body's blocks either way, tidied alike), the control's reads,
    and a bound control's node. What Office JS refuses fails alone, as
    ``'error: ' + code + ': ' + message`` (check 57's ``failed``, the snippet's).
    """
    if mode not in ("run", "report"):
        raise ValueError(f"check 63 has no {mode!r} mode")
    body = package.body
    result: dict[str, Any] = {"check": 63, "mode": mode, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_63(p.text)) is not None]
    result["caseCount"] = len(regions)
    try:
        for region in regions:
            tag = f"c63-{region['n']}"
            block = region["shape"].startswith("block")
            entry = dict(region)
            if mode == "run" and region["op"] != "open":
                try:
                    _track_63(package, region.get("mode") == "tracked")
                    entry["result"] = _apply_63(body, region, tag, block=block)
                except Exception as error:  # noqa: BLE001 - what Office JS refuses fails alone, as in the snippet
                    entry["error"] = failed_57(error)
                _track_63(package, False)
            entry["after"] = _report_63(package, body, region, tag, DATA_NS_63, node=region["shape"].endswith("-bound"))
            result["cases"].append(entry)
    finally:
        _track_63(package, False)
    return result


def _make_64(body: Any, region: dict[str, Any], tag: str) -> dict[str, Any]:
    """A ``made-`` case's control, made with ``insert_content_control``: around the region's empty paragraph, or at the
    point after ``PLAIN<n>A`` (the end of that word if ``"After"`` is refused), tagged (the snippet's ``make``)."""
    from docx4j_py.wml import el

    n = region["n"]
    out: dict[str, Any] = {}
    if region["shape"] == "made-block":
        control = _content_paragraphs_63(body, n)[0].insert_content_control()
        # control.tag = ...: this package's tag is read-only and put_property writes it (CR-003 section 17.9 item 3)
        control.put_property(el.tag(val=tag))
        out["at"] = "paragraph"
    else:
        hit = find(body, f"PLAIN{n}A")
        for where in ("After", "End"):
            try:
                control = hit.get_range(where).insert_content_control()
                control.put_property(el.tag(val=tag))
                out["at"] = where
                break
            except Exception as error:  # noqa: BLE001 - the snippet tries the next
                out[f"getRange({where})"] = failed_57(error)
        if "at" not in out:
            raise LookupError(f"no control made: {out}")
    out["ooxml"] = attempt_57(lambda: body_of_57(region_of_63(body, n).get_xml()))
    return out


def run_check_64(package: Any, mode: str) -> dict[str, Any]:
    """``CHECK64.runCheck`` in its ``run`` or ``report`` mode, on a loaded package (the snippet at core-ts ``020c039``).

    Check 63's run, on controls in Word's form (``w:placeholder`` naming
    ``DefaultPlaceholder_-1854013440`` and Word's glossary), with
    ``control-text-start``; a ``made-`` case's region holds no control, and the
    run makes one first, tracking off (:func:`_make_64`), then writes. A bound
    control's node is read for the ``-bound`` and ``-value`` shapes.
    """
    if mode not in ("run", "report"):
        raise ValueError(f"check 64 has no {mode!r} mode")
    body = package.body
    result: dict[str, Any] = {"check": 64, "mode": mode, "host": None, "cases": []}
    regions = [region for p in body.paragraphs if (region := parse_region_64(p.text)) is not None]
    result["caseCount"] = len(regions)
    try:
        for region in regions:
            tag = f"c64-{region['n']}"
            shape = region["shape"]
            # the snippet's paragraphFor: a run's paragraph for a shape naming "run", else the control's first
            block = re.search(r"(^|-)run(-|$)", shape) is None
            entry = dict(region)
            if mode == "run" and shape.startswith("made-"):
                try:
                    _track_63(package, False)
                    entry["made"] = _make_64(body, region, tag)
                except Exception as error:  # noqa: BLE001 - the snippet's .catch(failed)
                    entry["made"] = failed_57(error)
            if mode == "run" and not isinstance(entry.get("made"), str):
                try:
                    _track_63(package, region["mode"] == "tracked")
                    entry["result"] = _apply_63(body, region, tag, block=block)
                except Exception as error:  # noqa: BLE001 - what Office JS refuses fails alone, as in the snippet
                    entry["error"] = failed_57(error)
                _track_63(package, False)
            entry["after"] = _report_63(
                package, body, region, tag, DATA_NS_64, node=re.search(r"-(bound|value)$", shape) is not None
            )
            result["cases"].append(entry)
    finally:
        _track_63(package, False)
    return result
