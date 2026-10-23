"""docx4j-core-ts's Office JS checks 52 and 53, ported to the content API: CR-013 Phase B's oracles.

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

Each read is attempted alone, as the snippets sync each read alone, so that
what fails fails alone, as ``'error: ' + (code or message)``. A member this
package does not offer (``Range.footnotes``, ``Paragraph.get_text_ranges``,
``Body.footnotes``, ``ContentControl.showing_placeholder_text``) is
``error: NotImplemented: ...``, as core-ts's shim reports one.

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
    "run_check_52",
    "run_check_53",
    "run_check_54",
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
