"""A write at a paragraph's end, as Word 365 places it: CR-019 Phase D, the port of core-ts's CR-002 sections 63.23 to 63.27.

    .venv-fork/bin/python -m pytest tests/content/test_paragraph_end.py

docx4j-core-ts ran Word 365 checks 69 and 70 on where Office JS's writes at a
paragraph's End go (its sections 63.23 to 63.26; Word 365 16.0.20326.20232,
Design Mode off, 2026-10-25), after this package's CR-019 Phase C found its own
port writing such text into a trailing placeholder control, where no reader saw
it (CR-019 section 13.4 item 1). The user decided D12 to D14 there ("D12: as
word; D13: ok; D14: as word"), and core-ts implemented them at ``09dc3d2`` (its
section 63.27): :meth:`Paragraph._write_at_end` here.

* **Check 69:** a paragraph ending in a run-level control showing its
  placeholder (Word's form, the bare form, or alone), a hidden run, a control
  holding text, or nothing more, written ``NEW<n>`` at its End by nine routes.
* **Check 70:** a paragraph ending in a bookmark, a comment's markers, a tracked
  deletion or insertion, a complex or simple field, a text box, or a control
  after a bold run or holding bold text, by check 69's routes but the body's.

``tests/office_js/checks.py`` ports the two snippets step for step
(:func:`run_check_69`, :func:`run_check_70`); this file holds this package's
run to Word's through core-ts's ``look69`` and ``look70``, then the regions of
this package's save against Word's save after the same run. What still differs
is pinned, with its reason: in check 70's run, Word's All Markup reading a
deletion's text, Word reading a ``w:fldSimple`` as a complex field, and Word's
``w:noProof`` after a text box (D13's departure); in the saves, the hidden
cases' run boundary, the simple field's form and ``w:noProof``. Then core-ts's
tests that are no replay: a link, a proofing mark, a paragraph of markers, the
tracked writes, line breaks and a search hit; and this package's own: the
bound controls' nodes and the routes core-ts's structure gives too.
"""

from __future__ import annotations

import hashlib
import importlib.util
import io
import json
import re
import zipfile
from typing import Any

import pytest
from conftest import ROOT

from docx4j_py import create_package, load
from docx4j_py.wml import to_xml

OFFICE_JS = ROOT / "tests" / "office_js"

_spec = importlib.util.spec_from_file_location("office_js_checks", OFFICE_JS / "checks.py")
checks = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(checks)

PLACEHOLDER = "Click or tap here to enter text."
DEFAULT_ENTRY = "DefaultPlaceholder_-1854013440"
AUTHOR_70 = "Check 70"

#: The files, as core-ts 09dc3d2 has them: each check's file (git-ignored there, written by build.mjs before Word's run:
#: 2026-10-25 13:39 for check 69's run at 14:13, 14:28 for check 70's at 15:45, each the same bytes as the share's
#: copy Word read), Word's two JSONs (the share's result-report.txt and result-run.txt, byte for byte) and Word's two
#: saves, which test/check69/ and test/check70/'s manifest.json list with these SHA-1s.
SHA1 = {
    "69-paragraph-end.docx": "7d34ddc02421749705da4fc0927f8410164a1acc",
    "69-paragraph-end-word365-report.json": "b1a528a2206cd05bd8304e08e12c53fbd88e1bb9",
    "69-paragraph-end-word365-run.json": "eb7379a98eee6dc7a0cf64b7dabd3120a3f6f653",
    "69-opened-word365.docx": "8230ae767568bcb3d924c07b4492efac0ca7b8a6",
    "69-run-word365.docx": "b55905f67f58fe512d29d373ebcebd08b462d305",
    "70-paragraph-endings.docx": "5be28b67e17be98f0822d366ac244fba51f0f024",
    "70-paragraph-endings-word365-report.json": "549686057efc66c65207205670b5cc9fe15ba42a",
    "70-paragraph-endings-word365-run.json": "27c71c8944bd0695bf4f173c961677ef4494322c",
    "70-opened-word365.docx": "f561e8dd762cd4bb1f399924019f14535adcd9eb",
    "70-run-word365.docx": "651a59bfeba71da1f62c76aa4eafed092c7beadb",
}

# ---------------------------------------------------------------------------
# the two files' cases and paragraphs (core-ts 09dc3d2's check69/build.mjs and check70/build.mjs)
# ---------------------------------------------------------------------------

SHAPES_69 = ["plain", "ph-word", "ph-bare", "hidden", "control-text", "ph-only"]
PARAGRAPH_OPS = ["paragraph-end", "end-before", "end-after", "end-start", "end-end", "end-replace", "content-end"]
SHAPES_70 = [
    "bookmark-span",
    "bookmark-point",
    "comment",
    "tracked-del",
    "tracked-ins",
    "field-complex",
    "field-simple",
    "textbox",
    "bold-ph",
    "bold-control-text",
    "control-bold-text",
]
OPS_70 = [*PARAGRAPH_OPS, "cell-body-end"]


def cases_69() -> list[dict[str, Any]]:
    """Check 69's 49 cases in document order (``check69Cases``): each shape by the seven paragraph routes and a cell's,
    and the document body's last."""
    out: list[dict[str, Any]] = []
    for shape in SHAPES_69:
        for op in [*PARAGRAPH_OPS, "cell-body-end"]:
            out.append({"n": len(out) + 1, "op": op, "shape": shape})
    out.append({"n": len(out) + 1, "op": "body-end", "shape": "ph-word"})
    return out


def cases_70() -> list[dict[str, Any]]:
    """Check 70's 88 cases in document order (``check70Cases``)."""
    out: list[dict[str, Any]] = []
    for shape in SHAPES_70:
        for op in OPS_70:
            out.append({"n": len(out) + 1, "op": op, "shape": shape})
    return out


def _t(text: str, r_pr: str = "") -> str:
    return f'<w:r>{f"<w:rPr>{r_pr}</w:rPr>" if r_pr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _run(inner: str, r_pr: str = "") -> str:
    return f'<w:r>{f"<w:rPr>{r_pr}</w:rPr>" if r_pr else ""}{inner}</w:r>'


_STYLE = '<w:rStyle w:val="PlaceholderText"/>'


def paragraph_69(case: dict[str, Any]) -> str:
    """A check 69 case's paragraph, as ``build.mjs`` writes it (``paragraphOf``)."""
    n = case["n"]

    def control(kind: str) -> str:
        entry = f'<w:placeholder><w:docPart w:val="{DEFAULT_ENTRY}"/></w:placeholder>' if kind == "word" else ""
        flag = "" if kind == "value" else "<w:showingPlcHdr/>"
        inner = _t(f"VALUE{n}") if kind == "value" else _t(PLACEHOLDER, _STYLE)
        return (
            f'<w:sdt><w:sdtPr><w:alias w:val="Check 69 control {n}"/><w:tag w:val="c69-{n}"/><w:id w:val="{69000 + n}"/>'
            f"{entry}{flag}</w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>"
        )

    lead = _t(f"TEXT{n} ")
    return "<w:p>" + {
        "plain": lead,
        "ph-word": lead + control("word"),
        "ph-bare": lead + control("bare"),
        "hidden": lead + _t(f"HIDDEN{n}", "<w:vanish/>"),
        "control-text": lead + control("value"),
        "ph-only": control("word"),
    }[case["shape"]] + "</w:p>"


def _text_box(identifier: int, text: str) -> str:
    """check70/build.mjs's ``textBox``: an inline text box in ``mc:AlternateContent`` with a plain-text Fallback."""
    wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    return _run(
        '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="1828800" cy="365760"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{identifier}" '
        f'name="Text Box {identifier}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="{wps}"><wps:wsp>'
        '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="365760"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="0050DC"/></a:solidFill></a:ln></wps:spPr>'
        f"<wps:txbx><w:txbxContent><w:p>{_t(text)}</w:p></w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" '
        'anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>'
        f'</mc:Choice><mc:Fallback><w:t xml:space="preserve">{text}</w:t></mc:Fallback></mc:AlternateContent>'
    )


def paragraph_70(case: dict[str, Any]) -> str:
    """A check 70 case's paragraph, as ``build.mjs`` writes it (``paragraphOf``)."""
    n = case["n"]
    revision = f'w:author="{AUTHOR_70}" w:date="2026-10-25T00:00:00Z"'
    bold = "<w:b/>"

    def control(kind: str) -> str:
        entry = (
            f'<w:placeholder><w:docPart w:val="{DEFAULT_ENTRY}"/></w:placeholder><w:showingPlcHdr/>' if kind == "ph" else ""
        )
        inner = _t(PLACEHOLDER, _STYLE) if kind == "ph" else _t(f"VALUE{n}", bold if kind == "bold-value" else "")
        return (
            f'<w:sdt><w:sdtPr><w:alias w:val="Check 70 control {n}"/><w:tag w:val="c70-{n}"/><w:id w:val="{70000 + n}"/>'
            f"{entry}</w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>"
        )

    lead = _t(f"TEXT{n} ")
    deleted = f'<w:delText xml:space="preserve">DELETED{n}</w:delText>'
    inner = {
        "bookmark-span": f'<w:bookmarkStart w:id="{n}" w:name="BM{n}"/>{lead}<w:bookmarkEnd w:id="{n}"/>',
        "bookmark-point": f'{lead}<w:bookmarkStart w:id="{n}" w:name="PT{n}"/><w:bookmarkEnd w:id="{n}"/>',
        "comment": f'<w:commentRangeStart w:id="{n}"/>{lead}<w:commentRangeEnd w:id="{n}"/>'
        + _run(f'<w:commentReference w:id="{n}"/>', '<w:rStyle w:val="CommentReference"/>'),
        "tracked-del": f'{lead}<w:del w:id="{7000 + n}" {revision}>' + _run(deleted) + "</w:del>",
        "tracked-ins": f'{lead}<w:ins w:id="{7000 + n}" {revision}>{_t(f"INSERTED{n}")}</w:ins>',
        "field-complex": lead
        + _run('<w:fldChar w:fldCharType="begin"/>')
        + _run(f'<w:instrText xml:space="preserve"> QUOTE "FIELD{n}" </w:instrText>')
        + _run('<w:fldChar w:fldCharType="separate"/>')
        + _t(f"FIELD{n}")
        + _run('<w:fldChar w:fldCharType="end"/>'),
        "field-simple": f'{lead}<w:fldSimple w:instr=" QUOTE &quot;FIELD{n}&quot; ">{_t(f"FIELD{n}")}</w:fldSimple>',
        "textbox": lead + _text_box(70000 + n, f"BOX{n}"),
        "bold-ph": _t(f"TEXT{n} ", bold) + control("ph"),
        "bold-control-text": _t(f"TEXT{n} ", bold) + control("value"),
        "control-bold-text": lead + control("bold-value"),
    }[case["shape"]]
    return f"<w:p>{inner}</w:p>"


# ---------------------------------------------------------------------------
# core-ts's readings of a case (placeholders.test.mjs at 09dc3d2)
# ---------------------------------------------------------------------------


def _items_69(xml: str) -> str:
    out = []
    for match in re.finditer(r"<w:sdt>([\s\S]*?)</w:sdt>|<w:r(?: [^>]*)?>([\s\S]*?)</w:r>", xml):
        sdt, run = match.group(1), match.group(2)
        if sdt is not None:
            content = re.search(r"<w:sdtContent>([\s\S]*?)</w:sdtContent>", sdt)
            flag = "P " if "<w:showingPlcHdr/>" in sdt else ""
            out.append(f"[{flag}{_items_69(content.group(1) if content else '')}]")
            continue
        text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", run))
        if "<w:vanish/>" in run:
            out.append(f"{{h:{text}}}")
        elif '<w:rStyle w:val="PlaceholderText"/>' in run:
            out.append("{ph}" if text == PLACEHOLDER else f"{{S:{text}}}")
        else:
            out.append(text)
    return "".join(out)


def look_69(case: dict[str, Any]) -> str:
    """A case's paragraph after check 69, as ``<n> <items> | <text>`` (core-ts ``look69``).

    Plain runs' text joined, a hidden run ``{h:...}``, a run in Placeholder
    Text ``{ph}`` (the default text) or ``{S:...}``, a control ``[...]`` with
    ``P `` where it shows its placeholder; then Office JS's paragraph text.
    """
    after = case.get("after")
    if isinstance(after, str):
        return f"{case['n']} {after}"
    paragraphs = [m.group(0) for m in re.finditer(r"<w:p(?: [^>]*)?>[\s\S]*?</w:p>", after["ooxml"])]
    paragraphs = [p for p in paragraphs if not re.search(r">(BEGIN|END)\d+[ <]", p)]
    return f"{case['n']} {_items_69(paragraphs[-1] if paragraphs else '')} | {after['paragraph']}"


def word_69(case: dict[str, Any]) -> str:
    """Word's answer for a check 69 case (core-ts ``word69``): ``NEW<n>`` after everything, outside a control; after a
    hidden run hidden, but by the End range's ``"After"``."""
    n, shape = case["n"], case["shape"]
    hidden = shape == "hidden" and case["op"] != "end-after"
    return {
        "plain": f"{n} TEXT{n} NEW{n} | TEXT{n} NEW{n}",
        "ph-word": f"{n} TEXT{n} [P {{ph}}]NEW{n} | TEXT{n} NEW{n}",
        "ph-bare": f"{n} TEXT{n} [P {{ph}}]NEW{n} | TEXT{n} NEW{n}",
        "hidden": f"{n} TEXT{n} {{h:HIDDEN{n}NEW{n}}} | TEXT{n} "
        if hidden
        else f"{n} TEXT{n} {{h:HIDDEN{n}}}NEW{n} | TEXT{n} NEW{n}",
        "control-text": f"{n} TEXT{n} [VALUE{n}]NEW{n} | TEXT{n} VALUE{n}NEW{n}",
        "ph-only": f"{n} [P {{ph}}]NEW{n} | NEW{n}",
    }[shape]


def _returned_69(case: dict[str, Any]) -> str:
    """What the write returned in Word: the text, but ``""`` where it was written hidden."""
    return "" if case["shape"] == "hidden" and case["op"] != "end-after" else f"NEW{case['n']}"


def _tidy(xml: str) -> str:
    return re.sub(r' w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"', "", xml)


def region_69(xml: str, n: int) -> str:
    """A case's region in a save's document part, revision and paragraph ids left out: BEGIN's paragraph to END's, or to
    the section properties for the last case (core-ts ``region69``)."""
    tidy = _tidy(xml)
    start = tidy.index("</w:p>", tidy.index(f">BEGIN{n} ")) + len("</w:p>")
    end = tidy.find(f">END{n}<")
    return tidy[start : tidy.index("<w:sectPr") if end < 0 else tidy.rindex("<w:p>", 0, end)]


_ITEMS_70 = re.compile(
    r'<w:sdt>([\s\S]*?)</w:sdt>|<w:(del|ins) [^>]*>([\s\S]*?)</w:\2>|<w:fldSimple [^>]*>([\s\S]*?)</w:fldSimple>'
    r'|<w:bookmarkStart [^>]*w:name="([^"]*)"[^>]*/>|<w:bookmarkEnd [^>]*/>|<w:commentRangeStart [^>]*/>'
    r"|<w:commentRangeEnd [^>]*/>|<w:r(?: [^>]*)?>([\s\S]*?)</w:r>"
)


def items_70(xml: str) -> str:
    """A paragraph's items as tokens (core-ts ``items70``).

    A bookmark ``<name>`` ... ``</b>``, a comment range ``<c>`` ... ``</c>``
    and its reference ``{cref}``, a revision ``[del:...]`` or ``[ins:...]``, a
    complex field ``{f`` instruction ``i``, ``|``, result, ``}``, a simple
    field ``{fs:...}``, a text box ``{box}``, a control ``[...]`` (``P ``
    showing its placeholder, ``{ph}`` the default text), a bold run
    ``*...*``, other run properties ``{<rPr>:...}``.
    """
    out = []
    for match in _ITEMS_70.finditer(xml):
        whole = match.group(0)
        sdt, kind, revision, simple, bookmark, run = match.groups()
        if sdt is not None:
            content = re.search(r"<w:sdtContent>([\s\S]*?)</w:sdtContent>", sdt)
            flag = "P " if "<w:showingPlcHdr/>" in sdt else ""
            out.append(f"[{flag}{items_70(content.group(1) if content else '')}]")
        elif kind is not None:
            out.append(f"[{kind}:{items_70(revision)}]")
        elif simple is not None:
            out.append(f"{{fs:{items_70(simple)}}}")
        elif bookmark is not None:
            out.append(f"<{bookmark}>")
        elif whole.startswith("<w:bookmarkEnd"):
            out.append("</b>")
        elif whole.startswith("<w:commentRange"):
            out.append("<c>" if whole.startswith("<w:commentRangeStart") else "</c>")
        else:
            field = re.search(r'<w:fldChar w:fldCharType="(begin|separate|end)"/>', run)
            if field:
                out.append({"begin": "{f", "separate": "|", "end": "}"}[field.group(1)])
            elif "<w:instrText" in run:
                out.append("i")
            elif "<w:commentReference " in run:
                out.append("{cref}")
            elif re.search(r"<w:drawing>|<mc:AlternateContent>", run):
                out.append("{box}")
            else:
                text = "".join(re.findall(r"<w:(?:t|delText)(?: [^>]*)?>([^<]*)</w:(?:t|delText)>", run))
                found = re.search(r"<w:rPr>([\s\S]*?)</w:rPr>", run)
                r_pr = found.group(1) if found else ""
                if "PlaceholderText" in r_pr:
                    out.append("{ph}" if text == PLACEHOLDER else f"{{S:{text}}}")
                elif r_pr == "<w:b/>":
                    out.append(f"*{text}*")
                else:
                    out.append(text if r_pr == "" else f"{{{r_pr}:{text}}}")
    return "".join(out)


def look_70(case: dict[str, Any]) -> str:
    """A case after check 70 (core-ts ``look70``): its paragraph (the one holding ``TEXT<n> ``, a text box's own left
    out) as tokens, and Office JS's text."""
    after = case.get("after")
    if isinstance(after, str):
        return f"{case['n']} {after}"
    xml = re.sub(r"<w:txbxContent>[\s\S]*?</w:txbxContent>", "<w:txbxContent/>", after["ooxml"])
    paragraphs = [m.group(0) for m in re.finditer(r"<w:p(?: [^>]*)?>[\s\S]*?</w:p>", xml)]
    found = next((p for p in paragraphs if f">TEXT{case['n']} " in p), "")
    return f"{case['n']} {items_70(found)} | {after['paragraph']}"


def items_70_of(case: dict[str, Any]) -> str:
    return look_70(case).split(" | ")[0]


def _items_in(n: int, xml: str) -> str:
    """:func:`items_70_of` a region's markup, its paragraph's text left empty."""
    return items_70_of({"n": n, "after": {"ooxml": xml, "paragraph": ""}})


#: The four routes Word gave ``w:noProof`` after a text box (check 70's cases 57, 58, 63 and 64).
NO_PROOF_ROUTES = ["paragraph-end", "end-before", "content-end", "cell-body-end"]


def word_70(case: dict[str, Any]) -> str:
    """Word's answer for a check 70 case (core-ts ``word70``): ``NEW<n>`` after everything the paragraph holds; after a
    text box ``w:noProof`` by four routes; a ``w:fldSimple`` read as a complex field; a deletion's text read."""
    n, shape = case["n"], case["shape"]

    def after(ending: str, text: str = "") -> str:
        return f"{n} {ending}NEW{n} | TEXT{n} {text}NEW{n}"

    if shape == "textbox" and case["op"] in NO_PROOF_ROUTES:
        return f"{n} TEXT{n} {{box}}{{<w:noProof/>:NEW{n}}} | TEXT{n} NEW{n}"
    return {
        "bookmark-span": after(f"<BM{n}>TEXT{n} </b>"),
        "bookmark-point": after(f"TEXT{n} <PT{n}></b>"),
        "comment": after(f"<c>TEXT{n} </c>{{cref}}"),
        "tracked-del": after(f"TEXT{n} [del:DELETED{n}]", f"DELETED{n}"),
        "tracked-ins": after(f"TEXT{n} [ins:INSERTED{n}]", f"INSERTED{n}"),
        "field-complex": after(f"TEXT{n} {{fi|FIELD{n}}}", f"FIELD{n}"),
        # Word reads a w:fldSimple as a complex field, and writes it so (its save on opening)
        "field-simple": after(f"TEXT{n} {{fi|FIELD{n}}}", f"FIELD{n}"),
        "textbox": after(f"TEXT{n} {{box}}"),
        "bold-ph": after(f"*TEXT{n} *[P {{ph}}]"),
        "bold-control-text": after(f"*TEXT{n} *[VALUE{n}]", f"VALUE{n}"),
        "control-bold-text": after(f"TEXT{n} [*VALUE{n}*]", f"VALUE{n}"),
    }[shape]


def ours_70(case: dict[str, Any]) -> str:
    """This package's answer for a check 70 case (core-ts's ``shapeOf``): ``NEW<n>`` after every ending, by every route,
    plain."""
    n, shape = case["n"], case["shape"]
    ending, text = {
        "bookmark-span": (f"<BM{n}>TEXT{n} </b>", ""),
        "bookmark-point": (f"TEXT{n} <PT{n}></b>", ""),
        "comment": (f"<c>TEXT{n} </c>{{cref}}", ""),
        "tracked-del": (f"TEXT{n} [del:DELETED{n}]", ""),
        "tracked-ins": (f"TEXT{n} [ins:INSERTED{n}]", f"INSERTED{n}"),
        "field-complex": (f"TEXT{n} {{fi|FIELD{n}}}", f"FIELD{n}"),
        "field-simple": (f"TEXT{n} {{fs:FIELD{n}}}", f"FIELD{n}"),
        "textbox": (f"TEXT{n} {{box}}", ""),
        "bold-ph": (f"*TEXT{n} *[P {{ph}}]", ""),
        "bold-control-text": (f"*TEXT{n} *[VALUE{n}]", f"VALUE{n}"),
        "control-bold-text": (f"TEXT{n} [*VALUE{n}*]", f"VALUE{n}"),
    }[shape]
    return f"{n} {ending}NEW{n} | TEXT{n} {text}NEW{n}"


def placed(case: dict[str, Any]) -> str:
    """Where ``NEW<n>`` went: a simple field read as Word reads it, and Word's ``w:noProof`` after a text box left out
    (D13's departure)."""
    tokens = re.sub(r"\{fs:(FIELD\d+)\}", r"{fi|\1}", items_70_of(case))
    return re.sub(r"\{<w:noProof/>:(NEW\d+)\}", r"\1", tokens)


def region_70(xml: str, n: int) -> str:
    """A case's region in a save of check 70, revision and paragraph ids left out (core-ts ``region70``)."""
    tidy = _tidy(xml)
    start = tidy.index("</w:p>", tidy.index(f">BEGIN{n} ")) + len("</w:p>")
    return tidy[start : tidy.rindex("<w:p>", 0, tidy.index(f">END{n}<"))]


# ---------------------------------------------------------------------------
# the helpers
# ---------------------------------------------------------------------------


def word(name: str) -> dict:
    return json.loads((OFFICE_JS / name).read_text(encoding="utf-8"))


def _parts(data: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name).decode("utf-8") for name in archive.namelist() if name.endswith((".xml", ".rels"))}


def _document(name: str) -> str:
    return _parts((OFFICE_JS / name).read_bytes())["word/document.xml"]


def run(check: int, mode: str = "run") -> tuple[Any, dict[str, Any]]:
    """This package's run of a check, after its report as Word's was (the snippet's Report only, then Run the check)."""
    name = {69: "69-paragraph-end", 70: "70-paragraph-endings"}[check]
    package = load(OFFICE_JS / f"{name}.docx")
    runner = getattr(checks, f"run_check_{check}")
    report = runner(package, "report")
    return package, report if mode == "report" else runner(package, "run")


def _manifest() -> dict[str, str]:
    lines = (OFFICE_JS / "SHA1SUMS").read_text(encoding="utf-8").splitlines()
    return dict(reversed(line.split("  ", 1)) for line in lines if line)


# ---------------------------------------------------------------------------
# check 69 (core-ts 09dc3d2, its sections 63.23 and 63.24)
# ---------------------------------------------------------------------------


def test_the_files_are_core_tss_as_copied_and_word_365s_jsons_are_of_these_cases_run_with_design_mode_off():
    manifest = _manifest()
    for name, sha1 in SHA1.items():
        assert manifest[name] == sha1, name
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == sha1, name
    for check, name, cases in ((69, "69-paragraph-end", cases_69()), (70, "70-paragraph-endings", cases_70())):
        for mode in ("report", "run"):
            measured = word(f"{name}-word365-{mode}.json")
            assert [measured["check"], measured["mode"], measured["host"]["version"], measured["caseCount"]] == [
                check,
                mode,
                "16.0.20326.20232",
                len(cases),
            ]
            assert [{k: c[k] for k in ("n", "op", "shape")} for c in measured["cases"]] == cases
        for save in (f"{check}-opened-word365.docx", f"{check}-run-word365.docx"):
            assert "formsDesign" not in _parts((OFFICE_JS / save).read_bytes())["word/settings.xml"], save


def test_check_69s_file_and_snippet_agree_each_paragraph_as_built_the_last_cases_the_documents_last():
    package = load(OFFICE_JS / "69-paragraph-end.docx")
    body = package.body
    cases = cases_69()
    assert len(cases) == 49
    assert [r for p in body.paragraphs if (r := checks.parse_region_69(p.text)) is not None] == cases
    # a control for each case of the four shapes that have one, Word's form naming the default entry
    controlled = [c for c in cases if c["shape"] in ("ph-word", "ph-bare", "control-text", "ph-only")]
    controls = body.content_controls
    assert [c.tag for c in controls] == [f"c69-{c['n']}" for c in controlled]
    assert [c.placeholder_entry for c in controls] == [
        DEFAULT_ENTRY if c["shape"] in ("ph-word", "ph-only") else None for c in controlled
    ]
    assert [c.is_showing_placeholder for c in controls] == [c["shape"] != "control-text" for c in controlled]
    xml = _document("69-paragraph-end.docx")
    for case in cases:
        at = xml.index(f">BEGIN{case['n']} ")
        start = xml.index("</w:p>", at) + len("</w:p>")
        end = xml.index("<w:sectPr>") if case["op"] == "body-end" else xml.rindex("<w:p>", 0, xml.index(f">END{case['n']}<"))
        region = xml[start:end]
        assert ("<w:tbl>" in region) == (case["op"] == "cell-body-end"), case["n"]
        assert paragraph_69(case) in region, case["n"]
    assert re.search(re.escape(paragraph_69(cases[48])) + r"<w:sectPr>[^<]*(<w:pg\w+ [^>]*/>)+</w:sectPr></w:body>", xml)
    assert "word/glossary/styles.xml" in zipfile.ZipFile(OFFICE_JS / "69-paragraph-end.docx").namelist()


def test_word_365s_answers_check_69_a_write_at_a_paragraphs_end_goes_after_everything_it_holds_hidden_after_a_hidden_run():
    assert not [c for c in word("69-paragraph-end-word365-report.json")["cases"] if c.get("error")]
    measured = word("69-paragraph-end-word365-run.json")
    assert not [c for c in measured["cases"] if c.get("error")], "Office JS refused nothing"
    # every route alike, but the End range's "After" after a hidden run (case 27)
    assert [look_69(c) for c in measured["cases"]] == [word_69(c) for c in cases_69()]
    assert [c["result"]["returned"] for c in measured["cases"]] == [_returned_69(c) for c in cases_69()]
    # Word's save: NEW<n> a run of its own, with no properties but a hidden run's w:vanish; each control as it was
    saved = _document("69-run-word365.docx")
    kept = (
        r"<w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:r><w:rPr><w:rStyle w:val=\"PlaceholderText\"/>"
        r"</w:rPr><w:t>Click or tap here to enter text\.</w:t></w:r></w:sdtContent></w:sdt><w:r>"
    )
    for case in measured["cases"]:
        region = region_69(saved, case["n"])
        vanish = "<w:rPr><w:vanish/></w:rPr>" if _returned_69(case) == "" else ""
        assert re.search(rf"<w:r>{re.escape(vanish)}<w:t>NEW{case['n']}</w:t></w:r></w:p>", region), case["n"]
        if case["shape"].startswith("ph"):
            assert re.search(kept, region), case["n"]
    # Word's save on opening holds every paragraph as built
    opened = _document("69-opened-word365.docx")
    for case in cases_69():
        built = {"n": case["n"], "after": {"ooxml": paragraph_69(case), "paragraph": ""}}
        assert look_69({"n": case["n"], "after": {"ooxml": region_69(opened, case["n"]), "paragraph": ""}}) == look_69(built)


def test_check_69s_report_through_the_port_reads_every_region_as_word_365_does():
    _package, ours = run(69, "report")
    measured = word("69-paragraph-end-word365-report.json")
    assert ours["caseCount"] == 49 and ours["title"] == measured["title"]
    assert [look_69(c) for c in ours["cases"]] == [look_69(c) for c in measured["cases"]]
    # a placeholder and a hidden run read as nothing, a control's value as its text
    assert [look_69(c) for c in ours["cases"] if c["op"] == "paragraph-end"] == [
        "1 TEXT1  | TEXT1 ",
        "9 TEXT9 [P {ph}] | TEXT9 ",
        "17 TEXT17 [P {ph}] | TEXT17 ",
        "25 TEXT25 {h:HIDDEN25} | TEXT25 ",
        "33 TEXT33 [VALUE33] | TEXT33 VALUE33",
        "41 [P {ph}] | ",
    ]


def test_check_69s_run_through_the_port_is_word_365s_in_every_case():
    _package, ours = run(69)
    measured = word("69-paragraph-end-word365-run.json")
    assert not [c for c in ours["cases"] if c.get("error")]
    # at 81cbe45 (CR-019 Phase C) 40 of the 49 differed: the text went inside a trailing placeholder control, behind
    # the prompt with the flag on (9, 11 to 17, 19 to 24, 41 and 43 to 49: read as nothing), or before it (the End
    # range's "Before", 10, 18 and 42); after a hidden run plain (25, 26, 28 to 32), before it by "Before"; inside a
    # trailing control holding text (33 to 40). Only the plain paragraphs and case 27 were Word's
    assert [look_69(c) for c in ours["cases"]] == [look_69(c) for c in measured["cases"]]
    assert [c["result"]["returned"] for c in ours["cases"]] == [c["result"]["returned"] for c in measured["cases"]]
    for theirs, mine in zip(measured["cases"], ours["cases"], strict=True):
        control = theirs["after"]["control"]
        assert (mine["after"]["control"] is None) == (control is None), theirs["n"]
        if control is not None:
            assert {k: mine["after"]["control"][k] for k in ("type", "text")} == {k: control[k] for k in ("type", "text")}
            # each control as it was: one showing its placeholder keeps it (Word gives no showingPlaceholderText)
            assert mine["after"]["showingPlaceholderText"] == (theirs["shape"] != "control-text"), theirs["n"]


#: The regions of this package's save after check 69's run that differ from Word's save after the same run, by
#: core-ts's items, each with its reason: after a hidden run Word's save writes NEW<n> as a run of its own in the hidden
#: run's look, where this package extends the hidden run's last w:t (D13), as Word's own run read it back through
#: getOoxml (its JSON's `{h:HIDDEN<n>NEW<n>}`): a run boundary, the same text in the same look.
SAVE_DIFFERS_69 = {n: "a run boundary: Word's save splits the hidden run" for n in (25, 26, 28, 29, 30, 31, 32)}


def test_check_69s_saved_regions_are_words_save_after_the_run_but_for_the_hidden_runs_boundary():
    package, _ours = run(69)
    saved = _parts(package.save())["word/document.xml"]
    words = _document("69-run-word365.docx")
    differs = {}
    for case in cases_69():
        n = case["n"]
        mine, theirs = _items_69(region_69(saved, n)), _items_69(region_69(words, n))
        if mine != theirs:
            differs[n] = (mine, theirs)
    assert sorted(differs) == sorted(SAVE_DIFFERS_69)
    for n, (mine, theirs) in differs.items():
        assert (mine, theirs) == (f"TEXT{n} {{h:HIDDEN{n}NEW{n}}}", f"TEXT{n} {{h:HIDDEN{n}}}{{h:NEW{n}}}"), n
    # the saved markup (core-ts's): the last run extended, but by "After", which writes a run of its own with no
    # properties (D14); hidden after a hidden run; else NEW<n> a run of its own after the control, its flag kept
    for case in cases_69():
        n, region = case["n"], region_69(saved, case["n"])
        hidden = case["shape"] == "hidden" and case["op"] != "end-after"
        if case["shape"] == "plain" and case["op"] != "end-after":
            assert re.search(rf'<w:r><w:t xml:space="preserve">TEXT{n} NEW{n}</w:t></w:r></w:p>', region), n
        elif hidden:
            assert re.search(rf'<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">HIDDEN{n}NEW{n}</w:t></w:r></w:p>', region), n
        else:
            assert re.search(rf"</w:(sdt|r)><w:r><w:t>NEW{n}</w:t></w:r></w:p>", region), n


# ---------------------------------------------------------------------------
# check 70 (core-ts 09dc3d2, its sections 63.25 and 63.26)
# ---------------------------------------------------------------------------


def test_check_70s_file_and_snippet_agree_each_paragraphs_ending_as_built():
    package = load(OFFICE_JS / "70-paragraph-endings.docx")
    body = package.body
    cases = cases_70()
    assert len(cases) == 88
    assert [r for p in body.paragraphs if (r := checks.parse_region_69(p.text)) is not None] == cases
    # the controls of the D13 shapes, the placeholder in Word's form
    controlled = [c for c in cases if c["shape"] in ("bold-ph", "bold-control-text", "control-bold-text")]
    assert [(c.tag, c.is_showing_placeholder, c.placeholder_entry) for c in body.content_controls] == [
        (f"c70-{c['n']}", c["shape"] == "bold-ph", DEFAULT_ENTRY if c["shape"] == "bold-ph" else None) for c in controlled
    ]
    xml = _document("70-paragraph-endings.docx")
    for case in cases:
        start = xml.index("</w:p>", xml.index(f">BEGIN{case['n']} ")) + len("</w:p>")
        region = xml[start : xml.rindex("<w:p>", 0, xml.index(f">END{case['n']}<"))]
        assert ("<w:tbl>" in region) == (case["op"] == "cell-body-end"), case["n"]
        assert paragraph_70(case) in region, case["n"]
    endings = [re.sub(r"\d+", "#", _items_in(c["n"], paragraph_70(c))) for c in cases]
    assert list(dict.fromkeys(endings)) == [
        "# <BM#>TEXT# </b>",
        "# TEXT# <PT#></b>",
        "# <c>TEXT# </c>{cref}",
        "# TEXT# [del:DELETED#]",
        "# TEXT# [ins:INSERTED#]",
        "# TEXT# {fi|FIELD#}",
        "# TEXT# {fs:FIELD#}",
        "# TEXT# {box}",
        "# *TEXT# *[P {ph}]",
        "# *TEXT# *[VALUE#]",
        "# TEXT# [*VALUE#*]",
    ]
    # a comment for each comment case, and the revisions in the tracked cases, as the listing reads them
    comments = _parts((OFFICE_JS / "70-paragraph-endings.docx").read_bytes())["word/comments.xml"]
    assert [int(m) for m in re.findall(r'<w:comment w:id="(\d+)"', comments)] == [
        c["n"] for c in cases if c["shape"] == "comment"
    ]
    changes = [(change.type, change.text, change.author) for change in body.get_tracked_changes()]
    assert changes == [
        ("Deleted", f"DELETED{c['n']}", AUTHOR_70) if c["shape"] == "tracked-del" else ("Added", f"INSERTED{c['n']}", AUTHOR_70)
        for c in cases
        if c["shape"].startswith("tracked")
    ]


def test_word_365s_answers_check_70_after_everything_into_no_bookmark_comment_revision_field_or_text_box_plain_after_a_control():
    measured = word("70-paragraph-endings-word365-run.json")
    assert not [c for c in measured["cases"] if c.get("error")], "Office JS refused nothing"
    assert [look_70(c) for c in measured["cases"]] == [word_70(c) for c in cases_70()]
    assert [c["result"]["returned"] for c in measured["cases"]] == [f"NEW{c['n']}" for c in cases_70()]
    # Word's save holds the same; NEW<n> a run of its own, plain but after a text box (w:noProof, by four routes): not
    # the bold of the run before a control (the user: "NEW65 is not bold; NEW73 is not bold"), nor a comment
    # reference's style
    saved = _document("70-run-word365.docx")
    for case in measured["cases"]:
        region = region_70(saved, case["n"])
        assert _items_in(case["n"], region) == items_70_of(case), case["n"]
        proof = "<w:rPr><w:noProof/></w:rPr>" if "{<w:noProof/>:NEW" in look_70(case) else ""
        assert re.search(rf"<w:r>{re.escape(proof)}<w:t>NEW{case['n']}</w:t></w:r></w:p>", region), case["n"]
    # Word's save on opening holds every paragraph as built, but a w:fldSimple written as a complex field
    opened = _document("70-opened-word365.docx")
    for case in cases_70():
        built = re.sub(r"\{fs:(FIELD\d+)\}", r"{fi|\1}", _items_in(case["n"], paragraph_70(case)))
        assert _items_in(case["n"], region_70(opened, case["n"])) == built, case["n"]


#: The cases of check 70's report where this package reads otherwise than Word, each with its reason, as core-ts pins
#: them: Word's All Markup reads a deletion's text, which Paragraph.text reads as Word's default view does (check 53);
#: Word reads a w:fldSimple as a complex field, which this package keeps as written.
REPORT_DIFFERS_70 = {
    **{n: "All Markup reads the deletion's text" for n in range(25, 33)},
    **{n: "Word reads the w:fldSimple as a complex field" for n in range(49, 57)},
}


def test_check_70s_report_through_the_port_reads_every_region_as_word_365_does_but_a_deletion_and_a_simple_field():
    _package, ours = run(70, "report")
    measured = word("70-paragraph-endings-word365-report.json")
    assert ours["caseCount"] == 88 and ours["title"] == measured["title"]
    differs = [t["n"] for t, m in zip(measured["cases"], ours["cases"], strict=True) if look_70(t) != look_70(m)]
    assert differs == sorted(REPORT_DIFFERS_70)
    for theirs, mine in zip(measured["cases"], ours["cases"], strict=True):
        if theirs["shape"] == "tracked-del":
            n = theirs["n"]
            assert (look_70(theirs), look_70(mine)) == (
                f"{n} TEXT{n} [del:DELETED{n}] | TEXT{n} DELETED{n}",
                f"{n} TEXT{n} [del:DELETED{n}] | TEXT{n} ",
            )


def test_check_70s_run_through_the_port_is_written_where_word_365_wrote_in_every_case():
    package, ours = run(70)
    measured = word("70-paragraph-endings-word365-run.json")
    assert not [c for c in ours["cases"] if c.get("error")]
    # after every ending, by every route, plain (D12, D13); at 81cbe45 (CR-019 Phase C) every one of the 88 differed:
    # the text went inside a spanning bookmark, a comment's range, a field's result (both forms), another author's
    # tracked insertion, and a control holding text; before a point bookmark, a deletion and a text box; and after a
    # bold run before a placeholder control into the control, bold, behind the prompt (65, 67 to 72), or before it by
    # the End range's "Before" (66)
    assert [look_70(c) for c in ours["cases"]] == [ours_70(c) for c in cases_70()]
    assert [c["result"]["returned"] for c in ours["cases"]] == [f"NEW{c['n']}" for c in cases_70()]
    # where NEW<n> went is Word's in all 88
    assert [t["n"] for t, m in zip(measured["cases"], ours["cases"], strict=True) if placed(t) != placed(m)] == []
    # what is left, none of it placement: Word's All Markup reads a deletion's text (check 53: Paragraph.text follows
    # Word's default view), a w:fldSimple Word reads as a complex field, and the four routes giving w:noProof after a
    # text box (D13's departure: it only stops proofing)
    differ = [t for t, m in zip(measured["cases"], ours["cases"], strict=True) if look_70(t) != look_70(m)]
    assert list(dict.fromkeys(c["shape"] for c in differ)) == ["tracked-del", "field-simple", "textbox"]
    assert [c["op"] for c in differ if c["shape"] == "textbox"] == NO_PROOF_ROUTES
    # the tracked insertion stays as it was, and the new text is outside it, untracked
    assert [c for c in package.body.get_tracked_changes() if "NEW" in c.text] == []


#: The regions of this package's save after check 70's run that differ from Word's save after the same run, each
#: with its reason: Word's save writes a w:fldSimple as a complex field (as its save on opening does), and Word gave
#: NEW<n> after a text box w:noProof by four routes, which D13 leaves out.
SAVE_DIFFERS_70 = {
    **{n: "Word writes the w:fldSimple as a complex field" for n in range(49, 57)},
    **{n: "Word's w:noProof after the text box (D13's departure)" for n in (57, 58, 63, 64)},
}


def test_check_70s_saved_regions_are_words_save_after_the_run_but_a_simple_fields_form_and_w_no_proof():
    package, _ours = run(70)
    saved = _parts(package.save())["word/document.xml"]
    words = _document("70-run-word365.docx")
    differs = {}
    for case in cases_70():
        n = case["n"]
        mine, theirs = _items_in(n, region_70(saved, n)), _items_in(n, region_70(words, n))
        if mine != theirs:
            differs[n] = (mine, theirs)
    assert sorted(differs) == sorted(SAVE_DIFFERS_70)
    for n, (mine, theirs) in differs.items():
        if n < 57:
            assert (mine, theirs) == (f"{n} TEXT{n} {{fs:FIELD{n}}}NEW{n}", f"{n} TEXT{n} {{fi|FIELD{n}}}NEW{n}"), n
        else:
            assert (mine, theirs) == (f"{n} TEXT{n} {{box}}NEW{n}", f"{n} TEXT{n} {{box}}{{<w:noProof/>:NEW{n}}}"), n
    # NEW<n> a run of its own after everything, with no properties, in every region
    for case in cases_70():
        assert re.search(rf"<w:r><w:t>NEW{case['n']}</w:t></w:r></w:p>", region_70(saved, case["n"])), case["n"]


# ---------------------------------------------------------------------------
# core-ts's tests that are no replay (placeholders.test.mjs at 09dc3d2)
# ---------------------------------------------------------------------------


def with_body(body: str) -> Any:
    """A created package whose body is `body` (core-ts ``withBody``)."""
    package = create_package()
    package.body.insert_xml(body, location="End")
    return package


def first_paragraph_xml(package: Any) -> str:
    """The body's first paragraph as saved, ``xml:space``, ids, dates and namespace declarations left out."""
    xml = to_xml(package.main_document_part.contents.body.content[0])
    return re.sub(r' xmlns:\w+="[^"]*"| xml:space="preserve"| w:(?:date|id|author)="[^"]*"| w16du:dateUtc="[^"]*"', "", xml)


_PLACEHOLDER_CONTROL = (
    '<w:sdt><w:sdtPr><w:id w:val="9"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:rPr>'
    f'<w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>{PLACEHOLDER}</w:t></w:r></w:sdtContent></w:sdt>'
)
_VALUE_CONTROL = '<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:r><w:t>VALUE</w:t></w:r></w:sdtContent></w:sdt>'
_TEXT = '<w:r><w:t xml:space="preserve">TEXT </w:t></w:r>'
_HIDDEN = "<w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r>"


def test_a_write_at_the_end_past_a_link_goes_after_it_plain_a_proofing_mark_is_no_item_a_paragraph_of_markers_gets_a_run():
    link = with_body(
        '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="x"><w:r><w:rPr>'
        '<w:rStyle w:val="Hyperlink"/></w:rPr><w:t>here</w:t></w:r></w:hyperlink></w:p>'
    )
    link.body.paragraphs[0].insert_text(".", location="End")
    assert first_paragraph_xml(link).endswith("</w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>")
    # a proofing mark is no item, so the misspelt bold word before it lends its bold
    proof = with_body(
        '<w:p><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t>Teh</w:t></w:r>'
        '<w:proofErr w:type="spellEnd"/></w:p>'
    )
    proof.body.paragraphs[0].insert_text(" cat", location="End")
    assert first_paragraph_xml(proof).endswith(
        '<w:r><w:rPr><w:b/></w:rPr><w:t>Teh cat</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>'
    )
    # a paragraph with no runs: a run of its own, after a bookmark it holds
    marks = with_body('<w:p><w:bookmarkStart w:id="1" w:name="B"/><w:bookmarkEnd w:id="1"/></w:p>')
    assert marks.body.paragraphs[0].insert_text("X", location="End").text == "X"
    assert first_paragraph_xml(marks).endswith("<w:bookmarkEnd/><w:r><w:t>X</w:t></w:r></w:p>")


def test_tracked_a_write_at_the_end_is_a_w_ins_after_everything_or_joins_this_authors_own_insertion_that_ends_it():
    package = with_body(f"<w:p>{_TEXT}{_PLACEHOLDER_CONTROL}</w:p>")
    package.change_tracking_mode = "TrackAll"
    paragraph = package.body.paragraphs[0]
    assert paragraph.insert_text("NEW", location="End").text == "NEW"
    assert re.search(r"</w:sdt><w:ins(?: [^>]*)?><w:r><w:t>NEW</w:t></w:r></w:ins></w:p>$", first_paragraph_xml(package))
    # the same author again: the insertion grows, in its run's look
    paragraph.get_range("End").insert_text("ER", location="Before")
    assert re.search(r"</w:sdt><w:ins(?: [^>]*)?><w:r><w:t>NEWER</w:t></w:r></w:ins></w:p>$", first_paragraph_xml(package))
    assert [(c.type, c.text) for c in package.body.get_tracked_changes()] == [("Added", "NEWER")]
    # the control kept its placeholder, its flag on
    control = package.body.content_controls[0]
    assert (control.is_showing_placeholder, control.text, paragraph.text) == (True, "", "TEXT NEWER")


def test_a_write_at_the_end_holding_a_line_break_makes_paragraphs_after_everything_after_a_hidden_run_visible():
    package = with_body(f"<w:p>{_TEXT}{_VALUE_CONTROL}</w:p>")
    returned = package.body.paragraphs[0].insert_text("A\nB", location="End")
    assert [p.text for p in package.body.paragraphs] == ["TEXT VALUEA", "B"]
    assert returned.text == "A\rB"
    assert package.body.content_controls[0].text == "VALUE"
    hidden = with_body(f"<w:p>{_TEXT}{_HIDDEN}</w:p>")
    hidden.body.paragraphs[0].insert_text("A\nB", location="End")
    assert [p.text for p in hidden.body.paragraphs] == ["TEXT A", "B"]
    # without a break, after the hidden run, hidden (D13): the range reads nothing, as Word's
    one = with_body(f"<w:p>{_TEXT}{_HIDDEN}</w:p>")
    assert one.body.paragraphs[0].insert_text("A", location="End").text == ""
    assert one.body.paragraphs[0].text == "TEXT "


def test_a_search_hit_ending_at_a_paragraphs_end_is_no_end_route_text_after_it_goes_inside_a_trailing_controls_text():
    package = with_body(f"<w:p>{_TEXT}{_VALUE_CONTROL}</w:p>")
    [hit] = package.body.search("VALUE")
    hit.insert_text("!", location="After")
    assert package.body.content_controls[0].text == "VALUE!"


# ---------------------------------------------------------------------------
# this package's own
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "route",
    [
        "paragraph.insert_text",
        "end.Before",
        "end.After",
        "end.Start",
        "end.End",
        "end.Replace",
        "content.End",
        "content.After",
        "body.insert_text",
        "body.get_range(End)",
        "cell.insert_text",
    ],
)
@pytest.mark.parametrize("tracked", [False, True])
def test_every_route_to_a_paragraphs_end_writes_after_a_trailing_placeholder_control_which_keeps_its_placeholder(route, tracked):
    table = route.startswith("cell")
    paragraph_xml = f"<w:p>{_TEXT}{_PLACEHOLDER_CONTROL}</w:p>"
    package = with_body(
        f"<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>{paragraph_xml}</w:tc></w:tr></w:tbl>"
        if table
        else paragraph_xml
    )
    if tracked:
        package.change_tracking_mode = "TrackAll"
    body = package.body
    paragraph = body.paragraphs[0]
    owner, _, where = route.partition(".")
    if route == "paragraph.insert_text":
        returned = paragraph.insert_text("NEW", location="End")
    elif owner in ("end", "content"):
        returned = paragraph.get_range(owner.capitalize()).insert_text("NEW", location=where)
    elif route == "body.insert_text":
        returned = body.insert_text("NEW", location="End")
    elif route == "body.get_range(End)":
        returned = body.get_range("End").insert_text("NEW", location="Before")
    else:
        returned = body.tables[0].cell(0, 0).insert_text("NEW", location="End")
    paragraph = body.paragraphs[0]
    control = body.content_controls[0]
    assert (returned.text, paragraph.text, control.text, control.is_showing_placeholder) == ("NEW", "TEXT NEW", "", True)
    xml = to_xml(paragraph.element)
    assert re.search(r"</w:sdt>(<w:ins [^>]*>)?<w:r><w:t>NEW</w:t></w:r>(</w:ins>)?</w:p>$", re.sub(r" xmlns:\w+=\"[^\"]*\"", "", xml))
    assert ("<w:ins " in xml) == tracked


def test_the_end_range_follows_the_text_it_put_before_it_and_covers_what_it_replaced_with():
    package = with_body(f"<w:p>{_TEXT}{_PLACEHOLDER_CONTROL}</w:p>")
    paragraph = package.body.paragraphs[0]
    end = paragraph.get_range("End")
    end.insert_text("A", location="Before")
    # still collapsed at the end, so still the paragraph's End
    assert (end.start, end.end, end.text) == (6, 6, "")
    end.insert_text("B", location="Start")
    assert (end.start, end.end, paragraph.text) == (7, 7, "TEXT AB")
    end.insert_text("C", location="Replace")
    # it covers its text now, so it is no longer collapsed: its writes are an ordinary range's
    assert (end.start, end.end, end.text) == (7, 8, "C")
    content = paragraph.get_range("Content")
    content.insert_text("D", location="After")
    assert (content.text, paragraph.text) == ("TEXT ABC", "TEXT ABCD")
    # a range made from one is an ordinary range, and so is one no longer ending at its paragraph's end
    assert paragraph.get_range("Content").get_range("End").end_of is None
    assert paragraph.get_range("Whole").end_of is None and paragraph.get_range("Start").end_of is None


def test_the_content_ranges_after_writes_plain_after_a_hidden_run_its_end_hidden_d14_as_core_ts_takes_it():
    package = with_body(f"<w:p>{_TEXT}{_HIDDEN}</w:p>")
    paragraph = package.body.paragraphs[0]
    assert paragraph.get_range("Content").insert_text("A", location="After").text == "A"
    assert paragraph.text == "TEXT A"
    # now the last item is the plain run "A": the End lends its look, nothing hidden
    assert paragraph.get_range("Content").insert_text("B", location="End").text == "B"
    assert paragraph.text == "TEXT AB"
    xml = first_paragraph_xml(package)
    assert xml.endswith("<w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r><w:r><w:t>AB</w:t></w:r></w:p>")


def test_a_bound_block_controls_node_takes_a_write_at_its_paragraphs_end_a_trailing_bound_run_level_controls_does_not():
    package = create_package()
    part = package.custom_xml_parts.add('<g xmlns="http://example.com/g"><block/><run>VAL</run></g>')
    binding = (
        '<w:dataBinding w:prefixMappings="xmlns:ns0=\'http://example.com/g\'" w:xpath="/ns0:g[1]/ns0:{0}[1]" '
        'w:storeItemID="{{11111111-2222-4333-8444-555555555555}}"/><w:text/>'
    )
    package.body.insert_xml(
        f'<w:sdt><w:sdtPr><w:id w:val="1"/>{binding.format("block")}</w:sdtPr><w:sdtContent>'
        '<w:p><w:r><w:t>BLOCK</w:t></w:r></w:p></w:sdtContent></w:sdt>'
        f'<w:p><w:r><w:t xml:space="preserve">To: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="2"/>{binding.format("run")}'
        "</w:sdtPr><w:sdtContent><w:r><w:t>VAL</w:t></w:r></w:sdtContent></w:sdt></w:p>",
        location="End",
    )
    block, inline = package.body.get_content_controls()
    for control, name in ((block, "block"), (inline, "run")):
        control.xml_mapping.set_mapping(f"/ns0:g[1]/ns0:{name}[1]", "xmlns:ns0='http://example.com/g'", part)
    block.paragraphs[0].insert_text("ED", location="End")
    assert part.select_single_node("/ns0:g[1]/ns0:block[1]").text == "BLOCKED"
    package.body.paragraphs[1].insert_text("UE", location="End")
    # after the control, outside it: the control and its node as they were
    assert (inline.text, package.body.paragraphs[1].text) == ("VAL", "To: VALUE")
    assert part.select_single_node("/ns0:g[1]/ns0:run[1]").text == "VAL"


def test_a_break_at_the_end_of_a_paragraph_ending_in_a_fields_instruction_is_refused_before_anything_changes():
    package = with_body(
        '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> QUOTE "x" </w:instrText></w:r></w:p>'
    )
    paragraph = package.body.paragraphs[0]
    before = to_xml(paragraph.element)
    with pytest.raises(Exception) as raised:
        paragraph.insert_text("B\nC", location="End")
    assert getattr(raised.value, "code", None) == "range.splits_field"
    assert to_xml(paragraph.element) == before and len(package.body.paragraphs) == 1
