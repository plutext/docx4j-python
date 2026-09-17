#!/usr/bin/env python
"""Write ``tests/python_docx_subset.json`` from an installed python-docx.

CR-003 section 3.6 promises one committed list of the python-docx members
:mod:`docx4j_py.docx` implements, and a test that keeps the facade equal to it
(``tests/content/test_python_docx_subset.py``). The list is not written here: it
is **derived** by introspecting python-docx itself, at the pinned version
:data:`PINNED`, so that the promise is against python-docx's real API and not
against someone's memory of it.

python-docx is **not a dependency of this package**. Install it in a throwaway
virtual environment and run this script with that interpreter:

    uv venv /tmp/pydocx && VIRTUAL_ENV=/tmp/pydocx uv pip install python-docx==1.2.0
    /tmp/pydocx/bin/python scripts/python_docx_subset.py            # rewrite the JSON
    /tmp/pydocx/bin/python scripts/python_docx_subset.py --check     # fail if it would change

When python-docx is not importable the script **says so and changes nothing**:
the JSON is committed, so the test runs in ``.venv-fork`` like everything else.
An optional ``parity`` extra in ``pyproject.toml`` names the same version, for
the parity test of section 20.

**What the list contains.** For each class of :data:`INTERFACES`, every public
member found on its MRO below ``object``: its name, its kind (``property``,
``method`` or ``attribute``), whether it is read-only, and the class in the MRO
that declares it. Each member is tagged ``implemented: true`` unless
:data:`OMITTED` names it, which is the hand-kept list of deliberate omissions
with a one-line reason each. Beside the interfaces the file carries the
``Length`` constructors and their EMU factors, and every ``WD_*`` / ``MSO_*``
enumeration with its members and their ``w:val`` strings, so that the test can
assert the values and not just the names.
"""

from __future__ import annotations

import argparse
import importlib
import inspect
import json
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent

#: Where the committed list lives.
TARGET = ROOT / "tests" / "python_docx_subset.json"

#: The python-docx version the promise is against.
PINNED = "1.2.0"

#: ``module:ClassName`` for every python-docx class the facade offers, and the
#: name :mod:`docx4j_py.docx` exports it under.
INTERFACES: dict[str, str] = {
    "Document": "docx.document:Document",
    "Paragraph": "docx.text.paragraph:Paragraph",
    "Run": "docx.text.run:Run",
    "Font": "docx.text.font:Font",
    "ParagraphFormat": "docx.text.parfmt:ParagraphFormat",
    "ColorFormat": "docx.dml.color:ColorFormat",
    "Hyperlink": "docx.text.hyperlink:Hyperlink",
    "RenderedPageBreak": "docx.text.pagebreak:RenderedPageBreak",
    "Table": "docx.table:Table",
    "_Row": "docx.table:_Row",
    "_Rows": "docx.table:_Rows",
    "_Column": "docx.table:_Column",
    "_Columns": "docx.table:_Columns",
    "_Cell": "docx.table:_Cell",
    "Section": "docx.section:Section",
    "Sections": "docx.section:Sections",
    "_Header": "docx.section:_Header",
    "_Footer": "docx.section:_Footer",
    "InlineShape": "docx.shape:InlineShape",
    "InlineShapes": "docx.shape:InlineShapes",
    "Styles": "docx.styles.styles:Styles",
    "BaseStyle": "docx.styles.style:BaseStyle",
    "CharacterStyle": "docx.styles.style:CharacterStyle",
    "ParagraphStyle": "docx.styles.style:ParagraphStyle",
    "_TableStyle": "docx.styles.style:_TableStyle",
    "_NumberingStyle": "docx.styles.style:_NumberingStyle",
    "CoreProperties": "docx.opc.coreprops:CoreProperties",
    "Settings": "docx.settings:Settings",
    "Comment": "docx.comments:Comment",
    "Comments": "docx.comments:Comments",
    "Length": "docx.shared:Length",
    "RGBColor": "docx.shared:RGBColor",
    "Drawing": "docx.drawing:Drawing",
}

#: ``Interface.member`` -> why the facade does not offer it. The whole list of
#: deliberate omissions; anything not here must be on the facade class.
OMITTED: dict[str, str] = {
    "_Cell.merge": (
        "python-docx's rectangular merge is about 150 lines of grid surgery whose "
        "failure mode in Word is a repair prompt; Phase J left it out rather than "
        "half-implement it (CR-003 section 20)"
    ),
    "Styles.latent_styles": (
        "371 w:lsdException entries behind an API nothing in the wild uses; the typed "
        "w:latentStyles is document.styles.element.latent_styles"
    ),
    "ParagraphFormat.tab_stops": (
        "the TabStops collection is a class of its own with its own add/clear verbs; "
        "w:tabs is document-order typed content on paragraph_format.element.p_pr.tabs"
    ),
    "Comments.add_comment": (
        "an unanchored comment, which Word shows nothing for; Phase G has no verb for "
        "one and Document.add_comment(runs, text) is the anchored form"
    ),
    "Run.mark_comment_range": (
        "python-docx's internal helper for add_comment; Phase G places the three "
        "markers itself, hoisting them out of a w:ins (CR-003 section 15.2)"
    ),
    "RenderedPageBreak.preceding_paragraph_fragment": (
        "splits a paragraph at a rendered page break into two detached Paragraphs; "
        "nothing in this engine renders, so nothing writes the break it needs"
    ),
    "RenderedPageBreak.following_paragraph_fragment": (
        "the other half of the same split; see preceding_paragraph_fragment"
    ),
    "Drawing.image": (
        "python-docx's Image proxy is a whole image-format parser (blob, dpi, "
        "content_type, native size); this engine reads the header it needs in "
        "docx4j_py.wml.pictures and hands the bytes out through the content API's "
        "InlinePicture.get_bytes() and .image_part"
    ),
}

#: The ``Length`` constructors, and the EMU each unit is worth.
CONSTRUCTORS: dict[str, int] = {
    "Emu": 1,
    "Twips": 635,
    "Pt": 12700,
    "Mm": 36000,
    "Cm": 360000,
    "Inches": 914400,
}

#: ``module:EnumName`` for every enumeration the facade offers.
ENUMS: dict[str, str] = {
    "WD_PARAGRAPH_ALIGNMENT": "docx.enum.text:WD_PARAGRAPH_ALIGNMENT",
    "WD_BREAK_TYPE": "docx.enum.text:WD_BREAK_TYPE",
    "WD_COLOR_INDEX": "docx.enum.text:WD_COLOR_INDEX",
    "WD_LINE_SPACING": "docx.enum.text:WD_LINE_SPACING",
    "WD_TAB_ALIGNMENT": "docx.enum.text:WD_TAB_ALIGNMENT",
    "WD_TAB_LEADER": "docx.enum.text:WD_TAB_LEADER",
    "WD_UNDERLINE": "docx.enum.text:WD_UNDERLINE",
    "WD_STYLE_TYPE": "docx.enum.style:WD_STYLE_TYPE",
    "WD_CELL_VERTICAL_ALIGNMENT": "docx.enum.table:WD_CELL_VERTICAL_ALIGNMENT",
    "WD_ROW_HEIGHT_RULE": "docx.enum.table:WD_ROW_HEIGHT_RULE",
    "WD_TABLE_ALIGNMENT": "docx.enum.table:WD_TABLE_ALIGNMENT",
    "WD_TABLE_DIRECTION": "docx.enum.table:WD_TABLE_DIRECTION",
    "WD_HEADER_FOOTER_INDEX": "docx.enum.section:WD_HEADER_FOOTER_INDEX",
    "WD_ORIENTATION": "docx.enum.section:WD_ORIENTATION",
    "WD_SECTION_START": "docx.enum.section:WD_SECTION_START",
    "WD_INLINE_SHAPE_TYPE": "docx.enum.shape:WD_INLINE_SHAPE_TYPE",
    "MSO_COLOR_TYPE": "docx.enum.dml:MSO_COLOR_TYPE",
    "MSO_THEME_COLOR_INDEX": "docx.enum.dml:MSO_THEME_COLOR_INDEX",
}

#: Enumerations the facade deliberately does not offer, and why.
OMITTED_ENUMS: dict[str, str] = {
    "WD_BUILTIN_STYLE": (
        "163 members naming Word's built-in styles by an MS API integer; "
        "docx4j_py.model.content.styles.BUILT_IN_STYLES is the same list keyed by "
        "style id and display name, which is what a style setter takes"
    ),
}


def load(path: str) -> Any:
    """``module:Name`` -> the object."""
    module_name, _, attribute = path.partition(":")
    return getattr(importlib.import_module(module_name), attribute)


def kind_of(member: Any) -> str:
    """``"property"``, ``"method"`` or ``"attribute"`` for a class attribute.

    python-docx's ``lazyproperty`` and ``write_only_property`` are properties as
    far as a caller is concerned, and that is what the promise is about.
    """
    if isinstance(member, property):
        return "property"
    if type(member).__name__ in ("lazyproperty", "member_descriptor", "cached_property"):
        return "property"
    if isinstance(member, (staticmethod, classmethod)):
        return "method"
    if inspect.isfunction(member) or inspect.ismethoddescriptor(member):
        return "method"
    return "attribute"


def read_only(member: Any) -> bool:
    """True when the member cannot be assigned to."""
    if isinstance(member, property):
        return member.fset is None
    if type(member).__name__ == "lazyproperty":
        return True
    return kind_of(member) == "method"


def members_of(cls: type) -> list[dict[str, Any]]:
    """Every public member on a class's MRO below ``object``, in name order."""
    found: dict[str, dict[str, Any]] = {}
    for klass in cls.__mro__:
        if klass is object or klass.__module__.startswith(("builtins", "typing")):
            continue
        for name, member in vars(klass).items():
            if name.startswith("_") or name in found:
                continue
            found[name] = {
                "name": name,
                "kind": kind_of(member),
                "read_only": read_only(member),
                "declared_by": klass.__name__,
            }
    return [found[name] for name in sorted(found)]


def build() -> dict[str, Any]:
    """The whole JSON document, from the installed python-docx."""
    interfaces: dict[str, list[dict[str, Any]]] = {}
    for name, path in INTERFACES.items():
        entries = members_of(load(path))
        for entry in entries:
            reason = OMITTED.get(f"{name}.{entry['name']}")
            entry["implemented"] = reason is None
            if reason is not None:
                entry["reason"] = reason
        interfaces[name] = entries

    enums: dict[str, dict[str, Any]] = {}
    for name, path in ENUMS.items():
        enum_cls = load(path)
        enums[name] = {
            "members": {
                member.name: {
                    "value": int(member.value),
                    "xml_value": getattr(member, "xml_value", None),
                }
                for member in enum_cls
            }
        }

    total = sum(len(entries) for entries in interfaces.values())
    implemented = sum(
        1 for entries in interfaces.values() for entry in entries if entry["implemented"]
    )
    return {
        "_": (
            "The python-docx members docx4j_py.docx promises, derived from an installed "
            f"python-docx {PINNED} by scripts/python_docx_subset.py. Do not edit by hand; "
            "run the script in a throwaway venv that has python-docx."
        ),
        "source": f"python-docx {PINNED}, introspected",
        "version": PINNED,
        "counts": {"interfaces": len(interfaces), "members": total, "implemented": implemented},
        "constructors": CONSTRUCTORS,
        "omitted_enums": OMITTED_ENUMS,
        "enums": enums,
        "interfaces": interfaces,
    }


def main(argv: list[str] | None = None) -> int:
    """Write, or check, ``tests/python_docx_subset.json``."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="fail if the file would change")
    args = parser.parse_args(argv)

    try:
        version = importlib.import_module("docx").__version__
    except ImportError:
        print(
            "python-docx is not importable here, so the committed "
            f"{TARGET.name} is unchanged. Install python-docx=={PINNED} in a "
            "throwaway venv and run this script with that interpreter."
        )
        return 0

    if version != PINNED:
        print(f"python-docx {version} is installed; this list is pinned to {PINNED}")
        return 1

    document = build()
    text = json.dumps(document, indent=2, ensure_ascii=False) + "\n"
    if args.check:
        current = TARGET.read_text(encoding="utf-8") if TARGET.exists() else ""
        if current != text:
            print(f"{TARGET} is out of date; run scripts/python_docx_subset.py")
            return 1
        print(f"{TARGET} is up to date")
        return 0

    TARGET.write_text(text, encoding="utf-8")
    counts = document["counts"]
    print(
        f"{TARGET}: {counts['interfaces']} interfaces, {counts['members']} members, "
        f"{counts['implemented']} implemented"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
