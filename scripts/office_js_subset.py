#!/usr/bin/env python
"""Write ``tests/office_js_subset.json`` from docx4j-core-ts's subset declaration.

CR-003 section 5 promises one committed list of the Office JS members this
package implements, and a test that keeps the package equal to it. The list is
not written here: it is *derived* from
``docx4j-core-ts/test/office-js-subset.ts``, the TypeScript engine's
compile-time assignability check, so that the two engines cannot drift apart in
silence. That file declares a ``namespace OfficeSubset`` of interfaces copied
from ``@types/office-js`` with ``load``/``sync``/``context`` removed and arrays
for collections; this script parses the interfaces, snake_cases the member
names, tags each member with the CR-003 phase that owns it, and writes the JSON.

One block is **not** derived: :data:`AHEAD_OF_TS`, the list interfaces of CR-003
Phase H. docx4j-core-ts's own phase H is still proposed, so its declaration has
no ``List`` or ``ListItem`` and none of the list members of ``Paragraph`` and
``Body``; the block is written here in the same TypeScript, parsed by the same
parser, and deleted when the TypeScript declaration gains them.

    scripts/office_js_subset.py                 # rewrite tests/office_js_subset.json
    scripts/office_js_subset.py --check         # fail if it would change
    scripts/office_js_subset.py --source PATH   # a checkout somewhere else

When the sibling checkout is absent the script says so and changes nothing: the
JSON is committed, so the test runs anywhere.

**How the names map.** ``snake_case`` of the Office JS name, with two
exceptions, both of which CR-003 section 3.2 spells out and :data:`ALIASES`
records: ``Table.getCell`` is ``cell(row_index, cell_index)``, python-docx's
name for the same thing, and ``InlinePicture.getBase64ImageSrc`` is
``get_base64()``, whose ``get_bytes()`` twin is the one a Python caller
actually wants (section 3.1: "every ``..._from_base64`` has a
``..._from_bytes`` twin"). Everything else is mechanical: ``styleBuiltIn`` is
``style_built_in``, ``getRange`` is ``get_range``,
``insertInlinePictureFromBase64`` is ``insert_inline_picture_from_base64``.
A member whose doc comment says "extension" is marked as one and the test does
not require it.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent

#: Where the TypeScript declaration lives when the sibling checkout is present.
DEFAULT_SOURCE = ROOT.parent / "docx4j-core-ts" / "test" / "office-js-subset.ts"

#: Where the committed list lives.
TARGET = ROOT / "tests" / "office_js_subset.json"

#: Office JS members this package spells differently, and why (see the module
#: docstring). ``Interface.officeJsName`` -> the Python name.
ALIASES: dict[str, str] = {
    "Table.getCell": "cell",
    "InlinePicture.getBase64ImageSrc": "get_base64",
}

#: **Ahead of docx4j-core-ts; delete when its Phase H lands.** The TypeScript
#: engine's ``office-js-subset.ts`` has no list interfaces, because its own
#: phase H is still proposed (its CR-002 section 3.9), and this package's Phase
#: H landed first (CR-003 section 18). The block is written in the same
#: TypeScript the script parses, so the members it produces have exactly the
#: shape every other interface's do; when the TypeScript declaration gains the
#: interfaces, delete this constant and the two lines that splice it in, and the
#: JSON will not change.
AHEAD_OF_TS = """
  export interface List {
    readonly id: number;
    readonly levelTypes: ArrayLike<string>;
    readonly paragraphs: ArrayLike<Paragraph>;
    levelExists(level: number): boolean;
    getLevelParagraphs(level: number): ArrayLike<Paragraph>;
    getLevelString(level: number): string;
    setLevelNumbering(level: number, listNumbering: string, formatString?: string): void;
    setLevelBullet(level: number, listBullet: string, charCode?: number, fontName?: string): void;
    setLevelIndents(level: number, textIndent: number, bulletNumberPickerIndent: number): void;
    insertParagraph(paragraphText: string, insertLocation: InsertLocation): Paragraph;
  }
  export interface ListItem {
    level: number;
    readonly listString: string;
    readonly siblingIndex: number;
    getAncestor(parentOnly?: boolean): Paragraph;
    getDescendants(directChildrenOnly?: boolean): ArrayLike<Paragraph>;
  }
  export interface ParagraphLists {
    readonly isListItem: boolean;
    readonly list: List;
    readonly listItem: ListItem;
    startNewList(): List;
    attachToList(listId: number, level: number): void;
    detachFromList(): void;
  }
  export interface BodyLists {
    readonly lists: ArrayLike<List>;
  }
"""

#: Where the members of the two "ahead of" interfaces really belong: Office JS
#: has them on ``Paragraph`` and ``Body``, and this package does too.
AHEAD_OF_TS_MERGE = {"ParagraphLists": "Paragraph", "BodyLists": "Body"}

#: The CR-003 phase that owns each interface, where every member is one phase's.
PHASE_BY_INTERFACE: dict[str, str] = {
    "SearchOptions": "B",
    "Font": "B",
    "Range": "B",
    "Paragraph": "B",
    "Body": "B",
    "Table": "C",
    "TableRow": "C",
    "TableCell": "C",
    "InlinePicture": "C",
    "ContentControl": "C",
    "Comment": "G",
    "TrackedChange": "F",
    "Document": "F",
    "XmlMapping": "E",
    "CustomXmlPart": "E",
    "CustomXmlNode": "E",
    "CustomXmlPartCollection": "E",
    "CustomXmlPrefixMappingCollection": "E",
    "CheckboxContentControl": "E",
    "DatePickerContentControl": "E",
    "ContentControlListItem": "E",
    "ListContentControl": "E",
    "PictureContentControl": "E",
    "RepeatingSectionContentControl": "E",
    "List": "H",
    "ListItem": "H",
    "ParagraphLists": "H",
    "BodyLists": "H",
}

#: Members of a Phase B interface that a later phase owns. ``Interface.name``
#: in Office JS's spelling; the phases are CR-003 section 8's.
PHASE_BY_MEMBER: dict[str, str] = {
    "Body.getComments": "G",
    "Body.getTrackedChanges": "F",
    "Body.tables": "C",
    "Body.contentControls": "C",
    "Body.inlinePictures": "C",
    "Body.insertTable": "C",
    "Body.insertInlinePictureFromBase64": "C",
    "Paragraph.getComments": "G",
    "Paragraph.insertComment": "G",
    "Paragraph.getTrackedChanges": "F",
    "Range.getComments": "G",
    "Range.insertComment": "G",
    "Range.getTrackedChanges": "F",
    # the typed content-control kinds are phase E, on a phase C interface
    "ContentControl.xmlMapping": "E",
    "ContentControl.placeholderText": "E",
    "ContentControl.appearance": "E",
    "ContentControl.color": "E",
    "ContentControl.cannotDelete": "E",
    "ContentControl.cannotEdit": "E",
    "ContentControl.removeWhenEdited": "E",
    "ContentControl.checkboxContentControl": "E",
    "ContentControl.datePickerContentControl": "E",
    "ContentControl.dropDownListContentControl": "E",
    "ContentControl.comboBoxContentControl": "E",
    "ContentControl.pictureContentControl": "E",
    "ContentControl.repeatingSectionContentControl": "E",
    "ContentControl.groupContentControl": "E",
    # the list members of Body and Paragraph, ahead of docx4j-core-ts
    "Body.lists": "H",
    "Paragraph.isListItem": "H",
    "Paragraph.list": "H",
    "Paragraph.listItem": "H",
    "Paragraph.startNewList": "H",
    "Paragraph.attachToList": "H",
    "Paragraph.detachFromList": "H",
}

_INTERFACE = re.compile(r"export interface (\w+)\s*\{")
_MEMBER = re.compile(r"^(readonly\s+)?(\w+)\s*(\(|\??\s*:)")
_COMMENTS = re.compile(r"/\*.*?\*/|//[^\n]*", re.DOTALL)
_CAMEL = re.compile(r"(?<=[a-z0-9])(?=[A-Z])")


def snake(name: str) -> str:
    """``styleBuiltIn`` to ``style_built_in``; the mechanical half of the mapping."""
    return _CAMEL.sub("_", name).lower()


def python_name(interface: str, member: str) -> str:
    """The Python spelling of an Office JS member, :data:`ALIASES` first."""
    alias = ALIASES.get(f"{interface}.{member}")
    return alias if alias is not None else snake(member)


def phase_of(interface: str, member: str) -> str:
    """The CR-003 phase that owns a member."""
    return PHASE_BY_MEMBER.get(f"{interface}.{member}", PHASE_BY_INTERFACE.get(interface, "?"))


def _bodies(text: str) -> list[tuple[str, str]]:
    """``(interface name, the text between its braces)``, in declaration order."""
    out: list[tuple[str, str]] = []
    for match in _INTERFACE.finditer(text):
        depth = 1
        index = match.end()
        while index < len(text) and depth:
            if text[index] == "{":
                depth += 1
            elif text[index] == "}":
                depth -= 1
            index += 1
        out.append((match.group(1), text[match.end() : index - 1]))
    return out


def _declarations(body: str) -> list[str]:
    """An interface body split into member declarations, comments kept.

    Split on the semicolons that are not inside a nested type --- an inline
    object type (``ArrayLike<{ prefix: string }>``) or a method's argument list
    --- so that a member written on one line with its neighbours comes out as
    one member and an inline type's fields do not come out as members at all.
    """
    parts: list[str] = []
    buffer: list[str] = []
    depth = 0
    index = 0
    while index < len(body):
        if body.startswith("//", index):
            end = body.find("\n", index)
            end = len(body) if end < 0 else end
            buffer.append(body[index:end])
            index = end
            continue
        if body.startswith("/*", index):
            end = body.find("*/", index)
            end = len(body) if end < 0 else end + 2
            buffer.append(body[index:end])
            index = end
            continue
        char = body[index]
        if char in "{([<":
            depth += 1
        elif char in "})]>":
            depth -= 1
        elif char == ";" and depth <= 0:
            parts.append("".join(buffer))
            buffer = []
            index += 1
            continue
        buffer.append(char)
        index += 1
    if "".join(buffer).strip():
        parts.append("".join(buffer))
    return parts


def parse(text: str) -> dict[str, list[dict[str, Any]]]:
    """Every ``export interface`` of the declaration, with its members.

    A member is a method when its name is followed by a parenthesis and a
    property otherwise; ``readonly`` makes a property read-only, which the JSON
    keeps because the test uses it to tell a settable view property from a
    computed one. The doc comment above a member is carried along, and one
    naming an "extension" marks the member as one, which the test does not
    require of this package.
    """
    out: dict[str, list[dict[str, Any]]] = {}
    for interface, body in _bodies(text):
        members: list[dict[str, Any]] = []
        for declaration in _declarations(body):
            note = " ".join(
                part.strip("/*").strip()
                for part in _COMMENTS.findall(declaration)
                for part in [part]
            )
            note = " ".join(
                line.strip().lstrip("*").strip() for line in note.splitlines()
            ).strip()
            stripped = _COMMENTS.sub(" ", declaration).strip()
            match = _MEMBER.match(stripped)
            if match is None:
                continue
            name = match.group(2)
            kind = "method" if match.group(3) == "(" else "property"
            members.append(
                {
                    "office_js": name,
                    "name": python_name(interface, name),
                    "kind": kind,
                    "readonly": bool(match.group(1)) and kind == "property",
                    "phase": phase_of(interface, name),
                    "extension": "extension" in note.lower(),
                    **({"note": note} if note else {}),
                }
            )
        if members:
            out[interface] = members
    return out


def build(source: Path) -> dict[str, Any]:
    """The whole JSON document, from the TypeScript declaration.

    Plus :data:`AHEAD_OF_TS`, the list interfaces this package implements and
    docx4j-core-ts does not declare yet; its ``ParagraphLists`` and
    ``BodyLists`` members are merged into ``Paragraph`` and ``Body``, which is
    where Office JS has them.
    """
    interfaces = parse(source.read_text(encoding="utf-8"))
    for name, members in parse(AHEAD_OF_TS).items():
        target = AHEAD_OF_TS_MERGE.get(name, name)
        for member in members:
            member["phase"] = phase_of(target, member["office_js"])
            member["name"] = python_name(target, member["office_js"])
        interfaces.setdefault(target, []).extend(members)
    return {
        "_": (
            "The Office JS members docx4j-python promises, derived from "
            "docx4j-core-ts/test/office-js-subset.ts by scripts/office_js_subset.py. "
            "Do not edit by hand; run the script."
        ),
        "source": "docx4j-core-ts/test/office-js-subset.ts",
        "aliases": ALIASES,
        "interfaces": {name: members for name, members in sorted(interfaces.items()) if members},
    }


def main(argv: list[str] | None = None) -> int:
    """Write, or check, ``tests/office_js_subset.json``."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--check", action="store_true", help="fail if the file would change")
    args = parser.parse_args(argv)

    if not args.source.exists():
        print(f"{args.source} is not there; the committed {TARGET.name} is unchanged")
        return 0

    document = build(args.source)
    text = json.dumps(document, indent=2, ensure_ascii=False) + "\n"
    if args.check:
        current = TARGET.read_text(encoding="utf-8") if TARGET.exists() else ""
        if current != text:
            print(f"{TARGET} is out of date; run scripts/office_js_subset.py")
            return 1
        print(f"{TARGET} is up to date")
        return 0

    TARGET.write_text(text, encoding="utf-8")
    total = sum(len(members) for members in document["interfaces"].values())
    print(f"{TARGET}: {len(document['interfaces'])} interfaces, {total} members")
    return 0


if __name__ == "__main__":
    sys.exit(main())
