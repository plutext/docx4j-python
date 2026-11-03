#!/usr/bin/env python3
"""docx-platform-tests adapter, protocol version 1, for docx4j-python.

The suite's runner calls this once per scenario::

    run_adapter.py --protocol-version 1 --operation op.json --input in.docx --output out.docx

Exit 0: the output is written. 1: an error, the traceback on stderr. 2: the
operation is declined, the reason one line on stdout, nothing written. 3: another
protocol version.

Standard library plus ``docx4j_py`` (the ``docx4j`` distribution on PyPI), so it
runs under whatever interpreter has that installed. README.md says, operation by
operation, what is library API and what is glue.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import traceback

PROTOCOL_VERSION = "1"


class Declined(Exception):
    """An operation, or an input to one, docx4j-python has no API for: exit 2."""


class InputError(Exception):
    """The input does not hold what the operation addresses (no such text, table or row): exit 1."""


# ---------------------------------------------------------------------------
# small pieces of glue shared by several operations
# ---------------------------------------------------------------------------

#: The ``runFormatting`` keys the protocol defines; any other is declined rather than ignored.
RUN_FORMATTING_KEYS = {"bold", "italic", "fontSizeHalfPoints"}


def literal(text):
    """`text` as a plain search query: ``^`` starts one of Word's Find codes, and ``^^`` is a caret."""
    return text.replace("^", "^^")


def first_occurrence(package, text):
    """The first match of `text` in the main document, paragraphs in document order, case-sensitive.

    The content API's ``Body.search``: a match is found on a paragraph's whole
    text, so it may cross runs; tables and content controls are descended into.
    """
    if not text:
        raise InputError("the text to find is empty")
    hits = package.body.search(literal(text), match_case=True, limit=1)
    if not hits:
        raise InputError(f"{text!r} is not in any paragraph of the main document")
    return hits[0]


def apply_run_formatting(font, formatting, *, has_text=True):
    """``runFormatting`` onto a content-API ``Font`` (a paragraph's or a range's)."""
    if not formatting:
        return
    unknown = set(formatting) - RUN_FORMATTING_KEYS
    if unknown:
        raise Declined(
            f"runFormatting keys this adapter does not know: {', '.join(sorted(unknown))}"
        )
    if not has_text:
        raise Declined(
            "runFormatting on a paragraph with no text: there is no run to carry it"
        )
    if "bold" in formatting:
        font.bold = bool(formatting["bold"])
    if "italic" in formatting:
        font.italic = bool(formatting["italic"])
    if "fontSizeHalfPoints" in formatting:
        # the content API's size is in points
        font.size = formatting["fontSizeHalfPoints"] / 2


def table_at(package, index):
    """The `index`-th table of the main document's body (its own tables, as Office JS ``Body.tables``)."""
    tables = package.body.tables
    index = int(index)
    if not 0 <= index < len(tables):
        raise InputError(f"no table {index}; the body has {len(tables)}")
    return tables[index]


def set_default_header_or_footer(document, kind, text):
    """The last section's default header or footer: made where there is none, its content replaced where there is.

    The facade's ``Section.header`` / ``footer`` and ``is_linked_to_previous``
    (which creates the part, its relationship, its content type and the
    reference), then the content API's ``Body.insert_text(..., "Replace")`` on
    that part's body. Glue: a body with no ``w:sectPr`` is given an empty one
    first, which is what Word assumes such a body has.
    """
    from docx4j_py.wml import el

    if len(document.sections) == 0:
        body = document.element.body
        body.sect_pr = el.sectPr()
        body.sect_pr.parent = body
    section = document.sections[-1]
    definition = section.header if kind == "header" else section.footer
    definition.is_linked_to_previous = False
    definition.part.body.insert_text(text, location="Replace")


# ---------------------------------------------------------------------------
# the operations
# ---------------------------------------------------------------------------


def load(path):
    from docx4j_py import load as load_package

    return load_package(path)


def new_package():
    from docx4j_py import create_package

    return create_package()


def accept_all(op, source):
    from docx4j_py.model import revisions

    package = load(source)
    revisions.accept_all(package)
    return package


def reject_all(op, source):
    from docx4j_py.model import revisions

    package = load(source)
    revisions.reject_all(package)
    return package


def replace_first(op, source):
    package = load(source)
    first_occurrence(package, op["findText"]).insert_text(
        op["replaceText"], location="Replace"
    )
    return package


def format_first(op, source):
    package = load(source)
    hit = first_occurrence(package, op["findText"])
    apply_run_formatting(hit.font, op["runFormatting"])
    return package


def compose_paragraphs(op, source):
    package = new_package()
    for descriptor in op["paragraphDescriptorList"]:
        text = descriptor["paragraphText"]
        paragraph = package.body.insert_paragraph(text)
        apply_run_formatting(
            paragraph.font, descriptor.get("runFormatting"), has_text=bool(text)
        )
    return package


def compose_table(op, source):
    rows = op["tableCellTextRows"]
    if not rows or not rows[0]:
        raise InputError("tableCellTextRows is empty")
    columns = len(rows[0])
    if any(len(row) != columns for row in rows):
        raise Declined(
            "tableCellTextRows is ragged; insert_table makes a rectangular grid"
        )
    package = new_package()
    package.body.insert_table(len(rows), columns, values=rows)
    return package


def compose_hyperlink(op, source):
    package = new_package()
    paragraph = package.body.insert_paragraph(op["hyperlinkDisplayText"])
    paragraph.get_range("Content").hyperlink = op["hyperlinkTargetUrl"]
    return package


def compose_header(op, source):
    from docx4j_py.docx import Document

    document = Document()  # create_package()
    document.package.body.insert_paragraph(op["bodyText"])
    set_default_header_or_footer(document, "header", op["headerText"])
    return document.package


def compose_numbered_list(op, source):
    from docx4j_py.model.content.lists import NUMBERING

    texts = op["listItemTexts"]
    levels = op.get("itemLevels")
    if levels is None:
        levels = [0] * len(texts)
    if not texts or len(levels) != len(texts):
        raise InputError("listItemTexts is empty, or itemLevels does not match it")
    number_format = op["numberFormat"]
    if number_format == "bullet":
        kind, numbering = "Bullet", None
    else:
        numbering = {fmt: name for name, fmt in NUMBERING.items()}.get(number_format)
        if numbering is None:
            raise Declined(
                f"numberFormat {number_format!r}: setLevelNumbering writes "
                f"{', '.join(sorted(NUMBERING.values()))} and the bullet set only"
            )
        kind = "Number"
    package = new_package()
    paragraphs = [package.body.insert_paragraph(text) for text in texts]
    new_list = paragraphs[0].start_new_list(kind=kind)
    for paragraph, level in zip(paragraphs, levels):
        paragraph.attach_to_list(new_list.id, int(level))
    if numbering is not None:
        for level in sorted({int(level) for level in levels}):
            new_list.set_level_numbering(level, numbering)
    return package


def compose_compatibility_mode(op, source):
    from docx4j_py.model.content import ContentError

    package = new_package()
    try:
        package.compatibility_mode = op["compatibilityMode"]
    except ContentError as error:
        if getattr(error, "code", None) == "compatibility.mode_invalid":
            raise Declined(str(error)) from None
        raise
    package.body.insert_paragraph(op["bodyText"])
    return package


def append_paragraph(op, source):
    package = load(source)
    text = op["paragraphText"]
    paragraph = package.body.insert_paragraph(text, location="End")
    apply_run_formatting(paragraph.font, op.get("runFormatting"), has_text=bool(text))
    return package


def insert_paragraph_after(op, source):
    package = load(source)
    paragraph = first_occurrence(package, op["anchorText"]).paragraph
    paragraph.insert_paragraph(op["paragraphText"], location="After")
    return package


def append_table_row(op, source):
    package = load(source)
    table = table_at(package, op["tableIndex"])
    texts = op["cellTexts"]
    columns = len(table.column_widths_twips())
    if len(texts) > columns:
        raise Declined(
            f"{len(texts)} cell texts for a table of {columns} columns; add_rows keeps the table's grid"
        )
    table.add_rows(1, location="End", values=[texts])
    return package


def delete_table_row(op, source):
    package = load(source)
    table = table_at(package, op["tableIndex"])
    row = int(op["rowIndex"])
    if not 0 <= row < table.row_count:
        raise InputError(f"no row {row}; the table has {table.row_count}")
    table.delete_rows(row, 1)
    return package


def set_table_cell_text(op, source):
    package = load(source)
    table = table_at(package, op["tableIndex"])
    table.cell(int(op["rowIndex"]), int(op["columnIndex"])).value = op[
        "replacementCellText"
    ]
    return package


def merge_cells(op, source):
    raise Declined(
        "docx4j-python has no cell-merge API: the content API adds and removes rows and columns only, "
        "and the python-docx facade's _Cell.merge raises NotImplementedError (CR-003 section 20)"
    )


def add_comment(op, source):
    from docx4j_py.model.content import Author

    package = load(source)
    package.author = Author(
        op["commentAuthorName"], initials=op["commentAuthorInitials"]
    )
    first_occurrence(package, op["anchorText"]).insert_comment(op["commentText"])
    return package


def remove_all_comments(op, source):
    package = load(source)
    for body in package.bodies():  # the main document, headers, footers and notes
        for (
            comment
        ) in body.get_comments():  # replies are nested; delete() takes the thread
            comment.delete()
    return package


def apply_paragraph_style(op, source):
    package = load(source)
    paragraph = first_occurrence(package, op["anchorText"]).paragraph
    paragraph.style = op["paragraphStyleId"]
    return package


def apply_numbering(op, source):
    package = load(source)
    paragraph = first_occurrence(package, op["anchorText"]).paragraph
    paragraph.attach_to_list(
        int(op["numberingInstanceId"]), int(op["indentationLevel"])
    )
    return package


def set_default_footer(op, source):
    from docx4j_py.docx import Document

    document = Document(source)
    set_default_header_or_footer(document, "footer", op["footerText"])
    return document.package


OPERATIONS = {
    "acceptAllTrackedChanges": accept_all,
    "rejectAllTrackedChanges": reject_all,
    "replaceFirstTextOccurrence": replace_first,
    "composeDocumentWithParagraphs": compose_paragraphs,
    "composeDocumentWithTable": compose_table,
    "composeDocumentWithHyperlink": compose_hyperlink,
    "composeDocumentWithHeaderText": compose_header,
    "composeDocumentWithNumberedList": compose_numbered_list,
    "composeDocumentWithCompatibilityMode": compose_compatibility_mode,
    "appendParagraphWithText": append_paragraph,
    "insertParagraphAfterAnchorText": insert_paragraph_after,
    "appendTableRow": append_table_row,
    "deleteTableRowAtIndex": delete_table_row,
    "setTableCellText": set_table_cell_text,
    "mergeTableCellsInRow": merge_cells,
    "addCommentOnFirstTextOccurrence": add_comment,
    "removeAllComments": remove_all_comments,
    "applyParagraphStyleToAnchor": apply_paragraph_style,
    "applyNumberingToAnchorParagraph": apply_numbering,
    "formatFirstTextOccurrence": format_first,
    "setDefaultFooterText": set_default_footer,
}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--protocol-version", required=True)
    parser.add_argument("--operation", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args(argv)

    if args.protocol_version != PROTOCOL_VERSION:
        print(
            f"the docx4j-python adapter speaks protocol v{PROTOCOL_VERSION}, not v{args.protocol_version}"
        )
        return 3

    with open(args.operation, encoding="utf-8") as handle:
        op = json.load(handle)
    name = op.get("operationName")
    handler = OPERATIONS.get(name)
    if handler is None:
        print(f"the docx4j-python adapter does not implement operation {name!r}")
        return 2
    existed = os.path.exists(args.output)
    try:
        package = handler(op, args.input)
        package.save(
            args.output
        )  # the default save: the normal form Word reads, normalising on
    except Declined as declined:
        _discard(args.output, existed)
        print(str(declined).replace("\n", " "))
        return 2
    except InputError as error:
        _discard(args.output, existed)
        print(f"{name}: {error}", file=sys.stderr)
        return 1
    except Exception:  # noqa: BLE001 - every failure is an error, its traceback the diagnostics
        _discard(args.output, existed)
        traceback.print_exc()
        return 1
    return 0


def _discard(path, existed):
    """Nothing is left at --output after a failure (a save may have begun), unless it was there before."""
    if existed:
        return
    try:
        os.remove(path)
    except OSError:
        pass


if __name__ == "__main__":
    sys.exit(main())
