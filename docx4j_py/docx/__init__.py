"""``docx4j_py.docx``: a structural subset of python-docx's API over this engine.

CR-003 section 3.6, Phase J (section 20). python-docx is the API in every
tutorial, in every model's memory and in every "write me a script that makes a
Word document" answer, so code already written --- by people or by models ---
runs here:

    >>> from docx4j_py.docx import Document, Pt, Inches
    >>> doc = Document()                                     # or Document("in.docx")
    >>> _ = doc.add_heading("Report", level=1)
    >>> p = doc.add_paragraph("Plain, ")
    >>> p.add_run("bold").bold = True
    >>> p.style = "Heading 2"                                # display name, as python-docx
    >>> doc.add_table(rows=2, cols=2).cell(0, 0).text = "a"
    >>> for paragraph in doc.paragraphs: _ = paragraph.style.name, paragraph.text
    >>> data = doc.save()                                    # or doc.save("out.docx")

**Names and semantics are python-docx's**, at the version
``tests/python_docx_subset.json`` pins; ``tests/content/test_python_docx_subset.py``
asserts every member that list marks implemented. The Office JS vocabulary of
:mod:`docx4j_py.model.content` remains the primary API and the one the MCP
tools are written over; this namespace exists so existing code runs.

**What python-docx code gets from underneath that python-docx does not have.**
The whole schema (``paragraph._p`` is the typed ``P``, not lxml, and
``paragraph.view`` is the content API's ``Paragraph``), parts never touched
written back byte for byte, a
:class:`~docx4j_py.model.content.reports.ChangeReport` for every mutation
(``doc.package.last_change``), ``dry_run``, tracked changes, comments, lists,
addresses and custom XML --- all of it through ``doc.package`` and
``doc.part.body``.

**What is deliberately not here** is in CR-003 section 20 and in
``tests/python_docx_subset.json``: ``_Cell.merge``, ``Styles.latent_styles``,
``ParagraphFormat.tab_stops``, and python-docx's own ``lxml`` internals, which
were never the promise (decided question 7).
"""

from __future__ import annotations

from docx4j_py.docx.comments import Comment, Comments
from docx4j_py.docx.coreprops import CoreProperties
from docx4j_py.docx.document import Document
from docx4j_py.docx.enums import (
    MSO_COLOR_TYPE,
    MSO_THEME_COLOR,
    MSO_THEME_COLOR_INDEX,
    WD_ALIGN_PARAGRAPH,
    WD_ALIGN_VERTICAL,
    WD_BREAK,
    WD_BREAK_TYPE,
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_COLOR,
    WD_COLOR_INDEX,
    WD_HEADER_FOOTER,
    WD_HEADER_FOOTER_INDEX,
    WD_INLINE_SHAPE,
    WD_INLINE_SHAPE_TYPE,
    WD_LINE_SPACING,
    WD_ORIENT,
    WD_ORIENTATION,
    WD_PARAGRAPH_ALIGNMENT,
    WD_ROW_HEIGHT,
    WD_ROW_HEIGHT_RULE,
    WD_SECTION,
    WD_SECTION_START,
    WD_STYLE_TYPE,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
    WD_TABLE_ALIGNMENT,
    WD_TABLE_DIRECTION,
    WD_UNDERLINE,
)
from docx4j_py.docx.section import Section, Sections, _Footer, _Header
from docx4j_py.docx.settings import Settings
from docx4j_py.docx.shape import Drawing, InlineShape, InlineShapes
from docx4j_py.docx.shared import (
    Cm,
    Emu,
    Inches,
    Length,
    Mm,
    Pt,
    RGBColor,
    Twips,
)
from docx4j_py.docx.styles import (
    BaseStyle,
    CharacterStyle,
    ParagraphStyle,
    Styles,
    _NumberingStyle,
    _TableStyle,
)
from docx4j_py.docx.table import Table, _Cell, _Column, _Columns, _Row, _Rows
from docx4j_py.docx.text import (
    ColorFormat,
    Font,
    Hyperlink,
    Paragraph,
    ParagraphFormat,
    RenderedPageBreak,
    Run,
)

__all__ = [
    "MSO_COLOR_TYPE",
    "MSO_THEME_COLOR",
    "MSO_THEME_COLOR_INDEX",
    "WD_ALIGN_PARAGRAPH",
    "WD_ALIGN_VERTICAL",
    "WD_BREAK",
    "WD_BREAK_TYPE",
    "WD_CELL_VERTICAL_ALIGNMENT",
    "WD_COLOR",
    "WD_COLOR_INDEX",
    "WD_HEADER_FOOTER",
    "WD_HEADER_FOOTER_INDEX",
    "WD_INLINE_SHAPE",
    "WD_INLINE_SHAPE_TYPE",
    "WD_LINE_SPACING",
    "WD_ORIENT",
    "WD_ORIENTATION",
    "WD_PARAGRAPH_ALIGNMENT",
    "WD_ROW_HEIGHT",
    "WD_ROW_HEIGHT_RULE",
    "WD_SECTION",
    "WD_SECTION_START",
    "WD_STYLE_TYPE",
    "WD_TABLE_ALIGNMENT",
    "WD_TABLE_DIRECTION",
    "WD_TAB_ALIGNMENT",
    "WD_TAB_LEADER",
    "WD_UNDERLINE",
    "BaseStyle",
    "CharacterStyle",
    "Cm",
    "ColorFormat",
    "Comment",
    "Comments",
    "CoreProperties",
    "Document",
    "Drawing",
    "Emu",
    "Font",
    "Hyperlink",
    "Inches",
    "InlineShape",
    "InlineShapes",
    "Length",
    "Mm",
    "Paragraph",
    "ParagraphFormat",
    "ParagraphStyle",
    "Pt",
    "RGBColor",
    "RenderedPageBreak",
    "Run",
    "Section",
    "Sections",
    "Settings",
    "Styles",
    "Table",
    "Twips",
    "_Cell",
    "_Column",
    "_Columns",
    "_Footer",
    "_Header",
    "_NumberingStyle",
    "_Row",
    "_Rows",
    "_TableStyle",
]
