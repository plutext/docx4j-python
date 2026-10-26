"""What every facade view shares: its story, and the block-container verbs.

CR-003 section 3.6, Phase J. python-docx's proxies hold an lxml element and a
``_parent`` that can reach the part; the facade's hold a typed element from
CR-001's model and a :class:`Story`, which is the content API's
:class:`~docx4j_py.model.content.Body` for the part the element lives in ---
``/word/document.xml``, a header, a footer --- plus the
:class:`~docx4j_py.docx.document.Document` the caller started from.

Everything that mutates goes through the content API, or, where python-docx has
a member the content API does not, through the typed tree inside a
:class:`~docx4j_py.model.content.reports.recording` block, so that
``pkg.last_change`` and ``dry_run`` keep working for facade calls exactly as
they do for the Office JS ones (CR-003 decided question 4).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from docx4j_py.model.content.body import Body
from docx4j_py.model.content.reports import recording
from docx4j_py.wml import P, Tbl

if TYPE_CHECKING:  # pragma: no cover - annotations only
    from docx4j_py.docx.document import Document
    from docx4j_py.docx.styles import ParagraphStyle

__all__ = ["BlockItemContainer", "Story", "recording"]


class Story:
    """The part a view belongs to, and the document it was reached through.

    One per part rather than one per view: the content API's ``Body`` finds the
    live list holding any element from its parent pointers, so a single body
    over ``/word/document.xml`` serves a paragraph in a table cell as well as
    one at the top level.
    """

    __slots__ = ("body", "document")

    def __init__(self, body: Body, document: Document | None = None) -> None:
        """Wrap the content API body of a part."""
        #: The content API :class:`~docx4j_py.model.content.Body` of the part.
        self.body = body
        #: The :class:`~docx4j_py.docx.document.Document`, or None for a part
        #: reached without one (a header read straight off the package).
        self.document = document

    @property
    def part(self) -> Any:
        """The engine part: python-docx's ``.part``, and the same promise."""
        return self.body.part

    @property
    def package(self) -> Any:
        """The ``WordprocessingMLPackage``. An extension; python-docx has none."""
        return self.body.package

    def __repr__(self) -> str:
        """``<Story /word/document.xml>``."""
        return f"<Story {getattr(self.part, 'part_name', '?')}>"


class BlockItemContainer:
    """The base of every proxy that holds paragraphs and tables.

    python-docx's ``BlockItemContainer``: ``_Body``, ``_Cell``, ``_Header``,
    ``_Footer`` and 1.2's ``Comment`` all answer ``paragraphs``, ``tables``,
    ``add_paragraph``, ``add_table`` and ``iter_inner_content``.

    **Only direct children count**, as in python-docx: a paragraph inside a
    table cell is not in ``document.paragraphs``, and neither is one inside a
    ``w:ins`` or a content control. The content API's ``body.paragraphs``
    descends into all three, and stays reachable as ``document.part.body``
    (section 20, departure 5).
    """

    __slots__ = ()

    @property
    def _story(self) -> Story:
        """The part this container is in."""
        raise NotImplementedError

    @property
    def _container(self) -> Any:
        """The typed element whose block children this container is: a ``w:body``,
        ``w:tc``, ``w:hdr``, ``w:ftr`` or ``w:comment``."""
        raise NotImplementedError

    @property
    def _content_body(self) -> Body:
        """A content API body scoped to this container."""
        story = self._story
        container = self._container
        if story.body.container is container:
            return story.body
        return story.body.sub(container)

    # -- reading -----------------------------------------------------------

    @property
    def paragraphs(self) -> list[Any]:
        """The ``w:p`` children of this container, in document order."""
        from docx4j_py.docx.text import Paragraph

        story = self._story
        return [
            Paragraph(item, story)
            for item in self._content_body.content
            if isinstance(item, P)
        ]

    @property
    def tables(self) -> list[Any]:
        """The ``w:tbl`` children of this container, in document order."""
        from docx4j_py.docx.table import Table

        story = self._story
        return [
            Table(item, story) for item in self._content_body.content if isinstance(item, Tbl)
        ]

    def iter_inner_content(self) -> Iterator[Any]:
        """Each ``Paragraph`` or ``Table`` child, in document order (python-docx 1.2)."""
        from docx4j_py.docx.table import Table
        from docx4j_py.docx.text import Paragraph

        story = self._story
        for item in list(self._content_body.content):
            if isinstance(item, P):
                yield Paragraph(item, story)
            elif isinstance(item, Tbl):
                yield Table(item, story)

    # -- writing -----------------------------------------------------------

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Any:
        """A paragraph at the end of this container, with `text` in one run.

        A tab in `text` becomes ``w:tab`` and a newline or carriage return a
        ``w:br``, as python-docx writes them; any other control character is
        refused with ``ValueError``, as python-docx 1.2.0 refuses it (CR-013
        Phase A). `style` may be a display name, a style id or a style object;
        a built-in style the document does not define is **defined** rather
        than left dangling (CR-003 section 14.9).
        """
        from docx4j_py.docx.text import Paragraph, check_python_docx_text

        check_python_docx_text(text)
        body = self._content_body
        with recording(body, "docx.add_paragraph"):
            view = body.insert_paragraph(text.replace("\r", "\n") if text else "")
            paragraph = Paragraph(view.element, self._story)
            if style is not None:
                paragraph.style = style
            return paragraph

    def add_table(self, rows: int, cols: int, width: Any = None) -> Any:
        """A table of `rows` by `cols` at the end of this container.

        `width` is python-docx's third argument, the total width the columns
        divide evenly; None asks the content API for the section's own text
        width, which is what its ``insert_table`` uses and what python-docx's
        ``Document.add_table`` computes for itself.

        ``rows=0`` is python-docx's way of building a table to add rows to
        later; the content API refuses a table with no rows, so one is built
        and then removed.
        """
        return self._add_table(rows, cols, width)

    def _add_table(self, rows: int, cols: int, width: Any = None, before: Any = None) -> Any:
        """:meth:`add_table`, the table at the end or, given `before` (a content API paragraph), in front of it."""
        from docx4j_py.docx.table import Table
        from docx4j_py.model.content.table import insert_table_into

        if cols < 1:
            raise ValueError(f"a table needs at least one column, not {cols}")
        widths = None if width is None else column_widths(width, cols)  # before anything changes (CR-022)
        body = self._content_body
        with recording(body, "docx.add_table"):
            if before is None:
                view = body.insert_table(max(rows, 1), cols)
            else:
                view = insert_table_into(body, max(rows, 1), cols, location="Before", target=before)
            table = Table(view.element, self._story)
            if widths is not None:
                table._set_column_widths(widths)
            if rows < 1:
                for row in list(view.element.content):
                    view.element.content.remove(row)
            return table


def column_widths(width: Any, columns: int) -> list[int]:
    """`width` (a python-docx ``Length``) split evenly between `columns`, in twips, each within Word's range.

    What ``add_table(rows, cols, width)`` writes on every cell's ``w:tcW``
    (and on ``w:gridCol``): each column is refused beyond -32,768 to 32,767
    twips, which Word would read wrapped (CR-022 Q4), before the table is
    built. The widths are the facade's own sums, so the message is in twips.
    """
    from docx4j_py.docx.shared import twips_of
    from docx4j_py.model.values import OWN, measure_in_range

    return [
        measure_in_range("w:tcW/@w:w", each, OWN, label="column width", context="dxa")
        for each in _even(twips_of(width) or 0, columns)
    ]


def within_rule(
    attribute: str,
    value: int,
    *,
    setter: str,
    rule_attribute: str,
    rule: str | None,
    absent: str,
    label: str | None = None,
) -> None:
    """`value`, already set, within the range of `rule` (None: `rule_attribute` absent, which Word reads as `absent`).

    For a rule setter: ``line_spacing_rule`` and ``_Row.height_rule`` change how Word reads
    a value already in the tree (a line of 0 twips at least is no line; exact, Word reads
    it as at least 0), so the value is held to the new rule's row before the rule is
    written (CR-022 section 10, decision 4). `setter` is the facade property that sets the
    value (``line_spacing``, ``height``), `label` what the message calls it (the setter's
    name by default). ``MeasureRangeError``, its hint naming the value and the rule.
    """
    from docx4j_py.model.values import OWN, measure_in_range

    rule_text = f'{rule_attribute}="{rule}"' if rule is not None else f"no {rule_attribute} ({absent})"
    label = label or setter.replace("_", " ")
    unit = "240ths of a line" if attribute == "w:spacing/@w:line" and (rule or absent) == "auto" else "twips"
    measure_in_range(
        attribute,
        value,
        OWN,
        label=f"{label}, under {rule_text},",
        context=rule,
        hint=(
            f"the {label} already set, {value:,} {unit}, is beyond the range of {rule_text}, {{span}}: "
            f"set {setter} within it first, or choose another rule"
        ),
    )


def _even(total: int, columns: int) -> list[int]:
    """`total` twips split between `columns`, the remainder on the last one."""
    each = total // columns
    widths = [each] * columns
    widths[-1] = total - each * (columns - 1)
    return widths
