"""Plain text from the block model (WD section 7.1). CR-011 section 3.2 item 4.

docx4j's ``org.docx4j.model.text.PlainText``, as core-ts CR-012 ports it,
less what CR-011 Phase C brings (REQ-605's listed and grid forms, the metadata
block, alt text, page markers, ``layout: indent``).

Each paragraph is one line, a break one more (REQ-603), the text exactly as
the document holds it (REQ-602: tabs, both hyphens, zero-width and tag
characters kept), U+000A between lines and after the last. The characters are
section 6.1's: ``w:tab`` and ``w:ptab`` a tab, ``w:br`` and ``w:cr`` a line
feed, ``w:noBreakHyphen`` U+2011, ``w:softHyphen`` U+00AD, ``w:sym`` the
character its ``w:char`` names.

A table is a row a line, its cells a tab apart, a cell's paragraphs a space
apart (REQ-605's simple form, ``tables: tsv``). What the extractor adds
follows the WD's bracket convention (REQ-707): ``[n]`` at a note's reference
and at the start of the note. In the ``markup`` view insertions are
``{++text++}``, deletions ``{--text--}`` and comments ``{>>text<<}`` at their
references, replies after them (REQ-702, REQ-504), each run's marker its own,
flat, never one inside another; under ``markup_detail`` a marker's opening is
followed by ``author:`` or ``author, date:``.
"""

from __future__ import annotations

from docx4j_py.model.text.result import TextBlock, TextComment, TextResult, TextRevision, TextRun

__all__ = ["plain_text"]


def plain_text(result: TextResult) -> str:
    """The plain text of a result (WD sections 7.1 and 7.5)."""
    writer = _Writer(result)
    for block in result.blocks:
        writer.block(block)
        writer.out.append("\n")
    return "".join(writer.out)


class _Writer:
    __slots__ = ("detail", "options", "out", "result")

    def __init__(self, result: TextResult) -> None:
        self.result = result
        self.options = result.options
        self.detail = result.options.markup_detail
        self.out: list[str] = []

    def block(self, block: TextBlock) -> None:
        if block.kind == "TABLE":
            self.table(block)
        else:
            self.paragraph(block)

    def paragraph(self, block: TextBlock) -> None:
        if block.list_label is not None:
            self.out.append(block.list_label)
        runs = block.runs
        i = 0
        while i < len(runs):
            revision = runs[i].revision
            j = i
            while j < len(runs) and self._same_marker(revision, runs[j].revision):
                j += 1
            if revision is not None:
                self.out.append("{++" if revision.type == "INSERTION" else "{--")
                self._detail(revision.author, revision.date)
            for k in range(i, j):
                self.run(runs[k])
            if revision is not None:
                self.out.append("++}" if revision.type == "INSERTION" else "--}")
            i = j

    def _same_marker(self, a: TextRevision | None, b: TextRevision | None) -> bool:
        """Runs under one marker: the same type, and the same author and date where the detail shows them."""
        if a is None or b is None:
            return a is b
        if a.type != b.type:
            return False
        if self.detail == "author":
            return a.author == b.author
        if self.detail == "author_date":
            return a.author == b.author and a.date == b.date
        return True

    def _detail(self, author: str | None, date: str | None) -> None:
        if self.detail == "author":
            self.out.append(f"{author or ''}: ")
        elif self.detail == "author_date":
            self.out.append(f"{author or ''}, {date or ''}: ")

    def run(self, run: TextRun) -> None:
        kind = run.kind
        if kind in ("TEXT", "FIELD_CODE", "LIST_LABEL"):
            self.out.append(run.text)
        elif kind == "NOTE_REFERENCE":
            if self.options.note_references and self.options.notes and run.note_number > 0:
                self.out.append(f"[{run.note_number}]")
        elif kind == "NOTE_MARK":
            if run.note_number > 0:
                self.out.append(f"[{run.note_number}]")
        elif kind == "COMMENT_REFERENCE":
            comments = self.result.comments
            comment = comments.get(run.comment_id or "")
            if comment is None or (comment.parent_id is not None and comment.parent_id in comments):
                return
            self.comment(comment)
            for reply in comment.replies:
                self.comment(reply)

    def comment(self, comment: TextComment) -> None:
        self.out.append("{>>")
        self._detail(comment.author, comment.date)
        self.out.append(f"{comment.text}<<}}")

    def _text_of(self, block: TextBlock) -> str:
        """A block written on its own, for a table cell's value."""
        saved = self.out
        self.out = []
        try:
            self.block(block)
            return "".join(self.out)
        finally:
            self.out = saved

    def table(self, table: TextBlock) -> None:
        """REQ-605's simple form: a row a line, its cells a tab apart, a cell's paragraphs a space apart."""
        for r, row in enumerate(table.rows):
            if r > 0:
                self.out.append("\n")
            for c, cell in enumerate(row.cells):
                if c > 0:
                    self.out.append("\t")
                self.out.append(" ".join(self._values(cell.blocks)))

    def _values(self, blocks: list[TextBlock], out: list[str] | None = None) -> list[str]:
        """A cell's paragraphs, a nested table's included, empty ones left out."""
        if out is None:
            out = []
        for b in blocks:
            if b.kind == "TABLE":
                for row in b.rows:
                    for cell in row.cells:
                        self._values(cell.blocks, out)
            else:
                text = self._text_of(b)
                if text:
                    out.append(text)
        return out
