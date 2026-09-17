"""``Document``: what ``from docx4j_py.docx import Document`` gives you.

CR-003 section 3.6, Phase J. python-docx's ``docx.Document`` is a *function*
that opens a package and hands back a ``docx.document.Document``; here it is
the class itself, so ``Document("in.docx")`` reads the same and ``isinstance``
works. What is underneath is a
:class:`~docx4j_py.openpackaging.packages.WordprocessingMLPackage`, reachable
as ``document.package``, and the content API is reachable as
``document.part.body`` --- which is the point of the facade: python-docx code
runs, and everything python-docx has no word for (tracked changes, comments,
addresses, custom XML, byte-for-byte untouched parts) is one attribute away.

``Document()`` with no argument is ``create_package()``: **A4**, 2.54 cm
margins, Word 2013 compatibility mode and docx4j's default styles. python-docx's
own default template is US Letter with its own style set; copying it would mean
shipping a second template, so the difference is a recorded departure (section
20, departure 1) rather than a bug.
"""

from __future__ import annotations

import os
from collections.abc import Iterator
from typing import IO, Any

from docx4j_py.docx._base import BlockItemContainer, Story, recording
from docx4j_py.docx.enums import WD_BREAK, WD_SECTION
from docx4j_py.docx.shared import Length
from docx4j_py.wml import Body as BodyElement
from docx4j_py.wml import P, el

__all__ = ["Document", "read_image"]


def read_image(image_path_or_stream: str | os.PathLike[str] | IO[bytes] | bytes) -> bytes:
    """The bytes of an image, from a path, a stream or the bytes themselves."""
    if isinstance(image_path_or_stream, (bytes, bytearray, memoryview)):
        return bytes(image_path_or_stream)
    read = getattr(image_path_or_stream, "read", None)
    if callable(read):
        seek = getattr(image_path_or_stream, "seek", None)
        if callable(seek):
            try:
                seek(0)
            except (OSError, ValueError):  # pragma: no cover - a non-seekable stream
                pass
        return read()
    with open(image_path_or_stream, "rb") as handle:
        return handle.read()


class Document(BlockItemContainer):
    """A WordprocessingML document: python-docx's ``Document``."""

    __slots__ = ("_body_story", "_package")

    def __init__(
        self, path_or_stream: str | os.PathLike[str] | IO[bytes] | bytes | None = None
    ) -> None:
        """Open a ``.docx``, or create one when nothing is given.

        Args:
            path_or_stream: a path, a ``PathLike``, a readable binary stream,
                the bytes of a package, or None for a new document.
        """
        from docx4j_py import WordprocessingMLPackage, create_package
        from docx4j_py.model.content.body import body_of

        if path_or_stream is None:
            package = create_package()
        else:
            package = WordprocessingMLPackage.load(path_or_stream)
        self._package = package
        self._body_story = Story(body_of(package.get_main_document_part()), self)

    def __repr__(self) -> str:
        """``<docx4j_py.docx Document, 12 paragraphs>``."""
        return f"<docx4j_py.docx Document, {len(self.paragraphs)} paragraphs>"

    # -- what is underneath ------------------------------------------------

    @property
    def package(self) -> Any:
        """The ``WordprocessingMLPackage``. An extension; python-docx has no package."""
        return self._package

    @property
    def part(self) -> Any:
        """The main document part. An extension in kind: python-docx's is its own class."""
        return self._package.get_main_document_part()

    @property
    def element(self) -> Any:
        """The typed ``w:document``. python-docx's is an lxml element."""
        return self.part.contents

    @property
    def _element(self) -> Any:
        """python-docx's private name for :attr:`element`."""
        return self.element

    @property
    def _story(self) -> Story:
        """The main document part, as a story."""
        return self._body_story

    @property
    def _container(self) -> BodyElement:
        """The ``w:body``."""
        return self._body_story.body.container

    # -- reading -----------------------------------------------------------

    @property
    def core_properties(self) -> Any:
        """``/docProps/core.xml`` as python-docx's :class:`CoreProperties`."""
        from docx4j_py.docx.coreprops import CoreProperties

        return CoreProperties(self._package)

    @property
    def settings(self) -> Any:
        """``/word/settings.xml`` as python-docx's ``Settings``."""
        from docx4j_py.docx.settings import Settings

        return Settings(self._package)

    @property
    def styles(self) -> Any:
        """The styles the document defines. **Unmarshals** ``/word/styles.xml``."""
        from docx4j_py.docx.styles import Styles

        return Styles(self._package)

    @property
    def sections(self) -> Any:
        """The document's sections, in order."""
        from docx4j_py.docx.section import Sections

        return Sections(self._body_story, self)

    @property
    def inline_shapes(self) -> Any:
        """Every ``wp:inline`` in the main document part."""
        from docx4j_py.docx.shape import InlineShapes

        return InlineShapes(self._body_story)

    @property
    def comments(self) -> Any:
        """The comments the document carries (python-docx 1.2)."""
        from docx4j_py.docx.comments import Comments

        return Comments(self)

    # -- writing -----------------------------------------------------------

    def add_heading(self, text: str = "", level: int = 1) -> Any:
        """A heading paragraph at the end: level 0 is ``Title``, 1 to 9 ``Heading N``.

        The style is applied through the content API's ``ensure_style``, so a
        built-in heading the document does not define is **defined** rather than
        left dangling (CR-003 section 14.9) --- which matters here, because
        ``create_package()`` defines Heading 1 to 4 and no more.
        """
        if not 0 <= level <= 9:
            raise ValueError(f"level must be in range 0-9, got {level}")
        style = "Title" if level == 0 else f"Heading {level}"
        return self.add_paragraph(text, style)

    def add_page_break(self) -> Any:
        """A paragraph holding nothing but a page break."""
        with recording(self._body_story.body, "docx.add_page_break"):
            paragraph = self.add_paragraph()
            paragraph.add_run().add_break(WD_BREAK.PAGE)
            return paragraph

    def add_table(self, rows: int, cols: int, style: Any = None) -> Any:
        """A table of `rows` by `cols` at the end, `style` applied.

        The columns divide the section's text width, which is what python-docx
        computes for itself from ``page_width`` less the margins and what the
        content API's ``insert_table`` reads from ``w:sectPr``.
        """
        with recording(self._body_story.body, "docx.add_table"):
            table = super().add_table(rows, cols, None)
            table.style = style
            return table

    def add_picture(
        self,
        image_path_or_stream: str | os.PathLike[str] | IO[bytes] | bytes,
        width: int | Length | None = None,
        height: int | Length | None = None,
    ) -> Any:
        """A picture in a new paragraph at the end, as its own ``InlineShape``."""
        with recording(self._body_story.body, "docx.add_picture"):
            run = self.add_paragraph().add_run()
            return run.add_picture(image_path_or_stream, width, height)

    def add_section(self, start_type: Any = WD_SECTION.NEW_PAGE) -> Any:
        """A new section at the end, its page setup cloned from the last one.

        python-docx's rule exactly: the sentinel ``w:body/w:sectPr`` is copied
        into a new paragraph at the end --- which becomes the *previous*
        section's break --- and the sentinel keeps the page setup but loses its
        header and footer references, so the new section inherits them.
        """
        from docx4j_py.child import deep_copy, link_parents
        from docx4j_py.docx.section import Section
        from docx4j_py.wml import PPr

        body = self._container
        with recording(self._body_story.body, "docx.add_section"):
            sentinel = getattr(body, "sect_pr", None)
            if sentinel is None:
                sentinel = el.sectPr()
                body.sect_pr = sentinel
                sentinel.parent = body
            clone = deep_copy(sentinel)
            paragraph = P(p_pr=PPr(sect_pr=clone))
            link_parents(paragraph)
            self._body_story.body.content.append(paragraph)
            references = sentinel.header_reference_or_footer_reference
            while len(references):
                del references[0]
            section = Section(sentinel, self._body_story, self)
            section.start_type = start_type
            return section

    def add_comment(
        self,
        runs: Any,
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
    ) -> Any:
        """A comment anchored on `runs` (python-docx 1.2), over CR-003 Phase G.

        `runs` is one :class:`~docx4j_py.docx.text.Run` or a non-empty sequence
        of them; only the first and the last are used, and a single run serves
        as both. The four comment parts are created if the document has none,
        with their relationships and content types, and ``CommentText`` /
        ``CommentReference`` are added to ``/word/styles.xml``.
        """
        from docx4j_py.docx.comments import add_comment

        return add_comment(self, runs, text, author, initials)

    def save(self, path_or_stream: str | os.PathLike[str] | IO[bytes] | None = None) -> Any:
        """Save the document; with no argument, return the bytes.

        Only the parts that were unmarshalled are re-marshalled; every other
        part is written back byte for byte from the source (CR-002's rule).
        """
        if path_or_stream is None:
            return self._package.save()
        return self._package.save(path_or_stream)

    # -- the content of the body ------------------------------------------

    def iter_inner_content(self) -> Iterator[Any]:
        """Each ``Paragraph`` or ``Table`` of the body, in document order."""
        return super().iter_inner_content()
