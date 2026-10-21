"""``InlineShape`` and ``InlineShapes``: a ``wp:inline`` as python-docx has it.

CR-003 section 3.6, Phase J. python-docx's ``InlineShape`` is a thin proxy over
the ``wp:inline`` element: ``width`` and ``height`` in EMU, and a ``type`` read
from the ``a:graphicData/@uri``. It sets **only** the axis it is given, where
the content API's :class:`~docx4j_py.model.content.InlinePicture` keeps the
aspect ratio; both are here, the facade's following python-docx (section 20,
departure 7) and the content API's a property access away through
``document.part.body``.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

from docx4j_py.docx._base import Story, recording
from docx4j_py.docx.enums import WD_INLINE_SHAPE
from docx4j_py.docx.shared import Emu, Length

__all__ = ["Drawing", "InlineShape", "InlineShapes"]

#: The ``a:graphicData/@uri`` of each shape kind python-docx knows.
_URI_PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_URI_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


class Drawing:
    """A ``w:drawing``, as ``Run.iter_inner_content`` yields it (python-docx 1.2).

    python-docx offers only the element; so does this.
    """

    __slots__ = ("_drawing", "_parent")

    def __init__(self, drawing: Any, parent: Any = None) -> None:
        """Wrap the ``w:drawing``."""
        self._drawing = drawing
        self._parent = parent

    @property
    def element(self) -> Any:
        """The typed ``w:drawing``."""
        return self._drawing

    @property
    def part(self) -> Any:
        """The part this drawing is in."""
        return None if self._parent is None else self._parent.part

    @property
    def has_picture(self) -> bool:
        """True when the drawing holds an embedded picture rather than a chart.

        False for a **linked** picture, as python-docx's is: the bytes are on
        the filesystem, not in the package.
        """
        inline = inline_of(self._drawing)
        pic = _pic_of(inline) if inline is not None else None
        if pic is None:
            return False
        blip = getattr(getattr(pic, "blip_fill", None), "blip", None)
        return blip is not None and getattr(blip, "embed", None) is not None


def inline_of(drawing: Any) -> Any | None:
    """The ``wp:inline`` of a ``w:drawing``, or None for a floating shape.

    The content API's ``_inline_of``, which is the same walk over the one
    ``anchor_or_inline`` choice list the generated ``Drawing`` has.
    """
    from docx4j_py.model.content.picture import _inline_of

    return _inline_of(drawing)


def _straight_in_a_branch(drawing: Any) -> bool:
    """True for a ``w:drawing`` whose parent is an ``mc:Choice`` or ``mc:Fallback``.

    python-docx's inline shapes are ``//w:p/w:r/w:drawing/wp:inline``: a
    drawing straight in a run's ``mc:AlternateContent`` branch is not a run's
    child, so python-docx does not list it. Before CR-012 Phase B such a
    drawing was a wildcard tree that no walk of typed nodes visited; now it
    is a typed ``Drawing``, and this keeps it out, as python-docx does.
    """
    parent = getattr(drawing, "parent", None)
    cls = type(parent)
    return cls.__module__ == "docx4j_py.mce" and cls.__name__ in (
        "AlternateContentChoice",
        "AlternateContentFallback",
    )


def _pic_of(inline: Any) -> Any | None:
    """The ``pic:pic`` inside a ``wp:inline``'s graphic data, or None."""
    graphic = getattr(inline, "graphic", None)
    data = getattr(graphic, "graphic_data", None) if graphic is not None else None
    for item in getattr(data, "any_element", None) or ():
        if type(item).__name__ == "Pic":
            return item
    return None


class InlineShape:
    """A ``wp:inline``: python-docx's ``InlineShape``."""

    __slots__ = ("_inline", "_story")

    def __init__(self, inline: Any, story: Story | None = None) -> None:
        """Wrap the ``wp:inline``."""
        self._inline = inline
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one ``wp:inline`` are equal."""
        return isinstance(other, InlineShape) and other._inline is self._inline

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._inline))

    def __repr__(self) -> str:
        """``<docx4j_py.docx InlineShape PICTURE 1828800x1371600>``."""
        return (
            f"<docx4j_py.docx InlineShape {self.type.name} "
            f"{int(self.width)}x{int(self.height)}>"
        )

    @property
    def element(self) -> Any:
        """The typed ``wp:inline``; python-docx spells it ``_inline``."""
        return self._inline

    @property
    def width(self) -> Length:
        """The display width in EMU."""
        return Emu(int(self._inline.extent.cx))

    @width.setter
    def width(self, cx: int | Length) -> None:
        self._set(cx=int(cx))

    @property
    def height(self) -> Length:
        """The display height in EMU."""
        return Emu(int(self._inline.extent.cy))

    @height.setter
    def height(self, cy: int | Length) -> None:
        self._set(cy=int(cy))

    def _set(self, *, cx: int | None = None, cy: int | None = None) -> None:
        """Write one or both axes on the extent and on the picture's ``a:ext``."""
        with _maybe_recording(self._story, "docx.inline_shape.size"):
            extent = self._inline.extent
            if cx is not None:
                extent.cx = cx
            if cy is not None:
                extent.cy = cy
            pic = _pic_of(self._inline)
            xfrm = getattr(getattr(pic, "sp_pr", None), "xfrm", None) if pic is not None else None
            ext = getattr(xfrm, "ext", None) if xfrm is not None else None
            if ext is not None:
                if cx is not None:
                    ext.cx = cx
                if cy is not None:
                    ext.cy = cy

    @property
    def type(self) -> WD_INLINE_SHAPE:
        """What the shape holds; ``PICTURE`` for everything this engine writes."""
        graphic = getattr(self._inline, "graphic", None)
        data = getattr(graphic, "graphic_data", None) if graphic is not None else None
        uri = getattr(data, "uri", None)
        if uri == _URI_PICTURE:
            pic = _pic_of(self._inline)
            blip = getattr(getattr(pic, "blip_fill", None), "blip", None)
            if blip is not None and getattr(blip, "link", None) is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == _URI_CHART:
            return WD_INLINE_SHAPE.CHART
        if uri == _URI_DIAGRAM:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED


class _maybe_recording:
    """A recording block when there is a story to record against."""

    __slots__ = ("inner",)

    def __init__(self, story: Story | None, operation: str) -> None:
        """Open a report if a body is reachable."""
        self.inner = None if story is None else recording(story.body, operation)

    def __enter__(self) -> Any:
        """The recorder, or None."""
        return None if self.inner is None else self.inner.__enter__()

    def __exit__(self, *exc_info: object) -> None:
        """Close the report, if there was one."""
        if self.inner is not None:
            self.inner.__exit__(*exc_info)


class InlineShapes:
    """``document.inline_shapes``: every ``w:p/w:r/w:drawing/wp:inline``."""

    __slots__ = ("_story",)

    def __init__(self, story: Story) -> None:
        """Wrap the part whose inline shapes these are."""
        self._story = story

    @property
    def part(self) -> Any:
        """The part these shapes are in."""
        return self._story.part

    def __repr__(self) -> str:
        """``<docx4j_py.docx InlineShapes, 2 shapes>``."""
        return f"<docx4j_py.docx InlineShapes, {len(self)} shapes>"

    @property
    def _inlines(self) -> list[Any]:
        from docx4j_py.traversal import element_name, iter_nodes

        w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"
        out: list[Any] = []
        for node in iter_nodes(self._story.body.container):
            if element_name(node) == w and not _straight_in_a_branch(node):
                inline = inline_of(node)
                if inline is not None:
                    out.append(inline)
        return out

    def __len__(self) -> int:
        """How many inline shapes the part holds."""
        return len(self._inlines)

    def __iter__(self) -> Iterator[InlineShape]:
        """Each shape, in document order."""
        return (InlineShape(inline, self._story) for inline in self._inlines)

    def __getitem__(self, idx: int) -> InlineShape:
        """The shape at `idx`."""
        try:
            return InlineShape(self._inlines[idx], self._story)
        except IndexError:
            raise IndexError(f"inline shape index [{idx}] out of range") from None
