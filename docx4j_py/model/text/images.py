"""A picture as the document gives it, and its description (WD REQ-514). CR-011 section 6.

docx4j's ``TextImage``, ``ImageDescription``, ``ImageDescriber`` and
``AltTextDescriber`` (its CR-046 phase 2). One describer per reading,
:func:`alt_text_describer` by default; a host may replace or wrap it, a
captioning model for instance, through ``TextOptions(image_describer=...)``,
and what it generates it marks ``GENERATED``, which plain text writes
``[image, generated: ...]``.

A describer reads what the document says of a picture, and may read its part;
an SVG's ``<text>`` elements and the glyph runs of EMF and WMF are not the
document's text and the default never reads them (REQ-514).
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any, Literal

__all__ = [
    "ImageDescriber",
    "ImageDescription",
    "TextImage",
    "alt_text_describer",
]

#: Where a description came from (docx4j ``ImageDescription.Provenance``).
Provenance = Literal["AUTHOR", "GENERATED"]


@dataclass(frozen=True, slots=True)
class ImageDescription:
    """An image's description and where it came from. docx4j ``ImageDescription``.

    ``AUTHOR`` for the document's own text (``wp:docPr/@descr`` or ``@name``,
    VML ``@alt``), ``GENERATED`` for a describer's.
    """

    text: str
    provenance: Provenance = "AUTHOR"

    @classmethod
    def author(cls, text: str) -> ImageDescription:
        """An author's description."""
        return cls(text, "AUTHOR")

    @classmethod
    def generated(cls, text: str) -> ImageDescription:
        """A generated description."""
        return cls(text, "GENERATED")

    def __str__(self) -> str:
        """The text, ``generated: `` before a generated one."""
        return f"generated: {self.text}" if self.provenance == "GENERATED" else self.text


@dataclass(frozen=True, slots=True, eq=False)
class TextImage:
    """A picture as the document gives it, for an :data:`ImageDescriber`. docx4j ``TextImage``.

    The title is never read as text (REQ-514's note on ``@title``); it is here
    for a describer that wants it.
    """

    #: The alt text: ``wp:docPr/@descr`` (or a nested picture's
    #: ``pic:cNvPr/@descr``), VML ``v:shape/@alt``; None for none.
    descr: str | None
    #: ``wp:docPr/@name`` (``pic:cNvPr/@name``); None for a VML picture.
    name: str | None
    #: ``wp:docPr/@title``; None for none.
    title: str | None
    #: Marked decorative (``adec:decorative val="1"`` in ``wp:docPr``'s extensions).
    decorative: bool
    #: A VML picture (``v:shape`` with ``v:imagedata``).
    vml: bool
    #: The image part (``a:blip/@r:embed``, else its ``asvg:svgBlip``'s;
    #: ``v:imagedata/@r:id``), or None for a linked or missing one.
    part: Any = field(default=None, repr=False)
    #: What the picture was read from: a ``wp:inline``, ``wp:anchor``, ``pic:pic`` or ``v:shape``.
    source: Any = field(default=None, repr=False)


#: Gives an image its description, or None for none (docx4j ``ImageDescriber``).
ImageDescriber = Callable[[TextImage], "ImageDescription | None"]


def alt_text_describer(image: TextImage | None) -> ImageDescription | None:
    """The default describer (docx4j ``AltTextDescriber``): the author's alt text, else the name.

    ``wp:docPr/@descr``, else the picture's ``@name``; for VML,
    ``v:shape/@alt``; nothing for a picture marked decorative. Every
    description it gives is the author's.
    """
    if image is None or image.decorative:
        return None
    if image.descr:
        return ImageDescription.author(image.descr)
    if image.name:
        return ImageDescription.author(image.name)
    return None
