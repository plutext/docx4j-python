"""``InlinePicture``: Office JS ``Word.InlinePicture`` over a ``w:drawing``.

CR-003 section 3.2, Phase C. The counterpart of docx4j's
``BinaryPartAbstractImage.createImagePart`` followed by ``createImageInline``:
an :class:`~docx4j_py.openpackaging.parts.binary_part.ImagePart` under a free
``/word/media/imageN.<ext>``, a relationship from the part the body belongs to
(the main document, a header, a footer), and the ``wp:inline`` CR-003 Phase A's
:func:`~docx4j_py.wml.pictures.inline_picture` writes, sized from the image's
own header (:func:`~docx4j_py.wml.pictures.image_size`) and scaled down to the
text column as docx4j's ``CxCy.scale`` does.

No Pillow: the four header readers are Phase A's, in
:mod:`docx4j_py.wml.pictures`, and an image none of them understands raises a
:class:`~docx4j_py.model.content.errors.BuilderError` naming the four formats.

**What a dry run does with the part it adds** (CR-003 section 12.5): the
``ImagePart`` and its relationship go into the *real* package, because a
:class:`~docx4j_py.model.content.trial.TrialPackage` shares its part map. Phase
C therefore gives the trial an undo log --- :func:`note_added_part` tells it
what was added, and leaving the ``with`` block removes the part, its
relationship and its content-type override again --- so a dry run of
``insert_inline_picture`` really does leave the document as it was. Everything
else about a trial is unchanged.

Sizes are **points** on the way in and out, as Office JS reports them
(``EMU = points * 12700``).
"""

from __future__ import annotations

import base64
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, link_parents
from docx4j_py.model.content.enums import BodyLocation, TextLocation
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.reports import recording
from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.binary_part import ImagePart
from docx4j_py.traversal import element_name
from docx4j_py.wml import R, el, emu_for, image_size, inline_picture, to_xml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = [
    "EMU_PER_POINT",
    "EMU_PER_TWIP",
    "IMAGE_FORMATS",
    "InlinePicture",
    "NewPicture",
    "add_image",
    "free_image_name",
    "insert_picture_into_body",
    "insert_picture_into_paragraph",
    "next_drawing_id",
    "note_added_part",
    "pictures_of",
    "writable_width_emu",
]

#: English Metric Units per point and per twip (docx4j ``UnitsOfMeasurement``).
EMU_PER_POINT = 12700
EMU_PER_TWIP = 635

#: Office JS ``Word.ImageFormat``, by the content type of the part embedded.
IMAGE_FORMATS: dict[str, str] = {
    ContentTypes.IMAGE_PNG: "Png",
    ContentTypes.IMAGE_X_PNG: "Png",
    ContentTypes.IMAGE_JPEG: "Jpeg",
    ContentTypes.IMAGE_GIF: "Gif",
    ContentTypes.IMAGE_BMP: "Bmp",
    ContentTypes.IMAGE_TIFF: "Tiff",
    ContentTypes.IMAGE_SVG: "Svg",
    ContentTypes.IMAGE_EMF: "Emf",
    ContentTypes.IMAGE_EMF2: "Emf",
    ContentTypes.IMAGE_WMF: "Wmf",
}

#: The content type of each format :func:`~docx4j_py.wml.pictures.image_size`
#: reads; the file extension is the format's own name, as docx4j's part names.
_CONTENT_TYPE_BY_FORMAT: dict[str, str] = {
    "png": ContentTypes.IMAGE_PNG,
    "jpeg": ContentTypes.IMAGE_JPEG,
    "gif": ContentTypes.IMAGE_GIF,
    "bmp": ContentTypes.IMAGE_BMP,
}

_W_DRAWING = element_name(el.drawing())
_INLINE_TYPE = "Inline"


# ---------------------------------------------------------------------------
# adding the part (docx4j BinaryPartAbstractImage.createImagePart)
# ---------------------------------------------------------------------------


def free_image_name(package: Any, extension: str) -> PartName:
    """``/word/media/imageN.<ext>``, the first N free (docx4j ``getNewPartName``).

    Deterministic by construction --- it is derived from the parts the package
    already holds --- which is what CR-003 section 3.4 asks of an allocated
    name, and more useful than a number from ``pkg.id_generator()`` would be.
    """
    index = 1
    while True:
        candidate = PartName.of(f"/word/media/image{index}.{extension}")
        if package.get_part(candidate) is None:
            return candidate
        index += 1


def note_added_part(
    package: Any,
    part: Any,
    relationship: Any,
    source: Any,
    *,
    added_content_type: bool = True,
) -> None:
    """Tell a trial about a part added during it, so it can be un-added.

    A no-op on a real package. CR-003 section 12.5 required anything that adds
    a part to *say* what a dry run does with it; giving
    :class:`~docx4j_py.model.content.trial.TrialPackage` an undo log was
    cheaper than a parts overlay and makes the answer "nothing".

    Args:
        package: the body's package --- a
            :class:`~docx4j_py.model.content.trial.TrialPackage` or the real one.
        part: the part that was added.
        relationship: the relationship that was written for it.
        source: the part the relationship lives on.
        added_content_type: whether the content-type override was added by this
            call, and so should go with it.
    """
    note = getattr(package, "note_added_part", None)
    if note is not None:
        note(part, relationship, source, added_content_type=added_content_type)


def _source_part_of(body: Body) -> Any:
    """The part a new relationship is added to: the body's own part."""
    part = getattr(body, "part", None)
    if part is None:
        raise ContentError(
            "this body has no part, so an image cannot be related to it",
            code="picture.no_part",
            hint="insert the picture into a body reached from a package",
        )
    package = getattr(part, "package", None)
    if package is None:
        raise ContentError(
            f"{type(part).__name__} is not in a package yet",
            code="picture.no_package",
            hint="add the part to a package before adding an image to it",
        )
    return getattr(part, "_wrapped", part)


def writable_width_emu(body: Any) -> int | None:
    """The width of the text column in EMU, from ``w:sectPr``, or None.

    CR-003 section 10.4: the number ``emu_for(..., max_width_emu=)`` needs, and
    the reason an image wider than the page comes out scaled down rather than
    cropped by Word.
    """
    container = getattr(body, "container", body)
    sect_pr = getattr(container, "sect_pr", None)
    pg_sz = getattr(sect_pr, "pg_sz", None) if sect_pr is not None else None
    width = getattr(pg_sz, "w", None) if pg_sz is not None else None
    if width is None:
        return None
    pg_mar = getattr(sect_pr, "pg_mar", None)
    left = getattr(pg_mar, "left", None) if pg_mar is not None else None
    right = getattr(pg_mar, "right", None) if pg_mar is not None else None
    try:
        twips = int(width) - int(left or 0) - int(right or 0)
    except (TypeError, ValueError):  # pragma: no cover - a malformed w:sectPr
        return None
    return twips * EMU_PER_TWIP if twips > 0 else None


def next_drawing_id(scope: Any) -> int:
    """An id no other drawing in this part uses (``wp:docPr/@id``).

    One above the highest in use, as :func:`~docx4j_py.wml.sdt.next_sdt_id`
    allocates a control's, so the same document and the same calls give the
    same bytes.
    """
    from docx4j_py.dml.main import CTNonVisualDrawingProps
    from docx4j_py.traversal import find

    highest = 0
    for props in find(scope, CTNonVisualDrawingProps, mce="all"):
        value = getattr(props, "id", None)
        if isinstance(value, int) and value > highest:
            highest = value
    return highest + 1


class NewPicture:
    """What :func:`add_image` made: the run, the drawing, the part and the id."""

    __slots__ = ("drawing", "image_part", "rel_id", "run")

    def __init__(self, run: R, drawing: Any, image_part: ImagePart, rel_id: str) -> None:
        """Hold the four pieces the caller has to place."""
        #: The ``w:r`` holding the drawing; the caller inserts it.
        self.run = run
        #: The ``w:drawing``.
        self.drawing = drawing
        #: The ``ImagePart`` added to the package.
        self.image_part = image_part
        #: The relationship id the drawing refers to.
        self.rel_id = rel_id

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the part name and the relationship id."""
        return {"part": str(self.image_part.part_name), "rel_id": self.rel_id}

    def __repr__(self) -> str:
        """``<NewPicture /word/media/image1.png rId5>``."""
        return f"<NewPicture {self.image_part.part_name} {self.rel_id}>"


def add_image(
    body: Body,
    data: bytes,
    *,
    width: float | None = None,
    height: float | None = None,
    alt_text_description: str = "",
    alt_text_title: str | None = None,
    name: str | None = None,
) -> NewPicture:
    """Add the image to the package and build the run that shows it.

    docx4j's ``createImagePart`` then ``createImageInline``, in that order:

    1. the bytes are read for their format and pixel size (no Pillow);
    2. an ``ImagePart`` goes in under the first free
       ``/word/media/imageN.<ext>``;
    3. a relationship is written on the part the body belongs to;
    4. the ``wp:inline`` is built, at the image's natural size unless `width` or
       `height` says otherwise, scaled down to the text column when it is wider.

    Raises:
        BuilderError: the bytes are not a PNG, JPEG, GIF or BMP.
        ContentError: the body has no part, or its part is in no package.
    """
    source = _source_part_of(body)
    package = source.package
    info = image_size(data)
    content_type = _CONTENT_TYPE_BY_FORMAT[info.format]
    part = ImagePart(free_image_name(package, info.format), content_type)
    part.set_bytes(data)
    had_override = package.content_type_manager.get_override_content_type(part.part_name)
    rel = source.add_target_part(part)
    note_added_part(
        getattr(body, "package", None),
        part,
        rel,
        source,
        added_content_type=had_override is None,
    )

    size = emu_for(info, max_width_emu=writable_width_emu(body))
    cx, cy = size.cx, size.cy
    if width is not None:
        wanted = round(width * EMU_PER_POINT)
        cy = round(height * EMU_PER_POINT) if height is not None else max(1, round(cy * wanted / cx))
        cx = wanted
    elif height is not None:
        wanted = round(height * EMU_PER_POINT)
        cx = max(1, round(cx * wanted / cy))
        cy = wanted

    identifier = next_drawing_id(getattr(body, "container", body))
    drawing = inline_picture(
        rel.id,
        cx=cx,
        cy=cy,
        id=identifier,
        name=name or f"Picture {identifier}",
        descr=alt_text_description,
        title=alt_text_title,
        pic_name=part.part_name.name.rpartition("/")[2],
    )
    run = R(content=ChildList([drawing]))
    link_parents(run)
    return NewPicture(run, drawing, part, rel.id)


# ---------------------------------------------------------------------------
# the view
# ---------------------------------------------------------------------------


def _inline_of(drawing: Any) -> Any:
    """The ``wp:inline`` of a ``w:drawing``, or None for a floating shape."""
    for item in getattr(drawing, "anchor_or_inline", None) or ():
        if type(item).__name__ == _INLINE_TYPE:
            return item
    return None


def pictures_of(paragraph: Paragraph) -> list[InlinePicture]:
    """Every inline picture in a paragraph, in order (Office JS ``inlinePictures``)."""
    from docx4j_py.model.content.text_model import runs_of

    out: list[InlinePicture] = []
    for run in runs_of(paragraph.element):
        for item in getattr(run, "content", None) or ():
            if element_name(item) != _W_DRAWING:
                continue
            if _inline_of(item) is not None:
                out.append(InlinePicture(item, run, paragraph))
    return out


class InlinePicture:
    """A subset of Office JS ``Word.InlinePicture`` over a ``w:drawing``."""

    __slots__ = ("element", "paragraph", "run")

    def __init__(self, element: Any, run: R, paragraph: Paragraph) -> None:
        """Build the view over the drawing, the run holding it and its paragraph."""
        #: The ``w:drawing``.
        self.element = element
        #: The ``w:r`` the drawing is in.
        self.run = run
        #: The :class:`~docx4j_py.model.content.Paragraph` the run is in.
        self.paragraph = paragraph

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:drawing`` are equal."""
        return isinstance(other, InlinePicture) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self.element))

    def __repr__(self) -> str:
        """``<InlinePicture body/3 Png 120.0x90.0pt 'a pangolin'>``."""
        alt = self.alt_text_description
        where = self.paragraph.address
        size = f"{self.width:g}x{self.height:g}pt"
        return f"<InlinePicture {where} {self.image_format} {size} {alt!r}>"

    # -- the tree ----------------------------------------------------------

    @property
    def inline(self) -> Any:
        """The ``wp:inline``.

        Raises:
            ContentError: the drawing is a ``wp:anchor``, a floating shape,
                which Office JS does not call an inline picture.
        """
        found = _inline_of(self.element)
        if found is None:
            raise ContentError(
                "this drawing is a floating shape (wp:anchor), not an inline picture",
                code="picture.not_inline",
                hint="Office JS's InlinePicture is wp:inline only",
            )
        return found

    def _pic(self) -> Any:
        """The ``pic:pic`` inside the graphic data, or None."""
        graphic = getattr(self.inline, "graphic", None)
        data = getattr(graphic, "graphic_data", None) if graphic is not None else None
        for item in getattr(data, "any_element", None) or ():
            if type(item).__name__ == "Pic":
                return item
        return None

    @property
    def rel_id(self) -> str | None:
        """The relationship id of the image part (``a:blip/@r:embed``). Extension."""
        pic = self._pic()
        blip = getattr(getattr(pic, "blip_fill", None), "blip", None) if pic is not None else None
        return getattr(blip, "embed", None) or getattr(blip, "link", None)

    @property
    def image_part(self) -> ImagePart | None:
        """The image part this picture embeds, or None. Extension."""
        rel_id = self.rel_id
        part = self.paragraph.parent_body.part
        rels = getattr(part, "relationships_part", None) if part is not None else None
        if not rel_id or rels is None:
            return None
        found = rels.get_part(rel_id)
        return found if isinstance(found, ImagePart) else None

    # -- size --------------------------------------------------------------

    @property
    def width(self) -> float:
        """The width in points, as Office JS reports it."""
        return self.inline.extent.cx / EMU_PER_POINT

    @width.setter
    def width(self, points: float) -> None:
        self._resize(round(points * EMU_PER_POINT), None)

    @property
    def height(self) -> float:
        """The height in points, as Office JS reports it."""
        return self.inline.extent.cy / EMU_PER_POINT

    @height.setter
    def height(self, points: float) -> None:
        self._resize(None, round(points * EMU_PER_POINT))

    def _resize(self, cx: int | None, cy: int | None) -> None:
        """Resize, keeping the aspect ratio for the axis not given."""
        with recording(self.paragraph.parent_body, "format") as change:
            extent = self.inline.extent
            change.text(before=f"size={extent.cx}x{extent.cy}")
            ratio = extent.cy / extent.cx if extent.cx else 1.0
            width = cx if cx is not None else max(1, round((cy or 1) / (ratio or 1)))
            height = cy if cy is not None else max(1, round(width * ratio))
            extent.cx, extent.cy = width, height
            pic = self._pic()
            xfrm = getattr(getattr(pic, "sp_pr", None), "xfrm", None) if pic is not None else None
            ext = getattr(xfrm, "ext", None) if xfrm is not None else None
            if ext is not None:
                ext.cx, ext.cy = width, height
            change.touched(self.paragraph)
            change.text(after=f"size={width}x{height}")

    # -- alt text ----------------------------------------------------------

    @property
    def alt_text_description(self) -> str:
        """``wp:docPr/@descr``, which Word calls the alt text."""
        return self.inline.doc_pr.descr or ""

    @alt_text_description.setter
    def alt_text_description(self, text: str) -> None:
        with recording(self.paragraph.parent_body, "format") as change:
            change.text(before=f"alt_text_description={self.alt_text_description}")
            self.inline.doc_pr.descr = text
            pic = self._pic()
            props = getattr(getattr(pic, "nv_pic_pr", None), "c_nv_pr", None)
            if props is not None:
                props.descr = text
            change.touched(self.paragraph)
            change.text(after=f"alt_text_description={text}")

    @property
    def alt_text_title(self) -> str | None:
        """``wp:docPr/@title``; None when the attribute is absent."""
        return self.inline.doc_pr.title

    @alt_text_title.setter
    def alt_text_title(self, text: str | None) -> None:
        with recording(self.paragraph.parent_body, "format") as change:
            change.text(before=f"alt_text_title={self.alt_text_title}")
            self.inline.doc_pr.title = text
            change.touched(self.paragraph)
            change.text(after=f"alt_text_title={text}")

    # -- the bytes ---------------------------------------------------------

    @property
    def image_format(self) -> str:
        """Office JS ``Word.ImageFormat``: ``"Png"``, ``"Jpeg"``, ..., ``"Unsupported"``."""
        part = self.image_part
        if part is None:
            return "Unsupported"
        return IMAGE_FORMATS.get(part.content_type, "Unsupported")

    def get_bytes(self) -> bytes:
        """The image's bytes. The twin of :meth:`get_base64` (CR-003 section 3.1).

        Raises:
            ContentError: the picture has no image part --- a linked image, or
                one whose relationship is gone.
        """
        part = self.image_part
        if part is None:
            raise ContentError(
                f"this picture has no image part (r:embed is {self.rel_id!r})",
                code="picture.no_image_part",
                hint="a linked image is not embedded; read it from its r:link target",
            )
        return part.data

    def get_base64(self) -> str:
        """The image's bytes as base64 (Office JS ``getBase64ImageSrc``)."""
        return base64.b64encode(self.get_bytes()).decode("ascii")

    # -- editing -----------------------------------------------------------

    def delete(self) -> None:
        """Remove the picture; the run goes too when the picture was all it held."""
        with recording(self.paragraph.parent_body, "delete") as change:
            content = getattr(self.run, "content", None)
            if content is not None:
                for index, item in enumerate(content):
                    if item is self.element:
                        del content[index]
                        break
            if not content:
                self.paragraph._remove_empty_runs()
            change.touched(self.paragraph)
            change.text(after="")

    def get_xml(self) -> str:
        """The ``w:drawing`` as XML, with docx4j's prefixes. Extension."""
        return to_xml(self.element)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a picture."""
        part = self.image_part
        return {
            "kind": "picture",
            "address": self.paragraph.address,
            "width": self.width,
            "height": self.height,
            "image_format": self.image_format,
            "alt_text_description": self.alt_text_description,
            "alt_text_title": self.alt_text_title,
            "rel_id": self.rel_id,
            "part": str(part.part_name) if part is not None else None,
        }


# ---------------------------------------------------------------------------
# the verbs
# ---------------------------------------------------------------------------


def insert_picture_into_body(
    body: Body,
    data: bytes,
    *,
    location: BodyLocation = "End",
    **options: Any,
) -> InlinePicture:
    """What :meth:`Body.insert_inline_picture` is. See its docstring."""
    with recording(body, "insert_inline_picture") as change:
        made = add_image(body, data, **options)
        paragraph = body.insert_paragraph("", location=location)
        paragraph.element.content.append(made.run)
        link_parents(paragraph.element)
        made.run.parent = paragraph.element
        change.touched(paragraph)
        _record_part(change, made)
        return InlinePicture(made.drawing, made.run, paragraph)


def insert_picture_into_paragraph(
    paragraph: Paragraph,
    data: bytes,
    *,
    location: TextLocation = "End",
    **options: Any,
) -> InlinePicture:
    """What :meth:`Paragraph.insert_inline_picture` is. See its docstring."""
    body = paragraph.parent_body
    with recording(body, "insert_inline_picture") as change:
        made = add_image(body, data, **options)
        if location == "Replace":
            paragraph.element.content.clear()
            paragraph.element.content.append(made.run)
            made.run.parent = paragraph.element
            link_parents(paragraph.element)
        elif location in ("Start", "End"):
            paragraph.insert_items_at(0 if location == "Start" else len(paragraph.text), [made.run])
        else:
            raise ContentError(
                f"insert_inline_picture takes Start, End or Replace, not {location!r}",
                code="location.invalid",
                hint="use Body.insert_inline_picture for a picture in a paragraph of its own",
            )
        change.touched(paragraph)
        _record_part(change, made)
        return InlinePicture(made.drawing, made.run, paragraph)


def _record_part(change: Any, made: NewPicture) -> None:
    """Name the image part and the relationships part in ``parts_touched``."""
    parts = getattr(change, "parts", None)
    if parts is None:
        return
    for name in (str(made.image_part.part_name), _rels_name(made.image_part)):
        if name and name not in parts:
            parts.append(name)


def _rels_name(part: ImagePart) -> str | None:
    rels = getattr(part.owning_relationship_part, "part_name", None)
    return str(rels) if rels is not None else None
