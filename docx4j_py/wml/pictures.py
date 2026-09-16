"""Inline pictures: the ``w:drawing`` docx4j writes, and the image headers.

CR-003 Phase A (section 3.3), the tree half of "add a picture". Three pieces,
none of which knows about parts, relationships or files:

:func:`inline_picture`
    the ``w:drawing`` holding one ``wp:inline``, element for element and
    attribute for attribute as docx4j's
    ``BinaryPartAbstractImage.createImageInline`` writes it. The image part and
    its relationship belong to the package, so the relationship id is a plain
    string here; CR-003 Phase C's ``insert_inline_picture`` makes the part and
    then calls this.
:func:`image_size`
    the pixel size and the density of a PNG, JPEG, GIF or BMP, read from the
    header. CR-003 section 4: **no Pillow** --- the four readers are a hundred
    lines, and an image whose bytes this cannot read is an error that says so
    rather than a wrong size.
:func:`emu_for`
    those pixels as English Metric Units, scaled down to a maximum width the
    way docx4j's ``CxCy.scale`` scales an image wider than the text column.

    >>> from docx4j_py.wml import el, r, inline_picture, image_size, emu_for
    >>> info = image_size(png_bytes)  # doctest: +SKIP
    >>> size = emu_for(info, max_width_emu=5486400)  # doctest: +SKIP
    >>> run = r(inline_picture(rel.id, cx=size.cx, cy=size.cy, name="logo.png"))
    ... # doctest: +SKIP
"""

from __future__ import annotations

import dataclasses
import struct
from typing import Any, Literal

from docx4j_py.child import ChildList, link_parents
from docx4j_py.dml import main as _a
from docx4j_py.dml import picture as _pic
from docx4j_py.dml import wordprocessing_drawing as _wp
from docx4j_py.dml.main import el as a_el
from docx4j_py.dml.picture import el as pic_el
from docx4j_py.dml.wordprocessing_drawing import el as wp_el
from docx4j_py.wml import Drawing, el
from docx4j_py.wml.builders import BuilderError

__all__ = [
    "DEFAULT_DPI",
    "EMU_PER_INCH",
    "PICTURE_URI",
    "EmuSize",
    "ImageInfo",
    "emu_for",
    "image_size",
    "inline_picture",
]

#: English Metric Units per inch: docx4j's ``UnitsOfMeasurement`` EMU_RATIO.
EMU_PER_INCH = 914400

#: The density assumed for an image whose header does not give one. docx4j's
#: configured default (``docx4j.dpi``) and Word's own assumption.
DEFAULT_DPI = 96.0

#: The ``a:graphicData/@uri`` of a DrawingML picture.
PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"

#: The image formats :func:`image_size` reads, which are the four Word takes
#: without conversion (docx4j converts anything else with ImageMagick).
ImageFormat = Literal["png", "jpeg", "gif", "bmp"]


@dataclasses.dataclass(frozen=True, slots=True)
class ImageInfo:
    """What an image's header says about its size (docx4j's ``ImageInfo``)."""

    width_px: int
    """The width in pixels."""
    height_px: int
    """The height in pixels."""
    dpi_x: float
    """The horizontal density; :data:`DEFAULT_DPI` when the header has none."""
    dpi_y: float
    """The vertical density; :data:`DEFAULT_DPI` when the header has none."""
    format: ImageFormat
    """``"png"``, ``"jpeg"``, ``"gif"`` or ``"bmp"``."""
    has_density: bool = False
    """True when the density was read rather than assumed."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, for a tool result."""
        return dataclasses.asdict(self)


@dataclasses.dataclass(frozen=True, slots=True)
class EmuSize:
    """An extent in English Metric Units (docx4j's ``CxCy``)."""

    cx: int
    """The width in EMU."""
    cy: int
    """The height in EMU."""
    scaled: bool = False
    """True when the image had to be scaled down to fit."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, for a tool result."""
        return dataclasses.asdict(self)


# ---------------------------------------------------------------------------
# the headers
# ---------------------------------------------------------------------------

_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
_METRES_PER_INCH = 0.0254

#: The JPEG start-of-frame markers, which carry the size. Every ``SOF`` except
#: ``C4`` (define Huffman table), ``C8`` (JPEG extension) and ``CC`` (define
#: arithmetic coding), which share the range but are not frames.
_JPEG_SOF = frozenset(
    {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}
)


def _png_size(data: bytes) -> ImageInfo:
    """PNG: ``IHDR`` for the pixels, ``pHYs`` for the density."""
    width, height = struct.unpack_from(">II", data, 16)
    dpi_x = dpi_y = DEFAULT_DPI
    has_density = False
    offset = 8
    while offset + 8 <= len(data):
        length, kind = struct.unpack_from(">I4s", data, offset)
        if kind == b"pHYs" and offset + 8 + 9 <= len(data):
            ppu_x, ppu_y, unit = struct.unpack_from(">IIB", data, offset + 8)
            # unit 1 is the metre; unit 0 means "aspect ratio only", no density
            if unit == 1 and ppu_x and ppu_y:
                dpi_x = ppu_x * _METRES_PER_INCH
                dpi_y = ppu_y * _METRES_PER_INCH
                has_density = True
            break
        if kind == b"IDAT" or kind == b"IEND":
            break
        offset += 12 + length  # length, type, data, CRC
    return ImageInfo(width, height, dpi_x, dpi_y, "png", has_density)


def _jpeg_size(data: bytes) -> ImageInfo:
    """JPEG: the first ``SOFn`` for the pixels, the ``JFIF`` ``APP0`` density."""
    dpi_x = dpi_y = DEFAULT_DPI
    has_density = False
    offset = 2
    while offset + 4 <= len(data):
        if data[offset] != 0xFF:
            offset += 1
            continue
        marker = data[offset + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:  # standalone
            offset += 2
            continue
        length = struct.unpack_from(">H", data, offset + 2)[0]
        segment = offset + 4
        if (
            marker == 0xE0
            and data[segment : segment + 5] == b"JFIF\x00"
            and segment + 12 <= len(data)
        ):
            units, x_density, y_density = struct.unpack_from(">BHH", data, segment + 7)
            if x_density and y_density:
                if units == 1:  # dots per inch
                    dpi_x, dpi_y = float(x_density), float(y_density)
                    has_density = True
                elif units == 2:  # dots per centimetre
                    dpi_x, dpi_y = x_density * 2.54, y_density * 2.54
                    has_density = True
        elif marker in _JPEG_SOF:
            height, width = struct.unpack_from(">HH", data, segment + 1)
            return ImageInfo(width, height, dpi_x, dpi_y, "jpeg", has_density)
        elif marker == 0xDA:  # start of scan: no frame header before the data
            break
        offset += 2 + length
    raise BuilderError(
        "no JPEG start-of-frame marker: the bytes are not a whole JPEG",
        code="image.truncated",
        hint="pass the complete image part, not the first few hundred bytes",
    )


def _gif_size(data: bytes) -> ImageInfo:
    """GIF: the logical screen descriptor. GIF carries no density."""
    width, height = struct.unpack_from("<HH", data, 6)
    return ImageInfo(width, height, DEFAULT_DPI, DEFAULT_DPI, "gif", False)


def _bmp_size(data: bytes) -> ImageInfo:
    """BMP: the DIB header, whose density is in pixels per metre."""
    header_size = struct.unpack_from("<I", data, 14)[0]
    if header_size == 12:  # BITMAPCOREHEADER: 16-bit, unsigned, no density
        width, height = struct.unpack_from("<HH", data, 18)
        return ImageInfo(width, height, DEFAULT_DPI, DEFAULT_DPI, "bmp", False)
    width, height = struct.unpack_from("<ii", data, 18)
    dpi_x = dpi_y = DEFAULT_DPI
    has_density = False
    if header_size >= 40 and len(data) >= 46:
        ppm_x, ppm_y = struct.unpack_from("<ii", data, 38)
        if ppm_x > 0 and ppm_y > 0:
            dpi_x = ppm_x * _METRES_PER_INCH
            dpi_y = ppm_y * _METRES_PER_INCH
            has_density = True
    # a negative height means the rows are stored top down; the size is |height|
    return ImageInfo(width, abs(height), dpi_x, dpi_y, "bmp", has_density)


def image_size(data: bytes) -> ImageInfo:
    """The pixel size and the density of an image, from its header.

    PNG (``IHDR``, ``pHYs``), JPEG (``SOFn``, the ``JFIF`` ``APP0`` density),
    GIF and BMP, which are the four formats Word embeds without converting.
    Where the header gives no density :data:`DEFAULT_DPI` is assumed, as docx4j
    and Word do, and :attr:`ImageInfo.has_density` says which happened.

    Raises:
        BuilderError: if the bytes are not one of the four formats, or are
            truncated before the size.
    """
    if data.startswith(_PNG_SIGNATURE) and len(data) >= 24:
        return _png_size(data)
    if data.startswith(b"\xff\xd8"):
        return _jpeg_size(data)
    if data.startswith((b"GIF87a", b"GIF89a")) and len(data) >= 10:
        return _gif_size(data)
    if data.startswith(b"BM") and len(data) >= 26:
        return _bmp_size(data)
    raise BuilderError(
        f"not a PNG, JPEG, GIF or BMP image (first bytes {data[:8]!r})",
        code="image.unsupported_format",
        hint="convert the image to PNG or JPEG first; Word embeds only these four",
    )


def emu_for(info: ImageInfo, *, max_width_emu: int | None = None) -> EmuSize:
    """An image's natural size in EMU, scaled down to `max_width_emu`.

    ``px / dpi * 914400`` per axis, which is docx4j's ``CxCy.scale``: the
    natural size when it fits, and otherwise the maximum width with the height
    reduced in proportion, so the aspect ratio is kept (``noChangeAspect`` in
    the ``wp:inline`` says the same thing to Word). `max_width_emu` is the
    width of the text column, which the caller knows from ``w:sectPr`` and this
    module does not.
    """
    cx = info.width_px / info.dpi_x * EMU_PER_INCH
    cy = info.height_px / info.dpi_y * EMU_PER_INCH
    if max_width_emu is not None and cx > max_width_emu:
        return EmuSize(int(max_width_emu), round(cy * max_width_emu / cx), True)
    return EmuSize(round(cx), round(cy), False)


# ---------------------------------------------------------------------------
# the drawing
# ---------------------------------------------------------------------------


def inline_picture(
    rel_id: str,
    *,
    cx: int,
    cy: int,
    id: int = 1,  # noqa: A002 - wp:docPr's attribute name
    name: str | None = None,
    descr: str = "",
    title: str | None = None,
    pic_id: int = 0,
    pic_name: str | None = None,
    link: bool = False,
) -> Drawing:
    """A ``w:drawing`` holding one ``wp:inline`` for an image part.

    Element for element what docx4j's
    ``BinaryPartAbstractImage.createImageInline`` writes: ``wp:extent``, a zero
    ``wp:effectExtent``, ``wp:docPr``, a ``wp:cNvGraphicFramePr`` whose
    ``a:graphicFrameLocks`` sets ``noChangeAspect``, and an ``a:graphic`` whose
    ``a:graphicData`` holds a ``pic:pic`` with the blip, a stretch fill and a
    ``rect`` preset geometry.

    Returns the :class:`Drawing`, so the call reads
    ``r(inline_picture(rel.id, cx=…, cy=…))``.

    Args:
        rel_id: the relationship id of the image part, which becomes
            ``a:blip/@r:embed`` (or ``@r:link``; see `link`). The part and the
            relationship are the package's business, not this module's.
        cx: the width in EMU (:func:`emu_for` gives one from the image).
        cy: the height in EMU.
        id: ``wp:docPr/@id``, unique in the document. docx4j's ``id1``.
        name: ``wp:docPr/@name`` and ``pic:cNvPr/@name``, what Word shows in
            the selection pane; docx4j's ``filenameHint``, usually the original
            file name. Defaults to ``"Picture <id>"``, as Word names one.
        descr: ``wp:docPr/@descr``, Office JS's ``altTextDescription``;
            docx4j's ``altText``. Empty by default, as docx4j writes it.
        title: ``wp:docPr/@title``, Office JS's ``altTextTitle``. Absent by
            default, and absent from what docx4j writes; the schema copy gained
            the attribute for this (``schemas/PATCHES.md`` 5).
        pic_id: ``pic:cNvPr/@id``, docx4j's ``id2``, a second id Word does not
            show. 0 by default, which is what Word itself writes.
        pic_name: ``pic:cNvPr/@name``. docx4j writes its one ``filenameHint``
            in both places, which is the default here too; Word writes the file
            name here and ``Picture N`` above, so the two can differ.
        link: write ``r:link`` rather than ``r:embed``, for an image that is
            linked rather than embedded (docx4j's ``link``).

    Raises:
        BuilderError: if `cx` or `cy` is not a positive number of EMU.
    """
    if cx <= 0 or cy <= 0:
        raise BuilderError(
            f"a picture needs a positive extent, not cx={cx} cy={cy}",
            code="picture.empty_extent",
            hint="use emu_for(image_size(data)) to size it from the image itself",
        )
    if name is None:
        name = f"Picture {id}"
    if pic_name is None:
        pic_name = name

    def doc_props(props_id: int, props_name: str, *, described: bool) -> Any:
        props = _a.CTNonVisualDrawingProps(id=props_id, name=props_name)
        if described:
            props.descr = descr
            if title is not None:
                props.title = title
        return props

    picture = pic_el.pic(
        nv_pic_pr=pic_el.nvPicPr(
            c_nv_pr=doc_props(pic_id, pic_name, described=False),
            c_nv_pic_pr=pic_el.cNvPicPr(),
        ),
        blip_fill=pic_el.blipFill(
            blip=a_el.blip_blip(link=rel_id) if link else a_el.blip_blip(embed=rel_id),
            tile_or_stretch=ChildList([a_el.stretch(fill_rect=a_el.fillRect())]),
        ),
        sp_pr=pic_el.spPr(
            xfrm=a_el.xfrm(
                off=a_el.off(x=0, y=0),
                ext=a_el.positive_size2_d(cx=cx, cy=cy),
            ),
            cust_geom_or_prst_geom=ChildList([a_el.prstGeom(prst="rect", av_lst=a_el.avLst())]),
        ),
    )

    inline = wp_el.inline(
        dist_t=0,
        dist_b=0,
        dist_l=0,
        dist_r=0,
        extent=wp_el.extent(cx=cx, cy=cy),
        effect_extent=wp_el.effectExtent(l=0, t=0, r=0, b=0),
        doc_pr=doc_props(id, name, described=True),
        c_nv_graphic_frame_pr=wp_el.cNvGraphicFramePr(
            graphic_frame_locks=a_el.graphicFrameLocks(no_change_aspect=True)
        ),
        graphic=a_el.graphic(
            graphic_data=a_el.graphicData(uri=PICTURE_URI, any_element=ChildList([picture]))
        ),
    )
    drawing = el.drawing(anchor_or_inline=ChildList([inline]))
    link_parents(drawing)
    return drawing


# ``_pic`` and ``_wp`` are imported for the module's type names to resolve in
# documentation and for a reader to see which packages this builds from.
_ = (_pic.Pic, _wp.Inline)
