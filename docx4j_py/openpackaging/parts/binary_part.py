"""Parts whose content is bytes. docx4j ``BinaryPart`` and its subclasses.

CR-002 section 5.4. One class for images rather than docx4j's nine
(``ImagePngPart``, ``ImageJpegPart``, ``MetafileEmfPart``, ...): the kind of an
image is its content type, and a Python consumer switches on that string more
naturally than on a class. The departure is recorded in CR-002 section 12; the
TypeScript engine made the same one.
"""

from __future__ import annotations

from docx4j_py.openpackaging.content_types import ContentTypes, is_stored_uncompressed
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part

__all__ = [
    "AlternativeFormatInputPart",
    "BinaryPart",
    "EmbeddedPackagePart",
    "ImagePart",
    "ObfuscatedFontPart",
    "OleObjectBinaryPart",
    "TrueTypeFontPart",
    "VbaProjectBinaryPart",
]


class BinaryPart(Part):
    """A part held as bytes: images, fonts, embeddings, anything unrecognised."""

    __slots__ = ("_bytes",)

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = ContentTypes.OCTET_STREAM,
        relationship_type: str = "",
    ) -> None:
        """Build the part; its bytes come from the container until they are set."""
        super().__init__(part_name, content_type, relationship_type)
        self._bytes: bytes | None = None

    @property
    def is_loaded(self) -> bool:
        """Whether the bytes have been read or set."""
        return self._bytes is not None

    @property
    def compress(self) -> bool:
        """False for already-compressed media: a zip stores those, as Word does."""
        return not is_stored_uncompressed(self.content_type)

    @property
    def data(self) -> bytes:
        """The bytes, read from the source container on first access."""
        if self._bytes is None:
            loaded = self._source_bytes()
            if loaded is None:
                raise Docx4JException(
                    f"Part {self.part_name} has no content and no source container"
                )
            self._bytes = loaded
        return self._bytes

    def set_bytes(self, data: bytes) -> None:
        """Replace the bytes."""
        self._bytes = bytes(data)

    @property
    def bytes_for_save(self) -> bytes:
        """:attr:`data`."""
        return self.data

    def __len__(self) -> int:
        """The number of bytes, without loading when the container knows."""
        if self._bytes is not None:
            return len(self._bytes)
        package = self.package
        store = package.source_part_store if package is not None else None
        if store is not None:
            size = store.size(self.part_name.store_name)
            if size is not None:
                return size
        return len(self.data)


class ImagePart(BinaryPart):
    """An image of any kind. docx4j ``BinaryPartAbstractImage`` and its subclasses.

    The kind is :attr:`content_type` (``image/png``, ``image/x-emf``, ...).
    """

    __slots__ = ()

    def __init__(
        self, part_name: PartName | str, content_type: str = ContentTypes.IMAGE_PNG
    ) -> None:
        """Build an image part with the ``image`` relationship type."""
        super().__init__(part_name, content_type, Namespaces.IMAGE)


class EmbeddedPackagePart(BinaryPart):
    """An embedded Office package: the workbook behind a chart, say. docx4j ``EmbeddedPackagePart``."""

    __slots__ = ()

    def __init__(
        self, part_name: PartName | str, content_type: str = ContentTypes.OCTET_STREAM
    ) -> None:
        """Build the part with the ``package`` relationship type."""
        super().__init__(part_name, content_type, Namespaces.EMBEDDED_PKG)

    @property
    def compress(self) -> bool:
        """False: it is a zip already. docx4j ``shouldCompress``."""
        return False


class OleObjectBinaryPart(BinaryPart):
    """An OLE object's binary. docx4j ``OleObjectBinaryPart``."""

    __slots__ = ()

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = ContentTypes.OFFICEDOCUMENT_OLE_OBJECT,
    ) -> None:
        """Build the part with the ``oleObject`` relationship type."""
        super().__init__(part_name, content_type, Namespaces.OLE_OBJECT)


class ObfuscatedFontPart(BinaryPart):
    """An embedded, obfuscated font. docx4j ``ObfuscatedFontPart``."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the part with the ``font`` relationship type."""
        super().__init__(part_name, ContentTypes.OFFICEDOCUMENT_FONT, Namespaces.FONT)


class TrueTypeFontPart(BinaryPart):
    """An embedded TrueType font. docx4j ``TrueTypeFontPart``."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the part with the ``font`` relationship type."""
        super().__init__(part_name, ContentTypes.TRUETYPE_FONT, Namespaces.FONT)


class VbaProjectBinaryPart(BinaryPart):
    """``vbaProject.bin``. docx4j ``VbaProjectBinaryPart``."""

    __slots__ = ()

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = ContentTypes.OFFICEDOCUMENT_VBA_PROJECT,
    ) -> None:
        """Build the part with the ``vbaProject`` relationship type."""
        super().__init__(part_name, content_type, Namespaces.VBA_PROJECT)


class AlternativeFormatInputPart(BinaryPart):
    """An ``altChunk``'s content: HTML, RTF, text, or another docx.

    docx4j ``AlternativeFormatInputPart``. A WordprocessingML chunk is a zip
    already and is stored rather than deflated, as docx4j's ``shouldCompress``
    decides.
    """

    __slots__ = ()

    def __init__(
        self, part_name: PartName | str, content_type: str = ContentTypes.OCTET_STREAM
    ) -> None:
        """Build the part with the ``aFChunk`` relationship type."""
        super().__init__(part_name, content_type, Namespaces.AF)

    @property
    def compress(self) -> bool:
        """False for a WordprocessingML chunk, which is itself a zip."""
        ct = self.content_type or ""
        return not ("wordprocessingml" in ct or "ms-word" in ct)
