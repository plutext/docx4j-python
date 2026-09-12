"""``[Content_Types].xml``: the constants and the manager.

CR-002 section 5.3, docx4j ``org.docx4j.openpackaging.contenttype``. Two element
types, ``Default`` (extension -> content type) and ``Override`` (part name ->
content type); the engine parses and writes them with lxml because the schema
that describes them, ``xsd/contentTypes/opc-contentTypes.xsd``, is not in the
generated set.

:class:`ContentTypes` is docx4j's constant class, name for name and string for
string; every value was transcribed from
``org/docx4j/openpackaging/contenttype/ContentTypes.java``.
"""

from __future__ import annotations

from lxml import etree

from docx4j_py.openpackaging.exceptions import InvalidFormatException
from docx4j_py.openpackaging.part_name import PartName

__all__ = [
    "CONTENT_TYPES_NS",
    "IMAGE_CONTENT_TYPES_BY_EXTENSION",
    "ContentTypeManager",
    "ContentTypes",
    "is_stored_uncompressed",
    "is_xml_content_type",
]

#: The namespace of ``[Content_Types].xml``. docx4j ``ContentTypeManager.TYPES_NAMESPACE_URI``.
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


class ContentTypes:
    """docx4j's ``ContentTypes``, transcribed.

    A plain class of class attributes rather than an enum: docx4j's are
    ``public static final String`` and the values are what travel, never the
    names.
    """

    CONTENT_TYPES_PART = "application/vnd.openxmlformats-package.content-types+xml"
    RELATIONSHIPS_PART = "application/vnd.openxmlformats-package.relationships+xml"
    APPLICATION_XML = "application/xml"
    XML = "text/xml"

    PACKAGE_COREPROPERTIES = "application/vnd.openxmlformats-package.core-properties+xml"
    OFFICEDOCUMENT_CUSTOMPROPERTIES = (
        "application/vnd.openxmlformats-officedocument.custom-properties+xml"
    )
    OFFICEDOCUMENT_EXTENDEDPROPERTIES = (
        "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    )
    OFFICEDOCUMENT_CUSTOMXML_DATASTORAGEPROPERTIES = (
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
    )
    #: Deliberately the same string as :attr:`APPLICATION_XML`, as in docx4j.
    OFFICEDOCUMENT_CUSTOMXML_DATASTORAGE = "application/xml"
    OFFICEDOCUMENT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
    OFFICEDOCUMENT_THEME_OVERRIDE = (
        "application/vnd.openxmlformats-officedocument.themeOverride+xml"
    )
    OFFICEDOCUMENT_FONT = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
    TRUETYPE_FONT = "application/x-font-ttf"
    OFFICEDOCUMENT_VBA_PROJECT = "application/vnd.ms-office.vbaProject"
    OFFICEDOCUMENT_VBA_DATA = "application/vnd.ms-word.vbaData+xml"
    OFFICEDOCUMENT_VBA_PROJECT_SIGNATURE = "application/vnd.ms-office.vbaProjectSignature"
    OFFICEDOCUMENT_OLE_OBJECT = "application/vnd.openxmlformats-officedocument.oleObject"
    OFFICEDOCUMENT_ACTIVEX_OBJECT = "application/vnd.ms-office.activeX"
    OFFICEDOCUMENT_ACTIVEX_XML_OBJECT = "application/vnd.ms-office.activeX+xml"
    WEB_EXTENSION_TASKPANES = "application/vnd.ms-office.webextensiontaskpanes+xml"
    WEB_EXTENSION_WEBEXTENSION = "application/vnd.ms-office.webextension+xml"

    WORDPROCESSINGML_DOCUMENT = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    )
    WORDPROCESSINGML_DOCUMENT_MACROENABLED = (
        "application/vnd.ms-word.document.macroEnabled.main+xml"
    )
    WORDPROCESSINGML_TEMPLATE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
    )
    WORDPROCESSINGML_TEMPLATE_MACROENABLED = (
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
    )
    WORDPROCESSINGML_COMMENTS = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    )
    WORDPROCESSINGML_COMMENTS_EXTENDED = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    )
    WORDPROCESSINGML_COMMENTS_IDS = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
    )
    #: Word 2018; not in docx4j's table, needed for a current Word's comments.
    WORDPROCESSINGML_COMMENTS_EXTENSIBLE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
    )
    WORDPROCESSINGML_ENDNOTES = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
    )
    WORDPROCESSINGML_FONTTABLE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    )
    WORDPROCESSINGML_FOOTER = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    )
    WORDPROCESSINGML_FOOTNOTES = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    )
    WORDPROCESSINGML_GLOSSARYDOCUMENT = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
    )
    WORDPROCESSINGML_HEADER = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    )
    MS_WORD_KEYMAP = "application/vnd.ms-word.keyMapCustomizations+xml"
    WORDPROCESSINGML_NUMBERING = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
    )
    WORDPROCESSINGML_PEOPLE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
    )
    WORDPROCESSINGML_SETTINGS = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
    )
    WORDPROCESSINGML_STYLES = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
    )
    WORDPROCESSINGML_STYLESWITHEFFECTS = "application/vnd.ms-word.stylesWithEffects+xml"
    WORDPROCESSINGML_WEBSETTINGS = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
    )

    DRAWINGML_DIAGRAM_DATA = (
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
    )
    DRAWINGML_DIAGRAM_LAYOUT = (
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
    )
    DRAWINGML_DIAGRAM_COLORS = (
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"
    )
    DRAWINGML_DIAGRAM_STYLE = (
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
    )
    DRAWINGML_DIAGRAM_LAYOUT_HEADER = (
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayoutHeader+xml"
    )
    DRAWINGML_DIAGRAM_DRAWING = "application/vnd.ms-office.drawingml.diagramDrawing+xml"
    DRAWINGML_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    DRAWINGML_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
    DRAWINGML_CHART_SHAPES = (
        "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"
    )
    CHART_EX = "application/vnd.ms-office.chartex+xml"
    CHART_STYLE = "application/vnd.ms-office.chartstyle+xml"
    CHART_COLOR_STYLE = "application/vnd.ms-office.chartcolorstyle+xml"
    VML_DRAWING = "application/vnd.openxmlformats-officedocument.vmlDrawing"

    PRESENTATION = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    PRESENTATIONML_MAIN = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    )
    PRESENTATIONML_MACROENABLED = (
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    )
    PRESENTATIONML_TEMPLATE = (
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
    )
    PRESENTATIONML_TEMPLATE_MACROENABLED = (
        "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
    )
    PRESENTATIONML_SLIDESHOW = (
        "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
    )
    PRESENTATIONML_COMMENTS = (
        "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
    )
    PRESENTATIONML_COMMENT_AUTHORS = (
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
    )
    PRESENTATIONML_FONT_DATA = "application/x-fontdata"
    PRESENTATIONML_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    PRESENTATIONML_SLIDE_MASTER = (
        "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
    )
    PRESENTATIONML_SLIDE_LAYOUT = (
        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
    )
    PRESENTATIONML_TABLE_STYLES = (
        "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"
    )
    PRESENTATIONML_PRES_PROPS = (
        "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
    )
    PRESENTATIONML_VIEW_PROPS = (
        "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
    )
    PRESENTATIONML_TAGS = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
    PRESENTATIONML_NOTES_SLIDE = (
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
    )
    PRESENTATIONML_NOTES_MASTER = (
        "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
    )
    PRESENTATIONML_HANDOUT_MASTER = (
        "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
    )

    SPREADSHEETML_WORKBOOK = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    SPREADSHEETML_WORKBOOK_MACROENABLED = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    SPREADSHEETML_TEMPLATE = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
    )
    SPREADSHEETML_TEMPLATE_MACROENABLED = "application/vnd.ms-excel.template.macroEnabled.main+xml"
    SPREADSHEETML_PRINTER_SETTINGS = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
    )
    SPREADSHEETML_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    SPREADSHEETML_WORKSHEET = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    )
    SPREADSHEETML_CHARTSHEET = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
    )
    SPREADSHEETML_CALC_CHAIN = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
    )
    SPREADSHEETML_SHARED_STRINGS = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    )
    SPREADSHEETML_PIVOT_CACHE_DEFINITION = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
    )
    SPREADSHEETML_PIVOT_CACHE_RECORDS = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
    )
    SPREADSHEETML_PIVOT_TABLE = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    )
    SPREADSHEETML_QUERY_TABLE = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"
    )
    SPREADSHEETML_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    SPREADSHEETML_COMMENTS = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
    )
    SPREADSHEETML_CONNECTIONS = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"
    )
    SPREADSHEETML_EXTERNAL_LINK = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
    )

    DIGITAL_SIGNATURE_ORIGIN_PART = (
        "application/vnd.openxmlformats-package.digital-signature-origin"
    )
    DIGITAL_SIGNATURE_XML_SIGNATURE_PART = (
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
    )
    INK_ML = "application/inkml+xml"

    IMAGE_EMF = "image/x-emf"
    IMAGE_EMF2 = "image/emf"
    IMAGE_WMF = "image/x-wmf"
    IMAGE_GIF = "image/gif"
    IMAGE_JPEG = "image/jpeg"
    IMAGE_JPEG_XR = "image/vnd.ms-photo"
    IMAGE_PICT = "image/pict"
    IMAGE_PNG = "image/png"
    IMAGE_X_PNG = "image/x-png"
    IMAGE_TIFF = "image/tiff"
    IMAGE_WEBP = "image/webp"
    IMAGE_EPS = "application/postscript"
    IMAGE_BMP = "image/bmp"
    IMAGE_SVG = "image/svg+xml"
    OCTET_STREAM = "application/octet-stream"

    EXTENSION_EMF = "emf"
    EXTENSION_WMF = "wmf"
    EXTENSION_GIF = "gif"
    EXTENSION_JPG_1 = "jpg"
    EXTENSION_JPG_2 = "jpeg"
    EXTENSION_JPEG_XR = "wdp"
    EXTENSION_PNG = "png"
    EXTENSION_TIFF = "tiff"
    EXTENSION_WEBP = "webp"
    EXTENSION_EPS = "eps"
    EXTENSION_BMP = "bmp"


#: Content type by lower-case extension, for a package that names an image in a
#: relationship but forgot its content type.
IMAGE_CONTENT_TYPES_BY_EXTENSION: dict[str, str] = {
    "png": ContentTypes.IMAGE_PNG,
    "jpg": ContentTypes.IMAGE_JPEG,
    "jpeg": ContentTypes.IMAGE_JPEG,
    "gif": ContentTypes.IMAGE_GIF,
    "bmp": ContentTypes.IMAGE_BMP,
    "tif": ContentTypes.IMAGE_TIFF,
    "tiff": ContentTypes.IMAGE_TIFF,
    "emf": ContentTypes.IMAGE_EMF,
    "wmf": ContentTypes.IMAGE_WMF,
    "svg": ContentTypes.IMAGE_SVG,
    "webp": ContentTypes.IMAGE_WEBP,
    "wdp": ContentTypes.IMAGE_JPEG_XR,
    "eps": ContentTypes.IMAGE_EPS,
}


def is_xml_content_type(content_type: str | None) -> bool:
    """Whether bytes of this content type are XML."""
    if content_type is None:
        return False
    return (
        content_type.endswith("+xml")
        or content_type == ContentTypes.APPLICATION_XML
        or content_type == ContentTypes.XML
    )


def is_stored_uncompressed(content_type: str | None) -> bool:
    """Already-compressed media: stored rather than deflated in a zip.

    docx4j ``ZipPartStore.shouldCompress``, the content-type half: PNG, JPEG
    and GIF. EMF and WMF are *not* in the list --- they are metafiles and
    deflate well --- which is what CR-002 section 4 means by "emf/wmf left
    deflated". The part-class half (embedded packages, some altChunks) is
    :attr:`Part.compress`.
    """
    return content_type in (
        ContentTypes.IMAGE_PNG,
        ContentTypes.IMAGE_X_PNG,
        ContentTypes.IMAGE_JPEG,
        ContentTypes.IMAGE_GIF,
    )


class ContentTypeManager:
    """``[Content_Types].xml``. docx4j ``ContentTypeManager``.

    Two tables: ``Default`` by lower-cased extension and ``Override`` by
    lower-cased part name, each remembering the spelling it was given so that
    the file is written back as it was read.
    """

    __slots__ = ("_defaults", "_overrides")

    def __init__(self) -> None:
        """An empty manager; :meth:`create_default` is the useful constructor."""
        # extension (lower case) -> (extension as written, content type)
        self._defaults: dict[str, tuple[str, str]] = {}
        # part name key -> (part name as written, content type)
        self._overrides: dict[str, tuple[str, str]] = {}

    # -- construction ------------------------------------------------------

    @classmethod
    def create_default(cls) -> ContentTypeManager:
        """The two defaults every package has: ``rels`` and ``xml``."""
        ctm = cls()
        ctm.add_default_content_type("rels", ContentTypes.RELATIONSHIPS_PART)
        ctm.add_default_content_type("xml", ContentTypes.APPLICATION_XML)
        return ctm

    @classmethod
    def parse(cls, data: bytes) -> ContentTypeManager:
        """Parse ``[Content_Types].xml``."""
        ctm = cls()
        parser = etree.XMLParser(resolve_entities=False, huge_tree=True)
        try:
            root = etree.fromstring(data, parser)
        except etree.XMLSyntaxError as exc:
            raise InvalidFormatException(f"Bad [Content_Types].xml: {exc}") from exc
        if etree.QName(root).localname != "Types":
            raise InvalidFormatException(f"[Content_Types].xml: expected Types, found {root.tag}")
        for element in root:
            if not isinstance(element.tag, str):
                continue
            local = etree.QName(element).localname
            if local == "Default":
                extension = element.get("Extension")
                content_type = element.get("ContentType")
                if extension is not None and content_type is not None:
                    ctm.add_default_content_type(extension, content_type)
            elif local == "Override":
                part_name = element.get("PartName")
                content_type = element.get("ContentType")
                if part_name is not None and content_type is not None:
                    ctm.add_override_content_type(part_name, content_type)
        return ctm

    # -- writing -----------------------------------------------------------

    def to_bytes(self) -> bytes:
        """The file as Word writes it: the declaration, defaults, overrides."""
        root = etree.Element(f"{{{CONTENT_TYPES_NS}}}Types", nsmap={None: CONTENT_TYPES_NS})
        for extension, content_type in self._defaults.values():
            etree.SubElement(
                root,
                f"{{{CONTENT_TYPES_NS}}}Default",
                Extension=extension,
                ContentType=content_type,
            )
        for part_name, content_type in self._overrides.values():
            etree.SubElement(
                root,
                f"{{{CONTENT_TYPES_NS}}}Override",
                PartName=part_name,
                ContentType=content_type,
            )
        from docx4j_py.openpackaging.parts.xml_part import XML_DECLARATION

        # lxml's own declaration uses single quotes; Word writes double, and so
        # does every other part this engine writes.
        return XML_DECLARATION + etree.tostring(root, encoding="utf-8")

    # -- lookup ------------------------------------------------------------

    def get_content_type(self, part_name: PartName | str) -> str | None:
        """The content type of a part: its override, else its extension's default."""
        name = PartName.of(part_name)
        override = self._overrides.get(name.key)
        if override is not None:
            return override[1]
        default = self._defaults.get(name.extension.lower())
        return default[1] if default is not None else None

    def get_default_content_type(self, extension: str) -> str | None:
        """The ``Default`` for an extension."""
        entry = self._defaults.get(extension.lower())
        return entry[1] if entry is not None else None

    def get_override_content_type(self, part_name: PartName | str) -> str | None:
        """The ``Override`` for a part, if it has one."""
        entry = self._overrides.get(PartName.of(part_name).key)
        return entry[1] if entry is not None else None

    def is_content_type_registered(self, content_type: str) -> bool:
        """Whether any default or override names this content type."""
        return any(ct == content_type for _e, ct in self._defaults.values()) or any(
            ct == content_type for _n, ct in self._overrides.values()
        )

    def get_part_names_for_content_type(self, content_type: str) -> list[PartName]:
        """Every part overridden to a content type. docx4j ``getPartNameOverridenByContentType``."""
        return [PartName.of(name) for name, ct in self._overrides.values() if ct == content_type]

    # -- mutation ----------------------------------------------------------

    def add_default_content_type(self, extension: str, content_type: str) -> None:
        """Register a ``Default``."""
        self._defaults[extension.lower()] = (extension, content_type)

    def add_override_content_type(self, part_name: PartName | str, content_type: str) -> None:
        """Register an ``Override``."""
        name = PartName.of(part_name)
        self._overrides[name.key] = (name.name, content_type)

    def add_content_type(self, part_name: PartName | str, content_type: str) -> None:
        """Register a part's content type the way ``addPart`` does.

        docx4j writes a ``Default`` only where one already covers the extension
        with the same value (its JPEG/GIF/PNG special cases compare a content
        type against an *extension* string and so never fire, which this does
        not reproduce: the behaviour it produces is the one below, an Override
        unless a Default already says the same thing).
        """
        name = PartName.of(part_name)
        if self.get_default_content_type(name.extension) == content_type:
            return
        self.add_override_content_type(name, content_type)

    def remove_content_type(self, part_name: PartName | str) -> None:
        """Remove a part's override, or the default for its extension."""
        name = PartName.of(part_name)
        if self._overrides.pop(name.key, None) is not None:
            return
        self._defaults.pop(name.extension.lower(), None)

    def remove_override_content_type(self, part_name: PartName | str) -> None:
        """Remove a part's override."""
        self._overrides.pop(PartName.of(part_name).key, None)

    def remove_default_content_type(self, extension: str) -> None:
        """Remove an extension's default."""
        self._defaults.pop(extension.lower(), None)

    # -- introspection -----------------------------------------------------

    @property
    def defaults(self) -> dict[str, str]:
        """Extension as written -> content type."""
        return {extension: ct for extension, ct in self._defaults.values()}

    @property
    def overrides(self) -> dict[str, str]:
        """Part name as written -> content type."""
        return {name: ct for name, ct in self._overrides.values()}

    def __repr__(self) -> str:
        """``ContentTypeManager(2 defaults, 11 overrides)``."""
        return (
            f"ContentTypeManager({len(self._defaults)} defaults, {len(self._overrides)} overrides)"
        )
