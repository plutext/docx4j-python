"""Relationship types and the package namespaces.

docx4j ``org.docx4j.openpackaging.parts.relationships.Namespaces``, transcribed
constant for constant. A *relationship type* is the URI in a
``<Relationship Type="...">``; it is what tells a consumer what a part is for,
and together with the content type it is what the registry dispatches on.

Several constants share a value on purpose, as in docx4j: ``DOCUMENT`` is also
``PRESENTATIONML_MAIN`` and ``SPREADSHEETML_WORKBOOK``, ``COMMENTS`` is also the
PML and SML comments type, ``FONT`` is also ``PRESENTATIONML_FONT_DATA``, and
``THUMBNAIL`` is ``METADATA_THUMBNAIL``. The names are kept so that docx4j code
reads across.
"""

from __future__ import annotations

__all__ = ["Namespaces"]

_OPC = "http://schemas.openxmlformats.org/package/2006/relationships"
_OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MS = "http://schemas.microsoft.com/office"


class Namespaces:
    """docx4j's ``Namespaces``: the package namespaces and every relationship type."""

    # -- namespaces --------------------------------------------------------
    NAMESPACE_PREFIX_TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006"
    NAMESPACE_PREFIX_STRICT = "http://purl.oclc.org/ooxml/officeDocument"
    CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
    RELATIONSHIPS = _OPC
    RELATIONSHIPS_OFFICEDOC = _OD
    NS_WORD12 = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W_NAMESPACE_DECLARATION = f'xmlns:w="{NS_WORD12}"'
    PKG_XML = "http://schemas.microsoft.com/office/2006/xmlPackage"
    MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    XML_EVENTS = "http://www.w3.org/2001/xml-events"
    XFORMS = "http://www.w3.org/2002/xforms"
    XML_SCHEMA = "http://www.w3.org/2001/XMLSchema"
    WEB_EXTENSION_TASKPANES = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
    WEB_EXTENSION_WEBEXTENSION = (
        "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    )

    # -- relationship types: the package ----------------------------------
    DIGITAL_SIGNATURE = f"{_OPC}/digital-signature/signature"
    DIGITAL_SIGNATURE_CERTIFICATE = f"{_OPC}/digital-signature/certificate"
    DIGITAL_SIGNATURE_ORIGIN = f"{_OPC}/digital-signature/origin"
    PROPERTIES_CORE = f"{_OPC}/metadata/core-properties"
    THUMBNAIL = f"{_OPC}/metadata/thumbnail"
    METADATA_THUMBNAIL = THUMBNAIL

    # -- relationship types: officeDocument --------------------------------
    PROPERTIES_EXTENDED = f"{_OD}/extended-properties"
    PROPERTIES_CUSTOM = f"{_OD}/custom-properties"
    PROPERTIES_COVERPAGE = f"{_MS}/2006/coverPageProps"
    IMAGE = f"{_OD}/image"
    DOCUMENT = f"{_OD}/officeDocument"
    DOCUMENT_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
    HYPERLINK = f"{_OD}/hyperlink"
    CUSTOM_XML_DATA_STORAGE = f"{_OD}/customXml"
    CUSTOM_XML_DATA_STORAGE_PROPERTIES = f"{_OD}/customXmlProps"
    WEB_SETTINGS = f"{_OD}/webSettings"
    SETTINGS = f"{_OD}/settings"
    STYLES = f"{_OD}/styles"
    THEME = f"{_OD}/theme"
    THEME_OVERRIDE = f"{_OD}/themeOverride"
    FONT_TABLE = f"{_OD}/fontTable"
    FONT = f"{_OD}/font"
    HEADER = f"{_OD}/header"
    FOOTER = f"{_OD}/footer"
    GLOSSARY_DOCUMENT = f"{_OD}/glossaryDocument"
    NUMBERING = f"{_OD}/numbering"
    FOOTNOTES = f"{_OD}/footnotes"
    ENDNOTES = f"{_OD}/endnotes"
    COMMENTS = f"{_OD}/comments"
    #: Note the capital F: [MS-OE376] writes ``aFChunk`` where the spec says
    #: ``afChunk``, and Word writes what [MS-OE376] says.
    AF = f"{_OD}/aFChunk"
    SUBDOCUMENT = f"{_OD}/subDocument"
    ATTACHED_TEMPLATE = f"{_OD}/attachedTemplate"
    OLE_OBJECT = f"{_OD}/oleObject"
    ACTIVEX_XML_OBJECT = f"{_OD}/control"
    EMBEDDED_PKG = f"{_OD}/package"
    VML = f"{_OD}/vmlDrawing"

    # -- relationship types: Microsoft extensions --------------------------
    COMMENTS_EXTENDED = f"{_MS}/2011/relationships/commentsExtended"
    COMMENTS_IDS = f"{_MS}/2016/09/relationships/commentsIds"
    COMMENTS_EXTENSIBLE = f"{_MS}/2018/08/relationships/commentsExtensible"
    OFFICE_2011_PEOPLE = f"{_MS}/2011/relationships/people"
    VBA_PROJECT = f"{_MS}/2006/relationships/vbaProject"
    VBA_DATA_WORD = f"{_MS}/2006/relationships/wordVbaData"
    KEYMAP = f"{_MS}/2006/relationships/keyMapCustomizations"
    VBA_PROJECT_SIGNATURE = f"{_MS}/2006/relationships/vbaProjectSignature"
    CHART_EX = f"{_MS}/2014/relationships/chartEx"
    CHART_COLOR_STYLE = f"{_MS}/2011/relationships/chartColorStyle"
    CHART_STYLE = f"{_MS}/2011/relationships/chartStyle"
    DRAWINGML_DIAGRAM_DRAWING = f"{_MS}/2007/relationships/diagramDrawing"
    GRAPHIC_FRAME_DOC = f"{_MS}/2006/relationships/graphicFrameDoc"

    # -- relationship types: DrawingML -------------------------------------
    DRAWINGML_DIAGRAM_DATA = f"{_OD}/diagramData"
    DRAWINGML_DIAGRAM_LAYOUT = f"{_OD}/diagramLayout"
    DRAWINGML_DIAGRAM_COLORS = f"{_OD}/diagramColors"
    DRAWINGML_DIAGRAM_STYLE = f"{_OD}/diagramQuickStyle"
    DRAWINGML_DIAGRAM_LAYOUT_HEADER = f"{_OD}/diagramLayoutHeader"
    CHART_USER_SHAPES = f"{_OD}/chartUserShapes"

    # -- relationship types: PresentationML --------------------------------
    PRESENTATIONML_MAIN = DOCUMENT
    PRESENTATIONML_COMMENTS = COMMENTS
    PRESENTATIONML_COMMENT_AUTHORS = f"{_OD}/commentAuthors"
    PRESENTATIONML_FONT_DATA = FONT
    PRESENTATIONML_SLIDE = f"{_OD}/slide"
    PRESENTATIONML_SLIDE_MASTER = f"{_OD}/slideMaster"
    PRESENTATIONML_SLIDE_LAYOUT = f"{_OD}/slideLayout"
    PRESENTATIONML_TABLE_STYLES = f"{_OD}/tableStyles"
    PRESENTATIONML_PRES_PROPS = f"{_OD}/presProps"
    PRESENTATIONML_VIEW_PROPS = f"{_OD}/viewProps"
    PRESENTATIONML_TAGS = f"{_OD}/tags"
    PRESENTATIONML_NOTES_SLIDE = f"{_OD}/notesSlide"
    PRESENTATIONML_NOTES_MASTER = f"{_OD}/notesMaster"
    PRESENTATIONML_HANDOUT_MASTER = f"{_OD}/handoutMaster"

    # -- relationship types: SpreadsheetML ---------------------------------
    SPREADSHEETML_WORKBOOK = DOCUMENT
    SPREADSHEETML_WORKSHEET = f"{_OD}/worksheet"
    SPREADSHEETML_CHARTSHEET = f"{_OD}/chartsheet"
    SPREADSHEETML_PRINTER_SETTINGS = f"{_OD}/printerSettings"
    SPREADSHEETML_CALC_CHAIN = f"{_OD}/calcChain"
    SPREADSHEETML_SHARED_STRINGS = f"{_OD}/sharedStrings"
    SPREADSHEETML_STYLES = STYLES
    SPREADSHEETML_DRAWING = f"{_OD}/drawing"
    SPREADSHEETML_CHART = f"{_OD}/chart"
    SPREADSHEETML_COMMENTS = COMMENTS
    SPREADSHEETML_PIVOT_TABLE = f"{_OD}/pivotTable"
    SPREADSHEETML_QUERY_TABLE = f"{_OD}/queryTable"
    SPREADSHEETML_TABLE = f"{_OD}/table"
    SPREADSHEETML_CONNECTIONS = f"{_OD}/connections"
    SPREADSHEETML_PIVOT_CACHE_DEFINITION = f"{_OD}/pivotCacheDefinition"
    SPREADSHEETML_PIVOT_CACHE_RECORDS = f"{_OD}/pivotCacheRecords"
    SPREADSHEETML_EXTERNAL_LINK = f"{_OD}/externalLink"


#: The relationship types whose target is a package's main part. docx4j
#: ``PackageRelsUtil.getNameOfMainPart``.
MAIN_PART_RELATIONSHIP_TYPES: tuple[str, ...] = (
    Namespaces.DOCUMENT,
    Namespaces.GRAPHIC_FRAME_DOC,
    Namespaces.DRAWINGML_DIAGRAM_LAYOUT,
    Namespaces.DOCUMENT_STRICT,
)
