"""``WordprocessingMLPackage``: a ``.docx``. docx4j ``WordprocessingMLPackage``.

CR-002 section 5.5. The package class for the four WordprocessingML main content
types (document, macro-enabled document, template, macro-enabled template), the
part shortcuts docx4j's ``DocumentPart`` keeps, and ``create_package``.

``create_package`` is docx4j's ``createPackage(PageSizePaper, boolean)``, part
for part: a main document part with one section, docx4j's own default styles,
a settings part with the ``w:compat`` Word writes into a new document (section
12.10), core and extended document properties, and --- since CR-002 Phase B
(section 15.3), a departure docx4j's own ``addDefaultThemePart`` supports but
its ``createPackage`` does not take --- **a theme part**, so that the
``minorHAnsi`` the default styles name resolves to a face. No numbering part and
no font table, as docx4j leaves it.
"""

from __future__ import annotations

import datetime
import os
import re
from typing import IO, TYPE_CHECKING, Any, Literal

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.openpackaging.packages.opc_package import (
    OpcPackage,
    check_kind,
    new_memory_package,
    register_package_class,
)
from docx4j_py.openpackaging.parts.docprops import (
    DocPropsCorePart,
    DocPropsExtendedPart,
)
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.wml import (
    DocumentSettingsPart,
    FontTablePart,
    FooterPart,
    HeaderPart,
    MainDocumentPart,
    NumberingDefinitionsPart,
    StyleDefinitionsPart,
)
from docx4j_py.openpackaging.stores import PartStore

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.parts.dml import ThemePart

__all__ = ["PAGE_SIZES", "PageSizePaper", "WordprocessingMLPackage"]

#: docx4j's ``PageSizePaper``: six values, and B4JIS is the only B size.
PageSizePaper = Literal["LETTER", "LEGAL", "A3", "A4", "A5", "B4JIS"]

#: Portrait width and height in twips and the ``w:code``, from docx4j's
#: ``PageDimensions.setPgSize``. Landscape swaps width and height and sets
#: ``w:orient="landscape"``; portrait leaves ``w:orient`` unset, as docx4j does.
PAGE_SIZES: dict[str, tuple[int, int, int]] = {
    "LETTER": (12240, 15840, 1),
    "LEGAL": (12240, 20160, 5),
    "A3": (16839, 23814, 8),
    "A4": (11907, 16839, 9),
    "A5": (8391, 11907, 11),
    "B4JIS": (14572, 20639, 12),
}

#: docx4j's ``MarginsWellKnown``: top, bottom, left, right in twips. ``NORMAL``
#: is the default and the only one ``create_package`` offers. ``w:header``,
#: ``w:footer`` and ``w:gutter`` are deliberately **not** written: docx4j's
#: ``setMargins`` does not write them either.
PAGE_MARGINS: dict[str, tuple[int, int, int, int]] = {
    "NORMAL": (1440, 1440, 1440, 1440),
    "NARROW": (720, 720, 720, 720),
    "MODERATE": (1440, 1440, 1080, 1080),
    "WIDE": (1440, 1440, 2880, 2880),
}

MAIN_CONTENT_TYPES = (
    ContentTypes.WORDPROCESSINGML_DOCUMENT,
    ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED,
    ContentTypes.WORDPROCESSINGML_TEMPLATE,
    ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED,
)


class WordprocessingMLPackage(OpcPackage):
    """A WordprocessingML package: a ``.docx``, ``.docm``, ``.dotx`` or ``.dotm``."""

    __slots__ = (
        "_font_mapper",
        "_glossary_entries",
        "_property_resolver",
        "_run_font_selector",
        "default_theme",
        "main_document_part",
    )

    def __init__(self) -> None:
        """An empty package with no main document part yet."""
        super().__init__()
        self.main_document_part: MainDocumentPart | None = None
        self._property_resolver: Any = None
        self._font_mapper: Any = None
        self._run_font_selector: Any = None
        #: Which Office theme this package's **themeless** font references
        #: resolve against, and which theme part ``create_package`` gives it:
        #: ``"2023"`` (Aptos, Word 365 and the default), ``"2013"`` (Calibri) or
        #: ``"2007"`` (Cambria). docx4j's one JVM-wide ``docx4j.fonts.defaultTheme``
        #: is per package here, because a server may hold several at once
        #: (CR-002 section 15.3). None takes the process-wide setting.
        self.default_theme: str | None = None
        # CR-019 section 4.9: the glossary's placeholder entries, read privately
        # and with those written since waiting for the save
        # (``docx4j_py.model.content.placeholders``), made on first use. A plain
        # container: the ``glossary_entries`` property is installed by
        # ``docx4j_py.model.content.register()``, and this layer imports nothing
        # from the content API.
        self._glossary_entries: Any = None

    # -- the resolution utilities (CR-002 section 6) -----------------------

    def get_property_resolver(self, refresh: bool = False) -> Any:
        """The package's :class:`~docx4j_py.model.property_resolver.PropertyResolver`.

        docx4j ``MainDocumentPart.getPropertyResolver(boolean)``: one per
        package, made on first use and kept. `refresh` re-reads the styles part
        and throws the caches away, which is what a caller who has **modified or
        removed** a style needs; a style merely *added* is picked up on its own.

        The styles and numbering parts are read without being unmarshalled, so
        building a resolver costs a document no byte of its round trip.

        The import is inside the method because the engine imports nothing from
        the model at module level (``tests/openpackaging/test_threads_and_import.py``).
        """
        from docx4j_py.model.property_resolver import PropertyResolver

        if self._property_resolver is None:
            self._property_resolver = PropertyResolver(self)
        elif refresh:
            self._property_resolver.refresh()
        return self._property_resolver

    def refresh_property_resolver(self) -> Any:
        """Re-read the styles part into the resolver, if one has been built.

        What every writer of ``/word/styles.xml`` calls:
        :func:`docx4j_py.model.content.styles.ensure_style` through
        :func:`docx4j_py.model.listnumbering.invalidate_styles`. The run font
        selector's cached default font goes with it, and the font mapper --- a
        style may name a font --- is thrown away.
        """
        if self._property_resolver is not None:
            self._property_resolver.refresh()
        if self._run_font_selector is not None:
            self._run_font_selector.refresh()
        self._font_mapper = None
        return self._property_resolver

    def get_run_font_selector(self, refresh: bool = False) -> Any:
        """The package's :class:`~docx4j_py.model.fonts.RunFontSelector`.

        docx4j ``MainDocumentPart.getRunFontSelector``: one per package, made on
        first use and kept. It reads the theme part's ``a:theme`` and the
        settings part's ``w:themeFontLang``, both with ``read_contents()``, so
        resolving a document's fonts costs it no byte of its round trip.
        """
        from docx4j_py.model.fonts import RunFontSelector

        if refresh:
            self._run_font_selector = None
        if self._run_font_selector is None:
            theme_part = self.theme_part
            theme = None if theme_part is None else theme_part.read_contents()
            settings = self.document_settings_part
            theme_font_lang = None
            if settings is not None:
                theme_font_lang = settings.read_contents().theme_font_lang
            self._run_font_selector = RunFontSelector(
                resolver=self.get_property_resolver(),
                theme=theme,
                theme_font_lang=theme_font_lang,
                default_theme=self.default_theme,
            )
        return self._run_font_selector

    @property
    def run_font_selector(self) -> Any:
        """The selector where one has been built, else None. Never builds one."""
        return self._run_font_selector

    def get_font_mapper(self, refresh: bool = False) -> Any:
        """The package's font mapper, populated over the fonts the document uses.

        docx4j ``WordprocessingMLPackage.getFontMapper`` /
        ``setFontMapper(Mapper)``: an
        :class:`~docx4j_py.model.fonts.IdentityPlusMapper` over the faces
        docx4j's four font jars carry, which is the environment the parity
        goldens record. A consumer with real fonts calls
        :meth:`set_font_mapper` with its own.
        """
        from docx4j_py.model.fonts import IdentityPlusMapper

        if refresh:
            self._font_mapper = None
        if self._font_mapper is None:
            self.set_font_mapper(IdentityPlusMapper())
        return self._font_mapper

    def set_font_mapper(self, mapper: Any, populate: bool = True) -> Any:
        """Use `mapper` for this package, running its passes over the fonts in use."""
        main = self.main_document_part
        if populate and main is not None:
            font_table = self.font_table_part
            fonts = None if font_table is None else font_table.read_contents()
            mapper.populate(sorted(main.fonts_in_use()), fonts)
        self._font_mapper = mapper
        return mapper

    @property
    def font_mapper(self) -> Any:
        """:meth:`get_font_mapper`, so that ``pkg.font_mapper`` reads as docx4j's does."""
        return self.get_font_mapper()

    # -- loading -----------------------------------------------------------

    @classmethod
    def load(
        cls,
        source: str | os.PathLike[str] | bytes | IO[bytes] | PartStore,
        *,
        options: LoadOptions | None = None,
    ) -> WordprocessingMLPackage:
        """Load a ``.docx``; raises if the main part is not a ``w:document``."""
        package = OpcPackage.load(source, options=options)
        check_kind(package, WordprocessingMLPackage, "WordprocessingML")
        return package  # type: ignore[return-value]

    def set_part_shortcut(self, part: Part, relationship_type: str) -> bool:
        """Keep the main document part, plus what ``OpcPackage`` keeps."""
        if relationship_type in (Namespaces.DOCUMENT, Namespaces.DOCUMENT_STRICT):
            self.main_document_part = part  # type: ignore[assignment]
            return True
        return super().set_part_shortcut(part, relationship_type)

    # -- the part shortcuts ------------------------------------------------

    def get_main_document_part(self) -> MainDocumentPart:
        """The main document part; raises when there is none."""
        from docx4j_py.openpackaging.exceptions import Docx4JException

        if self.main_document_part is None:
            raise Docx4JException("This package has no main document part")
        return self.main_document_part

    @property
    def style_definitions_part(self) -> StyleDefinitionsPart | None:
        """``/word/styles.xml``, through the main document part's relationships."""
        return self._shortcut("style_definitions_part")

    @property
    def numbering_definitions_part(self) -> NumberingDefinitionsPart | None:
        """``/word/numbering.xml``."""
        return self._shortcut("numbering_definitions_part")

    @property
    def font_table_part(self) -> FontTablePart | None:
        """``/word/fontTable.xml``."""
        return self._shortcut("font_table_part")

    @property
    def document_settings_part(self) -> DocumentSettingsPart | None:
        """``/word/settings.xml``."""
        return self._shortcut("document_settings_part")

    @property
    def web_settings_part(self) -> Any:
        """``/word/webSettings.xml``."""
        return self._shortcut("web_settings_part")

    @property
    def theme_part(self) -> ThemePart | None:
        """``/word/theme/theme1.xml``."""
        return self._shortcut("theme_part")

    def header_parts(self) -> list[HeaderPart]:
        """Every header part, in relationship order."""
        main = self.main_document_part
        return main.header_parts() if main is not None else []

    def footer_parts(self) -> list[FooterPart]:
        """Every footer part, in relationship order."""
        main = self.main_document_part
        return main.footer_parts() if main is not None else []

    def _shortcut(self, attribute: str) -> Any:
        main = self.main_document_part
        return getattr(main, attribute, None) if main is not None else None

    # -- creating ----------------------------------------------------------

    @classmethod
    def create_package(
        cls,
        *,
        page_size: PageSizePaper = "A4",
        landscape: bool = False,
        margins: str = "NORMAL",
        default_theme: str | None = None,
    ) -> WordprocessingMLPackage:
        """A new, empty document. docx4j ``WordprocessingMLPackage.createPackage``.

        What it creates, in docx4j's order:

        * ``/word/document.xml`` with an empty body and one ``w:sectPr`` carrying
          the page size for `page_size` and 2.54 cm margins;
        * ``/word/styles.xml``, docx4j's own default styles, related **from the
          document part** (not from the package), as docx4j relates it;
        * ``/docProps/core.xml`` with ``dcterms:created`` and ``dcterms:modified``
          set to now (UTC, seconds), and ``/docProps/app.xml`` with
          ``Application`` (``docx4j-python``) and ``AppVersion`` (``XX.YYYY``, the
          only form Word accepts), both related from the package. docx4j writes
          these only when its ``docx4j.dc.write`` / ``docx4j.App.write``
          properties are set; a created document should not carry empty parts, so
          here they are always written (CR-002 section 12.8);
        * ``/word/settings.xml`` with the ``w:compat`` Word writes into a new
          document (:data:`~docx4j_py.openpackaging.parts.wml.NEW_DOCUMENT_COMPAT`):
          ``compatibilityMode`` **15**, so the document targets **Word 2013 and
          later** and does not open in Word's Compatibility Mode, plus the five
          settings Word writes beside it. docx4j writes only
          ``overrideTableStyleFontSizeAndJustification`` and no mode at all,
          which leaves a created document in Word 2007 format; CR-002 section
          12.10 records the departure, and ``pkg.compatibility_mode`` changes it.

        Args:
            page_size: one of :data:`PAGE_SIZES`; ``A4`` as docx4j's code
                defaults (its sample ``docx4j.properties`` says ``LETTER``, but
                the code says A4).
            landscape: swap width and height and write ``w:orient``.
            margins: one of :data:`PAGE_MARGINS`.
            default_theme: ``"2023"`` (the default: Aptos, what Word 365
                writes), ``"2013"`` (Calibri) or ``"2007"`` (Cambria) --- which
                Office theme the new document's theme part carries.
        """
        from docx4j_py.wml import Body, CtSectPrPgMar, CtSectPrPgSz, Document, SectPr

        if page_size not in PAGE_SIZES:
            raise ValueError(f"Unknown page size {page_size!r}; one of {sorted(PAGE_SIZES)}")
        if margins not in PAGE_MARGINS:
            raise ValueError(f"Unknown margins {margins!r}; one of {sorted(PAGE_MARGINS)}")

        package: WordprocessingMLPackage = new_memory_package(cls)  # type: ignore[assignment]
        # CR-003 section 3.4: a created document always gets paragraph ids, so
        # that every paragraph an agent makes has an address that survives an
        # insert before it. A loaded document decides for itself (``None``).
        package.assigns_para_ids = True
        package.content_type_manager.add_default_content_type(
            "rels", ContentTypes.RELATIONSHIPS_PART
        )
        package.content_type_manager.add_default_content_type("xml", ContentTypes.APPLICATION_XML)

        width, height, code = PAGE_SIZES[page_size]
        page_sz = CtSectPrPgSz(
            w=height if landscape else width,
            h=width if landscape else height,
            code=code,
        )
        if landscape:
            page_sz.orient = "landscape"
        top, bottom, left, right = PAGE_MARGINS[margins]
        page_mar = CtSectPrPgMar(top=top, bottom=bottom, left=left, right=right)

        document = Document(body=Body(sect_pr=SectPr(pg_sz=page_sz, pg_mar=page_mar)))
        main = MainDocumentPart()
        main.set_contents(document)
        package.add_target_part(main)

        styles = StyleDefinitionsPart()
        styles.package = package
        styles.unmarshal_default_styles()
        main.add_target_part(styles)

        core = DocPropsCorePart()
        core.set_contents(default_core_properties())
        package.add_target_part(core)

        app = DocPropsExtendedPart()
        app.set_contents(default_extended_properties())
        package.add_target_part(app)

        settings = DocumentSettingsPart()
        main.add_target_part(settings)
        settings.set_contents(_empty_settings())
        settings.set_default_compat_settings()

        # A theme part, as docx4j's addDefaultThemePart does, so that a package
        # created by either library resolves its minorHAnsi defaults to the same
        # face (CR-002 section 15.3). docx4j's createPackage does not add one;
        # the departure is deliberate and is recorded there.
        from docx4j_py.model.fonts import theme_of_setting
        from docx4j_py.openpackaging.parts.dml import ThemePart

        package.default_theme = theme_of_setting(default_theme).value
        theme = ThemePart()
        theme.set_xml(theme_of_setting(default_theme).xml)
        main.add_target_part(theme)

        return package


#: What ``Application`` says in a document this library creates.
APPLICATION_NAME = "docx4j-python"


def app_version() -> str:
    """The installed version in Word's ``XX.YYYY`` form.

    docx4j's ``Save`` warns that anything else (a ``-SNAPSHOT`` suffix, say) makes
    Word 2010 x64 report the document as corrupt, so ``0.0.1`` becomes ``0.0001``:
    the major version, a dot, then minor and patch as two digits each.
    """
    from importlib.metadata import PackageNotFoundError, version

    try:
        raw = version("docx4j-py")
    except PackageNotFoundError:
        raw = "0.0.1"
    parts = [int(x) for x in re.findall(r"\d+", raw)[:3]] + [0, 0, 0]
    major, minor, patch = parts[0], parts[1], parts[2]
    return f"{major}.{min(minor, 99):02d}{min(patch, 99):02d}"


def default_core_properties(now: datetime.datetime | None = None) -> Any:
    """``cp:coreProperties`` for a new document: ``created`` and ``modified``.

    Both are ``dcterms:W3CDTF`` values, which the serialiser writes with the
    ``xsi:type`` Word expects. Creator and last-modified-by are the user's
    business, not the library's (docx4j's default puts its own name there only
    when ``docx4j.dc.write`` is set); set them on ``contents`` if wanted.
    """
    from docx4j_xsdata.models.datatype import XmlDateTime

    from docx4j_py.docprops.core import CoreProperties
    from docx4j_py.docprops.dcterms import W3CDTF

    if now is None:
        now = datetime.datetime.now(datetime.timezone.utc)
    stamp = XmlDateTime.from_datetime(now.astimezone(datetime.timezone.utc).replace(microsecond=0))
    return CoreProperties(created=W3CDTF(value=stamp), modified=W3CDTF(value=stamp))


def default_extended_properties() -> Any:
    """``Properties`` for a new document: ``Application`` and ``AppVersion``."""
    from docx4j_py.docprops.extended import Properties

    return Properties(application=APPLICATION_NAME, app_version=app_version())


def _empty_settings() -> Any:
    from docx4j_py.wml import Settings

    return Settings()


register_package_class(list(MAIN_CONTENT_TYPES), WordprocessingMLPackage)
