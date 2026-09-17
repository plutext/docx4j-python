"""The WordprocessingML parts. docx4j ``openpackaging.parts.WordprocessingML``.

CR-002 section 5.4. One class per part kind, each an
:class:`~docx4j_py.openpackaging.parts.xml_part.XmlPart` over the model class its
root element is. The default part names are docx4j's.

:class:`DocumentPart` is docx4j's: the shortcuts a main or glossary document
keeps to the parts it relates to, filled in by ``set_part_shortcut`` as the
loader walks the relationships.

A departure from docx4j's layout, recorded in CR-002 section 12: these are one
module rather than a package of one file per class. Python's import cost is per
module and the classes are four lines each.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, ClassVar

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.default_xml_part import VMLPart
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.xml_part import XmlPart

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.wml import (
        Document,
        Fonts,
        Numbering,
        Styles,
    )

__all__ = [
    "NEW_DOCUMENT_COMPAT",
    "WORD_COMPAT_URI",
    "CommentsExtendedPart",
    "CommentsPart",
    "DocumentPart",
    "DocumentSettingsPart",
    "EndnotesPart",
    "FontTablePart",
    "FooterPart",
    "FootnotesPart",
    "GlossaryDocumentPart",
    "HeaderPart",
    "KeyMapCustomizationsPart",
    "MainDocumentPart",
    "NumberingDefinitionsPart",
    "PeoplePart",
    "StyleDefinitionsPart",
    "VMLPart",
    "VbaDataPart",
    "WebSettingsPart",
]

#: The ``w:uri`` every one of Word's own ``w:compatSetting``\ s carries.
WORD_COMPAT_URI = "http://schemas.microsoft.com/office/word"

#: The ``w:compat`` Word writes into a **new** document, in its own order:
#: ``compatibilityMode`` 15 (Word 2013, and what Word 2016, 2019 and 365 write
#: too) and the five settings beside it. ``create_package`` writes exactly this,
#: so that a created document is not opened in Word's Compatibility Mode
#: (CR-002 section 12.10, 2026-09-17). The content API's
#: ``pkg.compatibility_mode`` reads and rewrites the first of them.
NEW_DOCUMENT_COMPAT: tuple[tuple[str, str], ...] = (
    ("compatibilityMode", "15"),
    ("overrideTableStyleFontSizeAndJustification", "1"),
    ("enableOpenTypeFeatures", "1"),
    ("doNotFlipMirrorIndents", "1"),
    ("differentiateMultirowTableHeaders", "1"),
    ("useWord2013TrackBottomHyphenation", "0"),
)

_W = Namespaces.NS_WORD12
_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
_WNE = "http://schemas.microsoft.com/office/word/2006/wordml"


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


class DocumentPart(XmlPart[Any]):
    """What a main or glossary document knows about the parts around it.

    docx4j ``DocumentPart``: the typed shortcuts, filled in by
    :meth:`set_part_shortcut` while the loader walks this part's relationships.
    """

    __slots__ = (
        "comments_extended_part",
        "comments_extensible_part",
        "comments_ids_part",
        "comments_part",
        "document_settings_part",
        "endnotes_part",
        "font_table_part",
        "footnotes_part",
        "numbering_definitions_part",
        "people_part",
        "style_definitions_part",
        "theme_part",
        "web_settings_part",
    )

    #: Relationship type -> the attribute it fills in. docx4j's switch.
    _SHORTCUTS: ClassVar[dict[str, str]] = {
        Namespaces.STYLES: "style_definitions_part",
        Namespaces.NUMBERING: "numbering_definitions_part",
        Namespaces.FONT_TABLE: "font_table_part",
        Namespaces.THEME: "theme_part",
        Namespaces.SETTINGS: "document_settings_part",
        Namespaces.WEB_SETTINGS: "web_settings_part",
        Namespaces.COMMENTS: "comments_part",
        Namespaces.COMMENTS_EXTENDED: "comments_extended_part",
        Namespaces.COMMENTS_IDS: "comments_ids_part",
        Namespaces.COMMENTS_EXTENSIBLE: "comments_extensible_part",
        Namespaces.FOOTNOTES: "footnotes_part",
        Namespaces.ENDNOTES: "endnotes_part",
        Namespaces.OFFICE_2011_PEOPLE: "people_part",
    }

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        """Build the part with every shortcut unset."""
        super().__init__(*args, **kwargs)
        for attribute in DocumentPart._SHORTCUTS.values():
            setattr(self, attribute, None)

    def set_part_shortcut(self, part: Part, relationship_type: str) -> bool:
        """Fill in a shortcut for a well-known relationship type."""
        attribute = DocumentPart._SHORTCUTS.get(relationship_type)
        if attribute is None:
            return False
        setattr(self, attribute, part)
        return True

    # -- the parts there can be several of ---------------------------------

    def _targets(self, relationship_type: str) -> list[Part]:
        rels = self.relationships_part
        if rels is None:
            return []
        out = []
        for rel in rels.get_relationships_by_type(relationship_type):
            part = rels.get_part(rel)
            if part is not None:
                out.append(part)
        return out

    def header_parts(self) -> list[HeaderPart]:
        """Every header this document relates to, in relationship order."""
        return self._targets(Namespaces.HEADER)  # type: ignore[return-value]

    def footer_parts(self) -> list[FooterPart]:
        """Every footer this document relates to, in relationship order."""
        return self._targets(Namespaces.FOOTER)  # type: ignore[return-value]


class MainDocumentPart(DocumentPart):
    """``/word/document.xml``. docx4j ``MainDocumentPart``."""

    __slots__ = ("glossary_document_part", "key_map_customizations_part")

    model_class_path = "docx4j_py.wml.Document"

    def __init__(
        self,
        part_name: PartName | str = "/word/document.xml",
        content_type: str = ContentTypes.WORDPROCESSINGML_DOCUMENT,
    ) -> None:
        """Build the main document part."""
        super().__init__(part_name, content_type, Namespaces.DOCUMENT, _w("document"))
        self.glossary_document_part: GlossaryDocumentPart | None = None
        self.key_map_customizations_part: KeyMapCustomizationsPart | None = None

    def set_part_shortcut(self, part: Part, relationship_type: str) -> bool:
        """Also keep the glossary document and the key-map customisations."""
        if relationship_type == Namespaces.GLOSSARY_DOCUMENT:
            self.glossary_document_part = part  # type: ignore[assignment]
            return True
        if relationship_type == Namespaces.KEYMAP:
            self.key_map_customizations_part = part  # type: ignore[assignment]
            return True
        return super().set_part_shortcut(part, relationship_type)

    @property
    def contents(self) -> Document:
        """The ``w:document``."""
        return super().contents

    @property
    def body_element(self) -> Any:
        """The typed ``w:body``, created if the document has none.

        CR-002 called this ``body``. CR-003 Phase B gives ``part.body`` to the
        content API's :class:`~docx4j_py.model.content.Body` view --- registered
        on :class:`~docx4j_py.openpackaging.parts.xml_part.XmlPart` by the
        content module, which this layer never imports --- so the element keeps
        the longer name, and ``part.contents.body`` is the same object.
        """
        from docx4j_py.wml import Body

        document = self.contents
        if document.body is None:
            document.body = Body()
            document.body.parent = document
        return document.body


class GlossaryDocumentPart(DocumentPart):
    """``/word/glossary/document.xml``. docx4j ``GlossaryDocumentPart``."""

    __slots__ = ()

    model_class_path = "docx4j_py.wml.GlossaryDocument"

    def __init__(self, part_name: PartName | str = "/word/glossary/document.xml") -> None:
        """Build the glossary document part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_GLOSSARYDOCUMENT,
            Namespaces.GLOSSARY_DOCUMENT,
            _w("glossaryDocument"),
        )


class StyleDefinitionsPart(XmlPart["Styles"]):
    """``/word/styles.xml``. docx4j ``StyleDefinitionsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Styles"

    def __init__(
        self,
        part_name: PartName | str = "/word/styles.xml",
        content_type: str = ContentTypes.WORDPROCESSINGML_STYLES,
    ) -> None:
        """Build the styles part."""
        super().__init__(part_name, content_type, Namespaces.STYLES, _w("styles"))

    def unmarshal_default_styles(self) -> Styles:
        """Load docx4j's own default styles into this part.

        docx4j ``StyleDefinitionsPart.unmarshalDefaultStyles``; the resource is
        docx4j's ``styles.xml``, copied under ``docx4j_py/resources/openpackaging``.
        """
        from docx4j_py.openpackaging.resources import default_part_bytes

        self.set_bytes(default_part_bytes("styles.xml"))
        return self.contents


class NumberingDefinitionsPart(XmlPart["Numbering"]):
    """``/word/numbering.xml``. docx4j ``NumberingDefinitionsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Numbering"

    def __init__(self, part_name: PartName | str = "/word/numbering.xml") -> None:
        """Build the numbering part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_NUMBERING,
            Namespaces.NUMBERING,
            _w("numbering"),
        )

    def unmarshal_default_numbering(self) -> Numbering:
        """Load docx4j's own default numbering definitions into this part."""
        from docx4j_py.openpackaging.resources import default_part_bytes

        self.set_bytes(default_part_bytes("numbering.xml"))
        return self.contents

    # -- list numbering (CR-002 section 6.2, delivered by CR-003 Phase H) ---

    def get_emulator(self, reset: bool = False) -> Any:
        """The numbering emulator of this part's package. docx4j ``getEmulator``.

        :class:`docx4j_py.model.listnumbering.Emulator`: the definitions read
        (from bytes --- this does **not** unmarshal the part) and the counters
        to walk a story with. `reset` throws the definitions and the counted
        labels away, as docx4j's ``getEmulator(true)`` does.

        The import is inside the method because the engine imports nothing from
        the model at module level (CR-001 section 13.5, and
        ``tests/openpackaging/test_threads_and_import.py``).
        """
        from docx4j_py.model.listnumbering import emulator_of

        emulator = emulator_of(self.package)
        if emulator is not None and reset:
            emulator.refresh()
        return emulator

    @property
    def numbering_state(self) -> Any:
        """The counters the state-less emulator calls use. docx4j ``getNumberingState``."""
        emulator = self.get_emulator()
        return None if emulator is None else emulator.numbering_state

    @property
    def abstract_list_definitions(self) -> Any:
        """``w:abstractNumId`` -> the definition. docx4j ``getAbstractListDefinitions``."""
        emulator = self.get_emulator()
        return {} if emulator is None else emulator.abstract_list_definitions

    @property
    def instance_list_definitions(self) -> Any:
        """``w:numId`` -> the definition. docx4j ``getInstanceListDefinitions``."""
        emulator = self.get_emulator()
        return {} if emulator is None else emulator.instance_list_definitions


class FontTablePart(XmlPart["Fonts"]):
    """``/word/fontTable.xml``. docx4j ``FontTablePart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Fonts"

    def __init__(self, part_name: PartName | str = "/word/fontTable.xml") -> None:
        """Build the font table part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_FONTTABLE,
            Namespaces.FONT_TABLE,
            _w("fonts"),
        )

    def unmarshal_default_fonts(self) -> Fonts:
        """Load docx4j's own default font table into this part."""
        from docx4j_py.openpackaging.resources import default_part_bytes

        self.set_bytes(default_part_bytes("fontTable.xml"))
        return self.contents


class DocumentSettingsPart(XmlPart["Settings"]):
    """``/word/settings.xml``. docx4j ``DocumentSettingsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Settings"

    def __init__(self, part_name: PartName | str = "/word/settings.xml") -> None:
        """Build the settings part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_SETTINGS,
            Namespaces.SETTINGS,
            _w("settings"),
        )

    def set_compat_setting(self, name: str, value: str) -> None:
        """Write one of Word's own ``w:compatSetting``\\ s, appending it in order."""
        from docx4j_py.wml import CTCompat, CTCompatSetting

        settings = self.contents
        if settings.compat is None:
            settings.compat = CTCompat()
            settings.compat.parent = settings
        settings.compat.compat_setting.append(
            CTCompatSetting(name=name, uri=WORD_COMPAT_URI, val=value)
        )

    def set_override_table_style_font_size_and_justification(self, value: bool) -> None:
        """Write Word's ``overrideTableStyleFontSizeAndJustification`` compat setting.

        docx4j's ``createPackage`` sets it on every new document; without it Word
        applies a table style's font size and justification to direct
        formatting, which is not what anyone means.
        """
        self.set_compat_setting(
            "overrideTableStyleFontSizeAndJustification", "1" if value else "0"
        )

    def set_default_compat_settings(self) -> None:
        """Write the whole ``w:compat`` Word writes into a new document.

        :data:`NEW_DOCUMENT_COMPAT`, in Word's own order, ``compatibilityMode``
        15 first: a document this library creates targets **Word 2013 and
        later**, and so opens without Word's *Compatibility Mode* title bar
        (CR-002 section 12.10). ``pkg.compatibility_mode`` changes it.
        """
        for name, value in NEW_DOCUMENT_COMPAT:
            self.set_compat_setting(name, value)


class WebSettingsPart(XmlPart["WebSettings"]):
    """``/word/webSettings.xml``. docx4j ``WebSettingsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.WebSettings"

    def __init__(self, part_name: PartName | str = "/word/webSettings.xml") -> None:
        """Build the web settings part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_WEBSETTINGS,
            Namespaces.WEB_SETTINGS,
            _w("webSettings"),
        )


class HeaderPart(XmlPart["Hdr"]):
    """``/word/headerN.xml``. docx4j ``HeaderPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Hdr"

    def __init__(self, part_name: PartName | str = "/word/header1.xml") -> None:
        """Build a header part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_HEADER,
            Namespaces.HEADER,
            _w("hdr"),
        )


class FooterPart(XmlPart["Ftr"]):
    """``/word/footerN.xml``. docx4j ``FooterPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Ftr"

    def __init__(self, part_name: PartName | str = "/word/footer1.xml") -> None:
        """Build a footer part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_FOOTER,
            Namespaces.FOOTER,
            _w("ftr"),
        )


class FootnotesPart(XmlPart["Footnotes"]):
    """``/word/footnotes.xml``. docx4j ``FootnotesPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Footnotes"

    def __init__(self, part_name: PartName | str = "/word/footnotes.xml") -> None:
        """Build the footnotes part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_FOOTNOTES,
            Namespaces.FOOTNOTES,
            _w("footnotes"),
        )


class EndnotesPart(XmlPart["Endnotes"]):
    """``/word/endnotes.xml``. docx4j ``EndnotesPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Endnotes"

    def __init__(self, part_name: PartName | str = "/word/endnotes.xml") -> None:
        """Build the endnotes part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_ENDNOTES,
            Namespaces.ENDNOTES,
            _w("endnotes"),
        )


class CommentsPart(XmlPart["Comments"]):
    """``/word/comments.xml``. docx4j ``CommentsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wml.Comments"

    def __init__(self, part_name: PartName | str = "/word/comments.xml") -> None:
        """Build the comments part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_COMMENTS,
            Namespaces.COMMENTS,
            _w("comments"),
        )


class CommentsExtendedPart(XmlPart[Any]):
    """``/word/commentsExtended.xml``, ``w15:commentsEx``. docx4j ``CommentsExtendedPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.w15.CommentsEx"

    def __init__(self, part_name: PartName | str = "/word/commentsExtended.xml") -> None:
        """Build the extended comments part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENDED,
            Namespaces.COMMENTS_EXTENDED,
            f"{{{_W15}}}commentsEx",
        )


class PeoplePart(XmlPart[Any]):
    """``/word/people.xml``, ``w15:people``. docx4j ``PeoplePart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.w15.People"

    def __init__(self, part_name: PartName | str = "/word/people.xml") -> None:
        """Build the people part."""
        super().__init__(
            part_name,
            ContentTypes.WORDPROCESSINGML_PEOPLE,
            Namespaces.OFFICE_2011_PEOPLE,
            f"{{{_W15}}}people",
        )


class KeyMapCustomizationsPart(XmlPart[Any]):
    """``/word/customizations.xml``, ``wne:tcg``. docx4j ``KeyMapCustomizationsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wne.Tcg"

    def __init__(self, part_name: PartName | str = "/word/customizations.xml") -> None:
        """Build the key-map customisations part."""
        super().__init__(
            part_name, ContentTypes.MS_WORD_KEYMAP, Namespaces.KEYMAP, f"{{{_WNE}}}tcg"
        )


class VbaDataPart(XmlPart[Any]):
    """``/word/vbaData.xml``, ``wne:vbaSuppData``. docx4j ``VbaDataPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.wne.VbaSuppData"

    def __init__(self, part_name: PartName | str = "/word/vbaData.xml") -> None:
        """Build the VBA data part."""
        super().__init__(
            part_name,
            ContentTypes.OFFICEDOCUMENT_VBA_DATA,
            Namespaces.VBA_DATA_WORD,
            f"{{{_WNE}}}vbaSuppData",
        )
