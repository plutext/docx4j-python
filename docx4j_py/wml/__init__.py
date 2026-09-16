from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_xsdata.models.datatype import XmlDateTime

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class BooleanDefaultFalse(Child):
    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class BooleanDefaultTrue(Child):
    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAttr(Child):
    class Meta:
        name = "CT_Attr"

    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAutoCaption(Child):
    class Meta:
        name = "CT_AutoCaption"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    caption: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCnf(Child):
    class Meta:
        name = "CT_Cnf"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 12,
            "pattern": r"[01]*",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColumn(Child):
    class Meta:
        name = "CT_Column"

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    space: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCompatSetting(Child):
    class Meta:
        name = "CT_CompatSetting"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTControl(Child):
    class Meta:
        name = "CT_Control"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shapeid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataBinding(Child):
    class Meta:
        name = "CT_DataBinding"

    prefix_mappings: None | str = field(
        default=None,
        metadata={
            "name": "prefixMappings",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    xpath: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    store_item_id: None | str = field(
        default=None,
        metadata={
            "name": "storeItemID",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDecimalNumber(Child):
    class Meta:
        name = "CT_DecimalNumber"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartName(Child):
    class Meta:
        name = "CT_DocPartName"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    decorated: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocVar(Child):
    class Meta:
        name = "CT_DocVar"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEmpty(Child):
    class Meta:
        name = "CT_Empty"


@dataclass(slots=True, kw_only=True)
class CTFFName(Child):
    class Meta:
        name = "CT_FFName"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_length": 65,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFitText(Child):
    class Meta:
        name = "CT_FitText"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFtnEdnRef(Child):
    class Meta:
        name = "CT_FtnEdnRef"

    custom_mark_follows: None | bool = field(
        default=None,
        metadata={
            "name": "customMarkFollows",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFtnEdnSepRef(Child):
    class Meta:
        name = "CT_FtnEdnSepRef"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGuid(Child):
    class Meta:
        name = "CT_Guid"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTKinsoku(Child):
    class Meta:
        name = "CT_Kinsoku"

    lang: None | bytes | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 2,
            "format": "base16",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLang(Child):
    class Meta:
        name = "CT_Lang"

    val: None | bytes | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 2,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLanguage(Child):
    class Meta:
        name = "CT_Language"

    val: None | bytes | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 2,
            "format": "base16",
        },
    )
    east_asia: None | bytes | str = field(
        default=None,
        metadata={
            "name": "eastAsia",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 2,
            "format": "base16",
        },
    )
    bidi: None | bytes | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 2,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLongHexNumber(Child):
    class Meta:
        name = "CT_LongHexNumber"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMacroName(Child):
    class Meta:
        name = "CT_MacroName"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_length": 33,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarkup(Child):
    class Meta:
        name = "CT_Markup"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPaperSource(Child):
    class Meta:
        name = "CT_PaperSource"

    first: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    other: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureBase(Child):
    class Meta:
        name = "CT_PictureBase"

    schemas_microsoft_com_vml_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "urn:schemas-microsoft-com:vml",
            "sequence": 1,
        },
    )
    schemas_microsoft_com_office_office_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "sequence": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPixelsMeasure(Child):
    class Meta:
        name = "CT_PixelsMeasure"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTReadingModeInkLockDown(Child):
    class Meta:
        name = "CT_ReadingModeInkLockDown"

    actual_pg: None | bool = field(
        default=None,
        metadata={
            "name": "actualPg",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    font_sz: None | int = field(
        default=None,
        metadata={
            "name": "fontSz",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRel(Child):
    class Meta:
        name = "CT_Rel"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSaveThroughXslt(Child):
    class Meta:
        name = "CT_SaveThroughXslt"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    solution_id: None | str = field(
        default=None,
        metadata={
            "name": "solutionID",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtListItem(Child):
    class Meta:
        name = "CT_SdtListItem"

    display_text: None | str = field(
        default=None,
        metadata={
            "name": "displayText",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    value: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtText(Child):
    class Meta:
        name = "CT_SdtText"

    multi_line: None | bool = field(
        default=None,
        metadata={
            "name": "multiLine",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeDefaults(Child):
    class Meta:
        name = "CT_ShapeDefaults"

    schemas_microsoft_com_office_office_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShortHexNumber(Child):
    class Meta:
        name = "CT_ShortHexNumber"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSignedHpsMeasure(Child):
    class Meta:
        name = "CT_SignedHpsMeasure"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSignedTwipsMeasure(Child):
    class Meta:
        name = "CT_SignedTwipsMeasure"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSmartTagType(Child):
    class Meta:
        name = "CT_SmartTagType"

    namespaceuri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    url: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTString(Child):
    class Meta:
        name = "CT_String"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStylePaneFilter(Child):
    class Meta:
        name = "CT_StylePaneFilter"

    all_styles: None | bool = field(
        default=None,
        metadata={
            "name": "allStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    custom_styles: None | bool = field(
        default=None,
        metadata={
            "name": "customStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    latent_styles: None | bool = field(
        default=None,
        metadata={
            "name": "latentStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    styles_in_use: None | bool = field(
        default=None,
        metadata={
            "name": "stylesInUse",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    heading_styles: None | bool = field(
        default=None,
        metadata={
            "name": "headingStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    numbering_styles: None | bool = field(
        default=None,
        metadata={
            "name": "numberingStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    table_styles: None | bool = field(
        default=None,
        metadata={
            "name": "tableStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    direct_formatting_on_runs: None | bool = field(
        default=None,
        metadata={
            "name": "directFormattingOnRuns",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    direct_formatting_on_paragraphs: None | bool = field(
        default=None,
        metadata={
            "name": "directFormattingOnParagraphs",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    direct_formatting_on_numbering: None | bool = field(
        default=None,
        metadata={
            "name": "directFormattingOnNumbering",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    direct_formatting_on_tables: None | bool = field(
        default=None,
        metadata={
            "name": "directFormattingOnTables",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    clear_formatting: None | bool = field(
        default=None,
        metadata={
            "name": "clearFormatting",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    top3_heading_styles: None | bool = field(
        default=None,
        metadata={
            "name": "top3HeadingStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    visible_styles: None | bool = field(
        default=None,
        metadata={
            "name": "visibleStyles",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    alternate_style_names: None | bool = field(
        default=None,
        metadata={
            "name": "alternateStyleNames",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextScale(Child):
    class Meta:
        name = "CT_TextScale"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "min_inclusive": 0,
            "max_inclusive": 600,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrackChangesView(Child):
    class Meta:
        name = "CT_TrackChangesView"

    markup: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    comments: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    ins_del: None | bool = field(
        default=None,
        metadata={
            "name": "insDel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    formatting: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    ink_annotations: None | bool = field(
        default=None,
        metadata={
            "name": "inkAnnotations",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTwipsMeasure(Child):
    class Meta:
        name = "CT_TwipsMeasure"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUcharHexNumber(Child):
    class Meta:
        name = "CT_UcharHexNumber"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWritingStyle(Child):
    class Meta:
        name = "CT_WritingStyle"

    lang: None | bytes | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 2,
            "format": "base16",
        },
    )
    vendor_id: None | int = field(
        default=None,
        metadata={
            "name": "vendorID",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dll_version: None | int = field(
        default=None,
        metadata={
            "name": "dllVersion",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    nl_check: None | bool = field(
        default=None,
        metadata={
            "name": "nlCheck",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    check_style: None | bool = field(
        default=None,
        metadata={
            "name": "checkStyle",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    app_name: None | str = field(
        default=None,
        metadata={
            "name": "appName",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtCustomXmlPrPlaceholder(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtDocPartCategoryName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtDocPartPrDescription(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtDocPartPrStyle(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtEdnPropsNumStart(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFfddlistDefault(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFfddlistListEntry(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFfddlistResult(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFftextInputDefault(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFftextInputFormat(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFftextInputMaxLength(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtFontFamilyVal(Enum):
    DECORATIVE = "decorative"
    MODERN = "modern"
    ROMAN = "roman"
    SCRIPT = "script"
    SWISS = "swiss"
    AUTO = "auto"


@dataclass(slots=True, kw_only=True)
class CtFrameName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFrameSz(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFramesetSz(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFtnPropsNumStart(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlLegacy(Child):
    class Meta:
        global_type = False

    legacy: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    legacy_space: None | int = field(
        default=None,
        metadata={
            "name": "legacySpace",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    legacy_indent: None | int = field(
        default=None,
        metadata={
            "name": "legacyIndent",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlLvlPicBulletId(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlLvlRestart(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlLvlText(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    null: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlPStyle(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlStart(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtLvlSuffVal(Enum):
    TAB = "tab"
    SPACE = "space"
    NOTHING = "nothing"


@dataclass(slots=True, kw_only=True)
class CtMailMergeActiveRecord(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMailMergeAddressFieldName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMailMergeCheckErrors(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMailMergeConnectString(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMailMergeMailSubject(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMailMergeQuery(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeAnnotationRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeCommentReference(Child):
    class Meta:
        global_type = False

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeContinuationSeparator(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeCr(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeDayLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeDayShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeEndnoteRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeFootnoteRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeLastRenderedPageBreak(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeMonthLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeMonthShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeNoBreakHyphen(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangePgNum(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeSeparator(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeSoftHyphen(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeSym(Child):
    class Meta:
        global_type = False

    font: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    char: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeTab(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeYearLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangeYearShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtOdsoFieldMapDataColumn(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtOdsoFieldMapDataMappedName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtOdsoFieldMapDataName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtOdsoColDelim(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtOdsoTable(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtOdsoUdl(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseDivId(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseInd(Child):
    class Meta:
        global_type = False

    left: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left_chars: None | int = field(
        default=None,
        metadata={
            "name": "leftChars",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right_chars: None | int = field(
        default=None,
        metadata={
            "name": "rightChars",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    start: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    start_chars: None | int = field(
        default=None,
        metadata={
            "name": "startChars",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    end: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    end_chars: None | int = field(
        default=None,
        metadata={
            "name": "endChars",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hanging: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hanging_chars: None | int = field(
        default=None,
        metadata={
            "name": "hangingChars",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    first_line: None | int = field(
        default=None,
        metadata={
            "name": "firstLine",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    first_line_chars: None | int = field(
        default=None,
        metadata={
            "name": "firstLineChars",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseNumPrIlvl(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseNumPrNumId(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseOutlineLvl(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBasePStyle(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtPprBaseTextAlignmentVal(Enum):
    TOP = "top"
    CENTER = "center"
    BASELINE = "baseline"
    BOTTOM = "bottom"
    AUTO = "auto"


class CtPermStartEdGrp(Enum):
    NONE = "none"
    EVERYONE = "everyone"
    ADMINISTRATORS = "administrators"
    CONTRIBUTORS = "contributors"
    EDITORS = "editors"
    OWNERS = "owners"
    CURRENT = "current"


@dataclass(slots=True, kw_only=True)
class CtPlaceholderDocPart(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtRAnnotationRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRCommentReference(Child):
    class Meta:
        global_type = False

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtRContinuationSeparator(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRCr(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRDayLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRDayShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtREndnoteRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRFootnoteRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRLastRenderedPageBreak(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRMonthLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRMonthShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRNoBreakHyphen(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRPgNum(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRSeparator(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRSoftHyphen(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRSym(Child):
    class Meta:
        global_type = False

    font: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    char: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtRTab(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRYearLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRYearShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtRecipientDataColumn(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSdtDateDateFormat(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSdtDocPartDocPartCategory(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSdtDocPartDocPartGallery(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSdtPrAlias(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSdtPrBibliography(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSdtPrCitation(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSdtPrEquation(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSdtPrGroup(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSdtPrPicture(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSdtPrRichText(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSectPrBasePgMar(Child):
    class Meta:
        global_type = False

    top: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    header: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    footer: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    gutter: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtSectPrBaseTypeVal(Enum):
    NEXT_PAGE = "nextPage"
    NEXT_COLUMN = "nextColumn"
    CONTINUOUS = "continuous"
    EVEN_PAGE = "evenPage"
    ODD_PAGE = "oddPage"


@dataclass(slots=True, kw_only=True)
class CtSectPrPgMar(Child):
    class Meta:
        global_type = False

    top: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    header: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    footer: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    gutter: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtSectPrTypeVal(Enum):
    NEXT_PAGE = "nextPage"
    NEXT_COLUMN = "nextColumn"
    CONTINUOUS = "continuous"
    EVEN_PAGE = "evenPage"
    ODD_PAGE = "oddPage"


@dataclass(slots=True, kw_only=True)
class CtSettingsAttachedSchema(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsBookFoldPrintingSheets(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsClickAndTypeStyle(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsConsecutiveHyphenLimit(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsDecimalSymbol(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsDefaultTableStyle(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsDisplayHorizontalDrawingGridEvery(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsDisplayVerticalDrawingGridEvery(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsForceUpgrade(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtSettingsListSeparator(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSettingsSummaryLength(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTblPrBaseTblStyle(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTblPrBaseTblStyleColBandSize(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTblPrBaseTblStyleRowBandSize(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtTblWidthType(Enum):
    NIL = "nil"
    PCT = "pct"
    DXA = "dxa"
    AUTO = "auto"


@dataclass(slots=True, kw_only=True)
class CtTcPrInnerGridSpan(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class CtTcPrInnerHMergeVal(Enum):
    CONTINUE = "continue"
    RESTART = "restart"


class CtTcPrInnerVMergeVal(Enum):
    CONTINUE = "continue"
    RESTART = "restart"


class CtTextDirectionVal(Enum):
    LR_TB = "lrTb"
    TB_RL = "tbRl"
    BT_LR = "btLr"
    LR_TB_V = "lrTbV"
    TB_RL_V = "tbRlV"
    TB_LR_V = "tbLrV"


@dataclass(slots=True, kw_only=True)
class CtTrPrBaseDivId(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTrPrBaseGridAfter(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTrPrBaseGridBefore(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtWebSettingsEncoding(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtWebSettingsPixelsPerInch(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class FontPanose(Child):
    class Meta:
        name = "CT_Panose"

    val: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 10,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class FontSig(Child):
    class Meta:
        name = "CT_FontSig"

    usb0: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    usb1: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    usb2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    usb3: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    csb0: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    csb1: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class HdrFtrRef(Enum):
    EVEN = "even"
    DEFAULT = "default"
    FIRST = "first"


@dataclass(slots=True, kw_only=True)
class HpsMeasure(Child):
    class Meta:
        name = "CT_HpsMeasure"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Id(Child):
    class Meta:
        name = "id"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class JcEnumeration(Enum):
    LEFT = "left"
    START = "start"
    CENTER = "center"
    RIGHT = "right"
    END = "end"
    BOTH = "both"
    MEDIUM_KASHIDA = "mediumKashida"
    DISTRIBUTE = "distribute"
    NUM_TAB = "numTab"
    HIGH_KASHIDA = "highKashida"
    LOW_KASHIDA = "lowKashida"
    THAI_DISTRIBUTE = "thaiDistribute"


class NumberFormat(Enum):
    DECIMAL = "decimal"
    UPPER_ROMAN = "upperRoman"
    LOWER_ROMAN = "lowerRoman"
    UPPER_LETTER = "upperLetter"
    LOWER_LETTER = "lowerLetter"
    ORDINAL = "ordinal"
    CARDINAL_TEXT = "cardinalText"
    ORDINAL_TEXT = "ordinalText"
    HEX = "hex"
    CHICAGO = "chicago"
    IDEOGRAPH_DIGITAL = "ideographDigital"
    JAPANESE_COUNTING = "japaneseCounting"
    AIUEO = "aiueo"
    IROHA = "iroha"
    DECIMAL_FULL_WIDTH = "decimalFullWidth"
    DECIMAL_HALF_WIDTH = "decimalHalfWidth"
    JAPANESE_LEGAL = "japaneseLegal"
    JAPANESE_DIGITAL_TEN_THOUSAND = "japaneseDigitalTenThousand"
    DECIMAL_ENCLOSED_CIRCLE = "decimalEnclosedCircle"
    DECIMAL_FULL_WIDTH2 = "decimalFullWidth2"
    AIUEO_FULL_WIDTH = "aiueoFullWidth"
    IROHA_FULL_WIDTH = "irohaFullWidth"
    DECIMAL_ZERO = "decimalZero"
    BULLET = "bullet"
    GANADA = "ganada"
    CHOSUNG = "chosung"
    DECIMAL_ENCLOSED_FULLSTOP = "decimalEnclosedFullstop"
    DECIMAL_ENCLOSED_PAREN = "decimalEnclosedParen"
    DECIMAL_ENCLOSED_CIRCLE_CHINESE = "decimalEnclosedCircleChinese"
    IDEOGRAPH_ENCLOSED_CIRCLE = "ideographEnclosedCircle"
    IDEOGRAPH_TRADITIONAL = "ideographTraditional"
    IDEOGRAPH_ZODIAC = "ideographZodiac"
    IDEOGRAPH_ZODIAC_TRADITIONAL = "ideographZodiacTraditional"
    TAIWANESE_COUNTING = "taiwaneseCounting"
    IDEOGRAPH_LEGAL_TRADITIONAL = "ideographLegalTraditional"
    TAIWANESE_COUNTING_THOUSAND = "taiwaneseCountingThousand"
    TAIWANESE_DIGITAL = "taiwaneseDigital"
    CHINESE_COUNTING = "chineseCounting"
    CHINESE_LEGAL_SIMPLIFIED = "chineseLegalSimplified"
    CHINESE_COUNTING_THOUSAND = "chineseCountingThousand"
    KOREAN_DIGITAL = "koreanDigital"
    KOREAN_COUNTING = "koreanCounting"
    KOREAN_LEGAL = "koreanLegal"
    KOREAN_DIGITAL2 = "koreanDigital2"
    VIETNAMESE_COUNTING = "vietnameseCounting"
    RUSSIAN_LOWER = "russianLower"
    RUSSIAN_UPPER = "russianUpper"
    NONE = "none"
    NUMBER_IN_DASH = "numberInDash"
    HEBREW1 = "hebrew1"
    HEBREW2 = "hebrew2"
    ARABIC_ALPHA = "arabicAlpha"
    ARABIC_ABJAD = "arabicAbjad"
    HINDI_VOWELS = "hindiVowels"
    HINDI_CONSONANTS = "hindiConsonants"
    HINDI_NUMBERS = "hindiNumbers"
    HINDI_COUNTING = "hindiCounting"
    THAI_LETTERS = "thaiLetters"
    THAI_NUMBERS = "thaiNumbers"
    THAI_COUNTING = "thaiCounting"


@dataclass(slots=True, kw_only=True)
class RStyle(Child):
    class Meta:
        name = "rStyle"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class STAlgClass(Enum):
    HASH = "hash"


class STAlgType(Enum):
    TYPE_ANY = "typeAny"


class STAnnotationVMerge(Enum):
    CONT = "cont"
    REST = "rest"


class STBorder(Enum):
    NIL = "nil"
    NONE = "none"
    SINGLE = "single"
    THICK = "thick"
    DOUBLE = "double"
    DOTTED = "dotted"
    DASHED = "dashed"
    DOT_DASH = "dotDash"
    DOT_DOT_DASH = "dotDotDash"
    TRIPLE = "triple"
    THIN_THICK_SMALL_GAP = "thinThickSmallGap"
    THICK_THIN_SMALL_GAP = "thickThinSmallGap"
    THIN_THICK_THIN_SMALL_GAP = "thinThickThinSmallGap"
    THIN_THICK_MEDIUM_GAP = "thinThickMediumGap"
    THICK_THIN_MEDIUM_GAP = "thickThinMediumGap"
    THIN_THICK_THIN_MEDIUM_GAP = "thinThickThinMediumGap"
    THIN_THICK_LARGE_GAP = "thinThickLargeGap"
    THICK_THIN_LARGE_GAP = "thickThinLargeGap"
    THIN_THICK_THIN_LARGE_GAP = "thinThickThinLargeGap"
    WAVE = "wave"
    DOUBLE_WAVE = "doubleWave"
    DASH_SMALL_GAP = "dashSmallGap"
    DASH_DOT_STROKED = "dashDotStroked"
    THREE_DEMBOSS = "threeDEmboss"
    THREE_DENGRAVE = "threeDEngrave"
    OUTSET = "outset"
    INSET = "inset"
    APPLES = "apples"
    ARCHED_SCALLOPS = "archedScallops"
    BABY_PACIFIER = "babyPacifier"
    BABY_RATTLE = "babyRattle"
    BALLOONS3_COLORS = "balloons3Colors"
    BALLOONS_HOT_AIR = "balloonsHotAir"
    BASIC_BLACK_DASHES = "basicBlackDashes"
    BASIC_BLACK_DOTS = "basicBlackDots"
    BASIC_BLACK_SQUARES = "basicBlackSquares"
    BASIC_THIN_LINES = "basicThinLines"
    BASIC_WHITE_DASHES = "basicWhiteDashes"
    BASIC_WHITE_DOTS = "basicWhiteDots"
    BASIC_WHITE_SQUARES = "basicWhiteSquares"
    BASIC_WIDE_INLINE = "basicWideInline"
    BASIC_WIDE_MIDLINE = "basicWideMidline"
    BASIC_WIDE_OUTLINE = "basicWideOutline"
    BATS = "bats"
    BIRDS = "birds"
    BIRDS_FLIGHT = "birdsFlight"
    CABINS = "cabins"
    CAKE_SLICE = "cakeSlice"
    CANDY_CORN = "candyCorn"
    CELTIC_KNOTWORK = "celticKnotwork"
    CERTIFICATE_BANNER = "certificateBanner"
    CHAIN_LINK = "chainLink"
    CHAMPAGNE_BOTTLE = "champagneBottle"
    CHECKED_BAR_BLACK = "checkedBarBlack"
    CHECKED_BAR_COLOR = "checkedBarColor"
    CHECKERED = "checkered"
    CHRISTMAS_TREE = "christmasTree"
    CIRCLES_LINES = "circlesLines"
    CIRCLES_RECTANGLES = "circlesRectangles"
    CLASSICAL_WAVE = "classicalWave"
    CLOCKS = "clocks"
    COMPASS = "compass"
    CONFETTI = "confetti"
    CONFETTI_GRAYS = "confettiGrays"
    CONFETTI_OUTLINE = "confettiOutline"
    CONFETTI_STREAMERS = "confettiStreamers"
    CONFETTI_WHITE = "confettiWhite"
    CORNER_TRIANGLES = "cornerTriangles"
    COUPON_CUTOUT_DASHES = "couponCutoutDashes"
    COUPON_CUTOUT_DOTS = "couponCutoutDots"
    CRAZY_MAZE = "crazyMaze"
    CREATURES_BUTTERFLY = "creaturesButterfly"
    CREATURES_FISH = "creaturesFish"
    CREATURES_INSECTS = "creaturesInsects"
    CREATURES_LADY_BUG = "creaturesLadyBug"
    CROSS_STITCH = "crossStitch"
    CUP = "cup"
    DECO_ARCH = "decoArch"
    DECO_ARCH_COLOR = "decoArchColor"
    DECO_BLOCKS = "decoBlocks"
    DIAMONDS_GRAY = "diamondsGray"
    DOUBLE_D = "doubleD"
    DOUBLE_DIAMONDS = "doubleDiamonds"
    EARTH1 = "earth1"
    EARTH2 = "earth2"
    ECLIPSING_SQUARES1 = "eclipsingSquares1"
    ECLIPSING_SQUARES2 = "eclipsingSquares2"
    EGGS_BLACK = "eggsBlack"
    FANS = "fans"
    FILM = "film"
    FIRECRACKERS = "firecrackers"
    FLOWERS_BLOCK_PRINT = "flowersBlockPrint"
    FLOWERS_DAISIES = "flowersDaisies"
    FLOWERS_MODERN1 = "flowersModern1"
    FLOWERS_MODERN2 = "flowersModern2"
    FLOWERS_PANSY = "flowersPansy"
    FLOWERS_RED_ROSE = "flowersRedRose"
    FLOWERS_ROSES = "flowersRoses"
    FLOWERS_TEACUP = "flowersTeacup"
    FLOWERS_TINY = "flowersTiny"
    GEMS = "gems"
    GINGERBREAD_MAN = "gingerbreadMan"
    GRADIENT = "gradient"
    HANDMADE1 = "handmade1"
    HANDMADE2 = "handmade2"
    HEART_BALLOON = "heartBalloon"
    HEART_GRAY = "heartGray"
    HEARTS = "hearts"
    HEEBIE_JEEBIES = "heebieJeebies"
    HOLLY = "holly"
    HOUSE_FUNKY = "houseFunky"
    HYPNOTIC = "hypnotic"
    ICE_CREAM_CONES = "iceCreamCones"
    LIGHT_BULB = "lightBulb"
    LIGHTNING1 = "lightning1"
    LIGHTNING2 = "lightning2"
    MAP_PINS = "mapPins"
    MAPLE_LEAF = "mapleLeaf"
    MAPLE_MUFFINS = "mapleMuffins"
    MARQUEE = "marquee"
    MARQUEE_TOOTHED = "marqueeToothed"
    MOONS = "moons"
    MOSAIC = "mosaic"
    MUSIC_NOTES = "musicNotes"
    NORTHWEST = "northwest"
    OVALS = "ovals"
    PACKAGES = "packages"
    PALMS_BLACK = "palmsBlack"
    PALMS_COLOR = "palmsColor"
    PAPER_CLIPS = "paperClips"
    PAPYRUS = "papyrus"
    PARTY_FAVOR = "partyFavor"
    PARTY_GLASS = "partyGlass"
    PENCILS = "pencils"
    PEOPLE = "people"
    PEOPLE_WAVING = "peopleWaving"
    PEOPLE_HATS = "peopleHats"
    POINSETTIAS = "poinsettias"
    POSTAGE_STAMP = "postageStamp"
    PUMPKIN1 = "pumpkin1"
    PUSH_PIN_NOTE2 = "pushPinNote2"
    PUSH_PIN_NOTE1 = "pushPinNote1"
    PYRAMIDS = "pyramids"
    PYRAMIDS_ABOVE = "pyramidsAbove"
    QUADRANTS = "quadrants"
    RINGS = "rings"
    SAFARI = "safari"
    SAWTOOTH = "sawtooth"
    SAWTOOTH_GRAY = "sawtoothGray"
    SCARED_CAT = "scaredCat"
    SEATTLE = "seattle"
    SHADOWED_SQUARES = "shadowedSquares"
    SHARKS_TEETH = "sharksTeeth"
    SHOREBIRD_TRACKS = "shorebirdTracks"
    SKYROCKET = "skyrocket"
    SNOWFLAKE_FANCY = "snowflakeFancy"
    SNOWFLAKES = "snowflakes"
    SOMBRERO = "sombrero"
    SOUTHWEST = "southwest"
    STARS = "stars"
    STARS_TOP = "starsTop"
    STARS3D = "stars3d"
    STARS_BLACK = "starsBlack"
    STARS_SHADOWED = "starsShadowed"
    SUN = "sun"
    SWIRLIGIG = "swirligig"
    TORN_PAPER = "tornPaper"
    TORN_PAPER_BLACK = "tornPaperBlack"
    TREES = "trees"
    TRIANGLE_PARTY = "triangleParty"
    TRIANGLES = "triangles"
    TRIBAL1 = "tribal1"
    TRIBAL2 = "tribal2"
    TRIBAL3 = "tribal3"
    TRIBAL4 = "tribal4"
    TRIBAL5 = "tribal5"
    TRIBAL6 = "tribal6"
    TWISTED_LINES1 = "twistedLines1"
    TWISTED_LINES2 = "twistedLines2"
    VINE = "vine"
    WAVELINE = "waveline"
    WEAVING_ANGLES = "weavingAngles"
    WEAVING_BRAID = "weavingBraid"
    WEAVING_RIBBON = "weavingRibbon"
    WEAVING_STRIPS = "weavingStrips"
    WHITE_FLOWERS = "whiteFlowers"
    WOODWORK = "woodwork"
    X_ILLUSIONS = "xIllusions"
    ZANY_TRIANGLES = "zanyTriangles"
    ZIG_ZAG = "zigZag"
    ZIG_ZAG_STITCH = "zigZagStitch"


class STBrClear(Enum):
    NONE = "none"
    LEFT = "left"
    RIGHT = "right"
    ALL = "all"


class STBrType(Enum):
    PAGE = "page"
    COLUMN = "column"
    TEXT_WRAPPING = "textWrapping"


class STCalendarType(Enum):
    GREGORIAN = "gregorian"
    HIJRI = "hijri"
    HEBREW = "hebrew"
    TAIWAN = "taiwan"
    JAPAN = "japan"
    THAI = "thai"
    KOREA = "korea"
    SAKA = "saka"
    GREGORIAN_XLIT_ENGLISH = "gregorianXlitEnglish"
    GREGORIAN_XLIT_FRENCH = "gregorianXlitFrench"


class STCaptionPos(Enum):
    ABOVE = "above"
    BELOW = "below"
    LEFT = "left"
    RIGHT = "right"


class STChapterSep(Enum):
    HYPHEN = "hyphen"
    PERIOD = "period"
    COLON = "colon"
    EM_DASH = "emDash"
    EN_DASH = "enDash"


class STCharacterSpacing(Enum):
    DO_NOT_COMPRESS = "doNotCompress"
    COMPRESS_PUNCTUATION = "compressPunctuation"
    COMPRESS_PUNCTUATION_AND_JAPANESE_KANA = "compressPunctuationAndJapaneseKana"


class STColorSchemeIndex(Enum):
    DARK1 = "dark1"
    LIGHT1 = "light1"
    DARK2 = "dark2"
    LIGHT2 = "light2"
    ACCENT1 = "accent1"
    ACCENT2 = "accent2"
    ACCENT3 = "accent3"
    ACCENT4 = "accent4"
    ACCENT5 = "accent5"
    ACCENT6 = "accent6"
    HYPERLINK = "hyperlink"
    FOLLOWED_HYPERLINK = "followedHyperlink"


class STCombineBrackets(Enum):
    NONE = "none"
    ROUND = "round"
    SQUARE = "square"
    ANGLE = "angle"
    CURLY = "curly"


class STCryptProv(Enum):
    RSA_AES = "rsaAES"
    RSA_FULL = "rsaFull"


class STDirection(Enum):
    LTR = "ltr"
    RTL = "rtl"


class STDisplacedByCustomXml(Enum):
    NEXT = "next"
    PREV = "prev"


class STDocGrid(Enum):
    DEFAULT = "default"
    LINES = "lines"
    LINES_AND_CHARS = "linesAndChars"
    SNAP_TO_CHARS = "snapToChars"


class STDocPartBehavior(Enum):
    CONTENT = "content"
    P = "p"
    PG = "pg"


class STDocPartGallery(Enum):
    PLACEHOLDER = "placeholder"
    ANY = "any"
    DEFAULT = "default"
    DOC_PARTS = "docParts"
    COVER_PG = "coverPg"
    EQ = "eq"
    FTRS = "ftrs"
    HDRS = "hdrs"
    PG_NUM = "pgNum"
    TBLS = "tbls"
    WATERMARKS = "watermarks"
    AUTO_TXT = "autoTxt"
    TXT_BOX = "txtBox"
    PG_NUM_T = "pgNumT"
    PG_NUM_B = "pgNumB"
    PG_NUM_MARGINS = "pgNumMargins"
    TBL_OF_CONTENTS = "tblOfContents"
    BIB = "bib"
    CUST_QUICK_PARTS = "custQuickParts"
    CUST_COVER_PG = "custCoverPg"
    CUST_EQ = "custEq"
    CUST_FTRS = "custFtrs"
    CUST_HDRS = "custHdrs"
    CUST_PG_NUM = "custPgNum"
    CUST_TBLS = "custTbls"
    CUST_WATERMARKS = "custWatermarks"
    CUST_AUTO_TXT = "custAutoTxt"
    CUST_TXT_BOX = "custTxtBox"
    CUST_PG_NUM_T = "custPgNumT"
    CUST_PG_NUM_B = "custPgNumB"
    CUST_PG_NUM_MARGINS = "custPgNumMargins"
    CUST_TBL_OF_CONTENTS = "custTblOfContents"
    CUST_BIB = "custBib"
    CUSTOM1 = "custom1"
    CUSTOM2 = "custom2"
    CUSTOM3 = "custom3"
    CUSTOM4 = "custom4"
    CUSTOM5 = "custom5"


class STDocPartType(Enum):
    NONE = "none"
    NORMAL = "normal"
    AUTO_EXP = "autoExp"
    TOOLBAR = "toolbar"
    SPELLER = "speller"
    FORM_FLD = "formFld"
    BB_PLC_HDR = "bbPlcHdr"


class STDocProtect(Enum):
    NONE = "none"
    READ_ONLY = "readOnly"
    COMMENTS = "comments"
    TRACKED_CHANGES = "trackedChanges"
    FORMS = "forms"


class STDocType(Enum):
    NOT_SPECIFIED = "notSpecified"
    LETTER = "letter"
    E_MAIL = "eMail"


class STDropCap(Enum):
    NONE = "none"
    DROP = "drop"
    MARGIN = "margin"


class STEdnPos(Enum):
    SECT_END = "sectEnd"
    DOC_END = "docEnd"


class STEm(Enum):
    NONE = "none"
    DOT = "dot"
    COMMA = "comma"
    CIRCLE = "circle"
    UNDER_DOT = "underDot"


class STFFTextType(Enum):
    REGULAR = "regular"
    NUMBER = "number"
    DATE = "date"
    CURRENT_TIME = "currentTime"
    CURRENT_DATE = "currentDate"
    CALCULATED = "calculated"


class STFldCharType(Enum):
    BEGIN = "begin"
    SEPARATE = "separate"
    END = "end"


class STFrameLayout(Enum):
    ROWS = "rows"
    COLS = "cols"
    NONE = "none"


class STFrameScrollbar(Enum):
    ON = "on"
    OFF = "off"
    AUTO = "auto"


class STFtnEdn(Enum):
    NORMAL = "normal"
    SEPARATOR = "separator"
    CONTINUATION_SEPARATOR = "continuationSeparator"
    CONTINUATION_NOTICE = "continuationNotice"


class STFtnPos(Enum):
    PAGE_BOTTOM = "pageBottom"
    BENEATH_TEXT = "beneathText"
    SECT_END = "sectEnd"
    DOC_END = "docEnd"


class STHAnchor(Enum):
    TEXT = "text"
    MARGIN = "margin"
    PAGE = "page"


class STHeightRule(Enum):
    AUTO = "auto"
    EXACT = "exact"
    AT_LEAST = "atLeast"


class STHexColorAuto(Enum):
    AUTO = "auto"


class STHint(Enum):
    DEFAULT = "default"
    EAST_ASIA = "eastAsia"
    CS = "cs"


class STInfoTextType(Enum):
    TEXT = "text"
    AUTO_TEXT = "autoText"


class STLineNumberRestart(Enum):
    NEW_PAGE = "newPage"
    NEW_SECTION = "newSection"
    CONTINUOUS = "continuous"


class STLineSpacingRule(Enum):
    AUTO = "auto"
    EXACT = "exact"
    AT_LEAST = "atLeast"


class STLock(Enum):
    SDT_LOCKED = "sdtLocked"
    CONTENT_LOCKED = "contentLocked"
    UNLOCKED = "unlocked"
    SDT_CONTENT_LOCKED = "sdtContentLocked"


class STMailMergeDataType(Enum):
    TEXT_FILE = "textFile"
    DATABASE = "database"
    SPREADSHEET = "spreadsheet"
    QUERY = "query"
    ODBC = "odbc"
    NATIVE = "native"


class STMailMergeDest(Enum):
    NEW_DOCUMENT = "newDocument"
    PRINTER = "printer"
    EMAIL = "email"
    FAX = "fax"


class STMailMergeDocType(Enum):
    CATALOG = "catalog"
    ENVELOPES = "envelopes"
    MAILING_LABELS = "mailingLabels"
    FORM_LETTERS = "formLetters"
    EMAIL = "email"
    FAX = "fax"


class STMailMergeOdsoFMDFieldType(Enum):
    NULL = "null"
    DB_COLUMN = "dbColumn"


class STMailMergeSourceType(Enum):
    DATABASE = "database"
    ADDRESS_BOOK = "addressBook"
    DOCUMENT1 = "document1"
    DOCUMENT2 = "document2"
    TEXT = "text"
    EMAIL = "email"
    NATIVE = "native"
    LEGACY = "legacy"
    MASTER = "master"


class STPTabAlignment(Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class STPTabLeader(Enum):
    NONE = "none"
    DOT = "dot"
    HYPHEN = "hyphen"
    UNDERSCORE = "underscore"
    MIDDLE_DOT = "middleDot"


class STPTabRelativeTo(Enum):
    MARGIN = "margin"
    INDENT = "indent"


class STPageBorderDisplay(Enum):
    ALL_PAGES = "allPages"
    FIRST_PAGE = "firstPage"
    NOT_FIRST_PAGE = "notFirstPage"


class STPageBorderOffset(Enum):
    PAGE = "page"
    TEXT = "text"


class STPageBorderZOrder(Enum):
    FRONT = "front"
    BACK = "back"


class STPageOrientation(Enum):
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"


class STPitch(Enum):
    FIXED = "fixed"
    VARIABLE = "variable"
    DEFAULT = "default"


class STProof(Enum):
    CLEAN = "clean"
    DIRTY = "dirty"


class STRestartNumber(Enum):
    CONTINUOUS = "continuous"
    EACH_SECT = "eachSect"
    EACH_PAGE = "eachPage"


class STRubyAlign(Enum):
    CENTER = "center"
    DISTRIBUTE_LETTER = "distributeLetter"
    DISTRIBUTE_SPACE = "distributeSpace"
    LEFT = "left"
    RIGHT = "right"
    RIGHT_VERTICAL = "rightVertical"


class STSdtDateMappingType(Enum):
    TEXT = "text"
    DATE = "date"
    DATE_TIME = "dateTime"


class STShd(Enum):
    NIL = "nil"
    CLEAR = "clear"
    SOLID = "solid"
    HORZ_STRIPE = "horzStripe"
    VERT_STRIPE = "vertStripe"
    REVERSE_DIAG_STRIPE = "reverseDiagStripe"
    DIAG_STRIPE = "diagStripe"
    HORZ_CROSS = "horzCross"
    DIAG_CROSS = "diagCross"
    THIN_HORZ_STRIPE = "thinHorzStripe"
    THIN_VERT_STRIPE = "thinVertStripe"
    THIN_REVERSE_DIAG_STRIPE = "thinReverseDiagStripe"
    THIN_DIAG_STRIPE = "thinDiagStripe"
    THIN_HORZ_CROSS = "thinHorzCross"
    THIN_DIAG_CROSS = "thinDiagCross"
    PCT5 = "pct5"
    PCT10 = "pct10"
    PCT12 = "pct12"
    PCT15 = "pct15"
    PCT20 = "pct20"
    PCT25 = "pct25"
    PCT30 = "pct30"
    PCT35 = "pct35"
    PCT37 = "pct37"
    PCT40 = "pct40"
    PCT45 = "pct45"
    PCT50 = "pct50"
    PCT55 = "pct55"
    PCT60 = "pct60"
    PCT62 = "pct62"
    PCT65 = "pct65"
    PCT70 = "pct70"
    PCT75 = "pct75"
    PCT80 = "pct80"
    PCT85 = "pct85"
    PCT87 = "pct87"
    PCT90 = "pct90"
    PCT95 = "pct95"


class STTabJc(Enum):
    CLEAR = "clear"
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    DECIMAL = "decimal"
    BAR = "bar"
    NUM = "num"


class STTabTlc(Enum):
    NONE = "none"
    DOT = "dot"
    HYPHEN = "hyphen"
    UNDERSCORE = "underscore"
    HEAVY = "heavy"
    MIDDLE_DOT = "middleDot"


class STTblLayoutType(Enum):
    FIXED = "fixed"
    AUTOFIT = "autofit"


class STTblOverlap(Enum):
    NEVER = "never"
    OVERLAP = "overlap"


class STTblStyleOverrideType(Enum):
    WHOLE_TABLE = "wholeTable"
    FIRST_ROW = "firstRow"
    LAST_ROW = "lastRow"
    FIRST_COL = "firstCol"
    LAST_COL = "lastCol"
    BAND1_VERT = "band1Vert"
    BAND2_VERT = "band2Vert"
    BAND1_HORZ = "band1Horz"
    BAND2_HORZ = "band2Horz"
    NE_CELL = "neCell"
    NW_CELL = "nwCell"
    SE_CELL = "seCell"
    SW_CELL = "swCell"


class STTextEffect(Enum):
    BLINK_BACKGROUND = "blinkBackground"
    LIGHTS = "lights"
    ANTS_BLACK = "antsBlack"
    ANTS_RED = "antsRed"
    SHIMMER = "shimmer"
    SPARKLE = "sparkle"
    NONE = "none"


class STTextboxTightWrap(Enum):
    NONE = "none"
    ALL_LINES = "allLines"
    FIRST_AND_LAST_LINE = "firstAndLastLine"
    FIRST_LINE_ONLY = "firstLineOnly"
    LAST_LINE_ONLY = "lastLineOnly"


class STTheme(Enum):
    MAJOR_EAST_ASIA = "majorEastAsia"
    MAJOR_BIDI = "majorBidi"
    MAJOR_ASCII = "majorAscii"
    MAJOR_HANSI = "majorHAnsi"
    MINOR_EAST_ASIA = "minorEastAsia"
    MINOR_BIDI = "minorBidi"
    MINOR_ASCII = "minorAscii"
    MINOR_HANSI = "minorHAnsi"


class STThemeColor(Enum):
    DARK1 = "dark1"
    LIGHT1 = "light1"
    DARK2 = "dark2"
    LIGHT2 = "light2"
    ACCENT1 = "accent1"
    ACCENT2 = "accent2"
    ACCENT3 = "accent3"
    ACCENT4 = "accent4"
    ACCENT5 = "accent5"
    ACCENT6 = "accent6"
    HYPERLINK = "hyperlink"
    FOLLOWED_HYPERLINK = "followedHyperlink"
    NONE = "none"
    BACKGROUND1 = "background1"
    TEXT1 = "text1"
    BACKGROUND2 = "background2"
    TEXT2 = "text2"


class STVAnchor(Enum):
    TEXT = "text"
    MARGIN = "margin"
    PAGE = "page"


class STVerticalAlignRun(Enum):
    BASELINE = "baseline"
    SUPERSCRIPT = "superscript"
    SUBSCRIPT = "subscript"


class STVerticalJc(Enum):
    TOP = "top"
    CENTER = "center"
    BOTH = "both"
    BOTTOM = "bottom"


class STView(Enum):
    NONE = "none"
    PRINT = "print"
    OUTLINE = "outline"
    MASTER_PAGES = "masterPages"
    NORMAL = "normal"
    WEB = "web"


class STWrap(Enum):
    AUTO = "auto"
    NOT_BESIDE = "notBeside"
    AROUND = "around"
    TIGHT = "tight"
    THROUGH = "through"
    NONE = "none"


class STXAlign(Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    INSIDE = "inside"
    OUTSIDE = "outside"


class STYAlign(Enum):
    INLINE = "inline"
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"
    INSIDE = "inside"
    OUTSIDE = "outside"


class STZoom(Enum):
    NONE = "none"
    FULL_PAGE = "fullPage"
    BEST_FIT = "bestFit"
    TEXT_FIT = "textFit"


class StOnOff(Enum):
    TRUE = "true"
    FALSE = "false"
    ON = "on"
    OFF = "off"
    VALUE_0 = "0"
    VALUE_1 = "1"


class StTargetScreenSz(Enum):
    VALUE_544X376 = "544x376"
    VALUE_640X480 = "640x480"
    VALUE_720X512 = "720x512"
    VALUE_800X600 = "800x600"
    VALUE_1024X768 = "1024x768"
    VALUE_1152X882 = "1152x882"
    VALUE_1152X900 = "1152x900"
    VALUE_1280X1024 = "1280x1024"
    VALUE_1600X1200 = "1600x1200"
    VALUE_1800X1440 = "1800x1440"
    VALUE_1920X1200 = "1920x1200"


@dataclass(slots=True, kw_only=True)
class Tag(Child):
    class Meta:
        name = "tag"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TblGridCol(Child):
    class Meta:
        name = "CT_TblGridCol"

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class UnderlineEnumeration(Enum):
    SINGLE = "single"
    WORDS = "words"
    DOUBLE = "double"
    THICK = "thick"
    DOTTED = "dotted"
    DOTTED_HEAVY = "dottedHeavy"
    DASH = "dash"
    DASHED_HEAVY = "dashedHeavy"
    DASH_LONG = "dashLong"
    DASH_LONG_HEAVY = "dashLongHeavy"
    DOT_DASH = "dotDash"
    DASH_DOT_HEAVY = "dashDotHeavy"
    DOT_DOT_DASH = "dotDotDash"
    DASH_DOT_DOT_HEAVY = "dashDotDotHeavy"
    WAVE = "wave"
    WAVY_HEAVY = "wavyHeavy"
    WAVY_DOUBLE = "wavyDouble"
    NONE = "none"


@dataclass(slots=True, kw_only=True)
class FontsFontAltName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class HighlightVal(Enum):
    BLACK = "black"
    BLUE = "blue"
    CYAN = "cyan"
    GREEN = "green"
    MAGENTA = "magenta"
    RED = "red"
    YELLOW = "yellow"
    WHITE = "white"
    DARK_BLUE = "darkBlue"
    DARK_CYAN = "darkCyan"
    DARK_GREEN = "darkGreen"
    DARK_MAGENTA = "darkMagenta"
    DARK_RED = "darkRed"
    DARK_YELLOW = "darkYellow"
    DARK_GRAY = "darkGray"
    LIGHT_GRAY = "lightGray"
    NONE = "none"


class NumberingAbstractNumMultiLevelTypeVal(Enum):
    SINGLE_LEVEL = "singleLevel"
    MULTILEVEL = "multilevel"
    HYBRID_MULTILEVEL = "hybridMultilevel"


@dataclass(slots=True, kw_only=True)
class NumberingAbstractNumName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingAbstractNumNumStyleLink(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingAbstractNumStyleLink(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingNumIdMacAtCleanup(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingNumAbstractNumId(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingNumLvlOverrideStartOverride(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class ProofErrType(Enum):
    SPELL_START = "spellStart"
    SPELL_END = "spellEnd"
    GRAM_START = "gramStart"
    GRAM_END = "gramEnd"


@dataclass(slots=True, kw_only=True)
class RAnnotationRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RCommentReference(Child):
    class Meta:
        global_type = False

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RContinuationSeparator(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RCr(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RDayLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RDayShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class REndnoteRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RFootnoteRef(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RLastRenderedPageBreak(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RMonthLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RMonthShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RNoBreakHyphen(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RPgNum(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RSeparator(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RSoftHyphen(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RSym(Child):
    class Meta:
        global_type = False

    font: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    char: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RTab(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RYearLong(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RYearShort(Child):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class StyleAliases(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class StyleBasedOn(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class StyleLink(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class StyleName(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class StyleNext(Child):
    class Meta:
        global_type = False

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


class StyleType(Enum):
    PARAGRAPH = "paragraph"
    CHARACTER = "character"
    TABLE = "table"
    NUMBERING = "numbering"


@dataclass(slots=True, kw_only=True)
class StyleUiPriority(Child):
    class Meta:
        global_type = False

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class StylesLatentStylesLsdException(Child):
    class Meta:
        global_type = False

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    locked: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    ui_priority: None | int = field(
        default=None,
        metadata={
            "name": "uiPriority",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    semi_hidden: None | bool = field(
        default=None,
        metadata={
            "name": "semiHidden",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    unhide_when_used: None | bool = field(
        default=None,
        metadata={
            "name": "unhideWhenUsed",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    q_format: None | bool = field(
        default=None,
        metadata={
            "name": "qFormat",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Br(Child):
    class Meta:
        name = "br"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    type_value: None | STBrType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    clear: None | STBrClear = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAltChunkPr(Child):
    class Meta:
        name = "CT_AltChunkPr"

    match_src: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "matchSrc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAutoCaptions(Child):
    class Meta:
        name = "CT_AutoCaptions"

    auto_caption: list[CTAutoCaption] = field(
        default_factory=ChildList,
        metadata={
            "name": "autoCaption",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBackground(CTPictureBase):
    class Meta:
        name = "CT_Background"

    color: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBorder(Child):
    class Meta:
        name = "CT_Border"

    val: None | STBorder = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    space: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shadow: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    frame: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCalendarType(Child):
    class Meta:
        name = "CT_CalendarType"

    val: None | STCalendarType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCaption(Child):
    class Meta:
        name = "CT_Caption"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pos: None | STCaptionPos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    chap_num: None | bool = field(
        default=None,
        metadata={
            "name": "chapNum",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    heading: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_label: None | bool = field(
        default=None,
        metadata={
            "name": "noLabel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    num_fmt: None | NumberFormat = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sep: None | STChapterSep = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCharacterSpacing(Child):
    class Meta:
        name = "CT_CharacterSpacing"

    val: None | STCharacterSpacing = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColor(Child):
    class Meta:
        name = "CT_Color"

    val: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorSchemeMapping(Child):
    class Meta:
        name = "CT_ColorSchemeMapping"

    bg1: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    t1: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bg2: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    t2: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    accent1: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    accent2: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    accent3: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    accent4: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    accent5: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    accent6: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hyperlink: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    followed_hyperlink: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "name": "followedHyperlink",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColumns(Child):
    class Meta:
        name = "CT_Columns"

    col: list[CTColumn] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_occurs": 45,
        },
    )
    equal_width: None | bool = field(
        default=None,
        metadata={
            "name": "equalWidth",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    space: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sep: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCompat(Child):
    class Meta:
        name = "CT_Compat"

    use_single_borderfor_contiguous_cells: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useSingleBorderforContiguousCells",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    wp_justification: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "wpJustification",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_tab_hang_ind: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noTabHangInd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_leading: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noLeading",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    space_for_ul: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "spaceForUL",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_column_balance: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noColumnBalance",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    balance_single_byte_double_byte_width: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "balanceSingleByteDoubleByteWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_extra_line_spacing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noExtraLineSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_leave_backslash_alone: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotLeaveBackslashAlone",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ul_trail_space: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "ulTrailSpace",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_expand_shift_return: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotExpandShiftReturn",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spacing_in_whole_points: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "spacingInWholePoints",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    line_wrap_like_word6: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "lineWrapLikeWord6",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    print_body_text_before_header: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "printBodyTextBeforeHeader",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    print_col_black: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "printColBlack",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    wp_space_width: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "wpSpaceWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    show_breaks_in_frames: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "showBreaksInFrames",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sub_font_by_size: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "subFontBySize",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_bottom_spacing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressBottomSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_top_spacing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressTopSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_spacing_at_top_of_page: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressSpacingAtTopOfPage",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_top_spacing_wp: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressTopSpacingWP",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_sp_bf_after_pg_brk: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressSpBfAfterPgBrk",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    swap_borders_facing_pages: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "swapBordersFacingPages",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    conv_mail_merge_esc: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "convMailMergeEsc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    truncate_font_heights_like_wp6: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "truncateFontHeightsLikeWP6",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mw_small_caps: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "mwSmallCaps",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_printer_metrics: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "usePrinterMetrics",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_suppress_paragraph_borders: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotSuppressParagraphBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    wrap_trail_spaces: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "wrapTrailSpaces",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    footnote_layout_like_ww8: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "footnoteLayoutLikeWW8",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shape_layout_like_ww8: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "shapeLayoutLikeWW8",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    align_tables_row_by_row: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "alignTablesRowByRow",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    forget_last_tab_alignment: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "forgetLastTabAlignment",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    adjust_line_height_in_table: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "adjustLineHeightInTable",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    auto_space_like_word95: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autoSpaceLikeWord95",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_space_raise_lower: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noSpaceRaiseLower",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_use_htmlparagraph_auto_spacing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotUseHTMLParagraphAutoSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    layout_raw_table_width: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "layoutRawTableWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    layout_table_rows_apart: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "layoutTableRowsApart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_word97_line_break_rules: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useWord97LineBreakRules",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_break_wrapped_tables: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotBreakWrappedTables",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_snap_to_grid_in_cell: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotSnapToGridInCell",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    select_fld_with_first_or_last_char: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "selectFldWithFirstOrLastChar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    apply_breaking_rules: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "applyBreakingRules",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_wrap_text_with_punct: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotWrapTextWithPunct",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_use_east_asian_break_rules: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotUseEastAsianBreakRules",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_word2002_table_style_rules: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useWord2002TableStyleRules",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    grow_autofit: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "growAutofit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_felayout: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useFELayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_normal_style_for_list: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useNormalStyleForList",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_use_indent_as_numbering_tab_stop: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotUseIndentAsNumberingTabStop",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_alt_kinsoku_line_break_rules: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useAltKinsokuLineBreakRules",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    allow_space_of_same_style_in_table: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "allowSpaceOfSameStyleInTable",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_suppress_indentation: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotSuppressIndentation",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_autofit_constrained_tables: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotAutofitConstrainedTables",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    autofit_to_first_fixed_width_cell: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autofitToFirstFixedWidthCell",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    underline_tab_in_num_list: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "underlineTabInNumList",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    display_hangul_fixed_width: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "displayHangulFixedWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    split_pg_break_and_para_mark: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "splitPgBreakAndParaMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_vert_align_cell_with_sp: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotVertAlignCellWithSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_break_constrained_forced_table: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotBreakConstrainedForcedTable",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_vert_align_in_txbx: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotVertAlignInTxbx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_ansi_kerning_pairs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useAnsiKerningPairs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cached_col_balance: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "cachedColBalance",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    compat_setting: list[CTCompatSetting] = field(
        default_factory=ChildList,
        metadata={
            "name": "compatSetting",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomXmlPr(Child):
    class Meta:
        name = "CT_CustomXmlPr"

    placeholder: None | CtCustomXmlPrPlaceholder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    attr: list[CTAttr] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocGrid(Child):
    class Meta:
        name = "CT_DocGrid"

    type_value: None | STDocGrid = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    line_pitch: None | int = field(
        default=None,
        metadata={
            "name": "linePitch",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    char_space: None | int = field(
        default=None,
        metadata={
            "name": "charSpace",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartBehavior(Child):
    class Meta:
        name = "CT_DocPartBehavior"

    val: None | STDocPartBehavior = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartGallery(Child):
    class Meta:
        name = "CT_DocPartGallery"

    val: None | STDocPartGallery = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartType(Child):
    class Meta:
        name = "CT_DocPartType"

    val: None | STDocPartType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocProtect(Child):
    class Meta:
        name = "CT_DocProtect"

    edit: None | STDocProtect = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    formatting: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    enforcement: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    crypt_provider_type: None | STCryptProv = field(
        default=None,
        metadata={
            "name": "cryptProviderType",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_algorithm_class: None | STAlgClass = field(
        default=None,
        metadata={
            "name": "cryptAlgorithmClass",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_algorithm_type: None | STAlgType = field(
        default=None,
        metadata={
            "name": "cryptAlgorithmType",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_algorithm_sid: None | int = field(
        default=None,
        metadata={
            "name": "cryptAlgorithmSid",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_spin_count: None | int = field(
        default=None,
        metadata={
            "name": "cryptSpinCount",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_provider: None | str = field(
        default=None,
        metadata={
            "name": "cryptProvider",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    alg_id_ext: None | str = field(
        default=None,
        metadata={
            "name": "algIdExt",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    alg_id_ext_source: None | str = field(
        default=None,
        metadata={
            "name": "algIdExtSource",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_provider_type_ext: None | str = field(
        default=None,
        metadata={
            "name": "cryptProviderTypeExt",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_provider_type_ext_source: None | str = field(
        default=None,
        metadata={
            "name": "cryptProviderTypeExtSource",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hash: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "format": "base64",
        },
    )
    salt: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocRsids(Child):
    class Meta:
        name = "CT_DocRsids"

    rsid_root: None | CTLongHexNumber = field(
        default=None,
        metadata={
            "name": "rsidRoot",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid: list[CTLongHexNumber] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocType(Child):
    class Meta:
        name = "CT_DocType"

    val: None | STDocType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocVars(Child):
    class Meta:
        name = "CT_DocVars"

    doc_var: list[CTDocVar] = field(
        default_factory=ChildList,
        metadata={
            "name": "docVar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEastAsianLayout(Child):
    class Meta:
        name = "CT_EastAsianLayout"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    combine: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    combine_brackets: None | STCombineBrackets = field(
        default=None,
        metadata={
            "name": "combineBrackets",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vert: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    vert_compress: None | bool = field(
        default=None,
        metadata={
            "name": "vertCompress",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEdnPos(Child):
    class Meta:
        name = "CT_EdnPos"

    val: None | STEdnPos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEm(Child):
    class Meta:
        name = "CT_Em"

    val: None | STEm = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFCheckBox(Child):
    class Meta:
        name = "CT_FFCheckBox"

    size_or_size_auto: None | HpsMeasure | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "size",
                    "type": ForwardRef("HpsMeasure"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sizeAuto",
                    "type": ForwardRef("BooleanDefaultTrue"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    default: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    checked: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFDDList(Child):
    class Meta:
        name = "CT_FFDDList"

    result: None | CtFfddlistResult = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    default: None | CtFfddlistDefault = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    list_entry: list[CtFfddlistListEntry] = field(
        default_factory=ChildList,
        metadata={
            "name": "listEntry",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFHelpText(Child):
    class Meta:
        name = "CT_FFHelpText"

    type_value: None | STInfoTextType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_length": 256,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFStatusText(Child):
    class Meta:
        name = "CT_FFStatusText"

    type_value: None | STInfoTextType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_length": 140,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFTextType(Child):
    class Meta:
        name = "CT_FFTextType"

    val: None | STFFTextType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFrameLayout(Child):
    class Meta:
        name = "CT_FrameLayout"

    val: None | STFrameLayout = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFramePr(Child):
    class Meta:
        name = "CT_FramePr"

    drop_cap: None | STDropCap = field(
        default=None,
        metadata={
            "name": "dropCap",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lines: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_space: None | int = field(
        default=None,
        metadata={
            "name": "vSpace",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_space: None | int = field(
        default=None,
        metadata={
            "name": "hSpace",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    wrap: None | STWrap = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_anchor: None | STHAnchor = field(
        default=None,
        metadata={
            "name": "hAnchor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_anchor: None | STVAnchor = field(
        default=None,
        metadata={
            "name": "vAnchor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    x: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    x_align: None | STXAlign = field(
        default=None,
        metadata={
            "name": "xAlign",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    y: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    y_align: None | STYAlign = field(
        default=None,
        metadata={
            "name": "yAlign",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_rule: None | STHeightRule = field(
        default=None,
        metadata={
            "name": "hRule",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    anchor_lock: None | bool = field(
        default=None,
        metadata={
            "name": "anchorLock",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFrameScrollbar(Child):
    class Meta:
        name = "CT_FrameScrollbar"

    val: None | STFrameScrollbar = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFtnPos(Child):
    class Meta:
        name = "CT_FtnPos"

    val: None | STFtnPos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHeight(Child):
    class Meta:
        name = "CT_Height"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_rule: None | STHeightRule = field(
        default=None,
        metadata={
            "name": "hRule",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineNumber(Child):
    class Meta:
        name = "CT_LineNumber"

    count_by: None | int = field(
        default=None,
        metadata={
            "name": "countBy",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    start: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    distance: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    restart: None | STLineNumberRestart = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLock(Child):
    class Meta:
        name = "CT_Lock"

    val: None | STLock = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMailMergeDataType(Child):
    class Meta:
        name = "CT_MailMergeDataType"

    val: None | STMailMergeDataType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMailMergeDest(Child):
    class Meta:
        name = "CT_MailMergeDest"

    val: None | STMailMergeDest = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMailMergeDocType(Child):
    class Meta:
        name = "CT_MailMergeDocType"

    val: None | STMailMergeDocType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMailMergeOdsoFMDFieldType(Child):
    class Meta:
        name = "CT_MailMergeOdsoFMDFieldType"

    val: None | STMailMergeOdsoFMDFieldType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMailMergeSourceType(Child):
    class Meta:
        name = "CT_MailMergeSourceType"

    val: None | STMailMergeSourceType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarkupRange(CTMarkup):
    class Meta:
        name = "CT_MarkupRange"

    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMoveFromRangeEnd(CTMarkup):
    class Meta:
        name = "CT_MoveFromRangeEnd"

    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMoveToRangeEnd(CTMarkup):
    class Meta:
        name = "CT_MoveToRangeEnd"

    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumRestart(Child):
    class Meta:
        name = "CT_NumRestart"

    val: None | STRestartNumber = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTObject(CTPictureBase):
    class Meta:
        name = "CT_Object"

    control: None | CTControl = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dxa_orig: None | int = field(
        default=None,
        metadata={
            "name": "dxaOrig",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dya_orig: None | int = field(
        default=None,
        metadata={
            "name": "dyaOrig",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPageNumber(Child):
    class Meta:
        name = "CT_PageNumber"

    fmt: None | NumberFormat = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    start: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    chap_style: None | int = field(
        default=None,
        metadata={
            "name": "chapStyle",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    chap_sep: None | STChapterSep = field(
        default=None,
        metadata={
            "name": "chapSep",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPerm(Child):
    class Meta:
        name = "CT_Perm"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPlaceholder(Child):
    class Meta:
        name = "CT_Placeholder"

    doc_part: None | CtPlaceholderDocPart = field(
        default=None,
        metadata={
            "name": "docPart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTProof(Child):
    class Meta:
        name = "CT_Proof"

    spelling: None | STProof = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    grammar: None | STProof = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRecipientData(Child):
    class Meta:
        name = "CT_RecipientData"

    active: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    column: None | CtRecipientDataColumn = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    unique_tag: None | bytes = field(
        default=None,
        metadata={
            "name": "uniqueTag",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRubyAlign(Child):
    class Meta:
        name = "CT_RubyAlign"

    val: None | STRubyAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtComboBox(Child):
    class Meta:
        name = "CT_SdtComboBox"

    list_item: list[CTSdtListItem] = field(
        default_factory=ChildList,
        metadata={
            "name": "listItem",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    last_value: None | str = field(
        default=None,
        metadata={
            "name": "lastValue",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtDateMappingType(Child):
    class Meta:
        name = "CT_SdtDateMappingType"

    val: None | STSdtDateMappingType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtDocPart(Child):
    class Meta:
        name = "CT_SdtDocPart"

    doc_part_gallery: None | CtSdtDocPartDocPartGallery = field(
        default=None,
        metadata={
            "name": "docPartGallery",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_part_category: None | CtSdtDocPartDocPartCategory = field(
        default=None,
        metadata={
            "name": "docPartCategory",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_part_unique: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "docPartUnique",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtDropDownList(Child):
    class Meta:
        name = "CT_SdtDropDownList"

    list_item: list[CTSdtListItem] = field(
        default_factory=ChildList,
        metadata={
            "name": "listItem",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    last_value: None | str = field(
        default=None,
        metadata={
            "name": "lastValue",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShd(Child):
    class Meta:
        name = "CT_Shd"

    val: None | STShd = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fill: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_fill: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeFill",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_fill_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeFillTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_fill_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeFillShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSmartTagPr(Child):
    class Meta:
        name = "CT_SmartTagPr"

    attr: list[CTAttr] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTabStop(Child):
    class Meta:
        name = "CT_TabStop"

    val: None | STTabJc = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    leader: None | STTabTlc = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pos: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTargetScreenSz(Child):
    class Meta:
        name = "CT_TargetScreenSz"

    val: None | StTargetScreenSz = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblLayoutType(Child):
    class Meta:
        name = "CT_TblLayoutType"

    type_value: None | STTblLayoutType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblLook(Child):
    class Meta:
        name = "CT_TblLook"

    first_row: None | SharedTypesStOnOff = field(
        default=None,
        metadata={
            "name": "firstRow",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    last_row: None | SharedTypesStOnOff = field(
        default=None,
        metadata={
            "name": "lastRow",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    first_column: None | SharedTypesStOnOff = field(
        default=None,
        metadata={
            "name": "firstColumn",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    last_column: None | SharedTypesStOnOff = field(
        default=None,
        metadata={
            "name": "lastColumn",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_hband: None | SharedTypesStOnOff = field(
        default=None,
        metadata={
            "name": "noHBand",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_vband: None | SharedTypesStOnOff = field(
        default=None,
        metadata={
            "name": "noVBand",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblOverlap(Child):
    class Meta:
        name = "CT_TblOverlap"

    val: None | STTblOverlap = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblPPr(Child):
    class Meta:
        name = "CT_TblPPr"

    left_from_text: None | int = field(
        default=None,
        metadata={
            "name": "leftFromText",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right_from_text: None | int = field(
        default=None,
        metadata={
            "name": "rightFromText",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    top_from_text: None | int = field(
        default=None,
        metadata={
            "name": "topFromText",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom_from_text: None | int = field(
        default=None,
        metadata={
            "name": "bottomFromText",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vert_anchor: None | STVAnchor = field(
        default=None,
        metadata={
            "name": "vertAnchor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    horz_anchor: None | STHAnchor = field(
        default=None,
        metadata={
            "name": "horzAnchor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tblp_xspec: None | STXAlign = field(
        default=None,
        metadata={
            "name": "tblpXSpec",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tblp_x: None | int = field(
        default=None,
        metadata={
            "name": "tblpX",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tblp_yspec: None | STYAlign = field(
        default=None,
        metadata={
            "name": "tblpYSpec",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tblp_y: None | int = field(
        default=None,
        metadata={
            "name": "tblpY",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextEffect(Child):
    class Meta:
        name = "CT_TextEffect"

    val: None | STTextEffect = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextboxTightWrap(Child):
    class Meta:
        name = "CT_TextboxTightWrap"

    val: None | STTextboxTightWrap = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrackChange(CTMarkup):
    class Meta:
        name = "CT_TrackChange"

    author: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    date: None | XmlDateTime = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTVerticalAlignRun(Child):
    class Meta:
        name = "CT_VerticalAlignRun"

    val: None | STVerticalAlignRun = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTVerticalJc(Child):
    class Meta:
        name = "CT_VerticalJc"

    val: None | STVerticalJc = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTView(Child):
    class Meta:
        name = "CT_View"

    val: None | STView = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWriteProtection(Child):
    class Meta:
        name = "CT_WriteProtection"

    recommended: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    crypt_provider_type: None | STCryptProv = field(
        default=None,
        metadata={
            "name": "cryptProviderType",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_algorithm_class: None | STAlgClass = field(
        default=None,
        metadata={
            "name": "cryptAlgorithmClass",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_algorithm_type: None | STAlgType = field(
        default=None,
        metadata={
            "name": "cryptAlgorithmType",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_algorithm_sid: None | int = field(
        default=None,
        metadata={
            "name": "cryptAlgorithmSid",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_spin_count: None | int = field(
        default=None,
        metadata={
            "name": "cryptSpinCount",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_provider: None | str = field(
        default=None,
        metadata={
            "name": "cryptProvider",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    alg_id_ext: None | str = field(
        default=None,
        metadata={
            "name": "algIdExt",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    alg_id_ext_source: None | str = field(
        default=None,
        metadata={
            "name": "algIdExtSource",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_provider_type_ext: None | str = field(
        default=None,
        metadata={
            "name": "cryptProviderTypeExt",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    crypt_provider_type_ext_source: None | str = field(
        default=None,
        metadata={
            "name": "cryptProviderTypeExtSource",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hash: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "format": "base64",
        },
    )
    salt: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTZoom(Child):
    class Meta:
        name = "CT_Zoom"

    val: None | STZoom = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    percent: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFramesetSplitbarColor(Child):
    class Meta:
        global_type = False

    val: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtLvlSuff(Child):
    class Meta:
        global_type = False

    val: None | CtLvlSuffVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtMathRunTrackChangePtab(Child):
    class Meta:
        global_type = False

    alignment: None | STPTabAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    relative_to: None | STPTabRelativeTo = field(
        default=None,
        metadata={
            "name": "relativeTo",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    leader: None | STPTabLeader = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseSpacing(Child):
    class Meta:
        global_type = False

    before: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    before_lines: None | int = field(
        default=None,
        metadata={
            "name": "beforeLines",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    before_autospacing: None | bool = field(
        default=None,
        metadata={
            "name": "beforeAutospacing",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "false",
        },
    )
    after: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    after_lines: None | int = field(
        default=None,
        metadata={
            "name": "afterLines",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    after_autospacing: None | bool = field(
        default=None,
        metadata={
            "name": "afterAutospacing",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "false",
        },
    )
    line: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    line_rule: None | STLineSpacingRule = field(
        default=None,
        metadata={
            "name": "lineRule",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseTextAlignment(Child):
    class Meta:
        global_type = False

    val: None | CtPprBaseTextAlignmentVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtRPtab(Child):
    class Meta:
        global_type = False

    alignment: None | STPTabAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    relative_to: None | STPTabRelativeTo = field(
        default=None,
        metadata={
            "name": "relativeTo",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    leader: None | STPTabLeader = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSectPrBasePgSz(Child):
    class Meta:
        global_type = False

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    orient: None | STPageOrientation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    code: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSectPrBaseType(Child):
    class Meta:
        global_type = False

    val: None | CtSectPrBaseTypeVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSectPrPgSz(Child):
    class Meta:
        global_type = False

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    orient: None | STPageOrientation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    code: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSectPrType(Child):
    class Meta:
        global_type = False

    val: None | CtSectPrTypeVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTcPrInnerHMerge(Child):
    class Meta:
        global_type = False

    val: None | CtTcPrInnerHMergeVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTcPrInnerVMerge(Child):
    class Meta:
        global_type = False

    val: None | CtTcPrInnerVMergeVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Color(Child):
    class Meta:
        name = "color"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    val: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CommentRangeEnd(Child):
    class Meta:
        name = "commentRangeEnd"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CommentRangeStart(Child):
    class Meta:
        name = "commentRangeStart"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class DelText(Child):
    class Meta:
        name = "delText"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    value: str = field(default="")
    space: None | SpaceValue = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )


@dataclass(slots=True, kw_only=True)
class Drawing(Child):
    class Meta:
        name = "CT_Drawing"

    anchor_or_inline: list[Anchor | Inline] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "anchor",
                    "type": ForwardRef("Anchor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                },
                {
                    "name": "inline",
                    "type": ForwardRef("Inline"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class FontFamily(Child):
    class Meta:
        name = "CT_FontFamily"

    val: None | CtFontFamilyVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class FontPitch(Child):
    class Meta:
        name = "CT_Pitch"

    val: None | STPitch = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class FontRel(CTRel):
    class Meta:
        name = "CT_FontRel"

    font_key: None | str = field(
        default=None,
        metadata={
            "name": "fontKey",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    subsetted: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class FooterReference(CTRel):
    class Meta:
        name = "footerReference"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    type_value: None | HdrFtrRef = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class HeaderReference(CTRel):
    class Meta:
        name = "headerReference"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    type_value: None | HdrFtrRef = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Highlight(Child):
    class Meta:
        name = "highlight"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    val: None | HighlightVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Jc(Child):
    class Meta:
        name = "CT_Jc"

    val: None | JcEnumeration = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumFmt(Child):
    class Meta:
        name = "CT_NumFmt"

    val: None | NumberFormat = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Pict(CTPictureBase):
    class Meta:
        name = "CT_Picture"

    movie: None | CTRel = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    control: None | CTControl = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class ProofErr(Child):
    class Meta:
        name = "proofErr"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    type_value: None | ProofErrType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RFonts(Child):
    class Meta:
        name = "rFonts"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    hint: None | STHint = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ascii: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_ansi: None | str = field(
        default=None,
        metadata={
            "name": "hAnsi",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    east_asia: None | str = field(
        default=None,
        metadata={
            "name": "eastAsia",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cs: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ascii_theme: None | STTheme = field(
        default=None,
        metadata={
            "name": "asciiTheme",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_ansi_theme: None | STTheme = field(
        default=None,
        metadata={
            "name": "hAnsiTheme",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    east_asia_theme: None | STTheme = field(
        default=None,
        metadata={
            "name": "eastAsiaTheme",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cstheme: None | STTheme = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TblGridBase(Child):
    class Meta:
        name = "CT_TblGridBase"

    grid_col: list[TblGridCol] = field(
        default_factory=ChildList,
        metadata={
            "name": "gridCol",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TblWidth(Child):
    class Meta:
        name = "CT_TblWidth"

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    type_value: None | CtTblWidthType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Text(Child):
    class Meta:
        name = "CT_Text"

    value: str = field(default="")
    space: None | SpaceValue = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )


@dataclass(slots=True, kw_only=True)
class TextDirection(Child):
    class Meta:
        name = "CT_TextDirection"

    val: None | CtTextDirectionVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class U(Child):
    class Meta:
        name = "u"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    val: None | UnderlineEnumeration = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: None | STHexColorAuto | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "length": 3,
            "format": "base16",
        },
    )
    theme_color: None | STThemeColor = field(
        default=None,
        metadata={
            "name": "themeColor",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_tint: None | str = field(
        default=None,
        metadata={
            "name": "themeTint",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_shade: None | str = field(
        default=None,
        metadata={
            "name": "themeShade",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CalcOnExit(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CantSplit(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlDelRangeEnd1(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlDelRangeEnd2(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlInsRangeEnd1(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlInsRangeEnd2(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveFromRangeEnd1(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveFromRangeEnd2(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveToRangeEnd1(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveToRangeEnd2(CTMarkup):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Enabled(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class EndnoteReference1(CTFtnEdnRef):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class EndnoteReference2(CTFtnEdnRef):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class EntryMacro(CTMacroName):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ExitMacro(CTMacroName):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class FootnoteReference1(CTFtnEdnRef):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class FootnoteReference2(CTFtnEdnRef):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Hidden(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class NumberingAbstractNumMultiLevelType(Child):
    class Meta:
        global_type = False

    val: None | NumberingAbstractNumMultiLevelTypeVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RPtab(Child):
    class Meta:
        global_type = False

    alignment: None | STPTabAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    relative_to: None | STPTabRelativeTo = field(
        default=None,
        metadata={
            "name": "relativeTo",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    leader: None | STPTabLeader = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class ShowingPlcHdr(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class StylesLatentStyles(Child):
    class Meta:
        global_type = False

    lsd_exception: list[StylesLatentStylesLsdException] = field(
        default_factory=ChildList,
        metadata={
            "name": "lsdException",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    def_locked_state: None | bool = field(
        default=None,
        metadata={
            "name": "defLockedState",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    def_uipriority: None | int = field(
        default=None,
        metadata={
            "name": "defUIPriority",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    def_semi_hidden: None | bool = field(
        default=None,
        metadata={
            "name": "defSemiHidden",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    def_unhide_when_used: None | bool = field(
        default=None,
        metadata={
            "name": "defUnhideWhenUsed",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    def_qformat: None | bool = field(
        default=None,
        metadata={
            "name": "defQFormat",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "false",
        },
    )
    count: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TblHeader(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Temporary(BooleanDefaultTrue):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTAltChunk(Child):
    class Meta:
        name = "CT_AltChunk"

    alt_chunk_pr: None | CTAltChunkPr = field(
        default=None,
        metadata={
            "name": "altChunkPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBookmarkRange(CTMarkupRange):
    class Meta:
        name = "CT_BookmarkRange"

    col_first: None | int = field(
        default=None,
        metadata={
            "name": "colFirst",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    col_last: None | int = field(
        default=None,
        metadata={
            "name": "colLast",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCaptions(Child):
    class Meta:
        name = "CT_Captions"

    caption: list[CTCaption] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "min_occurs": 1,
        },
    )
    auto_captions: None | CTAutoCaptions = field(
        default=None,
        metadata={
            "name": "autoCaptions",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCellMergeTrackChange(CTTrackChange):
    class Meta:
        name = "CT_CellMergeTrackChange"

    v_merge: None | STAnnotationVMerge = field(
        default=None,
        metadata={
            "name": "vMerge",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_merge_orig: None | STAnnotationVMerge = field(
        default=None,
        metadata={
            "name": "vMergeOrig",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDivBdr(Child):
    class Meta:
        name = "CT_DivBdr"

    top: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartBehaviors(Child):
    class Meta:
        name = "CT_DocPartBehaviors"

    behavior: list[CTDocPartBehavior] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartCategory(Child):
    class Meta:
        name = "CT_DocPartCategory"

    name: None | CtDocPartCategoryName = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    gallery: None | CTDocPartGallery = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartTypes(Child):
    class Meta:
        name = "CT_DocPartTypes"

    type_value: list[CTDocPartType] = field(
        default_factory=ChildList,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    all: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEdnProps(Child):
    class Meta:
        name = "CT_EdnProps"

    pos: None | CTEdnPos = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_fmt: None | NumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_start: None | CtEdnPropsNumStart = field(
        default=None,
        metadata={
            "name": "numStart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_restart: None | CTNumRestart = field(
        default=None,
        metadata={
            "name": "numRestart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFTextInput(Child):
    class Meta:
        name = "CT_FFTextInput"

    type_value: None | CTFFTextType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    default: None | CtFftextInputDefault = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    max_length: None | CtFftextInputMaxLength = field(
        default=None,
        metadata={
            "name": "maxLength",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    format: None | CtFftextInputFormat = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFrame(Child):
    class Meta:
        name = "CT_Frame"

    sz: None | CtFrameSz = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    name: None | CtFrameName = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    source_file_name: None | CTRel = field(
        default=None,
        metadata={
            "name": "sourceFileName",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mar_w: None | CTPixelsMeasure = field(
        default=None,
        metadata={
            "name": "marW",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mar_h: None | CTPixelsMeasure = field(
        default=None,
        metadata={
            "name": "marH",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    scrollbar: None | CTFrameScrollbar = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_resize_allowed: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noResizeAllowed",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    linked_to_file: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "linkedToFile",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFramesetSplitbar(Child):
    class Meta:
        name = "CT_FramesetSplitbar"

    w: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: None | CtFramesetSplitbarColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_border: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noBorder",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    flat_borders: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "flatBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFtnProps(Child):
    class Meta:
        name = "CT_FtnProps"

    pos: None | CTFtnPos = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_fmt: None | NumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_start: None | CtFtnPropsNumStart = field(
        default=None,
        metadata={
            "name": "numStart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_restart: None | CTNumRestart = field(
        default=None,
        metadata={
            "name": "numRestart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOdsoFieldMapData(Child):
    class Meta:
        name = "CT_OdsoFieldMapData"

    type_value: None | CTMailMergeOdsoFMDFieldType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    name: None | CtOdsoFieldMapDataName = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mapped_name: None | CtOdsoFieldMapDataMappedName = field(
        default=None,
        metadata={
            "name": "mappedName",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    column: None | CtOdsoFieldMapDataColumn = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lid: None | CTLang = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dynamic_address: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "dynamicAddress",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTParaRPrOriginal(Child):
    class Meta:
        name = "CT_ParaRPrOriginal"

    ins: None | CTTrackChange = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    del_value: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "del",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    move_from: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "moveFrom",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    move_to: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "moveTo",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_style: list[RStyle] = field(
        default_factory=ChildList,
        metadata={
            "name": "rStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_fonts: list[RFonts] = field(
        default_factory=ChildList,
        metadata={
            "name": "rFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b_cs: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "bCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i_cs: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "iCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    caps: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    small_caps: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "smallCaps",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    strike: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dstrike: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    outline: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shadow: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    emboss: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    imprint: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_proof: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "noProof",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    snap_to_grid: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "snapToGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vanish: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    web_hidden: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "webHidden",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: list[Color] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spacing: list[CTSignedTwipsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    w: list[CTTextScale] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    kern: list[HpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    position: list[CTSignedHpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz: list[HpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz_cs: list[HpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "name": "szCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    highlight: list[Highlight] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    u: list[U] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    effect: list[CTTextEffect] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bdr: list[CTBorder] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: list[CTShd] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fit_text: list[CTFitText] = field(
        default_factory=ChildList,
        metadata={
            "name": "fitText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vert_align: list[CTVerticalAlignRun] = field(
        default_factory=ChildList,
        metadata={
            "name": "vertAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rtl: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cs: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    em: list[CTEm] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lang: list[CTLanguage] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    east_asian_layout: list[CTEastAsianLayout] = field(
        default_factory=ChildList,
        metadata={
            "name": "eastAsianLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spec_vanish: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "specVanish",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    o_math: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "oMath",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    glow: list[CTGlow] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    schemas_microsoft_com_office_word_2010_wordml_shadow: list[Shadow] = field(
        default_factory=ChildList,
        metadata={
            "name": "shadow",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    reflection: list[CTReflection] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_outline: list[CTTextOutlineEffect] = field(
        default_factory=ChildList,
        metadata={
            "name": "textOutline",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_fill: list[CTFillTextEffect] = field(
        default_factory=ChildList,
        metadata={
            "name": "textFill",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    scene3d: list[CTScene3D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    props3d: list[CTProps3D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    ligatures: list[CTLigatures] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_form: list[CTNumForm] = field(
        default_factory=ChildList,
        metadata={
            "name": "numForm",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_spacing: list[CTNumSpacing] = field(
        default_factory=ChildList,
        metadata={
            "name": "numSpacing",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    stylistic_sets: list[CTStylisticSets] = field(
        default_factory=ChildList,
        metadata={
            "name": "stylisticSets",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    cntxt_alts: list[CTOnOff] = field(
        default_factory=ChildList,
        metadata={
            "name": "cntxtAlts",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRecipients(Child):
    class Meta:
        name = "CT_Recipients"

    recipient_data: list[CTRecipientData] = field(
        default_factory=ChildList,
        metadata={
            "name": "recipientData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "min_occurs": 1,
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRubyPr(Child):
    class Meta:
        name = "CT_RubyPr"

    ruby_align: None | CTRubyAlign = field(
        default=None,
        metadata={
            "name": "rubyAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hps: None | HpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hps_raise: None | HpsMeasure = field(
        default=None,
        metadata={
            "name": "hpsRaise",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hps_base_text: None | HpsMeasure = field(
        default=None,
        metadata={
            "name": "hpsBaseText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lid: None | CTLang = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dirty: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtDate(Child):
    class Meta:
        name = "CT_SdtDate"

    date_format: None | CtSdtDateDateFormat = field(
        default=None,
        metadata={
            "name": "dateFormat",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lid: None | CTLang = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    store_mapped_data_as: None | CTSdtDateMappingType = field(
        default=None,
        metadata={
            "name": "storeMappedDataAs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    calendar: None | CTCalendarType = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    full_date: None | XmlDateTime = field(
        default=None,
        metadata={
            "name": "fullDate",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblCellMar(Child):
    class Meta:
        name = "CT_TblCellMar"

    top: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    start: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    end: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblGridChange(CTMarkup):
    class Meta:
        name = "CT_TblGridChange"

    tbl_grid: None | TblGridBase = field(
        default=None,
        metadata={
            "name": "tblGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrackChangeNumbering(CTTrackChange):
    class Meta:
        name = "CT_TrackChangeNumbering"

    original: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrackChangeRange(CTTrackChange):
    class Meta:
        name = "CT_TrackChangeRange"

    displaced_by_custom_xml: None | STDisplacedByCustomXml = field(
        default=None,
        metadata={
            "name": "displacedByCustomXml",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBasePBdr(Child):
    class Meta:
        global_type = False

    top: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    between: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bar: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtRprChangeRPr(Child):
    class Meta:
        global_type = False

    r_style: list[RStyle] = field(
        default_factory=ChildList,
        metadata={
            "name": "rStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_fonts: list[RFonts] = field(
        default_factory=ChildList,
        metadata={
            "name": "rFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b_cs: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "bCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i_cs: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "iCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    caps: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    small_caps: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "smallCaps",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    strike: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dstrike: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    outline: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shadow: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    emboss: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    imprint: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_proof: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "noProof",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    snap_to_grid: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "snapToGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vanish: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    web_hidden: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "webHidden",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: list[Color] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spacing: list[CTSignedTwipsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    w: list[CTTextScale] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    kern: list[HpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    position: list[CTSignedHpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz: list[HpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz_cs: list[HpsMeasure] = field(
        default_factory=ChildList,
        metadata={
            "name": "szCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    highlight: list[Highlight] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    u: list[U] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    effect: list[CTTextEffect] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bdr: list[CTBorder] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: list[CTShd] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fit_text: list[CTFitText] = field(
        default_factory=ChildList,
        metadata={
            "name": "fitText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vert_align: list[CTVerticalAlignRun] = field(
        default_factory=ChildList,
        metadata={
            "name": "vertAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rtl: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cs: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    em: list[CTEm] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lang: list[CTLanguage] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    east_asian_layout: list[CTEastAsianLayout] = field(
        default_factory=ChildList,
        metadata={
            "name": "eastAsianLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spec_vanish: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "specVanish",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    o_math: list[BooleanDefaultTrue] = field(
        default_factory=ChildList,
        metadata={
            "name": "oMath",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    glow: list[CTGlow] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    schemas_microsoft_com_office_word_2010_wordml_shadow: list[Shadow] = field(
        default_factory=ChildList,
        metadata={
            "name": "shadow",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    reflection: list[CTReflection] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_outline: list[CTTextOutlineEffect] = field(
        default_factory=ChildList,
        metadata={
            "name": "textOutline",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_fill: list[CTFillTextEffect] = field(
        default_factory=ChildList,
        metadata={
            "name": "textFill",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    scene3d: list[CTScene3D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    props3d: list[CTProps3D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    ligatures: list[CTLigatures] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_form: list[CTNumForm] = field(
        default_factory=ChildList,
        metadata={
            "name": "numForm",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_spacing: list[CTNumSpacing] = field(
        default_factory=ChildList,
        metadata={
            "name": "numSpacing",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    stylistic_sets: list[CTStylisticSets] = field(
        default_factory=ChildList,
        metadata={
            "name": "stylisticSets",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    cntxt_alts: list[CTOnOff] = field(
        default_factory=ChildList,
        metadata={
            "name": "cntxtAlts",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSectPrBasePgBorders(Child):
    class Meta:
        global_type = False

    top: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    z_order: None | STPageBorderZOrder = field(
        default=None,
        metadata={
            "name": "zOrder",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    display: None | STPageBorderDisplay = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    offset_from: None | STPageBorderOffset = field(
        default=None,
        metadata={
            "name": "offsetFrom",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSectPrPgBorders(Child):
    class Meta:
        global_type = False

    top: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    z_order: None | STPageBorderZOrder = field(
        default=None,
        metadata={
            "name": "zOrder",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    display: None | STPageBorderDisplay = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    offset_from: None | STPageBorderOffset = field(
        default=None,
        metadata={
            "name": "offsetFrom",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtTcPrInnerTcBorders(Child):
    class Meta:
        global_type = False

    top: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    start: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    end: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    inside_h: None | CTBorder = field(
        default=None,
        metadata={
            "name": "insideH",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    inside_v: None | CTBorder = field(
        default=None,
        metadata={
            "name": "insideV",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tl2br: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tr2bl: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RDelInstrText(Text):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RInstrText(Text):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RT(Text):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RangePermissionStart(CTPerm):
    class Meta:
        name = "CT_PermStart"

    ed_grp: None | CtPermStartEdGrp = field(
        default=None,
        metadata={
            "name": "edGrp",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ed: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    col_first: None | int = field(
        default=None,
        metadata={
            "name": "colFirst",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    col_last: None | int = field(
        default=None,
        metadata={
            "name": "colLast",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Tabs(Child):
    class Meta:
        name = "CT_Tabs"

    tab: list[CTTabStop] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class TblBorders(Child):
    class Meta:
        name = "CT_TblBorders"

    top: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    inside_h: None | CTBorder = field(
        default=None,
        metadata={
            "name": "insideH",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    inside_v: None | CTBorder = field(
        default=None,
        metadata={
            "name": "insideV",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TcMar(Child):
    class Meta:
        name = "CT_TcMar"

    top: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    left: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bottom: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    right: None | TblWidth = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CellDel(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CellIns(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlDelRangeStart1(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlDelRangeStart2(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlInsRangeStart1(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlInsRangeStart2(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveFromRangeStart1(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveFromRangeStart2(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveToRangeStart1(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlMoveToRangeStart2(CTTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class DelInstrText(Text):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class DocPartList(CTSdtDocPart):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class DocPartObj(CTSdtDocPart):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class FontsFont(Child):
    class Meta:
        global_type = False

    alt_name: None | FontsFontAltName = field(
        default=None,
        metadata={
            "name": "altName",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    panose1: None | FontPanose = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    charset: None | CTUcharHexNumber = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    family: None | FontFamily = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    not_true_type: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "notTrueType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pitch: None | FontPitch = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sig: None | FontSig = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    embed_regular: None | FontRel = field(
        default=None,
        metadata={
            "name": "embedRegular",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    embed_bold: None | FontRel = field(
        default=None,
        metadata={
            "name": "embedBold",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    embed_italic: None | FontRel = field(
        default=None,
        metadata={
            "name": "embedItalic",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    embed_bold_italic: None | FontRel = field(
        default=None,
        metadata={
            "name": "embedBoldItalic",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class InstrText(Text):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class NumberingNumPicBullet(Child):
    class Meta:
        global_type = False

    pict_or_drawing: None | Pict | Drawing = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "pict",
                    "type": ForwardRef("Pict"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "drawing",
                    "type": ForwardRef("Drawing"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    num_pic_bullet_id: None | int = field(
        default=None,
        metadata={
            "name": "numPicBulletId",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class T(Text):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class TblCellSpacing(TblWidth):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class WAfter(TblWidth):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class WBefore(TblWidth):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTBookmark(CTBookmarkRange):
    class Meta:
        name = "CT_Bookmark"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDiv(Child):
    class Meta:
        name = "CT_Div"

    block_quote: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "blockQuote",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    body_div: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bodyDiv",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mar_left: None | CTSignedTwipsMeasure = field(
        default=None,
        metadata={
            "name": "marLeft",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mar_right: None | CTSignedTwipsMeasure = field(
        default=None,
        metadata={
            "name": "marRight",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mar_top: None | CTSignedTwipsMeasure = field(
        default=None,
        metadata={
            "name": "marTop",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mar_bottom: None | CTSignedTwipsMeasure = field(
        default=None,
        metadata={
            "name": "marBottom",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    div_bdr: None | CTDivBdr = field(
        default=None,
        metadata={
            "name": "divBdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    divs_child: list[CTDivs] = field(
        default_factory=ChildList,
        metadata={
            "name": "divsChild",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPartPr(Child):
    class Meta:
        name = "CT_DocPartPr"

    content: list[
        CTDocPartName
        | CtDocPartPrStyle
        | CTDocPartCategory
        | CTDocPartTypes
        | CTDocPartBehaviors
        | CtDocPartPrDescription
        | CTGuid
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "name",
                    "type": ForwardRef("CTDocPartName"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "style",
                    "type": ForwardRef("CtDocPartPrStyle"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "category",
                    "type": ForwardRef("CTDocPartCategory"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "types",
                    "type": ForwardRef("CTDocPartTypes"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "behaviors",
                    "type": ForwardRef("CTDocPartBehaviors"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "description",
                    "type": ForwardRef("CtDocPartPrDescription"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "guid",
                    "type": ForwardRef("CTGuid"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEdnDocProps(CTEdnProps):
    class Meta:
        name = "CT_EdnDocProps"

    endnote: list[CTFtnEdnSepRef] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_occurs": 3,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFFData(Child):
    class Meta:
        name = "CT_FFData"

    content: list[
        CTFFName
        | Enabled
        | CalcOnExit
        | EntryMacro
        | ExitMacro
        | CTFFHelpText
        | CTFFStatusText
        | CTFFCheckBox
        | CTFFDDList
        | CTFFTextInput
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "name",
                    "type": ForwardRef("CTFFName"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "enabled",
                    "type": ForwardRef("Enabled"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "calcOnExit",
                    "type": ForwardRef("CalcOnExit"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "entryMacro",
                    "type": ForwardRef("EntryMacro"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "exitMacro",
                    "type": ForwardRef("ExitMacro"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "helpText",
                    "type": ForwardRef("CTFFHelpText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "statusText",
                    "type": ForwardRef("CTFFStatusText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "checkBox",
                    "type": ForwardRef("CTFFCheckBox"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ddList",
                    "type": ForwardRef("CTFFDDList"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "textInput",
                    "type": ForwardRef("CTFFTextInput"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFrameset(Child):
    class Meta:
        name = "CT_Frameset"

    sz: None | CtFramesetSz = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    frameset_splitbar: None | CTFramesetSplitbar = field(
        default=None,
        metadata={
            "name": "framesetSplitbar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    frame_layout: None | CTFrameLayout = field(
        default=None,
        metadata={
            "name": "frameLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    frameset_or_frame: list[CTFrameset | CTFrame] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "frameset",
                    "type": ForwardRef("CTFrameset"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "frame",
                    "type": ForwardRef("CTFrame"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFtnDocProps(CTFtnProps):
    class Meta:
        name = "CT_FtnDocProps"

    footnote: list[CTFtnEdnSepRef] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_occurs": 3,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOdso(Child):
    class Meta:
        name = "CT_Odso"

    udl: None | CtOdsoUdl = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    table: None | CtOdsoTable = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    src: None | CTRel = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    col_delim: None | CtOdsoColDelim = field(
        default=None,
        metadata={
            "name": "colDelim",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    type_value: None | CTMailMergeSourceType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    f_hdr: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "fHdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    field_map_data: list[CTOdsoFieldMapData] = field(
        default_factory=ChildList,
        metadata={
            "name": "fieldMapData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    recipient_data: list[CTRel] = field(
        default_factory=ChildList,
        metadata={
            "name": "recipientData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRPrChange(CTTrackChange):
    class Meta:
        name = "CT_RPrChange"

    r_pr: None | CtRprChangeRPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblPrBase(Child):
    class Meta:
        name = "CT_TblPrBase"

    tbl_style: None | CtTblPrBaseTblStyle = field(
        default=None,
        metadata={
            "name": "tblStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tblp_pr: None | CTTblPPr = field(
        default=None,
        metadata={
            "name": "tblpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_overlap: None | CTTblOverlap = field(
        default=None,
        metadata={
            "name": "tblOverlap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bidi_visual: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bidiVisual",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_style_row_band_size: None | CtTblPrBaseTblStyleRowBandSize = field(
        default=None,
        metadata={
            "name": "tblStyleRowBandSize",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_style_col_band_size: None | CtTblPrBaseTblStyleColBandSize = field(
        default=None,
        metadata={
            "name": "tblStyleColBandSize",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_w: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tblW",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    jc: None | Jc = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_cell_spacing: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tblCellSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_ind: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tblInd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_borders: None | TblBorders = field(
        default=None,
        metadata={
            "name": "tblBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: None | CTShd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_layout: None | CTTblLayoutType = field(
        default=None,
        metadata={
            "name": "tblLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_cell_mar: None | CTTblCellMar = field(
        default=None,
        metadata={
            "name": "tblCellMar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_look: None | CTTblLook = field(
        default=None,
        metadata={
            "name": "tblLook",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_caption: None | CTString = field(
        default=None,
        metadata={
            "name": "tblCaption",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_description: None | CTString = field(
        default=None,
        metadata={
            "name": "tblDescription",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblPrExBase(Child):
    class Meta:
        name = "CT_TblPrExBase"

    tbl_w: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tblW",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    jc: None | Jc = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_cell_spacing: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tblCellSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_ind: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tblInd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_borders: None | TblBorders = field(
        default=None,
        metadata={
            "name": "tblBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: None | CTShd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_layout: None | CTTblLayoutType = field(
        default=None,
        metadata={
            "name": "tblLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_cell_mar: None | CTTblCellMar = field(
        default=None,
        metadata={
            "name": "tblCellMar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_look: None | CTTblLook = field(
        default=None,
        metadata={
            "name": "tblLook",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_caption: None | CTString = field(
        default=None,
        metadata={
            "name": "tblCaption",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrPrBase(Child):
    class Meta:
        name = "CT_TrPrBase"

    content: list[
        CTCnf
        | CtTrPrBaseDivId
        | CtTrPrBaseGridBefore
        | CtTrPrBaseGridAfter
        | WBefore
        | WAfter
        | CantSplit
        | CTHeight
        | TblHeader
        | TblCellSpacing
        | Jc
        | Hidden
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cnfStyle",
                    "type": ForwardRef("CTCnf"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "divId",
                    "type": ForwardRef("CtTrPrBaseDivId"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "gridBefore",
                    "type": ForwardRef("CtTrPrBaseGridBefore"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "gridAfter",
                    "type": ForwardRef("CtTrPrBaseGridAfter"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "wBefore",
                    "type": ForwardRef("WBefore"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "wAfter",
                    "type": ForwardRef("WAfter"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "cantSplit",
                    "type": ForwardRef("CantSplit"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "trHeight",
                    "type": ForwardRef("CTHeight"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tblHeader",
                    "type": ForwardRef("TblHeader"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tblCellSpacing",
                    "type": ForwardRef("TblCellSpacing"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "jc",
                    "type": ForwardRef("Jc"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hidden",
                    "type": ForwardRef("Hidden"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CtPprBaseNumPr(Child):
    class Meta:
        global_type = False

    ilvl: None | CtPprBaseNumPrIlvl = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_id: None | CtPprBaseNumPrNumId = field(
        default=None,
        metadata={
            "name": "numId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    numbering_change: None | CTTrackChangeNumbering = field(
        default=None,
        metadata={
            "name": "numberingChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ins: None | CTTrackChange = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Fonts(Child):
    class Meta:
        name = "fonts"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    font: list[FontsFont] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class ParaRPrChange(CTTrackChange):
    class Meta:
        name = "CT_ParaRPrChange"

    r_pr: None | CTParaRPrOriginal = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class SectPrBase(Child):
    class Meta:
        name = "CT_SectPrBase"

    footnote_pr: None | CTFtnProps = field(
        default=None,
        metadata={
            "name": "footnotePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    endnote_pr: None | CTEdnProps = field(
        default=None,
        metadata={
            "name": "endnotePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    type_value: None | CtSectPrBaseType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_sz: None | CtSectPrBasePgSz = field(
        default=None,
        metadata={
            "name": "pgSz",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_mar: None | CtSectPrBasePgMar = field(
        default=None,
        metadata={
            "name": "pgMar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    paper_src: None | CTPaperSource = field(
        default=None,
        metadata={
            "name": "paperSrc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_borders: None | CtSectPrBasePgBorders = field(
        default=None,
        metadata={
            "name": "pgBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ln_num_type: None | CTLineNumber = field(
        default=None,
        metadata={
            "name": "lnNumType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_num_type: None | CTPageNumber = field(
        default=None,
        metadata={
            "name": "pgNumType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cols: None | CTColumns = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    form_prot: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "formProt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_align: None | CTVerticalJc = field(
        default=None,
        metadata={
            "name": "vAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_endnote: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noEndnote",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    title_pg: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "titlePg",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    text_direction: None | TextDirection = field(
        default=None,
        metadata={
            "name": "textDirection",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bidi: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rtl_gutter: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "rtlGutter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_grid: None | CTDocGrid = field(
        default=None,
        metadata={
            "name": "docGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    printer_settings: None | CTRel = field(
        default=None,
        metadata={
            "name": "printerSettings",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    footnote_columns: None | W15CtdecimalNumber = field(
        default=None,
        metadata={
            "name": "footnoteColumns",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    rsid_rpr: None | str = field(
        default=None,
        metadata={
            "name": "rsidRPr",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_del: None | str = field(
        default=None,
        metadata={
            "name": "rsidDel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_r: None | str = field(
        default=None,
        metadata={
            "name": "rsidR",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_sect: None | str = field(
        default=None,
        metadata={
            "name": "rsidSect",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TblGrid(TblGridBase):
    class Meta:
        name = "CT_TblGrid"

    tbl_grid_change: None | CTTblGridChange = field(
        default=None,
        metadata={
            "name": "tblGridChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TcPrInner(Child):
    class Meta:
        name = "CT_TcPrInner"

    cnf_style: None | CTCnf = field(
        default=None,
        metadata={
            "name": "cnfStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tc_w: None | TblWidth = field(
        default=None,
        metadata={
            "name": "tcW",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    grid_span: None | CtTcPrInnerGridSpan = field(
        default=None,
        metadata={
            "name": "gridSpan",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    h_merge: None | CtTcPrInnerHMerge = field(
        default=None,
        metadata={
            "name": "hMerge",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_merge: None | CtTcPrInnerVMerge = field(
        default=None,
        metadata={
            "name": "vMerge",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tc_borders: None | CtTcPrInnerTcBorders = field(
        default=None,
        metadata={
            "name": "tcBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: None | CTShd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_wrap: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noWrap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tc_mar: None | TcMar = field(
        default=None,
        metadata={
            "name": "tcMar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    text_direction: None | TextDirection = field(
        default=None,
        metadata={
            "name": "textDirection",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tc_fit_text: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "tcFitText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_align: None | CTVerticalJc = field(
        default=None,
        metadata={
            "name": "vAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hide_mark: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "hideMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cell_ins_or_cell_del_or_cell_merge: None | CellIns | CellDel | CTCellMergeTrackChange = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cellIns",
                    "type": ForwardRef("CellIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "cellDel",
                    "type": ForwardRef("CellDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "cellMerge",
                    "type": ForwardRef("CTCellMergeTrackChange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Recipients(CTRecipients):
    class Meta:
        name = "recipients"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTDivs(Child):
    class Meta:
        name = "CT_Divs"

    div: list[CTDiv] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMailMerge(Child):
    class Meta:
        name = "CT_MailMerge"

    main_document_type: None | CTMailMergeDocType = field(
        default=None,
        metadata={
            "name": "mainDocumentType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    link_to_query: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "linkToQuery",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    data_type: None | CTMailMergeDataType = field(
        default=None,
        metadata={
            "name": "dataType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    connect_string: None | CtMailMergeConnectString = field(
        default=None,
        metadata={
            "name": "connectString",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    query: None | CtMailMergeQuery = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    data_source: None | CTRel = field(
        default=None,
        metadata={
            "name": "dataSource",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    header_source: None | CTRel = field(
        default=None,
        metadata={
            "name": "headerSource",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_suppress_blank_lines: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotSuppressBlankLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    destination: None | CTMailMergeDest = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    address_field_name: None | CtMailMergeAddressFieldName = field(
        default=None,
        metadata={
            "name": "addressFieldName",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mail_subject: None | CtMailMergeMailSubject = field(
        default=None,
        metadata={
            "name": "mailSubject",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mail_as_attachment: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "mailAsAttachment",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    view_merged_data: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "viewMergedData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    active_record: None | CtMailMergeActiveRecord = field(
        default=None,
        metadata={
            "name": "activeRecord",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    check_errors: None | CtMailMergeCheckErrors = field(
        default=None,
        metadata={
            "name": "checkErrors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    odso: None | CTOdso = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMoveBookmark(CTBookmark):
    class Meta:
        name = "CT_MoveBookmark"

    author: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    date: None | XmlDateTime = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSectPrChange(CTTrackChange):
    class Meta:
        name = "CT_SectPrChange"

    sect_pr: None | SectPrBase = field(
        default=None,
        metadata={
            "name": "sectPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblPrChange(CTTrackChange):
    class Meta:
        name = "CT_TblPrChange"

    tbl_pr: None | CTTblPrBase = field(
        default=None,
        metadata={
            "name": "tblPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblPrExChange(CTTrackChange):
    class Meta:
        name = "CT_TblPrExChange"

    tbl_pr_ex: None | CTTblPrExBase = field(
        default=None,
        metadata={
            "name": "tblPrEx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTcPrChange(CTTrackChange):
    class Meta:
        name = "CT_TcPrChange"

    tc_pr: None | TcPrInner = field(
        default=None,
        metadata={
            "name": "tcPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrPrChange(CTTrackChange):
    class Meta:
        name = "CT_TrPrChange"

    tr_pr: None | CTTrPrBase = field(
        default=None,
        metadata={
            "name": "trPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class FldChar(Child):
    class Meta:
        name = "CT_FldChar"

    fld_data_or_ff_data_or_numbering_change: None | Text | CTFFData | CTTrackChangeNumbering = (
        field(
            default=None,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "fldData",
                        "type": ForwardRef("Text"),
                        "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    },
                    {
                        "name": "ffData",
                        "type": ForwardRef("CTFFData"),
                        "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    },
                    {
                        "name": "numberingChange",
                        "type": ForwardRef("CTTrackChangeNumbering"),
                        "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    },
                ),
            },
        )
    )
    fld_char_type: None | STFldCharType = field(
        default=None,
        metadata={
            "name": "fldCharType",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fld_lock: None | bool = field(
        default=None,
        metadata={
            "name": "fldLock",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    dirty: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class PPrBase(Child):
    class Meta:
        name = "CT_PPrBase"

    p_style: None | CtPprBasePStyle = field(
        default=None,
        metadata={
            "name": "pStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    keep_next: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "keepNext",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    keep_lines: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "keepLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    page_break_before: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "pageBreakBefore",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    frame_pr: None | CTFramePr = field(
        default=None,
        metadata={
            "name": "framePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    widow_control: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "widowControl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_pr: None | CtPprBaseNumPr = field(
        default=None,
        metadata={
            "name": "numPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_line_numbers: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressLineNumbers",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    p_bdr: None | CtPprBasePBdr = field(
        default=None,
        metadata={
            "name": "pBdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: None | CTShd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tabs: None | Tabs = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_auto_hyphens: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressAutoHyphens",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    kinsoku: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    word_wrap: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "wordWrap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    overflow_punct: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "overflowPunct",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    top_line_punct: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "topLinePunct",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    auto_space_de: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autoSpaceDE",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    auto_space_dn: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autoSpaceDN",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bidi: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    adjust_right_ind: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "adjustRightInd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    snap_to_grid: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "snapToGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spacing: None | CtPprBaseSpacing = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ind: None | CtPprBaseInd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    contextual_spacing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "contextualSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mirror_indents: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "mirrorIndents",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suppress_overlap: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "suppressOverlap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    jc: None | Jc = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    text_direction: None | TextDirection = field(
        default=None,
        metadata={
            "name": "textDirection",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    text_alignment: None | CtPprBaseTextAlignment = field(
        default=None,
        metadata={
            "name": "textAlignment",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    textbox_tight_wrap: None | CTTextboxTightWrap = field(
        default=None,
        metadata={
            "name": "textboxTightWrap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    outline_lvl: None | CtPprBaseOutlineLvl = field(
        default=None,
        metadata={
            "name": "outlineLvl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    div_id: None | CtPprBaseDivId = field(
        default=None,
        metadata={
            "name": "divId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cnf_style: None | CTCnf = field(
        default=None,
        metadata={
            "name": "cnfStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    collapsed: None | Collapsed = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class ParaRPr(Child):
    class Meta:
        name = "CT_ParaRPr"

    ins: None | CTTrackChange = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    del_value: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "del",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    move_from: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "moveFrom",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    move_to: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "moveTo",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_style: None | RStyle = field(
        default=None,
        metadata={
            "name": "rStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_fonts: None | RFonts = field(
        default=None,
        metadata={
            "name": "rFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b_cs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i_cs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "iCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    caps: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    small_caps: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "smallCaps",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    strike: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dstrike: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    outline: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shadow: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    emboss: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    imprint: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_proof: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noProof",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    snap_to_grid: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "snapToGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vanish: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    web_hidden: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "webHidden",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: None | Color = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spacing: None | CTSignedTwipsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    w: None | CTTextScale = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    kern: None | HpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    position: None | CTSignedHpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz: None | HpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz_cs: None | HpsMeasure = field(
        default=None,
        metadata={
            "name": "szCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    highlight: None | Highlight = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    u: None | U = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    effect: None | CTTextEffect = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bdr: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: None | CTShd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fit_text: None | CTFitText = field(
        default=None,
        metadata={
            "name": "fitText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vert_align: None | CTVerticalAlignRun = field(
        default=None,
        metadata={
            "name": "vertAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rtl: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    em: None | CTEm = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lang: None | CTLanguage = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    east_asian_layout: None | CTEastAsianLayout = field(
        default=None,
        metadata={
            "name": "eastAsianLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spec_vanish: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "specVanish",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    o_math: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "oMath",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    glow: None | CTGlow = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    schemas_microsoft_com_office_word_2010_wordml_shadow: None | Shadow = field(
        default=None,
        metadata={
            "name": "shadow",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    reflection: None | CTReflection = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_outline: None | CTTextOutlineEffect = field(
        default=None,
        metadata={
            "name": "textOutline",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_fill: None | CTFillTextEffect = field(
        default=None,
        metadata={
            "name": "textFill",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    props3d: None | CTProps3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    ligatures: None | CTLigatures = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_form: None | CTNumForm = field(
        default=None,
        metadata={
            "name": "numForm",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_spacing: None | CTNumSpacing = field(
        default=None,
        metadata={
            "name": "numSpacing",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    stylistic_sets: None | CTStylisticSets = field(
        default=None,
        metadata={
            "name": "stylisticSets",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    cntxt_alts: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "cntxtAlts",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    r_pr_change: None | ParaRPrChange = field(
        default=None,
        metadata={
            "name": "rPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RPr(Child):
    class Meta:
        name = "CT_RPr"

    r_style: None | RStyle = field(
        default=None,
        metadata={
            "name": "rStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_fonts: None | RFonts = field(
        default=None,
        metadata={
            "name": "rFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    b_cs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    i_cs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "iCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    caps: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    small_caps: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "smallCaps",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    strike: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    dstrike: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    outline: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shadow: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    emboss: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    imprint: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_proof: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noProof",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    snap_to_grid: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "snapToGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vanish: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    web_hidden: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "webHidden",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    color: None | Color = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spacing: None | CTSignedTwipsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    w: None | CTTextScale = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    kern: None | HpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    position: None | CTSignedHpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz: None | HpsMeasure = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sz_cs: None | HpsMeasure = field(
        default=None,
        metadata={
            "name": "szCs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    highlight: None | Highlight = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    u: None | U = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    effect: None | CTTextEffect = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bdr: None | CTBorder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    shd: None | CTShd = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fit_text: None | CTFitText = field(
        default=None,
        metadata={
            "name": "fitText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    vert_align: None | CTVerticalAlignRun = field(
        default=None,
        metadata={
            "name": "vertAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rtl: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cs: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    em: None | CTEm = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lang: None | CTLanguage = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    east_asian_layout: None | CTEastAsianLayout = field(
        default=None,
        metadata={
            "name": "eastAsianLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    spec_vanish: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "specVanish",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    o_math: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "oMath",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    glow: None | CTGlow = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    schemas_microsoft_com_office_word_2010_wordml_shadow: None | Shadow = field(
        default=None,
        metadata={
            "name": "shadow",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    reflection: None | CTReflection = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_outline: None | CTTextOutlineEffect = field(
        default=None,
        metadata={
            "name": "textOutline",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_fill: None | CTFillTextEffect = field(
        default=None,
        metadata={
            "name": "textFill",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    props3d: None | CTProps3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    ligatures: None | CTLigatures = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_form: None | CTNumForm = field(
        default=None,
        metadata={
            "name": "numForm",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    num_spacing: None | CTNumSpacing = field(
        default=None,
        metadata={
            "name": "numSpacing",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    stylistic_sets: None | CTStylisticSets = field(
        default=None,
        metadata={
            "name": "stylisticSets",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    cntxt_alts: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "cntxtAlts",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    r_pr_change: None | CTRPrChange = field(
        default=None,
        metadata={
            "name": "rPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Del2(CTRPrChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Ins2(CTRPrChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTPPrChange(CTTrackChange):
    class Meta:
        name = "CT_PPrChange"

    p_pr: None | PPrBase = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRPrDefault(Child):
    class Meta:
        name = "CT_RPrDefault"

    r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtEndPr(Child):
    class Meta:
        name = "CT_SdtEndPr"

    r_pr: list[RPr] = field(
        default_factory=ChildList,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSettings(Child):
    class Meta:
        name = "CT_Settings"

    write_protection: None | CTWriteProtection = field(
        default=None,
        metadata={
            "name": "writeProtection",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    view: None | CTView = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    zoom: None | CTZoom = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    remove_personal_information: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "removePersonalInformation",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    remove_date_and_time: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "removeDateAndTime",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_display_page_boundaries: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotDisplayPageBoundaries",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    display_background_shape: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "displayBackgroundShape",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    print_post_script_over_text: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "printPostScriptOverText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    print_fractional_character_width: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "printFractionalCharacterWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    print_forms_data: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "printFormsData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    embed_true_type_fonts: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "embedTrueTypeFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    embed_system_fonts: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "embedSystemFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_subset_fonts: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "saveSubsetFonts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_forms_data: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "saveFormsData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mirror_margins: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "mirrorMargins",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    align_borders_and_edges: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "alignBordersAndEdges",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    borders_do_not_surround_header: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bordersDoNotSurroundHeader",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    borders_do_not_surround_footer: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bordersDoNotSurroundFooter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    gutter_at_top: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "gutterAtTop",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hide_spelling_errors: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "hideSpellingErrors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hide_grammatical_errors: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "hideGrammaticalErrors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    active_writing_style: list[CTWritingStyle] = field(
        default_factory=ChildList,
        metadata={
            "name": "activeWritingStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    proof_state: None | CTProof = field(
        default=None,
        metadata={
            "name": "proofState",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    forms_design: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "formsDesign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    attached_template: None | CTRel = field(
        default=None,
        metadata={
            "name": "attachedTemplate",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    link_styles: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "linkStyles",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    style_pane_format_filter: None | CTStylePaneFilter = field(
        default=None,
        metadata={
            "name": "stylePaneFormatFilter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    style_pane_sort_method: None | CTShortHexNumber = field(
        default=None,
        metadata={
            "name": "stylePaneSortMethod",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    document_type: None | CTDocType = field(
        default=None,
        metadata={
            "name": "documentType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    mail_merge: None | CTMailMerge = field(
        default=None,
        metadata={
            "name": "mailMerge",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    revision_view: None | CTTrackChangesView = field(
        default=None,
        metadata={
            "name": "revisionView",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    track_revisions: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "trackRevisions",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_track_moves: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotTrackMoves",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_track_formatting: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotTrackFormatting",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    document_protection: None | CTDocProtect = field(
        default=None,
        metadata={
            "name": "documentProtection",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    auto_format_override: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autoFormatOverride",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    style_lock_theme: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "styleLockTheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    style_lock_qfset: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "styleLockQFSet",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    default_tab_stop: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "defaultTabStop",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    auto_hyphenation: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autoHyphenation",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    consecutive_hyphen_limit: None | CtSettingsConsecutiveHyphenLimit = field(
        default=None,
        metadata={
            "name": "consecutiveHyphenLimit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hyphenation_zone: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "hyphenationZone",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_hyphenate_caps: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotHyphenateCaps",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    show_envelope: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "showEnvelope",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    summary_length: None | CtSettingsSummaryLength = field(
        default=None,
        metadata={
            "name": "summaryLength",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    click_and_type_style: None | CtSettingsClickAndTypeStyle = field(
        default=None,
        metadata={
            "name": "clickAndTypeStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    default_table_style: None | CtSettingsDefaultTableStyle = field(
        default=None,
        metadata={
            "name": "defaultTableStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    even_and_odd_headers: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "evenAndOddHeaders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    book_fold_rev_printing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bookFoldRevPrinting",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    book_fold_printing: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "bookFoldPrinting",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    book_fold_printing_sheets: None | CtSettingsBookFoldPrintingSheets = field(
        default=None,
        metadata={
            "name": "bookFoldPrintingSheets",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    drawing_grid_horizontal_spacing: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "drawingGridHorizontalSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    drawing_grid_vertical_spacing: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "drawingGridVerticalSpacing",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    display_horizontal_drawing_grid_every: None | CtSettingsDisplayHorizontalDrawingGridEvery = (
        field(
            default=None,
            metadata={
                "name": "displayHorizontalDrawingGridEvery",
                "type": "Element",
                "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            },
        )
    )
    display_vertical_drawing_grid_every: None | CtSettingsDisplayVerticalDrawingGridEvery = field(
        default=None,
        metadata={
            "name": "displayVerticalDrawingGridEvery",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_use_margins_for_drawing_grid_origin: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotUseMarginsForDrawingGridOrigin",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    drawing_grid_horizontal_origin: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "drawingGridHorizontalOrigin",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    drawing_grid_vertical_origin: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "drawingGridVerticalOrigin",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_shade_form_data: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotShadeFormData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_punctuation_kerning: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noPunctuationKerning",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    character_spacing_control: None | CTCharacterSpacing = field(
        default=None,
        metadata={
            "name": "characterSpacingControl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    print_two_on_one: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "printTwoOnOne",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    strict_first_and_last_chars: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "strictFirstAndLastChars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_line_breaks_after: None | CTKinsoku = field(
        default=None,
        metadata={
            "name": "noLineBreaksAfter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_line_breaks_before: None | CTKinsoku = field(
        default=None,
        metadata={
            "name": "noLineBreaksBefore",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_preview_picture: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "savePreviewPicture",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_validate_against_schema: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotValidateAgainstSchema",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_invalid_xml: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "saveInvalidXml",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ignore_mixed_content: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "ignoreMixedContent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    always_show_placeholder_text: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "alwaysShowPlaceholderText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_demarcate_invalid_xml: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotDemarcateInvalidXml",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_xml_data_only: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "saveXmlDataOnly",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    use_xsltwhen_saving: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "useXSLTWhenSaving",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_through_xslt: None | CTSaveThroughXslt = field(
        default=None,
        metadata={
            "name": "saveThroughXslt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    show_xmltags: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "showXMLTags",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    always_merge_empty_namespace: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "alwaysMergeEmptyNamespace",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    update_fields: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "updateFields",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    hdr_shape_defaults: None | CTShapeDefaults = field(
        default=None,
        metadata={
            "name": "hdrShapeDefaults",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    footnote_pr: None | CTFtnDocProps = field(
        default=None,
        metadata={
            "name": "footnotePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    endnote_pr: None | CTEdnDocProps = field(
        default=None,
        metadata={
            "name": "endnotePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    compat: None | CTCompat = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_vars: None | CTDocVars = field(
        default=None,
        metadata={
            "name": "docVars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsids: None | CTDocRsids = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    math_pr: None | MathPr = field(
        default=None,
        metadata={
            "name": "mathPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ui_compat97_to2003: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "uiCompat97To2003",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    attached_schema: list[CtSettingsAttachedSchema] = field(
        default_factory=ChildList,
        metadata={
            "name": "attachedSchema",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    theme_font_lang: None | CTLanguage = field(
        default=None,
        metadata={
            "name": "themeFontLang",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    clr_scheme_mapping: None | CTColorSchemeMapping = field(
        default=None,
        metadata={
            "name": "clrSchemeMapping",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_include_subdocs_in_stats: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotIncludeSubdocsInStats",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_auto_compress_pictures: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotAutoCompressPictures",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    force_upgrade: None | CtSettingsForceUpgrade = field(
        default=None,
        metadata={
            "name": "forceUpgrade",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    captions: None | CTCaptions = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    read_mode_ink_lock_down: None | CTReadingModeInkLockDown = field(
        default=None,
        metadata={
            "name": "readModeInkLockDown",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    smart_tag_type: list[CTSmartTagType] = field(
        default_factory=ChildList,
        metadata={
            "name": "smartTagType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    schema_library: None | SchemaLibrary = field(
        default=None,
        metadata={
            "name": "schemaLibrary",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
        },
    )
    shape_defaults: None | CTShapeDefaults = field(
        default=None,
        metadata={
            "name": "shapeDefaults",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_embed_smart_tags: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotEmbedSmartTags",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    decimal_symbol: None | CtSettingsDecimalSymbol = field(
        default=None,
        metadata={
            "name": "decimalSymbol",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    list_separator: None | CtSettingsListSeparator = field(
        default=None,
        metadata={
            "name": "listSeparator",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_id: None | W14DocId = field(
        default=None,
        metadata={
            "name": "docId",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chart_tracking_ref_based: None | W15BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "chartTrackingRefBased",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    schemas_microsoft_com_office_word_2012_wordml_doc_id: None | W15DocId = field(
        default=None,
        metadata={
            "name": "docId",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    conflict_mode: None | ConflictMode = field(
        default=None,
        metadata={
            "name": "conflictMode",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    discard_image_editing_data: None | DiscardImageEditingData = field(
        default=None,
        metadata={
            "name": "discardImageEditingData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    default_image_dpi: None | DefaultImageDpi = field(
        default=None,
        metadata={
            "name": "defaultImageDpi",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblPrEx(CTTblPrExBase):
    class Meta:
        name = "CT_TblPrEx"

    tbl_pr_ex_change: None | CTTblPrExChange = field(
        default=None,
        metadata={
            "name": "tblPrExChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWebSettings(Child):
    class Meta:
        name = "CT_WebSettings"

    frameset: None | CTFrameset = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    divs: None | CTDivs = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    encoding: None | CtWebSettingsEncoding = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    optimize_for_browser: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "optimizeForBrowser",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rely_on_vml: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "relyOnVML",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    allow_png: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "allowPNG",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_rely_on_css: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotRelyOnCSS",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_save_as_single_file: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotSaveAsSingleFile",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_organize_in_folder: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotOrganizeInFolder",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    do_not_use_long_file_names: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotUseLongFileNames",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pixels_per_inch: None | CtWebSettingsPixelsPerInch = field(
        default=None,
        metadata={
            "name": "pixelsPerInch",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    target_screen_sz: None | CTTargetScreenSz = field(
        default=None,
        metadata={
            "name": "targetScreenSz",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    save_smart_tags_as_xml: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "saveSmartTagsAsXml",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class SdtPr(Child):
    class Meta:
        name = "CT_SdtPr"

    content: list[
        RPr
        | CtSdtPrAlias
        | CTLock
        | CTPlaceholder
        | ShowingPlcHdr
        | CTDataBinding
        | Temporary
        | Id
        | Tag
        | CtSdtPrEquation
        | CTSdtComboBox
        | CTSdtDate
        | DocPartObj
        | DocPartList
        | CTSdtDropDownList
        | CtSdtPrPicture
        | CtSdtPrRichText
        | CTSdtText
        | CtSdtPrCitation
        | CtSdtPrGroup
        | CtSdtPrBibliography
        | Checkbox
        | W14Ctempty
        | Appearance
        | W15Ctcolor
        | W15CtdataBinding
        | RepeatingSection
        | W15Ctempty
        | WebExtensionCreated
        | WebExtensionLinked
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "rPr",
                    "type": ForwardRef("RPr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "alias",
                    "type": ForwardRef("CtSdtPrAlias"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("CTLock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "placeholder",
                    "type": ForwardRef("CTPlaceholder"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "showingPlcHdr",
                    "type": ForwardRef("ShowingPlcHdr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dataBinding",
                    "type": ForwardRef("CTDataBinding"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "temporary",
                    "type": ForwardRef("Temporary"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "id",
                    "type": ForwardRef("Id"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tag",
                    "type": ForwardRef("Tag"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "equation",
                    "type": ForwardRef("CtSdtPrEquation"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "comboBox",
                    "type": ForwardRef("CTSdtComboBox"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "date",
                    "type": ForwardRef("CTSdtDate"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "docPartObj",
                    "type": ForwardRef("DocPartObj"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "docPartList",
                    "type": ForwardRef("DocPartList"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dropDownList",
                    "type": ForwardRef("CTSdtDropDownList"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "picture",
                    "type": ForwardRef("CtSdtPrPicture"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "richText",
                    "type": ForwardRef("CtSdtPrRichText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "text",
                    "type": ForwardRef("CTSdtText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "citation",
                    "type": ForwardRef("CtSdtPrCitation"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "group",
                    "type": ForwardRef("CtSdtPrGroup"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bibliography",
                    "type": ForwardRef("CtSdtPrBibliography"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "checkbox",
                    "type": ForwardRef("Checkbox"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "entityPicker",
                    "type": ForwardRef("W14Ctempty"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "appearance",
                    "type": ForwardRef("Appearance"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
                {
                    "name": "color",
                    "type": ForwardRef("W15Ctcolor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
                {
                    "name": "dataBinding",
                    "type": ForwardRef("W15CtdataBinding"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
                {
                    "name": "repeatingSection",
                    "type": ForwardRef("RepeatingSection"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
                {
                    "name": "repeatingSectionItem",
                    "type": ForwardRef("W15Ctempty"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
                {
                    "name": "webExtensionCreated",
                    "type": ForwardRef("WebExtensionCreated"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
                {
                    "name": "webExtensionLinked",
                    "type": ForwardRef("WebExtensionLinked"),
                    "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class SectPr(Child):
    class Meta:
        name = "CT_SectPr"

    header_reference_or_footer_reference: list[HeaderReference | FooterReference] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "headerReference",
                    "type": ForwardRef("HeaderReference"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "max_occurs": 6,
                },
                {
                    "name": "footerReference",
                    "type": ForwardRef("FooterReference"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "max_occurs": 6,
                },
            ),
            "max_occurs": 6,
        },
    )
    footnote_pr: None | CTFtnProps = field(
        default=None,
        metadata={
            "name": "footnotePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    endnote_pr: None | CTEdnProps = field(
        default=None,
        metadata={
            "name": "endnotePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    type_value: None | CtSectPrType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_sz: None | CtSectPrPgSz = field(
        default=None,
        metadata={
            "name": "pgSz",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_mar: None | CtSectPrPgMar = field(
        default=None,
        metadata={
            "name": "pgMar",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    paper_src: None | CTPaperSource = field(
        default=None,
        metadata={
            "name": "paperSrc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_borders: None | CtSectPrPgBorders = field(
        default=None,
        metadata={
            "name": "pgBorders",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ln_num_type: None | CTLineNumber = field(
        default=None,
        metadata={
            "name": "lnNumType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    pg_num_type: None | CTPageNumber = field(
        default=None,
        metadata={
            "name": "pgNumType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    cols: None | CTColumns = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    form_prot: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "formProt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    v_align: None | CTVerticalJc = field(
        default=None,
        metadata={
            "name": "vAlign",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    no_endnote: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "noEndnote",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    title_pg: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "titlePg",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    text_direction: None | TextDirection = field(
        default=None,
        metadata={
            "name": "textDirection",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    bidi: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rtl_gutter: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "rtlGutter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_grid: None | CTDocGrid = field(
        default=None,
        metadata={
            "name": "docGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    printer_settings: None | CTRel = field(
        default=None,
        metadata={
            "name": "printerSettings",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    footnote_columns: None | W15CtdecimalNumber = field(
        default=None,
        metadata={
            "name": "footnoteColumns",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    sect_pr_change: None | CTSectPrChange = field(
        default=None,
        metadata={
            "name": "sectPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_rpr: None | str = field(
        default=None,
        metadata={
            "name": "rsidRPr",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_del: None | str = field(
        default=None,
        metadata={
            "name": "rsidDel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_r: None | str = field(
        default=None,
        metadata={
            "name": "rsidR",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_sect: None | str = field(
        default=None,
        metadata={
            "name": "rsidSect",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TblPr(CTTblPrBase):
    class Meta:
        name = "CT_TblPr"

    tbl_pr_change: None | CTTblPrChange = field(
        default=None,
        metadata={
            "name": "tblPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TcPr(TcPrInner):
    class Meta:
        name = "CT_TcPr"

    tc_pr_change: None | CTTcPrChange = field(
        default=None,
        metadata={
            "name": "tcPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class TrPr(CTTrPrBase):
    class Meta:
        name = "CT_TrPr"

    ins: None | CTTrackChange = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    del_value: None | CTTrackChange = field(
        default=None,
        metadata={
            "name": "del",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tr_pr_change: None | CTTrPrChange = field(
        default=None,
        metadata={
            "name": "trPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class DocDefaultsRPrDefault(Child):
    class Meta:
        global_type = False

    r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class MoveFromRangeStart1(CTMoveBookmark):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class MoveFromRangeStart2(CTMoveBookmark):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class MoveToRangeStart1(CTMoveBookmark):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class MoveToRangeStart2(CTMoveBookmark):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class PPr(PPrBase):
    class Meta:
        name = "CT_PPr"

    r_pr: None | ParaRPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sect_pr: None | SectPr = field(
        default=None,
        metadata={
            "name": "sectPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    p_pr_change: None | CTPPrChange = field(
        default=None,
        metadata={
            "name": "pPrChange",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RunTrackChange(CTTrackChange):
    class Meta:
        name = "CT_RunTrackChange"

    content: list[
        CTAcc
        | CTBar
        | CTBox
        | CTBorderBox
        | CTD
        | CTEqArr
        | CTF
        | CTFunc
        | CTGroupChr
        | CTLimLow
        | CTLimUpp
        | CTM
        | CTNary
        | CTPhant
        | CTRad
        | CTSPre
        | CTSSub
        | CTSSubSup
        | CTSSup
        | CTR
        | CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "acc",
                    "type": ForwardRef("CTAcc"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "bar",
                    "type": ForwardRef("CTBar"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "box",
                    "type": ForwardRef("CTBox"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "borderBox",
                    "type": ForwardRef("CTBorderBox"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "d",
                    "type": ForwardRef("CTD"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "eqArr",
                    "type": ForwardRef("CTEqArr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "f",
                    "type": ForwardRef("CTF"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "func",
                    "type": ForwardRef("CTFunc"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "groupChr",
                    "type": ForwardRef("CTGroupChr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "limLow",
                    "type": ForwardRef("CTLimLow"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "limUpp",
                    "type": ForwardRef("CTLimUpp"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "m",
                    "type": ForwardRef("CTM"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "nary",
                    "type": ForwardRef("CTNary"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "phant",
                    "type": ForwardRef("CTPhant"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "rad",
                    "type": ForwardRef("CTRad"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sPre",
                    "type": ForwardRef("CTSPre"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSub",
                    "type": ForwardRef("CTSSub"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSubSup",
                    "type": ForwardRef("CTSSubSup"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSup",
                    "type": ForwardRef("CTSSup"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "r",
                    "type": ForwardRef("CTR"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Settings(CTSettings):
    class Meta:
        name = "settings"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class WebSettings(CTWebSettings):
    class Meta:
        name = "webSettings"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTPPrDefault(Child):
    class Meta:
        name = "CT_PPrDefault"

    p_pr: None | PPr = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTblStylePr(Child):
    class Meta:
        name = "CT_TblStylePr"

    p_pr: None | PPr = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_pr: None | CTTblPrBase = field(
        default=None,
        metadata={
            "name": "tblPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tr_pr: None | TrPr = field(
        default=None,
        metadata={
            "name": "trPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tc_pr: None | TcPr = field(
        default=None,
        metadata={
            "name": "tcPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    type_value: None | STTblStyleOverrideType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Lvl(Child):
    class Meta:
        name = "CT_Lvl"

    start: None | CtLvlStart = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_fmt: None | NumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl_restart: None | CtLvlLvlRestart = field(
        default=None,
        metadata={
            "name": "lvlRestart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    p_style: None | CtLvlPStyle = field(
        default=None,
        metadata={
            "name": "pStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    is_lgl: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "isLgl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    suff: None | CtLvlSuff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl_text: None | CtLvlLvlText = field(
        default=None,
        metadata={
            "name": "lvlText",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl_pic_bullet_id: None | CtLvlLvlPicBulletId = field(
        default=None,
        metadata={
            "name": "lvlPicBulletId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    legacy: None | CtLvlLegacy = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl_jc: None | Jc = field(
        default=None,
        metadata={
            "name": "lvlJc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    p_pr: None | PPr = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ilvl: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tplc: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tentative: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class DocDefaultsPPrDefault(Child):
    class Meta:
        global_type = False

    p_pr: None | PPr = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class MoveFrom1(RunTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class MoveFrom2(RunTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class MoveTo1(RunTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class MoveTo2(RunTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTRubyContent(Child):
    class Meta:
        name = "CT_RubyContent"

    content: list[
        R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSimpleField(Child):
    class Meta:
        name = "CT_SimpleField"

    fld_data: None | Text = field(
        default=None,
        metadata={
            "name": "fldData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtSimpleFieldHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtSimpleFieldHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    instr: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    fld_lock: None | bool = field(
        default=None,
        metadata={
            "name": "fldLock",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    dirty: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class DocDefaults(Child):
    class Meta:
        name = "docDefaults"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    r_pr_default: None | DocDefaultsRPrDefault = field(
        default=None,
        metadata={
            "name": "rPrDefault",
            "type": "Element",
        },
    )
    p_pr_default: None | DocDefaultsPPrDefault = field(
        default=None,
        metadata={
            "name": "pPrDefault",
            "type": "Element",
        },
    )


@dataclass(slots=True, kw_only=True)
class Style(Child):
    class Meta:
        name = "style"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    name: None | StyleName = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    aliases: None | StyleAliases = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    based_on: None | StyleBasedOn = field(
        default=None,
        metadata={
            "name": "basedOn",
            "type": "Element",
        },
    )
    next: None | StyleNext = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    link: None | StyleLink = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    auto_redefine: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "autoRedefine",
            "type": "Element",
        },
    )
    hidden: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    ui_priority: None | StyleUiPriority = field(
        default=None,
        metadata={
            "name": "uiPriority",
            "type": "Element",
        },
    )
    semi_hidden: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "semiHidden",
            "type": "Element",
        },
    )
    unhide_when_used: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "unhideWhenUsed",
            "type": "Element",
        },
    )
    q_format: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "qFormat",
            "type": "Element",
        },
    )
    locked: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    personal: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    personal_compose: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "personalCompose",
            "type": "Element",
        },
    )
    personal_reply: None | BooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "personalReply",
            "type": "Element",
        },
    )
    rsid: None | CTLongHexNumber = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    p_pr: None | PPr = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
        },
    )
    r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
        },
    )
    tbl_pr: None | CTTblPrBase = field(
        default=None,
        metadata={
            "name": "tblPr",
            "type": "Element",
        },
    )
    tr_pr: None | TrPr = field(
        default=None,
        metadata={
            "name": "trPr",
            "type": "Element",
        },
    )
    tc_pr: None | TcPr = field(
        default=None,
        metadata={
            "name": "tcPr",
            "type": "Element",
        },
    )
    tbl_style_pr: list[CTTblStylePr] = field(
        default_factory=ChildList,
        metadata={
            "name": "tblStylePr",
            "type": "Element",
        },
    )
    type_value: None | StyleType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    style_id: None | str = field(
        default=None,
        metadata={
            "name": "styleId",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    default: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "false",
        },
    )
    custom_style: None | bool = field(
        default=None,
        metadata={
            "name": "customStyle",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingAbstractNum(Child):
    class Meta:
        global_type = False

    nsid: None | CTLongHexNumber = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    multi_level_type: None | NumberingAbstractNumMultiLevelType = field(
        default=None,
        metadata={
            "name": "multiLevelType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tmpl: None | CTLongHexNumber = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    name: None | NumberingAbstractNumName = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    style_link: None | NumberingAbstractNumStyleLink = field(
        default=None,
        metadata={
            "name": "styleLink",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    num_style_link: None | NumberingAbstractNumNumStyleLink = field(
        default=None,
        metadata={
            "name": "numStyleLink",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl: list[Lvl] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_occurs": 9,
        },
    )
    abstract_num_id: None | int = field(
        default=None,
        metadata={
            "name": "abstractNumId",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingNumLvlOverride(Child):
    class Meta:
        global_type = False

    start_override: None | NumberingNumLvlOverrideStartOverride = field(
        default=None,
        metadata={
            "name": "startOverride",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl: None | Lvl = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ilvl: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRuby(Child):
    class Meta:
        name = "CT_Ruby"

    ruby_pr: None | CTRubyPr = field(
        default=None,
        metadata={
            "name": "rubyPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rt: None | CTRubyContent = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ruby_base: None | CTRubyContent = field(
        default=None,
        metadata={
            "name": "rubyBase",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Styles(Child):
    class Meta:
        name = "styles"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    doc_defaults: None | DocDefaults = field(
        default=None,
        metadata={
            "name": "docDefaults",
            "type": "Element",
        },
    )
    latent_styles: None | StylesLatentStyles = field(
        default=None,
        metadata={
            "name": "latentStyles",
            "type": "Element",
        },
    )
    style: list[Style] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class NumberingNum(Child):
    class Meta:
        global_type = False

    abstract_num_id: None | NumberingNumAbstractNumId = field(
        default=None,
        metadata={
            "name": "abstractNumId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    lvl_override: list[NumberingNumLvlOverride] = field(
        default_factory=ChildList,
        metadata={
            "name": "lvlOverride",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "max_occurs": 9,
        },
    )
    num_id: None | int = field(
        default=None,
        metadata={
            "name": "numId",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMathRunTrackChange(CTTrackChange):
    class Meta:
        name = "CT_MathRunTrackChange"

    r_pr: None | CTRPR = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    schemas_openxmlformats_org_wordprocessingml_2006_main_r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        Br
        | RT
        | DelText
        | RInstrText
        | RDelInstrText
        | CtMathRunTrackChangeNoBreakHyphen
        | CtMathRunTrackChangeSoftHyphen
        | CtMathRunTrackChangeDayShort
        | CtMathRunTrackChangeMonthShort
        | CtMathRunTrackChangeYearShort
        | CtMathRunTrackChangeDayLong
        | CtMathRunTrackChangeMonthLong
        | CtMathRunTrackChangeYearLong
        | CtMathRunTrackChangeAnnotationRef
        | CtMathRunTrackChangeFootnoteRef
        | CtMathRunTrackChangeEndnoteRef
        | CtMathRunTrackChangeSeparator
        | CtMathRunTrackChangeContinuationSeparator
        | CtMathRunTrackChangeSym
        | CtMathRunTrackChangePgNum
        | CtMathRunTrackChangeCr
        | CtMathRunTrackChangeTab
        | CTObject
        | Pict
        | FldChar
        | CTRuby
        | FootnoteReference2
        | EndnoteReference2
        | CtMathRunTrackChangeCommentReference
        | Drawing
        | CtMathRunTrackChangePtab
        | CtMathRunTrackChangeLastRenderedPageBreak
        | AlternateContent
        | CTText
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "br",
                    "type": ForwardRef("Br"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "t",
                    "type": ForwardRef("RT"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "delText",
                    "type": ForwardRef("DelText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "instrText",
                    "type": ForwardRef("RInstrText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "delInstrText",
                    "type": ForwardRef("RDelInstrText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "noBreakHyphen",
                    "type": ForwardRef("CtMathRunTrackChangeNoBreakHyphen"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "softHyphen",
                    "type": ForwardRef("CtMathRunTrackChangeSoftHyphen"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dayShort",
                    "type": ForwardRef("CtMathRunTrackChangeDayShort"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "monthShort",
                    "type": ForwardRef("CtMathRunTrackChangeMonthShort"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "yearShort",
                    "type": ForwardRef("CtMathRunTrackChangeYearShort"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dayLong",
                    "type": ForwardRef("CtMathRunTrackChangeDayLong"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "monthLong",
                    "type": ForwardRef("CtMathRunTrackChangeMonthLong"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "yearLong",
                    "type": ForwardRef("CtMathRunTrackChangeYearLong"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "annotationRef",
                    "type": ForwardRef("CtMathRunTrackChangeAnnotationRef"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "footnoteRef",
                    "type": ForwardRef("CtMathRunTrackChangeFootnoteRef"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "endnoteRef",
                    "type": ForwardRef("CtMathRunTrackChangeEndnoteRef"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "separator",
                    "type": ForwardRef("CtMathRunTrackChangeSeparator"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "continuationSeparator",
                    "type": ForwardRef("CtMathRunTrackChangeContinuationSeparator"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sym",
                    "type": ForwardRef("CtMathRunTrackChangeSym"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "pgNum",
                    "type": ForwardRef("CtMathRunTrackChangePgNum"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "cr",
                    "type": ForwardRef("CtMathRunTrackChangeCr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tab",
                    "type": ForwardRef("CtMathRunTrackChangeTab"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "object",
                    "type": ForwardRef("CTObject"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "pict",
                    "type": ForwardRef("Pict"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "fldChar",
                    "type": ForwardRef("FldChar"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ruby",
                    "type": ForwardRef("CTRuby"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "footnoteReference",
                    "type": ForwardRef("FootnoteReference2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "endnoteReference",
                    "type": ForwardRef("EndnoteReference2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentReference",
                    "type": ForwardRef("CtMathRunTrackChangeCommentReference"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "drawing",
                    "type": ForwardRef("Drawing"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ptab",
                    "type": ForwardRef("CtMathRunTrackChangePtab"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "lastRenderedPageBreak",
                    "type": ForwardRef("CtMathRunTrackChangeLastRenderedPageBreak"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "AlternateContent",
                    "type": ForwardRef("AlternateContent"),
                    "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                },
                {
                    "name": "t",
                    "type": ForwardRef("CTText"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Numbering(Child):
    class Meta:
        name = "numbering"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    num_pic_bullet: list[NumberingNumPicBullet] = field(
        default_factory=ChildList,
        metadata={
            "name": "numPicBullet",
            "type": "Element",
        },
    )
    abstract_num: list[NumberingAbstractNum] = field(
        default_factory=ChildList,
        metadata={
            "name": "abstractNum",
            "type": "Element",
        },
    )
    num: list[NumberingNum] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
        },
    )
    num_id_mac_at_cleanup: None | NumberingNumIdMacAtCleanup = field(
        default=None,
        metadata={
            "name": "numIdMacAtCleanup",
            "type": "Element",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class R(Child):
    class Meta:
        name = "r"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    r_pr: None | RPr = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
        },
    )
    content: list[
        Br
        | RT
        | DelText
        | RInstrText
        | RDelInstrText
        | RNoBreakHyphen
        | RSoftHyphen
        | RDayShort
        | RMonthShort
        | RYearShort
        | RDayLong
        | RMonthLong
        | RYearLong
        | RAnnotationRef
        | RFootnoteRef
        | REndnoteRef
        | RSeparator
        | RContinuationSeparator
        | RSym
        | RPgNum
        | RCr
        | RTab
        | CTObject
        | Pict
        | FldChar
        | CTRuby
        | FootnoteReference2
        | EndnoteReference2
        | RCommentReference
        | Drawing
        | RPtab
        | RLastRenderedPageBreak
        | AlternateContent
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "br",
                    "type": ForwardRef("Br"),
                },
                {
                    "name": "t",
                    "type": ForwardRef("RT"),
                },
                {
                    "name": "delText",
                    "type": ForwardRef("DelText"),
                },
                {
                    "name": "instrText",
                    "type": ForwardRef("RInstrText"),
                },
                {
                    "name": "delInstrText",
                    "type": ForwardRef("RDelInstrText"),
                },
                {
                    "name": "noBreakHyphen",
                    "type": ForwardRef("RNoBreakHyphen"),
                },
                {
                    "name": "softHyphen",
                    "type": ForwardRef("RSoftHyphen"),
                },
                {
                    "name": "dayShort",
                    "type": ForwardRef("RDayShort"),
                },
                {
                    "name": "monthShort",
                    "type": ForwardRef("RMonthShort"),
                },
                {
                    "name": "yearShort",
                    "type": ForwardRef("RYearShort"),
                },
                {
                    "name": "dayLong",
                    "type": ForwardRef("RDayLong"),
                },
                {
                    "name": "monthLong",
                    "type": ForwardRef("RMonthLong"),
                },
                {
                    "name": "yearLong",
                    "type": ForwardRef("RYearLong"),
                },
                {
                    "name": "annotationRef",
                    "type": ForwardRef("RAnnotationRef"),
                },
                {
                    "name": "footnoteRef",
                    "type": ForwardRef("RFootnoteRef"),
                },
                {
                    "name": "endnoteRef",
                    "type": ForwardRef("REndnoteRef"),
                },
                {
                    "name": "separator",
                    "type": ForwardRef("RSeparator"),
                },
                {
                    "name": "continuationSeparator",
                    "type": ForwardRef("RContinuationSeparator"),
                },
                {
                    "name": "sym",
                    "type": ForwardRef("RSym"),
                },
                {
                    "name": "pgNum",
                    "type": ForwardRef("RPgNum"),
                },
                {
                    "name": "cr",
                    "type": ForwardRef("RCr"),
                },
                {
                    "name": "tab",
                    "type": ForwardRef("RTab"),
                },
                {
                    "name": "object",
                    "type": ForwardRef("CTObject"),
                },
                {
                    "name": "pict",
                    "type": ForwardRef("Pict"),
                },
                {
                    "name": "fldChar",
                    "type": ForwardRef("FldChar"),
                },
                {
                    "name": "ruby",
                    "type": ForwardRef("CTRuby"),
                },
                {
                    "name": "footnoteReference",
                    "type": ForwardRef("FootnoteReference2"),
                },
                {
                    "name": "endnoteReference",
                    "type": ForwardRef("EndnoteReference2"),
                },
                {
                    "name": "commentReference",
                    "type": ForwardRef("RCommentReference"),
                },
                {
                    "name": "drawing",
                    "type": ForwardRef("Drawing"),
                },
                {
                    "name": "ptab",
                    "type": ForwardRef("RPtab"),
                },
                {
                    "name": "lastRenderedPageBreak",
                    "type": ForwardRef("RLastRenderedPageBreak"),
                },
                {
                    "name": "AlternateContent",
                    "type": ForwardRef("AlternateContent"),
                    "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                },
            ),
        },
    )
    rsid_rpr: None | str = field(
        default=None,
        metadata={
            "name": "rsidRPr",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_del: None | str = field(
        default=None,
        metadata={
            "name": "rsidDel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_r: None | str = field(
        default=None,
        metadata={
            "name": "rsidR",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtDir(Child):
    class Meta:
        name = "CT_Dir"

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtDirHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtDirHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    val: None | STDirection = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Del1(CTMathRunTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Ins1(CTMathRunTrackChange):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CtBdo(Child):
    class Meta:
        name = "CT_Bdo"

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtBdoHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtBdoHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    val: None | STDirection = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtContentRun(Child):
    class Meta:
        name = "CT_SdtContentRun"

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtSdtContentRunHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtSdtContentRunHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class SdtRun(Child):
    class Meta:
        name = "CT_SdtRun"

    sdt_pr: None | SdtPr = field(
        default=None,
        metadata={
            "name": "sdtPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sdt_end_pr: None | CTSdtEndPr = field(
        default=None,
        metadata={
            "name": "sdtEndPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sdt_content: None | CTSdtContentRun = field(
        default=None,
        metadata={
            "name": "sdtContent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomXmlRun(Child):
    class Meta:
        name = "CT_CustomXmlRun"

    custom_xml_pr: None | CTCustomXmlPr = field(
        default=None,
        metadata={
            "name": "customXmlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtCustomXmlRunHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtCustomXmlRunHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    element: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSmartTagRun(Child):
    class Meta:
        name = "CT_SmartTagRun"

    smart_tag_pr: None | CTSmartTagPr = field(
        default=None,
        metadata={
            "name": "smartTagPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtSmartTagRunHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtSmartTagRunHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    element: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class RunIns(CTTrackChange):
    class Meta:
        name = "ins"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAcc
        | CTBar
        | CTBox
        | CTBorderBox
        | CTD
        | CTEqArr
        | CTF
        | CTFunc
        | CTGroupChr
        | CTLimLow
        | CTLimUpp
        | CTM
        | CTNary
        | CTPhant
        | CTRad
        | CTSPre
        | CTSSub
        | CTSSubSup
        | CTSSup
        | CTR
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "acc",
                    "type": ForwardRef("CTAcc"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "bar",
                    "type": ForwardRef("CTBar"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "box",
                    "type": ForwardRef("CTBox"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "borderBox",
                    "type": ForwardRef("CTBorderBox"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "d",
                    "type": ForwardRef("CTD"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "eqArr",
                    "type": ForwardRef("CTEqArr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "f",
                    "type": ForwardRef("CTF"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "func",
                    "type": ForwardRef("CTFunc"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "groupChr",
                    "type": ForwardRef("CTGroupChr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "limLow",
                    "type": ForwardRef("CTLimLow"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "limUpp",
                    "type": ForwardRef("CTLimUpp"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "m",
                    "type": ForwardRef("CTM"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "nary",
                    "type": ForwardRef("CTNary"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "phant",
                    "type": ForwardRef("CTPhant"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "rad",
                    "type": ForwardRef("CTRad"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sPre",
                    "type": ForwardRef("CTSPre"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSub",
                    "type": ForwardRef("CTSSub"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSubSup",
                    "type": ForwardRef("CTSSubSup"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSup",
                    "type": ForwardRef("CTSSup"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "r",
                    "type": ForwardRef("CTR"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class RunDel(CTTrackChange):
    class Meta:
        name = "del"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAcc
        | CTBar
        | CTBox
        | CTBorderBox
        | CTD
        | CTEqArr
        | CTF
        | CTFunc
        | CTGroupChr
        | CTLimLow
        | CTLimUpp
        | CTM
        | CTNary
        | CTPhant
        | CTRad
        | CTSPre
        | CTSSub
        | CTSSubSup
        | CTSSup
        | CTR
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "acc",
                    "type": ForwardRef("CTAcc"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "bar",
                    "type": ForwardRef("CTBar"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "box",
                    "type": ForwardRef("CTBox"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "borderBox",
                    "type": ForwardRef("CTBorderBox"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "d",
                    "type": ForwardRef("CTD"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "eqArr",
                    "type": ForwardRef("CTEqArr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "f",
                    "type": ForwardRef("CTF"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "func",
                    "type": ForwardRef("CTFunc"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "groupChr",
                    "type": ForwardRef("CTGroupChr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "limLow",
                    "type": ForwardRef("CTLimLow"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "limUpp",
                    "type": ForwardRef("CTLimUpp"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "m",
                    "type": ForwardRef("CTM"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "nary",
                    "type": ForwardRef("CTNary"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "phant",
                    "type": ForwardRef("CTPhant"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "rad",
                    "type": ForwardRef("CTRad"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sPre",
                    "type": ForwardRef("CTSPre"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSub",
                    "type": ForwardRef("CTSSub"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSubSup",
                    "type": ForwardRef("CTSSubSup"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sSup",
                    "type": ForwardRef("CTSSup"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "r",
                    "type": ForwardRef("CTR"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtContentCell(Child):
    class Meta:
        name = "CT_SdtContentCell"

    content: list[
        Tc
        | CTCustomXmlCell
        | CTSdtCell
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tc",
                    "type": ForwardRef("Tc"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlCell"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("CTSdtCell"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtContentRow(Child):
    class Meta:
        name = "CT_SdtContentRow"

    content: list[
        Tr
        | CTCustomXmlRow
        | CTSdtRow
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tr",
                    "type": ForwardRef("Tr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRow"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("CTSdtRow"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CtBdoHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtBdoHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtBdoHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtCustomXmlRunHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtCustomXmlRunHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtCustomXmlRunHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtDirHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtDirHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtDirHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSdtContentRunHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtSdtContentRunHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtSdtContentRunHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSimpleFieldHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtSimpleFieldHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtSimpleFieldHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtSmartTagRunHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | CtSmartTagRunHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("CtSmartTagRunHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class PHyperlink(Child):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | PHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("PHyperlink"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_location: None | str = field(
        default=None,
        metadata={
            "name": "docLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "schema_default": "true",
        },
    )
    anchor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtCell(Child):
    class Meta:
        name = "CT_SdtCell"

    sdt_pr: None | SdtPr = field(
        default=None,
        metadata={
            "name": "sdtPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sdt_end_pr: None | CTSdtEndPr = field(
        default=None,
        metadata={
            "name": "sdtEndPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sdt_content: None | CTSdtContentCell = field(
        default=None,
        metadata={
            "name": "sdtContent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtRow(Child):
    class Meta:
        name = "CT_SdtRow"

    sdt_pr: None | SdtPr = field(
        default=None,
        metadata={
            "name": "sdtPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sdt_end_pr: None | CTSdtEndPr = field(
        default=None,
        metadata={
            "name": "sdtEndPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    sdt_content: None | CTSdtContentRow = field(
        default=None,
        metadata={
            "name": "sdtContent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class P(Child):
    class Meta:
        name = "p"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    p_pr: None | PPr = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
        },
    )
    content: list[
        CTCustomXmlRun
        | CTSmartTagRun
        | SdtRun
        | CtDir
        | CtBdo
        | R
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTSimpleField
        | PHyperlink
        | CTRel
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRun"),
                },
                {
                    "name": "smartTag",
                    "type": ForwardRef("CTSmartTagRun"),
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtRun"),
                },
                {
                    "name": "dir",
                    "type": ForwardRef("CtDir"),
                },
                {
                    "name": "bdo",
                    "type": ForwardRef("CtBdo"),
                },
                {
                    "name": "r",
                    "type": ForwardRef("R"),
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "fldSimple",
                    "type": ForwardRef("CTSimpleField"),
                },
                {
                    "name": "hyperlink",
                    "type": ForwardRef("PHyperlink"),
                },
                {
                    "name": "subDoc",
                    "type": ForwardRef("CTRel"),
                },
            ),
        },
    )
    rsid_rpr: None | str = field(
        default=None,
        metadata={
            "name": "rsidRPr",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_r: None | str = field(
        default=None,
        metadata={
            "name": "rsidR",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_del: None | str = field(
        default=None,
        metadata={
            "name": "rsidDel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_p: None | str = field(
        default=None,
        metadata={
            "name": "rsidP",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_rdefault: None | str = field(
        default=None,
        metadata={
            "name": "rsidRDefault",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    para_id: None | str = field(
        default=None,
        metadata={
            "name": "paraId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_id: None | str = field(
        default=None,
        metadata={
            "name": "textId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomXmlCell(Child):
    class Meta:
        name = "CT_CustomXmlCell"

    custom_xml_pr: None | CTCustomXmlPr = field(
        default=None,
        metadata={
            "name": "customXmlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        Tc
        | CTCustomXmlCell
        | CTSdtCell
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tc",
                    "type": ForwardRef("Tc"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlCell"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("CTSdtCell"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    element: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomXmlRow(Child):
    class Meta:
        name = "CT_CustomXmlRow"

    custom_xml_pr: None | CTCustomXmlPr = field(
        default=None,
        metadata={
            "name": "customXmlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        Tr
        | CTCustomXmlRow
        | CTSdtRow
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tr",
                    "type": ForwardRef("Tr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRow"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("CTSdtRow"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    element: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class SdtContentBlock(Child):
    class Meta:
        name = "CT_SdtContentBlock"

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class SdtBlock(Child):
    class Meta:
        name = "sdt"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    sdt_pr: None | SdtPr = field(
        default=None,
        metadata={
            "name": "sdtPr",
            "type": "Element",
        },
    )
    sdt_end_pr: None | CTSdtEndPr = field(
        default=None,
        metadata={
            "name": "sdtEndPr",
            "type": "Element",
        },
    )
    sdt_content: None | SdtContentBlock = field(
        default=None,
        metadata={
            "name": "sdtContent",
            "type": "Element",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomXmlBlock(Child):
    class Meta:
        name = "CT_CustomXmlBlock"

    custom_xml_pr: None | CTCustomXmlPr = field(
        default=None,
        metadata={
            "name": "customXmlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    element: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Tc(Child):
    class Meta:
        name = "CT_Tc"

    tc_pr: None | TcPr = field(
        default=None,
        metadata={
            "name": "tcPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Tr(Child):
    class Meta:
        name = "tr"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    tbl_pr_ex: None | CTTblPrEx = field(
        default=None,
        metadata={
            "name": "tblPrEx",
            "type": "Element",
        },
    )
    tr_pr: None | TrPr = field(
        default=None,
        metadata={
            "name": "trPr",
            "type": "Element",
        },
    )
    content: list[
        Tc
        | CTCustomXmlCell
        | CTSdtCell
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tc",
                    "type": ForwardRef("Tc"),
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlCell"),
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("CTSdtCell"),
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )
    rsid_rpr: None | str = field(
        default=None,
        metadata={
            "name": "rsidRPr",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_r: None | str = field(
        default=None,
        metadata={
            "name": "rsidR",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_del: None | str = field(
        default=None,
        metadata={
            "name": "rsidDel",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    rsid_tr: None | str = field(
        default=None,
        metadata={
            "name": "rsidTr",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    para_id: None | str = field(
        default=None,
        metadata={
            "name": "paraId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    text_id: None | str = field(
        default=None,
        metadata={
            "name": "textId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Tbl(Child):
    class Meta:
        name = "CT_Tbl"

    tbl_pr: None | TblPr = field(
        default=None,
        metadata={
            "name": "tblPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    tbl_grid: None | TblGrid = field(
        default=None,
        metadata={
            "name": "tblGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    content: list[
        Tr
        | CTCustomXmlRow
        | CTSdtRow
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tr",
                    "type": ForwardRef("Tr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlRow"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("CTSdtRow"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Body1(Child):
    class Meta:
        name = "CT_Body"

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    sect_pr: None | SectPr = field(
        default=None,
        metadata={
            "name": "sectPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFtnEdn(Child):
    class Meta:
        name = "CT_FtnEdn"

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    type_value: None | STFtnEdn = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTxbxContent(Child):
    class Meta:
        name = "CT_TxbxContent"

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Ftr(Child):
    class Meta:
        name = "ftr"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                },
            ),
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class Hdr(Child):
    class Meta:
        name = "hdr"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                },
            ),
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CommentsComment(CTTrackChange):
    class Meta:
        global_type = False

    content: list[
        CTCustomXmlBlock
        | SdtBlock
        | P
        | Tbl
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart2
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart2
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart2
        | CustomXmlInsRangeEnd2
        | CustomXmlDelRangeStart2
        | CustomXmlDelRangeEnd2
        | CustomXmlMoveFromRangeStart2
        | CustomXmlMoveFromRangeEnd2
        | CustomXmlMoveToRangeStart2
        | CustomXmlMoveToRangeEnd2
        | RunIns
        | RunDel
        | MoveFrom2
        | MoveTo2
        | OMathPara
        | OMath
        | CTAltChunk
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "customXml",
                    "type": ForwardRef("CTCustomXmlBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sdt",
                    "type": ForwardRef("SdtBlock"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "p",
                    "type": ForwardRef("P"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tbl",
                    "type": ForwardRef("Tbl"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "proofErr",
                    "type": ForwardRef("ProofErr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permStart",
                    "type": ForwardRef("RangePermissionStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "permEnd",
                    "type": ForwardRef("CTPerm"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkStart",
                    "type": ForwardRef("CTBookmark"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "bookmarkEnd",
                    "type": ForwardRef("CTMarkupRange"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeStart",
                    "type": ForwardRef("MoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFromRangeEnd",
                    "type": ForwardRef("CTMoveFromRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeEnd",
                    "type": ForwardRef("CTMoveToRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveToRangeStart",
                    "type": ForwardRef("MoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeStart",
                    "type": ForwardRef("CommentRangeStart"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentRangeEnd",
                    "type": ForwardRef("CommentRangeEnd"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeStart",
                    "type": ForwardRef("CustomXmlInsRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("RunIns"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("RunDel"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveFrom",
                    "type": ForwardRef("MoveFrom2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "oMathPara",
                    "type": ForwardRef("OMathPara"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "oMath",
                    "type": ForwardRef("OMath"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "altChunk",
                    "type": ForwardRef("CTAltChunk"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )
    initials: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocPart(Child):
    class Meta:
        name = "CT_DocPart"

    doc_part_pr: None | CTDocPartPr = field(
        default=None,
        metadata={
            "name": "docPartPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    doc_part_body: None | Body1 = field(
        default=None,
        metadata={
            "name": "docPartBody",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEndnotes(Child):
    class Meta:
        name = "CT_Endnotes"

    endnote: list[CTFtnEdn] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFootnotes(Child):
    class Meta:
        name = "CT_Footnotes"

    footnote: list[CTFtnEdn] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class Comments(Child):
    class Meta:
        name = "comments"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    comment: list[CommentsComment] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class Body(Body1):
    class Meta:
        name = "body"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class TxbxContent(CTTxbxContent):
    class Meta:
        name = "txbxContent"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTDocParts(Child):
    class Meta:
        name = "CT_DocParts"

    doc_part: list[CTDocPart] = field(
        default_factory=ChildList,
        metadata={
            "name": "docPart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Document(Child):
    class Meta:
        name = "document"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    background: None | CTBackground = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    body: None | Body = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class Endnotes(CTEndnotes):
    class Meta:
        name = "endnotes"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class Footnotes(CTFootnotes):
    class Meta:
        name = "footnotes"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(slots=True, kw_only=True)
class GlossaryDocument(Child):
    class Meta:
        name = "glossaryDocument"
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    background: None | CTBackground = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    doc_parts: None | CTDocParts = field(
        default=None,
        metadata={
            "name": "docParts",
            "type": "Element",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.customxml import SchemaLibrary
from docx4j_py.dml.wordprocessing_drawing import (
    Anchor,
    Inline,
)
from docx4j_py.math import (
    CTD,
    CTF,
    CTM,
    CTR,
    CTRPR,
    CTAcc,
    CTBar,
    CTBorderBox,
    CTBox,
    CTEqArr,
    CTFunc,
    CTGroupChr,
    CTLimLow,
    CTLimUpp,
    CTNary,
    CTPhant,
    CTRad,
    CTSPre,
    CTSSub,
    CTSSubSup,
    CTSSup,
    CTText,
    MathPr,
    OMath,
    OMathPara,
)
from docx4j_py.mce import AlternateContent
from docx4j_py.shared_types import StOnOff as SharedTypesStOnOff
from docx4j_py.w14 import (
    Checkbox,
    ConflictMode,
    CTFillTextEffect,
    CTGlow,
    CTLigatures,
    CTNumForm,
    CTNumSpacing,
    CTOnOff,
    CTProps3D,
    CTReflection,
    CTScene3D,
    CTStylisticSets,
    CTTextOutlineEffect,
    DefaultImageDpi,
    DiscardImageEditingData,
    Shadow,
)
from docx4j_py.w14 import (
    CTEmpty as W14Ctempty,
)
from docx4j_py.w14 import (
    DocId as W14DocId,
)
from docx4j_py.w15 import (
    Appearance,
    Collapsed,
    RepeatingSection,
    WebExtensionCreated,
    WebExtensionLinked,
)
from docx4j_py.w15 import (
    BooleanDefaultTrue as W15BooleanDefaultTrue,
)
from docx4j_py.w15 import (
    CTColor as W15Ctcolor,
)
from docx4j_py.w15 import (
    CTDataBinding as W15CtdataBinding,
)
from docx4j_py.w15 import (
    CTDecimalNumber as W15CtdecimalNumber,
)
from docx4j_py.w15 import (
    CTEmpty as W15Ctempty,
)
from docx4j_py.w15 import (
    DocId as W15DocId,
)
from docx4j_py.xml_ns import SpaceValue


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.wml import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "ANY_NS": "docx4j_py.wml.sdt",
    "BuilderError": "docx4j_py.wml.builders",
    "DEFAULT_DPI": "docx4j_py.wml.pictures",
    "EMU_PER_INCH": "docx4j_py.wml.pictures",
    "EmuSize": "docx4j_py.wml.pictures",
    "FragmentError": "docx4j_py.fragments",
    "HIGHLIGHT_COLORS": "docx4j_py.wml.builders",
    "ImageInfo": "docx4j_py.wml.pictures",
    "KIND_BY_QNAME": "docx4j_py.wml.sdt",
    "PICTURE_URI": "docx4j_py.wml.pictures",
    "RPR_BASE_FIELDS": "docx4j_py.wml.builders",
    "RPrElement": "docx4j_py.wml.builders",
    "RUN_OPTIONS": "docx4j_py.wml.builders",
    "SDT_FORMS": "docx4j_py.wml.sdt",
    "SDT_KINDS": "docx4j_py.wml.sdt",
    "SdtForm": "docx4j_py.wml.sdt",
    "SdtKind": "docx4j_py.wml.sdt",
    "UNDERLINE": "docx4j_py.wml.builders",
    "W14_NS": "docx4j_py.wml.sdt",
    "W15_NS": "docx4j_py.wml.sdt",
    "W_AND_W15": "docx4j_py.wml.sdt",
    "W_NS": "docx4j_py.wml.sdt",
    "apply_run_options": "docx4j_py.wml.builders",
    "br": "docx4j_py.wml.builders",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.wml.el",
    "element_name": "docx4j_py.traversal",
    "emu_for": "docx4j_py.wml.pictures",
    "find": "docx4j_py.traversal",
    "highlight_hex_value": "docx4j_py.wml.builders",
    "highlight_name_for_color": "docx4j_py.wml.builders",
    "image_size": "docx4j_py.wml.pictures",
    "inline_picture": "docx4j_py.wml.pictures",
    "iter_nodes": "docx4j_py.traversal",
    "link_parents": "docx4j_py.child",
    "next_sdt_id": "docx4j_py.wml.sdt",
    "p": "docx4j_py.wml.builders",
    "r": "docx4j_py.wml.builders",
    "read_run_options": "docx4j_py.wml.builders",
    "rpr_from_elements": "docx4j_py.wml.builders",
    "rpr_to_elements": "docx4j_py.wml.builders",
    "run_items_of": "docx4j_py.traversal",
    "sdt": "docx4j_py.wml.sdt",
    "sdt_kind_of": "docx4j_py.wml.sdt",
    "sdt_pr": "docx4j_py.wml.sdt",
    "sdt_property": "docx4j_py.wml.sdt",
    "t": "docx4j_py.wml.builders",
    "tab": "docx4j_py.wml.builders",
    "tbl": "docx4j_py.wml.builders",
    "tc": "docx4j_py.wml.builders",
    "text_of": "docx4j_py.traversal",
    "to_xml": "docx4j_py.fragments",
    "tr": "docx4j_py.wml.builders",
    "walk": "docx4j_py.traversal",
    "walk_all": "docx4j_py.traversal",
    "warm_up": "docx4j_py.runtime",
    "wml": "docx4j_py.fragments"
}


def __getattr__(name: str) -> object:
    """Import a Phase C helper, or the ``el`` submodule, on first use."""
    target = _PHASE_C.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    from importlib import import_module

    value = import_module(target) if name == "el" else getattr(import_module(target), name)
    globals()[name] = value
    return value
