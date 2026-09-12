from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2014/chartex"


@dataclass(slots=True, kw_only=True)
class CTAddress(Child):
    class Meta:
        name = "CT_Address"

    address1: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    country_region: None | str = field(
        default=None,
        metadata={
            "name": "countryRegion",
            "type": "Attribute",
        },
    )
    admin_district1: None | str = field(
        default=None,
        metadata={
            "name": "adminDistrict1",
            "type": "Attribute",
        },
    )
    admin_district2: None | str = field(
        default=None,
        metadata={
            "name": "adminDistrict2",
            "type": "Attribute",
        },
    )
    postal_code: None | str = field(
        default=None,
        metadata={
            "name": "postalCode",
            "type": "Attribute",
        },
    )
    locality: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    iso_country_code: None | str = field(
        default=None,
        metadata={
            "name": "isoCountryCode",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAggregation(Child):
    class Meta:
        name = "CT_Aggregation"


@dataclass(slots=True, kw_only=True)
class CTAxisId(Child):
    class Meta:
        name = "CT_AxisId"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCopyrights(Child):
    class Meta:
        name = "CT_Copyrights"

    copyright: list[str] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataId(Child):
    class Meta:
        name = "CT_DataId"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataLabelHidden(Child):
    class Meta:
        name = "CT_DataLabelHidden"

    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataLabelVisibilities(Child):
    class Meta:
        name = "CT_DataLabelVisibilities"

    series_name: None | bool = field(
        default=None,
        metadata={
            "name": "seriesName",
            "type": "Attribute",
        },
    )
    category_name: None | bool = field(
        default=None,
        metadata={
            "name": "categoryName",
            "type": "Attribute",
        },
    )
    value: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExtension(Child):
    class Meta:
        name = "CT_Extension"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExternalData(Child):
    class Meta:
        name = "CT_ExternalData"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    auto_update: None | bool = field(
        default=None,
        metadata={
            "name": "autoUpdate",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExtremeValueColorPosition(Child):
    class Meta:
        name = "CT_ExtremeValueColorPosition"


@dataclass(slots=True, kw_only=True)
class CTGeoChildEntities(Child):
    class Meta:
        name = "CT_GeoChildEntities"

    geo_hierarchy_entity: list[str] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoHierarchyEntity",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoParentEntitiesQuery(Child):
    class Meta:
        name = "CT_GeoParentEntitiesQuery"

    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoParentEntity(Child):
    class Meta:
        name = "CT_GeoParentEntity"

    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoPolygon(Child):
    class Meta:
        name = "CT_GeoPolygon"

    polygon_id: None | str = field(
        default=None,
        metadata={
            "name": "polygonId",
            "type": "Attribute",
        },
    )
    num_points: None | int = field(
        default=None,
        metadata={
            "name": "numPoints",
            "type": "Attribute",
        },
    )
    pca_rings: None | str = field(
        default=None,
        metadata={
            "name": "pcaRings",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHeaderFooter(Child):
    class Meta:
        name = "CT_HeaderFooter"

    odd_header: None | str = field(
        default=None,
        metadata={
            "name": "oddHeader",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    odd_footer: None | str = field(
        default=None,
        metadata={
            "name": "oddFooter",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    even_header: None | str = field(
        default=None,
        metadata={
            "name": "evenHeader",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    even_footer: None | str = field(
        default=None,
        metadata={
            "name": "evenFooter",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    first_header: None | str = field(
        default=None,
        metadata={
            "name": "firstHeader",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    first_footer: None | str = field(
        default=None,
        metadata={
            "name": "firstFooter",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    align_with_margins: None | bool = field(
        default=None,
        metadata={
            "name": "alignWithMargins",
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    different_odd_even: None | bool = field(
        default=None,
        metadata={
            "name": "differentOddEven",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    different_first: None | bool = field(
        default=None,
        metadata={
            "name": "differentFirst",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumberColorPosition(Child):
    class Meta:
        name = "CT_NumberColorPosition"

    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumberFormat(Child):
    class Meta:
        name = "CT_NumberFormat"

    format_code: None | str = field(
        default=None,
        metadata={
            "name": "formatCode",
            "type": "Attribute",
        },
    )
    source_linked: None | bool = field(
        default=None,
        metadata={
            "name": "sourceLinked",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumericValue(Child):
    class Meta:
        name = "CT_NumericValue"

    value: None | float = field(default=None)
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPageMargins(Child):
    class Meta:
        name = "CT_PageMargins"

    l: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    r: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    t: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    b: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    header: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    footer: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPercentageColorPosition(Child):
    class Meta:
        name = "CT_PercentageColorPosition"

    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelId(Child):
    class Meta:
        name = "CT_RelId"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSeriesElementVisibilities(Child):
    class Meta:
        name = "CT_SeriesElementVisibilities"

    connector_lines: None | bool = field(
        default=None,
        metadata={
            "name": "connectorLines",
            "type": "Attribute",
        },
    )
    mean_line: None | bool = field(
        default=None,
        metadata={
            "name": "meanLine",
            "type": "Attribute",
        },
    )
    mean_marker: None | bool = field(
        default=None,
        metadata={
            "name": "meanMarker",
            "type": "Attribute",
        },
    )
    nonoutliers: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    outliers: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStringValue(Child):
    class Meta:
        name = "CT_StringValue"

    value: str = field(default="")
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSubtotalIndex(Child):
    class Meta:
        name = "CT_SubtotalIndex"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class STAxisUnit(Enum):
    HUNDREDS = "hundreds"
    THOUSANDS = "thousands"
    TEN_THOUSANDS = "tenThousands"
    HUNDRED_THOUSANDS = "hundredThousands"
    MILLIONS = "millions"
    TEN_MILLIONS = "tenMillions"
    HUNDRED_MILLIONS = "hundredMillions"
    BILLIONS = "billions"
    TRILLIONS = "trillions"
    PERCENTAGE = "percentage"


class STDataLabelPos(Enum):
    BEST_FIT = "bestFit"
    B = "b"
    CTR = "ctr"
    IN_BASE = "inBase"
    IN_END = "inEnd"
    L = "l"
    OUT_END = "outEnd"
    R = "r"
    T = "t"


class STEntityType(Enum):
    ADDRESS = "Address"
    ADMIN_DISTRICT = "AdminDistrict"
    ADMIN_DISTRICT2 = "AdminDistrict2"
    ADMIN_DISTRICT3 = "AdminDistrict3"
    CONTINENT = "Continent"
    COUNTRY_REGION = "CountryRegion"
    LOCALITY = "Locality"
    OCEAN = "Ocean"
    PLANET = "Planet"
    POSTAL_CODE = "PostalCode"
    REGION = "Region"
    UNSUPPORTED = "Unsupported"


class STFormulaDirection(Enum):
    COL = "col"
    ROW = "row"


class STGeoMappingLevel(Enum):
    DATA_ONLY = "dataOnly"
    POSTAL_CODE = "postalCode"
    COUNTY = "county"
    STATE = "state"
    COUNTRY_REGION = "countryRegion"
    COUNTRY_REGION_LIST = "countryRegionList"
    WORLD = "world"


class STGeoProjectionType(Enum):
    MERCATOR = "mercator"
    MILLER = "miller"
    ROBINSON = "robinson"
    ALBERS = "albers"


class STIntervalClosedSide(Enum):
    L = "l"
    R = "r"


class STNumericDimensionType(Enum):
    VAL = "val"
    X = "x"
    Y = "y"
    SIZE = "size"
    COLOR_VAL = "colorVal"


class STPageOrientation(Enum):
    DEFAULT = "default"
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"


class STParentLabelLayout(Enum):
    NONE = "none"
    BANNER = "banner"
    OVERLAPPING = "overlapping"


class STPosAlign(Enum):
    MIN = "min"
    CTR = "ctr"
    MAX = "max"


class STQuartileMethod(Enum):
    INCLUSIVE = "inclusive"
    EXCLUSIVE = "exclusive"


class STRegionLabelLayout(Enum):
    NONE = "none"
    BEST_FIT_ONLY = "bestFitOnly"
    SHOW_ALL = "showAll"


class STSeriesLayout(Enum):
    BOX_WHISKER = "boxWhisker"
    CLUSTERED_COLUMN = "clusteredColumn"
    FUNNEL = "funnel"
    PARETO_LINE = "paretoLine"
    REGION_MAP = "regionMap"
    SUNBURST = "sunburst"
    TREEMAP = "treemap"
    WATERFALL = "waterfall"


class STSidePos(Enum):
    L = "l"
    T = "t"
    R = "r"
    B = "b"


class STStringDimensionType(Enum):
    CAT = "cat"
    COLOR_STR = "colorStr"
    ENTITY_ID = "entityId"


class STTickMarksType(Enum):
    IN = "in"
    OUT = "out"
    CROSS = "cross"
    NONE = "none"


class StDoubleOrAutomaticValue(Enum):
    AUTO = "auto"


class StGapWidthRatioValue(Enum):
    AUTO = "auto"


class StValueAxisUnitValue(Enum):
    AUTO = "auto"


@dataclass(slots=True, kw_only=True)
class CTBinning(Child):
    class Meta:
        name = "CT_Binning"

    bin_size_or_bin_count: None | float | int = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "binSize",
                    "type": float,
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "binCount",
                    "type": int,
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    interval_closed: None | STIntervalClosedSide = field(
        default=None,
        metadata={
            "name": "intervalClosed",
            "type": "Attribute",
        },
    )
    underflow: None | float | StDoubleOrAutomaticValue = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    overflow: None | float | StDoubleOrAutomaticValue = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCategoryAxisScaling(Child):
    class Meta:
        name = "CT_CategoryAxisScaling"

    gap_width: None | float | StGapWidthRatioValue = field(
        default=None,
        metadata={
            "name": "gapWidth",
            "type": "Attribute",
            "min_inclusive": 0.0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExtensionList(Child):
    class Meta:
        name = "CT_ExtensionList"

    ext: list[CTExtension] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFormula(Child):
    class Meta:
        name = "CT_Formula"

    value: str = field(default="")
    dir: None | STFormulaDirection = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STFormulaDirection.COL,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoChildTypes(Child):
    class Meta:
        name = "CT_GeoChildTypes"

    entity_type: list[STEntityType] = field(
        default_factory=ChildList,
        metadata={
            "name": "entityType",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataEntityQuery(Child):
    class Meta:
        name = "CT_GeoDataEntityQuery"

    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )
    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataPointQuery(Child):
    class Meta:
        name = "CT_GeoDataPointQuery"

    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )
    latitude: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    longitude: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataPointToEntityQuery(Child):
    class Meta:
        name = "CT_GeoDataPointToEntityQuery"

    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )
    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoEntity(Child):
    class Meta:
        name = "CT_GeoEntity"

    entity_name: None | str = field(
        default=None,
        metadata={
            "name": "entityName",
            "type": "Attribute",
        },
    )
    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoHierarchyEntity(Child):
    class Meta:
        name = "CT_GeoHierarchyEntity"

    entity_name: None | str = field(
        default=None,
        metadata={
            "name": "entityName",
            "type": "Attribute",
        },
    )
    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )
    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoLocation(Child):
    class Meta:
        name = "CT_GeoLocation"

    address: None | CTAddress = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    latitude: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    longitude: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    entity_name: None | str = field(
        default=None,
        metadata={
            "name": "entityName",
            "type": "Attribute",
        },
    )
    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoLocationQuery(Child):
    class Meta:
        name = "CT_GeoLocationQuery"

    country_region: None | str = field(
        default=None,
        metadata={
            "name": "countryRegion",
            "type": "Attribute",
        },
    )
    admin_district1: None | str = field(
        default=None,
        metadata={
            "name": "adminDistrict1",
            "type": "Attribute",
        },
    )
    admin_district2: None | str = field(
        default=None,
        metadata={
            "name": "adminDistrict2",
            "type": "Attribute",
        },
    )
    postal_code: None | str = field(
        default=None,
        metadata={
            "name": "postalCode",
            "type": "Attribute",
        },
    )
    entity_type: None | STEntityType = field(
        default=None,
        metadata={
            "name": "entityType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoPolygons(Child):
    class Meta:
        name = "CT_GeoPolygons"

    geo_polygon: list[CTGeoPolygon] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoPolygon",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumericLevel(Child):
    class Meta:
        name = "CT_NumericLevel"

    pt: list[CTNumericValue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    pt_count: None | int = field(
        default=None,
        metadata={
            "name": "ptCount",
            "type": "Attribute",
        },
    )
    format_code: None | str = field(
        default=None,
        metadata={
            "name": "formatCode",
            "type": "Attribute",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPageSetup(Child):
    class Meta:
        name = "CT_PageSetup"

    paper_size: None | int = field(
        default=None,
        metadata={
            "name": "paperSize",
            "type": "Attribute",
            "schema_default": "1",
        },
    )
    first_page_number: None | int = field(
        default=None,
        metadata={
            "name": "firstPageNumber",
            "type": "Attribute",
            "schema_default": "1",
        },
    )
    orientation: None | STPageOrientation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STPageOrientation.DEFAULT,
        },
    )
    black_and_white: None | bool = field(
        default=None,
        metadata={
            "name": "blackAndWhite",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    draft: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    use_first_page_number: None | bool = field(
        default=None,
        metadata={
            "name": "useFirstPageNumber",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    horizontal_dpi: None | int = field(
        default=None,
        metadata={
            "name": "horizontalDpi",
            "type": "Attribute",
            "schema_default": "600",
        },
    )
    vertical_dpi: None | int = field(
        default=None,
        metadata={
            "name": "verticalDpi",
            "type": "Attribute",
            "schema_default": "600",
        },
    )
    copies: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "1",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTParentLabelLayout(Child):
    class Meta:
        name = "CT_ParentLabelLayout"

    val: None | STParentLabelLayout = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRegionLabelLayout(Child):
    class Meta:
        name = "CT_RegionLabelLayout"

    val: None | STRegionLabelLayout = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStatistics(Child):
    class Meta:
        name = "CT_Statistics"

    quartile_method: None | STQuartileMethod = field(
        default=None,
        metadata={
            "name": "quartileMethod",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStringLevel(Child):
    class Meta:
        name = "CT_StringLevel"

    pt: list[CTStringValue] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    pt_count: None | int = field(
        default=None,
        metadata={
            "name": "ptCount",
            "type": "Attribute",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSubtotals(Child):
    class Meta:
        name = "CT_Subtotals"

    idx: list[CTSubtotalIndex] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTValueAxisScaling(Child):
    class Meta:
        name = "CT_ValueAxisScaling"

    max: None | float | StDoubleOrAutomaticValue = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    min: None | float | StDoubleOrAutomaticValue = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    major_unit: None | float | StValueAxisUnitValue = field(
        default=None,
        metadata={
            "name": "majorUnit",
            "type": "Attribute",
            "min_exclusive": 0.0,
        },
    )
    minor_unit: None | float | StValueAxisUnitValue = field(
        default=None,
        metadata={
            "name": "minorUnit",
            "type": "Attribute",
            "min_exclusive": 0.0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTValueColorEndPosition(Child):
    class Meta:
        name = "CT_ValueColorEndPosition"

    extreme_value_or_number_or_percent: (
        None | CTExtremeValueColorPosition | CTNumberColorPosition | CTPercentageColorPosition
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "extremeValue",
                    "type": ForwardRef("CTExtremeValueColorPosition"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "number",
                    "type": ForwardRef("CTNumberColorPosition"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "percent",
                    "type": ForwardRef("CTPercentageColorPosition"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTValueColorMiddlePosition(Child):
    class Meta:
        name = "CT_ValueColorMiddlePosition"

    number_or_percent: None | CTNumberColorPosition | CTPercentageColorPosition = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "number",
                    "type": ForwardRef("CTNumberColorPosition"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "percent",
                    "type": ForwardRef("CTPercentageColorPosition"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTValueColors(Child):
    class Meta:
        name = "CT_ValueColors"

    min_color: None | CTSolidColorFillProperties = field(
        default=None,
        metadata={
            "name": "minColor",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    mid_color: None | CTSolidColorFillProperties = field(
        default=None,
        metadata={
            "name": "midColor",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    max_color: None | CTSolidColorFillProperties = field(
        default=None,
        metadata={
            "name": "maxColor",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class Chart(CTRelId):
    class Meta:
        name = "chart"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chartex"


@dataclass(slots=True, kw_only=True)
class CTDataLabel(Child):
    class Meta:
        name = "CT_DataLabel"

    num_fmt: None | CTNumberFormat = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    visibility: None | CTDataLabelVisibilities = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    separator: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    pos: None | STDataLabelPos = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataPoint(Child):
    class Meta:
        name = "CT_DataPoint"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFormatOverride(Child):
    class Meta:
        name = "CT_FormatOverride"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoChildEntitiesQuery(Child):
    class Meta:
        name = "CT_GeoChildEntitiesQuery"

    geo_child_types: None | CTGeoChildTypes = field(
        default=None,
        metadata={
            "name": "geoChildTypes",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoData(Child):
    class Meta:
        name = "CT_GeoData"

    geo_polygons: None | CTGeoPolygons = field(
        default=None,
        metadata={
            "name": "geoPolygons",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    copyrights: None | CTCopyrights = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    entity_name: None | str = field(
        default=None,
        metadata={
            "name": "entityName",
            "type": "Attribute",
        },
    )
    entity_id: None | str = field(
        default=None,
        metadata={
            "name": "entityId",
            "type": "Attribute",
        },
    )
    east: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    west: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    north: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    south: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataPointToEntityQueryResult(Child):
    class Meta:
        name = "CT_GeoDataPointToEntityQueryResult"

    geo_data_point_query: None | CTGeoDataPointQuery = field(
        default=None,
        metadata={
            "name": "geoDataPointQuery",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_data_point_to_entity_query: None | CTGeoDataPointToEntityQuery = field(
        default=None,
        metadata={
            "name": "geoDataPointToEntityQuery",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoLocations(Child):
    class Meta:
        name = "CT_GeoLocations"

    geo_location: None | CTGeoLocation = field(
        default=None,
        metadata={
            "name": "geoLocation",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoParentEntitiesQueryResult(Child):
    class Meta:
        name = "CT_GeoParentEntitiesQueryResult"

    geo_parent_entities_query: None | CTGeoParentEntitiesQuery = field(
        default=None,
        metadata={
            "name": "geoParentEntitiesQuery",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_entity: None | CTGeoEntity = field(
        default=None,
        metadata={
            "name": "geoEntity",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_parent_entity: None | CTGeoParentEntity = field(
        default=None,
        metadata={
            "name": "geoParentEntity",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGridlines(Child):
    class Meta:
        name = "CT_Gridlines"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLegend(Child):
    class Meta:
        name = "CT_Legend"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    pos: None | STSidePos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STSidePos.R,
        },
    )
    align: None | STPosAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STPosAlign.CTR,
        },
    )
    overlay: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPlotSurface(Child):
    class Meta:
        name = "CT_PlotSurface"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPrintSettings(Child):
    class Meta:
        name = "CT_PrintSettings"

    header_footer: None | CTHeaderFooter = field(
        default=None,
        metadata={
            "name": "headerFooter",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    page_margins: None | CTPageMargins = field(
        default=None,
        metadata={
            "name": "pageMargins",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    page_setup: None | CTPageSetup = field(
        default=None,
        metadata={
            "name": "pageSetup",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextData(Child):
    class Meta:
        name = "CT_TextData"

    f_or_v: list[CTFormula | str] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "f",
                    "type": ForwardRef("CTFormula"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "v",
                    "type": str,
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
            "max_occurs": 2,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTickLabels(Child):
    class Meta:
        name = "CT_TickLabels"

    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTickMarks(Child):
    class Meta:
        name = "CT_TickMarks"

    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    type_value: None | STTickMarksType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTValueColorPositions(Child):
    class Meta:
        name = "CT_ValueColorPositions"

    min: None | CTValueColorEndPosition = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    mid: None | CTValueColorMiddlePosition = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    max: None | CTValueColorEndPosition = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    count: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "2",
            "min_inclusive": 2,
            "max_inclusive": 3,
        },
    )


@dataclass(slots=True, kw_only=True)
class F(CTFormula):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Nf(CTFormula):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTDataLabels(Child):
    class Meta:
        name = "CT_DataLabels"

    num_fmt: None | CTNumberFormat = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    visibility: None | CTDataLabelVisibilities = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    separator: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    data_label: list[CTDataLabel] = field(
        default_factory=ChildList,
        metadata={
            "name": "dataLabel",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    data_label_hidden: list[CTDataLabelHidden] = field(
        default_factory=ChildList,
        metadata={
            "name": "dataLabelHidden",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    pos: None | STDataLabelPos = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFormatOverrides(Child):
    class Meta:
        name = "CT_FormatOverrides"

    fmt_ovr: list[CTFormatOverride] = field(
        default_factory=ChildList,
        metadata={
            "name": "fmtOvr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoChildEntitiesQueryResult(Child):
    class Meta:
        name = "CT_GeoChildEntitiesQueryResult"

    geo_child_entities_query: None | CTGeoChildEntitiesQuery = field(
        default=None,
        metadata={
            "name": "geoChildEntitiesQuery",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_child_entities: None | CTGeoChildEntities = field(
        default=None,
        metadata={
            "name": "geoChildEntities",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataEntityQueryResult(Child):
    class Meta:
        name = "CT_GeoDataEntityQueryResult"

    geo_data_entity_query: None | CTGeoDataEntityQuery = field(
        default=None,
        metadata={
            "name": "geoDataEntityQuery",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_data: None | CTGeoData = field(
        default=None,
        metadata={
            "name": "geoData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataPointToEntityQueryResults(Child):
    class Meta:
        name = "CT_GeoDataPointToEntityQueryResults"

    geo_data_point_to_entity_query_result: list[CTGeoDataPointToEntityQueryResult] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoDataPointToEntityQueryResult",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoLocationQueryResult(Child):
    class Meta:
        name = "CT_GeoLocationQueryResult"

    geo_location_query: None | CTGeoLocationQuery = field(
        default=None,
        metadata={
            "name": "geoLocationQuery",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_locations: None | CTGeoLocations = field(
        default=None,
        metadata={
            "name": "geoLocations",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoParentEntitiesQueryResults(Child):
    class Meta:
        name = "CT_GeoParentEntitiesQueryResults"

    geo_parent_entities_query_result: list[CTGeoParentEntitiesQueryResult] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoParentEntitiesQueryResult",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumericDimension(Child):
    class Meta:
        name = "CT_NumericDimension"

    f_or_nf_or_lvl: list[F | Nf | CTNumericLevel] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "f",
                    "type": ForwardRef("F"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "nf",
                    "type": ForwardRef("Nf"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "lvl",
                    "type": ForwardRef("CTNumericLevel"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    type_value: None | STNumericDimensionType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStringDimension(Child):
    class Meta:
        name = "CT_StringDimension"

    f_or_nf_or_lvl: list[F | Nf | CTStringLevel] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "f",
                    "type": ForwardRef("F"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "nf",
                    "type": ForwardRef("Nf"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "lvl",
                    "type": ForwardRef("CTStringLevel"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    type_value: None | STStringDimensionType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTText(Child):
    class Meta:
        name = "CT_Text"

    tx_data_or_rich: None | CTTextData | CTTextBody = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "txData",
                    "type": ForwardRef("CTTextData"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "rich",
                    "type": ForwardRef("CTTextBody"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAxisTitle(Child):
    class Meta:
        name = "CT_AxisTitle"

    tx: None | CTText = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAxisUnitsLabel(Child):
    class Meta:
        name = "CT_AxisUnitsLabel"

    tx: None | CTText = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartTitle(Child):
    class Meta:
        name = "CT_ChartTitle"

    tx: None | CTText = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    pos: None | STSidePos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STSidePos.T,
        },
    )
    align: None | STPosAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STPosAlign.CTR,
        },
    )
    overlay: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTData(Child):
    class Meta:
        name = "CT_Data"

    num_dim_or_str_dim: list[CTNumericDimension | CTStringDimension] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "numDim",
                    "type": ForwardRef("CTNumericDimension"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "strDim",
                    "type": ForwardRef("CTStringDimension"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoChildEntitiesQueryResults(Child):
    class Meta:
        name = "CT_GeoChildEntitiesQueryResults"

    geo_child_entities_query_result: list[CTGeoChildEntitiesQueryResult] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoChildEntitiesQueryResult",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoDataEntityQueryResults(Child):
    class Meta:
        name = "CT_GeoDataEntityQueryResults"

    geo_data_entity_query_result: list[CTGeoDataEntityQueryResult] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoDataEntityQueryResult",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoLocationQueryResults(Child):
    class Meta:
        name = "CT_GeoLocationQueryResults"

    geo_location_query_result: list[CTGeoLocationQueryResult] = field(
        default_factory=ChildList,
        metadata={
            "name": "geoLocationQueryResult",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAxisUnits(Child):
    class Meta:
        name = "CT_AxisUnits"

    units_label: None | CTAxisUnitsLabel = field(
        default=None,
        metadata={
            "name": "unitsLabel",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    unit: None | STAxisUnit = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartData(Child):
    class Meta:
        name = "CT_ChartData"

    external_data: None | CTExternalData = field(
        default=None,
        metadata={
            "name": "externalData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    data: list[CTData] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
            "min_occurs": 1,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTClear(Child):
    class Meta:
        name = "CT_Clear"

    geo_location_query_results: None | CTGeoLocationQueryResults = field(
        default=None,
        metadata={
            "name": "geoLocationQueryResults",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_data_entity_query_results: None | CTGeoDataEntityQueryResults = field(
        default=None,
        metadata={
            "name": "geoDataEntityQueryResults",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_data_point_to_entity_query_results: None | CTGeoDataPointToEntityQueryResults = field(
        default=None,
        metadata={
            "name": "geoDataPointToEntityQueryResults",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_child_entities_query_results: None | CTGeoChildEntitiesQueryResults = field(
        default=None,
        metadata={
            "name": "geoChildEntitiesQueryResults",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    geo_parent_entities_query_results: None | CTGeoParentEntitiesQueryResults = field(
        default=None,
        metadata={
            "name": "geoParentEntitiesQueryResults",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAxis(Child):
    class Meta:
        name = "CT_Axis"

    cat_scaling_or_val_scaling: None | CTCategoryAxisScaling | CTValueAxisScaling = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "catScaling",
                    "type": ForwardRef("CTCategoryAxisScaling"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "valScaling",
                    "type": ForwardRef("CTValueAxisScaling"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    title: None | CTAxisTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    units: None | CTAxisUnits = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    major_gridlines: None | CTGridlines = field(
        default=None,
        metadata={
            "name": "majorGridlines",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    minor_gridlines: None | CTGridlines = field(
        default=None,
        metadata={
            "name": "minorGridlines",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    major_tick_marks: None | CTTickMarks = field(
        default=None,
        metadata={
            "name": "majorTickMarks",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    minor_tick_marks: None | CTTickMarks = field(
        default=None,
        metadata={
            "name": "minorTickMarks",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tick_labels: None | CTTickLabels = field(
        default=None,
        metadata={
            "name": "tickLabels",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    num_fmt: None | CTNumberFormat = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    hidden: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeoCache(Child):
    class Meta:
        name = "CT_GeoCache"

    binary_or_clear: list[bytes | CTClear] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "binary",
                    "type": bytes,
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                    "format": "base64",
                },
                {
                    "name": "clear",
                    "type": ForwardRef("CTClear"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    provider: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeography(Child):
    class Meta:
        name = "CT_Geography"

    geo_cache: None | CTGeoCache = field(
        default=None,
        metadata={
            "name": "geoCache",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    projection_type: None | STGeoProjectionType = field(
        default=None,
        metadata={
            "name": "projectionType",
            "type": "Attribute",
        },
    )
    viewed_region_type: None | STGeoMappingLevel = field(
        default=None,
        metadata={
            "name": "viewedRegionType",
            "type": "Attribute",
        },
    )
    culture_language: None | str = field(
        default=None,
        metadata={
            "name": "cultureLanguage",
            "type": "Attribute",
        },
    )
    culture_region: None | str = field(
        default=None,
        metadata={
            "name": "cultureRegion",
            "type": "Attribute",
        },
    )
    attribution: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSeriesLayoutProperties(Child):
    class Meta:
        name = "CT_SeriesLayoutProperties"

    parent_label_layout: None | CTParentLabelLayout = field(
        default=None,
        metadata={
            "name": "parentLabelLayout",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    region_label_layout: None | CTRegionLabelLayout = field(
        default=None,
        metadata={
            "name": "regionLabelLayout",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    visibility: None | CTSeriesElementVisibilities = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    aggregation_or_binning: None | CTAggregation | CTBinning = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "aggregation",
                    "type": ForwardRef("CTAggregation"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
                {
                    "name": "binning",
                    "type": ForwardRef("CTBinning"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                },
            ),
        },
    )
    geography: None | CTGeography = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    statistics: None | CTStatistics = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    subtotals: None | CTSubtotals = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSeries(Child):
    class Meta:
        name = "CT_Series"

    tx: None | CTText = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    value_colors: None | CTValueColors = field(
        default=None,
        metadata={
            "name": "valueColors",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    value_color_positions: None | CTValueColorPositions = field(
        default=None,
        metadata={
            "name": "valueColorPositions",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    data_pt: list[CTDataPoint] = field(
        default_factory=ChildList,
        metadata={
            "name": "dataPt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    data_labels: None | CTDataLabels = field(
        default=None,
        metadata={
            "name": "dataLabels",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    data_id: None | CTDataId = field(
        default=None,
        metadata={
            "name": "dataId",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    layout_pr: None | CTSeriesLayoutProperties = field(
        default=None,
        metadata={
            "name": "layoutPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    axis_id: list[CTAxisId] = field(
        default_factory=ChildList,
        metadata={
            "name": "axisId",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    layout_id: None | STSeriesLayout = field(
        default=None,
        metadata={
            "name": "layoutId",
            "type": "Attribute",
        },
    )
    hidden: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    owner_idx: None | int = field(
        default=None,
        metadata={
            "name": "ownerIdx",
            "type": "Attribute",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
        },
    )
    format_idx: None | int = field(
        default=None,
        metadata={
            "name": "formatIdx",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPlotAreaRegion(Child):
    class Meta:
        name = "CT_PlotAreaRegion"

    plot_surface: None | CTPlotSurface = field(
        default=None,
        metadata={
            "name": "plotSurface",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    series: list[CTSeries] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPlotArea(Child):
    class Meta:
        name = "CT_PlotArea"

    plot_area_region: None | CTPlotAreaRegion = field(
        default=None,
        metadata={
            "name": "plotAreaRegion",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    axis: list[CTAxis] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChart(Child):
    class Meta:
        name = "CT_Chart"

    title: None | CTChartTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    plot_area: None | CTPlotArea = field(
        default=None,
        metadata={
            "name": "plotArea",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    legend: None | CTLegend = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartSpace(Child):
    class Meta:
        name = "CT_ChartSpace"

    chart_data: None | CTChartData = field(
        default=None,
        metadata={
            "name": "chartData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    chart: None | CTChart = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    clr_map_ovr: None | CTColorMapping = field(
        default=None,
        metadata={
            "name": "clrMapOvr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    fmt_ovrs: None | CTFormatOverrides = field(
        default=None,
        metadata={
            "name": "fmtOvrs",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    print_settings: None | CTPrintSettings = field(
        default=None,
        metadata={
            "name": "printSettings",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        },
    )


@dataclass(slots=True, kw_only=True)
class ChartSpace(CTChartSpace):
    class Meta:
        name = "chartSpace"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chartex"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTColorMapping,
    CTShapeProperties,
    CTSolidColorFillProperties,
    CTTextBody,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.chartex_2014 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.chartex_2014.el",
    "element_name": "docx4j_py.traversal",
    "find": "docx4j_py.traversal",
    "iter_nodes": "docx4j_py.traversal",
    "link_parents": "docx4j_py.child",
    "text_of": "docx4j_py.traversal",
    "to_xml": "docx4j_py.fragments",
    "walk": "docx4j_py.traversal",
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
