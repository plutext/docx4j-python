"""The DrawingML parts: theme, charts and diagrams.

CR-002 section 5.4: "chart/diagram parts typed where the model has classes".
Six of them do (chart, chart user shapes, the four diagram parts) plus the theme
and the theme override; the three Microsoft extension parts --- chart style,
chart colour style and ``chartex`` --- have no classes in this build's schema
closure, so they are :class:`DefaultXmlPart`\\ s and lose nothing.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.default_xml_part import DefaultXmlPart
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.xml_part import XmlPart

__all__ = [
    "ChartColorStylePart",
    "ChartExSpacePart",
    "ChartPart",
    "ChartStylePart",
    "DiagramColorsPart",
    "DiagramDataPart",
    "DiagramDrawingPart",
    "DiagramLayoutHeaderPart",
    "DiagramLayoutPart",
    "DiagramStylePart",
    "DrawingPart",
    "ThemeOverridePart",
    "ThemePart",
]

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


class ThemePart(XmlPart[Any]):
    """``/word/theme/theme1.xml``, ``a:theme``. docx4j ``ThemePart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.main.Theme"

    def __init__(self, part_name: PartName | str = "/word/theme/theme1.xml") -> None:
        """Build the theme part."""
        super().__init__(
            part_name,
            ContentTypes.OFFICEDOCUMENT_THEME,
            Namespaces.THEME,
            f"{{{_A}}}theme",
        )


class ThemeOverridePart(XmlPart[Any]):
    """``a:themeOverride``. docx4j ``ThemeOverridePart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.main.ThemeOverride"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the theme override part."""
        super().__init__(
            part_name,
            ContentTypes.OFFICEDOCUMENT_THEME_OVERRIDE,
            Namespaces.THEME_OVERRIDE,
            f"{{{_A}}}themeOverride",
        )


class ChartPart(XmlPart[Any]):
    """``c:chartSpace``. docx4j ``DrawingML.Chart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.chart.ChartSpace"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the chart part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_CHART,
            Namespaces.SPREADSHEETML_CHART,
            f"{{{_C}}}chartSpace",
        )


class DrawingPart(XmlPart[Any]):
    """``xdr:wsDr``, a spreadsheet drawing. docx4j ``DrawingML.Drawing``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.spreadsheet_drawing.WsDr"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the drawing part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DRAWING,
            Namespaces.SPREADSHEETML_DRAWING,
            f"{{{_XDR}}}wsDr",
        )


class DiagramDataPart(XmlPart[Any]):
    """``dgm:dataModel``. docx4j ``DiagramDataPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.diagram.DataModel"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the diagram data part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DIAGRAM_DATA,
            Namespaces.DRAWINGML_DIAGRAM_DATA,
            f"{{{_DGM}}}dataModel",
        )


class DiagramLayoutPart(XmlPart[Any]):
    """``dgm:layoutDef``. docx4j ``DiagramLayoutPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.diagram.LayoutDef"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the diagram layout part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DIAGRAM_LAYOUT,
            Namespaces.DRAWINGML_DIAGRAM_LAYOUT,
            f"{{{_DGM}}}layoutDef",
        )


class DiagramLayoutHeaderPart(XmlPart[Any]):
    """``dgm:layoutDefHdr``. docx4j ``DiagramLayoutHeaderPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.diagram.LayoutDefHdr"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the diagram layout header part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DIAGRAM_LAYOUT_HEADER,
            Namespaces.DRAWINGML_DIAGRAM_LAYOUT_HEADER,
            f"{{{_DGM}}}layoutDefHdr",
        )


class DiagramStylePart(XmlPart[Any]):
    """``dgm:styleDef``. docx4j ``DiagramStylePart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.diagram.StyleDef"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the diagram style part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DIAGRAM_STYLE,
            Namespaces.DRAWINGML_DIAGRAM_STYLE,
            f"{{{_DGM}}}styleDef",
        )


class DiagramColorsPart(XmlPart[Any]):
    """``dgm:colorsDef``. docx4j ``DiagramColorsPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.dml.diagram.ColorsDef"

    def __init__(self, part_name: PartName | str) -> None:
        """Build the diagram colours part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DIAGRAM_COLORS,
            Namespaces.DRAWINGML_DIAGRAM_COLORS,
            f"{{{_DGM}}}colorsDef",
        )


class ChartShapePart(DefaultXmlPart):
    """A chart's user shapes, ``cdr:userShapes``. docx4j ``ChartShapePart``.

    A tree, not a typed part: this build's ``dml.chart_drawing`` has no global
    element for ``cdr:userShapes`` (the one the generator produced sits in the
    chart namespace), so typing it would risk writing the wrong root name.
    """

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the chart shapes part."""
        super().__init__(
            part_name, ContentTypes.DRAWINGML_CHART_SHAPES, Namespaces.CHART_USER_SHAPES
        )


class ChartStylePart(DefaultXmlPart):
    """``cs:chartStyle``, a Microsoft extension. A tree until the model types it."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the chart style part."""
        super().__init__(part_name, ContentTypes.CHART_STYLE, Namespaces.CHART_STYLE)


class ChartColorStylePart(DefaultXmlPart):
    """``cs:colorStyle``, a Microsoft extension. A tree until the model types it."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the chart colour style part."""
        super().__init__(part_name, ContentTypes.CHART_COLOR_STYLE, Namespaces.CHART_COLOR_STYLE)


class ChartExSpacePart(DefaultXmlPart):
    """``cx:chartSpace``, a Microsoft extension. A tree until the model types it."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the chartex part."""
        super().__init__(part_name, ContentTypes.CHART_EX, Namespaces.CHART_EX)


class DiagramDrawingPart(DefaultXmlPart):
    """``dsp:drawing``, the Microsoft diagram drawing. A tree until the model types it."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the diagram drawing part."""
        super().__init__(
            part_name,
            ContentTypes.DRAWINGML_DIAGRAM_DRAWING,
            Namespaces.DRAWINGML_DIAGRAM_DRAWING,
        )
