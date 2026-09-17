"""python-docx's ``WD_*`` enumerations, with its member names and XML values.

CR-003 section 3.6, Phase J. python-docx's enums are ``int``-valued (the
number the MS Word API gives the member) and carry the ``w:val`` string
separately, so ``WD_ALIGN_PARAGRAPH.CENTER`` is ``1`` and its XML value is
``"center"``. Code in the wild compares members, prints them and passes them
straight into a setter, so both halves are kept: :class:`BaseXmlEnum` is
python-docx's own base class, re-implemented here rather than imported, since
python-docx is not a dependency of this package.

The content API speaks Office JS's strings (``"Centered"``, ``"Single"``), so
:data:`ALIGNMENT_TO_OFFICE_JS` and :data:`UNDERLINE_TO_OFFICE_JS` are where the
two vocabularies meet; everything else in this module is written straight into
the typed tree.
"""

from __future__ import annotations

import enum
from typing import Any, Self

__all__ = [
    "ALIGNMENT_TO_OFFICE_JS",
    "MSO_COLOR_TYPE",
    "MSO_THEME_COLOR",
    "MSO_THEME_COLOR_INDEX",
    "UNDERLINE_TO_OFFICE_JS",
    "WD_ALIGN_PARAGRAPH",
    "WD_ALIGN_VERTICAL",
    "WD_BREAK",
    "WD_BREAK_TYPE",
    "WD_CELL_VERTICAL_ALIGNMENT",
    "WD_COLOR",
    "WD_COLOR_INDEX",
    "WD_HEADER_FOOTER",
    "WD_HEADER_FOOTER_INDEX",
    "WD_INLINE_SHAPE",
    "WD_INLINE_SHAPE_TYPE",
    "WD_LINE_SPACING",
    "WD_ORIENT",
    "WD_ORIENTATION",
    "WD_PARAGRAPH_ALIGNMENT",
    "WD_ROW_HEIGHT",
    "WD_ROW_HEIGHT_RULE",
    "WD_SECTION",
    "WD_SECTION_START",
    "WD_STYLE_TYPE",
    "WD_TABLE_ALIGNMENT",
    "WD_TABLE_DIRECTION",
    "WD_TAB_ALIGNMENT",
    "WD_TAB_LEADER",
    "WD_UNDERLINE",
    "BaseEnum",
    "BaseXmlEnum",
]


class BaseEnum(int, enum.Enum):
    """An enum whose value is the MS Word API's integer, with no XML value."""

    def __new__(cls, ms_api_value: int, docstring: str) -> Self:
        """The member, carrying its documentation as ``__doc__``."""
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.__doc__ = docstring.strip()
        return self

    def __str__(self) -> str:
        """``'CENTER (1)'``, as python-docx prints it."""
        return f"{self.name} ({self.value})"


class BaseXmlEnum(int, enum.Enum):
    """An enum whose value is an integer and which also maps a ``w:val`` string."""

    xml_value: str | None

    def __new__(cls, ms_api_value: int, xml_value: str | None, docstring: str) -> Self:
        """The member, with its XML value and its documentation."""
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.xml_value = xml_value
        self.__doc__ = docstring.strip()
        return self

    def __str__(self) -> str:
        """``'CENTER (1)'``."""
        return f"{self.name} ({self.value})"

    @classmethod
    def from_xml(cls, xml_value: str | None) -> Any:
        """The member whose ``w:val`` is `xml_value`."""
        member = next((m for m in cls if m.xml_value == xml_value), None)
        if member is None:
            raise ValueError(f"{cls.__name__} has no XML mapping for {xml_value!r}")
        return member

    @classmethod
    def to_xml(cls, value: Any) -> str | None:
        """The ``w:val`` of a member, or of the integer that names one."""
        member = cls(value)
        xml_value = member.xml_value
        if not xml_value:
            raise ValueError(f"{cls.__name__}.{member.name} has no XML representation")
        return xml_value


# ---------------------------------------------------------------------------
# text
# ---------------------------------------------------------------------------


class WD_PARAGRAPH_ALIGNMENT(BaseXmlEnum):
    """``w:jc``: paragraph justification. MS API ``WdParagraphAlignment``."""

    LEFT = (0, "left", "Left-aligned.")
    CENTER = (1, "center", "Center-aligned.")
    RIGHT = (2, "right", "Right-aligned.")
    JUSTIFY = (3, "both", "Fully justified.")
    DISTRIBUTE = (4, "distribute", "Characters distributed to fill the width.")
    JUSTIFY_MED = (5, "mediumKashida", "Justified with a medium compression ratio.")
    JUSTIFY_HI = (7, "highKashida", "Justified with a high compression ratio.")
    JUSTIFY_LOW = (8, "lowKashida", "Justified with a low compression ratio.")
    THAI_JUSTIFY = (9, "thaiDistribute", "Justified for Thai layout.")


WD_ALIGN_PARAGRAPH = WD_PARAGRAPH_ALIGNMENT


class WD_BREAK_TYPE(enum.Enum):
    """The break kinds ``Run.add_break`` takes. MS API ``WdBreakType``.

    ``LINE_CLEAR_ALL`` and ``TEXT_WRAPPING`` are both 11, as they are in
    python-docx: the MS API has only ``TEXT_WRAPPING`` and python-docx added
    the other name for consistency with the two ``LINE_CLEAR_*`` members.
    """

    COLUMN = 8
    LINE = 6
    LINE_CLEAR_LEFT = 9
    LINE_CLEAR_RIGHT = 10
    LINE_CLEAR_ALL = 11
    PAGE = 7
    SECTION_CONTINUOUS = 3
    SECTION_EVEN_PAGE = 4
    SECTION_NEXT_PAGE = 2
    SECTION_ODD_PAGE = 5
    TEXT_WRAPPING = 11  # noqa: PIE796 - the same value as LINE_CLEAR_ALL, as in python-docx


WD_BREAK = WD_BREAK_TYPE


class WD_COLOR_INDEX(BaseXmlEnum):
    """``w:highlight``: Word's highlighter colours. MS API ``WdColorIndex``."""

    INHERITED = (-1, None, "Inherited from the style hierarchy.")
    AUTO = (0, "default", "Automatic colour, usually black.")
    BLACK = (1, "black", "Black.")
    BLUE = (2, "blue", "Blue.")
    BRIGHT_GREEN = (4, "green", "Bright green.")
    DARK_BLUE = (9, "darkBlue", "Dark blue.")
    DARK_RED = (13, "darkRed", "Dark red.")
    DARK_YELLOW = (14, "darkYellow", "Dark yellow.")
    GRAY_25 = (16, "lightGray", "25% grey.")
    GRAY_50 = (15, "darkGray", "50% grey.")
    GREEN = (11, "darkGreen", "Green.")
    PINK = (5, "magenta", "Pink.")
    RED = (6, "red", "Red.")
    TEAL = (10, "darkCyan", "Teal.")
    TURQUOISE = (3, "cyan", "Turquoise.")
    VIOLET = (12, "darkMagenta", "Violet.")
    WHITE = (8, "white", "White.")
    YELLOW = (7, "yellow", "Yellow.")


WD_COLOR = WD_COLOR_INDEX


class WD_LINE_SPACING(BaseXmlEnum):
    """``w:spacing/@lineRule``, plus the three multiples. MS API ``WdLineSpacing``."""

    SINGLE = (0, "UNMAPPED", "Single spaced; a multiple of 1.")
    ONE_POINT_FIVE = (1, "UNMAPPED", "Space-and-a-half; a multiple of 1.5.")
    DOUBLE = (2, "UNMAPPED", "Double spaced; a multiple of 2.")
    AT_LEAST = (3, "atLeast", "At least the stated amount.")
    EXACTLY = (4, "exact", "Exactly the stated amount.")
    MULTIPLE = (5, "auto", "A multiple of the line height.")


class WD_TAB_ALIGNMENT(BaseXmlEnum):
    """``w:tab/@val``. MS API ``WdTabAlignment``."""

    LEFT = (0, "left", "Left-aligned.")
    CENTER = (1, "center", "Center-aligned.")
    RIGHT = (2, "right", "Right-aligned.")
    DECIMAL = (3, "decimal", "Decimal-aligned.")
    BAR = (4, "bar", "Bar-aligned.")
    LIST = (6, "list", "List-aligned (deprecated).")
    CLEAR = (101, "clear", "Clears an inherited tab stop.")
    END = (102, "end", "Right-aligned (deprecated).")
    NUM = (103, "num", "Left-aligned (deprecated).")
    START = (104, "start", "Left-aligned (deprecated).")


class WD_TAB_LEADER(BaseXmlEnum):
    """``w:tab/@leader``. MS API ``WdTabLeader``."""

    SPACES = (0, "none", "Spaces; the default.")
    DOTS = (1, "dot", "Dots.")
    DASHES = (2, "hyphen", "Dashes.")
    LINES = (3, "underscore", "Double lines.")
    HEAVY = (4, "heavy", "A heavy line.")
    MIDDLE_DOT = (5, "middleDot", "A vertically centred dot.")


class WD_UNDERLINE(BaseXmlEnum):
    """``w:u/@val``. MS API ``WdUnderline``."""

    INHERITED = (-1, None, "Inherited from the containing paragraph.")
    NONE = (0, "none", "No underline, overriding any inherited value.")
    SINGLE = (1, "single", "A single line; read back as True.")
    WORDS = (2, "words", "Individual words only.")
    DOUBLE = (3, "double", "A double line.")
    DOTTED = (4, "dotted", "Dots.")
    THICK = (6, "thick", "A single thick line.")
    DASH = (7, "dash", "Dashes.")
    DOT_DASH = (9, "dotDash", "Alternating dots and dashes.")
    DOT_DOT_DASH = (10, "dotDotDash", "An alternating dot-dot-dash pattern.")
    WAVY = (11, "wave", "A single wavy line.")
    DOTTED_HEAVY = (20, "dottedHeavy", "Heavy dots.")
    DASH_HEAVY = (23, "dashedHeavy", "Heavy dashes.")
    DOT_DASH_HEAVY = (25, "dashDotHeavy", "Alternating heavy dots and dashes.")
    DOT_DOT_DASH_HEAVY = (26, "dashDotDotHeavy", "A heavy dot-dot-dash pattern.")
    WAVY_HEAVY = (27, "wavyHeavy", "A heavy wavy line.")
    DASH_LONG = (39, "dashLong", "Long dashes.")
    WAVY_DOUBLE = (43, "wavyDouble", "A double wavy line.")
    DASH_LONG_HEAVY = (55, "dashLongHeavy", "Long heavy dashes.")


# ---------------------------------------------------------------------------
# styles
# ---------------------------------------------------------------------------


class WD_STYLE_TYPE(BaseXmlEnum):
    """``w:style/@w:type``. MS API ``WdStyleType``."""

    CHARACTER = (2, "character", "A character style.")
    LIST = (4, "numbering", "A list style.")
    PARAGRAPH = (1, "paragraph", "A paragraph style.")
    TABLE = (3, "table", "A table style.")


# ---------------------------------------------------------------------------
# tables
# ---------------------------------------------------------------------------


class WD_CELL_VERTICAL_ALIGNMENT(BaseXmlEnum):
    """``w:vAlign`` in a ``w:tcPr``. MS API ``WdCellVerticalAlignment``."""

    TOP = (0, "top", "Aligned to the top border.")
    CENTER = (1, "center", "Centred vertically.")
    BOTTOM = (3, "bottom", "Aligned to the bottom border.")
    BOTH = (101, "both", "In the spec but not in Word; best avoided.")


WD_ALIGN_VERTICAL = WD_CELL_VERTICAL_ALIGNMENT


class WD_ROW_HEIGHT_RULE(BaseXmlEnum):
    """``w:trHeight/@hRule``. MS API ``WdRowHeightRule``."""

    AUTO = (0, "auto", "Tall enough for the tallest cell.")
    AT_LEAST = (1, "atLeast", "At least the stated height.")
    EXACTLY = (2, "exact", "Exactly the stated height.")


WD_ROW_HEIGHT = WD_ROW_HEIGHT_RULE


class WD_TABLE_ALIGNMENT(BaseXmlEnum):
    """``w:tblPr/w:jc``. MS API ``WdRowAlignment``."""

    LEFT = (0, "left", "Left-aligned.")
    CENTER = (1, "center", "Centred between the margins.")
    RIGHT = (2, "right", "Right-aligned.")


class WD_TABLE_DIRECTION(BaseEnum):
    """``w:bidiVisual``: which side the first column is on."""

    LTR = (0, "First column leftmost.")
    RTL = (1, "First column rightmost.")


# ---------------------------------------------------------------------------
# colour, from DrawingML
# ---------------------------------------------------------------------------


class MSO_COLOR_TYPE(BaseEnum):
    """How a colour is specified. MS API ``MsoColorType``."""

    RGB = (1, "An RGBColor value.")
    THEME = (2, "One of the preset theme colours.")
    AUTO = (101, "Determined by the application.")


class MSO_THEME_COLOR_INDEX(BaseXmlEnum):
    """``w:color/@themeColor``. MS API ``MsoThemeColorIndex``."""

    NOT_THEME_COLOR = (0, "UNMAPPED", "Not a theme colour.")
    ACCENT_1 = (5, "accent1", "Accent 1.")
    ACCENT_2 = (6, "accent2", "Accent 2.")
    ACCENT_3 = (7, "accent3", "Accent 3.")
    ACCENT_4 = (8, "accent4", "Accent 4.")
    ACCENT_5 = (9, "accent5", "Accent 5.")
    ACCENT_6 = (10, "accent6", "Accent 6.")
    BACKGROUND_1 = (14, "background1", "Background 1.")
    BACKGROUND_2 = (16, "background2", "Background 2.")
    DARK_1 = (1, "dark1", "Dark 1.")
    DARK_2 = (3, "dark2", "Dark 2.")
    FOLLOWED_HYPERLINK = (12, "followedHyperlink", "A clicked hyperlink.")
    HYPERLINK = (11, "hyperlink", "A hyperlink.")
    LIGHT_1 = (2, "light1", "Light 1.")
    LIGHT_2 = (4, "light2", "Light 2.")
    TEXT_1 = (13, "text1", "Text 1.")
    TEXT_2 = (15, "text2", "Text 2.")


MSO_THEME_COLOR = MSO_THEME_COLOR_INDEX


# ---------------------------------------------------------------------------
# sections and shapes
# ---------------------------------------------------------------------------


class WD_HEADER_FOOTER_INDEX(BaseXmlEnum):
    """``w:headerReference/@w:type``. MS API ``WdHeaderFooterIndex``."""

    PRIMARY = (1, "default", "Odd pages, or all of them.")
    FIRST_PAGE = (2, "first", "The first page of the section.")
    EVEN_PAGE = (3, "even", "Even pages of a recto/verso section.")


WD_HEADER_FOOTER = WD_HEADER_FOOTER_INDEX


class WD_ORIENTATION(BaseXmlEnum):
    """``w:pgSz/@orient``. MS API ``WdOrientation``."""

    PORTRAIT = (0, "portrait", "Portrait.")
    LANDSCAPE = (1, "landscape", "Landscape.")


WD_ORIENT = WD_ORIENTATION


class WD_SECTION_START(BaseXmlEnum):
    """``w:sectPr/w:type/@w:val``. MS API ``WdSectionStart``."""

    CONTINUOUS = (0, "continuous", "Continuous section break.")
    NEW_COLUMN = (1, "nextColumn", "New column section break.")
    NEW_PAGE = (2, "nextPage", "New page section break.")
    EVEN_PAGE = (3, "evenPage", "Even-page section break.")
    ODD_PAGE = (4, "oddPage", "Begins on the next odd page.")


WD_SECTION = WD_SECTION_START


class WD_INLINE_SHAPE_TYPE(enum.Enum):
    """What a ``wp:inline`` holds. MS API ``WdInlineShapeType``."""

    CHART = 12
    LINKED_PICTURE = 4
    PICTURE = 3
    SMART_ART = 15
    NOT_IMPLEMENTED = -6


WD_INLINE_SHAPE = WD_INLINE_SHAPE_TYPE


# ---------------------------------------------------------------------------
# where python-docx's members meet the content API's Office JS strings
# ---------------------------------------------------------------------------

#: ``WD_ALIGN_PARAGRAPH`` to the ``Word.Alignment`` value the content API takes.
ALIGNMENT_TO_OFFICE_JS: dict[WD_PARAGRAPH_ALIGNMENT, str] = {
    WD_PARAGRAPH_ALIGNMENT.LEFT: "Left",
    WD_PARAGRAPH_ALIGNMENT.CENTER: "Centered",
    WD_PARAGRAPH_ALIGNMENT.RIGHT: "Right",
    WD_PARAGRAPH_ALIGNMENT.JUSTIFY: "Justified",
}

#: ``WD_UNDERLINE`` to the ``Word.UnderlineType`` value the content API takes.
UNDERLINE_TO_OFFICE_JS: dict[WD_UNDERLINE, str] = {
    WD_UNDERLINE.NONE: "None",
    WD_UNDERLINE.SINGLE: "Single",
    WD_UNDERLINE.WORDS: "Word",
    WD_UNDERLINE.DOUBLE: "Double",
    WD_UNDERLINE.DOTTED: "Dotted",
    WD_UNDERLINE.THICK: "Thick",
    WD_UNDERLINE.DASH: "Dash",
    WD_UNDERLINE.DOT_DASH: "DotLine",
    WD_UNDERLINE.DOT_DOT_DASH: "DotDotLine",
    WD_UNDERLINE.WAVY: "Wave",
    WD_UNDERLINE.DOTTED_HEAVY: "DottedHeavy",
    WD_UNDERLINE.DASH_HEAVY: "DashLineHeavy",
    WD_UNDERLINE.DOT_DASH_HEAVY: "DotLineHeavy",
    WD_UNDERLINE.DOT_DOT_DASH_HEAVY: "DotDotLineHeavy",
    WD_UNDERLINE.WAVY_HEAVY: "WaveHeavy",
    WD_UNDERLINE.DASH_LONG: "DashLineLong",
    WD_UNDERLINE.WAVY_DOUBLE: "WaveDouble",
    WD_UNDERLINE.DASH_LONG_HEAVY: "DashLineLongHeavy",
}
