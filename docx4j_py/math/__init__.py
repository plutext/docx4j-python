from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@dataclass(slots=True, kw_only=True)
class CTChar(Child):
    class Meta:
        name = "CT_Char"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "max_length": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInteger2(Child):
    class Meta:
        name = "CT_Integer2"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_inclusive": -2,
            "max_inclusive": 2,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInteger255(Child):
    class Meta:
        name = "CT_Integer255"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_inclusive": 1,
            "max_inclusive": 255,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMR(Child):
    class Meta:
        name = "CT_MR"

    e: list[CTOMathArg] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTManualBreak(Child):
    class Meta:
        name = "CT_ManualBreak"

    aln_at: None | int = field(
        default=None,
        metadata={
            "name": "alnAt",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_inclusive": 1,
            "max_inclusive": 255,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSpacingRule(Child):
    class Meta:
        name = "CT_SpacingRule"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_inclusive": 0,
            "max_inclusive": 4,
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
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTwipsMeasure(Child):
    class Meta:
        name = "CT_TwipsMeasure"

    val: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "pattern": r"[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUnSignedInteger(Child):
    class Meta:
        name = "CT_UnSignedInteger"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


class STBreakBin(Enum):
    BEFORE = "before"
    AFTER = "after"
    REPEAT = "repeat"


class STFType(Enum):
    BAR = "bar"
    SKW = "skw"
    LIN = "lin"
    NO_BAR = "noBar"


class STJc(Enum):
    LEFT = "left"
    RIGHT = "right"
    CENTER = "center"
    CENTER_GROUP = "centerGroup"


class STLimLoc(Enum):
    UND_OVR = "undOvr"
    SUB_SUP = "subSup"


class STScript(Enum):
    ROMAN = "roman"
    SCRIPT = "script"
    FRAKTUR = "fraktur"
    DOUBLE_STRUCK = "double-struck"
    SANS_SERIF = "sans-serif"
    MONOSPACE = "monospace"


class STShp(Enum):
    CENTERED = "centered"
    MATCH = "match"


class STStyle(Enum):
    P = "p"
    B = "b"
    I = "i"
    BI = "bi"


class STTopBot(Enum):
    TOP = "top"
    BOT = "bot"


class StBreakBinSub(Enum):
    HYPHEN_MINUS_HYPHEN_MINUS = "--"
    HYPHEN_MINUS_PLUS_SIGN = "-+"
    PLUS_SIGN_HYPHEN_MINUS = "+-"


@dataclass(slots=True, kw_only=True)
class CTBreakBin(Child):
    class Meta:
        name = "CT_BreakBin"

    val: None | STBreakBin = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBreakBinSub(Child):
    class Meta:
        name = "CT_BreakBinSub"

    val: None | StBreakBinSub = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCtrlPr(Child):
    class Meta:
        name = "CT_CtrlPr"

    r_pr_or_ins_or_del: None | RPr | Ins2 | Del2 = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "rPr",
                    "type": ForwardRef("RPr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ins",
                    "type": ForwardRef("Ins2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("Del2"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFType(Child):
    class Meta:
        name = "CT_FType"

    val: None | STFType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLimLoc(Child):
    class Meta:
        name = "CT_LimLoc"

    val: None | STLimLoc = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOMathArgPr(Child):
    class Meta:
        name = "CT_OMathArgPr"

    arg_sz: None | CTInteger2 = field(
        default=None,
        metadata={
            "name": "argSz",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOMathJc(Child):
    class Meta:
        name = "CT_OMathJc"

    val: None | STJc = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOnOff(Child):
    class Meta:
        name = "CT_OnOff"

    val: None | StOnOff = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScript(Child):
    class Meta:
        name = "CT_Script"

    val: None | STScript = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShp(Child):
    class Meta:
        name = "CT_Shp"

    val: None | STShp = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyle(Child):
    class Meta:
        name = "CT_Style"

    val: None | STStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTText(Child):
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
class CTTopBot(Child):
    class Meta:
        name = "CT_TopBot"

    val: None | STTopBot = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTXAlign(Child):
    class Meta:
        name = "CT_XAlign"

    val: None | StXalign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTYAlign(Child):
    class Meta:
        name = "CT_YAlign"

    val: None | StYalign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAccPr(Child):
    class Meta:
        name = "CT_AccPr"

    chr: None | CTChar = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBarPr(Child):
    class Meta:
        name = "CT_BarPr"

    pos: None | CTTopBot = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBorderBoxPr(Child):
    class Meta:
        name = "CT_BorderBoxPr"

    hide_top: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "hideTop",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    hide_bot: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "hideBot",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    hide_left: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "hideLeft",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    hide_right: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "hideRight",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    strike_h: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "strikeH",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    strike_v: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "strikeV",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    strike_bltr: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "strikeBLTR",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    strike_tlbr: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "strikeTLBR",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBoxPr(Child):
    class Meta:
        name = "CT_BoxPr"

    op_emu: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "opEmu",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    no_break: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "noBreak",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    diff: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    brk: None | CTManualBreak = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    aln: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDPr(Child):
    class Meta:
        name = "CT_DPr"

    beg_chr: None | CTChar = field(
        default=None,
        metadata={
            "name": "begChr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sep_chr: None | CTChar = field(
        default=None,
        metadata={
            "name": "sepChr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    end_chr: None | CTChar = field(
        default=None,
        metadata={
            "name": "endChr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    grow: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    shp: None | CTShp = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEqArrPr(Child):
    class Meta:
        name = "CT_EqArrPr"

    base_jc: None | CTYAlign = field(
        default=None,
        metadata={
            "name": "baseJc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    max_dist: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "maxDist",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    obj_dist: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "objDist",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    r_sp_rule: None | CTSpacingRule = field(
        default=None,
        metadata={
            "name": "rSpRule",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    r_sp: None | CTUnSignedInteger = field(
        default=None,
        metadata={
            "name": "rSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFPr(Child):
    class Meta:
        name = "CT_FPr"

    type_value: None | CTFType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFuncPr(Child):
    class Meta:
        name = "CT_FuncPr"

    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupChrPr(Child):
    class Meta:
        name = "CT_GroupChrPr"

    chr: None | CTChar = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    pos: None | CTTopBot = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    vert_jc: None | CTTopBot = field(
        default=None,
        metadata={
            "name": "vertJc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLimLowPr(Child):
    class Meta:
        name = "CT_LimLowPr"

    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLimUppPr(Child):
    class Meta:
        name = "CT_LimUppPr"

    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMCPr(Child):
    class Meta:
        name = "CT_MCPr"

    count: None | CTInteger255 = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    mc_jc: None | CTXAlign = field(
        default=None,
        metadata={
            "name": "mcJc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMathPr(Child):
    class Meta:
        name = "CT_MathPr"

    math_font: None | CTString = field(
        default=None,
        metadata={
            "name": "mathFont",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    brk_bin: None | CTBreakBin = field(
        default=None,
        metadata={
            "name": "brkBin",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    brk_bin_sub: None | CTBreakBinSub = field(
        default=None,
        metadata={
            "name": "brkBinSub",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    small_frac: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "smallFrac",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    disp_def: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "dispDef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    l_margin: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "lMargin",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    r_margin: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "rMargin",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    def_jc: None | CTOMathJc = field(
        default=None,
        metadata={
            "name": "defJc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    pre_sp: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "preSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    post_sp: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "postSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    inter_sp: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "interSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    intra_sp: None | CTTwipsMeasure = field(
        default=None,
        metadata={
            "name": "intraSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    wrap_indent_or_wrap_right: None | CTTwipsMeasure | CTOnOff = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "wrapIndent",
                    "type": ForwardRef("CTTwipsMeasure"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "wrapRight",
                    "type": ForwardRef("CTOnOff"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
        },
    )
    int_lim: None | CTLimLoc = field(
        default=None,
        metadata={
            "name": "intLim",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    nary_lim: None | CTLimLoc = field(
        default=None,
        metadata={
            "name": "naryLim",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNaryPr(Child):
    class Meta:
        name = "CT_NaryPr"

    chr: None | CTChar = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    lim_loc: None | CTLimLoc = field(
        default=None,
        metadata={
            "name": "limLoc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    grow: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sub_hide: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "subHide",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sup_hide: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "supHide",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOMathParaPr(Child):
    class Meta:
        name = "CT_OMathParaPr"

    jc: None | CTOMathJc = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPhantPr(Child):
    class Meta:
        name = "CT_PhantPr"

    show: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    zero_wid: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "zeroWid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    zero_asc: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "zeroAsc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    zero_desc: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "zeroDesc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    transp: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRPR(Child):
    class Meta:
        name = "CT_RPR"

    lit: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    nor_or_scr_or_sty: list[CTOnOff | CTScript | CTStyle] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "nor",
                    "type": ForwardRef("CTOnOff"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "scr",
                    "type": ForwardRef("CTScript"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
                {
                    "name": "sty",
                    "type": ForwardRef("CTStyle"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
                },
            ),
            "max_occurs": 2,
        },
    )
    brk: None | CTManualBreak = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    aln: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRadPr(Child):
    class Meta:
        name = "CT_RadPr"

    deg_hide: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "degHide",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSPrePr(Child):
    class Meta:
        name = "CT_SPrePr"

    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSSubPr(Child):
    class Meta:
        name = "CT_SSubPr"

    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSSubSupPr(Child):
    class Meta:
        name = "CT_SSubSupPr"

    aln_scr: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "alnScr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSSupPr(Child):
    class Meta:
        name = "CT_SSupPr"

    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAcc(Child):
    class Meta:
        name = "CT_Acc"

    acc_pr: None | CTAccPr = field(
        default=None,
        metadata={
            "name": "accPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBar(Child):
    class Meta:
        name = "CT_Bar"

    bar_pr: None | CTBarPr = field(
        default=None,
        metadata={
            "name": "barPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBorderBox(Child):
    class Meta:
        name = "CT_BorderBox"

    border_box_pr: None | CTBorderBoxPr = field(
        default=None,
        metadata={
            "name": "borderBoxPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBox(Child):
    class Meta:
        name = "CT_Box"

    box_pr: None | CTBoxPr = field(
        default=None,
        metadata={
            "name": "boxPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTD(Child):
    class Meta:
        name = "CT_D"

    d_pr: None | CTDPr = field(
        default=None,
        metadata={
            "name": "dPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: list[CTOMathArg] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEqArr(Child):
    class Meta:
        name = "CT_EqArr"

    eq_arr_pr: None | CTEqArrPr = field(
        default=None,
        metadata={
            "name": "eqArrPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: list[CTOMathArg] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTF(Child):
    class Meta:
        name = "CT_F"

    f_pr: None | CTFPr = field(
        default=None,
        metadata={
            "name": "fPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    num: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    den: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFunc(Child):
    class Meta:
        name = "CT_Func"

    func_pr: None | CTFuncPr = field(
        default=None,
        metadata={
            "name": "funcPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    f_name: None | CTOMathArg = field(
        default=None,
        metadata={
            "name": "fName",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupChr(Child):
    class Meta:
        name = "CT_GroupChr"

    group_chr_pr: None | CTGroupChrPr = field(
        default=None,
        metadata={
            "name": "groupChrPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLimLow(Child):
    class Meta:
        name = "CT_LimLow"

    lim_low_pr: None | CTLimLowPr = field(
        default=None,
        metadata={
            "name": "limLowPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    lim: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLimUpp(Child):
    class Meta:
        name = "CT_LimUpp"

    lim_upp_pr: None | CTLimUppPr = field(
        default=None,
        metadata={
            "name": "limUppPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    lim: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMC(Child):
    class Meta:
        name = "CT_MC"

    mc_pr: None | CTMCPr = field(
        default=None,
        metadata={
            "name": "mcPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNary(Child):
    class Meta:
        name = "CT_Nary"

    nary_pr: None | CTNaryPr = field(
        default=None,
        metadata={
            "name": "naryPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sub: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sup: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOMathPara(Child):
    class Meta:
        name = "CT_OMathPara"

    o_math_para_pr: None | CTOMathParaPr = field(
        default=None,
        metadata={
            "name": "oMathParaPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    o_math: list[CTOMath] = field(
        default_factory=ChildList,
        metadata={
            "name": "oMath",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPhant(Child):
    class Meta:
        name = "CT_Phant"

    phant_pr: None | CTPhantPr = field(
        default=None,
        metadata={
            "name": "phantPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTR(Child):
    class Meta:
        name = "CT_R"

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
        | T
        | DelText
        | InstrText
        | DelInstrText
        | CtRNoBreakHyphen
        | CtRSoftHyphen
        | CtRDayShort
        | CtRMonthShort
        | CtRYearShort
        | CtRDayLong
        | CtRMonthLong
        | CtRYearLong
        | CtRAnnotationRef
        | CtRFootnoteRef
        | CtREndnoteRef
        | CtRSeparator
        | CtRContinuationSeparator
        | CtRSym
        | CtRPgNum
        | CtRCr
        | CtRTab
        | CTObject
        | Pict
        | FldChar
        | CTRuby
        | FootnoteReference1
        | EndnoteReference1
        | CtRCommentReference
        | Drawing
        | CtRPtab
        | CtRLastRenderedPageBreak
        | AlternateContent
        | CTText
        | Ins1
        | Del1
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
                    "type": ForwardRef("T"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "delText",
                    "type": ForwardRef("DelText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "instrText",
                    "type": ForwardRef("InstrText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "delInstrText",
                    "type": ForwardRef("DelInstrText"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "noBreakHyphen",
                    "type": ForwardRef("CtRNoBreakHyphen"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "softHyphen",
                    "type": ForwardRef("CtRSoftHyphen"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dayShort",
                    "type": ForwardRef("CtRDayShort"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "monthShort",
                    "type": ForwardRef("CtRMonthShort"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "yearShort",
                    "type": ForwardRef("CtRYearShort"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "dayLong",
                    "type": ForwardRef("CtRDayLong"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "monthLong",
                    "type": ForwardRef("CtRMonthLong"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "yearLong",
                    "type": ForwardRef("CtRYearLong"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "annotationRef",
                    "type": ForwardRef("CtRAnnotationRef"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "footnoteRef",
                    "type": ForwardRef("CtRFootnoteRef"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "endnoteRef",
                    "type": ForwardRef("CtREndnoteRef"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "separator",
                    "type": ForwardRef("CtRSeparator"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "continuationSeparator",
                    "type": ForwardRef("CtRContinuationSeparator"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "sym",
                    "type": ForwardRef("CtRSym"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "pgNum",
                    "type": ForwardRef("CtRPgNum"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "cr",
                    "type": ForwardRef("CtRCr"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "tab",
                    "type": ForwardRef("CtRTab"),
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
                    "type": ForwardRef("FootnoteReference1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "endnoteReference",
                    "type": ForwardRef("EndnoteReference1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "commentReference",
                    "type": ForwardRef("CtRCommentReference"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "drawing",
                    "type": ForwardRef("Drawing"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "ptab",
                    "type": ForwardRef("CtRPtab"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "lastRenderedPageBreak",
                    "type": ForwardRef("CtRLastRenderedPageBreak"),
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
                {
                    "name": "ins",
                    "type": ForwardRef("Ins1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "del",
                    "type": ForwardRef("Del1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRad(Child):
    class Meta:
        name = "CT_Rad"

    rad_pr: None | CTRadPr = field(
        default=None,
        metadata={
            "name": "radPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    deg: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSPre(Child):
    class Meta:
        name = "CT_SPre"

    s_pre_pr: None | CTSPrePr = field(
        default=None,
        metadata={
            "name": "sPrePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sub: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sup: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSSub(Child):
    class Meta:
        name = "CT_SSub"

    s_sub_pr: None | CTSSubPr = field(
        default=None,
        metadata={
            "name": "sSubPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sub: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSSubSup(Child):
    class Meta:
        name = "CT_SSubSup"

    s_sub_sup_pr: None | CTSSubSupPr = field(
        default=None,
        metadata={
            "name": "sSubSupPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sub: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sup: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSSup(Child):
    class Meta:
        name = "CT_SSup"

    s_sup_pr: None | CTSSupPr = field(
        default=None,
        metadata={
            "name": "sSupPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    e: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    sup: None | CTOMathArg = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class MathPr(CTMathPr):
    class Meta:
        name = "mathPr"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@dataclass(slots=True, kw_only=True)
class CTMCS(Child):
    class Meta:
        name = "CT_MCS"

    mc: list[CTMC] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class OMathPara(CTOMathPara):
    class Meta:
        name = "oMathPara"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@dataclass(slots=True, kw_only=True)
class CTMPr(Child):
    class Meta:
        name = "CT_MPr"

    base_jc: None | CTYAlign = field(
        default=None,
        metadata={
            "name": "baseJc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    plc_hide: None | CTOnOff = field(
        default=None,
        metadata={
            "name": "plcHide",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    r_sp_rule: None | CTSpacingRule = field(
        default=None,
        metadata={
            "name": "rSpRule",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    c_gp_rule: None | CTSpacingRule = field(
        default=None,
        metadata={
            "name": "cGpRule",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    r_sp: None | CTUnSignedInteger = field(
        default=None,
        metadata={
            "name": "rSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    c_sp: None | CTUnSignedInteger = field(
        default=None,
        metadata={
            "name": "cSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    c_gp: None | CTUnSignedInteger = field(
        default=None,
        metadata={
            "name": "cGp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    mcs: None | CTMCS = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTM(Child):
    class Meta:
        name = "CT_M"

    m_pr: None | CTMPr = field(
        default=None,
        metadata={
            "name": "mPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
    mr: list[CTMR] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOMath(Child):
    class Meta:
        name = "CT_OMath"

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
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart1
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart1
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart1
        | CustomXmlInsRangeEnd1
        | CustomXmlDelRangeStart1
        | CustomXmlDelRangeEnd1
        | CustomXmlMoveFromRangeStart1
        | CustomXmlMoveFromRangeEnd1
        | CustomXmlMoveToRangeStart1
        | CustomXmlMoveToRangeEnd1
        | RunIns
        | RunDel
        | MoveFrom1
        | MoveTo1
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
                    "type": ForwardRef("MoveFromRangeStart1"),
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
                    "type": ForwardRef("MoveToRangeStart1"),
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
                    "type": ForwardRef("CustomXmlInsRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd1"),
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
                    "type": ForwardRef("MoveFrom1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo1"),
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
class OMath(CTOMath):
    class Meta:
        name = "oMath"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@dataclass(slots=True, kw_only=True)
class CTOMathArg(Child):
    class Meta:
        name = "CT_OMathArg"

    arg_pr: None | CTOMathArgPr = field(
        default=None,
        metadata={
            "name": "argPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )
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
        | ProofErr
        | RangePermissionStart
        | CTPerm
        | CTBookmark
        | CTMarkupRange
        | MoveFromRangeStart1
        | CTMoveFromRangeEnd
        | CTMoveToRangeEnd
        | MoveToRangeStart1
        | CommentRangeStart
        | CommentRangeEnd
        | CustomXmlInsRangeStart1
        | CustomXmlInsRangeEnd1
        | CustomXmlDelRangeStart1
        | CustomXmlDelRangeEnd1
        | CustomXmlMoveFromRangeStart1
        | CustomXmlMoveFromRangeEnd1
        | CustomXmlMoveToRangeStart1
        | CustomXmlMoveToRangeEnd1
        | RunIns
        | RunDel
        | MoveFrom1
        | MoveTo1
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
                    "type": ForwardRef("MoveFromRangeStart1"),
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
                    "type": ForwardRef("MoveToRangeStart1"),
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
                    "type": ForwardRef("CustomXmlInsRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlInsRangeEnd",
                    "type": ForwardRef("CustomXmlInsRangeEnd1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeStart",
                    "type": ForwardRef("CustomXmlDelRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlDelRangeEnd",
                    "type": ForwardRef("CustomXmlDelRangeEnd1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeStart",
                    "type": ForwardRef("CustomXmlMoveFromRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveFromRangeEnd",
                    "type": ForwardRef("CustomXmlMoveFromRangeEnd1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeStart",
                    "type": ForwardRef("CustomXmlMoveToRangeStart1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "customXmlMoveToRangeEnd",
                    "type": ForwardRef("CustomXmlMoveToRangeEnd1"),
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
                    "type": ForwardRef("MoveFrom1"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("MoveTo1"),
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
    ctrl_pr: None | CTCtrlPr = field(
        default=None,
        metadata={
            "name": "ctrlPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.mce import AlternateContent
from docx4j_py.shared_types import (
    StOnOff,
    StXalign,
    StYalign,
)
from docx4j_py.wml import (
    Br,
    CommentRangeEnd,
    CommentRangeStart,
    CTBookmark,
    CTMarkupRange,
    CTMoveFromRangeEnd,
    CTMoveToRangeEnd,
    CTObject,
    CTPerm,
    CtRAnnotationRef,
    CtRCommentReference,
    CtRContinuationSeparator,
    CtRCr,
    CtRDayLong,
    CtRDayShort,
    CtREndnoteRef,
    CtRFootnoteRef,
    CtRLastRenderedPageBreak,
    CtRMonthLong,
    CtRMonthShort,
    CtRNoBreakHyphen,
    CtRPgNum,
    CtRPtab,
    CtRSeparator,
    CtRSoftHyphen,
    CtRSym,
    CtRTab,
    CTRuby,
    CtRYearLong,
    CtRYearShort,
    CustomXmlDelRangeEnd1,
    CustomXmlDelRangeStart1,
    CustomXmlInsRangeEnd1,
    CustomXmlInsRangeStart1,
    CustomXmlMoveFromRangeEnd1,
    CustomXmlMoveFromRangeStart1,
    CustomXmlMoveToRangeEnd1,
    CustomXmlMoveToRangeStart1,
    Del1,
    Del2,
    DelInstrText,
    DelText,
    Drawing,
    EndnoteReference1,
    FldChar,
    FootnoteReference1,
    Ins1,
    Ins2,
    InstrText,
    MoveFrom1,
    MoveFromRangeStart1,
    MoveTo1,
    MoveToRangeStart1,
    Pict,
    ProofErr,
    RangePermissionStart,
    RPr,
    RunDel,
    RunIns,
    T,
)
from docx4j_py.xml_ns import SpaceValue


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.math import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.math.el",
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
