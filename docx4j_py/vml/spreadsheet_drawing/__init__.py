from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "urn:schemas-microsoft-com:office:excel"


@dataclass(slots=True, kw_only=True)
class Accel(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Accel2(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Anchor(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Checked(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Column(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class DropLines(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class DropStyle(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Dx(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class FmlaGroup(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class FmlaLink(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class FmlaMacro(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class FmlaPict(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class FmlaRange(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class FmlaTxbx(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Inc(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Lct(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class ListItem(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Max(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Min(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class MultiSel(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Page(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Row(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


class STCF(Enum):
    PICT_OLD = "PictOld"
    PICT = "Pict"
    BITMAP = "Bitmap"
    PICT_PRINT = "PictPrint"
    PICT_SCREEN = "PictScreen"


class STObjectType(Enum):
    BUTTON = "Button"
    CHECKBOX = "Checkbox"
    DIALOG = "Dialog"
    DROP = "Drop"
    EDIT = "Edit"
    GBOX = "GBox"
    LABEL = "Label"
    LINE_A = "LineA"
    LIST = "List"
    MOVIE = "Movie"
    NOTE = "Note"
    PICT = "Pict"
    RADIO = "Radio"
    RECT_A = "RectA"
    SCROLL = "Scroll"
    SPIN = "Spin"
    SHAPE = "Shape"
    GROUP = "Group"
    RECT = "Rect"


class StTrueFalseBlank(Enum):
    TRUE = "True"
    T = "t"
    FALSE = "False"
    F = "f"
    VALUE = ""


@dataclass(slots=True, kw_only=True)
class ScriptExtended(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class ScriptLanguage(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class ScriptLocation(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class ScriptText(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Sel(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class SelType(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class TextHalign(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class TextValign(Child):
    class Meta:
        global_type = False

    value: None | str = field(default=None)


@dataclass(slots=True, kw_only=True)
class Vtedit(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Val(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class WidthMin(Child):
    class Meta:
        global_type = False

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class AutoFill(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class AutoLine(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class AutoPict(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class AutoScale(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Camera(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Cancel(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class ColHidden(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Colored(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Dde(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Default(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class DefaultSize(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Disabled(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Dismiss(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class FirstButton(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Help(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Horiz(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class JustLastX(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class LockText(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Locked(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class MapOcx(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class MoveWithCells(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class MultiLine(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class NoThreeD(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class NoThreeD2(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class PrintObject(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class RecalcAlways(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class RowHidden(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class SecretEdit(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class SizeWithCells(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Uiobj(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Vscroll(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class ValidIds(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class Visible(Child):
    class Meta:
        global_type = False

    value: None | StTrueFalseBlank = field(default=None)


@dataclass(slots=True, kw_only=True)
class CTClientData(Child):
    class Meta:
        name = "CT_ClientData"

    content: list[
        MoveWithCells
        | SizeWithCells
        | Anchor
        | Locked
        | DefaultSize
        | PrintObject
        | Disabled
        | AutoFill
        | AutoLine
        | AutoPict
        | FmlaMacro
        | TextHalign
        | TextValign
        | LockText
        | JustLastX
        | SecretEdit
        | Default
        | Help
        | Cancel
        | Dismiss
        | Accel
        | Accel2
        | Row
        | Column
        | Visible
        | RowHidden
        | ColHidden
        | Vtedit
        | MultiLine
        | Vscroll
        | ValidIds
        | FmlaRange
        | WidthMin
        | Sel
        | NoThreeD2
        | SelType
        | MultiSel
        | Lct
        | ListItem
        | DropStyle
        | Colored
        | DropLines
        | Checked
        | FmlaLink
        | FmlaPict
        | NoThreeD
        | FirstButton
        | FmlaGroup
        | Val
        | Min
        | Max
        | Inc
        | Page
        | Horiz
        | Dx
        | MapOcx
        | STCF
        | Camera
        | RecalcAlways
        | AutoScale
        | Dde
        | Uiobj
        | ScriptText
        | ScriptExtended
        | ScriptLanguage
        | ScriptLocation
        | FmlaTxbx
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "MoveWithCells",
                    "type": ForwardRef("MoveWithCells"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "SizeWithCells",
                    "type": ForwardRef("SizeWithCells"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Anchor",
                    "type": ForwardRef("Anchor"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Locked",
                    "type": ForwardRef("Locked"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "DefaultSize",
                    "type": ForwardRef("DefaultSize"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "PrintObject",
                    "type": ForwardRef("PrintObject"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Disabled",
                    "type": ForwardRef("Disabled"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "AutoFill",
                    "type": ForwardRef("AutoFill"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "AutoLine",
                    "type": ForwardRef("AutoLine"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "AutoPict",
                    "type": ForwardRef("AutoPict"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FmlaMacro",
                    "type": ForwardRef("FmlaMacro"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "TextHAlign",
                    "type": ForwardRef("TextHalign"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "TextVAlign",
                    "type": ForwardRef("TextValign"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "LockText",
                    "type": ForwardRef("LockText"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "JustLastX",
                    "type": ForwardRef("JustLastX"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "SecretEdit",
                    "type": ForwardRef("SecretEdit"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Default",
                    "type": ForwardRef("Default"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Help",
                    "type": ForwardRef("Help"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Cancel",
                    "type": ForwardRef("Cancel"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Dismiss",
                    "type": ForwardRef("Dismiss"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Accel",
                    "type": ForwardRef("Accel"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Accel2",
                    "type": ForwardRef("Accel2"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Row",
                    "type": ForwardRef("Row"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Column",
                    "type": ForwardRef("Column"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Visible",
                    "type": ForwardRef("Visible"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "RowHidden",
                    "type": ForwardRef("RowHidden"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ColHidden",
                    "type": ForwardRef("ColHidden"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "VTEdit",
                    "type": ForwardRef("Vtedit"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "MultiLine",
                    "type": ForwardRef("MultiLine"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "VScroll",
                    "type": ForwardRef("Vscroll"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ValidIds",
                    "type": ForwardRef("ValidIds"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FmlaRange",
                    "type": ForwardRef("FmlaRange"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "WidthMin",
                    "type": ForwardRef("WidthMin"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Sel",
                    "type": ForwardRef("Sel"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "NoThreeD2",
                    "type": ForwardRef("NoThreeD2"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "SelType",
                    "type": ForwardRef("SelType"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "MultiSel",
                    "type": ForwardRef("MultiSel"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "LCT",
                    "type": ForwardRef("Lct"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ListItem",
                    "type": ForwardRef("ListItem"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "DropStyle",
                    "type": ForwardRef("DropStyle"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Colored",
                    "type": ForwardRef("Colored"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "DropLines",
                    "type": ForwardRef("DropLines"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Checked",
                    "type": ForwardRef("Checked"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FmlaLink",
                    "type": ForwardRef("FmlaLink"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FmlaPict",
                    "type": ForwardRef("FmlaPict"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "NoThreeD",
                    "type": ForwardRef("NoThreeD"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FirstButton",
                    "type": ForwardRef("FirstButton"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FmlaGroup",
                    "type": ForwardRef("FmlaGroup"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Val",
                    "type": ForwardRef("Val"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Min",
                    "type": ForwardRef("Min"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Max",
                    "type": ForwardRef("Max"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Inc",
                    "type": ForwardRef("Inc"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Page",
                    "type": ForwardRef("Page"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Horiz",
                    "type": ForwardRef("Horiz"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Dx",
                    "type": ForwardRef("Dx"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "MapOCX",
                    "type": ForwardRef("MapOcx"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "CF",
                    "type": ForwardRef("STCF"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "Camera",
                    "type": ForwardRef("Camera"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "RecalcAlways",
                    "type": ForwardRef("RecalcAlways"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "AutoScale",
                    "type": ForwardRef("AutoScale"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "DDE",
                    "type": ForwardRef("Dde"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "UIObj",
                    "type": ForwardRef("Uiobj"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ScriptText",
                    "type": ForwardRef("ScriptText"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ScriptExtended",
                    "type": ForwardRef("ScriptExtended"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ScriptLanguage",
                    "type": ForwardRef("ScriptLanguage"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "ScriptLocation",
                    "type": ForwardRef("ScriptLocation"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "FmlaTxbx",
                    "type": ForwardRef("FmlaTxbx"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
            ),
        },
    )
    object_type: None | STObjectType = field(
        default=None,
        metadata={
            "name": "ObjectType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class ClientData(CTClientData):
    class Meta:
        namespace = "urn:schemas-microsoft-com:office:excel"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.vml.spreadsheet_drawing import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.vml.spreadsheet_drawing.el",
    "element_name": "docx4j_py.traversal",
    "find": "docx4j_py.traversal",
    "iter_nodes": "docx4j_py.traversal",
    "link_parents": "docx4j_py.child",
    "run_items_of": "docx4j_py.traversal",
    "text_of": "docx4j_py.traversal",
    "to_xml": "docx4j_py.fragments",
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
