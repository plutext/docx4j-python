from importlib import import_module
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx4j_py.word.msink import (
        Context,
        CTCtxLink,
        CTCtxNode,
        CTProperty,
        STDir,
        STKnownCtxNodeType,
        STKnownSemanticType,
    )
    from docx4j_py.word.wp14 import (
        CTSizeRelH,
        CTSizeRelV,
        PctPosHoffset,
        PctPosVoffset,
        SizeRelH,
        SizeRelV,
        STSizeRelFromH,
        STSizeRelFromV,
    )
    from docx4j_py.word.wp15 import (
        CTWebVideoPr,
        WebVideoPr,
    )
    from docx4j_py.word.wpc import (
        CTWordprocessingCanvas,
        Wpc,
    )
    from docx4j_py.word.wpg import (
        CTGraphicFrame,
        CTWordprocessingGroup,
        Wgp,
    )
    from docx4j_py.word.wps import (
        CTLinkedTextboxInformation,
        CTTextboxInfo,
        CTWordprocessingShape,
        Wsp,
    )

__all__ = [
    "CTCtxLink",
    "CTCtxNode",
    "CTProperty",
    "STDir",
    "STKnownCtxNodeType",
    "STKnownSemanticType",
    "Context",
    "CTSizeRelH",
    "CTSizeRelV",
    "STSizeRelFromH",
    "STSizeRelFromV",
    "PctPosHoffset",
    "PctPosVoffset",
    "SizeRelH",
    "SizeRelV",
    "CTWebVideoPr",
    "WebVideoPr",
    "CTWordprocessingCanvas",
    "Wpc",
    "CTGraphicFrame",
    "CTWordprocessingGroup",
    "Wgp",
    "CTLinkedTextboxInformation",
    "CTTextboxInfo",
    "CTWordprocessingShape",
    "Wsp",
]

# The re-exports above are resolved on first use, so that importing this
# package never forces a module of it to be imported before the modules it
# depends on.
__lazy_imports__ = {
    "CTCtxLink": ("docx4j_py.word.msink", "CTCtxLink"),
    "CTCtxNode": ("docx4j_py.word.msink", "CTCtxNode"),
    "CTProperty": ("docx4j_py.word.msink", "CTProperty"),
    "STDir": ("docx4j_py.word.msink", "STDir"),
    "STKnownCtxNodeType": ("docx4j_py.word.msink", "STKnownCtxNodeType"),
    "STKnownSemanticType": ("docx4j_py.word.msink", "STKnownSemanticType"),
    "Context": ("docx4j_py.word.msink", "Context"),
    "CTSizeRelH": ("docx4j_py.word.wp14", "CTSizeRelH"),
    "CTSizeRelV": ("docx4j_py.word.wp14", "CTSizeRelV"),
    "STSizeRelFromH": ("docx4j_py.word.wp14", "STSizeRelFromH"),
    "STSizeRelFromV": ("docx4j_py.word.wp14", "STSizeRelFromV"),
    "PctPosHoffset": ("docx4j_py.word.wp14", "PctPosHoffset"),
    "PctPosVoffset": ("docx4j_py.word.wp14", "PctPosVoffset"),
    "SizeRelH": ("docx4j_py.word.wp14", "SizeRelH"),
    "SizeRelV": ("docx4j_py.word.wp14", "SizeRelV"),
    "CTWebVideoPr": ("docx4j_py.word.wp15", "CTWebVideoPr"),
    "WebVideoPr": ("docx4j_py.word.wp15", "WebVideoPr"),
    "CTWordprocessingCanvas": ("docx4j_py.word.wpc", "CTWordprocessingCanvas"),
    "Wpc": ("docx4j_py.word.wpc", "Wpc"),
    "CTGraphicFrame": ("docx4j_py.word.wpg", "CTGraphicFrame"),
    "CTWordprocessingGroup": ("docx4j_py.word.wpg", "CTWordprocessingGroup"),
    "Wgp": ("docx4j_py.word.wpg", "Wgp"),
    "CTLinkedTextboxInformation": ("docx4j_py.word.wps", "CTLinkedTextboxInformation"),
    "CTTextboxInfo": ("docx4j_py.word.wps", "CTTextboxInfo"),
    "CTWordprocessingShape": ("docx4j_py.word.wps", "CTWordprocessingShape"),
    "Wsp": ("docx4j_py.word.wps", "Wsp"),
}


def __getattr__(name: str) -> object:
    """Import a re-exported class on first use."""
    target = __lazy_imports__.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

    value = getattr(import_module(target[0]), target[1])
    globals()[name] = value
    return value
