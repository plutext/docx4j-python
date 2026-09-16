from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2006/wordml"


@dataclass(slots=True, kw_only=True)
class CTAcd(Child):
    class Meta:
        name = "CT_Acd"

    arg_value: None | str = field(
        default=None,
        metadata={
            "name": "argValue",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    fci_based_on: None | str = field(
        default=None,
        metadata={
            "name": "fciBasedOn",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    fci_index_based_on: None | bytes = field(
        default=None,
        metadata={
            "name": "fciIndexBasedOn",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )
    acd_name: None | str = field(
        default=None,
        metadata={
            "name": "acdName",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAcdKeymap(Child):
    class Meta:
        name = "CT_AcdKeymap"

    acd_name: None | str = field(
        default=None,
        metadata={
            "name": "acdName",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDocEvents(Child):
    class Meta:
        name = "CT_DocEvents"

    event_doc_new: None | object = field(
        default=None,
        metadata={
            "name": "eventDocNew",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_open: None | object = field(
        default=None,
        metadata={
            "name": "eventDocOpen",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_close: None | object = field(
        default=None,
        metadata={
            "name": "eventDocClose",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_sync: None | object = field(
        default=None,
        metadata={
            "name": "eventDocSync",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_xml_after_insert: None | object = field(
        default=None,
        metadata={
            "name": "eventDocXmlAfterInsert",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_xml_before_delete: None | object = field(
        default=None,
        metadata={
            "name": "eventDocXmlBeforeDelete",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_content_control_after_insert: None | object = field(
        default=None,
        metadata={
            "name": "eventDocContentControlAfterInsert",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_content_control_before_delete: None | object = field(
        default=None,
        metadata={
            "name": "eventDocContentControlBeforeDelete",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_content_control_on_exit: None | object = field(
        default=None,
        metadata={
            "name": "eventDocContentControlOnExit",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_content_control_on_enter: None | object = field(
        default=None,
        metadata={
            "name": "eventDocContentControlOnEnter",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_store_update: None | object = field(
        default=None,
        metadata={
            "name": "eventDocStoreUpdate",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_content_control_content_update: None | object = field(
        default=None,
        metadata={
            "name": "eventDocContentControlContentUpdate",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    event_doc_building_block_after_insert: None | object = field(
        default=None,
        metadata={
            "name": "eventDocBuildingBlockAfterInsert",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFci(Child):
    class Meta:
        name = "CT_Fci"

    fci_name: None | str = field(
        default=None,
        metadata={
            "name": "fciName",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    fci_index: None | bytes = field(
        default=None,
        metadata={
            "name": "fciIndex",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )
    sw_arg: None | bytes = field(
        default=None,
        metadata={
            "name": "swArg",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLongHexNumber(Child):
    class Meta:
        name = "CT_LongHexNumber"

    val: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 4,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMacroWll(Child):
    class Meta:
        name = "CT_MacroWll"

    macro_name: None | str = field(
        default=None,
        metadata={
            "name": "macroName",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMcd(Child):
    class Meta:
        name = "CT_Mcd"

    macro_name: None | str = field(
        default=None,
        metadata={
            "name": "macroName",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    menu_help: None | str = field(
        default=None,
        metadata={
            "name": "menuHelp",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    b_encrypt: None | str = field(
        default=None,
        metadata={
            "name": "bEncrypt",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    cmg: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
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


class StOnOff(Enum):
    TRUE = "true"
    FALSE = "false"
    ON = "on"
    OFF = "off"
    VALUE_0 = "0"
    VALUE_1 = "1"


@dataclass(slots=True, kw_only=True)
class CTAcdManifest(Child):
    class Meta:
        name = "CT_AcdManifest"

    acd_entry: list[CTAcdKeymap] = field(
        default_factory=ChildList,
        metadata={
            "name": "acdEntry",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAcds(Child):
    class Meta:
        name = "CT_Acds"

    acd: list[CTAcd] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMcds(Child):
    class Meta:
        name = "CT_Mcds"

    mcd: list[CTMcd] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Macro(CTMacroWll):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Wll(CTMacroWll):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTKeymap(Child):
    class Meta:
        name = "CT_Keymap"

    content: None | CTFci | Macro | CTAcdKeymap | Wll | CTLongHexNumber = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "fci",
                    "type": ForwardRef("CTFci"),
                    "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
                },
                {
                    "name": "macro",
                    "type": ForwardRef("Macro"),
                    "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
                },
                {
                    "name": "acd",
                    "type": ForwardRef("CTAcdKeymap"),
                    "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
                },
                {
                    "name": "wll",
                    "type": ForwardRef("Wll"),
                    "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
                },
                {
                    "name": "wch",
                    "type": ForwardRef("CTLongHexNumber"),
                    "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
                },
            ),
        },
    )
    chm_primary: None | bytes = field(
        default=None,
        metadata={
            "name": "chmPrimary",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )
    chm_secondary: None | bytes = field(
        default=None,
        metadata={
            "name": "chmSecondary",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )
    kcm_primary: None | bytes = field(
        default=None,
        metadata={
            "name": "kcmPrimary",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )
    kcm_secondary: None | bytes = field(
        default=None,
        metadata={
            "name": "kcmSecondary",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "length": 2,
            "format": "base16",
        },
    )
    mask: None | StOnOff = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTToolbars(Child):
    class Meta:
        name = "CT_Toolbars"

    acd_manifest: None | CTAcdManifest = field(
        default=None,
        metadata={
            "name": "acdManifest",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    toolbar_data: None | CTRel = field(
        default=None,
        metadata={
            "name": "toolbarData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTVbaSuppData(Child):
    class Meta:
        name = "CT_VbaSuppData"

    doc_events: None | CTDocEvents = field(
        default=None,
        metadata={
            "name": "docEvents",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    mcds: None | CTMcds = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
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
class CTKeymaps(Child):
    class Meta:
        name = "CT_Keymaps"

    keymap: list[CTKeymap] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class VbaSuppData(CTVbaSuppData):
    class Meta:
        name = "vbaSuppData"
        namespace = "http://schemas.microsoft.com/office/word/2006/wordml"


@dataclass(slots=True, kw_only=True)
class CTTcg(Child):
    class Meta:
        name = "CT_Tcg"

    keymaps: None | CTKeymaps = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    keymaps_bad: None | CTKeymaps = field(
        default=None,
        metadata={
            "name": "keymapsBad",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    toolbars: None | CTToolbars = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
        },
    )
    acds: None | CTAcds = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2006/wordml",
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
class Tcg(CTTcg):
    class Meta:
        name = "tcg"
        namespace = "http://schemas.microsoft.com/office/word/2006/wordml"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.wne import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.wne.el",
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
