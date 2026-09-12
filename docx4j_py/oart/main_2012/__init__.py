from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2012/main"


@dataclass(slots=True, kw_only=True)
class CTNonVisualGroupProps(Child):
    class Meta:
        name = "CT_NonVisualGroupProps"

    is_legacy_group: None | bool = field(
        default=None,
        metadata={
            "name": "isLegacyGroup",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTObjectPr(Child):
    class Meta:
        name = "CT_ObjectPr"

    object_id: None | str = field(
        default=None,
        metadata={
            "name": "objectId",
            "type": "Attribute",
        },
    )
    is_active_x: None | bool = field(
        default=None,
        metadata={
            "name": "isActiveX",
            "type": "Attribute",
        },
    )
    link_type: None | str = field(
        default=None,
        metadata={
            "name": "linkType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSignatureLine(Child):
    class Meta:
        name = "CT_SignatureLine"

    is_signature_line: None | bool = field(
        default=None,
        metadata={
            "name": "isSignatureLine",
            "type": "Attribute",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    prov_id: None | str = field(
        default=None,
        metadata={
            "name": "provId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    signing_instructions_set: None | bool = field(
        default=None,
        metadata={
            "name": "signingInstructionsSet",
            "type": "Attribute",
        },
    )
    allow_comments: None | bool = field(
        default=None,
        metadata={
            "name": "allowComments",
            "type": "Attribute",
        },
    )
    show_sign_date: None | bool = field(
        default=None,
        metadata={
            "name": "showSignDate",
            "type": "Attribute",
        },
    )
    suggested_signer: None | str = field(
        default=None,
        metadata={
            "name": "suggestedSigner",
            "type": "Attribute",
        },
    )
    suggested_signer2: None | str = field(
        default=None,
        metadata={
            "name": "suggestedSigner2",
            "type": "Attribute",
        },
    )
    suggested_signer_email: None | str = field(
        default=None,
        metadata={
            "name": "suggestedSignerEmail",
            "type": "Attribute",
        },
    )
    signing_instructions: None | str = field(
        default=None,
        metadata={
            "name": "signingInstructions",
            "type": "Attribute",
        },
    )
    addl_xml: None | str = field(
        default=None,
        metadata={
            "name": "addlXml",
            "type": "Attribute",
        },
    )
    sig_prov_url: None | str = field(
        default=None,
        metadata={
            "name": "sigProvUrl",
            "type": "Attribute",
        },
    )


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
class CTBackgroundPr(Child):
    class Meta:
        name = "CT_BackgroundPr"

    bw_mode: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwMode",
            "type": "Attribute",
        },
    )
    bw_pure: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwPure",
            "type": "Attribute",
        },
    )
    bw_normal: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwNormal",
            "type": "Attribute",
        },
    )
    target_screen_size: None | StTargetScreenSz = field(
        default=None,
        metadata={
            "name": "targetScreenSize",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class NonVisualGroupProps(CTNonVisualGroupProps):
    class Meta:
        name = "nonVisualGroupProps"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/main"


@dataclass(slots=True, kw_only=True)
class ObjectPr(CTObjectPr):
    class Meta:
        name = "objectPr"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/main"


@dataclass(slots=True, kw_only=True)
class SignatureLine(CTSignatureLine):
    class Meta:
        name = "signatureLine"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/main"


@dataclass(slots=True, kw_only=True)
class BackgroundPr(CTBackgroundPr):
    class Meta:
        name = "backgroundPr"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/main"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import STBlackWhiteMode


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.main_2012 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.main_2012.el",
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
