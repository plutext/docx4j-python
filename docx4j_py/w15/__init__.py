from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_py.child import Child, ChildList
from docx4j_py.wml import (
    BooleanDefaultTrue as WmlBooleanDefaultTrue,
)
from docx4j_py.wml import (
    CTColor as WmlCtcolor,
)
from docx4j_py.wml import (
    CTDataBinding as WmlCtdataBinding,
)
from docx4j_py.wml import (
    CTDecimalNumber as WmlCtdecimalNumber,
)
from docx4j_py.wml import (
    CTEmpty as WmlCtempty,
)

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTGuid(Child):
    class Meta:
        name = "CT_Guid"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresenceInfo(Child):
    class Meta:
        name = "CT_PresenceInfo"

    provider_id: None | str = field(
        default=None,
        metadata={
            "name": "providerId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    user_id: None | str = field(
        default=None,
        metadata={
            "name": "userId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )


class STSdtAppearance(Enum):
    BOUNDING_BOX = "boundingBox"
    TAGS = "tags"
    HIDDEN = "hidden"


@dataclass(slots=True, kw_only=True)
class BooleanDefaultTrue(WmlBooleanDefaultTrue):
    class Meta:
        name = "chartTrackingRefBased"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTColor(WmlCtcolor):
    class Meta:
        name = "color"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTCommentEx(Child):
    class Meta:
        name = "CT_CommentEx"

    para_id: None | str = field(
        default=None,
        metadata={
            "name": "paraId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    para_id_parent: None | str = field(
        default=None,
        metadata={
            "name": "paraIdParent",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    done: None | StOnOff = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataBinding(WmlCtdataBinding):
    class Meta:
        name = "dataBinding"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTDecimalNumber(WmlCtdecimalNumber):
    class Meta:
        name = "footnoteColumns"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTEmpty(WmlCtempty):
    class Meta:
        name = "repeatingSectionItem"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTPerson(Child):
    class Meta:
        name = "CT_Person"

    presence_info: None | CTPresenceInfo = field(
        default=None,
        metadata={
            "name": "presenceInfo",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    author: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    contact: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtAppearance(Child):
    class Meta:
        name = "CT_SdtAppearance"

    val: None | STSdtAppearance = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtRepeatedSection(Child):
    class Meta:
        name = "CT_SdtRepeatedSection"

    section_title: None | CTString = field(
        default=None,
        metadata={
            "name": "sectionTitle",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )
    do_not_allow_insert_delete_section: None | WmlBooleanDefaultTrue = field(
        default=None,
        metadata={
            "name": "doNotAllowInsertDeleteSection",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Collapsed(WmlBooleanDefaultTrue):
    class Meta:
        name = "collapsed"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class DocId(CTGuid):
    class Meta:
        name = "docId"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class WebExtensionCreated(WmlBooleanDefaultTrue):
    class Meta:
        name = "webExtensionCreated"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class WebExtensionLinked(WmlBooleanDefaultTrue):
    class Meta:
        name = "webExtensionLinked"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CTCommentsEx(Child):
    class Meta:
        name = "CT_CommentsEx"

    comment_ex: list[CTCommentEx] = field(
        default_factory=ChildList,
        metadata={
            "name": "commentEx",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
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
class CTPeople(Child):
    class Meta:
        name = "CT_People"

    person: list[CTPerson] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2012/wordml",
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
class Appearance(CTSdtAppearance):
    class Meta:
        name = "appearance"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class RepeatingSection(CTSdtRepeatedSection):
    class Meta:
        name = "repeatingSection"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class CommentsEx(CTCommentsEx):
    class Meta:
        name = "commentsEx"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


@dataclass(slots=True, kw_only=True)
class People(CTPeople):
    class Meta:
        name = "people"
        namespace = "http://schemas.microsoft.com/office/word/2012/wordml"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.wml import (
    CTString,
    StOnOff,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.w15 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.w15.el",
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
