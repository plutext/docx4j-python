from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, ForwardRef

from docx4j_xsdata.models.datatype import XmlDate, XmlDateTime, XmlPeriod

from docx4j_py.child import Child, ChildList
from docx4j_py.docprops.dc import SimpleLiteral as DcSimpleLiteral

__NAMESPACE__ = "http://purl.org/dc/terms/"


@dataclass(slots=True, kw_only=True)
class SimpleLiteral(Child):
    class Meta:
        name = "alternative"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Abstract(Child):
    class Meta:
        name = "abstract"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class AccessRights(Child):
    class Meta:
        name = "accessRights"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Audience(Child):
    class Meta:
        name = "audience"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Available(Child):
    class Meta:
        name = "available"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class BibliographicCitation(Child):
    class Meta:
        name = "bibliographicCitation"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class ConformsTo(Child):
    class Meta:
        name = "conformsTo"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Created(Child):
    class Meta:
        name = "created"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class DateAccepted(Child):
    class Meta:
        name = "dateAccepted"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class DateCopyrighted(Child):
    class Meta:
        name = "dateCopyrighted"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class DateSubmitted(Child):
    class Meta:
        name = "dateSubmitted"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class EducationLevel(Child):
    class Meta:
        name = "educationLevel"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Extent(Child):
    class Meta:
        name = "extent"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class HasFormat(Child):
    class Meta:
        name = "hasFormat"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class HasPart(Child):
    class Meta:
        name = "hasPart"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class HasVersion(Child):
    class Meta:
        name = "hasVersion"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class IsFormatOf(Child):
    class Meta:
        name = "isFormatOf"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class IsPartOf(Child):
    class Meta:
        name = "isPartOf"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class IsReferencedBy(Child):
    class Meta:
        name = "isReferencedBy"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class IsReplacedBy(Child):
    class Meta:
        name = "isReplacedBy"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class IsRequiredBy(Child):
    class Meta:
        name = "isRequiredBy"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class IsVersionOf(Child):
    class Meta:
        name = "isVersionOf"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Issued(Child):
    class Meta:
        name = "issued"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Mediator(Child):
    class Meta:
        name = "mediator"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Medium(Child):
    class Meta:
        name = "medium"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Modified(Child):
    class Meta:
        name = "modified"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class References(Child):
    class Meta:
        name = "references"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Replaces(Child):
    class Meta:
        name = "replaces"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Requires(Child):
    class Meta:
        name = "requires"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Spatial(Child):
    class Meta:
        name = "spatial"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class TableOfContents(Child):
    class Meta:
        name = "tableOfContents"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Temporal(Child):
    class Meta:
        name = "temporal"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Valid(Child):
    class Meta:
        name = "valid"
        namespace = "http://purl.org/dc/terms/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Box(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class DCMIType(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: None | DcmitypeDcmitype = field(default=None)
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class DDC(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class ElementOrRefinementContainer(Child):
    class Meta:
        name = "elementOrRefinementContainer"

    content: list[
        EducationLevel
        | Mediator
        | Audience
        | AccessRights
        | Rights
        | Temporal
        | Spatial
        | Coverage
        | ConformsTo
        | HasFormat
        | IsFormatOf
        | References
        | IsReferencedBy
        | HasPart
        | IsPartOf
        | Requires
        | IsRequiredBy
        | Replaces
        | IsReplacedBy
        | HasVersion
        | IsVersionOf
        | Relation
        | Language
        | Source
        | BibliographicCitation
        | Identifier
        | Medium
        | Extent
        | Format
        | Type
        | DateSubmitted
        | DateCopyrighted
        | DateAccepted
        | Modified
        | Issued
        | Available
        | Valid
        | Created
        | Date
        | Contributor
        | Publisher
        | Abstract
        | TableOfContents
        | Description
        | Subject
        | Creator
        | SimpleLiteral
        | Title
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "educationLevel",
                    "type": ForwardRef("EducationLevel"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "mediator",
                    "type": ForwardRef("Mediator"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "audience",
                    "type": ForwardRef("Audience"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "accessRights",
                    "type": ForwardRef("AccessRights"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "rights",
                    "type": ForwardRef("Rights"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "temporal",
                    "type": ForwardRef("Temporal"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "spatial",
                    "type": ForwardRef("Spatial"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "coverage",
                    "type": ForwardRef("Coverage"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "conformsTo",
                    "type": ForwardRef("ConformsTo"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "hasFormat",
                    "type": ForwardRef("HasFormat"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "isFormatOf",
                    "type": ForwardRef("IsFormatOf"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "references",
                    "type": ForwardRef("References"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "isReferencedBy",
                    "type": ForwardRef("IsReferencedBy"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "hasPart",
                    "type": ForwardRef("HasPart"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "isPartOf",
                    "type": ForwardRef("IsPartOf"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "requires",
                    "type": ForwardRef("Requires"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "isRequiredBy",
                    "type": ForwardRef("IsRequiredBy"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "replaces",
                    "type": ForwardRef("Replaces"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "isReplacedBy",
                    "type": ForwardRef("IsReplacedBy"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "hasVersion",
                    "type": ForwardRef("HasVersion"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "isVersionOf",
                    "type": ForwardRef("IsVersionOf"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "relation",
                    "type": ForwardRef("Relation"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "language",
                    "type": ForwardRef("Language"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "source",
                    "type": ForwardRef("Source"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "bibliographicCitation",
                    "type": ForwardRef("BibliographicCitation"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "identifier",
                    "type": ForwardRef("Identifier"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "medium",
                    "type": ForwardRef("Medium"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "extent",
                    "type": ForwardRef("Extent"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "format",
                    "type": ForwardRef("Format"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "type",
                    "type": ForwardRef("Type"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "dateSubmitted",
                    "type": ForwardRef("DateSubmitted"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "dateCopyrighted",
                    "type": ForwardRef("DateCopyrighted"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "dateAccepted",
                    "type": ForwardRef("DateAccepted"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "modified",
                    "type": ForwardRef("Modified"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "issued",
                    "type": ForwardRef("Issued"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "available",
                    "type": ForwardRef("Available"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "valid",
                    "type": ForwardRef("Valid"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "created",
                    "type": ForwardRef("Created"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "date",
                    "type": ForwardRef("Date"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "contributor",
                    "type": ForwardRef("Contributor"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "publisher",
                    "type": ForwardRef("Publisher"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "abstract",
                    "type": ForwardRef("Abstract"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "tableOfContents",
                    "type": ForwardRef("TableOfContents"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "description",
                    "type": ForwardRef("Description"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "subject",
                    "type": ForwardRef("Subject"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "creator",
                    "type": ForwardRef("Creator"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
                {
                    "name": "alternative",
                    "type": ForwardRef("SimpleLiteral"),
                    "namespace": "http://purl.org/dc/terms/",
                },
                {
                    "name": "title",
                    "type": ForwardRef("Title"),
                    "namespace": "http://purl.org/dc/elements/1.1/",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class IMT(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class ISO3166(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class ISO6392(DcSimpleLiteral):
    class Meta:
        name = "ISO639-2"

    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class LCC(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class LCSH(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class MESH(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class Period(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class Point(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class RFC1766(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class RFC3066(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class TGN(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class UDC(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class URI(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: str = field(default="")
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


@dataclass(slots=True, kw_only=True)
class W3CDTF(DcSimpleLiteral):
    content: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )
    value: None | XmlPeriod | XmlDate | XmlDateTime = field(default=None)
    lang: Any = field(
        init=False,
        default=None,
        metadata={
            "type": "Ignore",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.docprops.dc import (
    Contributor,
    Coverage,
    Creator,
    Date,
    Description,
    Format,
    Identifier,
    Language,
    Publisher,
    Relation,
    Rights,
    Source,
    Subject,
    Title,
    Type,
)
from docx4j_py.docprops.dcmitype import DCMIType as DcmitypeDcmitype


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.docprops.dcterms import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.docprops.dcterms.el",
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
