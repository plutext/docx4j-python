from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://purl.org/dc/elements/1.1/"


@dataclass(slots=True, kw_only=True)
class Contributor(Child):
    class Meta:
        name = "contributor"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Coverage(Child):
    class Meta:
        name = "coverage"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Creator(Child):
    class Meta:
        name = "creator"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Date(Child):
    class Meta:
        name = "date"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Description(Child):
    class Meta:
        name = "description"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Format(Child):
    class Meta:
        name = "format"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Identifier(Child):
    class Meta:
        name = "identifier"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Language(Child):
    class Meta:
        name = "language"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Publisher(Child):
    class Meta:
        name = "publisher"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Relation(Child):
    class Meta:
        name = "relation"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Rights(Child):
    class Meta:
        name = "rights"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Source(Child):
    class Meta:
        name = "source"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Subject(Child):
    class Meta:
        name = "subject"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Title(Child):
    class Meta:
        name = "title"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class Type(Child):
    class Meta:
        name = "type"
        namespace = "http://purl.org/dc/elements/1.1/"

    any_element: None | object = field(
        default=None,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class ElementContainer(Child):
    class Meta:
        name = "elementContainer"

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
        | DctermsSimpleLiteral
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
                    "type": ForwardRef("DctermsSimpleLiteral"),
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
class SimpleLiteral(Child):
    lang: None | str | LangValue = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    content: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
            "mixed": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class AnyType(SimpleLiteral):
    class Meta:
        name = "any"
        namespace = "http://purl.org/dc/elements/1.1/"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.docprops.dcterms import (
    Abstract,
    AccessRights,
    Audience,
    Available,
    BibliographicCitation,
    ConformsTo,
    Created,
    DateAccepted,
    DateCopyrighted,
    DateSubmitted,
    EducationLevel,
    Extent,
    HasFormat,
    HasPart,
    HasVersion,
    IsFormatOf,
    IsPartOf,
    IsReferencedBy,
    IsReplacedBy,
    IsRequiredBy,
    Issued,
    IsVersionOf,
    Mediator,
    Medium,
    Modified,
    References,
    Replaces,
    Requires,
    Spatial,
    TableOfContents,
    Temporal,
    Valid,
)
from docx4j_py.docprops.dcterms import (
    SimpleLiteral as DctermsSimpleLiteral,
)
from docx4j_py.xml_ns import LangValue


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.docprops.dc import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.docprops.dc.el",
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
