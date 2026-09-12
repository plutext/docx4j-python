from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_xsdata.models.datatype import XmlDateTime

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"


@dataclass(slots=True, kw_only=True)
class CoreProperties(Child):
    class Meta:
        name = "coreProperties"
        namespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

    category: None | str = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    content_status: None | str = field(
        default=None,
        metadata={
            "name": "contentStatus",
            "type": "Element",
        },
    )
    content_type: None | str = field(
        default=None,
        metadata={
            "name": "contentType",
            "type": "Element",
        },
    )
    created: None | Created = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://purl.org/dc/terms/",
        },
    )
    creator: None | Creator = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://purl.org/dc/elements/1.1/",
        },
    )
    description: None | Abstract | TableOfContents | Description = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
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
            ),
        },
    )
    identifier: None | BibliographicCitation | Identifier = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
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
            ),
        },
    )
    keywords: None | str = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    language: None | Language = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://purl.org/dc/elements/1.1/",
        },
    )
    last_modified_by: None | str = field(
        default=None,
        metadata={
            "name": "lastModifiedBy",
            "type": "Element",
        },
    )
    last_printed: None | XmlDateTime = field(
        default=None,
        metadata={
            "name": "lastPrinted",
            "type": "Element",
        },
    )
    modified: None | Modified = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://purl.org/dc/terms/",
        },
    )
    revision: None | str = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )
    subject: None | Subject = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://purl.org/dc/elements/1.1/",
        },
    )
    title: None | SimpleLiteral | Title = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
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
    version: None | str = field(
        default=None,
        metadata={
            "type": "Element",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.docprops.dc import (
    Creator,
    Description,
    Identifier,
    Language,
    Subject,
    Title,
)
from docx4j_py.docprops.dcterms import (
    Abstract,
    BibliographicCitation,
    Created,
    Modified,
    SimpleLiteral,
    TableOfContents,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.docprops.core import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.docprops.core.el",
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
