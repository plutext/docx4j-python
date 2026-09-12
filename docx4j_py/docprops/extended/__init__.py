from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


@dataclass(slots=True, kw_only=True)
class PropertiesDigSig(Child):
    class Meta:
        global_type = False

    blob: None | Blob = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        },
    )


@dataclass(slots=True, kw_only=True)
class PropertiesHlinks(Child):
    class Meta:
        global_type = False

    vector: None | Vector = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        },
    )


@dataclass(slots=True, kw_only=True)
class PropertiesHeadingPairs(Child):
    class Meta:
        global_type = False

    vector: None | Vector = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        },
    )


@dataclass(slots=True, kw_only=True)
class PropertiesTitlesOfParts(Child):
    class Meta:
        global_type = False

    vector: None | Vector = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        },
    )


@dataclass(slots=True, kw_only=True)
class Properties(Child):
    class Meta:
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

    template: None | str = field(
        default=None,
        metadata={
            "name": "Template",
            "type": "Element",
        },
    )
    manager: None | str = field(
        default=None,
        metadata={
            "name": "Manager",
            "type": "Element",
        },
    )
    company: None | str = field(
        default=None,
        metadata={
            "name": "Company",
            "type": "Element",
        },
    )
    pages: None | int = field(
        default=None,
        metadata={
            "name": "Pages",
            "type": "Element",
        },
    )
    words: None | int = field(
        default=None,
        metadata={
            "name": "Words",
            "type": "Element",
        },
    )
    characters: None | int = field(
        default=None,
        metadata={
            "name": "Characters",
            "type": "Element",
        },
    )
    presentation_format: None | str = field(
        default=None,
        metadata={
            "name": "PresentationFormat",
            "type": "Element",
        },
    )
    lines: None | int = field(
        default=None,
        metadata={
            "name": "Lines",
            "type": "Element",
        },
    )
    paragraphs: None | int = field(
        default=None,
        metadata={
            "name": "Paragraphs",
            "type": "Element",
        },
    )
    slides: None | int = field(
        default=None,
        metadata={
            "name": "Slides",
            "type": "Element",
        },
    )
    notes: None | int = field(
        default=None,
        metadata={
            "name": "Notes",
            "type": "Element",
        },
    )
    total_time: None | int = field(
        default=None,
        metadata={
            "name": "TotalTime",
            "type": "Element",
        },
    )
    hidden_slides: None | int = field(
        default=None,
        metadata={
            "name": "HiddenSlides",
            "type": "Element",
        },
    )
    mmclips: None | int = field(
        default=None,
        metadata={
            "name": "MMClips",
            "type": "Element",
        },
    )
    scale_crop: None | bool = field(
        default=None,
        metadata={
            "name": "ScaleCrop",
            "type": "Element",
        },
    )
    heading_pairs: None | PropertiesHeadingPairs = field(
        default=None,
        metadata={
            "name": "HeadingPairs",
            "type": "Element",
        },
    )
    titles_of_parts: None | PropertiesTitlesOfParts = field(
        default=None,
        metadata={
            "name": "TitlesOfParts",
            "type": "Element",
        },
    )
    links_up_to_date: None | bool = field(
        default=None,
        metadata={
            "name": "LinksUpToDate",
            "type": "Element",
        },
    )
    characters_with_spaces: None | int = field(
        default=None,
        metadata={
            "name": "CharactersWithSpaces",
            "type": "Element",
        },
    )
    shared_doc: None | bool = field(
        default=None,
        metadata={
            "name": "SharedDoc",
            "type": "Element",
        },
    )
    hyperlink_base: None | str = field(
        default=None,
        metadata={
            "name": "HyperlinkBase",
            "type": "Element",
        },
    )
    hlinks: None | PropertiesHlinks = field(
        default=None,
        metadata={
            "name": "HLinks",
            "type": "Element",
        },
    )
    hyperlinks_changed: None | bool = field(
        default=None,
        metadata={
            "name": "HyperlinksChanged",
            "type": "Element",
        },
    )
    dig_sig: None | PropertiesDigSig = field(
        default=None,
        metadata={
            "name": "DigSig",
            "type": "Element",
        },
    )
    application: None | str = field(
        default=None,
        metadata={
            "name": "Application",
            "type": "Element",
        },
    )
    app_version: None | str = field(
        default=None,
        metadata={
            "name": "AppVersion",
            "type": "Element",
        },
    )
    doc_security: None | int = field(
        default=None,
        metadata={
            "name": "DocSecurity",
            "type": "Element",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.docprops.variant_types import (
    Blob,
    Vector,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.docprops.extended import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.docprops.extended.el",
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
