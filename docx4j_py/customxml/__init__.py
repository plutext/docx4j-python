from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/schemaLibrary/2006/main"


@dataclass(slots=True, kw_only=True)
class SchemaLibrarySchema(Child):
    class Meta:
        global_type = False

    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
            "schema_default": "",
        },
    )
    manifest_location: None | str = field(
        default=None,
        metadata={
            "name": "manifestLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
        },
    )
    schema_location: None | str = field(
        default=None,
        metadata={
            "name": "schemaLocation",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class SchemaLibrary(Child):
    class Meta:
        name = "schemaLibrary"
        namespace = "http://schemas.openxmlformats.org/schemaLibrary/2006/main"

    schema: list[SchemaLibrarySchema] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.customxml import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.customxml.el",
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
