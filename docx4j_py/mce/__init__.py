from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/markup-compatibility/2006"


@dataclass(slots=True, kw_only=True)
class AlternateContentChoice(Child):
    class Meta:
        global_type = False

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )
    requires: None | str = field(
        default=None,
        metadata={
            "name": "Requires",
            "type": "Attribute",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    must_understand: None | str = field(
        default=None,
        metadata={
            "name": "MustUnderstand",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    process_content: None | str = field(
        default=None,
        metadata={
            "name": "ProcessContent",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )


@dataclass(slots=True, kw_only=True)
class AlternateContentFallback(Child):
    class Meta:
        global_type = False

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    must_understand: None | str = field(
        default=None,
        metadata={
            "name": "MustUnderstand",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    process_content: None | str = field(
        default=None,
        metadata={
            "name": "ProcessContent",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )


@dataclass(slots=True, kw_only=True)
class AlternateContent(Child):
    class Meta:
        namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006"

    choice: list[AlternateContentChoice] = field(
        default_factory=ChildList,
        metadata={
            "name": "Choice",
            "type": "Element",
        },
    )
    fallback: None | AlternateContentFallback = field(
        default=None,
        metadata={
            "name": "Fallback",
            "type": "Element",
        },
    )
    ignorable: None | str = field(
        default=None,
        metadata={
            "name": "Ignorable",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    must_understand: None | str = field(
        default=None,
        metadata={
            "name": "MustUnderstand",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )
    process_content: None | str = field(
        default=None,
        metadata={
            "name": "ProcessContent",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        },
    )


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.mce import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.mce.el",
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
