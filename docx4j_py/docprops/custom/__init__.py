from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"


@dataclass(slots=True, kw_only=True)
class PropertiesProperty(Child):
    class Meta:
        global_type = False

    content: (
        None
        | Vector
        | Array
        | Blob
        | Oblob
        | Empty
        | Null
        | I1
        | I2
        | I4
        | I8
        | Int
        | Ui1
        | Ui2
        | Ui4
        | Ui8
        | Uint
        | R4
        | R8
        | DecimalType
        | Lpstr
        | Lpwstr
        | Bstr
        | Date
        | Filetime
        | Bool
        | Cy
        | Error
        | Stream
        | Ostream
        | Storage
        | Ostorage
        | Vstream
        | Clsid
        | Cf
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "vector",
                    "type": ForwardRef("Vector"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "array",
                    "type": ForwardRef("Array"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "blob",
                    "type": ForwardRef("Blob"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "oblob",
                    "type": ForwardRef("Oblob"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "empty",
                    "type": ForwardRef("Empty"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "null",
                    "type": ForwardRef("Null"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "i1",
                    "type": ForwardRef("I1"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "i2",
                    "type": ForwardRef("I2"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "i4",
                    "type": ForwardRef("I4"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "i8",
                    "type": ForwardRef("I8"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "int",
                    "type": ForwardRef("Int"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "ui1",
                    "type": ForwardRef("Ui1"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "ui2",
                    "type": ForwardRef("Ui2"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "ui4",
                    "type": ForwardRef("Ui4"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "ui8",
                    "type": ForwardRef("Ui8"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "uint",
                    "type": ForwardRef("Uint"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "r4",
                    "type": ForwardRef("R4"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "r8",
                    "type": ForwardRef("R8"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "decimal",
                    "type": ForwardRef("DecimalType"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "lpstr",
                    "type": ForwardRef("Lpstr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "lpwstr",
                    "type": ForwardRef("Lpwstr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "bstr",
                    "type": ForwardRef("Bstr"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "date",
                    "type": ForwardRef("Date"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "filetime",
                    "type": ForwardRef("Filetime"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "bool",
                    "type": ForwardRef("Bool"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "cy",
                    "type": ForwardRef("Cy"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "error",
                    "type": ForwardRef("Error"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "stream",
                    "type": ForwardRef("Stream"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "ostream",
                    "type": ForwardRef("Ostream"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "storage",
                    "type": ForwardRef("Storage"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "ostorage",
                    "type": ForwardRef("Ostorage"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "vstream",
                    "type": ForwardRef("Vstream"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "clsid",
                    "type": ForwardRef("Clsid"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "cf",
                    "type": ForwardRef("Cf"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
            ),
        },
    )
    fmtid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\s*\{[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}\}\s*",
        },
    )
    pid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    link_target: None | str = field(
        default=None,
        metadata={
            "name": "linkTarget",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Properties(Child):
    class Meta:
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

    property: list[PropertiesProperty] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.docprops.variant_types import (
    I1,
    I2,
    I4,
    I8,
    R4,
    R8,
    Array,
    Blob,
    Bool,
    Bstr,
    Cf,
    Clsid,
    Cy,
    Date,
    DecimalType,
    Empty,
    Error,
    Filetime,
    Int,
    Lpstr,
    Lpwstr,
    Null,
    Oblob,
    Ostorage,
    Ostream,
    Storage,
    Stream,
    Ui1,
    Ui2,
    Ui4,
    Ui8,
    Uint,
    Vector,
    Vstream,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.docprops.custom import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.docprops.custom.el",
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
