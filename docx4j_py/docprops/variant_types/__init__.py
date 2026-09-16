from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import ForwardRef

from docx4j_xsdata.models.datatype import XmlDateTime

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


class CtArrayBaseType(Enum):
    VARIANT = "variant"
    I1 = "i1"
    I2 = "i2"
    I4 = "i4"
    INT = "int"
    UI1 = "ui1"
    UI2 = "ui2"
    UI4 = "ui4"
    UINT = "uint"
    R4 = "r4"
    R8 = "r8"
    DECIMAL = "decimal"
    BSTR = "bstr"
    DATE = "date"
    BOOL = "bool"
    CY = "cy"
    ERROR = "error"


class CtVectorBaseType(Enum):
    VARIANT = "variant"
    I1 = "i1"
    I2 = "i2"
    I4 = "i4"
    I8 = "i8"
    UI1 = "ui1"
    UI2 = "ui2"
    UI4 = "ui4"
    UI8 = "ui8"
    R4 = "r4"
    R8 = "r8"
    LPSTR = "lpstr"
    LPWSTR = "lpwstr"
    BSTR = "bstr"
    DATE = "date"
    FILETIME = "filetime"
    BOOL = "bool"
    CY = "cy"
    ERROR = "error"
    CLSID = "clsid"
    CF = "cf"


@dataclass(slots=True, kw_only=True)
class Cf1(Child):
    class Meta:
        name = "CT_Cf"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )
    format: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\-1|\-2|\-3|[1-9]+|0",
        },
    )


@dataclass(slots=True, kw_only=True)
class Empty1(Child):
    class Meta:
        name = "CT_Empty"


@dataclass(slots=True, kw_only=True)
class Null1(Child):
    class Meta:
        name = "CT_Null"


@dataclass(slots=True, kw_only=True)
class Vstream1(Child):
    class Meta:
        name = "CT_Vstream"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )
    version: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\s*\{[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}\}\s*",
        },
    )


@dataclass(slots=True, kw_only=True)
class Blob(Child):
    class Meta:
        name = "blob"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class Bool(Child):
    class Meta:
        name = "bool"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | bool = field(default=None)


@dataclass(slots=True, kw_only=True)
class Bstr(Child):
    class Meta:
        name = "bstr"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: str = field(default="")


@dataclass(slots=True, kw_only=True)
class Clsid(Child):
    class Meta:
        name = "clsid"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: str = field(
        default="",
        metadata={
            "pattern": r"\s*\{[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}\}\s*",
        },
    )


@dataclass(slots=True, kw_only=True)
class Cy(Child):
    class Meta:
        name = "cy"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: str = field(
        default="",
        metadata={
            "pattern": r"\s*[0-9]*\.[0-9]{4}\s*",
        },
    )


@dataclass(slots=True, kw_only=True)
class Date(Child):
    class Meta:
        name = "date"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | XmlDateTime = field(default=None)


@dataclass(slots=True, kw_only=True)
class DecimalType(Child):
    class Meta:
        name = "decimal"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | Decimal = field(default=None)


@dataclass(slots=True, kw_only=True)
class Error(Child):
    class Meta:
        name = "error"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: str = field(
        default="",
        metadata={
            "pattern": r"\s*0x[0-9A-Za-z]{8}\s*",
        },
    )


@dataclass(slots=True, kw_only=True)
class Filetime(Child):
    class Meta:
        name = "filetime"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | XmlDateTime = field(default=None)


@dataclass(slots=True, kw_only=True)
class I1(Child):
    class Meta:
        name = "i1"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class I2(Child):
    class Meta:
        name = "i2"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class I4(Child):
    class Meta:
        name = "i4"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class I8(Child):
    class Meta:
        name = "i8"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Int(Child):
    class Meta:
        name = "int"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Lpstr(Child):
    class Meta:
        name = "lpstr"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: str = field(default="")


@dataclass(slots=True, kw_only=True)
class Lpwstr(Child):
    class Meta:
        name = "lpwstr"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: str = field(default="")


@dataclass(slots=True, kw_only=True)
class Oblob(Child):
    class Meta:
        name = "oblob"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class Ostorage(Child):
    class Meta:
        name = "ostorage"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class Ostream(Child):
    class Meta:
        name = "ostream"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class R4(Child):
    class Meta:
        name = "r4"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | float = field(default=None)


@dataclass(slots=True, kw_only=True)
class R8(Child):
    class Meta:
        name = "r8"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | float = field(default=None)


@dataclass(slots=True, kw_only=True)
class Storage(Child):
    class Meta:
        name = "storage"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class Stream(Child):
    class Meta:
        name = "stream"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class Ui1(Child):
    class Meta:
        name = "ui1"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Ui2(Child):
    class Meta:
        name = "ui2"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Ui4(Child):
    class Meta:
        name = "ui4"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Ui8(Child):
    class Meta:
        name = "ui8"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Uint(Child):
    class Meta:
        name = "uint"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class Array1(Child):
    class Meta:
        name = "CT_Array"

    content: list[
        Variant
        | I1
        | I2
        | I4
        | Int
        | Ui1
        | Ui2
        | Ui4
        | Uint
        | R4
        | R8
        | DecimalType
        | Bstr
        | Date
        | Bool
        | Error
        | Cy
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "variant",
                    "type": ForwardRef("Variant"),
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
                    "name": "bool",
                    "type": ForwardRef("Bool"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "error",
                    "type": ForwardRef("Error"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
                {
                    "name": "cy",
                    "type": ForwardRef("Cy"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
            ),
        },
    )
    l_bounds: None | int = field(
        default=None,
        metadata={
            "name": "lBounds",
            "type": "Attribute",
        },
    )
    u_bounds: None | int = field(
        default=None,
        metadata={
            "name": "uBounds",
            "type": "Attribute",
        },
    )
    base_type: None | CtArrayBaseType = field(
        default=None,
        metadata={
            "name": "baseType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Cf(Cf1):
    class Meta:
        name = "cf"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass(slots=True, kw_only=True)
class Empty(Empty1):
    class Meta:
        name = "empty"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass(slots=True, kw_only=True)
class Null(Null1):
    class Meta:
        name = "null"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass(slots=True, kw_only=True)
class Vstream(Vstream1):
    class Meta:
        name = "vstream"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass(slots=True, kw_only=True)
class Vector1(Child):
    class Meta:
        name = "CT_Vector"

    content: list[
        Variant
        | I1
        | I2
        | I4
        | I8
        | Ui1
        | Ui2
        | Ui4
        | Ui8
        | R4
        | R8
        | Lpstr
        | Lpwstr
        | Bstr
        | Date
        | Filetime
        | Bool
        | Cy
        | Error
        | Clsid
        | Cf
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "variant",
                    "type": ForwardRef("Variant"),
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
    base_type: None | CtVectorBaseType = field(
        default=None,
        metadata={
            "name": "baseType",
            "type": "Attribute",
        },
    )
    size: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Array(Array1):
    class Meta:
        name = "array"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass(slots=True, kw_only=True)
class Vector(Vector1):
    class Meta:
        name = "vector"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass(slots=True, kw_only=True)
class Variant1(Child):
    class Meta:
        name = "CT_Variant"

    content: (
        None
        | Variant
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
                    "name": "variant",
                    "type": ForwardRef("Variant"),
                    "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                },
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


@dataclass(slots=True, kw_only=True)
class Variant(Variant1):
    class Meta:
        name = "variant"
        namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.docprops.variant_types import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.docprops.variant_types.el",
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
