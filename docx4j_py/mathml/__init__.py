from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class CiType(Child):
    class Meta:
        name = "ci.type"

    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
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
class ConstantType(Child):
    class Meta:
        name = "constant.type"

    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class OperatorType(Child):
    class Meta:
        name = "operator.type"

    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class SepType(Child):
    class Meta:
        name = "sep.type"


class CnTypeType(Enum):
    INTEGER = "integer"
    RATIONAL = "rational"
    REAL = "real"


@dataclass(slots=True, kw_only=True)
class Abs(OperatorType):
    class Meta:
        name = "abs"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class And(OperatorType):
    class Meta:
        name = "and"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Arccos(OperatorType):
    class Meta:
        name = "arccos"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Arcsin(OperatorType):
    class Meta:
        name = "arcsin"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Arctan(OperatorType):
    class Meta:
        name = "arctan"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Ceiling(OperatorType):
    class Meta:
        name = "ceiling"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Ci(CiType):
    class Meta:
        name = "ci"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Cos(OperatorType):
    class Meta:
        name = "cos"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Divide(OperatorType):
    class Meta:
        name = "divide"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Eq(OperatorType):
    class Meta:
        name = "eq"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Exp(OperatorType):
    class Meta:
        name = "exp"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Exponentiale(ConstantType):
    class Meta:
        name = "exponentiale"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class FalseType(ConstantType):
    class Meta:
        name = "false"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Floor(OperatorType):
    class Meta:
        name = "floor"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Geq(OperatorType):
    class Meta:
        name = "geq"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Gt(OperatorType):
    class Meta:
        name = "gt"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Leq(OperatorType):
    class Meta:
        name = "leq"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Ln(OperatorType):
    class Meta:
        name = "ln"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Log(OperatorType):
    class Meta:
        name = "log"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Lt(OperatorType):
    class Meta:
        name = "lt"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Max(OperatorType):
    class Meta:
        name = "max"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Min(OperatorType):
    class Meta:
        name = "min"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Minus(OperatorType):
    class Meta:
        name = "minus"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Neq(OperatorType):
    class Meta:
        name = "neq"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Not(OperatorType):
    class Meta:
        name = "not"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Or(OperatorType):
    class Meta:
        name = "or"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Pi(ConstantType):
    class Meta:
        name = "pi"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Plus(OperatorType):
    class Meta:
        name = "plus"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Power(OperatorType):
    class Meta:
        name = "power"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Product(OperatorType):
    class Meta:
        name = "product"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Quotient(OperatorType):
    class Meta:
        name = "quotient"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Rem(OperatorType):
    class Meta:
        name = "rem"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Root(OperatorType):
    class Meta:
        name = "root"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Sep(SepType):
    class Meta:
        name = "sep"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Sin(OperatorType):
    class Meta:
        name = "sin"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Sum(OperatorType):
    class Meta:
        name = "sum"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Tan(OperatorType):
    class Meta:
        name = "tan"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Times(OperatorType):
    class Meta:
        name = "times"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class TrueType(ConstantType):
    class Meta:
        name = "true"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Xor(OperatorType):
    class Meta:
        name = "xor"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class CnType(Child):
    class Meta:
        name = "cn.type"

    type_value: None | CnTypeType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )
    content: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
            "mixed": True,
            "choices": (
                {
                    "name": "sep",
                    "type": ForwardRef("Sep"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Cn(CnType):
    class Meta:
        name = "cn"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class QualifierType(Child):
    class Meta:
        name = "qualifier.type"

    content: (
        None | Cn | Ci | Exponentiale | Pi | TrueType | FalseType | Apply | Logbase | Degree
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cn",
                    "type": ForwardRef("Cn"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "ci",
                    "type": ForwardRef("Ci"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "exponentiale",
                    "type": ForwardRef("Exponentiale"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "pi",
                    "type": ForwardRef("Pi"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "true",
                    "type": ForwardRef("TrueType"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "false",
                    "type": ForwardRef("FalseType"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "apply",
                    "type": ForwardRef("Apply"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "logbase",
                    "type": ForwardRef("Logbase"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "degree",
                    "type": ForwardRef("Degree"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
            ),
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
class Degree(QualifierType):
    class Meta:
        name = "degree"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class Logbase(QualifierType):
    class Meta:
        name = "logbase"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class ApplyType(Child):
    class Meta:
        name = "apply.type"

    content: list[
        Abs
        | Floor
        | Ceiling
        | Quotient
        | Divide
        | Rem
        | Minus
        | Plus
        | Times
        | Power
        | Root
        | Max
        | Min
        | And
        | Or
        | Xor
        | Not
        | Exponentiale
        | Pi
        | TrueType
        | FalseType
        | Eq
        | Neq
        | Leq
        | Lt
        | Geq
        | Gt
        | Exp
        | Ln
        | Log
        | Sin
        | Cos
        | Tan
        | Arcsin
        | Arccos
        | Arctan
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "abs",
                    "type": ForwardRef("Abs"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "floor",
                    "type": ForwardRef("Floor"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "ceiling",
                    "type": ForwardRef("Ceiling"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "quotient",
                    "type": ForwardRef("Quotient"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "divide",
                    "type": ForwardRef("Divide"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "rem",
                    "type": ForwardRef("Rem"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "minus",
                    "type": ForwardRef("Minus"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "plus",
                    "type": ForwardRef("Plus"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "times",
                    "type": ForwardRef("Times"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "power",
                    "type": ForwardRef("Power"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "root",
                    "type": ForwardRef("Root"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "max",
                    "type": ForwardRef("Max"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "min",
                    "type": ForwardRef("Min"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "and",
                    "type": ForwardRef("And"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "or",
                    "type": ForwardRef("Or"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "xor",
                    "type": ForwardRef("Xor"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "not",
                    "type": ForwardRef("Not"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "exponentiale",
                    "type": ForwardRef("Exponentiale"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "pi",
                    "type": ForwardRef("Pi"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "true",
                    "type": ForwardRef("TrueType"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "false",
                    "type": ForwardRef("FalseType"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "eq",
                    "type": ForwardRef("Eq"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "neq",
                    "type": ForwardRef("Neq"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "leq",
                    "type": ForwardRef("Leq"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "lt",
                    "type": ForwardRef("Lt"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "geq",
                    "type": ForwardRef("Geq"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "gt",
                    "type": ForwardRef("Gt"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "exp",
                    "type": ForwardRef("Exp"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "ln",
                    "type": ForwardRef("Ln"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "log",
                    "type": ForwardRef("Log"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "sin",
                    "type": ForwardRef("Sin"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "cos",
                    "type": ForwardRef("Cos"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "tan",
                    "type": ForwardRef("Tan"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "arcsin",
                    "type": ForwardRef("Arcsin"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "arccos",
                    "type": ForwardRef("Arccos"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "arctan",
                    "type": ForwardRef("Arctan"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
            ),
        },
    )
    content_1: list[Cn | Ci | Apply | Logbase | Degree] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cn",
                    "type": ForwardRef("Cn"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "ci",
                    "type": ForwardRef("Ci"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "apply",
                    "type": ForwardRef("Apply"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "logbase",
                    "type": ForwardRef("Logbase"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "degree",
                    "type": ForwardRef("Degree"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
            ),
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
class Apply(ApplyType):
    class Meta:
        name = "apply"
        namespace = "http://www.w3.org/1998/Math/MathML"


@dataclass(slots=True, kw_only=True)
class MathType(Child):
    class Meta:
        name = "math.type"

    content: (
        None | Cn | Ci | Exponentiale | Pi | TrueType | FalseType | Apply | Logbase | Degree
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cn",
                    "type": ForwardRef("Cn"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "ci",
                    "type": ForwardRef("Ci"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "exponentiale",
                    "type": ForwardRef("Exponentiale"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "pi",
                    "type": ForwardRef("Pi"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "true",
                    "type": ForwardRef("TrueType"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "false",
                    "type": ForwardRef("FalseType"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "apply",
                    "type": ForwardRef("Apply"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "logbase",
                    "type": ForwardRef("Logbase"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "degree",
                    "type": ForwardRef("Degree"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
            ),
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
class Math(MathType):
    class Meta:
        name = "math"
        namespace = "http://www.w3.org/1998/Math/MathML"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.mathml import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.mathml.el",
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
