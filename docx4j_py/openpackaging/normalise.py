"""What a save writes: Word's reading of a value the parse kept as its text, written plainly. CR-018 section 4.7.

The user's direction (2026-10-24): "a normaliser would be good. A user who edits
a file then blames us is a real concern: they are entitled to assume we don't
write dodgy content. I'd be inclined to make the normaliser opt-out." So a save
writes every value CR-018 covers in a **normal form that Word reads exactly as it
read the source**: a save never changes Word's meaning, and never writes what
Word would repair or misread.

**Phase A's rows** (the on/off values), for an ST_OnOff attribute the parse
kept as its text (the fork's ``exact_types``, :data:`docx4j_py.runtime.EXACT_TYPES`):

* a value Word reads as **on** (``on``, ``""``, ``" false"``, ``TRUE``,
  ``yes``, ...) is written **bare**, the attribute dropped, where the schema's
  default is on (an element's ``w:val``, ``BooleanDefaultTrue``); elsewhere
  (``w:fldLock``, ``w14:val``) as ``1``;
* a value Word reads as **off** (only ``off`` can be kept: ``false`` and ``0``
  bind as ``False``) is written ``0``, the element kept: an off value can
  override a style, and the save looks nothing up.

**What it does not touch.** A value that bound as it is (``true``, ``false``,
``1``, ``0``) is the serialiser's, written ``1`` or ``0`` as today, so "bare"
is the form of a rewritten value only. DrawingML's and docProps' booleans are
``xsd:boolean``, not Word's ST_OnOff, and keep their spelling. An on/off value
set in code keeps today's form; a colour or an enumeration set in code is
written in the normal form like any other (Phase B's rows read the render).

**Phase B's rows** (the enumerations and the colours, CR-018 sections 4.2,
4.3 and 4.7; docx4j's CR-060 D3 and CR-061), found by name in each part's
**render** (:func:`normalise_render`), since most of what they rewrite binds as
it is (a lower-case hex, a bare ``<w:highlight/>``) and so is not in
``part.kept``:

* ``w:color/@w:val`` and ``w:u/@w:color``: a value Word reads (a classic HTML
  name, hex in any case, ``#``, padding) is written ``RRGGBB``; one it rejects
  is dropped, with its ``w:color`` where nothing else is on it, since Word
  reads it as absent; beside a ``w:themeColor`` a rejected ``w:val`` is written
  ``auto``, as ``w:val`` is required;
* ``w:u/@w:val`` outside the enumeration is dropped, and a bare ``<w:u/>`` with
  it, as Word's save drops it; ``w:u``'s six on/off words are written ``none``;
* a ``w:highlight``, ``w:vertAlign`` or ``w:em`` whose value is missing or
  outside its enumeration is dropped, as Word's repair drops it;
* ``w:shd``'s fill and pattern colour, and a border's colour: ``RRGGBB``, and
  ``auto`` for a value Word rejects, which Word reads as an automatic that
  overrides a style's;
* a ``w:background`` whose colour Word rejects is dropped, as Word's save
  drops it (round seven): no page colour.

**CR-021's rows** (integers, sections 4.3, 4.5 and 4.6; docx4j CR-065 D3 (a)), in
the same typed walk as the on/off rows:

* an integer attribute of ``w``, ``w14`` or ``w15`` (:data:`INTEGER_NAMESPACES`)
  that the parse kept as its text (a decimal, ``""``) is written as Word reads
  it, :func:`docx4j_py.model.values.integer`, a decimal integer: ``"720.5"`` as
  ``720``, ``"-121.5"`` as ``-121``, ``""`` as ``0``; a form Word refuses
  (``"abc"``, ``"1e3"``, ``" 240 "``) is dropped, as Word refuses the document;
* elsewhere (DrawingML, VML, charts: CR-021 D5) a kept value is written as
  Python's ``int()`` binds it where it can, and as read where it cannot, so
  that the output there is what the parse before the fork's stage 14 gave; in
  the union fields that hold an integer (:data:`UNION_INTEGER_ATTRIBUTES`, none
  in ``w``), whose values bind as text without a record, the same rule is
  applied by name in the render;
* an rsid attribute (``w:rsidR``, ``w:rsidRPr``, any ``w:rsid*``) whose value
  is ``""`` is dropped, as Word's save drops it (D6): an empty ``hexBinary`` is
  lexically valid, so it is found by value, in the parts whose source holds an
  empty value (``XmlPart._empty_value``) as well as those that kept one.

A value Word reads as it is (``"240"``) is the serialiser's and is not touched;
neither is a value equal to its style's, which Word's own save leaves out (D7).

**Scope and cost.** Only parts the save re-marshals (a part never unmarshalled
is still written byte for byte). The on/off rows walk only the parts whose
parse kept a value (:attr:`XmlPart.kept`), so a save with nothing odd in it
costs one list test per part; such a part's tree is walked once, so that a kept
value an edit copied (a run split, ``w:rPrChange``'s old properties, which
Reject All restores) is written in the normal form too. The enumeration and
colour rows cost one ``iter`` over about twenty tags of each render, a third of
a millisecond over ``Symbols.docx``'s 790 KB main part.

**The tree keeps what was read.** The on/off rows are written into the tree for
the save and taken back out after it, and the enumeration and colour rows into
the render only, as docx4j's marshal-only adapters leave the object in memory
as it was (its CR-060 D3, CR-061): a later ``save(normalise=False)`` still
writes the source's spelling, and ``part.xml`` outside a save writes the tree as
it stands.

**Not silent.** Each rewrite is a :class:`Normalised`, in the save's
:class:`SaveReport` (``package.last_save``). ``save(normalise=False)``, or
``LoadOptions(normalise_on_save=False)`` for a package, keeps the source's
spelling, for a tool that must change nothing it was not asked to.
"""

from __future__ import annotations

import contextlib
import contextvars
import dataclasses
import threading
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from docx4j_py.model.values import integer, on_off_normal_form

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

__all__ = [
    "INTEGER_NAMESPACES",
    "ON_OFF_NAMESPACES",
    "UNION_INTEGER_ATTRIBUTES",
    "Normalised",
    "SaveReport",
    "integer_attributes",
    "normalise_render",
    "normalise_tree",
    "normalised_package",
    "normalised_tree",
    "on_off_attributes",
    "rendered_rows",
    "rsid_attributes",
]

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

#: The namespaces whose ``bool`` attributes are Word's ST_OnOff (docx4j's ``wml.xsd``
#: types ST_OnOff as ``xsd:boolean``; ``w14`` and ``w15`` carry docx4j's own patch of it).
ON_OFF_NAMESPACES: frozenset[str] = frozenset({W, W14, W15})

#: The namespaces whose integer attributes are read by Word's rule (CR-021 D5, Q3): where
#: Word was measured (docx4j CR-065 phase 0). Elsewhere a kept integer keeps today's output.
INTEGER_NAMESPACES: frozenset[str] = frozenset({W, W14, W15})

#: The union fields that hold an integer (CR-021 D5; none is in ``w``, ``w14`` or ``w15``), by
#: element and attribute, as the model declares them (``tests/content/test_integer_values.py``
#: derives the table from it): a value Python's ``int()`` takes is written as that integer, as
#: the parse before the fork's stage 14 bound it. Found by name in the render, since such a
#: value binds as the ``str`` member without a record.
UNION_INTEGER_ATTRIBUTES: dict[str, tuple[str, ...]] = {
    "{http://schemas.microsoft.com/ink/2010/main}context": ("semanticType",),
    "{http://schemas.microsoft.com/office/drawing/2008/diagram}sp": ("modelId",),
    "{http://schemas.microsoft.com/office/drawing/2012/chartStyle}styleClr": ("val",),
    "{http://schemas.microsoft.com/office/drawing/2018/animation/model3d}animPr": ("count",),
    "{http://schemas.openxmlformats.org/drawingml/2006/diagram}cxn": (
        "modelId", "srcId", "destId", "parTransId", "sibTransId",
    ),
    "{http://schemas.openxmlformats.org/drawingml/2006/diagram}if": ("val",),
    "{http://schemas.openxmlformats.org/drawingml/2006/diagram}param": ("val",),
    "{http://schemas.openxmlformats.org/drawingml/2006/diagram}prSet": ("presAssocID",),
    "{http://schemas.openxmlformats.org/drawingml/2006/diagram}pt": ("modelId", "cxnId"),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}ahPolar": ("minR", "maxR", "minAng", "maxAng"),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}ahXY": ("minX", "maxX", "minY", "maxY"),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}arcTo": ("wR", "hR", "stAng", "swAng"),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}cxn": ("ang",),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}pos": ("x", "y"),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}pt": ("x", "y"),
    "{http://schemas.openxmlformats.org/drawingml/2006/main}rect": ("l", "t", "r", "b"),
    **{
        f"{{http://schemas.openxmlformats.org/officeDocument/2006/math}}{name}": (
            "{http://schemas.openxmlformats.org/officeDocument/2006/math}val",
        )
        for name in ("interSp", "intraSp", "lMargin", "postSp", "preSp", "rMargin", "wrapIndent")
    },
}

#: The namespace whose defaults have been audited (docx4j #697, CR-059; CR-018 Q10): only
#: there does a save drop an on value whose schema default is on.
_AUDITED = W


@dataclasses.dataclass(frozen=True, slots=True)
class Normalised:
    """One value a save wrote in the normal form.

    Attributes:
        part: the part's name, ``/word/document.xml``.
        path: where it is: the element's path from the root, each step a
            prefixed name with its 1-based position among the siblings of that
            name, then ``@`` and the attribute (``w:document/w:body/w:p[3]/w:r[2]/w:rPr/w:b/@w:val``);
            for an element dropped (Phase B's rows), the element's path alone.
        before: the value as read (and as the tree still holds it); for an
            element dropped, the element as read (``<w:highlight w:val="YELLOW"/>``).
        after: what the save wrote: ``"1"`` or ``"0"``, ``"FF0000"``,
            ``"auto"``, ``"none"``; or None for the attribute, or the element,
            dropped.
    """

    part: str
    path: str
    before: str
    after: str | None


@dataclasses.dataclass(slots=True)
class SaveReport:
    """What a save did beyond writing the tree: ``package.last_save``.

    Attributes:
        normalise: whether the normal form was written (``save(normalise=)``).
        normalised: each value written in it, in part and document order.
        added: the names of the parts the save's hooks added
            (:func:`~docx4j_py.openpackaging.save.run_save_hooks`; CR-019
            section 4.9: a glossary a placeholder needed), in the order added.
        checked: whether the save ran the validator first (``save(check=)``,
            CR-020 Phase B).
        findings: what that check found that the source already held, and so
            saved as it is (:func:`docx4j_py.validator.check_for_save`): a table
            cell that does not end with a paragraph, an integer measurement
            past what Word stores. Empty where the check found nothing, or did
            not run.
    """

    normalise: bool = True
    normalised: list[Normalised] = dataclasses.field(default_factory=list)
    added: list[str] = dataclasses.field(default_factory=list)
    checked: bool = False
    findings: list[Any] = dataclasses.field(default_factory=list)

    @property
    def out_of_range(self) -> list[Any]:
        """The ``value.out_of_range`` findings: docx4j-core-ts's ``lastSave.outOfRange`` (its CR-017 P3)."""
        return [finding for finding in self.findings if finding.code == "value.out_of_range"]

    def by_part(self, order: list[str]) -> None:
        """Put :attr:`normalised` in the order the save wrote the parts, each part's rows in their own order."""
        rank = {part: index for index, part in enumerate(order)}
        self.normalised.sort(key=lambda record: rank.get(record.part, len(rank)))


# ---------------------------------------------------------------------------
# which attributes are on/off
# ---------------------------------------------------------------------------

_VARS: dict[type, tuple[tuple[Any, bool], ...]] = {}
_VARS_LOCK = threading.Lock()


def _namespace(qname: str | None) -> str | None:
    if not qname or not qname.startswith("{"):
        return None
    return qname[1:].split("}", 1)[0]


def on_off_attributes(cls: type) -> tuple[tuple[Any, bool], ...]:
    """A class's ST_OnOff attributes: (its ``XmlVar``, whether absence is on and audited), cached per class.

    An attribute is ST_OnOff when its one type is ``bool`` and its name is in
    :data:`ON_OFF_NAMESPACES`. Absence is on where the attribute is ``w``'s and
    the schema default, typed (the fork's ``XmlVar.schema_default_value``, stage
    13 item 31), is ``True``: an element's ``w:val`` (``BooleanDefaultTrue``),
    ``w:cols``' ``w:equalWidth``.
    """
    found = _VARS.get(cls)
    if found is not None:
        return found
    from docx4j_py.runtime import context

    out = []
    for var in context().build(cls).get_attribute_vars():
        if tuple(var.types) != (bool,):
            continue
        namespace = _namespace(var.qname)
        if namespace not in ON_OFF_NAMESPACES:
            continue
        out.append((var, namespace == _AUDITED and getattr(var, "schema_default_value", None) is True))
    found = tuple(out)
    with _VARS_LOCK:
        _VARS[cls] = found
    return found


_INTEGER_VARS: dict[type, tuple[tuple[Any, bool], ...]] = {}
_RSID_VARS: dict[type, tuple[Any, ...]] = {}


def integer_attributes(cls: type) -> tuple[tuple[Any, bool], ...]:
    """A class's integer attributes: (its ``XmlVar``, whether Word's rule applies), cached per class.

    An attribute is an integer one when its one type is ``int`` (a token list
    of them too, DrawingML's ``dgm:forEach/@st``). Word's rule
    (:func:`~docx4j_py.model.values.integer`) applies where the attribute is in
    :data:`INTEGER_NAMESPACES` (CR-021 D5); elsewhere a kept value is written as
    Python's ``int()`` binds it, where it can.
    """
    found = _INTEGER_VARS.get(cls)
    if found is not None:
        return found
    from docx4j_py.runtime import context

    found = tuple(
        (var, _namespace(var.qname) in INTEGER_NAMESPACES)
        for var in context().build(cls).get_attribute_vars()
        if tuple(var.types) == (int,)
    )
    with _VARS_LOCK:
        _INTEGER_VARS[cls] = found
    return found


def rsid_attributes(cls: type) -> tuple[Any, ...]:
    """A class's rsid attributes (``w:rsidR``, ``w:rsidRPr``, ``w:rsidDel``, any ``w:rsid*``), cached per class (D6)."""
    found = _RSID_VARS.get(cls)
    if found is not None:
        return found
    from docx4j_py.runtime import context

    found = tuple(
        var
        for var in context().build(cls).get_attribute_vars()
        if _namespace(var.qname) == W and var.qname.split("}", 1)[1].startswith("rsid")
    )
    with _VARS_LOCK:
        _RSID_VARS[cls] = found
    return found


def _python_int(text: str, tokens: bool = False) -> str | None:
    """What Python's ``int()`` binds `text` as, written; a token list token by token; None where it cannot (D5)."""
    try:
        if tokens:
            return " ".join(str(int(token)) for token in text.split())
        return str(int(text))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# the typed tree
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class _Change:
    node: Any
    field: str
    qname: str
    before: str
    #: what the field holds for the save: a ``bool`` (the on/off rows), an ``int`` or a token list's text
    #: (the integer rows), or None for the attribute dropped
    after: Any
    #: what the save writes, for the record: ``"1"``, ``"0"``, ``"720"``; None for the attribute dropped
    written: str | None


#: ``xsd:boolean``'s four literals: a value spelt so binds as a ``bool`` and is the serialiser's.
_EXACT = frozenset({"true", "false", "1", "0"})


def _changes(root: Any) -> list[_Change]:
    """Every kept ST_OnOff and integer value, and every empty rsid, under `root`, every ``mc`` branch included.

    Each with its normal form: the on/off rows (CR-018 section 4.7), the integer
    rows (CR-021 sections 4.3 and 4.5) and the empty rsid (section 4.6).
    """
    from docx4j_py.child import is_any_element
    from docx4j_py.runtime import context
    from docx4j_py.traversal import iter_nodes

    out: list[_Change] = []
    for node in iter_nodes(root, context=context(), mce="all"):
        if is_any_element(node) or not dataclasses.is_dataclass(node):
            continue
        cls = type(node)
        for var, absent_is_on in on_off_attributes(cls):
            value = getattr(node, var.name, None)
            # an exact literal as a str is code's (``checked.val = "1"``), and Word reads it as written
            if isinstance(value, str) and value not in _EXACT:
                after = on_off_normal_form(value, absent_is_on)
                out.append(_Change(node, var.name, var.qname, value, after, _written(after)))
        for var, word in integer_attributes(cls):
            value = getattr(node, var.name, None)
            if not isinstance(value, str):
                continue  # bound as an int (or absent): the serialiser's, as before
            if word:
                read = integer(value)
                written = None if read is None else str(read)
                if written != value:  # "240" as a str is code's, and Word reads it as written
                    out.append(_Change(node, var.name, var.qname, value, read, written))
                continue
            bound = _python_int(value, bool(var.tokens))
            if bound is not None and bound != value:
                after: Any = bound if var.tokens else int(bound)
                out.append(_Change(node, var.name, var.qname, value, after, bound))
        for var in rsid_attributes(cls):
            if getattr(node, var.name, None) == "":
                out.append(_Change(node, var.name, var.qname, "", None, None))
    return out


def _prefixed(qname: str) -> str:
    from docx4j_py.namespaces import NAMESPACE_TO_PREFIX

    namespace = _namespace(qname)
    if namespace is None:
        return qname
    local = qname.split("}", 1)[1]
    prefix = NAMESPACE_TO_PREFIX.get(namespace)
    return f"{prefix}:{local}" if prefix else qname


def _path(node: Any, attribute: str) -> str:
    """`node`'s path from its root, positions counted among same-named siblings, then the attribute."""
    from docx4j_py.child import iter_children
    from docx4j_py.runtime import context
    from docx4j_py.traversal import element_name

    steps: list[str] = []
    current = node
    while current is not None:
        parent = getattr(current, "parent", None)
        if parent is None:
            steps.append(_prefixed(element_name(current) or type(current).__name__))
            break
        siblings = list(iter_children(parent, context=context(), mce="all"))
        # the name is the field's (w:b is a BooleanDefaultTrue, with no name of its own)
        name = next((qname for qname, child in siblings if child is current), None) or element_name(current) or ""
        same = [child for qname, child in siblings if qname == name]
        position = next((i for i, child in enumerate(same, 1) if child is current), 0)
        steps.append(f"{_prefixed(name)}[{position}]" if len(same) > 1 else _prefixed(name))
        current = parent
    return "/".join(reversed(steps)) + f"/@{_prefixed(attribute)}"


def _written(value: bool | None) -> str | None:
    return None if value is None else ("1" if value else "0")


@contextlib.contextmanager
def _applied(changes: list[_Change]) -> Iterator[None]:
    try:
        for change in changes:
            setattr(change.node, change.field, change.after)
        yield
    finally:
        for change in reversed(changes):
            setattr(change.node, change.field, change.before)


@contextlib.contextmanager
def normalised_tree(root: Any, *, part: str = "") -> Iterator[list[Normalised]]:
    """Write the normal form into one typed tree for the block, and take it back out after it.

    For a caller that serialises a tree itself (``scripts/roundtrip.py``); a
    save uses :func:`normalised_package`. The list is filled before the block
    runs.
    """
    changes = _changes(root)
    records = [Normalised(part, _path(c.node, c.qname), c.before, c.written) for c in changes]
    with _applied(changes):
        yield records


@contextlib.contextmanager
def normalised_package(package: OpcPackage, *, enabled: bool = True) -> Iterator[SaveReport]:
    """Write the normal form into every part the save re-marshals, for the save; take it back out after.

    The parts are those holding a live tree (:attr:`XmlPart.contents` read)
    whose parse kept a value (:attr:`XmlPart.kept`) or whose source holds an
    empty attribute value (``XmlPart._empty_value``: an empty rsid, CR-021
    D6, binds without a record). The report is filled before the block runs and
    stays the save's.
    """
    report = SaveReport(normalise=enabled)
    applied: list[_Change] = []
    active = _Active(report) if enabled else None
    if enabled:
        from docx4j_py.openpackaging.parts.xml_part import XmlPart

        for part in package.parts.parts():
            if not isinstance(part, XmlPart) or not (part.kept or getattr(part, "_empty_value", False)):
                continue
            tree = part._contents  # the live tree only: a private one is not saved
            if tree is None:
                continue
            for change in _changes(tree):
                report.normalised.append(
                    Normalised(str(part.part_name), _path(change.node, change.qname), change.before, change.written)
                )
                applied.append(change)
    token = _ACTIVE.set(active)
    try:
        with _applied(applied):
            yield report
    finally:
        _ACTIVE.reset(token)
    if active is not None:
        report.by_part(active.order)


# ---------------------------------------------------------------------------
# an lxml tree: the same normal form over a source, for canon.py
# ---------------------------------------------------------------------------

_LXML_TABLE: dict[tuple[str, str], bool] | None = None
_LXML_INTEGERS: tuple[dict[tuple[str, str], str], dict[tuple[str, str, str], str]] | None = None
_LXML_BOUND: tuple[dict[str, set[type]], dict[tuple[str, type], set[type]], set[type]] | None = None
_LXML_LOCK = threading.Lock()


def _bound() -> tuple[dict[str, set[type]], dict[tuple[str, type], set[type]], set[type]]:
    """Every class each element name binds (a field of another class, or a root), built once from the model.

    Also the classes each (element name, class) binding is a field of, and the
    classes whose metadata cannot be built. A root is a class whose ``Meta``
    declares a ``namespace`` as well as a ``name`` (a global element), as
    :func:`~docx4j_py.traversal.element_name` reads it, not ``XmlMeta.qname``,
    which xsdata builds for a type class against whatever parent namespace it
    met first (``CT_PPr`` before a parse, ``{...}CT_PPr`` after one), so that
    the tables are the same whatever ran before them.
    """
    global _LXML_BOUND
    if _LXML_BOUND is not None:
        return _LXML_BOUND
    import inspect
    from importlib import import_module

    from docx4j_py.el_index import EL_MODULES
    from docx4j_py.runtime import context

    bound: dict[str, set[type]] = {}
    holders: dict[tuple[str, type], set[type]] = {}
    seen: set[type] = set()
    broken: set[type] = set()
    for module_name in sorted(set(EL_MODULES.values())):
        module = import_module(module_name.removesuffix(".el"))
        for cls in vars(module).values():
            if not (inspect.isclass(cls) and dataclasses.is_dataclass(cls)) or cls in seen:
                continue
            seen.add(cls)
            try:
                meta = context().build(cls)
            except Exception:  # noqa: BLE001 - inkml's MappingType names MathML's Math, which no module defines
                broken.add(cls)
                continue
            declared = vars(cls).get("Meta")
            name, namespace = getattr(declared, "name", None), getattr(declared, "namespace", None)
            if name and namespace:
                bound.setdefault(f"{{{namespace}}}{name}", set()).add(cls)
            for var in meta.get_element_vars():
                for each in (var, *(var.elements or {}).values()):
                    for tp in each.types:
                        if dataclasses.is_dataclass(tp):
                            bound.setdefault(each.qname, set()).add(tp)
                            holders.setdefault((each.qname, tp), set()).add(cls)
    with _LXML_LOCK:
        _LXML_BOUND = (bound, holders, broken)
    return _LXML_BOUND


def _lxml_table() -> dict[tuple[str, str], bool]:
    """(element qname, attribute qname) to "absence is on", for every ST_OnOff attribute the model types.

    Built once from the model: every class that an element name binds, and that
    class's :func:`on_off_attributes`. An attribute that one binding of an
    element name types ``bool`` and another does not is left out: the typed
    parse would keep it in one place and not the other, so the lxml side cannot
    say.
    """
    global _LXML_TABLE
    if _LXML_TABLE is not None:
        return _LXML_TABLE
    bound, _holders, broken = _bound()
    table: dict[tuple[str, str], bool] = {}
    for element, classes in bound.items():
        per_class = [{var.qname: on for var, on in on_off_attributes(cls)} for cls in classes if cls not in broken]
        if not per_class:
            continue
        names = set().union(*per_class)
        for attribute in names:
            values = {found.get(attribute) for found in per_class if attribute in found}
            everywhere = all(attribute in found for found in per_class)
            if everywhere and len(values) == 1:
                table[(element, attribute)] = values.pop()
    with _LXML_LOCK:
        _LXML_TABLE = table
    return table


def _integer_kind(var: Any, word: bool) -> str:
    return "word" if word else ("tokens" if var.tokens else "python")


def _integer_tables() -> tuple[dict[tuple[str, str], str], dict[tuple[str, str, str], str]]:
    """The integer attributes the model types, by name, for the lxml side: CR-021 Q5.

    The first table maps (element qname, attribute qname) to the attribute's
    kind --- ``"word"`` (Word's rule, :data:`INTEGER_NAMESPACES`), ``"python"``
    or ``"tokens"`` (D5's) --- where **every class that declares the attribute**
    among those the element name binds types it ``int``, alike (Q5: classes
    that do not declare it, ``w:bottom`` in ``w:tblCellMar`` for a border's
    ``w:sz``, do not count). That leaves three pairs whose element name binds a
    class typing the attribute otherwise (``w:sz@w:val``: ``w:frame``'s is a
    ``str``; ``w:default@w:val``; ``w:start@w:val``: a border edge's is
    ``ST_Border``). The second table settles those by the element's holder,
    (holder qname, element qname, attribute qname), where every class bound
    under that holder agrees: ``w:rPr/w:sz``, ``w:ddList/w:default``,
    ``w:lvl/w:start``.
    """
    global _LXML_INTEGERS
    if _LXML_INTEGERS is not None:
        return _LXML_INTEGERS
    from docx4j_py.runtime import context

    bound, holders, broken = _bound()
    names_of: dict[type, set[str]] = {}
    for element, classes in bound.items():
        for cls in classes:
            names_of.setdefault(cls, set()).add(element)

    def kinds_of(classes: set[type]) -> dict[str, set[str | None]]:
        """Each attribute some class declares, and the kinds the declaring classes give it (None: not an int)."""
        out: dict[str, set[str | None]] = {}
        for cls in classes:
            if cls in broken:
                continue
            integers = {var.qname: _integer_kind(var, word) for var, word in integer_attributes(cls)}
            for var in context().build(cls).get_attribute_vars():
                out.setdefault(var.qname, set()).add(integers.get(var.qname))
        return out

    table: dict[tuple[str, str], str] = {}
    by_holder: dict[tuple[str, str, str], str] = {}
    for element, classes in bound.items():
        if not element.startswith("{"):
            continue  # a class's own type name (``CT_Bevel``), which no element in a part has
        for attribute, kinds in kinds_of(classes).items():
            if kinds == {None}:
                continue
            if len(kinds) == 1:
                table[(element, attribute)] = next(iter(kinds))  # type: ignore[assignment]
                continue
            # the element name binds classes that disagree: by the holder each is bound under
            under: dict[str, set[type]] = {}
            for cls in classes:
                for holder in holders.get((element, cls), ()):
                    for name in names_of.get(holder, ()):
                        under.setdefault(name, set()).add(cls)
            for holder_name, there in under.items():
                if not holder_name.startswith("{"):
                    continue
                found = kinds_of(there).get(attribute)
                if found is not None and len(found) == 1 and None not in found:
                    by_holder[(holder_name, element, attribute)] = next(iter(found))  # type: ignore[arg-type]
    with _LXML_LOCK:
        _LXML_INTEGERS = (table, by_holder)
    return _LXML_INTEGERS


def _integer_normal_form(value: str, kind: str, exact: bool) -> str | None:
    """What a save writes for an integer attribute's `value`, by its kind: CR-021 sections 4.3 and 4.5.

    ``"word"``: Word's reading (:func:`~docx4j_py.model.values.integer`), or
    None for a form Word refuses; where the parse is not exact for ``int``
    (before the fork's stage 14) a value Python's ``int()`` takes binds and the
    serialiser writes that integer, so it is what the save writes. ``"python"``
    and ``"tokens"`` (D5): Python's ``int()`` where it takes the value, else the
    value as read.
    """
    if kind == "word":
        if not exact:
            bound = _python_int(value)
            if bound is not None:
                return bound
        read = integer(value)
        return None if read is None else str(read)
    bound = _python_int(value, kind == "tokens")
    return value if bound is None else bound


def _exact_integers() -> bool:
    """Whether the engine's parse reads an integer exactly (the fork's stage 14, CR-021 phase 0)."""
    from docx4j_py.runtime import EXACT_TYPES

    return int in EXACT_TYPES


def normalise_tree(root: Any) -> list[tuple[str, str | None, str, str | None]]:
    """Write the normal form into an lxml tree of a part, in place: (element path, attribute, before, after).

    The same rows as a save's, found by name rather than by the parse: the
    on/off rows (:func:`docx4j_py.model.values.on_off_normal_form`; an
    ST_OnOff attribute whose value is not exactly ``true``, ``false``, ``1`` or
    ``0``); the integer rows (:func:`_integer_tables`, CR-021: any value the
    save would write otherwise, ``"720.5"`` as ``720``, a value the
    serialiser re-spells too, ``"+240"`` as ``240``); the empty rsid, dropped;
    then the rows a save applies to its render, the very functions
    (:func:`normalise_render`: the enumerations and colours, attribute None for
    an element dropped, and the union integers). ``scripts/canon.py``'s
    ``word-normal-form`` applies it to a source before comparing a normalised
    save with it, so that nothing else may move.
    """
    table = _lxml_table()
    integers, by_holder = _integer_tables()
    exact_integers = _exact_integers()
    exact = _EXACT
    rsid = _W + "rsid"
    out: list[tuple[str, str, str, str | None]] = []
    for element in root.iter():
        tag = element.tag
        if not isinstance(tag, str):
            continue
        holder: str | None = None
        for attribute, value in list(element.attrib.items()):
            absent_is_on = table.get((tag, attribute))
            if absent_is_on is not None:
                if value in exact:
                    continue
                after = _written(on_off_normal_form(value, absent_is_on))
            elif attribute.startswith(rsid) and value == "":
                after = None
            else:
                kind = integers.get((tag, attribute))
                if kind is None:
                    if holder is None:
                        parent = element.getparent()
                        holder = parent.tag if parent is not None else ""
                    kind = by_holder.get((holder, tag, attribute))
                    if kind is None:
                        continue
                after = _integer_normal_form(value, kind, exact_integers)
                if after == value:
                    continue
            if after is None:
                del element.attrib[attribute]
            else:
                element.set(attribute, after)
            out.append((element.getroottree().getpath(element), attribute, value, after))
    rows = [*_enumeration_and_colour_rows(root), *_union_rows(root)]
    out.extend((row.element.getroottree().getpath(row.element), row.attribute, row.before, row.after) for row in rows)
    _apply_rows(rows)
    return out


# ---------------------------------------------------------------------------
# Phase B: the enumerations and the colours, over an lxml tree
# ---------------------------------------------------------------------------

_W = "{" + W + "}"
_VAL = _W + "val"
_COLOR = _W + "color"
_FILL = _W + "fill"
_THEME = (_W + "themeColor", _W + "themeTint", _W + "themeShade")
_RPR = _W + "rPr"
#: The elements whose children are a border's edges (``CT_Border``): a paragraph's, a
#: table's, a cell's, a page's and an HTML frame's (``webSettings``).
_BORDER_SETS = frozenset(_W + name for name in ("pBdr", "tblBorders", "tcBorders", "pgBorders", "divBdr"))
_BORDER_EDGES = frozenset(
    _W + name
    for name in ("top", "left", "bottom", "right", "start", "end", "between", "bar", "insideH", "insideV", "tl2br", "tr2bl")
)
_BACKGROUND = _W + "background"
_BACKGROUND_HOLDERS = frozenset({_W + "document", _W + "glossaryDocument"})
#: The run properties Phase B reads, each found as a child of a ``w:rPr`` (a run's, a style's,
#: the defaults', a numbering level's, a paragraph mark's, a revision's former properties).
_RUN_ELEMENTS = frozenset(_W + name for name in ("color", "u", "highlight", "vertAlign", "em", "bdr"))
_TAGS = (*sorted(_RUN_ELEMENTS), _W + "shd", *sorted(_BORDER_EDGES), _BACKGROUND)


@dataclasses.dataclass(slots=True)
class _Row:
    """One rewrite of :func:`_enumeration_and_colour_rows`: an attribute rewritten or dropped, or an element dropped."""

    element: Any
    attribute: str | None  # None: the element itself is dropped
    before: str
    after: str | None


def _enumeration_readers() -> dict[str, Any]:
    from docx4j_py.model.values import emphasis_mark, highlight, vertical_alignment

    return {_W + "highlight": highlight, _W + "vertAlign": vertical_alignment, _W + "em": emphasis_mark}


def _fill_or_border(element: Any, attribute: str, rows: list[_Row]) -> None:
    """A fill's, a pattern's or a border's colour: one Word reads as ``RRGGBB``; one it rejects ``auto``."""
    from docx4j_py.model.values import AUTO, colour

    value = element.get(attribute)
    if value is None:
        return
    read = colour(value)
    after = AUTO if read is None else read
    if after != value:
        rows.append(_Row(element, attribute, value, after))


def _enumeration_and_colour_rows(root: Any) -> list[_Row]:
    """CR-018 section 4.7's enumeration and colour rows, as rewrites of an lxml tree (not yet applied).

    Found by name, in the ``w`` namespace only:

    * ``w:color`` in a ``w:rPr``: a value Word reads written ``RRGGBB`` (a name,
      lower-case hex, ``#``, padding); one it rejects (``Auto``, ``F00``,
      ``""``) dropped with its element, as Word's save drops it, or written
      ``auto`` beside a theme attribute, since ``w:val`` is required and the
      theme colour is what Word draws; a ``w:color`` with no ``w:val`` and no
      theme attribute dropped;
    * ``w:u`` in a ``w:rPr``: a ``w:val`` outside the enumeration dropped, one
      of the six on/off words written ``none``; a ``w:color`` Word reads written
      ``RRGGBB``, one it rejects dropped; the element dropped where nothing is
      left on it (a bare ``<w:u/>``, as Word's save drops it);
    * ``w:highlight``, ``w:vertAlign`` and ``w:em`` in a ``w:rPr`` whose value
      is missing or outside the enumeration: dropped, as Word's repair drops
      them (Q6; docx4j's D3);
    * ``w:shd``'s ``w:fill`` and ``w:color``, a border edge's ``w:color``
      (``w:pBdr``, ``w:tblBorders``, ``w:tcBorders``, ``w:pgBorders``,
      ``w:divBdr``, and a run's ``w:bdr``): ``RRGGBB``, or ``auto`` for a value
      Word rejects, which it reads as an automatic that overrides a style's;
    * ``w:background``'s ``w:color``: ``RRGGBB``; a value Word rejects drops
      the element, as Word's save does (round seven), or is written ``auto``
      where a theme attribute or content is on it.

    A value already in the form is left alone, ``auto`` included.
    """
    from docx4j_py.model.values import AUTO, colour, underline

    readers = _enumeration_readers()
    rows: list[_Row] = []
    for element in root.iter(*_TAGS):
        parent = element.getparent()
        if parent is None:
            continue
        tag, holder = element.tag, parent.tag
        if tag in _RUN_ELEMENTS:
            if holder != _RPR:
                continue
            if tag == _COLOR:
                value = element.get(_VAL)
                themed = any(element.get(name) is not None for name in _THEME)
                read = None if value is None else colour(value)
                if read is None:
                    if not themed:
                        rows.append(_Row(element, None, _describe(element), None))
                    elif value is not None:
                        rows.append(_Row(element, _VAL, value, AUTO))
                elif read != value:
                    rows.append(_Row(element, _VAL, value, read))
            elif tag == _W + "u":
                changes: list[_Row] = []
                value = element.get(_VAL)
                if value is not None:
                    member = underline(value)
                    if member is None:
                        changes.append(_Row(element, _VAL, value, None))
                    elif member.value != value:
                        changes.append(_Row(element, _VAL, value, member.value))
                tint = element.get(_COLOR)
                if tint is not None:
                    read = colour(tint)
                    if read is None:
                        changes.append(_Row(element, _COLOR, tint, None))
                    elif read != tint:
                        changes.append(_Row(element, _COLOR, tint, read))
                dropped = {row.attribute for row in changes if row.after is None}
                if not (set(element.attrib) - dropped):
                    rows.append(_Row(element, None, _describe(element), None))
                else:
                    rows.extend(changes)
            elif tag == _W + "bdr":
                _fill_or_border(element, _COLOR, rows)
            elif readers[tag](element.get(_VAL)) is None:
                rows.append(_Row(element, None, _describe(element), None))
        elif tag == _W + "shd":
            _fill_or_border(element, _FILL, rows)
            _fill_or_border(element, _COLOR, rows)
        elif tag in _BORDER_EDGES:
            if holder in _BORDER_SETS:
                _fill_or_border(element, _COLOR, rows)
        elif tag == _BACKGROUND and holder in _BACKGROUND_HOLDERS:
            value = element.get(_COLOR)
            if value is None:
                continue
            read = colour(value)
            if read is None:
                themed = any(element.get(name) is not None for name in _THEME)
                if themed or len(element) or (element.text or "").strip():
                    rows.append(_Row(element, _COLOR, value, AUTO))
                else:
                    rows.append(_Row(element, None, _describe(element), None))
            elif read != value:
                rows.append(_Row(element, _COLOR, value, read))
    return rows


_UNION_TAGS = tuple(sorted(UNION_INTEGER_ATTRIBUTES))


def _union_rows(root: Any) -> list[_Row]:
    """CR-021 D5's union rows, as rewrites of an lxml tree (not yet applied): found by name.

    In an attribute of :data:`UNION_INTEGER_ATTRIBUTES` a value Python's
    ``int()`` takes is written as that integer (``" 240"`` as ``240``), as the
    parse before the fork's stage 14 bound and wrote it; any other value (a
    guide name, a decimal) is left as it is. A value bound as an integer is
    the serialiser's already, so nothing is found for it.
    """
    rows: list[_Row] = []
    for element in root.iter(*_UNION_TAGS):
        for attribute in UNION_INTEGER_ATTRIBUTES[element.tag]:
            value = element.get(attribute)
            if value is None:
                continue
            bound = _python_int(value)
            if bound is not None and bound != value:
                rows.append(_Row(element, attribute, value, bound))
    return rows


def _apply_rows(rows: list[_Row]) -> None:
    """Write :func:`_enumeration_and_colour_rows`' rewrites into their tree, a dropped element's tail kept."""
    for row in rows:
        element = row.element
        if row.attribute is not None:
            if row.after is None:
                del element.attrib[row.attribute]
            else:
                element.set(row.attribute, row.after)
            continue
        parent = element.getparent()
        if parent is None:
            continue
        if element.tail:
            previous = element.getprevious()
            if previous is not None:
                previous.tail = (previous.tail or "") + element.tail
            else:
                parent.text = (parent.text or "") + element.tail
        parent.remove(element)


def _prefixed_tag(tag: str) -> str:
    return _prefixed(tag) if tag.startswith("{") else tag


def _describe(element: Any) -> str:
    """An element as a short tag, ``<w:highlight w:val="YELLOW"/>``: what a dropped element was."""
    attributes = "".join(f' {_prefixed_tag(name)}="{value}"' for name, value in element.attrib.items())
    return f"<{_prefixed_tag(element.tag)}{attributes}/>"


def _lxml_path(element: Any, attribute: str | None) -> str:
    """:func:`_path`'s form over an lxml tree: prefixed names, a position where a name repeats among siblings."""
    steps: list[str] = []
    current = element
    while current is not None:
        name = _prefixed_tag(current.tag)
        parent = current.getparent()
        if parent is None:
            steps.append(name)
            break
        same = [child for child in parent if child.tag == current.tag]
        steps.append(f"{name}[{same.index(current) + 1}]" if len(same) > 1 else name)
        current = parent
    path = "/".join(reversed(steps))
    return path if attribute is None else f"{path}/@{_prefixed_tag(attribute)}"


#: The save under way in this context, for :func:`normalise_render`: its report and the
#: parts in the order the save marshalled them. None outside a normalising save.
_ACTIVE: contextvars.ContextVar[_Active | None] = contextvars.ContextVar("docx4j_py_normalising_save", default=None)


@dataclasses.dataclass(slots=True)
class _Active:
    report: SaveReport
    order: list[str] = dataclasses.field(default_factory=list)


def normalise_render(root: Any, part: str = "") -> list[Normalised]:
    """Write the render's rows into a part's rendered lxml tree, if a normalising save is under way.

    :meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart._marshal` calls it on
    every render. Outside a save that normalises (:func:`normalised_package`) it
    does nothing, so ``part.xml`` is unchanged; inside one the rewrites go into
    the save's report. The rows are CR-018 Phase B's enumerations and colours
    and CR-021 D5's union integers, found by name over the render (an ``iter``
    of about twenty tags each, a third of a millisecond over ``Symbols.docx``'s
    790 KB main part for the first), so the typed tree is never touched and a
    value set in code is written in the form too.
    """
    active = _ACTIVE.get()
    if active is None:
        return []
    if part not in active.order:
        active.order.append(part)
    records = rendered_rows(root, part)
    active.report.normalised.extend(records)
    return records


def rendered_rows(root: Any, part: str = "") -> list[Normalised]:
    """Write the render's rows into an lxml tree, in place, and return them as :class:`Normalised`.

    The enumeration and colour rows (CR-018 Phase B), then the union integer
    rows (CR-021 D5). For a caller that renders a tree itself
    (``scripts/roundtrip.py``), after :func:`normalised_tree`'s typed rows.
    """
    rows = [*_enumeration_and_colour_rows(root), *_union_rows(root)]
    records = [Normalised(part, _lxml_path(row.element, row.attribute), row.before, row.after) for row in rows]
    _apply_rows(rows)
    return records
