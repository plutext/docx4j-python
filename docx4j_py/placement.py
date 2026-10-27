"""What each typed list may hold: the table, the rule and the messages (CR-020 Phase A).

The generated types already say what every list may hold: ``P.content`` is a
union of 35 element classes and ``R.content`` of 34, and neither includes
``Tbl``. A type checker catches ``r.content.append(Tbl())``; code that is only
run never did, and Word opened the file with "unreadable content". This module
is the cold half of the check :class:`docx4j_py.child.ChildList` makes as code
adds to a list; the hot half, one dictionary lookup per item, stays in
``child.py``.

**The table (D1).** For each (class, list field), read lazily from the class's
``XmlMeta`` through the shared ``XmlContext`` (:func:`docx4j_py.runtime.context`)
and cached: the classes of each element choice and each choice's element name
(:class:`Entry`). A field with a wildcard, or with no element choices, has no
table and is not checked (28 of the model's 537 list fields). Strings are
allowed only where the field is mixed, and a value of a simple type only where
one of the field's choices is of a simple type (a ``list[str]`` element field,
``c:separator`` in ``c:dLbl``).

**The rule (D2).** An item goes in when

1. its class is one of the field's element classes, exactly; or
2. its element name (:func:`docx4j_py.traversal.element_name`) is one of the
   field's choices, its class has the **shape** of that choice's class
   (:func:`shape_of`: the attributes and the fields with the element names
   they hold, and one level further down the same of each child's class), and
   the serialiser writes it under that name in this field (:func:`written_as`);
   or
3. it is a ``str`` and the field is mixed, or a simple value and a choice is
   simple.

A ``DerivedElement`` is judged by its qualified name and the value it wraps; an
``AnyElement`` goes only where there is a wildcard, so never into a checked
list. Anything else is refused with :class:`MisplacedContentError`
(``target.misplaced``), whose message names the holder's field and the item and
whose hint says where the item goes (D6, :func:`refusal`).

Rule 2 is what lets the content API's hyperlink write through: a ``PHyperlink``
put in a run-level control's content, whose own class for ``w:hyperlink`` is
``CtSdtContentRunHyperlink``, the same XML. Its third condition is the
implementation's (CR-020 section 10.3): the serialiser writes an item whose class
is not one of the field's under the class's own name, which for most of the
same-shape classes is not the element's (``MoveFrom1`` in ``P.content`` is
written ``<MoveFrom1>``), so the check accepts only what is saved as the element
the field takes.

Nothing here is built at import: the first check of a (class, field) builds
its table, and the first refusal that needs a derived hint builds the reverse
index over every generated class (about 0.3 s, once per process).
"""

from __future__ import annotations

import dataclasses
from typing import Any

__all__ = [
    "Entry",
    "census",
    "entry",
    "holders_of",
    "is_allowed",
    "model_classes",
    "property_holders_of",
    "refusal",
    "shape_of",
    "verdict",
    "verdicts",
    "written_as",
]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


W_P = _w("p")
W_R = _w("r")
W_T = _w("t")
W_TBL = _w("tbl")
W_TR = _w("tr")
W_TC = _w("tc")
W_SDT = _w("sdt")
W_SDT_CONTENT = _w("sdtContent")
W_CUSTOM_XML = _w("customXml")
W_SECT_PR = _w("sectPr")


class Entry:
    """What one (class, list field) may hold, as the class metadata says.

    Attributes:
        cls: the holder's class.
        field: the list field's name.
        var: the field's ``XmlVar``.
        compound: whether the field is a compound (``XmlType.ELEMENTS``) one,
            rather than a list of one element.
        classes: every model class of every choice.
        by_name: element qualified name -> the model classes of that choice.
        simple: whether a choice is of a simple type (a value, not a class).
        mixed: whether the field is mixed content.
        namespace: the holder's namespace, as the serialiser passes it down.
        verdicts: item class -> whether it goes in; the exact classes are
            filled in on construction, the rest as they are judged.
    """

    __slots__ = (
        "by_name",
        "classes",
        "cls",
        "compound",
        "field",
        "mixed",
        "namespace",
        "simple",
        "var",
        "verdicts",
    )

    def __init__(self, cls: type, field: str, var: Any, choices: list[tuple[str, Any]], namespace: str | None) -> None:
        """Read the choices; nothing is judged yet but the exact classes."""
        self.cls = cls
        self.field = field
        self.var = var
        self.compound = bool(var.is_elements)
        self.mixed = bool(var.mixed)
        self.namespace = namespace
        by_name: dict[str, tuple[type, ...]] = {}
        classes: set[type] = set()
        simple = False
        for qname, alternative in choices:
            models = tuple(tp for tp in alternative.types if _is_model(tp))
            if len(models) < len(alternative.types):
                simple = True
            by_name[qname] = by_name.get(qname, ()) + models
            classes.update(models)
        self.by_name = by_name
        self.classes = frozenset(classes)
        self.simple = simple
        self.verdicts: dict[Any, bool] = dict.fromkeys(classes, True)

    def __repr__(self) -> str:
        """``Entry(P.content, 35 names)``."""
        return f"Entry({self.cls.__name__}.{self.field}, {len(self.by_name)} names)"


# ---------------------------------------------------------------------------
# the table (D1)
# ---------------------------------------------------------------------------

_entries: dict[tuple[type, str], Entry | None] = {}


def _context() -> Any:
    from docx4j_py.runtime import context

    return context()


def _is_model(tp: Any) -> bool:
    return isinstance(tp, type) and dataclasses.is_dataclass(tp)


def _meta(cls: type) -> Any:
    """The class's ``XmlMeta`` in the shared context, or None for a class xsdata cannot build."""
    try:
        return _context().build(cls)
    except Exception:  # noqa: BLE001 - not a model class, or one whose annotations do not resolve
        return None


def _element_vars(meta: Any) -> list[Any]:
    """The element, compound, wildcard and text vars, one per field name."""
    seen: set[str] = set()
    out = []
    for var in meta.get_element_vars():
        if var.name not in seen:
            seen.add(var.name)
            out.append(var)
    return out


def _build(cls: type, field: str) -> Entry | None:
    meta = _meta(cls)
    if meta is None:
        return None
    var = next((v for v in _element_vars(meta) if v.name == field), None)
    if var is None or var.is_wildcard or var.wildcards:
        return None
    if var.is_elements:
        if not var.elements:
            return None
        choices = list(var.elements.items())
    elif var.is_element:
        choices = [(var.qname, var)]
    else:
        return None
    return Entry(cls, field, var, choices, getattr(meta, "namespace", None))


def entry(cls: type, field: str) -> Entry | None:
    """The table for one (class, list field), or None where the field is not checked.

    None for a wildcard field, a field with no element choices, a field that
    is not an element list (an attribute's tokens), and a class xsdata cannot
    build (a hand-written ``Child``).
    """
    key = (cls, field)
    try:
        return _entries[key]
    except KeyError:
        pass
    built = _build(cls, field)
    _entries[key] = built
    return built


def verdicts(cls: type, field: str) -> dict[Any, bool] | None:
    """The verdict cache ``ChildList`` consults per item, or None where the field is not checked."""
    found = entry(cls, field)
    return None if found is None else found.verdicts


# ---------------------------------------------------------------------------
# the shape (rule 2)
# ---------------------------------------------------------------------------

_WILDCARD = "\x00wildcard"
_shapes: dict[type, Any] = {}
_shapes_below: dict[type, Any] = {}


def _attributes(meta: Any) -> frozenset[str]:
    return frozenset(var.qname for var in meta.get_attribute_vars())


def _shape_below(cls: type) -> Any:
    """One level of a class's shape: its attributes and each field's element names."""
    try:
        return _shapes_below[cls]
    except KeyError:
        pass
    meta = _meta(cls)
    result: Any = None
    if meta is not None:
        fields = []
        for var in _element_vars(meta):
            if var.is_wildcard:
                fields.append((var.name, _WILDCARD))
            elif var.elements:
                fields.append((var.name, frozenset(var.elements)))
            else:
                fields.append((var.name, frozenset((var.qname,))))
        result = (_attributes(meta), frozenset(fields))
    _shapes_below[cls] = result
    return result


def shape_of(cls: type) -> Any:
    """A class's shape: what it writes, compared two levels deep.

    Its attributes, and each field with the element names it holds and, for
    each of them, the one-level shape of the choice's classes: the
    attributes and the fields with their element names. Two classes of one
    shape write interchangeable XML as far as these levels go; below them,
    the item's own lists are checked against its own class. None for a class
    that is not a model.

    Over the model, 71 element names in compound fields have more than one
    class; this calls the classes of 60 of them one shape, and of 11
    different: ``w:sdt`` and ``w:customXml`` (one per form), ``w:ins``,
    ``w:del`` and ``w:name``; ``a:alphaMod``, ``a:lum``, ``a:tint``,
    ``a:xfrm`` and ``a:effect``; chartex's ``cx:lvl`` (CR-020 section 2.3).
    """
    try:
        return _shapes[cls]
    except KeyError:
        pass
    meta = _meta(cls)
    result: Any = None
    if meta is not None:
        fields = []
        for var in _element_vars(meta):
            if var.is_wildcard:
                fields.append((var.name, _WILDCARD))
                continue
            alternatives = var.elements.items() if var.elements else ((var.qname, var),)
            fields.append(
                (
                    var.name,
                    frozenset(
                        (qname, frozenset(_shape_below(tp) for tp in alternative.types if _is_model(tp)))
                        for qname, alternative in alternatives
                    ),
                )
            )
        result = (_attributes(meta), frozenset(fields))
    _shapes[cls] = result
    return result


# ---------------------------------------------------------------------------
# what the serialiser writes (rule 2's third condition)
# ---------------------------------------------------------------------------

_branch_names_cache: dict[type, str] | None = None


def _branch_names() -> dict[type, str]:
    global _branch_names_cache
    if _branch_names_cache is None:
        from docx4j_py.runtime import _branch_element_names

        _branch_names_cache = dict(_branch_element_names())
    return _branch_names_cache


def _target_qname(cls: type, namespace: str | None) -> str | None:
    return _context().fetch(cls, namespace).target_qname


def written_as(found: Entry, cls: type) -> str | None:
    """The element name the serialiser writes an instance of `cls` under in this field, or None.

    None where it writes an ``xsi:type`` (Word writes none) or refuses. The
    fork's ``XmlSerializer.convert_choice`` and ``xsi_type``, step for step:

    * a compound field: a class one of its choices declares, or a subclass of
      one, is written under that choice's name, with an ``xsi:type`` unless
      the class has no type name of its own; any other class is written under
      its own name, the type's, else the one :func:`docx4j_py.runtime.serializer_config`'s
      ``element_names`` gives it, else its metadata's ``qname`` (which is the
      class's name for most intermediate classes: ``MoveFrom1`` is written
      ``<MoveFrom1>``);
    * a list of one element: always that element's name, with an ``xsi:type``
      unless the class has no type name of its own.

    ``XmlMeta.qname`` is read here on purpose, from the shared context the
    serialiser uses, because it is what the serialiser writes; it is not an
    element name otherwise (``traversal.element_name`` is).
    """
    var = found.var
    namespace = found.namespace
    if found.compound:
        choice = var.find_clazz_choice(cls)
        if choice is not None:
            if cls in choice.types:
                return choice.qname
            target = _target_qname(cls, namespace)
            return choice.qname if target is None or target == choice.qname else None
        meta = _context().fetch(cls, namespace)
        return meta.target_qname or _branch_names().get(cls) or meta.qname
    if cls in var.types:
        return var.qname
    clazz = var.clazz
    if clazz is None:
        return None
    derived = issubclass(cls, clazz) or any(base is not object and issubclass(cls, base) for base in clazz.__bases__)
    if not derived:
        return None
    target = _target_qname(cls, namespace)
    return var.qname if target is None or target == var.qname else None


# ---------------------------------------------------------------------------
# the rule (D2)
# ---------------------------------------------------------------------------


def _class_types() -> Any:
    return _context().class_type


def _same_element(found: Entry, cls: type) -> bool:
    """Rule 2: the element name is a choice's, the shape that choice's class's, and it is written so."""
    from docx4j_py.traversal import element_name

    name = element_name(cls)
    if name is None:
        return False
    candidates = found.by_name.get(name)
    if not candidates:
        return False
    if written_as(found, cls) != name:
        return False
    shape = shape_of(cls)
    return shape is not None and any(shape_of(candidate) == shape for candidate in candidates)


def _derived_allowed(found: Entry, item: Any) -> bool:
    """A ``DerivedElement``: its name a choice's, and its value that choice's class or of its shape."""
    candidates = found.by_name.get(getattr(item, "qname", None))
    if candidates is None:
        return False
    value = getattr(item, "value", None)
    if not _is_model(type(value)):
        return found.simple
    cls = type(value)
    if cls in candidates:
        return True
    target = _target_qname(cls, found.namespace)
    if target is not None and target != item.qname:
        return False  # written with an xsi:type
    shape = shape_of(cls)
    return shape is not None and any(shape_of(candidate) == shape for candidate in candidates)


def verdict(cls: type, field: str, item: Any) -> bool:
    """Whether `item` goes into `cls`'s `field` (D2), judging and caching what the table does not hold.

    A verdict that depends only on the item's class is cached in the
    table's :attr:`Entry.verdicts`, so each class is judged once per field;
    a ``DerivedElement``'s depends on its name and value and is judged each
    time. True for a field that is not checked.
    """
    found = entry(cls, field)
    if found is None:
        return True
    item_cls = item.__class__
    known = found.verdicts.get(item_cls)
    if known is not None:
        return known
    types = _class_types()
    if isinstance(item, types.derived_element):
        return _derived_allowed(found, item)
    if isinstance(item, types.any_element):
        result = False  # a checked list has no wildcard
    elif isinstance(item, str):
        result = found.mixed or found.simple
    elif _is_model(item_cls):
        result = _same_element(found, item_cls)
    elif item is None:
        result = False
    else:
        result = found.simple  # a value of a simple type
    found.verdicts[item_cls] = result
    return result


def is_allowed(holder: Any, field: str, item: Any) -> bool:
    """Whether `item` may go into `holder`'s list `field` (a class or an instance), as the check judges it."""
    cls = holder if isinstance(holder, type) else holder.__class__
    return verdict(cls, field, item)


# ---------------------------------------------------------------------------
# the model, and the reverse index the derived hints read
# ---------------------------------------------------------------------------

_model_classes: list[type] | None = None
_reverse: dict[str, list[tuple[type, str]]] | None = None
_reverse_single: dict[str, list[tuple[type, str]]] | None = None


def model_classes() -> list[type]:
    """Every generated ``Child`` dataclass, namespace by namespace (1,969 of them)."""
    global _model_classes
    if _model_classes is None:
        from importlib import import_module

        from docx4j_py.child import Child
        from docx4j_py.el_index import EL_MODULES

        modules = sorted({path.rsplit(".", 1)[0] for path in EL_MODULES.values()})
        seen: set[type] = set()
        out: list[type] = []
        for name in modules:
            module = import_module(name)
            for value in vars(module).values():
                if (
                    isinstance(value, type)
                    and value.__module__ == name
                    and value not in seen
                    and issubclass(value, Child)
                    and dataclasses.is_dataclass(value)
                ):
                    seen.add(value)
                    out.append(value)
        _model_classes = out
    return _model_classes


def _build_reverse() -> None:
    global _reverse, _reverse_single
    from docx4j_py.child import _list_field_names

    lists: dict[str, list[tuple[type, str]]] = {}
    singles: dict[str, list[tuple[type, str]]] = {}
    for cls in model_classes():
        meta = _meta(cls)
        if meta is None:
            continue
        list_fields = set(_list_field_names(cls))
        for var in _element_vars(meta):
            if var.is_wildcard or var.is_text:
                continue
            names = var.elements if var.elements else (var.qname,)
            target = lists if var.name in list_fields else singles
            for name in names:
                target.setdefault(name, []).append((cls, var.name))
    _reverse, _reverse_single = lists, singles


#: The holders a derived hint names first, in this order; the rest follow by name.
_FIRST = (
    "Body",
    "Tc",
    "P",
    "R",
    "Tbl",
    "Tr",
    "SdtContentBlock",
    "CTSdtContentRun",
    "PHyperlink",
    "Hdr",
    "Ftr",
    "CTFtnEdn",
    "CommentsComment",
    "CTTxbxContent",
    "PPr",
    "RPr",
    "SectPr",
    "TblPr",
    "TrPr",
    "TcPr",
)


def _order(holders: list[tuple[type, str]]) -> list[tuple[type, str]]:
    def key(pair: tuple[type, str]) -> tuple[int, str, str]:
        name = pair[0].__name__
        rank = _FIRST.index(name) if name in _FIRST and pair[0].__module__ == "docx4j_py.wml" else len(_FIRST)
        return (rank, name, pair[1])

    return sorted(holders, key=key)


def holders_of(name: str) -> list[tuple[type, str]]:
    """Every (class, list field) whose choices include the element `name`, well-known holders first."""
    if _reverse is None:
        _build_reverse()
    return _order((_reverse or {}).get(name, []))


def property_holders_of(name: str) -> list[tuple[type, str]]:
    """Every (class, single-valued field) that holds the element `name`, well-known holders first."""
    if _reverse_single is None:
        _build_reverse()
    return _order((_reverse_single or {}).get(name, []))


def census() -> dict[str, int]:
    """The table's numbers over the whole model, for CR-020 section 10 and its tests.

    ``classes`` generated ``Child`` dataclasses; ``unbuildable`` of them
    xsdata cannot build; ``pairs`` (class, list field) pairs of the classes it
    can; ``checked`` of them with a table, ``unchecked`` without (wildcards
    and fields with no element choices); ``compound`` and ``element`` the
    checked ones by kind; ``mixed`` list fields that are mixed content;
    ``simple`` checked fields with a choice of a simple type; ``names``
    element names in compound fields with more than one class, ``one_shape``
    and ``differing`` of them.
    """
    from docx4j_py.child import _list_field_names

    counts = dict.fromkeys(
        (
            "classes",
            "unbuildable",
            "pairs",
            "checked",
            "unchecked",
            "compound",
            "element",
            "mixed",
            "simple",
            "names",
            "one_shape",
            "differing",
        ),
        0,
    )
    classes_by_name: dict[str, set[type]] = {}
    for cls in model_classes():
        counts["classes"] += 1
        meta = _meta(cls)
        if meta is None:
            counts["unbuildable"] += 1
            continue
        for var in _element_vars(meta):
            if var.elements:
                for qname, alternative in var.elements.items():
                    classes_by_name.setdefault(qname, set()).update(tp for tp in alternative.types if _is_model(tp))
        for field in _list_field_names(cls):
            counts["pairs"] += 1
            found = entry(cls, field)
            var = next((v for v in _element_vars(meta) if v.name == field), None)
            if var is not None and var.mixed:
                counts["mixed"] += 1
            if found is None:
                counts["unchecked"] += 1
                continue
            counts["checked"] += 1
            counts["compound" if found.compound else "element"] += 1
            if found.simple:
                counts["simple"] += 1
    for classes in classes_by_name.values():
        if len(classes) > 1:
            counts["names"] += 1
            if len({shape_of(cls) for cls in classes}) == 1:
                counts["one_shape"] += 1
            else:
                counts["differing"] += 1
    return counts


def differing_names() -> list[str]:
    """The element names in compound fields whose classes have more than one shape (section 2.3's eleven)."""
    classes_by_name: dict[str, set[type]] = {}
    for cls in model_classes():
        meta = _meta(cls)
        if meta is None:
            continue
        for var in _element_vars(meta):
            for qname, alternative in (var.elements or {}).items():
                classes_by_name.setdefault(qname, set()).update(tp for tp in alternative.types if _is_model(tp))
    return sorted(
        name
        for name, classes in classes_by_name.items()
        if len(classes) > 1 and len({shape_of(cls) for cls in classes}) > 1
    )


# ---------------------------------------------------------------------------
# the messages (D6)
# ---------------------------------------------------------------------------


def _prefixed(name: str | None) -> str:
    """``{w}tbl`` as ``w:tbl``; a namespace with no prefix in docx4j's table kept whole."""
    if not name:
        return "?"
    if not name.startswith("{"):
        return name
    from docx4j_py.namespaces import NAMESPACE_TO_PREFIX

    uri, _, local = name[1:].partition("}")
    prefix = NAMESPACE_TO_PREFIX.get(uri)
    return f"{prefix}:{local}" if prefix else name


def _local(name: str | None) -> str:
    return (name or "?").rpartition("}")[2]


def _article(word: str) -> str:
    """"a" or "an" before a class name as it is read aloud: an ``R``, an ``RT``, a ``RunIns``."""
    first = word[:1]
    if first in "AEIOU":
        return "an"
    if first in "FHLMNRSX" and (len(word) == 1 or word[1:2].isupper()):
        return "an"
    return "a"


def _name_of(item: Any) -> str | None:
    """The item's element name, as far as it has one."""
    types = _class_types()
    if isinstance(item, (types.derived_element, types.any_element)):
        qname = getattr(item, "qname", None)
        return qname if isinstance(qname, str) else None
    if _is_model(type(item)):
        from docx4j_py.traversal import element_name

        return element_name(item)
    return None


def _describe(item: Any, name: str | None) -> str:
    """``a `Tbl` (`w:tbl`)``, ``a `str```."""
    cls = type(item).__name__
    shown = f"{_article(cls)} `{cls}`"
    if name:
        shown += f" (`{_prefixed(name)}`)"
    return shown


def _holder_local(owner: Any) -> str:
    """What a holder is called in a hint's code: its element's local name, else its class's."""
    from docx4j_py.traversal import element_name

    name = element_name(owner)
    return _local(name) if name else type(owner).__name__.lower()


def _listing(pairs: list[tuple[type, str]], limit: int = 5) -> str:
    shown = ", ".join(f"`{cls.__name__}.{field}`" for cls, field in pairs[:limit])
    return shown + (", ..." if len(pairs) > limit else "")


#: The forms of a content control and of a custom XML element: class name -> (adjective, form, factory).
_SDT_FORMS = {
    "SdtBlock": ("block-level", "block", "el.sdt"),
    "SdtRun": ("run-level", "run", "el.sdt_run"),
    "CTSdtRow": ("row-level", "row", "el.sdt_row"),
    "CTSdtCell": ("cell-level", "cell", "el.sdt_cell"),
}
_CUSTOM_XML_FORMS = {
    "CTCustomXmlBlock": ("block-level", "block", "el.customXml"),
    "CTCustomXmlRun": ("run-level", "run", "el.custom_xml_run"),
    "CTCustomXmlRow": ("row-level", "row", "el.custom_xml_row"),
    "CTCustomXmlCell": ("cell-level", "cell", "el.custom_xml_cell"),
}

#: Run content, what each is called in a hint and how a run of it is built.
_RUN_CONTENT = {
    _w("t"): ("text", 'r("...")'),
    _w("delText"): ("deleted text", "el.r(content=[el.delText(...)])"),
    _w("instrText"): ("a field code", "el.r(content=[el.instrText(...)])"),
    _w("drawing"): ("a drawing", "el.r(content=[drawing])"),
    _w("pict"): ("a picture", "el.r(content=[pict])"),
    _w("br"): ("a break", "r(br())"),
    _w("tab"): ("a tab", "r(tab())"),
    _w("cr"): ("a carriage return", "el.r(content=[el.cr()])"),
    _w("sym"): ("a symbol", "el.r(content=[el.sym(...)])"),
    _w("fldChar"): ("a field character", "el.r(content=[el.fldChar(...)])"),
    _w("footnoteReference"): ("a footnote reference", "el.r(content=[el.footnoteReference(...)])"),
    _w("endnoteReference"): ("an endnote reference", "el.r(content=[el.endnoteReference(...)])"),
    _w("commentReference"): ("a comment reference", "el.r(content=[el.commentReference(...)])"),
    _w("object"): ("an embedded object", "el.r(content=[obj])"),
}

_TABLE_SIBLING = "a table is the paragraph's sibling: put it in the body, a table cell or a block-level content control"
_PARAGRAPH_BLOCK = (
    "a paragraph is a block, not inline content: put it in the body, a table cell or a block-level content control"
)


def _kinds(found: Entry) -> tuple[bool, bool, bool, bool, bool]:
    """What a list is: (holds runs, holds blocks, holds rows, holds cells, is a run's own content)."""
    names = found.by_name
    return (W_R in names, W_P in names, W_TR in names, W_TC in names, W_T in names)


def _wml_hint(found: Entry, owner: Any, item: Any, name: str) -> str | None:
    """The hand-written hints: docx4j-25's six mistakes, docx4j CR-064's table and Java's ``checkParent``."""
    runs, blocks, rows, cells, run_content = _kinds(found)
    inline = runs or run_content
    if name == W_TBL:
        if inline:
            return _TABLE_SIBLING
        if rows or cells:
            return "a nested table goes in a table cell (`Tc`), not straight in a table or a row"
    elif name == W_P:
        if inline:
            return _PARAGRAPH_BLOCK
        if rows:
            return "a paragraph goes in a cell, in a row: `el.tr(content=[el.tc(content=[p(...)])])`"
        if cells:
            return "a paragraph goes in a cell: `el.tc(content=[p(...)])`"
    elif name == W_R:
        if run_content:
            return "a run can't hold a run: put it in the paragraph, beside this one"
        if blocks:
            return "a run goes in a paragraph: `p(r(...))`"
        if rows or cells:
            return "a run goes in a paragraph, in a cell: `el.tc(content=[p(r(...))])`"
    elif name in _RUN_CONTENT and not run_content:
        what, example = _RUN_CONTENT[name]
        if runs:
            return f"{what} goes in a run: `{example}`"
        if blocks:
            return f"{what} goes in a run, in a paragraph: `p({example})`"
        if rows or cells:
            return f"{what} goes in a run, in a paragraph, in a cell"
    elif name == W_TC:
        if rows:
            return "a table holds rows: `el.tr(content=[el.tc(...)])`"
        if blocks or inline:
            return "a cell goes in a row, in a table: `el.tbl(content=[el.tr(content=[el.tc(...)])])`"
    elif name == W_TR:
        if cells:
            return "a row can't hold a row: put it in the table, beside this one"
        if blocks or inline:
            return "a row goes in a table: `el.tbl(content=[el.tr(...)])`"
    elif name == W_SDT_CONTENT:
        return (
            "a `w:sdtContent` is a content control's own part, not content: "
            "nest the control that holds it (`sdt(...)` builds both)"
        )
    if name in (W_SDT, W_CUSTOM_XML) and name in found.by_name:
        forms = _SDT_FORMS if name == W_SDT else _CUSTOM_XML_FORMS
        mine = forms.get(type(item).__name__)
        theirs = next((forms[c.__name__] for c in found.by_name[name] if c.__name__ in forms), None)
        if mine is not None and theirs is not None:
            what = "content control" if name == W_SDT else "custom XML element"
            where = (
                "in a paragraph"
                if inline
                else "among a table's rows"
                if rows
                else "among a row's cells"
                if cells
                else "at block level"
            )
            builder = f' or `sdt(..., form="{theirs[1]}")`' if name == W_SDT else ""
            return (
                f"this is {_article(mine[0])} {mine[0]} {what} (`{type(item).__name__}`); "
                f"{where} use the {theirs[1]} form, `{theirs[2]}(...)`{builder}"
            )
    return None


def _string_hint(found: Entry, holder: str) -> str:
    runs, blocks, rows, cells, run_content = _kinds(found)
    if run_content:
        return 'a string is not an element: wrap it, `t("...")`'
    if runs:
        return 'a string is not an element: wrap it, `r("...")`'
    if blocks:
        return 'a string is not an element: wrap it, `p("...")`'
    if rows or cells:
        return 'a string is not an element: text goes in a cell, `tc("...")`'
    return f"a string is not an element: `{holder}` takes {_takes(found)}"


def _takes(found: Entry, limit: int = 5) -> str:
    names = sorted(found.by_name)
    shown = ", ".join(f"`{_prefixed(n)}`" for n in names[:limit])
    return shown + (f" and {len(names) - limit} more" if len(names) > limit else "")


def _property_hint(owner: Any, item: Any, name: str) -> str | None:
    """`pPr` is a property, not content: set `p.p_pr`."""
    meta = _meta(type(owner))
    local = _local(name)
    if meta is not None:
        for var in _element_vars(meta):
            if var.is_element and not var.elements and var.qname == name and not var.list_element:
                hint = f"`{local}` is a property, not content: set `{_holder_local(owner)}.{var.name}`"
                if name == W_SECT_PR:
                    hint += ", or `p_pr.sect_pr` on the paragraph that ends the section"
                return hint
    return None


def _hint(found: Entry, owner: Any, field: str, item: Any, name: str | None) -> str:
    holder = f"{type(owner).__name__}.{field}"
    if isinstance(item, str):
        return _string_hint(found, holder)
    if item is None:
        return f"None is not an element: `{holder}` takes {_takes(found)}"
    if name is None:
        return f"`{holder}` takes {_takes(found)}"
    if name.startswith(f"{{{_W}}}"):
        hint = _wml_hint(found, owner, item, name)
        if hint is not None:
            return hint
    if name in found.by_name:
        theirs = found.by_name[name]
        if theirs:
            return (
                f"`{holder}` takes `{_prefixed(name)}` in another form, "
                f"{_article(theirs[0].__name__)} `{theirs[0].__name__}`; build that one"
                + (f", `{_factory_of(theirs[0])}(...)`" if _factory_of(theirs[0]) else "")
            )
        return f"`{holder}` takes `{_prefixed(name)}` as a value, not {_describe(item, None)}"
    hint = _property_hint(owner, item, name)
    if hint is not None:
        return hint
    holders = holders_of(name)
    singles = property_holders_of(name)
    local = _local(name)
    # a property element (CT_*Pr, by OOXML's own naming) is set, not added, even where some list takes it
    if singles and (not holders or local.endswith(("Pr", "PrEx", "PrChange"))):
        return f"`{local}` is a property, not content: set it as {_listing(singles)}"
    if holders:
        return f"`{_prefixed(name)}` goes in: {_listing(holders)}"
    return f"`{holder}` takes {_takes(found)}"


_factories: dict[type, str | None] = {}


def _factory_of(cls: type) -> str | None:
    """``el.<name>`` for the factory that builds exactly `cls`, or None."""
    try:
        return _factories[cls]
    except KeyError:
        pass
    found: str | None = None
    try:
        from importlib import import_module

        el = import_module(f"{cls.__module__}.el")
    except Exception:  # noqa: BLE001 - a module with no el
        el = None
    if el is not None:
        for name in getattr(el, "__all__", ()):
            factory = getattr(el, name, None)
            returns = getattr(factory, "__annotations__", {}).get("return") if callable(factory) else None
            if returns == cls.__name__ or returns is cls:
                found = f"el.{name}"
                break
    _factories[cls] = found
    return found


def refusal(owner: Any, field: str, item: Any) -> Exception:
    """The :class:`~docx4j_py.model.content.errors.MisplacedContentError` for an item refused (D5, D6).

    The message names the holder's field and the item, both its class and its
    element name ("`R.content` can't hold a `Tbl` (`w:tbl`)"); the hint says
    where the item goes: hand-written for the common mistakes, derived from
    the reverse index for the rest ("`w:x` goes in: `Body.content`, ...").
    """
    from docx4j_py.model.content.errors import MisplacedContentError

    found = entry(type(owner), field)
    name = _name_of(item)
    message = f"`{type(owner).__name__}.{field}` can't hold {_describe(item, name)}"
    hint = _hint(found, owner, field, item, name) if found is not None else None
    error = MisplacedContentError(message, hint=hint)
    error.holder = owner
    error.field = field
    error.item = item
    error.element = name
    return error
