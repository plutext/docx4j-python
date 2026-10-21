"""Parent pointers, parent-aware lists and subtree copying.

This is the runtime half of CR-001 section 5: the Python analogue of docx4j's
``org.jvnet.jaxb2_commons.ppp.Child`` interface, its ``ArrayListDocx4j``
collection type and ``XmlUtils.deepCopy``.

Three pieces, and nothing in them knows about WordprocessingML:

``Child``
    the base class the generator adds to every generated dataclass
    (``<Extensions><Extension type="class" class=".*" import="docx4j_py.child.Child"/>``).
    It contributes one ``__slots__`` entry, ``parent``, which is *not* a
    dataclass field: it is absent from ``__init__``, ``__eq__``, ``__repr__``
    and from anything xsdata serialises.

``ChildList``
    the ``default_factory`` of every generated list field
    (``<ListFactory>docx4j_py.child.ChildList</ListFactory>``). Appending to it
    sets the item's parent, so a tree built by hand is wired as it is built.

``link_parents``
    a single depth-first pass over a parsed tree that wires everything a
    constructor could not: single-valued element fields, and lists that were
    handed to a constructor as plain lists. It is driven by xsdata's
    ``XmlContext`` metadata, so it needs no generated code.

Example:
    >>> from docx4j_py.child import link_parents, deep_copy
    >>> doc = parser.from_bytes(data, Document)  # doctest: +SKIP
    >>> link_parents(doc)  # doctest: +SKIP
    >>> copy_of_body = deep_copy(doc.body)  # doctest: +SKIP
"""

from __future__ import annotations

import copy
import dataclasses
from collections.abc import Iterable, Iterator
from typing import Any, Self

__all__ = [
    "MCE_MODES",
    "MC_PREFIX_ATTRIBUTES",
    "BoundPrefixes",
    "Child",
    "ChildList",
    "deep_copy",
    "deep_copy_as",
    "is_alternate_content",
    "is_any_element",
    "iter_children",
    "iter_tree",
    "link_children",
    "link_parents",
    "mce_branch",
    "mce_branch_content",
    "mce_choices",
    "mce_fallback",
    "other_bindings_seen",
    "record_bindings",
    "requires_namespaces",
    "set_source_bindings",
    "source_bindings",
    "understands",
    "understood_namespaces",
]

#: The markup-compatibility namespace and the three element names of it that
#: the child enumeration has to understand (CR-002 section 5.6).
MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MC_ALTERNATE_CONTENT = f"{{{MCE_NS}}}AlternateContent"
MC_CHOICE = f"{{{MCE_NS}}}Choice"
MC_FALLBACK = f"{{{MCE_NS}}}Fallback"

#: What the ``mce=`` argument of :func:`iter_children` and of the traversal
#: functions accepts.
MCE_MODES = ("resolve", "all", "none")

# A key planted in the ``copy.deepcopy`` memo by the outermost ``__deepcopy__``
# so that the nested ones do not each relink the same subtree (that would make
# a copy quadratic). ``deepcopy`` only ever keys the memo by ``id()``, an int,
# so a str key cannot collide.
_DEEPCOPY_ROOT = "docx4j_py.child:deepcopy-root"


class ChildList(list):
    """A list that adopts what is put into it.

    Every mutation that introduces an item sets ``item.parent`` to the object
    that owns the list; every mutation that removes one clears it, unless
    :attr:`orphan_on_remove` is turned off. Items that are not :class:`Child`
    instances (tokens, ``AnyElement`` wildcards, plain strings) are ignored.

    The owner is set by :meth:`Child.__post_init__` on the instance the
    generator's ``default_factory=ChildList`` produced, and again by
    :func:`link_parents`. A ``ChildList`` with no owner is inert: it still
    behaves as a list, it just has no parent to hand out.

    Attributes:
        orphan_on_remove: class level switch; when True (the default) removing
            an item sets its parent back to None.
    """

    __slots__ = ("_owner",)

    orphan_on_remove: bool = True

    def __init__(self, iterable: Iterable[Any] = (), owner: Any = None) -> None:
        """Build the list and, if an owner is given, adopt its contents."""
        super().__init__(iterable)
        self._owner = owner
        if owner is not None:
            self._adopt_all()

    # -- ownership ---------------------------------------------------------

    @property
    def owner(self) -> Any:
        """The object whose field this list is, or None."""
        return self._owner

    def set_owner(self, owner: Any) -> None:
        """Make `owner` the parent of everything in this list, now and later."""
        self._owner = owner
        self._adopt_all()

    def _adopt(self, item: Any) -> Any:
        if isinstance(item, Child):
            item.parent = self._owner
        return item

    def _adopt_all(self) -> None:
        owner = self._owner
        for item in self:
            if isinstance(item, Child):
                item.parent = owner

    def _orphan(self, item: Any) -> None:
        if self.orphan_on_remove and isinstance(item, Child):
            item.parent = None

    # -- mutations that adopt ----------------------------------------------

    def append(self, item: Any) -> None:
        """Append an item and make the owner its parent."""
        super().append(self._adopt(item))

    def insert(self, index: int, item: Any) -> None:
        """Insert an item and make the owner its parent."""
        super().insert(index, self._adopt(item))

    def extend(self, iterable: Iterable[Any]) -> None:
        """Append every item of `iterable` and adopt each one."""
        items = [self._adopt(item) for item in iterable]
        super().extend(items)

    def __iadd__(self, iterable: Iterable[Any]) -> Self:
        """Implement ``list += iterable``, adopting every item."""
        self.extend(iterable)
        return self

    def __setitem__(self, index: Any, value: Any) -> None:
        """Assign by index or slice, orphaning what goes and adopting what comes."""
        if isinstance(index, slice):
            for old in self[index]:
                self._orphan(old)
            values = [self._adopt(item) for item in value]
            super().__setitem__(index, values)
        else:
            self._orphan(self[index])
            super().__setitem__(index, self._adopt(value))

    # -- mutations that orphan ---------------------------------------------

    def __delitem__(self, index: Any) -> None:
        """Delete by index or slice and clear the parent of what goes."""
        if isinstance(index, slice):
            for old in self[index]:
                self._orphan(old)
        else:
            self._orphan(self[index])
        super().__delitem__(index)

    def remove(self, item: Any) -> None:
        """Remove the first occurrence of `item` and clear its parent."""
        super().remove(item)
        self._orphan(item)

    def pop(self, index: int = -1) -> Any:
        """Remove and return an item, clearing its parent."""
        item = super().pop(index)
        self._orphan(item)
        return item

    def clear(self) -> None:
        """Remove every item, clearing each one's parent."""
        for item in self:
            self._orphan(item)
        super().clear()

    # -- copying and pickling ----------------------------------------------

    def __copy__(self) -> Self:
        """Return an unowned shallow copy; the items keep their parents."""
        return ChildList(self)

    def __deepcopy__(self, memo: dict) -> Self:
        """Return an unowned deep copy. The owner is never followed."""
        new = ChildList()
        memo[id(self)] = new
        for item in self:
            list.append(new, copy.deepcopy(item, memo))
        return new

    def __reduce__(self) -> tuple:
        """Pickle as a plain unowned list of items.

        The owner is deliberately not pickled; call :func:`link_parents` on the
        unpickled root to restore the pointers.
        """
        return (ChildList, (list(self),))


class Child:
    """A node that knows its parent.

    The generator adds this as a base class to every generated dataclass, so
    the generated source reads::

        @dataclass(slots=True, kw_only=True)
        class P(Child):
            p_pr: None | PPr = field(default=None, metadata={...})
            content: list[object] = field(default_factory=ChildList, metadata={...})

    ``parent`` is a slot rather than a dataclass field on purpose: it takes no
    ``XmlType.IGNORE`` metadata, it is not a constructor argument, and a parsed
    tree compares equal to an identically shaped constructed one.
    """

    __slots__ = ("parent",)

    def __post_init__(self) -> None:
        """Initialise the parent slot and bind every list field to this object.

        ``dataclasses`` calls this from the generated ``__init__`` because the
        method is visible on the class, inherited or not. A subclass that needs
        its own ``__post_init__`` must call ``super().__post_init__()``.
        """
        self.parent = None
        for name in _list_field_names(self.__class__):
            value = getattr(self, name, None)
            if isinstance(value, ChildList):
                value.set_owner(self)
            elif isinstance(value, list):
                # a plain list handed to the constructor, e.g. P(content=[r]);
                # promote it so the parent pointers work from here on
                object.__setattr__(self, name, ChildList(value, owner=self))

    def get_parent(self) -> Any:
        """Return the containing object, or None if this node is detached."""
        return getattr(self, "parent", None)

    def set_parent(self, parent: Any) -> None:
        """Set the containing object. Pass None to detach."""
        self.parent = parent

    def __deepcopy__(self, memo: dict) -> Self:
        """Deep copy the subtree without following the parent pointer.

        The copy is detached: its own ``parent`` is None, and its descendants
        are relinked to point inside the copy rather than inside the original.
        Only the outermost call relinks, so copying is linear.
        """
        cls = self.__class__
        new = cls.__new__(cls)
        memo[id(self)] = new

        is_root = _DEEPCOPY_ROOT not in memo
        if is_root:
            memo[_DEEPCOPY_ROOT] = True

        for name in _copyable_names(cls):
            try:
                value = getattr(self, name)
            except AttributeError:
                continue
            setattr(new, name, copy.deepcopy(value, memo))

        new.parent = None
        if is_root:
            del memo[_DEEPCOPY_ROOT]
            link_parents(new)
        return new


# ---------------------------------------------------------------------------
# per class caches
# ---------------------------------------------------------------------------

# field names whose generated default_factory is a ChildList
_list_fields_cache: dict[type, tuple[str, ...]] = {}
# field names __deepcopy__ has to carry over (everything but ``parent``)
_copyable_cache: dict[type, tuple[str, ...]] = {}
# field names that can hold a child object, for the link_parents walk
_child_fields_cache: dict[type, tuple[str, ...]] = {}


def _list_field_names(cls: type) -> tuple[str, ...]:
    """Names of the fields the generator gave a ChildList default factory."""
    try:
        return _list_fields_cache[cls]
    except KeyError:
        pass
    try:
        fields = dataclasses.fields(cls)
    except TypeError:  # not a dataclass; a hand written Child subclass
        names: tuple[str, ...] = ()
    else:
        names = tuple(
            f.name
            for f in fields
            if isinstance(f.default_factory, type) and issubclass(f.default_factory, ChildList)
        )
    _list_fields_cache[cls] = names
    return names


def _copyable_names(cls: type) -> tuple[str, ...]:
    """Names ``__deepcopy__`` must carry over, ``parent`` excluded."""
    try:
        return _copyable_cache[cls]
    except KeyError:
        pass
    try:
        names = tuple(f.name for f in dataclasses.fields(cls))
    except TypeError:
        seen: list[str] = []
        for klass in cls.__mro__:
            for name in getattr(klass, "__slots__", ()):
                if name != "parent" and name not in seen:
                    seen.append(name)
        names = tuple(seen)
    _copyable_cache[cls] = names
    return names


def _xml_context() -> Any:
    """Return a shared xsdata XmlContext, or None if the runtime is absent."""
    global _shared_context
    if _shared_context is _UNSET:
        try:
            from docx4j_xsdata.formats.dataclass.context import XmlContext
        except ImportError:  # pragma: no cover - only in a runtime-less install
            _shared_context = None
        else:
            _shared_context = XmlContext()
    return _shared_context


_UNSET = object()
_shared_context: Any = _UNSET


def _child_field_names(cls: type, context: Any) -> tuple[str, ...]:
    """Names of the fields that can hold a child object.

    Taken from the xsdata class metadata when the class has any: the element,
    wildcard, compound and text vars, and never the attribute vars, which saves
    the walk from touching thousands of strings and enums. Classes xsdata does
    not know (hand written dataclasses in the tests, for one) fall back to every
    dataclass field.
    """
    try:
        return _child_fields_cache[cls]
    except KeyError:
        pass

    names: tuple[str, ...] = ()
    if context is not None:
        try:
            meta = context.build(cls)
        except Exception:  # noqa: BLE001 - not an xsdata model, use the fallback
            meta = None
        if meta is not None:
            names = tuple(dict.fromkeys(var.name for var in meta.get_element_vars()))
    if not names:
        try:
            names = tuple(f.name for f in dataclasses.fields(cls))
        except TypeError:
            names = ()

    _child_fields_cache[cls] = names
    return names


# ---------------------------------------------------------------------------
# the walk
# ---------------------------------------------------------------------------


def _any_element_children(value: Any) -> list | None:
    """Return the children of an xsdata ``AnyElement``, or None."""
    children = getattr(value, "children", None)
    if children is not None and getattr(value, "qname", _UNSET) is not _UNSET:
        return children if isinstance(children, list) else None
    return None


#: A wildcard field's place in a child plan, where an element field has its
#: qualified name: a wildcard's typed child is named by what it is
#: (:func:`wildcard_child_name`), not by the field (CR-012 Phase B).
WILDCARD_FIELD = "\x00wildcard"

# field name -> the element qualified name it holds (WILDCARD_FIELD for a
# wildcard), or, for a compound ``content`` field, a mapping of item class to
# qualified name
_child_plan_cache: dict[type, tuple[tuple[str, str | None, dict[type, str] | None], ...]] = {}
# just the field names of the plan above, for the one-level link_children
_child_names_cache: dict[type, tuple[str, ...]] = {}


def _child_plan(
    cls: type, context: Any
) -> tuple[tuple[str, str | None, dict[type, str] | None], ...]:
    """The child fields of `cls`, each with the element name it writes.

    This is the one enumeration :func:`link_parents`, :func:`iter_children` and
    ``docx4j_py.traversal`` share: what the xsdata metadata says a class's
    children are, and what each of them is *called* in the document. The name
    matters because xsdata's element identity lives in two different places --- a
    single-valued field carries its own ``qname``, while the alternatives of a
    compound ``content`` field carry theirs in the field's ``choices`` --- and
    because the intermediate choice classes (CR-001 section 13.3) have no
    element name of their own at all: ``RT`` is ``w:t`` only because ``R.content``
    says so.
    """
    try:
        return _child_plan_cache[cls]
    except KeyError:
        pass

    plan: list[tuple[str, str | None, dict[type, str] | None]] = []
    meta = None
    if context is not None:
        try:
            meta = context.build(cls)
        except Exception:  # noqa: BLE001 - not an xsdata model, fall back below
            meta = None

    if meta is not None:
        seen: set[str] = set()
        for var in meta.get_element_vars():
            if var.name in seen:
                continue
            seen.add(var.name)
            if var.elements:
                choices: dict[type, str] = {}
                ambiguous: set[type] = set()
                for qname, alt in var.elements.items():
                    for tp in alt.types:
                        if tp in choices and choices[tp] != qname:
                            ambiguous.add(tp)
                        choices[tp] = qname
                for tp in ambiguous:
                    choices.pop(tp, None)
                plan.append((var.name, None, choices))
            elif var.is_wildcard:
                plan.append((var.name, WILDCARD_FIELD, None))
            else:
                plan.append((var.name, var.qname, None))
    else:
        try:
            plan = [(f.name, None, None) for f in dataclasses.fields(cls)]
        except TypeError:
            plan = []

    result = tuple(plan)
    _child_plan_cache[cls] = result
    return result


_wildcard_child_names: dict[type, str | None] = {}


def wildcard_child_name(item: Any) -> str | None:
    """The element name a typed child of a wildcard is written under: the serialiser's.

    No field names such an object --- a branch's ``R``, ``Drawing`` or
    ``RT``, the parser's hook typed it (CR-012) --- so the serialiser names it
    by its ``element_names`` table (:func:`docx4j_py.runtime._branch_element_names`:
    the ``w:sdt`` kinds, ``Tr``, ``Tc``, docx4j's root elements and the other
    classes a branch's container binds them to), else by the class's own
    global element declaration (``R`` is ``w:r``, ``P`` is ``w:p``). This
    answers the same, so that a walk names a branch's child as it names the
    same element elsewhere. None where neither says.
    """
    cls = item.__class__
    try:
        return _wildcard_child_names[cls]
    except KeyError:
        pass
    from docx4j_py.runtime import _branch_element_names

    name = _branch_element_names().get(cls)
    if name is None:
        meta = vars(cls).get("Meta")
        namespace = getattr(meta, "namespace", None)
        local = getattr(meta, "name", None)
        name = f"{{{namespace}}}{local}" if namespace and local else None
    _wildcard_child_names[cls] = name
    return name


def _qname_of(item: Any, choices: dict[type, str] | None, fallback: str | None) -> str | None:
    """The element name an item is written under, as far as the metadata knows."""
    if fallback is WILDCARD_FIELD:
        return wildcard_child_name(item)
    if choices is None:
        return fallback
    name = choices.get(item.__class__)
    if name is not None:
        return name
    for tp, qname in choices.items():
        if isinstance(item, tp):
            return qname
    return None


# ---------------------------------------------------------------------------
# markup compatibility (CR-002 section 5.6)
# ---------------------------------------------------------------------------


def is_any_element(value: Any) -> bool:
    """True for an xsdata ``AnyElement`` wildcard node.

    A wildcard node is what the parser builds for content the schema declares
    as ``xsd:any``: an ``mc:Choice``'s shape, a ``pic:pic`` inside
    ``a:graphicData``, a custom XML part's tree. It carries ``qname``,
    ``attributes``, ``children`` and ``text`` rather than fields, so anything
    that rewrites references (CR-003 section 4's ``insert_ooxml``) has to know
    the difference; :func:`docx4j_py.traversal.walk_all` is the walk that does.
    """
    return (
        not isinstance(value, Child)
        and getattr(value, "qname", _UNSET) is not _UNSET
        and isinstance(getattr(value, "children", None), list)
    )


#: The name this module used before CR-003 Phase A made it public.
_is_any_element = is_any_element


_understood: frozenset[str] | None = None


def understood_prefixes() -> frozenset[str]:
    """The prefixes this build of the model can parse.

    ``docx4j_py.namespaces.UNDERSTOOD``'s keys, generated from the set of
    generated packages; see CR-002 section 5.6: what the model can type. Since
    CR-012 Q8 neither ``mce_preprocess`` nor the traversal's ``mce="resolve"``
    selects by it: both read :func:`understood_namespaces`, docx4j's published
    list, as Word 365 does on CR-012 section 3.1's probes.
    """
    global _understood
    if _understood is None:
        try:
            from docx4j_py.namespaces import UNDERSTOOD
        except Exception:  # noqa: BLE001 - a model-less install; understand nothing
            _understood = frozenset()
        else:
            _understood = frozenset(UNDERSTOOD)
    return _understood


#: docx4j's published understood set, copied verbatim (CR-011 section 3.3).
UNDERSTOOD_NAMESPACES_RESOURCE = "understood-namespaces.txt"

_understood_namespaces: frozenset[str] | None = None


def understood_namespaces() -> frozenset[str]:
    """The namespaces a reader of text understands (WD REQ-306): docx4j's published list.

    ``docx4j_py/resources/understood-namespaces.txt``, copied verbatim from
    docx4j's ``org/docx4j/jaxb/understood-namespaces.txt`` (118 URIs at
    ``198d015a3``), the set docx4j's ``McSelection.understoodNamespaces``
    reads for its TEXT and READ purposes. Wider than
    :func:`understood_prefixes`: a namespace whose content a reader of text
    can reach, typed or not (VML's text box, chartex), is understood here.
    """
    global _understood_namespaces
    if _understood_namespaces is None:
        from importlib.resources import files

        text = files("docx4j_py.resources").joinpath(UNDERSTOOD_NAMESPACES_RESOURCE).read_text("utf-8")
        _understood_namespaces = frozenset(line.strip() for line in text.splitlines() if line.strip())
    return _understood_namespaces


class BoundPrefixes(str):
    """An ``mc`` attribute's value, with the namespaces its prefixes were bound to where it was parsed.

    CR-011 Q3, docx4j CR-052's ``McSourceBindings``. A binding keeps an
    attribute's string and not the declarations in scope at the element, and
    a producer may bind any prefix (docx4j's probe case 6 binds the wps
    namespace to ``shape2010``, and declares ``zz9`` on the
    ``mc:AlternateContent``). The fork's parser hands each ``mc:Choice`` it
    builds, typed or a wildcard, to :func:`record_bindings` with the
    declarations in scope (``ParserConfig.on_namespaces``), which replaces its
    ``Requires``, ``mc:Ignorable`` and ``mc:MustUnderstand`` values with these.

    It is the string it was, so equality and serialisation see the string.
    :attr:`namespaces` travels with it wherever the Choice goes: a move into
    another tree, ``deep_copy`` (it is immutable, and copies as itself), a
    pickle. An edit that sets a plain string drops it, and that value is then
    resolved through the prefix table; :func:`set_source_bindings` records one
    by hand. A ``str`` subclass cannot have slots, so the namespaces are its
    one instance attribute.
    """

    namespaces: dict[str, str | None]

    def __new__(cls, value: str, namespaces: Iterable[tuple[str, str | None]] = ()) -> BoundPrefixes:
        """The value, and each prefix it names with the URI it was bound to (None for none)."""
        self = super().__new__(cls, value)
        self.namespaces = dict(namespaces)
        return self

    def __reduce__(self) -> tuple[Any, ...]:
        """Pickled with its namespaces."""
        return (BoundPrefixes, (str(self), tuple(self.namespaces.items())))

    def __copy__(self) -> BoundPrefixes:
        """Immutable: itself."""
        return self

    def __deepcopy__(self, memo: dict) -> BoundPrefixes:
        """Immutable: itself."""
        return self


#: An ``mc:Choice``'s attributes that name prefixes: the typed field, and the
#: key a wildcard node keeps it under.
MC_PREFIX_ATTRIBUTES: tuple[tuple[str, str], ...] = (
    ("requires", "Requires"),
    ("ignorable", f"{{{MCE_NS}}}Ignorable"),
    ("must_understand", f"{{{MCE_NS}}}MustUnderstand"),
)

#: Whether any binding recorded in this process differs from the prefix
#: table's: until one does, a save can place every prefix from the table alone
#: and need not read the records (:mod:`docx4j_py.mc_declarations`).
_other_bindings_seen = False


def other_bindings_seen() -> bool:
    """Whether a binding the prefix table does not make has been recorded in this process.

    A private prefix (``zz9``), an alias (``shape2010``), a table prefix bound
    to another namespace, or a prefix bound to nothing.
    """
    return _other_bindings_seen


def record_bindings(qname: str, obj: Any, ns_map: dict) -> None:
    """The fork parser's ``on_namespaces`` hook for ``mc:Choice``: its prefix-naming values recorded.

    Typed (``AlternateContentChoice.requires``, ``.ignorable``,
    ``.must_understand``) or a wildcard (``AnyElement.attributes``). Only the
    prefixes each value names are kept, each with the URI it is bound to in
    scope there, or None.
    """
    if _is_any_element(obj):
        attributes = obj.attributes
        if not attributes:
            return
        for _field, key in MC_PREFIX_ATTRIBUTES:
            value = attributes.get(key)
            if isinstance(value, str) and value.split():
                attributes[key] = _recorded(value, ns_map)
        return
    for field, _key in MC_PREFIX_ATTRIBUTES:
        value = getattr(obj, field, None)
        if isinstance(value, str) and value.split():
            setattr(obj, field, _recorded(value, ns_map))


def _recorded(value: str, bindings: Any) -> BoundPrefixes:
    global _other_bindings_seen
    pairs = tuple((prefix, bindings.get(prefix) or None) for prefix in dict.fromkeys(value.split()))
    if not _other_bindings_seen:
        from docx4j_py.namespaces import PREFIXES

        if any(uri is None or PREFIXES.get(prefix) != uri for prefix, uri in pairs):
            _other_bindings_seen = True
    return BoundPrefixes(value, pairs)


def source_bindings(choice: Any) -> dict[str, str | None]:
    """What a Choice's prefix-naming values were bound to in its source; empty where nothing was recorded.

    docx4j ``McSourceBindings.of``: ``Requires``, ``mc:Ignorable`` and
    ``mc:MustUnderstand`` together, typed or a wildcard.
    """
    out: dict[str, str | None] = {}
    for value in _prefix_values(choice):
        if isinstance(value, BoundPrefixes):
            for prefix, uri in value.namespaces.items():
                out.setdefault(prefix, uri)
    return out


def set_source_bindings(choice: Any, bindings: dict[str, str]) -> None:
    """Record what a Choice's prefixes mean, for a Choice built or renamed in code. docx4j ``McSourceBindings.set``.

    Each of its ``Requires``, ``mc:Ignorable`` and ``mc:MustUnderstand``
    values is recorded with `bindings`' URI for each prefix it names.
    """
    if _is_any_element(choice):
        attributes = choice.attributes or {}
        for _field, key in MC_PREFIX_ATTRIBUTES:
            value = attributes.get(key)
            if isinstance(value, str) and value.split():
                attributes[key] = _recorded(str(value), bindings)
        return
    for field, _key in MC_PREFIX_ATTRIBUTES:
        value = getattr(choice, field, None)
        if isinstance(value, str) and value.split():
            setattr(choice, field, _recorded(str(value), bindings))


def _prefix_values(choice: Any) -> list[Any]:
    if _is_any_element(choice):
        attributes = choice.attributes or {}
        return [attributes.get(key) for _field, key in MC_PREFIX_ATTRIBUTES]
    return [getattr(choice, field, None) for field, _key in MC_PREFIX_ATTRIBUTES]


def requires_of(choice: Any) -> str | None:
    """An ``mc:Choice``'s ``Requires``, typed or a wildcard."""
    if _is_any_element(choice):
        return (choice.attributes or {}).get("Requires")
    return getattr(choice, "requires", None)


def requires_namespaces(choice: Any) -> list[str | None]:
    """The namespace URIs a Choice's ``Requires`` names, one per prefix, in order; None for a prefix bound to none.

    docx4j ``McSelection.requiresNamespaces`` (WD REQ-303): matched by URI,
    through the declarations recorded where the Choice was parsed
    (:class:`BoundPrefixes`). Where none were recorded --- a Choice built in
    code, or one whose ``Requires`` was set since --- docx4j's prefix table
    (:data:`docx4j_py.namespaces.PREFIXES`) is used instead, as REQ-303
    permits. Empty for an empty or whitespace-only ``Requires``.
    """
    requires = requires_of(choice)
    if not isinstance(requires, str):
        return []
    prefixes = requires.split()
    recorded = requires.namespaces if isinstance(requires, BoundPrefixes) else None
    if recorded is None:
        from docx4j_py.namespaces import PREFIXES

        return [PREFIXES.get(prefix) for prefix in prefixes]
    return [recorded.get(prefix) for prefix in prefixes]


def understands(choice: Any) -> bool:
    """Whether a reader of text understands a Choice (WD REQ-302 to REQ-304).

    Its ``Requires`` names at least one prefix (REQ-304: an empty one is not
    understood), and every prefix resolves to a namespace in
    :func:`understood_namespaces` (REQ-303: by URI). docx4j
    ``McSelection.understands``.
    """
    uris = requires_namespaces(choice)
    if not uris:
        return False
    known = understood_namespaces()
    return all(uri is not None and uri in known for uri in uris)


def is_alternate_content(node: Any) -> bool:
    """True for an ``mc:AlternateContent`` node, typed or as a wildcard."""
    if _is_any_element(node):
        return node.qname == MC_ALTERNATE_CONTENT
    cls = node.__class__
    return (
        cls.__module__ == "docx4j_py.mce"
        and cls.__name__ == "AlternateContent"
        and hasattr(node, "choice")
    )


def mce_choices(node: Any) -> list[Any]:
    """The ``mc:Choice`` branches of an ``mc:AlternateContent``, typed or wildcard, in order."""
    if _is_any_element(node):
        return [child for child in node.children or () if getattr(child, "qname", None) == MC_CHOICE]
    return list(getattr(node, "choice", None) or ())


def mce_fallback(node: Any) -> Any | None:
    """The ``mc:Fallback`` of an ``mc:AlternateContent``, typed or wildcard, or None."""
    if _is_any_element(node):
        for child in node.children or ():
            if getattr(child, "qname", None) == MC_FALLBACK:
                return child
        return None
    return getattr(node, "fallback", None)


def mce_branch(node: Any) -> Any | None:
    """The branch of an ``mc:AlternateContent`` a reader of text takes (CR-011 section 3.2 item 1).

    docx4j's TEXT purpose (``McSelection.selectedBranch(TEXT, ...)``), which
    is also its ``McMode.READ`` since 17.3.4: the **first** ``mc:Choice`` that
    :func:`understands`, by namespace URI over docx4j's published list (WD
    REQ-302, 303, 304, 306); else the ``mc:Fallback``; else, for an element
    with Choices and no Fallback, its first Choice (REQ-305: a reading rule).
    None for an element with no branch at all.

    Both shapes are handled: the typed ``docx4j_py.mce.AlternateContent`` the
    parser builds where the schema declares the element, and the ``AnyElement``
    one it builds inside a wildcard.
    """
    choices = mce_choices(node)
    for choice in choices:
        if understands(choice):
            return choice
    fallback = mce_fallback(node)
    if fallback is not None:
        return fallback
    return choices[0] if choices else None


def mce_branch_content(branch: Any) -> list[Any]:
    """What a branch holds: a typed one's ``any_element``, a wildcard one's children; empty for None."""
    if branch is None:
        return []
    if _is_any_element(branch):
        return [child for child in branch.children or () if not isinstance(child, str)]
    return list(getattr(branch, "any_element", None) or ())


def iter_children(
    obj: Any,
    *,
    context: Any = None,
    mce: str = "resolve",
    wildcards: bool = False,
) -> Iterator[tuple[str | None, Any]]:
    """Yield ``(element qualified name, child)`` for every child of `obj`.

    Driven entirely by the class metadata, so it works on any generated package
    with no per-class code. ``AnyElement`` wildcards are descended into: their
    own children are yielded under their ``qname``. A typed object a wildcard
    holds --- a kept branch's ``R`` or, since CR-012 Phase B, its ``Drawing``
    or ``RT`` --- is yielded under the name it is written under
    (:func:`wildcard_child_name`), as the same element is elsewhere; before
    Phase B it was yielded under the field's name, ``any_element``. The name
    is None where the metadata cannot say (an ambiguous compound
    alternative).

    Args:
        context: an ``XmlContext`` to read the metadata from.
        mce: what to do with ``mc:AlternateContent`` (CR-002 section 5.6).

            ``"resolve"`` (the default)
                the view a reader of text takes: the element is transparent
                and its children are the children of :func:`mce_branch`'s
                branch --- the first ``mc:Choice`` whose ``Requires``
                namespaces are all in docx4j's published understood list
                (matched by URI, CR-011 section 3.3), else ``mc:Fallback``,
                else the first ``mc:Choice``.
            ``"all"``
                every branch, ``mc:Choice`` and ``mc:Fallback`` nodes included.
                This is what :func:`link_parents` and the serialiser need: the
                tree is lossless and both branches are written back.
            ``"none"``
                ``mc:AlternateContent`` is a leaf.
        wildcards: also yield ``AnyElement`` wildcard nodes themselves, under
            their own ``qname``, instead of only the typed objects below them.
            Off by default, so ``walk``, ``find`` and ``link_parents`` see only
            model objects; ``text_of`` turns it on, because the content of an
            ``mc:Choice`` is a wildcard tree and its text is real text.
    """
    if context is None:
        context = _xml_context()

    if mce != "all" and is_alternate_content(obj):
        if mce == "none":
            return
        branch = mce_branch(obj)
        if branch is not None:
            yield from iter_children(branch, context=context, mce=mce, wildcards=wildcards)
        return

    if _is_any_element(obj):
        for inner in obj.children or ():
            yield from _one_child(inner, None, None, mce, wildcards, context)
        return

    for name, qname, choices in _child_plan(obj.__class__, context):
        value = getattr(obj, name, None)
        if value is None:
            continue
        items = value if isinstance(value, (list, tuple)) else (value,)
        for item in items:
            yield from _one_child(item, qname, choices, mce, wildcards, context)


def _one_child(
    item: Any,
    qname: str | None,
    choices: dict[type, str] | None,
    mce: str,
    wildcards: bool,
    context: Any,
) -> Iterator[tuple[str | None, Any]]:
    if mce != "all" and is_alternate_content(item):
        # The resolved view is docx4j's preprocessed one: the element itself
        # disappears and the chosen branch's children stand in its place, so a
        # caller never sees an mc:AlternateContent, an mc:Choice or an
        # mc:Fallback, and never sees the same content twice.
        if mce != "none":
            yield from iter_children(item, context=context, mce=mce, wildcards=wildcards)
        return
    if isinstance(item, Child):
        yield _qname_of(item, choices, qname), item
    elif wildcards and _is_any_element(item):
        yield item.qname if isinstance(item.qname, str) else None, item
    else:
        yield from _wildcard_children(item, mce)


def _wildcard_children(value: Any, mce: str = "all") -> Iterator[tuple[str | None, Any]]:
    """Yield the typed objects an ``AnyElement`` wildcard holds, at any depth.

    `mce` is honoured on the way down, so that a caller asking for the resolved
    view does not reach the branch of an ``mc:AlternateContent`` that a
    consumer would ignore, even when that branch is buried in a wildcard tree.
    """
    if mce != "all" and is_alternate_content(value):
        if mce == "none":
            return
        branch = mce_branch(value)
        if branch is not None:
            yield from _wildcard_children(branch, mce)
        return
    children = _any_element_children(value)
    if not children:
        return
    for inner in children:
        if isinstance(inner, Child):
            yield wildcard_child_name(inner), inner
        else:
            yield from _wildcard_children(inner, mce)


def link_children(obj: Any, *, context: Any = None) -> None:
    """Wire the parent pointer of `obj`'s **immediate** children, and no more.

    One level of what :func:`link_parents` does, from the same ``_child_plan``
    enumeration and by the same rules: every ``Child`` in a list field or in a
    single-valued field gets `obj` as its parent, and every :class:`ChildList`
    is bound to `obj` as its owner. A typed object nested inside an
    ``AnyElement`` wildcard keeps the parent it has, exactly as
    :func:`link_parents` leaves it.

    This is the parser's hook (CR-002 section 12.14). xsdata binds bottom up,
    so by the time an object exists its own children exist and are bound, and
    calling this on every object the parser builds ---
    ``ParserConfig(on_bind=...)``, which :func:`docx4j_py.runtime.parser` fills
    in --- wires the whole tree as it is parsed, with no second walk.
    :func:`link_parents` stays for the trees this does not see: hand built
    ones, ``deep_copy``, and anything a caller assembles field by field.

    Args:
        obj: the object whose children to wire.
        context: an ``XmlContext`` to take class metadata from; the parser's
            own is the cheapest one to pass. Defaults to a shared one.
    """
    cls = obj.__class__
    try:
        names = _child_names_cache[cls]
    except KeyError:
        if context is None:
            context = _xml_context()
        names = _child_names_cache[cls] = tuple(
            entry[0] for entry in _child_plan(cls, context)
        )

    for name in names:
        value = getattr(obj, name, None)
        if value is None:
            continue
        if isinstance(value, Child):
            value.parent = obj
        elif isinstance(value, ChildList):
            value.set_owner(obj)
        elif isinstance(value, (list, tuple)):
            for item in value:
                if isinstance(item, Child):
                    item.parent = obj


def link_parents(root: Any, *, context: Any = None) -> int:
    """Wire the parent pointer of every node below `root`.

    A single depth-first pass, driven by the xsdata class metadata, so it works
    on any generated package without per-class code. It handles compound
    (``content``) list fields, single-valued element fields, nested and derived
    element classes, and it descends through ``AnyElement`` wildcards without
    tripping over them. ``root.parent`` itself is left alone.

    Every ``ChildList`` it meets is bound to its owner on the way, so a list
    that reached a constructor as a plain list, or was replaced wholesale, ends
    up as usable as one the constructor built.

    **Markup compatibility does not apply here.** ``link_parents`` is
    ``mce="all"`` and always will be: every node of the tree is written back on
    save, so every node needs its parent, the branch of an
    ``mc:AlternateContent`` a consumer would ignore included (CR-002 section
    5.6).

    Args:
        root: the object to start from, typically a part root just parsed.
        context: an ``XmlContext`` to take class metadata from; the parser's
            own context is the cheapest one to pass. Defaults to a shared one.

    Returns:
        The number of nodes visited, root included.
    """
    if context is None:
        context = _xml_context()

    visited = 0
    stack: list[Any] = [root]
    seen: set[int] = set()

    while stack:
        obj = stack.pop()
        key = id(obj)
        if key in seen:
            continue
        seen.add(key)
        visited += 1

        for name, _qname, _choices in _child_plan(obj.__class__, context):
            value = getattr(obj, name, None)
            if value is None:
                continue
            if isinstance(value, ChildList):
                value.set_owner(obj)
                for item in value:
                    if isinstance(item, Child):
                        stack.append(item)
                    else:
                        stack.extend(c for _q, c in _wildcard_children(item))
            elif isinstance(value, (list, tuple)):
                for item in value:
                    if isinstance(item, Child):
                        item.parent = obj
                        stack.append(item)
                    else:
                        stack.extend(c for _q, c in _wildcard_children(item))
            elif isinstance(value, Child):
                value.parent = obj
                stack.append(value)
            else:
                stack.extend(c for _q, c in _wildcard_children(value))

    return visited


def iter_tree(root: Any, *, context: Any = None, mce: str = "all") -> Iterator[Any]:
    """Yield `root` and every :class:`Child` below it, depth first.

    The traversal is the one :func:`link_parents` uses, without the writing;
    it is what a parent checker or a ``walk``/``find`` helper is built on. It
    is ``mce="all"`` by default for the same reason ``link_parents`` is: a
    parent checker has to reach every node that will be written back.
    """
    if context is None:
        context = _xml_context()

    stack: list[Any] = [root]
    seen: set[int] = set()
    while stack:
        obj = stack.pop()
        key = id(obj)
        if key in seen:
            continue
        seen.add(key)
        yield obj

        for _qname, child in iter_children(obj, context=context, mce=mce):
            stack.append(child)


def deep_copy[T](obj: T, parent: Any = None) -> T:
    """Copy a subtree, optionally attaching the copy to `parent`.

    The docx4j idiom (``XmlUtils.deepCopy``, and ``deepCopy(value, parent)`` in
    docx4j-core-ts): the copy's descendants point inside the copy, the copy
    itself points at `parent`, and the original document is never followed, let
    alone copied.
    """
    new = copy.deepcopy(obj)
    if isinstance(new, Child):
        new.parent = parent
    return new


def deep_copy_as[T](obj: Any, cls: type[T], parent: Any = None) -> T:
    """Copy a subtree and re-type the copy as `cls`, keeping what `cls` declares.

    The reason this exists is ``xsi:type``. xsdata writes it whenever the class
    of the value in a field differs from the class the field declares, which is
    valid XML Schema and is not what Word writes: ``w:pPrChange/w:pPr`` is
    declared ``CT_PPrBase``, so putting a ``PPr`` (which extends it, adding
    ``w:rPr``, ``w:sectPr`` and ``w:pPrChange``) in there marshals

    .. code-block:: xml

        <w:pPr xsi:type="w:CT_PPr">

    where Word writes a plain ``<w:pPr>``. Copying *as* ``PPrBase`` gives the
    plain element, and drops the three fields a ``w:pPrChange`` must not hold
    anyway. It is the Python form of docx4j-generated-objects-ts's
    ``deepCopyAs`` (its CR-003 section 2) and of docx4j's habit of building a
    ``PPrBase`` by hand.

    `cls` is normally a base class of `obj`'s, but a sibling works too: the
    copy keeps every field `cls` declares that `obj` has a value for, and
    nothing else. Parents inside the copy are linked, and the copy's own parent
    is `parent`, as :func:`deep_copy`'s is.

    Args:
        obj: the subtree to copy.
        cls: the class to build. Every generated class takes keyword
            arguments and every field is optional (CR-001 section 3), so this
            is always possible.
        parent: what to attach the copy to; None leaves it detached.

    Raises:
        TypeError: if `cls` is not a dataclass, or declares no field that
            `obj` has --- which would silently return an empty element.
    """
    try:
        fields = dataclasses.fields(cls)
    except TypeError:
        raise TypeError(
            f"deep_copy_as needs a generated class, not {cls!r} "
            "(pass the base or sibling class to re-type as, e.g. PPrBase)"
        ) from None

    values: dict[str, Any] = {}
    for field in fields:
        value = getattr(obj, field.name, None)
        if value is None or (isinstance(value, list) and not value):
            continue
        values[field.name] = copy.deepcopy(value)

    if not values and not isinstance(obj, cls):
        theirs = ", ".join(sorted(f.name for f in fields)[:6]) or "no fields"
        raise TypeError(
            f"{type(obj).__name__} has nothing {cls.__name__} declares "
            f"({cls.__name__} has {theirs}); it is neither a base nor a sibling"
        )

    new = cls(**values)
    link_parents(new)
    if isinstance(new, Child):
        new.parent = parent
    return new
