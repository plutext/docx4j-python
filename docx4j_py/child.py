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
    "Child",
    "ChildList",
    "deep_copy",
    "deep_copy_as",
    "is_any_element",
    "iter_children",
    "iter_tree",
    "link_parents",
    "mce_branch",
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


# field name -> the element qualified name it holds, or, for a compound
# ``content`` field, a mapping of item class to qualified name
_child_plan_cache: dict[type, tuple[tuple[str, str | None, dict[type, str] | None], ...]] = {}


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


def _qname_of(item: Any, choices: dict[type, str] | None, fallback: str | None) -> str | None:
    """The element name an item is written under, as far as the metadata knows."""
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
    generated packages; see CR-002 section 5.6.
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


def _requires_understood(requires: str | None) -> bool:
    known = understood_prefixes()
    return all(prefix in known for prefix in (requires or "").split())


def mce_branch(node: Any) -> Any | None:
    """The branch of an ``mc:AlternateContent`` a consumer like this one takes.

    ECMA-376 Part 3: the **first** ``mc:Choice`` whose ``Requires`` prefixes are
    *all* understood, and ``mc:Fallback`` if there is no such choice. "Under-
    stood" is :func:`understood_prefixes`, which is the set of namespaces this
    build generated classes for. Returns None when there is no choice and no
    fallback --- which is legal and means "ignore this element".

    Both shapes are handled: the typed ``docx4j_py.mce.AlternateContent`` the
    parser builds where the schema declares the element, and the ``AnyElement``
    one it builds inside a wildcard.
    """
    if _is_any_element(node):
        fallback = None
        for child in node.children or ():
            qname = getattr(child, "qname", None)
            if qname == MC_CHOICE:
                requires = (getattr(child, "attributes", None) or {}).get("Requires")
                if _requires_understood(requires):
                    return child
            elif qname == MC_FALLBACK and fallback is None:
                fallback = child
        return fallback

    for choice in getattr(node, "choice", None) or ():
        if _requires_understood(getattr(choice, "requires", None)):
            return choice
    return getattr(node, "fallback", None)


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
    own children are yielded under their ``qname``. The name is None where the
    metadata cannot say (a wildcard holding a typed object, an ambiguous
    compound alternative).

    Args:
        context: an ``XmlContext`` to read the metadata from.
        mce: what to do with ``mc:AlternateContent`` (CR-002 section 5.6).

            ``"resolve"`` (the default)
                the view Word takes: the element is transparent and its
                children are the children of the branch a consumer that
                understands :func:`understood_prefixes` would take --- the
                first ``mc:Choice`` whose ``Requires`` prefixes are all
                understood, else ``mc:Fallback``.
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
            yield getattr(inner, "qname", None), inner
        else:
            yield from _wildcard_children(inner, mce)


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
