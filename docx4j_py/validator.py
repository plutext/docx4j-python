"""The validator: what a tree, a part or a package holds that Word reads otherwise, or not at all (CR-020 Phase B).

CR-020 D7 (section 4.7). Phase A refuses misplaced content as code adds it to a
typed list (:mod:`docx4j_py.placement`, ``ChildList``); this walks a tree and
reports what got past that check, and what no add-time check can see:

**Misplaced content** (``target.misplaced``, :class:`MisplacedContent`). Every
item of every typed list, and every single-valued field's value, held to D2,
Phase A's rule, by :func:`docx4j_py.placement.verdict` (reused, not
re-implemented): a plain list assigned whole (``run.content = [Tbl()]``), a
single field assigned (``p.p_pr = R()``), ``object.__setattr__``, and anything
built under :func:`docx4j_py.child.unchecked`. The message and hint are the
add-time error's (:func:`docx4j_py.placement.refusal`), so one search finds both.

**A table cell that does not end with a paragraph** (``cell.no_end_paragraph``,
:class:`CellWithoutEndParagraph`), as CR-016 Phase A reads it
(:func:`docx4j_py.model.content.table_rules.ends_with_paragraph`): past range
markup, through block-level controls and custom XML, and through the branch of
a kept ``mc:AlternateContent`` Word reads. Word opens such a file as unreadable
content.

**An integer measurement past what Word stores** (``value.out_of_range``,
:class:`ValueOutOfRange`): an attribute of CR-022's limits table
(:func:`docx4j_py.model.values.integer_limit`) whose value, as Word reads it
lexically (:func:`docx4j_py.model.values.integer`, CR-021, first), Word does not
keep as written. Word's reading is :mod:`docx4j_py.model.word_integers`' wrap
model. The fields are docx4j-core-ts CR-017's in snake_case, its ``storage``
strings exactly. Read-side semantics (section 4.7): 65,535 where Word's save
writes it ``-1`` is in range, and so is a ``-1`` there; ``w:col/@w:space``'s -1
is Word's sentinel; a row height under ``hRule="auto"``, an unmeasured
attribute and an attribute with no row (a table width under ``auto``) are not
checked; the write side's 65,534 cap is not a read limit.

Nothing is raised for a finding, and nothing is rewritten: a finding is plain
data with ``to_dict()``. A part never unmarshalled is read privately
(:meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart.read_contents`), so it
is still written byte for byte. ``package.save(check=True)``
(:func:`check_for_save`) is the one place a finding raises.
"""

from __future__ import annotations

import dataclasses
import re
from collections import Counter
from enum import Enum
from typing import Any, ClassVar, NamedTuple

__all__ = [
    "CellWithoutEndParagraph",
    "Finding",
    "MisplacedContent",
    "ValueOutOfRange",
    "check_for_save",
    "findings_in",
    "misplaced",
    "part_findings",
]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(prefixed: str) -> str:
    """``w:ind`` as ``{...main}ind``."""
    return f"{{{_W}}}{prefixed.removeprefix('w:')}"


# ---------------------------------------------------------------------------
# the findings
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class Finding:
    """One thing the validator found: where it is, and what it is.

    Attributes:
        message: what is wrong, in a sentence.
        part_name: the part, ``/word/document.xml``; None for a tree walked on
            its own.
        path: where it is, from the part's root, each step a prefixed name and
            its 1-based position among its siblings of that name
            (``/w:document/w:body[1]/w:p[3]/w:pPr[1]/w:ind[1]/@w:left``),
            docx4j-core-ts's form; an attribute's path ends with it.
        address: the content API's address of the innermost block holding it
            (``w14:5A2B1C3D``, ``body/3/0/1``), where the part has a body.
        node: the object the finding is about, for a caller that fixes it (the
            item misplaced, the cell, the element carrying the attribute); not
            in :meth:`to_dict`, nor compared.
    """

    #: The finding's code, as the content API's errors spell theirs.
    code: ClassVar[str] = "content.finding"
    #: Fields :meth:`to_dict` keeps when None, because None there means something.
    _keep_none: ClassVar[tuple[str, ...]] = ()

    message: str
    part_name: str | None = None
    path: str | None = None
    address: str | None = None
    node: Any = dataclasses.field(default=None, repr=False, compare=False)

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: ``code`` and the location first, then the finding's own fields, then the message.

        None is left out, as the other reports leave it (CR-003 section 3.1),
        except where it carries meaning (:class:`ValueOutOfRange`'s
        ``word_reads`` and ``range``).
        """
        out: dict[str, Any] = {"code": self.code}
        names = [f.name for f in dataclasses.fields(self) if f.name not in ("node", "message", "hint")]
        for name in (*names, "message", "hint"):
            if not hasattr(self, name):
                continue
            value = getattr(self, name)
            if value is None and name not in self._keep_none:
                continue
            out[name] = list(value) if isinstance(value, tuple) else value
        return out


@dataclasses.dataclass(frozen=True, slots=True)
class MisplacedContent(Finding):
    """An item a typed list, or a single-valued field, cannot hold (D2): ``target.misplaced``.

    Attributes:
        holder: the holder's class (``"R"``).
        field: the field (``"content"``).
        index: the item's position in the list; None for a single-valued field.
        item: the item's class (``"Tbl"``, ``"str"``).
        element: the item's element name, prefixed (``"w:tbl"``), where it has one.
        hint: where the item goes, as the add-time error says it.
    """

    code: ClassVar[str] = "target.misplaced"

    holder: str | None = None
    field: str | None = None
    index: int | None = None
    item: str | None = None
    element: str | None = None
    hint: str | None = None


@dataclasses.dataclass(frozen=True, slots=True)
class CellWithoutEndParagraph(Finding):
    """A table cell whose last block, as Word reads it, is not a paragraph (CR-016): ``cell.no_end_paragraph``.

    Attributes:
        ends_with: what the cell ends with instead: ``"w:tbl"``, or None for a
            cell that holds no block at all.
        hint: Word's repair, and this package's.
    """

    code: ClassVar[str] = "cell.no_end_paragraph"

    ends_with: str | None = None
    hint: str | None = None


@dataclasses.dataclass(frozen=True, slots=True)
class ValueOutOfRange(Finding):
    """An integer attribute Word 365 reads differently from its value (CR-022's read side): ``value.out_of_range``.

    The fields are docx4j-core-ts CR-017's (``ValueOutOfRangeWarning``) in
    snake_case, agreed for all the ports (CR-020 section 4.7).

    Attributes:
        element: the element carrying the attribute, prefixed (``"w:ind"``).
        attribute: the attribute, prefixed (``"w:left"``).
        value: the attribute as written: an ``int`` where it binds as one, else
            the text the parse kept (``"7205001.5"``).
        word_reads: what Word reads (-3959 for 7205001 in ``w:left``); None
            where Word computes a value (a negative column width).
        storage: how Word stores the attribute, the objects package's string
            (``"int16"``, ``"uint16, then 144..31681"``, ``"int16, the sign the
            rule"``, ...): :attr:`docx4j_py.model.word_integers.Storage.token`.
        range: ``(min, max)``, the values Word keeps as written there, in the
            attribute's own unit.
        read_as: where ``w:firstLine`` and ``w:hanging``'s shared field moved
            the value, the attribute Word reads it in (``"w:hanging"``).
        rule: for a line, a row's height or a frame's height, the rule Word
            reads with it (``"exact"``, ``"atLeast"``, ``"auto"``).
    """

    code: ClassVar[str] = "value.out_of_range"
    _keep_none: ClassVar[tuple[str, ...]] = ("word_reads", "range")

    element: str | None = None
    attribute: str | None = None
    value: int | str | None = None
    word_reads: int | None = None
    storage: str | None = None
    range: tuple[int, int] | None = None
    read_as: str | None = None
    rule: str | None = None


# ---------------------------------------------------------------------------
# the names a path is made of
# ---------------------------------------------------------------------------

_prefixed_names: dict[str, str] = {}


def _prefixed(qname: str | None) -> str:
    """``{...main}p`` as ``w:p``; a namespace the prefix table does not know kept whole."""
    if not qname:
        return "?"
    found = _prefixed_names.get(qname)
    if found is not None:
        return found
    shown = qname
    if qname.startswith("{"):
        from docx4j_py.namespaces import NAMESPACE_TO_PREFIX

        uri, _, local = qname[1:].partition("}")
        prefix = NAMESPACE_TO_PREFIX.get(uri)
        if prefix:
            shown = f"{prefix}:{local}"
    _prefixed_names[qname] = shown
    return shown


#: A node's place, a plain tuple for the walk's sake: (its parent's trail, its element name, its 1-based position
#: among its siblings of that name or None for the root, the node). The path's string is built only for a finding.
_Trail = tuple


def _path(trail: _Trail | None, attribute: str | None = None) -> str:
    steps: list[str] = []
    while trail is not None:
        parent, name, index, _node = trail
        steps.append(_prefixed(name) if index is None else f"{_prefixed(name)}[{index}]")
        trail = parent
    path = "/" + "/".join(reversed(steps))
    return f"{path}/@{attribute}" if attribute else path


# ---------------------------------------------------------------------------
# the ranged attributes (CR-022's table)
# ---------------------------------------------------------------------------


class _Ranged(NamedTuple):
    attribute: str  # the attribute's qualified name
    local: str  # its local name, for the shared first-line and hanging field
    shown: str  # prefixed
    holder: str | None  # the element the row's element must be under, qualified
    path: str  # the row's attribute path
    context: str | None  # the sibling attribute the row's range depends on, qualified
    storage: Any  # word_integers.Storage


_ranged_index: dict[str, tuple[_Ranged, ...]] | None = None


def _ranged() -> dict[str, tuple[_Ranged, ...]]:
    """Element qualified name -> the limits table's attributes on it, one per attribute path."""
    global _ranged_index
    if _ranged_index is not None:
        return _ranged_index
    from docx4j_py.model.values import integer_limits
    from docx4j_py.model.word_integers import storage_of

    index: dict[str, list[_Ranged]] = {}
    seen: set[str] = set()
    for row in integer_limits():
        if row.attribute in seen:
            continue
        seen.add(row.attribute)
        name = row.name
        index.setdefault(_q(row.element), []).append(
            _Ranged(
                attribute=_q(name),
                local=name.removeprefix("w:"),
                shown=name,
                holder=_q(row.holder) if row.holder else None,
                path=row.attribute,
                context=_q(row.context[0]) if row.context else None,
                storage=storage_of(row.attribute),
            )
        )
    _ranged_index = {element: tuple(rows) for element, rows in index.items()}
    return _ranged_index


_attribute_vars: dict[type, dict[str, Any]] = {}


def _attributes_of(cls: type) -> dict[str, Any]:
    """A class's attribute vars by qualified name; empty for a class xsdata cannot build."""
    try:
        return _attribute_vars[cls]
    except KeyError:
        pass
    from docx4j_py.runtime import context

    try:
        found = {var.qname: var for var in context().build(cls).get_attribute_vars()}
    except Exception:  # noqa: BLE001 - not a model class
        found = {}
    _attribute_vars[cls] = found
    return found


def _token(value: Any, var: Any) -> str | None:
    """An enumerated sibling attribute's value as Word reads it: the member's value, or None (unset)."""
    if value is None:
        return None
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, str):
        from docx4j_py.model.values import in_enumeration

        for tp in var.types:
            if isinstance(tp, type) and issubclass(tp, Enum):
                member = in_enumeration(value, tp)
                return None if member is None else member.value
        return value
    return None


_limits: dict[tuple[str, str | None], Any] = {}


def _limit(path: str, context: str | None) -> tuple[Any, int, int, frozenset[int]] | None:
    """The row for an attribute path under a context, its range as read and Word's own spellings beside it; cached.

    None where the attribute is not checked: unmeasured, a row height under ``hRule="auto"``, or no row under this
    context (a table width under ``auto``, ``nil`` or no ``w:type``: CR-022 Q3).
    """
    key = (path, context)
    try:
        return _limits[key]
    except KeyError:
        pass
    from docx4j_py.model.values import integer, integer_limit
    from docx4j_py.model.word_integers import read_range

    try:
        row = integer_limit(path, context)
    except LookupError:
        row = None
    found = None
    if row is not None:
        low, high = read_range(row)
        kept = {integer(spelling) for _stored, spelling in row.word_writes}
        if row.sentinel is not None:
            kept.add(row.sentinel)
        found = (row, low, high, frozenset(k for k in kept if k is not None))
    _limits[key] = found
    return found


_RULE_IN_MESSAGE = re.compile(r"\brule\b|clearing exact")


def _describe(reading: Any, local: str) -> str:
    """core-ts's ``describeReading``: ``-3959``, ``atLeast 1``, ``32768 in w:hanging``, ``a width Word computes``."""
    if reading.value is None:
        return "a width Word computes" if "computed" in reading.storage else "not known"
    where = f" in w:{reading.attribute}" if reading.attribute is not None and reading.attribute != local else ""
    rule = f"{reading.rule} " if reading.rule is not None and _RULE_IN_MESSAGE.search(reading.storage) else ""
    return f"{rule}{reading.value}{where}"


# ---------------------------------------------------------------------------
# the walk
# ---------------------------------------------------------------------------

_plans: dict[type, tuple[tuple[str, Any, Any, Any], ...]] = {}


def _plan(cls: type) -> tuple[tuple[str, Any, Any, Any], ...]:
    """A class's child fields, each with the name it holds, its choices and its placement table (or None)."""
    try:
        return _plans[cls]
    except KeyError:
        pass
    from docx4j_py import placement
    from docx4j_py.child import _child_plan
    from docx4j_py.runtime import context

    plan = tuple(
        (name, qname, choices, placement.entry(cls, name) if qname is not None or choices is not None else None)
        for name, qname, choices in _child_plan(cls, context())
    )
    _plans[cls] = plan
    return plan


class _Walker:
    """One walk of one tree: the findings in document order."""

    def __init__(self, part_name: str | None, container: Any = None, prefix: str | None = None) -> None:
        from docx4j_py import placement
        from docx4j_py.child import (
            WILDCARD_FIELD,
            Child,
            _qname_of,
            wildcard_child_name,
        )
        from docx4j_py.runtime import context
        from docx4j_py.traversal import element_name
        from docx4j_py.wml import Tc

        types = context().class_type
        self.any_element = types.any_element
        self.derived_element = types.derived_element
        self.child = Child
        self.wildcard = WILDCARD_FIELD
        self.wildcard_name = wildcard_child_name
        self.qname_of = _qname_of
        self.element_name = element_name
        self.verdict = placement.verdict
        self.written_as = placement.written_as
        self.tc = Tc
        self.ranged = _ranged()
        self.part_name = part_name
        self.container = container
        self.prefix = prefix
        self.findings: list[Finding] = []

    # -- the walk ----------------------------------------------------------

    def run(self, root: Any, root_name: str | None) -> list[Finding]:
        name = root_name or self.element_name(root) or type(root).__name__
        stack: list[tuple[Any, _Trail, str | None, str | None]] = [(root, (None, name, None, root), name, None)]
        while stack:
            node, trail, name, parent_name = stack.pop()
            if isinstance(node, self.any_element):
                children = self._wildcard_items(node, trail)
            else:
                children = self._visit(node, trail, name, parent_name)
            if children:
                children.reverse()
                stack.extend(children)
        return self.findings

    def _visit(self, node: Any, trail: _Trail, name: str | None, parent_name: str | None) -> list:
        cls = node.__class__
        if cls is self.tc:
            self._cell(node, trail)
        ranged = self.ranged.get(name) if name is not None else None
        if ranged:
            self._range(node, cls, ranged, trail, name, parent_name)
        children: list = []
        counts: dict[str, int] = {}
        for field, qname, choices, found in _plan(cls):
            value = getattr(node, field, None)
            if value is None:
                continue
            if isinstance(value, (list, tuple)):
                for index, item in enumerate(value):
                    self._item(node, cls, field, qname, choices, found, item, index, trail, name, counts, children)
            else:
                self._item(node, cls, field, qname, choices, found, value, None, trail, name, counts, children)
        return children

    def _step(self, name: str, counts: dict[str, int], trail: _Trail, node: Any) -> _Trail:
        count = counts.get(name, 0) + 1
        counts[name] = count
        return (trail, name, count, node)

    def _item(
        self,
        holder: Any,
        cls: type,
        field: str,
        qname: Any,
        choices: Any,
        found: Any,
        item: Any,
        index: int | None,
        trail: _Trail,
        holder_name: str | None,
        counts: dict[str, int],
        children: list,
    ) -> None:
        ok = True
        if found is not None and not found.verdicts.get(item.__class__):
            ok = self.verdict(cls, field, item)
        if isinstance(item, self.child):
            if qname is self.wildcard:
                item_name = self.wildcard_name(item)
            elif not ok:
                # a misplaced item is named by what it is, as core-ts's paths name it, not as the serialiser would
                # write it (``<w:CT_Tbl/>`` for a table in a run, its type's name)
                item_name = None
            elif choices is not None:
                item_name = self.qname_of(item, choices, None)
                if item_name is None and found is not None:
                    item_name = self.written_as(found, item.__class__)
            else:
                item_name = qname
            if item_name is None:
                item_name = self.element_name(item) or item.__class__.__name__
            item_trail = self._step(item_name, counts, trail, item)
            if not ok:
                self._misplaced(holder, field, item, index, item_trail)
            children.append((item, item_trail, item_name, holder_name))
            return
        if isinstance(item, self.derived_element):
            item_name = item.qname if isinstance(item.qname, str) else None
            item_trail = self._step(item_name or "?", counts, trail, item)
            if not ok:
                self._misplaced(holder, field, item, index, item_trail)
            if isinstance(item.value, self.child):
                children.append((item.value, item_trail, item_name, holder_name))
            return
        if isinstance(item, self.any_element):
            item_name = item.qname if isinstance(item.qname, str) else None
            item_trail = self._step(item_name or "?", counts, trail, item)
            if not ok:
                self._misplaced(holder, field, item, index, item_trail)
            children.append((item, item_trail, item_name, holder_name))
            return
        if not ok:  # a string, None, a value: the holder's path, as core-ts gives a string's
            self._misplaced(holder, field, item, index, trail)

    def _wildcard_items(self, node: Any, trail: _Trail) -> list:
        """An ``AnyElement``'s children: wildcard trees and the typed objects in them, named as they are written."""
        children: list = []
        counts: dict[str, int] = {}
        for item in node.children or ():
            if isinstance(item, self.child):
                name = self.wildcard_name(item) or item.__class__.__name__
                children.append((item, self._step(name, counts, trail, item), name, node.qname))
            elif isinstance(item, self.any_element):
                name = item.qname if isinstance(item.qname, str) else "?"
                children.append((item, self._step(name, counts, trail, item), name, node.qname))
        return children

    # -- the checks --------------------------------------------------------

    def _misplaced(self, holder: Any, field: str, item: Any, index: int | None, trail: _Trail) -> None:
        from docx4j_py import placement

        error = placement.refusal(holder, field, item)
        self.findings.append(
            MisplacedContent(
                message=error.message,
                part_name=self.part_name,
                path=_path(trail),
                address=self._address(trail),
                node=item,
                holder=type(holder).__name__,
                field=field,
                index=index,
                item=type(item).__name__,
                element=_prefixed(error.element) if error.element else None,
                hint=error.hint,
            )
        )

    def _cell(self, tc: Any, trail: _Trail) -> None:
        from docx4j_py.model.content.table_rules import _last_block, paragraph_end_of
        from docx4j_py.model.content.text_model import block_children_of
        from docx4j_py.wml import P, Tbl

        if paragraph_end_of(tc) is not None:
            return
        items = tc.content
        ends_with = None
        for _ in range(32):
            last = _last_block(items)
            if last is None or isinstance(last, P):
                break
            if isinstance(last, Tbl):
                ends_with = "w:tbl"
                break
            items = block_children_of(last)
        what = "a table (`w:tbl`)" if ends_with else "no paragraph or table"
        self.findings.append(
            CellWithoutEndParagraph(
                message=f"a table cell must end with a paragraph, as Word reads it; this one ends with {what}",
                part_name=self.part_name,
                path=_path(trail),
                address=self._address(trail),
                node=tc,
                ends_with=ends_with,
                hint=(
                    "Word opens the file as unreadable content, and its repair appends an empty paragraph to the "
                    "cell: `table_rules.end_with_paragraph(tc)` makes the same repair"
                ),
            )
        )

    def _range(
        self, node: Any, cls: type, ranged: tuple[_Ranged, ...], trail: _Trail, name: str, parent_name: str | None
    ) -> None:
        from docx4j_py.model.values import integer
        from docx4j_py.model.word_integers import reading

        attributes = _attributes_of(cls)
        for entry in ranged:
            var = attributes.get(entry.attribute)
            if var is None:
                continue
            if entry.holder is not None and parent_name != entry.holder:
                continue
            raw = getattr(node, var.name, None)
            if raw is None:
                continue
            value = integer(raw)
            if value is None:
                continue  # not a number Word reads: CR-021's, dropped by the save's normal form
            context = None
            if entry.context is not None:
                sibling = attributes.get(entry.context)
                context = None if sibling is None else _token(getattr(node, sibling.name, None), sibling)
            limit = _limit(entry.path, context)
            if limit is None:
                continue  # unmeasured, a row height under auto, or no row under this context (a width under auto)
            _row, low, high, kept = limit
            if low <= value <= high or value in kept:
                continue  # in range, or Word's own spelling (-1 for 65,535, a column space's sentinel)
            storage = entry.storage
            rule = None
            if storage.sign_rule is not None:
                sibling = attributes.get(_q(storage.sign_rule))
                rule = None if sibling is None else _token(getattr(node, sibling.name, None), sibling)
            read = reading(storage, entry.local, value, rule)
            shown = _prefixed(name)
            path = _path(trail, entry.shown)
            prefix = f"{self.part_name}: " if self.part_name else ""
            self.findings.append(
                ValueOutOfRange(
                    message=(
                        f"{prefix}{path}={raw} is past what Word 365 stores there "
                        f"({read.storage}, {low} to {high}); Word reads it as {_describe(read, entry.local)}"
                    ),
                    part_name=self.part_name,
                    path=path,
                    address=self._address(trail),
                    node=node,
                    element=shown,
                    attribute=entry.shown,
                    value=raw,
                    word_reads=read.value,
                    storage=read.storage,
                    range=(low, high),
                    read_as=f"w:{read.attribute}" if read.attribute is not None and read.attribute != entry.local else None,
                    rule=read.rule,
                )
            )

    # -- the address -------------------------------------------------------

    def _address(self, trail: _Trail) -> str | None:
        """The content API's address of the innermost block on the trail: ``address_of``'s, walked down the trail."""
        if self.container is None:
            return None
        from docx4j_py.model.content.addresses import para_id_address
        from docx4j_py.model.content.text_model import block_children_of

        chain: list[Any] = []
        current: _Trail | None = trail
        while current is not None:
            chain.append(current[3])
            current = current[0]
        chain.reverse()
        start = next((i for i, node in enumerate(chain) if node is self.container), None)
        if start is None:
            return None
        indices: list[int] = []
        block = None
        for parent, child in zip(chain[start:], chain[start + 1 :], strict=False):
            items = block_children_of(parent)
            if items is None:
                continue
            position = next((i for i, item in enumerate(items) if item is child), None)
            if position is None:
                continue
            indices.append(position)
            block = child
        if block is None:
            return None
        para_id = getattr(block, "para_id", None)
        if para_id:
            return para_id_address(para_id)
        return "/".join([self.prefix or "body", *(str(i) for i in indices)])


# ---------------------------------------------------------------------------
# trees, parts and packages
# ---------------------------------------------------------------------------


def findings_in(root: Any, *, part_name: str | None = None, root_name: str | None = None) -> list[Finding]:
    """Every finding in a tree, in document order (a part's own walk is :func:`part_findings`)."""
    return _Walker(part_name).run(root, root_name)


#: The part kinds that have a body, and so content API addresses (``addresses._PART_PREFIXES``).
_BODY_KINDS = ("maindocument", "header", "footer", "footnotes", "endnotes", "comments")


def _body_of(part: Any, root: Any) -> tuple[Any, str | None]:
    """The block container a part's addresses count from, and their prefix; (None, None) for a part with no body."""
    from docx4j_py.model.content.addresses import part_kind, prefix_for_part
    from docx4j_py.model.content.text_model import block_list_of

    if not part_kind(part).lower().startswith(_BODY_KINDS):
        return None, None
    container = getattr(root, "body", None)
    if container is None:
        container = root
    if block_list_of(container) is None:
        return None, None
    return container, prefix_for_part(part)


def _tree_of(part: Any) -> Any:
    """The part's live tree where it is unmarshalled, else its private one: nothing is marked for re-marshalling."""
    return part.contents if part.is_unmarshalled else part.read_contents()


def part_findings(part: Any, tree: Any = None) -> list[Finding]:
    """Every finding in one XML part (`tree`, else its own tree), with the part's name and its addresses."""
    from docx4j_py.traversal import element_name

    root = _tree_of(part) if tree is None else tree
    if root is None:
        return []
    container, prefix = _body_of(part, root)
    walker = _Walker(str(part.part_name), container, prefix)
    return walker.run(root, getattr(part, "root_name", None) or element_name(root))


def _xml_parts(package: Any, unmarshalled_only: bool) -> list[Any]:
    from docx4j_py.openpackaging.parts.xml_part import XmlPart

    out = []
    for part in package.parts.parts():
        if not isinstance(part, XmlPart) or part.model_class_path is None:
            continue
        if unmarshalled_only and not part.is_unmarshalled:
            continue
        if not (part.is_unmarshalled or part.is_loaded or _in_store(part)):
            continue  # nothing to read: a part with no content and no source, which no save can write either
        out.append(part)
    return out


def _in_store(part: Any) -> bool:
    """Whether the package's source container holds the part's bytes (``Part._source_bytes``'s test, unread)."""
    from urllib.parse import unquote

    package = getattr(part, "package", None)
    store = getattr(package, "source_part_store", None)
    if store is None:
        return False
    name = part.part_name.store_name
    return bool(store.has(name) or store.has(unquote(name)))


def misplaced(root: Any, *, unmarshalled_only: bool = False) -> list[Finding]:
    """Every finding in a tree, a part or a package (CR-020 D7, Q5): what :func:`docx4j_py.traversal.misplaced` is.

    Args:
        root: a model object (its tree is walked, paths from it), an
            ``XmlPart``, or a package (each of its typed XML parts, in the order
            they were added).
        unmarshalled_only: for a package, only the parts already unmarshalled:
            those a save re-marshals. Otherwise a part never unmarshalled is
            read privately and stays as it was, written byte for byte.
    """
    from docx4j_py.openpackaging.parts.xml_part import XmlPart

    if isinstance(root, XmlPart):
        return part_findings(root)
    parts = getattr(root, "parts", None)
    if parts is not None and hasattr(parts, "parts") and hasattr(root, "relationships_part"):
        out: list[Finding] = []
        for part in _xml_parts(root, unmarshalled_only):
            out.extend(part_findings(part))
        return out
    return findings_in(root)


# ---------------------------------------------------------------------------
# the save's check (CR-020 Q5)
# ---------------------------------------------------------------------------


def _key(finding: Finding) -> tuple:
    """What a finding is, apart from where: so that one in the source is known again after an edit moved it."""
    if isinstance(finding, ValueOutOfRange):
        return (finding.code, finding.element, finding.attribute, finding.value)
    if isinstance(finding, CellWithoutEndParagraph):
        from docx4j_py.traversal import text_of

        return (finding.code, finding.ends_with, text_of(finding.node, mce="all"))
    return (finding.code, finding.path)


def check_for_save(package: Any) -> list[Finding]:
    """``save(check=True)``'s check: raise for what this session introduced; return what the source already held.

    Over every part the save re-marshals (the unmarshalled ones; a part never
    unmarshalled is written byte for byte, as it was read):

    * **misplaced content raises**, always: a parse never produces it (CR-020
      section 2.1), so it is code's;
    * **a cell or an out-of-range value raises where the session introduced
      it**: one the part's source already held, found again by what it is
      (the attribute and its value; the cell's text and what it ends with) in
      a fresh parse of the source, made only when the part has such a finding,
      is **carried**: returned, recorded on the save's report
      (``package.last_save.findings``), and saved as it is. Nothing is
      repaired on a save (CLAUDE.md; CR-016 D4).

    Raises:
        SaveCheckError: the findings this session introduced, the first's code
            as the error's.
    """
    carried: list[Finding] = []
    introduced: list[Finding] = []
    for part in _xml_parts(package, unmarshalled_only=True):
        found = part_findings(part)
        if not found:
            continue
        rest = []
        for finding in found:
            (introduced if isinstance(finding, MisplacedContent) else rest).append(finding)
        if not rest:
            continue
        source = part.source_tree() if hasattr(part, "source_tree") else None
        held: Counter = Counter()
        if source is not None:
            held.update(_key(f) for f in part_findings(part, source) if not isinstance(f, MisplacedContent))
        for finding in rest:
            key = _key(finding)
            if held[key] > 0:
                held[key] -= 1
                carried.append(finding)
            else:
                introduced.append(finding)
    if introduced:
        from docx4j_py.model.content.errors import SaveCheckError

        raise SaveCheckError.of(introduced)
    return carried
