"""Addresses that survive edits: the three forms of CR-003 section 3.4.

An agent does not hold object references across tool calls; it holds a string.
Three forms, accepted wherever a block-level target is and reported everywhere a
view is:

**paraId first.** ``"w14:5A2B1C3D"``, the ``w14:paraId`` Word writes on every
paragraph it saves. Stable across every edit, inserts before it included.
:attr:`~docx4j_py.model.content.paragraph.Paragraph.address` reports it when the
paragraph has one.

**Ordinal second.** ``"body/3"``, ``"body/4/0/1/0"`` (a paragraph in the second
cell of the first row of the fifth block), ``"header:rId3/0"``,
``"footer:rId5/2"``, ``"footnotes/1/0"``, ``"comments/2/0"``. One index per
:func:`~docx4j_py.model.content.text_model.block_children_of` step, so a table's
rows, their cells and a content control's children are all just indices and the
four ``w:sdtContent`` levels are invisible (CR-003 section 4).
:attr:`~docx4j_py.model.content.paragraph.Paragraph.ordinal` always reports it.
Stable until an insert or a delete before it, and
:class:`~docx4j_py.model.content.reports.ChangeReport` says which ordinals moved.

**Text third.** ``paragraph_at(contains="Chapter 1")``, the first match, which is
what the MCP servers' ``near_text`` tools do.

Every :class:`~docx4j_py.model.content.errors.AddressError` names the nearest
surviving address --- the same prefix, the closest ordinal --- and says to call
``outline()``.

The prefix is the part: ``"body"`` for the main document, ``"header:<relId>"``
and ``"footer:<relId>"`` (there may be several of each), and ``"footnotes"``,
``"endnotes"`` and ``"comments"`` (there is at most one of each). A package
resolves any of them; a body resolves its own and hands the rest to its package.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from docx4j_py.child import iter_children
from docx4j_py.model.content.errors import AddressError, ContentError
from docx4j_py.model.content.text_model import block_children_of
from docx4j_py.wml import P, Tr

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = [
    "PARA_ID_PREFIX",
    "address_of",
    "assign_para_ids",
    "body_for_prefix",
    "element_at",
    "ensure_para_ids",
    "identified",
    "nearest_address",
    "ordinal_of",
    "package_bodies",
    "para_id_address",
    "paragraph_at",
    "path_of",
    "record_para_ids",
    "resolve_path",
]

#: What a paraId address begins with. ``w14:`` because that is the element's
#: namespace prefix, and because it cannot be confused with an ordinal.
PARA_ID_PREFIX = "w14:"

#: How far up the parent chain an ordinal is looked for before giving up. A
#: document nested this deep is a loop, not a document.
_MAX_DEPTH = 64

#: How many paraIds are drawn before the generator is declared stuck.
_MAX_ATTEMPTS = 64


# ---------------------------------------------------------------------------
# ordinals
# ---------------------------------------------------------------------------


def _index_of(items: list, element: Any) -> int | None:
    """The index of `element` in `items`, by identity, or None."""
    for index, item in enumerate(items):
        if item is element:
            return index
    return None


def path_of(body: Body, element: Any) -> list[int] | None:
    """The indices from `body`'s container down to `element`, or None.

    Walked **up** the parent pointers rather than down the tree, so it costs the
    nesting depth and not the document. A level whose block list does not hold
    the child is transparent and contributes no index, which is what makes the
    four ``w:sdtContent`` classes invisible to a path.
    """
    if element is body.container:
        return []
    path: list[int] = []
    current = element
    for _ in range(_MAX_DEPTH):
        parent = getattr(current, "parent", None)
        if parent is None:
            return None
        items = block_children_of(parent)
        if items is not None:
            index = _index_of(items, current)
            if index is not None:
                path.append(index)
        if parent is body.container:
            path.reverse()
            return path
        current = parent
    return None


def ordinal_of(body: Body, element: Any) -> str | None:
    """The ordinal address of a block in `body` (``"body/4/0/1/0"``), or None."""
    path = path_of(body, element)
    if path is None:
        return None
    return "/".join([body.prefix, *(str(index) for index in path)])


def address_of(body: Body, target: Any) -> str:
    """The address of any block: its paraId when it has one, else its ordinal.

    Args:
        target: a :class:`~docx4j_py.model.content.paragraph.Paragraph`, a
            :class:`~docx4j_py.model.content.body.Block` (which is what a table
            or a content control is until Phase C), or the element itself.

    Raises:
        AddressError: the block is not in this body.
    """
    element = getattr(target, "element", target)
    para_id = getattr(element, "para_id", None)
    if para_id:
        return para_id_address(para_id)
    ordinal = ordinal_of(body, element)
    if ordinal is None:
        raise AddressError(
            f"that {type(element).__name__} is not in {body.prefix}",
            code="address.not_here",
            hint="call outline() on the body that holds it, or address_of on its own body",
        )
    return ordinal


def para_id_address(para_id: str) -> str:
    """``"5A2B1C3D"`` as the address ``"w14:5A2B1C3D"``."""
    return f"{PARA_ID_PREFIX}{para_id}"


# ---------------------------------------------------------------------------
# resolving
# ---------------------------------------------------------------------------


def resolve_path(body: Body, path: list[int]) -> tuple[Any, list] | None:
    """The element and the live list at an ordinal path, or None.

    Each index is a step through
    :func:`~docx4j_py.model.content.text_model.block_children_of`, so a table's
    rows, a row's cells and a control's content are all one index each.
    """
    element: Any = None
    container: list | None = None
    current: Any = body.container
    for index in path:
        items = block_children_of(current)
        if items is None or index < 0 or index >= len(items):
            return None
        container = items
        element = items[index]
        current = element
    if element is None or container is None:
        return None
    return element, container


def _parse_ordinal(body: Body, address: str) -> list[int] | None:
    """The indices of an ordinal address under this body's prefix, or None."""
    if address == body.prefix:
        return []
    if not address.startswith(body.prefix + "/"):
        return None
    tail = address[len(body.prefix) + 1 :]
    path: list[int] = []
    for segment in tail.split("/"):
        if not segment.isdigit():
            return None
        path.append(int(segment))
    return path


def _paragraph_by_para_id(body: Body, para_id: str) -> Paragraph | None:
    wanted = para_id.lower()
    for paragraph in body.iter_paragraphs():
        current = paragraph.para_id
        if current and current.lower() == wanted:
            return paragraph
    return None


def _paragraph_containing(body: Body, text: str) -> Paragraph | None:
    for paragraph in body.iter_paragraphs():
        if text in paragraph.text:
            return paragraph
    return None


def element_at(body: Body, address: str) -> Any:
    """The view of the block at an address. What ``body.element_at`` is.

    Accepts a paraId (``"w14:5A2B1C3D"``) and an ordinal (``"body/3"``). An
    ordinal whose prefix names another part is resolved through the package,
    so ``pkg.body.element_at("header:rId8/0")`` works.

    Raises:
        AddressError: nothing lives there; the message names the nearest
            surviving address and says to call ``outline()``.
    """
    found = _element_at_or_none(body, address)
    if found is None:
        raise _not_found(body, address)
    return found


def _element_at_or_none(body: Body, address: str) -> Any:
    if not isinstance(address, str) or not address:
        raise ContentError(
            f"an address is a string such as 'body/3' or 'w14:5A2B1C3D', not {address!r}",
            code="address.invalid",
            hint="call outline() to list the current addresses",
        )
    if address.startswith(PARA_ID_PREFIX):
        wanted = address[len(PARA_ID_PREFIX) :]
        paragraph = _paragraph_by_para_id(body, wanted)
        if paragraph is not None:
            return paragraph
        for other in _sibling_bodies(body):
            paragraph = _paragraph_by_para_id(other, wanted)
            if paragraph is not None:
                return paragraph
        return None

    target = body
    path = _parse_ordinal(body, address)
    if path is None:
        target = _body_for_address(body, address)
        if target is None:
            return None
        path = _parse_ordinal(target, address)
        if path is None:
            return None
    found = resolve_path(target, path)
    if found is None:
        return None
    element, _container = found
    return target.view_for(element)


def paragraph_at(
    body: Body,
    address: str | None = None,
    *,
    contains: str | None = None,
    para_id: str | None = None,
) -> Paragraph:
    """The paragraph at an address, holding some text, or with a paraId.

    Exactly one of the three has to be given; they are CR-003 section 3.4's
    three address forms, in its order of preference.

    Raises:
        AddressError: nothing matches, or what does is not a paragraph.
    """
    from docx4j_py.model.content.paragraph import Paragraph

    given = [x for x in (address, contains, para_id) if x is not None]
    if len(given) != 1:
        raise ContentError(
            "paragraph_at takes exactly one of address, contains= or para_id=",
            code="address.ambiguous",
            hint="paragraph_at('body/3'), paragraph_at(contains='Chapter 1') "
            "or paragraph_at(para_id='5A2B1C3D')",
        )

    if para_id is not None:
        found = _element_at_or_none(body, para_id_address(para_id))
        if found is None:
            raise _not_found(body, para_id_address(para_id))
    elif contains is not None:
        found = _paragraph_containing(body, contains)
        if found is None:
            for other in _sibling_bodies(body):
                found = _paragraph_containing(other, contains)
                if found is not None:
                    break
        if found is None:
            raise AddressError(
                f"no paragraph contains {contains!r}",
                code="address.no_match",
                hint="call find() for a case-insensitive search, or outline() for the text",
            )
    else:
        found = element_at(body, address)  # type: ignore[arg-type]

    if not isinstance(found, Paragraph):
        raise AddressError(
            f"{address!r} is a {getattr(found, 'name', type(found).__name__)}, not a paragraph",
            code="address.not_a_paragraph",
            hint="use element_at() for a table or a content control",
        )
    return found


# ---------------------------------------------------------------------------
# the nearest surviving address
# ---------------------------------------------------------------------------


def _addresses_under(body: Body) -> list[tuple[list[int], str]]:
    """Every ordinal path in a body, with its address, in document order."""
    out: list[tuple[list[int], str]] = []

    def visit(items: list, path: list[int]) -> None:
        for index, item in enumerate(items):
            here = [*path, index]
            out.append((here, "/".join([body.prefix, *(str(i) for i in here)])))
            children = block_children_of(item)
            if children is not None:
                visit(children, here)

    visit(body.content, [])
    return out


def _score(wanted: list[int], candidate: list[int]) -> tuple[int, int, int]:
    """How close `candidate` is to `wanted`: bigger is closer."""
    shared = 0
    for a, b in zip(wanted, candidate):
        if a != b:
            break
        shared += 1
    gap = 0
    if shared < len(wanted) and shared < len(candidate):
        gap = abs(wanted[shared] - candidate[shared])
    # the deepest shared path first, then a block at the same depth --- an
    # address one level up is the container, not a neighbour --- and then the
    # smallest difference at the index that differs
    return (shared, -abs(len(candidate) - len(wanted)), -gap)


def nearest_address(body: Body, address: str) -> str | None:
    """The surviving address closest to one that is gone, or None.

    The same prefix and the closest ordinal, which is what CR-003 section 3.4
    asks an :class:`~docx4j_py.model.content.errors.AddressError` to name: the
    deepest shared path, then the smallest difference at the first index that
    differs.
    """
    target = body
    path = _parse_ordinal(body, address)
    if path is None:
        if address.startswith(PARA_ID_PREFIX):
            return None
        other = _body_for_address(body, address)
        if other is None:
            return None
        target = other
        path = _parse_ordinal(target, address)
        if path is None:
            return None
    candidates = _addresses_under(target)
    if not candidates:
        return None
    best = max(candidates, key=lambda item: _score(path, item[0]))
    return best[1]


#: The text model's segments after an anchor (WD REQ-519): what ``outline()`` and
#: ``to_markdown(addresses=True)`` report for a text box's paragraph, and the content API
#: does not resolve (CR-011 Q9).
_MODEL_SEGMENTS = ("/textbox:", "/chart:", "/smartart:", "/object:", "/altchunk:")
#: The text model's story prefixes the content API spells otherwise (``footnotes/1/0``) or not at all.
_MODEL_STORIES = ("footnote:", "endnote:", "comment:", "glossary:")


def _not_found(body: Body, address: str) -> AddressError:
    """The error CR-003 section 3.4 asks for, nearest address and all."""
    if any(segment in address for segment in _MODEL_SEGMENTS) or address.startswith(_MODEL_STORIES):
        return AddressError(
            f"{address!r} is the text model's address of content the content API does not resolve "
            "(a text box's, a chart's, a note's paragraph); it is for reading (CR-011 Q9)",
            code="address.read_only",
            hint="read it with outline(), to_markdown() or docx4j_py.model.text.extract(); "
            "edit a text box through its anchor, a note through its part's body (footnotes/<n>/<i>)",
        )
    nearest = nearest_address(body, address)
    if address.startswith(PARA_ID_PREFIX):
        message = f"no paragraph has the paraId {address[len(PARA_ID_PREFIX) :]!r}"
        hint = "call outline() to list current addresses; paraIds are reported as para_id"
    elif nearest is not None:
        message = f"nothing at {address!r}; the nearest surviving address is {nearest!r}"
        hint = f"use {nearest!r}, or call outline() to list current addresses"
    else:
        message = f"nothing at {address!r}"
        hint = "call outline() to list current addresses"
    return AddressError(message, code="address.not_found", hint=hint)


# ---------------------------------------------------------------------------
# the bodies of a package
# ---------------------------------------------------------------------------

#: Part class name prefix -> the address prefix, and whether it is per-relationship.
_PART_PREFIXES: tuple[tuple[str, str, bool], ...] = (
    ("maindocument", "body", False),
    ("header", "header", True),
    ("footer", "footer", True),
    ("footnotes", "footnotes", False),
    ("endnotes", "endnotes", False),
    ("comments", "comments", False),
)


def part_kind(part: Any) -> str:
    """The class name a prefix is decided from, a trial part's original included."""
    return type(getattr(part, "_wrapped", part)).__name__


def prefix_for_part(part: Any) -> str:
    """The address prefix a part's body reports under (``"header:rId8"``)."""
    name = part_kind(part).lower()
    for start, prefix, per_relationship in _PART_PREFIXES:
        if not name.startswith(start):
            continue
        if not per_relationship:
            return prefix
        rel_id = _rel_id_of(part)
        return f"{prefix}:{rel_id}" if rel_id else prefix
    return "body"


def _rel_id_of(part: Any) -> str | None:
    """The relationship id a header or footer is reached by, when there is one."""
    for relationship in getattr(part, "source_relationships", None) or ():
        rel_id = getattr(relationship, "id", None)
        if rel_id:
            return str(rel_id)
    return None


def package_parts(package: Any) -> Iterator[Any]:
    """Every part of a package that has a body, main document part first."""
    main = getattr(package, "main_document_part", None)
    if main is None:
        get_main = getattr(package, "get_main_document_part", None)
        if get_main is not None:
            try:
                main = get_main()
            except Exception:  # noqa: BLE001 - a package with no main part has no bodies
                main = None
    if main is None:
        return
    yield main
    for name in ("header_parts", "footer_parts"):
        getter = getattr(package, name, None)
        if getter is not None:
            yield from getter()
    for name in ("footnotes_part", "endnotes_part", "comments_part"):
        part = getattr(main, name, None)
        if part is not None:
            yield part


def package_bodies(package: Any) -> Iterator[Body]:
    """A :class:`~docx4j_py.model.content.body.Body` per part that has one."""
    from docx4j_py.model.content.body import body_of

    for part in package_parts(package):
        try:
            yield body_of(part)
        except ContentError:  # pragma: no cover - a notes part with nothing in it
            continue


def body_for_prefix(package: Any, prefix: str) -> Body | None:
    """The body whose address prefix is `prefix`, or None."""
    for body in package_bodies(package):
        if body.prefix == prefix:
            return body
    return None


def _sibling_bodies(body: Body) -> Iterator[Body]:
    """The package's other bodies, for an address this one cannot resolve."""
    package = body.package
    if package is None:
        return
    for other in package_bodies(package):
        if other.prefix != body.prefix:
            yield other


def _body_for_address(body: Body, address: str) -> Body | None:
    """The body an ordinal address belongs to, through the package."""
    prefix = address.split("/", 1)[0]
    if prefix == body.prefix:
        return body
    package = body.package
    if package is None:
        return None
    return body_for_prefix(package, prefix)


# ---------------------------------------------------------------------------
# paragraph ids (CR-003 section 3.4, determinism)
# ---------------------------------------------------------------------------


def identified(root: Any) -> Iterator[Any]:
    """Every ``w:p`` and ``w:tr`` at or under `root`, in document order, at any depth.

    In the branch Word reads, so not a ``mc:Fallback``: Word reads a ``mc:Choice``'s
    id and writes it over the Fallback's, and never reads the Fallback's own
    (CR-038 section 8). A text box's paragraphs are reached, in a run as they are.
    """
    if isinstance(root, (P, Tr)):
        yield root
    for _name, child in iter_children(root):
        yield from identified(child)


def _story_of(body: Body) -> Any:
    """The tree of the part `body` stands in: the whole story, since a cell's own body sees only that cell."""
    contents = getattr(body.part, "contents", None) if body.part is not None else None
    return contents if contents is not None else body.container


def para_ids_of(body: Body) -> set[str]:
    """Every ``w14:paraId`` in the part `body` stands in, rows' and text boxes' too, upper-cased as Word writes them."""
    return {item.para_id.upper() for item in identified(_story_of(body)) if item.para_id}


def _taken(package: Any, body: Body) -> set[str]:
    """The ids already used in this part, computed once and then kept in step.

    A set on the package, keyed by the part, rather than a walk per insert:
    assigning an id to each of two thousand new paragraphs would otherwise be
    quadratic. Per **part** and not per package, because looking at every part
    would unmarshal the headers, the footers and the notes of a document that
    only ever had its body edited, and CR-002's promise is that an untouched part
    is written back byte for byte. Two parts of one document colliding on a
    31-bit id is not a risk worth that. It isn't reduced on a delete, so an id in
    it may no longer be held: :func:`assign_para_ids` checks before replacing one.
    """
    cache = getattr(package, "_para_ids_taken", None)
    if not isinstance(cache, dict):
        cache = {}
        package._para_ids_taken = cache
    key = str(body.part.part_name) if body.part is not None else body.prefix
    taken = cache.get(key)
    if taken is None:
        taken = para_ids_of(body)
        cache[key] = taken
    return taken


def assigns_para_ids(package: Any, body: Body) -> bool:
    """Whether new paragraphs and rows in this body should be given a ``w14:paraId``.

    CR-003 section 3.4: "new paragraphs get one when the document already uses
    them (and always in a created document)". ``package.assigns_para_ids`` is
    the override, set to True by ``create_package`` and settable by a caller who
    wants stable handles on a document that has none. The whole part decides, not
    the body (CR-038 section 3.2), and a header or footer made in this session
    takes the main part's answer, since a new footer with no ids costs a file with
    no comments every id it holds (section 8).
    """
    if package is None:
        return False
    choice = getattr(package, "assigns_para_ids", None)
    if choice is not None:
        return bool(choice)
    return bool(_taken(package, body)) or (_made_here(body.part) and _main_uses_ids(package))


def _made_here(part: Any) -> bool:
    """Whether `part` is a header or footer this session made. The glossary isn't one: Word writes no ids there."""
    from docx4j_py.openpackaging.parts.wml import FooterPart, HeaderPart

    return isinstance(part, (HeaderPart, FooterPart)) and part._source_bytes() is None


def _main_uses_ids(package: Any) -> bool:
    """Whether the main part holds an id, read privately, so that asking marks nothing."""
    main = getattr(package, "main_document_part", None)
    tree = main.read_contents() if main is not None else None
    return tree is not None and any(item.para_id for item in identified(tree))


def _fresh_id(package: Any, taken: set[str]) -> str | None:
    """A new id from the package's seedable generator, so that the same seed and calls give the same bytes."""
    generator = package.id_generator(derive_from=taken)
    for _attempt in range(_MAX_ATTEMPTS):
        candidate = f"{generator.randrange(1, 0x7FFFFFFF):08X}"
        if candidate not in taken:
            taken.add(candidate)
            return candidate
    return None  # pragma: no cover - 64 collisions in a 2^31 space


def assign_para_ids(body: Body, element: Any) -> list[str]:
    """Give every ``w:p`` and ``w:tr`` at or under `element`, just added, the id it needs; returns the ids given.

    Word renumbers a file's every id if one is missing (CR-038 section 8), so this
    walks all of what was added. An id already there stays where it's free. One
    the part holds elsewhere is replaced, as Word itself gives the later of two a
    new id, so that a copy never answers its original's address: the set may be
    stale, so the tree decides, and a paragraph put back in its place keeps its id.
    """
    package = body.package
    if not assigns_para_ids(package, body):
        return []
    taken = _taken(package, body)
    given: list[str] = []
    held: Counter[str] | None = None
    for item in identified(element):
        if item.para_id:
            key = item.para_id.upper()
            if key not in taken:
                taken.add(key)
                continue
            if held is None:
                held = Counter(found.para_id.upper() for found in identified(_story_of(body)) if found.para_id)
            if held[key] <= 1:
                continue
            held[key] -= 1
        fresh = _fresh_id(package, taken)
        if fresh is not None:
            item.para_id = fresh
            given.append(fresh)
    return given


def record_para_ids(change: Any, body: Body, element: Any) -> None:
    """:func:`assign_para_ids`, each id recorded on `change`."""
    for para_id in assign_para_ids(body, element):
        change.created(para_id)


def ensure_para_ids(body: Body) -> list[str]:
    """Give every paragraph and row that lacks one a ``w14:paraId``. Extension.

    For an agent that wants stable handles on a legacy document: after this,
    every paragraph's :attr:`~docx4j_py.model.content.paragraph.Paragraph.address`
    is a paraId and survives every insert and delete before it, and Word keeps
    the ids, since it renumbers them all if a paragraph or a row has none (CR-038
    section 8). Text boxes are reached. A document's own repeated ids stay.

    **It re-marshals the part.** Writing an attribute on every paragraph means
    ``document.xml`` is rebuilt from the tree on save rather than copied byte
    for byte, so a document whose bytes matter should be left alone. The ids
    come from the package's seedable generator, so two runs over the same
    document with the same seed produce the same ones.

    Returns:
        The ids assigned, in document order.
    """
    package = body.package
    if package is None:
        raise ContentError(
            "this body has no package, so there is no id generator",
            code="paraid.no_package",
            hint="call ensure_para_ids on a body reached from a package",
        )
    taken = _taken(package, body)
    assigned: list[str] = []
    for item in identified(body.container):
        if not item.para_id:
            fresh = _fresh_id(package, taken)
            if fresh is not None:
                item.para_id = fresh
                assigned.append(fresh)
    return assigned
