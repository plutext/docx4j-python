"""``Body``: Office JS ``Word.Body`` over any container of block-level content.

CR-003 section 3.2. One class serves the main document's ``w:body``, a header,
a footer, the footnotes, the endnotes, the comments, a table cell and a content
control's ``w:sdtContent``, because in this model they all keep their children
in one list and :func:`~docx4j_py.model.content.text_model.block_children_of`
finds it. ``Body.sub(container, prefix)`` is how a nested one is made, and the
prefix it carries is the address prefix Phase D will report from.

``Body`` is a :class:`collections.abc.Sequence` of its block-level children
(``len(body)``, ``body[3]``, ``for block in body``), and ``"Chapter 2" in body``
is a text search, as CR-003 section 3.1 asks. ``body.content`` is still the live
``ChildList``: nothing is wrapped, and a caller who appends to it keeps working
(section 3.13).

The part registration of section 5 happens in this package's ``__init__``:
``XmlPart.body`` and ``WordprocessingMLPackage.body`` are installed from here,
and the parts layer never imports this one.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterator, Sequence
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, link_parents
from docx4j_py.model.content.enums import (
    BodyLocation,
    BreakTypeValue,
    Location,
    RangeLocation,
    TextLocation,
    TextView,
)
from docx4j_py.model.content.errors import ContentError, InvalidTargetError
from docx4j_py.model.content.paragraph import Paragraph
from docx4j_py.model.content.text_model import block_children_of, block_list_of
from docx4j_py.namespaces import PREFIXES
from docx4j_py.runtime import context
from docx4j_py.traversal import element_name
from docx4j_py.wml import P, R, br, r, to_xml, wml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.range import Range

__all__ = ["Block", "Body", "body_of"]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{_W}}}p"

_PREFIX_FOR: dict[str, str] = {uri: prefix for prefix, uri in PREFIXES.items()}


def qualified(qname: str | None) -> str:
    """``{namespace}local`` as ``w:local``, for a message a human reads."""
    if not qname:
        return "?"
    if not qname.startswith("{"):
        return qname
    uri, _, local = qname[1:].partition("}")
    prefix = _PREFIX_FOR.get(uri)
    return f"{prefix}:{local}" if prefix else local


@dataclasses.dataclass(frozen=True, slots=True)
class Block:
    """A block-level child that is not a paragraph: a table, a content control.

    ``Table`` and ``ContentControl`` views are Phase C; until then a body hands
    out this, which carries the element and the list holding it so that the
    tree is one attribute away.
    """

    element: Any
    """The element: a ``w:tbl``, a ``w:sdt``, a ``w:customXml``, ..."""
    container: list
    """The live list holding it."""

    @property
    def name(self) -> str:
        """The element's qualified name, as ``w:tbl``."""
        return qualified(element_name(self.element))

    @property
    def text(self) -> str:
        """The block's text, as docx4j's ``TextUtils`` reads it."""
        from docx4j_py.traversal import text_of

        return text_of(self.element)

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the element's name and its text."""
        return {"kind": self.name, "text": self.text}

    def __repr__(self) -> str:
        """``<Block w:tbl>``."""
        return f"<Block {self.name}>"


class Body(Sequence):
    """A subset of Office JS ``Word.Body`` over a container of block content."""

    __slots__ = ("container", "package", "part", "prefix")

    def __init__(
        self,
        part: Any,
        container: Any,
        prefix: str = "body",
        package: Any = None,
    ) -> None:
        """Build the view.

        Args:
            part: the part whose tree this is; None for a detached container.
            container: the typed container --- a ``w:body``, ``w:hdr``,
                ``w:ftr``, ``w:tc``, ``w:sdtContent``, ``w:footnotes``, ...
            prefix: the address prefix this body reports under (``"body"``,
                ``"header:rId3"``); addresses themselves are Phase D.
            package: the package, for the styles part; taken from `part` when
                not given.
        """
        #: The part whose tree this is, or None.
        self.part = part
        #: The typed container object.
        self.container = container
        #: The address prefix (CR-003 section 3.4); Phase D reports it.
        self.prefix = prefix
        #: The package, for style names and for the id generator.
        self.package = package if package is not None else getattr(part, "package", None)

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same container are equal."""
        return isinstance(other, Body) and other.container is self.container

    def __hash__(self) -> int:
        """Hashes by the container's identity."""
        return hash(id(self.container))

    def __repr__(self) -> str:
        """``<Body body, 12 blocks>``."""
        return f"<Body {self.prefix}, {len(self)} blocks>"

    # -- the tree ----------------------------------------------------------

    @property
    def content(self) -> list:
        """The live block-level list (docx4j ``getContent()``); ``w:sectPr`` is not in it."""
        found = block_list_of(self.container)
        if found is None:
            raise ContentError(
                f"{qualified(element_name(self.container))} holds no block-level content",
                code="body.no_content",
                hint="a Body is over a w:body, w:hdr, w:ftr, w:tc, w:sdtContent or a notes part",
            )
        owner, field = found
        items = getattr(owner, field)
        if not isinstance(items, ChildList):
            items = ChildList(items, owner=owner)
            setattr(owner, field, items)
        return items

    def _owner(self) -> Any:
        """The object whose list :attr:`content` is, for parent links."""
        found = block_list_of(self.container)
        return found[0] if found is not None else self.container

    def sub(self, container: Any, prefix: str | None = None) -> Body:
        """A body over a container nested in this one: a cell, a control.

        The part and the package are this body's, and the prefix carries on
        from this one, so that the addresses a nested body reports in Phase D
        continue this body's (``"body/4/0/1"`` for a cell).
        """
        return Body(self.part, container, prefix or self.prefix, self.package)

    # -- Sequence ----------------------------------------------------------

    def __len__(self) -> int:
        """The number of block-level children."""
        return len(self.content)

    def __getitem__(self, index: Any) -> Any:
        """The view of the block at `index`; a slice gives a list of views."""
        items = self.content
        if isinstance(index, slice):
            return [self.view_for(item) for item in items[index]]
        return self.view_for(items[index])

    def __iter__(self) -> Iterator[Any]:
        """The block views, in document order."""
        return iter(self.iter_blocks())

    def __contains__(self, item: object) -> bool:
        """``"Chapter 2" in body`` is a text search; anything else is identity."""
        if isinstance(item, str):
            return item in self.text
        if isinstance(item, Paragraph):
            return any(item.element is block for block in self.content)
        return any(item is block or item == block for block in self.content)

    def view_for(self, element: Any) -> Any:
        """The view of a block: a :class:`Paragraph`, else a :class:`Block`."""
        container = self._container_of(element)
        if isinstance(element, P):
            return Paragraph(element, container, self)
        return Block(element, container)

    def paragraph_for(self, element: P) -> Paragraph:
        """The view of a ``w:p`` anywhere in this body's tree."""
        return Paragraph(element, self._container_of(element), self)

    def _container_of(self, element: Any) -> list:
        """The live list holding an element, from its parent pointer."""
        parent = getattr(element, "parent", None)
        if parent is not None:
            items = block_children_of(parent)
            if items is not None and any(item is element for item in items):
                return items
        return self.content

    # -- reading -----------------------------------------------------------

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Every paragraph in document order, tables and controls descended into.

        A content control of any of the four forms is transparent, and so is a
        table: its rows, their cells and the cells' blocks are all the same
        ``content`` list in this model, so the descent is one function
        (CR-003 section 4: children of every ``w:sdt`` form are visible here).
        """
        return list(self.iter_paragraphs())

    def iter_paragraphs(self) -> Iterator[Paragraph]:
        """:attr:`paragraphs`, lazily, for a document too big to hold at once."""

        def visit(items: list) -> Iterator[Paragraph]:
            for item in items:
                if isinstance(item, P):
                    yield Paragraph(item, items, self)
                    continue
                children = block_children_of(item)
                if children is not None:
                    yield from visit(children)

        yield from visit(self.content)

    def iter_blocks(self) -> Iterator[Any]:
        """This body's own block-level children, as views, in order."""
        items = self.content
        for item in list(items):
            yield Paragraph(item, items, self) if isinstance(item, P) else Block(item, items)

    @property
    def text(self) -> str:
        """The text, a paragraph per line: Word's main story, the accepted view.

        No text boxes (CR-003 decided question 8): a text box lives in a
        ``w:drawing``, which the text model does not descend into.
        """
        return "\n".join(paragraph.text for paragraph in self.iter_paragraphs())

    def get_text(self, *, view: TextView = "accepted", max_chars: int | None = None) -> str:
        """The text, with a budget and a choice of view (CR-003 section 3.4).

        Args:
            view: ``"accepted"`` (the default) or ``"original"``, the document
                as it read before its tracked changes.
            max_chars: stop after this many characters, so that a 200-page
                document can be read into a context window that cannot hold it.
        """
        out: list[str] = []
        total = 0
        for paragraph in self.iter_paragraphs():
            line = paragraph.get_text(view=view)
            out.append(line)
            total += len(line) + 1
            if max_chars is not None and total > max_chars:
                return "\n".join(out)[:max_chars]
        text = "\n".join(out)
        return text if max_chars is None else text[:max_chars]

    def get_xml(self) -> str:
        """The part's XML, or the container's when this body has no part."""
        if self.part is not None:
            return self.part.get_xml()
        return to_xml(self.container)

    # -- inserting ---------------------------------------------------------

    def insert_paragraph(
        self,
        text: str = "",
        *,
        location: BodyLocation = "End",
        style: str | None = None,
    ) -> Paragraph:
        """Append or prepend a paragraph of the text (Office JS ``insertParagraph``).

        Args:
            text: the paragraph's text; tabs and newlines become ``w:tab`` and
                ``w:br``, as the ``r`` builder writes them.
            location: ``"End"`` (the default) or ``"Start"``.
            style: a style display name, stored name or id.

        Returns:
            The new paragraph's view.
        """
        element = P(content=ChildList([r(text)] if text else []))
        view = self.insert_element(element, location=location)
        if style is not None:
            view.style = style
        return view  # type: ignore[return-value]

    def insert_text(self, text: str, *, location: TextLocation = "End") -> Range:
        """Text at the start of the first paragraph, the end of the last, or instead of all.

        Returns:
            The :class:`~docx4j_py.model.content.Range` of the inserted text.
        """
        paragraphs = self.paragraphs
        if location == "Replace" or not paragraphs:
            if location == "Replace":
                self.clear()
            return self.insert_paragraph(text).get_range("Content")
        if location == "Start":
            return paragraphs[0].insert_text(text, location="Start")
        return paragraphs[-1].insert_text(text, location="End")

    def insert_break(
        self,
        type: BreakTypeValue = "Page",
        *,
        location: BodyLocation = "End",
    ) -> None:
        """A new paragraph holding a page or line break, at the start or the end."""
        paragraph = self.insert_paragraph("", location=location)
        paragraph.element.content.append(
            R(content=ChildList([br("page" if type == "Page" else None)]))
        )
        link_parents(paragraph.element)

    def insert_xml(
        self,
        xml: str,
        *,
        location: Location = "End",
        target: Any = None,
    ) -> list[Any]:
        """Insert a WordprocessingML fragment and return the views of what went in.

        The fragment is what is written *inside* ``document.xml`` --- one or
        more ``w:p``, ``w:tbl``, ``w:sdt`` --- and the prefixes are declared for
        it (:func:`docx4j_py.wml.wml`, resolved in a ``w:body``'s scope, so a
        ``w:sdt`` comes back as the block form). Views rather than a range,
        because a fragment may bring several blocks (CR-003 section 4).

        Args:
            xml: the fragment.
            location: ``"Start"``, ``"End"`` (the default), ``"Before"`` or
                ``"After"``; the last two need `target`.
            target: the paragraph or block to insert relative to.
        """
        elements = wml.all(xml, wrapper="body")
        if not elements:
            return []
        self.insert_element(elements, location=location, target=target)
        return [self.view_for(element) for element in elements]

    def insert_element(
        self,
        element: Any,
        *,
        location: Location = "End",
        target: Any = None,
    ) -> Any:
        """Insert block-level elements (docx4j ``addObject``), parents linked.

        The elements are checked against what the container really holds ---
        the choices of its compound field, read from the model's own metadata
        --- so the error names what was passed and what the container takes,
        rather than guessing from a list written here.

        Args:
            element: one element or a list of them.
            location: ``"Start"``, ``"End"`` (the default), ``"Before"`` or
                ``"After"``.
            target: for ``"Before"`` and ``"After"``, the paragraph or block to
                go next to.

        Returns:
            The view of the first inserted element.

        Raises:
            InvalidTargetError: the container cannot hold one of the elements.
            ContentError: ``"Before"`` or ``"After"`` with no target.
        """
        elements = list(element) if isinstance(element, (list, tuple)) else [element]
        if not elements:
            raise ContentError(
                "nothing to insert",
                code="insert.empty",
                hint="pass an element, or a list with at least one in it",
            )

        if location in ("Before", "After"):
            if target is None:
                raise ContentError(
                    f"{location!r} needs a target paragraph or block",
                    code="target.missing",
                    hint="pass target=paragraph, or use location='Start' / 'End'",
                )
            container = getattr(target, "container", None)
            element_of = getattr(target, "element", target)
            if container is None:
                container = self._container_of(element_of)
            index = next(
                (i for i, item in enumerate(container) if item is element_of),
                len(container),
            )
            if location == "After":
                index += 1
            owner = getattr(element_of, "parent", None) or self._owner()
        elif location in ("Start", "End"):
            container = self.content
            index = 0 if location == "Start" else len(container)
            owner = self._owner()
        else:
            raise ContentError(
                f"insert_element takes Start, End, Before or After, not {location!r}",
                code="location.invalid",
                hint="Replace is for insert_text and for a Range",
            )

        self._check(elements, owner)
        for offset, item in enumerate(elements):
            container.insert(index + offset, item)
            link_parents(item)
            item.parent = owner
            if isinstance(item, P):
                self._assign_para_id(item)
        return self.view_for(elements[0])

    def _check(self, elements: list, owner: Any) -> None:
        """Refuse an element the owner's block list cannot hold."""
        allowed = _accepted_names(owner)
        if allowed is None:
            return
        for item in elements:
            name = element_name(item)
            if name in allowed:
                continue
            if any(isinstance(item, cls) for classes in allowed.values() for cls in classes):
                continue
            takes = ", ".join(qualified(q) for q in list(allowed)[:8])
            more = f" and {len(allowed) - 8} more" if len(allowed) > 8 else ""
            raise InvalidTargetError(
                f"{qualified(element_name(owner))} cannot hold "
                f"{qualified(name)} ({type(item).__name__}); it takes {takes}{more}",
                hint=(
                    "wrap a run in a w:p, or insert it into a paragraph with "
                    "paragraph.insert_xml()"
                ),
            )

    # -- paragraph ids (CR-003 section 3.4, determinism) -------------------

    def _para_ids(self) -> set[str]:
        return {p.para_id for p in self.iter_paragraphs() if p.para_id}

    def _assign_para_id(self, element: P) -> None:
        """Give a new paragraph a ``w14:paraId`` when the document uses them."""
        if element.para_id:
            return
        taken = self._para_ids()
        if not taken:
            return
        package = self.package
        if package is None:
            return
        generator = package.id_generator(derive_from=taken)
        for _attempt in range(64):
            candidate = f"{generator.randrange(1, 0x7FFFFFFF):08X}"
            if candidate not in taken:
                element.para_id = candidate
                return

    # -- searching and editing --------------------------------------------

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text`, over every paragraph, in document order.

        A match is found on each paragraph's whole text, so it **spans runs**
        freely; only formatting one splits them (CR-003 section 8).

        Args:
            text: what to look for.
            **options: ``match_case`` (default False, which is
                ``re.IGNORECASE``), ``match_whole_word`` (Unicode word
                boundaries), ``match_wildcards`` (Word's ``?`` and ``*``, and
                ``[]``, ``<``, ``>``, ``@``, ``{}``), and ``limit``.
        """
        limit = options.pop("limit", None)
        out: list[Range] = []
        for paragraph in self.iter_paragraphs():
            for hit in paragraph.search(text, **options):
                out.append(hit)
                if limit is not None and len(out) >= limit:
                    return out
        return out

    def replace_text(self, find: str, replace: str, **options: Any) -> int:
        """Replace every match, last first so the offsets stay valid.

        Returns:
            How many were replaced.
        """
        count = 0
        for paragraph in self.paragraphs:
            count += paragraph.replace_text(find, replace, **options)
        return count

    def clear(self) -> None:
        """Remove every block. The section properties stay where they are."""
        self.content.clear()

    def get_range(self, location: RangeLocation = "Whole") -> Range:
        """A range over this body's first or last paragraph (Office JS's shape)."""
        paragraphs = self.paragraphs
        if not paragraphs:
            return self.insert_paragraph("").get_range("Content")
        if location == "Start":
            return paragraphs[0].get_range("Start")
        if location == "End":
            return paragraphs[-1].get_range("End")
        return paragraphs[0].get_range("Whole" if location == "Whole" else location)

    # -- docx4j names ------------------------------------------------------

    def add_paragraph_of_text(self, text: str) -> P:
        """docx4j ``MainDocumentPart.addParagraphOfText``; returns the ``w:p``."""
        return self.insert_paragraph(text).element

    def add_styled_paragraph_of_text(self, style_id: str, text: str) -> P:
        """docx4j ``addStyledParagraphOfText``; `style_id` is a style id."""
        paragraph = self.insert_paragraph(text)
        paragraph.style_id = style_id
        return paragraph.element

    def add_object(self, element: Any) -> None:
        """docx4j ``addObject``: append a block-level element."""
        self.insert_element(element)

    def get_content(self) -> list:
        """docx4j ``getContent()``: :attr:`content`."""
        return self.content


# ---------------------------------------------------------------------------
# what a container may hold, from the model's own metadata
# ---------------------------------------------------------------------------

_accepted_cache: dict[tuple[type, str], dict[str, tuple[type, ...]]] = {}


def _accepted_names(owner: Any) -> dict[str, tuple[type, ...]] | None:
    """The element names the owner's block list holds, and the classes for each.

    Read from the ``XmlContext`` metadata: a compound field carries its
    alternatives in ``elements``, and a plain repeated element field carries
    its own ``qname``. Nothing is listed here, so a container this code has
    never heard of validates correctly all the same.
    """
    found = block_list_of(owner)
    if found is None:
        return None
    holder, field = found
    key = (type(holder), field)
    try:
        return _accepted_cache[key]
    except KeyError:
        pass
    try:
        meta = context().build(type(holder))
    except Exception:  # noqa: BLE001 - not a model class; validate nothing
        _accepted_cache[key] = {}
        return None
    allowed: dict[str, tuple[type, ...]] = {}
    for var in meta.get_element_vars():
        if var.name != field:
            continue
        if var.elements:
            for qname, alternative in var.elements.items():
                allowed[qname] = tuple(alternative.types)
        elif var.qname:
            allowed[var.qname] = tuple(var.types)
    _accepted_cache[key] = allowed
    return allowed or None


# ---------------------------------------------------------------------------
# the registration of CR-003 section 5
# ---------------------------------------------------------------------------

#: Relationship-type suffix -> the address prefix a part's body reports under.
_PREFIXES: dict[str, str] = {
    "header": "header",
    "footer": "footer",
    "footnotes": "footnote",
    "endnotes": "endnote",
    "comments": "comment",
}


def _prefix_for(part: Any) -> str:
    name = type(part).__name__
    for suffix, prefix in _PREFIXES.items():
        if name.lower().startswith(suffix):
            rel_id = _rel_id_of(part)
            return f"{prefix}:{rel_id}" if rel_id else prefix
    return "body"


def _rel_id_of(part: Any) -> str | None:
    """The relationship id a header or footer is reached by, when there is one."""
    source = getattr(part, "source_relationships", None) or ()
    for relationship in source:
        rel_id = getattr(relationship, "id", None)
        if rel_id:
            return str(rel_id)
    return None


def body_of(target: Any) -> Body:
    """The :class:`Body` of a part or a package. What ``part.body`` is.

    Args:
        target: a ``MainDocumentPart``, ``HeaderPart``, ``FooterPart``,
            ``FootnotesPart``, ``EndnotesPart``, ``CommentsPart``, or a
            ``WordprocessingMLPackage`` (whose main document part answers).

    Raises:
        ContentError: the target has no block-level content.
    """
    main = getattr(target, "main_document_part", None)
    if main is not None or hasattr(target, "get_main_document_part"):
        package = target
        part = main if main is not None else target.get_main_document_part()
        if part is None:
            raise ContentError(
                "this package has no main document part",
                code="body.no_part",
                hint="load a .docx, or call create_package()",
            )
        return body_of(part)

    contents = target.contents
    container = getattr(contents, "body", None)
    if container is None:
        container = contents
    if block_list_of(container) is None:
        raise ContentError(
            f"{type(target).__name__} has no block-level content, so it has no body",
            code="body.no_content",
            hint="body is on the document, header, footer, footnotes, endnotes and comments parts",
        )
    package = getattr(target, "package", None)
    return Body(target, container, _prefix_for(target), package)


def register_body() -> None:
    """Install ``body`` on the parts and the package. See this package's ``__init__``."""
    from docx4j_py.model.content import register

    register()
