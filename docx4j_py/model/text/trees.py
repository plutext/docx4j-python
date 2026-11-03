"""What one reading of a package walks, read before the walk begins. CR-011 section 3.2.

core-ts CR-012's ``PackageTrees`` and ``TextContext``: every story's tree,
read **privately** where its part was not unmarshalled
(``XmlPart.read_contents``), so that reading a document costs no part its
byte-for-byte round trip (REQ-204); the property resolver and the numbering
emulator, which read their parts without unmarshalling them; and what a story
needs from the rest of the document: the note references' order (``[n]``),
the numbering states per story, the islands of wildcard text boxes, and each
view's plan, made once.

The styles, numbering and settings parts are consulted, not emitted: one that
cannot be read (its XML not well formed, CR-031) is read as if the package
lacked it, and the reading says so with ``part.not_read`` where it first needs
the part (WD 0.66 REQ-801, CR-031 section 4.6; :meth:`PackageTrees.part_not_read`).
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.revisions import trees as rtrees
from docx4j_py.openpackaging.exceptions import Docx4JException

__all__ = ["Context", "PackageTrees", "not_readable"]


def _real(part: Any) -> Any:
    """The real part a trial's copy stands for; the part itself otherwise."""
    return getattr(part, "_wrapped", part)


def _read(part: Any, unreadable: dict[int, tuple[Any, Exception]] | None = None) -> Any:
    """``read_contents()``: None for a part that is absent, and for one that cannot be read.

    A part that cannot be read is told from one that is absent (CR-027
    section 4.6): `unreadable`, where given, keeps it and its exception, by
    the real part (a trial's copy stands for its real part).
    """
    if part is None:
        return None
    try:
        return part.read_contents()
    except Exception as error:  # noqa: BLE001 - a part that will not read is kept, and its story says so
        if unreadable is not None:
            unreadable[id(_real(part))] = (part, error)
        return None


class _NoTree:
    """An lxml parser target that builds nothing: the parse checks that the XML is well formed, no more."""

    def close(self) -> None:
        return None


def _well_formed(part: Any) -> bool:
    """Whether a part's content is well-formed XML, without reading it: a parse of its bytes that builds no tree.

    True for a part already read (unmarshalled), False for one whose read
    failed (``XmlPart.is_content_unreadable``) and for one with no bytes.
    """
    from lxml import etree

    real = _real(part)
    if getattr(real, "is_content_unreadable", False):
        return False
    if getattr(real, "is_unmarshalled", False):
        return True
    try:
        etree.fromstring(real.xml, etree.XMLParser(target=_NoTree(), resolve_entities=False, huge_tree=True))
    except Exception:  # noqa: BLE001 - not well formed, or no bytes: the typed read says which
        return False
    return True


def _root_reason(error: BaseException) -> str:
    """The deepest ``__cause__``'s message on one line, its whitespace collapsed; its class's name where it has none.

    For XML that is not well formed the deepest cause is lxml's
    ``XMLSyntaxError``, and its ``msg`` is read, without the ``(<string>,
    line 2)`` its ``str()`` appends: the same text as the engine's "Problem with
    part ..." message (CR-031 section 4.7).
    """
    from lxml import etree

    root, seen = error, {id(error)}
    while root.__cause__ is not None and id(root.__cause__) not in seen:
        root = root.__cause__
        seen.add(id(root))
    text = root.msg if isinstance(root, etree.XMLSyntaxError) and root.msg else str(root)
    message = " ".join(str(text).split())
    if message:
        return message
    kind = type(root)
    return kind.__qualname__ if kind.__module__ == "builtins" else f"{kind.__module__}.{kind.__qualname__}"


def not_readable(part: Any, error: BaseException) -> Docx4JException:
    """Why a story part's text cannot be read, naming it. docx4j ``JaxbXmlPart.getReadableContents("read as text")``.

    ``The content of /word/document.xml could not be read, so it cannot be
    read as text: <reason>``, the reason the deepest cause's message on one
    line. The caller raises it from `error`, the part's own exception.
    """
    return Docx4JException(
        f"The content of {part.part_name} could not be read, so it cannot be read as text: {_root_reason(error)}"
    )


class PackageTrees:
    """The trees and resolution utilities of one package, as one reading sees them."""

    __slots__ = (
        "_anchored",
        "_consulted",
        "comments",
        "comments_extended",
        "document",
        "emulator",
        "endnotes",
        "footnotes",
        "headers_footers",
        "main",
        "numbering_part",
        "package",
        "parts_not_read",
        "resolver",
        "settings",
        "settings_part",
        "styles_part",
        "unreadable",
    )

    def __init__(self, package: Any) -> None:
        """Read every story's tree, the settings, and the resolver and emulator where the package has their parts.

        A story part that cannot be read (the main part, the notes, the
        comments, a header or a footer) is kept in :attr:`unreadable`, its tree
        None as an absent part's is. A styles or settings part that cannot be
        read is kept in :attr:`parts_not_read`, and no resolver is built over
        such a styles part; the numbering part is read the first time a
        paragraph needs it (:meth:`part_not_read`).
        """
        self.package = package
        main = getattr(package, "main_document_part", None) if package is not None else None
        #: The story parts that could not be read, by ``id`` of the real part: (part, exception).
        self.unreadable: dict[int, tuple[Any, Exception]] = {}
        #: The parts the reading consults and does not emit (the styles, the numbering, the settings) that could not
        #: be read, by ``id`` of the real part: (part, exception). CR-031 section 4.6.
        self.parts_not_read: dict[int, tuple[Any, Exception]] = {}
        #: The parts :meth:`part_not_read` has read, by ``id`` of the real part: the tree, or None.
        self._consulted: dict[int, Any] = {}
        #: The styles, numbering and settings parts, where the package has them.
        self.styles_part: Any = None
        self.numbering_part: Any = None
        self.settings_part: Any = None
        #: The main document part, and its ``w:document``.
        self.main = main
        self.document = _read(main, self.unreadable)
        self.resolver: Any = None
        self.emulator: Any = None
        self.settings: Any = None
        #: (part, tree) of the footnotes, endnotes and comments parts, or None.
        self.footnotes: tuple[Any, Any] | None = None
        self.endnotes: tuple[Any, Any] | None = None
        self.comments: tuple[Any, Any] | None = None
        self.comments_extended: Any = None
        #: Each header and footer part the main part relates to, by ``id`` of the
        #: real part (a trial's copy stands for its real part): (part, tree).
        self.headers_footers: dict[int, tuple[Any, Any]] = {}
        self._anchored: set[str] | None = None
        # an unreadable main part's package is read still: the reading raises, or (nested) says so, and a
        # header or footer of it can be read alone (docx4j's Texts.of(Part))
        if main is None:
            return
        self.styles_part = getattr(package, "style_definitions_part", None)
        # a styles part that cannot be read: no resolver over it, the reading as for a package without styles
        if self.styles_part is not None and self.part_not_read(self.styles_part) is None:
            try:
                self.resolver = package.get_property_resolver()
            except Exception:  # noqa: BLE001 - hidden text and indents then not resolved through styles
                self.resolver = None
        self.numbering_part = getattr(main, "numbering_definitions_part", None)
        if self.numbering_part is not None:
            from docx4j_py.model.listnumbering.emulator import emulator_of

            try:
                self.emulator = emulator_of(package)
            except Exception:  # noqa: BLE001 - lists then not counted
                self.emulator = None
        self.settings_part = getattr(main, "document_settings_part", None)
        self.settings = self._consult(self.settings_part)
        for name in ("footnotes", "endnotes", "comments"):
            part = getattr(main, f"{name}_part", None)
            tree = _read(part, self.unreadable)
            if tree is not None:
                setattr(self, name, (part, tree))
        self.comments_extended = _read(getattr(main, "comments_extended_part", None))
        try:
            parts = [*main.header_parts(), *main.footer_parts()]
        except Exception:  # noqa: BLE001 - no relationships: no headers
            parts = []
        for part in parts:
            tree = _read(part, self.unreadable)
            if tree is not None:
                self.headers_footers[id(_real(part))] = (part, tree)

    def unreadable_error(self, part: Any) -> Exception | None:
        """Why `part` could not be read, or None: it was read, it is absent, or it is no story part."""
        found = None if part is None else self.unreadable.get(id(_real(part)))
        return None if found is None else found[1]

    def part_not_read(self, part: Any) -> Exception | None:
        """Why `part`, a part the reading consults and does not emit, could not be read; None where it was or is absent.

        The styles, the numbering and the settings (CR-031 section 4.6): read
        privately the first time it is asked (``read_contents()``, so a part
        that cannot be read is parsed once, the engine keeping its failure,
        ``XmlPart.unreadable``), and kept in :attr:`parts_not_read` where its
        content could not be read. The reading then goes on as for a package
        without the part, and says so with ``part.not_read``. A read refused
        for another reason (a strict load's skipped content) is not this: the
        caller reads as before, without the part, silently.
        """
        if part is None:
            return None
        self._consult(part)
        found = self.parts_not_read.get(id(_real(part)))
        return None if found is None else found[1]

    def _consult(self, part: Any) -> Any:
        """A consulted part's tree, read privately once; None where it is absent or will not read (:meth:`part_not_read`).

        The numbering part is the exception: nothing in the text model reads
        it typed (the emulator reads it with lxml), so where its XML is well
        formed (:func:`_well_formed`, a parse that builds no tree) it is not
        read here, and is consulted as readable; only where it is not is it
        read, for the engine's record of why.
        """
        if part is None:
            return None
        key = id(_real(part))
        if key in self._consulted:
            return self._consulted[key]
        if self.numbering_part is not None and _real(part) is _real(self.numbering_part) and _well_formed(part):
            self._consulted[key] = None
            return None
        tree = None
        try:
            tree = part.read_contents()
        except Exception as error:  # noqa: BLE001 - the part's own record says whether its content could not be read
            if getattr(_real(part), "unreadable", None) is error:
                self.parts_not_read[key] = (part, error)
        self._consulted[key] = tree
        return tree

    def known_not_read(self, part: Any) -> bool:
        """Whether `part` was consulted already and could not be read; reads nothing."""
        return part is not None and id(_real(part)) in self.parts_not_read

    @property
    def body(self) -> Any:
        """The ``w:body``, or None."""
        return None if self.document is None else self.document.body

    def header_or_footer(self, rel_id: str | None) -> tuple[Any, Any] | None:
        """The (part, tree) of the header or footer a relationship of the main part names, or None."""
        found = self.related(self.main, rel_id)
        return None if found is None else self.headers_footers.get(id(_real(found)))

    def relationship_id(self, part: Any) -> str:
        """The id of the main part's relationship to `part`; its name where there is none (docx4j ``Texts.rId``)."""
        rels = getattr(self.main, "relationships_part", None) if self.main is not None else None
        for rel in getattr(rels, "list", None) or ():
            if _real(self.related(self.main, rel.id)) is _real(part):
                return rel.id
        return str(part.part_name)

    def related(self, part: Any, rel_id: str | None) -> Any:
        """The part a relationship of `part` names, or None."""
        if part is None or not rel_id:
            return None
        rels = getattr(part, "relationships_part", None)
        if rels is None:
            return None
        try:
            return rels.get_part(rel_id)
        except Exception:  # noqa: BLE001 - a broken relationship names nothing
            return None

    def anchored_comments(self) -> set[str]:
        """The ids of the comments anchored in the document (REQ-504, WD 0.20).

        A ``w:commentReference`` or ``w:commentRangeStart`` in the main
        document, a header or a footer, or a note, in any branch.
        """
        if self._anchored is not None:
            return self._anchored
        ids: set[str] = set()
        roots = [self.document, *(tree for _part, tree in self.headers_footers.values())]
        for holder in (self.footnotes, self.endnotes):
            if holder is not None:
                roots.append(holder[1])

        def enter(node: Any) -> bool:
            if rtrees.name(node) in ("commentReference", "commentRangeStart"):
                value = rtrees.attribute(node, "id")
                if value is not None:
                    ids.add(str(value))
            return True

        for root in roots:
            if root is not None:
                rtrees.walk(root, enter)
        self._anchored = ids
        return ids

    def stories(self, only_with_revisions: bool = True) -> list[Any]:
        """Every story over these trees, for the views' plan: the main body, each note and comment, the headers and footers.

        With `only_with_revisions`, a part whose bytes hold no revision markup
        is left out (it has nothing to plan), as the engine's listing does.
        """
        from docx4j_py.model.revisions.lister import _may_hold_revisions
        from docx4j_py.model.revisions.story import Story

        out: list[Story] = []
        body = self.body
        if body is not None and (not only_with_revisions or _may_hold_revisions(self.main)):
            out.append(Story(self.package, self.main, body.content, body))
        for holder, field in ((self.footnotes, "footnote"), (self.endnotes, "endnote"), (self.comments, "comment")):
            if holder is None:
                continue
            part, tree = holder
            if only_with_revisions and not _may_hold_revisions(part):
                continue
            for note in getattr(tree, field, None) or ():
                out.append(Story(self.package, part, note.content, note))
        for part, tree in self.headers_footers.values():
            if only_with_revisions and not _may_hold_revisions(part):
                continue
            items = rtrees.block_items(tree)
            if items is not None:
                out.append(Story(self.package, part, items, tree))
        return out


class Context:
    """What a story needs from the rest of the document, for one reading in one view.

    The note references' order, the numbering states per story, the islands
    of wildcard text boxes (shared with the plan's listing), and the plans of
    Accept All and Reject All, each made once.
    """

    __slots__ = ("_plans", "endnotes", "footnotes", "islands", "numbering_states", "trees")

    def __init__(self, package_trees: PackageTrees) -> None:
        """A fresh context over the trees."""
        from docx4j_py.model.listnumbering.definitions import NumberingStates

        self.trees = package_trees
        #: Notes referenced, in reference order: id to number, 0 for a note with a custom mark.
        self.footnotes: dict[str, int] = {}
        self.endnotes: dict[str, int] = {}
        self.numbering_states = NumberingStates()
        #: The typed islands of wildcard text boxes, by ``id`` of the wildcard node.
        self.islands: dict[int, Any] = {}
        self._plans: dict[bool, Any] = {}

    def note_number(self, note_type: str, note_id: str, custom_mark: bool) -> int:
        """A note's number, given at its first reference: in reference order, footnotes and endnotes apart (REQ-503)."""
        numbers = self.footnotes if note_type == "footnote" else self.endnotes
        found = numbers.get(note_id)
        if found is None:
            found = 0 if custom_mark else 1 + sum(1 for n in numbers.values() if n > 0)
            numbers[note_id] = found
        return found

    def plan(self, accepted: bool) -> Any:
        """The plan of Accept All (`accepted`) or Reject All over every story with revisions, made once; None for none."""
        if accepted in self._plans:
            return self._plans[accepted]
        from docx4j_py.model.revisions.lister import Lister
        from docx4j_py.model.text.view import ResolvedView

        stories = self.trees.stories()
        plan = None
        if stories:
            changes = Lister(self.trees.package, False, stories, self.islands).list()
            if changes:
                plan = ResolvedView(changes, accepted)
        self._plans[accepted] = plan
        return plan
