"""What one reading of a package walks, read before the walk begins. CR-011 section 3.2.

core-ts CR-012's ``PackageTrees`` and ``TextContext``: every story's tree,
read **privately** where its part was not unmarshalled
(``XmlPart.read_contents``), so that reading a document costs no part its
byte-for-byte round trip (REQ-204); the property resolver and the numbering
emulator, which read their parts without unmarshalling them; and what a story
needs from the rest of the document: the note references' order (``[n]``),
the numbering states per story, the islands of wildcard text boxes, and each
view's plan, made once.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.revisions import trees as rtrees

__all__ = ["Context", "PackageTrees"]


def _real(part: Any) -> Any:
    """The real part a trial's copy stands for; the part itself otherwise."""
    return getattr(part, "_wrapped", part)


def _read(part: Any) -> Any:
    """``read_contents()``, None for a part that is absent or cannot be read."""
    if part is None:
        return None
    try:
        return part.read_contents()
    except Exception:  # noqa: BLE001 - a part that will not read holds no text here
        return None


class PackageTrees:
    """The trees and resolution utilities of one package, as one reading sees them."""

    __slots__ = (
        "_anchored",
        "comments",
        "comments_extended",
        "document",
        "emulator",
        "endnotes",
        "footnotes",
        "headers_footers",
        "main",
        "package",
        "resolver",
        "settings",
    )

    def __init__(self, package: Any) -> None:
        """Read every story's tree, the settings, and the resolver and emulator where the package has their parts."""
        self.package = package
        main = getattr(package, "main_document_part", None) if package is not None else None
        #: The main document part, and its ``w:document``.
        self.main = main
        self.document = _read(main)
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
        if main is None or self.document is None:
            return
        if getattr(package, "style_definitions_part", None) is not None:
            try:
                self.resolver = package.get_property_resolver()
            except Exception:  # noqa: BLE001 - hidden text and indents then not resolved through styles
                self.resolver = None
        if getattr(main, "numbering_definitions_part", None) is not None:
            from docx4j_py.model.listnumbering.emulator import emulator_of

            try:
                self.emulator = emulator_of(package)
            except Exception:  # noqa: BLE001 - lists then not counted
                self.emulator = None
        self.settings = _read(getattr(main, "document_settings_part", None))
        for name in ("footnotes", "endnotes", "comments"):
            part = getattr(main, f"{name}_part", None)
            tree = _read(part)
            if tree is not None:
                setattr(self, name, (part, tree))
        self.comments_extended = _read(getattr(main, "comments_extended_part", None))
        try:
            parts = [*main.header_parts(), *main.footer_parts()]
        except Exception:  # noqa: BLE001 - no relationships: no headers
            parts = []
        for part in parts:
            tree = _read(part)
            if tree is not None:
                self.headers_footers[id(_real(part))] = (part, tree)

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
