"""Which tracked changes to resolve. docx4j ``TrackedChangeFilter``.

CR-010 section 5 (Phase B), docx4j CR-043 section 3.1: by author, type, kind,
date, part, and range --- a content control's content, or a bookmark's extent.
Every condition set must hold; a condition not set holds for all. The
conditions apply to each **single** change: of a group (touching changes of
one author and type) only the pieces that match are resolved (U19). A range is
read once, from the package as it stands, before anything is resolved.

::

    from docx4j_py.model import revisions
    from docx4j_py.model.revisions import TrackedChangeFilter

    revisions.reject_all(pkg, TrackedChangeFilter(authors=["Author B"]))
    revisions.accept_all(pkg, TrackedChangeFilter().types("Added").within_bookmark("Terms"))
    revisions.tracked_changes(pkg, TrackedChangeFilter(kinds=[revisions.Kind.RUN_INSERTION]))

The conditions can be given as keywords or set one by one, each setter
returning the filter, as docx4j's builder does. Anywhere a filter is taken, a
plain callable over a single change still is (CR-010 section 10.3 item 5).
"""

from __future__ import annotations

import datetime
import re
from collections.abc import Callable, Iterable
from typing import Any

from docx4j_py.model.revisions import trees

__all__ = ["Filter", "TrackedChangeFilter", "matcher"]


class TrackedChangeFilter:
    """Which changes ``accept_all``, ``reject_all`` and ``tracked_changes`` take."""

    __slots__ = ("_authors", "_bookmark", "_control", "_from", "_kinds", "_parts", "_to", "_types")

    def __init__(
        self,
        *,
        authors: Iterable[str] | None = None,
        types: Iterable[str] | None = None,
        kinds: Iterable[str] | None = None,
        date_from: datetime.datetime | None = None,
        date_to: datetime.datetime | None = None,
        parts: Iterable[Any] | None = None,
        within: Any = None,
        within_bookmark: str | None = None,
    ) -> None:
        """Build the filter; every keyword is a condition, and one not given holds for all.

        Args:
            authors: changes by these authors (``w:author``).
            types: changes of these types: Office JS's ``"Added"``,
                ``"Deleted"``, ``"Formatted"``.
            kinds: changes of these kinds (a single change's:
                :class:`~docx4j_py.model.revisions.Kind`, or its value
                ``"run_insertion"``, ``"paragraph_mark_deletion"``, ...).
            date_from: changes made at or after this.
            date_to: changes made before this. A change with no date matches
                no date range; a naive datetime is read as UTC. A change's date
                is read as Word means it (:attr:`TrackedChange.date`:
                ``w16du:dateUtc``, else ``w:date`` on the local wall clock).
            parts: changes in these parts (a part, or a part name such as
                ``"/word/footnotes.xml"``).
            within: changes inside this content control's content (the
                control as it stands in the package: ``SdtBlock``, ``SdtRun``,
                a row's or a cell's, or the content API's ``ContentControl``).
            within_bookmark: changes between the start and the end of the
                bookmark of this name, in whichever story it is in.
        """
        self._authors: frozenset[str] | None = None
        self._types: frozenset[str] | None = None
        self._kinds: frozenset[str] | None = None
        self._from: datetime.datetime | None = None
        self._to: datetime.datetime | None = None
        self._parts: tuple[Any, ...] | None = None
        self._control: Any = None
        self._bookmark: str | None = None
        if authors is not None:
            self.authors(*authors)
        if types is not None:
            self.types(*types)
        if kinds is not None:
            self.kinds(*kinds)
        if date_from is not None or date_to is not None:
            self.between(date_from, date_to)
        if parts is not None:
            self.parts(*parts)
        if within is not None:
            self.within(within)
        if within_bookmark is not None:
            self.within_bookmark(within_bookmark)

    # -- the conditions, docx4j's builder ------------------------------------

    def authors(self, *authors: str) -> TrackedChangeFilter:
        """Changes by these authors (``w:author``)."""
        self._authors = frozenset(authors)
        return self

    def types(self, *types: str) -> TrackedChangeFilter:
        """Changes of these types: ``"Added"``, ``"Deleted"``, ``"Formatted"``."""
        self._types = frozenset(str(t) for t in types)
        return self

    def kinds(self, *kinds: str) -> TrackedChangeFilter:
        """Changes of these kinds: a single change's ``kind`` (``"run_insertion"``, ...)."""
        self._kinds = frozenset(str(k) for k in kinds)
        return self

    def between(
        self, date_from: datetime.datetime | None, date_to: datetime.datetime | None
    ) -> TrackedChangeFilter:
        """Changes made at or after `date_from` and before `date_to`; either may be None for no bound.

        A change with no date matches no date range. A naive datetime is read
        as UTC, as a change's date is given.
        """
        self._from = _utc(date_from)
        self._to = _utc(date_to)
        return self

    def parts(self, *parts: Any) -> TrackedChangeFilter:
        """Changes in these parts: a part, or its name (``"/word/document.xml"``)."""
        self._parts = tuple(parts)
        return self

    def within(self, control: Any) -> TrackedChangeFilter:
        """Changes inside this content control's content, as the control stands in the package."""
        self._control = control
        self._bookmark = None
        return self

    def within_bookmark(self, name: str) -> TrackedChangeFilter:
        """Changes between the start and the end of the bookmark of this name, in whichever story it is in."""
        self._bookmark = name
        self._control = None
        return self

    # -- matching -------------------------------------------------------------

    def range(self, package: Any) -> dict[int, Any] | None:
        """The range, read once from the package before anything is resolved: the markup objects inside it.

        None when the filter has no range. Keyed by identity; the values hold
        the objects, so that no id is reused while the range is in use.

        Raises:
            ValueError: no bookmark of the name is in any story.
        """
        if self._control is None and self._bookmark is None:
            return None
        inside: dict[int, Any] = {}
        if self._control is not None:
            control = getattr(self._control, "element", self._control)
            content = getattr(control, "sdt_content", None)
            if content is None:
                return inside

            def enter(value: Any) -> bool:
                inside[id(value)] = value
                return True

            trees.walk(content, enter)
            return inside
        from docx4j_py.model.revisions.lister import package_stories, story_parts

        for story in package_stories(package):
            walk = _BookmarkWalk(self._bookmark, inside)
            walk.list(story.root)
            if walk.found:
                return inside
        # a story part with no revision markup is not read for the listing, and
        # a bookmark there has nothing to resolve inside it
        pattern = re.compile(rb"bookmarkStart[^>]*\bw:name=\"" + re.escape(self._bookmark.encode("utf-8")) + rb"\"")
        for part in story_parts(package):
            data = None if getattr(part, "is_unmarshalled", True) else _raw_bytes(part)
            if data is not None and pattern.search(data):
                return inside
        raise ValueError(f"no bookmark named {self._bookmark!r}")

    def matches(self, piece: Any, range_: dict[int, Any] | None) -> bool:
        """True when the single change `piece` meets every condition set."""
        if self._authors is not None and piece.author not in self._authors:
            return False
        if self._types is not None and piece.type not in self._types:
            return False
        if self._kinds is not None and str(piece.kind) not in self._kinds:
            return False
        if self._from is not None or self._to is not None:
            date = piece.date
            if date is None:
                return False
            if self._from is not None and date < self._from:
                return False
            if self._to is not None and not date < self._to:
                return False
        if self._parts is not None and not _in_parts(piece.part, self._parts):
            return False
        return range_ is None or _in_range(piece, range_)

    def bind(self, package: Any) -> Callable[[Any], bool]:
        """A predicate over single changes with this filter's range read now, from `package`."""
        range_ = self.range(package)
        return lambda piece: self.matches(piece, range_)

    def __repr__(self) -> str:
        """``TrackedChangeFilter(authors=['A'], kinds=['run_insertion'])``, the conditions set."""
        shown = []
        for label, value in (
            ("authors", self._authors),
            ("types", self._types),
            ("kinds", self._kinds),
        ):
            if value is not None:
                shown.append(f"{label}={sorted(value)!r}")
        if self._from is not None:
            shown.append(f"date_from={self._from.isoformat()!r}")
        if self._to is not None:
            shown.append(f"date_to={self._to.isoformat()!r}")
        if self._parts is not None:
            shown.append(f"parts={[_part_name(p) for p in self._parts]!r}")
        if self._control is not None:
            shown.append("within=<control>")
        if self._bookmark is not None:
            shown.append(f"within_bookmark={self._bookmark!r}")
        return f"TrackedChangeFilter({', '.join(shown)})"


#: What the module functions take as a filter: a :class:`TrackedChangeFilter`,
#: or any callable answering for one single change whether it is taken.
Filter = TrackedChangeFilter | Callable[[Any], bool]


def matcher(filter: Filter | None, package: Any) -> Callable[[Any], bool] | None:
    """A filter as a predicate over single changes, its range (if any) read once, now; None for none."""
    if filter is None:
        return None
    if isinstance(filter, TrackedChangeFilter):
        return filter.bind(package)
    return filter


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _raw_bytes(part: Any) -> bytes | None:
    """An unread part's bytes as they were loaded; None when there are none to read."""
    try:
        return part._raw_bytes()
    except Exception:  # noqa: BLE001 - no bytes to read
        return None


def _utc(value: datetime.datetime | None) -> datetime.datetime | None:
    if value is None:
        return None
    if value.tzinfo is None:
        return value.replace(tzinfo=datetime.UTC)
    return value


def _part_name(part: Any) -> str:
    if isinstance(part, str):
        return part
    name = getattr(part, "part_name", None)
    return str(name) if name is not None else ""


def _in_parts(part: Any, parts: tuple[Any, ...]) -> bool:
    for wanted in parts:
        if isinstance(wanted, str):
            if part is not None and _part_name(part) == wanted:
                return True
        elif wanted is part:
            return True
    return False


def _in_range(piece: Any, range_: dict[int, Any]) -> bool:
    """True when where `piece` stands was inside the range.

    A change in a text box kept as a wildcard tree stands in a typed island
    made after the range was read: it is inside when its text box was.
    """
    if id(piece.anchor()) in range_:
        return True
    island = piece.story.island
    return island is not None and id(island.node) in range_


class _BookmarkWalk:
    """A story walked in document order, a paragraph's mark after its content, noting what stands inside a bookmark.

    The bookmark's end is matched by its start's id.
    """

    __slots__ = ("found", "inside", "mark_id", "name", "open")

    def __init__(self, name: str, inside: dict[int, Any]) -> None:
        self.name = name
        self.inside = inside
        self.mark_id: Any = None
        self.open = False
        self.found = False

    def list(self, items: list) -> None:
        for item in items:
            self.node(item)

    def node(self, item: Any) -> None:
        p_class = trees.cls("P")
        if isinstance(item, p_class):
            self.paragraph(item)
            return

        def enter(value: Any) -> bool:
            local = trees.name(value)
            if (
                local == "bookmarkStart"
                and not self.found
                and trees.attribute(value, "name") == self.name
            ):
                self.mark_id = trees.marker_id(value)
                self.open = True
                self.found = True
                return False
            if local == "bookmarkEnd" and self.open and trees.marker_id(value) == self.mark_id:
                self.open = False
                return False
            if isinstance(value, p_class):
                self.paragraph(value)
                return False
            if self.open:
                self.inside[id(value)] = value
            return True

        trees.walk(item, enter)

    def paragraph(self, p: Any) -> None:
        """A paragraph: its properties' record first, then its content, then its mark."""
        p_pr = p.p_pr
        if self.open and p_pr is not None and p_pr.p_pr_change is not None:
            self.inside[id(p_pr.p_pr_change)] = p_pr.p_pr_change
        self.list(p.content)
        self.mark(p)

    def mark(self, p: Any) -> None:
        if not self.open:
            return
        self.inside[id(p)] = p
        p_pr = p.p_pr
        if p_pr is None:
            return
        r_pr = p_pr.r_pr
        if r_pr is not None:
            for value in (r_pr, r_pr.ins, r_pr.del_value, r_pr.r_pr_change):
                if value is not None:
                    self.inside[id(value)] = value
        if p_pr.sect_pr is not None:
            self.inside[id(p_pr.sect_pr)] = p_pr.sect_pr
