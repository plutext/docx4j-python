"""A document's tracked changes counted. docx4j ``TrackedChangesReport``.

CR-010 section 5 (Phase B), docx4j CR-043 section 3.1: the listing counted by
kind, by author and by part, with :meth:`TrackedChangesReport.to_dict` the form
an MCP tool reports (docx4j's ``toJson``). Every kind listed is resolved by
``accept_all`` and ``reject_all``, so there is no "unresolved" count: docx4j's
phase 1 had one, and its phase 2 removed it.

::

    revisions.report(pkg).to_dict()
    # {'total': 3, 'by_kind': {'run_insertion': 2, 'table_properties': 1},
    #  'by_author': {'Jason': 3}, 'by_part': {'/word/document.xml': 3}}
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.revisions.change import Kind, TrackedChange

__all__ = ["TrackedChangesReport"]


class TrackedChangesReport:
    """Counts by kind, author and part of a listing."""

    __slots__ = ("by_author", "by_kind", "by_part", "total")

    def __init__(self, changes: list[TrackedChange]) -> None:
        """Count `changes` (``tracked_changes(pkg)``'s listing)."""
        counted: dict[Kind, int] = {}
        by_author: dict[str, int] = {}
        by_part: dict[str, int] = {}
        for change in changes:
            counted[change.kind] = counted.get(change.kind, 0) + 1
            author = change.author or ""
            by_author[author] = by_author.get(author, 0) + 1
            part_name = getattr(change.part, "part_name", None)
            part = str(part_name) if part_name is not None else ""
            by_part[part] = by_part.get(part, 0) + 1
        #: How many changes are listed.
        self.total: int = len(changes)
        #: By kind, in :class:`~docx4j_py.model.revisions.Kind`'s order.
        self.by_kind: dict[Kind, int] = {kind: counted[kind] for kind in Kind if kind in counted}
        #: By author (``w:author``; ``""`` for none), in the order first listed.
        self.by_author: dict[str, int] = by_author
        #: By part name (``/word/document.xml``, ``/word/footnotes.xml``, ...), in the order first listed.
        self.by_part: dict[str, int] = by_part

    def to_dict(self) -> dict[str, Any]:
        """``{"total": 3, "by_kind": {"run_insertion": 2, ...}, "by_author": {...}, "by_part": {...}}``.

        Plain strings and ints, ready for ``json.dumps``.
        """
        return {
            "total": self.total,
            "by_kind": {str(kind): count for kind, count in self.by_kind.items()},
            "by_author": dict(self.by_author),
            "by_part": dict(self.by_part),
        }

    def __repr__(self) -> str:
        """The dictionary form."""
        return f"TrackedChangesReport({self.to_dict()!r})"
