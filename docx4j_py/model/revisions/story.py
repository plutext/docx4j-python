"""One story's content. docx4j ``org.docx4j.model.revisions.Story``.

CR-010 section 3.1: a part's body, a header, a footer, a note, a comment, or a
text box. A paragraph join never crosses its boundary, and a move's halves pair
within it. A text box that this package keeps as a wildcard tree (CR-002's
lossless MCE) is a story over a **typed island** (CR-010 section 3.5,
:mod:`.islands`), which the story carries so that resolving writes it back.
"""

from __future__ import annotations

from typing import Any

__all__ = ["Story"]


class Story:
    """A story: the part it is in, its live block list, and what holds that list."""

    __slots__ = ("island", "package", "part", "root", "root_owner", "twin")

    def __init__(
        self,
        package: Any,
        part: Any,
        root: list,
        root_owner: Any,
        twin: bool = False,
        island: Any = None,
    ) -> None:
        """Build the story.

        Args:
            package: the package, or None for a detached container.
            part: the part the story is in (for a text box, the part the box
                is in), or None.
            root: the live list of the story's block-level content.
            root_owner: what holds `root`: a ``w:body``, a ``w:hdr``, a note,
                a ``w:txbxContent``.
            twin: True for a text box in a second branch of
                ``mc:AlternateContent``: its changes are twins of the first
                branch's, not listed on their own.
            island: the :class:`~docx4j_py.model.revisions.islands.Island` a
                wildcard text box was parsed into, or None.
        """
        self.package = package
        self.part = part
        self.root = root
        self.root_owner = root_owner
        self.twin = twin
        self.island = island

    def __repr__(self) -> str:
        """``<Story /word/document.xml Body>``."""
        where = getattr(self.part, "part_name", None)
        kind = type(self.root_owner).__name__
        return f"<Story {where} {kind}{' twin' if self.twin else ''}{' island' if self.island else ''}>"
