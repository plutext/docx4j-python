"""What content removed by resolving takes with it. docx4j ``Removals``.

CR-043 section 3.2; CR-010 section 4. The notes whose references went (Word
removes a note with its reference), the comments whose anchors went (with their
``commentsExtended``, ``commentsIds`` and ``commentsExtensible`` entries and
their replies, and with the last comment the comment parts themselves, as Word
365 removes them: docx4j ``d18693760``, CR-010 section 12.8), bookmarks with a
start or an end in removed content (removed whole, C11), and the relationships
and parts nothing references any more (D7: images, charts, OLE objects, a
dropped section's headers and footers).
Recorded as content goes, settled by :meth:`Removals.flush` once the resolution
is done, against what still stands in every story.

Two things docx4j does not need:

* **the wildcard trees.** A comment's markers, a bookmark or a picture's
  ``r:embed`` may stand inside a text box, which this package keeps as a
  wildcard tree (CR-002's lossless MCE). Every scan and every removal here
  reads typed and wildcard nodes alike, by element name and ``w:id``; and the
  typed islands a resolution changed are written back **before** anything is
  settled (:mod:`.islands`).
* **parts not unmarshalled.** A story part nobody has read is scanned in its
  bytes (with lxml) and unmarshalled only when something in it has to go, so
  that a part the resolution never touches is still written back byte for
  byte (CR-002's round-trip rule).

A body that is not a real package's --- a dry run's copy, a detached container
--- is resolved **locally**: nothing outside its own part is touched (no note,
comment entry, relationship or part removed), because the other parts are the
real document's (CR-010 section 10.3).
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from lxml import etree

from docx4j_py.child import is_any_element
from docx4j_py.model.revisions import trees
from docx4j_py.model.revisions.story import Story

__all__ = ["Removals"]

_log = logging.getLogger(__name__)

_W = trees.W_NS


class Removals:
    """What one resolution removed, settled once it is done."""

    __slots__ = (
        "_index_cache",
        "bookmarks",
        "comments",
        "comments_whole",
        "endnotes",
        "footnotes",
        "islands",
        "local",
        "package",
        "relationships",
    )

    def __init__(self, package: Any, *, local: bool = False) -> None:
        """Start recording.

        Args:
            package: the package resolved.
            local: True for a body that is not a real package's: nothing outside
                the story parts' own content is touched.
        """
        self.package = package
        self.local = local or not _is_package(package)
        self.footnotes: set[int] = set()
        self.endnotes: set[int] = set()
        #: comments some of whose markup went: removed when nothing references
        #: them, else their orphan markers go
        self.comments: set[int] = set()
        #: comments to remove whole (their range held removed text, and is now empty)
        self.comments_whole: set[int] = set()
        self.bookmarks: dict[int, tuple[Any, set[int]]] = {}
        self.relationships: dict[int, tuple[Any, set[str]]] = {}
        #: the typed islands resolved, written back before anything is settled
        self.islands: list[Any] = []
        self._index_cache: dict[int, dict[str, set[int]]] = {}

    @classmethod
    def for_story(cls, story: Story) -> Removals:
        """A recorder for a resolution in `story`'s package."""
        return cls(story.package)

    # -- recording -----------------------------------------------------------

    def touched(self, story: Story) -> None:
        """Note that a resolution changed `story`: a typed island is written back before the flush."""
        island = story.island
        if island is not None and not island.dirty:
            island.dirty = True
            self.islands.append(island)

    def discard(self, story: Story, subtree: Any) -> None:
        """Records what `subtree`, about to be removed from a story, references."""
        self._discard(_target_of(story), story.part, subtree)

    def _discard(self, bookmark_root: Any, part: Any, subtree: Any) -> None:
        def enter(node: Any) -> bool:
            local = trees.name(node)
            if local in ("footnoteReference", "endnoteReference"):
                found = trees.marker_id(node)
                if found is not None:
                    (self.footnotes if local == "footnoteReference" else self.endnotes).add(found)
            elif local in ("commentReference", "commentRangeStart", "commentRangeEnd"):
                found = trees.marker_id(node)
                if found is not None:
                    self.comments.add(found)
            elif local in ("bookmarkStart", "bookmarkEnd") and bookmark_root is not None:
                found = trees.marker_id(node)
                if found is not None:
                    self.bookmarks.setdefault(id(bookmark_root), (bookmark_root, set()))[1].add(found)
            return True

        def relationship(rel_id: str) -> None:
            if part is not None:
                self.relationships.setdefault(id(part), (part, set()))[1].add(rel_id)

        trees.walk(subtree, enter, relationship)

    def remove_comment(self, comment_id: int | None) -> None:
        """A comment to remove whole: its range held removed content, and is now empty."""
        if comment_id is not None:
            self.comments_whole.add(comment_id)

    # -- settling ------------------------------------------------------------

    def _pending(self) -> bool:
        return bool(
            self.footnotes
            or self.endnotes
            or self.comments
            or self.comments_whole
            or self.bookmarks
            or self.relationships
        )

    def write_back(self) -> None:
        """Every typed island this resolution changed, back into its wildcard tree (newest first)."""
        while self.islands:
            self.islands.pop().write_back()

    def flush(self) -> None:
        """Settles everything recorded; what that removes in turn (a note's own picture) is settled too."""
        self.write_back()
        rounds = 0
        while self._pending() and rounds < 20:
            rounds += 1
            self._index_cache.clear()
            self._flush_comments()
            self._flush_notes()
            self._flush_bookmarks()
            self._flush_relationships()

    # -- the parts -----------------------------------------------------------

    def story_parts(self) -> list[Any]:
        """Every part whose content is a story's; none when resolving locally."""
        if self.local:
            return []
        from docx4j_py.model.revisions.lister import story_parts

        return story_parts(self.package)

    def _index(self, part: Any) -> dict[str, set[int]]:
        """``local name -> {w:id}`` of every WordprocessingML element in a part, typed or not."""
        key = id(part)
        found = self._index_cache.get(key)
        if found is not None:
            return found
        index: dict[str, set[int]] = {}
        if _unread(part):
            try:
                root = etree.fromstring(part._raw_bytes())
            except Exception:  # noqa: BLE001 - unreadable: nothing found in it
                root = None
            if root is not None:
                id_attribute = f"{{{_W}}}id"
                for element in root.iter():
                    tag = element.tag
                    if not isinstance(tag, str) or not tag.startswith(f"{{{_W}}}"):
                        continue
                    value = element.get(id_attribute)
                    if value is None:
                        continue
                    try:
                        index.setdefault(tag.rpartition("}")[2], set()).add(int(value))
                    except ValueError:
                        continue
        else:

            def enter(node: Any) -> bool:
                local = trees.name(node)
                if local:
                    value = trees.marker_id(node)
                    if value is not None:
                        index.setdefault(local, set()).add(value)
                return True

            trees.walk(_root_of(part), enter)
        self._index_cache[key] = index
        return index

    # -- comments ------------------------------------------------------------

    def _flush_comments(self) -> None:
        if not self.comments and not self.comments_whole:
            return
        candidates = set(self.comments)
        whole = set(self.comments_whole)
        self.comments.clear()
        self.comments_whole.clear()
        if self.local:
            return
        parts = self.story_parts()
        references: set[int] = set()
        starts: set[int] = set()
        ends: set[int] = set()
        for part in parts:
            index = self._index(part)
            references |= index.get("commentReference", set())
            starts |= index.get("commentRangeStart", set())
            ends |= index.get("commentRangeEnd", set())
        for comment_id in sorted(candidates):
            if comment_id in whole:
                continue
            if comment_id not in references:
                # its anchor is gone: Word keeps no comment without one (check 32)
                if (
                    self._comment_entry(comment_id) is not None
                    or comment_id in starts
                    or comment_id in ends
                ):
                    whole.add(comment_id)
            elif (comment_id in starts) != (comment_id in ends):
                # half a range: the half left goes, the comment stays on its reference
                for part in parts:
                    self._remove_from_part(
                        part,
                        {"commentRangeStart", "commentRangeEnd"},
                        {comment_id},
                        lambda item, cid=comment_id: _is_comment_marker(item, cid),
                    )
        done: set[int] = set()
        for comment_id in sorted(whole):
            self._remove_comment_whole(comment_id, parts, done)
        if done:
            self._drop_empty_comment_parts()

    def _drop_empty_comment_parts(self) -> None:
        """With the last comment a resolution takes away, the comment parts go (docx4j ``dropEmptyCommentParts``).

        As Word 365 removes them (CR-010 section 12.8; docx4j-python's artefact
        8 rejected, core-ts check 32): the comments part, its extended, ids and
        extensible parts, their relationships and content types; and the people
        part, where no tracked change is left to name an author (Word's removal
        was seen with none left). Never when resolving locally: a dry run's
        body touches no part outside its own (CR-010 section 10.3 item 3).
        """
        if self.local:
            return
        main = self.package.main_document_part
        comments_part = main.comments_part
        if comments_part is None or getattr(comments_part.contents, "comment", None):
            return
        from docx4j_py.openpackaging.parts.namespaces import Namespaces

        self._drop_part(main, comments_part, Namespaces.COMMENTS)
        self._drop_part(main, main.comments_extended_part, Namespaces.COMMENTS_EXTENDED)
        self._drop_part(main, main.comments_ids_part, Namespaces.COMMENTS_IDS)
        self._drop_part(main, main.comments_extensible_part, Namespaces.COMMENTS_EXTENSIBLE)
        if main.people_part is not None:
            from docx4j_py.model.revisions.lister import Lister

            if not Lister(self.package, False).list():
                self._drop_part(main, main.people_part, Namespaces.OFFICE_2011_PEOPLE)
        self._index_cache.clear()

    def _drop_part(self, main: Any, part: Any, relationship_type: str) -> None:
        """A part of the main document's: its relationship and shortcut go, and the part if nothing else targets it."""
        if part is None:
            return
        rels = main.relationships_part
        rel = rels.get_rel(part.part_name) if rels is not None else None
        if rel is not None:
            rels.remove_relationship(rel)
        main.set_part_shortcut(None, relationship_type)
        self._remove_if_untargeted(part)

    def _comments_root(self) -> Any:
        main = getattr(self.package, "main_document_part", None)
        part = getattr(main, "comments_part", None) if main is not None else None
        return None if part is None else part.contents

    def _comment_entry(self, comment_id: int) -> Any:
        root = self._comments_root()
        for comment in getattr(root, "comment", None) or ():
            if comment.id == comment_id:
                return comment
        return None

    def _remove_comment_whole(self, comment_id: int, parts: list[Any], done: set[int]) -> None:
        """A comment taken away: markers and reference run in every story, its entry, its extended entries, its replies."""
        if comment_id in done:
            return
        done.add(comment_id)
        for part in parts:
            self._remove_from_part(
                part,
                {"commentRangeStart", "commentRangeEnd", "commentReference"},
                {comment_id},
                lambda item, cid=comment_id: _is_comment_marker(item, cid)
                or _is_reference_run(item, cid)
                or _is_comment_reference(item, cid),
            )
        main = self.package.main_document_part
        comments_part = main.comments_part
        entry = self._comment_entry(comment_id)
        if entry is None:
            return
        para_id = _last_para_id(entry)
        self._discard(comments_part, comments_part, entry.content)
        comments = comments_part.contents.comment
        for index, item in enumerate(comments):
            if item is entry:
                del comments[index]
                break
        self._index_cache.clear()
        if para_id is None:
            return
        extended = main.comments_extended_part
        if extended is not None:
            entries = getattr(extended.contents, "comment_ex", None)
            if entries is not None:
                replies: list[str] = []
                for index in range(len(entries) - 1, -1, -1):
                    ex = entries[index]
                    if (ex.para_id or "").upper() == para_id.upper():
                        del entries[index]
                    elif (getattr(ex, "para_id_parent", None) or "").upper() == para_id.upper():
                        replies.append(ex.para_id)
                for reply in replies:
                    reply_id = self._comment_id_by_para_id(reply)
                    if reply_id is not None:
                        self._remove_comment_whole(reply_id, parts, done)
        from docx4j_py.openpackaging.parts.wml.comments import (
            remove_comment_id,
            remove_extensible,
        )

        durable_id = remove_comment_id(main.comments_ids_part, para_id)
        if durable_id:
            remove_extensible(main.comments_extensible_part, durable_id)

    def _comment_id_by_para_id(self, para_id: str | None) -> int | None:
        if not para_id:
            return None
        root = self._comments_root()
        for comment in getattr(root, "comment", None) or ():
            last = _last_para_id(comment)
            if last is not None and last.upper() == para_id.upper():
                return comment.id
        return None

    # -- notes ---------------------------------------------------------------

    def _flush_notes(self) -> None:
        """A note whose reference went is removed from its part, unless a reference to it still stands.

        Word: the editor's ``review/`` saves for footnotes, C12 and C12b for
        endnotes. Separators are never referenced and never go.
        """
        if not self.footnotes and not self.endnotes:
            return
        footnotes, endnotes = set(self.footnotes), set(self.endnotes)
        self.footnotes.clear()
        self.endnotes.clear()
        if self.local:
            return
        fn_standing: set[int] = set()
        en_standing: set[int] = set()
        for part in self.story_parts():
            index = self._index(part)
            fn_standing |= index.get("footnoteReference", set())
            en_standing |= index.get("endnoteReference", set())
        main = self.package.main_document_part
        if main.footnotes_part is not None and footnotes:
            self._prune(main.footnotes_part, main.footnotes_part.contents.footnote, footnotes, fn_standing)
        if main.endnotes_part is not None and endnotes:
            self._prune(main.endnotes_part, main.endnotes_part.contents.endnote, endnotes, en_standing)

    def _prune(self, part: Any, notes: list, ids: set[int], standing: set[int]) -> None:
        for index in range(len(notes) - 1, -1, -1):
            note = notes[index]
            kind = getattr(getattr(note, "type_value", None), "value", None)
            if kind is not None and kind != "normal":
                continue
            if note.id not in ids or note.id in standing:
                continue
            self._discard(part, part, note.content)
            del notes[index]
        self._index_cache.clear()

    # -- bookmarks -----------------------------------------------------------

    def _flush_bookmarks(self) -> None:
        """A bookmark with its start or its end in removed content goes whole (C11, C11b)."""
        if not self.bookmarks:
            return
        todo = list(self.bookmarks.values())
        self.bookmarks.clear()
        for target, ids in todo:

            def matcher(item: Any, ids: set[int] = ids) -> bool:
                return trees.name(item) in ("bookmarkStart", "bookmarkEnd") and trees.marker_id(item) in ids

            if _is_part(target):
                self._remove_from_part(target, {"bookmarkStart", "bookmarkEnd"}, ids, matcher)
            else:
                _remove_from(target, matcher)

    # -- relationships and parts ---------------------------------------------

    def _flush_relationships(self) -> None:
        """A relationship nothing in its part references goes, and its part with it when nothing targets that (D7)."""
        if not self.relationships:
            return
        todo = list(self.relationships.values())
        self.relationships.clear()
        if self.local:
            return
        from docx4j_py.openpackaging.parts.relationships_part import is_external

        package = self.package
        for part, ids in todo:
            if not _is_part(part) or package.get_part(part.part_name) is not part:
                continue
            rels = part.relationships_part
            if rels is None:
                continue
            used: set[str] = set()
            trees.walk(_root_of(part), None, used.add)
            for rel_id in sorted(ids):
                if rel_id in used:
                    continue
                rel = rels.get_relationship_by_id(rel_id)
                if rel is None:
                    continue
                target = None if is_external(rel) else rels.get_part(rel)
                rels.remove_relationship(rel)
                if target is not None:
                    self._remove_if_untargeted(target)

    def _remove_if_untargeted(self, target: Any) -> None:
        """Removes a part nothing targets any more, and what only it targeted in turn."""
        from docx4j_py.openpackaging.parts.relationships_part import is_external

        package = self.package
        if self._is_targeted(target.part_name):
            return
        own = target.relationships_part
        children: list[Any] = []
        if own is not None:
            for rel in list(own.list):
                if is_external(rel):
                    continue
                child = own.get_part(rel)
                if child is not None:
                    children.append(child)
        _log.debug("Removing %s: nothing references it now", target.part_name)
        package.parts.remove(target.part_name)
        package.content_type_manager.remove_override_content_type(target.part_name)
        for child in children:
            if package.get_part(child.part_name) is child:
                self._remove_if_untargeted(child)

    def _is_targeted(self, name: Any) -> bool:
        package = self.package
        if package.relationships_part is not None and package.relationships_part.get_rel(name) is not None:
            return True
        for part in list(package.parts.values()):
            rels = getattr(part, "relationships_part", None)
            if rels is not None and rels.get_rel(name) is not None:
                return True
        return False

    # -- removing from every list --------------------------------------------

    def _remove_from_part(self, part: Any, names: set[str], ids: set[int] | None, match: Callable[[Any], bool]) -> int:
        """Removes what `match` matches from every list of a part; a part not read is read only if it holds one."""
        if _is_part(part) and _unread(part):
            index = self._index(part)
            if not any((ids is None and name in index) or (ids is not None and index.get(name, set()) & ids) for name in names):
                return 0
        removed = _remove_from(_root_of(part), match)
        if removed:
            self._index_cache.pop(id(part), None)
        return removed

    def remove_everywhere(self, root: Any, match: Callable[[Any], bool]) -> int:
        """Removes what `match` matches from every list under `root`, typed or wildcard."""
        return _remove_from(root, match)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _is_package(package: Any) -> bool:
    if package is None:
        return False
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

    return isinstance(package, OpcPackage)


def _is_part(value: Any) -> bool:
    from docx4j_py.openpackaging.parts.part import Part

    return isinstance(value, Part)


def _unread(part: Any) -> bool:
    """True for an XML part whose tree nobody has unmarshalled."""
    return _is_part(part) and not getattr(part, "is_unmarshalled", True)


def _root_of(target: Any) -> Any:
    """The tree of a part (unmarshalling it), or the typed root itself."""
    if _is_part(target):
        return target.contents
    return target


def _target_of(story: Story) -> Any:
    """What a story's bookmarks are removed from: its part, or its root for a detached body."""
    if story.part is not None and _is_part(story.part):
        return story.part
    return story.root_owner


def _is_comment_marker(item: Any, comment_id: int) -> bool:
    return trees.name(item) in ("commentRangeStart", "commentRangeEnd") and trees.marker_id(item) == comment_id


def _is_comment_reference(item: Any, comment_id: int) -> bool:
    return trees.name(item) == "commentReference" and trees.marker_id(item) == comment_id


def _is_reference_run(item: Any, comment_id: int) -> bool:
    """A run that holds the comment's reference and nothing else."""
    if trees.name(item) != "r":
        return False
    content = item.children if is_any_element(item) else getattr(item, "content", None)
    found = False
    for child in content or ():
        if isinstance(child, str):
            continue
        if _is_comment_reference(child, comment_id):
            found = True
        elif is_any_element(child) and trees.name(child) == "rPr":
            continue
        else:
            return False
    return found


def _last_para_id(comment: Any) -> str | None:
    p_class = trees.cls("P")
    found = None
    for item in getattr(comment, "content", None) or ():
        if isinstance(item, p_class) and item.para_id:
            found = item.para_id
    return found


_list_fields_cache: dict[type, tuple[str, ...]] = {}


def _list_fields(klass: type) -> tuple[str, ...]:
    try:
        return _list_fields_cache[klass]
    except KeyError:
        pass
    import dataclasses

    try:
        names = tuple(
            field.name
            for field in dataclasses.fields(klass)
            if field.metadata.get("type") not in ("Attribute", "Attributes")
        )
    except TypeError:
        names = ()
    _list_fields_cache[klass] = names
    return names


def _remove_from(root: Any, match: Callable[[Any], bool]) -> int:
    """Removes the items `match` matches from every list under `root` --- block, run, run-content, math, wildcard."""
    import dataclasses

    removed = 0
    stack = [root]
    seen: set[int] = set()
    while stack:
        node = stack.pop()
        if node is None or isinstance(node, str) or id(node) in seen:
            continue
        seen.add(id(node))
        if is_any_element(node):
            items = node.children
            if items:
                for index in range(len(items) - 1, -1, -1):
                    if not isinstance(items[index], str) and match(items[index]):
                        del items[index]
                        removed += 1
                stack.extend(c for c in items if not isinstance(c, str))
            continue
        if not dataclasses.is_dataclass(node):
            continue
        for name in _list_fields(node.__class__):
            value = getattr(node, name, None)
            if value is None or isinstance(value, str):
                continue
            if isinstance(value, list):
                for index in range(len(value) - 1, -1, -1):
                    item = value[index]
                    if not isinstance(item, str) and match(item):
                        del value[index]
                        removed += 1
                stack.extend(value)
            else:
                stack.append(value)
    return removed
