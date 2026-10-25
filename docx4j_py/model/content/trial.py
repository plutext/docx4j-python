"""``with pkg.dry_run() as trial:`` --- edits applied to a copy and thrown away.

CR-003 section 3.4. A server offers a preview, and an agent checks a
``replace_text`` count before committing to it, by making the calls against a
*trial* package: a package-like object whose bodies are views over
:func:`~docx4j_py.child.deep_copy` copies of the affected parts' trees. The
:class:`~docx4j_py.model.content.reports.ChangeReport`\\ s land in
``trial.changes``; on the way out the copies are dropped and the real package is
byte for byte what it was.

::

    with pkg.dry_run() as trial:
        count = trial.body.replace_text("colour", "color")
        trial.last_change            # what it would do
    pkg.body.replace_text("colour", "color")   # commit, having seen the number

What it copies and what it does not:

* a part's tree is copied the **first time the trial asks for its body**, and
  only that part's; a document whose header is never touched costs nothing.
* the copy is of the live tree where the real package has unmarshalled the
  part, and otherwise of its private reading (``read_contents``, CR-002
  section 16), so a trial's first look at a part the real package has not
  read leaves that part unread and saved byte for byte. Since CR-013 Phase A
  every ``Paragraph.text`` asks the styles which runs they hide, so a trial
  copies ``styles.xml`` the first time it reads a paragraph. A dry run over
  ``pkg.body`` after ``pkg.body`` has been read costs a ``deepcopy`` of
  ``w:document`` and nothing else.
* **a part added during the trial is un-added on the way out.** Phase D added
  none; Phase C's ``insert_inline_picture`` and ``insert_ooxml`` add image
  parts to the real package's part map, which the trial shares, so they report
  what they added through :meth:`TrialPackage.note_added_part` and
  :meth:`TrialPackage.discard_added_parts` removes the part, its relationship
  and its content-type entry again when the block ends. A part a trial
  **creates** for another reason --- the ``numbering.xml`` Phase K's
  ``insert_markdown`` makes when a list needs one --- is still not un-created,
  because nothing reports it; the docstring of anything that adds a part says
  which of the two it is.
* **a relationship added with no part behind it** --- the external hyperlink
  relationship ``Range.hyperlink`` writes (CR-003 section 21) --- is reported
  through :func:`note_added_relationship` and taken away again by
  :meth:`TrialPackage.discard_added_relationships`, with the relationships part
  itself when the call created it, and the part's id allocator put back so that
  the commit after a trial allocates the id the trial reported.
* ``trial.save()`` is refused: a trial is not a document. So a glossary entry
  a placeholder in a trial names (CR-019 section 4.9) waits in the trial's own
  ``glossary_entries``, which start with the real package's waiting ones, and
  goes with it.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.child import deep_copy, link_parents
from docx4j_py.model.content.errors import ContentError

__all__ = ["TrialPackage", "TrialPart", "dry_run", "note_added_relationship"]


class TrialPart:
    """A part whose contents are a deep copy. Everything else is the real part."""

    __slots__ = ("_contents", "_modified", "_tree", "_wrapped", "package")

    def __init__(self, part: Any, package: Any) -> None:
        """Copy the part's tree; the copy is what this part's contents are.

        An lxml part --- a
        :class:`~docx4j_py.openpackaging.parts.default_xml_part.DefaultXmlPart`,
        which is what ``w16cid:commentsIds`` and ``w16cex:commentsExtensible``
        are (CR-001 section 13.5) --- is copied with ``copy.deepcopy`` of its
        element instead, so that a trial that edits one leaves the real part
        alone. Parsing the real part to copy it costs nothing in bytes: lxml
        writes back what it read, which a test pins.
        """
        import copy

        #: The real part. :func:`~docx4j_py.model.content.addresses.part_kind`
        #: reads it, so a trial header still reports ``header:rId8``.
        self._wrapped = part
        self._contents: Any = None
        self._tree: Any = None
        self._modified = False
        if hasattr(part, "set_tree"):
            self._tree = copy.deepcopy(part.tree)
        else:
            # the real part's private reading where it is not unmarshalled, so that
            # the trial's look leaves it unread (a reader never writes into it, and
            # this is a copy of it); the live tree where it is
            read = getattr(part, "read_contents", None)
            unread = not getattr(part, "is_unmarshalled", True) and callable(read)
            self._contents = deep_copy(read() if unread else part.contents)
            link_parents(self._contents)
        #: The trial package.
        self.package = package

    @property
    def contents(self) -> Any:
        """The copy."""
        return self._contents

    def read_contents(self) -> Any:
        """The copy, as a reader that must not unmarshal the real part reads it (the text model, CR-011)."""
        return self._contents

    def header_parts(self) -> list[Any]:
        """The trial's copies of the headers the real part relates to."""
        return [self.package.trial_part(part) for part in self._wrapped.header_parts()]

    def footer_parts(self) -> list[Any]:
        """The trial's copies of the footers the real part relates to."""
        return [self.package.trial_part(part) for part in self._wrapped.footer_parts()]

    @property
    def tree(self) -> Any:
        """The copied lxml element, for a part the model does not type."""
        return self._tree

    def set_tree(self, tree: Any) -> None:
        """Replace the copied element; the real part is not touched."""
        self._tree = tree
        self._modified = True

    @property
    def is_modified(self) -> bool:
        """Whether the trial's copy has been edited (CR-003 Phase E)."""
        return self._modified

    def mark_modified(self) -> None:
        """Adopt the **copy**; spelled out so it does not reach the real part.

        ``__getattr__`` would otherwise hand this to the wrapped part and make a
        trial's node edit re-marshal the real one.
        """
        self._modified = True

    @property
    def body_element(self) -> Any:
        """The copy's ``w:body``, for a main document part."""
        return getattr(self._contents, "body", self._contents)

    @property
    def is_unmarshalled(self) -> bool:
        """True: a trial part is a tree by construction."""
        return True

    @property
    def footnotes_part(self) -> Any:
        """The trial's copy of the footnotes part, for a document part."""
        return self.package.trial_part(self._wrapped.footnotes_part)

    @property
    def endnotes_part(self) -> Any:
        """The trial's copy of the endnotes part."""
        return self.package.trial_part(self._wrapped.endnotes_part)

    @property
    def comments_part(self) -> Any:
        """The trial's copy of the comments part."""
        return self.package.trial_part(self._wrapped.comments_part)

    @property
    def comments_extended_part(self) -> Any:
        """The trial's copy of ``/word/commentsExtended.xml`` (CR-003 Phase G)."""
        return self.package.trial_part(self._wrapped.comments_extended_part)

    @property
    def comments_ids_part(self) -> Any:
        """The trial's copy of ``/word/commentsIds.xml``, an lxml part."""
        return self.package.trial_part(self._wrapped.comments_ids_part)

    @property
    def comments_extensible_part(self) -> Any:
        """The trial's copy of ``/word/commentsExtensible.xml``, an lxml part."""
        return self.package.trial_part(self._wrapped.comments_extensible_part)

    @property
    def people_part(self) -> Any:
        """The trial's copy of ``/word/people.xml``."""
        return self.package.trial_part(self._wrapped.people_part)

    def get_xml(self) -> str:
        """The copy as XML, marshalled exactly as the real part would be."""
        if self._tree is not None:
            from lxml import etree

            from docx4j_py.openpackaging.parts.xml_part import XML_DECLARATION

            return (XML_DECLARATION + etree.tostring(self._tree, encoding="utf-8")).decode("utf-8")
        return self._wrapped._marshal(self._contents).decode("utf-8")

    def __getattr__(self, name: str) -> Any:
        """Anything else --- the part name, the relationships --- is the real part's."""
        return getattr(self._wrapped, name)

    def __repr__(self) -> str:
        """``<TrialPart /word/document.xml>``."""
        return f"<TrialPart {self._wrapped.part_name}>"


class TrialPackage:
    """A package-like object over copies of the parts a trial touches."""

    __slots__ = (
        "__weakref__",
        "_added",
        "_added_relationships",
        "_assigns_para_ids",
        "_author",
        "_change_tracker",
        "_change_tracking_mode",
        "_changes",
        "_compatibility_mode",
        "_current_change",
        "_custom_xml_parts",
        "_glossary_entries",
        "_id_rng",
        "_id_seed",
        "_inserted_paragraphs",
        "_numbering_emulator",
        "_package",
        "_para_ids_taken",
        "_parts",
        "_property_resolver",
        "_split_runs",
        "_split_spaces",
        "_tracked_change_date",
    )

    def __init__(self, package: Any) -> None:
        """Wrap a package; nothing is copied until a body is asked for."""
        self._package = package
        self._parts: dict[int, TrialPart] = {}
        #: The trial's own custom XML collection, over its copies (Phase E).
        self._custom_xml_parts: Any = None
        #: The trial's own glossary entries (CR-019 section 4.9): the real
        #: package's waiting ones copied in, those the trial names thrown away
        #: with it, the glossary read as the real package's.
        self._glossary_entries: Any = None
        self._added: list[tuple[Any, Any, Any, bool]] = []
        self._added_relationships: list[tuple[Any, Any, bool, int | None]] = []
        self._changes: list[Any] = []
        self._current_change: Any = None
        self._assigns_para_ids = package.assigns_para_ids
        self._para_ids_taken: dict[str, set[str]] | None = None
        # who a trial's comments are by: the real package's identity unless the
        # trial is given one of its own, which stays with the trial
        self._author: Any = None
        # change tracking is the trial's own too (CR-003 Phase F): a trial that
        # turns tracking on writes ``w:trackRevisions`` into its **copy** of the
        # settings part, and the real document's mode is not touched
        self._change_tracking_mode: Any = None
        self._change_tracker: Any = None
        # and so is the compatibility mode (CR-003 section 17.11): a trial that
        # raises it writes into the trial's **copy** of the settings part
        self._compatibility_mode: Any = None
        self._tracked_change_date: Any = package.tracked_change_date
        #: The halves the trial's own splits made (CR-003 section 16.12).
        self._split_runs: dict[int, int] = {}
        #: The paragraphs the trial inserted (CR-003 section 16.12).
        self._inserted_paragraphs: set[int] = set()
        #: The ``w:t``\ s whose ``xml:space`` the trial's splits set.
        self._split_spaces: set[int] = set()
        #: The trial's own numbering emulator (CR-003 Phase H), over the trial's
        #: copy of ``/word/numbering.xml``: a trial that starts a list counts in
        #: its own definitions and leaves the real package's emulator alone.
        self._numbering_emulator: Any = None
        #: The trial's own property resolver (CR-011 Phase D), over its copy of
        #: ``/word/styles.xml``: a style a trial's ``insert_markdown`` adds is
        #: the trial's, and the text model reads the trial's headings by it.
        self._property_resolver: Any = None
        # a trial allocates ids from its own generator, seeded like the real
        # one, so that a trial and the commit that follows it agree
        self._id_seed = package.id_seed
        self._id_rng: Any = None

    # -- the parts ---------------------------------------------------------

    def trial_part(self, part: Any) -> Any:
        """The trial's copy of a part, made on first use."""
        if part is None:
            return None
        if isinstance(part, TrialPart):
            return part
        key = id(part)
        found = self._parts.get(key)
        if found is None:
            found = TrialPart(part, self)
            self._parts[key] = found
        return found

    @property
    def main_document_part(self) -> Any:
        """The trial's copy of the main document part."""
        return self.trial_part(self._package.main_document_part)

    def get_main_document_part(self) -> Any:
        """docx4j's spelling of :attr:`main_document_part`."""
        return self.main_document_part

    def header_parts(self) -> list[Any]:
        """The trial's copies of the header parts."""
        return [self.trial_part(part) for part in self._package.header_parts()]

    def footer_parts(self) -> list[Any]:
        """The trial's copies of the footer parts."""
        return [self.trial_part(part) for part in self._package.footer_parts()]

    def get_property_resolver(self, refresh: bool = False) -> Any:
        """The trial's own :class:`~docx4j_py.model.property_resolver.PropertyResolver`, over its copies.

        Spelled out rather than left to ``__getattr__``, which would answer
        from the real package's styles: the text model's reading of a trial
        (``to_markdown``, ``outline``) resolves a paragraph's outline level
        through the styles the trial has, a style the trial added among them.
        """
        from docx4j_py.model.property_resolver import PropertyResolver

        if self._property_resolver is None:
            self._property_resolver = PropertyResolver(self)
        elif refresh:
            self._property_resolver.refresh()
        return self._property_resolver

    def refresh_property_resolver(self) -> Any:
        """Re-read the trial's copy of the styles part into its resolver, if one has been built."""
        if self._property_resolver is not None:
            self._property_resolver.refresh()
        return self._property_resolver

    @property
    def style_definitions_part(self) -> Any:
        """The trial's copy of ``/word/styles.xml``.

        Spelled out for the same reason the agent surface is, and for one more:
        CR-003 Phase K's ``insert_markdown`` *writes* to this part when the
        markdown needs a style the document does not define, and a trial must
        not leave that behind on the real document.
        """
        return self.trial_part(self._package.style_definitions_part)

    @property
    def numbering_definitions_part(self) -> Any:
        """The trial's copy of ``/word/numbering.xml``, or None.

        A trial that *creates* the part cannot un-create it (CR-003 section
        12.5): the new part goes into the real package's map, empty, and what
        the trial wrote into it goes with the copy.
        """
        return self.trial_part(self._package.numbering_definitions_part)

    @property
    def parts(self) -> Any:
        """The real package's part map; what a trial adds to it is undone on the way out."""
        return self._package.parts

    def note_added_part(
        self, part: Any, relationship: Any, source: Any, *, added_content_type: bool = True
    ) -> None:
        """Record a part a trial verb added, so :meth:`discard_added_parts` can undo it.

        Called by :func:`docx4j_py.model.content.picture.note_added_part`, which
        is a no-op on a real package. CR-003 section 12.5 asked for the question
        to be answered; this answers it with "nothing is left behind".
        """
        self._added.append((part, relationship, source, added_content_type))

    def discard_added_parts(self) -> list[Any]:
        """Remove the parts this trial added; returns their names, newest first."""
        removed: list[Any] = []
        package = self._package
        for part, relationship, source, added_content_type in reversed(self._added):
            rels = getattr(source, "relationships_part", None) if source is not None else None
            if rels is not None and relationship is not None:
                rels.remove_relationship(relationship)
            package.parts.remove(part.part_name)
            if added_content_type:
                package.content_type_manager.remove_override_content_type(part.part_name)
            # a part with a well-known relationship type left a shortcut behind
            # (``main.comments_part`` and its four siblings, CR-003 Phase G);
            # clearing it is what stops the next call finding a part the package
            # no longer holds
            shortcut = getattr(source, "set_part_shortcut", None)
            if shortcut is not None and part.relationship_type:
                shortcut(None, part.relationship_type)
            # a custom XML part is indexed by its ``ds:itemID`` as well
            # (CR-003 Phase E), and an index entry left behind would make
            # ``pkg.custom_xml_parts`` report a part the package no longer has
            index = package.custom_xml_data_storage_parts
            for key in [key for key, value in index.items() if value is part]:
                del index[key]
            removed.append(part.part_name)
        self._added.clear()
        return removed

    def note_added_relationship(
        self, relationships_part: Any, relationship: Any, *, created: bool, allocator: int | None
    ) -> None:
        """Record a relationship a trial verb added with no part behind it.

        Called by :func:`note_added_relationship`, which is a no-op on a real
        package; :meth:`discard_added_relationships` undoes it.
        """
        self._added_relationships.append((relationships_part, relationship, created, allocator))

    def discard_added_relationships(self) -> int:
        """Remove the relationships this trial added; returns how many.

        Newest first, so that each relationships part's id allocator ends where
        it was before the trial's first one. A relationships part the trial
        created is detached from its source part again once it is empty.
        """
        count = 0
        for rels, relationship, created, allocator in reversed(self._added_relationships):
            rels.remove_relationship(relationship)
            if allocator is not None:
                # the allocator only counts upwards; putting it back is what
                # makes the commit after the trial allocate the trial's id
                rels._next_id = allocator
            source = getattr(rels, "source_p", None)
            if created and len(rels) == 0 and getattr(source, "relationships_part", None) is rels:
                source.relationships_part = None
            count += 1
        self._added_relationships.clear()
        return count

    # -- the agent surface -------------------------------------------------
    #
    # These have to be spelled out rather than left to ``__getattr__``: the real
    # package has them too, and delegating would answer about the real document.

    @property
    def body(self) -> Any:
        """The trial's body: a view over the copy of ``word/document.xml``."""
        from docx4j_py.model.content.body import body_of

        return body_of(self)

    def outline(self, **options: Any) -> Any:
        """The trial's outline, over the copies."""
        from docx4j_py.model.content.reports import package_outline

        return package_outline(self, **options)

    def describe(self) -> Any:
        """The trial's description. The parts are the real package's."""
        from docx4j_py.model.content.describe import describe_package

        return describe_package(self)

    def element_at(self, address: str) -> Any:
        """The block at an address, in the trial's copies."""
        from docx4j_py.model.content.addresses import element_at

        return element_at(self.body, address)

    def paragraph_at(self, address: str | None = None, **options: Any) -> Any:
        """The paragraph at an address, in the trial's copies."""
        from docx4j_py.model.content.addresses import paragraph_at

        return paragraph_at(self.body, address, **options)

    def find(self, text: str, **options: Any) -> list[Any]:
        """Search the trial's copies."""
        from docx4j_py.model.content.reports import find_in

        return find_in(self.body, text, **options)

    def bodies(self) -> list[Any]:
        """Every body of the trial: the copies, not the real parts'."""
        from docx4j_py.model.content.addresses import package_bodies

        return list(package_bodies(self))

    def to_markdown(self, **options: Any) -> str:
        """The trial's markdown, over the copies (CR-003 Phase K)."""
        return self.body.to_markdown(**options)  # type: ignore[no-any-return]

    def markdown_budget(self, max_chars: Any = None, **options: Any) -> Any:
        """The trial's markdown with the truncation flag."""
        return self.body.markdown_budget(max_chars, **options)

    def insert_markdown(self, markdown: str, **options: Any) -> Any:
        """Insert markdown into the trial's copy of the main document part.

        A markdown fragment with a list needs a numbering part, and a part the
        trial adds is added to the **real** package's part map and stays there
        (12.5): a trial of such a fragment leaves an empty ``numbering.xml``
        behind, and the abstract numbering it wrote goes with the copy.
        """
        return self.body.insert_markdown(markdown, **options)

    @property
    def custom_xml_parts(self) -> Any:
        """The trial's custom XML parts, over **copies** of the real ones.

        Spelled out rather than left to ``__getattr__`` for the reason the rest
        of the agent surface is: the real package has the member too, and
        delegating would edit the real document. A part is copied the first time
        the trial asks the collection for it (CR-003 section 14.4), and a part
        the trial adds is un-added on the way out.
        """
        from docx4j_py.model.customxml.parts import custom_xml_parts_of

        return custom_xml_parts_of(self)

    @property
    def glossary_entries(self) -> Any:
        """The trial's own glossary entries (CR-019 section 4.9).

        Spelled out rather than left to ``__getattr__``, which would hand the
        trial the real package's, and a placeholder the trial writes would then
        put its entry into the real document's glossary on its next save.
        """
        from docx4j_py.model.content.placeholders import glossary_entries_of

        return glossary_entries_of(self)

    def dry_run(self) -> Any:
        """A trial of a trial. Allowed, and as cheap as the first."""
        return dry_run(self)

    @property
    def changes(self) -> list[Any]:
        """The reports of what the trial did. The real package's are untouched."""
        return self._changes

    @property
    def last_change(self) -> Any:
        """The trial's last report, or None."""
        return self._changes[-1] if self._changes else None

    @property
    def assigns_para_ids(self) -> bool | None:
        """As the real package's."""
        return self._assigns_para_ids

    @assigns_para_ids.setter
    def assigns_para_ids(self, value: bool | None) -> None:
        self._assigns_para_ids = value
        self._para_ids_taken = None

    @property
    def author(self) -> Any:
        """Who the trial's comments are by: the real package's, unless set here."""
        return self._author if self._author is not None else self._package.author

    @author.setter
    def author(self, value: Any) -> None:
        from docx4j_py.model.content.comments import Author

        self._author = Author(value) if isinstance(value, str) else value

    @property
    def change_tracking_mode(self) -> str:
        """The trial's mode: the real package's, until the trial sets its own.

        Setting it here writes ``w:trackRevisions`` into the trial's **copy** of
        ``/word/settings.xml``, so a dry run of "turn tracking on and edit"
        leaves the real document's settings part exactly as it was.
        """
        from docx4j_py.model.content.tracking import mode_of

        if self._change_tracking_mode is None:
            return str(self._package.change_tracking_mode)
        return str(mode_of(self))

    @change_tracking_mode.setter
    def change_tracking_mode(self, value: str | None) -> None:
        from docx4j_py.model.content.tracking import set_mode

        set_mode(self, None if value is None else str(value))

    @property
    def compatibility_mode(self) -> int:
        """The trial's mode: the real package's, until the trial sets its own.

        Setting it here writes ``w:compatSetting`` into the trial's **copy** of
        ``/word/settings.xml``, so a dry run of "raise the mode and edit" leaves
        the real document's settings part exactly as it was.
        """
        from docx4j_py.model.content.compatibility import mode_of

        if self._compatibility_mode is None:
            return int(self._package.compatibility_mode)
        return int(mode_of(self))

    @compatibility_mode.setter
    def compatibility_mode(self, value: Any) -> None:
        from docx4j_py.model.content.compatibility import set_mode

        set_mode(self, value)

    @property
    def document_settings_part(self) -> Any:
        """``/word/settings.xml``: the **real** part until the trial writes to it.

        A read --- ``describe()``'s ``tracking_on``, the mode getter --- reads
        the real part's bytes with lxml and leaves it as it was (CR-003 section
        15.3's rule, applied to this part). Only
        :meth:`writable_settings_part`, which the mode setter calls, makes the
        copy; from then on this answers with it.
        """
        real = self._package.document_settings_part
        if real is None:
            return None
        return self._parts.get(id(real), real)

    def writable_settings_part(self) -> Any:
        """The trial's **copy** of the settings part, made now. The mode setter's."""
        return self.trial_part(self._package.document_settings_part)

    @property
    def tracked_change_date(self) -> Any:
        """The trial's fixed revision date; the real package's by default."""
        return self._tracked_change_date

    @tracked_change_date.setter
    def tracked_change_date(self, value: Any) -> None:
        self._tracked_change_date = value

    def get_tracked_changes(self) -> list[Any]:
        """The trial's tracked changes, over the copies."""
        return list(self.body.get_tracked_changes())

    @property
    def id_seed(self) -> int | None:
        """The trial's own seed, taken from the real package's."""
        return self._id_seed

    @id_seed.setter
    def id_seed(self, value: int | None) -> None:
        self._id_seed = value
        self._id_rng = None

    def id_generator(self, *, derive_from: object = ()) -> Any:
        """A generator of the trial's own, seeded as the real one is."""
        if self._id_rng is None:
            import random
            import zlib

            seed = self._id_seed
            if seed is None:
                material = repr(sorted(str(x) for x in derive_from)).encode("utf-8")
                seed = zlib.crc32(material) ^ 0x646F6378
            self._id_rng = random.Random(seed)
        return self._id_rng

    def save(self, *args: Any, **kwargs: Any) -> Any:
        """Refused: a trial is discarded, not written."""
        raise ContentError(
            "a dry run cannot be saved; it is discarded when the block ends",
            code="dry_run.save",
            hint="leave the with block and make the same calls on the real package",
        )

    def __getattr__(self, name: str) -> Any:
        """Everything else --- ``describe``, ``outline``, the part shortcuts."""
        return getattr(self._package, name)

    def __repr__(self) -> str:
        """``<TrialPackage over WordprocessingMLPackage(18 parts), 2 changes>``."""
        return f"<TrialPackage over {self._package!r}, {len(self._changes)} changes>"


class dry_run:
    """``with pkg.dry_run() as trial:``. See this module's docstring."""

    __slots__ = ("package", "trial")

    def __init__(self, package: Any) -> None:
        """Prepare a trial over `package`. Nothing is copied yet."""
        self.package = package
        self.trial = TrialPackage(package)

    def __enter__(self) -> TrialPackage:
        """The trial package."""
        return self.trial

    def __exit__(self, *exc_info: object) -> None:
        """Drop the copies and what the trial added. The document is as it was."""
        self.trial.discard_added_relationships()
        self.trial.discard_added_parts()
        self.trial._parts.clear()


def note_added_relationship(
    package: Any,
    relationships_part: Any,
    relationship: Any,
    *,
    created: bool,
    allocator: int | None,
) -> None:
    """Tell a trial about a relationship added during it, so it can be un-added.

    The counterpart, for a relationship with no part behind it, of
    :func:`docx4j_py.model.content.picture.note_added_part`; a no-op on a real
    package.

    Args:
        package: the body's package --- a :class:`TrialPackage` or the real one.
        relationships_part: the ``.rels`` the relationship went into.
        relationship: the relationship.
        created: whether this call created `relationships_part`.
        allocator: the part's id allocator before the call, to be put back.
    """
    note = getattr(package, "note_added_relationship", None)
    if note is not None:
        note(relationships_part, relationship, created=created, allocator=allocator)
