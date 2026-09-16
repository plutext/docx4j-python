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
* the copy is of the *unmarshalled* tree, so a part the real package has not
  read is read (and therefore marked for re-marshalling on the real package)
  by the trial's first look at it. A dry run over ``pkg.body`` after
  ``pkg.body`` has been read costs a ``deepcopy`` of ``w:document`` and nothing
  else.
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
* ``trial.save()`` is refused: a trial is not a document.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.child import deep_copy, link_parents
from docx4j_py.model.content.errors import ContentError

__all__ = ["TrialPackage", "TrialPart", "dry_run"]


class TrialPart:
    """A part whose contents are a deep copy. Everything else is the real part."""

    __slots__ = ("_contents", "_tree", "_wrapped", "package")

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
        if hasattr(part, "set_tree"):
            self._tree = copy.deepcopy(part.tree)
        else:
            self._contents = deep_copy(part.contents)
            link_parents(self._contents)
        #: The trial package.
        self.package = package

    @property
    def contents(self) -> Any:
        """The copy."""
        return self._contents

    @property
    def tree(self) -> Any:
        """The copied lxml element, for a part the model does not type."""
        return self._tree

    def set_tree(self, tree: Any) -> None:
        """Replace the copied element; the real part is not touched."""
        self._tree = tree

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
        "_assigns_para_ids",
        "_changes",
        "_current_change",
        "_id_rng",
        "_id_seed",
        "_package",
        "_para_ids_taken",
        "_parts",
    )

    def __init__(self, package: Any) -> None:
        """Wrap a package; nothing is copied until a body is asked for."""
        self._package = package
        self._parts: dict[int, TrialPart] = {}
        self._added: list[tuple[Any, Any, Any, bool]] = []
        self._changes: list[Any] = []
        self._current_change: Any = None
        self._assigns_para_ids = package.assigns_para_ids
        self._para_ids_taken: dict[str, set[str]] | None = None
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
            removed.append(part.part_name)
        self._added.clear()
        return removed

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
        """Drop the copies and the parts the trial added. The document is as it was."""
        self.trial.discard_added_parts()
        self.trial._parts.clear()
