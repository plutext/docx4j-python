"""``OpcPackage``: an Open Packaging Conventions package. docx4j ``OpcPackage``.

CR-002 section 5.5. The content types, the package relationships, and the parts
reachable from them.

``OpcPackage.load`` sniffs the kind of package: it reads ``[Content_Types].xml``
and ``/_rels/.rels``, follows the ``officeDocument`` relationship, and picks the
package class from that part's content type. A ``.docx`` therefore comes back as
a :class:`~docx4j_py.openpackaging.packages.wordprocessingml_package.WordprocessingMLPackage`
whichever class you called ``load`` on, and a ``.pptx`` or an ``.xlsx`` comes
back as a plain :class:`OpcPackage` until CR-002 Phase C types their main parts.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Callable, Iterable
from pathlib import Path
from typing import IO, Any

from docx4j_py.openpackaging.content_types import ContentTypeManager, ContentTypes
from docx4j_py.openpackaging.exceptions import Docx4JException, InvalidFormatException
from docx4j_py.openpackaging.load import LoadOptions, load_package
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.parts_map import Parts
from docx4j_py.openpackaging.parts.relationships_part import (
    AddPartBehaviour,
    RelationshipsPart,
)
from docx4j_py.openpackaging.parts.xml_part import XmlPart
from docx4j_py.openpackaging.save import check_before_save, save_package
from docx4j_py.openpackaging.stores import (
    MemoryPartStore,
    PartSink,
    PartStore,
    ZipPartSink,
    ZipPartStore,
    store_for,
)

__all__ = ["OpcPackage", "register_package_class"]

#: Main-part content type -> package class. Filled in by the package modules as
#: they are imported, so that this module needs no forward reference to them.
_PACKAGE_CLASSES: dict[str, type] = {}


def register_package_class(content_types: str | list[str] | tuple[str, ...], cls: type) -> None:
    """Say which package class a main part's content type means."""
    if isinstance(content_types, str):
        content_types = [content_types]
    for content_type in content_types:
        _PACKAGE_CLASSES[content_type] = cls


class OpcPackage:
    """An OPC package. docx4j ``OpcPackage``.

    It is also a :class:`RelationshipSource`: the package itself owns
    ``/_rels/.rels``, and its pseudo part name is ``/``.
    """

    __slots__ = (
        "__weakref__",
        "_assigns_para_ids",
        "_author",
        "_change_tracker",
        "_change_tracking_mode",
        "_changes",
        "_compatibility_mode",
        "_current_change",
        "_custom_xml_parts",
        "_date_time_zone",
        "_id_rng",
        "_id_seed",
        "_inserted_paragraphs",
        "_mc_load_mode",
        "_numbering_emulator",
        "_para_ids_taken",
        "_pending_removals",
        "_split_runs",
        "_split_spaces",
        "_tracked_change_date",
        "content_type_manager",
        "custom_xml_data_storage_parts",
        "doc_props_core_part",
        "doc_props_custom_part",
        "doc_props_extended_part",
        "last_save",
        "load_options",
        "parts",
        "relationships_part",
        "save_hooks",
        "source_part_store",
        "was_strict",
        "write_times_on_save",
    )

    #: The pseudo part name of the package: the source of the package relationships.
    part_name = PartName.ROOT

    def __init__(self) -> None:
        """An empty package with the two default content types and no parts."""
        self.content_type_manager = ContentTypeManager.create_default()
        self.parts = Parts()
        self.custom_xml_data_storage_parts: dict[str, Any] = {}
        # CR-003 Phase E: the content API's view over those parts, made on first
        # use. A plain container; this layer still imports nothing from the
        # content API.
        self._custom_xml_parts: Any = None
        self.source_part_store: PartStore | None = None
        self.load_options = LoadOptions()
        #: What the last save did beyond writing the tree (CR-018 section 4.7): a
        #: :class:`~docx4j_py.openpackaging.normalise.SaveReport`, each value it
        #: wrote in Word's normal form in ``normalised``. None before the first save.
        self.last_save: Any = None
        # CR-012 Phase A: this package's own mc_load_mode, over the load
        # options' (None: theirs)
        self._mc_load_mode: str | None = None
        self.was_strict = False
        self.doc_props_core_part: Any = None
        self.doc_props_extended_part: Any = None
        self.doc_props_custom_part: Any = None
        self.relationships_part: RelationshipsPart = RelationshipsPart.create_package_rels()
        self.relationships_part.source_p = self
        self.relationships_part.package = self
        # CR-003 section 3.4, determinism: the ids the content API allocates
        # (``w14:paraId`` in Phase B) come from one generator per package, and
        # it is seedable so that the same document and the same calls give the
        # same bytes. None means "derive the seed from the document's own
        # state", which is what makes a run reproducible without being asked.
        self._id_seed: int | None = None
        self._id_rng: Any = None
        # CR-003 section 3.4, decided question 4: a ChangeReport is recorded for
        # every mutating content-API call, always, because a server needs one on
        # every call anyway. The list is the caller's to clear. These are plain
        # containers: this layer still imports nothing from the content API.
        self._changes: list[Any] = []
        self._current_change: Any = None
        self._assigns_para_ids: bool | None = None
        self._para_ids_taken: dict[str, set[str]] | None = None
        # CR-003 section 3.9: who a comment (and, with Phase F, a tracked
        # change) is by. The ``Author`` class is the content layer's, so the
        # storage is here and the ``author`` property is installed by
        # ``docx4j_py.model.content.register()``, as ``body`` is.
        self._author: Any = None
        # CR-003 section 3.8: the change-tracking mode (over ``w:trackRevisions``
        # in the settings part, cached here so that a loop of edits reads the
        # part once), the tracker the content API writes revisions through, and
        # the fixed date a revision carries instead of the wall clock. Plain
        # containers again: the ``change_tracking_mode`` and
        # ``tracked_change_date`` properties are installed by
        # ``docx4j_py.model.content.register()``, as ``author`` is.
        self._change_tracking_mode: Any = None
        self._change_tracker: Any = None
        self._tracked_change_date: Any = None
        # CR-025 Phase A: the client's time zone for this package (a tzinfo, or
        # None for the process's), over which every instant it holds is shown.
        # Never saved. The ``date_time_zone`` and ``effective_date_time_zone``
        # properties are installed by ``docx4j_py.model.content.register()``,
        # as ``tracked_change_date`` is.
        self._date_time_zone: Any = None
        #: Whether a save writes the save's time into the core properties
        #: (CR-025 Phase A, the portfolio's date rule's point 6; docx4j-core-ts
        #: ``writeTimesOnSave``, docx4j's ``docx4j.dcterms.write``): False, the
        #: default, leaves ``dcterms:created`` and ``dcterms:modified`` as they
        #: are; True sets ``dcterms:modified`` to now, UTC to the second as
        #: docx4j writes it, and ``dcterms:created`` too where the core part has none
        #: and was not read from the source. A package with no core part gets
        #: none. ``last_save.times`` is the stamp written
        #: (:func:`~docx4j_py.openpackaging.save.write_core_times`).
        self.write_times_on_save = False
        # CR-003 section 3.4 (2026-09-17): the Word version the document
        # declares, over ``w:compatSetting compatibilityMode``, cached here for
        # the same reason and read the same way --- with lxml, so that a read
        # leaves the settings part byte for byte. The ``compatibility_mode``
        # property is installed by ``docx4j_py.model.content.register()``.
        self._compatibility_mode: Any = None
        # CR-003 section 16.12: ``id(tail) -> id(head)`` for every ``w:r`` this
        # package's splits have made. A rejected revision puts the two halves of
        # a split run back side by side, and this is what says they are *those*
        # two halves --- runs a document keeps apart are never joined, and
        # neither is a half beside a run it never belonged to. In memory only: a
        # reload starts with none, and a reject then leaves the runs split.
        self._split_runs: dict[int, int] = {}
        # and the ``id()`` of every ``w:p`` this package's content API inserted
        # while tracking. With the mark of an inserted paragraph sometimes on
        # the paragraph *before* it (CR-003 section 16.10), the mark no longer
        # says who made the paragraph, and a property setter has to know: a
        # paragraph this session inserted records no ``w:pPrChange``.
        self._inserted_paragraphs: set[int] = set()
        # and the ``id()`` of every ``w:t`` whose ``xml:space="preserve"`` a
        # split put there. Only those may lose it when the halves are joined:
        # an attribute the document itself wrote stays (section 16.12).
        self._split_spaces: set[int] = set()
        # CR-003 Phase H: the numbering emulator of this package
        # (``docx4j_py.model.listnumbering``), made on first use and holding the
        # list definitions and the counted labels of each story. Here for the
        # same reason as the four above: the engine holds the state, the content
        # API puts it there, and this layer imports nothing from it.
        self._numbering_emulator: Any = None
        # CR-013 Phase C: what the content API's deletions took with them (a
        # note, a comment, a relationship whose reference went), settled when the
        # write ends through the revisions engine's ``Removals``. A plain
        # container again: the content API puts it here and settles it.
        self._pending_removals: Any = None
        #: What a save runs before it writes anything (CR-019 section 4.9): each
        #: callable is given the package and returns the names of the parts it
        #: added, or None. The content API registers one where it has writes
        #: that wait for the save (a glossary entry a placeholder names), so this
        #: layer runs them without importing it. A plain list, in the order they
        #: were registered; ``package.last_save.added`` lists what they added.
        self.save_hooks: list[Callable[[OpcPackage], Iterable[str] | None]] = []

    # -- markup compatibility at block level (CR-012 Phase A) -------------

    @property
    def mc_load_mode(self) -> str:
        """What a parse of this package's parts does with block-, row- and cell-level ``mc:AlternateContent``.

        ``"preserve"`` (the default) or ``"word"``: this package's own setting
        where one was made, else :attr:`load_options`' (``LoadOptions.mc_load_mode``).
        docx4j ``OpcPackage.getMcLoadMode`` over ``McLoadMode.of``. It is
        read when each part is **first parsed**, so set it before reading a
        part: a part already parsed keeps its tree, and a tree a private reader
        parsed under the other mode is not adopted by ``contents`` (CR-002
        section 16) but parsed again.
        """
        return self._mc_load_mode or self.load_options.mc_load_mode

    @mc_load_mode.setter
    def mc_load_mode(self, value: str) -> None:
        from docx4j_py.openpackaging.mc_load_mode import check_mode

        self._mc_load_mode = check_mode(value)

    # -- the agent surface's state (CR-003 section 3.4) --------------------

    @property
    def changes(self) -> list[Any]:
        """Every ``ChangeReport`` since the last ``pkg.changes.clear()``.

        A plain list, so ``clear()``, slicing and ``len`` are Python's. CR-003
        section 3.4: recorded always, cleared by the caller.
        """
        return self._changes

    @property
    def last_change(self) -> Any:
        """The most recent ``ChangeReport``, or None before the first edit."""
        return self._changes[-1] if self._changes else None

    @property
    def assigns_para_ids(self) -> bool | None:
        """Whether a new paragraph is given a ``w14:paraId``; None decides per document.

        CR-003 section 3.4: "new paragraphs get one when the document already
        uses them (and always in a created document)". None is that rule;
        ``create_package`` sets True, and a caller who wants stable handles on a
        document that has none can set it too (or call
        ``body.ensure_para_ids()`` for the paragraphs already there).
        """
        return self._assigns_para_ids

    @assigns_para_ids.setter
    def assigns_para_ids(self, value: bool | None) -> None:
        self._assigns_para_ids = value
        self._para_ids_taken = None

    # -- deterministic ids (CR-003 section 3.4) ----------------------------

    @property
    def id_seed(self) -> int | None:
        """The seed of this package's id generator, or None for the derived one.

        The content API allocates ids of its own --- a ``w14:paraId`` for a new
        paragraph while Phase B is all there is, revision ids and content
        control ids later --- and CR-003 section 3.4 requires that the same
        document and the same calls produce the same bytes. Setting this fixes
        the sequence, which is what a test or a reproducible agent run wants;
        leaving it None derives a seed from the document's own state, so a run
        is reproducible without anyone having asked.
        """
        return self._id_seed

    @id_seed.setter
    def id_seed(self, value: int | None) -> None:
        self._id_seed = value
        self._id_rng = None

    def id_generator(self, *, derive_from: object = ()) -> Any:
        """The package's ``random.Random``, made on first use.

        Args:
            derive_from: what to derive the seed from when :attr:`id_seed` is
                None --- for paragraph ids, the ids the document already uses.
        """
        if self._id_rng is None:
            import random
            import zlib

            seed = self._id_seed
            if seed is None:
                material = repr(sorted(str(x) for x in derive_from)).encode("utf-8")
                seed = zlib.crc32(material) ^ 0x646F6378
            self._id_rng = random.Random(seed)
        return self._id_rng

    # -- the package is a relationship source ------------------------------

    @property
    def package(self) -> OpcPackage:
        """Itself. A part answers the same question, so the loader can ask either."""
        return self

    @property
    def content_type(self) -> str | None:
        """The main part's content type, which is what the package *is*."""
        main = self.get_main_part()
        return main.content_type if main is not None else None

    def get_relationships_part(self, create_if_absent: bool = True) -> RelationshipsPart:
        """``/_rels/.rels``. Always there."""
        return self.relationships_part

    def set_part_shortcut(self, part: Part, relationship_type: str) -> bool:
        """Keep the three document-properties parts. docx4j ``OpcPackage.setPartShortcut``."""
        if relationship_type == Namespaces.PROPERTIES_CORE:
            self.doc_props_core_part = part
            return True
        if relationship_type == Namespaces.PROPERTIES_EXTENDED:
            self.doc_props_extended_part = part
            return True
        if relationship_type == Namespaces.PROPERTIES_CUSTOM:
            self.doc_props_custom_part = part
            return True
        return False

    def add_target_part(
        self,
        target: Part,
        mode: AddPartBehaviour | str | None = None,
        rel_id: str | None = None,
    ) -> Any:
        """Add a part as a target of the *package* relationships."""
        target.package = self
        rel = self.relationships_part.add_part(target, mode, rel_id)
        self.set_part_shortcut(target, target.relationship_type)
        return rel

    # -- loading -----------------------------------------------------------

    @classmethod
    def load(
        cls,
        source: str | os.PathLike[str] | bytes | IO[bytes] | PartStore,
        *,
        options: LoadOptions | None = None,
    ) -> OpcPackage:
        """Load a package from a path, a directory, bytes, a file object or a store.

        The class of the result is decided by the main part's content type, not
        by the class this was called on; a subclass's ``load`` checks the result
        and raises if it is the wrong kind.
        """
        store = store_for(source)
        return load_package(store, _package_for, options)

    # -- parts -------------------------------------------------------------

    def get_part(self, part_name: PartName | str) -> Part | None:
        """The part with this name, or None. Case-insensitive."""
        try:
            return self.parts[part_name]
        except KeyError:
            return None

    def get_main_part(self) -> Part | None:
        """The part the ``officeDocument`` relationship targets."""
        rels = self.relationships_part
        for relationship_type in (Namespaces.DOCUMENT, Namespaces.DOCUMENT_STRICT):
            rel = rels.get_relationship_by_type(relationship_type)
            if rel is not None:
                return rels.get_part(rel)
        return None

    def unmarshal_all(self) -> None:
        """Unmarshal every :class:`XmlPart`, for callers who prefer no laziness."""
        for part in self.parts.parts():
            if isinstance(part, XmlPart):
                part.contents

    @property
    def skipped(self) -> list[tuple[PartName, Any]]:
        """What lenient parsing dropped, over every part that was unmarshalled.

        Each entry names its part. Empty is the good answer and the round-trip
        harness's gate (CR-002 section 8).
        """
        out: list[tuple[PartName, Any]] = []
        for part in self.parts.parts():
            if isinstance(part, XmlPart):
                out.extend((part.part_name, item) for item in part.skipped)
        rels = self.relationships_part
        out.extend((rels.part_name, item) for item in rels.skipped)
        return out

    @property
    def kept(self) -> list[tuple[PartName, Any]]:
        """The values the parse bound as the text they were read with, over every part that was unmarshalled.

        CR-018 Q16: each entry names its part and is the fork's ``KeptValue``
        (``reason``, ``qname``, ``text``, ``field_name``, ``parent_class``,
        ``path``). ``"exact"`` is a value the exact parse kept on purpose
        (``w:val=" false"``, which XSD would have trimmed and Word reads as on);
        ``"invalid"`` is one no rule converts (POI's ``"off"``, ``"Yellow"``).
        Not skipped content, and ``LoadOptions(strict=True)`` does not raise on
        it: the text is in the tree, :mod:`docx4j_py.model.values` reads it as
        Word does, and a save writes Word's reading of it (``save(normalise=)``).
        """
        out: list[tuple[PartName, Any]] = []
        for part in self.parts.parts():
            if isinstance(part, XmlPart):
                out.extend((part.part_name, item) for item in part.kept)
        rels = self.relationships_part
        out.extend((rels.part_name, item) for item in rels.kept)
        return out

    # -- saving ------------------------------------------------------------

    def save(
        self,
        target: str | os.PathLike[str] | IO[bytes] | None = None,
        *,
        normalise: bool | None = None,
        check: bool | None = None,
    ) -> bytes | None:
        """Write the package as a zip.

        Args:
            target: a path, a binary file object, or None for bytes.
            normalise: write each value the parse kept as its text in the
                normal form Word reads exactly as it read the source (CR-018
                section 4.7; :mod:`~docx4j_py.openpackaging.normalise`): an
                on/off value Word reads as on bare, one it reads as off ``0``.
                None (the default) is ``load_options.normalise_on_save``, which
                is True; False keeps the source's spelling. Only the parts the
                save re-marshals are touched, and each rewrite is listed in
                :attr:`last_save`.
            check: run the validator (:meth:`misplaced`) over the parts the
                save re-marshals first, before the target is opened (CR-020
                Phase B): what this session introduced raises
                :class:`~docx4j_py.model.content.errors.SaveCheckError` and
                nothing is written; what the source already held is saved as
                it is, listed in ``last_save.findings``. None (the default) is
                :func:`~docx4j_py.openpackaging.save.checking`'s default, off
                outside the test suite.

        Returns:
            The bytes when `target` is None, else None.

        Where :attr:`write_times_on_save` is True, the save writes its time
        into ``dcterms:modified`` first (CR-025 D4,
        :func:`~docx4j_py.openpackaging.save.write_core_times`); off by
        default, so the core properties are saved as they are.

        Saving **over the file the package was loaded from** is supported and is
        what decided question 4 chose: the zip is written to a temporary file
        beside the target and renamed over it, because the source zip is still
        open for the lazy loads that saving itself performs. The source store is
        then reopened, so the package stays usable.
        """
        # the check runs before the target is opened, so a refusal leaves it as it was
        checked = check_before_save(self, check)
        if target is None:
            sink = ZipPartSink(None)
            return save_package(self, sink, normalise=normalise, check=False, _checked=checked)

        if isinstance(target, (str, os.PathLike)):
            path = Path(target)
            if self._is_source_path(path):
                return self._save_over(path, normalise=normalise, checked=checked)
            sink = ZipPartSink(str(path))
            save_package(self, sink, normalise=normalise, check=False, _checked=checked)
            return None

        sink = ZipPartSink(target)
        save_package(self, sink, normalise=normalise, check=False, _checked=checked)
        return None

    def save_to(self, sink: PartSink, *, normalise: bool | None = None, check: bool | None = None) -> Any:
        """Write the package to any sink and return what the sink produces. `normalise` and `check` are :meth:`save`'s."""
        return save_package(self, sink, normalise=normalise, check=check)

    def misplaced(self, *, unmarshalled_only: bool = False) -> list[Any]:
        """Every finding of the validator in the package's typed XML parts (CR-020 Phase B, Q5).

        :func:`docx4j_py.traversal.misplaced` over the package: misplaced
        content (``target.misplaced``), a table cell that does not end with a
        paragraph (``cell.no_end_paragraph``) and an integer measurement past
        what Word stores (``value.out_of_range``), each a
        :class:`docx4j_py.validator.Finding` with ``to_dict()``, in part and
        document order. Nothing is raised for a finding and nothing is
        rewritten: a part never unmarshalled is read privately
        (:meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart.read_contents`)
        and is still written byte for byte; `unmarshalled_only` walks only
        the parts a save re-marshals.
        """
        from docx4j_py.validator import misplaced

        return misplaced(self, unmarshalled_only=unmarshalled_only)

    def _is_source_path(self, path: Path) -> bool:
        store = self.source_part_store
        source = getattr(store, "path", None)
        if source is None:
            return False
        try:
            return Path(source).resolve() == path.resolve()
        except OSError:  # pragma: no cover - a path that cannot be resolved is not ours
            return False

    def _save_over(self, path: Path, *, normalise: bool | None = None, checked: list[Any] | None = None) -> None:
        """Temporary file beside the target, then rename. Decided question 4."""
        directory = path.parent
        handle, temporary = tempfile.mkstemp(
            prefix=path.name + ".", suffix=".tmp", dir=str(directory)
        )
        os.close(handle)
        try:
            sink = ZipPartSink(temporary)
            save_package(self, sink, normalise=normalise, check=False, _checked=checked)
            os.replace(temporary, path)
        except BaseException:
            Path(temporary).unlink(missing_ok=True)
            raise
        store = self.source_part_store
        if isinstance(store, ZipPartStore):
            store.reopen()

    # -- lifecycle ---------------------------------------------------------

    def close(self) -> None:
        """Close the source container. The package is not usable afterwards."""
        store = self.source_part_store
        close = getattr(store, "close", None)
        if close is not None:
            close()

    def __enter__(self) -> OpcPackage:
        """Return the package; the container closes on the way out."""
        return self

    def __exit__(self, *exc_info: object) -> None:
        """Close the source container."""
        self.close()

    def __repr__(self) -> str:
        """``WordprocessingMLPackage(18 parts)``."""
        return f"{type(self).__name__}({len(self.parts)} parts)"


def _package_for(main_content_type: str | None) -> OpcPackage:
    """The package object a main part's content type calls for.

    docx4j ``ContentTypeManager.createPackage``. An unknown content type is a
    plain :class:`OpcPackage`, which loads and saves perfectly well --- it just
    has no typed main part.
    """
    # importing the package modules registers their classes
    from docx4j_py.openpackaging.packages import wordprocessingml_package  # noqa: F401

    cls = _PACKAGE_CLASSES.get(main_content_type or "")
    if cls is None:
        return OpcPackage()
    return cls()


def new_memory_package(cls: type = OpcPackage) -> OpcPackage:
    """An empty package backed by a :class:`MemoryPartStore`. Used by ``create_package``."""
    package = cls()
    package.source_part_store = MemoryPartStore()
    return package


def check_kind(package: OpcPackage, cls: type, what: str) -> None:
    """Raise unless a loaded package is of the expected class."""
    if not isinstance(package, cls):
        main = package.get_main_part()
        found = main.content_type if main is not None else "missing"
        raise InvalidFormatException(f"Not a {what} package: the main part is {found}")


_ = (Docx4JException, ContentTypes)  # re-exported for callers of this module
