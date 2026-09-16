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
from docx4j_py.openpackaging.save import save_package
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
        "_changes",
        "_current_change",
        "_id_rng",
        "_id_seed",
        "_para_ids_taken",
        "content_type_manager",
        "custom_xml_data_storage_parts",
        "doc_props_core_part",
        "doc_props_custom_part",
        "doc_props_extended_part",
        "load_options",
        "parts",
        "relationships_part",
        "source_part_store",
        "was_strict",
    )

    #: The pseudo part name of the package: the source of the package relationships.
    part_name = PartName.ROOT

    def __init__(self) -> None:
        """An empty package with the two default content types and no parts."""
        self.content_type_manager = ContentTypeManager.create_default()
        self.parts = Parts()
        self.custom_xml_data_storage_parts: dict[str, Any] = {}
        self.source_part_store: PartStore | None = None
        self.load_options = LoadOptions()
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

    # -- saving ------------------------------------------------------------

    def save(self, target: str | os.PathLike[str] | IO[bytes] | None = None) -> bytes | None:
        """Write the package as a zip.

        Args:
            target: a path, a binary file object, or None for bytes.

        Returns:
            The bytes when `target` is None, else None.

        Saving **over the file the package was loaded from** is supported and is
        what decided question 4 chose: the zip is written to a temporary file
        beside the target and renamed over it, because the source zip is still
        open for the lazy loads that saving itself performs. The source store is
        then reopened, so the package stays usable.
        """
        if target is None:
            sink = ZipPartSink(None)
            return save_package(self, sink)

        if isinstance(target, (str, os.PathLike)):
            path = Path(target)
            if self._is_source_path(path):
                return self._save_over(path)
            sink = ZipPartSink(str(path))
            save_package(self, sink)
            return None

        sink = ZipPartSink(target)
        save_package(self, sink)
        return None

    def save_to(self, sink: PartSink) -> Any:
        """Write the package to any sink and return what the sink produces."""
        return save_package(self, sink)

    def _is_source_path(self, path: Path) -> bool:
        store = self.source_part_store
        source = getattr(store, "path", None)
        if source is None:
            return False
        try:
            return Path(source).resolve() == path.resolve()
        except OSError:  # pragma: no cover - a path that cannot be resolved is not ours
            return False

    def _save_over(self, path: Path) -> None:
        """Temporary file beside the target, then rename. Decided question 4."""
        directory = path.parent
        handle, temporary = tempfile.mkstemp(
            prefix=path.name + ".", suffix=".tmp", dir=str(directory)
        )
        os.close(handle)
        try:
            sink = ZipPartSink(temporary)
            save_package(self, sink)
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
