"""Writing a package to a container. docx4j ``io3.Save``.

CR-002 section 5.5, and the order is docx4j's:

1. ``[Content_Types].xml`` **first**, as Word writes it;
2. ``/_rels/.rels``;
3. every part reachable through relationships, depth first, each written once,
   and each part's own ``.rels`` after it (skipped when it holds nothing, as
   docx4j skips it).

What each part contributes is :attr:`Part.bytes_for_save`: a relationships part
and an unmarshalled :class:`XmlPart` are marshalled; everything else is the
bytes it was loaded with, copied from the source store **without being decoded**.
There is no dirty tracking (decided question 2): a part the caller unmarshalled
is re-marshalled even if nothing in it changed, which is docx4j's rule and is
the safe one, and the cost is bounded by what the caller chose to unmarshal.

Before any of it, the package's ``save_hooks`` run (:func:`run_save_hooks`,
CR-019 section 4.9): what the content API keeps for the save, such as the
glossary entries a placeholder names, goes into the parts then.

**The core properties' times** (CR-025 Phase A, the portfolio's date rule's
point 6). Off by default: a save leaves ``dcterms:created`` and
``dcterms:modified`` as they are. ``pkg.write_times_on_save = True`` makes each
save write its time there first (:func:`write_core_times`), as docx4j's
``docx4j.dcterms.write`` and docx4j-core-ts's ``writeTimesOnSave`` do.

**The check** (CR-020 Phase B, Q5). ``save(check=True)`` runs the validator
(:func:`docx4j_py.validator.check_for_save`) over every part the save
re-marshals, before the hooks and before the target is opened: what this
session introduced raises :class:`~docx4j_py.model.content.errors.SaveCheckError`
and nothing is written; what the source already held is saved as it is and
recorded on ``package.last_save.findings``. ``check`` None is the default of
the moment, :func:`checking`'s, which is False: the test suite turns it on
(``tests/conftest.py``), so every save the engine's own edits make is held to
the validator.
"""

from __future__ import annotations

import datetime
from collections.abc import Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from typing import TYPE_CHECKING, Any

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.part_name import CONTENT_TYPES_NAME
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.relationships_part import (
    RelationshipsPart,
    is_external,
)
from docx4j_py.openpackaging.stores import PartSink

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

__all__ = ["check_before_save", "checking", "run_save_hooks", "save_package", "saves_check", "write_core_times"]

#: What a save does when its own `check` is not given (CR-020 Q5): False, unless :func:`checking` says otherwise.
#: A context variable, so per thread and per asyncio task, as :func:`docx4j_py.child.unchecked` is.
_CHECKING: ContextVar[bool] = ContextVar("docx4j_py.openpackaging.save.checking", default=False)


@contextmanager
def checking(enabled: bool = True) -> Iterator[None]:
    """Make every save in the block whose `check` is not given run the validator first (or not), CR-020 Q5.

    ``tests/conftest.py`` turns it on for the whole suite; a test that saves a
    tree it built invalid on purpose turns it off around that save, or passes
    ``check=False``. The default outside any block is off.
    """
    token = _CHECKING.set(enabled)
    try:
        yield
    finally:
        _CHECKING.reset(token)


def saves_check(check: bool | None = None) -> bool:
    """Whether a save given `check` runs the validator: `check` itself, or :func:`checking`'s default for None."""
    return _CHECKING.get() if check is None else bool(check)


def check_before_save(package: OpcPackage, check: bool | None = None) -> list[Any] | None:
    """The check a save makes before it writes anything: None where it does not run, else what the source held.

    Raises:
        SaveCheckError: what this session introduced
            (:func:`docx4j_py.validator.check_for_save`).
    """
    if not saves_check(check):
        return None
    from docx4j_py.validator import check_for_save

    return check_for_save(package)


def save_package(
    package: OpcPackage,
    sink: PartSink,
    *,
    normalise: bool | None = None,
    check: bool | None = None,
    _checked: list[Any] | None = None,
) -> Any:
    """Write a package to a sink and return whatever the sink produces.

    `normalise` (CR-018 section 4.7): write each value a parse kept as its text
    in Word's normal form, for this save only
    (:func:`~docx4j_py.openpackaging.normalise.normalised_package`). None is the
    package's ``load_options.normalise_on_save``. What it rewrote is
    ``package.last_save``.

    `check` (CR-020 Phase B): run the validator first, before the hooks and
    before anything is put in the sink (:func:`check_before_save`); None is
    :func:`checking`'s default. `_checked` is the check's result where the
    caller ran it already, before it opened the sink
    (:meth:`~docx4j_py.openpackaging.packages.opc_package.OpcPackage.save`).
    """
    from docx4j_py.openpackaging.normalise import normalised_package

    findings = _checked if _checked is not None else check_before_save(package, check)
    if normalise is None:
        normalise = bool(getattr(getattr(package, "load_options", None), "normalise_on_save", True))
    times = write_core_times(package)
    added = run_save_hooks(package)
    with normalised_package(package, enabled=normalise) as report:
        sink.put(
            CONTENT_TYPES_NAME,
            package.content_type_manager.to_bytes(),
            content_type=ContentTypes.CONTENT_TYPES_PART,
        )
        saver = _Saver(package, sink)
        saver.save_rels(package.relationships_part)
        result = sink.finish()
    report.added = added
    report.times = times
    report.checked = findings is not None
    report.findings = list(findings or ())
    if hasattr(package, "last_save"):
        package.last_save = report
    return result


def _now() -> datetime.datetime:
    """Now, in UTC: the save's time (a test replaces it)."""
    return datetime.datetime.now(datetime.UTC)


def write_core_times(package: OpcPackage) -> str | None:
    """Where ``pkg.write_times_on_save``: the save's time into the core properties; the stamp written, or None.

    CR-025 D4, after docx4j's ``DocPropsCorePart.writeTimes()`` (17.3.3, its
    CR-036; under ``docx4j.dcterms.write``) and docx4j-core-ts's
    ``writeTimesOnSave``: ``dcterms:modified`` is now, and ``dcterms:created``
    too where the core part has none and was not read from the source (a part
    ``create_package`` or the caller made). In UTC with a ``Z``, to the
    second, as docx4j and core-ts write them and as ``create_package`` writes
    ``created``, so that ``modified`` is never before it; Word writes the
    minute (CR-025 section 2, Q3). A package with no core part gets none, as
    docx4j leaves it.
    ``cp:lastModifiedBy`` is not touched (docx4j's is ``docx4j.dc.write``).
    The part is unmarshalled for the write, so this save marshals it.
    """
    if not getattr(package, "write_times_on_save", False):
        return None
    part = getattr(package, "doc_props_core_part", None)
    if part is None:
        return None
    core = part.contents
    if core is None:
        return None
    from docx4j_xsdata.models.datatype import XmlDateTime

    from docx4j_py.docprops.dcterms import W3CDTF

    stamp = XmlDateTime.from_datetime(_now().astimezone(datetime.UTC).replace(microsecond=0))
    if getattr(core, "created", None) is None and not _from_source(package, part):
        core.created = W3CDTF(value=stamp)
    core.modified = W3CDTF(value=stamp)
    return str(stamp)


def _from_source(package: OpcPackage, part: Part) -> bool:
    """Whether `part` was read from the package's source container, rather than made in this session."""
    store = getattr(package, "source_part_store", None)
    if store is None:
        return False
    try:
        return bool(store.has(part.part_name.store_name))
    except Exception:  # noqa: BLE001 - a store that cannot say holds nothing of ours
        return False


def run_save_hooks(package: OpcPackage) -> list[str]:
    """Run the package's :attr:`~OpcPackage.save_hooks`, in order, before anything is written; what they added.

    CR-019 section 4.9: the content API's writes that wait for the save (a
    glossary entry a placeholder names) reach the package here, so that this
    layer imports nothing from it. Each hook is given the package and returns
    the names of the parts it added, or None.
    """
    added: list[str] = []
    for hook in tuple(getattr(package, "save_hooks", None) or ()):
        names = hook(package)
        if names:
            added.extend(str(name) for name in names)
    return added


class _Saver:
    """The recursion. docx4j's ``Save`` methods, as one object."""

    __slots__ = ("handled", "package", "sink")

    def __init__(self, package: OpcPackage, sink: PartSink) -> None:
        self.package = package
        self.sink = sink
        self.handled: set[str] = set()

    def save_rels(self, rels: RelationshipsPart) -> None:
        """Write a relationships part, then every part it targets."""
        self.sink.put(
            rels.part_name.store_name,
            rels.bytes_for_save,
            content_type=ContentTypes.RELATIONSHIPS_PART,
        )
        for rel in rels.list:
            if rel.type_value == Namespaces.HYPERLINK or is_external(rel):
                continue
            name = rels.resolve_target(rel)
            part = self.package.get_part(name)
            if part is None:
                raise Docx4JException(
                    f"Part {name} (target of {rel.id} in {rels.part_name}) is not in the package"
                )
            self.save_part(part)

    def save_part(self, part: Part) -> None:
        """Write one part and, after it, its own relationships."""
        key = part.part_name.key
        if key in self.handled:
            return
        self.handled.add(key)
        try:
            data = part.bytes_for_save
        except Docx4JException:
            raise
        except Exception as exc:  # noqa: BLE001 - name the part
            raise Docx4JException(f"Problem saving part {part.part_name}: {exc}") from exc
        self.sink.put(
            part.part_name.store_name,
            data,
            content_type=part.content_type,
            compress=part.compress,
        )
        own = part.relationships_part
        if own is not None and len(own.list) > 0:
            self.save_rels(own)
