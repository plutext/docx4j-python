"""``pkg.compatibility_mode``: which Word a document says it is written for.

CR-003 section 3.4 (the identity settings, beside ``pkg.author`` and
``pkg.change_tracking_mode``), added 2026-09-17. Word stores it in
``/word/settings.xml`` as one of its own compatibility settings::

    <w:compat>
      <w:compatSetting w:name="compatibilityMode"
                       w:uri="http://schemas.microsoft.com/office/word"
                       w:val="15"/>
      …
    </w:compat>

and the value is a **Word version**, not a file-format version: 11 is Word 2003,
12 Word 2007, 14 Word 2010, and 15 Word 2013 --- which is also what Word 2016,
2019 and 365 write, there being no 16. A document with no setting at all is
what Word assumes it to be: **12**, Word 2007, which is why it shows
*Compatibility Mode* in the title bar for one.

::

    pkg.compatibility_mode                 # 14 for a Word 2010 document
    pkg.compatibility_mode = 15            # target Word 2013 and later

**A read does not unmarshal the settings part.** The setting is read with lxml
from the bytes the part would be saved as, exactly as ``change_tracking_mode``
and ``describe()`` read ``w:trackRevisions``, and the answer is cached on the
package, so a document whose mode is only read keeps ``/word/settings.xml`` byte
for byte. Setting it unmarshals the part, as changing it means, and records a
:class:`~docx4j_py.model.content.reports.ChangeReport` naming
``/word/settings.xml``.

**Nothing is downgraded or upgraded silently.** A verb that writes a feature
Word 2013 introduced --- a resolved comment (``w15:done``), a repeating section,
a ``w15:dataBinding``, ``w15:color``, ``w15:appearance`` --- into a document
whose mode is below 15 adds a line to that call's ``ChangeReport.warnings``
naming the feature and the fix (:func:`warn_below`). The mode itself is never
changed for the caller: it is a decision about who can open the document.
"""

from __future__ import annotations

from enum import IntEnum
from typing import Any, Literal

from docx4j_py.model.content.errors import ContentError
from docx4j_py.namespaces import WML_NS

__all__ = [
    "COMPATIBILITY_MODES",
    "COMPAT_URI",
    "DEFAULT_COMPATIBILITY_MODE",
    "NEW_DOCUMENT_MODE",
    "WORD_VERSIONS",
    "CompatibilityMode",
    "CompatibilityModeValue",
    "compat_settings_of",
    "mode_of",
    "set_mode",
    "warn_below",
]


class CompatibilityMode(IntEnum):
    """The Word versions ``w:compatSetting compatibilityMode`` names.

    An ``IntEnum``, so ``pkg.compatibility_mode == 15`` and
    ``pkg.compatibility_mode is CompatibilityMode.WORD_2013`` are both true of a
    Word 2013 document and the value that crosses into JSON is the number Word
    writes.
    """

    WORD_2003 = 11
    WORD_2007 = 12
    WORD_2010 = 14
    #: And Word 2016, 2019 and 365: Word has never written 16.
    WORD_2013 = 15


#: The modes a document may declare, in order.
COMPATIBILITY_MODES: tuple[int, ...] = tuple(int(mode) for mode in CompatibilityMode)

#: What a signature takes: the number, or the enum member beside it.
CompatibilityModeValue = Literal[11, 12, 14, 15]

#: The ``w:uri`` every one of Word's own compatibility settings carries.
COMPAT_URI = "http://schemas.microsoft.com/office/word"

#: The ``w:name`` of the one this module is about.
COMPATIBILITY_SETTING = "compatibilityMode"

#: What Word assumes a document with no setting to be: Word 2007.
DEFAULT_COMPATIBILITY_MODE = int(CompatibilityMode.WORD_2007)

#: What a document this library creates declares: Word 2013 and later. The
#: whole ``w:compat`` that goes with it is the engine's ``NEW_DOCUMENT_COMPAT``
#: (:mod:`docx4j_py.openpackaging.parts.wml`), which owns it because
#: ``create_package`` writes it and the engine imports nothing from here
#: (CR-002 section 12.10).
NEW_DOCUMENT_MODE = int(CompatibilityMode.WORD_2013)

#: What each mode is called, for the warning's message.
WORD_VERSIONS: dict[int, str] = {11: "2003", 12: "2007", 14: "2010", 15: "2013"}

_W = f"{{{WML_NS}}}"


def _w(name: str) -> str:
    return f"{_W}{name}"


# ---------------------------------------------------------------------------
# reading, without unmarshalling
# ---------------------------------------------------------------------------


def mode_of(package: Any) -> int:
    """``pkg.compatibility_mode``: the Word version the document declares.

    :data:`DEFAULT_COMPATIBILITY_MODE` (12, Word 2007) when the document says
    nothing, which is what Word assumes. Read with lxml and cached on the
    package, so ``/word/settings.xml`` is left byte for byte.
    """
    cached = getattr(package, "_compatibility_mode", None)
    if cached is not None:
        return int(cached)
    mode = _read_mode(package)
    try:
        package._compatibility_mode = mode
    except AttributeError:  # pragma: no cover - a package-like object with no slot
        pass
    return mode


def _read_mode(package: Any) -> int:
    """The declared mode as the file has it, through the tree or through lxml."""
    part = getattr(package, "document_settings_part", None)
    if part is None:
        return DEFAULT_COMPATIBILITY_MODE
    if getattr(part, "is_unmarshalled", False):
        setting = _setting_in_tree(part.contents, COMPATIBILITY_SETTING)
        value = getattr(setting, "val", None) if setting is not None else None
        return _as_mode(value)
    from docx4j_py.model.content.describe import _root

    root = _root(part)
    if root is None:
        return DEFAULT_COMPATIBILITY_MODE
    compat = root.find(_w("compat"))
    if compat is None:
        return DEFAULT_COMPATIBILITY_MODE
    for node in compat.findall(_w("compatSetting")):
        if node.get(_w("name")) == COMPATIBILITY_SETTING and node.get(_w("uri")) == COMPAT_URI:
            return _as_mode(node.get(_w("val")))
    return DEFAULT_COMPATIBILITY_MODE


def _as_mode(value: Any) -> int:
    """A ``w:val`` as an int, or the default for one that is not a number."""
    try:
        return int(str(value))
    except (TypeError, ValueError):
        return DEFAULT_COMPATIBILITY_MODE


def compat_settings_of(settings: Any) -> list[Any]:
    """The ``w:compatSetting`` children of a ``w:settings``, or ``[]``."""
    compat = getattr(settings, "compat", None)
    return list(getattr(compat, "compat_setting", ()) or ()) if compat is not None else []


def _setting_in_tree(settings: Any, name: str) -> Any:
    """The ``w:compatSetting`` of Word's own with this name, or None."""
    for setting in compat_settings_of(settings):
        if setting.name == name and setting.uri == COMPAT_URI:
            return setting
    return None


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------


def set_mode(package: Any, value: Any) -> list[str]:
    """Write ``compatibilityMode``; returns the part names touched.

    Unmarshals the settings part, and creates one --- with its relationship and
    its content type --- when the document has none, as
    :func:`~docx4j_py.model.content.tracking.set_mode` does.

    Raises:
        ContentError: `value` is not one of :data:`COMPATIBILITY_MODES`.
    """
    from docx4j_py.model.content.tracking import (
        create_settings_part,
        settings_part_for_write,
    )
    from docx4j_py.wml import CTCompat, CTCompatSetting

    try:
        mode = int(value)
    except (TypeError, ValueError):
        mode = -1
    if mode not in COMPATIBILITY_MODES:
        raise ContentError(
            f"compatibility_mode takes {', '.join(str(m) for m in COMPATIBILITY_MODES)}, "
            f"not {value!r}",
            code="compatibility.mode_invalid",
            hint="pkg.compatibility_mode = 15 targets Word 2013 and later",
        )
    part = settings_part_for_write(package)
    if part is None:
        part = create_settings_part(package)
    settings = part.contents
    if settings.compat is None:
        settings.compat = CTCompat()
        settings.compat.parent = settings
    setting = _setting_in_tree(settings, COMPATIBILITY_SETTING)
    if setting is None:
        setting = CTCompatSetting(name=COMPATIBILITY_SETTING, uri=COMPAT_URI, val=str(mode))
        settings.compat.compat_setting.insert(0, setting)
    else:
        setting.val = str(mode)
    _remember(package, mode)
    return [str(part.part_name)]


def _remember(package: Any, mode: int) -> None:
    """Cache the mode the package now declares."""
    try:
        package._compatibility_mode = mode
    except AttributeError:  # pragma: no cover - a package-like object with no slot
        pass


# ---------------------------------------------------------------------------
# the warning a verb records (CR-003 section 17.11)
# ---------------------------------------------------------------------------


def warn_below(
    package: Any, change: Any, feature: str, *, needed: int = NEW_DOCUMENT_MODE
) -> None:
    """Warn, in this call's report, that `feature` needs a newer compatibility mode.

    Word does not refuse such markup; it reports it in the Compatibility Checker
    when the document is saved, and may drop it (CR-003 section 15.7's resolved
    comment). So the verb writes what it was asked to write and says so, and the
    caller decides whether to raise the mode.
    """
    if change is None or not getattr(change, "active", True):
        return
    mode = _package_mode(package)
    if mode >= needed:
        return
    version = WORD_VERSIONS.get(needed, str(needed))
    change.warn(
        f"{feature} needs compatibility mode {needed} (Word {version} and later); this document "
        f"declares {mode}. Word will report it in the Compatibility Checker. "
        f"Set pkg.compatibility_mode = {needed} to target that version."
    )


def _package_mode(package: Any) -> int:
    """The mode of a package, a trial package, or nothing at all."""
    if package is None:
        return NEW_DOCUMENT_MODE
    found = getattr(package, "compatibility_mode", None)
    return int(found) if found is not None else mode_of(package)
