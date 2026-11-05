"""What a text includes and how it is written: the specification's section 10 options. CR-011 section 3.1.

docx4j's ``org.docx4j.model.text.TextOptions`` and ``TextView`` (its CR-046),
as core-ts CR-012 ports them. The option names are the specification's
(``text_boxes``, ``note_references``, ``field_codes``) verbatim, which in
Python are the attribute names too (CR-011 Q1, core-ts CR-012's D1), so that
the suite's scenarios and an MCP tool's arguments pass through unchanged.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Mapping
from typing import Any, Literal, get_args

from docx4j_py.model.content.enums import TextView
from docx4j_py.model.text.images import ImageDescriber, alt_text_describer
from docx4j_py.openpackaging.exceptions import Docx4JException

__all__ = [
    "NOT_YET",
    "OPTION_NAMES",
    "EmbeddedFilter",
    "TextInclude",
    "TextOptionError",
    "TextOptions",
    "TextView",
    "view_of",
]

#: The stories ``include`` adds (REQ-501, REQ-502).
TextInclude = Literal["headers", "footers", "comments", "glossary"]
_INCLUDES: tuple[str, ...] = get_args(TextInclude)

#: Whether an object is read under ``embedded: true``, given its ``TextObject`` and address (docx4j ``EmbeddedFilter``).
EmbeddedFilter = Callable[[Any, str], bool]

#: The option names of WD section 10, in its order (docx4j ``TextOptions.NAMES``).
OPTION_NAMES: tuple[str, ...] = (
    "view",
    "include",
    "text_boxes",
    "notes",
    "note_references",
    "comments",
    "hidden",
    "field_codes",
    "fields",
    "toc",
    "placeholders",
    "bindings",
    "list_labels",
    "charts",
    "smartart",
    "embedded",
    "altchunks",
    "alt_text",
    "metadata",
    "symbols",
    "math",
    "tables",
    "markup_detail",
    "page_markers",
    "layout",
    "warnings",
    "duplicate_entries",
    "profile",
)

#: What a caller sets by name, recorded so that a profile leaves it (REQ-1002): every option but ``profile`` itself.
_RECORDED = frozenset(OPTION_NAMES) - {"profile"}

#: The boolean options, with their statements.
_BOOLEANS: dict[str, str] = {
    "text_boxes": "REQ-518",
    "notes": "REQ-503",
    "note_references": "REQ-503",
    "comments": "REQ-504",
    "hidden": "REQ-507",
    "field_codes": "REQ-505",
    "toc": "REQ-520",
    "placeholders": "REQ-508",
    "list_labels": "REQ-510",
    "charts": "REQ-511",
    "smartart": "REQ-511",
    "embedded": "REQ-512",
    "altchunks": "REQ-513",
    "alt_text": "REQ-514",
    "metadata": "REQ-515",
    "page_markers": "REQ-708",
    "warnings": "REQ-801",
}

#: The options with named values, their values and their statements.
_CHOICES: dict[str, tuple[tuple[str, ...], str]] = {
    "fields": (("cached", "update"), "REQ-506"),
    "bindings": (("bound", "surface"), "REQ-509"),
    "symbols": (("asis", "unicode"), "REQ-601"),
    "math": (("latex", "unicodemath"), "REQ-604"),
    "tables": (("auto", "tsv", "grid"), "REQ-605"),
    "markup_detail": (("type", "author", "author_date"), "REQ-702"),
    "layout": (("none", "indent"), "REQ-709"),
    "duplicate_entries": (("refuse", "first"), "REQ-206"),
}

#: ``word-text`` (REQ-1002, WD 0.41): Word's Save As plain text as the options reach it, by the specification's
#: names. docx4j's ``TextOptions.WORD_TEXT_VALUES`` at ``0cf4fb38a``, its 24, with ``list_labels`` as docx4j
#: ``87c421731`` sets it (WD 0.84, CR-027 section 4.1); ``math`` is ``latex``, not the WD's ``unicodemath``, which
#: :meth:`TextOptions.check_implemented` refuses. The tag's ``drop_caps: asis`` and ``small_caps: upper`` come with
#: their options in the CR-069 port (CR-029).
_WORD_TEXT: tuple[tuple[str, Any], ...] = (
    ("view", "accepted"),
    ("include", ("headers", "footers", "comments")),
    ("text_boxes", False),
    ("notes", True),
    ("note_references", True),
    ("comments", True),
    ("hidden", True),
    ("field_codes", False),
    ("fields", "cached"),
    ("toc", True),
    ("placeholders", True),
    ("bindings", "bound"),
    ("list_labels", True),
    ("charts", False),
    ("smartart", False),
    ("embedded", False),
    ("altchunks", True),
    ("alt_text", False),
    ("metadata", False),
    ("symbols", "asis"),
    ("math", "latex"),
    ("tables", "tsv"),
    ("page_markers", False),
    ("layout", "none"),
)

#: The profiles by name, as the WD writes them (REQ-1002).
_PROFILES: dict[str, tuple[tuple[str, Any], ...]] = {"word-text": _WORD_TEXT}

#: The values accepted and not yet acted on, with their statements. Empty since
#: CR-011 Phase C (section 13), which acts on each of section 10's values but the
#: two refused (:meth:`TextOptions.check_implemented`); kept, and
#: :meth:`TextOptions.not_yet`, so that a value added to section 10 before it is
#: implemented has a place to be named (REQ-1001: "document any it does not
#: implement").
NOT_YET: dict[tuple[str, Any], str] = {}


class TextOptionError(Docx4JException):
    """An option refused (REQ-1001): an unknown name, a value the specification does not define, or one not implemented.

    Attributes:
        option: the option's name.
        statement: the requirement the option belongs to.
        reason: ``"unknown"``, ``"invalid"`` or ``"not_implemented"``.
    """

    def __init__(self, message: str, option: str, statement: str, reason: str) -> None:
        """The refusal, with what it refused and why."""
        super().__init__(message)
        self.option = option
        self.statement = statement
        self.reason = reason


def view_of(value: Any) -> TextView:
    """A view by name: ``accepted``, ``original`` or ``markup``; ``current`` is ``accepted`` (REQ-401)."""
    name = str(value if value is not None else "").strip().lower()
    if name == "current":
        return "accepted"
    if name in ("accepted", "original", "markup"):
        return name  # type: ignore[return-value]
    raise TextOptionError(
        f"view: {value!r} is not accepted, original, markup or current (WD REQ-401)", "view", "REQ-401", "invalid"
    )


def _bool(name: str, value: Any, statement: str) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value if value is not None else "").strip().lower()
    if text == "true":
        return True
    if text == "false":
        return False
    raise TextOptionError(f"{name}: {value!r} is not true or false (WD REQ-1001)", name, statement, "invalid")


def _choice(name: str, value: Any, values: Iterable[str], statement: str) -> str:
    text = str(value if value is not None else "").strip().lower()
    values = tuple(values)
    if text in values:
        return text
    raise TextOptionError(
        f"{name}: {value!r} is not one of {', '.join(values)} (WD REQ-1001)", name, statement, "invalid"
    )


class TextOptions:
    """What a text includes and how it is written (WD section 10, REQ-1001). docx4j ``TextOptions``.

    Every option name and value of section 10 is accepted, by :meth:`set`,
    :meth:`of` or keyword. The defaults are the specification's: the
    ``accepted`` view, text boxes, notes and their references, placeholders and
    warnings on; comments, hidden text, field codes, the table of contents
    (REQ-520) and list labels off.

    **Refused** (:class:`TextOptionError`, reason ``not_implemented``), as
    docx4j refuses them: ``math: unicodemath`` (no OMML to UnicodeMath
    converter) and ``symbols: unicode`` (no glyph tables).

    ``page_markers: true`` writes ``[page n]`` where Word last broke a page
    (``w:lastRenderedPageBreak``, REQ-708): the markers are current only when
    the package was last paginated after its last edit.

    ``profile`` (REQ-1002) names a preset of option values, applied before the
    caller's own: an option the caller sets, before or after the profile, by a
    keyword, :meth:`set`, :meth:`of`, a mapping in any order, :meth:`include` or
    an attribute, stands, and the preset fills the rest. ``word-text`` (matched
    without case, with a hyphen or an underscore) is Word's Save As plain text as
    the options reach it: the ``accepted`` view, hidden text in, headers, footers
    and comments in, text boxes out, notes and their references in, list labels
    in, fields cached, no field codes, the table of contents in, placeholders in,
    bound values, chunks expanded, symbols as they are, tables as tab-separated
    rows, no page markers, no indentation, and no alt text, metadata, charts,
    SmartArt or embedded packages. It is the nearest the options reach, not a
    mimic. **What it does not reproduce** (REQ-1002's NOTE): the notes' and
    comments' forms (Word writes the notes after the body, each led by its mark,
    and marks a comment's anchor with the author's initials); CRLF lines and CR
    cells, in the saving system's ANSI code page by default, where the text here
    has U+000A, tabs and Unicode; the one space Word's text save writes after
    every list label, where the text here has the level's ``w:suff``, as Word
    draws it; folded hyphens, U+2011 and U+00AD kept here; and Word's linear math
    (``math: unicodemath`` is not implemented, so the preset leaves ``math`` at
    ``latex``, where the WD's value is ``unicodemath``). Three accidents of Word's
    own are left out on purpose: the partial update Word's open runs where a DATE,
    TIME or SAVEDATE field is present, the SYMBOL omission, and the stale PAGE
    number. The WD's ``drop_caps: asis`` and ``small_caps: upper`` come with their
    options (CR-029): until then ``w:smallCaps`` text is as stored.

    ``duplicate_entries`` (``refuse``, ``first``; REQ-206) is the **load's**
    option, accepted here as REQ-1001 requires (CR-032 Q5): a reading is of a
    package already loaded, so it does not decide again. Pass it to the load,
    ``LoadOptions(duplicate_entries=...)``; the load's setting holds, and
    under ``first`` the result's ``package.duplicate_entry`` warnings say what
    was read.

    Beyond section 10, ``image_describer`` (REQ-514): what gives an image its
    description, the author's alt text by default; and ``field_user``
    (REQ-506, CR-025 D10): the user ``USERNAME``, ``USERINITIALS`` and
    ``USERADDRESS`` show under ``fields: update``, a
    :class:`~docx4j_py.model.fields.FieldUser`, a name, or a mapping of
    ``name``, ``initials`` and ``address``; None by default, and then the three
    keep their cached result, as docx4j's extraction keeps them; and
    ``embedded_filter`` (REQ-522): under ``embedded: true``, a callable
    ``(obj, address) -> bool`` asked once per object that would be read, before
    the reading, with the object's
    :class:`~docx4j_py.model.text.result.TextObject` and the address its content
    is read at (``body/3/object:0``), every object counted, declined or not. A
    workbook, a presentation and an OLE 1.0 package are asked too, though this
    package does not read them (CR-011 Q7), so that a host can take them. An
    object declined contributes nothing and keeps its run and its description; a
    filter that raises declines, with ``content.not_read`` naming the object's
    part at the paragraph's address. None by default: every object is read.
    """

    __slots__ = (
        "_comments",
        "_explicit",
        "_include",
        "_profile",
        "alt_text",
        "altchunks",
        "bindings",
        "charts",
        "duplicate_entries",
        "embedded",
        "embedded_filter",
        "field_codes",
        "field_user",
        "fields",
        "hidden",
        "image_describer",
        "layout",
        "list_labels",
        "markup_detail",
        "math",
        "metadata",
        "note_references",
        "notes",
        "page_markers",
        "placeholders",
        "smartart",
        "symbols",
        "tables",
        "text_boxes",
        "toc",
        "view",
        "warnings",
    )

    def __init__(self, **options: Any) -> None:
        """The defaults, then each option given by its specification name."""
        #: The options a caller has set, by every route: a profile leaves them (REQ-1002).
        object.__setattr__(self, "_explicit", set())
        self._profile: str | None = None
        self.view: TextView = "accepted"
        self._include: list[str] = []
        self.text_boxes = True
        self.notes = True
        self.note_references = True
        self._comments = False
        self.hidden = False
        self.field_codes = False
        self.fields = "cached"
        #: Whether a table of contents is read: false by default (REQ-520,
        #: docx4j CR-046 D12). Under ``toc: false`` nothing of a TOC field is
        #: read, its code included under ``field_codes``, and a paragraph with
        #: nothing else in it is no paragraph.
        self.toc = False
        self.placeholders = True
        self.bindings = "bound"
        self.list_labels = False
        self.charts = True
        self.smartart = True
        self.embedded = False
        self.altchunks = True
        self.alt_text = False
        self.metadata = False
        self.symbols = "asis"
        self.math = "latex"
        self.tables = "auto"
        self.markup_detail = "type"
        self.page_markers = False
        self.layout = "none"
        self.warnings = True
        #: The load's option (REQ-206, CR-032 Q5): accepted, and read by whoever
        #: loads the package (``LoadOptions(duplicate_entries=...)``); a reading
        #: of a package loaded already does not act on it.
        self.duplicate_entries = "refuse"
        #: What gives an image its description (REQ-514): beyond section 10, a
        #: host's hook, as docx4j's ``TextOptions.imageDescriber``. The
        #: author's alt text by default (:func:`~docx4j_py.model.text.images.alt_text_describer`).
        self.image_describer: ImageDescriber = alt_text_describer
        #: The user the user fields show under ``fields: update`` (REQ-506):
        #: beyond section 10, a host's, as docx4j's ``FieldContext.user`` is the
        #: updater's. None by default: the fields keep their cached result.
        self.field_user: Any = None
        #: Whether an object is read under ``embedded: true`` (REQ-522): beyond
        #: section 10, a host's hook, as docx4j's ``TextOptions.embeddedFilter``.
        #: None by default: every object is read.
        self.embedded_filter: EmbeddedFilter | None = None
        self._explicit.clear()  # a default is nobody's: past the recording, or a profile would fill nothing
        for name, value in options.items():
            self.set(name, value)

    def __setattr__(self, name: str, value: Any) -> None:
        """An attribute, recorded as the caller's where it is an option (CR-027 Q5)."""
        object.__setattr__(self, name, value)
        if name in _RECORDED:
            self._explicit.add(name)

    @classmethod
    def of(cls, options: TextOptions | Mapping[str, Any] | None = None, **more: Any) -> TextOptions:
        """Options from the specification's names and values (REQ-1001); those not named keep their defaults."""
        if isinstance(options, TextOptions):
            if not more:
                return options
            out = options.copy()
        else:
            out = cls()
            for name, value in (options or {}).items():
                if value is not None:
                    out.set(name, value)
        for name, value in more.items():
            if value is not None:
                out.set(name, value)
        return out

    def set(self, name: str, value: Any) -> TextOptions:
        """Set one option by its specification name; raises :class:`TextOptionError` for one section 10 lacks."""
        self._explicit.add(self._apply(name, value))
        return self

    def _apply(self, name: str, value: Any) -> str:
        """Set one option without recording it as the caller's, as a profile's value is set; gives the name as read."""
        key = (name or "").strip()
        assign = object.__setattr__
        if key == "profile":
            self._apply_profile(value)
        elif key == "view":
            assign(self, "view", view_of(value))
        elif key == "include":
            self._set_include(value)
        elif key == "comments":
            assign(self, "_comments", _bool(key, value, "REQ-504"))
        elif key in _BOOLEANS:
            assign(self, key, _bool(key, value, _BOOLEANS[key]))
        elif key in _CHOICES:
            values, statement = _CHOICES[key]
            assign(self, key, _choice(key, value, values, statement))
        elif key == "image_describer":
            if value is not None and not callable(value):
                raise TextOptionError(
                    f"image_describer: {value!r} is not callable (WD REQ-514)", key, "REQ-514", "invalid"
                )
            assign(self, key, value or alt_text_describer)
        elif key == "embedded_filter":
            if value is not None and not callable(value):
                raise TextOptionError(
                    f"embedded_filter: {value!r} is not callable (WD REQ-522)", key, "REQ-522", "invalid"
                )
            assign(self, key, value)
        elif key == "field_user":
            from docx4j_py.model.fields.context import FieldUser

            try:
                assign(self, key, FieldUser.of(value))
            except TypeError as error:
                raise TextOptionError(f"field_user: {error} (WD REQ-506)", key, "REQ-506", "invalid") from error
        else:
            raise TextOptionError(
                f"Unknown text option {name!r}; the options are {', '.join(OPTION_NAMES)} (WD REQ-1001)",
                str(name),
                "REQ-1001",
                "unknown",
            )
        return key

    def _apply_profile(self, value: Any) -> None:
        """A profile by name: its values for every option the caller has not set (REQ-1002). None takes nothing back."""
        name = None if value is None else str(value).strip().lower().replace("_", "-")
        if name is not None and name not in _PROFILES:
            raise TextOptionError(
                f"profile: {value!r} is not one of {', '.join(_PROFILES)} (WD REQ-1002)",
                "profile",
                "REQ-1002",
                "invalid",
            )
        object.__setattr__(self, "_profile", name)
        for option, preset in _PROFILES.get(name, ()):
            if option not in self._explicit:
                self._apply(option, preset)

    @property
    def profile(self) -> str | None:
        """The profile applied (``word-text``), or None. Assigning one applies it, as :meth:`set` does."""
        return self._profile

    @profile.setter
    def profile(self, value: Any) -> None:
        self._apply_profile(value)

    def _set_include(self, value: Any) -> None:
        if value is None:
            items: list[Any] = []
        elif isinstance(value, str):
            items = [item for item in re.split(r"[,\s]+", value) if item]
        else:
            items = list(value)
        self._include = list(dict.fromkeys(_choice("include", item, _INCLUDES, "REQ-501") for item in items))

    def include(self, *stories: str) -> TextOptions:
        """Add stories to ``include`` (docx4j's builder ``include(Include...)``)."""
        for story in stories:
            name = _choice("include", story, _INCLUDES, "REQ-501")
            if name not in self._include:
                self._include.append(name)
        self._explicit.add("include")
        return self

    @property
    def includes(self) -> frozenset[str]:
        """The stories ``include`` names; ``comments: true`` adds ``comments``."""
        out = set(self._include)
        if self._comments:
            out.add("comments")
        return frozenset(out)

    @property
    def comments(self) -> bool:
        """``comments: true`` or ``include: comments``."""
        return self._comments or "comments" in self._include

    @comments.setter
    def comments(self, value: bool) -> None:
        self._comments = bool(value)

    def copy(self) -> TextOptions:
        """A copy, with what the caller has set and the profile as they stand."""
        out = TextOptions()
        for name in self.__slots__:
            value = getattr(self, name)
            object.__setattr__(out, name, type(value)(value) if isinstance(value, (list, set)) else value)
        return out

    def check_implemented(self) -> None:
        """Raise for a value this package refuses (REQ-1001: documented, and refused rather than ignored)."""
        if self.math == "unicodemath":
            raise TextOptionError(
                "Not implemented (WD REQ-1001): math: unicodemath (no OMML to UnicodeMath converter)",
                "math",
                "REQ-604",
                "not_implemented",
            )
        if self.symbols == "unicode":
            raise TextOptionError(
                "Not implemented (WD REQ-1001): symbols: unicode (no glyph tables for Symbol and Wingdings)",
                "symbols",
                "REQ-601",
                "not_implemented",
            )

    def not_yet(self) -> list[str]:
        """The options in effect that are accepted and not yet acted on (:data:`NOT_YET`), as ``name: value``."""
        out = []
        for name in OPTION_NAMES:
            if name in ("view", "include"):
                continue
            value = self.comments if name == "comments" else getattr(self, name)
            if (name, value) in NOT_YET:
                shown = str(value).lower() if isinstance(value, bool) else value
                out.append(f"{name}: {shown}")
        return out

    def to_dict(self) -> dict[str, Any]:
        """The options by the specification's names, in section 10's order: ``profile`` last, None for none."""
        out: dict[str, Any] = {}
        for name in OPTION_NAMES:
            if name == "include":
                out[name] = list(self._include)
            elif name == "comments":
                out[name] = self._comments
            else:
                out[name] = getattr(self, name)
        return out

    def __repr__(self) -> str:
        """The options that differ from the defaults."""
        defaults = TextOptions().to_dict()
        changed = {k: v for k, v in self.to_dict().items() if defaults[k] != v}
        return f"TextOptions({', '.join(f'{k}={v!r}' for k, v in changed.items())})"
