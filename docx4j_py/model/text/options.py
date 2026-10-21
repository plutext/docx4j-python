"""What a text includes and how it is written: the specification's section 10 options. CR-011 section 3.1.

docx4j's ``org.docx4j.model.text.TextOptions`` and ``TextView`` (its CR-046),
as core-ts CR-012 ports them. The option names are the specification's
(``text_boxes``, ``note_references``, ``field_codes``) verbatim, which in
Python are the attribute names too (CR-011 Q1, core-ts CR-012's D1), so that
the suite's scenarios and an MCP tool's arguments pass through unchanged.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from typing import Any, Literal, get_args

from docx4j_py.model.content.enums import TextView
from docx4j_py.openpackaging.exceptions import Docx4JException

__all__ = [
    "NOT_YET",
    "OPTION_NAMES",
    "TextInclude",
    "TextOptionError",
    "TextOptions",
    "TextView",
    "view_of",
]

#: The stories ``include`` adds (REQ-501, REQ-502).
TextInclude = Literal["headers", "footers", "comments", "glossary"]
_INCLUDES: tuple[str, ...] = get_args(TextInclude)

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
)

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
}

#: The values accepted and not yet acted on, with their statements: CR-011
#: Phase C's content (section 6). Each is read as its option's default until
#: then, documented here as REQ-1001 asks rather than refused, since a refusal
#: would fail checks of the suite that the default's text meets.
NOT_YET: dict[tuple[str, Any], str] = {
    ("fields", "update"): "REQ-506",
    ("bindings", "bound"): "REQ-509",
    ("charts", True): "REQ-511",
    ("smartart", True): "REQ-511",
    ("embedded", True): "REQ-512",
    ("altchunks", True): "REQ-513",
    ("alt_text", True): "REQ-514",
    ("metadata", True): "REQ-515",
    ("tables", "grid"): "REQ-605",
    ("page_markers", True): "REQ-708",
    ("layout", "indent"): "REQ-709",
}


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

    **Not yet** (CR-011 Phase C; :data:`NOT_YET`): ``fields: update``,
    ``embedded: true``, ``alt_text: true``, ``metadata: true``, ``tables:
    grid``, ``page_markers: true`` and ``layout: indent`` are accepted and read
    as their defaults; under the default ``charts``, ``smartart``,
    ``altchunks`` and ``bindings`` a chart or SmartArt contributes nothing, an
    ``altChunk`` is not expanded and a bound control is read by its surface.
    In ``tables: auto`` every table is written in REQ-605's simple form.
    """

    __slots__ = (
        "_comments",
        "_include",
        "alt_text",
        "altchunks",
        "bindings",
        "charts",
        "embedded",
        "field_codes",
        "fields",
        "hidden",
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
        for name, value in options.items():
            self.set(name, value)

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
        key = (name or "").strip()
        if key == "view":
            self.view = view_of(value)
        elif key == "include":
            self._set_include(value)
        elif key == "comments":
            self._comments = _bool(key, value, "REQ-504")
        elif key in _BOOLEANS:
            setattr(self, key, _bool(key, value, _BOOLEANS[key]))
        elif key in _CHOICES:
            values, statement = _CHOICES[key]
            setattr(self, key, _choice(key, value, values, statement))
        else:
            raise TextOptionError(
                f"Unknown text option {name!r}; the options are {', '.join(OPTION_NAMES)} (WD REQ-1001)",
                str(name),
                "REQ-1001",
                "unknown",
            )
        return self

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
        """A copy."""
        out = TextOptions()
        for name in self.__slots__:
            value = getattr(self, name)
            setattr(out, name, list(value) if isinstance(value, list) else value)
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
        """The options by the specification's names, in section 10's order."""
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
