"""``Font``: Office JS ``Word.Font`` over the direct run properties of a scope.

CR-003 section 3.2. A ``Font`` is not a thing in the tree: it is a view over the
``w:rPr`` of however many runs its scope covers --- every run of a paragraph, or
just the runs a :class:`~docx4j_py.model.content.Range` has split out.

**Reads report the first run in scope's effective formatting**, as Office JS
does: the document defaults, the paragraph style's run properties, the run's
character style and then its direct formatting, resolved by CR-002 Phase B's
:class:`~docx4j_py.model.property_resolver.PropertyResolver` (CR-002 section
15.1; CR-003 section 6, which this closes). ``paragraph.get_font(direct=True)``
and ``range.get_font(direct=True)`` read the run's own ``w:rPr`` instead, which
is what a caller reproducing markup wants and what these reads were until Phase
B; a font over something with no package behind it reads direct, because there
is no hierarchy to resolve against.

**A theme colour** reads as Word draws it in the effective read: the
``w:themeColor`` (with its tint or shade) through the document's theme, not the
``w:val`` cache beside it (CR-009 Phase B, :attr:`Font.color`); the direct read
keeps ``w:val``, and :meth:`Font.to_dict` carries the theme attributes as
written.

The **python-docx facade** (:mod:`docx4j_py.docx`) keeps python-docx's own
tri-state semantics --- ``None`` means *inherited*, and python-docx does not
report effective values either --- and is deliberately not switched (CR-003
section 20.3).

**Writes apply to every run in scope.** Both directions go through CR-001's one
name-to-``w:rPr`` mapping, :func:`~docx4j_py.wml.builders.apply_run_options` and
:func:`~docx4j_py.wml.builders.read_run_options`, which the ``r(text, **opts)``
builder shares --- so a property a builder sets is the property this reads back,
and the ``w:bCs`` / ``w:iCs`` / ``w:szCs`` complex-script twins of section 3.12
are written by that one mapping rather than here.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docx4j_py.model.content.enums import UnderlineValue
from docx4j_py.wml import RPr, apply_run_options, read_run_options

__all__ = ["Font", "MixedFont"]

#: ``Font.to_dict()``'s keys for ``w:color``'s theme attributes, as written (CR-009 Phase B).
THEME_KEYS = ("theme_color", "theme_tint", "theme_shade")


def _theme_attributes(color: Any) -> dict[str, str | None]:
    """``w:color``'s ``w:themeColor``, ``w:themeTint`` and ``w:themeShade`` as written; None where absent.

    An ``ST_ThemeColor`` member reads as its value (``"accent1"``), and a value
    the exact parse kept as its text (CR-018) as that text.
    """
    if color is None:
        return dict.fromkeys(THEME_KEYS)
    reference = color.theme_color
    return {
        "theme_color": None if reference is None else str(getattr(reference, "value", reference)),
        "theme_tint": color.theme_tint,
        "theme_shade": color.theme_shade,
    }


def _drawn(resolver: Any, color: Any) -> str | None:
    """``"#RRGGBB"``, the colour Word draws, where an effective ``w:color`` has a theme reference; else None.

    CR-009 Phase B (question 2), core-ts's ``readingOf`` at ``76fb45d``: the
    ``w:themeColor`` with its tint or shade, through the document's theme and
    ``w:clrSchemeMapping``, over the ``w:val`` cache, Word's own value kept
    where it is within one unit
    (:meth:`~docx4j_py.model.property_resolver.PropertyResolver.resolve_color`).
    None where there is no reference (or it is ``none``), and where the
    reference resolves to automatic: the read of ``w:val`` stands then, so an
    automatic or unstated colour reads as it did (``"auto"``, ``""``; CR-018
    rule 3).
    """
    if color is None:
        return None
    reference = color.theme_color
    if reference is None or getattr(reference, "value", reference) == "none":
        return None
    resolved = resolver.resolve_color(color)
    return None if resolved is None else f"#{resolved.hex()}"


class Font:
    """A subset of Office JS ``Word.Font``, over the ``w:rPr`` of a scope."""

    __slots__ = ("_context", "_direct", "_guard", "_holders", "_record", "_scope", "_tracking")

    def __init__(
        self,
        holders: Callable[[], list[Any]],
        *,
        scope: str = "",
        record: Any = None,
        tracking: Callable[[], Any] | None = None,
        context: Callable[[], tuple[Any, Any]] | None = None,
        direct: bool = False,
        guard: Callable[[], None] | None = None,
    ) -> None:
        """Build the view.

        Args:
            holders: called on every read and every write; returns the objects
                whose ``r_pr`` this font is over, in document order. A callable
                rather than a list because a range splits its runs at the
                boundaries the moment a write asks for them.
            scope: what this font is over, for :meth:`__repr__`.
            record: the ``paragraph.formatting(name, read)`` context manager, so
                that a write records a ``"format"``
                :class:`~docx4j_py.model.content.reports.ChangeReport` with the
                values either side (CR-003 section 3.4). None for a font over
                something with no paragraph behind it.
            tracking: called on every write; returns the
                :class:`~docx4j_py.model.content.tracking.FontTracking` of a
                package whose ``change_tracking_mode`` is on, or None. When it
                returns one, each run's properties as they stand are recorded in
                ``w:rPrChange`` before the write (CR-003 section 3.8).
            context: called on every read; returns
                ``(property_resolver, w:pPr)`` --- the resolver of the package
                this font's runs are in and the ``w:pPr`` of the paragraph they
                are in --- or ``(None, None)`` where there is no package to
                resolve against. None for a font that never resolves.
            direct: read the run's own ``w:rPr`` and not the effective one.
            guard: called before every write, before anything is split or
                recorded; it raises to refuse the write. A range's and a
                paragraph's refuse one over text read from a kept
                ``mc:AlternateContent`` branch (CR-011 Phase F).
        """
        self._holders = holders
        self._scope = scope
        self._record = record
        self._tracking = tracking
        self._context = context
        self._direct = direct
        self._guard = guard

    # -- the mapping -------------------------------------------------------

    def _direct_r_pr(self) -> Any:
        holders = self._holders()
        return getattr(holders[0], "r_pr", None) if holders else None

    def _read(self) -> dict[str, Any]:
        r_pr = self._direct_r_pr()
        if not self._direct and self._context is not None:
            resolver, p_pr = self._context()
            if resolver is not None:
                effective = resolver.get_effective_rpr(r_pr, p_pr)
                out = read_run_options(effective)
                # The reads the effective w:rPr cannot answer on its own: a run
                # whose face comes from the theme has a w:asciiTheme and no
                # w:ascii, so the name is the selector's (CR-002 section 15.3);
                # and a theme colour is drawn in the theme's colour, not the
                # w:val beside it (CR-009 Phase B).
                selector = self._selector()
                if selector is not None:
                    out["name"] = selector.ascii_font_name(effective)
                color = getattr(effective, "color", None)
                drawn = _drawn(resolver, color)
                if drawn is not None:
                    out["color"] = drawn
                out.update(_theme_attributes(color))
                return out
        out = read_run_options(r_pr)
        out.update(_theme_attributes(getattr(r_pr, "color", None)))
        return out

    def _selector(self) -> Any:
        """The package's run font selector, where one has been built."""
        if self._context is None:
            return None
        resolver, _p_pr = self._context()
        package = getattr(resolver, "package", None)
        if package is None:
            return None
        try:
            return package.get_run_font_selector()
        except Exception:  # noqa: BLE001 - a name is not worth an exception
            return None

    def _read_direct(self) -> dict[str, Any]:
        r_pr = self._direct_r_pr()
        out = read_run_options(r_pr)
        out.update(_theme_attributes(getattr(r_pr, "color", None)))
        return out

    def _apply(self, **options: Any) -> None:
        if self._guard is not None:
            self._guard()
        if self._record is None:
            self._write(**options)
            return
        name = next(iter(options))
        # the change report is about what the *markup* says, either side of the
        # write, so it reads direct formatting even where this font resolves
        with self._record(name, lambda: self._read_direct().get(name)):
            self._write(**options)

    def _write(self, **options: Any) -> None:
        tracking = self._tracking() if self._tracking is not None else None
        for holder in self._holders():
            r_pr = getattr(holder, "r_pr", None)
            if r_pr is None:
                r_pr = RPr()
                holder.r_pr = r_pr
                r_pr.parent = holder
            # the properties as they are now become w:rPrChange, unless this run
            # is this author's own insertion (CR-003 section 3.8)
            if tracking is not None and tracking.records(holder):
                tracking.tracker.record_r_pr_change(r_pr)
            apply_run_options(r_pr, **options)

    # -- Office JS's properties -------------------------------------------

    @property
    def bold(self) -> bool:
        """``w:b`` (and ``w:bCs``). Office JS ``Word.Font.bold``."""
        return bool(self._read()["bold"])

    @bold.setter
    def bold(self, value: bool) -> None:
        self._apply(bold=bool(value))

    @property
    def italic(self) -> bool:
        """``w:i`` (and ``w:iCs``). Office JS ``Word.Font.italic``."""
        return bool(self._read()["italic"])

    @italic.setter
    def italic(self, value: bool) -> None:
        self._apply(italic=bool(value))

    @property
    def strike_through(self) -> bool:
        """``w:strike``. Office JS ``Word.Font.strikeThrough``."""
        return bool(self._read()["strike_through"])

    @strike_through.setter
    def strike_through(self, value: bool) -> None:
        self._apply(strike_through=bool(value))

    @property
    def double_strike_through(self) -> bool:
        """``w:dstrike``. Office JS ``Word.Font.doubleStrikeThrough``."""
        return bool(self._read()["double_strike_through"])

    @double_strike_through.setter
    def double_strike_through(self, value: bool) -> None:
        self._apply(double_strike_through=bool(value))

    @property
    def subscript(self) -> bool:
        """``w:vertAlign w:val="subscript"``. Office JS ``Word.Font.subscript``."""
        return bool(self._read()["subscript"])

    @subscript.setter
    def subscript(self, value: bool) -> None:
        self._apply(subscript=bool(value))

    @property
    def superscript(self) -> bool:
        """``w:vertAlign w:val="superscript"``. Office JS ``Word.Font.superscript``."""
        return bool(self._read()["superscript"])

    @superscript.setter
    def superscript(self, value: bool) -> None:
        self._apply(superscript=bool(value))

    @property
    def underline(self) -> UnderlineValue:
        """``w:u``, as an Office JS ``Word.UnderlineType`` name.

        ``"None"`` when there is none. The value as Word reads it (CR-018 Phase
        B): one outside the enumeration (``SINGLE``, ``bogus``, ``""``) is no
        underline of the run's own, so ``"None"`` here and the style's beneath
        in the effective read, where until 0.2.0 it read ``"Mixed"``; ``w:u``'s
        six on/off words are ``none``. ``"Mixed"`` is a range whose runs differ
        (:class:`MixedFont`). Setting ``"None"`` removes the element.
        """
        return self._read()["underline"]  # type: ignore[return-value]

    @underline.setter
    def underline(self, value: UnderlineValue | bool) -> None:
        self._apply(underline=value)

    @property
    def name(self) -> str:
        """The face the run's ASCII text is set in.

        ``w:rFonts/@w:ascii``, or the face ``w:asciiTheme`` names, else
        ``w:hAnsi``, else the document default --- the effective read, through
        :meth:`~docx4j_py.model.fonts.RunFontSelector.ascii_font_name` (CR-002
        section 15.3). ``get_font(direct=True).name`` is the run's own
        ``w:ascii``, ``""`` where it sets none.

        Setting writes ``w:ascii`` and ``w:hAnsi`` together and clears the theme
        fonts, as Word does; ``""`` removes ``w:rFonts``.
        """
        return self._read()["name"]

    @name.setter
    def name(self, value: str) -> None:
        self._apply(name=value)

    @property
    def size(self) -> float:
        """The size in points (``w:sz`` is half-points); ``0`` when unset.

        Setting ``0`` removes ``w:sz`` and ``w:szCs``, which is what writing
        back what a silent run reads means.
        """
        return self._read()["size"]

    @size.setter
    def size(self, value: float) -> None:
        self._apply(size=value)

    @property
    def color(self) -> str:
        """``w:color`` as ``"#RRGGBB"``, which is what Office JS reports.

        **A theme colour reads as Word draws it** in the effective read
        (CR-009 Phase B, question 2): where the effective ``w:color`` has a
        ``w:themeColor``, the theme's colour, its ``w:themeTint`` or
        ``w:themeShade`` applied (the tint where both are), through the
        document's theme and ``w:clrSchemeMapping``, and not the ``w:val``
        beside it, which is a cache Word keeps and can be stale; Word's own
        value where it is within one unit (:class:`~docx4j_py.model.color.WordColors`).
        So a Heading 1 written ``365F91`` under ``accent1`` shaded ``BF`` reads
        ``#0F4761`` in the Office 2023 theme, as Word 365 reads it. The direct
        read (``get_font(direct=True)``) keeps ``w:val`` as written, and
        :meth:`to_dict` gives the reference itself (``theme_color``,
        ``theme_tint``, ``theme_shade``).

        Otherwise as Word reads ``w:val`` (CR-018 Phase B, rule 3): six hex
        digits in any case, with a leading ``#`` or surrounding spaces, or a
        classic HTML colour name (``red``, ``DarkRed``), as ``"#RRGGBB"``;
        exactly ``auto`` as ``"auto"``; ``""`` when the run sets none, and for a
        value Word rejects (``Auto``, ``F00``, ``grey``), which it reads as
        absent --- the effective read then gives the style's colour beneath.
        Setting accepts ``"#FF0000"``, ``"FF0000"`` and a name, which is
        written as its ``RRGGBB`` with no theme reference.
        """
        return self._read()["color"]

    @color.setter
    def color(self, value: str) -> None:
        self._apply(color=value)

    @property
    def highlight_color(self) -> str | None:
        """``w:highlight`` as ``"#RRGGBB"``, or None when there is none.

        As Word reads it (CR-018 Phase B): a value outside the enumeration
        (``YELLOW``, ``bogus``) is None, as is ``none``; and the effective read
        never takes a character style's highlight, which Word does not draw
        (round six), where a paragraph style's and the run's own are drawn.
        Setting takes a highlight name (``"yellow"``) or one of its ``#RRGGBB``
        values; None removes it.
        """
        return self._read()["highlight_color"]

    @highlight_color.setter
    def highlight_color(self, value: str | None) -> None:
        self._apply(highlight_color=value)

    # -- extensions --------------------------------------------------------

    @property
    def style(self) -> str:
        """The character style id (``w:rStyle``); ``""`` when none. Extension."""
        return self._read()["style"]

    @style.setter
    def style(self, value: str) -> None:
        self._apply(style=value)

    def to_dict(self) -> dict[str, Any]:
        """Every property this font reports, JSON-ready. Extension.

        With ``theme_color``, ``theme_tint`` and ``theme_shade``, the
        ``w:color``'s theme attributes as written (effective or direct as this
        font reads; None where absent), so a reader sees both the reference and
        the colour it draws (CR-009 Phase B).
        """
        return dict(self._read())

    def __repr__(self) -> str:
        """``<Font bold italic 12.0pt #FF0000>``, or ``<Font plain>``."""
        read = self._read()
        parts = [name for name in ("bold", "italic", "strike_through") if read[name]]
        if read["underline"] != "None":
            parts.append(f"underline={read['underline']}")
        if read["name"]:
            parts.append(read["name"])
        if read["size"]:
            parts.append(f"{read['size']}pt")
        if read["color"]:
            parts.append(read["color"])
        inside = " ".join(parts) or "plain"
        scope = f" {self._scope}" if self._scope else ""
        return f"<Font{scope} {inside}>"


class MixedFont(Font):
    """The :class:`Font` of a range across paragraphs: each run read in its own paragraph. Extension.

    CR-013 Phase B (core-ts ``MixedFont``, ``Font.mts`` at ``ef57729``; Q4): a
    property where the runs differ reads as Office JS reads it (core-ts checks
    52 and 53) --- ``None``, ``""`` for the name, the colour and the
    highlight, and ``"Mixed"`` for the underline --- so the getters are typed
    ``bool | None`` and ``float | None``. Where they agree, that value. A
    write (CR-013 Phase C) goes to `write`: every run and mark in the range,
    as Word formats them (:func:`~docx4j_py.model.content.span_edits.format_span`).
    """

    __slots__ = ("_readings", "_write_across")

    def __init__(
        self, readings: Callable[[], list[dict[str, Any]]], write: Callable[[dict[str, Any]], None]
    ) -> None:
        """Build the view.

        Args:
            readings: called on every read; each run's
                :func:`~docx4j_py.wml.builders.read_run_options` dict, effective
                or direct as the range asked, read in its own paragraph.
            write: called by every setter with the one option written
                (``{"bold": True}``).
        """
        super().__init__(list)
        self._readings = readings
        self._write_across = write

    def _value(self, key: str, mixed: Any = None) -> Any:
        values = [reading.get(key) for reading in self._readings()]
        if not values:
            return None
        first = values[0]
        return first if all(value == first for value in values) else mixed

    def _flag(self, key: str) -> bool | None:
        values = [bool(reading.get(key)) for reading in self._readings()]
        if not values:
            return None
        return values[0] if all(value == values[0] for value in values) else None

    def _read(self) -> dict[str, Any]:
        return {
            "bold": self._flag("bold"),
            "italic": self._flag("italic"),
            "strike_through": self._flag("strike_through"),
            "double_strike_through": self._flag("double_strike_through"),
            "subscript": self._flag("subscript"),
            "superscript": self._flag("superscript"),
            "underline": self._value("underline", "Mixed") or "Mixed",
            "name": self._value("name", "") or "",
            "size": self._value("size"),
            "color": self._value("color", "") or "",
            "highlight_color": self._highlight(),
            "style": self._value("style", "") or "",
            # w:color's theme attributes where every run agrees, else None (CR-009 Phase B)
            **{key: self._value(key) for key in THEME_KEYS},
        }

    def _highlight(self) -> str | None:
        values = [reading.get("highlight_color") for reading in self._readings()]
        if not values:
            return None
        return values[0] if all(value == values[0] for value in values) else ""

    def _apply(self, **options: Any) -> None:
        self._write_across(options)

    # the getters, typed for a value that may be mixed (Q4)

    @property
    def bold(self) -> bool | None:  # type: ignore[override]
        """``w:b`` where every run agrees, else None."""
        return self._read()["bold"]

    @bold.setter
    def bold(self, value: bool) -> None:
        self._apply(bold=value)

    @property
    def italic(self) -> bool | None:  # type: ignore[override]
        """``w:i`` where every run agrees, else None."""
        return self._read()["italic"]

    @italic.setter
    def italic(self, value: bool) -> None:
        self._apply(italic=value)

    @property
    def strike_through(self) -> bool | None:  # type: ignore[override]
        """``w:strike`` where every run agrees, else None."""
        return self._read()["strike_through"]

    @strike_through.setter
    def strike_through(self, value: bool) -> None:
        self._apply(strike_through=value)

    @property
    def double_strike_through(self) -> bool | None:  # type: ignore[override]
        """``w:dstrike`` where every run agrees, else None."""
        return self._read()["double_strike_through"]

    @double_strike_through.setter
    def double_strike_through(self, value: bool) -> None:
        self._apply(double_strike_through=value)

    @property
    def subscript(self) -> bool | None:  # type: ignore[override]
        """Subscript where every run agrees, else None."""
        return self._read()["subscript"]

    @subscript.setter
    def subscript(self, value: bool) -> None:
        self._apply(subscript=value)

    @property
    def superscript(self) -> bool | None:  # type: ignore[override]
        """Superscript where every run agrees, else None."""
        return self._read()["superscript"]

    @superscript.setter
    def superscript(self, value: bool) -> None:
        self._apply(superscript=value)

    @property
    def underline(self) -> UnderlineValue:
        """The underline where every run agrees, else ``"Mixed"``."""
        return self._read()["underline"]  # type: ignore[no-any-return]

    @underline.setter
    def underline(self, value: UnderlineValue | bool) -> None:
        self._apply(underline=value)

    @property
    def name(self) -> str:
        """The face where every run agrees, else ``""``."""
        return self._read()["name"]  # type: ignore[no-any-return]

    @name.setter
    def name(self, value: str) -> None:
        self._apply(name=value)

    @property
    def size(self) -> float | None:  # type: ignore[override]
        """The size in points where every run agrees, else None."""
        return self._read()["size"]  # type: ignore[no-any-return]

    @size.setter
    def size(self, value: float) -> None:
        self._apply(size=value)

    @property
    def color(self) -> str:
        """The colour where every run agrees, else ``""``."""
        return self._read()["color"]  # type: ignore[no-any-return]

    @color.setter
    def color(self, value: str) -> None:
        self._apply(color=value)

    @property
    def highlight_color(self) -> str | None:
        """The highlight where every run agrees (None for none), else ``""``."""
        return self._read()["highlight_color"]  # type: ignore[no-any-return]

    @highlight_color.setter
    def highlight_color(self, value: str | None) -> None:
        self._apply(highlight_color=value)

    @property
    def style(self) -> str:
        """The character style id where every run agrees, else ``""``. Extension."""
        return self._read()["style"]  # type: ignore[no-any-return]

    @style.setter
    def style(self, value: str) -> None:
        self._apply(style=value)

    def __repr__(self) -> str:
        """``<Font across paragraphs bold=None ...>``: the read, None where the runs differ."""
        read = self._read()
        inside = " ".join(f"{name}={read[name]!r}" for name in ("bold", "italic", "underline", "name", "size"))
        return f"<Font across paragraphs {inside}>"
