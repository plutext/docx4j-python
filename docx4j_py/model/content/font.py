"""``Font``: Office JS ``Word.Font`` over the direct run properties of a scope.

CR-003 section 3.2. A ``Font`` is not a thing in the tree: it is a view over the
``w:rPr`` of however many runs its scope covers --- the runs of a paragraph's
text, or the runs a :class:`~docx4j_py.model.content.Range` covers.

**Reads report every run in scope**, as Office JS does (CR-016 Phase B; core-ts
``Font.mts`` at ``222cc20``, its CR-002 section 59, held to Word 365 by core-ts
checks 52 and 53): each run's *effective* formatting --- the document defaults,
the paragraph style's run properties, the run's character style and then its
direct formatting, resolved by CR-002 Phase B's
:class:`~docx4j_py.model.property_resolver.PropertyResolver` (CR-002 section
15.1; CR-003 section 6) --- and a property where the runs differ reads
``None``, ``""`` for the name, the colour and the highlight, and ``"Mixed"`` for
the underline; where they agree, their value. So the getters are typed
``bool | None`` and ``float | None`` (CR-016 Q10). With **no run** in scope, a
paragraph with no text reads its mark's run properties (``w:pPr/w:rPr``
without its revisions; core-ts check 52 case 60), and a collapsed range its
paragraph's style. **A read splits no run** (CR-016 Q12); only a write does.
``paragraph.get_font(direct=True)`` and ``range.get_font(direct=True)`` read the
runs' own ``w:rPr`` by the same rule, which is what a caller reproducing markup
wants; a font over something with no package behind it reads direct, because
there is no hierarchy to resolve against.

**An automatic colour** (``w:val="auto"``, or none anywhere) reads ``#000000``
in the effective view, as Office JS reads it (core-ts check 53; CR-016 Phase B,
the user's decision of 2026-10-24 in CR-009 section 12.9); the direct view says
what is written, ``"auto"`` or ``""``. It is the API's answer, not the drawn
pixel: Word draws automatic text white over a dark shading.

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

import dataclasses
from collections.abc import Callable
from typing import Any

from docx4j_py.model.content.enums import UnderlineValue
from docx4j_py.model.values import AUTO, CallerLimit, measure_in_range
from docx4j_py.wml import RPr, apply_run_options, read_run_options

__all__ = ["Font", "MixedFont", "mark_run_properties"]

#: What an automatic colour reads in the effective view, as Office JS reads it (CR-016 Phase B).
AUTOMATIC = "#000000"

#: Office JS's own limit for ``font.size``: refused above 1,638 pt and at 0 (core-ts check 71; CR-022
#: section 3.2, Q1), though Word stores to 32,767.5 pt. 0 here keeps its documented meaning, the
#: element removed (CR-022 section 10, departure 2), so it never reaches the check.
OFFICE_JS_FONT_SIZE = CallerLimit(low=0, high=1638, low_open=True)

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
    reference resolves to automatic: the read of ``w:val`` stands then, and an
    automatic or unstated colour reads ``#000000`` (:data:`AUTOMATIC`; CR-016
    Phase B, as core-ts's ``readingOf`` at ``76fb45d``).
    """
    if color is None:
        return None
    reference = color.theme_color
    if reference is None or getattr(reference, "value", reference) == "none":
        return None
    resolved = resolver.resolve_color(color)
    return None if resolved is None else f"#{resolved.hex()}"


#: ``w:rPr``'s fields, which a mark's ``w:pPr/w:rPr`` shares but for its revisions.
_RPR_FIELDS = tuple(field.name for field in dataclasses.fields(RPr) if field.name != "r_pr_change")


def mark_run_properties(paragraph: Any) -> Any:
    """A paragraph mark's run properties as a ``w:rPr`` a :class:`Font` reads, or None. Extension.

    ``w:pPr/w:rPr`` without the mark's own revisions (``w:ins``, ``w:del``,
    ``w:moveFrom``, ``w:moveTo``, ``w:rPrChange``): core-ts ``markHolderOf``
    (``Font.mts`` at ``222cc20``; CR-016 Phase B). A new ``w:rPr`` over the
    mark's own property elements, for reading; it is in no tree.
    """
    p_pr = getattr(paragraph, "p_pr", None)
    mark = getattr(p_pr, "r_pr", None) if p_pr is not None else None
    if mark is None:
        return None
    return RPr(**{name: getattr(mark, name, None) for name in _RPR_FIELDS})


def _reading(r_pr: Any, resolver: Any, p_pr: Any, selector: Any) -> dict[str, Any]:
    """One run's reading: its effective properties where there is a resolver, else its own (CR-016 Phase B)."""
    if resolver is None:
        out = read_run_options(r_pr)
        out.update(_theme_attributes(getattr(r_pr, "color", None)))
        return out
    effective = resolver.get_effective_rpr(r_pr, p_pr)
    out = read_run_options(effective)
    # The reads the effective w:rPr cannot answer on its own: a run whose face
    # comes from the theme has a w:asciiTheme and no w:ascii, so the name is the
    # selector's (CR-002 section 15.3); a theme colour is drawn in the theme's
    # colour, not the w:val beside it (CR-009 Phase B); and an automatic colour
    # reads #000000, as Office JS reads it (CR-016 Phase B).
    if selector is not None:
        out["name"] = selector.ascii_font_name(effective)
    color = getattr(effective, "color", None)
    drawn = _drawn(resolver, color)
    if drawn is not None:
        out["color"] = drawn
    elif out["color"] in ("", AUTO):
        out["color"] = AUTOMATIC
    out.update(_theme_attributes(color))
    return out


_FLAGS = ("bold", "italic", "strike_through", "double_strike_through", "subscript", "superscript")


def _combined(readings: list[dict[str, Any]]) -> dict[str, Any]:
    """Every run's reading as one, Office JS's way (core-ts ``Font.value`` at ``222cc20``; CR-016 Phase B).

    A property where every reading agrees reads that value; where they differ,
    ``None``, ``""`` for the name, the colour, the highlight and the style, and
    ``"Mixed"`` for the underline. With no reading at all (a range across
    paragraphs holding no run), ``None``, ``""`` and ``"Mixed"`` the same way.
    """

    if len(readings) == 1:
        # one run: its reading, as the rule below gives it, without comparing
        out = dict(readings[0])
        for key in ("name", "color", "style"):
            out[key] = out[key] or ""
        out["underline"] = out["underline"] or "Mixed"
        return out

    def value(key: str, mixed: Any = None) -> Any:
        if not readings:
            return mixed
        first = readings[0].get(key)
        for reading in readings[1:]:
            if reading.get(key) != first:
                return mixed
        return first

    out = {key: value(key) for key in _FLAGS}
    out["underline"] = value("underline", "Mixed") or "Mixed"
    out["name"] = value("name", "") or ""
    out["size"] = value("size")
    out["color"] = value("color", "") or ""
    out["highlight_color"] = value("highlight_color", "") if readings else None
    out["style"] = value("style", "") or ""
    # w:color's theme attributes where every run agrees, else None (CR-009 Phase B)
    out.update({key: value(key) for key in THEME_KEYS})
    return out


class Font:
    """A subset of Office JS ``Word.Font``, over the ``w:rPr`` of a scope."""

    __slots__ = (
        "_context",
        "_direct",
        "_guard",
        "_holders",
        "_mark",
        "_reads",
        "_record",
        "_scope",
        "_tracking",
    )

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
        reads: Callable[[], list[Any]] | None = None,
        mark: Callable[[], Any] | None = None,
    ) -> None:
        """Build the view.

        Args:
            holders: called on every write; returns the objects whose ``r_pr``
                a write formats, in document order. A callable rather than a
                list because a range splits its runs at the boundaries the
                moment a write asks for them. Every read too, where `reads` is
                None.
            scope: what this font is over, for :meth:`__repr__`.
            record: the ``paragraph.formatting(name, read)`` context manager, so
                that a write records a ``"format"``
                :class:`~docx4j_py.model.content.reports.ChangeReport` with the
                direct values either side (CR-003 section 3.4). None for a font
                over something with no paragraph behind it.
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
            direct: read the runs' own ``w:rPr`` and not the effective one.
            guard: called before every write, before anything is split or
                recorded; it raises to refuse the write. A range's and a
                paragraph's refuse one over text read from a kept
                ``mc:AlternateContent`` branch (CR-011 Phase F).
            reads: called on every read; the runs a read takes, splitting
                nothing (CR-016 Q12): a paragraph's text's, a range's.
            mark: called where a read takes no run; the run properties read in
                their place (:func:`mark_run_properties`, a paragraph's mark),
                or None, which reads the paragraph's style (CR-016 Phase B).
        """
        self._holders = holders
        self._scope = scope
        self._record = record
        self._tracking = tracking
        self._context = context
        self._direct = direct
        self._guard = guard
        self._reads = reads
        self._mark = mark

    # -- the mapping -------------------------------------------------------

    def _read_properties(self) -> list[Any]:
        """The ``w:rPr`` of each run a read takes, in order; the mark's (or None) where it takes none."""
        holders = (self._reads if self._reads is not None else self._holders)()
        if holders:
            return [getattr(holder, "r_pr", None) for holder in holders]
        return [self._mark() if self._mark is not None else None]

    def _readings(self, *, direct: bool | None = None) -> list[dict[str, Any]]:
        """Each run's reading, effective or direct (:func:`_reading`): the same object read once."""
        return self.read_each(self._read_properties(), direct=self._direct if direct is None else direct)

    def read_each(self, properties: list[Any], *, direct: bool = False) -> list[dict[str, Any]]:
        """Each of `properties` read as this font reads a run, in this font's paragraph. Extension.

        The package's resolver, the paragraph's ``w:pPr`` and the run font
        selector are found once; a ``w:rPr`` object met twice (None, a run with
        no properties) is read once.
        """
        resolver = p_pr = selector = None
        if not direct and self._context is not None:
            resolver, p_pr = self._context()
            if resolver is not None:
                selector = self._selector(resolver)
        out: list[dict[str, Any]] = []
        seen: dict[int, dict[str, Any]] = {}
        for r_pr in properties:
            found = seen.get(id(r_pr))
            if found is None:
                found = _reading(r_pr, resolver, p_pr, selector)
                seen[id(r_pr)] = found
            out.append(found)
        return out

    def _read(self) -> dict[str, Any]:
        return _combined(self._readings())

    @staticmethod
    def _selector(resolver: Any) -> Any:
        """The package's run font selector, where one has been built."""
        package = getattr(resolver, "package", None)
        if package is None:
            return None
        try:
            return package.get_run_font_selector()
        except Exception:  # noqa: BLE001 - a name is not worth an exception
            return None

    def _read_direct(self) -> dict[str, Any]:
        return _combined(self._readings(direct=True))

    def _apply(self, **options: Any) -> None:
        if self._guard is not None:
            self._guard()
        if self._record is None:
            self._write(**options)
            return
        name = next(iter(options))
        # the change report is about what the *markup* says, either side of the
        # write, so it reads direct formatting even where this font resolves:
        # every run's, None where they differ (CR-016 Phase B)
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
    def bold(self) -> bool | None:
        """``w:b`` (and ``w:bCs``); None where the runs differ. Office JS ``Word.Font.bold``."""
        return self._read()["bold"]

    @bold.setter
    def bold(self, value: bool) -> None:
        self._apply(bold=bool(value))

    @property
    def italic(self) -> bool | None:
        """``w:i`` (and ``w:iCs``); None where the runs differ. Office JS ``Word.Font.italic``."""
        return self._read()["italic"]

    @italic.setter
    def italic(self, value: bool) -> None:
        self._apply(italic=bool(value))

    @property
    def strike_through(self) -> bool | None:
        """``w:strike``; None where the runs differ. Office JS ``Word.Font.strikeThrough``."""
        return self._read()["strike_through"]

    @strike_through.setter
    def strike_through(self, value: bool) -> None:
        self._apply(strike_through=bool(value))

    @property
    def double_strike_through(self) -> bool | None:
        """``w:dstrike``; None where the runs differ. Office JS ``Word.Font.doubleStrikeThrough``."""
        return self._read()["double_strike_through"]

    @double_strike_through.setter
    def double_strike_through(self, value: bool) -> None:
        self._apply(double_strike_through=bool(value))

    @property
    def subscript(self) -> bool | None:
        """``w:vertAlign w:val="subscript"``; None where the runs differ. Office JS ``Word.Font.subscript``."""
        return self._read()["subscript"]

    @subscript.setter
    def subscript(self, value: bool) -> None:
        self._apply(subscript=bool(value))

    @property
    def superscript(self) -> bool | None:
        """``w:vertAlign w:val="superscript"``; None where the runs differ. Office JS ``Word.Font.superscript``."""
        return self._read()["superscript"]

    @superscript.setter
    def superscript(self, value: bool) -> None:
        self._apply(superscript=bool(value))

    @property
    def underline(self) -> UnderlineValue:
        """``w:u``, as an Office JS ``Word.UnderlineType`` name; ``"Mixed"`` where the runs differ.

        ``"None"`` when there is none. The value as Word reads it (CR-018 Phase
        B): one outside the enumeration (``SINGLE``, ``bogus``, ``""``) is no
        underline of the run's own, so ``"None"`` here and the style's beneath
        in the effective read, where until 0.2.0 it read ``"Mixed"``; ``w:u``'s
        six on/off words are ``none``. Setting ``"None"`` removes the element.
        """
        return self._read()["underline"]  # type: ignore[no-any-return]

    @underline.setter
    def underline(self, value: UnderlineValue | bool) -> None:
        self._apply(underline=value)

    @property
    def name(self) -> str:
        """The face the runs' ASCII text is set in; ``""`` where the runs differ.

        ``w:rFonts/@w:ascii``, or the face ``w:asciiTheme`` names, else
        ``w:hAnsi``, else the document default --- the effective read, through
        :meth:`~docx4j_py.model.fonts.RunFontSelector.ascii_font_name` (CR-002
        section 15.3). ``get_font(direct=True).name`` is the runs' own
        ``w:ascii``, ``""`` where they set none.

        Setting writes ``w:ascii`` and ``w:hAnsi`` together and clears the theme
        fonts, as Word does; ``""`` removes ``w:rFonts``.
        """
        return self._read()["name"]  # type: ignore[no-any-return]

    @name.setter
    def name(self, value: str) -> None:
        self._apply(name=value)

    @property
    def size(self) -> float | None:
        """The size in points (``w:sz`` is half-points); None where the runs differ.

        ``0`` in the direct view where the runs set none. Setting ``0`` removes
        ``w:sz`` and ``w:szCs``, which is what writing back what a silent run
        reads means.

        Setting a size above 1,638 pt or below 0 is refused before anything
        changes, as Office JS refuses it (CR-022 Q1): ``MeasureRangeError``,
        ``value.out_of_range``. Office JS refuses 0 too; here 0 keeps its
        meaning, the element removed (CR-022 section 10, departure 2).
        """
        return self._read()["size"]  # type: ignore[no-any-return]

    @size.setter
    def size(self, value: float) -> None:
        if value is not None and value != 0:
            measure_in_range("w:rPr/w:sz/@w:val", value, label="font size", limit=OFFICE_JS_FONT_SIZE)
        self._apply(size=value)

    @property
    def color(self) -> str:
        """``w:color`` as ``"#RRGGBB"``, which is what Office JS reports; ``""`` where the runs differ.

        **Automatic reads ``#000000``** in the effective read --- ``w:val="auto"``,
        or no colour anywhere --- as Office JS reads it (CR-016 Phase B; core-ts
        check 53); the direct read says what is written, ``"auto"`` or ``""``.

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
        classic HTML colour name (``red``, ``DarkRed``), as ``"#RRGGBB"``; a
        value Word rejects (``Auto``, ``F00``, ``grey``) is read as absent ---
        the effective read then gives the style's colour beneath, and the direct
        read ``""``. Setting accepts ``"#FF0000"``, ``"FF0000"`` and a name,
        which is written as its ``RRGGBB`` with no theme reference.
        """
        return self._read()["color"]  # type: ignore[no-any-return]

    @color.setter
    def color(self, value: str) -> None:
        self._apply(color=value)

    @property
    def highlight_color(self) -> str | None:
        """``w:highlight`` as ``"#RRGGBB"``, None when there is none, ``""`` where the runs differ.

        As Word reads it (CR-018 Phase B): a value outside the enumeration
        (``YELLOW``, ``bogus``) is None, as is ``none``; and the effective read
        never takes a character style's highlight, which Word does not draw
        (round six), where a paragraph style's and the run's own are drawn.
        Setting takes a highlight name (``"yellow"``) or one of its ``#RRGGBB``
        values; None removes it.
        """
        return self._read()["highlight_color"]  # type: ignore[no-any-return]

    @highlight_color.setter
    def highlight_color(self, value: str | None) -> None:
        self._apply(highlight_color=value)

    # -- extensions --------------------------------------------------------

    @property
    def style(self) -> str:
        """The character style id (``w:rStyle``); ``""`` when none, or where the runs differ. Extension."""
        return self._read()["style"]  # type: ignore[no-any-return]

    @style.setter
    def style(self, value: str) -> None:
        self._apply(style=value)

    def to_dict(self) -> dict[str, Any]:
        """Every property this font reports, JSON-ready, as the getters read them. Extension.

        With ``theme_color``, ``theme_tint`` and ``theme_shade``, the
        ``w:color``'s theme attributes as written (effective or direct as this
        font reads; None where absent, or where the runs differ), so a reader
        sees both the reference and the colour it draws (CR-009 Phase B).
        ``color`` is the effective reading's, ``#000000`` for automatic
        (CR-016 Phase B).
        """
        return dict(self._read())

    def __repr__(self) -> str:
        """``<Font bold italic 12.0pt #FF0000>``, or ``<Font plain>``; a property where the runs differ is left out."""
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

    CR-013 Phase B (core-ts ``MixedFont``, ``Font.mts`` at ``ef57729``; Q4),
    and since CR-016 Phase B :class:`Font`'s own rule (core-ts at
    ``222cc20``): a property where the runs differ reads as Office JS reads
    it (core-ts checks 52 and 53) --- ``None``, ``""`` for the name, the
    colour and the highlight, and ``"Mixed"`` for the underline --- and where
    they agree, that value. A write (CR-013 Phase C) goes to `write`: every
    run and mark in the range, as Word formats them
    (:func:`~docx4j_py.model.content.span_edits.format_span`).
    """

    __slots__ = ("_readings_of", "_write_across")

    def __init__(
        self, readings: Callable[[], list[dict[str, Any]]], write: Callable[[dict[str, Any]], None]
    ) -> None:
        """Build the view.

        Args:
            readings: called on every read; each run's reading, effective or
                direct as the range asked, read in its own paragraph
                (:meth:`Font.read_each`).
            write: called by every setter with the one option written
                (``{"bold": True}``).
        """
        super().__init__(list)
        self._readings_of = readings
        self._write_across = write

    def _readings(self, *, direct: bool | None = None) -> list[dict[str, Any]]:
        return self._readings_of()

    def _apply(self, **options: Any) -> None:
        self._write_across(options)

    def __repr__(self) -> str:
        """``<Font across paragraphs bold=None ...>``: the read, None where the runs differ."""
        read = self._read()
        inside = " ".join(f"{name}={read[name]!r}" for name in ("bold", "italic", "underline", "name", "size"))
        return f"<Font across paragraphs {inside}>"
