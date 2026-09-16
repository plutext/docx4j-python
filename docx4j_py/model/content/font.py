"""``Font``: Office JS ``Word.Font`` over the direct run properties of a scope.

CR-003 section 3.2. A ``Font`` is not a thing in the tree: it is a view over the
``w:rPr`` of however many runs its scope covers --- every run of a paragraph, or
just the runs a :class:`~docx4j_py.model.content.Range` has split out.

**Reads report the first run in scope's direct formatting.** Office JS reports
*effective* formatting, which needs the style hierarchy and the document
defaults; CR-002 Phase B's ``PropertyResolver`` supplies those, and CR-003
section 6 says so explicitly: until it lands the views read direct formatting
and say so. ``tests/content/test_font.py`` carries the styled case marked
``xfail`` so that the day the resolver arrives the test flips rather than being
written then.

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

__all__ = ["Font"]


class Font:
    """A subset of Office JS ``Word.Font``, over the ``w:rPr`` of a scope."""

    __slots__ = ("_holders", "_record", "_scope")

    def __init__(
        self,
        holders: Callable[[], list[Any]],
        *,
        scope: str = "",
        record: Any = None,
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
        """
        self._holders = holders
        self._scope = scope
        self._record = record

    # -- the mapping -------------------------------------------------------

    def _read(self) -> dict[str, Any]:
        holders = self._holders()
        return read_run_options(getattr(holders[0], "r_pr", None) if holders else None)

    def _apply(self, **options: Any) -> None:
        if self._record is None:
            self._write(**options)
            return
        name = next(iter(options))
        with self._record(name, lambda: self._read().get(name)):
            self._write(**options)

    def _write(self, **options: Any) -> None:
        for holder in self._holders():
            r_pr = getattr(holder, "r_pr", None)
            if r_pr is None:
                r_pr = RPr()
                holder.r_pr = r_pr
                r_pr.parent = holder
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

        ``"None"`` when there is none, ``"Mixed"`` for a ``w:u`` value Office JS
        has no name for. Setting ``"None"`` removes the element.
        """
        return self._read()["underline"]  # type: ignore[return-value]

    @underline.setter
    def underline(self, value: UnderlineValue | bool) -> None:
        self._apply(underline=value)

    @property
    def name(self) -> str:
        """``w:rFonts/@w:ascii``; ``""`` when the run sets none.

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

        A value that is not a hex triple (``auto``, a theme colour) comes back
        as it stands; ``""`` when the run sets none. Setting accepts both
        ``"#FF0000"`` and ``"FF0000"``.
        """
        return self._read()["color"]

    @color.setter
    def color(self, value: str) -> None:
        self._apply(color=value)

    @property
    def highlight_color(self) -> str | None:
        """``w:highlight`` as ``"#RRGGBB"``, or None when there is none.

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
        """Every property this font reports, JSON-ready. Extension."""
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
