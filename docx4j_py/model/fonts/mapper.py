"""``Mapper``: a document font name to a physical font. docx4j ``org.docx4j.fonts.Mapper``.

CR-002 section 15.3. One order of precedence for every mapper (docx4j CR-016
phase 3), :meth:`Mapper.populate_font_mappings` first and then the passes
``WordprocessingMLPackage.setFontMapper`` runs after it:

1. the font itself where the environment has it, by its name --- ``INSTALLED``
2. the document's embedded form of it --- ``EMBEDDED``
3. the mapper's own answer, a variant of the name for
   :class:`IdentityPlusMapper` --- ``MAPPER_OWN``
4. :meth:`~Mapper.add_metrically_compatible_substitutes`: ``font-substitutes.xml``,
   best first --- ``METRIC_CLONE``, ``MEASURED_STAND_IN``, ``CLASS``
5. :meth:`~Mapper.add_alt_name_substitutes`: the document's own ``w:altName``,
   the chain followed --- ``ALT_NAME``
6. :meth:`~Mapper.add_class_based_substitutes`: a face of the same class --- ``CLASS``
7. :meth:`~Mapper.add_word_default_substitutes`: what Word itself draws a font it
   cannot find --- ``WORD_DEFAULT``
8. :meth:`~Mapper.add_no_bold_face_aliases`: re-maps a family with no bold face of
   its own

and what is left is ``UNMAPPED``. ``BestMatchingMapper``'s own step (a panose
match, then ``FontSubstitutions.xml``) would go at 3 and
:meth:`~Mapper.add_mapper_substitutes` between 6 and 7; it needs a font file's
panose and is a later CR, so ``add_mapper_substitutes`` is the empty hook
docx4j's is and :class:`IdentityPlusMapper` is the only mapper here.

**What is not ported** (CR-002 section 6.3): everything about *files* --- the
embedded-font metrics, the glyph checks, the FOP configuration,
``WordLineMetrics``, ``WidthFactors`` and the per-script coverage choices a
conversion records against a decision. :meth:`~Mapper.register_line_metrics_alias`
is kept because the alt-name and Word-default passes are written over it and a
consumer that does measure lines needs to know whose metrics Word used, but
nothing here reads a metric.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterable
from typing import Any

from docx4j_py.model.fonts.data import (
    EAST_ASIAN_FAMILIES,
    LINE_METRIC_FAMILIES,
    MS_FONT_NAMES,
    MS_FONTS_WITH_BOLD,
    SUBSTITUTE_ROWS,
    Substitute,
)
from docx4j_py.model.fonts.fallback import (
    left_to_the_document_default,
    select_by_class,
    substitution_class,
)
from docx4j_py.model.fonts.physical_font import PhysicalFont
from docx4j_py.model.fonts.registry import DEFAULT_FONT_REGISTRY, FontRegistry

__all__ = [
    "FONT_FALLBACK",
    "UNKNOWN_ERROR",
    "FontDecision",
    "IdentityPlusMapper",
    "Mapper",
    "font_table_of",
    "has_bold_face",
    "is_east_asian_family",
    "is_known_family",
    "word_default_for",
]

#: What the report says where nothing was measured.
UNKNOWN_ERROR = "unknown"

#: Word's own fallback where a document font resolves to nothing at all.
FONT_FALLBACK = "Times New Roman"

#: docx4j ``WordLineMetrics.ALIASES``: families the table reaches through a built-in alias.
_LINE_METRIC_ALIASES = {
    "helvetica": "arial",
    "helvetica neue": "arial",
    "helveticaneue": "arial",
    "helv": "arial",
    "arialmt": "arial",
    "timesnewromanpsmt": "times new roman",
}


def is_east_asian_family(document_font: str | None) -> bool:
    """docx4j ``WordLineMetrics.isEastAsianFamily``: the table's optional field 7."""
    if document_font is None:
        return False
    key = document_font.strip().lower()
    alias = _LINE_METRIC_ALIASES.get(key)
    families = EAST_ASIAN_FAMILIES()
    return key in families or (alias is not None and alias in families)


def _is_line_metric_family(document_font_name: str | None) -> bool:
    """docx4j ``WordLineMetrics.isTableFamily``: the table's families and its aliases."""
    if document_font_name is None:
        return False
    key = document_font_name.strip().lower()
    families = LINE_METRIC_FAMILIES()
    if key in families:
        return True
    alias = _LINE_METRIC_ALIASES.get(key)
    return alias is not None and alias in families


def is_known_family(document_font_name: str | None) -> bool:
    """Whether docx4j's tables know this family.

    ``MicrosoftFonts.xml``, ``word-line-metrics`` (512 Microsoft and Office cloud
    families) or ``FontSubstitutions.xml``.

    A known family the environment merely lacks is substituted for its widths
    (the passes before :meth:`Mapper.add_word_default_substitutes`), since Word on
    the author's machine had the font; an unknown one is what Word itself could
    not find either. "Known" is
    :func:`~docx4j_py.model.fonts.fallback.substitution_class`, **not**
    ``class_of``, which will also guess a class from a name that merely ends in
    "Sans": docx4j measured a corporate face that fell between the two and was
    drawn in the document default's serif throughout. A family deliberately left
    to the document default is "known" for this purpose too, the measurement that
    put it there being that the default beat a stand-in.
    """
    if document_font_name is None:
        return False
    name = document_font_name.strip()
    if name in MS_FONT_NAMES():
        return True
    if _is_line_metric_family(name):
        return True
    if substitution_class(name) != "UNKNOWN":
        return True
    return left_to_the_document_default(name)


_WEIGHT_WORDS = frozenset(
    {
        "light",
        "semilight",
        "semibold",
        "demibold",
        "medium",
        "black",
        "thin",
        "extralight",
        "ultralight",
        "heavy",
    }
)


def has_bold_face(document_font_name: str | None) -> bool:
    """Whether this document font has a bold face of its own.

    The ``MicrosoftFonts.xml`` entry where there is one (Franklin Gothic Book has
    none), else a family whose name ends in a weight word (Calibri Light, Segoe
    UI Semibold, Arial Black) is a single weight in Windows' font model and has
    none; any other family is taken to have one.
    """
    if document_font_name is None:
        return True
    name = document_font_name.strip()
    if name in MS_FONT_NAMES():
        return name in MS_FONTS_WITH_BOLD()
    last = name.lower()
    space = last.rfind(" ")
    if space > 0:
        last = last[space + 1 :]
    return last not in _WEIGHT_WORDS


def _source_of(substitute: Substitute | None) -> str:
    """The source a table row's quality word says."""
    if substitute is None:
        return "CLASS"
    if substitute.quality == "metric":
        return "METRIC_CLONE"
    if substitute.quality == "measured":
        return "MEASURED_STAND_IN"
    return "CLASS"


def font_table_of(fonts: Any) -> dict[str, Any]:
    """``word/fontTable.xml``'s entries, keyed by lower-cased name."""
    table: dict[str, Any] = {}
    for font in getattr(fonts, "font", None) or ():
        if font is not None and font.name is not None:
            table[font.name.strip().lower()] = font
    return table


def word_default_for(document_font_name: str, table: dict[str, Any]) -> str:
    """The font Word draws an unknown font in, by its ``w:fonts`` entry.

    Measured (docx4j CR-016 probe ``fonts-unresolvable``, Word 365): a face whose
    entry says ``w:family="roman"`` is drawn in **Cambria**, one saying ``swiss``
    in **Calibri**, one with an entry that names no family in Calibri, one with no
    ``w:fonts`` entry at all in Cambria; an exact panose does not change that
    (Arial's gave Calibri), and the document's default font is never used.
    """
    entry = table.get(document_font_name.strip().lower())
    if entry is None:
        return "Cambria"
    # the family: the entry's own, else the first along its w:altName chain
    seen: set[str] = set()
    while entry is not None and entry.name is not None:
        key = entry.name.strip().lower()
        if key in seen:
            break
        seen.add(key)
        family = getattr(entry.family, "val", None)
        family = getattr(family, "value", family)
        if family is not None:
            value = str(family).strip().lower()
            if value == "roman":
                return "Cambria"
            if value == "modern":
                return "Courier New"
            return "Calibri"  # swiss, script, decorative, auto
        alt = getattr(entry.alt_name, "val", None)
        entry = None if alt is None else table.get(alt.strip().lower())
    return "Calibri"


@dataclasses.dataclass(frozen=True, slots=True)
class FontDecision:
    """What was decided for one document font, and why. docx4j ``FontDecision``."""

    document_font: str
    """The name as the document has it."""
    source: str
    """``INSTALLED``, ``EMBEDDED``, ``METRIC_CLONE``, ``MEASURED_STAND_IN``,
    ``ALT_NAME``, ``CLASS``, ``WORD_DEFAULT``, ``MAPPER_OWN``, ``SYMBOL`` or
    ``UNMAPPED``."""
    via: str | None = None
    """How it got there --- the ``w:altName`` chain, the class, Word's default
    family --- or None where the source says it all."""
    width_error: str | None = None
    """What is known of the substitute's width error, or None."""
    physical_font: PhysicalFont | None = None
    """The physical font as the mapping now stands; None where nothing mapped it."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, as the parity goldens record a decision."""
        return {
            "source": self.source,
            "via": self.via,
            "physicalFont": None if self.physical_font is None else self.physical_font.name,
        }


class Mapper:
    """Maps the document's font names to physical fonts, over a :class:`FontRegistry`.

    One per package, as in docx4j.
    """

    __slots__ = (
        "_bold_forms",
        "_bold_italic_forms",
        "_decisions",
        "_font_mappings",
        "_italic_forms",
        "_line_metrics_aliases",
        "_regular_forms",
        "_resolved_via",
        "_word_defaulted",
        "registry",
    )

    def __init__(self, registry: FontRegistry | None = None) -> None:
        """Build the mapper over `registry`, docx4j's font jars' faces by default."""
        #: The fonts the consumer has.
        self.registry: FontRegistry = (
            DEFAULT_FONT_REGISTRY if registry is None else registry
        )
        self._font_mappings: dict[str, PhysicalFont] = {}
        self._regular_forms: dict[str, PhysicalFont] = {}
        self._bold_forms: dict[str, PhysicalFont] = {}
        self._italic_forms: dict[str, PhysicalFont] = {}
        self._bold_italic_forms: dict[str, PhysicalFont] = {}
        self._decisions: dict[str, dict[str, Any]] = {}
        self._word_defaulted: set[str] = set()
        self._line_metrics_aliases: dict[str, str] = {}
        self._resolved_via: str | None = None

    # -- the mapping -------------------------------------------------------

    def get(self, key: str | None) -> PhysicalFont | None:
        """The physical font a document font maps to, by case-insensitive name.

        None for a slot nothing names, which docx4j's ``Mapper.get`` has
        tolerated since 17.1.1.
        """
        if key is None:
            return None
        return self._font_mappings.get(key.lower())

    def put(self, key: str, font: PhysicalFont) -> None:
        """Map a document font to a physical font."""
        self._font_mappings[key.lower()] = font

    def __len__(self) -> int:
        """How many document fonts are mapped."""
        return len(self._font_mappings)

    def get_decision(self, document_font: str | None) -> FontDecision | None:
        """What was decided for this document font, or None where no pass met it."""
        if document_font is None:
            return None
        record = self._decisions.get(document_font.strip().lower())
        if record is None:
            return None
        return FontDecision(
            record["documentFont"],
            record["source"],
            record["via"],
            record["widthError"],
            self.get(record["documentFont"]),
        )

    def get_decisions(self) -> list[FontDecision]:
        """Every decision, ordered by the document font's name, case-insensitively.

        So that two runs over the same document give the same list.
        """
        out = [
            FontDecision(
                r["documentFont"], r["source"], r["via"], r["widthError"], self.get(r["documentFont"])
            )
            for r in self._decisions.values()
        ]
        out.sort(key=lambda d: d.document_font.lower())
        return out

    def line_metrics_family(self, document_font: str | None) -> str | None:
        """Whose Word line metrics a document font takes, where a pass resolved it elsewhere.

        None where it takes its own. Per mapper, so per package: docx4j's was a
        JVM-wide map until 17.1.1, and one document's ``w:altName`` answered for
        every later one.
        """
        if document_font is None:
            return None
        return self._line_metrics_aliases.get(document_font.strip().lower())

    def register_line_metrics_alias(
        self, document_font: str | None, family: str | None
    ) -> None:
        """Record whose line metrics a document font takes."""
        if document_font is None or family is None:
            return
        # a font with metrics of its own takes them; there is nothing to alias
        if _is_line_metric_family(document_font):
            return
        self._line_metrics_aliases[document_font.strip().lower()] = family.strip()

    def _decide(
        self,
        document_font: str,
        source: str,
        via: str | None = None,
        width_error: str | None = None,
    ) -> None:
        self._decisions[document_font.strip().lower()] = {
            "documentFont": document_font.strip(),
            "source": source,
            "via": via,
            "widthError": width_error,
        }

    # -- embedded faces ----------------------------------------------------

    def register_regular_form(self, name: str, font: PhysicalFont | None) -> None:
        """The document's embedded regular form of a font."""
        _set(self._regular_forms, name, font)

    def register_bold_form(self, name: str, font: PhysicalFont | None) -> None:
        """The document's embedded bold form of a font."""
        _set(self._bold_forms, name, font)

    def register_italic_form(self, name: str, font: PhysicalFont | None) -> None:
        """The document's embedded italic form of a font."""
        _set(self._italic_forms, name, font)

    def register_bold_italic_form(self, name: str, font: PhysicalFont | None) -> None:
        """The document's embedded bold-italic form of a font."""
        _set(self._bold_italic_forms, name, font)

    def is_embedded(self, name: str) -> bool:
        """Whether the document embeds this font."""
        return (
            name in self._regular_forms
            or name in self._bold_forms
            or name in self._italic_forms
            or name in self._bold_italic_forms
        )

    def _installed_or_embedded(self, name: str) -> PhysicalFont | None:
        """The installed font of this name, else the document's embedded form of it.

        docx4j CR-016 Decisions 2: the installed font wins, for all four faces
        and in both mappers.
        """
        return (
            self.registry.get(name)
            or self._regular_forms.get(name)
            or self._bold_forms.get(name)
            or self._italic_forms.get(name)
            or self._bold_italic_forms.get(name)
        )

    # -- the passes --------------------------------------------------------

    def resolve_document_font(
        self, document_font_name: str, font_table_entry: Any
    ) -> PhysicalFont | None:
        """The mapper's own answer for a font neither installed nor embedded.

        None where it has none. :attr:`_resolved_via` says how it was reached.
        """
        return None

    def populate_font_mappings(self, document_font_names: Iterable[str], fonts: Any) -> None:
        """An entry for each of the document's fonts: steps 1 to 3 of the precedence."""
        table = font_table_of(fonts)
        for document_font_name in document_font_names:
            if not document_font_name or not document_font_name.strip():
                continue
            if self.get(document_font_name) is not None:
                continue  # already mapped; its decision stands
            font = self._installed_or_embedded(document_font_name)
            source: str | None = None
            if font is not None:
                source = (
                    "INSTALLED"
                    if self.registry.get(document_font_name) is not None
                    else "EMBEDDED"
                )
            else:
                self._resolved_via = None
                font = self.resolve_document_font(
                    document_font_name, table.get(document_font_name.strip().lower())
                )
                if font is not None:
                    source = "MAPPER_OWN"
            if font is None:
                # recorded, not merely logged: a later pass overwrites this where
                # it maps the font, and what is left says which fonts reached the
                # consumer with nothing
                self._decide(document_font_name, "UNMAPPED")
            else:
                self.put(document_font_name, font)
                via = None
                if source == "MAPPER_OWN":
                    via = self._resolved_via or f"the mapper's own answer: {font.name}"
                self._decide(document_font_name, source or "UNMAPPED", via)

    def add_metrically_compatible_substitutes(self) -> None:
        """``font-substitutes.xml``: the first open substitute this environment has.

        The measurements that chose each substitute are in docx4j's
        ``Mapper.addMetricallyCompatibleSubstitutes`` javadoc, and the table's
        ``error`` attributes quote them.
        """
        for row in SUBSTITUTE_ROWS:
            before = self.get(row.document_font)
            taken: Substitute | None = None
            for candidate in row.substitutes:
                if self.registry.get(candidate.font) is not None:
                    taken = candidate
                    break
            if taken is None:
                continue
            self.add_metrically_compatible_substitute(row.document_font, taken.font)
            after = self.get(row.document_font)
            if after is not None and after is not before:
                self._decide(
                    row.document_font,
                    _source_of(taken),
                    None,
                    taken.error
                    if taken.error is not None
                    else (UNKNOWN_ERROR if taken.quality == "class" else None),
                )

    def add_metrically_compatible_substitute(
        self, document_font: str, substitute: str
    ) -> None:
        """Map a document font to an open substitute.

        Unless the environment has the font itself or the document embeds it: the
        embedded font is what the author intended, and it is the only thing
        certain to be available.
        """
        if self.is_embedded(document_font):
            return
        if self.registry.get(document_font) is not None:
            return
        font = self.registry.get(substitute)
        if font is not None:
            self.put(document_font, font)

    def add_alt_name_substitutes(
        self, document_font_names: Iterable[str] | None, fonts: Any
    ) -> None:
        """``w:altName`` in ``word/fontTable.xml`` (ECMA-376 17.8.3.1).

        "The name of an alternate font which shall be used if the font specified
        is not available": the document author's own answer to a missing font,
        and Word takes it.

        After the metric clones --- a clone of the document's own font is a
        better answer than any alias --- and before the class pass, which is only
        a guess from the name. The chain is followed: the alternate may itself be
        absent and name an alternate (docx4j CR-016 probe ``fonts-unresolvable``
        (e), (f)); a few hops, never a cycle.
        """
        alt_names: dict[str, str] = {}
        for font in getattr(fonts, "font", None) or ():
            alt = getattr(getattr(font, "alt_name", None), "val", None)
            if font is None or font.name is None or alt is None or not alt.strip():
                continue
            alt_names[font.name.strip().lower()] = alt.strip()
        if not alt_names:
            return

        names = document_font_names if document_font_names is not None else list(alt_names)
        for document_font_name in names:
            if not document_font_name or not document_font_name.strip():
                continue
            if self.get(document_font_name) is not None:
                continue  # already mapped
            if self.is_embedded(document_font_name):
                continue
            if self.registry.get(document_font_name) is not None:
                continue  # installed; identity

            alt: str | None = alt_names.get(document_font_name.strip().lower())
            seen = {document_font_name.strip().lower()}
            font = None
            resolved_alt: str | None = None
            east_asian_hop: str | None = None
            chain: list[str] = []
            while alt is not None and alt.strip().lower() not in seen:
                seen.add(alt.strip().lower())
                chain.append(alt)
                # A hop this environment lacks but whose family the line-metrics
                # table flags East Asian: Word had that font and drew it, and its
                # line is 1.3 x its usWin box, which no substitute comes near. The
                # chain goes on for the *width* substitution; only the line box is
                # taken here.
                if (
                    east_asian_hop is None
                    and is_east_asian_family(alt)
                    and self.registry.get(alt) is None
                ):
                    east_asian_hop = alt
                font = (
                    self.registry.get(alt)
                    or self.registry.get(f"{alt} Regular")
                    or self.get(alt)
                )
                if font is not None:
                    resolved_alt = alt
                    break
                alt = alt_names.get(alt.strip().lower())
            # The line box is registered even where the chain found no face: Word
            # had the alternate and drew it, so the line is its, whatever renders
            # the glyphs.
            if east_asian_hop is not None:
                self.register_line_metrics_alias(document_font_name, east_asian_hop)
            if font is None:
                continue

            self.put(document_font_name, font)
            if east_asian_hop is None:
                self.register_line_metrics_alias(document_font_name, resolved_alt)
            self._decide(
                document_font_name, "ALT_NAME", f"w:altName {' -> '.join(chain)}"
            )

    def wants_class_based_substitutes(self) -> bool:
        """Whether the class pass applies. True for every mapper since docx4j 17.1.1."""
        return True

    def add_class_based_substitutes(self, document_font_names: Iterable[str] | None) -> None:
        """Map whatever is still unmapped to a font of the same class.

        Without it an unmapped font falls back to whatever the document's
        *default* font maps to, which is a Times clone standing in for a sans as
        often as not. Deliberately conservative: a condensed face is left
        unmapped, its widths being further from an ordinary condensed face's than
        the default's are.
        """
        if document_font_names is None:
            return
        for document_font_name in document_font_names:
            if not document_font_name or not document_font_name.strip():
                continue
            if self.get(document_font_name) is not None:
                continue
            if self.is_embedded(document_font_name):
                continue
            if self.registry.get(document_font_name) is not None:
                continue

            font = select_by_class(self.registry, document_font_name)
            if font is not None:
                self.put(document_font_name, font)
                self._decide(
                    document_font_name,
                    "CLASS",
                    f"a face of the same class: {substitution_class(document_font_name)}",
                    UNKNOWN_ERROR,
                )

    def add_mapper_substitutes(
        self, document_font_names: Iterable[str] | None, fonts: Any
    ) -> None:
        """The mapper's own guesses for what the measured passes left.

        ``BestMatchingMapper``'s panose match and ``FontSubstitutions.xml``.
        Nothing here: that mapper needs font files.
        """

    def add_word_default_substitutes(
        self, document_font_names: Iterable[str] | None, fonts: Any
    ) -> None:
        """Word's own answer for a font it cannot find, for whatever is still unmapped.

        Cambria or Calibri (or Courier New) by the ``w:fonts`` entry's family
        (:func:`word_default_for`), mapped to whatever this mapper maps that font
        to. Applied only to a family none of docx4j's tables know
        (:func:`is_known_family`).
        """
        if document_font_names is None:
            return
        table = font_table_of(fonts)
        for document_font_name in document_font_names:
            if not document_font_name or not document_font_name.strip():
                continue
            if self.get(document_font_name) is not None:
                continue
            if self.is_embedded(document_font_name):
                continue
            if self.registry.get(document_font_name) is not None:
                continue
            if is_known_family(document_font_name):
                continue

            word_font = word_default_for(document_font_name, table)
            font = (
                self.get(word_font)
                or self.registry.get(word_font)
                or select_by_class(self.registry, word_font)
            )
            if font is None:
                continue
            self.put(document_font_name, font)
            self._word_defaulted.add(document_font_name.strip().lower())
            self.register_line_metrics_alias(document_font_name, word_font)
            self._decide(
                document_font_name,
                "WORD_DEFAULT",
                f"Word's own default for a font it cannot find: {word_font}",
                UNKNOWN_ERROR,
            )

    def add_no_bold_face_aliases(self, document_font_names: Iterable[str] | None) -> None:
        """An alias reporting no bold face, for each font that has none of its own.

        So that the consumer synthesises bold at the regular advances as Word
        does. Last of the passes: it re-maps, and it is not a source of its own
        --- the decision says so by reporting a synthetic bold face.
        """
        if document_font_names is None:
            return
        for document_font_name in document_font_names:
            if not document_font_name or not document_font_name.strip():
                continue
            if has_bold_face(document_font_name):
                continue
            if self.is_embedded(document_font_name):
                continue  # the embedded forms say what the document has
            # A font Word itself could not find is substituted whole --- Calibri's
            # real bold for its w:b --- however its name reads: "EnBW DIN Pro
            # Light" is Calibri Bold in Word, not a synthesised Calibri. Measured
            # (docx4j CR-016 phase 4 gate): with the alias that document scored
            # 0.8441, without it 0.8783.
            if document_font_name.strip().lower() in self._word_defaulted:
                continue
            font = self.get(document_font_name)
            if font is None or font.no_bold_face:
                continue
            self._font_mappings[document_font_name.lower()] = font.no_bold_face_alias()

    def populate(self, document_font_names: Iterable[str], fonts: Any = None) -> None:
        """Every pass in order, as docx4j's ``setFontMapper`` runs them.

        Args:
            document_font_names: the document's font names
                (``MainDocumentPart.fonts_in_use()``).
            fonts: the font table part's content, for ``w:altName`` and ``w:family``.
        """
        names = list(document_font_names)
        self.populate_font_mappings(names, fonts)
        self.add_metrically_compatible_substitutes()
        self.add_alt_name_substitutes(names, fonts)
        if self.wants_class_based_substitutes():
            self.add_class_based_substitutes(names)
        self.add_mapper_substitutes(names, fonts)
        self.add_word_default_substitutes(names, fonts)
        self.add_no_bold_face_aliases(names)

    def __repr__(self) -> str:
        """``<IdentityPlusMapper 5 fonts>``."""
        return f"<{type(self).__name__} {len(self._font_mappings)} fonts>"


#: The order a family whose plain face is absent is asked for. Bold before italic
#: since docx4j 17.1.1: with italic first, "Franklin Gothic Demi" was set upright
#: in Franklin Gothic Demi Italic, and a bold upright is nearer a missing regular
#: than an italic is.
_VARIANTS = (" regular", " bold", " italic", " bold italic")


class IdentityPlusMapper(Mapper):
    """docx4j ``org.docx4j.fonts.IdentityPlusMapper``: the default mapper.

    Maps a document font to the physical font of the same name where the
    environment has one --- an identity mapping --- plus the shared passes in
    :class:`Mapper` (the metric clones, ``w:altName``, a face of the same class,
    Word's own default) for the fonts it lacks. Measured on three real-document
    corpora in six font environments, including boxes with no Microsoft fonts and
    images with no fonts at all, it is ahead of docx4j's ``BestMatchingMapper``
    in every one of them (docx4j CR-016 phase 0c).

    It lives in this module rather than one of its own --- docx4j gives it a
    class file --- because it is twelve lines over :class:`Mapper` and Python's
    import cost is per module (the departure CR-002 section 12.3 records for the
    WordprocessingML parts).
    """

    __slots__ = ()

    def resolve_document_font(
        self, document_font_name: str, font_table_entry: Any
    ) -> PhysicalFont | None:
        """A variant of the name: ``Noto Sans Symbols`` is registered as ``… Regular``."""
        self._resolved_via = None
        for variant in _VARIANTS:
            font = self.registry.get(document_font_name + variant)
            if font is not None:
                self._resolved_via = (
                    f"a variant of the name: {document_font_name}{variant}"
                )
                return font
        return None


def _set(mapping: dict[str, PhysicalFont], name: str, font: PhysicalFont | None) -> None:
    if font is None:
        mapping.pop(name, None)
    else:
        mapping[name] = font
