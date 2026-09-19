"""``StyleUtil``: the overlay. docx4j ``org.docx4j.model.styles.StyleUtil``.

CR-002 section 6.1, the catalogue-driven half: ``apply``, ``apply_style_level``,
``is_empty``, ``unset`` and ``has_direct_formatting``, every one of them derived
from the tables of :mod:`docx4j_py.model.property_catalogue` by iteration rather
than hand-written per element, so the four cannot drift apart.

Not ported, as docx4j-core-ts did not port them either: the 73 ``areEqual``
overloads and ``StyleTree`` / ``Node`` / ``Tree`` / ``BrokenStyleRemediator``,
which serve docx4j's HTML and CSS export.

Names are docx4j's; where Java overloads on the argument type, the type is in
the name (``apply_rpr``, ``apply_ppr_base``, ``apply_tbl_pr`` and so on), as the
TypeScript port does.
"""

from __future__ import annotations

import logging
from typing import Any, Protocol

from docx4j_py.model.property_catalogue import (
    CELL,
    IND,
    PARAGRAPH,
    RUN,
    TABLE,
    TOGGLE_NAMES,
    TOGGLES,
    apply_catalogue,
    copy_leaf,
    has_direct_formatting_of,
    is_empty_of,
    is_true,
    named_property,
    unset_catalogue,
)

__all__ = [
    "CYCLE_LIMIT",
    "NumberingIndents",
    "apply_ppr",
    "apply_ppr_base",
    "apply_rpr",
    "apply_style",
    "apply_style_level",
    "apply_tbl_pr",
    "apply_tbl_style_pr",
    "apply_tbl_style_pr_list",
    "apply_tc_pr",
    "apply_toggles",
    "apply_tr_pr",
    "has_direct_formatting_ppr",
    "has_direct_formatting_rpr",
    "is_cyclic",
    "is_empty_ppr",
    "is_empty_ppr_base",
    "is_empty_rpr",
    "is_empty_style",
    "is_empty_tbl_pr",
    "is_empty_tbl_style_pr",
    "is_empty_tbl_style_pr_list",
    "is_empty_tc_pr",
    "is_empty_tr_pr",
    "numbering_off",
    "throw_on_cyclic_styles",
    "throws_on_cyclic_styles",
    "toggle",
    "unset_ppr_base",
    "unset_rpr",
]

logger = logging.getLogger(__name__)


class NumberingIndents(Protocol):
    """What the paragraph merge needs from the numbering part: a level's ``w:ind``."""

    def get_ind(self, num_pr: Any) -> Any:
        """The indent the level a ``w:numPr`` names contributes, or None."""
        ...  # pragma: no cover


def _style_type(style: Any) -> str | None:
    value = getattr(style, "type_value", None)
    return None if value is None else str(getattr(value, "value", value))


CHARACTER_STYLE = "character"
PARAGRAPH_STYLE = "paragraph"
NUMBERING_STYLE = "numbering"
TABLE_STYLE = "table"


# ---------------------------------------------------------------- is_empty


def is_empty_rpr(r_pr: Any) -> bool:
    """docx4j ``isEmpty(RPr)`` / ``isEmpty(ParaRPr)``."""
    return is_empty_of(RUN, r_pr)


def is_empty_ppr_base(p_pr_base: Any) -> bool:
    """docx4j ``isEmpty(PPrBase)``."""
    return is_empty_of(PARAGRAPH, p_pr_base)


def is_empty_ppr(p_pr: Any) -> bool:
    """docx4j ``isEmpty(PPr)``: the base members, the paragraph mark and the section.

    ``isEmpty(SectPr)`` is "anything non-null is non-empty" in docx4j, whose own
    comment calls the implementation "quite basic"; it is all ``isEmptyPPr`` uses
    it for.
    """
    return p_pr is None or (
        is_empty_ppr_base(p_pr)
        and is_empty_rpr(getattr(p_pr, "r_pr", None))
        and getattr(p_pr, "sect_pr", None) is None
    )


def is_empty_tbl_pr(tbl_pr: Any) -> bool:
    """docx4j ``isEmpty(CTTblPrBase)``."""
    return is_empty_of(TABLE, tbl_pr)


def is_empty_tc_pr(tc_pr: Any) -> bool:
    """docx4j ``isEmpty(TcPr)``."""
    return is_empty_of(CELL, tc_pr)


def is_empty_tr_pr(tr_pr: Any) -> bool:
    """docx4j ``isEmpty(TrPr)``: the generated model keeps a row's properties as one list."""
    return tr_pr is None or not tr_pr.content


def is_empty_tbl_style_pr(pr: Any) -> bool:
    """docx4j ``isEmpty(CTTblStylePr)``."""
    return pr is None or (
        is_empty_ppr(pr.p_pr)
        and is_empty_rpr(pr.r_pr)
        and is_empty_tbl_pr(pr.tbl_pr)
        and is_empty_tr_pr(pr.tr_pr)
        and is_empty_tc_pr(pr.tc_pr)
        and pr.type_value is None
    )


def is_empty_tbl_style_pr_list(entries: Any) -> bool:
    """Whether a ``w:tblStylePr`` list says nothing."""
    if not entries:
        return True
    return not any(not is_empty_tbl_style_pr(pr) for pr in entries)


def is_empty_style(style: Any) -> bool:
    """docx4j ``isEmpty(Style)``: what a style of that type can carry."""
    if style is None:
        return True
    kind = _style_type(style)
    if kind == CHARACTER_STYLE:
        return is_empty_rpr(style.r_pr)
    if kind in (PARAGRAPH_STYLE, NUMBERING_STYLE):
        return is_empty_ppr(style.p_pr) and is_empty_rpr(style.r_pr)
    return (
        is_empty_ppr(style.p_pr)
        and is_empty_rpr(style.r_pr)
        and is_empty_tbl_pr(style.tbl_pr)
        and is_empty_tc_pr(style.tc_pr)
        and is_empty_tbl_style_pr_list(style.tbl_style_pr)
    )


# ---------------------------------------------------------------- direct formatting


def has_direct_formatting_rpr(r_pr: Any) -> bool:
    """Whether a run's ``w:rPr`` states any formatting of its own.

    Any member of the run table other than the style reference is present; an
    element with no attributes counts, because it was stated.
    """
    return has_direct_formatting_of(RUN, r_pr)


def has_direct_formatting_ppr(p_pr: Any) -> bool:
    """Whether a paragraph's ``w:pPr`` states any formatting of its own.

    The paragraph mark's ``w:rPr`` and a ``w:sectPr`` are not formatting of the
    paragraph.
    """
    return has_direct_formatting_of(PARAGRAPH, p_pr)


# ---------------------------------------------------------------- unset


def unset_rpr(source: Any, destination: Any) -> None:
    """Unset in `destination` every ``w:rPr`` member that `source` states."""
    unset_catalogue(RUN, source, destination)


def unset_ppr_base(source: Any, destination: Any) -> None:
    """Unset in `destination` every ``PPrBase`` member that `source` states."""
    unset_catalogue(PARAGRAPH, source, destination)


# ---------------------------------------------------------------- apply: runs


def _skip_run(source: Any) -> bool:
    """Run properties with nothing to say --- a ``w:rFonts`` with only a ``w:hint`` still says something."""
    if source is None:
        return True
    r_fonts = getattr(source, "r_fonts", None)
    hint = r_fonts is not None and r_fonts.hint is not None
    return is_empty_of(RUN, source) and not hint


def apply_rpr(source: Any, destination: Any, cls: type | None = None) -> Any:
    """Apply run properties over run properties, in place. docx4j ``apply(RPr, RPr)``.

    A ``w:rPr`` or a ``w:pPr/w:rPr`` over either. Returns the destination,
    created as `cls` (``RPr`` by default) when it was None and the source had
    something to say.
    """
    if _skip_run(source):
        return destination
    out = destination
    if out is None:
        if cls is None:
            from docx4j_py.wml import RPr

            cls = RPr
        out = cls()
    apply_catalogue(RUN, source, out)
    return out


def apply_style_level(source: Any, destination: Any, document_defaults: Any) -> None:
    """Apply one **level** of the style hierarchy over the levels beneath it.

    The **toggle properties** are combined as ECMA-376-1 17.7.3 requires rather
    than overridden. Only a level boundary goes through here --- the table
    style's contribution against the paragraph style's, and the paragraph
    style's against the character style's. Within one level a style's
    ``w:basedOn`` chain is *not* XORed (17.7.3 takes the first value found
    walking that chain, which is what merging the chain root-first with
    :func:`apply_rpr` gives), and direct formatting is not a level either: an
    explicit value there is used as it stands, so it goes through
    :func:`apply_rpr` too.
    """
    if source is None or destination is None:
        return
    apply_catalogue(RUN, source, destination, TOGGLE_NAMES)
    apply_toggles(source, destination, document_defaults, destination)


def apply_toggles(upper: Any, lower: Any, document_defaults: Any, destination: Any) -> None:
    """Combine the twelve toggle properties of two levels and write them into a third ``w:rPr``.

    Nothing else of either is touched.
    """
    if destination is None:
        return
    for entry in TOGGLES:
        value = toggle(
            None if upper is None else entry.get(upper),
            None if lower is None else entry.get(lower),
            None if document_defaults is None else entry.get(document_defaults),
        )
        entry.set(destination, value)


def toggle(upper: Any, lower: Any, document_default: Any) -> Any:
    """One toggle property at one level boundary (ECMA-376-1 17.7.3).

    As Word's ``toggle-levels`` golden settled it (docx4j CR-015, 2026-09-18):

    * an upper level that says nothing is no boundary at all, so the lower value
      stands and the document defaults' true is not forced back on;
    * ``w:docDefaults`` is a base value, never a term of the XOR: where the
      defaults say true and the upper level states the property, the effective
      value is true;
    * an explicit false at a style level is a term of the XOR like any other
      value, so false XOR lower = lower and only a stated *true* inverts what is
      beneath it. Word draws a run in a character style stating
      ``<w:b w:val="0"/>`` over a bold paragraph style **bold**. Where no level
      beneath states the property, the one level's own false stands rather than
      the property going absent.

    Direct formatting never comes through here: an explicit value there wins
    outright.
    """
    from docx4j_py.wml import BooleanDefaultTrue

    if upper is None:
        return lower
    if document_default is not None and is_true(document_default):
        return copy_leaf(document_default)
    # false XOR lower = lower; with no level beneath it, the level's own false stands
    if not is_true(upper):
        return lower if lower is not None else upper
    return BooleanDefaultTrue(val=not is_true(lower))


# ---------------------------------------------------------------- apply: paragraphs


def numbering_off(num_pr: Any) -> bool:
    """Whether this ``w:numPr`` says "not numbered": ``w:numId w:val="0"`` (ECMA-376 17.9.18)."""
    return num_pr is not None and num_pr.num_id is not None and num_pr.num_id.val == 0


def _drop_numbering_ind(ind: Any, from_level: Any) -> None:
    """Take out of an indent the components a numbering level contributed."""
    for name in ("left", "right", "hanging", "first_line", "start", "end"):
        contributed = getattr(from_level, name)
        if contributed is not None and contributed == getattr(ind, name):
            setattr(ind, name, None)


def _ind_property() -> Any:
    entry = named_property(PARAGRAPH, IND)
    if entry is None:  # pragma: no cover - the table is a constant
        raise AssertionError("the PARAGRAPH table has no w:ind entry")
    return entry


def apply_ppr_base(source: Any, destination: Any, numbering: Any = None) -> None:
    """Apply paragraph properties over paragraph properties, in place.

    Every member but ``w:ind``, then the indent the numbering level this layer
    brings in contributes, then ``w:ind`` itself (an indent stated by the layer
    overrides the level's).
    """
    if source is None or is_empty_ppr_base(source):
        return

    from docx4j_py.wml import CtPprBaseNumPr

    # A snapshot: the w:numPr merge writes into the destination object, so
    # holding a reference here would hand the numbering-off rule below the
    # w:numId 0 it is looking past.
    inherited_num_pr = None
    if destination.num_pr is not None:
        inherited_num_pr = CtPprBaseNumPr(
            num_id=destination.num_pr.num_id, ilvl=destination.num_pr.ilvl
        )

    apply_catalogue(PARAGRAPH, source, destination, (IND,))

    # ECMA-376 17.9.18: w:numId w:val="0" takes the paragraph out of the list, so
    # neither the level's label nor its w:ind applies --- only the paragraph's
    # own. Only the components the inherited level contributed are dropped, so a
    # w:ind the style states itself survives.
    if (
        numbering is not None
        and numbering_off(source.num_pr)
        and inherited_num_pr is not None
        and inherited_num_pr.num_id is not None
        and not numbering_off(inherited_num_pr)
        and destination.ind is not None
    ):
        from_level = numbering.get_ind(inherited_num_pr)
        if from_level is not None:
            own = copy_leaf(destination.ind)
            _drop_numbering_ind(own, from_level)
            destination.ind = own

    # The indent of the level of the MERGED w:numPr, since a layer may state only
    # w:ilvl (the list from the style) or only w:numId (the level from the
    # style): CR-015's probe styles-numpr-ilvl-only.
    ind_entry = _ind_property()
    if (
        numbering is not None
        and source.num_pr is not None
        and destination.num_pr is not None
        and not numbering_off(destination.num_pr)
    ):
        num_ind = numbering.get_ind(destination.num_pr)
        if num_ind is not None:
            destination.ind = ind_entry.merge(num_ind, destination.ind)
    # an indent in the layer overrides any the numbering brought in
    destination.ind = ind_entry.merge(source.ind, destination.ind)


def apply_ppr(source: Any, destination: Any, numbering: Any = None) -> Any:
    """Apply a ``w:pPr`` over a ``w:pPr``: the base members and the paragraph mark's ``w:rPr``.

    A ``w:sectPr`` is not a paragraph property and is not merged (docx4j CR-015
    phase 4).
    """
    from docx4j_py.wml import ParaRPr, PPr

    if source is None or is_empty_ppr(source):
        return destination
    out = PPr() if destination is None else destination
    apply_ppr_base(source, out, numbering)
    out.r_pr = apply_rpr(getattr(source, "r_pr", None), out.r_pr, ParaRPr)
    return out


# ---------------------------------------------------------------- apply: tables


def apply_tbl_pr(source: Any, destination: Any) -> Any:
    """Apply a ``w:tblPr`` over a ``w:tblPr``. docx4j ``apply(CTTblPrBase, CTTblPrBase)``."""
    from docx4j_py.wml import CTTblPrBase

    if source is None or is_empty_tbl_pr(source):
        return destination
    out = CTTblPrBase() if destination is None else destination
    apply_catalogue(TABLE, source, out)
    return out


def apply_tc_pr(source: Any, destination: Any) -> Any:
    """Apply a ``w:tcPr`` over a ``w:tcPr``. docx4j ``apply(TcPr, TcPr)``."""
    from docx4j_py.wml import TcPr

    if source is None or is_empty_tc_pr(source):
        return destination
    out = TcPr() if destination is None else destination
    apply_catalogue(CELL, source, out)
    return out


def apply_tr_pr(source: Any, destination: Any) -> Any:
    """Apply a ``w:trPr`` over a ``w:trPr``.

    The generated model keeps a row's properties as one element list: an element
    of a class the destination already has replaces it, and one it lacks is
    added. Each generated class in that list is one element name, so the class is
    the name.
    """
    from docx4j_py.wml import TrPr

    if is_empty_tr_pr(source):
        return destination
    out = TrPr() if destination is None else destination
    for element in source.content:
        for index in range(len(out.content) - 1, -1, -1):
            if type(out.content[index]) is type(element):
                del out.content[index]
        out.content.append(copy_leaf(element))
    return out


def apply_tbl_style_pr_list(source: Any, destination: Any) -> None:
    """Merge a table style's conditional formatting (``w:tblStylePr``) into another's.

    *Per condition*: an entry for a condition the destination already has is
    applied on top of it, and one it lacks is added. So a style ``w:basedOn``
    another inherits the conditions it does not restate, and restating a
    condition overrides only the properties it names.
    """
    if is_empty_tbl_style_pr_list(source) or destination is None:
        return
    for source_pr in source:
        if source_pr is None:
            continue
        existing = None
        if source_pr.type_value is not None:
            existing = next(
                (
                    d
                    for d in destination
                    if d is not None and d.type_value == source_pr.type_value
                ),
                None,
            )
        if existing is None:
            copy = apply_tbl_style_pr(source_pr, None)
            if copy is not None:
                destination.append(copy)
        else:
            apply_tbl_style_pr(source_pr, existing)


def apply_tbl_style_pr(source: Any, destination: Any) -> Any:
    """Apply one ``w:tblStylePr`` over another. docx4j ``apply(CTTblStylePr, CTTblStylePr)``."""
    from docx4j_py.wml import CTTblStylePr, RPr

    if is_empty_tbl_style_pr(source):
        return destination
    out = CTTblStylePr() if destination is None else destination
    out.p_pr = apply_ppr(source.p_pr, out.p_pr)
    out.r_pr = apply_rpr(source.r_pr, out.r_pr, RPr)
    out.tbl_pr = apply_tbl_pr(source.tbl_pr, out.tbl_pr)
    out.tc_pr = apply_tc_pr(source.tc_pr, out.tc_pr)
    out.tr_pr = apply_tr_pr(source.tr_pr, out.tr_pr)
    out.type_value = source.type_value
    return out


def apply_style(source: Any, destination: Any) -> Any:
    """Apply one style's own properties over another's. docx4j ``apply(Style, Style)``.

    It does not climb the ``w:basedOn`` chain;
    :meth:`~docx4j_py.model.property_resolver.PropertyResolver.ancestry` does
    that.
    """
    from docx4j_py.wml import RPr

    if is_empty_style(source):
        return destination
    if destination.type_value is None:
        destination.type_value = source.type_value
    kind = _style_type(source)
    if kind == CHARACTER_STYLE:
        destination.r_pr = apply_rpr(source.r_pr, destination.r_pr, RPr)
        return destination
    if kind in (PARAGRAPH_STYLE, NUMBERING_STYLE):
        destination.p_pr = apply_ppr(source.p_pr, destination.p_pr)
        destination.r_pr = apply_rpr(source.r_pr, destination.r_pr, RPr)
        return destination
    destination.tbl_pr = apply_tbl_pr(source.tbl_pr, destination.tbl_pr)
    destination.tr_pr = apply_tr_pr(source.tr_pr, destination.tr_pr)
    destination.tc_pr = apply_tc_pr(source.tc_pr, destination.tc_pr)
    apply_tbl_style_pr_list(source.tbl_style_pr, destination.tbl_style_pr)
    destination.p_pr = apply_ppr(source.p_pr, destination.p_pr)
    destination.r_pr = apply_rpr(source.r_pr, destination.r_pr, RPr)
    return destination


# ---------------------------------------------------------------- cycles

#: docx4j's hard limit on deep ``w:basedOn`` hierarchies.
CYCLE_LIMIT = 32

_throw_on_cyclic = False


def is_cyclic(style_id: str, seen: list[str]) -> bool:
    """Whether the chain has come back to a style it has already seen.

    Or has gone deeper than :data:`CYCLE_LIMIT`. The caller degrades gracefully
    by using as much of the hierarchy as it has; :func:`throw_on_cyclic_styles`
    makes it raise instead, which is docx4j's
    ``docx4j.openpackaging.exceptions.CyclicStylesException.throw`` property.
    """
    if style_id not in seen and len(seen) <= CYCLE_LIMIT:
        return False
    logger.warning("Cycle detected in style basedOn hierarchy for: %s - stopping", style_id)
    if _throw_on_cyclic:
        from docx4j_py.openpackaging.exceptions import CyclicStylesException

        raise CyclicStylesException(f"Cycle detected in style basedOn hierarchy for: {style_id}")
    return True


def throw_on_cyclic_styles(value: bool) -> None:
    """Turn raising on a style cycle on or off. docx4j's property, off by default."""
    global _throw_on_cyclic  # docx4j's property is process wide too
    _throw_on_cyclic = value


def throws_on_cyclic_styles() -> bool:
    """Whether a style cycle raises."""
    return _throw_on_cyclic
