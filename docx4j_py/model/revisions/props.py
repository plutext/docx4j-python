"""Recorded properties back into live ones. docx4j ``Props``.

CR-043 section 3.4: the records are not the live types. A ``w:rPrChange/w:rPr``
is ``CtRprChangeRPr`` (every ``EG_RPrBase`` member a list), a mark's is
``CTParaRPrOriginal``, a ``w:pPrChange/w:pPr`` is ``PPrBase``. docx4j goes
through their XML; here the generated classes share their field names, so a
record is copied member by member into the live class
(:func:`~docx4j_py.wml.rpr_from_elements` for run properties, as Phase F did).

Check 25's list rules are here (:func:`restore_p_pr`): a restored ``w:ind``
equal to the list level's own is dropped, and a ``w:numPr`` naming a list with
no ``w:ilvl`` is given ``w:ilvl`` 0.
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.model.revisions import trees

__all__ = [
    "prune_paragraph",
    "prune_run",
    "r_pr_of",
    "restore_mark_properties",
    "restore_p_pr",
    "restore_r_pr",
]


def r_pr_of(recorded: Any) -> Any:
    """A recorded ``w:rPr`` (``CtRprChangeRPr``) as a live ``w:rPr``; None when it records none."""
    if recorded is None:
        return None
    from docx4j_py.wml import RPr, rpr_from_elements

    live = rpr_from_elements(recorded, cls=RPr)
    return None if trees.is_empty(live) else live


def restore_r_pr(run: Any, record: Any) -> None:
    """A run's direct formatting replaced by what its ``w:rPrChange`` recorded, wholesale (checks 24, 28)."""
    run.r_pr = r_pr_of(getattr(record, "r_pr", None))


def restore_mark_properties(p: Any, record: Any) -> None:
    """A mark's formatting put back as its ``w:rPrChange`` recorded it, its own revisions kept (check 27)."""
    p_pr = p.p_pr
    if p_pr is None or p_pr.r_pr is None:
        return
    from docx4j_py.wml import ParaRPr, rpr_from_elements

    current = p_pr.r_pr
    recorded = getattr(record, "r_pr", None)
    restored = ParaRPr() if recorded is None else rpr_from_elements(recorded, cls=ParaRPr)
    restored.ins = current.ins
    restored.del_value = current.del_value
    restored.move_from = current.move_from
    restored.move_to = current.move_to
    restored.r_pr_change = None
    p_pr.r_pr = restored
    from docx4j_py.child import link_parents

    link_parents(restored)
    restored.parent = p_pr
    prune_paragraph(p)


def restore_p_pr(p: Any, recorded: Any, package: Any) -> None:
    """A paragraph's properties replaced by what its ``w:pPrChange`` recorded.

    The mark's run properties and the section stay (the record is a
    ``CT_PPrBase``, which has neither); then check 25's two list rules, and the
    husk is pruned.
    """
    from docx4j_py.child import deep_copy, link_parents
    from docx4j_py.wml import el

    current = p.p_pr
    if current is None:
        return
    keep_r_pr, keep_sect_pr = current.r_pr, current.sect_pr
    fresh = el.pPr()
    if recorded is not None:
        for field in dataclasses.fields(recorded):
            value = getattr(recorded, field.name, None)
            if value is None or (isinstance(value, list) and not value):
                continue
            setattr(fresh, field.name, deep_copy(value))
    fresh.r_pr = keep_r_pr
    fresh.sect_pr = keep_sect_pr
    fresh.p_pr_change = None
    p.p_pr = fresh
    link_parents(fresh)
    fresh.parent = p
    _drop_level_indent(fresh, package)
    _write_recorded_level(fresh)
    prune_paragraph(p)


def prune_paragraph(p: Any) -> None:
    """Drops ``w:pPr/w:rPr`` and ``w:pPr`` once nothing is left in them, so that resolving leaves no husk."""
    p_pr = p.p_pr
    if p_pr is None:
        return
    if p_pr.r_pr is not None and trees.is_empty(p_pr.r_pr):
        p_pr.r_pr = None
    if trees.is_empty(p_pr):
        p.p_pr = None


def prune_run(run: Any) -> None:
    """Drops a run's ``w:rPr`` once nothing is left in it."""
    if run.r_pr is not None and trees.is_empty(run.r_pr):
        run.r_pr = None


# ---------------------------------------------------------------------------
# check 25's list rules
# ---------------------------------------------------------------------------


def _drop_level_indent(p_pr: Any, package: Any) -> None:
    """A restored ``w:ind`` whose every attribute is the list level's is dropped.

    Word records a list paragraph's indent in ``w:pPrChange`` even where the
    list level gives it, and on reject writes no indent back (check 18 section
    9, check 25).
    """
    if p_pr.num_pr is None or p_pr.ind is None:
        return
    main = getattr(package, "main_document_part", None) if package is not None else None
    numbering = getattr(main, "numbering_definitions_part", None) if main is not None else None
    if numbering is None:
        return
    try:
        emulator = numbering.get_emulator()
        level = emulator.get_ind(p_pr.num_pr) if emulator is not None else None
    except Exception:  # noqa: BLE001 - a numbering part we cannot read decides nothing
        return
    if level is None or not _indent_within(p_pr.ind, level):
        return
    p_pr.ind = None


def _indent_within(ind: Any, level: Any) -> bool:
    """True when every attribute `ind` sets, `level` sets to the same value."""
    found = False
    for field in dataclasses.fields(ind):
        mine = getattr(ind, field.name, None)
        if mine is None:
            continue
        found = True
        theirs = getattr(level, field.name, None)
        if theirs is None or str(mine) != str(theirs):
            return False
    return found


def _write_recorded_level(p_pr: Any) -> None:
    """A restored ``w:numPr`` that names a list but no level is at level 0, and Word writes the 0.

    Check 25's ``25e``, check 18 section 9.
    """
    num_pr = p_pr.num_pr
    if num_pr is None or num_pr.num_id is None or num_pr.ilvl is not None:
        return
    value = getattr(num_pr.num_id, "val", None)
    if value is None or str(value) == "0":
        return
    from docx4j_py.wml import el

    ilvl = el.ilvl(val=0)
    num_pr.ilvl = ilvl
    ilvl.parent = num_pr
