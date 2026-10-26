"""Every write of an integer measurement goes through the check: CR-022 section 4.3 (D3), the census.

Word keeps each integer attribute in a field of its own width and reads a value past it
wrapped (CR-022 sections 3.1 and 3.3), so every setter that writes one goes through
:func:`docx4j_py.model.values.measure_in_range`, which refuses a value Word would not read as
written, before anything changes.

**The census.** Every write of a WordprocessingML integer attribute under
``docx4j_py/model/content``, ``docx4j_py/docx`` and the hand-written modules of
``docx4j_py/wml`` (:data:`MODULES`) either goes through the check or is in :data:`LISTED`
with the reason it need not. A write is:

* an assignment ``<object>.<field> = <value>`` where ``<field>`` is the Python name of an
  integer attribute of ``w``, ``w14`` or ``w15`` anywhere in the model (:func:`integer_fields`),
  ``<object>`` is not bare ``self`` (a view's own state) and ``<value>`` not a constant;
* ``setattr(<object>, <name>, <value>)``, its name such a field or not a constant, and its
  value not a constant;
* a keyword argument ``<field>=<value>`` to a model class or an ``el`` factory, where the
  class it makes types ``<field>`` an integer attribute of ``w``, ``w14`` or ``w15``, and the
  value not a constant.

A write **goes through the check** when its value holds a call of ``measure_in_range``, or a
name the same function bound from one (directly or through another such name). The key is
``<file>::<scope>::<the write as written>``, so one listed write covers the same write
repeated in the same function. :func:`test_every_measurement_write_is_checked_or_listed`
keeps the list whole, the way ``tests/content/test_integer_readers.py`` keeps CR-021's: a
new write fails it until it goes through ``measure_in_range`` or is listed here with its
reason.
"""

from __future__ import annotations

import ast
import inspect
from functools import cache
from importlib import import_module
from pathlib import Path
from typing import Any

import docx4j_py

#: The package imported: the checkout's, or the installed wheel's under scripts/release_check.sh.
PACKAGE = Path(docx4j_py.__file__).parent

#: The directories read whole, and the hand-written modules of ``docx4j_py/wml``.
DIRECTORIES = ("model/content", "docx")
MODULES = ("wml/builders.py", "wml/pictures.py", "wml/sdt.py")

CHECK = "measure_in_range"

NOT_A_MEASUREMENT = (
    "an integer attribute that is no measurement: an id, a count, a level, a list's numId or "
    "abstractNumId, a level's start, a style's priority, a grid span; ids and numIds are 32 bits "
    "or wider, which a 16-bit row would refuse (CR-022 section 3)"
)
NOT_A_FIELD = "not a model field: a span's or a range's own offsets, or a report's"
COPY = (
    "a field copied as it stands from another model object (a deep copy, a split run, a restored "
    "w:pPr): the value was already in the tree, which is CR-020 Phase B's to report (CR-022 D5)"
)
NO_ROW = (
    "w:gridCol/@w:w, which Word recomputes from the cells (CR-022 section 3.3): no row, so not "
    "checked (Q3); the w:tcW written beside it is"
)
COMPUTED = (
    "a width this code computes from values the caller's check already held, or from the page "
    "and the grid, not taken from a caller"
)
CONSTANT_EXPR = "a constant expression: a level's default indents, a line multiple's 240ths"
STRING = "a string value written to a w:val that is no integer measurement: a compatibility setting, an alignment"
NAME_ONLY = (
    "a field that is no integer attribute here, only named like one elsewhere in the model: an element "
    "(w:b, a toggle, a child, a content list), xml:space, a relationship id, a core property"
)

#: Every write of an integer attribute that does not go through ``measure_in_range``, with why.
LISTED: dict[str, str] = {
    # the python-docx facade
    "docx/coreprops.py::CoreProperties._set_date::setattr(properties, field, ...)": NAME_ONLY,
    "docx/coreprops.py::CoreProperties._set_text::setattr(properties, field, ...)": NAME_ONLY,
    "docx/styles.py::BaseStyle._set_toggle::setattr(self._style, field, ...)": NAME_ONLY,
    "docx/styles.py::BaseStyle.priority::el.uiPriority(val=...)": NOT_A_MEASUREMENT,
    "docx/styles.py::BaseStyle.priority::self._style.ui_priority =": NOT_A_MEASUREMENT,
    "docx/table.py::Table._set_column_widths::column.w =": NO_ROW,
    "docx/table.py::Table._set_column_widths::el.tcW(w=...)": COMPUTED,
    "docx/table.py::_Column.width::self._grid_col.w =": NO_ROW,
    "docx/text.py::Font._set_bool::setattr(r_pr, each, ...)": NAME_ONLY,
    "docx/text.py::ParagraphFormat._child::setattr(p_pr, name, ...)": NAME_ONLY,
    "docx/text.py::ParagraphFormat._set_toggle::setattr(p_pr, field, ...)": NAME_ONLY,
    "docx/text.py::ParagraphFormat.line_spacing_rule::spacing.line =": CONSTANT_EXPR,
    "docx/text.py::TabStop.alignment::self._element.val =": STRING,
    # the content API
    "model/content/__init__.py::register::setattr(WordprocessingMLPackage, name, ...)": NOT_A_FIELD,
    "model/content/body.py::Body.content::setattr(owner, field, ...)": NAME_ONLY,
    "model/content/comments.py::_add_comment::el.comment(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_append::setattr(owner, field, ...)": NAME_ONLY,
    "model/content/comments.py::_place_across::el.commentRangeEnd(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_place_across::el.commentRangeStart(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_place_after::el.commentRangeEnd(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_place_after::el.commentRangeStart(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_place_around::el.commentRangeEnd(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_place_around::el.commentRangeStart(id=...)": NOT_A_MEASUREMENT,
    "model/content/comments.py::_reference_run::el.commentReference(id=...)": NOT_A_MEASUREMENT,
    "model/content/compatibility.py::set_mode::setting.val =": STRING,
    "model/content/controls.py::ContentControl.insert_copy_after::el.id_(val=...)": NOT_A_MEASUREMENT,
    "model/content/hyperlinks.py::_revision_like::clone.id =": NOT_A_MEASUREMENT,
    "model/content/hyperlinks.py::_wrap::link.id =": NAME_ONLY,
    "model/content/hyperlinks.py::_wrap::span.end =": NOT_A_FIELD,
    "model/content/hyperlinks.py::_wrap::span.start =": NOT_A_FIELD,
    "model/content/lists.py::List._writable_level::el.lvlOverride(ilvl=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::List.set_level_starting_number::CtLvlStart(val=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::List.set_level_starting_number::lvl.start =": NOT_A_MEASUREMENT,
    "model/content/lists.py::ListItem.level::el.ilvl(val=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::ListItem.level::num_pr.ilvl =": NOT_A_MEASUREMENT,
    "model/content/lists.py::_append::setattr(numbering, field, ...)": NAME_ONLY,
    "model/content/lists.py::_blank_level::el.ind(hanging=...)": CONSTANT_EXPR,
    "model/content/lists.py::_blank_level::el.ind(left=...)": CONSTANT_EXPR,
    "model/content/lists.py::_blank_level::el.lvl(ilvl=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::_set_num_pr::el.ilvl(val=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::_set_num_pr::el.numId(val=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::add_abstract_definition::abstract.abstract_num_id =": NOT_A_MEASUREMENT,
    "model/content/lists.py::add_abstract_definition::el.abstractNumId(val=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::add_abstract_definition::el.num(num_id=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::restart_over::el.abstractNumId(val=...)": NOT_A_MEASUREMENT,
    "model/content/lists.py::restart_over::el.num(num_id=...)": NOT_A_MEASUREMENT,
    "model/content/ooxml.py::_remove_keeping_mark::span.end =": NOT_A_FIELD,
    "model/content/ooxml.py::_visit_references.typed::setattr(node, field, ...)": NAME_ONLY,
    "model/content/paragraph.py::Paragraph.outline_level::el.outlineLvl(val=...)": NOT_A_MEASUREMENT,
    "model/content/paragraph.py::_carry_attributes::setattr(made, name, ...)": COPY,
    "model/content/span_edits.py::look_of::setattr(out, name, ...)": COPY,
    "model/content/split.py::_holder_like::clone.id =": NOT_A_MEASUREMENT,
    "model/content/styles.py::_reconcile_style_numbering::num_pr.num_id.val =": NOT_A_MEASUREMENT,
    "model/content/text_model.py::split_at::setattr(second, name, ...)": COPY,
    "model/content/tracking.py::_delete_math_structure::setattr(structure, key, ...)": NAME_ONLY,
    "model/content/tracking.py::_rename::made.space =": NAME_ONLY,
    "model/content/tracking.py::restore_p_pr::setattr(fresh, field.name, ...)": COPY,
    # the builders
    "wml/builders.py::apply_run_options::r_pr.b =": NAME_ONLY,
    "wml/builders.py::rpr_from_elements::setattr(target, field_name, ...)": COPY,
    "wml/builders.py::tbl::el.gridCol(w=...)": NO_ROW,
    "wml/builders.py::tc::el.gridSpan(val=...)": NOT_A_MEASUREMENT,
    "wml/sdt.py::sdt_pr::el.id_(val=...)": NOT_A_MEASUREMENT,
}


# ---------------------------------------------------------------------------
# the model's integer attributes
# ---------------------------------------------------------------------------


@cache
def integer_fields() -> frozenset[str]:
    """The Python names of every integer attribute of ``w``, ``w14`` and ``w15`` in the model."""
    from docx4j_py.openpackaging.normalise import _bound, integer_attributes

    bound, _holders, broken = _bound()
    return frozenset(
        var.name
        for classes in bound.values()
        for cls in classes
        if cls not in broken
        for var, word in integer_attributes(cls)
        if word
    )


@cache
def _integer_fields_of(cls: type) -> frozenset[str]:
    from docx4j_py.openpackaging.normalise import integer_attributes

    return frozenset(var.name for var, word in integer_attributes(cls) if word)


@cache
def _model_classes() -> frozenset[type]:
    """Every class of the generated model: each one an element name binds."""
    from docx4j_py.openpackaging.normalise import _bound

    bound, _holders, broken = _bound()
    return frozenset(cls for classes in bound.values() for cls in classes if cls not in broken)


def _model_class(obj: Any) -> type | None:
    """The model class `obj` makes: itself, or what an ``el`` factory returns; None otherwise."""
    if inspect.isclass(obj):
        return obj if obj in _model_classes() else None
    module = getattr(obj, "__module__", "") or ""
    if callable(obj) and module.startswith("docx4j_py.") and module.endswith(".el"):
        try:
            made = obj()
        except TypeError:
            return None
        return type(made) if type(made) in _model_classes() else None
    return None


def _resolve(module: Any, func: ast.expr) -> type | None:
    """The model class a call's function makes, read from the scanned module or ``docx4j_py.wml``."""
    from docx4j_py import wml

    if isinstance(func, ast.Name):
        for namespace in (module, wml):
            if hasattr(namespace, func.id):
                return _model_class(getattr(namespace, func.id))
        return None
    if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name):
        for namespace in (module, wml):
            owner = getattr(namespace, func.value.id, None)
            if owner is not None and hasattr(owner, func.attr):
                return _model_class(getattr(owner, func.attr))
    return None


# ---------------------------------------------------------------------------
# the scanner
# ---------------------------------------------------------------------------


def _calls_check(node: ast.AST, checked: set[str]) -> bool:
    """Whether `node` holds a call of ``measure_in_range``, or a name bound from one."""
    for each in ast.walk(node):
        if isinstance(each, ast.Call):
            func = each.func
            name = func.id if isinstance(func, ast.Name) else func.attr if isinstance(func, ast.Attribute) else None
            if name == CHECK:
                return True
        if isinstance(each, ast.Name) and each.id in checked:
            return True
    return False


def _bindings(function: ast.AST) -> list[tuple[list[str], ast.AST]]:
    """Each (names, value) a function binds: assignments, loop and comprehension targets, ``with ... as``."""
    out: list[tuple[list[str], ast.AST]] = []

    def names(target: ast.AST) -> list[str]:
        return [n.id for n in ast.walk(target) if isinstance(n, ast.Name)]

    def visit(node: ast.AST) -> None:
        for child in ast.iter_child_nodes(node):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Lambda)):
                continue
            if isinstance(child, ast.Assign):
                out.append(([n for t in child.targets for n in names(t)], child.value))
            elif isinstance(child, (ast.AnnAssign, ast.AugAssign, ast.NamedExpr)) and child.value is not None:
                out.append((names(child.target), child.value))
            elif isinstance(child, (ast.For, ast.AsyncFor, ast.comprehension)):
                out.append((names(child.target), child.iter))
            elif isinstance(child, ast.withitem) and child.optional_vars is not None:
                out.append((names(child.optional_vars), child.context_expr))
            visit(child)

    visit(function)
    return out


def _checked_names(function: ast.AST) -> set[str]:
    """The names `function` binds from a value that goes through the check, to a fixed point."""
    bindings = _bindings(function)
    checked: set[str] = set()
    while True:
        more = {name for names, value in bindings if _calls_check(value, checked) for name in names}
        if more <= checked:
            return checked
        checked |= more


def writes_in(tree: ast.AST, relative: str, module: Any) -> dict[str, bool]:
    """Every write of an integer attribute in one module's `tree`, and whether it goes through the check."""
    fields = integer_fields()
    found: dict[str, bool] = {}

    def record(key: str, value: ast.AST, checked: set[str]) -> None:
        found[key] = found.get(key, True) and _calls_check(value, checked)

    def visit(node: ast.AST, scope: list[str], checked: set[str]) -> None:
        for child in ast.iter_child_nodes(node):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                visit(child, [*scope, child.name], checked | _checked_names(child))
                continue
            if isinstance(child, ast.ClassDef):
                visit(child, [*scope, child.name], set())
                continue
            where = f"{relative}::{'.'.join(scope) or '<module>'}"
            if isinstance(child, (ast.Assign, ast.AnnAssign)) and child.value is not None:
                targets = child.targets if isinstance(child, ast.Assign) else [child.target]
                for target in targets:
                    if (
                        isinstance(target, ast.Attribute)
                        and target.attr in fields
                        and not (isinstance(target.value, ast.Name) and target.value.id == "self")
                        and not isinstance(child.value, ast.Constant)
                    ):
                        record(f"{where}::{ast.unparse(target)} =", child.value, checked)
            if isinstance(child, ast.Call):
                func = child.func
                if isinstance(func, ast.Name) and func.id == "setattr" and len(child.args) == 3:
                    owner, name, value = child.args
                    named = isinstance(name, ast.Constant) and isinstance(name.value, str)
                    if (not named or name.value in fields) and not isinstance(value, ast.Constant):
                        record(f"{where}::setattr({ast.unparse(owner)}, {ast.unparse(name)}, ...)", value, checked)
                else:
                    made = None
                    for keyword in child.keywords:
                        if keyword.arg not in fields or isinstance(keyword.value, ast.Constant):
                            continue
                        made = made or _resolve(module, func)
                        if made is not None and keyword.arg in _integer_fields_of(made):
                            record(f"{where}::{ast.unparse(func)}({keyword.arg}=...)", keyword.value, checked)
            visit(child, scope, checked)

    visit(tree, [], set())
    return found


def writes_in_source() -> dict[str, bool]:
    """Every write of an integer attribute in the census's files, and whether it goes through the check."""
    paths = [path for root in DIRECTORIES for path in sorted((PACKAGE / root).rglob("*.py"))]
    paths += [PACKAGE / module for module in MODULES]
    found: dict[str, bool] = {}
    for path in paths:
        relative = path.relative_to(PACKAGE).as_posix()
        name = "docx4j_py." + relative.removesuffix(".py").replace("/", ".").removesuffix(".__init__")
        found.update(writes_in(ast.parse(path.read_text(encoding="utf-8")), relative, import_module(name)))
    return found


# ---------------------------------------------------------------------------
# the tests
# ---------------------------------------------------------------------------


def test_every_measurement_write_is_checked_or_listed() -> None:
    """A new write of an integer attribute fails this until it goes through ``measure_in_range`` or is listed."""
    found = writes_in_source()
    unchecked = {key for key, ok in found.items() if not ok}
    assert sorted(unchecked - set(LISTED)) == [], "write it through measure_in_range, or list it with its reason"
    assert sorted(set(LISTED) - unchecked) == [], "listed, but no longer an unchecked write: take it out"


def test_every_listed_write_has_a_reason() -> None:
    reasons = {NOT_A_MEASUREMENT, NOT_A_FIELD, COPY, NO_ROW, COMPUTED, CONSTANT_EXPR, STRING, NAME_ONLY}
    assert set(LISTED.values()) <= reasons


def test_the_setters_cr_022_names_go_through_the_check() -> None:
    """The writes of CR-022 section 2.2, each through ``measure_in_range``: the census found them all."""
    checked = {key for key, ok in writes_in_source().items() if ok}
    for prefix in (
        "model/content/paragraph.py::Paragraph._set_indent::",
        "model/content/paragraph.py::Paragraph.first_line_indent::",
        "model/content/paragraph.py::Paragraph.space_before::",
        "model/content/paragraph.py::Paragraph.space_after::",
        "model/content/paragraph.py::Paragraph.line_spacing::",
        "model/content/table.py::TableCell.width::",
        "model/content/lists.py::List.set_level_indents::",
        "docx/section.py::Section.page_width::",
        "docx/section.py::Section.page_height::",
        "docx/section.py::Section._set_margin::",
        "docx/text.py::Font.size::",
        "docx/text.py::ParagraphFormat._set_ind::",
        "docx/text.py::ParagraphFormat.first_line_indent::",
        "docx/text.py::ParagraphFormat._set_spacing::",
        "docx/text.py::ParagraphFormat.line_spacing::",
        "docx/text.py::TabStops.add_tab_stop::",
        "docx/text.py::TabStop.position::",
        "docx/table.py::Table.add_column::",
        "docx/table.py::_Row.height::",
        "docx/table.py::_Cell.width::",
        "wml/builders.py::apply_run_options::",
        "wml/builders.py::tc::",
        "wml/builders.py::tbl::",
    ):
        assert any(key.startswith(prefix) for key in checked), prefix


def test_the_census_sees_through_a_bound_name() -> None:
    """The scanner's own rule: a value bound from ``measure_in_range``, directly or not, is checked."""
    tree = ast.parse(
        "def f(x):\n"
        "    a = measure_in_range('w:ind/@w:left', x, label='l')\n"
        "    b = a\n"
        "    c = round(x * 20)\n"
        "    for d in (a, b):\n"
        "        pass\n"
    )
    function = tree.body[0]
    assert _checked_names(function) == {"a", "b", "d"}
    assert not _calls_check(ast.parse("round(x * 20)").body[0], {"a"})
    assert _calls_check(ast.parse("values.measure_in_range(1)").body[0], set())


def test_the_census_catches_a_write_that_skips_the_check() -> None:
    """A raw conversion written to an integer field, to an ``el`` factory's keyword or by ``setattr`` is found."""
    from docx4j_py.model.content import paragraph

    tree = ast.parse(
        "def raw(ind, points):\n"
        "    ind.left = round(points * 20)\n"
        "    el.ind(hanging=round(points * 20))\n"
        "    setattr(ind, 'right', round(points * 20))\n"
        "    el.jc(val=points)\n"
        "def checked(ind, points):\n"
        "    twips = measure_in_range('w:ind/@w:left', points, label='left indent')\n"
        "    ind.left = twips\n"
        "    el.ind(left=twips)\n"
    )
    found = writes_in(tree, "x.py", paragraph)
    assert found == {
        "x.py::raw::ind.left =": False,
        "x.py::raw::el.ind(hanging=...)": False,
        "x.py::raw::setattr(ind, 'right', ...)": False,
        "x.py::checked::ind.left =": True,
        "x.py::checked::el.ind(left=...)": True,
    }, "w:jc/@w:val is an enumeration, not an integer attribute: not a write here"
