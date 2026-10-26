"""Every reader of an integer attribute reads it as Word does: CR-021 section 4.2 (D2), the census.

A WordprocessingML integer attribute the parse could not bind is kept as its text
(``"720.5"``, ``""``: CR-021 section 2.1), so a reader that reads such a field with
``int()`` raises and one that reads it with ``float()`` disagrees with Word, which
truncates (``"720.5"`` is 720 twips, ``""`` 0). Every reader of one goes through
:func:`docx4j_py.model.values.integer`.

**The census.** Every call of ``int(...)`` or ``float(...)`` under ``docx4j_py/model``
and ``docx4j_py/docx``, and in the hand-written modules beside the generated model
that read it (:data:`MODULES`: ``wml/builders.py``, ``wml/pictures.py``, ``wml/sdt.py``,
``traversal.py``), either goes through ``values.integer`` (its argument holds a call
of ``integer``) or is in :data:`LISTED` with the reason it is not a reader of a
WordprocessingML integer attribute. :func:`test_every_int_and_float_is_listed` keeps
the list whole, the way ``tests/openpackaging/test_private_readers.py`` keeps
``read_contents``' callers: a new ``int(`` or ``float(`` fails it until it goes
through ``values.integer`` or is listed here with its reason. The key is
``<file>::<scope>::<the call as written>``, so one listed call covers the same
expression repeated in the same function.

The readers that go through ``values.integer`` (CR-021 section 10 has the list):
the content API's twips, indents, spacing, outline levels, widths, ids and page
size (``paragraph.py``, ``api_script.py``, ``table.py``, ``picture.py``,
``describe.py``, ``comments.py``, ``controls.py``, ``reports.py``, ``lists.py``);
the python-docx facade's (``docx/section.py``, ``table.py``, ``text.py``,
``styles.py``); the numbering emulator's indents and its ``w:numId``, ``w:ilvl``
and ``w:abstractNumId`` keys (``listnumbering/definitions.py``'s ``_int`` and
``decimal``, ``emulator.py``'s ``_as_int``); the revisions engine's widths, spans
and marker ids (``revisions/tables.py``, ``trees.py``, ``removals.py``); and the
text model's (``text/extractor.py``, ``markdown.py``, ``view.py``); and
``read_run_options``' size (``wml/builders.py``), under ``Font`` and ``api_script``.
"""

from __future__ import annotations

import ast
from pathlib import Path

import docx4j_py

#: The package imported: the checkout's, or the installed wheel's under scripts/release_check.sh.
PACKAGE = Path(docx4j_py.__file__).parent

#: The directories read whole, and the hand-written modules elsewhere that read the model.
DIRECTORIES = ("model", "docx")
MODULES = ("wml/builders.py", "wml/pictures.py", "wml/sdt.py", "traversal.py")

ARGUMENT = "a caller's argument (a level, an id, a length, a code point), not a model field"
LENGTH = "a python-docx Length or a number the facade computed, not a model field"
HEX = "hexadecimal text (ST_ShortHexNumber, ST_HexColor, a colour component), not an integer attribute"
TEXT = "text that is no integer attribute: a style id, an address, a field switch, Markdown, lvlText"
CONSTANT = "a constant"
NUMBER = "a number this code computed (a rounded float, a chart value, a date), not a model field"
OUTSIDE = (
    "an integer attribute outside w, w14 and w15 (DrawingML, charts, SmartArt), read with Python's int() as "
    "before: Word was measured in WordprocessingML only (CR-021 D5, Q3)"
)
STRING = "an xsd:string value read as a number by its own rule (w:compatSetting/@w:val, cp:revision)"
MODE = "the compatibility mode, an int the package already holds or a member of CompatibilityMode"
KEY = "the numbering emulator's numId or ilvl key, a decimal definitions.decimal wrote through values.integer"
OWN = "values.integer itself: ASCII digits its expression matched"
TABLE = "the CR-022 limits table's own numbers (resources/integer-limits.json), not a model field"

#: Every ``int(`` and ``float(`` under ``model/`` and ``docx/`` that does not go through ``values.integer``.
LISTED: dict[str, str] = {
    # the python-docx facade
    "docx/coreprops.py::CoreProperties.revision::int(text)": STRING,
    "docx/shape.py::InlineShape.__repr__::int(self.height)": LENGTH,
    "docx/shape.py::InlineShape.__repr__::int(self.width)": LENGTH,
    "docx/shape.py::InlineShape.height::int(cy)": ARGUMENT,
    "docx/shape.py::InlineShape.height::int(self._inline.extent.cy)": OUTSIDE,
    "docx/shape.py::InlineShape.width::int(cx)": ARGUMENT,
    "docx/shape.py::InlineShape.width::int(self._inline.extent.cx)": OUTSIDE,
    "docx/shared.py::Cm.__new__::int(cm * Length._EMUS_PER_CM)": ARGUMENT,
    "docx/shared.py::Emu.__new__::int(emu)": ARGUMENT,
    "docx/shared.py::Inches.__new__::int(inches * Length._EMUS_PER_INCH)": ARGUMENT,
    "docx/shared.py::Length.cm::float(self._EMUS_PER_CM)": CONSTANT,
    "docx/shared.py::Length.inches::float(self._EMUS_PER_INCH)": CONSTANT,
    "docx/shared.py::Length.mm::float(self._EMUS_PER_MM)": CONSTANT,
    "docx/shared.py::Length.pt::float(self._EMUS_PER_PT)": CONSTANT,
    "docx/shared.py::Length.twips::float(self._EMUS_PER_TWIP)": CONSTANT,
    "docx/shared.py::Mm.__new__::int(mm * Length._EMUS_PER_MM)": ARGUMENT,
    "docx/shared.py::Pt.__new__::int(points * Length._EMUS_PER_PT)": ARGUMENT,
    "docx/shared.py::RGBColor.from_string::int(text[2:4], 16)": HEX,
    "docx/shared.py::RGBColor.from_string::int(text[4:6], 16)": HEX,
    "docx/shared.py::RGBColor.from_string::int(text[:2], 16)": HEX,
    "docx/shared.py::Twips.__new__::int(twips * Length._EMUS_PER_TWIP)": ARGUMENT,
    "docx/shared.py::length_or_none::int(value)": ARGUMENT,
    "docx/shared.py::points_of::int(value)": ARGUMENT,
    "docx/shared.py::twips_of::int(value)": ARGUMENT,
    "docx/styles.py::BaseStyle.priority::int(value)": ARGUMENT,
    "docx/table.py::Table.table_direction::int(WD_TABLE_DIRECTION.RTL)": CONSTANT,
    "docx/table.py::Table.table_direction::int(value)": ARGUMENT,
    "docx/text.py::ParagraphFormat.first_line_indent::int(value)": LENGTH,
    "docx/text.py::ParagraphFormat.line_spacing::int(line)": LENGTH,
    "docx/text.py::ParagraphFormat.line_spacing_rule::int(Twips(twips))": CONSTANT,
    "docx/text.py::ParagraphFormat.line_spacing_rule::int(line)": LENGTH,
    # charts (c:, cx:), outside WordprocessingML
    "model/chart/chartex.py::ChartExModel.read::int(a.get('id', '0'))": OUTSIDE,
    "model/chart/chartex.py::ChartExModel.read::int(d.get('id', '0'))": OUTSIDE,
    "model/chart/chartex.py::ChartExModel.read::int(data_id.get('val'))": OUTSIDE,
    "model/chart/chartex.py::_dimensions::float(text)": NUMBER,
    "model/chart/chartex.py::_dimensions::int(item.get('ptCount', '0'))": OUTSIDE,
    "model/chart/chartex.py::_dimensions::int(pt.get('idx', '0'))": OUTSIDE,
    "model/chart/chartex.py::_dimensions::int(x)": NUMBER,
    "model/chart/chartex.py::java_double::int(exponent)": NUMBER,
    "model/chart/formats.py::_excel_date::int((value - whole) * 86400000 + 0.5)": NUMBER,
    "model/chart/formats.py::_general::float(_TO_10_SF.plus(Decimal(value)))": NUMBER,
    "model/chart/formats.py::format_raw_cell_contents::float(_number_text(value))": NUMBER,
    "model/chart/model.py::ChartAxis.of::int(ax_id)": OUTSIDE,
    "model/chart/model.py::ChartAxis.of::int(cross)": OUTSIDE,
    "model/chart/model.py::ChartData._numbers::int(pt.idx)": OUTSIDE,
    "model/chart/model.py::ChartGroup.of::int(_val(a))": OUTSIDE,
    "model/chart/model.py::ChartSeries.of::int(idx)": OUTSIDE,
    "model/chart/model.py::ChartSeries.of::int(order)": OUTSIDE,
    "model/chart/model.py::_points::int(pt.idx)": OUTSIDE,
    "model/chart/table.py::format_number::float(text.strip())": NUMBER,
    # colour
    "model/color/resolved.py::ResolvedColor.__post_init__::int(self.alpha)": NUMBER,
    "model/color/resolved.py::ResolvedColor.__post_init__::int(self.blue)": NUMBER,
    "model/color/resolved.py::ResolvedColor.__post_init__::int(self.green)": NUMBER,
    "model/color/resolved.py::ResolvedColor.__post_init__::int(self.red)": NUMBER,
    "model/color/resolved.py::ResolvedColor.of_hex::int(digits, 16)": HEX,
    "model/color/word_colors.py::_hex_byte::int(text, 16)": HEX,
    # the content API
    "model/content/__init__.py::_set_package_compatibility_mode::int(value)": ARGUMENT,
    "model/content/addresses.py::_parse_ordinal::int(segment)": TEXT,
    "model/content/api_script.py::_num::float(value)": NUMBER,
    "model/content/api_script.py::_num::int(rounded)": NUMBER,
    "model/content/comments.py::_block_marker_at::float('inf')": CONSTANT,
    "model/content/compatibility.py::<module>::int(CompatibilityMode.WORD_2007)": MODE,
    "model/content/compatibility.py::<module>::int(CompatibilityMode.WORD_2013)": MODE,
    "model/content/compatibility.py::<module>::int(mode)": MODE,
    "model/content/compatibility.py::_as_mode::int(str(value))": STRING,
    "model/content/compatibility.py::_package_mode::int(found)": MODE,
    "model/content/compatibility.py::mode_of::int(cached)": MODE,
    "model/content/compatibility.py::set_mode::int(value)": ARGUMENT,
    "model/content/lists.py::List.__init__::int(list_id)": ARGUMENT,
    "model/content/lists.py::List._writable_level::int(level)": ARGUMENT,
    "model/content/lists.py::List.get_level_font::int(level)": ARGUMENT,
    "model/content/lists.py::List.get_level_paragraphs::int(level)": ARGUMENT,
    "model/content/lists.py::List.get_level_string::int(level)": ARGUMENT,
    "model/content/lists.py::List.level_exists::int(level)": ARGUMENT,
    "model/content/lists.py::List.set_level_alignment::int(level)": ARGUMENT,
    "model/content/lists.py::List.set_level_bullet::int(char_code)": ARGUMENT,
    "model/content/lists.py::List.set_level_bullet::int(level)": ARGUMENT,
    "model/content/lists.py::List.set_level_indents::float(bullet_number_picker_indentation)": ARGUMENT,
    "model/content/lists.py::List.set_level_indents::float(text_indent)": ARGUMENT,
    "model/content/lists.py::List.set_level_indents::int(level)": ARGUMENT,
    "model/content/lists.py::List.set_level_numbering::int(level)": ARGUMENT,
    "model/content/lists.py::List.set_level_starting_number::int(level)": ARGUMENT,
    "model/content/lists.py::List.set_level_starting_number::int(starting_number)": ARGUMENT,
    "model/content/lists.py::ListItem.level::int(value)": ARGUMENT,
    "model/content/lists.py::attach_to_list::int(level)": ARGUMENT,
    "model/content/lists.py::attach_to_list::int(list_id)": ARGUMENT,
    "model/content/lists.py::list_of::int(label.num_id)": KEY,
    "model/content/lists.py::lists_of::int(label.num_id)": KEY,
    "model/content/range.py::Range.get_tracked_changes.at::float('-inf')": CONSTANT,
    "model/content/range.py::Range.get_tracked_changes.at::float('inf')": CONSTANT,
    "model/content/range.py::Range.get_tracked_changes.at::float('nan')": CONSTANT,
    "model/content/reports.py::heading_level_of::int(match.group(1))": TEXT,
    "model/content/text_model.py::_model_character::int(code, 16)": HEX,
    "model/content/text_model.py::_sym_text::int(value, 16)": HEX,
    "model/content/text_model.py::_wildcard_text::int(value, 16)": HEX,
    "model/content/text_model.py::office_matches::float('inf')": CONSTANT,
    "model/content/trial.py::TrialPackage.compatibility_mode::int(mode_of(self))": MODE,
    "model/content/trial.py::TrialPackage.compatibility_mode::int(self._package.compatibility_mode)": MODE,
    # custom XML, fields, Markdown in
    "model/customxml/kinds.py::_symbol_of::int(text, 16)": HEX,
    "model/fields/formatting.py::_format_general::int(value)": NUMBER,
    "model/fields/formatting.py::_parse_date::int(found.group(0))": TEXT,
    "model/fields/formatting.py::_parse_double::float(text.replace('Infinity', 'inf'))": TEXT,
    "model/markdown/importer.py::_Importer.blocks::int(token.tag[1:])": TEXT,
    "model/markdown/importer.py::_Importer.list::int(token.attrs.get('start', 1))": TEXT,
    # numbering
    "model/listnumbering/definitions.py::ListNumberingDefinition.current_number_string::int(following)": TEXT,
    "model/listnumbering/definitions.py::ListNumberingDefinition.current_number_string::int(ilvl)": KEY,
    "model/listnumbering/definitions.py::ListNumberingDefinition.increment_counter::int(ilvl)": KEY,
    # styles
    "model/property_resolver.py::PropertyResolver.get_lvl_from_heading_style::int(rest)": TEXT,
    "model/property_resolver.py::heading_level_by_name::int(match.group(1))": TEXT,
    # revisions and the text model
    "model/revisions/trees.py::item_text::int(code, 16)": HEX,
    "model/text/drawings.py::smart_art::int(c.src_ord or 0)": OUTSIDE,
    "model/text/extractor.py::Extractor._run_item::int(code, 16)": HEX,
    "model/text/extractor.py::_at_address::int(rest.rstrip(']'))": TEXT,
    "model/text/field_values.py::FieldValues._sequence::int(reset.strip())": TEXT,
    # the reader itself
    "model/values.py::integer::int(digits)": OWN,
    # CR-022's check: a caller's measurement converted as its setter always converted it, and shown
    "model/values.py::_decimal::float(value)": NUMBER,
    "model/values.py::_decimal::int(value)": NUMBER,
    "model/values.py::_emu_secondary::int(value)": LENGTH,
    "model/values.py::_from_emu::float(_EMU_PER[unit])": CONSTANT,
    "model/values.py::_from_emu::int(value)": LENGTH,
    "model/values.py::_from_lines::int(value * 152400)": ARGUMENT,
    "model/values.py::_given::float(value)": NUMBER,
    "model/values.py::_given::int(number)": NUMBER,
    "model/values.py::_given::int(value)": NUMBER,
    "model/values.py::_load_limits::int(k)": TABLE,
    "model/values.py::_load_limits::int(raw['max'])": TABLE,
    "model/values.py::_load_limits::int(raw['min'])": TABLE,
    # the hand-written modules beside the model
    "traversal.py::_sym_text::int(value, 16)": HEX,
    "wml/pictures.py::_jpeg_size::float(x_density)": NUMBER,
    "wml/pictures.py::_jpeg_size::float(y_density)": NUMBER,
    "wml/pictures.py::emu_for::int(max_width_emu)": ARGUMENT,
}


def _through_integer(call: ast.Call) -> bool:
    """Whether the call's argument goes through ``values.integer``: it holds a call of ``integer``."""
    for arg in call.args:
        for node in ast.walk(arg):
            if isinstance(node, ast.Call):
                func = node.func
                if (isinstance(func, ast.Name) and func.id == "integer") or (
                    isinstance(func, ast.Attribute) and func.attr == "integer"
                ):
                    return True
    return False


def calls_in_source() -> set[str]:
    """Every ``int(`` and ``float(`` in the census's files that does not go through ``values.integer``."""
    found: set[str] = set()
    paths = [path for root in DIRECTORIES for path in sorted((PACKAGE / root).rglob("*.py"))]
    paths += [PACKAGE / module for module in MODULES]
    for path in paths:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        relative = path.relative_to(PACKAGE).as_posix()

        def visit(node: ast.AST, scope: list[str], relative: str = relative) -> None:
            for child in ast.iter_child_nodes(node):
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                    visit(child, [*scope, child.name])
                    continue
                if (
                    isinstance(child, ast.Call)
                    and isinstance(child.func, ast.Name)
                    and child.func.id in ("int", "float")
                    and not _through_integer(child)
                ):
                    found.add(f"{relative}::{'.'.join(scope) or '<module>'}::{ast.unparse(child)}")
                visit(child, scope)

        visit(tree, [])
    return found


def test_every_int_and_float_is_listed() -> None:
    """A new ``int(`` or ``float(`` fails this until it goes through ``values.integer`` or is listed with a reason."""
    found = calls_in_source()
    assert sorted(found - set(LISTED)) == [], "read it through values.integer, or list it with its reason"
    assert sorted(set(LISTED) - found) == [], "listed, but no longer in the source: take it out"


def test_every_listed_call_has_a_reason() -> None:
    reasons = {ARGUMENT, LENGTH, HEX, TEXT, CONSTANT, NUMBER, OUTSIDE, STRING, MODE, KEY, OWN, TABLE}
    assert set(LISTED.values()) <= reasons
    # the readers CR-021 section 2.3 named are no longer listed: each goes through values.integer
    gone = (
        "model/content/paragraph.py::",
        "model/content/table.py::",
        "model/content/api_script.py::_Emitter",
        "model/listnumbering/emulator.py::_as_int",
        "model/listnumbering/definitions.py::_int",
        "model/revisions/tables.py::",
        "docx/section.py::",
    )
    assert [key for key in LISTED if key.startswith(gone)] == []


def test_the_census_sees_through_integer() -> None:
    """The scanner's own rule: ``float(integer(x) or 0)`` goes through ``values.integer``; ``int(x)`` does not."""
    tree = ast.parse("float(integer(x) or 0)\nint(x)\nfloat(values.integer(y))\n")
    calls = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in ("int", "float")
    ]
    assert [_through_integer(call) for call in calls] == [True, False, True]
