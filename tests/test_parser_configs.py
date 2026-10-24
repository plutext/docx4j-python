"""Every parse in the tree is configured by ``docx4j_py.runtime.parser_config()``. CR-018 section 4.1a.

The fork's stage 13 makes the exact reading of a boolean or an enumeration a
``ParserConfig`` option (``exact_types``, item 29), off by default, and
``runtime.parser_config()`` turns it on (:data:`docx4j_py.runtime.EXACT_TYPES`).
A parser built on a ``ParserConfig`` of its own would read ``w:b w:val=" false"``
as ``False`` again, which Word reads as on, and a save would write it ``0``: the
one save over Word's probes that changed meaning (CR-018 finding 1). The fork
session's warning was exactly that, so this walks ``docx4j_py/``, ``scripts/``
and ``tests/`` for a ``ParserConfig(...)`` call, and for an ``XmlParser`` or a
``TreeParser`` built without a config, and allows only the places that need
otherwise, each with its reason.
"""

from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

#: (path, enclosing function or "" for module level): why the call is not ``parser_config()``'s.
ALLOWED: dict[tuple[str, str], str] = {
    ("docx4j_py/runtime.py", "parser_config"): "the one place the engine's config is made",
    ("scripts/roundtrip.py", "_upstream_configs"): "the upstream runtime (--runtime xsdata), REPORT.md's baseline",
    ("scripts/bench.py", "main"): "upstream xsdata 26.2 over the baseline bindings (.venv), REPORT.md's benchmark",
}

PARSERS = ("XmlParser", "TreeParser")


def _name(func: ast.expr) -> str | None:
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute):
        return func.attr
    return None


def _calls(path: Path) -> list[tuple[str, str, int, ast.Call]]:
    """Every ``ParserConfig(...)``, ``XmlParser(...)`` and ``TreeParser(...)`` call: (name, function, line, call)."""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    out: list[tuple[str, str, int, ast.Call]] = []

    def visit(node: ast.AST, function: str) -> None:
        for child in ast.iter_child_nodes(node):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                visit(child, function or child.name)
                continue
            if isinstance(child, ast.Call):
                name = _name(child.func)
                if name == "ParserConfig" or name in PARSERS:
                    out.append((name, function, child.lineno, child))
            visit(child, function)

    visit(tree, "")
    return out


def _sources() -> list[Path]:
    out = []
    for folder in ("docx4j_py", "scripts", "tests"):
        out.extend(sorted((ROOT / folder).rglob("*.py")))
    return [path for path in out if "__pycache__" not in path.parts]


def test_every_parser_config_is_runtime_s() -> None:
    offenders = []
    used = set()
    for path in _sources():
        relative = path.relative_to(ROOT).as_posix()
        for name, function, line, call in _calls(path):
            if name == "ParserConfig":
                key = (relative, function)
                if key in ALLOWED:
                    used.add(key)
                    continue
                offenders.append(f"{relative}:{line}: ParserConfig(...) in {function or 'the module'}")
            else:
                configured = any(keyword.arg == "config" for keyword in call.keywords) or len(call.args) >= 1
                if not configured:
                    offenders.append(f"{relative}:{line}: {name}(...) with the default ParserConfig")
    assert not offenders, "build the config with docx4j_py.runtime.parser_config():\n" + "\n".join(offenders)
    assert used == set(ALLOWED), sorted(set(ALLOWED) - used)  # an allowance nothing needs any more goes


def test_the_engine_s_config_reads_exactly() -> None:
    from enum import Enum

    from docx4j_py.runtime import EXACT_TYPES, parser, parser_config

    assert frozenset({bool, Enum}) == EXACT_TYPES
    assert parser_config().exact_types == EXACT_TYPES
    assert parser().config.exact_types == EXACT_TYPES


def test_a_kept_value_is_recorded_and_written_as_read() -> None:
    """``" false"`` kept on purpose (no warning), ``"off"`` kept as invalid; both in ``part.kept``, saved as read."""
    import io
    import warnings
    import zipfile

    from docx4j_py.openpackaging.api import load

    package = load(ROOT / "tests" / "fixtures" / "onoff-fun" / "props-b-case.docx")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        _ = package.main_document_part.contents
    kept = {k.text: k for _part, k in package.kept}
    assert kept[" false"].reason == "exact" and kept["false "].reason == "exact"
    assert kept["FALSE"].reason == kept["OFF"].reason == kept["00"].reason == "invalid"
    assert kept[" false"].path.endswith("/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b/@"
                                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
    assert kept[" false"].field_name == "val" and kept[" false"].parent_class == "BooleanDefaultTrue"
    messages = " ".join(str(w.message) for w in caught)
    assert "FALSE" in messages and "` false`" not in messages  # the exact keep raises no warning
    assert package.main_document_part.kept == [k for _p, k in package.kept if _p == package.main_document_part.part_name]
    assert not package.skipped  # a kept value is not skipped content
    xml = zipfile.ZipFile(io.BytesIO(package.save(normalise=False))).read("word/document.xml").decode()
    assert 'w:val=" false"' in xml and 'w:val="false "' in xml


def test_a_strict_load_does_not_refuse_a_value_word_reads() -> None:
    """CR-018 section 4.11: ``LoadOptions(strict=True)`` raises on skipped content only, never on a kept value.

    The fork's exact mode raises ``ParserError`` for a kept value under
    ``fail_on_converter_warnings``; no parse of the engine's turns that on (a
    strict load is ``lenient=True`` with the skipped report, and ``wml(...)``
    raises on skipped content), so neither starts failing on ``" false"``.
    """
    from docx4j_py.fragments import wml
    from docx4j_py.openpackaging.api import load
    from docx4j_py.openpackaging.load import LoadOptions

    package = load(ROOT / "tests" / "fixtures" / "onoff-fun" / "props-b-case.docx", options=LoadOptions(strict=True))
    _ = package.main_document_part.contents
    assert {k.reason for _p, k in package.kept} == {"exact", "invalid"}
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    assert wml(f'<w:b {w} w:val=" false"/>').val == " false"
    assert wml(f'<w:b {w} w:val="off"/>').val == "off"
