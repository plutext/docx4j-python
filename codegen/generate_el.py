#!/usr/bin/env python
"""Generate the per-namespace ``el`` modules from the runtime class metadata.

CR-001 section 6.2, Phase C. docx4j's ``ObjectFactory`` knew two things xsdata's
generated classes do not say out loud: *which element name a class stands for*,
and *which class an element name takes in a given scope*. Both are in the
``XmlContext`` metadata the fork builds for every generated dataclass --- in the
element vars of each class and in the ``choices`` of each compound field --- so
``el`` is derived from the metadata, never from the schema.

What it writes
--------------

``docx4j_py/<namespace package>/el.py``
    one function per element name of that namespace, plus ``QNAME_TO_CLASS``,
    ``CLASS_TO_QNAME``, ``ELEMENTS`` and ``__all__``.
``docx4j_py/el_index.py``
    namespace URI -> ``el`` module name, so a fragment parser can resolve a root
    element without importing every namespace.
``codegen/el_tables/<name>.json``
    the committed, human-reviewed table: which class each name resolved to, the
    scopes it was seen in, and every collision with the scope-qualified name
    each losing class was given. The ``overrides`` object at the top of each
    file is the *input* a human edits; everything else is output.

Module to package
-----------------

The generator emits one *module* per namespace (``docx4j_py/wml.py``), but
``el`` has to be reachable as ``docx4j_py.wml.el``, so this script first turns
each namespace module into a package of the same name
(``docx4j_py/wml/__init__.py``). Nothing else changes: the module *name* is
unchanged, so every generated cross-module import, the import-order manifest in
``docx4j_py/__init__.py`` and the lazy ``__getattr__`` of the package inits all
keep working. See CR-001 section 14.

Run through ``codegen/generate.sh``; ``--check`` proves the whole output is
reproducible.
"""

from __future__ import annotations

import argparse
import builtins
import dataclasses
import importlib
import json
import keyword
import pkgutil
import re
import shutil
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
PACKAGE = ROOT / "docx4j_py"
TABLES = Path(__file__).resolve().parent / "el_tables"

GLOBAL_SCOPE = "(global element)"


# ---------------------------------------------------------------------------
# names
# ---------------------------------------------------------------------------

_SNAKE_1 = re.compile(r"(.)([A-Z][a-z]+)")
_SNAKE_2 = re.compile(r"([a-z0-9])([A-Z])")
_NOT_IDENT = re.compile(r"[^0-9a-zA-Z_]")

#: Names that may not be used as a plain function name in a module.
RESERVED = (
    frozenset(keyword.kwlist) | frozenset(keyword.softkwlist) | frozenset(dir(builtins))
)


def snake(name: str) -> str:
    """``CTSdtRow`` -> ``sdt_row``, ``PHyperlink`` -> ``p_hyperlink``.

    A leading ``CT``/``Ct`` is dropped first: it is XJC's rendering of the
    schema's ``CT_`` prefix and says nothing about the element.
    """
    if len(name) > 2 and name[:2] in ("CT", "Ct") and name[2].isupper():
        name = name[2:]
    return _SNAKE_2.sub(r"\1_\2", _SNAKE_1.sub(r"\1_\2", name)).lower()


def identifier(name: str) -> str:
    """Turn an XML local name into a usable Python function name.

    The element name is kept **verbatim** --- ``el.pPr`` is ``w:pPr`` --- so
    that what a user writes matches what the document holds. Two adjustments:
    characters an identifier cannot hold become ``_``, and a name that is a
    Python keyword, soft keyword or builtin takes a trailing underscore
    (``el.del_``, ``el.type_``, ``el.object_``). ``el.name`` needs none: it is
    neither.
    """
    out = _NOT_IDENT.sub("_", name)
    if not out or out[0].isdigit():
        out = f"_{out}"
    if out in RESERVED:
        out = f"{out}_"
    return out


# ---------------------------------------------------------------------------
# the model
# ---------------------------------------------------------------------------

#: The hand-written modules that sit beside the generated ones, by file name:
#: the runtime of CR-001 sections 5 to 7 and the builders of section 6.2, which
#: CR-003 Phase A split into ``builders.py``, ``pictures.py`` and ``sdt.py``.
#: ``codegen/clean.py`` keeps the same files (by path, since it has to know
#: which package each lives in); this list is what the generator must not treat
#: as a namespace module and must not import while it is building ``el``.
HAND_WRITTEN: frozenset[str] = frozenset(
    {
        "__init__.py",
        "child.py",
        "el.py",
        "builders.py",
        "pictures.py",
        "sdt.py",
        "namespaces.py",
        "runtime.py",
        "fragments.py",
        "traversal.py",
    }
)


def modules_to_packages(package: Path) -> list[str]:
    """Turn every generated namespace *module* into a package of that name.

    Returns the list of module names that hold a namespace model.
    """
    converted = []
    for path in sorted(package.rglob("*.py")):
        if path.name in HAND_WRITTEN:
            continue
        if "__NAMESPACE__" not in path.read_text(encoding="utf-8")[:4000]:
            continue
        target = path.with_suffix("")
        if target.is_dir():
            # the directory survived `clean.py` because a hand-written file
            # lives in it (docx4j_py/wml/builders.py); merge into it. With no
            # __init__.py it was a namespace-package portion, which loses to
            # the module, so the imports up to here were the module's.
            shutil.move(str(path), str(target / "__init__.py"))
        else:
            tmp = target.with_name(target.name + "__el_tmp")
            tmp.mkdir()
            shutil.move(str(path), str(tmp / "__init__.py"))
            shutil.move(str(tmp), str(target))
        converted.append(target)
    names = []
    for path in sorted(package.rglob("__init__.py")):
        text = path.read_text(encoding="utf-8")[:4000]
        if "__NAMESPACE__" in text:
            rel = path.parent.relative_to(package.parent)
            names.append(".".join(rel.parts))
    return names


def load_model(module_names: list[str]) -> tuple[dict[type, str], dict[str, str]]:
    """Import every generated module; return class -> name and namespace -> module."""
    import docx4j_py

    for name in module_names:
        importlib.import_module(name)
    # everything else the package holds, so that no class is missed
    for info in pkgutil.walk_packages(docx4j_py.__path__, "docx4j_py."):
        if info.name.rpartition(".")[2] + ".py" in HAND_WRITTEN:
            continue
        try:
            importlib.import_module(info.name)
        except Exception:  # noqa: BLE001 - a module that cannot import is not our business
            continue

    classes: dict[type, str] = {}
    for name in module_names:
        mod = sys.modules[name]
        for attr, obj in vars(mod).items():
            if (
                isinstance(obj, type)
                and dataclasses.is_dataclass(obj)
                and obj.__module__ == name
            ):
                classes[obj] = attr
    namespaces = {sys.modules[n].__NAMESPACE__: n for n in module_names}
    return classes, namespaces


def collect(classes: dict[type, str], context: Any) -> tuple[dict, dict, dict, dict]:
    """Walk the metadata: element qname -> {class: set of scope classes}.

    Two sources, and only these two:

    * ``meta.elements`` --- the single-valued and repeated element fields;
    * ``meta.choices[*].elements`` --- the compound ``content`` fields, where
      xsdata keeps the element name of each alternative.

    A class whose own ``meta.qname`` carries a namespace is a *global element
    declaration* (``w:document``, ``w:p``, ``w:ins`` as ``RunIns``); it is
    recorded under the pseudo-scope ``(global element)`` so that part roots,
    which no field refers to, are in the table too.
    """
    table: dict[str, dict[type, set[str]]] = {}
    globals_: dict[str, set[type]] = {}
    metas: dict[type, Any] = {}
    simple: dict[str, set[str]] = {}

    def add(qname: str, target: type, scope: str) -> None:
        table.setdefault(qname, {}).setdefault(target, set()).add(scope)

    def add_simple(qname: str, types: Any) -> None:
        simple.setdefault(qname, set()).update(t.__name__ for t in types)

    for cls, cls_name in sorted(
        classes.items(), key=lambda kv: (kv[0].__module__, kv[1])
    ):
        try:
            meta = context.build(cls)
        except Exception as exc:  # noqa: BLE001
            print(f"  warning: no metadata for {cls.__module__}.{cls_name}: {exc}")
            continue
        metas[cls] = meta

        if meta.qname.startswith("{"):
            globals_.setdefault(meta.qname, set()).add(cls)
            add(meta.qname, cls, GLOBAL_SCOPE)

        for vars_ in meta.elements.values():
            for var in vars_:
                if any(tp in classes for tp in var.types):
                    for tp in var.types:
                        if tp in classes:
                            add(var.qname, tp, cls_name)
                else:
                    add_simple(var.qname, var.types)
        for choice in meta.choices:
            for qname, alt in choice.elements.items():
                if any(tp in classes for tp in alt.types):
                    for tp in alt.types:
                        if tp in classes:
                            add(qname, tp, cls_name)
                else:
                    add_simple(qname, alt.types)

    # a class whose metadata cannot be built is no use as a constructor; drop it
    # and say so, and drop any element name left with no candidate at all.
    for qname in list(table):
        for cls in list(table[qname]):
            if cls not in metas:
                print(
                    f"  warning: {classes[cls]} dropped from el ({qname}): no metadata"
                )
                del table[qname][cls]
        if not table[qname]:
            del table[qname]

    return table, globals_, metas, simple


# ---------------------------------------------------------------------------
# resolution
# ---------------------------------------------------------------------------


def resolve_namespace(
    uri: str,
    entries: dict[str, dict[type, set[str]]],
    globals_: dict[str, set[type]],
    classes: dict[type, str],
    overrides: dict[str, dict[str, str]],
    reach: dict[str, int],
) -> dict[str, Any]:
    """Pick the default class of every element name, and name the losers.

    The default, in order of precedence:

    1. an explicit entry in the table's ``overrides.defaults``;
    2. the class that *is* the global element declaration for that name
       (``SdtBlock`` for ``w:sdt``, ``RunIns`` for ``w:ins``), which is docx4j's
       own answer, taken from the schema's ``jaxb:class`` customisation;
    3. the class whose scopes are the most *reachable*: each scope class is
       weighted by how many scopes declare it in turn, so ``w:tab`` in ``R``
       (which 42 classes hold) beats ``w:tab`` in ``CTMathRunTrackChange``
       (which two do), where a plain count of scopes had it the other way round;
    4. the class declared in the most scopes;
    5. the alphabetically first class name, so the output never depends on
       dictionary order.

    Every other class of a colliding name gets a **scope-qualified name**:
    ``snake(class name)`` with a leading ``ct_`` already dropped by
    :func:`snake` --- ``el.sdt_run``, ``el.sdt_row``, ``el.sdt_cell``,
    ``el.custom_xml_run``, ``el.simple_field_hyperlink`` --- and, when that tag
    is not unique in the namespace (``TblWidth`` serves ``w:left``, ``w:right``,
    ``w:top``, ...), the element name in front of it: ``el.left_tbl_width``.
    An ``overrides.aliases`` entry keyed ``"<element>:<ClassName>"`` replaces
    the computed name.
    """
    default_overrides = overrides.get("defaults", {})
    alias_overrides = overrides.get("aliases", {})

    def weights(scopes: set[str]) -> int:
        """How reachable a candidate's declaring scopes are, all told."""
        return sum(reach.get(s, 0) + 1 for s in scopes if s != GLOBAL_SCOPE)

    resolved: dict[str, dict[str, Any]] = {}
    for qname, candidates in sorted(entries.items()):
        local = qname.split("}", 1)[1]
        want = default_overrides.get(local)
        chosen: type | None = None
        reason = ""
        if want:
            for cls in candidates:
                if classes[cls] == want:
                    chosen, reason = cls, "override"
                    break
            if chosen is None:
                print(f"  warning: override {local} -> {want} names no candidate class")
        if chosen is None:
            decl = globals_.get(qname, set()) & set(candidates)
            if len(decl) == 1:
                chosen, reason = next(iter(decl)), "global element declaration"
        if chosen is None:
            chosen = min(
                candidates,
                key=lambda c: (
                    -weights(candidates[c]),
                    -len(candidates[c]),
                    classes[c],
                ),
            )
            reason = "most reachable scope" if len(candidates) > 1 else "only candidate"
        resolved[local] = {
            "qname": qname,
            "default": chosen,
            "reason": reason,
            "candidates": candidates,
        }

    # primary names first, so an alias can never take one
    taken: dict[str, tuple[str, type]] = {}
    for local, info in sorted(resolved.items()):
        name = identifier(local)
        if name in taken:
            raise SystemExit(
                f"element name clash in {uri}: {local} and {taken[name][0]}"
            )
        taken[name] = (local, info["default"])
        info["name"] = name

    # then the scope-qualified aliases; a tag used by two elements is qualified
    tag_users: dict[str, list[tuple[str, type]]] = {}
    for local, info in sorted(resolved.items()):
        for cls in sorted(info["candidates"], key=lambda c: classes[c]):
            if cls is info["default"]:
                continue
            tag_users.setdefault(snake(classes[cls]), []).append((local, cls))

    for local, info in sorted(resolved.items()):
        aliases: dict[str, str] = {}
        for cls in sorted(info["candidates"], key=lambda c: classes[c]):
            if cls is info["default"]:
                continue
            override = alias_overrides.get(f"{local}:{classes[cls]}")
            if override:
                name = override
            else:
                tag = snake(classes[cls])
                name = (
                    tag
                    if len(tag_users[tag]) == 1 and tag not in taken
                    else f"{snake(local)}_{tag}"
                )
                while name in taken:
                    name = f"{name}_"
            taken[name] = (local, cls)
            aliases[name] = classes[cls]
        info["aliases"] = aliases

    return resolved


# ---------------------------------------------------------------------------
# emitting
# ---------------------------------------------------------------------------

HEADER = '''"""``el`` for the {uri} namespace.

Generated by ``codegen/generate_el.py`` from the ``XmlContext`` class metadata;
do not edit. The naming rules, the collisions and the class each name resolves
to are in ``codegen/el_tables/{table}.json``, which a human reviews.

    >>> from {module} import el
    >>> el.{sample}()  # doctest: +SKIP

Every function takes the dataclass's own keyword arguments and passes them
straight through. A class that carries text (``w:t``, ``w:delText``,
``w:instrText``: anything whose metadata has a text value field) also takes the
text as the one positional argument, and sets ``xml:space="preserve"`` when the
text begins or ends with whitespace or holds two in a row, which is what
docx4j's ``t()`` does.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.runtime import text_element

'''


def emit_module(
    path: Path,
    uri: str,
    module: str,
    table_name: str,
    resolved: dict[str, dict[str, Any]],
    classes: dict[type, str],
    metas: dict[type, Any],
) -> None:
    """Write one ``el.py``."""
    used: dict[str, set[str]] = {}
    for info in resolved.values():
        for cls in info["candidates"]:
            used.setdefault(cls.__module__, set()).add(classes[cls])

    # Every name this module is about to define as a function. Where a class
    # name is one of them --- `mce:AlternateContent`, `Relationships`,
    # `Relationship`: the element name is the class name --- the function would
    # shadow the class it imported, so `el.Relationships()` would recurse for
    # ever and the tables below would hold the function instead of the class.
    # Those classes are imported under a leading underscore instead.
    func_names = {info["name"] for info in resolved.values()}
    func_names |= {a for info in resolved.values() for a in info["aliases"]}

    def local_of(cls_name: str) -> str:
        return f"_{cls_name}" if cls_name in func_names else cls_name

    sample = next(iter(sorted(resolved)), "p")
    out = [
        HEADER.format(
            uri=uri, module=module, table=table_name, sample=identifier(sample)
        )
    ]
    for mod in sorted(used):
        names = sorted(used[mod])
        out.append(
            f"from {mod} import (\n"
            + "".join(
                f"    {n} as {local_of(n)},\n" if local_of(n) != n else f"    {n},\n"
                for n in names
            )
            + ")\n"
        )
    out.append("\n__NAMESPACE__ = " + json.dumps(uri) + "\n\n")

    all_names: list[str] = []
    bodies: list[str] = []
    qname_to_class: list[str] = []
    class_to_qname: dict[str, str] = {}
    elements: list[str] = []

    def fn(name: str, cls: type, qname: str, local: str, note: str) -> None:
        cls_name = local_of(classes[cls])
        meta = metas[cls]
        text_var = meta.text
        doc = (
            f'    """``{{{uri}}}{local}`` -> '
            f':class:`{cls.__module__}.{classes[cls]}`{note}"""\n'
        )
        if text_var is not None:
            space = next(
                (v.name for q, v in meta.attributes.items() if q.endswith("}space")),
                None,
            )
            bodies.append(
                f"def {name}(text: str | None = None, /, **kwargs: Any) -> {cls_name}:\n"
                + doc
                + f"    return text_element({cls_name}, text, kwargs,"
                f" {json.dumps(text_var.name)}, {json.dumps(space)})\n"
            )
        else:
            bodies.append(
                f"def {name}(**kwargs: Any) -> {cls_name}:\n"
                + doc
                + f"    return {cls_name}(**kwargs)\n"
            )
        all_names.append(name)

    seen_class: dict[str, str | None] = {}
    for local, info in sorted(resolved.items()):
        cls = info["default"]
        qname = info["qname"]
        extra = ""
        if info["aliases"]:
            extra = (
                f", the {info['reason']} of "
                + str(len(info["candidates"]))
                + " classes this name takes; see "
                + ", ".join(f"``{a}``" for a in sorted(info["aliases"]))
                + "."
            )
        else:
            extra = "."
        fn(info["name"], cls, qname, local, extra)
        qname_to_class.append(f"    {json.dumps(qname)}: {local_of(classes[cls])},")
        elements.append(f"    {json.dumps(local)}: {local_of(classes[cls])},")
        for alias, alias_cls_name in sorted(info["aliases"].items()):
            alias_cls = next(
                c for c in info["candidates"] if classes[c] == alias_cls_name
            )
            fn(alias, alias_cls, qname, local, f", the {alias_cls_name} form.")
        for c in info["candidates"]:
            name = classes[c]
            seen_class[name] = (
                qname
                if name not in seen_class
                else (qname if seen_class[name] == qname else None)
            )

    for cls_name, qname in sorted(seen_class.items()):
        if qname is not None:
            class_to_qname[cls_name] = qname

    out.append("\n".join(bodies))
    out.append("\n#: Element qualified name -> the class ``el.<name>`` builds.\n")
    out.append(
        "QNAME_TO_CLASS: dict[str, type] = {\n" + "\n".join(qname_to_class) + "\n}\n"
    )
    out.append("\n#: Local element name -> the class ``el.<name>`` builds.\n")
    out.append("ELEMENTS: dict[str, type] = {\n" + "\n".join(elements) + "\n}\n")
    out.append(
        "\n#: Class -> the one element qualified name it stands for, where there is one.\n"
    )
    out.append(
        "CLASS_TO_QNAME: dict[type, str] = {\n"
        + "\n".join(
            f"    {local_of(n)}: {json.dumps(q)},"
            for n, q in sorted(class_to_qname.items())
        )
        + "\n}\n"
    )
    out.append(
        "\n__all__ = [\n"
        + "".join(
            f"    {json.dumps(n)},\n"
            for n in sorted(
                [*all_names, "CLASS_TO_QNAME", "ELEMENTS", "QNAME_TO_CLASS"]
            )
        )
        + "]\n"
    )
    path.write_text("".join(out), encoding="utf-8")


def emit_index(path: Path, namespaces: dict[str, str]) -> None:
    """Write ``docx4j_py/el_index.py``: the namespace URI index.

    Two tables, both keyed off the set of namespaces that actually have
    generated bindings:

    ``EL_MODULES``
        namespace URI -> ``el`` module, so a fragment parser can resolve one
        root element without importing every namespace.
    ``UNDERSTOOD``
        prefix -> namespace URI, CR-002 section 5.6: the namespaces this build
        of the model understands, which is what ``mc:Choice``'s ``Requires``
        is tested against. Generated rather than hand kept, so that adding a
        schema to ``codegen/generate.sh`` widens the MCE resolver by itself.
        The prefix is docx4j's, from ``docx4j_py.namespaces``.
    """
    from docx4j_py.namespaces import NAMESPACE_TO_PREFIX

    lines = [
        '"""Namespace URI index for the generated bindings.\n\n'
        "Generated by ``codegen/generate_el.py``; do not edit.\n\n"
        "``EL_MODULES`` is namespace URI -> the ``el`` module for it, which lets a\n"
        "fragment parser resolve one root element without importing every\n"
        "namespace. ``UNDERSTOOD`` is the inverse restricted to a prefix: the\n"
        "namespaces this build understands, which is what ``mc:Choice``'s\n"
        "``Requires`` is tested against (CR-002 section 5.6). It is re-exported\n"
        "as ``docx4j_py.namespaces.UNDERSTOOD``.\n"
        '"""\n\nfrom __future__ import annotations\n\n',
        "EL_MODULES: dict[str, str] = {\n",
    ]
    for uri, module in sorted(namespaces.items()):
        lines.append(f"    {json.dumps(uri)}: {json.dumps(module + '.el')},\n")
    lines.append("}\n\n")

    understood: dict[str, str] = {}
    missing: list[str] = []
    for uri in sorted(namespaces):
        prefix = NAMESPACE_TO_PREFIX.get(uri)
        if prefix is None:
            missing.append(uri)
        else:
            understood[prefix] = uri
    if missing:
        raise SystemExit(
            "codegen/generate_el.py: these generated namespaces have no prefix in "
            "docx4j_py.namespaces.PREFIXES, so UNDERSTOOD cannot name them:\n  "
            + "\n  ".join(missing)
        )

    lines.append("UNDERSTOOD: dict[str, str] = {\n")
    for prefix, uri in sorted(understood.items()):
        lines.append(f"    {json.dumps(prefix)}: {json.dumps(uri)},\n")
    lines.append('}\n\n__all__ = ["EL_MODULES", "UNDERSTOOD"]\n')
    path.write_text("".join(lines), encoding="utf-8")


def table_name_for(module: str) -> str:
    return module.removeprefix("docx4j_py.").replace(".", "_")


#: Where each Phase C name a namespace package re-exports actually lives.
PHASE_C_EXPORTS: dict[str, str] = {
    "el": "",  # the submodule of this package
    "wml": "docx4j_py.fragments",
    "to_xml": "docx4j_py.fragments",
    "FragmentError": "docx4j_py.fragments",
    "walk": "docx4j_py.traversal",
    "iter_nodes": "docx4j_py.traversal",
    "find": "docx4j_py.traversal",
    "text_of": "docx4j_py.traversal",
    "element_name": "docx4j_py.traversal",
    "warm_up": "docx4j_py.runtime",
    "link_parents": "docx4j_py.child",
    "deep_copy": "docx4j_py.child",
    # CR-003 Phase A
    "walk_all": "docx4j_py.traversal",
    "run_items_of": "docx4j_py.traversal",
    "deep_copy_as": "docx4j_py.child",
}

_FOOTER_MARKER = "# CR-001 Phase C: el, the fragment helpers and the text sugar"

FOOTER = '''

{marker}, reached
# through this package as CR-001 section 6.2 writes them
# (``from {module} import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {exports}


def __getattr__(name: str) -> object:
    """Import a Phase C helper, or the ``el`` submodule, on first use."""
    target = _PHASE_C.get(name)
    if target is None:
        raise AttributeError(f"module {{__name__!r}} has no attribute {{name!r}}")
    from importlib import import_module

    value = import_module(target) if name == "el" else getattr(import_module(target), name)
    globals()[name] = value
    return value
'''


def append_footer(init: Path, module: str) -> int:
    """Give a namespace package the Phase C re-exports, lazily.

    ``docx4j_py.wml`` has to answer ``el``, ``p``, ``r``, ``t``, ``wml``,
    ``to_xml``, ``text_of``, ``walk`` and ``find``, and the package's
    ``__init__`` is generated, so the footer is generated too. Every one of
    those names imports the classes this module defines, so the import has to
    happen after them, which a module ``__getattr__`` (PEP 562) arranges.
    """
    text = init.read_text(encoding="utf-8")
    if _FOOTER_MARKER in text:
        text = text[: text.index(_FOOTER_MARKER)].rstrip("\n# ") + "\n"
    exports = dict(PHASE_C_EXPORTS)
    exports["el"] = f"{module}.el"
    # every hand-written module of the namespace package (docx4j_py/wml/ has
    # builders.py, pictures.py and sdt.py after CR-003 Phase A): its __all__
    # is re-exported from the package, so a user writes
    # ``from docx4j_py.wml import inline_picture`` and never has to know which
    # of them it lives in. codegen/clean.py keeps the same files.
    for hand_written in sorted(init.parent.glob("*.py")):
        if hand_written.name in ("__init__.py", "el.py"):
            continue
        source = hand_written.read_text(encoding="utf-8")
        match = re.search(r"^__all__ = \[(.*?)^\]", source, re.S | re.M)
        if match:
            for name in re.findall(r'"([^"]+)"', match.group(1)):
                exports[name] = f"{module}.{hand_written.stem}"
    init.write_text(
        text
        + FOOTER.format(
            marker=_FOOTER_MARKER,
            module=module,
            exports=json.dumps(dict(sorted(exports.items())), indent=4).replace(
                "\n}", "\n}"
            ),
        ),
        encoding="utf-8",
    )
    return len(exports)


#: CR-002 section 7: what ``docx4j_py`` itself re-exports, for the common case
#: (``from docx4j_py import WordprocessingMLPackage``). The engine is hand
#: written and lives under ``docx4j_py/openpackaging/``, but ``docx4j_py``'s own
#: ``__init__`` is generated --- it carries the import-order manifest --- so the
#: re-export has to be generated with it.
_ENGINE_MARKER = "# CR-002 section 7: the engine, re-exported for the common case"

ENGINE_FOOTER = '''

{marker}.
# Lazily, through a module __getattr__ (PEP 562): the engine imports the model,
# and the model is what this module's manifest above has just finished
# importing, so an eager import here would be a cycle.
_ENGINE: dict[str, str] = {exports}


def __getattr__(name: str) -> object:
    """Import an engine name on first use. CR-002 section 7."""
    target = _ENGINE.get(name)
    if target is None:
        raise AttributeError(f"module {{__name__!r}} has no attribute {{name!r}}")
    from importlib import import_module

    value = getattr(import_module(target), name)
    globals()[name] = value
    return value
'''

#: Name -> the module it lives in.
ENGINE_EXPORTS: dict[str, str] = {
    "OpcPackage": "docx4j_py.openpackaging",
    "WordprocessingMLPackage": "docx4j_py.openpackaging",
    "LoadOptions": "docx4j_py.openpackaging",
    "PartName": "docx4j_py.openpackaging",
    "Docx4JException": "docx4j_py.openpackaging",
    "load": "docx4j_py.openpackaging.api",
    "create_package": "docx4j_py.openpackaging.api",
}


def append_engine_footer(init: Path) -> int:
    """Give ``docx4j_py/__init__.py`` the CR-002 section 7 re-exports."""
    text = init.read_text(encoding="utf-8")
    if _ENGINE_MARKER in text:
        text = text[: text.index(_ENGINE_MARKER)].rstrip("\n# ") + "\n"
    init.write_text(
        text
        + ENGINE_FOOTER.format(
            marker=_ENGINE_MARKER,
            exports=json.dumps(dict(sorted(ENGINE_EXPORTS.items())), indent=4),
        ),
        encoding="utf-8",
    )
    return len(ENGINE_EXPORTS)


def write_table(
    path: Path,
    uri: str,
    module: str,
    resolved: dict[str, dict[str, Any]],
    classes: dict[type, str],
    overrides: dict[str, dict[str, str]],
    simple: dict[str, list[str]],
) -> None:
    """Write the committed, human-reviewed table."""
    collisions = {
        local: {
            "default": {
                "name": info["name"],
                "class": classes[info["default"]],
                "chosen_by": info["reason"],
                "scopes": sorted(info["candidates"][info["default"]]),
            },
            "scoped": [
                {
                    "name": alias,
                    "class": cls_name,
                    "scopes": sorted(
                        next(
                            v
                            for c, v in info["candidates"].items()
                            if classes[c] == cls_name
                        )
                    ),
                }
                for alias, cls_name in sorted(info["aliases"].items())
            ],
        }
        for local, info in sorted(resolved.items())
        if info["aliases"]
    }
    renamed = {
        local: info["name"]
        for local, info in sorted(resolved.items())
        if info["name"] != local
    }
    doc = {
        "namespace": uri,
        "module": module,
        "generated_by": "codegen/generate_el.py",
        "note": (
            "Output, except for `overrides`, which is the input a human edits: "
            "`defaults` maps an element name to the class `el.<name>` must build, "
            "`aliases` maps '<element>:<ClassName>' to the scope-qualified name that "
            "class must get. Regenerate with codegen/generate.sh."
        ),
        "overrides": overrides,
        "counts": {
            "element_names": len(resolved),
            "simple_type_elements": len(simple),
            "scoped_names": sum(len(i["aliases"]) for i in resolved.values()),
            "keyword_renames": len(renamed),
            "collisions": len(collisions),
        },
        "note_simple_type_elements": (
            "Element names whose schema type is a simple type, so xsdata generated no "
            "class for them and `el` has no function: the parent class holds the value "
            "directly (w:recipientData.uniqueTag is bytes). Listed, not built."
        ),
        "simple_type_elements": simple,
        "keyword_renames": renamed,
        "collisions": collisions,
        "elements": {
            local: classes[info["default"]] for local, info in sorted(resolved.items())
        },
    }
    path.write_text(json.dumps(doc, indent=2, sort_keys=False) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--no-convert", action="store_true", help="assume the packages exist"
    )
    args = ap.parse_args(argv)

    sys.path.insert(0, str(ROOT))
    module_names = (
        [
            ".".join(p.parent.relative_to(PACKAGE.parent).parts)
            for p in sorted(PACKAGE.rglob("__init__.py"))
            if "__NAMESPACE__" in p.read_text(encoding="utf-8")[:4000]
        ]
        if args.no_convert
        else modules_to_packages(PACKAGE)
    )
    print(f"namespace modules: {len(module_names)}")

    from docx4j_xsdata.formats.dataclass.context import XmlContext

    context = XmlContext()
    classes, namespaces = load_model(module_names)
    print(f"classes: {len(classes)}")

    table, globals_, metas, simple = collect(classes, context)
    print(
        f"element names: {len(table)} over {len({q.split('}')[0] for q in table})} namespaces"
    )

    # How many scopes declare each class, anywhere: the weight rule 3 of
    # resolve_namespace uses to tell a run-level element from a maths one.
    reach: dict[str, int] = {}
    for candidates in table.values():
        for cls, scopes in candidates.items():
            reach[classes[cls]] = reach.get(classes[cls], 0) + len(
                [s for s in scopes if s != GLOBAL_SCOPE]
            )

    by_uri: dict[str, dict[str, dict[type, set[str]]]] = {}
    for qname, candidates in table.items():
        uri = qname[1:].split("}", 1)[0]
        by_uri.setdefault(uri, {})[qname] = candidates

    TABLES.mkdir(exist_ok=True)
    emitted: dict[str, str] = {}
    totals = {"names": 0, "scoped": 0, "renames": 0, "collisions": 0}
    for uri, module in sorted(namespaces.items(), key=lambda kv: kv[1]):
        entries = by_uri.get(uri, {})
        name = table_name_for(module)
        table_path = TABLES / f"{name}.json"
        overrides: dict[str, dict[str, str]] = {"defaults": {}, "aliases": {}}
        if table_path.exists():
            previous = json.loads(table_path.read_text(encoding="utf-8"))
            overrides = previous.get("overrides") or overrides
            overrides.setdefault("defaults", {})
            overrides.setdefault("aliases", {})
        resolved = resolve_namespace(uri, entries, globals_, classes, overrides, reach)
        path = PACKAGE.parent / Path(*module.split("."), "el.py")
        emit_module(path, uri, module, name, resolved, classes, metas)
        write_table(
            table_path,
            uri,
            module,
            resolved,
            classes,
            overrides,
            {
                q.split("}", 1)[1]: sorted(types)
                for q, types in sorted(simple.items())
                if q.startswith("{" + uri + "}")
            },
        )
        append_footer(PACKAGE.parent / Path(*module.split(".")) / "__init__.py", module)
        emitted[uri] = module
        totals["names"] += len(resolved)
        totals["scoped"] += sum(len(i["aliases"]) for i in resolved.values())
        totals["renames"] += sum(
            1 for local, i in resolved.items() if i["name"] != local
        )
        totals["collisions"] += sum(1 for i in resolved.values() if i["aliases"])

    emit_index(PACKAGE / "el_index.py", emitted)
    engine = append_engine_footer(PACKAGE / "__init__.py")
    print(f"engine re-exports on docx4j_py: {engine}")
    print(
        f"el modules: {len(emitted)}; element names {totals['names']}, "
        f"scope-qualified {totals['scoped']}, keyword renames {totals['renames']}, "
        f"collisions {totals['collisions']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
