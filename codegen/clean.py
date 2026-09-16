#!/usr/bin/env python
"""Delete the previous generation's modules from ``docx4j_py/``.

Called by ``codegen/generate.sh`` before the generator runs, so that a
namespace that has been renamed or removed cannot leave a stale module behind
and so that the generator's own import check sees only what it just wrote.

Everything hand written under ``docx4j_py/`` is listed in :data:`KEEP`, by path
rather than by name: after CR-001 Phase C a namespace is a *package*
(``docx4j_py/wml/``) whose ``__init__.py`` and ``el.py`` are generated and
whose ``builders.py``, ``pictures.py`` and ``sdt.py`` are not, so
whole-directory deletion is no longer enough.
"""

from pathlib import Path

#: Paths under ``docx4j_py/`` that ``codegen/generate.sh`` does not own,
#: relative to the package root.
KEEP = {
    "child.py",  # CR-001 section 5: Child, ChildList, link_parents, deep_copy
    "namespaces.py",  # section 7: docx4j's prefix table
    "runtime.py",  # section 8: the shared context, warm_up
    "fragments.py",  # section 6.2: wml(...) and to_xml(...)
    "traversal.py",  # section 6.2: walk, find, text_of
    "resources",  # the parts warm_up parses, and docx4j's default styles etc.
    "wml/builders.py",  # section 6.2: the p/r/t/tbl/tr/tc sugar
    "wml/pictures.py",  # CR-003 Phase A: inline_picture, image_size, emu_for
    "wml/sdt.py",  # CR-003 Phase A: the w:sdt family
    "openpackaging",  # CR-002: the engine, all of it hand written
}

ROOT = Path(__file__).resolve().parent.parent / "docx4j_py"


def main() -> int:
    kept = {ROOT / k for k in KEEP}

    for pycache in sorted(ROOT.rglob("__pycache__"), key=lambda p: -len(p.parts)):
        for item in pycache.iterdir():
            item.unlink()
        pycache.rmdir()

    removed = 0
    for path in sorted(ROOT.rglob("*"), key=lambda p: -len(p.parts)):
        if path in kept or any(parent in kept for parent in path.parents):
            continue
        if path.is_file():
            path.unlink()
            removed += 1
    # then whatever directory is left with nothing in it. A directory that only
    # a kept file lives in (docx4j_py/wml/, which keeps builders.py) survives,
    # and with no __init__.py in it the generator's fresh docx4j_py/wml.py
    # still wins the import, which is why generate_el.py can merge the two.
    for path in sorted(ROOT.rglob("*"), key=lambda p: -len(p.parts)):
        if path.is_dir() and not any(path.iterdir()):
            path.rmdir()
            removed += 1

    print(f"removed {removed} generated path(s) from docx4j_py/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
