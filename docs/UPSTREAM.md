# Upstream reports to file against tefra/xsdata

Deferred on 2026-09-12: fix locally in the fork first, report upstream later with the fork's patch
attached. Each entry: the bug, how to reproduce it from this repository, and the fork commit once
one exists.

| # | Bug | Reproduce | Status |
|---|---|---|---|
| 1 | `--structure-style namespaces` emits mutually importing modules for schemas with cross-namespace type cycles (WML embeds DML, DML embeds WML through `a:graphicData`); the generated package cannot be imported. `namespace-clusters` refuses to generate at all. | `xsdata generate --structure-style namespaces schemas/wml/wml.xsd`, then `python -c "import <package>"`. REPORT.md section 3.1. | to report; fixed in the fork, commit `2c3c4f03` (`DeferredImports`), 2026-09-12 |
| 2 | `Error: Missing inner class` for an anonymous complexType declared inside an element that belongs to a recursive group (`w:dir`, `w:bdo` in `EG_PContent` / `EG_RunLevelElts`). | Generate from the unpatched docx4j `xsd/wml/wml.xsd`. REPORT.md section 2.1; the local workaround is the marked patch in `schemas/wml/wml.xsd`. | to report; generator fix later |
| 3 | Lenient parsing (`fail_on_unknown_properties=False`) discards elements and attributes with no warning, log record or counter. Arguably a feature request: a skipped-content report. | Parse `samples/2010-sample1.docx` `word/fontTable.xml`; `mc:Ignorable` is dropped silently. REPORT.md section 7.5. | to report as a feature request; fixed in the fork, commit `84fd5138` (`ParserConfig(skipped_report=True)`), 2026-09-12 |
| 4 | Under `--structure-style namespaces`, a namespace URI that is a prefix of another (the 2018 animation namespaces) yields a module and a package of the same name in one directory (`animation.py` beside `animation/`); the package shadows the module and its classes are unreachable. | Generate `schemas/wml/wml.xsd` with the namespaces style. Worked around with a module-name substitution in the config. | to report; workaround only |
| 5 | Under the namespaces layout a cross-module circular reference is emitted as a quoted forward reference but never imported (`NameError: name 'CtRuby' is not defined` from the OMML module). | Same generation as 4. | to report; fixed in the fork, commit `2c3c4f03` |
| 6 | `strongly_connected_components` (`xsdata/utils/graphs.py`) walks `set(edges)`, so the components it yields come out in a different order in every process; the order is an input to `DesignateClassPackages`. It is invisible upstream because nothing writes it to a file, and it stops being invisible the moment something does -- the fork's import-order manifest was reshuffled on every regeneration. | `codegen/generate.sh` twice and `diff -r` the output, before the fix. | to report **only if `DeferredImports` (report 1) is upstreamed**, or as a small determinism fix on its own; fixed in the fork, commit `e7e1310d` |

| 7 | Under the namespaces layout, `DeferredImports` misses a cross-module reference that appears **only inside a compound field's `choices`**: the generated `inkml` module names `ForwardRef("Math")` in the choices of `mapping` and never imports MathML's `Math`, so building the class metadata raises `NameError: name 'Math' is not defined`. Report 5's fix covers plain field references, not choice ones. | `.venv-fork/bin/python -c "from docx4j_xsdata.formats.dataclass.context import XmlContext; import docx4j_py.inkml as i; XmlContext().build(i.MappingType)"`. Found 2026-09-12 building `el` (CR-001 section 14.8 point 4). | to report; **not yet fixed in the fork** |

| 8 | Under `--structure-style namespaces`, an `el`-style factory function whose name equals the class it imports shadows that class. Not an xsdata bug — it is this repository's `codegen/generate_el.py` — but it is worth mentioning when reporting report 1, because it is a hazard of any generator that writes helper functions beside imported classes: `def Relationships(**kw) -> Relationships: return Relationships(**kw)` recurses for ever, and the generated lookup tables end up holding the function. | Generate `schemas/relationships.xsd`; `docx4j_py.relationships.el.Relationships()` before the fix. | **fixed here** 2026-09-12 (CR-002 Phase A): the class is imported as `_Relationships` when a function of that name is emitted. Nothing to report upstream |

Nothing new surfaced in the Phase B integration stage beyond report 6; Phase C added report 7, and
CR-002 Phase A added 8, which is ours rather than upstream's. The engine itself needed **no** fork
change: the runtime's skipped-content report, numeric booleans and `ListFactory` were enough to
read and write every part of sixteen real documents losslessly. Two things that looked
like upstream bugs were not: `RenameDuplicateClasses` numbering the classes `DisambiguateChoices`
invents (`T1`, `Ins2`) is correct behaviour for classes that have no xml identity, and the
explosion of those suffixes under a JAXB-derived name table was the fork's `<ClassNames>` handler
applying the table too eagerly, fixed in commit `930ccd57`.

Also worth mentioning upstream when reporting, not bugs: `--kw-only` is documented but not a
26.2 option (keyword-only is unconditional); the generator shells out to `ruff` by bare name and
fails with `FileNotFoundError` when the virtual environment's `bin` is not on `PATH`.
