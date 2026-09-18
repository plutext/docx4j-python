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

| 9 | The XML serializer's namespace bookkeeping is O(prefixes) per element. `EventHandler.start_tag` copies the whole prefix-URI map for every element and `start_namespaces` then compares the copy against the parent's entry by entry, so a large user `ns_map` costs one dict copy and one lookup per prefix per element -- and the answer is "declare nothing" for every element but the root. `add_namespace` adds `prefix_exists`, a linear scan of the map's values, per element and per attribute. | `.venv-fork/bin/python scripts/bench_marshal.py`: `samples/Symbols.docx`, 772 KB and 24,758 elements, with docx4j's 130-prefix table. Before: 481 ms, `start_namespaces` the largest single function at 0.39 s of a 2.0 s profile, 3.2 M `dict.get`. After: 336 ms. | to offer upstream; in the fork, commit `bdcc8a95` (the frames share the map copy-on-write and index its URIs) |

| 10 | The serializer asks for an `xsi:type` for every element value: `convert_value` to `convert_any_type` to `convert_xsi_type` to `convert_dataclass`, three `isinstance` checks deep, when the value's class **is** the field's declared class and `xsi_type()` therefore answers `None` by construction. The decision is a property of the field, not of the value. | Same benchmark: `convert_xsi_type` called 378,244 times for one part, always answering `None`. 336 ms to 306 ms. | to offer upstream; in the fork, commit `fd28e1eb` (`XmlVar.fast_element`, filled lazily beside the `namespace_matches` upstream already caches there) |

| 11 | `XmlMeta.get_element_vars` and `get_attribute_vars` chain and `sorted` the class's vars on every call, and the serializer calls both once per **instance**; the answer is a property of the class. 49,516 `sorted` calls for one 24,758-element part. | Same benchmark; 306 ms to 262 ms with both lists built once and kept on the meta. Upstream's only callers are `next_value` and `next_attribute`, which iterate and never modify. | to offer upstream; in the fork, commit `1d8b1f46` |

| 12 | `EventGenerator` routes every event of a nested element through four generator frames between the parent and the child (`convert_value`, `convert_list`, or `convert_elements` and `convert_choice`), each resumed once per event that passes through it; the branch they pick is fixed by the field's metadata. `EventHandler.write` also builds a list per event (`for name, *args in events`). | Same benchmark: 1.0 M generator resumptions for 79,960 events. 262 ms to 249 ms. | to offer upstream; in the fork, commit `8f2fe666` |


Reports 9 to 12 are the fork's CR-001 Phase B, the serialiser (`~/git/docx4j-xsdata`,
`docs/fork/CR-001-serialiser.md` section 9 and `docs/fork/CHANGES.md` stage 5). Each is a
separate, self-contained diff over one of upstream's files that leaves every byte of the output
and every one of upstream's 1,122 tests alone, which is why each is listed on its own: together
they take a 772 KB main part from 481 ms to 249 ms, 1.57 to 3.03 MiB/s, and none of them depends
on any other fork change. A fifth idea from the same CR --- collapsing a start tag and its
attributes into one event --- was implemented, measured at nothing, and reverted, so there is
nothing to offer there.

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
