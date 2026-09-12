# codegen: how the bindings are generated

CR-001 Phases B and C. Almost everything under `docx4j_py/` is generated; this directory holds
the inputs and the scripts that produce it.

| | |
|---|---|
| `generate.sh` | the one command that regenerates `docx4j_py/`; `--check` proves it is reproducible |
| `clean.py` | deletes the previous run's output, keeping the hand-written files it lists |
| `../.xsdata.phase-b.xml` | the generator configuration |
| `../schemas/` | docx4j's XSD tree with marked patches, see `schemas/PATCHES.md` |
| `derive_names.py`, `names/` | the docx4j class-name tables, see `names/README.md` |
| `generate_el.py`, `el_tables/` | the per-namespace `el` object factory, [below](#el-the-object-factory) |

The hand-written files under `docx4j_py/`, all listed in `clean.py`'s `KEEP`:

| | CR-001 |
|---|---|
| `child.py` | section 5: `Child`, `ChildList`, `link_parents`, `iter_children`, `deep_copy`; and CR-002 section 5.6's `mce` mode and `mce_branch` |
| `namespaces.py` | section 7: docx4j's prefix table |
| `runtime.py` | section 8: the shared `XmlContext`, the parser and serialiser factories, `warm_up()` |
| `fragments.py` | section 6.2: `wml(...)` and `to_xml(...)` |
| `traversal.py` | section 6.2: `walk`, `iter_nodes`, `find`, `text_of`, `element_name`, each with CR-002's `mce` mode |
| `wml/builders.py` | section 6.2: the `p` / `r` / `t` / `tbl` sugar |
| `resources/` | the parts `warm_up()` parses, and (CR-002) docx4j's default `styles.xml`, `numbering.xml` and `fontTable.xml` |
| `openpackaging/` | **CR-002, the engine**: Open Packaging, parts, load and save, all hand written |

## The command

```bash
codegen/generate.sh            # regenerate
codegen/generate.sh --check    # regenerate twice and prove the output is byte identical
```

which is, in essence,

```bash
export PATH="$PWD/.venv-fork/bin:$PATH"      # the generator shells out to `ruff` by bare name
python codegen/clean.py
docx4j-xsdata generate -c .xsdata.phase-b.xml schemas/docx4j_python__ROOT.xsd
python codegen/generate_el.py
```

`schemas/docx4j_python__ROOT.xsd` is the one SOURCE the generator takes, and the only file under
`schemas/` that is not docx4j's: the CLI accepts a single entry point, and CR-002 section 9 added
four more of them beside `wml/wml.xsd` — `relationships.xsd` and the three `docProps/` schemas —
so they are collected there by `xsd:import`.

`.venv-fork` has the fork (`docx4j-xsdata`, the renamed xsdata 26.2 + seven options) installed
editable from `~/git/docx4j-xsdata`. 4 seconds. The warnings are the six unmodified xsdata gives
(five duplicate `xml:` types, one absent `CT_GeoHierarchyEntity`) plus about 225 from
`<ClassNames>`, which are the name table reporting what it deliberately did not do: 33 entries
whose docx4j value is an inner Java class (`R.Cr`, `P.Hyperlink`) and therefore not a Python
class name, and the element entries that name a class the `types` section already names. Both
are expected and are explained in `names/README.md` and CR-001 section 13.3.

`clean.py` deletes everything under `docx4j_py/` that is not in its `KEEP` list, so that a
renamed namespace cannot leave a stale module behind and the generator's own import check (it
validates its output by importing the whole package) sees only what it just wrote. The
pre-Phase-B baseline cannot live there at all and moved to `baseline/docx4j_py/generated/` for
exactly that reason.

One subtlety, and the reason `KEEP` holds paths rather than names: `docx4j_py/wml/builders.py` is
hand written and lives *inside* a generated package. `clean.py` leaves the directory in place
with nothing in it but `builders.py` and **no `__init__.py`**, which makes it a namespace-package
portion — and a namespace-package portion loses the import to a module of the same name, so
xsdata's fresh `docx4j_py/wml.py` still wins and its import check passes. `generate_el.py` then
merges the two by moving the module in as the package's `__init__.py`.

**Regeneration is reproducible.** `codegen/generate.sh --check` regenerates twice and diffs: every
file under `docx4j_py/` byte identical — 135 generated modules, the import-order manifest and the
CR-002 re-export footer at the bottom of `docx4j_py/__init__.py`, every `el.py`, and the 25
hand-written engine modules `clean.py` keeps. It used to shuffle on every run, because the fork's `render_import_order` walked the
strongly connected components of an unordered graph; the walk is sorted now (fork
`docs/fork/CHANGES.md` stage 4, `docs/UPSTREAM.md` report 6).

## What the configuration says

The five fork generator options (fork `docs/fork/CHANGES.md` stages 2 to 4):

```xml
<AllOptional>true</AllOptional>                       <!-- every field None | T = None -->
<SchemaDefaults>metadata</SchemaDefaults>             <!-- schema default in metadata, not the field -->
<ListFactory>docx4j_py.child.ChildList</ListFactory>  <!-- default_factory of every list field -->
<DeferredImports>true</DeferredImports>               <!-- makes the namespaces layout importable -->
<ClassNames>codegen/names</ClassNames>                <!-- docx4j's class names, see names/README.md -->
```

plus, new in Phase B:

```xml
<Structure>namespaces</Structure>      <!-- one module per XML namespace -->
<UnnestClasses>true</UnnestClasses>    <!-- R.T becomes a top level class -->
<Extensions>
  <Extension type="class" class=".*" import="docx4j_py.child.Child"/>
</Extensions>
```

and a `<Substitutions>` table of one `type="package"` entry per namespace, which is what turns
`org.openxmlformats.schemas.wordprocessingml.2006.main` into `docx4j_py.wml`.

`applyIfDerived` is deliberately left off on the extension: a class that already extends a
generated class (the element-specific `RunIns` over `CT_TrackChange`) inherits `Child` from its
base rather than listing it twice. Enumerations use a different template and are never touched.
Result: **1,795 generated dataclasses, every one of them a `Child`; 0 without** (plus 325 enums,
which the extension never touches).

## The layout

61 namespace packages (one `__init__.py` of classes and one `el.py` each) --- 53 of CR-001 plus
the eight CR-002 section 9 added, `relationships` and the seven under `docprops` --- 5 grouping
packages, 135 generated modules in all; the class half is 79,412 lines and 2.4 MB, the `el` half 21,699 lines and
1.1 MB. **A namespace is a package, not a module** (Phase C): `docx4j_py/wml/__init__.py` holds
the classes and `docx4j_py/wml/el.py` the factory. The module *name* is unchanged, so everything
below reads as it did in Phase B.

| module | namespace | docx4j |
|---|---|---|
| `docx4j_py.wml` | `…/wordprocessingml/2006/main` | `org.docx4j.wml` |
| `docx4j_py.dml.main` | `…/drawingml/2006/main` | `org.docx4j.dml` |
| `docx4j_py.dml.picture` | `…/drawingml/2006/picture` | `org.docx4j.dml.picture` |
| `docx4j_py.dml.wordprocessing_drawing` | `…/drawingml/2006/wordprocessingDrawing` | `org.docx4j.dml.wordprocessingDrawing` |
| `docx4j_py.dml.chart`, `.chart_drawing`, `.diagram`, `.locked_canvas`, `.spreadsheet_drawing`, `.compatibility` | the other `drawingml/2006/*` | `org.docx4j.dml.*` |
| `docx4j_py.math` | `…/officeDocument/2006/math` | `org.docx4j.math` |
| `docx4j_py.mce` | `…/markup-compatibility/2006` | `org.docx4j.mce` |
| `docx4j_py.shared_types` | `…/officeDocument/2006/sharedTypes` | `org.docx4j.sharedtypes` |
| `docx4j_py.customxml` | `…/schemaLibrary/2006/main` | `org.docx4j.customxml` |
| `docx4j_py.w14`, `docx4j_py.w15` | `…/office/word/2010|2012/wordml` | `org.docx4j.w14`, `.w15` |
| `docx4j_py.wne` | `…/office/word/2006/wordml` | `org.docx4j.com.microsoft.…word.x2006.wordml` |
| `docx4j_py.word.wpc`, `.wp14`, `.wpg`, `.wps`, `.wp15`, `.msink` | the Word 2010/2012 drawing namespaces | `org.docx4j.com.microsoft.…word.x2010.*` |
| `docx4j_py.oart.*` (28 modules) | `schemas.microsoft.com/office/drawing/*`, thememl, the Excel and PowerPoint drawing extensions | docx4j's `xsd/odrawxml` |
| `docx4j_py.mathml`, `docx4j_py.inkml`, `docx4j_py.xml_ns` | MathML, InkML, the `xml:` namespace | |

Every module is importable from any entry point (`DeferredImports`), and each package `__init__`
re-exports its modules' classes through a module `__getattr__`, so `from docx4j_py.dml import
Graphic` works even though the classes live in `docx4j_py/dml/main.py`.

Three departures from CR-001 section 6.1, all recorded in CR-001 section 13:

* **`docx4j_py.dml` is a package, and DML main is `docx4j_py.dml.main`.** In Java
  `org.docx4j.dml` holds both classes and subpackages; in Python a name is either a module or a
  package. The lazy `__init__` makes `docx4j_py.dml.Graphic` work anyway.
* **`docx4j_py/__init__.py` belongs to the generator.** It carries the import-order manifest
  `DeferredImports` needs, so importing anything under `docx4j_py` — `docx4j_py.child`
  included — imports the whole model (0.59 s warm, 164 modules). Lazy per-namespace import is Phase D.
  One consequence: the baseline cannot live inside `docx4j_py/` any more, since importing it
  would run the manifest, which needs `docx4j_xsdata`. It moved to `baseline/docx4j_py/generated`
  (see **The baseline** below) and is untouched.
* **No `docx4j_py.vml` and no `docx4j_py.w16*`.** Neither namespace is in `schemas/wml/wml.xsd`'s
  import closure: `w:pict` takes VML through an `xsd:any`, and the 2016 namespaces are not in
  docx4j's WML schema set at all. Phase D.

## Class names

**Names in this output are docx4j's**, through `<ClassNames>codegen/names</ClassNames>`. `w:pPr`'s
type is `PPr`, `w:t`'s is `Text`, the two element-specific subclasses of `CT_RunTrackChange` are
`RunIns` and `RunDel`, `ST_Jc` is `JcEnumeration`, and the part roots are `Document`, `Styles`,
`Numbering`, `Fonts`, `Comments`, `Hdr`, `Ftr`, `Settings`, `WebSettings`, `Footnotes`,
`Endnotes`, `GlossaryDocument`. `tests/test_names.py` pins them, and pins that a rename never
touched `Meta.name`, which is still the name the schema gave.

The tables are matched to the generator by the `namespace` key **inside** each JSON, not by the
file name, so `codegen/names/dml.json` feeding `docx4j_py.dml.main` is not a mismatch.

**77 of the 2,120 classes carry a numeric suffix** (79 by a naive count of names ending in a
digit; `CTInteger2` and `CTInteger255` of OMML are schema names). Against 82 with no name table
at all, so the table is free — but getting there needed two rules in the fork, because docx4j's
JAXB model has one class where xsdata has several (CR-001 section 13.3, fork
`docs/fork/CHANGES.md` stage 4). Applied without them the table produced 333 suffixes and lost
`Text`, `BooleanDefaultTrue` and `CTMarkup` to `Text1`, `BooleanDefaultTrue1`, `Ctmarkup1`.

What keeps a suffix, and why it does not matter:

* **72 intermediate choice classes.** Where a compound field has two elements of the same type
  the generator invents a class per element name to tell them apart — `w:t`, `w:instrText` and
  `w:delInstrText` are all `CT_Text` — and those classes have no xml name of their own. docx4j has
  no class for them either: JAXB used one `JAXBElement` per name over the shared type. Each is a
  subclass of the class the table *did* name, so `isinstance(x, Text)` holds.
* **5 type classes behind a same-named global element**: `Body1`, `Graphic1`, `Pic1`, `Anchor1`,
  `Inline1`. The element class keeps the bare name, and it is the one the containing field refers
  to (`Document.body` is a `Body`).

`codegen/names/wml_overrides.json` is the one hand-written file in `names/`: three **scoped**
entries, `t@r`, `instrText@r` and `delInstrText@r`, which name the run's three text classes `RT`,
`RInstrText` and `RDelInstrText` instead of `T2`, `InstrText2` and `DelInstrText2`. That is the
same shape the generator already gives their siblings (`R.Cr` is `RCr`) and the shape docx4j gives
them in Java. `derive_names.py` never writes that file.

## Results, 2026-09-12

Round trip, `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata`,
50 parts from 12 documents in `samples/`, no `--ignore-defaults`, no `--tolerant-factory`,
no `--nsmap-from-source`:

| | Phase A (namespaces, no schema patch) | Phase B, project half | **Phase B final** |
|---|---:|---:|---:|
| parts parsed and serialised | 50 | 50 | **50** |
| canonically identical | 19 | 29 | **50** |
| element-dropped / element-added / child-order / text | 0 | 0 | **0** |
| attribute-added | 0 | 0 | **0** |
| attribute-dropped | 10 | 0 | **0** |
| attribute-value | 3,034 | 3,034 | **0** |
| boolean respellings (not a difference) | — | — | 34 |
| skipped elements and attributes | silent | silent | **0, and measured** |

The 10 dropped attributes were `mc:Ignorable` on `/w:fonts`, closed by the `anyAttribute` patch
(`schemas/PATCHES.md`). The 3,034 `attribute-value` differences were all the boolean spelling and
are closed by `SerializerConfig(bool_format="numeric")` (CR-001 open question 5), which the
harness now passes by default for the fork runtime. That leaves 34 the other way round — one
`styles.xml` whose producer wrote `true`/`false` where Word writes `1`/`0` — which
`scripts/canon.py` counts as its own `boolean-spelling` category and excludes from the identical
verdict, because `1` and `true` are the same `xsd:boolean`. With `--no-bool-numeric` the same
corpus gives 3,034 of them, so `numeric` is 89 times closer to what the corpus holds.

The harness prints a verdict block and **exits non-zero** on any parse failure, any real
difference, or any skipped content.

`scripts/checks.py out/phase-b`: 50/50 parts preserve child order exactly; 3,709 `w:t` nodes and
114 `xml:space` attributes compared with 0 mismatches; w14/w15 node and attribute counts equal on
all 6 parts that carry them; `mc:AlternateContent` and `mc:Fallback` counts equal on all 3. Six
parts declare `mc:Ignorable` prefixes the fixed `ns_map` does not (`w16se`, `w16cid`) — REPORT.md
7.2b, the engine's job (CR-002); it is six parts rather than four now precisely because
`mc:Ignorable` survives on more roots.

Parent pointers, `scripts/parents.py --models-module docx4j_py.wml`: 50 parts, 34,147 nodes,
**0 parent problems**, and serialising a tree before and after `link_parents` gives identical
bytes.

| cost | share of parse |
|---|---:|
| `Child.__post_init__` + `ChildList` during parse | ~5% (noise band over 39 parts) |
| `link_parents` after parse | **9.4%** (37 ms against 400 ms); **7.2%** since Phase C put it on the shared child enumeration |
| together | ~14%, against CR-001 section 8's 20% budget |

Import and throughput: `docx4j_py.wml` imports in **0.86 s cold, 0.59 s warm** and pulls in 164
modules — the whole model, because `docx4j_py/__init__.py` carries the import-order manifest.
Parse runs at 2.1 MiB/s with the lxml handler. `tests/test_performance.py` guards both.

## The baseline

`baseline/docx4j_py/generated/` is the output of **unmodified** xsdata 26.2 — the `clusters`
layout of `.xsdata.xml` that REPORT.md measured — kept for comparison and untouched. It imports
`xsdata`, which only `.venv` has, and its 1,625 modules import each other as
`docx4j_py.generated.<module>`, so it needs an (empty) `docx4j_py` package above it. Putting
`baseline/` in front of the repository root on the path shadows the Phase B package with it, and
`scripts/roundtrip.py` does that automatically for any `--package` under `docx4j_py.generated`:

```bash
.venv/bin/python scripts/roundtrip.py                      # the baseline
PYTHONPATH=baseline .venv/bin/python -c "import docx4j_py.generated.document"
```

## Tests

```bash
.venv-fork/bin/python -m pytest                  # 576 passed
.venv-fork/bin/python -m pytest -m "not slow"    # without the corpus round trip, the warm-up
                                                 # measurement, the el regeneration and the timings
```

Phase B: `tests/test_child.py` (parent pointers, `ChildList`, `deep_copy`), `tests/test_names.py`
(the docx4j names and `Meta.name`), `tests/test_skipped_report.py` (the fork's skipped-content
report and the fixture that proves it catches an unknown element and an unknown attribute),
`tests/test_roundtrip.py` (the harness over `samples/`, zero differences) and
`tests/test_performance.py` (import time, parse throughput, the `link_parents` budget).

CR-002 Phase A: `tests/openpackaging/` --- 295 tests over the engine, and the ones that bear on
this directory are `test_mce.py` (`UNDERSTOOD` is the set of generated namespaces and nothing more)
and `test_threads_and_import.py` (no engine module imports a model class at module level, so the
engine costs the model plus 24 ms and never the other way round).

Phase C: `tests/test_el.py` (every element name the metadata reaches has an entry, every entry
builds and serialises under its own name, the naming and collision rules),
`tests/test_codegen_el.py` (a table per `el` module, and regenerating changes nothing),
`tests/test_builders.py` (the sugar and the run-option mapping, mirroring the TypeScript's
`smoke.mjs`), `tests/test_fragments.py` (fourteen fragments through `to_xml(wml(x))`),
`tests/test_traversal.py` (`walk`, `find`, and `text_of` against an independent lxml extraction
over all 13 sample documents), `tests/test_runtime.py` (the prefix table, `xml:space`,
`warm_up`) and `tests/test_threads.py` (a shared context across eight threads).

## `el`, the object factory

`generate_el.py` walks the `XmlContext` metadata of every generated class — the element vars, the
`choices` of every compound field, and every class that is a global element declaration — and
writes, per namespace, a function for each element name of that namespace:

```python
from docx4j_py.wml import el
el.p(content=[el.r(content=[el.t("Hello World")])])
```

It also turns each namespace module into a package (above), writes `docx4j_py/el_index.py`
(namespace URI → `el` module) and a footer on each package `__init__.py` that resolves `el` and
the hand-written helpers lazily.

**2,058 element names over 58 namespaces**, 225 scope-qualified names, 35 renamed for a Python
keyword or builtin, 153 names that take more than one class. WML alone: 604, 135, 7 and 83.

`el_tables/<namespace>.json` is committed **for a human to review**: which class each name
resolved to and why (`chosen_by`), the scopes each was seen in, every collision with the
scope-qualified name its losing classes were given, the keyword renames, and the element names
that have no class at all because their schema type is a simple type. The `overrides` object at
the top of each file is the only **input**: `defaults` forces the class a bare name builds and
`aliases` renames a scope-qualified name. WML needs four defaults and five aliases; every other
namespace needs none. The naming rules and the reasoning are CR-001 section 14.2.
