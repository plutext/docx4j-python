# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

docx4j-python is the Python port of docx4j (Java, sibling checkout `../docx4j`). The OOXML object
model is **generated** from docx4j's own schema tree by a fork of xsdata; the engine (Open
Packaging, parts, load and save, MCE) is **hand written** over it. The design follows
docx4j-core-ts (`../docx4j-core-ts`) so the two engines stay recognisably one design, and Java
docx4j is the behavioural oracle: names, defaults and quirks are docx4j's unless a CR records a
departure.

The design lives in `docs/change-requests/`, indexed in its `README.md`. Each CR's
implementation-notes sections record what was actually built, the numbers and every deliberate
departure; read them before changing the area they cover. Status, 2026-10-24:

- **CR-001** (the object model): Phases A to C implemented; D (PML and SML) proposed.
- **CR-002** (the engine): Phases A and B implemented. B (section 15) is `PropertyResolver` and
  `StyleUtil` over a property catalogue, the numbering emulator completed and
  `org.docx4j.fonts`' decision half, with **zero differences on all 46 of docx4j's parity
  goldens with a fixture here**, which this repository consumes from
  `../docx4j-core-ts/test/golden/` and never builds. C (PML, SML, flat OPC, less the flat OPC read
  half) proposed. The xsdata fork's CR-001 (`~/git/docx4j-xsdata/docs/fork/CR-001-serialiser.md`)
  is the throughput work: its Phase A is this engine's `XmlPart._marshal` (CR-002 section 12.13,
  `Symbols.docx` 887 ms to 484 ms), and B to D are the fork's.
- **CR-003** (the content API): every phase implemented (sections 10 to 20; 21 is
  `Range.hyperlink`, 22 docx4j CR-043's hand-off). With Phase E it covers every non-extension
  member of `tests/office_js_subset.json`; with Phase J, `docx4j_py.docx` covers 305 of the 312
  members of `tests/python_docx_subset.json`.
- **CR-004** (PresentationML, a python-pptx facade), **CR-005** (SpreadsheetML, an openpyxl
  facade) and **CR-006** (the MCP server's contract; code in `../docx4j-mcp-py`): proposed.
- **CR-007** (PyPI): closed 2026-09-20. `docx4j 0.1.0` (import name `docx4j_py`) is on PyPI over
  `docx4j-xsdata 26.2.1`, both published by trusted publishing from a GitHub release
  (`RELEASING.md`; the fork's is its `REBASING.md` "Releasing"). The gates are
  `scripts/release_check.sh` plus the Word check. `pyproject.toml` depends on
  `docx4j-xsdata[lxml]>=26.2.1,<26.3`; `CHANGELOG.md` is the record, and `docs/pypi-readme.md`
  the PyPI page, every fence executed by `tests/test_readme.py`.
- **CR-008** (table conditional formatting): proposed. **CR-009** (theme colours): Phase A
  implemented; B implemented 2026-10-24 (`model/color/`, the WordprocessingML half of docx4j's
  `org.docx4j.model.color` at `a9e0303ff`; `PropertyResolver.theme_colors` and `resolve_color`;
  `Font.color`'s effective view reads a theme colour as Word draws it, the direct view `w:val`;
  docx4j's vectors in `tests/fixtures/theme-colors/`); C proposed.
- **CR-010** (tracked changes as Word resolves them, `model/revisions/`): complete, 2026-10-21.
- **CR-011** (one text model, `model/text/`): Phases A to D implemented, 2026-10-21. E, the HTML
  reader, is proposed. Its fork change (stage 10) was released in docx4j-xsdata 26.2.2
  (2026-10-22).
- **CR-012** (`mc:AlternateContent` kept at every level, its branches typed): decided 2026-10-21;
  Phases A and A2 (docx4j CR-051 phases 1 and 2) implemented and Word-checked; B (docx4j's 61
  root elements typed in a branch) and C (VML in the model, typed wherever it stands, with the
  `w:object` order fix) implemented and Word-checked 2026-10-22. Its fork changes (stages 11 and
  12) were released with CR-011's in docx4j-xsdata 26.2.2.
- **CR-013** (ranges across paragraphs, the port of docx4j-core-ts CR-014 on Office JS's
  characters): complete, 2026-10-23. A, Office JS's characters (`\r` marks, U+000B breaks, Word's
  own writers); B, a `Range` with two ends and the reads; C, the writes (`span_edits.py`,
  core-ts's phases 2 and 3 and check 57); D, the records and CR-003 section 23.4's review point
  (keep parity). Each phase that changed output is Word-checked. The Office JS subset is read from
  core-ts's committed `HEAD` (246 members).
- **CR-014** (content control data binding; phases 1, 1b and 1c from docx4j CR-058) and **CR-015**
  (Markdown output safety): proposed. **CR-016** (core-ts's cell, font, placeholder and wildcard
  changes): decided 2026-10-24. Phase A (a table cell ends with a paragraph, as Word reads it:
  `model/content/table_rules.py`, docx4j's `TableRules`) implemented 2026-10-24; B implemented
  2026-10-24 (`Font` reads every run in scope, `None`/`""`/`"Mixed"` where they differ, a paragraph
  its text's runs, else a placeholder's, else its mark, an automatic colour `#000000` in the
  effective view, a read splitting nothing; a run in a control showing its placeholder is no text to
  every Office JS reader, `hidden_runs_of`; `ContentControl.insert_text` writes in its place and
  clears the flag; `paragraph.text = ...` over a run-level one is refused, `range.holds_placeholder`;
  no `\r` after a deleted mark in the default view, none after an inserted one in the original); C
  implemented 2026-10-24 (Word's wildcards match a paragraph's mark where `^p` does, `^13` is one
  and a wildcard pattern `re.DOTALL`; a paragraph searches its `"Whole"` range, a body or block
  control its `"Content"`; a wildcard `replace_text` refuses a hit holding a mark,
  `range.replaces_mark`; no search hit cuts a grapheme cluster; `find` stays a paragraph at a
  time); D done 2026-10-24 (the records, and the Word check of A to C together, passed): complete.
- **CR-018** (run properties read as Word 365 reads them, the counterpart of docx4j CR-060 and
  CR-061; the fork's stage 13, released as docx4j-xsdata 26.2.3): decided 2026-10-24; Phases A and
  C done (Word-checked); B and D done 2026-10-24 (Word-checked): complete. **0.2.0 waits for CR-016, CR-018 and CR-009 Phase B** (the
  user, 2026-10-24); all three are done. **It also waits for core-ts's CR-002 section 63** (Word's writes into a
  control showing its placeholder, checks 63 and 64) **and its port here** (the user, 2026-10-25; CR-016 section 13.5). core-ts implemented it at `020c039`; the port is
  **CR-019** (proposed 2026-10-25: A the writes, B placeholders in Word's form and the glossary, C the records and the Word check).

`REPORT.md` is the original xsdata feasibility experiment. WordprocessingML only until CR-001 D
and CR-002 C.

## Environments and commands

Two virtual environments, neither committed:

- `.venv-fork`: Python 3.14 with the fork `docx4j-xsdata` installed editable from
  `~/git/docx4j-xsdata` (branch `docx4j`, local only, nothing pushed; its `docs/fork/CHANGES.md`
  lists every difference from upstream and `docs/fork/REBASING.md` the procedure). The generated
  bindings import `docx4j_xsdata`, so **everything except the baseline runs with this one**.
- `.venv`: upstream `xsdata[cli,lxml]==26.2`, only for the pre-Phase-B baseline in
  `baseline/docx4j_py/generated/` (`pyproject.toml` extra `baseline`).

```bash
.venv-fork/bin/python -m pytest                        # the suite
.venv-fork/bin/python -m pytest -m "not slow"          # without the corpus round trip, el regeneration and timings
.venv-fork/bin/python -m pytest tests/openpackaging    # the engine (CR-002)
.venv-fork/bin/python -m pytest tests/test_child.py::TestChildBase::test_parent_is_a_slot_not_a_field

codegen/generate.sh                                    # regenerate docx4j_py/ from schemas/
codegen/generate.sh --check                            # regenerate twice, prove the output byte identical

.venv-fork/bin/python scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata
.venv-fork/bin/python scripts/checks.py out/phase-b    # targeted fidelity checks over roundtrip.py's artefacts
.venv-fork/bin/python scripts/parents.py               # parent pointers over the corpus, and their cost
.venv-fork/bin/python scripts/mce_branches.py survey   # CR-012 Phase B: what a kept mc branch holds (corpus, readers)
.venv-fork/bin/python scripts/vml_census.py run --out out/x   # CR-012 Phase C: VML in every part, census and saves (then compare two runs)
.venv-fork/bin/python scripts/threads.py               # the thread-safety check
.venv-fork/bin/python scripts/bench_marshal.py         # parse and marshal, part by part (--json)
.venv-fork/bin/python scripts/acceptance.py            # writes out/acceptance/, the twelve Word checklist documents
.venv-fork/bin/python scripts/generate_fonts.py        # docx4j_py/resources/fonts/ from ../docx4j
scripts/release_check.sh                               # CR-007: build the wheel and sdist into dist/, twine check, a clean
                                                       # 3.12 venv with the wheel and the fork from PyPI, the smoke script,
                                                       # the fast suite against the wheel, out/acceptance/ from the wheel
scripts/compare_acceptance.py out/acceptance-checkout out/acceptance   # canonical, dates scrubbed

.venv/bin/python scripts/roundtrip.py                  # the baseline, upstream xsdata
python codegen/derive_names.py                         # re-derive codegen/names/ from docx4j's XJC output
```

`pyproject.toml` sets `addopts = "-q --strict-markers"` and turns `DeprecationWarning` into an
error; the only marker is `slow` (the corpus round trip, the el regeneration and, since CR-007,
`tests/test_performance.py`, whose thresholds are checkout-machine numbers). `scripts/roundtrip.py`
exits non-zero on any parse failure, any real difference or any skipped content.
`codegen/generate.sh` puts `$DOCX4J_PY_VENV/bin` (default `.venv-fork/bin`) on `PATH` itself,
because the generator shells out to `ruff` by bare name; CI sets `DOCX4J_PY_VENV` to the
interpreter's prefix. `derive_names.py` reads
`../docx4j/docx4j-generated-objects/target/generated-sources/xjc` by default
(`codegen/names/README.md`).

CI (`.github/workflows/`, CR-007 section 4.4): `test.yml` runs the fast suite on 3.12, 3.13 and
3.14 with the fork **from PyPI** and the full suite plus `generate.sh --check` on 3.14;
`publish.yml` publishes to PyPI from a GitHub release of tag `v<version>` by trusted publishing
(TestPyPI on `workflow_dispatch`). `RELEASING.md` is the procedure; never tag, push or publish
from a session --- those are the user's.

`scripts/reproduce.sh` reproduces `REPORT.md` with upstream xsdata: it regenerates the
git-ignored baseline in `baseline/docx4j_py/generated/` (running xsdata inside `baseline/`, since
xsdata writes relative to the working directory) and never touches `docx4j_py/`.

`tests/README.md` maps each test file to its CR section and holds the manual Word acceptance
checklist, with the record of each run. `tests/test_readme.py` executes every ```` ```python ````
fence of `docs/pypi-readme.md` and `README.md` in order (a README fence that is a fragment is
marked `<!-- fragment -->` on the line before it; the PyPI readme may carry no marker), so a
readme example that stops working fails the suite. After any change to marshalling, the prefix table, content
types, the zip writer, `create_package`, the markdown importer or the content API's inserts,
regenerate `out/acceptance/` and ask for a Word check: saved output must open without a repair
prompt. Markup that makes Word hang or quietly refuse an operation is logged in
`../docx4j-portfolio/docs/word-hangs.md`; add to it after any check that finds one.

## Generated and hand-written code

Almost everything under `docx4j_py/` is generated and owned by `codegen/generate.sh`. The
hand-written files are exactly those in `codegen/clean.py`'s `KEEP`, listed by path:

- `py.typed`: the PEP 561 marker, empty (CR-007)
- `child.py`: `Child`, `ChildList`, `link_parents`, `iter_children` (with the `mce` mode),
  `deep_copy`
- `namespaces.py`: docx4j's prefix table (`UNDERSTOOD` in it is re-exported from generated
  `el_index.py`)
- `runtime.py`: the shared `XmlContext`, parser and serialiser factories, `warm_up()`
- `fragments.py`: `wml(...)`, `to_xml(...)`
- `traversal.py`: `walk`, `iter_nodes`, `find`, `text_of`, `element_name`
- `mc_declarations.py`: where a save declares the prefixes an `mc` attribute names, on docx4j
  CR-052's placement (CR-011 Q11)
- `wml/builders.py`: the `p` / `r` / `t` / `tbl` / `br` / `tab` sugar (with `wml/pictures.py`
  and `wml/sdt.py` beside it)
- `model/`: CR-002 Phase B's resolution utilities — `property_catalogue.py`, `style_util.py`,
  `property_resolver.py`, `listnumbering/` and `fonts/` (`RunFontSelector`, the `Mapper` family,
  `FontFallback`, `fonts_in_use`) — and CR-003's content API: `model/content/` (the `Body`,
  `Paragraph`, `Range`, `Font`,
  `Table`, `TableRow`, `TableCell`, `InlinePicture`, `ContentControl`, `Comment` and
  `TrackedChange` views, `insert_ooxml`, addresses, `Outline`, `describe()`, `ChangeReport`,
  `dry_run`, `Author`, the `ChangeTracker`), `model/customxml/` (`CustomXmlPart`,
  `CustomXmlNode`, `XmlMapping`, the typed control kinds, the bindings, `describe()` / `fill()`),
  `model/markdown/` (markdown in; markdown out is the content API's readers and their address
  comments over the text model's `markdown`) and `model/sessions.py` (`DocumentSession`). Then
  CR-010's `model/revisions/` (the tracked-changes engine: `tracked_changes`, `accept_all`,
  `reject_all`, `report`) and CR-011's text model:
  `model/text/` (`extract`, `TextOptions`, `TextResult`, `plain_text`, `markdown`, the
  commonmark-java renderer port), with Phase C's `model/chart/` (a chart as a table and a
  description), `model/math/` (OMML to LaTeX), `model/fields/` (field formatting switches) and
  `model/java_text.py` (Java's `DecimalFormat` and `SimpleDateFormat`, for those two); CR-009's
  `model/color/` (`ResolvedColor`, `ThemeColors`, `WordColors`, whose `normalise` and `stated` are
  `values.py`'s; `theme_colors._slot_colour` is Phase C's seam for `DrawingMLColors`); and
  CR-018's `model/values.py` (Word's on/off, enumeration and colour readers)
- `resources/`: the parts `warm_up()` parses, docx4j's default styles, numbering and fontTable,
  `understood-namespaces.txt` (docx4j's published understood list, copied at `198d015a3`; CR-011),
  and `resources/fonts/` — docx4j's substitution tables as committed JSON with the three Office
  theme parts, written by `scripts/generate_fonts.py` and never hand-edited
- `docx/`: CR-003 Phase J's python-docx facade (`Document`, `Paragraph`, `Run`, `Font`, `Table`,
  `Section`, the `Length` units and the `WD_*` enums), a separate namespace with python-docx's names
  over the content API; `scripts/python_docx_subset.py` derives its member list from an installed
  python-docx 1.2.0, which is never a dependency
- `openpackaging/`: the whole engine

**Never hand-edit a generated file**: the namespace packages' `__init__.py` and `el.py`
(`wml/`, `dml/`, `math/`, `oart/`, `docprops/`, `relationships/`, `vml/` and the rest),
`docx4j_py/__init__.py` (it carries the import-order manifest and the CR-002 re-export footer) and
`docx4j_py/el_index.py`. `clean.py` deletes everything not in `KEEP` before each run. A new
hand-written file under `docx4j_py/` must be added to `KEEP`, or the next regeneration deletes it.
Change the output by changing the inputs below, regenerate, and run `generate.sh --check`.

## How regeneration works

`generate.sh` runs `codegen/clean.py`, then
`docx4j-xsdata generate -c .xsdata.phase-b.xml schemas/docx4j_python__ROOT.xsd`, then
`codegen/generate_el.py`. The inputs, and the only places a change to the model belongs:

- **The fork** (`~/git/docx4j-xsdata`). Generator options in `.xsdata.phase-b.xml`:
  `AllOptional` (every field `None | T = None`), `SchemaDefaults=metadata` (the schema default in
  field metadata, never filled on parse), `ListFactory=docx4j_py.child.ChildList`,
  `DeferredImports` (makes the one-module-per-namespace layout importable), `ClassNames`,
  `Structure=namespaces`, `UnnestClasses`, an extension that makes every class a `Child`, and a
  `Substitutions` table mapping each namespace to its `docx4j_py.*` module. The runtime adds the
  skipped-content report (`ParserConfig(skipped_report=True)`) and numeric booleans. A generator
  or runtime bug is fixed in the fork, with tests under its `tests/fork/`, and recorded in
  `docs/UPSTREAM.md` for later reporting upstream.
- **The schema copy** (`schemas/`). docx4j's `xsd/` tree copied verbatim apart from the patches in
  `schemas/PATCHES.md`, each marked in the XSD by a comment beginning `docx4j-python`
  (`grep -rn 'docx4j-python' schemas/`). `default="..."` is never stripped. Record any new patch
  in `PATCHES.md` with its marker. `PATCHES.md` names the docx4j commit the copy is from; a
  refresh re-copies the closure verbatim and re-patches by marker (CR-001 section 14.10 is the
  precedent, with the checks it made). `docx4j_python__ROOT.xsd` is ours: the single entry point that
  collects `wml/wml.xsd`, `relationships.xsd` and the `docProps/` schemas.
- **The name tables** (`codegen/names/*.json`). Schema type and element names mapped to docx4j's
  Java class names, derived mechanically by `derive_names.py`; only `wml_overrides.json` is hand
  written (the scoped `t@r` and friends, giving `RT`, `RInstrText`, `RDelInstrText`). Tables are
  matched by the `namespace` key inside each file, not by file name. `Meta.name` always stays the
  schema's name. The expected ~280 `ClassNames` warnings are explained in `codegen/names/README.md`
  and CR-001 section 13.3. `root_elements.json` (CR-012 Phase B), also `derive_names.py`'s, lists
  docx4j's root elements this model does not declare globally; `generate_el.py` writes it into
  `el_index.py` as `ROOT_ELEMENTS` and `ROOT_ELEMENT_BINDINGS`.
- **The `el` tables** (`codegen/el_tables/<namespace>.json`). Written by `generate_el.py` from the
  `XmlContext` metadata, never from the schema, and committed for review. The `overrides` object at
  the top of each file (`defaults`, `aliases`) is the only input; the rest is output. The naming
  rules are CR-001 section 14.2.

`generate_el.py` also turns each generated namespace module into a package of the same name (so
`el` is `docx4j_py.wml.el`) and writes `el_index.py`. The module name never changes. This is why
`KEEP` holds paths: `clean.py` leaves `wml/` holding only `builders.py`, with no `__init__.py`.

## The object model

- Every generated dataclass is a slotted `Child`. `parent` is a slot, not a field, so equality,
  `repr` and serialisation ignore it. `link_parents` wires a parsed tree from class metadata;
  `ChildList` (the `default_factory` of every list field) sets the parent on append, insert,
  extend and assignment, as docx4j's `ArrayListDocx4j` does. Assigning a single-valued field
  (`p.p_pr = ppr`) sets no parent until the next `link_parents`.
- `deep_copy(obj, parent=None)` copies fields only and relinks the copy's descendants. Plain
  `copy.deepcopy` through the parent slot would copy the whole document.
- `el.<elementName>(...)` is the object factory, one function per element name per namespace. Use
  it rather than guessing a class: 72 intermediate choice classes (`RT` for `w:t` in a run) have no
  element name of their own, and some names map to several classes (`el.sdt` is `SdtBlock`,
  `el.sdt_run` is `SdtRun`). `to_xml` needs the same table to name a root.
- Importing anything under `docx4j_py` imports the whole model (about 0.6 s), because
  `docx4j_py/__init__.py` carries the import-order manifest. The engine must not import a model
  class at module level: typed parts name their class as a string (`model_class_path`), and
  `tests/openpackaging/test_threads_and_import.py` enforces it.
- Threads: one `XmlContext` per process, one `ParserConfig` (so one `XmlParser`) per parse.
  `ParserConfig.skipped` is per-parse state, and sharing a config corrupts the skipped-content
  report (CR-001 section 14.7). `runtime.parser()` returns a new one on every call.
- Don't use `XmlMeta.qname` for element names (it caches whichever parent namespace it met first);
  use `element_name`.

## The engine

`docx4j_py/openpackaging/` mirrors docx4j's `org.docx4j.openpackaging` and `io3`, layer by layer:
`PartStore` / `PartSink` (zip, directory, memory) in `stores.py`; `PartName`; `ContentTypeManager`;
`parts/` (`Part`, `BinaryPart`, `XmlPart[T]` with lazy `contents`, `DefaultXmlPart` as an lxml tree,
`RelationshipsPart`, `PartRegistry` mapping content type to part class, the typed WML, DML and
docProps parts --- `parts/wml/` is a package, whose `comments.py` is the part half of CR-003
Phase G); `packages/` (`OpcPackage`, `WordprocessingMLPackage`, `create_package`);
`load.py` (Load3) and `save.py` (Save); `mce.py`, the optional preprocessor; `api.py`, `load()`
and `create_package()`. Everything is synchronous. Departures from docx4j are listed in CR-002
section 12.3.

## Round-trip and fidelity rules

- A part that is never unmarshalled is written back **byte for byte** from the source store.
  Reading `contents` marks a part for re-marshalling. There is no dirty tracking (docx4j's rule).
- A table cell ends with a paragraph, as Word reads it, after every content-API edit (CR-016
  Phase A: `ends_with_paragraph` and `end_with_paragraph`, through `table_rules.py`). A file that
  already breaks the rule is never repaired on load or save, only when an edit touches that cell.
- An unmarshalled part saved unchanged must be **canonically identical** to the source
  (`scripts/canon.py`'s prefix-insensitive normal form). The one tolerated category is
  `boolean-spelling` (`1` against `true`). Every other difference category must be zero. The one
  exception is a load in `LoadOptions(mc_load_mode="word")` (CR-012 section 4.1 item 3): a part
  holding a block-, row- or cell-level `mc:AlternateContent` is saved as Word's own save writes it,
  the branches Word does not read gone. The default, `preserve`, keeps the rule. Since CR-018 Phase A
  a second category is tolerated, `word-normal-form`: the parse is exact (`runtime.parser_config()`
  sets the fork's `exact_types={bool, Enum}`, so a value not exactly a lexical form is kept as its
  text and listed in `part.kept`), and a save normalises by default (`package.save(normalise=True)`,
  `LoadOptions(normalise_on_save=...)`), writing each kept value in the form Word reads as it read
  the source (an on/off value Word reads as on bare, one it reads as off `0`; since Phase B every
  enumeration and colour too: a colour as upper-case `RRGGBB`; a rejected text or underline colour,
  a missing or invalid `w:highlight`/`w:vertAlign`/`w:em` and a bare `<w:u/>` dropped; a rejected
  fill or border colour `auto`; `w:u`'s on/off words `none`, those rows found by name in each part's
  render, `normalise.normalise_render`, so the tree keeps what was read), recorded in
  `package.last_save.normalised`. `canon.compare(..., normalise=True)` writes the same normal form
  into the source before comparing, so every other category must still be zero;
  `save(normalise=False)` keeps the source's spelling and is held to `boolean-spelling` alone. Every
  `ParserConfig` in the tree is `parser_config()`'s (`tests/test_parser_configs.py`), and every
  on/off, enumeration and colour reader goes through `docx4j_py/model/values.py` (Word's rules: off
  for exactly `false`, `off` or `0`; an enumeration value outside it unset; a colour six hex digits,
  a classic HTML name, or exactly `auto`).
- **Nothing may be dropped silently.** Loading is lenient, with a per-part report (`part.skipped`,
  `package.skipped`); `LoadOptions(strict=True)` and `wml(...)` (strict by default) raise. Empty is
  the norm for WordprocessingML over `samples/`. A non-empty report is a fidelity bug, usually
  fixed by a marked schema patch rather than by loosening the check.
- Booleans serialise numeric (`1`/`0`), as Word writes them; the docProps parts use
  `bool_format = "words"`. VML's `t`/`f` are a string enumeration, written as read. A parsed
  float or bytes value keeps its source spelling (the fork's `lexical_types`, stage 12;
  `runtime.LEXICAL_TYPES`): `o:spt="202"`, Word's wrapped `o:gfxdata` and a lower-case `w:color`
  are saved as read, a value set in code in the converter's spelling. On write, `XmlPart.xml`
  re-declares every prefix that `mc:Ignorable` names.
- MCE is lossless by default: every branch of `mc:AlternateContent` is kept, and traversal takes
  `mce="resolve"` (the default), `"all"` or `"none"`. Since CR-012 Phase B a branch's children are
  typed where docx4j types them: docx4j's 61 WordprocessingML root elements (`w:drawing`,
  `w:pict`, `w:t`, `w:tbl` and the rest), bound in a branch as the container holding the element
  binds them, elsewhere in a lax wildcard as their root class, through the fork's stage 11 hooks
  (`runtime.branch_child_class` and the serialiser's `element_names`); any other name stays a
  wildcard (`AnyElement`) tree. Since Phase C VML (`v:`, `o:`, `w10:`, in `docx4j_py.vml.*`) is in
  the model, typed wherever it stands: under `w:pict` and `w:object`, in a branch or not, and in
  `settings.xml`'s shape defaults. A vmlDrawing part (`VMLPart`) stays an lxml tree.
  `LoadOptions(mce_preprocess=True)` replaces each element with the branch Word would take. Since
  CR-012 Phase A (Q8),
  it and `mc_load_mode="word"` select over docx4j's published understood list
  (`resources/understood-namespaces.txt`), matched by namespace URI, so they take the branch Word
  365 takes. `UNDERSTOOD` is generated, deliberately narrower than the prefix table, and stays
  the prefix table's own list.
- Block-, row- and cell-level `mc:AlternateContent` (CR-012 Phase A): `LoadOptions(mc_load_mode=
  "preserve")`, the default, keeps the element whole. The readers read Word's branch at block
  level, and an edit inside a kept branch is refused (`address.read_only`). `"word"` replaces it
  on load with Word's branch and keeps what it took out in `part.resolved_alternate_content`.
- `tests/openpackaging/test_round_trip.py` and `tests/test_roundtrip.py` hold these over the
  16-document corpus in `samples/` (docx4j fixtures, Apache-2.0). The known chart and
  spreadsheet-drawing `mc:AlternateContent` gap (CR-002 section 12.5) is pinned so that closing it
  is visible.

`out/`, `dist/`, `baseline/`, `.venv*` and `.claude/` are git-ignored; `out/` and `dist/` are rebuilt by the scripts.

## Portfolio task registry

This repository's change requests are indexed, with their dependencies on work in the other
docx4j repositories, in `../docx4j-portfolio/tasks.yaml` (ids `<repo>/<CR>[.<phase>]`; this
repository's key is `python`).

- When a CR's status changes (a phase lands; a CR is proposed, deferred or abandoned) or its
  **Depends on** changes, the matching entry in `tasks.yaml` changes in the same session
  (`status`, `depends_on`; a new entry for a new CR or phase).
- **Don't edit `tasks.yaml` directly** (the user's rule, 2026-10-17). Send the change to the
  docx4j-portfolio session, the one working in `~/git/docx4j-portfolio` (find it by that directory
  in `ListAgents`, not by its name's suffix, which changes): the entry id, the exact new
  `status`, `depends_on` and `notes` (or the whole entry), and any CR Status line to re-baseline
  in `scripts/status-snapshots.yaml`. It applies the change, runs `tasks.py check`, `graph` and
  `accept`, and commits. The same holds for `scripts/status-snapshots.yaml` and
  `docs/dependencies.md`.
- Only if no docx4j-portfolio session is running may this session edit them directly. Then commit
  only your own hunks, through a temporary index built from `HEAD` immediately before the commit
  (never the shared index, and check `HEAD` hasn't moved), and run
  `python3 ../docx4j-portfolio/scripts/tasks.py check` (then `accept`, and `graph` if dependencies
  changed).
- Before starting a CR or phase, check `python3 ../docx4j-portfolio/scripts/tasks.py blocked`: it
  may be waiting on work in another repository.
