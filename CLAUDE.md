# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

docx4j-python is the Python port of docx4j (Java, sibling checkout `../docx4j`). The OOXML object
model is **generated** from docx4j's own schema tree by a fork of xsdata; the engine (Open
Packaging, parts, load and save, MCE) is **hand written** over it. The design follows
docx4j-core-ts (`../docx4j-core-ts`) so the two engines stay recognisably one design, and Java
docx4j is the behavioural oracle: names, defaults and quirks are docx4j's unless a CR records a
departure.

The design lives in `docs/change-requests/`. Each CR's implementation-notes sections (CR-001
sections 12 to 14, CR-002 section 12, CR-003 sections 10 to 14) record what was actually built,
the numbers and every deliberate departure; read them before changing the area they cover.
Status: CR-001 (object model) Phases A to C implemented, D proposed; CR-002 (engine) Phase A
implemented, B (resolution utilities, the Java parity harness and golden files) and C (PML, SML,
flat OPC, less the flat OPC read half) proposed; CR-003 (content API) Phases A, B, C, D and K
implemented, E to J proposed.
`REPORT.md` is the original xsdata feasibility experiment.

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
.venv-fork/bin/python scripts/threads.py               # the thread-safety check
.venv-fork/bin/python scripts/acceptance.py            # writes out/acceptance/, the six Word checklist documents

.venv/bin/python scripts/roundtrip.py                  # the baseline, upstream xsdata
python codegen/derive_names.py                         # re-derive codegen/names/ from docx4j's XJC output
```

`pyproject.toml` sets `addopts = "-q --strict-markers"` and turns `DeprecationWarning` into an
error; the only marker is `slow`. `scripts/roundtrip.py` exits non-zero on any parse failure, any
real difference or any skipped content. `codegen/generate.sh` puts `.venv-fork/bin` on `PATH`
itself, because the generator shells out to `ruff` by bare name. `derive_names.py` reads
`../docx4j/docx4j-generated-objects/target/generated-sources/xjc` by default
(`codegen/names/README.md`).

`scripts/reproduce.sh` reproduces `REPORT.md` with upstream xsdata and dates from before Phase B:
it runs `rm -rf docx4j_py` and generates the clusters layout there, which would delete the
hand-written code. Do not run it as it stands.

`tests/README.md` maps each test file to its CR section and holds the manual Word acceptance
checklist, with the record of each run. After any change to marshalling, the prefix table, content
types, the zip writer, `create_package`, the markdown importer or the content API's inserts,
regenerate `out/acceptance/` and ask for a Word check: saved output must open without a repair
prompt.

## Generated and hand-written code

Almost everything under `docx4j_py/` is generated and owned by `codegen/generate.sh`. The
hand-written files are exactly those in `codegen/clean.py`'s `KEEP`, listed by path:

- `child.py`: `Child`, `ChildList`, `link_parents`, `iter_children` (with the `mce` mode),
  `deep_copy`
- `namespaces.py`: docx4j's prefix table (`UNDERSTOOD` in it is re-exported from generated
  `el_index.py`)
- `runtime.py`: the shared `XmlContext`, parser and serialiser factories, `warm_up()`
- `fragments.py`: `wml(...)`, `to_xml(...)`
- `traversal.py`: `walk`, `iter_nodes`, `find`, `text_of`, `element_name`
- `wml/builders.py`: the `p` / `r` / `t` / `tbl` / `br` / `tab` sugar (with `wml/pictures.py`
  and `wml/sdt.py` beside it)
- `model/`: CR-003's content API — `model/content/` (the `Body`, `Paragraph`, `Range`, `Font`,
  `Table`, `TableRow`, `TableCell`, `InlinePicture` and `ContentControl` views, `insert_ooxml`,
  addresses, `Outline`, `describe()`, `ChangeReport`, `dry_run`), `model/markdown/` (markdown out
  and in) and `model/sessions.py` (`DocumentSession`)
- `resources/`: the parts `warm_up()` parses, and docx4j's default styles, numbering and fontTable
- `openpackaging/`: the whole engine

**Never hand-edit a generated file**: the namespace packages' `__init__.py` and `el.py`
(`wml/`, `dml/`, `math/`, `oart/`, `docprops/`, `relationships/` and the rest),
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
  in `PATCHES.md` with its marker. `docx4j_python__ROOT.xsd` is ours: the single entry point that
  collects `wml/wml.xsd`, `relationships.xsd` and the `docProps/` schemas.
- **The name tables** (`codegen/names/*.json`). Schema type and element names mapped to docx4j's
  Java class names, derived mechanically by `derive_names.py`; only `wml_overrides.json` is hand
  written (the scoped `t@r` and friends, giving `RT`, `RInstrText`, `RDelInstrText`). Tables are
  matched by the `namespace` key inside each file, not by file name. `Meta.name` always stays the
  schema's name. The expected ~225 `ClassNames` warnings are explained in `codegen/names/README.md`
  and CR-001 section 13.3.
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
docProps parts); `packages/` (`OpcPackage`, `WordprocessingMLPackage`, `create_package`);
`load.py` (Load3) and `save.py` (Save); `mce.py`, the optional preprocessor; `api.py`, `load()`
and `create_package()`. Everything is synchronous. Departures from docx4j are listed in CR-002
section 12.3.

## Round-trip and fidelity rules

- A part that is never unmarshalled is written back **byte for byte** from the source store.
  Reading `contents` marks a part for re-marshalling. There is no dirty tracking (docx4j's rule).
- An unmarshalled part saved unchanged must be **canonically identical** to the source
  (`scripts/canon.py`'s prefix-insensitive normal form). The one tolerated category is
  `boolean-spelling` (`1` against `true`). Every other difference category must be zero.
- **Nothing may be dropped silently.** Loading is lenient, with a per-part report (`part.skipped`,
  `package.skipped`); `LoadOptions(strict=True)` and `wml(...)` (strict by default) raise. Empty is
  the norm for WordprocessingML over `samples/`. A non-empty report is a fidelity bug, usually
  fixed by a marked schema patch rather than by loosening the check.
- Booleans serialise numeric (`1`/`0`), as Word writes them; the docProps parts use
  `bool_format = "words"`. On write, `XmlPart.xml` re-declares every prefix that `mc:Ignorable`
  names.
- MCE is lossless by default: both branches of `mc:AlternateContent` are kept as wildcard
  (`AnyElement`) trees, and traversal takes `mce="resolve"` (the default), `"all"` or `"none"`.
  `LoadOptions(mce_preprocess=True)` replaces each one with the branch Word would take, which is
  the only way to get typed content such as `Drawing` inside a choice. `UNDERSTOOD` is generated
  and deliberately narrower than the prefix table.
- `tests/openpackaging/test_round_trip.py` and `tests/test_roundtrip.py` hold these over the
  16-document corpus in `samples/` (docx4j fixtures, Apache-2.0). The known chart and
  spreadsheet-drawing `mc:AlternateContent` gap (CR-002 section 12.5) is pinned so that closing it
  is visible.

`out/`, `baseline/`, `.venv*` and `.claude/` are git-ignored; `out/` is rebuilt by the scripts.

## Portfolio task registry

This repository's change requests are indexed, with their dependencies on work in the other
docx4j repositories, in `../docx4j-portfolio/tasks.yaml` (ids `<repo>/<CR>[.<phase>]`; this
repository's key is `python`).

- When a CR's status changes (a phase lands; a CR is proposed, deferred or abandoned) or its
  **Depends on** changes, update the matching entry in `tasks.yaml` in the same session (`status`,
  `depends_on`; add an entry for a new CR or phase).
- Then run `python3 ../docx4j-portfolio/scripts/tasks.py check`. It reports `CHANGED` for each CR
  whose Status line was edited; once the registry entry agrees, run `tasks.py accept` (and
  `tasks.py graph` if dependencies changed).
- Before starting a CR or phase, check `python3 ../docx4j-portfolio/scripts/tasks.py blocked`: it
  may be waiting on work in another repository.
