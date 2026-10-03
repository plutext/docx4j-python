# CR-008: Table conditions in the `PropertyResolver` --- a paragraph's table context handed to the resolver, and the table style decided by name

**Status:** Proposed 2026-10-04, on docx4j CR-030's hand-off (its section 9). Nothing implemented.
Amended the same day for CR-030's D11 (docx4j `0349796f9`): a band size stated nowhere means no
banding (section 3.1).
Not scheduled; Jason decides. Three questions are open (section 9).
**Depends on:** docx4j CR-030 (done 2026-10-04: commits `3577e5181`, `ecaa95c1d`, `8c387f484`,
`446d65c5d`, `843ac12df`, and `0349796f9` for D11, on `VERSION_17_3_1`, local and not yet pushed) and this repository's
CR-002 Phase B (implemented, section 15), whose item 6 of section 15.1 --- "table conditional
formatting stays out, as docx4j's own `table-conditions` CR has it" --- this CR lifts. The parity
gate of Phase A also needs goldens which do not exist yet (section 6 and question 1).
**Counterpart:** docx4j's
`docs/developer/change-requests/CR-030-table-conditions-in-the-resolver.md` (section 4 the design,
section 5 the Word probes T1 to T8 and their readings, section 6 what each phase did, section 9 the
hand-off) and docx4j-core-ts's `CR-007-table-conditions-in-the-resolver.md` (proposed the same
day; the same port, so the two engines stay one design).

## 1. Why

CR-002 Phase B ported `PropertyResolver` as docx4j had it on 2026-09-19, and docx4j then had no
answer to "what does the table style give this paragraph": the conditional formats were applied by
a preprocess which rewrote the document (`ParagraphStylesInTableFix`) and again, separately, by
the table writers. So Phase B left them out, and everything here which reports effective
formatting --- `Paragraph`'s and `Font`'s effective reads, `RunFontSelector`, `fonts_in_use` ---
answers for a paragraph in a table cell as though it were outside the table. A header row which
is bold because the table style's `firstRow` says so reads as not bold. For an agent asking
`describe()` or `font.bold` of a cell, that is a wrong answer about the commonest kind of table.

docx4j CR-030 moved the rule into the resolver: the caller hands it the paragraph's table context,
as the numbering part is already handed to it, and the resolver composes the table style's
contribution between the document defaults and the paragraph style. It also measured, with Word,
two rules which change answers this package already gives (section 2). This CR is the port.

## 2. What already disagrees with docx4j (checked 2026-10-04)

One thing, in code Phase B shipped. `get_effective_table_style` and `reaches_default_table_style`
(`docx4j_py/model/property_resolver.py`) decide Word's built-in Normal Table by the default table
style's **id**, as docx4j did until 17.3.1 (from CR-015's probe `styles-table-default`). Word goes
by the style's **name** (CR-030 probes T1, T2 and T7), and docx4j now does:

| Case | docx4j at `843ac12df` (Word's reading) | This package today |
|---|---|---|
| A table naming no style, where the `w:default` table style is not named "Normal Table" | that style as written, with nothing beneath it (T7: its margins of 300, or none) | the built-in (108 twips left and right) |
| A chain which reaches a `w:default` style named otherwise | every style as written, no built-in | the built-in, with the default style's own layer skipped |
| A chain which reaches a style named "Normal Table" that is not the default | the built-in stands in for it (T2 (a): margins of 108) | that style as written, no built-in |
| Styles above a style named "Normal Table" in its `w:basedOn` chain | never reached: the walk ends at "Normal Table" | applied |

The usual document --- default table style `TableNormal`, named "Normal Table", based on nothing
--- gets the same answer either way, which is why docx4j-core-ts's re-check of the goldens this
package consumes found zero differences against `843ac12df` (its CR-007 section 3; 46 fixtures,
none with a case from the table). `Table.effective_table_style`, `reaches_default_table_style`,
`style` and `style_built_in` (`model/content/table.py`) inherit the rule and move with it.

Nothing else disagrees, because nothing else is here yet: no in-context resolution, and so no
[MS-DOCX] exception to get wrong (section 3.4).

## 3. What is ported

Names are docx4j's in `snake_case`, as everywhere in `docx4j_py.model`. A new hand-written
package `docx4j_py/model/table/` mirrors `org.docx4j.model.table` (`model` is already in
`codegen/clean.py`'s `KEEP` as a directory, so nothing is added there).

### 3.1 `TableStyleConditions` (`model/table/conditions.py`)

The condition arithmetic, as static functions over the generated classes: `Look` (from
`w:tblLook`, the attributes or the legacy `w:val` bitmask; `DEFAULT` is Word's 04A0), the band
sizes, `from_cnf` (the `w:cnfStyle` cache of a row, cell or paragraph), `at_position` (the
conditions of a cell from its row and grid column), `resolve` (the caches where written, else the
position, gated by the look), `applicable` (a style's `w:tblStylePr` entries for a condition set,
in ECMA-376-1 17.7.6 order, `PRECEDENCE`), `formats_text`, and `key`. The writers' helpers
(`row_conditions`, `h_band_rows`, `v_band_cols`, `conditional_tr_pr`, `conditional_tc_pr`,
`has_tbl_header`) are ported with it although nothing here calls them yet: they are small, and
CR-003's table views and a later layout consumer read rows and cells the same way.

**A band size stated nowhere means no banding** (CR-030 D11, docx4j `0349796f9`, after the
hand-off; Word's probes T9 `tables-banding-merged-row` and T10 `tables-banding-band-size`). Word
bands no row of a table where neither the style chain nor the table's own `w:tblPr` states a
`w:tblStyleRowBandSize`, whatever `w:tblLook` and the style's band conditions say; a size stated
by the table alone is enough. So `row_band_size` and `col_band_size` return **0** where absent
(docx4j took it as 1 until this commit), and `at_position`, `h_band_rows` and `v_band_cols` treat
0 as no banding; a `w:cnfStyle` cache stays authoritative. The columns follow the rows' rule,
which Word was not asked about. T10 also confirms that a row whose every cell continues a
vertical merge still counts for the bands after it. The port is from `0349796f9`, not from the
hand-off's `843ac12df`.

### 3.2 `TableContext` and `CellContext` (`model/table/context.py`)

- **`TableContext`**, from `resolver.table_context(tbl)`: the table style the table resolves to
  (section 3.3's rule), the look (the table's `w:tblLook`, else the style's, else 04A0), the band
  sizes, each row's index, and each cell's first grid column and span (`w:gridBefore`, then the
  spans, nested tables excluded). `for_paragraph(tr, tc, p_pr)` and `for_cell(tr, tc)` give a
  `CellContext`. **The resolver keeps no `TableContext`**: it is a reading of the content, the
  resolver's contract covers the styles part only, and the caller holds it while in the table.
- **`CellContext`**, a frozen dataclass: the table style id, the conditions, the applicable
  `w:tblStylePr` entries which format text, and a structured cache key (table style id, condition
  set) --- a frozen value with `__eq__` and `__hash__`, never a joined string.

### 3.3 The resolver

```python
resolver.table_context(tbl) -> TableContext
resolver.cell_context_of(p) -> CellContext | None

resolver.get_effective_ppr(p_pr, cell_context=None)
resolver.get_effective_rpr(r_pr, p_pr=None, cell_context=None)
resolver.get_effective_paragraph_mark_rpr(p_pr, cell_context=None)
resolver.get_table_style_chain(style_id)
resolver.applies_table_style_size_jc_exception() -> bool
```

Java's overloads become a keyword argument. **`cell_context=None` gives exactly today's answer**,
and so does a context whose `formats_text()` is false; the 45 goldens Phase B holds are the proof
that nothing moved.

    effective pPr = docDefaults.pPr + table(ctx).pPr + chain pPr(style) [+ exception] + direct
    effective rPr = docDefaults.rPr + level(table(ctx).rPr, chain rPr(style)) [+ exception]
                    + level(character style) + direct
    table(ctx)    = the table style's own pPr / rPr, then ctx's applicable conditions in order

The table and the paragraph style are two levels of the style hierarchy, so the twelve toggle
properties combine across them (`style_util`'s toggle rule, ECMA-376-1 17.7.3) rather than the
paragraph style's overriding. The composition is cached per (paragraph style id, `CellContext`
key), and each table style's merged `w:basedOn` chain per style id; `refresh()` clears both.
`get_effective_table_style` keeps its contract of a new `w:style` each call by copying the cached
chain before merging the table's own `w:tblPr` over it.

**The name rule** replaces section 2's: a table's style is its `w:tblStyle` if that names one,
else the `w:default` table style whatever its name. Walking the chain, a style *named* "Normal
Table" contributes Word's built-in (108 twips left and right, nothing to text) in place of its own
definition and the walk ends there; every other style applies as written, with nothing beneath
it. The built-in also stands in for a missing style and where there is no default table style. A
style with no `w:name` (Google Docs writes them) is not "Normal Table".

**`cell_context_of(p)`** climbs the parent pointers to the nearest `Tc`, then the nearest `Tr`,
then the nearest `Tbl`, through whatever lies between (`w:sdt`, `w:customXml`, `w:smartTag`). It
returns `None` at a story root reached first --- `CTTxbxContent`, a footnote or endnote, a
comment, `Hdr`, `Ftr`, `Body` --- which is Word's reading: a text box anchored in a cell takes
none of the table's formatting (T3), and neither does a note referenced from one (T8). It also
returns `None` where the chain breaks. That is rarer here than in Java, since `ChildList` sets
the parent on append; the case to document is a single-valued assignment before the next
`link_parents`. Each call builds a `TableContext`, a walk of the table's rows, so it is for one
paragraph; code resolving every paragraph uses the tracker.

**`TableContextTracker`** (`model/table/tracker.py`): what a walker composes to keep the table
stack, with an empty stack inside `w:txbxContent`. It sits over `traversal.walk` /
`iter_children` in the default `mce="resolve"` mode. A paragraph inside an unresolved
`mc:AlternateContent` is not typed content here (CR-002 section 12.5) and gets no context, which
is the existing limit of every typed read and not a new one.

### 3.4 The [MS-DOCX] exception

`overrideTableStyleFontSizeAndJustification`: where it applies, a paragraph style which is the
default paragraph style, or which states no size (justification) of its own, and which resolves to
12pt (left), gives way to the table level's size (justification). Measured with Word 365 (T5, T6):
it applies **below compatibility mode 15 where the setting is not on** --- a document stating no
`compatibilityMode` is mode 12 --- and **never in mode 15**, where Word ignores a stated 0 and
re-saves it as 1. Read once per `refresh()`.

The resolver reads the mode and the setting from the settings part without unmarshalling it, as
it reads the styles part (`read_contents()`, CR-002 section 15) and as `pkg.compatibility_mode`
already does through lxml (CR-003 section 17.11). The reading moves to where the resolver can
reach it without importing the content API; `model/content/compatibility.py` keeps its public
names over it.

Two consequences worth stating. A document from `create_package()` declares mode 15 (CR-002
section 12.10), so the exception never applies to a created document. And a resolver built from a
styles tree alone, with no package, has no settings to read: question 3.

### 3.5 Not ported

The synthetic style ids (`styleIdFor`, `syntheticStyle`, `sourceStyleOf`) exist for docx4j's HTML
and FO XSLT pathways, which rewrite the document; nothing here does, so a paragraph never names
one and `sourceParagraphStyleOf` reduces to the paragraph's own style. The preprocess, the FO
visitor and the table writers have no counterpart here.

## 4. Who passes the context

- **`RunFontSelector`** (`model/fonts/run_font_selector.py`): takes a `cell_context` beside the
  `p_pr` it already takes and passes it to `get_effective_rpr`, so the font decided for a run in
  a cell follows the table style's `w:rFonts`, bold and italic.
- **`fonts_in_use`** (`model/fonts/fonts_in_use.py`): today a walk for names, which adds the
  `w:rFonts` of every `w:tblStylePr` of each style in use whether or not a cell is under that
  condition. docx4j's `FontsAnalysis` now walks with the tracker and resolves each run in context
  (`843ac12df`; 38 of its 617 corpus documents change). `walk_stories` takes the tracker and the
  spans follow.
- **The content API**: `Paragraph`'s effective paragraph properties and mark properties
  (`model/content/paragraph.py`) and `Font`'s effective reads (`model/content/font.py`) pass
  `resolver.cell_context_of(p)`; a view over a paragraph resolves one paragraph at a time, which
  is what `cell_context_of` is for. `describe()` and anything else reading through those views
  follow. Writes stay direct, as now.
- **Unchanged, deliberately**: the numbering `Emulator` (it resolves a style id, and docx4j's own
  fix there, D1, was to stop synthetic ids reaching it); markdown out, which keeps ignoring table
  styles as docx4j's does (CR-030 decision 4); the python-docx facade's tri-state `Font`, which
  reports direct formatting as python-docx does.

This is a change of answers in a released surface (`docx4j 0.1.0`): effective reads in a table
cell start reporting what Word shows. It goes in `CHANGELOG.md` under the next minor version, with
the name rule.

## 5. Where each piece lives

| Piece | File |
|---|---|
| `TableStyleConditions`, `Look` | `docx4j_py/model/table/conditions.py` (new) |
| `TableContext`, `CellContext` | `docx4j_py/model/table/context.py` (new) |
| `TableContextTracker` | `docx4j_py/model/table/tracker.py` (new) |
| the name rule, the chain cache, `table_context`, `cell_context_of`, the `cell_context` argument, the exception | `docx4j_py/model/property_resolver.py` |
| the mode and the setting, read without unmarshalling | beside the resolver; `model/content/compatibility.py` over it |
| the callers | `model/fonts/run_font_selector.py`, `model/fonts/fonts_in_use.py`, `model/content/paragraph.py`, `model/content/font.py`, `model/content/table.py` (docstrings: "reaches the default" becomes the name rule) |

No generated file, schema, name table or `el` table changes. As everywhere under `model/`, the
generated classes are imported inside functions, never at module level.

## 6. Tests

- **Nothing moves without a context.** The existing golden tests
  (`tests/content/test_properties_goldens.py`, `test_fonts_goldens.py`,
  `test_numbering_goldens.py`) stay at zero differences, before and after.
- **Word's readings, as fixtures.** CR-030's probes T1 to T8 are built by
  `docx4j-layout-fidelity`'s `Corpus.java`, and their Word readings are in CR-030 sections 5.1 and
  5.3. The twelve `tables-*` documents are copied into `tests/fixtures/` with provenance, and a
  test per probe asserts the measured reading: T1, T2 and T7 for the name rule (text and cell
  margins), T3 and T8 for the story boundaries, T5 and T6 for the exception in modes 12, 14 and
  15, with the setting on, off and absent. The two banding probes of D11 (T9, T10) join them:
  no band size, no bands; the table's own size alone; the fully merged row counted. These hold the rules whether or not a golden exists.
- **Units** for the condition arithmetic: `w:gridBefore`, spans, band sizes other than 1 and absent (0, no banding), the
  look's gates, the `w:cnfStyle` caches against the position, nested tables, a row-level and a
  cell-level `w:sdt`.
- **Parity.** The Java harness which writes the goldens lives in `../docx4j-core-ts/test/java/`
  and this repository consumes its output (CR-002 section 12.12). CR-030 section 9 asks that the
  goldens gain, per table paragraph, the table style id, the condition key and the in-context
  `pPr`, `rPr` and mark `rPr`, and core-ts's CR-007 phase 1 adds fixtures which exercise the name
  rule. When those are committed there, `tests/content/goldens.py` reads the new entries and the
  bar is the one Phase B set: **zero differences**. Until then the gate is the probes above
  (question 1).
- **Content API**: a header-row cell under a `firstRow` condition reads bold through `Font` and
  through `describe()`; the same paragraph in a text box inside the cell does not.
- **Cost**: `get_effective_rpr` with a context, cached, measured against Phase B's 34 us without
  one; `cell_context_of` over a large table measured once, so the tracker's reason is a number.

No Word check is needed: nothing in this CR changes what is written.

## 7. Phasing and effort

| Phase | What | Days |
|---|---|---|
| A | `model/table/`, the name rule, the chain cache, `table_context`, `cell_context_of`, the `cell_context` argument, the exception; the probe fixtures and the units; CR-002 section 15.1 item 6 amended | 2 to 3 |
| B | the callers of section 4: `RunFontSelector`, `fonts_in_use`, the content API's effective reads; `CHANGELOG.md`; the readmes where they show an effective read | 1 |

The name rule (section 2) is the first commit of Phase A and stands alone: it is a correction to
shipped behaviour and needs none of the rest. The parity gate closes whichever phase is open when
the extended goldens arrive.

## 8. Risks

- **docx4j `VERSION_17_3_1` is not pushed**, so the oracle is a local tree. The commits are named
  above; if the branch is rebased before it is pushed, the hashes here and in the tests'
  provenance are re-stated.
- **The goldens are not this repository's to extend.** If core-ts's CR-007 is not scheduled, the
  in-context parity gate stays open and the probes are the only measure. They are Word's own
  readings, which is the stronger oracle, but they are twelve documents and not a corpus.
- **`cell_context_of` per paragraph is quadratic over a large table.** The content API resolves
  one paragraph per view, so `describe()` or `to_api_script` over every cell of a long table
  would pay it. Phase B measures this; if it shows, those walkers take the tracker.
- **Answers change under callers of 0.1.0.** They change towards Word, and only in table cells
  and for the four cases of section 2.

## 9. Open questions (recommendations for decision)

1. **Wait for core-ts's CR-007 goldens, or go direct to the Java?** *Recommendation: go direct.*
   Phase A is built against docx4j's code and Word's readings of T1 to T8, which are available
   now; the in-context goldens are added to the parity tests when core-ts commits them. Waiting
   ties a correction to shipped behaviour (the name rule) to another repository's schedule. The
   alternative, if the parity bar must hold from the first commit, is to ask core-ts to land the
   harness and golden extension first: it is Java-side work, independent of its TypeScript port.
2. **Ship the name rule ahead of the rest?** *Recommendation: yes*, as Phase A's first commit and
   in the next release whether or not the rest is ready. It is small, it is measured, and it
   makes `Table.style` right for a document whose default table style has been renamed.
3. **A resolver with no package** (`PropertyResolver(styles=...)`): does the exception apply?
   *Recommendation: do what docx4j's `CompatibilityOptions.of(null)` does*, to be read in Phase
   A and recorded; if the Java cannot be built without a package, treat no settings as no
   `compatibilityMode`, which is mode 12 with the setting absent, so the exception applies.
