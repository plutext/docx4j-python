# CR-026: docx4j CR-067 phases 2 and 3's fields in the text model --- formulas, conditions, SET and nested fields; the counts, the file fields, the list numbers and SYMBOL, under `fields: update`

**Status:** Proposed 2026-10-29, at the user's word ("Yes, have them join the CR-067 phase 2 formula fields as a follow-on CR here"). Three phases (section 6): A, phase 2's fields (`=`, COMPARE, IF, SET, ASK and FILLIN, nested fields, a line break in an instruction); B, phase 3's (NUMWORDS, NUMCHARS, TEMPLATE, FILENAME, FILESIZE, LISTNUM, AUTONUM, AUTONUMLGL, AUTONUMOUT, SYMBOL); C, the records. The questions are in section 8, each with a recommendation. Nothing built.

*The number:* CR-025 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057.

**Depends on:**
- **docx4j CR-067 phase 2** (`f9992c7eb`, its section 13.2, D15 to D18): `=`, COMPARE, IF, SET and nested fields, measured against Word 365 (Jason's F9 run of the expression probe, 115 cases and a table, all matched; regional format English (Australia)). The merge round (`2a0e0605e`, 13.4, D19 to D21) is the mergers' and has no counterpart here (section 4.6).
- **docx4j CR-067 phase 3** (`5bbeb3868`, 14.2, D22 to D25; `317420c7e`, 14.4, D27 to D29; `e1ffa9a74`, 14.5, D26 built 2026-10-29): the counts, the file fields, INDEX, the address fields, and in round 2 the list numbers and SYMBOL, measured from Word's PDFs, since Word writes no result for those five into the file.
- **docx4j's follow-on, `b76c341e7`** (CR-067 section 14.6, 2026-10-29): `FieldValues.COMPUTED` gains ASK and FILLIN, as the WD has said since 0.38 (docx4j-12's answer to this session the same day: "the WD's reading is the intended one, and the code is the omission"); cached mode unchanged.
- **docx4j's memo, `bc7e36144`** (CR-067 section 14.7, 2026-10-29): Jason's five Word saves of `one-file-many-readings.docx` (`f6d87da4a`), scored against every reader. Word evaluates a field's code without its hidden text (section 2.1). docx4j's `FieldsPreprocessor` and `Extractor` leave a hidden run's instruction text out of the instruction the same day; the fix was in docx4j's working tree, not yet committed, when this was written. The portfolio passed the rule on, and it replaces the IF probe's C12 idea that a computed result takes the formatting of its branch.
- **The WD** (`../docx4j-portfolio/docs/text_extraction_specification_wd.md`), 0.37, 0.38, 0.39 and 0.40: REQ-505 (a SET's stored result never shown) and REQ-506's `fields: update` list. 0.34 to 0.40 await Jason's ruling; this CR follows the draft as it stands and moves with it.
- **This repository:** CR-025 Phase B (`23fa5ce`: `model/fields/`'s `FieldInstruction`, `FieldContext`, the registry, `FieldValues` over them) and CR-011 section 6 (`fields: update`).
- **Counterparts:** docx4j as above; docx4j-core-ts, whose CR-012 extraction takes `fields: update` from the WD and has not ported CR-067's phases (none of its CRs names them, 2026-10-29); docx4j-mcp's `extract_text` through the WD.

## 1. Why

A document's formula, IF, list number or word count keeps its cached result in this text model under `fields: update` today: `FieldValues` computes CR-025's 20 names and no other. For most fields that is a stale value at worst. For five it is wrong in every Word-saved file: Word paints LISTNUM, AUTONUM, AUTONUMLGL, AUTONUMOUT and SYMBOL live and never writes their results (docx4j's round 2: every case of `listnum-probe-r2-word365.docx` still holds the probe's `OLD` after Word's update and save), so their cached text is whatever produced the file first, and only `fields: update` can give what Word shows.

Measured at `cd697f6`, `extract(..., fields="update")` over docx4j's probes leaves `OLD` in 121 of the expression probe's 136 lines, 68 of the phase 3 probe's 100, all 15 of round 2's cases and all 7 of the word-count probe's. The same reading drops a `w:cr` inside a field's instruction (found 2026-10-28: `TextField.instr` reads `' QUOTE "ab" '` for `a⏎b`), which is where Word writes the line breaks of its own ADDRESSBLOCK format (docx4j's D29).

docx4j's text model computes all of these, with Word's measurements behind each (its D11: the inspector's set grows with the evaluator, the WD's list updated each phase). This CR brings the port to docx4j's set at `b76c341e7`, held to the same probes.

## 2. What Word does

### 2.1 Phase 2 (docx4j CR-067 section 13.2)

- **The formula.** The default number format is two decimals with trailing zeros dropped (`=1/3` 0.33; a whole number plain; no overflow at 10^12). The error texts, verbatim: `!Zero Divide`, `!Unexpected End of Formula`, `!Missing Operator`, `!Syntax Error, <token>`, `!Undefined Bookmark, <NAME>` (upper-cased). INT truncates toward zero, ROUND rounds half away from zero, MOD takes the dividend's sign, COUNT counts numbers only. `1,000` reads as a thousand and is echoed with its separator (`1,001`). The functions: ABS, INT, ROUND, MOD, SIGN, SUM, AVERAGE, MIN, MAX, COUNT, PRODUCT, AND, OR, NOT, TRUE, FALSE, IF, DEFINED. A bookmark's text is a number in a formula. Cells (`A1`, `ABOVE`, `LEFT`, ranges) are read from the field's table.
- **COMPARE and IF** (D16). A bare operand naming a bookmark is the bookmark's text; quoted, it is the string; a quoted number is a string (`"10" > "9"` is 0); strings compare case-sensitively; `*` and `?` in the right operand are wildcards, for `=` and `<>` only. An operator written without spaces (`IF 5=5 yes no`) is `Error! Unknown op code for conditional.`
- **SET** (D15) stores its value as the field's result under a bookmark of its name and Word displays nothing: the PDF is blank where the file holds the value. A REF, a formula or an IF after it reads the value, in document order.
- **Nested fields** (D17) are evaluated inside out, each result one operand of the outer instruction (quoted where empty or holding a space or a quote). Word writes a nested field's result as `w:instrText`, since it lies inside the outer field's instruction.
- **The number symbols** (D18) are the document's `w:decimalSymbol` and `w:listSeparator`, else `.` and `,`; `\#` formats the unrounded value, and `\*` applies after it.
- **Hidden text in a code** (`bc7e36144`, the memo's sentence 9, `IF 1 = 1 "<hidden>confidential" none`): Word evaluates the code without its hidden text, so the true branch is empty and Word shows nothing, both on open and after F9. Round 5 had shown that a hidden stored result does reach Word's plain text, so F9 wrote nothing rather than a hidden "confidential". The hidden run stays in the file; only the instruction evaluated leaves it out. docx4j's update had printed "confidential" visibly, and docx4j fixed that the same day.
- `=2^0.5` shows 1.4 in Word; docx4j gives 1.41, unexplained (its section 9).

### 2.2 ASK and FILLIN (docx4j's follow-on, the WD 0.38)

Outside a merge an extractor supplies no answers, so each gives its `\d` default (empty without one). An ASK shows nothing, as Word shows nothing for it, and defines its bookmark with that text, so a REF or an IF after it reads it; a FILLIN shows the default.

### 2.3 Phase 3 (docx4j CR-067 sections 14.2, 14.4 and 14.5)

- **NUMWORDS and NUMCHARS** (D25): the main story's whitespace-separated words and its characters without spaces; no header, no footnote; a field's own result counts. Word's 36 and 212 on the word-count probe.
- **TEMPLATE**: the extended properties' `Template` (`Normal.dotm`); `\p`, the template's path, keeps the cached result.
- **FILENAME and FILESIZE** (D25, D26): the name and size the caller supplies on the package (`OpcPackage.setSourceName`, `setSourceSize`, request attributes never saved); without them the cached result. FILENAME shows the last path segment, `\p` all of it; FILESIZE `\k` and `\m` round by a thousand (16820 bytes showed 17).
- **LISTNUM** (D27): the level is `\l`, else the paragraph's heading level, else the level after the nearest heading before it, else 1. The three named lists (NumberDefault, which a bare LISTNUM is; LegalDefault; OutlineDefault) share one set of nine counters. Counting a level resets the deeper ones; a level started under a parent that has not counted gives the parent 1; `\s` sets the count. NumberDefault paints `1) a) i) (1) (a) (i) 1. a. i.`, OutlineDefault `I. A. 1. a) (1) (a) (i) (a) (i)`, LegalDefault the counts from level 1 joined by periods plus a period (`2.1.14.`).
- **AUTONUM, AUTONUMLGL and AUTONUMOUT** (D27): their own counters, one per heading level and one for body text; a field counts at its paragraph's level and resets the deeper ones; a heading with no field changes nothing. AUTONUM is the count and `.` (or `\s`); AUTONUMLGL the levels counted so far joined by periods and `.` (`\e` drops it); AUTONUMOUT `I. A. 1.` by level.
- **SYMBOL** (D28): the code (decimal or `0x`) is a character, Windows-1252 for a text font; a symbol-encoded font's code is its glyph, painted from U+F000 plus the code (Word's `w:sym`); `\u` a code point; `\f` the font, `\s` the size; a code Word cannot read paints `###` in bold.
- **A line break in an instruction** (D29) is a line break: a `w:cr` or `w:br` among the instruction runs, or `&#xA;` in a `w:fldSimple`'s `w:instr`. A hand-written `\n` is the letter n.
- A field in a header, footer, footnote or endnote keeps its cached result for the list numbers (D27). INDEX keeps its cached paragraphs, needing layout for its page numbers (D22).

## 3. What exists here (measured 2026-10-29 at `cd697f6`)

- **The registry** (`model/fields/evaluators.py`) holds docx4j's phase 1 set and DATE, TIME and PRINTDATE; nothing of phases 2 or 3. `FieldInstruction` already names a formula `=` (`' =2+3 '` gives `=` and `['2+3']`), as docx4j's does.
- **`FieldValues.COMPUTED`** (`model/text/field_values.py`) is CR-025's 20 names.
- **`FieldContext`** (`model/fields/context.py`) has the package, zone, clock, language, user, section, document, `computes`, `sequences` and `bookmarks`; no node, no cells, no file name or size, no list numbers, no answers.
- **The switches** (`model/fields/formatting.py`) are docx4j's `FormattingSwitchHelper` at `198d015a3`, with the decimal symbol read from the settings; phase 2's `applyGeneralSwitch` (the `\*` switch alone, after a formula's own number format) is not here.
- **The walk** (`model/text/extractor.py`) is docx4j's `Extractor` before phase 2 in six places: a nested field's result is not appended to the outer's instruction; a `w:instrText` after a field's separate is appended to its instruction where docx4j's walk now reads it as the stored result; text inside a field's code is not given to the nested field it belongs to; a `w:cr` or `w:br` in the code is not a line break of the instruction (`_text` writes it only under `field_codes: true`; the instruction gathers `w:instrText` alone, `_run_item`); a SET's stored result is shown; and a hidden run's `w:instrText` is gathered into the instruction (the memo's sentence 9 gives `' IF 1 = 1 "confidential" "none" '`).
- **The package** (`openpackaging/packages/opc_package.py`) has `date_time_zone` (CR-025 Phase A), and no source name or size.
- **The suite.** No fixture under `samples/` or `tests/` holds one of these fields outside docx4j's three CR-046 FORMTEXT files, whose control paragraph C is `{ = 2+3 }` cached as `WRONG`. Under `fields: update` it becomes 5, which is what Word's own save after F9 holds (`cr046-formtext-unprotected-after-f9.docx`) and what docx4j's extraction now gives. `test_formtext_as_word_updates_it` checks only paragraphs A and B, so it still passes; its comment ("C, a formula, is not computed") changes.

## 4. Design

### 4.1 The evaluators (Phases A and B)

Each evaluator is ported from docx4j at `317420c7e` into `model/fields/`, one module per docx4j class, under its own names, registered as docx4j registers it:

- **Phase A.** `formula.py` (docx4j's `Formula`: the parser and evaluator, the error texts, the functions, the bookmark references, the cells through the context), `conditional.py` (`Conditional`: COMPARE and IF), and the registrations of `=`, COMPARE, IF, SET, ASK and FILLIN. `FieldContext` gains `as_operand` (D17's quoting), `cells` (None by default; docx4j's `FieldContext.Cells`, the updater's view, kept as the seam), `answers` (empty) and the number symbols (D18). `formatting.py` gains `apply_general_switch`.
- **Phase B.** `list_numbers.py` (`ListNumbers`: the two counter sets over the main document, computed once per pass, looked up by the field's node), `symbol_field.py` (`SymbolField`), the story counts in `sources.py` (`FieldSources.storyCounts`: its own walk over the body, a break or tab a space, a paragraph's end a newline), and the registrations of NUMWORDS, NUMCHARS, TEMPLATE, FILENAME, FILESIZE, LISTNUM, AUTONUM, AUTONUMLGL, AUTONUMOUT and SYMBOL. `FieldContext` gains `node`, `file_name` and `file_size`, the last two defaulting to the package's (4.3).

### 4.2 The text model (Phases A and B)

`FieldValues.COMPUTED` takes docx4j's set: phase A adds `=`, COMPARE, IF, SET, ASK and FILLIN; phase B the ten of section 2.3. A SET's and an ASK's result is empty, their values in the context's bookmarks. `compute` takes the field's node (its `w:fldSimple`, or the `w:fldChar` beginning it) for the list numbers.

The walk takes docx4j's six changes (section 3), all in Phase A; only the node waits for Phase B:

1. a nested field's result, computed else stored, is appended to the outer's instruction as an operand (`as_operand`);
2. a `w:instrText` after a field's separate is its stored result (Word's form for a nested field's result);
3. text inside an outer field's code belongs to the innermost separated field there, as its stored result;
4. a `w:cr` or `w:br` in a field's code is a newline of its instruction (D29), in every mode;
5. a SET's stored result is never shown, in every mode (REQ-505);
6. a hidden run's `w:instrText` before the field's separate is left out of the instruction, though the run itself is kept. "Hidden" is the walk's own reading of the run, the `hidden` that `_run_item` already takes (its own `w:vanish`, else its style's), as docx4j's `Extractor` uses `isHiddenRun` (section 2.1).

Items 1 to 4 and 6 change `TextField.instr` in the block model for the fields they touch, in every mode, since the instruction is read the same way whatever the `fields` option (Q4). Item 5 changes the default text of a document holding a SET (Q4). With `field_codes: true` the codes are written as they stand, as today. A hidden code run is already absent there, as it is from docx4j's `field_codes` column for the memo, so item 6 is consistent with that output.

A formula naming a table cell keeps its cached result: the walk supplies no cells, as docx4j's supplies none and the WD says (Q6).

### 4.3 The package's source name and size (Phase B)

`pkg.source_name` (a `str` or None) and `pkg.source_size` (an `int` or None), after docx4j's D26: request attributes beside `date_time_zone`, never saved, not set by `load()`, reset by a caller reusing a package. `FieldContext` reads them as the defaults of `file_name` and `file_size`, the per-call values overriding; FILENAME and FILESIZE keep their cached result without them (Q5).

### 4.4 The numbers and the text

The formula's numbers are Python floats, as docx4j's are Java doubles, and the default format and `\#` go through `model/java_text.py`'s `DecimalFormat`, which `formatting.py` already uses. The case where docx4j and Word differ (`=2^0.5`, 1.41 against 1.4) is docx4j's here too, held with docx4j's tolerance and listed as a departure from Word.

### 4.5 Fixtures

Copied unedited from `../docx4j/docx4j-core-tests/src/test/resources/fields/`, each directory with a `SHA1SUMS`:

- `cr067-phase2/` into `tests/fixtures/expr-probe-cr067/`: the probe, Word's save, the README with the predictions, `run-notes.txt`.
- From `cr067-phase3/` into `tests/fixtures/phase3-probe-cr067/`: `phase3-probe-r1.docx`, its Word save and PDF; `listnum-probe-r2.docx`, its save and PDF; `numwords-probe-r1.docx`, its save and PDF; `address-probe-r1-attached.docx` (Word's own ADDRESSBLOCK, its line breaks as `w:cr`, for D29's reading); both READMEs and `run-notes.txt`. The PDFs are the record of what Word painted for the five live fields, which docx4j's tests encode as literals. The merge files are left out (section 4.6).
- `cr067-fields-fun/` into `tests/fixtures/fields-fun-cr067/`, from `bc7e36144` on: the memo `one-file-many-readings.docx`, Word's five saves (`-word365-asis.docx`, `.txt` and `.pdf`; `-word365-f9.txt` and `.pdf`), the README, `scored.tsv`, `docx4j-readers.tsv` and `local-readers.txt`.

### 4.6 Not here

- **A write-back.** There is no FieldUpdater here. docx4j's D15 (bookmarking a SET's result), D17's updater half, D22 (INDEX regenerated), D27's writes and SYMBOL's run would come with a FieldUpdater port, its own CR if wanted (Q1).
- **INDEX** keeps its cached paragraphs, needing layout (D22); nothing changes for it.
- **The merge fields.** MERGEFIELD, MERGEREC, MERGESEQ, NEXT, NEXTIF, SKIPIF, ADDRESSBLOCK and GREETINGLINE keep their cached result under `fields: update`, as docx4j's text model keeps them, a merge and not an extraction supplying the record. This repository has no merger (CR-002 names mail merge out of scope), so the merge evaluators are not registered (Q1).
- **DATE, TIME and PRINTDATE** stay out of `fields: update`, CR-025 Q6.

## 5. Tests and oracles

- **The expression probe** (`ExpressionProbeWordTest`'s text-model half): Word's save is the oracle, every case and the table, through `extract(..., fields="update")` over the probe; the table's cell formulas keep `OLD` (no cells, 4.2); F07 (`=2^0.5`) and F20 read with docx4j's tolerance.
- **`FormulaTest` and `ConditionalTest`**, ported whole through `evaluate()`: the operators, switches, error texts, functions, bookmarks, cells through a context's `cells`, a document with `,` and `;` symbols; IF, COMPARE, bare bookmark names, SET, the operand quoting.
- **ASK and FILLIN**: `FieldValuesPromptsTest` (`b76c341e7`): an ASK with `\d "Jane"` and a REF to its bookmark reading Jane, a FILLIN with `\d "blue"` and one without, and under `fields: cached` the stored results; with section 2.2's IF after an ASK and a locked field cached.
- **Phase 3** (`Phase3ProbeWordTest`'s text-model halves, at `e1ffa9a74`): `fileFieldsAsWord` through `pkg.source_name` and `source_size`; `wordCountsAsWord`; `listNumbersAsWord` and `listNumbersRound2AsWord` (L01 to L19, A01 to A15, R01 to R15 against the PDFs' numbers); `listNumbersInTextAndHtml`'s text half; `sourceNameAndSizeOnThePackage`; `symbolsAsWord` (S01 to S14, the characters).
- **D29**: Word's own ADDRESSBLOCK read from `address-probe-r1-attached.docx`, its instruction carrying the newlines, its result cached; a `w:br` in a code; `&#xA;` in a `w:fldSimple`.
- **docx4j's own extraction** of the four probes, run from docx4j's build at `b76c341e7` (`Texts.plainText` under `fields: update`, the JVM in Sydney's zone), matched byte for byte, as CR-025 Phase B matched phase 1's.
- **The memo** (`FieldsFunMemoTest`'s text-model half, once docx4j commits it): under `fields: cached` the ten sentences are the stored column. Under `fields: update` they are docx4j's update column as fixed: 30, 750, "yes", 4 and an empty sentence 9, with 2020 and "two or later" kept, since DATE and PAGE are not computed. The facade's text does not change. Sentence 9's `TextField.instr` is `' IF 1 = 1 "" "none" '` in every mode.
- **The default mode**: a SET's stored result absent with `fields` left as `cached`; an outer field's `TextField.instr` holding its nested results.
- **The gates**: the suite; `scripts/roundtrip.py` 50 of 50; `compare_acceptance.py` 0 differ (nothing saved changes); `codegen/generate.sh --check` for the new files, each in `codegen/clean.py`'s `KEEP`; the text-model goldens and the agreement TSV unchanged.

No Word check is expected: extraction output is not a file, and nothing this CR adds is written to one.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | Phase 2's fields: `formula.py`, `conditional.py`, SET, ASK and FILLIN; `FieldContext`'s operand quoting, cells seam, answers and number symbols; `apply_general_switch`; the walk's six changes (4.2), D29's line break among them; the expression probe's and the memo's fixtures and replays, `FormulaTest`, `ConditionalTest` | 1 day |
| B | Phase 3's fields: `list_numbers.py`, `symbol_field.py`, the story counts, TEMPLATE, FILENAME and FILESIZE; `pkg.source_name` and `source_size`; the node in the walk; the phase 3 fixtures and replays | 1 day |
| C | The records: CHANGELOG (Added: the fields under `fields: update`, `pkg.source_name` and `source_size`; Changed: a SET's stored result no longer shown, a nested field's result in its outer's instruction; Fixed: a line break in an instruction), README and `docs/pypi-readme.md` where they list `fields: update`'s set, CLAUDE.md (the status line, the new files in its list), the index's row; the portfolio's `python/CR-026.*` entries through the docx4j-portfolio session | ¼ day |

## 7. Risks

- **Default output moves** for a document holding a SET (its stored value goes from the text) or a nested field (the outer's `instr` in the block model). Section 3's scan finds neither in the suite's fixtures, so no golden or agreement row should move; a caller reading the block model's `instr` sees it change.
- **The formula's grammar** is the largest piece (docx4j's `Formula` is 558 lines). `FormulaTest` and the probe's 115 cases are the gate, and docx4j's own extraction the byte-for-byte check.
- **The list numbers depend on heading levels.** docx4j's rule (`ListNumbers.headingLevel`) is the paragraph's effective outline level, else the number in its style's name (`heading 2`) or id (`Heading2`). The port takes that reader as written rather than the text model's own `outline_level` (which reads `w:outlineLvl` or a `HeadingN` style, `extractor.py` around line 1046), so the two engines agree on a heading-free style with an outline level; the probes are the gate.
- **The WD moves.** 0.37 to 0.40 await Jason's ruling; if a ruling changes REQ-505 or REQ-506, this port follows in its next phase.

## 8. Questions for the user

| Q | Question | Recommendation |
|---|---|---|
| Q1 | Scope: the text model's `fields: update` and the registry only, with no FieldUpdater (no write-back of a result, no INDEX regeneration) and no merge evaluators (4.6)? | Yes. Extraction is what this repository has a consumer for; a FieldUpdater port, with D15, D22 and D27's writes, is a CR of its own if a caller wants fields updated in a saved file. |
| Q2 | The list numbers and SYMBOL (LISTNUM, AUTONUM, AUTONUMLGL, AUTONUMOUT, SYMBOL) in the set, beside the counts and the file fields? | Yes. The WD (0.40) lists them and docx4j's text model computes them. Word never writes their results, so in a Word-saved file their cached text is stale and only `fields: update` gives what Word shows (2.3). This session called them display-only on 2026-10-28; that is the reason to compute them, not to leave them out. |
| Q3 | ASK and FILLIN under `fields: update`: their `\d` default, ASK defining its bookmark and showing nothing (2.2)? | Yes, settled by docx4j the same day (`b76c341e7`); recorded so that the decision is the user's too. |
| Q4 | The default mode's two changes: a SET's stored result never shown (REQ-505), and a field's `instr` in the block model holding its nested fields' results but not a hidden run's code text (4.2 items 1 and 6)? | Yes, as docx4j's walk does both in every mode: Word displays nothing for a SET, and the instruction is the one a `fields: update` evaluates. CHANGELOG's Changed names both. |
| Q5 | FILENAME and FILESIZE: `pkg.source_name` and `pkg.source_size`, set by the caller only; or set by `load(path)` from the path's name and size as well (4.3)? | The caller only, as docx4j's D26. A load from bytes or a stream has neither, a server's file is a temporary name, and the ports then read alike. A caller who wants Word's FILENAME sets it in one line. |
| Q6 | A formula naming a table cell under `fields: update`: keep the cached result, as docx4j and the WD do, or read the cells from the walk's table? | Keep the cached result. The walk would need the cell values Word reads (formula results before it, in document order), docx4j's text model does not supply them, and the outputs would then differ. `FieldContext.cells` stays the seam for a caller of `evaluate()`. |
| Q7 | Phasing: A (phase 2's fields and the walk's changes), B (phase 3's), C (the records), each built and verified before the next, as CR-025's were? | Yes. |

## 9. Effort

2¼ to 2½ days: A 1, B 1, C ¼.
