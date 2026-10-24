# CR-016: A cell ends with a paragraph; `Font`, placeholders and deleted marks read as Office JS reads them; wildcards across marks --- the port of docx4j-core-ts `222cc20` and `3ba1e68`

**Status:** Proposed 2026-10-24, at the user's request ("draft CR-016 now"). It ports to this package the behaviour changes docx4j-core-ts made on its `main` (unpushed) in `222cc20` and `3ba1e68`, which CR-003 sections 23.3 and 23.4 make owed here: member-for-member and behaviour parity with core-ts. CR-003 section 24 took the one member they added (`ContentControl.showing_placeholder_text`) and left the behaviour for this decision. Six items: **Decided 2026-10-24** (the user: "CR-016 Q1 the release will wait for it; the other questions answered as recommended"): 0.2.0 waits for CR-016 (Q1, against its recommendation); every other question as recommended, checks 58's and 59's files Apache-2.0 here (Q2). Phase B builds on CR-018 Phase B's readers (section 4.3's note). **Phase A implemented 2026-10-24** (section 10): `table_rules.py` (docx4j's `ends_with_paragraph` and `end_with_paragraph`, exported); `Paragraph.delete`, `Body.clear`, `ContentControl.delete` and `Body.insert_ooxml("Replace")` keep a cell's end paragraph, emptied, its mark never marked under tracking; a table after it splits it at its end; docx4j's repair elsewhere and in what goes in (D3); nothing on load (D4); a tracked `TableCell.value` a tracked deletion (Q11); the facade's `_Cell.add_table` python-docx's (Q13). Check 52's cases 86 to 95 Word's in the run, the save and Accept All but the new table's paragraph (Q4) and case 94's value; over the corpus's 2,435 cells, no write leaves one without a paragraph end, and `Range.delete()` moved in none. One pre-existing expectation moved. Its Word check is Phase D's (files built, section 10.5). Phases B to D not started.
1. Word's table rule (core-ts CR-002 section 55): a cell ends with a paragraph.
2. `Font` reading every run in scope (section 59).
3. A control showing its placeholder read as no text (core-ts CR-012 section 18).
4. A deleted paragraph mark read as nothing (CR-012 section 19).
5. The paragraph tracked-change listing kept and pinned (CR-002 section 58.1, K1).
6. Word's wildcards (CR-002 section 32): K4, and K7 at `3ba1e68`.

Three things core-ts did in the same commits are done here already (section 3.8). Four phases are proposed (section 5): A, the cell rule; B, the readers; C, the wildcards; D, the Word check and the records. They come to 4½ to 5½ days. Fourteen questions, decided 2026-10-24 (section 8). **Phase A is implemented** (section 10); B to D are not started.

*The number:* an earlier, undrafted idea used "CR-016" for a port of docx4j CR-057 (`insert_markdown` reading local image files). That idea now takes CR-017.

**Depends on:**
- **docx4j-core-ts**, cited by commit only:
  - `222cc20`: CR-002 sections 55 (55.6, 55.7), 58.1, 59 and 59.1; CR-012 sections 18, 19, 18.1 and 19.1; CR-002 section 32's K4 to K6; Word checks 58 and 59, written and run.
  - `3ba1e68`: CR-002 section 32's K7.
  - `4837d04`: CR-014 section 12.3. It is here already (section 3.8) and is cited only so that the record is whole.
  - `ebb218e` (2026-09-28): CR-002 section 32 item 1's grapheme filter, which Q8 proposes to take here.
- **docx4j** `18a3a31f4` (`VERSION_17_3_4`), for item 1's semantics: `org.docx4j.model.table.TableRules` (`endsWithParagraph`, `endWithParagraph`), `TableRulesTest`, and `docs/developer/word-table-rules.md`.
- **This repository:**
  - CR-013 (complete): the span edits, the story layout (`story.py`), `split.py`, `office_matches` and `^p`.
  - CR-010: Accept All keeps a cell's last paragraph; the listing; the views.
  - CR-011: Phase F's rule that nothing is edited inside a kept branch; `hidden_runs_of`, the hidden-text rule shared with the text model; the text model's `office_offset`.
  - CR-012 Phase A: kept block-level `mc:AlternateContent`, read on Word's branch.
  - CR-003 section 24 (`15dde99`): `showing_placeholder_text`.
  - CR-014 and CR-009 Phase B (both proposed) are touched, not needed (section 4.8).
- **No fork change.** No schema, generator or runtime change is foreseen; `codegen/generate.sh --check` is a gate only to prove it.
- **The sibling checkout** `../docx4j-core-ts`, for the subset (unchanged: 247 members, section 1) and, at the time of the port, for the oracles' sources (section 6.1).

**Counterpart:**
- docx4j-core-ts's `docs/change-requests/CR-002-content-api.md` at `3ba1e68` (sections 32, 55 to 55.7, 58, 58.1, 59, 59.1) and `CR-012-text-model.md` at `3ba1e68` (sections 18, 19, 18.1 and 19.1). Neither file changed between `222cc20` and `3ba1e68`, except section 32's K7.
- Its code in `src/model/content/`:
  - `tableRules.mts`, new;
  - `Font.mts`, `Paragraph.mts`, `Body.mts`, `ContentControl.mts`, `Range.mts`, `Table.mts`;
  - `search.mts` (`wildcardMatchesAcross`, `refuseWildcardMarkReplacement`);
  - `spanEdits.mts` (`keepCellMarkLast`), `story.mts` (`shownInDefaultView`) and `tree.mts` (`inShowingPlaceholder`).
- Its `src/model/customxml/xhtml.mts` and `src/html/`, which have no counterpart here.
- Its tests:
  - `test/table-rules.test.mjs`, 7 tests;
  - `test/readers.test.mjs`: checks 58 and 59, and the tests of sections 18 and 19;
  - `test/ranges.test.mjs`: check 52's cases 86 to 95, and the fonts check 53 now compares;
  - `test/content.test.mjs`: the font tests.
- Its `test/README.md`, checks 58 and 59.

## 1. Why

**The decision.** CR-003 section 23.3 keeps member-for-member Office JS parity with core-ts, and parity is behaviour too. Section 23.4's review, after CR-013, recommended keeping it. core-ts `222cc20` added one member, which CR-003 section 24 ported the same day (`15dde99`), and changed behaviour in six places. Section 24 left those for the user's decision: "parity behaviour, not members".

**The subset passes.** `tests/office_js_subset.json` has 247 members, derived from core-ts's committed `HEAD`. `test/office-js-subset.ts` didn't change in `3ba1e68` or `4837d04`, and `tests/content/test_office_js_subset.py` passes (10 tests, 2026-10-24). So nothing here holds 0.2.0's release gate (Q1).

**What it costs today.** Section 2 measured each item. Three findings are defects that matter outside parity:
- **Edits leave cells Word can't read.** On Word 365, docx4j's probes 02, 05, 09 and 12 measured that a cell whose last block is a table opens as unreadable content. A cell with no block breaks the schema (`CT_Tc`). Today, several edits here leave one or the other: `cell.body.insert_table(..., location="End")`, `Paragraph.delete()` of a cell's last paragraph, `cell.body.clear()`, `insert_markdown` of a table into a cell, `ContentControl.delete(keep_content=False)` of a control holding a cell's paragraph, and `insert_xml` or `insert_ooxml` of a table at a cell's end: eleven edits in all (section 2.1). 0.1.0 does the same.
- **A tracked `TableCell.value = "..."` loses the cell's old content untracked.** It is not a deletion, and Reject All can't give it back.
- **A value written into a control showing its placeholder stays a placeholder.** `insert_text` leaves `w:showingPlcHdr` on, so `update_from_content_controls()` skips a bound control as holding no value, and the value never reaches its custom XML.

## 2. What exists here (measured 2026-10-24 at `e7e0b20`)

Each probe below is read only: a package built with `create_package()` and `insert_xml`, or the files Word read, loaded from the share. The probes and a scratch port of checks 58 and 59 (`c58.py`, compared through core-ts's own views) are in this session's scratchpad. Nothing under `docx4j_py/` or `tests/` was changed.

### 2.1 A cell's last paragraph (item 1)

core-ts CR-002 section 55.2's table, run here. A cell holds either one list paragraph (p: List Paragraph, numbered, underlined), or a nested table and then that paragraph (tp). The middle column is the cell's last block after the edit, read through block-level controls.

| edit | the cell after it, here | tracked (Accept All / Reject All) |
|---|---|---|
| `cell.body.insert_table(1, 1, location="End")` (p) | ends in `w:tbl` | `w:tbl` / `w:p` |
| `cell.body.insert_xml(tbl, location="End")`; `Body.insert_ooxml(tbl, location="End")`; `Paragraph.insert_xml` and `insert_ooxml` of a table `"After"` the cell's last paragraph; `cell.body.insert_markdown` of a table, and `Paragraph.insert_markdown(..., location="After")` | ends in `w:tbl` | --- |
| `Paragraph.delete()` of the cell's only paragraph (p) | **no block** | the paragraph stays, its **mark marked deleted**, which Word never writes; Accept All keeps a paragraph (CR-010) |
| `Paragraph.delete()` of the last paragraph, after a nested table (tp) | ends in `w:tbl` | the same: mark marked deleted |
| `cell.body.clear()` (p, and tp) | **no block** | every mark marked deleted |
| `cell.body.insert_ooxml(tbl, location="Replace")` (p) | ends in `w:tbl` | --- |
| `ContentControl.delete(keep_content=False)` of `insert_content_control()`'s control, or of one around the cell's only paragraph (TC3) | **no block** | the same; a control's delete is not tracked (CR-003 section 4) |
| `ContentControl.delete(keep_content=False)` of a control around the last paragraph, after a nested table | ends in `w:tbl` | the same |
| `Range.insert_ooxml(tbl, location="Replace")` on the cell's paragraph, or on the cell body's `"Content"` range | keeps the rule: the table, then the paragraph emptied, its `w:pPr` kept (Word's form in case 94) | the same |
| `Range.delete()` of the cell paragraph's whole range; `TableCell.value = ""`; `insert_paragraph(..., "End")`; deleting the nested table; `insert_table(..., "Start")` | keeps the rule | keeps it |
| `TableCell.value = "X"` (tp), tracked | one paragraph `X` | **the nested table and the list paragraph go untracked**; the one change listed is `Added X` |

- **Over the corpus**, nothing is broken already. The corpus is `out/cr012-b-corpus/`, CR-013's (`COMMITS.txt`: this repository at `f1bb1d9`, core-ts at `df5b4af`, docx4j at `901a922a3`). Of its 418 distinct documents, none has a cell whose last block, read through `w:sdt` and `w:customXml`, isn't a paragraph (0 of 2,495 main-part cells). So D4 (Q3) has nothing to repair there.
- **No XHTML pass or html module exists here.** CR-014 phase 2 and CR-011 Phase E name them for later. The markdown importer is the nearest thing. Its own tables can't nest (GFM), but it puts blocks in through `Body.insert_element`, which is where a table at a cell's end has to be caught.

### 2.2 Check 52's cases 86 to 95 through the port

`tests/content/test_ranges.py` stops check 52's write tests at case 85 (`RANGE_CASES`), because section 55 wasn't implemented. Run to 95 here, through the same views:

| test | regions of 86 to 95 differing from Word here | core-ts at `222cc20` |
|---|---|---|
| the run (`run_check_52(..., "run")` against Word's run) | 86, 87, 88, 89, 91, 92, 94, 95. 90 and 93 agree only because the run's view hides a tracked case's text | 86 and 87 (the new table's own paragraph, Q4), 94 (the value returned) |
| the markup of each region, against Word's save after the run | all ten: no kept paragraph (88, 89, 92, 94, 95); a cell's last mark marked deleted (90, 91, 93); no split before the table (86, 87) | 86 and 87, the new table's paragraph alone |
| Accept All of this package's save, against Word's Accept All | 86, 87, 88, 89, 91, 92, 94, 95 | 86 and 87 |

### 2.3 `Font` (item 2)

`Font` reads the first run's effective properties (`font.py`, `_direct_r_pr`). `MixedFont` reads every run only across paragraphs (CR-013 Q4).

| read | here | core-ts at `222cc20` (Word where measured) |
|---|---|---|
| `paragraph.font` over a bold, red, yellow-highlighted run and an italic, underlined 14 pt Arial run | `True`, `False`, `"None"`, `"Aptos"`, `11.0`, `"#FF0000"`, `"#FFFF00"` (the first run's) | `None`, `None`, `"Mixed"`, `""`, `None`, `""`, `""` (check 53 cases 8 and 31; 39 of check 52's paragraph reads) |
| the same paragraph's `"Content"` range | the first run's | the same as the paragraph |
| a run with no colour anywhere; `w:color w:val="auto"` | `""`; `"auto"` | `"#000000"` both (check 53); the direct view still `""` and `"auto"` |
| across two paragraphs, one of each | `""` (`MixedFont`: the two differ) | `"#000000"` |
| an empty Heading 1 paragraph whose mark has `w:b w:val="0"` | `bold True`: the style's | `False`: the mark's (check 52 case 60) |
| Heading 1's theme colour | `"#365F91"` | the same; Word reads `#0F4761` through the theme. That is CR-009 Phase B here, and CR-008 phase 2 there |

- **Check 53** (its Simple Markup JSON, colours compared too): 12 cases' fonts differ here. Eight are an automatic colour (2, 7, 8, 9, 19, 20, 21, 24). Three are the theme colour (3, 4, 5). Two are runs that differ: case 8's `italic`, and case 31's colour and highlight. In core-ts after section 59, only 3, 4 and 5 differ.
- **Check 52**, on Word's save after Accept All: 40 of 1,269 paragraph font reads differ (29 regions), as CR-013 section 6.1 counted them. In core-ts, none does.
- **Found: a font read splits runs.** `Range.get_font`'s `holders()` calls `split_runs_around` on every call, a read's too. Reading `paragraph.search("TWO")[0].font.bold` on a paragraph of one run leaves three runs. Loaded from a save and read the same way, the next save writes 3 runs where the source had 1. core-ts splits only on a write (`holders(write)`).

### 2.4 A control showing its placeholder (item 3)

core-ts's `readers.test.mjs` test of section 18, run here: a block control, and a run-level control between `A ` and ` B`, each showing `Click here.` in the `PlaceholderText` style.

| read or write | here | core-ts (Word: check 53 cases 34 and 35, check 58 cases 11, 12 and 20 to 22) |
|---|---|---|
| the block control's `text`, the run-level one's, and their paragraph's | `Click here.`, `Click here.`, `A Click here. B` | `""`, `""`, `A  B` |
| `body.text` | `Click here.\rA Click here. B` | `\rA  B` |
| `body.search("Click")` | 2 hits | none (K3, Q5). Word finds hits that read `""` |
| `placeholder_text` | `Click here.` for both | the same, read from the runs (Q6). Word reads `null` |
| `insert_text("FILLED", location="End")` | `Click here.FILLED`; still showing its placeholder; the `PlaceholderText` style kept | `FILLED`; not showing; the style dropped. Word's save has no `w:showingPlcHdr` and a plain run |
| `location="Start"` | `FILLEDClick here.`, still showing | `FILLED` |
| `location="Replace"` | `FILLED`, **still `w:showingPlcHdr`** | `FILLED`, cleared |
| the text model (`plain_text`) | `Click here.\nA Click here. B` | the same. Neither engine changes the text model, which follows docx4j's semantics |

- **The bound control.** A bound `w:text` control showing its placeholder took `insert_text("VALUE", location="Replace")`, and then `update_from_content_controls()` ran. The node stays `""`, and the control is skipped as "a control showing its placeholder holds no value" (`bindings.py`, `update_from_control`).
- **Over the corpus**, 14 of the 418 documents hold 21 controls showing their placeholder.

### 2.5 A deleted paragraph mark (item 4)

core-ts's test of section 19, run here. `ONE`'s mark is deleted, then `TWO`, then a one-cell table whose paragraph `CELL` has a deleted mark, then `LAST`.

| read | here | core-ts (Word: check 58 cases 6 to 10) |
|---|---|---|
| `ONE`'s `mark_text` | `\r` | `""` |
| `ONE`'s `"Whole"` range | `ONE\r` | `ONE` |
| from it through `TWO`'s `"Whole"` range | `ONE\rTWO\r` | `ONETWO\r` |
| `body.text` | `ONE\rTWO\rCELL\r\nLAST` | `ONETWO\rCELL\r\nLAST` |
| `len(body.paragraphs)` | 4 | 4 (they stay apart) |
| the cell's paragraph's `mark_text` | `\t` | `\t` (the cell's end, whatever its mark) |
| `get_text(view="original")` | `ONE\rTWO\rCELL\r\nLAST` | the same |
| an inserted mark: `get_text(view="original")` | `ONE\rTWO` | `ONETWO`: Reject All's view writes no inserted mark |
| `search("^p")` | matches the deleted mark | doesn't |
| the text model (`plain_text`), both views | `ONETWO` in the accepted view, and an inserted mark joined in the original: each view already reads what its resolution leaves | unchanged |

Over the corpus, 78 of the 418 documents hold 291 deleted paragraph marks, and 81 hold an inserted one.

### 2.6 The paragraph listing (item 5)

Through core-ts's own `listing` view (`readers.test.mjs`: the types of the region's changes, and of each paragraph's, its whole range's and its content range's), check 58's cases 1, 5 and 14 to 19 are Word's here. Cases 2, 3 and 4 differ exactly as core-ts pins them: this package lists each deleted paragraph's change in that paragraph, where Word lists nothing. So K1 is this package's state already, and only the pin is owed. Accepting or rejecting a paragraph's part (check 58's Edit run, cases 14 and 16) resolves that part only, as in Word.

### 2.7 Wildcards (item 6)

`search_pattern` (`text_model.py`) makes `?` into `.`. A Python string is code points, so `?` matches an emoji whole already, and **K4 has nothing to port here**.

| query (`match_wildcards=True`) | here | Word (check 58 case 13, check 59) | core-ts |
|---|---|---|---|
| `?` over `x😀y a😀b كَتَبَ z` | 16 hits: each code point, the Arabic letters and their fathas apart; no mark | `x`, 😀, `y`, ` `, `a`, 😀, `b`, ` `, ` `, `z`, `\r`: nothing in the Arabic word | Word's |
| `[!a-z ]` | the two emoji and the six Arabic code points | `\r` alone | Word's |
| `[!a-z ]@` | 😀, 😀, the Arabic word | `\r` alone | 😀, 😀, the Arabic word, `\r` (K6) |
| `Q?Z`, `Q^13Z` and `Q*Z` over the paragraphs `Q` and `Z` | nothing | `Q\rZ` each | Word's |
| `^13` | the literal `^13` | each mark | Word's |
| a paragraph's `?` | `Q` | `Q`, `\r` | Word's |
| a cell holding `Q` and `Z`: its body's `?` | `Q`, `Z` | `Q`, `\r`, `Z` | Word's |
| `body.replace_text("Q?Z", "X", match_wildcards=True)` | 0 replaced | --- | refused (`NotSupportedError`), nothing changed |

**Found: no grapheme filter.** `find_all` returns a hit that cuts a grapheme cluster: `?` matches an Arabic letter without the fatha after it. core-ts has dropped such hits since `ebb218e` (2026-09-28: `matchesOf`, every search, `findAll` and `replaceText`), which is why its `?` and `[!a-z ]` agree with Word over the Arabic word. `is_grapheme_boundary` (CR-003 section 3.12) is here already; it serves the edits' run splits, not the search.

### 2.8 Checks 58 and 59 through a scratch port

The scratch port runs on the files Word read, from the share, and is compared through core-ts's views (check 58's `listing`, its `placeholder` view without `placeholderText`, and its edit view):

| check | Word's here | differing | core-ts after `3ba1e68` |
|---|---|---|---|
| 58, read (Simple Markup) | 10 of 22 (1, 5, 9, 10, 14 to 19) | 2, 3, 4 (K1, as core-ts); 6, 7, 8 (section 19); 11, 12, 20, 21, 22 (section 18); 13 (wildcards) | differs in 2, 3, 4 (K1), 12 (K3) and 13 (K6) |
| 58, edit | 7 of 9 | 21 and 22: `"End"` writes after the placeholder | none |
| 59 | 0 of 5 | every case: no wildcard matches a mark | none |

### 2.9 Found on the way, each taken into a phase or a question

1. **A tracked `TableCell.value` removes the cell's content untracked** (section 2.1). The setter calls `body.content.clear()` and then `insert_paragraph`, as 0.1.0's did. core-ts's setter clears through `clearContent(false)`, which is tracked. Q11.
2. **A `Range.font` read splits runs** (section 2.3). Q12.
3. **Writing into a placeholder leaves it a placeholder**, so a bound value never reaches its node (section 2.4). Section 18's write fixes it, in Phase B.
4. **No grapheme filter on search** (section 2.7). Q8.
5. **The facade's `_Cell.add_table`** (`docx/table.py`) calls `body.insert_table(rows, cols)` at `"End"` and then `add_paragraph()`, python-docx's own rule. Under section 55, `insert_table` at a cell's end would leave a paragraph after the table itself, and the facade would add a second. Q13.
6. **Two members core-ts's section 55.2 lists are absent here:** `ContentControl.clear()` and `Paragraph.insert_table`. Neither is in the subset, and neither is needed.
7. **core-ts's `ContentControl.insertText` syncs a bound node at once** (`syncBoundNode`). Here `update_from_content_controls()` writes the node. That is CR-014's ground, not this CR's (section 4.8 item 3).

## 3. What core-ts built

### 3.1 CR-002 section 55: a cell ends with a paragraph, as Word reads it (`222cc20`)

**The rule** is docx4j's `TableRules` (`18a3a31f4`), measured on Word 365 in twelve probes:
- a cell holds a block;
- its last block, as Word reads it, is a `w:p`, looked for past range markup, through block-level controls and custom XML, and through the one branch of an `mc:AlternateContent` Word reads;
- Word's repair appends an empty `w:p` after any wrapper, and Word's own edits never break the rule.

**D1 to D4** were decided 2026-10-22: at the edit; follow Word where measured; the cells of inserted content repaired too; a file that already breaks the rule left as it is until an edit touches the cell.

**Check 52's cases 86 to 95** (Word 365, 2026-10-22) answered D2. Each is a one-row table whose first cell holds a list item: alone (TC1), after a nested table (TC2), or in a block control (TC3).
- `Paragraph.delete()` (88, 89) and `body.clear()` (92) of the paragraph the cell ends with empty it, keeping its `w:pPr` whole.
- Tracked, they delete its text and leave its mark unmarked (90, 91, 93).
- `ContentControl.delete(false)` (95) leaves the paragraph, emptied, outside any control.
- `body.insertOoxml(table, 'Replace')` (94) puts the table before the cell's paragraph, which is kept, emptied.
- `body.insertTable(…, 'End')` (86, 87) splits the cell's paragraph at its end, with the table between the halves:
  - the first half keeps the text and gains its last run's `w:u` as mark properties, and is the inserted mark when tracked;
  - the paragraph after the table is the cell's own mark, unmarked;
  - Word also gives the new table's own paragraph the cell paragraph's properties, which core-ts doesn't follow (Q4).

**As built** (section 55.7):
- **`tableRules.mts`:** `endsWithParagraph`, `endWithParagraph`, and, for the edits, `paragraphEndOf`, `wrappingCellOf`, `endsCell` and `endCellsWithParagraphs`.
- **`Paragraph.delete()`** keeps the cell's end paragraph, emptied, where removing it would leave the cell without one. Tracked, it never marks that mark.
- **`Body.clear()`** keeps it too, through `clearContent(keepCellEnd)`. `insertText(…, 'Replace')` and `TableCell.value` call `clearContent(false)`, since they put in a paragraph of their own next.
- **`Body.insertOoxml(…, 'Replace')`** puts content ending in a table before the kept paragraph. **`ContentControl.delete()`** keeps the paragraph in the control's place.
- **A table after the cell's end paragraph**, in the cell's own content, is `keepCellMarkLast` (`spanEdits.mts`), on CR-002 section 60.8's split at the end.
- **Everything else goes through `Body.insertElement`**, which ends the cell with docx4j's repair, plain under tracking, and repairs the cells of what goes in before it goes in (D3). The XHTML pass and the html module do the same.
- **Held to Word:** check 52's cases 86 to 95 in the run, the save and Accept All.
- **Tested:** `test/table-rules.test.mjs`, 7 tests, which fail on `906d5e3`'s build.
- **Swept:** 602 documents, each top-level cell given four edits, 4,097 cells an edit. Every cell that changed is one the old build left without a paragraph, or with its last mark marked deleted.

### 3.2 CR-002 section 59: `Font` reads every run in scope (`222cc20`)

- **Where runs differ**, `Font` reads as Office JS does: `null`, `''` for the name, the colour and the highlight, and `'Mixed'` for the underline. `MixedFont` became `Font` with its own readings and its own write. The getters keep Office JS's `boolean` and `number` types (Jason: "Office JS's types kept").
- **A paragraph's read takes the runs of `Paragraph.text`** (`getRange('Content').runs`): Word's default view, so no deleted run and no equation's math runs. With none, it takes a placeholder's runs, which Word's font reads though its text doesn't (check 52 case 13 after Accept All).
- **A paragraph with no text** reads its mark (`markHolderOf`: `w:pPr/w:rPr` without the mark's revisions). So does a range holding the mark and no run. A collapsed range reads the paragraph's style, as before.
- **An automatic colour** reads `#000000` in the effective view; the direct view says what is written.
- **Theme colours** are not in it (core-ts's CR-008 phase 2).
- **A write** still goes to every run.
- **Held to Word:** check 53's fonts (all but cases 3 to 5's theme colour), and check 52's paragraph fonts in all three run tests.

### 3.3 CR-012 section 18: a control showing its placeholder (`222cc20`)

- **Reads:** `inShowingPlaceholder` (`tree.mts`) finds a control with `w:showingPlcHdr` up a run's parents, run- or block-level, and `hiddenRunsOf` takes its runs as hidden. So `Paragraph.text`, `Body.text`, `ContentControl.text`, a range's text and its offsets read no placeholder, and `search` finds nothing there. **K3** is a departure from Word, which finds hits that read `''`.
- **`ContentControl.insertText`** into such a control, block- or run-level, clears `w:showingPlcHdr` and puts the text in the placeholder's place whatever the location, as Word does with `'End'` (check 58 cases 20 to 22). It drops the `PlaceholderText` character style from the runs.
- **`placeholderText`** reads the control's runs, a departure: Word, with no glossary part to read, reads `null`.
- **`Paragraph.font`** of a paragraph whose only content is a placeholder reads the placeholder's runs.
- **The text model is unchanged.**
- **Measured by check 58**, run 2026-10-23, Word 365 16.0.20326.20232: a run-level control reads empty too (case 11). Writing replaces the placeholder at `'End'` as at `'Replace'`.

### 3.4 CR-012 section 19: a deleted mark reads as nothing in Word's default view (`222cc20`)

- **The layout:** `storyLayout` writes no `\r` after a paragraph whose mark is tracked as deleted (`shownInDefaultView`), and `Paragraph.markText` is `''` for one.
- **A cell's last paragraph** keeps the cell's `\t`.
- **The paragraphs, their offsets and the paragraph collection** stay as they are.
- **`^p`** doesn't match such a mark.
- **`Body.getText({ view: 'original' })`** takes the original view's rule: a deleted mark written, an inserted one not.
- **Measured by check 58**, cases 6 to 10: Word opened a deleted mark at a cell's end without a repair, and read it as the cell's end.

### 3.5 CR-002 section 58.1: the paragraph listing kept (`222cc20`)

- **K1, kept.** Word's listing differed between two files of the same shape: check 53's case 25 listed the next paragraph's part of a change through a mark, and check 58's cases 14 to 17 didn't. It also lists nothing in three of check 58's shapes: cases 2 and 3, two deleted paragraphs in a row, and case 4, a paragraph whose text alone is deleted.
- **The pin.** core-ts's listing is Word's in cases 1, 5 and 14 to 19. `readers.test.mjs` pins it against Word's JSON, with the differences at cases 2 to 4.
- **K2, a range's listing cut to the range: not now.** Case 5's content range is recorded as a difference.

### 3.6 CR-002 section 32: Word's wildcards (`ebb218e`, `222cc20`, `3ba1e68`)

- **`ebb218e`** (2026-09-28): `matchesOf` drops a hit whose start or end is not a grapheme boundary (`Intl.Segmenter`, only for a text with a code unit at or above U+0300).
- **K4** (`222cc20`): `?` takes a surrogate pair whole. **K6:** the classes over the emoji and the Arabic letters are not followed; `[!a-z ]@` keeps finding them.
- **K7** (`3ba1e68`, check 59): a wildcard matches a paragraph mark exactly where `^p` does.
  - It matches a mark between body paragraphs, one inside a cell, and one inside a block control, the control's last included.
  - It never matches a cell's or a row's end, nor anything across a table's edge.
  - `^13` is the mark's wildcard code. `?` and `*` cross a mark (`wildcardMatchesAcross`).
- **The scopes.**
  - `Range.search` over a range across paragraphs, or holding its mark, joins each run of paragraphs that are in one cell, or in none, by their own marks: `markText` `\r`, never a cell's end or a deleted mark, and the last only where the range holds it.
  - `Paragraph.search` searches its `'Whole'` range, its own mark included.
  - `Body.search`, and so a block control's, searches its `'Content'` range, without the last mark (check 59 case 5).
- **`replaceText` with wildcards** refuses, before anything changes, where a hit holds a mark (`refuseWildcardMarkReplacement`), as a plain `^p` query is refused.

### 3.7 Checks 58 and 59 (`222cc20`; core-ts `test/README.md`)

- **Check 58** was run 2026-10-23 in Word 365 (16.0.20326.20232, PC, Script Lab): a Read run in Simple Markup and one in All Markup, then an Edit run in All Markup. It has 22 cases:
  - 1 to 5: a paragraph's listing;
  - 6 to 8: a range across a deleted mark;
  - 9 and 10: a deleted mark at a cell's end;
  - 11 and 12: a placeholder, read and searched;
  - 13: wildcards over an emoji and an Arabic word;
  - 14 to 19: a paragraph's part of a change, accepted and rejected;
  - 20 to 22: a placeholder written into.

  Word opened the file without a repair. Its save after the three runs is committed (`test/check58/`, with a manifest).
- **Check 59** was run the same day. It is read only, and Word saved nothing. It has five regions: two body paragraphs; a cell of one paragraph; a cell of two; a two-by-two table; a block control of two. Each is searched with wildcards in four scopes.
- **In both checks the file Word read is git-ignored in core-ts**, built by `test/fixtures/check5x/build.mjs`. The share's copy is the faithful source, as for CR-013's checks (section 6.1).

### 3.8 Already here

- **`ContentControl.showing_placeholder_text`**: CR-003 section 24 (`15dde99`). Check 53's oracle still lists it, Word having returned it undefined.
- **docx4j CR-043 section 17's Accept All rules**, which core-ts carried in `222cc20` and, as `e55a83b8e` revised them, in `4837d04`: CR-010 sections 13, 13.3a and 14 (`3864d77`). The three engines agree on them (CR-010 section 14).
- **core-ts CR-014 section 14**, the `deleteUntracked` fix, which was this package's finding: CR-013 section 14.5.

## 4. Design here

Follow core-ts, rule for rule, except where Python convention or a rule decided here says otherwise. Section 4.9 lists the departures, each a question.

### 4.1 The items, one by one

| item | core-ts's change | Word check behind it | here today (section 2) | what changes here | readers that move, and what must hold or move |
|---|---|---|---|---|---|
| 1. A cell ends with a paragraph (CR-002 section 55) | the edits keep the cell's end paragraph, emptied, its `w:pPr` kept; tracked, never marked; a table at the cell's end splits that paragraph at its end; elsewhere docx4j's repair; D3 | check 52 cases 86 to 95 (run, save, Accept All); docx4j's twelve probes | eleven edits leave a cell ending in a table or holding no block; tracked, the last mark is marked deleted; a tracked `TableCell.value` loses content | `table_rules.py`; `Paragraph.delete`, `Body.clear`, `insert_text("Replace")`, `TableCell.value`, `ContentControl.delete`, `Body.insert_ooxml("Replace")`, `Body.insert_element` (and so every insert, `insert_markdown` among them); D3; the facade's `_Cell.add_table` (Q13) | **no reader moves**. Saved output changes, so there is a **Word check**. The cell sweep (section 6.2); check 52 to case 95 in its three write tests |
| 2. `Font` reads every run (section 59) | `null`, `''` or `'Mixed'` where the runs differ; the paragraph's text's runs; the mark where there is no text; automatic `#000000` | checks 52 and 53 | the first run's; automatic `""` or `"auto"`; an empty paragraph reads its style; **a read splits runs** | `Font` takes `MixedFont`'s rule; the mark; `#000000`; no split on a read (Q12); getters `bool \| None` (Q10) | `Paragraph.font` and `Range.font`. The readers' audit reads no font today, so it gains one (section 6.2). Check 53 to 3 differences, check 52's paragraph fonts compared |
| 3. A placeholder is no text (CR-012 section 18) | its runs hidden to the Office JS readers; no search hit (K3); `insert_text` replaces it at any location and clears it; `placeholder_text` from the runs | check 53 cases 34 and 35; check 58 cases 11, 12 and 20 to 22 | read as text everywhere; `insert_text` keeps `w:showingPlcHdr` | `hidden_runs_of` takes a placeholder's runs; `ContentControl.insert_text`; `placeholder_text` | `Paragraph.text`, `Body.text`, `get_text`, `ContentControl.text`, `find`, `search`, `element_at`, `TableCell.value`, and the text model's `office_offset` for a placeholder's runs (to `None`); the text model's text doesn't move. Saved output of `insert_text` changes, so there is a **Word check** |
| 4. A deleted mark is nothing (CR-012 section 19) | no `\r` after a deleted mark in the layout; `mark_text` `""`; a cell's `\t` kept; the original view drops an inserted mark | check 58 cases 6 to 10 | `\r` after every mark, in both views | `story_layout`'s mark rule by view; `mark_text_of` | `Body.text`, `get_text` in both views, a range's text across paragraphs, a `"Whole"` range, `^p`, and a block control's, a cell's and a comment's text where they hold such a mark. The paragraph views and offsets don't move |
| 5. The paragraph listing kept (CR-002 section 58.1, K1) | nothing changed; pinned against Word's JSON | check 58 cases 1 to 5 and 14 to 19 (against check 53 case 25) | the same as core-ts's | the pin, in check 58's oracle | **nothing moves** |
| 6. Wildcards (CR-002 section 32) | K4 `?` a code point; K7 a mark matched where `^p` matches; `replace_text` refuses a hit holding a mark; `ebb218e`'s grapheme filter, since 2026-09-28 | check 58 case 13; check 59 | K4 already holds; no mark matched; no grapheme filter | `search_pattern`'s `^13` and `(?s)`; a wildcard search across marks; the scopes; the refusal (Q9); the filter (Q8) | `search` and `find` with `match_wildcards`; plain `search` and `find` where a hit cuts a cluster (the filter). The audit's `body.search` and `body.find` may move only there |

### 4.2 The cell rule (Phase A)

**The module.** `docx4j_py/model/content/table_rules.py`, core-ts's `tableRules.mts` and docx4j's `TableRules`:
- **docx4j's two functions**, by docx4j's names in Python's spelling, exported from `docx4j_py.model.content` as core-ts exports them:
  - `ends_with_paragraph(contents)`: the last block past range markup, through `w:sdt`, `w:customXml` and the branch the readers take of a kept `mc:AlternateContent` (CR-011 Phase F, CR-012 Phase A);
  - `end_with_paragraph(tc)`: Word's repair, an empty `w:p` after any wrapper. It returns the paragraph, so that a caller can give it a `w14:paraId`.
- **The helpers the edits use**, internal: `paragraph_end_of`, `wrapping_cell_of`, `ends_cell` and `end_cells_with_paragraphs`.
- `model/` is in `codegen/clean.py`'s `KEEP` whole, so nothing is added there.

**The edits:**
- **`Paragraph.delete()`.** Of the paragraph a cell ends with, where removing it would leave the cell without one, it empties the paragraph and keeps its `w:pPr`. Tracked, it deletes the paragraph's text and leaves its mark unmarked. Its `ChangeReport` records the text gone and no block shifted, since the paragraph stays where it was.
- **`Body.clear()`.** It keeps the paragraph the cell ended with, emptied, outside any control it was in. Tracked, it goes through `Paragraph.delete()`. A private `_clear_content(keep_cell_end=...)` serves `insert_text(..., location="Replace")` and `TableCell.value`, which put in their own paragraph next. Under tracking that clear is tracked, which is Q11's fix.
- **`ContentControl.delete()`** of a block control holding the paragraph a cell ends with leaves that paragraph, emptied, in the control's place. With `keep_content=True` nothing changes, since the paragraph stays anyway.
- **`Body.insert_ooxml(..., location="Replace")`** puts content ending in a table before the cell's kept paragraph (case 94). Content ending in a paragraph replaces it, as now. `Range.insert_ooxml` already gives case 94's form (section 2.1).
- **A table put after the paragraph a cell ends with**, in the cell's own content (`insert_table`, `insert_xml`, `insert_ooxml`, `insert_markdown` or `insert_element` at `"End"`, or `"After"` that paragraph), splits that paragraph at its end (`keep_cell_mark_last`). This uses `split.py`'s split at a paragraph's end, core-ts's section 60.8, which CR-013 took. Then:
  - the first half's new mark takes the formatting of the text before it, and is the inserted mark when tracked;
  - the old `w:pPr` goes after the table, in an empty paragraph, unmarked;
  - the paragraph keeps its text, so that a view of it still holds it.
- **Everything else**, through `Body.insert_element`, the hub every insert goes through: docx4j's repair, plain under tracking.
- **D3.** `insert_element` repairs the cells of what goes in, at any depth, before it goes in, so that under tracking the repair is part of the insertion. `wml(...)` fragments, `insert_xml`, `insert_ooxml` and the markdown importer reach it through `insert_element`.
- **D4.** A file that breaks the rule stays as it was read until an edit touches that cell. An untouched part is copied byte for byte.

**Kept branches** (CR-011 Phase F, CR-012 Phase A). The rule reads a kept block-level `mc:AlternateContent` on the branch Word reads, as docx4j and core-ts do.
- An untracked `clear()`, or a control's delete, that takes such an element whole and leaves the cell without its end paragraph keeps that paragraph, Word's branch's, emptied and moved out, as core-ts does.
- An edit *inside* a kept branch stays refused (`address.read_only`), so a `Paragraph.delete()` of a branch's paragraph is refused here, where core-ts edits it. That is Phase F's existing departure, not a new one.
- Tracked, a kept element is refused as now (`kept_content_under`).

**The facade** (Q13). `docx4j_py.docx`'s `_Cell.add_table` keeps python-docx 1.2.0's result: the cell's paragraph untouched, the table, then one plain empty paragraph. It gets there by putting the paragraph in first and the table before it, so that section 55's split never runs. `tests/content/test_python_docx_parity.py` holds it.

**What is not ported:**
- core-ts's XHTML pass and html module, which have no counterpart here.
- When CR-014 phase 2 (rich-text and XHTML bindings) or CR-011 Phase E (the HTML reader) builds cells, each ends them with a paragraph through `end_with_paragraph`. This CR's Phase D notes that in both CRs.

### 4.3 `Font` (Phase B)

- **Every run in scope**, read as `MixedFont` reads them now: `None` where they differ, `""` for the name, the colour and the highlight, and `"Mixed"` for the underline.
  - `MixedFont` keeps only its own readings across paragraphs and its write (`span_edits.format_span`), as core-ts's does.
  - `format_span`'s rule (a toggle written only where a run doesn't read it already) reads each run alone, so it doesn't move.
- **A paragraph's read takes `get_range("Content").runs`**: `Paragraph.text`'s runs in Word's default view, with no deleted run and no math runs. With none, it takes a placeholder's runs.
  - `Range.runs` here is core-ts's code already (`range.py`): a zero-width hidden segment at a range's edge is not read, and one inside it is.
- **No run.** A paragraph with no text reads its mark's run properties (`w:pPr/w:rPr` without `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` and `w:rPrChange`). So does a range holding the mark and no run. A collapsed range reads the paragraph's style, as now.
- **An automatic colour** (`w:val="auto"`, or none anywhere) reads `"#000000"` in the effective view. `get_font(direct=True)` still says what is written. Theme colours are CR-009 Phase B's; check 53's cases 3 to 5 stay a recorded difference until it lands.
- **No split on a read** (Q12). `holders` takes `write`, as core-ts's does, and only a write splits the runs at the span's edges.
- **The types** (Q10). `Font`'s getters are typed `bool | None` and `float | None`, as `MixedFont`'s are since CR-013 Q4.
- **The python-docx facade's `Font`** keeps python-docx's tri-state semantics (CR-003 section 20.3), and is untouched.
- **The `ChangeReport`** of a font write records the direct value before the write. Where the runs differ, that is now `None`, not the first run's value.

*Note (2026-10-24): CR-018 Phase B lands first, on today's `Font`, changing the readers and the merge beneath it (enumerations outside their enumeration unset, the colour reader, a character style's highlight not drawn). This phase's `Font` reading every run then builds on those readers.*

### 4.4 Placeholders (Phase B)

- **`hidden_runs_of`** (`text_model.py`) takes a run inside a control with `w:showingPlcHdr` on, block- or run-level, found up the run's parents, as hidden. That is core-ts's `inShowingPlaceholder` inside `hiddenRunsOf`.
  - Every Office JS reader goes through it already: `Paragraph.text`, the story layout, `Range`, `search`, `office_matches` and the comment and change readers.
  - So do the text model's office spans (`extractor.py`, `_office_spans`). A placeholder's run then has no `office_offset` (`None`) and keeps its text. The text model reads a placeholder as text, flagged `placeholder`, as now: docx4j's semantics, as core-ts keeps them.
  - The audit's `office_offset` invariant (CR-013 Q11) holds with those runs left out.
- **K3** (Q5): a placeholder holds no text to search, as hidden text holds none.
- **`ContentControl.insert_text`** into a block- or run-level control showing its placeholder:
  - it clears `w:showingPlcHdr`, and writes in the placeholder's place at any location, as `"Replace"`;
  - it removes a `PlaceholderText` `w:rStyle` from the runs it leaves;
  - CR-013 section 11's break rules hold, since the write takes the `"Replace"` path; so will CR-014's one-paragraph rule, when it lands.
- **`placeholder_text`** (Q6) reads the control's runs while it shows the placeholder: a run-level control's content, a block control's paragraphs `\r` apart. So the setter's text reads back.

### 4.5 A deleted mark (Phase B)

- **The layout.** `story_layout` (`story.py`) takes a mark rule by view:
  - the default, the accepted view, writes no `\r` after a paragraph whose mark is deleted;
  - the original view writes none after an inserted mark;
  - a cell's last paragraph keeps the cell's `\t` whatever its mark.
- **`mark_text_of`** gives `""` for a deleted mark, and `\t` before that for a cell's last.
- **What follows:**
  - `Body.text` and `get_text` read through the layout;
  - a `"Whole"` range's text ends with `mark_text`;
  - `^p` reads `mark_text == "\r"` (`range.py`), so it stops matching a deleted mark;
  - the paragraph views, their offsets and the collection don't change.
- **Not touched:** `TrackedChange.text`, which keeps docx4j's characters, as the revisions goldens hold it (CR-013 section 3.8).

### 4.6 The paragraph listing (K1)

Nothing in the code changes. Check 58's port pins the listing against Word's JSON, as core-ts's test does:
- the region's changes, and each paragraph's, its whole range's and its content range's, by type;
- the differences named at cases 2, 3 and 4, and case 2's paragraphs pinned as core-ts pins them.

The Edit run's cases 14 and 16 hold that accepting or rejecting a paragraph's part resolves that part only. K2 is recorded as a difference (Q7).

### 4.7 Wildcards (Phase C)

- **K4:** nothing to do (section 2.7). A test holds `x?y` over `x😀y` to Word's.
- **`search_pattern`'s wildcard branch.**
  - `^13` is `\r`.
  - The pattern is compiled with `re.DOTALL`, so that `?` and `*` cross a mark, as core-ts's `[\s\S]` does.
  - `.` already crosses `\r` in Python, but not `\n`, and a row's `\r\n` is never in a wildcard's text, since no hit crosses a cell. So `DOTALL` is for clarity, and a test pins it.
- **A wildcard search across marks**, `Range._search_wildcards` (core-ts's `searchWildcards` and `wildcardMatchesAcross`):
  - it groups the range's paragraphs into runs in one cell, or in none, by `parent_table_cell`;
  - it joins each run's text by the marks a hit may hold: a paragraph's own (`mark_text == "\r"`), never a cell's end or a deleted mark, and the last only where the range holds it;
  - it maps each hit back to a `Range` with an end paragraph, holding its mark where it ends with one;
  - hidden text is left out, and a picture is U+FFFC, as in `office_matches`.
- **The scopes.**
  - `Paragraph.search` with wildcards searches its `"Whole"` range.
  - `Body.search`, and through it `ContentControl.search` and a cell's body, searches its `"Content"` range.
  - `Range.search` searches across marks where it spans paragraphs or holds its mark.
- **`replace_text` with wildcards** (Q9). It finds every hit first, with the same options over the same scope; if any hit holds a mark, it refuses before anything changes, with `SpanError` `range.replaces_mark` and a hint to search and edit the hits' ranges. That is the `^p` refusal's code (`refuse_mark_replacement`), and section 4.8 item 5 covers the rule.
- **The grapheme filter** (Q8, core-ts `ebb218e`).
  - `find_all` and the wildcard path drop a hit whose start or end is not a grapheme boundary, using `text_model.is_grapheme_boundary`.
  - It is checked only for a text with a code point at or above U+0300, core-ts's guard, so Latin text pays nothing.
  - It applies to every search and every replacement, as core-ts's does.

### 4.8 Interactions with this package's own rules

1. **Section 55 against CR-013's span edits.** CR-013 Phase C keeps a held last mark that has no paragraph after it in its list (a cell's last, a control's, the story's, one before a table), and refuses a range into or out of a table (`range.partial_table`). So a span already leaves a cell ending with its paragraph, and `Range.delete()` of a cell's whole range keeps it (section 2.1).
   - Phase A brings the other verbs to the same rule, and makes `span_edits`' "can this mark go" test (`span_edits.py`, around line 629) read `table_rules.ends_cell`, so that one rule serves both.
   - core-ts's `deleteUntracked` fix in the same commit is CR-013 section 14.5's, here already.
   - The cell sweep adds `Range.delete()` of each cell's whole range, which must not move.
2. **Section 59 against CR-013's `MixedFont`.** `MixedFont`'s reading rule moves into `Font`, and CR-013 Q4's types become `Font`'s too (Q10).
   - A range within one paragraph stops reading its first run. That is the change check 53 asks for, cases 8 and 31.
   - `range.font` across paragraphs reads as now, but for the automatic colour, which becomes `#000000` where it read `""`.
3. **Section 18 against CR-014's binding work (proposed).**
   - `update_from_control` skips a control showing its placeholder. With section 18's write, a value written into a bound placeholder clears the flag and reaches its node at the next `update_from_content_controls()` (section 2.4's probe is the test).
   - `apply_bindings` writes docx4j's placeholder for an empty value (`PLACEHOLDER_TEXT`, `w:showingPlcHdr`), which then reads as no text, as Word reads it.
   - The bound write-back reads `docx4j_text`, which doesn't use the hidden rule, so it doesn't move.
   - core-ts's immediate node sync on `insertText` is a difference CR-014 should record. It is not this CR's.
4. **Section 19 against CR-010's and CR-011's views.**
   - The text model reads each view as its resolution leaves it already (section 2.5): the accepted view joins at a deleted mark, the original at an inserted one. Section 19 brings `get_text(view=...)` and `Body.text` into line with it.
   - CR-010's listing and `TrackedChange.text` don't move. CR-013 Q11's `office_offset` is per paragraph and doesn't move with marks.
5. **Wildcards against `replace_text`'s check-every-match-first rule** (CR-013 section 4.9 item 3). The refusal is that rule: every hit is found before anything changes, and the call refuses if any holds a mark. Phase F's check of branch text per hit stays as it is. core-ts's refusal also comes before anything changes, so this item adds no departure.

### 4.9 Departures from Word, and from core-ts (each new one a question)

**Departures core-ts made from Word, and whether to take them:**
1. **The new table's own paragraph** (check 52 cases 86 and 87). Word gives it the cell paragraph's properties; core-ts doesn't follow (Q4).
2. **K3**, no search hit in a placeholder (Q5).
3. **`placeholder_text` read from the runs** (Q6).
4. **K1**, the paragraph listing as it is, and **K2**, a range's listing not cut to the range (Q7).
5. **K6**, a class over the emoji and the Arabic letters not followed (Q8).
6. **`replace_text` with wildcards refusing a hit that holds a mark** (Q9). Word would replace it, joining paragraphs; what that leaves is unmeasured.

**The standing ones, which this CR doesn't revisit:**
- `Other` after a nested table: core-ts CR-002 section 57, CR-013 Q6. It is asserted in cases 89, 91 and 94 as in core-ts.
- Word's bold filler of five spaces in an emptied control, not copied (check 52 cases 13 and 26).
- `Body.insert_ooxml` returning the views, not a `Range` (case 94's value).

**This package's departures from core-ts:**
- **The getters typed `bool | None`** (Q10).
- **Errors as `SpanError` codes** (Q9), CR-013 Q5's precedent.
- **The facade's `_Cell.add_table`**, python-docx's (Q13).
- **`[!a-z ]` over an emoji.** With Python's code points the class matches the emoji whole, where core-ts's UTF-16 half-matches are dropped by its filter and so agree with Word. Recorded beside K6 (Q8).
- **No edit inside a kept branch**: Phase F's, standing (section 4.2).

### 4.10 Errors

No new code is proposed. `replace_text` with wildcards reuses `SpanError` `range.replaces_mark` (Q9), and its hint names the way on: "search for the hits and edit their ranges". The cell rule raises nothing; it is a repair, as docx4j's is.

### 4.11 Not in this CR

- core-ts's K2 (a range's listing cut to the range), and Word's properties for a new table's own paragraph (Q4).
- Theme colours in `Font.color`: CR-009 Phase B.
- An XHTML pass and an html module: none exists here. CR-014 phase 2 and CR-011 Phase E are to end their cells with a paragraph when they come.
- `ContentControl.clear()` and `Paragraph.insert_table`, which the subset doesn't declare.
- core-ts's immediate sync of a bound node (CR-014's).
- What the text model reads of a placeholder: docx4j's question, asked there first, as core-ts says.

## 5. Phases

| phase | what | size |
|---|---|---|
| A | **The cell rule.** `table_rules.py`; the edits of section 4.2, D3 and D4; `TableCell.value` tracked (Q11); the facade's `_Cell.add_table` (Q13); check 52's three write tests run to case 95, with Word's `Other` read as section 57 has it; `tests/content/test_table_rules.py`, core-ts's 7 tests ported with docx4j's `TableRulesTest` probes; the cell sweep (section 6.2) | 1 to 1½ days |
| B | **The readers and the placeholder write.** Sections 4.3 to 4.6: `Font` (Q10, Q12), placeholders (Q5, Q6), deleted marks, K1's pin (Q7); check 58's port in its read and edit modes, with its fixtures (Q2); checks 52's and 53's tests comparing fonts, placeholders and lengths as core-ts's now do; the readers' audit before and after, with a font reader | 2 to 2½ days |
| C | **Wildcards.** Section 4.7: `^13`, the search across marks and its scopes, the refusal (Q9), the grapheme filter (Q8); check 59's port and fixture; check 58 case 13 | 1 day |
| D | **The Word check and the records.** Section 6.3's check; `tests/README.md`'s map and provenance; the CHANGELOG; the index row; a line in CR-003 section 24 (its "not taken now" taken) and notes in CR-014 and CR-011 (section 4.2) | ½ day |

**Order: A, B, C, D**, each committed and verified on its own.
- **A first.** It fixes saved output Word repairs, a defect 0.1.0 and 0.2.0 share, and nothing in it waits on the readers.
- **B before C.** Check 58's port and fixtures are B's, and C's case 13 runs on them.
- **D's Word check** covers A's and B's writes together: B's placeholder write changes saved output too.

## 6. Tests and the gate

### 6.1 The oracles

**The fixtures** (Q2). The bytes Word read and Word's own save, copied byte for byte with their SHA-1s into `tests/office_js/` and its `SHA1SUMS`, as for checks 48 to 57:

| check | file Word read (share `__tmp/`) | SHA-1 | Word's answers and save (core-ts commit) |
|---|---|---|---|
| 58 | `58/58-measures.docx` | `03647705a7461f9ec6a441213f8e96d3ca69a32e` | `test/fixtures/check58/58-measures-word365-read-simple.json` (`d299b8e6...`), `-read-all.json` (`989f8a9f...`) and `-edit.json` (`b90f8187...`) at `222cc20`; Word's save after the three runs, `test/check58/58-measures-word365.docx` (`0e5184c5ad383c67c68dd08d16ed26fe58e41a8c`, its `manifest.json`'s; the share's copy the same bytes) |
| 59 | `59/59-wildcard-marks.docx` | `30e8d448f10639f6fa27fb0bb7c00c920ac70117` | `test/fixtures/check59/59-wildcard-marks-word365.json` (`400556cd...`) at `222cc20`; read only, no save |
| 52, cases 86 to 95 | here already (CR-013 Q8) | `e014014d...` | here already: the run's and Accept All's JSON, Word's two saves |
| 53 | here already | `aea48f83...` | here already |

- **The snippets.** The share's `check58-script-lab.js` (`dca635d5...`) and `.html`, and `check59-script-lab.js` (`abd3bfa5...`) and `.html`, are the commits' byte for byte, at `222cc20` and at `3ba1e68`.
- **The JSONs.** Each JSON at `222cc20` parses equal to the share's `result-read-simple.txt`, `result-read-all.txt`, `result-edit.txt` and `result.txt`.
- **The builders.** core-ts's builders (`build.mjs`) are git-ignored output there, so the share's copies are the faithful source of the files Word read.

**The snippets in Python.** `tests/office_js/checks.py` gains `run_check_58` (its `read` and `edit` modes) and `run_check_59`, with the same JSON shape and `vis()`, following checks 52 to 57's ports:
- `parse_region_58` and `parse_region_59`;
- `QUERIES`;
- the ops `changes`, `range`, `cell`, `placeholder`, `search`, `accept`, `reject` and `fill`.

This session's scratch port (section 2.8) shows that every op maps onto the content API as it stands: no member is missing, and no op errors.

**What the tests hold:**
- **Check 52**, cases 86 to 95 joining the three write tests:
  - with section 59 the paragraph fonts are compared in all three, as core-ts's are, the tracked cases without their fonts, since Word's run was in All Markup;
  - Word's run: the differences gain 13 (Word's bold filler, now that fonts are compared), 86 and 87 (Q4) and 94 (the value returned), as core-ts's;
  - Word's save, region by region: they gain 86 and 87, on the new table's paragraph alone, that line pinned;
  - Word's Accept All after the run: 13, 26, 49 to 56, 60, 66, 86 and 87, as core-ts's.
  - Accept All on Word's own save gives Word's Accept All in all 95 regions, case 26 now included.
- **Check 53**: the differences shrink to core-ts's. The fonts are 3, 4 and 5 (the theme colour). The placeholder cases 34 and 35 are `placeholderText` and `showingPlaceholderText` alone. Case 33 (`body.text`'s length) is Word's.
- **Check 58**, on the file Word read:
  - the read mode against Word's Simple Markup run, and the edit mode against Word's Edit run, through core-ts's views;
  - the differences listed: read 2, 3 and 4 (K1), 12 (K3) and 13 (K6, and `[!a-z ]` over the emoji, Q8);
  - case 2's paragraphs pinned as core-ts pins them.
  - Beyond core-ts's test: cases 20 to 22's markup against Word's save after the run (`w:showingPlcHdr` gone, the run plain).
- **Check 59**: every scope Word's. The refusals: `body.replace_text("Q?Z", ...)` and a paragraph's `replace_text("?", ...)` refuse with nothing changed, and `replace_text("Q", "X")` replaces.
- **The rest of core-ts's tests**, each ported, its strings kept:
  - `table-rules.test.mjs` (7);
  - `readers.test.mjs`'s tests of sections 18 and 19 (section 2.4's and 2.5's probes, inverted to core-ts's answers);
  - `content.test.mjs`'s font tests: runs that differ, the name where they agree, the automatic colour, a paragraph with no text reading its mark, a collapsed range reading the style, a deleted run left out of the read.

### 6.2 The gates

**Phase A:**
- **The cell sweep.** Over `out/cr012-b-corpus/`'s 418 distinct documents, every main-part cell takes five edits, each on a fresh load, on the code before and after:
  1. delete the cell's last paragraph, untracked;
  2. the same, tracked;
  3. `clear()`;
  4. `insert_table(..., location="End")`;
  5. `Range.delete()` of the cell's whole range.

  It must find that:
  - every cell that changed is one the old code left without a paragraph, or with its last mark marked deleted;
  - no cell ends without a paragraph afterwards;
  - the fifth edit doesn't move;
  - the refusals are the same, counted by code.
- **The facade's parity test** passes.
- **The suite**, fast and full.

**Phase B:**
- **The readers' audit, before and after**, as CR-013 section 6.2 ran it: `scripts/mce_branches.py readers --holding all` over the same corpus, its `COMMITS.txt` unchanged.
  - A font reader is added to the script **before** the "before" run: each paragraph's `bold`, `italic`, `underline`, `name`, `size`, `color` and `highlight_color`, and each table cell's.
  - **Every answer that moves must be explained**, mechanically: it holds a placeholder (section 2.4's 14 documents), a deleted or inserted mark (the 78 and 81), runs that differ, an empty paragraph, or an automatic colour.
  - **Expected to move:**
    - `body.paragraphs`, `body.get_text` in both views, `body.element_at`, `body.content_controls`, `body.comments`, `body.tables`, `body.find` and `body.search`;
    - `text.extract`, in a placeholder run's `office_offset` alone;
    - the new font reader.
  - **Expected not to move:**
    - `text.plain_text`, `text.markdown`, `package.to_markdown`, `body.outline` and `package.describe`;
    - `revisions.*` and both resolutions;
    - `docx.*`, `body.inline_pictures`, `main.fonts_in_use` and `main.styles_in_use`.
- **The `office_offset` invariant** stays at 0 disagreements. The runs that lose their offset are counted, and each must be a placeholder's.
- **The rest:** the suite; CR-011's ratchet and its goldens, which are unchanged since the text model doesn't move; CR-010's Word-save comparisons; the parity goldens; `scripts/roundtrip.py`; `generate.sh --check`; and acceptance (`compare_acceptance.py` against a copy made first).
  - No existing artefact edits a cell's end or writes into a placeholder, so each should come out unchanged. Section 6.3's artefact is new.
  - `tests/test_readme.py` runs any fence that changes.

**Phase C:**
- **The audit's `body.search` and `body.find`** may move only where a hit cuts a grapheme cluster. Each move is listed with its cluster.
- **The filter's cost** is measured over the corpus, as core-ts did: it is nothing below U+0300.
- **The suite.**

**Every phase:** the full suite with no failure (5,926 passed at `600930b`), the subset test, `generate.sh --check` and the round trip, 50 of 50.

### 6.3 The Word check (Phase D)

CLAUDE.md asks for a Word check after any change to the content API's inserts. Phase A changes deletes, clears and inserts in cells, and Phase B's `insert_text` changes a control's markup. Staged in `__tmp/python-cr016/`:
1. **File 1, the cells.**
   - Each edit of section 2.1's table, on a cell holding a list item, and on one holding a nested table and then a list item, written untracked.
   - The same tracked, in a second copy, with the engine's Accept All and Reject All beside it.
   - The facade's `_Cell.add_table`.
   - **Expected:** no repair prompt. Each cell ends with a paragraph, the list item's numbering kept. In **All Markup**, a tracked delete of a cell's last paragraph shows its text deleted and no deleted mark. Word's Accept All and Reject All agree with the engine's.
2. **File 2, the placeholders.**
   - A block and a run-level control showing a placeholder, each written into at `"End"`, `"Start"` and `"Replace"`.
   - A bound one written and then `update_from_content_controls()`.
   - **Expected:** no repair prompt. Each control shows the written text in the run's own style, not as a placeholder. The bound control's custom XML holds the value.
3. **File 0 (Q14), optional.** Today's output of file 1's untracked edits, written by `e7e0b20`. **Expected:** Word repairs it, as docx4j's probes predict. That would turn section 1's defect from a prediction into a measurement of this package's own output.

## 7. Risks

1. **Readers move again in `0.x`.** 0.2.0 moves the characters. This CR moves `Body.text`, the paragraph text and the offsets again where a document holds a placeholder or a deleted or inserted mark: 14, 78 and 81 of the corpus's 418 documents. A caller storing offsets across a release sees them move. The audit makes every moved answer explicable, and the CHANGELOG entry names each rule.
2. **A placeholder becomes zero-width text.** An edit at its place meets CR-013's rule for hidden text, under which a hidden run in a span is an object that goes with it. So `paragraph.text = paragraph.text` over a paragraph holding a run-level placeholder control could take the control. Phase B tests it: either the control stays, or the write refuses, and the docstring says which.
3. **`None` in a single paragraph's `Font`.** `font.size + 2` raises where the runs differ, where it used to add to the first run's size. `0.x` allows the change, and the typing says it.
4. **The cell rule touches every insert.** A path that misses `insert_element` would still leave a cell ending in a table. Section 2.1's list is the test list, one test per path, and the sweep checks every cell. D3 runs before the insertion, so a caller's fragment is never inserted broken.
5. **The facade.** python-docx's own rule (a plain `<w:p/>` after a cell's table) and section 55's split must not stack (Q13). The parity test reads both files back with python-docx.
6. **The Python ports could misread a snippet**, as CR-013 risk 3 said. The file-and-snippet agreement tests (every `BEGIN` parses, every token and tag once) and the JSONs' case lists guard against it.
7. **The grapheme filter moves plain search.** A hit inside a decomposed accent (`e` before U+0301) goes, in `search`, `find` and `replace_text`. That is core-ts's behaviour and the point of it, but it is a move, and the audit lists it.
8. **core-ts moves on.** Its `main` is unpushed and live. This CR ports `222cc20` and `3ba1e68`, cited by commit, and records what it ported. A later core-ts change is a later amendment, as CR-013's phase 3 was.

## 8. Questions (decided 2026-10-24)

**Decided 2026-10-24** (the user: "CR-016 Q1 the release will wait for it; the other questions answered as recommended"). **Q1 is decided against its recommendation: 0.2.0 waits for CR-016.** Q2 to Q14 are decided as recommended. Q2's confirmation is given with them: checks 58's and 59's files may be used here under Apache-2.0. The questions stand below as they were asked.

1. **Q1. Does 0.2.0 wait for CR-016?** *Recommended: no.* **Decided: yes, it waits** (the user, 2026-10-24). The subset passes (247 members, section 1), CR-016 changes behaviour and no member, and 0.2.0 is otherwise ready. CR-016 goes in the next release with its own CHANGELOG entry. Phase A alone fixes cells Word repairs, a defect 0.1.0 and 0.2.0 share, so if the user wants it sooner it can go out first as a point release.
2. **Q2. Check 58's and 59's files here under Apache-2.0.** The files are:
   - the two files Word read, from the share: `58-measures.docx` and `59-wildcard-marks.docx`;
   - Word's save of check 58, from core-ts `222cc20`, where the share holds the same bytes;
   - the four JSONs, which are core-ts's (Apache-2.0).

   *Recommended: copy them byte for byte into `tests/office_js/` with their SHA-1s, as CR-013 Q8 did, once the user confirms they may be used here under Apache-2.0.* They are core-ts's builders' output and the user's own runs.
3. **Q3. core-ts's D1 to D4 for the cell rule.** At the edit; follow Word where measured; the cells of inserted content repaired (D3); a file that breaks the rule left as it is until an edit touches the cell (D4). *Recommended: adopt all four as core-ts decided them.* The corpus holds no broken cell, so D4 costs nothing there.
4. **Q4. The new table's own paragraph at a cell's end.** In check 52 cases 86 and 87, Word gives it the cell paragraph's `w:pPr` (List Paragraph, numbered `2.`). *Recommended: don't follow, as core-ts doesn't.* It is `insert_table`'s behaviour wherever a table goes, not the cell rule's, and it is recorded as a difference in the two cases.
5. **Q5. K3: a placeholder has no text to search.** Word finds hits there that read `""`. *Recommended: take it.* A zero-width hit is no range a caller can use, and it matches hidden text, which holds none.
6. **Q6. `placeholder_text` read from the runs.** Word reads `null` with no glossary part. *Recommended: take it*, so that the setter's text reads back. This package writes no glossary part either (CR-003 section 4).
7. **Q7. K1, the paragraph listing as it is, and K2, a range's listing not cut to the range.** *Recommended: keep the listing and pin it against check 58, as core-ts does; K2 not now.* Word's own answer differed between two files of the same shape. This package's listing is Word's in eight of check 58's eleven listing cases already. K2 needs views of a change's parts that neither engine has.
8. **Q8. The grapheme filter, and K6.** core-ts has dropped a hit that cuts a grapheme cluster since `ebb218e`; this package doesn't (section 2.7). *Recommended: take the filter, using `is_grapheme_boundary`, and record K6 with one more difference.* In Python a class matches an emoji whole, so `[!a-z ]` finds it where Word finds nothing. Following Word's class rule over non-BMP characters would need more checks than one paragraph gives.
9. **Q9. `replace_text` with wildcards over a hit holding a mark.** core-ts throws `NotSupportedError`. *Recommended: refuse before anything changes with `SpanError` `range.replaces_mark`*, the `^p` refusal's code, with a hint to search and edit the hits' ranges. That follows CR-013 Q5's scheme and `replace_text`'s check-every-match-first rule.
10. **Q10. `Font`'s getter types.** core-ts kept Office JS's `boolean` and `number`. *Recommended: type them `bool | None` and `float | None`*, as `MixedFont`'s are since CR-013 Q4: Python's typing says what the value can be, and the subset checks the kind, not the type.
11. **Q11. A tracked `TableCell.value` (found).** It removes the cell's content untracked. *Recommended: fix it in Phase A*, as core-ts's setter has it (`clearContent(false)`, tracked): the old content is deleted as a tracked change, then the new paragraph is inserted, so that Reject All gives the cell back. 0.1.0 has the defect.
12. **Q12. A `Range.font` read splits runs (found).** *Recommended: fix it in Phase B*, by splitting only on a write, as core-ts does. A read then leaves the tree, and the next save, as they were.
13. **Q13. The facade's `_Cell.add_table`.** *Recommended: keep python-docx 1.2.0's result exactly*, one plain paragraph after the table and the cell's paragraph untouched, by inserting the paragraph and then the table before it. The facade's promise is python-docx's behaviour (CR-003 section 20.3), and the parity test holds it.
14. **Q14. The Word check's file 0.** *Recommended: include it.* It is one more file, written by today's code, and Word's repair of it would record the defect section 1 predicts from docx4j's probes. Without it, the rule rests on docx4j's measurements alone.

## 9. Effort

| phase | days |
|---|---|
| A, the cell rule | 1 to 1½ |
| B, the readers and the placeholder write | 2 to 2½ |
| C, wildcards | 1 |
| D, the Word check and the records | ½ |
| **total** | **4½ to 5½** |

core-ts's own estimates were half a day to a day for section 55, a day for section 59, and half a day each for sections 18 and 19, each with its check. The larger part here is the oracles, as in CR-013: two more snippet ports, held to Word's JSON case for case, and three tests of check 52 run to case 95. Phase B's audit also needs a font reader first. CR-013 took one calendar day for an estimate of 12½ to 13½ (CR-003 section 23.4), so these numbers are the CRs' usual unit, not a forecast of the calendar.

## 10. Phase A implementation notes (2026-10-24)

Phase A as built, at the user's go (CR-016 decided 2026-10-24: "CR-016 Q1 the release will wait for it; the other questions answered as recommended"), on CR-018 complete (`955547c`). The oracles are docx4j-core-ts `222cc20` (its CR-002 section 55 and 55.7, `tableRules.mts`, the content files it changed and `test/table-rules.test.mjs`), read from the commit and never from core-ts's working tree, and docx4j `18a3a31f4` (`TableRules`, `TableRulesTest`). Phases B to D follow in their own briefs; CR-016's Word check is Phase D's, so Phase A's files for it are built and not staged (10.5).

**Commits.**
- `73ae725`: Phase A: `table_rules.py`, the edits, the facade, `tests/content/test_table_rules.py`, check 52's write tests to case 95, the parity script's cell `add_table`, one pre-existing expectation (10.7).
- `cfde1db`: `end_cells_with_paragraphs` without a closure, found by the cost measure (10.8).
- `c37769e` (the coordinator's, not this phase's): the eight `w:anchorLock` pins out of `tests/content/goldens.py`, docx4j-core-ts having committed its golden regeneration (`76fb45d`). It lands between the "before" and "after" columns of 10.8, and is why the after column has no failure.
- the records, the commit after `cfde1db`: this section, the Status line, the index row, the CHANGELOG's 0.2.0 bullets and `tests/README.md`'s map.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the rule | `docx4j_py/model/content/table_rules.py` (new; `model/` is in `KEEP` whole) | `ends_with_paragraph(contents)` (a `w:tc` or any block list, docx4j's two overloads) and `end_with_paragraph(tc)` (Word's repair; returns the paragraph appended, None where the rule held, where docx4j returns a boolean), exported from `docx4j_py.model.content`; for the edits `CellEnd`, `paragraph_end_of`, `wrapping_cell_of`, `ends_cell` and `end_cells_with_paragraphs` (D3) |
| the hub | `body.py`, `Body._insert_element` and `_keep_cell_end` | D3 before the insertion; after it, where the cell no longer ends with a paragraph, `keep_cell_mark_last` when a table went after the cell's end paragraph in the cell's own content, else docx4j's repair, plain under tracking, given a `w14:paraId` |
| the split | `span_edits.py`, `keep_cell_mark_last` | core-ts's `keepCellMarkLast`: the paragraph keeps its text and its element; its new `w:pPr` is `split._properties_for_head` with the mark in the formatting of the text before the end (`_r_pr_before`), inserted under tracking; its old `w:pPr` on an empty paragraph appended after the table |
| the deletes | `paragraph.py` (`delete`, `_delete`, `_tracked_delete`, `_ends_its_cell_alone`), `body.py` (`clear`, `_clear_content`), `controls.py` (`ContentControl.delete`), `ooxml.py` (`insert_ooxml_into_body` at `"Replace"`) | section 4.2's rules, core-ts's code for code |
| `TableCell.value` (Q11) | `table.py` | `body._clear_content(keep_cell_end=False)`, tracked, then `insert_paragraph` |
| the span edits | `span_edits.py`, `_removable` and `_delete_mark` | "a cell's last" is `ends_cell` (section 4.8 item 1), where it was `mark_text == "\t"` |
| the facade (Q13) | `docx/table.py` (`_Cell.add_table`), `docx/_base.py` (`_add_table`), `table.py` (`insert_table_into(..., target=)`) | the paragraph first, the table before it |
| the oracles | `tests/content/test_table_rules.py` (40 tests), `test_ranges.py`, `test_python_docx_parity.py` | 10.3 |

### 10.2 The rules as built

1. **The rule** (docx4j's): the last block among a cell's items, a block being a `w:p`, `w:tbl`, block `w:sdt`, block `w:customXml` or `mc:AlternateContent`, everything else range markup; through a control's content, custom XML's and the branch `block_children_of` reads of a kept `mc:AlternateContent` (`mce_branch`, docx4j's `McMode.READ`), down to a `w:p` (it holds) or a `w:tbl` or nothing (it breaks). An untyped (wildcard) `mc:AlternateContent` reads as breaking it; none is in a cell over the corpus.
2. **The paragraph a cell ends with** (`ends_cell`): up the parent pointers from the paragraph through the wrappers (a block control, its `w:sdtContent`, custom XML, an `mc:AlternateContent` and its branches) to a `w:tc`, and the rule's paragraph is this one. A paragraph in a text box, a header or the body has no cell.
3. **`Paragraph.delete()`**: untracked, the paragraph comes out; if its cell then breaks the rule it goes back where it was, its content emptied and its `w:pPr` kept, and the report records the text gone and no shift. Tracked, its text is deleted and its mark left unmarked whenever it is the cell's end paragraph (cases 90 and 91), whether or not a paragraph comes before it; this author's own empty insertion is taken back, as before, unless its cell would break.
4. **`Body.clear()`**: untracked, everything goes; if the cell then breaks the rule, the paragraph it ended with comes back, emptied, at the end of this body's content (a cell's body: outside any control it was in; a control's body: inside the control), else docx4j's repair. Tracked, every paragraph goes through `_delete`, so the cell's end paragraph keeps its mark. `_clear_content(keep_cell_end=False)` is the old behaviour for `insert_text(..., "Replace")` and `TableCell.value`.
5. **`ContentControl.delete()`** of a block control in a cell: where the cell then breaks the rule, the paragraph it ended with, emptied, takes the control's place; a cell left ending in a table some other way (a control holding only a table, unwrapped) gets docx4j's repair. Either `keep_content`.
6. **`Body.insert_ooxml(..., "Replace")`** in a cell or a cell's control: content ending in a table is put before the kept paragraph (case 94); content ending in a paragraph replaces it, as before.
7. **A table after the cell's end paragraph** in the cell's own list (`insert_table`, `insert_xml`, `insert_ooxml`, `insert_markdown`, `insert_element` at `"End"`, and `Paragraph.insert_xml`, `insert_ooxml`, `insert_markdown`, `Range.insert_ooxml` at `"After"`): `keep_cell_mark_last`. In a control's list instead, docx4j's repair after the control.
8. **D3**: `end_cells_with_paragraphs` walks what goes in through tables (every row and cell, through row- and cell-level controls and custom XML and every branch of a kept `mc:AlternateContent`), block controls and custom XML, never into a paragraph, nested cells first, before the insertion, so under tracking the repair paragraph is marked inserted with the table (`["TBL+ins", "P[|ins]"]`).
9. **D4**: nothing runs on load; an untouched part is copied byte for byte; the first edit through `insert_element` into a broken cell repairs it.

### 10.3 The oracles

**Check 52, cases 86 to 95** (`RANGE_CASES` 85 to 95; Word 365's run, its save after the run and its Accept All, `tests/office_js/`):

| test | differing regions of 86 to 95 | why |
|---|---|---|
| the run (`run_check_52(..., "run")` against Word's run) | 86, 87 (`paragraphs`); 94 (`result`) | 86 and 87: Word gives the new table's own paragraph the cell paragraph's `w:pPr`, and so its numbering, which is `insert_table`'s wherever it inserts (Q4); 94: this package's `Body.insert_ooxml` returns the views (the table, read `NEWOOXML`), Office JS a `Range` (`NEWOOXML\r\n`). Both core-ts's. Word's `Other` style for the cell's paragraph after a nested table (89, 91, 94) is read as its markup says, List Paragraph (`_other_as_ours`, core-ts's `listParagraphForOther`; CR-013 Q6) |
| the markup of each region against Word's save | 86 and 87, on line 3 alone, pinned: Word's `    P[ListParagraph/num|u] "NEWCELL"{u}` and `    P[ListParagraph/num|ins,u+change] ins:"NEWCELL"{u}` against this package's `    P[|] "NEWCELL"` and `    P[|ins] ins:"NEWCELL"`, core-ts's lines exactly | Q4. 88, 89, 90, 91, 92, 93, 94 and 95 are Word's line for line; 88, 90, 94 and 95 pinned |
| Accept All of this package's save against Word's Accept All | 86, 87 | Q4 again; every cell's last paragraph accepted is Word's |

Before Phase A (CR-016 section 2.2) the run differed in eight of the ten, the markup in all ten and Accept All in eight. The tests' other expectations are unchanged: the run's 26 regions differing before case 86, Accept All's 49 to 56, 60 and 66, the markup's 1, 13, 60, 66, 67, 68 and 79 to 85.

**`test_table_rules.py`** (40 tests):
- docx4j's `TableRulesTest`, its four tests by their names, with its twelve probes and the empty cell, each through the `w:tc` and through its list. core-ts's first two tests are the same four, merged, and are these.
- core-ts's other five, their strings and shapes kept (`P[style,jc|mark] text`, `TBL+ins`, `SDT{...}`). One answer differs, recorded (10.6 item 1): `Range.insert_ooxml(table, "Replace")` over the cell paragraph's `"Content"` or `"Whole"` range.
- core-ts's table of insertions at a cell's end, extended from its three verbs to eight, each untracked, tracked, accepted and rejected.
- This package's own: Q11; a control's delete keeping its content, and a broken cell unwrapped; a control's body cleared or replaced in a cell; a kept `mc:AlternateContent` in a cell (Word's branch read; a `Paragraph.delete()` inside it refused `address.read_only`, CR-011 Phase F's standing departure; an untracked `clear()` keeping the Choice's paragraph, emptied, moved out into the cell; a tracked `clear()` refused); D3 under tracking; the repair plain under tracking; custom XML and range markers; the facade's `_Cell.add_table`; one `ChangeReport` per write and its `dry_run` for seven writes, untracked and tracked (but a control's delete, which is not tracked): 13 cases, the real package's main part unchanged by each dry run, and every cell ending with a paragraph in the trial and after the commit; the exports.

**The facade** (Q13): `test_python_docx_parity.py` gains a script of its own, a cell's text underlined, `cell.add_table(2, 2)` holding `INNER`, then `cell.add_paragraph("AFTER")`, run through both libraries and read back by python-docx 1.2.0: the cell's inner content is the same, `[p "CELL" (no w:pPr, underlined), tbl 2x2, p "" (no w:pPr, no run), p "AFTER"]`. Run in a throwaway environment (`uv venv` in the session's scratchpad, python-docx 1.2.0, docx4j-xsdata 26.2.3 from PyPI, markdown-it-py, pytest; the checkout imported from its `tests/content/conftest.py`): 14 passed at `73ae725`, and 14 at `955547c` with the new test copied in, since `955547c`'s facade added the table at the end before any rule existed. Without Q13's change the new code would split the cell's paragraph and then add python-docx's paragraph too.

### 10.4 The cell sweep (section 6.2)

`out/cr016-a-sweep/cell_sweep.py` (git-ignored, with its two JSONs and `compare.txt`): over `out/cr012-b-corpus/` (`COMMITS.txt` unchanged), its 422 distinct documents by SHA-1 (4 Strict ones fail to load, as in every audit), every one of the **2,435 main-part cells** (every depth, through controls and Word's branch; section 2.1 counted 2,495 another way), given nine writes, each on a fresh load, untracked (`change_tracking_mode` set `"Off"` explicitly, since a document whose settings track revisions loads tracking them) and tracked: `Paragraph.delete()` of the cell's last own paragraph, `body.clear()`, `insert_table(..., "End")`, `insert_xml` and `insert_markdown` of a table at `"End"`, `insert_ooxml(table, "Replace")`, `ContentControl.delete(keep_content=False)` of the last block control of the cell's own list (11 cells have one), `TableCell.value`, and `Range.delete()` of the cell's whole range. Each saved, loaded again and accepted; tracked, also rejected. The cell's shape (every block, each paragraph's style, alignment, numbering, `w:pPrChange`, mark properties and text with its revisions), whether it ends with a paragraph (the script's own reading of docx4j's rule, the same on both codes) and whether its last mark is marked deleted. On `955547c` (in a `git worktree` in the session's scratchpad) and on the code of `73ae725`; 635 s and 647 s on 12 processes.

| write | cells edited (tracked) | before: the cell ending without a paragraph (holding no block) | before, tracked: the cell's last mark marked deleted by the write | after |
|---|---:|---:|---:|---|
| `Paragraph.delete()` | 2,427 (2,232) | 2,396 (2,383) | 2,232 | every cell ends with a paragraph, untracked and tracked, accepted and rejected |
| `body.clear()` | 2,429 (2,427) | 2,429 (2,429) | 2,232, and 195 the source's already | the same |
| `insert_table`, `insert_xml`, `insert_markdown` at `"End"` | 2,429 each (2,429 tracked) | 2,429 each, tracked too; Accept All 1,908 | --- | the same |
| `insert_ooxml(table, "Replace")` | 2,429 (2,427) | 2,429, tracked 2,427; Accept All 1,908 and 1,906 | --- | the same |
| `ContentControl.delete(keep_content=False)` | 11 | 2 (2) | --- | the same |
| `TableCell.value` | 2,429 (2,427) | 0 | --- | the same |
| `Range.delete()` of the whole range | 2,427 (2,427) | 0 | --- | **unchanged in every cell**, untracked and tracked |

- **Every row that moved is explained**, mechanically (`compare.txt`, 0 unexplained): the old code left the cell without a paragraph end or without a block (every untracked and tracked insert, `clear`, `insert_ooxml`, `Paragraph.delete`, the 2 control deletes), or marked its last mark deleted (2,232 tracked `clear`s and 2,232 tracked `Paragraph.delete`s; the 195 other cells' last marks the source deletes already); or, for a tracked `TableCell.value`, the old content is now a tracked deletion (2,427, Q11).
- **The refusals are the same, counted by code**, but for two: a tracked `TableCell.value` over a cell holding a kept `mc:AlternateContent` is now refused `address.read_only` (2 cells), as a tracked `clear()` of it always was, where the old setter removed it untracked. Otherwise `address.read_only` 6 or 8 per write (cells in a kept branch) and `tracking.already_deleted` 195 (a tracked delete of a paragraph whose mark the source already deletes), before and after.
- **Accept All of the tracked write equals Accept All of the untracked one** in every cell but 25: 13 `Paragraph.delete()`s of a cell's last paragraph after another paragraph (untracked it goes, the cell ending with the one before; tracked its text goes and its mark stays, so Accept All leaves it empty: core-ts's rule, Word unmeasured in this shape), and 4 each of `clear`, `insert_ooxml(..., "Replace")` and `value` in a cell whose content is a block control (a tracked deletion of a control's content leaves the control, CR-003 section 4). Before: 1,905 and 1,906 of `Paragraph.delete` and `clear` differed.
- **Reject All of the tracked write gives the document's own Reject All** in every cell but: the 3 controls deleted (a control's delete is not tracked); 4 each of the table inserts and 2 `insert_ooxml` Replaces in a cell whose last block is a block control or a kept `mc:AlternateContent` (the table goes after the wrapper and docx4j's repair paragraph after it, plain under tracking, stays when the table is rejected); and 74 tracked `value`s, 72 of them by the restored paragraph's mark properties alone (CR-003 section 16.10's final-mark shift records the `w:pPr` in a `w:pPrChange`, which has no mark properties, so a mark's `w:b` or `w:i` does not come back) and 2 the new paragraph after a cell's control (nothing before it to carry its mark, so it is not marked). Before, a tracked `value` failed to give the cell back in 1,691 cells.
- **A kept paragraph keeps its `w:pPr`, a deleted mark among it**: 195 cells end with a paragraph whose mark the source already deletes, after every write that keeps or leaves it (196 for `Paragraph.delete`: one cell whose last paragraph went, leaving the one before, whose mark the source deletes). Word reads such a mark at a cell's end as the cell's end (check 58 cases 9 and 10).
- 195 cells (69 to 195 by write) stand in rows the document's own revisions remove, so Accept All or Reject All of the save has no such cell; they are counted as gone, not compared.

### 10.5 The Word check's files (built, not staged)

`out/cr016-a-word-check/` (git-ignored): `build.py`, its `README.txt` with what Word should show, and `MANIFEST-*.txt` with the SHA-1s and each file's cells breaking the rule, every file read back strictly.
- **File 1** (`73ae725`'s code): `cr016a-1-untracked.docx` (21 regions: section 2.1's edits on a cell holding a numbered, underlined list item and on one holding a nested table and then that item), `cr016a-1-tracked.docx` (19, a control's delete not being tracked), the engine's `-tracked-accepted.docx` and `-tracked-rejected.docx`, and `cr016a-1-facade.docx` (`_Cell.add_table`). 0 cells breaking the rule in each.
- **File 0** (Q14, `955547c`'s code, the same script): `cr016a-0-at-955547c-untracked.docx`, 16 cells breaking the rule (every insert at the end, both deletes, `clear`, the Replaces and both control deletes); Word should repair it. The same run's tracked files (10 cells breaking it, its Accept All 10) and facade file are kept beside it for reference.
- Phase D adds Phase B's placeholder file and stages them together.

### 10.6 Departures and choices, with why

1. **`Range.insert_ooxml(table, "Replace")` over a cell paragraph's range** gives `[TBL, P(emptied, its w:pPr)]` here, core-ts `[P(its text gone, mark u), TBL, P(its w:pPr)]`. This package's `Range.insert_ooxml` deletes the range first and puts the blocks before the paragraph it leaves (CR-013 Phase C); the result keeps the rule and is Word's form for `Body.insertOoxml` (case 94). Unmeasured in Word; asserted as this package's.
2. **`end_with_paragraph` returns the paragraph**, docx4j a boolean, as core-ts and section 4.2 have it; truthy where docx4j's is true.
3. **The span edits' test of a cell's last mark is `ends_cell`** (section 4.8 item 1), where it was `mark_text == "\t"`. The two agree in every cell that keeps the rule; in a broken one (the last block a table) `mark_text` still reads the last paragraph as the cell's end, and a span edit can now take its mark, as Word reads it. No corpus cell is broken, and `Range.delete()` moved in none (10.4).
4. **A tracked delete of a cell's end paragraph never marks its mark**, even with a paragraph before it in the cell, where the untracked delete removes it: core-ts's rule (`trackedDelete(tracker, cellEnd)`), which section 4.2 adopts. Word measured only the paragraph alone at the top of its cell (cases 90, 91). The sweep's 13 cells are this.
5. **Q11's tracked `TableCell.value`** marks the old paragraph's mark deleted and, by this package's final-mark rule (CR-003 section 16.10), inserted too, the new paragraph keeping the cell's final mark unmarked with a `w:pPrChange` of the old properties; core-ts marks the new paragraph's own mark inserted. Accept All gives the value, Reject All the old content (but a mark's own run properties, 10.4); the Word check reads the form (10.5).
6. **D3 skips paragraphs**, as core-ts's `endCellsWithParagraphs` does: a table in a text box inside an inserted paragraph is not repaired.
7. **The facade's `_Cell.add_table`** uses an internal `target` on `insert_table_into`; `Body.insert_table` keeps Office JS's two locations.
8. **Choices for the user**: none taken against core-ts or docx4j but item 1, recorded.

### 10.7 Pre-existing tests whose expectation changed

One, in `73ae725`: `tests/content/test_word_writers.py::test_the_one_paragraph_writers_write_crlf_as_one_break[TableCell.value-tracked]` allows one more paragraph for a tracked `TableCell.value`: the new table's own empty paragraph stays as a tracked deletion until accepted (Q11), where the setter removed it untracked. In `test_ranges.py` the three write tests now run to case 95 and name 86, 87 and 94 (10.3); nothing they held before 86 moved. No golden, ratchet row, Word-save comparison or acceptance artefact moved.

### 10.8 The gates

| gate | before (`955547c`) | after (`cfde1db`, on `c37769e`) |
|---|---|---|
| the fast suite (`-m "not slow"`) | 4,910 passed, 8 skipped, 1,307 deselected, 3 failed (2 min 37 s): the 3 goldens whose 8 `w:anchorLock` pins XPASS(strict) under docx4j-core-ts's then uncommitted regeneration ("3 known") | 4,953 passed, 8 skipped, 1,307 deselected, 0 failed (2 min 38 s) |
| the full suite | 6,217 passed, 8 skipped, 3 failed, the same 3 (4 min 4 s) | 6,260 passed, 8 skipped, 0 failed (4 min 8 s); at `73ae725` the same, 4,953 and 6,260 |
| the new tests | --- | 40 in `test_table_rules.py`; `test_ranges.py` 37, as before |
| CR-011's ratchet (`tests/text/conformance.py`) | 1,291 passed, 0 failed, 293 pending of 1,584; goldens 12 of 12 | the same; `conformance.tsv` byte for byte |
| `tests/text` | 1,212 | 1,212 |
| `tests/revisions`; CR-010's Word-save comparisons (`test_word_saves.py`, `test_compare_word_resolution.py`) | 614; 76 | 614; 76; `scripts/compare_word_resolution.py` 0 lines either side, on the acceptance copy before and the regeneration after |
| `roundtrip.py` over `samples/` | `--normalise` and `--no-normalise`: 50 of 50, 0 differences, 34 `boolean-spelling`, 0 kept, 0 normalised, 0 skipped | the same, both ways |
| `codegen/generate.sh --check` | --- | byte identical, 301 modules; the tree unchanged |
| acceptance (`compare_acceptance.py` against `out/acceptance-before-cr016a/`, copied first) | a regeneration at `955547c` equal to the copy: 17 artefacts, 243 entries, 0 differing beyond the dates | 17 artefacts, 243 entries, 0 differing beyond the dates: no artefact edits a cell's end |
| python-docx 1.2.0's parity (10.3) | 13 passed (14 with the new test) | 14 passed |
| the cell sweep (10.4) | 14 of its 18 writes (nine, untracked and tracked) leave cells without a paragraph end or with their last mark marked deleted: each insert at the end 2,429, untracked and tracked; `clear` 2,429; `insert_ooxml` Replace 2,429 and 2,427 tracked; `Paragraph.delete` 2,396; a control's delete 2, both ways; a tracked `clear` and `Paragraph.delete` the last mark of 2,232 each | 0 cells ending without a paragraph, after any write, Accept All or Reject All |
| the cost: `out/cr016-a-sweep/bench_cells.py <root>`, three rounds interleaved, the median of 7 runs each (ms) | 2,000 `insert_paragraph` in the body 44.3 to 46.3; 500 in a cell 12.8 to 16.9; 1,000 inserts and 999 `Paragraph.delete`s in the body 141.7 to 145.7; 300 and 300 in a cell 22.7 to 23.2; 300 `insert_xml` of a table in the body 177.2 to 179.3; 100 in a cell 60.8 to 61.2 | 47.6 to 48.6; 14.5 to 15.1; 145.1 to 146.9; 24.0 to 24.6; 176.8 to 180.2; 64.8 to 66.0 |

The suites' difference is the 40 new tests and the 3 goldens `c37769e` released. The cost is about 1.5 µs an insert and 1 µs a delete in the body, within the rounds' noise elsewhere but for a table inserted in a cell: `wrapping_cell_of` (0.3 µs where there is no cell) and `end_cells_with_paragraphs` (0.27 µs over a paragraph, once its closure went, `cfde1db`) on every insert, `ends_cell` on every delete; in a cell, the end paragraph found before the insert and the rule read after it, and a table's D3 walk (about 40 µs more for a table inserted in a cell, the split included).

### 10.9 Found, and not done

1. **A tracked `TableCell.value`'s Reject All loses the old paragraph's mark properties** (72 cells of the sweep): the final-mark shift's `w:pPrChange` records a `PPrBase`, which has no `w:rPr`. CR-003 section 16.10's, not Phase A's; the text, the paragraph properties and the cell come back.
2. **docx4j's repair under tracking stays after Reject All** where a table went after a cell's last control or kept branch (6 cells of the sweep): plain, as D2 and core-ts have it; marking it inserted would make the cell's final mark an inserted one, which Word cannot reject (CR-003 section 16.10).
3. **A tracked deletion of a control's content leaves the control** (4 cells of the sweep), so Accept All of a tracked `clear()` keeps an empty control where the untracked one removes it: CR-003 section 4's, standing.
4. **An own empty insertion holding an empty run is not taken back** by a tracked delete (`_tracked_delete`'s test is "no content"): a tracked `value` on a table this author has just inserted tracked leaves the cell's old empty paragraph as a tracked deletion (10.7). core-ts's test is the same.
5. **`Paragraph.mark_text` still reads a broken cell's last paragraph as `\t`** (10.6 item 3). No reader moves in Phase A.
6. **CR-014 phase 2 and CR-011 Phase E**, when they build cells, end them with `end_with_paragraph` (section 4.2); Phase D notes it in both CRs.
7. During the sweep this session deleted one file of its own outside the scratchpad, `out/cr016-a-sweep/try.json`, a 20-document trial of the sweep it had just written; recorded because the brief allowed deletions in the scratchpad only.
8. **`CLAUDE.md`** should name Phase A (left to the coordinator, per the brief; the sentence proposed is in the report).
