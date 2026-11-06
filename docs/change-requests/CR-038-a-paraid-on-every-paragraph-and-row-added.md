# CR-038: A `w14:paraId` on every paragraph and row added (the port of core-ts's CR-002 section 66)

**Status:** Decided 2026-11-06 (the user: "CR-038 as recommended, please implement"): Q1 replace, Q2 a new header or footer takes ids and the glossary doesn't, Q3 this release. Proposed the same day ("write the CR"), on docx4j-core-ts `1e02687` and the portfolio's DI-026. It gives every paragraph and table row this package adds a `w14:paraId` where the document uses them, since one missing makes Word renumber every id in the file. **Phases A and B built the same day (section 9); the Word check is the user's.** Three Word rounds are in section 8, and they settle Q2's case: a new footer with no ids renumbers a whole file, and what protects a file is a comment.

**Depends on:** nothing here. The rule is core-ts's Word checks 79 and 80 (`4600413`, `test/check79/`, `test/check80/`).

**Counterpart:** docx4j-core-ts CR-002 section 66, built (`1e02687`), whose `assignParaIds` this follows. docx4j has the same gap, and whether it changes is the user's call there. The portfolio records all three as DI-026.

## 1. Why

Word 365 keeps a file's ids at open only if every paragraph and table row holds one. One missing and it renumbers the whole part (Word 16.0.20326.20232, 2026-11-06, one run each, through the Word bridge):

| File | Ids Word kept |
|---|---|
| Word's own save | 17 of 17, through two saves |
| A built file, one paragraph's id removed | 0 of 11 |
| A built file with a table whose row has no id | 0 of 14 |
| The same file with an id on the `w:tr` | 14 of 14 |
| core-ts's output, every new paragraph and row with an id and no `w14:textId` | all, in all four files |
| The same with one new row's id removed | 0 of 28 |

So `w14:textId` isn't needed, and the row's id counts as the paragraph's does. Section 8 refines the rule, and every result so far fits it: **in a file with no comments, one missing id in any story makes Word renumber every id in the file. A file with comments keeps its ids and fills the gap.** Compatibility mode plays no part, and a new footer is enough, so most of this package's table and part routes are affected.

Here a `w14:paraId` is the stable address (CR-003 section 3.4): `Paragraph.address`, `paragraph_at(para_id=...)`, `created_para_ids`. A caller who holds one expects it to last. Today it lasts until Word next opens a file this package added a table to. Nothing is corrupted, as far as these checks show: what goes is every id in the file, not just the new ones.

## 2. What exists here (measured 2026-11-06 at `3bec61e`)

Each route ran from a fresh load of `tests/fixtures/check76/76t18-hosts-control-word365.docx`, Word's save, with an id on every `w:p` and `w:tr` in the body, header, notes and comments, and three tables. The count is on the saved bytes, part by part. A read and save lost nothing.

**29 of 80 routes add a paragraph or row with no id:**

| Cause | Routes |
|---|---|
| A table made or extended: the rows and their cell paragraphs | `Body.insert_table` (Start, End, values, tracked, in a cell, in a control); `Table.add_rows` (Start, End, values, tracked); `TableRow.insert_rows` (Before, After); a Markdown table; the facade's `add_table`, `Table.add_row`, `_Cell.add_table` |
| Anything nested in a fragment | `insert_ooxml` and `insert_xml` with a `w:tbl` (Body, Paragraph, Range, cell); `insert_ooxml` with a `w:p` in a `w:sdt` or a text box |
| Single routes | the facade's `Table.add_column` and `add_section`; a comment's second paragraph onward (`insert_comment("a\nb")`, `Comment.content`) |
| A new part | the facade's `footer.add_paragraph` where no footer exists (2 paragraphs); `ContentControl.show_placeholder` (the glossary, 1) |

The other 50 add none without an id, but for `ContentControl.insert_copy_after`, which is the 80th, below. They are the paragraph routes: `insert_paragraph` everywhere, `insert_text` with a break, page and section breaks, a `w:p` through `insert_ooxml` and `insert_xml`, Markdown other than a table, pictures, one-paragraph comments and replies, and the facade's equivalents.

**Beside those, measured the same way:**
- **Duplicates.** `insert_copy_after` repeats 2 ids. Copying through `get_xml` repeats them too: `body.insert_xml(paragraph.get_xml())` 1, `body.insert_xml(table.get_xml())` 3, and a fragment naming an id already in the file 1. Word's reaction to a duplicate isn't measured.
- **A new cell's own paragraph.** `add_rows(1)[0].cells[0].insert_paragraph("x")` leaves that paragraph with no id either, in a file full of them, since the cell's `Body` sees only the cell.
- **`Body.ensure_para_ids`** skips rows: after `insert_table` it left 0 paragraphs and 2 rows without an id. It skips text boxes too.
- **`create_package`.** `insert_paragraph` then `insert_table` saves 4 paragraphs and 2 rows with none, so a file made here is affected.

**The code** (`model/content/addresses.py`):
- `assign_para_id(body, element)` takes one paragraph. `Body.insert_element` calls it only `if isinstance(item, P)`, so a table's content, and anything else nested, is passed over. `Table.add_rows` and `TableRow.insert_rows` insert into the table directly and don't call it.
- `assign_para_id` returns at once where `element.para_id` is set, so a copied id stays.
- `_taken` is per `body.prefix`, and `assigns_para_ids` asks whether that body's set is empty, so a body with no ids yet answers no.
- `para_ids_of` walks `iter_paragraphs`, which sees no row and no text box. `iter_children` reaches both, a text box in a kept branch included.

## 3. Design

### 3.1 One walker

`assign_para_ids(body, element)` replaces `assign_para_id`: every `P` and `Tr` at or under `element`, at any depth, in the branch Word reads (`iter_children`'s default, `mce="resolve"`), in document order, takes an id where it needs one. A `mc:Fallback` isn't entered: Word reads the `mc:Choice`'s id and writes it over the `mc:Fallback`'s (section 8), so an id there is never read. One generator and one `taken` set serve the walk. It returns the ids it gave.

### 3.2 The story decides, not the `Body`

As core-ts found, a cell's own `Body` would read as a document with none. `_taken` becomes per **part**, read from the part's whole tree, rows and text boxes included, so `para_ids_of` sees what it has been missing. `assigns_para_ids` asks that set. The override stays, and `create_package` keeps setting it.

Reading a part already being edited costs no byte-for-byte promise, and the set is still built once and kept in step, so two thousand new paragraphs stay linear. The walk now enters runs, once a part: Phase A times it on `Symbols.docx`.

### 3.3 The seams

| Where | Change |
|---|---|
| `Body.insert_element` | the walker on every item, not only a `P`. This covers `insert_table`, `insert_ooxml`, `insert_xml` and the Markdown importer, which reach it, and the cell paragraphs `end_cells_with_paragraphs` adds, since it runs first |
| `Table.add_rows`, `TableRow.insert_rows` | the walker on each row added |
| The five other `assign_para_id` callers (`body.py` twice, `span_edits.py`, `controls.py`, `split.py`) | the walker, on the one paragraph |
| `ContentControl`: a control's text over several paragraphs, and the first paragraph made in an empty one (`controls.py`) | the walker. Neither is a gap in the shapes measured, so a test holds each, as the read says they could be |
| `ContentControl.insert_copy_after` | the walker on the copy (3.4) |
| The facade's `Table.add_column` and `add_section` | the walker on the new cells' paragraphs and the new paragraph |
| `Comment`'s writers (`comments.py`) | every paragraph of a comment, not only the first, through the comments' own `_para_ids_taken` |

Ids go on at creation, not on save, as core-ts's do: a save reports and doesn't rewrite, and an id is wanted as soon as the content exists, for its address.

### 3.4 An id already there (Q1)

The walker keeps an id that's free, and replaces one the part already holds. So `insert_copy_after` and a fragment copied through `get_xml` take fresh ids, and a fragment's own unique ids stay. This goes beyond core-ts, which fills only what's missing.

**A repeat Word writes itself isn't a collision.** Word's own save gives a `mc:Choice`'s text box and its `mc:Fallback`'s copy the same ids (9 in `document.xml` of `extraction-probe.word-resave.docx`, and 1 in a header). The walker doesn't enter a `mc:Fallback` (3.1), so it never meets the repeat, and a fragment holding such a pair keeps it. Found after this section was first written, building the mode round.

**Word does the same itself.** Given two paragraphs with one id, it kept the first and gave the later one a new id, with no prompt (section 8). Here the copy takes the new id wherever it lands, even before its original, since that's what keeps an address a caller already holds.

The reason doesn't wait on Word: two paragraphs answering one `paragraph_at(para_id=...)` is this package's own bug, whatever Word does with the file. `taken` isn't reduced on a delete, so an id deleted and re-inserted through `get_xml` comes back as a new one. That loses nothing a caller could rely on, since the old paragraph had gone.

### 3.5 Parts other than the main document (Q2)

An existing part that uses ids follows 3.2, comments included. The open case is a new part, where there's no tree to ask: a new footer, and the glossary. Word's rule does reach past `document.xml` (section 8), and core-ts's section 66 leaves those parts out.

Recommended: a new header or footer takes the main part's answer. It has to: a new footer with no ids renumbered all 18 of a file's ids, and with ids kept them (section 8). It would read the main part privately, as the text model does, so it marks nothing. **The glossary stays without,** since Word writes none there and drops one it's given.

### 3.6 `ensure_para_ids`

It uses the walker, so rows and text boxes get ids too, and after it the rule holds for that part. Its docstring says so, and still warns that it re-marshals. It stays the caller's tool: nothing calls it for them.

### 3.7 The report

`ChangeReport.created_para_ids` holds every id a call gave, rows included, so `insert_table(2, 2)` lists six. Its name stays. `tests/content/test_word_writers.py` pins a count there, which moves only where a route gains ids.

### 3.8 Not here

- **A file that arrives with a gap isn't repaired.** A save doesn't rewrite, so Word renumbers it as it would have. `ensure_para_ids` is there for a caller who wants it.
- **A document using no ids stays so,** but for `create_package`'s override.
- **`w14:textId`** isn't written, as today, and check 80 shows Word doesn't need it.
- **Code that appends to a tree itself** (`tbl.content.append(...)`) isn't a route. It's the caller's.
- **`apply_bindings`** isn't measured. It writes through the control's writers, so 3.3 should reach it, and a test holds that.

## 4. Tests and oracles

`tests/content/test_para_ids.py`, over `76t18`'s Word save:

- **Every route of section 2,** the 80 and the copies: no paragraph or row without an id in the part it touched, none repeated, and every one of Word's ids still there, on the saved bytes.
- **A census:** every public `insert_*` and `add_*` on the content classes and the facade is in the route table, or named as adding no block, so that a new route fails until it's listed. The gap came from routes nobody listed.
- **The story:** a new cell's paragraph; a nested table; a table in a control; a text box in a fragment, in a kept branch.
- **Ids already there (Q1):** a fragment's unique id kept, a colliding one replaced, a copy's replaced wherever it lands, the original's untouched, a `mc:Fallback` not entered, a pair left as Word wrote it.
- **Other parts (Q2):** a comment's every paragraph; an existing header; a new footer and the glossary as decided.
- **A document with no ids** gets none from any table route. `create_package` gets them from all.
- **`ensure_para_ids`** leaves no row or text box without one, and returns them in document order.
- **Determinism:** the same seed and calls give the same bytes.
- **The cost:** two thousand paragraphs inserted stays linear.

The oracle is the rule of section 1, so the tests hold the count. Only Word can hold the outcome, which is the Word check.

## 5. Phases

| Phase | What | Effort |
|---|---|---|
| A | 3.1 to 3.7 and the tests | ½ day |
| B | The records: CHANGELOG, `tests/README.md`, the index, the registry, DI-026 through the portfolio; the Word check's files | ¼ day, then the Word run |

**The Word check.** check 80's shape, on this package's saves of `79c-save-1.word.docx`, a file that renumbers: rows added, a table added, columns through the facade, then all of them; the last through a save, and one with a new row's id taken off as the control. Each is read for how many ids Word kept. How it's run is the user's call.

## 6. Risks

1. **Ids shift for a seeded run.** The generator now draws for rows and cell paragraphs, so a later paragraph's id differs from today's for the same seed and calls. Nothing promises today's values, but a test that pins one will move, and `scripts/compare_acceptance.py` should be run.
2. **An address changes form.** A new table's paragraphs answer with a `w14:` address where they had a positional one. That's the point, but a caller matching the old form sees it.
3. **The rule is inferred.** It fits every result so far, but comments' part in it rests on two natural files and one comment added here, and why isn't known. The fix doesn't depend on the explanation: ids on everything is what Word's own save has, in every file measured. The Word check is a second run, on this package's output.
4. **The ports differ for a while.** Q1 goes beyond core-ts, and docx4j has the gap. DI-026 records it.

## 7. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Does the walker replace an id the part already holds, or only fill what's missing, as core-ts does? | **Replace.** Two paragraphs answering one address is a bug here, and Word does the same: measured on two files, it kept the first and gave the later one a new id, with no prompt and no other id moved (section 8). |
| Q2 | Does a part made in this session (a new footer, the glossary) take ids, or stay as it is, as core-ts's do? | **A new header or footer takes them, the glossary doesn't.** Measured (section 8): a new footer with no ids renumbered all 18 of a file's ids, and with ids kept them. Word writes none in a glossary and drops one it's given. A glossary with no ids on a file that renumbers isn't measured, but Word's own saves are that shape. |
| Q3 | Which release? | **This one.** It's about ¾ day, tables are a common route, and its Word check can share a run with CR-037's. |

## 8. Q2's Word run (2026-11-06)

The user: "CR-038 Q2 confirm word behaviour via portfolio's word-bridge". 11 inputs and 27 jobs, built in `out/cr038-q2-word-check/` (`build.py`, `compare.py`), staged and run by the portfolio session through the Word bridge, on Word 16.0.20326.20232. Every input is a Word save with ids on everything, changed in one way, each checked before it went. Two `save` runs each, and a `check` on five. All opened silently. The results are under `bridge/portfolio/<job>/` on the share, and these counts are this session's, from the saved bytes.

**All 11 sources turned out to be compatibility mode 12,** where the checks of section 1 are mode 15. That wasn't looked at when they were chosen. The mode round below shows it wasn't what mattered.

| Input | `document.xml` ids kept | The part in question |
|---|---|---|
| `q2a` control | 24 of 24 | as read |
| `q2b` a body paragraph's id off: the positive control | **23 of 23, the gap filled** | |
| `q2c` a header's id off | 24 of 24 | its other id kept, the gap filled |
| `q2d` a comment's second paragraph's id off | 24 of 24 | its other id kept, the gap filled |
| `q2e` read and saved here, no edit | 24 of 24 | as read |
| `q2f` this package's new footer, no ids | 24 of 24 | Word wrote ids on both paragraphs |
| `q2g` the same with ids | 24 of 24 | 2 of 2 kept |
| `q2h` this package's new glossary, no id | 24 of 24 | Word wrote none |
| `q2i` the same with an id | 24 of 24 | 0 of 1: Word dropped it |
| `q2j` control with a real footnote | 20 of 20 | the two real notes kept |
| `q2k` a real footnote's last paragraph's id off | **0 of 20, in both runs** | run 1: no ids in any part. Run 2: a full new set in every part |

Every save gave the separator paragraphs of `footnotes.xml` and `endnotes.xml` new ids, controls included, so they're left out of those counts. The five `check`s kept every body id at open, as the saves did.

**What it shows:**
1. **The positive control didn't fire.** `check` and `save` agree on `q2b`, so the instrument isn't the problem: the rule doesn't hold in this file. So "nothing moved" in `q2c` to `q2i` says nothing about a mode 15 file.
2. **Mode looked like the explanation, and the mode round below refutes it.** Position isn't it either: `79f`'s gap was a fifth paragraph and `q2b`'s a third.
3. **A gap outside `document.xml` can reach it.** `q2k` lost every body id, in a file where a gap in the body lost none. It's one file, the two runs disagree with each other, and it isn't understood. So the rule isn't simply per part.
4. **Ids in a new footer did no harm,** and Word writes them itself when they're missing.
5. **Word writes no ids in a glossary,** and drops one it's given, as docx4j's session counted.

**What it settles for Q2:** the glossary, and that a footer's ids are harmless on such a file.

### The mode round (2026-11-06)

The user, through the portfolio session: "Run the mode 12 against mode 15 round"; it serves docx4j's CR-077 phase 0 and core-ts's question of a row against a paragraph too. Built in `out/cr038-modes-word-check/`: each pair differs only in the mode digits in `settings.xml`, checked on every pair. 57 of 76 jobs ran, all silent. It stopped at `e-control-m12-save-1`, where Word's Compatibility Checker stopped `SaveAs` on the probe file (shapes and text boxes, collapsed comments), a dialog the bridge doesn't answer. Those saves aren't needed, since the first two rows of files answer the mode question.

Ids of the source's that Word kept, in `document.xml`. Each cell is two saves and, where there's a `check`, the reading at open, and they agree:

| File | Gap | Mode 15 | Mode 12 |
|---|---|---|---|
| `79c-save-1.word.docx` (core-ts's) | none | 18 of 18 | 18 of 18 |
| | a body paragraph | **0 of 17** | **0 of 17** |
| | the last row | **0 of 17** | **0 of 17** |
| `76t18`'s Word save (`q2b`) | a body paragraph | 23 of 23, the gap filled | 23 of 23, the gap filled |
| `76t26`'s Word save (`q2k`) | a real footnote's last paragraph | **0 of 20** | **0 of 20** |
| `extraction-probe.word-resave.docx` | a body paragraph, a row, a header, a footer, a footnote, an endnote, a keyed comment | nothing lost, each gap filled | not run |

**What it shows:**
1. **The mode doesn't decide.** Every pair behaved alike.
2. **Nor does a row against a paragraph.** `79c` renumbered on both.
3. **The file decides.** The same body gap renumbers `79c` and not `76t18` or the probe, at either mode.
4. **The lead, not a finding:** `comments.xml`, `commentsExtended.xml`, `commentsIds.xml` and a header are in both files that keep their ids, and in neither that renumbers. `commentsExtended` refers to paragraphs by id, which would be a reason for Word to keep them. But it's two files a side, and comments and headers come together in both, so the two can't be told apart yet.
5. **`q2k` is steady after all.** All five new saves wrote a full new set, and the `check`s kept 0 of 20 at open, at both modes. `q2`'s save with no ids at all is one of seven.
6. **For Q2 this is less than it looks.** Every "nothing moved" for a part other than `document.xml` comes from a file that keeps its ids anyway, `q2k` aside, where a gap outside `document.xml` reached it.

**Batch 2, at mode 15, on the probe** (8 jobs, all silent, on the bridge at `7a46241`; two saves each, agreeing), read by position in `positions.py`:
- **Two paragraphs with one id:** no prompt. The first kept it, the later one took a new id, different in each save, and nothing else changed.
- **A `mc:Choice` paragraph with no id:** Word gave it a new id and wrote the same over the `mc:Fallback`'s, so the Fallback's own id went.
- **A `mc:Fallback` paragraph with no id:** the Choice kept its id and Word wrote it over the gap. All 9 pairs came out equal in all four saves. So Word reads the Choice's id and doesn't read the Fallback's.
- **A nested row with no id:** filled, nothing else lost.

The probe keeps its ids, so these say what Word does with the id in question, not whether such a gap renumbers a file. That's unmeasured for a Choice, a Fallback and a nested row.

### The third round (2026-11-06)

The user: "tell the portfolio agent to use word-bridge to do the third round on 79c". Nine inputs in `out/cr038-r3-word-check/`, all on `79c-save-1.word.docx`, mode 15, each with two saves and a `check`, the three agreeing in every row. 27 jobs, all silent. A comment or header was added through this package's save, so `r-py-gap` is the positive control in that form.

| Input | `document.xml` ids Word kept | The new part |
|---|---|---|
| `r-py-control`: read and saved here | 18 of 18 | |
| `r-py-gap`: a body gap, the positive control | **0 of 17** | |
| `r-footer-no-ids`: this package's new footer, no ids, no body gap | **0 of 18** | Word wrote ids on both paragraphs |
| `r-footer-ids`: the same with ids | 18 of 18 | 2 of 2 kept |
| `r-comment`: a comment added, no gap | 18 of 18 | 1 of 1 kept |
| `r-comment-gap`: the same with a body gap | **17 of 17, the gap filled** | 1 of 1 kept |
| `r-header`: a header added, ids on it, no gap | 18 of 18 | 2 of 2 kept |
| `r-header-gap`: the same with a body gap | **0 of 17** | 0 of 2: renumbered too |
| `r-duplicate`: two body paragraphs with one id | 17 of 17 | |

**What it shows:**
1. **Q2's case is settled.** A new footer with no ids costs every id in the body, at open. With ids, none.
2. **Comments protect a file, and a header doesn't.** Adding one comment stopped the renumbering; adding a header didn't. That's the lead above, narrowed.
3. **The rule reaches across parts:** a gap in the body renumbered the header too.
4. **A duplicate doesn't renumber,** even here: the first kept its id, the later one took a new one, different in each save, and the other 16 stayed.

**The rule, then:** in a file with no comments, one missing id in any story renumbers every id. A file with comments keeps its ids and fills the gap. It fits every result of the three rounds and of checks 79 and 80: `79c`, `79f`, `80e`, `76t26` (`q2k`) and `r-header-gap` have no comments and renumbered, and `76t18`, the probe and `r-comment-gap` have them and kept. It's inferred: one comment added, and two files that held them. Which of the three comment parts does it, and why, isn't known. `commentsExtended.xml` refers to paragraphs by id, which would be a reason.

## 9. Phases A and B: implementation notes (2026-11-06)

### 9.1 What was built

As section 3, in `model/content/addresses.py`: `identified` (the walk), `assign_para_ids`, `record_para_ids`, `para_ids_of` and `_taken` per part, and `assigns_para_ids` asking the main part for a header or footer made here, through `read_contents`, so it marks nothing. `assign_para_id` is gone. The seams are section 3.3's, and the new header or footer's is `_add_definition`.

**Section 3.4 needed one thing it didn't say.** The set isn't reduced on a delete, and is built after the first insert has landed, so an id in it is only a possible collision. The walker counts the part's ids before replacing one, and only then: a paragraph put back in its place keeps its id, and a fragment with no ids never pays for the walk.

### 9.2 What moved

- **A new table's paragraphs** answer with a `w14:` address. One test pinned `body/1/1/1/0`, and now holds the new form.
- **`created_para_ids`** lists rows too.
- **`test_private_readers.py`** lists `_main_uses_ids` as a private reader, and a reader there reaches it.
- **Seeds:** rows are now in the set a seed derives from, so a document's generated ids differ from before for the same calls. No test pinned one.

### 9.3 Verified

- **The route probe of section 2,** re-run: 84 routes clean, where 29 of the first 80 had gaps. The one left is `show_placeholder`'s glossary, as decided, and `_Cell.merge` isn't implemented.
- **Tests:** `tests/content/test_para_ids.py`, 128, with a census of every public `insert_` and `add_` method.
- **Mutants:** 30, the walker and every seam. 23 caught at first. All seven survivors were gaps in the tests, among them two control seams nothing had reached, and all are caught now, by tests that pass.
- **The suite:** 11,628 passed, 8 skipped, with CR-035 built on top.
- **The cost:** the whole-part walk takes 33 ms on `Symbols.docx`, once a part, where the old paragraph walk took under 1 ms.

### 9.4 Not here

- **A load repair's new paragraph** has no id: CR-037's empty paragraph after a held `w:p`, and CR-030's. A repair isn't a route, and such a file already has a shape Word would read its own way.

