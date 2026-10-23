# CR-013: Ranges across paragraphs, as Office JS's `Word.Range` is --- the port of docx4j-core-ts CR-014, on Office JS's characters

**Status:** Decided 2026-10-23, every question as recommended (the user: "CR-013 accept recs for Q1 to Q11. For Q8, confirmed Apache-2.0 here"): Phase A takes Office JS's whole character set (Q1), the writers map each character back to the element it reads from and refuse U+0002 (Q2), and 0.2.0 waits for Phases A to C (Q10). It was proposed the same day, at the user's request, to port docx4j-core-ts CR-014 (ranges across paragraphs) so that the Office JS subset keeps parity with core-ts, as CR-003 section 23.3 decided ("for now let's aim to maintain parity with core-ts and see how that goes"), and because 0.2.0 waits for it (the user: "0.2.0 waits for the port, which needs a go for a new CR (CR-013, ranges across paragraphs)"). Its characters were decided the same day (the user: "core-ts's ranges across paragraphs rely on Word's own characters: \r at the end of each paragraph, and U+000B for a line break - we should do that here too"): a paragraph mark is `\r` in `Body.text` and in multi-paragraph text, a `w:br` is U+000B, and `Body.text` takes Word's form, as core-ts CR-012 phase 4 built it. That supersedes CR-011 section 15 item 2, which records it (`c4ca052`), and it is Phase A here, before the range work. Four phases: A, Office JS's characters (section 4.1); B, the reads (core-ts phase 1, `ef57729`); C, the writes (core-ts phase 2, `53db5bc`); D, the Word check and the records. Section 4.1 also takes two findings of the same day, from core-ts's probe of its writers (`87bbb38`) and the same probe here: every text writer is in the map, and a bound control's write-back keeps docx4j's characters. Phase A started 2026-10-23 at the user's go ("start Phase A") and was implemented the same day (`da7aa12`, `e7e54e6`, `316f427`, `7369c63`, `1bc9b7d`, `a465b7a`; section 10): the Office JS readers read Office JS's characters and every text writer takes them back; checks 48, 48b and 51 agree with Word 365 paragraph for paragraph (23 of 23, 25 of 25, 25 of 25); the readers' audit over 422 documents moved only answers the character table explains; the full suite is 5,389 passed, the one failure the stale Office JS subset that Phase B regenerates. Its Word check's first run (section 10.10): files 1 and 2 opened, each note explained by the markup. File 3 failed: Word recovered a block content control, bound with no type, that held two paragraphs. 0.1.0 writes the same, so it predates Phase A. Variants isolated the cause ("3a unreadable; others 3b 3c 3d open fine"): Word refuses a bound block control of no type holding two paragraphs, a content API defect 0.1.0 shares, and everything Phase A changed opens. B to D are not started.

**Depends on:**
- **docx4j-core-ts**, cited by commit only (its working tree holds uncommitted edits to the same files, a `test/check54/` among them, section 7):
  - CR-014's decisions (`a85421d`), phase 0 and check 52 (`8fdccf8`), phase 1 and check 53 (`ef57729`), phase 2 (`53db5bc`);
  - CR-012 phase 4 (`9307bdb`), the readers in Office JS's characters, held to checks 48 (`30b559e`), 48b (`5c6cb5c`) and 51 (`9307bdb`);
  - CR-002 sections 55 and 57 at `53db5bc`; sections 58 and 59 (`d3b3170`), proposed there, are referred to and not ported.
- **This repository:** CR-003 Phases B, C, F and G (`Range`, `Body`, `Paragraph`, the tables, controls, tracking and comments); CR-010 (the resolution's `join_with_next`, `model/revisions/resolver.py`, which Phase C's joins reuse as core-ts reuses its `joinWithNext`); CR-011 Phase F (the readers on Word's branch, and the edit rule over branch text); CR-012 Phases A and A2 (a view in a kept branch read-only).
- **No fork change.** No schema, generator or runtime change is foreseen; `codegen/generate.sh --check` is a gate only to prove it.
- **The sibling checkout** `../docx4j-core-ts`, for `tests/office_js_subset.json` (section 6.4) and, at the time of the port, for the oracles' sources (section 6.1).

**Counterpart:** docx4j-core-ts's `docs/change-requests/CR-014-ranges-across-paragraphs.md` at `53db5bc` (sections 5, the decisions; 6.2, check 52's run; 9, phase 1; 9.7, check 53; 10, phase 2), with its code in `src/model/content/` (`Range.mts`, `story.mts` new at `ef57729`, `spanEdits.mts` new at `53db5bc`) and its tests (`test/ranges.test.mjs`, 18 tests at `53db5bc`); its `test/README.md` checks 52 and 53; and its `docs/change-requests/CR-012-text-model.md` at `9307bdb` (sections 12 and 17, the characters and the readers). core-ts's CR-014 is the design this CR follows; where this one departs, section 4.9 says why.

## 1. Why

**The decision.** CR-003 section 23 (2026-10-23) keeps member-for-member Office JS parity with core-ts, with no pin: `tests/office_js_subset.json` stays derived from core-ts's `test/office-js-subset.ts` by `scripts/office_js_subset.py`, and a member core-ts adds is a member owed here, as `Range.hyperlink` was (CR-003 section 21). Section 23.3 makes this port its own CR and holds 0.2.0 for it. Until it lands, `tests/content/test_office_js_subset.py::test_the_script_would_write_what_is_committed` fails wherever core-ts is checked out beside this repository; it fails at `c4ca052`.

**The gap, member by member.** core-ts `ef57729` took the subset from 231 members to 241; `53db5bc` changed no member. Derived from `git show 53db5bc:test/office-js-subset.ts` into a scratch file with this repository's own script (a read-only probe) and compared with the committed JSON, the ten are:

| Office JS member | here | kind | what CR-014 makes it |
|---|---|---|---|
| `Paragraph.tableNestingLevel` | missing | read-only property | 0 outside a table, 1 in a cell, 2 in a nested table's cell |
| `Range.contentControls` | missing | read-only property | the controls the range overlaps, not one holding the whole range (check 53) |
| `Range.expandTo` | missing | method | the range from whichever start comes first to whichever end comes last, across paragraphs |
| `Range.isEmpty` | missing | read-only property | `text == ""` |
| `Range.tables` | missing | read-only property | the tables the range covers some of; inside one cell, that cell's (check 52) |
| `Table.getRange` | missing | method | `'Whole'` ending `\r\n`, `'Start'`, `'End'`, `'After'` |
| `Body.getRange` | `get_range` exists | method | `'Whole'` from the first paragraph to the last, with its mark; here it is the **first paragraph's** range |
| `Comment.getRange` | `get_range` exists | method | one `Range` across the paragraphs the comment spans (D2); here a `list[Range]`, one per paragraph |
| `Range.paragraphs` | exists | read-only property | every paragraph the range touches, a table's cells' among them; here always `[self.paragraph]` |
| `Range.styleBuiltIn` | exists (with a setter) | read-only property in core-ts's subset | check 53's rule (section 3.3); here the paragraph's style |

The probe's kinds: `Body.get_range` and `Comment.get_range` are methods, `Range.paragraphs` a property and `Range.style_built_in` a settable property; the other six are absent. The test asserts the kind, not `readonly`, so the setter may stay (section 4.3).

**Members already in the subset whose behaviour CR-014 moves.** They change no line of the JSON, but parity is behaviour too (CR-003 section 23.3):
- `ContentControl.get_range()` of a block, row or cell control: here the **first paragraph's** range (`controls.py`, `get_range`), core-ts one range over the control (section 3.4);
- `TrackedChange.get_range()`: here the change's part in its first paragraph (`tracked_change.py`), core-ts the group across paragraphs, by check 53's rules;
- `Paragraph.get_range("Whole")`: here the text alone, core-ts the text and the mark (CR-012 phase 4);
- `Range.get_range()`: core-ts adds `'After'`;
- `Range.style`: check 53's rule within one paragraph too;
- `Range.font`, `get_tracked_changes()`, `get_comments()`, `search()`, `hyperlink`, `text` and every write, across paragraphs;
- `Body.text`, `TableCell.value` and `Table.values`, a block control's `text` and `Comment.content`, in Office JS's form (Phase A).

**What it costs today.** core-ts's CR-014 section 1 calls its one-paragraph `Range` "silently partial", and stopped it with `SpanningRange` (its CR-012 section 17.12) before porting ranges. This package never had that stopgap, and 0.1.0 on PyPI has the defect:
- `control.get_range().delete()` on a block control of three paragraphs deletes the first paragraph's text only, and `get_range("End")` is that first paragraph's end (measured, section 2.3);
- `body.get_range()` reads and edits the body's first paragraph;
- an add-in's or an agent's code that takes a range from a control, a comment, a change or `expandTo` can't run here.

**The characters.** core-ts's ranges across paragraphs read and are measured in Office JS's characters: check 52's every expected text (`HEAD3B\rBODY3A`, a table's `\t` and `\r\n`, a note reference `\u0002`) is Word's character stream. This package's readers kept their own characters (CR-011 section 15 item 2), and the user has now decided otherwise (the Status). So Phase A ports core-ts CR-012 phase 4's readers first, as far as Q1 decides.

## 2. What exists here (measured 2026-10-23 at `7eb481b`; the code is unchanged at `c4ca052`)

### 2.1 The `Range`

`docx4j_py/model/content/range.py` (677 lines): "a span within one paragraph". `__slots__ = ("end", "paragraph", "start")`; `start` and `end` are code points into the paragraph's accepted-view text (CR-003 decided question 3), recomputed from the tree on every use. There is no end paragraph, no mark and no story.
- `text` is `self.paragraph.text[self.start:self.end]`; `paragraphs` is `[self.paragraph]`; `style`, `style_id` and `style_built_in` are the paragraph's.
- `font` reads the first run's effective properties; a write splits the runs on grapheme boundaries (CR-003 section 3.12).
- `get_range(location)` takes `RangeLocation = Literal["Whole", "Start", "End", "Content"]` (`enums.py`), and `"Whole"` and `"Content"` are the same range.
- `get_tracked_changes()` filters the paragraph's listing by overlap; `get_comments()` reads that paragraph's comment markers (`comments._ids_in_range`).
- Every write is one paragraph's `splice`. The Phase F edit rule (`refuse_branch_text`) applies to every one.

### 2.2 The characters, against Word 365

`Paragraph.text` is `text_of_view(element)`, the text of `text_model.segments_of`, whose `_CONSTANT_TEXT` and `item_text` give: `w:t` its value, `w:tab` `\t`, `w:br` and `w:cr` `\n` (every break type), `w:noBreakHyphen` U+2011, `w:softHyphen` U+00AD, `w:sym` the character its `w:char` names. A note reference, a `w:ptab`, a picture and an equation give nothing; hidden text is read. `Body.text` joins the paragraphs with `\n`, each cell's paragraph a line of its own (`body.py`, `text`). A read-only probe loaded the very files Word 365 read, from the share (section 6.1), and compared each body paragraph's `text` with Word's JSON:

| check | file | paragraphs equal to Word's | the others |
|---|---|---|---|
| 48, view A (Word's default view; `48-text-views-word365.txt`, `5c6cb5c`) | `__tmp/48/48-text.docx` | 13 of 23 | a note reference (U+0002), the four breaks (U+000B, U+000C, U+000E, U+000D), `w:ptab` (`0`), the two hyphens (U+001E, U+001F), a Wingdings `w:sym` (`(`), hidden text (out), the equation (` 𝑥=1 `) |
| 48b, take 2 (`48b-text-word365.json`, `5c6cb5c`) | `__tmp/48/48b-text.docx` | 13 of 25 | the same, and case 19's symbols and case 20's run hidden through its character style |
| 51 (`51-text-word365.json`, `9307bdb`) | `__tmp/51/51-text.docx` | 16 of 25 | a note reference, and the equations of cases 6 to 13 |

**Over the corpus.** A read-only census of the 418 distinct documents (by content) of `out/cr012-b-corpus/` that parse, the corpus CR-011 section 16.5 audited (`COMMITS.txt`: this repository at `f1bb1d9`, core-ts at `df5b4af`, docx4j at `901a922a3`), counted the main-part paragraphs holding each item whose character differs. Of 10,957 paragraphs: note references 124 (86 documents), a line `w:br` 72 (71), equations 40 (16), page breaks 23 (17), a run with its own `w:vanish` 22 (10; a style's `w:vanish` not counted), `w:sym` 12 (10), soft hyphens 11 (9), column breaks 8 (8), non-breaking hyphens 8 (8). Every document with more than one paragraph has a different `Body.text`.

**One reader here already has Word's form.** `TrackedChange.text` is docx4j's (`model/revisions/texts.py`): a paragraph mark `\r`, a row's cells a tab apart and then `\r\n`. So today a change across a mark reads `HEAD1B\rBODY1A ` (Word's, check 52 case 1) while `Body.text` reads that join `\n`. The decided change makes the two agree on the mark.

### 2.3 What `get_range` answers today

Measured by a read-only probe on a document built with `insert_xml`: a heading, a centred italic paragraph, a paragraph with a break, a hidden run and a tab, a block control of three paragraphs (check 51's shape), a 2 x 2 table, and a paragraph with an `Emphasis` run.

| call | here | core-ts at `53db5bc` (Word 365 where measured) |
|---|---|---|
| `paragraph.get_range("Whole").text` | `HEAD1A HEAD1B`, the same as `"Content"` | `HEAD1A HEAD1B\r` (check 48); a cell's last paragraph ends `\t` |
| `control.get_range("Whole")`, block control `CCP1`..`CCP3` | `CCP1`, the first paragraph's | `CCP1\rCCP2\rCCP3\r` (check 51) |
| `control.get_range("End")` | offset 4 of `CCP1` | the end of `CCP3` |
| `body.get_range("Whole")` | the first paragraph's range | the first paragraph's start to the last's end, with its mark |
| `body.get_range()` on an empty body | **inserts a paragraph** (0 blocks, then 1) and returns its range | throws (`This body holds no paragraph`) |
| `table.get_range()` | no such member | `'Whole'` ending `\r\n` (check 52 case 1) |
| `comment.get_range()`, a comment from `BODY1B` to `LIST1A` | `[<Range 'BODY1B'>, <Range 'LIST1A '>]` | one `Range`, `BODY1B\rLIST1A` (check 52 case 1) |
| `change.get_range()`, an insertion from `INS` through its mark into `PLAIN1A` | `INS`, in its first paragraph | across both paragraphs (check 53 case 26: ` LIST26B\rPLAIN26A `) |
| `change.get_range()`, a deletion through a mark | collapsed at its start, offset 7 | the same (check 53 case 25, Simple Markup) |
| `range.style` over an `Emphasis` run | `Normal` | `Emphasis` (check 53 case 7) |

### 2.4 The four members that exist by name

- **`Body.get_range`** reads the first paragraph, so every location but `"End"` is partial. It writes on an empty body, where CR-011 section 14.4 has every read leave the parts as they were (Q7).
- **`Comment.get_range`** returns `list[Range]`, CR-003 section 4's rule ("a `Range` per paragraph it touches"). Callers rely on the list: `api_script.py` (`comment.get_range()` filtered by paragraph, line 683), `README.md`'s comment fence (`comment.get_range()[0].text`, line 733, which `tests/test_readme.py` executes), and the tests: six lines index it (`tests/content/test_comments.py` five, `tests/agent/test_comment_workflows.py` one), of 19 test lines that call a comment's `get_range()`. D2 keeps that list as `get_ranges()`.
- **`Range.paragraphs`** is always one paragraph.
- **`Range.style_built_in`** reads the paragraph's value. core-ts's reads check 53's rule, and its subset declares it read-only; this package's setter sets the paragraph's style, as `Range.style`'s does, and may stay (section 4.3).

### 2.5 Found on the way, not fixed here (each is part of a phase below)

1. **`Range.get_comments()` across a comment's paragraphs.** `_ids_in_range` reads only the range's own paragraph's markers. Measured on a comment from `BODY1B` to `LIST1A`: a range on `LIST1A`, inside the comment, lists nothing, because its paragraph holds only the end marker, at 7; a range on `LIST1B`, just after the comment, lists it, since both edges are inclusive. Phase B reads the comment's extent over the story (section 4.6).
2. **`Body.get_range()` writes**, as above (Q7).
3. **Office JS's characters can't be written as text.** U+0002, U+000B, U+000C, U+000E, U+001E and U+001F are not XML 1.0 characters. Measured: `paragraph.text = "c\x0bd"` here keeps the character in a `w:t`, and the save raises `Docx4JException` (`All strings must be XML compatible`). Measured again on 2026-10-23 at `e87d15d` with U+000B and U+0002 through every text writer: `Body.insert_paragraph`, `Paragraph.text = ...`, `Range.insert_text`, `Body.replace_text`, `Range.insert_comment`, `Comment.content = ...`, `TableCell.value = ...`, a content control's `insert_text`, and the facade's `add_paragraph` and `Run.text = ...`. Each keeps the character, none refuses it, and each save raises the same `Docx4JException` from lxml's serialiser, with no `code`: nothing malformed is written, but the error comes at save, far from the write. On core-ts's built `dist/` (its working tree's build, so not a commit; a read-only Node probe), the same set writes `<w:t>c\u000bd</w:t>` raw, which is not well-formed XML, and core-ts reloads it. That is for core-ts's session; here, Q2 decides the writers.
4. **The facade shares a reader.** `docx4j_py/docx/text.py`'s `run_text`, python-docx's `Run.text`, takes `w:t` and `w:sym` from `text_model.item_text`. Phase A must leave `item_text` as it is and give the Office JS readers a map of their own, or the facade's text moves (CR-011 Q5).

## 3. What core-ts built

### 3.1 The decisions (`a85421d`, CR-014 section 5)

All four were decided 2026-10-22 as recommended:
1. **D1, the shape.** One `Range` class with an end paragraph, a single-paragraph range unchanged; not a second class.
2. **D2, `Comment.getRange()`.** One `Range`, as Office JS returns; the per-paragraph array kept as `getRanges()`. It breaks a caller that indexes the array, which 0.x allows.
3. **D3, tables.** A range holding whole tables reads and deletes them. One starting or ending inside a table reads in `body.text`'s form, and its writes refuse unless check 52 showed a simple rule. It showed Word's own refusal (section 3.2).
4. **D4, the editor.** Additive for the editor (its ED-008 section 8.8); phase 1 needn't wait. Nothing of the editor's is in this repository, so D4 has no counterpart here.

### 3.2 Check 52 (`8fdccf8`; `test/README.md` check 52; CR-014 section 6.2)

95 cases, each a region `BEGIN<n> <kind> <op> <tracking> <range>` ... `END<n>`, run once in Word 365 (16.0.20326.20232) in All Markup on 2026-10-22, with Word's saves after the run and after Accept All (`test/check52/`, SHA-1 `15f50530...` and `3d2a3736...`). The kinds: HB, BH, LB, BL, HMB, TBL, INTO, CELL, HID, FN, CC, W1, W2, and TC1 to TC3 for CR-002 section 55. What Word answered:

**Reads:**
- `text` is `body.text`'s form between the ends: a mark `\r`, a table's cells `\t` apart and its rows ended `\r\n`, a note reference `\u0002`, hidden text left out.
- `getRange('Whole')` is the range itself, not widened. `'Content'` is the same without a last mark. `'Start'` and `'End'` are collapsed at the two ends, and `'After'` just after the end, in the next paragraph where the range holds its mark.
- `paragraphs` is every paragraph touched, a table's cells' too. `expandTo` doesn't care which end comes first.
- `font` reads `null` where runs differ: `''` for the name, `Mixed` for the underline.
- A tracked change across paragraphs is one change, and its `getRange()` spans it. A comment's `getRange()` spans its paragraphs (`BODY1B\rLIST1A`).
- `^p` matches a paragraph's mark outside a table, never a cell's or row's end. That is core-ts's phase 3 (section 3.6).

**Writes:**
- **`delete()`** joins the paragraphs in the properties of the paragraph the range **ends** in, each run keeping its own formatting. Blocks between go whole, a table too.
- **A last mark in the range goes with it:** a paragraph's whole range deletes the paragraph, and a range ending at a mark joins its head to the next paragraph, in that paragraph's properties.
- **`Replace`** is `delete()` then the text, keeping a last mark.
  - Untracked, the new text takes the first replaced character's formatting as direct run properties (a heading's font, size, colour and bold, in a body paragraph).
  - Tracked, the `w:ins` goes before the `w:del`, and the new text has no formatting of its own.
- **Inside a table:** a range starting or ending inside one makes Word refuse `delete()` and `Replace` (`GeneralException`, nothing changed). Word does apply `font`, `insertParagraph` and `insertComment` there.
- **`font`** writes every run in the range, a table's cells too, and each mark it holds. `font.bold` adds `w:b` only where a run isn't bold already (no `w:b`, and no tracked change, on a heading's run bold through its style).
- **`insertParagraph`** splits the paragraph at the range's start (`'Before'`) or end (`'After'`), the new paragraph in that paragraph's properties; at a paragraph's start, or a range ending with its mark, nothing is split.
- **`insertText`** at `'Start'` and `'End'` goes at the first and last character, before a last mark.
- **Notes and hidden runs.** A footnote whose reference a deletion takes leaves `footnotes.xml`. A hidden run in the range goes with it; tracked, both stand inside the `w:del`.
- **Under tracking:**
  - a deletion is one `Deleted` change, and the paragraphs stay until accepted;
  - every mark in the range is marked deleted (`w:pPr/w:rPr/w:del`), the last paragraph's mark kept, **no `w:pPrChange`** written;
  - a table in the range has every row and every cell's mark deleted;
  - Accept All gives the untracked answer.
- **Measured, and not copied:** on a control's content range Word's `insertParagraph` adds an empty paragraph as well.

### 3.3 Check 53 (`ef57729`; CR-014 section 9.7)

35 read-only cases, run in Simple Markup and All Markup on 2026-10-22, measuring every choice phase 1 had made where Word was unmeasured. Confirmed: a range across paragraphs of different paragraph styles reads `Normal`, of one style that style; a table's `'Start'`, `'End'` and `'After'`; a cell control's ranges; a range over an inline control lists it. Changed to Word's answer:
- **The style rule.** A range holding its paragraphs whole reads their paragraph style, shared, else `Normal`. Any other range reads its runs' character style where they share one (`Emphasis`), and `''` (`Other`) where their character styles differ (a footnote reference among plain text); with none, the paragraph style. This holds within one paragraph too.
- **`MixedFont`:** a colour or highlight where runs differ reads `''`.
- **Controls:**
  - a row control's text is its cells a `\t` apart;
  - a row control's, or a table-ending control's, `'Content'` has no row end, and `'Whole'` ends `\t\t`;
  - a range lists the controls it overlaps, not one holding it.
- **Tracked changes:**
  - a deletion's range is collapsed (empty in Simple Markup);
  - an inserted mark is `\r` in its own paragraph;
  - a lone deleted mark is collapsed at its paragraph's end;
  - an inserted paragraph reads `...\r` in that paragraph alone.
- **`expandTo` across stories** throws `ItemNotFound`.

### 3.4 Phase 1, the reads (`ef57729`; CR-014 section 9)

- **One `Range` with two ends:** `paragraph` and `start`, and `endParagraph` holding `end`, which is `paragraph` for a range within one paragraph. Flags: `whole` (a paragraph's whole content, pictures at its edges included, CR-002 section 48), `mark` (a `'Whole'` range: the text ends with the last paragraph's mark), and `endMark` (what replaces that mark: a table's `\r\n`).
- **`story.mts`:**
  - `storyOf(p)` is the outermost block list above a paragraph; a text box's content and an unread branch of a kept `mc:AlternateContent` are stories of their own.
  - `storyLayout(items, textOf)` is that story's text in `body.text`'s form, with each paragraph's place in it. `Body.text` became the layout's text, its last mark dropped, unchanged in every file of its sweep.
- **The reads:**
  - `text` is the layout's slice, plus the mark for a `'Whole'` range; `paragraphs`; `isEmpty` is `text === ''`;
  - `style` and `styleBuiltIn` by check 53's rule; `font` across paragraphs a `MixedFont`, a single paragraph's still its first run's;
  - `search` a paragraph at a time;
  - `getTrackedChanges()` the story's listing filtered by overlap, since a paragraph's own listing cuts a change at its edge; `getComments()` each paragraph's, deduplicated;
  - `getRange` with `'After'`; `expandTo`, refusing two stories with `ItemNotFound`;
  - `tables`, `contentControls`, and `parentTableCell` and `parentTable` (extensions, the start's); `hyperlink` the first link among its paragraphs.
- **Every write refuses** with `NotSupportedError` until phase 2, `getOoxml` and `toApiScript` too.
- **`Paragraph.tableNestingLevel`**, counted through `parentTableCell`.
- **The views:**
  - `ContentControl.getRange()` of a block, row or cell control is one range, `SpanningRange` retired;
  - `Comment.getRange()` is one range (`undefined` with no markers in the body), the old array `getRanges()`;
  - `TrackedChange.getRange()` of a group runs from its first piece to its last, by check 53's rules;
  - `Body.getRange()`; and `Table.getRange()`.
- **Held to Word** by running check 52's own snippet through core-ts's Word shim (`reads` mode on the file, `report` mode on Word's save after Accept All), and check 53's (section 6.1 lists the differences asserted as they stand). A sweep of 806 `.docx` against `96f735b`: `Body.text`, each paragraph's text and whole range, each cell's text and `getRanges()` unchanged; 24 control ranges, 59 comment ranges and 472 tracked-change ranges moved as intended.

### 3.5 Phase 2, the writes (`53db5bc`; CR-014 section 10)

Every write across paragraphs, and a delete or tracked `Replace` of a range holding its mark, goes through `spanEdits.mts` (485 lines):
- **`extentOf`:** the innermost block list holding both ends, and the block holding each end. A range starting or ending inside a block it doesn't hold whole is refused before anything changes, with `GeneralException` inside a table (Word's own refusal) and `NotSupportedError` inside a content control (unmeasured).
- **`delete()`, untracked:**
  - the head's tail and the tail's head go, and the blocks between go whole; the two paragraphs join through the resolution's `joinWithNext`, in the last paragraph's properties, keeping its `w14:paraId`;
  - a held last mark goes, joining the paragraph after, unless none follows in the same list (a cell's last paragraph, a control's, the body's, one before a table), where the mark stays (CR-002 section 55's rule: a cell keeps its last paragraph);
  - a note whose reference goes is queued to leave on save; a section break in the first mark goes with it, its header and footer relationships released.
- **`delete()`, tracked:**
  - nothing is joined: the text is a `w:del`, every held mark marked deleted, no `w:pPrChange`, a table's rows `w:trPr/w:del`;
  - a mark another author already deleted is left as it is.
- **`Replace`:**
  - untracked, in `lookOf`'s first-character look: the run's own properties, and each style property the joined paragraph's style doesn't give, written direct; neutral values such as `w:b w:val="0"` where the source had none;
  - tracked, `w:ins` first.
- **`font`:**
  - every run and held mark, the toggles written only where the run or mark doesn't read them already;
  - tracked, a mark's `w:pPr/w:rPr/w:rPrChange` is new (`recordMarkRPrChange`).
- **`insertText`** at the ends.
- **`insertParagraph`:**
  - splits a paragraph (`splitParagraph`): a link, a revision or a smart tag the split falls in is cut in two, a revision's second half under a new id;
  - a control, a field, an equation or a kept `mc:AlternateContent` there is refused;
  - tracked, the first half's mark and the new paragraph are inserted.
- **`replaceText`** a paragraph at a time.
- **Changed for one paragraph:** a range holding its paragraph's mark deletes the paragraph, joining the next (check 52 case 14); a tracked `Replace` of such a range writes `w:ins` first.
- **Refused under tracking** (`assertTrackable`): a range holding a whole block that is not a paragraph or a table (a content control, custom XML, an `w:altChunk`, a kept `mc:AlternateContent`). Untracked such a block goes whole.
- **Held to Word:**
  - check 52's snippet in `run` mode through the shim, against Word's run, case by case;
  - Accept All after it, against Word's report after Accept All;
  - each region's markup, a line per paragraph, against Word's save after the run.
  - The differences are listed (section 6.1).
- **The sweep:** 808 `.docx`, six writes each on a range from mid-paragraph to mid-paragraph one to four later, untracked and tracked. Every write not refused agreed both ways: Accept All of the tracked write equal to the untracked one accepted, and Reject All equal to the original.
- **Not measured** (CR-014 section 10.6): ten choices, among them a single-paragraph `insertParagraph` still going before or after the paragraph, the toggles other than bold, `lookOf`'s other neutral values, and a tracked `Replace` within one paragraph not holding its mark (CR-002's `w:del` first).

### 3.6 Left by core-ts at `53db5bc`

CR-014 phase 3, comments and content controls inserted on a range across paragraphs and `^p` in `search` (its sections 4 and 10.7), is not committed. core-ts's working tree at the time of writing holds uncommitted edits to `Range.mts`, `spanEdits.mts`, `search.mts` and the tests, and an untracked `test/check54/`. They are not cited here. Also left there: `Range.footnotes` and `Paragraph.getTextRanges`; CR-002 section 55 (Word's table rule, decided, not implemented at `53db5bc`); section 58 (a paragraph's listing as its parts of each change, proposed `d3b3170`); section 59 (`Font` where runs differ within a paragraph, an automatic colour, proposed `d3b3170`).

### 3.7 core-ts CR-002 sections 55 and 57 (at `53db5bc`)

- **Section 55, Word's table rule:** a cell's last block, as Word reads it, is a `w:p`, and Word's own edits never leave a cell without its last paragraph. Check 52's cases 86 to 95 measured the edits: `Paragraph.delete()` or `body.clear()` of a cell's last paragraph empties it and keeps its `w:pPr`; tracked, only its text is deleted. Decided (D1 to D4 as recommended), not implemented. Phase C here takes from it only what core-ts's phase 2 does: a held last mark with no paragraph after it in the same list stays. The section's own port is separate (section 4.10).
- **Section 57, a departure:** Office JS reads a paragraph after a nested table in a cell, unless Normal, as style `Other`. core-ts reads its own `w:pStyle`, so that a heading's or a list item's style is not hidden. Recorded at the user's word ("record difference 1 as an intentional departure somewhere"). Q6 takes it here too.

### 3.8 The characters (core-ts CR-012 phase 4, `9307bdb`, its sections 17.1, 17.2 and 17.7)

`Paragraph.text` became its segments' text in Office JS's characters, one reader for text, offsets, `Range`, `search` and every edit:

| item | Office JS (check) | here today |
|---|---|---|
| paragraph mark in `Body.text`, multi-paragraph text | `\r` (checks 30, 31, 51) | `\n` |
| table row, cells | cells `\t` apart, the row ended `\r\n`, a cell's last mark not written (checks 30, 31, 51) | a line per cell paragraph |
| a paragraph's `'Whole'` range | its text and `\r`; `\t` for a cell's last (check 48) | its text |
| a cell's `value`, a block control's `text`, a comment's `content` | their content's `body.text`, `\r` between paragraphs (check 51) | `\n` between paragraphs |
| `w:br` line, page, column | U+000B, U+000C, U+000E (check 48) | `\n` |
| `w:cr` | U+000D | `\n` |
| footnote or endnote reference; a note's own mark | U+0002 (checks 48, 51) | nothing |
| `w:ptab` | `0` | nothing |
| non-breaking, optional hyphen | U+001E, U+001F | U+2011, U+00AD |
| any `w:sym` | `(` (check 48b) | the character `w:char` names |
| hidden text (own `w:vanish`, else the styles') | left out, in every view (checks 48, 48b) | read |
| an equation | one segment, Word's text of it (`mathText.mts`, check 51) | nothing |
| a comment reference, a field code, a picture | nothing | nothing |

`search` (`officeMatches`, `search.mts`) became Office JS's: a picture is U+FFFC, a place a plain query doesn't match across and `^g` matches; hidden text is not there; an optional hyphen is passed over unless named (`^-`); Word's Find codes `^f ^e ^l ^m ^t ^~ ^- ^g ^^`; a plain `-` matches a non-breaking hyphen. Kept off it (its section 17.7): `TrackedChange.text`, which keeps docx4j's characters as the revisions goldens hold them. Hidden text in an edit is an object to the span, as a picture is; a span may take an equation whole untracked, and cutting one, inserting inside one or any tracked edit taking one is refused.

## 4. Design here

Follow core-ts, member for member and rule for rule, except where Python convention or a rule decided here says otherwise; section 4.9 lists each departure.

### 4.1 Office JS's characters (Phase A)

**The decided rule** (the user, 2026-10-23). It covers four things:
- a paragraph mark is `\r` in `Body.text` and in every multi-paragraph text;
- a table in `Body.text` is Word's form: a row's cells `\t` apart, a cell's last mark not written, each row ended `\r\n`, a row-level control's rows the same;
- `Body.text`'s last paragraph's mark is not written;
- a `w:br` (a line break) is U+000B.

"Multi-paragraph text" reads on a cell's `value` and `Table.values`, a block control's `text` and `Comment.content`, all of which are their content's `Body.text` in core-ts, and on a range across paragraphs. Q1 asks the user to confirm that reading, with the rest of the set.

**The rest of phase 4's set** is Q1: U+0002 for a note reference and a note's mark, U+000C and U+000E for a page and a column break, U+000D for `w:cr`, `0` for `w:ptab`, U+001E and U+001F for the hyphens, `(` for `w:sym`, hidden text left out by the text model's rule, an equation as Word's text of it, a paragraph's `'Whole'` range ending with its mark, and `search`'s Find codes, U+FFFC and the optional hyphen. Recommended: all together (Q1).

**Where it lives.**
- A map of Office JS's characters in `text_model.py` beside `_CONSTANT_TEXT`, used by `segments_of`.
- `item_text` keeps its current answers for `traversal.text_of` and the facade's `run_text` (section 2.5 item 4).
- Hidden text through the text model's rule: the extractor's REQ-507 test at `model/text/extractor.py` line 1232, a run's own `w:vanish`, else its styles'. It is cached per style pair, as core-ts's `PropertyResolver.isVanishedByStyles` is.
- Word's text of an equation is a port of core-ts's `mathText.mts` (209 lines at `9307bdb`) into `model/content/math_text.py`, held to check 51. It is not `model/math/` (OMML to LaTeX), which is the text model's.

**The readers that move:**
- `Paragraph.text`, `get_text` in `accepted` and `original`, `text_budget`;
- `Body.text`, `get_text` and `text_budget`, through a story layout (section 4.2);
- `Range.text`; `TableCell.text` and `value`, and so `TableRow.values` and `Table.values` (from `text_of` to the cell's `Body.text`); a block control's `text` (already its `Body`'s) and a run control's (from `text_of` to its segments); `Comment.content` (already its `Body`'s);
- `search`, `find`, `replace_text`'s matches, `SearchHit`'s context, `paragraph_at(contains=)`;
- `Paragraph.to_dict()` and `repr`, and a `ChangeReport`'s `text_before` and `text_after`;
- the text model's `office_offset` (below);
- the offsets' own walkers, which count where the characters' lengths change (U+0002, `0`, hidden text, an equation): the comment markers' (`comments._visit_runs`) and a revision's collapsed place (`tracked_change._accepted_offset_of`). They move onto `segments_of`'s offsets, so that one reader counts.

**The readers that don't move,** each documented as not Office JS's:
- `traversal.text_of`, docx4j's `TextUtils`, as core-ts's `textOf` stays (CR-011 Q4);
- `Block.text` and `Table.text`, this package's extensions over `text_of`, a line per paragraph, as core-ts's `Table.text`;
- the python-docx facade's readers (CR-011 Q5): `run_text` and `holder_text`;
- `TrackedChange.text`, docx4j's characters (CR-011 Q6; core-ts CR-012 section 17.7), which already has the `\r` mark and Word's table form and keeps `\n` for a break;
- the text model, its plain text and Markdown, `outline()` and `describe()`.

**Offsets and `office_offset`.**
- Under the two decided characters no offset in a paragraph moves: `\n` and U+000B are one code point each, and a `Range`'s offsets are a paragraph's, not `Body.text`'s.
- Under Q1's full set, offsets move in a paragraph holding a note reference (+1 each), a `w:ptab` (+1), hidden text (minus its length) or an equation (plus its text's length). Over the corpus that is at least 124, 22 and 40 paragraphs (section 2.2).
- A text model run keeps its `office_offset` where the item's length agrees (`extractor.py`, `para.office_span`), as core-ts's `Extractor.mts` assigns it.
- Phase F's audit holds "agreement with `get_text()` character for character", and a break (`\n` in the model, U+000B in Office JS's text) would break it. Q11 restates it as agreement under the character map, every run with an offset comparing equal once Office JS's character is read as the model's.

**The writers** (Q2). A text read from a paragraph and written back must save, and six of the characters are not XML.
- Recommended, so that a write reads back as the item it read: U+000B a `w:br`; U+000C a page break; U+000E a column break; U+000D a `w:cr`; U+001E a `w:noBreakHyphen`; U+001F a `w:softHyphen`; `\t` a `w:tab` and `\n` a `w:br`, as now.
- U+0002 is refused: no text makes a note.
- `(` and `0` are ordinary text, so writing a `w:sym`'s or a `w:ptab`'s read text back is lossy, as it is in Word.
- It applies in every text writer, the ones section 2.5 item 3 measured: `set_text`, `splice`, `insert_text` (a content control's included), `replace_text`, `insert_paragraph`, `insert_comment`, `Paragraph.text = ...`, `TableCell.value = ...` and `Comment.content = ...`. In a comment both `\r` and `\n` are paragraph marks, so that `Comment.content` written back makes the paragraphs it read (core-ts's D5). Today the writer (`_comment_paragraphs`, `comments.py` line 545) splits on `\n` only, and a `\r` stays in a `w:t`, saved as `&#13;` (measured 2026-10-23: `"one\rtwo\nthree"` saves two paragraphs, `one&#13;two` and `three`). Phase A splits on both.
- The other characters XML 1.0 excludes (U+0000 to U+0008, U+000F to U+001D, U+FFFE, U+FFFF, a lone surrogate) are refused at the write with U+0002, `ContentError` `text.not_writable`, before anything changes. Today they fail at save (section 2.5 item 3), so only the place of the error moves. core-ts proposes the same set, under Office JS's `InvalidArgument` (its CR-002 section 60, D2, proposed at `87bbb38`).
- lxml's refusal at save stays as the guard behind them (core-ts's D3).
- **The python-docx facade's writers** (`add_paragraph`, `add_run`, `Run.text = ...`) keep python-docx 1.2.0's behaviour, measured in Phase A, since the facade is python-docx's names and semantics (CR-003 Phase J).
- **A bound control's write-back keeps docx4j's characters.** `update_from_control` (`model/customxml/bindings.py`) writes `control.text` into the control's custom XML node, and Phase A moves a control's `text` into Office JS's characters. Unchanged, a control holding a `w:br` would write U+000B into the node, which lxml refuses, and a control of several paragraphs would join them with `\r`. The write-back reads `text_of`'s characters instead, as now: a break `\n`, the hyphens U+2011 and U+00AD. That is docx4j's text, as core-ts's D4 has it. core-ts found the same regression in its unreleased CR-012 phase 4 at `87bbb38`. Phase A's tests pin a bound control with a break, and one of two paragraphs, written back as they are today.

**What it breaks or moves for callers** (each goes into the CHANGELOG's 0.2.0 entry, not yet released):
- **`Body.text`**: "a paragraph per line" becomes `\r` per paragraph with Word's table form. `README.md` (line 41) and `docs/pypi-readme.md` (line 27) say "a paragraph per line"; both fences are executed by `tests/test_readme.py`. A caller's `body.text.split("\n")` gets one string.
- **`str.splitlines()`** splits at `\r`, `\r\n`, U+000B, U+000C and U+001E, and not at U+000E, U+001F or U+0002 (checked in Python 3). Under Q1's full set, `splitlines()` cuts a word at a non-breaking hyphen. The docstrings and the readme say to split paragraphs with `split("\r")`.
- **A break in a paragraph** reads U+000B: `paragraph.text == "a\nb"` is false where it was true. **The CHANGELOG's Phase F entry** says "A break stays `\n` where Office JS gives U+000B" (`CHANGELOG.md` line 191), and Phase A rewrites that line.
- **A cell's value and `Table.values`**, a block control's text and a comment's content join with `\r`.
- **Stored offsets** move only under Q1's full set (above); addresses (paraId or ordinal) don't.
- **`paragraph_at(contains="a\nb")`** no longer finds a paragraph whose break was `\n`.
- **The python-docx facade** doesn't move (Q5), provided `item_text` is left as it is.
- **`TrackedChange.text`** doesn't move. So a change's `text` and its `get_range().text` differ for a break (`\n` against U+000B) and, under Q1's full set, for a note reference and hidden text, as in core-ts.
- **`text_of`** doesn't move, so `Table.text` (a line per paragraph) and `Table.values` (Office JS's) part company, as in core-ts.
- **Under Q1's full set:** hidden text leaves `Paragraph.text` and `search` (a hidden word is not found), so `paragraph.text = paragraph.text` drops the hidden runs, as core-ts's does; a search for a literal `^` changes meaning (`^^` is a caret); an edit that cuts an equation is refused.

### 4.2 One `Range`, two ends (D1)

```python
class Range:                     # Word.Range: [start, end) from one paragraph to the same or a later one in its story
    __slots__ = ("end", "end_mark", "mark", "paragraph", "start", "whole", "_end_paragraph")
    def __init__(self, paragraph, start, end, *, end_paragraph=None, mark=False, end_mark=None, whole=False): ...
    end_paragraph: Paragraph     # the paragraph holding `end`; `paragraph` for a range within one
    spans_paragraphs: bool       # extension (core-ts's spansParagraphs)
    holds_end_mark: bool         # extension: a 'Whole' range ending at its last paragraph's end, or a table's end
```

- A range within one paragraph is exactly today's: the constructor's three positional arguments are unchanged, and every range this package makes now is one.
- A range across paragraphs is made by `expand_to` or a view. A caller's `Range(p1, s, e, end_paragraph=p2)` is checked: both paragraphs in one story, `p2` not before `p1`. Otherwise `SpanError` `range.other_story` or `range.order`.
- `whole` and `mark` are core-ts's flags. Under Q1's full set `mark` adds the `\r` to `text`; under the two decided characters alone it is structural only (section 4.9 item 4).
- `__eq__` and `__hash__` take both paragraphs, both offsets and `mark`. `__len__` is `len(self.text)` across paragraphs and `end - start` within one, as now.
- `repr` names both ends: `<Range 4:body/5:6 'HEAD3B\rBODY3A'>`. `to_dict()` adds `end_address` and `mark` only for a range across paragraphs, so a single paragraph's JSON is unchanged.
- **The story** (`model/content/story.py`, new, after `story.mts`). `story_of(p)` climbs the parent pointers to the outermost block list: the body's, a header's, a footer's, a note's, a comment's, a text box's content. A kept `mc:AlternateContent`'s unread branch is a story of its own; Word's branch is part of the story around it, as `Body.paragraphs` already reads it (CR-012 Phase A). `story_layout(items, text_of)` gives the text in `Body.text`'s form and each paragraph's place, through `block_children_of`, `rows_of` and `cells_of` (`text_model.py`). `Body.text` becomes the layout's text with its last mark dropped (Phase A), so the two can't disagree.

### 4.3 The reads (Phase B), member by member

| member | here | rule |
|---|---|---|
| `Range.text` | property | the layout's slice between the ends; with `mark`, the last paragraph's mark (`\r`, `\t` for a cell's last) or `end_mark` (`\r\n`) |
| `Range.paragraphs` | property, `list[Paragraph]` | every paragraph between the ends in the layout, a table's cells' among them |
| `Range.is_empty` | property | `text == ""` |
| `Range.style`, `style_built_in` | properties; setters kept for one paragraph, refused across | check 53's rule, within one paragraph too; `Other` where character styles differ; CR-002 section 57's departure (Q6) |
| `Range.font` | property | within one paragraph the first run's, as now; across paragraphs `None` where runs differ, `""` for name, colour and highlight, `"Mixed"` for underline (Q4) |
| `Range.search` | method | a paragraph at a time, from `start` in the first to `end` in the last |
| `Range.get_tracked_changes()` | method | across paragraphs, the story's listing filtered by overlap; within one, the paragraph's listing, its overlap test taught a change that starts or ends elsewhere |
| `Range.get_comments()` | method | the comments whose extent over the story overlaps the range (fixes section 2.5 item 1) |
| `Range.get_range(location)` | method; `RangeLocation` gains `"After"` | `"Whole"` the range itself; `"Content"` without a last mark; `"Start"`, `"End"` collapsed; `"After"` collapsed after the end, at the next paragraph's start where the range holds its mark |
| `Range.expand_to(other)` | method | from the earlier start to the later end, a mark-holding end counting as later; `SpanError` `range.other_story` across stories (Office JS's `ItemNotFound`) |
| `Range.tables` | property, `list[Table]` | check 52's rule: a table the range covers some of; inside one cell, that cell's tables |
| `Range.content_controls` | property, `list[ContentControl]` | the controls the range overlaps, not one holding the whole range, in document order (check 53) |
| `Range.parent_table_cell`, `parent_table` | properties, extensions | the start's cell and table, or None (core-ts's) |
| `Range.hyperlink` | read | the first link among its paragraphs; the setter refused across paragraphs |
| `Range.runs` | property, extension | every run of every paragraph between the ends |
| `Range.get_text(view=)` | method | as now: `accepted` only |
| `Paragraph.table_nesting_level` | property, `int` | 0 outside a table, 1 in a cell, 2 in a nested table's, through `parent_table_cell` |
| `Body.get_range(location)` | method | `"Whole"` from the first paragraph's start to the last's end with its mark, `"Content"` without, `"Start"`, `"End"`; a body of one paragraph that paragraph's; no paragraph: `ContentError` `body.no_paragraph` (Q7) |
| `Table.get_range(location)` | method, `Literal["Whole", "Start", "End", "After"]` | `"Whole"` from the first cell's start to the last cell's end, `end_mark` `\r\n`; `"Start"`, `"End"`; `"After"` at the next paragraph |
| `Comment.get_range()` | method, a `Range` or None | one range from the start marker to the end marker over the story (D2); a start marker between blocks starts at the next paragraph, an end marker between blocks ends at the previous paragraph's end; None with no markers in the body |
| `Comment.get_ranges()` | method, `list[Range]`, extension | today's `get_range()`, unchanged; `api_script.py` and `README.md` move to it |
| `ContentControl.get_range(location)` | method | a block, row or cell control: one range from its first paragraph's start to its last's end, `"Whole"` with the mark; a row control or a table-ending control: `"Content"` ending with the last cell's text, `"Whole"` with `\t\t` (check 53); `"End"` the last paragraph's end |
| `TrackedChange.get_range()` | method | a group from its first piece to its last across paragraphs; a mark ends its own paragraph's range, holding it; a deletion collapsed where it starts; an inserted lone mark its mark, a deleted lone mark collapsed at its paragraph's end (check 53) |

- **Every write across paragraphs refuses** in Phase B, before anything changes, with `SpanError` `range.spans_paragraphs`, naming Phase C. That covers `text`, `insert_text`, `insert_paragraph`, `insert_ooxml`, `delete`, `replace_text`, `insert_comment`, `insert_content_control`, `hyperlink`, `style`, `style_id`, `style_built_in` and `font` writes, and `get_xml` and `to_api_script`, whose one `w:p` can't hold several blocks.
- **The subset** (`scripts/office_js_subset.py`) is re-derived at Phase B's end: 241 members, every non-extension one implemented but the two null-object members omitted as before (CR-003 section 18.2 item 12). `tests/content/test_office_js_subset.py`'s phase tables need no change: the new members fall to phases B, C and G, all implemented.

### 4.4 The writes (Phase C)

`docx4j_py/model/content/span_edits.py` (new) is the port of `spanEdits.mts` at `53db5bc`. `Range` calls it for every write across paragraphs, and for a delete or a tracked `Replace` of a range holding its mark.
- **`extent_of`**, as core-ts's. A range starting or ending inside a table is refused as Word refuses it: `SpanError` `range.partial_table` (Q5), for `delete` and `Replace` only. One starting or ending inside a content control is refused with `range.partial_control`.
- **`delete`**:
  - untracked, the head's tail, the blocks between and the tail's head go, and the paragraphs join through the resolution's `join_with_next` (`model/revisions/resolver.py` line 455), as Accept All joins, in the last paragraph's properties and `w14:paraId`;
  - a held last mark goes with the paragraph after joined in, unless none follows in the same list;
  - a note whose reference goes leaves on save, unless a reference stands then;
  - a section break in the first mark goes with its relationships released.
- **`delete`, tracked:**
  - the text is a `w:del`, hidden runs and note references inside it;
  - every held mark is marked deleted with no `w:pPrChange`; a table's rows and their cells' marks are deleted;
  - a mark already deleted is skipped. `tracking.mark_paragraph_deleted` raises `tracking.already_deleted` today, and span edits take a skipping form of it, as core-ts's `markDeletedInPlace` fix did.
- **`insert_text(..., location="Replace")`**: `delete` keeping a last mark, then the text where the paragraphs joined, untracked in `look_of`'s look, tracked with the `w:ins` first.
- **`font` writes** reach every run and held mark, toggles only where not read already. Tracked, a mark's `w:pPr/w:rPr/w:rPrChange` is new in `tracking.py`, recorded and settled as a run's is. This package doesn't settle a run's either (CR-003 section 21.3 item 9).
- **`insert_text` at the ends** goes at the first and last character, before a last mark.
- **`insert_paragraph`** splits a paragraph (`split_paragraph`), cutting a link, a revision or a smart tag in two, and refuses a control, a field, an equation or a kept element there. A single paragraph's `insert_paragraph` is unchanged, before or after the paragraph (core-ts section 10.6 item 1).
- **`replace_text`** goes a paragraph at a time, every match checked before the first is replaced (Phase F's rule; core-ts refuses paragraph by paragraph).
- **Changed for one paragraph:** `paragraph.get_range("Whole").delete()` deletes the paragraph and joins the next where one follows in the list, where today it empties it (measured, section 2.3); a tracked `Replace` of a range holding its mark writes `w:ins` first.
- **Refused under tracking** (`assert_trackable`): a range holding a whole block that is not a paragraph or a table. Untracked such a block goes whole.
- **The range after a write:** collapsed where a deletion closed up (tracked, where it starts); over the new text after `Replace`; over its text after `insert_paragraph` and `insert_text`.
- **The agent surface.** Each write records one `ChangeReport` (CR-003 section 3.4): `addresses` are every paragraph touched and every block removed; `moved` is the ordinals shifted by each removal (`reports.moved_by_delete`); `text_before` and `text_after` are the range's text; `parts_touched` gains `footnotes.xml` or `endnotes.xml` where a note goes. A write in a `dry_run` is undone with the trial (CR-003 section 12.5).

### 4.5 Kept `mc:AlternateContent` branches

Phase F's rule and CR-012 Phase A's hold for every paragraph a range touches; core-ts's ED-009 editing in place (its CR-012 section 12.2) is not ported (CR-011 section 15 item 3).
- **Reads:**
  - a range may run through text read from a run- or paragraph-level branch, and through the paragraphs of Word's branch of a block-, row- or cell-level element, which are part of the story as `Body.paragraphs` reads them (CR-012 Phase A, D2's addresses);
  - an unread branch is a story of its own, so `expand_to` into one refuses (`range.other_story`).
- **Writes, each refused before anything changes:**
  - a range with an end inside a kept block-, row- or cell-level element's branch refuses every write with `address.read_only`, since its paragraphs' views are read-only (`Body.body_for`, `_BranchBody`);
  - the head's and the tail's pieces obey `touches_branch_text`, so a span holding a branch's text in part or whole refuses with `address.read_only` (Phase F 16.3);
  - a `font` write is refused where any paragraph's piece holds branch text (`Font`'s `guard`), and a split inside a kept element is refused.
- **What goes whole.** A paragraph between the ends goes whole untracked, any run- or paragraph-level element in it with it, as an untracked `Paragraph.delete()` takes it (Phase F 16.3). A kept block-level element the range holds whole goes whole untracked, as an untracked `Body.clear()` takes it.
- **Under tracking,** either refuses through `kept_content_under`, as a tracked `Paragraph.delete()` or `Body.clear()` does, and through `assert_trackable`.
- **A difference from core-ts:** core-ts edits Word's branch in place, so a delete across paragraphs through a run-level branch's text succeeds there and is refused here. It is the departure CR-011 Phase F already records, and the tests assert it.

### 4.6 Tables, comments, content controls and tracked changes

- **Tables** (D3, check 52): a range across a whole table reads it in Word's form and deletes it whole. One starting or ending inside a table reads, and `delete` and `Replace` refuse with `range.partial_table`; `font` and `insert_paragraph` apply. A held last mark before a table, or a cell's last paragraph's, stays (CR-002 section 55's rule as core-ts's phase 2 has it). `Paragraph.table_nesting_level`, `Range.tables` and `Table.get_range` are section 4.3's. CR-002 section 55's own repair is not in this CR.
- **Comments:**
  - `Comment.get_range()` and `get_ranges()` (D2);
  - `Range.get_comments()` over the comment's extent in the story (section 2.5 item 1);
  - `insert_comment` across paragraphs stays refused (`range.spans_paragraphs`) until core-ts's phase 3 is ported (Q10), so the facade's `add_comment` keeps refusing runs in two paragraphs (`comment.spans_paragraphs`, `docx/comments.py`);
  - a deletion across paragraphs leaves a comment's markers where they fall (core-ts section 10.6 item 8), and Accept All's comment rule (CR-010) resolves them.
- **Content controls:**
  - `ContentControl.get_range()` and `Range.content_controls` (section 4.3);
  - `insert_content_control` across paragraphs stays refused (core-ts's phase 3);
  - a range ending inside a control refuses `delete` and `Replace` (`range.partial_control`); one holding a control whole takes it untracked and refuses tracked.
- **Tracked changes:**
  - `TrackedChange.get_range()` and `Range.get_tracked_changes()` (section 4.3);
  - a paragraph's own listing stays as it is, cutting a change at its edge, as core-ts's does until its CR-002 section 58;
  - the listing's grouping where Word's differs (after a tracked `Replace` Office JS lists the insertion alone; formatting across a table's cells is one change in Word) is CR-010's and is asserted as it stands, as core-ts asserts it.

### 4.7 Errors

Section 3.1's one hierarchy. Every refusal comes before anything changes, and each code is stable for a tool result:
- `SpanError`:
  - `range.spans_paragraphs`: a write Phase C doesn't bring, such as a comment, a control, `insert_ooxml`, a style or a link;
  - `range.partial_table`: Word's own refusal (check 52's INTO and CELL);
  - `range.partial_control`;
  - `range.other_story`: `expand_to` across stories, Office JS's `ItemNotFound`;
  - `range.order`;
  - `range.not_trackable`: `assert_trackable`;
- `AddressError` `address.read_only`: a kept branch (section 4.5);
- `ContentError` `body.no_paragraph` (Q7) and `text.not_writable` (U+0002, Q2).

The hints name the way on, as CR-003 section 3.1 requires: "write each paragraph's range", "turn change tracking off", or "delete the table, then the paragraphs".

### 4.8 What Python convention changes

- Lists, not collections; `None` for Office JS's `undefined` and `null`.
- `Literal` locations, each with a `StrEnum` (`RangeLocation` gains `"After"`; new `TableRangeLocation` and `BodyRangeLocation`).
- `Font` getters typed `bool | None` and `float | None`, since a value across paragraphs may be mixed (Q4).
- Keyword-only constructor arguments for the new ends and flags.
- `expand_to` returns a new `Range` and leaves both operands as they were, as Office JS's does.

### 4.9 Departures from core-ts, each with why

1. **Errors are `SpanError` codes**, not `GeneralException` or `NotSupportedError` (Q5). CR-003 section 21.3 item 3 is the precedent; the codes carry Word's meaning.
2. **No editing in place through a kept branch** (section 4.5): CR-011 Phase F's rule, which core-ts's ED-009 doesn't have.
3. **`replace_text` checks every match first,** where core-ts refuses paragraph by paragraph. This is Phase F's rule.
4. **Under Q1's narrower answer only**, a `"Whole"` range would carry `mark` without adding `\r` to its text, and `is_empty` would need a structural definition. Under the recommended answer this departure does not arise.
5. **The writers map Office JS's characters to their elements** (Q2), where core-ts writes them raw into `w:t` (section 2.5 item 3).
6. **The oracles run as a Python port of the snippets**, not through an Office JS shim. This package has none, and the snippets' report shape is kept so that Word's JSON compares field for field (section 6.1).

### 4.10 Not in this CR

- core-ts's CR-014 phase 3, `^p`, and comments and controls inserted on such a range (Q10);
- ranges across stories or a text box's edge (core-ts's scope too);
- `Range.footnotes` and `Paragraph.get_text_ranges`, which neither subset declares;
- core-ts CR-002 section 55's repair, section 58's paragraph listing and section 59's `Font`;
- the facade's `add_comment` across paragraphs;
- the HTML reader (CR-011 Phase E).

## 5. Phases

| phase | what | size |
|---|---|---|
| A | Office JS's characters: the two decided, and Q1's set if taken; the story layout behind `Body.text`; the writers' map (Q2); the readers that move off `text_of`; `office_offset` (Q11); checks 48, 48b and 51 as oracles; the readers' audit before and after; README, pypi-readme and CHANGELOG | 2 days (the two decided), 4 days (Q1's full set, with `math_text.py` and `search`'s Find codes) |
| B | The reads: `Range` with two ends and the story; the ten members and the views of section 4.3; Comment's D2; every write across paragraphs refused; the subset re-derived (241); checks 52 and 53 as oracles (fixtures, the Python port of their snippets' `reads` and `report` modes); the sweep | 2½ to 3 days |
| C | The writes: `span_edits.py`, the single-paragraph changes, the refusals, `ChangeReport` and `dry_run`; check 52's `run` mode, Word's markup and Word's Accept All as oracles; Reject All's round trip; the write sweep; acceptance regenerated with an artefact of ranges across paragraphs | 3½ to 4 days |
| D | The Word check of Phase C's artefact (and of Phase A's writes, if any output changed); `tests/README.md`'s records; CR-003 section 23's review point (what it cost, what it gave); the README row | ½ day |

A comes first because every text B and C compare with Word's is in Office JS's characters. B before C, as in core-ts. Each phase is committed and reported on its own; A is a release-visible change of its own and has its own CHANGELOG entry.

## 6. Tests and the gate

### 6.1 The oracles: core-ts's checks, ported from their commits

**The fixtures** (Q8): the bytes Word 365 read, and Word's own saves, copied byte for byte with a SHA-1 manifest into `tests/office_js/`. The directory is outside `tests/fixtures/`, as core-ts keeps `test/check52/` outside its `fixtures/`, so that no sweep of `tests/fixtures/` moves (`test_hyperlinks.py` reads every `.docx` there, and check 48's case 14 is a link). Their source:

| check | file Word read (share `__tmp/`) | SHA-1 | Word's answers (core-ts commit) |
|---|---|---|---|
| 48 | `48/48-text.docx` | `8dec79fd...` | `test/fixtures/check48/48-text-word365.json` (`30b559e`), `48-text-views-word365.txt` (`5c6cb5c`) |
| 48b | `48/48b-text.docx` | `ca98e9c0...` | `48b-text-word365.json` (`5c6cb5c`) |
| 51 | `51/51-text.docx` | `7a6a0102...` | `test/fixtures/check51/51-text-word365.json` (`9307bdb`) |
| 52 | `52/52-ranges.docx` | `e014014d...` | `test/fixtures/check52/52-ranges-word365-run.json` and `-accepted.json` (`8fdccf8`); Word's saves `test/check52/52-ranges-word365.docx` (`15f50530...`) and `-accepted.docx` (`3d2a3736...`), the share's copies the same bytes |
| 53 | `53/53-reads.docx` | `aea48f83...` | `test/fixtures/check53/53-reads-word365-simple.json` and `-all.json` (`ef57729`) |

The share's snippets are the commits' (`check52-script-lab.js` is `8fdccf8`'s, the run's; `check53-script-lab.js` is `ef57729`'s; both compared byte for byte). core-ts's `.docx` builders (`build.mjs`) are git-ignored output there, so a copy of the files Word read is the faithful source.

**The snippets in Python.** `tests/office_js/checks.py` ports `CHECK52.runCheck`'s three modes and check 53's, against the content API, with the same JSON shape and the same `vis()`:
- `parse_region`, the case parser;
- `range_of`, made from `search(..., match_case=True, match_whole_word=True)` and `expand_to`;
- `reads`, `apply` and `report`.

They are held, case by case, to Word's JSON:
- **Phase A** (checks 48, 48b and 51): every paragraph's `text` and whole range, `Body.text`, each query of 48's and 48b's searches, a cell's value and body text, the controls' text, a comment's content and the notes' bodies.
  - Check 48's file is held to 48b's run A (Word's default view), as core-ts holds it. Its own JSON is All Markup's: case 11's deleted text and the six hidden- and deleted-word queries are asserted as core-ts asserts them.
  - Check 48's `spans` need `expand_to`, so they join in Phase B.
  - Today 13 of 23, 13 of 25 and 16 of 25 paragraphs agree (section 2.2). After Phase A: every paragraph under Q1's full set; under the two decided characters, check 48's and 48b's case 3 and nothing else, every other difference pinned.
- **Phase B:**
  - check 52's `reads` mode on its file against the run's `reads` (16 kinds), and its `report` mode on Word's save after Accept All against Word's report;
  - the tracked changes a range lists on Word's save after the run;
  - check 53's snippet against the Simple Markup JSON.
- **Phase C:**
  - check 52's `run` mode on its file against Word's run, cases 1 to 85;
  - Accept All after it, saved and reloaded, against Word's report after Accept All;
  - each region's markup, a line per paragraph as core-ts's test writes it, against Word's save after the run;
  - core-ts's pinned lines for cases 16, 29, 39, 56 and 70.

**The differences core-ts asserts as they stand,** asserted here the same way, each with its reason:
- `marks` (`^p`, phase 3);
- Word's `Other` after a nested table (TC2; cases 89, 91 and 94; check 53's cases 10 and 12): Q6;
- `Range.footnotes` and `getTextRanges`, not implemented (the bold cases' words);
- a control showing its placeholder (cases 13, 26; check 53's 34 and 35), the readers';
- paragraph `font` (40 of 1,269 reads), CR-002's;
- the listing's grouping (11 regions);
- Word's extra empty paragraph on a control's content range (cases 60 and 66);
- the mark formatting Word gives on `insert_paragraph("After")` at a paragraph's end (cases 67 and 68, core-ts's departure);
- `insert_comment` across paragraphs (cases 79 to 85, phase 3);
- the Simple Markup readers of check 53 (case 25's joined paragraphs, case 33's `Body.text`);
- `font.color` (`#000000` against `""`, theme colours: CR-009's);
- CR-002 section 55's cases 86 to 95.

Anything else that differs is a failure. Wherever this package's own rules differ from core-ts's (section 4.9), the test asserts this package's answer and names the item.

**core-ts's `ranges.test.mjs`** (18 tests at `53db5bc`) is ported test for test into `tests/content/test_ranges.py`, its strings kept. That includes the file-and-snippet agreement (every `BEGIN` parses, every token is in the file once), the `TABLE_OOXML` package read, the run and accepted reports being of the same 95 cases, the saves' SHA-1s, and `accept_all` on Word's save after the run against Word's Accept All in 94 of 95 regions (case 26 Word's filler). core-ts's `readers.test.mjs` checks 48, 48b and 51 become `tests/content/test_office_js_characters.py`.

### 6.2 Phase A's gates

- **The readers' audit, before and after,** over the corpus CR-011 section 16.5 used (`out/cr012-b-corpus/`, its `COMMITS.txt` unchanged), with `scripts/mce_branches.py readers`. Phase F ran it with `--holding either` (384 documents); the characters move documents with no `mc` at all, so a new `--holding all` takes every distinct document (419 by content; 418 parse).
  - **Every answer that moves must be the character table's**, checked mechanically: each moved paragraph's after-text equals its segments read with the new map, it holds at least one item of the table (or is multi-paragraph text), and its before-text is the old map's.
  - **Expected to move:** `body.paragraphs`, `body.get_text` in both views, `body.element_at`, `body.find`, `body.search` and `body.tables` (cells of several paragraphs, or holding a break).
  - **Expected not to move:** `text.extract`'s text, `text.plain_text`, `text.markdown`, `package.to_markdown`, `body.outline`, `package.describe`, `revisions.tracked_changes` and both resolutions' text and canonical forms, `docx.paragraphs`, `docx.tables` and `docx.inline_shapes`, `body.inline_pictures`, `main.fonts_in_use` and `main.styles_in_use`.
  - **`text.extract`'s run segmentation** may move only as Q11 decides, and each move is reported.
- **The `office_offset` invariant:** `text.office_offsets`' disagreements stay 0, read under Q11's rule. Runs with an offset are counted before and after, and every change is explained (a note reference, hidden text, an equation).
- **The Word-measured fixtures:** checks 48, 48b and 51 (section 6.1).
- **The rest:** the suite, fast and full; CR-011's ratchet and goldens; CR-010's Word-save comparisons; the parity goldens; `scripts/roundtrip.py`; `generate.sh --check`; and acceptance (`compare_acceptance.py` against a copy made first).
  - Expected unchanged but for any artefact whose text a write holds a mapped character in. If any part differs, Phase A gets a Word check of its own.
  - `tests/test_readme.py` runs the changed fences.

### 6.3 Phases B and C's gates

- **The sweeps,** before and after, over the same corpus:
  - **Phase B** (core-ts section 9.5's): `Body.text`, each paragraph's text and whole range, each cell's text and `Comment.get_ranges()` unchanged; control, comment and tracked-change ranges moving as section 4.3 says, with the counts reported.
  - **Phase C** (core-ts section 10.5's): in each document one range from the middle of a paragraph to the middle of one up to four later. Six writes (`delete`, `Replace`, bold, `insert_paragraph` `"Before"` and `"After"`, and a `"Whole"` range's `delete`), each untracked, then tracked on a fresh load, saved and reloaded after every write. Accept All of the tracked write must equal the untracked write accepted, and Reject All must equal the original rejected, in each paragraph's text, style and alignment; refusals are counted by code.
- **Phase C's own tests** (core-ts's): the refusals before anything changes, a table, a control and a branch in both modes; Reject All after each tracked write giving back the original; a footnote leaving on save; a table's, a cell's, a control's and the body's whole range deleted; a split inside a link, and one refused inside a control; `insert_text` and `replace_text` across paragraphs; the `ChangeReport` and `dry_run` of each write.
- **The subset:** `tests/office_js_subset.json` re-derived from core-ts at `ef57729` or later (241 members); `test_the_script_would_write_what_is_committed` passing beside the sibling checkout (Q9); `test_every_implemented_member_is_there_with_the_right_kind` passing with every non-extension member implemented but the two omitted null-object members.
- **The rest:** the existing suite; the full suite; the ratchet; CR-010's comparisons; the parity goldens; the round trip; acceptance.
  - Phase B is read-only: acceptance unchanged but for Q7, which touches no artefact.
  - Phase C regenerates `out/acceptance/` with a new artefact.

### 6.4 The Word check (Phase D)

CLAUDE.md asks for a Word check after any change to the content API's inserts, and Phase C changes deletes, joins and splits. An artefact, built from a loaded document like artefact 8, holds ranges across paragraphs: from a heading's middle to a body paragraph's, a whole table between, a list item, a `"Whole"` range, a bold, a split. It is made untracked, and tracked in a second copy, with the engine's Accept All and Reject All beside them. Expected: no repair prompt; the untracked copy reads as check 52's untracked answers; the tracked copy in **All Markup** shows one deletion across the marks, with no formatting change on the joined paragraph; Word's Accept All and Reject All agree with the engine's. Phase A adds a file only if its writes changed an artefact (section 6.2).

## 7. Risks

1. **The characters are the largest change.** Every caller of `Body.text` and every paragraph with a break moves (section 4.1); under Q1's full set, stored offsets move too. 0.2.0 is not yet released, so the change ships once, in 0.2.0, with Phase F's. The audit (section 6.2) makes every moved answer explicable.
2. **Joins under tracking.** core-ts's risk too: where CR-002 sections 29 and 38 found Word's subtleties. Phase C reuses the resolution's `join_with_next` rather than writing a join, and Word's markup in check 52's save is the oracle.
3. **The oracles are Office JS code, and this package has no shim.** The Python port of the snippets could misread one: a missed `load`, an empty collection. The file-and-snippet agreement test, and running the port in `report` mode on Word's own accepted save (which must read as Word's report), catch a port that reads wrongly.
4. **core-ts is mid-edit.** Its working tree holds uncommitted CR-014 work (`Range.mts`, `spanEdits.mts`, `search.mts`, a `test/check54/`). If that changes `test/office-js-subset.ts` before this port lands, the gate moves under it. Q9 makes the derivation read core-ts's committed `HEAD`; this CR ports `53db5bc` and records what it ported.
5. **Offsets audited in one place.** core-ts's risk 3: code comparing `end` with `len(paragraph.text)` assuming one paragraph. Phase B audits every consumer of `Range` here: `hyperlinks.py`, `comments.py`, `ooxml.py`, `api_script.py`, `customxml/insert.py`, `reports.py` and the facade's `comments.py`. Each one either handles `end_paragraph` or refuses a range across paragraphs.
6. **Two readers of one comment.** D2 changes `Comment.get_range()`'s type. `api_script.py` and the README fence move to `get_ranges()` in Phase B, and the MCP server's contract (CR-006, proposed) is written against the new shape.
7. **Hidden text.** Under Q1's full set, `paragraph.text = paragraph.text` drops hidden runs, as core-ts's does. That is a real loss for a careless caller, and the setter's docstring says so.

## 8. Questions (decided 2026-10-23, each as recommended)

**Decided 2026-10-23, every one as recommended** (the user: "CR-013 accept recs for Q1 to Q11. For Q8, confirmed Apache-2.0 here"). Q8's files are the five files Word read from the share and Word's two saves of check 52. They may be used here under Apache-2.0 on that confirmation, and `tests/office_js/`'s manifest and `tests/README.md`'s provenance table record it, as they do for `mc-branches.docx`. The questions stand below as they were asked.

1. **Q1. Office JS's characters beyond the two decided: all of core-ts CR-012 phase 4's measured rules together, or only the paragraph mark and the line break?** The rest is: U+0002 for a note reference and a note's mark, U+000C and U+000E for a page and a column break, U+000D for `w:cr`, `0` for `w:ptab`, U+001E and U+001F for the hyphens, `(` for `w:sym`, hidden text left out by the text model's rule, an equation as Word's text of it, a paragraph's `"Whole"` range ending with its mark, cell, control and comment text as their content's `Body.text` (arguably within "multi-paragraph text" already), and `search`'s Find codes, U+FFFC for a picture and the optional hyphen passed over. *Recommended: all together, as one rule.*
   - They are one set of Word measurements (checks 48, 48b and 51, with 30 and 31), and core-ts adopted them as one rule (its D2, D6).
   - Parity is the aim (CR-003 section 23.3), and check 52's oracle embeds them: FN's text holds U+0002, HID's leaves the hidden run out, W1's whole range ends `\r`. Taking the two alone leaves those cases pinned.
   - Taken alone, the two leave `\n` meaning a page break, a column break and a `w:cr`, beside `\r\n`'s row end, so a `\n` would be ambiguous. Taken together, no item reads `\n` at all.
   - Taken alone, a `"Whole"` range must carry a mark it doesn't show (section 4.9 item 4).
   - Taking them later moves callers' offsets in a second release. Taking them now moves them once, in 0.2.0.
   - The cost: about 2 more days (`math_text.py`, the hidden rule in `segments_of`, the Find codes), offsets moving in at least 186 corpus paragraphs (124 with a note reference, 22 with hidden text, 40 with an equation), and `splitlines()` cutting a word at a non-breaking hyphen (section 4.1).
2. **Q2. Writing the characters back.** Six of them are not XML characters, and today a read text written back fails at save (section 2.5 item 3). *Recommended: the writers map each to the element it reads from:* U+000B `w:br`, U+000C a page break, U+000E a column break, U+000D `w:cr`, U+001E `w:noBreakHyphen`, U+001F `w:softHyphen`; `\n` stays a `w:br` and `\t` a `w:tab`, for 0.1.0's callers; U+0002 is refused (`text.not_writable`). Also: core-ts's build writes them raw into `w:t`, not well-formed XML, which goes to core-ts's session. Whether Word's `insertText` makes a paragraph of `\r` is unmeasured, and a check could ask it later.
3. **Q3. core-ts's D1 to D4, as decided there.** One `Range` with an end paragraph (D1); `Comment.get_range()` one `Range` or None, with `get_ranges()` the list (D2); whole tables read and deleted, Word's refusal inside one (D3); D4 is the editor's and has no counterpart here. D2 breaks `comment.get_range()[0]` on 0.x: `README.md`'s fence, `api_script.py` and the six test lines that index it move. *Recommended: adopt all as core-ts decided them.*
4. **Q4. `Font` where runs differ.** Across paragraphs core-ts reads `null`, `''` for the name, colour and highlight, and `Mixed` for the underline. *Recommended: `None`, `""` and `"Mixed"`, the getters typed `bool | None` and `float | None`.* A single paragraph keeps reading its first run until core-ts's CR-002 section 59 is decided there.
5. **Q5. Word's refusals as errors.** core-ts throws `GeneralException` for a range into or out of a table (Word's own) and `NotSupportedError` for writes not ported. *Recommended: `SpanError` with stable codes* (`range.partial_table`, `range.partial_control`, `range.spans_paragraphs`, `range.other_story`, `range.not_trackable`), the message saying Word refuses it too. That is CR-003 section 3.1's scheme and section 21.3 item 3's precedent.
6. **Q6. core-ts CR-002 section 57's departure.** Office JS reads a paragraph after a nested table in a cell as `Other` unless Normal. *Recommended: take the departure here too*, reading the paragraph's `w:pStyle`, for core-ts's reason (a heading's or a list item's style isn't hidden), asserted against Word's answer in checks 52 and 53.
7. **Q7. `Body.get_range()` on an empty body.** Today it inserts a paragraph, a read that writes. *Recommended: refuse with `ContentError` `body.no_paragraph`, as core-ts throws*, since CR-011 section 14.4 has every read leave the parts as they were.
8. **Q8. The oracles' fixtures.** *Recommended: copy the five files Word read from the share, and Word's two saves of check 52 from core-ts's commit, byte for byte with a SHA-1 manifest into `tests/office_js/` (outside `tests/fixtures/`), with Word's JSONs.* Porting core-ts's builders to Python would give files Word never read, built over this package's created package. The files are core-ts's builders' output (core-ts is Apache-2.0) and the user's own runs. As for `mc-branches.docx` (CR-011 section 16.11), the user is asked to confirm they may be used here under Apache-2.0.
9. **Q9. The subset's source.** `scripts/office_js_subset.py` reads core-ts's working-tree file, and core-ts is mid-edit (section 7 item 4). *Recommended: read core-ts's committed `HEAD` (`git show HEAD:test/office-js-subset.ts`) when the checkout is a git repository, falling back to the file.* It follows core-ts's commits rather than its edits; it is not a pin, so section 23.3's decision stands.
10. **Q10. core-ts's phase 3.** Its comments and controls on such a range, and `^p`, are uncommitted there. *Recommended: 0.2.0 waits for Phases A to C here (the subset's ten members and core-ts's committed phases 1 and 2), and phase 3 is ported as an amendment to this CR when core-ts commits it.* CR-003 section 23.3 names "the port of core-ts CR-014 phases 1 and 2".
11. **Q11. The `office_offset` invariant.** Phase F's audit holds each text model run's text equal to `get_text()` at its `office_offset`, character for character, and a break is `\n` in the model and U+000B in Office JS's text. *Recommended: keep the offsets as the extractor assigns them now (by item length, as core-ts's `Extractor.mts` does) and restate the invariant as agreement under the character map*, so the model's runs and REQ-706's JSON keep docx4j's segmentation. Splitting runs at every differing character would move the text model's JSON in every paragraph with a break.

## 9. Effort

| phase | days |
|---|---|
| A, the characters | 2 (the two decided) or 4 (Q1's full set, recommended) |
| B, the reads | 2½ to 3 |
| C, the writes | 3½ to 4 |
| D, the Word check and the records | ½ |
| **total** | **8½ to 9½** with the two characters alone; **10½ to 11½** with Q1's full set |

The larger part is not the members. It is the characters, and the oracles: a Python port of two Office JS snippets, held to Word's JSON case for case, where core-ts ran its snippets through its own shim. core-ts's own estimates for its phases 1 and 2 were 1½ to 2 and 2 to 3 days.

## 10. Phase A implementation notes (2026-10-23)

Section 4.1 as built, at the user's go ("start Phase A"), Q1's whole set and Q2's writers, with section 4.1's additions of 2026-10-23. The Office JS readers read Word's own character stream, as docx4j-core-ts CR-012 phase 4 (`9307bdb`) built it and Word 365 measured it (checks 48, 48b, 51); every text writer of the content API takes those characters back as the elements they read from; U+0002 and the other characters XML 1.0 excludes are refused at the write. On the files Word read, every paragraph of checks 48 (Word's default view), 48b and 51 now reads as Word 365's (13 of 23, 13 of 25 and 16 of 25 before). Nothing of Phases B and C is built: no `Range` across paragraphs, no new subset member, no `Comment.get_ranges()`, no check 52 or 53.

**Commits.**
- `da7aa12`: the audit first (section 6.2): `scripts/mce_branches.py readers --holding all`, and two readers Phase A moves (a block control's `text`, a comment's `content`).
- `e7e54e6`: the fixtures (Q8) in `tests/office_js/` with `SHA1SUMS`, and `tests/README.md`'s provenance rows.
- `316f427`: the pieces, none read yet: `model/content/math_text.py`, `model/content/story.py`, `PropertyResolver.is_vanished_by_styles`.
- `7369c63`: the readers and the writers, the new tests and the pre-existing expectations that move (10.12).
- `1bc9b7d`: the audit's `office_offset` agreement read under Q11's map; `README.md`, `docs/pypi-readme.md`, the CHANGELOG.
- The commit adding this section.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the characters | `model/content/text_model.py` | `OFFICE_CHARACTERS` and `office_item_text` (core-ts's `officeItemText`); `item_text` unchanged, docx4j's, which the facade's `run_text` reads (section 2.5 item 4) |
| the walk | `text_model.segments_of` | `hidden=` (a predicate, `hidden_runs_of`), `objects=` (pictures and hidden items as zero-width segments, `Segment.object` and `Segment.hidden`), `characters="docx4j"` (0.1.0's walker exactly), `visit=` (core-ts's hook: every element passed that adds no text, with its offset); an equation is one segment (`Segment.math`, its `item` and `run` the `m:oMath`) |
| hidden text | `text_model.hidden_runs_of`, `PropertyResolver.is_vanished_by_styles` | the text model's rule (REQ-507, `extractor._run`): a run's own `w:vanish`, else its paragraph style's and character style's; cached per pair, a pair naming a missing style not cached, and answered False at once where no style nor the defaults state `w:vanish` |
| equations | `model/content/math_text.py` (new, KEEP covers `model/`) | the port of `mathText.mts` at `9307bdb`: `math_text_of`, `math_letter`, `is_math` |
| the story layout | `model/content/story.py` (new) | `story_layout(items, text_of)` with each paragraph's `Place`; `mark_text_of`, core-ts's `lastInCell`; `story_of` is Phase B's |
| search | `text_model.office_matches`, `FIND_CODES`, `OBJECT_REPLACEMENT` | the port of `officeMatches` at `9307bdb`; every search site uses it (`Paragraph`, `Range`, `Body` through its paragraphs, `find_in`, a control's) |
| the writers | `text_model.WRITTEN_AS`, `text_items`, `text_run`, `explode_mapped`, `check_writable`, `unwritable` | Q2's map; a text written into a `w:t` by the edit primitives is then split around each mapped character, inside the same run, so the placing and the formatting are the primitives' own |
| equations in edits | `text_model.refuse_equation`, `refuse_tracked_equation`, `Paragraph.refuse_equation` | core-ts's `assertPlaceable`; pre-checks for a tracked `Body.clear()`, row and table deletions and `replace_text` |
| Q11 | `text_model.office_as_model` | Office JS's characters read as the text model's, offset for offset |
| readers moved | `paragraph.py`, `body.py`, `range.py`, `table.py`, `controls.py`, `comments.py`, `tracked_change.py`, `reports.py`, `model/text/extractor.py` | 10.3 |
| writers moved | `paragraph.py` (`text` setter, `splice`, `insert_paragraph`, the tracked deletion and insertion), `body.py`, `table.py` (`value`, `values`, `insert_table`, `add_rows`), `comments.py`, `controls.py`, `tracking.py` (a deleted row's cells) | 10.4 |
| the write-back | `controls.ContentControl.docx4j_text` (extension), `customxml/bindings.update_from_control` | 10.4 |
| the facade | `docx/text.py` (`check_python_docx_text`, `Run.text`, `add_text`), `docx/_base.py` (`add_paragraph`), `docx/comments.py` (`Comment.text`, `add_comment`) | 10.4 |
| the trial | `model/content/trial.py` | 10.5 |
| docstrings | every reader and writer above; `model/text/result.py`'s `office_offset` | the Q4 boilerplate kept, Office JS's characters added |
| tests | `tests/content/test_office_js_characters.py` (127, new); 26 pre-existing tests in 17 files (10.12) | 10.6 to 10.9 |
| fixtures | `tests/office_js/` | 10.6 |

### 10.2 The characters as built

Section 3.8's table, every row: a `w:br` U+000B (U+000C `w:type="page"`, U+000E `"column"`), a `w:cr` U+000D, a `w:ptab` `0`, the hyphens U+001E and U+001F, any `w:sym` `(`, a footnote or endnote reference and a note's own mark U+0002, an equation one segment of Word's text of it, hidden text nothing, a comment reference, a field code and a picture nothing. A kept branch's wildcard items read the same table.
- **Hidden runs** have no offset. With `objects=True` their text items, and a drawing in them, are listed as zero-width segments marked `hidden`, so that a deletion can find them; every other reader passes them by.
- **An equation** at run-holder level (`m:oMath`, `m:oMathPara`, in a paragraph or in a run holder) is one segment, not editable, its `item` and `run` the element. `Range.runs` and so a `font` write skip it, as core-ts's do.
- **`characters="docx4j"`** is the 0.1.0 walker exactly (`item_text`, no hidden filter, no equation), which a bound control's write-back reads (10.4). The audit's check (10.7) uses it as "the old map".

### 10.3 The readers

**Moved:**
- `Paragraph.text`, `get_text` in `accepted` and `original`, `text_budget`, `segments()` (now with `hidden_runs()`), `to_dict()`, `repr`;
- `Body.text`, `get_text` (the budget now cut from the whole layout) and `text_budget`: `story_layout` over the body's content, its last `\r` dropped;
- `Range.text`: a slice, and for a paragraph's `'Whole'` range (`Range.mark`, 10.11 item 1) the paragraph's `mark_text`;
- `TableCell.text` and `value`: the cell's `Body.text`; so `TableRow.values` and `Table.values`;
- `ContentControl.text`: a block, row or cell control's content's `Body.text`; a run control's span of `Paragraph.text`;
- `Comment.content`: its body's `Body.text`, unchanged in code;
- `search`, `find`, `replace_text`'s matches, `SearchHit`'s context, `paragraph_at(contains=)`, a `ChangeReport`'s `text_before` and `text_after`;
- the offsets' own walkers: the comment markers (`comments._markers_in_paragraph`, through `segments_of`'s `visit`, with the package's resolver passed where the offsets are read) and a revision's collapsed place (`tracked_change._accepted_offset_of`, the same way);
- the text model's `office_offset`: `extractor._office_spans` reads the same segments with the same hidden rule and resolver; a run keeps its offset where the item's length agrees (Q11), and the merge rule is unchanged.

**Not moved, each documented as not Office JS's:** `traversal.text_of`; `Block.text` and `Table.text` (a line per paragraph, docx4j's characters); the facade's `run_text`, `holder_text` and, since this phase, its `Comment.text` (10.4); `TrackedChange.text`; the text model, its plain text and Markdown, `outline()` and `describe()`.

### 10.4 The writers

- **The map** (Q2): U+000B and `\n` a `w:br`, U+000C a page break, U+000E a column break, U+000D a `w:cr`, U+001E a `w:noBreakHyphen`, U+001F a `w:softHyphen`, `\t` a `w:tab`. Each character is one element, so `\r\n` in a paragraph's text is a `w:cr` and a `w:br`; in a comment it is one mark.
- **Every text writer**: `Paragraph.text = ...` (untracked a run of `text_items`; tracked the splice), `splice` and so `insert_text` on a paragraph, a range, a body and a run control's span, `replace_text` on all three, `insert_paragraph` on a body, a paragraph, a range, a control and a list, `Body.insert_text`, `TableCell.value = ...`, `TableRow.values` and `Table.values`, `insert_table`'s and `add_rows`' values, `insert_comment`, `Comment.content = ...`, `reply`, a control's `insert_text` and `insert_paragraph`. Untracked and tracked: the tracked insertion's `w:t` is split the same way inside its `w:ins`.
- **Measured before** (a read-only probe at `da7aa12`): the splice-based writers (`Paragraph.text = ...`, `insert_text`, `replace_text`) wrote `\t` and `\n` raw into a `w:t`; only the paragraph, cell, table and comment writers mapped them, through the `r` builder. So "as now" in section 4.1 was true of those; the others now map them too.
- **The refusals**: `check_writable`, at the top of each writer and again in `splice`, refuses U+0002 ("no text makes a note"), U+0000 to U+0008, U+000F to U+001D, a lone surrogate, U+FFFE and U+FFFF with `ContentError` `text.not_writable`, naming the character and its offset, before anything changes. A table's `values` and `insert_table` check every value first. lxml's refusal at save stays behind it.
- **A comment** splits on `\r\n`, `\r` and `\n` (`_comment_paragraphs`); `"one\rtwo\nthree"` now makes three paragraphs where 0.1.0 saved `one&#13;two` and `three`.
- **A bound control's write-back** reads `ContentControl.docx4j_text`: a run control `text_of`'s, any other its paragraphs' 0.1.0 text a `\n` apart. Measured at `e7e54e6` in a scratch worktree and again after: a run control holding a break writes `line one\nline two`, a block control of two paragraphs (the second holding a break) `first\nsecond\nthird`, the same both times; the tests pin both.
- **The python-docx facade, measured in python-docx 1.2.0** (a throwaway venv in the session's scratch directory, as `scripts/python_docx_subset.py` says; never a dependency):

  | write | python-docx 1.2.0 | 0.1.0 here | Phase A here |
  |---|---|---|---|
  | `add_paragraph`, `add_run`, `Run.text =`, `Paragraph.text =`, `cell.text =` with `\t` | `w:tab` | `w:tab` | `w:tab` |
  | the same with `\n` | `w:br` | `w:br` | `w:br` |
  | the same with `\r` | `w:br` (`\r\n` two) | `w:br` in `add_paragraph`, `&#13;` in a `w:t` in the rest | `w:br` (`\r\n` two) |
  | any other control character (U+000B, U+000C, U+000E, U+001E, U+001F, U+0002, U+0000) | `ValueError` at the write | accepted, failing at save | `ValueError` at the write |
  | `add_comment(text=)` | `\n` a paragraph, `\r` a `w:br` | `\n` a paragraph, `\r` `&#13;` | `\n` a paragraph, `\r` a `w:br` |
  | `Comment.text` | paragraphs `\n` apart | the content API's `content` | `holder_text` of each paragraph, `\n` apart |

  `Comment.text` had to move off the content API's `content`, which is `\r`-joined now; for a comment of plain runs the answer is 0.1.0's.

### 10.5 Edits over hidden text and equations

- **Hidden text** (core-ts's rule, and Word's per check 52's HID, which Phase C holds): a hidden run strictly inside a deleted span goes with it; tracked, it is inside the `w:del`. At the span's edges it goes only when the span is the paragraph's whole text or the caller says `whole` (`Paragraph.text = ...` tracked, a tracked `Paragraph.delete()`, a deleted row's cells), core-ts's `WHOLE`. So `paragraph.text = paragraph.text` drops the hidden runs, tracked or not (risk 7, the setter's docstring). An insertion beside hidden text goes beside it, never into it.
- **An equation**: an edit may take one whole, untracked; cutting one or inserting inside one is refused (`SpanError` `range.cuts_equation`), and so is `split_at` inside one, so a `font`, comment, control or link whose span edge falls inside one refuses; under change tracking any edit taking one is refused (`TrackedChangeError` `tracking.equation`). A tracked `Body.clear()`, a tracked row or table deletion and `replace_text` check first, so nothing changes before the refusal. `Range.get_xml` keeps an equation whole where its span holds part of it, as it keeps a kept branch.
- **The dry run.** Every `Paragraph.text` now asks the styles which runs they hide, so a trial's first paragraph read made a `TrialPart` of `styles.xml`, which copied `part.contents` and so unmarshalled the real part: `tests/agent/test_dry_run.py::test_the_real_package_is_byte_for_byte_what_it_was` failed (`word/styles.xml` re-marshalled). A trial now copies a part the real package has not unmarshalled from its private reading (`read_contents`, CR-002 section 16) and leaves it unread. The module's docstring says so.

### 10.6 The oracles: checks 48, 48b and 51

**The fixtures** (Q8): `48-text.docx`, `48b-text.docx` and `51-text.docx` from the user's share, `__tmp/48/` and `__tmp/51/` (SHA-1 `8dec79fdd0dffa36051d572c6093e6515321ac01`, `ca98e9c0bd6062233383ae9181cbd1893190f7dd`, `7a6a0102051387e17d04e2eed87c46373c955005`, each checked against the brief's); `48-text-word365.json` (`30b559e`), `48-text-views-word365.txt` and `48b-text-word365.json` (`5c6cb5c`), `51-text-word365.json` (`9307bdb`), each by `git show` and checked against its git blob. `SHA1SUMS` holds all seven, and the first test checks it.

| oracle | before (`e7e54e6`) | after |
|---|---|---|
| check 48, view A (`48-text-views-word365.txt`), paragraphs | 13 of 23 | 23 of 23 |
| check 48b take 2, paragraphs | 13 of 25 | 25 of 25 |
| check 51, paragraphs | 16 of 25 | 25 of 25 |
| 48b run A's searches | 9 of 12 | 12 of 12 |
| 48b take 2's searches | 6 of 12 | 12 of 12 |
| check 48's own searches (All Markup) | 22 of 41 | 35 of 41: the six view-dependent queries core-ts asserts as they stand |

Also held, as core-ts holds them: check 48's every paragraph's `'Whole'` range (case 11's All Markup text read in Simple Markup) and `Body.text` (its table's row end `\t\t` read as Word's default view's `\r\n`); check 51's every `'Whole'` range, `Body.text`, `Table.values`, each cell's `value`, body text and paragraphs, a cell's last paragraph's `'Whole'` ending `\t`, the block control's and the inline control's text, the inline control's `'Whole'` and `'Content'`, the comment's content, and the footnote's and the endnote's paragraphs, body text and last paragraph's `'Whole'`. **Pinned as Phase B's:** the block control's `get_range("Whole")`, `CCP1\r` here (its first paragraph's) where Word gives `CCP1\rCCP2\rCCP3\r`. Ported from `readers.test.mjs` too: the hidden-text edit test, the search test (`^f` and `^e` apart, a range's own span), and the equation test (`math_letter`, a cut refused, a format leaving the math runs alone, taken whole, tracked refused). Its text-box, `paragraphAt` and `outline()` tests are core-ts's D3 to D5 and not this phase's.

### 10.7 The readers' audit, before and after

`scripts/mce_branches.py readers --holding all` over `out/cr012-b-corpus/`, its `COMMITS.txt` unchanged (this repository `f1bb1d9`, core-ts `df5b4af`, docx4j `901a922a3`): before at `da7aa12` (`out/cr013-a/readers-before.json`), after at `1bc9b7d` (`readers-after.json`), `compare-readers` in `compare.txt`. **422 distinct documents**, not section 6.2's 419: 419 `.docx`, 2 `.docm` and a `.dotm`; 418 load, the four Strict documents failing both times (4 reader errors each).

| reader | documents moved |
|---|---|
| `body.get_text`, `body.get_text.original` | 356 each |
| `body.paragraphs`, `body.element_at` | 105 each |
| `text.office_offsets` | 89 |
| `body.tables` | 20 |
| `body.find` | 18 |
| `body.content_controls` | 16 |
| `body.search` | 10 |
| `body.comments` | 4 |

**Not moved in any document:** `text.extract` (the text model's JSON, its run segmentation included: Q11 moved no run), `text.plain_text`, `text.markdown`, `package.to_markdown`, `body.outline`, `package.describe`, `revisions.tracked_changes`, both resolutions' text and canonical forms, `docx.paragraphs`, `docx.tables`, `docx.inline_shapes`, `body.inline_pictures`, `main.fonts_in_use`, `main.styles_in_use`.

**The mechanical check** (`out/cr013-a/check_audit.py`, its output `check-audit.txt`), over each of the 357 documents with a moved answer, loaded again at the after commit: **0 problems.**
- Each paragraph's before text is the old map's (`text_of_view(characters="docx4j")`), its after text `Paragraph.text`, and each of the **250 paragraphs that moved** holds an item of the table: a note reference 54, a break 45, both 58, an equation 40, hidden text 30, a `w:sym` 12, both hyphens 8, an optional hyphen 3.
- Each document's `get_text` in both views is, before, the old `\n` join of the old texts, and after, Office JS's layout of the after texts by the check's own walk (not `story.py`'s).
- Each cell's value is, before, `text_of`'s and, after, its content's layout; each control's and comment's text likewise.
- Each `find` hit lies in its paragraph's after text. `search`'s hits went from 61,146 to 61,122: the 24 lost are all `e`s in hidden text (10 documents, 2 or 4 each, every one counted); `find`'s, at most 20 a word, stayed 8,475, their contexts and offsets moving.

### 10.8 The `office_offset` invariant (Q11)

Runs with an offset: **16,224 before and 16,224 after**; disagreements with `get_text()`: **0 before** (character for character) and **0 after** (under the character map, `office_as_model`, with the package's hidden rule). 214 paragraph views' offsets moved, each explained: a note's own mark at the start of a note's paragraph (181), a note reference (27), an equation (6). No run gained or lost an offset, and no block's run segmentation moved (`text.extract` unchanged).

### 10.9 The gates

| gate | before | after |
|---|---|---|
| the fast suite (`-m "not slow"`) | 3,982 passed, 8 skipped, 1,280 deselected, 1 failed | 4,109 passed, 8 skipped, 1,280 deselected, 1 failed |
| the full suite | 5,262 passed, 8 skipped, 1 failed (3 min 03 s) | 5,389 passed, 8 skipped, 1 failed (3 min 15 s) |
| `tests/content/test_office_js_characters.py` | --- | 127 passed |
| CR-011's ratchet (`tests/text/conformance.py`) | 1,291 passed, 0 failed, 293 pending of 1,584; goldens 12 of 12 | the same; `conformance.tsv` byte for byte |
| `tests/text/test_conformance.py`; `tests/text` | 24; 1,188 | 24; 1,188 |
| CR-010's Word-save comparisons (`test_word_saves.py`, `test_compare_word_resolution.py`); `tests/revisions` | 68; 550 | 68; 550 |
| the parity goldens (numbering, properties, fonts) | 277, `PINS` empty | 277, `PINS` empty |
| `scripts/roundtrip.py` | --- | 50 of 50 canonically identical, 0 differences, 0 skipped, 34 benign respellings; exit 0 |
| `codegen/generate.sh --check` | --- | byte identical (296 modules); no generated file changed |
| acceptance (`compare_acceptance.py` against `out/acceptance-before-cr013a/`) | copied first; equal canonically to a regeneration at `da7aa12` | 13 artefacts, 175 entries, 0 differing beyond the dates; against the regeneration at `da7aa12`, 167 entries the same bytes and 8 in the dates only |

- **The one failure, both times:** `tests/content/test_office_js_subset.py::test_the_script_would_write_what_is_committed`, core-ts's subset having moved (section 1); Phase B re-derives the subset.
- **The new tests are the whole difference** in the suites (127 and 127).
- **Acceptance did not move**, since no artefact writes one of the mapped characters through a splice. So the Word check (10.10) is the writers' own.

### 10.10 The Word check (staged 2026-10-23; not yet run)

The writers' map changes what the content API's inserts write, so CLAUDE.md asks for one. In `~/vbShares/Office 2016/__tmp/python-cr013a/`, built by `out/cr013-a-word-check/build.py` and each read back here with `LoadOptions(strict=True)`, nothing skipped:
- `cr013a-SOURCE-writers.docx`, a created document of nine labelled cases, and its two saves: `cr013a-1-writers-untracked.docx` and `cr013a-2-writers-tracked.docx` (author `CR-013 check`). Through `Body.insert_paragraph`, `Paragraph.text = ...`, `Range.insert_text` at an end and in place of a span, `replace_text`, `TableCell.value = ...`, `insert_table`'s values, `insert_comment` and `Comment.content = ...`, and a run control's `insert_text`: U+000B, U+000C, U+000E, U+000D, U+001E, U+001F, `\t` and `\n`, and comments holding `\r`, `\n` and U+000B.
- `cr013a-SOURCE-bound-control.docx` and `cr013a-3-bound-control.docx`: a multi-line plain-text control and a block control bound to a custom XML part, written back after the first holds a break and the second two paragraphs.
- `README.txt` says what to look at in each file and tells the user to view file 2 in **All Markup**.
- Expected: each opens without a repair prompt and shows the README's text. File 3's question is what Word shows in each bound control after refreshing it from the XML's `\n`.

**Run 1, 2026-10-23** (the user's notes, `__tmp/python-cr013a/results.txt`, with `1.PNG` and `cr013a-SOURCE-bound-control.PNG`):
- **Files 1 and 2** opened; no repair was reported. Each of the user's notes, read against the markup:
  - **Case 3**, "SEED3 before COLA": as built. The label ends with `SEED3`, and the insertion goes at its end. The README didn't say so.
  - **Case 4**, "a new paragraph (pilcrow) after CRA": the save holds `CRA<w:cr/>CRB END4` in one `w:p`, and Word draws the `w:cr` as a paragraph mark. That agrees with check 48, where Word reads a `w:cr` as U+000D, a paragraph mark's character. The README's "Word draws it as a line break" was wrong. The writer stays: a `w:cr` is the element U+000D reads from.
  - **Case 7**, "no PAGE BREAK or other char between PAGE7A and PAGE7B": the cell holds `PAGE7A<w:br w:type="page"/>PAGE7B`, and Word draws nothing for a page break in a table cell. That Word ignores a page break in a cell is believed, not measured here.
  - **Case 8**, "COMP3B is on a line of its own": as built. The comment's third paragraph is `COMP3<w:br/>COMP3B`, a line break.
  - **Case 9**: the user confirmed it is not shown as a tracked insertion, as the README said (section 10.13).
- **The bound source** shows "World" and "Note", which is what its custom XML holds (`<to>World</to><note>Note</note>`). The README didn't say what the source holds.
- **File 3 failed:** Word could not open it and recovered it (`cr013a-3-bound-control-word365-recovered.docx`).
  - The recovery adds `<w:text/>` to the block control (no type, with a `w:dataBinding`), so Word reads it as plain text. It joins the control's two paragraphs into one (`Notesecond paragraph`). In the node, a newline becomes `\r\n`.
  - **0.1.0 writes the same** block control and the same node: measured in a `v0.1.0` worktree, loading the source, `insert_paragraph("second paragraph", location="End")` on the block control, then `update_from_content_controls()`. So the failure predates Phase A, and this check found it.
  - Four variants, each file 3 with one change undone, are staged beside it (`README-3-variants.txt`): 3a keeps two paragraphs with the node back to `Note`; 3b has one paragraph with the node keeping its newline; 3c changes only the plain control; 3d is the block control unbound. Their results decide the fix.

**Run 2, the variants, 2026-10-23** (the user: "3a unreadable; others 3b 3c 3d open fine"):
- **3a fails**: two paragraphs, with the node back to `Note`. **3b, 3c and 3d open**.
- **The cause:** Word refuses a block control that has no type and a `w:dataBinding` and holds more than one paragraph. Neither the node's newline (3b) nor the plain-text control's break and its write-back (3c) is the cause. Unbound, two paragraphs open (3d).
- **What Phase A changed passes in Word:** the writers' markup in files 1 and 2, and the write-back in docx4j's characters (3b, 3c).
- **The defect** is the content API's, and 0.1.0 has it. A control's `insert_paragraph` (and any write that adds a block to its content) gives such a control a second paragraph, and the save is one Word must repair. Its fix waits on the user (section 10.13).

### 10.11 Departures, with why

1. **`Range.mark` lands in Phase A**, keyword-only (section 4.8), within one paragraph: the `'Whole'` range ending with its mark is Q1's and section 6.1's Phase A gate, and core-ts's phase 4 added the same flag before its CR-014. Section 4.2's other flags (`whole`, `end_mark`) and the second end are Phase B's. `__eq__` and `__hash__` take `mark`; `__len__` stays `end - start`.
2. **The edge rule for hidden text** is computed, not flagged: a span that is the paragraph's whole text takes the hidden runs at its edges, whatever made the range, where core-ts's `WHOLE` comes only from a paragraph's `'Whole'` or `'Content'` range. A search hit covering a whole paragraph's text takes them here and not in core-ts. Phase B's `whole` flag can replace it.
3. **The equation refusals' codes** are new and not in section 4.7: `SpanError` `range.cuts_equation` and `TrackedChangeError` `tracking.equation` (core-ts's `NotSupportedError`).
4. **A tracked deletion that takes an equation is refused**, core-ts's rule as section 3.8 records it, though section 4.1's list of what moves for callers names only "an edit that cuts an equation". It reaches a tracked `Paragraph.delete()`, `Body.clear()`, a row's or table's deletion and a replacement, where 0.1.0's tracked deletion skipped the equation (it was no text) and left it undeleted. `tests/content/test_tracking.py`'s `every_tracked_verb` now deletes the first paragraph from the third on holding no equation, since `samples/2010-sample1.docx`'s third is its equation (10.12).
5. **A picture in a deleted span stays**, this package's rule (core-ts's CR-002 section 48 is not ported); a drawing in a hidden run goes with the run, and `search`'s `^g` passes it by.
6. **A dry run copies an unread part privately** (10.5): needed to keep the trial's promise once `Paragraph.text` reads the styles.
7. **The facade's `Comment.text`** is python-docx's `holder_text` join, not 0.1.0's content API text; the same for a comment of plain runs (10.4).
8. **`story_layout` lays out a list of rows in Word's row form**, so a row-level control's `text` is its cells `\t` apart with each row ended `\r\n`. Check 53's rule for a row control (cells `\t` apart, `'Whole'` ending `\t\t`) is Phase B's.
9. **The audit gained two readers** (`body.content_controls`, `body.comments`) in its first commit, so that the before run has them.

### 10.12 Pre-existing tests whose expectation changed, and why

Each moves with the characters, the layout or the `'Whole'` range's mark; each carries a comment naming CR-013 Phase A.
- **The paragraph mark, `\r`:** `test_body.py::test_insert_paragraph_text_and_the_docx4j_aliases`, `test_comments.py::test_setting_the_content_replaces_the_paragraphs_and_keeps_the_para_id`, `test_controls.py::test_insert_into_a_block_control`, `test_controls_typed.py::test_a_group_control_wraps_the_whole_body`, `test_table.py::test_a_cells_body_carries_the_address_prefix`, `revisions/test_content_api.py::test_get_text_keeps_office_jss_filtered_text`, `openpackaging/test_mce_typed.py::test_the_text_of_readers_read_a_paragraph_level_branchs_runs_as_at_run_level` (a table's row form).
- **The breaks:** `test_body.py::test_insert_text_and_insert_break` (U+000C), `test_paragraph.py::test_insert_break_in_and_around_the_paragraph` (U+000B, U+000C), `test_text_model.py::test_the_segments_of_a_paragraph` (U+000B, the hyphens, `(`; `text_of` keeps docx4j's), and `test_word_branch_text.py`: case 2's break is U+000B, the pin `test_case_2s_break_is_a_newline_where_office_js_gives_u000b` renamed `test_case_2s_break_is_u000b_as_office_js_gives_it` (CR-011 section 15 item 2, superseded), and `text_of`'s agreement and the offsets read under the map.
- **The `'Whole'` range's mark:** `test_body.py::test_clear_and_get_range`, `test_comments.py::test_get_comments_on_a_paragraph_and_on_a_range`, `test_range.py::test_the_original_view_is_not_offered_on_a_span`, `agent/test_errors.py::test_current_is_the_accepted_view`.
- **Hidden text left out:** `agent/test_outline.py::test_the_entries_text_is_the_text_models` (`Paragraph.text` no longer keeps it, so the two readers agree); `text/test_texts.py::test_hidden_text_is_the_gap_between_offsets` became `test_hidden_text_is_in_neither_text`, and `test_a_runs_office_offset_is_where_its_text_is_in_get_text` reads under the map.
- **Equations:** `test_tracking.py::test_rejecting_both_mark_forms_of_a_loaded_document_leaves_the_source` (the equation paragraph reads its text, not `""`), and `every_tracked_verb` (10.11 item 4), which `test_rejecting_what_was_appended_at_the_end_follows_word` and `test_accepting_every_tracked_verb_gives_what_the_untracked_run_gives` run over `2010-sample1.docx`.
- **The dry run:** `agent/test_dry_run.py::test_the_real_package_is_byte_for_byte_what_it_was` failed until 10.5's fix; its expectation is unchanged.

### 10.13 Found, and not done

- **`to_api_script`'s note on a `w:tab` or `w:br` outside the leading run** (CR-003 section 19.3) is now stale: `insert_text` writes both as elements, so such a script is exact. The note and the `exact` flag are left as they were (conservative).
- **Two writers that ignore change tracking**, before this phase too: `TableCell.value = ...` clears the cell's old content untracked, and a run control's `insert_text` at its `Start` or `End` edits the control's items directly. The tracked Word-check file shows both (its README says so).
- **Phase B's:** a block control's, a comment's, a tracked change's and the body's ranges across paragraphs; `Range.is_empty`; check 53's row-control text; `story_of`.
- **The Word check** of 10.10 has had two runs (above); its record goes into `tests/README.md` when Phase A closes.
- **The scratch worktree** of `e7e54e6` used for the before measurements was removed after use; the measurements are in this section.

### 10.14 Word's own writers, measured after the build (docx4j-core-ts check 55, `a923aeb`)

The user ran core-ts's check 55 in Word 365 (16.0.20326.20232) after Phase A was built. It has 58 cases, each writing `x`, a character, `y`. The writers were `insertText` at the end (also tracked), `insertParagraph` after, `insertComment` and `comment.content`. The snippet builds each text with `String.fromCodePoint`, so the characters are real ones. core-ts's CR-002 section 60.3 records it.
- **Agrees with Q2 as built:**
  - `\t`, U+000B, U+000C, U+000E, U+001E and U+001F became their elements in every writer, tracked too;
  - in a comment, `\r`, `\n` and `\r\n` each made a paragraph, `\r\n` as one.
- **Differs:**
  1. In `insertText`, `\n`, `\r` and `\r\n` each made a new paragraph, `\r\n` as one, its mark marked inserted when tracked. Here `\n` is a `w:br` and `\r` a `w:cr`. Run 1's case 4 agrees with Word's direction: Word draws a `w:cr` as a paragraph mark.
  2. `insertParagraph` wrote the escapes as text (a backslash, then `n`).
  3. U+0002 and U+0001 were gone by the save, and U+0007 at once. Word refused none of them; here they are refused (`text.not_writable`).
- **Open in core-ts, the user's:**
  - E1: `insertText` makes a paragraph, as Word does; recommended there.
  - E2: `insertParagraph` keeps Q2's map as a recorded departure; recommended there, and that is what this phase built.
  - E3: the excluded characters dropped, as Word drops them; recommended there.
  - E4: a check 55b for bound controls.
  
  E1 and E3, if taken, amend Q2 here. Both ports agreed to follow Word's measurement together.
- **Bound controls:** a data point for E4. In run 1, Word's own save of the recovered file 3 wrote the plain-text control's line break into its node as `\r\n` (`<to>line one\r\nline two</to>`). Phase A writes `\n`, docx4j's character, and Word opened that (3b, 3c).
