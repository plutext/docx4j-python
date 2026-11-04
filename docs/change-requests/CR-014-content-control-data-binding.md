# CR-014: Content control data binding --- a bound control Word reads as plain text holds one paragraph (phase 1); rich text, XHTML and picture bindings and the OpenDoPE processor named for later

**Status:** Proposed 2026-10-23, at the user's request ("draft the new CR here for content control data binding with phase 1 only proposed for now"), after CR-013 Phase A's Word check found a file this package writes that Word cannot open (CR-013 section 10.10): a block content control that has no type and a `w:dataBinding` and holds two paragraphs. 0.1.0 writes it too. Data binding here has been CR-003 Phase E's (sections 3.7, 4 and 17), and no CR has owned it since; this one does. Phase 1 (sections 3 and 4) is proposed. core-ts's Word check 56 (`38dcae9`) measured the rule's scope for both ports: Word can't open **any** bound block control holding a second block, whether it has no type, `w:text` or `w:richText`, and whether the second block is a paragraph or a table. On open Word reads each as plain text. So such a control keeps one paragraph: the writes that would give it a second are refused, or joined where Word joins them, `set_mapping` refuses a control that already holds more, and `apply_bindings` writes one paragraph, as docx4j's `bind.xslt` does. Phases 2 (rich-text, XHTML and picture bindings) and 3 (the OpenDoPE processor) are named in section 4 and not proposed. Nothing is implemented; phase 1 waits on the user's answers (section 7) and go. Phase 1b (section 2.9, Q10), a bound date formatted as Word formats it, was proposed the same day from docx4j's hand-off of its CR-058 phase 1 (`d8d8bfb71`). A probe here found `am/pm` in a `w:dateFormat` written as `a5/p5`, names in English only, and the DATE field converter's `h` running 0 to 11. Q10 was decided the same day as recommended (the user: "CR-014 Q10 agree"): phase 1b goes with phase 1, the names from a committed CLDR-derived table. Q1 to Q9 are open. Phase 1c (section 2.10, Q11), a store item ID that names no part resolved as Word repairs it, was proposed the same day from docx4j's second hand-off (`35cdeaee6`). Q11 is open. **Decided 2026-11-02** (the user: "CR-014's questions Q1 to Q9 and Q11 as recommended, noting for Q7 that 0.2.0 is already out"; section 7): Q1 to Q9 and Q11 as recommended, with Q10 decided 2026-10-23, so phases 1, 1b and 1c go together. Q7: 0.2.0 shipped on 2026-10-25 without phase 1, so it goes in the next release. **Amended 2026-10-28** (CR-025 Phase C): Q10's nominal default stands, and Word's bound-date conversion is to be a separate per-package opt-in, `pkg.bound_dates_converted_as_word`, off by default (the ports' date rule, point 3; section 7, Q10). **Phases 1, 1b and 1c implemented 2026-11-05** (sections 9 to 13; `c794583` and `871be28`, `ce7f5c1`, `d0fe5e8`): the refusals, the join, `set_mapping`'s refusal, one paragraph from `apply_bindings`, the reading and Q9's bindings as text; a bound date through the field code's converter, in CLDR's names from a committed table, with the opt-in; and `StoreItemResolver`'s port. **Word-checked 2026-11-05** (the user, every file opened without a repair prompt, section 13.2): phases 1, 1b and 1c are complete. Section 2.9's "Word's documentation, not Word's run" is corrected there: docx4j's `9b9d053cb` measured `h` as 1 to 12.

**Depends on:**
- **This repository:** CR-003 Phase E (`docx4j_py/model/customxml/`, `ContentControl`; sections 3.7, 4 and 17); CR-013 (its writers: Phase A's amendment, section 11, makes `insert_text` split paragraphs, and Phase C writes across them; both must keep this rule); CR-002 Phase C's flat OPC sink, for phase 2 only.
- **docx4j**, cited at `30fb41024` (`VERSION_17_3_5`): `org.docx4j.model.datastorage` (`bind.xslt`, `BindingTraverserXSLT`, `UpdateXmlFromDocumentSurface`, `OpenDoPEHandler`). For phase 1b, `d8d8bfb71` (`VERSION_17_3_5`, unpushed; its CR-058 phase 1): `BindingTraverserXSLT.xpathDateRun` and `FormattingSwitchHelper.formatDate`. For phase 1c, `35cdeaee6` (its CR-058 phase 2): `StoreItemResolver`, `BoundValues`.
- **docx4j-core-ts:** its CR-002 phase E (the binding step this package ported), its CR-002 section 61 (the same defect there: D1 to D4 at `80a933d`, check 56 and H1 to H4 at `38dcae9`, decided 2026-10-23, not yet implemented), check 55b (`1e61bb1`, Word's bound controls), and its CR-005 (`05e1370`, the OpenDoPE processor: phase A done, phase B proposed, sections 8 and 9 released in its 0.4.0).

**Counterpart:** docx4j `org.docx4j.model.datastorage`; docx4j-core-ts CR-002 phase E and section 61, and CR-005.

## 1. Why

1. **A file Word cannot open.** CR-013 Phase A's Word check (its section 10.10) bound a block control with no type element to a custom XML node, gave it a second paragraph with `ContentControl.insert_paragraph(location="End")`, and wrote it back with `update_from_content_controls()`. Word 365 could not open the save and repaired it: it added `<w:text/>` to the control, so reading it as plain text, and joined the two paragraphs into one. CLAUDE.md's rule is that saved output opens without a repair prompt.
2. **It is not new.** 0.1.0 writes the same control and the same node (measured in a `v0.1.0` worktree, CR-013 section 10.10).
3. **No CR owns binding.** CR-003 Phase E built the binding step (`apply_bindings`, `update_from_content_controls`, `XmlMapping`, `describe()`, `fill()`) and section 17.6 listed what it left out: OpenDoPE's conditions, repeats and handlers; picture bindings; explicitly rich-text bindings. core-ts has a CR for that ground (its CR-005), and this package has none.

## 2. What exists here (measured 2026-10-23 at `c5b466f`)

### 2.1 Binding today (CR-003 Phase E)

`docx4j_py/model/customxml/bindings.py` (500 lines) is docx4j's `BindingHandler.applyBindings` and its reverse, `UpdateXmlFromDocumentSurface`, with CR-003 section 4's departures:
- **containers are never bound** (a repeating section and its items, a group, a building-block gallery, a control holding a control, a block control holding a table: `is_container`);
- **a picture** and **an explicitly rich-text control** (`w:sdtPr/w:richText`) are counted as skipped, not failed;
- **a value goes into the control's first paragraph** for a block, row or cell control, keeping its `w:pPr` "and everything around it" (`ContentControl.set_bound_content`): the control's other paragraphs stay;
- a control with no type element is `"RichText"` to `sdt_kind_of` (as Office JS reports it), and `is_multi_line` is true for it, so a newline in its value becomes a `w:br`;
- the reverse writes `ContentControl.docx4j_text` since CR-013 Phase A: a break `\n`, paragraphs a `\n` apart.

### 2.2 What Word refuses (CR-013 section 10.10, the user's runs of 2026-10-23)

File 3 had a multi-line plain-text control and a block control with no type, both bound. Each variant is file 3 with one change undone:

| file | the block control | its node | Word |
|---|---|---|---|
| 3 | two paragraphs | `Note\nsecond paragraph` | could not open; repaired |
| 3a | two paragraphs | `Note` | **could not open** |
| 3b | one paragraph | `Note\nsecond paragraph` | opens |
| 3c | as the source (one paragraph) | `Note` (only the plain control changed) | opens |
| 3d | two paragraphs, `w:dataBinding` removed | `Note\nsecond paragraph` | opens |

So Word refuses **a bound block control with no type holding more than one paragraph**. The node's newline is not the cause (3b), nor the multi-line plain control's break (3c), and unbound the same two paragraphs open (3d).

### 2.3 The writes that make such a control today

On the check's source (`cr013a-SOURCE-bound-control.docx`: a block control with no type, one paragraph, bound), each write on its own (a probe, not committed):

| write | the control's content after | refused? |
|---|---|---|
| `control.insert_paragraph("X", location="End")` | `P`, `P` | no |
| `control.insert_paragraph("X", location="Start")` | `P`, `P` | no |
| `control.paragraphs[0].insert_paragraph("X", location="After")` | `P`, `P` | no |
| `control.paragraphs[0].get_range().insert_paragraph("X", location="After")` | `P`, `P` | no |
| `control.body().insert_table(1, 1, values=[["X"]])` | `P`, `Tbl` | no |
| `control.body().insert_markdown("A\n\nB")` | `P`, `P`, `P` | no |
| `control.body().insert_ooxml(<w:p>)` | `P`, `P` | no |
| `control.insert_text("X")`, and with `\r` | `P` | no (Phase A; section 11's E1 will make `\r` a paragraph) |

Every block-adding write succeeds, and each save would be one Word must repair.

### 2.4 `apply_bindings` over such a control

On file 3 (two paragraphs, node `Note\nsecond paragraph`), `apply_bindings()` writes the value into the first paragraph and keeps the second: the control reads `Note` U+000B `second paragraph` `\r` `second paragraph`, still two paragraphs, still a file Word cannot open.

### 2.5 `set_mapping` on a control that already holds more

A block control with no type holding two paragraphs, its mapping deleted and set again with `xml_mapping.set_mapping(xpath, prefix_mappings, part)`: `True`, two paragraphs, bound. The save is file 3a's shape.

### 2.6 docx4j (`30fb41024`)

- **Binding** (`bind.xslt`, the template at line 946, `w:dataBinding and not(w:richText) and not(w:docPartGallery)`): the control's `w:sdtContent` is **replaced** by one structure: a block control becomes one `w:p` (its `w:pPr` copied), a row or a table one row of one cell. So docx4j's bind leaves a bound control with one paragraph; this package's departure (section 2.1) is what keeps the second.
- **docx4j's own note** (`bind.xslt`, its XHTML template, line 333): "For all except the run-level case, we need to be able to have block level content. Which means we need to be inserting into a rich text control, which in turn means there can't be a w:sdtPr/w:dataBinding." That is Word's rule as 2.2 measured it.
- **The reverse** (`UpdateXmlFromDocumentSurface`, lines 295 to 336): `TextUtils.getText` for every bound control, except `w:richText` or the tag `od:progid=Word.Document`, whose block content is written as a flat OPC package.
- **No guard** against a second paragraph: docx4j has no content API that inserts one.

### 2.7 docx4j-core-ts

- Its session confirmed the same defect there (2026-10-23): `insertParagraph` into a bound block control with no type is not refused, and the save keeps two paragraphs.
- **Its CR-002 section 61** (`80a933d`, decided by the user there: "D1 to D6 make sense; record then commit CR-002 section 61"), not yet implemented:
  - D1: the rule covers bound block-level controls of plain text, a `w:text` element or no type element. That is one step wider than 2.2's measured trigger. `w:richText` and a node holding flat OPC are not covered.
  - D2: a write that would give one a second block is refused before anything changes, not repaired. A `\n` or `\r` written inside one becomes a `w:br` in a multi-line control and is dropped in a single-line one (check 55b's behaviour; its E1's paragraph doesn't apply there).
  - D3: the error is Word's own, as check 56 shows it, else `NotSupportedError`.
  - D4: a file already holding such a control is read and saved unchanged.
- **Check 56** (`38dcae9`, its sections 61.3 and 61.4; Word 365 16.0.20326.20232):
  - Four files made by core-ts's writers: no type with two paragraphs; `w:text` multi-line with two paragraphs; `w:richText` with two paragraphs; no type with a table before its paragraph. **All four** were unreadable to Word.
  - Every repair made the control plain text (`<w:text/>`; multi-line kept for the `w:text` one, `w:richText` replaced) and refilled it from its node, so the second block was lost. That is why file 3's repair read "Notesecond paragraph": its node held that text.
  - Office JS's block writes, 16 cases on four kinds of bound block control (no type, `w:text` multi-line, `w:text`, `w:richText`):
    - On open Word reads every bound block control as plain text. It gives the untyped and `w:richText` ones `<w:text/>` (single-line), and every lookup reports `PlainText`.
    - `insertTable` threw `InvalidArgument`, and `insertOoxml` of two paragraphs threw `GeneralException`, in every kind; nothing changed.
    - `insertParagraph('NEWPARA', 'End')`: in a single-line control the text joined the paragraph, and the node got `start1\r\nNEWPARA1`. In a multi-line control Word added a second paragraph and then couldn't open its own save.
    - `insertText('x\ny', 'End')`: a `w:br` in a multi-line control, the break dropped in a single-line one.
- **Its decisions H1 to H4** (the user there: "recommendations H1 to H4 are ok"), superseding D1:
  - H1: the rule covers every bound block-level control, untyped, `w:text` or `w:richText`. A control whose node holds flat OPC is left as it is (not measured).
  - H2: `insertParagraph` into one follows Word where Word's file is readable. In a single-line control the text joins its paragraph, and the node follows the control. In a multi-line control it is refused (`NotSupportedError`).
  - H3: Word's error codes where measured (`InvalidArgument` for `insertTable`, `GeneralException` for `insertOoxml` and `insertXml`), `NotSupportedError` for a block write Word wasn't given.
  - H4: a reader reports a bound block control with no type or `w:richText` (its node not flat OPC) as `PlainText`, single-line, as Office JS does. The file is left as it is.
- **Implemented there** (`6285dfa`, sections 60 and 61), and after the consistency review of 2026-10-23, `a18f53a` (its section 61.6) took this CR's Q4 and Q5: `setMapping` and `setMappingByNode` refuse, before anything changes, to bind a plain-text or bound block control holding more than one paragraph, and `applyBindings` over one holding more writes one paragraph, as `bind.xslt` does. So 3.3 and 3.4 are the two ports' one rule.
- Check 55b (`1e61bb1`, Word 365) measured Word's own bound controls: a line break goes to the node as `\r\n`; `\n` or `\r` written into a run-level bound plain-text control moves it to block level, in a paragraph of its own; Word's `contentControls` collections list no plain-text control (`getContentControls({ types })` finds them). It did not measure a second paragraph in a bound block control.

### 2.8 Not measured

- a plain-text (`w:text`) block control, bound or not, holding two paragraphs;
- an explicitly rich-text (`w:richText`) block control bound to a plain-text node, holding two paragraphs;
- a bound block control with no type holding a paragraph and a table;
- a bound date, drop-down or combo-box block control holding two paragraphs;
- what Word's Office JS `insertParagraph` does in a bound block control (refuses, converts the binding, or writes it).

Check 56 (2.7) answered the first three and the last. Still unmeasured: the date and list kinds, the cell and row levels, and a bound control whose node holds flat OPC.
- a cell- or row-level bound control holding more than one paragraph.

### 2.9 A bound date, as Word formats it (docx4j CR-058 phase 1, handed off 2026-10-23)

docx4j's session sent this on 2026-10-23: docx4j CR-058 phase 1, `d8d8bfb71` on `VERSION_17_3_5`, unpushed, checked in `../docx4j`. Word 365's rules for a bound `w:date` control were measured that day on docx4j-portfolio's `bind-typed-formats` probe (share `Office 2016/bind-fun`; docx4j CR-058 section 8 has the table; fixture `core-tests` `OpenDoPE/typed-controls-dates.docx`, expectations in `BoundDateFormatTest`). Each rule is set against what this package does, probed 2026-10-23 at `9de11a2` (`customxml/kinds.py`'s `format_date`, `bindings.py`'s `date_of`):

| Word 365 | here |
|---|---|
| `w:dateFormat` is Word's date picture language, a DATE field's `\@` syntax; docx4j now formats with `FormattingSwitchHelper.formatDate(wordFormat, date, lang)` | its own .NET-token formatter, separate from the field code's converter (`model/fields/formatting.py`) |
| `ddd` the weekday ("Thu"), `dddd`, `MMM`, `MMMM` | the same, in English |
| `am/pm` and `AM/PM` both upper-case "AM"/"PM" | **wrong**: `h:mm am/pm` gives `2:05 a5/p5`, `AM/PM` gives `A10/P10`; only .NET's `tt` gives "PM" |
| `h` runs 1 to 12. The probe's noon value (`12:05Z`) showed as 11:05 PM at UTC+11, where `h` and `K` agree, so this row said "Word's documentation, not Word's run"; **docx4j's next probe measured it** (`9b9d053cb`, `bind-typed-twelve`: `01:05Z` showed `12:05 PM` at UTC+11, section 10.2) | right in `format_date`. **The field code's converter maps `h` to Java's `K`** (0 to 11), which docx4j changed in the same commit (its finding 2). Probed: a DATE field's `h:mm am/pm` reads "0:05 PM" at noon and "0:05 AM" at midnight |
| names in the `w:lid` language ("22 Oktober 2026", "Okt. 2026" for de-DE) | English names only; `w:lid` ignored (CR-003 section 17) |
| no `w:lid`: Word's editing language, not the run's `w:lang`; docx4j takes the document default (`w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val`), else the platform's | `"en-US"`, unused |
| a stored value that is not a date (`not-a-date`, `2026-13-45`) is shown as stored, plain | the same (`date_of` gives None, the text as stored) |
| an empty value shows the placeholder (`w:showingPlcHdr`) | the same (docx4j's rule, CR-003 Phase E) |
| a stored dateTime's `Z` is UTC, shown in local time; docx4j keeps the nominal fields by default, and its `dateTimeZone` property converts | nominal (`Z` stripped), docx4j's default; no zone option |
| a date-only value is not shifted | not shifted |

docx4j CR-058's second item (a `w:storeItemID` that names no part, which Word repairs) wasn't handed off with phase 1, and isn't looked at here.

### 2.10 A store item ID that names no part (docx4j CR-058 phase 2, handed off 2026-10-23)

docx4j's session sent this second hand-off on 2026-10-23: docx4j CR-058 phase 2, `35cdeaee6` on `VERSION_17_3_5`, unpushed, checked in `../docx4j`. Word 365 was measured on 22 and 23 October 2026 (docx4j CR-058 section 8), on the fixtures `core-tests` `OpenDoPE/bind-plain.docx`, `bind-plain-two-parts.docx` and `bind-plain-no-part.docx`, with the expectations in `StoreItemRepairTest`. Word's rules:
- **Binding.** Where a `w:dataBinding`'s `w:storeItemID` names no custom XML part, Word evaluates the binding against the part whose root element matches the XPath's first step. The match is by namespace and local name, read with the binding's own prefix mappings.
- **Two parts.** Where two parts fit, Word takes the first in part order whose XPath selects a node: B7, with `b7` in both parts, takes part 1; B11, with `b11` only in part 2, takes part 2.
- **Saving.** Word's save writes that part's ID into the control.
- **No part.** Where no part fits (the XPath in another namespace, or a root no part has), the control keeps its surface and its stale ID, the binding untouched.

docx4j now does the same. `StoreItemResolver` is new, and every lookup goes through it: each traverser's text, date, picture, checkbox and XHTML paths, `UpdateXmlFromDocumentSurface`, and the text model's `BoundValues`, which writes nothing. The generators write the resolved ID back into the `w:dataBinding`, in memory. Where no part fits, both traversers leave the control as it is. The XSLT traverser used to empty it, and the NonXSLT one wrote `[missing!]`. A property, `strictStoreItemIds`, restores the old failure.

**Here** (read at `bfdd8be`):
- `XmlMapping.custom_xml_part` finds a part by the ID alone (`collection.get_item`).
- `apply_bindings` skips such a control with the code `no_part`, its surface and ID untouched. That is already Word's answer where no part fits, and never docx4j's old failure. But where a part fits, Word shows that part's value and this package keeps the stale surface.
- `update_from_content_controls` skips it the same way.
- The text model's `BoundValues._root` finds a part by the ID alone, so it reads the surface where Word reads the fitting part's value.

## 3. Phase 1: a bound control Word reads as plain text holds one paragraph

### 3.1 The rule

A **value control** is a block-level content control that has a `w:dataBinding` and is not a container (section 2.1): by check 56, one with no type, `w:text` or `w:richText` (H1). A control whose node holds flat OPC is left out, unmeasured (Q2). A value control's `w:sdtContent` holds **one paragraph**, and nothing else. It is multi-line only when it is `w:text` with `w:multiLine` set. A control with no type or `w:richText` is single-line, since Word gives it a `<w:text/>` on open (H4).

### 3.2 Writes that would break it are refused

Every write that would add a block to a value control's content is refused, before anything changes, tracked or not, except `insert_paragraph` into a single-line one, whose text joins the paragraph, as Word's does (H2; Q3): every row of 2.3's table that adds a block, and any other block writer the audit finds. The check sits where blocks are inserted into a `w:sdtContent` (`Body.insert_element` for a control's `body()`, a paragraph's `insert_paragraph` `"Before"` and `"After"`, `Range.insert_paragraph`, `insert_ooxml`, `insert_markdown`, `insert_table`, and CR-013 section 11's paragraph split), so that one helper answers for all. A test sweeps every block writer of the content API over a value control and asserts each refusal and the unchanged tree.

What stays allowed: text into the one paragraph (`insert_text` without a paragraph mark, `Paragraph.text = ...`, `replace_text`), formatting, `delete`, and `xml_mapping.delete()`, after which the control is free.

### 3.3 Binding a control that already holds more

`XmlMapping.set_mapping` and `set_mapping_by_node` refuse to bind a control of a value kind whose content is more than one paragraph (Q4), with the same error. The hint names the way on: reduce it to one paragraph, or leave it unbound.

### 3.4 `apply_bindings`

For a value control holding more than one block (a document made elsewhere: Word itself would not open it), `apply_bindings` writes the value into the first paragraph and removes the rest, as docx4j's `bind.xslt` replaces the content (2.6) and as Word's repair joined it (Q5). The entry in `BindingResult` says so. CR-003 section 4's departure (the first paragraph, everything around it kept) stays for a row or cell control, which docx4j would reduce to one cell.

### 3.5 What doesn't change

- `update_from_content_controls`: a node holding a newline opens (2.2's 3b), and CR-013 Phase A's characters stand (its F1 to F3 are CR-013 section 11's).
- Unbound controls: 3d opens.
- Loading and saving: a file already holding such a control is read and saved unchanged, CLAUDE.md's fidelity rule and core-ts's D4.

### 3.5a The reading (H4)

`ContentControl.type` reports a value control with no type or `w:richText` as `"PlainText"`, and `is_multi_line` as false, as Office JS does (H4; Q9). Today they read `"RichText"` and true (2.1), so `apply_bindings` turns a newline in its value into a `w:br`. Word drops the break instead. The file isn't changed by the reading.
- A row-, cell- or run-level control: Q1 measures the first two; a run-level control holds runs, and CR-013 section 11's E1, as decided, refuses a paragraph mark inside one.

### 3.6 With CR-013

- **Section 11's E1** makes `insert_text` and `replace_text` split a paragraph at `\n` and `\r`. Inside a value control it doesn't: a `\n` or `\r` becomes a `w:br` in a multi-line `w:text` control and is dropped in any other, as check 56 measured. That exception is built with CR-013's amendment (section 11.1, corrected for check 56), so E1 never gives a value control a second paragraph.
- **Phase C**'s writes across paragraphs keep the rule through the same helper.

### 3.7 The error

`BindingError` (CR-003 section 3.1), code `binding.one_paragraph` (Q6): "a content control bound to custom XML holds one paragraph: Word reads it as plain text"; hint: "write the text into its paragraph, or call xml_mapping.delete() first". A refusal reports nothing changed (`ChangeReport` and `dry_run` as every refusal).

## 4. Phases

| phase | what | status | size |
|---|---|---|---|
| 1 | Section 3: the rule (H1), the refusals and the join (3.2, 3.3), `apply_bindings` (3.4), the reading (3.5a); the tests; a Word check of `apply_bindings`' output and of the writes still allowed | **implemented 2026-11-05** (section 9; decided Q1 to Q9, 2026-11-02); Word-checked 2026-11-05 | 1 to 1½ days, with the check |
| 1b | **A bound date as Word formats it** (2.9, docx4j CR-058 phase 1, `d8d8bfb71`): the binder formats `w:dateFormat` through the field code's converter, as docx4j now does. `h` becomes 1 to 12 there (the DATE field too) and the AM/PM names upper-case; names come in the `w:lid` language, else the document default's; docx4j's `BoundDateFormatTest` (DF1 to DF13) and `DateFieldLanguageTest.twelveHourClockRunsOneToTwelve` are ported as the oracle, on docx4j's fixture | **implemented 2026-11-05** (section 10; decided Q10, 2026-10-23, amended 2026-10-28); Word-checked 2026-11-05 | ½ to 1 day |
| 1c | **A store item ID that names no part** (2.10, docx4j CR-058 phase 2, `35cdeaee6`): a port of `StoreItemResolver`, which `apply_bindings`, `update_from_content_controls` and the text model's `BoundValues` share. `apply_bindings` writes the resolved ID back, as Word's save does; `no_part` stays where no part fits. `StoreItemRepairTest` and its three fixtures are the oracle | **implemented 2026-11-05** (section 11; decided Q11, 2026-11-02); Word-checked 2026-11-05 | ½ day |
| 2 | **Rich content in bound controls**, for parity with core-ts CR-005 sections 8 and 9 and with docx4j: explicitly rich-text bindings both ways (a flat OPC package in the node, docx4j's `od:progid=Word.Document`; needs CR-002 Phase C's flat OPC sink), escaped XHTML (`od:ContentType=application/xhtml+xml`, core-ts section 8, which needs an XHTML importer here), and picture bindings (core-ts section 9: the `a:blip` replaced from base64) | named, not proposed | to be estimated |
| 3 | **The OpenDoPE processor**, core-ts CR-005 phase B and docx4j's `OpenDoPEHandler`: conditions, repeats with their contextualised XPaths, the integrity steps and removal; it needs an XPath 2 engine (core-ts uses FontoXPath; here `elementpath` would be a new dependency) | named, not proposed | weeks, as core-ts estimates its phase B |

Phases 2 and 3 are outside CR-003 section 23's Office JS parity, so each is the user's choice; the case for them is a `bind` tool for an agent (CR-006) and a template workflow from Python.

**From CR-016 (2026-10-24).**
- **Cells.** Phase 2 builds block content from XHTML and from a flat OPC package. Each table cell it builds ends with a paragraph, through `table_rules.end_with_paragraph` (CR-016 Phase A, its section 4.2).
- **Placeholders.** A bound control showing its placeholder reads as no text since CR-016 Phase B. `ContentControl.insert_text` into it clears `w:showingPlcHdr`, and the value reaches its node at the next `update_from_content_controls()` (CR-016 section 4.8 item 3). `update_from_control` still skips a control that shows its placeholder. core-ts's `insertText` syncs the bound node at once (`syncBoundNode`); here it doesn't. The difference is recorded here, for this CR to decide. **Since [CR-019](CR-019-writing-into-a-placeholder.md) (D7, 2026-10-25) it is gone:** every text write into a bound control writes its node at once, by every route, as core-ts's does.
- **An empty node** (CR-016 section 13.4 item 1, Word 365 2026-10-24). Word doesn't refresh a bound control showing its placeholder from an empty node: its save keeps `w:showingPlcHdr`, the control's own prompt run ("Click or tap here to enter text.") and the empty node. `apply_bindings` here writes docx4j's placeholder run (`runs_for_value("")`, "Click here to enter text.") in place of the control's own. That is docx4j's `BindingTraverser`, so it is parity with docx4j and not CR-016's change. Whether to keep the control's own prompt is this CR's to decide. **CR-019 decided it (its Q3, 2026-10-25, with core-ts and docx4j's CR-063):** the binding pass's empty value now shows the control's placeholder in Word's form (`show_placeholder()`): the entry the control names where the glossary holds it, else Word's default entry for its kind (`DefaultPlaceholder_-1854013440`, "Click or tap here to enter text.", for a text control; a date picker and a list have their own), written into the glossary on save. Only a kind with no default entry (a picture, a check box) keeps docx4j's run, whose words are Word 365's now (`bindings.PLACEHOLDER_TEXT`).
- **core-ts's checks 63 and 64** (its CR-002 section 63, `2070694`; CR-016 section 13.5). In the bare form (no `w:placeholder`, no glossary part, what docx4j's binding writes) Office JS's writes into a bound control never reached its node, and Word lost them on reopening. In Word's form (`w:placeholder` naming a glossary entry) every write reached it. core-ts decided to write the node on every text write into a bound control (its D7), and to write placeholders in Word's form after a check 65 (D8); docx4j's binding is docx4j's to decide. Both bear on the `syncBoundNode` difference above and on the empty-node prompt. **CR-019 ported both (2026-10-25):** every write into a bound control writes its node, in both forms (check 63's run, made again with Design Mode off, showed Word's own writes reaching the node in the bare form too, so D7 is Word's behaviour), and the binding pass and made controls write placeholders in Word's form (D8, D9; its section 12).

**From core-ts CR-005 sections 10 and 10.1 (2026-10-25, uncommitted there when noted).** For an empty node bound to a picture control, the user chose D1 (a2): Word's form (`w:placeholder` and `w:showingPlcHdr`), showing docx4j's own placeholder image, `picture-placeholder.png`. Here `apply_bindings` skips picture controls (reason `"picture"`), so nothing changes now. When picture binding is ported, under this CR, that is the rule both ports follow.

## 5. Tests and the gate (phase 1)

- **The scope** is check 56's (2.7), cited by commit; no check of this package's own comes first.
- **Unit tests** (`tests/content/test_value_controls.py`, beside `test_controls_typed.py`, which holds the binding tests today):
  - each block writer of 2.3 and the audit refused on a value control, the tree, the parts and the reports unchanged, tracked and not;
  - the writes 3.2 allows still work;
  - `set_mapping` and `set_mapping_by_node` refused on a two-paragraph control of a value kind, allowed on one paragraph;
  - `apply_bindings` over file 3's shape writes one paragraph, its `w:pPr` kept, and its entry says so;
  - an unbound control and a container unaffected.
- **The rest:** the suite, `generate.sh --check`, `roundtrip.py`, acceptance (`compare_acceptance.py` against a copy made first; artefact 9, the template filled, is the one that binds).
- **A Word check** of `apply_bindings`' output over 3a's shape and of a value control after the allowed writes, each opening without a repair prompt.

## 6. Risks

1. **Refusing too much or too little.** The rule's kinds are check 56's (H1). The date and list kinds and the cell and row levels are unmeasured, and the rule leaves a control whose node holds flat OPC alone.
2. **Callers who relied on it.** A caller inserting a paragraph into a bound control got a file Word repairs; the refusal turns a silent corruption into an error, and the hint says how to proceed.
3. **Word versions.** Measured in Word 365 only, which is CLAUDE.md's oracle.

## 7. Questions (decided 2026-11-02; Q10 2026-10-23)

The user, 2026-11-02: "CR-014's questions Q1 to Q9 and Q11 as recommended, noting for Q7 that 0.2.0 is already out". Every question below is decided as recommended. Q7 is answered by the date: 0.2.0 was released on 2026-10-25 without this CR, so phase 1 goes in the next release.

1. **Q1. Measure the rule's scope first?** *Answered by core-ts's check 56* (2.7): every bound block control, a table as a second block too. No check of this package's own is needed for phase 1. The date and list kinds and the cell and row levels stay unmeasured (2.8).
2. **Q2. Which kinds are value controls?** *Recommended: core-ts's H1, every bound block control (no type, `w:text`, `w:richText`), its node not flat OPC*, so that the two ports hold one rule.
3. **Q3. Refuse, or join?** *Recommended: as core-ts's H2, Word's own*: `insert_paragraph` into a single-line value control joins its text to the paragraph, as Word's does. Every other block write is refused (3.2), including `insert_paragraph` into a multi-line one, where Word's own result can't be opened.
4. **Q4. `set_mapping` on a control already holding more?** *Recommended: refuse*, so a binding is never made that Word cannot open. Office JS's `setMapping` there is unmeasured.
5. **Q5. `apply_bindings` over a value control holding more?** *Recommended: write one paragraph, as docx4j's `bind.xslt` does* (3.4). The alternative is to skip it with a reason and change nothing, leaving a file Word cannot open.
6. **Q6. The error.** *Recommended: `BindingError`, code `binding.one_paragraph`, its message naming Word's own error where check 56 measured one* (`InvalidArgument` for a table, `GeneralException` for OOXML; core-ts's H3). This package's codes are CR-003 section 3.1's, not Office JS's.
7. **Q7. Release.** *Recommended: in 0.2.0*, since it fixes files Word cannot open. **Decided 2026-11-02: in the next release**, since 0.2.0 is already out (released 2026-10-25). A 0.1.x fix release is the user's call, with the OLE `w:object` order fix (CR-012 section 12.15) and CR-015.
8. **Q8. Phases 2 and 3.** *Recommended: leave them named and unproposed* until there is a consumer (a `bind` tool, a template workflow); phase 2 waits on CR-002 Phase C's flat OPC sink in any case.
9. **Q9. The reading (H4).** *Recommended: as core-ts's H4*: a value control with no type or `w:richText` reads as `"PlainText"`, single-line (3.5a), so `apply_bindings` and the writers drop a newline there, as Word does. With it, `update_from_content_controls` writes a bound `w:richText` block control's text to its node, as core-ts does, where CR-003 section 4 skips it today. CR-013 section 11.2 measured the cost of the skip: a bound `w:richText` control written `NoteDxy` shows `NoteD` in Word, its node's text, because Word refreshes the control from its node on open.

10. **Q10. A bound date as Word formats it (2.9)?** *Recommended: take docx4j CR-058 phase 1 as phase 1b, with phase 1.* **Decided 2026-10-23 as recommended** (the user: "CR-014 Q10 agree"): phase 1b goes with phase 1, the names from a committed CLDR-derived table, and no `dateTimeZone` option.
    - The binder formats through `model/fields/formatting.py`'s converter, one date picture language for a control and a DATE field, as docx4j now has it.
    - `h` maps to 1 to 12 (Java's `h`, not `K`), as docx4j's next probe measured (`9b9d053cb`; section 10.2), and the AM/PM names upper-case. This moves a DATE field's text only at hours 0 and 12, and CR-011's ratchet and goldens show whether any fixture holds one.
    - The language is `w:lid`, else the document default's. The month and day names need a locale source: either CLDR data, through a new dependency such as `babel`, or a committed table generated from it, as `resources/fonts/` is from docx4j. *Recommended: a committed table* for the languages docx4j's tests and Word's measurements name, with English where a language is missing and the result's entry saying so. That keeps the package free of a dependency.
    - No `dateTimeZone` option, unless a caller asks for one: the nominal default is docx4j's.
    - **Amended 2026-10-28** by the ports' date rule (`../docx4j-portfolio/docs/date_handling_across_ports.md` section 3.6,
      point 3, the user's ruling) and CR-025 section 9: the nominal default stands, and Word's conversion is a separate
      per-package opt-in, off by default, after docx4j's `OpcPackage.setBoundDatesConvertedAsWord` (its CR-058 D6 as
      amended): `pkg.bound_dates_converted_as_word`, which reads the stored value's `Z` or offset as an instant and
      converts it into `pkg.effective_date_time_zone` (CR-025 D1), as Word does on a machine in that zone. Setting
      `pkg.date_time_zone` alone does not turn it on, and there is no process-wide fallback (CR-025 D2). Phase 1b
      builds it. Since CR-025 Phase B2 a date-only `w:fullDate` already reads as midnight (`values.date_time`), and a
      save writes a date-only or offset-free one as Word's save does.

11. **Q11. A store item ID that names no part (2.10)?** *Recommended: take docx4j CR-058 phase 2 as phase 1c, with phase 1.*
    - A port of `StoreItemResolver` (`customxml/`). The part the ID names comes first. Failing that, the part whose root matches the XPath's first step under the binding's prefix mappings; of several, the first in part order whose XPath selects a node, else the first.
    - `apply_bindings`, `update_from_content_controls` and `BoundValues` go through it. `apply_bindings` writes the resolved ID into the `w:dataBinding` and says so in its entry. `BoundValues` writes nothing.
    - **`XmlMapping.custom_xml_part` keeps naming only the part the ID names.** It is Office JS's read (`customXmlPart`), and what Office JS gives for a stale ID is unmeasured.
    - **No strict option.** This package never failed on a stale ID, so there is no old behaviour to restore.
    - The fixtures are docx4j's (Apache-2.0), copied with a SHA-1 manifest, as `samples/` holds docx4j's.

## 8. Effort

| phase | days |
|---|---|
| 1, with its own Word check (Q1's scope from core-ts's check 56) | 1 |
| 1b, a bound date as Word formats it (Q10) | ½ to 1 |
| 1c, a store item ID that names no part (Q11) | ½ |
| 2 | not estimated |
| 3 | weeks (core-ts's phase B estimate) |

## 9. Phase 1 (implementation notes, 2026-11-05)

Commits `c794583`, on `3a884a5`, and `871be28`. Section 2's measurements were taken at `c5b466f`; the code had moved under CR-013 section 11 and CR-019, so 2.3 was measured again first.

### 9.1 Section 2.3's table at `3a884a5`

The same source, each write on its own, untracked and tracked:

| write | untracked | tracked |
|---|---|---|
| `control.insert_paragraph("X", location="End")` and `"Start"` | `P`, `P` | `P`, `P` |
| `control.paragraphs[0].insert_paragraph("X", location="After")` | `P`, `P` | `P`, `P` |
| `control.paragraphs[0].get_range().insert_paragraph("X", location="After")` | `P`, `P` | `P`, `P` |
| `control.body().insert_table(1, 1, values=[["X"]])` | `P`, `Tbl` | `P`, `Tbl` |
| `control.body().insert_markdown("A\n\nB")` | `P`, `P`, `P` | `P`, `P`, `P` |
| `control.body().insert_ooxml(<w:p>)` | `P`, `P` | `P`, `P` |
| `control.insert_text("a\rb")` | `P`, `Noteab` (CR-013 section 11) | the same |
| `control.insert_text("a\nb", location="Replace")` | `P`, `ab` | **`P`, `P`**, the node `\nab` |

Every block write still went in. The last row is new: `Body.insert_text("Replace")` clears the body, and under tracking what it held stays, so the new paragraph was the control's second, by a write section 3.2 allows. Four more allowed writes reached no node (9.6).

### 9.2 What landed

| what | where | section |
|---|---|---|
| `is_bound_block_control` reads the element (`sdt_kind_of`), and `is_container` too, so `type` can ask it | `controls.py`, `bindings.py` | 3.5a, the brief's trap |
| `ContentControl.type` reads `"PlainText"`, and through it `is_multi_line` false, `controls_of_types`, `to_dict`, `repr`, `BindingEntry.kind` | `controls.py` | 3.5a |
| `describe()` from bytes reads the same (`_reads_as_plain_text`) | `template.py` | 3.5a |
| the helper, `refuse_second_block`, with `value_control_of`, `one_paragraph_error`, `other_blocks_of` and `_is_block_item` | `controls.py` | 3.2, 3.7 |
| `Body._insert_element` calls it; `Body._place_for` is where a block goes, shared with `Body.refuse_blocks`, the check for a verb with work of its own first; `insert_element(..., replacing=)` | `body.py` | 3.2 |
| the verbs with work first: `insert_table_into` (a style), `insert_picture_into_body` (an image part), `content_of(check=)` (parts copied), `blocks_for(check=)` (styles and numbering, from the parse alone), `Paragraph.refuse_beside` | `table.py`, `picture.py`, `ooxml.py`, `importer.py`, `paragraph.py` | 3.2 |
| `refuse_split`, under `split_paragraph`, `insert_paragraph_at` and `_replace` | `split.py` | 3.2, 3.6 |
| a control around a value control's paragraph: `insert_content_control_in_paragraph`, `_in_body`, `wrap_paragraphs` | `insert.py` | 3.2 |
| H2: `ContentControl._join_paragraph` | `controls.py` | 3.2, Q3 |
| `XmlMapping._write` refuses first (`_refuse_second_block`) | `mapping.py` | 3.3, Q4 |
| `set_bound_content` keeps the first paragraph alone (`other_blocks_of`, `queue_removed`); the entry says so | `controls.py`, `bindings.py` | 3.4, Q5 |
| `BindingNote`, `BindingEntry.notes`: what a write did beyond its value, also a warning on the report | `bindings.py` | 3.4, 2.9, 2.10 |
| `_skip_reason`: `w:richText` skipped only where `type` still reads `"RichText"` | `bindings.py` | Q9 |
| `Body.insert_text("Replace")` in a value control writes in place of its text | `body.py` | 3.2 |
| `Range._replace`: a break as the control takes it, across paragraphs too | `range.py` | 3.6 |

**The refusal.** `BindingError`, `binding.one_paragraph`. Its message is 3.7's after the verb (`insert_table refused: a content control bound to custom XML holds one paragraph: Word reads it as plain text; Word's own answer here is InvalidArgument (core-ts check 56)`), and its hint 3.7's. `InvalidArgument` is named for `insert_table`, `GeneralException` for `insert_ooxml` and `insert_xml`, at every level; any other write names none.

**What `set_mapping` tests.** The raw kind and `holds_one_paragraph`, not `is_bound_block_control`, since the control isn't bound yet, and only where it would write `w:dataBinding`: a control already holding a `w15:dataBinding` keeps that spelling (`_write`) and may hold blocks. It refuses only once a part resolves the XPath, so a path selecting nothing still answers False.

### 9.3 The audit: every block writer

`SWEEP` in `tests/content/test_value_controls.py` holds each of these on each of check 56's four kinds, tracked and not, with the document, the part names, the data, the styles and the report count unchanged.

| module | block writers | how each is held |
|---|---|---|
| `controls.py` | `insert_paragraph` | H2 at `"Start"`, `"End"`; `"Before"`, `"After"` go outside, allowed |
| `paragraph.py` | `insert_paragraph`, `insert_break` (`"Before"`, `"After"`), `insert_xml`, `insert_ooxml`, `insert_markdown`, `insert_content_control` | the helper; one `w:p` at `"Start"` or `"End"` merges and is allowed; `"Replace"` by one `w:p` allowed untracked |
| `range.py` | `insert_paragraph`, `insert_ooxml`, `insert_content_control` across paragraphs | the helper; one `w:p` merges |
| `body.py` | `insert_paragraph`, `insert_break`, `insert_table`, `insert_xml`, `insert_ooxml` (`"Replace"` too), `insert_markdown`, `insert_inline_picture`, `insert_element`, `insert_content_control`, `add_object`, `add_paragraph_of_text`, `add_styled_paragraph_of_text` | the helper |
| `span_edits.py` | `insert_paragraph_at`, `place_replacement` | through `refuse_split`, `Paragraph.insert_paragraph` and `insert_element`; `place_replacement` reaches a value control with its breaks already taken out (`_replace`). `keep_cell_mark_last` is a cell's. `delete_span` joins |
| `split.py` | `split_paragraph` | `refuse_split` |
| `lists.py` | `List.insert_paragraph` | through `Paragraph.insert_paragraph` |
| `table.py` | `TableCell`'s, `Table.insert_paragraph` | a table in a control makes it a container, so none reach one |
| `docx/` | `Paragraph.insert_paragraph_before`; `add_paragraph`, `add_table`, `add_picture`, `add_heading` on a `BlockItemContainer` | through the content API; the facade's containers are the body, a header, a footer, a cell and a comment, never a control, so only the first can reach one. The direct `content.append`s are a run's, a row's or the body's own section paragraph |
| `customxml/template.py` | none: `fill` writes nodes and calls `apply_bindings` | `_kind_of` read for 3.5a |
| `customxml/insert.py` | the three block forms of `insert_content_control` | the helper, before the control takes anything; `insert_content_control_in_range` is run-level (13.3, O4) |
| `kinds.py`, `insert_copy_after` | a repeating section's item, a copy of a control | beside a control, not in one |

### 9.4 The writes that stay allowed

`ALLOWED` holds fifteen writes on the four kinds, tracked and not: one paragraph, the text as written, and the node equal to it. `Paragraph.delete` and `Body.clear` are allowed, and leave the node (13.3, O2).

### 9.5 `apply_bindings` over more

File 3's shape in each of the four kinds gives one paragraph, its `w:pPr` kept. The extra blocks go through `queue_removed`, and what they held goes when the pass ends, a comment whose reference went with it: `recording_on`, which `apply_bindings` runs under, now settles a write's removals as `recording` always did. A marker beside the paragraph stays. An empty value in a value control holding more shows its placeholder in one paragraph.

### 9.6 Found, and fixed here

Each is a write section 3.2 allows that Word would have undone, since it refreshes a bound control from its node on open:

1. **`Paragraph.text = ...` untracked** wrote no node (it is no splice); tracked it did.
2. **One `w:p` merged** by `Paragraph.insert_xml`, `insert_ooxml` and `insert_markdown` wrote no node.
3. **`Range.insert_ooxml` of one `w:p`** left the node empty: the removal wrote it, the merge didn't. `Paragraph.write_bound_nodes` is the node write for 2 and 3, and CR-019's D7 is now true of them.
4. **A tracked `control.insert_text(..., "Replace")`** gave the control a second paragraph (9.1).

### 9.7 Pre-existing tests whose expectation changed, and why

- `test_word_writers.py`: `test_a_bound_block_control_holds_one_paragraph`'s reader line, which said "the reader is CR-014's": the bound rows now read `PlainText`. `CHECK_56_NODES` loses case 15, which pinned Q9's skip: the node is now Word's, `start15xy`.
- `test_office_js_characters.py`: `test_a_bound_controls_write_back_keeps_docx4js_characters` bound a two-paragraph control with `set_mapping`, now refused; it binds with `put_property`, as a file made elsewhere holds it.

## 10. Phase 1b (implementation notes, 2026-11-05)

Commit `ce7f5c1`.

### 10.1 What landed

- **`format_word_date(picture, when, lang)`** (`fields/formatting.py`), docx4j's `formatDate(String, Date, String)`, with `uses_date_names`. `kinds.format_date` keeps its name and signature, over it: one picture language for a control and a DATE field. The .NET token table is gone.
- **`h` and `hh` are 1 to 12** in `_DATE_ITEMS`, for a DATE field too. `tests/text/test_fields.py`'s pinned `K:mm` changed with it.
- **`resources/date-names.json`**, written by `scripts/generate_date_names.py` through Babel 2.18.0, CLDR 47, and never hand-edited. Babel is the script's alone (`python3 scripts/generate_date_names.py`, `--check`): no dependency of the package, and not in `.venv-fork`. The names are CLDR's format names, as they stand; docx4j's two rules, a short month's trailing period off and AM and PM upper-cased, are `model/fields/date_names.py`'s reader. The table's strings were compared with `java.text.DateFormatSymbols` on OpenJDK 21: the same for every locale in it. `hatchling` packs everything git tracks under `docx4j_py/`, so it ships; the wheel built from this tree holds it (13.1).
- **The languages** are Q10's, those docx4j's tests and Word's measurements name: `en`, `en-AU`, `de`, `fr`, `tr`. A tag is read whole, then by its language.
- **`simple_date_format(pattern, when, names=None)`**: a language's names; and only an ASCII letter is a pattern letter now, as in Java, so `yyyy年M月d日` works.
- **`date_of(value, zone=None)`** is docx4j's `parseStoredDate`, both forms: `xs:dateTime` and `xs:date`, the nominal fields, an offset read past; with a zone a date and time carrying a `Z` or offset is an instant, shown there.
- **`document_language(package)`**: `w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val`, read with lxml from `part.xml`, once a pass, so that binding a date leaves a styles part nobody read as it is.
- **`pkg.bound_dates_converted_as_word`**, False, on `OpcPackage` and the trial, and `dates.bound_date_zone(package)`, docx4j's `conversionZone`: None unless the package opts in, then `effective_zone`. `pkg.date_time_zone` alone turns nothing on. No process-wide setting (CR-025 D2).
- **`BindingNote`s**: `binding.date_language_missing` where a language the table lacks meets a picture that writes a name (a picture of numbers says nothing), and `binding.date_picture_unsupported`.

### 10.2 The oracle, and the record corrected

`tests/content/test_bound_dates.py` ports `BoundDateFormatTest`, at docx4j's current state on `VERSION_17_3_6`, on its fixtures: DF1 to DF13, the document default language, the opt-in, and `twelveOClockAndOffsetsAsWordShowsThem` in both readings, with `typed-controls-twelve.docx`, new since the CR's hand-off. `twelveHourClockRunsOneToTwelve` is ported whole, and `DateFieldLanguageTest`'s two language cases.

**Section 2.9's "Word's documentation, not Word's run" is superseded.** docx4j's `9b9d053cb` measured it: `bind-typed-twelve`, Word 365 at UTC+11, TW1 `2026-10-22T01:05:00Z` under `h:mm AM/PM` showed `12:05 PM`, TW2 `13:05Z` under `hh:mm am/pm` `12:05 AM`; an explicit offset is honoured and no offset is local. So `h` is 1 to 12 on Word's run, and the table above and Q10 now say so.

**CR-011's ratchet and the goldens: nothing moved.** The change is live only at hours 0 and 12. `tests/text/test_conformance.py` and the field probe's test pass as they stood, and the committed report is unchanged.

## 11. Phase 1c (implementation notes, 2026-11-05)

Commit `d0fe5e8`.

`model/customxml/store_items.py` is docx4j's `StoreItemResolver`: `first_step`, `resolve`, `resolved_store_item_id`, `repair` as `_repair`. The parts considered are `collection.items` less the property stores (`built_in`), docx4j's `instanceof CustomXmlDataStoragePart`. Part order is docx4j's, the part name compared as text, so `item10` comes before `item2`; Word's rule is measured on two parts only.

`XmlMapping` gains **`resolved_part`** and **`resolved_node`**, which `apply_binding_to`, `update_from_control`, `update_from_content_controls`, `write_control_to_node` and `_node_holds_flat_opc` read. `custom_xml_part` and `custom_xml_node` stay strict (Q11), and so does `to_dict()["resolves"]`. `BoundValues` asks `_repaired_root` where `_root` finds nothing, keyed by ID and XPath, since one stale ID fits two parts (B7 and B11).

`apply_bindings` writes `resolved_store_item_id`, the part's ID upper-cased, into the `w:dataBinding` or `w15:dataBinding`, in memory, and says so (`binding.store_item_repaired`, naming both IDs and the part, also a warning on the report). `BindingEntry.store_item_id` is then the new ID. `update_from_content_controls` writes the fitting part's node and leaves the ID: that direction only reads the document, as docx4j's.

`tests/content/test_store_item_repair.py` ports `StoreItemRepairTest`, but for `theStrictPropertyFailsTheBind` (no strict option, Q11), on its three fixtures.

`test_controls_typed.py`'s `test_a_binding_whose_part_is_gone_is_reported_not_raised` had a part that fits, so its stale ID is now repaired; its XPath names a root no part has.

## 12. Departures, with why

From core-ts (`6285dfa`, `a18f53a`):

1. **A container is no value control** (3.1). core-ts's `isBoundBlockControl` doesn't test it, so a bound block control holding a table or a control reads `PlainText` there and refuses writes; here it reads `RichText` and is left, as `apply_bindings` leaves it in both. The file is one Word can't open whichever.
2. **A marker is no block** (`holds_one_paragraph`, CR-019 section 4.5). core-ts's `BLOCK_LEVEL_TYPES` holds bookmarks and the rest, so a bookmark beside the paragraph refuses there, and is removed by its `applyBindings`.
3. **`set_mapping` on an unbound untyped or `w:richText` control holding more refuses** (3.3). core-ts's `usesW15` gives one `w15:dataBinding`, which may hold blocks; this package writes `w:dataBinding` there, as CR-003 built it, so 3.3's refusal is what keeps the file readable.
4. **One error** (Q6): `BindingError` with Word's name in its message, where core-ts throws `InvalidArgument`, `GeneralException` or `NotSupportedError`.
5. **H2 refuses first, and a placeholder goes by `_write_placeholder`.** core-ts takes the placeholder's runs before its multi-line refusal, so that refusal changes the control; and its text takes a plain run. Here nothing changes, and the text takes the control's look (CR-019's D11).
6. **`"Replace"` of a value control writes in place, tracked or not.** core-ts's `Body.insertText` clears and inserts. Here the paragraph and its `w:pPr` stay, where 0.2.0 made a new one; tracked it was 9.1's second paragraph.
7. **A `\n` in a Replace across a control holding two follows the control** (`Range._replace`), where core-ts's split refuses it. The deletion joins the paragraphs, so the write ends in one.
8. **`Paragraph.text = ...` and the merges write the node** (9.6). core-ts's untracked setter doesn't either; it is CR-019's D7 applied.
9. **A tracked `"Replace"` of the paragraph by one `w:p` is refused** (`Paragraph.replaced_block`, `Body.refuse_blocks`): the paragraph stays deleted beside the new one, two paragraphs.
10. **`insert_markdown` counts from the parse**, to refuse before a style is defined: one paragraph, heading or code block is one; a list, a quote, a table or anything else is more, a one-item list included. core-ts has no such writer.

From docx4j (`VERSION_17_3_6`):

11. **No `w:lid` and no document default is English**, DF7, where docx4j takes the JVM's locale. A process's locale is no document's; `DEFAULT_DATE_LANGUAGE`, what this package always used.
12. **A picture this can't write shows the stored text with a note**, where docx4j's traverser writes `[missing!]`.
13. **A language the table lacks is English with a note**, where docx4j has the JVM's data. Q10.
14. **`tt` and text in double quotes stay readable** through `kinds.format_date`, respelt to `am/pm` and single quotes before the converter: 0.2.0 read them, Java's formatter throws on them, and Word is unmeasured (O7). `h:mm tt` keeps `2:05 PM`.
15. **`w:fullDate` is set from the value**, CR-003 section 4's departure, where docx4j's bind leaves it; a value that is no date leaves it, where Word's save drops it.
16. **A stored value with a space before its time is no date now**, as in docx4j and `xs:dateTime`. `date_of` read it until phase 1b.
17. **A DATE field's language stays US English** (CR-025 D8): docx4j's field path takes one, and phase 1b gives the bound date alone its language.
18. **No `strictStoreItemIds`, no memo per document** (Q11), and the well-known IDs answer None in `resolve`: docx4j reaches its resolver past them, and this package special-cases none of them elsewhere (CR-003 section 4).
19. **The ID is written back only for a control the pass applied.** A skipped one keeps it: a picture, a container, a `w:richText` one at run level. docx4j's NonXSLT traverser writes it for a picture and its XSLT one doesn't.
20. **A part that fits and selects nothing is `no_match`, the control and its ID left.** Word's no-part probes have no root match, so this shape is unmeasured.

From core-ts again (found by the Word check, section 13.2):

21. **A placeholder write makes no styles part.** Where a package has no `word/styles.xml`, `require_placeholder_style` leaves it as it is, so the placeholder isn't grey in Word. core-ts's `requireStyles` makes the part on save, with its whole default style set and `PlaceholderText` (measured there 2026-11-05 at `b07134f`; that file not opened in Word). Kept (the user, 2026-11-05: "decision stands"). The portfolio's `docs/differences_between_implementations.md` holds it as DI-001.

## 13. The gates, the Word check and what's open

### 13.1 The gates

At `871be28`, on Python 3.14.7, `out/acceptance-before-cr014/` built at `3a884a5` first:

| gate | result |
|---|---|
| the full suite | 10,619 tests, none failed (9,907 at `3a884a5`). New, 712: `test_value_controls.py` 589, the sweep 45 writers on 4 kinds, tracked and not, 360, and the allowed writes 15 on 4 kinds, tracked and not, 120; `test_bound_dates.py` 77; `test_store_item_repair.py` 40; `test_fields.py` 6 |
| `codegen/generate.sh --check` | two generations of `docx4j_py` byte identical, 320 modules; nothing generated moved, and `KEEP` keeps the two new modules |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | `ROUND TRIP CLEAN: 50/50 parts canonically identical, 0 differences, 0 skipped (34 benign respellings)` |
| `scripts/acceptance.py`, `scripts/compare_acceptance.py out/acceptance-before-cr014 out/acceptance` | 17 artefacts, 246 entries, 0 differ beyond the wall-clock dates, artefact 9, the filled template, among them |
| CR-011's ratchet and goldens | as they stood (section 10.2) |
| the wheel | built with `hatchling` in a scratch environment: `resources/date-names.json`, `model/fields/date_names.py` and `model/customxml/store_items.py` are in it. `scripts/release_check.sh` is the release's, and wasn't run |

Python 3.12 and 3.13 are CI's; this ran on 3.14 alone.

### 13.2 The Word check (run 2026-11-05: passed)

`out/cr014-word-check/`, built by its `build.py` (`BUILD-LOG.txt`, `SHA1SUMS`), each file read back strictly with nothing skipped. `README.txt` says per file what to open and what Word should show, and where:

| file | what |
|---|---|
| `cr014-1-bind-over-two-paragraphs.docx`, `cr014-1b-3a-shape-bound.docx` | `apply_bindings`' output over 3a's shape, in each kind, and over CR-013's own source |
| `cr014-2-allowed-writes.docx`, `cr014-2b-allowed-writes-tracked.docx` | a value control after each allowed write: fifteen cases, untracked and tracked |
| `cr014-3-dates-nominal.docx` to `cr014-3f-twelve-as-written-unbound.docx` | bound dates, nominal and as Word at UTC+11, and two copies unbound so that Word shows what was written |
| `cr014-4-store-item-repaired.docx`, `cr014-4b-two-parts.docx`, `cr014-4c-no-part.docx` | a repaired store item ID, two parts, and none |

**The run** (the user, Word 365, 2026-11-05; the files on the share, `__tmp/cr014-word-check/`, `results.txt`). All ten steps passed: every file opened without a repair prompt, and each control showed what `README.txt` lists, but for the three notes below.

Two things neither port had measured:
- **Case 13 (step 3).** Word keeps the bold and the centring of a bound control's content through its refresh from the node.
- **Tracked writes (step 4).** Each of the fifteen cases showed its markup as listed, so Word's refresh on open leaves a bound control's revisions in place.

The notes, none a fault of the output:
- **`cr014-3-dates-nominal.docx`, DF2 to DF4.** Reported against the table's "cr014-3 wrote" column, because Word showed `2:05 AM`, `11:05 PM` and `02:05`. The file holds `3:05 PM`, `12:05 PM` and `15:05` (read from its `document.xml`, the SHA-1 as staged), and Word replaces them from the node on open, as step 5 says. So a bound file can't show that column in Word: `README.txt`'s table misled here. Only the unbound `cr014-3e` and `cr014-3f` can, and both passed step 7.
- **A placeholder not grey: DF10 in `cr014-3e`, B5 in `cr014-4-store-item-repaired.docx`.** `README.txt` was wrong to promise grey. The nine files built from docx4j's fixtures have no `word/styles.xml`, so `w:rStyle` names `PlaceholderText` in a file that doesn't define it, and `require_placeholder_style` leaves a package with no styles part as it is (CR-019 section 12). The text and bindings are right. **Kept as it is** (the user, 2026-11-05: "grey behaviour ok as is"): a placeholder write makes no styles part (departure 21, where core-ts does).
- **Case 15, tracked (step 4).** The inserted pilcrow shows before "second paragraph 15". That is the file's markup and Word's own rule at the end of a container: the preceding paragraph's mark is the inserted one, the final mark left alone (CR-003 section 16.10).

### 13.3 Open points

Each is unmeasured or undecided; none blocks the phases.

- **O1. `describe()` and `fill()` stay strict.** Q11 names three readers and they aren't among them, so a stale binding reads no value, and `fill` of its key is skipped `no_part` while the pass after repairs it. Going through the resolver is two lines.
- **O2. `Paragraph.delete` and `Body.clear` leave the node.** A value control emptied that way refills from its node in Word. core-ts's do the same. A `"Content"` range's delete writes it.
- **O3. `Paragraph.insert_break("Line" or "Page")` at `"Start"` or `"End"`** of a single-line value control writes the `w:br`, and no node. It is no block and no text writer; F3 drops such a break from text. Office JS's `insertBreak` there is unmeasured.
- **O4. A run-level control made inside a value control's paragraph** isn't refused: it holds no block. It makes the outer control a container. Word's answer is unmeasured.
- **O5. `Paragraph.insert_inline_picture`** into a value control's paragraph isn't refused, for the same reason.
- **O6. Word binds a run-level control the same way** (core-ts section 61.7: B8's `PlainText`, RX1's bind). Section 3.5a is block-level, as core-ts's H4, so a run-level `w:richText` control stays skipped and an untyped one `RichText`, both ports alike.
- **O7. `tt` and double quotes** in a `w:dateFormat`: departure 14, Word unmeasured.
- **O8. The names are CLDR's, not Word's, in places**, as docx4j's: German `Sept` where Word writes `Sep`, `en-AU`'s `Sept`, `June` and `July`, a German weekday's `Do.`. And five locales; another is one line of the script.
- **O9. The date and list kinds, and the row and cell forms,** stay outside the rule: 2.8's unmeasured.
- **O10. `update_from_content_controls` drops a date node's `Z`** (`iso_of`, since CR-003): `2026-10-22T15:05:00Z` goes back as `2026-10-22T15:05:00`, which Word reads as local. Not changed by this CR.
