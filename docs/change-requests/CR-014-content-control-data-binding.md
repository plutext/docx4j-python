# CR-014: Content control data binding --- a bound control Word reads as plain text holds one paragraph (phase 1); rich text, XHTML and picture bindings and the OpenDoPE processor named for later

**Status:** Proposed 2026-10-23, at the user's request ("draft the new CR here for content control data binding with phase 1 only proposed for now"), after CR-013 Phase A's Word check found a file this package writes that Word cannot open (CR-013 section 10.10): a block content control that has no type and a `w:dataBinding` and holds two paragraphs. 0.1.0 writes it too. Data binding here has been CR-003 Phase E's (sections 3.7, 4 and 17), and no CR has owned it since; this one does. Phase 1 (sections 3 and 4) is proposed. core-ts's Word check 56 (`38dcae9`) measured the rule's scope for both ports: Word can't open **any** bound block control holding a second block, whether it has no type, `w:text` or `w:richText`, and whether the second block is a paragraph or a table. On open Word reads each as plain text. So such a control keeps one paragraph: the writes that would give it a second are refused, or joined where Word joins them, `set_mapping` refuses a control that already holds more, and `apply_bindings` writes one paragraph, as docx4j's `bind.xslt` does. Phases 2 (rich-text, XHTML and picture bindings) and 3 (the OpenDoPE processor) are named in section 4 and not proposed. Nothing is implemented; phase 1 waits on the user's answers (section 7) and go. Phase 1b (section 2.9, Q10), a bound date formatted as Word formats it, was proposed the same day from docx4j's hand-off of its CR-058 phase 1 (`d8d8bfb71`). A probe here found `am/pm` in a `w:dateFormat` written as `a5/p5`, names in English only, and the DATE field converter's `h` running 0 to 11. Q10 was decided the same day as recommended (the user: "CR-014 Q10 agree"): phase 1b goes with phase 1, the names from a committed CLDR-derived table. Q1 to Q9 are open. Phase 1c (section 2.10, Q11), a store item ID that names no part resolved as Word repairs it, was proposed the same day from docx4j's second hand-off (`35cdeaee6`). Q11 is open.

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
| `h` runs 1 to 12: **Word's documentation, not Word's run**. The probe's noon value (`12:05Z`) showed as 11:05 PM at UTC+11, where `h` and `K` agree; docx4j CR-058 sections 1 and 8 say so, and docx4j's next probe takes a 12:xx local value | right in `format_date`. **The field code's converter maps `h` to Java's `K`** (0 to 11), which docx4j changed in the same commit (its finding 2). Probed: a DATE field's `h:mm am/pm` reads "0:05 PM" at noon and "0:05 AM" at midnight |
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
| 1 | Section 3: the rule (H1), the refusals and the join (3.2, 3.3), `apply_bindings` (3.4), the reading (3.5a); the tests; a Word check of `apply_bindings`' output and of the writes still allowed | **proposed** | 1 to 1½ days, with the check |
| 1b | **A bound date as Word formats it** (2.9, docx4j CR-058 phase 1, `d8d8bfb71`): the binder formats `w:dateFormat` through the field code's converter, as docx4j now does. `h` becomes 1 to 12 there (the DATE field too) and the AM/PM names upper-case; names come in the `w:lid` language, else the document default's; docx4j's `BoundDateFormatTest` (DF1 to DF13) and `DateFieldLanguageTest.twelveHourClockRunsOneToTwelve` are ported as the oracle, on docx4j's fixture | **decided** (Q10, 2026-10-23); with phase 1 | ½ to 1 day |
| 1c | **A store item ID that names no part** (2.10, docx4j CR-058 phase 2, `35cdeaee6`): a port of `StoreItemResolver`, which `apply_bindings`, `update_from_content_controls` and the text model's `BoundValues` share. `apply_bindings` writes the resolved ID back, as Word's save does; `no_part` stays where no part fits. `StoreItemRepairTest` and its three fixtures are the oracle | **proposed** (Q11) | ½ day |
| 2 | **Rich content in bound controls**, for parity with core-ts CR-005 sections 8 and 9 and with docx4j: explicitly rich-text bindings both ways (a flat OPC package in the node, docx4j's `od:progid=Word.Document`; needs CR-002 Phase C's flat OPC sink), escaped XHTML (`od:ContentType=application/xhtml+xml`, core-ts section 8, which needs an XHTML importer here), and picture bindings (core-ts section 9: the `a:blip` replaced from base64) | named, not proposed | to be estimated |
| 3 | **The OpenDoPE processor**, core-ts CR-005 phase B and docx4j's `OpenDoPEHandler`: conditions, repeats with their contextualised XPaths, the integrity steps and removal; it needs an XPath 2 engine (core-ts uses FontoXPath; here `elementpath` would be a new dependency) | named, not proposed | weeks, as core-ts estimates its phase B |

Phases 2 and 3 are outside CR-003 section 23's Office JS parity, so each is the user's choice; the case for them is a `bind` tool for an agent (CR-006) and a template workflow from Python.

**From CR-016 (2026-10-24).**
- **Cells.** Phase 2 builds block content from XHTML and from a flat OPC package. Each table cell it builds ends with a paragraph, through `table_rules.end_with_paragraph` (CR-016 Phase A, its section 4.2).
- **Placeholders.** A bound control showing its placeholder reads as no text since CR-016 Phase B. `ContentControl.insert_text` into it clears `w:showingPlcHdr`, and the value reaches its node at the next `update_from_content_controls()` (CR-016 section 4.8 item 3). `update_from_control` still skips a control that shows its placeholder. core-ts's `insertText` syncs the bound node at once (`syncBoundNode`); here it doesn't. The difference is recorded here, for this CR to decide.

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

## 7. Questions (for the user's decision; each with a recommendation)

1. **Q1. Measure the rule's scope first?** *Answered by core-ts's check 56* (2.7): every bound block control, a table as a second block too. No check of this package's own is needed for phase 1. The date and list kinds and the cell and row levels stay unmeasured (2.8).
2. **Q2. Which kinds are value controls?** *Recommended: core-ts's H1, every bound block control (no type, `w:text`, `w:richText`), its node not flat OPC*, so that the two ports hold one rule.
3. **Q3. Refuse, or join?** *Recommended: as core-ts's H2, Word's own*: `insert_paragraph` into a single-line value control joins its text to the paragraph, as Word's does. Every other block write is refused (3.2), including `insert_paragraph` into a multi-line one, where Word's own result can't be opened.
4. **Q4. `set_mapping` on a control already holding more?** *Recommended: refuse*, so a binding is never made that Word cannot open. Office JS's `setMapping` there is unmeasured.
5. **Q5. `apply_bindings` over a value control holding more?** *Recommended: write one paragraph, as docx4j's `bind.xslt` does* (3.4). The alternative is to skip it with a reason and change nothing, leaving a file Word cannot open.
6. **Q6. The error.** *Recommended: `BindingError`, code `binding.one_paragraph`, its message naming Word's own error where check 56 measured one* (`InvalidArgument` for a table, `GeneralException` for OOXML; core-ts's H3). This package's codes are CR-003 section 3.1's, not Office JS's.
7. **Q7. Release.** *Recommended: in 0.2.0*, since it fixes files Word cannot open. A 0.1.x fix release is the user's call, with the OLE `w:object` order fix (CR-012 section 12.15) and CR-015.
8. **Q8. Phases 2 and 3.** *Recommended: leave them named and unproposed* until there is a consumer (a `bind` tool, a template workflow); phase 2 waits on CR-002 Phase C's flat OPC sink in any case.
9. **Q9. The reading (H4).** *Recommended: as core-ts's H4*: a value control with no type or `w:richText` reads as `"PlainText"`, single-line (3.5a), so `apply_bindings` and the writers drop a newline there, as Word does. With it, `update_from_content_controls` writes a bound `w:richText` block control's text to its node, as core-ts does, where CR-003 section 4 skips it today. CR-013 section 11.2 measured the cost of the skip: a bound `w:richText` control written `NoteDxy` shows `NoteD` in Word, its node's text, because Word refreshes the control from its node on open.

10. **Q10. A bound date as Word formats it (2.9)?** *Recommended: take docx4j CR-058 phase 1 as phase 1b, with phase 1.* **Decided 2026-10-23 as recommended** (the user: "CR-014 Q10 agree"): phase 1b goes with phase 1, the names from a committed CLDR-derived table, and no `dateTimeZone` option.
    - The binder formats through `model/fields/formatting.py`'s converter, one date picture language for a control and a DATE field, as docx4j now has it.
    - `h` maps to 1 to 12 (Java's `h`, not `K`), on Word's documentation until docx4j's next probe measures a 12:xx local value, and the AM/PM names upper-case. This moves a DATE field's text only at hours 0 and 12, and CR-011's ratchet and goldens show whether any fixture holds one.
    - The language is `w:lid`, else the document default's. The month and day names need a locale source: either CLDR data, through a new dependency such as `babel`, or a committed table generated from it, as `resources/fonts/` is from docx4j. *Recommended: a committed table* for the languages docx4j's tests and Word's measurements name, with English where a language is missing and the result's entry saying so. That keeps the package free of a dependency.
    - No `dateTimeZone` option, unless a caller asks for one: the nominal default is docx4j's.

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
