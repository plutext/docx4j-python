# Tests

```bash
.venv-fork/bin/python -m pytest                    # everything, 1,205 tests
.venv-fork/bin/python -m pytest -m "not slow"      # without the corpus round trip and the timings
.venv-fork/bin/python -m pytest tests/openpackaging # the engine, CR-002
.venv-fork/bin/python -m pytest tests/content      # the views, CR-003 Phase B
.venv-fork/bin/python -m pytest tests/agent        # the agent surface, CR-003 Phase D
```

| file | CR | what it pins |
|---|---|---|
| `test_child.py` | CR-001 §5 | parent pointers, `ChildList`, `deep_copy` |
| `test_names.py` | CR-001 §6 | docx4j's class names, and that renaming never touched `Meta.name` |
| `test_skipped_report.py` | CR-001 §7 | the fork's skipped-content report catches an unknown element and attribute |
| `test_roundtrip.py` | CR-001 §9 | `scripts/roundtrip.py` over `samples/`, zero differences |
| `test_performance.py` | CR-001 §8 | model import time, parse throughput, the `link_parents` budget |
| `test_el.py`, `test_codegen_el.py` | CR-001 §6.2 | the `el` factory and its reviewed tables |
| `test_builders.py`, `test_fragments.py` | CR-001 §6.2 | the `p`/`r`/`t`/`tbl` sugar, `wml(...)` and `to_xml` |
| `test_traversal.py` | CR-001 §6.2 | `walk`, `find`, `text_of` against an independent lxml extraction |
| `test_runtime.py`, `test_threads.py` | CR-001 §8 | the prefix table, `warm_up`, a shared `XmlContext` |
| `openpackaging/test_part_name.py` | CR-002 §5.1 | OPC part names: the conformance clauses, resolve/relativize, case-insensitivity |
| `openpackaging/test_content_types.py` | CR-002 §5.3, §5.4 | `[Content_Types].xml` and docx4j's `newPartForContentType` table |
| `openpackaging/test_round_trip.py` | CR-002 §1, §8 | load / save / reload over every sample, byte-identical and canonically identical |
| `openpackaging/test_parts_and_relationships.py` | CR-002 §5.2, §5.5 | the relationship graph, `AddPartBehaviour`, `create_package`, `mc:Ignorable` on write |
| `openpackaging/test_mce.py` | CR-002 §5.6 | `mce="resolve"` / `"all"` / `"none"`, and the preprocessor |
| `openpackaging/test_threads_and_import.py` | CR-002 §3, §7 | eight threads over whole packages, and the engine's import cost |
| `content/test_body.py`, `test_paragraph.py`, `test_range.py`, `test_font.py` | CR-003 §3.2 | the views: the verbs, the style semantics, search across runs |
| `content/test_text_model.py` | CR-003 §3.12 | segments, grapheme-safe splitting, the search options |
| `content/test_errors_and_parts.py` | CR-003 §3.1 | the error hierarchy, and untouched parts still byte-identical |
| `content/test_table.py` | CR-003 §3.2, §4 | `Table`, `TableRow`, `TableCell`: `insert_table` sized from `w:sectPr`, `values`, `add_rows` / `insert_rows` / `delete_rows`, `header_row_count` on leading rows only, a nested table, and rows that descend into an OpenDoPE row-level `w:sdt` |
| `content/test_picture.py` | CR-003 §3.2, §4 | `InlinePicture`: the free `/word/media/imageN.<ext>`, the relationship from the body's own part, the header readers, points in and out, `delete`, and a `dry_run` that un-adds the part |
| `content/test_ooxml.py` | CR-003 §3.2, §4 | `insert_ooxml` from a flat OPC `pkg:package`: parts copied under free names with fresh relationship ids, numeric ids untouched, styles not merged; and `FlatOpcStore` |
| `content/test_controls.py` | CR-003 §3.2, §4 | `ContentControl` over all four `w:sdt` forms: the reads, `get_range()` exact for a run control, the row and cell refusals, `delete(keep_content=)` |
| `content/test_customxml.py` | CR-003 §3.7 | the custom XML model: the collection and `get_item` brace- and case-insensitive, the node model over lxml, XPath with Word's `xmlns:ns0='…'` string and the canonical positional path, the part-level XPath editors, `add()` with its properties part and the relationship from the main document part, `delete()` unlinking the mappings it can reach, a `dry_run` that un-adds a part, and **the promise this phase exists for**: reading every part and node leaves `customXml/item1.xml`, `customXml/itemProps1.xml` and `word/document.xml` byte for byte, while one node mutation marks that part and no other |
| `content/test_controls_typed.py` | CR-003 §3.7, §4 | the seven typed kinds (checkbox, date picker, drop-down, combo box, picture, repeating section, group --- the last three of the first four on documents this engine builds, since no fixture has them), the `w:sdtPr` properties and the two halves of `w:lock`, `placeholder_text` refused over content, `XmlMapping` reading `w:dataBinding` **and** `w15:dataBinding` and resolving before it writes, `insert_content_control` at all three levels with the two refusals, `apply_bindings` / `update_from_content_controls` in both directions for text, checkbox, date and list, and the write-through on `insert_text` with tracking off **and on** |
| `content/test_comments.py` | CR-003 §3.9, §4 | comments read from three documents (one Word wrote before `w14:paraId`, one Word wrote with all five parts, one thread built here); `insert_comment` on a range and on a paragraph, `reply`, `resolved`, `content`, `delete`; the markers hoisted out of a `w:ins`; a span across a run holder refused; the byte-for-byte promises for the two side parts and for `styles.xml` |
| `content/test_tracking.py` | CR-003 §3.8, §4 | change tracking: the mode over `w:trackRevisions` (a read leaves the settings part byte for byte, a write creates it when absent), the revision ids above the highest annotation id and apart from the comment ids, Word's rules one by one (a same-author `w:ins` extended, an insertion taken back, `w:del` before `w:ins`, `w:t` to `w:delText`, runs split at the boundaries), the paragraph-mark and row forms, `w:rPrChange` / `w:pPrChange` recorded once, `TrackedChange` with accept and reject of every kind, the paragraph join, the two views, `replace_text` tracked and not, and the two regressions the Word check found (section 16.9: a comment on text inside this author's own `w:ins` keeps its markers and its reference run **outside** it and survives both verbs; section 16.10: a deleted row's cells are emptied as Word empties them, and a paragraph appended at the end of a container leaves the final mark alone and marks the one before it; section 16.11: rejecting either mark form over a **loaded** document, through `reject_all()` and one change at a time, leaves the source document with no revision markup at all; **section 16.12, the invariant**: every tracked verb over a loaded document and then `reject_all()` leaves the body **canonically** what it was --- `scripts/canon.py`'s normal form, every `w14:paraId` still on its paragraph, no markup left --- over three documents, with the accept-side mirror beside it) |
| `content/test_office_js_subset.py` | CR-003 §3.4 (TS) | the committed Office JS member list |
| `content/test_markdown.py` | CR-003 §3.5 | markdown out and in, one construct at a time; the address comment and its regex; what `styles.xml` and `numbering.xml` are touched for |
| `agent/test_addresses.py` | CR-003 §3.4 | the three address forms, the nearest-address error, `ensure_para_ids` |
| `agent/test_outline.py` | CR-003 §3.4 | `Outline`, `depth`, `headings_only`, `limit`, the stats, the headers |
| `agent/test_find_and_describe.py` | CR-003 §3.4 | `SearchHit` and its range, and that `describe()` unmarshals nothing |
| `agent/test_change_report.py` | CR-003 §3.4, q4 | one report per call, `moved`, and the measured cost |
| `agent/test_dry_run.py` | CR-003 §3.4 | a trial edits a copy; the real package is byte for byte |
| `agent/test_sessions.py` | CR-003 §3.4, q5 | `DocumentSession`, the sweep, and eight threads on two handles |
| `agent/test_budgets.py` | CR-003 §7 | a 200-page document: outline under 64 KB, headings under 8 KB |
| `agent/test_scenarios.py` | CR-003 §7 | scripted tool-shaped sessions over the corpus, and determinism |
| `agent/test_errors.py` | CR-003 §3.4 | every error's `code` and `hint`, including where to split a span |
| `agent/test_tables_and_pictures.py` | CR-003 §7 | a paragraph in a cell addressed and edited, the `ChangeReport` it gives, `dry_run` of `insert_table` and of a picture, determinism of a part name and a relationship id, and `outline()` over nested tables |
| `agent/test_audit_trail.py` | CR-003 §3.4, §3.8 | the audit trail end to end: `pkg.author`, tracking on, an edit by address, `replace_text` after a `dry_run` count, a comment explaining it, `get_tracked_changes()` as JSON, save and reload, and `accept_all()` giving exactly what the same calls with tracking off give; determinism under a fixed seed and a fixed `tracked_change_date` |
| `agent/test_comment_workflows.py` | CR-003 §3.4, §3.9 | the audit trail: `pkg.author`, `find` then `insert_comment`, the `ChangeReport` a tool returns, a reply and a resolution, a thread's `to_dict()` under a 2 KB budget, `dry_run` leaving all four created parts un-created, and determinism |
| `agent/test_template_workflows.py` | CR-003 §3.7, §7 | the template story tool-shaped: `describe()` on the invoice (20 bindings, 2 repeats, the values, the repeat each binding is inside) reading the main part as **bytes** and leaving it byte for byte, the addresses once the body has been read, the `to_json()` size under 12 KB with a picture binding's base64 cut to 200 characters, `fill()` by XPath and by tag, a key that matches no binding reported rather than raised, `fill()` with a whole part as a string, a repeat reported but not expanded, a `dry_run` of a fill leaving the document alone, determinism of `add()` under a fixed seed, and describe → fill → save → reload end to end |
| `agent/test_markdown_workflows.py` | CR-003 §3.5, §7 | the coarse workflow (markdown in, markdown out) and the fine one (read with addresses, edit by address); `dry_run`, determinism, the markdown budget |
| `test_codegen_el.py` (the last two) | CR-003 §13 | every tracked hand-written path under `docx4j_py/` is inside `codegen/clean.py`'s `KEEP` |

## Fixtures

`samples/` holds sixteen documents. Thirteen `.docx` and one `.dotm` come from
docx4j's `sample-docs` and `docx4j-samples-docx4j` (Apache-2.0); `loadAndSave.pptx`
and `loadAndSave.xlsx` come from `docx4j-core-tests/src/test/resources`
(Apache-2.0) and are there for the generic loading path — CR-002 Phase A does not
type PresentationML or SpreadsheetML, and those two prove that it does not need to
in order to round-trip them.

`tests/fixtures/` holds the documents a single phase needs, copied from docx4j
(Apache-2.0) and listed here with their origin:

| file | origin | what it carries |
|---|---|---|
| `unknown_content.xml` | written here | CR-001 §7: an element and an attribute the bindings do not know |
| `lists.docx` | docx4j `docx4j-core-tests/src/test/resources/numbering_indentation.docx` | four `w:numPr` paragraphs on one `decimal` numbering definition, with a real `numbering.xml` |
| `hyperlink.docx` | docx4j `docx4j-core-tests/src/test/resources/AlteredParts/hyperlink.docx` | one `w:hyperlink` with an external relationship |
| `comments.docx` | docx4j `docx4j-core-tests/src/test/resources/AlteredParts/comments-one.docx` | one comment, for the `{>>...<<}` of `view="markup"` and for CR-003 Phase G's reads; no `w14:paraId`, and `w:comments` is the only comment part |
| `comments-modern.docx` | docx4j `docx4j-core-tests/src/test/resources/loadAndSave.docx` | Word-written, with **all five** comment parts: `w15:commentsEx` with `w15:done`, `w:people` with an Active-Directory email, `w16cid:commentsIds` and `w16cex:commentsExtensible`. The one comment has no reply, so the **thread** fixture is built in `tests/content/conftest.py` (`threaded_package()`) rather than fabricated (CR-003 section 15) |
| `footnotes.docx` | docx4j `docx4j-samples-docx4j/sample-docs/2010/w14_mcIgnorable-in-other-parts.docx` | a `w:footnoteReference` and the footnotes part it points at |
| `nested-table.docx` | docx4j `docx4j-layout-fidelity`'s corpus (`Corpus.java`'s `table-nested`) | a one-by-three table whose middle cell holds a two-by-two table, for CR-003 §3.2's `TableCell.tables` and `Table.parent_table_cell` |
| `tracked-pprchange.docx` | docx4j `docx4j-samples-docx4j/sample-docs/unmarshallFromTemplateDirtyExample.docx` | Word-written: a `w:pPrChange` with an empty `CT_PPrBase` original, a paragraph mark marked inserted (`w:pPr/w:rPr/w:ins`) and two run insertions, all by `jharrop` in 2012. The one document in any of the three checkouts with a `w:pPrChange` |

`samples/sample-docx.docx` carries one `w:ins` and one `w:del`, which is what the
CriticMarkup test and the stats test read. **A move, a `w:rPrChange` and a tracked
row** are in no document in docx4j, in docx4j-core-ts or here, so
`tests/content/conftest.py`'s `moves_package()` builds one in the open, as
`threaded_package()` does for a comment thread; it is not presented as a Word
document, because it is not one (CR-003 section 16.5).

**No `.docm` is in the corpus**: docx4j's repository has none to copy. `Normal.dotm`
covers the macro-enabled path (`application/vnd.ms-word.template.macroEnabledTemplate.main+xml`),
and the registry maps the `.docm` content type to the same `MainDocumentPart`.

## Word acceptance (manual)

Record each run here, in the form docx4j-core-ts's `test/README.md` uses. When a check finds
markup that makes Word **hang** or quietly refuse an operation, also add an entry to
`../docx4j-portfolio/docs/word-hangs.md`, the log shared across the three engines.

> Last run: 2026-09-17, Word (version not recorded), after CR-003 Phase F, artefact 8, second
> check. **Passed.** The first check the same day found the deleted row highlighted but not struck
> through, the appended paragraph invisible and **Reject All hanging Word**; the causes and the
> fixes are CR-003 sections 16.10 to 16.12 and the first entry of
> `../docx4j-portfolio/docs/word-hangs.md`. After the fixes every item in the row holds.
>
> Previous run: 2026-09-17, Word (version not recorded), after CR-003 Phase G, artefact 7.
> **Passed.** Three threads by Claude, the replies nested, the resolved thread resolved. The
> pane lists them in document order of their anchors (whole paragraph, then "first", then
> "document"), not in insertion order; the row above now says so. On save Word's Compatibility
> Checker flagged the one resolved thread ("collapsed"), because the source is in Word 2010
> compatibility mode; expected, and recorded in the row and in CR-003 section 15.7.
>
> Previous run: 2026-09-17, Word (version not recorded), after the CR-003 section 14.9 style fix,
> artefact 6 re-opened. **Passed**: "Tables and pictures" is Heading 1.
>
> Run: 2026-09-17, Word (version not recorded), after CR-003 Phase C, all six artefacts.
> **Passed except artefact 6**, whose "Tables and pictures" rendered as Normal rather than
> Heading 1. Word does not create a definition for a dangling `w:pStyle`; CR-003 section 14.9
> records the correction and a setter now adds the definition.
>
> Last run: 2026-09-17, Word (version not recorded), after CR-003 Phase K, all five artefacts.
> **Passed.** Every check in the table below held, including artefact 5's list glyphs, the
> restarting ordered list, the code styles, the link and the table, and a save-close-reopen.
>
> Previous run: 2026-09-12, Word (version not recorded), after CR-002 Phase A. **Passed.**
> 1: image with reflection, no text, as in the source document; 2: the two added paragraphs
> visible; 3 and 4: correct. No repair prompts. Re-checked after the docProps fix: Windows
> Explorer's Details tab shows Program name `docx4j-python` and Content created / Date last
> saved from the core part, rendered in local time.

Saved output must open in Word **without a repair prompt**. After any change to
marshalling, the prefix table, content types, the zip writer or `create_package`,
regenerate the artefacts and check by hand:

```bash
.venv-fork/bin/python scripts/acceptance.py      # writes out/acceptance/
```

`out/` is in `.gitignore`, so the nine files are built rather than committed; the script takes
about two seconds. It is deterministic apart from the `w:date` of a comment, which is the wall
clock, as Word writes it (CR-003 section 15.2 item 1); artefact 8's revision dates are the wall
clock for the same reason, since the script does not fix `pkg.tracked_change_date`.

### The nine artefacts

| file | what it exercises | what to look for in Word |
|---|---|---|
| `out/acceptance/1-untouched-round-trip.docx` | `2016_image_with_text_effects.docx` loaded and saved with **nothing** unmarshalled: every part is the source's bytes, and only `[Content_Types].xml` and the `.rels` parts come from this engine | opens with no repair prompt; the image, its text effects and the text are all there. A repair here means the zip writer, the content types or the relationships are wrong, and nothing else can be trusted. |
| `out/acceptance/2-remarshalled-round-trip.docx` | the same document with `document.xml`, `styles.xml` and `settings.xml` **unmarshalled and re-serialised**, and two paragraphs added (one through the `p()` builder, one through `el`) | no repair prompt; the two added paragraphs are at the end and the second keeps its leading and trailing spaces (`xml:space="preserve"`); the styles pane is unchanged; `w16se`/`w16cid` in `mc:Ignorable` did not upset Word. This is the test of §5.6. |
| `out/acceptance/3-created.docx` | `WordprocessingMLPackage.create_package()`: nothing came from a container | no repair prompt; A4 portrait with 2.54 cm margins; "Created by docx4j-python" is Heading 1; the styles pane offers Normal and Heading 1 to 4; the third paragraph is bold red 14 pt followed by plain text. |
| `out/acceptance/4-created-with-image.docx` | a created document plus an `ImagePart` added through `add_target_part` and placed with the `w:drawing` CR-003 Phase A's `inline_picture` builds, sized by `image_size` / `emu_for` from the PNG's own header | no repair prompt; the picture appears at its natural size, 9.00 cm by 6.56 cm (340 × 248 px at 96 dpi); right-click → Size shows those dimensions and **Lock aspect ratio ticked** (the `a:graphicFrameLocks noChangeAspect` the old hand-written fragment did not write); Alt Text shows the description; the image survives a Word save-and-reopen. |
| `out/acceptance/5-markdown-built.docx` | one markdown string through CR-003 Phase K's `insert_markdown`: headings, emphasis, inline code, a nested bullet list, an ordered list, a block quote, a fenced code block, a hyperlink and a GFM pipe table. It writes `styles.xml` (the styles from docx4j's `KnownStyles.xml`, plus `CodeChar` and `SourceCode`, which Word has no built-in equivalent of), creates `numbering.xml` from nothing, and adds an external relationship | no repair prompt; **Heading 1** and **Heading 2** appear in the navigation pane; the bullet list shows Word's own bullet glyphs with the nested level indented and using the second glyph; the ordered list is numbered 1 to 4 and **restarts at 1** (it is its own `w:num`); the quotation is in the Quote style; `x = 1` is in a grey Consolas block and `inline code` in grey Consolas within the paragraph; the link is blue, underlined and **Ctrl-click opens docx4java.org**; the table has Table Grid borders, a bold first row and the Total column right-aligned. Then **save from Word, close, reopen**: still clean, and the list numbering has not changed. |
| `out/acceptance/6-tables-and-pictures.docx` | a **loaded** document (`samples/2010-sample1.docx`) edited through CR-003 Phase C's content API: `insert_table(3, 3, values=…, style="TableGrid")`, `header_row_count = 1`, `add_rows`, `paragraph.insert_inline_picture(width=180)`, and a flat OPC `pkg:package` through `insert_ooxml` which brings a heading, a second picture and a second table | no repair prompt; **Tables and pictures** is a Heading 1 --- blue, bold, 14 pt, and in the navigation pane --- because the setter **added** the definition to `styles.xml`, which the source document does not carry (CR-003 section 14.9; Word does *not* supply it, which is what this row used to claim and what the 2026-09-17 run disproved); the **pasted** heading below is deliberately Normal, because `insert_ooxml` does not merge the source package's styles (CR-003 section 4); the first table has Table Grid borders, four rows, and its header row repeats if you force a page break inside it (Table Properties → Row → *Repeat as header row* is ticked for row 1 only); the pangolin under "A picture inserted at this paragraph:" is 6.35 cm wide with **Lock aspect ratio ticked** and Alt Text "Pangolin" / "A pangolin"; below it the pasted heading, a second, smaller pangolin (4.23 cm, alt text "the same pangolin") and a two-by-two Table Grid table are all there, and the two images are **different parts** (`/word/media/image1.png` and `image2.png`). Then **save from Word, close, reopen**: still clean. |
| `out/acceptance/7-comments.docx` | a **loaded** document (`samples/2010-sample1.docx`, which has **no** comment parts) given, through CR-003 Phase G's content API, a comment on the range `find("first")` returned, a reply in the same thread, a second comment with a reply that is then **resolved**, and a third on a whole paragraph. All four comment parts are created with their relationships and content types, and `CommentText`, `CommentTextChar` and `CommentReference` are added to `styles.xml` | no repair prompt; **Review → Show Comments** (or the markup pane) shows three threads by **Claude**, initials **C**, with today's date, in **document order of their anchors** (the whole-paragraph comment first, since its range starts at the paragraph's start, then **first**, then **document**), not in the order the script inserted them; the **first**-range thread's reply is **nested under it**, not a thread of its own; the second thread is shown **Resolved** (greyed out, with a *Reopen* button) and its reply is nested too; the highlighted range of the first thread is exactly the word **first** in "My first 2010 document.", of the second exactly **document**, and of the third the **whole first paragraph**; the comment text is in Word's Comment Text style at 10 pt. Then **save from Word, close, reopen**: still clean, and the threads, the nesting and the resolved state all survive. **Expected on that save**: Word's Compatibility Checker reports "Comments which have been collapsed will no longer be collapsed", 1 occurrence. The source is a Word 2010 document (`compatibilityMode` 14) and a resolved thread (`w15:done`) is a Word 2013 feature; Word says the same when a comment is resolved in Word on such a document. *Continue* is fine; it is not a defect in the output. |

| `out/acceptance/8-tracked-changes.docx` | the same **loaded** document with `pkg.change_tracking_mode = "TrackAll"` and then, through CR-003 Phase F's content API, `replace_text("first", "second")`, a paragraph inserted **after the first one**, another **appended at the very end** (the final-mark rule of section 16.10), a deleted paragraph, `font.bold = True` on another, a table row added and a row deleted, and a comment explaining the replacement. The table and two paragraphs are written **before** the mode goes on, so that there is something of the document's own to delete and to re-format and so that the body ends with a paragraph | no repair prompt; **Review → Track Changes is on** and the Reviewing pane lists every change with **Claude** and today's date: *Inserted* "second" and *Deleted* "first" in the first paragraph, an *Inserted* paragraph ("Added by an agent, right after the first paragraph.") **visible near the top**, another at the very end, a *Deleted* paragraph ("This paragraph will be deleted, with its mark.") whose paragraph mark is deleted too, a *Formatted* run ("This paragraph will be made bold." shown bold with a formatting balloon), an *Inserted* table row (North / Q2 / 300) and a *Deleted* table row (East / Q2 / 180), which is **still shown and struck through** --- its cell text struck, not merely highlighted --- until it is accepted. The comment by **Claude** is anchored on **second**. Then, on a copy, **Review → Accept All**: the document reads "My second 2010 document.", the deleted paragraph and the East row are gone, both added paragraphs and the North row are there, the bold stays, and the comment survives. On another copy, **Reject All** --- which must **not hang** --- : the first paragraph reads "first" again, the deleted paragraph and the East row are back, the North row and both added paragraphs are gone, the bold is gone, and the comment is still there. Then **save from Word, close, reopen**: still clean. **Expected on that save**: the Compatibility Checker may report the resolved-comment / collapsed-comment note again, for the reason artefact 7's row gives (the source is a Word 2010 document). |

| `out/acceptance/9-template-filled.docx` | `samples/invoice2013.docx`, a Word-authored template with twenty bindings over `/customXml/item1.xml`, **filled** through CR-003 Phase E's `pkg.custom_xml_parts.fill()`: a new company and contact, a new invoice number, the VAT checkbox cleared, a new date, and new data for the first line item of a `w15:repeatingSection`. Then a **new** custom XML part through `add()` --- `/customXml/item2.xml` with its `/customXml/itemProps2.xml`, a brace-wrapped upper-case `ds:itemID` and the relationship from the main document part, as docx4j's `addPropertiesPart` writes it --- a new paragraph, a run-level content control in it, and a mapping to the new part's `<by>` node | no repair prompt; the bound controls show the **new** values on open, before Word has refreshed anything: the customer is **Acme Manufacturing Ltd**, the contact **Ada Lovelace**, the invoice number **INV-2026-0917**, the VAT checkbox is **empty** (☐, in MS Gothic) and the date reads **17 September 2026** in the template's own `d MMMM yyyy` format. The line-item row shows **ACME-9 / Anvil, large / 2 / 199.00**; there is still **one** line item --- a repeating section is a container and is never bound, so the repeat is **not** expanded and Word will not add rows of its own (CR-003 section 17). At the end, "Approved by: **Grace Hopper**", where *Grace Hopper* is the new control, bound and showing its value. **Developer → XML Mapping Pane** (turn the Developer tab on in File → Options → Customize Ribbon if it is not there) offers **two** custom XML parts in its drop-down: the invoice's, and the new `http://example.com/approval` one whose tree shows `approval/by` and `approval/status`; clicking into the "Grace Hopper" control highlights `by` in the pane. Then **save from Word, close, reopen**: still clean, and every value above is unchanged. **Expected**: Word may report the Compatibility Checker note about the picture control or the repeating section on an older compatibility mode; the source is Word 2013 and the repeat is untouched, so *Continue* is fine. |

### Checks worth making on every one

1. **No repair prompt**, and no "Word found unreadable content" dialogue.
2. **Save from Word**, close, reopen: still clean. Word rewriting the file is the
   strongest signal that it understood it.
3. **Properties.** Word's File → Info panel shows the core part's dates under
   *Related Dates* (`create_package` writes `created` and `modified`) and, for
   the round trips, Word's own author and title. It never shows `Application` or
   `AppVersion`: those are visible only outside Word, in Windows Explorer's file
   Properties → Details tab as *Program name*, or by unzipping and reading
   `docProps/app.xml` (`create_package` writes `docx4j-python` and `0.0001`).
   For 3 and 4, unzipping is the check that counts.
4. For 2, 4, 6 and 7, **Review → Compare** against the source is a quick way to see that
   only the intended change happened.

### What has *not* been checked

* PowerPoint and Excel: `loadAndSave.pptx` and `loadAndSave.xlsx` round-trip
  byte for byte through the generic path, but nobody has opened the output.
* Flat OPC (`pkg:package`), which is Phase C.
* A macro-enabled document with a real VBA project.
