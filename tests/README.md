# Tests

```bash
.venv-fork/bin/python -m pytest                    # everything, 940 tests
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
| `comments.docx` | docx4j `docx4j-core-tests/src/test/resources/AlteredParts/comments-one.docx` | one comment, for the `{>>...<<}` of `view="markup"` |
| `footnotes.docx` | docx4j `docx4j-samples-docx4j/sample-docs/2010/w14_mcIgnorable-in-other-parts.docx` | a `w:footnoteReference` and the footnotes part it points at |
| `nested-table.docx` | docx4j `docx4j-layout-fidelity`'s corpus (`Corpus.java`'s `table-nested`) | a one-by-three table whose middle cell holds a two-by-two table, for CR-003 §3.2's `TableCell.tables` and `Table.parent_table_cell` |

Tracked changes need no fixture of their own: `samples/sample-docx.docx` already
carries one `w:ins` and one `w:del`, which is what the CriticMarkup test reads.

**No `.docm` is in the corpus**: docx4j's repository has none to copy. `Normal.dotm`
covers the macro-enabled path (`application/vnd.ms-word.template.macroEnabledTemplate.main+xml`),
and the registry maps the `.docm` content type to the same `MainDocumentPart`.

## Word acceptance (manual)

Record each run here, in the form docx4j-core-ts's `test/README.md` uses:

> Last run: 2026-09-17, Word (version not recorded), after the CR-003 section 14.9 style fix,
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

`out/` is in `.gitignore`, so the six files are built rather than committed; the script is
deterministic and takes about two seconds.

### The six artefacts

| file | what it exercises | what to look for in Word |
|---|---|---|
| `out/acceptance/1-untouched-round-trip.docx` | `2016_image_with_text_effects.docx` loaded and saved with **nothing** unmarshalled: every part is the source's bytes, and only `[Content_Types].xml` and the `.rels` parts come from this engine | opens with no repair prompt; the image, its text effects and the text are all there. A repair here means the zip writer, the content types or the relationships are wrong, and nothing else can be trusted. |
| `out/acceptance/2-remarshalled-round-trip.docx` | the same document with `document.xml`, `styles.xml` and `settings.xml` **unmarshalled and re-serialised**, and two paragraphs added (one through the `p()` builder, one through `el`) | no repair prompt; the two added paragraphs are at the end and the second keeps its leading and trailing spaces (`xml:space="preserve"`); the styles pane is unchanged; `w16se`/`w16cid` in `mc:Ignorable` did not upset Word. This is the test of §5.6. |
| `out/acceptance/3-created.docx` | `WordprocessingMLPackage.create_package()`: nothing came from a container | no repair prompt; A4 portrait with 2.54 cm margins; "Created by docx4j-python" is Heading 1; the styles pane offers Normal and Heading 1 to 4; the third paragraph is bold red 14 pt followed by plain text. |
| `out/acceptance/4-created-with-image.docx` | a created document plus an `ImagePart` added through `add_target_part` and placed with the `w:drawing` CR-003 Phase A's `inline_picture` builds, sized by `image_size` / `emu_for` from the PNG's own header | no repair prompt; the picture appears at its natural size, 9.00 cm by 6.56 cm (340 × 248 px at 96 dpi); right-click → Size shows those dimensions and **Lock aspect ratio ticked** (the `a:graphicFrameLocks noChangeAspect` the old hand-written fragment did not write); Alt Text shows the description; the image survives a Word save-and-reopen. |
| `out/acceptance/5-markdown-built.docx` | one markdown string through CR-003 Phase K's `insert_markdown`: headings, emphasis, inline code, a nested bullet list, an ordered list, a block quote, a fenced code block, a hyperlink and a GFM pipe table. It writes `styles.xml` (the styles from docx4j's `KnownStyles.xml`, plus `CodeChar` and `SourceCode`, which Word has no built-in equivalent of), creates `numbering.xml` from nothing, and adds an external relationship | no repair prompt; **Heading 1** and **Heading 2** appear in the navigation pane; the bullet list shows Word's own bullet glyphs with the nested level indented and using the second glyph; the ordered list is numbered 1 to 4 and **restarts at 1** (it is its own `w:num`); the quotation is in the Quote style; `x = 1` is in a grey Consolas block and `inline code` in grey Consolas within the paragraph; the link is blue, underlined and **Ctrl-click opens docx4java.org**; the table has Table Grid borders, a bold first row and the Total column right-aligned. Then **save from Word, close, reopen**: still clean, and the list numbering has not changed. |
| `out/acceptance/6-tables-and-pictures.docx` | a **loaded** document (`samples/2010-sample1.docx`) edited through CR-003 Phase C's content API: `insert_table(3, 3, values=…, style="TableGrid")`, `header_row_count = 1`, `add_rows`, `paragraph.insert_inline_picture(width=180)`, and a flat OPC `pkg:package` through `insert_ooxml` which brings a heading, a second picture and a second table | no repair prompt; **Tables and pictures** is a Heading 1 --- blue, bold, 14 pt, and in the navigation pane --- because the setter **added** the definition to `styles.xml`, which the source document does not carry (CR-003 section 14.9; Word does *not* supply it, which is what this row used to claim and what the 2026-09-17 run disproved); the **pasted** heading below is deliberately Normal, because `insert_ooxml` does not merge the source package's styles (CR-003 section 4); the first table has Table Grid borders, four rows, and its header row repeats if you force a page break inside it (Table Properties → Row → *Repeat as header row* is ticked for row 1 only); the pangolin under "A picture inserted at this paragraph:" is 6.35 cm wide with **Lock aspect ratio ticked** and Alt Text "Pangolin" / "A pangolin"; below it the pasted heading, a second, smaller pangolin (4.23 cm, alt text "the same pangolin") and a two-by-two Table Grid table are all there, and the two images are **different parts** (`/word/media/image1.png` and `image2.png`). Then **save from Word, close, reopen**: still clean. |

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
4. For 2, 4 and 6, **Review → Compare** against the source is a quick way to see that
   only the intended change happened.

### What has *not* been checked

* PowerPoint and Excel: `loadAndSave.pptx` and `loadAndSave.xlsx` round-trip
  byte for byte through the generic path, but nobody has opened the output.
* Flat OPC (`pkg:package`), which is Phase C.
* A macro-enabled document with a real VBA project.
