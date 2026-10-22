# CR-011: One text model --- the port of docx4j CR-046

**Status:** Proposed 2026-10-18, on docx4j CR-046 phase 4's hand-off (the goldens at docx4j
`d9acdf5e2`, the specification at WD 0.14), as the portfolio paper's section 10 row for this
package proposes: `../docx4j-portfolio/docs/text_extraction_behaviour_model.md`, whose section 9
recommendations the user accepted on 2026-10-17 (portfolio `2d8fd36`). It takes in CR-010 Phase
C's view agreement (CR-010 sections 6 and 12.3). **Decided 2026-10-21** (the user: "go for
CR-011"): every question of section 9 on its recommendation. Q4 is the revised answer (the Office
JS readers keep Office JS's text, as core-ts CR-012's D2), and Q11's fix waits for docx4j CR-052's
hand-off (its note). **Phase A and Q10 implemented the same day** (section 11): the suite at
`198d015a3`, its harness and the baseline as a ratchet, 941 pass, 410 fail, 233 pending of
1,584. **Phase B implemented the same day** (section 12), with Q3 and with Q11 on docx4j CR-052's
placement, which landed as `5df2e253e`: `docx4j_py.model.text`; 1,146 pass, 205 fail, 233 pending,
goldens 6 of 12; the agreement test at zero. Q11's Word check passed the same day, and Q3's
change to the xsdata fork (stage 10) needs a fork release before a docx4j-python release. **Phase C
implemented the same day** (section 13), the port of docx4j CR-046 phase 2: fields updated with
docx4j's switches, bound values, charts and SmartArt (`docx4j_py.model.chart`), WordArt, embedded
WordprocessingML, the glossary, plain-text `w:altChunk`s (the HTML reader absent, Q13), pictures'
alt text, metadata, equations as LaTeX, listed and grid tables, page markers and `layout: indent`;
1,284 pass, 7 fail (Markdown's, Phase D's), 293 pending, goldens 9 of 12; Q7's 60 pending by name.
**Phase D implemented the same day** (section 14), the port of docx4j CR-046 phase 3:
`markdown(result)`, docx4j-markdown's exporter with commonmark-java's renderer ported, byte for
byte docx4j's on 1,385 of 1,385 exports; `to_markdown` and `outline()` on the model, the Office JS
readers kept on Office JS's text (Q4); a read leaves every part's bytes as they were; CR-010's ten
pins closed. 1,291 pass, 0 fail, 293 pending (docx4j's 19, the inspector's 214, Q7's 60), goldens
12 of 12: section 3.6's gate is met and **the four phases are done**. Open: Q7's content (after
CR-001 D), the HTML reader as its own phase (Q12, Q13), the `dry_run` comment defect (12.3), and
the fork release, which came out as docx4j-xsdata 26.2.2 on 2026-10-22. **Q4 revised 2026-10-22**
(the user: "revise CR-011 Q4 so that Paragraph.text, get_text and search read the branch Word
shows"): Office JS's own text reads a kept `mc:AlternateContent`'s selected branch (core-ts check
49), so the Office JS readers, kept on it, read Word's branch too; section 15 proposes the change as
Phase F, as core-ts CR-012 phase 4 made it there and as docx4j's readers of text select
(`McSelection`'s TEXT purpose). **Phase F started 2026-10-22** at the user's go ("go for phase F,
now please").
**Depends on:** docx4j CR-046 phases 0 to 3, done 2026-10-18 (`c0c72ecef`, `65a527d47`,
`f71cdd299`, `8da15bb8c`, `608e40023`, `187e5a367`, `f59828e62`, `d9acdf5e2` on `VERSION_17_3_4`;
from `608e40023` on unpushed, so the local clone is the oracle); the specification
`../docx4j-portfolio/docs/text_extraction_specification_wd.md` at WD 0.14 (portfolio `cf56600`);
this repository's CR-010 (implemented; its ten `xfail(strict)` pins are this CR's to close); for
Q3, a fork release. **Updated 2026-10-19** on docx4j-25's notes of `a41de133b` and `3dcbb3764`
(both unpushed). The suite was then cited at `3dcbb3764`. That commit adds a third fixture,
`extraction-probe-structure.docx`, with cases 55 to 57 (a table of contents, an interrupted list,
an embedded workbook's hidden sheets) and four goldens; the eight earlier goldens are unchanged
since `d9acdf5e2`. The specification was then at WD 0.15 (portfolio `ac06ed7`, `7affcd0`). A
suite row may now combine options, and every row needs its scenario's statements (REQ-1201;
sections 3.6 and 6). The choices docx4j CR-046 phase 2 made where WD 0.8 was silent
are decided and normative (Jason, 19 October 2026): REQ-519's addresses, apart from the text-box
segment, which stays a proposal (section 3.2); REQ-515's plain-text property lines; REQ-604's markup
form of an equation; and REQ-512's hidden sheets and slides under `hidden: true`.
**Counterpart:** docx4j's `docs/developer/change-requests/CR-046-text-extraction-behaviour-model.md`
(section 2 the design, 4 the phases, 5 every statement and where docx4j meets it, 6 its decisions),
and its code, `org.docx4j.model.text` (27 classes, 6,203 lines at `d9acdf5e2`; at `198d015a3`, 42
and 12,456: the text model's 25 and 6,092, CR-053's HTML reader 3 and 973, and CR-050's inspector,
out of scope here (section 3.7), 14 and 5,391), and the harness
`org.docx4j.textextraction.ExtractionConformance` with its engine adapter `TextsAdapter` and
`ExtractionInspection`, which names the inspection tables.
docx4j-core-ts's CR-012 (`docs/change-requests/CR-012-text-model.md`) is the sibling port,
proposed 2026-10-19 and **decided by Jason on 2026-10-21**: D1 to D9, all as recommended (core-ts
`9b302ad`, pushed). Five of its decisions touch this CR: D1 (options keyed by the specification's
section 10 names verbatim) Q1, D2 (the Office JS readers keep Office JS's text) Q4 and section
3.5, D4 (the text-box address) section 3.2, D5 (text-box views read-only) Q9, and D9 (REQ-510 in
the emulator first) risk 3. D2 decides core-ts's readers only (its section 6.1); Q4 here is still
the user's.

**Updated 2026-10-21** (the user: "start CR-011 re-cite"). The suite is cited at docx4j `198d015a3`
and the specification at WD 0.28 (portfolio `f2046e5`). Since `3dcbb3764`, docx4j CR-050 (the
inspector: `8d32bf928`, `1ae8f1479`, `4d9345199`, `c2643af0d`, `86449c013`) and CR-053 (`1d108c0af`,
`4a8833eea`, `198d015a3`) have changed the suite; CR-053's hand-off (`d2d0a3861`) names those three
and `35e304770`, its writing half. All are on `VERSION_17_3_4`, pushed the same day with the branch
at `4b9c71288`, which does not touch the suite. The suite has five fixtures. The fourth,
`extraction-probe-inspect.docx`, holds cases 58 to 67 (97 token rows, no golden). The fifth,
`extraction-probe-html.docx`, is read only under `html reader present`, and only by an inspector.
CR-050's inspection rows need REQ-901, 902 and 904 to 907 (the counts, REQ-903) and pend here
(section 3.7). docx4j's own run is 1,565 passed and 19 pending of 1,584 assertions with no HTML
reader, and 1,526 and 58 with one. CR-053 changes REQ-513 (WD 0.23 to 0.28). An HTML, XHTML or MHT
chunk (an MHT's first `text/html` part) is read only by an *HTML reader*: a parser of the [HTML]
algorithm and a reader of its tree as blocks. Without one, the chunk is not expanded and raises
`altchunk.not_expanded` with its content type. The suite's scenario is now `html reader absent` (it
was `xhtml importer absent`). A chunk is decoded by its BOM; else a meta charset in its first 1,024
bytes; else UTF-8 where valid; else windows-1252. There is no XML-declaration step and no UTF-16
sniff. What `display: none` hides in a chunk, by Word's subset of CSS, is hidden text, and a chunk
of more than 50,000 elements is not expanded. Section 6's earlier `lxml.html` and section 3.6's
built-in MHT expansion predated this: libxml2's HTML parser does not follow the [HTML] algorithm
(Q12). Both are replaced, and the reader of the tree is Q13's. The harness's port counts docx4j's
1,584 (section 3.6), and the baseline is rerun (section 1: 941 pass, 410 fail, 233 pending; 3 of 12
goldens). WD 0.16 to 0.28's other changes in this CR's scope need no change of design: REQ-504
leaves out a comment with no anchor (0.20); REQ-519's text-box segment is decided (0.27; section
3.2); REQ-1201 and REQ-1203 make a row need its scenario's capability (0.16, 0.27; section 3.6);
section 10's options are as they were. REQ-202's new exemption for a reader of an emulated API
(0.27) bears on Q4 and Q5.

## 1. Why

The paper (its section 1) found that the three engines answer "what is the text of this
document?" differently, and that this package answers it five ways. Measured here on 2026-10-18
against docx4j's conformance suite (then two fixtures, 54 cases, 941 checks), again on
2026-10-19 against the suite at `3dcbb3764` (three fixtures, 57 cases, 1,023 checks), and on
2026-10-21 at `198d015a3` (five fixtures, 67 cases, 1,584 checks):

- **Five paths, five answers.** Of the probe's 164 tokens in the accepted view, `get_text()` gets
  121 to their expected count, `text_of` 121, `to_markdown` 119, the revisions engine's tree walk
  123 and the python-docx facade 106; the five disagree with each other on 25 tokens.
- **Against the suite**, an adapter over today's `get_text(view=)` and `to_markdown(view=)`, with
  docx4j's statement set advertised so that only docx4j's 51 pending checks pend: **591 pass, 299
  fail** at `d9acdf5e2`, and **648 pass, 324 fail** at `3dcbb3764`. Of the 25 new failures, 15 are
  content that Q7 defers: case 53's hidden slide under `embedded: true; hidden: true` (3, from
  `a41de133b`) and case 57's embedded workbook (12). Another 6 are case 55's table-of-contents
  entries, read by default (REQ-520). The other 4 are the new fixture's goldens. docx4j's own run is
  890 pass, 0 fail, 51 pending at `d9acdf5e2` (42 need the inspector, 9 an XHTML importer), and
  972 pass, 0 fail, 51 pending at `3dcbb3764`.
- **At `198d015a3`** the same adapter, advertising docx4j's statements less the inspector's
  (REQ-901 to 907, REQ-1204), gives **941 pass, 410 fail, 233 pending** of 1,584. The 233 are
  docx4j's 19 (the checks that need an HTML reader) and the inspector's 214 (186 `inspect` rows, 4
  inspections, 24 counts; section 3.7). On the three earlier fixtures 3 fewer fail. Case 36's
  chunk text passes, now expected only with an HTML reader. Its new warning row fails, since
  nothing here raises one. Case 47's warning row pends, now under `html reader present`. The fourth
  fixture's 477 checks give 290 pass, 89 fail and 98 pending; the fifth's one check pends. Of the
  89, 43 are in the default scenarios: text boxes (18), hidden text (9), bound values (6), notes
  (3), run-level `mc:AlternateContent` (3), the chunk warning (3) and a comment in markup (1). The
  other 46 are options that today's paths ignore: `field_codes: true` (18), `embedded: true` alone
  or with `hidden: true` (15, of which case 66's workbook, 9, is content Q7 defers), `include:
  headers` (6), `alt_text: true` (3), the chunk warning under `altchunks: false` (3) and `comments:
  true` (1). docx4j's own run is 1,565 pass, 0 fail, 19 pending, and with docx4j-ImportXHTML 1,526
  and 58.
- **Goldens: 3 of 8** pass (the external fixture's three plain-text views); the probe's four
  differ by 66 to 106 lines, the external Markdown by a dropped linked picture. At `3dcbb3764` the
  count is 3 of 12. The structure fixture's three plain-text goldens fail on the table of
  contents' two entries (REQ-520). Its Markdown golden fails on those entries and on the
  interrupted list: the indented paragraph's continuation, `4. 2.` written `  2.`, and the second
  list's `1)` written `1.` (REQ-703, REQ-510). At `198d015a3` it is still 3 of 12: the probe's
  four, recut by `1d108c0af` without case 36's line, differ by 65 to 104 lines, the others as
  before.
- **The failures, by class** (169 in the default scenarios at `d9acdf5e2`): text boxes and
  `mc:Choice` content not read (49), charts and SmartArt (27), `w:altChunk` not expanded and no
  warning (24), notes not appended (12), no plain-text markup view (10), hidden text included (9),
  math and ruby (9), no metadata channel (9), run-level `mc:AlternateContent` not read (6), bound
  values not read (6), joins, table layout and note markers in the goldens (5), comments in markup
  (3). At `198d015a3` they are 219: text boxes and `mc:Choice` 67, charts and SmartArt 27,
  `w:altChunk` 24, hidden text 18, notes 15, bound values 12, the plain-text markup view 10, math
  and ruby 9, metadata 9, run-level `mc:AlternateContent` 9, the goldens 9, the table of contents 6,
  comments in markup 4.
- **The views are filters, not resolutions** (paper section 3.1, REQ-401 to 406). `get_text()`
  against the text of a reloaded, resolved copy differs on 6 of CR-010's 11 fixtures and the probe
  in the accepted view and on 4 in the original; over docx4j's 139 revisions and text-extraction
  fixtures, on 79 and 80. CR-010 Phase C pinned the first set (`tests/revisions/test_content_api.py`).
- **`get_text(view="markup")` is neither view**: 5,093 characters on the probe, against 5,166
  accepted and 5,148 original. It is a filter that keeps neither holder. Section 9, Q10.
- **`mc:AlternateContent` is selected by prefix, and an empty `Requires` is understood** by the
  traversal (`docx4j_py/child.py:532`, `all()` over nothing), where the preprocessor
  (`docx4j_py/openpackaging/mce.py:150`) and the decided rule (REQ-304) say it is not. The
  understood set is 63 URIs against docx4j's published 118 (section 3.3).
- **Numbering labels**: over `LevelOwnerTest`'s three Word documents (61 body paragraphs), the
  emulator differs from Word on 27, every one a level that names no `w:pStyle` and the counts that
  follow it (REQ-510).

docx4j has the model, held to Word by the suite and its goldens, and its readers all on it. This
CR ports it, so that an agent reading a document through any of the three engines, or through any
reader of this one, gets one text.

## 2. What exists here

The five paths (file and line at `99d46a1`):

| # | Path | Readers | What it does today | Becomes |
|---|---|---|---|---|
| 1 | `traversal.text_of` (`traversal.py:432`) | `Block.text`, `Table.text`, `TableCell.text`, `ContentControl.text`, the outline's table preview, Markdown's nested table, the namespace packages' re-export | a fixed filter (`w:ins`, `w:moveTo` in; `w:del`, `w:moveFrom` out); `mce="resolve"` by prefix, losing a paragraph-level branch's wildcard runs; hidden text in; a line per cell paragraph | **deprecated** (REQ-202's exemption: kept, unused by any shipped reader), its readers moved to the model |
| 2 | `text_model.segments_of` / `text_of_view` (`model/content/text_model.py:194`, `259`) | `Paragraph`, `Body` and `Range` `text`, `get_text`, `text_budget`; `find_in`; the outline's counts and previews; `controls.set_text` | views as holder filters (`enums.py:61`); no mc; text boxes absent; hidden text in; a deleted row as empty lines; `"markup"` broken | **the model**, through the readers (Phase D) |
| 3 | Markdown `_render_run` / `_render_inlines` (`model/markdown/export.py:427`, `472`) | `Body.to_markdown`, `Paragraph.to_markdown`, `pkg.to_markdown` | `accepted` and `markup` (CriticMarkup, `export.py:472`), `original` raising (`export.py:968`); text boxes dropped with a warning (`export.py:402`); footnotes `[^n]`; comments in markup | **a serialiser of the model's blocks** (Phase D), keeping its GFM and CriticMarkup |
| 4 | `revisions/trees.item_text` / `append_text` (`trees.py:412`, `457`), through `revisions/texts.py:34` | `TrackedChange.text` | Office JS's form (`\r` for a mark, `\t` between cells, `\r\n` at a row's end), both branches, math | **stays**: it describes a change, not a document (Q6); the walk is reused |
| 5 | the python-docx facade's `run_text` / `holder_text` (`docx/text.py:128`, `147`) | `docx.Paragraph.text`, `Run.text`, cell text, comments | python-docx 1.2.0's semantics: `w:ins` and `w:del` both skipped, no `sdt`, breaks give nothing | **stays** as a declared emulation (Q5) |

Two smaller walks read text for their own ends (Markdown's comment text through lxml, at
`export.py:255`; the offset walks of `comments.py:267` and `tracked_change.py:461`); Phase D
moves the first and leaves the offset walks, which index rather than read.

What the model can stand on:

- **The view by resolution.** `revisions.accept_all` / `reject_all` cover every story and agree
  with Word on 22 of 22 comparisons (CR-010 section 12.7); `pkg.dry_run()` gives a resolved copy
  and leaves the package as it was (reload and resolve: 25 to 29 ms on the probe). The `Lister`
  and `trees.append_text` are already view-aware walks, math included.
- **The MCE trees.** Both branches are kept on load; `revisions/islands.py` already types a
  `w:txbxContent` into a story it can walk (CR-010 section 3.5).
- **CriticMarkup** (`export.py:448`, `472` to `500`). Written into plain text, it passes all 29
  markup checks of cases 9 to 16 and 28. It nests a deletion inside an insertion
  (`{++A {--B--}++}`) where the golden writes them side by side (`{++A ++}{--B--}`).
- **The numbering emulator.** `labels_for` (`model/listnumbering/emulator.py:967`) gives labels
  and counts; `style_linked_elsewhere` (`emulator.py:400`) is docx4j 17.1.0's rule, matching by
  `w:pStyle` only. REQ-510's owner is docx4j's `Emulator.levelStyleId` (`Emulator.java:427` to
  `454`).
- **The rest.** Hidden text: `PropertyResolver.get_effective_rpr` (`property_resolver.py:484`).
  Bound values: `XmlMapping.custom_xml_node` (`customxml/mapping.py:103`). Charts and SmartArt:
  the typed `dml.chart` and `dml.diagram` models. The block: `OutlineEntry`
  (`reports.py:133`: address, para id, style id, level, text). HTML for `w:altChunk`: nothing that
  conforms, since `lxml.html`'s libxml2 parser does not follow [HTML]'s algorithm (REQ-513, Q12).
  Not available: an OLE reader (`olefile` is not a dependency), a field evaluator, a number-format
  engine, and any SpreadsheetML or PresentationML model (CR-001 D).

**Found on the way, outside this CR.** A re-marshalled save drops a namespace declaration that
only an `mc:Choice`'s `Requires` uses: the probe declares `xmlns:shape2010` on its
`mc:AlternateContent` (case 6), and the saved `document.xml` keeps `Requires="shape2010"` with
the prefix undeclared. The file still parses, but the prefix no longer resolves, which MCE treats
as an error. This breaks CLAUDE.md's "nothing may be dropped silently", and it is an engine
defect (CR-002 section 12.3), recommended for a fix ahead of this CR (Q11). It is the same gap as
REQ-303's: the typed tree keeps the prefix and not its URI.

## 3. Design

### 3.1 The package

`docx4j_py/model/text/` (under `model/`, already in `KEEP`), mirroring `org.docx4j.model.text`
module for class where Python allows:

- `TextView`: `accepted` (the default), `original`, `markup`; `current` accepted as a synonym
  (REQ-401). It extends the content API's existing `TextView` (`enums.py:61`) rather than adding
  a second one.
- `TextOptions`: one field per WD section 10 option, with the WD's names and values verbatim
  (they are already snake case: `text_boxes`, `note_references`, `field_codes`), and
  `TextOptions.of(mapping)` (REQ-1001). An unimplemented value raises, naming its statement, as
  docx4j's section 2.6 does.
- `extract(target, options=None, **options) -> TextResult`, over a package, a part or a content
  API container (docx4j's `Texts.of`). The package is never modified (REQ-204).
- `TextResult`: `blocks` (the block model, a list of `TextBlock`), `warnings` (`TextWarning`:
  kind, address, branch), `metadata` (the document properties and `fields_update_on_open`).
- `TextBlock` per REQ-706: address, para id, style, the text in the view, runs (in `markup`,
  each with its revision), the flags, and the typed objects (field, image, object, chart,
  SmartArt, numbering: `num_id`, `level`, `format`, `level_text`, `start`, `value`,
  `abstract_num_id`).
- The serialisers: `plain_text(result)` (WD sections 7.1 and 7.5) and `markdown(result)` (7.3,
  REQ-703), both written from the blocks, so that no reader can drift from another (REQ-202). The
  block model's JSON is `TextResult.to_dict()`, for CR-006.

### 3.2 The four layers (REQ-201)

1. **Resolve.** Each `mc:AlternateContent` gives one branch by the TEXT purpose: the first
   `mc:Choice` all of whose `Requires` namespaces are understood, else the `mc:Fallback`
   (REQ-301, 302); `Requires` matched by **URI**, not prefix (REQ-303, section 3.4); an empty
   `Requires` not understood (REQ-304); an element with Choices and no Fallback, none understood,
   gives its first Choice (REQ-305); `mce.branches_differ` warned where the branches' text
   differs (REQ-307). A text box's content is read through the typed island of CR-010 section 3.5.
2. **View.** As docx4j's section 2.3: **the walk is the implementation, the copy is its oracle.**
   The walk does to the text what CR-010's resolver does to a copy, without making one --- a
   deleted mark joins in `accepted`, an inserted one in `original`, a row is there or not, a field
   goes as a unit, a note whose reference is deleted is absent, and the joined block takes the
   first paragraph's para id (REQ-402 to 405). It is built on `revisions/trees.py`'s holders and
   `append_text`. `markup` returns every run as stored, with its revision, and resolves nothing
   (REQ-403). The agreement test (REQ-406) holds the walk equal to `extract` over a `dry_run`
   copy resolved by `revisions.accept_all` / `reject_all`, over CR-010's eleven fixtures, the
   probe and, where the sibling checkout is present, docx4j's 139.
3. **Scope.** The stories and content classes of WD section 5 under section 10's options: the main
   story; text boxes after their anchoring paragraph, once (REQ-518), with `textbox:` addresses
   (REQ-519, core-ts CR-002 section 3.3's grammar; decided in WD 0.15, and its text-box segment in
   WD 0.27, which adopts core-ts CR-012's D4: docx4j's `body/3/textbox:0/1` as docx4j `608e40023`
   implements it); notes as
   `[n]` with their text appended (REQ-502); headers and footers
   by section under `include` (REQ-503), their blocks at `header:<rId>/<n>` and
   `footer:<rId>/<n>`, which `addresses.py` already writes (`footer:rId9/0` and `header:rId4/0`
   to `header:rId7/0` on the probe, as docx4j's oracle has them). docx4j's own were wrong before
   `8d32bf928`, its part's `toString()` prefixed (`org.docx4j.wml.Ftr@50bb1c1ffooter:rId9/0`, a
   shadowed local in `Extractor.headerOrFooter`), so the port copies the form, not that code
   before the fix; no fixture here was cut from it; comments in markup, never one with no anchor
   (REQ-504, WD 0.20); field codes and
   `field.no_result` (REQ-505); hidden text out by default, resolved through the styles (REQ-507);
   placeholders flagged (REQ-508); bound values (REQ-509); list labels with REQ-510's owner; the
   table of contents under `toc` (REQ-520); and Phase C's content classes (section 6).
4. **Serialise.** Characters as WD section 6.1 (with `w:ptab`); one line per paragraph and break;
   tables as `auto`, `tsv` or `grid` (REQ-605); math as LaTeX (REQ-604); then plain text or
   Markdown from the blocks.

### 3.3 The understood set

docx4j publishes it: `docx4j-core/src/main/resources/org/docx4j/jaxb/understood-namespaces.txt`,
118 URIs (REQ-306). This package's generated `UNDERSTOOD` holds 63. They share 60; 58 of
docx4j's are missing here --- for WordprocessingML `v`, `o`, `w10`, `w16`, `w16cex`, `w16du`,
`w16sdtdh`, `w16sdtfl`, `cx1`, the 2015 chartex namespaces, `r`, bibliography, custom XML, cover
page properties, `xmlPackage` and ink; the rest PresentationML (17), SpreadsheetML, encryption,
web extensions and slicers --- and 3 of ours are not docx4j's (`dcmitype`, `st`, and the XML
namespace, which docx4j's README excludes on purpose).

The list is copied verbatim into `docx4j_py/resources/understood-namespaces.txt`, with a test
asserting it against the sibling's when present (REQ-306). It is the set for the TEXT purpose and,
as docx4j's phase 3 made it for READ, for `mce="resolve"` traversal. **`UNDERSTOOD` stays what
`mce_preprocess` uses**: what the model can parse, deliberately narrower (CLAUDE.md). Measured over
352 packages (553 `mc:AlternateContent` in 65 of them): moving the traversal to the list by URI
changes the selected branch of 3 elements in 2 packages (`waterfall.docx`'s `cx1` twice, and
`numPicBullet-word365-AlternateContent.docx`'s numbering `v`); docx4j's TEXT rule against
`child.py`'s, 7 elements in 4 packages; and running `mce_preprocess` on the list instead of
`UNDERSTOOD` leaves all 65 skipped-content reports identical. Q2.

### 3.4 URIs, not prefixes

REQ-303 needs each `Requires` prefix's namespace URI, which the typed tree does not keep: an
`AlternateContentChoice` holds `requires` as a string, and the declarations in scope are gone
after the parse. The same gap is the save defect of section 2. **Recommended (Q3):** the fork's
parser records, on each `mc:Choice` (and each `mc:AlternateContent` that declares a namespace
itself), the URIs its `Requires` prefixes resolve to, in a slot that equality and serialisation
ignore, and the writer re-declares a prefix a `Requires` names that the output does not declare.
That fixes the save and gives REQ-303 its URIs at once. The alternative, a prefix table, is wrong
for any document that binds a known prefix to another URI, or a new prefix to a known one, as the
probe's case 6 does.

### 3.5 The readers

Phase D puts every shipped reader on the model (REQ-202). That is Q4's first answer. Under the
answer now recommended (Q4), which follows core-ts CR-012's D2, the first and fifth bullets change.
The Office JS readers keep Office JS's text, documented as not the text model: `text` on
`Paragraph`, `Body`, `Range`, `Table`, `TableCell` and `ContentControl`, `Comment.content`,
`TrackedChange.text`, `search`, and `get_text` as core-ts's `getText`. `find` goes with `search`.
`text_of` stays and is not deprecated. `outline()` and `to_markdown` go on the model as below,
and the model itself is reached through `extract`.

- `Paragraph`, `Body` and `Range` `text` and `get_text(view=)` become the model with the WD's
  defaults. This departs from Office JS, whose `body.paragraphs` leaves a text box out and whose
  paragraph text gives U+0002 for each note reference (core-ts check 42, its CR-010 section
  12.6); the departure is recorded (Q4).
- `find` searches the model's text in its default scope. A match in a text box carries a
  `textbox:` address; editing inside a box through a `Range` stays out of scope (Q9).
- `outline()` writes its entries from the blocks (REQ-706), and its counts follow them.
- `to_markdown` becomes `markdown(result)` over the blocks, with REQ-703's list rules (lists
  grouped by abstract numbering, follow-on paragraphs by indent, endnotes `[^e1]`, `HYPERLINK`
  fields as links, decorative pictures without alt text, linked pictures as `![alt]()`, the WD's
  25 monospace fonts, a header row only when the table says so), an `original` view, and
  CriticMarkup written flat.
- `text_of` is deprecated (a `DeprecationWarning` in its docstring and the CHANGELOG, not at
  call time, since the suite turns warnings into errors) and its seven readers move.
- Kept off the model, each documented as not the text model: `TrackedChange.text` (Office JS's
  form, Q6) and the python-docx facade (python-docx's semantics, Q5).

### 3.6 Tests and the oracle

- **The suite**, copied from docx4j at `198d015a3` into `tests/fixtures/text-extraction/`:
  `phase0/` (the five fixtures; the five TSVs --- 285 token rows, 201 findings of which 186 are the
  inspector's, 2 orders, 67 case statements, 12 goldens; the four inspection tables --- 186
  findings and 4 statuses, and 17 findings and 1 status for the HTML reader's fixture; and the
  twelve goldens), `GOLDENS.md` and `INSPECTION.md` (one level up in docx4j), `LevelOwnerTest`'s
  three Word documents, and the understood list. Provenance in `tests/README.md`, as CR-010's
  fixtures have it.
- **The harness**, `tests/text/conformance.py`: a port of `ExtractionConformance` --- REQ-1201's
  normalisation (CriticMarkup's brackets stripped, then whitespace, `\p{Z}`, `*`, `_`, `\`, `{`,
  `}` and `$`; U+200B, tag characters and both hyphens kept; an occurrence inside a longer token
  of the suite not counted), the scenarios (each token row in three views by default, its option
  column's scenario, and the supplemental ones), the combined option (since `a41de133b` and WD 0.15,
  a row's option column may name several options `; ` apart, as in `embedded: true; hidden: true`;
  the adapter sets each one, and the check needs each one's statement), the pending rule (a check
  pends when any statement it needs is not in the adapter's `implemented()`; since `3dcbb3764` a
  findings, order or golden row also needs its scenario's statements, as a token row does, and a
  golden, cut without an HTML reader, needs `html reader absent`), and `verify_golden` (Markdown
  image destinations written `…` on both sides). docx4j CR-050 and CR-053 add, since `3dcbb3764`:
  - the capability is `html reader present` or `absent` (WD 0.27; it was `xhtml importer`), and a
    token row whose option is `html reader present` (cases 36, 44 and 63) runs its default
    scenario under `html reader absent`;
  - `altchunks: false`'s supplemental rows take case 36 and `CHK63_SCRIPTED` as well;
  - an `inspect` findings row is checked once, in `accepted`, an inspection being the package's;
  - each of the first four fixtures' whole inspection is compared with `expected-inspection.tsv`
    by occurrence (REQ-1204), needing REQ-901, 902 and 904 to 907 and `html reader absent`;
  - the fifth fixture's is compared with `expected-inspection-html.tsv`, needing REQ-513 and `html
    reader present` as well, and a row with a `pending` reason pends on its own;
  - REQ-903's counts are checked for each of the four under six outputs (24 checks), needing
    REQ-903 and `html reader absent`;
  - a check needing any of REQ-901 to 907 is reported as the inspector's (CR-050).

  Its report is a TSV in docx4j's form. Advertising docx4j's statements, the port counts docx4j's
  1,584 assertions: 1,565 and 19 pending with `html reader absent`, 1,526 and 58 with `present`
  (this CR's survey mirror, 2026-10-21).
- **The adapter**, `tests/text/adapter.py`, docx4j's `TextsAdapter` protocol:
  `implemented()` and `extract(fixture, view, scenario, format) -> (text, findings)`; it has no
  `inspect` or `counts` (section 3.7). From Phase A on, `implemented()` is docx4j's statements less
  the inspector's: the target, not what this package meets (it was to grow phase by phase), so that
  a gap is a failure the ratchet holds rather than a pending check (section 11.3). It advertises
  `html reader absent` until the HTML reader exists (Q12, Q13). Then
  it takes the capability as a parameter, the reader switched off by an adapter-only hook for the
  absent run, and the suite runs twice. Absent is the gate's run: it checks the goldens and the
  default rows of cases 36, 44 and 63, and leaves docx4j's 19 pending (the 18 checks that need a
  reader, and the fifth fixture's inspection). Present, where html5lib imports, checks the 18 and
  leaves the goldens and the absent run's rows pending.
- **The report and the ratchet.** The adapter's run is committed as the suite's report,
  `tests/text/conformance.tsv`, in docx4j's form with this port's `FAIL` rows (docx4j's run stops
  at the first failure, since none of its checks fails). `tests/text/test_conformance.py` runs the
  suite and compares it with the report check by check, keyed by fixture, case, item, view and
  scenario. Any difference fails the test: a check newly failing, newly passing or newly pending,
  its reason changed, a check added or gone. The message counts the moves (`'FAIL -> PASS': 40`) and
  lists the first. A change that moves a check rewrites the report in the same commit
  (`.venv-fork/bin/python tests/text/conformance.py`). So the pytest suite is green while today's
  checks fail, each phase's progress is in its diff, and no regression passes unseen. The gate of
  Phase D is a report with no `FAIL` row.
- **The agreement test** (REQ-406): the walk against the resolved copy, as section 3.2.
- **CR-010's ten pins** (`tests/revisions/test_content_api.py`) turn to passing tests in Phase D;
  `xfail(strict=True)` makes each one's closing visible.
- **The gate**, at Phase D: in the `html reader absent` run no check fails; the only pending
  checks are docx4j's 19, the inspector's 214 (186 `inspect` rows, 4 inspections and 24 counts;
  section 3.7) and those of any content Q7 defers, each named; 12 of 12 goldens; the agreement
  test at zero differences; `LevelOwnerTest`'s 61 paragraphs as Word numbers them.

### 3.7 Out of scope

- **The inspector** (REQ-901 to 907, REQ-903's counts in a result, and REQ-1204's check of it; WD
  section 9 and section 12.1's inspector class): the paper's later CR (its D29), after the model.
  docx4j's is CR-050 (`Inspect.of`). Its 214 checks of the suite pend here (section 3.6).
  docx4j handed off CR-050 phase 3 on 2026-10-21 (`2853b7280`, unpushed; CR-050 done there), and
  this package acknowledged it the same day (recorded in docx4j `a9ec5680a`). The inspector's
  kinds are frozen, D21 adding `image.svg_text` and `image.metafile_text`. The reference is
  `text-extraction/support/inspect_oracle.py`, a stdlib-only implementation of every rule with
  its own WMF and EMF record reader. The fixtures and the expected rows go with it. A later
  inspector CR here would port it. The acknowledgement listed, kind by kind, what this engine
  already holds the data for, what needs a rule, and what it cannot see: metafile text, embedded
  fonts' glyphs, an HTML chunk's scripts (Q12) and Q7's embedded workbooks and presentations.
  The suite here stays at `198d015a3`, so phase 3's rows reach this package only at a re-cite.
- **Strict OOXML** (REQ-203, converted on load): an engine change, recommended as its own CR
  (Q8); the three Strict fixtures stay pending until then.
- **HTML** (REQ-705): this package has no HTML serialiser.
- **`math: unicodemath` and `symbols: unicode`**: not implemented in docx4j either.
- **Editing through the model**: it reads; the content API edits.

## 4. Phase A: the suite and the baseline

- Copy the suite (3.6) and port the harness and the adapter protocol.
- An adapter over today's paths records the baseline of section 1 as the suite's first report
  (941 pass, 410 fail, 233 pending of 1,584 at `198d015a3`; goldens 3 of 12), committed beside
  it.
- The agreement test, written now against today's views, `xfail(strict=True)` where they
  disagree.

## 5. Phase B: the model in plain text (docx4j CR-046 phase 1)

- `docx4j_py.model.text` (3.1): `TextView`, `TextOptions` with all 24 names, `extract`,
  `TextResult`, `TextBlock`, `TextWarning`, `plain_text`.
- Resolve (3.2 item 1, 3.3, 3.4): the TEXT selection by URI over the list; REQ-304, 305 and 307;
  `mce="resolve"` traversal moved to the list.
- The walk with the three views (3.2 item 2), and the agreement test passing.
- Scope (3.2 item 3): text boxes after their anchor with `textbox:` addresses, notes, headers and
  footers by section, comments in markup, field codes and `field.no_result`, run-level `mc`,
  hidden text through the resolver, placeholders, list labels with REQ-510's owner (the
  emulator's `level_style_id`, which moves labels for any reader of the emulator; Risk 3), the
  `toc` option.
- Serialise (3.2 item 4): characters, lines, `tsv` tables, CriticMarkup flat in `markup`.

## 6. Phase C: the other content (docx4j CR-046 phase 2)

- Fields: `fields: update` with the `\@` and `\#` switches (docx4j's `FormattingSwitchHelper`),
  `FORMTEXT` as Word updates it.
- Bound values (`bindings: bound`), from the custom XML part through `XmlMapping`.
- Charts (a chart's table and description, CR-047's axis titles), SmartArt and WordArt.
- Embedded objects: an embedded WordprocessingML package through this engine; the glossary
  document (REQ-516).
- `w:altChunk`: plain text; HTML, XHTML and MHT (its first `text/html` part) only through the HTML
  reader of Q12 (its block half Q13's), else `altchunk.not_expanded`; RTF not expanded, as in
  docx4j, for want of an RTF reader (REQ-513 since WD 0.23); `altchunks: false`.
- Images: `ImageDescriber`, alt text, decorative pictures.
- Metadata (REQ-515), and `fields_update_on_open`.
- Math as LaTeX (OMML to LaTeX, docx4j's `OmmlToLatex`), ruby's base text.
- Tables `listed` and `grid`; page markers; `layout: indent`.
- Deferred, pending and named (Q7): embedded SpreadsheetML and PresentationML packages, and OLE
  1.0 objects. REQ-512's hidden sheets and slides under `hidden: true`, decided on 2026-10-19, come
  with those packages. Until then, case 53's slides and case 57's workbook stay pending under
  `embedded: true` and `embedded: true; hidden: true`.

## 7. Phase D: every reader on the model (docx4j CR-046 phase 3)

- The readers of section 3.5, and `markdown(result)` with REQ-703.
- The gate of section 3.6; CR-010's pins closed.
- `CHANGELOG.md`: "Added" for `docx4j_py.model.text`; "Changed" for what every text reader now
  returns (text boxes, notes and `[n]`, hidden text out, joined marks, deleted rows gone, the table
  layout, math as LaTeX, charts, `w:altChunk` and bound values read, the TOC omitted, the
  `original` and `markup` views on `get_text`, the Markdown rules, selection by URI, numbering
  labels where two styles carry one level); "Deprecated" for `text_of`.
- The README and `docs/pypi-readme.md` fences, run by `tests/test_readme.py`; CR-003's text
  sections pointed here; CR-006's `read`, `outline`, `find` and `extract_text` given section 10's
  options and REQ-706's blocks, with REQ-801's warnings in their results.

No phase changes saved output, so none needs a Word check; the save fix of Q11 does, on its own.

## 8. Phasing and effort

| Phase | Content | Estimate |
|---|---|---|
| A | the suite, the harness, the adapter, the baseline, the agreement test pinned | 0.5 days |
| B | the model in plain text: resolve, the views, scope, serialise; REQ-510 | 4 days |
| C | the other content: fields, bindings, charts, SmartArt, embedded, `w:altChunk`, images, metadata, math, tables | 6 days |
| D | every reader on the model; Markdown on the blocks; the gate; the documentation | 3 days |
| F | the Office JS readers read Word's branch of a kept `mc:AlternateContent` (Q4 revised, section 15) | 2 days |

docx4j built it in four phases with its own engine underneath; this port has its code, its suite
and its goldens from the start. C is most of the work: docx4j's `FormattingSwitchHelper` alone is
1,092 lines, and its OMML conversion 668.

## 9. Questions

1. **Q1: the package and names**: `docx4j_py.model.text`; `extract(...)` for docx4j's
   `Texts.of`; docx4j's class names (`TextOptions`, `TextView`, `TextResult`, `TextBlock`,
   `TextWarning`); `plain_text` and `markdown` as functions; the WD's option names verbatim.
   *Recommended: as proposed.* core-ts CR-012's D1, decided, keys its options by the same names
   (`note_references`, `field_codes`).
2. **Q2: the understood set**: docx4j's list for the TEXT purpose and for `mce="resolve"`;
   `UNDERSTOOD` kept for `mce_preprocess` (3.3). *Recommended.* Revised for `mce_preprocess` by
   CR-012's Q8 (2026-10-21): on docx4j's CR-051 probes, `UNDERSTOOD` takes `w16cid` where Word
   365 takes `w16sdtfl`, so CR-012 recommends docx4j's list there too.
3. **Q3: how `Requires` gets its URIs**: the fork's parser records them and the writer
   re-declares (3.4), against a prefix table. *Recommended: the fork; it needs a fork release
   before this package releases.*
4. **Q4: `Body.text` and `get_text`**: the model with the WD's defaults, the Office JS departure
   recorded; against keeping Office JS's `body.text` and adding the model beside it.
   *Recommended (revised 2026-10-19): keep the Office JS readers on Office JS's text, as core-ts
   CR-012's D2 decides for core-ts (Jason, 2026-10-21), reader for reader (section 3.5).*
   - The two ports should give one answer. This package's design follows core-ts (CLAUDE.md), and
     this package's content API is held to Office JS's members (CR-003).
   - Q5 and Q6 already keep the python-docx facade and `TrackedChange.text` off the model, because
     their contract is another API's. D2 applies the same rule to the Office JS readers.
   - Jason ruled on 2026-10-19 that core-ts's `textOf` stays, not deprecated. `text_of` is its
     counterpart here.
   - The model behind `Body.text` would make it disagree with what Word's own Office JS returns for
     the same document: a text box left out, and U+0002 per note reference (core-ts check 42).
   - Until WD 0.27 either answer conflicted with REQ-202: under this one, the readers stay outside
     it. Like core-ts, this CR was to ask the specification for an exemption for a reader whose
     contract is an emulated API, documented as that API's text. Q5's note for the WD became that
     request.
   - **WD 0.27 grants it** (portfolio `804007d`, from core-ts CR-012's ask). REQ-202 now lets such a
     reader stay outside the layers "if it is documented as that API's text and not the text
     model, and each of the text model's runs records where its text lies in that API's text, so
     that the two map onto each other". The second condition is new to this CR. Under this answer
     each `TextRun` carries its source (its `w:r` and paragraph) and, where the run gives text to
     Office JS's `text`, its offset in its paragraph's `text`, as core-ts CR-012's `officeOffset`
     (its section 4.4), outside REQ-706's JSON. A run only the model writes (`[n]`, a list label,
     a field code) has none, and text that `text` holds and the model leaves out is the gap
     between two offsets. *Recommended*, with Phase D carrying the map.
   - It narrows the text-extraction paper's section 10 row ("one text model replacing the five"),
     which the user accepted on 2026-10-17. So it is the user's call. D2 is decided for core-ts
     only (its section 6.1): until Q4 is decided alike here, the two ports' paragraph text differs
     for the same document. Taking the first answer here would leave it differing for good.
   - **Revised 2026-10-22 (the user: "revise CR-011 Q4 so that Paragraph.text, get_text and
     search read the branch Word shows").** Office JS's text, measured, reads a kept
     `mc:AlternateContent`'s selected branch: core-ts check 49 (Word 365, Script Lab, 2026-10-20)
     gives `paragraph.text` the selected branch's words at run and paragraph level, never a
     Fallback's, and `body.search` finds a word that is only in a Choice and none that is only in
     a Fallback. So the Office JS readers keep Office JS's text, as decided, and that text now
     includes Word's branch. A text box stays out (check 42). Section 15 is the change, Phase F.
5. **Q5: the python-docx facade**: keeps python-docx's semantics as a declared emulation outside
   REQ-202, with a note proposed for the WD. *Recommended.* WD 0.27's exemption (Q4) covers it, so
   the note is no longer needed. Its mapping condition is met by each run's source: python-docx's
   `Paragraph.text` is its `w:r` and `w:hyperlink` children's text in order, so the facade `Run`
   a model run lies in, or that it lies in none (a run in a `w:ins` or a content control), follows
   from the run's `w:r`. *Recommended:* the source alone, no second offset.
6. **Q6: `TrackedChange.text`**: keeps Office JS's form, documented as describing a change, not
   the text model. *Recommended.*
7. **Q7: embedded SpreadsheetML, PresentationML and OLE 1.0**: deferred and pending, named in the
   gate, until CR-001 D gives this package their models; `olefile` not added. *Recommended.*
8. **Q8: Strict OOXML (REQ-203)**: its own engine CR. *Recommended.*
9. **Q9: `find` and text boxes**: a match in a box carries a `textbox:` address; editing inside a
   box through a `Range` is out of scope here. *Recommended*, and core-ts CR-012's D5, decided, does
   the same: its text-box views are read-only, their mutators throwing, with writes left to a later
   CR that maps a box's paragraphs between its two branches. Under the revised Q4, `find` stays with
   `search` on Office JS's text, and the `textbox:` address comes with `outline()`'s entries.
10. **Q10, ahead of this CR: `get_text(view="markup")`**, which returns neither view today.
    *Recommended: make it raise now, naming this CR, as a "Fixed" entry; Phase B gives it the
    real view.*
11. **Q11, ahead of this CR: the dropped `Requires` declaration** (section 2). *Recommended: fix
    it now as a CR-002 defect, by Q3's mechanism, with its own Word check.*
    docx4j's CR-049 (`ba40c53f4`) fixed the same class of fault in `XmlUtils`, which kept the
    prefixes to declare in state shared between marshals. Ours keeps none (measured 2026-10-19: a
    fragment serialised after `document.xml` declares only what it uses, and `styles.xml` only its
    own `mc:Ignorable` prefixes), so the gap is only a Choice prefix that nothing but `Requires`
    names. The check is `TextsPhase2Test`'s REQ-204 snapshot, which now compares namespace
    declarations. docx4j's CR-052 (phase 1 in progress, 2026-10-21) makes a save declare the
    `mc` prefixes the source bound, per Choice, where Word places them (its D4 (c)). Its phase 2
    hands off to this package, and it touches the same Choice record as Q3 and this question,
    so Q11's fix should take CR-052's placement once it lands.
12. **Q12: the HTML reader** (REQ-513 at WD 0.28). `html5lib`, a pure-Python parser of the
    [HTML] algorithm that builds an lxml tree, as an optional extra (`docx4j[html]`). Without it
    the reader is absent and a chunk raises `altchunk.not_expanded`, the suite's default run, and
    MHT goes the same way. *Recommended*, as docx4j keeps docx4j-ImportXHTML optional. The
    alternatives: `html5-parser` (C and faster, but its wheels are tied to lxml's version); or
    `lxml.html`, section 6's earlier plan, which does not follow the [HTML] algorithm.
    WD 0.16 to 0.28 add three duties. The reader fetches nothing a chunk refers to (REQ-513 since
    0.16), and html5lib fetches nothing. It does not expand a chunk of more than 50,000 elements,
    counted on the tree as parsed (0.27). And it decodes by WD section 9.2's Chunks rule (0.27,
    0.28), so html5lib is to be handed text, not bytes: its own sniffing ([HTML]'s prescan of
    1,024 bytes, then `chardet`'s guess where that is installed, then windows-1252) has no UTF-8
    step and is not Word's.
13. **Q13: the HTML reader's block half and Word's hiding rule** (REQ-513 at WD 0.27 and 0.28).
    A reader is a parser and a reader of its tree as blocks; Q12 chooses the parser.
    - docx4j's second half is docx4j-ImportXHTML. Before it, docx4j's own `HtmlChunks` rewrites
      the chunk's CSS to what Word applies (`HtmlChunks`, `HtmlChunkImport` and `JsoupHtml`, 973
      lines at `198d015a3`). This package has no importer, so the half would be written here: the
      tree read directly as paragraphs, headings, lists with their numbers and tables, a table's
      row groups in source order.
    - Its hidden text is REQ-513's rule: a `style` attribute and the chunk's own `style` rules;
      type, class, id and descendant selectors only; a rule inside any at-rule, or whose selector
      list holds one selector Word ignores, dropped whole; elements outside [HTML401] transparent;
      a host `w:style` defeating a type or class rule (`p`, `h1` to `h6`, `address`, `pre`,
      `code`, a class by its exact name); a descendant's own `display` showing its content.
    - REQ-513's chunk rules moved in six WD revisions over two days (0.23 to 0.28, Word's rounds 5
      to 12). The suite checks the hiding rule only through the fifth fixture's inspection, which
      needs the inspector (section 3.7). Its 18 checks under `html reader present` (cases 36, 44,
      47 and 63) check only whether a chunk is expanded.
    - core-ts CR-012 has no HTML reader either: its adapter advertises `html reader absent`, and
      REQ-513's chunks are its phase 5e's.

    *Recommended:* Phase C ships the reader absent: every HTML, XHTML and MHT chunk raises
    `altchunk.not_expanded`, the gate's run. The reader comes as a phase of its own after D, on
    Q12's answer: the parser, the decoder, the element limit and this block half, with docx4j's
    `HtmlChunks` and `AltChunkHtmlReadingTest` as its oracle. The alternative, the block half in
    Phase C, adds it to the phase that is most of this CR's work, against a rule still moving.

## 10. Risks

1. **Every text reader changes.** `Body.text`, `get_text`, `find`'s matches and offsets, the
   outline's counts and previews and `to_markdown` all return different text on the same document.
   This is a behaviour change in `0.x`, recorded in the CHANGELOG (section 7), and the README's
   examples show text.
2. **The WD is a working draft.** It is at 0.28 and its goldens can be recut, as `1d108c0af`
   recut the probe's four. Port from `198d015a3` and record the commit; a recut is a hand-off like
   any other.
3. **REQ-510 moves numbering labels** for every reader of the emulator, not only the text model.
   This package's parity goldens come from docx4j-core-ts (CLAUDE.md; 46 comparisons at zero
   differences), so they will differ until core-ts ports REQ-510 and regenerates; pin per
   comparison as CR-009 did (`tests/content/goldens.py` `PINS`). core-ts CR-012's D9, decided,
   makes `Emulator.levelStyleId` its phase 1, held to `LevelOwnerTest`'s Word labels. docx4j
   `VERSION_17_3_4` at `5c691fc1f`, which already reads the owner in its labels, gives the same
   answers on all 62 of core-ts's parity goldens (core-ts, 2026-10-21). So no golden is expected
   to move, and a pin would mean this port departs from docx4j's.
4. **The fork.** Q3 changes the fork's parser and writer, so this package cannot release on it
   until a fork release does (`RELEASING.md`: never against an unreleased fork).
5. **No number-format engine.** `\#` switches and chart number formats need one; port docx4j's
   subset, measured against the suite's cases, not a general engine.
6. **Speed.** The walk costs more than today's filter (`get_text` is 1 ms on the probe). Measure
   on CR-003 section 16.4's 200-page document in Phase B, and keep a call under a few times the
   filter's.

## 11. Phase A implementation notes (2026-10-21)

Phase A is done, with Q10. docx4j's text-extraction suite at `198d015a3` is in
`tests/fixtures/text-extraction/`. The port of its harness and an adapter over today's readers
are in `tests/text/`. The adapter's run is committed as the suite's first report: **941 pass,
410 fail and 233 pending of 1,584 checks, 3 of the 12 goldens**, section 1's baseline row for row.
The pytest suite holds every check to that report (the ratchet, 11.3). The agreement test
(REQ-406) pins today's views where they differ from a resolved copy. `get_text(view="markup")`
raises (Q10, 11.5). No reader's text changed.

### 11.1 What landed

| piece | where | what |
|---|---|---|
| the suite | `tests/fixtures/text-extraction/` (32 files) | `phase0/`: the five fixtures, the five TSVs, the four inspection tables and the twelve goldens; `GOLDENS.md`, `INSPECTION.md`; `level-owner/`, `LevelOwnerTest`'s three Word documents; `understood-namespaces.txt`. Each written by `git show 198d015a3:<path>`, byte for byte |
| the harness | `tests/text/conformance.py` (872 lines) | `Engine` (`implemented`, `extract`; `inspect` and `counts` raising, as docx4j's defaults throw), `Output`, `Inspection`, `run(engine, directory)`, `Run` (`checks`, `report()`, `summary()`), `Check`, `read_report`, docx4j's rules as functions (`table`, `normalize`, `count`, `verify`, `verify_order`, `verify_golden`, `image_destinations`, `requirements`, `scenario_requirements`, `row_requirements`, `text_requirements`, `phase`, `compare_inspection`), and `main`, which rewrites the report |
| the adapter | `tests/text/adapter.py` | `TextsAdapter` over `get_text` and `to_markdown`; `DOCX4J_IMPLEMENTED`, docx4j's set at `198d015a3`; `INSPECTOR` |
| the report | `tests/text/conformance.tsv` | 1,584 rows in docx4j's columns |
| the ratchet, the gate, the copy | `tests/text/test_conformance.py` | 23 tests |
| the agreement test | `tests/text/test_agreement.py`, `agreement-docx4j.tsv` | 24 tests over the copied fixtures and 2 of the pins themselves; 294 over docx4j's (`slow`) |
| Q10 | `model/content/text_model.py` (`require_view`, `VIEWS`), `paragraph.py`, `body.py`, `range.py`; `tests/agent/test_errors.py` | the view refused, two tests |

### 11.2 The harness as ported

Every rule of section 3.6, in docx4j's order: the token rows in three views under the default and
the option scenarios, the supplemental scenarios, the findings (an `inspect` row in `accepted`
only), the orders, the inspections and REQ-903's counts, the goldens. Ported from the Java where
the survey's mirror differed, none of which moves a count:

- a check's statements are a set, as Java's `LinkedHashSet`; the mirror's list named a statement
  twice in a pending reason when the row and its scenario both needed it;
- an empty marker raises, as in Java (the mirror looped);
- the tables are split as `Files.readAllLines` splits them (`\r\n`, `\r`, `\n`), and the goldens
  read as bytes, without Python's newline translation;
- an expectation is a count when it is ASCII digits (Java's `\d+`), not `str.isdigit()`;
- an unknown scenario or marking is an error of the suite and stops the run, as Java's
  `IllegalArgumentException` does;
- a golden's readiness is checked before its output is read.

With docx4j's whole statement set advertised and empty output behind it, the port makes docx4j's
1,584 checks: 1,565 and 19 pending under `html reader absent`, 1,526 and 58 under `present`
(`test_the_port_makes_docx4js_checks`). Under `absent` the 19 are case 36's, 44's and 47's rows
under `html reader present` (15), case 63's (3) and the fifth fixture's inspection.

**The one departure.** docx4j's run stops at the first failing check, so its report has `PASS`
and `PENDING` rows only. The port records a failing check as a `FAIL` row with its reason and goes
on. An engine's exception fails every check that reads its output, not the run. The reasons are
the port's own: Java's assertion messages without the context the row's columns give
(`expected 1, got 0`, `total: expected 2, got 1`, `marking del: expected 1, got 0`), and for a
golden the lines that differ (`normalised text differs: 43 lines of the golden missing, 22
extra`). Only a `CheckFailed` is a failure: an error of the suite or the harness stops the run.

### 11.3 The report and the ratchet

`tests/text/conformance.tsv` is the run's report: docx4j's header (`fixture`, `case`, `item`,
`view`, `scenario`, `status`, `reason`) and one row per check. `.venv-fork/bin/python
tests/text/conformance.py` runs the adapter, rewrites it and prints docx4j's summary line with the
failures and goldens added (`CR-046: 941 assertions passed, 410 failed, 233 pending of 1584;
goldens 3 of 12`).

`test_the_run_is_the_committed_report` runs the suite once per module and compares it with the
file by check. Any difference fails, and the message counts the moves and lists the first 25, with
the command that rewrites the report. Last, it compares the file's bytes with the run's, so the
order is held too. Flipping one row of the file by hand gives `1 checks differ from
conformance.tsv ({'PASS -> FAIL': 1})` and the row. A change that moves a check rewrites the
report in the same commit, and the diff is the progress. This is written into section 3.6.

**The adapter advertises the target** (section 3.6 had it grow phase by phase): docx4j's
statements at `198d015a3` less the inspector's (REQ-901 to 907, REQ-1204), with `CR-047` and
`html reader absent`. So every check docx4j makes is made here. Today's gaps are `FAIL` rows, and
a phase's progress is checks going from `FAIL` to `PASS`. Grown phase by phase, 410 failures
would instead have been pending checks indistinguishable from docx4j's, and each phase would have
changed `implemented()` and the report together. Only docx4j's 19 and the inspector's 214 pend,
and `test_only_docx4js_pending_checks_and_the_inspectors_pend` holds that. For Q7's deferred
content, Phase C's adapter will need a way to pend a named check, since its statement (REQ-512)
also covers content that is read.

**The adapter's `markup` plain text** is the filter `get_text(view="markup")` applied before
Q10: each paragraph's `text_of_view(view="markup")`, joined as `Body.get_text` joins. The
report is therefore section 1's measurement. Read through the call that now raises, 303
passing checks would fail, giving 638 pass, 713 fail and 2 of 12 goldens. That would measure the
refusal, not the readers. Phase B replaces both with the model's markup view.

### 11.4 The agreement test

The walk is `Body.get_text(view=)` on the fixture as loaded. The reference is a `dry_run` copy
resolved by `revisions.accept_all` (`accepted`) or `reject_all` (`original`), read by
`get_text()`. docx4j loads its copy a second time because a save drops probe case 6's `Requires`
prefix. A trial copies the parsed tree and saves nothing, so it keeps the prefix. On three
fixtures (`revisions-word15`, `review-for-word`, the probe) the trial's text equals a reloaded
copy's, and the real package's listing is unchanged after it. Only the default scope is read:
today's `get_text` reads the main story and takes no options, where docx4j's test also reads every
story with hidden text and field codes.

- **The copied twelve** (CR-010's eleven and the probe), 24 tests: 14 agree, and 10 are
  `xfail(strict=True)` with what differs, 6 in `accepted` and 4 in `original`, as section 1
  measured. Each is a deleted (in `original`, an inserted) paragraph mark not joined, or a
  paragraph the resolution removes left as an empty line. In `original` there is also `cr043`'s
  C01 inserted field result and C16 inserted link, kept. CR-010's five `get_text` pins in
  `tests/revisions/test_content_api.py` are the `accepted` half of these and stay as they are.
- **docx4j's fixtures**: the 147 revisions and text-extraction `.docx` at `198d015a3`, where the
  sibling checkout has them, `slow` (about 6 s). Section 1's 139 was an earlier count, and the
  working tree at `4b9c71288` has 153. The list and today's outcomes are committed in
  `agreement-docx4j.tsv`: 80 differ in `accepted`, 81 in `original` (section 1: 79 and 80 of
  139). A fixture docx4j adds later is not read, so the sibling moving on cannot fail the suite. A
  fixture docx4j changes in place can.

### 11.5 Q10

`require_view` in `model/content/text_model.py` refuses a view `get_text` cannot read. It is
called by `Paragraph.get_text` and `text_budget`, `Body.get_text` (and through it `text_budget`)
and `Range.get_text`. `"markup"` raises `ContentError` with the code `text.no_markup_view`
("get_text has no markup view yet: CR-011 Phase B gives it the text model's"). Its hint names the
views the reader takes and `to_markdown(view='markup')`. A range raises its `SpanError` for
`original` first, as before. Any other name raises `text.view_invalid`: the filter shows a
revision holder only in its own view, so every other name, `current` included, gave neither view
too. `current` becomes REQ-401's synonym in Phase B. In 0.1.0 a paragraph's and a body's markup
view was the text outside every tracked change (5,093 characters on the probe), and a range's was
the accepted text. No test or readme fence called it. The adapter is the one reader of that
filter, through `text_of_view` (11.3), and docx4j-mcp-py calls neither. CHANGELOG "Fixed".

### 11.6 The numbers

- The report: 941 pass, 410 fail, 233 pending of 1,584. By scenario: the default scenarios 669,
  219 and 190; `markdown` 2, 4 and 4; the options 270, 187 and 39. The 233 pending are 18 checks
  needing an HTML reader, the fifth fixture's inspection, the 186 `inspect` rows, 4 inspections
  and 24 counts. Goldens: 3 of 12, the external fixture's three plain-text views. The other nine
  differ by the lines section 1 gives.
- The harness: the adapter's run takes 1.35 s standalone, the model's import included.
- The tests: `tests/text` adds 49 items to the fast suite (39 pass, 10 `xfail`) in 1.7 s
  wall-clock, and 294 to the slow one (133 pass, 161 `xfail`) in about 6 s. `test_errors.py`
  gains two. The fast suite is 2,558 passed, 8 skipped and 20 xfailed, 334 deselected (82 s);
  the whole suite 2,731 passed, 8 skipped and 181 xfailed (105 s).

### 11.7 What Phase B needs

- The adapter's `extract` on `docx4j_py.model.text` (`extract`, `plain_text`, and Markdown in
  Phase D), its statement set unchanged; every commit that moves a check rewrites the report.
- The agreement test's walk becomes `extract`, and its pins close.
- `LevelOwnerTest`'s documents and the understood list are copied and not yet read: REQ-510's
  owner and section 3.3's `docx4j_py/resources/understood-namespaces.txt`.
- `get_text(view="markup")`'s refusal becomes the model's markup view.

## 12. Phase B implementation notes (2026-10-21)

Phase B is done, with Q3 and, at the coordinator's word, Q11 by docx4j CR-052's rule (phase 1,
`5df2e253e`). `docx4j_py.model.text` is the text model in plain text, the port of docx4j
CR-046 phase 1 after core-ts CR-012's phase 2. The adapter reads it, and the report moves from
941 pass, 410 fail and 233 pending to **1,146, 205 and 233 of 1,584, goldens 6 of 12**: 205
checks from `FAIL` to `PASS` and none the other way (12.9). The agreement test (REQ-406) holds
the walk equal to the resolved copy in **636 of 636** comparisons. `get_text(view="markup")` is
the model's view. The Office JS readers keep Office JS's text (Q4).

### 12.1 What landed

| piece | where | what |
|---|---|---|
| the fork | `~/git/docx4j-xsdata`, stage 10 (its `docs/fork/CHANGES.md`; report 19 in `docs/UPSTREAM.md`) | `ParserConfig(on_namespaces=..., namespaces_for=...)`: each object built for the elements named, typed or a wildcard, handed over with the namespace declarations in scope there. Nine tests in its `tests/fork/test_parser_namespaces.py` |
| the record | `docx4j_py/child.py`, `runtime.py` | `BoundPrefixes`, `record_bindings` (the hook, for `mc:Choice`), `source_bindings`, `set_source_bindings`; the TEXT selection: `understood_namespaces`, `requires_namespaces`, `understands`, `mce_branch` |
| the understood list | `docx4j_py/resources/understood-namespaces.txt` | docx4j's 118 URIs, copied verbatim from the suite's copy (REQ-306) |
| the writer | `docx4j_py/mc_declarations.py` (new, in `KEEP`), `openpackaging/parts/xml_part.py`, `fragments.py`, `model/revisions/islands.py` | docx4j CR-052's placement, on every marshal path (12.2) |
| the package | `docx4j_py/model/text/`: `__init__.py`, `options.py`, `result.py`, `trees.py`, `view.py`, `extractor.py`, `plain.py` (3,596 lines) | `extract`, `extract_resolved`, `plain_text`, `normalize`; `TextOptions` (the 26 names), `TextOptionError`, `TextView`; `TextResult`, `TextBlock`, `TextRun`, `TextRow`, `TextCell`, `TextField`, `TextNumbering`, `TextRevision`, `TextWarning`, `TextComment` |
| REQ-510 | `model/listnumbering/emulator.py`, `definitions.py` | `Emulator.level_style_id`, docx4j `Emulator.levelStyleId` (`d9acdf5e2`); `style_linked_elsewhere` asks it, a paragraph naming no style taken in the default style; `number(None)` resolved as an empty `w:pPr` (`c283dffdb`) |
| the readers | `model/content/text_model.py`, `paragraph.py`, `body.py`, `range.py`, `enums.py` | `get_text(view="markup")` and `text_budget` on the model; `current` read as `accepted`; `TextView` gains `"markup"` |
| the engine's share | `model/revisions/lister.py`, `model/content/trial.py` | `Lister(islands=)`, so the plan and the walk see one typed text box; a trial part's `read_contents`, `header_parts` and `footer_parts` |
| tests | `tests/text/test_texts.py` (46), `tests/openpackaging/test_mc_declarations.py` (18), `tests/openpackaging/test_mce.py` (+3), `tests/agent/test_errors.py` | the rows of `tests/README.md` |
| the harness | `tests/text/adapter.py`, `conformance.tsv`, `test_agreement.py`, `agreement-docx4j.tsv` | the adapter on the model; the report rewritten; the agreement test on `extract` |

### 12.2 Resolve, and the save (Q3, Q11)

- **The record.** The runtime parser hands each `mc:Choice` it builds, typed
  (`AlternateContentChoice`) or a wildcard (`AnyElement`), to `record_bindings`. Its `Requires`,
  `mc:Ignorable` and `mc:MustUnderstand` values are replaced by a `BoundPrefixes`: the same
  string, with each prefix it names and the URI that prefix had in scope. Equality and
  serialisation see the string. A move keeps the record, `deep_copy` and `copy.deepcopy` carry it
  (the value copies as itself), and an edit that sets a plain string drops it. The probe's eleven
  Choices load with their bindings, `shape2010` to the wps namespace among them.
- **The selection** (REQ-301 to 306, docx4j's TEXT purpose, which is also its `McMode.READ`): the
  first Choice whose `Requires` names at least one prefix, each bound to a namespace in docx4j's
  list; else the Fallback; else the first Choice. A Choice with no record (built in code, or its
  `Requires` set since) is matched through docx4j's prefix table, as REQ-303 permits. The
  traversal's `mce="resolve"` uses it, so `text_of`, `walk` and the content API read the same
  branch as the model (Q2). `mce_preprocess` keeps `UNDERSTOOD`, what the model can type.
- **REQ-307.** `mce.branches_differ` where the selected branch and another both hold text and
  differ after REQ-1201's normalisation, a branch's text being every `w:t`, `m:t` and `a:t` in
  it in every view, hidden or not (WD 0.29, the user's ruling for core-ts). The warning names the
  branch (`mc:Choice Requires="wps"`, `mc:Fallback`), never the text. Off under `warnings: false`.
- **The save** (docx4j CR-052 sections 4 and 9, its D3 and D4 (c)). After the cleanup and the
  root's `mc:Ignorable` reconciliation, every prefix an `mc:Choice`'s `Requires`, `mc:Ignorable`
  or `mc:MustUnderstand` names is placed: the source's binding, else the table's; a prefix the
  table binds as the source did goes on the root, as Word puts its fixed set; any other (a private
  `zz9`, the alias `shape2010`, a table prefix the source bound otherwise) goes on the Choice,
  ahead of `Requires`; a prefix the root already binds alike (its `mc:Ignorable` names it) is
  declared once, on the root; nothing goes on `mc:Fallback`. Where the Choice's content uses the
  namespace the prefix has in scope, the content's binding stands and a warning is logged (D3's
  clash). A prefix neither binds is left undeclared, with a warning. The root's `mc:Ignorable`
  prefixes now take the source's binding before the table's, as D3 orders it (it was the table's
  first). `fragments.to_xml` follows the same rule, and the typed islands of CR-010 declare a
  Choice's recorded bindings when they re-parse a text box, so a nested Choice keeps them.
- **Measured**: a re-marshalled save of the probe declares `zz9` on its three Choices and
  `shape2010` on its one, `<mc:Choice xmlns:zz9="http://example.com/zz9" Requires="zz9">`, and
  `wps` and `w14` on the root; `waterfall.docx` declares `cx1` on the root. Before, all of these
  were undeclared, which Word 365 repairs (docx4j CR-052 phase 0). Every case of docx4j's
  `McSourcePrefixesTest` that is not about Java's marshal paths is ported
  (`test_mc_declarations.py`, 18 tests), each checking every prefix an `mc` attribute of the saved
  `document.xml` names against the declarations in scope where it is named. **This changes saved
  output**, for a document whose Choices name such a prefix only. `out/acceptance/` regenerated is
  canonically identical to the artefacts before (`scripts/compare_acceptance.py`, 13 artefacts),
  and none of them has such a Choice. The Word check of the two fixtures saved is the user's,
  staged in `out/cr011-q11-word-check/` with a README. **It passed on 2026-10-21** (the user:
  "q11: both open fine"): the probe and `waterfall.docx` re-saved open with no repair prompt.
- **The cost.** The rendered tree is scanned for `mc:Choice`; the object tree is read for the
  records only once a binding the table does not make has been recorded in the process (`zz9`,
  `shape2010`), since until then the table gives every answer the records would. A declaration
  lxml would strip (an alias of a namespace in scope) is written with a placeholder namespace and
  given its URI in the serialised bytes.

### 12.3 The views (REQ-401 to 406)

As docx4j and core-ts build them: **the walk is the implementation, the copy its oracle.**
`view.py` is docx4j's `ResolvedView` as core-ts ports it over its own resolution, here over
CR-010's engine: the `Lister` over every story with revision markup (the stories read with
`read_contents`; a part whose bytes hold none is not listed), and each change planned as the
`Resolver` would resolve it --- a revision taken away, the moves replayed story by story last
half first (`_Moves.pieces_of`), an inserted field rejected whole, a row or cell, an equation's
revision or structure, a paragraph mark, an inserted control unwrapped, and the comments a
resolution takes whole (CR-010 section 12.9's rule, `a8bc6cda9`, U23's spans included). A text
box kept as a wildcard is one typed island shared by the listing and the walk
(`Lister(islands=)`), so the plan's identities are the walk's. `markup` plans nothing.

**The agreement test** (`test_agreement.py`): the walk, `plain_text(extract(pkg, view=))`, against
the same fixture loaded again, resolved by `revisions.accept_all` or `reject_all`, and read by
`extract_resolved` with no view logic; in two scopes, the default and every story with hidden
text and field codes. The copied twelve, 48 tests, and docx4j's 147 (`slow`), 588: **no
difference** (before: 10 of 24 and 80 and 81 of 147 in the default scope alone). The reference is
a second load, as docx4j's is, not a `dry_run`: **found on the way**, a trial keeps a comment the
resolution empties (`review-for-word`'s comment 0, rejected), which the walk drops and Word drops,
because the engine's flush does not take a trial's copy of a part for a part (`Removals._is_part`
asks for a `Part`). Not fixed here: a fix lets the flush drop a trial's emptied comment parts,
which must not reach the real package's relationships, and wants its own test.

### 12.4 Scope as built (WD section 5)

core-ts's `Extractor` over this model, its rules unchanged:

- **Text boxes** after their anchoring paragraph, or after the table it is in, once, addressed
  from it: `body/3/textbox:0/0`, `body/68/0/0/0/textbox:0/0` (REQ-518, REQ-519, D4). A box is found
  in a run's `w:drawing`, `w:pict` or `w:object`, typed (`wps:txbx`'s `TxbxContent`) or a wildcard
  `w:txbxContent` read as its island; in an `mc:AlternateContent`, in the selected branch only.
  `text_boxes: false` reads none. A box that would lose content as typed raises
  `content.not_read` at its anchor.
- **A kept branch** is a wildcard tree in this model (CR-002's lossless MCE; CR-012 not
  implemented): run content in it (`w:t`, `w:tab`, `w:sym`, a field's characters) is read by
  element name, and a wildcard run, revision, link or control at paragraph or block level is
  parsed typed for the reading.
- **Notes** in reference order, `[n]` at the reference and at the note's start, a note whose
  reference the view removes absent (REQ-402, REQ-503); `notes` and `note_references`.
- **Headers and footers** by section under `include` (REQ-502): each section's first, default and
  even ones as `w:titlePg` and `w:evenAndOddHeaders` show them, before and after its blocks, each
  part once, at `header:<rId>/<n>`.
- **Comments** at their references in `markup` and under `comments: true`, replies after their
  parent, never one with no anchor (REQ-504); `resolved` from `commentsExtended`.
- **Fields**: the cached result; the code under `field_codes` as `FIELD_CODE` runs; a field with
  no result `field.no_result` (REQ-505); a HYPERLINK field's target on its runs. The table of
  contents (a TOC field) left out whole under the default `toc: false`, its code too, and a
  paragraph with nothing else in it no paragraph (REQ-520).
- **Hidden text** left out, through the `PropertyResolver` (the run's `w:vanish`, else its
  paragraph and character styles'); a paragraph hidden mark and all has no line (REQ-507).
- **Placeholders** read and flagged; `placeholders: false` leaves them out (REQ-508).
- **List labels** under `list_labels`, counted per story by the emulator with REQ-510's owner;
  every numbered block has its `TextNumbering`, its `value` the label's count (REQ-510, REQ-706).
- **Equations** as their `m:t` text inside `$...$` or `$$...$$`, REQ-604's form for one not
  converted, and in `markup` the Reject All form deleted and the Accept All form inserted; ruby as
  its base.

### 12.5 Serialise (WD section 7.1)

`plain.py`: a line per paragraph, U+000A after each; section 6.1's characters (`w:tab` and
`w:ptab` a tab, `w:br` and `w:cr` a line feed, the two hyphens U+2011 and U+00AD, `w:sym` its
character), as stored (REQ-602); a table a row a line, its cells a tab apart and a cell's
paragraphs a space apart (REQ-605's simple form); `[n]`; and in `markup` CriticMarkup written
flat, consecutive runs of one kind of revision under one marker and never one marker inside
another, so a deletion inside an insertion reads `{++TC14_INS_KEPT ++}{--TC14_DEL_IN_INS--}` as
the golden has it (section 2's Markdown nests it); `markup_detail` adds `author:` or `author,
date:` and splits the markers by author or date.

### 12.6 The readers (Q4, Q10)

- `Paragraph.get_text(view="markup")` and `text_budget` are the model's markup view of the
  paragraph's own runs; `Body.get_text(view="markup")` and its `text_budget` the model's view of
  the body's content, text boxes after their anchors. Content is read in the package's view.
  Content with no package is read as it stands (no styles; in `accepted` and `original` run-level
  revisions only).
- `current` is `accepted` in every `get_text` (REQ-401's synonym, as section 11.5 said).
- In every other view the Office JS readers keep Office JS's text, as decided (Q4): `text`,
  `get_text`, `search`, `TrackedChange.text`, `text_of`, and the python-docx facade.
- Each `TextRun` carries its `source` (its `w:r` and `w:p`) and, where its text is in the content
  API's `get_text()` (`get_text(view="original")` in the original view), its `office_offset`
  there; within a run the two agree character for character, and runs stay apart where
  `get_text()` holds text between them (hidden text is the gap). Neither is in `to_dict()`.

### 12.7 REQ-510 and the parity goldens (risk 3)

`level_style_id` gives `LevelOwnerTest`'s owners, and the emulator and the text model number its
three Word documents as Word does, all 61 paragraphs (27 differed, section 1). The 46 parity
goldens of CR-002 still give zero differences (277 tests, `PINS` empty), as core-ts's 62 and
docx4j `5c691fc1f` led us to expect: none of their fixtures has a style that loses a level.
CR-010's Word-save comparisons are unchanged (68 tests).

### 12.8 Departures, with why

1. **Phase C's values are accepted, not refused** (section 3.1 said an unimplemented value
   raises). `fields: update`, `embedded: true`, `alt_text: true`, `metadata: true`, `tables:
   grid`, `page_markers: true` and `layout: indent` are read as their defaults, and listed in
   `NOT_YET` and `TextOptions.not_yet()`, as REQ-1001 allows ("document any it does not
   implement"). A refusal would fail every check under those scenarios that the default's text
   meets: 78 that pass today, against the ratchet's rule. `math: unicodemath` and `symbols:
   unicode` are refused, as docx4j refuses them.
2. **The record is the attribute's value**, a `str` subclass, not a slot (section 3.4): no
   generated class changes, so no regeneration, and the value carries it through a move and a
   copy and loses it with an edit, which is CR-052's D2 without a copy listener.
3. **The clash case of `McSourcePrefixesTest` does not arise as written.** The serialiser writes
   SpreadsheetML content as `s:`, not `x:`, so declaring Excel's `x` on the Choice rebinds
   nothing, and the test asserts that. The clash itself is tested with the root binding `zz9` one
   way and the Choice's content using it.
4. **`tables: auto` writes every table tab-separated**: the listed form is Phase C's.
5. **Equations are their text**, not LaTeX (Phase C).
6. **A range has no markup view**: `Range.get_text(view="markup")` raises `SpanError`
   (`range.no_markup_view`), as `original` does, its offsets being the accepted view's; the
   paragraph has it.
7. **The agreement reference is a second load**, not a `dry_run`, for 12.3's trial defect.
8. **A wildcard paragraph or block in a branch is parsed typed for the reading**, so the plan
   knows nothing in it. None of the 119 packages of `samples/` and `tests/fixtures/` has one:
   their branches hold run content (`w:drawing` 39, `w:pict` 19, `w:t` 2) and typed runs (14).
9. **Run boundaries follow `office_offset`**, as core-ts's (its departure 3).
10. **The glossary is not read** under `include: glossary` (REQ-516, Phase C).

### 12.9 The numbers

- **The report**: 1,146 pass, 205 fail, 233 pending of 1,584; goldens 6 of 12 (the three plain
  text goldens of the structure fixture join the external fixture's three). 205 moves, all `FAIL
  -> PASS`, by cause: text boxes and `mc:Choice` content 76 (cases 1 to 8, 17, 25 and its
  header's box, 48, 59, 67); field codes 36 (19, 20, 49, 61); hidden text 18 (18, 58); headers and
  footers by section 18 (26, 63); notes 16 (17, 27, 63); the views 9 (9, 10, 13, 14, 15);
  equations as their text and ruby 9 (32, 37); comments 8 (28, 63); the table of contents 6 (55);
  placeholders 3 (22); `fields_update_on_open` in the metadata 3 (50); goldens 3.
- **What fails** (205) is Phase C's and D's: embedded packages 69 (cases 39 to 42, 53, 57, 63,
  66); `w:altChunk` and its warnings 45 (36, 43 to 47, 63); charts and SmartArt 27 (34, 35);
  metadata 18 (38, 52); bound values 12 (23, 67); alt text 9 (33, 66); the glossary 9 (54);
  legacy WordArt and the watermark 6 (25, 63); Markdown 4 (21); and 6 goldens: the three
  Markdown ones (Phase D) and the probe's three plain ones, which differ only by Phase C's lines
  (17 of the golden's missing, 3 written in their place: the bound value's surface, the table in
  the simple form and the equation as its text), as
  `test_the_probes_goldens_differ_only_by_phase_cs_lines` holds.
- **The suites**: the fast suite 2,660 passed, 8 skipped, 10 xfailed and 628 deselected (88 s);
  the whole suite 3,288 passed, 8 skipped and 10 xfailed (133 s). The 10 are CR-010's pins
  (Phase D). The fork's suite 1,604 passed.
- **Unchanged**: `scripts/bench_marshal.py`, `Symbols.docx` parse 192 and 193 ms against 192 and
  200 at `d8e4b8b`, marshal 248 and 249 ms against 248 and 250, the 14 main parts 233 to 235 and
  297 to 298 ms against 233 to 243 and 297 to 301; `import docx4j_py` 644 to 648 ms against 646
  to 678, `docx4j_py.model.text` not among what it imports (`tests/test_performance.py` passes);
  `scripts/roundtrip.py` 50 of 50 parts canonically identical.
- **The cost** (CR-003 section 16.4's 200-page document, the median of five, the body
  unmarshalled): `get_text()` 9 ms; `plain_text(extract(pkg))` 34 ms in `accepted`, 33 in
  `original`, 29 in `markup`. With 3,900 tracked changes (`replace_text` tracked): `get_text()` 20
  ms, the model 94 to 100 ms, the plan among it. So three to five times the filter, within risk
  6's bound. The probe: 0.8 ms and 15 ms. The first profile spent two thirds of the time
  resolving each paragraph mark's properties; they are cached by the two style ids.

### 12.10 What Phase C and D need

- Phase C: section 6's content, each value of `NOT_YET` turned on with its statement, the
  adapter unchanged; the probe's goldens then exact.
- Phase D: `markdown(result)` and the Markdown readers on the blocks; `outline()` on the model;
  CR-010's ten pins; CR-006's tools.
- The `dry_run` comment defect of 12.3.
- A Word check of the probe and `waterfall.docx` saved after their main parts are unmarshalled
  (12.2), the user's: **passed 2026-10-21**.
- The fork's stage 10 needs a fork release before this package releases on it (risk 4).

## 13. Phase C implementation notes (2026-10-21)

Phase C is done: the port of docx4j CR-046 phase 2 at `198d015a3`, each value of section 6 acted
on, `NOT_YET` empty. The report moves from 1,146 pass, 205 fail and 233 pending to **1,284, 7 and
293 of 1,584, goldens 9 of 12**: 138 checks from `FAIL` to `PASS`, 60 from `FAIL` to `PENDING`
(Q7's deferred content, named), none the other way (13.7). Every plain-text golden is exact; the
7 failures are Markdown's, Phase D's. The agreement test stays at no difference (636 of 636).

### 13.1 What landed

| piece | where | what |
|---|---|---|
| metadata | `model/text/metadata.py` | docx4j's `Metadata`: `core`, `app` and `custom` maps in `TextResult.metadata` (REQ-515) |
| images | `model/text/images.py` | `TextImage`, `ImageDescription`, `ImageDescriber`, `alt_text_describer` (docx4j's `AltTextDescriber`) |
| SmartArt, DrawingML text | `model/text/drawings.py` | docx4j's `Drawings`: an `a:p`'s text; SmartArt's points in reading order |
| bound values | `model/text/bound_values.py` | docx4j's `BoundValues`, over `model/customxml/` and the property stores |
| one level down | `model/text/embedded.py` | docx4j's `AltChunks` and `Embedded` (its WordprocessingML half): `chunk_blocks`, `package_blocks`, `decode` |
| fields: update | `model/text/field_values.py` | docx4j's `FieldValues` |
| the walk | `model/text/extractor.py` | what a run anchors as docx4j's `Anchored` items (`_Box`, `_WordArt`, `_Chart`, `_SmartArt`, `_EmbeddedObject`); `read_nested`; `w:altChunk`; the glossary; bookmarks and the first reading; computed fields and FORMTEXT; page markers; equations by `OmmlToLatex` |
| plain text | `model/text/plain.py` | docx4j's `PlainText` whole: the listed and grid tables, charts, `[image: ...]`, `[page n]`, the metadata block, `layout: indent` |
| the block model | `model/text/result.py`, `options.py` | `TextRun.image`, `description`, `page_number`; `TextBlock.decorative`, `simple_form`, `chart`, `chart_title`, `chart_header`, `chart_rows`; `image.not_described`; `TextOptions(image_describer=)`; `NOT_YET` empty |
| charts | `model/chart/` (`model.py`, `chartex.py`, `table.py`, `description.py`, `formats.py`) | the reading half of docx4j's `org.docx4j.model.chart`: `ChartModel`, `ChartExModel`, `ChartTable`, `ChartDescription` (docx4j CR-047), and a measured subset of POI's `DataFormatter` |
| math | `model/math/` | docx4j's `OmmlToLatex` and `LatexSymbols` |
| fields | `model/fields/` | docx4j's `FldSimpleModel`, `FormattingSwitchHelper`, `DateFormatInferencer`, `NumberExtractor`, `DocPropertyResolver`, `DocVariableResolver` |
| Java's formats | `model/java_text.py` | the `DecimalFormat` and `SimpleDateFormat` the chart formats and the field switches share |
| the harness | `tests/text/conformance.py`, `adapter.py` | `Engine.deferred` and `adapter.DEFERRED` (13.5); the metadata flattened into the findings as docx4j's adapter flattens them |
| tests | `tests/text/test_texts_phase2.py` (48), `test_charts.py` (92), `test_fields.py` (287), `mini_docx.py` | docx4j's `TextsPhase2Test`, `ChartDescriptionTest`, `ChartTableTest` and its `FormattingSwitchHelper` tests, where not deferred; docx4j's `MiniDocx` |
| fixtures | `tests/fixtures/charts/` (36 fixtures, 33 JSON goldens), `text-extraction/formtext/`, `text-extraction/extraction-probe.word-resave.docx` | copied from docx4j at `198d015a3`, byte for byte |

Every new module is under `docx4j_py/model/`, which `KEEP` holds whole, and each is imported on
first use: `import docx4j_py` imports none of them.

### 13.2 The content as built

docx4j's rules, unchanged, each with its statement:

- **Metadata** (REQ-515): the core properties in the schema's order, the application properties
  that describe the document (the statistics left out), each custom property's typed value in the
  part's order (a vector, a blob or a class id left out), an empty string left out; in the
  metadata whatever the options, and in the text only under `metadata: true`, a line a kind
  before the body. Nothing is unmarshalled.
- **Pictures** (REQ-514): each an `IMAGE` run with its `TextImage` (`wp:docPr`'s or a nested
  `pic:cNvPr`'s `descr`, `name` and `title`, `adec:decorative`, VML's `alt`, the image part, an SVG
  picture's `asvg:svgBlip` part where `a:blip` names none) and the describer's description; the
  title never read. Plain text writes `[image: ...]` or `[image, generated: ...]` under `alt_text:
  true`; a decorative picture has no description, and a paragraph of decorative pictures alone is
  a `decorative` block (REQ-706). A describer that raises gives `image.not_described`, the image
  part's name its selector.
- **Objects** (REQ-512): each `w:object` an `OBJECT` run with its ProgID; an ActiveX control one
  with none. Under `embedded: true`, one reading level only, an `EmbeddedPackagePart`, or for a
  ProgID `Package` a binary part, follows its anchor at `body/n/object:k`.
- **Legacy WordArt** (REQ-511): a `v:shape`'s `v:textpath/@string` a one-paragraph text box
  (`body/69/textbox:0/0`), under `text_boxes`; case 63's watermark in its header so.
- **Charts** (REQ-511, CR-047): a `CHART` block after the anchoring paragraph (or its table),
  `body/103/chart:0`, in the anchoring story; its description's kind, axes and secondary series
  under `chart` in its JSON, its title, header and rows; plain text writes the note before the
  title, then the rows a tab apart. A chartex chart is read where its `mc:Choice` (`cx1`,
  understood) is selected. A chart part that cannot be read raises `content.not_read`.
- **SmartArt** (REQ-511): the data part's points from `doc` along `parOf` by `srcOrd`, the text
  of `node` and `asst` points, a line a paragraph, `SMARTART` blocks at `body/105/smartart:0/i`;
  `content.not_read` where the data part cannot be read.
- **Bound values** (REQ-509): a plain-text control's `w:dataBinding` evaluated with its own prefix
  mappings against the custom XML part its store item id names, or the core or application
  properties by their well-known ids; the first node's string value. Where it resolves and the
  control does not show its placeholder over an empty value, the control is its value (a
  block-level one a paragraph), and `sdt.binding_differs` is raised where the surface differs, in
  either `bindings` value; `bindings: surface` reads the surface.
- **`w:altChunk`** (REQ-513): by the declared content type, then the extension, then the bytes.
  Plain text as its lines, decoded by its BOM, else as UTF-8 where it is valid, else
  windows-1252; a WordprocessingML or Flat OPC chunk read one level down at
  `body/n/altchunk:0/...`, its own chunks not expanded (case 47's HTML one raises at
  `body/130/altchunk:0/body/1`); HTML, XHTML and MHT not expanded, there being no HTML reader
  (Q12, Q13); RTF never. Each not expanded raises `altchunk.not_expanded` with its content type;
  `altchunks: false` expands none.
- **One level down** (REQ-512, REQ-513; docx4j's `nested`): the reading's options less embedded
  packages, chunks and metadata; the addresses under the anchor's; the warnings joining the
  reading's; the main story's blocks in the story they are read into (`EMBEDDED` for a package,
  the chunk's own story for a chunk).
- **The glossary** (REQ-501, REQ-516): each `w:docPart`'s body at `glossary:i/j`, last, after the
  endnotes, under `include: glossary`.
- **Equations** (REQ-604): LaTeX by `OmmlToLatex` in the view, `$...$`, `$$...$$` for an
  `m:oMathPara`; in `markup` an equation with tracked changes its Reject All form deleted and its
  Accept All form inserted; one outside the subset its runs' text, as before.
- **Tables** (REQ-605): the simple form where the table is simple (no merge or split, no nested
  table, one non-empty paragraph a cell, every row's cell edges alike); else the listed format, by
  the rules Jason ruled on 18 October 2026; `tables: tsv`; `tables: grid` as pandoc's grid table.
- **Page markers** (REQ-708): each `w:lastRenderedPageBreak` of the main story a `PAGE_MARKER`
  run, its page numbered from 2, written `[page n]` under `page_markers: true`; the option's
  description says the markers are current only when the package was last paginated after its
  last edit. **`layout: indent`** (REQ-709): four spaces a 720-twip stop of a paragraph's
  first-line indent, outside tables.

### 13.3 Fields (REQ-506)

docx4j's `FieldValues` computes in the walk, without a copy: `DOCPROPERTY` (a custom property
first, a core or application one by its built-in name), `DOCVARIABLE`, `AUTHOR`, `TITLE`,
`SUBJECT`, `KEYWORDS`, `COMMENTS`, `LASTSAVEDBY`, `REVNUM`, `CREATEDATE` and `SAVEDATE` (an
argument given is the value), `SEQ` in reading order (`\r`, `\c`, `\h`), `REF` (not with `\n`,
`\r`, `\w`, `\p` or `\f`) and `QUOTE`; never a field with `w:fldLock`; a field whose value cannot
be found keeps its cached result. A `REF` may come before its bookmark, so the package's route
reads the document once with cached results for the bookmarks' text first; a part's or content's
route has none, and records each bookmark as it ends. A computed result replaces the cached one,
which is not read at all: neither its text, nor a note or comment reference, a picture, a page
marker, a field code or a nested field in it. A FORMTEXT is as Word's update leaves it (docx4j's
Office check, 18 October 2026): its default text, or five en spaces where it has none; kept as
typed where the document is protected for forms (`w:documentProtection w:edit="forms"`, enforced
as written), or where the input is not regular text or has a format.

The switches are docx4j's `FormattingSwitchHelper`, a measured subset (risk 5): `\@` on a value
in one of docx4j's inferred formats (its international set: day first, ISO 8601, `dd MMM yyyy`,
with times), Word's pattern converted to Java's as docx4j converts it (its `h` is Java's `K`, the
hour 0 to 11, and the spaces after a closing quote go inside it); `\#` as a Java `DecimalFormat`
rounding half even, the number drawn out of a `DOCPROPERTY`'s or `MERGEFIELD`'s text, `!Syntax
Error` for no number, and docx4j's refusals (`x` before the point, a character between digits,
Word's fraction-only results) leaving the value; `\*` with `Caps`, `FirstCap`, `Upper`, `Lower`,
`ARABIC`, `Roman` and `roman`. docx4j's 138 switch cases with its default properties all hold.

### 13.4 Charts and their numbers (CR-047, risk 5)

`model/chart/` reads a chart as Office shows it: a series with no name is `Series` and its index
plus one; a chart title with no text "Chart Title", or the single shown series' name; an axis
title with no text "Axis Title"; a pie its first series; Excel's filtered series not shown; a
deleted axis left out unless it has a title. `ChartDescription` gives Excel's English kinds,
the axis roles (scatter and bubble by position), the secondary series and the note, and its
`to_json` is docx4j's byte for byte on all 33 of docx4j's goldens. `ChartTable` writes
categories down and series across, multi-level categories a column a level, scatter's x and y.
`ChartExModel` reads a chartex part by element name, since this engine holds it as a tree
(CR-001 section 14.12).

The numbers are formatted as POI's `DataFormatter` in the US locale formats them, for a measured
subset (`formats.py`): `General` (POI's `ExcelGeneralNumberFormat`: ten significant digits
rounded half up), date and time codes by POI's conversion to a `SimpleDateFormat` pattern, and
other codes by POI's cleaning and Java's `DecimalFormat` rounding half up. A code outside it
(elapsed time `[h]`, fractions, engineering notation) is written as the cache holds it, as
docx4j writes it where its formatter throws. `java_text.DecimalFormat` was checked against Java
21's own on 23 patterns, among them `#.##` keeping its leading zero, `.00` not, and `-0` for a
negative that rounds to zero.

### 13.5 Q7's deferred content, and the harness

Embedded SpreadsheetML and PresentationML packages and OLE 1.0 packages are not read (Q7):
`olefile` is not added and this package has no SpreadsheetML or PresentationML model (CR-001 D).
Under `embedded: true` each raises `content.not_read` with its part's name (13.6, item 1).

REQ-512 covers the embedded document this phase reads, so the adapter cannot leave the statement
unadvertised. The harness gains `Engine.deferred(fixture, case, item, scenario)`: a check whose
statements are advertised pends with the reason the engine gives, and docx4j's engine gives none.
The adapter names 20 checks in `DEFERRED`, 60 with their views: case 39's workbook cell, case
41's OLE 1.0 text, case 53's slides (its nine tokens, its hidden slide under `embedded: true;
hidden: true`, and its order), case 57's four sheets (the hidden two under `embedded: true;
hidden: true`) and case 66's three cells. A token whose absence is checked (a hidden slide
without `hidden: true`, the speaker notes) still passes, so none goes from `PASS` to `PENDING`.

### 13.6 Departures, with why

1. **The deferred packages raise `content.not_read`** where docx4j reads them. Q7 defers them, and
   REQ-801's warnings are where a result says it is not the whole document; an empty reading
   would be silent. No check reads the warning.
2. **The harness's `deferred` hook** is this port's second departure from `ExtractionConformance`
   (13.5): docx4j has no deferred content.
3. **A story part's route is docx4j's `Texts.of(Part)`**, which Phase B did not follow: the main
   document part's reading is its story and text boxes, addressed `body`, with no notes and no
   first reading (it was the whole package's); a header's or footer's is addressed
   `header:rId4` in its own story (it was `content`, in the main story). `refOnThePartRoute`
   needs it; nothing in the package called the old one.
4. **A chart or SmartArt data part's root is checked** before it is read. This engine's lenient
   parser makes an empty object of any root, where docx4j's unmarshal refuses one; without the
   check `<notAChart/>` would be an empty chart, not `content.not_read`.
5. **A kept branch's drawing is read by element name.** A drawing in a kept `mc:Choice` is a
   wildcard tree here (CR-002's lossless MCE): its `wp:inline`, `c:chart`, `cx:chart`,
   `dgm:relIds`, `pic:pic`, `v:shape`, `w:object` and `w:control` are found by name, as docx4j
   finds them by class in its bound tree.
6. **POI's exponent `+` only on scientific notation.** POI adds it wherever its result holds an
   `E`, so a suffix such as `EUR` reads `E+UR`; this writes `5.00 EUR`.
7. **Field dates and numbers are in US English** whatever the document's language: docx4j's
   `FieldValues` passes no language, so its dates take the platform locale, which its tests set
   to English.
8. **A control's surface** (for `sdt.binding_differs`) is read in the selected branch of an
   `mc:AlternateContent`, where docx4j's reflective walk reads both.
9. **Bound values special-case the two property stores**, which the content API's `XmlMapping`
   does not (CR-003 section 4): docx4j's text model does, because Word shows them.
10. **`TextOptions(image_describer=)`** is docx4j's `TextOptions.imageDescriber`, a callable; it
    is not one of section 10's names, so `to_dict` leaves it out.
11. **Workbook and presentation tests of `TextsPhase2Test`** (`hiddenSheets`, `workbookCells`,
    the workbook, presentation and OLE 1.0 halves of `embeddedPackages`) are not ported (Q7);
    `test_charts.py` holds `workbookCells`' date, `3/15/23`, for the formats.

### 13.7 The numbers

- **The report**: 1,284 pass, 7 fail, 293 pending of 1,584; goldens 9 of 12 (the probe's three
  plain-text goldens join the six). 138 checks `FAIL -> PASS`, by cause: `w:altChunk` and its
  warnings 45 (cases 36, 43 to 47, 63); charts 18 (34); metadata 18 (38, 52); bound values 12 (23,
  67); SmartArt 9 (35); pictures' alt text 9 (33, 66); the embedded document 9 (40, 63); the
  glossary 9 (54, its order among them); legacy WordArt and the watermark 6 (25, 63); goldens 3.
  60 `FAIL -> PENDING`, Q7's (13.5). None `PASS -> FAIL`. The 7 failures are Markdown's (case 21
  and the three Markdown goldens), Phase D's. The 293 pending are docx4j's 19, the inspector's
  214 and Q7's 60. No check under `fields: update` moved: the suite's rows there hold cached
  results, and `TextsPhase2Test`'s cases hold the computed ones.
- **The suites**: the fast suite 3,088 passed, 8 skipped, 10 xfailed and 628 deselected (87 s);
  the whole suite 3,716 passed, 8 skipped and 10 xfailed (137 s); `tests/text` adds 428 to each.
  The agreement test 636 of 636, CR-002's 46 parity goldens, CR-010's Word-save comparisons,
  `test_readme.py` and `test_performance.py` pass.
- **Unchanged**: `scripts/bench_marshal.py`, `Symbols.docx` parse 189 to 194 ms and marshal 249 to
  260 ms, the 14 main parts 230 to 251 and 298 to 312 ms (Phase B: 192 to 193, 248 to 249, 233 to
  235, 297 to 298); `import docx4j_py` 628 to 640 ms (`-X importtime`; Phase B 644 to 648), none
  of the new modules imported; `scripts/roundtrip.py` 50 of 50 parts canonically identical. The
  engine is not touched.
- **The cost** (CR-003 section 16.4's 200-page document, the median of five, the body
  unmarshalled): `plain_text(extract(pkg))` 36 ms in `accepted`, 34 in `original`, 30 in `markup`
  (Phase B, measured beside it: 32, 36 and 30); `get_text()` 10 ms. The probe 18 ms (15), its
  two document chunks, Flat OPC chunk, chart and SmartArt now read; with every option on, 39 ms.
  The harness's run 2.7 s.

### 13.8 What Phase D needs

- `markdown(result)` and the Markdown readers on the blocks (REQ-703): the 7 failing checks, a
  chart's table and note, a picture as `![alt](…)`, a link from a HYPERLINK field, the original
  view; `outline()` on the model; CR-010's ten pins; CR-006's tools with section 10's options.
- Q7's content, after CR-001 D gives this package SpreadsheetML and PresentationML models, and a
  decision on an OLE 1.0 reader: then `DEFERRED` empties and `TextsPhase2Test`'s workbook,
  presentation and OLE halves are ported.
- The HTML reader as its own phase after D (Q12, Q13).
- Section 12.10's open items stand: the `dry_run` comment defect (12.3) and a fork release
  before this package releases on stage 10. The Word check of the probe and `waterfall.docx`
  saved after their main parts are unmarshalled passed on 2026-10-21.

## 14. Phase D implementation notes (2026-10-21)

Phase D is done: the port of docx4j CR-046 phase 3 at `198d015a3`. Markdown is written from the
blocks by docx4j-markdown's exporter, and every shipped reader the specification governs reads
the model: `to_markdown` (package, body, paragraph, table) and `outline()`. Under Q4 as revised,
the Office JS readers keep Office JS's text and say so. The report moves from 1,284 pass, 7 fail
and 293 pending to **1,291, 0 and 293 of 1,584, goldens 12 of 12**: 7 checks from `FAIL` to
`PASS`, none the other way. Against docx4j's own `MarkdownExporter`, built from a read-only export
of `198d015a3`, **1,385 of 1,385 exports are byte-identical** (277 documents, five readings each),
and so are the 61 `to_markdown` calls the test suite makes, pictures' destinations aside (14.5).
CR-010's ten pins are closed and no `xfail` is left in the suite.

### 14.1 What landed

| piece | where | what |
|---|---|---|
| the writer | `docx4j_py/model/text/commonmark.py` (807 lines, commonmark-java's BSD-2-Clause notice kept; `NOTICE` names it) | commonmark-java 0.30.0's `MarkdownWriter` and `CoreMarkdownNodeRenderer`, the Markdown renderers of its GFM tables, strikethrough and footnotes extensions, and docx4j-markdown's math renderer, for the nodes `WmlToMarkdown` builds; each node may carry a tag, and `render` records the line each tagged node's first character is written on |
| the exporter | `docx4j_py/model/text/markdown.py` (1,204 lines) | docx4j's `WmlToMarkdown`, `MarkdownExporter` and `MarkdownExportOptions`; `markdown(result, package=None, *, image_destination=None)` (Q1), `markdown_blocks(...)` (core-ts's `blocks`), `front_matter(result)` (REQ-704), `MarkdownImage`, `MarkdownBlock` (with `merged`), `MarkdownBlocks`; `TextNumbering.label` (docx4j's `isLabel`) |
| the readers | `model/markdown/export.py` (430 lines, from 1,071), `model/content/text_model.py` (`Reading`, `reading_of`), `model/content/__init__.py` | `body_markdown`, `paragraph_markdown`, `table_markdown`, `markdown_budget_of`, and the new `package_markdown` and `package_markdown_budget`, all on the model; `escape`, `address_comment`, `ADDRESS_COMMENT`, `CODE_STYLE_IDS`, `MONOSPACE_FONTS`, `MarkdownView` (gaining `ORIGINAL`) and `MarkdownViewValue` kept; the old renderer deleted |
| the outline | `model/content/reports.py` | `outline_of` over the reading (`_Model`); `OutlineEntry.textbox` |
| addresses | `model/content/addresses.py` | `address.read_only` for an address of the text model that the content API does not resolve |
| the trial | `model/content/trial.py` | `TrialPackage.get_property_resolver`, `refresh_property_resolver`: a trial's own resolver over its copies |
| the docstrings | `model/content/{body,paragraph,range,table,controls,comments,tracked_change}.py`, `traversal.py`, `docx/{text,table,comments}.py` | each Office JS reader, `text_of` and the facade documented as not the text model (Q4, Q5, Q6) |
| the fixes | `model/text/extractor.py`, `metadata.py` | a level's text and start; the core dates as written and `lastPrinted`; a table given as content one block (14.6) |
| the harness | `tests/text/adapter.py`, `conformance.tsv` | Markdown as docx4j's adapter writes it, `markdown(result, pkg)`; the report rewritten |
| tests | `tests/text/test_text_markdown.py` (53), `content/test_markdown.py` (+3), `agent/test_outline.py` (+3), `agent/test_markdown_workflows.py` (+2), `revisions/test_content_api.py` | docx4j-markdown's tests and this package's (14.8, `tests/README.md`) |
| documents | `CHANGELOG.md`, `README.md`, `docs/pypi-readme.md`, CR-003 sections 3.5 and 3.13, CR-006 section 3.1 | the entries of section 7 as Q4 revised adjusts them (14.3) |

Commits: `e0ec105` (Markdown from the blocks), `f2b8e58` (the two defects the byte check found),
`77b2f50` (the readers, CR-010's pins), `382f96f` (the writer's escaping), and this section with
the documents.

### 14.2 Markdown as ported (REQ-703, REQ-704)

docx4j's rules, unchanged, as core-ts's section 16 ports them:

- **Blocks.** A heading from its outline level (0 to 5), before the numbering. A list from its
  numbering: two `w:num` on one `w:abstractNum` are one list; a margin paragraph ends it and it
  goes on at Word's count; a paragraph indented to an item's text (or in `ListParagraph`) is a
  follow-on of the deepest item it reaches; an empty paragraph leaves it alone; an adjacent list
  takes the other delimiter or bullet; a nested list resumed after an interruption opens under an
  empty item showing its parent's count. A label level (`none`, or a `w:lvlText` with no `%n`) is
  a bold label. A code-styled paragraph, or one in a named paragraph style whose own font is
  mono, is a fence, consecutive ones one fence. A Quote or IntenseQuote paragraph is a block
  quote nested by its indent beyond the style's own (the package asked through its resolver). An
  empty paragraph with a bottom border of its own is `---`; a paragraph that is one `m:oMathPara`
  is display math.
- **Inlines.** Bold, italic, strike and code where a run adds them to its paragraph's style, a
  header row's bold not marked; links from a `w:hyperlink` or a HYPERLINK field; a picture with
  the describer's text; notes `[^1]` and `[^e1]`, `[^f-5]` for a custom mark, the definitions
  where the model reads them; `[page n]` under `page_markers`; in `markup` CriticMarkup flat,
  comments `{>>...<<}` at their reference with their replies, `markup_detail` splitting by
  author or date.
- **Tables.** GFM: the header the `w:tblHeader` rows (several joined ` / ` by grid column) or the
  first row, the alignment the header cell's first paragraph's, a span padded with empty cells, a
  merge's continuation empty, a cell's paragraphs `<br>` apart. A chart is its note, its title and
  its table, the value columns right-aligned.
- **The tags** (core-ts's departure 5): each node made for a block carries it, and the writer notes
  the line its first character lands on, so `markdown_blocks` cuts the Markdown into whole lines
  per block, joining to `markdown`'s. A block that writes nothing has `""`; `merged` says it is a
  code paragraph whose line is in the fence of a block before it, which the address comments need
  (14.3).
- **Pictures.** `image_destination`, a callable, stands for docx4j's `imageDirPath` and
  `imageTargetUri` (core-ts's departure 2): it is given the run, the part, its bytes, its content
  type, its number among the export's pictures and docx4j's file name (`image1.png`), and returns
  the link. The default is docx4j's data URI. An SVG picture's bytes are the part's as stored
  (core-ts's departure 3); in this engine an SVG part is an `ImagePart`, so no export of the 277
  differs by it.

### 14.3 The readers (section 3.5 under Q4 as revised)

- **`to_markdown`** is `markdown()` over the text model's reading, every keyword kept. The reading
  (`reading_of`) is the package's for the main body (docx4j `Texts.of(pkg)`: its story, text boxes
  and notes), the part's for a header or a footer, and the content route for any other body (a
  cell, a control, a comment, a notes part). `view` takes `accepted`, `original`, `markup` and
  `current`; anything else is `markdown.view_invalid`, as before. `pkg.to_markdown()` and
  `pkg.markdown_budget()` read the package itself, with nothing unmarshalled (14.4). A heading is
  the block's effective outline level through the styles, where 0.1.0 also took a `HeadingN`
  style id that names no style as one, so such a paragraph is now body text, as in docx4j; the
  outline's levels follow.
- **Addresses.** `addresses=True` writes each block's address in today's `ADDRESS_COMMENT` form, on
  its own line before the block's first line and indented as it is. A block of the body's own story
  has its content API address: `w14:` and its paraId where it has one, else the model's ordinal
  under the body's prefix, which is the content API's (measured: 2,088 main-story blocks over
  `samples/` and `tests/fixtures/`, each resolving with `element_at` to the element it was read
  from). Any other block keeps the model's address, read only (Q9): `body/3/textbox:0/0`,
  `body/103/chart:0`, `footnote:2/0`. `element_at` refuses those with `address.read_only` rather
  than `address.not_found`. A code paragraph merged into a fence has no comment, as in 0.1.0; an
  empty paragraph keeps its own.
- **`max_chars`** cuts after the last whole block that fits, trailing white space dropped, and in
  the middle of the first block only when that one is over the budget.
- **`Paragraph.to_markdown`** is the paragraph's lines in its story's Markdown, trailing white space
  dropped, so a list item has its document marker (`3. LST56\_PLUMS`, `4. 2. LST56\_QUINCE\_TWO`).
  A paragraph in a table cell, or a code paragraph merged into a fence, is written on its own. One
  joined into the paragraph before it is that block's. **`Table.to_markdown`** is the table's
  block, likewise: the brief named `table_markdown` among what had to keep working, and putting it
  on the model leaves no second Markdown in the package (REQ-202).
- **`outline()`** keeps one entry per block of the tree (controls, tables and their cells, `depth`,
  `limit`, `headings_only`, the addresses). An entry's text, its heading level and a table's
  first-row preview are the model's, in the accepted view with the defaults: a joined pair's text
  on the first paragraph's entry and the second's empty; a hidden or a table-of-contents paragraph
  empty. The counts are the model's blocks', the notes' left out (`paragraphs`, `tables`, `words`,
  `chars`); the probe's are 171 paragraphs and 5 tables. Each text box's paragraphs follow their
  anchor at their `textbox:` addresses with `OutlineEntry.textbox` (Q9). For a box anchored in a
  cell they come after the table, as the model reads them (REQ-518). Charts, SmartArt and chunks
  have no entries. `stats.tracked_changes` and `describe().tracking_on` are untouched; CR-010's
  tests of them pass unchanged.
- **Kept off the model** (Q4 revised, Q5, Q6), each now saying so in its docstring: `text` on
  `Paragraph`, `Body` (and its `Block`), `Range`, `Table`, `TableCell` and `ContentControl`;
  `Comment.content`; `get_text` in `accepted` and `original`; `search` and `find`;
  `TrackedChange.text`; `text_of`; and the python-docx facade. **`text_of` is not deprecated**:
  Q4 revised supersedes section 7's "Deprecated" bullet, and the CHANGELOG deprecates nothing.

### 14.4 A read leaves every part as it was (decided with the coordinator, 2026-10-21)

**A behaviour change of the save after a read.** 0.1.0's `pkg.to_markdown()` went through
`pkg.body`, which unmarshals `word/document.xml`, and read the note definitions through
`part.contents`, which unmarshals `word/footnotes.xml`. So a save afterwards re-marshalled both,
canonically identical to the source. The text model reads every part with `read_contents`, and
`pkg.to_markdown()` reads the package without a body. So a save after it now writes both parts from
the source bytes. Measured over the 70 documents of `samples/`, `tests/fixtures/` and the probes:
`word/document.xml` in all 70 (`document22.xml` in one), `word/footnotes.xml` in the 7 with a
footnote reference, no other part's bytes changed. This is the fidelity rule: a part the caller
never asked to unmarshal is written back byte for byte. The coordinator accepted it: it needs no
Word check, and nothing should unmarshal a part on purpose to keep the old bytes. `Body.to_markdown`
on an existing `pkg.body` is as before, the body's part being unmarshalled already. **`outline()`
still unmarshals the notes parts**, through `package_bodies` (unchanged in this phase; 14.10).
`test_reading_the_markdown_unmarshals_nothing_and_a_save_is_the_source` (footnotes.docx, the probe)
holds it: no part unmarshalled, and every part of the save byte-identical to an untouched save and,
outside the content types and the relationships parts that every save writes, to the source. It
fails on the old reader, which leaves `/word/document.xml` and `/word/footnotes.xml` unmarshalled.

**Read, then edit by address.** `test_every_address_the_markdown_writes_resolves_once_the_body_is_read`
(`2010-sample1.docx` and the probe, both with ordinal addresses) reads `pkg.to_markdown(addresses=True)`
with nothing unmarshalled, then touches `pkg.body`. Every address written resolves to the element a
fresh reading reads there, with the text the Markdown showed: all 4 of `2010-sample1.docx`'s and 134
of the probe's 159, with `element_at` and, for paragraphs, `paragraph_at`. The probe's other 25
(text boxes, the chart, SmartArt, the chunks' paragraphs, the notes) are refused as
`address.read_only`. Then one edit by address. A first `pkg.body` after a read parses the main part a second
time (14.9).

### 14.5 Checked against docx4j itself (the brief's byte check)

docx4j at `198d015a3` was exported with `git archive` into the scratch directory and built there
offline, `mvn -o -q package -DskipTests -pl docx4j-markdown,docx4j-JAXB-ReferenceImpl -am`, with
`-Dmaven.repo.local` a scratch directory and `-Dmaven.repo.local.tail=~/.m2/repository`, which
Maven reads and never writes. Nothing was installed and the scratch repository stayed empty. A
runner wrote `new MarkdownExporter().export(Texts.of(pkg, options), pkg)`, docx4j's adapter's call,
for each document in five readings, as core-ts's sweep did: `accepted`, `markup`, `original`,
`accepted` with `metadata: true`, and `markup` with `markup_detail: author_date; comments: true`.
This package wrote `markdown(extract(pkg, options), pkg)` for the same.

- **The brief's set**: the suite's five fixtures, CR-010's eleven and the 147 docx4j fixtures of
  the agreement test (163 documents, 815 exports). The first comparison gave 665 identical. All 150
  differences were lists: the extractor's `TextNumbering` read `level.lvl_text`, which
  `ListLevel` does not have, and the `AttributeError` its `except` swallowed left every
  `level_text` None (so every level a label) and every `start` 1. Fixed (`e0ec105`), and docx4j's
  fallback to the abstract definition's level ported: **815 of 815 identical**, data URIs included.
- **Extended** to this package's `samples/` and `tests/fixtures/` (114 more documents, 1,385
  exports in all): 4 differed, all in `metadata: true`'s front matter, two defects of Phase C's
  `metadata.py`. `lastPrinted` was dropped, an `XmlDateTime` being neither a literal nor a W3CDTF.
  `created` and `modified` came out with xsdata's nine fractional digits (`.495027100` for the
  part's `.4950271`). Both are now read as the part writes them, as docx4j reads them (`f2b8e58`):
  **1,385 of 1,385 identical**.
- **The readers.** A pytest plugin saved, with its package, every main-body or package
  `to_markdown` the suite calls with no addresses and no budget (61 calls in `tests/content`,
  `tests/agent`, `tests/revisions` and the readme fences). docx4j's export of each saved package
  was compared with what the reader returned: **61 of 61 identical** with picture destinations
  masked. The first run's one difference was an empty document's `""` against docx4j's `"\n"`,
  fixed.

### 14.6 Found on the way, and fixed

- `TextNumbering`'s `level_text` and `start` (above), a Phase B defect; the block JSON's
  `levelText` and `start` are now docx4j's. Plain text never read them, so no check moved.
- The core properties' dates and `lastPrinted` (above), Phase C's.
- **A table given as content is one `TABLE` block** (`extract(tbl)`, a `Table` view), as docx4j's
  `Texts.of(pkg, content)` reads it: its `blockItems` gives a table none. It was its cells'
  paragraphs read flat. `Table.to_markdown` needs it.
- **A trial's own property resolver.** A dry run's `to_markdown` read the trial's headings through
  the real package's styles, `__getattr__` handing `get_property_resolver` through. So a heading
  `insert_markdown` had just defined in the trial's copy of `styles.xml` was body text
  (`test_a_dry_run_of_insert_markdown_previews_without_committing`). `TrialPackage` now has its
  own, refreshed when the trial writes a style.
- **The writer's escaping**, the profile's largest local cost (14.9): a `str.translate` table per
  set of characters in place of a character at a time (`382f96f`). Output unchanged, 1,385 of
  1,385.

### 14.7 Departures, with why

1. **`image_destination`, a callable**, for docx4j's `imageDirPath` and `imageTargetUri` (core-ts's
   departure 2): the exporter writes no file; a host that wants files writes them in the callable.
2. **The readers' pictures point at their parts** (`media/image1.png`, relative to `/word/`), as
   0.1.0's did, where `markdown()` writes docx4j's data URI. A reader's Markdown goes into a
   context window, and REQ-703 leaves the destination to the implementation. This is the one way
   the readers' Markdown is not docx4j's.
3. **An SVG picture's bytes are the part's as stored** (core-ts's departure 3), docx4j's being its
   `XmlPart` re-serialised; no export of the 277 differs by it here.
4. **`MarkdownBlock.merged`**, beside core-ts's two fields, for the address comments.
5. **A block's content API address in the comments** (paraId first), not the model's ordinal
   alone, which CR-003 section 12.7 decided and its tests hold: the model's address is the
   comment's only for the blocks the content API does not reach.
6. **`Paragraph.to_markdown` and `Table.to_markdown` drop trailing white space**: a block's lines
   end with the blank line before the next block, which an excerpt does not want. The document's
   Markdown (`pkg`, `Body`) ends as docx4j's does, with a line end.
7. **docx4j's deprecated `TrackedChangesPolicy` is not ported** (core-ts's departure 4): the view
   is the option. **docx4j's Markdown importer** (`MarkdownImporter`, `LatexToOmml`) is not ported
   either: the importer is CR-003's, on `markdown-it-py`. So the four of `MarkdownMathExportTest`'s
   six that read LaTeX through docx4j's importer or `LatexToOmml` are not ported, and the picture
   tests build their picture instead of importing it.
8. **A kept block-level `mc:AlternateContent` reads neither branch.** It is not in this engine's
   typed tree (CR-012, docx4j CR-051's gap, proposed): the part's skipped-content report names it.
   `MarkdownAlternateContentTest`'s kept load is pinned so; its `mce_preprocess` load reads the
   Choice, as docx4j's does. **Closed by CR-012 Phase A (2026-10-21, its section 9):** the schema
   copy admits the element at block level, the extractor's existing `mc` case reads Word's branch,
   and the pin now reads the Choice with nothing skipped, in both load modes.

### 14.8 Pre-existing tests whose expectation changed, and why

Every new output below is **docx4j's**: each `to_markdown` these tests call is among 14.5's 61,
identical to docx4j's export of the same package with picture destinations masked, and a
paragraph's or table's is a slice of that Markdown. The reason given is the rule that moves it.

`tests/content/test_markdown.py` (CR-003's Markdown tests):

| test | changed to | why |
|---|---|---|
| `test_a_heading_is_hashes_from_the_shared_heading_level` | `...Plain\n` | the document's Markdown ends with a line end (commonmark-java's `Document`) |
| `test_an_outline_level_is_a_heading_without_a_heading_style` | `...level\n` | the line end |
| `test_bold_italic_strike_and_code` | `` `x = 1`\n ``, with `CodeChar` defined first | a run is code by its font through its styles (REQ-703's monospace fonts); an undefined `CodeChar` gives no font, and docx4j writes `x = 1` plain, so the test defines the style to keep its subject |
| `test_a_monospace_font_is_code_when_no_style_says_so` | `` `mono`\n `` | the line end |
| `test_two_bold_runs_in_a_row_are_one_pair_of_markers` | `**one two**\n` | the line end |
| `test_a_hyperlink_reads_its_relationship_and_an_anchor_reads_its_anchor` | `...slashdot.org/)\n` | the line end |
| `test_an_internal_link_is_a_fragment` | the TOC's link absent from `to_markdown()`, present in `markdown(extract(pkg, toc=True), pkg)` | the table of contents is left out by default (REQ-520, `toc: false`) |
| `test_a_bullet_list_nests_by_ilvl` | `...nested\n` | the line end |
| `test_a_table_is_a_gfm_pipe_table` | `\|Region\|Total\|\n\|---\|---\|\n\|North\|120\|\n` | commonmark-java's table renderer: no padding inside the pipes, `---` |
| `test_a_vertically_merged_cells_first_row_keeps_its_text` | `\|A\|\n\|---\|\n\|Merged\|\n\|\|\n` | the same |
| `test_the_table_renderer_works_over_the_element_alone_for_phase_c` | `\|A\|\n\|---\|` (both forms) | the same, trailing white space dropped (departure 6) |
| `test_a_header_rows_bold_is_convention_and_is_not_marked` | `\|A\|B\|\n\|---\|---\|\n\|1\|2\|\n` | the same |
| `test_an_image_is_its_part_name_relative_to_word` | `...image2.jpeg)\n` | the line end; the destinations are kept (departure 2) |
| `test_a_footnote_reference_is_a_gfm_footnote_with_its_definition_at_the_end` | ends `[^1]: Hello\n` | the line end |
| `test_a_comment_is_criticmarkup_in_the_markup_view_only` | `One paragraph\n`, `One paragraph{>>One comment<<}\n` | the line end |
| `test_the_view_values_are_a_literal_with_a_strenum_beside_them` | `original` and `current` read; `final` refused | the views are REQ-401's |
| `test_a_code_style_paragraph_becomes_a_fence_and_consecutive_ones_merge` | `...```\n` | the line end |
| `test_a_quote_style_paragraph_is_a_block_quote` | `> Quoted.\n` | the line end |
| `test_markup_that_would_be_read_back_as_markup_is_escaped` | `1\. not a list`, and the line end | commonmark-java escapes an ordered-list marker's delimiter, not its first digit |
| `test_an_empty_paragraph_contributes_nothing` | `one\n\ntwo\n` | the line end |
| `test_a_link_becomes_a_hyperlink_with_an_external_relationship` | `...).\n` | the line end |
| `test_an_ordered_list_keeps_its_start_value` | `3. three\n4. four\n` | the line end |
| `test_a_code_block_becomes_source_code_paragraphs` | `...```\n` | the line end |
| `test_an_image_is_never_fetched_and_is_reported` | `...pangolin.png)\n` | the line end |
| `test_an_html_block_is_skipped_and_reported` | `Text.\n` | the line end |
| `test_the_whole_fragment_survives_a_save_and_a_reload` | `\|Region\|Total\|` | the table form |
| `test_reading_a_document_as_markdown_leaves_every_other_part_byte_for_byte` | `pkg.to_markdown()` changes no part; `pkg.body.to_markdown()` still only `word/document.xml` | 14.4 |

Elsewhere:

- `tests/content/test_comments.py::test_a_comment_is_a_body_of_its_own`: `Is this right?\n`, the
  line end.
- `tests/content/test_lists.py::test_a_two_level_list_with_a_restart_round_trips_through_markdown`:
  the line end.
- `tests/content/test_lists.py::test_a_loose_items_follow_on_paragraph_is_rendered_under_it`:
  `- one\n  \n  continued here\n\n- two\n`. commonmark-java's loose list writes the item's blank
  line with its prefix and separates every item, the list being loose throughout.
- `tests/content/test_table.py::test_to_markdown_is_the_renderer_over_the_element`:
  `|Region|Quarter|Total|`, `|---|---|---|`, the table form.
- `tests/agent/test_markdown_workflows.py::test_markdown_in_and_out_gives_back_the_same_structure`:
  `|Area|State|`, the table form.
- `tests/agent/test_scenarios.py`, two tests: the outline's text-box entries are left out of the
  paraId session, and in the corpus session each is checked to be refused as `address.read_only`
  rather than to resolve (Q9).
- `tests/revisions/test_content_api.py`, CR-010's ten pins (14.4's decision, below).
- The README's shown outputs, which nothing asserts: the table form, the line end, the budget's
  `chars` (1,585 for 1,610).

**CR-010's pins (the coordinator's decision, recorded).** The five `to_markdown` pins pass as they
were written, so their `xfail` marks are gone. The five `get_text` pins cannot pass: under Q4,
`get_text(view="accepted")` stays Office JS's filtered text, which does not join a deleted
paragraph mark. They are re-pointed to the model, on the same eleven fixtures:
`plain_text(extract(fixture, view="accepted")) == plain_text(extract(accept_all copy))`, which
holds on all eleven. `test_get_text_keeps_office_jss_filtered_text` holds `get_text`'s documented
form on `22-deleted` (`'One.\n\nThree.\nFour.\nFive \neight.\nNine.\n'` against Accept All's
`'One.\nThree.\nFour.Five eight.\nNine.\n'`). `DIFFERS` keeps the measured reasons as the record
of that difference. No CR-010 `xfail` is left.

### 14.9 The numbers

- **The report**: 1,291 pass, 0 fail, 293 pending of 1,584; goldens 12 of 12. 7 checks `FAIL ->
  PASS`, by cause: case 21's four link rows (`FLD21_URL_MARK` and `FLD21_FIELD_URL_MARK` under
  `markdown`, the original view and a HYPERLINK field's link) and the three Markdown goldens. None
  `PASS -> FAIL` or `PASS -> PENDING`. The 293 are docx4j's 19, the inspector's 214 and Q7's 60.
- **The gate** (section 3.6): no `FAIL` row; the agreement test at zero (636 of 636, the slow
  run's 588 among them); `LevelOwnerTest`'s 61 paragraphs as Word numbers them; CR-002's 46 parity
  goldens at zero (277 tests, `PINS` empty) and CR-010's Word-save comparisons (68) unchanged;
  `scripts/roundtrip.py` 50 of 50 parts canonically identical; `test_readme.py` and
  `test_performance.py` pass. The fast suite: 3,162 passed, 8 skipped and 628 deselected (97 s;
  Phase C: 3,088 passed, 8 skipped, 10 xfailed); the whole suite 3,790 passed and 8 skipped (144
  s; Phase C: 3,716, 8 and 10 `xfail`). No `xfail` is left.
- **Unchanged**: `scripts/bench_marshal.py`, measured beside Phase C's `48d00de` on this machine:
  `Symbols.docx` parse 202 to 218 ms and marshal 269 to 276 ms (`48d00de`: 204 to 235, 254 to
  264), the 14 main parts 250 to 263 and 320 to 329 ms (247 to 280, 305 to 317). `import
  docx4j_py` takes 639 to 685 ms (`48d00de`: 636 to 700) and imports none of the new modules.
- **The cost** (CR-003 section 16.4's 200-page document, the median of five or seven):
  - With the body parsed, as the fixture builds it: `to_markdown()` 48 ms (0.1.0: 8), `markup` 45
    (8), `addresses=True` 59 (49), `outline()` 50 ms (16); `markdown(extract(pkg), pkg)` 45 and
    `plain_text(extract(pkg))` 34.
  - **The read-then-body workflow**, loaded fresh from its bytes: `pkg.to_markdown()` then the first
    `pkg.body`. Before (0.1.0): **50 ms**. At `77b2f50`: **169 ms**, and body first then
    `to_markdown`, one parse, **134 ms**. The jump is mostly the model and the renderer; the
    main part's second parse is 35 ms, 26%. After the escaping fix (`382f96f`), 132 to 135 ms, and
    90 to 96 body first. The second parse is the same 39 to 42 ms, now about 40 to 45% of the
    rest. The coordinator chose to keep it: every part staying byte for byte after a read is worth
    it (14.4).
  - **The profile** of `pkg.to_markdown()` (cProfile, cumulative). Before the fix, the top three
    were the main part's `read_contents` parse (149 ms profiled), the commonmark render (145 ms,
    130 of it `_Writer._write`'s character loop) and `extract`'s walk (114 ms;
    `Extractor._paragraph` 106, `office_span` 31). After it, the parse (145 to 158 ms profiled,
    about 42 ms real), the walk (117 ms) and `WmlToMarkdown.convert` (27 ms), the render down to
    31 ms. Nothing else local is left: the walk is Phase B's, and the parse is the engine's.
    Unprofiled, `pkg.to_markdown()` went from 127 ms to 89 to 93 ms.

### 14.10 Follow-ups, and what is not done

- **`XmlPart.contents` adopting the private tree** (the engine's, CR-002, not this phase's).
  `read_contents` parses a part into a private tree and `contents` parses it again. Where nothing
  has mutated the private tree, `contents` could adopt it, for the main part, styles and
  numbering, which the resolvers read privately. That removes 14.9's second parse (about 40 ms on
  the 200-page document). It needs tests that a private tree is never mutated by a reader (the
  text model, the resolver, the emulator, the font selector), that an adopted tree is linked and
  reports its skipped content as an unmarshalled one does, that a save after adoption marshals
  what a fresh parse would, and the thread rule (one package per thread) unchanged.
- **`outline()` could move to `read_contents` without changing what it returns.**
  `package_outline` builds a body for every part through `package_bodies` (`body_of`, which
  unmarshals) only to keep the headers' and footers'. So the notes and comments parts are
  unmarshalled for nothing. Its structural walk could also run over the reading's trees, since
  the addresses and entries would be the same. Not done in this phase, at the coordinator's word.
- **Charts, SmartArt, embedded and chunk blocks have no outline entries**, only text boxes (Q9 asked
  for those).
- **A block-level `mc:AlternateContent` kept on load** is not read (CR-012), as 14.7 item 8 says.
  Closed by CR-012 Phase A (its section 9).
- Section 13.8's other items stand: Q7's content, the HTML reader (Q12, Q13), the `dry_run` comment
  defect of 12.3, and a fork release before this package releases on stage 10.
- Lint: the new modules pass `ruff check`; the files they touch keep their earlier findings
  (`extractor.py`'s `S110`, the import order of `paragraph.py` and `range.py`).

## 15. Phase F: the Office JS readers read Word's branch (Q4 revised 2026-10-22; started 2026-10-22)

**Why.** Q4 kept `text`, `get_text` and `search` on Office JS's text. This package's readers took
that to be the paragraph's own runs, a kept `mc:AlternateContent`'s branches passed by, and so did
core-ts until its CR-012 phase 4. Word 365 measures otherwise.
- **core-ts check 49** (2026-10-20, the user, Word 365, Script Lab), over the editor's
  `packages/editor-model/test/fixtures/mc/mc-branches.docx` (docx4j-ts-editor `4e02089`; Word's
  save with its report is `mc-branches-check49-word365.docx`, editor `3189640`, ED-009 section 9.2
  item 4). In all ten cases, `paragraph.text` gives the selected branch's words, never a Fallback's
  and never nothing. That covers run level and paragraph level: case 8's paragraph-level element,
  Choice "ying" and Fallback "yang", reads `bold ying plain ying`. A Choice requiring `x14ac`
  counts as understood. `body.search` finds a word that is only in a Choice, and finds none that is
  only in a Fallback.
- **core-ts check 42** (2026-10-18): a run-level text box, a wps Choice with its VML Fallback,
  gives `paragraph.text` nothing. `body.text`, `body.paragraphs` and `inlinePictures` leave the box
  out.
- **This package, on check 49's probe, at `05fdc74`:**
  - `Paragraph.text` is `''` for cases 1 to 5 and 8, and `"I prefer "` for 6, 7, 9 and 10. Word's
    `paragraph.text` is the branch's words.
  - `body.search("ying")` and `search("flat")` find nothing.
  - `text_of`, behind `Block.text`, cell text and `Table.values`, already reads Word's branch
    (`bold ying plain ying`). So a cell's text and its paragraph's text disagree today.

**The two other engines.**
- **core-ts** reads the selected branch at run and paragraph level since its CR-012 phase 4
  (`9307bdb`, 2026-10-21, held to check 49): in `Paragraph.text`, its offsets, `Range` and
  `search`.
- **docx4j** has no Office JS readers. Its readers of text select by `McSelection`'s TEXT purpose,
  the ISO rule over the published understood list (at `2853b7280`: `TraversalUtil`'s READ mode
  since 17.3.4, docx4j CR-046). This package's `child.mce_branch` over
  `understood_namespaces()` is the port of that rule, already used by `text_of` and the text model
  (REQ-302 to 306).
- So after this phase, every reader of text here takes the same branch as docx4j's and core-ts's.

**The change.**
1. **Readers:** `Paragraph.text` (so `Body.text`, `Range.text` as a slice, and a block control's
   text), `get_text` in `accepted` and `original`, `search` and `find`.
   - They read a run-level or paragraph-level `mc:AlternateContent` in a paragraph as the branch
     `mce_branch` selects. Its runs are read as the paragraph's own runs are read: the same
     characters, the same deletions out in `accepted` and insertions out in `original`.
   - A text box in a branch contributes nothing, as one outside a branch does (check 42).
   - Block level is unchanged: Phase A's `Body.paragraphs` already reaches the selected branch's
     paragraphs.
2. **The characters stay this package's.** A break is `\n` and a tab `\t`, inside a branch as
   outside it. Word gives U+000B for a `w:br` (check 49, case 2). core-ts's phase 4 moved to
   Word's characters throughout, along with hidden text and note references. That is a separate
   question, not this phase.
3. **Edits.** An offset or a `Range` over a branch's text reads, but an edit through it is refused
   with `address.read_only` (CR-012 Phase A's rule for an edit inside a kept branch). That covers
   `insert_text`, `font` writes, `splice`, `delete`, `insert_content_control` and tracked replace.
   - An edit elsewhere in the same paragraph works as before. Its offsets count the branch's text.
   - `Paragraph.text = ...` replaces the paragraph's content, as now, and takes the element with
     it. What Word does there is unmeasured.
   - core-ts edits a selected branch in place (its ED-009 section 12.2). Porting that is not this
     phase.
4. **The map** (WD 0.27's REQ-202 condition, 12.6). Each text model `TextRun` read from a branch's
   run gets its `office_offset` in the new text, and the offsets after it move.
5. **Unchanged:**
   - the text model and Markdown, which read the branch already;
   - `text_of`, which reads it already;
   - `TrackedChange.text` (Q6);
   - the python-docx facade (Q5): python-docx's `Paragraph.text` is its `w:r` and `w:hyperlink`
     children only;
   - `outline()`, which is the text model's.

**The gates.**
- Tests over check 49's probe, taken from the editor's history at `4e02089`. Each case's
  `Paragraph.text` is Word 365's, except case 2's U+000B, which is pinned as item 2's difference.
  `search("ying")` and `search("flat")` hit; `search("yang")` does not. An edit through a hit is
  refused, and an edit beside the element works.
- `samples/` and the CR-012 corpus: Phase B's readers' audit (`scripts/mce_branches.py readers
  --holding either`), before and after. Every answer that moves must be Word's branch's text.
  Nothing outside a branch may move.
- The suite, CR-011's ratchet and goldens, CR-010's Word-save comparisons, the parity goldens, and
  `office_offset`'s agreement with `get_text()` character for character (12.6).
- Acceptance is unchanged: none of its artefacts has a branch. A Word check only if a write path
  changes.

**Not in this phase:**
- Office JS's characters (item 2);
- edits in place inside a branch (item 3);
- `body.text` over a paragraph-level element, `range.text` over such a paragraph and
  `tableCell.value` for a cell holding one, which check 49 did not measure. They follow from
  `Paragraph.text` here by construction, and are recorded as unmeasured.
