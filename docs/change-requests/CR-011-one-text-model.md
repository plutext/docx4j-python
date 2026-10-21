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
1,584.
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
