# CR-011: One text model --- the port of docx4j CR-046

**Status:** Proposed 2026-10-18, on docx4j CR-046 phase 4's hand-off (the goldens at docx4j
`d9acdf5e2`, the specification at WD 0.14), as the portfolio paper's section 10 row for this
package proposes: `../docx4j-portfolio/docs/text_extraction_behaviour_model.md`, whose section 9
recommendations the user accepted on 2026-10-17 (portfolio `2d8fd36`). It takes in CR-010 Phase
C's view agreement (CR-010 sections 6 and 12.3). **Implementation waits on the user's go**; the
questions of section 9 carry recommendations.
**Depends on:** docx4j CR-046 phases 0 to 3, done 2026-10-18 (`c0c72ecef`, `65a527d47`,
`f71cdd299`, `8da15bb8c`, `608e40023`, `187e5a367`, `f59828e62`, `d9acdf5e2` on `VERSION_17_3_4`;
from `608e40023` on unpushed, so the local clone is the oracle); the specification
`../docx4j-portfolio/docs/text_extraction_specification_wd.md` at WD 0.14 (portfolio `cf56600`);
this repository's CR-010 (implemented; its ten `xfail(strict)` pins are this CR's to close); for
Q3, a fork release. **Updated 2026-10-19** on docx4j-25's notes of `a41de133b` and `3dcbb3764`
(both unpushed). The suite is now cited at `3dcbb3764`. That commit adds a third fixture,
`extraction-probe-structure.docx`, with cases 55 to 57 (a table of contents, an interrupted list,
an embedded workbook's hidden sheets) and four goldens; the eight earlier goldens are unchanged
since `d9acdf5e2`. The specification is at WD 0.15 (portfolio `ac06ed7`, `7affcd0`). A suite row
may now combine options, and every row needs its scenario's statements (REQ-1201; sections 3.6
and 6). The choices docx4j CR-046 phase 2 made where WD 0.8 was silent
are decided and normative (Jason, 19 October 2026): REQ-519's addresses, apart from the text-box
segment, which stays a proposal (section 3.2); REQ-515's plain-text property lines; REQ-604's markup
form of an equation; and REQ-512's hidden sheets and slides under `hidden: true`.
**Counterpart:** docx4j's `docs/developer/change-requests/CR-046-text-extraction-behaviour-model.md`
(section 2 the design, 4 the phases, 5 every statement and where docx4j meets it, 6 its decisions),
and its code, `org.docx4j.model.text` (27 classes, 6,203 lines) and the harness
`org.docx4j.textextraction.ExtractionConformance` with its engine adapter `TextsAdapter`.
docx4j-core-ts's CR-012 (`docs/change-requests/CR-012-text-model.md`) is the sibling port,
proposed 2026-10-19 and **decided by Jason on 2026-10-21**: D1 to D9, all as recommended (core-ts
`9b302ad`, pushed). Five of its decisions touch this CR: D1 (options keyed by the specification's
section 10 names verbatim) Q1, D2 (the Office JS readers keep Office JS's text) Q4 and section
3.5, D4 (the text-box address) section 3.2, D5 (text-box views read-only) Q9, and D9 (REQ-510 in
the emulator first) risk 3. D2 decides core-ts's readers only (its section 6.1); Q4 here is still
the user's.

**Since then (2026-10-21).** The suite and the specification have moved on. docx4j `198d015a3`
has five fixtures; the fifth is `extraction-probe-html.docx`, read only under `html reader present`.
It also has CR-050's inspection rows, pending without REQ-901 to 907. The specification is at WD
0.28 (portfolio `f2046e5`). docx4j CR-053's hand-off (`35e304770`, `1d108c0af`, `4a8833eea`,
`198d015a3`) changes REQ-513. An HTML, XHTML or MHT chunk (an MHT's first `text/html` part) is
read only by an *HTML reader*: a parser of the [HTML] algorithm and a reader of its tree as
blocks. Without one, the chunk is not expanded and raises `altchunk.not_expanded` with its
content type. The suite's scenario is now `html reader absent` (it was `xhtml importer absent`).
A chunk is decoded by its BOM; else a meta charset in its first 1,024 bytes; else UTF-8 where
valid; else windows-1252. There is no XML-declaration step and no UTF-16 sniff. Section 6's
`lxml.html` and section 3.6's built-in MHT expansion predate this: libxml2's HTML parser does not
follow the [HTML] algorithm (Q12). This CR still cites `3dcbb3764` and WD 0.15. The re-cite, with
a baseline rerun, waits on the user.

## 1. Why

The paper (its section 1) found that the three engines answer "what is the text of this
document?" differently, and that this package answers it five ways. Measured here on 2026-10-18
against docx4j's conformance suite (then two fixtures, 54 cases, 941 checks), and again on
2026-10-19 against the suite at `3dcbb3764` (three fixtures, 57 cases, 1,023 checks):

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
- **Goldens: 3 of 8** pass (the external fixture's three plain-text views); the probe's four
  differ by 66 to 106 lines, the external Markdown by a dropped linked picture. At `3dcbb3764` the
  count is 3 of 12. The structure fixture's three plain-text goldens fail on the table of
  contents' two entries (REQ-520). Its Markdown golden fails on those entries and on the
  interrupted list: the indented paragraph's continuation, `4. 2.` written `  2.`, and the second
  list's `1)` written `1.` (REQ-703, REQ-510).
- **The failures, by class** (169 in the default scenarios at `d9acdf5e2`): text boxes and `mc:Choice`
  content not read (49), charts and SmartArt (27), `w:altChunk` not expanded and no warning (24), notes
  not appended (12), no plain-text markup view (10), hidden text included (9), math and ruby (9),
  no metadata channel (9), run-level `mc:AlternateContent` not read (6), bound values not read
  (6), joins, table layout and note markers in the goldens (5), comments in markup (3).
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
  (`reports.py:133`: address, para id, style id, level, text). HTML for `w:altChunk`: `lxml.html`.
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
   (REQ-519, core-ts CR-002 section 3.3's grammar; decided in WD 0.15 apart from the text-box
   segment, still a proposal after core-ts CR-010's D8: REQ-519's NOTE 2 gives the final spelling to
   core-ts CR-012, whose D4, decided, keeps docx4j's `body/3/textbox:0/1` as docx4j `608e40023`
   implements it and asks the specification to mark the segment decided); notes as
   `[n]` with their text appended (REQ-502); headers and footers
   by section under `include` (REQ-503), their blocks at `header:<rId>/<n>` and
   `footer:<rId>/<n>`, which `addresses.py` already writes (`footer:rId9/0` and `header:rId4/0`
   to `header:rId7/0` on the probe, as docx4j's oracle has them). docx4j's own were wrong before
   `8d32bf928`, its part's `toString()` prefixed (`org.docx4j.wml.Ftr@50bb1c1ffooter:rId9/0`, a
   shadowed local in `Extractor.headerOrFooter`), so the port copies the form, not that code
   before the fix; no fixture here was cut from it; comments in markup (REQ-504); field codes and
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

- **The suite**, copied from docx4j at `3dcbb3764` into `tests/fixtures/text-extraction/`:
  `phase0/` (the three fixtures, the five TSVs --- 188 token rows, 26 findings, 2 orders, 57 case
  statements, 12 goldens --- and the twelve goldens), `GOLDENS.md` (one level up in docx4j),
  `LevelOwnerTest`'s three Word documents, and the understood list. Provenance in
  `tests/README.md`, as CR-010's fixtures have it.
- **The harness**, `tests/text/conformance.py`: a port of `ExtractionConformance` --- REQ-1201's
  normalisation (CriticMarkup's brackets stripped, then whitespace, `\p{Z}`, `*`, `_`, `\`, `{`,
  `}` and `$`; U+200B, tag characters and both hyphens kept; an occurrence inside a longer token
  of the suite not counted), the scenarios (each token row in three views by default, its option
  column's scenario, and the supplemental ones), the combined option (since `a41de133b` and WD 0.15,
  a row's option column may name several options `; ` apart, as in `embedded: true; hidden: true`;
  the adapter sets each one, and the check needs each one's statement), the pending rule (a check
  pends when any statement it needs is not in the adapter's `implemented()`; since `3dcbb3764` a
  findings, order or golden row also needs its scenario's statements, as a token row does, and a
  golden, cut without an XHTML importer, needs `xhtml importer absent`), and `verify_golden`
  (Markdown image destinations written `…` on both sides). Its report is a TSV in docx4j's form.
- **The adapter**, `tests/text/adapter.py`, docx4j's `TextsAdapter` protocol:
  `implemented()` (the REQ ids this package meets, grown phase by phase) and
  `extract(fixture, view, scenario, format) -> (text, findings)`. Until Phase C it advertises
  `xhtml importer absent`. Phase C's MHT expansion is built in here, not an optional dependency
  as docx4j-ImportXHTML is in docx4j. The adapter therefore then takes the capability as a
  parameter, with the expansion switched off by an adapter-only hook for the absent run. The suite
  runs twice: absent, which is the gate's run and checks the goldens and case 44's default rows;
  and present, which checks the nine rows that need an importer.
- **The agreement test** (REQ-406): the walk against the resolved copy, as section 3.2.
- **CR-010's ten pins** (`tests/revisions/test_content_api.py`) turn to passing tests in Phase D;
  `xfail(strict=True)` makes each one's closing visible.
- **The gate**, at Phase D: no check fails; the only pending checks are docx4j's 51 and those of
  any content Q7 defers, each named; 12 of 12 goldens; the agreement test at zero differences;
  `LevelOwnerTest`'s 61 paragraphs as Word numbers them.

### 3.7 Out of scope

- **The inspector** (REQ-901 to 904; WD section 12): the paper's later CR (its D29), after the
  model.
- **Strict OOXML** (REQ-203, converted on load): an engine change, recommended as its own CR
  (Q8); the three Strict fixtures stay pending until then.
- **HTML** (REQ-705): this package has no HTML serialiser.
- **`math: unicodemath` and `symbols: unicode`**: not implemented in docx4j either.
- **Editing through the model**: it reads; the content API edits.

## 4. Phase A: the suite and the baseline

- Copy the suite (3.6) and port the harness and the adapter protocol.
- An adapter over today's paths records the baseline of section 1 as the suite's first report
  (648 pass, 324 fail, 51 pending of 1,023 at `3dcbb3764`; goldens 3 of 12), committed beside it.
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
- `w:altChunk`: plain text, HTML and XHTML through `lxml.html`, MHT; `altchunks: false`.
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
   `UNDERSTOOD` kept for `mce_preprocess` (3.3). *Recommended.*
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
   - Either answer conflicts with REQ-202 as written: under this one, the readers stay outside it.
     Like core-ts, this CR would ask the specification for an exemption for a reader whose contract
     is an emulated API, documented as that API's text. Q5's note for the WD becomes that request.
   - It narrows the text-extraction paper's section 10 row ("one text model replacing the five"),
     which the user accepted on 2026-10-17. So it is the user's call. D2 is decided for core-ts
     only (its section 6.1): until Q4 is decided alike here, the two ports' paragraph text differs
     for the same document. Taking the first answer here would leave it differing for good.
5. **Q5: the python-docx facade**: keeps python-docx's semantics as a declared emulation outside
   REQ-202, with a note proposed for the WD. *Recommended.*
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
    declarations.
12. **Q12: the HTML reader** (REQ-513 at WD 0.28). `html5lib`, a pure-Python parser of the
    [HTML] algorithm that builds an lxml tree, as an optional extra (`docx4j[html]`). Without it
    the reader is absent and a chunk raises `altchunk.not_expanded`, the suite's default run, and
    MHT goes the same way. *Recommended*, as docx4j keeps docx4j-ImportXHTML optional. The
    alternatives: `html5-parser` (C and faster, but its wheels are tied to lxml's version); or
    `lxml.html`, section 6's plan, which does not follow the [HTML] algorithm.

## 10. Risks

1. **Every text reader changes.** `Body.text`, `get_text`, `find`'s matches and offsets, the
   outline's counts and previews and `to_markdown` all return different text on the same document.
   This is a behaviour change in `0.x`, recorded in the CHANGELOG (section 7), and the README's
   examples show text.
2. **The WD is a working draft.** It is at 0.15 and its goldens can be recut. Port from
   `3dcbb3764` and record the commit; a recut is a hand-off like any other.
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
