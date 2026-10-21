# Changelog

The distribution `docx4j` (import name `docx4j_py`), on PyPI at
<https://pypi.org/project/docx4j/>. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the versions [Semantic Versioning](https://semver.org/) as far as `0.x` goes: what each
surface promises in `0.x` is the table at the end of this file. Each entry summarises what a
change request delivered and links the CR's own implementation notes, which are the record; the
CRs are under [`docs/change-requests/`](docs/change-requests/README.md).

## [Unreleased]

### Added

- **Tracked changes listed, accepted and rejected as Word resolves them**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md), Phases A to C,
  sections [10](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md#10-phase-a-implementation-notes-2026-10-16)
  to [12](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md#12-phase-c-implementation-notes-2026-10-18)):
  `docx4j_py.model.revisions`, the port of docx4j 17.3.3's `org.docx4j.model.revisions` (its
  CR-043), with `tracked_changes(pkg)`, `accept_all(pkg)` and `reject_all(pkg)` over **every
  story** --- the main document, footnotes, endnotes, comments, headers and footers --- and inside
  them content controls, text boxes in both branches of `mc:AlternateContent` (made typed for the
  resolution and written back in place, the rest of the branch untouched) and equations. The
  content API's `get_tracked_changes()`, `accept()`, `reject()`, `accept_all()` and
  `reject_all()` are views over it. Every kind listed is resolved, a table's property records
  among them (`w:tblPrChange`, `w:tblGridChange`, `w:tblPrExChange`, `w:trPrChange`,
  `w:tcPrChange`: one change per table when they change anything, accepted by dropping them and
  rejected by restoring the table, grid, rows and cells wholesale, a merged grid collapsed as
  Word's save collapses it), a section's `w:sectPrChange` (rejected, merged over the section) and
  `w:cellIns` / `w:cellDel` (a cell removed and its neighbour widened); records that change
  nothing, `w:cellMerge`, `w:numberingChange` and `w:numPr/w:ins`, which nothing lists, are swept.
  `TrackedChangeFilter` (authors, types, kinds, a date range, parts, a content control's content
  or a bookmark's extent) is taken by `revisions.accept_all`, `reject_all` and
  `tracked_changes(pkg, filter)`, and applies to single changes, so a group partly matched has
  only its matching pieces resolved; a plain callable is a filter too. `revisions.report(pkg)`
  counts by kind, author and part, with `to_dict()`. Held to Word 365's own Accept All and Reject
  All saves: against the eleven fixtures of docx4j's agreement run, copied into
  `tests/fixtures/revisions/`, no line of the structural comparison differs, the table, row, cell
  and section properties included, by this package's port of docx4j's `DocumentDescription` and
  by docx4j's own.
- **`Range.hyperlink`** ([CR-003](docs/change-requests/CR-003-content-api.md), section
  [21](docs/change-requests/CR-003-content-api.md#21-rangehyperlink-2026-10-12)): Office JS's read
  and write of a link, the one member docx4j-core-ts added to the Office JS subset the content API
  promises (`tests/office_js_subset.json`, now 231 members). The value is `address#location`: the
  address becomes an external relationship of the range's own part (a header's paragraph relates
  from the header) and the `r:id`, the location the `w:anchor`; reading gives the link the range is
  in, or `""`. Setting follows what docx4j-core-ts measured in Word: the runs are split at the
  span's boundaries and wrapped in one `w:hyperlink` with `w:history`, each taking the `Hyperlink`
  character style in place of its own and losing its direct colour (the style defined in
  `styles.xml` when the document lacks it); a span over another link's edge takes that link in,
  one inside a link replaces it; `""` or None unlinks, the runs staying; a caret, or a span across a
  content control, a field or a tracked change, is refused with `SpanError` before anything
  changes. Under change tracking the restyle is a `w:rPrChange`, and a link over a tracked
  insertion holds the `w:ins`. The call records a `"hyperlink"` `ChangeReport` naming the `.rels`
  it wrote, and a `dry_run` now takes back a relationship added with no part behind it.
  `docs/user/hyperlinks.md` is brought up to date; `get_hyperlink_ranges`, an `insert_hyperlink`
  sugar, the facade's `add_hyperlink` and the fix for text appended after a link joining it remain
  proposals.
- **One text model: `docx4j_py.model.text`**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md), Phases B and C, sections
  [12](docs/change-requests/CR-011-one-text-model.md#12-phase-b-implementation-notes-2026-10-21)
  and [13](docs/change-requests/CR-011-one-text-model.md#13-phase-c-implementation-notes-2026-10-21)):
  the port of docx4j's `org.docx4j.model.text` (its CR-046 phase 1), after docx4j-core-ts's
  CR-012. `extract(pkg, view=, **options)` gives the block model (`TextResult`: `TextBlock`s with
  their runs, fields, numbering, addresses and flags; comments; warnings), and
  `plain_text(result)` its text. The views are `accepted`, `original` and `markup` (`current` a
  synonym): the first two the text of the document Accept All or Reject All leaves, worked out
  without resolving anything and held equal to CR-010's resolved copy over 159 fixtures in two
  scopes; `markup` every run as stored, written as CriticMarkup. The scope is the specification's:
  text boxes after their anchoring paragraph (`body/3/textbox:0/0`), notes in reference order as
  `[n]`, headers and footers by section under `include`, comments in `markup`, field codes under
  `field_codes`, hidden text left out through the styles, placeholders, list labels under
  `list_labels`, the table of contents only under `toc: true`; the options are the specification's
  26 names verbatim. A reading never changes the package. Phase C (its CR-046 phase 2) reads the
  rest of the document: the document properties in the result's metadata and, under `metadata:
  true`, before the body; each picture an `IMAGE` run described by the options' describer (the
  author's alt text, `[image: ...]` under `alt_text: true`; `image_describer=` a host's hook) and
  each object an `OBJECT` run; legacy WordArt; a chart as a `CHART` block with docx4j CR-047's
  note, its title and its table; SmartArt's lines in reading order; a bound plain-text control's
  bound value (`sdt.binding_differs`); a plain-text, WordprocessingML or Flat OPC `w:altChunk`
  expanded where it is (HTML, XHTML, MHT and RTF raising `altchunk.not_expanded`: there is no HTML
  reader); an embedded WordprocessingML package under `embedded: true`; the glossary under
  `include: glossary`; equations as LaTeX; REQ-605's listed and grid tables; `[page n]` under
  `page_markers`; `layout: indent`; and under `fields: update` the fields that need neither
  layout nor fetching computed (`DOCPROPERTY`, `DOCVARIABLE`, the core-property fields, `SEQ`,
  `REF`, `QUOTE`, with the `\@`, `\#` and `\*` switches) and FORMTEXT as Word's update leaves
  it. Embedded workbooks and presentations and OLE 1.0 packages are deferred and raise
  `content.not_read`. Against docx4j's text-extraction suite, 1,284 checks pass, 7 fail and 293
  pend of 1,584 (941, 410 and 233 at Phase A; 1,146, 205 and 233 at B), the 7 Markdown's (Phase
  D), the 293 docx4j's 19, the inspector's 214 and the 60 of the deferred content; 9 of the 12
  goldens are exact, every plain-text one.
- **Markdown from the text model: `docx4j_py.model.text.markdown`**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md), Phase D, section
  [14](docs/change-requests/CR-011-one-text-model.md#14-phase-d-implementation-notes-2026-10-21)):
  `markdown(result, package=None, *, image_destination=None)` writes a reading as docx4j-markdown's
  `MarkdownExporter` writes it (WD REQ-703, REQ-704), the port of its `WmlToMarkdown` with a port
  of commonmark-java 0.30.0's Markdown renderer (`docx4j_py.model.text.commonmark`, BSD-2-Clause,
  named in `NOTICE`), so the escaping, the list markers and the block separators are docx4j's:
  over 277 documents in five readings each, 1,385 of 1,385 exports are docx4j's byte for byte,
  pictures' data URIs included. `markdown_blocks(...)` gives each block's Markdown, joining to the
  whole; `MarkdownExporter` and `MarkdownExportOptions` are docx4j's classes, `image_destination`
  (a callable given the image, its bytes and docx4j's file name) standing for docx4j's
  `imageDirPath`. With it the text-extraction suite gives 1,291 checks passed, none failed and 293
  pending of 1,584, and 12 of 12 goldens.
- **Charts, equations and field switches as docx4j reads them**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md#13-phase-c-implementation-notes-2026-10-21)
  section 13): `docx4j_py.model.chart` (`ChartModel`, `ChartExModel`, `ChartTable`,
  `ChartDescription`: a chart's title, its data as Excel formats it, its kind in Excel's names,
  its axes and secondary series; held to docx4j's 33 JSON goldens), `docx4j_py.model.math`
  (`OmmlToLatex`, docx4j's OMML to LaTeX for its subset) and `docx4j_py.model.fields`
  (`FldSimpleModel`, `apply_formatting_switch`, the `DOCPROPERTY` and `DOCVARIABLE` resolvers;
  docx4j's 138 switch cases), each imported on first use.
- **`mc:AlternateContent` at block, row and cell level, and the load mode that says what to do with it**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) Phase A, section
  [9](docs/change-requests/CR-012-type-kept-mc-branches.md#9-phase-a-implementation-notes-2026-10-21),
  the port of docx4j's CR-051 phase 1): the schema copy, refreshed to docx4j `2853b7280`, admits the
  element in every block container (the body, a header or footer, a cell, a note, a comment, a
  text box, a block-level content control or custom XML element), among a table's rows and among a
  row's cells, as Word 365 reads it there. `LoadOptions(mc_block_level=...)`, and
  `pkg.mc_block_level` over it for one package, read when each part is first parsed:
  `"preserve"` (the default) keeps each such element whole with every branch, as at run and
  paragraph level, so a part saved unchanged is canonically identical to its source; `"word"`
  replaces each on load by the branch Word reads (the first Choice whose `Requires` namespaces are
  all in docx4j's published understood list, else the Fallback; one with no understood Choice and
  no Fallback kept whole, WD REQ-305; one with no branch removed), adds the selected Choice's
  prefixes to the root's `mc:Ignorable`, and records each element taken out with its address in
  `part.resolved_alternate_content` (`ResolvedAlternateContent`), so a save writes what Word's
  save writes, text boxes apart. A part without the element pays one substring test. Readers of
  text and structure read Word's branch at block level in `preserve` mode: the text model and
  Markdown (with `mce.branches_differ`, raised in `word` mode from the record), `outline()`,
  `Body.paragraphs`, `element_at` and `traversal`; a view over content in a kept branch is
  read-only, and every mutating verb refuses it with `address.read_only`. The python-docx facade
  and `describe()` pass the element by; at row and cell level so did the content API, the text
  model and the revisions engine, as docx4j's CR-051 phase 1 does, until Phase A2 (under Changed). Held to docx4j's
  19 `mc-fun` probes and Word 365's saves of 14 of them, copied into `tests/fixtures/mc-fun/`.

### Changed

- **`to_markdown` is the text model's** ([CR-011](docs/change-requests/CR-011-one-text-model.md)
  section [14](docs/change-requests/CR-011-one-text-model.md#14-phase-d-implementation-notes-2026-10-21),
  REQ-202): `pkg.to_markdown`, `Body.to_markdown`, `Paragraph.to_markdown`, `Table.to_markdown`
  and `markdown_budget` are `markdown()` over the text model's reading, every keyword kept
  (`addresses=`, `view=`, `max_chars=`). The Markdown is docx4j's, which moves what 0.1.0 wrote:
  each text box's paragraphs after its anchor; endnotes `[^e1]` beside footnotes `[^1]`; hidden
  text left out; fields' results, bound values, charts (their note, title and table), SmartArt
  and equations as LaTeX; the table of contents left out (REQ-520); a heading by its effective
  outline level through the styles (a `HeadingN` style id that names no style with one is body
  text, as in docx4j); a GFM table with no padding
  inside the pipes (`|A|B|` and `|---|---|`) and a header row from `w:tblHeader`; lists by their
  numbering, a list interrupted by a paragraph going on at Word's count and an adjacent list
  taking the other delimiter (`1)`), a paragraph indented to an item's text a follow-on paragraph
  of it; a run code by its font through its styles (a `CodeChar` that the package does not
  define marks nothing); commonmark-java's escaping (`1\. not a list`); the document's Markdown
  ending with a line end. `view` takes `original` (Reject All's text) and `current`; `markup` is
  CriticMarkup written flat. `addresses=True` still writes the content API's address (paraId,
  else the ordinal) for a block of the body's story; a text box's, a note's or a chart's block
  has the text model's (`body/3/textbox:0/0`), which `element_at` refuses as `address.read_only`.
  Pictures still point at their parts (`media/image1.png`), where `markdown()` writes docx4j's
  data URI. A paragraph is written as its story writes it (`3. Plums`, `4. 2. Quince`).
- **`to_markdown` is slower on large documents**: on CR-003 section 16.4's 200-page document,
  `pkg.to_markdown()` is about 48 ms where 0.1.0's was 8 ms with the body parsed, and about 90 ms
  where it was 53 ms read fresh from a file, the main part's parse included (CR-011 section 14.9).
  A first `pkg.body` after it no longer parses the main part again (CR-002 section 16, below).
- **A read leaves every part as it was**: `pkg.to_markdown()` and `pkg.markdown_budget()` read the
  package with nothing unmarshalled, so a save afterwards writes `word/document.xml`, and
  `word/footnotes.xml` in a document with a footnote reference, from the source bytes, where
  0.1.0 unmarshalled both (through `pkg.body` and for the note definitions) and re-marshalled
  them, canonically identical. A first `pkg.body` afterwards adopts the tree the reading parsed.
  `outline()` still unmarshals the notes parts, through the package's bodies.
- **A read followed by an edit parses a part once**
  ([CR-002](docs/change-requests/CR-002-engine.md) section
  [16](docs/change-requests/CR-002-engine.md#165-implementation-notes-2026-10-21)): a part's
  `contents` adopts the tree `read_contents()` already parsed, instead of parsing the part's bytes
  again. The readers that leave such a tree are `pkg.to_markdown()`, the text model, the property
  resolver, the numbering emulator's styles and the font selector. A tree is adopted only where the
  load options (`strict`, `mce_preprocess`) are the ones it was parsed under. On CR-003
  section 16.4's 200-page document, `pkg.to_markdown()` then the first `pkg.body` takes 88 to
  92 ms, down from 131 to 136 ms, which is what body first costs. A resolver read then
  `style_definitions_part.contents` costs one parse of the styles part, not two. Saved output does
  not change. A resolver built before the styles part is unmarshalled now holds the live styles
  rather than a copy; `refresh()` after a style edit is still the rule.
- **`outline()` reads the text model** (CR-011 section 14, REQ-706): each entry's text, its
  heading level (the effective outline level, through the styles: a `HeadingN` style id without
  an outline level is body text, a style with one a heading whatever its name) and a table's
  preview are the text model's accepted view, and `stats.paragraphs`, `tables`, `words` and
  `chars` count its blocks (a hidden or a table-of-contents paragraph none, a joined pair one).
  Each text box's paragraphs are entries after their anchor (after the table, for one in a cell)
  at their `textbox:` addresses, with `OutlineEntry.textbox` (Q9). The entries are still one per
  block of the tree. `stats.tracked_changes` is unchanged. On the 200-page document `outline()`
  takes about 50 ms, where it took 16.
- **Dates are written and read as Word 365 writes them**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md) section 12.9):
  every revision the tracker writes carries `w:date` on the process's **local wall clock** with a
  `Z` that does not mean UTC, and the true UTC in `w16du:dateUtc`, with `w16du` added to the
  part's `mc:Ignorable`. "Now" is to the minute, as Word stamps it. 0.1.0 wrote the UTC in
  `w:date`, to the second, so Word showed this package's revisions and comments shifted by the
  reader's UTC offset. A new comment's `w:date` takes the same local form, to the minute, and its
  `w16cex:dateUtc` (where the document has that part) the UTC. Reading follows docx4j:
  `TrackedChange.date` is `w16du:dateUtc` where present, else `w:date` read as local time in the
  process's zone; `Comment.creation_date` is the document's `w16cex:dateUtc` where it has one,
  else `w:date` read as local time. Both are returned in UTC. A date this package writes reads
  back as the same instant. A naive `pkg.tracked_change_date` is still UTC; the `w:date` bytes of
  a fixed date now depend on the time zone.
- **A comment whose range a resolution empties goes whole, wherever its markers were, as Word 365
  removes it** ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md)
  section [12.8](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md#128-the-word-check-2026-10-21),
  fix 2, the port of docx4j `d18693760`, its CR-043 section 15): rejecting an insertion or
  accepting a deletion removes a comment whose range start was inside it **or whose markers lie
  either side of it**, once nothing with text is left in its range --- markers, reference run,
  entry, its extended entries and its replies. 0.1.0 removed only such a comment's range markers,
  as Word 2010 and Word 15 do; a comment written by 0.1.0's tracker on inserted text now goes with
  its rejection, as in Word 365. With the last comment the resolution takes away, the `comments`,
  `commentsExtended`, `commentsIds` and `commentsExtensible` parts go too, with their
  relationships and content types, and `people.xml` once no tracked change is left. A dry run's
  body removes no part. **A comment whose range spans paragraphs goes too** (section
  [12.9](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md#129-words-date-convention-2026-10-21),
  the port of docx4j `a8bc6cda9`), as Word 365's Reject All of two inserted paragraphs commented
  whole removes it: such a comment is judged once the resolution has joined the paragraphs whose
  marks it took away, and goes whole when its start and end then stand in one paragraph with
  nothing with text between. One whose range still crosses a paragraph mark stays.
- **A comment on tracked-inserted text is written as Word writes it**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md) section
  [12.8](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md#128-the-word-check-2026-10-21)):
  `insert_comment` on text in a `w:ins` puts the range start inside the `w:ins`, and the range end
  and the reference run after it. When the comment ends inside the insertion, the `w:ins` is closed
  there and the rest becomes a second `w:ins` with the same author and date and a new id. This is
  the form Word 2010 and Word 15 wrote (docx4j's `measure-word15.docx`), and docx4j-core-ts's since
  0.3.1. 0.1.0 put both range markers outside the `w:ins`, which widened a comment on part of an
  insertion to the whole of it. The reference run is still never inside a `w:ins`. Accepting the
  insertion leaves the comment on its text. Rejecting it now removes the comment, its entries and
  its replies, as Word 16's Reject All does; 0.1.0's form kept it. Tracking on or off makes no
  difference. A `w:del` or a move keeps the markers outside it.
- **What one tracked change is, and the counts** ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md)
  section 10): `get_tracked_changes()` groups as Office JS and Word's Review pane do --- touching
  insertions, deletions or formatting changes by one author are **one** change, a paragraph's
  mark a piece whose text is `\r` (so a deleted paragraph is one change, `"Two.\r"`), rows that
  touch one change, a row's text its cells a tab apart then `\r\n` --- where 0.1.0 listed every
  piece of markup apart; `accept_all()` and `reject_all()` count the same way. `kind` is docx4j's
  finer kind (`"run_insertion"`, `"paragraph_mark_deletion"`, `"move"`, `"insertion"`, ...) where
  it was `"run"`, `"mark"`, `"row"`; `to_dict()` gains `"part"`, and a change gains `part` and
  `pieces`. **`pkg.get_tracked_changes()` lists every story**, not only the main document's. **A
  view is valid until the next accept or reject on its package**, as Office JS's is: a stale
  one's `accept()` / `reject()` raises `TrackedChangeError` (`tracking.gone`); list again. A
  paragraph join keeps the **next** paragraph's properties and `w14:paraId` in both directions,
  as Word's does (docx4j CR-043's C10 and C10b), where 0.1.0 kept the first one's id on an accept
  and the original paragraph on a reject; so rejecting a paragraph this package appended at the
  end of a body leaves the appended paragraph's element, with the properties it records (below:
  a heading's, since Phase C). A date is `w16du:dateUtc` where the markup has it, else `w:date`
  read as Word's local wall-clock time (below). `OutlineStats.tracked_changes` counts the same
  grouped listing, over the stories the outline shows (the body's, and the package outline's
  headers and footers too), and is 0 once `accept_all` has run; `describe()`'s `tracking_on` is
  unchanged, since accepting and rejecting never change `w:trackRevisions`.
- **The colour merge rules are docx4j 17.3.3's** ([CR-009](docs/change-requests/CR-009-theme-colours.md),
  Phase A, section [10](docs/change-requests/CR-009-theme-colours.md#10-phase-a-implementation-notes-2026-10-12)):
  in the property catalogue a `w:color` or `w:u` stating only a theme colour is no longer empty,
  so it is inherited, and a colour inherits **as one** --- where a layer states a value or a theme
  reference, its value, reference, tint and shade replace the ones beneath together (`w:color`,
  `w:u`'s colour, `w:shd`'s pattern colour and fill, each its own, and every border edge), where
  0.1.0 merged them attribute by attribute. Effective properties --- the `PropertyResolver`'s, and
  the content API's reads through it --- change only where a colour carries a theme reference, and
  both towards what Word draws and towards docx4j 17.3.3 (its CR-034): a run's own
  `w:color w:val="auto"` under a theme-coloured style is `auto`, no longer the style's reference,
  and a character style's theme colour no longer keeps a paragraph style's tint or shade. Over the
  64 documents of `samples/` and `tests/fixtures/` two effective answers change and no content-API
  answer does. The parity goldens, regenerated by docx4j-core-ts at docx4j `2e381ef26` the same
  day, agree in both places.
- **A paragraph mark's character style is resolved**, as docx4j 17.3.2 resolves it (its
  `fb2a764f7`; CR-002 section 15.1 item 8): `PropertyResolver.get_effective_paragraph_mark_rpr`
  applies the mark's `w:pPr/w:rPr/w:rStyle` as a style level, then the mark's direct formatting,
  where 0.1.0 copied the `w:rStyle` and never read it. It sizes an empty paragraph and starts a
  list label's run properties, so both follow the mark's character style, as Word draws them.
- **`mce="resolve"` reads the branch a reader of text takes**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) sections 3.3 and
  [12.2](docs/change-requests/CR-011-one-text-model.md#122-resolve-and-the-save-q3-q11), its Q2
  and Q3): the first `mc:Choice` whose `Requires` namespaces are all in docx4j's published
  understood list (118 URIs, now `docx4j_py/resources/understood-namespaces.txt`), matched by
  **namespace URI** through the declarations in scope where the Choice was parsed (a `wps` Choice
  under the prefix `shape2010` is understood); else the `mc:Fallback`; else, where there is none,
  the first Choice (before, nothing). An empty `Requires` is no longer understood. `text_of`,
  `walk`, `find` and the content API read through it. `mce_preprocess` keeps what the model can
  type (`UNDERSTOOD`). Each Choice's `Requires`, `mc:Ignorable` and `mc:MustUnderstand` carry the
  bindings they were parsed with (`child.source_bindings`); the fork's parser gives them
  (`docx4j-xsdata` stage 10, unreleased).
- **`get_text(view="markup")` is the text model's markup view**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) section
  [12.6](docs/change-requests/CR-011-one-text-model.md#126-the-readers-q4-q10)): on a body or a
  paragraph, and their `text_budget`s, the runs as stored with each revision as CriticMarkup and
  each comment at its reference; a range refuses it with `SpanError` (`range.no_markup_view`), as
  it refuses `original`. `current` is read as `accepted`. The other views, and every other Office
  JS reader, keep Office JS's text. Since Phase C (section 13) the markup view of a body has the
  text model's other content too: equations as LaTeX, charts, SmartArt, WordArt, bound values,
  expanded `w:altChunk`s and the table forms.
- **A numbering level belongs to one paragraph style**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) section
  [12.7](docs/change-requests/CR-011-one-text-model.md#127-req-510-and-the-parity-goldens-risk-3),
  WD REQ-510): the one its `w:lvl/w:pStyle` names, else the first paragraph style in `styles.xml`
  whose own `w:numPr` names a `w:num` on the same `w:abstractNum` at that level
  (`Emulator.level_style_id`, docx4j's `levelStyleId`). A paragraph numbered through any other
  style gets no label and is not counted, as in Word: the three documents of docx4j's
  `LevelOwnerTest` are numbered as Word numbers them, where 27 of 61 paragraphs differed. A
  paragraph naming no style is in the default paragraph style, and one with no `w:pPr` is resolved
  as one with an empty one (docx4j `c283dffdb`).
- **`mce_preprocess` selects over docx4j's published understood list**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) Q8, section
  [9](docs/change-requests/CR-012-type-kept-mc-branches.md#9-phase-a-implementation-notes-2026-10-21)),
  as a reader of text and the `word` mode do, where it selected over the namespaces the generated
  model types (`UNDERSTOOD`, which stays the prefix table's own list). It now takes the branch Word
  365 takes: `w16sdtfl` where it took `w16cid` on Choices written newest first, and a chart's
  chartex Choice (`cx1`) where it took the Fallback picture (`waterfall.docx`). No other document of
  `samples/`, `tests/fixtures/` or docx4j-core-ts's fixtures selects differently.
- **Tracked changes inside a block- or paragraph-level `mc:AlternateContent` are listed, accepted
  and rejected** ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section 4.1 item 6,
  docx4j's `TrackedChangesAlternateContentTest`): the changes of the branch Word reads are listed
  once, where the element stands, each with the other branches' change of its kind at its position
  as a twin that resolves with it; `accept_all` and `reject_all` resolve every branch. Until now
  such an element was read as holding nothing. The revisions engine's trees answer the element
  with Word's branch; a row- or cell-level one too since Phase A2 (below).
- **A view over content in a kept `mc:AlternateContent` branch is read-only**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section 4.1 item 5): an edit there,
  at any level, would change one branch alone, so the content API raises `AddressError`
  (`address.read_only`), as for a text box's address (CR-011 Q9). Load with
  `LoadOptions(mc_block_level="word")` to edit Word's branch of a block-level element in place.
- **A row- or cell-level `mc:AlternateContent` is read as Word reads it, in `preserve` mode too**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) Phase A2, section
  [10](docs/change-requests/CR-012-type-kept-mc-branches.md#10-phase-a2-implementation-notes-2026-10-21),
  the port of docx4j's CR-051 phase 2): Phase A kept such an element whole and its readers passed it
  by, so a row's or a cell's content in its branches was missing. Now:
  - the rows and cells in its branches are typed (`Tr`, `Tc`, where a branch's `w:tc` was a
    wildcard tree) and written back as `w:tr` and `w:tc`. Code that looked for an `AnyElement`
    `w:tc` in an `mc:Choice` among a row's cells now finds a `Tc`;
  - the text model and Markdown read Word's branch in the element's place, with
    `mce.branches_differ` at its address, the same in both modes. So do `Table.rows`,
    `TableRow.cells`, `Body.paragraphs`, `outline()` and the addresses (`body/1/1/0`, a row in the
    element's branch);
  - the views over that branch are read-only: every edit raises `address.read_only`, and so does a
    table's verb that would edit one of them, refused before anything changes;
  - a tracked deletion of a table, a row or a body that would mark content in a kept element is
    refused; untracked, the element goes whole;
  - the revisions engine lists Word's branch's row and cell changes once, the other branches' as
    twins that resolve with them, and removes a table left with no rows only when no branch has
    one;
  - `outline()` gives a cell's blocks under the table at their full paths through any wrapper, so
    a row-level content control's rows no longer come out as `tc` entries.

### Fixed

- **The text model's numbering, document dates and table content**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) section 14, found by the byte check
  against docx4j): a block's `TextNumbering` has its level's `levelText` and `start` (they were
  None and 1 for every list, an attribute misnamed); the core properties' dates are as the part
  writes them, `lastPrinted` among them (it was left out, and `created` had nine fractional
  digits for seven); a table given to `extract` as content is one `TABLE` block, as docx4j reads
  it (it was its cells' paragraphs). A dry run reads the styles it added, through a property
  resolver of its own.
- **A save declares every prefix an `mc:Choice` names**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) section
  [12.2](docs/change-requests/CR-011-one-text-model.md#122-resolve-and-the-save-q3-q11), its Q11,
  by docx4j CR-052's rule): a part re-marshalled after it was read left a prefix only a Choice's
  `Requires` named undeclared (docx4j's probe's `zz9` and `shape2010`, Word's own `cx1` on a
  chartex chart), which Word 365 repairs. Now each is declared where Word 365 puts it: a prefix
  bound as docx4j's table binds it on the root, any other on the `mc:Choice` that names it, ahead
  of `Requires`, with the binding the source gave it, which travels with the Choice when it is
  moved or copied; a Choice's own `mc:Ignorable` and `mc:MustUnderstand` likewise. `to_xml`
  follows the same rule. The root's `mc:Ignorable` prefixes take the source's binding before the
  table's.
- **`get_text(view="markup")` raises instead of returning neither view**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) section
  [11.5](docs/change-requests/CR-011-one-text-model.md#115-q10), its Q10). 0.1.0 took the value
  and answered with the text outside every tracked change on a body or a paragraph (5,093
  characters on docx4j's probe, against 5,166 accepted and 5,148 original), and with the accepted
  text on a range. `Body.get_text`, `Paragraph.get_text`, both `text_budget`s and `Range.get_text`
  now raise `ContentError` with the code `text.no_markup_view`, whose hint names the views they
  take and `to_markdown(view="markup")` for the changes as CriticMarkup. CR-011 Phase B gives
  `get_text` the text model's markup view (under "Changed"). Any other view name raises
  `text.view_invalid`, for the same reason.
- **A comment's range is read right when its markers are inside a `w:ins` or a hyperlink**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md) section 12.8).
  A marker's offset was counted from the start of the holder it sat in, not the paragraph's. So a
  range starting in an insertion or a link that does not open the paragraph read from the
  paragraph's start. Word writes the range start there for a comment on inserted text.
- **The tracked-change listing no longer reports a document clean when it is not**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md); CR-003
  section 22.3): after `accept_all()` 0.1.0 listed no change while revisions Word still shows
  remained in text boxes, on moved and formatted paragraph marks, in table, row, cell, grid and
  section records, at content-control markers and in equations. Each is now listed and resolved,
  and `OutlineStats.tracked_changes` counts the same listing.
- **Rejecting an inserted field takes the whole field** (docx4j CR-043's C01, C16), where 0.1.0
  left a lone `w:fldChar` separator and its result with no begin and no end.
- **Equations' tracked changes are listed and resolved**, in all three of Word's forms: inside
  `m:r`, an equation structure's `m:ctrlPr` (rejected, the structure goes whole) and `w:ins` /
  `w:del` holding math. The equations fixture lists the 12 changes Word counts, where 0.1.0
  listed none.
- **Moves Word wrote with range markers are resolved** (C15, check 18): the halves are paired by
  `w:name` on their range starts, the four range markers and the moved paragraph marks go, and
  edited and broken moves resolve as Word's check 20 saves do.
- **Rejecting a list change restores the list as Word does** (check 25): a restored `w:ind` equal
  to the list level's own is dropped, and a `w:numPr` with no `w:ilvl` gets `w:ilvl` 0.
- **A note goes with its reference** (C12; the editor's review file), **a bookmark touching
  removed content goes whole** (C11, C11b), **a comment range emptied by a change goes**, and a
  comment whose anchor goes takes its `commentsExtended`, `commentsIds` and `commentsExtensible`
  entries and its replies with it; an image, a dropped section's headers and footers or a
  hyperlink's relationship goes when nothing references it any more.
- **Rejecting a paragraph appended after a heading gives the heading back**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md) section 12;
  docx4j-core-ts's check 29). With tracking on, a paragraph appended at the end of a body or a cell
  takes the old final mark (CR-003 section 16.10), as Word's Enter at the end of a paragraph does;
  when its properties differ from the paragraph before it --- a Normal paragraph after a Heading 1
  --- it now records that paragraph's in a `w:pPrChange`, the markup Word writes, and a style set
  on it afterwards (`insert_paragraph(style=...)`, the facade's `add_heading`) is recorded against
  them. 0.1.0 wrote none, so Reject All, in Word as here, left the heading a Normal paragraph.
- **Headers, footers, notes and comments are resolved too**, by
  `docx4j_py.model.revisions.accept_all(pkg)` / `reject_all(pkg)`, as Word's Accept All covers
  them; `Body.accept_all()` resolves its own story, with Word's side effects in other parts.
- **`Range.insert_content_control` keeps what lies between the runs it wraps**
  ([CR-003](docs/change-requests/CR-003-content-api.md) section 21.9): the control takes the whole
  slice from the span's first run to its last, so a bookmark or a comment's range markers between
  them go inside in their place, where 0.1.0 moved the runs alone and left the rest after the
  control; and a tracked deletion or move between the runs is refused (`range.crosses_holder`)
  before anything changes, as `Range.hyperlink` refuses it. docx4j-core-ts found and fixed the same
  slice in its `insertContentControl`.
- **The chartex binding reads a histogram's bins**
  ([CR-001](docs/change-requests/CR-001-object-model.md) section 14.12). Excel writes them as
  `<cx:binSize val="10"/>` and `<cx:binCount val="6"/>`. `docx4j_py.oart.chartex_2014` typed them
  as element text, as [MS-ODRAWXML] does, so it read `""` and could not write the object back.
  - `CTBinning.bin_size_or_bin_count` is now a `CTBinSize` or a `CTBinCount`, each with a `val`.
  - The schema file is docx4j CR-040's (`28bdf831c`), copied verbatim.
  - Loading and saving a package was never affected: the engine keeps a chartex part as an XML tree.
- **The chartex binding reads `cx:externalData`'s `cx:autoUpdate`**
  ([CR-001](docs/change-requests/CR-001-object-model.md) section 14.12). Word and PowerPoint write
  it in the chartex namespace, `cx:autoUpdate="0"`, where [MS-ODRAWXML] declares it unqualified.
  So `CTExternalData.auto_update` read `None`, the skipped-content report named the attribute,
  and writing the object back left it out. It now reads `False` and is written back qualified.
  - The schema file is docx4j CR-040 phase 2's (`14fcb2f4e`), copied verbatim.
  - Loading and saving a package was never affected.
- **A block-, row- or cell-level `mc:AlternateContent` is no longer lost**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section 1.1 item 1): the parser
  skipped it (`part.skipped`, and `strict=True` raised), and a part read and saved lost it whole,
  every branch. It is kept whole by default, or resolved as Word reads it in `word` mode.
- **A loop of `insert_paragraph` is no longer quadratic in the body's length**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section 9.10): the view each insert
  returns was found by scanning the body's content for the new paragraph; `Body.view_for` now takes
  the list the caller knows, as `insert_*`, `body[i]` and iterating the body do.
- **`ContentControl.form` of a block-level control whose content begins with a kept
  `mc:AlternateContent`** ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section
  10.3) read `"run"`; it reads Word's branch and gives `"block"`.
- **A tracked `Body.clear()` over a kept block-level `mc:AlternateContent`**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section 10.6) raised
  `address.read_only` part-way, the paragraphs before the element already marked deleted; it is
  refused before anything changes.
- **A content control straight in an `mc:Choice` or `mc:Fallback` keeps its content**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section 1.1 item 2): every
  `w:sdt` there was typed as a block-level control, so a run-level one's runs were skipped and lost
  on save, and `wml(...)` refused it. It is now bound as its container binds it (`SdtBlock`,
  `SdtRun`, `CTSdtRow`, `CTSdtCell`) and written back as `w:sdt`, through the fork's stage 11
  (`docx4j-xsdata` after `26.2.1`; this package is never released against an unreleased fork).

## [0.1.0] --- 2026-09-20

The first release: WordprocessingML only. `.pptx` and `.xlsx` load and save through the generic
OPC path with every part byte for byte, but nothing in them is typed; PresentationML and
SpreadsheetML are CR-001 Phase D and CR-002 Phase C, for `0.2` and `0.3`.

### Added

- **The object model** ([CR-001](docs/change-requests/CR-001-object-model.md), Phases A to C,
  sections [12](docs/change-requests/CR-001-object-model.md#12-phase-a-implementation-notes-2026-09-12),
  [13](docs/change-requests/CR-001-object-model.md#13-phase-b-implementation-notes-2026-09-12) and
  [14](docs/change-requests/CR-001-object-model.md#14-phase-c-implementation-notes-2026-09-12)):
  the whole of WordprocessingML and what it embeds (DrawingML, OMML, VML, `w14` to `w16`, the
  relationships and `docProps` schemas) as slotted dataclasses generated from docx4j's own schema
  tree by the xsdata fork, one package per namespace under docx4j's class names; every object a
  `Child` with a parent pointer and `deep_copy`; the `el` object factory (2,061 element names over
  63 namespaces), the `p` / `r` / `t` / `tbl` sugar, `wml(...)` fragments and `to_xml`, `walk`,
  `find`, `text_of` and the rest of the traversal; lenient parsing with a per-part skipped-content
  report; `warm_up()`; docx4j's prefix table. Refreshed to docx4j `c363a969f` on 2026-09-20
  ([14.10](docs/change-requests/CR-001-object-model.md#1410-schema-refresh-2026-09-20) and
  [14.11](docs/change-requests/CR-001-object-model.md#1411-follow-up-refresh-to-docx4j-c363a969f-2026-09-20)).
- **The Open Packaging engine** ([CR-002](docs/change-requests/CR-002-engine.md), Phases A and B,
  sections [12](docs/change-requests/CR-002-engine.md#12-phase-a-implementation-notes-2026-09-12)
  and [15](docs/change-requests/CR-002-engine.md#15-phase-b-implementation-notes-2026-09-19)):
  `PartStore` and `PartSink` (zip, directory, memory, flat OPC), `PartName`, content types,
  relationships, typed parts with lazy `contents`, `load` and `save` with every untouched part
  written back byte for byte and every re-marshalled one canonically identical to its source,
  lossless `mc:AlternateContent` with a resolving traversal, `mc:Ignorable` on write,
  `create_package` (Word 2013+ compatibility mode, a theme part); and docx4j's resolution
  utilities --- `PropertyResolver` and `StyleUtil` over a property catalogue, the numbering
  emulator, and `org.docx4j.fonts`' decision half --- with zero differences on all 45 of docx4j's
  parity goldens. Departures from docx4j are section 12.3.
- **The content API** ([CR-003](docs/change-requests/CR-003-content-api.md), every phase,
  sections [10](docs/change-requests/CR-003-content-api.md#10-phase-a-implementation-notes-2026-09-16)
  to [20](docs/change-requests/CR-003-content-api.md#20-phase-j-implementation-notes-2026-09-18)):
  `Body`, `Paragraph`, `Range`, `Font`, `Table`, `TableRow`, `TableCell`, `InlinePicture`,
  `ContentControl`, `Comment`, `TrackedChange`, `List` and `ListItem` in Office JS's vocabulary;
  the agent surface --- paraId-first addresses, `outline()`, `find()`, `describe()`, a
  `ChangeReport` on every mutating call, `dry_run`, `DocumentSession`, errors with a `code` and a
  `hint`; markdown out and in with addresses; tracked changes and comments as the audit trail, with
  `accept_all()` and `reject_all()`; custom XML parts, XML mapping and the typed content-control
  kinds with `describe()` / `fill()` for templates; `insert_ooxml` from a flat OPC package;
  `to_api_script`, the reveal codes; and the **python-docx facade** `docx4j_py.docx` (305 of
  python-docx 1.2.0's 312 members, a parity test with python-docx as the judge).
- **Packaging** ([CR-007](docs/change-requests/CR-007-pypi-release.md), Phase B, section 10):
  the distribution `docx4j` with `py.typed`, `NOTICE`, this changelog, `RELEASING.md`, a short PyPI
  readme whose every example the suite executes (`tests/test_readme.py`), CI on Python 3.12 to
  3.14, and `scripts/release_check.sh`, which builds the wheel and gates it in a clean environment.

### Dependencies

- `docx4j-xsdata[lxml] >=26.2.1,<26.3` --- the xsdata fork the generated modules import, published
  as [docx4j-xsdata 26.2.1](https://pypi.org/project/docx4j-xsdata/26.2.1/) by CR-007 Phase A
  (section 9). The upper bound is where the fork's upstream base moves and the bindings are
  regenerated.
- `markdown-it-py >=3.0`, imported on the first `insert_markdown`.
- Python 3.12 or later.

### Notes

- **One generated field name is not what an earlier checkout had**: the `w:rPr` of `m:ctrlPr`
  (`math.CTCtrlPr`) was `schemas_openxmlformats_org_wordprocessingml_2006_main_r_pr` and is
  `r_pr_1`, with the `a:rPr` docx4j CR-025 admitted as `r_pr_2` (CR-001 section 14.11). Nothing
  released carried the old name, so this is a note, not a breaking change.
- `from docx4j_py.wml import sdt` is the content-control builder until something imports the
  module `docx4j_py.wml.sdt`, after which it is the module (CR-003 section 11.5); write
  `from docx4j_py.wml.sdt import sdt`, which is unambiguous.

### What `0.x` promises

| Surface | Stability in `0.x` |
|---|---|
| `docx4j_py.load`, `create_package`, `WordprocessingMLPackage`, the parts, `PartStore` / `PartSink` | Stable: docx4j's names and behaviour; departures are recorded in CR-002 section 12.3 |
| The content API (`Body`, `Paragraph`, `Range`, `Table`, ..., the agent surface, markdown) | Stable in shape: Office JS's vocabulary is borrowed, not invented; a member may be added in a minor, none renamed |
| `docx4j_py.docx`, the python-docx facade | Stable: its promise is `tests/python_docx_subset.json`, 305 of 312 members |
| The generated model (`docx4j_py.wml` and the rest, `el`) | Regenerates with the schema; class and field names are docx4j's and do not move, but a schema refresh adds fields and classes in a minor |
| `docx4j_py.model` (resolver, `StyleUtil`, fonts, numbering) | Stable: docx4j's, zero differences on the goldens |
| Anything under `codegen/`, `scripts/`, `baseline/` | Not shipped in the wheel; in the source distribution for the suite |

[Unreleased]: https://github.com/plutext/docx4j-python/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/plutext/docx4j-python/releases/tag/v0.1.0
