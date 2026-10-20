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

### Changed

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
  as written. `OutlineStats.tracked_changes` counts the same grouped listing, over the stories
  the outline shows (the body's, and the package outline's headers and footers too), and is 0
  once `accept_all` has run; `describe()`'s `tracking_on` is unchanged, since accepting and
  rejecting never change `w:trackRevisions`.
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

### Fixed

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
