# CR-028: manual redline to tracked changes --- the port of docx4j CR-068, an explicit, opt-in conversion

**Status:** Proposed 2026-11-02, at the user's word ("draft those 2 CRs"). docx4j CR-068's phase 3 handed the port to
this repository on 2026-10-31 (its section 13, docx4j-python-64). Five phases (section 6). The questions are in
section 8, each with a recommendation. Nothing built.

*The number:* CR-027's status line names CR-028 as the proposed number for this port.

**Depends on:**
- **docx4j CR-068** (registry `docx4j/CR-068`, done), as the release tag `docx4j-17.3.5` (`64fae4cd0`) has it:
  `529caac97` (phase 1, with `Texts.addresses` made public), `3fa2aab46` (phase 2, the opt-ins), `c39181b29` and
  `be09f382a` (Word's second and third runs), and `307c69e34` (the comment cards). Each is an ancestor of the tag.
  `60b6b2ca4` is phase 3's hand-off and changes only the CR's text. None of these files has changed on
  `VERSION_17_3_6` since.
- **docx4j's oracle and tests** at the tag: `docx4j-core-tests/src/test/resources/revisions/cr068/` (58 files),
  `ManualRedlineTest` (15) and `ManualRedlineOptInsTest` (12).
- **The portfolio:** `docs/mcp_tool_vocabulary.md` section 5 (the verb `manual_redline_to_tracked_changes`, whose
  result is `ManualRedlineReport.toJson`), and `docs/date_handling_across_ports.md` section 3.6 point 5 (a new
  revision is stamped to the minute).
- **This repository:**
  - CR-010's engine and its `DocumentDescription` port (`tests/revisions/description.py`);
  - CR-003 Phase F's tracker (`model/content/tracking.py`) and CR-025's `pkg.date_time_zone`;
  - CR-011's `TextBlock.source`, CR-012's typed branches, CR-002 section 16's `read_contents`, CR-020's placement.
- **Counterparts:**
  - docx4j-core-ts `core-ts/CR-019`, proposed, D1 to D8 awaiting the user. Its D1 makes the options a plain
    object, and its D2 stamps to the minute.
  - docx4j-mcp `mcp/manual-redline-tool`, built on a branch.
  - `../docx4j-mcp-py`, once CR-006 is built.
- **CR-027** (decided 2026-11-02, nothing built) plans to copy one file of the same oracle (section 5.1).

## 1. Why

A manual redline is a document whose author struck the old words through and underlined or coloured the new ones by
hand, instead of turning on Track Changes. The file holds formatting, not revisions. Measured at `ed12ae7` on
docx4j's probe base, `manual-redline.docx`:
- `accepted`, `original` and `markup` all give `M01 Rate: 1.316 1.335 per cent.`, and in a cell
  `M02 cell:\t1.3161.335`, a number the page never showed;
- `revisions.tracked_changes` lists one change, the base's own insertion (M07), so Accept All and Reject All leave
  every struck word.

That is right for an extraction, and the user ruled in docx4j that it stays so: no view treats a strike as a
deletion. docx4j built instead an operation that a reviewer runs on purpose. It turns the strike into a tracked
deletion by the author `manual redline`, so that Accept All, Reject All and every reader follow. This package has no
such operation.

## 2. What docx4j built

### 2.1 The default

`ManualRedline.toTrackedChanges(pkg[, options])` converts in place and returns a report. `preview(pkg[, options])`
is the same walk with nothing written; since phase 2 it no longer converts a clone. The classes are 1,009, 250 and
356 lines.

**What converts** (D2, D5, D6, D8, D10). A `w:r` whose own `w:rPr` holds `w:strike` or `w:dstrike` on is wrapped in
a `w:del`. Each `w:t` becomes `w:delText`, and tabs, breaks, symbols and drawings stay in the run. The strike is
removed, and an `w:rPr` left empty goes with it.

The walk covers the main document, the notes, the headers and footers, and their tables, content controls,
hyperlinks and smart tags. Text boxes are walked in every `mc:AlternateContent` branch: the first branch holding
boxes is reported, and the others convert as twins with their own ids, as Word writes them. The comments are only
reported (`comment`) without the opt-in. A run holding only a comment reference, a note's own mark, a separator or
`w:lastRenderedPageBreak` is neither converted nor reported.

**Left alone and reported:**
- a run in `w:ins`, `w:del`, `w:moveFrom` or `w:moveTo`, or whose strike is itself a `w:rPrChange`:
  `in_tracked_change`;
- a complex field's runs, `begin` to `end`, the depth carried across the story's paragraphs; a `w:fldSimple`'s:
  `field`;
- math: `math`;
- a run struck only by its styles: `style_strike`.

Consecutive runs left for one reason in one paragraph make one entry. A run whose own `w:strike w:val="0"` turns a
style's strike off needs nothing. `w:trackRevisions` is not touched (D12).

**Grouping** (D8). Adjacent converted runs of one content list share one `w:del`. A `w:proofErr` between them does
not part them; anything else does, a bookmark or a comment range among them.

**The separating space** (D13, on by default; `separatorSpace(false)` turns it off). It applies only where the
group's text neither starts nor ends with whitespace:
- (a) where the text shown before ends in whitespace, or there is none, and the text after begins with a space,
  that following space is taken;
- (b) otherwise, where nothing follows or what follows begins with one of `.,;:!?)]}”’»`, the space before is taken.

The space comes only from the first or last `w:t` of a plain run (not tracked, not a field's, not converted). It is
split off with that run's properties, less `w:rPrChange`. It goes into the group's `w:del` where the run is beside
it, else into a `w:del` of its own (M09, the space after a struck hyperlink). Word's own deletion took the space in
the main story only; docx4j takes it in every story.

**The paragraph mark** (D9 as revisited). A struck mark, not already tracked, in a paragraph that converted
something and shows nothing after Accept All, is deleted (`w:pPr/w:rPr/w:del`, its strike removed). That happens
only where the paragraph holds no `w:sectPr` and the next item of its container is a `w:p`. Otherwise the mark is
left and reported `paragraph_mark`: a cell's last paragraph, one before a table or a control, a story's last.

### 2.2 The opt-ins (phase 2; D7, D14 to D16)

- **`styleStrike(true)`.** A run struck only by its styles converts, kind `style_strike`. Its styles are its
  character style, its paragraph's style and the defaults, through `PropertyResolver.getEffectiveRPr`; a table
  style is not consulted. The run is given its own `w:strike w:val="0"` (or `w:dstrike`), so that Reject All shows
  it plain.
- **`underlineAsInsertion(true)`.** A run whose own `w:u` names a pattern other than `none` becomes a `w:ins`, and
  its `w:u` is removed. A `w:u` with no pattern inherits, and is not taken. A run both struck and underlined becomes
  a deletion with its underline kept, counted in `strikeAndUnderline`.
- **`colourAsInsertion(hex...)`.** A run becomes a `w:ins`, its colour removed, when its own `w:color` is a named
  `RRGGBB` value (compared upper-case, never a theme colour) and it has no underline, highlight or shading of its
  own. Anything not six hex digits is refused, `auto` among them.
- **`comments(true)`.** The comments convert, and a converted run there keeps its mark (D14). Word 365's comment
  card shows no tracked change. A deleted run without its strike shows plain, as if kept. One that keeps its strike
  is not drawn, and a "Missing content" notice points to the Revisions pane (`comment-cards.docx`, docx4j section
  12.7).
- **Spaces beside insertions.** An insertion takes its separating space by the rejected text (D16). The one space
  between a deletion and the insertion replacing it (a plain run holding exactly `" "` in one `w:t`) is written
  `w:ins` holding `w:del`, so that it goes in either outcome (D15).
- **`within(control)`, `withinBookmark(name)`, `parts(part...)`.** The range is `TrackedChangeFilter`'s, read over
  every story before anything converts. A run converts only inside it, and a mark only if its paragraph is.
- **`auditProperty(true)`.** `docProps/custom.xml` gains `docx4j:manualRedline`. Its value is the instant and the
  options' `toString()`: `2026-10-31T00:00:00Z; author=manual redline, date=2026-10-31T00:00:00Z,
  separatorSpace=true, auditProperty=true`. A preview writes none.

### 2.3 Ids, author, dates, and the report

- **Ids.** They count from one above the largest annotation id in the story parts, read first. That includes every
  `CTMarkup` (revisions, bookmarks, ranges, `w:comment`), the comment ranges and the comment references.
- **Author.** `manual redline` (`DEFAULT_AUTHOR`); an empty author is refused.
- **Date** (D4 as clarified). The operation's instant to the second by default; `date(null)` writes none.
  - `w:date` is its wall clock in `effectiveDateTimeZone()`, with a `Z`. It is the one Word shows: phase 0 found that
    Word 365 ignores `w16du:dateUtc`, and shows no date for a mark that has `dateUtc` alone.
  - `w16du:dateUtc` is the instant, and `w16du` joins the part's `mc:Ignorable`.
- **`ManualRedlineReport`** has one `Change` per `w:del`, `w:ins`, separating space and mark, one `Left` per run or
  mark left (reason, part, address, text), and counts. A change has a kind (`strike`, `double_strike`,
  `style_strike`, `underline`, `colour`, `separator_space`, `paragraph_mark`), a type (`deletion`, `insertion`,
  `both`), `own` (false for a space that went into the change beside it), the part, the address, the text and its
  runs.

  `Texts.addresses(pkg, options)` reads the addresses (`body/1`, `body/11/textbox:0/0`) before the conversion,
  mapping each block's source to its address by identity under the accepted view with notes, comments, headers and
  footers. A twin's paragraph takes its outer paragraph's address. `toJson()`, which docx4j-mcp returns, has the keys
  `preview`, `deletions`, `insertions`, `runs`, `strikeAndUnderline`, `byKind`, `byPart`,
  `changes[{kind,type,own,part,address,text}]` and `left[{reason,part,address,text}]`.

### 2.4 The oracle: Word 365's three runs (16.0.20326, the user's VM, 2026-10-31)

`cr068-probe.ps1` lists every revision (`cr068-log.txt`) and saves Word's Accept All and Reject All of each file as
docx and as text:

| File | What it is | Word's revisions |
|---|---|---|
| `manual-redline-d9.docx` | phase 0's hand-written D9 prediction: the default without D13 | 17 |
| `manual-redline-phase1.docx` | docx4j's default conversion | 18 |
| `manual-redline-phase2.docx` | the base with the opt-ins | 29 |
| `manual-redline-2-phase2.docx` | the second base, N01 to N10, with the opt-ins | 10 |

Accept All and Reject All left none. docx4j's tests compare its own Accept All and Reject All of each conversion
with Word's saves through `DocumentDescription`. For the second base they allow for three normalisations of Word's
save: N06's `w:u none` dropped, the unreferenced footnote removed, the empty header removed.

### 2.5 Where the CR's prose is narrower than the code

The code at the tag is what this CR ports:
- **Style strike.** Section 3.2 names "a style, a table style or the paragraph mark's style". The code consults the
  run's style, the paragraph's style and the defaults, as its section 11.2 says.
- **Strike and underline.** Section 3.3 says such a run is "reported `strike_and_underline`". No kind or reason has
  that name: the change is a `strike`, and the count `strikeAndUnderline` records it.
- **`mc:AlternateContent`.** Section 11.2 names block- and paragraph-level elements. The walk takes every branch of
  any `mc:AlternateContent` it meets in a block list, so row- and cell-level ones too. The probe's only one is
  run-level (the text box), so no test reaches the others.
- **The report.** Section 3.1 says "per story", but the code counts by part and by kind. The left-by-reason count and
  each change's runs are getters only, not in `toJson()`, and `toString()` is the JSON.
- **Unstated.** A group's kind is its runs' kind where they agree, else `strike` for a deletion and `underline` for
  an insertion. A style-struck run in a `w:ins` is reported `style_strike` without the opt-in, `in_tracked_change`
  with it.

## 3. What exists here (measured 2026-11-02 at `ed12ae7`)

Measured with scratchpad scripts reading docx4j's files at the tag; nothing in the repository changed.

- **Nothing converts a strike.** The content API can delete a struck range under `change_tracking_mode =
  "TrackAll"`, one range at a time, without D8, D9 or D13. It keeps the strike on the deleted run, as Word's own
  deletion does, so Reject All gives struck text back.
- **The reading half already agrees with Word.** `revisions.accept_all` and `reject_all` were run on docx4j's four
  conversions, and each result was saved, loaded and described by `tests/revisions/description.py`. Against Word's
  eight saves, 0 lines differ in all eight, the second base with docx4j's three normalisations. The port needs only
  the writing half.
- **The listing.** `revisions.tracked_changes` gives 18 on `-phase1` and 17 on `-d9`, as Word does, but 41 on
  `-phase2` (Word 29) and 15 on `manual-redline-2-phase2` (Word 10). The difference is D15's `w:ins` holding
  `w:del`: the listing gives an insertion with empty text plus a one-space deletion, where Word's pane folds the
  space into its neighbours. Office JS's grouping, which CR-010 follows, is unmeasured on this shape (Q10).
- **The text model** reads `-phase1` as docx4j's tests expect. The accepted view has all eleven lines that
  `theSeparatingSpace` names, none with two spaces, and the markup view has `M01 Rate: {--1.316 --}1.335 per cent.`.
  On the base, one pass over `extract`'s blocks (`source` to `address`) gives `body/1`, `body/19` and
  `body/11/textbox:0/0`. There is no public `addresses`.
- **The validator.** docx4j's four conversions load with nothing skipped, and `pkg.misplaced()` reports nothing. A
  `w:del` in a hyperlink, a control or a smart tag is placement-valid here, and so is a `w:ins` holding a `w:del`.
- **Style strike.** `pkg.get_property_resolver().get_effective_rpr(r_pr, p_pr)` reads the styles privately (CR-002
  section 6.1) and gives the probe's cases: M05's `StruckText` strike on, M06's own `w:strike w:val="0"` off, N07's
  `DoubleStruck` double strike on.
- **The tracker** (`tracking.py`), which exists only while the mode is on. `ChangeTracker.markup` writes the id,
  `pkg.author`'s name, `w:date` on the package zone's wall clock (`xml_date`) and `w16du:dateUtc` (`xml_date_utc`,
  after `declare_ignorable`); "now" is to the minute. Ids come from a counter cached on `pkg._change_tracker`, one
  above the highest `CTMarkup` id in the parts already unmarshalled, `w:comment` ids excluded.
- **`TrackedChangeFilter`** has `within`, `within_bookmark` and `parts` (a part or its name). Its `range()` walks
  `package_stories(package)`, which skips a notes, comments, header or footer part whose bytes hold no revision
  markup. A bookmark found only there gives an empty range: right for resolving, wrong for this CR.
- **`DocPropsCustomPart`** is typed, and the field resolvers and the text model read it. Nothing sets a property or
  adds the part.
- **`description.py`** is `DocumentDescription` at `00614748e` with `d18693760`, which is the class at the tag.
- **Fixtures.** `tests/fixtures/revisions/` holds copies of `cr043/`, `editor/`, `core-ts/` and `python/`, and
  nothing from `cr068/`; `tests/revisions/conftest.py` reads the rest from the sibling checkout. CR-027 Phase B plans
  `tests/fixtures/manual-redline-dates/` for one file of `cr068/`.
- **Acceptance.** `scripts/acceptance.py` writes twelve artefacts; 8 is tracked changes, with the engine's accepted
  and rejected copies.

## 4. Design

### 4.1 The API

A new module, `docx4j_py/model/revisions/manual_redline.py`, is docx4j's class as a module (Q1). It is listed in
`codegen/clean.py`'s `KEEP`.

```python
from docx4j_py.model.revisions import manual_redline

report = manual_redline.to_tracked_changes(pkg)                       # strike only, every story but comments
report = manual_redline.to_tracked_changes(pkg, style_strike=True, underline_as_insertion=True,
                                           colour_as_insertion=["FF0000"], comments=True)
options = manual_redline.ManualRedlineOptions().author("Legal review").within_bookmark("Schedule2")
report = manual_redline.preview(pkg, options)                         # the same report; nothing changed
print(report.to_json())
```

- **`to_tracked_changes(pkg, options=None, **more)` and `preview(...)`.** The keywords apply over `options`, as
  `extract`'s do. A `pkg` of None raises `ValueError`.
- **`ManualRedlineOptions`** has `TrackedChangeFilter`'s shape: a keyword constructor and docx4j's chainable setters
  in snake case. The options:
  - `author` (`DEFAULT_AUTHOR = "manual redline"`, never empty);
  - `date`, an aware `datetime`, a naive one read as UTC as `tracked_change_date` is; None writes no date, and the
    default is the operation's time;
  - `separator_space` (True), `style_strike`, `underline_as_insertion`, `comments`, `audit_property`;
  - `colour_as_insertion(*hex)` (`RRGGBB`, upper-cased, else `ValueError`);
  - `within(control)` (a `ContentControl` view or a `w:sdt`), `within_bookmark(name)`, and `parts(*parts)` (a part
    or its name).

  `str(options)` is docx4j's `toString()` text.
- **No verb** on `pkg`, `Body` or the python-docx facade, as in docx4j and core-ts.

### 4.2 The walk

The module ports `ManualRedline.Conversion` and its `Walk`, with section 2's behaviour (Q2). It is built on the
engine's own helpers:
- `lister.package_stories(pkg, only_with_revisions=False)`, `lister.box_refs`, `child.mce_branch`;
- `trees`' lists and readers, with `trees.fit` and `splice` so that a wrapper takes the class its holder keeps
  (CR-020 refuses any other);
- `resolver.delete_text`, and `tracking.copy_r_pr` for a space's new run.

Where this package's representation differs:
- **On/off, underline and colour** are read through `model/values.py` (`is_on`, `underline`, `colour`), as CLAUDE.md
  requires. So `colour_as_insertion("FF0000")` also takes a `w:color` of `red` or `#FF0000`, which Word draws as
  `FF0000` and which docx4j's text comparison misses (Q2).
- **A kept `mc:AlternateContent`** at block, row, cell or paragraph level converts in every branch: Word's branch is
  reported, and the others are twins, as in docx4j. That holds even though the content API refuses an edit inside a
  kept branch (`address.read_only`, CR-012; Q8). Under `mc_load_mode="word"`, only Word's branch is there.
- **A text box below a wildcard node** converts through a typed island (`islands.Island`), as `accept_all`'s sweep
  does. No corpus document has one.

### 4.3 Ids, author, date

- **Ids** (Q4). The conversion first finds the highest id, as docx4j does. It scans every `CTMarkup` in every story
  part (`w:comment` included), plus `CommentRangeStart`, `CommentRangeEnd` and `RCommentReference`. That is a
  superset of the tracker's scan. Afterwards, a tracker cached on the package has its counter raised past the last
  id written (a small new `ChangeTracker` method), so a later tracked edit cannot reuse one.
- **Author:** the options', not `pkg.author`.
- **Date** (Q3). By default the operation's time to the minute (`now_to_the_minute()`): the portfolio's point 5,
  this package's tracker and core-ts's D2, where docx4j stamps to the second. A caller's `date` keeps its seconds.
  The tracker's own functions write it, so the two writers cannot drift: `xml_date(instant, package_zone(pkg))`,
  `xml_date_utc(instant)`, and `declare_ignorable(root, "w16du")` on each part with a dated change. `date=None`
  writes neither.

### 4.4 Reading privately, writing what changes (Q5)

The walk reads every story part through `read_contents()`. Before its first write into a part it takes the part's
`contents`, which adopts that very tree (CR-002 section 16), so the walk's objects are the live ones. The conversion
checks that identity, and raises rather than write into a tree that would not be saved. A part nothing is written
into stays unmarshalled and is saved byte for byte: a header with nothing struck, or the comments without the
opt-in. `preview` unmarshals nothing. docx4j unmarshals every story part; the behaviour is the same, and this is the
engine's fidelity rule applied. The new caller joins the `read_contents` callers that
`tests/openpackaging/test_private_readers.py` lists.

### 4.5 The report, and `text.addresses`

- **`ManualRedlineReport`** has snake-case attributes: `preview`, `changes`, `left`, `deletions`, `insertions`,
  `runs`, `strike_and_underline`, `by_kind`, `by_part` and `left_by_reason`. `Change` and `Left` are slotted
  dataclasses; `Kind`, `Type` and `Reason` are `enum.StrEnum`s with docx4j's labels, as `revisions.Kind` is.
  `to_json()` is docx4j's JSON exactly, its keys, order and escaping (Q6); `to_dict()` is its `json.loads`, and
  `str()` the JSON.
- **`docx4j_py.model.text.addresses(pkg, options=None, **more)`**, public (Q7), is a read-only identity mapping
  from each block's `source` (the first block made from it, cells and text boxes included) to its address. It holds
  the objects, so that no `id()` is reused while it is alive, as `TrackedChangeFilter.range` does. The conversion
  reads it once, before writing, under docx4j's options (`notes`, `comments`, `include="headers, footers"`). A
  paragraph with no entry takes its outer paragraph's address. A failed read leaves the addresses None and the
  conversion goes on, as in docx4j.

### 4.6 Engine additions

- **`TrackedChangeFilter.range(package, every_story=False)`.** With `every_story=True` it walks
  `package_stories(package, only_with_revisions=False)`, as docx4j's walks every story. The conversion asks for it;
  the listing and the resolvers keep today's path.
- **Custom properties** (Q9), for `audit_property`.
  - `OpcPackage.add_doc_props_custom_part()` (docx4j's `addDocPropsCustomPart`) makes `/docProps/custom.xml`, its
    package relationship and its content type, and does nothing where the part already exists.
  - `DocPropsCustomPart.set_property(name, value)` (docx4j's `setProperty(String, String)`) writes or replaces a
    `vt:lpwstr` property with docx4j's fmtid `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}` and the next `pid` from 2.
    `get_property(name)` reads one back.
  - The property is `docx4j:manualRedline`, its value docx4j's text: the instant with a `Z` (or `no date`), `; `,
    and `str(options)`.

### 4.7 Not here

- **Out of scope:**
  - the MCP server (CR-006, `../docx4j-mcp-py`), which takes the verb once built;
  - docx4j-core-ts (its CR-019), and the python-docx facade;
  - any change to a view, an extraction default or a reader (the user's ruling in docx4j);
  - as in docx4j: detecting a redline, deleted-row and deleted-cell marks, moves, `w:rPrChange`, and turning
    tracking on. M04's struck `Gizmo` row stays after Accept All with empty cells, as in Word.
- **Deferred:**
  - the text model's `text.formatting_join` finding, which would point a reviewer here (docx4j CR-069 phase 4, in
    the inspector, not ported; CR-027 section 4.5);
  - Word on the web, on the Mac and with classic comments, unmeasured in docx4j;
  - the listing's D15 shape (Q10).

## 5. Tests and oracles

### 5.1 The oracle copied (Phase A)

`docx4j-core-tests/src/test/resources/revisions/cr068/` at `docx4j-17.3.5` is copied unedited into
`tests/fixtures/revisions/cr068/`, beside `cr043/`, with a `SHA1SUMS` and a `tests/README.md` row naming the tag. It
is 54 files, about 560 KB: every file but the four screenshots, which the row names as left in docx4j.
`manual-redline-dates.docx` is among them. CR-027 Phase B would read it there rather than in
`tests/fixtures/manual-redline-dates/`, so it is copied once, by whichever CR lands first (Q11); deciding Q11
includes amending CR-027 section 5.

### 5.2 docx4j's tests ported (Phases B and C)

`tests/revisions/test_manual_redline.py` ports the 15 and `test_manual_redline_opt_ins.py` the 12. They run on the
bases with `pkg.date_time_zone = "Australia/Sydney"` and `date=2026-10-31T00:00:00Z`; that date is on the minute, so
Q3 moves nothing.
- **Word's saves.** Each base is converted, accepted and rejected by `revisions.accept_all` and `reject_all`, saved
  and loaded, and must describe (`description.describe`) as Word's save does: the default as `word-phase1-*`;
  `separator_space=False` as `word-d9-*`, through docx4j's `asD9` relabelling; the opt-ins as `word-phase2-*` and
  `word-phase2b-*`, with the three normalisations.
- **The files Word was given** (`thePhase1FileIsTheConversion`, `thePhase2FilesAreTheConversion`). This
  conversion's story parts must match `-phase1`, `-phase2` and `manual-redline-2-phase2` under `scripts/canon.py`'s
  comparison, `boolean-spelling` tolerated. Any other category is investigated before it is accepted.
- **The counts.**
  - The default (`theReport`): 18 deletions, 17 runs, `byKind` `{strike=15, separator_space=7, paragraph_mark=1,
    double_strike=1}`, left `{style_strike=1, in_tracked_change=1, field=2, comment=1}`.
  - Every opt-in but the range and the audit (`theFirstBase`): 26 deletions, 15 insertions, 28 runs.
  - The second base: `strike_and_underline` 1, and nothing left.
- **The rest of docx4j's cases:**
  - the three views after, idempotence, and `preview` equal to the conversion's report but for `"preview"`;
  - ids unique and above 100, and each `w:del` in a paragraph, a hyperlink, an inline control or a smart tag (17);
  - the dates (Sydney, UTC, none) and the author; `w16du` declared on a document with no `mc:Ignorable`;
  - the separating-space, left-alone and paragraph-mark cases;
  - N07's own `w:dstrike w:val="0"`, N02's kept underline, D15's shape, and D14 in the comment;
  - colours, a bookmark's range, a control's range (M17, `body/19`) and `parts`;
  - the audit property, and none from a preview.

### 5.3 This package's own

- **Byte for byte** (4.4). After a preview, or a conversion of a document with nothing struck, a save gives the
  source's bytes; a default conversion leaves the comments part unmarshalled.
- **Ids** (4.3). A tracked edit after a conversion takes a fresh id, whether its tracker was made before or after.
- **Readers.** "Now" lands on the minute, and a given date keeps its seconds. `red` and `#FF0000` are taken by
  `FF0000`. `text.addresses` gives `body/1`, `body/19`, `body/11/textbox:0/0`, `footnote:1/0` and a header's
  address, and nothing for an untaken branch.
- **The engine additions.** `range(every_story=True)` finds a bookmark in a header with no revision markup. The
  custom property round-trips, and `DOCPROPERTY` reads it.
- **Placement.** Every save in the suite runs `save(check=True)` (`tests/conftest.py`), so a misplaced wrapper fails.

### 5.4 The Word check (Phase D)

This CR writes revision markup and can create a part with a content type, so CLAUDE.md requires a Word check.
`scripts/acceptance.py` gains artefact 13. It is built from the copied probe base with the fixed date in Sydney's
zone, so that `w:date` reads 11:00 on any machine:
- `13-manual-redline.docx`: the default conversion;
- `13-manual-redline-optins.docx`: style strike, underline, `FF0000`, comments and `audit_property`;
- `13-manual-redline-accepted.docx` and `-rejected.docx`: the engine's Accept All and Reject All of the first, each
  asserted to list nothing.

The user opens each in Word 365 under All Markup. What Word should show, from docx4j's runs on the same conversions:
1. **Opening.** All four open without a repair prompt.
2. **The default's 18 revisions.** 17 are by `manual redline` at 31/10/2026 11:00, in red strikethrough with change
   bars, and M07's insertion is by `CR-068 base` in a second author's colour. M05 (a style's strike), M08 and M08b
   (fields) and M16 (the comment) still show black strikethrough, and M12 keeps its own red.
3. **Accept All of the default**, and `-accepted.docx`:
   - `M01 Rate: 1.335 per cent.`, `M09 See example.com/new-terms for the terms.`, `M13 Fee: EUR 450 a month.`,
     `M14 kept words, end.`, `M15 Before after.`, `M17 Party: Acme Holdings Pty Ltd.`;
   - the footnote `the rate is 1.335.`, the text box `the price is $35.00.`, the header `M10 header: Final version`;
   - M11's struck paragraph joined to the next, no two spaces anywhere, and the Gizmo row kept with empty cells.
4. **Reject All of the default**, and `-rejected.docx`: the base's text, every converted run plain and not struck,
   and M07's insertion gone.
5. **The opt-ins.** 29 revisions: 23 in the main story, 2 in the footnote, 1 in the comment, 2 in the text box, 1
   in the header.
   - Accept All adds `M05 then plain words.` and `M12 The party shall promptly notify the other.`; Reject All gives
     `M01 Rate: 1.316 per cent.` and `M12 The party shall notify the other.`.
   - M16's card shows the "Missing content" notice, and the Revisions pane `withdrawn remark,` struck. That is
     predicted from `comment-cards.docx`'s C2: docx4j did not look at this file's card.
   - File > Info > Properties > Advanced Properties > Custom lists `docx4j:manualRedline`.

Differences go into `tests/README.md`'s checklist. A hang or a quiet refusal also goes into
`../docx4j-portfolio/docs/word-hangs.md`.

### 5.5 The gates

The suite; `scripts/roundtrip.py`, 50 of 50; `compare_acceptance.py`, differing in artefact 13 alone;
`codegen/generate.sh --check`; and `scripts/release_check.sh` before the next release.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | The oracle copied (5.1), with `SHA1SUMS` and its `tests/README.md` row; `text.addresses` (4.5) and its tests | ½ day |
| B | `manual_redline.py`'s default walk (4.2), ids, author and date (4.3), the private read (4.4), the report, `preview` and idempotence; `TrackedChangeFilter.range(every_story=)`; `test_manual_redline.py` (15) and 5.3's tests; held to `word-phase1-*` and `word-d9-*` | 1½ days |
| C | The opt-ins (2.2) and the custom-properties additions (4.6); `test_manual_redline_opt_ins.py` (12); held to `word-phase2-*` and `word-phase2b-*` | 1 day |
| D | Artefact 13 (5.4), `out/acceptance/` regenerated and compared; the user's Word check, recorded in `tests/README.md` | ½ day, and the user's check |
| E | The records: CHANGELOG's Added (`manual_redline`, `text.addresses`, the custom-properties methods); a `README.md` fence, which `tests/test_readme.py` runs; CLAUDE.md's status and file list; the index's row; Q10's note in CR-010; the portfolio's `python/CR-028.*` entries through the docx4j-portfolio session | ¼ day |

## 7. Risks

- **A strike that is not a redline**, such as "was $40" in a price list. The conversion is explicit; `preview`, the
  range and `parts` are the reviewer's tools. The author name keeps it reversible:
  `revisions.reject_all(pkg, TrackedChangeFilter(authors=["manual redline"]))` gives plain text back.
- **Adoption** (4.4). If `contents` re-parsed instead of adopting the tree, the writes would go into a tree that is
  never saved. The engine adopts whenever the load options are unchanged, and the identity check turns any miss into
  an error, not a silent loss.
- **Placement.** A `w:del` or `w:ins` in a hyperlink, a control or a smart tag is a scoped class here. `trees.fit`
  re-types it, and the suite's save check catches a miss. docx4j's conversions already pass `misplaced()` here.
- **The canonical comparison** (5.2) may show another category, such as `xml:space` placement. Each is investigated.
  Where docx4j's own output disagrees with Word, the finding goes to docx4j, as core-ts's section 7 says.
- **D15 in the listing** (section 3, Q10): more changes than Word's pane after an opt-in conversion.

## 8. Questions for the user

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The API: a module `model/revisions/manual_redline.py` with `to_tracked_changes` and `preview`; `ManualRedlineOptions` in `TrackedChangeFilter`'s shape (keywords, and docx4j's chainable setters in snake case); `**more` keywords as `extract` takes them; no verb on `pkg`, `Body` or the facade? | Yes. The call reads as docx4j's class does (`manual_redline.to_tracked_changes(pkg)`), and the options take their sibling's shape in `model/revisions`. core-ts's D1 makes the same choice in its own idiom. |
| Q2 | docx4j's rules as the code at the tag has them (section 2, with 2.5's wider reach), D1 to D16 as built, D14 included? The departures would be only those named here: Q3, Q5, Q6's key spelling, `parts` by name, and the colour read through `values.colour` (`red` and `#FF0000` counting as `FF0000`). | Yes. Word's three runs measured the rules, and this engine already resolves docx4j's conversions as Word does (section 3). The colour departure is CLAUDE.md's rule for colour readers, and Word's own reading (CR-018). |
| Q3 | The default stamp: to the minute (the portfolio's point 5, this package's tracker, core-ts's D2), or to the second (docx4j)? | The minute. Every stamp this package writes is to the minute, and Word's pane shows minutes. A caller's `date` keeps its seconds. The tests' date is on the minute, so no golden depends on the choice. |
| Q4 | Ids: docx4j's scan (every annotation id in every story part, comment ids and marks included), with a cached tracker's counter raised afterwards? | Yes. It is a superset of the tracker's space, so it only takes higher ids, and it gives docx4j's ids on any document. Without the raise, a tracked edit after a conversion could reuse an id. |
| Q5 | Read every story privately (`read_contents`), unmarshalling only the parts written into, so that `preview` and untouched parts stay byte for byte? | Yes. It is the engine's fidelity rule, and it changes no behaviour. docx4j has no private read. |
| Q6 | `to_json()` in docx4j's exact form (`strikeAndUnderline`, `byKind`, `byPart`), although `TrackedChangesReport.to_dict` here uses snake case? | Yes. The MCP vocabulary names `ManualRedlineReport.toJson` as the verb's result, so `docx4j-mcp-py` can return it unchanged. The attributes stay snake case. |
| Q7 | A public `docx4j_py.model.text.addresses(pkg, options=None, **more)`, an identity mapping from a block's source to its address? | Yes. It is docx4j's `Texts.addresses` (public since 17.3.5) and core-ts's D3, and it is one pass over `TextBlock.source`. |
| Q8 | Convert inside a kept `mc:AlternateContent`, at every level and in every branch, although the content API refuses edits inside a kept branch (`address.read_only`)? | Yes. That refusal guards a single edit at an address that a branch may not show. The conversion, like `accept_all`, treats the whole document, and docx4j converts every branch. |
| Q9 | `OpcPackage.add_doc_props_custom_part()` and `DocPropsCustomPart.set_property` / `get_property` as new engine API, with the property `docx4j:manualRedline` and docx4j's value text? | Yes. These are docx4j's names (core-ts's D5 adds the same), and the property then reads the same whichever port wrote it. `audit_property` is off by default. |
| Q10 | The listing of D15's `w:ins` holding `w:del`, as an empty insertion plus a one-space deletion where Word's pane folds the space into its neighbours: leave it, or change CR-010's grouping? | Leave it, and note it in CR-010. Office JS's grouping, which CR-010 follows, is unmeasured on this shape. A Script Lab read of `trackedChanges` on `13-manual-redline-optins.docx` at Phase D would settle it, if the user wants one. Accept All and Reject All are right either way. |
| Q11 | One copy of `revisions/cr068/` at the tag, without the screenshots, in `tests/fixtures/revisions/cr068/`, with CR-027 Phase B reading `manual-redline-dates.docx` there instead of from its own directory? | Yes. Each docx4j directory then has one copy here, beside `cr043/`. If CR-027 is to stay as written, copy `cr068/` without that file, and leave it to CR-027's directory. |

## 9. Effort

3¾ days, and the user's Word check: A ½, B 1½, C 1, D ½, E ¼.
