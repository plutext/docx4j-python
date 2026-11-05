# CR-027: the text model's WD 0.41 to 0.49 additions --- `profile: word-text`, revision and comment dates as the author's clock, embedded objects for a host, and the conformance suite re-pinned

**Status:** Proposed 2026-10-31, at the user's word ("fold it into CR-027 with profile and the CR-070 items"); **decided 2026-11-02**, every question as recommended (Q6 on 2026-11-01; section 8), and amended the same day (unreadable parts); **Phase A implemented 2026-11-02** (section 10: `1065287`, `5844c0b`, `c7d4c81`, `4ee1849`): the suite at the tag, `phase0/` 1,294 pass, 0 fail, 478 pending and `dpp/` 3 pass, 189 pending; **Phase B implemented 2026-11-03** (section 11: `a21e8cc`, `7f42fa1`): revision and comment dates in REQ-702's form, `TrackedChange.author_local_time` and `author_offset`, docx4j's `revisions/cr068/` copied at the tag; no suite row, agreement row or golden moved. **Amended 2026-11-05** (WD 0.84, section 4.1): `word-text` sets `list_labels: true`. **Phase C implemented 2026-11-05** (section 12: `d8f3faf`, `6f0d15a`): `TextObject` on every object run, the OLE 1.0 names (`model/text/ole.py`), `embedded_filter`, the JSON's `object` member; no suite row moved. **Phase D implemented 2026-11-05** (section 13: `1b6192c`, `6f0d15a`): `profile: word-text`, `dpp/` 9 pass, 0 fail, 183 pending, six of Word's own text saves checked. **Phase E done 2026-11-05** (section 14): the records. **Complete**; no Word check (section 5). It covers five
pieces of the text model (CR-011's `model/text/`) that the WD and docx4j have added since this repository's suite
copy was pinned at docx4j `198d015a3`:
- WD 0.41's `profile` (REQ-1002);
- WD 0.46 and 0.49's date form (REQ-702, REQ-706);
- docx4j CR-070's embedded objects (WD 0.47's REQ-512, REQ-522, REQ-513 and REQ-801);
- a part the text model cannot read, named (REQ-801, REQ-512; docx4j `4fdd80063`, added 2026-11-02);
- the suite re-pinned so that its rows show them.

Five phases (section 6). The questions are in section 8, each with a recommendation. Q6 decided 2026-11-01 (the user:
"CR-027 should copy docx4j's test suite from the 17.3.5 tag"). **Decided 2026-11-02** (the user: "CR-027's questions  Q1 to Q5 and Q7 as recommended"): Q1 to Q5 and Q7 as
recommended. Nothing built. **Reviewed 2026-11-02** by docx4j's session: its corrections are folded in below, and
no decision changes; Q5 is carried out differently (section 4.1). **Amended 2026-11-02** (the user: "yes", to adding
docx4j's change for a part that cannot be read): Phase A also names a part the text model cannot read (sections 2.4
and 4.6), ¼ day more.

*The number:* CR-026 is the latest CR here. CR-028 is the proposed number for the port of docx4j CR-068 (manual
redline to tracked changes), which the user has not yet decided.

**Depends on:**
- **docx4j CR-046 section 9, `0cf4fb38a`** (WD 0.41 REQ-1002, 2026-10-29; rebased from `bc842ac66`): `TextOptions.profile`,
  `TextOptionsProfileTest`.
- **docx4j `fc71f12e6`** (WD 0.46 REQ-702 and REQ-706, 2026-10-31; the date review's section 3.7 (c), Jason: "rule (c),
  yes, build it into 17.3.5"): `RevisionDates`, `TrackedChange.getAuthorLocalTime()` and `getAuthorOffset()`,
  `TextRevision.getDate()` and `TextComment.getDate()` in REQ-702's form; `RevisionDatesTest`, `MarkupDatesTest`, and
  `TextsTest`'s TC9 date form.
  WD 0.49 (docx4j-portfolio `a993e46`) states docx4j's three precision rules (section 2.2).
- **docx4j CR-070, `1560ab379`** (WD 0.47, 2026-10-31; D1 to D5 accepted, its section 7 the rules as built):
  `TextObject`, `EmbeddedFilter`, `altchunk.not_expanded`'s part name, a nested warning naming the outer part, and
  `expected-findings.tsv` re-keyed; `EmbeddedObjectsTest` (9). docx4j-12 handed it on the same day.
- **docx4j 17.3.5** (Maven Central 2026-11-01; tag `docx4j-17.3.5` on origin at `64fae4cd0`): the release holding all
  three, and the suite this CR copies (section 4.4, Q6).
- **docx4j `4fdd80063`** (`VERSION_17_3_6`, unpushed, after the tag; docx4j CR-065 section 9, 2026-11-02; Jason:
  "make those three throw a Docx4JException naming the part and its cause"): a main part the text model cannot read
  raises, naming the part and why; in a nested reading, and for every other story part, `content.not_read` instead.
  `UnreadablePartTest`'s four text-model tests (section 2.4). No suite row changes, so the copy stays at the tag.
- **The WD** (`../docx4j-portfolio/docs/text_extraction_specification_wd.md`): 0.34 to 0.42 committed at portfolio
  `e49964f`, 0.46 and 0.47 at `fb26239`, 0.48 at `fb3cfcb`, 0.49 at `a993e46`; this CR follows that text.
- **This repository:**
  - CR-011's text model and its conformance harness (`tests/text/conformance.py`, the suite copy in
    `tests/fixtures/text-extraction/` at `198d015a3`);
  - CR-010's `TrackedChange` (`model/revisions/change.py`);
  - CR-025's `pkg.date_time_zone`, which `TrackedChange.date` reads `w:date` in and which this CR leaves alone;
  - CR-003 Phase G's comment parts (`openpackaging/parts/wml/comments.py`, `dates_utc`).
- **Counterparts:** docx4j as above. docx4j-core-ts ports the same pieces as its own entries (the portfolio names
  core-ts/CR-019 and CR-012.22).

## 1. Why

Each piece is a WD requirement this package does not yet meet, measured 2026-10-31 at `5cbc01b`:

- **`profile`** is refused: `TextOptions.of(profile="word-text")` raises `TextOptionError` ("Unknown text option
  'profile'"). The preset's other 24 values at `0cf4fb38a` are already options here; the tag's preset has 26
  (section 2.1).
- **The dates.** `extractor._date_text` writes `w:date` as stored and never reads `dateUtc`. On docx4j's four measured
  cases (`revisions/cr068/manual-redline-dates.docx`, under `view: markup, markup_detail: author_date`), every date
  differs from REQ-702:

  | Case | Stored | Here | REQ-702 |
  |---|---|---|---|
  | D1 | `w:date` 00:00Z alone | `2026-10-31T00:00:00Z` | `2026-10-31T00:00:00` |
  | D2 | `dateUtc` 00:00Z alone | (nothing) | `2026-10-31T00:00:00Z` |
  | D3 | `w:date` 00:00Z, `dateUtc` 00:00Z | `2026-10-31T00:00:00Z` | `2026-10-31T00:00:00+00:00` |
  | D4 | `w:date` 11:00Z, `dateUtc` 00:00Z | `2026-10-31T11:00:00Z` | `2026-10-31T11:00:00+11:00` |

  A mark with neither date writes `author, : ` in plain text and Markdown, where REQ-702 writes `author: `.
- **The embedded objects.**
  - An object run carries its ProgID alone (`TextRun.prog_id`): no kind, relationship id, part, content type or OLE
    1.0 package names.
  - There is no per-object filter.
  - `altchunk.not_expanded`'s selector is the chunk's content type.
  - A warning raised one level down names the inner package's part.
- **The suite.** Its copy is docx4j's at `198d015a3` (2026-10-21). Since then four docx4j commits have changed the
  tables:
  - `2853b7280` and `4a37702eb`: CR-050 phases 3 and 4, the text a picture draws and a capability column;
  - `338beec18`: CR-066 phase 4, `value.out_of_range` in the inspection, cases 75 to 82;
  - `1560ab379`: CR-070's re-keyed rows.

  Net, the tables gained 34 token rows (285 to 319), 56 findings rows, 56 inspection rows and 15 case statements,
  with new columns in the inspection tables; `4a37702eb` deleted `expected-inspection-html-status.tsv` (section 4.4),
  and `2853b7280` and `338beec18` changed `extraction-probe-inspect.docx` itself. Until the copy moves, none of
  CR-070's rows can be checked here. A fifth commit, `fe77bea88` (a control bound to a cover-page property shows the
  property), added `bind-props.docx` at the folder root, in neither CR-011's list nor `dpp/`. What it fixed already
  holds here: `bind-props.docx` reads `PR4_COVER_ABSTRACT`, checked by hand, not by a test here (docx4j's is
  `BoundPropertiesTest`).

  By the 17.3.5 tag (`64fae4cd0`) four more commits have changed them, none of the tables above:
  - `f2432c3de` and `a8e57ddc5`: CR-066 phase 4's Word evidence (`phase0/evidence/word-open-results.md`), which is not
    copied;
  - `993e385e7` and `7326296b1`: CR-069's suite, `dpp/`, beside `phase0/`, and its inspector row in `INSPECTION.md`.

  `dpp/` is the WD's sixth fixture (WD 0.50). It holds:
  - Word for Microsoft 365's re-save of the probe, with ten cases;
  - each case alone (`bisect/`) and Word's caps measurements (`caps/`), with Word's own plain-text save of each;
  - its tables: 34 token rows, 8 findings and 2 goldens;
  - `word-saves.tsv`, 22 rows. Each compares a fixture's plain text under `profile: word-text` with Word's save,
    decoded as the row says (windows-1252, or UTF-8 with a byte order mark).

  Nearly all its rows state CR-069's requirements, which are not built here: REQ-521, REQ-607, REQ-710 and the
  inspector's `text.formatting_join` (section 9.2).

## 2. What docx4j built

### 2.1 `profile` (REQ-1002; docx4j CR-046 section 9)

`profile` names a preset of option values, applied before the caller's own. An option the caller sets by name is
recorded, and the preset fills only the options not recorded, so `profile` then `fields: update` and `fields:
update` then `profile` both give `update`.
- By name, `profile: word-text` works in a map in either order: matched case-insensitively, with a hyphen or an
  underscore.
- `toMap` carries `profile` last (null where there is none), so options round-trip by name.
- An unknown profile name is refused, naming REQ-1002.

`word-text` is Word's Save As plain text as the options reach it, at `0cf4fb38a` in 24 values:
- `view: accepted`, `include: [headers, footers, comments]`, `comments: true`, `hidden: true`;
- `text_boxes: false`, `notes: true`, `note_references: true`, `list_labels: true` (amended 2026-11-05, WD 0.84: `false` at `0cf4fb38a`);
- `fields: cached`, `field_codes: false`, `toc: true`, `placeholders: true`, `bindings: bound`, `altchunks: true`;
- `symbols: asis`, `tables: tsv`, `page_markers: false`, `layout: none`;
- `alt_text`, `metadata`, `charts`, `smartart` and `embedded` false;
- `math: latex`, not the WD's `unicodemath`, which docx4j does not implement.

At the tag `TextOptions.WORD_TEXT_VALUES` has 26: CR-069 added `drop_caps: asis` (`15c66769f`) and
`small_caps: upper` (`75b9bdb54`), and WD REQ-1002 lists both, though docx4j's CR-046 section 9 still says 24.

The WD's NOTE lists what the preset does not reproduce:
- the notes' and comments' forms, CRLF and CR cells, folded hyphens, and Word's linear math;
- three accidents left out on purpose: the partial update Word's open runs where a DATE, TIME or SAVEDATE field is
  present, the SYMBOL omission, and the stale PAGE number.

An implementation documents these.

### 2.2 Dates as the author's clock (REQ-702, REQ-706; `fc71f12e6`, WD 0.49)

Word writes two attributes:
- `w:date`, the author's wall clock, labelled `Z` although it is not UTC;
- `w16du:dateUtc` (`w16cex:dateUtc` for a comment), the instant.

Word's Reviewing Pane shows `w:date` and ignores `dateUtc`, and a mark with `dateUtc` alone shows no date (measured
2026-10-31, docx4j CR-068 phase 0). `RevisionDates` gives three readings:

- **`localTime(date)`**: `w:date`'s date and time as stored. Any designator (a `Z`, an offset) is not read, a missing
  time is 00:00:00 (a date alone is midnight), and fractions of a second are dropped.
- **`offset(date, dateUtc)`**: `w:date` less `dateUtc`, where both are present. It counts only when the difference is
  a whole number of quarter hours from -12:00 to +14:00, as a clock's offset is; else there is none (a pair written
  inconsistently gives no offset rather than a wrong one). `dateUtc` without a designator is UTC, and one with an
  offset has it applied; its fractions of a second are dropped too, since `utc` reads it through `localTime`.
- **`text(date, dateUtc)`**, REQ-702's form, in four cases:
  - `w:date` with an offset known: the local time with the offset after it (`2026-10-31T11:00:00+11:00`; `+00:00`
    for UTC, never `Z`);
  - `w:date` alone: the local time in ISO 8601's extended form with no designator (`2026-10-31T11:00:00`);
  - `dateUtc` alone: written with `Z`;
  - neither: none.

  The `Z` Word stores on `w:date` is never written, and no time is converted to another zone.

Where they are used:
- `TextRevision.getDate()` and `TextComment.getDate()` return the form.
- A comment's `dateUtc` is found through its last paragraph's `w14:paraId`, which `commentsIds` maps to a durable id
  and `commentsExtensible` maps to the date.
- Under `markup_detail: author_date` a mark with no date writes `author: ` alone, in plain text and Markdown alike.
- `TrackedChange` gains `getAuthorLocalTime()` and `getAuthorOffset()` beside `getDate()`, the instant, which is
  unchanged. With both known, the local time at the offset is the instant.

`RevisionDatesTest`'s cases are the rule's oracle:
- the measured four;
- offsets of -05:00, +05:30, +05:45, +14:00 and -12:00;
- none for seven minutes, for over fourteen hours and for a day behind;
- the other forms: no designator, an offset on `w:date`, a date alone, a fraction, and a `dateUtc` with an offset.

### 2.3 Embedded objects for a host (docx4j CR-070, REQ-512, REQ-522, REQ-513, REQ-801)

- **(a) The object's description.** Every object run gets a `TextObject`, whatever `embedded` says. Its members, as
  the WD names them for the block model:
  - `kind`: `embed` or `link` from `o:OLEObject/@Type`, `embed` where it is absent; `control` for a `w:control`, and
    for an `o:OLEObject` whose part's content type is an ActiveX control's (`application/vnd.ms-office.activeX+xml`
    or `application/vnd.ms-office.activeX`);
  - `relationshipId`, the `r:id`;
  - `progId`, which a `w:control` does not have;
  - `part`, the embedded or control part's name: absent for a linked object, even one whose relationship is
    internal, and for an unresolved relationship;
  - `contentType`, that part's;
  - for an OLE 1.0 package (`Extractor.object`: a ProgID beginning `package`, in any case, whose part is a
    `BinaryPart` and not an `EmbeddedPackagePart`): `label`, `fileName` and `command`, from its `Ole10Native` stream's
    plain fields (not the UTF-16 `label2`, `fileName2` and `command2`). An empty value is absent.

  These are the names Tika 3.3.3 gives such a package, so that a host's filter on `*.txt` meets `notes.txt`, not
  `oleObject1.bin`. A linked object is never read and never offered to the filter.
- **(b) The filter.** Under `embedded: true` a host's filter (`EmbeddedFilter`, `boolean read(TextObject, String
  address)`) is asked once per object the extractor would read, before it is read. The objects concerned are an
  embedded OOXML package and an OLE 1.0 package, never one level down. The address is REQ-519's object segment,
  `body/3/object:0`.
  - Its index counts every object the extractor would read, declined or not, so a decline moves no other address.
  - A declined object contributes nothing and keeps its run and description.
  - A filter that raises is taken as declining, and raises `content.not_read` with the object's part name at the
    anchoring paragraph's address.
  - With no filter, every object is read.
  - An inspection does not see the filter.
- **(c) The warnings.**
  - `altchunk.not_expanded`'s selector is the chunk's part name; the warning still carries the content type. The
    selector is the branch, else the part name, else the content type.
  - A warning raised inside an embedded package or a chunk that names a part names the outer package's or chunk's
    part, the one the host package has. That is REQ-801's `content.not_read` rule, applied to a nested chunk's
    `altchunk.not_expanded` and a nested `image.not_described`. Word names every HTML chunk `afchunk.htm`, so the
    inner name could find another part.
  - So `expected-findings.tsv`'s twelve `altchunk.not_expanded` rows are keyed by part name (`/word/afchunk.htm`,
    `/word/chunk43.txt` and so on to `/word/chunk63.htm`).
  - Case 36's row of count 2 splits into `/word/afchunk.htm` (`html reader absent`, 1) and `/word/chunk47.docx`
    (`default`, 1).

### 2.4 A part that cannot be read (`4fdd80063`; REQ-801, REQ-512)

docx4j's load returns for a part whose XML cannot be read (CR-065 D1): `getJaxbElement()` is null, and an untouched
save keeps its bytes. The text model now says so, where before 17.3.6 it gave an empty result for the main part,
threw a `NullPointerException` for the notes, and passed over the rest silently.
- **`JaxbXmlPart.getReadableContents(purpose)`**: the content, or a `Docx4JException` "The content of
  /word/document.xml could not be read, so it cannot be read as text: <reason>". The reason is the deepest cause's
  message on one line, its whitespace collapsed, or the cause's class name where it has none.
- **The main part.** `Extractor.extract` (`Texts.of(pkg)`), and `Texts.of(part)` for the main part, a header or a
  footer, throw it as a `Docx4JRuntimeException`, the part's own exception its cause.
- **Nested** (an embedded or chunked document): the main part gives `content.not_read` at the nested reading's root
  (its trailing `/` dropped), selector the part name, and the nested result is otherwise empty (REQ-512).
- **Every other story part** (`Extractor.storyContent`): footnotes, endnotes, comments and the glossary (no address),
  and a header or footer (its prefix as address), that cannot be read contribute nothing and raise
  `content.not_read` with the part name (REQ-801). The body is still read.
- **Tests:** `UnreadablePartTest`'s `theTextModelSaysWhichPartCouldNotBeReadAndWhy`,
  `theTextModelOfAHeaderSaysWhichPartCouldNotBeRead`, `anEmbeddedDocumentThatCannotBeReadIsContentNotRead` and
  `unreadableHeadersAndNotesAreContentNotRead`. Its others (CR-065's saves, `toHTML`) have no counterpart here.

## 3. What exists here (measured 2026-10-31 at `5cbc01b`)

- **`TextOptions`** (`model/text/options.py`) is a slotted class whose options are plain attributes. They are set by
  the constructor's keywords, by `set(name, value)` and by `of(...)`, and a caller can also assign one directly
  (`opts.hidden = True`). Nothing records which options a caller set. `check_implemented` refuses `math: unicodemath`
  as docx4j does, and `to_dict` writes every option.
- **The dates.**
  - `extractor._date_text` is the lexical form of `w:date`; `_revision` and the comment reader (around line 2100)
    use it.
  - `plain._Writer._detail` and `markdown._detail` write `f"{author or ''}, {date or ''}: "`.
  - `TrackedChange.date` (CR-010, CR-025) is the instant: `dateUtc` where present, else `w:date` read in the
    package's zone, a date alone as local midnight.
  - The content API's `Comment.creation_date` reads `dates_utc(parts)`, the same paraId-to-durable-id-to-date chain
    REQ-702 asks for, through the last paragraph with an entry.
- **The objects.**
  - `extractor._object` makes an `OBJECT` run with `prog_id` alone, and none for a `w:control`.
  - Under `embedded: true`, an `EmbeddedPackagePart` or an OLE 1.0 package's binary part is queued (`_EmbeddedObject`,
    with its address).
  - `embedded.package_blocks` reads a WordprocessingML package one level down; a workbook, a presentation and an OLE
    1.0 package are deferred (CR-011 Q7), giving nothing and raising `content.not_read`.
  - No module reads an OLE compound file, and nothing here parses `Ole10Native`.
  - `TextWarning` has `branch`, `content_type` and `part_name`, and its `selector` is the branch, else the content
    type, else the part name.
- **A part that cannot be read** (measured 2026-11-02 at `8f9e618`). `trees._read` (`model/text/trees.py`) returns
  None for a part that is absent and for one that cannot be read alike.
  - An unreadable main part gives an empty result with no warning.
  - Unreadable footnotes, endnotes, comments, headers and footers are left out with none (`PackageTrees.__init__`),
    and so is the glossary (`extractor._glossary`). That breaks REQ-801 as `extractor.py`'s own docstring states it.
  - `extract(header_part)` over an unreadable header raises `TypeError("not a story part ...")`, which names the
    wrong cause.
- **The harness** (`tests/text/conformance.py`, `test_conformance.py`) is docx4j's `ExtractionConformance` at
  `198d015a3`, over the copy in `tests/fixtures/text-extraction/`. CR-011's status records 1,291 pass, 0 fail and 293
  pending: docx4j's 19 that need an HTML reader, the inspector's 214 (CR-011 section 3.7) and Q7's 60.

## 4. Design

### 4.1 `profile` (Phase D)

- `TextOptions` gains a `profile` slot and a private `_explicit` set of the option names a caller has set.
- Every way a caller sets an option records it in `_explicit`: the constructor's keywords, `set`, `of`, and direct
  assignment to an option attribute. Direct assignment is caught by a `__setattr__` over the option names, one
  comparison per assignment.
- The defaults and `copy()` do not record. `__init__` assigns every default and `copy()` calls `setattr` over every
  slot (`options.py`), so through the recording `__setattr__` every option would be marked explicit and the preset
  would fill nothing; `of(opts, profile=...)` copies first and would fail the same way. So `__init__` assigns its
  defaults past the recording, and `copy()` carries `_explicit` and `profile` as they stand. This changes how Q5 is
  carried out, not Q5 (docx4j's review, 2026-11-02).
- The preset's own values go through a private path that does not record them, so `profile` then `fields="update"`
  and `fields="update"` then `profile` both give `update`, as docx4j's do.
- `set("profile", "word-text")`, `TextOptions(profile=...)` and a mapping with `profile` in any position accept the
  name case-insensitively, with a hyphen or an underscore.
- An unknown name raises `TextOptionError` naming REQ-1002 (reason `invalid`).
- `to_dict` carries `profile` last, None where there is none.
- `word-text`'s values are the 24 of `0cf4fb38a` (section 2.1), with `math: latex`, since `check_implemented` refuses
  `unicodemath` here too. The tag's other two, `drop_caps: asis` and `small_caps: upper`, come with their options in
  the CR-069 port (CR-029); until then `word-text` keeps `w:smallCaps` text as stored. No conformance row breaks:
  the `word-saves.tsv` rows that need them cite REQ-607 or REQ-521, which are not advertised here, and pend (4.4).
- The NOTE's list of what the preset does not reproduce goes on `TextOptions`' docstring, with `math` among it
  (Q5).
- **Amended 2026-11-05 (WD 0.84).** docx4j's preset sets `list_labels: true` since `87c421731` (the user's ruling of
  2026-11-04, after the tag), so both built presets hold REQ-1002's value and this one follows. The oracle for this one value is
  docx4j's `TextOptions` at that commit, not the tag. The NOTE's list gains what docx4j's Javadoc names: the one space Word's text
  save writes after every label, where the text model writes the level's `w:suff`. The dpp suite's 22 Word saves hold no list,
  so no row moves.

### 4.2 The dates (Phase B)

- **`model/revisions/dates.py`** (new, in `codegen/clean.py`'s `KEEP` beside `model/dates.py`) is the port of
  `RevisionDates`: `local_time(date) -> datetime | None` (naive), `offset(date, date_utc) -> timezone | None` and
  `text(date, date_utc) -> str | None`, over the model's `XmlDateTime` values and their lexical forms.
- **`TrackedChange`** gains `author_local_time` and `author_offset` (docx4j's two getters) beside `date`, which is
  unchanged (Q1). With both known, `author_local_time.replace(tzinfo=author_offset)` is `date`.
- **`TextRevision.date` and `TextComment.date`** take REQ-702's form (Q2).
  - `_revision` reads `date` and `date_utc` from the mark.
  - The comment reader takes `date_utc` from `dates_utc(parts)` by `_last_para_id`, the last paragraph carrying a
    `w14:paraId`, which already descends into the `mc` branch the walk reads, as docx4j's `lastParaId` takes the
    branch its TEXT selection chooses. That is not `creation_date`'s last paragraph with an entry.
  - Both dates are read to the second. `RevisionDates` drops `dateUtc`'s fractions as well as `w:date`'s, but
    `dates_utc` keeps microseconds (`comments.py`'s `_utc`), so the port truncates them, or a `dateUtc` of
    `...:00.5Z` would give no offset.
  - The content API's `Comment.creation_date` and `TrackedChange.date` are untouched: they are the instant, and
    CR-025's reading of them stands.
- **`plain` and `markdown`** write `author: ` alone where the date is None, under `author_date`.

### 4.3 Embedded objects (Phase C, and the warnings in Phase A)

- **`TextObject`** (`model/text/result.py`): a slotted dataclass on `TextRun.object` for every `OBJECT` run, None on
  any other. It has:
  - `kind` (`"embed"`, `"link"`, `"control"`), `relationship_id`, `prog_id`;
  - `part`, the part itself, not in the JSON;
  - `part_name`, `content_type`, `label`, `file_name`, `command`;
  - `source`, the `w:object` or a lone `w:control`, not in the JSON, as `TextImage.source` is.
- `TextRun.to_dict` writes an `"object"` member holding the WD's members in this JSON's spelling (`kind`,
  `relationship_id`, `part` = the part name, `content_type`, `label`, `file_name`, `command`), each left out where
  None.
- `extractor._object` makes the description as docx4j's `Extractor.object` and `control` do:
  - the kind from `o:OLEObject/@Type` (`embed` where it is absent), `control` for a `w:control` or an ActiveX part's
    content type, and an OLE 1.0 package as section 2.3 (a) has it;
  - the part through `trees.related`, and none for a link.
- **The OLE 1.0 names** (Q3) come from a small read-only reader of an OLE compound file:
  - `model/text/ole.py` (new): the header, the FAT and mini FAT, the directory, and one named stream read whole,
    `\x01Ole10Native`, its name matched case-insensitively (`getEntryCaseInsensitive`);
  - then `Ole10Native`'s plain `label`, `fileName` and `command`, read as docx4j's copy of POI's `Ole10Native`
    (`org.docx4j.org.apache.poi.poifs.filesystem`) reads them:
    - names only where `flags1` is 2 and the label's first byte is not a control (0x00 to 0x1F or 0x7F: Java tests
      it as a signed byte); otherwise the record is compact or unparsed and has none;
    - `label` and `fileName` null-terminated and UTF-8 (`readAsciiZ`, `getFromCompressedUTF8`; a malformed byte is
      U+FFFD, as Java decodes it);
    - `command` after its length, ISO-8859-1, its last byte dropped (`readAsciiLen`);
    - a string too long for POI's 1,024-byte limit, or any other failure in the record, leaves all three absent
      (`Embedded.ole10Names` catches it and returns null);
  - read when the run is made, so that the filter has them.

  A part that is not a compound file, or holds no `\x01Ole10Native`, gives none of the three. `olefile`, Q3's
  alternative, would replace only the container, not this parse.
- **`TextOptions(embedded_filter=...)`**: a callable `(obj: TextObject, address: str) -> bool`, None by default, set
  by name like `image_describer`, a host's hook beyond section 10's names.
  - It is asked in `_object` before an object is queued, for every object docx4j would read: an
    `EmbeddedPackagePart` of any kind and an OLE 1.0 package (Q4). A workbook, a presentation and an OLE 1.0 package
    are asked too, though this package defers reading them, so that a host can take them. Admitted, they raise
    `content.not_read` as today, where docx4j reads them: a departure in the output, not in the addresses.
  - The address index counts every object asked about, admitted or not.
  - A filter that raises declines, with `content.not_read` naming the object's part at the paragraph's address.
- **The warnings** (Phase A, since the re-pinned suite needs them):
  - `altchunk.not_expanded` carries `part_name` (the chunk's) beside `content_type`, and `TextWarning.selector`
    becomes the branch, else the part name, else the content type.
  - `Extractor.read_nested` (docx4j's `Extractor.nested`) takes the outer part, and every warning raised in the
    nested reading that carries a part name names the outer part instead, whatever its kind, as `Extractor.nested`
    does. Today those are `content.not_read`, `altchunk.not_expanded` and `image.not_described`; at the tag CR-069's
    `paragraph.drop_cap_joined` is one too, though docx4j CR-070 section 7 item 7 lists three.

### 4.4 The suite re-pinned (Phase A)

- `tests/fixtures/text-extraction/` is re-copied unedited from docx4j's `docx4j-core-tests/src/test/resources/text-extraction/`
  at the release tag `docx4j-17.3.5` (`64fae4cd0`; Q6, the user, 2026-11-01), with `SHA1SUMS`. The copy takes the files
  CR-011 took (`tests/README.md`'s rows) as they stand at the tag, less `phase0/expected-inspection-html-status.tsv`,
  which `4a37702eb` deleted, moving its row into `expected-inspection-status.tsv` under the capability column; and
  `dpp/` whole, its README included because its tables' columns are defined there. `tests/README.md`'s rows name the
  tag.
- `level-owner/` and `understood-namespaces.txt` stay as they are: they come from elsewhere in docx4j.
- The harness takes docx4j's changes to `ExtractionConformance`, `ExtractionFixtureChecks` and
  `ExtractionConformanceTest` between `198d015a3` and the tag:
  - the inspection tables' new columns (`capability` and those after it), the HTML-status table folded into the
    capability comparison as `4a37702eb` removed `htmlInspections`: `conformance.py`'s `HTML_STATUS` (lines 112 and
    817) and its count in `test_conformance.py` (line 208) go;
  - the findings keyed by part name;
  - the new cases' statements;
  - a run over `dpp/` beside `phase0/`;
  - `993e385e7`'s scenarios `drop_caps`, `small_caps` and `profile` (REQ-521, REQ-607, REQ-1002), which
    `scenario_requirements` (`conformance.py`) refuses today as unknown: `small_caps: upper` in
    `dpp/expected-tokens.tsv`, `profile: word-text` in all 22 `word-saves.tsv` rows;
  - `pendingReason`, which names in the report the CR-069 phase a pending row waits on;
  - `word-saves.tsv`'s check, `verifyWordSave`, as REQ-1002's NOTE reads a save:
    - the save is decoded by its byte order mark, else in the code page the row names;
    - the extraction is passed through that encoding, so a character the code page lacks is `?` on both sides;
    - the two are compared line by line, a line ending at CR LF, CR, LF or a tab, with trailing empty lines dropped;
  - a package Word saved, which lacks the builder's fixed ZIP timestamps (`ExtractionFixtureChecks.entries(..., built)`).

  That is all of docx4j CR-069's phase 0 (`dpp/`, `word-saves.tsv`, `verifyWordSave`, `entries(..., false)`), which
  its section 11.4 hands to this repository.
- Rows that need the inspector (REQ-901 to 907) stay pending, as its 214 do (CR-011 section 3.7). CR-050 phases 3
  and 4, CR-066 phase 4 and CR-069 phase 4 are inspector rows.
- `dpp/`'s rows pend on the existing rule wherever they cite a statement the engine does not advertise:
  - 33 of the 34 token rows (REQ-607 or REQ-710);
  - the 8 findings: 2 `paragraph.drop_cap_joined` (REQ-521) and 6 inspector rows;
  - both goldens;
  - all 22 of `word-saves.tsv`'s rows until Phase D, since their scenario is `profile: word-text` (REQ-1002).

  Two groups are checked earlier:
  - **Phase A:** the remaining token row, `H₂O` stored as U+2082 in `caps/caps-lang5.docx`, cites only REQ-602,
    which is advertised.
  - **Phase D:** six of `word-saves.tsv`'s rows cite only REQ-1002 and REQ-605, the table rule. They become checks
    once Phase D advertises REQ-1002, and Word's own text saves are then an oracle for `profile: word-text`. The
    other 16 cite REQ-521, REQ-607 or REQ-710.
- Phase A records what moves: every row that now passes, fails or pends, against 293 pending today. A text-model row
  that fails for a reason outside this CR is recorded and taken to the user, not fixed in passing.

### 4.5 Not here

- **The inspector** (docx4j's `Inspect`; D33's `fields.auto_update` and `fields.print_update`, `InspectionReport.markdown`):
  it is not ported, and its rows stay pending. It would be a CR of its own (CR-026 section 4.6).
- **Reading a workbook, a presentation or an OLE 1.0 package's content** stays deferred (CR-011 Q7). This CR gives
  a host what it needs to read them itself.
- **`math: unicodemath`** stays refused.
- **docx4j CR-069 from its phase 1** is the CR-069 port's (CR-029): case mapping, the drop-cap join and
  superscripts in the text, `small_caps` and `drop_caps` with the preset's two values (4.1), and `w:dstrike` as struck
  (its phase 4). Its phase 0 is this CR's Phase A (4.4).
- **docx4j `4a10eed5b`** (in the tag; WD 0.57 REQ-705) writes REQ-702's form in the HTML markup view's `<ins>` and
  `<del>` datetimes. It does not apply: there is no HTML writer here.
- **docx4j `02f9fc333`** (`VERSION_17_3_6`, after the tag) adds phase0 cases 83 to 85 in a sixth fixture there,
  `extraction-probe-opaque.docx`. It is outside this copy and comes with a later re-pin.

### 4.6 A part that cannot be read (Phase A)

- `trees._read` tells a part that is absent (None) from one that cannot be read, keeping the part and its exception
  (`PackageTrees.unreadable`).
- **The main part**, read at the top level by `extract(pkg)` or `extract(part)` (the main part, a header or a
  footer), raises `Docx4JException` (`openpackaging/exceptions.py`, the base `TextOptionError` already uses here)
  with docx4j's message, the part's exception its `__cause__`. Python has no unchecked exceptions, so docx4j's
  `Docx4JRuntimeException` has no other counterpart. The reason is the deepest `__cause__`'s message on one line, as
  `getReadableContents` writes it.
- **Nested**, `read_nested` gives `content.not_read` at the nested root, named for the outer part as 4.3's rule
  says, and reads nothing more of that package.
- **Every other story part** gives `content.not_read` with its part name, at docx4j's address (none for the notes,
  comments and glossary, the prefix for a header or footer), once per part; the body is read.
- No engine API is added: `XmlPart.read_contents` already raises, and the message is the text model's.

## 5. Tests and oracles

- **Phase A: the re-pinned suite.** `test_conformance.py` over the new copy, `phase0/` and `dpp/`, CR-070's twelve
  re-keyed rows and case 36's split among them; `dpp/`'s `H₂O` row checked and the rest recorded pending, each with
  its reason, `pendingReason`'s CR-069 phase among them; the three new scenarios mapped to their statements. The
  part names in the warnings, in `test_alt_chunks` (docx4j's `TextsPhase2Test.altChunks`). The nested chunk's warning
  names `/word/chunk47.docx`.
- **Phase A: a part that cannot be read.** `UnreadablePartTest`'s four text-model tests (2.4), over a package whose
  part's XML is broken as docx4j's test breaks it: the main part's message and cause; a header's; an embedded
  document's `content.not_read`; unreadable headers and notes as `content.not_read` with the body read.
- **Phase B: the dates.**
  - `RevisionDatesTest` ported whole (section 2.2's cases), with `trackedChanges` over `manual-redline-dates.docx`:
    D1's local time and no offset, D2's none, D3's `+00:00`, D4's `+11:00`, and the instant recovered from the two.
  - `MarkupDatesTest`: the four cases in plain text, in Markdown and as each run's `TextRevision.date`; a mark with no
    date as `{--A: gone--}`; a comment with and without its `dateUtc` (`R, 2026-10-31T11:00:00+11:00` and
    `R, 2026-10-31T11:00:00`), the parts built as docx4j's test builds them.
  - `TextsTest`'s TC9 dates as `fc71f12e6` changed them (`{++Probe, 2026-10-17T00:00:00: TC9_INSERTED ++}`, the
    `Z` gone).
  - The fixture `manual-redline-dates.docx` is read from `tests/fixtures/revisions/cr068/`, docx4j's
    `revisions/cr068/` copied once at the tag with `SHA1SUMS` (CR-028 Q11, decided 2026-11-02). Whichever of this Phase B
    and CR-028 Phase A lands first makes the copy.
- **Phase C: `EmbeddedObjectsTest`'s nine.**
  - The probe's five objects described.
  - An OLE 1.0 package's names, and none for a binary that is not one.
  - A linked object; a `w:control` in `w:pict`; an ActiveX `o:OLEObject`.
  - No filter, and a None one, read every object.
  - A filter declining one of two in a paragraph, its runs and addresses kept.
  - A filter over the probe, one level.
  - A failing filter.
  - The chunk warning's part name, a nested chunk's included.
  - The block model's JSON carries the `object` member.

  Four of the nine cannot be mirrored with docx4j's strings: they read content this package defers (section 4.5),
  which gives `content.not_read` here. `noFilterReadsEvery`, `filterDeclinesOneOfTwo` and `failingFilterDeclines`
  read `NOTES_TEXT`, `OTHER_TEXT` and `THIRD_TEXT` out of OLE 1.0 packages and assert no warning or exactly one;
  `filterOverTheProbe` expects `OLE41_NATIVE_TEXT` and `EMB53_SLIDE_ONE`. The named departure: the three filter tests
  are built over embedded OOXML documents, which this package reads, with docx4j's addresses and declines; over the
  probe, the four asks and `EMB40_DOCX_TEXT` are docx4j's, and the admitted OLE package and presentation give
  `content.not_read`.

  docx4j builds its OLE 1.0 packages at test time with POI (`EmbeddedObjectsTest.ole10`: `Ole10Native.writeOut` in a
  `POIFSFileSystem`); only the probe's `probe41.bin` is stored. Here the bytes come from a small compound-file writer
  in the tests, or are generated by docx4j and committed with `SHA1SUMS`. Phase C settles which; the recommendation
  is the test writer, which needs no Java run.
- **Phase D: `TextOptionsProfileTest`,** and `dpp/`'s six `word-saves.tsv` rows that need only REQ-1002 and REQ-605,
  checked against Word's saves. The values; both orders, through `set`, keywords, a mapping and attribute
  assignment; the map form; an unknown profile; the fields-fun memo under the preset
  (`tests/fixtures/fields-fun-cr067/`, CR-026's, or its own copy if CR-026 has not landed).
- **The gates.**
  - The suite.
  - `scripts/roundtrip.py`, 50 of 50.
  - `compare_acceptance.py`, 0 differ: nothing saved changes.
  - `codegen/generate.sh --check` for the new files, each in `KEEP`.
  - `tests/text/agreement-docx4j.tsv` and the goldens moving only where a row says the date's form changed.

No Word check: extraction output is not a file, and nothing this CR adds is written to one.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | The suite re-pinned at `docx4j-17.3.5` (4.4) and the harness's changes, `dpp/`, `word-saves.tsv`, the three scenarios and `pendingReason` included; `altchunk.not_expanded`'s part name and the nested warnings' outer part (4.3); a part that cannot be read (4.6, `UnreadablePartTest`'s four); every row's state recorded against 293 pending | 1¼ days |
| B | The dates: `model/revisions/dates.py`, `TrackedChange.author_local_time` and `author_offset`, `TextRevision.date` and `TextComment.date` in REQ-702's form, `author: ` alone; the two tests | ½ day |
| C | `TextObject` on every object run, the OLE 1.0 names (`model/text/ole.py`), `embedded_filter`, the JSON's `object` member; `EmbeddedObjectsTest`, four of its tests over OOXML objects, and the OLE 1.0 test bytes (5) | 1 day |
| D | `profile: word-text`; `TextOptionsProfileTest` | ¼ day |
| E | The records: CHANGELOG, with the Added and Changed entries listed after this table; CLAUDE.md's status and file list, and the index's row; the portfolio's `python/CR-027.*` entries through the docx4j-portfolio session | ¼ day |

Phase E's CHANGELOG entries:
- **Added:** `profile`, `TextObject` and `embedded_filter`, `TrackedChange.author_local_time` and `author_offset`.
- **Changed:** `TextRevision.date` and `TextComment.date` in REQ-702's form; `altchunk.not_expanded`'s selector the
  part name; a nested warning naming the outer part; a main part that cannot be read raises `Docx4JException`
  naming it, where the result was empty.
- **Fixed:** footnotes, endnotes, comments, a header, a footer or the glossary that cannot be read raise
  `content.not_read` (REQ-801), where they were left out silently.

## 7. Risks

- **The block model's dates change form** for every caller of `TextRevision.date` and `TextComment.date`, in every
  view, not only under `author_date` (Q2). docx4j's CHANGELOG states the same change for 17.3.5.
- **The re-pin may surface rows that moved for reasons outside this CR.** Docx4j's resources changed in nine commits
  between the pin and the tag (section 1). Phase A records each and takes any text-model failure to the user, rather than widening the CR.
- **The compound-file reader** is new code over a binary format (Q3). It reads only the one stream, is bounded by the
  part's size, and gives no names rather than raising on a malformed file. docx4j's fixtures are its oracle,
  POI's reading through `EmbeddedObjectsTest`.
- **`__setattr__` on `TextOptions`** (Q5) adds a check to every option assignment. It is a slotted class built a few
  times per extraction, so the cost is negligible; the alternative leaves direct assignment unrecorded.

## 8. Questions (decided 2026-11-02; Q6 2026-11-01)

The user, 2026-11-02: "CR-027's questions  Q1 to Q5 and Q7 as recommended". With Q6 (2026-11-01), every question below is decided.

| Q | Question | Recommendation |
|---|---|---|
| Q1 | `TrackedChange` gains `author_local_time` and `author_offset` (docx4j's `getAuthorLocalTime()` and `getAuthorOffset()`), beside `date`, the instant, unchanged? | Yes. They are docx4j's, they give what Word shows with no zone assumed, and `date` stays CR-025's reading. The content API's `Comment.creation_date` is unchanged. |
| Q2 | `TextRevision.date` and `TextComment.date` in REQ-702's form in the block model, in every view (they were `w:date` as stored)? | Yes, as docx4j's `TextRevision.getDate()` and `TextComment.getDate()` do since `fc71f12e6`. REQ-706 asks for it, and the block model and the serialisations then agree. CHANGELOG's Changed names it. |
| Q3 | An OLE 1.0 package's `label`, `file_name` and `command`: a small hand-written read-only reader of the compound file (`model/text/ole.py`), the `olefile` package as a new dependency, or leave the three absent (a named departure)? | The hand-written reader. It reads only the header, the allocation tables, the directory and one stream: a few hundred lines with no new dependency, held to docx4j's fixtures. `olefile` (BSD) is mature, but it would be this package's second runtime dependency for three strings. |
| Q4 | The filter asked for every object docx4j would read, including a workbook, a presentation and an OLE 1.0 package, which this package defers reading? | Yes. A host can then take the objects this package cannot read, which is the filter's purpose (REQ-522's NOTE). The addresses also agree with docx4j's. Admitted, they raise `content.not_read` as today. |
| Q5 | `profile`: every way a caller sets an option is recorded (keywords, `set`, `of`, a mapping, and direct attribute assignment through `__setattr__`), so either order gives the caller's value; `to_dict` carries `profile` last; `word-text`'s `math` is `latex`, documented beside the WD's NOTE? | Yes. It is docx4j's rule in Python's terms. Without the `__setattr__`, a caller who writes `opts.fields = "update"` and then `opts.set("profile", "word-text")` would get the preset's `cached`. |
| Q6 | The suite re-pinned to docx4j's current text-extraction resources, with the inspector's new rows pending as its 214 are, and every moved row recorded? | **Decided 2026-11-01:** copied from the release tag `docx4j-17.3.5` (`64fae4cd0`), the user: "CR-027 should copy docx4j's test suite from the 17.3.5 tag". The copy is the files CR-011 took, at the tag, and `dpp/` whole, so the pin stays a single docx4j commit, as CR-011 set it, and that commit is a release. |
| Q7 | Phasing: A (the suite and the warnings), B (the dates), C (the objects), D (the profile), E (the records), each built and verified before the next? | Yes. A first, so that B to D are measured against the suite they must pass. |

## 9. Effort

3¼ days: A 1¼, B ½, C 1, D ¼, E ¼.

## 10. Phase A (implementation notes, 2026-11-02)

Phase A as built: section 4.4 (the suite re-pinned and the harness's changes), section 4.3's warnings and
section 4.6 (a part that cannot be read), with section 5's two Phase A bullets. docx4j was read at its tag
`docx4j-17.3.5` (`64fae4cd0`), by `git show`, and at `4fdd80063` for section 4.6 only.

**Commits.** `1065287` (the copy and the harness, 10.1 to 10.4), `5844c0b` (the warnings, 10.5), `c7d4c81` (a part
that cannot be read, 10.6), and this section's.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the copy | `tests/fixtures/text-extraction/` | the files CR-011 took, as they stand at the tag, and `dpp/` whole; `SHA1SUMS` (84 files); `phase0/expected-inspection-html-status.tsv` removed |
| `.gitattributes` | the repository's root | `dpp/**/*.txt -text`, as docx4j's: Word's saves are compared as Word wrote them |
| the harness | `tests/text/conformance.py` | docx4j's changes from `198d015a3` to the tag (10.3); a run over `dpp/`; `main` writes both reports |
| the reports | `tests/text/conformance.tsv`, `conformance-dpp.tsv` (new) | `phase0/`'s and `dpp/`'s runs, in docx4j's form |
| the adapter | `tests/text/adapter.py` | `DOCX4J_IMPLEMENTED` at the tag (70), `NOT_ADVERTISED` (5, each with its reason) |
| the warnings | `model/text/result.py`, `extractor.py`, `embedded.py` | 10.5 |
| a part that cannot be read | `model/text/trees.py`, `extractor.py`, `__init__.py` | 10.6 |
| the tests | `tests/text/test_conformance.py`, `test_unreadable_parts.py` (new), `test_texts_phase2.py`; `tests/test_validator.py`, `tests/openpackaging/test_mce_typed.py` | 10.7, 10.9 |

No file under `docx4j_py/` is new, so `KEEP` is unchanged.

### 10.2 The copy

- Each file is `git show docx4j-17.3.5:docx4j-core-tests/src/test/resources/text-extraction/<path>`, byte for byte.
  `SHA1SUMS` lists them, and `test_conformance.py` holds the folder to it and, with the sibling checkout, each file to
  docx4j's blob at the tag.
- Changed since `198d015a3`: `GOLDENS.md`, `INSPECTION.md`, and in `phase0/` `case-statements.tsv`,
  `expected-findings.tsv`, `expected-inspection.tsv`, `expected-inspection-status.tsv`, `expected-tokens.tsv` and
  `extraction-probe-inspect.docx`. The rest are the same blobs.
- The tables: 319 token rows (+34), 257 findings (+56, 242 of them the inspector's), 82 case statements (+15), 2
  orders, 12 goldens; 242 inspection rows (+56) and 8 status rows, one per fixture and capability. `dpp/`: 34 token
  rows, 8 findings, 2 goldens, 22 Word saves. These are section 1's numbers.
- `phase0/expected-inspection-html.tsv` is kept: unchanged at the tag, the source of the reader-present rows, which
  `expected-inspection.tsv` now holds. The harness no longer reads it.
- Not copied, as before: `phase0/README.md`, `RUN.txt`, `baseline/`, `evidence/` and `support/`; the folder root's
  `README.md`, `expected-tokens.tsv`, `extraction-probe.docx` and `run-phase0.sh`; and `bind-props.docx` (section 1).
- `level-owner/` and `understood-namespaces.txt` stay as they were.

### 10.3 The harness

docx4j's `ExtractionConformance`, `ExtractionFixtureChecks` and `ExtractionConformanceTest`, `198d015a3` to the tag,
and `ExtractionInspection`'s constants:

- **The capability column** (`4a37702eb`). A fixture is compared against its rows of `any` and of the engine's
  environment, with that environment's status row. The other environment's rows pend by count ("15 rows of html reader
  present, not this run's environment"). A fixture with rows of an environment needs REQ-513. The inspection's
  scenario is the environment, where it was `default`. `_html_inspections` and `HTML_STATUS` are gone, `parsed` is
  one of `FIELDS`, and REQ-903's counts no longer need `html reader absent`.
- **The findings keyed by part name**: no harness change; the tables carry them.
- **The new case statements**: 82 cases, each token row citing its case's statements and fixture, tested.
- **CR-069's scenarios** `drop_caps`, `small_caps` and `profile` (REQ-521, REQ-607, REQ-1002).
- **`pending_reason`**: CR-069's latest phase among the missing statements, else CR-046's phase or the inspector.
- **`word-saves.tsv`**, through `verify_word_save`: the save decoded by its byte order mark, else in the row's code
  page; the extraction passed through that encoding; the lines split at CR LF, CR, LF or a tab, trailing empty lines
  dropped. Its tests are docx4j's `wordSaveIsReadByItsEncodingAndComparedByLine`.
- **A run over `dpp/`** beside `phase0/`, reported in `conformance-dpp.tsv`, as docx4j's test writes `dpp.tsv`.
- **`entries(data, built=True)`**: a package's members, with the builder's fixed 1980 timestamp where `built`.
  `phase0/`'s five carry it; `dpp/`'s bisect and caps files do not, and are read with `built=False`. No reader here
  needs it, as docx4j's adapter does for the inspector's selectors; it is tested as a fixture check.

**The adapter** advertises `DOCX4J_IMPLEMENTED`, docx4j's 70 at the tag, less the inspector's 8 and `NOT_ADVERTISED`:

- REQ-1002 (`profile: word-text`): Phase D;
- REQ-521, REQ-607, REQ-710: the port of docx4j CR-069 (CR-029);
- `text.formatting_join`: an inspector kind, CR-069 phase 4.

**The port's gate.** With docx4j's statements advertised and nothing behind them, the harness makes 1,772 checks over
`phase0/` (36 pending with no HTML reader, 40 with one) and 192 over `dpp/` (0 and 2). docx4j's own last run here
(`docx4j-core-tests/target/text-extraction/`, read only) was checked against them, less what came after the tag (the
sixth fixture's 28 checks, `02f9fc333`; the 12 reader-present goldens, `27887b28a`). The checks are the same 1,772 and
192, in docx4j's order, with docx4j's 36 and 0 pending checks and their reasons.

### 10.4 Every row, against 1,291 pass, 0 fail and 293 pending

| report | before (`198d015a3`'s copy) | after the copy (`1065287`) | after Phase A (`5844c0b`, `c7d4c81`) |
|---|---|---|---|
| `phase0/` | 1,291 pass, 0 fail, 293 pending of 1,584; goldens 12 of 12 | 1,258, 36, 478 of 1,772 | **1,294, 0, 478 of 1,772**; goldens 12 of 12 |
| `dpp/` | --- | 3, 0, 189 of 192 | **3, 0, 189 of 192**; goldens 0 of 2 |

The moves in `phase0/`, by check (fixture, case, item, view, scenario):

- **33 passing checks gone, 36 new ones passing.** The 11 `altchunk.not_expanded` rows keyed by content type, each
  checked in three views, are now docx4j CR-070's 12 keyed by the chunk's part name. Case 47's row was `html reader
  present` and is now `default`, so checked. The 36 failed after the copy ("expected 1, got None") and pass with the
  warnings (10.5).
- **10 pending checks gone.** The inspector's `inspect:altchunk` rows of cases 36, 44 and 63 under `default`, now one
  row per environment (3); case 47's `html reader present` warning in three views (3); the four inspections under
  `default`, now under the environment (4).
- **1 pending check's reason changed.** The HTML fixture's inspection under `html reader present`: the inspector's
  before, now "15 rows of html reader present, not this run's environment".
- **195 new pending checks.** The token checks of cases 68 to 82, which cite REQ-904 (CR-050 phase 3, CR-066 phase 4;
  123); new inspector rows of `default` and `html reader absent` (41) and of `html reader present` (18); the five
  inspections under the environment (5); the fifth fixture's counts (6); the other environment's rows by count (2).
- No check that passed fails. No row moved for a reason outside this CR.

Pending at the end, `phase0/`, 478:

- the inspector's 400: 382, and 18 that also need an HTML reader (CR-011 section 3.7);
- Q7's 60, by name (`adapter.DEFERRED`);
- docx4j's 15 extraction checks under `html reader present`;
- the other environment's rows, by count, for 3 fixtures.

docx4j's own 36 pending checks are among them. Nothing pends for `NOT_ADVERTISED` in `phase0/`.

`dpp/`: case 5's `H₂O`, stored as U+2082 in `caps/caps-lang5.docx`, passes in three views (REQ-602). The 189 pending:

| reason | checks |
|---|---|
| CR-069 phase 1 (REQ-607; with REQ-1002 for 13 Word saves) | 118 |
| CR-069 phase 2 (REQ-521: the two `paragraph.drop_cap_joined` rows in three views, and C02's save) | 7 |
| CR-069 phase 3 (REQ-710; the two goldens; two saves) | 52 |
| CR-069 phase 4 (`text.formatting_join`) | 4 |
| the inspector (`text.near_background`) | 2 |
| phase 1 or 2; REQ-1002 alone: the six saves Phase D makes checks | 6 |

### 10.5 The warnings (section 4.3)

- `altchunk.not_expanded` carries `part_name`, the chunk's, beside `content_type`.
- `TextWarning.selector` is the branch, else the part name, else the content type.
- `Extractor.read_nested(package, prefix, story, outer)`: every warning of the nested reading that names a part names
  `outer` instead, whatever its kind. `chunk_blocks` and `package_blocks` pass the chunk's or the embedded package's
  part. The probe's case 47 chunk holds `/word/nested.html`; its warning names `/word/chunk47.docx`.
- `TextWarning.__str__` is docx4j's `toString`: the kind, ` at ` and the address, then the branch, the content type and
  the part name, each in brackets where present. It wrote the selector alone, which after the change would have lost
  the content type. `test_alt_chunks` asserts docx4j's strings.
- Today the warnings that name a part are `content.not_read`, `altchunk.not_expanded` and `image.not_described`.

### 10.6 A part that cannot be read (section 4.6)

- `trees._read(part, unreadable)` records a part that cannot be read, with its exception, in
  `PackageTrees.unreadable`; `unreadable_error(part)` asks it. An absent part is still None.
- `PackageTrees` reads the rest of a package whose main part cannot be read, where it stopped. So a header of such a
  package can be read alone, as docx4j's `Texts.of(Part)` reads it.
- `trees.not_readable(part, error)` is docx4j's message: "The content of /word/document.xml could not be read, so it
  cannot be read as text: <reason>". The reason is the deepest `__cause__`'s message on one line, its whitespace
  collapsed, or the class's name where there is none. Here it is the parser's, "Failed to create target class
  `Document`".
- `Extractor.extract`: at the top level it raises that, from the part's exception. Nested, it gives `content.not_read`
  at the nested root (`body/0/object:0`, `body/1/altchunk:0`), which `read_nested` names for the outer part, and reads
  nothing more of that package. A chunk so read gives no `altchunk.not_expanded`: it was read.
- `extract(part)`, for the main part, a header or a footer, raises it too. A header that cannot be read raised
  `TypeError("not a story part")`.
- The footnotes, endnotes, comments, a header or footer and the glossary that cannot be read each give
  `content.not_read` with the part's name, once, at docx4j's address: a header's or footer's prefix, else none. Each is
  checked only where the options read it (`notes`, `comments` or markup, `include`), as docx4j's `storyContent` is.
- `extract()`'s docstring "Raises" names `Docx4JException` and `TypeError`.

### 10.7 The tests and the numbers

- `test_conformance.py`: 59 (24 before). The ratchet over both reports; the pending checks by cause; `dpp/`'s checks;
  the port's gate in four runs; the copy (`SHA1SUMS`, docx4j's blobs, the tables' counts, the case statements,
  `entries`); the harness's rules, with CR-069's three scenarios, `pending_reason`, the Word saves and the inspection
  by capability.
- `test_unreadable_parts.py`, new, 13: `UnreadablePartTest`'s four text-model tests; each other story part (6); a
  chunk; a header of an unreadable package; the part's route left as it was.
- `test_texts_phase2.py`: `test_alt_chunks` is docx4j's `altChunks` at the tag, the probe's nested chunk folded in
  (it was a test of its own); a nested `image.not_described` naming its chunk.
- The fast suite: 7,104 passed, 8 skipped; 7,055 at `80e5910` by count (49 new tests: 35, 13 and the validator's 1).
- The full suite: 8,803 passed, 8 skipped (10.10).
- ruff (`check`, `format --check`, at the files' 120 columns) on every file changed: no new finding. The findings there
  before are left as they were: `check`'s in `extractor.py` and `test_texts_phase2.py` (one each), and the formatter's
  hunks in `conformance.py`, `result.py`, `extractor.py`, `test_conformance.py`, `test_texts_phase2.py`,
  `test_validator.py` and `test_mce_typed.py`.

### 10.8 Departures, with why

1. **The tests' break.** docx4j cuts a closing tag short (`</w:body>` to `</w:body`). This package's parse recovers
   from that and reads the part: xsdata's lxml handler parses with `recover=True`. So the tests break the root's start
   tag (`<w:document` to `<<w:document`), which no recovery mends. 10.11 says what the recovery means.
2. **The exception.** docx4j throws a `Docx4JRuntimeException` whose cause is the `Docx4JException` naming the part,
   whose cause is the read's. Here it is the `Docx4JException`, its `__cause__` the part's own (`Problem with part
   /word/document.xml: ...`): Python has no unchecked exceptions (section 4.6).
3. **The reason's class name**, where a cause has no message, is the module-qualified name (bare for a builtin), not
   Java's.
4. **A failing Word save's reason** is `line n: Word wrote [...], the extraction [...]`, without docx4j's path: the
   port's reasons leave out what the row's columns give.
5. **`entries`** reads the DOS fields through `zipfile`, which gives the central directory's. docx4j's `dosTimes` reads
   them so since `4e4f8ce49`; before, the JDK's extended timestamp made the check depend on the time zone.
6. **The two reports** are this port's committed files. docx4j's test writes `conformance.tsv` and `dpp.tsv` under
   `target/`.

The harness's two departures (CR-011): a failing check is a `FAIL` row, and `Engine.deferred`.

### 10.9 Pre-existing tests whose expectation changed, and why

- `tests/openpackaging/test_mce_typed.py::test_the_facade_still_lists_a_drawing_in_a_run`: the inspect probe has 8
  inline pictures at the tag, case 66's 2 and 6 of cases 68 to 74. It asserted 2.
- `tests/test_validator.py::test_the_corpus_holds_nothing_but_the_probes_values`: the inspect probe's cases 75 to 82
  are CR-066 phase 4's out-of-range values, and `misplaced()` reports them. They are exempted as the CR-066 probes'
  are, and a new test holds them to the inspector's `value.out_of_range` rows: 16, the inspector's 17 less `w:gridCol`,
  which no round measured and `integer-limits.json` leaves out.
- `test_texts_phase2.py::test_alt_chunks` (10.5).

### 10.10 The gates

| gate | after |
|---|---|
| the fast suite | 7,104 passed, 8 skipped, 0 failed |
| the full suite | 8,803 passed, 8 skipped, 0 failed (6 min 32 s), at `c7d4c81`. A first run, beside other work, failed `test_performance.py`'s list check by 1.4 ms against its 201.6 ms limit; alone, and in the second run, it passed |
| `conformance.py` | `phase0/` 1,294, 0, 478 of 1,772; `dpp/` 3, 0, 189 of 192 |
| ruff on the changed files | no new finding |

No Word check: no file this package saves changes.

### 10.11 Found, and left open

1. **A part's XML is mended silently.** The lenient parse recovers from XML that is not well formed and reads what it
   can, with no skipped report, even under `LoadOptions(strict=True)`. A `word/document.xml` cut after "SEC" of its
   second paragraph reads as `FIRST PARAGRAPH`, `SEC`, its third paragraph gone. That breaks this repository's rule
   that nothing is dropped silently, and docx4j refuses such a part (CR-065 D1). It is the engine's and the fork's
   (`recover=True` in xsdata's lxml handler), not this CR's: for the user.
2. **`extract(part)` unmarshalled the part.** Telling a story part from content asked `hasattr(target, "contents")`,
   which unmarshals the part for real (REQ-204) and raised the part's own exception before the text model could name
   it. It now asks the class. Fixed, since section 4.6 needs it; a test holds both routes.
3. **`dpp/`'s 189 pending checks** wait on CR-029 and Phase D; the inspector's 400 on a CR of their own (section 4.5).

## 11. Phase B (implementation notes, 2026-11-03)

Phase B as built: section 4.2 (the dates), with section 5's Phase B bullets. The user ordered it on 2026-11-03 ("do
CR-027 phase b, then CR-030"). docx4j was read at its tag `docx4j-17.3.5` (`64fae4cd0`), by `git show`:
`RevisionDates`, `RevisionDatesTest`, `MarkupDatesTest`, `TextsTest`, and `fc71f12e6`'s changes to `TrackedChange`,
`Extractor`, `PlainText`, `TextRevision`, `TextComment` and `WmlToMarkdown`. None of what this phase ports changed
between `fc71f12e6` and the tag.

**Commits.** `a21e8cc` (the copy, 11.2), `7f42fa1` (the code and the tests, 11.1, 11.3 and 11.4), and this section's.

### 11.1 What landed

| piece | where | what |
|---|---|---|
| the rule | `model/revisions/dates.py` (new; in `KEEP`) | docx4j's `RevisionDates`: `local_time`, `offset` and `text`, with `MIN_OFFSET` and `MAX_OFFSET` |
| the change | `model/revisions/change.py` | `TrackedChange.author_local_time` and `author_offset` (Q1), a group's its first piece's; `date` unchanged |
| the extractor | `model/text/extractor.py` | `_revision` reads `w:date` and `w16du:dateUtc` from the mark, typed or wildcard; the comment reader takes `dateUtc` by `_last_para_id` (`_comment_dates_utc`); `_date_text` is gone |
| the writers | `model/text/plain.py`, `markdown.py` | under `author_date`, `author: ` alone where the date is None; the plain writer's docstring |
| the block model | `model/text/result.py` | `TextRevision.date` and `TextComment.date` document the form, and the old one before it |
| the copy | `tests/fixtures/revisions/cr068/`, `.gitattributes` | 11.2 |
| the tests | `tests/revisions/test_revision_dates.py`, `tests/text/test_markup_dates.py` (new); `tests/text/test_texts.py` | 11.4 |

The content API's `Comment.creation_date` and `TrackedChange.date` are untouched.

### 11.2 The copy

- `git archive docx4j-17.3.5 docx4j-core-tests/src/test/resources/revisions/cr068 | tar -x` into a scratch directory,
  then copied into `tests/fixtures/revisions/cr068/`. That is 54 files, about 566 KB: every file there but the four
  screenshots, which Q11 leaves in docx4j (CR-028 section 5.1). This is the one copy of that directory here, and
  CR-028 Phase A reads it rather than making its own.
- `SHA1SUMS` lists the 54. The `README.md` is this package's: it names the tag and its commit, each file, and the
  four left out. It is not in `SHA1SUMS`. Every file's git blob id is docx4j's at the tag.
- `.gitattributes` marks the whole directory `-text`, after Phase A's `dpp/**/*.txt` rule. The `.txt` files are
  Word's Save As text (CR LF lines, a bare CR between cells), and `cr068-log.txt` is UTF-8 with a byte order mark and
  CR LF lines. Marking the whole directory keeps every file to `SHA1SUMS` on a checkout with `core.autocrlf`.
- The sdist takes `tests/` whole (`pyproject.toml`), so it needs no exclude. The wheel holds no test.

### 11.3 The rule, as built

- **The inputs.** Each function takes an `XmlDateTime`, its lexical form or a `datetime`, read by
  `values.date_time`, CR-025's reader of a `w:date`'s form. The model keeps a `w:date` as an `XmlDateTime`, or as
  its text where the exact parse refuses it, a date alone among them (CR-025 Phase B2). `values.date_time` trims
  the value, reads a date alone as midnight, and reads anything else that is no `xsd:dateTime` as no date.
- **Fractions** of a second are dropped from every input, in `_clock`. `dates_utc` gives an aware `datetime` that
  keeps microseconds, so the truncation section 4.2 asks for is the rule's own, not a step in the comment reader. A
  `dateUtc` of `...:00.5Z` gives the offset (a test holds it).
- **`24:00:00`**, which `xsd:dateTime` allows, is the next day's midnight. That is how Java's
  `XMLGregorianCalendar` reads it (checked with JDK 21: `2026-10-31T24:00:00Z` reads as 2026-11-01 0:00:00).
- **`offset`** is `datetime.UTC` for zero, else a fixed `datetime.timezone`. **`text`** writes the local time with
  `isoformat(timespec="seconds")` (Java's `uuuu-MM-dd'T'HH:mm:ss`) and the offset as `+HH:MM` (Java's `xxx`).
- **The marks.** `_revision` serves every mark the text model reads: a run-level revision, a row's and a cell's
  marker, and an equation's `w:ins` and `w:del`. docx4j's `revision()` serves the same, the equation's since
  `fc71f12e6`. Every typed class among them is a `CTTrackChange`, which carries `date_utc`. A wildcard mark is read
  by its `{w16du}dateUtc` attribute.
- **A comment's `dateUtc`.** `_comment_dates_utc` builds a `CommentParts` over the comments tree this reading read
  privately, and asks `dates_utc`, the content API's own reader, so that one chain reads the two parts.
  `comment_parts_of` would unmarshal the comments part (`contents`) and break REQ-204. Neither lxml part is parsed:
  `_root_for_read` parses its bytes aside. A part that cannot be read gives no date, as docx4j's null part does.
- **`_last_para_id`** is unchanged: the last paragraph carrying a `w14:paraId`, descending into the `mc` branch the
  walk reads, as docx4j's `lastParaId` does. That is not `creation_date`'s last paragraph with an entry, and a test
  holds the difference. The two differ only where that paragraph has no entry. Every comment in this repository's
  fixtures that has a `commentsIds` entry has one paragraph, so none differs.

### 11.4 The tests and the numbers

- `tests/revisions/test_revision_dates.py`, new, 19 tests:
  - `RevisionDatesTest` whole: `measuredCases`, `offsets` and `otherForms`, each over the model's form (an
    `XmlDateTime`, or a date alone as text) and over the lexical form (6), and `trackedChanges` over
    `manual-redline-dates.docx` (1);
  - the copy: every file in `SHA1SUMS` and nothing else but the README, and each file docx4j's blob at the tag,
    the tag's other four the screenshots (2);
  - what is no date (6);
  - `24:00:00` (1);
  - a `datetime` read to the second (1);
  - a group's stamp, its first piece's (1);
  - `date` as CR-025 read it, in the engine and the content API (1).
- `tests/text/test_markup_dates.py`, new, 8 tests:
  - `MarkupDatesTest`'s two;
  - the comments in Markdown;
  - `markup_detail: author` unchanged;
  - the last paragraph carrying a `w14:paraId`, against `creation_date`;
  - a `dateUtc` with a fraction;
  - the block model's JSON in the `markup` view (Q2);
  - a reading that parses neither side part (REQ-204).
- `tests/text/test_texts.py`: `test_markup`, docx4j's `TextsTest.markup` at the tag. This repository had ported
  only its first assertion (in `test_the_views`), so TC9's `author` and `author_date` lines, the run's revision and
  the block's count are new here, TC9's date as `fc71f12e6` writes it (`{++Probe, 2026-10-17T00:00:00: TC9_INSERTED
  ++}`, the `Z` gone; `inserted.date == "2026-10-17T00:00:00"`).
- The fast suite: 7,132 passed, 8 skipped. Phase A's was 7,104, and the 28 are this phase's (19, 8 and 1).
- The full suite: 9,128 passed, 8 skipped. Phase A's was 8,803.
  - Of the 325 more, 28 are this phase's tests.
  - The other 297 come from `openpackaging/test_private_readers.py`: its corpus globs `tests/fixtures/revisions/`,
    so its eleven readers now read the copy's 27 documents, and all pass.
  - `revisions/test_every_fixture.py` already read the sibling's `cr068/`, and is unchanged.
- ruff (`check`, `format --check`, at 120 columns) on every file changed: no new finding. `dates.py`'s one `DTZ001`
  carries a `noqa`, since the author's clock is naive by design. The formatter would change the same number of lines
  in the five existing files as before (11, 61, 22, 20 and 37), and the new files are formatted. `extractor.py`'s one
  `check` finding (S110) is 10.7's, left as it was.

### 11.5 The suite, the agreement table and the goldens

- **The conformance suite.** No row of `phase0/` or `dpp/` cites REQ-702 or REQ-706, and none pends on either: the
  478 and 189 pending checks are 10.4's, for 10.4's reasons. `conformance.py`'s two reports are byte for byte the
  committed ones: `phase0/` 1,294 pass, 0 fail, 478 pending of 1,772, goldens 12 of 12; `dpp/` 3, 0 and 189 of 192.
  Nothing moves.
- **The inspector's rows.** The two `revision.author` rows (cases 9 and 14, under `options: view: markup;
  markup_detail: author_date`) carry docx4j's inspector's own date range, `2026-10-17T00:00:00Z to
  2026-10-17T00:00:00Z`. They pend with the inspector (CR-011 section 3.7), and this phase leaves them alone.
- **`tests/text/agreement-docx4j.tsv`** compares the plain text of the accepted and original views, which write no
  date. No line moves.
- **The goldens.** `phase0/`'s `.markup.txt` and `.md` are written under `markup_detail: type`, which writes no date.
  None moves, as docx4j's commit says of its own ("No suite golden moves").

### 11.6 Departures, with why

1. **A year outside 1 to 9999** is no date here. Python's `datetime` cannot hold one, and Java's `LocalDateTime` can.
   No producer writes one.
2. **One `dateUtc` per durable id.** `dates_utc` maps a durable id to the paragraph id of the last
   `w16cid:commentId` naming it. docx4j's `commentDatesUtc` maps each `commentId`'s paragraph id. So where two
   paragraph ids share one durable id, both get its date in docx4j and only the last here. Word writes one durable id
   per comment, and section 4.2 reads `dates_utc`, so the content API and the text model keep one reader.
3. **`dateUtc`'s lexical forms** are `dates_utc`'s `datetime.fromisoformat`, which also takes ISO 8601's basic form
   (`20261031T000000Z`). `XMLGregorianCalendar` refuses that form. Word writes the extended form.
4. **The types.** The local time is a naive `datetime`, for Java's `LocalDateTime`. The offset is a
   `datetime.timezone` (`datetime.UTC` for zero), for Java's `ZoneOffset`. With both known,
   `author_local_time.replace(tzinfo=author_offset) == date`, as section 4.2 has it.
5. **The inputs** are an `XmlDateTime`, its text or a `datetime`, where docx4j's are an `XMLGregorianCalendar`. The
   model keeps a date alone as text (CR-025 Phase B2), and `dates_utc` gives a `datetime`.
6. **docx4j's test helper `x(s)`** is ported twice: each case of the first three tests runs over the model's form and
   over the lexical form.

### 11.7 Pre-existing tests whose expectation changed, and why

None. No test pinned `TextRevision.date` or `TextComment.date` in the old form. No plain-text or Markdown test wrote
`author_date`: `test_private_readers.py` reads under it, but checks only the trees. `test_markup` is a new test in
`test_texts.py`, and none of that file's other tests changed.

### 11.8 The gates

| gate | after (`7f42fa1`) |
|---|---|
| the full suite | 9,128 passed, 8 skipped, 0 failed (6 min 42 s) |
| the fast suite | 7,132 passed, 8 skipped, 0 failed |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50 parts canonically identical, 0 differences, 0 skipped, 34 benign respellings (`boolean-spelling`) |
| `codegen/generate.sh --check` | two generations byte identical; the generated tree unchanged and `dates.py` kept |
| `scripts/compare_acceptance.py` | `out/acceptance/` built by `scripts/acceptance.py` before the change (at `9825b87`) and after it: 17 artefacts, 246 entries, 0 differ beyond the wall-clock dates |
| `conformance.py` | as 11.5, both reports byte for byte the committed ones |
| ruff on the changed files | no new finding |

No Word check: extraction output is not a file, and nothing this package saves changes.

### 11.9 Found, and left open

1. **`tests/README.md`'s `test_private_readers.py` row** says "the 71 of `samples/`, `fixtures/text-extraction/` and
   `fixtures/revisions/`". It was 96 before this phase and is 123 with the copy. It is not this phase's row, so it is
   left as it was.
2. **The inspector's date.** docx4j's inspector at the tag writes a mark's `w:date` with its `Z` (11.5). When the
   inspector is ported (section 4.5), its form is docx4j's inspector's, not REQ-702's.
3. Phase E's CHANGELOG entries for this phase are section 6's, and CLAUDE.md's status follows at Phase E. The index's
   row records Phase B now. The portfolio's `python/CR-027.B` goes through the docx4j-portfolio session.

## 12. Phase C (implementation notes, 2026-11-05)

Phase C as built: section 4.3's objects (the warnings were Phase A's), with section 5's Phase C bullet. The user ordered
C, D and E together on 2026-11-05. docx4j was read at its tag `docx4j-17.3.5` (`64fae4cd0`), by `git show`:
`TextObject`, `EmbeddedFilter`, `Embedded`, `Extractor.object`, `control` and `admitted`, `EmbeddedObjectsTest`, and its
copy of POI (`Ole10Native`, `IOUtils`, `StringUtil`, `LittleEndianByteArrayInputStream`, `DirectoryNode`). The first
three and the test are the same at `VERSION_17_3_6`'s head.

**Commits.** `d8f3faf` (the code and the tests), `6f0d15a` (two packages Word wrote, 12.3), and this section's.

### 12.1 What landed

| piece | where | what |
|---|---|---|
| the object | `model/text/result.py` | `TextObject`, `TextRun.object`, `TextRun.to_dict`'s `object` member |
| the reader | `model/text/ole.py` (new; in `KEEP`) | `root_stream`, `record_names`, `package_names` |
| the names | `model/text/embedded.py` | `ole10_names(part)`, docx4j's `Embedded.ole10Names` |
| the extractor | `model/text/extractor.py` | `_object`, `_ole_object`, `_control`, `_admitted`; `_object_run` takes the description |
| the option | `model/text/options.py` | `embedded_filter`, by name like `image_describer`; the `EmbeddedFilter` alias |
| the exports | `model/text/__init__.py` | `TextObject`, `EmbeddedFilter` |
| the tests | `tests/text/test_embedded_objects.py`, `test_ole.py` (new); `ole_writer.py` (new, not a test); `mini_docx.py` gains `external` | 12.5 |

### 12.2 The object and the filter, as built

- **`TextObject`** is a frozen slotted dataclass, as `TextImage` is. `part_name` and `content_type` are properties over
  `part`, as docx4j's getters are. `__str__` is docx4j's `toString`, in the WD's lower-case kinds.
- **`kind`:** a part whose content type is exactly one of the two ActiveX types is `control`, ahead of `link`, as docx4j
  orders them, though a link has no part to test. `o:OLEObject/@Type` is `link` only as exactly `Link`: a value the exact
  parse keeps as text reads as `embed`, as JAXB's null does.
- **An OLE 1.0 package** is a ProgID beginning `package` in any case, whose part is a `BinaryPart` and no
  `EmbeddedPackagePart`. Its names are read when the run is made, under every option, one level down too. Reading the
  bytes unmarshals nothing (REQ-204; a test holds it).
- **The filter** is asked in `_ole_object`, with the address `para.objects` gives, which counts before the question, so a
  decline moves no address. A link, a control and a plain OLE binary are not asked and take no index. Nested, `embedded`
  is off, so nothing is asked one level down.
- **A filter that raises** (any `Exception`) declines, with `content.not_read` at `prefix.rpartition("/")[0]`, the
  anchoring paragraph's address.
- **A package the filter admits and this package defers** raises `content.not_read` from `_EmbeddedObject.emit`, as it
  did, at the same address (Q4). One the filter declines raises nothing, so a host that takes them sees no warning.
- **The wildcard forms.** A `w:object` in a kept branch is no root element, so a wildcard (CR-012 Phase C). `_object`
  reads it, its `o:OLEObject` and its `w:control` by name, with the same members; `source` is then the `AnyElement`.

### 12.3 The reader, as built

- **`_CompoundFile`** reads the header, the FAT's sector list (the header's 109, then the DIFAT's), the directory's chain
  and, for one stream, its chain: the FAT's, or the mini FAT's and the mini stream's under 4,096 bytes. The sector size is
  the header's (512 or 4,096, its byte 30, as POI reads it); the mini sector and the cutoff are constants, as POI holds
  them.
- **The bound.** A chain or a DIFAT that meets itself, and a FAT sector listed twice, are refused, so a walk is bounded by
  the file. A stream is read only as far as its size, itself held to POI's 100,000,000 bytes. A free sector ends a chain,
  as one does in a damaged file POI reads.
- **`record_names`** is `Ole10Native(byte[], int)`, its failures kept: a record that ends at its data's end gives no names
  (POI's `readShort` throws there), and so does one whose UTF-16 names fail, though none is kept. A negative length is
  refused at once. In POI it reads on and fails at the next read or the string's conversion; the result is the same.
- **`label` and `file_name`** decode as Java decodes. Python's `errors="replace"` agrees with JDK 21 but for an encoded
  surrogate (`ED A0 80`), three U+FFFD in Python and one in Java, so those three bytes are replaced first.
- **`package_names`** catches everything, as `ole10Names` does. A compact or unparsed record gives None, where POI gives
  three nulls: the same to `TextObject`.

**Held to POI.** `EmbeddedObjectsTest`'s oracle is POI's reading, so a scratch harness, not committed, ran docx4j's
`Ole10Native` and `POIFSFileSystem` from `docx4j-core`'s jar against this reader:

| what | cases | POI gives names | differ |
|---|---|---|---|
| `Ole10Native` records: valid ones, each cut at every length, every limit at and past it, and 60,000 with random bytes changed | 60,327 | 22,110 | **0** |
| intact files: POI's two, `probe41.bin`, the writer's forms, the names without case, a storage of the name, no stream | 22 | | **0** |
| the same files damaged, 1 to 4 bytes changed, some cut short | 40,000 | | 714, 1.8%, each a damaged structure (12.6, item 5) |

The harness stands in the scratch directory, and needs `docx4j-core`'s jar and commons-io, so it is no test here. What
it found is: an encoded surrogate (12.3), a sector shift POI reads by its one byte, and counts POI reads as signed.

Two packages Word wrote are in the sibling checkout, `OLE/dragged txt.docx` and `OLE/create new bitmap.docx`. POI reads the
first's names and the second as `unparsed`, and so does this reader. `test_ole.py` reads them from `../docx4j`, skipped
without it, since section 5 settles no copy.

### 12.4 The test bytes

Section 5 left it to this phase: the test writer, as recommended. `tests/text/ole_writer.py` is `ole10_record` (POI's
`writeOut` in its parsed form) and `compound_file` (version 3 or 4, the streams under 4,096 bytes in the mini stream,
a DIFAT past 109 FAT sectors). The harness read the writer's files through POI too, so it writes what POI reads.

### 12.5 The tests and the numbers

- `tests/text/test_embedded_objects.py`, new, 25 tests:
  - `EmbeddedObjectsTest`'s nine, section 5's four over OOXML objects;
  - the object's edges: a link with an internal relationship, a relationship that does not resolve, an absent `Type`, the
    binary ActiveX type, a wildcard `w:object`, a run that is no object's (6);
  - the JSON and `__str__` (3);
  - the filter over docx4j's OLE 1.0 packages, as docx4j asks it, and declining them all; an object that would not be read;
    one level down (4);
  - the option (2);
  - a reading leaving every part as it was (1).
- `tests/text/test_ole.py`, new, 105 tests:
  - the compound file (36): POI's file, Word's two, eight stream sizes in each sector size (16), the DIFAT, the name without
    case, what gives no names, chains and a directory that meet themselves, 1,500 damaged files raising nothing;
  - the record (69): the control first byte's 32 values, `flags1`, a high first byte, a malformed name, POI's limits, the
    UTF-16 names, a record failing after its data, and the total size.
- The fast suite after this phase: 8,751 passed, 8 skipped, at `d8f3faf` with the user's commits beside it.
- ruff (`check`, `format --check`, at 120 columns): the new files are clean. `extractor.py` has `check`'s one S110 from
  10.7, and `format` the hunks it had.

### 12.6 Departures, with why

1. **Four of `EmbeddedObjectsTest`'s nine** are section 5's: the three filter tests over embedded documents, so their
   strings and addresses are `body/0/object:1/body/0`, where docx4j's are `body/0/object:1/0`, and `filterOverTheProbe`'s
   admitted OLE package and presentation are `content.not_read`. docx4j's asks and declines over OLE 1.0 packages are
   held by a test of their own.
2. **A filter's result is its truth**: `bool(...)`, so a filter returning None declines. Java's returns a `boolean`.
3. **`Exception`** for Java's `RuntimeException`: Python has no checked exceptions.
4. **The kinds** are the WD's lower case, as section 4.3 has them, where docx4j's enum is upper case, in `__str__` too.
5. **A damaged compound file.** POI reads the whole structure at load, this reader what one stream needs, so they
   differ only on a damaged file. A broken mini FAT with the stream in FAT sectors gives names here and fails in POI; a
   broken mini stream chain past the stream's end fails here and reads in POI. Every intact file agrees (12.3).
6. **`source`** of a wildcard `w:object` is an `AnyElement` (12.2). docx4j's is typed.

### 12.7 Pre-existing tests whose expectation changed, and why

None. `test_objects_are_object_runs_with_their_progids` stands: `prog_id` is unchanged.

### 12.8 Found, and left open

1. **A record ending at its data's end gives no names** (12.3), by POI's `readShort`, so by docx4j's. Word's package
   (`dragged txt.docx`) has its UTF-16 names after the data, so it reads. A writer that stops at the data would not, and
   no measurement says one exists: section 4.3's rule is followed.
2. **REQ-512's Markdown clause** ("the preview picture with the alt text `Embedded object: <ProgID>`") is built in neither
   docx4j nor here: no source has the string, and no suite row asks for it. It is CR-011's clause, not this phase's.

## 13. Phase D (implementation notes, 2026-11-05)

Phase D as built: section 4.1, its amendment of 2026-11-05 with it, and section 5's Phase D bullet. docx4j was read at its
tag, by `git show`: `TextOptions`, `TextOptionsProfileTest`. The one value after the tag is `87c421731`'s.

**Commits.** `1b6192c` (the code, the tests, the memo and `conformance-dpp.tsv`), `6f0d15a` (one more test, 13.9), and this section's.

### 13.1 What landed

| piece | where | what |
|---|---|---|
| the option | `model/text/options.py` | `profile`, last in `OPTION_NAMES` and `to_dict`; `_explicit`; `__setattr__`; `set` over a private `_apply`; `_apply_profile`; `_WORD_TEXT`, `_PROFILES` |
| the NOTE | `TextOptions`' docstring | what the preset does not reproduce |
| the adapter | `tests/text/adapter.py` | REQ-1002 advertised, out of `NOT_ADVERTISED` |
| the report | `tests/text/conformance-dpp.tsv` | 13.5 |
| the memo | `tests/fixtures/fields-fun-cr067/` | `one-file-many-readings.docx` at the tag, `SHA1SUMS`, `README.md`: section 5's own copy, CR-026 not having landed; the blob is docx4j's |
| the tests | `tests/text/test_text_options_profile.py` (new); `test_conformance.py`, `test_texts.py`, `test_duplicate_entry_warnings.py` | 13.6 |

### 13.2 The recording, as built

- **`__setattr__`** assigns, then records a name in `_RECORDED`, every option but `profile`. Beyond the options it costs
  one set lookup.
- **`__init__`** assigns its defaults through it and clears `_explicit` before the keywords, so a default is nobody's.
  That is section 4.1's "past the recording" by its result.
- **`set`** is `_apply` and a record. `_apply` writes through `object.__setattr__`, so the preset's own values record
  nothing. `include(...)`, docx4j's builder, records `include`.
- **`copy()`** writes every slot past the recording, `_explicit` as a new set.
- **`_apply_profile`** normalises the name (`strip`, `lower`, `_` to `-`, which is docx4j's `profileOf` turned round), and
  fills what `_explicit` lacks. None clears the name and takes nothing back, as docx4j's `profile(null)` does.
- **`to_dict`** carries `profile` last, so that `TextOptions.of(options.to_dict())` gives the same options: every other
  option is then the caller's, and the preset fills nothing.

### 13.3 The preset

`_WORD_TEXT` is section 2.1's 24 values, `list_labels: true` among them, and `math: latex`. `TextOptions.java`'s log after
`0cf4fb38a`, each commit's change to `WORD_TEXT_VALUES`:

| commit | the preset | here |
|---|---|---|
| `1560ab379` (CR-070) | nothing | |
| `75b9bdb54` (CR-069 phase 1) | `small_caps: upper` | CR-029's (4.1) |
| `15c66769f` (CR-069 phase 2) | `drop_caps: asis` | CR-029's (4.1) |
| `7af22e9dd` (WD 0.80) | nothing; a new option, `default_language`, left at its default | CR-029's |
| `87c421731` (WD 0.84) | `list_labels: true` | taken (4.1's amendment) |

So nothing moved beyond what section 4.1 says. docx4j's has 26 values, this 24.

### 13.4 The NOTE

`TextOptions`' docstring holds section 2.1's list with 4.1's addition, each with what the text has instead: the notes' and
comments' forms, CRLF lines and CR cells, the one space after every label, folded hyphens, Word's linear math and `math:
latex`, and the three accidents. It adds the ANSI code page, which WD REQ-1002's NOTE lists and section 2.1 does not, since
the WD asks for "at least those the NOTE below lists". A test holds each phrase.

### 13.5 The suite, the agreement table and the goldens

Against section 11.5's numbers:

| report | section 11.5 | after this phase |
|---|---|---|
| `phase0/` | 1,294 pass, 0 fail, 478 pending of 1,772; goldens 12 of 12 | the same, byte for byte |
| `dpp/` | 3, 0 and 189 of 192; goldens 0 of 2 | **9, 0 and 183 of 192**; goldens 0 of 2 |

Every row of `conformance-dpp.tsv` that moved, 22, all `word-saves.tsv`'s:

- **6 from PENDING to PASS:** `bisect/control.txt`, `C03`, `C04`, `C08`, `C09`, `C10`, as section 4.4 says. Word's own saves
  are now an oracle for `profile: word-text`.
- **16 still PENDING, the reason's statement list losing REQ-1002:** `C02` (REQ-521), `C05` and the twelve under `caps/`
  (REQ-607), `C06` and `C07` (REQ-710). Their phases are unchanged.
- The pending reasons `phase 1` (5) and `phase 2` (1) are gone. The rest are 10.4's: CR-069 phase 1, 118; phase 2, 7; phase 3,
  52; phase 4, 4; the inspector, 2.

`tests/text/agreement-docx4j.tsv` and the goldens do not move: the preset is opt-in, and no row of them uses it. The objects
add no text: `conformance.tsv` was byte for byte the same after Phase C.

**What the six do not show.** Each of the six also passes under the default options: their documents hold nothing the
preset changes. They show the preset takes nothing away from Word's text, not that each value is Word's. The 16 are where
the values bite, and they wait on CR-029.

### 13.6 The tests and the numbers

- `tests/text/test_text_options_profile.py`, new, 40 tests: `TextOptionsProfileTest`'s four, its unknown name over 5 forms;
  the 24 values, and nothing else moved; each route in either order (6); attribute assignment and `include`; `comments:
  false` (13.9); the defaults and `copy()` recording nothing; the name's 6 forms; `to_dict`, a round trip, `repr`;
  `check_implemented`; the memo's copy (2); the docstring's list; a list label under each `w:suff` (3).
- `test_conformance.py`: `test_dpp_holds_what_holds_before_cr_069` and `NOT_ADVERTISED`'s set.

### 13.7 Departures, with why

1. **`profile` is a property over a `_profile` slot**, where section 4.1 says a slot. As a slot, `opts.profile = "word-text"`
   would set the name and apply nothing, and `to_dict` would then say so falsely. Assigning applies it, as `comments` is a
   property over `_comments`.
2. **24 values, not the tag's 26** (4.1), and `list_labels` from after the tag (4.1's amendment).
3. **`TextOptionError`** (reason `invalid`) for Java's `IllegalArgumentException`, as every refusal here is.
4. **The ANSI code page** in the docstring, beyond section 2.1's list (13.4).
5. **`set` records any name it accepts**, as docx4j's does, the hooks (`image_describer`, `embedded_filter`, `field_user`)
   among them. `__setattr__` records only section 10's names. No preset names a hook, so neither shows.

### 13.8 Pre-existing tests whose expectation changed, and why

- `test_texts.py::test_the_options_are_section_10s_names`: 28 names, `profile` last.
- `test_duplicate_entry_warnings.py::test_the_option_is_accepted_by_name_with_its_two_values`: `duplicate_entries` is
  second to last.
- `test_conformance.py`'s two, for 13.5.

### 13.9 Found, and left open

1. **A caller's `comments: false` does not hold under the preset**, unless the caller sets `include` too. `comments` is
   true where `include` names it, as docx4j's `isComments()` is, and the preset's `include` names it. That is docx4j's
   reading, so it is built and pinned (`test_comments_false_under_the_preset_is_docx4js_reading`), but WD REQ-1002 says a
   caller's option overrides the preset's. For the user. **Taken up by [CR-035](CR-035-one-setting-for-comments.md)** (the user, 2026-11-05): one setting for comments.
2. **The Word-save comparison splits at a tab, not at a space.** WD 0.79 has it split at a space too and tolerate Word's
   space after a label. The harness is at the tag and `dpp/`'s 22 saves hold no list, so nothing shows. A later re-pin brings
   it.
3. **`default_language`** (docx4j `7af22e9dd`, after the tag) is no option here. CR-029 is amended for it.

## 14. Phase E (implementation notes, 2026-11-05)

The records, as section 6 lists them: `CHANGELOG.md`'s Added, Changed and Fixed entries for all of CR-027; `CLAUDE.md`'s
status and file list; the index's row; `tests/README.md`'s rows, the three new test files, `ole_writer.py`, the memo, and
the `dpp/` and option counts; this CR's Status line. The registry goes through the docx4j-portfolio session.

**Commits.** This section's.

### 14.1 The gates

Run together, after Phases C and D, at `6f0d15a`. The records' commit reflows one call in `extractor.py` and one comment in
`test_texts.py`, and `tests/text/` passed again after it.

| gate | result |
|---|---|
| the full suite | 10,789 passed, 8 skipped, 0 failed (7 min 55 s). At `1b6192c`, 10,786 |
| the fast suite | 8,800 collected, 8 skipped, by count, all within the full run. Its last run of its own, after Phase C: 8,751 passed, 8 skipped, 0 failed |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50 parts canonically identical, 0 differences, 0 skipped, 34 benign respellings (`boolean-spelling`) |
| `codegen/generate.sh --check` | two generations byte identical (321 modules); the generated tree unchanged, and `ole.py` kept |
| `scripts/compare_acceptance.py` | `out/acceptance/` copied before Phase C and rebuilt by `scripts/acceptance.py` after: 17 artefacts, 246 entries, 0 differ beyond the wall-clock dates. A build at `4095345` itself gives the same |
| `conformance.py` | as 13.5: `phase0/` 1,294, 0, 478 of 1,772, goldens 12 of 12; `dpp/` 9, 0, 183 of 192, goldens 0 of 2; each report byte for byte the committed one |
| `tests/text/agreement-docx4j.tsv` | unchanged |
| ruff on the changed files | no new finding. The new files are clean; `extractor.py`'s one `check` finding (S110) and the formatter's hunks in the existing files are as they were |

No Word check: section 5. Extraction output is not a file, and nothing this CR adds is written to one. `compare_acceptance.py`
says the same of what is saved.
