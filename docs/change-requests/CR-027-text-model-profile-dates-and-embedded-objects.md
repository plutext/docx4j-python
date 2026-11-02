# CR-027: the text model's WD 0.41 to 0.49 additions --- `profile: word-text`, revision and comment dates as the author's clock, embedded objects for a host, and the conformance suite re-pinned

**Status:** Proposed 2026-10-31, at the user's word ("fold it into CR-027 with profile and the CR-070 items"). It covers five
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
- `text_boxes: false`, `notes: true`, `note_references: true`, `list_labels: false`;
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
  - The fixture `manual-redline-dates.docx` is copied from docx4j's `revisions/cr068/` into
    `tests/fixtures/manual-redline-dates/`, with `SHA1SUMS`.
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
