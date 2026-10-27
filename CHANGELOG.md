# Changelog

The distribution `docx4j` (import name `docx4j_py`), on PyPI at
<https://pypi.org/project/docx4j/>. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the versions [Semantic Versioning](https://semver.org/) as far as `0.x` goes: what each
surface promises in `0.x` is the table at the end of this file. Each entry summarises what a
change request delivered and links the CR's own implementation notes, which are the record; the
CRs are under [`docs/change-requests/`](docs/change-requests/README.md).

## [Unreleased]

### Changed

- **A measurement a setter would write beyond Word's range is refused**
  ([CR-022](docs/change-requests/CR-022-integer-measurements-refused-beyond-words-range.md), sections
  [10](docs/change-requests/CR-022-integer-measurements-refused-beyond-words-range.md#10-phase-a-implementation-notes-2026-10-26)
  and [11](docs/change-requests/CR-022-integer-measurements-refused-beyond-words-range.md#11-phase-b-the-records-2026-10-27);
  the counterpart of docx4j CR-066 and docx4j-core-ts CR-002 section 65). Word 365 stores a twips or
  half-point measurement in 16 bits and keeps the low 16 bits of anything larger: 0.2.0's
  `paragraph.left_indent = 400000` wrote `w:left="8000000"`, which Word reads as -3,959 twips (docx4j
  CR-065's probe). Every setter that writes an integer measurement now raises `MeasureRangeError` (a
  `ContentError` and a `ValueError`, code `value.out_of_range`) **before it changes anything**, the
  message giving the value and Word's range in the caller's units and the hint what to pass.
  - **The content API** refuses where Word 365's own Office JS setters refuse, at their limits as
    measured ([CR-024](docs/change-requests/CR-024-office-js-setters-as-word-runs-them.md), core-ts's
    check 71b), each on the value as written: `space_before` and `space_after` 0 to 1,584 pt,
    `line_spacing` 0.75 to 1,584 pt (so `line_spacing = 0` is refused where it wrote `w:line="0"`),
    `Font.size` 1 to 1,638 pt after rounding to a half-point (0.9 is taken; 0 still removes `w:sz`
    and `w:szCs`, as documented since 0.1.0; a negative is refused), a level's text indent short of
    1,584 pt either way, a cell width that is not negative; and at the faithful range elsewhere: the indents
    (-1,638.4 to 1,638.35 pt; a first-line indent to 1,638.35, a hanging one to 1,638.4), a cell's
    width, a level's marker indent.
  - **The python-docx facade** refuses a `Length` Word would misread, where python-docx writes it:
    a section's page size and seven margins, a paragraph format's indents, spacing and line spacing
    (under the rule the write leaves), tab stops, `Font.size`, a row's height (under its rule), a
    cell's width and the column widths `add_table` and `add_column` divide out.
    `line_spacing_rule` and `height_rule` hold the value already set to the new rule's range.
  - **The builders** refuse `size=` (`r`, `p` and everything over them), `tc(width=)` and `tbl`'s
    column widths and their sum.
  - **The limits** are per attribute, measured, never assumed: 60 rows in
    `docx4j_py/resources/integer-limits.json` (`values.integer_limit`, `integer_limits()`,
    `unmeasured_attributes()`), from docx4j CR-066's probe rounds 5 and 6 (Word 365, 2026-10-26; the
    group files and Word's saves are fixtures in `tests/fixtures/integer-limits/`, and a test holds
    every row to them). An attribute the probe did not measure is not checked. The seven attributes
    Word writes 65,535 back in as `"-1"` (`w:before`, `w:after`, the five unsigned page margins) stop
    at 65,534; an exact line spacing or row height is 1 to 32,768 twips, from Word's saves; a negative
    in an attribute Word stores unsigned is refused.
  - A value already in a file is left as it is (CR-021 Q7: the wrap is never emulated); CR-020
    Phase B reports it with the same table and code. No sensible measurement is that large, so the
    only code this breaks set a value Word was already misreading.

- **An integer attribute is read exactly as written** ([CR-021](docs/change-requests/CR-021-integer-attributes-read-as-word-reads-them.md)
  phase 0; the fork's stage 14). `" 240 "`, `"1_000"` and digits from other scripts used to bind as
  integers through Python's `int()`, where Word refuses the document. They're now kept as their text
  and listed in `part.kept`. In WordprocessingML a save drops them; elsewhere it writes the integer,
  as before.

### Fixed

- **`Paragraph.line_spacing` writes a multiple of single spacing, as Word does**
  ([CR-024](docs/change-requests/CR-024-office-js-setters-as-word-runs-them.md); a defect 0.1.0 and
  0.2.0 shared, the port of docx4j-core-ts CR-002 section 65.9). It wrote `w:lineRule="exact"`, a
  fixed height, where Word 365's own `lineSpacing` setter writes `w:line` as 20 x the points under
  `w:lineRule="auto"` (core-ts's check 71b): 12 is single spacing, 24 double, whatever the font.
  It now writes Word's, replacing an exact or at-least rule already there, and reads back as set.
  `to_api_script` expresses an auto rule, stated or the default, and falls back to `insert_xml` for
  an exact or at-least one. The python-docx facade's `ParagraphFormat.line_spacing` is unchanged.

- **`List.set_level_indents` writes what Word writes**
  ([CR-023](docs/change-requests/CR-023-set-level-indents-as-word-writes-them.md); a defect 0.1.0
  and 0.2.0 shared, the port of docx4j-core-ts CR-002 section 65.5). Its second indent, Office JS's
  `bulletNumberPictureIndent`, is the marker's place relative to the text: a first-line indent,
  written `w:firstLine` when zero or more and `w:hanging` of its magnitude when negative, as Word
  365 writes it (core-ts's check 71). It was written to `w:hanging` as given, so the usual list,
  `set_level_indents(0, 36, -18)`, wrote `w:hanging="-360"`, and `(0, 36, 18)` wrote a hanging
  indent where Word writes a first-line one. `to_api_script` scripts a level's indents the same way.

- **Integer attributes are read and saved as Word 365 reads them**
  ([CR-021](docs/change-requests/CR-021-integer-attributes-read-as-word-reads-them.md), section
  [10](docs/change-requests/CR-021-integer-attributes-read-as-word-reads-them.md#10-phases-a-and-b-implementation-notes-2026-10-26);
  a defect 0.1.0 and 0.2.0 shared). A decimal where WordprocessingML wants an integer (Google Docs'
  `w:tblW w:w="10206.0"` and `w:pgSz w:w="11907.0"`, pandoc 2.2.1's `tblW`, the
  `w:hanging="141.99999999999994"` family) or an empty one (SQL Server Reporting Services 2012's
  `w:pgMar w:header=""`) was kept as its text and written back after an edit. The readers then
  disagreed: `Paragraph.left_indent` read "720.5" as 36.025 pt, and the python-docx facade's
  `section.page_width`, `header_distance` and table column widths raised `ValueError`.
  - **Reading:** `docx4j_py.model.values.integer` reads Word's form, `[+-]?[0-9]*(\.[0-9]*)?` or the
    empty string, truncated toward zero, the empty string 0. Every reader of an integer attribute
    goes through it, held by a census.
  - **Saving:** a save writes Word's integer and drops a form Word refuses (`"abc"`, `"1e3"`,
    surrounding whitespace). It drops an empty `w:rsid*` too, as Word does. Each change is recorded
    in `last_save.normalised`, and `save(normalise=False)` keeps the source.
  - **Scope:** WordprocessingML (`w`, `w14`, `w15`), where Word was measured. Elsewhere a save writes
    what 0.2.0 wrote.
  - **The rule** is Word's, from docx4j CR-065's shared probe (four rounds, Word 365, 2026-10-26).

- **The PyPI page's figures** ([`docs/pypi-readme.md`](docs/pypi-readme.md)): 0.2.0's page still
  described 0.1.0 ("What is and is not in 0.1", 45 parity goldens, a 1 MB wheel of 238 modules,
  6 MB installed, a 0.6 s import, and 148,000 lines of generated code, which was the whole
  package's source at 0.1.0). It now gives 0.2.0's: 46 goldens, a 1.5 MB wheel of 308 modules,
  7.5 MB installed, a 0.7 s import and 112,000 lines of generated code (the files outside
  `codegen/clean.py`'s `KEEP`), under "What is and is not in 0.x".

### Dependencies

- `docx4j-xsdata[lxml] >=26.2.4,<26.3`: the fork's
  [docx4j-xsdata 26.2.4](https://pypi.org/project/docx4j-xsdata/26.2.4/), published 2026-10-26 (its
  `docs/fork/CHANGES.md` stage 14, item 32: `ParserConfig(exact_types=)` takes `int`).

## [0.2.0] --- 2026-10-25

Still WordprocessingML only; PresentationML and SpreadsheetML remain CR-001 Phase D and CR-002
Phase C. Seven change requests are new and complete, each Word-checked:

- [CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md): tracked changes
  listed, accepted and rejected as Word 365 resolves them (`docx4j_py.model.revisions`).
- [CR-011](docs/change-requests/CR-011-one-text-model.md): one text model under every reader of
  text, with Markdown, charts, equations and field switches as docx4j reads them
  (`docx4j_py.model.text`). The Office JS readers (`Paragraph.text`, `get_text`, `search`) read the
  branch Word shows, as Office JS does (Phase F).
- [CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md): `mc:AlternateContent` kept at
  every level, with its branches typed, VML among them.
- [CR-013](docs/change-requests/CR-013-ranges-across-paragraphs.md): Office JS's own characters,
  which every text writer takes back (Phase A), and a `Range` that reads and writes across
  paragraphs as Office JS's does (Phases B and C).
- [CR-018](docs/change-requests/CR-018-run-properties-as-word-reads-them.md): every on/off value,
  enumeration and colour of the run properties read as Word 365 reads it, kept as written, and
  written on save as Word reads it. This closes a save 0.1.0 shared that changed meaning
  (`w:b w:val=" false"`, bold in Word, written `0`). It needs docx4j-xsdata 26.2.3.
- [CR-016](docs/change-requests/CR-016-cell-ends-fonts-placeholders-and-wildcards.md): Word's table
  rule, a cell ends with a paragraph, kept through every edit. This closes a defect 0.1.0 shared:
  eleven edits left a cell Word opens as unreadable content. `Font`, a control showing its
  placeholder and a deleted paragraph mark read as Office JS reads them. Word's wildcards match
  across a paragraph mark, and no hit cuts a grapheme cluster.
- [CR-019](docs/change-requests/CR-019-writing-into-a-placeholder.md): every write into a content
  control showing its placeholder replaces it as Word 365 does, by every route, and placeholders
  are written in Word's own form, naming a glossary entry, the glossary written on save where a
  file has none. This closes a defect 0.1.0 shared: a value written into a control showing its
  placeholder, by any route, left the control a placeholder. A write at a paragraph's end goes
  after everything the paragraph holds, as Word 365 writes it, which closes another: 0.1.0 wrote it
  into a trailing control, field result or tracked insertion (Phase D).

[CR-009](docs/change-requests/CR-009-theme-colours.md)'s Phase B resolves theme colours as Word
draws them, so `Font.color` reads a theme colour in the theme's colour rather than the cache written
beside it. With these come CR-003's `Range.hyperlink`, CR-009 Phase A's colour merge rules, CR-002
section 16's single parse of a part, docx4j 17.3.5's `DOCPROPERTY` changes (a name matched in any
case, `Keywords` a core property), and a fix to a defect 0.1.0 shipped: an embedded object would
not open in Word after a save.

### Added

- **Theme colours as Word draws them: `docx4j_py.model.color`**
  ([CR-009](docs/change-requests/CR-009-theme-colours.md) Phase B, section
  [12](docs/change-requests/CR-009-theme-colours.md#12-phase-b-implementation-notes-2026-10-24)),
  the WordprocessingML half of docx4j's `org.docx4j.model.color` (its CR-034, at `a9e0303ff`) as
  docx4j-core-ts ports it (`76fb45d`): `ResolvedColor` (red, green, blue and OOXML's alpha),
  `ThemeColors` (a docx's theme part, or the bundled theme `pkg.default_theme` names where it has
  none, with the settings' `w:clrSchemeMapping`) and `WordColors` (`resolve`, `of` for a
  `w:color`, a border or a `w:u`, `fill` and `pattern`; `normalise` and `stated` are
  `docx4j_py.model.values`' colour reader). A theme reference wins over the value beside it,
  `w:themeTint` and `w:themeShade` change the HSL luminance (the tint where both are given), and
  Word's written value is kept where it is within one unit of the computation. The
  `PropertyResolver` gains `theme_colors` (made once, read again on `refresh()`),
  `resolve_color`, `resolve_fill` and `resolve_pattern_color`. Held to docx4j's vectors, copied
  into `tests/fixtures/theme-colors/`: the 117 tint and shade triples of four corpora within one
  unit (76 exact), and every case of the P1 and P2 probes through a resolver at docx4j's own
  tallies. DrawingML and SpreadsheetML colours are Phase C's.

- **Word's table rule: `ends_with_paragraph` and `end_with_paragraph`**
  ([CR-016](docs/change-requests/CR-016-cell-ends-fonts-placeholders-and-wildcards.md) Phase A,
  section [10](docs/change-requests/CR-016-cell-ends-fonts-placeholders-and-wildcards.md#10-phase-a-implementation-notes-2026-10-24)),
  exported from `docx4j_py.model.content` (`model/content/table_rules.py`), docx4j's
  `TableRules` (`18a3a31f4`) as docx4j-core-ts ports it (`222cc20`): whether a cell's last block,
  as Word reads it (past range markup, through block-level controls, custom XML and the branch
  Word reads of a kept `mc:AlternateContent`), is a paragraph, for a `w:tc` or its list; and
  Word's repair, an empty `w:p` appended after any wrapper, returned (None where the rule held).
- **On/off values read as Word 365 reads them: `docx4j_py.model.values`**
  ([CR-018](docs/change-requests/CR-018-run-properties-as-word-reads-them.md) Phase A, section
  [10](docs/change-requests/CR-018-run-properties-as-word-reads-them.md#10-phase-a-and-c-implementation-notes-2026-10-24)),
  against the user's six rounds of Word 365 probes (2026-10-24), copied into
  `tests/fixtures/onoff-fun/`. `on_off(value, default=True)`: an ST_OnOff value is **off for
  exactly `false`, `off` or `0`**, case-sensitive and untrimmed, and **on for anything else**
  (`on`, `""`, `" false"`, `TRUE`, `yes`, ...); `is_on(element)` for a `w:b`-like element;
  `xsd_boolean` for DrawingML's `xsd:boolean`. Every reader of an on/off value goes through them:
  the resolver's toggles and `StyleUtil.toggle`, `Font`, the text model and Markdown, hidden
  text, the font selector, content controls, table header rows, `w:trackRevisions`, numbering,
  OMML, the custom XML kinds and the python-docx facade. Held to Word's own saves of the 28
  probes: over 376 markers, every reader's reading of the eleven measured toggles equals its
  reading of Word's save. Public, for a caller who reads the typed tree (Q13).
- **The enumerations and colour read as Word 365 reads them** (CR-018 Phase B, section
  [11](docs/change-requests/CR-018-run-properties-as-word-reads-them.md#11-phase-b-and-d-implementation-notes-2026-10-24)),
  in `docx4j_py.model.values`: `in_enumeration` and `underline`, `highlight`,
  `vertical_alignment`, `emphasis_mark` (a member only for exactly a member's value, so `SINGLE`,
  `YELLOW` and `bogus` are unset; `w:u`'s six on/off words read as `none`), and `colour(value)`
  (six hex digits in any case, trimmed, one `#` allowed, or one of the 140 classic HTML colour
  names, `NETSCAPE_COLOURS`, as `RRGGBB`; exactly `auto` as `AUTO`; anything else None, which
  is as if absent in a text or underline colour and an automatic that overrides a style's in a
  fill or a border). They are in the resolver's merge, so an invalid direct value lets the
  style's through, and under `Font.underline`, `highlight_color` and `color`, the facade's
  `underline`, `highlight_color`, `superscript`, `subscript` and `ColorFormat`. A character
  style's `w:highlight` is never drawn, as in Word. Held to Word's saves of the 33 probes (seven
  rounds; round seven's from docx4j `5163c2986`): every reader agrees at all 388 markers.
- **`effective(obj, name)`** (CR-018 Phase C, section 4.4): the one reader of a value the file
  leaves out, the stated value else the schema default typed (the fork's
  `XmlVar.schema_default_value`), in the namespaces whose defaults are audited (`w`, `w14`,
  `w15`; `LookupError` elsewhere). The model's fields never change. `w:fldLock`,
  `w:documentProtection/@w:enforcement`, `customMarkFollows`, `w:style/@w:default`,
  `w:customStyle`, `w:multiLine` and `w14:checked` read through it.
- **`part.kept` and `package.kept`** (CR-018 Q16): the values the parse kept as their text, the
  fork's `KeptValue`s, `"exact"` (kept on purpose, `w:val=" false"`) or `"invalid"` (`"off"`,
  `"Yellow"`); `describe()` counts them by reason (`"kept"`). Not skipped content:
  `LoadOptions(strict=True)` does not raise on them.
- **Placeholders in Word's form: `ContentControl.placeholder_entry`, `show_placeholder()`,
  `default_placeholder_entry` and `shows_words_form_placeholder`**
  ([CR-019](docs/change-requests/CR-019-writing-into-a-placeholder.md) Phase B, section
  [12](docs/change-requests/CR-019-writing-into-a-placeholder.md#12-phase-b-implementation-notes-2026-10-25);
  extensions, core-ts's `placeholderEntry`, `showPlaceholder`, `defaultPlaceholderEntry` and
  `showsWordsFormPlaceholder` at `aefc41f` and `b12e0c5`): `placeholder_entry` is the entry
  `w:placeholder/w:docPart/@w:val` names (`""` for a `w:placeholder` naming none, None without one),
  which decides the control's form; `show_placeholder()` shows a copy of that entry's content as the
  glossary holds it (core-ts check 68), with `w:showingPlcHdr`, a control naming none, or an entry
  the glossary lacks, made to name its kind's default (core-ts section 63.18, docx4j CR-063's D8);
  `default_placeholder_entry` is Word's default entry for the control's kind, a `DefaultEntry` (its
  name and text), None for a picture, a check box or a repeating section;
  `shows_words_form_placeholder` is whether an empty control shows a placeholder in Word's form, as
  the binding pass asks. With them, from Phase A, `shows_placeholder_content`: whether a text write
  replaces the placeholder (a block or run control showing it).
- **`package.glossary_entries`** (CR-019 section 4.9, D9; core-ts's `pkg.glossaryEntries`): the
  glossary's entries, read privately (`GlossaryEntries`: `text_of`, `body_of`, `has`, `pending`),
  and the entries the writes name, kept until the save writes them in Word's form; a `dry_run` has
  its own, dropped with it. `docx4j_py.model.content.placeholders` holds it, with
  `DEFAULT_ENTRIES`, `default_entry_for` and `DEFAULT_PLACEHOLDER_ENTRY`.
- **`OpcPackage.save_hooks` and `package.last_save.added`** (CR-019 section 12.5 item 4): callables
  the save runs before it writes anything, each given the package and returning the names of the
  parts it added, which `last_save.added` lists. The content API registers the glossary's there;
  the engine imports no model class.
- **Word 365's runs of core-ts's checks 63 to 70, and Word's default placeholders**
  (`tests/office_js/`, Apache-2.0 here; CR-019 sections 10.3, 11.4.1, 12.1, 12.10 and 15.1): each
  check's file, Word's JSONs and Word's saves, check 63's run 2 (Design Mode off) beside its first
  run, kept as the record of what Design Mode did; `tests/office_js/checks.py`'s ports of the eight
  Script Lab snippets (`run_check_63` to `run_check_70`); and `placeholders-word365.docx`, the user's
  Word 365 file of one control of each kind, the oracle of each kind's default entry.

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
  row's cells, as Word 365 reads it there. `LoadOptions(mc_load_mode=...)`, and
  `pkg.mc_load_mode` over it for one package, read when each part is first parsed:
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

- **A `DOCPROPERTY` name is matched in any case, and `Keywords` is a core property**, as docx4j
  `VERSION_17_3_5` reads them (its `ea5461ee9`, #244, measured in Word 365; and `f15bce627`, #277):
  `DocPropertyResolver` takes a custom property's name exactly first, else in any case, and
  the built-in names in any case. So under `fields: update` the text model computes
  `DOCPROPERTY client` from a property named `Client`, and `DOCPROPERTY title` from the title,
  where 0.1.0 kept the cached result. docx4j's probe and Word's save of it are in
  `tests/fixtures/docprop-case/`: every result Word's but a missing property's, which keeps its
  cached result where Word writes its error text, as docx4j keeps the old one
  (CR-011 section 17).
- **`Font.color` reads a theme colour as Word draws it** (CR-009 Phase B, question 2): where the
  effective `w:color` has a `w:themeColor`, the effective read reports the theme's colour with its
  tint or shade, through the document's theme and `w:clrSchemeMapping`, where 0.1.0 reported the
  `w:val` beside it, a cache that can be stale. So a Heading 1 in docx4j's default styles (written
  `365F91`, Office 2010's, under `accent1` shaded `BF`) reads `#0F4761` in a new document's Office
  2023 theme, as Word 365 reads it (Office JS check 53). The direct read
  (`get_font(direct=True)`) keeps `w:val`; an automatic or unstated colour still reads `"auto"`
  or `""`. `Font.to_dict()` gains `theme_color`, `theme_tint` and `theme_shade`, as written. Over
  the 728 documents of `samples/` and `out/cr012-b-corpus/`, 119 of 58,555 effective answers move,
  every one a theme reference meeting a stale cache (or `auto`) beside it; no direct answer
  moves. The python-docx facade's `ColorFormat` keeps python-docx's semantics (question 3). A
  `w:themeColor` or `w:themeFill` outside `ST_ThemeColor` (`Accent1`) is no reference, as docx4j
  reads it, so a style's colour beneath shows (CR-009 section 12.10 item 1).
- **A table cell ends with a paragraph, as Word reads it, after every edit** (CR-016 Phase A,
  core-ts CR-002 section 55, held to Word 365 in check 52's cases 86 to 95). `Paragraph.delete()`,
  `Body.clear()`, `ContentControl.delete(keep_content=False)` and `Body.insert_ooxml(...,
  location="Replace")` keep the paragraph a cell ends with, emptied, its `w:pPr` kept, where
  removing it would leave the cell without one (a control's in the control's place, outside it);
  tracked, its text is deleted and its mark never marked deleted, as Word does. A table put after
  that paragraph (`insert_table`, `insert_xml`, `insert_ooxml`, `insert_markdown`,
  `insert_element` at `"End"`, or `"After"` it) splits it at its end: it keeps its text, with a
  new mark in its last run's formatting (inserted, tracked), and its old `w:pPr` goes on an empty
  paragraph after the table. Anywhere else an insert leaves a cell ending in a table, docx4j's
  repair (an empty paragraph after any wrapper, plain under tracking); the cells of what goes in
  are repaired before it goes in. A file that already breaks the rule is left as read until an
  edit touches that cell; nothing is repaired on load. No reader moves.
- **`Font` reads every run in scope** (CR-016 Phase B, section
  [11](docs/change-requests/CR-016-cell-ends-fonts-placeholders-and-wildcards.md#11-phase-b-implementation-notes-2026-10-24);
  core-ts CR-002 section 59, held to Word 365 by Office JS checks 52 and 53): `Paragraph.font` and
  `Range.font` within one paragraph report each run's effective formatting, `None` where the runs
  differ (`""` for `name`, `color`, `highlight_color` and `style`, `"Mixed"` for `underline`), where
  0.1.0 reported the first run's, so `font.size + 2` can raise; the getters are typed `bool | None`
  and `float | None`. A paragraph's font reads the runs of its text (`get_range("Content").runs`: no
  deleted run, no hidden run, no field-code run), else a placeholder's runs, else its mark's run
  properties (`w:pPr/w:rPr`), where it read the first run or the paragraph style; a range holding
  its mark and no run reads the mark; a collapsed range its paragraph's style, as before. A font
  write's `ChangeReport` gives the runs' direct value before the write, `None` where they differ.
  A range across paragraphs (`MixedFont`) reads as before. Over the 418 documents of
  `out/cr012-b-corpus/`, 680 of 85,407 paragraph and cell font answers move to `None`, `""` or
  `"Mixed"`, 68 to an empty paragraph's mark and 15 to the text's runs (CR-016 section 11.4). The
  python-docx facade's `Font` keeps python-docx's semantics.
- **An automatic colour reads `#000000`** (CR-016 Phase B; the user's decision of 2026-10-24, CR-009
  section 12.9): the effective `Font.color` of a run whose colour is `w:val="auto"`, unstated or a
  value Word rejects reads `#000000`, as Office JS reads it, where 0.1.0 read `"auto"` or `""`;
  `get_font(direct=True).color` keeps `"auto"` and `""`. **`Font.to_dict()["color"]` follows the
  effective reading**, so it is `#000000` for most paragraphs, and `repr()` shows it. It is the
  API's answer, not the drawn pixel: Word draws automatic text white over a dark shading. 9,349
  corpus font answers move.
- **A control showing its placeholder reads as no text** (CR-016 Phase B; core-ts CR-012 section
  18, held to Word 365 by check 53 cases 34 and 35 and check 58 cases 11 and 12): the runs of a
  content control whose `w:showingPlcHdr` is on, at any level, are left out of `Paragraph.text`,
  `Body.text`, `get_text` in both views, `ContentControl.text`, a range's text and offsets,
  `TableCell.value`, `Table.values` and a comment's content, as hidden text is, and `search` and
  `find` find nothing in them (K3: Word finds hits reading `""`). `ContentControl.placeholder_text`
  reads the glossary entry since CR-019 (below). The text model is
  unchanged: it reads a placeholder as text, flagged, as docx4j does; its runs there have no
  `office_offset`, and the runs after them stand where `Paragraph.text` has them. In the corpus
  14 documents move.
- **Every text write into a control showing its placeholder writes in its place, as Word 365 does**
  (CR-016 Phase B for `ContentControl.insert_text`, check 58 cases 20 to 22;
  [CR-019](docs/change-requests/CR-019-writing-into-a-placeholder.md) Phases A and A2, sections
  [10](docs/change-requests/CR-019-writing-into-a-placeholder.md#10-phase-a-implementation-notes-2026-10-25)
  and [11.4](docs/change-requests/CR-019-writing-into-a-placeholder.md#114-phase-a2-implementation-notes-2026-10-25),
  for every route; the port of docx4j-core-ts CR-002 section 63 at `020c039` and `d84f3da`, held to
  Word 365 by core-ts's checks 63 (run 2, Design Mode off), 64 and 67): the control's `insert_text`
  at any location, its `"Content"` range's `insert_text`, and a block control's paragraph's
  `insert_text`, `text = ...` and `splice` take the placeholder's runs out, clear `w:showingPlcHdr`
  and write the text in one paragraph, in either form (a `w:placeholder` naming a glossary entry,
  Word's, or none, the bare form docx4j's binder writes). The text takes a copy of the control's own
  `w:sdtPr/w:rPr` and nothing of the placeholder run's, plain where the control has none (D11,
  core-ts check 68 and `c275a2a`). Tracked, only the insertion is recorded. The control's
  `insert_paragraph("Start" | "End")` empties the placeholder's paragraph, clears the flag and puts
  the new paragraph after it. 0.1.0 left the flag on in every route, and wrote beside the prompt at
  `"Start"` and `"End"` (under "Fixed").
- **A Replace or delete over a whole control showing its placeholder takes it** (CR-019 D3, check 63
  cases 12, 18 to 20, 26, 38, 44 to 46 and 52): `paragraph.text = ...`, a paragraph's
  `insert_text(..., "Replace")` or `"Content"` delete holding a run-level placeholder control, that
  control's `"Whole"` Replace or delete, and a paragraph's `"Content"` delete over a block control's
  only paragraph. Untracked the control goes, a block control leaving its paragraph emptied;
  tracked it is marked deleted between Word's `w:customXmlDelRange` markers, nothing inside it
  marked and its flag cleared, a Replace's insertion first, and the revisions engine resolves it as
  Word does. 0.1.0's `paragraph.text = ...` took the control too, reading its prompt as text; its
  span edits kept the control, emptied, as they still do for a run-level control holding text.
  With it **a tracked `Paragraph.insert_text(..., "Replace")` over an equation goes ahead**, the
  equation deleted in Word's form, where CR-013 kept that one refusal (`tracking.equation`; core-ts
  dropped it at `020c039`).
- **A content control's own ranges are inside it** (CR-019 D4 and Q4; core-ts's F2, check 64 cases 21
  and 42): a run-level control's `get_range("Content")` writes inside the control at every
  location, at `"Start"` with text showing too, where 0.1.0 wrote into the run before it; an empty
  run-level control's range stands where the control does, not at offset 0. A range made from one
  (`expand_to`, `get_range(...)`) is an ordinary range.
- **A paragraph inserted beside a block control's only paragraph goes outside the control** (CR-019
  D6 (b), check 63 cases 10, 11, 36 and 37): `Paragraph.insert_paragraph(..., "Before" | "After")`
  on the only paragraph of a block control showing its placeholder puts the new paragraph beside
  the control, which stays block-level (Word's save makes it run-level, which no reader here
  needs). 0.1.0 put it inside.
- **Every write into a bound control writes its node** (CR-019 D7; checks 63's run 2 and 64, where
  Word's writes reached the node in both forms): `Paragraph.splice`, and so every text writer,
  writes the custom XML node of each bound control around the span, after the edit and outside the
  tracked branch, where only a run-level control's `insert_text` did (`write_through`); a block
  control's `insert_text` writes it too. About 0.1 ms more a write into a bound control (CR-019
  section 10.8).
- **A `"Content"` delete of a bare-form control showing its placeholder shows Word's filler**
  (CR-019 D10 (a), check 63's run 2 cases 6, 17, 32 and 43; Word's filler of check 52's cases 13 and
  26): five spaces in one run, in a copy of the control's `w:sdtPr/w:rPr`, under the flag, tracked
  or not, nothing marked; so does an untracked `"Content"` delete that empties a bare-form control
  holding text. The Office JS readers read it as no text, the text model as the placeholder run it
  is. In Word's form the delete changes nothing, and a control holding text, emptied untracked,
  shows its entry again (check 66 case 3). 0.1.0 left a run-level control an empty `w:sdtContent`
  under the flag (under "Fixed").
- **`ContentControl.placeholder_text` reads the glossary entry, and is `str | None`** (CR-019 Q2,
  section
  [12](docs/change-requests/CR-019-writing-into-a-placeholder.md#12-phase-b-implementation-notes-2026-10-25);
  core-ts `aefc41f`'s D8, held to Word 365 by check 65): the text of the entry `w:placeholder`
  names, its paragraphs `\r` apart, the glossary read privately so it is saved byte for byte; None
  in the bare form or where the glossary lacks the entry, as Office JS reads it. 0.1.0 read the
  control's runs while the flag was on, else `""`, so a caller testing for `""`, or reading docx4j's
  prompt from a bound control, now reads None. **Setting it names an entry**: the kind's default
  text, or `""`, names Word's default entry for the kind, any other text a new entry of 32
  upper-case hex digits, written into the glossary on save; a control showing its placeholder, or
  holding nothing, shows the text in `PlaceholderText` with the flag, and one holding content keeps
  it. **`BindingError` `binding.holds_content`, which 0.1.0's setter raised over content, is gone**:
  the setter no longer refuses.
- **The binding pass gives an empty value Word's placeholder** (CR-019 Q3; core-ts `aefc41f` and
  `b12e0c5`, docx4j CR-063's D8, from the user's Word 365 file of one control of each kind,
  `tests/office_js/placeholders-word365.docx`): `apply_bindings` over an empty node shows the
  control's placeholder in Word's form (`show_placeholder()`), the entry it names where the glossary
  holds it, else Word's default entry for its kind: `DefaultPlaceholder_-1854013440` "Click or tap
  here to enter text." for rich and plain text, `-1854013437` "Click or tap to enter a date." for a
  date picker, `-1854013438` "Choose an item." for a drop-down list or a combo box, and a building
  block gallery's and a repeating section item's. A kind with none (a picture, a check box) keeps
  docx4j's run, whose words are Word 365's now: `bindings.PLACEHOLDER_TEXT` is "Click or tap here to
  enter text." (docx4j's D5; it was "Click here to enter text."). 0.1.0 wrote docx4j's run, in the
  bare form, for every kind; docx4j's binder still does, until its CR-063 lands.
- **A control made on an empty paragraph, or at a point, shows its kind's placeholder** (CR-019 D8,
  check 64 cases 15 to 21 and 36 to 42): `Paragraph.insert_content_control` on a paragraph holding
  nothing, and `Range.insert_content_control` on a collapsed range, make a rich text, plain text,
  date picker, drop-down list or combo box control showing Word's default entry for it, as Word's
  made controls do. `Body.insert_content_control` and a range across paragraphs make controls as
  before. 0.1.0 made an empty control on an empty paragraph, and wrapped the run at a point (under
  "Fixed").
- **A file that comes to show a placeholder gains a glossary** (CR-019 D9 (a), check 66; section 7's
  risk 3): where a control comes to name an entry the file lacks, the save writes the entry into the
  glossary in Word's form, making the glossary document part (`/word/glossary/document.xml`, its
  relationship from the main document part and its content type) where the file has none, and the
  glossary's styles part, with Default Paragraph Font and Placeholder Text, where those are missing
  (`package.last_save.added` lists the parts made). The main document's styles part gains Placeholder
  Text, Word 365's definition, where a placeholder is shown and it lacks the style, as a hyperlink's
  style is added (core-ts's `requireStyles`; CR-019 section 12.5 item 3). Entry names and `w:guid`s
  come from the package's seeded generator. An untouched glossary is saved byte for byte, and a file
  that shows no new placeholder is unchanged. Acceptance artefact 9's control, made at a point,
  gives that file a glossary.
- **A write at a paragraph's end goes after everything the paragraph holds, as Word 365 writes it**
  (CR-019 Phase D, section
  [15](docs/change-requests/CR-019-writing-into-a-placeholder.md#15-phase-d-implementation-notes-2026-10-25);
  the port of docx4j-core-ts CR-002 sections 63.24 to 63.27 at `09dc3d2`, held to Word 365 by
  core-ts's checks 69 and 70, the user's D12 to D14): `Paragraph.insert_text(..., "End")`, a
  paragraph's `get_range("End")` at any location and its `get_range("Content")` at `"End"` and
  `"After"`, and through them `Body.insert_text(..., "End")` and a cell's, write after the
  paragraph's last item, into none of the items: not a trailing run-level content control (one
  showing its placeholder keeps its prompt and flag), a hyperlink, a tracked insertion or deletion,
  a field, a bookmark's or a comment's markers, or a text box (D12). The text takes the last item's
  run properties only where that item is a run holding text, extending it, so after a hidden run it
  is hidden and the range returned reads `""`; after anything else it is a run with no properties,
  whatever runs are around a control (D13; Word's `w:noProof` after a text box is left out). The
  End range's `"After"`, and the Content range's, write a run with no properties (D14). Tracked, the
  text is a `w:ins` after everything, or joins this author's own insertion that ends the paragraph;
  a paragraph break in it makes paragraphs, written visible after a hidden run. A search hit or a
  `"Whole"` range ending at the paragraph's end writes as before. 0.1.0 extended the last run
  holding text: into a trailing control, a field's result or another author's insertion, inside a
  bookmark or a comment's range, before a point bookmark, a deletion, a text box or a picture
  (under "Fixed"), and into a trailing hidden run, hidden, as Word does, which CR-013 (below, never
  released) had made visible.
- **A deleted paragraph mark reads as nothing in Word's default view** (CR-016 Phase B; core-ts
  CR-012 section 19, held to Word 365 by check 58 cases 6 to 10): `Body.text`, `get_text()`, a
  range's text across paragraphs, a `"Whole"` range's text and a block control's, a cell's and a
  comment's text write no `\r` after a paragraph whose mark is tracked as deleted, and
  `get_text(view="original")` none after an inserted one, the marks each view's resolution keeps;
  a cell's last paragraph keeps the cell's `\t`. `Paragraph.mark_text` is `""` for a deleted mark,
  so `search("^p")` no longer matches one. The paragraphs, their offsets and `body.paragraphs` do
  not change, nor does `TrackedChange.text`. In the corpus `get_text()` moves in the 78 documents
  holding a deleted mark and the original view in the 81 holding an inserted one.
- **Word's wildcards match a paragraph mark where `^p` does** (CR-016 Phase C, section
  [12](docs/change-requests/CR-016-cell-ends-fonts-placeholders-and-wildcards.md#12-phase-c-implementation-notes-2026-10-24);
  core-ts CR-002 section 32's K7, held to Word 365 by Office JS check 59): with
  `match_wildcards=True`, `?`, `*` and a class cross a paragraph's own mark and `^13` is one, never
  a cell's or a row's end, a deleted mark, nor anything across a table's edge. `Paragraph.search`
  searches its `"Whole"` range, its mark included (`?` over `Q` finds `Q` and `\r`);
  `Body.search`, and so a block control's and a cell's body's, its `"Content"` range, without the
  last mark; a range across paragraphs, or holding its mark, its paragraphs' marks, a hit ending
  with one holding it. Before, no wildcard matched a mark and `^13` was three characters; inside a
  paragraph it now matches a `w:cr`, and a wildcard pattern is `re.DOTALL`, so `?` takes a line
  feed held in a `w:t`. `find` still reads a paragraph at a time. In the corpus no reader of the
  audit moves; its new wildcard reader finds 4,221 hits in 288 documents, each holding a mark.
- **`replace_text` with wildcards refuses a hit holding a mark** (CR-016 Phase C, Q9):
  `SpanError` `range.replaces_mark`, before anything changes, where any hit `search` finds holds a
  paragraph's mark (`body.replace_text("Q?Z", "X", match_wildcards=True)` over the paragraphs `Q`
  and `Z`; a paragraph's `replace_text("?", ...)`, whose hits now include its mark), as a plain
  `^p` query is refused. A body's replacement reads its `"Content"` range, so over a body of one
  paragraph `?` is replaced where that paragraph's own `replace_text("?", ...)` refuses (core-ts's
  rule). 0.1.0 replaced each paragraph's hits. So a query whose `?`, `*` or negated class can reach
  a paragraph's end is now refused where 0.1.0 replaced: `e?` over a paragraph ending in `e`,
  `[!a-z]` over any two paragraphs. Search for the hits and edit their ranges instead.
- **No search hit cuts a grapheme cluster** (CR-016 Phase C, Q8; core-ts `ebb218e`): `search`,
  `find` and `replace_text`, plain or wildcard, drop a hit whose start or end falls inside a
  cluster (`is_grapheme_boundary`): `cafe` no longer matches inside `cafe` + U+0301, nor `?` an
  Arabic letter without its fatha, so `?` and `[!a-z ]` over an Arabic word find what Word's do.
  Text with no code point at or above U+0300 is never filtered, and pays nothing. In the corpus no
  plain search moves; `?` and `[!a-z ]` lose 4 hits, in one Arabic phrase. A class still matches
  an emoji whole (a code point), where Word's matches nothing (CR-016 section 4.9).
- **A save writes Word's reading of a kept value, in a plain normal form** (CR-018 section 4.7,
  Q5 and Q17, the user: "they are entitled to assume we don't write dodgy content"): an on/off
  value the parse kept as its text is written bare where Word reads it as on (`"1"` where the
  attribute's absence is off), and `0` where Word reads it as off. Only the parts a save
  re-marshals, only kept values, each listed in `package.last_save.normalised`; the tree keeps
  what was read. `package.save(normalise=False)`, or `LoadOptions(normalise_on_save=False)`,
  keeps the source's spelling. `scripts/canon.py` tolerates a second category,
  `word-normal-form`, by writing the same normal form into the source before it compares, and
  `scripts/roundtrip.py` runs both ways (`--normalise`, `--no-normalise`). Phase B adds the
  enumerations' and colours' rows, found in each part's render so the tree is never touched: a
  colour Word reads is written `RRGGBB` (`red`, `ff0000`, `#FF0000`); a text or underline colour
  it rejects is dropped (`auto` beside a theme colour); a fill or border colour it rejects is
  written `auto`; a `w:highlight`, `w:vertAlign` or `w:em` with a missing or invalid value, and a
  bare `<w:u/>`, are dropped, as Word's repair and save drop them; `w:u`'s on/off words are
  written `none`; a `w:background` whose colour Word rejects is dropped (no page colour). A colour
  or an enumeration set in code is written in the form too.
- **`Font.color` reads a colour as Word draws it** (CR-018 Phase B): a classic HTML name or any
  hex spelling as `"#RRGGBB"`, and `""` for a value Word rejects (`Auto`, `F00`, `grey`), where
  it returned the string; `Font.underline` reads an invalid `w:u` as `"None"` (it read
  `"Mixed"`); setting `Font.color = "red"` writes `FF0000`. The facade's `color.type` is None for
  a rejected value, as for no `w:color`.
- **Booleans and enumerations are parsed exactly as written** (CR-018 section 4.1a): every parse
  is `runtime.parser_config()`'s, which sets the fork's `exact_types={bool, Enum}`
  (`runtime.EXACT_TYPES`), so `" false"` and `" yellow"` are kept as their text instead of
  trimmed; `tests/test_parser_configs.py` holds every `ParserConfig` in the tree to it.
- **`wml.xsd` refreshed from docx4j `46c727bf8`** (CR-018 section 4.5): docx4j's #697 and CR-059
  corrected the on/off attribute defaults (36 field defaults `true` to `false`, 10 gone, in the
  `schema_default` metadata only; `schemas/PATCHES.md`).

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
- **`Paragraph.text`, `get_text`, `search` and `find` read the branch Word shows of a run- or
  paragraph-level `mc:AlternateContent`, as Office JS does**
  ([CR-011](docs/change-requests/CR-011-one-text-model.md) Phase F, sections
  [15](docs/change-requests/CR-011-one-text-model.md#15-phase-f-the-office-js-readers-read-words-branch-q4-revised-2026-10-22-implemented-2026-10-22-section-16)
  and [16](docs/change-requests/CR-011-one-text-model.md#16-phase-f-implementation-notes-2026-10-22);
  Q4 revised 2026-10-22). Word 365's `paragraph.text` gives the selected branch's words, never a
  Fallback's and never nothing, and `body.search` finds a word only a Choice holds and none only a
  Fallback holds (docx4j-core-ts check 49). 0.1.0 passed the element by: `"I prefer "` where Word
  shows `I prefer w14`, and nothing for a paragraph whose text is all in a branch. The branch is
  the one `text_of`, the text model and docx4j's TEXT purpose take, so `Paragraph.text`,
  `Body.text`, `Range.text`, a block control's text, `get_text` in `accepted` and `original`,
  `search`, `find` and `replace_text`'s matches now agree with a cell's text and with the text
  model; a branch's insertions and deletions are filtered as the paragraph's own, and a text box
  in a branch still gives nothing (check 42). A break reads U+000B, as Office JS gives it, since
  CR-013 Phase A (below); until then it read `\n`.
  **An edit through that text is refused** with `address.read_only` before anything changes (an
  edit there would change Word's branch alone): an insertion strictly inside it, and any span
  holding it --- `insert_text`, `delete`, `font` and `style` writes through a `Range`, a
  paragraph's `font` writes, `insert_content_control`, `insert_comment`, `hyperlink` and
  `insert_ooxml` through a range, `replace_text` (every match checked first) and, with change
  tracking on, `Paragraph.text = ...`, `delete`, `Body.clear` and a row's or a table's deletion.
  An insertion at either end of the element's text goes beside the element, and an edit
  elsewhere in the paragraph works as before. **Offsets in such a paragraph count the branch's
  text**, so a stored `Range` or `SearchHit` offset after the element moves by its length, and
  each text model run read from a branch has its `office_offset`. Text model runs either side of
  a branch's boundary with the same formatting are now one run in `TextBlock.runs`, as docx4j
  writes them; their text is unchanged. The python-docx facade's readers are unchanged, and its
  `add_comment` on a run whose only text is a branch's is refused, where 0.1.0 anchored the
  comment at the paragraph's start.
- **The Office JS readers read Office JS's characters, and every text writer takes them back**
  ([CR-013](docs/change-requests/CR-013-ranges-across-paragraphs.md) Phase A, section
  [10](docs/change-requests/CR-013-ranges-across-paragraphs.md#10-phase-a-implementation-notes-2026-10-23);
  decided 2026-10-23, Q1 and Q2). The text is Word's own character stream, as docx4j-core-ts
  measured it in Word 365 (its checks 48, 48b and 51) and as `tests/office_js/` holds it to Word's
  answers paragraph for paragraph:
  - **`Body.text`** (and `get_text`, `text_budget`) ends each paragraph with `\r`, the last one's
    left out, and writes a table as Word does: a row's cells a `\t` apart, a cell's last mark not
    written, each row ended `\r\n`. 0.1.0 wrote a paragraph per line. `body.text.split("\n")` now
    gives one string; split paragraphs with `split("\r")`, not `splitlines()`, which also splits
    at U+000B, U+000C and U+001E.
  - **A cell's `value`** (so `TableRow.values` and `Table.values`), **a block control's `text`** and
    **a comment's `content`** are their content's `Body.text`, paragraphs `\r` apart; a run-level
    control's `text` is its span of `Paragraph.text`. `Table.text` and `Block.text` stay docx4j's,
    a line per paragraph, so `Table.text` and `Table.values` part company, as in core-ts.
  - **In a paragraph** (`Paragraph.text`, `get_text`, `Range.text`, `search`, `find`, `SearchHit`'s
    context, `paragraph_at(contains=)`, `to_dict`, a `ChangeReport`'s texts): a line break is
    U+000B, a page break U+000C, a column break U+000E, a `w:cr` U+000D, a `w:ptab` `0`, the
    non-breaking and optional hyphens U+001E and U+001F, any `w:sym` `(`, a footnote or endnote
    reference and a note's own mark U+0002, and an equation Word's text of it (Unicode's
    mathematical alphabets, U+2212 for a minus, each structure `,`-args-`.`). **Hidden text is left
    out**, a run's own `w:vanish` or its styles', so a hidden word is not found and
    `paragraph.text = paragraph.text` drops the hidden runs. `paragraph.text == "a\nb"` is false
    where it was true for a break.
  - **A paragraph's `'Whole'` range** (`get_range()`'s default) ends with its mark, `\r`, or `\t`
    for a cell's last paragraph; `'Content'` does not.
  - **`search`** takes Word's Find codes (`^f ^e ^l ^m ^t ^~ ^- ^g ^^`, so a literal `^` is `^^`),
    matches a picture's place with `^g` and not across it otherwise, passes over an optional hyphen
    unless the query names it, and lets a plain `-` match a non-breaking hyphen.
  - **Stored offsets move** in a paragraph holding a note reference, a `w:ptab`, hidden text or an
    equation; addresses do not. The text model's `office_offset` agrees with `Paragraph.text` under
    the character map (Q11), its segmentation unchanged.
  - **The writers** (`Paragraph.text = ...`, `insert_text`, `replace_text`, `insert_paragraph`,
    `Body.insert_text`, `TableCell.value = ...`, `values`, `insert_table` and `add_rows` values,
    `insert_comment`, `Comment.content = ...`, `reply`, a content control's `insert_text` and
    `insert_paragraph`) follow Word's own writers, as docx4j-core-ts's Word checks 55 and 55b
    measured them (CR-013 section
    [11](docs/change-requests/CR-013-ranges-across-paragraphs.md#11-phase-as-amendment-words-own-writers-decided-2026-10-23),
    decided 2026-10-23). U+000B is a `w:br`, U+000C and U+000E a page and a column break, U+001E
    and U+001F the hyphens and `\t` a `w:tab`, so a text read and written back reads as it did.
    0.1.0's `Paragraph.text = ...`, `insert_text` and `replace_text` wrote `\t` and `\n` raw into
    the `w:t` (the paragraph, cell, table and comment writers mapped them), and the six that are
    not XML characters failed only at the save.
    - **`insert_text` and `replace_text` make a new paragraph of each `\n` and each `\r`, `\r\n`
      one**, as Word's `insertText` does: on a paragraph, a range, the body, a cell and a block
      content control. Both paragraphs keep the paragraph's properties; the new one goes in front
      with a mark of its own, and the `w:p` keeps the last piece, its mark, its section break and
      its `w14:paraId`, so a `Paragraph` view goes on reading the last of them. The returned `Range`
      spans the paragraphs the text made, as Office JS's does (`x\ry`; CR-013 Phase B, below). With
      change tracking on, the new mark is marked inserted. A hyperlink,
      a tracked change or a smart tag the break falls in is cut in two; a text that would go into a
      run-level content control or a field, or a break inside an equation or text read from a kept
      `mc:AlternateContent`, is refused before anything changes (`SpanError`
      `range.splits_control`, `range.splits_field`, `range.cuts_equation`; `address.read_only`).
      Inside a `w:text` content control, bound or not, block or run, and a bound block control of
      no type or `w:richText` (Word can't open one holding a second block, and reads every one as
      plain text), a break makes no paragraph: it is a `w:br` where the control is a multi-line
      `w:text`, and nothing in any other, where a line, page or column break is dropped too and a
      bound control's node follows (docx4j-core-ts CR-002 sections 60 and 61, its Word checks 55b
      and 56). In a run-level control a `\n` or `\r` is refused (`range.splits_control`). 0.1.0
      wrote `\n` and `\r` raw into the `w:t`.
    - **Text written into a paragraph whose text is hidden** takes no hidden run's look, so it reads
      as text and its characters reach the map, and `"Start"` puts it before the leading hidden runs.
      At the paragraph's end, since CR-019 Phase D (above), text after a trailing hidden run is in
      that run's look, hidden, as Word 365 writes it, and as 0.1.0 wrote it into the hidden run's own
      `w:t`; its characters reach the map, and a paragraph break in it is written visible.
    - **A bound control's custom XML node** takes a non-breaking hyphen as `-`, as Word writes it;
      an optional hyphen stays U+00AD.
    - **`insert_paragraph` and `Paragraph.text = ...` write one paragraph**: `\n` and `\r\n` a
      `w:br`, `\r` a `w:cr`; so do a cell's `value`, a table's values and a comment's paragraphs.
      Word's `insertParagraph` writes the escapes as text instead, which this does not follow.
    - **U+0002 and the other characters XML 1.0 excludes are dropped**, with no error, as Word
      drops them; lxml's refusal at save stays behind that. So `paragraph.text = paragraph.text`
      on a paragraph holding a note reference loses the reference.
    - **In a comment** `\r`, `\n` and `\r\n` (one) are paragraph marks, where 0.1.0 split on `\n`
      and saved a `\r` as `&#13;`.
    - **New**: `get_content_controls(types=...)` on `Body`, `Paragraph`, `Range` and
      `ContentControl` (Office JS `getContentControls`, WordApi 1.5), filtering by a control's
      `type` as Word does; `content_controls` keeps listing every kind, where Word's lists rich-text
      controls only. Extensions: `Paragraph.refuse_paragraph_breaks`, `text_since` and
      `holds_objects`; `controls.is_bound_block_control`, `plain_text_lines_of`, `block_control_of`
      and `paragraph_lines_of`; `bindings.bound_value_of`; `text_model.drop_unwritable`,
      `for_writer`, `PARAGRAPH_BREAKS` and `LINE_CHARACTERS`; and `model/content/split.py`'s
      `split_paragraph` (core-ts's `splitParagraph`). `to_api_script` writes a run's line break as
      U+000B in an `insert_text` line.
  - **A deletion takes the pictures in its span**, as docx4j-core-ts's CR-002 section 48 has it
    (CR-013 section 11): a picture strictly inside a deleted or replaced span goes with it, and a
    paragraph's whole content holds those at its edges too (`insert_text("Replace")`,
    `Paragraph.text = ...`, a `'Whole'` range, a tracked `Paragraph.delete()`, which now takes a
    paragraph holding only a picture); with change tracking on, the picture's run goes into the
    `w:del`. A search hit beside a picture leaves it. `Range.get_xml` trims the pictures outside
    its span. 0.1.0 left every picture in place. The image part stays in the package, as it does
    when `InlinePicture.delete()` removes one.
  - **Equations in an edit**: an edit takes an equation whole, untracked, and never part of it
    (`SpanError` `range.cuts_equation`); with change tracking on, a deletion that would take one ---
    `Paragraph.delete()`, `Body.clear()`, a row's or a table's deletion, a replacement --- is
    refused with `TrackedChangeError` `tracking.equation`, as core-ts refuses it. 0.1.0's tracked
    deletion left the equation undeleted.
  - **Unchanged**: `traversal.text_of`, `TrackedChange.text` (docx4j's characters, as the
    revisions goldens hold them), the text model with its plain text and Markdown, `outline()` and
    `describe()`; and a bound content control's write-back into its custom XML, which keeps
    docx4j's characters (a break `\n`, paragraphs a `\n` apart). **The python-docx facade's
    readers are python-docx's, unchanged**, and its writers follow python-docx 1.2.0 as measured:
    a carriage return in `add_run`, `Run.text = ...`, `Paragraph.text = ...` and `cell.text = ...`
    is a `w:br` (0.1.0 wrote it raw), a comment's `\n` starts a paragraph and its `\r` is a `w:br`,
    any other control character is refused with `ValueError`, and `Comment.text` stays
    python-docx's, paragraphs a `\n` apart.
  - **`dry_run`** copies a part the real package has not read from its private reading, so a
    trial's first look at it leaves it unread: every `Paragraph.text` now asks the styles which
    runs they hide.
- **Ranges across paragraphs: the reads** ([CR-013](docs/change-requests/CR-013-ranges-across-paragraphs.md)
  Phase B, section [12](docs/change-requests/CR-013-ranges-across-paragraphs.md#12-phase-b-implementation-notes-2026-10-23);
  the port of docx4j-core-ts CR-014 phase 1, `ef57729`). A `Range` has two ends: `paragraph` and
  `start`, and `end_paragraph` holding `end`, the same paragraph for a range within one, which
  reads as it did. Held to Word 365 by core-ts's checks 52 and 53, run here through a Python port
  of their Script Lab snippets (`tests/office_js/checks.py`) on the files Word read and Word's
  saves:
  - **Across paragraphs** a range's `text` is `Body.text`'s form between its ends (a mark `\r`, a
    table's cells `\t` apart and its rows ended `\r\n`); `paragraphs` is every paragraph it
    touches, a table's cells' among them; `font` is a `MixedFont`, a property where the runs
    differ `None`, `""` for the name, the colour and the highlight, `"Mixed"` for the underline;
    `search` and `find` go a paragraph at a time; `get_tracked_changes()` lists the story's
    changes it overlaps; `hyperlink` reads the first link among its paragraphs.
  - **New members** (the Office JS subset, `tests/office_js_subset.json`, is 247 members, derived
    from core-ts's committed `test/office-js-subset.ts`): `ContentControl.showing_placeholder_text`
    (CR-003 section 24; `is_showing_placeholder` now reads `w:showingPlcHdr w:val="0"` as off);
    `Range.expand_to`, `is_empty`, `tables`
    and `content_controls`; `Table.get_range` (`'Whole'` ending `\r\n`, `'Start'`, `'End'`,
    `'After'`); `Paragraph.table_nesting_level`; `Range.get_range("After")`. Extensions:
    `Range.end_paragraph`, `spans_paragraphs`, `holds_end_mark`, `whole`, `end_mark`,
    `parent_table_cell`, `parent_table` and `require_one_paragraph`; `model/content/story.py`'s
    `story_of`; `RangeLocationName`, `TableRangeLocation` and `BodyRangeLocation`.
  - **The views give one range**: a block, row or cell content control's `get_range()` runs from its
    first paragraph to its last (0.1.0's was the first paragraph's, so `get_range().delete()`
    emptied that one paragraph); `Body.get_range()` from the body's first paragraph to its last
    (0.1.0's was the first paragraph's); a tracked change's from its first piece to its last, a
    deletion collapsed where it starts, an inserted mark `\r`, a deleted one collapsed at its
    paragraph's end (0.1.0's was its part in its first paragraph).
  - **`Comment.get_range()` is one `Range`, or None** with no markers in the body, as Office JS's
    (`BODY1B\rLIST1A` across paragraphs; D2). **Breaking for 0.x**: 0.1.0 returned a list, a range
    per paragraph, which is `Comment.get_ranges()` now; `comment.get_range()[0]` becomes
    `comment.get_range()`, or `comment.get_ranges()[0]`.
  - **`Body.get_range()` on a body with no paragraph** raises `ContentError` `body.no_paragraph`,
    where 0.1.0 inserted a paragraph (a read that wrote); on a notes or comments part's body, whose
    first and last paragraphs are in two notes, `SpanError` `range.other_story`.
  - **`Range.style` and `style_built_in`** read Word's rule, within one paragraph too: a range
    holding its paragraphs whole reads their shared paragraph style, else `Normal`; any other range
    its runs' shared character style (`Emphasis`), `""` (`Other`) where they differ. 0.1.0 read the
    paragraph's style. A paragraph after a nested table in a cell reads its own style, where Word
    reads `Other` (core-ts CR-002 section 57's departure, taken here).
  - **`Range.get_comments()`** reads each comment's extent over the story, so a range on any of a
    comment's paragraphs lists it; 0.1.0 read only the range's own paragraph's markers, so a range
    on a comment's last paragraph listed nothing.
  - **A row content control's `text`** is its cells a `\t` apart, the row end left out, and its
    `'Whole'` range ends `\t\t`, as Word's; a block control whose content ends with a table leaves
    the row end out too. A content control's level (`form`) is read past the markers its content
    opens with (a comment's, a bookmark's): Word's save of check 52 opens a block control with a
    `w:commentRangeStart`, which 0.1.0 read as a run-level control.
  - **A deletion or replacement starting or ending inside a table** is refused with `SpanError`
    `range.partial_table`, as Word refuses it, before anything changes. A `Range` made across
    stories, or with its end before its start, is refused (`range.other_story`, `range.order`). The
    writes across paragraphs are Phase C's (below).
- **Ranges across paragraphs: the writes** ([CR-013](docs/change-requests/CR-013-ranges-across-paragraphs.md)
  Phase C, section [14](docs/change-requests/CR-013-ranges-across-paragraphs.md#14-phase-c-implementation-notes-2026-10-23);
  the port of docx4j-core-ts CR-014 phases 2 and 3, `53db5bc`, `87bbb38` and `202a01d`). Every write
  of a `Range` works across paragraphs, untracked and tracked, as Word 365 makes it in core-ts's checks
  52 and 54 (run here through `tests/office_js/checks.py`), with the range a write leaves as Office
  JS's, and one `ChangeReport` a write, its `moved` each shifted block's ordinal before the call
  against its ordinal after, however many blocks the write removed:
  - **`delete()` and `insert_text(..., "Replace")`** join the paragraphs in the last one's
    properties; a paragraph mark the range holds goes, whatever ends the range, its emptied paragraph
    with it where no paragraph follows (a cell's last and a story's last kept); into or out of a block
    content control nothing joins, and the start's paragraph takes the end's properties. Tracked, the
    text is a `w:del` and every mark the range holds is deleted, with no `w:pPrChange`; a whole content
    control goes between Word's `w:customXmlDelRangeStart` and `End` markers, which the revisions engine
    lists as `Deleted` and resolves as Word does. An untracked replacement takes the first replaced
    character's look. Replacing a range that ends with a whole table is refused
    (`range.replaces_table`), as Word refuses it.
  - **`font`** writes each property only where a run or a held paragraph mark doesn't read it
    already, a mark's change tracked as Word's `w:rPrChange`; **`style`**, **`style_id`** and
    **`style_built_in`** set each paragraph's, tracked with a `w:pPrChange` each.
  - **`insert_paragraph`** splits the paragraph at the range's start (`"Before"`) or end (`"After"`)
    for every range, within one paragraph too (see below); `insert_text` at `"Start"` and `"End"`;
    `replace_text` a paragraph at a time, every match checked first; `insert_ooxml`, one paragraph's
    runs in place of the range.
  - **`insert_comment`** across paragraphs and across a table: the start before the range's first
    character, the end and the reference after its last.
  - **`insert_content_control`** wraps paragraphs held whole in one block control (tracked, Word's
    `w:customXmlInsRange*` markers); any other range across paragraphs is refused
    (`range.partial_paragraph`), as Word refuses it.
  - **`^p` in `search`** matches a paragraph's mark (never a cell's or a row's end); `replace_text`
    with `^p` is refused (`range.replaces_mark`).
  - **`get_xml` and `to_api_script`** of a range across paragraphs give the blocks a deletion would
    take, trimmed at the ends.
  - **A deletion takes a comment's reference with it**, within one paragraph too: the comment goes,
    at once untracked and on Accept All tracked, as Word's does. What a deletion takes (a note, a
    comment, a relationship nothing references any more) is settled once, when the write ends.
  - Still refused: a hyperlink across paragraphs (`range.spans_paragraphs`; Word writes a field),
    several blocks in place of a range, `Replace` across a control's edge, and under tracking a range
    holding an `w:altChunk` or a kept `mc:AlternateContent` (`range.not_trackable`).
- **A tracked replacement is the insertion first, then the deletion** (CR-013 section 13.1 item 7,
  as Word writes it, core-ts check 54 case 4), within one paragraph too; 0.1.0 wrote the `w:del` first.
  `Range.insert_paragraph` splits its paragraph for every range, within one paragraph too (check 54
  cases 1 to 3), where 0.1.0 put the new paragraph before or after the whole paragraph;
  `Paragraph.insert_paragraph` is unchanged.
- **A paragraph break inside a field's result splits the field** (CR-013 section 13.7, D1; core-ts
  check 57 cases 1 to 4), its begin, instruction and separator in the first paragraph and its end in
  the next; a `w:fldSimple` is first written as Word's complex field. Until now any break inside a
  field was refused (`range.splits_field`), which stays for a field's instruction.
- **A tracked deletion taking an equation whole is made in Word's form** (D2, check 57 cases 7 to 10):
  a `w:del` in each `m:r` and each structure's `m:ctrlPr`; it was refused (`tracking.equation`),
  and since CR-019 a tracked `Paragraph.insert_text(..., "Replace")` taking one is not refused
  either (above). Accept All drops an
  `m:oMath` it empties and an `m:oMathPara` left with none, docx4j's rule since its `ebb21c990` (F1),
  Reject All of a whole inserted equation too; the readers leave out a structure whose own revision
  the view hides (F2).
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
- **Accept All where a paragraph mark taken away has no paragraph to join is Word 365's**
  ([CR-010](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md) sections 13 and
  [14](docs/change-requests/CR-010-tracked-changes-as-word-resolves-them.md#14-section-13s-implementation-notes-2026-10-23),
  the port of docx4j `29dbf5964` and `e55a83b8e`, its CR-043 section 17, from docx4j-core-ts's
  check 54 and docx4j's text-left probe): an emptied paragraph before a table or a block content
  control goes; with text left, before a table its content joins the front of the table's first
  cell's first paragraph, which keeps its own properties, and before a control it takes the
  control's first paragraph's properties; a control's last paragraph keeps a bare mark, text left or
  not. 0.1.0 dropped the mark and kept the paragraph as it was in each. A kept block-level
  `mc:AlternateContent` next is read as the branch Word takes, but takes no content. Whichever
  resolution takes the mark away, Accept All or Reject All, as docx4j's. The text model's
  `accepted` and `original` views follow, and also read a content control deleted whole (Word's
  `w:customXmlDelRange*` form) as gone when accepted and kept when rejected, and a table whose
  every row is deleted as gone.
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
  (`docx4j-xsdata` stage 10, in 26.2.2).
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
  `LoadOptions(mc_load_mode="word")` to edit Word's branch of a block-level element in place.
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
  - a `w:cellDel` accepted, or a `w:cellIns` rejected, in a branch Word does not read is removed
    from its own branch with the change it twins, and that branch's neighbour widened; it was
    listed as a twin but kept its cell and its marker (section
    [10.10](docs/change-requests/CR-012-type-kept-mc-branches.md#1010-the-cell-twin-port-2026-10-22),
    the port of docx4j `60bd4b9ad`);
  - a table's property record, a `w:tcPrChange` or a `w:trPrChange`, in a branch Word does not
    read is listed and resolved with the table's one change, and a grid restored by a reject is
    collapsed along each branch's reading of a row. Such a record was not listed when only that
    branch held one, and was kept after accepting or rejecting (section
    [10.11](docs/change-requests/CR-012-type-kept-mc-branches.md#1011-the-table-records-port-2026-10-22),
    the port of docx4j `901a922a3`);
  - a table's change's text is the rows of the branch Word reads, where it held every branch's;
    and a grid restored by a reject, over a row with a cell-level element, counts every branch's
    spans from the old ones before writing any, so a cell in Word's branch no longer loses its
    span beside a cell outside the element (section
    [10.12](docs/change-requests/CR-012-type-kept-mc-branches.md#1012-the-change-text-and-grid-port-2026-10-22),
    the port of docx4j `84b202c5a`);
  - `outline()` gives a cell's blocks under the table at their full paths through any wrapper, so
    a row-level content control's rows no longer come out as `tc` entries.
- **A kept `mc:AlternateContent` branch's WordprocessingML children are typed**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) Phase B, section
  [11](docs/change-requests/CR-012-type-kept-mc-branches.md#11-phase-b-implementation-notes-2026-10-22)):
  docx4j's 61 WordprocessingML root elements that this model does not declare globally
  (`w:drawing`, `w:pict`, `w:t`, `w:tbl`, `w:hyperlink`, `w:bookmarkStart` and the rest, derived from
  docx4j's own classes into `codegen/names/root_elements.json`) are typed in a branch as the
  container holding the element binds them, and in any other lax wildcard as their root class, as
  JAXB binds an `@XmlRootElement`. **Code that looked for an `AnyElement` `w:drawing` in
  `any_element` now finds a `Drawing`**, and a `Pict`, an `RT` or a `Tbl` where it found a
  wildcard `w:pict`, `w:t` or `w:tbl`; a DrawingML text box in a Choice is a typed `TxbxContent`,
  while below a `Pict` the VML stays a wildcard tree (Phase C). A branch's `w:moveFrom`,
  `w:moveTo`, `w:instrText` and `w:customXml` stay wildcard trees. The typed children are written
  back under their element names, with no `xsi:type`: over 2,746 parts of 422 documents every save
  is canonically what it was, and 25 parts, each holding a branch, differ in bytes, in attribute
  order or in a boolean written `0` for `false`. `LoadOptions(mce_preprocess=True)` is no longer
  the only way to a typed `Drawing` in a choice. With it:
  - the walks (`walk`, `walk_all`, `iter_nodes`, `find`, `text_of`, `child.iter_children`) give a
    branch's typed child its element name, where they gave the field's, `any_element`: `text_of`
    reads a branch's `w:t`, and a paragraph-level branch's `w:r`, as it reads them elsewhere, and a
    paragraph in a block-level branch as one line, where it read each run as a line;
  - `Body.tables` reads a table in Word's branch of a block-level element, read-only, and
    `Body.paragraphs`, `outline()`, `find` and `get_text` read its paragraphs;
  - the python-docx facade's `inline_shapes` keeps a `w:drawing` straight in a branch out, as
    python-docx's `//w:p/w:r/w:drawing/wp:inline` does;
  - the text model reads a typed `cx:chart` as it read a wildcard one (`waterfall.docx`'s, in its
    Choice).
- **VML is typed** ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) Phase C,
  section [12](docs/change-requests/CR-012-type-kept-mc-branches.md#12-phase-c-implementation-notes-2026-10-22)):
  docx4j's five VML schemas (`v:`, `o:`, `w10:`, Excel's `x:` and PowerPoint's `pvml:`, at the
  schema copy's commit) are in the model, as `docx4j_py.vml.main`, `.officedrawing`,
  `.wordprocessing_drawing`, `.spreadsheet_drawing` and `.presentation_drawing`, each with its
  `el`. VML is typed wherever it stands, in a branch or not: below every `w:pict` and `w:object`,
  a header's watermark, a numbering picture bullet, and `settings.xml`'s `o:shapedefaults` and
  `o:shapelayout`, which nearly every Word document has. **Code that looked for an `AnyElement`
  `v:shape`, `v:imagedata`, `o:OLEObject` or a VML text box's `w:txbxContent` now finds a
  `Shape` (`CTShape`), a `CTImageData`, an `Oleobject` (`CTOLEObject`) or a `TxbxContent`**;
  `v:imagedata`'s `r:id` is the field
  `schemas_openxmlformats_org_office_document_2006_relationships_id`, beside VML's own `id`. VML's
  booleans (`ST_TrueFalse`, `t`/`f`/`true`/`false`) are a string enumeration and are written as
  read, and an `xsd:float` or base64 value keeps the spelling it was read with (the fork's
  `ParserConfig(lexical_types=)`), so `o:spt="202"` is not saved `202.0` nor Word's wrapped
  `o:gfxdata` on one line; a value set in code takes the converter's spelling. Over 2,746 parts of
  422 documents nothing is newly skipped, no save gains a canonical difference, and 57 parts
  differ in bytes, each in attribute order alone; `pandoc-code.docx`'s `styles.xml`, whose
  lower-case `w:color` values were written in upper case, is canonically identical to its source
  now. With it:
  - the text model reads legacy WordArt, a VML picture's description and an object's ProgID from
    the typed tree, as it read them from the wildcard one;
  - `insert_ooxml` rewrites every relationships-namespace attribute of a typed node, so a VML
    picture's `r:id` is rewritten;
  - a VML text box's runs count in `styles_in_use()` and `fonts_in_use()`, as docx4j's traversal
    counts them (`DefaultParagraphFont` joins `vml/textbox.docx`'s styles);
  - no text box is made a typed island any more over the corpus; one below a wildcard node (in an
    element no class binds) still is;
  - `docx4j_py.namespaces.UNDERSTOOD` names `v`, `o`, `w10`, `xvml` and `pvml`;
  - `import docx4j_py` takes about 100 ms longer (640 to 745 ms here), for the five modules;
  - **`Pict`, `CTObject` and `CTBackground` (`w:pict`, `w:object`, `w:background`) hold their
    VML and `o:` children in one list, `any_and_any`** (docx4j's `getAnyAndAny()`), in document
    order, where they had two, `schemas_microsoft_com_vml_element` and
    `schemas_microsoft_com_office_office_element` (CR-012 section
    [12.16](docs/change-requests/CR-012-type-kept-mc-branches.md#1216-the-wobject-order-fix-implementation-notes-2026-10-22),
    `schemas/PATCHES.md` patch 6; under "Fixed", what the two lists did to a save). inkml's
    `TraceGroupType` likewise holds its five kinds of child in one list, `content` (patch 7).

### Fixed

- **Edits left cells Word repairs** (CR-016 section 2.1, a defect 0.1.0 shared): eleven edits left
  a cell ending in a table (`insert_table(..., location="End")` in a cell, `insert_xml`,
  `insert_ooxml` and `insert_markdown` of a table at a cell's end) or holding no block at all
  (`Paragraph.delete()` of a cell's only paragraph, `body.clear()`, a control's delete), which
  Word 365 opens as unreadable content; tracked, a cell's last mark was marked deleted, which Word
  never writes. Over the 422 documents of `out/cr012-b-corpus/` and their 2,435 main-part cells,
  every cell nine writes edited, untracked and tracked, now ends with a paragraph, and so does
  every cell after Accept All and Reject All of them (CR-016 section 10.4). Word 365 repairs the Word
  check's edits as written before the fix, adding a paragraph to exactly the 16 cells this release
  ends with one, and opens them as written now without a repair prompt (CR-016 section 13).
- **A value written into a control showing its placeholder stayed a placeholder** (CR-016 section
  2.4 and CR-019 section 3, a defect 0.1.0 shared): every write left `w:showingPlcHdr` on, so the
  control stayed a placeholder: `update_from_content_controls()` skipped a bound one as holding no
  value, and this release's readers, which read a placeholder as no text, would not read the text
  at all. `ContentControl.insert_text` also left the `PlaceholderText` style on its runs and wrote
  beside the prompt at `"Start"` and `"End"`, and a block control's paragraph wrote into the
  prompt's own run. Every route now writes in the prompt's place and clears the flag (CR-016 Phase
  B for `insert_text`, CR-019 for the rest; under "Changed").
- **A collapsed range's `insert_content_control` wrapped the run around its point** (CR-019 section
  10.9 item 1, a defect 0.1.0 shared): a control made at the point after `ONE` in a paragraph
  `ONE TWO` held `ONE TWO`. It is now made empty at the point, the run split there, showing its
  kind's placeholder (check 64 cases 19 to 21 and 40 to 42), and two controls made at one point
  keep their order.
- **A run-level control's `"Content"` range wrote outside it at `"Start"`** (CR-019 D4 and Q4, a
  defect 0.1.0 shared): its `"Start"` wrote into the run before the control, whether the control
  held text or showed its placeholder. Its writes now land inside at every location (under
  "Changed").
- **A `"Content"` delete left an empty run-level placeholder that Word won't leave Design Mode over**
  (CR-019 section 11.1, core-ts check 63's run 2, control 53; a defect 0.1.0 shared): the delete
  left a run-level control showing its placeholder over an empty `w:sdtContent`, over which Word 365
  refused to switch Developer > Design Mode off. It now shows Word's own five-space filler (D10 (a),
  under "Changed").
- **A write at a paragraph's end before a trailing control showing its placeholder went into the
  control** (CR-019 section 13.4 item 1, fixed by Phase D, section
  [15](docs/change-requests/CR-019-writing-into-a-placeholder.md#15-phase-d-implementation-notes-2026-10-25);
  core-ts's check 69 cases 9 to 24 and 41 to 49, and check 70's 65 to 72; a defect 0.1.0 shared):
  `insert_text(..., "End")`, the paragraph's End range at any location and
  `Body.insert_text(..., "End")` extended the prompt's own run, so 0.1.0 wrote the text grey, in
  Placeholder Text, inside a control still flagged as showing its placeholder (`X Click here.Y`),
  and this release's readers, which read a placeholder as no text, would not have read it at all.
  It now goes after the control, outside it, plain, and the control keeps its prompt (under
  "Changed").
- **A write at a paragraph's end went into a trailing control holding text** (CR-019 Phase D;
  core-ts's check 69 cases 33 to 40, check 70's 73 to 88; a defect 0.1.0 shared): the text joined
  the control's value (`X VALY`). It went into a trailing field's result too, which a field update
  replaces (check 70's cases 41 to 56), and into another author's tracked insertion, untracked,
  where it read as theirs (check 70's 33 to 40). It now goes after each, outside it (under
  "Changed").
- **A `Range.font` read split runs** (CR-016 Q12, found 2026-10-24): reading
  `paragraph.search("TWO")[0].font.bold` split the paragraph's run at the hit's edges, so the next
  save wrote three runs where the source had one. A read now splits nothing; only a write does.
- **A replacement of a hit inside a grapheme cluster split the character** (CR-016 Phase C, Q8;
  found building its Word-check file): `replace_text("cafe", "tea")` over `cafe` + U+0301 left the
  accent on the replacement's last letter (`teá`), and tracked it inserted `tea` before the
  cluster and deleted nothing, the deletion's split having snapped back to the cluster's start.
  Such a hit is no longer found (under "Changed").
- **A tracked `TableCell.value` lost the cell's old content untracked** (CR-016 Q11): it is now a
  tracked deletion, the new paragraph inserted after it, so Reject All gives the cell back
  (`Table.values` and `TableRow.values` with it). The python-docx facade's `_Cell.add_table`
  keeps python-docx 1.2.0's result exactly (Q13): the cell's paragraph untouched, the table, one
  plain `<w:p/>`.
- **On/off values as Word reads them** (CR-018 Phase A, section 1.2's findings): `w:b
  w:val=" false"` and `"false "` (bold in Word) were trimmed, read off and **saved `0`**, the one
  save over Word's probes that changed meaning; they are kept, read on and saved bare. `"off"`
  read on in every reader (CharBold with a direct `"off"` read bold), so `w:vanish w:val="off"`
  hid text Word shows (from `Paragraph.text`, `search`, the text model and Markdown), and
  `range.font.bold = True` across paragraphs left an `"off"` run unbold; it reads off now. The
  facade read `w:val=""` off; `describe()` and the change tracker read `w:trackRevisions
  w:val="off"` on; `w:temporary` and `w:multiLine` read `"FALSE"` off. A bare `<w14:checked/>`
  read unchecked though its schema default is true (Phase C). `w15:done` read through the
  extractor never matched its enumeration member, so a typed comment's `resolved` was false.
- **Colours, underlines and highlights as Word reads them** (CR-018 Phase B, section 1.2 finding
  5): a colour name Word draws (`red`, `DarkRed`) came back from `Font.color` as the name and
  from the facade as None; the facade read five or seven hex digits as `FF0000`, as python-docx's
  slicing does, where Word draws automatic; `highlight_color` returned `"YELLOW"` and `"bogus"`
  raw; an invalid direct `w:u`, `w:highlight`, `w:vertAlign`, `w:em` or `w:color` overrode the
  style's value beneath, which Word shows; `w:u w:val="true"` over an underlining style read as
  the style's underline, where Word removes it; a character style's highlight was drawn.
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
  (`docx4j-xsdata` 26.2.2).
- **A comment's markers and a `Range.font` write within one paragraph stay inside the range**
  (CR-013 section 14.6): beside a run holding a tab or a break they took the whole run (a comment on
  `b\tc` in a run `a\tb\tc` covered all of it; bold on `b` made the whole run bold). The run is split
  around the range now, as core-ts's `splitRunsAround` splits it.
- **A tracked document's `Range.get_xml()` no longer gives the range's paragraph's other words as
  deleted** (CR-013 section 14.6): the copy was trimmed through a tracked splice. It is trimmed
  untracked now.
- **Text typed inside another author's tracked insertion splits it** (CR-013 section 14.2), the new
  `w:ins` between the halves, where it went after the whole insertion and the range returned read the
  wrong characters; and a tracked insertion between a tab and the text after it in one run stands
  there, where it went after the run.
- **An embedded object opens again after a save**
  ([CR-012](docs/change-requests/CR-012-type-kept-mc-branches.md) section
  [12.16](docs/change-requests/CR-012-type-kept-mc-branches.md#1216-the-wobject-order-fix-implementation-notes-2026-10-22)).
  Since 0.1.0, saving a document whose main part had been unmarshalled (any edit of the body)
  wrote each `w:object`'s `o:OLEObject` before its `v:shape`, and Word showed the object's icon but
  would not open the embedded object on a double-click. `CT_PictureBase`, the content model of
  `w:pict`, `w:object` and `w:background`, was a repeating sequence of a VML wildcard and an `o:`
  wildcard, bound as two lists that were written interleaved, one item of each in turn; a marked
  schema patch makes it one wildcard over both namespaces, which accepts the same documents and
  keeps the children in document order, as docx4j's JAXB binding does (`schemas/PATCHES.md`
  patch 6). A `w:pict` mixing VML and `o:` children was reordered the same way and is kept as
  read too. Over 2,746 parts of 422 documents, the 17 main parts whose `w:object`s were reordered
  are canonically identical to their sources now.

### Dependencies

- `docx4j-xsdata[lxml] >=26.2.3,<26.3`: the fork's release 3,
  [docx4j-xsdata 26.2.3](https://pypi.org/project/docx4j-xsdata/26.2.3/), published 2026-10-24
  (its `docs/fork/CHANGES.md` stage 13, items 29 to 31): `ParserConfig(exact_types=)`,
  `SkippedReport.kept` and `XmlVar.schema_default_value`, which CR-018 Phases A and C need.
- Before it, `docx4j-xsdata[lxml] >=26.2.2,<26.3` (was `>=26.2.1`): the fork's release 2,
  [docx4j-xsdata 26.2.2](https://pypi.org/project/docx4j-xsdata/26.2.2/), published 2026-10-22
  (its `docs/fork/CHANGES.md` stages 9 to 12, items 22 to 28). This release needs all four: stage 9's
  unknown attribute on a simple-typed element in the skipped report, stage 10's namespaces in
  scope for the text model (CR-011), stage 11's branch typing hooks (CR-012 Phases A to B) and
  stage 12's `lexical_types` (CR-012 Phase C). The `dev` extra's floor moves with it.

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

[Unreleased]: https://github.com/plutext/docx4j-python/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/plutext/docx4j-python/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/plutext/docx4j-python/releases/tag/v0.1.0
