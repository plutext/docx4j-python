# CR-025: Dates in the package's zone, and docx4j CR-067 phase 1's fields in the text model --- the portfolio's date rule, and the tier-1 fields under `fields: update`

**Status:** Proposed 2026-10-28, on the user's ruling on the cross-product date policy the same day (`../docx4j-portfolio/docs/date_handling_across_ports.md` section 3.6, committed at `4ac305b`: "I confirm option (a) as amended, and 3.5; also the three sub-items") and his instruction as the portfolio session relayed it: "commit, then have the ports do their work". Three phases (section 6): A, the package's time zone and the core properties' opt-in (the rule's points 2, 5 and 6, after docx4j's D12); B, docx4j CR-067 phase 1's fields in the text model, held to Word's probe (point 1 for the zone); C, the records. A and B are built at the coordinator's brief; C is not, beyond this CR's row in the index and the note for CR-014 in section 9. The names follow docx4j-core-ts's, at the coordinator's request the same day (section 4.5). The questions are in section 8, each with a recommendation. Where building needed an answer, the recommendation is built and listed for the user to decide, as CR-020 sections 10.12 and 11.9 did. **Phase A implemented 2026-10-28** (section 11; `cbafbc1`): `pkg.date_time_zone`, the five places an instant meets a clock, `pkg.write_times_on_save`, its stamp to the second after a finding while building (11.2, Q3 restated). **Phase B implemented 2026-10-28** (section 12; `23fa5ce`): docx4j's evaluator ported, the text model's `fields: update` at docx4j's set; Word's probe given back in 21 of 25 cases and the cached result in the 4 where Word writes an error text; docx4j's own extraction of the probe matched byte for byte in five modes; no golden moved. **Q1 to Q10 decided as recommended 2026-10-28** (the user: "CR-025's ten questions as recommended"; 12.10, nothing changed). No Word check: extraction is no file, and with the zone unset and the opt-in off a saved file is unchanged (acceptance 0 differ). **Phase B2 implemented 2026-10-28** (section 13; `3fe390d`), a phase the user ordered once docx4j's date-forms probe had run (its CR-067 section 12.2, `c58e36e64`): "item 3 as recommended; go on CR-025 B2". A core property's date alone or time with no offset is UTC, as Word reads it (docx4j's D13; Q8's reading ends, as Q8 foresaw); a date alone in a revision's or a comment's `w:date` is local midnight, and in a date control's `w:fullDate` midnight; a save writes a `w:date` or `w:fullDate` given as a date alone or a time with no offset in Word's form, the nominal clock to the minute labelled `Z`, and leaves every value with a `Z` or an offset as read. No file of the corpus or the acceptance set holds such a value, so no saved file there changed. **13.10's four points decided as recommended 2026-10-28** (the user: "four B2 points in CR-025 section 13.10, all as recommended"; 13.11): `PRINTDATE` is `cp:lastPrinted` and the facade's core dates are aware in UTC, both built; the other two keep what was built. **Phase C, the records, done 2026-10-28** (section 14, the user: "do CR-025 Phase C"). **CR-025 is complete.**

*The number:* CR-024 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057.

**Depends on:**
- **docx4j CR-067 phase 1** (`d6ca8dbee`): the tier-1 fields, measured against Word 365 (its section 11.3, Jason's F9 run, version 2608 build 20326.20232, 2026-10-28). Its probe, Word's save and the results are this CR's oracle (section 5).
- **docx4j CR-067 D12** (`c3e9920e5`): a time zone per package, `OpcPackage.setDateTimeZone`, `effectiveDateTimeZone()`, `FieldContext.zone`, and `TrackedChange.getDate()` reading a `w:date` without `dateUtc` in it (its section 12). docx4j CR-058 D6 as amended the same day: the bound-date conversion its own opt-in, off by default.
- **The portfolio's rule**, `docs/date_handling_across_ports.md` section 3.6 (`4ac305b`).
- **Counterparts:** docx4j as above; docx4j-core-ts, a package-level zone option in its own CR (the rule's section 3.4); docx4j-mcp, a `time_zone` argument per tool (docx4j CR-067 section 12); `../docx4j-mcp-py`, the same when CR-006 is built.
- **This repository:** CR-010 section 12.9 (Word's date convention, the readers and writers Phase A changes); CR-011 sections 6 and 13.3 (`fields: update`, the evaluator Phase B changes); CR-014, proposed and not built, whose bound dates the rule's point 3 governs (section 9).

## 1. Why

The ports agreed on the kinds of date without being told: a revision, a comment and a date field hold an instant, and a bound date a nominal day. What the user's ruling settles is the zone an instant is shown in, and that the zone belongs to the package, the client's, not the process's. In full:

1. **Kinds.** A revision stamp, a comment stamp and a date field hold an instant. A bound date holds a nominal day. An extractor gives every stored date as written and converts nothing.
2. **Zone.** An instant is rendered, and a `w:date` without `dateUtc` read, in the package's zone: a zone the caller sets on the package it is working on, from the client it serves. Unset, the process's zone. A per-call override where a port has one.
3. **Word's wart.** A bound date is the day its fields say by default. Word's conversion of it in a zone is a separate per-package opt-in, named for what it does, off by default, converting in the package's zone. Setting the zone does not turn it on.
4. **Other producers' forms.** Date-only and offset-free values are unmeasured in Word; no port changes its reading until a Word probe has been run. Each port documents its current reading meanwhile.
5. **Stamps.** A new revision or comment is stamped to the minute, as Word writes it.
6. **Core properties on save.** Untouched by default. docx4j's opt-in (`docx4j.dcterms.write`: set `dcterms:modified` to now, and `created` for a new package without one) is the shape every port offers, under one name per port, off by default.
7. **Principle.** Follow Word where its behaviour is correct for the kind of value; where it is a wart, the intent is the default and Word's behaviour an opt-in named for what it does.

This repository already follows points 1 and 5 (CR-010 section 12.9; the text model gives every stored date as written). Point 2 is half done: every instant is shown in the process's zone, with no way to name the client's. Point 6 has no opt-in. And the evaluator under `fields: update` is docx4j's `FieldValues` at `198d015a3`, before CR-067: it reads `CREATEDATE` as text, so its result is in no zone at all (section 3.2).

## 2. What Word does (docx4j CR-067 section 11.3)

docx4j's probe (`field-probe-r1.docx`, 25 cases, every field in the complex form with the cached result `OLD` but case 25's `w:fldSimple`) has the core properties title Probe Title, author Probe Author, revision 7, keywords "alpha, beta", created `2024-02-03T04:05:06Z` and modified `2025-06-07T08:09:10Z`; a bold bookmark `bm1` ("Bookmarked text"); a bookmark `bmNote` around a footnote reference; and a section break before cases 24 and 25. Jason opened it in Word 365 on a machine in Sydney's zone, pressed Ctrl+A and F9, and saved it (`field-probe-r1-word365.docx`). `results.txt` is his record:

| Case | Field | Word 365 on F9 | What it says |
|---|---|---|---|
| 1, 3 | `CREATEDATE`, `SAVEDATE` | `2/3/2024 3:05:00 PM`, `6/7/2025 6:09:00 PM` | the instant in the machine's zone (UTC+11, then +10), **to the minute** (`:06` and `:10` show as `:00`), in the short date and long time; `SAVEDATE` is the file's modified time, not the session's |
| 2, 4, 25 | the same with `\@` | `2024-02-03 15:05`, `7 June 2025`, `2024-02-03` | the switch as expected; the simple field updates as the complex one |
| 5 | `REVNUM` | 7 | the file's revision |
| 6, 7 | `INFO Title`, `INFO Author` | Probe Title, Probe Author | `INFO` is the field it names |
| 8, 9 | `INFO`, `INFO Nope` | `Error! Unknown info keyword.` | an error text |
| 10, 24 | `SECTION` | 1, 2 | |
| 11, 12 | `QUOTE "hello world"`, `QUOTE` | hello world, **empty** | no argument is empty, not an error |
| 13, 14, 16 | `REF bm1`, `\* Upper`, `\h` | Bookmarked text (in the bookmark's bold), BOOKMARKED TEXT, Bookmarked text | `\h` changes nothing in the result |
| 15, 19 | `REF nope`, `NOTEREF nope` | `Error! Reference source not found.`, `Error! Bookmark not defined.` | error texts |
| 17, 18 | `NOTEREF bmNote`, `\f` | 1, 1 (with `\f` in the `FootnoteReference` style) | |
| 20 to 22 | `USERNAME`, `USERINITIALS`, `USERADDRESS` | before F9 `OLDUSER` in all three; after, Author A, AA, **empty** | Word does not update them on open; a blank address shows nothing |
| 23 | `KEYWORDS` | alpha, beta | the control |

Word's Save As then wrote created and modified as the save's moment and revision 2: to Word the saved file is a new document.

**And the core properties.** Word writes `dcterms:created` and `dcterms:modified` to the minute. All 296 such values in this repository's 148 Word 365 saves (every `*word365*.docx` under `tests/`) end `:00Z`; Word's save of the probe wrote `2026-10-28T00:17:00Z`.

## 3. What exists here (measured 2026-10-28 at `3c63804`)

### 3.1 Where an instant becomes a clock time

Five places, all in the process's zone (`datetime.astimezone()` with no argument), with no caller option:

| Place | What it does |
|---|---|
| `model/content/tracking.py`, `xml_date` and `ChangeTracker.markup` | a new revision's `w:date`: the instant on the local wall clock with a `Z`; `w16du:dateUtc` in UTC beside it; "now" to the minute, or `pkg.tracked_change_date` as given |
| `model/content/comments.py`, the comment writer (`_new_comment`) | a new comment's `w:date`, the same way, to the minute; `w16cex:dateUtc` in UTC |
| `model/revisions/change.py`, `word_date` and `date_of`; `TrackedChange.date` | a `w:date` without `dateUtc` read as the local wall clock, its `Z` notwithstanding |
| `model/content/tracking.py`, `date_of` | the same, for the content API |
| `model/content/comments.py`, `Comment.creation_date` | a comment's `w:date` without `w16cex:dateUtc`, the same |

A `dateUtc` is UTC and stays so. A `w:date` with an explicit non-zero offset, which Word never writes, is read as it says.

### 3.2 The evaluator

`docx4j_py/model/text/field_values.py` is docx4j's `FieldValues` at `198d015a3` (CR-011 section 13.3): `DOCPROPERTY`, `DOCVARIABLE`, the document-information fields, `SEQ`, `REF` and `QUOTE`, through one parser (`FldSimpleModel`) and docx4j's switches (`model/fields/`). Over the probe, under `fields: update` (this repository at `3c63804`, any zone):

- `CREATEDATE` and `SAVEDATE` give the core property's lexical form, `2024-02-03T04:05:06Z`, with or without `\@`. The portfolio's section 2 says "with `\@` parsed nominally, the `Z` ignored"; measured, it is not even that: docx4j's inferred formats at `198d015a3` have no pattern with a `Z`, so the switch finds no date and leaves the text (cases 1 to 4 and 25).
- `INFO`, `SECTION`, `NOTEREF` and the user fields are not computed: each keeps `OLD` (cases 6 to 10, 17, 18, 24).
- `QUOTE` with no argument keeps `OLD` (case 12).
- `REVNUM`, `QUOTE "hello world"`, `REF` and `KEYWORDS` are Word's (cases 5, 11, 13, 14, 16, 23); the three error cases keep `OLD`, as docx4j's do (its D5).

So 12 of the 25 cases differ from docx4j's text model at `c3e9920e5` on the same file. docx4j's extraction of the probe, run here from its build of that commit (section 5), gives Word's result in every case but the three error texts (kept `OLD`, D5) and the three user fields (kept `OLDUSER`: its text model supplies no user).

### 3.3 The core properties on save

`create_package` writes `created` and `modified`, in UTC to the second. A save never touches them: a part never unmarshalled is written byte for byte, and nothing unmarshals the core part on a save.

## 4. Design

### 4.1 The package's zone (D1, D2; Phase A)

**D1.** `pkg.date_time_zone`: the client's zone for this package, a `datetime.tzinfo` or an IANA name (`"Asia/Tokyo"`, read as `zoneinfo.ZoneInfo`), None by default. Never saved: a request attribute, as docx4j's is, so a server reusing a loaded package sets it per request. A bad name or another type raises `ContentError` (`date_time_zone.invalid`). `pkg.effective_date_time_zone` is the zone used: the package's, else the process's, as an object (`docx4j_py.model.dates.PROCESS_ZONE`, a `tzinfo` that reads the process's zone at each instant exactly as `astimezone()` does). docx4j's `getDateTimeZone()` and `effectiveDateTimeZone()`. A dry run's trial starts with its package's and may be given its own, as its `tracked_change_date` may.

**D2.** The fallback is the process's zone, and nothing between. docx4j's chain has a JVM-wide property between the package and the JVM default (`BindingHandler.dateTimeZone`), because a JVM's default zone is awkward to set per deployment. A Python process's zone is `TZ`, which a deployment sets as easily as it would set a property, so a second process-wide setting would be a second `TZ`.

**D3.** Every place of section 3.1 uses the effective zone. Unset, the code path is today's (`astimezone()` with no argument), so output is byte for byte what it was. `dateUtc` stays UTC; a non-zero offset is still read as it says; the stamp is still to the minute and `tracked_change_date` still as given. The text model's dates (a revision's or a comment's in `markup` view, the metadata) are given as written, converting nothing (point 1), as today.

### 4.2 The core properties' opt-in (D4; Phase A)

**D4.** `pkg.write_times_on_save`, a package attribute, False by default: docx4j-core-ts's `writeTimesOnSave` in Python's spelling (section 4.5). When on, a save sets `dcterms:modified` to now, and `dcterms:created` too where the core part has none and was not read from the source (made by `create_package` or by the caller): docx4j's `DocPropsCorePart.writeTimes()` (17.3.3, its CR-036). In UTC with a `Z`, **to the second**, as docx4j and docx4j-core-ts write it. Word writes these to the minute (section 2), and point 5's "as Word writes it" could be read to reach them; but point 5 is about a revision's and a comment's stamp, point 6 names docx4j's shape, and `create_package` writes `created` to the second, so a stamp to the minute would put `modified` before `created` in a package made and saved within the minute (found while building, section 11.2). A package with no core part gets none. `cp:lastModifiedBy` is not touched: docx4j writes it under its own property, `docx4j.dc.write`, which the rule does not name. `last_save.times` is the stamp written, or None. docx4j-core-ts writes the second too, and makes a core part for a new package that has none (its `createPackage` writes an empty one); here `create_package` writes one with both dates, so the case does not arise, and a package with none is left without, as docx4j leaves it.

### 4.3 The evaluator (D5 to D7; Phase B)

**D5.** docx4j's evaluator, ported where the text model needs it: `docx4j_py.model.fields.FieldInstruction` (the one parser), `FieldContext` (the package, the zone override, a clock, the user, the section, and the pass's sequences, bookmarks and note numbers), `FieldUser`, and a registry by name (`docx4j_py.model.fields.evaluators`: `register`, `get`, `has`, `names`, `evaluate`), each entry a callable returning the result or None to keep the cached one (docx4j's D5), with a flag for the date fields that format themselves. The text model's `FieldValues` becomes a thin consumer over it with docx4j's `COMPUTED` set at `c3e9920e5`: `DOCPROPERTY`, `DOCVARIABLE`, `SEQ`, `REF`, `QUOTE`, the nine document-information fields, and phase 1's `INFO`, `SECTION`, `NOTEREF`, `USERNAME`, `USERINITIALS` and `USERADDRESS`.

**D6.** The parser: `FieldInstruction` is ported, with docx4j's accessors (`name`, `parameter_text`, `parameters`, `argument`, `has_switch`, `switch_value`, `switch_values`, `index_of_switch`, `value_at`, `unquote`), and `FldSimpleModel` becomes a subclass that keeps its one extra name, `parameter_string`. Its tokeniser is already the one docx4j moved into `FieldInstruction`, so nothing a caller reads changes; `FieldInstructionTest`'s 31 shapes hold both to it. `FldSimpleModel` is not deprecated with a warning: the suite turns a `DeprecationWarning` into an error, and the class is public in 0.2.0.

**D7.** No language reaches the text model's evaluator, as none reaches docx4j's inspector (`FieldContext.of(pkg)`): month and day names are English, as `model/java_text.py` writes them.

### 4.4 The fields (D8 to D14; Phase B)

**D8. `CREATEDATE` and `SAVEDATE`.** The core property's W3CDTF with an offset or a `Z` is an instant, truncated to the minute on the instant (docx4j's `floorMod` of the epoch), shown in the package's zone (the context's override, else `pkg.date_time_zone`, else the process's). With `\@`, the switch formats it; with `\@` and no value, the short date (`M/d/yy`), as Java's default short date instance gives it in docx4j. With no `\@`, Word's default: the short date and the long time, `M/d/yyyy h:mm:ss a` in US English (`2/3/2024 3:05:00 PM`), docx4j's choice for no language. docx4j gives another language the locale's short date and medium time, unmeasured; no language reaches this evaluator (D7). A value that is no date keeps the cached result, as docx4j's does (today the raw text was shown). An argument is ignored, as docx4j's `CoreDate` ignores it.

**D9. The forms Word has not been measured on** (the rule's point 4). A date only (`2024-02-03`, also `2024` and `2024-02`) and a time with no offset (`2025-06-07T08:09:10`) are read as today, nominally: the wall clock as written, converted by no zone. The field's own rules, the minute and Word's default format, apply to them as to any date the field shows. With no package zone set, this is docx4j's result too: it reads such a value as local time in the JVM's zone and shows it in the same zone. With one set it differs: docx4j shows the JVM-local reading converted to the package's zone. Docx4j's date-forms probe (its CR-067 section 12.1, staged for Jason) measures what Word does with all three forms; until then each port keeps its reading (Q8). **Phase B2 (section 13):** the probe has run, and Word reads both as UTC (created `2024-02-03` showed 11:00 in Sydney, modified `2025-06-07T08:09:10` 18:09), so both are now read as UTC, as docx4j's D13 reads them, and shown in the package's zone to the minute like any other instant.

**D10. The user fields.** `FieldUser(name, initials=None, address=None)`, after docx4j's `FieldContext.user(name, initials, address)`. The text model takes it as `TextOptions(field_user=...)`, an option beyond section 10 like `image_describer`, a host's hook (a `FieldUser`, a name alone, or a mapping with `name`, `initials` and `address`). Unset, the three fields keep their cached result, which is also what Word shows until F9 (cases 20 to 22). Set, `USERNAME` is the name, and `USERINITIALS` and `USERADDRESS` their values or empty where the user has none, as Word shows a blank one. docx4j's extraction has no way to supply a user: this is an extension, and with no user the two agree.

**D11. `QUOTE`** with no argument is empty (case 12). An unknown `INFO` name, a missing bookmark (`REF`, `NOTEREF`) keep the cached result, where Word writes its error text (docx4j's D5).

**D12. `INFO <name>`** evaluates the field it names, with its switches (`INFO Title \* Upper` is `TITLE \* Upper`), **only where that field is one the consumer computes**. docx4j's `INFO` asks the whole registry, so its extraction evaluates `INFO PrintDate` at the clock, though its text model keeps `PRINTDATE` itself cached; here it keeps the cached result. No name, an unknown one, or `INFO INFO` keeps the cached result.

**D13. `SECTION`** is the section of the main story the walk is in, from 1, as docx4j's text model counts it (the sections ended so far, plus one); elsewhere (a note, a header) the cached result. **`NOTEREF`** is the number of the footnote or endnote whose reference the bookmark holds: docx4j's `FieldSources.noteNumbers`, the main part walked in document order, each kind counted apart from 1, a custom mark not counted, separators skipped; the walk goes where docx4j's goes (not into a deletion, a move, a drawing, a picture, an object, an equation, ruby or `mc:AlternateContent`).

**D14. `DATE`, `TIME` and `PRINTDATE`** are registered, at the context's clock in its zone, for a caller of `evaluate`; the text model keeps their cached result, as docx4j's does (its `FieldValues` leaves them out of `COMPUTED`). Extraction stays deterministic (Q6).

**What extraction does not show** (pinned by the tests as write-back only): a `REF`'s result in the bookmark's formatting (case 13's bold), `NOTEREF \f`'s `FootnoteReference` style (case 18), and `\h`'s hyperlink. The text model gives text; a write-back is docx4j's `FieldUpdater`, which this repository has no port of.

### 4.5 The names, aligned across the ports (the coordinator, 2026-10-28)

The coordinator asked, the same day and before anything was built, that the three ports name these alike, taking docx4j-core-ts's names (its CR-018, which mirrors docx4j `c3e9920e5`; uncommitted then, committed later the same day as `d9c0cef`) in snake_case. So:

| What | docx4j | docx4j-core-ts | here |
|---|---|---|---|
| the package's zone | `OpcPackage.setDateTimeZone(ZoneId)`, `getDateTimeZone()` | `pkg.dateTimeZone` (an IANA id; a `RangeError` for an unknown one) | `pkg.date_time_zone` (a `tzinfo` or an IANA name; `ContentError` `date_time_zone.invalid`) |
| the zone in force | `effectiveDateTimeZone()` | `effectiveDateTimeZone()` | `pkg.effective_date_time_zone` |
| a process-wide fallback | `BindingHandler.dateTimeZone` | none | none (D2) |
| the core properties' opt-in | `docx4j.dcterms.write` (process-wide) | `pkg.writeTimesOnSave` | `pkg.write_times_on_save` |
| its stamp | the second | the second | the second (D4, Q3) |
| Word's bound-date conversion | `setBoundDatesConvertedAsWord(true)` | `pkg.boundDatesConvertedAsWord` | `pkg.bound_dates_converted_as_word`, when CR-014 is built (section 9) |
| a new package's core part | `created` and `modified` | empty, no dates | `created` and `modified`, to the second (Q10) |

The coordinator's earlier suggestion, `pkg.time_zone`, is replaced.

## 5. Tests and oracles

- **The probe**, copied unedited from `../docx4j/docx4j-core-tests/src/test/resources/fields/cr067-phase1/` into `tests/fixtures/field-probe-cr067/` with a `SHA1SUMS`: `field-probe-r1.docx`, `field-probe-r1-word365.docx`, `results.txt`. `FieldProbeWordTest`'s replay, ported through the text model: in Sydney's zone, every case Word's result as `results.txt` gives it, but the three error texts (the cached `OLD`, D11) and, with no user supplied, the three user fields (`OLDUSER`); with `FieldUser("Author A", "AA")`, those three Word's too. The zone pinned twice over: the process's, by `TZ`, and the package's, by `pkg.date_time_zone`, with the process elsewhere.
- **docx4j's own extraction of the probe**, run here from docx4j's build of `c3e9920e5` (`Texts.plainText` under `fields: update`, the JVM in Sydney's zone), byte for byte: the same plain text (section 11).
- **Word's save read with cached results** gives `results.txt` case for case: the oracle file and the record agree.
- **`FieldEvaluatorsTest`**, the cases that apply: the property fields, the pass's sequences and references, the date fields at a clock, a caller's own evaluator, the core-property dates, `REVNUM` and `INFO`, `SECTION`, `NOTEREF` and the user.
- **`PackageDateTimeZoneTest`**, the cases that apply: the fallback chain; the date fields in the package's zone (Tokyo), and the per-call override; a revision's `w:date` `2024-02-03T04:05:06Z` read in Tokyo as `2024-02-02T19:05:06Z`, its `dateUtc` unchanged. The bound-date case has no counterpart until CR-014.
- **`FieldInstructionTest`**: its 31 shapes through `FieldInstruction` and `FldSimpleModel`, every accessor compared.
- **Phase A's writers**: a revision and a comment stamped with a package zone set write `w:date` on that zone's wall clock and `dateUtc` unchanged; unset, byte for byte today's in two process zones.
- **The opt-in**: off, a save is byte for byte; on, `modified` (and `created` where D4 says) to the second in UTC, `last_save.times`, nothing else in the part changed, and a package with no core part untouched.
- **The gates**: the suite, `scripts/roundtrip.py` 50 of 50, `scripts/threads.py`, `compare_acceptance.py` 0 differ against `out/acceptance-cr020b/` with the zone unset and the opt-in off; `codegen/generate.sh --check` for the new files.
- **Phase B2: docx4j's date-forms probe** (its CR-067 section 12.2), copied unedited into `tests/fixtures/date-forms-cr067/` with a `SHA1SUMS`, and `DateFormsProbeWordTest`'s replay ported (section 13.6): the fields through the text model, the revisions, the date controls, Word's save, and this save against it.

No Word check is expected: extraction output is not a file, and with the zone unset and the opt-in off a saved file is unchanged. The opt-in, when on, re-marshals the core part, which the corpus round trip already holds canonically identical, and writes the form `create_package` writes, which Word has opened since 0.1.0.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | D1 to D4: `pkg.date_time_zone` and `effective_date_time_zone`, `model/dates.py`, the five sites of 3.1, the dcterms opt-in; the tests of section 5 for them | ½ day |
| B | D5 to D14: `FieldInstruction`, `FieldContext`, `FieldUser`, the registry, the evaluators, `FieldValues` over them, `TextOptions(field_user=)`; the probe, docx4j's output, `FieldEvaluatorsTest`, `FieldInstructionTest` | 1 day |
| B2 | Section 13, after docx4j's date-forms probe: a core property's bare date UTC (docx4j's D13); a date alone in a `w:date` local midnight, in a `w:fullDate` midnight; the save's normal form for a date-only or offset-free `w:date` and `w:fullDate` (`values.date_time`, `values.date_time_normal_form`, `normalise.DATE_ATTRIBUTES`); the probe's replay | ½ day |
| C | The records: CHANGELOG (Added: `date_time_zone`, `write_times_on_save`, the fields, `field_user`; Changed: `CREATEDATE`/`SAVEDATE`, `QUOTE`), README and `docs/pypi-readme.md` where they name the readers and the save, CLAUDE.md (the status line, the new files in its list), the index's row brought up to date, CR-014's note (section 9) folded into CR-014; the portfolio's `python/CR-025.*` entries through the docx4j-portfolio session | ¼ day |

## 7. Risks

- **Moved expectations.** A test or golden that read a `CREATEDATE` as text, or a `QUOTE` with no argument as cached, moves. Section 3.2's measurement and a scan of every fixture's instructions say none does: no fixture under `tests/` or `samples/` holds a tier-1 field but `QUOTE` with an argument, and the two `TIME` fields stay cached.
- **The process's zone, read as an object** (D1's `PROCESS_ZONE`), disagreeing with `astimezone()` at a gap or a fold. It is only reached through `effective_date_time_zone`; the sites themselves keep `astimezone()` when nothing is set, and a test holds the object to it at every half hour of a year in three zones with daylight time, one of them a half-hour change (Lord Howe).
- **A server that forgets to set the zone** shows its own: the rule's default, and why D1's attribute is on the package.

## 8. Questions for the user

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The zone's name and type: `pkg.date_time_zone`, a `tzinfo` or an IANA name, None for the process's; `pkg.effective_date_time_zone` the zone used (D1)? | Yes: docx4j-core-ts's `dateTimeZone` and `effectiveDateTimeZone()` in snake_case (section 4.5), taking a `tzinfo` too, as Python's own zones are objects. |
| Q2 | A process-wide fallback between the package and the process's zone, as docx4j's JVM-wide property is (D2)? | None. A Python process's zone is `TZ`; a second setting would be a second `TZ`. |
| Q3 | The opt-in's stamp: to the second, as docx4j and docx4j-core-ts write it, or to the minute, as Word writes it (296 of 296 values in Word's saves here) (D4)? | The second, the three ports alike. The minute would want `create_package`'s stamps to the minute too (Q10), or `modified` can precede `created`; if the user wants Word's minute, it is one constant here, and the same change in the other two ports. |
| Q4 | The opt-in's name and shape: `pkg.write_times_on_save`, a package attribute, as docx4j-core-ts's `writeTimesOnSave` (D4)? Or a save option as well, as `save(normalise=)` is? | The attribute alone, named as core-ts names it, so that the ports read alike; a per-save option can follow if a caller needs one. |
| Q5 | The parser: port `FieldInstruction`, `FldSimpleModel` a subclass kept, both held to the 31 shapes; or keep `FldSimpleModel` alone and hold it to them (D6)? | Port it. The evaluators read as docx4j's do, accessor for accessor, and the tokeniser is already the same; nothing a caller reads changes. |
| Q6 | `DATE`, `TIME` and `PRINTDATE` under `fields: update`: keep the cached result, as docx4j's text model does, or evaluate them at a clock (D14)? | Keep the cached result. A "now" makes extraction non-deterministic, and docx4j's extraction does not evaluate them. They are in the registry for a caller who supplies a clock. |
| Q7 | The user for `USERNAME`, `USERINITIALS` and `USERADDRESS`: `TextOptions(field_user=...)`; or `pkg.author`, the identity a revision and a comment carry; or none, as docx4j's extraction has none (D10)? | `field_user`. It is per call, as docx4j's `FieldContext.user` is, and a server's user is a request's. `pkg.author` has a default and is the content API's; using it would change extraction for every document with a user field. |
| Q8 | A date-only or offset-free core property with a package zone set: shown as written, this repository's reading; or read in the process's zone and converted to the package's, docx4j's (D9)? | As written, until docx4j's date-forms probe has run (point 4); then both ports follow Word. (Phase B2: the probe has run, and both follow Word, UTC; section 13.) |
| Q9 | `INFO <name>`: only the fields the consumer computes; or every field the registry has, as docx4j's, so that `INFO PrintDate` shows the clock in extraction (D12)? | Only the consumer's. docx4j's reach looks unintended (its text model keeps `PRINTDATE` cached); offer docx4j the narrower rule through the portfolio. |
| Q10 | `create_package` writes `created` and `modified` to the second. Leave it, or write them to the minute as Word does? | Leave it. The rule's point 6 is about saves; changing a new package's stamps moves every saved file's core part for no reader's gain. Raise it with the portfolio if a port wants one rule for all core stamps. |

## 9. A note for CR-014 (Phase C folds it in)

CR-014 is proposed, not built. Its Q10 (decided 2026-10-23: docx4j CR-058's nominal default, no `dateTimeZone` option) stands: a bound date is the day its fields say, on every machine. The rule's point 3 adds Word's conversion as a separate per-package opt-in, off by default, after docx4j's `OpcPackage.setBoundDatesConvertedAsWord(true)` (CR-058 D6 as amended 2026-10-28): when on, the stored value's `Z` or offset is read as an instant and converted to `pkg.effective_date_time_zone` (this CR's D1), as Word does on a machine in that zone; setting `pkg.date_time_zone` alone does not turn it on. A Python name in this CR's idiom would be `pkg.bound_dates_converted_as_word`, with no process-wide fallback (this CR's D2). It belongs to CR-014's phase 1b, which formats a bound date.

## 10. Effort

1½ to 2 days: A ½, B 1, C ¼.

## 11. Phase A implementation notes (2026-10-28)

Phase A as built, D1 to D4, at the coordinator's brief. Phase C (the records) is not this phase's: no CHANGELOG, README, `docs/pypi-readme.md` or CLAUDE.md change was made, beyond this CR's row in the index (with the draft, `df8d51a`) and `tests/README.md`'s rows for the new tests.

**Commits.** `df8d51a`, the draft and the index's row. `cbafbc1`, Phase A: `model/dates.py`, the five sites, the properties and the opt-in, and their tests. The dry run's own zone (11.1) landed with Phase B's commit.

### 11.1 What landed

| piece | where | what |
|---|---|---|
| `pkg.date_time_zone` | `model/content/__init__.py` (installed by `register()`, as `tracked_change_date` is); the slot `_date_time_zone` on `OpcPackage` | None by default; a `tzinfo` as given, an IANA name as `zoneinfo.ZoneInfo`; anything else `ContentError` `date_time_zone.invalid`, the zone set before kept |
| `pkg.effective_date_time_zone` | the same | the package's, else `PROCESS_ZONE` |
| `model/dates.py` | new, hand-written, in `KEEP` | `ProcessZone` and `PROCESS_ZONE` (the process's zone as a `tzinfo`, read at each instant as `astimezone()` reads it, its `fromutc` setting `fold` in a repeated hour); `zone_of`, `package_zone`, `effective_zone`, `wall_clock`, `instant_of`, `to_the_minute`. Each takes `zone=None` for the process's and then calls `astimezone()` with no argument, the code that ran before |
| the writers | `tracking.xml_date(value, zone=None)`, `ChangeTracker.markup`; the comment writer | `w:date` on the package's wall clock; `dateUtc` UTC as before |
| the readers | `revisions.change.word_date` and `date_of` (`zone=None`), `TrackedChange.date` (the story's package), `tracking.date_of`, `Comment.creation_date` (the body's package) | a `w:date` without `dateUtc` read in the package's zone; a non-zero offset as it says; `dateUtc` UTC |
| a dry run | `model/content/trial.py` | the trial's own zone, the real package's by default, as its `tracked_change_date` is; its stamps use it |
| `pkg.write_times_on_save` | a plain attribute on `OpcPackage`, False | D4 |
| `save.write_core_times` | `openpackaging/save.py`, from `save_package` after the check and before the hooks | `dcterms:modified` the save's time (`save._now()`, UTC) to the second; `created` where the core part has none and its name is not in the source store (`_from_source`); a `W3CDTF` written as `create_package` writes one, `xsi:type="dcterms:W3CDTF"` |
| `last_save.times` | `normalise.SaveReport.times` | the stamp, or None |

### 11.2 D4's precision, decided while building (Q3)

The draft recommended the minute, Word's. Built that way, a package made by `create_package` and saved with the opt-in in the same minute had `modified` before `created`: `create_package` writes `created` to the second (`2026-10-28T02:32:21Z`) and the stamp was `02:32:00Z`. The second, docx4j's and docx4j-core-ts's, has no such case, and makes the three ports alike; Q3 is restated with that and the recommendation is the second, which is what is built. Word's minute stays one constant away.

### 11.3 Departures, with why

1. **No process-wide fallback** (D2): docx4j's `BindingHandler.dateTimeZone` property has no counterpart; `TZ` is the process's.
2. **A `tzinfo` as well as a name** (D1): docx4j takes a `ZoneId`, core-ts an IANA id; Python's zones are objects, and a fixed offset (`datetime.timezone`) is a zone a caller may hold.
3. **`created` where the part was not read from the source**, not docx4j's "where the package is new" (`isNew()`): a `create_package` package has a memory store, so "new" is the core part's provenance here. A loaded file whose core part has no `created` keeps none, as docx4j's does; core-ts's further case, a new package with no core part given one, does not arise, `create_package` making one.
4. **The opt-in writes the second, not the minute** (11.2).

### 11.4 Kept as it was

The text model's dates (a revision's and a comment's in the `markup` view's detail, the metadata's) are given as written, the rule's point 1. `tracked_change_date` is used as given, to the second; "now" is to the minute, for a revision and a comment alike (point 5, already so since CR-010 section 12.9). `create_package`'s `created` and `modified` (to the second) are unchanged (Q10).

### 11.5 Tests and the numbers

- `tests/content/test_date_time_zone.py`, 24: the fallback chain; six refusals; nothing saved, a dry run's own zone; docx4j's revision case in two process zones, the unset reading in two, an explicit offset; a comment's `w:date` in the package's zone and `w16cex:dateUtc` in any; a revision and a comment stamped on Tokyo's clock in two process zones, the unset bytes in two; `PROCESS_ZONE` against `astimezone()` at all 17,520 half hours of 2026 in Sydney, New York and Lord Howe (a half-hour daylight change), both ways and both folds; `xml_date`'s old formula in three zones and twelve months.
- `tests/openpackaging/test_write_times.py`, 8: off, a loaded core part byte for byte and not unmarshalled; on, `modified` with Word's `xsi:type`, the rest of the part unchanged, every save its own time, `create_package`'s `created` kept, a part made here given `created`, a loaded one with none left without, no core part none made, the dates reloaded.

### 11.6 The gates (Phase A at `cbafbc1`)

| gate | result |
|---|---|
| the fast suite (`-m "not slow"`) | **6,821 passed**, 0 failed: one more than the baseline's 6,788 at `3c63804` and Phase A's 32 (12.8 says where the one is not), counted from pytest's progress marks (a doubled `-q` hid the summary line, and the count read 7 skips, where every later run shows the same 8 as the baseline: six README fragments and the two python-docx modules; 12.8 has the runs with the summary) |
| acceptance (`scripts/acceptance.py --out out/acceptance-cr025a`, then `compare_acceptance.py` against `out/acceptance-cr020b/`) | **17 artefacts, 246 entries, 0 differ** beyond the dates |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | **50 of 50 canonically identical**, 0 differences, 0 skipped, 34 `boolean-spelling` |
| `scripts/threads.py` | **passes** |
| `ruff check` on the changed files | no new finding; the three there before (`opc_package.py` 2, `save.py` 1) left as they were |

## 12. Phase B implementation notes (2026-10-28)

Phase B as built, D5 to D14 and section 5's oracle, at the coordinator's brief. docx4j at `c3e9920e5` was read for behaviour (`FieldEvaluators`, `FieldContext`, `FieldSources`, `FieldInstruction`, the text model's `FieldValues` and `Extractor`, and the four test classes), and run, from its own build, as an oracle (12.3). Phase C is not this phase's.

**Commits.** `23fa5ce`, Phase B (12.1 to 12.7; with it the dry run's own zone, 11.1, and the census rows of 12.6); and this section's, after it.

### 12.1 What landed

| piece | where | what |
|---|---|---|
| `FieldInstruction` | `model/fields/instruction.py`, new, in `KEEP` | docx4j's, accessor for accessor (`name`, `parameter_text`, `parameters` a copy, `argument`, `has_switch`, `switch_value`, `switch_values`, `index_of_switch`, `value_at`; `split_parameters`, `unquote`, `FieldInstruction.parse`) |
| `FldSimpleModel` | `model/fields/formatting.py` | a subclass of `FieldInstruction`, keeping `parameter_string`; `apply_formatting_switch` takes either |
| `format_date` | `model/fields/formatting.py` | docx4j's `formatDate(instruction, date, lang, zone)` over a wall clock: the `\@` picture, else Java's short date (`M/d/yy`) for a `*DATE*` field and short time (`h:mm a`) for another |
| `FieldContext`, `FieldUser` | `model/fields/context.py`, new, in `KEEP` | D5, D10; the zone (`zone`, the override, else the package's, else `PROCESS_ZONE`), `clock` and `now()`, `language`, `user`, `section`, `document`, `computes` and `evaluates()` (this port's, D12), `sequences`, `bookmarks`, `note_numbers()`, `read_bookmarks()`, `properties()`, `variables()` |
| the registry | `model/fields/evaluators.py`, new, in `KEEP` | `register(name, evaluator, formats_itself=)`, `get`, `has`, `formats_itself`, `names`, `evaluate`; docx4j's 23 names (12.2) |
| the sources | `model/fields/sources.py`, new, in `KEEP` | `note_numbers(document)` and `bookmark_texts(document)` over docx4j's walk; `parse_w3cdtf` (D8, D9) |
| `docx4j_py.model.fields` | `__init__.py` | exports `FieldInstruction`, `FieldContext`, `FieldUser`, `evaluate`, `register`, `format_date`, `parse_w3cdtf`, `split_parameters`, `unquote` beside the old names |
| `FieldValues` | `model/text/field_values.py` | a thin consumer: `COMPUTED` docx4j's 20 names, a `FieldContext` per reading over the walk's own `w:document` (nothing more is read), the section per field, the switches applied once unless the evaluator formats itself, an exception keeping the cached result |
| the walk | `model/text/extractor.py` | `_section_of(story)` (the main story's sections ended, plus one; None elsewhere) passed to each computed field, complex and simple; the user from the options |
| `TextOptions(field_user=)` | `model/text/options.py` | D10: a `FieldUser`, a name or a mapping of `name`, `initials`, `address`; anything else `TextOptionError` (`field_user`, REQ-506, `invalid`); beyond section 10, so not in `to_dict()`, as `image_describer` is not |

### 12.2 The evaluators as registered

`DOCPROPERTY`, `DOCVARIABLE`; `AUTHOR`, `COMMENTS`, `KEYWORDS`, `LASTSAVEDBY`, `SUBJECT`, `TITLE`, `REVNUM` (an argument is the value); `CREATEDATE`, `SAVEDATE` (docx4j's `CoreDate`, formatting itself); `SEQ`, `REF`, `QUOTE`; `DATE`, `TIME`, `PRINTDATE` (the clock, formatting themselves); `INFO`, `SECTION`, `NOTEREF`, `USERNAME`, `USERINITIALS`, `USERADDRESS`. The text model computes 20 of the 23, docx4j's `COMPUTED`: all but the three clock fields (D14). `KEYWORDS` now reads through `DocPropertyResolver` (`Keywords`, a core property since docx4j's #277), as docx4j's does, where the old `FieldValues._keywords` read the core part itself; `tests/openpackaging/test_private_readers.py`'s census loses that caller (12.6).

### 12.3 The oracle, case by case

**Word's results** (`results.txt`, Jason's F9 run), replayed through the text model, the process in Sydney's zone, `field_user=FieldUser("Author A", "AA")`; and again with the package's zone Sydney and the process in New York, Kolkata or UTC (`tests/text/test_field_probe.py`):

| Case | Field | Word 365 | here, `fields: update` | |
|---|---|---|---|---|
| 1 | `CREATEDATE` | 2/3/2024 3:05:00 PM | 2/3/2024 3:05:00 PM | Word's |
| 2 | `CREATEDATE \@ "yyyy-MM-dd HH:mm"` | 2024-02-03 15:05 | 2024-02-03 15:05 | Word's |
| 3 | `SAVEDATE` | 6/7/2025 6:09:00 PM | 6/7/2025 6:09:00 PM | Word's |
| 4 | `SAVEDATE \@ "d MMMM yyyy"` | 7 June 2025 | 7 June 2025 | Word's |
| 5 | `REVNUM` | 7 | 7 | Word's |
| 6 | `INFO Title` | Probe Title | Probe Title | Word's |
| 7 | `INFO Author` | Probe Author | Probe Author | Word's |
| 8 | `INFO` | Error! Unknown info keyword. | OLD | cached (D11, docx4j's D5) |
| 9 | `INFO Nope` | Error! Unknown info keyword. | OLD | cached (D11) |
| 10 | `SECTION` | 1 | 1 | Word's |
| 11 | `QUOTE "hello world"` | hello world | hello world | Word's |
| 12 | `QUOTE` | (empty) | (empty) | Word's |
| 13 | `REF bm1` | Bookmarked text, bold | Bookmarked text | Word's text; the bold is a write-back's, pinned |
| 14 | `REF bm1 \* Upper` | BOOKMARKED TEXT | BOOKMARKED TEXT | Word's |
| 15 | `REF nope` | Error! Reference source not found. | OLD | cached (D11) |
| 16 | `REF bm1 \h` | Bookmarked text | Bookmarked text | Word's |
| 17 | `NOTEREF bmNote` | 1 | 1 | Word's |
| 18 | `NOTEREF bmNote \f` | 1, in `FootnoteReference` | 1 | Word's text; the style is a write-back's, pinned |
| 19 | `NOTEREF nope` | Error! Bookmark not defined. | OLD | cached (D11) |
| 20 | `USERNAME` | Author A | Author A (no user: OLDUSER) | Word's with the user |
| 21 | `USERINITIALS` | AA | AA (no user: OLDUSER) | Word's with the user |
| 22 | `USERADDRESS` | (empty) | (empty) (no user: OLDUSER) | Word's with the user |
| 23 | `KEYWORDS` | alpha, beta | alpha, beta | Word's |
| 24 | `SECTION` | 2 | 2 | Word's |
| 25 | `CREATEDATE \@ "yyyy-MM-dd"`, simple | 2024-02-03 | 2024-02-03 | Word's |

21 of 25 are Word's; the 4 where Word writes an error text keep the cached result, as docx4j's do. Without a user the three user fields keep `OLDUSER`, what Word itself showed before F9. At `3c63804`, 12 of the 25 differed from Word's and docx4j's (section 3.2): cases 1 to 4 and 25 (the lexical form), 6, 7, 10, 12, 17, 18 and 24 (cached).

**docx4j's own extraction**, run here from docx4j's build of `c3e9920e5` (the `17.3.5-SNAPSHOT` jars in its `target/`; `ProbeText.java` and `ProbeMd.java` in the fixture directory; the JVM in Sydney's zone): byte for byte, `Texts.plainText` of the probe updated, updated with `setDateTimeZone(Asia/Tokyo)`, and cached, and of Word's save cached; and `MarkdownExporter` of the probe updated. Word's save updated agrees too (the save's moment in cases 1 to 4 and 25, revision 2: Word's Save As rewrote them), checked by hand and not kept. So the text model is at parity with docx4j's on this file in every mode tried.

**Word's save, read cached,** is `results.txt` in all 25 cases.

### 12.4 The tests

- `tests/text/test_field_probe.py`, 16 (12.3), the fixtures in `tests/fixtures/field-probe-cr067/` (three files copied from docx4j at `d6ca8dbee`, their git blob ids docx4j's; five outputs of docx4j and the two Java files that wrote them; `README.md`, `SHA1SUMS`).
- `tests/text/test_field_evaluators.py`, 32: `FieldEvaluatorsTest`'s six cases that apply (`fieldUpdaterUpdatesTheTierOneFieldsInPlace` and `fieldUpdaterUpdatesItsNamesOnly` are `FieldUpdater`'s), with its `fr-FR` date pinned in English (D7) and the defaults of `DATE` and `TIME` with no picture; a custom mark and a deletion not counted by `NOTEREF`; `PackageDateTimeZoneTest.theDateFieldsShowThePackagesZone` and the context's fallback; D9's four nominal forms in three package zones; four offset forms; no date, no value, an argument ignored; an empty picture; D12; D14; `SECTION` in the main story and a note; `field_user` refused; a read leaving the package byte for byte.
- `tests/text/test_field_instruction.py`, 71: `FieldInstructionTest`, its 31 shapes through both classes, held to what `FldSimpleModel` read at `3c63804` (computed from that commit's source into the test's table).
- `tests/text/conftest.py` gains the `zone` fixture, as `tests/content/` and `tests/revisions/` have it.

### 12.5 Departures, with why

1. **`INFO` asks the consumer's set** (D12, Q9). docx4j's `INFO` asks the registry, so its extraction evaluates `INFO PrintDate` (and `INFO Date`, `INFO Time`) at the clock; here they keep the cached result. Offered to docx4j through the portfolio (12.7).
2. **No language** (D7): the month and day names are English, so docx4j's `fr-FR` case is pinned as English. No language reaches docx4j's inspector either, so extraction agrees.
3. **The forms Word has not measured** (D9, Q8): nominal here whatever the package's zone; docx4j reads them in the JVM's zone and shows them in the package's. With no package zone set the two agree.
4. **A user for extraction** (D10, Q7): an extension; docx4j's text model has none, and with none the two agree.
5. **Not ported:** `FieldContext`'s part, file name, external resolver and `sectionOf`, which only docx4j's `FieldUpdater` and writers use; `FieldUpdater` itself (no write-back here).
6. **`format_date`'s short time is `h:mm a` with a space.** Java 20 and later (CLDR 42) put a narrow no-break space before the AM/PM marker in `getTimeInstance(SHORT)`; docx4j's output for a `TIME` with no picture depends on its JDK. The text model never reaches it (D14).

### 12.6 Moved expectations

None of the text model's goldens or ratchet moved: `tests/text/conformance.tsv` (the harness's report), the goldens, `test_docproperty_case.py`, `test_texts_phase2.py`'s `fields: update` cases and the agreement test pass unchanged. No fixture under `tests/` or `samples/` held a tier-1 field, a `CREATEDATE` or `SAVEDATE`, or a `QUOTE` with no argument before this phase (a scan of every instruction, section 7). Two censuses changed, both lists of the source rather than expectations of output: `tests/openpackaging/test_private_readers.py`'s `CALLERS` loses `model/text/field_values.py::FieldValues._keywords` (12.2), and its docstring says so; and CR-021's `tests/content/test_integer_readers.py` lists the new `int(` calls, the `SEQ` reset (moved from `field_values.py` to `evaluators.py`, its reason `TEXT` as before) and the nine of `parse_w3cdtf`, under a new reason, `W3CDTF` (a core property's date, no WordprocessingML integer attribute).

### 12.7 Found, and not done

1. **docx4j's `INFO` reach** (12.5 item 1): for docx4j, through the portfolio.
2. **The portfolio's section 2, python column,** says this repository read `CREATEDATE \@` "parsed nominally, the `Z` ignored"; measured, it gave the lexical form with or without `\@` (section 3.2). For the portfolio session's record.
3. **The portfolio's registry**: `python/CR-025.A` and `python/CR-025.B` want `done`, with sections 11 and 12 as their notes; the `tasks.yaml` rule (CLAUDE.md) is not this session's to apply.
4. **Two `conftest` modules on one command line.** `pytest tests/content/X.py tests/text/Y.py` imports `tests/text/conftest.py` as `conftest` first, so a `tests/content` module's `from conftest import ...` fails; the whole suite collects in an order that avoids it. Pre-existing; the new `tests/content/test_date_time_zone.py` imports nothing from `conftest` so as not to depend on it.
5. **Phase C**: CHANGELOG, README and `docs/pypi-readme.md`, CLAUDE.md (its status line, and the five new files in its list of hand-written ones), the index's row, CR-014's note (section 9).

### 12.8 The gates (at `23fa5ce`)

| gate | before (`3c63804`) | after |
|---|---|---|
| the fast suite (`-m "not slow"`) | 6,788 passed, 8 skipped (the brief's figure) | **6,940 passed, 8 skipped**, 0 failed, 1,457 deselected (3 min 26 s) |
| the full suite (`.venv-fork/bin/python -m pytest`) | not run here (8,245 passed, 8 skipped at CR-020's `5903d7d`) | **8,397 passed, 8 skipped**, 0 failed (5 min 0 s) |
| the count | --- | 6,940 is one more than 6,788 and the 151 new tests. pytest collects 6,947 items here (the python-docx module's skip is a collection skip, so 6,948 outcomes); outside the four golden files that parametrize over the sibling checkouts, a worktree of `3c63804` collects exactly these ids less the 151, and those four files and their inputs did not change, so this checkout held 6,796 items at `3c63804`, where 6,788 and 8 imply 6,795. The one was not found; no expectation moved for it |
| the new tests | --- | 151: Phase A's 32 (`test_date_time_zone.py` 24, `test_write_times.py` 8), Phase B's 119 (`test_field_probe.py` 16, `test_field_evaluators.py` 32, `test_field_instruction.py` 71) |
| `tests/test_readme.py` | --- | **44 passed, 6 skipped** (the six fences marked fragments) |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | --- | **50 of 50 canonically identical**, 0 differences, 0 skipped, 34 `boolean-spelling` |
| `scripts/threads.py` | --- | **passes** (no configuration lost or altered a document) |
| acceptance (`scripts/acceptance.py --out out/acceptance-cr025b`, `compare_acceptance.py` against `out/acceptance-cr020b/`) | --- | **17 artefacts, 246 entries, 0 differ** beyond the dates |
| `codegen/generate.sh --check` (a worktree at `23fa5ce` in this session's scratchpad) | --- | two generations byte identical, **316 modules** (CR-020's 311 and the five new); nothing generated differs from the commit, and `model/dates.py` and the four new `model/fields/` modules survive `clean.py` |
| `ruff check` on every Python file changed since `3c63804` | --- | no new finding; the four there before (`extractor.py` 1, `opc_package.py` 2, `save.py` 1) left as they were |
| docx4j's extraction of the probe (12.3) | 12 of 25 cases differ | **byte for byte** in five modes |

No Word check (section 5): the text model's output is no file, and the defaults leave saved files unchanged, which the acceptance comparison shows. `pkg.write_times_on_save` writes a core part in the form `create_package` has always written, which Word has opened since 0.1.0.

### 12.9 For the user to decide

Every question of section 8 is open; each is built as recommended, so an answer that agrees changes nothing:

1. **Q1, the names**: `pkg.date_time_zone` (a `tzinfo` or an IANA name) and `pkg.effective_date_time_zone`, after docx4j-core-ts's (section 4.5).
2. **Q2**: no process-wide fallback between the package's zone and `TZ`.
3. **Q3, restated while building** (11.2): the opt-in's stamp to the second, the three ports alike; the minute, Word's, would want `create_package`'s to the minute too.
4. **Q4**: `pkg.write_times_on_save` alone, no per-save option.
5. **Q5**: `FieldInstruction` ported, `FldSimpleModel` its subclass, not deprecated with a warning.
6. **Q6**: `DATE`, `TIME`, `PRINTDATE` cached under `fields: update`.
7. **Q7**: the user from `TextOptions(field_user=)`, never from `pkg.author`.
8. **Q8**: a date alone or a time with no offset nominal whatever the package's zone, until docx4j's date-forms probe has run.
9. **Q9**: `INFO` evaluates only what the consumer computes; offer docx4j the same.
10. **Q10**: `create_package`'s dates unchanged, to the second.

### 12.10 The decisions (2026-10-28)

The user, on section 8: "CR-025's ten questions as recommended". Each keeps what Phases A and B built, so no code, test or
generated file changed and no gate moved from 11.6's and 12.8's. The coordinator re-ran the fast suite at `ed416fd` (6,940
passed, 8 skipped; the one test beyond the 151 new ones is CR-020 Phase C's README fence, which the 6,788 baseline
predates) and exercised the zone, the opt-in and the probe's dates by hand (London, Kolkata, Los Angeles; Sydney's
"2/3/2024 3:05:00 PM").

1. **Q1:** `pkg.date_time_zone` (a `tzinfo` or an IANA name, None for the process's) and `pkg.effective_date_time_zone`.
2. **Q2:** no process-wide fallback between the package's zone and `TZ`.
3. **Q3:** `write_times_on_save` stamps to the second, as docx4j and docx4j-core-ts do.
4. **Q4:** `pkg.write_times_on_save` alone, no per-save option.
5. **Q5:** `FieldInstruction` ported, `FldSimpleModel` its subclass, with no deprecation warning.
6. **Q6:** `DATE`, `TIME` and `PRINTDATE` keep their cached result under `fields: update`.
7. **Q7:** the user fields' user from `TextOptions(field_user=)`, never from `pkg.author`.
8. **Q8:** a date-only or offset-free value is read as written, whatever the package's zone, until docx4j's date-forms probe
   has run.
9. **Q9:** `INFO` evaluates only what the consumer computes. Offered to docx4j the same day (docx4j-12), whose `INFO` reaches
   the clock fields in extraction (12.5 item 1).
10. **Q10:** `create_package`'s dates unchanged, to the second.

## 13. Phase B2 implementation notes (2026-10-28)

Phase B2 as built, at the coordinator's brief after the user's order (13.1). docx4j at `c58e36e64` was read for behaviour: its CR-067 sections 12.2 and 12.3, the commit's diff (`FieldSources.parseW3CDTF`, `FieldEvaluators`, `FieldValues`) and `DateFormsProbeWordTest`. Phase C (the records) is not this phase's: no CHANGELOG, README, `docs/pypi-readme.md`, CLAUDE.md or index change was made, beyond `tests/README.md`'s rows for the new tests and fixtures.

**Commits.** `3fe390d`, Phase B2 (13.3 to 13.6: the readers, the normal form, the fixtures, the tests); and this section's, after it.

### 13.1 Why

**Q8 foresaw it.** Q8 was decided as recommended (12.10 item 8): "a date-only or offset-free value is read as written, whatever the package's zone, until docx4j's date-forms probe has run", the recommendation going on "then both ports follow Word". The portfolio's rule (`date_handling_across_ports.md` section 3.6 point 4) says the same. The probe has run (Jason, Word 365 version 2608 build 20326.20232, Sydney's zone, 2026-10-28; docx4j CR-067 section 12.2, committed at `c58e36e64`), so Q8's reading ends and Word's begins: the decision changes its meaning as it said it would, and items 1 and 2 needed no new decision.

**Item 3 is the user's.** "item 3 as recommended; go on CR-025 B2" (2026-10-28, relayed by the coordinator). The recommendation it accepts: a save normalises only the forms the source did not write as Word does, a date alone and a date-time with no offset in a `w:date` or a `w:fullDate`, writing them in Word's form, the nominal clock to the minute labelled `Z`; a value already carrying a `Z` or an offset is left exactly as it is, with no truncation, because "truncating well-formed values would move many saved files for no gain".

**What this repository did at `2d206c0`** (the coordinator's measurement, confirmed here; `TZ=Australia/Sydney`, `pkg.date_time_zone` Sydney):

| what | here at `2d206c0` | Word |
|---|---|---|
| D01, `CREATEDATE` of created `2024-02-03` | `2024-02-03 00:00:00` | `2024-02-03 11:00:00` |
| D02, `SAVEDATE` of modified `2025-06-07T08:09:10` | `2025-06-07 08:09:00` | `2025-06-07 18:09:00` |
| D04, D05, the same without `\@` | `2/3/2024 12:00:00 AM`, `6/7/2025 8:09:00 AM` | `2/3/2024 11:00:00 AM`, `6/7/2025 6:09:00 PM` |
| D03, D06, `PRINTDATE` | `OLD`, cached (Q6) | `cp:lastPrinted` |
| R01, `w:date="2024-02-03"` | `TrackedChange.date` None: no `xsd:dateTime`, so the exact parse keeps it as its text (`part.kept`, a `ConverterWarning`) | local midnight |
| B01, `w:fullDate="2024-02-03"` | `DatePicker.full_date` None, kept the same way | the picker on 3 Feb 2024 |
| a save of the probe, its document read | every `w:date` and `w:fullDate` as read, `last_save.normalised` empty | the nominal clock to the minute labelled `Z` |

### 13.2 What Word does (docx4j CR-067 section 12.2)

The probe, `date-forms-r1.docx`, holds three stored forms, a date alone (`2024-02-03`), a time with no offset (`2025-06-07T08:09:10`) and a UTC time (`2025-01-02T03:04:05Z`), each in a core property (created, modified, lastPrinted, read by `CREATEDATE`, `SAVEDATE` and `PRINTDATE` with `\@ "yyyy-MM-dd HH:mm:ss"`, D01 to D03, and without a switch, D04 to D06), in a revision's `w:date` (three `w:ins`, R01 to R03), and in a date control's `w:fullDate` (B01 to B03). Jason opened it in Word 365 in Sydney's zone, pressed Ctrl+A and F9, read the balloons and the pickers, and saved it (`date-forms-r1-word365.docx`):

| Case | Word 365 | Taken as |
|---|---|---|
| D01, D04: created `2024-02-03` | `2024-02-03 11:00:00`; `2/3/2024 11:00:00 AM` | a date alone is **UTC midnight** (11:00 at UTC+11) |
| D02, D05: modified `2025-06-07T08:09:10` | `2025-06-07 18:09:00`; `6/7/2025 6:09:00 PM` | a time with no offset is **UTC** (18:09 at UTC+10), to the minute |
| D03, D06: lastPrinted `2025-01-02T03:04:05Z` | `2025-01-02 14:04:00`; `1/2/2025 2:04:00 PM` | `PRINTDATE` on F9 is `cp:lastPrinted` (docx4j's D14) |
| R01: `w:date="2024-02-03"` | `3/02/2024 12:00:00 AM` | a date alone is **local midnight** |
| R02: `2025-06-07T08:09:10` | `7/06/2025 8:09 AM` | the local clock as written |
| R03: `2025-01-02T03:04:05Z` | `2/01/2025 3:04 AM` | the local clock, the `Z` ignored (CR-010 section 12.9) |
| B01 to B03: `w:fullDate` in the three forms | the stored text; the picker on 3 Feb 2024, 7 Jun 2025, 2 Jan 2025 | the nominal day in every form |
| Word's save | every `w:date` and `w:fullDate` `2024-02-03T00:00:00Z`, `2025-06-07T08:09:00Z`, `2025-01-02T03:04:00Z`; lastPrinted `2025-01-02T03:04:00Z`; created and modified the save's moment | the nominal clock to the minute labelled `Z`, no zone converted |

So the three places disagree on a bare value: a core property's is UTC; a revision's and a control's are the local clock or the nominal day.

### 13.3 What landed

| piece | where | what |
|---|---|---|
| `values.date_time` | `model/values.py` (new section "Dates") | a `w:date` or `w:fullDate` as Word reads its form: an `xsd:dateTime` with a `Z`, an offset or neither is that clock (an `XmlDateTime`); a **date alone** is midnight on that day, its offset (if any) kept, as docx4j's `XMLGregorianCalendar` reads one; anything else None (`2025-06-07T08:09`, a month alone, no date). Takes the model's value, a string (trimmed) or a `datetime`. The one reader of the form, as `values.py` is for every other kind of value: the parse keeps what was written and the reader applies Word's rule |
| `values.date_time_normal_form` | the same | what a save writes for a `w:date` or `w:fullDate`: for a date alone or a date-time with no offset, the nominal clock to the minute labelled `Z`; None (leave it) for a value with a `Z` or an offset, or no date |
| item 1: `parse_w3cdtf` | `model/fields/sources.py` | docx4j's D13: a core property's value with no offset, and a date alone (also `2024`, `2024-02`), is **UTC**. It now returns an aware UTC `datetime` (or None), docx4j's `Date`, where Phase B returned `(datetime, has_offset)`: the flag had no other use, and the function is unreleased |
| `CREATEDATE`, `SAVEDATE` | `model/fields/evaluators.py`, `_core_date` | one path for every form: the instant to the minute, shown in the context's zone; D9's nominal branch is gone |
| item 2: the revision and comment readers | `model/revisions/change.py`: `_calendar` is `values.date_time`; `word_date`, `utc_date`, `date_of` over it | a date alone in a `w:date` is **local midnight** in the package's zone (else the process's): `TrackedChange.date` (every kind, through `date_of`), `tracking.date_of`, `Comment.creation_date` (all three readers go through `word_date`). The other forms read as before: no offset the local clock as written, a `Z` the local clock, a non-zero offset as it says. A `dateUtc` still wins; a date alone there is UTC midnight, as docx4j's `dateOf` reads one (`setTimezone(0)`) |
| the date control's reader | `model/customxml/kinds.py`, `DatePicker.full_date` | through `values.date_time`: a date alone is midnight, naive, the nominal day (B01), where it was None. The other forms as before (naive with no offset, aware with one) |
| item 3: `DATE_ATTRIBUTES` | `openpackaging/normalise.py` | the 25 elements and their attribute (13.4), a constant, as `UNION_INTEGER_ATTRIBUTES` is, held to the model by a test |
| `_date_rows` | the same; in `_render_rows` with Phase B's and CR-021's render rows | found by name in each render (`normalise_render`, so in a normalising save only), and in a source by `normalise_tree` (so `canon.compare(..., normalise=True)`) and `rendered_rows` (`scripts/roundtrip.py`): a value `date_time_normal_form` rewrites is written so, one `Normalised` row each, in `last_save.normalised`. The render rows are the right home: an offset-free value binds as it is and is in no `part.kept`, and a render row leaves the tree as read (13.5). 0.008 ms over `Symbols.docx`'s main part |
| `canon.compare` | `scripts/canon.py` | its docstring names the new rows; nothing else changed there |

### 13.4 The attributes, from the model

`tests/content/test_date_forms.py` derives the table from the model (every element name whose bound classes declare a `w`-namespace `xsd:dateTime` attribute) and holds `DATE_ATTRIBUTES` to it. **The `w` namespace has exactly two `xsd:dateTime` attributes, `w:date` and `w:fullDate`**, on these 25 elements, every class each name binds typing the attribute `XmlDateTime` alone:

- `w:date` on the revision elements of `w`: `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, `w:moveFromRangeStart`, `w:moveToRangeStart`, `w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange`, `w:tblPrExChange`, `w:trPrChange`, `w:tcPrChange`, `w:cellIns`, `w:cellDel`, `w:cellMerge`, `w:numberingChange`, `w:customXmlInsRangeStart`, `w:customXmlDelRangeStart`, `w:customXmlMoveFromRangeStart`, `w:customXmlMoveToRangeStart` (21; `w:ins` and `w:del` include a paragraph mark's, a row's and those in an `m:r` or `m:ctrlPr`);
- `w:date` on `w14`'s two, `w14:customXmlConflictInsRangeStart` and `w14:customXmlConflictDelRangeStart`, root elements in the model and in no content model (a typed document has nowhere to hold them, so a test covers them over a render);
- `w:date` on a comment, `w:comment`;
- `w:fullDate` on a date control's `w:date`.

42 classes carry a `w:date` (the revision elements' choice classes, `w14`'s and `CommentsComment`); a test reads a date alone through `date_of` on every one. Not in the table: `w16du:dateUtc` and `w16cex:dateUtc`, which are UTC and which Word writes only with a `Z`, and every other namespace's `xsd:dateTime` (`dcterms`, `cp:lastPrinted`, inkml's, the 2013 command's `dt`): the brief and the probe are about `w:date` and `w:fullDate`.

### 13.5 The normal form, and `w:fullDate`'s

- **`w:fullDate` takes `w:date`'s form.** Word's save of B01 wrote `2024-02-03T00:00:00Z` and of B02 (`2025-06-07T08:09:10`) `2025-06-07T08:09:00Z`: the nominal clock to the minute labelled `Z`, the seconds of an offset-free value cut, exactly as for R01 and R02. So one rule serves both, and an offset-free `w:fullDate` with seconds is cut to the minute, as Word's save cut it.
- **Word reads the form written as it read the source.** A revision's `2024-02-03T00:00:00Z` is local midnight (the `Z` ignored) as `2024-02-03` was; `2025-06-07T08:09:00Z` is 8:09 local as `2025-06-07T08:09:10` was, to the minute, which is what Word holds (its balloon showed 8:09 AM and its save wrote `:00`). A control's day is the same. Read back here, R02's instant loses its 10 seconds (13.6), as Word's own save loses them.
- **Left as read**: a value with a `Z` or an offset, its seconds and fraction included (`2025-01-02T03:04:05Z`, `...+10:00`, `...05.123Z`); a date alone with one (`2024-02-03Z`), which carries an offset by the user's rule though Word writes no date alone (unmeasured); anything that is no date (`2025-06-07T08:09`).
- **CR-018's discipline throughout**: only a part the save re-marshals is normalised (a part never read is the source's bytes, dates included, which the probe's replay holds); the render is rewritten and never the tree, so `save(normalise=False)` and `part.xml` outside a save write the source's spelling; a value set in code with no offset is written in the form too, as Phase B's colours are; `canon.compare(..., normalise=True)` writes the same rows into the source, so the round trip stays at zero differences beyond `word-normal-form`.

### 13.6 The oracle, case by case

`tests/text/test_date_forms_probe.py`, the fixtures in `tests/fixtures/date-forms-cr067/` (`date-forms-r1.docx`, `date-forms-r1-word365.docx`, `README.txt`, `results.txt` copied unedited from docx4j at `c58e36e64`, their git blob ids docx4j's; `README.md`, `SHA1SUMS`). docx4j's `DateFormsProbeWordTest`, its updater's Part 1 through the text model's `fields: update`; in Sydney's zone by the process, and by `pkg.date_time_zone` with the process in New York, Kolkata and UTC:

| Case | Word 365 | here at `2d206c0` | here now | |
|---|---|---|---|---|
| D01 `CREATEDATE \@`, date alone | `2024-02-03 11:00:00` | `2024-02-03 00:00:00` | `2024-02-03 11:00:00` | Word's |
| D02 `SAVEDATE \@`, no offset | `2025-06-07 18:09:00` | `2025-06-07 08:09:00` | `2025-06-07 18:09:00` | Word's |
| D03 `PRINTDATE \@`, `Z` | `2025-01-02 14:04:00` | `OLD` | `OLD` | cached (Q6), as docx4j's text model keeps it |
| D04 `CREATEDATE` | `2/3/2024 11:00:00 AM` | `2/3/2024 12:00:00 AM` | `2/3/2024 11:00:00 AM` | Word's |
| D05 `SAVEDATE` | `6/7/2025 6:09:00 PM` | `6/7/2025 8:09:00 AM` | `6/7/2025 6:09:00 PM` | Word's |
| D06 `PRINTDATE` | `1/2/2025 2:04:00 PM` | `OLD` | `OLD` | cached (Q6) |
| R01 `w:date`, date alone | 3/02/2024 12:00:00 AM | None | `2024-02-02T13:00:00Z` | Word's (local midnight, UTC+11) |
| R02, no offset | 7/06/2025 8:09 AM | `2025-06-06T22:09:10Z` | the same | Word's |
| R03, `Z` | 2/01/2025 3:04 AM | `2025-01-01T16:04:05Z` | the same | Word's |
| B01 `w:fullDate`, date alone | the picker on 3 Feb 2024 | None | `2024-02-03 00:00` (naive) | Word's day |
| B02, no offset | 7 Jun 2025 | `2025-06-07 08:09:10` | the same | Word's day |
| B03, `Z` | 2 Jan 2025 | `2025-01-02 03:04:05+00:00` | the same | Word's day |
| the save, R01 and B01 | `2024-02-03T00:00:00Z` | `2024-02-03` | `2024-02-03T00:00:00Z` | Word's |
| the save, R02 and B02 | `2025-06-07T08:09:00Z` | `2025-06-07T08:09:10` | `2025-06-07T08:09:00Z` | Word's |
| the save, R03 and B03 | `2025-01-02T03:04:00Z` | `2025-01-02T03:04:05Z` | `2025-01-02T03:04:05Z` | as read, the user's decision (13.7 item 1) |
| the save, lastPrinted | `2025-01-02T03:04:00Z` | as read | as read | the core part is not unmarshalled (13.7 item 2) |
| the save, created and modified | the save's moment | as read | as read | Word's Save As makes a new document; here D4's opt-in |

The rest of the replay: `results.txt` holds the six field cases; Word's save holds the nominal clock to the minute; this save's four rows are `last_save.normalised`, `canon.compare(..., normalise=True)` counts them as `word-normal-form` and nothing else; `save(normalise=False)` writes the probe's spelling; a save that never reads the document writes it byte for byte; this save reloaded reads R01 and R02 as Word's save reads them, R03 but for the seconds Word cut, and nothing is kept as text; `parse_w3cdtf`'s three forms, and the per-call zone (Tokyo: `2024-02-03 09:00:00`, `2025-06-07 17:09:00`); and `PRINTDATE` pinned as the registry's clock (13.7 item 3). 19 tests.

**docx4j's own extraction of this probe was not run.** Phase B ran docx4j's build of `c3e9920e5` (12.3); docx4j's `target/` jars are now a build of its uncommitted phase-2 work (`FieldSources`, `FieldValues`, `Extractor` modified in its tree), so they are no clean oracle for `c58e36e64`, and building docx4j is not this session's. The oracle is Word's record and docx4j's replay of it.

### 13.7 Departures, with why

1. **A value with a `Z` or an offset is not cut to the minute** (R03, B03): Word's save writes `03:04:00Z` where this one keeps `03:04:05Z`. The user's decision: truncating well-formed values would move many saved files for no gain. Word reads `:05` and `:00` alike (to the minute).
2. **The core part's dates are not touched by a save**: Word's save cut lastPrinted to the minute and made created and modified the save's moment; here a part never unmarshalled is the source's bytes (the engine's rule), and `pkg.write_times_on_save` (D4, off by default) writes `modified` only. Item 1 is a reading, not a write.
3. **docx4j's D14 is not ported**: since `c58e36e64` docx4j registers `PRINTDATE` as `cp:lastPrinted` (`CoreDate("LastPrinted")`), measured; this CR's D14 registers it at the context's clock, as docx4j did at `c3e9920e5`. Under `fields: update` both text models keep it cached (Q6; D03 and D06), so extraction agrees; only a caller of `evaluate` sees the difference, pinned by the replay. Not in the brief's three items; for the user (13.10 item 1).
4. **A save normalises, which docx4j's does not set out to**: docx4j has no normal form for these values (its CR-067 section 12.2 records Word's save, not a change to its own), and what its JAXB marshaller writes for a date alone was not checked here. CR-018's normaliser is this port's, opt-out, on the user's rule that a save writes what Word reads exactly as it read the source (CR-018 section 4.7).
5. **A date alone in `dateUtc` is UTC midnight**, as docx4j's `dateOf` reads one, where it was None (and the `w:date` beside it was used). Word writes neither; it follows docx4j rather than a measurement.
6. **A date alone with an offset** (`2024-02-03Z`, `2024-02-03+05:00`) is left as read by the save, by the user's rule for a value carrying one, though Word writes no date alone; read as midnight there (the `Z` as the local clock, a non-zero offset as it says). Unmeasured.

### 13.8 Moved expectations

- **`tests/text/test_field_evaluators.py`**, 32 tests to 29. `test_the_core_property_dates_revision_and_info`: `parse_w3cdtf` returns the aware instant, not `(instant, True)` (13.3). `test_a_date_alone_or_a_time_with_no_offset_is_nominal_in_any_zone` (12 cases: four forms in three package zones, every one the wall clock as written) is replaced by `test_a_date_alone_or_a_time_with_no_offset_is_utc` (9 cases: the four forms in the process's zone, Sydney, and two forms each in Tokyo and New York with `2024` there too): `2025-06-07T08:09:10` moves from `6/7/2025 8:09:00 AM` to `6:09:00 PM` in Sydney, `5:09:00 PM` in Tokyo, `4:09:00 AM` in New York; `2024-02-03` from `12:00:00 AM` to `11:00:00 AM`, `9:00:00 AM`, and `2/2/2024 7:00:00 PM`; `2024` and `2024-02` to 11:00 in Sydney. The module docstring and `tests/README.md`'s row say so.
- **The readers' results** for a date alone (none of them held by a test before): `TrackedChange.date`, `tracking.date_of`, `Comment.creation_date` and `DatePicker.full_date` give an instant or a midnight where they gave None.
- **CR-011's goldens and the ratchet**: nothing moved. `tests/text/conformance.tsv`, `agreement-docx4j.tsv`, the goldens, `test_field_probe.py` (Phase B's 16, all with `Z` values) and `test_texts_phase2.py`'s `fields: update` cases pass unchanged, and no fixture under `tests/` or `samples/` holds a bare core-property date (a scan of every `docProps/core.xml`, 13.9).
- **No census moved**: `parse_w3cdtf`'s nine `int(` calls are textually as CR-021's `test_integer_readers.py` lists them; `tests/openpackaging/test_private_readers.py` is unchanged.

### 13.9 The gates (Phase B2 at `3fe390d`)

| gate | before (`2d206c0`) | after |
|---|---|---|
| the fast suite (`-m "not slow"`) | 6,940 passed, 8 skipped (the brief's figure) | **6,998 passed, 8 skipped**, 0 failed, 1,457 deselected (3 min 42 s); the 13 warnings are the modules' before, none from a new test |
| the count | --- | 6,940 less Phase B's 32 `test_field_evaluators.py` tests, plus 90: that module's 29, `test_date_forms.py` 42, `test_date_forms_probe.py` 19 |
| the full suite | 8,397 passed, 8 skipped (12.8, at `23fa5ce`) | **8,455 passed, 8 skipped**, 0 failed (5 min 0 s): 8,397 and the same 58 |
| `tests/test_readme.py` | 44 passed, 6 skipped (12.8) | **44 passed, 6 skipped** (the six fences marked fragments) |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50, 0 differences, 34 `boolean-spelling` | **50 of 50 canonically identical**, 0 differences, 0 skipped, 34 `boolean-spelling`, **0 normalised**: no part of the corpus holds a date alone or an offset-free `w:date` or `w:fullDate`, so no new row shows under the tolerated categories |
| `scripts/threads.py` | --- | **passes** (no configuration lost or altered a document) |
| acceptance (`scripts/acceptance.py --out out/acceptance-cr025b2`, `compare_acceptance.py` against `out/acceptance-cr020b/`) | --- | **17 artefacts, 246 entries, 0 differ** beyond the dates; each set holds 106 `w:date` and 1 `w:fullDate`, every one with a `Z`, so none was normalised |
| the corpus and fixtures scanned, parsed (every `.docx` under `samples/` and `tests/`; docx4j-core-ts's `test/` and docx4j's test resources, 996 files) | --- | before the probe was copied, 419 files here with 1,023 `w:date` and 2 `w:fullDate`, every one with a `Z`; in the siblings 7,916 and 27 with a `Z`, and the only date-only and offset-free values anywhere are the probe's own four. Of every `docProps/core.xml` under `samples/`, `tests/` and `out/acceptance-cr020b/` (254 `created`, 254 `modified`, 32 `lastPrinted`), the probe's `created` and `modified` are the only bare ones |
| `ruff check` on the changed files | --- | no new finding; `scripts/canon.py`'s one (F401, `typing.Any`) was there before |
| `codegen/generate.sh --check` | --- | not needed: no file was added under `docx4j_py/` |

**No saved file of the corpus or the acceptance set changed.** The round trip writes 0 normalised values and the acceptance comparison 0 differences, because none of those files holds a value item 3 covers. The new normal form does change a save of a file that holds one, such as the probe: its four values are written in exactly the form Word's own save wrote for them (`date-forms-r1-word365.docx`, which Word wrote and so reads). So no acceptance artefact needs a new Word check. Because the normal form is a change to marshalled output (CLAUDE.md's rule), this save of the probe is written for one, should the user want it: `out/date-forms-cr025b2/date-forms-r1-saved.docx` (13.10 item 4).

### 13.10 For the user to decide

1. **docx4j's D14** (13.7 item 3): register `PRINTDATE` as `cp:lastPrinted` (to the minute, in the zone, the cached result where the document was never printed), as docx4j does since `c58e36e64` and as Word's F9 showed (D03, D06), in place of this CR's D14 clock? Recommendation: yes, docx4j's, a one-line registration; extraction is unchanged, since the text model keeps `PRINTDATE` cached (Q6). Not built: it is not one of Phase B2's three items.
2. **The python-docx facade's core dates.** `docx4j_py.docx`'s `core_properties.created`, `modified` and `last_printed` give a naive `datetime` for a bare value (`2024-02-03` as `2024-02-03 00:00`), where python-docx 1.2.0's `_parse_W3CDTF_to_datetime` gives every form aware in UTC, which is also Word's reading (D13). Recommendation: align them (`docx/coreprops.py`'s `_datetime_of`), a separate small change with its own test against python-docx's behaviour. Not built: the brief's item 1 is the field evaluation.
3. **A date alone with a `Z` or an offset** in a `w:date` or `w:fullDate` (13.7 item 6): left as read by the save, under the rule for a value carrying one. Recommendation: keep it; Word writes none, and the rule is the user's.
4. **A Word check** of this save of the probe, `out/date-forms-cr025b2/date-forms-r1-saved.docx` (sha1 `7984183607d83c01572990a2cdd2a7305503a9f9`), optional: it opens without a repair prompt; R01 to R03's balloons read 3/02/2024 12:00 AM, 7/06/2025 8:09 AM and 2/01/2025 3:04 AM in Sydney's zone (All Markup); B01 to B03's pickers are on 3 Feb 2024, 7 Jun 2025 and 2 Jan 2025. Recommendation: not needed for release, since the forms written are Word's own save's; worth the five minutes if the user wants the normal form seen in Word.

### 13.11 The decisions on 13.10 (2026-10-28), and the two changes they made

The user: "four B2 points in CR-025 section 13.10, all as recommended". The coordinator re-ran the fast suite at `a78c453`
(6,998 passed, 8 skipped) and the probe by hand (the process in New York, the package in Sydney: D01, D02, D04, D05 Word's;
R01 `2024-02-02T13:00Z`; a save's `w:date` and `w:fullDate` as 13 says, `normalise=False` the source's) before making them.

1. **docx4j's D14 is ported.** `PRINTDATE` is registered as `cp:lastPrinted` through `_core_date("LastPrinted")`, the
   path `CREATEDATE` and `SAVEDATE` take: to the minute, in the package's zone, Word's default format where there is
   no `\@`, and the cached result where the document was never printed (no value). The text model keeps `PRINTDATE`
   cached as before (Q6), so no extraction changes; `evaluate` gives Word's D03 and D06 (`2025-01-02 14:04:00`,
   `1/2/2025 2:04:00 PM`). On the way: `cp:lastPrinted` is a bare `XmlDateTime`, not a `W3CDTF`, and the resolver's
   `_first` found no text in it, so `DOCPROPERTY LastPrinted` had never resolved (a `FieldValueError`, the cached result).
   `_first` now gives its lexical form, as docx4j's resolver does (`XMLGregorianCalendar.toXMLFormat`); the text model's
   own metadata reader, which reads the part's text, is unchanged.
2. **The python-docx facade's core dates are aware in UTC**, as python-docx 1.2.0's `_parse_W3CDTF_to_datetime` gives
   every form (read from its source): `docx/coreprops.py`'s `_datetime_of` makes a naive value UTC and an aware one
   `astimezone(UTC)`, and reads a form the model kept as text with `parse_w3cdtf`. One departure kept: a fraction of a
   second survives here, where python-docx's 19-character slice drops it.
3. **A date alone with a `Z` or an offset** in a `w:date` or `w:fullDate` stays as read by the save.
4. **No Word check** of the probe's save: its forms are Word's own save's.

**Tests:** `test_date_forms_probe.py`'s clock test becomes `test_printdate_is_last_printed` (D03, D06 and `DOCPROPERTY
LastPrinted`) and `test_printdate_where_the_document_was_never_printed_keeps_its_cached_result`;
`test_field_evaluators.py`'s clock test reads `DATE` where it read `PRINTDATE`, and `PRINTDATE` with no `cp:lastPrinted`
is None; its `INFO` test asks `INFO Date` for the clock, `INFO PrintDate` restricted being None either way;
`test_docx_facade.py` gains `test_the_core_dates_are_aware_in_utc_in_every_form_as_python_docx_reads_them`.

## 14. Phase C, the records (2026-10-28)

At the user's "do CR-025 Phase C". Records only: no code, test or generated file changed, so no gate moved from 13.11's
(the fast suite 7,000 passed, 8 skipped); `tests/test_readme.py` runs the readmes as before.

- **`CHANGELOG.md`, Unreleased:** under "Added", `pkg.date_time_zone`, `pkg.write_times_on_save` and the fields under
  `fields: update` with `docx4j_py.model.fields`' new names; under "Changed", `CREATEDATE` and `SAVEDATE` as Word shows
  them (they gave the raw W3CDTF text, the change a caller of 0.2.0 sees) with `QUOTE`'s empty result, and the date-only
  and offset-free forms read and saved as Word does; under "Fixed", `DOCPROPERTY LastPrinted` and the facade's core
  dates.
- **`README.md`:** the audit trail's date paragraph names the package's zone and the save's opt-in; the text model's
  section says what `fields="update"` computes. **`docs/pypi-readme.md`:** the same in one clause.
- **`CLAUDE.md`:** CR-025's status bullet; `model/dates.py` and `model/fields/`' new modules in the hand-written list;
  the date rows in the round-trip rules' normal form.
- **`docs/change-requests/README.md`:** CR-025's row brought up to B2, 13.11 and this phase, and its summary corrected
  (the opt-in stamps to the second, Q3, where the draft said the minute).
- **CR-014:** section 9's note folded into its Q10 and its Status line: the nominal default stands, Word's conversion a
  separate per-package opt-in, `pkg.bound_dates_converted_as_word`, for phase 1b to build.
- **The portfolio:** `python/CR-025.C` and `python/CR-025` done, through the docx4j-portfolio session.
