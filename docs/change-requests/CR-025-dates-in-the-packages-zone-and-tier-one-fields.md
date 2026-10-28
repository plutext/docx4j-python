# CR-025: Dates in the package's zone, and docx4j CR-067 phase 1's fields in the text model --- the portfolio's date rule, and the tier-1 fields under `fields: update`

**Status:** Proposed 2026-10-28, on the user's ruling on the cross-product date policy the same day (`../docx4j-portfolio/docs/date_handling_across_ports.md` section 3.6, committed at `4ac305b`: "I confirm option (a) as amended, and 3.5; also the three sub-items") and his instruction as the portfolio session relayed it: "commit, then have the ports do their work". Three phases (section 6): A, the package's time zone and the core properties' opt-in (the rule's points 2, 5 and 6, after docx4j's D12); B, docx4j CR-067 phase 1's fields in the text model, held to Word's probe (point 1 for the zone); C, the records. A and B are built at the coordinator's brief; C is not, beyond this CR's row in the index and the note for CR-014 in section 9. The names follow docx4j-core-ts's, at the coordinator's request the same day (section 4.5). The questions are in section 8, each with a recommendation. Where building needed an answer, the recommendation is built and listed for the user to decide, as CR-020 sections 10.12 and 11.9 did.

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

**D1.** `pkg.date_time_zone`: the client's zone for this package, a `datetime.tzinfo` or an IANA name (`"Asia/Tokyo"`, read as `zoneinfo.ZoneInfo`), None by default. Never saved: a request attribute, as docx4j's is, so a server reusing a loaded package sets it per request. A bad name or another type raises `ContentError` (`date_time_zone.invalid`). `pkg.effective_date_time_zone` is the zone used: the package's, else the process's, as an object (`docx4j_py.model.dates.PROCESS_ZONE`, a `tzinfo` that reads the process's zone at each instant exactly as `astimezone()` does). docx4j's `getDateTimeZone()` and `effectiveDateTimeZone()`. A dry run's trial reads its package's.

**D2.** The fallback is the process's zone, and nothing between. docx4j's chain has a JVM-wide property between the package and the JVM default (`BindingHandler.dateTimeZone`), because a JVM's default zone is awkward to set per deployment. A Python process's zone is `TZ`, which a deployment sets as easily as it would set a property, so a second process-wide setting would be a second `TZ`.

**D3.** Every place of section 3.1 uses the effective zone. Unset, the code path is today's (`astimezone()` with no argument), so output is byte for byte what it was. `dateUtc` stays UTC; a non-zero offset is still read as it says; the stamp is still to the minute and `tracked_change_date` still as given. The text model's dates (a revision's or a comment's in `markup` view, the metadata) are given as written, converting nothing (point 1), as today.

### 4.2 The core properties' opt-in (D4; Phase A)

**D4.** `pkg.write_times_on_save`, a package attribute, False by default: docx4j-core-ts's `writeTimesOnSave` in Python's spelling (section 4.5). When on, a save sets `dcterms:modified` to now, and `dcterms:created` too where the core part has none and was not read from the source (made by `create_package` or by the caller): docx4j's `DocPropsCorePart.writeTimes()` (17.3.3, its CR-036). In UTC with a `Z`, **to the minute**: the rule's point 5 is "as Word writes it", and Word writes these to the minute (section 2), where docx4j writes the second. A package with no core part gets none. `cp:lastModifiedBy` is not touched: docx4j writes it under its own property, `docx4j.dc.write`, which the rule does not name. `last_save.times` is the stamp written, or None. docx4j-core-ts writes the second too, and makes a core part for a new package that has none (its `createPackage` writes an empty one); here `create_package` writes one with both dates, so the case does not arise, and a package with none is left without, as docx4j leaves it.

### 4.3 The evaluator (D5 to D7; Phase B)

**D5.** docx4j's evaluator, ported where the text model needs it: `docx4j_py.model.fields.FieldInstruction` (the one parser), `FieldContext` (the package, the zone override, a clock, the user, the section, and the pass's sequences, bookmarks and note numbers), `FieldUser`, and a registry by name (`docx4j_py.model.fields.evaluators`: `register`, `get`, `has`, `names`, `evaluate`), each entry a callable returning the result or None to keep the cached one (docx4j's D5), with a flag for the date fields that format themselves. The text model's `FieldValues` becomes a thin consumer over it with docx4j's `COMPUTED` set at `c3e9920e5`: `DOCPROPERTY`, `DOCVARIABLE`, `SEQ`, `REF`, `QUOTE`, the nine document-information fields, and phase 1's `INFO`, `SECTION`, `NOTEREF`, `USERNAME`, `USERINITIALS` and `USERADDRESS`.

**D6.** The parser: `FieldInstruction` is ported, with docx4j's accessors (`name`, `parameter_text`, `parameters`, `argument`, `has_switch`, `switch_value`, `switch_values`, `index_of_switch`, `value_at`, `unquote`), and `FldSimpleModel` becomes a subclass that keeps its one extra name, `parameter_string`. Its tokeniser is already the one docx4j moved into `FieldInstruction`, so nothing a caller reads changes; `FieldInstructionTest`'s 31 shapes hold both to it. `FldSimpleModel` is not deprecated with a warning: the suite turns a `DeprecationWarning` into an error, and the class is public in 0.2.0.

**D7.** No language reaches the text model's evaluator, as none reaches docx4j's inspector (`FieldContext.of(pkg)`): month and day names are English, as `model/java_text.py` writes them.

### 4.4 The fields (D8 to D14; Phase B)

**D8. `CREATEDATE` and `SAVEDATE`.** The core property's W3CDTF with an offset or a `Z` is an instant, truncated to the minute on the instant (docx4j's `floorMod` of the epoch), shown in the package's zone (the context's override, else `pkg.date_time_zone`, else the process's). With `\@`, the switch formats it; with `\@` and no value, the short date (`M/d/yy`), as Java's default short date instance gives it in docx4j. With no `\@`, Word's default: the short date and the long time, `M/d/yyyy h:mm:ss a` in US English (`2/3/2024 3:05:00 PM`), docx4j's choice for no language. docx4j gives another language the locale's short date and medium time, unmeasured; no language reaches this evaluator (D7). A value that is no date keeps the cached result, as docx4j's does (today the raw text was shown). An argument is ignored, as docx4j's `CoreDate` ignores it.

**D9. The forms Word has not been measured on** (the rule's point 4). A date only (`2024-02-03`, also `2024` and `2024-02`) and a time with no offset (`2025-06-07T08:09:10`) are read as today, nominally: the wall clock as written, converted by no zone. The field's own rules, the minute and Word's default format, apply to them as to any date the field shows. With no package zone set, this is docx4j's result too: it reads such a value as local time in the JVM's zone and shows it in the same zone. With one set it differs: docx4j shows the JVM-local reading converted to the package's zone. Docx4j's date-forms probe (its CR-067 section 12.1, staged for Jason) measures what Word does with all three forms; until then each port keeps its reading (Q8).

**D10. The user fields.** `FieldUser(name, initials=None, address=None)`, after docx4j's `FieldContext.user(name, initials, address)`. The text model takes it as `TextOptions(field_user=...)`, an option beyond section 10 like `image_describer`, a host's hook (a `FieldUser`, a name alone, or a mapping with `name`, `initials` and `address`). Unset, the three fields keep their cached result, which is also what Word shows until F9 (cases 20 to 22). Set, `USERNAME` is the name, and `USERINITIALS` and `USERADDRESS` their values or empty where the user has none, as Word shows a blank one. docx4j's extraction has no way to supply a user: this is an extension, and with no user the two agree.

**D11. `QUOTE`** with no argument is empty (case 12). An unknown `INFO` name, a missing bookmark (`REF`, `NOTEREF`) keep the cached result, where Word writes its error text (docx4j's D5).

**D12. `INFO <name>`** evaluates the field it names, with its switches (`INFO Title \* Upper` is `TITLE \* Upper`), **only where that field is one the consumer computes**. docx4j's `INFO` asks the whole registry, so its extraction evaluates `INFO PrintDate` at the clock, though its text model keeps `PRINTDATE` itself cached; here it keeps the cached result. No name, an unknown one, or `INFO INFO` keeps the cached result.

**D13. `SECTION`** is the section of the main story the walk is in, from 1, as docx4j's text model counts it (the sections ended so far, plus one); elsewhere (a note, a header) the cached result. **`NOTEREF`** is the number of the footnote or endnote whose reference the bookmark holds: docx4j's `FieldSources.noteNumbers`, the main part walked in document order, each kind counted apart from 1, a custom mark not counted, separators skipped; the walk goes where docx4j's goes (not into a deletion, a move, a drawing, a picture, an object or `mc:AlternateContent`).

**D14. `DATE`, `TIME` and `PRINTDATE`** are registered, at the context's clock in its zone, for a caller of `evaluate`; the text model keeps their cached result, as docx4j's does (its `FieldValues` leaves them out of `COMPUTED`). Extraction stays deterministic (Q6).

**What extraction does not show** (pinned by the tests as write-back only): a `REF`'s result in the bookmark's formatting (case 13's bold), `NOTEREF \f`'s `FootnoteReference` style (case 18), and `\h`'s hyperlink. The text model gives text; a write-back is docx4j's `FieldUpdater`, which this repository has no port of.

### 4.5 The names, aligned across the ports (the coordinator, 2026-10-28)

The coordinator asked, the same day and before anything was built, that the three ports name these alike, taking docx4j-core-ts's names (its CR-018, uncommitted, which mirrors docx4j `c3e9920e5`) in snake_case. So:

| What | docx4j | docx4j-core-ts | here |
|---|---|---|---|
| the package's zone | `OpcPackage.setDateTimeZone(ZoneId)`, `getDateTimeZone()` | `pkg.dateTimeZone` (an IANA id; a `RangeError` for an unknown one) | `pkg.date_time_zone` (a `tzinfo` or an IANA name; `ContentError` `date_time_zone.invalid`) |
| the zone in force | `effectiveDateTimeZone()` | `effectiveDateTimeZone()` | `pkg.effective_date_time_zone` |
| a process-wide fallback | `BindingHandler.dateTimeZone` | none | none (D2) |
| the core properties' opt-in | `docx4j.dcterms.write` (process-wide) | `pkg.writeTimesOnSave` | `pkg.write_times_on_save` |
| its stamp | the second | the second | the minute (D4, Q3) |
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
- **The opt-in**: off, a save is byte for byte; on, `modified` (and `created` where D4 says) to the minute in UTC, `last_save.times`, nothing else in the part changed, and a package with no core part untouched.
- **The gates**: the suite, `scripts/roundtrip.py` 50 of 50, `scripts/threads.py`, `compare_acceptance.py` 0 differ against `out/acceptance-cr020b/` with the zone unset and the opt-in off; `codegen/generate.sh --check` for the new files.

No Word check is expected: extraction output is not a file, and with the zone unset and the opt-in off a saved file is unchanged. The opt-in, when on, re-marshals the core part, which the corpus round trip already holds canonically identical, and writes the form `create_package` writes, which Word has opened since 0.1.0.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | D1 to D4: `pkg.date_time_zone` and `effective_date_time_zone`, `model/dates.py`, the five sites of 3.1, the dcterms opt-in; the tests of section 5 for them | ½ day |
| B | D5 to D14: `FieldInstruction`, `FieldContext`, `FieldUser`, the registry, the evaluators, `FieldValues` over them, `TextOptions(field_user=)`; the probe, docx4j's output, `FieldEvaluatorsTest`, `FieldInstructionTest` | 1 day |
| C | The records: CHANGELOG (Added: `date_time_zone`, `write_times_on_save`, the fields, `field_user`; Changed: `CREATEDATE`/`SAVEDATE`, `QUOTE`), README and `docs/pypi-readme.md` where they name the readers and the save, CLAUDE.md (the status line, the new files in its list), the index's row brought up to date, CR-014's note (section 9) folded into CR-014; the portfolio's `python/CR-025.*` entries through the docx4j-portfolio session | ¼ day |

## 7. Risks

- **Moved expectations.** A test or golden that read a `CREATEDATE` as text, or a `QUOTE` with no argument as cached, moves. Section 3.2's measurement and a scan of every fixture's instructions say none does: no fixture under `tests/` or `samples/` holds a tier-1 field but `QUOTE` with an argument, and the two `TIME` fields stay cached.
- **The process's zone, read as an object** (D1's `PROCESS_ZONE`), disagreeing with `astimezone()` at a gap or a fold. It is only reached through `effective_date_time_zone`; the sites themselves keep `astimezone()` when nothing is set, and a test holds the object to it across a year of instants in two zones with daylight time.
- **A server that forgets to set the zone** shows its own: the rule's default, and why D1's attribute is on the package.

## 8. Questions for the user

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The zone's name and type: `pkg.date_time_zone`, a `tzinfo` or an IANA name, None for the process's; `pkg.effective_date_time_zone` the zone used (D1)? | Yes: docx4j-core-ts's `dateTimeZone` and `effectiveDateTimeZone()` in snake_case (section 4.5), taking a `tzinfo` too, as Python's own zones are objects. |
| Q2 | A process-wide fallback between the package and the process's zone, as docx4j's JVM-wide property is (D2)? | None. A Python process's zone is `TZ`; a second setting would be a second `TZ`. |
| Q3 | The opt-in's stamp: to the minute, as Word writes it (296 of 296 values in Word's saves here), or to the second, as docx4j and docx4j-core-ts write it (D4)? | The minute. Point 5 says "as Word writes it", and Word writes these as it writes a revision's. Offer the other two ports the same through the portfolio; if the user prefers the three alike at the second, it is one constant here. |
| Q4 | The opt-in's name and shape: `pkg.write_times_on_save`, a package attribute, as docx4j-core-ts's `writeTimesOnSave` (D4)? Or a save option as well, as `save(normalise=)` is? | The attribute alone, named as core-ts names it, so that the ports read alike; a per-save option can follow if a caller needs one. |
| Q5 | The parser: port `FieldInstruction`, `FldSimpleModel` a subclass kept, both held to the 31 shapes; or keep `FldSimpleModel` alone and hold it to them (D6)? | Port it. The evaluators read as docx4j's do, accessor for accessor, and the tokeniser is already the same; nothing a caller reads changes. |
| Q6 | `DATE`, `TIME` and `PRINTDATE` under `fields: update`: keep the cached result, as docx4j's text model does, or evaluate them at a clock (D14)? | Keep the cached result. A "now" makes extraction non-deterministic, and docx4j's extraction does not evaluate them. They are in the registry for a caller who supplies a clock. |
| Q7 | The user for `USERNAME`, `USERINITIALS` and `USERADDRESS`: `TextOptions(field_user=...)`; or `pkg.author`, the identity a revision and a comment carry; or none, as docx4j's extraction has none (D10)? | `field_user`. It is per call, as docx4j's `FieldContext.user` is, and a server's user is a request's. `pkg.author` has a default and is the content API's; using it would change extraction for every document with a user field. |
| Q8 | A date-only or offset-free core property with a package zone set: shown as written, this repository's reading; or read in the process's zone and converted to the package's, docx4j's (D9)? | As written, until docx4j's date-forms probe has run (point 4); then both ports follow Word. |
| Q9 | `INFO <name>`: only the fields the consumer computes; or every field the registry has, as docx4j's, so that `INFO PrintDate` shows the clock in extraction (D12)? | Only the consumer's. docx4j's reach looks unintended (its text model keeps `PRINTDATE` cached); offer docx4j the narrower rule through the portfolio. |
| Q10 | `create_package` writes `created` and `modified` to the second. Leave it, or write them to the minute as Word does? | Leave it. The rule's point 6 is about saves; changing a new package's stamps moves every saved file's core part for no reader's gain. Raise it with the portfolio if a port wants one rule for all core stamps. |

## 9. A note for CR-014 (Phase C folds it in)

CR-014 is proposed, not built. Its Q10 (decided 2026-10-23: docx4j CR-058's nominal default, no `dateTimeZone` option) stands: a bound date is the day its fields say, on every machine. The rule's point 3 adds Word's conversion as a separate per-package opt-in, off by default, after docx4j's `OpcPackage.setBoundDatesConvertedAsWord(true)` (CR-058 D6 as amended 2026-10-28): when on, the stored value's `Z` or offset is read as an instant and converted to `pkg.effective_date_time_zone` (this CR's D1), as Word does on a machine in that zone; setting `pkg.date_time_zone` alone does not turn it on. A Python name in this CR's idiom would be `pkg.bound_dates_converted_as_word`, with no process-wide fallback (this CR's D2). It belongs to CR-014's phase 1b, which formats a bound date.

## 10. Effort

1½ to 2 days: A ½, B 1, C ¼.
