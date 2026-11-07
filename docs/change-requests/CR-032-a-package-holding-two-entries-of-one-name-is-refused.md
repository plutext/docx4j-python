# CR-032: A package holding two entries of one name is refused --- WD REQ-206, with `duplicate_entries: first`

**Status:** Proposed 2026-11-04, at the user's word ("draft CR-032 to CR-034"). A package whose container holds two
entries of one name is refused when loaded, the refusal naming the entry. Today the last copy is read silently, and an
untouched save drops the first. `LoadOptions(duplicate_entries="first")` reads the first copy instead, raises
`package.duplicate_entry` once per name, and a save writes only the copy read. The rule is the user's (4 November 2026,
in the docx4j-core-ts session: "refuse the package, with an optional config option to read the first copy and warn"),
with its five details ruled as recommended the same day (WD 0.77, portfolio `b0d333d`). Two phases (section 6), about ½
day (section 9). **Decided 2026-11-04** (the user: "CR-032 as recommended, please build"): Q1 to Q8 as
recommended, Q3 as revised the same day to match docx4j-core-ts's CR-024 D4, a `DuplicateEntryException` subclass of
`InvalidFormatException` (section 8). Word check 75, run by the user for core-ts the same day, confirms the key
(section 3.1). **Phase A implemented 2026-11-04** (section 10: `beef1cf`): every store keeps the first
entry of a name and counts the rest, the zip store reading and sizing by the first `ZipInfo`; the load refuses a
duplicate before reading anything, `DuplicateEntryException` naming each; `duplicate_entries="first"` reads the first,
`DuplicateEntry` first on `package.warnings`, and saves only the copy read; the text model's `package.duplicate_entry`
with `TextWarning.detail`, and `duplicate_entries` by name in `TextOptions`; 67 tests, full suite 9,900 passed.
**Phase B done 2026-11-04** (section 10: the records). **Complete** (no Word check, Q8). **Amended 2026-11-07** for WD 0.100: more than one `officeDocument` relationship is refused too, the same switch reading the last (section 11).

**Depends on:** nothing in another repository. The statement is the WD's REQ-206 at 0.77
(`../docx4j-portfolio/docs/text_extraction_specification_wd.md`). docx4j and core-ts each build their own, and since
the code and the option are the WD's, no separate agreement is needed. The repair policy
(`../docx4j-portfolio/docs/repair_policy_across_ports.md`, P1 to P6 decided 2026-11-04) carves REQ-206 out of R1: a
duplicate entry is refused in every operation, read-only ones included.

## 1. Why

Measured 2026-11-04 at `94c52db` (scratchpad `dup/probe.py`): a package built here, then re-zipped with a second
`word/document.xml` after the first, the first holding `FIRST` and the second `SECOND`.

| case | read | reported | untouched save |
|---|---|---|---|
| the same name twice | `SECOND` | nothing: `skipped` empty, no warning | one `word/document.xml`, holding `SECOND` |
| the second spelled `Word/Document.xml` | `SECOND` | nothing | one `word/document.xml`, holding `SECOND` |

So the package reads a copy no one chose and says nothing, and its untouched save drops the other copy, which breaks
the rule that nothing is dropped silently. Word 365 reads the first entry, but only after its unreadable-content prompt
and a repair (the identifier's `extraction-probe-dupmain`, 81 configurations, 3 November 2026, in REQ-206's NOTE).
Under the repair policy that is tier 2. Apache POI, Tika, the Open XML SDK and LibreOffice refuse; docx4j,
python-docx and all three of our engines read the last copy silently.

## 2. What exists here (measured 2026-11-04 at `94c52db`)

### 2.1 The stores

Every store keeps its names in `_CaseInsensitiveIndex` (`docx4j_py/openpackaging/stores.py:97`), keyed by the name with
backslashes folded to slashes and the leading slash dropped (`_normalise`, line 50), then lower-cased. `add`
(line 108) appends a new key to `_order`, the listing order, and **always overwrites** the stored spelling:

```python
key = stored.lower()
if key not in self._index:
    self._order.append(stored)
self._index[key] = stored
```

- **`ZipPartStore`** (line 152) builds the index from `infolist()`, which is the central directory's order:
  `zipfile` reads only the central directory. `load` reads `self._zip.read(stored)` by name. `zipfile`'s `NameToInfo`
  holds the last entry of each exact name, so an exact duplicate reads the last entry. A case variant resolves to the
  last spelling, and so also reads the last entry. `part_names()` lists the first spelling, while `load` reads the
  last: the two disagree.
- **`DirectoryPartStore`** (line 233): on a case-sensitive filesystem, `word/document.xml` and `word/Document.xml`
  are two files and one key. The same overwrite happens.
- **`MemoryPartStore`** (line 284) and **`FlatOpcStore`** (line 353): `put` replaces. A flat OPC document with two
  `pkg:part` of one `pkg:name` reads the last.

### 2.2 Our own saves write no duplicate

`PartName` equality is case-insensitive, OPC [M1.12] (`part_name.py:122`), so a package cannot hold two parts whose
names differ only in case. `ZipPartSink.put` skips a name it has already written (`stores.py:491`).

### 2.3 Where the store is made

`OpcPackage.load` calls `store_for(source)` (`stores.py:578`), which takes no options, and then
`load_package(store, ..., options)` (`opc_package.py:370`). `load_package` (`load.py:123`) reads
`[Content_Types].xml`, then `_rels/.rels`, then walks the relationships. The text model's embedded packages and
`altChunk`s load through the same `load()` (`model/text/embedded.py:101`), as does `insert_ooxml`'s flat OPC
(`model/content/ooxml.py:155`).

## 3. The statement

REQ-206 at WD 0.77:

- **Refuse by default.** A package whose zip holds two entries of one name MUST be refused, the refusal naming the
  entry.
- **What counts as one name.** Two names are one when they are equivalent as OPC part names, which is
  case-insensitive. The rule covers every entry, `[Content_Types].xml` and the relationships parts included.
- **The option.** An implementation MAY offer `duplicate_entries: first`. Under it, it:
  - reads the first entry of each duplicated name in the central directory (not local-header order) and ignores the
    rest;
  - raises `package.duplicate_entry` once per name, with the selector the entry name as the first entry spells it and
    `detail` the number of entries;
  - on a save, writes only the copy it read.
- **The table.** WD section 10's option table has the row `duplicate_entries | refuse, first | refuse | REQ-206`, so
  REQ-1001 requires the name to be accepted.

### 3.1 Word, measured: check 75 (2026-11-04)

The user ran docx4j-core-ts's check 75 in three rounds (core-ts `test/README.md` item 75, `test/fixtures/check75/`).
Every file reported unreadable content and was repaired.

- **75a, an ASCII case pair** (`word/headerA.xml` then `word/headera.xml`, the relationship naming `headera.xml`):
  - Word takes them as one name and reads the **first** entry, showing headerA's header;
  - its save keeps that copy as `word/header1.xml`.

  So `duplicate_entries: first` by central-directory order is Word's reading, even where the relationship names the
  second spelling exactly. This package's store today reads the last copy, which shows a different part.
- **75b to 75k, non-ASCII part names:** Word 365 reads no part whose name holds a non-ASCII character, raw UTF-8 or
  percent-encoded, even a single entry with no twin. It reports unreadable content, and its repair drops the part.
  - Round 3's controls show that the check's zip writer is not the cause: 75j, a single ASCII header through it,
    opens cleanly, and 75k, the encoded name written by fflate, is dropped too.
  - So Word can't tell `É` from `é`, and the fold of section 4.1 isn't Word's to decide.
  - A percent-encoded name is ASCII, so lower-casing it folds ASCII letters only, which is OPC [M1.12]'s letter.
  - A raw non-ASCII pair, which `str.lower()` also folds, is a package Word repairs anyway.
  - The key stands as Q1 recommends.
- **On check 75's pairs, this package's key** (`_CaseInsensitiveIndex`, measured at `97a04cb`) and core-ts's
  (`PartName.key`) give the same counts:
  - 75a: one key;
  - `header%C3%89.xml` against `header%C3%A9.xml`: two;
  - `header%c3%a9.xml` against `header%C3%A9.xml`: one;
  - a raw `headerÉ.xml` against `headeré.xml`: one.
- **Not here:** this package reads a part whose name holds a non-ASCII character, where Word drops it. That is a
  question of its own for the WD, not REQ-206's.

### 3.2 docx4j-core-ts's CR-024

core-ts's sibling CR is CR-024, decided 2026-11-04: "decisions as recommended, subject to: Word's reading of a
non-ASCII case pair is unmeasured - let's measure it", which check 75 then measured. It matches this CR in where the
check sits, the stores, the option, the warning and its fields, the save, and nested packages. It differs in three
points:
- **Backslashes:** this package's store folds them in every lookup, and core-ts's does not. Each key is its own
  store's lookup rule, so `word\document.xml` beside `word/document.xml` is refused here and not there.
- **Strict loading:** core-ts has no `strict` option, so Q4 is this package's alone.
- **The release:** core-ts ships CR-024 with its 0.7.0.

The refusal's type, a difference in shape only, is aligned by Q3 as decided.

## 4. Design

### 4.1 One name (D1, Q1)

- **The check sits in `_CaseInsensitiveIndex.add`.** A name whose key is already present is a duplicate. The first
  entry keeps its spelling and position, and the index counts the copies per key: `duplicates()` returns
  `{first spelling: count}` in listing order.
- **The key is the store's own:** backslashes folded, the leading slash dropped, `str.lower()`. This is the store's
  lookup and `PartName`'s equality, so any two entries the store would confuse are one name.
- **A named departure from OPC's letter.** OPC [M1.12] compares part names as case-insensitive ASCII strings, but
  `str.lower()` also folds non-ASCII letters (`É` and `é`). The store cannot tell such a pair apart, so refusing it is
  the safe reading. Word's reading of a non-ASCII case pair is unmeasured.
- **Percent-decoding is not in the key.** `resolve` tries the exact name before the decoded one, so two spellings that
  differ only by percent-encoding never confuse the store, and they are not refused.

### 4.2 Every store (D2, Q2)

- **Zip:** the order is the central directory's.
- **Directory:** the order of the store's walk. Only case variants can collide.
- **Memory and flat OPC:** the order the parts are given: a mapping's iteration order, or the flat document's
  `pkg:part` order.
- **A sink keeps its replace semantics.** `MemoryPartStore.put` still replaces, as a sink's write. Duplicates are
  counted only while a store is filled from its source: the constructor's mapping and `FlatOpcStore.parse`.
- **A third-party `PartStore`** without `duplicates()` is not checked, and the docstring says so.

### 4.3 The first copy, read (D3)

`ZipPartStore` keeps the first `ZipInfo` per key and reads by it (`self._zip.read(info)`), never by name. `load`,
`size` and `part_names` then agree. `reopen()` rebuilds the index the same way.

### 4.4 The refusal (D4, Q3)

`load_package` asks the store for its duplicates before it reads `[Content_Types].xml`. Under `refuse`, it raises
`DuplicateEntryException`, a subclass of `InvalidFormatException` (Q3 as decided, core-ts CR-024 D4):

> Duplicate entry word/document.xml in this package: 2 entries (REQ-206). LoadOptions(duplicate_entries="first") reads the first

Every duplicated name is listed, in listing order. The exception carries `code = "package.duplicate_entry"` and
`entries`, the `{name: count}` mapping, so a host can act on it without parsing the message. Nothing has been read at
that point.

### 4.5 The option (D5)

`LoadOptions(duplicate_entries="refuse" | "first")`, default `"refuse"`. `__post_init__` refuses any other value with
`ValueError` naming REQ-206, as `check_mode` does for `mc_load_mode` (`mc_load_mode.py:136`).

### 4.6 The warning (D6, Q4)

- **The record.** Under `"first"`, the package records a `DuplicateEntry` per name: a new frozen dataclass in
  `openpackaging/`, `code = "package.duplicate_entry"`, `entry` (the first copy's spelling), `count`.
- **Where it shows.** `package.warnings` lists these first, in listing order, ahead of the parts' records. It is not
  skipped content.
- **Strict loading.** `LoadOptions(strict=True)` does not refuse a duplicate under `"first"`, since the option is the
  caller's explicit choice.

### 4.7 The save (D7)

Under `"first"`, a save writes only the copy read, with no change to `save.py`:

- an untouched part's bytes come from `store.load`, which is now the first copy;
- `[Content_Types].xml` and the relationships parts are written from what was parsed, which is the first copy;
- an entry no relationship names is not carried, as today.

The tests hold all three.

### 4.8 The text model (D8, Q5, Q7)

- **The warning in the result.** `extract(pkg)` carries the package's `DuplicateEntry` records into
  `TextResult.warnings` as `TextWarning(kind="package.duplicate_entry")`, the selector the entry name.
- **The count.** `TextWarning` gains `detail: int | None`, the count, which `to_dict` includes only when set. This is
  the field REQ-801's later warnings need too: `part.not_read`'s count kept, WD 0.76 (3).
- **The option name.** `TextOptions` accepts `duplicate_entries` (`refuse`, `first`), as REQ-1001 requires, and
  documents it as the load's option. The two adapters (`tests/text/adapter.py`, `platform-tests-adapter/`) pass it to
  `LoadOptions`. `extract` over a package already loaded doesn't decide again: the load's setting holds, and its
  warnings say what was read.

### 4.9 Packages inside packages (D9, Q6)

An embedded document or a flat OPC chunk holding a duplicate is refused by its own load, under the default. The text
model already reads a chunk or an embedded package that will not load as not expanded or `content.not_read`
(`embedded.py:117`). `insert_ooxml` refuses a flat OPC fragment with two `pkg:part` of one name.

### 4.10 Not here

- **Several main parts.** `package.multiple_main_parts`, a package whose package relationships hold more than one
  `officeDocument` relationship, was a recommended WD statement still open. The user ruled it on 7 November, and section
  11 builds it on this CR's option.
- **R1's save gate and the `repair` switch** are CR-033's. If the WD's `repair: word` later absorbs
  `duplicate_entries: first` (repair policy P4), this option stays as its alias.

### 4.11 The CHANGELOG lines (Phase B)

- **Changed:** a package holding two entries of one name, OPC part names compared without case, is refused,
  `DuplicateEntryException` (an `InvalidFormatException`) naming each entry. It was the last copy read silently, and an untouched save dropped the
  first.
- **Added:** `LoadOptions(duplicate_entries="first")`; `DuplicateEntry` (`package.duplicate_entry`) on
  `package.warnings` and in `TextResult.warnings`; `TextWarning.detail`.

## 5. Tests and oracles

`tests/openpackaging/test_duplicate_entries.py` builds each zip in the test, as the probe does. The first copy always
holds `FIRST` and the second `SECOND`.

- **The main part, the same name twice:** refused, the message naming it. Under `"first"`: reads `FIRST`, one warning
  with count 2, and an untouched save holds one entry holding `FIRST`.
- **The case variant `Word/Document.xml`:** the same, with the first's spelling named.
- **Three copies:** count 3.
- **`[Content_Types].xml` twice:** refused. Under `"first"`, the first's types hold; the second copy gives a different
  override to prove it.
- **`_rels/.rels` twice, and `word/_rels/document.xml.rels` twice:** the same, the first's relationships.
- **A media binary twice, and an entry no relationship names twice:** refused, since the rule covers every entry.
  Under `"first"`, the binary's first copy is saved and the orphan is not carried.
- **Spellings:** `word\document.xml` beside `word/document.xml` is one name. A non-ASCII case pair is one name, which
  pins the departure. A percent-encoded and a decoded spelling are two names, and not refused.
- **The other stores:** a `DirectoryPartStore` with case-variant files (skipped where the filesystem folds case); a
  `FlatOpcStore` with two `pkg:part` of one `pkg:name`; `insert_ooxml` of such a fragment.
- **The option:** `LoadOptions(duplicate_entries="last")` raises `ValueError`. `strict=True` with `"first"` loads.
- **The text model:** `extract` under `"first"` carries the warning, with `detail` 2. An embedded document with a
  duplicate reads `content.not_read`.
- **Our own saves:** every sample saved and re-read records no duplicate.

**Oracles and checks:**
- The oracle is REQ-206 and its NOTE. The identifier's probe isn't committed there yet.
- No Word check (Q8): the refusal writes nothing, and the option's save is an ordinary package from the same writer.
- The adapter advertises REQ-206 from Phase A.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | D1 to D9: the index, the stores, the refusal, the option, `DuplicateEntry`, the text model's warning and option, the tests | ⅜ day |
| B | The records: CHANGELOG (4.11), CLAUDE.md ("Nothing may be dropped silently" and the engine paragraph), `tests/README.md`, this CR's implementation notes, the index | ⅛ day |

## 7. Risks

- **A file that loads today is refused.** That is the ruling. The CHANGELOG puts it under Changed, and the message
  names the option.
- **A host that builds its own `PartStore`** gets no check (4.2). This is documented, not enforced.
- **No cost.** The index is built once, as now.

## 8. Questions (decided 2026-11-04)

The user, 2026-11-04: "CR-032 as recommended, please build". Every question is decided as recommended, Q3 as
revised.

| # | Question | Recommendation |
|---|---|---|
| Q1 | The collision key: the store's own (`str.lower()`, backslashes folded, the leading slash dropped), or OPC [M1.12]'s ASCII fold? | **The store's own**, so any two entries the store would confuse are refused; the non-ASCII pair is a named departure, unmeasured in Word. |
| Q2 | Every store, or the zip store alone? | **Every store**: a directory's case variants and a flat OPC's repeated `pkg:name` are the same confusion. |
| Q3 | The refusal: `InvalidFormatException` with `code` and `entries`? | **A subclass of it, `DuplicateEntryException`**, carrying `code` and `entries`. This revises the first recommendation (the class itself) to match core-ts's CR-024 D4, so a caller can catch the case alone and still catches it as a malformed package. |
| Q4 | Does `strict=True` refuse a duplicate under `"first"`? | **No**: the option is the caller's explicit choice, and a duplicate is not skipped content. |
| Q5 | `TextOptions(duplicate_entries=...)`. | **Accepted** for REQ-1001 and documented as the load's option; the adapters pass it to `LoadOptions`; `extract` over a loaded package does not decide again. |
| Q6 | A package inside a package. | **Its own load's default**, refuse, which the text model already reads as `content.not_read`. |
| Q7 | `TextWarning.detail`. | **Add it**, in `to_dict` only when set, for this count and for REQ-801's counts to come. |
| Q8 | A Word check. | **None**: nothing new is written. |

## 9. Effort

About ½ day: Phase A ⅜, Phase B ⅛. No fork change, no Word check.

## 10. Phase A (implementation notes, 2026-11-04)

Phase A as built, with Phase B's records. Every question is as decided (section 8), Q3 as revised. No Word check
(Q8): a package with no duplicate loads and saves as before, the refusal writes nothing, and the option's save is an
ordinary package from the same writer.

**Commits.** `beef1cf` (the code and the tests), and this section's (the records: this section, the status line, the
index row, CHANGELOG, CLAUDE.md, `tests/README.md`).

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the index | `openpackaging/stores.py`, `_CaseInsensitiveIndex` | `add` indexes a container's entry: a key already held keeps its first spelling and position, and the entry is counted and gets None; `put` is a sink's write, a key already held keeping its first spelling, nothing counted; `duplicates()` gives `{first spelling: count}` in listing order. The key is unchanged (`_normalise`, then `str.lower()`), and `remove` forgets a key's count |
| the zip store | `ZipPartStore._index_entries` | the index and each key's first `ZipInfo` from `infolist()`, the central directory's order; `load` reads `self._zip.read(info)` and `size` the info's `file_size`, never by name; `reopen()` rebuilds both the same way |
| the other stores | `DirectoryPartStore`, `MemoryPartStore._fill`, `FlatOpcStore.parse` | the directory's sorted walk counts through the index's constructor; a memory store's constructor mapping and a flat OPC document's `pkg:part`s fill through `_fill`, which keeps the first of a name and its content type, so a second `pkg:part`'s type never becomes an override; `MemoryPartStore.put` keeps its replace semantics. Each store answers `duplicates()`, which is not a member of the `PartStore` protocol (a third-party store stays one, unchecked; its docstring says so) |
| the refusal | `openpackaging/exceptions.py`, `DuplicateEntryException(InvalidFormatException)` | `code = "package.duplicate_entry"`, `entries` (`{name: count}`); the message section 4.4's for one name, and `Duplicate entries in this package: _rels/.rels, 3 entries; word/document.xml, 2 entries (REQ-206). LoadOptions(duplicate_entries="first") reads the first of each` for several. Exported from `docx4j_py.openpackaging` beside `InvalidFormatException` |
| the check | `openpackaging/load.py`, `load_package` | the store's `duplicates()` through `getattr`, before `[Content_Types].xml` is read; any value but `"first"` refuses |
| the option | `LoadOptions.duplicate_entries` | `"refuse"` (the default) or `"first"`, `DUPLICATE_ENTRIES`; `__post_init__` raises `ValueError` naming WD REQ-206 for any other value, as `check_mode` does |
| the record | `load.py`, `DuplicateEntry` | frozen, `code` (`package.duplicate_entry`), `entry`, `count`, `message`, `to_dict()`, as `BrOutsideRun`; one per name under `"first"`, kept in the package's `_duplicate_entries` slot (`opc_package.py`), listed first by `package.warnings`, logged once at `WARNING` on `docx4j_py.openpackaging`. Not skipped content, so `strict=True` loads (Q4) |
| the text model | `model/text/result.py`, `options.py`, `extractor.py` | `TextWarning.detail` (`int \| None`, in `to_dict` only where set, in `__str__` bracketed) and `PACKAGE_DUPLICATE_ENTRY`; `Extractor._duplicate_entries` carries the package's records into the result first on a package's reading (`extract(pkg)`, `extract_resolved`), the selector the entry (`part_name`), `detail` the count; `TextOptions.duplicate_entries` (`refuse`, `first`, REQ-206) in `_CHOICES` and last in `OPTION_NAMES`, after `warnings` as in WD section 10's table, documented as the load's option and acted on by no reading |
| the adapter | `tests/text/adapter.py` | a scenario's `duplicate_entries` passed to `LoadOptions`, the package cached per fixture and value; `BEYOND_DOCX4J` advertises REQ-206 (section 5's "the adapter advertises REQ-206") |
| `load()` | `openpackaging/api.py` | its docstring names the refusal and the option |

### 10.2 Before and after

Measured with a probe in this session's scratchpad (`cr032/probe/before_after.py`), over the tests' zips (section
5's cases, built from a package made here), against `0cb5cde`'s package (archived into the scratchpad) and the change.

| case | at `0cb5cde` | Phase A |
|---|---|---|
| `word/document.xml` twice | the last copy read, nothing said; an untouched save one entry, the last | refused, `DuplicateEntryException` naming it; under `"first"` the first read, `DuplicateEntry(word/document.xml, 2)`, the save one entry, the first's bytes |
| `Word/Document.xml` after it | the last spelling's bytes read through either name; `part_names()` the first spelling | the same refusal, the first spelling named; under `"first"` the first read |
| `Word/Document.xml` first (core-ts's decoy-first) | the second, the relationship's exact spelling | refused; under `"first"` the decoy's bytes, as Word's 75a reads it |
| check 75a, `headerA.xml` then `headera.xml` | `headera.xml`'s header | refused; under `"first"` `headerA.xml`'s, Word's reading, saved as `word/headera.xml` |
| `[Content_Types].xml` twice, the second typing the main part a template | the second's types | refused; under `"first"` the first's, the save writing them |
| `_rels/.rels` twice, the second naming `word/other.xml` | `word/other.xml` the main part, `word/document.xml` dropped by an untouched save | refused; under `"first"` `word/document.xml`, the save without `word/other.xml` |
| `word/_rels/document.xml.rels` twice | the second's relationships | refused; under `"first"` the first's |
| a media binary twice | the last copy, saved | refused; under `"first"` the first, saved |
| an entry no relationship names, twice | not read, not saved | refused (the rule covers every entry); under `"first"` not carried |
| a flat OPC document, two `pkg:part` of one name | the last, its content type the override | refused; under `"first"` the first, with its type |
| `insert_ooxml` of it | the last part's body inserted | refused, `ooxml.no_main_part` (10.3 item 4) |

### 10.3 Departures, with why

1. **`DuplicateEntry` is in `load.py`**, beside `LoadOptions`, rather than a module of its own as `BrOutsideRun` has
   (`breaks.py`, which holds CR-030's pass): the record is the load's, and an existing module was preferred. It is
   exported from `docx4j_py.openpackaging`, as `BrOutsideRun` is.
2. **The index has two entry points**: `add` for a container's entry, which counts, and `put` for a sink's write,
   which replaces. A put of a case variant now keeps the first spelling, where before the index took the new
   spelling and the store kept the old bytes under the old one, unreachable; no caller wrote a case variant through a
   sink (a save writes each part name once, `PartName` folding case).
3. **Several names in one message** (section 4.4 gives one): `Duplicate entries in this package: <name>, <count>
   entries; ... (REQ-206). LoadOptions(duplicate_entries="first") reads the first of each`, in listing order.
4. **`insert_ooxml`'s refusal keeps its one code**, `ooxml.no_main_part`, quoting the refusal in its message with the
   exception as `__cause__` (CR-003 section 12.1 item 11: one code for an incoming package that cannot be read). The
   quoted message names `LoadOptions(duplicate_entries="first")`, which `insert_ooxml` does not take (10.7 item 3).
5. **An embedded document holding a duplicate is `content.not_read` at its object's paragraph** (`body/0`), not at a
   nested root, since its load is refused before a nested reading starts: the address of any embedded package that
   will not load (`embedded.py`'s `package_blocks`). A chunk is `altchunk.not_expanded`, as for any chunk that will
   not load (Q6).
6. **The warning is carried on a package's reading only**, `extract(pkg)` and `extract_resolved`, at the start, so it
   comes first; a part's or content's reading carries none. Its selector is `TextWarning.part_name`, holding the
   entry's name as the store spells it, with no leading slash (`word/document.xml`, `[Content_Types].xml`), where the
   other kinds' part names have one: the WD's "the entry name as the first entry spells it".
7. **The load logs each record** at `WARNING`, as CR-030's parse logs `BrOutsideRun`; section 4.6 says nothing of
   a log.
8. **Any value but `"first"` refuses at the load**, so a `LoadOptions` whose field is set after construction, past
   `__post_init__`, can never read a duplicate silently.
9. **The adapter advertises REQ-206 beyond docx4j's statement set** (`BEYOND_DOCX4J`), where it advertised docx4j's
   at `docx4j-17.3.5` less what this package lacks: no check of the copied suites names REQ-206, so neither committed
   report moves. `platform-tests-adapter/run_adapter.py` loads with no options, so it is unchanged.

### 10.4 Tests and numbers

`tests/openpackaging/test_duplicate_entries.py`, new, 56 tests (0.5 s), each zip built in the test from a package
made here (a paragraph, a header and an image), the first copy holding `FIRST`:

- **the main part** (6): the same name twice, the case variant and three copies, each refused with the message,
  `code` and `entries`, and under `"first"` read, warned with its count, saved as one entry of the first's bytes;
- **the decoy first** (1), **check 75a** (2), **an untouched save** (1) writing the first copy's bytes unparsed;
- **the content types, the package's and a part's relationships twice** (3), the first's holding and saved;
- **a media binary, an orphan, several names** (3): the first copy saved, the orphan not carried, one message listing
  every name in listing order;
- **the refusal reads nothing** (1, a counting store); **the warning** (1) first on `package.warnings` ahead of the
  first copy's `BrOutsideRun`, logged, `to_dict()`; **strict** (2); **the option** (4 values refused, the default);
  **a clean package** (1);
- **the key** (9): over the zip store, check 75's pairs (75a one name, `%C3%89` against `%C3%A9` two, `%c3%a9`
  against `%C3%A9` one, a raw `É` against `é` one), an encoded against a decoded spelling two, a backslash and a
  leading slash folded (7); a backslash twin of the main part refused, and an encoded and a decoded name loading (2);
- **the zip store** (2): `part_names`, `load` and `size` agreeing on the first `ZipInfo` where `zipfile` reads the
  last; `reopen()`;
- **the other stores** (6): a directory's case variants (skipped where the file system folds case; it does not
  here), a memory store's mapping counted and `put` replacing uncounted, a memory store refused, a flat OPC document
  (the first's content type kept), `insert_ooxml` refusing, and a third-party store without `duplicates()` loading
  unchecked;
- **our own saves** (13): every `.docx` in `samples/`, loaded, saved and re-read, holding no duplicate.

`tests/text/test_duplicate_entry_warnings.py`, new, 11 tests (0.5 s): `extract` under `"first"` carrying the warning
(`detail` 2, `to_dict`, `__str__`), first and once in every view; a warning without a count writing no `detail`; a
reading not deciding again; the option by name (its values, case, `to_dict`, `repr`, `copy`) and refused (3 values);
the adapter loading under the scenario's value, advertising REQ-206, refusing by default; an embedded document
`content.not_read` and a chunk `altchunk.not_expanded`.

67 new tests in all. Mutations show they bite (runtime patches, `cr032/plugin/mut_plugin.py`): the zip store reading
by name fails 12; `duplicates()` returning nothing fails 38; a duplicate taking the index's spelling fails 3; `put`
counting fails 1; the text model's carry off fails 4; `package.warnings` listing the records last fails 1. At
`0cb5cde` both files fail at collection (none of the names exist).

**The cost**, best of thirty, CPython 3.14.6, interleaved with `0cb5cde`'s package twice (`cr032/probe/cost.py`): a `ZipPartStore` over `samples/Symbols.docx` 0.049 and 0.047 ms at `0cb5cde`, 0.047 and 0.049 ms with the change; `load()` 0.511 and 0.511 ms against 0.518 and 0.525 ms. `Headers.docx` 0.709 and 0.699 ms against 0.714 and 0.706 ms, `2010-sample1.docx` 0.485 and 0.486 ms against 0.499 and 0.505 ms: the index is built once, as before, with a second dictionary of `ZipInfo`s; within noise, at most 3%.

ruff (`check`): the two new test files clean; the changed files find what they found before (4 in `stores.py`, 1 in
`load.py`, 2 in `opc_package.py`, 2 in `extractor.py`, 1 in `test_conformance.py`).

### 10.5 Pre-existing tests whose expectation changed, and why

1. `tests/text/test_texts.py::test_the_options_are_section_10s_names`: 27 option names, not 26, the last two
   `warnings` and `duplicate_entries`, section 10's order (Q5).
2. `tests/text/test_conformance.py::test_the_adapter_advertises_the_target_less_the_inspector`: the adapter's
   statements gain `BEYOND_DOCX4J`'s REQ-206 (10.3 item 9), which the test now holds is in no docx4j set.

No existing test builds a zip holding a duplicate: the fast suite run with `zipfile`'s `Duplicate name` warning made
an error (`-W "error:Duplicate name:UserWarning"`) passed whole, and so did the full suite with the refusal in. No
test was skipped or loosened.

### 10.6 The gates

| gate | result |
|---|---|
| the full suite, slow included, the save check on (`tests/conftest.py`) | 9,900 passed, 8 skipped, 0 failed (6 min 43 s) at `beef1cf`, and the same again (6 min 46 s) before this section's commit; the 13 warnings the fork's `ConverterWarning`s of other tests, as before; `test_performance.py` passed in both. `0cb5cde` had 9,833 passed and 8 skipped (7 min 38 s), run this session first: the 67 are this phase's. The fast suite with `-W "error:Duplicate name:UserWarning"` passed whole (10.5) |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50 parts canonically identical, 0 differences, 0 skipped, 34 benign respellings (`boolean-spelling`), 0 parts repaired |
| `codegen/generate.sh --check`, `scripts/compare_acceptance.py`, `scripts/release_check.sh` | not run: no generated file, no marshalling, no packaging and no save of a package without a duplicate changed |

### 10.7 Found, and left open

1. **The conformance harness does not know the option.** `tests/text/conformance.py` is the port of docx4j's harness
   at its tag, and its `_OPTION_STATEMENTS` has no `duplicate_entries`, so a scenario naming it would be refused there
   as unknown. No row of the copied suites names it; when a suite carries REQ-206's rows, the harness's table needs
   the row `duplicate_entries: REQ-206`, as docx4j's will.
2. **Word drops a part whose name holds a non-ASCII character** (section 3.1), where this package reads it. Not
   REQ-206's, and recorded for the WD; nothing here changes.
3. **`insert_ooxml`'s message names an option its caller cannot pass** (10.3 item 4). A host pasting a flat OPC
   fragment with a duplicate has no `"first"`; whether it should is the WD's question, not this CR's.
4. **`package.multiple_main_parts`** (section 4.10): ruled and built, section 11.
5. **The registry**, through the docx4j-portfolio session: `python/CR-032` complete.

## 11. Two main parts: WD 0.100's amendment (implementation notes, 2026-11-07)

The user: "fix the 2 gaps identified in connection with WD 0.100". REQ-206's statement, ruled by the user on 7 November in the docx4j session and written into WD 0.100 (portfolio `9034ad0`): a package whose package relationships hold more than one `officeDocument` relationship is refused, naming the targets; under the same switch an implementation reads the last relationship's target, as a fresh Word does, and raises `package.multiple_main_parts` once, the selector that part's name and `detail` the number of relationships. An unreferenced part is no case: an orphan, never read, dropped on save.

Before this, `load_package` took the first relationship and said nothing, so it read a different part from a fresh Word's.

### 11.1 What landed

- **`MultipleMainPartsException`**, an `InvalidFormatException`, with `code` and `targets`, each target as a part name in the relationships' order. Raised by `load_package` under `duplicate_entries="refuse"`, the default.
- **`MultipleMainParts`** in `load.py`, `part_name` and `count`, the record under `duplicate_entries="first"`, first on `package.warnings` after any `DuplicateEntry`, and logged.
- **`TextWarning.PACKAGE_MULTIPLE_MAIN_PARTS`**, carried by `Extractor._load_warnings`, renamed from `_duplicate_entries`, as `package._duplicate_entries` is now `package._load_warnings`: each holds both records.

### 11.2 What the build settled beyond the statement

- **The other relationships go at load**, with the Override of each part they named. `get_main_part`, `main_document_part`, the loader and the save then see one main part, and nothing else needs to know. So a save holds one main part, as Word's repaired save does, and reloads under the default with nothing said.
- **The count is of relationships**, so a transitional and a strict one are two, and two to the same target are two. The second is refused as the statement reads, and the portfolio's `one-target` round (`ot-*`, four files twice) measured it: Word repairs two or three relationships to one part, however spelled, so it isn't an ordinary package. Under the switch its part is read and keeps its content type.
- **The switch's name reads oddly here:** `duplicate_entries="first"` reads the **last** relationship. It's the WD's switch and the WD's value, so the name stays, and the docstring and the refusal's message say which.
- **`strict=True`** refuses nothing more under the switch, as for a duplicated entry (Q4).
- **An orphan's Override stays.** `ContentTypeManager.to_bytes` writes the source's Overrides as parsed, so a save drops an orphan part but keeps its Override, as it always has. That's what a save writes, CR-033's question, and unmeasured in Word.

### 11.3 The oracle

The tests build their zips (`tests/openpackaging/test_multiple_main_parts.py`). Word's saves of the `pw-two-main-rels-*` round aren't an oracle: in those runs Word read the first relationship, and the WD's `open-order` round shows why, its choice alternating with each repaired open in one Word session, the first open of a session reading the last. A refusal writes nothing, and the switch's save is an ordinary package. Word opened that save silently, both runs, and an orphan's save with its Override too (CR-039 section 9.7).
