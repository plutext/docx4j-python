# CR-032: A package holding two entries of one name is refused --- WD REQ-206, with `duplicate_entries: first`

**Status:** Proposed 2026-11-04, at the user's word ("draft CR-032 to CR-034"). A package whose container holds two
entries of one name is refused when loaded, the refusal naming the entry. Today the last copy is read silently, and an
untouched save drops the first. `LoadOptions(duplicate_entries="first")` reads the first copy instead, raises
`package.duplicate_entry` once per name, and a save writes only the copy read. The rule is the user's (4 November 2026,
in the docx4j-core-ts session: "refuse the package, with an optional config option to read the first copy and warn"),
with its five details ruled as recommended the same day (WD 0.77, portfolio `b0d333d`). Two phases (section 6), about ½
day (section 9). The questions are in section 8, each with a recommendation. Nothing built.

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
`InvalidFormatException`:

> Duplicate entry word/document.xml in this package: 2 entries (REQ-206). LoadOptions(duplicate_entries="first") reads the first

Every duplicated name is listed, in listing order. The exception gains `code = "package.duplicate_entry"` and
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
  `officeDocument` relationship, is a recommended WD statement still open. If the user rules it, it fits this CR's
  option as a Phase A2.
- **R1's save gate and the `repair` switch** are CR-033's. If the WD's `repair: word` later absorbs
  `duplicate_entries: first` (repair policy P4), this option stays as its alias.

### 4.11 The CHANGELOG lines (Phase B)

- **Changed:** a package holding two entries of one name, OPC part names compared without case, is refused,
  `InvalidFormatException` naming each entry. It was the last copy read silently, and an untouched save dropped the
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

## 8. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | The collision key: the store's own (`str.lower()`, backslashes folded, the leading slash dropped), or OPC [M1.12]'s ASCII fold? | **The store's own**, so any two entries the store would confuse are refused; the non-ASCII pair is a named departure, unmeasured in Word. |
| Q2 | Every store, or the zip store alone? | **Every store**: a directory's case variants and a flat OPC's repeated `pkg:name` are the same confusion. |
| Q3 | The refusal: `InvalidFormatException` with `code` and `entries`? | **Yes**: the class `load` already raises for a malformed package, and the code is the WD's. |
| Q4 | Does `strict=True` refuse a duplicate under `"first"`? | **No**: the option is the caller's explicit choice, and a duplicate is not skipped content. |
| Q5 | `TextOptions(duplicate_entries=...)`. | **Accepted** for REQ-1001 and documented as the load's option; the adapters pass it to `LoadOptions`; `extract` over a loaded package does not decide again. |
| Q6 | A package inside a package. | **Its own load's default**, refuse, which the text model already reads as `content.not_read`. |
| Q7 | `TextWarning.detail`. | **Add it**, in `to_dict` only when set, for this count and for REQ-801's counts to come. |
| Q8 | A Word check. | **None**: nothing new is written. |

## 9. Effort

About ½ day: Phase A ⅜, Phase B ⅛. No fork change, no Word check.
