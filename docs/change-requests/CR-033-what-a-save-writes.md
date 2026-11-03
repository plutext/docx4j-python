# CR-033: What a save writes --- the session's findings warned on every save (repair policy P6), a save that would lose content refused unless `repair: word` (R1), and a relationships part read as far as its error

**Status:** Proposed 2026-11-04, at the user's word ("draft CR-032 to CR-034"). It builds python's half of the repair
policy (`../docx4j-portfolio/docs/repair_policy_across_ports.md`), whose P1 to P6 the user decided as recommended on
2026-11-04 (section 4.1 there, portfolio `6085a74`):
- **P6.** The default save runs the validator's walk over the parts it re-marshals and records what the session
  introduced on `last_save`, where today it records nothing. `save(check=True)` still refuses.
- **P2 as R1, with the gate at `save()`.** A save that would lose content is refused, naming the part, unless the
  switch is on. An untouched save of the source bytes stays allowed.
- **Optionally, WD 0.76 (3).** A relationships part not well formed is read as the relationships parsed before the
  error. The WD allows it and no longer requires it, so CR-031's refusal stays conforming.

Four phases (section 6); C is optional. About 1⅛ days without C, 1⅝ with it (section 9). The questions are in section
8, Q1 being the repair policy's open P7. Nothing built.

**Depends on:**
- **CR-032**, for `TextWarning.detail` (Phase C's count kept) and for the repair policy's carve-out: a duplicate entry is
  refused in every operation, so this gate never sees one.
- **The repair policy's P7** (open, the user's): whether R1's gate covers content a lenient load skipped. Phase B is
  written for "yes", the document's recommendation.
- **P4** (open): the switch's name. `repair: none | word` is the recommended WD option, which becomes the WD's when the
  first engine builds it, as this CR would.

## 1. Why

### 1.1 The default save says nothing of what it writes wrong (P6)

The content API refuses most misplaced content as code adds it (CR-020 Phase A). What it cannot see is left to the
validator:
- a wrong object put in a single-valued field (`p.p_pr = ...`);
- a list assigned whole;
- a table cell left without its closing paragraph;
- an out-of-range measurement.

The validator runs only under `save(check=True)` or `package.misplaced()`. The default save, `check=False`, writes all
of these silently. That is the departure the repair policy's P6 names. The user ruled the remedy, a warning on
`last_save`, together with the cost it has to pay for.

Measured 2026-11-04 at `f23bbbb` (scratchpad `p6/`; warm, medians of 9, `save(check=True)` against `check=False` on
one loaded package with its main part re-marshalled):

| document | document.xml | save | with the walk | added |
|---|---|---|---|---|
| `sample-docx.docx` | 17 KiB | 9.1 ms | 9.7 ms | +0.7 ms (7%) |
| `tables.docx` | 51 KiB | 54.5 ms | 56.7 ms | +2.2 ms (4%) |
| about 50 pages (`sample-docx.docx`'s body ten times) | 148 KiB | 57.7 ms | 64.1 ms | +6.3 ms (11%) |
| `Symbols.docx` | 772 KiB | 254.5 ms | 289.1 ms | +34.6 ms (14%) |

- **Per operation:** across a whole load, edit and save, that is about 6% on 50 pages and 8% on `Symbols.docx`.
- **A document that already holds a finding doubles its save.** One `w:ind w:left="7205001"` in the source took the
  50-page save from 57.9 to 119.3 ms, and `Symbols.docx` from 255.3 to 509.1 ms. To tell what the session introduced
  from what the file held, `check_for_save` (`validator.py:850`) parses the part's source again
  (`XmlPart.source_tree`, `xml_part.py:363`) on every save.
- **None of the 16 samples holds a finding.**

### 1.2 A save that loses content (R1)

Measured at `94c52db` (scratchpad `cr033/skipped.py`): a `w:unknownThing` in a paragraph, loaded leniently.

| | `part.skipped` | the save keeps it | `last_save` says |
|---|---|---|---|
| the part never read | — | yes, byte for byte | nothing, and nothing needed |
| the part read, so re-marshalled | 1 | **no** | nothing |

The load logs the skip at WARNING and keeps it in `part.skipped`. The save drops it and records nothing. Under R1 that
is a lossy write, refused by default. It is the only one python makes today:
- CR-032 refuses duplicate entries in every operation.
- A part not well formed is either written byte for byte or refused when consulted (CR-031).
- CR-030's repaired form is Word's own save of it (tier 1).
- `mc_load_mode="word"` writes Word's save, by the caller's choice at load.

### 1.3 A relationships part cut short (WD 0.76 (3), optional)

CR-031 refuses the load of a package whose relationships part is not well formed (D5). Text extraction of such a file
therefore returns nothing, where Word, after its repair prompt, keeps every relationship before the break. Ruled on
2026-11-04 at WD 0.76: an implementation MAY refuse; one that doesn't reads the relationships parsed before the error,
raises `part.not_read` naming the relationships part with the count kept, and reads as absent the parts that only a lost
relationship named.

lxml gives exactly that prefix (scratchpad `wd073/prefix.py`). On docx4j's hand-off file, a `document.xml.rels` cut
after its third relationship, an `iterparse` keeps `rId1` to `rId3` (styles, numbering, settings) and stops at
"Premature end of data in tag Relationships", as Word does.

Reading the prefix is safe only with R1's gate in place: a save that rewrites the relationships from the prefix drops
the parts the lost ones reached.

## 2. What exists here (measured 2026-11-04 at `94c52db`)

- **`save_package`** (`openpackaging/save.py:100`):
  - runs `check_before_save` first;
  - then writes the core times and runs the save hooks;
  - then writes inside `normalised_package`.
  - It records the check's carried findings in `last_save.findings`, and `checked` whether it ran.
  - `OpcPackage.save` runs the check before it opens the sink (`opc_package.py:505`).
- **`check_for_save`** (`validator.py:850`) walks `_xml_parts(package, unmarshalled_only=True)`:
  - A misplaced finding is always the session's: a parse never produces one.
  - Any other finding is matched against the source's, by key, from `part_findings(part, part.source_tree())`.
  - It raises `SaveCheckError` for what the session introduced, and returns what the source held.
- **`checking()`** (`save.py:69`) is the context variable that sets the default for `check=None`, off outside the test
  suite.
- **`SaveReport`** (`normalise.py:237`) has `normalise`, `normalised`, `added`, `checked`, `findings`, `times`, and
  `out_of_range` derived.
- **`part.skipped`** is per part, and `package.skipped` gathers it (`opc_package.py:397`). Nothing in `save.py` reads it.
- **A relationships part not well formed refuses the load.** The load reads each relationships part's contents
  (`RelationshipsPart.relationships`, `relationships_part.py:151`). For a part not well formed, `XmlPart`'s refusal
  (`xml_part.py:358`, "Problem with part ...: <lxml's message>") ends the load (CR-031 D5; `load.py:145` documents it).

## 3. The policy, as it reaches python

Section 3 of the repair policy records python's position (`21674d7`, `4c4a1b7`), and R1 is restated around it:
- **A read-only operation** reads what Word's repair would give, marked.
- **A write that would lose content** is refused by default, naming the part, unless the switch is on.
- **An untouched save of the source bytes** is lossless and allowed.
- **Where the gate sits** is each engine's. Python has no read-only load, so here it sits at the part's first
  consultation and at `save()`.

Section 5 records what each engine saves. P6's row: "warn on `last_save`, the source's findings cached once per part".

## 4. Design

### 4.1 The session's findings on every save (Phase A; D1, Q2, Q3)

- **The modes.** `check` takes three meanings:
  - `None`, the default: walk and record;
  - `True`: walk and refuse, as now;
  - `False`, given explicitly: no walk, the documented opt-out for a caller that saves in a loop and accepts the risk.

  `checking(True)` keeps making `None` refuse for the suite.
- **The default save.** It runs `check_for_save`'s walk without raising. What the session introduced goes to the new
  `last_save.introduced` (`list[Finding]`, in part and document order), and what the source held to `last_save.findings`,
  as under `check=True`.
- **The log.** A non-empty `introduced` also logs one WARNING line on `docx4j_py.openpackaging`, as the load logs a
  skip: "save wrote 2 findings this session introduced; the first: target.misplaced at /word/document.xml: ...;
  last_save.introduced holds them".
- **`checked`.** `last_save.checked` is True whenever the walk ran.
- **The refusal.** `SaveCheckError` still comes only from `check=True`.

### 4.2 The source's findings, once per part (Phase A; D2)

`check_for_save` asks the part for `source_findings()`: the `Counter` of finding keys in its source, computed once from
`source_tree()` and kept on the part (a new slot). The source bytes never change while the package holds the store.
`ZipPartStore.reopen()`, after a save over the same path, makes the package's parts read from the new file; the cache
is dropped then, keyed by the store's identity and a reopen count. A document carrying a finding then pays one parse
of that part's source per package, not one per save.

### 4.3 The gate (Phase B; D3, Q1, Q4)

- **When it runs.** Before anything is put in the sink, beside the check, `save_package` asks each part it will
  re-marshal what writing it would lose. Today that means `part.skipped`, when not empty. Phase C adds a relationships
  part read as a prefix.
- **The refusal.** Under `repair="none"` (the default), a loss raises `SaveLossError`, a `Docx4JException` with
  `code = "save.content_lost"` and `losses`, one `ContentLost` per part: its name, the kind (`skipped`, or
  `relationships_cut`) and the count:

  > Saving would lose content: /word/document.xml holds 1 item the load skipped (part.skipped). LoadOptions(repair="word") writes the part without it

- **Under `repair="word"`.** The save writes, and records each loss in `last_save.repaired`.
- **What the gate never refuses:**
  - a part never re-marshalled, which is written byte for byte;
  - an unreadable part, which is written byte for byte (CR-031).
- **What the gate never sees:** what the caller removed through the API, which is the caller's edit, not a loss.
- **The code** is a proposal for the WD's vocabulary (R2).

### 4.4 The switch (Phase B; D4, Q5)

- **The option.** `LoadOptions(repair="none" | "word")`, default `"none"`, checked in `__post_init__` as
  `mc_load_mode` is.
- **Per package.** `pkg.repair` overrides it for one package, as `pkg.mc_load_mode` does. A save takes no `repair`
  argument: the switch is the caller's Yes to Word's prompt, given in advance.
- **The load reads the same either way** (R1: read-only operations get the repaired reading). The switch decides only
  whether a lossy save may be written.

### 4.5 A relationships part read as far as its error (Phase C, optional; D5, Q6)

- **The load.** `RelationshipsPart` overrides its parse: on lxml's `XMLSyntaxError` it reads the complete
  `Relationship` children of the root that precede the error, by `etree.iterparse` (scratchpad `wd073/prefix.py`), and records
  `cut = RelationshipsCut(part_name, kept, reason)`. The load carries on with those relationships.
- **The parts the lost relationships named** are never reached, so they are absent: a missing comments part reads as
  one, the property parts as WD 0.76 (1) has it (CR-034).
- **The text model** raises `part.not_read` naming the relationships part, with `detail` the count kept (CR-032's
  field), once.
- **The save.** A relationships part with a `cut` is a loss to the gate, in every save, the untouched one included.
  The save writes the relationships from the prefix, and the parts the lost ones named are not written. Writing the cut
  bytes back would not be lossless either: the parts after the cut were never reached. So under `repair="none"`, a
  package read with a cut cannot be saved; under `"word"`, it saves as Word's repaired save does.
- **A part inside an embedded package** gets the same reading, by its own load.
- **CR-031's records.** D5 and Q3 are superseded where this lands, `load.py`'s docstring with them; CR-031 gains a
  pointer here.

### 4.6 Not here

- **The one rule for every non-story part** (a part cut late read as its complete children before the error) is an
  open WD item, the user's. It would reach styles, numbering and settings, reverse CR-031 A2, and move the text model's
  `part.not_read` readings.
- **`package.multiple_main_parts`** is CR-032's, if ruled.
- **A Word check** is not needed for A or B: nothing new is written. C's repaired save is ordinary output from the same
  writer. Q7 asks whether to open one anyway.

### 4.7 The CHANGELOG lines (Phase D)

- **Changed:**
  - The default save runs the validator's walk over the parts it re-marshals, and records what the session introduced
    in `last_save.introduced`, with one WARNING log line. `save(check=False)` skips the walk.
  - A save that would drop content the load skipped is refused (`SaveLossError`, `save.content_lost`) unless
    `LoadOptions(repair="word")`.
- **Added:** `LoadOptions(repair=)`, `pkg.repair`, `last_save.introduced`, `last_save.repaired`, `SaveLossError`,
  `ContentLost`.
- **Phase C, if built:**
  - **Changed:** a relationships part not well formed is read as far as its error, where the load was refused.
  - **Added:** `RelationshipsCut`, `part.not_read` with its count.
- **Fixed:** a document carrying a validator finding no longer parses its source again on every checked save.

## 5. Tests and oracles

**Phase A**, `tests/openpackaging/test_save_findings.py`:
- **The routes:** a wrong single-valued field, a list assigned whole, a cell emptied, and an out-of-range value set
  under `unchecked()` each land in `last_save.introduced` on a default save, with the WARNING line, and `check=True`
  still raises.
- **The opt-out:** `check=False` records none and walks nothing; a counter on `part_findings` proves it.
- **Carried findings:** a source finding goes to `last_save.findings`, not `introduced`.
- **The cache:** a second save of a document carrying a finding parses no source (the counter), and after `reopen()` it
  parses again once.
- **The guard:** `test_performance.py` gains a guard on `Symbols.docx`'s default save, at the measured +14% with the
  suite's usual margin.
- **The suite's own saves:** every save passes through the walk, with `checking(True)` as today.

**Phase B**, `test_save_gate.py`:
- the probe of 1.2, untouched: it saves, byte for byte;
- re-marshalled: it is refused, the message naming the part and the option;
- under `repair="word"`, it saves, with `last_save.repaired` naming the loss;
- `pkg.repair` overrides the load's option;
- an unreadable part next to it is still written byte for byte;
- an API removal is not a loss.

**Phase C**, `test_relationships_cut.py`, on docx4j's `broken-part-document-rels.docx` and
`broken-part-package-rels.docx` (e2af12b74, copied with `SHA1SUMS` as CR-031 Phase A copied `part-not-read/`), against
their Word 365 saves:
- the document's text and its comment references, the comments read as absent;
- the property fields read as CR-034 gives them;
- `part.not_read` with its count;
- every save refused under `"none"`; under `"word"`, the saved relationships are the prefix;
- a cut inside a `Relationship` element (no complete child) keeps none.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | D1, D2: the default save's walk and record, `check=False`'s opt-out, the per-part cache, the tests and the guard | ½ day |
| B | D3, D4: the gate over `part.skipped`, `repair`, `SaveLossError`, the tests | ⅜ day |
| C (optional) | D5: the relationships prefix, CR-031 D5 superseded, the fixtures and tests | ½ day |
| D | The records: CHANGELOG, CLAUDE.md (the engine paragraph and "Nothing may be dropped silently"), `tests/README.md`, CR-020 and CR-031 pointers, the index | ¼ day |

## 7. Risks

- **The default save gets slower:** 4 to 14% of the save on the measured documents (1.1). `check=False` is the
  documented way out.
- **A save that worked today raises:** a re-marshalled part with skipped content, under Phase B. That is rare (the
  samples' report is empty) and is the point of R1. The message names the option.
- **Phase C reverses a decided CR-031 point.** The WD now allows either reading, so it's the user's choice (Q6).

## 8. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | The repair policy's P7: does the gate cover content a lenient load skipped? | **Yes, from the start**: it's python's one lossy write today, and dropping an element Word does not understand is the engine's own repair (R4). |
| Q2 | `check=False` given explicitly: skip the walk, or still record? | **Skip it**, as a named, documented opt-out; the default `None` records. |
| Q3 | The log line as well as `last_save.introduced`? | **Yes, one line per save**, as the load logs a skip, so a caller who never reads `last_save` still sees it. |
| Q4 | The gate's code, `save.content_lost`, and `SaveLossError`. | **As proposed**, sent to the WD as R2's vocabulary. |
| Q5 | The switch: on the load (`LoadOptions(repair=)`, `pkg.repair`), or also per save? | **On the load and the package only**: it is the Yes to Word's prompt, given in advance. |
| Q6 | Build Phase C, the relationships prefix? | **Yes**: an extraction then reads what Word reads where today it reads nothing, and the gate keeps the save safe. Either reading conforms. |
| Q7 | A Word check of Phase C's repaired save? | **One, optional**: the two files saved under `repair="word"`, opened for a prompt. |

## 9. Effort

About 1⅛ days without Phase C (A ½, B ⅜, D ¼), about 1⅝ with it. No fork change.
