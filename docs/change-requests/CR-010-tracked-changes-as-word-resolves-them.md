# CR-010: Tracked changes as Word resolves them --- the port of docx4j CR-043

**Status:** Proposed 2026-10-16, on docx4j CR-043's hand-off (CR-003 section 22), and scheduled by
the user the same day ("yes" to the port). The questions of section 9 carry recommendations that
follow docx4j's own decisions on the same questions (its D1, D4, D5 and D7). Phase A proceeds on
them, and each stays open to the user until Phase C lands.
**Depends on:** docx4j CR-043, done 2026-10-16 (`062bd4a44`, `737c32139` and `00614748e` on
`VERSION_17_3_3`, unpushed, shipping in 17.3.3), and this repository's CR-003 Phase F (section 16),
whose listing, `accept()`, `reject()`, `accept_all()` and `reject_all()` this replaces.
**Counterpart:** docx4j's `docs/developer/change-requests/CR-043-tracked-changes-accept-reject-api.md`
(section 3.1 the API, 3.2 the rules with their Word evidence, 3.4 what the port must mind, 11 what
is unmeasured, 14 the agreement run), and its code, `org.docx4j.model.revisions` in docx4j-core.
docx4j-core-ts's `src/model/content/TrackedChange.mts` and `tableRevisions.mts` are the engine docx4j
ported, and the nearer model for Python.

## 1. Why

docx4j's agreement run (CR-003 section 22) put this package's `body.accept_all()` and
`body.reject_all()` against Word 365's own Accept All and Reject All saves of eleven Word-made
documents, through docx4j's structural comparison. They differ on seven. The tracking code is
unchanged since `v0.1.0`, so these are the released package's results. Measured here:

- **The listing reports a document clean when it is not.** After `accept_all()`,
  `get_tracked_changes()` lists 0 changes, while revisions remain that Word still shows:
  - in text boxes, both branches;
  - on moved and formatted paragraph marks;
  - in table, row, cell, grid and section property records;
  - at content-control markers;
  - in math.

  `OutlineStats.tracked_changes` counts the same listing, and CR-006 plans MCP tools over it.
- **Rejecting an inserted field breaks it**: a lone `w:fldChar separate` and its result survive.
- **Equations are invisible**: the equations fixture lists 0 changes.
- **Moves Word wrote with range markers**, list paragraphs on reject, a note whose reference goes,
  bookmarks and emptied comment ranges all resolve differently from Word.

docx4j now has the measured rule set, held to Word's saves in every story by 141 tests, and
docx4j-core-ts agrees with it closely. This CR ports it, so the three engines resolve tracked
changes as one design.

## 2. What exists here (CR-003 Phase F)

| Piece | Where | Stays or goes |
|---|---|---|
| The tracker: the markup every tracked mutation writes | `model/content/tracking.py` (1,149 lines) | **stays**, out of scope (docx4j CR-043 section 6) |
| The view: `TrackedChange`, `tracked_changes_of_paragraph` / `_row` / `_body`, `join_with_next`, the accept and reject helpers | `model/content/tracked_change.py` (850 lines) | the view **stays, over the new engine**; the listing and resolving code **goes** |
| `Body.get_tracked_changes` / `accept_all` / `reject_all`, `Paragraph.` and `Range.get_tracked_changes`, `pkg.get_tracked_changes()` | `model/content/body.py`, `paragraph.py`, `range.py`, `__init__.py` | **stay**, delegating |
| `OutlineStats.tracked_changes`, `to_markdown(view="accepted"\|"markup")` | `model/content/reports.py`, `model/markdown/` | stay; the count follows the listing, and the views must agree with it (Phase C) |
| The tests | `tests/content/test_tracking.py` (65 tests, 1,200 lines), `tests/agent/test_audit_trail.py` | kept; a test that encodes a rule Word contradicts is changed, and each change is recorded |

## 3. Design

### 3.1 Where the engine lives

A new hand-written package, **`docx4j_py/model/revisions/`**, added to `codegen/clean.py`'s
`KEEP`. Its modules mirror docx4j's classes, one module each, so the two read side by side:

| docx4j | Here | What |
|---|---|---|
| `TrackedChanges` | `__init__.py` | `tracked_changes(pkg)`, `accept_all(pkg, filter=None)`, `reject_all(pkg, filter=None)`, `report(pkg)` |
| `Lister` | `lister.py` | the listing: Office JS's grouping, in document order, part by part |
| `TrackedChange` | `change.py` | the engine's change: `type`, `kind`, `author`, `date`, `text`, `part`, `pieces`, `accept()`, `reject()` |
| `Resolver`, `Removals`, `Trees`, `Texts`, `Props` | `resolver.py`, `removals.py`, `trees.py`, `texts.py`, `props.py` | resolving, and its side effects on notes, comments, bookmarks and parts |
| `Tables` | `tables.py` | table, grid, row, cell and section records |
| `Story`, `Generation` | `story.py`, `generation.py` | the story parts and their walk; staleness of views |
| `TrackedChangeFilter` | `filter.py` | authors, types, kinds, dates, parts, a range |
| `TrackedChangesReport` | `report.py` | counts by kind, author and part, with `to_dict()` |

The names are Python's (`snake_case`, module functions where docx4j has static methods); the
vocabulary is docx4j's and Office JS's. The engine imports no model class at module level, as
the engine rule requires (CLAUDE.md), and `docx4j_py.model.revisions` is imported on first use.

### 3.2 One view

The content API's `TrackedChange` (`model/content/tracked_change.py`) becomes the view over the
engine's change, rather than a second class of the same name:

- **Office JS's members keep their shape**: `type`, `author`, `date`, `text`, `get_range()`,
  `accept()`, `reject()`, `to_dict()`.
- **Its extensions** (`id`, `element`, `kind`, `target`, `address`) gain docx4j's `part` and
  `pieces`.
- `Body.get_tracked_changes()`, `Paragraph.get_tracked_changes()` and `Range.get_tracked_changes()`
  filter the engine's listing to their scope.
- `pkg.get_tracked_changes()` lists every story.

A view is valid until the next accept or reject on its package, as docx4j's is: a join removes a
paragraph, and Office JS throws on a stale change. A stale view raises
`TrackedChangeError("tracking.gone")`, which is the code Phase F already has.

### 3.3 Scope of each verb

- **`Body.accept_all()` and `reject_all()`** resolve their own story completely: its paragraphs and
  tables, content controls, text boxes in both branches of `mc:AlternateContent`, and math.
  Word's side effects in other parts follow: a note goes with its reference, a comment with its
  range, and a part with its last reference (D7).
- **`docx4j_py.model.revisions.accept_all(pkg)` and `reject_all(pkg)`** resolve every story, as
  Word's Accept All does: main, headers, footers, footnotes, endnotes, comments.
- **Both return the count** of changes resolved, as docx4j's do.
- **`w:trackRevisions` is never changed** by accepting or rejecting (docx4j section 3.3).

### 3.4 The rules

The rules are **docx4j CR-043 section 3.2, row by row**, each with the Word check that measured
it. This CR does not restate them; Phase A's notes record any row built differently, and why.
Section 11's unmeasured rows (U1 to U20) are followed as docx4j follows them (Q5), so the ports
agree where Word is silent.

The rows that change behaviour here most:

- **grouping** as Office JS groups, measured;
- **resolving order**: last first, list again until nothing is left, then sweep what is never
  listed;
- **moves** paired by `w:name` on their range starts, with the four range markers removed;
- **fields resolved as units**;
- **check 25's list rules** on reject;
- **a paragraph join keeps the next paragraph's `w:pPr` and `w14:paraId`**;
- **the side effects**: notes, comments and their extended parts, bookmarks, drawings and their
  parts;
- **dates** from `w16du:dateUtc` where present;
- **a text box's revisions listed once**, from the Choice branch, with the Fallback's twin resolved
  with it by position.

### 3.5 Text boxes in a lossless tree

docx4j keeps both branches of `mc:AlternateContent` as typed JAXB. This package keeps them as
**wildcard `AnyElement` trees** (CR-002's lossless MCE; checked on the probe: the only
`AlternateContent` holds an `AnyElement` per branch, and `iter_nodes` does not descend into
them). So the text boxes' revisions are out of reach of a typed engine, which is why Phase F's
listing missed them.

Recommended (Q3): **typed islands.**
1. Each `w:txbxContent` found inside a wildcard branch is serialised and parsed into its typed
   class by the runtime parser, as `wml(...)` parses a fragment, with the namespaces in scope
   carried over.
2. It is resolved like any other story.
3. It is written back into the `AnyElement` tree in place, and the rest of the branch is
   untouched.

Nothing here does this today: `LoadOptions(mce_preprocess=True)` chooses a branch in the bytes
before parsing, and `child.py` only walks typed objects that already sit inside a wildcard. So the
conversion both ways is Phase A's to build and test.

The alternative, a second resolver over element names, would duplicate every rule.

### 3.6 Tests, and the oracle

- **Word's saved files are the oracle**, compared structurally rather than byte for byte:
  - each paragraph's text, style, alignment, indent, list level and runs' formatting;
  - each table's rows and properties;
  - comments and notes.

  `tests/revisions/description.py` ports docx4j-core-tests' `DocumentDescription` (541 lines),
  including what it allows for: Word re-splits runs on save, renumbers ids, turns an accepted
  `HYPERLINK` field into a `w:hyperlink` (C16) and writes a row's indent as
  `w:gridBefore`/`w:wBefore` (C23).
- **The fixtures** are docx4j's `docx4j-core-tests/src/test/resources/revisions/` (`cr043/`,
  `core-ts/checkNN/`, `editor/`) and `tracked-changes-equations.docx`. All are Word-made from
  engine-built test files; none is a business document. Recommended (Q4):
  - copy the agreement run's eleven fixtures, and Word's saves of them, into
    `tests/fixtures/revisions/`, with provenance in `tests/README.md`, so CI holds the core;
  - read the rest from the sibling checkout, skipping without it, as CR-009's theme vectors do.
- **docx4j's tests are ported as tests here**: `TrackedChangesWordTest`, `TrackedChangesTest`,
  `TrackedChangesOneChangeTest` and phase 2's table tests. Each fixture is resolved whole, and one
  change at a time.
- **The agreement run is the gate.** docx4j's `revisions/agreement/py_resolve.py` against
  `PortsAgreement`: every line of python's column in CR-043 section 14 goes to zero, or the
  difference is recorded here as a departure. The Java half is docx4j's to run, so Phase C asks
  docx4j's session for the rerun. The same comparison against Word's saves runs here on every
  test run.
- **This package's own markup** must resolve as well. That is what the tracker writes and what
  `moves_package()` builds, so Phase F's tests stay.

### 3.7 Out of scope

As docx4j CR-043 section 6:
- recording edits, which is the tracker's job and unchanged;
- comparing documents;
- legacy markup beyond section 3.2's rows;
- showing changes in PDF or HTML;
- pptx and xlsx;
- Office JS's own quirks (the Review tab is the oracle).

The python-docx facade has no tracked-change members, and gains none.

## 4. Phase A: the engine, the body and every story

docx4j's phase 1. The package of section 3.1:
- the listing, with Office JS's grouping, dates and text;
- `accept()` and `reject()` per change, and `accept_all` / `reject_all` with the order and the
  repeat;
- runs at any holder depth, fields as units, moves, paragraph marks, run, mark and paragraph
  formatting with check 25's list rules, and rows;
- the side effects: comments with their extended parts and replies, notes, bookmarks, and drawings
  with their parts (D7);
- every story part, content controls, text boxes as typed islands (3.5), and math in its three
  forms.

The content API's views are rewired over it: section 3.2, `Body` / `Paragraph` / `Range` /
package.

Tests: `DocumentDescription`, the copied fixtures, and docx4j's phase 1 tests (66 there). The
agreement fixtures' accept and reject compared with Word's saves, as far as phase 1 reaches.

## 5. Phase B: tables, sections, the legacy forms, filters, the report

docx4j's phase 2:
- **the records**: `w:tblPrChange`, `w:tblGridChange`, `w:trPrChange`, `w:tcPrChange` and
  `w:tblPrExChange`, one change per table (or per row); reject restores them wholesale and
  collapses a merged grid (C23);
- **`w:sectPrChange`**, merged on reject;
- **`w:cellIns` / `w:cellDel`** by check 27's rules, and `w:cellMerge`;
- **the legacy forms swept** (D4);
- **`TrackedChangeFilter`**, its conditions applied to single changes (U19);
- **`report(pkg)`**.

Tests: docx4j's phase 2 tests (75 there).

## 6. Phase C: integration

- **The content API's reads agree with the engine.**
  - `OutlineStats.tracked_changes` follows the listing, and `describe()`'s `tracking_on` is
    unchanged.
  - `to_markdown(view="accepted")` reads as `accept_all()` would leave the document, and
    `view="markup"` shows every listed change.
  - A test holds the accepted view's agreement over the fixtures.
- **The README's audit-trail example** and `docs/pypi-readme.md` are updated where the counts
  changed, and executed by `tests/test_readme.py`.
- **`CHANGELOG.md`** gets a "Changed" entry for the grouping, the counts and the stale view, and a
  "Fixed" entry for each finding of section 1.
- **`out/acceptance/`** artefact 8 (tracked changes) is regenerated, with an accepted and a rejected
  copy beside it, and **a Word check is asked for**: each copy must open without a repair prompt
  and show no revisions in Word's Review pane.
- **The hand-off to docx4j's session:** rerun the agreement, and record python's column.
- **CR-003 section 16** and section 22 are pointed here, and CR-006 section 7's tracked-change
  tools are noted as backed by this.

## 7. Phasing and effort

| Phase | Content | Estimate |
|---|---|---|
| A | the engine: every story, runs, fields, moves, marks, formatting, rows, side effects, text boxes, math; the views rewired | 3 days |
| B | table, grid, row, cell and section records; legacy forms; filters; the report | 1.5 days |
| C | the content API's reads, docs, CHANGELOG, acceptance artefact and Word check, the agreement rerun | 1 day |

docx4j's took a week and a half from scratch. This port has docx4j's code, its tests and the
measured rules from the start.

## 8. Risks

1. **Counts change.** Grouping as Office JS groups changes what `get_tracked_changes()` returns
   and what `accept_all()` counts on the same document. This is a behaviour change in `0.x` (Q7),
   recorded in the CHANGELOG. The README's audit-trail example shows counts.
2. **Speed.** Phase F accepts 3,900 changes in 30 ms in one pass. Listing again until nothing is
   left costs more. Measure it in Phase A against the same document, and keep the cost of a
   relist under a few times a single pass.
3. **Parent pointers.** Splicing moves nodes between lists. `ChildList` sets parents on insert,
   but a single-valued field set by assignment does not (CLAUDE.md), so `link_parents` runs over
   what was touched.
4. **Lossless MCE.** A typed island written back must leave the rest of its branch byte for byte
   (3.5); a test holds it.
5. **docx4j moves.** CR-043 is unpushed, and a fix there after this port is a hand-off like any
   other. Port from `00614748e`, and record the commit.

## 9. Questions

Each recommendation follows docx4j's decision on the same question where it made one.

1. **Q1: the package and names**: `docx4j_py.model.revisions`, module functions
   `tracked_changes`, `accept_all`, `reject_all` and `report`, and docx4j's class names as
   modules. *Recommended: as proposed (docx4j D1).*
2. **Q2: one view**: the content API's `TrackedChange` becomes the view over the engine's change,
   rather than a second public class. *Recommended: yes.*
3. **Q3: text boxes in the lossless tree**: typed islands (3.5), rather than a second resolver
   over element names, or giving up lossless MCE for the text boxes. *Recommended: typed
   islands.*
4. **Q4: the fixtures**: copy the agreement run's eleven fixtures and Word's saves of them into
   `tests/fixtures/revisions/`, and read the rest from the sibling checkout, skipping without it.
   *Recommended: yes.*
5. **Q5: where Word is unmeasured**: follow docx4j CR-043 section 11's rules, as docx4j follows
   core-ts. *Recommended: yes (docx4j D5).*
6. **Q6: the legacy forms and orphaned parts**: resolve the forms Word no longer writes by check
   27's rules, and remove a part when its last reference goes. *Recommended: yes, both (docx4j D4
   and D7).*
7. **Q7: the counts change in `0.x`**: grouping and counting as Office JS does, recorded as a
   "Changed" entry. *Recommended: yes; the 0.x promise (CR-007) allows it, and the old counts
   were not Word's.*
