# CR-010: Tracked changes as Word resolves them --- the port of docx4j CR-043

**Status:** Proposed 2026-10-16, on docx4j CR-043's hand-off (CR-003 section 22), and scheduled by
the user the same day ("yes" to the port). The questions of section 9 carry recommendations that
follow docx4j's own decisions on the same questions (its D1, D4, D5 and D7). Phase A proceeds on
them, and each stays open to the user until Phase C lands. **Phase A implemented 2026-10-16**
(section 10): `docx4j_py.model.revisions`, every story, the content API rewired over it, and no
line differing from Word's saves of the eleven agreement fixtures but the table and section
records, which are Phase B's. **Phase B implemented 2026-10-16** (section 11): the table, grid,
row, cell and section records resolved, the legacy forms swept, `TrackedChangeFilter` and
`report(pkg)`; against Word's saves of the eleven fixtures (fifteen comparisons) **no line
differs with the records shown**, by this package's `DocumentDescription` port and by docx4j's
own. **Phase C implemented 2026-10-18** (section 12): the content API's counts held to the
engine's listing, the text and Markdown views' agreement moved to CR-011 by the user's decision
(their disagreement today pinned), the tracker's appended paragraph in Word's check 29 form,
artefact 8 regenerated with an accepted and a rejected copy, and the documentation; the fifteen
comparisons still at zero. **The user's Word check of artefact 8 (2026-10-21, section 12.8)**
found one disagreement: Word's Reject All removes a comment that the engine keeps. Its fix waits
on the user's decision.
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

  *The second and third items moved to CR-011 (section 12.3), by the user's decision of
  2026-10-17 on the portfolio's text-extraction paper (`docs/text_extraction_behaviour_model.md`,
  portfolio `2d8fd36`): its section 3.1 defines the views by Word's resolutions, and its section 10
  gives this package a CR-011 that "folds CR-010 phase C's Markdown agreement", sequenced after
  docx4j CR-046 phase 0's goldens. Phase C pins today's disagreement instead, measured.*
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
*Phase C has landed (2026-10-18): the seven now stand as decided, on these recommendations,
unless the user reopens them.*

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

## 10. Phase A implementation notes (2026-10-16)

Phase A is done, on section 9's recommendations (Q1 to Q7). `docx4j_py.model.revisions` lists,
accepts and rejects tracked changes **as Word resolves them** in every story, ported from docx4j's
`org.docx4j.model.revisions` at `00614748e` (`VERSION_17_3_3`); the content API's views are
rewired over it. Against Word 365's own Accept All and Reject All saves of the eleven agreement
fixtures, compared by the port of `DocumentDescription`, **no line differs** with the table and
section properties left out (Phase A's comparison, as docx4j's phase 1 made it), and what differs
with them shown is the table, grid, row, cell and section records alone, which are Phase B's.

```python
from docx4j_py import load
from docx4j_py.model import revisions

pkg = load("in.docx")
for change in revisions.tracked_changes(pkg):        # every story, Office JS's grouping
    print(change.kind, change.author, repr(change.text), change.part.part_name)
revisions.accept_all(pkg)                            # or reject_all; the count listed
pkg.body.reject_all()                                # one story, with Word's side effects
```

### 10.1 What landed

| piece | where | lines | signature |
|---|---|---:|---|
| the module functions (docx4j `TrackedChanges`) | `model/revisions/__init__.py` | 245 | `tracked_changes(pkg)`, `accept_all(pkg, filter=None)`, `reject_all(pkg, filter=None)` (each returns the count listed), `resolve_all(pkg, accept, filter=None, *, stories=None, key=None)`, `tracked_changes_of(pkg, stories, key=None)`, `body_stories(body)`; `Kind`, `TrackedChange`, `Story`, `is_resolvable` |
| the change (`TrackedChange`) | `change.py` | 564 | `TrackedChange` with `type` (`"Added"` / `"Deleted"` / `"Formatted"`), `kind` (`Kind`, a `StrEnum`: docx4j's 21 kinds in snake case), `author`, `date`, `id`, `text`, `part`, `pieces`, `resolvable`, `twins`, `accept()`, `reject()`; the targets (`RunT`, `MarkT`, `RunPropsT`, `ParaPropsT`, `MarkPropsT`, `RowT`, `GroupT`, `FieldT`, `HolderT`, `MathRunT`, `MathStructT`, `TableT`, `SectionT`, `CellT`); `date_of` |
| the listing (`Lister`) | `lister.py` | 849 | `Lister(pkg, attach_twins, stories=None)` with `list()`, `stories()`, `paragraph_changes(p, container, story)`; `package_stories(pkg)`, `story_parts(pkg)`, `groupable`, `BREAK` |
| the resolution (`Resolver`, its `Moves`) | `resolver.py` | 805 | `resolve(change, accept, removals)`, `resolve_pieces(pieces, accept, removals)`, `restore_text` / `delete_text` |
| the side effects (`Removals`) | `removals.py` | 603 | `Removals(pkg, local=False)` with `discard(story, subtree)`, `remove_comment(id)`, `touched(story)`, `write_back()`, `flush()`, `remove_everywhere(root, match)` |
| the tree (`Trees`) | `trees.py` | 614 | `name`, `attribute`, `marker_id`, `revision_of`, `run_items`, `block_items`, `children`, `list_of`, `is_holder`, `is_math`, `item_text`, `has_text`, `text`, `paragraph_text`, `walk` (typed and wildcard nodes, `r:` ids), `is_empty`, `fit`, `splice`; `Revision`, `View` |
| the text (`Texts`), the records (`Props`) | `texts.py`, `props.py` | 109, 176 | `text_of(change)`, `row_text(tr, view)`; `r_pr_of`, `restore_r_pr`, `restore_p_pr` (with check 25's two list rules), `restore_mark_properties`, `prune_paragraph`, `prune_run` |
| the table records, listed (`Tables`' listing half) | `tables.py` | 102 | `has_records`, `record_of`, `change_text`, `rows_of`, `cells_of` |
| `Story`, `Generation` | `story.py`, `generation.py` | 57, 61 | `Story(package, part, root, root_owner, twin=False, island=None)`; `Generation.of(key)`, `is_current()`, `advance()` |
| the typed islands (no docx4j counterpart) | `islands.py` | 178 | `Island(node, holder)` with `typed`, `dirty`, `write_back()`; `holds_revisions(node)`, `REVISION_LOCALS` |
| the content API's view, rewritten | `model/content/tracked_change.py` | 464 (was 850) | `TrackedChange(change, body)` with Office JS's `type`, `author`, `date`, `text`, `get_range()`, `accept()`, `reject()`, `to_dict()` and the extensions `id`, `element`, `kind`, `target`, `address`, `part`, `pieces`, `paragraph`; `tracked_changes_of_body`, `tracked_changes_of_paragraph`, `tracked_changes_of_package`, `accept_or_reject_body`, `count_tracked_changes` |
| the verbs, delegating | `body.py`, `reports.py`, `content/__init__.py` | +38 | `Body.get_tracked_changes` / `accept_all` / `reject_all` over `body_stories`; `OutlineStats.tracked_changes` counts the same listing; `pkg.get_tracked_changes()` lists every story |

`codegen/clean.py`'s `KEEP` already holds `model` whole, so the package needed no entry; the
comment on it names `model/revisions/` now, and `codegen/generate.sh --check` reproduced the model
byte for byte with the package in place. The old listing and resolving --- Phase F's
`TrackedChangeTarget`, `tracked_changes_of_row`, `folded_in_row`, `changes_in_row`,
`join_with_next` and the accept and reject helpers --- are gone. `tracking.py`, the writer, is
unchanged.

### 10.2 The rules as built

Every row of docx4j CR-043 section 3.2 that phase 1 covers is built as docx4j builds it, and
section 11's unmeasured rows that reach Phase A (U1, U2, U3, U4, U6, U7, U12, U15, U16, U17, U18,
U19) as docx4j follows them. In order of how much they changed what this package did:

1. **Grouping** (section 3.1): touching pieces of one type by one author are one change, a mark a
   piece whose text is `\r`, rows grouping with the rows and inserted text touching them, a row
   its cells a tab apart then `\r\n`; a nested revision listed apart; moves never grouped, each
   half of each moved paragraph a change (U18). `Paragraph.get_tracked_changes()` is Office JS's
   (check 22: the changes starting in the paragraph, cut at its end), which docx4j has no
   counterpart for and core-ts's `trackedChangesOfParagraph` rule supplies.
2. **Order**: last first, pieces of a group runs last first then marks last first; listed again
   until nothing is left (20 rounds at most, two rounds with no progress stop it), then the sweep
   (the stray `w:customXmlIns*` markers: docx4j phase 1's sweep).
3. **Fields as units** (C01, C16): Word's inserted field --- a `w:ins` around begin and
   instruction, plain separator and result, a `w:ins` around end --- is one `field_insertion`,
   rejected whole; a field straddling an edge resolves as runs (U2).
4. **Moves** by `w:name` on the range starts, the four markers removed, moved marks resolved last
   first, an emptied moved paragraph removed; edited and broken moves as check 20 (the destination's
   typed text removed on reject, a broken pair dissolved into plain `w:ins` / `w:del` keeping its
   id). Check 18's move from either half and every check 20 case agree with Word's one-change
   saves (10.4).
5. **Paragraph marks**: a join keeps the **next** paragraph's `w:pPr` and `w14:paraId` in both
   directions (C10, C10b), markers between the two going into it, and never crosses a container
   (U1).
6. **Formatting**: a run's `w:rPrChange` restored wholesale, a mark's with its own revisions kept,
   `w:pPrChange` restored keeping the mark's `w:rPr` and the `w:sectPr`, with check 25's rules (a
   restored `w:ind` equal to the level's dropped, through the numbering emulator's `get_ind`; a
   `w:numPr` with no `w:ilvl` given `w:ilvl` 0).
7. **Side effects**, settled after each resolution against what stands in every story: a note
   with its reference (C12; the review file's footnote); a comment whose range start was inside
   removed content, with its `commentsExtended`, `commentsIds` and `commentsExtensible` entries
   and its replies (U3), one whose markers lie outside losing its emptied range; bookmarks with a
   start or end inside removed content, or left empty around it (C11, C11b, U15); relationships
   nothing references and the parts only they targeted, with their content-type overrides (D7:
   C04's and C05's picture, C06's dropped headers and footers, a hyperlink's relationship in a
   deleted row, U4).
8. **Every story** (main, footnotes, endnotes, comments, headers and footers in section order),
   content controls (Word's `w:customXmlInsRangeStart` form, inserted and rejected whole, C16b),
   **text boxes** listed once from `mc:Choice` with the `mc:Fallback` twin paired by kind and
   position and resolved with it (C22), and **math** in its three forms (inside `m:r`; an
   `m:ctrlPr`'s, the structure going whole, U6; `w:ins` / `w:del` holding math, U7): the
   equations fixture lists the 12 changes Word counts.
9. **Dates**: `w16du:dateUtc` first (10.3 item 1 for `w:date`).

### 10.3 Departures, all deliberate

1. **`w:date` is read as it is written**, a `Z` as UTC, where docx4j reads it as Word's local
   wall-clock time (check 36). `w16du:dateUtc` is preferred, as docx4j does. This package's own
   tracker writes the true UTC in `w:date` and no `w16du:dateUtc` (CR-003 section 16.1), so
   docx4j's reading would put every revision this package makes off by the reader's offset, and
   the tracker is out of this CR's scope. A Word 365 file carries `w16du:dateUtc` and reads the
   same either way; a Word 2010 or 15 file reads its local time as UTC, as Phase F did. Making the
   tracker write Word's convention and `w16du:dateUtc` (core-ts's tracker does) would let the
   reading follow docx4j's; a later CR's.
2. **The table, grid, row, cell and section records are listed by docx4j phase 1's rule** --- a
   table with any record is one `table_properties` change, a `w:sectPrChange` a
   `section_properties` one, a `w:cellIns` / `w:cellDel` a `cell_insertion` / `cell_deletion` ---
   and left for Phase B: `accept()` / `reject()` on one raise `TrackedChangeError`
   (`tracking.unsupported`), `accept_all` / `reject_all` leave them listed and do not count them,
   and `TrackedChange.resolvable` says which is which. Phase 2's refinement (records that change
   nothing not listed, and swept) and the legacy forms' sweep come with Phase B's resolution.
3. **A body that is not a real package's is resolved locally.** A dry run's body (and a detached
   container) resolves its own story, and touches nothing outside its own part --- no note, comment
   entry, relationship or part is removed --- because the other parts are the real document's and a
   trial must leave it byte for byte (CR-003 section 3.4). docx4j has no dry run; its conversion
   mode (`acceptAllForConversion`) is the nearest thing, for the same reason.
4. **The halves of a run this session split are joined again** after a resolution through the
   content API (`Body._tidy_runs`, CR-003 section 16.12), as Phase F did. docx4j has no such split
   memory; Word re-splits runs on save anyway, and the description merges runs, so nothing an
   oracle compares changes.
5. **The vocabulary is Python's**: `kind` is a `Kind`, a `StrEnum` of docx4j's kinds in snake case
   (`"run_insertion"`, `"paragraph_mark_deletion"`, `"move"`, ...), and `type` Office JS's strings,
   as the content API already used them. A filter is any callable over a piece (U19's rule); the
   named `TrackedChangeFilter` is Phase B's.
6. **Model-specific mechanics with no difference in behaviour**: content moved between lists is
   re-typed where the target list keeps that element as another class (`trees.fit`: the generated
   model has one class per element per scope, `MoveFromRangeStart1` between paragraphs and `2`
   inside one); a math `w:ins`'s `m:rPr` and `w:rPr` become the run's where the run has none
   (fields here, members of the content list in JAXB); a story part whose bytes hold no revision
   markup is not unmarshalled for the listing, and one nobody has read is scanned in its bytes for
   the side effects and unmarshalled only when something in it has to go, so that a part a
   resolution never touches is still written back byte for byte.

The CR's own scope rule (section 3.3), not a departure from CR-043's rules: `Body.accept_all()` and
`reject_all()` resolve their own story with Word's side effects in other parts, and
`revisions.accept_all(pkg)` / `reject_all(pkg)` every story. docx4j's `py_resolve.py` calls the
body's, so on the probe and on check 18 its run shows the header, footer and note revisions a
body leaves (10.5).

### 10.4 The typed islands

CR-002 keeps both `mc:AlternateContent` branches, and the shapes inside them, as wildcard
`AnyElement` trees: a text box's `w:txbxContent` sits under `wps:txbx` or `v:textbox`, elements
nothing in this build is bound to, so it and everything in it are wildcards too. `islands.py`
makes one typed for as long as it is resolved:

- **in**: the wildcard subtree is written out as an lxml element (its `qname`, attributes, text
  and tails, under docx4j's prefixes) and parsed as `TxbxContent` by the runtime parser; anything
  skipped raises, so nothing is lost on the way in;
- **resolved** as a story like any other (`Story.island`), its part the part the box is in, so a
  picture's relationship in it is the host part's;
- **out** (`Island.write_back`, before anything is settled and before the next listing reads the
  tree): serialised with `to_xml` and parsed again **inside an element no class is bound to**, so
  the parser keeps `w:txbxContent` a wildcard exactly as it did reading the document; the new node
  takes the old one's place in its parent's `children` and its `tail`. Nothing else in the branch
  is re-parsed or re-serialised.

A text box with no revision markup in it (`holds_revisions`) is never made an island.
`test_a_typed_island_leaves_the_rest_of_its_branch_unchanged` holds risk 4: over the probe's C22,
the `mc:AlternateContent` before and after `accept_all` is identical once each `w:txbxContent` is
cut out (the namespace declarations aside, which follow use: `w16du` goes with the revisions that
carried `w16du:dateUtc`); `test_a_text_box_with_no_revision_is_never_made_an_island` that none is
made over Word's own clean save. The side effects read wildcard trees too (by element name and
`w:id`), so a comment reference or a bookmark inside a text box counts as standing.

### 10.5 The numbers

**The tests.** `tests/revisions/`: 348, of which 283 read docx4j's sibling checkout and skip
without it.

| file | tests | what |
|---|---:|---|
| `description.py` | --- | the port of `DocumentDescription` (541 Java lines), its normalisations included: runs merged, ids, dates and rsids, field codes (an accepted `HYPERLINK` field reads as Word's `w:hyperlink`, C16), inherited header references, the order of comments, list numbers by first appearance, empty parts, and Word's written defaults (`w:tblLook`, `w:tcW` auto 0, `w:cols w:space="720"`, a row's `w:tblPrEx` repeating the table or holding zero margins, C23); `records=False` leaves the table, row, cell and section properties out; `differing` is `PortsAgreement`'s count |
| `test_word_saves.py` | 31 | the port of `TrackedChangesWordTest`'s fixtures: 17 whole resolutions (every story) against Word's saves, with the records shown too where a fixture has none; 13 of them again through `Body.accept_all()` where the body is the document; the probe leaving only the two Phase B records listed |
| `test_one_change.py` | 30 | `TrackedChangesOneChangeTest`: check 22 (4, copied), check 18's move from either half (4) and check 20's edited and broken moves (22), from the sibling |
| `test_tracked_changes.py` | 30 | `TrackedChangesTest` as far as Phase A reaches (grouping, the equations' 12, the text-box twin, dates, a stale view, a piece on its own, a filter on pieces, U1, the join, U2, C01, C11/C11b, U4, U3, math structures; the table, section and cell records listed), and the islands, every story, a note with its reference, a picture's part with its last reference, unread parts left unread |
| `test_every_fixture.py` | 257 | every one of docx4j's 128 revisions fixtures resolved whole both ways and one change at a time, nothing resolvable left, saved and loaded again |

`tests/content/test_tracking.py` is 68 (65 before; 10.6), `tests/agent/test_audit_trail.py` 6.
The fast suite is **2,346 passed, 8 skipped** (1,995 and 8 before), the whole suite **2,386
passed, 8 skipped** (2,035 and 8).

**The agreement run.** docx4j's `revisions/agreement/py_resolve.py` (which calls
`body.accept_all()` / `body.reject_all()`) at HEAD before and after, each output described
against Word's own save by `description.py`; lines differing with the records shown (as
`PortsAgreement` counts them), and in brackets with them left out. "Every story" is
`revisions.accept_all(pkg)` / `reject_all(pkg)` over the same fixture.

| fixture | before (0.1.0) | after, the body's verbs | after, every story |
|---|---|---|---|
| `cr043-tracked` (accept / reject) | 42 (36) / 60 (44) | 18 (12) / 34 (18) | 6 (0) / 16 (0) |
| `tracked-changes-equations` | 10 (10) / 10 (10) | 0 / 0 | 0 / 0 |
| `revisions-word15` (check 18) | 15 (7) / 24 (14) | 10 (2) / 10 (0) | 8 (0) / 10 (0) |
| `review-for-word` (Word 15's saves; Word 2010's alike) | 2 (2) / 4 (4) | 0 / 0 | 0 / 0 |
| `28a` (check 28) | 12 (0) / 12 (0) | 12 (0) / 12 (0) | 12 (0) / 12 (0) |
| `25a`, `25e` (reject) | 6 (6), 2 (2) | 0, 0 | 0, 0 |
| `26a`, `26d` (reject), `33a` (accept) | 0 | 0 | 0 |

The "before" column reproduces docx4j's account (CR-003 section 22.2) line for line, which is the
check on the description's port. What is left with the records shown is the table, grid, row,
cell and section records (C23's merged grid and C07's margin, check 18's table and final section,
check 28's two tables): Phase B. What the body's verbs leave beyond that is the probe's header,
footer and endnote revisions and check 18's footnote revision, which are those stories' own
(section 3.3). `22-deleted` has no whole-document Word save; its one-change saves pass.

**The timings**, CR-003 section 16.4's measurement redone (its 200-page document, 1,950 tracked
replacements, 3,900 revisions; saved, loaded and its body read; the median of five, this machine):

| call | 0.1.0 | Phase A |
|---|---:|---|
| `body.get_tracked_changes()` | 10.6 ms | 21.0 ms |
| `body.accept_all()` | 32.1 ms | 69.9 ms |
| `body.reject_all()` | 40.5 ms | 86.7 ms |
| `pkg.outline()` | 35.5 ms | 44.6 ms |

The relist costs **2.2 times** the single pass (risk 2's bound is "a few times"): `accept_all`
lists twice --- once to resolve, once to find nothing left --- where Phase F listed once, and the
sweep and the side effects' settling are passes of their own. Two optimisations made it so: a run
is searched for text boxes only when it holds a `w:drawing`, `w:pict`, `w:object` or
`mc:AlternateContent` (the first measurement was 100 ms to list and 236 ms to accept), and the
sweep walks content lists only.

**The acceptance artefacts**: `scripts/acceptance.py` into a scratch directory and
`scripts/compare_acceptance.py` against `out/acceptance/`: all 11 artefacts identical,
canonically, dates scrubbed. Nothing was regenerated and no Word check is asked for: that is
Phase C's.

### 10.6 Phase F's tests that Word contradicts, changed

Each was changed to Word's rule, not the other way round; the row of docx4j CR-043 section 3.2 (or
3.1) that settles it is named.

| test (`tests/content/test_tracking.py` unless named) | was | is | settled by |
|---|---|---|---|
| `test_a_paragraph_inserted_in_the_middle_carries_its_own_mark` | `[("mark", "Added"), ("run", "Added")]` | `[("insertion", "Added", "A new paragraph.\r")]`, pieces `run_insertion`, `paragraph_mark_insertion` | 3.1 grouping (check 23, "ins-paragraph") |
| `test_a_paragraph_appended_at_the_end_leaves_the_final_mark_alone` | two changes; `accept_all() == 2` | one, `"\rAppended at the end."`; `== 1` | 3.1 grouping |
| `test_rejecting_an_appended_paragraph_joins_it_into_the_one_before`, renamed `..._into_the_one_after` | `reject_all() == 2`; the surviving paragraph is the first's id and keeps `Heading1` | `== 1`; the survivor is the appended paragraph, its id and its properties | 3.2 "inserted paragraph mark: reject joins with the next, keeping the next's properties" (C10) |
| `test_a_fragment_appended_at_the_end_shifts_every_mark_back_one` | `reject_all() == 6` | `== 1` | 3.1 grouping |
| `test_deleting_a_paragraph_marks_the_mark_and_the_content` | `[("mark", "Deleted"), ("run", "Deleted")]` | `[("deletion", "Deleted", "The quick … dog.\r")]` | 3.1 (check 22: a deleted paragraph is one change) |
| `test_a_deleted_row_is_emptied_as_Word_empties_one` | `("row", "Deleted", "a\nb")` | `("row_deletion", "Deleted", "a\tb\r\n")` | 3.1, a row's text (checks 30, 31) |
| `test_a_row_delete_and_a_table_delete_mark_rather_than_remove` | two row changes | one, of two pieces | 3.1, rows that touch group (check 31) |
| `test_an_inserted_table_marks_every_row_and_every_paragraph` | `[("row", "Added"), ("row", "Added")]`; `reject_all() == 2` | `[("row_insertion", "Added")]`, two pieces; `== 1` | 3.1 (check 31) |
| `test_markdown_inserted_while_tracking_is_on_is_one_report_and_all_insertions` | `len(added) >= 4` | one change of four or more pieces | 3.1 grouping |
| `test_the_tracked_changes_of_a_word_written_document` | five types, `changes[1].kind == "mark"` | three, `changes[1]` an `insertion` of a run and a mark | 3.1 grouping |
| `test_accepting_a_deleted_paragraph_mark_joins_with_the_next` | the first paragraph's `w14:paraId` survives | the second's | 3.2 "deleted paragraph mark: the result keeps the next paragraph's `w:pPr` and `w14:paraId`" (C10b) |
| `test_rejecting_both_mark_forms_of_a_loaded_document_leaves_the_source` | kinds `mark, run, mark, run`; counts 4; rejecting `reversed(get_tracked_changes())` one by one | two `insertion`s; counts 2; list, reject the last, list again | 3.1 grouping; a view is valid until the next accept or reject (3.1) |
| `test_rejecting_every_tracked_verb_gives_the_document_back` (3 documents) | the verbs appended a paragraph and markdown at the body's end | they insert them after the second paragraph; the body is still canonically the source | C10 (next row) |
| new: `test_rejecting_what_was_appended_at_the_end_follows_word` (3 documents) | --- | appended at the end and rejected: the text and every paragraph id but the last are the source's, no markup is left, and the last paragraph is the appended one | 3.2, C10 |

And for the vocabulary (section 3.2's `kind`, not a Word rule): `kind` values `"run"`, `"mark"`,
`"row"` became docx4j's in `test_a_run_this_author_inserted_takes_formatting_with_no_r_pr_change`,
`test_row_insertions_and_deletions_go_on_the_tr_pr_and_the_row_stays`,
`test_a_content_control_inherits_tracking_and_its_delete_does_not`,
`test_the_moves_the_formatting_change_and_the_rows_of_the_built_document` and
`test_a_tracked_change_has_a_range_a_repr_and_a_to_dict` (whose `to_dict` also gains `"part"`, as
does `tests/agent/test_audit_trail.py::test_the_audit_trail_is_the_documents_own`).

**What the C10 change means for this package's own markup.** A paragraph appended at the end of a
body carries no mark: Phase F's tracker marks the paragraph before it (CR-003 section 16.10), as
Word does when Enter is pressed at the end. Rejected, Word joins that paragraph into the appended
one and keeps the appended one's properties and id --- and so does this engine now. Word's own
form also records a `w:pPrChange` on the new paragraph when its properties differ (check 29),
which is what carries the original properties back; the tracker writes none, so rejecting an
append at the end of a Heading 1 leaves a Normal paragraph, in Word and here. That is the
tracker's to fix (write check 29's `w:pPrChange`, or give the appended paragraph the properties it
takes over), outside this CR's scope (section 2); Phase F's 16.10 heuristic hid it on this
package's side only. *Done in Phase C (section 12.2): the tracker writes
check 29's `w:pPrChange`.*

### 10.7 What Phase B needs

- `tables.py`'s resolving half (docx4j `Tables`): `accept` / `reject` of a table's records with
  the canonical comparison (`record_of` listing only records that change something, phase 2's
  `change_text`), the merged grid collapsed and a row's redundant `w:tblPrEx` dropped (C23),
  `w:sectPrChange` merged (`rejectSection`, C07, U13), cells inserted and deleted (`removeCell`,
  check 27), `w:cellMerge` (U8); `resolver.py`'s `TableT` / `SectionT` / `CellT` branches, which
  return now; `change.UNRESOLVED_KINDS` emptied and `TrackedChange.resolvable` retired.
- The sweep's phase 2 half: records that change nothing, `w:cellMerge`, `w:numberingChange` and
  `w:numPr/w:ins` (D4, U9).
- `TrackedChangeFilter` (authors, types, kinds, dates, parts, a control's content, a bookmark's
  extent read once) in `filter.py`, behind the `filter` the module functions already take, and
  `tracked_changes(pkg, filter)`; `report(pkg)` with `to_dict()`.
- `tests/revisions/test_word_saves.py` comparing with the records shown everywhere (the
  `no_records` column goes), and docx4j's phase 2 tests (75).

Phase C, beyond section 6: the tracker's appended-paragraph form (10.6), and the `w:date`
convention (10.3 item 1), if the user wants them.

## 11. Phase B implementation notes (2026-10-16)

Phase B is done, on section 9's recommendations, ported from docx4j CR-043 phase 2 (`737c32139`,
read as it stands at `00614748e`). The records Phase A listed and left --- a table's
`w:tblPrChange`, `w:tblGridChange`, `w:tblPrExChange`, `w:trPrChange` and `w:tcPrChange`, a
section's `w:sectPrChange`, a cell's `w:cellIns` and `w:cellDel` --- are accepted and rejected as
Word's saves have them; the legacy forms nothing lists are swept; a filter and a report are
added. Against Word 365's Accept All and Reject All saves of the eleven agreement fixtures **no
line differs with the table, row, cell and section properties shown**, compared by
`description.py` and, independently, by docx4j's own Java `DocumentDescription` (11.4). No case
was found where docx4j's rule and Word's save disagree.

```python
from docx4j_py.model import revisions
from docx4j_py.model.revisions import TrackedChangeFilter

revisions.reject_all(pkg, TrackedChangeFilter(authors=["Author B"]))     # one reviewer's
revisions.accept_all(pkg, TrackedChangeFilter().types("Added").within_bookmark("Terms"))
revisions.tracked_changes(pkg, TrackedChangeFilter(kinds=[revisions.Kind.TABLE_PROPERTIES]))
revisions.report(pkg).to_dict()   # {"total": 37, "by_kind": {...}, "by_author": {...}, "by_part": {...}}
```

### 11.1 What landed

| piece | where | lines | signature |
|---|---|---:|---|
| the table, cell and section records (docx4j `Tables`, the resolving half) | `model/revisions/tables.py` | 567 (was 102) | `in_change(tr)`, `has_records(tbl)`, `record_of(tbl)` (None when the records change nothing), `change_text(tbl)`, `accept(tbl)`, `reject(tbl)`, `prune_row(tr)`, `collapse_grid(tbl)`, `remove_cell(tc, tr, removals, story)`, `cell_marker(tc_pr, local)`, `resolve_cell_merge(tc_pr, accept)`, `reject_section(sect_pr)`; `rows_of`, `cells_of` as before |
| properties through their XML (docx4j `Props.canonical`, `inner`, `convert`) | `props.py` | 268 (was 176) | `element_of(value, local)`, `inner(element, skip, out)`, `canonical(value, local, skip=None)`, `convert(recorded, local, live)` |
| the filter (docx4j `TrackedChangeFilter`) | `filter.py` (new) | 365 | `TrackedChangeFilter(*, authors, types, kinds, date_from, date_to, parts, within, within_bookmark)` and the setters `authors(*)`, `types(*)`, `kinds(*)`, `between(from, to)`, `parts(*)`, `within(control)`, `within_bookmark(name)`, each returning the filter; `range(pkg)`, `matches(piece, range)`, `bind(pkg)`; `Filter` (a `TrackedChangeFilter` or a callable), `matcher(filter, pkg)` |
| the report (docx4j `TrackedChangesReport`) | `report.py` (new) | 65 | `TrackedChangesReport` with `total`, `by_kind`, `by_author`, `by_part`, `to_dict()` |
| the module functions | `__init__.py` | 352 (was 245) | `tracked_changes(pkg, filter=None)`, `accept_all(pkg, filter=None)`, `reject_all(pkg, filter=None)`, `report(pkg)`; the sweep's phase-2 half; `TrackedChangeFilter`, `TrackedChangesReport` and `Filter` exported, `is_resolvable` gone |
| the change | `change.py` | 566 | `anchor()` (where a change stands, for a range); `UNRESOLVED_KINDS`, `is_resolvable`, `TrackedChange.resolvable` and the `tracking.unsupported` refusal gone, as docx4j's `Kind.isResolvable()` went |
| the resolver, the lister | `resolver.py`, `lister.py` | +30, +16 | the `TableT`, `SectionT` and `CellT` branches (`_Op.cell`); a table listed through `record_of`; `box_refs` public for the sweep |

`Body.accept_all()`'s docstring says the records are resolved; nothing else in the content API
changed, because its verbs already delegated to the engine: `TrackedChange.accept()` / `reject()`
on a record now works, and `accept_all()` counts records. `codegen/generate.sh --check` reproduced
the model byte for byte with the two new modules in place (`KEEP` holds `model` whole).

### 11.2 The rules as built

Every row of docx4j CR-043 section 3.2 that phase 2 covers, and section 11's U8, U9, U10, U13,
U19 and U20, as docx4j builds them:

1. **Listed when they change something** (checks 24 and 26's rule; U20): a table's records are
   compared as canonical XML (attributes sorted, rsids left out) with the record and the row's
   and cell's own markers left out; records that change nothing, and a grid's record alone, are
   not listed. A table's change's text is every row's where the table's own properties or grid
   changed, else the changed rows': the probe's C23 is `"C23 r2c1\tC23 r2c2\r\n"`.
2. **Accepted**: every record goes, a `w:trPr` left empty with it.
3. **Rejected**, wholesale: the table's `w:tblPr` and grid, and each row and cell in the change,
   as recorded (through their XML, `props.convert`); a row in the change with no `w:trPrChange`
   cleared (check 28); the row's own `w:ins` / `w:del` and the cell's `w:cellIns` / `w:cellDel` /
   `w:cellMerge` kept. A merged grid is collapsed: the column edges no row's cells fall on are
   merged, the columns past every row's last cell dropped with the rows' `w:gridAfter` /
   `w:wAfter`, the spans recounted --- C23's five columns (720, 1280, 720, 1280, 720) back to two
   (2000, 2000), as Word's save has them.
4. **Either way**, a row's `w:tblPrEx` that only repeats the table's `w:tblPr` goes (C23).
5. **A section's record rejected is merged** (check 18, C07): the recorded members laid over the
   section, the rest left, a recorded `w:docGrid` of pitch 0 meaning none; the header and footer
   references stay (U13). Accepted, the record goes.
6. **Cells** (check 27, U10): an inserted cell accepted, or a deleted one rejected, loses its
   marker; otherwise the cell goes and the one before it (else after it) spans its grid columns,
   the widths added when both are `dxa`.
7. **The sweep** (D4), after the rounds and only with no filter: records that change nothing,
   accepted either way; `w:cellMerge`, a rejected one giving back the `w:vMerge` its
   `w:vMergeOrig` recorded (U8); `w:numberingChange` in `w:numPr` and in `w:fldChar`, dropped
   either way; `w:numPr/w:ins`, the marker dropped on accept, the `w:numPr` on reject (U9); and
   phase 1's stray `w:customXmlIns*` markers. A record's own content is never looked into.
8. **The filter** (section 3.1): every condition set must hold; it applies to single changes,
   so a group partly matched has only its matching pieces resolved (U19); a range --- a content
   control's content, a bookmark's extent walked in document order with a paragraph's mark after
   its content --- is read once, before anything is resolved; `tracked_changes(pkg, filter)` lists
   the changes any of whose pieces it matches; `accept_all` / `reject_all` count those, and sweep
   nothing (what is swept belongs to no change).
9. **The report**: counts by kind (in `Kind`'s order), author and part (in the order first
   listed). There is no unresolved count: every kind listed is resolved, so the probe's
   `accept_all` now counts 37 where Phase A counted 35.

### 11.3 Departures

None of docx4j's rules is built differently because Word's saves say otherwise: all fifteen
comparisons are at zero with docx4j's rules as they are. What differs is vocabulary and
mechanics:

1. **The filter's vocabulary is Python's.** The conditions are keywords of the constructor as
   well as docx4j's chained setters; types are Office JS's strings (`"Added"`, `"Deleted"`,
   `"Formatted"`), kinds a `Kind` or its value; a part may be given by its name
   (`"/word/footnotes.xml"`) as well as itself; `within` takes the model's control or the content
   API's `ContentControl`; a naive datetime is read as UTC. A plain callable over a single change
   is still a filter (10.3 item 5), with no range and no sweep.
2. **A change in a text box kept as a wildcard tree is inside a range when its text box was.**
   Its markup is a typed island made after the range is read (3.5), so the range holds the
   wildcard `w:txbxContent`, not the island's objects; docx4j holds the typed objects themselves.
   The difference shows only for a change a resolution brings out inside the box in a later round
   (a move's nested deletion), which is outside the range here.
3. **A bookmark in a story part with no revision markup** is found in the part's bytes, without
   unmarshalling it (10.3 item 6's rule), and gives an empty range; docx4j walks every story. A
   name found nowhere raises `ValueError` (docx4j `IllegalArgumentException`).
4. **`report(pkg).to_dict()`** is docx4j's `toJson()` in Python's names: `total`, `by_kind` (the
   kinds' snake-case values), `by_author`, `by_part`; a change with no author counts under `""`
   where docx4j writes `"null"`.
5. **Mechanics with no difference in behaviour**: `props.convert` and `canonical` go through
   `to_xml` and the runtime parser (a conversion that would lose anything raises); the sweep walks
   content lists only --- every form it sweeps stands on a table, a cell, a paragraph or a run ---
   and makes a text box an island only when its wildcard tree holds revision markup, writing it
   back only when the sweep changed it; a cell's `w:cellIns`, `w:cellDel` and `w:cellMerge` are
   one choice field in the generated model (`cell_ins_or_cell_del_or_cell_merge`), as the schema
   has them.

### 11.4 The numbers

**The tests.** `tests/revisions/`: **360** (348 before), of which 283 read docx4j's sibling
checkout and skip without it.

| file | tests | what changed |
|---|---:|---|
| `test_tracked_changes.py` | 42 (30) | docx4j's phase 2 tests ported: `tableRecords`, `recordsThatChangeNothing`, `sectionRecordMerges`, `cellsInsertedAndDeleted`, `legacyFormsSwept`, `filterByAuthorAndType`, `filterByKindSplitsAGroup`, `filterByDate`, `filterWithinABookmark`, `filterWithinAControl`; and a field's `w:numberingChange` swept, a record resolved through the content API, a filter by part (a part or its name), a filter with no condition sweeping nothing, the report over the probe; three Phase A tests replaced (11.5) |
| `test_word_saves.py` | 31 (31) | every comparison with the records shown |
| `test_every_fixture.py` | 257 (257) | nothing at all left, not "nothing resolvable" |
| `test_one_change.py` | 30 (30) | unchanged |

The fast suite is **2,358 passed, 8 skipped** (2,346 and 8 before), the whole suite **2,398
passed, 8 skipped** (2,386 and 8).

**Word's saves.** Each agreement fixture resolved by `revisions.accept_all` / `reject_all` (every
story), saved, and compared with Word's own save; lines differing with the records shown, and in
brackets with them left out, by this package's `description.py` and by docx4j's
`DocumentDescription` (its Java class at `00614748e`, run over the same saved files):

| comparison | before, `description.py` | before, Java | after, `description.py` | after, Java |
|---|---|---|---|---|
| `cr043-tracked` accept | 6 (0) | 6 (0) | **0 (0)** | **0 (0)** |
| `cr043-tracked` reject | 16 (0) | 16 (0) | **0 (0)** | **0 (0)** |
| `revisions-word15` accept (check 18) | 8 (0) | 8 (0) | **0 (0)** | **0 (0)** |
| `revisions-word15` reject | 10 (0) | 10 (0) | **0 (0)** | **0 (0)** |
| `28a-changes` accept (check 28) | 12 (0) | 12 (0) | **0 (0)** | **0 (0)** |
| `28a-changes` reject | 12 (0) | 12 (0) | **0 (0)** | **0 (0)** |
| `tracked-changes-equations` accept, reject; `review-for-word` accept, reject (Word 15's); `26a`, `26d`, `25a`, `25e` reject; `33a` accept | 0 (0) | 0 (0) | 0 (0) | 0 (0) |

`test_word_saves.py` holds the same and Word 2010's two saves of `review-for-word`. Through the
body's verbs (`body.accept_all()`, which docx4j's `py_resolve.py` calls) what is left is what
Phase A left with the records hidden --- the probe's header, footer and endnote revisions, 12 / 18
lines, and check 18's footnote revision, 2 / 0 --- which are those stories' own (3.3); the records
add nothing.

**The timings**, Phase A's 200-page benchmark (`body`'s 3,900 revisions; the median of five, two
runs each, this machine): no change.

| call | Phase A | Phase B |
|---|---:|---:|
| `body.get_tracked_changes()` | 20.8, 21.5 ms | 21.4, 21.1 ms |
| `body.accept_all()` | 70.5, 69.5 ms | 70.7, 69.8 ms |
| `body.reject_all()` | 86.9, 85.6 ms | 83.4, 84.5 ms |
| `pkg.outline()` | 44.1, 44.0 ms | 43.6, 43.0 ms |

**The acceptance artefacts**: `scripts/acceptance.py` into a scratch directory and
`scripts/compare_acceptance.py` against `out/acceptance/`: all 11 identical, canonically, dates
scrubbed. Nothing was regenerated and no Word check is asked for: that is Phase C's.

### 11.5 Tests changed

Each encoded Phase A's "listed but not resolved"; none encoded a rule Word contradicts.

| test | was | is |
|---|---|---|
| `test_word_saves.py::CASES` and `test_resolving_every_story_gives_words_save` | a fourth column, "holds no Phase B records"; compared with `records=False`, and with `records=True` only where the column was set | three columns; every case compared with `records=True` (a superset of `records=False`'s lines) |
| `test_word_saves.py::test_a_bodys_accept_all_gives_words_save_where_the_body_is_the_document` | the body's listing empty or holding only table and section records; `records=False` | the listing empty; `records=True` |
| `test_word_saves.py::test_the_probe_resolved_leaves_only_phase_bs_records_listed` | count 35; the two records left, not resolvable | renamed `test_the_probe_resolved_leaves_nothing_listed`: 37 listed, a table's and a section's records among them; count 37; nothing left |
| `revision_helpers.unresolved` (used by the three files above and `test_every_fixture.py`) | what is left that `resolvable` says this release resolves | everything left |
| `test_every_fixture.py::test_every_change_one_by_one` | the listing filtered by `resolvable` | the whole listing |
| `test_tracked_changes.py::test_a_tables_records_are_listed_and_left_for_phase_b` | listed, `resolvable` false, `reject()` raising `tracking.unsupported`, `accept_all` counting 1 and leaving it | `test_a_tables_records_are_one_change_accepted_or_rejected`, docx4j's `tableRecords`: rejected restores the recorded `w:tblW`, the report counts it, `accept_all` counts 1 and keeps the current one |
| `test_tracked_changes.py::test_a_sections_record_is_listed` | listed, counted by `outline()` | `test_a_sections_record_rejected_is_merged`, docx4j's `sectionRecordMerges`, the listing and the `outline()` count kept |
| `test_tracked_changes.py::test_a_cell_insertion_is_listed` | listed | `test_u10_cells_inserted_and_deleted`, docx4j's `cellsInsertedAndDeleted` |
| `test_tracked_changes.py::test_every_story_is_resolved_by_the_package_verbs` | the body's records left | nothing left |
| `test_tracked_changes.py::test_a_bodys_verbs_resolve_its_own_story_and_leave_the_others` | the main document part among the parts left, for its records | exactly the header, footer and endnotes parts left |

`test_a_filter_resolves_the_pieces_it_matches` is kept, its docstring saying it holds the plain
callable. No test of CR-003 Phase F's (`tests/content/`, `tests/agent/`) encoded the records, and
none changed.

### 11.6 What Phase C needs

Section 6 as written: the content API's reads agreeing with the engine (`OutlineStats`, the
markdown views and a test of the accepted view over the fixtures), the README and the PyPI page
where counts changed, the CHANGELOG's "Changed" and "Fixed" entries, artefact 8
regenerated with an accepted and a rejected copy and a Word check, CR-003 sections 16 and 22
pointed here, and the hand-off to docx4j's session to rerun the agreement --- where python's column
through the body's verbs is now the other stories' revisions alone (11.4). Beyond it, if the user
wants them: the tracker's appended-paragraph form (check 29, 10.6) and the `w:date` convention
(10.3 item 1). *Phase C (section 12) did all of it but the Markdown items, which moved to CR-011,
the agreement rerun, which is docx4j's session's, and the `w:date` convention, left out.*

## 12. Phase C implementation notes (2026-10-18)

Phase C is done, on section 9's recommendations, with section 6's two Markdown items moved to
CR-011 by the user's decision (12.3). The content API's counts are held to the engine's listing
(12.4); the text and Markdown views' disagreement with `accept_all` is measured and pinned for
CR-011; the tracker writes Word's check 29 form for a paragraph appended at the end of a
container, which section 10.6 left open, so that rejecting an append after a heading gives the
heading back (12.2); acceptance artefact 8 ends with a heading and has an accepted and a rejected
copy beside it, with a script to put Word's own saves against them (12.5); and the CHANGELOG,
CR-003 sections 16 and 22 and CR-006 section 7 point here. The fifteen comparisons with Word's
saves are still at zero, by `description.py` and by docx4j's Java `DocumentDescription`. **No
Word check has been made**: the user's check of artefact 8 is what this phase waits on.

```python
from docx4j_py.model import revisions
from docx4j_py.model.content import Author

pkg.author = Author("Author B")
pkg.change_tracking_mode = "TrackAll"
pkg.body.insert_paragraph("New")     # the body ends with a Heading 1: Word's 29a form
# the heading:  <w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:ins .../></w:rPr></w:pPr>
# "New":        <w:pPr><w:pPrChange ...><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr>
revisions.reject_all(pkg)            # one paragraph, "Case a heading", Heading 1
```

### 12.1 What landed

| piece | where | lines | signature |
|---|---|---:|---|
| check 29 in the tracker | `model/content/tracking.py` | 1,311 (was 1,149) | `final_mark_properties(tracker, previous)` (the old final mark's properties, read before the shift), `record_final_mark(tracker, paragraph, original)` (29a's `w:pPrChange` when they differ), `same_properties(p_pr, original)` (`CT_PPrBase`'s members compared in place), `holds_final_mark(paragraph, container)`, `settle_final_mark(tracker, paragraph, container)`, `hand_back_final_mark(tracker, leaving, previous)`; `track_inserted_blocks` calls the first two |
| the paragraph view | `model/content/paragraph.py` | 1,682 (was 1,661) | `_p_pr()` records a `w:pPrChange` on the paragraph holding the old final mark, as on one the document had; `formatting(...)` settles the record after a write (`_settle_final_mark`); `_tracked_delete` hands the record back when an append is taken back |
| `OutlineStats.tracked_changes`'s docstring | `model/content/reports.py` | +5 | it said "`w:ins`, `w:del` and `w:rPrChange` elements", stale since CR-003 Phase F; now the engine's listing over the outline's stories |
| acceptance artefact 8 | `scripts/acceptance.py` | 717 (was 687) | a Heading 1 ends the body before the mode goes on; `8-tracked-changes-accepted.docx` and `8-tracked-changes-rejected.docx` written beside it by `revisions.accept_all` / `reject_all`, each asserted to list nothing when read back; a build may return several files |
| Word's saves against ours | `scripts/compare_word_resolution.py` (new) | 97 | `compare(ours, word, *, records=True) -> (lines, diff)`; `main()`: artefact 8's two pairs (`out/acceptance/8-tracked-changes-{accepted,rejected}.docx` against `out/acceptance-word/8-word-{accepted,rejected}.docx`) or any pairs given; exit 0, 1 on a difference, 2 on a missing file |
| the tests | `tests/revisions/test_appended_paragraph.py`, `test_content_api.py`, `test_compare_word_resolution.py` (new) | 300, 119, 75 | 12.5 |
| the fixtures | `tests/fixtures/revisions/core-ts/check29/` | --- | `29-input.docx`, `29a-heading-end-tracked.docx`, `29b-normal-end-tracked.docx`, from docx4j-core-tests at `00614748e` (docx4j-core-ts's check 29, `25b50f5`); provenance in `tests/README.md` |

`codegen/generate.sh --check` reproduced the model byte for byte; `KEEP` holds `model` whole.
Nothing in `docx4j_py/model/revisions/` changed: the engine already resolves Word's 29a markup as
Word does, so check 29 was the tracker's alone.

### 12.2 Check 29, as built

docx4j-core-ts's check 29 (its CR-002, "Enter at the end of a paragraph") measured Word's Enter
at the end of a paragraph with tracking on: the paragraph before the caret ends with the **new**
mark, inserted by the typist, and the new line keeps the **old** one. After a Heading 1, whose
next style is Normal, the new line turns Normal and Word records that in a `w:pPrChange` on it
holding Heading 1 (29a); after a Normal paragraph, nothing (29b). CR-003 section 16.10's shift
already gave an appended paragraph the old final mark; the tracker now gives it the record too.

1. **Where**: wherever 16.10's shift applies --- the insert's paragraphs end their container and
   a paragraph the document had stands before them --- which is every route that appends, since
   all go through `Body.insert_element`: `Body.insert_paragraph`, `Paragraph.insert_paragraph`
   (`"After"` on the last paragraph), `insert_xml`, `insert_ooxml`, `insert_markdown`, a cell's
   and a content control's `insert_paragraph`, and the facade's `add_paragraph` and
   `add_heading`. In the middle of a container nothing changes: each inserted paragraph carries
   its own mark, and there is no old mark to record.
2. **What is recorded**: the properties of the paragraph that ended the container before the
   insert, as `CT_PPrBase` holds them (the mark's `w:rPr`, the `w:sectPr` and any record left
   out) --- or what that paragraph's own `w:pPrChange` records, the first original winning as
   `record_p_pr_change` has it. A record this session wrote on a paragraph it appended earlier
   is the old mark's and moves on with it: two appends after a heading leave the record on the
   second only.
3. **When**: only when the insert's last paragraph's properties differ, member by member.
4. **After the insert**: the paragraph holding the old final mark (`holds_final_mark`: last in
   its container, its own mark unmarked, the mark before it inserted) records a `w:pPrChange` on
   a property write as a paragraph the document had does, where a paragraph this session
   inserted otherwise records none (16.12 item 4). So `insert_paragraph(style=...)` and the
   facade's `add_heading`, which set the style after the insert, record it against the old
   mark: a Heading 2 appended after a Normal paragraph records Normal's empty `<w:pPr/>`. A write that
   brings the paragraph back to what its record holds drops the record (`settle_final_mark`, after
   every write through `formatting`), so `insert_paragraph("New", style="Heading 1")` after a
   heading records nothing, as 29b would have it.
5. **Taken back**: `Paragraph.delete()` on an appended paragraph (this author's own insertion,
   16.10) hands its record back to the paragraph before it when that is one this session
   appended; a paragraph the document had keeps its own properties, which are the old mark's.
6. **Resolved**: by the engine unchanged. Rejected, the record restores Heading 1 and the join
   keeps the next paragraph's properties (C10), so the heading comes back; accepted, the record
   goes and the appended paragraph keeps its own.

The markup, ours over 29-input cut after the heading, beside Word's 29a (ids and dates aside):

| | the heading's `w:pPr` | the new paragraph |
|---|---|---|
| ours | `<w:pStyle w:val="Heading1"/><w:rPr><w:ins w:id="2" w:author="Author B" w:date="…"/></w:rPr>` | `<w:pPr><w:pPrChange w:id="3" w:author="Author B" w:date="…"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr><w:ins w:id="1" …><w:r><w:t>New</w:t></w:r></w:ins>` |
| Word's 29a | `<w:pStyle w:val="Heading1"/><w:rPr><w:ins w:id="0" w:author="Author B" w:date="…" w16du:dateUtc="…"/></w:rPr>` | `<w:pPr><w:pPrChange w:id="1" w:author="Author B" w:date="…" w16du:dateUtc="…"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr><w:ins w:id="2" …><w:r><w:t>New</w:t></w:r></w:ins>` |

The one difference is `w16du:dateUtc`, which this tracker does not write (10.3 item 1, out of
this phase's scope). Described by `description.py` and by docx4j's Java `DocumentDescription`,
ours and Word's are identical for 29a and 29b, tracked, accepted and rejected: six pairs, by each.

Departures, each where Word is unmeasured:

1. **Another author's record on the paragraph before** is what the new record records, and is
   left where it is. Word's 29d moves an old mark's own revision on to the new line; this does
   not move another author's revision, only this session's own record (item 2).
2. **The mark's run properties** (`w:pPr/w:rPr`) are neither compared nor recorded: 29a's heading
   has none, and `w:pPrChange` cannot hold them.
3. **With no paragraph before** (an empty container, or a table) nothing is recorded, as 16.10
   marks nothing there.

### 12.3 The views: moved to CR-011

Section 6 asked for `to_markdown(view="accepted")` to read as `accept_all()` would leave the
document, `view="markup"` to show every listed change, and a test of the accepted view's
agreement. On 2026-10-17 the user accepted every recommendation of the portfolio's
text-extraction paper (`../docx4j-portfolio/docs/text_extraction_behaviour_model.md` section 9,
portfolio `2d8fd36`). Its section 3.1 defines the views **by resolution** --- `accepted` is the
text after Accept All on every story --- where both ports filter runs by their holders, and its
section 10 gives docx4j-python a CR-011, "one text model replacing the five", which "folds CR-010
phase C's Markdown agreement"; its section 10a sequences the ports after docx4j CR-046 phase 0's
goldens. So the three items are CR-011's, and none was built here.

What this phase does instead is pin the disagreement, measured over the eleven agreement
fixtures: `get_text(view="accepted")` and `to_markdown(view="accepted")` against the same view
after `revisions.accept_all`. Lines differing (ndiff's `+` and `-` lines):

| fixture | `get_text` | `to_markdown` | what |
|---|---:|---:|---|
| `cr043-tracked` | 15 | 14 | C06's, C09's and C10b's deleted marks not joined; C13's deleted table and a deleted row of C14's left as empty lines |
| `revisions-word15` | 7 | 5 | a deleted mark not joined ("Join this paragraph with the next one."); a moved paragraph's emptied source and a deleted row's three cells as lines |
| `review-for-word` | 7 | 5 | the same deleted mark; a deleted row's cells and a deleted paragraph as lines |
| `33a-deleted` | 1 | 2 | the deleted heading left (an empty line; `##` in the Markdown) |
| `22-deleted` | 5 | 6 | the paper's example: `Four.\nFive \neight.` for `Four.Five eight.`; "Two." as an empty line |
| the other six | 0 | 0 | |

`tests/revisions/test_content_api.py` holds the six that agree and pins the five that differ as
`xfail(strict=True)`, each with its reason, for both views: CR-011 closing one turns its test
into an unexpected pass, which fails the suite until the pin is removed. Measured too and not
pinned, for CR-011: `get_text(view="original")` against the text after `revisions.reject_all`
differs on three (`cr043-tracked` 13 lines, `revisions-word15` 7, `review-for-word` 10). The
`markup` view was not measured against the listing: what it should show for a mark, a row or a
property change is the paper's D9, CR-011's to build.

### 12.4 The counts, as found

`OutlineStats.tracked_changes` already counted the engine's listing (Phase A rewired it through
`count_tracked_changes`, the same `Lister` over `body_stories(body)` that `get_tracked_changes()`
lists), and `describe()`'s `tracking_on` reads `w:trackRevisions`, which no resolution writes
(section 3.3). Both hold over the eleven fixtures, and `test_content_api.py` now says so:

- **every body's** `outline()` count is what that body's `get_tracked_changes()` lists, and what
  the engine lists in that body's part;
- **the package's** `outline()` adds the headers' and footers', which are the stories an outline
  shows: on the probe 32 of `pkg.get_tracked_changes()`'s 37, the five endnote revisions outside
  it; on check 18's file 26 of 28, the two footnote revisions outside it; elsewhere equal.
  `pkg.get_tracked_changes()` and `revisions.report(pkg).total` count every story;
- after `revisions.accept_all`, saved or not, every count is 0;
- `tracking_on` is `change_tracking_mode != "Off"`, and is the same after `accept_all` and
  `reject_all`, saved and loaded.

The only change was the field's docstring (12.1). Whether the package outline should count the
notes' and comments' revisions as well is a question for the user, not a defect: an outline
shows neither, and its other counts (paragraphs, words) are of the same stories.

### 12.5 The numbers

**The tests.** `tests/revisions/`: **454** (360 before), of which 283 read docx4j's sibling
checkout and skip without it, as before.

| file | tests | what |
|---|---:|---|
| `test_appended_paragraph.py` (new) | 19 | check 29: the `w:pPr` written beside 29a's; 29a and 29b described as Word's files are (the paragraphs after the case cut from both); resolved both ways as Word's markup resolves; Reject All giving one Heading 1, Accept All both paragraphs with their own properties; every route that appends, the facade's and a table cell's included; `Paragraph.insert_paragraph("After")` copying a heading and needing no record; a style set after the insert; a fragment; the record moving with the old mark and coming back when the append is taken back; none in the middle |
| `test_content_api.py` (new) | 55 | 12.4's counts (11) and `tracking_on` both ways (22); 12.3's views (22: 12 passing, 10 `xfail`) |
| `test_compare_word_resolution.py` (new) | 20 | the script: each of `test_word_saves.py`'s 17 cases resolved, saved and compared with Word's save at 0 lines; artefact 8's layout with the probe's files; a difference counted and shown; pairs given, a save missing |

The fast suite is **2,442 passed, 8 skipped, 10 xfailed** (2,358 passed and 8 skipped before),
the whole suite **2,482 passed, 8 skipped, 10 xfailed** (2,398 and 8).

**Word's saves.** The fifteen comparisons rerun (each agreement fixture resolved by
`revisions.accept_all` / `reject_all`, saved, described against Word's save): **0 lines** in every
one, with the records shown and without, by `description.py` and by docx4j's Java
`DocumentDescription` --- as Phase B left them. `test_word_saves.py` holds the same.

**Tracked appends**, 1,000 paragraphs appended at the end of a created body with tracking on
(the median of five, this machine; before is `da7f8fd`):

| appends | before | after |
|---|---:|---:|
| Normal, nothing recorded | 58 ms | 63 ms |
| alternating Heading 1 and Normal, every append recording | 69 ms | 87 ms |
| `style="Heading 1"` on each after a heading: a record written at the insert, settled by the style | 76 ms | 108 ms |

The first version deep-copied both `w:pPr`s to compare them (111 ms alternating);
`same_properties` compares members in place, and the record is copied only when it is written.
The listing and resolving timings are Phase B's: nothing in the engine changed.

**The acceptance artefacts.** `out/acceptance/` was copied aside, regenerated with
`scripts/acceptance.py` and compared with `scripts/compare_acceptance.py` over the eleven
artefacts both have: 147 entries compared, **ten artefacts identical** once the dates are
scrubbed (1 to 7 and 9 to 11), and **artefact 8 differing**, in exactly `word/document.xml` (the
heading, its inserted mark, the record on the appended paragraph, the ids after them) and
`word/styles.xml` (the `Heading1` definition the style setter adds, the source defining none).
Artefact 8 lists 9 changes: the replacement's two, the paragraph inserted near the top, the
deleted and inserted rows, the deleted paragraph, the bold run, the appended paragraph with the
heading's mark (`"\rAnd one appended at the very end."`), and the appended paragraph's record.
Its two copies list none, read back. Rejected, the body ends with "A heading the appended
paragraph follows" in Heading 1 (before this phase, Normal); accepted, with that heading in
Heading 1 and the appended paragraph after it in Normal. The regeneration is recorded in `tests/README.md`, and artefact
8's row says what to look for.

### 12.6 Tests changed

| test | was | is | why |
|---|---|---|---|
| `tests/content/test_tracking.py::test_rejecting_an_appended_paragraph_joins_it_into_the_one_after` | `reject_all() == 1`; the survivor is the appended paragraph with its own properties, `style_id != "Heading1"`, "as Word's reject leaves it" | `== 2` (the group and the record); the survivor is still the appended paragraph (its id, C10), now `style_id == "Heading1"`, and no `w:pPrChange` is left | the appended paragraph records the heading it followed, as Word's 29a does (12.2); 10.6 had named this as the tracker's to fix |
| `tests/content/test_tracking.py::test_rejecting_what_was_appended_at_the_end_follows_word` (3 documents) | its docstring: the last paragraph "carries the appended paragraph's properties", the missing `w:pPrChange` being the tracker's, out of scope | one assertion more: the last paragraph's `w:pPr` is the source's last paragraph's, and the docstring says why | `sample-docx.docx`'s last paragraph states `w:ind w:left="0"`, which the appended paragraph now records and the reject gives back; the assertion fails on the tracker before this phase, for that document |

No other existing test changed. The invariant test beside them (`test_rejecting_every_tracked_verb_gives_the_document_back`,
the body canonically the source after every tracked verb and `reject_all()`) inserts in the
middle since Phase A (10.6), and passes unchanged.

### 12.7 What is left

- ~~**The user's Word check of artefact 8**~~ **made 2026-10-21** (section 12.8): the heading
  and the appended paragraph as designed, Accept All at 0 lines from Word's, and one
  disagreement left, the comment on Reject All, whose fix is the user's decision.
- ~~**The hand-off to docx4j's session**~~ **done 2026-10-18.** Asked the same day; docx4j's
  session reran CR-043 section 14's python column with `PortsAgreement` (docx4j `1b3a68ccf`, this
  repository at `e690217`, read only) over the eleven fixtures: through
  `revisions.accept_all` / `reject_all`, **0 lines differ in all 22 comparisons** (11 fixtures,
  accept and reject; section 14 had had the probe at 42 / 60). Through `Body.accept_all`, the
  probe differs by 12 / 18 lines and check 18 by 2 / 0, the rest 0, as measured here (11.4), and
  that session's grouping by story puts every one of those lines in a header, a footer, the
  footnotes or the endnotes, none in the main story: section 3.3's design. On this package's
  suggestion its `py_resolve.py` now calls `revisions.accept_all` / `reject_all` ("python: every
  story" in its README), and section 14 carries a dated note of the rerun and of the departures
  of 10.3 item 1 and 11.3; committed in docx4j as `ad0aa771e` (`VERSION_17_3_4`, unpushed).
- **CR-011**: the text and Markdown views by resolution (12.3).
- **The `w:date` convention** (10.3 item 1), if the user wants it: Word's local wall-clock
  `w:date` and `w16du:dateUtc`, the one difference left in 12.2's markup.

### 12.8 The Word check (2026-10-21)

The user opened artefact 8 and its two copies in Word 16 (`AppVersion` 16.0000 in its saves) and
saved Word's own Accept All and Reject All of it; the four files are in the user's share,
`__tmp/python-8/`. No repair prompt, on any of them. `scripts/compare_word_resolution.py` over
Word's two saves:

| comparison | lines differing |
|---|---|
| `8-tracked-changes-accepted.docx` against Word's Accept All | **0** |
| `8-tracked-changes-rejected.docx` against Word's Reject All | **5**: the comment |

The heading comes back as Heading 1 on Reject All in both, the deleted paragraph comes back, the
bold goes, and Accept All leaves the appended paragraph Normal: check 29's `w:pPrChange` does what
section 12.2 built it for.

**Three of the user's observations are what Word shows, not what it reads.** The user saw no
formatting change on the appended last paragraph, the deleted paragraph's mark not deleted, and no
formatting balloon on the bold paragraph. Word's save of the document, unresolved, keeps all three
records exactly as written (the mark's `w:del`, the `w:rPrChange` with its empty `w:rPr`, the
`w:pPrChange` holding `Heading1`), and Word resolved each one: its Accept All leaves no empty
paragraph where the deleted one was, so the mark went with its text, and its Reject All removes the
bold and gives the heading back. In the user's screenshot the deleted paragraph's ¶ is in the
revision colour where an untouched paragraph's is black. A formatting balloon depends on Show Markup
→ Balloons, and a `w:pPrChange` on a paragraph whose text is inserted is the form Word itself writes
at Enter after a heading (core-ts check 29a). The checklist row promised more than Word displays,
and is corrected (`tests/README.md`, artefact 8).

**The disagreement: Word's Reject All removes Claude's comment; the engine keeps it.** The comment
is on "second", the inserted half of `replace_text("first", "second")`. The tracker writes its range
markers and its reference run **outside** the `w:ins` (CR-003 section 16.9, the hoist of section
15.6). Word's Reject All removed everything belonging to it: the markers, the reference run and the
entry. With it went the comments, `commentsExtended`, `commentsIds` and `people` parts, since it was
the only comment. The engine keeps the comment and its reference run and drops only the range
markers. The rule it follows is docx4j's and core-ts's `removeRevision`: a comment whose markers lie
outside the revision keeps its entry and its reference, measured on Word 2010's and Word 15's Reject
All of the editor's `review-for-word.docx`, the same form. core-ts CR-002 (the item corrected by its
check 32) recorded that current Word on the outside form was not measured. **This check measures it:
current Word removes the comment.** The older Words kept it. With the start marker **inside** the
`w:ins`, which is Word's own form and core-ts's tracker's since its 0.3.1, current Word, docx4j and
both engines already remove it.

Two changes would close it, and both are the user's to make:

1. **The tracker writes Word's form**: the range start inside the `w:ins`, the end and the
   reference run after it, as core-ts's tracker has done since check 32. A comment on
   tracked-inserted text then goes with a Reject All under every rule measured: current Word's,
   docx4j's and this engine's. Accept All leaves it on its word, as now. This reverses section
   16.9's aim, a comment that survives the rejection of what it explains, which current Word does
   not allow in either form. Its cost is the tracker's marker site, the three tests 16.9 added (a
   comment now goes on `reject_all()` rather than surviving with an empty range), and a new Word
   check of artefact 8.
2. **The resolver follows current Word on the outside form**: such a comment is removed whole once
   the resolution leaves no text in its range, wherever its markers were, so files written by
   docx4j-python 0.1.0 and core-ts 0.3.0 resolve as current Word resolves them. This changes
   docx4j's rule, and the agreement run holds the ports to docx4j. It belongs in docx4j first,
   then core-ts and here. The review file's Word 2010 and Word 15 comparisons
   (`test_word_saves.py`) would then differ by that comment, and would be kept as older Word's
   behaviour.

Recommended: the first now, as the rest of Phase C, and the second handed to docx4j's session and
core-ts with this check's files, to be taken here when docx4j decides it.
