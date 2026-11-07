# CR-040: An `mc:AlternateContent` that selects nothing, read as Word reads it (WD REQ-305 and REQ-306)

**Status:** Complete 2026-11-07: Phases A to C done, the Word check passed through the bridge at the user's word (section 8.4), for this release. Decided 2026-11-07 (the user: "CR-040 as recommended, Q6 this release"). Proposed the same day ("write CR-040"), on the ruling "REQ-305 itself changes for all three engines". docx4j published the two sets at `824565363`.

**Depends on:** `portfolio/wd-req-305-follows-word`, for docx4j's published list. WD 0.98 (`d24ab36`) sets the order: the WD, then docx4j's list, then each engine. **CR-039** is built, and its `rw-24` path gets the warning here.

**Counterpart:** the rule is for all three engines, each on the user's word in its own session. Neither docx4j nor core-ts had a CR raising `mce.nothing_selected` when this was written.

## 1. Why

Where an `mc:AlternateContent` has no `mc:Fallback` and no `mc:Choice` the reader understands, REQ-305 used to select the first `Choice`, a documented departure from ISO 29500-3. Word removes such an element and shows neither branch, so the old rule returned words Word doesn't show, in a namespace of the author's choosing. WD 0.97 changed it, and 0.98 added a warning, in the user's words: "selecting nothing should raise a warning".

**The rule can't change alone.** The old rule happened to give the right answer for namespaces Word knows and the list lacks, such as `wpi`, `oel` and `cx3`. Removing it before the list grows would lose words Word shows, with nothing said. So the list comes first, and the warning makes any namespace still missing visible.

## 2. What exists here (measured 2026-11-07 at `f7c5652`)

| Route, on `rw-25` (a `Choice` Word doesn't understand, no `Fallback`, in a body) | Reads |
|---|---|
| The default load: `extract`, `Body.paragraphs` | `CHOICE WORDS`, no warning |
| `mc_load_mode="word"` | `CHOICE WORDS`: the element is kept whole, and the reader takes its first `Choice` |
| `mce_preprocess=True` | Neither branch, already as Word, with no warning |

Word shows `Case in body` and `LAST LINE` alone.

- **The rule is two lines.** `child.mce_branch` ends `return choices[0] if choices else None`, and `mc_load_mode._splice` keeps such an element whole. Every reader already takes None, since an element with no branch at all gives it: the extractor's five sites, `story.py`, `text_model.py`, `revisions/`.
- **What moves:** with both lines changed for one run, three tests failed across `tests/text`, `tests/openpackaging`, `tests/content` and `tests/test_child.py`, each pinning REQ-305 by name. The rest of the suite wasn't run.
- **`rw-24`, at a root:** CR-039 already reads an empty document, with no `mce.nothing_selected`.
- **How common:** of 133 `mc:AlternateContent` in the 719 local packages, one selects nothing, `rw-24`, under today's list and docx4j's new one.
- **The list:** this package's has 118 URIs, docx4j's at `198d015a3`. Word understands 105 of them, and the other 13 are mostly document-properties and encryption namespaces, there because readers traverse them in their own parts. So the user ruled for two sets ("two sets", through the portfolio), and WD 0.99's REQ-306 has an *understood set* for every markup compatibility question and a *traversal set* for what a reader can read. The understood set measured so far is 114 URIs: the 105 and `wpi`, `oel`, `cx3`, `cx4`, `cx6`, `cx7`, `c16r2`, `ax` and `xsi`. `test_the_understood_list_is_docx4js_verbatim` fails today because of docx4j's working copy.
- **Every use here is the understood set:** `child.understands`, `mce.py`, `mc_load_mode.py`, `repeated._refused` and `breaks._gives_its_text`. Nothing in this package asks the traversal question of the list.
- **The new set moves one selection.** `phase0/extraction-probe-inspect.docx` holds a `Choice` requiring `wpi` with a `Fallback`. Today it reads the `Fallback`, and with the understood set the `Choice`, as Word does. This is a docx4j fixture, so its expected output is docx4j's to re-baseline. Taking the 13 off moves nothing in the 719 packages.

## 3. Design

### 3.1 The understood set (Phase A)

A new resource, the understood set, in docx4j's published form once it's committed, read by `understood_namespaces()` in place of today's list. Never ahead of it: REQ-306 has additions made to the published lists first. Where the set is published and what the file is called are still open in the WD.

Today's `understood-namespaces.txt` is then the traversal set. Nothing here reads it, so it stays only if REQ-306's "publish two sets" needs it in the package (Q1). `UNDERSTOOD`, the prefix table's own narrower list, doesn't change.

### 3.2 The rule (Phase B)

- `child.mce_branch` returns None where no `Choice` is understood and there's no `Fallback`.
- `mc_load_mode._splice` takes such an element out, as it does one with no branch at all, and records it in `part.resolved_alternate_content`. WD 0.97 dropped the clause on keeping it whole.
- `mce_preprocess` already does this.
- A default load keeps the element in the tree and in a save, as every kept branch is (CR-012). Only a reader sees nothing.
- An edit can't reach it, since nothing in it has an address.

### 3.3 The warning (Phase B)

`mce.nothing_selected`, on the text result (REQ-801): selector the part's name, `detail` the namespaces the `mc:Choice` children require, as URIs. `requires_namespaces` already resolves them.

- **In a story:** one call at the extractor's five sites, where `mce_branch` answers None for an element that has Choices.
- **At a root:** a record on `part.warnings` from `repeated._taken`, carried to the text result as `part.unexpected_root` is.
- **Once per element:** the address tells them apart, and `warn` keeps each once (Q2).

### 3.4 What it leaves

- `mce_preprocess` raises no warning: it runs before any reading, and the element is gone. The text model can't see it (Q3).
- An element with no `Choice` and no `Fallback` isn't REQ-305's case, and stays silent.

## 4. Tests and oracles

- **Word's files:** `rw-25` joins `tests/fixtures/repeated/`, and the `req305` and `understood-choice` rounds' files with it (Q5). Each namespace Word understands shows the `Choice`, and each it doesn't shows neither.
- **The rule:** the three tests that pin the old rule change, by each load mode.
- **The warning:** selector, `detail`, one per element, at a root (`rw-24`), and none where a branch is selected.
- **The set:** the verbatim test held to docx4j's published understood set, with a row for each of the 13 taken off and the 9 added.
- **`phase0/`:** the suite re-pinned at the docx4j commit that holds the list, with `extraction-probe-inspect.docx`'s row as docx4j re-baselines it.

## 5. Phases

| Phase | Work | Cost |
|---|---|---|
| A | The understood set, the suite's copy, the `phase0/` re-pin | ¼ day, once docx4j publishes it |
| B | The rule, the warning, `rw-24`, the tests | ½ day |
| C | The records and the Word check of a `word`-mode save | ¼ day |

A goes first, and B never ahead of it. A is right by Word on its own.

## 6. Risks

- **A namespace Word knows and the list lacks** loses words. The warning says so, naming the URI. Eleven were found by asking, and others may exist.
- **Released behaviour changes.** A caller who got a `Choice`'s words gets none, with a warning. It needs a line under Changed in `CHANGELOG.md`.
- **`phase0/` depends on docx4j's re-baseline.** If that's late, Phase A waits or pins one row.
- **The code's name** is the WD's draft, not yet agreed with the engines.

## 7. Questions (decided 2026-11-07, as recommended, Q6 this release)

| # | Question | Decision |
|---|---|---|
| Q1 | Take docx4j's understood set only once published, and keep the traversal list in the package? | **Once published,** since REQ-306 has the published lists first. **Keep the traversal list** as the published copy it is, with nothing reading it, until the WD settles what a Python package must publish. |
| Q2 | One warning per element, or per part? | **Per element**, with its address, as `mce.branches_differ` is. The selector is the part's name either way, and a caller needs to find it. |
| Q3 | Does `mce_preprocess` say anything? | **No.** It's the caller's opt-in to resolving at load. It could log, but no result carries it. |
| Q4 | `detail`'s form: one string of URIs separated by spaces, or a list? | **One string,** as `Requires` is, so `detail` stays a scalar. To be agreed with docx4j and core-ts. |
| Q5 | Which Word files join the fixtures? | **`rw-25`, and the `req305` and `understood-choice` rounds.** Each is small, and they're this rule's only evidence. |
| Q6 | Which release? | **This one.** |

## 8. Phases A and B: implementation notes (2026-11-07)

### 8.1 What was built

- **The two sets:** `resources/understood-namespaces.txt` is docx4j's at `824565363`, 114 URIs, and `traversal-namespaces.txt` beside it is the old 118, which nothing reads. docx4j kept the file's name for the understood set, so no reader's path changed.
- **The rule:** `child.mce_branch` ends at the `mc:Fallback`. `mc_load_mode._splice` takes such an element out, which leaves `kept` with nothing to do, so it went.
- **The warning:** `TextWarning.MCE_NOTHING_SELECTED`, from `Extractor._nothing_selected`, reached from `_compare_branches`, which already stood at every site and takes a word-mode record too. `_compare_branches` now takes the story, for the part's name. At a root, `ReadAsWord.nothing_selected`.
- **`tests/fixtures/nothing-selected/`:** 23 inputs with Word's saves and text saves, 306 KB.

### 8.2 What the build settled beyond the design

- **Only two rows of the suite moved.** docx4j's `expected-inspection.tsv` at `824565363` also holds three rows for a later fixture, so the copy here takes the two `wpi` rows alone, and `test_conformance.py` holds it to both commits.
- **`wpi` was the old tests' example of a namespace not understood.** It's understood now, so `test_mce_levels.py` uses `w98`, as docx4j's `McProbeDocx` does. A prefix the table doesn't know goes on its `mc:Choice`, by docx4j CR-052's rule.
- **`repeated._refused` and `breaks._gives_its_text`** follow the understood set too, as REQ-306's "every markup compatibility question" and REQ-318 have it.
- **`detail` can be empty,** where the `Requires` prefixes bind to nothing.
- **`word` mode takes the element out, where docx4j's `06628f550` keeps it whole in both modes.** The user, 2026-11-07: "keeping ours as built, suggest to docx4j that in word mode, it behaves as we do". Word's save leaves it out, and `preserve` is there for a caller who wants it kept. Whether docx4j follows is its own.

### 8.3 Verified

- **The whole suite:** 12,074 passed, 8 skipped, none failing.
- **`scripts/roundtrip.py`:** clean, 50 of 50 parts, 0 differences, 0 skipped.
- **13 mutants:** every one caught, one only after a test for an unbound prefix.
- **Word's own files:** all 23 read as Word shows them, and the warning is raised on the ten where Word shows neither branch and on none of the rest.

### 8.4 The Word check: passed

Through the Word bridge on 16.0.20326.20232, built over `c991788`, `bridge/staging/cr040-word-check/`, with `out/cr040-word-check/build.py` and `compare.py` (git-ignored). 25 inputs, each run twice: `mc_load_mode="word"` saves of the 23 fixtures, and 2 with a paragraph added through `Body`.

- **Word opened all 50 silently,** with no dialog.
- **The two runs agree** on all 25, in words, sections and properties, leaving rsids aside.
- **All 25 match Word's own reading of the source,** compared from the raw saves. That covers the elements a word-mode save takes out, and the `Choice` content a namespace new to the set promotes, with its prefix in `mc:Ignorable`. The 2 edited files hold the new paragraph once.
