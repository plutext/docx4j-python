# CR-040: An `mc:AlternateContent` that selects nothing, read as Word reads it (WD REQ-305 at 0.98)

**Status:** Proposed 2026-11-07 (the user: "write CR-040"), on the user's ruling that day: "REQ-305 itself changes for all three engines". Nothing is built, and nothing can be until docx4j's list is committed (section 3.1).

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
- **The list:** this package's has 118 URIs, docx4j's at `198d015a3`. docx4j's working copy has 127, uncommitted: nine added and none removed, as many as the namespaces Word was measured to understand beyond the old list (`wpi`, `oel`, `cx3`, `cx4`, `cx6`, `cx7`, `c16r2`, `ax`, `xsi`). `test_the_understood_list_is_docx4js_verbatim` fails today because of that working copy.
- **The list moves one selection.** `phase0/extraction-probe-inspect.docx` holds a `Choice` requiring `wpi` with a `Fallback`. Today it reads the `Fallback`, and with the new list the `Choice`, as Word does. This is a docx4j fixture, so its expected output is docx4j's to re-baseline.

## 3. Design

### 3.1 The list (Phase A)

Copy docx4j's published list once it's committed, to `resources/understood-namespaces.txt` and the suite's copy. Never ahead of it: REQ-306 has additions made to the published list first. `UNDERSTOOD`, the prefix table's own narrower list, doesn't change.

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
- **The list:** `test_the_understood_list_is_docx4js_verbatim` passes again, at 127 or whatever docx4j commits.
- **`phase0/`:** the suite re-pinned at the docx4j commit that holds the list, with `extraction-probe-inspect.docx`'s row as docx4j re-baselines it.

## 5. Phases

| Phase | Work | Cost |
|---|---|---|
| A | The list, the suite's copy, the `phase0/` re-pin | ¼ day, once docx4j commits |
| B | The rule, the warning, `rw-24`, the tests | ½ day |
| C | The records and the Word check of a `word`-mode save | ¼ day |

A goes first, and B never ahead of it. A is right by Word on its own.

## 6. Risks

- **A namespace Word knows and the list lacks** loses words. The warning says so, naming the URI. Eleven were found by asking, and others may exist.
- **Released behaviour changes.** A caller who got a `Choice`'s words gets none, with a warning. It needs a line under Changed in `CHANGELOG.md`.
- **`phase0/` depends on docx4j's re-baseline.** If that's late, Phase A waits or pins one row.
- **The code's name** is the WD's draft, not yet agreed with the engines.

## 7. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Take docx4j's list only once committed, or copy the working copy now? | **Once committed.** REQ-306 has the published list first, and a working copy can still change. |
| Q2 | One warning per element, or per part? | **Per element**, with its address, as `mce.branches_differ` is. The selector is the part's name either way, and a caller needs to find it. |
| Q3 | Does `mce_preprocess` say anything? | **No.** It's the caller's opt-in to resolving at load. It could log, but no result carries it. |
| Q4 | `detail`'s form: one string of URIs separated by spaces, or a list? | **One string,** as `Requires` is, so `detail` stays a scalar. To be agreed with docx4j and core-ts. |
| Q5 | Which Word files join the fixtures? | **`rw-25`, and the `req305` and `understood-choice` rounds.** Each is small, and they're this rule's only evidence. |
| Q6 | Which release? | **The user's.** With CR-039 if docx4j's list is in time. It shouldn't hold anything else. |
