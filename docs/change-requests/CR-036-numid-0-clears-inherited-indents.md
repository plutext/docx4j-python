# CR-036: `w:numId 0` clears every inherited left, hanging and first-line indent, docx4j's rule of 2026-11-05

**Status:** Proposed and decided 2026-11-06, at the user's word ("build CR-036"), on the docx4j-portfolio session's hand-off of docx4j `9e3100351` (`VERSION_17_3_6`, unpushed when this was written). Implemented 2026-11-06 (section 6), unreleased. No Word check: nothing saved changes, and the rule rests on Word's goldens, through docx4j.

**Depends on:** nothing here. The rule is in docx4j's history.

**Counterpart:** docx4j `StyleUtil` at `9e3100351`; docx4j-core-ts CR-027, the same change, whose seam this follows.

## 1. Why

Where a paragraph or a style turns numbering off with `w:numId w:val="0"`, the `PropertyResolver` read the indent wrong. It dropped only what equalled the level's indent, on the comment "a `w:ind` the style states itself survives", which no probe had tried. Word 365 clears every left, hanging and first-line indent inherited so far, wherever stated, and keeps the right one. Only a `w:ind` stated with the `w:numId 0`, or below it, counts.

The evidence is docx4j's: probe `numbering-off-style-ind`, Word's goldens of 5 November 2026, modes 12 and 15 alike (`docx4j-export-fo/docs/word-layout-rules.md`, anchor `s28numid0`). Word wasn't run for this CR.

## 2. Before this CR (measured 2026-11-05)

docx4j's `NumIdZeroIndentTest` cases through `get_effective_ppr`, in twips. The level is left 1440, hanging 1152. `OwnInd` is a numbered style stating left 454, hanging 454, right 57. `Switch` is a style based on it holding `w:numId 0`. `OnBase` is a numbered style based on one stating left 720, hanging 360.

| Case | Shape | Word | Here |
|---|---|---|---|
| N1 | `OwnInd`, direct `w:numId 0` | right 57 | **Wrong:** left 454, hanging 454, right 57 |
| N4 | `OnBase`, direct `w:numId 0` | nothing | Right |
| N7 | `Switch` | right 57 | **Wrong:** left 454, hanging 454, right 57 |
| N8 | `Switch` with a `w:ind` of its own, left 300, hanging 100 | left 300, hanging 100, right 57 | Right |
| N10 | N1 with a direct left 300 | left 300, right 57 | **Wrong:** left 300, hanging 454, right 57 |
| N11 | N1 with a direct hanging 200 | hanging 200, right 57 | **Wrong:** left 454, hanging 200, right 57 |
| N14 | `Switch` with a direct left 142 | left 142, right 57 | **Wrong:** left 142, hanging 454, right 57 |

**What reads it:** `Paragraph.left_indent` and the other indent getters (`effective_p_pr`), the text model's `indent`, and Markdown's style indent. Nothing writes from it: `_set_indent` reads the paragraph's own `w:ind`, and the emulator's read of a style's effective properties takes only its `w:numPr`.

## 3. Design

In `model/style_util.py`, as docx4j's `StyleUtil` and core-ts's `styleUtil.mts`:

- `_drop_numbering_ind` goes, and with it the test of each component against the level's. `_drop_left_and_first_line` clears `left`, `left_chars`, `start`, `start_chars`, `hanging`, `hanging_chars`, `first_line` and `first_line_chars`, and leaves the rest. So a right indent equal to the level's now stays, where the old rule dropped it.
- The branch in `apply_ppr_base` stops asking `get_ind`. It asks whether the inherited `w:numPr` names a list that exists, through `instance_list_definitions`, which `NumberingIndents` gains and the emulator already has. Where it names none the rule doesn't apply.
- A level with no `w:ind` no longer stops the branch. That follows docx4j on a shape its document calls unmeasured.
- Only what is inherited so far clears. A `w:ind` stated with the `w:numId 0`, or below it, still applies, because the merge follows the branch.

**A departure in form, not in rule:** the list's `w:numId` and the `w:numId 0` itself are read as Word reads them (CR-021). `"1.0"` is list 1, as `get_ind` already reads it, and `"0.0"` and `"0.9"` turn numbering off, as the numbering emulator already reads them, where `numbering_off` compared with 0. docx4j's `BigInteger` refuses these forms at the parse.

## 4. Tests and oracles

`tests/content/test_numid_zero_indent.py`, built in memory:

- docx4j's nine cases, the seven above and its two earlier ones.
- A first-line indent; a style asked by name; every left and first-line form clearing and every right form staying; a right indent equal to the level's.
- A `w:numId` naming no list, and no numbering part, leaving the indent alone.
- A level with no `w:ind`, clearing as docx4j's does, marked as unmeasured.
- The CR-021 readings on both sides, and a form Word refuses turning nothing off.
- A second `w:numId 0` clearing nothing more, nothing inherited changing nothing, and a read changing no style.
- The content API in points and the text model's block.

No parity golden holds this: no fixture has the shape. A fixture of it can join the goldens when they're regenerated on a docx4j holding `9e3100351`.

## 5. Not here

- `_drop_level_indent` in `model/revisions/props.py` is a rule of rejecting a tracked change, a different mechanism, and doesn't change.

## 6. As built (2026-11-06)

`model/style_util.py` and the test file, as sections 3 and 4.

- **The 25 tests** are section 4's. Fourteen mutants are each caught: each of the eight names kept, the right and the end cleared, no list test, either `w:numId` not read as Word reads it, and the level's `w:ind` gating the branch. A fifteenth, dropping `copy_leaf`, survives: in these shapes the merge hands the branch a fresh object, so the copy is defensive, kept as docx4j's `deepCopy` is.
- **The corpus:** none of the 8,467 paragraphs in the 453 `.docx` files that load under `samples/` and `tests/fixtures/` has a different resolved indent under this rule. So these files don't exercise it, and section 4's tests are what hold it.
- **An unreadable numbering part:** the same as before. The text model reads as without the part, and the content API raises the part's refusal.
- **The fast suite:** 8,855 passed, 8 skipped (8,830 before).
