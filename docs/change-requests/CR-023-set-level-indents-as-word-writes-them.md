# CR-023: `List.set_level_indents` as Word 365 writes it --- the marker's indent is a first-line indent, negative for hanging; the port of docx4j-core-ts CR-002 section 65.5 (`fd898b9`)

**Status:** Proposed and implemented 2026-10-26, at the user's request ("do The set_level_indents fix"), after core-ts's heads-up that its check 71 (Word 365, the user's run, 2026-10-26) found its own reading wrong, and the same defect was measured here. A defect 0.1.0 and 0.2.0 shipped (CR-003 Phase H, `ef93251`). The Word check is staged; see section 5.

*The number:* CR-022 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057.

**Depends on:** nothing.
- **Counterpart:** docx4j-core-ts CR-002 section 65.5 (`fd898b9`, its `List.mts` and `test/lists.test.mjs`).
- **Word's evidence:** core-ts's check 71 save (`test/check71/71-run-word365.docx`, SHA-1 `99783d4d…`, copied here as `tests/office_js/71-run-word365.docx`), docx4j CR-066 sections 3.2 and 3.4 (round 6), and this repository's CR-022 section 3.3.

## 1. The defect

Office JS's `List.setLevelIndents(level, textIndent, bulletNumberPictureIndent)` takes two indents in points. CR-003 Phase H (section 19, item 6) read the second as the level's hanging indent and wrote it to `w:hanging` as given. Word 365 reads it as the marker's place relative to the text, the level's **first-line** indent:
- written `w:firstLine` when it's zero or more;
- written `w:hanging` with its magnitude when it's negative.

Check 71's case 86, `setLevelIndents(0, -1, 0)`, was saved by Word as `w:left="-20" w:firstLine="0"`. Word keeps first line and hanging in one signed 16-bit field: first line F is stored as F, and hanging H as -H (core-ts check 71 on the API side, round 6 on the file side).

Measured here at `b38a0ed`, before the fix:

| Call | This package wrote | Word writes |
|---|---|---|
| `set_level_indents(0, 36, -18)`, the usual list | `w:left="720" w:hanging="-360"`, a negative in an unsigned attribute | `w:left="720" w:hanging="360"` |
| `set_level_indents(0, 36, 18)` | `w:left="720" w:hanging="360"`, a hanging indent | `w:left="720" w:firstLine="360"`, a first-line indent |
| `set_level_indents(0, -1, 0)` | `w:left="-20" w:hanging="0"` | `w:left="-20" w:firstLine="0"` (case 86) |

`to_api_script` inverted the same mapping: it wrote `set_level_indents(ilvl, left, hanging)`, and it reported a level's first-line indent as beyond the three level setters.

## 2. The fix (D1, D2)

- **D1, `List.set_level_indents`:** `text_indent` is `w:left`, as before. The second indent is rounded to twips, and written as `w:firstLine` when it's zero or more, or as `w:hanging` of its magnitude when it's negative; the level holds one or the other. The parameter keeps its name, `bullet_number_picker_indentation`, so a caller passing it by keyword is unaffected.
- **D2, `to_api_script`:** a level's indents become `set_level_indents(ilvl, left, first_line)`, or `-hanging` where the level has a hanging indent. Running the script writes the same `w:ind`. A first-line indent is no longer listed as beyond the setters; `w:right` still is. A level holding both attributes, which Word can't store, keeps `w:firstLine` in the script, with a note.

**Not in this CR: a range check.** core-ts's fix also holds the values to Word's ranges (left -32,768 to 32,767; first line 0 to 32,767; hanging 0 to 32,768). Here that's CR-022's guard, decided 2026-10-26, which covers every setter: this one by Office JS's own limit for a level's text indent (±1,638 pt), and the marker by section 3.3's single field. Its census lists this setter.

## 3. Tests

- `tests/content/test_lists.py`:
  - **Updated:** `test_set_level_indents_is_in_points` was pinning the old mapping, and now calls `(0, 36, -18)` and expects `hanging` 360.
  - **New:** `test_set_level_indents_marker_is_a_first_line_or_hanging_indent` covers five markers, -18, 18, 0, -0.4 and 0.4.
  - **New:** `test_set_level_indents_writes_what_word_365_writes` holds case 86's `w:ind` to Word's save, attribute for attribute.
- `tests/content/test_api_script.py`: `test_set_level_indents_round_trips_a_first_line_or_a_hanging_indent` checks that the script names the marker, and that running it writes the same `w:ind`, for a first-line and a hanging indent.

## 4. Records

CHANGELOG [Unreleased], Fixed. CR-003 section 19, item 6, notes that this CR supersedes its mapping. `tests/README.md`, the index and CLAUDE.md are updated.

## 5. The Word check

`out/cr023-word-check/` is staged at `vbShares/Office 2016/__tmp/python-cr023/`: three lists, each set by one of section 1's calls. The question is whether each opens without a repair and shows the indent its label names. Its result goes in section 6.

## 6. Implementation notes (2026-10-26)

To be filled in once the Word check is in: the commit, the gates' numbers, and the Word check's result.
