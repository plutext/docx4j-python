# CR-024: The content API's spacing, line, list indent and font size setters as Word 365's own --- a line spacing written under the auto rule, and Word's own limits; the port of docx4j-core-ts CR-002 section 65.9 (`b7c2680`)

**Status:** Proposed and implemented 2026-10-28, at the user's request, after core-ts's check 71b (CR-022 section 11.2): "point 2: agreed" (Word's own limits) and "point 1: switch to "auto" matching Word and core-ts". **The Word check is pending** (section 5).

*The number:* CR-023 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057.

**Depends on:** nothing.
- **Counterpart:** docx4j-core-ts CR-002 sections 65.8 and 65.9 (`b7c2680`, 2026-10-27, unreleased): its `Paragraph.mts` writes `lineRule="auto"`, and its `wordRange.mts` rows take Word's own limits.
- **Word's evidence:** core-ts's check 71b, run by the user in Word 365 16.0.20326 on 2026-10-27: its run JSON (`test/fixtures/check71b/71b-setter-limits-word365-run.json`, SHA-1 `24d606ed…`, copied here unedited as `tests/office_js/71b-setter-limits-word365-run.json`) and Word's save (`test/check71b/71b-run-word365.docx`).
- **The record across ports:** `../docx4j-portfolio/docs/office_js_departures.md`, written by the portfolio session on the user's word the same day. Office JS's reference pages state no range, rounding or line rule for any of these members, so each finding is undocumented rather than contradicted.

## 1. What check 71b found

**Word's own `lineSpacing` setter writes an auto rule.** It writes `w:line` as 20 x the points under `w:lineRule="auto"`: 0.75 pt is `w:line="15" w:lineRule="auto"`, and 12 pt is 240, single spacing. That is Office JS's documented reading ("In the Word UI, this value is divided by 12"): the points are a multiple of 12-point lines, and Word lays the line out as a multiple of the font's single spacing. `Paragraph.line_spacing` here wrote `w:lineRule="exact"` (CR-003 Phase B), so a line spacing set here laid out as a fixed height where Word's is a multiple. Check 71 couldn't show it: Word refused every line value that check set.

**Word's own setters stop below what Word stores**, with `InvalidArgument` and nothing written. CR-022 had taken these limits from check 71, which set only 1,584 and 1,638 pt and found 1,638 refused:

| Setter | Word takes | The content API took (CR-022) |
|---|---|---|
| `spaceBefore`, `spaceAfter` | 0 to 1,584 pt, 31,680 twips | 0 up to, not including, 1,638 pt |
| `lineSpacing` | 0.75 to 1,584 pt, 15 to 31,680 under auto | 0.05 up to, not including, 1,638 pt, exact |
| `setLevelIndents`' text indent | -1,583.95 to 1,583.95 pt, 31,679 twips either way | above -1,638.4, below 1,638 pt |
| `font.size` | 1 to 1,638 pt after rounding to a half-point, 2 to 3,276 | above 0, up to 1,638 pt |

Word takes `font.size = 0.9` and writes it as `w:sz="2"`, so its limit is on the value as written, not on the value given. A refused `setLevelIndents` leaves the list `startNewList` made with Word's default level, 36 pt left and 18 hanging.

## 2. The change (D1 to D3)

- **D1, `Paragraph.line_spacing`:** writes `w:line` as 20 x the points under `w:lineRule="auto"`, replacing an exact or at-least rule already there. The getter is unchanged: it already read an auto line as 240ths of a 12-point line, so a value set reads back as set. The python-docx facade's `ParagraphFormat.line_spacing` follows python-docx and is not changed.
- **D2, the limits:** `space_before` and `space_after` take 0 to 1,584 pt, `line_spacing` 0.75 to 1,584, `List.set_level_indents`' text indent -1,583.95 to 1,583.95, and `Font.size` 1 to 1,638 after rounding. Each is a `CallerLimit` with the new `stored` flag, which bounds the value as the setter writes it, in the attribute's unit, as core-ts's rows do (`OFFICE_JS_SPACING`, the new `OFFICE_JS_LINE_SPACING`, `OFFICE_JS_TEXT_INDENT`, `OFFICE_JS_FONT_SIZE`). The refusal's message is CR-022's: `space before 5000 pt (100,000 twips) is beyond Word's range, 0 to 1,584 pt, where Office JS's setter stops`.
- **D3, `to_api_script`:** an auto line rule, stated or the default, is expressed as `line_spacing`, and an exact or at-least rule falls back to `insert_xml`, as core-ts's `toApiScript` does. Where the source leaves the rule to its default, the block is marked inexact with a note, because running the script states the rule.

## 3. Departures, and what is kept

1. **`space_before = 0` writes `w:before="0"`**, where Word writes nothing when the style already gives 0 (check 71b case 1, core-ts's one listed difference). The file reads the same.
2. **`Font.size = 0` still removes the size**, where Word refuses 0 (CR-022 section 10.13, decision 1).
3. **The limits bound the value as written, replacing CR-022 section 10.8 item 5**, which bounded the value given. Word's font setter decides it: 0.9 pt is taken. For the other three, every value check 71b set is on the 0.05 pt grid, where both readings agree. Between grid points the value as written decides, as in core-ts: 1,584.02 pt of space before is written as 31,680 twips and accepted, and a text indent of 1,583.99 pt, written as 31,680, is refused.
4. **The facade and the builders keep the faithful range** (CR-022 Q4, Q5). Only the content API speaks for Office JS.

## 4. Tests

- **New, `tests/content/test_office_js_setters.py`:** each of check 71b's 51 cases through the setter of the same name. The setter refuses exactly where Word refused, 31 cases, changing nothing; elsewhere it reads back Word's value and writes Word's attributes, `w:spacing`'s and `w:sz`'s value for value, and the level's `w:ind`. Case 1 is the one difference, pinned.
- **Updated, `tests/content/test_measurement_setters.py`:** the ranges of the four setters and their messages; `test_a_line_spacing_is_written_auto_as_words_setter_writes_it` replaces the exact-rule test.
- **Updated:** `tests/content/test_paragraph.py` (the rule written), `tests/content/test_api_script.py` (an auto rule expressed, stated and default; exact and at-least fall back), `tests/content/test_office_js_characters.py` (the fixture in `SHA1SUMS`, 101 files).

## 5. The Word check

`out/cr024-word-check/` is staged at `vbShares/Office 2016/__tmp/python-cr024/`. It holds paragraphs whose line spacing is set by `line_spacing`, from 0.75 pt to 36, and one in a large font, so the multiple shows. The question is whether the file opens without a repair and Word's Paragraph dialog names each line spacing as its label says. Its result goes in section 6.

## 6. Implementation notes (2026-10-28)

- **What changed:** `values.CallerLimit` gained `stored`, and `_PER_POINT` the auto line's unit, 240ths of a line at 20
  to the point; the four limits in `paragraph.py`, `lists.py` and `font.py`; `Paragraph.line_spacing`'s write;
  `api_script.py`'s spacing branch and its inexact note. Nothing generated, marshalled or loaded changed.
- **Against Word's run, by hand as well as by the test:** 1,584 pt of space before is written as 31,680 twips and 1,584.05
  refused; a line of 0.75 pt is `w:line="15" w:lineRule="auto"` and 0.7 refused; a font of 0.9 pt is `w:sz="2"`, 0.5
  and 1,638.4 refused, 1,638.05 written as 3,276; a text indent of ±1,583.95 pt taken and ±1,584 refused. The
  messages read `line spacing 0.7 pt (14 240ths of a line) is beyond Word's range, 0.75 to 1,584 pt, where Office
  JS's setter stops`.
- **No other port's user is affected here:** `../docx4j-mcp-py` sets none of these members.

**The gates.**

| gate | result |
|---|---|
| the fast suite (`-m "not slow"`) | **6,438 passed, 7 skipped**, 0 failed: 55 more than CR-020 Phase A's 6,383, the 53 of `test_office_js_setters.py` and `test_api_script.py`'s 2 |
| `tests/content/test_api_script.py` with its slow tests (every sample's script executed) | passes |
| acceptance (`compare_acceptance.py`, `out/acceptance/` from CR-020 Phase A against `out/acceptance-cr024/` from this change) | **17 artefacts, 246 entries, 0 differ** beyond the dates: no acceptance document sets these members through the content API |
| `ruff check` on the changed files | passes |

Not run, since nothing they cover changed: the corpus round trip, `generate.sh --check`, `scripts/threads.py`.
