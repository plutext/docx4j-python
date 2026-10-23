# CR-018: Run properties read as Word 365 reads them --- on/off values, the enumerations and `w:color`, one reader for each, and `effective()` for a value the file leaves out

**Status:** Proposed 2026-10-24, at the user's request through the portfolio session ("yes please", to "Shall I ask docx4j-25, the TypeScript session and the Python session each to draft its CR now, giving them these rules and the probes?"). **Nothing has started.** docx4j (CR-060) and docx4j-core-ts (its CR-015) drafted theirs the same day against the same rules; docx4j's was accepted that morning and is being implemented (section 3.1). Word's rules come from four rounds of probes the user ran in Word 365 on 2026-10-24 (section 1.1). Five things are proposed: **Revised the same day** (the user: "assume we will do a fork release. review the rest of the CR on the basis that anything better done in the fork can be done there"): the parse's exactness for `bool` and enumerations, a record of the values it keeps, and the typed schema default move into the fork's stage 13, released as 26.2.3 (section 4.1a, Q2, Q3, Q16), as the fork session sketched them. The reading rules stay in `docx4j_py`. Round five's answers are folded in (section 3.1b), and Q14 to Q16 are new.
- **(a)** one ST_OnOff reader, used by every reader of an on/off value here, with the value kept as written so that Word's rule can be applied to it;
- **(b)** the enumerations `u`, `highlight`, `vertAlign` and `em` read as unset when the value is outside the enumeration, in the resolver's merge as well as in the readers;
- **(c)** one `w:color` reader: hex, the classic HTML colour names, and automatic;
- **(d)** `effective(obj, "field")`, the one explicit reader of a value the file leaves out, over the schema defaults the model already carries;
- **(e)** the `wml.xsd` refresh from docx4j `46c727bf8`, which corrects those defaults.

Four phases are proposed (section 5), in the order A, C, B, D: A, on/off and the refresh; B, the enumerations and colour, on CR-016 Phase B; C, `effective()`; D, round five, the Word check and the records. They come to 3 to 4½ days. Thirteen questions wait on the user (section 8). Q1 recommends that Phase A go into 0.2.0 and the rest after it.

*The number:* CR-016 is the latest CR here. Its header reserves CR-017 for a port of docx4j CR-057 (`insert_markdown` reading local image files), which is not drafted, so this one is CR-018.

**Depends on:**
- **The probes.** `~/vbShares/Office 2016/onoff-fun/`: 20 probes and Word 365's save of each (`<name>-word365.docx`), the PDFs and round 1's screenshots, and `README.txt` (SHA-1 `2af2a57b91392e3205780315398c7262a810f71c`). The docx4j-portfolio session built them (its `onoff/build.py` to `build4.py`), and the user ran them on 2026-10-24. Appendix A lists the 40 files' SHA-1s.
- **docx4j:**
  - **CR-060**, the counterpart. Proposed at `38ccaedeb` and accepted at `5d7564778` (`VERSION_17_3_5`, unpushed), D1 to D6 as recommended. Its Phase 0 is in progress.
  - CR-045 (2026-10-16, `"on"` and `"off"` read as booleans), CR-060's predecessor.
  - #697 (`77ba784d6`) and CR-059 (`ac97b56c5`, `46c727bf8`) for the refresh, and `82c64bf14` for `w:anchorLock` (section 4.9).
- **docx4j-core-ts:**
  - its counterpart, `docs/change-requests/CR-015-run-properties-as-word-reads-them.md`, which was **uncommitted** in its working tree when this was drafted. It is cited by name and by its registry entry only. Nothing here is read from a working tree.
  - `05fcf39`, where jsonix 3.4.1 reads `"on"` and `"off"` as booleans, and its fixture `test/fixtures/onoff/pandoc-code.docx`.
- **The xsdata fork's stage 13, released as 26.2.3** (section 4.1a; revised 2026-10-24, the user: "assume we will do a fork release"). Three items, all off by default: `ParserConfig(exact_types=...)` for `bool` and `Enum` (item 29), `SkippedReport.kept` (item 30) and `XmlVar.schema_default_value` (item 31), sketched by the fork session on 2026-10-24 (its probes in `/tmp/claude-1001/-home-jharrop-git-docx4j-xsdata/43b7a91a-08cd-4dec-ae85-2406e681cfe1/scratchpad/onoff/`). Phase A needs the release on PyPI before CI, which installs the fork from PyPI, can run it; `pyproject.toml`'s floor becomes `>=26.2.3,<26.3`.
- **This repository:**
  - CR-016 Phase B (`Font` reading every run in scope), which Phase B's readers sit under (section 4.9 item 1).
  - CR-002 Phase B's catalogue, `StyleUtil` and `PropertyResolver`, whose merge and toggle rules Phase A and B change.
  - CR-011's text model and its ratchet; CR-010 and CR-013's tracked formatting; CR-003 Phase J's facade.
  - CR-009 Phase B and CR-008 (both proposed), which are touched but not needed.
- **The sibling checkouts** `../docx4j`, for the refresh (`git diff 2853b7280 46c727bf8 -- xsd/wml/wml.xsd`), and `../docx4j-core-ts`, for the parity goldens.

**Counterparts:**
- docx4j CR-060 at `5d7564778`: its section 1 (the rules), section 2 (docx4j measured), section 3 (an `OnOffAdapter`; `Underline`, `HighlightColor` and `FontColor`; one colour normaliser), and D1 to D6.
- core-ts CR-015, uncommitted. Its registry entry (`core-ts/CR-015`) names its D1 to D5, an Office JS check 60, a fifth probe round of invalid values over styles, and phase 0 copying the probes into `test/onoff-fun/`.

## 1. Why

### 1.1 Word's rules (Word 365, four rounds, 2026-10-24)

The rules as the portfolio session read the user's runs. Each was checked here against Word's own saves (section 2.6), and all 90 colour markers agree with rule 3 and the table.
1. **On/off (ST_OnOff).** Measured on `b`, `i`, `caps`, `smallCaps`, `strike`, `dstrike`, `outline`, `shadow`, `emboss`, `imprint` and `vanish`; by extension `bCs`, `iCs` and the rest.
   - A value is **off for exactly `false`, `off` or `0`**: case-sensitive and untrimmed.
   - It is **on for anything else**: bare, `true`, `on`, `1`, `""`, `" "`, `TRUE`, `FALSE`, `OFF`, `False`, `" false"`, `"false "`, `00`, `no`, `yes`.
   - Word shows no repair prompt.
   - Word saves an on value bare. An off value is dropped, or kept as `w:val="0"` where it overrides a style (props-b-styles S2, S3, S4, S6).
   - The toggle (ParaBold and CharBold, S7) is not bold.
2. **Enumerations** (`u`, `highlight`, `vertAlign`, `em`). A value outside the enumeration is unset, as if absent, and the case must be exact.
   - A bare `<w:u/>`, or one with only `w:color`, is no direct underline: over CharUnderline the style's underline shows (S11).
   - `none` over a style removes the underline (S10).
   - An invalid `w:u` opens with no repair prompt. A missing or invalid value on `highlight`, `vertAlign` or `em` makes Word repair the file, dropping the element.
3. **`w:color`.**
   - Six hex digits, any case, with an optional leading `#` and surrounding whitespace trimmed.
   - Or a name from the classic HTML (Netscape) colour table, case-insensitive. That table spells the grays with an a and has `lightgrey`, not `lightgray`. It has no `grey`, no CSS 4 names, no X11-only names, no system colours and no spaces.
   - Anything else is automatic: `F00`, `0xFF0000`, five or seven digits, `GG0000` and `""`. `auto`, in any case, is automatic too.
   - Word saves a name, and lower-case hex, as upper-case hex.

**The table, from the saves.** Word wrote a hex value for 41 distinct names (48 markers, the case variants among them), and each value is the CSS 3 keyword's. It dropped 18 names. Seven of the 18 are exactly the seven keywords CSS 3 has and the classic table lacks: `grey`, `darkgrey`, `dimgrey`, `slategrey`, `lightslategrey`, `darkslategrey` and `lightgray`. Round 4 probed all seven, so the fit with the 140-name table is not a guess about the grays. The other 11 dropped are `darkYellow`; `rebeccapurple`, `transparent` and `currentcolor`; the X11-only `lightgoldenrod`, `violetred` and `navyblue`; `dark red`; and the system colours `window`, `buttonface` and `windowtext`. That leaves 99 of the table's 140 names unmeasured (Q7).

### 1.2 What it costs today

Section 2 measures it. Five findings change what a document means or does:
1. **A save that changes meaning.** `w:b w:val=" false"` and `"false "` are bold in Word. Here the bool converter trims them, binds them `False`, reads them off, and **saves them as `w:val="0"`**, so they are not bold in Word after a save (props-b-case K5, K6). This is the one save over the 20 probes that changes meaning (section 2.7).
2. **`"off"` reads as on, in every reader**, for all eleven measured properties, and over a style too: CharBold with a direct `"off"` reads bold (S3).
3. **`w:vanish w:val="off"` loses its text.** The run reads as hidden, so it is missing from `Paragraph.text`, `Body.text`, `search`, the text model and Markdown, where Word shows it (T10).
4. **A write that does nothing.** `range.font.bold = True` over a range across paragraphs skips a run whose `w:b` is `"off"`, because the run "reads bold already" (`format_span`'s reads-shortcut, `span_edits._reads`). The run's markup is left `<w:b w:val="off"/>`, and Word still shows it not bold.
5. **Colours and underlines read wrongly.**
   - A colour Word draws, such as `red` or `DarkRed`, comes back from `Font.color` as the name, and as `None` from the facade.
   - An invalid `w:u` reads `"Mixed"` where Word draws no underline.
   - `highlight_color` returns `"YELLOW"` or `"bogus"` raw.

**The corpus says these values are rare in what Word and docx4j write** (section 2.10). 1 document of 422 holds an on/off spelling other than bare, `true`, `false`, `1` and `0`, and none holds an invalid enumeration or a named colour. **No reader's answer over the corpus would move.** Other producers write them, though. docx4j's #585 is Apache POI writing `w:val="off"` (docx4j CR-060 section 3.1); pandoc's reference document carries `"on"` and `"off"` (core-ts `05fcf39`); and the first edition of ECMA-376 allowed both. So the fix is for other producers' files, and finding 1 is for anyone's.

## 2. What exists here (measured 2026-10-24 at `7545995`)

Everything below is read only. It was measured on the working tree at `7545995`, which holds the uncommitted port of docx4j `82c64bf14` (`w:anchorLock` read as stated, `w:framePr` only; section 4.9 item 6) and nothing else. The scripts are in this session's scratchpad (`cr018/scripts/`):
- `probe_readers.py`: every reader below, over each probe and over Word's save of it;
- `compare.py`: where each reader's answer on a probe differs from its answer on Word's save;
- `pydocx_probe.py`: python-docx 1.2.0 over the probes;
- `wordsaves.py` and `check_table.py`: the saves against the rules;
- `save_meaning.py`: the canonical comparison of section 2.7;
- `census.py` and `padded.py`: section 2.10;
- `defaults.py`: section 2.9;
- `write_over_off.py`, `span_over_off.py`, `vanish_off.py` and `checkbox.py`: the findings.

Nothing under `docx4j_py/`, `tests/`, `schemas/` or `codegen/` was changed.

**The oracle.** Word's save of a probe holds only Word's own spellings: bare for on, a value dropped where Word reads it as absent, upper-case hex. Read by this package's readers, which read those spellings correctly, it gives Word's answer. A reader **agrees with Word** on a marker when its answer on the probe equals its answer on Word's save. Every probe paragraph is `label: MARKER (end)`, and only the marker run carries the property tested. There are 20 probes with 227 marker runs.

**Loading.** Every probe loads. Nothing raises and nothing is skipped: `package.skipped` is empty on all 20, and on Word's 20 saves. A value the converter can't parse gives a `ConverterWarning` and is bound as its source string (CR-002 section 12.3 item 7), and a save writes it back as read.

### 2.1 How a value is bound

| type | converter | what it binds |
|---|---|---|
| ST_OnOff, an element's `w:val` (`BooleanDefaultTrue.val`) and the attributes ECMA-376 types ST_OnOff, which docx4j's copy types `xsd:boolean` | the fork's `BoolConverter`, upstream's | `True` for `true`/`1` and `False` for `false`/`0`, **after `value.strip()`** (`converter.py:296`). Anything else (`on`, `off`, `""`, `TRUE`, `yes`) fails, warns and stays the source `str` |
| `w:u/@w:val`, `w:highlight/@w:val`, `w:vertAlign/@w:val`, `w:em/@w:val` | the `EnumConverter`, exact | the member, or the source `str` for anything else (`"SINGLE"`, `"YELLOW"`, `"Superscript"`, `"DOT"`, `""`) |
| `w:color/@w:val`, `None \| STHexColorAuto \| bytes` | `auto`, else `xsd:hexBinary`, which trims | `bytes` for any even number of hex digits, with stage 12's lexical spelling kept (`ff0000` and `" FF0000"` are saved as read); `b""` for `""`; the `str` for `auto` in another case, a name, `#FF0000`, an odd number of digits, `GG0000` |

### 2.2 On/off (rounds 1 and 2)

`b` in every value. `i`, `caps`, `smallCaps`, `strike`, `dstrike`, `outline`, `shadow`, `emboss`, `imprint` and `vanish` bare, `"off"`, `"0"`, `""` and `"FALSE"`. The readers are:
- the resolver, `is_true` over the effective `w:rPr`, all eleven;
- `Font`, `b`, `i`, `strike` and `dstrike` only, since this package's `Font` (Office JS's subset) reads no others;
- the text model's runs, `b`, `i`, `strike` and `vanish`, and Markdown;
- the facade, all eleven.

| `w:val` | Word | bound here | resolver, `Font`, text model, Markdown | facade | python-docx 1.2.0 | saved here | Word's save |
|---|---|---|---|---|---|---|---|
| absent (`<w:b/>`) | on | `None` | on | on | `True` | `<w:b/>` | `<w:b/>` |
| `true`, `1` | on | `True` | on | on | `True` | `"1"` | bare |
| `false`, `0` | off | `False` | off | off | `False` | `"0"` | dropped; `"0"` over a style |
| `on` | on | `"on"`, with a warning | on | on | `True` | as read | bare |
| `off` | **off** | `"off"`, with a warning | **on** | **on** | `False` | as read | dropped; `"0"` over a style |
| `""` | on | `""` | on | **off** (`bool("")`) | raises `InvalidXmlError` | as read | bare |
| `" "`, `TRUE`, `FALSE`, `OFF`, `False`, `Off`, `00`, `no`, `yes` | on | the `str` | on | on | raises `InvalidXmlError` | as read | bare |
| `" false"`, `"false "` | **on** | **`False`**, trimmed | **off** | **off** | raises `InvalidXmlError` | **`"0"`**: off | bare |

**Over styles** (props-b-styles), every reader is Word's in S1 to S11 but one:
- **S3**, CharBold with a direct `"off"`, reads bold in every reader. Word draws it not bold, and its save writes `w:val="0"`.
- The toggle S7 (ParaBold and CharBold) reads not bold, as Word draws it, and S8 adds a direct `<w:b/>` and reads bold, as Word does.
- `StyleUtil.toggle` (`style_util.py:275`) reads a style level's `"off"` through `is_true` as a stated true, so a character style stating `"off"` over a bold paragraph style would invert it to not bold. Word reads it as a stated false, which leaves the paragraph style's bold standing. No probe has this case; it follows from rule 1 and the toggle rule CR-002 section 15 settled.

**Vanish** (T10). `w:vanish w:val="off"` reads hidden, so the run's text is missing. `Paragraph.text` reads `'T10 w:vanish w:val="off":  (end)'`, `search("vanish off")` finds nothing, and the text model and Markdown leave it out. In Word the text shows. The cause is `hidden_runs_of`'s `_on` (`text_model.py:827`) and the extractor's `_on`.

**A write over `"off"`.** `Range.font.bold = True` on one run writes `<w:b/><w:bCs/>` over the `"off"`, which is right. Over a range across paragraphs (V4 to V6 of onoff-valid), `format_span` leaves `<w:b w:val="off"/>` as it was, because `_reads` asks `read_run_options`, which reads `"off"` as bold already. In Word V5 stays not bold.

### 2.3 The enumerations (round 2)

| value | Word | bound here | `Font` | facade | python-docx 1.2.0 | saved here |
|---|---|---|---|---|---|---|
| `<w:u/>`, bare; `<w:u w:color="FF0000"/>` | no direct underline; over CharUnderline its underline shows (S11) | `val` `None` | `"None"`; S11 `"Single"`, through the merge, which keeps the style's `val` | `None` | `None` | as read |
| `w:u` `""`, `SINGLE`, `Single`, `NONE`, `bogus`, `true`, `1`, `on`, `off`, `0`, `false` (11 markers) | unset, no repair | the `str` | **`"Mixed"`** | `None` | raises `ValueError` | as read |
| `w:highlight` `YELLOW`, `bogus` | unset; Word repairs the file | the `str` | **`"YELLOW"`, `"bogus"`** | `None` | raises `ValueError` | as read |
| `w:vertAlign` `Superscript`, `bogus` | unset; Word repairs | the `str` | not superscript | not superscript | not superscript | as read |
| `w:em` `DOT` | unset; Word repairs | the `str` | (nothing reads `em`) | (python-docx has no `em`) | --- | as read |
| bare `highlight`, `vertAlign`, `em` | unset; Word repairs | `val` `None` | unset | unset | raises `InvalidXmlError` (`highlight`, `vertAlign`) | as read |

**Over a style it would differ.** The resolver's merge takes any stated `w:val` over the style's. `merge_u` uses `_attr(source.val, ...)`, and `highlight`, `vertAlign` and `em` are `_val` leaves that `is_empty_val` calls stated. So a direct `w:u w:val="bogus"` over CharUnderline reads `"Mixed"`, and a direct `w:vertAlign w:val="bogus"` over a superscript style reads not superscript. By rule 2, Word reads the style's value in both. No probe has an invalid value over a style that states the same property; S11 has a missing one. Round five measures this (Q8).

### 2.4 `w:color` (rounds 2 to 4)

| value | Word | bound here | `Font.color` | facade `color.rgb` | python-docx 1.2.0 | saved here |
|---|---|---|---|---|---|---|
| six hex digits, any case (`FF0000`, `ff0000`, `Ff0000`) | that colour | `bytes`, the spelling kept | `#FF0000` | `FF0000` | `FF0000` | as read |
| `"FF0000 "`, `" FF0000"` | `FF0000` | `bytes`, since `hexBinary` trims | `#FF0000` | `FF0000` | `FF0000`; **`" FF0000"` is `0FF000`** | as read |
| `#FF0000` | `FF0000` | `str` | `"#FF0000"` (right by accident: the string as it is) | `FF0000` | raises `ValueError` | as read |
| a name in the table, any case (41 names, 48 markers) | its hex | `str` | **the name** (`"red"`, `"DarkRed"`) | **`None`** | raises `ValueError` | as read |
| `auto` | automatic | `str` | `"auto"`; CR-016 proposes `"#000000"` | `None` | `None` (type `AUTO`) | as read |
| `Auto`, `AUTO` | automatic | `str` | **`"Auto"`, `"AUTO"`** | `None` | raises `ValueError` | as read |
| a name outside the table (18); `F00`, `0xFF0000`, `GG0000` | automatic | `str` | **the string** | `None` | raises `ValueError` | as read |
| five or seven digits (`FF000`, `FF00000`) | automatic | `str` | **the string** | **`FF0000`**: the facade copies python-docx's slicing parse | `FF0000` | as read |
| `""`; `<w:color/>` | automatic | `b""`; `val` `None` | `""` | `None` | raises | as read |

### 2.5 The readers' answers against Word's saves

Of the 227 marker reads, the markers where each reader's answer on the probe differs from its answer on Word's save:

| reader | differ | on/off | underline | highlight | colour |
|---|---:|---:|---:|---:|---:|
| the resolver's toggles (`is_true`) | 14 | 14 | --- | --- | --- |
| `Font` | 95 | 8 | 11 | 2 | 74 |
| the text model's runs | 6, and the T10 run missing | 6 | --- | --- | --- |
| Markdown | 7 | 7 | --- | --- | --- |
| the facade | 75 | 25 | 0 | 0 | 50 |

Notes on the counts:
- `Font`'s 74 colour differences include `auto` read `"auto"` (K1), which CR-016 proposes to change. `Font` has no reader of `caps` and the like.
- The facade's 25 on/off differences are 12 `"off"` read on, 2 `" false"` read off, and 11 `""` read off.
- **python-docx 1.2.0**, for section 4.8, raises on 121 of the 227 marker reads (`InvalidXmlError` or `ValueError`). Where it answers, it differs from Word in only 3: five and seven digits read `FF0000`, and `" FF0000"` reads `0FF000`. It reads `"off"` off and `"on"` on (`ST_OnOff.convert_from_xml` accepts exactly `1`, `0`, `true`, `false`, `on` and `off`).

### 2.6 Word's saves against the rules

`check_table.py` takes each of the 90 colour markers (rounds 2 to 4) and computes rule 3 with the 140-name table (CSS 3's 147 keywords less its six `grey` spellings and `lightgray`). It agrees with Word's save in all 90. Rounds 1 and 2's on/off markers agree with rule 1 in every save: bare where Word read on, dropped or `"0"` where it read off. The enumeration markers agree with rule 2: each invalid or missing value is gone from Word's save.

### 2.7 What a save writes, and what the canon tolerates

`scripts/canon.py` tolerates one category, **`boolean-spelling`**: `{"1", "true"}` against each other, and `{"0", "false"}` (CR-001 section 9). Any other attribute difference is `attribute-value`, and element text is `text`. Neither is tolerated.

Over the 20 probes, each main part read (so re-marshalled) and saved, then compared with its source:

| converter | `boolean-spelling` | other differences |
|---|---:|---|
| today's, which trims | 4 (`true` to `1`, `false` to `0`) | **2 `attribute-value`**: `' false' -> '0'`, `'false ' -> '0'` (props-b-case), the save that changes meaning |
| a non-trimming one registered in the process (section 4.1) | 4 | **none** |

Every other odd spelling is saved as read today, because a value the converter rejects stays a `str`: `"off"`, `"on"`, `"TRUE"`, `""`, `"red"`, `"SINGLE"` and `"YELLOW"`. Lower-case hex keeps its spelling through stage 12's lexical `bytes`.

### 2.8 The readers of an on/off value here

`grep` for `.val`, `is not False` and `bool(` over the hand-written code finds these:

| # | where | what it reads | rule today | Word's? |
|---|---|---|---|---|
| 1 | `wml/builders.py` `_on` (186), through `read_run_options` | `Font`'s `bold`, `italic`, `strike_through`, `double_strike_through`; `span_edits._reads` (`format_span`'s shortcut); `api_script._run_format` | `val is not False` | `"off"` on; `" false"` off |
| 2 | `model/property_catalogue.py` `is_true` (108), through `style_util.toggle` (275) | the toggle XOR at every style level | the same | the same |
| 3 | `model/property_resolver.py` `is_vanished_by_styles` (553) | a style's `w:vanish` | the same | the same |
| 4 | `model/content/text_model.py` `_on` (827), through `hidden_runs_of` | hidden runs for `Paragraph.text`, `Body.text`, `search`, the offsets | the same | the same |
| 5 | `model/text/extractor.py` `_on` (167), `_on_of` (2472) | `b`, `i`, `strike`, `vanish` of runs and marks; `showingPlcHdr`; `updateFields`, `evenAndOddHeaders`, `titlePg`, `contextualSpacing`, `tblHeader` | the same | the same |
| 6 | `model/fonts/run_font_selector.py` `is_on` (94) | `cs`, `rtl`, `b`, `i` for the font slot | the same | the same |
| 7 | `model/math/omml_to_latex.py` `_on` (481) | a math run's `w:b`, `w:i` | the same | the same |
| 8 | `model/content/controls.py` `is_showing_placeholder` (729) | `w:showingPlcHdr` | the same | the same |
| 9 | `model/content/table.py` `TableRow.is_header` (837) | `w:tblHeader` | the same | the same |
| 10 | `model/content/tracking.py` `_track_revisions_in` (1581, the tree) | `w:trackRevisions` | `getattr(value, "val", True) is not False` | the same |
| 11 | the same, the lxml branch (1590); `model/content/describe.py` `_tracking_on` (387) | `w:trackRevisions` | `not in ("0", "false")` | **`"off"` on** |
| 12 | `model/content/controls.py` `remove_when_edited` (832), `is_multi_line` (941), `plain_text_lines_of` (1314) | `w:temporary`, `w:text/@w:multiLine` | `str(value).lower() not in ("0", "false")` | **`"FALSE"` off, `"off"` on** |
| 13 | `model/customxml/kinds.py` `allow_insert_delete_section` (596) | `w15:doNotAllowInsertDeleteSection` | `str(value).lower() in ("0", "false")` | the same |
| 14 | `model/customxml/kinds.py` `is_checked` (144) | `w14:checked` | `str(value).lower() in ("1", "true", "on")` | **a bare `<w14:checked/>` reads unchecked**, though its schema default is true (`checkbox.py`); `"TRUE"` reads checked |
| 15 | `model/text/extractor.py` (719) | `w:documentProtection/@w:enforcement` | `in (True, "1", "true", "on")` | `"TRUE"` and `""` off |
| 16 | `model/text/extractor.py` `_written` (2533) | `w:fldLock`, "as written, false when absent" | `in (True, 1, "1", "true", "on")` | the same |
| 17 | `model/listnumbering/definitions.py` `_on` (81), over lxml | numbering's on/off children | `not in ("0", "false", "off")` | **Word's already** |
| 18 | `model/content/math_text.py` `_on` (220) | `m:` on/off, an enumeration with `on`/`off` | `not in (False, "off", "0", 0, "false")` | Word's already, for OMML |
| 19 | the facade: `docx/text.py` `_get_bool` (403, `Font`'s twenty) and `_get_toggle` (845, `ParagraphFormat`'s four), `docx/section.py` (253), `docx/settings.py` (56), `docx/styles.py` (153), `docx/table.py` (370) | as named | `True if value is None else bool(value)` | **`""` off**, `"off"` on, `" false"` off |
| 20 | `tests/revisions/description.py` (462), a test helper | `w:val` of a Word-save description | `in (None, True, 1, "1", "true", "on")` | `"TRUE"` off |

**Not ST_OnOff.** `model/chart/model.py` `_bool` (51) reads DrawingML's `xsd:boolean`, where XSD's rule (whitespace collapsed) is the right one. Today it relies on the converter's trim, so a non-trimming converter would move it unless it trims itself (section 4.1). `w:tblLook`'s six attributes are `SharedTypesStOnOff`, an enumeration with `on` and `off`, and no reader reads them yet (CR-008 will).

### 2.9 The schema defaults

- **In the model.** Every field is `None | T = None`. The XSD default is kept only as `metadata={"schema_default": ...}` (CR-001's `SchemaDefaults=metadata`), which the fork exposes as `XmlVar.schema_default`. It is never filled on parse, never written on save, and no hand-written code reads it.
- **The numbers.** 546 generated fields declare one, every one an attribute, and 51 of them are in `wml`. Counted with inherited fields, as `dataclasses.fields` sees them, there are 666, 62 in `wml`.
- **The hazards** for a reader over them: 116 enum defaults are already members; 49 `str` fields default to `""`; three are `"NaN"`; and a stated value can be an unconverted `str`.
- **The copy is old.** `schemas/` is from docx4j `2853b7280`, so it still has the 2008 on/off defaults that docx4j's #697 (`77ba784d6`) and CR-059 (`ac97b56c5`, then `46c727bf8`) corrected: 30 to false, 10 with no default, `equalWidth` true.
- **The refresh, dry-run on a scratch copy.** `git -C ../docx4j diff 2853b7280 46c727bf8 -- xsd/wml/wml.xsd` applies to a copy of `schemas/wml/wml.xsd` with `patch`: 29 hunks, clean, at offsets of up to 66 lines. The 17 `docx4j-python` markers are intact. `xsd/` changed in no other file in that range, and `wml.xsd` hasn't changed in docx4j since `46c727bf8`. Regenerating changes only `wml` metadata: 36 values true to false, and 10 dropped (measured by this repository's session before this draft; not regenerated here).

### 2.10 The corpus census

`census.py` reads the 422 distinct documents of `out/cr012-b-corpus/` (`COMMITS.txt`: this repository `f1bb1d9`, core-ts `df5b4af`, docx4j `901a922a3`) with lxml. It covers every WordprocessingML part and the 206 element names `wml.xsd` declares `BooleanDefaultTrue`.

| what | count |
|---|---|
| on/off elements | 32,328: 32,052 bare, 274 `true`/`false`/`1`/`0`, **1 `"on"`, 1 `"off"`**; no whitespace, case, empty or other spelling |
| documents holding any other spelling | **1**: core-ts's `test/fixtures/onoff/pandoc-code.docx` (pandoc's reference document, `Microsoft Word 12.0.0` in its `app.xml`), with `w:tblHeader w:val="on"` and `w:wordWrap w:val="off"` in its Source Code style |
| answers that would move | **none**: no reader here reads `w:wordWrap`, and `"on"` reads on today too |
| `w:u` | 412 valid; none bare without a colour; none invalid |
| `w:vertAlign`, `w:highlight` | 148 and 7, all valid; none bare |
| `w:color` in a `w:rPr` | 8,630 upper-case hex, **20 lower-case hex** (the same pandoc document's syntax colours, which read the same), 17 `auto`; **no name, no `#`, no padding, nothing else** |
| a boolean spelling padded with whitespace, in **any** part of any document (attribute or element text) | **0**, over 441 distinct documents: the corpus, `tests/` and `samples/` |
| this repository's `tests/` and `samples/` at `HEAD` (179 distinct documents) | no odd on/off, enumeration or colour value at all |

So the parity goldens, CR-011's ratchet, the round trip over `samples/` and the acceptance artefacts hold none of these values. Each is a gate that must not move (section 6.2).

### 2.11 Found on the way, each taken into a phase or a question

1. **`w:vanish w:val="off"` hides text Word shows** (section 2.2). Phase A.
2. **A write across paragraphs leaves `"off"`** (section 2.2). Phase A, through `_reads`.
3. **A bare `<w14:checked/>` reads unchecked.** Its schema default is true. Word always writes `w14:val`, so it is unmeasured. Phase C, through `effective()` (section 4.4).
4. **Two lxml readers of `w:trackRevisions` read `"off"` as on** (section 2.8 items 11 and 12). Phase A.
5. **The facade's colour reads five and seven digits as `FF0000`**, copying python-docx's slicing parse, where Word reads automatic. Q9.
6. **The chart reader leans on the converter's trim** (section 2.8). Phase A.
7. **`em` has no reader.** Only the merge sees it. Phase B keeps the merge right and adds no reader.

## 3. What the counterparts propose

### 3.1 docx4j CR-060 (accepted 2026-10-24, `5d7564778`)

docx4j binds ST_OnOff as `xsd:boolean` (since 2008, `6cbe2e057`). JAXB's boolean parse turns `"off"` into `null`, which `isVal()` reads as true. It trims `" false"` to false, and a save writes `"true"` or `"false"`. So docx4j has this package's two on/off defects, and a third this package doesn't have: **a docx4j save writes `"off"` as on**.

Its decisions, as accepted:
- **D1**: an `OnOffAdapter` (`XmlAdapter<String, Boolean>`) bound in `wml.xsd` to every ST_OnOff attribute. It reads exactly `false`, `off` or `0` as `FALSE` and anything else as `TRUE`, and leaves the fields' types alone. The preprocessor route the user suggested (`39641e21c`'s throwing adapter and a template) was weighed and not taken.
- **D2**: a save writes `true`/`false`.
- **D3**: invalid enumerations read as absent. `highlight`, `vertAlign` and `em` with a missing or invalid value are **dropped on save**, as Word repairs them, and `w:u` is kept.
- **D4**: a colour's source is kept, and rendered normalised.
- **D5**: the full 140-name table.
- **D6**: underline styles and emphasis marks in rendering go to a CR of their own.

### 3.1a docx4j CR-060 as implemented (2026-10-24, `87e6ace56`, `893658f68`, `868a1797b`)

docx4j's session handed it off the same day. The commits are on `VERSION_17_3_5`, unpushed, and checked in `../docx4j`. docx4j is this package's behavioural oracle, so its built answers are the ones to port against. Where this CR departs from them, section 4.10 says why.

- **The fixtures and the oracle.** `docx4j-core-tests/src/test/resources/runprops/word365/` holds the 20 probes, Word's 20 saves and a README (41 files at `868a1797b`). `RunPropertiesWord365Test` takes Word's save as the oracle: each marker's reading in the probe must equal its reading in Word's save, through the resolver, the text model, HTML's direct styling, PDF's effective styling and a round trip. The readings are compared as what is drawn: the font size and language from Word's docDefaults are left out, off counts as absent, and colours are upper-case hex. This package's measurements (section 2.5) use the same method.
  - The fixtures are in docx4j's Apache-2.0 tree, as `samples/` are. Q11 can therefore copy them from docx4j at `868a1797b`, byte for byte with their SHA-1s, rather than from the share. The user's confirmation is then needed only for the round-five probes, if any.
- **On/off:** off for exactly `false`, `off` or `0`, else on, applied to 46 of docx4j's 48 on/off attributes. The two left out, `checkStyle` (`w:activeWritingStyle`) and `actualPg` (`w:readModeInkLockDown`), are required, so XJC makes them primitive booleans that an adapter can't bind; only Word reads either. Here `AllOptional` makes every attribute `None | T`, so no exclusion is needed. docx4j writes `true`/`false` (its D2; this CR's Q5 keeps the source spelling).
- **Colour:**
  - Trim, then six hex digits in any case with an optional `#`, written upper-case. Else one of the 140 classic HTML names, case-insensitive (docx4j-core `org/docx4j/model/color/ColorTables.java` `HTML`, the table this CR's 140-name list must equal). Else automatic.
  - It applies to `w:color` and `w:u/@w:color` only. `w:shd` and border colours are untouched, since names there are unmeasured.
  - An invalid direct `w:color` over a style's colour keeps the direct value, automatic. That is unmeasured, and the same as this CR's Q8.
- **Highlight:** case-exact `ST_HighlightColor` (16 values and `none`). Anything else is absent, in the resolver too, so a style's highlight beneath shows. This is the same as this CR's (b).
- **`w:u` with no value** (bare, colour only, or invalid): no underline is drawn.
- **On save (docx4j's D3):**
  - A `highlight`, `vertAlign` or `em` with a missing or invalid value is left out, on `w:rPr` and on `w:pPr`'s `w:rPr` only.
  - A revision's former properties (`w:rPrChange/w:rPr`, `CT_ParaRPrOriginal`) are written as they are.
  - `w:u` is written bare, as before. An invalid `w:u` over a style is unmeasured (core-ts's D5).
  - This CR's Q6 keeps them on save.
- **docx4j's inspector** (`InspectXml.on`) ignored case; it now uses the same rule.

### 3.1b Round five, run 2026-10-24 (docx4j CR-060 section 12)

The portfolio built the probes and the user ran them in Word 365 the same day: `props-u-styles`, `colour-styles`, `props-enum-styles` and `colour-elsewhere` in `~/vbShares/Office 2016/onoff-fun/`, with the README's round-5 section. docx4j's session reported the results; its CR-060 section 12 has each marker.
- **An invalid value over a style that states the same property reads as absent, so the style's value shows.** This holds for an invalid or bare `w:u`, an invalid `w:color`, and an invalid `highlight`, `vertAlign` or `em`. Word's save drops each.
  - **The exception:** `w:u w:val="true"` over a style removes the style's underline, and Word saves it as `"none"`. The other on/off words over a style are unmeasured.
  - Keeping the value on save (Q5, Q6) therefore keeps the meaning, apart from the repair prompt that `highlight`, `vertAlign` and `em` bring.
  - Word drew no highlight from a character style in the repaired file, which is an effect of the repair.
- **Colour names are read in `w:u/@w:color`, `w:shd/@w:fill` and border colours too.** An invalid fill drops the shading.

This changes two things in this CR:
- **Q8's colour half turns.** An invalid colour is as if absent, not a stated automatic: the merge skips it as it skips an invalid enumeration, so the style's colour shows. docx4j's reading of it (the direct value kept) was unmeasured when it was built, and its session has the measurement.
- **The colour reader's scope widens** to `w:shd/@w:fill` and the border colours (`w:pBdr`, `w:tcBorders`, `w:tblBorders`, `w:bdr` and their sides), each with the same rule. An invalid fill reads as no shading. docx4j's reader covers only `w:color` and `w:u/@w:color` (section 3.1a).

**The portfolio's spot-check adds four details** (the w:u and enumeration saves agree):
- Over an underlining style, a colour-only `w:u`, or `bogus` plus a colour, gives the style's underline in that colour.
- An explicit `"auto"` colour does override a style's colour. `auto` or `AUTO` over a style is otherwise unmeasured.
- Over a style, F00 in `w:u/@w:color` and in a border colour gives automatic, and in `w:shd/@w:fill` drops the shading.
- Word drew no highlight from a character style, though a paragraph style's highlight showed. The file had been repaired, so this waits for a clean probe before it counts as Word's rule.

The single-value `w:u w:val="true"` exception goes into Phase B as a pinned case. Whether the other on/off words over a style behave the same is a sixth-round question, if one is wanted.

### 3.2 docx4j-core-ts CR-015 (proposed, uncommitted)

Known here only from its registry entry. jsonix fails the whole part on an on/off value outside `true`/`false`/`1`/`0`/`on`/`off`, or on an enumeration value outside its enumeration. Its D1 puts the leniency in the parse (jsonix and the generated objects), so it depends on a jsonix change. Its D2 keeps each invalid value's spelling until round five is measured, then drops those it shows to be as if absent. It plans an Office JS check 60 and a fifth probe round of invalid values over styles, and copies the probes into `test/onoff-fun/`.

### 3.3 The xsdata fork session's answer (2026-10-24)

*Its first answer, below, recommended a registered converter with no fork change. With the user taking a fork release as given, its later sketch is section 4.1a's, which supersedes items 3, 5 and 6 below.*

1. **The trim** is upstream's `BoolConverter.deserialize` (`value.strip()`). That is right by XSD, since `xs:boolean`'s whiteSpace is collapse, but Word doesn't follow XSD here. The parser passes attribute values and element text unmodified, so element text trims too (`<ep:ScaleCrop> false</ep:ScaleCrop>` binds false and is saved `0`).
2. **No lexical bool is possible.** `bool` can't be subclassed, an `int` subclass breaks `is False`, and stage 12's `lexical_types` refuses `bool`. The only lossless form is the source `str`, which the conversion-failure path already binds.
3. **The fix needs no fork change.** `converter.register_converter(bool, WordBool())`, upstream's public API, where `WordBool` matches exactly `true`/`1` and `false`/`0` and raises for anything else, so the value is bound as its `str` and saved as read.
   - The fork's per-var converter plans carry the factory's generation, so a registration after they are built is picked up.
   - Measured: all 15 spellings keep their meaning on save, the four canonical ones still bind bools, and the cost is nil (18.2 µs either way).
4. **`"on"` and `"off"` stay a reader's.** Accepting them in the converter would save `"off"` as `0`, which breaks CR-012's bar of canonical identity.
5. **The scope.** `converter` is one module-level `ConverterFactory` (`converter.py:861`) that 11 fork modules import directly. There is no per-parser or per-context converter and no injection point, so a registration reaches every `docx4j_xsdata` parse and serialise in the process.
   - **Upstream `xsdata` is unaffected**: the rename makes the two converters separate objects. Probed in a venv holding both, the fork rejects `" false"` after the registration and `xsdata` still returns `False`.
   - `codegen/generate.sh` runs the generator in a separate process.
   - Upstream's own documentation presents global registration, a built-in type's override included, as the way to do this.
6. **For tests.** `unregister_converter(bool)` doesn't restore the original: `type_converter(bool)` then falls back by MRO to `IntConverter`. A test that swaps the converter restores it with `register_converter(bool, ...)`.

This session confirmed the code read only. `runtime.py` uses the fork's global through `XmlParser` and `XmlSerializer`, and never holds a converter of its own. `docx4j_py`'s own import doesn't import `runtime` (`load` and `parser()` do), and section 2.7's second row is its measurement.

## 4. Design here

### 4.1 (a) The one ST_OnOff reader, and keeping the value as written

**The reader.** A new hand-written module, `docx4j_py/model/values.py`. `model/` is in `codegen/clean.py`'s `KEEP` whole, so it needs no `KEEP` entry, and `wml/builders.py` already imports from `model/`.

```python
OFF = frozenset({"false", "off", "0"})          # rule 1: exactly these, case-sensitive, untrimmed

def on_off(value, default=True) -> bool:
    """An ST_OnOff value as Word 365 reads it: None is `default` (the schema's
    true for an element's w:val); a bool is itself; an enumeration member
    (w:tblLook's, OMML's) reads by its value; any other value is off only if
    it is exactly "false", "off" or "0"."""

def is_on(element) -> bool:
    """A w:b-like element: absent, False; present, on_off(element.val)."""

def xsd_boolean(value, default) -> bool:
    """xsd:boolean as XSD reads it, for the namespaces that aren't Word's ST_OnOff
    (DrawingML): whitespace collapsed, then the four literals."""
```

**Every reader of section 2.8 calls it.** Items 1 to 16 and 19 go through `is_on` or `on_off`, the lxml readers (11, 17) through `on_off` of the attribute string, and item 20, the test helper, too. Items 17 and 18 are Word's already and are switched only so that the rule has one home. Items 15 and 16 are attributes, and read through `on_off(effective(...))` once Phase C has `effective()` (section 4.4); until then, `on_off(value, default=False)`. `model/chart/model.py`'s `_bool` takes `xsd_boolean`.

**What this fixes.**
- `"off"` reads off everywhere, so `w:vanish w:val="off"` shows its text, and `format_span` writes `w:b` over an `"off"` run.
- `""` reads on in the facade.
- A style level's `"off"` is a stated false in `StyleUtil.toggle`.
- `w:trackRevisions w:val="off"` reads off in `describe()` and in tracking.
- `"FALSE"` reads on in `w:temporary` and `w:multiLine`.
- `" false"` reads on, once the converter keeps it (below).

**Keeping `" false"` as written: the options.**

| option | what | fork change | `" false"` kept and read on | cost |
|---|---|---|---|---|
| 1. A non-trimming bool converter, registered by `docx4j_py` (the first draft's recommendation; superseded) | `runtime.py` registers `ExactBoolConverter` (the fork session's `WordBool`) on import: the four literals exactly, anything else kept as its `str` with a `ConverterWarning`, as `"off"` is today | none, and no release | yes | process-wide over `docx4j_xsdata`'s `bool` fields (Q3) |
| **2. A fork `ParserConfig` option, off by default** (recommended, section 4.1a) | about ten lines and a test in the fork (`tests/fork/`), recorded in `docs/fork/CHANGES.md`; `runtime.parser_config()` turns it on | yes: **a 26.2.3 release** before CI (which installs the fork from PyPI) or a docx4j-python release can use it | yes | per parser, so narrower in reach; the same fields |
| 3. A lexical `bool`, as stage 12's floats and bytes | impossible: `bool` can't be subclassed, and an `int` subclass breaks `is False` | --- | --- | --- |
| 4. A schema patch typing the ST_OnOff attributes `xsd:string` | 66 `wml` fields become `str`, and `BooleanDefaultTrue.val` with them | none; a regeneration | yes | the model's types and every writer change; numeric booleans lost; not recommended by the fork session either |
| 5. The converter accepts `on`/`off` | binds them as bools | --- | no: `"off"` would be saved `0`, which breaks canonical identity | --- |
| 6. Recover the source spelling at read time (lxml beside the tree) | --- | no | read on, but **still saved `0`** | a second parse, and the save still changes meaning |

**Revised 2026-10-24: option 2 is recommended (Q2), as section 4.1a has it**, since the user takes a fork release as given. Option 1 was the first draft's, because it needed nothing from the fork; its measurement stands for both, since the two bind the same values (section 2.7: the probes' saves canonically identical to their sources). Option 1's notes, kept for the record:
- **The registration** is at `docx4j_py.runtime`'s import, before any parser exists, and is idempotent. Every engine parse goes through `runtime.parser()`. A caller who builds an `XmlParser` on `docx4j_xsdata` without importing `runtime` gets upstream's trim, and the docstring says so. `scripts/roundtrip.py`, `scripts/parents.py` and the tests that build a `ParserConfig` of their own (`test_skipped_report.py`, `test_performance.py`, `test_child.py`) import `runtime`, so that the harness measures what the engine does.
- **Other `xsd:boolean` fields** (DrawingML, docProps' element text, every namespace) keep a padded or re-cased value as a `str` instead of trimming it. That is lossless, and saved as read. Their one reader here, the chart's `_bool`, reads it by `xsd_boolean`, with the same answer as today. There are 0 such values in 441 documents (section 2.10).
- **A value set in code** is a `bool` and is written numeric, as now.

### 4.1a The fork's stage 13 (26.2.3), and why each item is the fork's

Revised 2026-10-24, the user: "assume we will do a fork release. review the rest of the CR on the basis that anything better done in the fork can be done there." The fork session sketched each item (design only). docx4j-python's rule for the split: **the fork keeps what was written, exactly and explicitly; `docx4j_py` decides what it means.**

1. **Item 29, `ParserConfig(exact_types: frozenset[type] = frozenset())`**, accepting `bool` and `Enum` (`ValueError` for anything else), beside `skipped_report` and `lexical_types`.
   - **What it does.** For a type in the set, the converter matches its lexical forms exactly, with no XSD whitespace collapse. `bool` takes exactly `true`, `false`, `1` and `0`. An enumeration's member matches only if its value equals the text, case and spacing and all. Anything else binds the raw text by the failure path, and is saved as read.
   - **Why the fork:** the trim happens there (`BoolConverter.deserialize`'s `strip()`, converter.py:296; `EnumConverter.deserialize` strips and re-joins tokens, converter.py:642). An explicit per-parse option replaces option 1's global registration: no mutation at import, no import-order trap for the scripts, and no global to swap and restore in tests.
   - **How.** `ParserUtils.parse_var`, through which every XML conversion passes, passes the set to the converter. Stage 6's per-var plan isn't keyed by it: stage 12's `plan[3]` widens to a kind code (none, LEXICAL, BOOL, ENUM), so a value pays only the check it pays today.
   - **Scope: every field of the two types, in every namespace** (311 bool vars, wml's 70 among them). The mode is lossless everywhere: an exact spelling binds as before, and anything else is kept. A narrower scope would need a schema patch and three generator pieces, so the fork session advises against it, and so does this CR.
   - **Enumerations, measured by the fork session:** today `w:highlight " yellow"`, `"yellow "` and `"\tyellow\n"` bind `YELLOW` and are saved `"yellow"`, as do `w:u " single"` and `w:vertAlign "superscript "`. Under the option they are kept and saved as read. That is lossless whatever Word's whitespace rule turns out to be. Word's rule is unmeasured, and only decides what the reader does with a kept `" yellow"`. Until a probe settles it, this CR reads such a value as **outside the enumeration, unset** (Word is case-exact, and untrimmed for on/off), a pinned case (Q15).
   - **Unions:** only the member type in the set becomes exact. A `bool|int` given `" 1"` still gives `int` 1, pinned in the fork's tests.
   - **ints** normalise too (`" 12"`, `"007"`) and stay outside the set, unmeasured (section 4.12).
2. **Item 30, `SkippedReport.kept`**: when the skipped report is on, a list of `KeptValue(qname, parent_class, path, text, type, reason)`.
   - `reason` is `"exact"` for a value kept on purpose (it would have converted under XSD's collapse: `" false"`, `" yellow"`), or `"invalid"` for a genuine failure (`"off"` for `xs:boolean`, `"Yellow"`, `"bogus"`).
   - **No `ConverterWarning` for `"exact"`**, since it's what the caller asked for, and it isn't a failure under `fail_on_converter_warnings`. `"invalid"` keeps its warning and its strictness.
   - It is separate from the report's items, so `len(report)`, iteration and `roundtrip.py`'s "nothing skipped" bar are unchanged: a kept value isn't skipped content.
   - It is per config, so it keeps this package's one-config-per-thread rule. Collecting warnings instead needs `catch_warnings`, which is process-global.
   - **Why the fork:** only the parser can tell the two kinds apart. `docx4j_py` surfaces them as `part.kept` and `package.kept` beside `part.skipped` (CR-002 section 12.3 item 7's "value" entries), and Q16 asks whether `describe()` lists them.
3. **Item 31, `XmlVar.schema_default_value`**: the schema default typed once when the var is built, cached on the meta. `schema_default` stays lexical, so its documented contract stands.
   - **Rules:** an enumeration default is the member; no default is None; a lexical default that doesn't convert keeps its string (only hand-written metadata can do that, since the generator checks its defaults).
   - **Why the fork:** it is generic, the analogue of Jsonix's typed `defaultValue`. `effective()` then reads one slot instead of keeping its own per-class cache (section 4.4).
   - **The policy stays here:** which namespaces are audited, and `LookupError` elsewhere (Q10).

**Size:** about 100 lines and 25 fork tests, all off by default, one commit per item, recorded in the fork's `docs/fork/CHANGES.md` as stage 13. The fork session builds it on the user's go. **The gate before the release:** this package's suite with `exact_types={bool, Enum}`, `roundtrip.py`, and `bench_marshal` plus `Symbols.docx` with the option on and off, interleaved, since `w:b`, `w:u` and `w:highlight` are hot.

**Then, here:**
- `runtime.parser_config()` sets `exact_types={bool, Enum}` and `skipped_report=True`, as it does the rest.
- **A test asserts that every `ParserConfig` in the tree comes from `parser_config()`**, so that a script that builds its own can't read `" false"` as `False` (the fork session's warning).
- The floor goes to `>=26.2.3,<26.3`, and `.venv-fork`'s editable install is refreshed.

**What stays here, and why:**
- **The reading rules**: on/off, the enumerations, colour, the facade and the style merge. They are Word's reading of WordprocessingML, not a binding. The fork serves every namespace, and only keeps what was written.
- **What a save writes** (Q5) needs nothing from the fork. A bool can't keep its spelling (it can't be subclassed), so `true`/`false` still become `1`/`0`, the tolerated `boolean-spelling` difference.

### 4.2 (b) The enumerations: unset when outside, in the merge as well as the readers

```python
def in_enumeration(value, enumeration):
    """A member of `enumeration`, or None: a member is itself; a str is the member
    whose value it is exactly (case and all); anything else, None (rule 2)."""
```

Two places use it, because an invalid direct value must not override a style's value (section 2.3):
1. **The resolver's merge** (`property_catalogue.py`). `merge_u` takes `w:val` only when it is in `UnderlineEnumeration`, so a bare `w:u`, a colour-only one and an invalid one all leave the style's `val`, as S11 measured. `w:u`'s `w:color` and theme attributes merge on their own, as now. The three `_val` leaves `highlight`, `vertAlign` and `em` get an emptiness test that also counts a value outside `HighlightVal`, `STVerticalAlignRun` or `STEm` as empty, so the style's comes through. The other `_val` leaves keep `is_empty_val`. This is what docx4j has for free: JAXB gives `null` for such a value.
2. **The readers.**
   - `read_run_options`: an invalid `w:u` reads `"None"` (the direct view) where it read `"Mixed"`; an invalid `highlight` reads `None`; `vertAlign` already reads right.
   - The facade's `underline`, `highlight_color`, `superscript` and `subscript` read through it. They already answer `None` and not-superscript here.
   - Nothing reads `em`, and the merge is all it needs.

The kept `str` is never rewritten: a save writes what was read (section 4.7, Q6). Under section 4.1a's `exact_types`, a padded member (`" yellow"`) is kept too, and reads as outside the enumeration until Word's whitespace is measured (Q15).

### 4.3 (c) The colour reader

```python
def colour(value) -> str | None:
    """w:color/@w:val as Word 365 reads it: 'RRGGBB' upper case, or None for automatic.
    bytes of length 3 are the colour; a str is trimmed, loses one leading '#', and is
    the colour if six hex digits; else a name in NETSCAPE_COLOURS, case-insensitive;
    anything else ('auto' in any case, F00, 0xFF0000, five or seven digits, '') is None."""
```

**The table.** `NETSCAPE_COLOURS`, the 140 names with their hex, written into `values.py` and tested against CSS 3's 147 less the seven (Q7).

**Its readers:**
- **The scope, widened by round five:** `w:color/@w:val`, `w:u/@w:color`, `w:shd/@w:fill` and `w:shd/@w:color`, and every border's `w:color` (`w:pBdr`, `w:tblBorders`, `w:tcBorders`, `w:bdr`, `w:pgBorders`), each by the same rule. An invalid `w:fill` reads as no shading.
- `read_run_options`'s `color` (so `Font.color`): `#RRGGBB` for a colour. Every automatic spelling reads as `auto` reads: `"auto"` today, and `"#000000"` if CR-016's proposal lands.
- The facade's `ColorFormat.rgb`: an `RGBColor`, or `None`. Its `type` is `AUTO` for every automatic spelling and `RGB` for a colour (Q9).

**Order.** A stated `w:themeColor` is CR-009 Phase B's and wins over `w:val`, as Word reads it; `colour()` reads `w:val` only.

**Over a style** (Q8, answered by round five, section 3.1b). An invalid value is as if absent, so the style's colour shows: the merge skips a value `colour()` reads as `None` unless it is a stated `auto`, which does override. The merge therefore gains a colour emptiness test beside the enumerations' (section 4.2).

### 4.4 (d) `effective(obj, "field")`, the reader of a value the file leaves out

```python
def effective(obj, name):
    """The value `obj` states for attribute `name`, else the schema's default, typed;
    else None. Never the field's own reading: the field stays None when absent."""
```

- **Read from the fork's `XmlVar.schema_default_value`** (section 4.1a item 3), typed once when the meta is built, through `runtime.context().build(type(obj)).get_attribute_vars()`, a name-to-var dict built once per class. (The first draft typed each default here with `converter.deserialize`; that moves into the fork.)
- **The stated value wins**, an unconverted `str` included (`" false"`, `"red"`). The lexical readers above it decide what that means.
- **Attribute vars only.** All 546 defaults are attributes.
- **The hazards, answered as the schema has them:** `""` defaults are `""`, `"NaN"` is `float("nan")`, and both are listed in the tests.
- **Scope (Q10).** The `w` namespace, whose defaults docx4j audited (#697, CR-059) and Phase A refreshes, and `w14` and `w15` once Phase C has checked their few defaults (their on/off type is docx4j's own patch, default true). Anywhere else it raises `LookupError` naming the namespace, until someone audits that namespace's defaults. That follows the docx4j session's warning: XJC's getters returning the default hid absence and caused #697, #641, the autospacing bug, the CR-059 `anchorLock` regression and the FontEmbedder bug. So this is an explicit helper, and the model's fields never change. It also follows Jsonix's advice that one generic reader over the metadata keeps the audited XSD the single source of truth.
- **How it relates to (a) to (c).** `effective()` answers **absence**; (a) to (c) answer **a present value's spelling**.
  - An on/off element is `on_off(effective(element, "val"))`. `BooleanDefaultTrue.val`'s default is `"true"`, so that is the same answer as `is_on`, with the rule in the schema rather than in code.
  - An attribute ECMA types ST_OnOff, such as `w:fldLock`, `w:dirty`, `w:enforcement` or `w:anchorLock`, is `on_off(effective(obj, "fld_lock"))`, and the refresh makes its absence false. That is how section 2.8's items 15 and 16 read.
  - `w14:checked`'s default is `"true"`, so a bare `<w14:checked/>` reads checked (section 2.11 item 3).
  - The enumerations and `w:color` have no default to read: their `w:val` is required, and a missing one is unset (rule 2) or automatic (rule 3).

### 4.5 (e) The `wml.xsd` refresh from docx4j `46c727bf8`

1. `git -C ../docx4j diff 2853b7280 46c727bf8 -- xsd/wml/wml.xsd | patch schemas/wml/wml.xsd`: 29 hunks, clean, the patches' markers intact (section 2.9). That range touches no other file under `xsd/`, so `PATCHES.md` can name `46c727bf8` as the copy's commit.
2. **`schemas/PATCHES.md`** gains the refresh's section, in section 9's form: the commits (#697 `77ba784d6`, CR-059 `ac97b56c5` and `46c727bf8`), the hunks, the checks, and that no patch changed.
3. **`codegen/generate.sh`, then `--check`.** Only `wml` metadata moves: 36 defaults true to false, and 10 gone.
4. **A test** compares each `w` schema default in the model with docx4j's `wml.xsd` at `46c727bf8`, parsed directly, so that the next refresh is checked mechanically.

Nothing reads the defaults until Phase C, so the refresh changes no answer.

### 4.6 Where each piece lives

| piece | where | why |
|---|---|---|
| keeping `" false"` (and `" yellow"`) as written | the fork, `ParserConfig(exact_types={bool, Enum})` (stage 13 item 29), set by `runtime.parser_config()` | the trim happens in the converter; an explicit per-parse option, not a global registration (Q2, section 4.1a) |
| recording what was kept | the fork, `SkippedReport.kept` (item 30); `part.kept` and `package.kept` here | only the parser can tell kept-on-purpose from invalid (Q16) |
| a default typed | the fork, `XmlVar.schema_default_value` (item 31) | generic, and typed once with the meta; the audit policy stays here (Q10) |
| `on_off`, `is_on`, `xsd_boolean`, `in_enumeration`, `colour`, the table, `effective` | `docx4j_py/model/values.py`, new, hand-written | the rules are Word's reading, not a binding: the parse keeps the value, and a reader applies the rule, which is the fork session's advice (keep the reader in `docx4j_py`; the fork keeps what was written, section 4.1a); one module that every layer can import, the builders included |
| an invalid enumeration not overriding a style | `model/property_catalogue.py` (`merge_u`, `is_empty_val`) | the merge is where a stated value wins |
| each reader of section 2.8 | where it is, switched to the module | one rule, and one test of each reader against Word's saves |
| the defaults | `schemas/wml/wml.xsd`, by the refresh | the schema copy is where a change to the model belongs (CLAUDE.md) |
| not in the fork | --- | `on`/`off` in the converter would change saves (section 3.3 item 4); the rules are Word's reading of WordprocessingML, and the fork serves every namespace: the fork keeps what was written, `docx4j_py` decides what it means |

### 4.7 What a save writes

The portfolio's rule is that **a save never changes meaning**. This package's own rule is stricter: an unmarshalled part saved unchanged is canonically identical to its source, `boolean-spelling` the one tolerated category (CLAUDE.md, CR-001 section 9).
- **Recommended (Q5): a save keeps the source's spelling.** It already does for everything the converter rejects. With section 4.1's converter it does for `" false"` too, and `true`/`false` keep becoming `1`/`0`, the tolerated category. So `"off"`, `"on"`, `"TRUE"`, `""`, `"red"`, `"ff0000"`, `"SINGLE"` and `"YELLOW"` are written as read, and Word reads each as it read the source. Section 2.7 measured it over the probes.
- **The alternative: Word's form.** That is bare for on; `0` for an off value over a style, or dropped; upper-case hex; and repaired enumerations dropped. It would change bytes Word doesn't need changed, and `canon.py` would have to learn that `"off"` equals `"0"` for ST_OnOff attributes only, a rule it can't see from the XML.
- If wanted, a later explicit operation (`normalise()`, say) can write Word's form, and nothing in this CR stops it.
- **A value written in code** keeps today's form: `el.b()` is bare, and `val=False` is `"0"`. **An edit that copies properties** keeps the `str`: a run split, a recorded `w:rPrChange` (`record_r_pr_change`), `deep_copy`.
- **Invalid `highlight`, `vertAlign` or `em` (Q6).** Word repairs the source, dropping them; docx4j's accepted D3 drops them on save. The recommendation here is to keep them. "Nothing may be dropped silently" is this package's rule, a save keeps exactly the repair prompt the source already had, and the meaning is the same either way, since both Word and this package read them as absent. core-ts's D2 keeps them too until round five. An edit that writes the property replaces them.

### 4.8 The facade

python-docx 1.2.0's own reading is the facade's promise (CR-003 Phase J). Measured (section 2.5), python-docx:
- accepts exactly `1`, `0`, `true`, `false`, `on` and `off` for ST_OnOff, case-sensitive and untrimmed, and reads them as Word does;
- **raises `InvalidXmlError` for every other on/off value**, `""`, `" false"` and `"TRUE"` among them;
- raises `ValueError` for an invalid `w:u`, an invalid `w:highlight`, any colour name, `Auto`, `#FF0000`, `F00` and `""`;
- answers wrongly, by slicing, for five or seven hex digits (`FF0000`) and for `" FF0000"` (`0FF000`).

Over the probes it raises on 121 of 227 marker reads, and where it answers it differs from Word in those 3 colours only.

So **following Word contradicts python-docx's answer only where python-docx misparses a colour.** Everywhere else it either agrees with python-docx or answers where python-docx raises. CR-003 section 20.3 items 9 and 10 already answer where python-docx refuses.
- **Recommended (Q9): Word's reading throughout**, with python-docx's tri-state shape kept: `None` for not stated, and for an enumeration Word reads as unset.
- The facade then reads `"off"` off (python-docx's and Word's answer; it reads on today) and `""` on (it reads off today), and colour names as their hex.
- It is recorded as a departure in CR-003 section 20.3.
- The alternative, raising where python-docx raises, would make the facade fail on files Word opens without a prompt.

### 4.9 Interactions

1. **CR-016's `Font` reading every run** (core-ts section 59). Phase B's readers sit under it, in `read_run_options` and the merge, so B lands on CR-016 B's `Font`. With both, a paragraph of `"off"` runs and absent ones reads `bold False`, not `None`, since they agree. Whatever CR-016 decides for `auto` (it proposes `"#000000"`) becomes the answer for every automatic spelling.
2. **CR-013 and CR-010's tracked formatting.** A tracked `Font` write records the run's properties as they stand in `w:rPrChange` (`record_r_pr_change`, a copy), so a kept `"off"` is recorded as written. Reject All restores it whole, and the text model's original view reads it through the same readers. `format_span`'s `_reads` decides what a tracked write across paragraphs touches, so a run with `"off"` now gains its `w:b` and its `w:rPrChange`. Nothing in CR-010's listing reads a value.
3. **CR-011's text model, its ratchet and its goldens.** The extractor's readers switch (section 2.8 item 5). The ratchet runs docx4j's text-extraction suite over docx4j's fixtures, which hold no odd value (section 2.10), so `conformance.tsv` mustn't move. If docx4j's CR-060 adds on/off checks to `ExtractionConformance`, they come with the next port of it.
4. **The parity goldens.** core-ts's goldens come from the docx4j release, so they move when docx4j's readers do. Their fixtures hold no odd value: `fonts-cs-off.docx`'s `w:cs` is `false` or bare. So docx4j CR-060 shouldn't move the 46 goldens here, and a golden that docx4j adds over the probes comes with the next regeneration.
5. **CR-009 Phase B** (theme colours): a stated `w:themeColor` wins over `colour(w:val)`.
6. **The `w:anchorLock` port of docx4j `82c64bf14`**, uncommitted in this working tree (`docx4j_py/model/property_catalogue.py`, `tests/content/test_properties.py`). Related work only; this CR leaves those two files to it. It reads `w:anchorLock` as stated, since docx4j's CR-059 corrected its default to false, and the refresh makes this copy's default false too. It should land before or with Phase A, which touches the same file (`is_empty_val`, `merge_u`).
7. **CR-008** (table conditional formatting) will read `w:tblLook`'s six ST_OnOff attributes, an enumeration with `on` and `off`, through `on_off`.
8. **CR-012**: a kept branch's typed content is read by the same readers. VML's `t`/`f` is a string enumeration, untouched.
9. **CR-014**: `w:temporary` and `w:text/@w:multiLine` read through `on_off` (section 2.8 item 12).

### 4.10 Departures

**From Word.** None in reading. On save, the source's spelling (Q5) and the repaired enumerations (Q6) are kept, where Word writes its own form.

**From docx4j CR-060:**
1. **The bound value.** docx4j's adapter reads the value into a `Boolean`; here the value is kept as a `str` and read by a function. The answers are the same.
2. **D2.** docx4j writes `true`/`false`; here the source's spelling is kept (Q5).
3. **D3.** docx4j drops invalid `highlight`, `vertAlign` and `em` on save; here they are kept (Q6).
4. **The scope.** docx4j's adapter reaches every ST_OnOff attribute by type. Here the readers are switched one by one (section 2.8's 20 rows), and the converter reaches every `bool` field.

**From core-ts CR-015**, as its registry entry states it: none foreseen. Its leniency lives in jsonix's parse, and this package's parse is lenient already.

### 4.11 Errors

No new error. A value Word reads is never refused: loading stays lenient, and `LoadOptions(strict=True)` raises on skipped content, not on a kept value, as now. Under the fork's exact mode (section 4.1a) a value kept on purpose (`" false"`, `" yellow"`) raises no `ConverterWarning`, while a genuine failure (`"off"` for `xs:boolean`, `"Yellow"`) keeps its warning. Both are recorded in the fork's `SkippedReport.kept` and surfaced as `part.kept` (Q16), CR-002 section 12.3 item 7's "value" entries.

### 4.12 Not in this CR

- **Underline styles and emphasis marks in rendering** (docx4j's D6). Nothing here renders.
- **An operation that writes Word's form** (section 4.7).
- **The other converters' trims**, for example `w:sz w:val=" 24"`. Unmeasured.
- **`w:tblLook`'s readers** (CR-008).

## 5. Phases

| phase | what | size |
|---|---|---|
| A | **On/off and the refresh.** Section 4.5's refresh, `PATCHES.md` and the regeneration. Section 4.1a: on the fork's 26.2.3, `runtime.parser_config()` sets `exact_types={bool, Enum}`, `part.kept` and `package.kept` surface the record, and a test holds every `ParserConfig` in the tree to `parser_config()`; the floor `>=26.2.3`. Section 4.1: `model/values.py` with `on_off`, `is_on` and `xsd_boolean`; every reader of section 2.8 switched, the facade's six and the chart's `_bool` with them. The fixtures (Q11) and the on/off tests (section 6.1). Section 6.2's gates | 1 to 1½ days |
| B | **The enumerations and colour**, on CR-016 Phase B's `Font`. Sections 4.2 and 4.3: `in_enumeration` in the merge and the readers; `colour()` and `NETSCAPE_COLOURS`; `Font.underline`, `highlight_color` and `color`; the facade's `ColorFormat`. The tests over rounds 2 to 4 | 1 to 1½ days |
| C | **`effective()`.** Section 4.4: the per-class table, the scope (Q10), its tests over the `w` namespace's defaults (and `w14`'s and `w15`'s, once checked) and the hazards; `w14:checked`, `w:fldLock` and `w:enforcement` read through it | ½ day |
| D | **Round five, the Word check and the records.** Round five's answers folded in (Q12): the pins of Q7 and Q8 confirmed or turned. The Word check (section 6.3). The records: `tests/README.md`'s map and the fixtures' provenance, the CHANGELOG, CR-003 section 20.3's new departure (Q9), and the index row | ½ to 1 day |

**Order: A, C, B, D**, each committed and verified on its own.
- **The fork's stage 13 before A** (section 4.1a), built by the fork session and released as 26.2.3 by the user.
- **A first here.** It closes the one save that changes meaning, and the `"off"` readings, with nothing from CR-016. Q1 proposes it for 0.2.0, which then waits for 26.2.3 too.
- **C next.** It needs only A's refresh.
- **B after CR-016 Phase B**, whose `Font` it changes underneath.
- **D last.** Its Word check covers A's and B's saves together. Round five can run as soon as it is built, and its answers are needed before B's pins are final.

## 6. Tests and the gates

### 6.1 The oracles

**The fixtures (Q11).** The 20 probes and Word's 20 saves (Appendix A), copied byte for byte from the share into **`tests/fixtures/onoff-fun/`** with a `SHA1SUMS` manifest, as `tests/fixtures/mc-fun/` was. The provenance goes into `tests/README.md`'s table: built by the docx4j-portfolio session (`onoff/build.py` to `build4.py`), run by the user in Word 365 on 2026-10-24, and the share's `README.txt` with its SHA-1. The PDFs and the round 1 screenshots stay on the share; the saves carry Word's answers.
- **The name.** `onoff-fun` is the share's own name, and core-ts's CR-015 uses it (`test/onoff-fun/`). It sits beside `mc-fun/`, the precedent of a share's probe folder in this tree, and it holds the colour and enumeration rounds too.
- **If docx4j's Phase 0** commits the same files first (core-tests `src/test/resources/runprops/word365/`), each copy is checked against its blob there too.

**The tests**, `tests/content/test_run_property_values.py` (new):
1. **Every marker, every reader, against Word's save**, as section 2 measured. A reader's answer on the probe must equal its answer on Word's save. The readers are the resolver's toggles, `Font` (direct and effective), the text model's runs, Markdown, the facade, `Paragraph.text` and `search` (for T10), and `format_span`'s write.
2. **The rules as tables**, so that the oracle isn't only agreement with a save:
   - `on_off` over rule 1's 15 spellings, an enumeration member, and `None` with each default;
   - `in_enumeration` over round 2's values;
   - `colour` over every colour marker, and over the 140 names against CSS 3 less the seven.
3. **The saves.** Each probe's main part read and saved is canonically identical to its source, `boolean-spelling` aside (section 2.7's second row). `" false"` and `"false "` are written as read.
4. **The parse** (on the fork's `exact_types`): the four literals bind bools, an exact member binds it, and anything else is kept as its `str` and recorded in `part.kept` with its reason; every `ParserConfig` in the tree comes from `runtime.parser_config()` (a test that walks the tree). The fork's own 25 tests cover the option itself.
5. **The facade against python-docx 1.2.0** where it is installed (the `parity` extra), as the facade's parity test is: python-docx's answer wherever it answers and agrees with Word, Word's where it raises, and the three colour misparses listed.
6. **`effective()`** (Phase C): each `w` default against docx4j's XSD at `46c727bf8`; the hazards; the `LookupError` outside the scope; `w14:checked` bare.

### 6.2 The gates

**Phase A:**
- **The readers' audit, before and after**, as CR-013 section 6.2 ran it: `scripts/mce_branches.py readers --holding all` over `out/cr012-b-corpus/`, its `COMMITS.txt` unchanged. **Expected: no answer moves.** The census (section 2.10) found one document with odd values, and no reader reads its `w:wordWrap`. Any move is a finding, explained before the phase lands.
- **The corpus round trip with the new converter.** Over the 422 documents, no part gains a canonical difference and nothing is newly skipped (CR-012's bar). `scripts/roundtrip.py` over `samples/` stays at 50 of 50, with the same `boolean-spelling` count.
- **The rest:**
  - `generate.sh --check` after the refresh;
  - the 46 parity goldens, zero differences;
  - CR-011's ratchet unmoved;
  - CR-010's Word-save comparisons;
  - acceptance, regenerated and compared (`compare_acceptance.py` against a copy made first), which no artefact should change;
  - `scripts/threads.py`, since the registration is import-time state;
  - `tests/test_performance.py`, since the converter is on every parse (the fork session measured nil);
  - the full suite.

**Phase B:**
- **The audit again, with CR-016 B's font reader.** Expected: no move (no invalid enumeration or named colour in the corpus).
- **The colour markers.** For each of the 90, `Font.color` and the facade's `color.rgb` equal `colour()` of the value in Word's save.

**Phase C:**
- **The audit, unchanged.** The readers C switches answer absence as they did, but for a bare `<w14:checked/>`, and the corpus, `tests/` and `samples/` hold 2 `w14:checked`, neither bare.

**Every phase:** the full suite with no failure, the two subset tests, `generate.sh --check` and the round trip.

### 6.3 The Word check (Phase D)

CLAUDE.md asks for a Word check after any change to marshalling. Phase A changes what a save writes for `" false"`, and B's merge changes nothing written. Staged in `__tmp/python-cr018/`:
1. **The 20 probes, each loaded here, every part read (so re-marshalled), and saved.**
   - **Expected:** each opens as its source does. There is no repair prompt but the source's own (`props-enum-missing` and `props-enum-other` prompt, as their sources do; Q6).
   - Every marker draws as the source's does, and the share's PDFs are the reference. K5 and K6 are bold.
2. **The edits.**
   - `range.font.bold = True` across paragraphs over `"off"` runs;
   - `font.bold = False` over `" false"`;
   - `Font.color` written over a named colour;
   - the facade's `font.underline = None` over an invalid `w:u`.
   - **Expected:** no repair prompt, and each run as written.
3. **Round five's probe** (Q12), if built by then, run as rounds 1 to 4 were.

## 7. Risks

1. **A parser built without `parser_config()`** reads `" false"` as `False` and `" yellow"` as `YELLOW`, as today. The test of section 6.1 item 4 holds every `ParserConfig` in this tree to `parser_config()`; a caller's own parser is the caller's choice. (The first draft's risk, a process-wide registration, is gone with option 1.)
2. **A caller's own reader meets a kept `str`.** Code outside this package that tests `b.val is not False` now meets `" false"` as a `str`. That reads on, which is Word's answer, as `"off"` read on before. `on_off` is public (Q13), and the CHANGELOG names it.
3. **Readers move for other producers' files.** POI's `"off"` (docx4j #585) now reads off everywhere, and pandoc's `w:wordWrap w:val="off"` too, though nothing reads `w:wordWrap`. That is the point, but a caller comparing output across versions sees it. The census found no move in the corpus.
4. **Unmeasured extensions.** These are taken by extension and are round five's:
   - on/off outside `w:rPr` (`w:keepNext`, `w:tblHeader`, `w:titlePg`, `w:trackRevisions`, a mark's `w:vanish`);
   - `bCs` and `iCs`;
   - `w14` and `w15`;
   - the ST_OnOff attributes (`w:fldLock` and the rest);
   - invalid values over styles (Q8);
   - the table's 99 unmeasured names (Q7).
5. **The enumeration merge** changes the resolver's output for an invalid direct value. docx4j's JAXB already nulls such a value, so the two engines agree after the change, and no golden holds one.
6. **The counterparts move.** docx4j's CR-060 is accepted and being implemented; core-ts's CR-015 is uncommitted and may change. This CR follows the rules, not either implementation, and section 4.10 records where it departs. A later change there is a later amendment here.

## 8. Questions (open; each with a recommendation)

1. **Q1. Does 0.2.0 wait for CR-018?** *Recommended: yes for Phase A only (the refresh and the on/off reader), as this repository's session recommended to the user on 2026-10-24; B to D in the next release.* Phase A closes a save that changes meaning (`" false"` written `"0"`, finding 1), which 0.1.0 has too. It needs the fork's 26.2.3 (section 4.1a) and nothing from CR-016, and it is a day or a day and a half here. The refresh changes metadata only.
2. **Q2. How does `" false"` survive the parse?** *Recommended (revised 2026-10-24, the user assuming a fork release): the fork's `ParserConfig(exact_types={bool, Enum})`, stage 13 item 29, set by `runtime.parser_config()`* (section 4.1a). It is explicit per parse, with no global mutation, no import-order trap and no test swapping, and it covers the enumerations' padding too. The first draft's option 1, a converter registered through `register_converter`, needs no release, binds the same values, and stays the fallback if 26.2.3 is delayed.
3. **Q3. The fork's stage 13: all three items in 26.2.3?** *Recommended: yes*, items 29 (`exact_types`), 30 (`SkippedReport.kept`) and 31 (`XmlVar.schema_default_value`), about 100 lines and 25 tests, all off by default, built by the fork session on the user's go and gated by this package's suite with the option on, `roundtrip.py` and the interleaved benchmark (section 4.1a). Item 31 alone could wait, since only Phase C uses it, but it is small and independent. (The first draft's Q3, whether a process-wide registration is acceptable, falls away with option 1.)
4. **Q4. The readers' scope: run properties, or every on/off reader?** *Recommended: every reader of section 2.8, through the one function.* The rule is ST_OnOff's, the cost is the same, and a narrower fix would leave `w:trackRevisions w:val="off"` and `w:tblHeader` reading as they do. The readers outside `w:rPr` are by extension until round five.
5. **Q5. What does a save write?** This is the portfolio's D-question. *Recommended: the source's spelling* (section 4.7). It is this package's round-trip rule, canonical identity with `boolean-spelling` the one tolerated difference; it never changes meaning; and it needs no change to `canon.py`. Word's form can be a later explicit operation.
6. **Q6. Invalid or missing `highlight`, `vertAlign`, `em` on save.** *Recommended: keep them as read, departing from docx4j's accepted D3.* Nothing is dropped silently here, the meaning is the same either way, and a save keeps the prompt the source already had. core-ts's D2 keeps them until round five. Revisit with round five if Word turns out to read any of them otherwise.
7. **Q7. The colour table.** *Recommended: all 140 names of the classic table, as docx4j's D5 takes them*, with the 41 measured names as tests. Round five measures the other 99 in one file. All seven of CSS 3's extra keywords were probed and dropped, so the fit is strong.
8. **Q8. A value outside the rules, over a style that states the same property.** *Answered by round five (section 3.1b): both read as if absent, so the merge skips them and the style's value shows*, an enumeration and a colour alike. The one exception, `w:u w:val="true"` removing the style's underline, is a pinned case in Phase B. (The original recommendation, a colour as a stated automatic, is withdrawn.) Rule 2 says "as if absent", and S11 is the missing-value case of it. Rule 3 says "automatic", and Word's save can't tell the two readings apart without a style.
9. **Q9. The facade: Word's reading, or python-docx's?** *Recommended: Word's throughout, keeping python-docx's tri-state shape*, recorded in CR-003 section 20.3. python-docx raises on 121 of the 227 probe reads and agrees with Word wherever it answers, but for three colour misparses (section 4.8). So following Word departs from python-docx's answers only there, and the facade stops failing or misreading files Word opens without a prompt.
10. **Q10. `effective()`'s scope: all 546 defaults, or `wml`'s?** *Recommended: the `w` namespace (51 defaults), with `w14` and `w15` added once Phase C has checked their few, raising `LookupError` elsewhere until a namespace's defaults are audited.* Only `wml`'s defaults have been audited, by docx4j's #697 and CR-059. Unaudited defaults were how docx4j's getters went wrong, and no one has audited DrawingML's.
11. **Q11. The probes as fixtures, under Apache-2.0.** *Recommended: copy the 40 files byte for byte into `tests/fixtures/onoff-fun/` with a `SHA1SUMS`* (section 6.1, Appendix A), once the user confirms they may be used here under Apache-2.0, as CR-013 Q8 did. They are the portfolio's builds and the user's own runs.
12. **Q12. A fifth probe round.** *Recommended: yes, built by the portfolio session as rounds 1 to 4 were, and shared with docx4j and core-ts* (core-ts plans one too). It covers:
    - invalid `u`, `highlight`, `vertAlign` and `color` over a style stating each;
    - `bCs` and `iCs`;
    - `"off"` on `w:keepNext`, `w:tblHeader`, `w:titlePg`, a mark's `w:vanish`, and `w:trackRevisions` in settings;
    - a bare `<w14:checked/>`;
    - `w:fldLock` with `" false"` and `"TRUE"`;
    - the 99 unmeasured names.

    Phase B's pins wait for it, and nothing else does.
13. **Q13. Are the readers public?** *Recommended: yes, `docx4j_py.model.values` (`on_off`, `is_on`, `in_enumeration`, `colour`, `effective`), documented in the CHANGELOG.* A caller reading the typed tree meets the kept strings (risk 2) and needs Word's rule as much as this package does.
14. **Q14. Q11's source, after docx4j CR-060.** *Recommended: copy the 40 files from docx4j's `runprops/word365/` at `868a1797b`* (section 3.1a), byte for byte with their SHA-1s, as `samples/` holds docx4j's fixtures under Apache-2.0. The share then serves only round five's files and any later ones, which still need Q11's confirmation.
15. **Q15. A padded enumeration value (`" yellow"`), kept by the fork's exact mode.** *Recommended: read it as outside the enumeration, unset, pinned until a probe measures Word's whitespace for `u`, `highlight`, `vertAlign` and `em`.* Word is case-exact there and untrimmed for on/off, so untrimmed is the likelier rule; the save keeps it as read either way, so the reading is the only thing at stake. A sixth probe round can take it with the on/off words over a style (section 3.1b).
16. **Q16. Surfacing the fork's `SkippedReport.kept`.** *Recommended: `part.kept` and `package.kept` beside `part.skipped`, each a list of the fork's `KeptValue`, and `describe()` counting them by reason (`"exact"`, `"invalid"`), not listing them.* An agent learns that a document holds values Word reads by its own rule (POI's `"off"`, padded booleans) without a warning per value; `LoadOptions(strict=True)` still raises on skipped content only. That is CR-002 section 12.3 item 7's "value" entries, here at last.

## 9. Effort

| phase | days |
|---|---|
| the fork's stage 13 (section 4.1a), in the fork session, and the 26.2.3 release (the user's) | ½ to 1 |
| A, on/off and the refresh | 1 to 1½ |
| B, the enumerations and colour | 1 to 1½ |
| C, `effective()` | ½ |
| D, round five, the Word check and the records | ½ to 1 |
| **total** | **3½ to 5½**, the fork's included |

docx4j's own estimate was a day and a half for the same ground (CR-060 section 5). The larger part here is the readers: section 2.8's twenty rows, in some twenty modules, where docx4j's adapter reaches them all by type, and each is tested against Word's saves. Phase A's audit and the corpus round trip are the other part. As in CR-016, these are the CRs' usual unit, not a forecast of the calendar.

## Appendix A. The probes (SHA-1, as on the share on 2026-10-24)

`README.txt` is `2af2a57b91392e3205780315398c7262a810f71c`. The 40 `.docx` files are 20 probes and Word 365's saves, 400 KB in all:

```
e372163d90b474180a01bfc756e689a96642f99b  colour-case.docx
0c30c8c8cfff9f1c8d7b99766a56493e75557db3  colour-case-word365.docx
7c9f18f31293919a929b9c904b21eec904b2a669  colour-hex.docx
e8456d8ff64917c7028a8dc4afc81225c32dfe82  colour-hex-word365.docx
cee2adc4e3c6aefcd9e1108fc9788058f4b30c4d  colour-names-css.docx
4c12a90671558bdfef3e85379db22d7d7697b9e8  colour-names-css-word365.docx
2491353261ba8851c2d72c4df990b6b60bd89132  colour-names-highlight.docx
e2e86617865e528c060c4cdefdc706eb806bc0c7  colour-names-highlight-word365.docx
65978aa2e0b38b826d454725aa5f7e18f7c9587a  colour-table-grey.docx
e9c2a44213885a90002e46738a8c2a1bcd2656c5  colour-table-grey-word365.docx
aca5e1daceb16381ab8bd170996a1089874efb28  colour-table-other.docx
d67f1557624e45597ba4becde76877a0f5313038  colour-table-other-word365.docx
6e789e6b04a2d08e7669c01ef999bb39dda04abc  onoff-empty.docx
0db1a6e84df66878bffcb6454fff43dcfbac9b95  onoff-empty-word365.docx
c6f6c165ca7587e791c94c65c52d6252178c477d  onoff-space.docx
b52d9a75b23f9de743d5bcee67f9e63d7d61c723  onoff-space-word365.docx
661da9c49c2e2fa19d8e9ce7d1d373a0dc135085  onoff-TRUE.docx
58d5b2e935afe2211f69c2790eadf48f2bba77df  onoff-TRUE-word365.docx
8e174d7fca73636b8947a26af8a71b8ddece5233  onoff-valid.docx
43b5692586c0b4cfc0176e1eef8191eb6d68723e  onoff-valid-word365.docx
80c151c4138791f1bc7a97504160cbb20459159f  onoff-yes.docx
6d2fb94532205693f09091af6c9d46db27c6e931  onoff-yes-word365.docx
3bbd0b163a8407afd86f58035417fab7c75a83e7  props-b-case.docx
6b1cdb50e17c1cf252637ccd2be524afc0b29aaf  props-b-case-word365.docx
276fa157dd903b292175b86c1b8ed5cd85829a6b  props-b-styles.docx
a361bc879ecd8bcf724b65e0246988fcb72185df  props-b-styles-word365.docx
3272a071db219943f37d4ce64d147d75b2890c2d  props-enum-missing.docx
020e70a4a6ad354f3d433e5d3fec689da557d27e  props-enum-missing-word365.docx
30c79b39ed46db09e27d3d06287e0de075ccdf6b  props-enum-other.docx
a72b231728c611901a3a54ac375cbb322b5cc4b9  props-enum-other-word365.docx
2f4d51a81ddb482f32fcd2acb1c229f831b0716c  props-toggles-1.docx
26361f73f4bd33b7ca5c8eaaa5ce6764b3596fff  props-toggles-1-word365.docx
903d3b1040f85a20d7213e6257ff737999ef4cf7  props-toggles-2.docx
f51ea13ad04ea388d599eae0e0b7ebc35794ba61  props-toggles-2-word365.docx
8f941072898265474d75d98573d7395139c18add  props-u-case.docx
deb116c8b3868beb3d4471735428467cf6f02748  props-u-case-word365.docx
2f4228da02345d34f9ff54a0652479caa94ffed1  props-u-other.docx
c506430aa29a324491d3d9ea0786b80b47480dcf  props-u-other-word365.docx
67f0ee04edd25010cd05043d85dd160041520a9d  props-u-valid.docx
63ec049068b4612763b2a4767f1e8b86933f2ec8  props-u-valid-word365.docx
```
