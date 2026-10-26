# CR-022: Integer measurements refused beyond Word's range as code sets them --- the write-side guard, after Word's 16-bit wrap

**Status:** Proposed 2026-10-26, at the user's request ("option 1, and also add the out-of-range report to CR-020's Phase B list"). It follows the user's suspicion the same day "that we should guard against users setting twip measurements to large integers in each of docx4j, python and core-ts". The portfolio session coordinates it as one CR per port. Round 5 of the shared Word probe is the shared phase 0. It is docx4j CR-066's phase 0 (docx4j-25, `1be8b6515`, local), on the share at `vbShares/Office 2016/__tmp/cr066/`: 191 cases in six groups, 22 single-attribute families and 48 singles. Checked here: all 76 files load with nothing kept or skipped. **Round 5 was run 2026-10-26** (the user, Word 365: all six groups opened clean, no repair). **Phase 0 is done.** The range is per attribute, not per schema type (section 3.1), so D1 is a table per attribute, and Q1 is restated. **Two more measurements came the same day:**
- core-ts's check 71 measured Office JS's own setters (section 3.2), which revises Q1's recommendation;
- round 6, the file side of the attributes round 5 left unmeasured (481 cases, `__tmp/cr066/round6/`), **was run 2026-10-26** (the user: "cr066/round6 all opened cleanly"). docx4j-25's reading is section 3.3, and with it D1's table covers every attribute Word could measure. **Nothing has started.** Phase 0 gives the limits, and nothing here is built before it. Three phases (section 6), 1 to 1½ days after round 5. **Decided 2026-10-26** (the user: "CR-022 questions take the recommended answers"; section 8): Q1 to Q7 as recommended. **Nothing has started.**

*The number:* CR-021 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057.

**Depends on:**
- **Round 5 of the shared probe, docx4j CR-066's phase 0**, for this CR's limits table. Per attribute type it holds 32,767, 32,768, 65,535, 65,536 and 65,537; the signed ends -1, -32,768 and -32,769; the wrap controls 7,205,001 and 1,205,003; and each type's user-interface or schema limit and the value past it. Its `cases.txt` gives, for each case, the value if Word wraps at 16 bits and the value if it clamps. The question per family is wrap, clamp, keep or refuse.
- **Counterparts:** docx4j CR-066, the write side (its recommendations, awaiting the user: a WARN at marshal and unmarshal with the value Word will read, never a throw at save and never a rewrite on read; a hard guard where docx4j computes twips from units, `TblFactory`), core-ts's content API setters (it is asked what Office JS does with the same values), and docx4j-generated-objects-ts's builders, whose CR-015 table of 178 attributes may give the per-type table.
- **This repository:** CR-020 Phase B reports an out-of-range value already in a tree. That's the read side, and the user added it to CR-020 the same day; it uses this CR's table. CR-021 reads decimals and the other forms in such attributes, and its Q7 recommends not emulating the wrap on read.

## 1. Why

Word 365 stores twips measurements in 16 bits. An integer too large to fit isn't refused or clamped: Word keeps its low 16 bits. The evidence is an accident in CR-065's probe. A slip in docx4j-25's generator put three large integers into every family file's styles, settings and header, and Word's round 2 saves (`vbShares/Office 2016/__tmp/cr065/families/cr065-{B..H}-word365.docx`, checked here and by the portfolio session) show each one wrapped:

| Where | In the file | Word wrote | What happened |
|---|---|---|---|
| `styles.xml`, `w:ind w:left` (signed twips) | 7205001 | -3959 | 7205001 mod 65536 = 61577, read as signed: -3959 |
| `settings.xml`, `w:defaultTabStop` (unsigned twips) | 7205002 | 61578 | mod 65536 |
| a header's `w:spacing w:before` (unsigned twips) | 1205003 | 25355 | mod 65536 |

The regenerated probe no longer carries the slip, so these three saves are the record.

Nothing here refuses such a value as code sets it (section 2). The user who sets a 500 cm page gets a file that Word opens at a size nobody asked for, with no error and nothing in the file to say so.

## 2. What exists here (measured 2026-10-26 at `f5f5245`)

### 2.1 Nothing refuses a value Word will misread

- **The content API:**
  - `Paragraph.left_indent = 400000.0` (points) saves `w:ind w:left="8000000"`, with the matching `w:hanging` when `first_line_indent = -400000.0`;
  - `space_before = 5000.0` saves `w:spacing w:before="100000"`.
- **The python-docx facade:**
  - `paragraph_format.left_indent = Cm(1000)` saves `w:left="566929"`;
  - `section.page_width = Cm(500)` saves `w:pgSz w:w="283465"`.
- **What Word would show,** by the measured wrap (an inference, not checked in Word for these values):
  - the 500 cm page opens 37.6 cm wide (283465 mod 65536 = 21321 twips);
  - the indent becomes about 8 cm (4608 twips);
  - the space before about 24 cm (34464 twips).

### 2.2 Where code sets a measurement

- **The content API:** about 13 writes, each a conversion from points (`* TWIPS_PER_POINT`, or half-points for `Font.size`):
  - `model/content/paragraph.py`: the indents and the spacing;
  - `table.py`: widths and heights;
  - `lists.py`: a level's indents;
  - `font.py`: `Font.size`.
- **The python-docx facade:** every setter converts through `docx/shared.py`'s `Length` (`.twips`, `.emu`), about 12 conversions.
- **The tree-layer builders:** `wml/builders.py`'s `size` writes half-points (`round(size * 2)`).
- **Not setters, and out of this CR's scope:**
  - the parse, which reads (CR-021 for its forms; CR-020 Phase B reports a range violation);
  - the save normaliser, which writes back what was read;
  - a value assigned straight to a generated field (`ind.left = 8000000`), which CR-020 Phase B's validator reports, for the reason CR-020 Q4 gives against a `__setattr__`.

### 2.3 What the model knows about an attribute's type

The generated fields carry neither the schema's simple type nor its facets: `CtPprBaseInd.left`'s metadata is its type (`Attribute`) and namespace alone. `ST_TextScale`'s `minInclusive 0`, `maxInclusive 600` in `schemas/wml/wml.xsd` don't reach the field. So a limit per type has to be derived from the schema copy: each integer attribute's simple type (`ST_TwipsMeasure`, `ST_SignedTwipsMeasure`, `ST_HpsMeasure`, `ST_EighthPointMeasure`, `ST_PointMeasure`, `ST_TextScale`, `ST_DecimalNumber`, ...). CR-021 counts 224 integer-only attribute fields in `w`, under 70 names.

## 3. What is known, and what round 5 must answer

- **Measured:** the 16-bit wrap in twips, signed and unsigned, at three values far past the limit (1.2 million and 7.2 million).
- **Unmeasured:**
  - the exact boundaries;
  - whether Word wraps, clamps or refuses between its user interface's limits and 16 bits;
  - every other family:
    - half-points (`w:sz`; Word's largest font is 1,638 pt, 3,276 half-points);
    - eighth-points (a border's `w:sz`);
    - points (a border's `w:space`);
    - the percent (`w:w`, which the schema caps at 600);
    - decimal numbers;
  - what Office JS does with the same values (core-ts is asked).
- **Wider, and not to be caught:** comment and bookmark ids and `w:numId` are 32-bit decimal numbers, and DrawingML's EMU sizes are 64-bit. A per-type table that wrongly lists them would refuse legitimate values.
- **Word's user interface stops sooner than 16 bits for some measures:** page size goes to 22 in (55.9 cm, 31,680 twips), where 16 bits signed reaches 57.8 cm and unsigned 115.6 cm.

### 3.1 Round 5, measured (docx4j CR-066 section 3.1; the user's run, 2026-10-26)

No value was refused or repaired. A negative in a schema-unsigned attribute is accepted everywhere, and read as its two's complement. What Word keeps is the width of the field in Word's binary document structures: `dyaBefore` unsigned, `dyaTop` signed, `dyaRowHeight` and `lspd.dyaLine` signed with the sign as the rule, `brc` in 8 and 5 bits. docx4j-25 expects that to be stable across Word versions. Checked here in Word's saves of groups P, B and R (for example `w:before` 65535 → "-1", 65536 → 0, 1205003 → 25355; a border's `w:sz` 256 → 0, 32767 → 255; `w:sz` 3277 kept, 65536 → 0).

| Attribute | What Word does | Read faithfully |
|---|---|---|
| `w:ind/@w:left`; `w:spacing/@w:line` with `lineRule="auto"` (negatives written); `w:rPr/w:spacing/@w:val` (31681 kept); `w:position/@w:val` (3169 kept); `@w:w` of `w:tblW`, `w:tcW` and `w:tblInd`, in `dxa` and in `pct` (5001 and 32767 kept); `w:pgMar/@w:top` (`ST_SignedTwipsMeasure`, as is `bottom`; 32768 → -32768, -1 kept) | signed 16 bits, wraps; no interface clamp | -32,768 to 32,767 |
| `w:spacing/@w:before` (32768 kept, 65536 → 0, 1205003 → 25355; 0xFFFF is written as "-1", so 65535 and -1 both come back as "-1"); `w:sz/@w:val` (3277, 32768 and 65535 kept; 65536 → 0; -1 → 65535); text scale `w:w/@w:val` (601 and 65535 kept, past the schema's 600; 65536 → 0; -1 → 65535) | unsigned 16 bits, wraps | 0 to 65,535 |
| `w:pgSz/@w:h` | unsigned 16 bits, with a minimum of 144 and no maximum (65535 kept, 65536 → 144) | 144 to 65,535 |
| `w:tab/@w:pos` | signed 16 bits, then clamped to ±31,680 (32767 → 31680, 32768 → -31680, -32769 → 31680) | -31,680 to 31,680 |
| `w:pgSz/@w:w` | unsigned 16 bits, then clamped to 144 to 31,681 (32767, 65535 and -1 → 31681; 65536 → 144; 1205003 → 25355) | 144 to 31,681 |
| `w:trHeight/@w:val` | signed 16 bits, the sign carrying the rule: negative means `exact` (32768 → `exact` 32768; 65535 and -1 → `exact` 1; 65536 dropped; 65537 → `atLeast` 1) | 0 to 32,767 |
| `w:spacing/@w:line` with `lineRule="exact"` | stored as the negative, so after the wrap the sign re-derives the rule (exact 65535 → `atLeast` 1; exact -32768 → exact 32768; exact 7205001 → `atLeast` 3959) | 0 to 32,767 |
| a border's `@w:sz` | 8 bits, mod 256 (256 → 0, 32767 → 255, -1 → 255) | 0 to 255 |
| a border's `@w:space` | 5 bits, mod 32 (32 → 0, 255 → 31, -1 → 31); no clamp at the interface's 6 pt | 0 to 31 |
| `w:pgNumType/@w:start` | signed 32 bits, wraps (2147483648 → -2147483648; 4294967295 → -1; 4294967296 → 0) | -2,147,483,648 to 2,147,483,647 |

**Unmeasured:**
- the other 26 `ST_TwipsMeasure` attributes (`gridCol`, the other margins, `hanging`, `firstLine`, the distances);
- `w:spacing/@w:after`, assumed as `before`;
- the other border edges, assumed as `top`.

**Word's interface limits are mostly not enforced:**
- 1,638 pt fonts (3,276 half-points), the 600% scale, the 6 pt border space and the 22 in page are all kept past their limits;
- the exceptions are the tab stop's and the page width's clamps, and the page's minimum of 144 twips.

So the "faithful" column is the guard's range: inside it, what code sets is what Word reads.

### 3.2 Office JS's own setters (core-ts's check 71; docx4j CR-066 section 3.2)

core-ts's check 71 (96 cases, Word 365 16.0.20326, the user's run, 2026-10-26) sets values through Office JS and reads Word's save:
- **The indents and `TableCell.columnWidth` aren't guarded.** `leftIndent`, `rightIndent`, `firstLineIndent` and `columnWidth` take any value and store it wrapped exactly as the file reader does: 1638.4 pt → -32768 twips, 400,000 pt → 4608. Only a negative `columnWidth` is refused.
- **`w:firstLine` and `w:hanging` are one signed field.** A first line F is stored as F, a hanging H as -H, so first line is 0 to 32,767 and hanging 0 to 32,768, and a wrap crosses between them: `firstLineIndent` 1638.4 pt writes `w:hanging="32768"`.
- **Refused at Office JS's own limit, below storage,** with `InvalidArgument` and nothing written:
  - `spaceBefore`, `spaceAfter` and `lineSpacing` at 1,638 pt and up, and every negative;
  - `setLevelIndents`' text indent at 1,638 pt and up and at -1,638.4;
  - `font.size` above 1,638 and at 0.

So Office JS guards some measures, at Word's interface limits, and leaves the rest to wrap.

### 3.3 Round 6, measured (docx4j CR-066 section 3.4; the user's run, 2026-10-26)

Everything opened; nothing was refused or repaired. docx4j-25's reading, spot-checked here in Word's saves (group A's `w:firstLine` 32768 → `w:hanging="32768"`, 65535 → hanging 1, 65536 → firstLine 0, -1 → hanging 1; `w:defaultTabStop` 32768 and 65535 kept, 65536 → 0):

| Attribute | What Word does | Read faithfully |
|---|---|---|
| `w:ind/@w:firstLine` and `@w:hanging` | **one signed 16-bit field**: first line F stored as F, hanging H as -H, after the wrap; written as `firstLine` when 0 or more, as `hanging` with the magnitude when negative (firstLine 7205001 → hanging 3959; hanging 7205001 → firstLine 3959) | first line 0 to 32,767; hanging 0 to 32,768 |
| `w:spacing/@w:after`; `w:pgMar` `left`, `right`, `header`, `footer`, `gutter`; `w:defaultTabStop`; `w:hyphenationZone`; `w:szCs`, `w:kern` | unsigned 16 bits (0xFFFF written "-1" in the first two) | 0 to 65,535 |
| `w:spacing/@w:line` with `lineRule="atLeast"` | signed 16 bits stored positive; a wrapped negative becomes `exact` (65535 → exact 1) | 0 to 32,767 |
| `w:ind/@w:right` (`start` and `end` read and written back as `left` and `right`); `w:pgMar/@w:bottom`; `w:framePr` `w`, `hSpace`, `vSpace`, `x`, `y` (0 not written); `w:tblCellSpacing`, `w:tblCellMar`, `w:tcMar` in `dxa`; `w:tblpPr`'s six (0 not written); `w:cols/@w:space`; `w:lnNumType/@w:distance` | signed 16 bits | -32,768 to 32,767 |
| `w:framePr/@w:h` | 15 bits and a flag: bit 15 turns exact off (32768 → 0 with no `hRule`; 65535 → 32767) | 0 to 32,767 |
| `w:col/@w:w` | signed 16 bits, then a negative replaced by a computed width (9360); 32767 kept though wider than the page | 0 to 32,767 |
| `w:col/@w:space` | signed 16 bits, then a negative becomes -1 | 0 to 32,767 |
| `w:trHeight/@w:val` with `hRule="exact"` | signed 16 bits, the sign as the rule (as round 5) | 0 to 32,767 |
| `w:trHeight/@w:val` with `hRule="auto"` | discarded at every value | (no reading) |
| every border family (a run's `w:bdr`, `w:tcBorders`, `w:tblBorders`, `w:pgBorders`) | `w:sz` 8 bits, `w:space` 5 bits, as round 5's paragraph borders | 0 to 255; 0 to 31 |
| `w:fitText/@w:val` | not 16-bit: all ten values kept as written | (no limit found) |

**Not measurable:**
- `w:wBefore` and `w:wAfter`: Word drops them with an inconsistent `gridBefore`;
- `w:gridCol`: Word recomputes it, even under a fixed layout;
- a level's `w:legacy`: Word 365 drops it from `numbering.xml`.

These go unchecked (Q3).

Group E (the digitless forms) belongs to CR-021: Word reads `"."`, `"+"`, `"-"` and `"+."` as 0.

## 4. Design

### 4.1 The limits table (D1)

One committed table, keyed by attribute: (element, attribute), with the context where the range depends on it (`w:spacing/@w:line` by `lineRule`). Each row gives the faithful range of section 3.1 and cites its round 5 cases. It's hand-written from round 5 and lives under `docx4j_py/resources/`, which `KEEP` already holds.

An attribute round 5 didn't measure has no row and isn't checked (Q3). A row's element name has to match every class that element name binds, as CR-021's canon table does. That's checked by a test against the model's integer fields, so a per-holder class can't slip past.

A per-type table, derived from `schemas/`, was the draft's design. It's dropped: round 5 shows one schema type (`ST_TwipsMeasure`) holding attributes of five different widths.

### 4.2 One check (D2)

`docx4j_py.model.values.measure_in_range(owner, field, value)` raises `MeasureRangeError(ContentError, ValueError)`, with code `value.out_of_range` (Q2). The message names the property, the value in the caller's units and in the file's, and Word's range. For example: "left indent 400000 pt (8,000,000 twips) is beyond Word's range, -1,638 to 1,638 pt". `ValueError` is a base so that a caller catching the language's error for a bad value catches this one.

### 4.3 Where it runs (D3)

Every write of section 2.2: the content API's setters, the facade's (at the point a `Length` is written to a field), and the builders. A census like CR-021's D2 holds the list whole: every write of an integer measurement under `docx4j_py/model/content`, `docx4j_py/docx` and `docx4j_py/wml/builders.py` goes through the check, and a new one fails until it does.

### 4.4 Which limit (D4, Q1)

**The content API's setters, which use Office JS's vocabulary, refuse where Office JS refuses** (section 3.2):
- space before and after, and line spacing, from 1,638 pt and for a negative;
- a level's text indent beyond ±1,638 pt;
- a font size above 1,638 pt or at 0.

**Everywhere else, including the indents and column widths that Office JS lets wrap, every setter refuses beyond the faithful range of section 3.1, per attribute.** That's a guard where Office JS has none. The facade and the builders use the faithful range throughout, since they don't speak Office JS.

The faithful range of section 3.1, per attribute. That's Word's storage width, narrowed where Word clamps: the tab stop's ±31,680, the page width's 144 to 31,681, the page height's minimum of 144. Word's other interface limits aren't enforced by Word itself (section 3.1), so refusing at them would refuse values Word reads as written.

"Unsigned" means Word's storage (section 3.1's unsigned rows). For every attribute round 5 measured, that matches the schema, with one exception: `w:trHeight/@w:val` is `ST_TwipsMeasure` (unsigned) in the schema, but signed in Word, with the sign as the rule (-1 → exact 1). `w:pgMar/@w:top` is signed in both (`ST_SignedTwipsMeasure`), and a negative top margin is Word's exact top margin, so a negative there is accepted. A negative in an attribute Word stores unsigned, or in `w:trHeight/@w:val` (where `w:hRule` says exact), is refused as a caller's mistake (Q7). Word keeps `w:before="-1"` as 0xFFFF, the same stored value as 65,535, and writes it back as "-1"; what it shows on screen is unmeasured. docx4j CR-066 only reports these, which is the read side's position; this CR's is the write side's.

### 4.5 The read side is CR-020's (D5)

A value already out of range in a file, or assigned straight to a field, is reported by CR-020 Phase B's validator with this CR's table. It's never rewritten: CR-021 Q7 recommends against emulating the wrap, which would swap one wrong number for another.

## 5. Tests and oracles

- **Round 5's cases as fixtures** (`tests/fixtures/integer-limits/`: the six groups and Word's saves, read-only from `vbShares/Office 2016/__tmp/cr066/`). For each measured attribute, both ends of the faithful range are accepted and one past each is refused, through every setter that writes that attribute. The test's expectations come from Word's saves.
- **The census of D3.**
- **The examples of section 2.1** are refused, each with its code and hint.
- **A test that no 32-bit or 64-bit type is in the table:** the ids, `numId`, the EMUs.
- **No Word check:** a refusal writes nothing. The gates are the suite, `scripts/roundtrip.py` 50/50 and `scripts/compare_acceptance.py` 0 differ.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| 0 | Round 5, docx4j CR-066's phase 0, shared: built by docx4j-25, run by the user. **Done 2026-10-26** (section 3.1) | (docx4j-25's and the user's) |
| A | D1 to D4: the two tables, `measure_in_range`, every setter through it, the census, the fixtures | 1 day |
| B | The records: CHANGELOG ("Changed": a refusal of a value that was accepted), the README and `docs/pypi-readme.md` where they name the setters, CLAUDE.md, the index; CR-020 Phase B told the table is ready | ¼ to ½ day |

## 7. Risks

- **A wrong table refuses a legitimate value.** Only round 5's measured types have rows, and section 5's last test keeps the wide types out.
- **A user-interface limit (Q1) refuses a value Word reads faithfully,** for example a page wider than 22 in that Word renders. That's why the storage limit is the recommendation.
- **Code that set an out-of-range value on purpose breaks.** No sensible measurement is that large, and the CHANGELOG says so.

## 8. Questions (decided 2026-10-26)

The user, 2026-10-26: "CR-022 questions take the recommended answers". So:
- the content API refuses where Office JS refuses, at Office JS's limits, and every setter refuses beyond the faithful range per attribute elsewhere (Q1);
- the error is `MeasureRangeError(ContentError, ValueError)` with code `value.out_of_range` (Q2);
- only the measured attributes are checked, with `w:after` and the other border edges by analogy (Q3);
- the facade (Q4) and the builders (Q5) refuse too;
- the hint is in the caller's units (Q6);
- a negative in a Word-unsigned attribute, or in `w:trHeight/@w:val`, is refused (Q7).

| Q | Question | Recommendation |
|---|---|---|
| Q1 | Refuse beyond the faithful range per attribute (section 3.1: the storage width, narrowed where Word clamps), or at Word's interface limits (22 in for a page, 1,638 pt for a font)? | Revised after check 71 (section 3.2). The content API refuses where Office JS refuses (spacing, line spacing, a level's text indent, font size, at Office JS's limits), and at the faithful range elsewhere. The facade and the builders use the faithful range. Round 5 shows Word's file reader keeps values past most interface limits, so only the API Office JS defines uses them. |
| Q2 | The error: a new `MeasureRangeError(ContentError, ValueError)` with code `value.out_of_range`? | Yes. |
| Q3 | Which attributes: only those round 5 measured (with `w:spacing/@w:after` and the other border edges by analogy), or every integer attribute by an assumed width? | Only those measured, plus the two analogies, which docx4j-25 also assumes. An attribute without a row isn't checked. |
| Q4 | The python-docx facade refuses too, though python-docx itself accepts any `Length`? | Yes. python-docx writes the same broken file; the facade's promise is python-docx's names and semantics, and a refusal is a better semantics for a value Word misreads. |
| Q5 | The builders (`p(..., size=...)`, `r(..., size=...)`) refuse too? | Yes, through the same check. |
| Q6 | The hint in the caller's units (points, cm, `Length`) as well as twips? | Yes: the caller set points or a `Length`, and twips alone would not tell them what to change. |
| Q7 | A negative in an attribute Word stores unsigned (`w:before`, `w:sz`, the text scale, the page size), or in `w:trHeight/@w:val`: refuse it at the setters, or allow it, since Word accepts it and keeps it as a large positive value or an exact height (docx4j only reports it)? | Refuse it. No caller setting a negative space before or font size means a large positive one, and a row's exact height is `w:hRule`'s to say. A negative in a signed attribute (`w:pgMar/@w:top`, the indents) stays accepted. |

## 9. Effort

1 to 1½ days here after round 5: A 1, B ¼ to ½.

## 10. Phase A implementation notes (2026-10-26)

Phase A as built, D1 to D4 and section 5's tests, at the coordinator's brief after the decision of section 8. Phase B (the records) is next and isn't started: no CHANGELOG, README, `docs/pypi-readme.md`, CLAUDE.md or index change here. docx4j CR-066 (sections 3.1 and 3.4) and docx4j-core-ts CR-002 section 65 (`fd898b9`: `wordRange.mts`, `measurements.test.mjs`, check 71's run) were read for behaviour only.

**Commits.**
- `e46f3c1`: the fixtures, `tests/fixtures/integer-limits/` (10.2).
- `ac288fa`: D1 and D2: the limits table, `values.measure_in_range`, `MeasureRangeError`, and the table's tests (`test_integer_limits.py`). Nothing calls the check yet.
- `3f66840`: D3 and D4: every setter through the check, the setter tests, the census (`test_measurement_setters.py`, `test_measurement_writes.py`), `tests/README.md`'s rows.
- the commit after `3f66840`: this section.

Two commits by other sessions landed on `main` between the decision (`9ef7507`) and `e46f3c1`, both documents only (`31d548b`, CR-023's close; `a1e8abc`, CR-020 section 4.7), so the before-runs below, taken at `9ef7507`, are the code `e46f3c1` started from.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the table (D1) | `docx4j_py/resources/integer-limits.json` (`resources` is in `KEEP`) | hand-written, 60 rows, 82 (attribute, context) keys once a row naming several attributes is expanded, 7 unmeasured attributes (10.3) |
| the check (D2) | `docx4j_py/model/values.py` | `measure_in_range(attribute, value, units=POINTS, *, label, context=None, limit=None)`; `integer_limit(attribute, context)`, `integer_limits()`, `unmeasured_attributes()`; `IntegerLimit`, `Units` (`POINTS`, `EMU`, `LINES`, `OWN`), `CallerLimit` (10.4) |
| the error | `docx4j_py/model/content/errors.py`, exported from `docx4j_py.model.content` | `MeasureRangeError(ContentError, ValueError)`, `default_code = "value.out_of_range"`, with `attribute`, `value` and `stored` |
| the setters (D3, D4) | `model/content/paragraph.py`, `table.py`, `lists.py`, `font.py`; `docx/section.py`, `text.py`, `table.py`, `_base.py`; `wml/builders.py` | 10.5 |
| the census | `tests/content/test_measurement_writes.py` | 92 writes of an integer attribute, 30 through the check, 62 listed with a reason (10.6) |
| the oracle | `tests/content/test_integer_limits.py`, `tests/content/test_measurement_setters.py` | 10.7 |

### 10.2 The fixtures

`tests/fixtures/integer-limits/`, copied read-only from `vbShares/Office 2016/__tmp/cr066/` on 2026-10-26 and checked against the share file by file (SHA-1): `round5/` holds round 5's six group files (`cr066-{B,N,P,R,S,T}.docx`), Word 365's save of each, `cases.txt` (191 cases, the control `Z-1` among them) and `README.txt`; `round6/` holds round 6's eight group files (`cr066-r6-{A,E,F,N,P,R,S,T}.docx`), their saves, the twenty `D1` and `D2` singles (`w:defaultTabStop` and `w:hyphenationZone`, one document-wide value per file) and their saves, `cases.txt` (489 cases) and `README.txt`. No `~$` lock file was there. Not copied: the families and the other singles (fallbacks no round needed, since Word opened every group) and `families.txt` (its content is in `cases.txt`). `README.md` names the sources (docx4j CR-066 phase 0, rounds 5 and 6, the user's Word 365 runs of 2026-10-26, docx4j-25's generator, round 5 at `1be8b6515`) and how a case is laid out; `SHA1SUMS` holds the 72 files.

### 10.3 The table as built (D1)

Each row gives its attribute paths, a context where the range depends on a sibling attribute (`null` for that attribute absent), the faithful range (inclusive, in the row's unit), Word's storage as found, and the case ids it rests on. A path names a holder where the element name alone isn't enough: `w:rPr/w:sz` (`w:frame/w:sz` is text), and the two cell-margin rows, which round 6 measured on one edge each. Two optional fields record how Word's save spells a reading: `word_writes` (65535 saved as `"-1"`, in `w:before`, `w:after` and the five unsigned margins) and `sentinel` (`w:col/@w:space`, where every negative becomes -1).

| Attribute (context) | Faithful range | Cases |
|---|---|---|
| `w:ind/@w:left`, `@w:start`, `@w:right`, `@w:end` | -32,768 to 32,767 twips | r5 P2; r6 A6, A5, A7 |
| `w:ind/@w:firstLine` | 0 to 32,767 twips | r6 A1 |
| `w:ind/@w:hanging` | 0 to 32,768 twips | r6 A2 |
| `w:spacing/@w:before`, `@w:after` | 0 to 65,535 twips | r5 P1, N1; r6 A3 |
| `w:spacing/@w:line` (`lineRule` auto, absent) | -32,768 to 32,767 240ths of a line | r5 P3 |
| `w:spacing/@w:line` (exact) | **1 to 32,768** twips (departure 1) | r5 P4 |
| `w:spacing/@w:line` (atLeast) | 0 to 32,767 twips | r6 A4 |
| `w:spacing/@w:val` (a run's) | -32,768 to 32,767 twips | r5 R3 |
| `w:position/@w:val` | -32,768 to 32,767 half-points | r5 R2 |
| `w:rPr/w:sz/@w:val`, `w:szCs/@w:val`, `w:kern/@w:val` | 0 to 65,535 half-points | r5 R1, N1; r6 R2, R3 |
| `w:w/@w:val` (text scale) | 0 to 65,535 percent | r5 R4, N1 |
| `w:tblW/@w:w`, `w:tcW/@w:w` (`type` dxa; pct) | -32,768 to 32,767 twips; fiftieths of a percent | r5 T1, T2; T5, T4 |
| `w:tblInd/@w:w`, `w:tblCellSpacing/@w:w`, `w:tblCellMar/w:left/@w:w`, `w:tcMar/w:top/@w:w` (dxa) | -32,768 to 32,767 twips | r5 T3; r6 T4, T5, T6 |
| `w:trHeight/@w:val` (`hRule` atLeast, absent) | 0 to 32,767 twips | r5 T6, N1 |
| `w:trHeight/@w:val` (exact) | **1 to 32,768** twips (departure 1) | r6 T1 |
| `w:tab/@w:pos` | -31,680 to 31,680 twips | r5 P5 |
| `w:pgSz/@w:w`; `@w:h` | 144 to 31,681; 144 to 65,535 twips | r5 S1, N1; S2 |
| `w:pgMar/@w:top`, `@w:bottom` | -32,768 to 32,767 twips | r5 S3, N1; r6 S6 |
| `w:pgMar/@w:left`, `@w:right`, `@w:header`, `@w:footer`, `@w:gutter` | 0 to 65,535 twips | r6 S1 to S5 |
| `w:pgNumType/@w:start` | -2,147,483,648 to 2,147,483,647 | r5 S4 |
| `w:cols/@w:space`; `w:col/@w:w`, `@w:space`; `w:lnNumType/@w:distance` | -32,768 to 32,767; 0 to 32,767; 0 to 32,767; -32,768 to 32,767 twips | r6 S7; S8, S9; S10 |
| `w:framePr/@w:w`, `@w:hSpace`, `@w:vSpace`, `@w:x`, `@w:y`; `@w:h` | -32,768 to 32,767; 0 to 32,767 twips | r6 F1, F3 to F6; F2 |
| `w:tblpPr`'s `leftFromText`, `rightFromText`, `topFromText`, `bottomFromText`, `tblpX`, `tblpY` | -32,768 to 32,767 twips | r6 P1 to P6 |
| `w:defaultTabStop/@w:val`, `w:hyphenationZone/@w:val` | 0 to 65,535 twips | r6 D1, D2 |
| `w:top/@w:sz`, `w:bdr/@w:sz`; `@w:space` | 0 to 255 eighths of a point; 0 to 31 points | r5 B1, B2, N1; r6 R4, R5, T9, T10, S11 |
| the other eleven border edges' `@w:sz`, `@w:space` (`left`, `bottom`, `right`, `start`, `end`, `between`, `bar`, `insideH`, `insideV`, `tl2br`, `tr2bl`) | as `w:top`'s | by analogy (Q3) |

That is 60 rows, every one of section 3.1's and 3.3's measurable attributes. **Unmeasured** (no row, so not checked; Q3): `w:trHeight/@w:val` under `hRule="auto"` (discarded at every value, T2), `w:gridCol/@w:w` (recomputed, T3), `w:wBefore` and `w:wAfter` (dropped, T7, T8), `w:legacy`'s two (dropped, N1) and `w:fitText/@w:val` (kept at every value, R1: no limit found). Every case of both rounds is cited once, by a row or by the unmeasured list, but for the two controls and round 6's group E, which is CR-021's: 612 cases by rows (191 inside their row's range, 421 outside), 58 by the unmeasured list.

Not in the table, by design: anything 32 bits or wider (the comment, bookmark and revision ids, `w:numId`, `w:abstractNumId`), and anything outside `w` (DrawingML's EMUs). Every row's range fits 16 bits but `w:pgNumType/@w:start`'s.

### 10.4 The check (D2)

`measure_in_range("w:ind/@w:left", points, label="left indent")` looks up the row (by attribute and context), converts the caller's value to the row's unit **with the conversion the setter always made** (`round(points * 20)`, python-docx's `Length.twips`, `round(Length.pt * 2)`, `twips_of(Emu(int(lines * Twips(240))))`, or the value itself for the builders' twips), and returns it if the row admits it and the caller's limit (a `CallerLimit`, in the caller's units) admits the caller's value; otherwise it raises `MeasureRangeError`. An attribute the table lists as unmeasured, or a context it lists (`w:trHeight` under auto), is converted and not checked. An attribute the table doesn't know raises `LookupError`: a setter's mistake, not a caller's. A NaN or an infinity is refused, `value.out_of_range` too. 1.4 µs a call.

The message is the CR's (Q6): `left indent 400000 pt (8,000,000 twips) is beyond Word's range, -1,638.4 to 1,638.35 pt`, exactly section 4.2's example with the decided range. Its hint is the range: `pass a left indent from -1,638.4 to 1,638.35 pt`. Where Office JS's limit binds, the range is the two ranges together, an open end said so, and the message says whose: `space before 5000 pt (100,000 twips) is beyond Word's range, 0 to 1,638 pt, 1,638 excluded, where Office JS's setter stops`. For a `Length` the value is shown in EMU with cm (pt for a size) and the attribute's unit, and the range in EMU with cm, each cm bound rounded inward: `page width 180000000 EMU (500 cm, 283,465 twips) is beyond Word's range, 91,440 to 20,117,435 EMU (0.254 to 55.881 cm)`. A line multiple shows lines, its bounds to three places inward (`-136.533 to 136.529 lines`).

### 10.5 The setters (D3, D4)

Each computes its value through the check **before** it creates or records anything (a `w:pPr`, a `w:ind`, a `w:pPrChange` under tracking, a split run, a `ChangeReport`), so a refusal leaves the tree as it was. Where the old code converted inside `with recording(...)` or after creating the element, the conversion moved up.

- **The content API, at Office JS's limits** (section 3.2, Q1): `Paragraph.space_before` and `space_after` (`OFFICE_JS_SPACING`: 0 up to, not including, 1,638 pt) and `line_spacing` (the same, with the exact row, so 0.05 pt up); `List.set_level_indents`' text indent (`OFFICE_JS_TEXT_INDENT`: above -1,638.4, below 1,638 pt); `Font.size` (`OFFICE_JS_FONT_SIZE`: above 0, up to 1,638 pt; 0 removes, departure 2). `TableCell.width` and `column_width` refuse a negative, as Office JS refuses it (`OFFICE_JS_CELL_WIDTH`; departure 3).
- **The content API, at the faithful range**, where Office JS lets the value wrap: `Paragraph.left_indent` and `right_indent` (by whichever of `w:left`/`w:start` or `w:right`/`w:end` the paragraph writes), `first_line_indent` (`w:firstLine` 0 to 1,638.35 pt, `w:hanging` to 1,638.4 pt), `TableCell.width` up to 1,638.35 pt, and `set_level_indents`' marker (the first-line field).
- **The python-docx facade, at the faithful range** (Q4), a `Length` in EMU: `Section.page_width`, `page_height` and the seven margins (`left_margin`, `right_margin`, `top_margin`, `bottom_margin`, `header_distance`, `footer_distance`, `gutter`); `ParagraphFormat.left_indent`, `right_indent`, `first_line_indent`, `space_before`, `space_after`, `line_spacing` (a `Length` under the rule the write leaves, exact or at least; a multiple under auto, in 240ths of a line); `TabStops.add_tab_stop` and `TabStop.position`; `Font.size` (half-points); `_Row.height` (under its row's rule: at least or absent, exact; not checked under auto, and a rule outside the enumeration read as unset, as CR-018 reads one); `_Cell.width`; `Table.add_column` and `add_table`'s and `_Cell.add_table`'s columns (`_base.column_widths`, before the table, or the paragraph `_Cell.add_table` puts in first, goes in).
- **The builders, at the faithful range** (Q5): `apply_run_options`' `size` (so `r(..., size=)`, `p(..., size=)` and everything over them), checked first, before any option is applied; `tc(width=)`; `tbl`'s column widths and their sum (`w:tblW`).
- `None` keeps its meaning everywhere: unset in the facade, left alone in `Font.size` and the builders. `w:gridCol` is written unchecked (`_Column.width`, `_set_column_widths`, `tbl`), with no row (10.3).

### 10.6 The census (D3)

`tests/content/test_measurement_writes.py` walks the AST of every module under `docx4j_py/model/content` and `docx4j_py/docx` and of `wml/builders.py`, `wml/pictures.py` and `wml/sdt.py` for every write of an integer attribute of `w`, `w14` or `w15`: an assignment to a field whose Python name is such an attribute's anywhere in the model (93 names; not to bare `self`, not of a constant), a `setattr` naming one or naming a variable, and a keyword to a model class or an `el` factory whose class types that keyword an integer attribute (the class resolved from the module's names). A write goes through the check when its value holds a call of `measure_in_range` or a name the same function bound from one, to a fixed point. Every other write is listed, keyed `<file>::<scope>::<the write>`; a new one fails, as CR-021's census does. 92 writes: 30 through the check, and 62 listed:
- 30 an integer attribute that is no measurement (ids, `w:numId`, `w:abstractNumId`, `w:ilvl`, a level's start, `w:outlineLvl`, `w:uiPriority`, `w:gridSpan`);
- 14 a field that is no integer attribute where it is written, only named like one elsewhere (`w:b`, toggles, children, content lists, `xml:space`, a relationship id, core properties);
- 5 a field copied as it stands (a deep copy, a split run, a restored `w:pPr`);
- 4 not a model field (a span's offsets, a class's methods);
- 3 `w:gridCol`, with no row;
- 3 a constant expression (a new level's default indents, a line multiple's 240ths);
- 2 a string value in a `w:val` (`w:compatSetting`, a tab's alignment);
- 1 a width computed from widths its caller checked (`_set_column_widths`' `w:tcW`).

CR-021's census (`test_integer_readers.py`) lists the check's own `int()` and `float()` calls: 131 listed now, 13 added (12 in `values.py`, a new reason `TABLE` for the table's own numbers; the facade's `first_line_indent`) and 2 gone (the facade's `Font.size` and `line_spacing` conversions, now the check's).

### 10.7 Tests and the numbers

- **`test_integer_limits.py`** (838 tests): the fixtures against `SHA1SUMS`; the rounds' case counts; every case cited once; each cited case of its row (element, attribute, the context the source states) and its value in the source; **each row against Word's saves**: for 58 rows with cases, 612 readings, every value inside the row's range kept (the value Word saves under the same rule, 65535 read through `word_writes`) and every value outside changed (wrapped, clamped, its rule flipped, its sentinel, or dropped), a boundary kept and one past it changed wherever the probe set them; the unmeasured attributes as Word treated them; each row's element against every class it binds in the model (`normalise._bound()` and `integer_attributes`, as CR-021's canon table reads the model): at least one class declares the attribute, every declaring class types it an integer of `w` and declares the row's context attribute, and a holder is needed for its typing or named in the row's note as the only one measured; the border rows on border classes only; the wide attributes absent and every row within 16 bits but `w:pgNumType/@w:start`. Fed the CR's own `0 to 32,767` for an exact line or an exact height, the test fails on P4-2 or T1-2; fed 31,680 for the page width, on S1-2.
- **`test_measurement_setters.py`** (37 tests): each setter of 10.5 at both ends of its range and one past each, the code, a `ValueError` and a `ContentError`, the hint naming the range, the tree and `package.changes` unchanged; `w:start`/`w:end` paragraphs; the first-line field; under change tracking, no `w:pPrChange`; a range's font splitting no run when refused; NaN and infinity; the facade's line spacing under each rule, a row's height under each rule (and an unknown rule), `add_table`'s and `_Cell.add_table`'s columns refused before anything goes in; the builders; **section 2.1's four examples**, refused with the CR's message.
- **`test_measurement_writes.py`** (5 tests): the census, its reasons, the section 2.2 setters found through the check, and the scanner's own rules (a bound name; a raw write to a field, to an `el` keyword and by `setattr` caught, an enumeration's `w:val` not).

### 10.8 Departures, with why

1. **An exact line and an exact row height are 1 to 32,768 twips, where sections 3.1 and 3.3 give 0 to 32,767.** Word stores an exact value as its negative, so the stored sign gives the rule back. Exact 32,768 is stored as -32,768 and read as exact 32,768: Word's saves keep it (P4-2; T1-2), and exact -32,768 comes back as exact 32,768 too (P4-7, T1-6). Exact 0 is stored as 0, as exact 65,536 is, which Word saves as `atLeast` 0 (P4-4) or drops (T1-4). The CR's bound would refuse a value Word keeps and accept one it changes, and section 5's test, a boundary kept and one past it changed, can't hold with it (10.7). docx4j-core-ts's `spacingLineExact` is 1 to 32,768 too. So `Paragraph.line_spacing` refuses below 0.05 pt (0 included), the facade's exact `Twips(0)` is refused and `Twips(32768)` accepted, and a row's exact height is 1 to 32,768. At least, the ranges are the CR's, 0 to 32,767.
2. **`Font.size = 0` keeps removing `w:sz` and `w:szCs`,** where Office JS refuses 0 (check 71, case 95). The removal has been documented on `Font.size` since CR-003 Phase B ("Setting `0` removes `w:sz` and `w:szCs`, which is what writing back what a silent run reads means": a run with no size reads 0, so writing back what was read must not ask for a 0 pt font), and `test_font.py::test_the_empty_values_remove_the_element` holds it. Following Office JS would break both, so 0 never reaches the check; every other value is held to Office JS's limit, above 0 up to 1,638 pt, and a negative is refused. The builders' `size=0` removes too, the same mapping (`apply_run_options`). **For the user to decide** (10.11).
3. **`TableCell.width` refuses a negative.** D4's list of Office JS's refusals doesn't name it, but section 3.2 records it ("Only a negative `columnWidth` is refused"), and Q1 is to refuse where Office JS refuses. The faithful row alone (signed) would accept it. The facade's `_Cell.width` and the builders' `tc(width=)` accept a negative, as the faithful range does (Q4, Q5).
4. **The content API's Office JS limits apply with the faithful range, not instead of it.** The two agree except at an exact line's 0 (departure 1), which Office JS was not asked; check 71 never set `lineSpacing` to 0. The setter refuses there.
5. **Office JS's limits are on the caller's value, before rounding**, as Office JS's are stated: 1,637.99 pt of space before is accepted and written as 32,760 twips, the twips 1,638 pt would write; 1,638 is refused. Where exactly Office JS stops between the values check 71 set (1,584 and 1,638 pt for spacing; 1 to 1,638 for a font) is unmeasured (core-ts's D2 asks for a check 71b); this follows the CR's statement of them.
6. **Two fields the table carries for the readings, not for the check:** `word_writes` and `sentinel` (10.3). Without them the fixture test would read Word's `"-1"` for 65,535 as a change, and `w:col/@w:space`'s `-1` as a kept -1.
7. **The cell-margin rows are the measured edges only** (`w:tblCellMar/w:left`, `w:tcMar/w:top`), though section 3.3's table names the containers: Q3 takes only the border edges by analogy. The other margin edges have no row.
8. **`add_table`'s column widths are shown in twips**, not in the caller's `Length`: they're the facade's own division of it (`_base.column_widths`), and a column's twips are what is refused.
9. **`_Row.height` under `hRule="auto"` is not checked**, by the table's unmeasured entry: Word discards it at every value (T2). A height set under auto and a rule changed afterwards is 10.10 item 1.

### 10.9 Existing semantics kept

- `wml/builders.py`'s `size=0` removes `w:sz` and `w:szCs`: not a measurement written, so not checked. Kept.
- `None`: unset in the facade (`page_width = None` and the margins, indents, spacing, line spacing, row height and cell width remove the attribute or the element, as before); left alone by `Font.size` and the builders. Kept, and tested for the section, the paragraph format and the font.
- `Font.size = 0` removes the size (departure 2). The content API has no other documented meaning for a value Office JS refuses: `line_spacing = 0` was undocumented (the getter's "0 when the paragraph sets none" is a reading), and it wrote `w:line="0" w:lineRule="exact"`, which Word reads as at least 0.
- **No pre-existing test's expectation changed.** No test asserted a value now refused; the only edits to existing tests are CR-021's census list (10.6) and `tests/README.md`'s rows.

### 10.10 Found, and not done

1. **The rule setters can move a value out of its range without a check.** The facade's `line_spacing_rule` changes `w:lineRule` under an existing `w:line` (an auto line of -100 set to `EXACTLY` is exact -100, which Word reads as at least 100), and `_Row.height_rule` changes `w:hRule` under an existing height (32,768 at least is refused, but 32,768 exact set to at least is accepted and read as exact). They write the rule, not the measurement, so the census doesn't see them; they could check the existing value under the new rule.
2. **`insert_table` (and the Markdown importer's tables) over a section whose page width is already beyond Word's range** now refuse, through `tbl`'s check (`table width ... twips`), where before they wrote a width Word wraps. Only a file's own out-of-range `w:pgSz` reaches it (CR-020 Phase B's to report); `writable_width` could clamp, but CR-021 Q7 is against emulating Word's reading.
3. **Writers outside the census's scope**: the Markdown importer's nested-quote indent grows 720 twips a level, beyond 32,767 at 46 levels (`model/markdown/importer.py`, `context_p_pr`); the revisions engine's grid and width merges (`model/revisions/tables.py`) sum widths already in the tree. Neither takes a caller's measurement. The scanner run over `model/markdown`, `model/customxml` and `model/revisions` finds 29 writes, none through the check.
4. **`api_script`** writes a script from a document's values, so a file already out of range gives a script that now raises when it is run. It reads; it doesn't write.
5. **CR-020 Phase B's finding** (its section 4.7, agreed with core-ts's CR-017 the same day) wants `storage` as a token (`"int16"`, `"uint16"`, `"8 bits"`, `"5 bits"`, `"int32"`) and `word_reads`, what Word reads. The table's `storage` is prose beginning with the token in most rows (not `w:firstLine`'s, `w:hanging`'s or `w:framePr/@w:h`'s), and there is no model of the wrap. Both are Phase B's of CR-020 to add, or a field here.
6. **Differences from docx4j-core-ts's setters** (`wordRange.mts` at `fd898b9`), where this follows the CR: core-ts holds `spaceBefore`, `spaceAfter`, `font.size` and the list text indent at Word's storage, pending its D2, where these are at Office JS's limits (Q1); core-ts stops `w:before` and `w:after` at 65,534 because Word writes 65,535 back as `"-1"`, where the CR's range is 0 to 65,535 and these only matter in the facade (the content API stops at 1,638 pt); core-ts's message also says what Word would read (`Word 365 would read it as 4608`), which needs the wrap model of item 5; core-ts has a percentage cell width, which this content API doesn't write. The exact line agrees (1 to 32,768), as do the indents, the first-line field and a cell's width.

### 10.11 For the user to decide

1. **`Font.size = 0`** (departure 2): keep the documented removal (as built), or follow Office JS and refuse 0, changing CR-003's documented behaviour and `test_the_empty_values_remove_the_element`.
2. **Departure 1**: the exact rows at 1 to 32,768, from Word's saves, and with them `Paragraph.line_spacing = 0` refused.
3. **65,535 in the unsigned twips attributes** (`w:before`, `w:after`, the five unsigned margins): accepted, as the CR's range says, though Word's save writes it back as `"-1"`, which another reader (and CR-020's report, unless it reads `word_writes`) takes for -1; core-ts stops at 65,534. Only the facade and the builders reach it.
4. **10.10 item 1**: whether the facade's rule setters should check the value they leave under the new rule.

### 10.12 The gates

| gate | before (`9ef7507`) | after (`3f66840`) |
|---|---|---|
| the full suite (`.venv-fork/bin/python -m pytest`) | 6,770 passed, 8 skipped (derived: 7,650 less the 880 new tests) | **7,650 passed, 8 skipped** (4 min 47 s) |
| the fast suite (`-m "not slow"`) | 5,389 passed, 8 skipped, 1,381 deselected: the existing tests, run once the setters had changed and before any new test was written; none failed but CR-021's census, which then listed the check's conversions | 6,269 passed, 8 skipped, 1,381 deselected (3 min 16 s) |
| the new tests | --- | 838 in `test_integer_limits.py`, 37 in `test_measurement_setters.py`, 5 in `test_measurement_writes.py` |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | --- | **50 of 50** canonically identical, 0 differences, 0 skipped, 34 `boolean-spelling` |
| acceptance (`scripts/compare_acceptance.py`, `out/acceptance/` regenerated at `9ef7507` in a scratch worktree against the checkout's) | --- | **17 artefacts, 246 entries, 0 differ** beyond the dates: no acceptance document sets an out-of-range value |
| `codegen/generate.sh --check` (a scratch worktree at `3f66840`) | --- | two generations byte identical, 308 modules; nothing generated differs from the commit, and `resources/integer-limits.json` survives `clean.py` |
| `ac288fa` on its own (a scratch worktree) | --- | `tests/content/`, `test_builders.py`, `test_builders_phase_a.py`: 3,449 passed, 3 skipped (the worktree's skips) |
| the censuses | --- | 92 writes, 30 through the check, 62 listed; CR-021's 131 listed |

No Word check (section 5): a refusal writes nothing, and the acceptance documents are unchanged.
