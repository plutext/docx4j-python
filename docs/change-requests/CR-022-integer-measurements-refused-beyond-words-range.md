# CR-022: Integer measurements refused beyond Word's range as code sets them --- the write-side guard, after Word's 16-bit wrap

**Status:** Proposed 2026-10-26, at the user's request ("option 1, and also add the out-of-range report to CR-020's Phase B list"). It follows the user's suspicion the same day "that we should guard against users setting twip measurements to large integers in each of docx4j, python and core-ts". The portfolio session coordinates it as one CR per port. Round 5 of the shared Word probe is the shared phase 0. It is docx4j CR-066's phase 0 (docx4j-25, `1be8b6515`, local), on the share at `vbShares/Office 2016/__tmp/cr066/`: 191 cases in six groups, 22 single-attribute families and 48 singles. Checked here: all 76 files load with nothing kept or skipped. **Round 5 was run 2026-10-26** (the user, Word 365: all six groups opened clean, no repair). **Phase 0 is done.** The range is per attribute, not per schema type (section 3.1), so D1 is a table per attribute, and Q1 is restated. **Nothing has started.** Phase 0 gives the limits, and nothing here is built before it. Three phases (section 6), 1 to 1½ days after round 5. Seven questions wait on the user (section 8).

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

## 8. Questions

| Q | Question | Recommendation |
|---|---|---|
| Q1 | Refuse beyond the faithful range per attribute (section 3.1: the storage width, narrowed where Word clamps), or at Word's interface limits (22 in for a page, 1,638 pt for a font)? | The faithful range. Round 5 shows Word keeps values past most interface limits, so refusing there would refuse what Word reads as written. |
| Q2 | The error: a new `MeasureRangeError(ContentError, ValueError)` with code `value.out_of_range`? | Yes. |
| Q3 | Which attributes: only those round 5 measured (with `w:spacing/@w:after` and the other border edges by analogy), or every integer attribute by an assumed width? | Only those measured, plus the two analogies, which docx4j-25 also assumes. An attribute without a row isn't checked. |
| Q4 | The python-docx facade refuses too, though python-docx itself accepts any `Length`? | Yes. python-docx writes the same broken file; the facade's promise is python-docx's names and semantics, and a refusal is a better semantics for a value Word misreads. |
| Q5 | The builders (`p(..., size=...)`, `r(..., size=...)`) refuse too? | Yes, through the same check. |
| Q6 | The hint in the caller's units (points, cm, `Length`) as well as twips? | Yes: the caller set points or a `Length`, and twips alone would not tell them what to change. |
| Q7 | A negative in an attribute Word stores unsigned (`w:before`, `w:sz`, the text scale, the page size), or in `w:trHeight/@w:val`: refuse it at the setters, or allow it, since Word accepts it and keeps it as a large positive value or an exact height (docx4j only reports it)? | Refuse it. No caller setting a negative space before or font size means a large positive one, and a row's exact height is `w:hRule`'s to say. A negative in a signed attribute (`w:pgMar/@w:top`, the indents) stays accepted. |

## 9. Effort

1 to 1½ days here after round 5: A 1, B ¼ to ½.
