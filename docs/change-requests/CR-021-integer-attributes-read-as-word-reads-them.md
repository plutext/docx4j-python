# CR-021: Integer attributes read as Word reads them --- a decimal truncated toward zero, the empty string 0, every other form refused; the port of docx4j CR-065's rule

**Status:** Proposed 2026-10-26, at the user's request ("yes please draft the CR"). The portfolio session had relayed the user's request the same day: propose a fix for a decimal in a WordprocessingML integer attribute, which docx4j repairs on load and this package saves back. The rule comes from docx4j CR-065's phase 0: one Word probe shared by the three ports, 65 cases over four rounds, run by the user in Word 365 on 2026-10-26 and checked here against Word's own saves (section 3). Four phases (section 6): 0, the fork's stage 14 (Q2); A, the reader; B, the save; C, the records and a Word check. 2 to 3 days here, plus half a day in the fork. **Decided 2026-10-26** (the user: "CR-021 Q1 to Q6 and Q8 as recommended; Q7 discussed separately"; section 8). Q7 was discussed separately. Its outcome so far is CR-022, the write-side guard, and an out-of-range report added to CR-020 Phase B. This CR doesn't emulate the wrap, as recommended, unless that discussion decides otherwise. **Started 2026-10-26** (the user: "do CR-021"). **Phases 0, A and B are done (2026-10-26; section 10).** Phase 0 is the fork's stage 14 (item 32), released as docx4j-xsdata 26.2.4 on the user's go, with `int` now in `runtime.EXACT_TYPES`; the commits here are `2d7e46c`, `e158f2f`, `406d98d`, `44e2888` and `7fe0be8`. Each was verified by hand (section 10.12). **Phase C is done:** the records, and the Word check, which passed 2026-10-26 (the user: "cr021: everything opened. cr021-edges gets normal margin settings"; section 10.13). **CR-021 is complete, 2026-10-26.** One unmeasured difference from docx4j remains: `"."`, `"+"` and `"-"` read 0 here and are dropped by docx4j. Word accepts all three, and in the check's files the two readings can't be told apart (section 10.13).

*The number:* CR-020 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057.

**Depends on:**
- **docx4j CR-065** phase 0 (done, `aacaa4c45` and the three commits before it) for the measurements, and its D3 for the rule, so that the ports read identically. Its phase 1 (a part that can't be read is never replaced by an empty one) and phase 2 (the lenient parse in Java) are docx4j's own.
- **The xsdata fork**, for phase 0 only: `int` added to `ParserConfig(exact_types=...)` as stage 14, released as docx4j-xsdata 26.2.4 (Q2). Without it, phases A to C still fix the decimals and the empty string, but not the forms of section 2.2.
- **Counterparts:** docx4j-generated-objects-ts CR-015, revised the same day to the same rule.

## 1. Why

Real producers write decimals where WordprocessingML wants an integer, and Word opens the files:
- Google Docs, 2014 and 2015: `<w:tblW w:w="10206.0" w:type="dxa"/>`, `<w:pgSz w:h="16839.0" w:w="11907.0"/>`;
- pandoc 2.2.1, 2018, in `tblW` (docx4j issue #298);
- the `w:hanging="141.99999999999994"` family, seen in 2017 and 2025.

SQL Server Reporting Services 2012 writes empty values: `w:pgMar w:header="" w:footer="" w:gutter=""` and `w:rsidR=""`.

This package keeps such a value as its text. A save after an edit writes it back exactly, so the output is as broken as the input. That runs against the user's rule for the normaliser (2026-10-24): "A user who edits a file then blames us is a real concern: they are entitled to assume we don't write dodgy content." Its readers disagree with each other about the value, and one of them raises (section 2.3).

## 2. What exists here (measured 2026-10-26 at `a09c1a0`)

### 2.1 The parse keeps a decimal as its text

The probe: one document holding every shape above, five integer attributes docx4j doesn't round (section 3.2), and SSRS's empties. A decimal or an empty string fails the fork's `int` converter:
- a `ConverterWarning`, and the field holds the source string;
- since the fork's stage 13 (26.2.3, CR-018) it's listed in `part.kept` and `package.kept` as `KeptValue(reason="invalid")`, with the field, the class and the path: 23 of 23 in the probe, SSRS's three empty `pgMar` values included;
- a save with the part never read is byte for byte;
- a save after an edit writes every value back as read. The normaliser rewrites none of them (`last_save.normalised` is `[]`), and `normalise=False` gives the same bytes.

### 2.2 The parse also accepts forms Word refuses

The fork's `int` converter is Python's `int()`, so some strings bind as integers without a warning or a kept record. A save then writes the integer:

| Source | Binds as | Kept | Word (section 3) |
|---|---|---|---|
| `" 240 "`, `"\t240\n"` | 240 | no | refuses the document |
| `"1_000"` | 1000 | no | refuses (not a form Word reads) |
| `"٢٤٠"` (Arabic-Indic digits) | 240 | no | refuses (not a form Word reads) |
| `"+240"`, `"0240"` | 240 | no | reads 240 |

So this package reads, and quietly repairs, documents Word won't open. That's more lenient than Word and than docx4j's D3, and nothing records it. CR-018 met the same thing for on/off values and enumerations, and fixed it in the fork: the parse accepts exactly a lexical form and keeps anything else as its text (`exact_types`). `int` isn't one of those types.

### 2.3 The readers disagree, and one raises

Three readings of one field:
- **The content API reads a float.** `Paragraph`'s readers use `float(value) / TWIPS_PER_POINT` (`model/content/paragraph.py`, `table.py`, `api_script.py`: about 17 sites). "720.5" reads 36.025 pt, and "2.5" reads 0.125 pt.
- **The python-docx facade reads an `int` and raises.** Its readers use `int()` (`docx/section.py`, `table.py`, `text.py`, `styles.py`, `shared.py`: about 23 sites). `ValueError` comes from `section.page_width` and `page_height` on "11907.0" (Google Docs' shape), from `header_distance`, `footer_distance` and `gutter` on "" (SSRS's), and from table column widths on "4500.5".
- **Others have their own helpers:** the numbering emulator's `_as_int` and `definitions._int`. `revisions/tables.py:796` adds two cell widths with `int()` and `model/content/table.py:446` reads the grid with `int()`, so both raise on a decimal.

### 2.4 The model's integer attributes

- **WordprocessingML (`w`, `w14`, `w15`):** 224 integer-only attribute fields in `w`, under 70 distinct names, and none is a union with `str`. `w14` has 61 more and `w15` one.
- **Elsewhere:** 539 more integer-only attribute fields (DrawingML 279, `oart` 126, VML 108, `word` 13, docProps 7, OMML 5, InkML 1), and 31 unions of `int` with `str` (DrawingML 27, `oart` 3, `word` 2, and OMML's `CTTwipsMeasure.val`). In a union, a decimal binds as a `str` without a warning.

### 2.5 The normaliser, and the canon's mirror of it

CR-018's on/off rows walk the typed tree. `on_off_attributes(cls)` gives a class's `bool` attribute vars, cached. `_changes(root)` finds a `str` held in a `bool` field (a kept value) and its normal form. `_applied` sets each such field to its normal form for the render and restores it afterwards, so the tree keeps what was read. Integers fit the same three functions.

`scripts/canon.py`'s `compare(normalise=True)` applies the same normal form to the source with an lxml mirror, `_lxml_table()`. It holds an (element, attribute) pair only where every class that element name binds types the attribute the same way. Built the same way for integers over `w`, that gives 193 unambiguous pairs and 28 ambiguous ones. The ambiguous ones are border edges and cell margins sharing element names: `w:bottom@w:sz` is a border's, and `w:bottom` in `w:tblCellMar` has no `sz`.

### 2.6 The empty rsid

An rsid is `ST_LongHexNumber`, an `xsd:hexBinary`. An empty `hexBinary` is lexically valid, so `rsidR=""` parses without a warning or a kept record, and a save writes it back as `""`.

## 3. What Word does, and docx4j

### 3.1 Word 365's reading (docx4j CR-065 phase 0, the user's run, 2026-10-26)

Four rounds, 65 cases, in families: `w:spacing/@w:before` in every form; `w:ind/@w:left` with positive and negative halves; docx4j's listed attributes; borders; tabs; run properties; tables; sections; and `styles.xml`, `settings.xml` and a header. The files and Word's saves are at `vbShares/Office 2016/__tmp/cr065/`.

**Word's lexical form for an integer attribute** is an optional sign (`+` or `-`), digits, and optionally a dot and digits, where either digit group may be empty; or the empty string. As a regular expression it's `[+-]?[0-9]*(\.[0-9]*)?`, non-empty, or `""`. There's no surrounding whitespace: XML Schema's whitespace collapse doesn't apply here, unlike the on/off reading CR-018 measured. **The fraction is dropped: a decimal is truncated toward zero. The empty string reads 0. Any other form makes Word refuse the whole document,** not just the attribute: a comma (`"240,5"`), hex (`"0x10"`), an exponent (`"1.5e2"`, `"1e3"`), letters (`"abc"`), or surrounding spaces (`" 240.5 "`, `" 240 "`).

Checked here in Word's saves of families A1, A2a to A2d, A6, B, C and H:

| Read | Word wrote |
|---|---|
| 120.5, 120.4 / 121.5 / 122.5 / 240.0 / 141.99999999999994 | 120 / 121 / 122 / 240 / 141 |
| `w:ind/@w:left` 720.5, 720.4 / 721.5 / -120.5 / -121.5 / -1.5 / -0.5 | 720 / 721 / -120 / -121 / -1 / no `w:ind` (0, the default) |
| `w:spacing/@w:after` 141.5; `@w:line` 259.2 (`lineRule="auto"`); `w:hanging` 141.99999999999994 | 141; 259; 141 |
| `w:pgSz` 11907.5 by 16838.5; `w:pgMar/@w:top` 1133.5 | 11907 by 16838; 1133 |
| SSRS: `pgMar header="" footer="" gutter=""`; `sectPr rsidR="" rsidSect=""`; a lone `top=""` | 0, 0, 0; the empty rsids dropped; 0 |
| `"240."`, `".5"`, `"+240"`, `"+240.5"` | 240, no `w:spacing` (0), 240, 240 |
| `" 240.5 "`, `" 240 "`, `"240,5"`, `"0x10"`, `"1.5e2"`, `"1e3"`, `"abc"` | the document refused |

docx4j-25 reports the rest, also measured:
- border `@w:sz` 7.5 → 7, `@w:space` 1.8 → 1 and 7.2 → 7, `w:between` too;
- tab `@w:pos` 1440.5 → 1440, text scale 90.5 → 90, `w:rPr/w:spacing` 20.5 → 20, `w:position` 6.5 → 6;
- `tblW` 9000.5 → 9000 and 10206.0 → 10206, `tblInd` 10.5 → 10;
- `trHeight` 400.7 → 400, and the `pct` widths 50.5 → 50;
- a style's `w:ind/@w:left` 720.5 → 720, `defaultTabStop` 720.5 → 720, a header's `w:before` 120.5 → 120.

Two asides from the run:
- **Defaults vanish in Word's saves.** A value equal to the style's default after truncation isn't written: `w:sz` 22.5 became 22, the document default. Word stores direct formatting as a difference from the style.
- **Word's twips values are 16-bit.** Sentinel integers left by a slip in the probe's generator came back modulo 2^16.

Word recomputes the table grid on open, so `gridCol` decimals have no reading you can observe.

### 3.2 docx4j today (measured here through JPype, 17.3.4 and a 17.3.5-SNAPSHOT build)

- **Its stylesheet rounds half-even** (`format-number(., '#')`): 2.5 → 2, 3.5 → 4, 4500.5 → 4500, 5705.5 → 5706. That's wrong for Word on every half that rounds up, and on `141.99999999999994` (docx4j gives 142, Word 141).
- **It covers only a fixed list.** A decimal anywhere else (`w:spacing/@w:before`, `w:ind/@w:left`, a border's `@w:sz`, a tab's `@w:pos`, a `pct` width) leaves the main document unmarshalled. An edit then starts an empty document, and the save loses all the content. docx4j-25 traced it to a `NumberFormatException` that docx4j's event handler never sees; CR-065's phase 1 fixes it.

**CR-065's D3 (a):** read Word's form exactly as Word reads it. Any other form, whitespace included, drops the attribute with a WARN, without trimming.

## 4. Design

### 4.1 One integer reader (D1)

`docx4j_py.model.values.integer(value) -> int | None`:
- an `int` is itself;
- a `str` in Word's form reads its integer part with its sign. The truncation is done on the string, not through `float`: `float("2.9999999999999999")` is 3.0, where the string's integer part is 2. Word's reading of more digits than a double holds is unmeasured: `141.99999999999994` reads 141 by either route. Phase C's Word check adds the case (Q8). `""`, and a form whose digit groups are both empty (`"."`, `"+"`, `"-"`), read 0, by the lexical test as CR-065 D3 states it; those are unmeasured. **docx4j differs here:** its phase 2 (`fe449c35c`, 2026-10-26) requires at least one digit, and drops `"-"`, `"."`, `"+"` and `"+."` with a WARN. It reads `"239.99999999999999"` as 239, as this package does, since neither parses through a double. Phase C's Word check measures Word's own reading (`cr021-edge-E2-raw.docx` to `E4`), and both ports follow it;
- anything else reads `None`, unset, because Word refuses it.

Digits are ASCII `[0-9]`. Python's `\d` matches every Unicode digit, so it isn't used. The 16-bit wrap isn't emulated (Q7).

### 4.2 Every reader goes through it (D2)

The content API's twips readers, the facade's, the numbering emulator's `_as_int` and `definitions._int`, `revisions/tables.py`'s width sum and `model/content/table.py`'s grid all call `values.integer`, and so does the resolver wherever it does arithmetic. So a read agrees with what a save writes, the facade stops raising, and "720.5" reads 36.0 pt, as Word shows it.

**A census holds the list whole**, the way `tests/openpackaging/test_private_readers.py` holds `read_contents`' callers: every `int(` or `float(` applied to a model field under `docx4j_py/model` and `docx4j_py/docx` is listed, and a new one fails until it's listed or goes through `values.integer`.

### 4.3 The save writes Word's integer (D3)

The normaliser gains integer rows beside its on/off rows, in the same typed walk:
- `integer_attributes(cls)`: a class's attribute vars whose one type is `int`, cached per class;
- `_changes`: a `str` held in such a field gets its normal form, `values.integer(text)`, written as a decimal integer;
- a refused form (`None`) drops the attribute;
- each change is recorded in `last_save.normalised` with the source text, and `normalise=False` keeps the source spelling, as for CR-018's rows.

The canon's mirror gains the same rows. Its lxml table holds an (element, attribute) pair where every class that declares the attribute types it `int` (Q5). That settles 25 of section 2.5's 28 ambiguous pairs. The three left have another type in some holder: `w:sz@w:val` (`int` or `str`), `w:default@w:val` (`bool`, `int` or `str`) and `w:start@w:val` (`int`, or `ST_Border` on a border edge). The canon leaves those three to the typed path.

### 4.4 The exact parse (D4, phase 0, Q2)

`int` joins `bool` and `Enum` in the fork's `exact_types` (stage 14): the converter binds only `[+-]?[0-9]+`, exactly, and keeps anything else as its text with `reason="invalid"`. Section 2.2's forms are then recorded, and read as Word reads them: refused, so unset and dropped at save. That gives parity with docx4j's D3.

Without phase 0, D1 to D3 still fix the decimals and the empty string. The forms of section 2.2 would go on binding as integers and being repaired without a record.

### 4.5 Scope (D5)

- **The rows apply in `w`, `w14` and `w15`**, where Word was measured.
- **Elsewhere, today's output is kept exactly.** Phase 0's exact parse would otherwise turn a padded integer in a DrawingML attribute into kept text, which a save would write back as read. For any namespace outside those three, the normaliser writes what Python's `int()` would have bound where it parses, and leaves the value alone otherwise, so a DrawingML decimal is untouched, as today. A later probe can widen the scope (Q3).
- **The unions** of section 2.4. Stage 14 makes `int` exact wherever the int converter runs, unions included (the fork session's uniform rule, accepted 2026-10-26). In the 35 union fields that contain `int` (none in `w`, `w14` or `w15`), a value Python's `int()` accepts but the exact parse refuses would then bind as a `str` (31 fields with a `str` member), or be kept (4 without). So that today's output holds there too, the normaliser writes such a value as the integer Python's `int()` gives, wherever the field's types include `int` and its namespace isn't one of the three.

### 4.6 The empty rsid and SSRS (D6)

- **SSRS's empty `pgMar` values** are kept integers, so D1 and D3 cover them: written 0, as Word writes them. That differs from docx4j's stylesheet, which drops `gutter`; CR-065's D5 (a) follows Word.
- **An rsid attribute whose value is `""`** is dropped at save, as Word drops it (section 3.1), and recorded. It isn't a kept value (section 2.6), so this is the one row found by value, not by the parse.

### 4.7 What a save still doesn't do (D7)

- **A part never read is written back byte for byte**, so a value Word refuses survives only in a part nobody read. A save after an edit never writes one.
- **A value equal to its style's default** after truncation is written, where Word's save leaves it out. It means the same, and it's a tolerated difference from Word's save.

## 5. Tests and oracles

- **The probe as fixtures:** CR-065's probe files and Word's saves go into `tests/fixtures/integers/`, as CR-018's went into `onoff-fun/`. For every case, `values.integer` gives Word's value, or `None` where Word refused. For every family, the normalised save writes Word's value, allowing for the defaults Word leaves out (D7).
- **The exact parse** (after phase 0): each form of section 2.2 is kept, reads `None` and is dropped.
- **The census of D2**, and the facade's three `ValueError`s gone.
- **The canon:** `compare(normalise=True)` counts each rewrite as `word-normal-form` over the probe, and nothing else moves.
- **The gates:** `scripts/roundtrip.py` 50/50, `scripts/compare_acceptance.py` 0 differ, `codegen/generate.sh --check`.
- **A Word check in phase C, because saved output changes:** the normalised saves of the probe's families open without a repair and show Word's values. One case is a dropped required attribute (`pgMar top="abc"` dropped), so that Q4 is answered.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| 0 | The fork's stage 14 (`int` in `exact_types`) and the docx4j-xsdata 26.2.4 release, by the fork session; the pin here moves to `>=26.2.4,<26.3` (Q2) | ½ day in the fork |
| A | D1, D2: `values.integer`, every reader through it, the census, the fixtures | 1 day |
| B | D3, D5, D6: the normaliser's rows, the canon's, the empty rsid; D4's tests once phase 0 is released | ½ to 1 day |
| C | The records (CHANGELOG, CLAUDE.md, the index), and the Word check | ¼ day, and the user's Word run |

A and B don't wait for phase 0. Only the tests of section 2.2's forms do.

## 7. Risks

- **A reader missed.** The census of D2 holds the list.
- **Truncation through `float`** misreads values near an integer. D1 truncates the string.
- **Phase 0 changes how every integer in the model parses.** D5 keeps the output outside `w` exactly as it is today, and the corpus round trip is the gate.
- **Dropping a refused form in a required attribute** leaves an element the schema calls incomplete. Word refuses such a source anyway, and the Word check (Q4) says whether it opens the result.

## 8. Questions (decided 2026-10-26)

The user, 2026-10-26: "CR-021 Q1 to Q6 and Q8 as recommended; Q7 discussed separately".

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The rule: Word's measured reading, with a refused form read as unset and dropped at save (docx4j CR-065 D3 (a))? | Yes, so that the three ports read identically. |
| Q2 | Phase 0: make the parse exact for `int` in the fork (stage 14, 26.2.4), or leave Python's lenient `int()`, which reads section 2.2's forms as 240 without a record? | The fork change, for parity and so that nothing is repaired unrecorded. |
| Q3 | Scope: `w`, `w14` and `w15`, or every namespace? | `w`, `w14` and `w15`: Word was measured there; elsewhere, today's output exactly (D5) until a probe says more. |
| Q4 | A refused form in a required attribute: drop it, or write 0? | Drop it, as docx4j does, with the Word check confirming Word opens the result. |
| Q5 | The canon's lxml table: an attribute typed `int` in every class that declares it, or in every class the element name binds (the on/off rule)? | Every class that declares it, which settles 25 of section 2.5's 28 pairs; the canon leaves the other three to the typed path. |
| Q6 | The empty rsid: drop it at save? | Yes, as Word's save drops it. |
| Q7 | Word's 16-bit wrap: emulate it? | No. Only a value beyond 16 bits would show it: 32,767 twips signed (about 58 cm) or 65,535 unsigned (about 116 cm). Record it as a known difference. **Discussed separately (2026-10-26):** the user chose a guard where code sets a measurement (CR-022) and a report of an out-of-range value in a tree (CR-020 Phase B). This CR doesn't emulate the wrap on read unless that discussion decides otherwise. |
| Q8 | A Word check in phase C, with the unmeasured edges (`"2.9999999999999999"`, `"."`, `"+"`, a dropped required attribute)? | Yes: saved output changes, and those four cases are cheap to add. |

## 9. Effort

2 to 3 days here (A 1, B ½ to 1, C ¼, plus the user's Word run), and half a day in the fork for phase 0.

## 10. Phases A and B implementation notes (2026-10-26)

Phases A and B as built, at the user's go ("do CR-021"), on docx4j-xsdata 26.2.3's parse, where `int` is not exact (`.venv-fork`'s editable fork carried stage 14 from mid-session, but `runtime.EXACT_TYPES` stayed `{bool, Enum}` until phase 0, so every parse was 26.2.3's); then phase 0's half here (10.10), at the coordinator's brief, once the fork's stage 14 (item 32, `9c3065e5`) was released as docx4j-xsdata 26.2.4 (tag `docx4j-v26.2.4`, fork commit `0b407b64`). Phase C (the records and the Word check) is the user's next step. docx4j CR-065 D3 (a) is the rule; docx4j-generated-objects-ts CR-015 the counterpart, read for behaviour only.

**Commits.**
- `2d7e46c`: the fixtures, `tests/fixtures/integers/` (10.2).
- `e158f2f`: Phase A: `values.integer`, every reader through it, the census, the probe's oracle.
- `406d98d`: Phase B: the normaliser's integer rows, D5's rows, the empty rsid, the canon's mirror, their tests.
- `44e2888`: phase 0 here: `runtime.EXACT_TYPES` gains `int`, the floor `>=26.2.4`, `docs/UPSTREAM.md` row 27 (10.10).
- the commit after `44e2888`: this section and `tests/README.md`'s map and fixture rows.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the reader (D1) | `docx4j_py/model/values.py` | `integer(value) -> int \| None` (10.3 item 1) |
| the readers (D2) | 22 modules (10.4) | every reader of a WordprocessingML integer attribute through `values.integer`; the numbering emulator's `w:numId`, `w:ilvl` and `w:abstractNumId` keys through `definitions.decimal` |
| the census | `tests/content/test_integer_readers.py` | 120 listed `int(` and `float(` calls, each with its reason; a new one fails (10.5) |
| the save (D3, D5, D6) | `docx4j_py/openpackaging/normalise.py` | `integer_attributes(cls)` and `rsid_attributes(cls)`, cached per class, beside `on_off_attributes`; `_changes` finds the integer rows and the empty rsid in the same typed walk; `UNION_INTEGER_ATTRIBUTES` and `_union_rows`, found by name in the render; `INTEGER_NAMESPACES` |
| the trigger | `docx4j_py/openpackaging/parts/xml_part.py` | `XmlPart._empty_value`: whether the parsed source holds `""` or `''` (an empty rsid binds without a record) |
| the canon's mirror | `normalise.py`, `scripts/canon.py` | `normalise_tree` gains the integer rows (`_integer_tables`, Q5, with three pairs settled by holder), the empty rsid and the union rows; `_bound()` builds the model tables from `Meta` (10.7 item 6) |
| the oracle and the rows | `tests/content/test_integer_values.py`, `tests/content/test_integer_normal_form.py` | 10.6 |

### 10.2 The fixtures

`tests/fixtures/integers/`: the 22 family files of the share's `__tmp/cr065/families/` (no `~$` lock file was there) and Word 365's saves of the 14 Word opened, with the share's `cases.txt` and `README.txt` (`1a530e6a2a45`), copied read-only on 2026-10-26, each the same bytes as the share's (`cmp`); this package's `README.md` names the source (docx4j CR-065 phase 0, `aacaa4c45`, rounds two to four, the user's runs of 2026-10-26) and `SHA1SUMS` holds all 39 files. The all-cases file and the structure control are not copied: Word refused the first, and the second has no decimal.

Two things the README records that a test relies on:
- **Word's saves of `B` to `I` are round two's**, of the first generation (08:30 to 08:32), whose styles, settings and header in `B` to `H` carried the generator's sentinel integers, which Word wrote modulo 2^16 (`w:defaultTabStop` 61578, the style's indent -3959, the header's 25355). The sources here are the regenerated files (08:35), with plain integers outside the cases. The cases' own paragraphs, tables and sections hold the same values in both, so the tests compare cases only.
- **The case ids** are the regenerated probe's, as `cases.txt` gives them, one higher from `B20` on than the first design (after `A19` was added); Word's saves carry the same labels.

### 10.3 The rules as built

1. **`values.integer(value)`**: an `int` is itself (a `bool` is None: no integer attribute holds one); a `str` that `re.fullmatch(r"([+-]?)([0-9]*)(?:\.[0-9]*)?", value)` matches reads as `int(digits)` with its sign, or 0 where the digit group before the dot is empty (`""`, `"."`, `"+"`, `"-"`, `".5"`, `"-0.5"`); anything else, a `float` or `bytes` included, is None. ASCII `[0-9]` and `fullmatch`, so no other script's digits, no trailing newline (which `$` would let through) and no whitespace. The truncation is on the string: `"2.9999999999999999"` is 2, where `float()` gives 3.0. A digit string longer than Python converts (`sys.get_int_max_str_digits()`) is None rather than an exception. No 16-bit wrap (Q7).
2. **A reader** reads a refused form as unset, as if absent: a fallback the reader already had applies (`w:left` refused, `w:start` read; no `w:outlineLvl`, the style's heading level; no `w:ilvl`, level 0; no width, the default). `""` is a stated 0, so where a reader chose between two attributes with `or` (the numbering emulator's and `definitions.Indent.of`'s `w:left or w:start`), it now takes the first stated (`_stated`): an empty `w:left` is 0 rather than falling through to `w:start`.
3. **The save's typed rows (D3)**, in `_changes`, over the parts whose parse kept a value or whose source holds an empty value: a `str` in an `int` attribute of `w`, `w14` or `w15` is written as `str(values.integer(text))`, or dropped where that is None; a `str` already in that form (`"240"`, set in code) is no row. Each row is a `Normalised(part, path, before, after)` in `last_save.normalised`; `_applied` sets the field (to the `int`, or None) for the render and restores the text after it, so the tree keeps what was read and `save(normalise=False)` writes the source's spelling.
4. **Outside `w`, `w14` and `w15` (D5)**: a `str` in an `int` attribute is written as `str(int(text))` where Python's `int()` takes it (a token list, DrawingML's `dgm:forEach/@st`, `@cnt`, `@step`, token by token), and left as read where it does not. Under 26.2.3 no such row ever fires (a value `int()` takes is never kept); under the exact parse it keeps 26.2.3's output. **The unions**: in the 35 union fields that hold an `int` (none in `w`, `w14` or `w15`), `UNION_INTEGER_ATTRIBUTES` (23 elements, 42 (element, attribute) pairs; one field, msink's `CT_CtxNode/@semanticType`, no element binds), a value `int()` takes is written as that integer, found by name in each render, since with stage 14 such a value binds as the `str` member without a record.
5. **The empty rsid (D6)**: an attribute of `w` whose local name starts with `rsid`, valued `""`, is dropped and recorded (`before ""`, `after None`), found by value in the same walk (`rsid_attributes(cls)`). A part enters the walk when its parse kept a value **or** `XmlPart._empty_value` is set (the parsed bytes hold `""` or `''`: two substring tests, 0.43 ms over `Symbols.docx`'s 790 KB main part, where the parse is about 200 ms).
6. **The canon's mirror**: `normalise_tree` rewrites, by name, every attribute the save would write otherwise: the on/off rows as before; an integer attribute by its kind, from `_integer_tables()`: `"word"` (Word's reading; before stage 14 a value Python's `int()` takes is the serialiser's integer, so `"+240"` and `" 240 "` are `240`), `"python"` or `"tokens"` (D5's); an empty rsid dropped; then the render's rows, the very functions (CR-018 Phase B's and the unions'). The table holds a pair where **every class that declares the attribute** types it `int` alike (Q5): 217 pairs of a `w` element and a `w` attribute (192 declared by every class the name binds, 25 not), 53 more on `w14` elements, 1 on a `w15` element; 622 `"python"` pairs (3 of them on `w` or `w14` elements, whose attribute is another namespace's) and 9 `"tokens"`. The three pairs whose element name also binds a class typing the attribute otherwise are settled by the element's holder, where every class bound there agrees: `w:rPr/w:sz@w:val`, `w:ddList/w:default@w:val` and `w:lvl/w:start@w:val` (10.7 item 4). Building the integer tables costs 0.4 s once, in a process that compares, on top of the model scan the on/off table already made.

### 10.4 The readers

Through `values.integer` now, by module (CR-021 section 2.3's list, then what it did not name):
- **the content API**: `Paragraph.left_indent`, `right_indent`, `first_line_indent`, `space_before`, `space_after`, `line_spacing`, `outline_level` (`paragraph.py`); `api_script`'s indents, spacing, line and outline level; `writable_width`, `Table.column_widths_twips` (the grid, `table.py:446`, and the first row's `w:tcW`) and `TableCell.width` (`table.py`); **not named:** `picture.writable_width_emu` (`w:pgSz`, `w:pgMar`), `describe()`'s page in both branches (the typed one read `float()`; the lxml one, over a part never read, raised `ValueError` on `"abc"`), `heading_level_of` (`reports.py`, whose `int()` raised on a decimal `w:outlineLvl`), the comment and marker ids (`comments.py`), `ContentControl.id` (`controls.py`), and `lists.py`'s `w:abstractNumId`, `w:numId` and `w:ilvl` reads (`next_abstract_num_id`, `next_num_id`, `_ilvl`);
- **the python-docx facade**: `Section`'s lengths (`section.py`, the `ValueError`s on `"11907.0"` and `""`), `_Column.width`, `_Cell.width`, `_Row.height`, `grid_span`, `grid_cols_before`/`after` (`table.py`, the `ValueError` on `"4500.5"`), `BaseStyle.priority` (`styles.py`), `Font.size`, `ParagraphFormat`'s indents and spacing, `TabStop.position` and the tab order (`text.py`); `shared.py`'s `int()`s read caller lengths and are listed;
- **`read_run_options`' size** (`wml/builders.py`, under `Font`, `api_script` and `span_edits`), **not named**: it read `size / 2`, a `TypeError` on a kept `"22.5"`, with no `int(` for the census to find;
- **the numbering emulator**: `definitions._int` and `emulator._as_int`; **not named:** its keys. It keys definitions and levels by `w:numId`, `w:ilvl` and `w:abstractNumId` as text, read from the model and from lxml; through `definitions.decimal` (`str(values.integer(text))`) `"1.0"` and `"1"` are one list, as Word reads them. Without it, `lists_of`'s `int(label.num_id)` raised on a `"1.0"` present in both the paragraph and `numbering.xml`;
- **the revisions engine**: `tables._grid_before`, `_span`, `collapse_grid`'s widths, `remove_cell`'s width sum (`tables.py:796`), `reject_section`'s `w:linePitch`; `trees.marker_id` and `removals`' id index;
- **the text model**: the extractor's outline level, `w:ilvl` and `_twips` (indents, `w:gridSpan`), Markdown's style indent, `ResolvedView.removes_comment`;
- **the resolver**: no arithmetic on an integer field was found (an AST scan over `property_resolver.py`, `property_catalogue.py` and `style_util.py`); its one comparison, a heading style's `w:outlineLvl` against its name's level, is left (10.11 item 2).

### 10.5 The census

`tests/content/test_integer_readers.py` walks the AST of every module under `docx4j_py/model` and `docx4j_py/docx`, and of `wml/builders.py`, `wml/pictures.py`, `wml/sdt.py` and `traversal.py`, for every call of `int` or `float` whose argument holds no call of `integer`. Each must be in `LISTED`, keyed `<file>::<scope>::<the call>`, with one of eleven reasons; `test_every_int_and_float_is_listed` fails on one found and not listed, or listed and gone, as `test_private_readers.py` holds `read_contents`' callers. 120 are listed: 35 a caller's argument, 15 an integer attribute outside `w`, `w14` and `w15` (charts, `wp:extent`, SmartArt's `srcOrd`; D5), 14 a number the code computed, 12 constants, 12 hexadecimal text, 11 other text (style ids, addresses, field switches, Markdown, `lvlText`), 7 the compatibility mode, 4 a python-docx `Length`, 4 the emulator's decimal keys, 2 an `xsd:string` read as a number (`w:compatSetting/@w:val`, `cp:revision`), 1 `values.integer`'s own. Under `model/` and `docx/` there were 200 such calls at `d3dc747`, counting repeats; 129 are left, 6 of them `describe()`'s `float(integer(...) or 0)`, which go through it.

### 10.6 The oracle and the numbers

**Over the probe** (`test_integer_values.py`):
- `cases.txt`'s 65 cases are found where `LOCATIONS` says, 70 values (`G50`, `H56` and `H57` hold two, `H60` three), each holding the value `cases.txt` names;
- **`values.integer` of every source value is Word's integer, or None where Word refused the document**: 55 values against Word's save, 8 refused cases (`A11` to `A16`, `A18`, `A19`; `A5`'s three exponents were refused as a group), 4 values Word left out as the style's (`A08`, `A17`, `B25` 0, `F41` 22), and the grid's three (`G50`, `G51`), whose reading Word's recomputed grid hides;
- the content API's paragraph readings over `A1`, `A2a` to `A2d`, `A6`, `B` and `C`, the facade's over `A1`, `B`, `C`, `E` and `H` (lengths, sections, tables), the resolver's and the text model's reading of `I62`'s style indent, and `Font`'s size over `F41`, each equal over the probe and over Word's save (a value Word omits read as stated, 0);
- **the saves**: 92 rewrites over the 22 family files under 26.2.3's parse (88 kept values and `H60`'s four empty rsids; 95 with phase 0's, 10.10; per family `A` 17, `A1` 6, `A2` 4, `A2a`, `A2b`, `A2d`, `A2e`, `A3`, `A4`, `A6`, `A7` 1, `A5` 3, `B` 8, `C` 7, `D` 4, `E` 2, `F` 4, `G` 12, `H` 14, `I` 3; `A2c`'s `"+240"` and `A2f`'s `" 240 "` bind as 240 under 26.2.3); **every normalised save of the 14 families Word opened writes Word's value case by case** (75 values: 62, and `Z65` in the 13 other families), the four omitted values as Word read them and the grid truncated (4500, 5705, 1234); `H60`'s sectPr has no rsid left; `compare(normalise=True)` finds nothing but `word-normal-form` in any part, its count the save's rows plus the serialiser's re-spellings (`"+240"` in `A`, `A2` and `A2c`, and before stage 14 `" 240 "` in `A`, `A2` and `A2f`); `save(normalise=False)` writes every kept value as read, its only differences those re-spellings;
- `scripts/roundtrip.py` over the 36 files of `tests/fixtures/integers/` (122 parts; it reads no `settings.xml`): `--normalise` 122 of 122 canonically identical, 91 normalised, 97 `word-normal-form`, 748 `boolean-spelling`, 87 kept; `--no-normalise` 118 of 122, its 6 `attribute-value` differences the re-spellings.

**The rows on their own** (`test_integer_normal_form.py`, 27 tests, and 16 that run once `int` is exact, 10.10; 5 of them at `406d98d`): Word's rows in a typed tree and the tree restored; D5's typed rows and the token lists; the union rows in the render, and `UNION_INTEGER_ATTRIBUTES` derived from the model; an empty rsid dropped in a part whose parse kept nothing; a part with nothing odd not walked (`Symbols.docx`); a `str` set in code in a part that kept nothing written as set (CR-018's rule for on/off values set in code); the canon's tables and its normal form both sides of stage 14; `normalise_tree` by name and by holder.

### 10.7 Departures and choices, with why

1. **The census reaches past `model/` and `docx/`**, to the four hand-written modules beside the model that read it, because one of them held a reader: `read_run_options`' `size / 2` (10.4).
2. **A part is walked for the empty rsid when its source holds an empty value** (`XmlPart._empty_value`), as well as when its parse kept one. The brief's "the same walk" ran only over parts with kept values, so a part with an SSRS `w:rsidR=""` and nothing kept (a header) would have kept it. The alternative, the empty rsid found by name in every render, costs an XPath over every attribute, 2.9 ms over `Symbols.docx`'s main part per save, where the parse-time test costs 0.43 ms per parse; a false positive (`""` in text) costs a walk.
3. **The union rows are the render's, not the walk's**: once stage 14 makes `int` exact, a union value `int()` would take binds as the `str` member with no record, so no walk triggered by `part.kept` would find it (CR-018 Phase B's reason for its rows, 11.7 item 1). A consequence: a union value set in code (`a:pt x=" 240"`) is written `240`, where an integer set in code elsewhere is written as set.
4. **The canon settles `w:sz@w:val`, `w:default@w:val` and `w:start@w:val` by holder** rather than leaving them to the typed path: left out, `F41`'s `w:rPr/w:sz w:val="22.5"`, which the save writes 22, would be an `attribute-value` difference in `compare(normalise=True)` over family `F`, and section 5's "nothing else moves" would not hold. The holder rule is the Q5 rule applied one level up, derived from the model; `w:frame/w:sz` (a `str`) and a border edge's `w:start` are not in it.
5. **The canon counts the serialiser's re-spelling of a bound integer as `word-normal-form`** (`"+240"` as `240`; before stage 14 `" 240 "` too): the save writes the same, but has no text to record, since the parse bound an `int`. So over a part the canon's count is the save's rows plus those (10.6). Today's canon reported them as `attribute-value`; none is in `samples/`.
6. **The model tables are built from `Meta`, not `XmlMeta.qname`** (`_bound()`, which the on/off table shares): xsdata builds a type class's `qname` against the first parent namespace it met (CLAUDE.md's warning), so the on/off table held 75 keys that are type names (`CT_Bookmark`, or `{w}CT_Bookmark` once a parse has built the class under a parent), and the integer table's counts moved with test order (the fast suite found it). The on/off table's 286 real entries are the same either way; only the junk keys are gone.
7. **`values.integer(True)` is None, and a `float` or `bytes` None**: none is an integer attribute's value, and a `bool` is an `int` to Python.
8. **The emulator's keys go through `values.integer`** (`definitions.decimal`), beyond the brief's `_as_int` and `_int` (10.4); with them, `_stated` for `w:left`/`w:start` (10.3 item 2).
9. **A kept `str` already in Word's form is no row** (`"240"`), as an exact on/off literal held as a `str` is none (CR-018).

### 10.8 Pre-existing tests whose expectation changed

None in Phases A and B: no test edited, no golden, no ratchet row, no Word-save comparison, no acceptance artefact moved. Phase 0's are in 10.10.

### 10.9 The gates (Phases A and B, 26.2.3's lenient `int`)

Before is `d3dc747`, measured where a checkout is needed (the suite) by subtracting the new tests from the after run, since this tree was never reset: a run of the full suite in a scratch worktree at `d3dc747` gives 5,388 passed and 639 skipped, its skips the goldens and fixtures a worktree has no sibling checkouts for, so its numbers are not the checkout's.

| gate | before (`d3dc747`) | after (`406d98d`) |
|---|---|---|
| the full suite | 6,581 passed, 8 skipped (derived) | 6,735 passed, 13 skipped (4 min 45 s): the 159 new tests, 5 of them section 2.2's, skipped until `int` is exact |
| the fast suite (`-m "not slow"`) | --- | 5,354 passed, 13 skipped, 1,381 deselected (3 min 17 s) |
| the new tests | --- | 124 in `test_integer_values.py`, 32 in `test_integer_normal_form.py` (5 skipped), 3 in `test_integer_readers.py` |
| `roundtrip.py` over `samples/` (before: in the scratch worktree) | `--normalise` and `--no-normalise`: 50 of 50, 0 differences, 34 `boolean-spelling`, 0 skipped | the same, both ways, 0 kept, 0 normalised |
| the same over `tests/fixtures/integers/` (36 files, 122 parts) | --- | `--normalise`: 122 of 122, 91 normalised, 97 `word-normal-form`, 748 `boolean-spelling`, 87 kept; `--no-normalise`: 118 of 122, 6 `attribute-value` (the serialiser's re-spelt `"+240"` and `" 240 "`, which 26.2.3's parse binds) |
| `codegen/generate.sh --check` (a scratch worktree at `406d98d`) | --- | two generations byte identical, 308 modules; nothing generated differs from the commit |
| acceptance (`compare_acceptance.py` against `out/acceptance/` regenerated at `d3dc747` in a scratch worktree) | --- | 17 artefacts, 246 entries, 0 differ beyond the dates |
| the census | --- | 120 listed, every other call through `values.integer` |
| the cost | --- | the render's union rows 0.006 ms over `Symbols.docx`'s main part (none of their names is in it; lxml's `iter` by name finds none), the parse's empty-value test 0.43 ms; the canon's tables 0.4 s once |

### 10.10 Phase 0 here: `int` exact (docx4j-xsdata 26.2.4)

At the coordinator's brief (2026-10-26), once docx4j-xsdata 26.2.4 was on PyPI (tag `docx4j-v26.2.4`, fork commit `0b407b64`; stage 14 is `9c3065e5` and `ef1a7c50`), after Phases A and B were committed and their gates had passed. One commit, `44e2888`:
- **`runtime.EXACT_TYPES` is `{bool, Enum, int}`**, its comment and `parser_config`'s docstring citing the fork's stage 14 (item 32) and this phase. An integer now converts only from `[+-]?[0-9]+` as written; anything else is kept as its text: XSD's whitespace alone (`" 240 "`, `"\t240\n"`) on purpose (`reason="exact"`, no warning), anything else no XSD rule converts (`"1_000"`, `"٢٤٠"`, fullwidth `"２４０"`, `"\u00a0240"`, and the decimals and `""` as before) with the warning (`"invalid"`). `values.integer` reads each of section 2.2's forms as None, and the save drops it (D3); outside `w`, `w14` and `w15` D5 writes it as 26.2.3 did.
- **`pyproject.toml`'s floor** is `docx4j-xsdata[lxml]>=26.2.4,<26.3` in the dependencies and `docx4j-xsdata[cli,lxml]>=26.2.4,<26.3` in the `dev` extra. `RELEASING.md` named `>=26.2.1` (stale since 26.2.2) and names `>=26.2.4` now; the workflows name no version (`test.yml` installs the `dev` extra from PyPI). CHANGELOG's dependency line is phase C's, with the other records.
- **`.venv-fork`**'s editable install refreshed (`uv pip install -p .venv-fork -e '../docx4j-xsdata[cli,lxml]'`): it reports 26.2.4. Nothing else installed.
- **`docs/UPSTREAM.md` row 27**, as the fork drafted it.

**Tests whose expectation moved.** `tests/test_parser_configs.py::test_the_engine_s_config_reads_exactly` pins `{bool, Enum, int}`. Nothing else: the tests Phases A and B added that see the difference are written both ways, keyed on `runtime.EXACT_TYPES` (`test_the_rows_of_each_family`: 95 rows, `A19`'s `" 240 "` now kept and dropped in `A`, `A2` and `A2f`, where it bound as 240; `test_a_normalising_save_moves_only_the_rewritten_values`: the canon's re-spellings are `"+240"`'s alone; `test_a_save_that_keeps_the_spelling_changes_nothing`: `" 240 "` written as read). Section 2.2's tests, which key on `runtime.EXACT_TYPES` (the coordinator's correction: not on the fork's own `EXACT_TYPES`, which allows `int` from 26.2.4 whatever the engine asks), run: six forms each kept with its reason, warned exactly where `"invalid"`, read None, `space_before` 0, dropped at save, the tree keeping it; `"+240"` and `"0240"` bind as 240; a padded DrawingML integer (`a:ext/@cx=" 5"`) kept and written `5`; and the 42 (element, attribute) pairs of the union fields, for nine spellings each, write what 26.2.3's parse wrote. A scratch check over the same 42 pairs with twelve spellings (504 cases) found no difference either, and a corpus check (421 unique documents of `out/cr012-b-corpus/`, every typed part read and the package saved, normalising, once with 26.2.3's lenient `int` and once exact) found no part whose save differs and no value newly kept: no document of the corpus holds one of section 2.2's forms. Over the same corpus the normaliser's integer and rsid rows never fire (its 22 rows are CR-018's, in `pandoc-code.docx`).

| gate, `int` exact (`44e2888`) | result |
|---|---|
| the full suite | 6,751 passed, 8 skipped (4 min 53 s): `406d98d`'s 6,735 and 13, section 2.2's 5 skipped tests now run, with 11 more of `test_integer_normal_form.py` (two more forms, the reasons, the union pairs' nine spellings); 6,581 and 8 before, with the 170 new tests |
| `roundtrip.py` over `samples/` | `--normalise` and `--no-normalise`: 50 of 50, 0 differences, 34 `boolean-spelling`, 0 kept, 0 normalised, 0 skipped: as before |
| the same over `tests/fixtures/integers/` (122 parts) | `--normalise`: 122 of 122, 94 normalised, 97 `word-normal-form`, 748 `boolean-spelling`, 90 kept; `--no-normalise`: 119 of 122, 3 `attribute-value` (`"+240"`, which binds) |
| `codegen/generate.sh --check` (a scratch worktree at `44e2888`, the fork at 26.2.4) | two generations byte identical, 308 modules; nothing generated differs from the commit |
| acceptance (against the same `d3dc747` regeneration) | 17 artefacts, 246 entries, 0 differ beyond the dates |
| the cost | the parse of `Symbols.docx`'s main part, 26.2.3's lenient `int` against the exact, interleaved, 15 runs each on a loaded machine (load average 12): min 198.7 against 203.0 ms, median 209.3 against 211.6 ms; the 13 samples' main parts 248.9 against 250.5 ms, median 262.0 against 261.8 ms: within the noise (the fork's fast path is `str.isascii() and str.isdigit()`) |

### 10.11 Found, and not done

1. **Readers outside `w`, `w14` and `w15` are as they were** (D5): `InlinePicture.width` divides `wp:extent/@cx`, and the charts read `c:idx` with `int()`; a kept decimal there still raises. Word was not measured there (Q3).
2. **The resolver's heading-level comparison** (`_heading_layer`): a style whose `w:outlineLvl` is a kept `"0.5"` is copied with the name's level even where Word reads the same level. Harmless (the copy states what Word reads); left.
3. **The text model's `indent_left`** takes `w:left` when present, refused or not, where `Paragraph.left_indent` falls back to `w:start` for a refused `w:left`. Only a document Word refuses can tell them apart.
4. **`describe()`'s lxml branch** reads Word's rule (a refused form 0) while its typed branch reads what the parse bound, which before phase 0 is 240 for `" 240 "`; with phase 0 they agree.
5. **Values equal to the style's are written** (D7): Word's save leaves out `A08`, `A17`, `B25` and `F41`; this package writes the value Word read. The same meaning.
6. **Phase C**: CHANGELOG, CLAUDE.md, the index and the Status line (the brief leaves them to the user), and the Word check with the unmeasured edges (Q8: `"2.9999999999999999"`, `"."`, `"+"`, a dropped required attribute; `test_a_required_attribute_refused_is_dropped_too` builds the last).
7. **The 16-bit wrap** is not emulated (Q7); CR-022 is its guard.
8. **`XmlPart._empty_value` has false positives**: `""` or `''` in a part's text (two quotation marks typed together) sends the part through the typed walk, about 30 ms over a part the size of `Symbols.docx`'s, for no row. Not measured over the corpus; a tighter test (`=""` with XML's optional spacing) would cost a regular expression over the bytes.
9. **CHANGELOG's dependency line for 26.2.4** is phase C's, with the other records; `pyproject.toml` and `RELEASING.md` name it already.

### 10.12 Review (2026-10-26)

Verified by hand at `7fe0be8`:
- **The fast suite:** 5,370 passed, 8 skipped. The full suite was the agent's: 6,751 passed, 8 skipped.
- **`scripts/roundtrip.py`:** 50/50 canonically identical, 0 skipped. The pin is `>=26.2.4,<26.3` in both places, and `EXACT_TYPES` is `{bool, Enum, int}`.
- **The readers,** on the morning's probe (section 2.1's shapes):
  - "720.5" reads 36.0 pt, "2.5" 0.1 pt, "8.5" 0.4 pt;
  - the facade's `page_width` is 11907 twips, and `header_distance`, `footer_distance` and `gutter` are 0, with no `ValueError`.
- **Untouched and normalised saves:** a save with the part never read is byte for byte. A save after an edit writes every value truncated (3.5 → 3, 1.8 → 1, 141.99999999999994 → 141), SSRS's empties as 0, and the empty rsids dropped. That's 25 records in `last_save.normalised`, and `normalise=False` keeps the source.
- **The canon:** `compare(normalise=True)` counts the 25 as `word-normal-form`. The other differences it shows come from the probe's own out-of-order `w:jc` before `w:tabs`, not from this CR.
- **Section 2.2's forms:**
  - `" 240 "` is kept as `exact`;
  - `"1_000"`, `"٢٤٠"`, `"２４０"` and `"\u00a0240"` are kept as `invalid`;
  - each reads as unset and is dropped at save;
  - `"+240"` and `"0240"` bind as 240.
- **`values.integer`'s edges:**
  - `"2.9999999999999999"` → 2;
  - `"."`, `"+"`, `"-"` and `"-0.5"` → 0;
  - `"abc"`, `"1e3"`, `" 240"` and `"240,5"` → None, and so does `True`.

**The departures of section 10.7 are accepted on review.** The census's wider scope (1) found a real reader in `wml/builders.py`. The parse-time flag for `=""` (2) is two substring tests, and a false positive costs one walk at save. The canon settling the three ambiguous pairs by holder (4) and counting the serialiser's re-spellings (5) keep "nothing else moves" true over the probe.

One limitation is noted, and it isn't new: a source `"+240"` is written `"240"` even under `save(normalise=False)`, because it binds as an integer. That was so before this CR.

**The Word check** (phase C) is staged at `vbShares/Office 2016/__tmp/python-cr021/` (`out/cr021-word-check/build.py`, built at `7fe0be8`):
- the 22 families saved through this package, 95 rewrites in all; the 8 families Word refused as sources now have the refused values dropped;
- `cr021-edges.docx`, a dropped required `pgMar top="abc"`;
- five raw files for Word's reading of the unmeasured `"239.99999999999999"`, `"."`, `"+"` and `"-"`.

### 10.13 Phase C's Word check (the user, Word 365, 2026-10-26): passed

The user: "cr021: everything opened. cr021-edges gets normal margin settings". The files were built at `7fe0be8` and staged at `__tmp/python-cr021/`.
- **The families:** all 22 of this package's saves opened without a repair prompt, the 8 whose probe files Word had refused outright among them, with their refused values dropped.
- **The dropped required attribute** (Q4): `cr021-edges.docx`, whose `w:pgMar/@w:top="abc"` the save dropped, opened with Word's Normal margins, a 2.54 cm top. So dropping a refused form in a required attribute gives a file Word opens, with its default.
- **The four unmeasured edges** (`tests/fixtures/integers-edges/`, held by `tests/content/test_integer_edges.py`):
  - E1 `w:spacing/@w:before="239.99999999999999"`: Word saved 239. It truncates the string, not a double, as `values.integer` and docx4j do.
  - E2 `"."`, E3 `"+"` and E4 `w:ind/@w:left="-"`: Word opened each, so they aren't refused forms, and its saves leave the attribute out. Word omits direct formatting equal to the style's, and the style's value here is 0, so this package's 0 and docx4j's drop can't be told apart in these files. They'd differ only under a style whose value isn't 0. A probe with such a style would settle it; none is planned.
- **The right margin in `cr021-B.docx`:** case B27's text runs one character per line down the right margin. B27 is the probe's `w:ind/@w:left="10206.0"`, 18 cm, wider than A4's 15.9 cm text column. Word's own save of the probe has the same 10206, so Word renders the original probe the same way. It's expected.
