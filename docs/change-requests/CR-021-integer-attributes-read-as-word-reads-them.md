# CR-021: Integer attributes read as Word reads them --- a decimal truncated toward zero, the empty string 0, every other form refused; the port of docx4j CR-065's rule

**Status:** Proposed 2026-10-26, at the user's request ("yes please draft the CR"). The portfolio session had relayed the user's request the same day: propose a fix for a decimal in a WordprocessingML integer attribute, which docx4j repairs on load and this package saves back. The rule comes from docx4j CR-065's phase 0: one Word probe shared by the three ports, 65 cases over four rounds, run by the user in Word 365 on 2026-10-26 and checked here against Word's own saves (section 3). **Nothing has started.** Four phases (section 6): 0, the fork's stage 14 (Q2); A, the reader; B, the save; C, the records and a Word check. 2 to 3 days here, plus half a day in the fork. Eight questions wait on the user (section 8).

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
- a `str` in Word's form reads its integer part with its sign. The truncation is done on the string, not through `float`: `float("2.9999999999999999")` is 3.0, where the string's integer part is 2. Word's reading of more digits than a double holds is unmeasured: `141.99999999999994` reads 141 by either route. Phase C's Word check adds the case (Q8). `""`, and a form whose digit groups are both empty (`"."`, `"+"`), read 0, as docx4j's test does; those two are also unmeasured;
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
- **The unions** of section 2.4 aren't integers to the parse, and stay as they are.

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

## 8. Questions

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The rule: Word's measured reading, with a refused form read as unset and dropped at save (docx4j CR-065 D3 (a))? | Yes, so that the three ports read identically. |
| Q2 | Phase 0: make the parse exact for `int` in the fork (stage 14, 26.2.4), or leave Python's lenient `int()`, which reads section 2.2's forms as 240 without a record? | The fork change, for parity and so that nothing is repaired unrecorded. |
| Q3 | Scope: `w`, `w14` and `w15`, or every namespace? | `w`, `w14` and `w15`: Word was measured there; elsewhere, today's output exactly (D5) until a probe says more. |
| Q4 | A refused form in a required attribute: drop it, or write 0? | Drop it, as docx4j does, with the Word check confirming Word opens the result. |
| Q5 | The canon's lxml table: an attribute typed `int` in every class that declares it, or in every class the element name binds (the on/off rule)? | Every class that declares it, which settles 25 of section 2.5's 28 pairs; the canon leaves the other three to the typed path. |
| Q6 | The empty rsid: drop it at save? | Yes, as Word's save drops it. |
| Q7 | Word's 16-bit wrap: emulate it? | No. Only a value beyond 16 bits would show it: 32,767 twips signed (about 58 cm) or 65,535 unsigned (about 116 cm). Record it as a known difference. |
| Q8 | A Word check in phase C, with the unmeasured edges (`"2.9999999999999999"`, `"."`, `"+"`, a dropped required attribute)? | Yes: saved output changes, and those four cases are cheap to add. |

## 9. Effort

2 to 3 days here (A 1, B ½ to 1, C ¼, plus the user's Word run), and half a day in the fork for phase 0.
