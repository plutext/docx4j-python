# CR-037: An unknown `w:` element in a content host, read as Word reads it (the port of core-ts's CR-026 D7)

**Status:** Proposed 2026-11-06, at the user's word ("yes, Python should follow D7"), on docx4j-core-ts `80a4707`. It reads what an unknown `w:` element holds in the fifteen hosts Word 365 reads it in, where this package drops the element with all it holds. Two phases, about 1 day, and a Word check. Nothing built. Questions in section 7, each with a recommendation. **Amended the same day** after check 76's round 6: in `w:pPr` and `w:rPr` what Word does depends on what the element holds (section 2), which adds Q5.

**Depends on:**
- **core-ts's check 76** (`80a4707`, `test/check76/`, its CR-026 section 4.4): the oracle. Round 6 (`76t27`, `76t28`) is uncommitted there, and on the share in `bridge/portfolio/`.
- **CR-030 section 11** (complete): the pass this one runs beside, and `part.skipped` for what a pass takes out.

**Counterpart:** docx4j-core-ts CR-026 section 3.7, decided, not built, with its D8 open. docx4j CR-076, proposed and uncommitted: the same rule, its D7 the same open point as D8. docx4j drops the element with all it holds today (its section 2.1), so until CR-076 is built this is a difference from it, recorded as DI-024.

## 1. Why

`w:madeUp` isn't a WordprocessingML element. Word 365 doesn't treat it as one thing. In a content host it takes the tags away and reads what they held, at any depth. In a `w:t` it drops the element with all it holds. In `w:pPr` and `w:rPr` it depends on what the element holds: a property goes, a run in a `w:pPr` is read at the head of the paragraph, and a run in a `w:rPr` makes Word refuse the file.

Here the fork skips the element, so what it held goes too. `part.skipped` says so, and `strict=True` refuses, so nothing goes silently. But the words Word shows are missing from what this package reads, and a save of a part that was read loses them.

## 2. What exists here (measured 2026-11-06)

Check 76's 29 `76t` sources against Word's save of each, read through the text model with headers, footers, comments and notes on. Round 6 is the portfolio session's bridge run, twice each, Word 16.0.20326.20232; its results are read here, not Word's hand saves.

| | Files | Here |
|---|---|---|
| Content hosts (`76t`, `76t1` to `76t4`, `76t7` to `76t17`, `76t19` to `76t25`) | 23 | **Wrong:** Word's text is missing, each element listed in `part.skipped` |
| A property in `w:pPr` or `w:rPr` (`76t5`, `76t6`) | 2 | Right: the element goes, as in Word |
| A run in `w:pPr` (`76t27`) | 1 | **Wrong:** Word opens it silently and reads `HIDDEN WORDSTARGET TEXT`, one run. Here the words are missing, the element listed in `part.skipped` |
| A run in `w:rPr` (`76t28`) | 1 | Word refuses the file: "Word experienced an error trying to open the file." Here it reads, the words missing, the element listed, as any file Word refuses is read and reported |
| Controls with no such element (`76t18`, `76t26`) | 2 | Right |

`w:t` itself is CR-030 section 11's, and reads as Word.

- **The hosts:** `w:p`, `w:body`, `w:tc`, `w:r`, `w:tbl`, `w:tr`, `w:sdtContent`, `w:hyperlink`, `w:ins`, `w:hdr`, `w:comment`, `w:fldSimple`, `w:smartTag`, `w:customXml` and `w:footnote`.
- **`mc:Ignorable`** plays no part, measured in a `w:p` only (`76t8`).
- **How common:** none of the 445 local `.docx` files outside `check76/` has any skipped entry, and core-ts found no such element in its 720 packages. This is correctness on a rare shape, not a common loss.
- **The seam:** `XmlPart._parse` already parses, reads the skipped report, repairs the bytes and parses again, for CR-030 alone. A clean part pays the test of an empty report.
- **What the report can't say:** `SkippedNode` names an element the parse found no place for. That's a name the model doesn't hold, or a known element in the wrong place, and only the first is this CR's.

## 3. Design

### 3.1 The pass

A new `openpackaging/read_through.py`, over lxml, as `breaks.py` is.

- **What it takes:** an element in the `w:` namespace whose name the model doesn't hold, whose nearest ancestor with a known name is one of the fifteen hosts (so `76t7`'s nesting works). What it held takes its place, in order, with its text and tails.
- **A run in a `w:pPr`'s unknown element (Q5):** the run goes to the head of the paragraph, and the tags are reported. Only a direct run, the shape measured, and only there.
- **What it leaves:** every other skipped element, as today. That is a property in `w:pPr` or `w:rPr` and anything in `w:t`, measured; a run in `w:rPr`, which Word refuses; anything else such an element holds in `w:pPr`, unmeasured; and every host not measured: `w:ftr`, `w:endnote`, `w:del`, `w:sdtPr`, `w:trPr`, `w:tcPr`, `w:tblPr` and `w:sectPr`, as core-ts's list has it. Adding a host later is then the whole change.
- **What its host doesn't admit:** unmeasured. The pass doesn't look, so a `w:p` left in a `w:p` is skipped by the second parse and listed, as any misplaced element is.

### 3.2 One repair, one reparse

`_parse` asks each pass whether the report triggers it. It runs read-through first and CR-030's pass second on one lxml tree, then parses once. The order matters: a `w:br` an unknown element held directly in a `w:p` is then CR-030's rule 1's to place. A clean part still pays only the test of an empty report.

### 3.3 The model's names

`generate_el.py` writes the `w:` namespace's element names into `el_index.py` as a frozenset, from the same metadata as `codegen/el_tables/wml.json`'s 604 elements. The fork doesn't change, and `generate.sh --check` holds it.

### 3.4 The report (Q1)

One `SkippedNode` for each element taken, through the seam CR-030 section 11 added, so `part.skipped` lists it and `strict=True` refuses, though no words go. This is what core-ts's section 3.7 does, and the two ports share the report (CR-033 Q9).

### 3.5 The save and the round trip

A part that was read saves in the shape read through, as Word's own save writes it: the tags gone, the content kept. An unread part stays byte for byte. CLAUDE.md's round-trip rule gains this beside CR-030's exception: `canon.compare(..., repaired=True)` and `scripts/roundtrip.py` run the pass, counted as a category of their own.

CR-033's gate isn't built. When it is, a changed part holding such tags is core-ts's case in its section 3.7: refused until the caller accepts the loss, though no words go. That stays CR-033's to decide.

### 3.6 Fragments (Q2)

`_parse_one` and `_repair_breaks` refuse any loss unless `lenient`. The pass runs in both, so `insert_ooxml` refuses such a fragment, and `lenient` reads it through. `wml()` stays strict.

## 4. Tests and oracles

- Check 76's 29 sources and Word's saves copied into `tests/fixtures/check76/`, unedited, against core-ts's `manifest.json`. `76t28` has no save. `76t27`'s is a bridge save, so only its body is compared.
- Each source against Word's save in tokens and plain text, over the part it touches: the 23 and `76t27` as Word reads them, and the four unchanged. `76t28` read and reported.
- `part.skipped`'s entry for each, a nested element's two, and none for a control.
- A known element in the wrong place left alone. An unmeasured host left alone. A `w:br` held in a `w:p`, placed by rule 1.
- Save and reload; a save against its repaired source; an untouched save byte for byte.
- `strict=True` and `insert_ooxml` refusing, `lenient` accepting.
- A clean part parsed once.

## 5. Phases

| Phase | What | Effort |
|---|---|---|
| A | The names table, the pass, `_parse`'s one repair, the fragment route, `canon` and `roundtrip.py`, the fixtures and tests | ¾ day |
| B | The records: CHANGELOG, `CLAUDE.md`'s round-trip rule and engine paragraph, `tests/README.md`, the index, the registry; the Word check's files | ¼ day, then the Word run |

The Word check: this package's save of each of the 23 and `76t27`, read beside Word's own, since a save changes. Phase A's fixtures wait for core-ts to commit round 6.

## 6. Risks

1. **A pass that moves content is the riskiest kind of repair.** It only ever acts on a part whose report names such an element, and none of the local files has one, so it cannot change what a clean part reads. The corpus stays the guard for that.
2. **`strict=True` refuses a file Word opens silently.** That's what it says today, and Q1 keeps it.
3. **Unmeasured hosts stay dropped,** so `w:ftr` and `w:endnote` lose words where Word probably keeps them. It's the same in core-ts, and the report names each.
4. **docx4j still drops.** Its CR-076 proposes the same rule. Until that's built the two differ on these files, which sits with DI-024.

## 7. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Does an element read through go in `part.skipped`, so `strict=True` refuses, or in `part.warnings`, as CR-030's `br.outside_run` does for a repair that loses nothing? | **`part.skipped`.** The tags are lost, whatever the words do. core-ts does this, and asked that the ports agree. |
| Q2 | Does `insert_ooxml` refuse such a fragment, unless `lenient`? | **Yes.** It's a caller's fragment, and refusing what a read can't keep is that route's rule. core-ts hasn't settled its route, so this goes to it. |
| Q3 | Exactly core-ts's fifteen hosts, or measure `w:ftr`, `w:endnote` and the rest first? | **The fifteen now.** Each added host is one line, and the shape is rare. docx4j's CR-076 D1 recommends the same. |
| Q4 | Which release? | **The next,** with CR-030 section 11. It should not hold a release up. |
| Q5 | A run in a `w:pPr`'s unknown element (`76t27`): move it to the head of the paragraph as Word does, or leave it dropped and reported? | **Move it, a direct run alone.** It's the same loss of words Word shows, and a few lines. core-ts's D8 and docx4j's CR-076 D7 are this question, each open with the same recommendation, so one answer does for all three. |
