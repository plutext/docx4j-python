# CR-039: A repeated child and an `mc:AlternateContent` root, read as Word reads them

**Status:** Complete 2026-11-07: Phases B and C built, Phase A folded into B, and the Word check passed through the bridge at the user's word (section 9.4). Decided 2026-11-07 (the user: "yes update Q9 then take decisions as recommended, commit then start phase b"): Q9 changed to read both roots as Word does, every other question as recommended. Proposed 2026-11-06 (the user: "write the CR - we may defer implementation of it indefinitely..."), on the objects session's message of that day and section 7.15 of the portfolio's `text_extraction_behaviour_model.md` (uncommitted when read). Nothing is built and nothing waits on it. Phases B and C are planned (the user, 2026-11-07: "CR-039: we'll implement phases B and C shortly"), with D after them, which answers Q7. Phase A is folded into B and the rest dropped (Q10). Which release they go in isn't said. Phase B built 2026-11-07 (section 8), then amended to WD 0.98's REQ-318 to REQ-320, with Phase C and a `w:body` in a `w:body` (section 9, the user: "go"). Amended 2026-11-07 on core-ts's message of that day, which relays the user's word that this repository consider its CRs: section 2.4 is new, and sections 3.2, 3.4, 3.5 and 3.6 and Q2 to Q6 changed, with Q9 new. As first written, Phase B refused every root but `w:document`, and Word opens several of them.

**Depends on:** nothing to build. The WD says nothing about either shape, and the portfolio drafts it (Q1). Phase B starts ahead of that at the user's word and is amended to it, as CR-037 was. **CR-033** is the save's half (section 2.3). Nothing here waits on phase 0 of the two counterparts (Word check 81), which isn't scheduled and has no date: core-ts's CR-028 may be deferred indefinitely too. Phase B rests on rounds already run, and Phase C builds what's measured. Phase 0 would only widen C.

**Counterpart:** docx4j CR-079 and core-ts CR-028, both proposed with nothing built, the second keeping the first's case names and sharing its phase 0. This CR was written without having read them, and now follows core-ts's where it can, since the two engines stay one design. Today docx4j and core-ts keep the second of two bodies where this package keeps the first (section 7.15's table).

## 1. Why

Word opens a `word/document.xml` that gives a child twice, or whose root is `mc:AlternateContent`, and reads what it holds, with no message on all but one file (`rw-22`, section 2.2). This package keeps the first of a repeated child and reads nothing under such a root. So the text differs from what Word shows, and the ports differ from each other: the same two-body file gives different words here and in docx4j.

Nobody is known to write these files. None of the 445 local `.docx` files outside `check76/` has any skipped entry (CR-037 section 2), and all 834 Word packages core-ts read have `w:document` (its CR-028 section 2.4), so this is correctness on a rare input, which is why it can wait. Two things stop it being only a curiosity:

- **Three of the thirteen repeated shapes aren't malformed.** `w:b` and `w:sz` given twice are valid in ISO 29500-4:2016 Transitional, where `EG_RPrBase` is unbounded (the objects session's reading; docx4j's `wml.xsd` doesn't allow them). This package reads the first and Word the last.
- **A file can show a reader one text and Word another.** That's how content gets past a text-based check. This is an inference from the table in section 2: no such file has been seen.

## 2. What exists here (measured 2026-11-06 at `fac442b`)

The inputs are the Word bridge's, in the share's `bridge/portfolio/` as `rc-*` and `rw-*`, built by the objects session, each run twice on Word 16.0.20326.20232, the two equal by section 7.15. Word's column here is from the `-a` run's `.txt` and `.word.docx`; a box's words are from the save alone, since Word's text save leaves boxes out. The probe read each `-a` input and changed nothing: `load`, `.contents`, `extract`, `save`, the facade's `Document`. Its controls agree with Word.

### 2.1 A child given twice

All thirteen differ from Word, and each second copy is listed in `part.skipped`.

| File | Given twice | Word | Here |
|---|---|---|---|
| `10-body` | `w:body` | Both, joined; the second's section, `w:pgSz` 15840 | First; its section, 11906 |
| `11-sdtContent` | `w:sdtContent` in a block `w:sdt` | Both, joined | First |
| `12-txbxContent` | `w:txbxContent` in a VML box | Both, joined, **the second first** | First |
| `20-pPr-values` | `w:pPr` | `w:jc` right, the later | center, the first |
| `21-pPr-merge` | `w:pPr` | Merged: `w:ind` and `w:jc` | `w:jc` only |
| `22-rPr-values` | `w:rPr` | `w:sz` 60, the later | 40, the first |
| `23-rPr-merge` | `w:rPr` | Merged: `w:b` and `w:i` | `w:b` only |
| `30-jc` | `w:jc` | right, the last | center, the first |
| `31-b-on-off` | `w:b` on, then off | Off | On |
| `32-b-off-on` | `w:b` off, then on | On | Off |
| `33-sz` | `w:sz` | 60, the last | 40, the first |
| `40-sectPr` | `w:sectPr` | The later, 15840 | The first, 11906 |
| `41-tblPr` | `w:tblPr` | The later, `w:jc` right | The first, center |

In `40-sectPr` and `41-tblPr` both copies hold the same properties, so whether Word merges them wasn't tested.

### 2.2 The root

The parse is handed the class, so the root's name is never compared: any root binds as `Document`. On these thirteen its children are skipped and the tree is empty, and the facade raises `ContentError`, since a document with no `w:body` has no body. A root with a `w:body` straight under it would read as a normal document: `<foo><w:body>` binds its body with an empty report (tried 2026-11-07; Word's reading of that file isn't measured).

| File | Root of `word/document.xml` | Word | Text here | `part.skipped` |
|---|---|---|---|---|
| `rw-20`, `rw-21` | `mc:AlternateContent`, a branch taken | Opens silently, that branch's words | Empty | `mc:Choice`, `mc:Fallback` |
| `rw-23`, `rw-26`, `rw-27` | The same, the branch holding two `w:document` | Joined silently; the second's section, margins included | Empty | The same |
| `rw-24` | The same, nothing selected | Opens silently, empty | Empty, which agrees | `mc:Choice` |
| `rw-22` | The same, a `Fallback` with no `Choice` | Repaired, keeping its words | Empty | `mc:Fallback` |
| `rw-10` to `rw-13` | `foo`, `z:foo`, `w:foo` | Refused | Empty | Each `w:document` |
| `rw-30`, `rw-31` | `w:wordDocument`, `pkg:package` | Refused | Empty | `w:body`; each `pkg:part` |

`LoadOptions(mce_preprocess=True)` and `mc_load_mode="word"` give the same empty text (measured on `rw-20`, `21`, `23`, `24`): `resolve_alternate_content`'s root "is not itself replaced". `mc_load_mode.branch_taken`, called on the root, already answers Word's branch on all seven `mc` root files, and `None` on `rw-24`.

`rw-25`, a `Choice` not understood with no `Fallback` in the body, is a different question. Here it reads the `Choice`, as REQ-305 requires, and Word shows neither. The user ruled on 2026-11-07 that REQ-305 itself changes for all three engines, and this CR doesn't carry that change.

### 2.3 What a caller sees, and the save

- **A report, and only that.** The load logs each item at WARNING, `part.skipped` and `package.skipped` list it, and `strict=True` refuses (tried on `10-body` and `rw-20`). `extract`'s `TextResult.warnings` is empty on all 26, so a caller holding only the text can't tell.
- **A save after `.contents` writes the loss.** `10-body` saves one body, 405 bytes of 713. Each of the thirteen root files saves a 141-byte `w:document` with no body, so the whole document goes. With `.contents` unread, `extract` then save keeps the source bytes (tried on `10-body` and `rw-20`).

The second point isn't this CR's to fix. It's CR-033 section 1.2's lossy write: CR-033's gate refuses such a save (Q1 decided yes, 2026-11-05) and its Q8 copies a part read and not changed. CR-033 isn't built. So deferring this CR leaves the text as it is, and the save's loss stays until CR-033 is built, whatever happens here.

### 2.4 Another story's root (measured 2026-11-07 at `0b9ee09`)

core-ts built these and read Word's saves (its CR-028 sections 1.1 and 2.4), each file twice with the two runs agreeing, in the share's `bridge/portfolio/` as `or-*`. Word has no one answer. The probe here read each and tried a write in memory. `or-00`, the control, agrees.

| File | Root of `word/document.xml` | Word | Text here | `part.skipped` |
|---|---|---|---|---|
| `or-10`, `or-11` | `w:hdr`, `w:ftr` | Silent, a blank page. Its save has an empty body, the words in the `separator` footnote with nothing referring to them; whether they ever draw isn't measured | Empty, which agrees with Word's text | Each `w:p` |
| `or-12`, `or-13`, `or-15` | `w:footnotes`, `w:endnotes` | Refused | Empty | Each note |
| `or-14` | `w:comments` | Repair offered, Show Repairs empty. Its save holds the comment | Empty | The `w:comment` |
| `or-30` | `w:styles` | Silent, empty | Empty, which agrees | The `w:style` |
| `or-32` | `w:body` with no `w:document` | Silent, the words in the body | **Empty** | `w:p`, `w:sectPr` |
| `or-31` | `w:madeUp` holding a `w:p` | Silent, the words in the body | **Empty** | The `w:p` |
| `or-50` | `mc:AlternateContent`, `w:hdr` in a `w14` `Choice`, `w:document` in the `Fallback` | As `or-10`, the `Fallback`'s words nowhere in its save | Empty, which agrees | `mc:Choice`, `mc:Fallback` |
| `or-40`, `or-41` | `w:hdr`, `w:document`, under a header's content type | Repair offered, then refused after Yes | A plain `OpcPackage`, and `extract` gives `""` | Not read |

- **A write can't go nowhere here.** core-ts loses a write through `Body` on these roots. Here `body_of` raises `ContentError` (`body.no_content`) on all ten with a `MainDocumentPart`, and the facade's `add_paragraph` the same. So core-ts's D10 has no counterpart here.
- **The save** after `.contents` writes the same 141-byte `w:document`, and section 2.3's point holds.
- **A package with no main part is quieter than it should be.** On `or-40` and `or-41`, `extract` gives empty text with no warning, where core-ts refuses and the facade raises `InvalidFormatException`. `body_of` raises `AttributeError`, which names nothing useful.
- **What decides isn't known.** `w:madeUp` holding a `w:p` opens and `w:foo` holding a `w:document` (`rw-12`) is refused. Name and content both differ.
- **A root not in this table isn't measured here.**

## 3. Design

### 3.1 The pass

`openpackaging/repeated.py`, a pass over lxml at `XmlPart._parse`'s seam, beside `read_through` and `breaks` and through the same repair and one reparse. Order: the root, the repeated children, then CR-037's and CR-030's passes on the same tree, since a `w:document` taken from a branch may hold what they repair.

**The trigger** is the skipped report, as for both passes: a clean part pays the test of an empty report. Every shape in section 2 is already in it. `parent_qname` can be a type name (`CT_PPr`) in whichever namespace the metadata met first, so the test compares local names.

### 3.2 The root (Phase B)

- **`mc:AlternateContent`:** replaced by what `branch_taken` answers, and the rows below then apply to what the branch holds, as `or-50` needs. One `w:document` becomes the root; two are joined as two `w:body` are (section 3.3); `None` leaves an empty `w:document` holding an empty `w:body`, which the facade can read, as core-ts's D4 has it.
- **The roots Word refuses (WD REQ-318):** `w:footnotes`, `w:endnotes`, `w:tc`, Word 2003's `wordDocument` and `pkg:package`; an unknown `w:` root holding a `w:document` or, directly, a `w:r`; and a root in no namespace, or in one neither understood nor in its `mc:Ignorable`, holding a `w:document` or a `w:p`. Refused as a part not well formed is (CR-031, and Q4): `Docx4JException` naming the part, kept as `part.unreadable`, the load lazy, an untouched save writing its bytes.
- **The roots Word reads (REQ-319):** a `w:body`, wrapped in a `w:document`; a `w:p`, as a body's one paragraph; a `w:sdtContent`, as a body holding what it holds; and an unknown `w:` root holding `w:p`, `w:tbl`, a block `w:sdt` or a further unknown element, wrapped so that CR-037's pass takes its tags off. A `w:sectPr` directly in that root has no effect, a `w:body` in it is the body, and empty it is an empty document.
- **Every other root yields nothing (REQ-320),** with section 3.6's warning: no text, what it held in `part.skipped`, `Body` refusing. That covers `w:hdr`, `w:ftr`, `w:styles` and `w:comments`, whose words aren't a reading to copy (core-ts CR-028 section 4), a root in a namespace its `mc:Ignorable` names, and anything unmeasured, among them `foo` holding a `w:body`.
- **Main part only.** Every measurement is `word/document.xml`. A header's or a note's root stays as it is unless a round measures it (core-ts's D5).
- **The root's name is read once,** in this one part, since `<foo/>` leaves no report and `<foo><w:body>` an empty one.

### 3.3 The repeated child (Phase C)

| Rule | Elements | Measured |
|---|---|---|
| Join in order | `w:body`, `w:sdtContent` | A body in a `w:document`, a block `w:sdt` |
| Merge, the later value winning | `w:pPr`, `w:rPr` | One property each way, one level deep |
| Keep the last | `w:jc`, `w:b`, `w:sz` | One file each (Q5) |
| Keep the later | `w:sectPr`, `w:tblPr` | Same properties in both copies |

Joining bodies keeps the last final `w:sectPr` stated and drops the others, a body with none being passed over: in `rt-17` the third body has none and the second's stands. A `w:sectPr` in a paragraph's `w:pPr` stays (`rt-15`).

A `w:body` in a `w:body`, at any depth, gives way to what it holds in place, the same section rule applying (`nd-20` to `nd-24`). The WD doesn't require it.

Everything else stays as it is today: the first kept, the later listed. That covers a second `w:txbxContent` (Q6), a run-level or cell-level `w:sdtContent`, and a second `w:tcPr` or `w:trPr`. Phase 0's list covers most of those, and whether "the last" means the whole element or each attribute (`w:ind w:left` then `w:ind w:right`), since `rc-30` to `rc-33` hold one attribute each.

### 3.4 The report (Q3)

One warning a part with a count, on `part.warnings` and `package.warnings`, as `BrOutsideRun` is, since a file read as Word reads it has lost nothing: `part.read_as_word`, naming the root where it wasn't `w:document` and counting the children joined. What goes whole goes in `part.skipped` beside it: a value a later one replaced, an earlier `w:sectPr`, a branch not taken. So `strict=True` refuses those, and a part that only joined loads under it, as CR-030's repaired part does.

That's core-ts CR-028 section 3.3. It replaces this section's first draft, one `SkippedNode` for each element, under which a join that loses nothing would have been refused at a save. `part.unexpected_root` is decided here (the user, 2026-11-07), and `part.read_as_word` is a proposal for core-ts's D8.

### 3.5 The save and the round trip

A part read and repaired saves in that shape, Word's own save's. An unread part stays byte for byte. `canon.compare(..., repaired=True)` runs the source through the same pass and counts `repeated-child`, every other category still held to zero. That's a fourth exception in `CLAUDE.md`'s round-trip rule.

Once CR-033 is built, its gate refuses a re-marshalled part with `part.skipped` entries unless `repair="word"`. A join alone leaves none, so that part saves. A part that dropped a value or a branch needs the switch, which is core-ts's D3.

`wml()` and the fragment route stay as they are: a fragment is the caller's code, and `wml()` refuses what it can't bind.

### 3.6 A root left as it is (Phase B, Q10)

Phase A's two lasting pieces are in Phase B, and its third is dropped:

- **A root that isn't `w:document` and isn't read as Word reads it:** one warning on the part at the parse, on `part.warnings`, `package.warnings` and the text result, naming the root. This is core-ts's D11, under the code `part.unexpected_root`. REQ-801's list needs it added.
- **A package with no main part:** `extract` refuses, as core-ts's `Texts.of` does, and `body_of` raises `ContentError`.
- **Dropped:** `content.not_read` for a second container. After Phase C only unmeasured shapes would raise it.

## 4. Tests and oracles

- **Word's files are the oracle:** the 13 repeated and 13 root inputs, core-ts's 12 `or-*` files named in section 2.4, their controls and Word's saves, as fixtures under `tests/fixtures/repeated/` with a `SHA1SUMS` and a README, as `check76/` is. They come to under 1 MB, and today they exist only on the share and in `docx-tool-identifier`'s git-ignored `runs-local/` (Q8).
- **Text:** equal to Word's `.txt`, and to Word's save for a box.
- **Properties:** the target paragraph's `w:pPr` and `w:rPr`, the `w:tblPr` and the `w:sectPr` equal to Word's save, by `canon`.
- **Refusal:** the nine roots Word refuses, each named, with an untouched save keeping its bytes.
- **Roots Word opens:** each stays as it is today, with the warning.
- **Nothing else moves:** the corpus round trip, `check76/`, `part.skipped` empty on `samples/`.
- **The warnings:** each on its shapes, none on a clean file, and `extract` refusing `or-40` and `or-41`.

## 5. Phases

Each phase stands alone, and any of them can be left out.

| Phase | Work | Cost |
|---|---|---|
| A | Folded into B, the rest dropped (Q10) | none |
| B | The root: `mc:AlternateContent` read through `branch_taken`, a bare `w:body` and an unknown `w:` root read through, the roots Word refuses refused, the rest warned; two `w:body` joined, which two `w:document` need; `extract` and `body_of` on a package with no main part; canon; the fixtures | ¾ day |
| C | The other repeated children, with their tests | ¾ day |
| D | The records and the Word check of B's and C's saves, through the bridge at the user's word | ¼ day |


**What B and C add.** Estimates, set against `read_through.py` (364 lines) and its 1,128 lines of tests:

- **Code:** about 60 lines for B and 150 for C, in `repeated.py`, plus about 20 in `XmlPart._parse`, `canon.py` and `roundtrip.py`. Tests about 150 and 300 lines. No change to the model, the generator or the fork.
- **A clean file, Phase C:** nothing. The trigger is the skipped report, which is already read, and the module is imported only when it fires.
- **A clean file, Phase B:** one read of the main part's root name, the only cost every file pays, and only when the part is read. Timed at about 23 µs with a pull parser fed 2 KB at a time; `etree.iterparse` took 0.1 to 0.45 ms, since it reads 32 KB first. The skipped report can't replace this read, as section 2.2's `<foo><w:body>` shows.
- **A file the pass repairs:** that part parses twice, the route CR-030 and CR-037 already take, and if they fire too it's still one reparse. About 2.1 times the part's parse, summed from its steps on `samples/` rather than run: `Symbols.docx` (772 KB) 200 ms to about 418 ms, of which the pass's own lxml parse, walk and write are 18 ms. The lxml tree is transient.

## 6. Risks

- **One file per rule.** Each row of section 3.3 rests on one small file, two runs, with no styles part. A real document might join differently, above all a body with headers and several sections.
- **A rule beyond the measured names.** `w:vanish` given twice decides whether words show, and nothing here measures it (Q5).
- **Word has no one rule for the root.** Section 3.2's list is what was measured, and a shape off it gets today's reading.
- **`w:txbxContent`'s order is unexplained.** Copying it pins something nobody understands.
- **A root refused where today gives empty text.** A caller who got `""` gets an exception. Word refuses the same files, so that's the honest answer, but it's a change.
- **The evidence can go.** See Q8.

## 7. Questions (decided 2026-11-07: Q9 read as Word, the rest as recommended)

| # | Question | Decision |
|---|---|---|
| Q1 | Does the WD take these rules first? | **Yes.** It's the shared rule for three engines, and the portfolio drafts it from section 7.15 and this CR. Phase B starts now at the user's word and is amended to the WD as it lands. Phase C waits for it. |
| Q2 | Are the warnings worth having, and under which codes? | **Yes:** `part.unexpected_root` for a root left as it is, as Q10 has it, and no `content.not_read` for a second container. Agreed with core-ts and docx4j when they answer. |
| Q3 | How is what a pass does reported? | **One warning a part with a count, and `part.skipped` for what goes whole**, as core-ts's section 3.3 has it. A strict caller then loads a file that only joined. |
| Q4 | Are the roots Word refuses refused here, rather than read as an empty document? | **Yes, those measured, by name and shape.** Every other root stays as it is, with the warning, since Word opens several and what decides isn't known. |
| Q5 | Does "keep the last" cover every property in `w:pPr` and `w:rPr`, or only the three measured? | **Those measured, and any phase 0 adds if it runs**, as core-ts's D6. `w:vanish` is on its list now, since it decides whether words show, so it stays unread until then. |
| Q6 | Follow Word's second-first order in a VML box? | **Not without a second shape and three copies**, as core-ts's D7, which phase 0 would give. Until then a second box stays as it is, and Phase C goes without this row. |
| Q7 | Does this stay proposed, or move to deferred now? | **Decided 2026-11-07:** neither, phases B and C are to be built shortly. |
| Q8 | Keep the Word files in this repository now, even if nothing is built? | **Yes.** Rebuilding the inputs is easy, but Word's saves need the bridge again. It's under 1 MB of fixtures, and a deferred CR can outlive a share. |
| Q9 | Read a bare `w:body` and `w:madeUp` holding a `w:p` as Word does? | **Yes.** A scratch trial (2026-11-07) matched Word's text and saved body on `or-31` and `or-32`, and `Body` wrote to both. About 15 lines, and `w:madeUp` is CR-037's pass once wrapped. The rule rests on one file, and if it's wrong we read words from a file Word might refuse, which loses nothing. It goes ahead of core-ts's D11, which is undecided. |
| Q10 | With Phase A left out, do its root warning and `extract`'s refusal move into Phase B? | **Yes, both, and the second-container warning goes.** Section 3.2's roots that stay as today would otherwise read empty with nothing said, and `or-40` and `or-41` still give `""`. After Phase C, the second-container warning has only unmeasured shapes left. About an hour, though the code isn't agreed yet with core-ts and docx4j. |

## 8. Phase B: implementation notes (2026-11-07)

### 8.1 What was built

- **`openpackaging/repeated.py`:** `read_root`, `repair`, `triggers`, and the two records, `ReadAsWord` (`part.read_as_word`) and `UnexpectedRoot` (`part.unexpected_root`). 278 lines with its docstrings and the two records, over section 5's estimate of 60 for code alone.
- **`XmlPart._read_root`:** one read of the start tag where a part sets `reads_root`, which only `MainDocumentPart` does. It runs before `_parse` from both `_unmarshal` and `source_tree`, so a save's check reads the source as the parse did. A root read this way resets the part's source declarations, or `mc:Ignorable`'s prefix would go.
- **`read_through.repair_bytes`** runs `repeated.repair` first and returns a fifth value, `joined`.
- **The refusal:** `RootRefused` becomes the part's `Docx4JException`, kept as `part.unreadable`. `misplaced()` reports `part.root_refused`, a new finding, since the XML is well formed.
- **The rest:** `TextWarning.PART_UNEXPECTED_ROOT`, the root's name in `detail`, `extract` and `body_of` on a package with no main part, `canon.compare(..., main_part=True)` and its `read-as-word` category, `roundtrip.py`.
- **`tests/fixtures/repeated/`:** 48 inputs, 37 of Word's saves and their text saves, about 0.5 MB, `or-14-comments`' from its rerun.

### 8.2 What the build settled beyond the design

- **Two `w:body` are in Phase B,** since two `w:document` need the join. `rc-10` comes through the skipped report and `repair_bytes`, and a branch's documents through `read_root`.
- **A root read as Word reads it never needs a second `ReadAsWord`,** because `read_root` joins as it goes. Phase C will need one per part once a repeated `w:pPr` can stand under such a root.
- **Any other root is now written as a `w:document` with no body,** so `foo` holding a `w:body` yields nothing, as REQ-320 says. It used to bind. What it held is listed with its real parent, `p` under `w:hdr`, not `w:document`.
- **The clean cost is one start tag.** `etree.XMLPullParser` needs `close()` on a short part, whose start tag otherwise waits in the parser: a test on `<a/>` found that.
- **A later `w:body` with no `w:sectPr`** leaves the last one stated, moved to the end. Built ahead of the measurement, and `rt-16` and `rt-17` then held it, with `rt-15` for a `w:sectPr` in a `w:pPr`. All three are in the fixtures.
- **`rw-24` reads as Word does, an empty document,** by the user's ruling of 2026-11-07 ("rw-24 against REQ-305: as Word"). REQ-305 in a body (`rw-25`) was ruled after it, and that change isn't this CR's (section 2.2).

### 8.3 Verified

- **The whole suite:** 11,775 passed, 8 skipped, at `6851323`. After `detail` and the `rt-*` fixtures, the fast suite: 9,676 passed, 8 skipped.
- **`scripts/roundtrip.py`:** clean, 50 of 50 parts, 0 differences, 0 skipped.
- **21 mutants** of `repeated.py`, `xml_part.py`, `read_through.py` and the extractor: every one caught. Four survived the first run and a test or a removal followed: a branch's second element, text after its document, a prefix only the branch declares, and an unused merge, now gone.
- **Word's own files:** every file Word read matches its words and its section, a save reads again with nothing to say, and `Body` writes to each.
- **No Word check yet.** This changes what a save writes, so Phase D's is due, at the user's word.

### 8.4 Left to others

- **The WD:** Phase B is built to 0.96's draft of REQ-317 to REQ-321 and REQ-322's section rule. Its three review points were taken: `wordDocument`'s namespace, the section sentence, and the root's name in `detail`, which `TextWarning` now holds.
- **A `w:body` in a `w:body`** (`nd-20` to `nd-24`): Word reads the words in place, and here they're lost, listed in `part.skipped`. The WD requires nothing of it yet, and this CR doesn't cover it.
- **Q6:** `w:txbxContent` comes in reverse order with three and four copies (`rt-12`, `rt-13`), which meets one of its two conditions. A second shape is still missing.
- **The codes:** `part.unexpected_root` is decided here (the user, 2026-11-07: "makes sense"). `part.read_as_word` is a proposal, and `part.root_refused` is the validator's.
- **CR-033:** the save after `.contents` of a root that yields nothing still writes an empty `w:document`.

### 8.5 docx4j's `madeup-root` round (2026-11-07, resolved in section 9)

docx4j's round tested how far Word reads an unknown root: 20 files, each run twice, in the share's `bridge/portfolio/` as `mr-*`, with the table in its CR-079 section 1.1. Compared here from the raw saves at `af1d577`, Phase B agrees with Word on 13 in words or refusal, and section 3.2's rules are too wide or too narrow on the rest.

| File | Root | Word | Phase B |
|---|---|---|---|
| `mr-14` | `w:madeUp` holding a bare `w:r` | Refused | Empty, under `part.read_as_word`, which is false |
| `mr-15` | `w:madeUp` holding a `w:body` | The words, and that body's section | **Empty**, under `part.read_as_word`, which is false: the wrap leaves a `w:body` in a `w:body` |
| `mr-11` | `w:madeUp` with a `w:sectPr` directly in it | The words, the `w:sectPr` not read | The words, and it reads that `w:sectPr` |
| `mr-20`, `mr-21` | `z:madeUp`, `madeUp`, each holding a `w:p` | Refused | Empty, `part.unexpected_root` |
| `mr-31` | `w:tc` | Refused | Empty, `part.unexpected_root` |
| `mr-30`, `mr-32` | `w:sdtContent`, `w:p` | The words | **Empty**, `part.unexpected_root` |

Two `w:p`, a `w:tbl`, a block `w:sdt`, a further unknown element, paragraph properties and a style all read as Word does (`mr-10`, `mr-12`, `mr-13`, `mr-17`, `mr-18`, `mr-50`). So do an empty `w:madeUp` (`mr-16`), `z:madeUp` with `z` in `mc:Ignorable` (`mr-22`), and a branch's `w:madeUp` (`mr-40`).

Each rests on one file. WD 0.97's REQ-318 to REQ-320 don't have these yet: docx4j has proposed wording, and nothing is ruled, so the code stays as built. `mr-15` and section 8.4's `w:body` in a `w:body` are one fix.

## 9. Phase B's amendment and Phase C: implementation notes (2026-11-07)

### 9.1 What was built

- **The roots, to WD 0.98's REQ-318 to REQ-320:** section 3.2. `_refused` takes `w:tc`, a bare `w:r` in an unknown root and a `w:p` in a foreign one, and honours the root's `mc:Ignorable`. `_read_through` is new: a `w:p`, a `w:sdtContent`, and what an unknown root may hold. Section 8.5's seven differences are gone: `mr-*` agrees with Word on all 20 files, in words and refusals.
- **A `w:body` in a `w:body`:** `_flatten`. `nd-20` to `nd-24` match Word's words and section. It's what makes `mr-15`'s body the body, by another route, and the WD doesn't require it.
- **Phase C, REQ-322:** `repeated.repair` joins `w:sdtContent` in a block `w:sdt`, merges `w:pPr` and `w:rPr`, and keeps the last `w:jc`, `w:b`, `w:sz`, `w:sectPr` and `w:tblPr`. `triggers` names the nine children measured.
- **One warning a part:** `with_root`, gone in section 8, came back, since a repeated property can now stand under a root read.
- **`canon`:** a merge appends the later `w:pPr`'s children and a save writes the schema's order, so under a join `compare` sorts the children of `w:pPr` and `w:rPr` on both sides.
- **Fixtures:** `tests/fixtures/repeated/` adds `mr-*` and `nd-20` to `nd-24`.

### 9.2 What the build settled beyond the design

- **`w:sectPr` in a join is counted**, so `rc-10` reports 2 children read as one: a body and a section.
- **A block `w:sdt` is one in a `w:body` or a `w:tc`.** Only the body is measured. A run-level one and one in a `w:tr` stay as they are.
- **What a merge replaces is listed,** so `rc-20` is refused under `strict=True` and `rc-21` loads.
- **`w:orient` differs from Word on three files.** On `rt-10`, `rt-40` and `rt-50`, Word's save has the last `w:pgSz`'s size with `w:orient="landscape"` from an earlier copy, which the last copy doesn't state. Here the earlier section goes whole, so that attribute goes too. Text is the same, and REQ-322 stands. It hints that Word merges `w:sectPr`, which section 3.3 left open. Left as it is unless a round says more.
- **`mce.nothing_selected` isn't here.** WD 0.98 has REQ-305 raise it, at a root too, so `rw-24` should. It belongs with REQ-305's change, which waits for docx4j's list.

### 9.3 Verified

- **The whole suite:** 11,966 passed, 8 skipped, 1 failed. The failure is `test_the_understood_list_is_docx4js_verbatim`, which compares this package's list with the sibling's working copy, and docx4j's has an uncommitted change (127 URIs). It's REQ-305's work, and the list here doesn't change in this CR.
- **`scripts/roundtrip.py`:** clean, 50 of 50 parts, 0 differences, 0 skipped.
- **28 mutants** of `repeated.py`: every one caught. Two survived the first run, and a test and a simpler join followed.
- **Word's own files:** every file Word read matches its words, its section and its properties, and every file Word refused is refused.

### 9.4 The Word check: passed

Through the Word bridge on 16.0.20326.20232, built over `f7c5652`, `bridge/staging/cr039-word-check/`, with `out/cr039-word-check/build.py` and `compare.py` (git-ignored). 46 inputs, each run twice: this package's saves of 42 fixtures, read and saved, and 4 with a paragraph added through `Body`.

- **Word opened all 92 silently,** with no dialog. That includes `rw-22`, whose source Word had to repair.
- **The two runs agree** on all 46, in words, sections and properties, leaving rsids aside.
- **All 46 match Word's own reading of the source,** compared from the raw saves: this package's save, opened by Word, gives what Word gave from the original file. The 4 edited files hold the new paragraph once, and are otherwise the same.

So a save of each shape is an ordinary document Word reads as it read the source. The roots Word refuses aren't in the set, since a save writes their bytes as read.
