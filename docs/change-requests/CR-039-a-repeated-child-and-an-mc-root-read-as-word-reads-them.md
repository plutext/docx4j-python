# CR-039: A repeated child and an `mc:AlternateContent` root, read as Word reads them

**Status:** Proposed 2026-11-06 (the user: "write the CR - we may defer implementation of it indefinitely..."), on the objects session's message of that day and section 7.15 of the portfolio's `text_extraction_behaviour_model.md` (uncommitted when read). Nothing is built, nothing waits on it, and it isn't for the coming release. Whether it stays proposed or moves to deferred is Q7. Amended 2026-11-07 on core-ts's message of that day, which relays the user's word that this repository consider its CRs: section 2.4 is new, and sections 3.2, 3.4, 3.5 and 3.6 and Q2 to Q6 changed, with Q9 new. As first written, Phase B refused every root but `w:document`, and Word opens several of them.

**Depends on:** nothing to build. Phases B and C want a WD requirement first (Q1), since the WD says nothing about either shape. **CR-033** is the save's half (section 2.3). Nothing here waits on phase 0 of the two counterparts (Word check 81), which isn't scheduled and has no date: core-ts's CR-028 may be deferred indefinitely too. Phase B rests on rounds already run, and Phase C builds what's measured. Phase 0 would only widen C.

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

`rw-25`, a `Choice` not understood with no `Fallback` in the body, is a different question. Here it reads the `Choice`, as REQ-305 requires, and Word shows neither. Section 7.15 leaves that to the user, and this CR doesn't touch it.

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
- **The roots Word refuses:** `w:footnotes`, `w:endnotes`, `foo`, `z:foo` and `w:foo` holding a `w:document`, `w:wordDocument` and `pkg:package`. Refused as a part not well formed is (CR-031, and Q4): `Docx4JException` naming the part, kept as `part.unreadable`, the load lazy, an untouched save writing its bytes. They're listed by name and shape, never as "anything else", because `w:madeUp` holding a `w:p` opens and what decides isn't known.
- **Every other root stays as it is today,** with section 3.6's warning: empty text, its children in `part.skipped`, `Body` refusing. That covers `w:hdr`, `w:ftr`, `w:styles` and `w:comments`, whose words aren't a reading to copy (core-ts CR-028 section 4), a bare `w:body` and `w:madeUp` (Q9), and anything unmeasured, among them `foo` holding a `w:body`.
- **Main part only.** Every measurement is `word/document.xml`. A header's or a note's root stays as it is unless a round measures it (core-ts's D5).
- **The root's name is read once,** in this one part, since `<foo/>` leaves no report and `<foo><w:body>` an empty one.

### 3.3 The repeated child (Phase C)

| Rule | Elements | Measured |
|---|---|---|
| Join in order | `w:body`, `w:sdtContent` | A body in a `w:document`, a block `w:sdt` |
| Join, the second first | `w:txbxContent` | A VML box only (Q6) |
| Merge, the later value winning | `w:pPr`, `w:rPr` | One property each way, one level deep |
| Keep the last | `w:jc`, `w:b`, `w:sz` | One file each (Q5) |
| Keep the later | `w:sectPr`, `w:tblPr` | Same properties in both copies |

Joining two bodies drops the first body's final `w:sectPr`: Word's save holds one section, the second's. A `w:sectPr` in a paragraph's `w:pPr` stays. Only one section per body is measured.

Everything else stays as it is today: the first kept, the later listed. That covers a run-level or cell-level `w:sdtContent`, a DrawingML box, and a second `w:tcPr` or `w:trPr`. Phase 0's list covers most of those, and whether "the last" means the whole element or each attribute (`w:ind w:left` then `w:ind w:right`), since `rc-30` to `rc-33` hold one attribute each.

### 3.4 The report (Q3)

One warning a part with a count, on `part.warnings` and `package.warnings`, as `BrOutsideRun` is, since a file read as Word reads it has lost nothing. What goes whole goes in `part.skipped` beside it: a value a later one replaced, an earlier `w:sectPr`, a branch not taken. So `strict=True` refuses those, and a part that only joined loads under it, as CR-030's repaired part does.

That's core-ts CR-028 section 3.3. It replaces this section's first draft, one `SkippedNode` for each element, under which a join that loses nothing would have been refused at a save. The code is agreed with core-ts and docx4j first (their D8).

### 3.5 The save and the round trip

A part read and repaired saves in that shape, Word's own save's. An unread part stays byte for byte. `canon.compare(..., repaired=True)` runs the source through the same pass and counts `repeated-child`, every other category still held to zero. That's a fourth exception in `CLAUDE.md`'s round-trip rule.

Once CR-033 is built, its gate refuses a re-marshalled part with `part.skipped` entries unless `repair="word"`. A join alone leaves none, so that part saves. A part that dropped a value or a branch needs the switch, which is core-ts's D3.

`wml()` and the fragment route stay as they are: a fragment is the caller's code, and `wml()` refuses what it can't bind.

### 3.6 A warning (Phase A, Q2)

If the reading waits, the result can at least say so. Three small changes, none of which reads anything new:

- **A root that isn't `w:document`:** one warning on the part at the parse, on `part.warnings`, `package.warnings` and the text result, naming the root. This is core-ts's D11, and `part.unexpected_root` is the code proposed here, to be agreed with core-ts and docx4j. It needs section 3.2's read of the root's name, so that read comes with this phase.
- **A second container:** `extract` raises `content.not_read`, selector the part's name, where `part.skipped` holds `w:body`, `w:sdtContent` or `w:txbxContent`, the names that carry words. It stays silent for the rest, which keeps the WD's note that what REQ-308's section drops carries no warning (CR-026 D4). That note is about content Word drops too, and these hold content Word shows.
- **A package with no main part:** `extract` refuses, as core-ts's `Texts.of` does, and `body_of` raises `ContentError`.

REQ-801 needs a clause first: `content.not_read` is defined today for a chart, a SmartArt part and an embedded package, and the new code isn't in its list. Phases B and C remove most of what raises the first two.

## 4. Tests and oracles

- **Word's files are the oracle:** the 13 repeated and 13 root inputs, core-ts's 12 `or-*` files named in section 2.4, their controls and Word's saves, as fixtures under `tests/fixtures/repeated/` with a `SHA1SUMS` and a README, as `check76/` is. They come to under 1 MB, and today they exist only on the share and in `docx-tool-identifier`'s git-ignored `runs-local/` (Q8).
- **Text:** equal to Word's `.txt`, and to Word's save for a box.
- **Properties:** the target paragraph's `w:pPr` and `w:rPr`, the `w:tblPr` and the `w:sectPr` equal to Word's save, by `canon`.
- **Refusal:** the nine roots Word refuses, each named, with an untouched save keeping its bytes.
- **Roots Word opens:** each stays as it is today, with the warning.
- **Nothing else moves:** the corpus round trip, `check76/`, `part.skipped` empty on `samples/`.
- **Phase A:** each warning on its shapes, none on a skipped property or an unknown element, and `extract` refusing `or-40` and `or-41`.

## 5. Phases

Each phase stands alone, and any of them can be left out.

| Phase | Work | Cost |
|---|---|---|
| A | The warnings, once REQ-801 allows them, and `extract` refusing a package with no main part | ¼ day |
| B | The root: `mc:AlternateContent` read through `branch_taken`, the roots Word refuses refused | ½ day |
| C | The repeated child: `repeated.py`, canon, `roundtrip.py`, the fixtures and tests | 1 day |
| D | The records and the Word check of B's and C's saves, through the bridge at the user's word | ¼ day |

A is worth building only if B and C wait.

**What B and C add.** Estimates, set against `read_through.py` (364 lines) and its 1,128 lines of tests:

- **Code:** about 60 lines for B and 150 for C, in `repeated.py`, plus about 20 in `XmlPart._parse`, `canon.py` and `roundtrip.py`. Tests about 150 and 300 lines. No change to the model, the generator or the fork.
- **A clean file, Phase C:** nothing. The trigger is the skipped report, which is already read, and the module is imported only when it fires.
- **A clean file, Phase A or B:** one read of the main part's root name, the only cost every file pays, and only when the part is read. Timed at about 23 µs with a pull parser fed 2 KB at a time; `etree.iterparse` took 0.1 to 0.45 ms, since it reads 32 KB first. The skipped report can't replace this read, as section 2.2's `<foo><w:body>` shows.
- **A file the pass repairs:** that part parses twice, the route CR-030 and CR-037 already take, and if they fire too it's still one reparse. About 2.1 times the part's parse, summed from its steps on `samples/` rather than run: `Symbols.docx` (772 KB) 200 ms to about 418 ms, of which the pass's own lxml parse, walk and write are 18 ms. The lxml tree is transient.

## 6. Risks

- **One file per rule.** Each row of section 3.3 rests on one small file, two runs, with no styles part. A real document might join differently, above all a body with headers and several sections.
- **A rule beyond the measured names.** `w:vanish` given twice decides whether words show, and nothing here measures it (Q5).
- **Word has no one rule for the root.** Section 3.2's list is what was measured, and a shape off it gets today's reading.
- **`w:txbxContent`'s order is unexplained.** Copying it pins something nobody understands.
- **A root refused where today gives empty text.** A caller who got `""` gets an exception. Word refuses the same files, so that's the honest answer, but it's a change.
- **The evidence can go.** See Q8.

## 7. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Does the WD take these rules first? | **Yes, before B and C.** It's the shared rule for three engines, and without it this package would depart from the other two alone, though towards Word. The portfolio would draft it from section 7.15. |
| Q2 | If the reading waits, are the warnings worth having, and under which codes? | **Yes:** `part.unexpected_root` for a root and `content.not_read` for a second container, agreed with core-ts and docx4j first. It's the cheapest way to stop a silent empty text. |
| Q3 | How is what a pass does reported? | **One warning a part with a count, and `part.skipped` for what goes whole**, as core-ts's section 3.3 has it. A strict caller then loads a file that only joined. |
| Q4 | Are the roots Word refuses refused here, rather than read as an empty document? | **Yes, those measured, by name and shape.** Every other root stays as it is, with the warning, since Word opens several and what decides isn't known. |
| Q5 | Does "keep the last" cover every property in `w:pPr` and `w:rPr`, or only the three measured? | **Those measured, and any phase 0 adds if it runs**, as core-ts's D6. `w:vanish` is on its list now, since it decides whether words show, so it stays unread until then. |
| Q6 | Follow Word's second-first order in a VML box? | **Not without a second shape and three copies**, as core-ts's D7, which phase 0 would give. Until then a second box stays as it is, and Phase C goes without this row. |
| Q7 | Does this stay proposed, or move to deferred now? | **The user's.** The registry entry says proposed, with the note that it may wait indefinitely. |
| Q8 | Keep the Word files in this repository now, even if nothing is built? | **Yes.** Rebuilding the inputs is easy, but Word's saves need the bridge again. It's under 1 MB of fixtures, and a deferred CR can outlive a share. |
| Q9 | Read a bare `w:body` and `w:madeUp` holding a `w:p` as Word does? | **Not until a file holds one**, as core-ts's D11. They're the two roots Word opens silently with the words in the body, and the warning says this package shows none. |
