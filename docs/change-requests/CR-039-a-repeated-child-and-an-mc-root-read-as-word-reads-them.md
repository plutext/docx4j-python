# CR-039: A repeated child and an `mc:AlternateContent` root, read as Word reads them

**Status:** Proposed 2026-11-06 (the user: "write the CR - we may defer implementation of it indefinitely..."), on the objects session's message of that day and section 7.15 of the portfolio's `text_extraction_behaviour_model.md` (uncommitted when read). Nothing is built, nothing waits on it, and it isn't for the coming release. Whether it stays proposed or moves to deferred is Q7.

**Depends on:** nothing to build. Phases B and C want a WD requirement first (Q1), since the WD says nothing about either shape. **CR-033** is the save's half (section 2.3).

**Counterpart:** none. docx4j and core-ts lose a body too, keeping the second where this package keeps the first, and whether they change is the user's call there. The objects session measured all three (section 7.15's table).

## 1. Why

Word opens a `word/document.xml` that gives a child twice, or whose root is `mc:AlternateContent`, and reads what it holds, with no message on all but one file (`rw-22`, section 2.2). This package keeps the first of a repeated child and reads nothing under such a root. So the text differs from what Word shows, and the ports differ from each other: the same two-body file gives different words here and in docx4j.

Nobody is known to write these files. None of the 445 local `.docx` files outside `check76/` has any skipped entry (CR-037 section 2), so this is correctness on a rare input, which is why it can wait. Two things stop it being only a curiosity:

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

The parse is handed the class, so the root's name is never compared: any root binds as `Document`, its children are skipped and the tree is empty. The facade raises `ContentError` on all thirteen, since a document with no `w:body` has no body.

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

## 3. Design

### 3.1 The pass

`openpackaging/repeated.py`, a pass over lxml at `XmlPart._parse`'s seam, beside `read_through` and `breaks` and through the same repair and one reparse. Order: the root, the repeated children, then CR-037's and CR-030's passes on the same tree, since a `w:document` taken from a branch may hold what they repair.

**The trigger** is the skipped report, as for both passes: a clean part pays the test of an empty report. Every shape in section 2 is already in it. `parent_qname` can be a type name (`CT_PPr`) in whichever namespace the metadata met first, so the test compares local names.

### 3.2 The root (Phase B)

- **`mc:AlternateContent`:** replaced by what `branch_taken` answers. One `w:document` becomes the root; two are joined as two `w:body` are (section 3.3); `None` leaves an empty `w:document` holding an empty `w:body`, which the facade can read.
- **Any other root than `w:document`:** refused as a part not well formed is (CR-031, and Q4): `Docx4JException` naming the part, kept as `part.unreadable`, the load lazy, an untouched save writing its bytes.
- **Main part only.** Every measurement is `word/document.xml`. A header's or a note's root stays as it is. The refusal needs one read of the root's name, in this one part, since a `<foo/>` leaves no report to trigger on.

### 3.3 The repeated child (Phase C)

| Rule | Elements | Measured |
|---|---|---|
| Join in order | `w:body`, `w:sdtContent` | A body in a `w:document`, a block `w:sdt` |
| Join, the second first | `w:txbxContent` | A VML box only (Q6) |
| Merge, the later value winning | `w:pPr`, `w:rPr` | One property each way, one level deep |
| Keep the last | `w:jc`, `w:b`, `w:sz` | One file each (Q5) |
| Keep the later | `w:sectPr`, `w:tblPr` | Same properties in both copies |

Joining two bodies drops the first body's final `w:sectPr`: Word's save holds one section, the second's. A `w:sectPr` in a paragraph's `w:pPr` stays. Only one section per body is measured.

Everything else stays as it is today: the first kept, the later listed. That covers a run-level or cell-level `w:sdtContent`, a DrawingML box, and a second `w:tcPr` or `w:trPr`.

### 3.4 The report (Q3)

One `SkippedNode` for each element a pass takes, through CR-030 section 11's seam, so `part.skipped` lists it and `strict=True` refuses as it does today. That's CR-037's rule, and it changes nothing for a strict caller.

### 3.5 The save and the round trip

A part read and repaired saves in that shape, Word's own save's. An unread part stays byte for byte. `canon.compare(..., repaired=True)` runs the source through the same pass and counts `repeated-child`, every other category still held to zero. That's a fourth exception in `CLAUDE.md`'s round-trip rule.

Once CR-033 is built, its gate refuses a re-marshalled part with `part.skipped` entries unless `repair="word"`, so a part this pass read would need that switch, as a part CR-037's did. Nothing is lost by a join, so CR-033 may want to tell the two apart. That's for whichever is built second.

`wml()` and the fragment route stay as they are: a fragment is the caller's code, and `wml()` refuses what it can't bind.

### 3.6 A warning (Phase A, Q2)

If the reading waits, the text result can at least say so. `extract` raises `content.not_read`, selector the part's name, where the part's `part.skipped` holds `w:body`, `w:sdtContent`, `w:txbxContent`, `mc:Choice`, `mc:Fallback` or `w:document`, the names that carry words. It stays silent for the rest, which keeps the WD's note that what REQ-308's section drops carries no warning (CR-026 D4). That note is about content Word drops too, and these names hold content Word shows.

`content.not_read` is defined today for a chart, a SmartArt part and an embedded package, so REQ-801 needs a clause first. Phases B and C remove most of what raises it.

## 4. Tests and oracles

- **Word's files are the oracle:** the 13 repeated and 13 root inputs, their controls and Word's saves, as fixtures under `tests/fixtures/repeated/` with a `SHA1SUMS` and a README, as `check76/` is. They come to about 0.5 MB, and today they exist only on the share and in `docx-tool-identifier`'s git-ignored `runs-local/` (Q8).
- **Text:** equal to Word's `.txt`, and to Word's save for a box.
- **Properties:** the target paragraph's `w:pPr` and `w:rPr`, the `w:tblPr` and the `w:sectPr` equal to Word's save, by `canon`.
- **Refusal:** the six roots Word refuses, each named, with an untouched save keeping its bytes.
- **Nothing else moves:** the corpus round trip, `check76/`, `part.skipped` empty on `samples/`.
- **Phase A:** the warning on each shape, and none on a skipped property or an unknown element.

## 5. Phases

Each phase stands alone, and any of them can be left out.

| Phase | Work | Cost |
|---|---|---|
| A | The warning, once REQ-801 allows it | ¼ day |
| B | The root: `mc:AlternateContent` read through `branch_taken`, any other root refused | ½ day |
| C | The repeated child: `repeated.py`, canon, `roundtrip.py`, the fixtures and tests | 1 day |
| D | The records and the Word check of B's and C's saves, through the bridge at the user's word | ¼ day |

A is worth building only if B and C wait.

## 6. Risks

- **One file per rule.** Each row of section 3.3 rests on one small file, two runs, with no styles part. A real document might join differently, above all a body with headers and several sections.
- **A rule beyond the measured names.** `w:vanish` given twice decides whether words show, and nothing here measures it (Q5).
- **`w:txbxContent`'s order is unexplained.** Copying it pins something nobody understands.
- **A root refused where today gives empty text.** A caller who got `""` gets an exception. Word refuses the same files, so that's the honest answer, but it's a change.
- **The evidence can go.** See Q8.

## 7. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Does the WD take these rules first? | **Yes, before B and C.** It's the shared rule for three engines, and without it this package would depart from the other two alone, though towards Word. The portfolio would draft it from section 7.15. |
| Q2 | If the reading waits, is the warning worth having, and under which kind? | **Yes, under `content.not_read`**, with REQ-801 extended. It's the cheapest way to stop a silent empty text. A new kind buys nothing, and the selector rule already fits. |
| Q3 | Is what a pass takes listed in `part.skipped`, so that `strict=True` refuses? | **Yes**, as CR-037 does. Strict callers see no change. |
| Q4 | Are the roots Word refuses refused here, rather than read as an empty document? | **Yes.** Empty text with a success is the worst of the options, and `foo` isn't a Word document. |
| Q5 | Does "keep the last" cover every property in `w:pPr` and `w:rPr`, or only the three measured? | **Every one, after one more Word round over the schema's property list.** Without that round, only the three. `w:vanish` is the one that matters. |
| Q6 | Follow Word's second-first order in a VML box? | **Yes**, pinned as measured, for the VML box only. It held on both runs. A DrawingML box stays as it is until measured. |
| Q7 | Does this stay proposed, or move to deferred now? | **The user's.** The registry entry says proposed, with the note that it may wait indefinitely. |
| Q8 | Keep the Word files in this repository now, even if nothing is built? | **Yes.** Rebuilding the inputs is easy, but Word's saves need the bridge again. It's about 0.5 MB of fixtures, and a deferred CR can outlive a share. |
