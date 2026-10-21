# CR-012: Keep `mc:AlternateContent` at every level, and type what its branches hold, as docx4j does

**Status:** Proposed 2026-10-19, at the user's request ("docx4j and docx4j-generated-objects-ts both
retain all mc content ... do we? if not, please draft a CR"). Section 1 answers the question.
**Revised 2026-10-21** at the user's request ("assess CR-051 and modify your CR-012 as
appropriate"), after docx4j's CR-051. Section 1.1 corrects section 1's answer: an element at block,
row or cell level is not kept, and a content control directly in a branch loses its runs. Section
3.1 measures both on Jason's probes. A new Phase A (section 4.1) ports CR-051's phase 1. The 61
roots, which were Phase A, are now Phase B, and VML is Phase C. The questions of section 7 carry
recommendations. **Revised again 2026-10-21** for docx4j CR-051's D5 as Jason revised it the same
day (docx4j-25's note; docx4j's record of it not yet committed): docx4j's load-time mode defaults to
`preserve` from 17.4.0, and 17.3.x keeps `word`. Q7's recommendation moves to (b), `preserve` by
default, and section 4.1's readers' audit grows to match. **Decided 2026-10-21** (the user: "CR-012
questions section: i accept each of the recommendations"): Q1 to Q8 as recommended, Q7 on its
revised answer, `preserve` by default. By Q6, Phase A goes first and now. It starts once CR-002
section 16, which changes the same `XmlPart._unmarshal`, has landed. B follows A, and C waits for
its own go (Q4). **Phase A implemented the same day** (section 9), the port of docx4j CR-051 phase
1: the schema copy refreshed to docx4j `2853b7280`; the fork's stage 11 (a branch's child typed by
its container, and a generator fix that makes the output independent of the bytecode cache);
`LoadOptions(mc_block_level=)`, `preserve` by default; the readers reading Word's branch at block
level, an edit inside a kept branch refused; the revisions twins; and docx4j's 19 probes with Word
365's 14 saves. In `preserve` mode all 19 probes save canonically identical; in `word` mode all 14
read as Word's save. The acceptance artefacts are unchanged and CR-011's ratchet does not move.
**The Word check of three probe saves passed the same day** (the user: "__tmp/python-cr012a/ are all
good"; 9.12): **Phase A done**. Row and cell level are passed by; docx4j's phase 2 (`9db1a3cbd`),
which reads them, is ported as Phase A2 (section 4.4), started the same day at the user's go ("make
the CLAUDE.md edits, then do phase A2"). **Phase A2 implemented the same day** (section 10): a
branch's `w:tr` and `w:tc` typed by their container (the fork's stage 11 item 27, tests only); the
text model, the revisions engine (twins, the table-removal rule) and the content API (read-only) at
row and cell level; docx4j's phase 2 tests; the five row and cell probes read `w16sdtfl` and save
canonically identical in `preserve` mode; acceptance unchanged. **The Word check of six saves passed
on 2026-10-22** (the user: "python-cr012a2 each pass"; 10.9): **Phase A2 done**. The port of docx4j
`60bd4b9ad`, a twin `w:cellDel` or `w:cellIns` in a branch Word does not read resolved in its own
branch, landed the same day at the user's go ("go on the port"; 10.10), and its Word check of three
saves passed (the user: "python-cr012a2-twin passes"). The port of docx4j `901a922a3`, a table's
property records read and resolved through every branch, landed the same day at the user's go ("port
the fix, then start phase b"; 10.11), and its Word check of three saves passed (the user: "the Word
check of __tmp/python-cr012a2-records/ passes"). **Phase B started 2026-10-22** at the same go, its design
revised first (section 4.2): Phase A's fork hooks take the root-element table, so the fork needs no
code. The port of docx4j `84b202c5a`, a table change's text read from Word's branch and a merged
grid's spans counted along every branch before any is written, landed the same day at the user's go
("go on the port"; 10.12), and its Word check passed (the user: "python-cr012a2-grid passes").
**Phase B implemented the same day** (section 11): docx4j's 61 WordprocessingML root elements,
derived from docx4j's classes into `codegen/names/root_elements.json` and `el_index.py`, typed in a
branch by the container's binding and elsewhere by their root class; over 422 documents built from
commits every `w:drawing` in a Choice is a `Drawing` and every `w:pict` in a Fallback a `Pict`, no
new canonical difference and no skipped report changed; the readers' audit kept the python-docx
facade's `inline_shapes` as python-docx reads them; the `text_of` readers now read a
paragraph-level branch as a run-level one (11.4, kept); parse +21% and save +35% on the
403-branch document; acceptance and the ratchet unchanged; docx4j's agreement run 0 of 22. **The
Word check of four saves passed on 2026-10-22** (the user: "cr012b okay", file 2 "fine in All
Markup"; 11.10): **Phase B done**. **Phase C started 2026-10-22** at the user's go ("do phase c"),
section 4.3 revised first, and was **implemented the same day** (section 12): docx4j's five VML
schemas at the schema copy's commit, with no patch, so VML is typed wherever it stands (under every
`w:pict` and `w:object`, header watermarks, numbering picture bullets, `settings.xml`'s shape
defaults); over the 422 documents nothing is newly skipped and no save gains a canonical difference;
VML's `t`/`f` values are a string enumeration, written as read, and `o:spt` and `o:gfxdata` keep
their spelling through the fork's stage 12 (`ParserConfig(lexical_types=)`), which joins stages 9 to
11 in needing a fork release; the readers' audit over 384 documents moved one answer,
`styles_in_use()` on `vml/textbox.docx`, clearly right by docx4j's traversal (12.8); no corpus
document reaches an island any more, and the islands stay for a text box below a wildcard node
(12.9); the import costs about 100 ms more (cold 1.12 s), and the 403-branch document parses about
15% and saves about 10% slower; acceptance, the ratchet and CR-010's Word-save comparisons are
unchanged; docx4j's agreement run 0 of 24 (12.13). **Its Word check (12.14): files 1, 2 and 4
passed; file 3, an OLE object, failed** (the user: double-clicking "does not open the embedded
document"): the save writes `w:object`'s `o:OLEObject` before the `v:shape` it names, an order
older than this phase that `v0.1.0` writes too, and two order-only variants confirmed it as the
cause (the user: both "correct").
**Depends on:**
- **Phase A:**
  - docx4j CR-051 phase 1 (`6687cc4b8`, `1cff2cd5c` on `VERSION_17_3_4`, pushed), whose
    schema admissions the schema copy takes;
  - docx4j's published understood list (`understood-namespaces.txt`, CR-046), which CR-011's
    Q2 also ports;
  - a change to the xsdata fork's runtime (section 4.1 item 2);
  - CR-002 section 16 (implemented 2026-10-21), which changes the same `XmlPart._unmarshal` and
    records the load options a private parse used, the record the mode joins.
- **Phase A2:** docx4j CR-051 phase 2 (`9db1a3cbd` on `VERSION_17_3_4`, pushed), whose readers
  it ports.
- **Phase B:** Phase A, whose fork stage 11 hooks (items 25 and 26) take the root table; no
  further fork code (section 4.2, revised 2026-10-22).
- **Phase C:** docx4j CR-048 (`acd0defc9` on `VERSION_17_3_4`, pushed 2026-10-18, to ship in
  17.3.4), whose VML schema this phase copies, at `2853b7280`, the schema copy's commit, where
  `xsd/vml/` is unchanged since `acd0defc9` (section 4.3, revised 2026-10-22); and the xsdata
  fork's stage 12, item 28 (`ParserConfig(lexical_types=)`, `d084a267`, local), which joins
  stages 9 to 11 in needing a fork release (section 12.6).

docx4j-python cannot release on either fork change before a fork release, as CR-011's Q3 also
needs (`RELEASING.md`: never against an unreleased fork).

**Counterpart:** docx4j's CR-051
(`docs/developer/change-requests/CR-051-mc-alternate-content-block-level.md`, accepted, and
phase 1 done on 2026-10-20). Its phase 3 hands off to this repository: its readers at block
level, and whether its load model wants the two modes. docx4j-generated-objects-ts CR-009
(`../docx4j-generated-objects-ts/docs/change-requests/CR-009-type-kept-mc-branches.md`,
implemented and released in 0.4.0 on 2026-10-18), the same change in the TypeScript object
model. docx4j-core-ts CR-010 (keep every `mc` branch), whose phase 2 asked for it. docx4j's own
`@XmlRootElement` classes and unscoped `ObjectFactory` `@XmlElementDecl`s, which its JAXB
context types in a lax wildcard.

## 1. The answer

**We retain all of it where the schema admits the element**: in a run, in paragraph content and
in `w:numPicBullet`. Both branches of every such `mc:AlternateContent`, with every element and
attribute, are kept on load and written back. This has been the default since CR-002 (section
5.6, lossless MCE), held by the round-trip tests over `samples/`. It was measured again for
this CR over 371 documents and 2,471 WordprocessingML parts (section 3).

**We don't type all of it.** A kept branch's children are a wildcard (`any_element`). The
parser types one only when the model declares that element globally. `w:r` and `w:p` are global
here, so they come back typed. `w:drawing`, `w:pict` and `w:t` are declared only inside their
parents, so in a branch they stay `AnyElement` trees, and so does everything below them. Across
the corpus that is 493 `w:drawing`, 487 `w:pict` and 4 `w:t`. Below `w:pict`, VML is never typed,
in a branch or out of one, because the schema copy has no VML.

docx4j types both: its classes are JAXB root elements, and the VML package is in its context.
docx4j-generated-objects-ts has typed both since 0.4.0 (its CR-009).

What the gap costs today:
- `LoadOptions(mce_preprocess=True)` is the only way to get a typed `Drawing` inside a choice
  (CLAUDE.md).
- CR-010 had to build "islands", which re-parse a text box's `w:txbxContent` out of the wildcard
  tree so that revisions inside it can be resolved (`model/revisions/islands.py`).
- Of CR-011's baseline failures, 49 are text boxes and `mc:Choice` content that the text paths
  cannot see.

### 1.1 Where we don't (found 2026-10-21)

docx4j's CR-051 measured Word 365 reading `mc:AlternateContent` at block, row and cell level as
it reads it in a run. That covers the body, a cell, a table, a row, a header, a note, a text box
and every kind of content control. Section 3.1 runs this package over the same probes, and finds
two defects.

1. **Block, row and cell level: skipped, then lost.** The schema copy, like docx4j's before
   CR-051, admits the element only in run and paragraph content and `w:numPicBullet`.
   - Anywhere else the parser reports it in `part.skipped`, and `LoadOptions(strict=True)`
     raises.
   - A part that is read and saved loses the element whole, every branch with it; a part
     never read is written back byte for byte, as always.
   - Before CR-051, docx4j resolved such an element to its Fallback. This package keeps none of
     it.
   - None of section 3's 371 documents has one; their skipped reports would have shown it. Word
     reads such an element but never writes one (CR-051 phase 0), which is why CR-021's survey
     of 835 documents found none either.
2. **A content control directly in a branch takes the block class.** `w:sdt` is global here as
   `SdtBlock`, so a `w:sdt` that is a direct child of a kept branch is typed as a block-level
   control, whatever its level. In a paragraph's branch, its runs are then skipped, and lost on
   save. docx4j had the same defect: lax binding by name bound every one as `CTSdtRow`. It fixed
   it in CR-051 by binding a control by its container's kind (`McBranchSdt`). Nothing in the
   corpus has one.

Section 3.1 also finds that `LoadOptions(mce_preprocess=True)` takes `w16cid` where Word 365 takes
`w16sdtfl`, on every newest-first probe. `UNDERSTOOD` lacks `w16sdtfl`, `w16sdtdh`, `w16du` and
`w16cex`, all of which docx4j's published list has. CLAUDE.md says `mce_preprocess` takes "the
branch Word would take"; for Word 365 on these probes it doesn't.

## 2. What docx4j declares that this model doesn't

The source is docx4j's generated classes, in the build in its checkout
(`docx4j-generated-objects/target/generated-sources/xjc`, built 2026-10-18 on `VERSION_17_3_4`).
They declare 639 element names globally: the `@XmlRootElement` classes (191, nested ones
included) and every `@XmlElementDecl` with no `scope`.

**WordprocessingML: 99 global names in docx4j, 37 here, 62 missing.**
- 61 of the 62 are already in this repository's schema copy, as docx4j's own
  `annox:annotate ... @XmlRootElement(name="...")` annotations in `schemas/wml/wml.xsd`. The
  generator ignores them, which is why they aren't global here.
- The 62nd, `w:sdtPr`, is a root in docx4j's generated classes but not annotated in the schema.
- **`w:sdtContent` is left out.** It maps to four classes (`CTSdtContentRun`, `CTSdtContentCell`,
  `CTSdtContentRow`, `SdtContentBlock`), as it does in docx4j and objects-ts.

That leaves **61**:

`w:abstractNum`, `w:alias`, `w:altChunk`, `w:annotationRef`, `w:bibliography`, `w:bookmarkEnd`,
`w:bookmarkStart`, `w:citation`, `w:commentReference`, `w:continuationSeparator`, `w:cr`,
`w:dataBinding`, `w:dayLong`, `w:dayShort`, `w:drawing`, `w:endnoteRef`, `w:equation`,
`w:fldChar`, `w:fldSimple`, `w:footnoteRef`, `w:group`, `w:hyperlink`, `w:ind`, `w:jc`,
`w:lastRenderedPageBreak`, `w:lock`, `w:lvl`, `w:lvlOverride`, `w:monthLong`, `w:monthShort`,
`w:noBreakHyphen`, `w:num`, `w:numPicBullet`, `w:numPr`, `w:outlineLvl`, `w:pBdr`, `w:pPr`,
`w:permEnd`, `w:permStart`, `w:pgNum`, `w:pict`, `w:picture`, `w:ptab`, `w:rPr`, `w:richText`,
`w:sdtPr`, `w:sectPr`, `w:separator`, `w:smartTag`, `w:softHyphen`, `w:spacing`, `w:sym`, `w:t`,
`w:tab`, `w:tabs`, `w:tbl`, `w:tblGrid`, `w:tblPr`, `w:tc`, `w:yearLong`, `w:yearShort`.

Each takes the class the `el` table already gives its name: `Drawing`, `Pict`, `RT` for `w:t`,
`RTab` for `w:tab`, `PHyperlink` for `w:hyperlink`, `CtPprBaseSpacing` for `w:spacing`, and so
on. This is the same set as objects-ts's 61. Two class choices are worth stating:
- `w:pPr`, `w:sectPr`, `w:tblPr` and `w:tblGrid` take docx4j's root classes (`PPr`, `SectPr`,
  `TblPr`, `TblGrid`), not the `...Base` types, as objects-ts's list B also says.
- `w:spacing` takes `CtPprBaseSpacing`, the class of docx4j's root (`PPrBase.Spacing`,
  `@XmlRootElement(name = "spacing")`). objects-ts's list A gives `CTSignedTwipsMeasure`, the run
  property's type, which is not docx4j's root. A bare `w:spacing` never appears in a branch, so
  the difference is theoretical.

**Other vocabularies.** docx4j's other global names are in namespaces this model either doesn't
have (PresentationML, SpreadsheetML, VML, `o:`, `w10:` and the rest) or already declares
globally. The one exception is the custom properties' `property`, which never appears in a
branch and is left out.

## 3. Measured (2026-10-19)

The measurements used a prototype that edited no module. It added two hooks:
- **parsing:** in a wildcard that is not `processContents="skip"`, after `xsi:type` and the
  model's own global names, one of the 61 is parsed as a plain instance of its class;
- **saving:** an instance written without a field name takes its registered element name.

The corpus was 371 distinct documents: `samples/`, `tests/fixtures/`, docx4j-core-ts's
`test/fixtures/`, and docx4j's `docx4j-core-tests` and `docx4j-core` test resources. Every
WordprocessingML part (2,471) was loaded both ways, census-counted and marshalled. The scripts
are in `.claude/scratch/cr-012-mc/`: `globals.py`, `survey.py`, `deep.py`, `readers.py`,
`timing.py` and the pytest plugin `cr012_patch.py`.

**What a kept branch holds:**

| branch child | today | typed |
|---|---|---|
| `mc:Choice` / `w:drawing` | wildcard 491 | `Drawing` 491 |
| `mc:Choice` / `w:pict` | wildcard 1 | `Pict` 1 |
| `mc:Choice` / `w:t` | wildcard 4 | `RT` 4 |
| `mc:Fallback` / `w:pict` | wildcard 486 | `Pict` 486 |
| `mc:Fallback` / `w:drawing` | wildcard 2 | `Drawing` 2 |
| `mc:Choice` / `w:r` (global already) | `R` 12 | `R` 16 |
| `mc:Fallback` / `w:r` | `R` 8 | `R` 12 |
| `mc:AlternateContent` inside a wildcard tree | 4 | 0 (now typed, in a typed drawing) |

Three of the 61 occur in a wildcard anywhere in the corpus: `w:drawing`, `w:pict` and `w:t`. All
three are in an `mc:Choice` or `mc:Fallback`. No other wildcard in the model met one, so outside
the branches nothing changed.

**How deep the typing reaches.**
- **The Choice side is typed all the way down**, including the text box:
  - `samples/DrawingML_GraphicData_wps.docx`: 63 wildcard nodes become 61 typed objects, none left
    wildcard;
  - core-ts's `cr043-tracked.docx`: `Wsp`, `CTTextboxInfo`, `TxbxContent`, `RunIns`, `RunDel` and
    `DelText` are typed.
- **The Fallback side** gets a typed `Pict`, and everything below it stays wildcard: VML,
  `v:textbox` and the `w:txbxContent` inside it. That is 26 of 27 nodes in the first document and
  18 of 19 in the second. Typing that needs VML in the model (Phase C).

**Round trip.**
- **Canonical form:** typed and today's saves don't differ on any of the 2,471 parts.
- **Skipped reports:** the same as today's, part for part, so nothing is newly lost. The only
  skips in the corpus are Strict OOXML documents, a known gap (CR-002 section 12.3 item 7).
- **Bytes:** 20 parts differ in bytes only, in ways the round-trip rules already allow:
  - attribute order: a typed element writes the schema's order, so a `w:p`'s `w14:paraId` moves
    after its `w:rsid*`;
  - booleans: `wps:cNvSpPr txBox="true"` and `wps:bodyPr anchorCtr="false"` are written `1` and
    `0`, as on every typed element.

**Plain instances, not wrappers.** A first version of the prototype wrapped each typed child in
xsdata's `DerivedElement` to keep its element name. xsdata reads that wrapper as a type
substitution, so on save it wrote `xsi:type="w:CT_Drawing"` on `w:drawing` and
`xsi:type="w:CT_Picture"` on `w:pict`: 6 added attributes on `samples/`. The plain form, which is
what JAXB gives docx4j for a root class, writes none. A `w:r` in a branch is already a plain `R`
today, so the two forms stay consistent.

**Readers.** 23 corpus documents have `mc:AlternateContent` in `document.xml`. On all of them,
`get_text()`, `to_markdown()`, the inline-picture count, `outline()` and `get_tracked_changes()`
return what they return today. None of today's readers looks into the newly typed content. The
gain is for the readers that will: CR-011's text boxes, and the content API over kept branches
(core-ts CR-010 phase 3 is the counterpart).

**The suite with the prototype on.**
- **Fast:** 2,437 passed and 5 failed, against 2,442 passed without it. CR-010's fifteen
  Word-save comparisons pass, and its ten `xfail(strict)` pins stay xfailed.
  - Three `tests/openpackaging/test_mce.py` tests fail: `text_of` over a `w:t` that sits straight
    in a run's `mc:Choice` returns `"a"` for `"a-modern-fallback"`. `text_of` reads a wildcard
    `w:t` by name, and a typed one only where a run holds it. That is a traversal gap the change
    exposes, fixed in Phase B.
  - Two `tests/test_builders_phase_a.py` tests fail because they assert that `wps:wsp` inside a
    Choice is visited as a wildcard. Their expectation changes.
- **Slow:** 39 passed and 1 failed. The failure is the warm-up timing test (54.4 ms against
  52.1 ms), which passed 3 times out of 3 on rerun, with the prototype and without.

**Cost.** The prototype's hooks run on every node, so these numbers are an upper bound:
- `tc-p3-transforms.docx` (403 `mc:AlternateContent`, a 1.9 MB `document.xml`): parse 274 to
  339 ms (+24%), marshal 169 to 238 ms (+40%);
- the text-extraction probe (9 of them, about 90 KB): 4 to 7% on each.

### 3.1 Every level, on docx4j's probes (2026-10-21)

The probes are Jason's 25 `mc-fun` documents (`~/vbShares/Office 2016/mc-fun/`). docx4j's
`docx4j-core-tests/src/test/resources/mc-fun/` holds 16 of them, with 14 of Word 365's saves,
since CR-051 phase 1; the other nine are the pandemonium and plain variants and ED-009's three
run-level documents. Most carry ten Choices,
newest first, and a Fallback. Each was loaded, its main part read, and saved; the part was
compared with the source in `scripts/canon.py`'s form. The text is `get_text()`, by default and
under `mce_preprocess=True`. The Word 365 column is the branch Word kept in Jason's saves
(CR-051 section 1).

| level | probes | skipped | after a save | `get_text()` | `mce_preprocess` | Word 365 |
|---|---|---|---|---|---|---|
| run and paragraph content | 11 | 0 | identical | the branch's text absent | `w16cid` | `w16sdtfl` |
| block | 9 | 1 each | the element gone | nothing of it | `w16cid` | `w16sdtfl` |
| row and cell | 4 | 1 each | the element gone | nothing of it | `w16cid` | `w16sdtfl` |
| `mc-mixed` | 1 | 1 (the body's) | the body's gone | `SDTRUN_`, `HLINK_`, `INP_` absent | each `*_W14` | each `*_W14`, and `BODY_W14` |

- **The block rows:**
  - `mc-block`, `mc-block-newest-first`, `mc-block-nofallback` and `mc-block-pandemonium`;
  - `mc-sdt-block-newest-first` and `mc-tc-block-newest-first`;
  - the three text-box probes, `mc-txbx-block`, `mc-txbx-sdt-block` and `mc-txbx-tc-block`.
- **The row and cell rows:** `mc-row-`, `mc-sdt-row-`, `mc-cell-` and
  `mc-sdt-cell-newest-first`.
- **`get_text()` reads no kept branch at any level.** That is the content API's own reader gap,
  CR-011's (REQ-302), and outside this CR.
- **The `w:sdt` defect** shows on a fragment, a run-level control directly in a paragraph's
  `mc:Choice`. `wml(...)` raises `FragmentError`, naming the `w:r` that was skipped in
  `CT_SdtContentBlock`.

## 4. Proposal

### 4.1 Phase A: every level, as docx4j CR-051 phase 1

1. **The schema.** docx4j's three admissions (`EG_ContentBlockContent`, `EG_ContentRowContent`
   and `EG_ContentCellContent`, `1cff2cd5c`) come in by refreshing the schema copy from docx4j at
   that commit or later.
   - As CR-001 section 14.10 did: the closure re-copied verbatim, this repository's patches
     re-applied by marker, and `PATCHES.md` naming the commit.
   - Every block, row and cell container's content then admits `AlternateContent`, kept whole
     with wildcard branches as a paragraph's is today.
   - Regenerated, and `generate.sh --check`.
2. **A branch's container-dependent names.** In the fork's wildcard path, a child named `w:sdt`
   is typed by the kind of container the `mc:AlternateContent` stands in, as docx4j's
   `McBranchSdt` does:
   - `SdtBlock` in block content;
   - `SdtRun` in paragraph content;
   - `CTSdtRow` among a table's rows;
   - `CTSdtCell` among a row's cells.

   The same applies to `w:customXml` and `w:sdtContent`, should either become typed in a branch.
   - The change gets tests in the fork's `tests/fork/`, an entry in `CHANGES.md` and a row in
     `docs/UPSTREAM.md`.
   - Phase B's root table extends the same mechanism.
3. **The load-time mode** (CR-051 section 3.2 and its D5):
   `LoadOptions(mc_block_level="word" | "preserve")`, default `"preserve"` (Q7, revised from
   `"word"` with docx4j's D5). A package-level setting,
   `pkg.mc_block_level`, is read when each part is first unmarshalled, for packages made
   otherwise.
   - **`word`:** each element whose parent is a block, row or cell container is replaced by the
     content of the branch Word reads. The rule is ISO's: the first Choice whose `Requires` are
     understood, else the Fallback. It runs over docx4j's published understood list, matched by
     namespace URI, as docx4j's `McBlockLevel` and `McSplice` do (CR-051 section 3.2).
     - The outermost element goes first, and an element inside a discarded branch goes with it.
     - It is done on the lxml tree before the fork's parser runs, as `mce.py`'s preprocessor is.
       The selected content is then parsed typed by the ordinary path.
     - A part without the element pays one substring test.
   - **REQ-305:** an element with no understood Choice and no Fallback stays whole in both modes.
     One with no branch at all is removed in `word` mode and kept in `preserve`.
   - **`mc:Ignorable`:** the selected Choice's `Requires` prefixes are added to the root's
     `mc:Ignorable` where it doesn't already list them, with their declarations.
   - **The record:** each element taken out is kept on the part with its address
     (`part.resolved_alternate_content`). CR-011's `mce.branches_differ`, and a port of CR-050's
     inspector, then see every branch in `word` mode too.
   - **`preserve`:** nothing is applied, and the element stays whole, as item 1 makes it.
   - **Run and paragraph level are untouched** in both modes, as in docx4j: kept whole, as since
     CR-002.
   - **The round-trip rule's one exception.** In `word` mode a part holding a block-, row- or
     cell-level element isn't canonically identical when saved unchanged. It is what Word's own
     save writes (CR-051 phase 0), text boxes apart (CR-051 section 3.5).
     `tests/openpackaging/test_round_trip.py` and CLAUDE.md say so, and `preserve` keeps the
     rule.
4. **`mce_preprocess`** selects over docx4j's published list too, so that it takes the branch Word
   365 takes (Q8).
5. **The readers,** in `preserve` mode and for REQ-305's elements in either mode. An audit like
   CR-051's section 10 covers:
   - the content API, `traversal` and the revisions engine;
   - the Markdown exporter, `outline()`, `describe()` and the python-docx facade.

   Each reader either reads the branch `mce="resolve"` selects through the kept element, or
   passes the element by without error; none raises or miscounts. With `preserve` the default
   (Q7), a reader of text or structure must read Word's branch at block level, not pass it by.
   That covers the text model (closing CR-011 section 14.7 item 8), Markdown, `outline()`,
   `describe()` and the content API's collections. Otherwise the default load would hide
   content Word shows. An edit addressed inside a kept branch is refused (`address.read_only`,
   as a text box's is, CR-011 Q9) until a later CR maps edits between branches; a caller that
   edits there loads in `word` mode. Row and cell level are passed by, as docx4j's phase 2
   leaves them until then; recorded.
6. **The revisions engine at paragraph level.**
   - docx4j's revisions model reads a paragraph-level element's selected branch as text, and
     lists the other branches' changes as twins (CR-051 section 10.5 item 3,
     `TrackedChangesAlternateContentTest`).
   - This engine's `_has_text` reads such an element as holding no text (found in CR-010's U23
     port).
   - Ported here, held to docx4j's test.
7. **The probes:** docx4j's 16 with Word 365's 14 saves (the oracle of its
   `McBlockLevelProbesTest` and `McBlockLevelLoadTest`), copied into `tests/fixtures/mc-fun/`.

### 4.2 Phase B: the 61 WordprocessingML roots

**Revised 2026-10-22, before implementation, for what Phases A and A2 built.**
- **The fork has the two halves already.** Item 1 below, as first written, proposed a new table of
  root-element names in the fork's `XmlContext`. Phase A's fork stage 11 has since added both
  halves of it:
  - for parsing, a hook for a lax wildcard's child (`ParserConfig(wildcard_class=,
    wildcard_class_for=)`, item 25);
  - for saving, a table naming an object that no field names (`SerializerConfig(element_names=)`,
    item 26). Item 26 already names this phase as its next user.
- **So the table lives in this package, with no fork code.** It is generated, and handed to the
  fork through those two options. Q1's plain instances through a table, Q2's derivation and Q3 all
  stand.
- **What A and A2 already type.** After them, a branch's `w:sdt`, `w:tr` and `w:tc` are typed by
  their container (`runtime.BRANCH_CHILD_NAMES`). `w:tbl`, `w:bookmarkStart`, `w:drawing`, `w:pict`,
  `w:t` and the rest of the 61 are still wildcards in a branch (9.11, 10.8).

1. **The hook answers for the roots.** `runtime.branch_child_class`, the parser's hook, takes the
   61 in three steps:
   - **(a) In a branch, the container's binding.** For a name among the 61, as for `w:sdt`,
     `w:tr` and `w:tc` already, the class is the one the container that holds the
     `mc:AlternateContent` binds the name to. This is the rule A and A2 established: a branch's
     child is read as the container's own child is.
     - docx4j's root class is the class its containers bind, since docx4j has one class per
       type.
     - Where this model generated several classes for one type (`w:hyperlink` has seven), the
       container's class is the one the same element gets outside a branch.
   - **(b) Elsewhere, the root class.** Where (a) gives none, in any lax wildcard and in a branch
     whose container does not take the name, the name binds to its root class. That is what JAXB
     does with an `@XmlRootElement` name: section 6's first risk, every lax wildcard sees the
     table.
   - **(c) Otherwise None:** the lookup by name, as now.
   - **The parser side:** `wildcard_class_for` becomes `BRANCH_CHILD_NAMES` together with the
     roots' names.
   - **The saving side:** `_branch_element_names` gains each root class under its name. That table
     must be a function, class to name, and its build raises if it is not. The classes that (a)
     can give besides the root class are named too, each under the one name that gives it.
   - **No other name is typed.** A branch's `w:ins`, `w:instrText` or `w:delText` stays a wildcard
     tree, as docx4j keeps it a DOM element.
2. **The table, derived mechanically** (Q2, as first written).
   - `codegen/derive_names.py` writes `codegen/names/root_elements.json` from docx4j's XJC output,
     the source of the other name tables. It takes every `@XmlRootElement` class and unscoped
     `@XmlElementDecl` whose name this model doesn't declare globally, and gives each the class the
     `el` table names for it.
   - `w:sdtContent` is left out (Q3).
   - It is cross-checked against the schema copy's annox annotations (61 of the 62; `w:sdtPr` is
     the 62nd).
   - **Provenance.** The XJC output is a build in docx4j's checkout, not a commit. The cross-check
     against the schema copy, which is at docx4j `2853b7280`, is what ties the table to a commit.
     The implementation notes record both.
   - The file is committed for review and never hand-edited, like the other name tables.
3. **Generated, then registered.**
   - `codegen/generate_el.py` writes the table into `docx4j_py/el_index.py` as `ROOT_ELEMENTS`:
     qualified name to the class's path, as a string.
   - `docx4j_py/runtime.py`, hand-written, resolves it when the hook and the serialiser's table
     are first built. It never resolves it at import, since the engine imports no model class at
     module level.
   - `codegen/generate.sh --check` stays byte-identical.
4. **The readers the change exposes.**
   - `traversal.text_of` and `walk` read a typed child of a kept branch as they read it elsewhere,
     under every `mce` mode. This is the three `test_mce.py` failures of section 3.
   - The two `walk(mce="all")` tests expect `Wsp`.
   - **An audit**, as Phase A's (9.6). Every reader of text or structure is run over the corpus
     documents that hold a branch, before and after, and any answer that changes is listed with
     why. Two readers whose wildcard paths may now be dead:
     - the text model's typed copy of a branch (`Extractor._typed_block`);
     - `revisions/islands.py`, which reads a Choice's text box.
     Each stays only where it is still reached. Its docstring says where.
   - `Body.tables` reads a table in a block-level branch, and `find(P)` reaches a cell-level
     branch's paragraphs (9.11).
5. **The fork.** No code.
   - The hook already runs for every lax wildcard. Item 26's table already writes any class it is
     given.
   - If the fork's tests do not pin a behaviour this phase relies on, a tests-only item 28 joins
     stage 11, as item 27 did. Two candidates: the hook answering in a wildcard that is not a
     branch's, and a typed child holding a branch of its own.
6. **The record.**
   - CHANGELOG "Changed": a kept branch's WordprocessingML children are typed. Code that looks for
     an `AnyElement` `w:drawing` in `any_element` now finds a `Drawing`.
   - A note in CR-002 section 5.6.
   - CLAUDE.md's MCE paragraph: `mce_preprocess` is no longer the only way to get a typed
     `Drawing` in a choice. This is the coordinator's.

### 4.3 Phase C: VML in the model (the Fallback side)

Copy docx4j's `xsd/vml/` (`vml-main.xsd`, `vml-officedrawing.xsd`, `vml-wordprocessingdrawing.xsd`,
`vml-spreadsheetDrawing.xsd`, `vml-presentationDrawing.xsd`) at `acd0defc9` into the schema copy,
and reach them from `docx4j_python__ROOT.xsd`. That commit is docx4j CR-048, which adds
`w14:anchorId` to all eleven VML shape types and `v:group`, and `o:gfxdata` to the nine that
lacked it. Then re-derive the names and regenerate.

VML's elements are global in its schema, and `CT_PictureBase` already admits VML and `o:` with
`processContents="lax"`. So once VML is in the model, the parser types it wherever it appears,
with no root table:
- under every `w:pict` (542 in the corpus, about 55 of them outside a branch) and `w:object` (37);
- 1,923 `v:` elements in 37 documents;
- `v:textbox`'s `w:txbxContent`, which `CT_Textbox` declares.

The islands can then retire.

This changes legacy VML everywhere, not only in branches. CR-048 found that docx4j's own VML
binding was losing attributes before it. So the phase has its own survey and gate: zero newly
skipped, canonical identity over the corpus, and a Word check. It is proposed here so the whole
answer is in one place. It starts only after Phase B, on its own go (Q4).

**Revised 2026-10-22, before implementation** (the user: "do phase c"), for what Phases A to B
built and what the copy now is.

1. **The commit.**
   - The VML schemas are copied at `2853b7280`, the commit the rest of the schema copy is at
     (`schemas/PATCHES.md`). docx4j's `xsd/vml/` has not changed since `acd0defc9`, which is an
     ancestor, so the files are CR-048's.
   - Each file is checked against its blob at both commits.
   - `vml-main.xsd` imports `../wml/wml.xsd`, `../wml/w14_word_2010_wordml.xsd` and
     `../shared/shared-relationshipReference.xsd`, all three already in the copy, and the four other
     VML schemas.
2. **`vml__ROOT.xsd` is left out.** It declares a vmlDrawing part's root, `<xml>` in no namespace
   (docx4j CR-026). That part belongs to SpreadsheetML and PresentationML (CR-001 D, CR-002 C), and
   WordprocessingML never holds one.
3. **Names and modules.**
   - `derive_names.py` writes a name table for each VML namespace that docx4j's packages declare:
     `org.docx4j.vml`, `.officedrawing`, `.wordprocessingDrawing`, `.spreadsheetDrawing` and
     `.presentationDrawing`. It reads the XJC build, cross-checked against the schemas as Phase B's
     table was.
   - Each namespace gets a `Substitution` to a `docx4j_py` module, following the existing
     convention, and `el` tables.
   - The prefix table already has `v`, `o`, `w10`, `xvml` and `pvml`.
4. **Wider than the branches.** `CT_PictureBase`'s two lax wildcards (VML and `o:`) are under every
   `w:pict` and `w:object`, in a branch or not. `CT_ShapeDefaults`' `o:` wildcard is in
   `settings.xml` (`w:shapeDefaults`) and `w:hdrShapeDefaults`, so nearly every Word document's
   settings part changes from wildcard to typed. Headers' VML watermarks change too. The survey
   covers every WordprocessingML part for that reason.
5. **The gate, as section 4.3 says, made precise.**
   - Over the corpus Phase B's survey built from commits (`scripts/mce_branches.py corpus`): zero
     newly skipped, no new canonical difference, and every byte difference listed by kind.
   - **VML's booleans are a named risk.** VML spells them `t`/`f`/`true`/`false`
     (`ST_TrueFalse`, a string), and this model writes an `xsd:boolean` numerically. Any VML or `o:`
     attribute whose spelling changes is listed. Unless Word is shown to accept the new spelling,
     it is kept as the source spells it: the type stays a string, or a marked schema patch.
6. **The islands retire where they are no longer reached.**
   - `model/revisions/islands.py` reads a VML text box through a typed copy, as do the text model
     (`_typed_block`, `_typed_run_level`) and any other wildcard-to-typed path Phases A and B
     listed.
   - Each is removed only if nothing reaches it after the change, and its readers then read the
     typed tree in place.
   - The readers' audit is Phase B's (`scripts/mce_branches.py readers`, before and after). Every
     changed answer is listed with why, and any whose contract is unclear goes to the coordinator.
7. **Fork fixes, if needed.** The fork's generator or runtime may meet VML shapes no earlier schema
   had: mixed content, attribute groups, VML's `ST_` types. A defect found is fixed in the fork as a
   stage 12 item, with tests under `tests/fork/` and a `docs/UPSTREAM.md` row, as the earlier stages
   were. It joins stages 9 to 11 in needing a fork release.
8. **Cost.** Five namespaces of classes join the model. Import time, `warm_up()` and
   `bench_marshal.py` are measured before and after on `samples/`, `Symbols.docx` and a VML-heavy
   document. The CR-007 performance test's thresholds (`tests/test_performance.py`, slow) are
   checked, and a breach is reported, not loosened.
9. **The Word check.** Its README names the Review view (All Markup) for any file with tracked
   changes, and stages each unmodified source beside its save. The files:
   - a legacy VML text box with tracked changes: `cr043-tracked.docx`'s C22, whose Fallback is VML;
   - a VML watermark in a header;
   - a `w:object` (an OLE object's `v:shape`);
   - a VML-only document with no `mc:AlternateContent`;
   - `settings.xml`'s typed `o:shapedefaults`, through any of the above.

### 4.4 Phase A2: row and cell level, as docx4j CR-051 phase 2

The user's go, 2026-10-21 ("make the CLAUDE.md edits, then do phase A2"). docx4j's phase 2
(`9db1a3cbd`, pushed; its CR-051 section 12) teaches the walks that read rows and cells by type
the element, and lifts phase 1's carve-out. Phase A here kept row- and cell-level elements whole
from the start, but its readers pass them by (section 9.6). So in the default mode a row's or
cell's text in a kept branch is missing from the text model, and its tracked changes are not
listed. This phase closes both.

1. **Typed rows and cells in a branch (the fork, stage 11 extended).** Today a branch's `w:tr` or
   `w:tc` is a wildcard tree (9.6). Item 25's hook (`runtime.branch_child_class`) answers for
   `w:sdt` by what the container binds. It is extended to `w:tr` and `w:tc` by the same rule, read
   from the metadata: a row-level branch's `w:tr` is the class `CT_Tbl` binds, and a cell-level
   branch's `w:tc` the class `CT_Row` binds. Item 26's name table names them on write. docx4j
   binds them so by name. Nothing else in a branch changes: `w:p`, `w:tbl` and `w:r` stay as they
   are until Phase B. The probes stay canonically identical in both modes.
2. **The text model** (docx4j `Extractor.rows` and `cells`). An element among a table's rows or a
   row's cells is compared with its other branches (`mce.branches_differ` at the element's
   address), and Word's branch (the TEXT purpose) is walked in its place, as `blocks` does at
   block level. A row inside it is addressed `body/<tbl>/<i>/0`. In `word` mode the walks read the
   load's record for their container, so the warning is the same at the same address in both
   modes.
3. **The revisions engine** (docx4j `Lister.tableTokens`, `alternateRows`, `alternateCells`,
   `pairTwins`; `Resolver.row`; `Texts.rowText` and `cellText`).
   - Word's branch's row and cell changes are listed where the element stands, and every other
     branch's become twins by kind and position, as `_alternate_level` does at block level.
   - Accepting or rejecting one resolves its twin.
   - A table left with no rows is removed only when no branch has a row left.
   - The row and cell texts read rows and cells through content controls, custom XML and a kept
     element.
4. **The content API.** `Table.rows` and `TableRow.cells` read Word's branch, read-only, as
   `Body.paragraphs` does at block level, and an edit there raises `address.read_only`. `outline()`
   and `describe()` follow. Anything else that walks rows or cells by type is audited, as docx4j's
   section 12.1 surveyed every `instanceof Tr` and `Tc`. That covers `model/content/`, the property
   resolver's parent climb, the python-docx facade and Markdown's tables.
5. **Not ported: OpenDoPE's two cases** (docx4j's `RemovalHandler.xslt` row count,
   `OpenDoPEHandler.emptyKeepingTableShape`, `McRowCellLevelOpenDoPETest`). This package's
   template fill has no OpenDoPE conditions or repeats and removes no control (CR-003 section 17;
   `model/customxml/template.py`). Recorded.
6. **The tests**, docx4j's at `9db1a3cbd`, against the 19 probes:
   - `McBlockLevelLoadTest.preserveModeKeepsBlockLevelElements`, every container;
   - `McBlockLevelProbesTest.preserveModeKeepsEveryBranch` without its row and cell exception, and
     `preserveModeReadsAsWordsSave`: each probe with a Word 365 save reads as Word's save, in
     `preserve` mode and after a `preserve` save;
   - `McBranchSdtTest`'s row and cell controls typed in both branches;
   - `McBlockLevelReadersTest`'s row- and cell-level element (`body/3/1`, `body/3/2/1`), the
     warnings equal across the modes;
   - `TrackedChangesAlternateContentTest`'s row and cell cases and its two row-removal cases.
7. **The gate.**
   - The whole suite, `roundtrip.py` 50 of 50, and `generate.sh --check`.
   - CR-011's ratchet and agreement test unchanged, and CR-010's Word-save comparisons unchanged.
   - `out/acceptance/` canonically identical.
   - A Word check of the five row and cell probes saved in `preserve` mode and of one in `word`
     mode, since branch rows and cells are now written from typed objects. docx4j's own check of
     its `preserve` saves passed (its section 12.3).

## 5. Tests and the gate

- **Phase A tests,** in `tests/openpackaging/test_mce_levels.py`, over the 25 probes in both modes:
  - no skipped item on any probe, and `strict=True` loads them all;
  - **`word`:** the branch kept is Word 365's (`w16sdtfl` on the newest-first probes). Where
    Jason saved the probe from Word, the part is compared with Word's save through
    `tests/revisions/description.py`;
  - **`preserve`:** canonical identity with the source;
  - `mc-block-nofallback` (REQ-305) kept whole in `word` mode; an element with no branch removed
    in `word` mode and kept in `preserve`;
  - the record holds each element taken out, with its address; the root's `mc:Ignorable` gains
    the selected Choice's prefixes;
  - a `w:sdt` directly in a paragraph's, a table's and a row's branch is typed as its kind and
    round-trips whole, and `wml(...)` takes the section 3.1 fragment;
  - the readers' audit, one test per reader; the revisions twins against docx4j's
    `TrackedChangesAlternateContentTest`;
  - `mce_preprocess` takes `w16sdtfl` on the newest-first probes.
- **Phase A's gate:** CR-010's fifteen Word-save comparisons and the round-trip tests over
  `samples/` unchanged (no corpus document has a block-level element), and a Word check of a
  block-level probe saved in `word` mode.
- **Phase B tests,** in `tests/openpackaging/test_mce_typed.py`:
  - each of the 61 resolves through the context to its class;
  - a `w:drawing` in an `mc:Choice` and a `w:pict` in an `mc:Fallback`, under a `w:r`, parse
    typed and marshal back whole;
  - a typed child saved writes no `xsi:type`.
- **The census on `samples/`,** pinned: 3 `Drawing` in Choices, 3 `Pict` in Fallbacks, no
  wildcard `w:drawing` or `w:pict` left in a branch.
- **The corpus survey** of section 3, committed as `scripts/mce_branches.py`:
  - zero canonical differences against the sources;
  - every skipped report unchanged;
  - its census recorded in the implementation notes.
- **CR-010's gate held:** the fifteen Word-save comparisons at 0 lines, and docx4j-25 asked to
  rerun docx4j's agreement run (it calls `revisions.accept_all`), 0 of 22.
- **Acceptance:** `out/acceptance/` regenerated and compared canonically with the checkout's
  (`scripts/compare_acceptance.py`). A Word check is asked for, since branch content now
  marshals through typed classes.
- **Cost:** `scripts/bench_marshal.py` before and after, with the numbers recorded. The
  section 3 figures are the prototype's upper bound.

## 6. Risks

- **`word` mode, on request, changes a saved document** wherever block-, row- or cell-level
  alternate content is read: the branches Word does not read go. Word's own save does the same,
  and the record keeps them for inspection. The default, `preserve`, reproduces the input (Q7).
- **In the default `preserve` mode, a reader that passes a block-level element by hides content
  Word shows.** The audit (section 4.1 item 5) makes each reader of text or structure read Word's
  branch, and an edit inside a kept branch is refused rather than written to one branch only.
- **The schema refresh brings docx4j's other changes since `58cce65a8`.** They are listed, and
  each is checked, as CR-001 section 14.10 did.
- **The readers in `preserve` mode** meet an element they never met. The audit (section 4.1
  item 5) is the guard, as CR-051's section 10 was docx4j's.

- **Every lax wildcard sees the table,** not only `mc:Choice` and `mc:Fallback`. Over the corpus
  no other wildcard met one of the 61. A file whose extension content held one would now have it
  typed. That is lossless only where its class binds all of it, and the skipped report says when
  it doesn't. That is the same trade docx4j's lax wildcards make.
- **Typing exposes binding gaps the wildcard hid.** The corpus shows none for the 61. Phase C is
  where docx4j's VML gaps were (CR-048), and its gate is the guard.
- **Code that inspects `any_element`.** In this repository, the readers of section 4.2 item 4 and
  the islands. Outside it, any caller that looked for an `AnyElement` there; the CHANGELOG says
  so.
- **Cost:** up to +24% parse and +40% marshal on a document of 403 branches, by the prototype.
  The fork's version runs only in the wildcard path, and `bench_marshal.py` records the real
  figure.

## 7. Questions (decided 2026-10-21, each as recommended)

1. **Q1: Phase B as proposed:** the 61 roots, as plain instances, through a table in the fork's
   context. *Recommended:* it is what docx4j and objects-ts do, measured lossless here.
2. **Q2: where the table comes from:** `derive_names.py` over docx4j's XJC output, into
   `codegen/names/`, then into generated `el_index.py`. The alternative is the fork's generator
   reading the schema's annox annotations and writing each root name into the class's `Meta`.
   *Recommended:* `derive_names.py`. docx4j's Java is already the source for this model's
   names, it gives `w:sdtPr`, which the annotations lack, and it keeps the generator unchanged.
3. **Q3: `w:sdtContent` left out**, as objects-ts left it, since it has four classes.
   *Recommended.*
4. **Q4: Phase C, VML,** started only after Phase B lands, on its own go and with its own
   survey and Word check. *Recommended.*
5. **Q5: no switch back to the wildcard form.** `mce_preprocess=True` stays the other option,
   for a resolved tree. *Recommended:* the readers don't change, and a switch would keep two
   shapes of kept branch alive in every reader.
6. **Q6: Phase A first, and now.** It ends two losses of content that happen today (section
   1.1). Its fork part (section 4.1 item 2) can share one fork release with Phase B's.
   *Recommended.*
7. **Q7: the mode's default:**
   - (a) `word` by default, `preserve` on request, per package: docx4j's D5, which was the
     user's own proposal there;
   - (b) `preserve` by default.

   *Recommended (revised 2026-10-21): (b),* as docx4j's D5 now decides for 17.4.0.
   - "A library should not drop what it did not touch" is D5's reason, and it cites this
     package, which keeps every branch and selects on read (CLAUDE.md's lossless MCE). The
     round-trip rule then holds with no exception by default.
   - D5 keeps `word` in docx4j 17.3.x only so that a point release does not change what
     existing walkers see. Nothing here has such walkers: today the element is lost on save
     (section 1.1), so either default is a change.
   - The cost is section 4.1 item 5. Each reader of text or structure must read Word's branch at
     block level, and an edit inside a kept branch is refused until a later CR. Under (a) the
     readers see what Word shows without a change, and an edit lands in Word's branch.
   - The recommendation until this revision was (a), docx4j's D5 as first decided (the user's
     own proposal there).
8. **Q8: the understood set:** docx4j's published list, both for `word` mode and for
   `mce_preprocess`, shared with CR-011's Q2. *Recommended.* CR-011's Q2 had kept `UNDERSTOOD`
   for `mce_preprocess`. Section 3.1 shows it then misses Word 365's branch, so this revises
   that half of Q2. `UNDERSTOOD` stays as the prefix table's own list.

## 8. Phases and effort

| Phase | What | Effort |
|---|---|---|
| A | docx4j CR-051 phase 1: the schema refresh and regeneration; the fork's container-dependent `w:sdt`; the load-time mode (`preserve` by default, Q7) with its record, REQ-305 and `mc:Ignorable`; `mce_preprocess` on docx4j's list; the readers' audit, each reader of text or structure reading Word's branch at block level; the revisions twins; the 25 probes and Word's saves; a Word check | 3.5 days |
| A2 | docx4j CR-051 phase 2: the fork's hook extended to a branch's `w:tr` and `w:tc`; the text model's row and cell walks; the revisions engine's row and cell twins and the table-removal rule; the content API's rows and cells read-only; docx4j's phase 2 tests; a Word check | 1.5 days |
| B | the fork's root-element table and its tests; `derive_names.py`, `root_elements.json`, `ROOT_ELEMENTS`; the registration; traversal's typed branch children; tests, survey script, the record; then a fork release before a docx4j-python release | 1.5 days |
| C | docx4j's VML schemas at `acd0defc9`, names, regeneration; the VML survey and gate; the islands retired; Word check | 3 days |

## 9. Phase A implementation notes (2026-10-21)

Phase A is the port of docx4j CR-051 phase 1, read at docx4j `2853b7280`. Each block, row and
cell container now keeps `mc:AlternateContent`, and a load mode says what to do with it:
`preserve`, the default, keeps it whole, and `word` resolves it as Word reads it. A `w:sdt`
straight in a branch is typed by its container's kind. `mce_preprocess` selects over docx4j's
list. Every reader of text or structure reads Word's branch at block level, and an edit in a kept
branch is refused. The revisions engine lists a branch's changes with twins. The 19 probes and
Word 365's 14 saves are in `tests/fixtures/mc-fun/`.

**Commits.**

- docx4j-python:
  - `a4dd7bf`: the schema refresh;
  - `9e6593d`: the mode, the fork hooks, `mce_preprocess`, the readers, the revisions engine, the
    probes and the tests;
  - `b25eb2c` and `22eaa4c`: the text model reads a `word`-mode load's record;
  - `d4eed0b`: this section and the CHANGELOG;
  - `0b33234`: the text model passes row and cell level by; `Body.view_for` takes a known
    container (9.10).
- The fork (`~/git/docx4j-xsdata`, branch `docx4j`, local only):
  - `78103a57`: stage 11, item 24;
  - `54ec5220`: items 25 and 26.

### 9.1 What landed

| piece | where | what |
|---|---|---|
| the schema | `schemas/` (seven files), `schemas/PATCHES.md` | refreshed to docx4j `2853b7280` (9.2); fifteen content lists admit `AlternateContent` |
| the fork | `docx4j_xsdata/models/xsd.py`, `parsers/`, `serializers/` | stage 11 (9.3): a compositor's identity never reused; `ParserConfig(wildcard_class=, wildcard_class_for=)`; `SerializerConfig(element_names=)` |
| the hooks | `docx4j_py/runtime.py` | `branch_child_class` (a branch's `w:sdt` bound as its container binds it), `BRANCH_CHILD_NAMES`, the serialiser's `element_names` for the four `w:sdt` classes |
| the mode | `docx4j_py/openpackaging/mc_block_level.py` (new, under `openpackaging/` in `KEEP`), `load.py`, `packages/opc_package.py`, `parts/xml_part.py` | `LoadOptions(mc_block_level=)`, `pkg.mc_block_level`, `_ParseSettings.mc_block_level`, `splice_word`, `ResolvedAlternateContent`, `part.resolved_alternate_content` (9.4) |
| `mce_preprocess` | `openpackaging/mce.py`, `child.py` | selects over docx4j's published list (9.5) |
| the content API | `model/content/text_model.py`, `body.py`, `paragraph.py`, `table.py`, `controls.py`, `reports.py` | block lists answer a block-level element with Word's branch; a read-only body for views in a kept branch; `recording` refuses it (9.6); `Body.view_for(element, container)`, so a loop of inserts no longer scans the body (9.10) |
| the text model | `model/text/extractor.py` | `mce.branches_differ` from a `word`-mode load's record; row and cell walks pass the element by (9.6) |
| the revisions engine | `model/revisions/trees.py`, `lister.py`, `tables.py`, `change.py`, `__init__.py` | the trees answer the element with Word's branch; the lister's twins; `TrackedChange.twin`; row and cell finders pass the element by (9.6) |
| tests | `tests/openpackaging/test_mce_levels.py` (256), `tests/revisions/test_alternate_content.py` (9), `tests/openpackaging/test_adoption.py` (+5); `tests/agent/test_change_report.py`'s timing test (9.10) | section 5's Phase A tests (9.8) |
| fixtures | `tests/fixtures/mc-fun/` (33 files) | docx4j's probes and Word 365's saves, byte for byte |
| documents | `CHANGELOG.md`, `tests/README.md`, `docs/UPSTREAM.md` (reports 20 to 22), `tests/openpackaging/test_round_trip.py`'s docstring, the fork's `docs/fork/CHANGES.md` stage 11 | |

### 9.2 The schema refresh, and docx4j's changes since the copy

The procedure was CR-001 section 14.10's:

1. **The closure.** It was computed from the five entry points at both commits, following every
   `xsd:import` and `xsd:include`. It is the same 103 files at `c363a969f` and at `2853b7280`.
   `wml/wml__ROOT.xsd`, which nothing references, is unchanged.
2. **The re-patch.** Each patched file was merged three ways: docx4j at `c363a969f` was the base,
   the old copy ours, and docx4j at `2853b7280` theirs.
   - Only patch 4's region conflicted, and it took docx4j's version.
   - The lines by which each patched file now differs from docx4j at `2853b7280` are exactly the
     lines by which the old copy differed from docx4j at `c363a969f`, less patch 4's.
   - Every other file of the closure is docx4j's at `2853b7280`, byte for byte.

`git log c363a969f..2853b7280 -- xsd` lists eleven commits. Each was checked as follows.

| commit | docx4j's change | in the closure? | checked |
|---|---|---|---|
| `1cff2cd5c` | CR-051 phase 1: `mc:AlternateContent` in `EG_ContentBlockContent`, `EG_ContentRowContent`, `EG_ContentCellContent` (`wml/wml.xsd`, 21 lines with the header history) | yes | regenerated: `AlternateContent` joins the `content` union and choices of 15 classes (`Body1`, `Hdr`, `Ftr`, `Tc`, `CTFtnEdn`, `CommentsComment`, `CTTxbxContent`, `SdtContentBlock`, `CTCustomXmlBlock`, `Tbl`, `CTSdtContentRow`, `CTCustomXmlRow`, `Tr`, `CTSdtContentCell`, `CTCustomXmlCell`), 90 lines of `docx4j_py/wml/__init__.py` and nothing else; the same fifteen lists docx4j-generated-objects-ts reported |
| `dae2dfc8b` | `CT_Settings`: `w14:docId`, `w15:chartTrackingRefBased`, `w15:docId` in Word's order | yes | **retires patch 4**, as `PATCHES.md` foresaw: the order is the patch's, so the generated `Settings` and every saved `settings.xml` are unchanged |
| `8e8f6ea83` | `a:graphicData`'s wildcard `lax` (`dml/dml-graphicalObject.xsd`) and the vmlDrawing root's (`vml/vml__ROOT.xsd`) | the first file only | no generated change: the generator writes `process_contents` only for `skip`, and the parser treats `strict` as `lax` |
| `409aade38` | 17.1.1 renamed 17.2.0 in comments | six files (`wml.xsd`, `dml-chart.xsd`, `dml-spreadsheetDrawing.xsd`, `dml-textParagraph.xsd`, the mce schema, `oart14docprop.xsd`); `ROOT.xsd` and SML are outside | comments only; no generated change |
| `14fcb2f4e`, `28bdf831c` | CR-040: chartex `cx:autoUpdate` and the bins | yes | already in the copy (CR-001 section 14.12); byte-identical at `2853b7280` |
| `acd0defc9` | CR-048: VML `w14:anchorId`, `o:gfxdata` (`vml/vml-main.xsd`) | **no**: `wml.xsd` imports no VML; `CT_Picture`'s lax wildcard names the namespaces without an import | not copied. It is Phase C's (section 4.3), so no VML was typed and nothing needed asking |
| `df1e535b1` | CR-033 digital signatures (`digSig/`) | no | not copied |
| `5ec4e3ff9` | `FopConfig/fop.xsd` | no | not copied |
| `bde2feb3d`, `b1fa34947` | CR-027: SML calcFeatures and revision uids; `shared/shared-bibliography.xsd` | no | not copied (SML is CR-001 Phase D's; the bibliography schema is outside `wml.xsd`'s closure) |
| `b4ca0d98a` | CR-026: `ROOT.xsd`, `vml/vml__ROOT.xsd` | no | not copied |

**The name tables.** docx4j added no type, so the tables were not re-derived. The `el` tables are
unchanged, with 2,063 element names over 60 namespaces. There are 282 `ClassNames` warnings, as
before, and the module count is unchanged.

**The markers.** `grep -ro 'docx4j-python' schemas --include='*.xsd'` finds 26 matches:

- 24 patch markers: 21 wildcards, one for `CT_Dir`/`CT_Bdo`, and two for patch 3;
- `docx4j_python__ROOT.xsd`'s comment;
- docx4j's own mention of this repository in `office_drawing_2017_03_chart_.xsd` (since
  `c363a969f`), which is not a marker.

**`generate.sh --check` first failed** on the refreshed schemas. The cause was not the refresh:
the generator's compositor identity is taken from `id()`, which can be reused (9.3, item 24).
`CT_TextCharacterProperties`' underline choice merged with the fill choice into one field, but
only when the fork's bytecode was cached. So the check's two runs disagreed, and the committed
output came from the uncached form. With the fork's fix, the output is the same cached or not. It
equals the committed output apart from the fifteen admissions, and `--check` is byte-identical.
This is probably the earlier "stale bytecode" regeneration diff that sessions have worked around
by deleting `__pycache__`.

**The gate after the refresh alone:**
- `scripts/roundtrip.py`: 50 of 50, nothing skipped.
- The whole suite passes. One pin moved, 9.10's.
- CR-011's ratchet is unchanged.

### 9.3 The fork's stage 11

`docs/fork/CHANGES.md` stage 11 has the detail, and `docs/UPSTREAM.md` reports 20 to 22. Each
item is off by default.

- **Item 24 (`78103a57`): a compositor's identity is never reused.** `models/xsd.py`'s
  `compositor_id(obj)` replaces `id(self)` in the five compositors' restriction paths. Its value
  comes from a process-wide counter, read once per object and kept on it.
  `AddAttributeSubstitutions.prepare_substituted` draws from the same counter.
  - **Tests:** twelve in `tests/fork/test_compositor_identity.py`. Five fail without the fix.
  - **Upstream tests:** seven assertions that pinned `id()` were edited.
- **Item 25 (`54ec5220`): `ParserConfig(wildcard_class=, wildcard_class_for=)`.**
  - For a name asked for, a lax wildcard asks the hook for the class before the lookup by name.
    The hook gets the enclosing classes, innermost first, through a new `ElementNode.parent` slot
    and `ancestor_classes()`.
  - `runtime.branch_child_class` answers for `w:sdt`. Past the `mc` wrappers, the first class is
    the container, and the hook returns the class that container binds `w:sdt` to (`SdtBlock`,
    `SdtRun`, `CTSdtRow`, `CTSdtCell`). The rule is the schema's own binding, read from the
    metadata, not a table.
- **Item 26 (`54ec5220`): `SerializerConfig(element_names=)`.**
  - An object no field names (a wildcard's typed child, or the root) is written under the name
    the table gives its class, not its type's name (`CT_SdtRun`).
  - A field's own name always wins, and no `xsi:type` is written.
  - `runtime.serializer_config` names the four `w:sdt` classes. Phase B's root table extends it.
- **The fork's suite:** 1,620 passed and 1 skipped at stage 10. It is now 1,651 passed and 1
  skipped. Of the 19 new tests for items 25 and 26, 17 fail without them.

Section 3.1's fragment, a run-level control straight in a paragraph's `mc:Choice`, now goes
through `wml(...)` with nothing skipped. It parses as `SdtRun` and is written back canonically
identical.

### 9.4 The mode

- **The API.**
  - `LoadOptions(mc_block_level="preserve" | "word")` defaults to `"preserve"`. Any other value
    raises `ValueError`.
  - `pkg.mc_block_level` is this package's own setting, over the load options'. Setting it does not
    change a shared `LoadOptions`. docx4j has it as `OpcPackage.setMcBlockLevel` over
    `McBlockLevel.of`.
  - The mode is read when each part is first parsed. A part already parsed keeps its tree.
- **The parse settings.** The mode is the third field of `_ParseSettings` (CR-002 section 16.2).
  So a private tree parsed under one mode is never adopted under the other, and is parsed again.
  CR-002 section 16's new tests cover both directions, the same mode, and the load options changed
  in place.
- **Scope.** An element is in scope when the element it stands in, past any `mc` wrappers, is one
  of these containers:
  - a block container: `w:body`, `w:hdr`, `w:ftr`, `w:tc`, `w:footnote`, `w:endnote`, `w:comment`,
    `w:txbxContent`, `w:docPartBody`;
  - `w:tbl` (rows) or `w:tr` (cells);
  - a `w:sdtContent` or `w:customXml`, at the level of the content its control or element stands
    in.

  A run- or paragraph-level element, and one in a property, is never in scope. That is docx4j's
  `McSplice.inScope` over the effective container.
- **`word` mode.** It works on the lxml tree of the part's bytes, before the fork's parser, as
  `mce.py` does.
  - The elements are taken in document order, which is outermost first.
  - An element no longer reachable from the root (it was in a discarded branch) is skipped. So is
    one inside an element kept whole.
  - An element with no branch is removed and recorded.
  - One with no understood Choice and no Fallback is kept whole (REQ-305).
  - Otherwise the first understood Choice, else the Fallback, replaces the element in place. A
    Choice is understood when every prefix its `Requires` names is bound, in scope at the Choice,
    to a namespace on docx4j's list, matched by URI.
  - The selected Choice's `Requires` prefixes are added to the root's `mc:Ignorable`, as the
    conventional prefix of each namespace. The writer declares each one from the source's
    declarations or the prefix table.
  - Only a part from which something was taken is serialised again. Otherwise the parser reads
    the original bytes.
  - A part whose bytes do not hold `AlternateContent` pays one substring test.
  - `mce_preprocess`, when it is also on, runs after the splice.
- **The record.** `part.resolved_alternate_content` holds one `ResolvedAlternateContent` per
  element taken out, outermost first. Each has:
  - the element whole, as a typed `AlternateContent`;
  - the branch whose content was promoted (None where the element was removed);
  - the container's prefixed name and its level;
  - the container's address, such as `w:document/w:body[0]/w:tbl[1]/w:tr[0]/w:tc[1]`: each step
    the element's name with its index among its parent's children, property elements left out;
  - the index at which the element stood;
  - how many elements were promoted.

  The element is a parse of its own, with a `wildcard_class` that types a `w:sdt` by the record's
  level. `set_bytes` clears the record. `set_contents` leaves it, since it describes the last
  parse of the part's bytes.
- **The round-trip rule's one exception** is stated in `tests/openpackaging/test_round_trip.py`'s
  docstring. In `word` mode, a part holding such an element is not canonically identical when it
  is saved unchanged. `preserve` keeps the rule. CLAUDE.md is not edited (the brief), so its
  "Round-trip and fidelity rules" should gain the same sentence (9.11).

### 9.5 `mce_preprocess` over docx4j's list (Q8)

`mce.py` selects over `child.understood_namespaces()`, docx4j's published list of 118 namespaces
(the same file at `2853b7280` as at `198d015a3`). `UNDERSTOOD` stays the prefix table's own list.

**What moves.** Every part of `samples/`, `tests/fixtures/` and docx4j-core-ts's fixtures (416
documents, 63 parts with the element) was resolved both ways. 14 parts select differently:

- the eleven newest-first probes, which now take `w16sdtfl` where they took `w16cid`, as Word 365
  does;
- `waterfall.docx`, in three copies (`tests/fixtures/`, `tests/fixtures/charts/`, docx4j-core-ts's
  fixtures). Its chart's Choice requires `cx1`, which is on docx4j's list, so the Choice
  (`cx:chart`) is now taken over the Fallback picture.

No test of the suite pinned either answer, so no test moved.

### 9.6 The readers' audit, in `preserve` mode

| reader | block level | row and cell level | test |
|---|---|---|---|
| `traversal` (`walk`, `iter_nodes`, `find`, `text_of`) | Word's branch: `mce="resolve"` already selected at any level | Word's branch too (`mce="resolve"` is level-blind); a cell-level branch's `w:tc` is a wildcard until Phase B, so its text is read but `find(P)` does not reach inside | `test_traversal_reads_words_branch` |
| content API: `Body.paragraphs`, `iter_paragraphs`, `element_at`, `paragraph_at`, `address_of` | Word's branch: the block list of a block-level element is its selected branch's `any_element`, and of that branch itself, so a path names the branch index (`body/2/0`, the text model's preserve-mode address) | passed by: no view, no error (`block_list_of` answers None for an element in a row or cell container) | `test_the_body_reads_words_branch_at_block_level`, `test_the_body_passes_row_and_cell_level_by` |
| content API: `Body.tables`, `Table.rows`, `cells` | a `w:tbl` in a branch is a wildcard until Phase B: passed by | passed by | the same |
| an edit inside a kept branch | refused: each view over content in any kept branch (block, paragraph or run level) gets a read-only body (`Body.body_for`, `_BranchBody`), and `recording`, which every mutating verb opens, raises `AddressError` `address.read_only`; reading through the view works; in `word` mode the edit is made | the same, for the content the API does not hand out | `test_an_edit_inside_a_kept_branch_is_refused`, `test_an_edit_there_in_word_mode_is_made` |
| `outline()` | the element is an entry (`kind` `AlternateContent`, no text) whose children are Word's branch's blocks with the model's text; `stats` the model's | passed by | `test_the_outline_reads_words_branch` |
| `describe()` | passes by: it counts no block; nothing is skipped now | the same | `test_describe_passes_the_element_by` |
| the text model (`extract`, `plain_text`) | Word's branch, through `Extractor.blocks`' existing `mc` case: CR-011 section 14.7 item 8's gap was the schema, not the reader. `mce.branches_differ` at the element's address, and in `word` mode from the record at the same address (`b25eb2c`) | passed by: the row and cell walks skip the element (`0b33234`); no warning. Through the trees change of `9e6593d` the row walk had read a row-level branch (its `w:tr` is typed), and it was made to pass it by so that every reader of structure agrees, as the brief asks | `test_the_text_model_and_markdown_read_words_branch`, `test_the_text_models_warnings_are_the_same_in_both_modes_*` (26), `test_the_text_model_passes_row_and_cell_level_by` |
| Markdown (`to_markdown`, `markdown`) | follows the text model: closes CR-011 section 14.7 item 8 | as the text model | the same, and `test_text_markdown.py`'s moved pin (9.10) |
| the revisions engine | the trees answer the element with Word's branch (`branch_items`; docx4j `Trees.blockItems` and `runItems`), and the lister lists its changes once with the others' as twins (`_alternate_level`), at paragraph level too | passed by (`tables._rows` and `_cells`), as docx4j's phase 2 leaves its `Lister`'s rows and cells | `tests/revisions/test_alternate_content.py`, `test_the_revisions_engine_passes_row_and_cell_level_by` |
| the python-docx facade | passes by: `Document.paragraphs` is `w:body`'s own `w:p`, as python-docx's; editing them works | passes by | `test_the_python_docx_facade_passes_the_element_by` |
| `tests/revisions/description.py` (the Word-save comparisons' describer) | reads Word's branch through the trees | as the trees | the probe comparisons (9.7) |

No reader raises or miscounts on any probe. `Paragraph.text` and the other Office JS readers stay
Office JS's text, which does not descend into `mc:AlternateContent` (CR-011 Q4). That is
unchanged.

### 9.7 The probes, in both modes

The 19 probes and Word 365's 14 saves were copied from docx4j at `2853b7280` with
`git show 2853b7280:<path>`. Each was checked against its git blob with `git hash-object`.

- 16 probes came in at `1cff2cd5c`, and the three original-order probes (`mc-text`, `mc-run`,
  `mc-block`) at `a9ec5680a`. So this is 19 probes, where the brief said 16.
- The 25 of section 3.1 were Jason's share: these 19 and six pandemonium and plain variants that
  docx4j does not hold.

| | `preserve` | `word` |
|---|---|---|
| nothing skipped, `strict=True` loads | 19 of 19 | 19 of 19 |
| main part canonically identical after a save | 19 of 19, every element, Choice and Fallback written | the 13 with a block-, row- or cell-level element changed by design; the six run- and paragraph-level ones identical |
| Word 365's branch (paper section 7.8's six) | `w14` (original order) and `w16sdtfl` (newest first), nothing else | the same |
| described as Word 365's save (`description.py`) | --- | 0 lines on 11 of 14; `mc-txbx-block` and `mc-txbx-sdt-block` 2 lines, `mc-txbx-tc-block` 4 |
| read as Word 365's save (the text model, before and after the save) | --- | 14 of 14 |

The text-box probes' lines are the Word save's regenerated VML Fallback, which shows its text box
twice (CR-051 section 3.5). `mc-txbx-tc-block` also shows Word's recomputed table grid
(`4481 4536` for `4500 4500`). The test pins each line by kind. Of the run-level probes,
`mc-run` and `mc-sdt-run-newest-first` described 2 lines apart until the revisions trees read a
paragraph-level element's branch. Word's save resolves them, and this package keeps them whole in
both modes, as docx4j does. They now describe alike.

### 9.8 The numbers

- **The suite:**
  - the whole suite: 4,782 passed, 8 skipped (CR-002 section 16.5: 4,512 and 8);
  - the fast suite: 3,502 passed, 8 skipped, 1,280 deselected (3,232, 8 and 1,280 before);
  - the 270 new tests are `test_mce_levels.py` (256), `test_alternate_content.py` (9) and
    `test_adoption.py` (5);
  - no other test moved but 9.10's.
- **Unchanged:**
  - `scripts/roundtrip.py`: 50 of 50 canonically identical, 0 differences, 0 skipped, 34 benign
    respellings;
  - CR-011's ratchet: 1,291 pass, 0 fail, 293 pending of 1,584, goldens 12 of 12, with
    `conformance.tsv` unchanged after the refresh, after the readers and after the record reader;
  - the agreement test: 638 passed;
  - the parity goldens: 277 passed, `PINS` empty;
  - CR-010's Word-save comparisons (`test_word_saves.py` and `test_compare_word_resolution.py`):
    68 passed;
  - `scripts/threads.py`: no configuration lost or altered a document;
  - `generate.sh --check`: byte-identical.
- **`scripts/bench_marshal.py`** (`--runs 7`, three runs each, this machine):

  | | before | after |
  |---|---:|---:|
  | `Symbols.docx` parse | 190 to 202 ms | 193 to 194 ms |
  | `Symbols.docx` marshal | 247 to 250 ms | 247 to 252 ms |
  | the 14 main parts, parse | 231 to 245 ms | 234 to 237 ms |
  | the 14 main parts, marshal | 296 to 299 ms | 296 to 301 ms |

  This is within the band, as the hooks cost one attribute store per element node and a set
  lookup only for a typed child of a wildcard.
- **The substring test,** on CR-003 section 16.4's 200-page document (`tests/agent/conftest.py`'s
  `big_document()`, a 292 KB main part without the element), the median of 201: 72 µs, against a
  41 to 45 ms parse. In `word` mode a part with only run-level elements pays an lxml parse it
  discards: `DrawingML_GraphicData_wps.docx`'s 15 KB main part took 2.1 ms in `preserve` and 2.3 ms
  in `word`.
- **`import docx4j_py`:** 613 to 628 ms, against 620 to 630 before (seven runs each, `-X
  importtime`). That is 242 modules both times. The new module is imported only by a parse in
  `word` mode and by `LoadOptions`' check.
- **The acceptance artefacts:** `out/acceptance/` was regenerated and compared with a copy made
  before the first change (`out/acceptance-before-cr012a/`).
  - `scripts/compare_acceptance.py`: 13 artefacts, 175 entries, 0 differing beyond the wall-clock
    dates;
  - entry by entry, byte for byte, with the dates scrubbed: 175 of 175 identical.

  Saved output does not change for documents without the element.

### 9.9 Departures, with why

1. **`preserve` is the default**, against docx4j `2853b7280`'s code, which defaults to `word`.
   This is CR-012 Q7, decided on its revised answer, as docx4j's D5 decides for 17.4.0. docx4j
   17.3.x keeps `word` so that a point release does not change what existing walkers see.
2. **`preserve` keeps row- and cell-level elements whole.** docx4j's phase 1 still resolves them in
   `preserve` until its phase 2. Here they are kept, so the round-trip rule holds with no exception
   by default. The content API, the text model and the revisions engine pass them by without error,
   as docx4j's phase 1 leaves its readers. `traversal`'s `mce="resolve"` is level-blind and reads
   them, as it did before this phase. **Overtaken by Phase A2 (section 10):** docx4j's phase 2
   (`9db1a3cbd`) keeps them whole in `preserve` too, so this is no longer a departure, and the
   readers here now read Word's branch at row and cell level, an edit there refused.
3. **The splice is on the source's lxml tree, before the parse**, as section 4.1 item 3 asks.
   docx4j splices its unmarshalled objects after the unmarshal. Here the selected content is parsed
   in place, typed by the ordinary path, so the record's element is a parse of its own and not the
   objects the tree held.
4. **The record's address** indexes every step but the root among the parent's elements, with
   property elements left out. docx4j indexes only a step whose parent has a content list, and
   names an unindexed step by its first child of that name. Both are worked out at the splice.
5. **`mc-block-nofallback` is not REQ-305's case.** Section 5 says it is "kept whole in `word`
   mode", but its first Choice (`w16sdtfl`) is understood, so `word` mode promotes it, as docx4j's
   `noFallbackProbe` and Word 365's save do. REQ-305 (no understood Choice and no Fallback) is
   tested on a built element instead.
6. **Read-only reaches every kept branch,** at run and paragraph level too. An edit in a text box
   in a run's `mc:Choice`, or in a run of a paragraph's branch, would change one branch alone, so
   it is refused as a block-level branch's is (CR-011 Q9's rule for text boxes, by another route).
7. **`element_at` resolves an address into a block-level branch and returns a read-only view.** It
   does not refuse the address itself, so an agent can read by address, and the edit is what
   raises `address.read_only`.
8. **The python-docx facade passes the element by**, as python-docx reads `w:body`'s own
   paragraphs.
9. **When both are on, `mce_preprocess` runs after the `word` splice,** so the block-level element
   is recorded and the rest resolved.
10. **The fork's item 24 was not planned.** It fixes a generator defect the refresh exposed, and it
    shares stage 11 with the planned items 25 and 26.
11. **A `w:customXml` straight in a branch** stays a wildcard: section 4.1 item 2's "should either
    become typed". docx4j binds it by the container's kind at its splice. Here the `word` splice
    promotes it before the parse, so it is typed as its container takes it. In `preserve` it stays
    a wildcard until a root table names it.

### 9.10 Pre-existing tests whose expectation changed, and why

- `tests/text/test_text_markdown.py::test_markdown_follows_the_text_model_through_alternate_content`
  (CR-011 section 14.7 item 8's pin, in `a4dd7bf`):
  - **Then:** the kept load gave `before\n\nafter\n`, and the skipped report named the element.
  - **Now:** it gives `before\n\nCHOICEW14\n\nafter\n` with nothing skipped. That is docx4j's
    `MarkdownAlternateContentTest`: the Markdown follows the text model to Word's branch, in both
    modes.
- **In the fork:** seven upstream assertions that expected `id()` in a restriction path or a
  substituted choice now expect the compositor's integer (9.3).

- `tests/agent/test_change_report.py::test_the_report_costs_a_few_microseconds_and_does_not_grow`
  (CR-003 section 3.4), a timing test whose expectation is unchanged: the last 500 of 2,000 inserts
  within four times the first 500. Its method changed. With this phase's test module collected,
  the whole fast suite failed it every time (21 µs at the start, 190 to 250 µs at the end), and it
  passed alone and in `tests/agent`. A profile in the failing configuration found two causes:
  - `Body.view_for`, which `insert_paragraph` returns through, found the new paragraph's list by
    scanning the body's content for it: 2,013,006 identity tests over 2,001 inserts, quadratic,
    and older than this phase. `view_for` now takes the list where the caller knows it
    (`_insert_element`, `__getitem__`, `iter_blocks`), so the scan is gone and the loop costs
    0.22 s where it cost 0.52 s;
  - the rest was a garbage collection over the 350,000 objects the collected suite keeps alive,
    landing in the last measured block: with the collector paused around the test it passed, and
    with a plugin of the same shape that paused nothing it failed.

  The test now collects, then pauses the collector while it times, so it measures the call. It
  passes in the whole fast suite, alone and in `tests/agent` (`0b33234`).

Nothing else moved: the whole suite's difference from CR-002 section 16.5 is this phase's new
tests.

### 9.11 Found, and not done

- **Phase B's typing.** A `w:tbl`, `w:tc`, `w:drawing` and the rest of the 61 roots are still
  wildcards in a branch. So:
  - `Body.tables` passes a table in a block-level branch by;
  - the text model passes a cell-level branch by;
  - `find(P)` does not reach a cell-level branch's paragraphs.
- **Row and cell level in `preserve`, docx4j's CR-051 phase 2.** It has landed in docx4j as
  `9db1a3cbd` (unpushed; the coordinator's word, 2026-10-21): in `preserve` mode docx4j's readers
  now read Word's branch at row and cell level too. Porting it is a follow-up for this package.
  This phase keeps to `2853b7280`, its probes and tests included, and takes nothing from that
  commit. Here, until then, the content API's, the text model's and the revisions engine's row and
  cell walks pass the element by without error, and `mce.branches_differ` is not raised at row or
  cell level. **Ported as Phase A2 (section 10).**
- **The inspector.** CR-050's inspector, the record's second reader in docx4j, has no port here, so
  only the text model reads the record.
- **A record inside a wildcard text box.** An element inside a text box that is itself in a
  run-level branch has an address through the wildcard tree. The text model reads such a box
  through a typed copy, so in `word` mode its warning is not raised. No probe has the shape.
- **CLAUDE.md** (not edited, by the brief) should say two things:
  - "Round-trip and fidelity rules": an unmarshalled part saved unchanged is canonically identical
    unless it was loaded in `word` mode with a block-, row- or cell-level element;
  - its MCE paragraph: `mce_preprocess` takes Word 365's branch over docx4j's list.
- **A fork release.** Stages 9 to 11 are local, so this package is not released against them
  (`RELEASING.md`). That is the user's, as CR-011's Q3 already needed.
- **CR-011 section 14.7 item 8 and 14.10's matching bullet** now say they are closed here.
- **The Word check** passed on 2026-10-21 (9.12).

### 9.12 The Word check (passed 2026-10-21)

Three files are in `out/cr012-a-word-check/`, with a `README.txt` saying what each should show.
They are also copied to `~/vbShares/Office 2016/__tmp/python-cr012a/`. Each was loaded, its main
part unmarshalled so that this engine writes it, and saved.

- `cr012a-1-block-newest-first-word-mode.docx`: `mc-block-newest-first` in `word` mode. The body's
  element is replaced by the `w16sdtfl` paragraph, and the root carries `mc:Ignorable="w16sdtfl"`
  with its declaration.
- `cr012a-2-sdt-row-newest-first-word-mode.docx`: `mc-sdt-row-newest-first` in `word` mode. The
  element among a row-level control's rows is replaced by the `w16sdtfl` row.
- `cr012a-3-mixed-preserve-mode.docx`: `mc-mixed` in `preserve` mode. All four elements are kept
  whole, and each Choice's `w14` is declared.

Expected: each opens without a repair prompt and shows the `w16sdtfl` or `w14` branch, as the
source does in Word 365. **Passed on 2026-10-21** (the user: "__tmp/python-cr012a/ are all good"):
all three open with no repair prompt and show what the README says.

## 10. Phase A2 implementation notes (2026-10-21)

Phase A2 is the port of docx4j CR-051 phase 2, read at docx4j `9db1a3cbd` (its section 12, the
commit's diff and its tests), at the user's go ("make the CLAUDE.md edits, then do phase A2").
Phase A kept a row- or cell-level `mc:AlternateContent` whole in `preserve` mode, and its readers
passed it by (9.6). Now a branch's `w:tr` and `w:tc` are typed. The text model, the revisions
engine and the content API read Word's branch at row and cell level, as they read it at block
level, and an edit there is refused. The probes are Phase A's 19, unchanged at `9db1a3cbd` (each of
the 33 files checked against its blob there). docx4j's phase-2 tests are ported against them, and
its readers' probe is built here as docx4j builds it.

**Commits.**

- docx4j-python:
  - `8b5ea5b`: a row- or cell-level branch's `w:tr` and `w:tc` typed;
  - `5d5f8a5`: the revisions engine, with its tests;
  - `f32927a`: the text model and the content API, with docx4j's phase-2 tests;
  - `6358a37`: a test of the tracked clear over a block-level element;
  - `a92c583`: this section, the CHANGELOG, `tests/README.md`, `docs/UPSTREAM.md` report 23.
- The fork (`~/git/docx4j-xsdata`, branch `docx4j`, local only): `297ba8b7`, stage 11 item 27
  (tests only).

### 10.1 What landed

| piece | where | what |
|---|---|---|
| typed rows and cells | `docx4j_py/runtime.py` | `BRANCH_CHILD_NAMES` adds `w:tr` and `w:tc`, so `branch_child_class` binds a branch's row as the table binds it (`Tr`) and a branch's cell as the row binds it (`Tc`), in a row- or cell-level control's content and custom XML too; `element_names` writes them `w:tr` and `w:tc` (`Tc`'s own name is its type's, `CT_Tc`) |
| the record's parse | `openpackaging/mc_block_level.py` | `_record_hook` binds a `w:sdt`, `w:tr` or `w:tc` straight in a recorded element's branch as a container of the record's level binds it (`Body1`, `Tbl`, `Tr`), so a `word`-mode record of a cell-level element holds `Tc`s |
| the text model | `model/text/extractor.py` | `_rows` and `_cells`: the element compared with its other branches (`mce.branches_differ` at its address) and Word's branch walked in its place; a `word`-mode record read for the container (10.3) |
| the revisions engine | `model/revisions/lister.py`, `tables.py`, `resolver.py`, `texts.py` | `_row_items`, `_row_tokens`, `_cell_items`, `_cell_tokens`; `_alternate_level` takes its level's walk; `tables.rows_of` and `cells_of` read Word's branch; `tables.has_rows`; `Resolver.row` over it; `row_text` through `cells_of` and `rows_of` |
| the content API | `model/content/text_model.py`, `table.py`, `reports.py`, `body.py`, `controls.py` | `block_list_of` answers a row- or cell-level element with Word's branch; `rows_of` and `cells_of` through it; read-only row and cell views; `refuse_read_only` (factored out of `recording`); `kept_content_under` and the tracked deletions' refusal; `outline()`'s cells through any wrapper; `ContentControl.form` through a kept element |
| tests | `tests/openpackaging/test_mce_levels.py` (256 to 319), `tests/revisions/test_alternate_content.py` (9 to 19); the fork's `tests/fork/test_wildcard_class_local_names.py` (5) | 10.4 |
| documents | `CHANGELOG.md`, `tests/README.md`, `docs/UPSTREAM.md` (report 23), the fork's `docs/fork/CHANGES.md` (item 27) | |

### 10.2 The fork's item

The brief asked for stage 11 to be extended as item 27. **The fork needed no code.** Item 25's hook
is generic: the names it is asked for are this package's (`runtime.BRANCH_CHILD_NAMES`), and the
answer is `runtime.branch_child_class`, which reads the container's binding from the metadata.
Adding `w:tr` and `w:tc` there was enough. The probes saved canonically identical and nothing was
newly skipped, so no fork change was made.

`w:tc` is a new shape for the hook all the same. `CT_Tc` has no global element at all, so the lookup
by name leaves a cell-level branch's cell an `AnyElement` tree. That is lossless, but a reader of
typed cells cannot see it. `w:sdt`, by contrast, found a global element of the wrong class. So item
27 is five tests that pin what this package now relies on:

- without the hook such a name stays a wildcard tree and nothing is skipped;
- with it the name is typed as its container binds it;
- a child the hook typed is an enclosing class of what it holds, so a branch nested in it is bound
  by it;
- a global name is asked about before the lookup by name;
- item 26's table writes the typed children back under their element names, and they read back
  equal.

The five tests pass on the stage-11 code; they fail only if that behaviour regresses.
`CHANGES.md` has the entry, and `docs/UPSTREAM.md` report 23 records the shape as report 21's
second use. The fork's suite went from 1,651 passed and 1 skipped to 1,656 passed and 1 skipped.

### 10.3 The walkers, and what each does now

docx4j's section 12.1 surveyed every `instanceof Tr` and `Tc`. Here the survey was a grep across
`docx4j_py/model/`, `docx4j_py/docx/`, `traversal.py` and `child.py`. It looked for the row and cell
classes (`Tr`, `Tc`, `trees.cls("Tr")`), their names (`W_TR`, `W_TC`, `"tr"`, `"tc"`),
`rows_of` and `cells_of`, and the text model's `TextRow` and `TextCell`. Each walker found is
below, with what it does now in `preserve` mode. "Word's branch" is `mce_branch`'s: the first Choice
over docx4j's understood list, else the Fallback.

| walker | what it does now | before (Phase A) |
|---|---|---|
| **text model** `Extractor._rows`, `_cells` | Word's branch walked in the element's place; `mce.branches_differ` at the element's address (`body/1/1`); a row inside it is `body/1/1/0`; in `word` mode the record for the container raises the same warning at the same address | passed by, no warning |
| Markdown (`model/text/markdown.py`, `model/markdown/export.py`, `Table.to_markdown`), plain text (`model/text/plain.py`), the outline's preview | read the text model's `TextRow`s and `TextCell`s, so they follow it | missing the branch's rows and cells |
| **revisions** `Lister._table_tokens` → `_row_items`, `_row_tokens`, `_cell_items`, `_cell_tokens` | a row- or cell-level element goes to `_alternate_level` with its level's walk: Word's branch's changes listed where it stands, the others' as twins by kind and position; `accept_all`'s listing walks every branch | passed by: nothing listed |
| `tables.rows_of`, `cells_of` (and their callers `record_of`, `change_text`, `accept`, `reject`, `collapse_grid`, `remove_cell`, `in_change`, `has_records`) | through `trees.block_items`, which answers the element with Word's branch, as docx4j's `Lister.rowsOf` and `cellsOf`; a row or cell there is held by the branch's list | passed by |
| `tables.has_rows` (new), `Resolver.row` | a table goes only when no branch has a row left | `rows_of` empty |
| `texts.row_text`, `_cell_text` | through `cells_of` and `rows_of`: controls, custom XML and a kept element | direct `w:tc` children; a nested table's direct `w:tr`s |
| `revisions._sweep_content` | through `trees.children`, so Word's branch only, as docx4j's sweep | the same, at block level |
| **content API** `text_model.block_list_of`, `block_children_of` | a row- or cell-level element answers with Word's branch's list, and that branch with its own, so a path through the element names the branch's index | block level only |
| `text_model.rows_of`, `cells_of` (`_wrapped`) | descend into the element as into a row- or cell-level control; a row or cell there is held by the branch's list | passed by |
| `Table.rows`, `row_count`, `values`, `cell()`, `column_widths_twips`, `to_dict`; `TableRow.cells`, `cell_count`, `values`, `row_index`; `TableCell.cell_index`, `address`, `body` | read Word's branch; a row or cell view there has a read-only body (`TableRow.parent_body`, `TableCell._parent_body` through `Body.body_for`) | passed by |
| edits: `TableRow.insert_rows`, `delete`, `values`, `is_header`; `TableCell.value`, `width`, `insert_paragraph`, `insert_text` and its body's verbs; `Table.values`, `header_row_count`, `delete_rows` | `address.read_only` for a row or cell in a kept branch, refused before anything changes (`_refuse_in_branch`; the two setters that open no report ask `refuse_read_only`) | not reachable |
| tracked deletions: `Table.delete`, `Table.delete_rows`, `TableRow.delete`, `Body.clear` (through `tracking.track_deleted_row`, `delete_row_content`, `_delete_blocks`, `Body._row_elements`) | refused with `address.read_only` when they would mark content in a kept block-, row- or cell-level element (`kept_content_under`), since `cells_of` and the cells' blocks reach only Word's branch; untracked, the element goes whole | the row walk passed a row- or cell-level element by; `Body.clear` over a block-level element raised `address.read_only` part-way, at the branch's first paragraph, the paragraphs before it already marked (a Phase A gap, closed here) |
| `Table.add_rows`, `tracking.track_inserted_table` | unchanged: new rows go into the table's own content; a new table holds no element | the same |
| `table.cell_of` (`Paragraph.parent_table_cell`, `Table.parent_table_cell`) | the parent climb passes over the `mc` wrappers and finds the cell view in a branch | None: `rows_of` did not hold the row |
| `reports._table_shape`; `outline()`'s table walk (`_cells_with_paths`, new) | rows and columns of Word's branch; a cell's blocks are the table's children at their paths through any wrapper, the element or a row- or cell-level control, which is a path step and no entry of its own | rows and cells by index, so a row-level control's (and now an element's) rows came out as `tc` entries |
| `addresses.path_of`, `resolve_path`, `_addresses_under`; `Body.iter_paragraphs`, `paragraphs`, `element_at`; `comments.markers_of`; `controls._collect`; `lists_of`, `inline_pictures` (through the paragraphs) | through `block_children_of`, so through Word's branch, the views read-only | passed by |
| `ContentControl.form` | a leading kept element is read through to Word's branch | an `SdtBlock` whose content began with a block-level element read `"run"` (`mc-sdt-block-newest-first`; a Phase A gap) |
| `ContentControl._ancestor_table`, `table._ancestor` | parent climbs by element name, over the `mc` wrappers | the same |
| `api_script._table_grid` | direct children by name: an element among rows or cells is not expressible, so the whole table goes to `insert_ooxml` as XML | the same |
| `describe()` (`_styles_in_use`, `_revision_authors`) | `iter_nodes(mce="all")` now reaches a cell-level branch's typed cells, so a style or a revision's author there is counted (a built case: author `Zed` found, where it was not) | a wildcard `w:tc`'s content was not walked |
| the python-docx facade (`docx/table.py` `_rows_of`, `_cells_of`) | direct `w:tr` and `w:tc` children, as python-docx's `tr_lst` and `tc_lst`: passes the element by | the same |
| the property resolver | no parent climb in this port: it is handed a `w:pPr` and does not know the table | the same |
| `traversal` (`walk`, `iter_nodes`, `find`, `text_of`), `child.iter_children`, `link_parents` | level-blind as before; a cell-level branch's cells are typed, so `find(P)` reaches their paragraphs, and their parents are linked | a wildcard tree: its text read by name, `find(P)` not reaching inside |
| CR-010's islands | a text box straight in a run of a cell-level branch's cell (a bare `w:drawing`) is typed now and needs no island; one in a run-level `mc:AlternateContent` there is a wildcard tree and an island, as anywhere (Phase B) | the whole cell a wildcard tree: an island either way |

### 10.4 docx4j's tests at `9db1a3cbd`

Ported against the probes, as section 4.4 item 6 lists them.

| docx4j | here |
|---|---|
| `McBlockLevelLoadTest.preserveModeKeepsBlockLevelElements`, every container | `test_preserve_mode_keeps_every_container_and_the_text_model_reads_words_branch` (17): kept whole, nothing recorded, every branch, its content and the prefixes written, whole after a round trip, the text the same in both modes and Word's branch's |
| `McBlockLevelProbesTest.preserveModeKeepsEveryBranch` without its exception | already held for all 19 by Phase A's `test_preserve_saves_each_probe_canonically_identical` |
| `McBlockLevelProbesTest.preserveModeReadsAsWordsSave` | `test_preserve_mode_reads_as_word_365s_save` (14): before and after a `preserve` save |
| `McBranchSdtTest` (`staxPath`'s `check`, `xmlUtilsUnmarshalString`) | `test_a_sdt_in_either_branch_is_its_containers_kind_at_every_level` (both modes) and `test_a_fragment_types_a_sdt_in_either_branch_at_every_level`: `SdtBlock` ×4, `SdtRun` ×2, `CTSdtRow` ×2, `CTSdtCell` ×2 in `preserve`; the binder and preprocessor paths are JAXB's and have no counterpart |
| `McBlockLevelReadersTest.textModelWarnsFromTheRecord`, its probe | `_readers_probe()` built as docx4j builds it; `test_the_text_model_warns_from_the_record_at_row_and_cell_level_too`: `body/3/1` and `body/3/2/1` among nine or more, the same in both modes |
| `TrackedChangesAlternateContentTest`'s row and cell cases and two row-removal cases | `tests/revisions/test_alternate_content.py`: six for the cases, two for the removal |

Beyond docx4j's tests, the new ones cover:

- the typing (`test_a_row_or_cell_level_branch_holds_typed_rows_and_cells`, six containers);
- the record's parse;
- the content API's rows, cells and refusals, including the tracked deletions;
- `word` mode's edits;
- `outline()`, Markdown and `traversal`;
- the control's `form`;
- a row's and a nested table's revision text.

### 10.5 The numbers

- **The suite:**
  - the whole suite: 4,855 passed, 8 skipped (Phase A's 9.8: 4,782 and 8);
  - the 73 new tests: `test_mce_levels.py` 63, `test_alternate_content.py` 10;
  - nothing else moved but 10.7's.
- **Unchanged:**
  - `scripts/roundtrip.py`: 50 of 50 canonically identical, 0 differences, 0 skipped, 34 benign
    respellings;
  - CR-011's ratchet: 1,291 pass, 0 fail, 293 pending of 1,584, goldens 12 of 12, and
    `conformance.tsv` unchanged (`git status` clean);
  - the agreement test: 638 passed;
  - CR-010's Word-save comparisons (`test_word_saves.py`, `test_compare_word_resolution.py`): 68
    passed;
  - the parity goldens: 277 passed, `PINS` empty;
  - CR-002 section 16's tests (`test_private_readers.py`, `test_adoption.py`): 727 passed;
  - `scripts/threads.py`: no configuration lost or altered a document.
- **`generate.sh --check`** was not run: no input of the generated code changed (no schema, name
  table or `el` table), so nothing was regenerated.
- **The fork:** 1,656 passed, 1 skipped (1,651 before).
- **`scripts/bench_marshal.py`** (`--runs 7`, three runs each, this machine):

  | | before | after |
  |---|---:|---:|
  | `Symbols.docx` parse | 196 to 210 ms | 194 to 199 ms |
  | `Symbols.docx` marshal | 251 to 263 ms | 249 to 253 ms |
  | the 14 main parts, parse | 242 to 254 ms | 236 to 241 ms |
  | the 14 main parts, marshal | 301 to 313 ms | 298 to 303 ms |

  Within the band: the hook's name test is one set lookup, made only for a child of a wildcard.
- **`import docx4j_py`:** 613 to 646 ms, against 629 to 654 before (seven runs each, `-X
  importtime`), 242 modules both times.
- **The acceptance artefacts:** `out/acceptance/` was regenerated and compared with a copy made
  before the first change (`out/acceptance-before-cr012a2/`).
  - `scripts/compare_acceptance.py`: 13 artefacts, 175 entries, 0 differing beyond the wall-clock
    dates;
  - entry by entry, with the dates scrubbed: 175 of 175 identical, 167 of them byte for byte.

### 10.6 Departures, with why

1. **No fork code** (10.2). The brief expected the hook's extension in the fork. The hook's names
   are this package's, so the extension is `runtime.py`'s, and the fork's item is tests and a
   record.
2. **A tracked deletion that would mark content in a kept branch is refused.** This covers a
   table, its rows, a row and a body's `clear()`.
   - docx4j has no content API, so there is nothing to follow.
   - The marks would land in Word's branch alone, which is the edit section 4.1 item 5 refuses.
   - An untracked deletion takes the element whole and is made.
   - `Body.clear()` under tracking refuses a block-level element too, before it changes anything.
     In Phase A it raised part-way, at the branch's first paragraph, leaving the paragraphs before
     it marked deleted.
3. **`outline()`'s table walk treats any wrapper among rows or cells as a path step with no entry
   of its own.** That covers the element and a row- or cell-level content control.
   - Phase A's walk went by index, so a row-level control's rows came out as `tc` entries. The
     element would have done the same.
   - Now a cell's blocks are the table's children at their full paths, as `rows_of` and `cells_of`
     read them.
   - No test pinned the old shape.
4. **`ContentControl.form`** reads a leading kept element through. That fixes Phase A's `"run"`
   for a block-level control (10.3).
5. **The python-docx facade passes row- and cell-level elements by**, as python-docx reads a
   table's own `w:tr` and a row's own `w:tc`. This is 9.9's departure 8, at these levels.
6. **A twin `w:cellIns` or `w:cellDel` in a branch Word does not read is not resolved** (10.8),
   as in docx4j before `60bd4b9ad`: `remove_cell` finds cells through `cells_of`, which reads one
   branch. **Closed on 2026-10-22 by the port of `60bd4b9ad`** (`dde142e`, section 10.10): the
   twin is removed from its own branch, and that branch's neighbour widened.
7. **OpenDoPE's two cases are not ported**, as section 4.4 item 5 says:
   - docx4j's `RemovalHandler.xslt` row count, `OpenDoPEHandler.emptyKeepingTableShape` and
     `McRowCellLevelOpenDoPETest`;
   - this package's template fill has no OpenDoPE conditions or repeats and removes no control
     (CR-003 section 17, `model/customxml/template.py`);
   - so neither a row nor a control is ever removed or emptied there, and there is nothing to
     port.

### 10.7 Pre-existing tests whose expectation changed, and why

All five are Phase A's pins, in `tests/openpackaging/test_mce_levels.py`. Each pinned the pass-by
that this phase replaces.

- `test_the_body_passes_row_and_cell_level_by` (4) is now
  `test_the_body_reads_words_branch_at_row_and_cell_level`.
  - **Then:** no branch paragraph in `Body.paragraphs`, and only the plain rows and cells.
  - **Now:** the `w16sdtfl` paragraph is there, and the rows and cells are Word's branch's.
- `test_the_text_model_passes_row_and_cell_level_by` (4) is now
  `test_the_text_model_reads_words_branch_at_row_and_cell_level`.
  - **Then:** `I prefer\tbelow` with no branch row, `I prefer` with no branch cell, and `word`
    mode giving something else.
  - **Now:** `row\tw16sdtfl` and `I prefer\tw16sdtfl`, the same in both modes and the same as Word
    365's save.
- `test_the_revisions_engine_passes_row_and_cell_level_by` (1) is now
  `test_the_revisions_engine_walks_the_row_and_cell_probes`.
  - The probes hold no tracked change, so the expectation is unchanged: nothing listed, nothing
    accepted.
  - The name said "passes by", which is no longer so. It now takes all five row and cell probes in
    both modes.
- `test_the_text_models_warnings_are_the_same_in_both_modes_on_the_probes` now covers all 19 probes.
  It left out the four row and cell probes, whose warning Phase A did not raise.
- `test_the_text_models_warnings_are_the_same_in_both_modes_in_every_block_container` (11) is now
  `..._in_every_container` (17). The row and cell containers raise their one warning in both modes
  now.

Nothing else in the suite moved. The fork's upstream tests are untouched.

### 10.8 Found, and not done

- **A twin cell marker in an unread branch.** Take a `w:cellDel` (or `w:cellIns`) in each branch of a
  cell-level element. Accepting or rejecting the listed change removes Word's branch's cell, but
  the twin in the Fallback keeps its cell and its marker.
  - `tables.remove_cell` finds cells through `cells_of`, which reads Word's branch, so it does not
    find the twin.
  - docx4j's `Tables.removeCell` over `Lister.cellsOf` does the same.
  - Finding the twin is easy. Widening its neighbour is not: the neighbour may stand outside the
    element and be shared by every branch, and the primary's removal has already widened it.
  - Left, with docx4j, and worth telling docx4j-25. Word never writes a row- or cell-level element
    (CR-051 phase 0).
  - **docx4j fixed it in `60bd4b9ad`** (2026-10-22, on `VERSION_17_3_4`):
    - `Tables.removeCell` looks a cell that is not among `Lister.cellsOf(tr)` up in its own branch
      (`Lister.branchOf`, `Lister.cellsIn`), removes it and widens that branch's neighbour;
    - a twin alone in its branch is removed without widening;
    - its tests are `acceptingADeletedCellResolvesItsTwinInItsOwnBranch` and
      `acceptingADeletedCellAloneInEachBranchWidensTheRowsNeighbourOnce`;
    - one asymmetry is left (CR-051 section 12.5): a read branch of two cells beside a twin branch
      of one leaves the twin branch's grid a span short.
    - **Ported on 2026-10-22** at the user's go ("go on the port"), in `dde142e` (the code and
      its eight tests) and `85cc8eb` (section 10.10, the CHANGELOG, `tests/README.md`). The
      asymmetry is pinned as a known gap.
- **Phase B's typing.** A `w:p`, `w:tbl` or `w:r` in a branch is as before:
  - `w:p` and `w:r` are global, so typed already;
  - a `w:tbl` in a block-level branch is still a wildcard, so `Body.tables` passes a table there by
    (9.11); the text model reads it through a typed copy (`Extractor._typed_block`).
- **The inspector.** CR-050's inspector, docx4j's other reader of the record at row and cell level
  (`inspectorFindsTheRecordedBranchesWherePreserveModeHasThem`), has no port here (9.11).
- **docx4j's 17.4.0 switch (its 12.4)** is not this package's: `preserve` has been the default here
  since Phase A (Q7).
- **CLAUDE.md** is not edited, by the brief. Its CR-012 status line could say that Phase A2 is
  implemented, its Word check awaited. The Status line of this CR and the change-requests README
  row are left to the coordinator, by the brief.
- **A fork release.** Stages 9 to 11 are local, item 27 among them, as 9.11 says.

### 10.9 The Word check (passed 2026-10-22)

Six files are in `out/cr012-a2-word-check/`, with a `README.txt` saying what each should show.
They are also copied to `~/vbShares/Office 2016/__tmp/python-cr012a2/`. Each probe was loaded, its
main part unmarshalled so that this engine writes it (the branch rows and cells from typed `Tr` and
`Tc`), and saved.

- In `preserve` mode, five saves, each canonically identical to its source:
  - `cr012a2-1-row-newest-first-preserve-mode.docx`;
  - `cr012a2-2-sdt-row-newest-first-preserve-mode.docx`;
  - `cr012a2-3-cell-newest-first-preserve-mode.docx`;
  - `cr012a2-4-sdt-cell-newest-first-preserve-mode.docx`;
  - `cr012a2-5-tc-block-newest-first-preserve-mode.docx`.

  These are the five row and cell probes docx4j checked (its `mc-fun/<name>-docx4j-preserve.docx`
  on the share). The fifth is `mc-tc-block-newest-first`, a block-level element in a cell.
- In `word` mode: `cr012a2-6-row-newest-first-word-mode.docx`. Its row-level element is replaced by
  the `w16sdtfl` row, and the root's `mc:Ignorable` is `w16sdtfl`.

Expected: each opens without a repair prompt and shows Word 365's branch, `w16sdtfl`, where the
element stands, as Word 365's own save of the probe does. Each already reads so in the text model.
**Passed on 2026-10-22** (the user: "python-cr012a2 each pass"): all six open with no repair prompt
and show what the README says.

### 10.10 The cell-twin port (2026-10-22)

This is the port of docx4j `60bd4b9ad` (on `VERSION_17_3_4`, read by commit), at the user's go
("go on the port"). It closes 10.8's first item and 10.6's sixth departure. docx4j made the fix
after this package found the gap while porting Phase A2.

**Commits.**

- `dde142e`: `tables.branch_of`, `tables.cells_in` and the twin's lookup in `remove_cell`, with
  the tests;
- `85cc8eb`: this section, 10.6 item 6, 10.8, the CHANGELOG and `tests/README.md`.

**What landed** (`docx4j_py/model/revisions/tables.py`, beside `cells_of` and `_cells`):

- `branch_of(tc, tr)` is docx4j's `Lister.branchOf`. It climbs the parent pointers from a cell to
  its row and answers the `mc:Choice` or `mc:Fallback` it meets, typed or wildcard. A cell standing
  in the row itself, or in a cell-level control or custom XML element there, answers None.
- `cells_in(branch)` is docx4j's `Lister.cellsIn`. It gives a branch's cells, each with the live
  list holding it, through `_cells`, so through cell-level controls and custom XML. The list is a
  typed branch's `any_element` (a wildcard one's `children`), what `McSplice.contentOf` gives
  docx4j. `mce_branch_content` copies the list, so it is not used.
- `remove_cell` looks up a cell that is not among `cells_of(tr)` in its own branch. It removes the
  cell from that branch's list and widens that branch's neighbour, the cell before it, else the one
  after it. A twin alone in its branch is removed with nothing widened. `_index_in` is docx4j's
  `Tables.indexOf`.

Two things were checked before the port:

- **Parent pointers.** A branch's typed `Tc` has its branch as `parent`, both after
  `set_contents(wml(...))` and in a document saved and loaded again. The branch's parent is the
  element, and the element's the row. So `branch_of` needed nothing new.
- **The order of resolution.** `TrackedChange.accept()` and `reject()` resolve the listed change
  and then its twins, so the read branch's cell is gone when the twin is looked up. `accept_all`
  and `reject_all` resolve last first, so there the twin goes while the read branch's cell is still
  in place. Both orders give the same markup, and the tests run both.

**Tests** (`tests/revisions/test_alternate_content.py`, 19 to 27). The documents are built as
docx4j's test builds them, from its `TBL_PROPS` and its `cell`, `deletedCell` and `plainCell`.
The serialiser writes `w:gridSpan w:val="2"` and `w:tcW w:w="2000" w:type="dxa"`, so docx4j's
counts of serialised markup are asserted unchanged.

- docx4j's two, with its assertions:
  - `test_accepting_a_deleted_cell_resolves_its_twin_in_its_own_branch`: one change, with one
    twin; after the accept no `cellDel`, `>C<` or `>F<`, `Ckeep` and `Fkeep` kept, two
    `w:gridSpan w:val="2"`, two `w:tcW w:w="2000"`, and nothing listed;
  - `test_accepting_a_deleted_cell_alone_in_each_branch_widens_the_rows_neighbour_once`: one
    `w:gridSpan w:val="2"`, one `w:tcW w:w="2000"`, and the element kept, both branches empty of
    cells (asserted through `cells_in`; the serialiser writes `<mc:Choice Requires="w14"/>` and
    `<mc:Fallback/>`).
- Beyond docx4j's:
  - `w:cellIns` rejected, in both shapes, with the same assertions. The commit names the case, and
    docx4j has no test of it.
  - `accept_all` and `reject_all`. Each removes the marked cells in both shapes, with the counts
    above. Over the first shape with the other marker, each keeps every cell, the markers gone and
    nothing widened.
  - the pinned asymmetry, below;
  - `word` mode. The element is replaced on load by Word's branch, so there is one change with no
    twin, and accepting it widens `Ckeep` once.

Without the change, seven of the eight fail. The eighth, `word` mode, passes, as it should: that
mode has no twin.

**The numbers.**

- the fast suite (`-m "not slow"`): 3,583 passed, 8 skipped (3,575 and 8 before);
- `tests/revisions` and `tests/openpackaging/test_mce_levels.py`: 856 passed;
- CR-011's ratchet (`tests/text/conformance.py`): 1,291 pass, 0 fail, 293 pending of 1,584,
  goldens 12 of 12, `conformance.tsv` unchanged; `tests/text/test_conformance.py` 24 passed;
- the acceptance artefacts: `out/acceptance/` was regenerated and compared with a copy made before
  the change (`out/acceptance-before-cr012-twin/`). `scripts/compare_acceptance.py` found 13
  artefacts and 175 entries, 0 differing beyond the wall-clock dates;
- `generate.sh --check` was not run: no input of the generated code changed.

**Departures, with why.** Both are in form only. The behaviour is docx4j's.

1. **docx4j's guard's second clause is dropped.** docx4j returns after the removal when
   `neighbour == null || inOtherBranch && cells.size() == 1`. The second clause can never decide:
   - `cells` is the twin's branch's list, built before the removal, and the twin is in it;
   - if it holds one cell, that cell is the twin, at index 0;
   - then there is no cell at `at - 1` and none at `at + 1`, so `neighbour` is null already.

   So the port tests `neighbour is None` alone. It keeps docx4j's reasoning in a comment: the
   row's neighbour was widened when the read branch's cell went, and every branch spans the same
   grid columns. Without the clause, `inOtherBranch` has no other use, so there is no flag either.
2. **`cells_in` takes no row.** docx4j's `cellsIn(branch, tr)` passes the row to the `Row` records
   its `cells` builds. Here `_cells` builds `(cell, list)` pairs, which need no row.

**The asymmetry, pinned.** docx4j leaves it (its CR-051 section 12.5) and so does the port. The test
is `test_known_gap_pinned_a_lone_twin_beside_a_read_branch_of_two_cells_leaves_its_branch_short`.

- The shape: a Choice holding a deleted cell and a plain one (`Ckeep`), and a Fallback holding
  only the twin, which spans the same two grid columns (`w:gridSpan 2`). A plain cell (`keep`)
  follows the element.
- Before the accept, the row spans 3 grid columns read through either branch.
- Accepted, the read branch's cell goes, and its neighbour in its own branch, `Ckeep`, is widened.
  The twin, alone in its branch, goes with nothing widened, on the rule that the row's neighbour
  was widened already. Here it was not: `keep` stays one column, 1000 twips.
- After the accept, the row spans 3 columns through the Choice and 1 through the Fallback. The
  Fallback is short by the twin's span.
- No producer writes the shape: Word writes no row- or cell-level element at all (docx4j CR-051
  phase 0).
- A fix would widen the row's neighbour for a lone twin only when the read branch's cell had a
  neighbour in its own branch. That is docx4j's to decide first.

**The Word check (passed 2026-10-22).** Three files are in `out/cr012-a2-twin-word-check/`, with a
`README.txt` saying, for each, the source, what was done and what Word should show. They are also
copied to `~/vbShares/Office 2016/__tmp/python-cr012a2-twin/`.

- The sources are the tests' shapes, built with `create_package()` and `set_contents()`, with a
  heading paragraph before the table and one after it. The grid has two 1000-twip columns, where
  the tests have one, with single borders, so that a widened cell shows across both columns.
- The listed change was resolved with `accept()` or `reject()`, and the package saved from the
  typed objects.
- Each `w:document` declares `xmlns:w14`, which the Choice's `Requires` names. Each was read back
  with nothing skipped and no change left, and the text model reads the w14 Choice.

The three files:

- `cr012a2-twin-1-deleted-cell-beside-accepted.docx`: the first shape, accepted. Each branch is one
  cell of `w:gridSpan 2` and `w:tcW 2000`, `Ckeep` and `Fkeep`.
- `cr012a2-twin-2-deleted-cell-alone-accepted.docx`: the second shape, accepted. `keep` is widened
  once, and the element is left with an empty Choice and an empty Fallback.
- `cr012a2-twin-3-inserted-cell-beside-rejected.docx`: the first shape with `w:cellIns`, rejected,
  as file 1.

Expected: each opens with no repair prompt and no tracked change. The table shows the Choice's
cells, one cell across both columns: `Ckeep` in files 1 and 3, `keep` in file 2. None of `C`, `F`
or `Fkeep` shows. **Passed on 2026-10-22** (the user: "python-cr012a2-twin passes"): all three open
with no repair prompt and show what the README says, file 2's empty Choice and Fallback among them.

**Found, and not done.**

- **A table's property records in a branch Word does not read are neither listed nor resolved.**
  This is the same one-branch reading as the gap closed here, in the table records:
  - `record_of`, `in_change`, `has_records` and `reject` read cells through `cells_of` and rows
    through `rows_of`, so through Word's branch alone;
  - so a `w:tcPrChange` in a cell-level element's Fallback is never seen. Probed: with one in
    each branch, the table's change is listed once, with no twin. After `accept_all`, and after
    `reject_all`, the Fallback's cell keeps its `w:tcPrChange` and its properties;
  - a `w:trPrChange` in a row-level element's Fallback is the same (probed);
  - `collapse_grid` recounts Word's branch's spans alone, by reading;
  - docx4j at `60bd4b9ad` reads them the same way (`Tables.cells` and `rows` over
    `Lister.cellsOf` and `rowsOf`), by reading its code at that commit, not by a run;
  - worth telling docx4j-25. No producer writes the shape.
  - **Told docx4j-25 on 2026-10-22, and confirmed there.** docx4j fixed it the same day, in
    `901a922a3` (on `VERSION_17_3_4`), with two tests,
    `resolvingATablePropertyRecordReachesEveryBranchsCell` and
    `acceptingATablePropertyRecordReachesARowInAnotherBranch`:
    - `Tables.rows` and `cells` read every branch (`Lister.allRowsOf`, `allCellsOf`) for
      `recordOf`, `inChange`, `hasRecords`, accept and reject;
    - `collapseGrid` counts along `Lister.cellLayoutsOf(tr)`, one cell list per branch reading of
      the row;
    - `removeCell` still reads Word's branch for the neighbour;
    - docx4j also dropped the dead clause (departure 1 above).

    **Closed on 2026-10-22 by the port of `901a922a3`** (`4cd1798`, section 10.11), at the user's
    go ("port the fix"): the records are listed and resolved in every branch, and `collapse_grid`
    counts the grid along every branch reading of a row.

### 10.11 The table records port (2026-10-22)

This is the port of docx4j `901a922a3` (on `VERSION_17_3_4`, read by commit), at the user's go
("port the fix"). It closes 10.10's "Found, and not done" item. docx4j made the fix after this
package found the gap while porting `60bd4b9ad`.

**Commits.**

- `4cd1798`: `tables.all_rows_of`, `all_cells_of` and `cell_layouts_of`, the record paths and
  `collapse_grid` over them, with the tests;
- `d70b6ab`: this section, 10.10's item, the CHANGELOG and `tests/README.md`.

**What landed** (`docx4j_py/model/revisions/tables.py`, beside `rows_of`, `cells_of`, `cells_in`
and `branch_of`):

- `all_rows_of(tbl)` is docx4j's `Lister.allRowsOf`, and `all_cells_of(tr)` its `allCellsOf`.
  - They read as `rows_of` and `cells_of` read, through row- and cell-level controls and custom
    XML, but through every branch of an `mc:AlternateContent` (docx4j's `McMode.ALL`: its
    Choices, then its Fallback), where `rows_of` and `cells_of` read Word's.
  - Each gives `(row or cell, live list)` pairs. The order across branches means nothing.
- The record paths read them where they read `rows_of` and `cells_of`: `in_change`,
  `has_records`, `_row_differs`, `record_of`, `change_text`, `accept` and `reject`. These are
  docx4j's paths over `Tables.rows` and `cells`, which `901a922a3` points at `allRowsOf` and
  `allCellsOf`.
- `cell_layouts_of(tr)` is docx4j's `Lister.cellLayoutsOf`: one cell list per reading of the row.
  - Word's branch is first (`cells_of`).
  - Then, for each cell-level element among the cells and each of its branches other than Word's,
    the row read with that branch in the element's place. Every other element is still read by
    Word's branch.
  - The elements are found through every branch, controls and custom XML (`_cell_elements`,
    docx4j's `cellElements`); a row is read with one branch swapped by `_cells_reading` (docx4j's
    `cellsReading`). `_layouts(tr)` is docx4j's `Tables.layouts`.
- `collapse_grid` reads every branch's rows (`all_rows_of`) and gathers the grid edges along every
  layout of each row.
  - It sets the new spans along each layout, the last first, so Word's branch is done last.
  - The row's `w:gridBefore` is the row's own, read once before the layouts.
  - Its `w:gridAfter` is counted from where Word's branch's layout ends.
- `remove_cell` still reads Word's branch (`cells_of`) for its neighbour, then the twin's own
  branch (`cells_in`), as since 10.10. Its comment now says docx4j dropped the dead clause too.
- The helpers are `_every_branch(ac)` (docx4j `McSelection.branches(ac, McMode.ALL)`) and
  `_content_of(branch)` (docx4j `McSplice.contentOf`, a branch's live list). `has_rows` and
  `cells_in` use them now too, with no change in what they read.
- docx4j's dead guard clause is gone in `901a922a3`. 10.10's departure 1 dropped it here already,
  so docx4j now agrees and there was nothing to port.

**Where `collapse_grid` is reached.** `reject` calls it when the table's grid has a
`w:tblGridChange` (check 27, C23).

- A grid record alone has no author, so it is swept, not listed (`record_of`).
- So the tests give the table a `w:tblPrChange` as well, and reject its one change.
- So the public API reaches `collapse_grid` with a kept element, and it is tested there, not as a
  private function.

**Tests** (`tests/revisions/test_alternate_content.py`, 28 to 37). docx4j's documents are built as
its test builds them, from its `TBL_PROPS`, `cell`, `plainCell` and `plainRow`, and the new
`changedCell`: a red-shaded cell whose `w:tcPrChange` records it unshaded.

- docx4j's two, with its assertions:
  - `test_resolving_a_table_property_record_reaches_every_branchs_cell`: a `w:tcPrChange` in each
    branch's cell of a cell-level element.
    - One change is listed.
    - After `accept_all`: no `tcPrChange`, and two `w:fill="FF0000"`.
    - After `reject_all`: no `tcPrChange`, no fill, and two `<w:tcW w:w="1000"`.
  - `test_accepting_a_table_property_record_reaches_a_row_in_another_branch`: a `w:trPrChange`
    only on the Fallback's row of a row-level element, the row now `w:cantSplit` and recorded with
    no properties. One change is listed. After `accept_all`: no `trPrChange`, `<w:cantSplit/>`
    kept, and nothing listed.
- Beyond docx4j's:
  - The listed change's own `accept()` and `reject()` on the first shape, with docx4j's
    assertions. The change has no twin, and its text is the row's as Word reads it, `C\r\n`.
  - The Fallback row's `w:trPrChange` rejected, by its change and by `reject_all`. The row's
    `w:trPr` goes whole, since the record held none. The Choice's row is as it was. The change's
    text is `F\r\n`.
  - `has_records`, `in_change` and `record_of` for a record that only an unread branch holds:
    - a `w:tcPrChange` only in the Fallback's cell: the row is in the change, the table has
      records, and the record is the Fallback's; accepted, it goes;
    - a `w:trPrChange` only on the Fallback's row: the table has records; `all_rows_of` gives the
      Fallback's row in the change, `rows_of` only the Choice's.
  - `collapse_grid` along every layout: three tests, below.
  - The change's text over a row-level element, pinned (below).
  - `word` mode: the element is replaced on load by Word's branch, so there is one cell, one
    record and no twin. Accepted, one fill is left; rejected, none.

The three `collapse_grid` tests. Each recorded grid is of 1000-twip columns, and each table's
change is rejected:

- **Cell level.** A recorded grid of six. The row skips two columns (`w:gridBefore 2`); the
  Choice's cells span 1 and 3, the Fallback's 2 and 2.
  - Rejected: four columns (2000, 1000, 1000, 2000), `w:gridBefore 1`, the Choice's spans 1 and
    2, the Fallback's 2 and 1.
  - With Word's branch alone: three columns, and the Fallback left spanning five of them.
- **`w:gridAfter` from Word's branch**, in a shape no producer writes. The branches end at
  different columns, so that the rule decides.
  - A recorded grid of four. The Choice's cell spans 3, the Fallback's 2, and the row has
    `w:gridAfter 1` and `w:wAfter`.
  - Rejected: two columns (2000 and 1000), the fourth dropped. Word's branch ends on the last edge
    kept, so `w:gridAfter` and `w:wAfter` go, and the row's `w:trPr` with them.
  - Counted from the Fallback's end, `w:gridAfter` would stay 1. With the layouts walked first to
    last, the test fails.
- **Row level.** A recorded grid of four. The Choice's row spans 1 and 3, the Fallback's 2 and 2.
  Rejected: three columns (1000, 1000, 2000); the Choice's row spans 1 and 2, the Fallback's 2
  and 1.

Without the change, nine of the ten fail. The tenth, `word` mode, passes, as it should: that mode
has one branch. With `all_rows_of` and `all_cells_of` in place and only `collapse_grid`'s layouts
reverted to `cells_of`, the two cell-level grid tests fail and the row-level one passes.

**The numbers.**

- the fast suite (`-m "not slow"`): 3,593 passed, 8 skipped (3,583 and 8 before);
- `tests/revisions` and `tests/openpackaging/test_mce_levels.py`: 866 passed (856 before);
- CR-011's ratchet (`tests/text/conformance.py`): 1,291 pass, 0 fail, 293 pending of 1,584,
  goldens 12 of 12, `conformance.tsv` unchanged; `tests/text/test_conformance.py` 24 passed;
- the acceptance artefacts: `out/acceptance/` was regenerated and compared with a copy made before
  the change (`out/acceptance-before-cr012-records/`). `scripts/compare_acceptance.py` found 13
  artefacts and 175 entries, 0 differing beyond the wall-clock dates;
- `generate.sh --check` was not run: no input of the generated code changed.

**Departures, with why.** All are in form only. The behaviour is docx4j's.

1. **No `Tables.rows` and `cells` wrappers.** docx4j's record paths read `Tables.rows` and
   `cells`, private wrappers over `Lister`'s lists, and `901a922a3` repoints those two.
   - This module has no such wrappers: its record paths read `rows_of` and `cells_of` directly,
     and now `all_rows_of` and `all_cells_of`.
   - The names `_rows` and `_cells` here are `Lister`'s private walkers already.
   - `Tables.layouts` is `_layouts`.
2. **Pairs, with no row passed.** `all_cells_of`, `cell_layouts_of` and their walkers take no row.
   This is as `cells_in` takes none (10.10 departure 2): the `(cell, list)` pairs need none.
3. **Word's branch is compared by identity.** docx4j skips the branch `McSelection.branches(ac,
   McMode.READ)` gives with `List.contains`.
   - Here the one branch `mce_branch` gives is skipped by `is`, as the lister's `_alternate_level`
     skips it.
   - A dataclass's `==` compares fields. So two Choices with the same content would compare equal,
     and `in` would skip both.
4. **Every branch through `child.py`'s helpers.** `McMode.ALL` is `mce_choices`, then
   `mce_fallback`, as `has_rows` read them.
   - The lister's `_branch_objects` gives the same branches.
   - It is not imported, since the lister imports this module.

**Found, and not done: the text of a table's change.**

- docx4j's `changeText` reads the table's rows through `Tables.rows`, which `901a922a3` points at
  every branch. The commit's message names the record paths (`recordOf`, `inChange`,
  `hasRecords`, accept, reject, `collapseGrid`), not the text.
- For a changed row only a Fallback holds, that is wanted. Its change's text is that row's
  (`F\r\n`, tested), where reading Word's branch would give none.
- For a table whose own properties or grid changed, the change's text is every row's.
  - So it is now every branch's row: the Choice's and the Fallback's.
  - In the pinned test that is `A\tB\r\nX\tY\r\n`, where Word reads `A\tB\r\n`.
- The port keeps docx4j's reading, and pins it:
  `test_known_gap_pinned_a_whole_tables_change_text_reads_every_branchs_rows`.
- A fix would read Word's branch for a whole table, and every branch only for the changed rows.
  That is docx4j's to decide first: worth telling docx4j-25. No producer writes the shape.
- **Closed on 2026-10-22 by the port of docx4j `84b202c5a`** (`b545884`, section 10.12), at the
  user's go ("go on the port"). docx4j found the item by reading this section, and decided it
  otherwise than the fix above:
  - its `changeText` reads `Lister.rowsOf`, Word's branch, for both halves, so `change_text` reads
    `rows_of` again;
  - so the whole table's text is the Choice's row, `A\tB\r\n`, and the pin is now the expectation,
    `test_a_whole_tables_change_text_reads_the_branch_word_reads`;
  - and the change of a record only the Fallback's row holds has no text, `""`, where this
    section's test had `F\r\n`;
  - the record paths stay on `all_rows_of`.

**The Word check (passed 2026-10-22).** Three files are in `out/cr012-a2-records-word-check/`, with
a `README.txt` saying, for each, the source, what was done and what Word should show. They are also
copied to `~/vbShares/Office 2016/__tmp/python-cr012a2-records/`.

- The sources are docx4j's two tests' shapes, built with `create_package()` and `set_contents()`,
  with a heading paragraph before the table and one after it. The tests' `w:tblPr` is empty; here
  it has single borders on every edge, so that the table shows. The grid is one 1000-twip column.
- The listed change was resolved with `accept()` or `reject()`, and the package saved from the
  typed objects.
- Each `w:document` declares `xmlns:w14`, which the Choice's `Requires` names.
- Each was read back with `LoadOptions(strict=True)`: nothing skipped, no tracked change, no
  `w:tcPrChange` or `w:trPrChange` left. The text model reads the w14 Choice.

The three files:

- `cr012a2-records-1-cell-record-accepted.docx`: the cell-level shape, accepted. Both branches'
  cells keep the red shading, and neither has a record.
- `cr012a2-records-2-cell-record-rejected.docx`: the same shape, rejected. Both branches' cells
  are back to `w:tcW 1000`, unshaded.
- `cr012a2-records-3-fallback-row-record-accepted.docx`: the row-level shape, accepted. The
  Fallback's row keeps `w:cantSplit`, and has no record.

Expected: each opens with no repair prompt and no tracked change, and shows the Choice's cell, `C`:
red in file 1, plain in files 2 and 3. `F` is not shown. **Passed on 2026-10-22** (the user: "the
Word check of __tmp/python-cr012a2-records/ passes"): all three open with no repair prompt and show
what the README says.

### 10.12 The change text and grid port (2026-10-22)

This is the port of docx4j `84b202c5a` (on `VERSION_17_3_4`, read by commit), at the user's go
("go on the port"). It closes 10.11's "Found, and not done" item, and a defect in 10.11's
`collapse_grid` that this package shared with docx4j. docx4j made both fixes after the ports read
`901a922a3`: the change text from this package's 10.11, the grid from core-ts's port of the same
commit (its CR-013 phase 3).

**Commits.**

- `b545884`: `tables.change_text` on `rows_of`, and `collapse_grid`'s spans counted before any is
  written, with the tests;
- `44c7265`: this section, 10.11's item, the CHANGELOG and `tests/README.md`; `c8bb973`, the
  commits; `4f04e94`, the Word check.

The port was made in a git worktree while Phase B's agent worked in the main checkout, and its
commits were cherry-picked onto `main` after Phase B (there `ad5e885`, `017079b`, `5db2af6` and
`90df2c5`). The cherry-pick put this section before section 11; its text is unchanged.

**What landed** (`docx4j_py/model/revisions/tables.py`):

- `change_text` reads `rows_of` again (docx4j `Lister.rowsOf`, the branch Word reads), for both
  halves of its rule:
  - where the table's own properties or grid changed, the text is every row Word shows: a
    row-level element's Choice's rows, not its Fallback's;
  - otherwise it is the changed rows' among those. A row in the change that only a branch Word
    does not read holds is still listed and resolved, through the record paths, but adds no text,
    so its change's text is empty. docx4j points the whole loop at `Lister.rowsOf`; 10.11's
    suggested split (every branch for the changed rows) was not taken.
  - Whether a row is in the change still reads its cells through every branch (`in_change`,
    `_row_differs`, over `all_cells_of`), and a row's text still reads Word's branch of its cells
    (`row_text`).
  - The record paths stay on `all_rows_of` and `all_cells_of`: `in_change`, `has_records`,
    `_row_differs`, `record_of`, `accept` and `reject`.
- `collapse_grid` counts every layout's new spans before it writes any:
  - The counts are taken along each layout, the last first, from the spans as they were, into a
    map keyed by the cell's identity. A cell outside the element is in every layout, so it is
    counted in each, and the count along Word's layout, the last, is the one that stands.
  - Then each cell with a `w:tcPr` gets its count: a `w:gridSpan` when it is above 1, none
    otherwise.
  - Then the row's `w:gridBefore` and `w:gridAfter`. docx4j now sets both after the spans, where
    `w:gridBefore` was set before the loop. The order changes nothing: `before` is read once,
    before the layouts, and the counts do not read the row's members.
- **The defect** (core-ts's finding). The loop wrote each layout's spans as it counted them. A cell
  outside the element is in every layout. So Word's layout, counted last, read that cell with the
  span the Fallback's layout had already rewritten, in new-grid units, while the count was still
  in old-grid units. Every cell after it was then placed from the wrong edge.
  - It was probed here before the port, with the test file's `regridded(6, ...)` and `spanned`
    (columns of 1000), and with docx4j's own shape (columns of 500).
  - Either way the grid came back as 3 columns, which is right, and F got span 2. But C, the cell
    in Word's branch, got span 1 where 2 is right. Read through the Choice, the row spanned two
    columns of three.

**Tests** (`tests/revisions/test_alternate_content.py`, 36 to 39, and one changed). docx4j's
documents are built as its tests build them:

- its `w:tblPr`, centred and recorded with no properties (`CHANGED_TBL_PR`);
- its grids: three columns of 1000 now, six of 500 recorded (`regridded_in_500s`);
- its `spanned`: a `w:tcW` of 500 a column, with a `w:gridSpan` (`spanned_in_500s`);
- its `plainRow` (`plain_row`, already here).

The tests:

- docx4j's two, with its assertions:
  - `test_rejecting_a_grid_change_counts_every_branchs_spans_from_the_old_ones`: rejected, three
    `w:gridCol`, all `w:w="1000"`; two `w:gridSpan`, both `w:val="2"`; no `gridAfter` or
    `gridBefore`. Beyond docx4j's assertions: C and F are the cells spanning 2 (`cells_in`), and
    the row spans three columns through either branch.
  - `test_a_table_changes_text_reads_the_branch_word_reads`: one change, its text holding `C` and
    not `F`.
- The pin, turned into the expectation and renamed:
  `test_known_gap_pinned_a_whole_tables_change_text_reads_every_branchs_rows` is now
  `test_a_whole_tables_change_text_reads_the_branch_word_reads`. Its text is `A\tB\r\n`, where it
  was `A\tB\r\nX\tY\r\n`.
- Beyond docx4j's: `test_rejecting_a_grid_change_counts_a_cell_after_the_element_from_its_old_span`.
  - It puts the cell outside the element after it, where docx4j's shape has it before. Row 1 is
    the element, C | F over two old columns, then X over four. Row 2 is a, b and c over two each.
  - Rejected: three columns of 1000; X spans 2 and the others 1; the row spans three columns
    through either branch.
  - Before the fix X got span 1. The Fallback's layout wrote X's span 2, in new columns, and Word's
    then counted X from that: two old columns, where X spans four. Here the shared cell is
    miscounted itself; in docx4j's shape it is the cell after it that is placed from the wrong
    edge.
  - The shape gives X a span above 1 in the new grid on purpose. With C | F over four and X over
    two after them, probed, X's count comes out 1 either way and the old code passes by luck.
- Changed: 10.11's `test_rejecting_a_table_property_record_reaches_a_row_in_another_branch`. The
  Fallback row's change now has no text, `""`, where it had `F\r\n`. docx4j's per-row half reads
  `Lister.rowsOf` too, and the Choice's row is not in the change. Its other assertions are
  unchanged.

Without the fix, all five fail: the four new and the changed one. With only `change_text`
reverted, the three text tests fail; with only `collapse_grid` reverted, the two grid tests.

**The numbers.**

- the fast suite (`-m "not slow"`): 3,596 passed, 8 skipped (3,593 and 8 before);
- `tests/revisions` and `tests/openpackaging/test_mce_levels.py`: 869 passed (866 before);
- CR-011's ratchet (`tests/text/conformance.py`): 1,291 pass, 0 fail, 293 pending of 1,584,
  goldens 12 of 12, `conformance.tsv` unchanged;
- the acceptance artefacts: `out/acceptance/` was regenerated and compared with a copy made before
  the change (`out/acceptance-before/`). `scripts/compare_acceptance.py` found 13 artefacts and
  175 entries, 0 differing beyond the wall-clock dates;
- `generate.sh --check` was not run: no input of the generated code changed.

The port was made in a git worktree two levels below the checkout. There the tests that read the
sibling checkouts (`../docx4j`, `../docx4j-core-ts`) skip, so the two suite numbers above were
taken in a clone of the worktree placed beside links to those checkouts. That clone gives 3,593
passed and 8 skipped at the parent commit, as the checkout does.

**Departures, with why.** Both are in form only. The behaviour is docx4j's.

1. **The counts are keyed by `id()`.** docx4j's map is an `IdentityHashMap<Tc, Integer>`.
   - A generated dataclass is slotted with `eq=True`, so it is unhashable, and its `==` compares
     fields: two empty cells compare equal.
   - So the map here is `dict[int, tuple[Tc, int]]`, keyed by `id(tc)`, with the cell in the
     value. That keeps every cell alive for the loop, so no id is reused.
   - The cells are written in the dict's order, first counted first, where docx4j's identity map
     has no order. Each write touches one cell, so the order changes nothing.
2. **The test helpers' names.** docx4j's grid test's `spanned` is `spanned_in_500s` here, since the
   file's `spanned` (10.11's, a `w:gridSpan` alone) is taken. Its grids are `regridded_in_500s`,
   beside 10.11's `regridded` (columns of 1000). `element(choice, fallback)` builds the
   `mc:AlternateContent` for the new tests.

**The Word check (passed 2026-10-22).** One file is in `out/cr012-a2-grid-word-check/`, with a
`README.txt` saying the source, what was done and what Word should show. Both are also copied to
`~/vbShares/Office 2016/__tmp/python-cr012a2-grid/`.

- The source is docx4j's grid test's shape, built with `create_package()` and `set_contents()`,
  with a paragraph before the table and one after it.
- The test's `w:tblPr` is centred and recorded with no properties. Here the current and the
  recorded `w:tblPr` both have single borders on every edge, so that the table shows after the
  reject. The reject restores the recorded `w:tblPr` wholesale, so the table is not centred.
- `reject_all` rejected the table's one change, and the package was saved from the typed objects.
- The `w:document` declares `xmlns:w14`, which the Choice's `Requires` names.
- It was read back with `LoadOptions(strict=True)`:
  - nothing skipped, and no tracked change;
  - no `w:tblPrChange`, `w:tblGridChange`, `w:gridBefore` or `w:gridAfter`;
  - three `w:gridCol` of 1000; C and F span 2, and X, a, b and c span 1;
  - the text model reads the w14 Choice: the table is `X: a; C: b; C: c`, with no `F`.

The file is `cr012a2-grid-rejected.docx`.

Expected: it opens with no repair prompt and no tracked change. It shows a bordered table of three
equal columns: row 1 is X across one column and C across two, and row 2 is a, b and c. `F` is not
shown. **Passed on 2026-10-22** (the user: "python-cr012a2-grid passes"): it opens with no repair
prompt and shows what the README says.

## 11. Phase B implementation notes (2026-10-22)

Phase B types what a kept branch holds, as docx4j types it, at the user's go ("port the fix, then
start phase b"), on section 4.2 as revised before it started. docx4j's 61 WordprocessingML root
elements that this model does not declare globally are derived from docx4j's own classes, generated
into `docx4j_py/el_index.py` and handed to the fork's stage 11 hooks, with no fork code. In a branch
each is bound as the container holding the `mc:AlternateContent` binds it, and in any other lax
wildcard as its root class. A `w:drawing` in an `mc:Choice` is a `Drawing` now, down to the `wps`
shape and its `w:txbxContent`; a `w:pict` in an `mc:Fallback` is a `Pict`, its VML still a wildcard
tree (Phase C). The walks name a branch's typed child by its element name, and the readers read it
as they read the same element elsewhere.

**Commits.**

- docx4j-python:
  - `0d041e6`: `codegen/derive_names.py` writes `codegen/names/root_elements.json`; the table; the
    names README;
  - `44278d9`: `codegen/generate_el.py` writes `ROOT_ELEMENTS` and `ROOT_ELEMENT_BINDINGS` into
    `docx4j_py/el_index.py`; the regeneration;
  - `f02dcc7`: the runtime hook's three steps and the serialiser's element names; the `word`-mode
    record's hook;
  - `cb7581c`: the walks' names, and the readers (the text model's `_graphic_kind`,
    `Body.tables`, the python-docx facade's inline shapes, the docstrings of the wildcard readers);
    the two `walk_all` tests;
  - `69f61e0`: `tests/openpackaging/test_mce_typed.py`;
  - `a7ce5dd`: `scripts/mce_branches.py`;
  - `14d0e89`: the `text_of` readers' paragraph-level reading pinned for the coordinator (11.4);
  - `2bfcbd4`: this section, the CHANGELOG, CR-002 section 5.6's note,
    `tests/README.md` and `docs/UPSTREAM.md` report 24.
- The fork (`~/git/docx4j-xsdata`, branch `docx4j`): none. Its tests already pin what this phase
  relies on (11.7 item 8).

### 11.1 What landed

| piece | where | what |
|---|---|---|
| the derivation | `codegen/derive_names.py`, `codegen/names/root_elements.json`, `codegen/names/README.md` | docx4j's WordprocessingML roots not global here, each with the `el` table's class; the cross-check against the schema copy's annotations, which stops the run on a disagreement (11.2) |
| the generated table | `codegen/generate_el.py`, `docx4j_py/el_index.py` | `ROOT_ELEMENTS` (61 names to their root classes' paths) and `ROOT_ELEMENT_BINDINGS` (22 names to the 47 other classes a container binds them to), held to the model at generation |
| the hook | `docx4j_py/runtime.py` | `branch_child_class`: (a) `container_binding`, (b) `root_element_class`, (c) None; `wildcard_names()`, the parser's default `wildcard_class_for`; `element_names_table`, which refuses a class claimed by two names; `_branch_element_names` gains 107 classes, 113 in all |
| the record | `docx4j_py/openpackaging/mc_block_level.py` | `_record_hook` takes (a), then (a) with the level's container, then (b) |
| the walks | `docx4j_py/child.py` | `iter_children` names a wildcard field's typed child by `wildcard_child_name`, the serialiser's name for it, where it gave the field's name, `any_element` |
| the readers | `docx4j_py/model/text/extractor.py`, `model/content/body.py`, `docx/shape.py`, `model/revisions/islands.py` | `_graphic_kind` by element name; `Body.tables` through a block-level branch; the facade's `inline_shapes` without a branch's drawing; the docstrings say where each wildcard reader is still reached (11.4) |
| tests | `tests/openpackaging/test_mce_typed.py` (260); `tests/test_builders_phase_a.py` (2 changed) | 11.6, 11.8 |
| the survey | `scripts/mce_branches.py` | `corpus`, `survey`, `readers`, `compare-readers` (11.5) |
| documents | `CHANGELOG.md`, CR-002 section 5.6, `tests/README.md`, `docs/UPSTREAM.md` report 24 | |

### 11.2 The derivation and its provenance

**The source.** docx4j's XJC output in its checkout,
`docx4j-generated-objects/target/generated-sources/xjc`. It is a build, not a commit: its 2,709
`.java` files were written at 2026-10-20 14:11:44 +11:00 (`root_elements.json`'s `xjc_built`, the
newest file of `org.docx4j.wml`). That predates `1cff2cd5c` (17:32 the same day); docx4j's
`VERSION_17_3_4` was at `cf1abf001` then. The build's own commit is not recorded.

**The rule** (`build_root_elements`):

- docx4j's global names are every `@XmlRootElement`, nested classes included, and every
  `@XmlElementDecl` with no `scope`. The name tables' regexes see top-level classes only, so this
  pass has its own, which tracks a class's enclosing classes by indentation (`R.Tab`,
  `P.Hyperlink`, `PPrBase.Spacing`, `Numbering.Num.LvlOverride`).
- In the WordprocessingML namespace that is 99 names. The schema copy declares 37 globally (a
  top-level `xsd:element` of `wml/wml.xsd`), all 37 among the 99. So 62 are missing here, as
  section 2 counted.
- A name several docx4j classes claim as their root is left out. That is `w:sdtContent`, whose four
  classes a name alone cannot choose among: Q3's decision, made a rule.
- Each of the 61 left takes the class the `el` table names for it (`codegen/el_tables/wml.json`).
- The other vocabularies' missing names, in a namespace this model has, are listed and not taken:
  the custom properties' `property`, as section 2 said.

**The cross-check.** The names `wml/wml.xsd` annotates with `annox` `@XmlRootElement` and does not
declare globally must be exactly the 62 the rule finds, or the run stops with both lists. They are:
62 and 62. **`w:sdtPr` is annotated** in the schema copy, written `@XmlRootElement(name =
"sdtPr")` with spaces, and has been since the first copy. Section 2's count, which took it for the
one unannotated name, matched only the unspaced form (70 names that way, 71 with it). So the two
sources agree on all 62, and the cross-check ties the table to the schema copy's commit,
`2853b7280`. Between the XJC build's likely tree (`cf1abf001`) and `2853b7280`, `xsd/wml/wml.xsd`
changed only by `1cff2cd5c`, which touches no annotation; nothing under `xsd/` changed from
`2853b7280` to `901a922a3`.

**Not a `ClassNames` table.** The generator's `<ClassNames>` option reads every `*.json` in
`codegen/names/`. The file's entries are under `roots`, with no `types` or `elements` key, so the
generator takes nothing from it: the regeneration changed `el_index.py` alone, with the same 282
`ClassNames` warnings. Re-running `derive_names.py` rewrites the other 56 tables byte for byte, and
two runs write the same `root_elements.json`.

**Held to the model** (`generate_el.root_elements`). Each root must not be a global element of the
model, and its class must be the one `el` builds for the name, or the run stops and says to
re-derive. The bindings are what step (a) can give besides the root class. Its containers are the
122 classes that can hold an `mc:AlternateContent`: 45 by a field, 77 by a lax wildcard (a global
`mc:AlternateContent` is typed in any of them), the `mc` wrappers aside. A class bound to two names
stops the run; none is.

### 11.3 The (a)/(b) table

For each of the 61 names, every container's binding against the root class. "Containers" are the
122 of 11.2; a container binding none of the 61 is not listed.

| names | how many | step (a) gives | against (b) |
|---|---:|---|---|
| `w:altChunk` (10 containers), `w:bookmarkEnd`, `w:bookmarkStart`, `w:permEnd`, `w:permStart` (31 each), `w:drawing`, `w:pict` (6 each), `w:fldChar` (5), `w:fldSimple`, `w:smartTag` (14 each), `w:tbl` (11), `w:rPr` (5), `w:tc` (3), `w:sectPr` (2), `w:pPr`, `w:tblPr`, `w:tblGrid` (1 each) | 17 | the root class, in every container that binds the name | agree |
| `w:abstractNum`, `w:alias`, `w:bibliography`, `w:citation`, `w:dataBinding`, `w:equation`, `w:group`, `w:ind`, `w:jc`, `w:lock`, `w:lvl`, `w:lvlOverride`, `w:num`, `w:numPicBullet`, `w:numPr`, `w:outlineLvl`, `w:pBdr`, `w:picture`, `w:richText`, `w:sdtPr`, `w:spacing`, `w:tabs` | 22 | nothing: no container binds the name | (b) always |
| `w:annotationRef`, `w:commentReference`, `w:continuationSeparator`, `w:cr`, `w:dayLong`, `w:dayShort`, `w:endnoteRef`, `w:footnoteRef`, `w:lastRenderedPageBreak`, `w:monthLong`, `w:monthShort`, `w:noBreakHyphen`, `w:pgNum`, `w:ptab`, `w:separator`, `w:softHyphen`, `w:sym`, `w:tab`, `w:yearLong`, `w:yearShort` | 20 | the root class in `R`; in a math run (`CTR`) its `CtR...` class; in a math run's track change (`CTMathRunTrackChange` and its `Ins1`, `Del1`) its `CtMathRunTrackChange...` class | differ in the math run |
| `w:t` | 1 | `RT` in `R` and the math run's track change; `T` in a math run (`CTR`) | differ in the math run |
| `w:hyperlink` | 1 | each of its seven classes in its own container: `PHyperlink` in `P`, `CtSimpleFieldHyperlink` in `CTSimpleField`, and the bdo's, dir's, custom XML run's, run control's and smart tag's, each in its container and in itself (a hyperlink holding a branch) | differ for six of the seven kinds of container |

Section 4.2 item 1 expected them to agree except where this model generated several classes for one
type, such as `w:hyperlink`. That holds, and the exception is wider than the example:
- the run content names have three classes each, one per scope, as `w:hyperlink` has seven;
- the math run's own `w:t` is the `el` table's `t_math` class, `T`.

So 22 names take 47 other classes. Each is in the serialiser's table under its name
(`ROOT_ELEMENT_BINDINGS`). A test recomputes the containers apart from the generator and checks
every class step (a) gives against that table.

### 11.4 The readers' audit

Every reader of text or structure was run over every distinct corpus document holding
`mc:AlternateContent` in a WordprocessingML part (85, by content), at `f1bb1d9` before the first
change and at `14d0e89` after the last, and the answers diffed (`scripts/mce_branches.py readers`
and `compare-readers`). The readers: the content API's `Body.paragraphs` (addresses and text),
`tables` (values), `inline_pictures`, `element_at` (every paragraph's and table's address),
`outline()`, `describe()`, `get_text()`, `find` and `search`; the text model's `extract`,
`plain_text` and `markdown`, and `to_markdown`; `revisions.tracked_changes`; the python-docx
facade's `paragraphs`, `tables` and `inline_shapes`.

**The `text_of` readers read a paragraph-level branch's runs** (decided 2026-10-22 by the
coordinator: kept).

- `text_of` is behind `Block.text`, `Table.text`, `TableCell.text` (and so `Table.values`), a
  run-level control's text and `TrackedChange.text`.
- Before this phase it read a run-level branch's text, whose `w:t` was a wildcard read by name, but
  not a paragraph-level branch's runs: a typed `w:r` there was enumerated under the field's name,
  `any_element`, and no walk of text read it.
- Now it reads both, Word's branch, as its contract says ("only the branch a consumer that
  understands ... would take contributes") and as Word shows. A cell `a` + a branch's run
  `-PARA_CHOICE` reads `a-PARA_CHOICE` where it read `a`.
- `Paragraph.text` and `get_text`, Office JS's text by CR-011 Q4, read neither branch, before and
  after. So a cell's text and its paragraph's text can now differ at paragraph level, as they
  already could at run level.
- No corpus document has the shape. Kept, and tested:
  `test_the_text_of_readers_read_a_paragraph_level_branchs_runs_as_at_run_level`.
  Whether the `text_of` readers that call themselves Office JS's text should descend into
  `mc:AlternateContent` at all is CR-011 Q4's question.
- **The decision: kept.**
  - The `text_of` readers already read Word's branch at run level. This phase removes the one level
    at which they did not, so they now follow their own contract at every level. That is also what
    Word shows.
  - The split from `Paragraph.text`, `get_text` and the content API's inline pictures, which pass a
    branch by, is older than this phase: it existed at run level. Q4 kept those readers as they
    were, and this phase does not widen the split in kind.
  - What Office JS's own `paragraph.text` gives for a branch is measurable in Word (Script Lab over
    a probe). If it reads Word's branch, those readers should follow, which is a CR-011 Q4 revision
    for the user.

**The answers that changed over the corpus.**

1. **The python-docx facade's `inline_shapes`: 24 documents**, 488 drawings, each a `w:drawing`
   straight in a branch, joined the list (`tc-p3-transforms.docx` 403 of them). python-docx 1.2.0
   lists `//w:p/w:r/w:drawing/wp:inline` (`docx/shape.py`, `InlineShapes._inline_lst`): a drawing
   whose parent is an `mc:Choice` is no run's child, so **clearly wrong by python-docx's contract**.
   Acted on: the facade keeps a drawing straight in a branch out (`_straight_in_a_branch`), and
   the answer is the one before. No newly typed drawing stood inside a typed text box in a branch.
2. **`Table.values`, `Table.text` and `element_at`'s table text: 1 document**,
   `mc-txbx-tc-block.docx`. Its second cell holds a block-level element whose Choice paragraph has
   two runs, the text `CELL_CHOICE_PARA ` and a text box. It read `CELL_CHOICE_PARA \n`: the
   paragraph, typed `P` but named `any_element`, was not read as a paragraph, so each run was a
   line. Now `CELL_CHOICE_PARA `. **Clearly right**: a paragraph is one line, docx4j `TextUtils`'
   rule and `text_of`'s. Kept, and pinned (`test_a_paragraph_in_a_block_level_branch_is_one_line`).

Nothing else changed: the paragraphs, pictures, outline, `describe()`, `get_text`, `find`,
`search`, the text model, Markdown, the tracked changes and the facade's paragraphs and tables
answer as before on all 85.

**Beyond the corpus.** A probe built for the purpose holds a branch's `w:hyperlink`,
`w:fldSimple`, `w:smartTag`, `w:bookmarkStart`, `w:fldChar`, run content (`w:tab`, `w:sym`,
`w:noBreakHyphen`, `w:softHyphen`, `w:cr`), an inline picture, a `w:ins` and a table in a
block-level branch. Read with Phase A2's parser and with this phase's, through the code after this
phase:

- the text model, Markdown, `to_markdown`, the tracked changes, `describe()` and the facade answer
  the same both ways, so the text model's typed paths read what its wildcard paths read;
- the content API moves only for the table in the block-level branch. `Body.tables` lists it, read
  only (9.11's first item, below), and `Body.paragraphs`, `element_at`, `outline()` (a `table`
  entry where an `AnyElement` one stood), `find`, `search` and `get_text` reach its paragraph, as
  the text model already did. **Clearly right** by Phase A's rule: a reader of structure reads
  Word's branch at block level. Pinned
  (`test_the_content_api_reads_a_block_level_branchs_table_as_its_paragraphs`);
- `describe()` reads the parts' bytes with lxml, so typing cannot move it.

**9.11's two items.**

- **`Body.tables` reads a table in a block-level branch.** It still passed it by once the table
  was typed, since it read its own list's `w:tbl` items alone. It now reads through a block-level
  `mc:AlternateContent` to Word's branch, as `Body.paragraphs` does, its views read-only (every
  edit raises `address.read_only`, `add_rows` and `delete` included). Controls stay
  `control.tables`, as Office JS reports them. Tested
  (`test_body_tables_reads_a_table_in_a_block_level_branch`).
- **`find(P)` reaches a cell-level branch's paragraphs**: since Phase A2's typed cells, tested
  there (`test_mce_levels.py::test_traversal_reaches_a_cell_level_branchs_paragraphs`). Nothing more needed.

**Code that may now be dead.** Each was instrumented over the 85 documents, reading the text model
and listing, accepting and rejecting the tracked changes, both ways:

| code | untyped | typed | still reached by |
|---|---:|---:|---|
| `Extractor._typed_block` (a wildcard `w:p`, `w:tbl` or `w:sdt` in a block list) | 0 | 0 | no parse this package makes: in a branch `w:p` is global, `w:tbl` a root and `w:sdt` bound by its container; a tree built in code or parsed with no branch hook |
| `Extractor._typed_run_level` (a wildcard run-level item) | 0 | 0 | a branch's `w:moveFrom` and `w:moveTo`, which are neither global nor roots (no corpus document has one); `w:r`, `w:ins` and `w:del` are global, `w:hyperlink`, `w:fldSimple`, `w:smartTag` roots, `w:sdt` bound by its container |
| `Extractor._island` | 531 | 28 | VML text boxes |
| `revisions.islands.Island` | 566 | 48 | VML text boxes: every wildcard `w:txbxContent` left in the corpus, 514, is `v:textbox`'s, under `v:rect` (412) or `v:shape` (102) below a `Pict` |

Each keeps its code, and its docstring now says where it is still reached. The islands retire with
Phase C.

### 11.5 The survey and the census

`scripts/mce_branches.py`, section 5's survey. The corpus is built from commits by `git archive`
into `out/cr012-b-corpus/` (git-ignored), its `COMMITS.txt` recording them:

- docx4j-python `f1bb1d9`: `samples/`, `tests/fixtures/` (164 files);
- docx4j-core-ts `df5b4af`: `test/fixtures/` (252);
- docx4j `901a922a3`: `docx4j-core-tests/src/test/resources` (298).

That is 422 distinct documents and 2,746 WordprocessingML parts. `docx4j-core/src/test/resources`,
which the brief named, does not exist at `901a922a3`: `docx4j-core` has no `src/test`.

Each part is loaded twice: **untyped**, Phase A2's parser (the hook asked about `w:sdt`, `w:tr` and
`w:tc` only, no root class), and **typed**. The untyped pass gives the census taken at `f1bb1d9`
before any change, count for count.

**What a kept branch holds** (its direct children, over the corpus):

| branch child | untyped | typed |
|---|---:|---:|
| `mc:Choice` / `w:drawing` | wildcard 521 | `Drawing` 521 |
| `mc:Choice` / `w:pict` | wildcard 1 | `Pict` 1 |
| `mc:Choice` / `w:t` | wildcard 24 | `RT` 24 |
| `mc:Fallback` / `w:pict` | wildcard 491 | `Pict` 491 |
| `mc:Fallback` / `w:drawing` | wildcard 4 | `Drawing` 4 |
| `mc:Fallback` / `w:t` | wildcard 11 | `RT` 11 |
| `mc:Choice` / `w:r` (global) | `R` 47 | `R` 51 |
| `mc:Fallback` / `w:r` | `R` 16 | `R` 20 |
| `mc:Choice` / `w:p`, `w:tr`, `w:tc` | `P` 54, `Tr` 20, `Tc` 20 | the same |
| `mc:Fallback` / `w:p`, `w:tr`, `w:tc` | `P` 8, `Tr` 2, `Tc` 2 | the same |
| `mc:AlternateContent` inside a wildcard tree | 4 | 0 (typed, in a typed drawing; their branches' runs are the 4 more `R` each side) |

**How deep the typing reaches**, every node below a branch:

| | untyped | typed |
|---|---:|---:|
| in a Choice, wildcard | 17,104 | 19 (VML below the one `Pict` in a Choice) |
| in a Choice, typed | 959 | 21,238 |
| in a Fallback, wildcard | 5,981 | 5,383 (VML below a `Pict`) |
| in a Fallback, typed | 140 | 753 |

**The census on `samples/`**, pinned in `test_the_census_on_samples`, as section 5 expected: 3
`Drawing` in Choices, 3 `Pict` in Fallbacks, no wildcard `w:drawing` or `w:pict` left in a branch.
The Choices' 135 wildcard nodes are 185 typed objects, and nothing in a Choice is a wildcard; the
Fallbacks hold the 3 `Pict` and 28 wildcard VML nodes.

**The gate.**

- **Canonical form:** no part's typed save differs from its source where its untyped save does
  not, and no part's two saves differ canonically. 2,698 parts are identical to the source both
  ways. The other 48 are the same both ways, as before this phase: 24 Strict OOXML parts, whose
  skipped content is a known gap (CR-002 section 12.3 item 7); 23 where the schema's order is
  imposed (Word Online's `settings.xml` in `HelloWordOnline.docx`, `w:object`'s `v:shape` and
  `o:OLEObject` in the OLE fixtures); and one hex colour's case (`pandoc-code.docx`'s styles).
- **Skipped reports:** every one is the same both ways, part for part (24 parts have one, all
  Strict).
- **Bytes:** 25 parts differ, each holding a branch, as section 3 found: 19 in attribute order
  only (equal in C14N), 6 in boolean spelling (`wps:bodyPr anchorCtr="false"` and `wps:cNvSpPr
  txBox="true"` written `0` and `1`, 413 attributes, 403 of them `tc-p3-transforms.docx`'s). The
  script fails if a part with no branch differs.

### 11.6 The numbers

- **The suite:**
  - the whole suite: 5,133 passed, 8 skipped, in 3 minutes 8 seconds, with no timing flake (10.5's
    4,855, then 10.10's 8 and 10.11's 10, then this phase's 260);
  - the fast suite (`-m "not slow"`): 3,853 passed, 8 skipped, 1,280 deselected (3,593, 8 and
    1,280 before);
  - the 260 new tests are `tests/openpackaging/test_mce_typed.py`'s; nothing else moved but 11.8.
- **Unchanged:**
  - `scripts/roundtrip.py`: 50 of 50 canonically identical, 0 differences, 0 skipped, 34 benign
    respellings;
  - CR-011's ratchet: 1,291 pass, 0 fail, 293 pending of 1,584, goldens 12 of 12,
    `conformance.tsv` unchanged;
  - the agreement test: 638 passed;
  - CR-010's Word-save comparisons (`test_word_saves.py`, `test_compare_word_resolution.py`): 68
    passed; `tests/revisions`: 547 passed, its two known-gap pins unchanged;
  - the parity goldens: 277 passed, `PINS` empty;
  - `scripts/threads.py`: no configuration lost or altered a document;
  - `generate.sh --check`: byte identical.
- **`scripts/bench_marshal.py`** (`--runs 7`; `samples/` three runs, `tc-p3-transforms.docx` five,
  this machine):

  | | before | after |
  |---|---:|---:|
  | `Symbols.docx` parse | 204 to 211 ms | 195 to 201 ms |
  | `Symbols.docx` marshal | 254 to 264 ms | 248 to 252 ms |
  | the 14 main parts, parse | 249 to 257 ms | 238 to 244 ms |
  | the 14 main parts, marshal | 305 to 319 ms | 298 to 302 ms |
  | `tc-p3-transforms.docx` parse | 266 to 279 ms | 321 to 339 ms |
  | `tc-p3-transforms.docx` marshal | 168 to 182 ms | 234 to 240 ms |

  One of the five `tc-p3` parses before was 522 ms, an outlier left out. On `samples/` the change is
  within the band. On `tc-p3-transforms.docx` (403 branches, a 1.9 MB main part) parse is up about
  21% and marshal about 35%, near the prototype's upper bound (section 3: +24% and +40%):
  - the hook is not the cost: its 806 calls there take 0.4 ms;
  - the cost is the typing itself. The Choices' 13,345 wildcard nodes become 15,775 typed objects,
    each built with its parent pointer and written through the metadata, where a wildcard tree is
    built and written as plain nodes. That is Q1's trade, typed plain instances.
- **`import docx4j_py`:** 630 to 680 ms, against 632 to 672 before (seven runs each, `-X
  importtime`), 242 modules both times. The root table is strings, resolved on first use.
- **The acceptance artefacts:** `out/acceptance/` was regenerated and compared with a copy made
  before the first change (`out/acceptance-before-cr012-b/`). `scripts/compare_acceptance.py`
  found 13 artefacts and 175 entries, 0 differing beyond the wall-clock dates; entry by entry, 167
  byte for byte and 8 in the dates only.

### 11.7 Departures, with why

1. **The generated table has a second half, `ROOT_ELEMENT_BINDINGS`.** Section 4.2 item 3 names
   `ROOT_ELEMENTS` alone, and item 1 asks the serialiser to name every class step (a) can give. The
   runtime cannot list them without building every class's metadata. So the generator lists them,
   from the same metadata, and checks them a function. They are strings, like `ROOT_ELEMENTS`.
2. **"This model declares globally" is read from the schema copy** in `derive_names.py`, which is
   stdlib only: a top-level `xsd:element` of `wml/wml.xsd`. `generate_el.py` then holds the table
   to the model (`find_type`-equivalent: no root may be a global element class).
3. **The cross-check found `w:sdtPr` annotated** (11.2). The brief allowed the two sources to
   disagree on `w:sdtPr` only; they do not disagree at all.
4. **The (a)/(b) exception is wider than `w:hyperlink`** (11.3): the 20 run content names and
   `w:t` in a math run.
5. **The walks' fix is in `child.iter_children`**, the one enumeration every walk shares, not in
   `traversal` alone. So `walk`, `walk_all`, `iter_nodes`, `find`, `text_of` and the text model's
   address walk all name a branch's typed child as the serialiser writes it. It also mends a gap
   older than this phase: a branch's typed `w:r` and `w:p`, global since CR-001, were named
   `any_element` and not read by `text_of` (11.4).
6. **`Body.tables` reads through a block-level element**, 9.11's item, which Phase B's typing alone
   did not close.
7. **The python-docx facade keeps a drawing straight in a branch out of `inline_shapes`**, to keep
   python-docx's `//w:p/w:r/w:drawing/wp:inline` (11.4).
8. **No fork item 28.** The fork knows nothing of markup compatibility, so "a lax wildcard that is
   not a branch's" is any lax wildcard to it. Its
   `test_serializer_element_names.py::test_written_and_read_back_through_the_wildcard_class_hook`
   asks the hook in, and names the hook's class in, the lax wildcard of a plain `x:branch` element.
   Item 27 pins a typed child holding a branch of its own. Both of section 4.2 item 5's candidates
   are pinned, so nothing was added; the fork's suite was not run, since the fork did not change.
9. **The record's hook** (`mc_block_level._record_hook`) tries (a) with the parse's own ancestors,
   then (a) with the record level's container, then (b), where it took `branch_child_class` and
   the container. With (b) inside `branch_child_class`, the old order would have given a root class
   before the level's container could bind the name.
10. **"No other name is typed" is narrower than its examples.** `w:ins`, `w:del` and `w:delText` are
    global elements of this model (`RunIns`, `RunDel`, `DelText`), so a branch's were typed by the
    lookup by name before this phase. What stays a wildcard tree is a branch's `w:instrText`,
    `w:delInstrText`, `w:moveFrom`, `w:moveTo` and `w:customXml` (`w:moveFrom` and `w:customXml`
    tested, and none of them in `wildcard_names()`).
11. **`tc-p3-transforms.docx`** is docx4j's (`docx4j-core-tests/src/test/resources/theme-colors/`),
    not docx4j-core-ts's, and is in the corpus from there.

### 11.8 Pre-existing tests whose expectation changed, and why

Both are in `tests/test_builders_phase_a.py`, over `samples/DrawingML_GraphicData_wps.docx`, as
section 3 foresaw.

- `test_walk_all_reaches_the_wildcards_inside_an_alternate_content`.
  - **Then:** the `wps:wsp` in the `mc:Choice` was a wildcard, handed to the wildcard visitor, and
    wildcard nodes outnumbered typed ones.
  - **Now:** the typed visitor gets a `Wsp` named `{wps}wsp`; the wildcard visitor gets the
    Fallback's VML (`v:rect`), all `AnyElement`; no `wordprocessingShape` name is a wildcard. The
    count comparison is gone (80 typed, 26 wildcard).
- `test_walk_all_sees_the_attributes_on_a_wildcard`.
  - **Then:** it read `distL` and `w14:anchorId` from the Choice's wildcard `wp:anchor`.
  - **Now:** that is a typed `Anchor` (`dist_l` 114300, also asserted), and the wildcard whose
    attributes it reads is the Fallback's `v:rect` (`o:spid`, `w14:anchorId`).

Two kinds of test failed along the way and pass with their expectations unchanged:

- `tests/openpackaging/test_mce.py`'s three `text_of` tests (section 3's), with the hook alone: a
  typed `RT` was named `any_element`. The walks' fix (11.7 item 5) mends them.
- `tests/text/test_text_markdown.py::test_a_chartex_chart_from_its_choice_and_no_picture` and
  `tests/text/test_texts_phase2.py::test_a_chartex_chart_in_its_choice`: `waterfall.docx`'s
  `cx:chart` in its Choice, typed now inside a typed drawing, was met by the text model's
  `_graphic_kind` only as a wildcard before, so its typed path did not know chartex's class. It
  reads a typed item by its element name now.

Nothing else moved.

### 11.9 Found, and not done

- **The `text_of` question** (11.4): decided by the coordinator, kept. What Office JS's
  `paragraph.text` gives for a branch is the user's to measure, if wanted.
- **CLAUDE.md's MCE paragraph** says `mce_preprocess` is the only way to a typed `Drawing` in a
  choice; that is no longer so. The coordinator's, as section 4.2 item 6 says.
- **A fork release.** Stages 9 to 11 are local, as before; this phase adds no fork code, so it
  needs no more than Phases A and A2 did.
- **docx4j's agreement run** (it calls `revisions.accept_all`): rerun by docx4j-25 on 2026-10-22
  against `6a7a03c` (this phase, with 10.12's port), docx4j's side at `84b202c5a`'s revisions
  code: **0 lines differ in all 22 comparisons** (the eleven fixtures, accepted and rejected), and
  0 in the 2026-10-21 set's `review-for-word-word365-rejected` too, 24 of 24 as run. Section 5's
  gate holds.
- **The content API's inline pictures** (`Body.inline_pictures`, `Paragraph.inline_pictures`) pass
  a picture in a run-level branch by, before and after, as `Paragraph.text` passes its text by
  (Q4). The text model reads it. Not a change, and Q4's.
- **A drawing in a typed text box in a branch.** Its parent is a run, so the facade's
  `inline_shapes` would list it (only Word's branch, `mce="resolve"`), where python-docx lists it
  in every branch. No corpus document has one; before this phase the whole box was a wildcard tree.
- **A `word`-mode record inside a text box in a run-level branch** (9.11's item): a DrawingML box
  there is typed now, so the text model reads it directly rather than through a typed copy, and the
  record's warning may now be raised there. Not probed; no probe has the shape.
- **The corpus's 48 parts not canonically identical to their source**, the same before and after
  (11.5): Strict OOXML, schema order in `settings.xml` and `w:object`, and a colour's case. Outside
  this phase.

### 11.10 The Word check (passed 2026-10-22)

Four files are in `out/cr012-b-word-check/`, with a `README.txt` saying, for each, the source,
what was done and what Word should show. They are also copied to
`~/vbShares/Office 2016/__tmp/python-cr012b/`. Each was loaded, its main part unmarshalled so that
this engine writes the typed branches back, and saved. Each reads back with
`LoadOptions(strict=True)`, nothing skipped, its main part canonically identical to its source and
no `xsi:type` written.

- `cr012b-1-wps-shape.docx`: `samples/DrawingML_GraphicData_wps.docx`, a wps shape in a run-level
  Choice and its VML in the Fallback. It differs from an untyped save in attribute order only.
- `cr012b-2-tracked-text-box.docx`: docx4j-core-ts's `cr043-tracked.docx`. Its case C22 is a text
  box in a run-level Choice, now typed down to its `RunIns` and `RunDel`, holding a tracked deletion
  and insertion, with their twins in the VML Fallback. Attribute order only.
- `cr012b-3-txbx-run.docx`: `tests/fixtures/mc-fun/mc-txbx-run.docx`, a text box in a run-level
  element. The same bytes as an untyped save.
- `cr012b-4-theme-colour-shapes.docx`: `tc-p3-transforms.docx`, 403 rectangles in run-level
  elements. Its 403 typed `wps:bodyPr` are written `anchorCtr="0"` where the source has `false`.

Expected: each opens with no repair prompt and looks as its source does, C22's tracked changes
shown.

**Passed on 2026-10-22.** The user: "cr012b okay except cr012b-2 textbox tracked change content
not as expected", then, after a second look, "cr012b-2 fine in All Markup".
- **The first look** was in Simple Markup, where Word shows C22's text box as "a word to delete:.
  Inserted text goes at the end. C22 inserted.": the deletion hidden, the insertion plain.
- **The save was checked** against the source. C22's element differs only in where the `a` prefix is
  declared and in the spelling of one newline in a VML attribute (`&#10;` for `&#xA;`). The tracked
  deletion and insertion are in both branches, as in the source.
- **The README now names the view** (All Markup), and the unmodified source is staged beside the save
  as `cr012b-2-SOURCE-cr043-tracked.docx` for comparison.
- **In All Markup** C22's tracked changes show as in the source.

## 12. Phase C implementation notes (2026-10-22)

Phase C puts VML in the model, at the user's go ("do phase c"), on section 4.3 as revised before it
started. docx4j's five VML schemas, at the schema copy's commit, are in `schemas/vml/` and reached
from the root schema; their name tables are derived from docx4j's classes and cross-checked against
the copy; the regeneration adds five namespace packages and touches no other generated module. So
`CT_Picture`'s and `CT_Object`'s lax wildcards, and `CT_ShapeDefaults`', type VML by the global
elements now in the context, wherever it stands: below every `w:pict` and `w:object`, in a branch or
not, a header's watermark, a numbering picture bullet and `settings.xml`'s `o:shapedefaults`. A
Fallback's VML text box is a typed `TxbxContent`. VML's booleans are a string enumeration and keep
their spelling by the schema's own types; the two VML values the converter would have respelled,
`o:spt` and `o:gfxdata`, keep theirs through a fork option (stage 12). The text model reads typed VML
in place, `insert_ooxml` rewrites a typed `r:id`, and no corpus document makes a typed island any
more.

**Commits.**

- docx4j-python:
  - `6897210`: the schema copy, `schemas/vml/` and the root schema's imports, `schemas/PATCHES.md`;
  - `3dbd31a`: the five name tables, `codegen/derive_names.py`'s cross-check, the substitutions,
    `codegen/names/README.md`;
  - `58096b6`: the regeneration;
  - `a1a6abc`: `runtime.parser_config` sets the fork's `lexical_types`;
  - `e658512`: the readers (the text model's `_vml_shape` and `_object`, `insert_ooxml`'s
    `_rel_fields`), the docstrings that said VML was a wildcard tree, and the pre-existing tests
    whose expectation changed (12.12);
  - `b107e23` and `0800ad8`: `tests/openpackaging/test_vml.py`;
  - `ef86b1d`: `scripts/vml_census.py`, and `scripts/mce_branches.py readers --holding` with four
    more readers;
  - and the commit of this section, the CHANGELOG, CR-002 section 5.6's note, `tests/README.md`
    and `docs/UPSTREAM.md` report 25.
- The fork (`~/git/docx4j-xsdata`, branch `docx4j`, local only): `d084a267`, stage 12 item 28,
  `ParserConfig(lexical_types=)` (12.6).

### 12.1 What landed

| piece | where | what |
|---|---|---|
| the schema copy | `schemas/vml/` (five files), `schemas/docx4j_python__ROOT.xsd`, `schemas/PATCHES.md` | docx4j's `vml-main`, `vml-officedrawing`, `vml-wordprocessingdrawing`, `vml-spreadsheetDrawing` and `vml-presentationDrawing` at `2853b7280`, byte for byte; `vml__ROOT.xsd` left out; no patch (12.2, 12.4) |
| the names | `codegen/derive_names.py`, `codegen/names/vml*.json` (five), `codegen/names/README.md` | 92 types and 49 elements from docx4j's XJC build; `cross_check_vml` against the copy (12.3) |
| the modules | `.xsdata.phase-b.xml` | a `Substitution` per namespace: `vml.main`, `vml.officedrawing`, `vml.wordprocessing_drawing`, `vml.spreadsheet_drawing`, `vml.presentation_drawing` |
| the regeneration | `docx4j_py/vml/` (five packages, each with its `el`), `docx4j_py/__init__.py`, `docx4j_py/el_index.py`, `codegen/el_tables/vml_*.json` (five) | 167 classes; `EL_MODULES` and `UNDERSTOOD` gain the five namespaces; no other generated module changed |
| the fork | `docx4j_xsdata/models/datatype.py`, `parsers/config.py`, `parsers/utils.py`, `models/elements.py`, `serializers/mixins.py`; `tests/fork/test_lexical_types.py` | stage 12 item 28: a float or bytes value that keeps the spelling it was read with (12.6) |
| the runtime | `docx4j_py/runtime.py` | `LEXICAL_TYPES` (float, bytes) in every parser config |
| the readers | `docx4j_py/model/text/extractor.py`, `docx4j_py/model/content/ooxml.py` | `_vml_shape` and `_object` read a typed `v:shape` (its `v:textpath` string, `v:imagedata` `r:id`, `alt`) and `o:OLEObject` (ProgID, `r:id`), a wildcard tree still by name; `_rel_fields`: every relationships-namespace attribute field of a typed node (12.8) |
| the docstrings | `model/revisions/islands.py`, `model/revisions/resolver.py`, the text model's, `openpackaging/parts/default_xml_part.py` (`VMLPart`), `model/content/ooxml.py` | where an island is still reached; why a vmlDrawing part stays a tree (12.9) |
| tests | `tests/openpackaging/test_vml.py` (22); `test_mce.py` (+1), `test_mce_typed.py`, `test_builders_phase_a.py` (+1), `revisions/test_tracked_changes.py` | 12.10, 12.12 |
| the survey | `scripts/vml_census.py`; `scripts/mce_branches.py` | `run` and `compare`; `readers --holding` (`mc`, `vml` or `either`), and the fonts and styles in use and Accept All and Reject All as readers (12.7, 12.8) |
| documents | `CHANGELOG.md`, CR-002 section 5.6, `tests/README.md`, `docs/UPSTREAM.md` report 25; the fork's `docs/fork/CHANGES.md` stage 12 | |

### 12.2 The copy and its checks

- **From commits only.** Each file was taken with `git -C ../docx4j show 2853b7280:xsd/vml/<file>`,
  never from docx4j's working tree. Its `git hash-object` here equals its blob at `2853b7280` and at
  `acd0defc9` (docx4j CR-048): `git log acd0defc9..2853b7280 -- xsd/vml` is empty, and `acd0defc9`
  is an ancestor of `2853b7280`. The files are docx4j's in every byte, BOM and line endings
  included (CRLF; `vml-main.xsd`'s CR-048 lines LF), and the repository has no `.gitattributes`.
- **Reached from the root.** `docx4j_python__ROOT.xsd` imports all five. `vml-main.xsd`'s own
  imports, `../wml/wml.xsd`, `../wml/w14_word_2010_wordml.xsd` and
  `../shared/shared-relationshipReference.xsd`, were already in the copy. `wml/wml.xsd` imports no
  VML (`CT_Picture`'s wildcards name the namespaces without an import), so no file of the closure
  changed.
- **Left out:** `vml/vml__ROOT.xsd` (a vmlDrawing part's root, `<xml>` in no namespace,
  SpreadsheetML's and PresentationML's) and `vml/README.txt`.
- **Markers:** `grep -ro 'docx4j-python' schemas --include='*.xsd'` still finds 26; the VML files
  carry none. lxml's `XMLSchema` over the root reports the one error it reported before (the
  chartex QName with a trailing space, `PATCHES.md`'s verification).

### 12.3 Names and modules

**The source** is the XJC build Phase B used (11.2): `org.docx4j.vml`'s newest file was written at
2026-10-20 14:11:44 +11:00, the same time as `org.docx4j.wml`'s. The build carries docx4j CR-048
(`CTRect.java` has `anchorId` and `gfxdata`), so it postdates `acd0defc9`.

**The derivation needed no change.** `derive_names.py` takes every namespace of the schema copy
that a docx4j package declares, so with `schemas/vml/` in place it wrote five more tables:

| table | java package | types | elements |
|---|---|---:|---:|
| `vml.json` | `org.docx4j.vml` | 36 | 21 |
| `vml_officedrawing.json` | `org.docx4j.vml.officedrawing` | 42 | 18 |
| `vml_wordprocessing_drawing.json` | `org.docx4j.vml.wordprocessingDrawing` | 9 | 6 |
| `vml_spreadsheet_drawing.json` | `org.docx4j.vml.spreadsheetDrawing` | 3 | 2 |
| `vml_presentation_drawing.json` | `org.docx4j.vml.presentationDrawing` | 2 | 2 |

The other 56 tables and `root_elements.json` are rewritten byte for byte, before and after the
regeneration (every docx4j VML root is a global element of the VML schemas, so `root_elements.json`'s
`other_vocabularies` gains nothing). The run's dropped `@XmlElementDecl`s fall from 264 to 215: the
49 in the VML namespaces are routed to their tables.

**The cross-check** (`cross_check_vml`, made a check that stops the run, as Phase B's was). For each
VML namespace, `types` must name exactly the schema's complex types and nothing that is not a type of
the schema, and `elements` every global element of the schema and nothing but a local element of
simple type besides. It holds for all five. It lists three enumerations JAXB makes no class for
(`ST_TrueFalseBlank` in three schemas, `o:ST_Angle`, `o:ST_ScreenSize`: an enumeration value that is
no Java identifier, `""`, `"30"`, `"544,376"`), which the generator names itself (`StTrueFalseBlank`,
as `StOnOff` elsewhere), and `x:CF`, a local element of simple type docx4j's ObjectFactory declares
scoped to `CTClientData`. The cross-check ties the tables to `2853b7280`.

**The regeneration.**
- 63 namespace modules to 68, 1,802 classes to 1,969, 2,063 element names over 60 namespaces to
  2,189 over 65, 283 `.py` modules to 294.
- 282 `ClassNames` warnings to 330. The 48 new ones are all of one kind: a global VML element whose
  named type owns the docx4j name "is left alone" (`v:rect` of `CT_Rect`, the four `o:` stroke
  children of `CT_StrokeChild`, the four `w10:` borders of `CT_Border`), CR-001 section 13.3's
  case: the element's class keeps the generator's name (`Rect`, a subclass of `CTRect`).
- `docx4j_py/wml/` and every other module are unchanged. The import-order manifest gains the five
  packages, and `el_index.py` their `EL_MODULES` entries and `UNDERSTOOD`'s `v`, `o`, `w10`, `xvml`
  and `pvml`.
- **The `el` tables need no override:** no collision, one keyword rename (`o:complex`,
  `el.complex_`). `o:fill` and `v:fill`, `o:shadow` and `v:shadow` are in separate modules.
- `generate.sh --check`: byte identical.

### 12.4 Patches

**None.**
- The generator took every construct of the five schemas (attribute groups, VML's `ST_` types,
  `CT_Textbox`'s choice of `w:txbxContent` and a `##local` skip wildcard), with no fork change.
- Nothing is newly skipped over the survey (12.7), so no attribute or element Word writes is
  missing from the schemas: CR-048's `w14:anchorId` and `o:gfxdata` are there.
- The booleans keep their spelling by the schema's own types (12.5), and the floats and base64 by a
  fork option rather than a patch, so the types stay docx4j's: `o:spt` an `xsd:float`, as the
  objects-ts session asked docx4j to keep it (docx4j CR-048, D3).

### 12.5 Booleans

**VML declares no `xsd:boolean`.** Its booleans are `ST_TrueFalse` (`t`, `f`, `true`, `false`) and
`ST_TrueFalseBlank` (those and `""`), string enumerations in `vml-main.xsd`, `vml-officedrawing.xsd`
and `vml-spreadsheetDrawing.xsd`. The generator makes them enums (`STTrueFalse`,
`StTrueFalseBlank`), whose values are written as read, so the model's numeric booleans (CR-001 Q5)
never meet them.

- **No VML or `o:` attribute's spelling changes on a save.** Over the corpus there are 778
  boolean-valued VML attributes, 245 `t` and 533 `f` (no `true` or `false`), on `stroked` (464),
  `gradientshapeok` (84), `filled` (43), `o:preferrelative`, `o:extrusionok` and `aspectratio`
  (33 each) and ten others; none changed.
- A value outside an enumeration would be kept as its string, with the converter's warning; the
  survey met none.
- `test_the_boolean_spellings_are_kept` pins all four spellings.

**What did change spelling**, until stage 12, was no boolean: with VML typed and no option, 97
`o:spt` on `v:shapetype` (`202` written `202.0`, `75` as `75.0`, `136` as `136.0`, an `xsd:float`)
and 54 `o:gfxdata` (48 on `v:shape`, 4 on `v:oval`, 2 on `v:rect`; `xsd:base64Binary`, which Word
wraps at 76 characters with `&#xA;` and the converter wrote on one line), in 50 parts of 41
documents. Those were the only new canonical differences typing VML made. docx4j writes both rewrites
(its CR-048 D3), and Word read them in docx4j's checks; the gate here asks for none, so the fork keeps
the source's spelling (12.6). No integer or decimal VML value changed.

### 12.6 The fork's item

**Stage 12, item 28 (`d084a267`): `ParserConfig(lexical_types=)`.** `docs/fork/CHANGES.md` stage 12
has the detail, and `docs/UPSTREAM.md` report 25.

- **The option:** `lexical_types`, of `float` and `bytes`, off (empty) by default.
- **What it binds:** where the converter would write a value of one of them in another spelling than
  its source's, the value bound is `XmlLexicalFloat` (a slotted `float`) or `XmlLexicalBytes`,
  carrying the source text.
  - Either is equal to and hashed as the plain value, and pickled and copied with its spelling.
  - The serialiser writes the text.
  - Arithmetic, or a value set in code, gives a plain value, written in the converter's spelling.
  - A value spelled as the converter would spell it is bound plain.
- **The cost:** the converter plan gains a fourth entry, so a value of any other type pays one tuple
  index. Interleaved over 15 runs, `Symbols.docx` parses in 205 ms median with the option on and
  off.
- **docx4j-python** sets `frozenset({float, bytes})` (`runtime.LEXICAL_TYPES`) for every part, not
  VML alone. Its one other effect over the corpus is an improvement:
  `docx4j-core-ts/test/fixtures/onoff/pandoc-code.docx`'s `styles.xml`, whose 20 lower-case
  `w:color` values (`ST_HexColor`, hexBinary) were written in upper case, is canonically identical
  to its source for the first time.
- **The fork's suite:** 1,668 passed and 1 skipped, against 1,656 and 1 before. Twelve new tests;
  ten fail without the item, and the other two pin the behaviour with the option off.

It joins stages 9 to 11 in needing a fork release before a docx4j-python release.

### 12.7 The survey and the census

`scripts/vml_census.py`, section 4.3 item 5's survey, over Phase B's corpus. `out/cr012-b-corpus/`
was built from commits by `scripts/mce_branches.py corpus`, its `COMMITS.txt` recording
docx4j-python `f1bb1d9`, docx4j-core-ts `df5b4af` and docx4j `901a922a3`. It was kept, since
neither this repository's `samples/` and `tests/fixtures/` nor core-ts's `test/fixtures/` changed
after those commits. It is 422 distinct documents and 2,746 WordprocessingML parts.

VML is typed by the model, not by a hook, so the survey compares two runs.
- **Before:** `out/cr012-c-survey/before`, at `f062c25`, before any change.
- **After:** `out/cr012-c-survey/after`, at the last commit.
- Each run loads every part, records its census, skipped report and save, and compares the save with
  the source in `scripts/canon.py`'s form; `compare` sets the two runs' saves side by side. A run
  takes 22 seconds.

**The census**, every VML element by vocabulary, kind and place (an element in an `mc:Choice` or
`mc:Fallback` counted on that side):

| | before | after |
|---|---:|---:|
| `v:` in a Choice | wildcard 18 | typed 18 |
| `v:` in a Fallback | wildcard 1,106 | typed 1,106 |
| `v:` outside any branch | wildcard 975 | typed 975 |
| `o:` in a Choice | wildcard 1 | typed 1 |
| `o:` outside any branch | wildcard 1,046 | typed 1,041 |
| `w10:` in a Fallback | wildcard 50 | typed 50 |
| `w10:` outside any branch | wildcard 2 | typed 2 |
| `x:`, `pvml:` | none | none |

- The five `o:` fewer are `o:OLEObject`'s simple-typed children (`o:FieldCodes` 1, `o:LinkType` 2,
  `o:LockedField` 2), field values of `CTOLEObject` now, not nodes.
- **By part type:** the main document 2,052 (2,047 typed); settings 963 (`o:shapedefaults` 409,
  `o:shapelayout` and `o:idmap` 277 each); headers 144; numbering 38 (picture bullets); comments 1.
- **The holders:**
  - `w:pict`: 1 in a Choice, 494 in a Fallback, 43 outside;
  - `w:object`: 41, all outside;
  - `w:txbxContent` in a Fallback: wildcard 490 before, typed 490 after;
  - `w:txbxContent` outside a branch: 25 wildcard (VML) and 22 typed (DrawingML) before, 47 typed
    after.
- 333 parts hold VML. After, no VML element in any of them is a wildcard.
- **On `samples/`**, pinned in `test_the_census_on_samples`:
  - settings: 13 `o:shapedefaults`, 12 `o:shapelayout`, 12 `o:idmap`;
  - in the three Fallbacks: 2 `v:oval`, 1 `v:rect`, 1 `v:textbox` and its `w:txbxContent`;
  - all typed.

**The gate.**

- **Skipped:** no item is newly skipped. The same 24 parts have a report, all Strict OOXML (CR-002
  section 12.3 item 7).
- **Canonical form:** no part gained a difference. 2,698 parts were identical to their source
  before, and 2,699 are after; the one more is `pandoc-code.docx`'s styles (12.6). The other 47 are
  as before: 24 Strict parts, and 23 where the schema's order is imposed (three `settings.xml`, Word
  Online's among them; eighteen main parts, among them the `w:object`s of the OLE fixtures and the
  extraction probes, 12.13; two `styles.xml`), as Phase B listed them (11.5).
- **Bytes:** 58 parts differ from the before run's save.
  - 57 in attribute order alone, equal in C14N: a typed element writes the schema's order, where a
    wildcard kept the source's. The elements: `v:rect` 412, `v:shape` 163, `w:p` 65 (a VML text
    box's paragraphs, typed now: `w14:paraId` after `w:rsid*`, as Phase B found), `v:shapetype` 48,
    `v:imagedata` 47, `v:path` 33, `v:textpath` 8, `v:oval` 4, and one each of `o:shapedefaults`
    and `o:idmap`, in core-ts's `NumberingIndents.docx` settings, which writes `spidmax` before
    `v:ext`. By part type: 43 main documents, 10 headers, 2 numbering parts, 1 comments part, 1
    settings part.
  - 1, `pandoc-code.docx`'s styles, in 20 `w:color` values kept in lower case.
  - The other 2,688 saves are the same bytes as before: every `settings.xml` but one among them,
    since the typed `o:shapedefaults` writes `v:ext` and `spidmax` in the order Word does.

### 12.8 The readers' audit

Every reader of text or structure was run over every distinct corpus document whose WordprocessingML
parts hold `mc:AlternateContent` or declare a VML namespace, 384 of them (`readers --holding
either`), and the answers diffed.
- **The runs:** at `f062c25` before any change, in a worktree with the fork's new option off
  (`out/cr012-c-survey/readers-either-before.json`), and at the last commit. Its answers on Phase
  B's 85 documents and 17 readers equal those of a run made in the checkout before the fork
  changed.
- **The readers:** Phase B's 17 (11.4), and four more that a typed VML text box can move:
  - `MainDocumentPart.fonts_in_use()` and `styles_in_use()`;
  - every story's changes accepted and rejected (`revisions.accept_all` and `reject_all`): what is
    left, the text after, and each part's canonical form.

**For the coordinator's decision: none.**

**The one answer that changed:**

1. **`styles_in_use()`, 1 document,** `docx4j-core-tests/.../vml/textbox.docx`: `["Normal"]` before,
   `["DefaultParagraphFont", "Normal"]` after.
   - The document's VML text box (no `mc:AlternateContent`) holds a run with no `w:rStyle`, which
     the walk (`iter_nodes`) reaches now that the box is typed, so the default character style counts
     as used.
   - **Clearly right by docx4j's contract.** docx4j's `FontAndStyleFinder` walks with
     `TraversalUtil`, which descends `Pict`, `CTShape` and `CTTextbox` to the `w:txbxContent`
     (`TraversalUtil.getChildrenImpl` at `2853b7280`), and counts a run with no `w:rStyle` as the
     default character style's.
   - Kept. `fonts_in_use()` did not move: the box's run names no font.

**Nothing else changed** on the 384 documents: the content API's paragraphs, tables, inline
pictures, addresses, `outline()`, `describe()`, `get_text`, `find` and `search`; the text model,
plain text, Markdown and `to_markdown`; the tracked changes listed; Accept All and Reject All, their
text and every part's canonical form; the facade's paragraphs, tables and inline shapes. The same
four Strict documents fail to load both ways.
- **CR-010's twins:** the VML text box of `cr043-tracked.docx`'s C22, the Fallback twin of the
  Choice's (section 9's twins), is typed down to its `RunIns` and `RunDel` and is listed and resolved
  as before (`test_the_tracked_vml_text_box_is_typed_down_to_its_changes`, and the audit's Accept
  All and Reject All forms).

**What made the answers stay.** With VML typed and no reader changed, the fast suite lost 14 tests
to the readers, and 19 of CR-011's conformance checks went from PASS to FAIL:
- legacy WordArt, a VML picture's description and an `o:OLEObject`'s ProgID, which the text model
  read from a wildcard `v:shape`, `v:textpath`, `v:imagedata` and `o:OLEObject` by name;
- the embedded objects a ProgID leads to.

`_vml_shape` and `_object` now read the typed classes (`CTShape.content`, `CTTextPath.string`,
`CTImageData`'s `r:id`, `CTOLEObject.prog_id` and `id`), and a wildcard tree still by name.
`insert_ooxml`'s rewrite read typed nodes' relationship ids by a list of field names, among them
`id`. Typed `v:imagedata`'s `r:id` is a field of another name, beside VML's own `id`, so it reads
every relationships-namespace attribute field by the class's metadata now (`_rel_fields`), as the
revisions engine's walk does (`trees._rel_fields`).
`test_insert_ooxml_rewrites_a_typed_imagedatas_relationship_id` holds it.

### 12.9 The islands

Instrumented over the 422 documents, reading the text model and listing, accepting and rejecting
the tracked changes, at `f062c25` and after:

| code | before | after |
|---|---:|---:|
| `revisions.islands.Island` | 248 calls, 10 documents | 0 |
| `Extractor._island` | 203 calls, 9 documents | 0 |
| `Extractor._typed_block` | 0 | 0 |
| `Extractor._typed_run_level` | 0 | 0 |

**The islands do not retire:** something can still reach them.
- **Below a wildcard node every element is a wildcard,** so a text box there is a wildcard
  `w:txbxContent`. Such a node is:
  - an element no class binds (a VML-namespace element the schema does not declare,
    `<w:pict><v:unbound>`);
  - a branch's `w:moveFrom` or `w:moveTo`, neither global nor a docx4j root;
  - a tree built in code.
- **Who reaches one:**
  - for an unbound element, the text model, the listing and Accept All all go through islands;
  - for a branch's move, the text model reads it through `_typed_run_level`'s typed copy, and the
    listing does not walk it for text boxes.
- No corpus document has either.
- **Pinned:** `test_a_text_box_below_a_wildcard_node_is_still_a_wildcard_tree` and
  `test_a_text_box_below_a_wildcard_node_is_read_through_an_island`.
- **The docstrings** (`islands.py`, `_island`, the resolver's `_has_text`, the text model's) say so.

Section 4.3 item 6 said each retires "only if nothing reaches it". Their readers read the typed tree
in place wherever the parse types it, which is every corpus case. `_typed_block` and
`_typed_run_level` are not VML's, and stay as Phase B left them (11.4).

### 12.10 The numbers

- **The suite:**
  - the whole suite: 5,160 passed and 8 skipped in 3 minutes 21 seconds, against 5,136 passed and 8
    skipped before (3 minutes 13 seconds);
  - the fast suite (`-m "not slow"`): 3,880 passed, 8 skipped and 1,280 deselected, against 3,856,
    8 and 1,280 before;
  - the 24 new tests: `tests/openpackaging/test_vml.py`'s 22, `test_mce.py`'s
    `test_understood_names_vml_since_it_has_classes` and `test_builders_phase_a.py`'s
    `test_walk_all_reaches_both_branches_typed`;
  - `tests/test_performance.py` (slow) passes, with no rerun needed.
- **Unchanged:**
  - `scripts/roundtrip.py`: 50 of 50 canonically identical, 0 differences, 0 skipped, 34 benign
    respellings, exit 0;
  - CR-011's ratchet: 1,291 pass, 0 fail, 293 pending of 1,584, goldens 12 of 12;
    `conformance.tsv` unchanged (`test_conformance.py`, 24 passed);
  - CR-010's Word-save comparisons (`test_word_saves.py`, `test_compare_word_resolution.py`): 68
    passed, 0 lines;
  - `tests/revisions`: 550 passed, the known-gap pin in `test_alternate_content.py` unchanged;
  - `tests/openpackaging/test_threads_and_import.py` passes: the engine imports no model class at
    module level, and the text model imports the VML classes on first use (`_vml_classes`);
  - `scripts/threads.py`: no configuration lost or altered a document;
  - `generate.sh --check`: byte identical.
- **The acceptance artefacts:** `out/acceptance/` was regenerated and compared with the copy made
  before the first change (`out/acceptance-before-cr012-c/`). `scripts/compare_acceptance.py`
  found 13 artefacts and 175 entries, 0 differing beyond the wall-clock dates; entry by entry, 167
  byte for byte and 8 in the dates only.
- **Cost** (this machine; `import` in a fresh process, 7 runs; a cold import with an empty
  `PYTHONPYCACHEPREFIX`, 3 runs; `bench_marshal.py`, `samples/` three runs of 7 and four VML-heavy
  corpus documents three runs of 5):

  | | before | after |
  |---|---:|---:|
  | `import docx4j_py` | 629 to 676 ms (median 640) | 716 to 782 ms (median 744) |
  | cold import | 1,006 to 1,013 ms | 1,119 to 1,138 ms |
  | `warm_up()` | 46 to 51 ms | 50 to 63 ms |
  | `docx4j_py` modules after import and `warm_up()` | 100 | 106 |
  | `Symbols.docx` parse / marshal | 202 to 207 / 250 to 261 ms | 201 to 204 / 252 to 256 ms |
  | the 14 main parts of `samples/`, parse / marshal | 246 to 252 / 302 to 312 ms | 250 / 303 to 308 ms |
  | `tc-p3-transforms.docx` (806 VML elements in 403 Fallbacks) parse / marshal | 322 to 333 / 256 to 275 ms | 362 to 393 / 285 to 297 ms |
  | `extraction-probe.docx` (162) parse / marshal | 15 to 18 / 23 ms | 17 to 18 / 24 to 26 ms |
  | `loadAndSave.docx` (watermarks) parse / marshal | 7 / 8 to 9 ms | 7 to 8 / 9 ms |
  | `LegacyForms.docx` (VML only) parse / marshal | 2 / 2 ms | 2 / 2 ms |

  - **The import:** about 100 ms more, for the five modules. `-X importtime` gives
    `vml.officedrawing` with `vml.main` 61 ms, `vml.spreadsheet_drawing` 16 ms,
    `vml.wordprocessing_drawing` 6 ms and `vml.presentation_drawing` 1 ms. `vml.main` is 7,319
    lines: the eleven shape types each inline VML's attribute groups, some 60 fields a class. The
    cold import is 1.12 s against `test_performance.py`'s 2.0 s budget.
  - **A part with no VML** parses and marshals as before: the option's check is free (12.6).
  - **A VML-heavy part** pays for typing, as Phase B's Choices did: `tc-p3-transforms.docx` parses
    about 15% slower and marshals about 10% slower. That is less than Phase B's +21% and +35% on
    the same document, whose Fallbacks are each two VML elements.

### 12.11 Departures, with why

1. **A fork item for floats and base64, not booleans.** Section 4.3 item 5 named VML's booleans
   as the risk. They proved none (12.5); `o:spt` and `o:gfxdata` were the respellings. Section 4.3
   item 7 sends a defect VML exposes to the fork, and the phase's brief prefers a fork fix to a
   workaround here and to loosening a check:
   - a schema patch typing them as strings is a workaround here, and changes docx4j's types;
   - a `canon.py` category for them would loosen the check;
   - so the fork keeps a value's spelling (12.6).
2. **The option is on for every part,** not VML's alone: any float or bytes value keeps its
   spelling. Its one other effect in the corpus is `pandoc-code.docx`'s colours (12.6), a part
   canonically identical now that was not.
3. **The survey compares two runs,** where Phase B toggled its hook within one: VML is typed by the
   model's own classes, with no hook to switch off. The before run was made before any change.
4. **The readers' audit is wider than Phase B's.** It covers 384 documents, those declaring a VML
   namespace as well as those holding `mc:AlternateContent`, since VML outside a branch changed too.
   It has 21 readers, adding the fonts and styles in use and Accept All and Reject All. Its before
   run was made at `f062c25` in a worktree, after the fork had changed, with the fork's new option
   off; on Phase B's documents and readers it equals a run made before the fork changed.
5. **`cross_check_vml` stops the run,** as Phase B's root cross-check does, rather than being a one-off
   check.
6. **The islands stay** (12.9). Nothing in the corpus reaches them any more, but a text box below a
   wildcard node still can.
7. **The census counts nodes,** so the five simple-typed `o:` elements that are field values now
   drop out of the after column (12.7).
8. **Four Word-check files, not five.** `settings.xml`'s shape defaults are in all four, and both
   `w:shapeDefaults` and `w:hdrShapeDefaults` in file 2. No corpus document has more notable ones:
   273 settings parts have the standard `o:shapedefaults` and `o:shapelayout`, 132 an
   `o:shapedefaults` in `w:hdrShapeDefaults`, and none a fill, stroke or colour menu.

### 12.12 Pre-existing tests whose expectation changed, and why

- `tests/openpackaging/test_mce.py::test_understood_is_narrower_than_the_prefix_table`.
  - **Then:** `"v" not in UNDERSTOOD`, VML being the example of a prefix with no classes.
  - **Now:** `mv` (Mac VML, no schema) is the example.
  - `test_understood_names_vml_since_it_has_classes` holds the five VML prefixes. No branch
    selection reads `UNDERSTOOD` (Q8), so no reader moved.
- `tests/openpackaging/test_mce_typed.py::test_the_choice_is_typed_all_the_way_down_and_the_fallback_to_its_pict`,
  renamed `test_both_branches_are_typed_all_the_way_down`.
  - **Then:** the Fallback's `v:rect`, `v:textbox` and `w:txbxContent` were wildcards.
  - **Now:** they are `CTRect`, `CTTextbox` and `CTTxbxContent`, and no node in either branch is a
    wildcard.
- `tests/openpackaging/test_mce_typed.py::test_the_census_on_samples`.
  - **Then:** the Fallbacks' depth was 3 typed and 28 wildcard.
  - **Now:** 44 typed.
- `tests/test_builders_phase_a.py`'s walk_all tests over `samples/DrawingML_GraphicData_wps.docx`.
  The document holds no wildcard any more, so the wildcard visitor's tests read a fragment whose
  `w:pict` holds a VML-namespace element no class binds (`UNBOUND_VML`).
  - `test_walk_all_reaches_the_wildcards_inside_an_alternate_content`: the wildcard visitor gets the
    unbound element and its child.
  - `test_walk_all_reaches_both_branches_typed` (new): the sample's `Wsp` and `CTRect` go to the
    typed visitor, under their names, and no wildcard is met.
  - `test_walk_all_sees_the_attributes_on_a_wildcard`: `o:spid` and `w14:anchorId` are read from
    the unbound element's attributes and from the sample's typed `CTRect`'s fields.
  - `test_walk_all_reports_the_parent_of_a_wildcard`: over the fragment, two wildcards. Over the
    sample it would pass with none.
- `tests/revisions/test_tracked_changes.py::test_a_comment_across_paragraphs_in_a_text_box_is_judged_in_its_wildcard_tree`,
  renamed `test_a_comment_across_paragraphs_in_a_vml_text_box_is_judged_by_its_text`. Its
  expectation is unchanged; its premise was that the VML box is a wildcard tree at the flush.

Fourteen tests failed along the way and pass with their expectations unchanged:
- `tests/agent/test_outline.py::test_the_counts_follow_the_text_model` (170 paragraphs for 171);
- the plain-text and Markdown goldens of `extraction-probe` (`test_texts.py`,
  `test_text_markdown.py`);
- seven of `test_texts_phase2.py` (a VML picture's alt, objects' ProgIDs, legacy WordArt, embedded
  documents and their warnings);
- `test_conformance.py`'s two (19 checks PASS to FAIL).

All went with VML typed and the text model still reading it by name. `_vml_shape` and `_object`
mend them (12.8).

### 12.13 Found, and not done

- **`w:object`'s children are saved in the schema's order.** `CT_PictureBase` is a repeating
  sequence of a VML element and an `o:` element, and its two wildcard lists are written
  interleaved. So a `w:object`'s `v:shapetype`, `v:shape`, `o:OLEObject` is saved `v:shapetype`,
  `o:OLEObject`, `v:shape`.
  - It is older than this phase, the same with wildcards (11.5's 23 parts).
  - It is the one canonical difference among the Word-check files (file 3).
  - A marked schema patch making the sequence a choice, one compound list in document order, would
    end it. That is for the coordinator: it changes `CT_PictureBase`'s generated fields.
- **A tracked change in a branch's `w:moveFrom` or `w:moveTo`** is not listed. The move is a
  wildcard tree, which the listing does not walk; the text model reads it through a typed copy. No
  corpus document has one; this phase changed nothing there.
- **The import's 100 ms** (12.10) could come down only by importing VML on first need. That means
  the fork's context learning a namespace's module lazily, and is not attempted.
- **A vmlDrawing part (`VMLPart`) stays an lxml tree.** Its root's schema, `vml__ROOT.xsd`, is
  SpreadsheetML's and PresentationML's (CR-001 Phase D, CR-002 Phase C).
- **`x:` and `pvml:`** are in the model and in no corpus document's WordprocessingML; they are typed
  and untested beyond the regeneration.
- **A fork release.** Stages 9 to 12 are local. docx4j-python cannot release on them before one
  (`RELEASING.md`).
- **docx4j's agreement run** (it calls `revisions.accept_all`): rerun by docx4j-25 on 2026-10-22
  against `30abcd2` (this phase), through `.venv-fork` with the fork's stage 12 as installed,
  docx4j's side at `f62314d3b`'s revisions code (unchanged since `84b202c5a`): **0 lines differ in
  all 24 comparisons** (the eleven fixtures, accepted and rejected, and the 2026-10-21 set's
  `review-for-word-word365-rejected`, accepted and rejected). Section 5's gate holds, as the audit's
  Accept All and Reject All forms, unchanged on all 384 documents, said it would.
- **CLAUDE.md** described VML as untyped in places, for the coordinator: done in `30abcd2` (the
  generated packages' list names `vml/`, the CR-012 line, the MCE paragraph, the fork's
  `lexical_types` beside the boolean rule, `scripts/vml_census.py` among the commands). "WordprocessingML
  only until CR-001 D and CR-002 C" stays: VML is WordprocessingML's here.
- **`docs/UPSTREAM.md`**: rows 19 to 24 stood after the notes below the table, where they did not
  render as table rows, and report 25 was appended after them. This predates the phase. (The
  coordinator's check found each row once, not twice as first reported, and moved rows 19 to 25
  into the table.)

### 12.14 The Word check (2026-10-22: files 1, 2 and 4 passed; file 3 failed, on `w:object`'s order)

Four files are in `out/cr012-c-word-check/`, with a `README.txt` saying, for each, the source, what
was done and what Word should show. They are also copied to
`~/vbShares/Office 2016/__tmp/python-cr012c/`. Each unmodified source is beside its save as
`cr012c-<n>-SOURCE-<name>.docx`.

Each was loaded with every WordprocessingML part unmarshalled, so that this engine writes the typed
VML back, and saved. Each reads back with `LoadOptions(strict=True)` with nothing skipped and no
`xsi:type` written. Every WordprocessingML part is canonically identical to its source, file 3's
`w:object` order apart (12.13). Against the same save made by the engine before this phase, every
part is the same in C14N: the parts holding VML differ in attribute order only, and `settings.xml`
is the same bytes.

- `cr012c-1-tracked-vml-text-box.docx`: docx4j-core-ts's `cr043-tracked.docx`.
  - Its C22 Fallback, a VML text box, is typed down to its tracked deletion and insertion, the
    Choice's twins.
  - Word 365 shows the Choice's box. The README names Review > Display for Review > All Markup, and
    what Simple Markup shows instead.
- `cr012c-2-header-watermarks.docx`: docx4j's `loadAndSave.docx`.
  - Three headers each hold a VML text watermark ("SAMPLE") straight in a run.
  - Its settings have `w:shapeDefaults` and `w:hdrShapeDefaults`.
- `cr012c-3-ole-object.docx`: docx4j's `OLE/inserted doc.docx`, an embedded Word document shown as
  an icon (`w:object`, `v:shape`, `o:OLEObject`), with no `mc:AlternateContent`.
- `cr012c-4-vml-only-text-box.docx`: docx4j's `vml/textbox.docx`, a Word 2010 document with a
  legacy VML text box straight in a run, its `o:gfxdata` wrapped across lines, and no
  `mc:AlternateContent`.

Expected: each opens with no repair prompt and looks as its source does.

**The result, 2026-10-22** (the user: "__tmp/python-cr012c/ all good except the ole-object case,
double clicking does not open the embedded document (compare SOURCE, where it works)"):

- **Files 1, 2 and 4 passed.**
- **File 3 failed.** The icon shows, but double-clicking it does not open the embedded document;
  in the source it does. The coordinator's comparison of the two packages:
  - the embedding (`word/embeddings/Microsoft_Word_97_-_2003_Document.doc`), the image and
    `[Content_Types].xml` are the same bytes, and both relationships parts are the same in C14N;
  - in `document.xml` the one difference beyond attribute order is 12.13's: the save writes
    `o:OLEObject`, whose `ShapeID="_x0000_i1025"` names the shape, before the
    `v:shape id="_x0000_i1025"`, where the source has the shape first.
  - **The order is older than this phase and has shipped.** Saving the source with `f062c25` (before
    this phase) and with the `v0.1.0` tag writes the same order. So any document with an OLE object
    whose main part docx4j-python re-marshals (any edit of the body) loses the object's activation.
- **Confirming the cause.** Two files, each differing from its base only in that order, are staged
  in `~/vbShares/Office 2016/__tmp/python-cr012c-ole/` with the source and a README:
  `cr012c-3a-save-with-shape-first.docx` (the save, `o:OLEObject` moved back after the `v:shape`)
  and `cr012c-3b-SOURCE-with-OLEObject-first.docx` (Word's source, `o:OLEObject` moved before it).
  **Confirmed 2026-10-22** (the user: 3a, "Double-clicking should work. correct"; 3b,
  "Double-clicking should fail. correct"): the order alone decides it. With `o:OLEObject` before the
  `v:shape` its `ShapeID` names, Word shows the icon and quietly does not activate the object; with
  the shape first, it does.
- **The class of bug.** A repeating `xsd:sequence` of two or more particles is serialised
  interleaved, one particle's list then the next's, whatever the source's order. Over the schema
  copy there are two: `CT_PictureBase` (wml.xsd, the base of `CT_Picture`, `CT_Object` and
  `CT_Background`: `w:pict`, `w:object` and `w:background`), a repeating sequence of a VML wildcard
  and an `o:` wildcard, and inkml's `traceGroup.type`, which no corpus document holds. JAXB binds
  such a sequence as one content list, in document order (docx4j's `Pict.getAnyAndAny()`).

### 12.15 The `w:object` order fix (proposed 2026-10-22; waiting on the user's go)

12.14's file 3 fails because a save writes `w:object`'s children in an order Word does not
activate, and that order comes from the binding, not from VML: `CT_PictureBase` (wml.xsd, the base
of `CT_Picture`, `CT_Object` and `CT_Background`) is
`<xsd:sequence maxOccurs="unbounded">` of a lax `xsd:any` for VML and one for `o:`, which the
generator makes two lists (`schemas_microsoft_com_vml_element`,
`schemas_microsoft_com_office_office_element`) and the serialiser writes interleaved, one item of
each in turn. The fix is Phase C's, since its Word check needs it, and is also a fix to every
release so far.

1. **A marked schema patch** (`docx4j-python` comment, `schemas/PATCHES.md` entry): the repeating
   sequence becomes one `xsd:any processContents="lax"
   namespace="urn:schemas-microsoft-com:vml urn:schemas-microsoft-com:office:office"
   minOccurs="0" maxOccurs="unbounded"`. **It accepts the same documents:** both particles are
   optional, so `(v? o?)*` is `(v | o)*`. It is JAXB's own reading of the type, one content list
   in document order (docx4j's `Pict.getAnyAndAny()`).
2. **One generated list, in document order.** A scratch generation of both shapes with the fork
   (outside this repository) gave the pair two lists, written `v:shapetype`, `o:OLEObject`,
   `v:shape` from a source of `v:shapetype`, `v:shape`, `o:OLEObject`, and the single wildcard one
   list, written as read. The generator names the field
   `schemas_microsoft_com_vml_schemas_microsoft_com_office_office_element`; recommended: docx4j's
   name, `any_and_any`, by a field `Substitution` in `.xsdata.phase-b.xml`, if the fork's
   generator takes one for a field, and the generator's name otherwise.
3. **Every reader of the two old fields** moves to the one list: the text model's `_object` and
   `w:pict` walk, the revisions engine's trees, the content API's pictures, `insert_ooxml`, the
   census scripts and the tests that name them. Phase C has not been released, so the rename is
   one CHANGELOG entry with the rest of Phase C's.
4. **inkml's `traceGroup.type`** has the same shape (a repeating sequence of five elements). It is
   patched the same way only if every particle is optional (the same language), and is otherwise
   recorded and left; no corpus document holds ink.
5. **The gates:** Phase C's, rerun. The survey's 23 parts where the schema's order was imposed
   should lose those whose only cause was `CT_PictureBase` (the `w:object`s of the OLE fixtures and
   the extraction probes), with nothing newly skipped and no part gaining a difference. A test pins
   `w:object`'s order through a load and save. The suites, the round trip, acceptance and the
   readers' audit are as Phase C's. **The Word check:** file 3 saved again by the fixed engine,
   whose embedded document must open on a double-click, beside its source.
6. **A fix release** for the 0.1.0 line is the user's decision. The patch needs no fork change, so it
   could be made at the `v0.1.0` tag over `docx4j-xsdata 26.2.1`; `main` itself cannot release
   before the fork's stages 9 to 12 do.
