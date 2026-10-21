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
on 2026-10-22** (the user: "python-cr012a2 each pass"; 10.9): **Phase A2 done**.
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
- **Phase B:** a further change to the fork's runtime (section 4.2).
- **Phase C:** docx4j CR-048 (`acd0defc9` on `VERSION_17_3_4`, pushed 2026-10-18, to ship in
  17.3.4), whose VML schema this phase copies.

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

After Phase A, a block-level branch kept in `preserve` mode holds `w:p` and `w:tr` (global here),
`w:sdt` (typed by Phase A's kind rule), and `w:tbl`, `w:tc` and `w:bookmarkStart` (among the 61
below). So this phase types those branches too.

1. **The fork's runtime: root-element names.** The `XmlContext` gains a table from qualified
   element name to class, registered by the caller. JAXB gets the same thing from
   `@XmlRootElement`.
   - **Parsing:** in a wildcard whose `process_contents` isn't `"skip"`, the parser checks the
     table after `xsi:type` and the context's own `find_type(qname)`. A registered name gives a
     plain instance of its class.
   - **Saving:** an instance written without a field name (in a wildcard, or as the document
     root) takes its registered name instead of its type name. The table must be one-to-one,
     class to name, and registration checks it.
   - The change gets tests in the fork's `tests/fork/`, an entry in `docs/fork/CHANGES.md` and a
     row in this repository's `docs/UPSTREAM.md`.
2. **The table, derived mechanically.**
   - `codegen/derive_names.py` writes `codegen/names/root_elements.json` from docx4j's XJC
     output, the source of the other name tables. It takes every `@XmlRootElement` class and
     unscoped `@XmlElementDecl` whose name this model doesn't declare globally, and gives each
     the class the `el` table names for it.
   - `w:sdtContent` is left out (Q3).
   - It is cross-checked against the schema copy's annox annotations (61 of the 62).
   - The file is committed for review and never hand-edited, like the other name tables.
3. **Generated, then registered.**
   - `codegen/generate_el.py` writes the table into `docx4j_py/el_index.py` as `ROOT_ELEMENTS`,
     the module's class paths as strings.
   - `docx4j_py/runtime.context()`, hand-written, registers it when it builds the shared context.
   - `codegen/generate.sh --check` stays byte-identical.
4. **The readers the change exposes.**
   - `traversal.text_of` and `walk` read a typed child of a kept branch as they read it
     elsewhere (the three `test_mce.py` failures), under every `mce` mode.
   - The two `walk(mce="all")` tests expect `Wsp`.
   - `revisions/islands.py`'s docstring says that only the VML Fallback's text box still needs an
     island. The code stays as it is, since the Fallback still needs it.
5. **The record.**
   - CLAUDE.md's MCE paragraph: `mce_preprocess` is no longer the only way to get a typed
     `Drawing` in a choice.
   - A note in CR-002 section 5.6.
   - CHANGELOG "Changed": a kept branch's WordprocessingML children are typed. Code that looks
     for an `AnyElement` `w:drawing` in `any_element` now finds a `Drawing`.

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

**The Word check (Not yet checked).** Three files are in `out/cr012-a2-twin-word-check/`, with a
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
or `Fkeep` shows.

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
