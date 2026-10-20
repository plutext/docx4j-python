# CR-012: Keep `mc:AlternateContent` at every level, and type what its branches hold, as docx4j does

**Status:** Proposed 2026-10-19, at the user's request ("docx4j and docx4j-generated-objects-ts
both retain all mc content ... do we? if not, please draft a CR"). Section 1 answers the
question. **Revised 2026-10-21** at the user's request ("assess CR-051 and modify your CR-012 as
appropriate"), after docx4j's CR-051. Section 1.1 corrects section 1's answer: an element at
block, row or cell level is not kept, and a content control directly in a branch loses its runs.
Section 3.1 measures both on Jason's probes. A new Phase A (section 4.1) ports CR-051's phase 1.
The 61 roots, which were Phase A, are now Phase B, and VML is Phase C. The questions of section 7
carry recommendations. **Implementation waits on the user's go.**
**Depends on:**
- **Phase A:**
  - docx4j CR-051 phase 1 (`6687cc4b8`, `1cff2cd5c` on `VERSION_17_3_4`, unpushed), whose
    schema admissions the schema copy takes;
  - docx4j's published understood list (`understood-namespaces.txt`, CR-046), which CR-011's
    Q2 also ports;
  - a change to the xsdata fork's runtime (section 4.1 item 2).
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

The probes are Jason's 25 `mc-fun` documents (`~/vbShares/Office 2016/mc-fun/`, also in
`docx4j-core-tests/src/test/resources/mc-fun/` since CR-051 phase 1). Most carry ten Choices,
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
   `LoadOptions(mc_block_level="word" | "preserve")`, default `"word"`. A package-level setting,
   `pkg.mc_block_level`, is read when each part is first unmarshalled, for packages made
   otherwise.
   - **`word`:** each element whose parent is a block, row or cell container is replaced by the
     content of the branch Word reads. The rule is ISO's: the first Choice whose `Requires` are
     understood, else the Fallback. It runs over docx4j's published understood list, matched by
     namespace URI.
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
   passes the element by without error; none raises or miscounts. Row and cell level in
   `preserve` mode are passed by, as docx4j's phase 2 leaves them until then; recorded.
6. **The revisions engine at paragraph level.**
   - docx4j's revisions model reads a paragraph-level element's selected branch as text, and
     lists the other branches' changes as twins (CR-051 section 10.5 item 3,
     `TrackedChangesAlternateContentTest`).
   - This engine's `_has_text` reads such an element as holding no text (found in CR-010's U23
     port).
   - Ported here, held to docx4j's test.
7. **The probes:** docx4j's copies of the 25, with Word 365's saves, copied into
   `tests/fixtures/mc-fun/`.

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

- **`word` mode changes a saved document** wherever block-, row- or cell-level alternate content
  is read: the branches Word does not read go. Word's own save does the same, and the record
  keeps them for inspection. A caller that must reproduce its input loads in `preserve` (Q7).
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

## 7. Questions

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

   *Recommended: (a).* The content API, the python-docx facade and every reader then see what
   Word shows, with no reader changed, and a save writes what Word's save writes. Tools that must
   reproduce their input, such as CR-006's MCP edit tool, load in `preserve`. (b) keeps the
   round-trip rule whole, but leaves every reader blind to Word's branch at block level until
   each one learns the element.
8. **Q8: the understood set:** docx4j's published list, both for `word` mode and for
   `mce_preprocess`, shared with CR-011's Q2. *Recommended.* CR-011's Q2 had kept `UNDERSTOOD`
   for `mce_preprocess`. Section 3.1 shows it then misses Word 365's branch, so this revises
   that half of Q2. `UNDERSTOOD` stays as the prefix table's own list.

## 8. Phases and effort

| Phase | What | Effort |
|---|---|---|
| A | docx4j CR-051 phase 1: the schema refresh and regeneration; the fork's container-dependent `w:sdt`; the load-time mode with its record, REQ-305 and `mc:Ignorable`; `mce_preprocess` on docx4j's list; the readers' audit; the revisions twins; the 25 probes and Word's saves; a Word check | 3 days |
| B | the fork's root-element table and its tests; `derive_names.py`, `root_elements.json`, `ROOT_ELEMENTS`; the registration; traversal's typed branch children; tests, survey script, the record; then a fork release before a docx4j-python release | 1.5 days |
| C | docx4j's VML schemas at `acd0defc9`, names, regeneration; the VML survey and gate; the islands retired; Word check | 3 days |
