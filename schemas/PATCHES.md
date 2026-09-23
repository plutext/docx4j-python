# Patches to the schema copy

`schemas/` is docx4j's `xsd/` tree — the 95-file transitive closure of `wml/wml.xsd`, plus
`relationships.xsd` and `docProps/` (CR-002 section 9), docx4j's own `wml/wml__ROOT.xsd` (kept
from the first copy, referenced by nothing here) and `docx4j_python__ROOT.xsd`, which is not
a docx4j schema at all but the one entry point `codegen/generate.sh` hands the generator — copied
verbatim except for the patches listed here. Every patch is marked in the XSD by a
comment beginning `docx4j-python:` on the line before it, so
`grep -rn 'docx4j-python' schemas/` finds all of them and a re-copy from docx4j can be
re-patched mechanically.

**The copy is from docx4j commit `c363a969f` (2026-09-20; three files --- `shared/shared-math-2ed.xsd`, `dml/dml-textCharacter.xsd`, `odrawxml/office_drawing_2017_03_chart_.xsd`, docx4j's CR-025 and its c16r3 fix --- copied on top of the `58cce65a8` refresh the same day, CR-001 section 14.11; none of the three carries a patch).** It was first taken at `6b8048aab`
(2026-09-05, 90 files in the WML closure) and refreshed on 2026-09-20 (CR-001 section 14.10):
docx4j's CR-018, CR-021 phase 2, CR-023 and CR-024 and the lax mce wildcards changed nine files
of the closure (`wml/wml.xsd`, `wml/w14_word_2010_wordml.xsd`, `wml/w15_word_2012_wordml.xsd`,
`dml/dml-chart.xsd`, `dml/dml-documentProperties.xsd`, `dml/dml-spreadsheetDrawing.xsd`,
`dml/dml-textParagraph.xsd`, `mce/markup-compatibility-2006-MINIMAL.xsd`,
`odrawxml/oart14docprop.xsd`) and added five that `wml/wml.xsd` now imports (`wml/w15symex.xsd`,
`wml/w16cid.xsd`, `wml/w16du.xsd`, `wml/w16sdtdh.xsd`, `wml/w16sdtfl.xsd`, docx4j's CR-023).
Line endings are docx4j's: four files docx4j normalised to LF came across as LF
(`dml-textParagraph`, the mce schema, `oart14docprop`, `w15_word_2012_wordml`), and two the first
copy had normalised itself went back to docx4j's CRLF (`customXml/shared-customXmlSchemaProperties.xsd`,
`dml/dml-stylesheet.xsd`), which is why their diffs at the refresh are whole-file. `cei.xsd`,
`w16cex.xsd` and `w16.xsd` exist in docx4j's `xsd/wml/` but are outside `wml.xsd`'s closure (bound
through docx4j's `ROOT.xsd`) and are not copied; SML (`xsd/xlsx`, `sml`) is CR-001 Phase D's.

Nothing else is changed. In particular `default="..."` is **not** stripped — the schema
default is kept and the generator's `<SchemaDefaults>metadata</SchemaDefaults>` option puts
it in the field metadata instead of the field default (CR-001 section 7).

## 1. `CT_Dir` / `CT_Bdo`: anonymous types promoted (pre-existing, Phase A)

`wml/wml.xsd` line ~8631 and the two `xsd:complexType` declarations that follow.

`w:dir` and `w:bdo` were declared inside the recursive group `EG_ContentRunContent` with
anonymous `xsd:complexType`s that themselves reference `EG_PContent`. xsdata 26.2 fails with
`Error: Missing inner class` (`REPORT.md` section 2.1; deferred upstream report 2). The two
anonymous types are promoted to named `CT_Dir` and `CT_Bdo`; semantics are unchanged. XJC
accepts the original, so docx4j's own build is unaffected.

Its marker reads `docx4j-python patch:` rather than `docx4j-python:`; both match the grep.

**Status at the 2026-09-20 refresh: kept.** docx4j `58cce65a8` still declares the two anonymous
types inside the group.

## 2. `xsd:anyAttribute` on every part-root complexType (Phase B)

```xml
<!-- docx4j-python: part root; anyAttribute so mc:Ignorable and future
     root attributes survive the round trip (CR-001 section 7) -->
<xsd:anyAttribute namespace="##other" processContents="lax"/>
```

inserted as the last child of the complexType (XSD requires `anyAttribute` after all
`attribute` / `attributeGroup` declarations).

**Why.** A part root carries attributes the schema does not declare — `mc:Ignorable` above
all, and in future any `w16*` root attribute. xsdata drops an undeclared attribute silently
(`REPORT.md` 7.2 and 7.5); with the wildcard it maps them to an `other_attributes: dict`
field and round-trips them exactly. Ten of the corpus's fifty parts lost
`mc:Ignorable='w14 …'` on `/w:fonts` before this patch.

**Why `##other` and not `##any`.** `##other` is every namespace but the schema's own target
namespace, which is exactly the undeclared-attribute case (`mc:`, `w14:`, `w16se:`); it is
also the form `REPORT.md` 7.2 verified. `##any` would put the wildcard in competition with
the type's own declared attributes. Where the type already declares
`<xsd:attribute ref="mc:Ignorable"/>` the declared attribute still wins — a named attribute
use takes precedence over the wildcard — so `mc:Ignorable` stays a typed field and the
wildcard only catches the rest.

**Not stripped:** the existing `<xsd:attribute ref="mc:Ignorable" use="optional"/>`
declarations are left exactly where they were.

### The 21 part-root complexTypes patched

**Status at the 2026-09-20 refresh: all 21 kept.** docx4j's CR-018 and CR-023 declared
`mc:Ignorable` on six more of these roots (the last column), which makes `mc:Ignorable` a typed
field there as it already was on the others; the wildcard stays for every attribute the schema
does not declare (a future `w16*` root attribute), and docx4j added no `anyAttribute` of its own to
any of the 21. Each wildcard is re-inserted as the last child, after docx4j's new declaration.

| File | Type | Part | Had `mc:Ignorable` at `6b8048aab`? | At `58cce65a8`? |
|---|---|---|---|---|
| `wml/wml.xsd` | `w:document` (anonymous) | `word/document.xml` | yes | yes |
| `wml/wml.xsd` | `w:styles` (anonymous) | `word/styles.xml`, `stylesWithEffects.xml` | yes | yes |
| `wml/wml.xsd` | `w:numbering` (anonymous) | `word/numbering.xml` | yes | yes |
| `wml/wml.xsd` | `w:fonts` (anonymous) | `word/fontTable.xml` | **no** | yes (CR-023) |
| `wml/wml.xsd` | `w:comments` (anonymous) | `word/comments.xml` | **no** | yes (CR-018) |
| `wml/wml.xsd` | `w:hdr` (anonymous) | `word/headerN.xml` | yes | yes |
| `wml/wml.xsd` | `w:ftr` (anonymous) | `word/footerN.xml` | yes | yes |
| `wml/wml.xsd` | `w:glossaryDocument` (anonymous) | `word/glossary/document.xml` | **no** | yes (CR-023) |
| `wml/wml.xsd` | `CT_Settings` | `word/settings.xml` | yes | yes |
| `wml/wml.xsd` | `CT_WebSettings` | `word/webSettings.xml` | **no** | yes (CR-023) |
| `wml/wml.xsd` | `CT_Footnotes` | `word/footnotes.xml` | yes | yes |
| `wml/wml.xsd` | `CT_Endnotes` | `word/endnotes.xml` | yes | yes |
| `wml/wml.xsd` | `CT_Recipients` | mail-merge recipients part | **no** | no |
| `wml/w15_word_2012_wordml.xsd` | `CT_CommentsEx` | `word/commentsExtended.xml` | no | yes (CR-018) |
| `wml/w15_word_2012_wordml.xsd` | `CT_People` | `word/people.xml` | no | yes (CR-018) |
| `offmacro/office-word-2006-wordml.xsd` | `CT_VbaSuppData` | `word/vbaData.xml` | no | no |
| `offmacro/office-word-2006-wordml.xsd` | `CT_Tcg` | `word/attachedToolbars.bin` companion | no | no |
| `customXml/shared-customXmlSchemaProperties.xsd` | `sl:schemaLibrary` (anonymous) | custom XML schema library part | no | no |
| `dml/dml-stylesheet.xsd` | `a:theme` (anonymous) | `word/theme/themeN.xml` | no | no |
| `dml/dml-stylesheet.xsd` | `CT_BaseStylesOverride` | `a:themeOverride` part | no | no |
| `dml/dml-stylesheet.xsd` | `CT_EmptyElement` | `a:themeManager` part | no | no |

Deliberately **not** patched, though they are the types of global elements:

* `CT_Body`, `CT_TxbxContent`, `w:docDefaults`, `w:style`, `w:p`, `w:r`, … — they are never a
  part root, only content inside one.
* Everything in PML, SML and the chart/diagram namespaces — Phase D.

Known side effect: `CT_EmptyElement` is shared with `a:masterClrMapping` and
`a:overrideClrMapping`, which therefore also gain an `other_attributes` dict. It is lossless
and harmless, but it is the one patch that widens a type used outside a part root.

## Verification

* `lxml.etree.XMLSchema` parses each patched schema with exactly the same result as before the
  patch (`wml.xsd`, `w15_…`, `office-word-2006-wordml.xsd` all carry one pre-existing,
  unrelated error about a `chartex` QName; the customXml schema is valid before and after;
  `dml-stylesheet.xsd` was valid at `6b8048aab` and since docx4j's CR-021 phase 3 rewrote
  `odrawxml/oart14docprop.xsd` reports, patched or not, an import-order error about `a:AG_Locking`
  that only lxml's standalone parse of that one file sees — the generator starts from the root).
* `grep -ro 'docx4j-python' schemas --include='*.xsd' | wc -l` = **26** marker comments in the
  schemas: 21 wildcards, one for `CT_Dir`/`CT_Bdo`, three for patches 3 and 4 (the new
  complexType, the element's `type=`, and the swap), and one in `docx4j_python__ROOT.xsd`, which
  is ours in its entirety and not a patch. (It was 27 between CR-003 Phase A and the 2026-09-20
  refresh, when patch 5 was retired.) Twenty-five patch markers, then.
* The refresh itself: each re-patched file differs from the previous copy by exactly docx4j's
  own commits (the hunks of `git diff` between the two copies, less the marker lines, equal
  `git diff 6b8048aab..58cce65a8 -- xsd/<file>`).
* Round trip: `attribute-dropped` falls from 10 to 0 (see `codegen/README.md`).

## 3. `w:stylePaneFormatFilter`: `CT_StylePaneFilter` (CR-002 Phase A)

`wml/wml.xsd`, two marked places: the new `CT_StylePaneFilter` complexType, just after
`CT_ShortHexNumber`, and the `w:stylePaneFormatFilter` element declaration in `CT_Settings`,
whose `type=` changes from `CT_ShortHexNumber` to it.

**Status at the 2026-09-20 refresh: kept**; docx4j `58cce65a8` still types the element
`CT_ShortHexNumber` and declares no `CT_StylePaneFilter`.

**Why.** ECMA-376 2nd edition types this element `CT_StylePaneFilter`, with fifteen `ST_OnOff`
attributes; docx4j's copy is 1st edition and types it `CT_ShortHexNumber`, which declares only
`w:val`. Word writes all sixteen. CR-001 section 14.8 point 5 found that **15 attributes were
dropped from `samples/toc.docx`'s `settings.xml`**, which the skipped-content report caught; it is
0 now, and `settings.xml` and `webSettings.xml` joined the round trip that CR-002 section 12.2
measures.

```xml
<xsd:complexType name="CT_StylePaneFilter">
  <xsd:attribute name="allStyles" type="xsd:boolean" use="optional"/>
  … customStyles, latentStyles, stylesInUse, headingStyles, numberingStyles, tableStyles,
    directFormattingOnRuns, directFormattingOnParagraphs, directFormattingOnNumbering,
    directFormattingOnTables, clearFormatting, top3HeadingStyles, visibleStyles,
    alternateStyleNames …
  <xsd:attribute name="val" type="ST_ShortHexNumber" use="optional"/>
</xsd:complexType>
```

**Two choices worth stating.** The `ST_OnOff` attributes are written `xsd:boolean` rather than as
the schema's `ST_OnOff` enumeration (`true|false|on|off`), because the enumeration cannot hold
Word's `1`/`0`; `xsd:boolean` can hold all four of `1 0 true false` and the serialiser writes
`1`/`0` as Word does. A producer that wrote `on`/`off` *here* would lose them, and none does.
And `w:val` is kept `use="optional"` so that a 1st-edition document with only `w:val="3F01"` still
round-trips.

## 4. `w14:docId` before `w15:chartTrackingRefBased` (CR-002 Phase A)

`wml/wml.xsd`, in `CT_Settings`: the two element declarations are swapped.

**Status at the 2026-09-20 refresh: kept**; docx4j `58cce65a8` still declares
`w15:chartTrackingRefBased` first. **Subsumed at docx4j `dae2dfc8b`** (2026-09-23, on its
`VERSION_17_2_1` branch, reported by the docx4j session): docx4j now declares `w14:docId`,
`w15:chartTrackingRefBased`, `w15:docId` in that order --- **exactly this patch's order**,
measured on Word 365 output --- with `SettingsDocIdOrderTest` pinning it. So this patch retires at
the next refresh, and nothing this repository writes changes: the two corpus documents that carry
all three already write them in that order and the round trip is clean either way. docx4j's note
that **Word Online writes `w15:chartTrackingRefBased` first** is the accepted cost on both sides:
a serialiser that follows the schema reorders those two elements of a Word Online document it had
only read, which `canon.py` reports as `child-order`. No such document is in `samples/`.

**Why.** docx4j's copy declares `w15:chartTrackingRefBased` first. Every `settings.xml` in
`samples/` that has both writes `w14:docId` first, so a serialiser that followed the schema
reordered two elements of a part it had only read —
`2016_image_with_text_effects.docx` and `DrawingML_GraphicData_wps.docx` each showed a two-element
`child-order` difference. Both elements are `minOccurs="0"` and nothing depends on their relative
order.

## 5. `CT_NonVisualDrawingProps/@title` (CR-003 Phase A) — retired 2026-09-20

**Retired at the 2026-09-20 refresh, subsumed by docx4j commit `d487076d3` (CR-018 item 2,
2026-09-16):** docx4j's `dml-documentProperties.xsd` carries the identical attribute, so the copy
is docx4j's file verbatim and no marker remains. The text below is the record of the patch.

`dml/dml-documentProperties.xsd`, in `CT_NonVisualDrawingProps`, between `descr` and `hidden`:

```xml
<xsd:attribute name="title" type="xsd:string" use="optional" default=""/>
```

**Why.** `wp:docPr/@title` is Office JS's `InlinePicture.altTextTitle` and the twin of `@descr`
(`altTextDescription`), and CR-003 Phase A's `inline_picture(..., title=)` writes it. It is in
ECMA-376 4th edition Part 1 20.1.2.2.8 (Transitional) and Word reads and writes it; docx4j's
`xsd/` has carried it since its own CR-018 item 2, so this patch makes the copy **identical to
docx4j's current schema** rather than adding anything of ours — it is the one attribute
`diff -r schemas/dml ~/git/docx4j/xsd/dml` reported, and that diff is now empty for this file.

The whole regeneration it caused is seven lines: `title: None | str` on
`docx4j_py.dml.main.CTNonVisualDrawingProps`. No element name, and therefore no `el` entry,
changes; `codegen/generate.sh --check` passes.
