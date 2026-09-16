# Patches to the schema copy

`schemas/` is docx4j's `xsd/` tree — the 91-file transitive closure of `wml/wml.xsd`, plus
`relationships.xsd` and `docProps/` (CR-002 section 9) and `docx4j_python__ROOT.xsd`, which is not
a docx4j schema at all but the one entry point `codegen/generate.sh` hands the generator — copied
verbatim except for the patches listed here. Every patch is marked in the XSD by a
comment beginning `docx4j-python:` on the line before it, so
`grep -rn 'docx4j-python' schemas/` finds all of them and a re-copy from docx4j can be
re-patched mechanically.

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

| File | Type | Part | Already had `mc:Ignorable`? |
|---|---|---|---|
| `wml/wml.xsd` | `w:document` (anonymous) | `word/document.xml` | yes |
| `wml/wml.xsd` | `w:styles` (anonymous) | `word/styles.xml`, `stylesWithEffects.xml` | yes |
| `wml/wml.xsd` | `w:numbering` (anonymous) | `word/numbering.xml` | yes |
| `wml/wml.xsd` | `w:fonts` (anonymous) | `word/fontTable.xml` | **no** |
| `wml/wml.xsd` | `w:comments` (anonymous) | `word/comments.xml` | **no** |
| `wml/wml.xsd` | `w:hdr` (anonymous) | `word/headerN.xml` | yes |
| `wml/wml.xsd` | `w:ftr` (anonymous) | `word/footerN.xml` | yes |
| `wml/wml.xsd` | `w:glossaryDocument` (anonymous) | `word/glossary/document.xml` | **no** |
| `wml/wml.xsd` | `CT_Settings` | `word/settings.xml` | yes |
| `wml/wml.xsd` | `CT_WebSettings` | `word/webSettings.xml` | **no** |
| `wml/wml.xsd` | `CT_Footnotes` | `word/footnotes.xml` | yes |
| `wml/wml.xsd` | `CT_Endnotes` | `word/endnotes.xml` | yes |
| `wml/wml.xsd` | `CT_Recipients` | mail-merge recipients part | **no** |
| `wml/w15_word_2012_wordml.xsd` | `CT_CommentsEx` | `word/commentsExtended.xml` | no |
| `wml/w15_word_2012_wordml.xsd` | `CT_People` | `word/people.xml` | no |
| `offmacro/office-word-2006-wordml.xsd` | `CT_VbaSuppData` | `word/vbaData.xml` | no |
| `offmacro/office-word-2006-wordml.xsd` | `CT_Tcg` | `word/attachedToolbars.bin` companion | no |
| `customXml/shared-customXmlSchemaProperties.xsd` | `sl:schemaLibrary` (anonymous) | custom XML schema library part | no |
| `dml/dml-stylesheet.xsd` | `a:theme` (anonymous) | `word/theme/themeN.xml` | no |
| `dml/dml-stylesheet.xsd` | `CT_BaseStylesOverride` | `a:themeOverride` part | no |
| `dml/dml-stylesheet.xsd` | `CT_EmptyElement` | `a:themeManager` part | no |

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
  unrelated error about a `chartex` QName; `dml-stylesheet.xsd` and the customXml schema are
  valid before and after).
* `grep -ro 'docx4j-python' schemas --include='*.xsd' | wc -l` = **26** marker comments in the
  schemas: 21 wildcards, one for `CT_Dir`/`CT_Bdo`, three for patches 3 and 4 (the new
  complexType, the element's `type=`, and the swap), and one for patch 5 (`@title`; it was 25
  before CR-003 Phase A). A twenty-seventh is in `docx4j_python__ROOT.xsd`, which is ours in
  its entirety and not a patch.
* Round trip: `attribute-dropped` falls from 10 to 0 (see `codegen/README.md`).

## 3. `w:stylePaneFormatFilter`: `CT_StylePaneFilter` (CR-002 Phase A)

`wml/wml.xsd`, two marked places: the new `CT_StylePaneFilter` complexType, just after
`CT_ShortHexNumber`, and the `w:stylePaneFormatFilter` element declaration in `CT_Settings`,
whose `type=` changes from `CT_ShortHexNumber` to it.

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

**Why.** docx4j's copy declares `w15:chartTrackingRefBased` first. Every `settings.xml` in
`samples/` that has both writes `w14:docId` first, so a serialiser that followed the schema
reordered two elements of a part it had only read —
`2016_image_with_text_effects.docx` and `DrawingML_GraphicData_wps.docx` each showed a two-element
`child-order` difference. Both elements are `minOccurs="0"` and nothing depends on their relative
order.

## 5. `CT_NonVisualDrawingProps/@title` (CR-003 Phase A)

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
