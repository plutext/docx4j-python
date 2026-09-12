# docx4j name tables

These JSON files map OOXML **schema type names** (`CT_PPr`, `ST_Jc`, …) and **global element
names** (`p`, `document`, `t`, …) to the **Java class names docx4j uses** for them, one file per
XML namespace. They exist so that the classes the xsdata fork generates from `schemas/` can be
given exactly the names a docx4j user already knows — the fork's `<ClassNames>` generator option
reads them. Both sides are derived from the same customised schemas, so the keys line up: docx4j
ran XJC over `docx4j/xsd/`, we run xsdata over `schemas/` (a copy of the same files). The tables
are derived mechanically from docx4j's JAXB-annotated generated Java (`@XmlType`,
`@XmlRootElement`, `ObjectFactory`'s `@XmlElementDecl`). Every `<prefix>.json` is machine
written; the one hand-maintained file is `wml_overrides.json`, [below](#the-two-hand-written-entries-wml_overridesjson).
Regenerate with:

```
python codegen/derive_names.py \
    [--java-root /home/jharrop/git/docx4j/docx4j-generated-objects/target/generated-sources/xjc] \
    [--schemas-root schemas] [--out codegen/names] [--report /tmp/anomalies.json]
```

The script is stdlib-only and idempotent (two runs produce byte-identical files). Each file has
exactly four top-level keys, in this order: `namespace`, `java_package`, `types`, `elements`;
`types` and `elements` are sorted by key.

Which namespaces get a file: the **intersection** of (a) the `targetNamespace` values found in
`schemas/` — i.e. WML's import closure as staged in this repo — and (b) the namespaces docx4j's
Java packages declare in their `package-info.java`. 54 namespaces qualify (CR-002 section 9 added seven: the package relationships and the docProps family). Both directions of the
difference are listed under [Namespaces with no table](#namespaces-with-no-table).

## Files

54 files, 1425 type entries and 516 element entries in total.

| file | namespace | java package | types | elements |
| --- | --- | --- | ---: | ---: |
| `customxml.json` | `http://schemas.openxmlformats.org/schemaLibrary/2006/main` | `org.docx4j.customxml` | 0 | 1 |
| `dml.json` | `http://schemas.openxmlformats.org/drawingml/2006/main` | `org.docx4j.dml` | 270 | 35 |
| `dml_chart.json` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | `org.docx4j.dml.chart` | 165 | 3 |
| `dml_chart_2007.json` | `http://schemas.microsoft.com/office/drawing/2007/8/2/chart` | `org.docx4j.dml.chart.x2007` | 5 | 3 |
| `dml_chart_drawing.json` | `http://schemas.openxmlformats.org/drawingml/2006/chartDrawing` | `org.docx4j.dml.chartDrawing` | 14 | 0 |
| `dml_compatibility.json` | `http://schemas.openxmlformats.org/drawingml/2006/compatibility` | `org.docx4j.dml.compatibility` | 1 | 1 |
| `dml_diagram.json` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` | `org.docx4j.dml.diagram` | 109 | 12 |
| `dml_diagram2008.json` | `http://schemas.microsoft.com/office/drawing/2008/diagram` | `org.docx4j.dml.diagram2008` | 6 | 2 |
| `dml_picture.json` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | `org.docx4j.dml.picture` | 2 | 1 |
| `dml_spreadsheetdrawing.json` | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` | `org.docx4j.dml.spreadsheetdrawing` | 17 | 3 |
| `dml_wordprocessing_drawing.json` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` | `org.docx4j.dml.wordprocessingDrawing` | 16 | 2 |
| `doc_props_core.json` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | `org.docx4j.docProps.core` | 0 | 1 |
| `doc_props_core_dc_elements.json` | `http://purl.org/dc/elements/1.1/` | `org.docx4j.docProps.core.dc.elements` | 2 | 16 |
| `doc_props_core_dc_terms.json` | `http://purl.org/dc/terms/` | `org.docx4j.docProps.core.dc.terms` | 18 | 33 |
| `doc_props_custom.json` | `http://schemas.openxmlformats.org/officeDocument/2006/custom-properties` | `org.docx4j.docProps.custom` | 0 | 1 |
| `doc_props_extended.json` | `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties` | `org.docx4j.docProps.extended` | 0 | 1 |
| `doc_props_variant_types.json` | `http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes` | `org.docx4j.docProps.variantTypes` | 7 | 7 |
| `drawing_2010_chart_drawing.json` | `http://schemas.microsoft.com/office/drawing/2010/chartDrawing` | `org.docx4j.com.microsoft.schemas.office.drawing.x2010.chartDrawing` | 3 | 1 |
| `drawing_2010_diagram.json` | `http://schemas.microsoft.com/office/drawing/2010/diagram` | `org.docx4j.com.microsoft.schemas.office.drawing.x2010.diagram` | 1 | 2 |
| `drawing_2010_main.json` | `http://schemas.microsoft.com/office/drawing/2010/main` | `org.docx4j.com.microsoft.schemas.office.drawing.x2010.main` | 42 | 13 |
| `drawing_2012_chart.json` | `http://schemas.microsoft.com/office/drawing/2012/chart` | `org.docx4j.com.microsoft.schemas.office.drawing.x2012.chart` | 18 | 26 |
| `drawing_2012_chart_style.json` | `http://schemas.microsoft.com/office/drawing/2012/chartStyle` | `org.docx4j.com.microsoft.schemas.office.drawing.x2012.chartStyle` | 13 | 2 |
| `drawing_2012_main.json` | `http://schemas.microsoft.com/office/drawing/2012/main` | `org.docx4j.com.microsoft.schemas.office.drawing.x2012.main` | 4 | 4 |
| `drawing_2013_main_command.json` | `http://schemas.microsoft.com/office/drawing/2013/main/command` | `org.docx4j.com.microsoft.schemas.office.drawing.x2013.main.command` | 25 | 6 |
| `drawing_2014_chart.json` | `http://schemas.microsoft.com/office/drawing/2014/chart` | `org.docx4j.com.microsoft.schemas.office.drawing.x2014.chart` | 11 | 9 |
| `drawing_2014_chartex.json` | `http://schemas.microsoft.com/office/drawing/2014/chartex` | `org.docx4j.com.microsoft.schemas.office.drawing.x2014.chartex` | 110 | 4 |
| `drawing_2014_main.json` | `http://schemas.microsoft.com/office/drawing/2014/main` | `org.docx4j.com.microsoft.schemas.office.drawing.x2014.main` | 4 | 5 |
| `drawing_201611_diagram.json` | `http://schemas.microsoft.com/office/drawing/2016/11/diagram` | `org.docx4j.com.microsoft.schemas.office.drawing.x201611.diagram` | 4 | 1 |
| `drawing_201611_main.json` | `http://schemas.microsoft.com/office/drawing/2016/11/main` | `org.docx4j.com.microsoft.schemas.office.drawing.x201611.main` | 1 | 1 |
| `drawing_2016_svg_main.json` | `http://schemas.microsoft.com/office/drawing/2016/SVG/main` | `org.docx4j.com.microsoft.schemas.office.drawing.x2016.SVG.main` | 1 | 1 |
| `drawing_201703_chart.json` | `http://schemas.microsoft.com/office/drawing/2017/03/chart` | `org.docx4j.com.microsoft.schemas.office.drawing.x201703.chart` | 2 | 1 |
| `drawing_2017_decorative.json` | `http://schemas.microsoft.com/office/drawing/2017/decorative` | `org.docx4j.com.microsoft.schemas.office.drawing.x2017.decorative` | 1 | 1 |
| `drawing_2017_model3d.json` | `http://schemas.microsoft.com/office/drawing/2017/model3d` | `org.docx4j.com.microsoft.schemas.office.drawing.x2017.model3d` | 16 | 1 |
| `drawing_2018_animation.json` | `http://schemas.microsoft.com/office/drawing/2018/animation` | `org.docx4j.com.microsoft.schemas.office.drawing.x2018.animation` | 2 | 0 |
| `drawing_2018_animation_model3d.json` | `http://schemas.microsoft.com/office/drawing/2018/animation/model3d` | `org.docx4j.com.microsoft.schemas.office.drawing.x2018.animation.model3d` | 2 | 2 |
| `drawing_2018_hyperlinkcolor.json` | `http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor` | `org.docx4j.com.microsoft.schemas.office.drawing.x2018.hyperlinkcolor` | 2 | 1 |
| `excel_2010_spreadsheet_drawing.json` | `http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing` | `org.xlsx4j.schemas.microsoft.com.office.excel.x2010.spreadsheetDrawing` | 3 | 1 |
| `ink_2010_main.json` | `http://schemas.microsoft.com/ink/2010/main` | `org.docx4j.com.microsoft.schemas.ink.x2010.main` | 6 | 1 |
| `inkml.json` | `http://www.w3.org/2003/InkML` | `org.docx4j.org.w3.x2003.inkML` | 33 | 8 |
| `math.json` | `http://schemas.openxmlformats.org/officeDocument/2006/math` | `org.docx4j.math` | 80 | 23 |
| `mathml.json` | `http://www.w3.org/1998/Math/MathML` | `org.docx4j.org.w3.x1998.math.mathML` | 8 | 45 |
| `mce.json` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | `org.docx4j.mce` | 0 | 1 |
| `powerpoint_2014_ink_action.json` | `http://schemas.microsoft.com/office/powerpoint/2014/inkAction` | `org.docx4j.com.microsoft.schemas.office.powerpoint.x2014.inkAction` | 10 | 1 |
| `relationships.json` | `http://schemas.openxmlformats.org/package/2006/relationships` | `org.docx4j.relationships` | 0 | 2 |
| `thememl_2012_main.json` | `http://schemas.microsoft.com/office/thememl/2012/main` | `org.docx4j.com.microsoft.schemas.office.thememl.x2012.main` | 1 | 1 |
| `w14.json` | `http://schemas.microsoft.com/office/word/2010/wordml` | `org.docx4j.w14` | 53 | 33 |
| `w15.json` | `http://schemas.microsoft.com/office/word/2012/wordml` | `org.docx4j.w15` | 9 | 13 |
| `wml.json` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | `org.docx4j.wml` | 301 | 175 |
| `word_2006_wordml.json` | `http://schemas.microsoft.com/office/word/2006/wordml` | `org.docx4j.com.microsoft.schemas.office.word.x2006.wordml` | 16 | 2 |
| `word_2010_wordprocessing_canvas.json` | `http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas` | `org.docx4j.com.microsoft.schemas.office.word.x2010.wordprocessingCanvas` | 1 | 1 |
| `word_2010_wordprocessing_drawing.json` | `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing` | `org.docx4j.com.microsoft.schemas.office.word.x2010.wordprocessingDrawing` | 4 | 2 |
| `word_2010_wordprocessing_group.json` | `http://schemas.microsoft.com/office/word/2010/wordprocessingGroup` | `org.docx4j.com.microsoft.schemas.office.word.x2010.wordprocessingGroup` | 2 | 1 |
| `word_2010_wordprocessing_shape.json` | `http://schemas.microsoft.com/office/word/2010/wordprocessingShape` | `org.docx4j.com.microsoft.schemas.office.word.x2010.wordprocessingShape` | 3 | 1 |
| `word_2012_wordprocessing_drawing.json` | `http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing` | `org.docx4j.com.microsoft.schemas.office.word.x2012.wordprocessingDrawing` | 1 | 1 |

`org.docx4j.wml` has 333 `.java` files = 331 classes + `ObjectFactory` + `package-info`, and
331 = 301 named `@XmlType`s + 30 anonymous (`@XmlType(name = "")`) ones. The tables account for
every class in the package.

## How the tables are derived

* **`types`** — for each `*.java` (excluding `ObjectFactory.java` / `package-info.java`), the
  **top-level** `@XmlType(name = "…")` with a non-empty name, mapped to the file's base name.
  Enums (`@XmlEnum`) are included. `@XmlType(name = "")` — an anonymous type, i.e. a class XJC
  made for a global element rather than a named schema type — contributes to `elements` only.
* **`elements`** — (a) a top-level `@XmlRootElement(name = "…")`, plus (b) `ObjectFactory`
  `@XmlElementDecl(namespace = …, name = …)` whose following factory method's payload type is a
  docx4j class. A declaration whose `namespace` is not the ObjectFactory's own package namespace
  is routed to that namespace's table (see [cross-package classes](#classes-owned-by-another-java-package));
  if that namespace has no table, the declaration is dropped and counted below.
* All annotation regexes are anchored at column 0, so only top-level classes are matched;
  every nested/inner class in the generated sources is indented. Inner classes still appear as
  *values* when an `ObjectFactory` declaration points at one (`P.Hyperlink`) — see
  [nested class names](#values-that-name-a-nested-java-class).
* **Conflicts** — where one element name yields two or more different classes, `@XmlRootElement`
  wins over `ObjectFactory`. If the winning source is itself ambiguous, **no entry is emitted**
  and the case is listed below.

## Why docx4j's names differ from the schema names

Three separate mechanisms, and it matters which one applies:

1. **XJC's default name mangling.** `CT_Settings` has no customisation, so XJC strips the
   underscore and camel-cases: **`CTSettings`**. Same for `ST_OnOff` → `STOnOff`. This is the
   majority of entries.

2. **Inline `<jaxb:class name="…"/>` in the schema.** docx4j's copies of the OOXML XSDs (which is
   what `schemas/` holds) carry JAXB customisations inline. `bindings.xjb` — the only `.xjb` under
   `/home/jharrop/git/docx4j/xsd/` — is six lines and renames nothing:

   ```xml
   <jaxb:bindings xmlns:jaxb="https://jakarta.ee/xml/ns/jaxb"
                  jaxb:version="3.0">
       <jaxb:globalBindings collectionType="org.docx4j.list.ArrayListDocx4j" />
   </jaxb:bindings>
   ```

   Every rename is instead an `<xsd:appinfo><jaxb:class name="…"/>` in the XSD, e.g.

   ```xml
   <xsd:complexType name="CT_PPr">
       <xsd:annotation><xsd:appinfo><jaxb:class name="PPr"/></xsd:appinfo></xsd:annotation>
   ```

   which is why `CT_PPr` is `PPr` and not `CTPPr`. All 62 such customisations in `schemas/`:

   | schema file | declaration | `jaxb:class name` |
   | --- | --- | --- |
   | `wml/wml.xsd` | complexType `CT_HpsMeasure` | `HpsMeasure` |
   | `wml/wml.xsd` | complexType `CT_Jc` | `Jc` |
   | `wml/wml.xsd` | complexType `CT_TextDirection` | `TextDirection` |
   | `wml/wml.xsd` | complexType `CT_ParaRPrChange` | `ParaRPrChange` |
   | `wml/wml.xsd` | complexType `CT_RunTrackChange` | `RunTrackChange` |
   | `wml/wml.xsd` | element `commentRangeStart` | `CommentRangeStart` |
   | `wml/wml.xsd` | element `commentRangeEnd` | `CommentRangeEnd` |
   | `wml/wml.xsd` | complexType `CT_Tabs` | `Tabs` |
   | `wml/wml.xsd` | complexType `CT_PPrBase` | `PPrBase` |
   | `wml/wml.xsd` | complexType `CT_PPr` | `PPr` |
   | `wml/wml.xsd` | complexType `CT_Picture` | `pict` (class `Pict`) |
   | `wml/wml.xsd` | complexType `CT_Drawing` | `drawing` (class `Drawing`) |
   | `wml/wml.xsd` | complexType `CT_FldChar` | `FldChar` |
   | `wml/wml.xsd` | complexType `CT_SectPrBase` | `SectPrBase` |
   | `wml/wml.xsd` | complexType `CT_SectPr` | `SectPr` |
   | `wml/wml.xsd` | complexType `CT_PermStart` | `RangePermissionStart` |
   | `wml/wml.xsd` | complexType `CT_Text` | `Text` |
   | `wml/wml.xsd` | element `shadow` | `shadow2006` |
   | `wml/wml.xsd` | complexType `CT_RPr` | `RPr` |
   | `wml/wml.xsd` | complexType `CT_ParaRPr` | `ParaRPr` |
   | `wml/wml.xsd` | complexType `CT_SdtPr` | `SdtPr` |
   | `wml/wml.xsd` | complexType `CT_SdtContentBlock` | `SdtContentBlock` |
   | `wml/wml.xsd` | element `sdt` | `SdtBlock` |
   | `wml/wml.xsd` | complexType `CT_SdtBlock` | `SdtBlock` |
   | `wml/wml.xsd` | complexType `CT_SdtRun` | `SdtRun` |
   | `wml/wml.xsd` | complexType `CT_TblWidth` | `TblWidth` |
   | `wml/wml.xsd` | complexType `CT_TblGridCol` | `TblGridCol` |
   | `wml/wml.xsd` | complexType `CT_TblGridBase` | `TblGridBase` |
   | `wml/wml.xsd` | complexType `CT_TblGrid` | `TblGrid` |
   | `wml/wml.xsd` | complexType `CT_TcMar` | `TcMar` |
   | `wml/wml.xsd` | complexType `CT_TcPr` | `TcPr` |
   | `wml/wml.xsd` | complexType `CT_TcPrInner` | `TcPrInner` |
   | `wml/wml.xsd` | complexType `CT_Tc` | `Tc` |
   | `wml/wml.xsd` | complexType `CT_TrPr` | `TrPr` |
   | `wml/wml.xsd` | complexType `CT_TblBorders` | `TblBorders` |
   | `wml/wml.xsd` | complexType `CT_TblPr` | `TblPr` |
   | `wml/wml.xsd` | complexType `CT_Tbl` | `Tbl` |
   | `wml/wml.xsd` | complexType `CT_NumFmt` | `NumFmt` |
   | `wml/wml.xsd` | complexType `CT_Lvl` | `Lvl` |
   | `wml/wml.xsd` | complexType `CT_Panose` | `FontPanose` |
   | `wml/wml.xsd` | complexType `CT_FontFamily` | `FontFamily` |
   | `wml/wml.xsd` | complexType `CT_Pitch` | `FontPitch` |
   | `wml/wml.xsd` | complexType `CT_FontSig` | `FontSig` |
   | `wml/wml.xsd` | complexType `CT_FontRel` | `FontRel` |
   | `wml/wml.xsd` | element `ins` | `RunIns` |
   | `wml/wml.xsd` | element `del` | `RunDel` |
   | `wml/wml.xsd` | complexType `CT_Body` | `Body` |
   | `wml/wml.xsd` | complexType `CT_PPr` (2nd) | `PPr` |
   | `wml/wml.xsd` | complexType `CT_ParaRPr` (2nd) | `ParaRPr` |
   | `wml/w14_word_2010_wordml.xsd` | element `shadow` | `shadow14` |
   | `shared/shared-math-2ed.xsd` | element `rPr` (×2) | `rPrMath` |
   | `shared/shared-math-2ed.xsd` | element `t` (×2) | `tMath` |
   | `dml/dml-baseStylesheet.xsd` | complexType `CT_FontCollection` | `FontCollection` |
   | `dml/dml-baseStylesheet.xsd` | complexType `CT_BaseStyles` | `BaseStyles` |
   | `dml/dml-graphicalObject.xsd` | complexType `CT_GraphicalObjectData` | `graphicData` (class `GraphicData`) |
   | `dml/dml-graphicalObject.xsd` | complexType `CT_GraphicalObject` | `graphic` (class `Graphic`) |
   | `dml/dml-picture.xsd` | complexType `CT_Picture` | `pic` (class `Pic`) |
   | `dml/dml-textCharacter.xsd` | complexType `CT_TextFont` | `TextFont` |
   | `dml/dml-wordprocessingDrawing.xsd` | complexType `CT_Inline` | `inline` (class `Inline`) |
   | `dml/dml-wordprocessingDrawing.xsd` | complexType `CT_Anchor` | `anchor` (class `Anchor`) |

   Note the lowercase `jaxb:class` names (`pict`, `drawing`, `pic`, `graphic`, `inline`,
   `anchor`): XJC capitalises them, so the actual classes are `Pict`, `Drawing`, `Pic`, `Graphic`,
   `Inline`, `Anchor`.

3. **No schema type at all.** `w:p` is a *global element with an anonymous complexType* in
   `wml.xsd` (line 9369), so there is no `CT_P` and no `types` entry; XJC names the class after
   the element, giving `@XmlType(name = "")` + `@XmlRootElement(name = "p")` on class `P`. The
   same holds for `styles`, `document`, `delText`, and 27 other wml classes. They are reachable
   only through `elements`.

4. **The schema itself was edited.** docx4j renamed some schema components in its XSD copies, not
   via JAXB customisation. In particular `wml.xsd:2608` reads
   `<xsd:simpleType name="JcEnumeration"> <!-- was "ST_Jc" -->`, so **there is no `ST_Jc` in
   `wml.json`** — the key is `JcEnumeration` (value `JcEnumeration`). `ST_Jc` exists only in
   `math.json` (→ `STJc`). Likewise `wml.xsd:9611` notes `CT_TcPrInner` replaced `CT_TcPrBase`.
   Because `schemas/` is a copy of the same edited XSDs, the fork sees the same names, so this is
   only a hazard for anyone diffing against the stock ECMA-376 schemas.

## Anomalies

### Element collisions

Six element names map to more than one class. Before an entry is given up on, the candidates that
the `types` section **already names** are dropped: the generator reaches those classes by their
schema type name and needs no element entry for them. Four of the six fall to that rule, two do
not.

| file | element | resolution |
| --- | --- | --- |
| `wml.json` | `sdt` | **`SdtBlock`** -- the other three candidates are named by `CT_SdtRun`, `CT_SdtCell` and `CT_SdtRow`; `SdtBlock` is the anonymous type of the global `<xsd:element name="sdt">`, which nothing else names |
| `wml.json` | `sdtContent` | no entry needed; all four candidates are named by their types |
| `wml.json` | `customXml` | no entry needed; all four candidates are named by their types |
| `drawing_2014_chartex.json` | `lvl` | no entry needed; both candidates are named by their types |
| `math.json` | `rPr` | **unresolved**: `CTR.RPrMath` and `CTMathRunTrackChange.RPrMath` are inner Java classes, not expressible as a Python class name |
| `math.json` | `t` | **unresolved**: `CTR.TMath` and `CTMathRunTrackChange.TMath`, likewise |

Where a genuine ambiguity remains and the candidates *are* expressible, the script writes **scoped
keys** instead, `element@EnclosingType`, from JAXB's `@XmlElementDecl(scope=...)` translated back
through the `types` table into schema type names. The fork's `<ClassNames>` option understands
them (fork `docs/codegen/config.md`). On docx4j's WML closure no such case survives the narrowing
rule, so no scoped key is emitted today; the mechanism is exercised by
`codegen/names/wml_overrides.json`, which is described below.

The full candidate lists, before narrowing:

| file | element | source | candidate classes (scopes) |
| --- | --- | --- | --- |
| `wml.json` | `sdt` | `@XmlRootElement` | `CTSdtCell`, `CTSdtRow`, `SdtBlock`, `SdtRun` |
| `wml.json` | `sdtContent` | `@XmlRootElement` | `CTSdtContentCell`, `CTSdtContentRow`, `CTSdtContentRun`, `SdtContentBlock` |
| `wml.json` | `customXml` | `ObjectFactory` | `CTCustomXmlBlock` (scopes `Body`, `CTCustomXmlBlock`, `CTFtnEdn`, `CTTxbxContent`, `Comments.Comment`, `Ftr`, `Hdr`, `SdtContentBlock`, `Tc`); `CTCustomXmlCell` (`CTCustomXmlCell`, `CTSdtContentCell`, `Tr`); `CTCustomXmlRow` (`CTCustomXmlRow`, `CTSdtContentRow`, `Tbl`); `CTCustomXmlRun` (`CTCustomXmlRun`, `CTSdtContentRun`, `CTSimpleField`, `CTSmartTagRun`, `P`, `P.Bdo`, `P.Dir`, `P.Hyperlink`, `RunDel`, `RunIns`, `RunTrackChange`) |
| `math.json` | `rPr` | `ObjectFactory` | `CTR.RPrMath` (scope `CTR`, from `org.docx4j.math`); `CTMathRunTrackChange.RPrMath` (scope `CTMathRunTrackChange`, from `org.docx4j.wml`) |
| `math.json` | `t` | `ObjectFactory` | `CTR.TMath` (scope `CTR`, from `org.docx4j.math`); `CTMathRunTrackChange.TMath` (scope `CTMathRunTrackChange`, from `org.docx4j.wml`) |
| `drawing_2014_chartex.json` | `lvl` | `ObjectFactory` | `CTNumericLevel` (scope `CTNumericDimension`); `CTStringLevel` (scope `CTStringDimension`) |

`w:ins` / `w:del` are **not** collisions: run-level `ins`/`del` win via `@XmlRootElement` on
`RunIns` / `RunDel`, and the cell/row-level variants (`CTTrPrChange` etc.) are named types rather
than competing root elements. Likewise `w:t` resolves cleanly to `Text` via `@XmlRootElement`.

**Type-name collisions: none.** No two Java classes in one package claim the same `@XmlType` name.

### Values that name a nested Java class

33 element entries (all in `wml.json`) resolve to an **inner** class, so the value contains a dot
and is *not* a valid Python class name. The generator must flatten or otherwise handle these:

`alias` → `SdtPr.Alias`, `annotationRef` → `R.AnnotationRef`, `bdo` → `P.Bdo`,
`bibliography` → `SdtPr.Bibliography`, `citation` → `SdtPr.Citation`,
`commentReference` → `R.CommentReference`, `continuationSeparator` → `R.ContinuationSeparator`,
`cr` → `R.Cr`, `dayLong` → `R.DayLong`, `dayShort` → `R.DayShort`, `dir` → `P.Dir`,
`divId` → `CTTrPrBase.DivId`, `endnoteRef` → `R.EndnoteRef`, `equation` → `SdtPr.Equation`,
`footnoteRef` → `R.FootnoteRef`, `gridAfter` → `CTTrPrBase.GridAfter`,
`gridBefore` → `CTTrPrBase.GridBefore`, `group` → `SdtPr.Group`, `hyperlink` → `P.Hyperlink`,
`lastRenderedPageBreak` → `R.LastRenderedPageBreak`, `monthLong` → `R.MonthLong`,
`monthShort` → `R.MonthShort`, `noBreakHyphen` → `R.NoBreakHyphen`, `pgNum` → `R.PgNum`,
`picture` → `SdtPr.Picture`, `ptab` → `R.Ptab`, `richText` → `SdtPr.RichText`,
`separator` → `R.Separator`, `softHyphen` → `R.SoftHyphen`, `sym` → `R.Sym`, `tab` → `R.Tab`,
`yearLong` → `R.YearLong`, `yearShort` → `R.YearShort`.

### Anonymous types (`@XmlType(name = "")`)

Classes XJC created for a global element rather than a named schema type. They appear in
`elements` only, never in `types`.

| file | count |
| --- | ---: |
| `wml.json` | 30 |
| `customxml.json` | 1 |
| `dml.json` | 1 |
| `mce.json` | 1 |

No other emitted namespace has any, and no class anywhere lacks a top-level `@XmlType`.

### Classes owned by another Java package

37 element entries name a class that lives in a *different* Java package from the one that owns
the namespace — docx4j reuses a class across namespaces rather than generating a near-duplicate.
The `java_package` field of the file is therefore the package of *most* of its classes, not all.
Examples: `w15.json`'s `collapsed`/`color`/`dataBinding` → `BooleanDefaultTrue`/`CTColor`/
`CTDataBinding` from `org.docx4j.wml`; `w14.json`'s `customXmlConflictDelRangeStart` →
`CTTrackChange` from `org.docx4j.wml`; `drawing_2012_chart.json`'s `layout`/`numFmt`/`tx` →
`CTLayout`/`CTNumFmt`/`CTTx` from `org.docx4j.dml.chart`; `drawing_2010_main.json`'s
`hiddenFill`/`hiddenLine`/… → `CTFillProperties`/`CTLineProperties`/… from `org.docx4j.dml`;
`dml_diagram.json`'s `t` → `CTTextBody` from `org.docx4j.dml`; `dml_chart.json`'s `userShapes` →
`CTDrawing` from `org.docx4j.dml.chartDrawing`. Run the script with `--report` for the full list.

### Java class names that collide across namespaces (informational)

96 class names are used in more than one namespace — `CTColor` in `dml`, `w14`, `w15` and `wml`;
`CTBoolean` in `dml_chart`, `drawing_2010_diagram`, `drawing_2012_chart`, `drawing_2014_chart`;
`CTChart`/`CTChartSpace` in `dml_chart` and `drawing_2014_chartex`; `BooleanDefaultTrue` in `w15`
and `wml`; and so on. In Java these are distinct classes in distinct packages. If the fork puts
its generated modules in one flat Python namespace they *will* clash — keep one Python module per
XML namespace. Some of these are the same class shared across namespaces (previous section) and
some are genuinely different classes with the same name; the two cases cannot be distinguished
from the class name alone.

### One class, many keys within a namespace

`CT_PPr` → `PPr` *and* `pPr` → `PPr` is the normal shape (one type + its element) and occurs for
220 classes. Beyond that, 28 classes are the target of several **element** names in one namespace
— the generator cannot use the element name to pick a class name for these:

* `wml`: `BooleanDefaultTrue` (27 elements: `b`, `bCs`, `caps`, …), `Text` (`t`, `instrText`,
  `delInstrText`), `CTMarkup` (4), `CTTrackChange` (4), `HpsMeasure` (3), `TblWidth` (3),
  `CTFtnEdnRef` (2), `CTMacroName` (2), `CTMoveBookmark` (2), `CTSdtDocPart` (2),
  `RunTrackChange` (2)
* `dml`: `CTPercentage` (15: `red`, `green`, `blue`, `lum`, `sat`, … and their `Mod`/`Off`
  variants), `CTPositiveFixedPercentage` (3), `CTPositivePercentage` (2)
* `mathml`: `OperatorType` (34: `abs`, `and`, `plus`, `times`, …), `ConstantType` (4),
  `QualifierType` (2)
* `w14`: `CTPercentage` (6), `CTOnOff` (3), `CTPositiveFixedPercentage` (3), `CTMarkup` (2),
  `CTTrackChange` (2); `w15`: `BooleanDefaultTrue` (4)
* `drawing_2012_chart`: `CTBoolean` (4); `drawing_2014_chart`: `CTBoolean` (2);
  `drawing_2014_chartex`: `CTFormula` (2); `drawing_2014_main`: `CTIdentifier` (2);
  `dml_spreadsheetdrawing`: `CTMarker` (`from`, `to`)

### Python hazards

* **Element names that are Python keywords or shadow builtins** — 12. They are fine as class-name
  *keys*, but anything deriving a field/attribute name from them needs escaping:
  `wml` `del` (keyword) → `RunDel`, `dir` → `P.Dir`, `id` → `Id`, `object` → `CTObject`;
  `dml_spreadsheetdrawing` `from` (keyword) → `CTMarker`;
  `mathml` `and`, `not`, `or` (keywords) and `abs`, `max`, `min`, `sum` (builtins) → `OperatorType`.
* **Element names that are not valid Python identifiers** — none. Every element name in every
  table matches `[A-Za-z_][A-Za-z0-9_]*`.
* **Class names that are Python keywords or shadow builtins** — none. Every value is PascalCase.
  (`Text`, `Id`, `Object`… differ from the builtins only by case.)
* **Type-name keys that are not valid Python identifiers** — 40, all in `inkml.json`, of the form
  `ink.type`, `trace.type`, `annotationXML.type`, … (XSD type names containing a dot). They are
  keys, not identifiers, so this only matters if a consumer assumes keys are identifier-shaped.
* **Duplicate class names within one namespace** — none in `types` (a class name is a file name,
  which is unique per package); see the previous section for one class under several element keys.

### Skipped Java packages

Packages with `.java` files but no `package-info.java`, so no namespace can be read. All six are
in the WML closure or adjacent to it, so the namespaces they would have served get no table:

| java package | files | namespace it would have served |
| --- | ---: | --- |
| `org.docx4j.sharedtypes` | 7 | `…/officeDocument/2006/sharedTypes` (the `@XmlType` annotations do carry it) |
| `org.docx4j.dml.lockedCanvas` | 1 | `…/drawingml/2006/lockedCanvas` |
| `org.docx4j.com.microsoft.schemas.office.drawing.x2010.picture` | 1 | `…/office/drawing/2010/picture` |
| `org.docx4j.com.microsoft.schemas.office.drawing.x2014.chart.ac` | 1 | `…/office/drawing/2014/chart/ac` |
| `org.docx4j.com.microsoft.schemas.office.drawing.x2016.ink` | 3 | `…/office/drawing/2016/ink` |
| `org.docx4j.com.microsoft.schemas.office.drawing.x201612.diagram` | 1 | `…/office/drawing/2016/12/diagram` |

No package declared an empty `@XmlSchema` namespace, and no namespace is claimed by two packages.

## Namespaces with no table

### In the `schemas/` closure but with no usable docx4j package (8)

* `http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes` — `org.docx4j.sharedtypes`
  exists (7 enums, `ST_OnOff`, `ST_CalendarType`, …) but has no `package-info.java`; the namespace
  lives on the individual `@XmlType(namespace = …)` annotations instead.
* `http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas` — package exists, no `package-info.java`.
* `http://schemas.microsoft.com/office/drawing/2010/picture`,
  `http://schemas.microsoft.com/office/drawing/2014/chart/ac`,
  `http://schemas.microsoft.com/office/drawing/2016/ink`,
  `http://schemas.microsoft.com/office/drawing/2016/12/diagram` — ditto.
* `http://schemas.openxmlformats.org/officeDocument/2006/relationships` — attribute-only
  (`r:id`, `r:embed`); docx4j generates no classes for it. Note `org.docx4j.relationships` is the
  *package* relationships namespace (`…/package/2006/relationships`), a different thing.
* `http://www.w3.org/XML/1998/namespace` — `xml:space`, `xml:lang`; no classes.

### docx4j packages whose namespace is outside the closure (not emitted, 24 under `org.docx4j`)

All VML — `urn:schemas-microsoft-com:vml` (`org.docx4j.vml`), `…:office:word`
(`org.docx4j.vml.wordprocessingDrawing`), `…:office:office` (`org.docx4j.vml.officedrawing`),
`…:office:excel`, `…:office:powerpoint`, `urn:docx4j:vml:root` — plus the package-level parts:
`…/package/2006/relationships` (`org.docx4j.relationships`),
`…/package/2006/metadata/core-properties` (`org.docx4j.docProps.core`),
`http://purl.org/dc/elements/1.1/`, `http://purl.org/dc/terms/`,
`…/officeDocument/2006/extended-properties`, `…/officeDocument/2006/custom-properties`,
`…/officeDocument/2006/docPropsVTypes`, `…/officeDocument/2006/bibliography`,
`…/officeDocument/2006/customXml`, `…/office/2006/xmlPackage`, `…/office/2006/coverPageProps`,
`…/office/2006/encryption` (+ the two `keyEncryptor` namespaces),
`…/office/webextensions/{taskpanes,webextension}/2010/11`,
`…/office/word/2015/wordml/symex` (`org.docx4j.w15symex`), and
`…/office/word/2016/wordml/cid` (`org.docx4j.w16cid`).
The `org.xlsx4j.*` and `org.pptx4j.*` trees are likewise outside the closure except
`org.xlsx4j.schemas.microsoft.com.office.excel.x2010.spreadsheetDrawing`, which is in it and *is*
emitted.

**`vml` and `w16cid` in particular are worth a second look**: `w:pict` content is VML, and
`w16cid` is common in modern `.docx` files, but neither namespace is reachable from the `.xsd`
files currently in `schemas/`. If either is added to `schemas/`, re-running the script picks it up
automatically — no change to the script is needed.

### ObjectFactory declarations dropped because their namespace has no table

255 `@XmlElementDecl`s point at namespaces outside the emitted set and were discarded. The largest
groups: `…/spreadsheetml/2006/main` (52), `…/office/powerpoint/2010/main` (30),
`urn:schemas-microsoft-com:vml` (21), `…/officeDocument/2006/bibliography` (19),
`urn:schemas-microsoft-com:office:office` (18), `http://purl.org/dc/elements/1.1/` (16),
`…/office/powerpoint/2013/main/command` (8), `…/officeDocument/2006/docPropsVTypes` (7),
`…/office/powerpoint/2012/main` (6), `urn:schemas-microsoft-com:office:word` (6),
`…/presentationml/2006/main` (4), and 21 namespaces with 1–2 each.

### Declarations whose payload is a plain Java type, not a docx4j class

147 `@XmlElementDecl`s produce a `JAXBElement<…>` over a built-in type and are skipped, since
there is no docx4j class name to record: `String` (106), `BigInteger` (18), `byte[]` (6),
`Integer` (5), `Long` (3), `Short` (2), `XMLGregorianCalendar` (2), and one each of `BigDecimal`,
`Boolean`, `Byte`, `Double`, `Float`.

## The two hand written entries: `wml_overrides.json`

`derive_names.py` writes one `<prefix>.json` per namespace and never touches
`wml_overrides.json`, which carries three **scoped** element entries and, since CR-002 Phase A,
one type entry:

```json
"types":    { "CT_StylePaneFilter": "CTStylePaneFilter" },
"elements": { "t@r": "RT", "instrText@r": "RInstrText", "delInstrText@r": "RDelInstrText" }
```

`CT_StylePaneFilter` is there because the type **does not exist in docx4j's schema copy at all**:
`schemas/PATCHES.md` 3 adds it, so `derive_names.py` has no docx4j name to derive and xsdata would
call it `CtStylePaneFilter`. It is named the way docx4j names every `CT_` type it did not rename,
`CTShortHexNumber` and `CTSignedTwipsMeasure` being its neighbours.

`w:t`, `w:instrText` and `w:delInstrText` inside `w:r` all have the type `CT_Text`, so the
generator invents a class per element name to keep a compound field able to tell them apart.
docx4j has no class for them either -- JAXB used one `JAXBElement` per name over the shared
`Text` -- so there is no derived name to take, and left alone they come out `T2`, `InstrText2`
and `DelInstrText2`. `w:t` is the most used element in the model. The names chosen follow what
the generator already does for the *siblings* of the same field, `R.Cr` as `RCr` and `R.Tab` as
`RTab`, which is how docx4j names them in Java as well. All three are subclasses of `Text`.

The scope is the enclosing class as the generator holds it at the moment it invents the class:
the schema type name (`CT_P`), unless the type is used by exactly one global element and has been
reduced into it, which is the case here -- `CT_R` is `r`. It is read in the file's own namespace,
so it is wml's `w:r` and not OMML's `CT_R`, which also carries a `w:t`.

## Notes for the integration stage

1. `types` keys are the **type names as they appear in `schemas/`**, which are docx4j's edited
   copies of the OOXML XSDs. Feed the generator the same `schemas/` tree or the keys will not match
   (`JcEnumeration` vs `ST_Jc` is the sharpest example).
2. A class reachable only as an anonymous type (`P`, `Styles`, `Document`, `DelText`, …) has **no
   `types` entry**; the generator must consult `elements` for global elements with inline
   complexTypes, not just `types`.
3. 33 `elements` values contain a dot (inner Java classes). Decide on a flattening rule before
   generating.
4. Four of the six element collisions are resolved by the narrowing rule above and two
   (`m:rPr`, `m:t`) are not. Do not "fix" the remaining two by hand-picking a candidate in a
   derived file — put the choice in `wml_overrides.json` or another file the script does not
   write, so that re-running `derive_names.py` cannot silently revert it.
5. The same class name appears in several namespaces (96 cases); keep generated modules per
   namespace.
