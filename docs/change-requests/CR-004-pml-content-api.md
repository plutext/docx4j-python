# CR-004: A PresentationML content API, and a python-pptx facade

**Status:** Proposed 2026-09-19. Nothing implemented.
**Depends on:** CR-001 Phase D (the PML, DML chart and DML diagram namespaces generated, lazy
per-namespace import) and CR-002 Phase C (`PresentationMLPackage` with typed slide, layout,
master, notes and chart parts), both proposed; CR-003 (the content API for WordprocessingML,
implemented), whose conventions, agent surface and facade pattern this CR reuses rather than
restates. For the effort figures, CR-005 and the xsdata fork's CR-001 (the serialiser) are
independent of this one.
**Counterpart:** python-pptx 1.0.2 (the facade's oracle: `scripts/python_pptx_subset.py` derives
the member list from an installed copy, as `scripts/python_docx_subset.py` does for python-docx);
the PowerPoint JavaScript API (`PowerPoint.Presentation`, `Slide`, `Shape`, `TextFrame`,
`TextRange`, `Table`, `SlideLayout`, `SlideMaster`) for the vocabulary; docx4j's `pptx4j`
(`org.pptx4j.pml`, `PresentationMLPackage`, `SlidePart`, `Pptx4jException`) and
`docx4j-samples-pptx4j` for the semantics; the portfolio's `docs/python_ecosystem.md` for why.

## 1. Why

The pptx ecosystem in Python has one library, python-pptx: 3.5k stars, 451 open issues, 86 open
pull requests, and a design (lxml proxies over the slice of the schema it models) whose users hit
the schema wall first, on charts, SmartArt, animations and most of DrawingML. Nothing in Python
offers a lossless edit of a deck it does not fully model, and nothing offers a deck to an agent
the way CR-003 offers a document: an outline under a budget, addresses that survive edits, a
change report on every call, a dry run. CR-002 Phase C gives this engine the *package* (load,
save, untouched parts byte for byte, the whole schema typed); this CR gives it the content API
and the familiar names on top, which is what turns a package into something people use.

The pattern is CR-003's, and every rule of that CR that is not about WordprocessingML applies
here unchanged: Python conventions (section 3.1 there), the agent surface (3.4), the facade as a
separate namespace with the incumbent's names and semantics (3.6, section 20), the derived
subset list and the parity test with the incumbent as the judge, `ChangeReport` on every
mutation, determinism, `dry_run` over deep copies. This document says what is different because
the format is different.

## 2. What others do

| | reads | writes | what it cannot |
|---|---|---|---|
| python-pptx 1.0.2 | slides, layouts, masters, shapes, placeholders, text with runs and fonts, tables, pictures, charts (category, XY, bubble), notes | the same, `add_slide` from a layout, `add_textbox`, `add_picture`, `add_table`, `add_chart` with an embedded workbook it writes itself | animations, transitions, SmartArt, sections, comments, most DrawingML effects; anything it does not model is kept only because lxml keeps it, and a rewritten part is re-serialised |
| docx4j `pptx4j` (Java) | the whole schema typed | the whole schema | no content API to speak of: `SlidePart`, `MainPresentationPart.addSlide`, samples |
| PowerPoint JS | `Presentation.slides`, `Slide.shapes`, `Shape` geometry, `TextFrame.textRange`, `Table`, layouts and masters, tags, custom properties, `insertSlidesFromBase64` | the same; no charts at all (2026) | it is the vocabulary an add-in author already knows, and the one CR-003 chose for Word |

python-pptx's member count, measured by introspection on 2026-09-19 over the thirty classes its
users touch: **302 public members**. That is the size of the facade's promise, before the
omissions.

## 3. Design

### 3.1 Conventions

CR-003 section 3.1, verbatim: `snake_case` members with PowerPoint JS's string values and a
`StrEnum` beside each; defaults so the common call is short; collections are `list`s and views
are values over a live element; frozen result dataclasses with `to_dict()`; the error hierarchy
of `docx4j_py.model.content.errors` with `code` and `hint`; measurements in **points** in the
content API (PowerPoint JS reports points), EMU `Length` only in the facade (python-pptx's
unit).

### 3.2 The surface (`docx4j_py/model/presentation/`)

```python
class Presentation:               # PowerPoint.Presentation, over a PresentationMLPackage
    slides: list[Slide];  slide_masters: list[SlideMaster];  layouts: list[SlideLayout]
    slide_width: float;  slide_height: float                # points; p:sldSz
    title: str                                              # docProps/core.xml
    def insert_slide(self, *, layout: SlideLayout | str | None = None, location: Literal["Start", "End"] | int = "End") -> Slide
    def get_slide(self, address: str) -> Slide
    def outline(self, *, max_chars=80, budget=...) -> Outline;  def find(self, text, **options) -> list[SearchHit]
    def describe(self) -> dict;   def to_markdown(self, *, addresses=False, notes=False, max_chars=None) -> str
    def to_api_script(self, **options) -> str
    package: PresentationMLPackage;  last_change: ChangeReport | None

class Slide:                      # PowerPoint.Slide, over a SlidePart
    id: int;  index: int;  address: str                     # "slide:<p:sldId id>" first, "slides/3" second
    layout: SlideLayout;  master: SlideMaster
    shapes: list[Shape];  placeholders: list[Shape]         # document order (z-order); placeholders by type
    title: str | None                                       # the title placeholder's text
    notes: NotesSlide | None
    hidden: bool
    def insert_shape(self, kind: ShapeKind, *, left, top, width, height, text="") -> Shape   # text box, geometric shape
    def insert_picture(self, data: bytes, *, left, top, width=None, height=None, **options) -> Shape
    def insert_table(self, row_count, column_count, *, left, top, width, height, values=None) -> Shape
    def insert_chart(self, chart_type: ChartType, data: ChartData, *, left, top, width, height) -> Shape
    def insert_markdown(self, markdown: str) -> list[Shape]  # a title and a body from "# Title" + bullets
    def move(self, index: int) -> None;  def duplicate(self) -> Slide;  def delete(self) -> None
    def to_markdown(self, *, addresses=False) -> str;  def get_xml(self) -> str;  def to_dict(self) -> dict
    def get_comments(self) -> list[Comment]                 # modern (p188) and legacy comments, read; write in a later phase
    element: CTSlide;  part: SlidePart

class Shape:                      # PowerPoint.Shape, over p:sp / p:pic / p:graphicFrame / p:grpSp / p:cxnSp
    id: int;  name: str;  type: ShapeType                   # "GeometricShape" | "Image" | "Table" | "Chart" | "Group" | "Line" | "Placeholder" | "SmartArt" | "Unsupported"
    address: str                                            # "slide:257/shape:4" (p:cNvPr id), stable across edits
    left: float; top: float; width: float; height: float; rotation: float      # points; write moves the a:xfrm
    text_frame: TextFrame | None;  table: Table | None;  chart: Chart | None;  picture: Picture | None
    placeholder: Placeholder | None                         # p:nvPr/p:ph: type, idx, the layout shape it inherits from
    fill: Fill;  line: Line                                 # solid/none/gradient read, solid/none write; a:ln width and colour
    group: Shape | None;  shapes: list[Shape]               # a p:grpSp's children
    z_order: int;  def bring_forward / send_backward / to_front / to_back
    def delete(self) -> None;  def get_xml(self) -> str;  def to_dict(self) -> dict
    element: Any

class TextFrame:                  # PowerPoint.TextFrame over p:txBody
    text: str;  text_range: TextRange;  paragraphs: list[TextParagraph]
    auto_size: Literal["None", "ShapeToFitText", "TextToFitShape"];  word_wrap: bool
    vertical_alignment: Literal["Top", "Middle", "Bottom", ...];  margins (left/right/top/bottom: float)
    def insert_paragraph(self, text="", *, level=0, location="End") -> TextParagraph
    def clear(self) -> None

class TextRange:                  # PowerPoint.TextRange: a span within a text frame, or the whole
    text: str;  font: Font;  paragraph_format: ParagraphFormat       # alignment, indent level, bullet on/off
    def insert_text(self, text, *, location="End") -> TextRange
    def search(self, text, **options) -> list[TextRange];  def get_range(...)

class TextParagraph / TextRun    # a:p / a:r with a:rPr: the CR-003 Paragraph/Range split, DrawingML flavoured
class Font:                       # a:rPr: bold, italic, underline, size (points), name (latin), color (hex or scheme), strike, ...
class Table / TableRow / TableCell  # a:tbl: row_count, column_count, values, cell(r, c), merge reads, add_rows, column widths, cell text_frame
class Picture:                    # p:pic: image part, get_bytes(), crop, replace_image(data)
class Chart:                      # c:chart over the chart part: chart_type, categories, series (name, values), replace_data(ChartData), title
class Placeholder:                # type, idx, inherits: bool, layout_shape: Shape | None
class NotesSlide:                 # text_frame of the notes placeholder; to_markdown
class SlideLayout / SlideMaster   # name, shapes (read), placeholders
```

Semantics that are not CR-003's:

- **Geometry is the value.** A shape's `left`/`top`/`width`/`height` read the effective `a:xfrm`:
  the shape's own, else the placeholder's inherited from the layout, else the master's (python-pptx
  does the same walk). A write always writes the shape's own `a:xfrm`, as PowerPoint does when a
  placeholder is moved.
- **Text has two levels of formatting**, run (`a:rPr`) and paragraph (`a:pPr`: level, alignment,
  bullets through `a:buChar` / `a:buAutoNum` / `a:buNone`), and both inherit from the placeholder's
  list styles (`a:lstStyle`) and the master's text styles. Reads report the *direct* value, as
  CR-003's `Font` does until a resolver exists; a `PowerPointPropertyResolver` is this CR's
  analogue of CR-002 Phase B's, out of scope here and named in section 9.
- **A placeholder is a shape** with `placeholder` set, as in both python-pptx and PowerPoint JS;
  the slide's `placeholders` is the filtered list by `p:ph/@type` and `@idx`. `insert_slide` copies
  the layout's placeholders into the new slide with empty text, exactly as `add_slide` does.
- **Slides are parts.** Insert and delete are package operations: a new `SlidePart` under a free
  `/ppt/slides/slideN.xml`, its relationship from the presentation part, a `p:sldId` with a fresh
  id above the highest, the `[Content_Types]` override, and the layout relationship from the new
  slide. Delete removes the `p:sldId` and the relationship and leaves the part orphaned for the
  save to drop, as PowerPoint does. Notes slides come with their slide.
- **Charts** read the typed chart part (CR-001 Phase D's `c:` namespace): categories and series
  from `c:cat` / `c:val` caches. `replace_data` writes the caches and, when the chart has an
  embedded workbook, rewrites it through CR-005's SpreadsheetML package when that exists and
  otherwise leaves the workbook alone and says so in `ChangeReport.warnings` (PowerPoint
  re-derives the cache from the workbook only when a user opens the data). Chart *creation* is a
  later phase and is what needs the workbook writer.
- **SmartArt** (`dgm:` graphic data) reads as text through the data part and is otherwise
  `Unsupported`, with the whole `p:graphicFrame` available as XML. Animations (`p:timing`) and
  transitions are kept byte for byte and never touched.

### 3.3 The agent surface

CR-003 section 3.4 with slides as the block level:

- **Addresses:** `"slide:<id>"` and `"slide:<id>/shape:<id>"` first (`p:sldId/@id` and
  `p:cNvPr/@id`, both stable across inserts), ordinals `"slides/3/shapes/1"` second, text third
  (`presentation.shape_at(contains=...)`). Table cells extend a shape's address with `/r/c`.
- **`outline()`:** one entry per slide (address, title, layout name, shape count, notes length,
  hidden) with nested entries per shape (address, type, name, placeholder type, text truncated to
  `max_chars`) under the same `budget` rules and `truncated` flag as CR-003.
- **`find()`** across all text frames, tables and notes with snippets convertible to a `TextRange`.
- **`describe()`:** layouts and their placeholders, masters, theme fonts and colours, slide size,
  the chart types in use, the parts inventory --- what an agent needs before it inserts a slide.
- **`ChangeReport`** on every mutation, `parts_touched` naming slide, notes, chart, media and
  relationship parts; `dry_run()` over deep copies with the undo log for added parts (CR-003
  section 14.4); `DocumentSession` unchanged.
- **Markdown:** `to_markdown()` renders a deck as one document, `---` between slides, the title
  as a heading, body text as bullets with indent for level, tables as GFM, notes under a `>`
  quote when asked, pictures as `![name](slide:257/shape:4)`; every slide and shape carries its
  address as an HTML comment when `addresses=True`. `Slide.insert_markdown` fills a slide's title
  and body placeholders from `# Title` plus bullets (and a GFM table into a new table shape),
  which is what an agent writes when it is asked to "add a slide about X". Nothing beyond that
  is imported; CR-003 Phase K's importer is reused for the inline formatting.
- **`to_api_script`** as CR-003 Phase I, over slides and shapes, with `slide.insert_xml(...)` of
  the marshalled `p:sp` as the fallback.

### 3.4 The facade, `docx4j_py.pptx`

python-pptx's public API over the views, in python-pptx's names and semantics, as CR-003
section 3.6 / section 20 did for python-docx:

```python
from docx4j_py.pptx import Presentation, Inches, Pt
from docx4j_py.pptx.util import Emu
from docx4j_py.pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from docx4j_py.pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from docx4j_py.pptx.chart.data import CategoryChartData
from docx4j_py.pptx.enum.chart import XL_CHART_TYPE

prs = Presentation("in.pptx")                      # or Presentation() for a new deck
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "Quarterly report"
body = slide.placeholders[1].text_frame
body.text = "Revenue"
p = body.add_paragraph(); p.text = "up 12%"; p.level = 1
slide.shapes.add_picture("logo.png", Inches(1), Inches(1), width=Inches(2))
table = slide.shapes.add_table(2, 3, Inches(1), Inches(3), Inches(6), Inches(1)).table
table.cell(0, 0).text = "Region"
prs.save("out.pptx")
```

`Presentation`, `Slides`, `Slide`, `SlideLayout(s)`, `SlideMaster(s)`, `NotesSlide`,
`SlideShapes` (`add_textbox`, `add_picture`, `add_table`, `add_shape`, `add_chart`,
`add_group_shape`, `title`, `placeholders`), `Shape`, `Picture`, `GraphicFrame`,
`SlidePlaceholder`, `GroupShape`, `TextFrame`, `_Paragraph`, `_Run`, `Font`, `Table`, `_Cell`,
`_Row`, `_Column`, `Chart`, `CategoryChartData` / `XyChartData`, plots, series, axes,
`FillFormat`, `LineFormat`, `ColorFormat`, `Length` and its units, the `MSO_*` / `PP_*` / `XL_*`
enumerations with python-pptx's member names and integer values. The subset list
(`tests/python_pptx_subset.json`) is derived by introspecting python-pptx 1.0.2 in a throwaway
venv, with the hand-kept omissions and their reasons in the script; the parity test builds one
ordinary python-pptx script through both libraries and reads both files back **with
python-pptx**. python-pptx is never a dependency. The measured starting point is 302 members over
thirty classes; the omissions expected on the day are `add_chart` with an embedded workbook
until CR-005 exists (departure recorded, the chart is refused with a message), `add_movie`, and
python-pptx's `Image` parser (as `Drawing.image` was for docx).

## 4. Rules kept from CR-003 and its TypeScript counterpart

Everything in CR-003 section 4 that is not WordprocessingML-specific: a built-in name the
document lacks is defined, not dangled (the layout's placeholder analogue: a placeholder type the
layout lacks is refused with the layout's list); inserted content returns its views; nothing is
fetched from the network; an image is never fetched; the undo log for added parts; the wall clock
only in the docProps dates; ids from a seeded generator; every write through one primitive so
`ChangeReport` and `dry_run` cannot be bypassed. Two rules of its own:

- **Never rewrite a slide part a read touched only for its text.** `outline()`, `find()`,
  `describe()` and `to_markdown()` read slides with lxml, as CR-003's `describe()` does, so a
  deck read end to end saves byte for byte.
- **Ids are the file's.** `p:sldId/@id` and `p:cNvPr/@id` are reported as they stand and allocated
  above the highest in the part; nothing renumbers.

## 5. Where each piece lives

| Piece | Package |
|---|---|
| tree-layer builders for `p:sp`, `p:txBody`, `p:pic`, `p:graphicFrame` with `a:tbl`, the chart frame, `a:xfrm` | `docx4j_py.pml.builders`, `docx4j_py.dml.builders` (new, in `KEEP`) |
| `Presentation`, `Slide`, `Shape`, `TextFrame`, `TextRange`, `Table`, `Picture`, `Chart`, `Placeholder`, `NotesSlide`, layouts and masters | `docx4j_py.model.presentation` |
| addresses, `Outline`, `find`, `describe`, `ChangeReport`, `dry_run` | reused from `docx4j_py.model.content.addresses` / `.reports` / `.trial` with a slide-shaped container; the `DocumentSession` unchanged |
| markdown out and in | `docx4j_py.model.markdown.presentation` (over the views; the inline importer shared) |
| `to_api_script` | `docx4j_py.model.presentation.api_script` |
| the facade | `docx4j_py.pptx` (a separate namespace; python-pptx's module paths mirrored where code imports them: `pptx.util`, `pptx.enum.*`, `pptx.chart.data`) |
| the subset list and the parity test | `scripts/python_pptx_subset.py`, `tests/python_pptx_subset.json`, `tests/presentation/test_python_pptx_parity.py` |

## 6. What CR-001 Phase D and CR-002 Phase C supply, and what this CR does without

CR-001 D: the `pml`, `dml.chart`, `dml.diagram` namespace packages with docx4j's class names
(`CTSlide`, `CTShape`, `CTTextBody`, `CTGraphicalObjectFrame`, `CTChartSpace`) and `el` factories,
and the lazy per-namespace import that keeps `import docx4j_py.docx` from paying for PML. CR-002
C: `PresentationMLPackage` with `MainPresentationPart`, `SlidePart`, `SlideLayoutPart`,
`SlideMasterPart`, `NotesSlidePart`, `ChartPart`, `EmbeddedPackagePart`, `ThemePart`, their
registry entries and `create_package()` with docx4j's default deck (one blank layout set, the
theme, 4:3 or 16:9). Without a property resolver, `Font` and `ParagraphFormat` read direct
values and say so (section 3.2); without CR-005, chart creation is refused (section 3.4).

## 7. Tests

- **Fixtures:** docx4j's `loadAndSave.pptx` (already in `samples/`), the decks under
  `../docx4j/docx4j-samples-pptx4j/sample-docs/` (charts, tables, pictures, notes, a 16:9 deck),
  one deck saved by a current PowerPoint with a SmartArt, an animation and a modern comment.
- **Per phase**, in `tests/presentation/`: reads against known values; every mutation followed by
  save, reload and read back; untouched parts byte for byte through the content API; a deck read
  end to end saves byte for byte.
- **Agent scenarios**, in `tests/agent/`: outline, find, insert a slide by markdown, edit by
  address, `ChangeReport` checked, `dry_run` then commit; the budget test for a 200-slide deck.
- **The subset and parity tests** (section 3.4); `to_api_script`'s executed round trip.
- **PowerPoint acceptance**, manual, per phase in `tests/README.md`: the saved deck opens without
  a repair prompt; a slide inserted from a layout, a picture, a table, a chart with replaced data,
  a notes page; then save from PowerPoint, close, reopen.

## 8. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | Tree-layer builders (`sp`, `txBody`, `pic`, `graphicFrame`/`tbl`, `xfrm`), `el` coverage, the DrawingML text model (paragraph and run split, `a:rPr` to `Font`) | 2 days |
| B | `Presentation`, `Slide`, `Shape`, `TextFrame`, `TextRange`, `Font`, `Placeholder`: reads, geometry, text writes, `insert_shape`, `insert_slide` and `delete` as package operations, layouts and masters read | 4 days |
| C | `Table`, `Picture`, `Chart` reads and `replace_data`, groups and z-order, `NotesSlide`, comments read | 3 days |
| D | The agent surface: addresses, `outline()`, `find()`, `describe()`, `ChangeReport`, `dry_run`; `to_markdown` and `Slide.insert_markdown`; `to_api_script` | 4 days |
| E | The python-pptx facade, its subset list and parity test; the acceptance deck | 4 days |
| F | Chart creation with the embedded workbook (needs CR-005), modern comment writes, a `PowerPointPropertyResolver` for effective text formatting | 3 days, later |

Nineteen days for A to E; F when its dependencies land. Each phase ends with implementation
notes in this CR, the status line and the index updated, the README's example executed.

**Hand-off from docx4j CR-035 phase 2 (2026-10-12, docx4j `276325167` on `VERSION_17_3_3`,
unpushed; its CR section 4.2 the design, section 11 the record).** The chart read model exists in
Java and has a JSON oracle, which Phase C's `Chart` reads and Phase F's creation are measured
against the way CR-002 Phase B is measured against the parity goldens: consumed from the sibling
checkout, nothing copied. `org.docx4j.model.chart`: `Charts.in(pkg)`, every chart a docx, pptx or
xlsx places, in document order, as a `ChartRef` (host part, anchor, extent, name, the `c:` or
`cx:` chart part, whether it sat in an `mc:AlternateContent` Choice, as every chartex waterfall
does); `ChartModel.read(chart, hostTheme)`, groups, series, axes, title, legend and chart-wide
settings, the series through one adapter over the eight `ser` classes, series hidden by Excel's
chart filter (`c15:filtered*Series`) included and marked hidden, data keyed by `c:pt/@idx` with
`ptCount` (a gap is null, `#N/A` kept as text), fills automatic where unstated and solid colours
resolved in the host's theme through CR-034 (here CR-009 Phase C); `ChartExModel` from the chartex
caches; `ChartCells`, the data from the cells and where it differs from the caches. The **shown**
values follow MS-OI29500 and are part of the oracle: a title with no text shows "Chart Title" or
the name of a single named series (2.1.1561); a series with no name is "SeriesN" with N = idx + 1
(2.1.1565); `varyColors` is ignored with more than one series and on area, stock and filled radar
(2.1.1571); a pie shows only its first series (2.1.1519); data-label positions have per-kind
defaults (2.1.1456); each kind has its own supported `dispBlanksAs` values and default (2.1.1452).
The oracle: `../docx4j/docx4j-core-tests/src/test/resources/charts/json/`, 32 goldens (29 charts,
3 chartex), one per chart part of every fixture in `charts/` (its `README.md`), format version 1:
fixed member order, two-space indent, cache numbers as written, gaps as null. The fixtures include
16 Excel-made probes (`p2-*`: every chart kind, a combination chart on a secondary axis, a
chartsheet, gaps, multi-level categories, a filtered series and category, a chart pasted into
PowerPoint). A measured fact for `replace_data` and for creation: Excel hides a filtered category
by narrowing every series' `c:f` to a union of the visible cells, such as
`(Sheet1!$A$2:$A$3,Sheet1!$A$5)`, the cache holding only those points and the full range in
`c15:fullRef`; it does not use `categoryFilterExceptions`. CR-005 Phase G shares this oracle.

**Hand-off from docx4j CR-035 phase 4 (2026-10-12, docx4j `b7a811bbf` on `VERSION_17_3_3`,
unpushed; CR-035 is done; its CR sections 4.4 the design, 8 decisions D10 to D12, 11 every
measured rule).** A chart's data has a Java editor gated on Excel's own edits, which is the
oracle for `replace_data` (Phase C, once the embedded workbook can be written: question 3) and for
Phase F's creation. `org.docx4j.model.chart.ChartDataEditor(chart, workbook)`, or `of(chart)` for
an xlsx: `setValue`, `setBlank`, `setError`, `setCategory`, `setSeriesName`, `addCategory`,
`removeCategory`, `addSeries`, `removeSeries`, `setData`, `refreshCaches`, each keeping the cells,
the table, the chart's references and its caches in step, for series in columns or rows, Excel's
chart-filter unions and multi-level categories; under it limited cell writers in
`org.xlsx4j.model` (`CellRef`, `AreaRef`, `SheetCells`, `Tables`). The oracle:
`../docx4j/docx4j-core-tests/src/test/resources/charts/edit/`, 21 cases (`e1` to `e7`, `r1` to
`r14`), each a before and an after workbook made by Excel and a note of the `=SERIES` formulas
and table ranges, held by `ChartDataEditorTest` (27 tests: chart part, cells, spans, dimension and
table compared part by part with Excel's after-workbook); every docx4j output opened in Excel
without repair, with the same Select Data ranges. The rules Excel was measured to follow:

- **A new series**: `idx` and `order` one above every series', filtered ones included;
  `c16:uniqueId` the `idx` as 8 hex digits plus the chart's shared suffix; no data labels (r10);
  its formatting from the last series still in its automatic colour, recoloured by the colour
  style's cycle on fill, line and marker (r9, r11).
- **The table**: `ref` and `autoFilter` grow; a new `tableColumn` takes id max + 1, a fresh
  `xr3:uid` and its header cell's text as its name; a duplicate name becomes "Series 12" in the
  cell and the table (r7). A new cell takes the style of the cell before it (r8b); row spans are
  the extent of each 16-row block.
- **Caches**: `ptCount` is the number of cells named, one `c:pt` per non-empty cell (a blank point
  dropped, `ptCount` kept); an error cached as `#N/A`; a cache's `formatCode` kept as written
  (Excel spells built-in formats its own way, `d\-mmm`). Numbers compared as numbers: Excel writes
  2.2 as `2.2000000000000002`.
- **Excel's chart filter**: the unions and `c15:fullRef` grow, the filtered series' too (r12, r13).
- **D12**: deleting a middle table column, Excel keeps series by position (r1, r5);
  `removeSeries` removes the named series with its formatting. Removing series from the end keeps
  formatting by position, as python-pptx's `replace_data` does and docx4j's `setData` does.
- **D10**: formula cells are not overwritten unless asked; `fullCalcOnLoad` is set when the
  workbook has formulas.

docx4j's `326b88bd2` (CR-035 phase 3) also gives `ChartTable`, a chart as its title over a
categories-down, series-across table with numbers formatted by format code in the US locale, as
docx4j's markdown export writes it: the shape for a chart in `to_markdown` and `describe()`.

## 9. Open questions (recommendations for decision)

1. **Vocabulary: PowerPoint JS or python-pptx as the primary API.** Recommendation: PowerPoint
   JS for the content API (consistency with CR-003 and the TypeScript engine; an add-in author's
   names), python-pptx as the facade. The facade is where most Python users will live.
2. **Units.** Recommendation: points in the content API, EMU `Length` in the facade, as for docx.
3. **Chart data.** Recommendation: `replace_data` writes the caches now and the embedded workbook
   only once CR-005 exists; chart creation waits for that (Phase F). A hand-rolled minimal
   workbook writer, which is what python-pptx carries, is a second SpreadsheetML writer and is
   declined.
4. **SmartArt.** Recommendation: text read only, otherwise `Unsupported` with the XML available;
   no layout algorithm.
5. **Markdown into a slide.** Recommendation: title and body placeholders plus one table, nothing
   more; a deck from a whole markdown document (one slide per `#`) is a two-line loop over it and
   belongs in the MCP server (CR-006), not here.
6. **Effective text formatting.** Recommendation: direct values with a documented `inherited`
   read as `None`, and a resolver in Phase F; PowerPoint JS itself reports the effective value, so
   this is a recorded departure until then.
7. **Where the facade's module paths go.** Recommendation: mirror python-pptx's import paths
   (`docx4j_py.pptx.util`, `.enum.shapes`, `.chart.data`) because python-pptx code imports from
   them by name, unlike python-docx code which mostly imports from the top level.
