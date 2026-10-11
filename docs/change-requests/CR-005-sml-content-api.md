# CR-005: A SpreadsheetML content API, fidelity first, and an openpyxl facade

**Status:** Proposed 2026-09-19. Nothing implemented.
**Depends on:** CR-001 Phase D (the SML namespace generated, lazy import) and CR-002 Phase C
(`SpreadsheetMLPackage` with `WorkbookPart`, `WorksheetPart`, `SharedStringsPart`, `StylesPart`,
`CalcChainPart`, `TablePart`, chart and drawing parts), both proposed; CR-003 (implemented) for
the conventions, agent surface and facade pattern; **the xsdata fork's CR-001 (the serialiser)
before this starts**, because section 3.2's first constraint is only half an answer without it.
**Counterpart:** openpyxl 3.1.5 (the facade's oracle; the subset derived by introspection as for
python-docx), XlsxWriter 3.2 (what this CR deliberately does not compete with), the Excel
JavaScript API (`Excel.Workbook`, `Worksheet`, `Range`, `Table`, `NamedItem`, `Comment`,
`DataValidation`, `ConditionalFormat`, `Chart`) for the vocabulary, docx4j's `xlsx4j`
(`org.xlsx4j.sml`, `SpreadsheetMLPackage`, `WorksheetPart`, `SmlUtil`) for the semantics, and
the portfolio's `docs/python_ecosystem.md` for why.

## 1. Why

openpyxl is the default reader and writer of `.xlsx` in Python and the `pandas.read_excel`
engine; XlsxWriter is the best writer and cannot read. Neither edits an existing workbook without
loss: openpyxl keeps only what it models (shapes, some drawing content, slicers, sparklines and
anything it has not implemented are dropped on save, VBA only with `keep_vba`), and it is slow
and memory-heavy on large sheets unless the caller opts into read-only or write-only modes that
give up the other half. The pitch of this CR is one sentence: **edit an existing workbook without
losing anything.** The engine already has the half that matters (untouched parts byte for byte,
the whole schema typed); this CR adds a content API whose reads and writes are bounded by the
size of what is touched, not by the size of the workbook, and openpyxl's names on top.

## 2. Two constraints, first

They are in the summary because they shape everything below.

1. **Sheet data is not unmarshalled into typed objects by default.** A worksheet's `sheetData` is
   the bulk of any real workbook: a 50 MB `sheet1.xml` is ordinary. The typed model's serialiser
   runs at about 0.9 MiB/s and its parser at 2.6 MiB/s (measured 2026-09-18, docx4j-python
   `docs/python_ecosystem.md`), so unmarshalling that sheet is twenty seconds and marshalling it a
   minute. The xsdata fork's CR-001 aims to move those numbers, and even at its target they are
   not a worksheet's numbers. So: a `WorksheetPart` is read with **lxml** and kept as an lxml tree
   (the engine's `DefaultXmlPart` already does this for custom XML parts), cells are addressed in
   that tree through an index built once per sheet, writes edit the tree in place, and the save
   is `lxml.etree.tostring` (70 MiB/s). Typed objects are available **on demand** for anything
   outside `sheetData` (`sheetPr`, `cols`, `mergeCells`, `dataValidations`, `conditionalFormatting`,
   `tableParts`, `extLst`) through a typed view over the same lxml nodes, and for a single cell
   when a caller asks for `cell.element`. A sheet nobody touches is never parsed at all.
2. **Formula evaluation is out of scope.** Formulas are strings; cached values are read and
   reported as such (`Range.values` gives the cache, `Range.formulas` the text); a write that
   changes a formula, or a value a formula depends on, marks the workbook `fullCalcOnLoad` so
   Excel recalculates on open, and clears the affected cached values rather than leaving them
   stale. This is openpyxl's position too, and the `formulas` / `pycel` packages exist for the
   other one.

## 3. Design

### 3.1 Conventions

CR-003 section 3.1: `snake_case` with Excel JS's string values and `StrEnum`s; defaults so the
common call is short; `list`s and value views; frozen results with `to_dict()`; the error
hierarchy with `code` and `hint`. Two additions: **A1 references are the address form**
(`"Sheet1!B2:D10"`, what every user, Excel JS and openpyxl use), and **values are Python values**
(`int`, `float`, `str`, `bool`, `datetime`, `None`, or a `Formula`), converted at the boundary
with the workbook's date system (1900 or 1904) and the cell's number format deciding when a
serial is a date.

### 3.2 The surface (`docx4j_py/model/workbook/`)

```python
class Workbook:                   # Excel.Workbook, over a SpreadsheetMLPackage
    worksheets: list[Worksheet];  names: list[NamedItem];  tables: list[Table]
    active: Worksheet;  date_1904: bool;  calc_on_load: bool
    def add_worksheet(self, name=None, *, location="End") -> Worksheet
    def get_worksheet(self, name_or_address) -> Worksheet;  def get_range(self, address: str) -> Range   # "Sheet1!A1:B2"
    def outline(...) -> Outline;  def find(self, text, **options) -> list[SearchHit];  def describe(self) -> dict
    def to_markdown(self, *, addresses=False, max_rows=None, max_chars=None) -> str;  def to_api_script(...) -> str
    package: SpreadsheetMLPackage;  last_change: ChangeReport | None

class Worksheet:                  # Excel.Worksheet, over a WorksheetPart kept as lxml
    name: str;  index: int;  visibility: Literal["Visible", "Hidden", "VeryHidden"];  address: str
    used_range: Range | None;  tables: list[Table];  charts: list[Chart];  comments: list[Comment]
    freeze_panes: str | None;  auto_filter: str | None;  tab_color: str | None
    def get_range(self, address: str) -> Range;  def get_cell(self, row: int, column: int) -> Range   # 1-based, as Excel
    def get_used_range(self, *, values_only=False) -> Range
    def insert_rows(self, at: int, count=1);  def delete_rows(self, at, count=1);  the column twins   # references in formulas, names, tables, merges and validations shifted, as Excel does
    def add_table(self, address, *, name=None, style=None, header_row=True) -> Table
    def add_comment(self, address, text, *, author=None) -> Comment
    def insert_markdown(self, markdown: str, *, at: str = "A1") -> Range   # a GFM table into cells
    def delete(self) -> None;  def copy(self, *, name=None) -> Worksheet;  def to_dict(self) -> dict
    part: WorksheetPart;  element: lxml root (extension)

class Range:                      # Excel.Range: a rectangle on one sheet, one cell included
    address: str;  worksheet: Worksheet;  row_count: int;  column_count: int
    values: list[list[Any]]                                 # read: cached values as Python values; write: cells set, strings shared or inline (section 3.3)
    formulas: list[list[str | None]];  text: list[list[str]]   # formatted text through the number format, as Excel shows it
    number_format: str | list[list[str]];  format: RangeFormat  # font, fill, borders, alignment, wrap, column width, row height
    value / formula / number_format singular twins for a one-cell range
    merged: bool;  hyperlink: str | None;  data_validation: DataValidation | None
    def insert(self, shift: Literal["Down", "Right"]);  def delete(self, shift: Literal["Up", "Left"])
    def merge(self, *, across=False);  def unmerge();  def clear(self, apply_to: Literal["All", "Contents", "Formats"] = "All")
    def get_cell(self, r, c) -> Range;  def get_offset_range(...);  def get_resized_range(...);  def iter_rows(self, *, values_only=False)
    def copy_from(self, source: Range, *, values=True, formats=True)
    def to_markdown(self) -> str;  def to_dict(self) -> dict

class RangeFormat / Font / Fill / Borders / Alignment   # over the styles part's cellXfs registry (section 3.3)
class Table:                      # Excel.Table (ListObject): name, range, header_row, total_row, columns, style, add_rows(values)
class NamedItem:                  # defined name: name, refers_to, scope, value
class Comment:                    # legacy comment (author, text) and threaded comment (xl/threadedComments) read; legacy write
class DataValidation;  class ConditionalFormat   # read; the common kinds written (list, whole number, decimal, date; cell value, colour scale, data bar)
class Chart:                      # read: chart type, series references, title; anchor; no creation in this CR
class PivotTable:                 # read only: name, source, location
```

### 3.3 What is underneath: the three registries

- **The cell index.** On first access a sheet's lxml tree is walked once to map `(row, col)` to
  the `c` element (about 1 s per 50 MB, 70 MiB/s lxml parse plus the walk); `Range` reads go
  through it; a write inserts a `c` (and its `row`) in order when the cell is absent. Rows and
  cells stay in the file's order; nothing is rewritten that was not touched.
- **Shared strings.** `xl/sharedStrings.xml` is read with lxml and indexed; a written string is
  appended as a new shared string (Excel's own behaviour) unless the workbook already uses inline
  strings, in which case the write is inline. Rich text (`r` runs in an `si`) reads as the joined
  text and is preserved untouched unless the cell is rewritten.
- **Styles.** `xl/styles.xml`'s `cellXfs` is the registry: a `RangeFormat` write finds or appends
  an `xf` (and its `font` / `fill` / `border` / `numFmt`) exactly as openpyxl's `_writer` and Excel
  do, so two cells given the same format share one index and a format that exists is reused.
  The styles part is edited as lxml too; it is small but its `dxfs` and `cellStyles` must survive.

Everything else (`workbook.xml`, `[Content_Types].xml`, the `.rels`, `calcChain.xml`, tables,
comments, drawings, charts, pivot caches, VBA, custom XML, the printer settings binaries) goes
through the engine unchanged: typed on demand, byte for byte when untouched. `calcChain.xml` is
**deleted** on any formula edit, which is what Excel tolerates and openpyxl does; a workbook with
no formula edit keeps it.

### 3.4 The agent surface

CR-003 section 3.4 with sheets and ranges as the levels:

- **Addresses** are `"Sheet1!B2"` / `"Sheet1!A1:D10"`; a sheet by name (Excel's own stable handle;
  `sheetId` second, ordinal third); a table by name; a chart by `cNvPr` id. `AddressError` names
  the sheets and the used range.
- **`outline()`:** per sheet its used range, row and column counts, tables with their headers,
  names, charts, comments, whether it has formulas, and the first `max_rows` rows as text under
  the budget; `truncated` when the budget cuts. A 100-sheet workbook is a few KB.
- **`find()`** over values and formulas with cell addresses; **`describe()`** the workbook's
  shape for a template (tables, names, validations, the sheets an agent would fill), which is the
  spreadsheet analogue of CR-003's `describe()`.
- **`ChangeReport`** with `parts_touched`, the ranges written, the shared strings and styles added,
  and `warnings` (a formula left uncalculated, a cached value cleared, `calcChain` removed);
  `dry_run` over the lxml trees (a deep copy per touched sheet) with the added-part undo log.
- **Markdown:** `to_markdown()` renders each sheet's used range as a GFM table with a `## Sheet`
  heading, cut at `max_rows` with a trailing note, each row's address in a comment when asked;
  `insert_markdown` puts a GFM table into cells at an anchor. `to_api_script` emits the
  `get_range(...).values = ...` and format calls that reproduce a range, `insert_xml` of a sheet's
  fragment for the rest.

### 3.5 The facade, `docx4j_py.xlsx`

openpyxl's public API in openpyxl's names and semantics:

```python
from docx4j_py.xlsx import load_workbook, Workbook
from docx4j_py.xlsx.styles import Font, PatternFill, Alignment, Border, Side
from docx4j_py.xlsx.utils import get_column_letter

wb = load_workbook("in.xlsx")                       # or Workbook() for a new one
ws = wb.active
ws["A1"] = "Region";  ws["A1"].font = Font(bold=True)
ws.cell(row=2, column=1, value="North")
ws.append(["South", 98])
for row in ws.iter_rows(min_row=2, values_only=True): print(row)
ws.column_dimensions["A"].width = 18;  ws.freeze_panes = "A2";  ws.merge_cells("C1:D1")
wb.save("out.xlsx")
```

`Workbook`, `Worksheet`, `Cell`, `MergedCell`, `RowDimension`, `ColumnDimension`, the style
objects (`Font`, `PatternFill`, `GradientFill`, `Border`, `Side`, `Alignment`, `Protection`,
`NamedStyle`), `Table`, `DataValidation`, the conditional-format `Rule`s, `Comment`,
`DefinedName`, `PageMargins`, `PageSetup`, `SheetView`, `Image` (add), the chart classes
(`BarChart`, `LineChart`, `PieChart`, `Reference`, `Series`) for **creation** when Phase G lands,
`ReadOnlyWorksheet` and `WriteOnlyWorksheet` as the same class with the mode flags accepted and
ignored (this engine is always read-and-write and never loads a sheet it does not touch; recorded
as a departure), `openpyxl.utils` (`get_column_letter`, `column_index_from_string`,
`coordinate_to_tuple`, `range_boundaries`). The subset list (`tests/openpyxl_subset.json`) derived
by introspection of openpyxl 3.1.5 (178 members over the twenty-two classes measured on
2026-09-19; the whole public API is larger and the derivation script decides the class list); the
parity test builds one ordinary openpyxl script through both libraries and reads both files back
**with openpyxl**, plus the fidelity test openpyxl cannot pass: a real workbook with a shape, a
slicer and a pivot table, opened, one cell changed, saved, and every other part byte for byte.
openpyxl is never a dependency; `pandas` is not a target (section 9).

## 4. Rules

CR-003 section 4's format-independent rules, and:

- **Nothing is dropped.** A sheet's unknown content, its `extLst`, its `x14ac:` and `xr:`
  attributes, are lxml's to keep and it keeps them; the skipped-content report applies to the
  typed parts only and stays empty over the corpus.
- **A read never rewrites.** `outline()`, `find()`, `describe()`, `to_markdown()` and every
  `Range` read leave the sheet's bytes what they were; `DefaultXmlPart`'s "parsed for reading" is
  the mechanism (CR-003 section 17.3).
- **Shifting rows and columns shifts references** in formulas (the sheet's own and other sheets'
  references to it), defined names, tables, merged ranges, validations, conditional formats and
  hyperlinks, as Excel does; it is the one operation with a non-local effect and it says so in
  `ChangeReport.moved`.
- **Dates are the number format's decision**, as in Excel and openpyxl: a serial in a cell whose
  format is a date format reads as `datetime`; a `datetime` written gets a date format if the
  cell has none.
- **`fullCalcOnLoad` on any formula-affecting write**, `calcChain` removed, cached values of the
  cells written cleared; never a stale cache presented as a value without a warning.

## 5. Where each piece lives

| Piece | Package |
|---|---|
| the lxml worksheet part, the cell index, shared strings, the styles registry | `docx4j_py.openpackaging.parts.sml` (the engine, with CR-002 Phase C) and `docx4j_py.model.workbook.store` |
| `Workbook`, `Worksheet`, `Range`, `RangeFormat`, `Table`, `NamedItem`, `Comment`, validations, conditional formats, charts and pivots read | `docx4j_py.model.workbook` |
| the agent surface | reused from `docx4j_py.model.content.addresses` / `.reports` / `.trial` with a sheet-shaped container |
| markdown out and in, `to_api_script` | `docx4j_py.model.markdown.workbook`, `docx4j_py.model.workbook.api_script` |
| the facade | `docx4j_py.xlsx` with openpyxl's module paths mirrored (`.styles`, `.utils`, `.chart`, `.worksheet.table`) |
| the subset list and the parity test | `scripts/openpyxl_subset.py`, `tests/openpyxl_subset.json`, `tests/workbook/test_openpyxl_parity.py` |

## 6. What the engine supplies, and what this CR does without

CR-002 C: the package, part classes, `create_package()` with Excel's default workbook (one sheet,
the default styles and theme, `date1904` false). CR-001 D: the typed `sml` namespace for the
on-demand views. The xsdata fork's CR-001: the throughput that makes the on-demand typed views
tolerable for the medium-sized parts (`styles.xml` with thousands of `xf`s, a `table1.xml`, a
chart). Without it this CR still works, because its hot path never touches the typed
serialiser; it is a dependency for the experience, not the design.

## 7. Tests

- **Fixtures:** docx4j's `loadAndSave.xlsx` (in `samples/`), the workbooks under
  `../docx4j/docx4j-samples-xlsx4j/sample-docs/`, one saved by a current Excel with a table, a
  pivot table, a slicer, a shape, a chart, threaded comments and a data validation; a 50 MB
  generated sheet for the performance tests (built by the test, not committed).
- **Per phase**, in `tests/workbook/`: reads against known values (dates in both systems, shared
  and inline strings, rich text, booleans, errors); every write followed by save, reload and read
  back **with openpyxl** as well as with this engine; the fidelity test of section 3.5; the
  reference-shifting matrix for row and column inserts and deletes.
- **Performance**, marked `slow`: open the 50 MB workbook and read one cell in under two seconds;
  write one cell and save in under four; memory under three times the file size.
- **Agent scenarios** and the **subset and parity tests** as CR-004 section 7.
- **Excel acceptance**, manual, in `tests/README.md`: the saved workbook opens without a repair
  prompt, recalculates on open, the table and the pivot survive; save from Excel, close, reopen.

## 8. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | The store: lxml worksheet parts, the cell index, shared strings, the styles registry, the value conversions, `fullCalcOnLoad` and `calcChain` handling | 3 days |
| B | `Workbook`, `Worksheet`, `Range` reads: values, formulas, text, formats, used range, merges, names, tables read | 2 days |
| C | Writes: values and formulas, formats through the registry, merges, rows and columns with reference shifting, `add_worksheet`, `copy`, `delete` | 3 days |
| D | Tables, comments, data validation and conditional formats (read all, write the common kinds), hyperlinks | 2 days |
| E | The agent surface: addresses, `outline()`, `find()`, `describe()`, `ChangeReport`, `dry_run`; markdown out and in; `to_api_script` | 3 days |
| F | The openpyxl facade, its subset list, the parity and fidelity tests, the acceptance workbook | 4 days |
| G | Charts and images: read in D, creation here through the typed chart model; `Image.add` | 3 days, later |

Seventeen days for A to F. Section 2's constraints are Phase A's acceptance test.

**Hand-off from docx4j CR-035 phase 2 (2026-10-12, docx4j `276325167` on `VERSION_17_3_3`,
unpushed).** Phase D's `Chart` read and Phase G's creation have an oracle: docx4j's chart read
model (`org.docx4j.model.chart`: `Charts.in`, `ChartModel`, `ChartExModel`, `ChartCells`) and its
32 JSON goldens in `../docx4j/docx4j-core-tests/src/test/resources/charts/json/`, one per chart
part of every fixture in `charts/`, 16 of them Excel-made probes (`p2-*`) covering every chart
kind, a chartsheet, gaps and `#N/A`, multi-level categories, a filtered series and category, and a
combination chart on a secondary axis. The detail, the MS-OI29500 "shown" rules and Excel's
filtered-category markup (`c:f` narrowed to a union of the visible cells, the full range in
`c15:fullRef`) are recorded once, in CR-004 section 8, which this CR shares; `ChartCells` --- the
data from the cells against the caches --- is the piece a workbook needs that a deck does not.
Consumed from the sibling checkout, nothing copied, as the parity goldens are.

## 9. Open questions (recommendations for decision)

1. **Vocabulary.** Recommendation: Excel JS for the content API, openpyxl for the facade, for the
   same reasons as CR-004 question 1.
2. **The sheet store: lxml in place, or a compact cell table of Python values rebuilt on save.**
   Recommendation: lxml in place. It is the fidelity guarantee (untouched cells are untouched
   bytes), it is what the engine already does for custom XML, and a bulk `values` read over the
   index is fast enough; a compact table is XlsxWriter's design and the wrong trade here.
3. **Formula evaluation.** Recommendation: out of scope, permanently; `fullCalcOnLoad` and a
   warning. A caller who needs values now runs LibreOffice headless or `formulas`.
4. **A pandas engine.** Recommendation: no. `pd.read_excel(engine=...)` is a bulk read where
   openpyxl's read-only mode and `calamine` are the right tools; `Range.values` into a `DataFrame`
   is one line for the caller and is shown in the README.
5. **A write-only mode.** Recommendation: no; XlsxWriter owns that, and the engine's created
   workbook with `Range.values = rows` is adequate for the sizes an agent writes.
6. **Threaded comments.** Recommendation: read in Phase D, write legacy comments only; threaded
   comment writes (`xl/threadedComments`, `persons`) in a later phase after an Excel check, as
   CR-003 did with modern comments.
7. **VBA.** Recommendation: `.xlsm` loads and saves with `vbaProject.bin` untouched, as the engine
   already guarantees; no `keep_vba` flag because there is nothing to opt into.
8. **Sequencing against the xsdata fork's CR-001.** Recommendation: fork CR-001 phases A and B
   first, as the user asked on 2026-09-18; this CR's Phase A can start in parallel since its hot
   path is lxml.
