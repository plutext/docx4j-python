# CR-003: A content API for Python, for agents and MCP, in the vocabulary of Office JS

**Status:** Proposed 2026-09-16; revised the same day to follow Python conventions throughout
and to be designed for AI projects and MCP servers first; open questions decided 2026-09-16
(section 9); **Phase A implemented 2026-09-16** (section 10); **Phase B implemented 2026-09-16**
(section 11); **Phase D implemented 2026-09-16** (section 12); **Phase K implemented 2026-09-17**
(section 13); **Phase C implemented 2026-09-17** (section 14); **Phase G implemented
2026-09-17** (section 15); **Phase F implemented 2026-09-17** (section 16). Phases E, H, I and J
proposed.
**Depends on:** CR-001 Phases A to C (the model, `el`, the builders, `wml(...)`, `text_of`,
`walk`, `find`) and CR-002 Phase A (packages, parts, load and save), both implemented. Effective
formatting and list labels need CR-002 Phase B (`PropertyResolver`, the numbering `Emulator`);
until it lands the views read direct formatting and say so.
**Counterpart:** docx4j-core-ts [CR-002](../../../docx4j-core-ts/docs/change-requests/CR-002-content-api.md)
(phases B, C, D, E, F, G and I implemented 2026-09-10 to 2026-09-16; H proposed), whose object
model and rules this CR keeps so the two engines stay one design, and whose implementation
notes are folded into section 4 as rules; `plutext/docx4j-mcp` (`CR-mcp-server.md`: coarse
tools, markdown in and out, path policy, inline caps) for what an MCP server needs; python-docx
for what Python code written by people and by models already looks like; docx4j
`MainDocumentPart.addParagraphOfText` / `addObject` / `getContent`, `TextUtils`,
`BindingHandler`, `AcceptTrackedChanges`.

## 1. Why

CR-001's Hello World is honest and low-level:

```python
body.content.append(el.p(content=[el.r(content=[el.t("Hello World")])]))
```

The sugar removes the typing, but the tree is still the whole API: there is no "add a paragraph
after the one that says X", no "make this range bold" that splits runs at the boundaries, no
address an agent can hold across tool calls, no way to read a 200-page document into a context
window that holds twenty. Three consumers are in view, and the third is new to this CR:

- A **developer** who knows Word documents and wants the common operations to read the way
  they do in python-docx, docx4j or Office JS.
- An **agent writing Python** against this package: Claude Code, or code an LLM emits at run
  time. It writes what it has seen most, which is python-docx, and it recovers from errors by
  reading them, so the error text has to say what to do instead.
- An **agent behind an MCP server**, which never sees the API. It sees tools with JSON
  arguments and JSON results, a finite context window, and a human who will open the result in
  Word and wants to know what the agent changed. docx4j-mcp's position is that such tools should
  be few and coarse: a document in, a document out, markdown as the exchange format. That is
  right for creating and converting documents. Editing a document that exists needs one more
  thing, an *address*, so that "change the third paragraph of section 2" is one deterministic
  call and not a round trip through markdown. Both tool shapes rest on the same API; this CR
  makes sure the API serves both.

And a fourth reason: docx4j-core-ts and docx4j-python are one design in two languages. Where
the TypeScript content API has settled a shape or a rule, this CR uses it, so the documentation,
the MCP tool surface and the agents' training material read across, and a fix in one is a fix
for both. Where Python has a convention, the convention wins over the TypeScript spelling.

## 2. What others do

| Library | Building | Reading and editing | What an agent gets |
|---|---|---|---|
| python-docx | `Document()`, `doc.add_paragraph("x", style="Heading 1")`, `p.add_run("bold").bold = True`, `doc.add_table(rows, cols)`, `doc.add_picture(path, width=Inches(2))`, `doc.save(path)` | `doc.paragraphs`, `p.text`, `p.runs`, `p.style.name`, `table.cell(r, c).text` | the API models already know by heart; a thin typed slice over lxml; nothing outside it |
| docx4j (Java) | `mdp.addParagraphOfText`, `addStyledParagraphOfText`, `addObject`; `XmlUtils.unmarshalString` | `getContent()`, `TraversalUtil`, `TextUtils.extractText` | the whole schema; no addresses |
| Office JS | `body.insertParagraph(text, "End")`, `insertText`, `insertOoxml`, `insertTable` | `body.paragraphs`, `body.text`, `body.search(text, options)` returning ranges | the most documented Word API there is; insert locations `Start` / `End` / `Before` / `After` / `Replace` |
| docx4j-core-ts | Office JS's shapes over the docx4j tree; `insertXml`; `el.p()` | `search` across runs, `Range.font`, `outline()`, addresses, comments, tracked changes, XML mapping | addresses and `outline()` for agents; a `Word` shim so add-in code runs in Node |
| docx4j-mcp (Java) | `markdown_to_docx`, `html_to_docx`, `fill_template` | `docx_to_markdown`, `extract_text`, `describe_template` | few coarse tools, path allow-list, inline caps, `overwrite: true`; deliberately no fine-grained editing |
| Office-Word-MCP-Server | `add_paragraph(file, text, style)`, `add_heading`, `add_table`, `insert_line_or_paragraph_near_text` | `get_document_outline`, `get_paragraph_text_from_document(file, index)`, `find_text_in_document`, `search_and_replace`, `format_text(paragraph_index, start, end, ...)` | flat tools over python-docx; a paragraph index as the address, which shifts on every insert |

Two observations. Every developer library converges on the same verbs, and Office JS's insert
locations are the cleanest spelling of them. Every agent-facing server converges on addresses
and string payloads, and the ones built on python-docx use a paragraph index that breaks the
moment something is inserted. The API below takes Office JS's verbs, python-docx's habits, and
gives agents addresses that survive edits.

## 3. Design

### 3.1 Python conventions, the rules

These hold across every class in this CR and are what "follows Python conventions" means here.

- **Names**: `snake_case` members in Office JS's vocabulary (`insert_paragraph`, `style_built_in`,
  `parent_table_cell`). The *values* Office JS defines stay its strings, since they are data that
  cross into tool arguments: `"Start"`, `"End"`, `"Before"`, `"After"`, `"Replace"`, `"Centered"`,
  `"Single"`, `"TrackAll"`, `"Heading1"`. Each is a `typing.Literal` in the signature and a
  `StrEnum` for those who prefer a name (`InsertLocation.END == "End"`), so both spellings work
  and neither is second class.
- **The common call is short.** Every location and option has a default, and options are
  keyword-only: `body.insert_paragraph("Hello")` appends; `body.insert_paragraph("Hello",
  location="Start")` prepends; `search("fox", match_case=True)`. Positional arguments are for
  what is always given.
- **Properties for cheap reads and writes, methods for work.** `paragraph.text`, `font.bold`,
  `table.row_count` are properties; anything that takes an argument, allocates, or touches
  another part is a method. No `get_` prefix on a property; `get_text(view=...)` exists because
  it takes an argument.
- **Collections are sequences and views are values.** `body.paragraphs` is a `list[Paragraph]`
  in document order; `Body` itself is a `Sequence` of its block-level children (`len(body)`,
  `body[3]`, `for block in body`), and `"Chapter 2" in body` is a text search. A view holds the
  element and its container, is created on access, compares equal to another view of the same
  element, hashes, and has a `repr` that shows its address and a text preview:
  `<Paragraph body/3 'Chapter 2: The …'>`. `str(paragraph)` is its text.
- **Results are frozen dataclasses with `to_dict()`**, never tuples or ad-hoc dicts: `Outline`,
  `OutlineEntry`, `SearchHit`, `ChangeReport`, `SkippedEntry`. `to_dict()` gives JSON-ready
  primitives, which is what a tool result is; `dataclasses.asdict` works too. `to_json()` on
  `Outline` and `ChangeReport` because they are what servers return whole.
- **Errors are one hierarchy, and every message says what to do instead.** `Docx4JError` (from
  CR-002) with `ContentError`, `AddressError`, `InvalidTargetError`, `TrackedChangeError`,
  `BindingError`; each carries `code` (a stable string such as `"address.not_found"`), the
  human message, and `hint`, a sentence an agent can act on (`"call outline() to list current
  addresses"`). No bare `ValueError` for a document-model problem.
- **Paths are `str | os.PathLike`**, bytes are `bytes`, base64 is a `str` and never the only
  way in: every `..._from_base64` has a `..._from_bytes` twin, because Python callers have bytes.
- **Iteration is lazy where the document is large**: `body.iter_paragraphs()`,
  `body.iter_blocks()` and `search(..., limit=)` exist beside the list forms, for a
  thousand-page document read in a loop.
- **Type-complete and docstring-complete**, because that is what a model writing code against
  this API reads: `Literal` and `Self` in signatures, `__all__` in every module, a one-line
  docstring per public member saying what it does and what Office JS calls it, a
  `py.typed` marker.
- **No proxies, no magic.** A view is a plain object over the tree; `paragraph.element` is the
  `P`; `body.content` is the live `ChildList`. `__getattr__` tricks that make an unknown member
  raise something clever are not wanted; `AttributeError` is Python's message for that.

### 3.2 The surface (`docx4j_py/model/content/`)

```python
Location = Literal["Start", "End", "Before", "After", "Replace"]

class Body(Sequence[Paragraph | Table | ContentControl]):   # Word.Body: pkg.body, and on any part or
                                                            # container whose root has a content list
    paragraphs: list[Paragraph]                             # document order; tables and controls descended
    tables: list[Table]
    content_controls: list[ContentControl]
    inline_pictures: list[InlinePicture]
    text: str                                               # a paragraph per line; the accepted view; no text boxes
    def get_text(self, *, view: Literal["accepted", "original"] = "accepted", max_chars: int | None = None) -> str
    def iter_paragraphs(self) -> Iterator[Paragraph];  def iter_blocks(self) -> Iterator[Paragraph | Table | ContentControl]

    def insert_paragraph(self, text: str = "", *, location: Literal["Start", "End"] = "End", style: str | None = None) -> Paragraph
    def insert_text(self, text: str, *, location: Literal["Start", "End", "Replace"] = "End") -> Range
    def insert_table(self, row_count: int, column_count: int, *, location: Literal["Start", "End"] = "End", values: list[list[str]] | None = None, style: str | None = None) -> Table
    def insert_break(self, type: Literal["Page", "Line", "SectionNext", "SectionContinuous"] = "Page", *, location: Literal["Start", "End"] = "End") -> None
    def insert_inline_picture(self, data: bytes, *, location: Literal["Start", "End"] = "End", width: float | None = None, height: float | None = None, alt_text_description: str = "", alt_text_title: str | None = None) -> InlinePicture
    def insert_inline_picture_from_base64(self, base64: str, *, location="End", **options) -> InlinePicture
    def insert_ooxml(self, ooxml: str, *, location: Literal["Start", "End", "Replace"] = "End") -> list[Paragraph | Table]   # pkg:package, as Word; or a bare fragment
    def insert_xml(self, xml: str, *, location: Literal["Start", "End"] = "End") -> list[Paragraph | Table]                 # a w:p / w:tbl fragment, prefixes supplied
    def insert_markdown(self, markdown: str, *, location: Literal["Start", "End"] = "End") -> list[Paragraph | Table]       # 3.5
    def insert_element(self, element: Child, *, location: Location = "End", target: Paragraph | Table | None = None) -> None   # docx4j addObject; rejects what a body cannot hold
    def insert_content_control(self, kind: ContentControlType = "RichText") -> ContentControl
    def search(self, text: str, *, match_case: bool = False, match_whole_word: bool = False, match_wildcards: bool = False, limit: int | None = None) -> list[Range]
    def find(self, text: str, *, context: int = 40, limit: int = 20, **search_options) -> list[SearchHit]   # 3.4: hits with addresses and snippets
    def replace_text(self, find: str, replace: str, **search_options) -> int
    def clear(self) -> None
    def get_range(self, location: Literal["Whole", "Start", "End", "Content"] = "Whole") -> Range
    def get_comments(self) -> list[Comment]
    def get_tracked_changes(self) -> list[TrackedChange]
    def accept_all(self) -> None;  def reject_all(self) -> None
    def to_markdown(self, *, max_chars: int | None = None) -> str
    def get_xml(self) -> str
    # addresses (3.4)
    def outline(self, *, depth: int | None = None, max_chars: int = 80, headings_only: bool = False) -> Outline
    def element_at(self, address: str) -> Paragraph | Table | ContentControl
    def paragraph_at(self, address: str | None = None, *, contains: str | None = None, para_id: str | None = None) -> Paragraph
    def address_of(self, view) -> str
    # the tree
    content: ChildList;  part: XmlPart
    # docx4j aliases
    add_paragraph_of_text, add_styled_paragraph_of_text, add_object, get_content

class Paragraph:                  # Word.Paragraph
    text: str                                              # read: runs joined; write: one run replaces them
    style: str                                             # display name ("Heading 1"), section 4
    style_built_in: str                                    # Word.Style value ("Heading1") or "Other"
    style_id: str | None                                   # w:pStyle, docx4j's view
    alignment: Literal["Left", "Centered", "Right", "Justified", "Unknown"]
    font: Font                                             # over the runs
    left_indent: float; right_indent: float; first_line_indent: float
    space_after: float; space_before: float; line_spacing: float          # points, as Office JS
    outline_level: int
    is_list_item: bool;  list_item: ListItem | None;  list: List | None  # phase H
    parent_body: Body;  parent_table_cell: TableCell | None;  parent_content_control: ContentControl | None
    inline_pictures: list[InlinePicture];  content_controls: list[ContentControl]
    address: str;  para_id: str | None                     # the two stable handles (3.4)
    runs: list[R];  element: P
    def insert_text(self, text, *, location: Literal["Start", "End", "Replace"] = "End") -> Range
    def insert_paragraph(self, text="", *, location: Literal["Before", "After"] = "After", style=None) -> Paragraph
    def insert_break(self, type="Page", *, location: Literal["Before", "After", "Start", "End"] = "End") -> None
    def insert_inline_picture(self, data: bytes, *, location: Literal["Start", "End", "Replace"] = "End", **options) -> InlinePicture
    def insert_ooxml(self, ooxml, *, location=...) -> list[...];  def insert_xml(self, xml, *, location=...) -> list[...];  def insert_markdown(...)
    def insert_content_control(self, kind="RichText") -> ContentControl
    def insert_comment(self, text: str) -> Comment
    def search(self, text, **options) -> list[Range];  def find(self, text, **options) -> list[SearchHit]
    def replace_text(self, find, replace, **options) -> int
    def get_range(self, location="Whole") -> Range
    def get_text(self, *, view="accepted") -> str
    def get_comments(self) -> list[Comment];  def get_tracked_changes(self) -> list[TrackedChange]
    def start_new_list(self) -> List;  def attach_to_list(self, list_id: int, level: int = 0) -> None;  def detach_from_list(self) -> None
    def to_markdown(self) -> str;  def get_xml(self) -> str;  def to_dict(self) -> dict
    def delete(self) -> None

class Range:                      # Word.Range: a span within one paragraph, or a whole paragraph
    text: str;  font: Font;  paragraphs: list[Paragraph];  style: str;  style_built_in: str;  style_id: str | None
    start: int;  end: int                                  # code-point offsets within the paragraph's accepted text
    def insert_text(self, text, *, location: Location = "Replace") -> Range
    def insert_paragraph(self, text="", *, location: Literal["Before", "After"] = "After") -> Paragraph
    def insert_ooxml(self, ooxml, *, location: Literal["Before", "After", "Replace"] = "Replace") -> list[...]
    def insert_content_control(self, kind="RichText") -> ContentControl
    def insert_comment(self, text) -> Comment
    def search(self, text, **options) -> list[Range];  def replace_text(self, find, replace, **options) -> int
    def get_comments(self) -> list[Comment];  def get_tracked_changes(self) -> list[TrackedChange]
    def get_xml(self) -> str;  def delete(self) -> None
    runs: list[R]

class Font:                       # Word.Font, backed by w:rPr
    bold: bool; italic: bool; strike_through: bool; subscript: bool; superscript: bool
    underline: Literal["None", "Single", "Double", "Dotted", "Wave", "Thick", ...]
    name: str;  size: float;  color: str;  highlight_color: str | None
    def to_dict(self) -> dict

class Table:                      # Word.Table, backed by w:tbl
    row_count: int;  rows: list[TableRow];  values: list[list[str]]
    style: str;  style_built_in: str;  style_id: str | None;  header_row_count: int
    text: str;  address: str;  parent_table_cell: TableCell | None;  element: Tbl
    def cell(self, row_index: int, cell_index: int) -> TableCell           # Office JS getCell; python-docx's name
    def add_rows(self, row_count: int = 1, *, location: Literal["Start", "End"] = "End", values=None) -> list[TableRow]
    def delete_rows(self, row_index: int, row_count: int = 1) -> None
    def column_widths(self) -> list[float]
    def to_markdown(self) -> str;  def to_dict(self) -> dict;  def delete(self) -> None
class TableRow:   row_index, cell_count, cells, values, is_header, insert_rows(count=1, *, location="After", values=None), delete()
class TableCell:  body: Body, paragraphs, tables, text, value, insert_paragraph, insert_text, row_index, cell_index, parent_row, parent_table, width, column_width

class InlinePicture:              # Word.InlinePicture
    width: float;  height: float;  alt_text_description: str;  alt_text_title: str | None;  image_format: str
    def get_bytes(self) -> bytes;  def get_base64(self) -> str;  def delete(self) -> None
    paragraph: Paragraph;  rel_id: str;  inline: Inline;  image_part: ImagePart

class ContentControl:             # Word.ContentControl, backed by w:sdt in all four forms
    type: ContentControlType;  form: Literal["block", "run", "row", "cell"]
    tag: str;  title: str;  id: int;  text: str;  placeholder_text: str
    appearance: Literal["BoundingBox", "Tags", "Hidden"];  color: str;  cannot_delete: bool;  cannot_edit: bool;  remove_when_edited: bool
    xml_mapping: XmlMapping;  address: str
    paragraphs, tables, content_controls, inline_pictures
    checkbox_content_control, date_picker_content_control, drop_down_list_content_control, combo_box_content_control, picture_content_control, repeating_section_content_control, group_content_control
    def insert_text(...);  def insert_paragraph(...);  def search(...);  def get_range(...);  def get_xml(self) -> str;  def to_dict(self) -> dict
    def delete(self, *, keep_content: bool = True) -> None
    element: SdtBlock | SdtRun | CTSdtRow | CTSdtCell
```

`WordprocessingMLPackage` gains `body`, `outline()`, `paragraph_at`, `element_at`, `author`,
`change_tracking_mode`, `tracked_change_date`, `get_tracked_changes()`, `custom_xml_parts`,
`to_markdown()`, `describe()` (3.4); `MainDocumentPart`, `HeaderPart`, `FooterPart`,
`FootnotesPart`, `EndnotesPart` and `CommentsPart` gain `body`.

The Hello World:

```python
from docx4j_py import create_package

pkg = create_package()
pkg.body.insert_paragraph("Hello World")
pkg.save("hello.docx")
```

Editing what is there:

```python
body = pkg.body
title = body.insert_paragraph("Report", location="Start", style="Heading1")
title.alignment = "Centered"
hit = body.search("quick brown fox")[0]           # across runs
hit.font.italic = True                            # runs split at the boundaries
hit.insert_text("slow red fox")                   # Replace is a Range's default location
body.paragraphs[-1].insert_paragraph("The end.")  # After is a Paragraph's
body.text                                         # a paragraph per line
```

### 3.3 Tree layer additions (CR-001's `docx4j_py.wml.builders`)

The TypeScript objects package found, a phase at a time, that the content API wrote builders it
should have had from the tree. They are Phase A here, and live with `p`, `r`, `t` and `tbl`:

- `tr(cells, *, widths=None)` and `tc(blocks, *, width=None)`.
- `inline_picture(rel_id, *, cx, cy, id, name, descr="", title=None)`: the `w:drawing` /
  `wp:inline` exactly as docx4j's `BinaryPartAbstractImage.createImageInline` writes it.
- `sdt(content, *, kind, tag=None, title=None, id=None, form=None)`, `sdt_pr(**options)`,
  `next_sdt_id(root)`, `sdt_property(sdt_pr, local_name, namespace=W)` (looks in `w:` and
  `w15:`, because Word writes `w15:dataBinding` on a repeating section), `sdt_kind_of(sdt_pr)`.
- `rpr_to_elements(rpr)` and `rpr_from_elements(elements)`: `w:rPrChange/w:rPr` is a base-type
  element list where `w:rPr` has named properties; `EG_RPrBase` order, w14 effects included.
- `deep_copy_as(obj, cls)`: a copy typed as a base class (`PPrBase` for `w:pPrChange/w:pPr`).
  xsdata writes `xsi:type` when an instance's class differs from the declared field type, which
  is valid but not what Word writes.
- `walk_all(root, visitor, wildcard_visitor)`: `walk` that also enters wildcard content, for
  anything that rewrites references.
- `run_items_of(holder)`: the run list of a `w:p`, `w:hyperlink`, `w:ins`, `w:del`, `w:moveFrom`,
  `w:moveTo`, `w:smartTag`, `w:customXml` or run-form `w:sdtContent`, by the model's field names.

### 3.4 The agent surface

This is the section the revision adds. It is what an MCP server, or an agent writing Python,
needs beyond the views, and each item is in the library rather than the server so that every
server (and every notebook) gets the same behaviour.

**Addresses that survive edits.** Three forms, accepted wherever a `Paragraph | Table` target is
and reported everywhere a view is:

- **paraId first.** `"w14:5A2B1C3D"`, the `w14:paraId` Word writes on every paragraph it saves.
  Stable across every edit, including inserts before it. `Paragraph.address` reports it when the
  paragraph has one; new paragraphs get one when the document already uses them (and always in
  a created document, from a per-package generator that a test can seed).
- **Ordinal second.** `"body/3"`, `"body/4/0/1/0"` (a paragraph in a cell), `"header:rId3/0"`;
  the fallback for documents without paraIds and the form `outline()` uses to show structure.
  Stable until an insert or delete before it; `ChangeReport` (below) says which ordinals moved.
- **Text third.** `paragraph_at(contains="Chapter 1")`, first match; what the MCP servers'
  `near_text` tools do. Every `AddressError` names the nearest surviving address and says to
  call `outline()`.

**`outline()` with a budget.** A document is read into a context window that cannot hold it, so
the outline is what an agent reads first and it must be small, structured and enough to choose
an address from. `Outline` is a frozen dataclass: `entries` (one `OutlineEntry` per block:
`address`, `para_id`, `kind` (`paragraph` / `table` / `control`), `style_id`, `level` for
headings, `text` truncated to `max_chars` with an ellipsis, `chars` the untruncated length,
`rows` and `cols` for a table), `headers`, `footers`, `stats` (paragraphs, tables, words,
comments, tracked changes), and `to_dict()` / `to_json()`. `depth=` limits nesting into tables
and controls; `headings_only=True` gives the table of contents view; `to_markdown()` on the
outline renders it as a nested list, which is the cheapest thing to show a model.

**`find()` returns hits with context.** `search()` returns `Range`s for code; `find()` returns
`SearchHit(address, para_id, start, end, snippet, before, after)` for tools, so a server can
show an agent *where* the matches are without a second call per hit, with `limit` and
`context` (characters each side) as the budget.

**`describe()` tells an agent what it can use.** `pkg.describe()` returns the styles the
document defines (id, display name, kind, whether built in), the page setup, the parts present
(headers, footers, comments, custom XML with their namespaces), the authors of comments and
revisions, whether tracking is on, and the `skipped` report; as a dataclass with `to_dict()`.
It is docx4j-mcp's `describe_template` widened to any document, and the answer to "which style
should I use for a heading here".

**Every mutation returns what it did.** A `ChangeReport` (frozen, `to_dict()`, `to_json()`) is
available after any content-API call as `pkg.last_change` and accumulates in `pkg.changes` until
`pkg.changes.clear()`: the operation, the addresses touched, the addresses that moved (old to
new ordinal), the paraIds created, the text before and after for a text edit, the parts that
will be re-marshalled. A server returns it as the tool result so the agent verifies without
re-reading the document, and a human can be shown a summary of what the agent did.

**`dry_run`.** A `with pkg.dry_run() as trial:` block applies calls to a deep copy of the affected
part (`deep_copy` from CR-001) and reports the `ChangeReport` without touching the document; the
context manager discards the copy. It is how a server offers a preview, and how an agent checks
a `replace_text` count before committing to it.

**The audit trail is Word's own.** `pkg.author = Author("Claude", initials="C")`,
`pkg.change_tracking_mode = "TrackAll"` and every edit thereafter is a tracked change a human
reviews in Word with accept and reject, and `range.insert_comment("Changed because …")` is how
the agent explains itself in the document rather than in a chat log. This is the single most
useful thing the API does for an AI workflow and it costs nothing beyond phases F and G; the
README leads with it.

**Long-lived process, many documents.** CR-002 accepted a one-second import on the condition
that a process serves many documents. An MCP server therefore holds packages open across tool
calls: `docx4j_py.model.sessions.DocumentSession` is a small registry (`open(path) -> handle`,
`get(handle)`, `save(handle, path=None)`, `close(handle)`, an idle timeout, one `ParserConfig`
per thread as CR-001 section 14 requires), so that "open, edit, edit, save" is four cheap calls
and not four loads. It is in the library because every server would otherwise write it.

**Budgets everywhere.** `get_text(max_chars=)`, `to_markdown(max_chars=)`, `outline(max_chars=)`,
`find(limit=, context=)`, `iter_paragraphs()`; when a result is cut, the dataclass says so
(`truncated=True`, `chars=`) rather than silently ending. docx4j-mcp's inline cap and
`output_path` spill are the server's; the library gives it the numbers.

**Determinism.** Same document, same calls, same bytes: relationship ids, image part names,
`w:id`s and paraIds are allocated from the document's own state, and the paraId generator is
seedable (`pkg.id_seed`) for tests and for reproducible agent runs.

**Errors an agent can act on.** Every `ContentError` has `code` and `hint` (3.1). Inserting an
element a container cannot hold names what was passed and what the container takes; an address
that no longer exists names the nearest one; a style that does not exist lists the five closest
by name; a `Range` that crosses a run holder says where to split. The server passes the message
through; the agent's next call is usually right.

**The tool surface it implies.** The MCP tools are not this CR, but the API is designed so that
each is a few lines and the two shapes docx4j-mcp distinguishes both fit: the coarse tools
(`docx_to_markdown` is `pkg.to_markdown()`, `markdown_to_docx` is `create_package()` plus
`insert_markdown`, `extract_text` is `body.text`, `describe` is `pkg.describe()`) and the
in-place ones (`outline`, `find`, `get`, `insert`, `replace`, `format`, `delete`, `comment`,
`accept_changes`, each taking an address and returning a `ChangeReport`).

### 3.5 Markdown in and out

Markdown is what models write and read best, and docx4j-mcp made it the exchange format for
that reason (its `markdown_to_docx` and `docx_to_markdown` are the docx4j-markdown module,
Java). The Python engine needs its own, small and honest:

- `body.to_markdown()`, `paragraph.to_markdown()`, `table.to_markdown()`, `pkg.to_markdown()`:
  headings from outline level (style or `w:outlineLvl`), bold, italic, strike, code (`w:rStyle`
  of a code style, or a monospace font), links (`w:hyperlink` with its relationship), lists
  (through phase H's `List`, before it as `- ` and `1. ` from `w:numPr`), tables as GFM pipe
  tables, images as `![alt](media/imageN.png)` with the bytes reachable through the outline,
  footnotes as GFM footnotes, comments and tracked changes as `{==...==}` / `{++...++}` /
  `{--...--}` CriticMarkup when `view="markup"` and the accepted view otherwise; the paragraph's
  address as an HTML comment (`<!-- body/3 -->`) when `addresses=True`, so a model can read a
  document as markdown and then edit it by address.
- `body.insert_markdown(markdown)`: CommonMark plus GFM tables through `markdown-it-py`
  (the one dependency this CR adds; pure Python, permissive licence), producing paragraphs with
  the document's own heading, list and table styles (`describe()`'s style list is what it
  chooses from, with docx4j's default names as fallbacks and the numbering part created when
  a list needs one).

Round trip is not a goal (markdown cannot hold a docx); a document read as markdown, edited as
markdown and written back through `insert_markdown` is a coarse workflow that loses formatting
the markdown could not carry, and `to_markdown(addresses=True)` plus in-place edits is the fine
one. The README says which to use when.

### 3.6 A python-docx facade, so existing code and model memory keep working

docx4j-core-ts built a `Word` shim so that add-in code runs unchanged in Node. The Python
analogue is not Office JS, it is **python-docx**: the API in every tutorial, in every model's
memory, and in every "write me a script that makes a Word document" answer. `docx4j_py.docx`
offers a structural subset of python-docx's public API over this engine:

```python
from docx4j_py.docx import Document, Pt, Inches

doc = Document("in.docx")                                  # or Document() for a new one
doc.add_heading("Report", level=1)
p = doc.add_paragraph("Plain, ")
p.add_run("bold").bold = True
p.style = "Heading 2"                                      # display name, as python-docx
doc.add_table(rows=2, cols=2).cell(0, 0).text = "a"
doc.add_picture("logo.png", width=Inches(2))
for paragraph in doc.paragraphs: print(paragraph.style.name, paragraph.text)
doc.save("out.docx")
```

`Document`, `Paragraph`, `Run`, `Font`, `Table`, `_Row`, `_Cell`, `Section`, the `Length`
units (`Pt`, `Inches`, `Cm`, `Emu`, `Twips`), `WD_ALIGN_PARAGRAPH` and the other enums
python-docx code uses, with `python-docx`'s semantics where they differ from Office JS's
(`paragraph.style` is a style object with `.name`; `run.font.size` is a `Length`;
`add_paragraph(text, style)` returns the paragraph). What python-docx code gets from underneath
that python-docx does not have: the whole schema (`paragraph._p` is the typed `P`, not lxml),
byte-for-byte untouched parts, tracked changes and comments, addresses. The subset is a
committed member list (`tests/python_docx_subset.json`, from python-docx's documented API at a
pinned version) and a test asserts it; python-docx's own test suite is not run.

The Office JS vocabulary remains the primary API and the one the MCP tools are written over;
the facade exists so that code already written, by people or by models, runs.

### 3.7 Custom XML parts, XML mapping and typed content controls

The same surface as docx4j-core-ts section 3.5: WordApiDesktop 1.3's `CustomXmlPart`,
`CustomXmlNode`, `XmlMapping` and the typed content-control kinds, over the parts CR-002 holds as
lxml trees. Simpler here in one respect that removes a design section: **XPath is lxml's**,
XPath 1.0 complete with a namespace map, so there is no `XPathEngine`, no optional dependency,
no `load()` to warm and nothing asynchronous. `select_nodes(xpath, namespace_mappings=None)`
takes Word's `xmlns:ns0='...'` prefix-mapping string or the part's namespace manager.

```python
part = pkg.custom_xml_parts.get_item("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}")
part.select_single_node("/ns0:invoice/ns0:customer/ns0:name", "xmlns:ns0='urn:invoice'").text = "Acme"
control.xml_mapping.set_mapping("/ns0:invoice/ns0:total", "xmlns:ns0='urn:invoice'", part)
pkg.custom_xml_parts.apply_bindings()             # docx4j BindingHandler.applyBindings, what Word does on open
pkg.custom_xml_parts.update_from_content_controls()
```

For agents this is the template story docx4j-mcp's `describe_template` and `fill_template`
tell: `pkg.custom_xml_parts.describe()` returns the skeleton (the XPaths the document binds,
their types and repeats) as a dataclass, and `fill(data: dict | str)` sets the nodes and applies
the bindings. Reading a custom XML part does not mark it for re-marshalling; a mutation through a
node view does.

### 3.8 Change tracking and find-and-replace

`pkg.change_tracking_mode` (`"Off"` | `"TrackAll"` | `"TrackMineOnly"`) over `w:trackRevisions`,
`pkg.author`, `pkg.tracked_change_date`. While the mode is on, every mutation writes Word's
revision markup in the paragraph-level primitives, so `Range`, `Body`, `Table` and later phases
inherit it: `w:ins`, `w:del` with `w:delText`, a replacement as `w:del` then `w:ins`, paragraph
marks, `w:rPrChange` and `w:pPrChange` recorded once on the first write, row insertions and
deletions on `w:trPr`. `TrackedChange` (`type`, `author`, `date`, `text`, `accept()`,
`reject()`, `get_range()`, `to_dict()`) from `get_tracked_changes()` on body, paragraph and
range; `accept_all()` / `reject_all()` on the body, accepting as docx4j's
`AcceptTrackedChanges` does. The text model reads the accepted view throughout;
`get_text(view="original")` gives the other. `replace_text` is `search` then `insert_text` from
the last match to the first, tracked when the mode is on. Reading or writing the mode
unmarshals the settings part, which round-trips canonically (CR-002 section 12.4).

### 3.9 Comments

`get_comments()` on body, paragraph and range; `insert_comment(text)` on paragraph and range;
`Comment` with `author_name`, `author_email`, `initials`, `content`, `creation_date`
(`datetime | None`), `resolved` (`w15:done`), `replies`, `reply(text)`, `delete()`,
`get_range()`, `to_dict()`, over `w:comments`, `w15:commentsEx`, `w16cid:commentsIds`,
`w:people`, and `w16cex:commentsExtensible` kept in step when present. Any of the first four the
document lacks is created with its relationship, content type and the two comment styles.
Reads unmarshal three parts, writes five; a document whose comments are not read keeps all of
them byte for byte.

### 3.10 Lists

`List` over a `w:num` and its `w:abstractNum`, `ListItem` (`level`, `list_string`,
`sibling_index`), and on `Paragraph`: `is_list_item`, `list`, `list_item`, `start_new_list()`
(docx4j's default definitions, creating the numbering part when absent), `attach_to_list`,
`detach_from_list`. `list_string` needs CR-002 Phase B's `Emulator`; until then `None`, and
`to_markdown` numbers lists itself.

### 3.11 `to_api_script`

docx4j-core-ts's phase I is a `Word` shim plus `toApiScript`. The shim has no Python counterpart
worth building (3.6 is Python's version of "existing code runs"). `to_api_script(target)` does
carry over: the content-API calls, as Python source against `body`, that produce a body, a
paragraph, a range or an element, falling back to `body.insert_xml(...)` for what the verbs
cannot express, decided per paragraph so a paragraph comes out whole either way. For an agent it
is reveal codes: shown a document, it learns the calls that would make it.

### 3.12 Scripts and directionality

Offsets are code points. A split (`Paragraph.split_at`, behind `Range.font`, `insert_text` at a
boundary, `insert_comment`, `insert_content_control`) never falls inside a grapheme cluster:
`unicodedata` combining classes plus the ZWJ and variation selectors, and the `regex` module's
`\X` when installed. `match_whole_word` uses Unicode word boundaries; `match_case=False` uses
`re.IGNORECASE`. The run mapping writes `w:bCs`, `w:iCs` and `w:szCs` with their Latin twins;
`alignment` maps `start` / `end` as well as `left` / `right`; the indents read both the
transitional and the strict attribute names and write whichever the paragraph uses; `bidi` on
`Paragraph`, `rtl` on `Font`.

### 3.13 What stays as it is

- `contents` remains the typed tree and `content` lists remain `ChildList`s. Nothing is wrapped;
  a user who appends to `body.content` keeps working.
- The packaging layer keeps docx4j's names; only the content API takes Office JS's, and the
  facade python-docx's. Extensions beyond Office JS are marked in the docstring.
- Everything mutates the tree in place, marks the part for re-marshalling, links parents, and
  rejects an element a container cannot hold, naming what was passed.
- `text_of` (CR-001) is docx4j's `TextUtils` over any subtree and decides its own text-box rule
  under CR-002 question 7; `Body.text` is Word's main story, the accepted view, no text boxes.

## 4. What the TypeScript implementation learned, adopted here from the start

Each item cost docx4j-core-ts a correction or a departure note (its CR-002 sections 7 to 15).
They are rules here.

- **`style` is the display name, `style_built_in` the `Word.Style` value.** `style` reads the
  styles part's `w:name` when that part is unmarshalled (nothing is unmarshalled for it), with
  Word's stored lower-case built-in names (`heading 1`, `toc 1`, `annotation text`) mapped to
  display names through docx4j's `KnownStyles.xml`; otherwise the id with spaces inserted.
  Setting accepts a display name, a stored name or an id. `style_built_in` reads `"Other"` for a
  style that is not built in and refuses to be set to it. `style_id` is the docx4j-named
  extension; `outline()` reports ids. `BUILT_IN_STYLES` is one committed list.
- **`insert_xml`, `insert_ooxml`, `insert_markdown` and `insert_element` return the inserted
  views**, since a fragment may bring several blocks. At paragraph and range level a fragment of
  exactly one `w:p` has its runs merged into the paragraph, as Word's paste does.
- **`Body.tables` returns `Table` views** from the first phase, never elements.
- **Rows and cells descend into row- and cell-level content controls** (an OpenDoPE repeat wraps
  its `w:tr` in a `w:sdt`), so `rows` is what Word shows; `add_rows` puts new rows in the table's
  own content. Children of every `w:sdt` form are visible to `Body.paragraphs` and to addresses,
  and the four `sdtContent` levels are skipped in paths.
- **`header_row_count` counts leading rows** carrying `w:tblHeader`; `insert_table` sets no
  style unless `style=` is given, so a new table is borderless until `style_built_in = "TableGrid"`.
- **Pictures**: `/word/media/imageN.<ext>` with N free for that extension (docx4j's
  `getNewPartName`); a relationship from the part the body belongs to; the `wp:inline` from
  `inline_picture`; PNG, JPEG, GIF and BMP headers read for size and density, 96 dpi default,
  EMU as `px / dpi * 914400`, an image wider than the text area scaled down as `CxCy.scale`
  does. No Pillow dependency: the four header readers are a hundred lines.
- **`insert_ooxml` from a `pkg:package`**: every part the content references is copied under a
  free name with a fresh relationship id, its own relationships copied recursively keeping their
  ids; references in the inserted content are rewritten by attribute name and only when the
  incoming package really has a relationship of that id, so numeric ids are never touched; the
  walk enters wildcard content (`walk_all`). Styles and numbering are not merged.
- **`ContentControl.get_range()`** is exact for a run-level control and the first paragraph's
  range for the block, row and cell forms. `type` is `"RichText"` when `w:sdtPr` names no kind,
  `w:text` is `"PlainText"`. A row- or cell-level control refuses `insert_paragraph` at
  `"Start"` / `"End"` and `insert_text(..., location="Replace")`, naming what it holds.
  `insert_content_control` at range level refuses a span that crosses a run holder;
  `RepeatingSection` is refused at run level.
- **Comments**: one `Comment` type for comments and replies (`parent` is an extension); `id` is
  the OOXML `w:id`, an `int`; `creation_date` may be `None`; `get_range()` is a `Range` per
  paragraph and a comment whose markers sit at block level gets none; an empty range gets a
  reference run only; no `w16cex` part is created, only kept in step; the styles part is
  unmarshalled only when it lacks the two comment styles.
- **Tracking rules, Word's not just the markup**: a run already inside a `w:ins` by the same
  author is extended rather than nested; deleting text that author inserted takes it back;
  `w:del` before `w:ins` on a replacement; `w:t` becomes `w:delText` and `w:instrText`
  `w:delInstrText`; runs split at the span's boundaries. Revision ids come from a per-package
  counter above the highest `w:id` on any `CTMarkup` in the parts unmarshalled, **excluding**
  `w:comment/@w:id`, a separate space with its own allocator. A comment is not a revision: its
  markers are hoisted out of a `w:ins` or `w:del` the anchor run sits in. Deleting a content
  control is not tracked. `TrackMineOnly` is stored as `TrackAll`. An inserted paragraph carries
  its own mark. A deleted row stays in the tree until accepted. `accept_all()` and
  `reject_all()` are one pass in reverse document order and, unlike docx4j's conversion
  preprocessor, accepting also drops `w:rPrChange` and `w:pPrChange`. Editing deleted text
  raises, naming the author.
- **Bindings**: a bound control's `insert_text` writes through to the custom XML node, because
  Word refreshes a bound control from the data on open; containers are never bound; pictures and
  explicit rich text are counted as skipped; dates are formatted with `w:dateFormat` in the
  `w:lid` locale and `w:fullDate` set, a checkbox gets `w14:checked` and the glyph run from
  `w14:checkedState` / `uncheckedState`, a list shows the `display_text` whose `value` matches;
  the reverse writes `true` / `false`, the entry's value and the stored `w:fullDate`. Content
  goes into the **first paragraph** of a block, row or cell control; `w:placeholder` is kept;
  bindings are read from `w:dataBinding` **and** `w15:dataBinding`; the three well-known
  docProps store item ids are not special-cased; `placeholder_text` refuses to overwrite
  content; `add()` follows docx4j's `addPropertiesPart`, with the relationship **from the main
  document part**.
- **`Paragraph.font` is over the runs**, not the paragraph mark; `alignment` is `"Unknown"` when
  `w:jc` is absent; `outline_level` is `w:outlineLvl + 1`, 10 when absent.
- **A `Range`'s original view is not offered** (`get_text(view="original")` raises on a range),
  and `search` has no `view` option.
- **`to_api_script` fidelity**: a `w:tab` is `insert_text("\t")`; `w:proofErr`,
  `w:lastRenderedPageBreak` and bookmarks are dropped; a `w:br` is a run of its own; anything a
  verb cannot express falls back to `insert_xml`.

## 5. Where each piece lives

| Piece | Package | Why |
|---|---|---|
| `tr`, `tc`, `inline_picture`, the `sdt` family, `rpr_to_elements`, `deep_copy_as`, `walk_all`, `run_items_of` | `docx4j_py.wml.builders`, `docx4j_py.child`, `docx4j_py.traversal` | tree only |
| `Body`, `Paragraph`, `Range`, `Font`, `Table`, `TableRow`, `TableCell`, `InlinePicture`, `ContentControl`, the aliases, `insert_xml`, `insert_ooxml`, the text model | `docx4j_py.model.content` | needs parts |
| addresses, `Outline`, `SearchHit`, `describe()`, `ChangeReport`, `dry_run`, the error hierarchy | `docx4j_py.model.content.addresses`, `.reports`, `.errors` | the agent surface, over the views |
| `DocumentSession` | `docx4j_py.model.sessions` | every server would write it |
| markdown out and in | `docx4j_py.model.markdown` | over the views; `markdown-it-py` for parsing |
| the python-docx facade | `docx4j_py.docx` | a separate namespace with python-docx's names |
| `CustomXmlPart`, `CustomXmlNode`, `XmlMapping`, the typed kinds, bindings, `describe()` / `fill()` | `docx4j_py.model.customxml` | needs parts |
| `ChangeTracker`, `TrackedChange` | `docx4j_py.model.content.tracking` | the primitives call it |
| `Comment` and the comment parts | `docx4j_py.model.content.comments`, `docx4j_py.openpackaging.parts.wml.comments` | half view, half part |
| `List`, `ListItem` | `docx4j_py.model.content.lists` | needs the numbering part |
| `to_api_script` | `docx4j_py.model.content.api_script` | over the views |
| the two subset lists and their tests | `tests/office_js_subset.json`, `tests/python_docx_subset.json` | the compatibility promises live with the implementation |
| MCP tools | a `docx4j-mcp-py` product | product, not library; 3.4 says what it is a few lines over |

Import cycles: the parts layer never imports the content API statically. `XmlPart.body` is
provided through a registration the content module performs on import; the content module
imports the parts, never the reverse.

## 6. What CR-002 Phase B supplies, and what this CR does without it

`Font` reads report direct formatting of the first run in scope until `PropertyResolver` lands,
then effective formatting as Office JS reports it; documented as such, with a test marked to
flip. `ListItem.list_string` is `None` until the `Emulator`, and `to_markdown` numbers lists
itself. `insert_table` sizes the grid from `w:sectPr`, which needs nothing from Phase B.

## 7. Tests

- **Fixtures:** the 16 documents in `samples/` (`invoice2013.docx` already covers w15 repeating
  sections, a checkbox, a date, a picture control and twenty bindings), plus, copied from
  docx4j's samples as each phase needs them, a document with comments and replies, one with
  tracked changes, one with lists, one with text boxes, one with a hyperlink and a field.
- **Per phase**, in `tests/content/`: reads against known values; every mutation followed by
  save, reload and read back; untouched parts still byte-identical through the content API;
  `to_api_script` output executed against a fresh body and compared to the source.
- **Agent scenarios**, in `tests/agent/`: scripted sessions of tool-shaped calls (`outline`,
  `find`, edit by address, `ChangeReport` checked, `dry_run` then commit, tracked and commented
  edits accepted in a reload) over the corpus, and a token-budget test that every `to_dict()`
  result for a 200-page document stays under a stated size at the default budgets.
- **The two subset tests** (3.6 and the Office JS list of section 5).
- **Parity with docx4j-core-ts:** the same fixture, the same call sequence, the same saved
  `document.xml` canonically, run when the sibling checkout and Node are present; a difference
  is a bug in one engine or a departure to record in both CRs.
- **Word acceptance**, manual, per phase in `tests/README.md`: a content edit, an inserted
  picture, a comment thread, a tracked change accepted in Word, a bound control refreshed from
  its data, a markdown-built document.

## 8. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | Tree layer additions (3.3) — **implemented 2026-09-16, section 10** | 2 days |
| B | `Body`, `Paragraph`, `Range`, `Font`: the verbs, `style` per section 4, `search` across runs with grapheme-safe splitting, `insert_xml`, `insert_element`; the error hierarchy; the Office JS subset list — **implemented 2026-09-16, section 11** | 4 days |
| D | The agent surface (3.4): addresses, `Outline` with budgets, `find()`, `describe()`, `ChangeReport`, `dry_run`, `DocumentSession`, determinism; the agent scenario tests — **implemented 2026-09-16, section 12** | 4 days |
| K | Markdown out, with addresses, then in (3.5) --- **implemented 2026-09-17, section 13** | 3 days |
| C | `Table`, `TableRow`, `TableCell`, `InlinePicture` with the header readers, `insert_ooxml`, `ContentControl` reads and `delete` --- **implemented 2026-09-17, section 14** | 4 days |
| G | Comments (3.9) --- **implemented 2026-09-17, section 15** | 3 days |
| F | Change tracking and `replace_text` (3.8); the README's audit-trail example --- **implemented 2026-09-17, section 16** | 4 days |
| E | Custom XML, mapping, typed controls, `describe()` / `fill()` (3.7) | 4 days |
| J | The python-docx facade (3.6) and its subset test | 3 days |
| H | Lists (3.10) | 3 days |
| I | `to_api_script` (3.11) | 2 days |

B, D and K first: after them an agent can read any document within a budget, address any block,
edit it and see what changed, which is the MCP server's whole first release. G before F, as the
TypeScript engine found (tracking has to hoist comment markers). J after E so the facade sits on
a complete core. Each phase ships as a minor version.

## 9. Open questions (decided 2026-09-16)

All eleven recommendations below were accepted on 2026-09-16 and are now decisions.

1. **`snake_case` with Office JS values, or camelCase members for literal parity.** Recommendation:
   `snake_case`, Office JS string values, `StrEnum`s beside them; no camelCase aliases.
2. **Default locations** (`"End"` on a body and a paragraph's `insert_text`, `"After"` on a
   paragraph's `insert_paragraph`, `"Replace"` on a range). Recommendation: yes; Office JS has
   no defaults only because JavaScript positional arguments made them awkward, and the short
   call is what a model writes correctly.
3. **Code-point offsets.** Recommendation: yes; only `Range` carries offsets.
4. **`ChangeReport` on every call, or opt in.** Recommendation: always recorded, cheap (a few
   fields per call), cleared by the caller; the server needs it on every call anyway.
5. **`DocumentSession` in the library or in the server.** Recommendation: the library, since
   the thread rule and the idle-close logic are the library's knowledge and every server needs
   them.
6. **`markdown-it-py` as a dependency.** Recommendation: yes, for `insert_markdown` only,
   imported lazily; `to_markdown` needs nothing. (Done in Phase K, and no second one:
   `mdit-py-plugins` and `linkify-it-py` were both declined, which is what puts GFM footnote
   import out of scope --- section 13.4.)
7. **The python-docx facade: subset test against a pinned python-docx version, or run
   python-docx's own tests.** Recommendation: the subset list; python-docx's tests exercise its
   lxml internals, which are not the promise.
8. **`Body.text` and text boxes.** Recommendation: Word's main story, no text boxes, the accepted
   view; `text_of` decides its own rule under CR-002 question 7.
9. **A `Word` shim.** Recommendation: no; `to_api_script` only, and the python-docx facade is
   Python's "existing code runs".
10. **Grapheme clusters without a dependency.** Recommendation: `unicodedata` plus the ZWJ and
    variation-selector list, `regex`'s `\X` when installed.
11. **The Python MCP server.** Recommendation: a separate product repository once Phase K lands,
    with docx4j-mcp's tool names for the coarse tools, its path policy and inline caps, and the
    in-place tools of 3.4 as its second release; not this CR.

## 10. Phase A implementation notes (2026-09-16)

Phase A is done: the tree-layer additions of section 3.3, their tests, the acceptance artefact
and the README. The suite is **696 tests** (687 of them fast), against 581 at the end of CR-002
Phase A; **115** of the new ones are `tests/test_builders_phase_a.py`. One schema patch and one
regeneration were needed (10.2); nothing in `~/git/docx4j-xsdata` changed, so the fork is still
at CR-002 Phase A's commit.

```python
from docx4j_py.wml import (tr, tc, inline_picture, image_size, emu_for,
                           sdt, sdt_pr, sdt_property, sdt_kind_of, next_sdt_id,
                           rpr_to_elements, rpr_from_elements,
                           deep_copy_as, walk_all, run_items_of)
```

### 10.1 What landed

| Piece | Where | Signature |
|---|---|---|
| rows and cells | `docx4j_py.wml.builders` | `tr(cells, *, widths=None, header=False)`, `tc(blocks="", *, width=None, span=None)` |
| the picture | `docx4j_py.wml.pictures` (new) | `inline_picture(rel_id, *, cx, cy, id=1, name=None, descr="", title=None, pic_id=0, pic_name=None, link=False) -> Drawing` |
| the image headers | same | `image_size(data) -> ImageInfo(width_px, height_px, dpi_x, dpi_y, format, has_density)`, `emu_for(info, *, max_width_emu=None) -> EmuSize(cx, cy, scaled)` |
| content controls | `docx4j_py.wml.sdt` (new) | `sdt(content, *, kind="RichText", tag=None, title=None, id=None, form=None, alias=None, lock=None, placeholder=None)`, `sdt_pr(**the same options)`, `sdt_property(sdt_pr, local_name, namespace=W_AND_W15)`, `sdt_kind_of(sdt_pr) -> str`, `next_sdt_id(root) -> int` |
| run properties as elements | `docx4j_py.wml.builders` | `rpr_to_elements(rpr) -> list[RPrElement]`, `rpr_from_elements(elements, *, cls=RPr)` |
| the re-typing copy | `docx4j_py.child` | `deep_copy_as(obj, cls, parent=None)` |
| the wildcard walk | `docx4j_py.traversal` | `walk_all(root, visitor, wildcard_visitor=None, *, context=None, mce="all")` |
| the run list | `docx4j_py.traversal` | `run_items_of(holder) -> list | None` |

`builders.py` grew to 730 lines, so the picture and the content control went into sibling modules
as section 3.3's deliverable allows; every public name is on `docx4j_py.wml` all the same, because
the generator's package footer now re-exports the `__all__` of **every** hand-written module in a
namespace package rather than of `builders.py` alone (`codegen/generate_el.py`'s `HAND_WRITTEN` is
the one list of what the generator does not own, and `codegen/clean.py` keeps the same files).

`BuilderError(ValueError)` carries the `code` and `hint` section 3.1 asks of every error
(`sdt.form_mismatch`, `picture.empty_extent`, `image.unsupported_format`, `rpr.unknown_property`,
…). **Phase B must re-root it** under `Docx4JError` / `ContentError` when that hierarchy exists;
deriving from `ValueError` is what keeps the existing `pytest.raises(ValueError)` call sites
working until then, and the `code` strings are meant not to change.

### 10.2 What the generated model turned out to be

Three of the five things section 3.3 describes are shaped differently here from the TypeScript
model, and two of the five helpers are therefore thinner than the CR expected.

**`w:sdtPr` is one choice list, as the CR says.** `SdtPr` has a single compound `content`
field of 30 alternatives (`rPr`, `alias`, `lock`, `placeholder`, `showingPlcHdr`, `dataBinding`,
`temporary`, `id`, `tag`, the ten kind elements, and the w14/w15 ones: `w14:checkbox`,
`w14:entityPicker`, `w15:appearance`, `w15:color`, `w15:dataBinding`, `w15:repeatingSection`,
`w15:repeatingSectionItem`, `w15:webExtensionCreated`, `w15:webExtensionLinked`). It is docx4j's
`getRPrOrAliasOrLock()` in Python, so `sdt_property` and `sdt_kind_of` are the accessors the CR
called for. They are written over `iter_children`, which reads the qualified name of each
alternative out of the field's `choices` metadata, so the two `dataBinding` entries — `w:` is
`CTDataBinding`, `w15:` is `W15CtdataBinding` — are told apart by the model rather than by a
list here. The four container classes are `SdtBlock`, `SdtRun`, `CTSdtRow` and `CTSdtCell`, each
with `sdt_pr`, `sdt_end_pr` and `sdt_content`, and the four content classes are
`SdtContentBlock`, `CTSdtContentRun`, `CTSdtContentRow` and `CTSdtContentCell`. Only `SdtBlock` is
a global element declaration, so only it has an element name of its own: `to_xml` on the other
three needs `name="{…}sdt"`, which the tests do.

**`w:rPrChange/w:rPr` is *not* an element list here.** Section 3.3 expected the base-type element
list the TypeScript model has (`CTRPrChange.RPr.egrPrBase`). xsdata generated
`CtRprChangeRPr` — and `CTParaRPrOriginal`, for `w:pPr/w:rPr/w:rPrChange/w:rPr` — with **named
fields in `EG_RPrBase` order, each one a list**, because the group is unbounded there; `RPr` and
`ParaRPr` have the same names as single values. So:

- **`rpr_to_elements` and `rpr_from_elements` are thin**, as the deliverable allowed for: a
  field-by-field copy, with the single/list difference absorbed. `RPR_BASE_FIELDS` is computed
  from the model (the fields `RPr` and `CtRprChangeRPr` share, in `RPr`'s declaration order) and
  is 51 pairs, the twelve w14 text effects included and `w:rPrChange` excluded.
- **Order is not this code's problem.** Because they are named fields, xsdata writes them in
  declaration order, which is the schema's, whatever order the caller supplies. The TypeScript
  has to sort; Python cannot get it wrong.
- `rpr_to_elements` returns `RPrElement(name, qname, value)` rather than bare property objects,
  because the value alone does not identify the element: `w:b`, `w:i`, `w:caps` and fifteen more
  are all a `BooleanDefaultTrue`. `rpr_from_elements` also accepts `(name, value)` pairs and a
  whole `w:rPr` object of any of the four shapes.

**Every run holder keeps its children under the same field name, `content`.** The TypeScript has
to know three property names (`content`, `customXmlOrSmartTagOrSdt` for `w:ins`/`w:del`,
`accOrBarOrBox` for `w:moveFrom`/`w:moveTo`), and CR-001 section 14 warned this would differ per
holder. It does not: `P`, `PHyperlink`, `RunIns`, `RunDel`, `MoveFrom2`, `MoveTo2`,
`CTSmartTagRun`, `CTCustomXmlRun`, `CTSimpleField`, `CtDir`, `CtBdo` and the four
`w:sdtContent` classes all use `content`. `run_items_of` is therefore one element-name test
(`RUN_HOLDERS`) plus a reach through `sdt_content` for a `w:sdt`, and it is keyed on element
names rather than classes so that it costs `traversal.py` no import of the model.

**`deep_copy_as` has real work to do.** `CTPPrChange.p_pr` is declared `PPrBase` and `PPr`
extends it, so `pPrChange.p_pr = deep_copy(ppr)` marshals `<w:pPr xsi:type="w:CT_PPr">`, which is
valid and is not what Word writes. `deep_copy_as(ppr, PPrBase)` copies only the fields `PPrBase`
declares — dropping `w:rPr`, `w:sectPr` and `w:pPrChange`, which a `w:pPrChange` must not hold
anyway — and the `xsi:type` and the `xsi` declaration both go. A test asserts both halves.

### 10.3 Departures from section 3.3

1. **`inline_picture` takes `title`, and that needed a schema patch.** `CT_NonVisualDrawingProps`
   in this repository's schema copy had no `title` attribute. docx4j's own `xsd/` has carried it
   since docx4j's CR-018 item 2 (ECMA-376 4th edition Part 1 20.1.2.2.8, Transitional), and it
   was the *only* difference `diff -r schemas/dml ~/git/docx4j/xsd/dml` reported, so the patch
   copies it across rather than inventing anything (`schemas/PATCHES.md` 5). The regeneration it
   caused is seven lines, `codegen/generate.sh --check` passes, and no element name moved.
2. **`inline_picture` takes `pic_id` and `pic_name` as well.** docx4j's `createImageInline` has
   two ids (`id1` for `wp:docPr`, `id2` for `pic:cNvPr`) and one `filenameHint` for both names;
   Word writes `0` for the second id and puts the file name in `pic:cNvPr/@name` while
   `wp:docPr/@name` says "Picture 1". The defaults are docx4j's behaviour (`pic_id=0`,
   `pic_name` following `name`); the keywords exist so that a caller can write exactly what Word
   writes, which is what makes the comparison against `samples/Images.docx` an exact one. `link=`
   is docx4j's `link` argument, `r:link` instead of `r:embed`.
3. **`sdt(id=None)` writes no `w:id`.** The TypeScript gives a control a random 31-bit id. That
   would make `to_xml` non-deterministic, which section 3.4 forbids ("same document, same calls,
   same bytes"), so the id is omitted unless one is given and `next_sdt_id(root)` is the way to
   get one. `next_sdt_id` is likewise deterministic — one above the highest `w:id` in the tree,
   and the lowest free id if that would overflow `ST_DecimalNumber` — where the TypeScript
   randomises until it misses. Word assigns an id of its own to a control that has none.
4. **`sdt_property`'s default namespace is `w:` *and* `w15:`** (`W_AND_W15`), where the
   TypeScript defaults to `w:` alone and makes the caller ask for both. Section 3.3 asks for the
   two, and the reason is in section 4: 3 of the 20 bindings in `samples/invoice2013.docx` are
   `w15:dataBinding`. `W_NS`, `W14_NS`, `W15_NS` and `ANY_NS` are exported for the other cases.
5. **`sdt` and `sdt_pr` also take `lock` and `placeholder`**, written between the `w:id` and the
   kind element; and `alias` is accepted as a spelling of `title`, since the element is `w:alias`
   and Word's dialog says "Title" (giving both different values raises).
6. **`tr` takes `header=`** (`w:trPr/w:tblHeader`), which section 3.3 does not mention and
   `Table.header_row_count` in section 4 needs something to have written.
7. **`tc` gives an empty cell a `w:p`.** `tbl` used not to: `tbl([[[]]])` produced a `w:tc` with
   no paragraph, which makes Word repair the document. Every existing `tbl` test passes
   unchanged; this is the one output difference and it is a fix.
8. **`walk_all` defaults to `mce="all"`** where `walk` defaults to `"resolve"`, because a caller
   that rewrites references has to reach the branch a consumer would ignore: it is written back
   on save. `wildcard_visitor` is optional, so `walk_all(root, visitor)` is "`walk` that descends
   through wildcards".
9. **`emu_for` returns `EmuSize(cx, cy, scaled)`**, a frozen dataclass with `to_dict()`, not a
   tuple (section 3.1), and takes `max_width_emu` rather than docx4j's `PageDimensions`: the
   text-column width is the caller's knowledge, and the tree layer has no `w:sectPr`. The
   arithmetic is `CxCy.scale`'s: `px / dpi * 914400` per axis, the natural size when it fits, and
   otherwise the maximum width with the height in proportion.
10. **`image_size` reports `has_density`**, so a caller can tell a real `pHYs`/JFIF density from
    the 96 dpi default rather than guessing. JPEG density is read from the `JFIF` `APP0` segment
    only (units 1 dpi, 2 dots per cm), not from EXIF, which is what docx4j's reader does too.

### 10.4 What Phase B must know

- **The error hierarchy.** Re-root `BuilderError` under `Docx4JError` / `ContentError` and keep
  its `code` strings; `hint` is already the sentence section 3.1 wants.
- **Nothing here knows about parts.** `inline_picture` takes a relationship id as a string;
  making the `ImagePart`, finding a free `/word/media/imageN.<ext>` and adding the relationship
  are Phase C's, and `scripts/acceptance.py`'s `created_with_image` is the worked example of the
  three steps in order.
- **The text column.** `emu_for(..., max_width_emu=)` needs a number Phase B can compute from
  `w:sectPr` (`pgSz/@w` minus `pgMar/@left` and `@right`, twips × 635 EMU); the acceptance
  script hard-codes A4's 5,731,510 EMU until it can.
- **Parents of single-valued fields.** A constructor cannot link them (only `ChildList` adopts),
  so every builder here calls `link_parents` on what it returns. A view that assigns
  `paragraph.p_pr = …` must do the same, or call `link_parents` afterwards.
- **`run_items_of` is structural**: it returns the live list with nothing filtered, `w:del` and
  `w:moveFrom` included, because the original view of a revision needs them. The filtering stays
  in `text_of`.
- **`sdt_kind_of` returns `"RichText"` for an untyped control** and never `"Unknown"`, which is
  section 4's rule and what `ContentControl.type` should report.
- **`to_xml` on a run-, row- or cell-form `w:sdt` needs `name=`** (only `SdtBlock` is a global
  element declaration); `wml(xml, wrapper=…)` resolves the same ambiguity on the way in —
  `"p"` gives `SdtRun`, `"tbl"` gives `CTSdtRow`, `"tr"` gives `CTSdtCell`.
- **`DelText` is still not a `Text`** (CR-001 section 14.8 point 6), so a run-level filter that
  wants both must name both.

## 11. Phase B implementation notes (2026-09-16)

Phase B is done: `Body`, `Paragraph`, `Range` and `Font` (section 3.2), the style semantics and
the other rules of section 4 that they touch, search across runs with the grapheme-safe splitting
of section 3.12, `insert_xml` and `insert_element`, the error hierarchy of section 3.1, the
Office JS subset list of section 5, the tests of section 7 and the README. The suite is **776
tests** (775 fast plus one `xfail`, 45 s without the corpus round trip, 52 s with it), against
696 at the end of Phase A; **80** of the new ones are `tests/content/`. Nothing in
`~/git/docx4j-xsdata` changed, no schema patch was needed and the model was not regenerated;
`codegen/generate_el.py` changed only in its engine footer (11.5), and the footer it now writes is
byte-identical to what is committed.

```python
from docx4j_py import create_package

pkg = create_package()
pkg.body.insert_paragraph("Hello World")
pkg.save("hello.docx")
```

### 11.1 What landed

| Piece | Where | Notes |
|---|---|---|
| the error hierarchy | `docx4j_py/model/content/errors.py` | `Docx4JError` (CR-002's `Docx4JException`, the same class) → `ContentError(code, message, hint)` → `AddressError`, `InvalidTargetError`, `StyleError`, `SpanError`, `BuilderError`; each has `to_dict()` |
| the text model | `.text_model` | `Segment`, `segments_of`, `text_of_view`, `runs_of`, `split_at`, `set_text`, `block_list_of` / `block_children_of`, `search_pattern` / `find_all`, `is_grapheme_boundary` / `snap_back` / `snap_forward` / `grapheme_clusters` |
| the styles | `.styles` | `BUILT_IN_STYLES` (171 `BuiltInStyle(style_id, name, built_in, type, stored_name)`), `WORD_STYLE_VALUES` (the 50 `Word.Style` values, equal to the TypeScript engine's list), `built_in_of`, `id_of_built_in`, `display_name_of`, `style_name_of`, `style_id_of` |
| the values | `.enums` | `Location`, `BodyLocation`, `TextLocation`, `ParagraphLocation`, `RangeLocation`, `TextView`, `AlignmentValue`, `UnderlineValue`, `BreakTypeValue` as `Literal`s, and `InsertLocation`, `Alignment`, `UnderlineType`, `BreakType` as `StrEnum`s beside them |
| the views | `.body`, `.paragraph`, `.range`, `.font` | section 3.2's members that Phase B owns, plus `Block` for a block-level child that is not a paragraph (a `Table` view is Phase C) |
| the registration | `.__init__` | `XmlPart.body` and `WordprocessingMLPackage.body`, installed on import; the parts layer imports nothing from here |
| determinism | `OpcPackage.id_seed`, `.id_generator()` | section 3.4's seedable per-package generator; a new paragraph gets a `w14:paraId` from it when the document already uses them |
| the subset list | `scripts/office_js_subset.py`, `tests/office_js_subset.json` | 24 interfaces, 203 members, derived from docx4j-core-ts's `test/office-js-subset.ts` |

`Body` is one class over **any** container with a block-level list, because in this model a
`w:body`, a `w:hdr`, a `w:ftr`, a `w:tc`, a `w:tr`, a `w:tbl` and all four `w:sdtContent` classes
keep their children under the same field name, `content` (Phase A's finding, section 10.2,
turned out to reach further than the run holders). `block_children_of` is therefore the whole of
the descent: `Body.paragraphs` walks into tables, rows, cells and every `w:sdt` form with one
three-line visitor and no per-class code, and `Body.sub(container, prefix)` is a constructor call.
The three notes parts are the exception and they cost one rule: `w:footnotes` keeps its notes in
`footnote` and `w:comments` its comments in `comment`, so a container with no `content` list
answers with its single repeated element field **when the items of that field are themselves
block containers** — which is what tells `w:footnotes` (whose `w:footnote` has a `content` list)
apart from `w:styles` (whose `w:style` has not), and is why `styles_part.body` raises
`ContentError("body.no_content")` rather than handing out a body of style definitions.

`insert_element` validates against the `XmlContext` metadata rather than a list written here: the
container's block field is looked up, its compound `elements` (30 alternatives for a `w:body`)
are the accepted element names, and the message names both halves —

```
w:body cannot hold w:r (R); it takes w:customXml, w:sdt, w:p, w:tbl, w:proofErr, w:permStart,
w:permEnd, w:bookmarkStart and 22 more (wrap a run in a w:p, or insert it into a paragraph
with paragraph.insert_xml())
```

so a container this code has never heard of validates correctly, and the accepted list cannot go
stale against the schema.

### 11.2 Departures from section 3.2 and section 4, all deliberate

1. **`MainDocumentPart.body` is now the view; the element is `body_element`.** CR-002 gave the
   part a `body` property returning the typed `w:body`. Section 3.2 wants `part.body` to be the
   `Body` view, and section 5 wants it registered from the content module, so the element kept the
   longer name and `part.contents.body` is the same object. Nothing in the repository used the old
   name.
2. **`Body.tables`, `content_controls`, `inline_pictures`, `insert_table` and the picture verbs
   are not here.** Section 4 says `Body.tables` returns `Table` views "from the first phase", and
   it does — `Table` is Phase C, so Phase B offers none rather than the elements the TypeScript
   engine handed out and had to correct. `body[3]` for a table gives a `Block(element, container)`,
   which Phase C replaces with a `Table` without changing what `body.paragraphs` reports.
3. **`Font` keeps `double_strike_through` and `style`**, which Office JS's `Word.Font` has (the
   first) and does not (the second, `w:rStyle`); both were already in CR-001's
   `apply_run_options` / `read_run_options` vocabulary, so leaving them off `Font` would have made
   the one mapping two. They are marked as extensions in the docstrings.
4. **`Font.size = 0` removes `w:sz` and `w:szCs`.** Section 3.2 does not say, but `read_run_options`
   reports `0` for a run that sets no size, and writing back what was read must not ask Word for a
   zero-point font. `apply_run_options` grew the rule (the one change to Phase A's code beyond the
   error re-rooting), which also means `r(text, size=0)` now clears rather than writes.
5. **`insert_text` on a `Body` with no paragraphs makes one.** Office JS raises; CR-003 section 3.1
   says the common call is short, and "insert text into an empty document" is the commonest call
   an agent makes first.
6. **`Range.font` splits and does not merge**, as the deliverable allows: two adjacent runs whose
   `w:rPr` ends up identical stay two runs. Word tolerates it and a `to_api_script` phase can tidy.
7. **A `Range`'s error type is `SpanError`, not `RangeError`.** Python has no builtin of that name
   and `Range` here is the view, so the error is named for what it is about. The TypeScript name is
   in its docstring.
8. **`get_text(max_chars=)` truncates and does not say so.** The dataclass that reports
   `truncated=True` (section 3.4) is Phase D's `Outline`; `get_text` returns a `str` and there is
   nowhere to put the flag until then.
9. **`paragraph.insert_xml` defaults to `location="After"`**, matching `insert_paragraph`; the
   merging behaviour section 4 asks for is on `"Start"` and `"End"`, which have to be asked for.
10. **Setting a style unmarshals the styles part; reading one never does.** Section 4 requires the
    display name to come from `w:name` "when that part is unmarshalled", and separately requires a
    `StyleError` listing the five closest names for a style the document does not define. The
    second cannot be done without reading the part, so `style_id_of(..., validate=True)` unmarshals
    it and the setter uses that; every read path passes `validate=False` and a test asserts that
    reading `paragraph.style` leaves `styles.xml` untouched. A name that resolves to a **built-in**
    style id is written as it stands whether the document defines it or not, because Word creates
    the definition when it opens the file.
    **Corrected 2026-09-17:** Word does *not* create the definition --- it renders the paragraph
    as Normal --- so a built-in style the document does not define is now defined by the setter.
    See section 14.9.
11. **`w:delText` is read, in one place.** Section 3.2's text model excludes it, and it is excluded
    — except inside the `w:del` or `w:moveFrom` that removed it, which is the only thing
    `get_text(view="original")` can be made of. A `w:delText` outside a revision still contributes
    nothing.
12. **`read_run_options` reported a parsed `w:color` wrongly**, and Phase B fixed it: `ST_HexColor`
    is a union of `auto` and `xsd:hexBinary`, so a parsed `w:color/@w:val` is `bytes` and a built
    one is the `str` the builder wrote. Both now read back as `"#RRGGBB"`. This was a CR-001 defect
    that only a view over a *loaded* document could find.

### 11.3 Grapheme clusters without a dependency (section 3.12, decided question 10)

`is_grapheme_boundary(text, index)` is `unicodedata`'s general categories and combining classes
plus the five cases they do not cover: the zero-width joiner binds both ways (which is what holds
`U+1F468 ZWJ U+1F469 ZWJ U+1F467` together), a variation selector and an emoji skin-tone modifier
never begin a cluster, two regional indicators make one flag so only every second one is a
boundary, a virama (canonical combining class 9) binds the consonant after it so `स्ते` is one
cluster, and `CR LF` is one cluster. `grapheme_clusters` uses the `regex` module's `\X` when that
module happens to be installed and this function otherwise; `regex` is **not** installed in
`.venv-fork`, so the committed tests exercise the dependency-free path.

`split_at(offset, prefer="back"|"forward")` snaps before it splits and **returns the offset it
used**, which is the piece the CR did not specify and every caller needs: `Range.font` snaps its
start backwards and its end forwards, so formatting a span always covers whole clusters.
`Paragraph.splice` deliberately does *not* snap, so `replace_text` replaces exactly what `search`
matched.

### 11.4 The Office JS subset list, and how the mapping works

`scripts/office_js_subset.py` parses `docx4j-core-ts/test/office-js-subset.ts` — the TypeScript
engine's own compile-time assignability check — rather than restating it, so the two engines
cannot drift apart in silence. It writes `tests/office_js_subset.json`: 24 interfaces, 203
members, each with its Office JS name, its Python name, `property` or `method`, `readonly`, the
CR-003 phase that owns it, whether it is an extension, and the doc comment. `--check` fails if the
committed file is out of date, and a test runs it when the sibling checkout is present.

The mapping is mechanical `snake_case` (`styleBuiltIn` → `style_built_in`, `getRange` →
`get_range`, `insertInlinePictureFromBase64` → `insert_inline_picture_from_base64`) with exactly
**two** exceptions, both of them section 3.2's and both recorded in the JSON's own `aliases`:

| Office JS | Here | Why |
|---|---|---|
| `Table.getCell(rowIndex, cellIndex)` | `cell(row_index, cell_index)` | python-docx's name for the same thing (section 3.2), and section 3.1's "no `get_` prefix" for what is a lookup |
| `InlinePicture.getBase64ImageSrc()` | `get_base64()` | its `get_bytes()` twin is what a Python caller wants (section 3.1: bytes are never only reachable through base64) |

Three further shape differences are in the JSON as data rather than as aliases. Office JS's
`SearchOptions` is an *object*; here the three options are keyword arguments of `search`, so the
test asserts they are keyword-only parameters of `search_pattern` instead of members of a class.
Office JS's collections are proxies with `items`; here they are `list`s. And Office JS's
`load`/`sync`/`context` have no counterpart at all, which is why the TypeScript file removed them
before this script ever saw it.

The test asserts that every non-extension member Phase B owns exists on `Body`, `Paragraph`,
`Range` or `Font` with the right kind, and its assertion message lists the members of later phases
still missing — **157 of them, C: 52, E: 79, F: 13, G: 13** — so a phase that lands without
touching its members is caught by the message rather than by a failure.

### 11.5 The import direction, and the one place the two layers meet

Section 5 says the parts layer never imports the content API and that `XmlPart.body` is provided
by a registration the content module performs on import. Both halves hold, and the second needed
somewhere for the registration to be triggered from: `docx4j_py/__init__.py`, which is neither
layer and has just finished importing the model both rest on, now ends with
`import docx4j_py.model.content`. It costs the engine's import — milliseconds on top of the
model's second — and `codegen/generate_el.py`'s engine footer writes it, so a regeneration keeps
it. `docx4j_py/model/content/__init__.py` imports only `errors` eagerly and reaches `body_of`
through a lazy property getter, so registering costs nothing and `docx4j_py.wml.builders` can take
`BuilderError` from `docx4j_py.model.content.errors` (whose only import is
`docx4j_py.openpackaging.exceptions`) without the views importing the builders back.

Four names joined the top-level lazy exports: `ContentError`, `Paragraph`, `Range` and `Font`.
**`Body` did not**, because `docx4j_py.__all__` already promises that name for the model's
`w:body` class, and a body is reached as `pkg.body` rather than imported.

One wart found on the way, recorded rather than fixed: `docx4j_py.wml.sdt` is both a module and
the builder that module exports, so `from docx4j_py.wml import sdt` gives whichever was bound
first — the lazy re-export if nothing has imported the module, the module if something has. It is
CR-003 Phase A's and it is not new; `from docx4j_py.wml.sdt import sdt` is unambiguous. A later
phase should decide whether the generator's package footer ought to rename such a collision.

### 11.6 What Phase D must know

- **The prefix plumbing is in.** `Body.prefix` is the address prefix (`"body"`, `"header:rId8"`,
  `"footnote:rId5"`, and whatever `sub()` is given), `Body.sub(container, prefix)` carries it into
  a cell or a control, and `Paragraph.__repr__` already prints `<Paragraph body/3 '…'>` from it.
  Phase D adds `address`, `element_at`, `paragraph_at`, `address_of` and `outline()`; it does not
  have to thread anything new through the views.
- **The ordinal path is `block_children_of` all the way down**, and the four `w:sdtContent` levels
  are already invisible to it, which is section 4's "the four `sdtContent` levels are skipped in
  paths". A path is therefore the index in each successive `block_children_of`, with no special
  case for a table, a row, a cell or a content control.
- **The id generator is on the package, not on the body.** `pkg.id_seed = 1234` fixes the sequence;
  with no seed the generator is seeded from the paragraph ids the document already carries, so two
  runs over the same document give the same new ids. Phase D's revision ids and Phase C's control
  ids should come from `pkg.id_generator()` too, and the note of section 4 stands: comment ids are
  a separate space.
- **`Paragraph.parent_table_cell` is wired and returns None** until Phase C; `Range.paragraphs`,
  `Body.view_for` and `Body.paragraph_for` are the three places a Phase C `Table` view has to be
  slotted into, and none of them is on a hot path.
- **Every mutating verb already returns what it made** (`insert_paragraph` a `Paragraph`,
  `insert_text` a `Range`, `insert_xml` and `insert_element` the views), so `ChangeReport` has
  something to report about without any signature changing.
- **`Body.get_text(max_chars=)` and `search(limit=)` are the budgets that exist**; `iter_paragraphs`
  and `iter_blocks` are lazy, so `outline(max_chars=)` and `find(limit=, context=)` can be written
  over them without materialising a list per call.

## 12. Phase D implementation notes (2026-09-16)

Phase D is done: the agent surface of section 3.4 — the three address forms, `outline()` under a
budget, `find()` with context, `describe()`, a `ChangeReport` on every mutating call, `dry_run`,
`DocumentSession`, determinism, the errors an agent can act on — with the agent scenario tests
and the token-budget test of section 7 and the README's "For agents" section. The suite is **866
tests** (865 passing plus one `xfail`; 857 of them fast, 45 s), against 776 at the end of Phase
B; **89** of the new ones are `tests/agent/`. Nothing in `~/git/docx4j-xsdata` changed, no schema
patch was needed and the model was not regenerated; `codegen/generate_el.py` changed only in its
engine footer (one name, `DocumentSession`), and the footer it writes is byte-identical to what
is committed.

```python
from docx4j_py import load

pkg = load("in.docx")
pkg.outline().to_markdown()                              # read it into a context window
hit = pkg.find("quick brown fox")[0]                     # address, offsets, snippet
pkg.paragraph_at(hit.address).insert_paragraph("New")    # edit by address
pkg.last_change.to_json()                                # what that call did
```

### 12.1 What landed

| Piece | Where | Signature |
|---|---|---|
| the three address forms | `docx4j_py/model/content/addresses.py` (610 lines) | `path_of(body, element)`, `ordinal_of(body, element)`, `address_of(body, target)`, `resolve_path`, `element_at(body, address)`, `paragraph_at(body, address=None, *, contains=None, para_id=None)`, `nearest_address`, `package_bodies`, `prefix_for_part`, `assign_para_id`, `ensure_para_ids` |
| on the views | `.body`, `.paragraph` | `Paragraph.address`, `Paragraph.ordinal`, `Block.address` / `.ordinal` (which is what `Table.address` and `ContentControl.address` are until Phase C), `Body.address_of`, `Body.element_at`, `Body.paragraph_at`, `Body.outline`, `Body.find`, `Body.range_of`, `Body.text_budget`, `Body.ensure_para_ids` |
| the results | `.reports` (1,022 lines) | `Outline(entries, headers, footers, stats, truncated)`, `OutlineEntry(address, para_id, ordinal, kind, style_id, level, text, chars, truncated, rows, cols, children)`, `OutlineSection(prefix, part_name, entries)`, `OutlineStats(paragraphs, tables, words, chars, comments, tracked_changes, skipped)`, `SearchHit(address, para_id, ordinal, start, end, match, snippet, before, after)`, `ChangeReport(operation, addresses, moved, created_para_ids, text_before, text_after, parts_touched, at)`, `TextExcerpt(text, chars, truncated)` |
| the recorder | `.reports` | `recording(body, operation)`, `ChangeRecorder`, `NULL_RECORDER`, `moved_by_insert`, `moved_by_delete`, `container_prefix` |
| `describe()` | `.describe` (452 lines) | `Description(styles, page, parts, headers, footers, custom_xml, authors, tracking_on, skipped, application, created, modified)`, `StyleInfo(id, name, kind, built_in, in_use)`, `PageSetup(width_pt, height_pt, orientation, margins)`, `PartInfo(name, content_type, kind, unmarshalled)` |
| `dry_run` | `.trial` (296 lines) | `dry_run(package)`, `TrialPackage`, `TrialPart` |
| the package's half | `.__init__`'s `register()` | `pkg.outline()`, `pkg.describe()`, `pkg.element_at()`, `pkg.paragraph_at()`, `pkg.find()`, `pkg.bodies()`, `pkg.dry_run()`; `pkg.changes`, `pkg.last_change` and `pkg.assigns_para_ids` are on `OpcPackage` itself, because they are state rather than behaviour and need no import from this layer |
| the session | `docx4j_py/model/sessions.py` (362 lines) | `DocumentSession(*, idle_timeout=900, max_open=32)` with `open`, `add`, `get`, `use`, `lock_for`, `save`, `close`, `close_all`, `handles`, `documents`, `sweep`, `__enter__`/`__exit__`, `__len__`, `__contains__`; `OpenDocument(handle, package, source, lock, last_used)` |
| the top-level name | `codegen/generate_el.py`'s `ENGINE_EXPORTS` | `from docx4j_py import DocumentSession` — the one name a server imports rather than reaches through a package |

The **prefix plumbing of section 11.6 was enough**, as it promised: `Body` gained addresses
without threading anything new through the views, and `path_of` is `block_children_of` walked
*upwards* through the parent pointers, so an address costs the nesting depth rather than the
document. A level whose block list does not hold the child contributes no index, which is how the
four `w:sdtContent` classes stay invisible to a path without being named anywhere — the same
trick that made Phase B's descent one visitor.

### 12.2 Departures from section 3.4, all deliberate

1. **`outline()` takes a `limit`, default 300, and section 3.4 does not mention one.** It had to:
   section 7 asks for a stated size at the default budgets, and a 200-page document has 2,200
   blocks, which is 380 KB of JSON at any per-entry size worth having. `limit=None` asks for all
   of them and `stats` counts the whole body either way, so an agent reading a truncated outline
   still knows how big the document is and can ask for more. `DEFAULT_ENTRY_LIMIT` is one
   constant and a test pins it.
2. **The notes parts' address prefixes lost their relationship id.** Phase B's `_prefix_for` gave
   a footnotes part `footnote:rId5`; section 3.4's own examples are `footnotes/1/0` and
   `comments/2/0`, and there is at most one of each part, so the id said nothing. Headers and
   footers keep theirs (`header:rId3`, `footer:rId5`) because there may be several. One Phase B
   test changed.
3. **A created document always stamps `w14:paraId`; a loaded one decides for itself.** Section
   3.4 asks for exactly this ("and always in a created document"), and it needed somewhere to
   put the decision: `pkg.assigns_para_ids` is `None` for "as the document does", and
   `create_package` sets it True. Phase B's test that a new document's paragraphs have no paraId
   became two tests, one per half of the rule.
4. **The "document already uses them" test is per body, not per package.** Looking at every part
   to collect the ids in use would unmarshal the headers, the footers and the notes of a document
   that only ever had its body edited, and CR-002's promise is that an untouched part is written
   back byte for byte. The ids in use are cached on the package per address prefix, which also
   stops a loop of inserts being quadratic. Two parts of one document colliding on a 31-bit id is
   not a risk worth the promise.
5. **`describe()` reads bytes with lxml and unmarshals nothing**, which is the choice section 3.4
   offered and preferred. Every part it reads — styles, settings, `docProps`, the comments part,
   the custom XML parts — goes through `part.xml` (the source bytes for a part nobody has
   touched) and `etree.fromstring`. The **one** exception is the main document part, whose tree
   is used when it is already unmarshalled: a caller with a `body` has unmarshalled it anyway,
   and re-marshalling two thousand paragraphs to count their style references would be absurd.
   A test asserts that describing a document leaves every part byte for byte and marks nothing
   for re-marshalling.
6. **`StyleInfo.name` is the display name, not the stored `w:name`.** Word stores `heading 1` and
   shows `Heading 1`; `paragraph.style` takes and reports the display name (section 4), so
   `describe()`'s list is in the same vocabulary as the setter an agent will call next.
7. **`Description`, `StyleInfo`, `PageSetup` and `PartInfo` live in `.describe`, and `dry_run` in
   `.trial`**, where section 5 puts the whole agent surface in `.addresses` and `.reports`. Two
   more modules of 450 and 300 lines rather than one of 1,800; every name is exported from
   `docx4j_py.model.content` all the same, as Phase A did when `builders.py` grew (section 10.1).
8. **`SearchHit` carries `ordinal` and `match` as well as section 3.4's list.** `address` is the
   paraId when there is one, so without `ordinal` a tool result would not say *where* the hit is
   in a document Word has not stamped; and `match` saves a server slicing `snippet` to find out
   what actually matched. `hit.range(body)` and `body.range_of(hit)` both convert back, and the
   conversion goes through the **address**, so a hit that has been through JSON and a tool call
   still works.
9. **`find()` is on `Range` too**, not only on `Body` and `Paragraph`, because `search` is.
10. **`DocumentSession.save` requires `overwrite=True` to write over the file the document was
    opened from.** Section 3.4 says `save(handle, path=None)`; the first version took None to
    mean "the source", and the first test written against it overwrote a corpus sample. That is
    docx4j-mcp's `overwrite: true` and the reason it exists, so it is the rule here: no target is
    the bytes, a target is that path, and the source needs asking for.
11. **`DocumentSession.use(handle)` is the way to hold the lock**, beside `get(handle)`, which
    section 3.4 does not mention. `get` returns the package and releases the lock, which is right
    for a read; an edit that must be atomic against another tool call needs the lock for its
    whole length, and a context manager is how Python says that.
12. **`SpanError`'s "where to split" is a method nothing calls yet.**
    `Range.holder_boundaries()` returns the offsets at which a span enters or leaves a
    `w:hyperlink`, a run-level `w:sdt`, a `w:ins` or another run holder, and
    `Range.require_one_holder(operation)` raises naming them. Phase D has no operation that needs
    a single holder; Phase C's range-level `insert_content_control` and Phase G's
    `insert_comment` do, and section 3.4 asks for the error now. It is tested directly rather
    than through a caller.

### 12.3 The budgets, measured

The token-budget test builds section 7's 200-page document with the Phase A builders — 2,000
paragraphs (a `Heading1` every fortieth), 20 three-by-three tables, 2,180 paragraphs and 2,200
blocks in all (2,020 of them the body's own), built in 0.05 s — and measures the JSON a tool would return:

| call | bytes of UTF-8 JSON | budget |
|---|---:|---:|
| `outline().to_json()`, defaults | **53,352** | 64 KB |
| the same, every paragraph carrying a `w14:paraId` | **60,804** | 64 KB |
| `outline(headings_only=True).to_json()` | **6,469** (7,698 with paraIds) | 8 KB |
| `outline(headings_only=True).to_markdown()` | 640 | — |
| `find(limit=20)`, `to_dict()` each, at `context=40` | **4,143** | 8 KB |
| `outline(limit=None).to_json()` | 389,778 | — (the honest whole) |

Three things make the default fit. `to_dict()` leaves out what is None or empty, which is worth
about 25% on a body of plain paragraphs. A table's `text` is its **first row** rather than its
whole content (`"Region | Quarter | Total"`), which is the preview an agent chooses an address
from and costs one row. And `limit` cuts the entries while `stats` still counts everything, so
the result says `paragraphs: 2180, truncated: true` over 300 entries rather than lying by
omission. The margin is thinnest on `headings_only` with paraIds — 7,698 of 8,192 — which is
what a 50-heading document costs; a document with a hundred headings needs `limit=`.

### 12.4 What a `ChangeReport` costs, measured

Recorded on every mutating call, as decided question 4 requires, and measured against the same
call with the recorder switched off:

| | |
|---|---:|
| `insert_paragraph` on a loaded document | **38.8 µs** |
| the same, with the report | **43.0 µs** |
| the report's share | **≈ 4 µs, 11%** |
| the recorder alone (construct, three notes, `finish`) | 2.0 µs |
| one stored `ChangeReport` | 96 bytes plus its tuples |

The first version cost **47 µs**, not 4, and the reason is worth recording: `change.touched(view)`
asked `address_of` for the address, which for a paragraph with no paraId scans its container for
the element's index — so a loop of two thousand inserts was quadratic. An insert already **knows**
the index it chose, so `ChangeRecorder.active` now lets a verb that can build the address cheaply do so, and `container_prefix(body, container)` gives it the prefix (free for the body's own list,
the container's own ordinal otherwise). A test asserts that the last five hundred of two thousand
inserts take no more than four times what the first five hundred took.

`moved` is the one field that is not a few values: it lists every ordinal in the **touched
container** that shifted, so inserting at the start of a 200-block body reports 207 pairs.
Appending — the default location, and the common call — reports none at all, and the cost is
`O(blocks after the insertion point)` in one container rather than anything document-wide. A test
pins both halves.

### 12.5 `dry_run`, and what it cannot undo

`pkg.dry_run()` yields a `TrialPackage`: a package-like object that copies a part's tree with
CR-001's `deep_copy` the first time the trial asks for that part's body, and answers `body`,
`outline`, `describe`, `element_at`, `paragraph_at`, `find`, `bodies` and `dry_run` over the
copies. Those seven have to be spelled out rather than left to `__getattr__`, because the real
package has them too and delegating would answer about the real document — the one wart in an
otherwise mechanical wrapper. `TrialPart` keeps the original as `_wrapped`, which is why a trial
header still reports `header:rId8`: `prefix_for_part` reads the wrapped part's class.

The limits, in the module's docstring and in the README:

- A part the real package has **not** read is unmarshalled by the trial's first look at it, and
  that marks it for re-marshalling on the *real* package. A dry run over `pkg.body` after
  `pkg.body` has been read costs a `deepcopy` of `w:document` and nothing else.
- **Parts added during a trial are discarded with it, and cannot be un-added.** Phase D adds
  none; Phase C's `insert_inline_picture` would add an `ImagePart` to the real package's part map
  — the trial shares it — and the docstring of anything that adds a part must say so.
- `trial.save()` is refused, with a hint saying to leave the block and make the calls again.
- The trial's id generator is seeded from the real package's, so the paraIds a trial allocates
  are the ones the commit that follows will allocate. A test asserts it.

### 12.6 Determinism, and what "the same bytes" needed

`test_the_same_seed_and_the_same_calls_give_the_same_bytes` runs eleven calls over
`2010-sample1.docx` twice and compares the saved bytes; it also compares the `ChangeReport`s,
minus their timestamps. Two things had to hold that did not follow from Phase B. The ids in use
are cached per body, so the generator is derived from the same material on both runs whatever
order the calls come in; and `ensure_para_ids` walks in document order, so a legacy document
stamped twice gets the same ids. A created document is **not** byte-reproducible across runs and
this test does not pretend otherwise: `create_package` writes `dcterms:created` and
`dcterms:modified` as the current time (CR-002 section 12.8), which is a property of the
document, not of the agent surface.

### 12.7 What Phase K (markdown with addresses) needs

- **`to_markdown(addresses=True)` has its address already.** `paragraph.address` is the paraId
  when there is one and the ordinal otherwise, which is exactly what section 3.5 wants in the
  `<!-- body/3 -->` comment, and `Body.element_at` accepts it back. Emit `paragraph.address`
  rather than `paragraph.ordinal`, so that a model editing markdown and writing back by address
  is not defeated by an insert.
- **`Outline.to_markdown()` is not `Body.to_markdown()`** and should not grow into it. The
  outline's markdown is a nested list of *entries* under a budget; section 3.5's is the document's
  content with its runs' formatting. `heading_level_of(element)` — `w:outlineLvl + 1`, else a
  `HeadingN` style — is the one piece worth sharing, and it is exported.
- **`max_chars=` on `to_markdown` should return a `TextExcerpt`-shaped answer** or say in its
  docstring that it truncates silently, as `get_text` does. `TextExcerpt(text, chars, truncated)`
  is there for the first.
- **`insert_markdown` chooses styles from `describe()`**, as section 3.5 says, and
  `Description.style_names(kind="paragraph", in_use=True)` is the call: it is over the styles the
  document *defines*, with `in_use` from docx4j's `stylesInUse`, and it costs no unmarshalling.
- Every `insert_markdown` must open one `recording(body, "insert_markdown")`, so that a fragment
  of twenty blocks is one `ChangeReport` and not twenty. The nested verbs are no-ops inside it
  already.

### 12.8 What Phase C needs

- **`Table.address` and `ContentControl.address` are the ordinal**, which is what `Block.address`
  already returns; when the views land, `Body.view_for` is the one place to change and the
  addresses do not move, because a path is `block_children_of` all the way down and a table's
  rows and cells are already indices in it (`body/4/0/1/0` is table, row, cell, block).
- **`outline()` already reports `rows`, `cols` and a first-row preview** by walking the tree; a
  `Table` view should be slotted into `_table_shape` rather than duplicated beside it.
  `_is_row` accepts a `w:tr` wrapped in a row-level `w:sdt` or `w:customXml`, which is the
  OpenDoPE repeat of section 4.
- **Anything that adds a part must say what a `dry_run` does with it** (12.5), and take its ids
  from `pkg.id_generator()` so that section 3.4's determinism holds for image part names and
  relationship ids as it now does for paraIds.
- **Every new mutating verb opens one `recording(...)`** and, if it inserts at a known index,
  reports the address from that index rather than through `address_of` (12.4).
- `Range.require_one_holder("insert_content_control")` is written and tested; call it.

## 13. Phase K implementation notes (2026-09-17)

Phase K is done: markdown out with addresses and markdown in (section 3.5), the tests of section
7, acceptance artefact 5 and the README's "Markdown, coarse and fine". The suite is **941 tests**
(940 passing plus one `xfail`; 931 of them fast, 49 s, and 59 s for the whole), against 866 at the
end of Phase D; **75** of the new ones are `tests/content/test_markdown.py` (59),
`tests/agent/test_markdown_workflows.py` (14) and the two `KEEP` tests in `tests/test_codegen_el.py`.
Nothing in `~/git/docx4j-xsdata` changed, no schema patch was needed and the model was not
regenerated. One dependency joined `pyproject.toml`: `markdown-it-py>=3.0`, for `insert_markdown`
only, imported on the first call (decided question 6).

```python
from docx4j_py import load

pkg = load("in.docx")
markdown = pkg.to_markdown(addresses=True)          # read it, with a handle on every block
pkg.paragraph_at("body/11").insert_paragraph("Two tables follow.", location="After")
pkg.body.insert_markdown("## Appendix\n\n- one\n- two\n")   # or write whole blocks
```

### 13.1 What landed

| Piece | Where | Signature |
|---|---|---|
| markdown out | `docx4j_py/model/markdown/export.py` (1,037 lines) | `body_markdown(body, *, addresses=False, view="accepted", max_chars=None, footnotes=True)`, `markdown_budget_of(body, max_chars=None, ...) -> TextExcerpt`, `paragraph_markdown(paragraph, *, addresses=False, view="accepted")`, `table_markdown(element, context=None, **options)`, `address_comment(address, indent="")`, `escape(text)`; `ADDRESS_COMMENT`, `MarkdownViewValue` / `MarkdownView`, `CODE_STYLE_IDS`, `MONOSPACE_FONTS` |
| markdown in | `docx4j_py/model/markdown/importer.py` (912 lines) | `blocks_for(body, markdown, recorder) -> (elements, parts touched)`, `ensure_style(package, style_id, *, touched=None, defined=None)`, `style_ids_of(part) -> set[str]`, `parser()`; `STYLE_IDS`, `CUSTOM_STYLE_XML` |
| the one verb | `docx4j_py/model/markdown/__init__.py` (111 lines) | `insert_markdown_into(body, markdown, *, location="End", target=None) -> list[Paragraph \| Block]`, and every public name re-exported |
| on the views | `.body`, `.paragraph` | `Body.to_markdown`, `Body.markdown_budget`, `Body.insert_markdown`, `Paragraph.to_markdown`, `Paragraph.insert_markdown` |
| on the package | `.content.__init__`'s `register()` | `pkg.to_markdown()`, `pkg.markdown_budget()`, `pkg.insert_markdown()` |
| on a trial | `.trial` | `TrialPackage.to_markdown` / `markdown_budget` / `insert_markdown`, and its own `style_definitions_part` and `numbering_definitions_part` |
| nothing dropped silently | `.reports` | `ChangeReport.warnings`, `ChangeRecorder.warn(message)` |
| the regeneration guard | `codegen/clean.py`, `tests/test_codegen_el.py` | `KEEP` gains `"model"`; two tests assert that every tracked hand-written path under `docx4j_py/` is covered by `KEEP` and that nothing `KEEP` names is missing |

The behavioural oracle is Java docx4j's `docx4j-markdown`. `WmlToMarkdown`'s decisions are kept:
headings from the effective outline level and **before** the numbering test (a built-in `Heading`
style can carry a legacy `w:numPr`); emphasis wrapped round a *run* of equally formatted text
rather than round each `w:r`; a header row's bold is convention, not markup; `w:gridSpan` padded
with empty cells and a `w:vMerge` continuation left empty, because GFM has no spans; a field's
cached result kept and its instruction dropped; a code block from consecutive `SourceCode`
paragraphs. `MarkdownToWmlVisitor`'s are kept too: the same style ids, the nine-level
`w:abstractNum` built from a scanned per-depth signature, a bullet-only signature shared and any
ordered list given its own `w:num` so it restarts, `w:contextualSpacing` for tightness, a
follow-on paragraph in a list item indented rather than numbered, and a quote winning over a list.

### 13.2 Departures from section 3.5, each deliberate

1. **`view="markup"` is CriticMarkup, where Java's `MARKUP` policy is `~~strikethrough~~`.**
   Section 3.5 asks for `{++...++}` / `{--...--}` and `{>>...<<}`, and it is right to: Java's
   deletion markup is indistinguishable from a real `w:strike`, so a model reading it cannot tell
   an author's formatting from a reviewer's edit. Insertions are `{++...++}`, deletions and
   `w:moveFrom` are `{--...--}`, and a `w:commentReference` becomes `{>>the comment text<<}` after
   the run that carries it. `{==...==}` (CriticMarkup's highlight) is **not** written: a comment
   range's start and end are separate elements and pairing them is Phase G's job, so the comment
   is attached at its reference and the highlight is left for that phase.
2. **The address comment is on its own line before the block, at the block's own indentation.**
   Section 3.5 shows `<!-- body/3 -->` without saying where, and the two candidates are not
   equal. A trailing comment is cheaper in tokens but cannot be put on a table without adding a
   cell; an own-line comment is uniform across every block kind, which is what was asked for. The
   cost is recorded rather than hidden: **an addressed export is for reading and addressing, not
   for feeding back through `insert_markdown`**, because in CommonMark an HTML-comment line
   between two list items ends the list. That is fine, because the two workflows are separate ---
   the coarse one never asks for addresses. `ADDRESS_COMMENT` is the committed regex
   (`^[ \t]*<!--\s(?P<address>\S+)\s-->$`, multiline) and a test asserts that `Body.element_at`
   accepts everything it yields.
3. **An empty paragraph keeps its comment and nothing else.** Markdown has no empty paragraph and
   Java drops them; but dropping one under `addresses=True` would make a block unaddressable from
   the markdown, so the comment line stays on its own.
4. **Formatting is read direct, not effective.** CR-002 Phase B's `PropertyResolver` does not
   exist, so where Java compares a run's *effective* `w:rPr` against its paragraph style's
   baseline, this reads the run's own `w:rPr`. Section 6 allows exactly this. The visible
   difference is that a `Heading 1` whose style is bold does **not** come out as `# **Title**`
   here either (its runs carry no `w:b`), but a document that sets `w:b` directly on a heading's
   runs will, where Java would cancel it against the baseline. A test to flip belongs with the
   `PropertyResolver`.
5. **`in_use=True` is *not* the filter on `describe()`'s style list.** Section 12.7 suggested
   `Description.style_names(kind="paragraph", in_use=True)`. That is wrong for the job: a styles
   template defines `Heading 1` without ever having used it, and `in_use` is docx4j's
   `stylesInUse`, which reads the main document part's references. The choice is over every style
   the document **defines** — read from the styles part's bytes with lxml by `style_ids_of`, so a
   fragment whose styles are all there unmarshals nothing — falling back to `describe()`'s
   display names when an id does not match, and to `ensure_style` when neither does.
6. **`Paragraph.insert_markdown` accepts `"Start"` and `"End"` as well as `"Before"` and
   `"After"`.** Section 3.5's signature lists only the last two, but section 4's paragraph-level
   merge rule ("a fragment of exactly one `w:p` has its runs merged into the paragraph, as Word's
   paste does") has nowhere to live otherwise. `"After"` stays the default, so the short call is
   unchanged, and the merge behaviour is `insert_xml`'s, verbatim: it must be asked for.
7. **An image is never fetched, and becomes a link rather than its bare alt text.** docx4j-mcp's
   posture (and Java's `DefaultMarkdownImageHandler`, which declines a remote URL): an agent must
   not be able to make the library open a socket. The alt text becomes the link's text and the
   destination its target, so nothing in the markdown is lost, and the report's `warnings` say so
   for every image. Embedding a *local* file is Phase C's `insert_inline_picture`, and this should
   call it when it exists.
8. **`ChangeReport` gained a `warnings` field.** Section 3.5 says "nothing may be dropped
   silently" but the report had nowhere to say what was. A field on the frozen dataclass, left out
   of `to_dict()` when empty, was cheaper than a second result type and is what a tool returns
   anyway. `ChangeRecorder.warn` is its one writer.
9. **The content layer reaches the markdown module through a function-level import.** Section 5's
   rule is that `docx4j_py.model.markdown` imports the content layer and never the reverse, and at
   *module* level that holds exactly. But `body.to_markdown()` has to work without the caller
   having imported anything, and the registration trick that gives `XmlPart` its `body` cannot be
   used here because `Body` is itself lazily imported. So `Body.to_markdown`, `Body.markdown_budget`,
   `Body.insert_markdown`, `Paragraph.to_markdown` and `Paragraph.insert_markdown` are five
   two-line methods whose bodies import `docx4j_py.model.markdown`. No module-level cycle exists in
   either direction, `tests/openpackaging/test_threads_and_import.py` still passes, and
   `import docx4j_py` does not import the markdown package at all.
10. **The trial got two more spelled-out parts.** `insert_markdown` *writes* to `styles.xml` and
    `numbering.xml`, so `TrialPackage.style_definitions_part` and `.numbering_definitions_part`
    now return the trial's copies, as `main_document_part` does. Without them a dry run left the
    style it added on the real document, which is not what 12.5 promises. A part the trial
    **creates** still cannot be un-created: a trial of a fragment with a list leaves an empty
    `numbering.xml` in the real package (and its `[Content_Types].xml` entry and relationship),
    and what the trial wrote into it goes with the copy. A test pins both halves.
11. **One `MarkdownIt` serves the process.** A parser per call cost 600 µs of the 800 a
    one-paragraph fragment took. A `MarkdownIt` keeps no per-parse state — the block and inline
    states are made inside `parse` — so it can be shared where CR-001 section 14.7's `ParserConfig`
    could not. `parser()` is the accessor and the place the import happens.
12. **`tests/agent/test_markdown.py` is `test_markdown_workflows.py`.** pytest's prepend import
    mode refuses two test modules with the same basename in directories that have no `__init__.py`,
    and `tests/` has none by design (every file there does `from conftest import ...`).

### 13.3 The numbers, measured

The 200-page document is section 7's fixture, unchanged: 2,000 paragraphs with a `Heading1` every
fortieth, 20 three-by-three tables, 2,180 paragraphs and 2,200 blocks.

| call | |
|---|---:|
| its markdown | **208,172 characters** (208,172 bytes of UTF-8) |
| `to_markdown()` | **7 ms** |
| `to_markdown(addresses=True)` | 50 ms (245,442 characters) |
| `markdown_budget(16 KB)` | 8 ms |
| `get_text()`, for comparison | 9 ms |
| `outline()`, for comparison | 23 ms |
| `insert_markdown` of all 208 KB | **113 ms**, 2,020 blocks |
| of which markdown-it's own parse | 28 ms |
| `insert_markdown` of one paragraph | 205 µs |
| `insert_paragraph`, for comparison | 27 µs |
| `import markdown_it` | 17 ms, and **not** loaded by `import docx4j_py` (656 ms, unchanged) |

Reading is cheaper than the outline, which is the right way round: markdown *is* the cheap read.
`addresses=True` costs seven times as much, because every block asks `address_of` for its address
and a paragraph without a `w14:paraId` is an upward walk — the same cost 12.4 measured, and the
same fix would apply (a renderer that knows the index it is at could build the address itself);
50 ms for a 200-page document did not justify the complication. Writing is an order of magnitude
dearer than the equivalent verb, and three quarters of that is this code rather than the parser:
a `w:rPr` per run, a `w:pPr` per paragraph and a style lookup per styled block.

### 13.4 The limits

- **Math is out of scope.** Java's module converts OMML to LaTeX and back (`docx4j-markdown`'s
  CR-006); an `m:oMath` here contributes nothing to the markdown, and `$...$` in markdown is
  text. A later phase can port `OmmlToLatex` / `LatexToOmml`; nothing in this design is in its way.
- **GFM footnote *import* is out of scope**, and the decision was the dependency: `markdown-it-py`
  has no footnote rule, `mdit-py-plugins` is the only reasonable way to get one, and section 3.5
  allows exactly one dependency. Doing it by hand means a block rule *and* an inline rule *and* a
  definition collector, which is not "a few lines". So `[^1]` stays literal text on the way in,
  while footnotes are exported properly. A footnotes part is written by nothing in this phase.
- **Export-only constructs**: task list items (`- [x]`), YAML front matter, a fence's info string
  (` ```python ` comes back as ` ``` `), indented code blocks as distinct from fenced ones,
  reference-style links, and inline HTML. Java carries the first two; both would need
  `mdit-py-plugins`.
- **A loose list item's follow-on paragraph is not reattached on export.** Java appends a
  `ListParagraph`-styled paragraph with no `w:numPr` to the open list item; here it closes the
  list and becomes a plain paragraph. The import writes such paragraphs (indented, unnumbered), so
  a list with multi-paragraph items does not survive a coarse round trip as one list. It is a
  handful of lines in `_render_blocks` and belongs with Phase H, which owns lists.
- **Escaping is narrow and hand-written**, because there is no commonmark renderer to delegate to:
  `\`, `` ` ``, `*`, `_`, `[`, `]`, `<`, `>` inline, and a leading `#`, `>`, `+`, `-`, `=`, `|` or
  `N.` at the start of a rendered paragraph. It is deliberately not exhaustive; a paragraph of raw
  HTML-looking text will round-trip, a pathological one may not.
- **`mc:AlternateContent` and text boxes** contribute nothing, as they do to `Body.text`
  (decided question 8).
- **Headers, footers and endnotes** are not in `pkg.to_markdown()`: Java's exporter leaves them
  out too, and `header_part.body.to_markdown()` renders one when a caller wants it.
- The addressed export is not re-importable as markdown (13.2 item 2).

### 13.5 What Phase C needs

- **`Table.to_markdown()` hangs on `table_markdown(element, context)`**, which already takes the
  element and any context (a `Body`, a renderer context, or nothing). Add the method; do not
  duplicate the renderer, and in particular do not re-derive `_rows_of` / `_cells_of`, which
  already unwrap a row- or cell-level `w:sdt` as section 4 requires.
- **`insert_markdown` builds its tables with the Phase A `tbl` / `tr` / `tc` builders and sizes
  the grid from `w:sectPr`**, which is what `insert_table` is specified to do; when `insert_table`
  lands, `_Importer.table` should call it rather than keep its own copy of the sizing, and
  `writable_width()` is the one function to move.
- **An image should become a real `InlinePicture` when the destination is a local file**, through
  Phase C's `insert_inline_picture`. The remote case stays a link, always (13.2 item 7).
- `ContentControl` views change nothing here: the renderer descends through every `w:sdt` form
  already, because `block_children_of` does.

### 13.6 What Phase H needs

- **The list renderer is `_OpenList` in `export.py`**, which numbers items itself from the
  numbering part's `w:numFmt` and `w:start`, as section 6 says it must until the `Emulator` lands.
  When `ListItem.list_string` exists, the marker should come from it, and the four format families
  this does not distinguish (`lowerLetter`, `lowerRoman`, `upperLetter`, `upperRoman` all render
  as `N.`) come right for free.
- **A run of list paragraphs is one markdown block**, broken by any non-list paragraph and by a
  different top-level `w:numId`. That is Java's rule and it is why four list paragraphs separated
  by empty ones each restart at `1.` (the `tests/fixtures/lists.docx` test pins it). Phase H should
  decide whether a `List` view makes the "loose item's follow-on paragraph" rule of 13.4 worth
  implementing.
- **`_Numbering` in `importer.py` is docx4j's `ImportNumbering`**, including its per-depth
  signature scan and the "bullet-only lists share one `w:num`" rule. A `List` API should create
  definitions through it rather than beside it, and its id allocation (the next free
  `w:abstractNumId` and `w:numId` above the document's own) is what keeps section 3.4's
  determinism.

## 14. Phase C implementation notes (2026-09-17)

Phase C is done: `Table`, `TableRow`, `TableCell` and `InlinePicture` with the header readers of
Phase A, `insert_ooxml` over a flat OPC `pkg:package`, the `ContentControl` reads and `delete`,
the tests of section 7 and acceptance artefact 6. The suite is **1,011 tests** (1,010 passing
plus one `xfail`; 1,002 of them fast, 53 s, and 62 s for the whole), against 941 at the end of
Phase K; **71** of the new ones are `tests/content/test_table.py` (20), `test_picture.py` (14),
`test_ooxml.py` (13), `test_controls.py` (15) and `tests/agent/test_tables_and_pictures.py` (9).
Nothing in `~/git/docx4j-xsdata` changed, no schema patch was needed, the model was not
regenerated and `codegen/generate_el.py` did not change at all: every Phase C name is reached
through `docx4j_py.model.content`, so `ENGINE_EXPORTS` gained nothing.

```python
from docx4j_py import load

pkg = load("in.docx")
table = pkg.body.insert_table(3, 2, values=[["Name", "Value"]], style="TableGrid")
table.cell(1, 1).paragraphs[0].insert_text("120")
pkg.body.insert_inline_picture(png_bytes, width=180, alt_text_description="A pangolin")
pkg.body.insert_ooxml(package_xml)          # Word's own clipboard format
```

### 14.1 What landed

| Piece | Where | Signature |
|---|---|---|
| the tables | `docx4j_py/model/content/table.py` (896 lines) | `Table(element, container, parent_body)` with `row_count`, `rows`, `values` (get and set), `style` / `style_id` / `style_built_in`, `header_row_count` (get and set), `text`, `address`, `ordinal`, `parent_table_cell`, `cell(row_index, cell_index)`, `add_rows(row_count=1, *, location="End", values=None)`, `delete_rows(row_index, row_count=1)`, `column_widths()`, `column_widths_twips()`, `insert_rows`, `delete()`, `to_markdown()`, `get_xml()`, `to_dict()`; `TableRow` (`row_index`, `cell_count`, `cells`, `values`, `is_header`, `insert_rows`, `delete`); `TableCell` (`body`, `paragraphs`, `tables`, `text`, `value`, `row_index`, `cell_index`, `parent_row`, `parent_table`, `width`, `column_width`, `insert_paragraph`, `insert_text`) |
| the shared row walk | `.text_model` | `rows_of(table)`, `cells_of(row)` → `[(element, the live list holding it)]`, row- and cell-level `w:sdt` / `w:customXml` unwrapped |
| the sizing | `.table` | `writable_width(body) -> int` (twips), `insert_table_into(...)`, `cell_of(value, body)` |
| the pictures | `docx4j_py/model/content/picture.py` (601 lines) | `InlinePicture(element, run, paragraph)` with `width` / `height` (points, get and set), `alt_text_description`, `alt_text_title`, `image_format`, `rel_id`, `inline`, `image_part`, `get_bytes()`, `get_base64()`, `delete()`, `get_xml()`, `to_dict()`; `add_image(body, data, *, width, height, alt_text_description, alt_text_title, name) -> NewPicture`, `free_image_name`, `next_drawing_id`, `writable_width_emu`, `pictures_of`, `note_added_part` |
| the controls | `docx4j_py/model/content/controls.py` (613 lines) | `ContentControl(element, container, parent_body)` with `type`, `form`, `tag`, `title`, `id`, `text`, `paragraphs`, `tables`, `content_controls`, `inline_pictures`, `address`, `body()`, `parent_paragraph()`, `get_range(location="Whole")`, `search`, `insert_text`, `insert_paragraph`, `delete(*, keep_content=True)`, `get_xml()`, `to_dict()`; `controls_in`, `controls_in_paragraph`, `parent_control_of` |
| the paste | `docx4j_py/model/content/ooxml.py` (469 lines) | `content_of(ooxml, *, target=None, package=None)`, `is_flat_opc`, `rewrite_relationship_ids`, `REL_ATTRIBUTES`, `insert_ooxml_into_body` / `_paragraph` / `_range` |
| flat OPC, read | `docx4j_py/openpackaging/stores.py` (+89 lines) | `FlatOpcStore.parse(source) -> FlatOpcStore`, a `MemoryPartStore` subclass; `FlatOpcStore.NAMESPACE` |
| on the views | `.body`, `.paragraph`, `.range` | `Body.tables` / `content_controls` / `inline_pictures` / `insert_table` / `insert_inline_picture` / `insert_inline_picture_from_base64` / `insert_ooxml`; `Body.view_for` hands out `Table` and `ContentControl`; `Paragraph.parent_table_cell` / `parent_content_control` / `content_controls` / `inline_pictures` / `insert_inline_picture` / `insert_inline_picture_from_base64` / `insert_ooxml`; `Range.insert_ooxml` |
| the trial's undo log | `.trial` | `TrialPackage.note_added_part(part, relationship, source, *, added_content_type=True)`, `TrialPackage.discard_added_parts()` |
| the outline | `.reports` | `_table_shape` is now `rows_of` / `cells_of`; `_is_row` is gone |

`Body.view_for` was the **one** place section 12.8 said had to change, and it was: a path is
`block_children_of` all the way down, so `body/4/0/1/0` meant table, row, cell, block before the
views existed and means the same now. Two Phase B tests that asserted `Block` for a `w:tbl`
became `Table`; nothing else moved.

### 14.2 Departures from section 3.2 and section 4, all deliberate

1. **`Body.tables` is this body's own list, not every table under it.** Office JS's
   `Body.tables` is the body's top-level tables (a nested one is `cell.body.tables`, one inside a
   control is `control.tables`), and the TypeScript engine reads it the same way. `paragraphs`
   still descends everywhere, which is the asymmetry Office JS itself has. A test pins both
   halves over `tests/fixtures/nested-table.docx`.
2. **`column_widths()` is points; `column_widths_twips()` is the twips.** Section 3.2 declares
   `list[float]`, and every other measurement in this API is points (`TableCell.width`,
   `Paragraph.left_indent`). The TypeScript engine returns twips there; the twips are still one
   call away, and `add_rows` uses them.
3. **`add_rows` takes `row_count` first and `location` as a keyword**, where Office JS's is
   `addRows(location, rowCount, values)`. Decided question 2 and section 3.1: the location has a
   default and options are keyword-only, so the common call is `table.add_rows()`.
4. **`TableRow.insert_rows` and `Table.add_rows` refuse `row_count < 1`** with
   `table.no_rows` rather than doing nothing, and `insert_table` refuses a zero row or column
   count with `table.empty`. An agent that computed a count wrongly gets told.
5. **`Table.address` has no paraId form.** A `w:tr` carries `w14:paraId` in this schema but a
   `w:tbl` does not, so a table's address is always the ordinal --- which is what
   `Block.address` returned before, so no address changed (section 12.8's promise).
6. **`ContentControl.text` does not go through `text_of` for the block, row and cell forms.**
   `text_of` treats a `w:sdtContent` as a *run holder* (it is in its `RUN_HOLDERS` set, because a
   run-level control's content is one), so a block control's two paragraphs came back as one
   line, `"onetwo"`. The three block forms now answer through `body().text`, which is a line per
   paragraph; the run form still reads `text_of`, which is right for it. Found by a test.
7. **`insert_text(location="Start" | "End")` on a run-level control edits the control's own
   items**, not `Paragraph.splice` at the control's boundary. `splice` at a boundary extends the
   **neighbouring** run (Phase B, section 11.3: that is what makes `replace_text` replace exactly
   what `search` matched), and at a control's boundary the neighbour is a run *outside* it, so
   `insert_text("The ", location="Start")` landed before the control and the control's text did
   not change. Word puts text asked for at the start or the end of a control inside it.
   `"Replace"` is still `splice(start, end, text)`, which is correct because the whole span is
   inside. Found by a test.
8. **`ContentControl.form` is lower case** (`"block"`, `"run"`, `"row"`, `"cell"`), as section
   3.2 spells it; the TypeScript engine capitalises. It is not an Office JS value --- Office JS
   has no `form` at all --- so it is not data that crosses into a tool argument under an Office
   JS name, and section 3.1's rule does not bind it.
9. **`ContentControl.body()` is a method, not a property**, because it *creates* nothing but is
   a constructor call over the control's content and reads better as one; `paragraphs`, `tables`
   and `content_controls` are the properties a caller wants. `parent_paragraph()` likewise.
10. **The ids Phase C allocates come from the document's own state, not from
    `pkg.id_generator()`.** Section 12.8 asked for the generator "so that determinism holds for
    image part names and relationship ids". It holds — but a *random* generator is the wrong tool
    for these three: `/word/media/imageN.<ext>` is the first free N (docx4j's `getNewPartName`),
    a relationship id is the relationships part's own next free `rIdN`, and `wp:docPr/@id` is one
    above the highest in the part. All three are derived from the document, so the same document
    and the same calls give the same bytes — which a test asserts by running eleven calls twice
    and comparing the saved bytes — and the names stay readable. `id_generator()` goes on serving
    the paraIds, which have no such natural order.
11. **`insert_ooxml` reports `ooxml.no_main_part` for any incoming package it cannot read as a
    WordprocessingML document**, the underlying `InvalidFormatException` quoted in the message.
    A flat package with no `officeDocument` relationship fails inside the loader, and two codes
    for "this is not a document you can paste from" would be two codes an agent has to learn.
12. **`Table.parent_table_cell` and `Paragraph.parent_table_cell` share `cell_of`**, which walks
    **up** the parent pointers and is transparent to a cell-level `w:sdt`. A paragraph outside a
    table costs one `getattr` chain to the `w:body`; there is no scan.
13. **`_is_row` is gone.** Section 12.8 asked for the `Table` view to be slotted into
    `_table_shape`; `_is_row` accepted a `w:sdt` *as a row*, so an OpenDoPE repeat outlined as one
    row whose "cells" were its rows. `rows_of` unwraps it, and the outline now reports the invoice
    fixture's repeating section as the 2×4 table Word shows. That is a fix, and a test pins it.
14. **`Body.insert_ooxml` takes `target=`**, which section 3.2 does not list, so that
    `"Before"` and `"After"` work as they do for `insert_xml`; `"Replace"` clears the body first,
    as Office JS's does.

### 14.3 The flat OPC decision

`insert_ooxml` needs to read a `pkg:package`, and CR-002 Phase C (flat OPC) is proposed, not
implemented. The choice was a private reader inside `ooxml.py` or a real
`PartStore`. The `PartStore` protocol made it small enough that there was no argument:

```python
class FlatOpcStore(MemoryPartStore):
    NAMESPACE = "http://schemas.microsoft.com/office/2006/xmlPackage"

    @classmethod
    def parse(cls, source: str | bytes) -> FlatOpcStore: ...
```

**89 lines**, of which the parse is thirty: a `MemoryPartStore` already answers `part_names`,
`has`, `load` and `size`, so unpacking each `pkg:part` into it (a `pkg:xmlData`'s root element
serialised, or a `pkg:binaryData` base64-decoded) is the whole job, and `load()` takes any
`PartStore`. One thing had to be added that a zip does not need: a flat package carries **no
`[Content_Types].xml`** — each part states its own `pkg:contentType` — so one is synthesised from
those through `ContentTypeManager`, which is what docx4j's `FlatOpcXmlImporter` does.

So this delivers the **read** half of CR-002 Phase C's flat OPC item; writing one
(`FlatOpcXmlExporter`, a `PartSink`) is still that phase's, and a note in CR-002 section 12 says
so. `scripts/acceptance.py` and `tests/content/test_ooxml.py` each carry a 25-line `flat_opc()`
that writes one, which is the measure of how small the exporter would be.

### 14.4 `dry_run`, and the parts a verb adds

Section 12.5 recorded that a part added during a trial "is discarded with it, and cannot be
un-added", and required the docstring of anything that adds a part to say so. Phase C's two verbs
add parts, so the question was whether to say it or fix it. Fixing it turned out to be **26
lines**, so it is fixed:

- `TrialPackage.note_added_part(part, relationship, source, *, added_content_type)` appends to an
  undo log; the free function `picture.note_added_part` calls it and is a no-op on a real package,
  so neither verb knows what kind of package it is in.
- `dry_run.__exit__` calls `TrialPackage.discard_added_parts()`, which removes the relationship,
  the part and — only when this call added it — the content-type override.
- A test asserts that a dry run of `insert_inline_picture` and one of `insert_ooxml` each leave
  the real package **byte for byte**, relationship parts included.

A parts *overlay* was the alternative and is not small: a trial would need its own copy of every
relationships part the verb touches, its own `get_part`, and a rule for resolving across the two
maps. The undo log answers the same question at a twentieth of the size.

What still cannot be un-created is a part a trial **creates for a different reason** and does not
report: Phase K's `insert_markdown` makes an empty `numbering.xml` when the document has none.
That is unchanged and 13.2 item 10 still describes it; when Phase H owns lists it should report
through the same log.

### 14.5 The numbers, measured

Against `samples/2010-sample1.docx` loaded and its body read, with the 157 KB PNG out of
`samples/Images.docx`:

| call | |
|---|---:|
| `insert_table(3, 3, values=…)` | **214 µs** |
| `insert_inline_picture` (157 KB PNG) | **1.0 ms** |
| the same with a 75-byte PNG | 0.85 ms |
| of which `next_drawing_id`, a `find` over the part | 0.13 ms and rising |
| of which `image_size`, the header reader | 1.1 µs |
| `insert_ooxml`, a 226 KB `pkg:package` with one image | **3.7 ms** |
| `insert_ooxml`, a bare one-paragraph fragment | 435 µs |
| `insert_paragraph`, for comparison | 17 µs |

The picture's millisecond is **not** the image: a 75-byte PNG costs 85% of what a 157 KB one
does, and reading the header is a microsecond. It is the two scans that make the call
deterministic --- `free_image_name` over the package's parts and `next_drawing_id` over the part's
tree --- plus the relationship and the content type. `next_drawing_id` is the one that grows:
**8 ms** on the 200-page document of section 7, where the whole call is 9 ms. That is the cost
12.4 measured for `address_of` in another guise, and the same fix would apply (cache the highest
id per body on the package, as `_para_ids_taken` caches the paraIds); one picture in a 200-page
document at 9 ms did not justify it, and a phase that inserts pictures in a loop should.

The package paste is three quarters parse --- `load()` of a whole package, through
`FlatOpcStore` --- and one quarter copy.

On the 200-page document of section 7 (2,000 paragraphs, 20 three-by-three tables):

| call | |
|---|---:|
| `outline()` **before** Phase C (`_table_shape` as Phase D wrote it) | 20.8 ms |
| `outline()` **after** (the shared `rows_of` / `cells_of`) | **20.5 ms** |
| `body.tables` (the body's own list, 20 of them) | 0.1 ms |
| `body.content_controls` (none, but the walk is the same) | 1.7 ms |
| `body.inline_pictures` (none; a walk of every run) | 3.8 ms |

The outline is unchanged within the noise, which is what 12.8 wanted: the view was slotted in
rather than duplicated, and it costs nothing because `rows_of` is the same walk `block_children_of`
was doing, with one element-name test per child. The budget test of section 7 is untouched.

### 14.6 What Phase E needs

- **`sdt_property` and `sdt_kind_of` are the only accessors**, and they already look in `w:` and
  `w15:` (Phase A, section 10.3 item 4), which is what `w15:dataBinding` needs. `ContentControl`
  has a private `_property(local_name)` for the string-valued ones; a setter belongs beside it.
- **`ContentControl.form` decides from the content for `SdtBlock`.** A typed kind must not change
  that: a `w:picture` control is block-form in `invoice2013.docx` and run-form elsewhere.
- **`insert_content_control` is not written**, at any level. `Range.require_one_holder(
  "insert_content_control")` is written and tested (12.8) and is what the range-level one should
  call; `next_sdt_id(root)` (Phase A) is the id, and `sdt(content, form=…)` the builder.
  `RepeatingSection` is already refused at run level by the builder.
- **A bound control's `insert_text` must write through to the custom XML node** (section 4). The
  run-level path is `ContentControl._extend` and `Paragraph.splice`; both are in one place, so the
  write-through hook is one call.
- **`placeholder_text` needs `w:showingPlcHdr` and the control's own runs**, and `_extend` is the
  function that builds a run inside an empty control.

### 14.8 A Phase K correction, found while checking Phase C (2026-09-17)

The table renderer of section 13 compared `w:vMerge/@w:val` to `"restart"` through `str()`,
and the model's value is an enum whose `str()` is the member's qualified name, so the first row
of a vertically merged cell rendered empty (`samples/sample-docx.docx`'s "Vertical merge" cell).
It now compares the enum's value; `test_a_vertically_merged_cells_first_row_keeps_its_text` pins
it. Nothing else in section 13 changes.

### 14.9 A Phase B correction: a built-in style the document lacks is defined, not dangled (2026-09-17)

**What Word showed.** Acceptance artefact 6 does
`body.insert_paragraph("Tables and pictures", style="Heading 1")` on a loaded
`samples/2010-sample1.docx`, whose `styles.xml` defines six styles and no heading among them. The
saved paragraph carried `<w:pStyle w:val="Heading1"/>` and Word rendered it as **Normal**: no
repair prompt, no warning, the style simply gone. Section 11.2's item 10 had it backwards. Word
does not materialise a definition for a dangling `w:pStyle`; it falls back to Normal, and a
dangling `w:tblStyle` falls back to Table Normal the same way.

**The rule now.** A name that resolves through `BUILT_IN_STYLES` to a style the document's styles
part does not define gets its **definition added** --- docx4j's `KnownStyles.xml`, the same 164
styles Phase K's markdown import activates on demand --- before the id is written. A custom name
the document does not define still raises `StyleError` with `code="style.not_found"` and the five
closest names: nothing to copy from, and guessing is worse than refusing. Two cases still write
the id as it stands, because there is nothing to add it to or nothing to add:

- **a document with no styles part at all** --- `ensure_style` returns the id untouched, which is
  what it already did;
- **the seven modern table styles** (`PlainTable1` ... `PlainTable5`, `GridTable1Light`,
  `ListTable1Light`). They are in `BUILT_IN_STYLES` because Office JS's `Word.Style` has them, and
  they are *not* in `KnownStyles.xml`, which predates them. `define_built_in` swallows
  `style.not_creatable` for exactly this case and writes the id. Word has them latent and will
  most likely resolve them; if it does not, the fix is seven more definitions in the resource, not
  a different rule.

**What a setter may now touch.** Reads are unchanged: `style`, `style_id` and `style_built_in`
unmarshal nothing, and `test_reading_a_style_never_unmarshals_the_styles_part` pins it over a
whole document. A setter unmarshals `/word/styles.xml` only when it has to add a definition ---
`style_ids_of` reads the ids with lxml from the part's bytes first, so a style already defined
costs no re-marshalling --- and when it does add one, `/word/styles.xml` goes into the call's
`ChangeReport.parts_touched`, through `current_recorder(package).parts`. Artefact 6's heading
insert now reports `('/word/document.xml', '/word/styles.xml')`; a second
`insert_paragraph(style="Heading 1")` on the same document adds nothing and reports
`('/word/document.xml',)` alone.

The callers that pass `define=True`: `Paragraph.style` (and so `Range.style`, which delegates),
`Body.insert_paragraph(style=)` and `Paragraph.insert_paragraph(style=)` (both through
`view.style`), `Table.style`, `Body.insert_table(style=)`, and the two `style_built_in` setters,
which bypass `style_id_of` and call `define_built_in` on what `id_of_built_in` returns.
`Table.style` and `Table.style_built_in` now open their own `recording` so that the definition is
added inside it and the part reaches the report. **`style_id` stays raw**: "set `style_id` to
write the id as it stands" is the documented escape hatch, it is what the error's hint offers, and
`add_styled_paragraph_of_text` (docx4j's own API) goes through it.

**Where the machinery lives.** `ensure_style`, `style_ids_of`, `_known_styles` and
`CUSTOM_STYLE_XML` moved from `docx4j_py/model/markdown/importer.py`, where Phase K built them,
into `docx4j_py/model/content/styles.py`, where the style rule lives: the markdown package imports
the content layer and not the other way round. They are re-exported from
`docx4j_py.model.markdown.importer` and `docx4j_py.model.markdown`, so Phase K's `__all__` and its
tests are unchanged, and the `wml` import `CUSTOM_STYLE_XML` needs is deferred into the one branch
that uses it.

**Two consequences worth knowing.** docx4j's `Heading1` carries `w:link w:val="Heading1Char"` and
`w:numPr/w:numId w:val="3"`, and a document that has neither gets a definition pointing at both.
That is what Java docx4j's `ImportStyles` writes too, and Word resolves an unresolvable `w:link`
or `w:numId` by ignoring it. And a `dry_run` defines the style on the **trial's** copy of the part
only: `TrialPackage.style_definitions_part` already returned a trial part for Phase K's sake, and
`test_a_trial_defines_the_style_on_the_trial_only` pins that the real part is left alone.

**What this does not fix.** Artefact 6's *pasted* heading still renders as Normal, and should:
`insert_ooxml` takes the body of a flat OPC package and merges neither its styles nor its
numbering (section 4 --- that is docx4j's `MergeDocx`), so a `w:pStyle` the fragment brings dangles
by design. The acceptance row in `tests/README.md` now says so, so that the next Word check does
not read it as the same defect.

### 14.7 What Phase G (comments) and Phase F (tracking) need

- **`TableRow.delete` and `Table.delete_rows` remove the row.** Tracked, a deleted row must stay
  in the tree with a `w:trPr/w:del` until it is accepted (section 4), and `TrPr.content` is the
  choice list to put it in --- `TableRow.is_header` is the worked example of reading and writing
  that list. `Table.delete` should mark every row rather than removing the table.
- **A new row's paragraphs need `w:ins`** when tracking is on: `_row_element` is the one place
  rows are built, in `add_rows` and `insert_rows` alike.
- **Deleting a content control is not tracked** (section 4), and `ContentControl.delete` writes
  no revision markup, which is already right.
- **`InlinePicture.delete` prunes the run through `Paragraph._remove_empty_runs`**, which knows
  about `w:ins` and `w:del`; a tracked delete should wrap rather than prune.
- **A comment anchored in a cell** gets its `Range` from `cell.body`, which is an ordinary `Body`
  with the cell's address prefix, so nothing in Phase G has to know about tables.

## 15. Phase G implementation notes (2026-09-17)

Phase G is done: comments as section 3.9 specifies them and section 4's "Comments" bullet rules
them --- `get_comments()` on a body, a paragraph, a range and a content control,
`insert_comment(text)` on a paragraph and a range, `Comment` with its replies, its `resolved`
flag and its `delete()`, `WordprocessingMLPackage.author`, the five parts kept in step --- with
the tests of section 7, acceptance artefact 7 and the README's "The audit trail". The suite is
**1,067 tests** (1,066 passing plus one `xfail`; 1,058 of them fast, 53.7 s, and 60.0 s for the
whole), against 1,020 at the end of Phase C; **47** of the new ones are
`tests/content/test_comments.py` (39) and `tests/agent/test_comment_workflows.py` (8). Nothing in
`~/git/docx4j-xsdata` changed, no schema patch was needed, the model was not regenerated and
`codegen/generate_el.py` did not change at all.

```python
from docx4j_py import load
from docx4j_py.model.content import Author

pkg = load("in.docx")
pkg.author = Author("Claude", initials="C", email="claude@example.com")
comment = pkg.find("first")[0].range(pkg.body).insert_comment("Changed because the source says 'red'.")
comment.reply("Checked against the source; agreed.")
comment.resolved = True
```

### 15.1 What landed

| Piece | Where | Signature |
|---|---|---|
| the part half | `docx4j_py/openpackaging/parts/wml/comments.py` (335 lines) | `CommentParts(main, comments_part, comments)` with `comments_ex`, `people`, `ids_part`, `extensible_part`, `added`, `part_names`; `comment_parts_of(main) -> CommentParts | None`, `create_comment_parts(main) -> CommentParts`; the lxml accessors `durable_ids_in_use`, `durable_id_for`, `set_comment_id`, `remove_comment_id`, `set_extensible`, `remove_extensible`; `COMMENTS_IDS_NS`, `COMMENTS_EXTENSIBLE_NS`, `NEW_COMMENTS_IGNORABLE` |
| the view half | `docx4j_py/model/content/comments.py` (1,255 lines) | `Author(name, initials=None, email=None)` and `initials_of`, `DEFAULT_AUTHOR`, `author_of`; `CommentMarker(id, kind, item, owner, run, run_owner, paragraph, paragraph_container, offset)`, `markers_of`, `markers_of_paragraph`, `comment_ids_of`; `next_comment_id`; `ensure_comment_styles(package) -> tuple[str, ...]`; `Comment(element, parts, body)` with `id`, `author_name`, `initials`, `author_email`, `creation_date`, `para_id`, `comment_body`, `paragraphs`, `content` (get and set), `resolved` (get and set), `replies`, `parent`, `anchor_address`, `reply(text)`, `delete()`, `get_range() -> list[Range]`, `get_xml()`, `to_dict()`; `insert_comment_into(range, text)`, `comments_of(scope)` |
| on the views | `.body`, `.paragraph`, `.range`, `.controls` | `Body.get_comments`, `Paragraph.get_comments` / `insert_comment`, `Range.get_comments` / `insert_comment`, `ContentControl.get_comments` |
| on the package | `.__init__`'s `register()`, and `OpcPackage` | `pkg.author` (a property installed by `register()`), over the `_author` slot |
| the trial | `.trial` | `TrialPart` copies an lxml part as well as a typed one (`tree`, `set_tree`), and answers `comments_extended_part`, `comments_ids_part`, `comments_extensible_part` and `people_part`; `TrialPackage.author`; `discard_added_parts` clears the part shortcut an un-added part left behind |
| the fixture | `tests/fixtures/comments-modern.docx` | docx4j's `loadAndSave.docx`, the one document in any of the three checkouts with all five comment parts |

The **`parts/wml.py` module became a `parts/wml/` package**, which is the only structural change
outside the content layer: `__init__.py` is what `wml.py` was, character for character, so every
import path, `wml.__all__` and `test_every_typed_part_names_a_model_class_that_exists` are
unchanged. `codegen/clean.py`'s `KEEP` lists `openpackaging` whole, so it needed nothing.

**No registries.** docx4j-core-ts needed two (`setCommentPartsAccess`, `setCommentApi`) to keep
its import graph acyclic. Here the direction was already settled by section 5 --- the content
layer imports the parts layer and never the other way round --- so
`model/content/comments.py` simply imports `openpackaging/parts/wml/comments.py`, and the part
module knows nothing of `Comment`. That is 60 lines of seam the Python port does not have.

### 15.2 Departures from sections 3.9 and 4, all deliberate

1. **The comment ids come from the document's own state, not from `pkg.id_generator()`.**
   Section 3.9's "comment ids are their own space" is kept exactly --- `next_comment_id` is one
   above the highest `w:comment/@w:id` **and** the highest id on any marker in the body, and
   revision ids (Phase F) will be a different counter --- but the number is *derived*, as Phase C
   decided for part names and relationship ids (14.2 item 10). A random generator is the wrong
   tool for a counter that has a natural order. The `w14:paraId` of a comment's paragraphs and
   the `w16cid:durableId` **do** come from `pkg.id_generator()`, because those have no natural
   order; a test runs the whole audit-trail sequence twice under one seed and compares the saved
   bytes, every part but `word/comments.xml` exactly and that one with `w:date` masked, since the
   date a comment is written is the wall clock and not a property of the agent surface (12.6
   made the same exception for `dcterms:created`).
2. **Three styles are defined, not two.** Section 3.9 and section 4 both say "the two comment
   styles". `CommentText` and `CommentReference` are the two that are *referred to*; docx4j's
   `KnownStyles.xml` gives `CommentText` a `w:link w:val="CommentTextChar"`, and 14.9's whole
   point is that a definition should not dangle, so `CommentTextChar` goes in with them.
   `CommentSubject` and `CommentSubjectChar` are **not** added: Word stopped writing the subject
   line years ago, and nothing this phase writes refers to them. `ensure_comment_styles` returns
   the part names it touched, and reads the ids out of the part's *bytes* first
   (`style_ids_of`), so a document that already has the three keeps `/word/styles.xml` byte for
   byte --- which a test pins over `comments-modern.docx`.
3. **`Author` lives in the content layer; its storage is on `OpcPackage`.** 12.1 put state on
   `OpcPackage` "because it needs no import from this layer", and `pkg.author` breaks that rule:
   its *value* is a content-layer class. So the split is the same one `body` uses --- an
   `_author` slot on `OpcPackage`, and an `author` property installed on
   `WordprocessingMLPackage` by `register()`, through a new `_PACKAGE_PROPERTIES` table beside
   `_PACKAGE_MEMBERS` (a property is not a function, so it could not go in the old one). The
   setter also accepts a bare string (`pkg.author = "Claude"`), and refuses anything else with
   `code="author.invalid"`. `Author` is **not** a top-level `docx4j_py` export: that would mean
   regenerating `docx4j_py/__init__.py`'s footer for one name, and
   `from docx4j_py.model.content import Author` is where every other view comes from.
4. **`delete()` reads the thread from `w15:paraIdParent`, not from `replies`.** The TypeScript
   engine walks its in-memory reply list, so deleting a `Comment` built any way but through
   `getComments()` leaves its replies behind. `Comment._thread()` walks the w15 entries instead,
   which is the document's own answer and is correct for a view built anywhere. A test deletes a
   comment whose reply this view was never told about.
5. **A reply's `w:commentRangeEnd` goes *after* its parent's, not before.** Word nests the starts
   outward (`<start 0/><start 1/>`) and the ends inward (`<end 1/><end 0/>`); this writes
   `<end 0/><end 1/>`, as docx4j-core-ts does. Both anchor exactly the same text, and Word shows
   the thread correctly either way; keeping the two engines identical is worth more than
   matching Word's own byte order here. The three insertions are made **last first** --- the
   reference run, the end, the start --- because all three are usually in one list.
6. **`ContentControl.get_comments()` is implemented**, though `tests/office_js_subset.json` has
   no such member: Office JS's `ContentControl` has no `getComments`, but section 3.2 lists it,
   and a bound control an agent has just filled in is where a comment about it belongs. Marked an
   extension in the docstring, as section 3.13 requires.
7. **`Comment.anchor_address` is an extension section 3.9 does not list**, and it is what
   `repr`, `to_dict()` and `change.touched()` report: a comment's address is the address of the
   paragraph its first marker is in. `comment_body`, `paragraphs`, `para_id`, `element`,
   `get_xml()` and `to_dict()` are the other extensions, as they are in the TypeScript engine.
8. **`Comment.id` is `-1` for a `w:comment` with no `w:id`** rather than raising. `w:id` is
   required by the schema and Word always writes it; a document that has lost one should still
   be readable, and `-1` matches no marker, so such a comment is simply never in scope.
9. **A comment in a header or a footer is accepted**, as the TypeScript engine accepts it, and
   the `w:comment` goes into the *document's* comments part, which is the only one there is.
   Word does not write comments in headers and will not show them; a test pins the behaviour and
   says so rather than refusing, because refusing would cost an error code an agent has to learn
   for a case it will not meet.
10. **A newly created `w:comments` carries `mc:Ignorable="w14 w15"`, and the two w15 parts carry
    none.** Word writes `mc:Ignorable` on every part; on `w15:commentsEx` and `w15:people` it
    names prefixes that are not the root's own, and naming `w15` there would ask Word to ignore
    the root element. `w:comments` gets it because its paragraphs carry `w14:paraId`; both
    prefixes are in docx4j's prefix table, so `XmlPart.xml` declares them (the check section 3.9
    asked for).
11. **`w16cid` and `w16cex` are edited as lxml**, because the registry gives them
    `DefaultXmlPart` (CR-001 section 13.5). That is three small accessors per part rather than a
    schema patch, and the alternative --- adding the two namespaces to the schema closure --- is
    a CR-001 decision, not this phase's.
12. **`tests/agent/test_comments.py` is `tests/agent/test_comment_workflows.py`.** pytest cannot
    collect two test modules with the same basename without turning the test directories into
    packages, and every other pair in this repository already avoids it
    (`content/test_markdown.py` and `agent/test_markdown_workflows.py`).

### 15.3 The trial, and the parts a comment adds

`insert_comment` creates up to four parts, so 14.4's undo log is what makes a dry run honest, and
two things had to be added to it:

- **`TrialPart` now copies an lxml part.** `w16cid:commentsIds` is a `DefaultXmlPart`, which has
  a `tree` and no `contents`; a trial that did not copy it would write its entries into the real
  document. The copy is `copy.deepcopy` of the element, and parsing the real part to make it
  costs **nothing in bytes**: a test in CR-002's round trip already shows that lxml writes back
  what it read, and this phase measured both Word-written parts of `comments-modern.docx` as
  byte-identical after a parse and a serialise.
- **`discard_added_parts` clears the shortcut.** A picture part has no shortcut on the document
  part; a comments part has five of them (`main.comments_part` and its siblings). Removing the
  part without clearing `main.comments_part` left the next real call holding a part the package
  no longer had. One `set_part_shortcut(None, relationship_type)` per un-added part.
- `TrialPackage.author` is **trial-local**: reading it falls through to the real package, setting
  it does not reach back. A trial is a preview, not a place to change the document's identity.

`test_a_dry_run_of_insert_comment_leaves_the_package_byte_for_byte` runs the whole verb on
`2010-sample1.docx`, then asserts that all four parts are gone, their content-type overrides with
them, `main.comments_part` is None again, `pkg.body.get_comments()` is empty and the saved bytes
--- the relationship parts included --- are what they were. 12.5's **first** limitation is
unchanged and the second test says so in a comment: a trial's first look at `/word/styles.xml`
unmarshals the real part, so a caller who wants the bytes compared must read that part first, as
the caller would have anyway.

### 15.4 The numbers, measured

Against `samples/2010-sample1.docx` loaded and its body read, and against the 200-page document
of section 7 (2,000 paragraphs, 20 tables) with a comment on every tenth paragraph:

| call | |
|---|---:|
| `insert_comment`, first one on a document with **no** comment parts | **577 µs** |
| `insert_comment`, second one (the parts are there) | **167 µs** |
| `insert_paragraph`, for comparison (12.4) | 17 µs |
| `get_comments()` on `comments-modern.docx`, one comment, three parts unmarshalled | **1.04 ms** |
| `insert_comment` on the 200-page document, 200 of them | **3.5 ms** each |
| of which `next_comment_id`, a marker walk of the whole body | **2.9 ms** |
| of which `_para_ids_taken`, a walk of the comments part | 0.04 ms |
| `get_comments()` over 205 comments on the 200-page document | **6.0 ms** |
| of which the same marker walk | 2.9 ms |
| a two-comment thread's `to_dict()` as compact JSON | **461 bytes** |

The first call's 577 µs is the four parts, their relationships, their content types and the three
style definitions; the second call's 167 µs is the real cost of a comment, and two thirds of that
is splitting the runs at the span's boundaries (`segments_of` twice, `split_at` twice), which is
the same work `Range.font` does.

**`next_comment_id` is this phase's `next_drawing_id`** (14.5): a walk of the whole body, 2.9 ms
at 200 pages and rising, and it is *both* verbs' cost, because `get_comments` walks the markers
too. The same fix applies --- cache the highest id and the marker index per body on the package,
as `_para_ids_taken` caches the paraIds --- and the same judgement: one comment in a 200-page
document at 3.5 ms did not justify it, and a phase that comments in a loop should. The 200-comment
run above took 0.70 s in all, which is the honest shape of the quadratic.

The thread's 461 bytes is well inside the 2 KB a tool result should cost, and a test pins it.
Nothing here changes the budgets of 12.3: `outline()` gained no field, and `stats.comments` was
already counted with lxml from the comments part's bytes (12.2 item 5), which a test now shows
agrees with `len(get_comments())` counting replies.

### 15.5 The fixture decision

Section 7 asks for "a document with comments and replies, copied from docx4j's samples". There
is **no such document**. Every `.docx` in `~/git/docx4j`, `~/git/docx4j-core-ts` and this
repository was checked: fourteen have a `w:comments` part, exactly one ---
`docx4j-core-tests/src/test/resources/loadAndSave.docx` --- has all five, and it has a single
comment and no reply. docx4j-core-ts's own Phase G fixture, `comments-two.docx`, is two
*unrelated* comments in a `w:comments` part alone.

So the phase uses two:

- **`tests/fixtures/comments-modern.docx`**, `loadAndSave.docx` copied verbatim (Apache-2.0,
  55 KB). It is what Word writes: `mc:Ignorable` on every part, a `w15:commentEx` with
  `w15:done="0"`, a `w15:person` whose `w15:userId` is Word's Active-Directory
  `S::address::guid` form, a `w16cid:durableId` and a `w16cex:dateUtc`. It is what the
  byte-for-byte tests and the email, paraId and done-flag reads run against.
- **`conftest.threaded_package()`**, built here from XML written in the open. Two comments on one
  range, the second a **reply** through `w15:paraIdParent`, and a third **resolved** on another
  paragraph, with all five parts. It is declared as hand-built in `tests/README.md` and in its
  own comment; it is not presented as a Word document, because it is not one. If a genuine
  Word-written thread turns up it should replace it, and the tests would not change.

### 15.6 What Phase F needs

- **The marker hoist is written and tested, and is the piece Phase F inherits.**
  `_marker_site(paragraph, segment, after=)` answers "where does a marker for this run go": beside
  the run, but **outside** the `w:ins` / `w:del` / `w:moveTo` / `w:moveFrom` the run sits in,
  found through `segment.run.parent` and that holder's own parent. It is why
  `test_the_markers_are_hoisted_out_of_a_tracked_insertion` passes over
  `samples/sample-docx.docx`. What Phase F must not do is undo it: when tracking is on, a
  `w:commentRangeStart` must still be written **outside** any `w:ins` the anchor is in, and the
  reference run must not itself be wrapped in a `w:ins`, or accepting the revision would take the
  comment with it. The cost, recorded so it is not read as a bug: a comment on *part* of an
  insertion widens to the whole of it, which is what Word shows once the insertion is accepted.
- **`pkg.author` is the shared setting**, already installed and documented as such
  (`Author.name` / `initials` / `email`; `initials_of` derives the initials). Phase F should read
  `author_of(package)` from `model/content/comments.py` rather than add a second setting, and
  `pkg.tracked_change_date` (section 3.8) is the only identity field still missing.
- **The two id spaces are separate and stay separate** (section 4). `next_comment_id` is over
  `w:comment/@w:id` and the comment markers only; Phase F's revision counter is over every
  `CTMarkup` `w:id` in the parts unmarshalled, **excluding** `w:comment/@w:id`. Nothing in this
  phase touches a revision id.
- **The text model already hides deleted content from the markers.** `_visit_runs` enters a
  `w:ins` and a `w:moveTo` and not a `w:del` or a `w:moveFrom`, so a marker's offset is an
  offset into the *accepted* view, which is what `Range` uses. A `get_comments()` over the
  original view is not offered and should not be: a `Range`'s original view is refused already
  (section 4).
- **`Comment.delete()` is not tracked**, as `ContentControl.delete` is not (section 4, 14.7). A
  comment is not content; deleting one leaves no revision markup, and Phase F should leave that
  as it is.

### 15.7 What Word said about artefact 7 (2026-09-17)

The Word check passed, with two observations worth keeping. The comments pane lists threads in
**document order of their anchors**, not insertion order: the whole-paragraph comment comes first
because its `w:commentRangeStart` sits at the paragraph's start, before the range on "first"
inside the same paragraph. And on save Word's Compatibility Checker reported "Comments which have
been collapsed will no longer be collapsed", one occurrence: `samples/2010-sample1.docx` is a
Word 2010 document (`w:compatSetting compatibilityMode` 14) and a resolved thread (`w15:done`)
is a Word 2013 feature, so Word warns that saving in the older mode drops it. It is Word's own
downgrade warning, the same one a person resolving a comment in that document would see, and not
a defect in the markup. A `resolved = True` on a document whose compatibility mode is below 15
could add a `ChangeReport.warnings` line saying so, which would cost one lxml read of
`settings.xml`; not done, noted for whoever finds an agent confused by the dialog.


## 16. Phase F implementation notes (2026-09-17)

Phase F is done: change tracking as section 3.8 specifies it and section 4's "Tracking rules,
Word's not just the markup" rules it --- `pkg.change_tracking_mode` over `w:trackRevisions`,
`pkg.tracked_change_date`, the revision markup every mutation writes in the paragraph-level
primitives, `TrackedChange` with `accept()` and `reject()`, `get_tracked_changes()` on a body, a
paragraph, a range and the package, `accept_all()` / `reject_all()`, and `replace_text` at all
three levels --- with the tests of section 7, acceptance artefact 8 and the README's audit
trail, which now leads with both halves. The suite is **1,131 tests** (1,130 passing plus one
`xfail`; 1,122 of them fast, 54.8 s, and 64.4 s for the whole), against 1,067 at the end of
Phase G; **64** of the new ones are `tests/content/test_tracking.py` (58) and
`tests/agent/test_audit_trail.py` (6). Nothing in `~/git/docx4j-xsdata` changed, no schema patch
was needed, the model was not regenerated and `codegen/generate_el.py` did not change at all:
every Phase F name is reached through `docx4j_py.model.content`.

```python
from docx4j_py import load
from docx4j_py.model.content import Author

pkg = load("in.docx")
pkg.author = Author("Claude", initials="C")
pkg.change_tracking_mode = "TrackAll"
pkg.body.replace_text("colour", "color")     # a w:del and a w:ins each
pkg.get_tracked_changes()[0].to_dict()
pkg.body.accept_all()
```

### 16.1 What landed

| Piece | Where | Signature |
|---|---|---|
| the tracker | `docx4j_py/model/content/tracking.py` (825 lines) | `ChangeTracker(package, mode)` with `author`, `date`, `markup_roots()`, `next_id()`, `markup()`, `track_change()`, `ins(items)`, `deletion(items)`, `own_insertion(revision)`, `own_paragraph(p)`, `assert_editable(revision)`, `record_r_pr_change(r_pr)`, `record_p_pr_change(p_pr)`, `mark_paragraph_inserted` / `_deleted`, `mark_row_inserted` / `_deleted`; `tracker_of(package)`, `Revision` and `revision_of(run)`, `FontTracking`, `track_inserted_paragraph`, `wrap_new_runs`, `track_inserted_table`, `para_r_pr_of`, `row_pr_of`, `prune_paragraph_properties`, `mark_deleted` / `mark_inserted`, `to_deleted_text` / `to_restored_text`, `copy_r_pr`, `restore_r_pr`, `restore_p_pr`, `xml_date` / `date_of`, `mode_of` / `set_mode` |
| the view | `docx4j_py/model/content/tracked_change.py` (662 lines) | `TrackedChangeTarget(kind, value, element, owner, mark)`; `TrackedChange(target, paragraph, body)` with `type`, `author`, `date`, `text`, `get_range()`, `accept()`, `reject()`, `to_dict()`, and `id`, `element`, `kind`, `target`, `address` as extensions; `tracked_changes_of_paragraph`, `tracked_changes_of_row`, `tracked_changes_of_body`, `join_with_next` |
| the values | `.enums` | `ChangeTrackingModeValue` and `ChangeTracking`, `TrackedChangeTypeValue` and `TrackedChangeType`, with `CHANGE_TRACKING_MODES` and `TRACKED_CHANGE_TYPES` beside them |
| the error | `.errors` | `TrackedChangeError(ContentError)`, codes `tracking.deleted_text`, `tracking.already_deleted`, `tracking.gone`, `tracking.no_paragraph` |
| on the package | `.__init__`'s `register()`, and `OpcPackage` | `pkg.change_tracking_mode` and `pkg.tracked_change_date` (properties over the `_change_tracking_mode`, `_change_tracker` and `_tracked_change_date` slots), `pkg.get_tracked_changes()` |
| on the views | `.body`, `.paragraph`, `.range` | `Body.change_tracker` / `get_tracked_changes` / `accept_all` / `reject_all`, `Paragraph.change_tracker` / `font_tracking` / `get_tracked_changes`, `Range.get_tracked_changes`; `Table.change_tracker` |
| the tracked primitives | `.paragraph` (+397 lines) | `Paragraph.splice`'s tracked branch, `_delete_text`, `_insert_tracked`, `_isolate`, `_split_run_before` / `_after`, `_tracked_delete`, and the hook in the one `_p_pr()` accessor |
| the trial | `.trial` | `TrialPackage.change_tracking_mode` (trial-local), `tracked_change_date`, `get_tracked_changes()`, `document_settings_part` (the real part until written to) and `writable_settings_part()` |
| the fixture | `tests/fixtures/tracked-pprchange.docx`, `tests/content/conftest.py` | docx4j's `unmarshallFromTemplateDirtyExample.docx`; `moves_package()` for the forms no document carries |

`Body.insert_element` and `Paragraph.splice` are where the whole phase hangs from. Everything
else --- `Range.insert_text`, `Range.delete`, the `text` setter, `insert_xml`, `insert_ooxml`,
`insert_markdown`, `insert_table`, `ContentControl.insert_text`, `TableCell.value` --- already
ran through those two, so it was tracked the day the branch landed and the tests only had to
prove it (`test_a_content_control_inherits_tracking_and_its_delete_does_not`,
`test_markdown_inserted_while_tracking_is_on_is_one_report_and_all_insertions`).

### 16.2 Departures from sections 3.8 and 4, all deliberate

1. **`TrackedChange` is its own module**, `tracked_change.py`, as section 3.8 allowed
   ("or in `tracking.py`; record the choice"). The two halves share nothing but the small
   helpers `tracking.py` exports, they are 825 and 662 lines, and the writing half must not
   import the reading half (a primitive never needs a view). The CR's section 5 table says
   "`ChangeTracker`, `TrackedChange` | `docx4j_py.model.content.tracking`"; both names are
   re-exported from `docx4j_py.model.content`, which is where every other view comes from, so
   nothing outside this repository can tell.
2. **Reading the mode does not unmarshal the settings part.** Section 3.8 says "reading or
   writing the mode unmarshals the settings part"; reading it with lxml instead --- the same
   `_root(part)` `describe()`'s `tracking_on` uses (12.2 item 5) --- costs 3.2 ms once and
   **0.2 µs** thereafter (the answer is cached on the package), and it keeps a promise worth
   more than the simplicity: a document whose mode is only *read* saves `/word/settings.xml`
   byte for byte, which `test_reading_the_mode_leaves_the_settings_part_byte_for_byte` pins.
   Writing does unmarshal, because changing the flag means re-marshalling the part anyway.
   The cache is invalidated by the setter and is the only state the mode has.
3. **A tracked deletion does not reach a field instruction.** `w:instrText` is not text (the
   text model excludes it, 11.2 item 11), so a span never covers one and `_isolate` splits it
   into a run of its own: deleting "the whole field" through a `Range` deletes its result and
   leaves the instruction. Word deletes both. `to_deleted_text` / `to_restored_text` do rename
   `w:instrText` to `w:delInstrText` and back, and a test exercises them directly, so the
   markup half of section 4's rule is kept and the field-awareness is what is missing; a phase
   that owns fields should close it.
4. **`TrackMineOnly` is remembered for the session and stored as `TrackAll`.** Section 4 says
   only the second half. Reading the mode back in the same session gives what was asked for
   (the tracker carries it), and a reload gives `TrackAll`, which a test states in both
   directions. The distinction changes nothing this engine does: it matters only when a second
   author edits the same package, which is Office JS's concern.
5. **Accepting or rejecting away a table's last row takes the table with it.** docx4j's
   `AcceptTrackedChanges` leaves the husk, because it is a conversion preprocessor and a
   `w:tbl` with no `w:tr` never reaches the renderer's output. This document is saved again,
   and an empty `w:tbl` is not valid WordprocessingML, so `_remove_row` drops it. That is the
   same reasoning section 4 already applied to `w:rPrChange` and `w:pPrChange`.
6. **A paragraph this author inserted records no `w:pPrChange`, and a run inside a `w:ins` of
   ours no `w:rPrChange`.** The second is the TypeScript engine's `ownInsertions` rule; the
   first is its counterpart for `_p_pr()` and is new here, and it is what stops
   `body.insert_paragraph("x", style="Heading 1")` writing a formatting revision on a paragraph
   that did not exist a moment ago. `ChangeTracker.own_paragraph(p)` is the test.
7. **`replace_text` reports the pair, not the text either side.** `text_before` is `find` and
   `text_after` is `replace`, at body, paragraph and range level, because one call may touch
   two thousand paragraphs and "the text before" would then be one arbitrary paragraph's. The
   count is the return value; `ChangeReport` has no field for it and gained none. The three
   verbs open one `recording("replace_text")` and the nested ones are no-ops inside it, so a
   whole-body replace is one report (12.4's rule).
8. **`OutlineStats.tracked_changes` is counted through the view**, not structurally. Phase D
   counted `w:ins`, `w:del` and `w:rPrChange` by element name over `iter_nodes`; a paragraph
   mark's `w:ins` is a `CTTrackChange` in a single-valued field and `element_name` answers
   `None` for it (CR-001's warning about `XmlMeta.qname`, in another guise), so that count
   could never have agreed with `get_tracked_changes()`. Counting the collection instead makes
   them agree by construction **and is six times cheaper** (16.4).
9. **A `TrackedChange`'s `text` for a deletion is the deleted text.** Section 3.8 says only
   "text". `_revision_text` reads a `w:delText` inside the revision it belongs to, which is
   what the original view shows and what `to_markdown(view="markup")` renders between `{--`
   and `--}`; a test asserts the two agree over `samples/sample-docx.docx`.
10. **`TrackedChange.type` never answers `"None"`**, though the `Literal` carries it, because
    every piece of markup a change is over is an insertion, a deletion or a formatting change.
    Office JS's `"Unknown"` is not in the `Literal` at all: section 3.8 lists four values and
    `"Unknown"` is not one of them.
11. **A trial's mode is the trial's.** `TrialPackage.change_tracking_mode` falls through to the
    real package until the trial sets one, and setting it writes `w:trackRevisions` into the
    trial's **copy** of the settings part (`writable_settings_part()`), so
    `with pkg.dry_run() as trial: trial.change_tracking_mode = "TrackAll"` leaves the real
    document's mode alone. Reading it still reads the real part's bytes, so `describe()` on a
    trial costs no unmarshalling --- which is how the Phase K dry-run test stayed green.
    12.5's first limitation is unchanged: a trial that *writes* the mode unmarshals the real
    settings part, as it does for any part it copies.
12. **Corrected 2026-09-17 (16.10): "an inserted paragraph carries its own mark" holds in the
    middle of a container and nowhere else.** Section 4 and docx4j-core-ts's phase F both say a
    paragraph this API inserts marks its **own** mark, and Phase F shipped that everywhere. At
    the **end** of a container it is wrong and Word cannot handle it; the rule is now Word's own
    (mark the preceding paragraph's mark, leave the final one alone), and 16.10 says what Word
    did with the version that was not.

### 16.3 The one thing Word does that this does not

A `w:moveFrom` / `w:moveTo` pair is **read** (both `TrackedChange`s, both views, accept and
reject, the `moves_package()` fixture) and never **written**: nothing in the content API moves
content as a move. Word writes the pair when a user drags a selection; an agent's equivalent is
a delete and an insert, which is what this writes. `w:moveFromRangeStart` and its three
siblings are carried through untouched, as they always were. Section 3.8 asks for the forms to
be covered by `TrackedChange`, which they are.

### 16.4 The numbers, measured

Against the 200-page document of section 7 (2,000 paragraphs, 20 three-by-three tables), loaded
and its body read, with `pkg.tracked_change_date` fixed:

| call | |
|---|---:|
| `insert_text`, untracked | 18 µs |
| `insert_text`, **tracked** | **45 µs** |
| the **first** tracked call, which pays for the id scan | **7.4 ms** |
| the second | 81 µs |
| `replace_text("lazy", "energetic")` over the whole document, untracked (1,950 matches) | 186 ms |
| the same, **tracked** | **366 ms** |
| `get_tracked_changes()` over the 3,900 revisions it made | **45 ms** |
| `accept_all()` over those 3,900 | **30 ms** |
| `reject_all()` over those 3,900 | 39 ms |
| `outline()` with Phase D's structural counter | 20.6 ms |
| `outline()` with Phase F's counter (16.2 item 8) | **13.9 ms** |
| `change_tracking_mode`, first read (an lxml parse of `settings.xml`) | 3.2 ms |
| `change_tracking_mode`, every read after it | **0.2 µs** |
| `change_tracking_mode = "TrackAll"` (unmarshals the part) | 17.4 ms |

A tracked edit costs **two and a half times** an untracked one, and a tracked `replace_text`
just under twice: the extra is the run isolation (two `split_at`s and the run splitting, which
is the work `Range.font` already does) and the two wrappers. That is the honest price of
writing revision markup, and it is paid per edit rather than per document.

**The id scan is the one fixed cost**, 7.4 ms on a 200-page document, and it is paid **once per
package**: `ChangeTracker` caches the counter, and the tracker itself is cached on the package,
so the second tracked call is 81 µs. It walks every node of every part already unmarshalled and
tests `isinstance(node, CTMarkup)`, which the generated model makes exact --- `CTBookmark`,
`CTMarkupRange`, `CTTrackChange`, `CTRPrChange` and the rest all extend it, and
`CommentsComment` extends it too and is the one exclusion (section 4). A part the package has
not read is not unmarshalled for an id; the annotation ids of an untouched header cannot collide
with an edit to the body in any way Word minds.

`accept_all()` over 3,900 changes in 30 ms is a tenth of the edits that made them, because it is
one pass in reverse document order over a list already collected: no re-collection, no
re-walking, and a paragraph join never disturbs a change still to do.

Nothing here changes the budgets of 12.3: `outline()` gained no field and got faster, and
`TrackedChange.to_dict()` is 180 bytes of compact JSON, so the 20 an agent would show cost under
4 KB.

### 16.5 The fixture decision

Section 7 asks for "a document with tracked changes, copied from docx4j's samples". Every
`.docx` in `~/git/docx4j`, `~/git/docx4j-core-ts` and this repository was searched for the eight
forms. The result:

- `samples/sample-docx.docx` has one `w:ins` and one `w:del`, by "Jason Harrop", 2007.
  docx4j-core-ts's own Phase F fixture, `test/fixtures/tracked-changes.docx`, is **the same
  document** (the same ids, author and dates), so copying it would have added nothing.
- **`tests/fixtures/tracked-pprchange.docx`** is docx4j's
  `docx4j-samples-docx4j/sample-docs/unmarshallFromTemplateDirtyExample.docx` copied verbatim
  (Apache-2.0, 17 KB), and it is the **only** document in any of the three checkouts with a
  `w:pPrChange`. It is Word-written and also carries a paragraph mark marked inserted, which is
  the other form no other sample has.
- A **move**, a **`w:rPrChange`** and a **tracked row** are in none of them. So
  `tests/content/conftest.py`'s `moves_package()` builds one, in the open, from XML written to
  match what Word writes: a `w:moveFrom` / `w:moveTo` pair, a run whose `w:rPrChange` records
  the italic it used to be, an inserted row and a deleted row. It is declared as hand-built in
  `tests/README.md` and in its own docstring, exactly as Phase G's `threaded_package()` is
  (15.5). If a genuine Word-written one turns up it should replace it, and the tests would not
  change.

### 16.6 What Phase E needs

- **A bound control's `insert_text` writes through to the custom XML node** (section 4), and it
  now runs through a `Paragraph.splice` that may be writing revision markup. The write-through
  hook belongs *outside* the tracked branch: the data node takes the new text whatever the mode
  is, because Word refreshes a bound control from the data on open and a `w:del` in the control
  does not change the data. 14.6 named `ContentControl._extend` and `Paragraph.splice` as the
  one place; it still is.
- **`ContentControl.delete()` stays untracked** (section 4, 14.7), and a test now pins it while
  the mode is on.
- **The two id spaces are three.** Revision ids are `ChangeTracker.next_id()` over every
  `CTMarkup`, comment ids are `next_comment_id`, and `w:sdt/@w:id` is `next_sdt_id(root)`
  (Phase A, 10.3 item 3), which is a fourth thing again. A typed control Phase E creates should
  take its id from the last of those, not from the tracker.

### 16.7 What Phase H needs

- **A list paragraph's `w:numPr` is a paragraph property**, so `attach_to_list` and
  `detach_from_list` must go through `Paragraph._p_pr()` like every other property setter, and
  they will then record a `w:pPrChange` while tracking is on --- which is what Word does for a
  numbering change on an existing paragraph. `start_new_list()` on a paragraph this author
  inserted records nothing, by 16.2 item 6.
- **`w:numberingChange` is not written.** It is the deprecated numbering revision
  (ECMA-376 17.13.5.19) and Word writes `w:pPrChange` instead; `TrackedChange` does not offer
  it, and a document that carries one is left alone.
- **The numbering part is not tracked at all**: a revision is body markup, and a definition
  added to `/word/numbering.xml` is a part edit. Nothing in Phase H needs a tracked branch
  beyond the paragraph property one.

### 16.8 What the Word check is for

`scripts/acceptance.py` writes `8-tracked-changes.docx`: `samples/2010-sample1.docx` with the
author set, the mode on, a replacement, an inserted paragraph, a deleted paragraph, a bold run,
a table row added and one deleted, and a comment on the replacement. Two paragraphs and the
table are written **before** the mode goes on, so that there is something of the document's own
to delete and to re-format. What to look for is in `tests/README.md`: the Review pane listing
each change with its author and date, Accept All leaving the intended document, Reject All
restoring the pre-edit one, the comment surviving both, and a save-close-reopen with no repair
prompt. **Not yet run.** The Compatibility Checker note of 15.7 may appear again, for the same
reason: the source is a Word 2010 document.

### 16.9 A defect found the same day: a comment inside a same-author insertion

**The defect.** A comment on text that lay **wholly inside** this author's own `w:ins` --- which
is exactly what commenting on a `replace_text` gives, the commonest thing the audit trail of
section 3.4 does --- had its `w:commentRangeStart`, its `w:commentRangeEnd` **and its reference
run written inside the `w:ins`**. Accepting the revision was harmless, but rejecting it deleted
the comment with the insertion, which is precisely what section 15.6's marker hoist exists to
prevent.

**The cause was not in Phase G's hoist**, which is correct and untouched. It was a stale parent
pointer in `Paragraph._delete_text`. Moving the runs into the new `w:del` is
`ChangeTracker.deletion(items)`, which sets each run's `parent` to the `w:del`; the code then
asked `group[0].run.parent` for "what held this run", got the `w:del` it had just made, and gave
both the `w:del` and the `w:ins` of the replacement that as their `parent`. The two elements sat
in the right place in `w:p/@content` --- the XML was valid and every text test passed --- but
`_marker_site` walks `segment.run.parent` and then **that** element's parent to find the list to
hoist into, found the `w:del` instead of the `w:p`, could not find the `w:ins` in the `w:del`'s
own list, and fell through to its un-hoisted branch. `revision_of` was reading the same wrong
`owner`, so the fix repairs that too.

**The fix** is four lines: `_Target` now captures `parent` at collection time, before any run is
moved, and the `w:del` (and the anchor the following `w:ins` is placed against) take that. Three
tests pin it --- the markers and the reference run are **siblings** of the `w:ins` and the two
revisions are children of the `w:p`; the comment survives `accept_all()` with its range still
reading `"report"`; it survives `reject_all()` with an **empty** range where it was, and
`get_comments()` still lists it --- and a fourth does the same for a comment on a **whole
paragraph inserted while tracking is on**, where rejecting joins the paragraph into the one
before it and carries the markers along.

The lesson for anything else that moves elements while tracking: **read a parent before the
move, never after**. `wrap_new_runs`, `Paragraph.insert_break` and `Body.insert_break` set the
wrapper's parent explicitly and were never affected; `_insert_tracked`'s non-anchor path reads a
parent nothing has moved.

The suite was **1,125 tests** at that point (1,124 passing plus one `xfail`); the four new ones
are `tests/content/test_tracking.py`'s. Section 16.10, the same day, took it further.

### 16.10 What the Word check of artefact 8 found (2026-09-17)

The Word check of `8-tracked-changes.docx` **failed**, in three ways, and two rules came out of
it. Both are Word's, neither is docx4j's --- docx4j's `AcceptTrackedChanges` reads revision
markup and never writes any, so it could not have said.

**1. A deleted row was shown pink and not struck through.** The row carried
`w:trPr/w:del` and nothing else: its cells' runs were still plain `w:r`/`w:t` and their paragraph
marks were unmarked. Word writes a deleted row as the row mark **and** every run of every cell in
a `w:del` with `w:delText` **and** every cell paragraph's mark marked `w:pPr/w:rPr/w:del` --- the
exact mirror of the inserted row this phase already wrote correctly. A row marked but not emptied
is a row Word knows is going and whose text it still believes is live, which is what the pink
without strikethrough was saying.

`ChangeTracker.mark_row_deleted` is unchanged (it writes the row mark, and nothing else should);
the new `track_deleted_row(tracker, row, body)` wraps it and empties the row through
`Paragraph.delete_text_tracked` and `mark_paragraph_deleted` --- the **same primitives a
paragraph deletion uses**, so accepting and rejecting need no special case. `Table.delete_rows`,
`Table.delete`, `TableRow.delete` and `Body.clear` all go through it. A row **this author
inserted** is now taken back whole instead, as deleting text this author inserted is (section 4);
the caller removes it and there is no revision left behind.

**The row is still one `TrackedChange`, and that is a decision.** The cell-level revisions a row
insertion or deletion is made of --- same author, same direction --- are **folded into** the row's
change, because that is what Word's Reviewing pane shows and what an agent means by "the row I
deleted"; `folded_in_row(change)` is the function, accepting or rejecting the row processes them,
and anything else inside the row (another author's edit, a formatting revision) is **listed
separately and is not touched**. A three-cell deleted row is therefore one change and not seven.

**2 and 3. The appended paragraph could not be seen, and Reject All hung Word.** The inserted
paragraph was the body's last block and carried its own `w:pPr/w:rPr/w:ins`. **Word never marks a
container's final paragraph mark as inserted**, because that mark cannot be deleted: pressing
Enter at the end of the last paragraph marks the **preceding** paragraph's mark and gives the new
paragraph the original, unmarked final mark. A final mark marked inserted is a revision Word
cannot reject --- it hung, and the changes stayed as they were.

So `track_inserted_blocks(tracker, elements, container)` replaces the per-paragraph call in
`Body.insert_element`, which is where `insert_paragraph`, `insert_xml`, `insert_ooxml` and
`insert_markdown` all arrive. A paragraph with nothing after it but the other paragraphs of the
same insert **shifts its mark back one**: the paragraph before it takes it. For a fragment of
three appended blocks that marks the paragraph that was already there and the first two inserted
ones, and leaves the third --- the one that now ends the container --- unmarked, which is exactly
what Word writes for three presses of Enter. In the **middle** of a container nothing changes:
each inserted paragraph carries its own mark, as section 4 says and as docx4j-core-ts does. With
no paragraph before it (an empty container, or a table) nothing is marked and only the runs are
wrapped: there is no earlier mark for the break to live on, and rejecting then leaves the
paragraph empty rather than removing it.

Three things follow from the shift:

- **The take-back follows the mark.** `insertion_mark_of(paragraph, container)` answers "which
  mark records this paragraph's insertion" --- its own, or the previous paragraph's when it is
  last --- and `Paragraph.delete()` clears that one when it takes its own insertion back.
- **Rejecting an inserted mark keeps the surviving paragraph's own properties.**
  `join_with_next(..., keep_properties=True)`. docx4j's rule --- the joined paragraph takes the
  *next* one's `w:pPr` --- is right for **accepting a deleted mark**, where the mark that
  survives is the next one's; in the reject direction the paragraph that survives is one the
  document already had and nobody edited, and taking the new paragraph's (usually empty)
  properties would strip its style. A test appends to a `Heading1` paragraph and rejects it.
- **No fixture could be copied.** Every Word-written `.docx` in the three checkouts was searched
  again: none has a deleted table row, and none has a final paragraph whose mark is marked
  inserted --- which is itself the evidence that Word never writes one. The expected markup in
  the tests is therefore built from Word's rules and says so, as `moves_package()` does (16.5).

**The numbers.** The suite is **1,131 tests** (1,130 passing plus one `xfail`; 1,122 of them
fast, 54.8 s, and 64.4 s for the whole); `tests/content/test_tracking.py` is 58. Seven of its
tests pinned the form Word rejected --- an inserted paragraph's own mark at the end of a body, and
six changes for an inserted two-by-two table --- and were **changed to the new rules**, not the
other way round. Artefact 8 was rebuilt: the inserted paragraph now goes **after the first
paragraph**, where a human can see it, a second one is appended at the very end so that the
final-mark rule is exercised, and the table is written before both so that the body ends with a
paragraph. Accepting gives the intended document, rejecting gives back exactly the document that
was there before the mode went on, and the comment survives both.

### 16.11 The last of it: the mark goes with the break it recorded

The Word check's second rule (16.10) left one thing half done, and a second check of the
regenerated artefact found it: **rejecting an inserted paragraph mark joined the paragraphs but
left the mark behind**, so `reject_all()` over a loaded document ended with the text right, an
empty paragraph where the inserted one had been, and `get_tracked_changes()` still listing the
two `mark` changes. Accepting was clean throughout.

The cause is `keep_properties`, introduced in 16.10. docx4j's join replaces the surviving
paragraph's properties with the **next** one's, and that wholesale replacement is what used to
carry the `w:pPr/w:rPr/w:ins` away; keeping the surviving paragraph's own properties --- which
16.10 had to do, so that rejecting an appended paragraph does not strip the formatting of a
paragraph nobody edited --- left the mark sitting in them. `TrackedChange.reject()` now drops it
explicitly after the join, which is what "the break this mark recorded is gone" means. One line,
and `join_with_next`'s docstring says whose job it is.

Both forms were affected and neither was covered: the own-mark form in the middle of a container
and the shifted-back form at the end. The tests that did pass passed because they asserted the
**text** after rejecting and not the markup, and because they ran over a created package whose
paragraphs had no properties to keep. `test_rejecting_both_mark_forms_of_a_loaded_document_leaves_the_source`
now runs the coordinator's own reproduction --- a **loaded** document, a
paragraph inserted in the middle and one appended, rejected through `reject_all()` **and** one
change at a time in reverse --- and asserts three things the older tests did not: the body is the
source document, `get_tracked_changes()` is **empty**, and no `w:ins` or `w:del` is left in the
part. The two neighbouring tests gained the same two assertions.

The lesson, beside 16.10's: **a test that rejects must assert the markup is gone, not only that
the text came back.** Text equality passed through every version of this defect.
