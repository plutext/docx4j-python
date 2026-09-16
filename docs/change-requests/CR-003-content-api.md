# CR-003: A content API for Python, for agents and MCP, in the vocabulary of Office JS

**Status:** Proposed 2026-09-16; revised the same day to follow Python conventions throughout
and to be designed for AI projects and MCP servers first.
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
| A | Tree layer additions (3.3) | 2 days |
| B | `Body`, `Paragraph`, `Range`, `Font`: the verbs, `style` per section 4, `search` across runs with grapheme-safe splitting, `insert_xml`, `insert_element`; the error hierarchy; the Office JS subset list | 4 days |
| D | The agent surface (3.4): addresses, `Outline` with budgets, `find()`, `describe()`, `ChangeReport`, `dry_run`, `DocumentSession`, determinism; the agent scenario tests | 4 days |
| K | Markdown out, with addresses, then in (3.5) | 3 days |
| C | `Table`, `TableRow`, `TableCell`, `InlinePicture` with the header readers, `insert_ooxml`, `ContentControl` reads and `delete` | 4 days |
| G | Comments (3.9) | 3 days |
| F | Change tracking and `replace_text` (3.8); the README's audit-trail example | 4 days |
| E | Custom XML, mapping, typed controls, `describe()` / `fill()` (3.7) | 4 days |
| J | The python-docx facade (3.6) and its subset test | 3 days |
| H | Lists (3.10) | 3 days |
| I | `to_api_script` (3.11) | 2 days |

B, D and K first: after them an agent can read any document within a budget, address any block,
edit it and see what changed, which is the MCP server's whole first release. G before F, as the
TypeScript engine found (tracking has to hoist comment markers). J after E so the facade sits on
a complete core. Each phase ships as a minor version.

## 9. Open questions

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
   imported lazily; `to_markdown` needs nothing.
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
