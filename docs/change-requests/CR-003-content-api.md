# CR-003: A content API in the shape of Office JS, over the docx4j tree

**Status:** Proposed 2026-09-16.
**Depends on:** CR-001 Phases A to C (the model, `el`, the builders, `wml(...)`, `text_of`,
`walk`, `find`) and CR-002 Phase A (packages, parts, load and save), both implemented. Effective
formatting (`Font` reads as Office JS reports it) and list labels need CR-002 Phase B
(`PropertyResolver`, the numbering `Emulator`); until it lands the views read direct formatting
and say so.
**Counterpart:** docx4j-core-ts [CR-002](../../../docx4j-core-ts/docs/change-requests/CR-002-content-api.md)
(phases B, C, D, E, F, G and I implemented 2026-09-10 to 2026-09-16; H proposed), whose design this
CR takes as given and whose implementation notes (its sections 7 to 15) are folded into section 4
here as rules, so the Python engine does not rediscover them; docx4j `MainDocumentPart.addParagraphOfText`
/ `addStyledParagraphOfText` / `addObject` / `getContent`, `TraversalUtil`, `ClassFinder`,
`TextUtils`, `BindingHandler`, `AcceptTrackedChanges`; the tool surface of `plutext/docx4j-mcp`.

## 1. Why

CR-001's Hello World is honest and low-level:

```python
body.content.append(el.p(content=[el.r(content=[el.t("Hello World")])]))
```

The sugar (`p("Hello World", style="Heading1")`) removes the typing, but the tree is still the
whole API: there is no "add a paragraph after the one that says X", no "find the paragraph that
says X", no "make this range bold" that splits runs at the boundaries, no addresses an agent can
hold across tool calls. python-docx has the first of those and nothing typed underneath; the
Java docx4j has the tree and four helpers on the part.

Two kinds of consumer are in view, the same two as in TypeScript. A **developer** who knows Word
documents and wants the common operations to read the way they do in docx4j, python-docx or
Office JS. An **agent** (Claude through an MCP server, or code an LLM writes against this
package) that wants few operations, strings in and out, stable addresses for "the third
paragraph", and no way to produce an invalid tree by accident.

And a third reason that is Python's own: docx4j-core-ts and docx4j-python are one design in two
languages. Where the TypeScript content API has settled a name, a shape or a rule, this CR uses
it, so that the documentation, the MCP tool surface and the agents' training material read across
between the two, and a fix found in one is a fix for both.

## 2. What others do

| Library | Building | Reading and editing | Shape |
|---|---|---|---|
| docx4j (Java) | `mdp.addParagraphOfText("x")`, `addStyledParagraphOfText`, `addObject(o)`; `XmlUtils.unmarshalString` | `getContent()`, `TraversalUtil`, `ClassFinder`, XPath via a `Binder`, `TextUtils.extractText` | methods on the part; the tree everywhere else |
| Office JS | `body.insertParagraph(text, "End")`, `insertText`, `insertOoxml`, `insertTable(r, c, "Start", values)` | `body.paragraphs`, `body.text`, `body.search(text, options)` returning ranges | fluent proxies with insert locations `Start` / `End` / `Replace` / `Before` / `After` |
| python-docx | `document.add_paragraph("x", style="Heading 1")`, `paragraph.add_run("bold").bold = True` | `document.paragraphs`, `paragraph.text`, `paragraph.runs` | methods on document and paragraph, properties for formatting; a thin typed slice over lxml |
| docx4j-core-ts | `body.insertParagraph`, `p()`, `el.p()`, `insertXml`, `insertOoxml` | `body.paragraphs`, `search` across runs, `Range.font`, `outline()`, addresses, comments, tracked changes, custom XML mapping | Office JS's shapes, structurally assignable to a `Word.*` subset |
| Office-Word-MCP-Server and docx4j-mcp | `add_paragraph(file, text, style)`, `add_table`, `insert_line_or_paragraph_near_text` | `get_document_outline`, `find_text_in_document`, `search_and_replace`, `format_text(paragraph_index, start, end, ...)` | flat tools: text in, text out, an index or "near this text" as the address |

The convergence is what matters: a handful of verbs on the body or the part, addresses and string
payloads for agents. python-docx is the library a Python developer will compare this to, and its
vocabulary (`add_paragraph`, `paragraph.text`, `paragraph.runs`, `run.bold`) is close enough to
Office JS's that a python-docx user reads this API without a glossary.

## 3. Design: the shape of Office JS, in Python

Decision, carried from docx4j-core-ts: where Office JS has a name or a shape for an operation,
this API uses it. What is mimicked is the object model and the method surface, not the
execution model: Office JS proxies batch and `sync()`; here everything is in memory and
synchronous, including what marshals, since Python has no asynchronous import to wait on
(CR-002 section 1). Office JS code ported here drops `load` and `sync` and the `await`s.

Three Python-specific rules:

- **Names are Office JS's vocabulary in `snake_case`**: `insert_paragraph`, `insert_text`,
  `style_built_in`, `get_range`, `parent_table_cell`. PEP 8 and CR-001's field naming
  (`p_pr`, `r_pr`) leave no room for camelCase methods, and the vocabulary is what carries the
  parity, not the casing. The insert-location and enum *values* stay exactly Office JS's strings
  (`"Start"`, `"End"`, `"Before"`, `"After"`, `"Replace"`, `"Centered"`, `"Single"`,
  `"TrackAll"`, `"Heading1"`), because they are data, they appear in agents' tool arguments, and
  CR-001's builders already take them (`underline="Single"`). docx4j's helper names are aliases
  in `snake_case` (`add_paragraph_of_text`, `add_styled_paragraph_of_text`, `add_object`,
  `get_content`).
- **Collections are lists.** `body.paragraphs` is a `list[Paragraph]` in document order, so
  `len`, indexing, slicing, iteration and `[-1]` all work; there is no `items` and no
  `get_first()`. Views are light objects holding the element and its container, created on
  access and cheap to discard; two views of one element compare equal.
- **Offsets are code points**, since that is what a Python `str` indexes; Office JS uses UTF-16
  code units. Addresses (3.3) never carry offsets, so nothing crosses the boundary except a
  `Range`, which is Python's business. Grapheme clusters are respected when a run is split
  (3.10).

Units follow Office JS: points on the views (`left_indent`, `space_after`, `font.size`), twips
and half-points in the tree, converted on read and write.

### 3.1 The surface (`docx4j_py/model/content/`)

```python
class Body:                       # Word.Body (a subset): on WordprocessingMLPackage.body, and on any
                                  # part or container whose root has a content list (main document,
                                  # headers, footers, footnotes, endnotes, comments, a table cell, a
                                  # block content control)
    paragraphs: list[Paragraph]                        # document order; tables and content controls descended
    tables: list[Table]
    content_controls: list[ContentControl]             # nested ones included
    inline_pictures: list[InlinePicture]
    text: str                                          # a paragraph per line; the accepted view (3.6); no text boxes
    def get_text(self, *, view: Literal["accepted", "original"] = "accepted") -> str
    def insert_paragraph(self, text: str, location: Literal["Start", "End"]) -> Paragraph
    def insert_text(self, text: str, location: Literal["Start", "End", "Replace"]) -> Range
    def insert_table(self, row_count: int, column_count: int, location: Literal["Start", "End"], values: list[list[str]] | None = None) -> Table
    def insert_break(self, type: Literal["Page", "Line", "SectionNext", "SectionContinuous"], location: Literal["Start", "End"]) -> None
    def insert_inline_picture_from_base64(self, base64: str, location: Literal["Start", "End"], *, width: float | None = None, height: float | None = None, alt_text_description: str = "", alt_text_title: str | None = None) -> InlinePicture
    def insert_inline_picture(self, data: bytes, location, **options) -> InlinePicture        # extension: bytes in
    def insert_ooxml(self, ooxml: str, location: Literal["Start", "End", "Replace"]) -> list[Paragraph | Table]   # pkg:package, as Word; or a bare fragment
    def insert_xml(self, xml: str, location) -> list[Paragraph | Table]                       # extension: a w:p / w:tbl fragment, prefixes supplied
    def insert_element(self, element: Child, location: Literal["Start", "End", "Before", "After"], target: Paragraph | Table | None = None) -> None   # docx4j addObject
    def insert_content_control(self, kind: ContentControlType = "RichText") -> ContentControl # wraps the body's content
    def search(self, text: str, *, match_case: bool = False, match_whole_word: bool = False, match_wildcards: bool = False) -> list[Range]
    def replace_text(self, find: str, replace: str, **search_options) -> int
    def clear(self) -> None
    def get_range(self, location: Literal["Whole", "Start", "End", "Content"] = "Whole") -> Range
    def get_comments(self) -> list[Comment]
    def get_tracked_changes(self) -> list[TrackedChange]
    def accept_all(self) -> None;  def reject_all(self) -> None
    def get_xml(self) -> str
    # extensions
    content: ChildList                                 # the live list (docx4j getContent()); sectPr excluded
    part: XmlPart                                      # the part this body belongs to
    def element_at(self, address: str) -> Paragraph | Table | ContentControl | None
    def paragraph_at(self, address: str | None = None, *, contains: str | None = None) -> Paragraph | None
    def address_of(self, view) -> str
    # docx4j aliases
    add_paragraph_of_text, add_styled_paragraph_of_text, add_object, get_content

class Paragraph:                  # Word.Paragraph (a subset)
    text: str                                          # read: runs joined (w:t, w:tab, w:br, w:sym); write: one run replaces them
    style: str                                         # the style's display name ("Heading 1"), see section 4
    style_built_in: str                                # the Word.Style value ("Heading1"), or "Other"
    style_id: str | None                               # extension: w:pStyle as docx4j thinks of it
    alignment: Literal["Left", "Centered", "Right", "Justified", "Unknown"]
    font: Font                                         # over the runs; setting applies to every run
    left_indent: float; right_indent: float; first_line_indent: float
    space_after: float; space_before: float; line_spacing: float     # points, as Office JS
    outline_level: int
    is_list_item: bool; list_item: ListItem | None; list: List | None   # phase H
    parent_body: Body
    parent_table_cell: TableCell | None
    parent_content_control: ContentControl | None
    inline_pictures: list[InlinePicture]
    content_controls: list[ContentControl]
    def insert_text(self, text, location: Literal["Start", "End", "Replace"]) -> Range
    def insert_paragraph(self, text, location: Literal["Before", "After"]) -> Paragraph
    def insert_break(self, type, location) -> None
    def insert_inline_picture_from_base64(self, base64, location: Literal["Start", "End", "Replace"], **options) -> InlinePicture
    def insert_ooxml(self, ooxml, location) -> list[...];  def insert_xml(self, xml, location) -> list[...]
    def insert_content_control(self, kind="RichText") -> ContentControl            # wraps the paragraph
    def insert_comment(self, text: str) -> Comment
    def search(self, text, **options) -> list[Range]
    def replace_text(self, find, replace, **options) -> int
    def get_range(self, location="Whole") -> Range
    def get_text(self, *, view="accepted") -> str
    def get_comments(self) -> list[Comment];  def get_tracked_changes(self) -> list[TrackedChange]
    def start_new_list(self) -> List;  def attach_to_list(self, list_id: int, level: int) -> None;  def detach_from_list(self) -> None
    def get_xml(self) -> str
    def delete(self) -> None
    # extensions
    element: P                                         # the tree
    para_id: str | None                                # w14:paraId, the stable address (3.3)
    runs: list[R]

class Range:                      # Word.Range (a subset): a span within one paragraph, or a whole paragraph
    text: str
    font: Font
    paragraphs: list[Paragraph]
    style: str; style_built_in: str; style_id: str | None
    def insert_text(self, text, location: Literal["Before", "After", "Start", "End", "Replace"]) -> Range
    def insert_paragraph(self, text, location: Literal["Before", "After"]) -> Paragraph
    def insert_ooxml(self, ooxml, location: Literal["Before", "After", "Replace"]) -> list[...]
    def insert_content_control(self, kind="RichText") -> ContentControl            # wraps the runs, split at the boundaries
    def insert_comment(self, text) -> Comment
    def search(self, text, **options) -> list[Range]
    def replace_text(self, find, replace, **options) -> int
    def get_comments(self) -> list[Comment];  def get_tracked_changes(self) -> list[TrackedChange]
    def get_xml(self) -> str                          # a copy of the paragraph trimmed to the span
    def delete(self) -> None
    runs: list[R]                                      # extension: the runs it spans, split at the boundaries when set

class Font:                       # Word.Font (a subset), backed by w:rPr
    bold: bool; italic: bool; strike_through: bool; subscript: bool; superscript: bool
    underline: Literal["None", "Single", "Double", "Dotted", "Wave", "Thick", ...]
    name: str;  size: float;  color: str;  highlight_color: str | None
    # phase B of CR-002 makes reads effective (styles resolved); until then direct formatting of the first run

class Table:                      # Word.Table (a subset), backed by w:tbl
    row_count: int; rows: list[TableRow]; values: list[list[str]]
    style: str; style_built_in: str; style_id: str | None; header_row_count: int
    text: str; parent_table_cell: TableCell | None
    def get_cell(self, row_index: int, cell_index: int) -> TableCell
    def add_rows(self, location: Literal["Start", "End"], row_count: int, values=None) -> list[TableRow]
    def delete_rows(self, row_index: int, row_count: int = 1) -> None
    def column_widths(self) -> list[float]
    def delete(self) -> None
    element: Tbl

class TableRow:   row_index, cell_count, cells: list[TableCell], values, is_header, insert_rows(location, count, values), delete()
class TableCell:  body: Body, paragraphs, tables, text, value, insert_paragraph, insert_text, row_index, cell_index, parent_row, parent_table, width, column_width

class InlinePicture:              # Word.InlinePicture (a subset)
    width: float; height: float                        # points
    alt_text_description: str; alt_text_title: str | None; image_format: str
    def get_base64(self) -> str;  def get_bytes(self) -> bytes
    def delete(self) -> None
    paragraph: Paragraph; rel_id: str; inline: Inline; image_part: ImagePart

class ContentControl:             # Word.ContentControl (a subset), backed by w:sdt in all four forms
    type: ContentControlType                           # "RichText" | "PlainText" | "Picture" | "DatePicker" | "ComboBox" | "DropDownList" | "CheckBox" | "RepeatingSection" | "RepeatingSectionItem" | "Group" | "BuildingBlockGallery" | ...
    form: Literal["block", "run", "row", "cell"]       # extension
    tag: str; title: str; id: int; text: str; placeholder_text: str
    appearance: Literal["BoundingBox", "Tags", "Hidden"]; color: str; cannot_delete: bool; cannot_edit: bool; remove_when_edited: bool
    xml_mapping: XmlMapping
    paragraphs, tables, content_controls, inline_pictures
    checkbox_content_control, date_picker_content_control, drop_down_list_content_control, combo_box_content_control, picture_content_control, repeating_section_content_control, group_content_control   # the kind-specific views, None for the other kinds
    def insert_text(self, text, location) -> Range;  def insert_paragraph(...);  def search(...);  def get_range(...);  def get_xml(self) -> str
    def delete(self, keep_content: bool) -> None
    element: SdtBlock | SdtRun | CTSdtRow | CTSdtCell
```

`WordprocessingMLPackage` gains `body` (the main document part's), `outline()`, `paragraph_at`,
`element_at`, `author`, `change_tracking_mode`, `tracked_change_date`, `get_tracked_changes()`,
`custom_xml_parts`; `MainDocumentPart`, `HeaderPart`, `FooterPart`, `FootnotesPart`,
`EndnotesPart` and `CommentsPart` gain `body`.

The Hello World:

```python
from docx4j_py import create_package

pkg = create_package(page_size="A4")
pkg.body.insert_paragraph("Hello World", "End")
pkg.save("hello.docx")
```

And the same code against a live document in Word, for comparison:

```ts
await Word.run(async (context) => { context.document.body.insertParagraph('Hello World', 'End'); await context.sync(); });
```

Editing what is there:

```python
body = pkg.body
title = body.insert_paragraph("Report", "Start")
title.style_built_in = "Heading1"                 # title.style reads "Heading 1"
title.alignment = "Centered"
hit = body.search("quick brown fox")[0]           # across runs
hit.font.italic = True                            # runs split at the boundaries
hit.insert_text("slow red fox", "Replace")        # keeps the first replaced character's formatting
body.paragraphs[-1].insert_paragraph("The end.", "After")
body.text                                         # a paragraph per line
```

### 3.2 Tree layer additions (CR-001's `docx4j_py.wml.builders`)

The TypeScript objects package's CR-003 found, one phase at a time, that the content API wrote
builders it should have had from the tree layer. They are specified here up front, as Phase A,
and live with `p`, `r`, `t` and `tbl` because they need only the tree:

- `tr(cells, *, widths=None)` and `tc(blocks, *, width=None)`, so `add_rows` and `insert_rows`
  build one row of the table's widths.
- `inline_picture(rel_id, *, cx, cy, id, name, descr="", title=None)`: the `w:drawing` /
  `wp:inline` exactly as docx4j's `BinaryPartAbstractImage.createImageInline` writes it. The
  README's image example becomes `body.content.append(p(r(inline_picture(rel.id, cx=..., cy=...))))`.
- `sdt(content, *, kind, tag=None, title=None, id=None, form=None)`, `sdt_pr(**options)`,
  `next_sdt_id(root)`, `sdt_property(sdt_pr, local_name, namespace=W)` and `sdt_kind_of(sdt_pr)`:
  the four `w:sdt` forms and the choice list `w:sdtPr` is (docx4j's own shape). `sdt_property`
  looks in both `w:` and `w15:`, because Word writes `w15:dataBinding` on a repeating section.
- `rpr_to_elements(rpr)` and `rpr_from_elements(elements)`: `w:rPrChange/w:rPr` is a base-type
  element list where `w:rPr` has named properties, in the schema's `EG_RPrBase` order, w14
  effects included.
- `deep_copy_as(obj, cls)`: a copy typed as a base class (`PPrBase` for `w:pPrChange/w:pPr`).
  xsdata serialises by the declared field type and writes `xsi:type` when the instance's class
  differs, which is valid but not what Word writes; the copy must be the declared type.
- `walk_all(root, visitor, wildcard_visitor)`: `walk` that also enters wildcard (`AnyElement`)
  content, for anything that rewrites references (an SVG twin's `r:embed` in an `a:extLst`).
- `run_items_of(holder)`: the run list of a `w:p`, `w:hyperlink`, `w:ins`, `w:del`,
  `w:moveFrom`, `w:moveTo`, `w:smartTag`, `w:customXml`, `w:sdtContent` (run form), by the
  model's field names, so the text model reads inside tracked insertions.

### 3.3 Addresses and the agent surface

Three address forms, accepted wherever a `Paragraph | Table` target is, and returned by
`outline()`; the same forms as docx4j-core-ts, so an MCP tool set can serve both engines:

- **Ordinal:** `"body/3"` (the fourth block-level child of the body), `"body/3/1"` (a row, or a
  run), `"body/4/0/1/0"` (a paragraph in a cell), `"header:rId3/0"`. Stable until the content
  changes. Block-level children are paragraphs, tables, block content controls and `altChunk`s.
- **paraId:** `"w14:5A2B1C3D"` where Word wrote one; new paragraphs get one when the document
  already uses them. Stable across edits.
- **Text:** `paragraph_at(contains="Chapter 1")`, first match, as the MCP servers' `near_text`.

```python
pkg.outline()
# {"paragraphs": [{"address": "body/0", "para_id": "5A2B1C3D", "style_id": "Heading1", "text": "Chapter 1"}, ...],
#  "tables": [{"address": "body/4", "rows": 3, "cols": 2}], "headers": [...], "footers": [...]}
pkg.paragraph_at("body/3").insert_paragraph("New", "After")
pkg.paragraph_at(contains="Chapter 2").delete()
```

The MCP verbs map onto one-liners: `add_paragraph` is `body.insert_paragraph(text, "End")` plus
`style`; `insert_line_or_paragraph_near_text` is `body.search(text)[0].paragraphs[0].insert_paragraph(line, "After")`;
`delete_paragraph(index)` is `body.paragraph_at("body/3").delete()`; `get_document_outline` is
`outline()`; `search_and_replace` is `body.replace_text(a, b)`. A Python MCP server over this
package is a product (docx4j-mcp is Java; a `docx4j-mcp-py` would sit next to it) and not this CR.

### 3.4 Compatibility with Office JS, enforced

TypeScript enforces the subset with a compile-time assignability check against
`@types/office-js`. Python has no such check, so the contract is a **committed member list**:
`tests/office_js_subset.json`, generated by `scripts/office_js_subset.py` from docx4j-core-ts's
`test/office-js-subset.ts` (the interfaces, their members, property or method, and which are
extensions), snake_cased. A test asserts every non-extension member exists here with the same
kind, and lists what this engine adds beyond the TypeScript one so the two READMEs can say the
same thing. Regenerate when the TypeScript subset grows; the script reads the sibling checkout
when present and the committed file otherwise.

### 3.5 Custom XML parts, XML mapping and typed content controls

The same surface as docx4j-core-ts section 3.5, WordApiDesktop 1.3's `CustomXmlPart`,
`CustomXmlNode`, `XmlMapping` and the typed content-control kinds, over the same parts
(`CustomXmlDataStoragePart` is an lxml tree already, CR-002 section 5.4). It is simpler here in
one respect that removes a whole design section: **XPath is lxml's**, XPath 1.0 complete with a
namespace map, so there is no `XPathEngine`, no optional dependency, no `load()` to warm and
nothing asynchronous. `select_nodes(xpath, namespace_mappings=None)` takes Word's
`xmlns:ns0='...'` prefix-mapping string or the part's namespace manager.

```python
pkg.custom_xml_parts                     # list-like: items, get_by_namespace(ns), get_item(id), add(xml)
part = pkg.custom_xml_parts.get_item("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}")
part.select_single_node("/ns0:invoice/ns0:customer/ns0:name", "xmlns:ns0='urn:invoice'").text = "Acme"
control.xml_mapping.set_mapping("/ns0:invoice/ns0:total", "xmlns:ns0='urn:invoice'", part)
pkg.custom_xml_parts.apply_bindings()             # docx4j BindingHandler.applyBindings, what Word does on open
pkg.custom_xml_parts.update_from_content_controls()
```

`CustomXmlNode.xpath` reports the canonical `/ns0:a[1]/ns0:b[2]` form Word writes. Reading a
custom XML part does not mark it for re-marshalling; a mutation through a node view does
(`DefaultXmlPart.mark_modified()`), so a document whose data parts are only read keeps them byte
for byte.

### 3.6 Change tracking and find-and-replace

`pkg.change_tracking_mode` (`"Off"` | `"TrackAll"` | `"TrackMineOnly"`) over `w:trackRevisions`
in the settings part, created when the document has none; `pkg.author` (shared with comments)
and `pkg.tracked_change_date` (now, unless fixed). While the mode is on, every mutation of the
content API writes Word's revision markup instead of editing in place, in the paragraph-level
primitives so that `Range`, `Body`, `Table` and later phases inherit it: `w:ins`, `w:del` with
`w:delText`, a replacement as `w:del` then `w:ins`, an inserted paragraph's mark, a deleted
one's, `w:rPrChange` and `w:pPrChange` recorded once on the first write, `w:trPr/w:ins` and
`w:trPr/w:del` for rows. `TrackedChange` views (`type`, `author`, `date`, `text`, `accept()`,
`reject()`, `get_range()`) from `get_tracked_changes()` on body, paragraph and range;
`accept_all()` and `reject_all()` on the body, accepting as docx4j's `AcceptTrackedChanges`
does. The text model reads the accepted view throughout (`w:ins` in, `w:del` out), which is what
a reviewer and an agent want; `get_text(view="original")` on body and paragraph gives the other.
`replace_text(find, replace, **options)` is `search` then `insert_text(..., "Replace")` from the
last match to the first, tracked when the mode is on.

Reading or writing the mode unmarshals the settings part, which then re-marshals on save; that
is CR-002's rule and is acceptable because the settings part round-trips canonically (CR-002
section 12.4). No deferred write is needed, since nothing here is asynchronous.

### 3.7 Comments

`get_comments()` on body, paragraph and range; `insert_comment(text)` on paragraph and range;
`Comment` with `author_name`, `author_email`, `initials`, `content`, `creation_date`
(`datetime | None`), `resolved` (`w15:done`), `replies`, `reply(text)`, `delete()`,
`get_range()`, over the parts a current Word writes: `w:comments`, `w15:commentsEx` (threads
by `w15:paraIdParent`), `w16cid:commentsIds`, `w:people`, and `w16cex:commentsExtensible` kept
in step when present. Any of the first four the document lacks is created with its
relationship, content type and the `CommentText` and `CommentReference` styles. Reads unmarshal
three parts, writes five; a document whose comments are not read keeps all of them byte for
byte. The author is `pkg.author = Author(name, initials=None, email=None)`, default
`Author("docx4j-python")`.

### 3.8 Lists

`List` over a `w:num` and its `w:abstractNum` (`id`, `level_types`, `level_exists`,
`get_level_paragraphs`, `set_level_numbering`, `set_level_bullet`, `set_level_indents`, copying
the abstract definition first when other `w:num`s share it), `ListItem` (`level`, `list_string`,
`sibling_index`), and on `Paragraph`: `is_list_item`, `list`, `list_item`, `start_new_list()`
(docx4j's default definitions, `NumberingDefinitionsPart.unmarshal_default_numbering`, creating
the numbering part when absent), `attach_to_list(list_id, level)`, `detach_from_list()`.
`list_string` and `sibling_index` need CR-002 Phase B's `Emulator`; until then `list_string` is
`None`. A paragraph whose style carries `w:numPr` is a list item through the style.

### 3.9 `to_api_script`, and no `Word` shim

docx4j-core-ts's phase I is a `Word` object whose `Word.run(pkg, fn)` runs add-in code against
a package: proxies, `ClientResult`, `sync()`, the enums. Its purpose is running Office JS code
unchanged in Node, and it has no Python counterpart worth building: there is no Office JS code
in Python to run. What does carry over is **`to_api_script(target)`**, the reveal-codes
generator: the content-API calls, as Python source against `body`, that produce a body, a
paragraph, a range or a `w:p` / `w:tbl` element (`insert_paragraph` with the text, `style_built_in`,
`alignment`, the indents, one `insert_text` per run with the `font` assignments that differ from
the previous run, `insert_break`, `insert_table` for a plain grid), falling back to
`body.insert_xml(...)` with the marshalled fragment, decided per paragraph so that a paragraph
comes out whole either way. It is how an agent learns the API from a document it is shown, and
how an editor shows what a selection is. `to_source(element)`, emitting `el.p(...)` / `p(...)`
calls, is the tree-layer complement and a CR-001 follow-up.

### 3.10 Scripts and directionality

Offsets are code points (section 3). A split (`Paragraph.split_at`, behind `Range.font`,
`insert_text` at a range boundary, `insert_comment`, `insert_content_control`) must never fall
inside a grapheme cluster: `unicodedata` gives combining classes and the ZWJ and variation
selectors are a short list, which covers emoji sequences and Indic scripts without a
dependency; the `regex` module's `\X` is used when installed. `match_whole_word` uses Unicode
word boundaries (`\b` under `re` with `str` patterns is Unicode-aware); `match_case=False` uses
`re.IGNORECASE`, which is Unicode case-insensitive. The run mapping writes `w:bCs`, `w:iCs` and
`w:szCs` with their Latin twins, so complex-script text is formatted by the same calls;
`alignment` maps `start` / `end` as well as `left` / `right`, and the indent properties read
both `w:left` / `w:right` and the strict-form `w:start` / `w:end`, writing whichever the paragraph
already uses. `bidi` on `Paragraph` and `rtl` on `Font`.

### 3.11 What stays as it is

- `contents` remains the typed tree and `content` lists remain `ChildList`s. Nothing is wrapped
  in proxies; a user who appends to `body.content` directly keeps working (the `ChildList` links
  the parent; `deep_copy(subtree)` for a copy).
- The packaging layer keeps docx4j's names; only the content API takes Office JS's. Where
  Office JS has no name (the tree, addresses, `insert_xml`, `outline`, `get_bytes`), the
  extension is marked as such in the docstring.
- Everything mutates the tree in place, marks the part for re-marshalling (CR-002's rule), and
  links parents. Inserts reject an element the container cannot hold, naming what was passed:
  the guard an agent needs so that a wrong tree fails at insert time, not in Word.
- `text_of` (CR-001) and `Body.text` are two things: `text_of` is docx4j's `TextUtils` over any
  subtree and decides its own text-box rule (CR-002 open question 7); `Body.text` is Word's
  main story, the accepted view, no text boxes, and is what agents and the outline report.

## 4. What the TypeScript implementation learned, adopted here from the start

Each item below cost docx4j-core-ts a correction or a departure note (its CR-002 sections 7 to
15). They are rules here, not discoveries to make again.

- **`style` is the display name, `style_built_in` the `Word.Style` value.** `style` reads the
  styles part's `w:name` when that part is unmarshalled (nothing is unmarshalled for it), with
  Word's stored lower-case built-in names (`heading 1`, `toc 1`, `annotation text`) mapped to
  display names through docx4j's `KnownStyles.xml`; otherwise the id with spaces inserted. Setting
  accepts a display name, a stored name or an id. `style_built_in` reads `"Other"` for a style
  that is not built in and refuses to be set to it. `style_id` is the docx4j-named extension;
  `outline()` reports ids. `BUILT_IN_STYLES` is one committed list (`model/content/styles.py`).
- **`insert_xml`, `insert_ooxml` and `insert_element` return the inserted views**, since a
  fragment or a package may bring several blocks. At paragraph and range level a fragment of
  exactly one `w:p` has its runs merged into the paragraph, which is what Word's paste does.
- **`Body.tables` returns `Table` views** from the first phase, never elements.
- **Rows and cells descend into row- and cell-level content controls** (an OpenDoPE repeat
  wraps its `w:tr` in a `w:sdt`), so `rows` is what Word shows; `add_rows` puts new rows in the
  table's own content, never inside a repeat. Children of every `w:sdt` form are visible to
  `Body.paragraphs` and to addresses, and the four `sdtContent` levels are skipped in paths.
- **`header_row_count` counts leading rows** carrying `w:tblHeader`; `insert_table` sets no
  style, so a new table is borderless until `style_built_in = "TableGrid"`.
- **Pictures**: `/word/media/imageN.<ext>` with N free for that extension (docx4j's
  `getNewPartName`); a relationship from the part the body belongs to (main document, header or
  footer); the `wp:inline` from `inline_picture`; PNG, JPEG, GIF and BMP headers read for size
  and density (docx4j gets this from XML Graphics Commons' `ImageInfo`), 96 dpi default, EMU as
  `px / dpi * 914400`, an image wider than the text area scaled down as `CxCy.scale` does. No
  Pillow dependency: the four header readers are a hundred lines.
- **`insert_ooxml` from a `pkg:package`** (CR-002 Phase C's flat OPC store): every part the
  content references is copied under a free name with a fresh relationship id, its own
  relationships copied recursively keeping their ids; references in the inserted content are
  rewritten by attribute name (`embed`, `link`, `id`, `href`, the chart and diagram ones) and
  only when the incoming package really has a relationship of that id, so numeric ids
  (`w:bookmarkStart/@w:id`, `wp:docPr/@id`) are never touched; the walk enters wildcard content
  (`walk_all`). Styles and numbering are not merged (that is docx4j's MergeDocx).
- **`ContentControl.get_range()`** is exact for a run-level control and the first paragraph's
  range for the block, row and cell forms, because a `Range` is within one paragraph. `type` is
  `"RichText"` when `w:sdtPr` names no kind, `w:text` is `"PlainText"`. A row- or cell-level
  control refuses `insert_paragraph` at `"Start"` / `"End"` and `insert_text(..., "Replace")`,
  naming what it holds. `insert_content_control` at range level refuses a span that crosses a
  run holder (a hyperlink, a tracked insertion); `RepeatingSection` is refused at run level.
- **Comments**: one `Comment` type for comments and replies (`parent` is an extension); `id` is
  the OOXML `w:id`, an `int`; `creation_date` may be `None`; `get_range()` is a `Range` per
  paragraph and a comment whose markers sit at block level gets none; an empty range gets a
  reference run only; no `w16cex` part is created, only kept in step; reactions and comments in
  headers are out of scope; the styles part is unmarshalled only when it lacks the two comment
  styles.
- **Tracking rules, Word's not just the markup**: a run already inside a `w:ins` by the same
  author is extended rather than nested; deleting text that author inserted takes it back;
  `w:del` before `w:ins` on a replacement; `w:t` becomes `w:delText` and `w:instrText`
  `w:delInstrText`; runs split at the span's boundaries. Revision ids come from a per-package
  counter above the highest `w:id` on any `CTMarkup` in the parts unmarshalled, **excluding**
  `w:comment/@w:id`, which is a separate space with its own allocator. A comment is not a
  revision: its markers are hoisted out of a `w:ins` or `w:del` the anchor run sits in, so
  accepting or rejecting leaves the comment in place. Deleting a content control is not
  tracked. `TrackMineOnly` is stored as `TrackAll`, since `w:trackRevisions` is a flag. An
  inserted paragraph carries its own mark. A deleted row stays in the tree until accepted, so
  `row_count` reports it, as Word does. `accept_all()` and `reject_all()` are one pass in reverse
  document order and, unlike docx4j's conversion preprocessor, accepting also drops
  `w:rPrChange` and `w:pPrChange`. Editing deleted text raises, naming the author.
- **Bindings**: a bound control's `insert_text` writes through to the custom XML node, because
  Word refreshes a bound control from the data on open; containers (`RepeatingSection`,
  `RepeatingSectionItem`, `Group`, `BuildingBlockGallery`, anything holding a table or another
  control) are never bound; pictures and explicit rich text are counted as skipped, not bound;
  dates are formatted with `w:dateFormat` in the `w:lid` locale and `w:fullDate` set, a checkbox
  gets `w14:checked` and the ☒ / ☐ glyph run from `w14:checkedState` / `uncheckedState`, a list
  shows the `display_text` whose `value` matches; the reverse writes `true` / `false`, the
  entry's value and the stored `w:fullDate`. Content goes into the **first paragraph** of a
  block, row or cell control, keeping its `w:pPr`; `w:placeholder` is kept; bindings are read
  from `w:dataBinding` **and** `w15:dataBinding`; the three well-known docProps store item ids
  are not special-cased; `placeholder_text` refuses to overwrite content; `add()` follows
  docx4j's `addPropertiesPart` (`/customXml/itemN.xml`, `itemPropsN.xml`, a brace-wrapped
  upper-case UUID, `ds:schemaRefs`, the relationship **from the main document part**).
- **`Paragraph.font` is over the runs**, not the paragraph mark; `alignment` is `"Unknown"` when
  `w:jc` is absent; `outline_level` is `w:outlineLvl + 1`, 10 when absent.
- **A `Range`'s original view is not offered** (`get_text(view="original")` raises on a range,
  since offsets are accepted-view offsets), and `search` has no `view` option.
- **`to_api_script` fidelity**: a `w:tab` is `insert_text("\t")`; `w:proofErr`,
  `w:lastRenderedPageBreak` and bookmarks are dropped; a `w:br` is a run of its own; anything a
  verb cannot express (`w:pPr` beyond `pStyle`, `jc`, `ind`, `spacing` and `outlineLvl`;
  `w:rPr` outside the `Font` vocabulary, `w:rStyle` included; a `w:lineRule` other than
  `exact`) falls back to `insert_xml`.

## 5. Where each piece lives

| Piece | Package | Why |
|---|---|---|
| `tr`, `tc`, `inline_picture`, `sdt` and friends, `rpr_to_elements`, `deep_copy_as`, `walk_all`, `run_items_of` | `docx4j_py.wml.builders` and `docx4j_py.child` / `traversal` (CR-001's tree layer) | tree only; usable without a package |
| `Body`, `Paragraph`, `Range`, `Font`, `Table`, `TableRow`, `TableCell`, `InlinePicture`, `ContentControl`, the docx4j aliases, `insert_xml`, `insert_ooxml`, the text model (`segments_of`, `split_at`) | `docx4j_py.model.content` | needs parts (image parts and relationships, headers and footers by relationship) |
| addresses, `outline()`, `paragraph_at`, `element_at` | `docx4j_py.model.content.addresses` and the package | needs the package (headers, footers, footnotes) |
| `CustomXmlPart`, `CustomXmlNode`, `XmlMapping`, the typed kinds, `apply_bindings`, `update_from_content_controls` | `docx4j_py.model.customxml` | needs parts (the data part, its properties part, `storeItemID`) |
| `ChangeTracker`, `TrackedChange`, accept and reject | `docx4j_py.model.content.tracking` | the paragraph primitives call it |
| `Comment`, the comment parts' loading and creation | `docx4j_py.model.content.comments` and `docx4j_py.openpackaging.parts.wml.comments` | half view, half part |
| `List`, `ListItem` | `docx4j_py.model.content.lists` | needs the numbering part and, for labels, CR-002 Phase B |
| `to_api_script` | `docx4j_py.model.content.api_script` | over the views |
| the Office JS subset list and its test | `tests/office_js_subset.json`, `scripts/office_js_subset.py` | the compatibility promise lives with the implementation |
| MCP tools over the above | a `docx4j-mcp-py` product | product, not library |

Import cycles: the parts layer must not import the content API statically (CR-002 keeps the
graph one-directional). `XmlPart.body` is provided through a registration the content module
performs on import, as the TypeScript engine does with `setCommentApi`; the content module
imports the parts, never the reverse.

## 6. What CR-002 Phase B supplies, and what this CR does without it

`Font` reads report direct formatting of the first run in scope until `PropertyResolver` lands,
then effective formatting as Office JS reports it; the property is documented as such and a test
is marked to flip. `ListItem.list_string` and `sibling_index` are `None` until the `Emulator`.
`insert_table` sizes the grid to the section's text width from `w:sectPr`, which needs nothing
from Phase B. Nothing else waits.

## 7. Tests

- **Fixtures:** the 16 documents in `samples/` plus, copied from docx4j's samples as needed,
  `invoice2013.docx` (already there: w15 repeating sections, a checkbox, a date, a picture
  control, twenty bindings), a document with comments and replies, one with tracked changes,
  one with lists, one with text boxes, one with a hyperlink and a field.
- **Per phase**, in `tests/content/`: the views' reads against known values; every mutation
  followed by save, reload and read back; a canonical comparison of the untouched parts (CR-002's
  rule holds through the content API); the Office JS subset test (3.4); `to_api_script` output
  executed against a fresh body and the result compared to the source paragraph.
- **Parity with docx4j-core-ts:** the same fixture, the same call sequence, the same saved
  `document.xml` canonically. A small script runs the TypeScript engine when the sibling
  checkout and Node are present (skipped otherwise) and compares; the differences it finds are
  either a bug in one engine or a departure to record in both CRs.
- **Word acceptance**, manual, added to `tests/README.md` per phase as the TypeScript checklist
  does (items 5 to 9 there: a content edit, an inserted picture, a comment thread, a tracked
  change accepted in Word, a bound control refreshed from its data).

## 8. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | Tree layer additions (3.2): `tr`, `tc`, `inline_picture`, the `sdt` family, `rpr_to_elements` / `rpr_from_elements`, `deep_copy_as`, `walk_all`, `run_items_of`; tests against marshalled XML | 2 days |
| B | `Body`, `Paragraph`, `Range`, `Font`: `insert_paragraph`, `insert_text`, `text`, `style` / `style_built_in` / `style_id` as section 4, `alignment`, indents and spacing in points, `search` across runs with grapheme-safe splitting, `delete`, `get_xml`, `insert_xml`, `insert_element`; the subset list and test; README rewritten around `pkg.body` | 4 days |
| D | Addresses (ordinal, paraId, text), `outline()`, `paragraph_at`, `element_at`, `address_of` | 1 day |
| C | `Table`, `TableRow`, `TableCell`, `insert_table`, `InlinePicture` and both `insert_inline_picture` forms with the header readers, `insert_ooxml` from a `pkg:package` (after CR-002 Phase C's flat OPC store; a bare fragment before it), `ContentControl` reads and `delete` | 4 days |
| E | `custom_xml_parts`, `CustomXmlPart` / `CustomXmlNode` over lxml, `XmlMapping`, the typed kinds, `insert_content_control`, `apply_bindings` / `update_from_content_controls` | 4 days |
| G | Comments (3.7) | 3 days |
| F | Change tracking and `replace_text` (3.6), `TrackedChange`, accept and reject | 4 days |
| H | Lists (3.8); labels once CR-002 Phase B lands | 3 days |
| I | `to_api_script` (3.9) | 2 days |

B and D first, in that order, because they are what an agent needs; C, E, G, F, H, I in the
order the TypeScript engine took them, which its notes justify (comments before tracking, since
tracking has to hoist comment markers). Each phase ships as a minor version.

## 9. Open questions

1. **`snake_case` with Office JS values, or camelCase members for literal parity.** Recommendation:
   `snake_case` members, Office JS string values (section 3); no camelCase aliases. The parity
   is in the vocabulary and the values, and a `test/office_js_subset.json` mapping keeps it
   honest.
2. **Code-point offsets** (section 3). Recommendation: yes; only `Range` carries offsets and it
   never crosses to Office JS.
3. **A `Word` shim.** Recommendation: no (3.9); `to_api_script` only.
4. **`Body.text` and text boxes.** Recommendation: Word's main story, no text boxes, the accepted
   view; `text_of` decides its own rule under CR-002 question 7 and the two are documented as
   different things (3.11).
5. **Reading `change_tracking_mode` unmarshals the settings part.** Recommendation: accept it
   (3.6); the settings part round-trips canonically and the alternative, a deferred write, only
   exists in TypeScript because of asynchrony.
6. **`insert_ooxml` before CR-002 Phase C.** Recommendation: accept bare fragments from Phase C
   here and add the `pkg:package` form when the flat OPC store exists, rather than blocking on
   it; agents that produce `pkg:package` strings are add-in agents first.
7. **Grapheme clusters without a dependency.** Recommendation: `unicodedata` plus the ZWJ and
   variation-selector list, `regex`'s `\X` when installed; a test on an emoji ZWJ sequence and
   on Devanagari.
8. **The Python MCP server.** Recommendation: a separate product repository once Phase D lands,
   with the tool names of docx4j-mcp and Office-Word-MCP-Server; not this CR.
