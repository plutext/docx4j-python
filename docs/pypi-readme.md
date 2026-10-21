# docx4j for Python

[docx4j](https://www.docx4java.org) for Python: the whole of ECMA-376 WordprocessingML as typed
objects, generated from docx4j's own schema tree; docx4j's Open Packaging engine written by hand
over them (`.docx` in, typed parts and relationships, `.docx` out, every part you did not touch
written back byte for byte); a content API in the vocabulary of Office JS's `Word.Body`,
`Paragraph`, `Range`, `Table`, comments, tracked changes and lists, built for agents; and a
python-docx facade so code already written runs. The classes carry docx4j's names (`P`, `R`,
`PPr`, `Tbl`, `Document`, `Styles`), so fifteen years of docx4j documentation and examples read
across; the design is the same as [docx4j-core-ts](https://github.com/plutext/docx4j-core-ts).

```
pip install docx4j
```

The distribution is `docx4j`; the import name is `docx4j_py`, as `python-docx` imports as
`docx`. Python 3.12 or later.

## Open, edit, save

```python
from docx4j_py import load

pkg = load("in.docx")                                # a WordprocessingMLPackage
body = pkg.body                                      # Word.Body over word/document.xml

print(body.text)                                     # a paragraph per line
for paragraph in body.paragraphs:                    # tables and content controls descended into
    print(paragraph.style, paragraph.text)

title = body.insert_paragraph("Report", location="Start", style="Heading 1")
title.alignment = "Centered"
hit = body.search("quick brown fox")[0]              # matches span runs freely
hit.font.italic = True                               # the runs are split at the boundaries
hit.insert_text("slow red fox")                      # Replace is a Range's default location
body.paragraphs[-1].insert_paragraph("The end.")     # After is a Paragraph's

pkg.save("out.docx")                                 # only the parts you touched are re-marshalled
```

A part that is never touched is written back byte for byte; reading a part's `contents` --- which
`pkg.body` does for `word/document.xml` --- marks it for re-marshalling, and a re-marshalled part
is canonically identical to its source. Nothing is dropped silently: what lenient parsing could
not place is reported per part (`part.skipped`), and `LoadOptions(strict=True)` raises instead.

From nothing:

```python
from docx4j_py import create_package

pkg = create_package()                       # docx4j's createPackage: A4, one section, the default styles
body = pkg.body
body.insert_paragraph("Created by docx4j", style="Heading 1")
paragraph = body.insert_paragraph("One paragraph, three runs")
paragraph.search("three runs")[0].font.italic = True
table = body.insert_table(3, 2, values=[["Name", "Value"], ["a", "1"]], style="TableGrid")
table.header_row_count = 1
table.add_rows(1, values=[["b", "2"]])
pkg.save("hello.docx")
```

## For agents

An agent cannot hold a Python object across tool calls, and a 200-page document does not fit in
its context window. So the API gives it **addresses** --- strings that survive an edit --- and
**budgets** on everything it reads. Four calls are the loop: read the outline, find the text,
edit by address, check the report.

```python
from docx4j_py import load

pkg = load("in.docx")

outline = pkg.outline()                       # small, structured, enough to choose an address from
outline.to_markdown()                         # the cheapest thing to show a model
outline.to_json()                             # under 64 KB for a 200-page document

hits = pkg.find("quick brown fox")            # matches with their addresses and 40 characters either side
hits[0].address, hits[0].snippet              # ('w14:5A2B1C3D', '… over the quick brown fox, which …')

paragraph = pkg.paragraph_at(hits[0].address) # 'w14:5A2B1C3D', 'body/3', or contains='Chapter 1'
paragraph.insert_paragraph("Added by an agent.", location="After")

pkg.last_change.to_json()                     # what that call did, for the tool result
pkg.save("out.docx")
```

`w14:5A2B1C3D` is the `w14:paraId` Word writes and survives every edit; `body/3` is the ordinal.
Every mutating call records a `ChangeReport`; `pkg.dry_run()` applies calls to a copy and throws
it away; `describe()` reads what the document offers (its styles, page, parts, authors) without
unmarshalling anything; every error carries a `code` and a `hint` naming what to do instead.

An agent should leave the trail Word already has: tracked changes for what it did, and a comment
for why, where the human who opens the document will see them.

```python
from docx4j_py import load
from docx4j_py.model.content import Author

pkg = load("in.docx")
pkg.author = Author("Claude", initials="C", email="claude@example.com")
pkg.change_tracking_mode = "TrackAll"          # every edit from here is a revision

count = pkg.body.replace_text("document", "report")     # a w:del and a w:ins per hit
pkg.find("report")[0].range(pkg.body).insert_comment(
    "Changed 'document' to 'report': the brief asks for a report."
)
[c.to_dict() for c in pkg.get_tracked_changes()]        # for the tool result
pkg.body.to_markdown(view="markup")                     # {--deleted--}{++inserted++}{>>a comment<<}
pkg.save("out.docx")
```

Markdown goes both ways: `pkg.to_markdown(addresses=True)` puts each block's address in an HTML
comment before it, and `body.insert_markdown(text)` writes CommonMark plus GFM tables through the
document's own styles, creating the numbering part when there is none.

The text an agent reads is one text model, the port of docx4j's (`docx4j_py.model.text`): text
boxes after their anchors, notes, hidden text left out, the views by resolution. `to_markdown` and
`outline()` read through it, and so can a tool:

```python
from docx4j_py import load
from docx4j_py.model.text import extract, markdown, plain_text

pkg = load("in.docx")
result = extract(pkg, view="markup")          # the block model, CriticMarkup in this view
plain_text(result)                            # plain text, as the specification writes it
markdown(result, pkg)                         # docx4j-markdown's Markdown of the same reading
[block.to_dict() for block in result.blocks]  # the blocks as JSON
```

## python-docx code runs

`docx4j_py.docx` is a structural subset of python-docx's public API --- python-docx's names and
semantics, this engine underneath. Its promise is a committed member list derived from
python-docx 1.2.0: 305 of 312 members, the rest refused with a reason each. python-docx is not a
dependency.

```python
from docx4j_py.docx import Document, Pt

doc = Document("in.docx")                             # or Document() for a new one
doc.add_heading("Report", level=1)                    # a built-in style it lacks is defined
p = doc.add_paragraph("Plain, ")
p.add_run("bold").bold = True
p.runs[1].font.size = Pt(14)
doc.add_table(rows=2, cols=2, style="Table Grid").cell(0, 0).text = "a"
for paragraph in doc.paragraphs:
    print(paragraph.style.name, paragraph.text)

print(doc.package.last_change.to_dict())              # what python-docx does not have: a ChangeReport,
doc.save("out.docx")                                  # parts untouched byte for byte, the typed tree
```

## The object model and the engine

Every element has a constructor in its namespace's `el` module, with `p`, `r`, `t` and `tbl` as
sugar; `wml(...)` parses a fragment with docx4j's prefixes declared; `to_xml` is the inverse.
The tree stays reachable from every view (`paragraph.element` is the typed `P`).

```python
from docx4j_py.wml import el, p, r, to_xml, wml

para = p("Hello ", r("World", bold=True), style="Heading1")
same = wml('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>')
print(to_xml(el.p(content=[el.r(content=[el.t("Hello World")])]), pretty=True))
```

Under `docx4j_py.openpackaging`: `OpcPackage`, `WordprocessingMLPackage`, `PartName`,
`ContentTypeManager`, `RelationshipsPart`, the typed WordprocessingML parts, `PartStore` and
`PartSink` (zip, directory, memory, flat OPC); `docx4j_py.model` has docx4j's `PropertyResolver`,
`StyleUtil`, numbering emulator and font selection, with zero differences on docx4j's 45 parity
goldens.

```python
from docx4j_py import load
from docx4j_py.openpackaging import LoadOptions, MemoryPartSink
from docx4j_py.wml import warm_up

warm_up()                                             # a server does this once, before the first request
with load("in.docx", options=LoadOptions(strict=True)) as pkg:
    settings = pkg.get_part("/word/settings.xml").contents   # typed, lazily, on first access
    pkg.save_to(MemoryPartSink())
```

## What is and is not in 0.1

**WordprocessingML only.** `.pptx` and `.xlsx` load and save through the generic OPC path (every
part byte for byte), but nothing in them is typed: PresentationML and SpreadsheetML are
[CR-001 Phase D](https://github.com/plutext/docx4j-python/blob/main/docs/change-requests/CR-001-object-model.md)
(the object model) and
[CR-002 Phase C](https://github.com/plutext/docx4j-python/blob/main/docs/change-requests/CR-002-engine.md)
(the engine), and their content APIs are CR-004 and CR-005. Everything is synchronous.

**The wheel is large for pure Python** --- about 1 MB compressed, 238 modules, 6 MB installed
--- because it is the whole of WordprocessingML and what it embeds (DrawingML, OMML, VML, the
`w14` to `w16` extensions) as typed classes, 148,000 lines of generated code. Importing
`docx4j_py` imports all of it, about 0.6 s once per process; a long-running process is the
intended host.

What `0.x` promises, surface by surface:

| Surface | Stability in `0.x` |
|---|---|
| `docx4j_py.load`, `create_package`, `WordprocessingMLPackage`, the parts, `PartStore` / `PartSink` | Stable: docx4j's names and behaviour; departures are recorded in CR-002 section 12.3 |
| The content API (`Body`, `Paragraph`, `Range`, `Table`, ..., the agent surface, markdown) | Stable in shape: Office JS's vocabulary is borrowed, not invented; a member may be added in a minor, none renamed |
| `docx4j_py.docx`, the python-docx facade | Stable: its promise is the committed member list, 305 of 312 |
| The generated model (`docx4j_py.wml` and the rest, `el`) | Regenerates with the schema; class and field names are docx4j's and do not move, but a schema refresh adds fields and classes in a minor |
| `docx4j_py.model` (resolver, `StyleUtil`, fonts, numbering) | Stable: docx4j's, zero differences on the goldens |

## Links

- Repository: <https://github.com/plutext/docx4j-python> --- the full README, with the whole
  content API, the audit trail, lists, custom XML templates and `to_api_script` shown
- The design, as change requests: [CR-001](https://github.com/plutext/docx4j-python/blob/main/docs/change-requests/CR-001-object-model.md)
  the object model, [CR-002](https://github.com/plutext/docx4j-python/blob/main/docs/change-requests/CR-002-engine.md)
  the engine, [CR-003](https://github.com/plutext/docx4j-python/blob/main/docs/change-requests/CR-003-content-api.md)
  the content API
- [docx4j](https://www.docx4java.org), the Java original, and
  [docx4j-core-ts](https://github.com/plutext/docx4j-core-ts), the TypeScript engine of the same design
- [Changelog](https://github.com/plutext/docx4j-python/blob/main/CHANGELOG.md)

Apache-2.0, as docx4j is; the generated model imports
[docx4j-xsdata](https://pypi.org/project/docx4j-xsdata/), a fork of xsdata (MIT). Every fence
above is executed by the test suite (`tests/test_readme.py`).
