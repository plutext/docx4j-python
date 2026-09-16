# docx4j-python

docx4j for Python: the whole of ECMA-376 as typed objects, generated from docx4j's own schema
tree by a [fork of xsdata](https://github.com/tefra/xsdata), and docx4j's Open Packaging engine
written by hand over them: `.docx` in, typed parts and relationships, `.docx` out. The classes
carry docx4j's names (`P`, `R`, `Text`, `PPr`, `Tbl`, `Document`, `Styles`), so fifteen years of
docx4j documentation and examples read across. The design is the same as
[docx4j-core-ts](https://github.com/plutext/docx4j-core-ts), so the two read across too.

Status: the object model ([CR-001](docs/change-requests/CR-001-object-model.md) Phases A to C),
the Open Packaging engine ([CR-002](docs/change-requests/CR-002-engine.md) Phase A) and the
content API ([CR-003](docs/change-requests/CR-003-content-api.md) Phase A, the tree-layer
builders; Phase B, `Body`, `Paragraph`, `Range` and `Font` in Office JS's vocabulary; and
Phase D, the agent surface — addresses, `outline()`, `find()`, `describe()`, `ChangeReport`,
`dry_run` and `DocumentSession`, all implemented 2026-09-16) are in. Over 16 real documents,
every part not touched is written back byte for byte, all 141 typed WordprocessingML parts
unmarshal and re-serialise canonically identical to the source, and nothing is dropped. Saved
output opens in Word. Markdown in and out (Phase K) and tables, pictures and content controls
(Phase C) are next.

```python
from docx4j_py import load

pkg = load("in.docx")                                # a WordprocessingMLPackage; the kind is sniffed
body = pkg.body                                      # Word.Body over word/document.xml

print(body.text)                                     # a paragraph per line, the accepted view
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

A part that is never touched is written back byte for byte; reading `contents` --- which `pkg.body`
does for `word/document.xml` --- marks a part for re-marshalling. Everything is synchronous. `load`
takes a path, bytes, a file object or a `PartStore`; `save` takes a path, a file object or nothing
(bytes), and saving over the file you loaded is fine.

### From nothing to a `.docx`

```python
from docx4j_py import create_package

pkg = create_package()                                       # docx4j's createPackage: A4, one section
pkg.body.insert_paragraph("Hello World")
pkg.save("hello.docx")
```

Longer, with the verbs of Office JS's `Word.Body` (`docx4j_py.model.content`, CR-003):

```python
from docx4j_py import create_package

pkg = create_package(page_size="A4")        # docx4j's createPackage: one section, its default styles
body = pkg.body

body.insert_paragraph("Created by docx4j-python", style="Heading 1")
paragraph = body.insert_paragraph("One paragraph, three runs")
paragraph.search("three runs")[0].font.italic = True
body.insert_break("Page")
body.insert_xml("<w:tbl><w:tblPr/><w:tblGrid/>"
                "<w:tr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc></w:tr></w:tbl>")
pkg.save("hello.docx")
```

`create_package` writes `docProps/app.xml` (`Application`, `AppVersion`) and `docProps/core.xml`
(`created`, `modified`); author and title are yours to set on those parts' `contents`.

**The tree stays reachable.** Nothing is wrapped: `paragraph.element` is the `P`, `body.content`
is the live `ChildList`, and the low-level route of CR-001 works exactly as it did, on the same
objects the views hand out.

```python
from docx4j_py.wml import p, r, tbl, br, text_of

document = pkg.main_document_part.contents               # unmarshalled on first access, as docx4j does
body = document.body                                     # the typed w:body; pkg.body is the view over it
body.content.append(p("Plain, ", r("bold", bold=True), r(" and red.", color="FF0000")))
body.content.append(tbl([["Name", "Value"], ["a", "1"]], style="TableGrid"))
text_of(document)                                        # docx4j TextUtils, over any subtree
```

### Adding an image

An image is a part related from the main document part; the picture in the body refers to it by
relationship id. `inline_picture` writes the `w:drawing` exactly as docx4j's
`BinaryPartAbstractImage.createImageInline` does, and `image_size` reads the size and the density
out of the image's own header (PNG, JPEG, GIF and BMP, no Pillow), so the picture comes out the
size it really is:

```python
from docx4j_py.openpackaging import ImagePart, AddPartBehaviour
from docx4j_py.wml import el, r, emu_for, image_size, inline_picture

main = pkg.main_document_part
image = ImagePart("/word/media/image1.png")
image.set_bytes(png_bytes)
rel = main.add_target_part(image, AddPartBehaviour.RENAME_IF_NAME_EXISTS)   # rel.id is the r:embed

size = emu_for(image_size(png_bytes), max_width_emu=5731510)   # scaled to the text column
body.content.append(el.p(content=[
    r(inline_picture(rel.id, cx=size.cx, cy=size.cy, id=1, name="image1.png", descr="A pangolin"))
]))
```

`wml("<w:p><w:r><w:drawing>…")` still takes the whole thing as a fragment when you want to write
the XML yourself; it declares docx4j's prefix table for you.

`scripts/acceptance.py` builds exactly this, and three other documents, for the manual Word
checklist in [`tests/README.md`](tests/README.md).

### The object model

Every element name has a constructor in the namespace's `el` module, generated from the class
metadata so it always returns the right class for that element; `p`, `r`, `t`, `tbl`, `tr`, `tc`,
`br` and `tab` are the hand-written sugar on top:

```python
from docx4j_py.wml import el, p, r, t, P, R, Text, Tbl, Drawing

el.p(content=[el.r(content=[el.t("Hello World")])])   # docx4j's ObjectFactory, one function per element
el.t("  two spaces  ")                                # xml:space="preserve" set for you
el.pPr(); el.sdt_run(); el.del_(); el.t_math()        # 604 names in wml: scope-qualified (sdt_run) and
                                                      # keyword-renamed (del_) where they have to be
r("bold", bold=True, size=12, color="FF0000", underline="Single", font="Arial")   # Office JS names, as docx4j-core-ts
```

Fragments and XML: `wml(...)` parses one element written as it would appear inside
`document.xml`, with the standard prefixes declared for you (docx4j's
`XmlUtils.unmarshalString` with `W_NAMESPACE_DECLARATION`); `to_xml` is the inverse. Parsing is
strict by default: anything the model would drop raises, `lenient=True` reports it instead.

```python
from docx4j_py.wml import wml, to_xml

para = wml('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
           '<w:r><w:rPr><w14:glow w14:rad="101600"><w14:srgbClr w14:val="FF0000"/></w14:glow></w:rPr>'
           '<w:t>Hello</w:t></w:r></w:p>')
paras = wml.all("<w:p>...</w:p><w:p>...</w:p>")     # several siblings
to_xml(para)                                        # declarations trimmed to what the fragment uses
```

Marshalling goes the other way for any object in the tree, a `P` say, with `to_xml` (docx4j
`XmlUtils.marshaltoString`); the part as a whole is `part.xml`, the bytes `save` would write:

```python
from docx4j_py.wml import find, P, p, r

second = find(document, P)[1]                       # a paragraph from the loaded document
print(to_xml(second, pretty=True))
```

```xml
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidR="00D15781" w:rsidRDefault="00945132">
  <w:r>
    <w:t>This is a document exhibiting basic docx features.</w:t>
  </w:r>
  <w:r w:rsidR="00665DAE">
    <w:t xml:space="preserve">  </w:t>
  </w:r>
</w:p>
```

```python
print(to_xml(p("Hello ", r("World", bold=True), style="Heading1"), pretty=True))
```

```xml
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:r>
    <w:t xml:space="preserve">Hello </w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
    <w:t>World</w:t>
  </w:r>
</w:p>
```

```python
main.xml                                            # bytes: the whole word/document.xml, as save() writes it
repr(second)                                        # the dataclass: P(p_pr=None, content=[R(r_pr=None, content=[RT(value='This is ...
```

Traversal is metadata-driven, so it needs no per-class code and covers every namespace:

```python
from docx4j_py.wml import (text_of, walk, walk_all, find, iter_nodes, run_items_of,
                           deep_copy, deep_copy_as, PPrBase)

text_of(document)                          # w:t, w:tab, w:br, w:sym; w:delText and fields excluded, as docx4j
find(document, P)                          # docx4j ClassFinder; isinstance, so subclasses match
find(document, (Tbl, Drawing))
walk(document, lambda node, parent, name: print(name))   # docx4j TraversalUtil; return False to stop descending
for node in iter_nodes(document): ...
walk_all(document, on_node, on_wildcard)   # and into the xs:any content, where a stray r:embed hides
run_items_of(para)                         # the run list of a w:p, w:hyperlink, w:ins, w:sdt, ...

copy = deep_copy(para)                     # the subtree, not the document; copy.parent is None until placed
base = deep_copy_as(para.p_pr, PPrBase)    # a copy re-typed as a base class, so no xsi:type is written
```

Content controls have their own builders, over a `w:sdtPr` the model keeps as one choice list, as
docx4j does:

```python
from docx4j_py.wml import p, sdt, sdt_property, sdt_kind_of, next_sdt_id

control = sdt([p("Acme Ltd")], kind="RichText", tag="customer", id=next_sdt_id(document))
sdt_kind_of(control.sdt_pr)                # Office JS Word.ContentControlType: 'RichText'
sdt_property(control, "dataBinding")       # looks in w: and w15:, because Word writes w15:dataBinding
```

Every object is a `Child`: it knows its parent (`node.parent`), and appending to a content list
sets the child's parent, as docx4j's `ArrayListDocx4j` does. Objects are slotted dataclasses;
equality and `repr` ignore the parent, and so does serialisation.

### The engine

```python
from docx4j_py.openpackaging import (
    OpcPackage, WordprocessingMLPackage, LoadOptions, PartName,
    ImagePart, DefaultXmlPart, AddPartBehaviour, ContentTypes, Namespaces,
    ZipPartStore, DirectoryPartStore, MemoryPartStore, MemoryPartSink,
)

pkg = OpcPackage.load("in.pptx")                  # anything OPC loads; .docx comes back as WordprocessingMLPackage
pkg.parts["/ppt/presentation.xml"]                # case-insensitive, as OPC says; untyped parts are DefaultXmlPart (lxml)
pkg.content_type_manager.get_content_type("/ppt/slides/slide1.xml")
pkg.relationships_part.get_relationships_by_type(Namespaces.DOCUMENT)   # docx4j's constant names

with load("in.docx") as pkg:                      # a context manager: the zip stays open for lazy loads
    part = pkg.get_part("/word/settings.xml")
    part.is_unmarshalled                          # False: load cost the relationships and content types, nothing else
    settings = part.contents                      # now True, and this part will be re-marshalled on save
    part.skipped                                  # what lenient parsing dropped from it: empty is the norm
    pkg.skipped                                   # the same, over every unmarshalled part
    pkg.unmarshal_all()                           # for callers who prefer no laziness
    pkg.save()                                    # bytes
    pkg.save_to(MemoryPartSink())                 # or any sink: a directory, a zip, memory

load("in.docx", options=LoadOptions(strict=True))              # raise on the first dropped node instead
load("in.docx", options=LoadOptions(mce_preprocess=True))     # resolve mc:AlternateContent on read, as docx4j does
```

Markup compatibility: by default both branches of every `mc:AlternateContent` are kept, so an
unmarshalled part loses nothing on save. The branches are wildcard content, which means a
`w:drawing` inside one is not a typed `Drawing` until you opt into `mce_preprocess=True`, which
replaces each `mc:AlternateContent` with the branch Word would take, as docx4j's preprocessor
does. The traversal functions take `mce="resolve"` (default, descend into the understood
`mc:Choice`, else the `mc:Fallback`), `"all"` or `"none"`.

Under `docx4j_py.openpackaging`: `PartName` (OPC part names, case-insensitive equality,
relationship target resolution), `ContentTypeManager` and `ContentTypes`, `RelationshipsPart`
with `add_part` / `add_relationship` / `remove_part` and docx4j's `AddPartBehaviour`,
`Namespaces` for the relationship types, `PartRegistry` (content type to part class, extensible),
the typed WordprocessingML parts (`MainDocumentPart`, `StyleDefinitionsPart`,
`NumberingDefinitionsPart`, `FontTablePart`, `DocumentSettingsPart`, `HeaderPart`, `FooterPart`,
`FootnotesPart`, `EndnotesPart`, `CommentsPart` and the w15 and w16 comment parts,
`GlossaryDocumentPart`, `ThemePart`, the three `DocProps*Part`s, chart and diagram parts), and
`CustomXmlDataStoragePart`, `VMLPart` and `DefaultXmlPart` as lxml trees.

## For agents

An agent cannot hold a Python object across tool calls, and a 200-page document does not fit in
its context window. So the API gives it **addresses** — strings that survive an edit — and
**budgets** on everything it reads. Four calls are the whole loop: read the outline, find the
text, edit by address, check the report.

```python
from docx4j_py import load

pkg = load("in.docx")

outline = pkg.outline()                       # small, structured, enough to choose an address from
outline.to_markdown()                         # the cheapest thing to show a model
outline.stats.to_dict()                       # {'paragraphs': 412, 'tables': 9, 'words': 11_204, ...}
outline.to_json()                             # under 64 KB for a 200-page document

hits = pkg.find("quick brown fox")            # matches with their addresses and 40 characters either side
hits[0].address, hits[0].snippet              # ('w14:5A2B1C3D', '… over the quick brown fox, which …')

paragraph = pkg.paragraph_at(hits[0].address) # 'w14:5A2B1C3D', 'body/3', or contains='Chapter 1'
paragraph.insert_paragraph("Added by an agent.", location="After")

pkg.last_change.to_json()                     # what that call did, for the tool result
pkg.save("out.docx")
```

**Three address forms**, accepted wherever a block-level target is and reported everywhere a view
is. `w14:5A2B1C3D` is the `w14:paraId` Word writes, and it survives every edit including an insert
in front of it; `body/3` and `body/4/0/1/0` (a paragraph in a cell) are the ordinal, stable until
something is inserted or deleted before them; `contains="Chapter 1"` is the first text match.
`paragraph.address` is the paraId when there is one and the ordinal otherwise, `paragraph.ordinal`
is always the ordinal. A document Word has not stamped can be stamped here:

```python
pkg.id_seed = 1234                            # same seed, same ids, same bytes
pkg.body.ensure_para_ids()                    # every paragraph gets a stable handle
```

(That one re-marshals `word/document.xml`, because it writes an attribute on every paragraph.
A document created with `create_package()` stamps every paragraph it is given, so nothing is
needed there.)

**Every mutation says what it did.** A `ChangeReport` is recorded for every content-API call —
always, because a server needs one on every call anyway — and costs about 5 µs on a 40 µs
insert. `pkg.last_change` is the last; `pkg.changes` accumulates until `pkg.changes.clear()`.

```python
pkg.last_change.to_dict()
# {'operation': 'insert_paragraph',
#  'addresses': ['w14:1C24BE5D'],
#  'moved': [['body/4', 'body/5'], ['body/5', 'body/6']],   # the ordinals that shifted
#  'created_para_ids': ['1C24BE5D'],
#  'text_after': 'Added by an agent.',
#  'parts_touched': ['/word/document.xml'],
#  'at': '2026-09-16T12:46:08.749252+00:00'}
```

**Try it first.** A `dry_run` applies the calls to a deep copy of the parts it touches and throws
the copy away; the real package is byte for byte what it was.

```python
with pkg.dry_run() as trial:
    count = trial.body.replace_text("colour", "color")
    preview = trial.last_change.to_dict()      # what it would do

pkg.body.replace_text("colour", "color")       # commit, having seen the number
```

**What can I use here?** `describe()` is docx4j-mcp's `describe_template` widened to any document:
the styles it defines and which it uses, the page, the parts, the comment and revision authors,
whether tracking is on. It reads every part as bytes with lxml, so it unmarshals nothing and an
untouched part stays byte for byte.

```python
description = pkg.describe()
description.style_names(kind="paragraph", in_use=True)   # ['Heading 1', 'Heading 2', 'Caption', ...]
description.page.width_pt, description.page.orientation  # (595.3, 'portrait')
description.authors                                      # {'comments': [...], 'revisions': [...]}
```

**Budgets everywhere**, and when a result is cut the dataclass says so:
`outline(depth=, max_chars=80, headings_only=True, limit=300)`, `find(context=40, limit=20)`,
`get_text(max_chars=)`, and `text_budget(max_chars)` for the caller who needs the flag rather
than the string.

**Errors say what to do instead.** Every one carries a stable `code`, the message, and a `hint`
an agent can act on: an address that is gone names the nearest surviving one, an element a
container cannot hold names what it takes, a style that does not exist lists the five closest
names, a span that crosses a hyperlink says where to split.

```python
pkg.element_at("body/99")
# AddressError: nothing at 'body/99'; the nearest surviving address is 'body/6'
#               (use 'body/6', or call outline() to list current addresses)
```

### Serving many documents

Import costs about 0.7 s and builds no class metadata; the first parse of each class does. In a
server, import once and warm the context before the first request; parse and serialise each
part with its own `ParserConfig` (the engine does), and share the rest across threads:

```python
from docx4j_py.wml import warm_up
warm_up()            # 44 ms: parses an embedded document and styles part, so the first request does not
```

`scripts/threads.py` is the check: 8 threads over whole packages with a shared `XmlContext`,
output identical to the sequential run.

A server then holds packages open across tool calls, so that "open, edit, edit, save" is four
cheap calls and not four loads. `DocumentSession` is that registry, in the library because the
thread rule and the idle-close logic are the library's knowledge:

```python
from docx4j_py import DocumentSession

session = DocumentSession(idle_timeout=900, max_open=32)

handle = session.open("report.docx")          # a short opaque string, not a path
with session.use(handle) as pkg:              # the handle's lock is held: one call at a time
    pkg.body.paragraph_at(contains="Chapter 1").insert_paragraph("New", location="After")
    report = pkg.last_change
session.save(handle, "edited.docx")           # or save(handle) for the bytes, overwrite=True for the source
session.sweep()                               # closes what has been idle; no background thread
session.close(handle)
```

One lock per handle, so two tool calls on one document serialise and two on different documents
do not. There is no background thread: `sweep()` is the server's to call.

## Running it

The bindings import `docx4j_xsdata`, the forked runtime, which `.venv-fork` has installed
editable from `~/git/docx4j-xsdata` (branch `docx4j`; `docs/fork/CHANGES.md` there lists every
difference from upstream xsdata).

```bash
.venv-fork/bin/python -m pytest                       # the suite
.venv-fork/bin/python -m pytest -m "not slow"         # without the corpus round trip and the timings
.venv-fork/bin/python -m pytest tests/openpackaging   # the engine
.venv-fork/bin/python -m pytest tests/agent           # the agent surface: outline, find, addresses

codegen/generate.sh                                   # regenerate docx4j_py/ from schemas/
codegen/generate.sh --check                           # regenerate twice and prove it is reproducible

.venv-fork/bin/python scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata
.venv-fork/bin/python scripts/threads.py              # the thread-safety check
.venv-fork/bin/python scripts/acceptance.py           # the four documents for the Word checklist
```

[`codegen/README.md`](codegen/README.md) explains how the bindings are generated, the docx4j name
tables in `codegen/names/` and the `el` tables in `codegen/el_tables/`. `schemas/PATCHES.md`
lists every change to docx4j's schema copy.

## Layout

```
docx4j_py/            one package per XML namespace, generated; codegen/generate.sh owns it
  wml/__init__.py       the classes: P, R, Text, PPr, Tbl, Document, Styles, ...
  wml/el.py             the object factory: el.p, el.r, el.t, el.sdt_run, ...
  wml/builders.py       hand written: p, r, t, tbl, tr, tc, br, tab, the run options, rpr_to_elements
  wml/pictures.py       hand written: inline_picture, image_size, emu_for (CR-003 Phase A)
  wml/sdt.py            hand written: sdt, sdt_pr, sdt_property, sdt_kind_of, next_sdt_id
  dml/ math/ mce/ w14/ w15/ ...   the namespaces WordprocessingML embeds, the same shape
  relationships/ docprops/        the .rels classes and the three properties parts
  child.py              hand written: Child, ChildList, link_parents, iter_children, deep_copy(_as)
  namespaces.py         hand written: docx4j's prefix table; UNDERSTOOD is generated
  runtime.py            hand written: the shared XmlContext, warm_up()
  fragments.py          hand written: wml(...) and to_xml(...)
  traversal.py          hand written: walk, walk_all, iter_nodes, find, text_of, run_items_of
  openpackaging/        hand written, all of it: the engine (CR-002)
    part_name.py content_types.py stores.py load.py save.py mce.py resources.py api.py
    parts/              Part, BinaryPart, XmlPart, RelationshipsPart, the registry, the typed parts
    packages/           OpcPackage, WordprocessingMLPackage
  model/content/        hand written, all of it: the content API (CR-003 Phases B and D)
    body.py paragraph.py range.py font.py   the views, in Office JS's vocabulary
    text_model.py       the paragraph's text as segments, and grapheme-safe splitting
    styles.py enums.py errors.py            BUILT_IN_STYLES, the Literals and StrEnums, ContentError
    addresses.py        the three address forms, and the paraId generator's users
    reports.py          Outline, SearchHit, ChangeReport, TextExcerpt, and the recorder
    describe.py trial.py                    describe() (lxml, unmarshals nothing) and dry_run()
  model/sessions.py     hand written: DocumentSession, the registry a server keeps documents in
  resources/            the parts warm_up parses; docx4j's default styles, numbering, fontTable
                        and KnownStyles.xml
codegen/              the generator, the name tables and the el tables
schemas/              docx4j's xsd tree with marked patches (schemas/PATCHES.md)
scripts/              roundtrip.py, canon.py, checks.py, parents.py, bench.py, threads.py,
                      acceptance.py, office_js_subset.py
samples/              16 documents from docx4j (Apache-2.0): 13 .docx, a .dotm, a .pptx, an .xlsx
out/acceptance/       the four documents for the manual Word checklist
docs/change-requests/ the design: CR-001 the object model, CR-002 the engine,
                      CR-003 the content API (Phases A, B and D implemented 2026-09-16)
```

## Licence

Apache-2.0, as docx4j is (see [LICENSE](LICENSE)). The forked generator and runtime are MIT, as xsdata is.
The schemas under `schemas/` are derived from ECMA-376 and keep the terms stated in their own headers.
