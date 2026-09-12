# docx4j-python

docx4j for Python: the whole of ECMA-376 as typed objects, generated from docx4j's own schema
tree by a [fork of xsdata](https://github.com/tefra/xsdata), and docx4j's Open Packaging engine
written by hand over them: `.docx` in, typed parts and relationships, `.docx` out. The classes
carry docx4j's names (`P`, `R`, `Text`, `PPr`, `Tbl`, `Document`, `Styles`), so fifteen years of
docx4j documentation and examples read across. The design is the same as
[docx4j-core-ts](https://github.com/plutext/docx4j-core-ts), so the two read across too.

Status: the object model ([CR-001](docs/change-requests/CR-001-object-model.md) Phases A to C)
and the Open Packaging engine ([CR-002](docs/change-requests/CR-002-engine.md) Phase A) are
implemented. Over 16 real documents, every part not touched is written back byte for byte, all
141 typed WordprocessingML parts unmarshal and re-serialise canonically identical to the source,
and nothing is dropped. Saved output opens in Word. The resolution utilities (`PropertyResolver`,
list numbering, fonts: CR-002 Phase B) are next, then the content API (CR-003).

```python
from docx4j_py import load
from docx4j_py.wml import p, r, text_of

pkg = load("in.docx")                                    # a WordprocessingMLPackage; the kind is sniffed
main = pkg.main_document_part
document = main.contents                                 # unmarshalled on first access, as docx4j does
print(text_of(document))                                 # docx4j TextUtils: a paragraph per line

body = document.body
body.content.append(p("Hello World", style="Heading1")) # a paragraph, a run, the text, xml:space if needed
body.content.append(p("Plain, ", r("bold", bold=True), r(" and red.", color="FF0000")))

styles = pkg.style_definitions_part.contents             # Styles; header_parts(), footer_parts(), ...
pkg.save("out.docx")                                     # only the parts you unmarshalled are re-marshalled
```

A part that is never touched is written back byte for byte; reading `contents` marks a part for
re-marshalling. Everything is synchronous. `load` takes a path, bytes, a file object or a
`PartStore`; `save` takes a path, a file object or nothing (bytes), and saving over the file you
loaded is fine.

### From nothing to a `.docx`

```python
from docx4j_py import create_package
from docx4j_py.wml import p, r, tbl, br

pkg = create_package(page_size="A4")        # docx4j's createPackage: one section, its default styles
body = pkg.main_document_part.contents.body
body.content.append(p("Created by docx4j-python", style="Heading1"))
body.content.append(p("One paragraph, ", r("three runs", italic=True, size=14), r(", one break"), br()))
body.content.append(tbl([["Name", "Value"], ["a", "1"], ["b", "2"]], style="TableGrid"))
pkg.save("hello.docx")
```

`create_package` writes `docProps/app.xml` (`Application`, `AppVersion`) and `docProps/core.xml`
(`created`, `modified`); author and title are yours to set on those parts' `contents`.

### Adding an image

An image is a part related from the main document part; the picture in the body refers to it by
relationship id. The fragment is the shortest honest way to write the `w:drawing`, and it goes in
through `wml(...)`, which declares docx4j's prefix table for you:

```python
from docx4j_py.openpackaging import ImagePart, AddPartBehaviour
from docx4j_py.wml import wml

image = ImagePart("/word/media/image1.png")
image.set_bytes(png_bytes)
rel = main.add_target_part(image, AddPartBehaviour.RENAME_IF_NAME_EXISTS)   # rel.id is the r:embed

body.content.append(wml(f"""
  <w:p><w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="2857500" cy="1905000"/>
      <wp:docPr id="1" name="Picture 1"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="{rel.id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2857500" cy="1905000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>"""))
```

`scripts/acceptance.py` builds exactly this, and three other documents, for the manual Word
checklist in [`tests/README.md`](tests/README.md).

### The object model

Every element name has a constructor in the namespace's `el` module, generated from the class
metadata so it always returns the right class for that element; `p`, `r`, `t`, `tbl`, `br` and
`tab` are the hand-written sugar on top:

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
from docx4j_py.wml import text_of, walk, find, iter_nodes, deep_copy

text_of(document)                          # w:t, w:tab, w:br, w:sym; w:delText and fields excluded, as docx4j
find(document, P)                          # docx4j ClassFinder; isinstance, so subclasses match
find(document, (Tbl, Drawing))
walk(document, lambda node, parent, name: print(name))   # docx4j TraversalUtil; return False to stop descending
for node in iter_nodes(document): ...

copy = deep_copy(para)                     # the subtree, not the document; copy.parent is None until placed
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

## Running it

The bindings import `docx4j_xsdata`, the forked runtime, which `.venv-fork` has installed
editable from `~/git/docx4j-xsdata` (branch `docx4j`; `docs/fork/CHANGES.md` there lists every
difference from upstream xsdata).

```bash
.venv-fork/bin/python -m pytest                       # the suite
.venv-fork/bin/python -m pytest -m "not slow"         # without the corpus round trip and the timings
.venv-fork/bin/python -m pytest tests/openpackaging   # the engine

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
  wml/builders.py       hand written: p, r, t, tbl, br, tab and the run options
  dml/ math/ mce/ w14/ w15/ ...   the namespaces WordprocessingML embeds, the same shape
  relationships/ docprops/        the .rels classes and the three properties parts
  child.py              hand written: Child, ChildList, link_parents, iter_children, deep_copy
  namespaces.py         hand written: docx4j's prefix table; UNDERSTOOD is generated
  runtime.py            hand written: the shared XmlContext, warm_up()
  fragments.py          hand written: wml(...) and to_xml(...)
  traversal.py          hand written: walk, iter_nodes, find, text_of, with the mce mode
  openpackaging/        hand written, all of it: the engine (CR-002)
    part_name.py content_types.py stores.py load.py save.py mce.py resources.py api.py
    parts/              Part, BinaryPart, XmlPart, RelationshipsPart, the registry, the typed parts
    packages/           OpcPackage, WordprocessingMLPackage
  resources/            the parts warm_up parses; docx4j's default styles, numbering and fontTable
codegen/              the generator, the name tables and the el tables
schemas/              docx4j's xsd tree with marked patches (schemas/PATCHES.md)
scripts/              roundtrip.py, canon.py, checks.py, parents.py, bench.py, threads.py, acceptance.py
samples/              16 documents from docx4j (Apache-2.0): 13 .docx, a .dotm, a .pptx, an .xlsx
out/acceptance/       the four documents for the manual Word checklist
docs/change-requests/ the design: CR-001 the object model, CR-002 the engine
```

## Licence

Apache-2.0, as docx4j is. The forked generator and runtime are MIT, as xsdata is.
