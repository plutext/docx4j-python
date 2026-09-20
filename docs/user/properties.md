# Document properties: adding or altering a property in `docProps/`

A `.docx` carries its metadata in up to three **properties parts** under `docProps/`, each a
small XML part with its own schema, its own relationship from the package root and its own
content type:

| Part | Root | What is in it | Typed as |
|---|---|---|---|
| `docProps/core.xml` | `cp:coreProperties` | Dublin Core: title, creator, subject, description, keywords, category, dates, revision, `lastModifiedBy`, `contentStatus` | `docx4j_py.docprops.core.CoreProperties`, part `DocPropsCorePart` |
| `docProps/app.xml` | `Properties` (extended) | what the application recorded: `Application`, `AppVersion`, `Company`, `Manager`, `Template`, the statistics (`Pages`, `Words`, `Characters`, `TotalTime`), `HyperlinkBase`, `TitlesOfParts` | `docx4j_py.docprops.extended.Properties`, part `DocPropsExtendedPart` |
| `docProps/custom.xml` | `Properties` (custom) | user-defined name/value pairs, each a `property` with a `name`, a `pid` and a typed value (`vt:lpwstr`, `vt:i4`, `vt:bool`, `vt:filetime`, ...) | `docx4j_py.docprops.custom.Properties`, part `DocPropsCustomPart` |

All three are **typed parts** (CR-002 section 9): `part.contents` is the unmarshalled root, a
slotted dataclass whose fields are the schema's elements, and a part whose `contents` you have
read is re-marshalled on save while every other part is written back byte for byte. The
package holds them as `pkg.doc_props_core_part`, `pkg.doc_props_extended_part` and
`pkg.doc_props_custom_part`, any of which is `None` when the document has no such part (the
corpus's `toc.docx` has core and app but no custom; `create_package()` writes core and app).

Everything below was run on 2026-09-21 against `samples/toc.docx` and the output is what is
shown.

## 1. The short way: the python-docx facade, for core properties

python-docx's `document.core_properties` is sixteen read/write scalars over `core.xml`, and the
facade carries them (CR-003 Phase J, `docx4j_py/docx/coreprops.py`). It hides the schema's
shapes --- `dc:title` is not a string but a `SimpleLiteral` element, `dcterms:modified` is a
`W3CDTF` carrying an `xsi:type` --- and creates the part when a loaded document has none:

```python
import datetime as dt
from docx4j_py.docx import Document

d = Document("toc.docx")
cp = d.core_properties
print(cp.title, "|", cp.author, "|", cp.modified)
# United Nations Declaration on the Elimination ... | Frohmader | 2016-11-22 01:25:00+00:00

cp.title = "Declaration (annotated)"
cp.author = "Plutext"
cp.keywords = "UN; rights"
cp.modified = dt.datetime(2026, 9, 21, 9, 0, tzinfo=dt.UTC)
d.save("out.docx")
```

The sixteen: `author`, `category`, `comments`, `content_status`, `created`, `identifier`,
`keywords`, `language`, `last_modified_by`, `last_printed`, `modified`, `revision`, `subject`,
`title`, `version`, and python-docx's names for them (`author` is `dc:creator`, `comments` is
`dc:description`). Dates are timezone-aware `datetime`s and are written in UTC as `W3CDTF`. Each
assignment records a `ChangeReport` like every facade call.

The facade covers **core only**, because python-docx does. For `app.xml` and `custom.xml`, and
for a core property python-docx does not name, use the typed part.

## 2. The typed part: read, alter, add, remove

### 2.1 Core (`docProps/core.xml`)

```python
from docx4j_py import load
import docx4j_py.docprops.core as core_ns

pkg = load("toc.docx")
props = pkg.doc_props_core_part.contents           # CoreProperties

props.keywords                                       # None: absent
props.revision                                       # '2'
props.title                                          # Title(any_element=AnyElement(text='United Nations ...'))
```

Two shapes of field, as the schema has them:

- **Plain strings**: `cp:keywords`, `cp:revision`, `cp:lastModifiedBy`, `cp:category`,
  `cp:contentStatus`, `cp:version`. Assign a `str`; assign `None` to remove the element.

  ```python
  props.keywords = "UN; rights"
  props.revision = str(int(props.revision or 0) + 1)
  ```

- **Dublin Core elements**: `dc:title`, `dc:creator`, `dc:subject`, `dc:description`,
  `dc:identifier`, `dc:language` are typed `SimpleLiteral` wrappers with a wildcard child, so a
  parsed value comes back as a wrapper around an `AnyElement` whose `.text` is the string, and a
  new one is the wrapper class around the string:

  ```python
  props.title = core_ns.Title(any_element="Declaration (typed)")
  props.title.parent = props                          # a single-valued field sets no parent itself
  ```

  `props.title.any_element.text` reads a parsed one; `props.title.any_element` is the `str` you
  assigned until the next parse. (The facade's `_text_of` reads both; copy it if you need to.)

- **Dates**: `dcterms:created` and `dcterms:modified` are `W3CDTF` with an `xsi:type`, which the
  serialiser writes for you:

  ```python
  import datetime as dt
  from docx4j_xsdata.models.datatype import XmlDateTime
  from docx4j_py.docprops.dcterms import W3CDTF

  stamp = XmlDateTime.from_datetime(dt.datetime(2026, 9, 21, 9, 0, tzinfo=dt.UTC))
  props.modified = W3CDTF(value=stamp)
  props.modified.parent = props
  # <dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-21T09:00:00Z</dcterms:modified>
  ```

  `cp:lastPrinted` is a plain `XmlDateTime`, no wrapper.

`pkg.save("out.docx")` then writes:

```xml
<cp:coreProperties ...><dcterms:created xsi:type="dcterms:W3CDTF">2016-11-22T01:25:00Z</dcterms:created>
<dc:creator>Frohmader</dc:creator><cp:keywords>UN; rights</cp:keywords><cp:lastModifiedBy>jharrop</cp:lastModifiedBy>
<dcterms:modified xsi:type="dcterms:W3CDTF">2016-11-22T01:25:00Z</dcterms:modified><cp:revision>3</cp:revision>
<dc:title>Declaration (typed)</dc:title></cp:coreProperties>
```

### 2.2 Extended (`docProps/app.xml`)

Every field is a plain scalar (`str`, `int`, `bool`) except the vectors (`TitlesOfParts`,
`HeadingPairs`, `HLinks`), so this is the easy part:

```python
app = pkg.doc_props_extended_part.contents           # extended.Properties
app.application, app.app_version, app.company         # 'Microsoft Office Word', '14.0000', 'Women With Disabilities Australia'
app.pages, app.words                                  # 4, 1548

app.company = "Plutext Pty Ltd"
app.manager = "J. Harrop"                             # was None; the element is added
app.template = None                                   # removes <Template>
```

The schema's content model is `xsd:all`, so the elements carry no order; a re-marshalled
`app.xml` keeps the source order and a new element goes where the class declares it. Word
recomputes the statistics (`Pages`, `Words`, ...) when it saves, so writing those is
pointless; `Company`, `Manager`, `HyperlinkBase` and `Template` are the ones a program sets.

### 2.3 Custom (`docProps/custom.xml`)

A list of `property` elements, each with three attributes and one typed child:

- `fmtid` --- always `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`, the user-defined property set;
- `pid` --- an integer, **unique within the part and at least 2** (1 is reserved for the code
  page); Word numbers them 2, 3, 4, ... and does not require them contiguous;
- `name` --- what Word's File → Info → Properties → Advanced → Custom shows;
- the value, one of the `vt:` variant types: `vt:lpwstr` (text), `vt:i4` (32-bit integer),
  `vt:bool`, `vt:filetime` (a date), `vt:r8` (a double).

```python
from docx4j_py.docprops import variant_types as vt
from docx4j_py.docprops.custom import Properties as CustomProperties, PropertiesProperty
from docx4j_py.openpackaging.parts.docprops import DocPropsCustomPart

FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

part = pkg.doc_props_custom_part
if part is None:                                     # toc.docx has none: create it (section 3)
    part = DocPropsCustomPart()
    part.set_contents(CustomProperties())
    pkg.add_target_part(part)

custom = part.contents
next_pid = max((p.pid or 1 for p in custom.property), default=1) + 1
custom.property.append(
    PropertiesProperty(fmtid=FMTID, pid=next_pid, name="ProjectCode", content=vt.Lpwstr(value="DOCX4J-7"))
)
custom.property.append(
    PropertiesProperty(fmtid=FMTID, pid=next_pid + 1, name="Reviewed", content=vt.Bool(value=True))
)
custom.property.append(
    PropertiesProperty(fmtid=FMTID, pid=next_pid + 2, name="PageCount", content=vt.I4(value=4))
)
pkg.save("out.docx")
```

`custom.property` is a `ChildList`, so `append` sets the parent. The part writes:

```xml
<prop:Properties xmlns:prop="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <prop:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ProjectCode"><vt:lpwstr>DOCX4J-7</vt:lpwstr></prop:property>
  <prop:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Reviewed"><vt:bool>true</vt:bool></prop:property>
  <prop:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="PageCount"><vt:i4>4</vt:i4></prop:property>
</prop:Properties>
```

(`prop:` is docx4j's prefix for the namespace; Word writes it as the default namespace. Same
document.) Read back, `[(p.name, type(p.content).__name__, p.content.value) for p in
custom.property]` is `[('ProjectCode', 'Lpwstr', 'DOCX4J-7'), ('Reviewed', 'Bool', True),
('PageCount', 'I4', 4)]`.

To **alter** one, find it by name and replace its `content` (the value's type may change: a
`Lpwstr` today, an `I4` tomorrow); to **remove** one, remove it from the list. The `pid`s of the
others need no renumbering.

```python
for p in custom.property:
    if p.name == "ProjectCode":
        p.content = vt.Lpwstr(value="DOCX4J-8")
        p.content.parent = p
custom.property[:] = [p for p in custom.property if p.name != "Reviewed"]
```

A `vt:filetime` takes an `XmlDateTime` like the core dates: `vt.Filetime(value=stamp)`.

## 3. When the part does not exist

A loaded document may lack any of the three (`toc.docx` lacks `custom.xml`; a document made by
a minimal generator may lack all three). Creating one is three things, and `add_target_part`
does two of them:

```python
from docx4j_py.docprops.extended import Properties as ExtendedProperties
from docx4j_py.openpackaging.parts.docprops import DocPropsExtendedPart

if pkg.doc_props_extended_part is None:
    part = DocPropsExtendedPart()                                       # its part name and content type
    part.set_contents(ExtendedProperties(application="docx4j-python", company="Plutext Pty Ltd"))
    rel = pkg.add_target_part(part)                                     # the relationship from the package
                                                                        # root and the content-type override
    # rel.id == 'rId5', rel.type_value ends in 'extended-properties', rel.target == 'docProps/app.xml'
```

The three part classes carry their part name and content type; `add_target_part(part)` relates
the part from the package's root relationships and adds the `[Content_Types].xml` override, as
docx4j's `addTargetPart` does. **Check for the existing part first**: `add_target_part` with a
part whose name is already taken replaces it, which is what you want only if you meant it.
`create_package()`'s documents already have core and app (the facade's `core_properties`
creates core when it is absent, so a facade user never sees this).

## 4. What is recorded, and what is not

The facade route records a `ChangeReport` per assignment (`docx.core_properties.title`, ...),
and `pkg.dry_run()` covers it. The typed route is a tree edit: nothing is recorded unless you
wrap it in the content API's `recording` yourself, and `dry_run` sees only what the deep copy
sees. Either way, reading `part.contents` marks the part for re-marshalling on save (CR-002's
rule: no dirty tracking, docx4j's own), so a program that only *reads* the properties of many
documents and saves them should read through `part.read_contents()` (a private parse, the part
stays byte for byte) or not save at all.

Nothing in the properties parts is validated beyond the schema: a `pid` collision, a `fmtid`
that is not the user-defined set, a `dc:title` with a child element are all written as given,
and Word repairs or ignores them. Keep to the shapes above.

## 5. What would be nicer

Java docx4j's `DocPropsCustomPart` has `getProperty(name)` / `setProperty(name, value)` and
`DocPropsCorePart` / `DocPropsExtendedPart` have typed getters, so a Java user never handles a
`pid` or a `vt:` class. The equivalent here would be a small `pkg.properties` view --- `core`,
`app` and `custom` as mappings of plain Python values with the creation of section 3 folded in,
recording a `ChangeReport` --- and it is the shape CR-006's MCP server would want for a
`properties` tool. It is not in any CR yet; if it is wanted, it is half a day over what section 2
shows and belongs beside the facade's `coreprops.py`, which already has the core translation.
