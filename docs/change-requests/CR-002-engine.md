# CR-002: The engine: container interface, Open Packaging layer, typed parts, resolution utilities

**Status:** Proposed 2026-09-12; open questions decided 2026-09-12 (section 11); **Phase A implemented
2026-09-12** (section 12); Word acceptance passed 2026-09-12 (12.8). Phases B and C proposed.
**Depends on:** CR-001 Phases A to C (implemented 2026-09-12): the generated model under `docx4j_py`,
`Child` and `ChildList`, `link_parents`, `deep_copy`, `el`, `wml(...)`, `to_xml`, `text_of`,
`walk`, `find`, `warm_up()`, the prefix table in `docx4j_py.namespaces`, the fork runtime's
skipped-content report and numeric booleans. Two small additions to the model are listed in
section 9 and are pulled forward from CR-001 Phase D.
**Counterpart:** docx4j-core (`org.docx4j.openpackaging.*`, `io3.Load3`, `io3.Save`,
`io3.stores.PartStore`, `org.docx4j.model.PropertyResolver`, `org.docx4j.model.listnumbering.*`,
`org.docx4j.fonts.*`), docx4j VERSION_17_1_1; docx4j-core-ts CR-001 (Phase A implemented
2026-09-10), whose structure this CR follows so the two engines stay recognisably one design.

## 1. Summary

CR-001 gives WordprocessingML typed objects; this CR gives them a document to live in. The engine
reads and writes Office Open XML packages from zip files, bytes, file objects and directories,
exposes their parts as typed objects related by relationships exactly as docx4j does, and ports
the three utilities every docx4j consumer ends up needing: effective paragraph and run properties
(`PropertyResolver`), list numbering (`Emulator`) and font selection (`RunFontSelector` with a
pluggable `Mapper`).

Design rules, from docx4j and from docx4j-core-ts CR-001:

- **A part that is never touched is written back byte for byte.** Only parts whose content was
  unmarshalled are re-marshalled on save. This is what makes round trips safe with content the
  model does not know, and what keeps saving fast.
- **A part that is unmarshalled loses nothing silently.** The fork's skipped-content report is
  attached to every unmarshalled part and logged; a strict mode raises. This is the rule CR-001
  section 7 set, applied per part.
- **Names follow docx4j** (`org.docx4j.openpackaging`) in `snake_case`; departures are recorded
  in the implementation notes, as the TypeScript engine does in its section 12.
- **Everything is synchronous.** docx4j-core-ts needed an asynchronous `getContents()` because its
  runtime imports modules dynamically; Python has no such constraint, so `part.contents` is a
  plain property that unmarshals on first access, which is exactly docx4j's `getContents()`.

Scope is WordprocessingML end to end. PresentationML and SpreadsheetML get their package classes
and main parts once CR-001 Phase D has generated their schemas; until then they load as generic
parts through the same registry and are not typed.

## 2. Non-goals

Encryption (password-protected packages), digital signatures, VBA, the XSLT, FO and PDF export
line, mail merge, OpenDoPE binding, markdown, the diffx comparison, and the content API (CR-003)
are separate CRs or products, as in Java. Schema validation is not part of the engine. Flat OPC
(`pkg:package`) is Phase C, not Phase A: it matters for Office JS interchange, which is a
TypeScript concern first, and it needs the `xmlPackage.xsd` types from CR-001 Phase D.

## 3. Runtime and dependencies

- **Python:** 3.10 and later, as the fork runtime requires. Tested on 3.12 to 3.14.
- **Zip:** the standard library's `zipfile`. It reads the central directory once, inflates an
  entry on demand, handles ZIP64, and writes with per-entry compression choice. There is no
  reason for a third-party zip library; the choice is confined to one `PartStore` implementation
  anyway.
- **XML:** the fork runtime (`docx4j_xsdata`) with its lxml handler for typed parts; `lxml.etree`
  directly for what the model does not type (`DefaultXmlPart`, custom XML data storage parts,
  `[Content_Types].xml`, the optional MCE preprocessor). lxml is already a hard dependency of the
  runtime path we use, so this adds nothing.
- **Content types:** `[Content_Types].xml` is not in the generated set (docx4j generates it from
  `xsd/contentTypes/opc-contentTypes.xsd`; docx4j-core-ts hand-wrote it). Two element types,
  `Default` and `Override`; the engine parses and writes it by hand with lxml (section 5.3).
- **Threads:** one `ParserConfig` per thread, per CR-001 section 14; the engine never shares a
  parser config between parts being unmarshalled concurrently. The `XmlContext` is shared and
  built once, warmed by `warm_up()`.

## 4. The container interface

docx4j's `io3.stores.PartStore` separates *where bytes come from and go to* from *what the
package is*. The same split is the point here, because there are three containers from the
start: a zip, a directory, and memory, and flat OPC later.

```python
class PartStore(Protocol):
    """Where a package's bytes come from. docx4j: io3.stores.PartStore, the load half."""
    def part_names(self) -> Iterable[str]: ...          # as stored, no leading '/', container order
    def has(self, part_name: str) -> bool: ...
    def load(self, part_name: str) -> bytes: ...         # may inflate lazily on first access
    def size(self, part_name: str) -> int | None: ...    # if known without loading

class PartSink(Protocol):
    """Where a package's bytes go. docx4j: PartStore, the save half."""
    def put(self, part_name: str, data: bytes, *, content_type: str, compress: bool = True) -> None: ...
    def finish(self) -> Any: ...                         # completes the container; return type per sink
```

| Class | Source | Notes |
|---|---|---|
| `ZipPartStore` | path, bytes, or a binary file object | `zipfile.ZipFile`; central directory read once; an entry is inflated on first `load`; the file object stays open until the package is closed (`OpcPackage` is a context manager) |
| `DirectoryPartStore` | an unzipped directory | docx4j `UnzippedPartStore`; for diffing and for tests |
| `MemoryPartStore` | `dict[str, bytes]` | new packages, tests, `MemoryPartSink` results |
| `FlatOpcPartStore` (Phase C) | a `pkg:package` string or element | XML parts re-serialised to bytes, binary parts base64-decoded; synthesises `[Content_Types].xml` from the per-part `pkg:contentType` |
| `ZipPartSink` | path, or a binary file object, or bytes via `BytesIO` | writes `[Content_Types].xml` first, as Word does; deflates everything except already-compressed media (png, jpeg, gif, emf/wmf left deflated) which are stored |
| `DirectoryPartSink` | a directory | for diffing saved output |
| `MemoryPartSink` | | returns a `MemoryPartStore` |
| `FlatOpcPartSink` (Phase C) | | the `pkg:package` string |

The package keeps a reference to its **source store**. On save, a part that was never
unmarshalled, and whose bytes were never replaced, is copied from the source store to the sink
unchanged (docx4j `Save.saveRawXmlPart`, `OpcPackage.getSourcePartStore()`). Saving to the same
path the package was loaded from is supported by writing to a temporary file beside it and
renaming, since the source zip is still open for lazy loads.

## 5. The Open Packaging layer

Names follow `org.docx4j.openpackaging`, in `snake_case`.

### 5.1 `PartName`

Value object for an OPC part name: leading `/`, no trailing `/`, segments percent-encoded per OPC
Annex A. `PartName.resolve(source, target)` resolves a relationship target relative to the source
part's directory, with `../` handling and `TargetMode="External"` left untouched;
`PartName.rels_for(name)` gives `/word/_rels/document.xml.rels`; `extension`, `directory`,
`name`. Equality and hashing are case-insensitive, since Word writes mixed case; the stored
spelling is kept for writing.

### 5.2 `RelationshipsPart`

An `XmlPart[Relationships]` over the generated `docx4j_py.relationships` classes (section 9).
Package relationships live at `/_rels/.rels`; every part with relationships owns one, reached
through `part.relationships_part`. API as docx4j: `get_relationship_by_type`,
`get_relationships_by_type`, `get_relationship_by_id`, `get_part(rel)`, `get_rel(part_name)`,
`is_a_target(part_name)`, `add_part(part, mode)` (allocates the next free `rIdN`; `mode` is
docx4j's `AddPartBehaviour`: `OVERWRITE_IF_NAME_EXISTS`, `RENAME_IF_NAME_EXISTS`,
`REUSE_EXISTING`), `add_relationship`, `remove_part`, `remove_relationship`. `Namespaces` holds
the relationship type constants (`OFFICE_DOCUMENT`, `STYLES`, `NUMBERING`, `FONT_TABLE`, `HEADER`,
`FOOTER`, `IMAGE`, `HYPERLINK`, `CUSTOM_XML`, `COMMENTS`, `COMMENTS_EXTENDED`, ...), as docx4j's
`org.docx4j.openpackaging.parts.relationships.Namespaces`. The relationships namespace is the
default namespace on the root when written, as Word writes it.

### 5.3 `ContentTypeManager`

Parses and writes `[Content_Types].xml`: `Default` (extension to content type) and `Override`
(part name to content type). `get_content_type(part_name)`, `add_default_content_type`,
`add_override_content_type`, `remove_content_type`. `ContentTypes` holds the constants
(`WORDPROCESSINGML_DOCUMENT`, `WORDPROCESSINGML_DOCUMENT_MACROENABLED`, `WORDPROCESSINGML_TEMPLATE`,
`WORDPROCESSINGML_STYLES`, `IMAGE_PNG`, ...), as docx4j's `org.docx4j.openpackaging.contenttype.ContentTypes`.

### 5.4 Parts

```
Part                          part_name, content_type, relationship_type, package,
                              relationships_part, owning_relationship_part, source_relationships
├── BinaryPart                bytes (property, lazy), set_bytes()                          docx4j BinaryPart
│   ├── ImagePart             one class; the image kind is the content type               docx4j BinaryPartAbstractImage + Image*Part
│   ├── EmbeddedPackagePart, OleObjectBinaryPart, ObfuscatedFontPart, ...
├── XmlPart[T]                contents (property: unmarshals on first access), set_contents(T),
│   │                         is_unmarshalled, xml (bytes), root_name, skipped                docx4j JaxbXmlPart<E>
│   ├── RelationshipsPart[Relationships]
│   ├── MainDocumentPart[Document], StyleDefinitionsPart[Styles], NumberingDefinitionsPart[Numbering],
│   │   FontTablePart[Fonts], DocumentSettingsPart[Settings], WebSettingsPart[WebSettings],
│   │   HeaderPart[Hdr], FooterPart[Ftr], FootnotesPart[Footnotes], EndnotesPart[Endnotes],
│   │   CommentsPart[Comments], CommentsExtendedPart, CommentsIdsPart, CommentsExtensiblePart, PeoplePart,
│   │   GlossaryDocumentPart[GlossaryDocument], ThemePart[Theme], BibliographyPart,
│   │   DocPropsCorePart[CoreProperties], DocPropsExtendedPart[Properties], DocPropsCustomPart[Properties],
│   │   CustomXmlDataStoragePropertiesPart[DatastoreItem], AlternativeFormatInputPart (altChunk),
│   │   chart, chart style, chart colour style, diagram data / layout / style / colours parts (typed: the model has dml.chart and dml.diagram)
│   │   PML and SML main parts once CR-001 Phase D lands (Phase C here)
├── CustomXmlDataStoragePart  an lxml tree (a customer's own schema)                         docx4j CustomXmlDataStoragePart
├── VMLPart                   an lxml tree (Word writes an unqualified <xml> root)          docx4j VMLPart, same departure as docx4j-core-ts
└── DefaultXmlPart            an lxml tree, for XML the registry does not know               docx4j DefaultXmlPart
```

`XmlPart[T]` is the JAXB-part analogue. `T` is the root class from the model (`Document`,
`Styles`, ...). `contents` unmarshals on first access with the fork runtime's lxml handler, a
per-call `ParserConfig` (lenient, `skipped_report=True`), then `link_parents`; the part keeps the
report as `part.skipped` and logs each entry at warning level with the part name. `set_contents`
replaces the tree and links parents. `xml` gives the bytes as they will be written: the source
bytes if not unmarshalled, else the tree serialised with the prefix table, numeric booleans, and
the `mc:Ignorable` reconciliation of 5.6. `unmarshal_all()` on the package removes laziness for
callers that prefer.

`PartRegistry` maps a content type to a part class, with a fallback on relationship type for the
generic content types (`application/xml` custom XML parts are recognised by their relationship, as
docx4j's `ContentTypeManager.getPart(partName, rel)` does). It ships with docx4j's table
(`ContentTypeManager.newPartForContentType`) and is extensible: `registry.register(content_type,
PartClass)`. Unknown XML becomes `DefaultXmlPart`, anything else `BinaryPart`; nothing is dropped.

### 5.5 Packages

```python
class OpcPackage:                                          # docx4j OpcPackage
    @classmethod
    def load(cls, source: str | PathLike | bytes | BinaryIO | PartStore, *, options: LoadOptions | None = None) -> "OpcPackage"
    content_type_manager: ContentTypeManager
    relationships_part: RelationshipsPart                  # /_rels/.rels
    parts: Parts                                           # mapping PartName -> Part, docx4j Parts
    def get_part(self, part_name: PartName | str) -> Part | None
    custom_xml_data_storage_parts: dict[str, CustomXmlDataStoragePart]   # by itemId
    source_part_store: PartStore | None
    def unmarshal_all(self) -> None
    def save(self, target: str | PathLike | BinaryIO | None = None) -> bytes | None   # zip; bytes when target is None
    def save_to(self, sink: PartSink) -> Any
    def close(self) -> None                                # also a context manager
    skipped: list[SkippedEntry]                            # union over unmarshalled parts, each entry naming its part

class WordprocessingMLPackage(OpcPackage):                 # docx4j WordprocessingMLPackage
    @classmethod
    def load(...) -> "WordprocessingMLPackage"             # raises if the main part is not a w:document
    @classmethod
    def create_package(cls, *, page_size: PageSizePaper = "A4", landscape: bool = False) -> "WordprocessingMLPackage"
    main_document_part: MainDocumentPart
    # shortcuts as docx4j: style_definitions_part, numbering_definitions_part, font_table_part,
    # document_settings_part, theme_part, header_parts(), footer_parts()
    def get_property_resolver(self) -> PropertyResolver    # Phase B
    font_mapper: Mapper | None

class PresentationMLPackage(OpcPackage): ...               # Phase C
class SpreadsheetMLPackage(OpcPackage): ...                # Phase C
```

`OpcPackage.load` sniffs the kind: it reads `[Content_Types].xml` and `/_rels/.rels`, follows the
`officeDocument` relationship, and picks the package class from that part's content type. A path
or bytes is a zip (or a directory, if the path is one); a `PartStore` is used as is.

`LoadOptions`: `strict: bool = False` (raise on the first skipped node instead of reporting),
`mce_preprocess: bool = False` (section 5.6), `registry: PartRegistry | None`.

**Load** (docx4j `Load3`): content types, then package relationships, then for every internal
relationship resolve the target part name, create the part through the registry (or link the
already-created part: a part can be the target of several relationships), attach its own `.rels`
part if present, and recurse. Parts not reachable through relationships are not loaded, as in
docx4j and as Word treats them. Part contents are not unmarshalled during load.

**Save** (docx4j `Save`): write `[Content_Types].xml` from the manager; walk the relationship graph
from the package relationships writing each part once: relationship parts and unmarshalled
`XmlPart`s are marshalled, everything else is copied from the source store or from the bytes set
on the part. `MemoryPartStore` backs new packages.

### 5.6 Markup compatibility (`mc:AlternateContent`, `mc:Ignorable`)

**On read, the default is lossless.** CR-001 established that the model round-trips both branches
of `mc:AlternateContent` as typed `docx4j_py.mce` objects wherever the schema allows them, which
is everywhere Word writes them (the drawing-in-a-run case that forced docx4j-core-ts to preprocess
parses here; CR-001 section 7). So the engine does *not* resolve alternate content by default,
and the untaken branch is not lost when a part is re-marshalled. What consumers need instead is
the *view* Word takes: the traversal utilities of CR-001 (`walk`, `iter_nodes`, `find`,
`text_of`, and `iter_children` under them) gain an `mce` mode, `"resolve"` by default, that
descends into the first `mc:Choice` whose `Requires` prefixes are all understood (the set is the
generated model's namespaces, exposed as `docx4j_py.namespaces.UNDERSTOOD`), else into
`mc:Fallback`; `mce="all"` visits everything, `mce="none"` treats `mc:AlternateContent` as a leaf.
The content API (CR-003) uses the default. Word's own text extraction and docx4j's preprocessor
give the same view, so `text_of` over a document with text boxes does not double count.

`LoadOptions(mce_preprocess=True)` gives docx4j's behaviour: an lxml pass before unmarshalling
that replaces each `mc:AlternateContent` with the chosen branch (docx4j's `McPreprocessor`), for
callers that want a simpler tree and accept the loss for parts they unmarshal.

**On write,** the root element must declare every prefix named in `mc:Ignorable` even if unused in
the tree, and Word expects the conventional prefixes. `XmlPart.xml` serialises with the prefix
table (`docx4j_py.namespaces`, docx4j's `NamespacePrefixMapper`), keeps the declarations the tree
uses, and re-declares the prefixes `mc:Ignorable` names; a prefix named there that the table does
not know is kept with its source declaration (the part remembers the source root's namespace map
for this purpose). `to_xml` in CR-001 already does the fragment-level half; the part-level half is
this CR's, and the 6 parts CR-001 section 13 identified as affected are the acceptance test.

### 5.7 The `settings.xml` schema patch

CR-001 section 14.8 found that Word writes ECMA-376 second-edition `CT_StylePaneFilter` attributes
on `w:stylePaneFormatFilter`, which docx4j's schema copy types as `CT_ShortHexNumber`, so 15
attributes are dropped from `toc.docx`'s settings part. The patch (section 7 of CR-001 procedure,
marked `docx4j-python:` in `schemas/wml/wml.xsd`, listed in `schemas/PATCHES.md`) lands with
Phase A here because this CR is the first thing that saves a settings part; the round-trip harness
gains `settings.xml` and `webSettings.xml` for every sample.

## 6. WordprocessingML resolution utilities

These are ported from docx4j as-is; docx4j-core-ts CR-001 section 6 specifies the same API and
semantics and is not repeated in full. Where the Java and the TypeScript agree, the Python does
too.

### 6.1 `PropertyResolver` (docx4j `org.docx4j.model.PropertyResolver`)

Built from a `WordprocessingMLPackage`; styles, numbering and theme fonts read once; `refresh()`
after styles change. `get_document_default_ppr()`, `get_document_default_rpr()`,
`get_effective_ppr(ppr_or_style_id)`, `get_effective_rpr(rpr, ppr)` and the style-id forms,
`get_effective_table_style(tblpr)`, `get_lvl_from_heading_style`, `activate_style`,
`has_direct_rpr_formatting`. The overlay is docx4j's `StyleUtil.apply`, property-wise,
`deep_copy`ing so results never alias the style definitions; toggle properties follow ECMA-376
17.7.3; `basedOn` cycles raise `CyclicStylesException`. Results are plain `PPr` and `RPr` objects
from the model, parent `None`.

### 6.2 List numbering (docx4j `org.docx4j.model.listnumbering`)

`NumberingDefinitionsPart.get_emulator()` and the definitions: `ListNumberingDefinition` per
`w:num` (with `lvlOverride` / `startOverride` applied), `AbstractListNumberingDefinition` per
`w:abstractNum`, `ListLevel` per `w:lvl`. `Emulator.get_number(pkg, ppr)` and the
`(pkg, pstyle_val, num_id, ilvl)` form return docx4j's `ResultTriple`: `num_string`, `num_font`,
`is_bullet`, plus the level's `ind`. Counters are stateful and callers walk paragraphs in document
order. Number formats as in the TypeScript CR: `decimal`, `decimalZero`, `lowerLetter`,
`upperLetter`, `lowerRoman`, `upperRoman`, `bullet`, `none` now; the enclosed and East Asian
formats as docx4j has them later. `unmarshal_default_numbering()` embeds docx4j's default
definitions as a resource for CR-003's `start_new_list`.

### 6.3 Fonts (docx4j `org.docx4j.fonts`)

`RunFontSelector` decides which *document* font a run, or a character range within it, uses:
effective `rFonts` and `hint`, theme references resolved through `ThemePart`, and the script of
the characters by Unicode range. `Mapper` maps a document font name plus bold and italic to a
`PhysicalFont`. This CR ships `IdentityPlusMapper`: the name itself, docx4j's metric-compatible
substitutes (Arial and Liberation Sans, Times New Roman and Liberation Serif, Calibri and Carlito,
Cambria and Caladea, Courier New and Liberation Mono) and the font table's `altName`s. A
`BestMatchingMapper` over installed fonts and embedded `ObfuscatedFontPart`s, using `fontTools`,
is a later CR; the interface is fixed here. `MainDocumentPart.fonts_in_use()` and
`styles_in_use()` as in docx4j.

## 7. Package layout and exports

```
docx4j_py/
  openpackaging/
    part_name.py, content_types.py, stores.py (PartStore, PartSink and implementations),
    load.py, save.py, mce.py, exceptions.py
    parts/        part.py, binary_part.py, xml_part.py, default_xml_part.py, relationships_part.py,
                  namespaces.py (relationship types), registry.py, wml/, dml/, docprops/, customxml/
    packages/     opc_package.py, wordprocessingml_package.py, (presentationml_package.py, spreadsheetml_package.py in Phase C)
  model/          property_resolver.py, style_util.py, listnumbering/, fonts/
```

`docx4j_py.openpackaging` and `docx4j_py.model` are the public names; `docx4j_py` re-exports
`WordprocessingMLPackage`, `OpcPackage`, `load` and `create_package` for the common case. Import
cost of the engine on top of the model must stay under 100 ms; `zipfile` and `lxml` are already
imported by the runtime path.

## 8. Tests

- **Fixtures:** the 12 documents already in `samples/` (docx4j `sample-docs`, Apache-2.0), plus a
  document saved by a current Word with w16 markup, a `.dotm`, a `.docm`, one `.pptx` and one
  `.xlsx` for the generic loading path.
- **Round trip:** load, save, reload. Untouched parts byte-identical (they come from the source
  store); for an unmarshalled part saved unchanged, the canonical XML identical to the source,
  which `scripts/roundtrip.py` already proves for 50 parts and which this CR extends to every part
  of every fixture, settings and webSettings included, with the skipped-content verdict as the
  gate.
- **MCE:** the `mce="resolve"` view of `text_of` on the text-box fixtures equals the text Word
  reports; `mce_preprocess=True` produces the same tree as docx4j's preprocessor on the same
  input (a small Java harness, output committed).
- **Parity with docx4j:** golden files produced by a Java harness under `test/java/` (run by
  hand, output committed): for each fixture and paragraph, docx4j's effective `PPr` and `RPr` as
  XML and the numbering string; the Python resolver must produce the same. docx4j-core-ts planned
  the same harness and has not built it yet; build it here and share the golden files with that
  repository.
- **Word acceptance** is manual: saved output opens in Word without repair prompts. The checklist
  from docx4j-core-ts `test/README.md` is copied and extended with the settings part.
- **Threads:** `scripts/threads.py` extended to whole packages, with one `ParserConfig` per
  thread as CR-001 section 14 requires.
- CI: Python 3.12, 3.13, 3.14.

## 9. What the object model needs from this

Pulled forward from CR-001 Phase D, small and needed by Phase A here:

- **`docx4j_py.relationships`** generated from `schemas/relationships.xsd` (two classes,
  `Relationships` and `Relationship`, with `TargetMode`), with docx4j's names.
- **`docx4j_py.docprops`** generated from `schemas/docProps/` (core with its Dublin Core imports,
  extended, custom), for the three properties parts.
- The `w:stylePaneFormatFilter` schema patch of 5.7.
- **An `mce` mode on `iter_children`** and the traversal functions built on it (5.6), with
  `docx4j_py.namespaces.UNDERSTOOD` derived from the generated packages' namespaces at generation
  time rather than hand-kept.

Nothing else: `el`, `wml(...)`, `to_xml`, `deep_copy`, `link_parents`, `warm_up()` and the
skipped-content report suffice.

## 10. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | `openpackaging/` core: stores, `PartName`, content types, relationships, parts, registry, `OpcPackage` and `WordprocessingMLPackage`, load and save for zip, directory and memory, `create_package`, MCE view and preprocessor, `mc:Ignorable` on write, settings patch, round-trip and thread tests, Word acceptance checklist | 5 days |
| B | `PropertyResolver` and `StyleUtil`; numbering `Emulator`; `RunFontSelector` and `IdentityPlusMapper`; the Java parity harness and golden files | 6 days |
| C | PML and SML packages with main parts (after CR-001 Phase D); flat OPC store and sink; docs and examples (server, MCP) | 3 days |

Phase A alone is useful: a typed docx round trip with edits through `el` and the builders. CR-003
(the content API) can start on Phase A.

## 11. Open questions (decided 2026-09-12)

All six recommendations below were accepted on 2026-09-12 and are now decisions.

1. **MCE on read: lossless with a resolving view, or preprocess as docx4j does.** Recommendation:
   lossless by default with `mce="resolve"` traversal (5.6), `mce_preprocess=True` as the option.
   This is a deliberate departure from docx4j and from docx4j-core-ts, justified by the model
   handling both branches; it should be reviewed once CR-003 has used it.
2. **Dirty tracking.** docx4j re-marshals every unmarshalled part; a Python engine could compare
   the tree to the source. Recommendation: docx4j's rule, no dirty tracking. Simple, safe, and
   the cost is bounded by what the caller chose to unmarshal.
3. **`Parts` as a `dict` subclass or a mapping wrapper.** Recommendation: a `Mapping[PartName,
   Part]` wrapper with `add`, `remove`, and lookup by string, so case-insensitivity is enforced.
4. **Saving to the loaded path.** Recommendation: temporary file and rename, with the source zip
   closed and reopened afterwards so the package remains usable.
5. **Strict loading default.** Recommendation: lenient with the report, as CR-001 decided; strict
   is an option for tests and for callers who prefer to fail.
6. **Case-insensitive part names.** Recommendation: yes for equality and lookup, preserving the
   stored spelling on write, as docx4j-core-ts decided.
7. **(added 2026-09-12, undecided) Text boxes in `text_of`.** docx4j's `TextUtils` includes text
   inside `w:txbxContent` (it traverses everything); Word's main story and Office JS `body.text`
   exclude it. Recommendation: follow docx4j, include it, through the resolving view so a box is
   counted once; give `text_of` a `text_boxes=False` switch for the Word view, which CR-003's
   `body.text` will use. Needs the wildcard typing of 12.6 point 1 or a wildcard-aware walk of
   the box's content.

## 12. Phase A implementation notes (2026-09-12)

Phase A is done. The engine is `docx4j_py/openpackaging/`, 25 modules and 5,513 lines of hand
written Python over the generated model; the model gained the four things section 9 asked for.
No fork commit was needed: nothing in Phase A wanted a runtime feature `docx4j-xsdata` does not
already have, so `~/git/docx4j-xsdata` is unchanged at CR-001 Phase C's state.

### 12.1 What is in

```
docx4j_py/openpackaging/
  __init__.py            the public names of section 7, all of them
  api.py                 load() and create_package(), the two docx4j_py itself re-exports
  exceptions.py          Docx4JException and its three subclasses
  part_name.py           PartName: OPC Annex A, resolve/relativize, rels_for
  content_types.py       ContentTypes (transcribed) and ContentTypeManager
  stores.py              PartStore/PartSink and the six implementations
  load.py                LoadOptions, load_package  (docx4j io3.Load3)
  save.py                save_package               (docx4j io3.Save)
  mce.py                 the optional preprocessor  (docx4j mc-preprocessor.xslt)
  resources.py           docx4j's default styles, numbering and fontTable
  parts/
    part.py              Part
    binary_part.py       BinaryPart, ImagePart, EmbeddedPackagePart, OleObjectBinaryPart,
                         ObfuscatedFontPart, TrueTypeFontPart, VbaProjectBinaryPart,
                         AlternativeFormatInputPart
    xml_part.py          XmlPart[T]: lazy contents, skipped, xml (5.6 on write)
    default_xml_part.py  DefaultXmlPart, CustomXmlDataStoragePart, VMLPart
    relationships_part.py RelationshipsPart, AddPartBehaviour
    namespaces.py        Namespaces: every relationship type, transcribed
    registry.py          PartRegistry and docx4j's table
    parts_map.py         Parts
    wml.py, dml.py, docprops.py   the typed parts of 5.4
  packages/
    opc_package.py       OpcPackage
    wordprocessingml_package.py   WordprocessingMLPackage, create_package, PAGE_SIZES
```

and in the model (section 9): `docx4j_py.relationships`, `docx4j_py.docprops.*`, the
`CT_StylePaneFilter` schema patch, and the `mce` mode on `iter_children`, `walk`, `iter_nodes`,
`find` and `text_of` with a **generated** `docx4j_py.namespaces.UNDERSTOOD`.

### 12.2 The numbers

Corpus: 16 documents in `samples/` — the 13 `.docx` of CR-001, plus `Normal.dotm`,
`loadAndSave.pptx` and `loadAndSave.xlsx` (all docx4j, Apache-2.0). No `.docm` is in docx4j's
repository to copy; the `.dotm` covers the macro-enabled content type and the registry maps the
`.docm` one to the same `MainDocumentPart`.

| | value |
|---|---:|
| documents loaded | **16** |
| parts loaded through the relationship graph | **223** |
| relationship parts written | 56 |
| non-relationship entries on an untouched save | 223 |
| of those, **byte identical to the source** | **223 / 223** |
| typed XML parts unmarshalled and re-serialised (WordprocessingML) | **141** |
| of those, **canonically identical to the source** | **141 / 141** |
| real differences | **0** (34 boolean respellings, benign, as CR-001 section 13.1) |
| **skipped content, WordprocessingML** | **0** |
| skipped content, the pptx and the xlsx | 5 items in 3 parts, section 12.5 |
| engine import, on top of the model | **24 ms** (budget 100 ms) |
| eight threads, 39 whole-package jobs | **0** differences against the sequential run |
| tests | **576** in this repository, 295 of them the engine's |

The round trip that matters is now every part of every document rather than CR-001's 50: it
includes `settings.xml` and `webSettings.xml` (which the `CT_StylePaneFilter` patch made lossless),
`docProps/core.xml` and `docProps/app.xml`, `theme1.xml`, `footnotes.xml`, `endnotes.xml`,
`stylesWithEffects.xml` and the `w15` comment parts.

### 12.3 Departures from sections 4 to 7, all deliberate

1. **One `ImagePart`, not nine.** docx4j has `ImagePngPart`, `ImageJpegPart`, `MetafileEmfPart`
   and six more; the kind of an image *is* its content type and a Python consumer switches on
   that string more naturally than on a class. docx4j-core-ts made the same departure. The same
   applies to `DefaultXmlPart` standing in for docx4j's per-part classes where this build has no
   bindings (section 12.5).
2. **`parts/wml.py`, `parts/dml.py`, `parts/docprops.py` are modules, not packages.** Section 7's
   layout says `wml/`, `dml/`, `docprops/`, `customxml/`. The classes are four lines each and
   Python's import cost is per module; `customxml` folded into `docprops.py` because the only
   class in it, `CustomXmlDataStoragePropertiesPart`, is read for one attribute.
3. **`part.data` rather than `part.bytes` on `BinaryPart`.** `bytes` is a builtin and shadowing it
   on an object people call `part` reads badly. `set_bytes()` keeps docx4j's name.
4. **`XmlPart.model_class_path` is a string.** Every typed part names its model class as
   `"docx4j_py.wml.Document"` and resolves it on first use. Importing *anything* under `docx4j_py`
   runs the generator's import-order manifest and costs 0.6 s (CR-001 section 13.5 point 3), so an
   engine that named its classes eagerly would put that cost on `import docx4j_py.openpackaging`
   whether or not a document was ever opened. A test greps the engine for a module-level model
   import.
5. **A per-part `bool_format`.** CR-001 decided question 5 chose numeric booleans because Word
   writes `w:val="1"`. Word writes `<ScaleCrop>false</ScaleCrop>` in `docProps/app.xml`, so the
   three document-properties parts carry `bool_format = "words"`. Same `xsd:boolean`, and matching
   Word per part is what turns "equivalent" into "byte comparable".
6. **`mce_preprocess` resolves a choice by namespace, not by a property.** docx4j's XSLT compares
   the `Requires` *prefix tokens* against `docx4j.jaxb.mc.preferChoice`, which is empty by default,
   so docx4j out of the box always takes the fallback. This resolves each prefix against the
   element's own declarations and asks whether the namespace is one the model understands, which is
   what ECMA-376 Part 3 10.2.1 says and what Word does. docx4j's `mc:AlternateContent`-inside-a-run
   exemption is not reproduced either: it exists because docx4j's exporters resolve it later, and
   this model parses it wherever it appears.
7. **The MCE preprocessor is markup compatibility only.** docx4j's `mc-preprocessor.xslt` also
   imports Strict documents and repairs the malformed nesting Google Docs, pandoc and SSRS emit.
   The Strict import is here, at the relationship level (`RelationshipsPart.import_strict`, and
   `Load3`'s `wasStrict` flag); the nesting repairs are not, and are a later CR.
8. **An absolute relationship target is treated as external** even when `TargetMode="External"` is
   missing. `samples/tables.docx` has `file:///D:\…\logo.png`, and docx4j throws
   `IllegalArgumentException` on it from `URIHelper.resolvePartUri`; making one malformed
   relationship unload the whole package is not a useful behaviour.
9. **`Parts` is a `Mapping` wrapper** (decided question 3) with `add`, `remove` and `rename`;
   `part.part_name = ...` re-keys it, which docx4j does by hand at each call site.
10. **`PartName.extension` uses the last dot.** docx4j's `PartName.getExtension` does too, but its
    `ContentTypeManager.getPart` uses `indexOf(".")`; the last dot is the right answer and the two
    only disagree on a part name with a dot in a directory segment, which Word does not write.
11. **`RelationshipsPart` writes the relationships namespace as the default one**
    (`<Relationships xmlns="…">`), as Word does, through a `default_namespace` hook on `XmlPart`.
    Everything else keeps docx4j's prefix table.
12. **No `ImageBrokenPart`.** docx4j creates one when a content type is missing for an image;
    here the registry falls back to the extension's content type and keeps the bytes, and `Save`
    therefore has no part to skip.

Two behaviours that *look* like departures and are not: `AddPartBehaviour.RENAME_IF_NAME_EXISTS`
turns `image1.png` into `image12.png` rather than `image2.png` (docx4j's own arithmetic, pinned by
a test), and `ContentTypeManager.add_content_type` writes an `Override` for a PNG or a GIF
(docx4j's JPEG/GIF/PNG special cases compare a *content type* against the strings `"gif"` and
`"png"` and can never fire).

### 12.4 What the two schema patches bought

`schemas/PATCHES.md` gained two entries, both marked `docx4j-python:`:

* **`CT_StylePaneFilter`** (section 5.7): `w:stylePaneFormatFilter` was typed `CT_ShortHexNumber`,
  so Word's fifteen second-edition attributes were dropped from every `settings.xml`.
  `toc.docx` lost 15; it now loses 0, and `settings.xml` joined the round trip.
* **`w14:docId` before `w15:chartTrackingRefBased`** in `CT_Settings`. docx4j's copy declares them
  the other way round, and every `settings.xml` in the corpus that has both writes `w14:docId`
  first, so a serialiser following the schema reordered two elements of a part it had only read.

### 12.5 The one gap found, and it is not WordprocessingML

The DrawingML **chart** and **spreadsheet-drawing** schemas in this build do not declare
`mc:AlternateContent` on `c:chartSpace` or on the anchors, so unmarshalling one of those parts
drops it — the skipped-content report says so, which is the report doing its job. It costs 5 items
in 3 parts of `loadAndSave.pptx` and `loadAndSave.xlsx`. It does **not** affect any round trip,
because nothing unmarshals those parts unless a caller asks: both files still save byte for byte.
WordprocessingML has the wildcards it needs (CR-001's `anyAttribute` patch and the compound fields
the generator produces); the chart namespaces get the same treatment when CR-001 Phase D generates
PML and SML. `tests/openpackaging/test_round_trip.py` pins the gap so that closing it is visible.

Three parts are `DefaultXmlPart` (an lxml tree) rather than typed, for the same reason — their
namespaces are outside `schemas/wml/wml.xsd`'s import closure and so have no bindings in this
build: `commentsIds.xml` (`w16cid`), `commentsExtensible.xml` (`w16cex`) and
`customXml/itemPropsN.xml` (`ds:datastoreItem`, in `officeDocument/2006/customXml`). Nothing is
lost: a tree holds everything, and a part that is never parsed is copied verbatim. The chart style,
chart colour style, `chartex` and diagram-drawing parts are trees for the same reason.

### 12.6 MCE: what "resolve" actually changed

Decided question 1 kept the read lossless and put the resolution in the *view*. The mechanism is
`docx4j_py.child.iter_children(..., mce=...)`, and `walk`, `iter_nodes`, `find` and `text_of` take
it and default to `"resolve"`:

* **`"resolve"`** splices the chosen branch's children in place of the `mc:AlternateContent`, so a
  caller never sees the element, an `mc:Choice` or an `mc:Fallback` — exactly the tree docx4j's
  preprocessor produces, without rewriting anything. The branch is the first `mc:Choice` whose
  `Requires` prefixes are all in `docx4j_py.namespaces.UNDERSTOOD`, else the `mc:Fallback`.
* **`"all"`** is the whole tree, and is what `link_parents` and `iter_tree` use and always will:
  every node is written back on save, so every node needs its parent.
* **`"none"`** treats `mc:AlternateContent` as a leaf.

`UNDERSTOOD` is prefix → namespace for every *generated* package, written into
`docx4j_py/el_index.py` by `codegen/generate_el.py` and re-exported as
`docx4j_py.namespaces.UNDERSTOOD`. It is deliberately **narrower** than the prefix table: 61
namespaces against 130. A prefix the table knows but nothing can parse — `wpi`, `v`, the PML and
SML extensions — must not count as understood, because saying it does is the one mistake
`mc:AlternateContent` exists to prevent. Adding a schema to `codegen/generate.sh` widens the
resolver by itself.

**Effect on `text_of` over the corpus, corrected 2026-09-12.** The first write-up of this section
claimed that `text_of(document)` returned `DrawingML_GraphicData_wps.docx`'s text once and
`mce="all"` twice. Checked by hand while writing the README, that was wrong: the string in
question ("Milk glass") is a body paragraph outside the `mc:AlternateContent`; the text box says
"Funky chicken", once in the `wps` choice and once in the VML fallback, and **at document level
`text_of` reports it in no mode at all**, because a run's items are read by qualified name
(`w:t`, `w:tab`, `w:br`, `w:sym`) and an `mc:AlternateContent` among them contributes nothing.
Called on the `AlternateContent` object itself, `text_of` does what the design says: "Funky
chicken" once resolved, twice with `mce="all"`, and it reads the wildcard (`AnyElement`) text to
do it. The corpus test was passing vacuously (resolved equals all when neither sees the box) and
now tests the object directly.

Two facts behind this, both to carry forward:

1. **The branches of `mc:AlternateContent` are wildcard trees, not typed objects.** The mce
   schema declares them as `xs:any`, so a `w:drawing` inside a choice is an `AnyElement`, round
   trips losslessly, and is a typed `Drawing` only under `LoadOptions(mce_preprocess=True)`.
   `find(document, Drawing)` therefore misses drawings that Word wrapped in alternate content
   (every shape a current Word writes) unless preprocessing is on. Section 5.6's resolving view
   is still right about *which branch* to read; it was wrong to imply the branch is typed. A fork
   runtime feature could type wildcard content by the enclosing class's choices (the `R` that
   holds the `AlternateContent` knows that `w:drawing` is a `Drawing`), wrapping the result in a
   `DerivedElement` so it serialises under the right name. That is proposed for Phase C, and until
   then consumers that need typed drawings use `mce_preprocess=True`.
2. **Whether `text_of` should descend into text boxes is a decision, not a bug.** docx4j's
   `TextUtils` walks everything and so includes `w:txbxContent`; Word's main story, and Office JS
   `body.text`, exclude shapes. The Python `text_of` currently excludes them by accident of the
   run rule above. Open question 7 below records the choice; the README makes no claim either way.

### 12.7 Threads and import

`tests/openpackaging/test_threads_and_import.py` runs **39 whole-package jobs** — load, unmarshal
every typed part, save, split the zip — over the 13 `.docx` on **8 threads** against the one
process-wide `XmlContext`, and compares every byte against a sequential run: **0 differences, 0
failures**. A second test runs the skipped-content report the same way and finds every part's
report is its own, which is CR-001 section 14.7's rule (one `ParserConfig` per thread) holding at
the package level: `XmlPart` takes a fresh parser per unmarshal. A third creates 24 packages
concurrently, which is the one place the engine shares a resource (docx4j's default styles).

Import: **24 ms** for the engine on top of the model, against section 7's 100 ms budget, measured
in a fresh subprocess three times. Importing the engine *cold* is 681 ms, of which 657 ms is the
model, because `docx4j_py/__init__.py` carries the import-order manifest — CR-001 Phase D's lazy
per-namespace import is what moves that, not anything here.

### 12.8 Word acceptance

**Passed 2026-09-12.** `scripts/acceptance.py` writes the four artefacts `tests/README.md`'s
checklist needs into `out/acceptance/`: an untouched round trip, a round trip with the main part,
styles and settings re-marshalled and two paragraphs added, a document created from nothing, and a
created document with an image part placed by a `w:drawing`. The user opened all four in Word on
2026-09-12: no repair prompts; artefact 1 shows the image with its reflection and no text, which
is what its source (`samples/2016_image_with_text_effects.docx`, one paragraph holding the image,
no `w:t`, no comments part) contains; artefact 2 shows the two added paragraphs; artefacts 3 and 4
are correct. On artefact 1 the only entries that differ from the source are the two relationships
parts, which the Save rule always re-marshals: same relationships, attribute order
`Target, Type, Id` where Word writes `Id, Type, Target`, and the zip entry order follows the
relationship walk rather than the source. Both are cosmetic; writing Word's attribute order is a
cheap follow-up if byte-identical `.rels` is ever wanted.

The same run found that artefacts 3 and 4 carried an **empty** `docProps/app.xml` and
`docProps/core.xml`: no `Application`, no `AppVersion`, nothing in the core part. That is what
docx4j does unless `docx4j.App.write` and `docx4j.dc.write` are set, and it is wrong for a created
document. Fixed 2026-09-12: `create_package` now writes `Application` (`docx4j-python`) and
`AppVersion` in Word's `XX.YYYY` form (docx4j's `Save` notes that any other form, a `-SNAPSHOT`
suffix say, makes Word 2010 x64 report corruption; `0.0.1` becomes `0.0001`), and `dcterms:created`
and `dcterms:modified` as `W3CDTF` values at UTC second precision, which the serialiser writes
with the `xsi:type` Word expects. Creator and last-modified-by are left to the caller. The
extended-properties root is now written in the default namespace, as Word writes it. Loaded
documents are unaffected: their properties parts are copied byte for byte unless unmarshalled, and
nothing is rewritten at save time (decided question 2). Tests in
`tests/openpackaging/test_docprops_defaults.py`; artefacts 3 and 4 regenerated for a re-check.

### 12.9 Deferred to Phase B and C

* **Phase B**, unblocked and unchanged: `PropertyResolver` and `StyleUtil`; the numbering
  `Emulator`; `RunFontSelector` and `IdentityPlusMapper`; the Java parity harness and its golden
  files. `MainDocumentPart.fonts_in_use()` and `styles_in_use()` go with them.
* **Phase C**: `PresentationMLPackage` and `SpreadsheetMLPackage` with typed main parts (after
  CR-001 Phase D); `FlatOpcPartStore` and `FlatOpcPartSink`; the chart-namespace schema patches of
  12.5; docs and examples. **Half of the flat OPC item is done**: CR-003 Phase C needed to read a
  `pkg:package` for `insert_ooxml` and wrote
  `stores.FlatOpcStore`, a read-only `PartStore` of 89 lines that
  `load()` takes like any other (a `pkg:xmlData`'s root element serialised, a `pkg:binaryData`
  base64-decoded, and the `[Content_Types].xml` a flat package does not carry synthesised from
  each part's `pkg:contentType`, as docx4j's `FlatOpcXmlImporter` does). What remains is the
  **sink**: `FlatOpcXmlExporter`, which `tests/content/test_ooxml.py` and `scripts/acceptance.py`
  each approximate in 25 lines of test code. CR-003 section 14.3 records the decision.
* **Not started, not blocking**: `OpcPackage.clone()`, external resource loading
  (`Load.loadExternalTargets`), the digital-signature parts, and docx4j's
  `DrawingPropsIdTracker`, which only matters once something adds drawings in bulk.
