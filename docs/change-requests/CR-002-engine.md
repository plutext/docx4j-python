# CR-002: The engine: container interface, Open Packaging layer, typed parts, resolution utilities

**Status:** Proposed 2026-09-12; open questions decided 2026-09-12 (section 11); **Phase A implemented
2026-09-12** (section 12); Word acceptance passed 2026-09-12 (12.8); amended 2026-09-17, a created
document declares `compatibilityMode` 15 (section 12.10); amended again 2026-09-17, **section
6.2's numbering emulator landed early through CR-003 Phase H** (section 12.11), as the flat OPC
read half landed through its Phase C; amended 2026-09-19, the marshal is one pass (12.13) and the
parent pointers are wired during the parse (12.14); **Phase B implemented 2026-09-19** (section
15): `PropertyResolver`, `StyleUtil`, the numbering completed and the fonts, with **zero
differences on all 45 of docx4j's parity goldens** --- 984 style ids, 618 paragraphs, 1,113 runs,
8 tables, 183 numbered paragraphs and 1,112 font spans. Word acceptance of the five created
artefacts that gained a theme part is outstanding (15.5). Phase C proposed.
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

*The counting core of this section was implemented on 2026-09-17, ahead of Phase B, because
CR-003 Phase H needed list labels: section 12.11 records what landed and what is left.*

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
- **Parity with docx4j:** golden files produced by a Java harness: for each fixture and
  paragraph, docx4j's effective `PPr` and `RPr` as XML and the numbering result; the Python
  resolver must produce the same. *Decided 2026-09-19: the harness is built in docx4j-core-ts
  (`test/java/`, against a pinned docx4j build) and its JSON goldens are committed there under
  `test/golden/<fixture>.json`; this repository **consumes** those files and builds no harness
  of its own (section 12.12). The property values inside are XML strings docx4j marshalled, to
  be unmarshalled here and compared as object trees, never as text.* **Done 2026-09-19 (section
  15): all 45 fixtures are in `tests/fixtures/`, three test files read the goldens ---
  `tests/content/test_properties_goldens.py`, `test_numbering_goldens.py` and
  `test_fonts_goldens.py`, over the walk they share in `tests/content/goldens.py` --- and every
  recorded answer is reproduced. Zero differences.**
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
| B | `PropertyResolver` and `StyleUtil`; numbering `Emulator` (its core landed early, section 12.11); `RunFontSelector` and `IdentityPlusMapper`; the parity test over docx4j-core-ts's committed goldens (section 12.12; the harness is not built here) | 5 days |
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
   `Load3`'s `wasStrict` flag); the nesting repairs are not, and are a later CR. **Measured
   2026-09-21** on docx4j's four Word-saved strict fixtures
   (`docx4j-core-tests/src/test/resources/strict/`): `load()` sets `was_strict` and rewrites the
   relationship types, but the parts' *content* stays in the `purl.oclc.org` namespaces the model
   does not know, so the main part parses to a `Document` whose `w:body` is **skipped** (reported
   by the skipped-content report, so nothing is lost silently; `pkg.body` then raises
   `body.no_content`). A strict document is therefore loadable and saveable byte for byte but has
   no typed content. The part-level conversion is the later CR, and it takes docx4j's two fixes of
   2026-09-21 as rules from the start: (1) strict universal measures to transitional integers ---
   ×2 for the half-point family (`sz`, `szCs`, `kern`, `position`, `hps`, `hpsRaise`,
   `hpsBaseText`), ×20 for every other point value including `w:line` under `lineRule="auto"`
   (docx4j `7892cb501`); (2) the 2010 shape, group and canvas elements that strict keeps in the
   wordprocessingDrawing namespace (`wp:wsp`, `wp:spPr`, `wp:txbx`, `wp:bodyPr`, `wp:wgp`,
   `wp:wpc`, ...) map to `wps:` / `wpg:` / `wpc:` by root name, children by nearest such
   ancestor, not to transitional `wp:` (docx4j `3f4776c35`). The fixtures and
   `StrictLoadTest.java` are the oracle.

   **The repair side, measured the same day** against the XSLT's four repair classes (docx4j-11's
   list), on fragments through `wml()` strict and `runtime.parser(lenient=True)`: (1) a decimal
   in a twips attribute (Google Docs, pandoc: `w:spacing/@w:after="8.5"`,
   `w:ind/@w:hanging="141.99999999999994"`) and (2) SSRS's empty `w:rsidR=""` /
   `w:pgMar/@w:header=""` are **kept as the raw string** --- the fork's converter fails, warns
   (`ConverterWarning`), and leaves the `str` in the `int` field, so the value round-trips byte for
   byte and the content API's arithmetic still reads it (`space_after` 0.425 cm), but it is not
   typed and **the skipped-content report does not carry it**: a fidelity gap in the report, not in
   the file, and the later CR's first item (the report gains `kind="value"` for a converter
   failure; the rounding and dropping docx4j applies follow). (3) `wordml201011:*` attributes and
   (4) the malformed nesting (`w:r` in `w:r`, `w:p` in `w:r` or `w:hyperlink`) are **dropped and
   reported**, never silently: the subtree is a `SkippedNode(kind="element")` in `part.skipped`,
   `wml()` raises `FragmentError`, `LoadOptions(strict=True)` raises. No hoisting or splitting is
   done; docx4j 17.2.0's `MalformedNestingTest.java` (17 shapes) is the oracle when it is. The
   trigger differs from docx4j's retry-after-`UnmarshalException`: every parse is the lenient one
   with the report, once.
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

*As found, 2026-09-12.* The DrawingML **chart** and **spreadsheet-drawing** schemas in this build
do not declare `mc:AlternateContent` on `c:chartSpace` or on the anchors, so unmarshalling one of
those parts drops it — the skipped-content report says so, which is the report doing its job. It
costs 5 items in 3 parts of `loadAndSave.pptx` and `loadAndSave.xlsx`. It does **not** affect any
round trip, because nothing unmarshals those parts unless a caller asks: both files still save
byte for byte. WordprocessingML has the wildcards it needs (CR-001's `anyAttribute` patch and the
compound fields the generator produces); the chart namespaces get the same treatment when CR-001
Phase D generates PML and SML. `tests/openpackaging/test_round_trip.py` pins the gap so that
closing it is visible.

**Closed 2026-09-20, by the schema refresh** (CR-001 section 14.10). docx4j's CR-024
(`bcb4c7f57`) made `dml-chart.xsd` admit `mc:AlternateContent` beside `c:style` and
`dml-spreadsheetDrawing.xsd` admit it in `EG_Anchor`, `EG_ObjectChoices` and `CT_GroupShape`, and
the regenerated `dml.chart` and `dml.spreadsheet_drawing` packages carry it as a typed
`AlternateContent` in the compound fields (`ChartSpace.alternate_content`, the anchors' `content`).
The 5 items in 3 parts were **3 elements and 2 attributes**: the three `mc:AlternateContent`
(`c:chartSpace` in both charts, `xdr:oneCellAnchor` in the drawing) are kept now and the three
parts marshal canonically identical to their source, which
`test_the_pptx_and_xlsx_chart_parts_keep_their_alternate_content` asserts. **The two attributes
are a second gap the first one's count had hidden**, pinned in its place by
`test_the_c16r3_chart_boolean_is_the_one_known_unmarshal_gap`: `c16r3:dispNaAsBlank/@val`
(`dataDisplayOptions16`, [MS-ODRAWXML]'s 2017/03 chart namespace), one in each chart. Office
writes the attribute unqualified; docx4j's `odrawxml/office_drawing_2017_03_chart_.xsd` is
`attributeFormDefault="qualified"`, so the binding looks for `c16r3:val` and drops `val`. docx4j's
own JAXB binding has the same blind spot, so the fix is a hand-off to docx4j (the schema's
`attributeFormDefault`, or `form="unqualified"` on that one attribute) rather than a patch
here — a chart namespace is outside `schemas/PATCHES.md`'s WML remit, and it costs nothing on the
ordinary path, where those parts are never unmarshalled. **Closed the same day**: docx4j
`c363a969f` made the schema `attributeFormDefault="unqualified"` (and its prefix `c16r3`, not
`c173`); the copy was taken (CR-001 section 14.11) and
`test_the_c16r3_chart_boolean_keeps_its_val` now asserts the closed state --- every part of both
fixtures unmarshals with nothing skipped, and the option is written back `c16r3:dispNaAsBlank
val="1"`. There is no known unmarshal gap over `samples/` any more.

Two parts are `DefaultXmlPart` (an lxml tree) rather than typed, for the same reason — their
namespaces are outside `schemas/wml/wml.xsd`'s import closure and so have no bindings in this
build: `commentsExtensible.xml` (`w16cex`) and `customXml/itemPropsN.xml` (`ds:datastoreItem`, in
`officeDocument/2006/customXml`). `commentsIds.xml` (`w16cid`) was a third until the refresh
brought `w16cid.xsd` into the closure; its `CTCommentsIds` is generated now, though the part is
still registered as a tree (a typed part is CR-001 Phase D's or a later CR's to add). Nothing is
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

* ~~**Phase B**~~ **implemented 2026-09-19, section 15**: `PropertyResolver` and `StyleUtil`; the
  numbering `Emulator` completed (its counting core had landed early through CR-003 Phase H on
  2026-09-17, section 12.11, and 15.2 is the three things that section said were left);
  `RunFontSelector` and `IdentityPlusMapper` with `MainDocumentPart.fonts_in_use()` and
  `styles_in_use()`; and the parity test over the goldens docx4j-core-ts commits (**the harness
  is built there, not here**: decided 2026-09-19, section 12.12), which finds zero differences on
  all 45.
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

### 12.10 A created document is Word 2013, not Word 2007 (2026-09-17)

A dated amendment to Phase A's `create_package`, from the same session as CR-003
sections 17.10 and 17.11.

Phase A wrote one compatibility setting into `/word/settings.xml`,
`overrideTableStyleFontSizeAndJustification`, because that is what docx4j's
`createPackage` writes and 12.3's rule is that docx4j is the oracle. docx4j
writes **no `compatibilityMode` at all**, and Word reads a document that declares
none as **Word 2007**: every document this library created opened with
*Compatibility Mode* in Word's title bar, and Word applied its 2007 layout rules
to it. Nobody had looked at the title bar, and the Word acceptance checklist did
not ask.

`create_package` now writes what Word writes into a new document, in Word's own
order:

| `w:name` (all with `w:uri="http://schemas.microsoft.com/office/word"`) | `w:val` |
|---|---|
| `compatibilityMode` | **15** |
| `overrideTableStyleFontSizeAndJustification` | 1 |
| `enableOpenTypeFeatures` | 1 |
| `doNotFlipMirrorIndents` | 1 |
| `differentiateMultirowTableHeaders` | 1 |
| `useWord2013TrackBottomHyphenation` | 0 |

15 is Word 2013, and is what Word 2016, 2019 and 365 write too: there has never
been a 16. The table is `NEW_DOCUMENT_COMPAT` in
`docx4j_py/openpackaging/parts/wml/__init__.py`, beside `WORD_COMPAT_URI`, and
`DocumentSettingsPart.set_default_compat_settings()` writes it;
`set_override_table_style_font_size_and_justification` is still there, over the
new one-setting `set_compat_setting`, because it is docx4j's own member name.
The table lives in the engine rather than in the content API because
`create_package` is the engine's and the engine imports nothing from the content
API (CR-003 section 5).

**The departure from docx4j is deliberate and is this section.** A created
document targets Word 2013 and later; a caller who wants otherwise writes
`pkg.compatibility_mode = 12` (CR-003 sections 3.4 and 17.11), which is the
property this amendment exists alongside. Loaded documents are untouched, as
always: their settings part is copied byte for byte unless something unmarshals
it.

`test_a_created_document_saves_and_reloads` pins the whole `w:compat`, name,
value and URI, in Word's order. Artefacts 3, 4 and 5 were regenerated and the
Word check of them is outstanding; 12.8's rule (a change to `create_package`
means a new Word check) is why.

### 12.11 Section 6.2's numbering emulator landed early, through CR-003 Phase H (2026-09-17)

The way 12.9 records the flat OPC **read** half arriving with CR-003 Phase C:
CR-003 Phase H needed list labels (`Word.ListItem.listString`, and the markdown
exporter's markers), and a second counting engine to be thrown away when Phase B
arrives would have been the wrong thing to write. So section 6.2's **counting
core** is implemented, in the layout that section named:

`docx4j_py/model/listnumbering/`
    `definitions.py` --- `ListLevel`, `AbstractListNumberingDefinition`,
    `ListNumberingDefinition` (with `w:lvlOverride`, `w:startOverride` and the
    override `w:lvl` applied, and `w:numStyleLink` resolved as
    `NumberingDefinitionsPart.initialiseMaps` resolves it), `Counter`,
    `NumberingState` and `NumberingStates` (one per story, as docx4j's are).
    `formats.py` --- the eight number formats section 6.2 listed (`decimal`,
    `decimalZero`, `lowerLetter`, `upperLetter`, `lowerRoman`, `upperRoman`,
    `bullet`, `none`), with docx4j's fail-soft rule: an unknown `w:numFmt`, or a
    value a format cannot express, gives the decimal label and **one** logged
    warning per format. `emulator.py` --- `Emulator.get_number(package, p_pr,
    state)`, `get_number_of(package, p_style_val, num_id, ilvl, direct_num_pr,
    state)` and `peek(...)`, returning docx4j's `ResultTriple` (`num_string`,
    `num_font`, `is_bullet`, `ind`) under its newer name `NumberingResult`.

`NumberingDefinitionsPart` gains `get_emulator(reset=False)`, `numbering_state`,
`abstract_list_definitions` and `instance_list_definitions`, importing the model
inside the method as this layer must (12.7). `unmarshal_default_numbering()` was
already here, and `start_new_list` is what section 6.2 said would want it.

Against the eight Word-measured probes of docx4j's CR-014 and the numbering one
of its CR-015 --- six of which are now `tests/fixtures/numbering-*.docx` and
`styles-numpr-ilvl-only.docx` --- **every label this emulator produces is the
label Word paints**, including the two docx4j itself got wrong before CR-014
(`w:lvlRestart`, and a level reset but not yet used showing its start).

**What remains for Phase B**, and is deliberately not here:

* **Resolution through `PropertyResolver`.** docx4j's `Emulator.resolve` asks
  the resolver for the effective `w:pPr` of a paragraph style; with no resolver
  yet, `Emulator.effective_numbering` walks `w:basedOn` itself, attribute by
  attribute, and reads the styles part with lxml so that a read unmarshals
  nothing. It is the one method in the module that is not docx4j's own logic,
  and it is the one that goes when Phase B lands (CR-003 section 18.2 item 16).
* **The twenty exotic formats** (ordinal, cardinal text, Chinese, Hebrew,
  Russian, Thai, the enclosed circles): `formats.register` is the hook, and
  until they arrive each falls back to the decimal label with one warning.
* **`getInd`'s full form and the label's own run properties** (docx4j's
  `labelRPr`, the `w:lvlOverride/w:lvl` rule CR-014 probe P3 measured): the
  level's own `w:ind` and the linked style's `w:basedOn` chain are here, the
  rest belongs with `StyleUtil`.

Nothing in CR-002 Phase B changes shape because of this; the emulator is where
6.2 said it would be, and Phase B fills in the resolution underneath it.

### 12.12 The Java parity harness is docx4j-core-ts's; Python consumes its goldens (2026-09-19)

Section 8 planned a Java harness under this repository's `test/java/` and offered to share its
goldens with docx4j-core-ts. Decided the other way round on 2026-09-19, with the TypeScript
engine's Phase B implementation plan (docx4j-core-ts CR-001 section 14): the harness is a Maven
project in `docx4j-core-ts/test/java/` against a pinned local build of `../docx4j`
(17.1.1-SNAPSHOT), and it writes one JSON golden per fixture into
`docx4j-core-ts/test/golden/<fixture>.json`. Each golden carries a header (docx4j commit, date),
`styles` (the effective `w:pPr` and `w:rPr` per style id, the default paragraph style, the
document defaults), `stories` (main document, each header and footer, footnotes, endnotes,
comments: per block-level paragraph the effective `w:pPr`, the paragraph mark's `w:rPr`, the
numbering result --- `numString`, `isBullet`, `numFont`, `ind`, `ilvl`, `labelRPr`,
`notNumbered` --- under one `NumberingState` per story as docx4j's CR-014 defines them, and per
run the effective `w:rPr` and the font spans `RunFontSelector.documentFontFor` gives, folded by
code point, with the `IdentityPlusMapper` result) and `tables` (the effective table style and
whether the chain reaches the default). Property values are XML strings docx4j marshalled with
its prefixes, to be unmarshalled by each port and compared as object trees, never as text.

What this changes here:

* **No harness is built in this repository.** Phase B's parity test reads the committed goldens
  from the sibling checkout (`../docx4j-core-ts/test/golden/`) and skips when it is absent, the
  way `scripts/office_js_subset.py` and the python-docx parity test already behave. Section 8's
  estimate for Phase B drops by a day.
* **The fixtures keep their names on both sides.** The six numbering fixtures this repository
  copied from docx4j-layout-fidelity for Phase H (`numbering-default-style-numbered`,
  `numbering-lvlrestart`, `numbering-numstylelink-separate`, `numbering-shared-abstract`,
  `numbering-stories`, `styles-numpr-ilvl-only`) are copied there unchanged with provenance, so
  a golden's name matches a fixture here. The twenty `styles-*` and `fonts-*` probes and the
  docx4j-core-tests numbering and font documents come with them and are the fixtures Phase B
  adds under `tests/fixtures/` when it lands.
  docx4j-core-ts's parity set is 37 fixtures (its `test/fixtures/parity/`, provenance in its
  `test/README.md`): the 27 CR-014/015/016 probes regenerated from `Corpus.java`, nine documents
  from docx4j-core-tests, and `fonts-modesOfApplication.docx`. This repository's `lists.docx` is
  docx4j-core-tests' `numbering_indentation.docx` byte for byte, so its golden is that one.
* **Drift comes for free.** A weekly workflow in docx4j-core-ts reruns the harness against
  docx4j's `VERSION_17_1_1` head and opens a pull request when a golden moves; this repository
  sees the change when the sibling checkout is updated.
* **The emulator's tests stay.** `tests/content/test_listnumbering.py` pins the labels against
  the Word-measured answers of docx4j's CR-014 and CR-015 as it does now; the goldens add
  docx4j's own `labelRPr`, `ind` and `numFont` per paragraph, which section 12.11 lists as what
  remains, and docx4j's resolved `numId` / `ilvl` per paragraph if the golden carries them (asked
  for on 2026-09-19, so that `Emulator.effective_numbering` can be checked against
  `Emulator.resolve` directly rather than through the label).
* **The portfolio rule stands.** `python/CR-002.B` still waits on `docx4j/fo-fidelity-next`;
  the gate was lifted for `core-ts/CR-001.B` only. Its registry entry now also depends on
  `core-ts/CR-001.B`, since the goldens are that phase's output.

docx4j-core-ts's numbering port takes section 12.11's module split (definitions, state,
formats, emulator) and its "what remains for Phase B" list, so the two engines' numbering reads
alike.

**Consumed the same day.** `tests/content/test_numbering_goldens.py` holds the Phase H emulator
to 19 of the 45 goldens --- 285 paragraphs over 43 stories --- comparing `numString`, `isBullet`,
`ilvl`, `numId`, docx4j's own `numRef` (the resolution asked for above, which the harness
delivers from its new public `Emulator.numRefFor`) and the whole `stateAfter`: every counter and
every start override spent. `ind`, `indResolved`, `lvl`, `labelRPr`, `numFont`, `effectivePPr`,
`paragraphMarkRPr` and the runs' `fontSpans` are what Phase B still has to compare, and the test
says so. CR-003 section 18.10 records the fixtures copied, the story rule, the two recorded
departures the goldens surfaced (both `w:numId` 0) and the numbers.

### 12.13 The marshal is one pass, not three (2026-09-19)

Section 5.6's serialisation was the right shape and the wrong number of passes. `XmlPart._marshal`
rendered the object to a **string**, parsed that string back with lxml, cleaned the namespaces,
copied the whole tree into a new root to declare the prefixes `mc:Ignorable` names, and serialised
again: the document written twice and read once on every marshal. On `samples/Symbols.docx`'s
771 KB main part (24,758 elements) that was 887 ms, of which the copy alone --- `_redeclare`,
moving 24,758 nodes --- was 345 ms.

It is now one render and one serialisation:

* `runtime.tree_serializer()`, a new twin of `serializer()` with the same `serializer_config()`
  (numeric booleans), renders through the fork's `TreeSerializer` **into an lxml tree**;
* the `root_name` rename is `root.tag = ...` in place, not a tree copy;
* the namespace reconciliation is `etree.cleanup_namespaces(root)` followed --- only when
  `mc:Ignorable` names a prefix the cleanup dropped --- by a second pass with `top_nsmap` and
  `keep_ns_prefixes`, which declares those prefixes on the root without copying anything. Two
  passes rather than one so the declaration order does not move: the unused ignorable prefixes
  are declared after the used ones, as they were.

`_redeclare`, `_rename` and the module's `_LXML_PARSER` are gone.

**Symbols.docx: marshal 887 ms to 484 ms, 0.85 to 1.56 MiB/s.** The corpus's 14 main parts,
964 KB: 985 ms to 572 ms. Parse is untouched (302 ms, 2.49 MiB/s). Of the remaining 484 ms, 450 is
the fork's render and 19 ms is this engine, the largest piece of that being the full-tree walk for
`mc:Ignorable` attributes.

The output is the old output, part for part: `tests/openpackaging/test_marshal_one_pass.py` keeps
the old method as `reference_marshal` and compares all 150 typed XML parts of all 16 samples byte
for byte. It found one latent bug on the way, now fixed and the single deliberate difference from
any earlier run: the old method iterated a **`set`** of prefixes, so the order of the namespace
declarations on a saved root depended on `PYTHONHASHSEED` --- 49 of the 150 parts gave three
different byte strings over three seeds. Both sort now. The affected parts are canonically
identical either way, which is why the round trip never saw it.

This is Phase A of the fork's own CR-001 (`~/git/docx4j-xsdata/docs/fork/CR-001-serialiser.md`),
whose section 8 holds the benchmark tables, the stage-by-stage split, the `cProfile` its Phase B
starts from and what the tree and string writers do and do not share. Phases B to D are the fork's
and are where the other 450 ms is. The benchmark lives here, as that CR's section 3 asks:
`scripts/bench_marshal.py`.

### 12.14 The parent pointers are wired during the parse (2026-09-19)

`XmlPart._unmarshal` parsed the part and then walked the whole tree a second time with
`link_parents`, to set every node's `parent` slot and bind every `ChildList` to its owner
(CR-001 section 5). For `samples/Symbols.docx`'s main part that is 24,758 nodes visited
twice, and a `seen` set the size of the document, for work the parse could do as it goes.

The fork's CR-001 Phase D adds a **generic post-bind hook**,
`ParserConfig(on_bind: Callable[[Any], None] | None = None)`: `ElementNode.bind` calls it
with every object the parser builds, as soon as `class_factory` has built it. xsdata binds
bottom up, so at that moment the object's own children exist and are bound and its parent
does not exist yet --- exactly the moment a parent pointer wants to be set. It is off by
default and it is per config, so a parse owns its hook the way it owns its skipped report.

This engine fills it in:

* **`child.link_children(obj)`** is one level of `link_parents`, from the same
  `_child_plan` enumeration and by the same rules: every `Child` in a single-valued field
  or a list field gets `obj` as its parent, every `ChildList` is bound to `obj` as its
  owner, and a typed object nested inside an `AnyElement` wildcard keeps the parent it
  has, exactly as `link_parents` leaves it. It is one level and no deeper: called on every
  object bottom up, the levels compose into the whole tree.
* **`runtime.parser_config`** passes `functools.partial(link_children, context=context())`
  as `on_bind` unless the caller asks for something else. `parser(on_bind=None)` is the
  plain parser, and then the tree needs `link_parents` before anything reads `.parent`.
* **`XmlPart._unmarshal`** skips the `link_parents` walk when the config carries a hook,
  and does it when it does not. `set_contents` is unchanged: it takes an arbitrary tree
  from a caller, of unknown provenance, so it still walks it.

**`link_parents` is unchanged and stays.** Hand-built trees, `deep_copy`'s relink,
`set_contents`, the markdown importer, the content API's inserts and every unpickled root
need the full walk, and `scripts/parents.py` still measures the hook against it.

The numbers, on this machine in one session. Over the 50 typed parts of `samples/`
(`scripts/parents.py`, which now parses each part both ways and checks both trees) the
hooked parse is 3 to 5% ahead: 268 ms against 293, 274 against 282, 283 against 291 and
285 against 301, over four runs. On `Symbols.docx`'s main part alone it is a **wash** ---
189 ms against 193 --- because `link_parents` over that part was only 20 ms to begin with,
and the per-object callback costs about what the walk's own bookkeeping cost. The
fork's CR-001 section 10.4 records why the CR expected more: its "0.12 s of the parse is
`link_parents`" was a profiled cumulative time, not a real one.

It is kept regardless, for what is not on the clock: it is a capability the fork can offer
upstream on its own (docx4j-python `docs/UPSTREAM.md` report 15), it removes a second full
traversal and its `seen` set, and it puts the parent pointer where docx4j puts it --- at
construction --- so no window exists in which a parsed tree has no parents.

The rest of Phase D is the fork's alone and is what actually moved the number: the parse of
that part is **329 ms to 193 ms, 2.29 to 3.90 MiB/s**, and the corpus's 964 KB of main
parts 393 ms to 238 ms. With section 12.13's marshal at 250 ms, a 772 KB main part is now
443 ms to read and write where CR-001 opened at 1,114. The fork's
`docs/fork/CR-001-serialiser.md` section 10 has the tables, the profile and section 10.6's
recommendation to close that CR.

## 15. Phase B implementation notes (2026-09-19)

Phase B is done. Section 6's three resolution utilities are implemented, and the parity test
section 8 promised reads the goldens docx4j-core-ts commits (12.12) and finds **zero
differences on all 45 of them**, over the effective properties, the style resolutions, the
table styles, the list labels and counters, and the font spans.

The gate 12.12's last bullet recorded --- `python/CR-002.B` waiting on `docx4j/fo-fidelity-next`
--- was lifted by Jason on 2026-09-19, the goldens being that work's output and already
committed.

What was written, in four steps, each its own commit:

```
docx4j_py/model/
    property_catalogue.py   1,126   docx4j PropertyCatalogue: one entry per member of
                                    w:rPr, PPrBase, CT_TblPrBase and w:tcPr, in schema
                                    order, with its leaf merge rule and its emptiness
    style_util.py             561   docx4j StyleUtil, the catalogue-driven half
    property_resolver.py      731   docx4j PropertyResolver
    listnumbering/            +510  effective_numbering through the resolver, docx4j's
                                    whole format registry, getInd's full form, the
                                    typed views of a level (lvl, labelRPr, own_ind)
    fonts/                  2,817   RunFontSelector (661), Mapper and
                                    IdentityPlusMapper (769), FontFallback (328),
                                    fonts_in_use (238), the JSON loaders (216), the
                                    script table (167), ThemePart.getFont (125),
                                    the registry (140), defaultTheme (95),
                                    LanguageTagToScriptMapping and CJKToEnglish (100),
                                    PhysicalFont (83), the package exports (162)
docx4j_py/resources/fonts/  184 KB  docx4j's own tables as committed JSON, plus the
                                    three Office theme parts byte for byte
scripts/generate_fonts.py     375   what writes them, from ../docx4j
tests/content/              1,445   goldens.py and the five test files
```

Four modules of the engine changed: `XmlPart` gained `read_contents()`,
`WordprocessingMLPackage` the three accessors and a theme part in `create_package`,
`MainDocumentPart` the font walk, and `exceptions` a `CyclicStylesException`.
No fork commit was needed: nothing in Phase B wanted a runtime feature `docx4j-xsdata` does not
already have.

### 15.1 Step 1: property resolution

`PropertyResolver` is the resolution order of docx4j's CR-015 phase 2, and `StyleUtil` is the
overlay, both driven by the **property catalogue** rather than hand-written per element: one
table per properties element, one entry per schema member in schema order, each saying how to
read it, write it, merge a more specific value over an inherited one, whether it is empty and
whether it formats the content. `apply`, `is_empty`, `unset` and `has_direct_formatting` are all
derived from the tables by iteration, so the four cannot drift apart as docx4j's hand-kept lists
did before its CR-015 phase 1. Every leaf a merge carries is a **copy** (`deep_copy`), so an
effective `w:rPr`'s `w:b` is never the styles part's.

**Parity is zero differences** on all 45 goldens: the document defaults and `effectivePPr` /
`effectiveRPr` for every one of the **984** style ids, `effectivePPr` and `paragraphMarkRPr` for
every one of the **618** paragraphs of every story, `effectiveRPr` for every one of the **1,113**
runs, and the effective table style and `reachesDefaultTableStyle` of all **8** tables --- 1,602
`w:pPr` and 2,715 `w:rPr` comparisons, each as an **object tree** (ours through a marshal and an
unmarshal, the golden's XML string through an unmarshal), so neither prefixes nor attribute
order nor the class a member happens to be is part of the contract.

**Departures from section 6.1 and from the Java, all deliberate.**

1. **`CyclicStylesException` raises only when asked**, as docx4j does: its `isCyclic` reads
   `docx4j.openpackaging.exceptions.CyclicStylesException.throw`, false by default, and otherwise
   stops the walk and resolves with as much of the hierarchy as it has. Here that property is
   `style_util.throw_on_cyclic_styles(True)`. Section 6.1's "`basedOn` cycles raise
   `CyclicStylesException`" is the property's `true` branch, not the default. The TypeScript port
   took the same reading.
2. **The deprecated overloads are not ported**: the four-flag `getEffectiveRPr(styleId, ...)` and
   `getEffectiveRPrUsingPStyleRPr`, which docx4j's CR-015 phase 5 deprecated and which exist only
   for callers older than the fix; and `ImmutablePropertyResolver`, deprecated with no caller.
3. **`activate_style(style_id)` answers from what the package has.** docx4j reads `KnownStyles.xml`
   from its jar to activate a style the document lacks; this package has that resource, but it is
   the content API's `ensure_style` (CR-003 section 14.9) that reads it, with the numbering
   reconciliation a known style needs, so the resolver's string form answers `True` for a style
   already live and `False` (logged) otherwise. The object form works as docx4j's does,
   recursively for `w:basedOn` and `w:link`.
4. **The resolver reads the styles part with a new `XmlPart.read_contents()`**, which unmarshals
   into a *private* tree without marking the part unmarshalled. Without it, building a resolver
   would cost every document its byte-for-byte round trip of `word/styles.xml`, since this engine
   re-marshals any part that was unmarshalled --- and the content API now builds a resolver on
   every effective read. When something else unmarshals the part for real (the comment styles, a
   caller's `contents`), `refresh()` switches to that live tree and the private copy is dropped;
   `activate_style(style)` raises while the styles part is private, since it would otherwise
   write into a tree nobody will save. This is the TypeScript port's departure 6, arrived at
   independently for the same reason.
5. **Three flags on catalogue entries record docx4j quirks rather than hiding them**:
   `counts_empty=False` for `w:bidiVisual` (applied by `apply(CTTblPrBase)` but not counted by
   `isEmpty(CTTblPrBase)`), and both that and a "keep the destination" merge for `w:tblCaption`
   and `w:tblDescription`, which docx4j's hand-written pair neither counts nor carries. Likewise
   `w:start` / `w:end` on `w:tblCellMar` and the `w:tcBorders` members follow the Java's lists,
   not the schema's, although the generated classes carry them.
6. **Table conditional formatting stays out**, as docx4j's own `table-conditions` CR has it.
   `apply_tbl_style_pr_list` merges a style's `w:tblStylePr` per condition, because
   `apply(Style, Style)` does; nothing applies a condition to a cell.

**Quirks reproduced deliberately**, each because a golden records it:

* **`w:framePr` gains an explicit `w:anchorLock`.** docx4j reads it through `isAnchorLock()`,
  which answers `true` when the attribute is absent, and then applies that answer, so any merged
  `w:framePr` carries `w:anchorLock="true"` even when no layer stated it. The same reading makes
  `isEmpty(CTFramePr)` treat an explicit `w:anchorLock="0"` as the only attribute that makes a
  frame non-empty on its own.
* **A `w:rFonts` that states only a `w:hint`, or nothing at all, still creates the element** in
  the destination, which `RunFontSelector` relies on; a *null* source leaves the destination
  alone, which is the CR-015 phase 1 fix.
* **`isEmpty(SectPr)` is "anything non-null is non-empty"**, which is all `is_empty_ppr` uses it
  for; `w:sectPr` is not merged (CR-015 phase 4).
* **`get_effective_ppr(styleId)` of a *character* style** answers the document defaults plus that
  style's (empty) paragraph chain, not None: the golden records it for every style in the part.
* **`get_effective_rpr(styleId)` of a missing style is None**, where `get_effective_ppr(styleId)`
  of a missing style is the default paragraph style's.
* **The heading outline level comes from the style's `w:name`**, not its declaration: Word takes
  a built-in heading's level from "heading 1" ... "heading 9", the same in every locale, whatever
  `w:outlineLvl` the style declares. The layer is a copy where the two differ, and the styles
  part is not touched.
* **Word's 10 pt default is made explicit** on the resolver's private copy of the document
  defaults where nothing states a size (CR-015's probe `styles-no-size-anywhere`), and **Word's
  built-in Normal Table** underlies a table naming no style and one whose chain reaches the
  document's default table style --- the built-in, not the document's definition of that style
  (CR-015 phase 4, probe `styles-table-default`).

**Goldens questioned: none.** Every recorded value is reproduced; nothing was left failing.

**The content API switch (CR-003 section 6, closed).** `Font` reads go through
`get_effective_rpr(rPr, pPr)` of the first run in scope, and `Paragraph.alignment`,
`left_indent`, `right_indent`, `first_line_indent`, `space_before`, `space_after`,
`line_spacing` and `outline_level` read `get_effective_ppr(pPr)`; `Table.style` and
`style_built_in` report the style the table resolves to, so a table naming none reads *Normal
Table* rather than nothing. Writes are unchanged and still direct, and so is the
`ChangeReport` a format write records --- a report is about what the *markup* says either side
of the write. `paragraph.get_font(direct=True)` and `range.get_font(direct=True)` are the direct
reads, which is what a caller reproducing markup wants and what these reads were before;
`paragraph.effective_p_pr` and `paragraph.effective_paragraph_mark_r_pr` are the resolved objects
themselves. A body with **no package** --- a fragment --- reads direct and raises nothing.
`Alignment` keeps `"Unknown"` where nothing anywhere states a `w:jc`, which is a departure from
the TypeScript port, whose `Alignment` reads `'Left'` after resolution: the enum is CR-003
section 3.2's and changing it is a content-API decision, not this CR's.

**The python-docx facade is deliberately not switched.** python-docx's `Font` booleans are
tri-state and `None` means *inherited*; python-docx does not report effective values either, and
CR-003 section 20.3 says the facade keeps python-docx's semantics. `docx4j_py.docx` therefore
still reads the run's own `w:rPr`, and the parity test with python-docx as the judge is
unaffected.

### 15.2 Step 2: numbering completion

Section 12.11 listed three things the Phase H emulator still owed Phase B. All three are here.

**Resolution through the resolver.** `Emulator.effective_numbering` was the one method in the
module that was not docx4j's own logic --- it walked `w:basedOn` itself, attribute by attribute,
over an lxml reading of the styles part --- and it is now
`PropertyResolver.get_effective_ppr(styleId).getNumPr()`, which is what docx4j's `Emulator.resolve`
asks for. `default_paragraph_style_id` and the `w:basedOn` walk inside `styleLinkedElsewhere` went
the same way. The lxml reading stays as the fallback for a body with **no package**, which is the
only caller that cannot ask a resolver, and it is still what `read_definitions` uses to resolve a
`w:numStyleLink` (docx4j's `resolveLinkedAbstractNum` reads the raw style too).

`resolve` gained the two reason strings the hand-written walk could not produce, in docx4j's own
order: `"style 'X' has a w:numPr without a w:numId val"`, which needs to tell a `w:numPr` that
states nothing from no `w:numPr` at all, and `"cyclic styles at X"`, which is reachable now that
a cycle can be detected. Neither appears in a golden; both are docx4j's contract.

**The twenty exotic formats.** `formats.py` is docx4j's whole registry --- **twenty-nine**
formats, one for one with its `NumberFormat*` class and its values: ordinal, cardinalText,
ordinalText, hex, chicago, numberInDash, the full-width, Thai and Devanagari digit families, the
Russian, Arabic and Thai alphabets, hebrew1 with its 15 and 16 exceptions, the two Chinese
counting sets and the enclosed circles, beside the eight section 6.2 named. Still fail-soft: a
`w:numFmt` with no formatter, or a value a formatter cannot express, gives the decimal label and
one warning per format; the formatters themselves raise `ValueError`, as docx4j's throw
`NumberFormatException`, so a caller that wants to know still can.

**`getInd`'s full form, the label's `w:rPr` and the level itself.** `ListLevel` gained the typed
views of docx4j's JAXB plane, unmarshalled once from the lxml element and kept: `lvl` (the
abstract `w:lvl`), `override_lvl`, `r_pr` (the abstract level's, never the override's),
`label_r_pr` (the override level's `w:rPr` where the instance overrides this level with one, else
the abstract level's --- a replacement, not a merge, which CR-014 probe P3 measured) and
`own_ind`. `Emulator.ind_of(numId, ilvl)` is `NumberingDefinitionsPart.getInd`: the override
level asked first and the abstract level only where it answers nothing, each through
`getIndFromLvl` --- a `w:ind` the level states itself, else the first found walking the linked
`w:pStyle`'s `w:basedOn` chain, read through the resolver's `get_style` as docx4j reads it.
`NumberingResult` carries all of them, and its `ind` is now the **typed** `w:ind` docx4j hands
out (the `Indent` summary, which folds `w:start` into `left`, is `result.indent`).

`get_ind(numPr)` is the seam the paragraph merge asks per layer, so that a style stating a
`w:numPr` brings the level's indent with it. That is what makes `effectivePPr` parity possible at
all: the TypeScript port did the same walk narrowly in its step 2 and replaced it in step 3; here
it was written once, in the real module, because section 12.11 had already put the definitions
there.

**Parity.** `tests/content/test_numbering_goldens.py` now runs **all 45** goldens rather than 19
--- 618 paragraphs over 88 stories, 183 of them numbered over 21 documents --- and compares
`numFont`, `ind`, `indResolved`, `lvl` and `labelRPr` as object trees beside the `numString`,
`isBullet`, `ilvl`, `numId`, `numRef` and whole `stateAfter` it compared before. **Zero
differences.** The walk it shares with the two other golden tests moved to
`tests/content/goldens.py`.

### 15.3 Step 3: fonts

`docx4j_py/model/fonts/` is the **decision** half of `org.docx4j.fonts`. docx4j's
`RunFontSelector` is 2,572 lines, most of them XSL-FO output; what section 6.3 promised, and what
a renderer, a text extractor or a measurer needs, is the decision per character.

**What is in.** `RunFontSelector` (`document_fonts_of`, `default_font_of`, the theme lookup by
`w:themeFontLang`, `resolved_slots`, `complex_script_font` with `w:cs` and `w:rtl` **by value**,
`preamble_rule`, the [MS-OI29500] 17.3.2.26 table branch for branch, `span_script`, `is_emoji`,
`symbol_font_name`, `run_locale`, `document_font_for` and `spans`); `theme_fonts.theme_font_of`,
docx4j's `ThemePart.getFont(STTheme, CTLanguage)` over an `a:fontScheme` rather than a part, since
the selector, font discovery and the content API all resolve references and none of them should
have to hold a `ThemePart`; `Mapper` with the one precedence of docx4j's CR-016 phase 3 and
`FontDecision(source, via, width_error, physical_font)`, which is what the goldens record;
`IdentityPlusMapper`; `FontFallback`'s class and candidate half; `fonts_in_use` and
`styles_in_use` as CR-016 phase 4's walk for names; and the default theme.

**Parity is zero differences** on all 45 goldens: **1,112 font spans** over 1,113 runs,
`fontsInUse` and `stylesInUse` as sorted lists, the theme part, the default font, and 134
`IdentityPlusMapper` decisions (41 distinct).

**docx4j's tables are committed as JSON**, under `docx4j_py/resources/fonts/`, written by
`scripts/generate_fonts.py` from docx4j's own resources in docx4j's own order --- 30 substitute
rows and 419 OpenOffice VCL entries from `font-substitutes.xml` and `FontSubstitutions.xml`, 137
Microsoft families (41 with a bold face of their own) and 512 line-metric families (46 East
Asian) from `MicrosoftFonts.xml` and `word-line-metrics.properties` --- beside the three Office
theme parts copied byte for byte. Each file carries a `_source` object naming the docx4j commit
it came from; the four data files are **byte identical** at the goldens' commit (`7fba7a150`) and
at docx4j's head, which was checked. The registry of 43 physical faces is the fourth file and is
not parsed out of the jars: reading a name table would need a font library this package does not
depend on, so the list is the script's own, with `--check-jars` to compare it against the built
jars.

**What of the Java is deliberately not ported**, and why each is a later CR:

1. **Everything that makes XSL-FO**: `createElement` / `setAttribute` / `symbolSetAttribute`, the
   `RunFontCharacterVisitor` and the three output modes, `finish()`'s post-processes (`kernSpaces`,
   `characterScaling`, `noLigatures`, `smallCaps`, `applyLineHeight`, `markWidthFactor`) and
   `getCssProperty`. The answer here is a list of `FontSpan`; how a span is drawn is the
   consumer's.
2. **The glyph-coverage pass** (`glyphFallback`, `FontFallback.selectCovering` / `covers` /
   `needsCoverage`, `GlyphCheck`, `GlyphAdvances`, `TextMeasurer`) and with it
   **`BestMatchingMapper`**'s panose matching: all of it reads a font file's cmap and metrics.
   `add_mapper_substitutes` is the empty hook docx4j's base class has, so the later CR adds a
   mapper rather than changing this one.
3. **`WordLineMetrics` and `WidthFactors`**: line boxes and advance corrections are measurement,
   not selection. `register_line_metrics_alias` / `line_metrics_family` are kept, because the
   `w:altName` and Word-default passes are written over them and a consumer that does measure
   needs to know whose metrics Word used, but nothing here reads a metric.
4. **`arabicNumbering` and `capsAndSoftHyphens`**: both rewrite the run's *text* before the
   dispatch, and neither changes which font a character gets.
5. **`symbolRun`'s Unicode replacement** (`SymbolMapper`, `translateUnicode2SingleByte`): the
   selector answers "Symbol" or "Wingdings" for such a run, which is the document font and what a
   golden records; turning a private-use code point into the replacement character a substitute
   face can draw is output.
6. **`registerUsedFont`, `ownFont`, the FOP configuration, `FontReport`, `FontsAnalysis.usage`,
   `MetricsOnlyFonts` and `FontEnvironment`**: no FOP here, and each of the last five needs font
   files.

**Departures from the Java, all deliberate.**

1. **A `FontRegistry` is a parameter, not a static map.** docx4j's `PhysicalFonts` discovers the
   machine's fonts once, in the `Mapper` static initialiser. A library that runs on a server has
   no reason to think the machine's fonts are the author's, and reading a font file needs a
   dependency this package does not have, so a `Mapper` is constructed over a registry the caller
   supplies. `default_font_registry()` is the 43 faces docx4j's four font jars carry, read off
   `PhysicalFonts.getPhysicalFonts()` with system discovery off --- the environment CR-016 phase
   0c measured and the goldens record --- so a mapping computed here reproduces a golden on any
   machine. `PhysicalFonts.discover` over installed fonts is a later CR, as section 6.3 says.
2. **`PhysicalFont` is a name, a family and the no-bold-face flag**, with no `EmbedFontInfo`, no
   panose, no file URI and no `Typeface`.
3. **`spans()` folds by font alone**, which is what `document_font_for` per code point answers and
   what the harness records. docx4j's rendering walk additionally cuts a span where the *script*
   changes between two non-shared characters, so that the coverage pass can substitute a Greek
   stretch as a sibling span rather than nesting it; that cut serves the coverage pass, which is
   not ported. `span_script` is exported and each span reports its script, so a renderer can cut
   further.
4. **U+2190 to U+2BFF answers the hAnsi font.** docx4j asks whether the hAnsi font has the glyph
   and names Segoe UI Symbol where it has that face. Both are glyph checks; without them the
   branch answers hAnsi, which is also docx4j's answer in the goldens' font environment.
5. **The emoji font is a constructor argument, not a property**
   (`docx4j.fonts.RunFontSelector.EmojiFont`); unset by default, as docx4j's property is.
6. **`script_of` is a range table, not the Unicode database.** Java's
   `Character.UnicodeScript.of(cp).name()` has no equivalent in Python: `re` has no
   `\p{Script=...}` and `unicodedata` exposes no script property, and adding `regex` for one
   field would be a dependency for nothing. So `scripts.py` carries the ranges of the scripts a
   font selection distinguishes, with the combining ranges answering `INHERITED` and everything
   else `COMMON` --- a code point of a script the table does not carry joins the span it follows,
   which is the safe answer and the same one a shared character gets. **Nothing in the goldens
   depends on it**: the harness records no script per span, and the field is an extension of the
   TypeScript port's that this reproduces.
7. **`docx4j.fonts.defaultTheme` is per package**, `pkg.default_theme`, not one process-wide
   property: docx4j has a properties file and a JVM; a server here may hold several packages at
   once. The process-wide `default_theme_setting()` is the fallback and is what moves them all.
8. **The mapper's per-script choices are not recorded**: `FontDecision` carries what a golden
   records and not `recordScriptChoice` / `recordSymbolFace` / `getBoldFace` / `getLineBox` /
   `getWidthFactor`, which are the conversion's report of what the coverage pass and the metrics
   did.
9. **`IdentityPlusMapper` is in `mapper.py`**, not a module of its own: it is twelve lines over
   `Mapper` and Python's import cost is per module, which is 12.3's reason for the one-module
   WordprocessingML parts.

**Quirks reproduced deliberately**, each because a golden or a CR-016 probe records it: `w:cs
w:val="0"` turns an inherited complex-script flag **off**, and `w:rtl` alone on Latin text still
takes the cs font; a theme reference beats the explicit attribute beside it at every level, and
where the package has no theme part the reference resolves to the Office theme's Latin face, so
the explicit name is never used; a theme reference the theme *part* cannot answer gives Calibri in
the document default only, where the same reference on a run resolves to nothing and the explicit
attribute stands; an `a:font` entry whose `typeface` is empty answers the empty string, where a
script the list does not carry falls back to the collection's own `a:latin`; a run with no
`w:rFonts` at all gets the default font in the ascii and hAnsi slots only, never the eastAsia
slot, which would fire the preamble rule; `w:cs=""` (LibreOffice writes it in `w:docDefaults`) is
read as no complex-script font in `resolved_slots` but is returned as it is by
`complex_script_font`; `is_known_family` asks `substitution_class`, not `class_of`, so a name that
merely ends in "Sans" is not a known family and the Word-default pass acts on it; and
`add_no_bold_face_aliases` skips a Word-defaulted font, because a font Word itself could not find
is substituted whole, its real bold included.

**`create_package` gives a new document a theme part**, as docx4j's `addDefaultThemePart` does,
so that the `minorHAnsi` its default styles name resolves to a face rather than to nothing.
docx4j's own `createPackage` does not, and this is the departure: a created document now has
**six** parts rather than five, and `create_package(default_theme="2013")` gives it Office 2013's
theme instead of Word 365's. A loaded document is untouched, as always, and one with no theme
part still resolves its references from the setting --- which is where the golden
`fonts-missing-slots` gets its Aptos.

**`Font.name` is the resolved face.** It was `w:rFonts/@w:ascii`, so a run whose face comes from
the theme read `""`; it is now `RunFontSelector.ascii_font_name` of the effective `w:rPr` --- the
face `w:asciiTheme` names, else `w:hAnsi`, else the document default. `get_font(direct=True).name`
is unchanged.

**Goldens questioned: none.** Every recorded value is reproduced; nothing was left failing.

### 15.4 The numbers

On this machine, in one session; `min` of three cold loads each.

| | 2010-sample1 (4 p, 2 runs) | invoice2013 (41 runs) | loadAndSave (10 stories) | Symbols (160 p, 3,902 runs) |
|---|---:|---:|---:|---:|
| build the resolver | 3 ms | 4 ms | 12 ms | 4 ms |
| build the selector | 3 ms | 3 ms | 3 ms | 3 ms |
| build the mapper (with `fonts_in_use`) | 2 ms | 7 ms | 15 ms | 232 ms |
| `fonts_in_use()` alone | 6 ms | 12 ms | 27 ms | 248 ms |
| `spans()` per run | --- | 38 us | 49 us | **41 us** |

Building a resolver is the styles part's parse and a copy of the document defaults: 3 to 12 ms,
once per package, and **no part is unmarshalled**, so the document still saves byte for byte.
`get_effective_rpr` is **34 us** a call over `Symbols.docx`'s 3,902 runs (132 ms for all of them)
and 40 us over `loadAndSave.docx`'s; the style chains are cached per style id, so what a call
pays is the document defaults' `deep_copy` and the run's own layer.

`fonts_in_use()` is the expensive one, and it is the walk, not the fonts: 248 ms over
`Symbols.docx` is one `iter_nodes` traversal of a 24,758-node part, 63 us per run, and it is paid
once per package when something first asks for a mapper. The mapper's own passes over the 10
names that walk found are the remaining 8 ms.

The three parity tests take **9.5 s**, **2.3 s** and **2.3 s**; the two unit files 1.1 s each.
The suite is **1,942 tests**, all passing, with one skip (the corpus round trip is `slow`).
`scripts/roundtrip.py` is 50/50 parts canonically identical, zero differences, nothing skipped.

### 15.5 What is left

* **`BestMatchingMapper`, the glyph-coverage pass and font metrics**, which are one body of work
  and one dependency: a library that can read a font file's cmap, panose and advances.
  `add_mapper_substitutes` is the hook, `FontRegistry` the seam and
  `register_line_metrics_alias` the record they plug into, so none of the code here changes.
* **Table conditional formatting**: `w:tblStylePr` is merged per condition but no condition is
  applied to a cell. docx4j's own `table-conditions` CR is where that is settled.
* **`PhysicalFonts.discover`** over installed fonts, which the same dependency unlocks.
* **The `<scriptSubstitutes>` and `<widthFactors>` blocks of `font-substitutes.xml`**, which serve
  the two passes above and are deliberately not generated.
* ~~**A Word check** of the five acceptance artefacts that gained a theme part (3, 4, 5, 6 and 11)~~
  --- **passed in Word on 2026-09-19**, the same day: no repair prompt, the body in Aptos and the
  headings in Aptos Display, as `tests/README.md` records. Nothing of Phase B awaits a check.
