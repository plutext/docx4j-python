# CR-031: A part not well formed is refused when read, not silently mended --- the fork's lxml handler without recovery, docx4j CR-065 D1's refusal

**Status:** Proposed 2026-11-03, at the user's request ("yes", to drafting CR-030 and CR-031). A part whose XML is not well formed raises `Docx4JException`, naming the part and the parser's reason, when its content is first read, as docx4j's does (CR-065 D1), where today the xsdata fork's lxml handler recovers, reads what it can and drops the rest with no report, even under `LoadOptions(strict=True)`; the load stays lazy, an untouched save still writes the part's bytes, and a relationships part not well formed refuses the load. Three phases (section 6): 0, the fork's stage 15 (`ParserConfig(recover=...)`), released as docx4j-xsdata 26.2.5; A, the engine; B, the records. 1½ to 2 days here and ½ day in the fork; eleven questions, each with a recommendation (section 8). **Decided 2026-11-03**, Q1 to Q11 as recommended (the user: "CR-030 and CR-031 both decided as recommended"; section 8): the refusal through the fork's stage 15, a relationships part not well formed refusing the load and docx4j's reading of one as none handed to docx4j, the text model reading without an unreadable styles, numbering or settings part, marked by `part.not_read`, the failure recorded on the part, `part.not_well_formed` in the validator, no mended reading on request, `Docx4JException` with no new class, a generic part refused when read, the custom XML view left as it is, and no Word check. Phase 0 done 2026-11-03 (the fork's stage 15, `77398949`; docx4j-xsdata 26.2.5 on PyPI; `docs/UPSTREAM.md` row 29, `1fc04d3`). **Phase A implemented 2026-11-03** (section 10: `0f73213`, `57c0b94`): `runtime.RECOVER` False and the pin at 26.2.5; every read of a part not well formed refused with `Docx4JException` naming it and lxml's reason, the `mc` passes' too; `XmlPart.unreadable` and `is_content_unreadable`; a relationships part refusing the load; `part.not_well_formed` in the validator; `source_tree()` None; `part.not_read` held to docx4j's `part-not-read` fixture at `3ad164bdf`, its case-mapping lines pending until CR-029; 199 tests, full suite 9,495 passed. **Amended 2026-11-03** (the user: "part.not_read as recommended", WD 0.66 REQ-801): where the text model needed such a part and could not read it, it warns `part.not_read`, in place of a logged line (section 4.6; Phase A ¼ day more). **Amended again 2026-11-03** (the user: "make the list readers refuse; start phase B"; section 10.8 item 1): the content API's list readers refuse a numbering or styles part that is not well formed, as section 4.6 said, where Phase A's list emulator read it as absent; Phase A2, ¼ day, then Phase B. **Phase A2 implemented 2026-11-03** (section 11: `f1cfa01`): every reader over the list emulator raises the part's own `Docx4JException`, its `unreadable`, the styles part at the first read of the styles and the numbering part only over a paragraph that names a list (`Emulator.names_a_list`), the property resolver's readers there too (section 11.3); the text model unchanged, its emulator reading no styles it could not read and a paragraph meeting the numbering part's refusal resolved without it; the other raw readers measured and left as they are, the accept and reject engine's reading of an unreadable story part as empty put to the user (section 11.4); 294 tests, full suite 9,789 passed. **Phase B done 2026-11-03** (section 11.7: CHANGELOG, CLAUDE.md, `tests/README.md`, the index; the hand-offs to docx4j and docx4j-core-ts sent through the portfolio and answered, section 11.8). **CR-031 complete** (no Word check, Q11).

*The number:* CR-030 is the latest CR here, drafted beside this one the same day.

**Depends on:**
- **docx4j CR-065 D1** (accepted 2026-10-26; phase 1 `9ea65238a`): a part whose content cannot be read is never
  mended or replaced, the load returns, an untouched save keeps its bytes. Its section 9 (`4fdd80063`, 17.3.6,
  unreleased) is the text model's message, which CR-027 Phase A ported here (`c7d4c81`).
- **The xsdata fork**, for phase 0: stage 15 (item 33, the next free numbers at fork `cc59fab6`), released as
  docx4j-xsdata 26.2.5. Proposed registry ids `xsdata/stage-15` and `xsdata/release-26.2.5`, as for stage 14. The
  release is the user's to publish (`docs/fork/REBASING.md`, "Releasing"). Phase A's refusal waits for it; the rest of
  Phase A does not (section 6).
- **This repository:**
  - CR-002's lazy load and byte-for-byte rule (`XmlPart`, `save.py`);
  - CR-027 Phase A's unreadable story parts (`trees._read`, `not_readable`, `content.not_read`; its section 4.6);
  - CR-020 Phase B's validator (`misplaced()`, `save(check=True)`, `source_tree()`);
  - CR-012's `splice_word` and the MCE preprocessor, which parse with lxml before the fork does;
  - CR-030 (proposed today) puts its `w:br` pass in `XmlPart._parse` after the fork's parse. The two compose
    (section 4.4).
- **The WD's REQ-801 at 0.66** (portfolio `bfcd03a`, the user's ruling of 2026-11-03): `part.not_read` (section 4.6).
- **Counterparts:** docx4j as above; the portfolio sent `part.not_read`'s hand-offs to docx4j and core-ts with 0.66. docx4j-core-ts and docx4j-generated-objects-ts are not measured here (section 7).

## 1. Why

CLAUDE.md's rule is "**Nothing may be dropped silently.** Loading is lenient, with a per-part report". The fork's
`LxmlEventHandler.parse` (`docx4j_xsdata/formats/dataclass/parsers/handlers/lxml.py`, line 37) parses every typed
part with `etree.iterparse(..., recover=True)`, upstream xsdata's setting. libxml2's recovery mode reads past a
well-formedness error. It truncates at a cut, drops an undefined entity or a stray `&`, closes elements where it
guesses they end, and keeps the first of two roots. The binder then sees a tree that is whole as far as it can tell,
so nothing goes in `part.skipped`, and `LoadOptions(strict=True)` raises only where the guess happens to leave content
the model cannot place. CR-027 section 10.11 item 1 found it. Section 2 measures it:
- a `word/document.xml` cut after "SEC" of its second paragraph reads `FIRST PARAGRAPH`, `SEC`. The rest of the body
  is gone, its section properties and header reference with it, at the default and under `strict=True`, with an
  empty report;
- an edit, or only a read through the content API (`pkg.body.text`, which unmarshals the part), followed by a save
  writes the mended part, so the file loses what the read lost;
- a `word/_rels/document.xml.rels` cut after its third relationship loads without the header and numbering parts the
  last two named. An untouched save, with nothing read, then writes the package without them.

docx4j refuses such a part (CR-065 D1, and section 3 here): the load returns, the part's content is unreadable with
the parser's reason, every read says so, and an untouched save keeps the bytes. XML 1.0 (Fifth Edition) section 1.2
makes a well-formedness violation a *fatal error*, after which a conforming processor "MUST NOT continue normal
processing (i.e., it MUST NOT continue to pass character data and information about the document's logical
structure to the application in the normal way)". libxml2's recovery mode is outside that by design.

## 2. What exists here (measured 2026-11-03 at `9d869f4`)

CPython 3.14.6, lxml 6.1.3 (libxml2 2.14.6), the fork editable at `cc59fab6` (26.2.4's code). HEAD has since moved
to `7fe0516` (CR-030's document only), so the engine measured is the same. The probes are in this session's
scratchpad (`cr031/`: `probe.py`, `side.py`, `corpus.py`, `timing.py`, `routes.py`, `rss.py`,
`norecover_plugin.py`, `java/Cr031Probe.java`), not in the repository. "Recovery off" is the fork's handler
monkeypatched in the probe's process to pass `recover=False`, and nothing else; neither the fork nor the engine was
edited.

### 2.1 Where a part's XML is parsed

| Route | Parser | Today, for XML not well formed |
|---|---|---|
| The fork's `LxmlEventHandler.parse`: every typed part through `XmlPart._parse` (`contents`, `read_contents()`, `source_tree()`), so every story, styles, numbering, settings and docProps part and every `RelationshipsPart` | `iterparse(recover=True)` | **mended**, no report |
| `xml_part._root_ns_map` (the root's declarations, before the parse) | `iterparse(recover=True)`, the root's start event only | a sniff; `{}` on a raise |
| `CustomXmlDataStoragePropertiesPart.item_id` (`parts/docprops.py`, at load) | `iterparse(recover=True)`, the root's start event only | a sniff; None on a raise |
| `DefaultXmlPart.tree` (custom XML, `VMLPart`, any generic XML part) | `XMLParser()`, strict | raises lxml's `XMLSyntaxError`, which doesn't name the part |
| `mce.mce_preprocess_bytes` (`LoadOptions(mce_preprocess=True)`), run in `_parse` before its `try` | strict | raises lxml's `XMLSyntaxError` out of `contents` |
| `mc_load_mode.splice_word` (`mc_load_mode="word"`), the same, only for bytes holding `AlternateContent` | strict | the same; bytes without the marker go to the fork and are mended |
| `ContentTypeManager.parse` | strict | `InvalidFormatException("Bad [Content_Types].xml: ...")`, the load refused |
| `FlatOpcStore.parse`, `fragments.wml(...)`, `CustomXmlPart.set_xml` | strict | `Docx4JException`, `FragmentError`, `BindingError` (`binding.not_well_formed`) |
| The model's own readers of a part's bytes: `comments._root_for_read`, `load._item_id`'s fallback, `content/styles.py:522`, `reports.py:772`, `describe.py:221`, `listnumbering/emulator.py:93`, `customxml/template.py:449`, `text/bound_values.py:92`, `text/metadata.py:152`, `customxml/parts.py:667`, `revisions/removals.py:218`, `chart/model.py:38` | strict | each its own handling; none mends |

So recovery is in exactly three places: the fork's handler, which reads content, and two sniffs of a root's start tag.

### 2.2 Twenty malformations, today

The base document is made in code (`create_package`) and saved. It has four paragraphs: `FIRST PARAGRAPH`, `SECOND
PARAGRAPH` numbered from docx4j's default numbering, `HIDDEN BY STYLE` in a paragraph style with `w:vanish`, and
`THIRD PARAGRAPH`. It has a header reading `HEADER TEXT`. Each case changes one part, and each route starts from a
fresh load:
- an untouched save;
- `read_contents()` and `part.skipped`;
- `extract()` with headers;
- `get_property_resolver()`;
- `pkg.body.text`, then a save;
- an edit (a paragraph `NEW` appended to the broken part, or to the main part for a relationships part), then a save,
  reloaded.

The default and `strict=True` give the same in every case but the three noted.

| # | Case | What a read gives today | The report |
|---|---|---|---|
| 1 | `document.xml` cut after "SEC" | `FIRST PARAGRAPH`, `SEC`; the text model the same, and no header (the cut took `w:sectPr` and its reference) | none |
| 2 | cut mid-tag (`<w:pPr><w`) | `FIRST PARAGRAPH`, `SECOND PARAGRAPH`, `""`; the stub `<w` bound as an element `w` | 1 skipped (element `w` in `PPr`); strict raises "1 item(s) were skipped" |
| 3 | an undefined entity (`SECOND&nbsp;PARAGRAPH`) | `SECONDPARAGRAPH` | none |
| 4 | an element left open (`<w:spacing w:before="120">`) | `""`: the rest of the body nested inside `w:spacing` | 5 skipped (a `w:r` in `PPr`, three `w:p` and the `w:sectPr` in `P`); strict raises "5 item(s) were skipped" |
| 5 | `</w:body>` removed | all four paragraphs | none |
| 6 | `</w:body` (docx4j's break, CR-027 section 10.8) | all four paragraphs | none |
| 7 | a stray `&` | `SECOND  PARAGRAPH` | none |
| 8 | a stray `<` | `FIRST PARAGRAPH`, `SECOND `: everything after it gone | none |
| 9 | an attribute given twice (`w:before="120" w:before="240"`) | all four; the first value kept | 1 skipped (attribute `w:before`); strict raises "1 item(s) were skipped" |
| 10 | a control character (U+0001) | `SECONDPARAGRAPH` | none |
| 11 | a byte not UTF-8 (a lone `0xE9`) | `SECOND � PARAGRAPH`: U+FFFD, written by any later save | none |
| 12 | a second root after the first | all four; the second root (`EXTRA`) gone | none |
| 13 | `<<w:document` (CR-027's break) | refused: ``Problem with part /word/document.xml: Failed to create target class `Document` `` | - |
| 14 | `header1.xml` cut after "HEA" | `HEA`; the text model `HEA` too, with no warning | none |
| 15 | `styles.xml` cut inside `HiddenStyle`'s start tag | 23 styles, the last with no id; the resolver built; `HIDDEN BY STYLE` read as text in both the text model and `body.text` | none |
| 16 | `numbering.xml` cut before its first `w:num` | 2 `w:abstractNum`, no `w:num`: the list label gone (the Markdown loses `1. `) | none |
| 17 | `document.xml.rels` cut after its third relationship | the load succeeds without `/word/numbering.xml` and `/word/header1.xml`; **an untouched save writes the package without them** | none |
| 18 | `_rels/.rels` cut inside its first relationship | the load refused: `InvalidFormatException("No relationship of type officeDocument")` | - |
| 19 | `_rels/.rels` cut inside its second | the load refused: `InvalidFormatException("A part name shall not be empty [M1.1]")` | - |
| 20 | `[Content_Types].xml` cut | the load refused: `InvalidFormatException("Bad [Content_Types].xml: AttValue: ' expected, line 2, column 251 ...")` | - |

The routes, across the cases:
- **An untouched save is byte for byte** for every part in cases 1 to 16. Case 17 is the exception: the
  relationships part is read at load, so the save re-marshals it.
- **A read through the content API marks the main part**, so the save after `pkg.body.text` rewrites
  `document.xml` from the mended tree in cases 1 to 12: case 1's file then holds two paragraphs and no section
  properties.
- **An edit saves the mended part:** case 1 saves `FIRST PARAGRAPH`, `SEC`, `NEW`, and case 4 saves `""`, `NEW`.
- **`strict=True` raises only in cases 2, 4 and 9**, where the mended tree has content the model cannot place. Its
  reason is that content, not the XML.
- **The python-docx facade:** `docx4j_py.docx.Document(case 1)` gives the paragraphs `FIRST PARAGRAPH`, `SEC`.

### 2.3 The same with recovery off

Every read of the broken part is refused, at the default and under `strict=True` alike. Each read raises
`Docx4JException` "Problem with part <name>: <lxml's message>", the same message as for any other parse failure. An
untouched save stays byte for byte. lxml's messages:

| # | lxml's message (each followed by `(<string>, line 2)`) |
|---|---|
| 1 | `Premature end of data in tag t line 2, line 2, column 351` |
| 2 | `Couldn't find end of Start Tag w, line 2, column 396` |
| 3 | `Entity 'nbsp' not defined, line 2, column 360` |
| 4 | `Opening and ending tag mismatch: spacing line 2 and pPr, line 2, column 216` |
| 5 | `Opening and ending tag mismatch: body line 2 and document, line 2, column 719` |
| 6 | `expected '>', line 2, column 714` |
| 7 | `xmlParseEntityRef: no name, line 2, column 356` |
| 8 | `StartTag: invalid element name, line 2, column 356` |
| 9 | `Attribute w:before redefined, line 2, column 222` |
| 10 | `PCDATA invalid Char value 1, line 2, column 354` |
| 11 | `Invalid bytes in character encoding, line 2, column 355` |
| 12 | `Extra content at the end of the document, line 2, column 728` |
| 13 | `StartTag: invalid element name, line 2, column 2` |
| 14 | `Premature end of data in tag t line 2, line 2, column 97` |
| 15 | `Couldn't find end of Start Tag style, line 2, column 11071` |
| 16 | `Premature end of data in tag numbering line 2, line 2, column 4866` |
| 17 | the load raises: `Problem with part /word/_rels/document.xml.rels: Premature end of data in tag Relationships line 2, line 2, column 475` |
| 18 | the load raises: `Problem with part /_rels/.rels: AttValue: ' expected, line 2, column 115` |
| 19 | the load raises: `Problem with part /_rels/.rels: Specification mandates value for attribute Target, line 2, column 247` |
| 20 | unchanged (the content types were strict already) |

Downstream of the refusal, with recovery off:
- **The text model**, which CR-027 Phase A already handles:
  - cases 1 to 13 raise "The content of /word/document.xml could not be read, so it cannot be read as text: <lxml's
    message>";
  - case 14 reads the body and warns `content.not_read at header:rId5 (/word/header1.xml)`;
  - case 15 reads without styles, so `HIDDEN BY STYLE` is read as text, with no warning and no log;
  - case 16 reads without numbering, with no list label, no warning and no log (the output is today's).
- **`get_property_resolver()`** raises the styles part's exception in case 15, and so does `pkg.body.text`. In case
  16 `body.text` reads.
- **The fork's exception chain:** `XmlParser.parse` raises `ParserError(e)` inside its `except SyntaxError`, so
  lxml's `XMLSyntaxError` (with `lineno` and `position`) is its `__context__`, not its `__cause__`.

### 2.4 The other routes

- **The MCE preprocessor and Word mode** already refuse, but with the wrong exception. Case 1 under
  `LoadOptions(mce_preprocess=True)` raises lxml's own `XMLSyntaxError` out of `contents`, naming no part. Under
  `mc_load_mode="word"`, case 1 is mended (no `AlternateContent` in the bytes, so the splice never parses) and the
  same cut in a body that holds a block-level `mc:AlternateContent` raises lxml's `XMLSyntaxError`. The outcome
  depends on the content. `tests/openpackaging/test_mce_levels.py::test_a_corrupt_part_still_raises_in_word_mode`
  accepts either exception.
- **A custom XML part not well formed** (its `customXml/item1.xml` cut) loads, and registers by its properties part's
  id. Its `.tree` raises lxml's `XMLSyntaxError`. The content API's view swallows that: `document_element` is None
  and `get_xml()` returns `""`. Over `samples/invoice2013.docx` with its custom XML part cut in half,
  `apply_bindings()` applies 0 of its 20 bindings and changes nothing in the body. An untouched save is byte for
  byte.
- **The custom XML properties part cut** after its root's start tag still gives its `ds:itemID` through the
  recovering sniff.

### 2.5 The corpus

Every XML part (content type ending `xml`, and every `.rels`), parsed as the fork does but with `recover=False`, and
one level into each embedded package:

| Where | Packages | XML parts | Not well formed |
|---|---|---|---|
| `samples/` | 17 (1 nested) | 282, its one VML part included | 0 |
| `tests/fixtures/` | 424 (25 nested), and 1 loose XML file | 5,736 | 0 |
| the rest of `tests/` (`tests/office_js/` and the like) | 63 | 1,108 | 0 |

The 9 `text/html` parts (altChunk sources and `hw.html`) are not XML parts and no route parses them as XML. Five of
them are not well formed as XML, which is expected of HTML.

### 2.6 The suite with recovery off

The fork's handler patched in the test process (a scratchpad plugin, `-p norecover_plugin`):
- the fast suite: **3 failed, 7,101 passed**, 8 skipped, 1,699 deselected (220.6 s);
- the slow tests: 1,699 passed, 1 skipped (171.6 s), the corpus round trip included.

The three failures:
- `tests/content/test_integer_normal_form.py::test_an_empty_rsid_is_dropped_where_nothing_was_kept`. Its fixture
  inserts `w:rsidR=""` into the first `<w:p `, which already has a `w:rsidR`: "Attribute w:rsidR redefined". It
  passed because recovery kept the first value. The fixture is the bug: it should insert into a paragraph with no
  rsid, or replace the existing value.
- `tests/text/test_unreadable_parts.py::test_the_text_model_says_which_part_could_not_be_read_and_why` and
  `::test_the_text_model_of_a_header_says_which_part_could_not_be_read`. Each asserts the reason
  "Failed to create target class `Document`" (or `Hdr`), the reason recovery gave for `<<w:document`. lxml's reason
  replaces it.

### 2.7 The cost

`samples/Symbols.docx`'s main part (790,857 bytes), best of nine, interleaved, three runs:

| | `recover=True` | `recover=False` |
|---|---|---|
| lxml's `iterparse` alone, events consumed | 13.67 to 13.76 ms | 13.53 to 13.65 ms |
| the engine's parse (`runtime.parser().from_bytes(data, Document)`) | 196.4 to 199.2 ms | 196.2 to 197.9 ms |

Nothing measurable. Section 4.1 uses three more measures:
- a strict lxml tree handed to the fork (its `iterwalk` path) is 186.5 ms best, but its peak memory is 13.4 MiB
  above the start where the streaming parse's is 4.5 to 4.7 MiB (two runs each, `ru_maxrss`);
- the two routes give identical trees, kept and skipped reports and marshalled bytes for all 651 typed parts of
  `samples/` and `tests/fixtures/*.docx`;
- a strict lxml pre-pass alone costs 12.9 ms, 6.5% of the parse.

## 3. The counterpart: docx4j

### 3.1 How docx4j reads a part

Read in docx4j at `ac66f5518` (`VERSION_17_3_6`).
- **`JaxbXmlPart.unmarshal` and `JaxbXmlPartXPathAware.unmarshal`** (the main part, headers, styles, numbering and
  the rest) unmarshal through StAX (Woodstox). On a failure they may retry through the preprocessor, which parses the
  bytes with Xerces' `DocumentBuilder`. Neither parser recovers. Xerces' `SAXParseException` becomes a
  `JAXBException`, and `JaxbXmlPart.getContents()` keeps it as `unreadable`, a
  `Docx4JException("Problem with part /word/document.xml", cause)`. Every later read throws it again without
  another parse.
- **`getJaxbElement()`** logs it once and returns null. `isContentUnreadable()` tells "could not be read" from "no
  content yet". The conveniences refuse rather than replace (`contentsOrNew`, CR-065 phase 1).
- **`getReadableContents(purpose)`** gives the message the text model throws (CR-065 section 9).
- **An untouched save** writes the part's bytes from the source store.
- **A part held as DOM** (`XmlPart`, a custom XML part) is parsed eagerly by `Load3.getPart`, so one not well formed
  fails the load: "Failed to getPart", its root a `SAXParseException` (measured: section 2.4's cut item).
- **`RelationshipsPart.unmarshal(InputStream)`** catches every exception, prints its stack trace and keeps the
  `Relationships` it had (section 3.4).
- **`[Content_Types].xml`** refuses the load: "Couldn't get [Content_Types].xml from ZipFile".

### 3.2 docx4j over the same twenty files

`java/Cr031Probe.java`, OpenJDK 21.0.12. The classpath is `docx4j-core-17.3.6-SNAPSHOT.jar` as built in docx4j's
checkout on 2026-11-03 (it has `4fdd80063`'s text-model refusal), ahead of the 17.3.6-SNAPSHOT bundle (2026-11-01)
for its dependencies.

| Cases | docx4j |
|---|---|
| 1 to 13 (`document.xml`) | `load()` returns. `getContents()` throws `Problem with part /word/document.xml` and `isContentUnreadable()` is true. `Texts.of(pkg)` throws `Docx4JRuntimeException` "The content of /word/document.xml could not be read, so it cannot be read as text: <Xerces's message>": case 1 "XML document structures must start and end within the same entity.", case 3 "The entity "nbsp" was referenced, but not declared.", case 9 "Attribute "before" bound to namespace ... was already specified for element "w:spacing"", case 11 "Invalid byte 2 of 3-byte UTF-8 sequence." (a `MalformedByteSequenceException`). An untouched save is byte for byte |
| 14 (header) | the header is unreadable. `Texts.of` reads the body and warns `content.not_read at header:rId5 (/word/header1.xml)`. Byte for byte |
| 15 (styles) | the styles part is unreadable. `Texts.of` reads without styles (`HIDDEN BY STYLE` is text), with no warning. `getPropertyResolver()` throws a `NullPointerException` ("Cannot invoke ... Styles.getStyle() because the return value of ... StyleDefinitionsPart.getJaxbElement() is null") |
| 16 (numbering) | the numbering part is unreadable. `Texts.of` gives `SECOND PARAGRAPH` numbered with no format, with no warning |
| 17 (`document.xml.rels`) | **`load()` succeeds with no relationships for the main part.** Its text has no header, shows `HIDDEN BY STYLE` and has no list format. An untouched save loses `word/_rels/document.xml.rels`, `header1.xml`, `numbering.xml`, `settings.xml`, `styles.xml` and `theme/theme1.xml` |
| 18, 19 (`_rels/.rels`) | the load refused: "Error getting document from Zipped Part: _rels/.rels", its root a `NullPointerException` |
| 20 (content types) | the load refused: "Couldn't get [Content_Types].xml from ZipFile", its root a `WstxEOFException` |

So, except for a part's relationships, docx4j's behaviour is this CR's design (section 4).

### 3.3 Styles and numbering in docx4j

docx4j has no deliberate rule for a styles or numbering part it cannot read:
- `PropertyResolver`'s constructor throws a `NullPointerException` from `getDefaultTableStyle()`, which calls
  `getJaxbElement().getStyle()`;
- the text model's `Extractor` catches any exception building the resolver, logs WARN "No property resolver: hidden
  text and indents are not resolved through styles (...)" and reads without one;
- `Extractor.numbering` catches a `RuntimeException` per paragraph and logs it at DEBUG.

Neither is a `TextWarning`. REQ-801 names the story parts only.

### 3.4 Found: docx4j loads a part's relationships that are not well formed as none

`Load3.getRelationshipsPartFromZip` makes the part's `RelationshipsPart` (`getRelationshipsPart(true)`, empty), then
`RelationshipsPart.unmarshal(is)` swallows the parse failure. The part then has no relationships, every part those
named is never loaded, and an untouched save drops them all (case 17: six entries). For the package's own `.rels`,
the null content fails later as an NPE, so that load is refused. CR-065 D1's principle doesn't reach this method. It
is a hand-off to docx4j, through the portfolio session, with case 17's file as the reproduction (Q3).

### 3.5 Word

Not measured: no probe of the shared set has a part that is not well formed. Word's reading isn't this CR's oracle.
The rule is XML 1.0's and docx4j's, and nothing here changes what a save of a well-formed file writes (Q11).

### 3.6 xsdata upstream

Upstream `xsdata 26.2` (`.venv`) over `<a><b>one</b><b>two</b><b>thr`, cut:
- the lxml handler gives `A(b=['one', 'two', 'thr'])`;
- the native handler (`XmlEventHandler`) raises `ParserError: no element found: line 1, column 29`;
- the lxml handler with `process_xinclude=True`, which parses with `etree.parse` and the default parser, raises
  `ParserError: Premature end of data in tag b ...`.

So upstream's handlers disagree, and only the default one mends. It is upstream row 29 for `docs/UPSTREAM.md`,
"to report", with the fork's option as the patch to offer.

## 4. Design

### 4.1 Where: the fork's stage 15, set by the engine (D1, Q2)

**The fork** (stage 15, item 33, in `parsers/config.py`, `parsers/handlers/lxml.py` and `parsers/bases.py`):
- **The option.** `ParserConfig(recover: bool = True)`: whether the lxml handler recovers from XML that is not well
  formed. True is upstream's behaviour, so with the option at its default every behaviour is 26.2.4's, as for stages
  13 and 14.
- **The handler.** With `recover=False`, `LxmlEventHandler.parse` passes `recover=False` to `etree.iterparse`. The
  first fatal error raises `XMLSyntaxError`, which `XmlParser.parse` raises as `ParserError`. The objects built
  before it are discarded, and the skipped report is the caller's to ignore.
- **The cause.** `raise ParserError(e) from e`, so lxml's error, with `lineno` and `position`, is the `__cause__`.
  Upstream has no `from`, so it is the `__context__` today (section 2.3).
- **Unaffected:** the native handler, which never recovers; the `process_xinclude` path, already strict; and a tree
  source (`iterwalk`), which is already parsed.
- **Tests:** `tests/fork/test_recover.py`:
  - the default mends as before (a pin);
  - `recover=False` refuses each malformation of section 2.2 under the lxml handler, with lxml's message, the
    `XMLSyntaxError` as the cause;
  - a well-formed document parses identically under both settings, each `skipped` and `kept` record included;
  - a tree source is unaffected.
- **The release:** `CHANGES.md` stage 15 and release 26.2.5, by `REBASING.md` "Releasing" steps 1 to 5. Step 1's gate
  is docx4j-python against the wheel, with the option set as phase A sets it. The user publishes.

**The engine:**
- `runtime.parser_config` passes `recover=False` (`runtime.RECOVER`, beside `EXACT_TYPES`);
- `tests/test_parser_configs.py`, which holds every `ParserConfig` in the tree to `parser_config()`'s, gains the
  assertion;
- the pin moves to `docx4j-xsdata[lxml]>=26.2.5,<26.3`.

**The alternatives** (Q2), each measured in section 2.7:
- **(b) Engine only, a strict tree:** `XmlPart._parse` parses the bytes with a strict `etree.XMLParser` and hands
  the tree to the fork, whose `iterwalk` path never recovers. No fork release is needed, and `splice_word` and the MCE
  pass could share the tree instead of re-serialising it. It is 5% faster on `Symbols.docx` and identical over the
  651 parts, but holds the whole lxml tree at once: about 9 MiB more peak for a 790 KB part, about 11 times the part.
- **(c) Engine only, a handler subclass:** `UncheckedXmlParser(handler=...)` with a copy of the fork's `parse`. It
  costs nothing and needs no release, but copies the fork's internals (`process_context`, `EVENTS`) into the engine,
  which CLAUDE.md's rule ("a generator or runtime bug is fixed in the fork") avoids.
- **(d) A strict pre-pass** before the fork's parse: 6.5% on every parse of every part.

(a), the fork's option, costs nothing, keeps the parse streaming, and is the change to offer upstream.

### 4.2 What is refused, when, and how (D2, Q8)

- **When:** at the first read of the part's content: `contents`, `read_contents()` and `source_tree()` for a typed
  part, and `tree` for a `DefaultXmlPart`. Never at load, except for the relationships (section 4.5). A part is never
  read from partial content.
- **The exception:** `Docx4JException`, docx4j's, as for any part whose parse fails today. No new class (Q8).
- **The message:** today's form, `Problem with part /word/document.xml: Premature end of data in tag t line 2,
  line 2, column 351`. It is lxml's `msg`, without the `(<string>, line 2)` lxml appends for a source with no file
  name. lxml's `XMLSyntaxError` is the deepest `__cause__`, through the fork's `ParserError`, so a caller that must
  tell this failure from another can.
- **`_parse`'s other two parsers.** The splice (`splice_word`) and the MCE pass (`mce_preprocess_bytes`) move
  inside `_parse`'s `try`. Every route then raises the same `Docx4JException` (section 2.4), and Word mode refuses a
  part whether it holds `mc:AlternateContent` or not.
- **`DefaultXmlPart.tree`** raises the same, naming its part, from lxml's error.

### 4.3 The load and the save (D3, Q9)

- **The load stays lazy:** it parses the content types and the relationships, and the two sniffs read a root's start
  tag (section 4.4). A part that is not well formed and never read costs nothing.
- **A save writes such a part's bytes as read:** a part never read, a part whose read failed (its `_contents` stays
  None) and a part beside an edited one alike. That is CR-002's rule and docx4j's. The file's own fault is kept as it
  was. A part no longer comes back from a save mended and shorter.
- **A replacement** (`set_contents`, `set_bytes`, `set_tree`) is the caller's decision, as docx4j's
  `setJaxbElement` is.
- **A departure from docx4j:** a generic XML part (custom XML, VML) is parsed lazily here, on `tree`, and docx4j's
  eager parse fails the load for one not well formed (section 3.1). The engine keeps its lazy rule (Q9). The custom
  XML parts are indexed at load by their properties parts' `ds:itemID`, which the sniff reads without the item part.

### 4.4 The engine's own parsers (D4)

- **The two sniffs** (`_root_ns_map`, `item_id`) parse with `recover=False`. lxml yields the root's start event before
  an error later in the bytes (measured), so the result is the same for any part whose root start tag is itself well
  formed. A broken root gives `{}` and None, as a raise does today. After this, no parser in the tree recovers.
- **A census** holds it: a test over `docx4j_py/`'s source finds no `recover=True` (lxml's default is strict), and
  `parser_config().recover` is False.
- **CR-030's pass** (`breaks.repair_bytes`) runs only after the fork's parse succeeds, so on bytes already well
  formed. Its lxml parser must not recover either, and the census holds it to that.
- **The model's own raw readers** (section 2.1's last row) are strict already, and each keeps its handling, except
  the list emulator's, which refuses (section 4.6, amended 2026-11-03).

### 4.5 Relationships and content types (D5, Q3)

- **A relationships part that is not well formed refuses the load**, the package's or a part's:
  `Docx4JException("Problem with part /word/_rels/document.xml.rels: ...")` from `load()`. Without its relationships
  the load cannot know which parts exist. Reading the relationships before the error and dropping the rest is
  today's silent loss (case 17), and reading none is docx4j's (section 3.4). With recovery off this is what the
  engine already does (section 2.3), because `rels.list` reads the part at load. So D5 is a test and a sentence in
  `load()`'s docstring.
- **`[Content_Types].xml`** is refused as today (`InvalidFormatException`, "Bad [Content_Types].xml: ...").

### 4.6 Styles, numbering and settings (D6, Q4; amended 2026-11-03)

- **The content API refuses.** `get_property_resolver()` raises the styles part's `Docx4JException` (docx4j: an
  NPE), and so does every content API reader that resolves styles (`body.text` and the rest).
- **The list readers refuse too** (amended 2026-11-03; the user: "make the list readers refuse"; section 10.8 item
  1). `Paragraph.is_list_item`, `list_item` and every other reader over the list emulator read the numbering and
  styles parts with lxml (`listnumbering/emulator.py`'s `_root_of`, under `_ensure` and `_read_styles`). Phase A kept
  that reader by section 4.4, and it reads a part not well formed as absent: a numbered paragraph is no list item,
  with nothing said. Instead they raise the part's own refusal: the `Docx4JException` its `contents` raises
  ("Problem with part /word/numbering.xml: <lxml's message>", lxml's `XMLSyntaxError` the deepest cause), the same
  text, and recorded as its `unreadable` (D8).
  - **A styles part:** at the emulator's first read of the styles, so at the first list reader, as the resolver's
    readers refuse.
  - **A numbering part:** only where a reader needs a definition, at the same paragraphs as the text model's
    `part.not_read` (a `w:numPr`, direct or through a style; section 10.3). A list reader over a paragraph with no
    numbering raises nothing for the numbering part.
  - **The writers** refuse already: they read `contents`.
  - **The text model is unchanged.** It hands the emulator no part its consult recorded as not read, so it still
    reads as for a package that lacks the part, with one `part.not_read`, held by docx4j's fixture.
  - **The other raw readers** of section 2.1's last row keep their handling (section 4.4). Phase A2 measures how each
    reads a part not well formed and records it, unchanged.
- **The text model reads as for a package that lacks the part, and says so** with `part.not_read` (WD 0.66 REQ-801;
  the user, 2026-11-03: "part.not_read as recommended"). Q4 first asked the WD, and this section first proposed a
  logged line and no `TextWarning`.

**Measured 2026-11-03** at `9825b87`, with the fork's handler patched to `recover=False` as in section 2.6. The
document is made in code:
- a heading numbered through its style, a directly numbered paragraph, and a paragraph hidden by its style;
- a run in Symbol, and a plain paragraph;
- a default header and an even one, under `w:evenAndOddHeaders`.

Each of its styles, numbering and settings parts is cut to two thirds, and also removed:

| Part | Not well formed | Absent |
|---|---|---|
| styles | the heading's label gone, the hidden text shown; the direct label kept (`1.`, since the heading is no longer counted) | the same |
| numbering | every label gone, by style and direct | the same |
| settings | the even header gone; `fields_update_on_open` false; forms protection off | the same |
| font table | not consulted: nothing under `model/text/` reads one, and `create_package` makes none | the same |

So the reading here already proceeds as for a package that lacks the part. docx4j's defect is not here: there, an
unreadable styles part loses the direct labels too. Nor is its side finding: there, reading a package with no styles
part adds one. Each cut part's read raises with recovery off. Today, with recovery on, the cut settings part is
mended and keeps the even header, which came before the cut.

**The warning:**
- **Its form:** `TextWarning.PART_NOT_READ = "part.not_read"`, with no address and `part_name` its selector, once
  per part.
- **When:** where the reading first needs the part, not when `PackageTrees` is built:
  - **styles:** the first paragraph or run whose properties the reading resolves.
  - **numbering:** the first paragraph with a `w:numPr`, direct or through a style the reading could read. A
    document with no numbered paragraph raises nothing.
  - **settings:** the first read of the even header, forms protection or `fields_update_on_open`. Every result
    carries `fields_update_on_open` (REQ-801), so an unreadable settings part always raises it. With CR-029, the caps
    language (`w:themeFontLang`) joins these reads; today the text model maps no case.
  - **the font table:** never, until a reading consults it.
- **What raises it:** `PackageTrees` records each such part whose read failed, by name, with the part's exception
  (D8's `unreadable`), beside its story parts. A resolver or emulator that fails for any other reason is not this
  warning, and reads as today.
- **No log line:** the warning is the text model's one record of the failure.

### 4.7 The text model (D7)

CR-027 section 4.6 stands as built. What changes:
- **More parts are unreadable.** Every story part that is not well formed is now unreadable, not only those whose
  recovery bound nothing (CR-027's `<<w:document`).
- **The reason is lxml's:** "The content of /word/document.xml could not be read, so it cannot be read as text:
  Premature end of data in tag t line 2, line 2, column 351". `trees._root_reason` reads an `XMLSyntaxError`'s `msg`,
  so the reason is the same text as the engine's message.
- **CR-027's tests** (`tests/text/test_unreadable_parts.py`) go back to docx4j's own break (`</w:body>` to
  `</w:body`, `</w:hdr>` to `</w:hdr`), which retires CR-027 section 10.8's departure 1. Their reason becomes lxml's
  for it (`expected '>'`, at the cut's line and column).

### 4.8 The failure recorded (D8, Q5)

docx4j's `unreadable`, in this engine's terms:
- `XmlPart.unreadable` is the `Docx4JException` of the read that failed, or None, and `is_content_unreadable` is
  docx4j's `isContentUnreadable()`;
- a later `contents` or `read_contents()` raises it again without parsing, unless the parse settings have changed
  (`_ParseSettings`, as the private tree is keyed);
- `set_contents`, `set_bytes` and `set_xml` clear it.

The text model, the resolver and the validator each read the same parts. Without D8, a broken 50 MB part is parsed
to its error once per reader. `PackageTrees.unreadable` (CR-027) can then read the part's own record.

### 4.9 The validator and the save's check (D9, Q6)

- **`misplaced()` reports instead of raising.** Over a package or a part, it reports a part it cannot read as a
  finding, `part.not_well_formed`, with the part, lxml's message, the line and the column, and goes on to the next
  part. Today the walk reads a mended tree. After D1 it would raise on the first such part. Reporting fits its
  contract: it reports and changes nothing.
- **`save(check=True)`** checks unmarshalled parts only, so a part not well formed (never unmarshalled) isn't checked
  and is written byte for byte.
- **`source_tree()`** of a part replaced in code whose source isn't well formed gives None, not a raise. The source
  then held nothing readable, so every finding counts as introduced, where the check would otherwise fail the save.

### 4.10 Strict loading (D10)

`LoadOptions(strict=True)` adds nothing for well-formedness: a part that is not well formed is refused at every
setting. Today strict raised in 3 of the 13 `document.xml` cases, and for the wrong reason (section 2.2). Its
docstring says so. A strict load does not become an eager one. A caller who wants every part checked asks
`package.misplaced()` (D9).

### 4.11 No mended reading on request (Q7)

The opt-in would be `LoadOptions(recover=True)` passing the fork's option through, with one skipped item of kind
`recovered` per part so read (lxml's first error, its line and column), and `strict=True` raising on it. That needs
the fork to put lxml's error log into the skipped report as well. Not recommended unless someone asks: docx4j has no
such reading, and a mended part saved after an edit writes what libxml2 guessed.

### 4.12 Not here

- **The custom XML view's readers** (Q10). `CustomXmlPart.document_element` is None and `get_xml()` gives `""` for a
  part that is not well formed. They never mended (lxml's default parse), and the binding pass applies nothing from
  such a part.
- **A content API error code** for a part that cannot be read (`ContentError` with a `code`): CR-006's contract for
  the MCP server. Today it is the engine's `Docx4JException`, as it is for a part that fails to bind.
- **An eager load** that reads every part: `package.misplaced()` gives the check (D9).
- **Word's repair** of such a part: not measured, not emulated (section 3.5).

### 4.13 The CHANGELOG lines (Phase B)

- **Changed:** **A part whose XML is not well formed is refused when read, not mended.** `contents`,
  `read_contents()` and a generic XML part's `tree` raise `Docx4JException` naming the part and lxml's reason
  (`Problem with part /word/document.xml: Premature end of data in tag t line 2, line 2, column 351`). Before, lxml's
  recovery read the part up to its first error, or around it, and dropped the rest with no report, even under
  `LoadOptions(strict=True)`. So a file that read before, wrongly, now raises. The load is still lazy: a part never
  read costs nothing and is saved byte for byte, as docx4j does (CR-065 D1). The text model raises for a main part
  so read and gives `content.not_read` for another story part (CR-027), and reads as for a package without a styles,
  numbering or settings part so read, marked by the new warning `part.not_read` (WD 0.66 REQ-801). `misplaced()` reports such a part as `part.not_well_formed`. Needs docx4j-xsdata
  26.2.5.
- **Added:** `TextWarning.PART_NOT_READ` (`part.not_read`): a styles, numbering or settings part the reading needed
  and could not read, once per part, its selector the part's name.
- **Fixed:** a relationships part that is not well formed refuses the load. Before, the relationships after its first
  error were lost and an untouched save dropped the parts they named. `LoadOptions(mce_preprocess=True)` and
  `mc_load_mode="word"` raise `Docx4JException` naming the part, where they raised lxml's `XMLSyntaxError`, and Word
  mode no longer mends a part that holds no `mc:AlternateContent`.

## 5. Tests and oracles

- **The fork:** `tests/fork/test_recover.py` (section 4.1).
- **The engine:** `tests/openpackaging/test_not_well_formed.py`. The cases are section 2.2's twenty, built in code
  from one document this package makes and saves, each a replacement in the saved bytes as docx4j's `fixture` does.
  Over every case and at both `strict` settings:
  - `load()` returns, except for cases 17 to 20, which raise naming the part;
  - `contents` and `read_contents()` raise the part's `Docx4JException` with lxml's message, its deepest cause an
    `XMLSyntaxError`, and `source_tree()` gives None (section 4.9; corrected 2026-11-03, section 10.5 item 5);
  - `unreadable` and `is_content_unreadable` are set, and a second read raises the same exception without a parse;
  - an untouched save, a save after a read of another part, and a save after an edit of another part are each byte
    for byte for the broken part;
  - `set_contents` replaces it, and the save writes the replacement;
  - `misplaced()` reports `part.not_well_formed` and goes on;
  - `mce_preprocess=True` and `mc_load_mode="word"` raise the same, with and without `mc:AlternateContent`;
  - a custom XML part's `tree` raises naming its part;
  - the sniffs give the same over the corpus as with recovery.
- **docx4j's `UnreadablePartTest`,** its five engine tests ported over its own breaks (`</w:body`, `</w:hdr`, a
  `w:defaultTabStop` left open): the main part's and a header's reads refused, with the body still read; the
  settings refused; a replacement accepted and saved; a created package's parts unaffected. CR-027 ported its four
  text-model tests already.
- **The text model:**
  - case 14 gives `content.not_read` for the header;
  - section 4.6's measurement is a test: each cut part's reading equals its absent part's, with one `part.not_read`
    naming the part;
  - none for a numbering part in a document with no numbered paragraph;
  - one only, over a reading that needs the part many times;
  - nothing logged by the text model (`caplog`);
  - CR-027's tests with docx4j's breaks.
  - **docx4j's fixture** for `part.not_read` (its CR-046 section 10, `3ad164bdf`, unpushed; WD 0.68 Appendix A.8,
    portfolio `6ffa998`): `docx4j-core-tests/src/test/resources/text-extraction/part-not-read/`, copied, not depended
    on, at the commit Phase A starts from. Its eight files are the intact document; styles, numbering and settings,
    each not well formed and absent; and numbering unneeded. Each is held to `expected/<fixture>.<view>.txt` in the
    three views (`list_labels` on, `small_caps: upper`) and to `expected-warnings.tsv`: each not-well-formed file
    reads as its absent twin, plus one `part.not_read`. The settings fixture's caps run under `themeFontLang` tr-TR is
    recorded pending until CR-029 lands.
- **The census of section 4.4.**
- **The fixed fixture:** `test_an_empty_rsid_is_dropped_where_nothing_was_kept` inserts into a paragraph with no rsid.
- **The gates:**
  - the full suite;
  - `scripts/roundtrip.py`, 50 of 50;
  - `scripts/compare_acceptance.py`, 0 differ;
  - `codegen/generate.sh --check`;
  - `scripts/release_check.sh` with the fork from PyPI at 26.2.5;
  - the parse guard (`tests/test_performance.py`).
- **No Word check:** what a save writes for a well-formed file doesn't change, and a part that is not well formed is
  written as read (Q11).

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| 0 | The fork's stage 15 (item 33, `ParserConfig(recover=...)`, `ParserError` from lxml's error), its tests, `CHANGES.md`, and the docx4j-xsdata 26.2.5 release, which the user publishes; `docs/UPSTREAM.md` row 29 here | ½ day in the fork |
| A | D1's engine half (`runtime.RECOVER`, the pin, `test_parser_configs.py`) once 26.2.5 is out; D2 to D10: `_parse`'s `try`, `DefaultXmlPart.tree`, the sniffs, `unreadable`, the validator's finding, `source_tree()`, the text model's `part.not_read` and reason, the census, the tests of section 5, the three fixed tests | 1¼ to 1¾ days |
| A2 | The list readers' refusal (section 4.6, amended 2026-11-03): the emulator's reads of the styles and numbering parts raise the part's own `Docx4JException`, the text model handing it no part it recorded as not read; its tests; the other raw readers' handling measured and recorded | ¼ day |
| B | The records: CHANGELOG (section 4.13), CLAUDE.md (the status, and "Round-trip and fidelity rules"' "Nothing may be dropped silently" gaining the refusal), `tests/README.md`'s map, the index row, the portfolio's entries through the docx4j-portfolio session, the docx4j hand-off of section 3.4, and `part.not_well_formed` to docx4j-core-ts through the portfolio (Q6) | ¼ day |

Everything in Phase A except D1's line can land before 26.2.5. Under 26.2.4 the tests of the refusal itself are
marked to skip, as CR-021's were before its phase 0, and the rest of D2 to D10 runs.

## 7. Risks

- **A file that read before now raises.** That is the point of the change, but a caller who opened truncated or
  unescaped files and got some text now gets an exception, with the part and the reason. The CHANGELOG says so under
  Changed. Q7's opt-in is there if someone asks.
- **A broken relationships part now refuses the whole file** (Q3). Before, the file opened without the parts those
  relationships named, and a save lost them, so the old behaviour was silent loss, not service.
- **A broken styles part blocks the content API** for that file (D6). The text model still reads it.
- **Something else relying on recovery.** The corpus has no part that is not well formed (section 2.5), and the
  suite has three tests that relied on it (section 2.6), all named.
- **The release is the user's.** Phase A's refusal waits for docx4j-xsdata 26.2.5.
- **CR-030 is drafted beside this one** and changes `XmlPart._parse` too. Its pass runs after a successful parse
  (section 4.4), so the two compose in either order. Whichever lands second rebases its `_parse` hunk.
- **The other ports are unmeasured.** Whether docx4j-core-ts's or docx4j-generated-objects-ts's parser recovers is
  for the portfolio to ask, with section 2.2's cases as the probe.

## 8. Questions (decided 2026-11-03)

All eleven are decided as recommended (the user: "CR-030 and CR-031 both decided as recommended").

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The rule: a part whose XML is not well formed is refused when its content is read, never mended, and the load stays lazy, with an untouched save writing its bytes (docx4j CR-065 D1)? | Yes. It is CLAUDE.md's "nothing dropped silently", XML 1.0's fatal error and docx4j's rule, and it costs nothing (section 2.7). |
| Q2 | Where: (a) the fork's stage 15, `ParserConfig(recover=...)`, which the engine sets False; (b) the engine alone, a strict lxml tree handed to the fork's `iterwalk` path; (c) the engine alone, a subclass of the fork's handler; (d) a strict pre-pass? | (a). It is free, keeps the parse streaming, follows CLAUDE.md's rule that the fork is where runtime behaviour changes, and is the patch to offer upstream. (b) is the fallback if the release should wait: no release, 5% faster, but about 9 MiB more peak memory per 790 KB part. |
| Q3 | A relationships part not well formed, the package's or a part's: refuse the load, naming it? And hand docx4j's reading of a part's relationships as none (section 3.4) to docx4j? | Yes to both. Without its relationships the load cannot know the parts, and both today's partial reading and docx4j's empty one lose parts on an untouched save. |
| Q4 | A styles or numbering part not well formed: the content API raises (the resolver cannot be built), and the text model reads without it, as docx4j does, logging one warning naming the part? A `TextWarning` for it goes to the WD as a question? | Yes, as written. The content API writes, so it must not resolve without styles. The text model is a reader, and docx4j's reading is the oracle. Whether REQ-801 should cover non-story parts is the portfolio's question for the WD. **Amended 2026-11-03:** asked as WD 0.65's open item, and ruled by the user as `part.not_read` ("part.not_read as recommended"; WD 0.66 REQ-801): a `TextWarning` for a part the reading needed and could not read, in place of the log line (section 4.6). |
| Q5 | Record the failure on the part (`unreadable`, `is_content_unreadable`), raised again without a second parse until the content is replaced (docx4j's `unreadable` and `isContentUnreadable()`)? | Yes. It is docx4j's design, and it saves a parse per reader of a large broken part. |
| Q6 | The validator: `misplaced()` reports a part it cannot read as `part.not_well_formed` and goes on, and `source_tree()` gives None for a source that is not well formed? | Yes. `misplaced()` reports and never raises for content; the code goes to docx4j-core-ts through the portfolio, as CR-020's codes did. |
| Q7 | An opt-in to lxml's mended reading, `LoadOptions(recover=True)`, each part so read reported as recovered in `skipped`? | No, unless someone asks. docx4j has no such reading, and it needs more of the fork (lxml's error log in the report). |
| Q8 | The exception: `Docx4JException` with today's message form ("Problem with part /word/document.xml: <lxml's message>"), lxml's `XMLSyntaxError` the deepest cause, and no new class? | Yes. It is docx4j's hierarchy, a caller catching today's parse failures keeps working, and the cause carries the line and column. |
| Q9 | A generic XML part (custom XML, VML): refused when its `tree` is read, or at load, as docx4j's eager parse refuses it? | When read. The lazy load is CR-002's rule, and a custom XML part nobody reads should not refuse the file. |
| Q10 | The custom XML view's `document_element` (None) and `get_xml()` (`""`) for a part that is not well formed: change them here? | No: leave them, recorded (section 4.12). They never mended, and the view's contract is Office JS's, where such a part cannot arise. The engine's `tree` now raises naming the part, for a caller who must know. |
| Q11 | A Word check, or a Word probe of how Word treats such a part? | No. No saved output of a well-formed file changes, a part that is not well formed is written as read, and Word's repair isn't the oracle (section 3.5). |

## 9. Effort

1¾ to 2¼ days here (A 1¼ to 1¾, A2 ¼, B ¼; `part.not_read` added ¼ day to A on 2026-11-03, and A2 ¼ day the
same day), and ½ day in the fork for
phase 0. No Word run.

## 10. Phase A (implementation notes, 2026-11-03)

Phase A as built. The user ordered it on 2026-11-03 ("do the remain fork work and start CR-031 phase A"), after
phase 0: the fork's stage 15 (`77398949`), docx4j-xsdata 26.2.5 on PyPI, and `docs/UPSTREAM.md` row 29 (`1fc04d3`).
docx4j was read at `3ad164bdf` (its CR-046 section 10, local and unpushed) by `git show` and `git archive`,
read-only: `Extractor`'s `consult()` and its callers, `TextWarning.PART_NOT_READ`, `Emulator.resolverOf`,
`PartNotReadTest` and the `part-not-read` fixture; and `UnreadablePartTest` for the engine's five tests. Every
question is as decided (section 8), section 4.6 as amended. Phase B's records (CHANGELOG, CLAUDE.md, the docx4j
hand-off of section 3.4, the portfolio's entries) are not here.

**Commits.** `0f73213` (the copy, 10.2), `57c0b94` (the code and the tests, 10.1, 10.3 to 10.6), and this
section's (the records).

### 10.1 What landed

| piece | where | what |
|---|---|---|
| D1 | `runtime.py`; `pyproject.toml`, `RELEASING.md` | `RECOVER = False` beside `EXACT_TYPES`, set by `parser_config()` (`recover=`, the fork's stage 15); the pin `docx4j-xsdata[lxml]>=26.2.5,<26.3` in `dependencies` and the `dev` extra, and `RELEASING.md`'s statement of it. Nothing else names the floor (CI installs the dev extra; `release_check.sh` and the readmes name no version); CLAUDE.md's is Phase B's |
| D2 | `openpackaging/parts/xml_part.py` | `_parse`'s `try` holds `splice_word`, `mce_preprocess_bytes` and the fork's parse: "Problem with part <name>: <reason>", the reason lxml's `msg` (`_reason`) where an `XMLSyntaxError` is among the causes, else the error's own text as before; `syntax_error_of(error)` finds it |
| D3 | `openpackaging/parts/default_xml_part.py` | `DefaultXmlPart.tree` raises `Docx4JException` naming the part, from lxml's error |
| D4 | `xml_part._root_ns_map`, `CustomXmlDataStoragePropertiesPart.item_id` | the sniffs without `recover=True` (lxml's default, strict) |
| D5 | `openpackaging/load.py`, `api.py` | no code: `rels.list` reads the part at load, so the refusal of D2 refuses the load; the module, `load_package` and `load()` docstrings say so |
| D8 | `xml_part.py` | `XmlPart.unreadable` and `is_content_unreadable`; `_unmarshal` keeps the failure of `_parse` with the `_ParseSettings` it was read under and raises it again (`with_traceback(None)`) without a parse while they hold; a parse that succeeds, `set_contents`, `set_bytes` and `set_xml` clear it |
| D9 | `validator.py`, `traversal.py`, `opc_package.py`; `xml_part.source_tree()` | `PartNotWellFormed` (`part.not_well_formed`: `message` lxml's, `line`, `column` from lxml's `position`, `node` the part), made by `part_findings` where the part's own tree cannot be read for an `XMLSyntaxError`, so `misplaced()` of a package or a part reports it and goes on; `source_tree()` gives None where the source's parse fails |
| D7 | `model/text/trees.py` | `_root_reason` reads an `XMLSyntaxError`'s `msg` |
| D6 (`part.not_read`) | `model/text/trees.py`, `extractor.py`, `result.py`, `__init__.py` | `TextWarning.PART_NOT_READ`; `PackageTrees.styles_part`, `numbering_part`, `settings_part`, `parts_not_read`, `part_not_read(part)` and `known_not_read(part)`; the extractor's `consult(part)`, through `resolver` (now a property), `_numbering_emulator()`, `_counting_emulator()` and `_settings()` (10.3) |
| D10 | `load.py` | `LoadOptions.strict`'s docstring: it adds nothing for well-formedness, and does not make the load eager |
| the census | `tests/openpackaging/test_not_well_formed.py` | 10.4 |

### 10.2 The copy

- `git archive 3ad164bdf docx4j-core-tests/src/test/resources/text-extraction/part-not-read | tar -x` into the
  session's scratchpad, then copied into `tests/fixtures/part-not-read/`, not under `tests/fixtures/text-extraction/`,
  whose copy is pinned at `docx4j-17.3.5` (CR-027 Q6): the eight `.docx`, `expected/` (24 files),
  `expected-warnings.tsv` and the README as `docx4j-README.md`. Each file's git blob id is docx4j's at `3ad164bdf`
  (`git hash-object` against `git rev-parse 3ad164bdf:<path>`), and a test holds each to `SHA1SUMS` (34 files).
- The builder, `build_part_not_read.py`, stays in docx4j, as `br-placement/`'s do. The `README.md` is this
  package's: the commit, each file, and the pending lines (10.5 item 1). Every text file is LF, so no
  `.gitattributes` rule is needed.

### 10.3 `part.not_read`, as built

- **Where the parts are read.** `PackageTrees` reads the styles and the settings when it is built, as it did, now
  through `_consult(part)`: `read_contents()` once, the tree kept, and the part recorded in `parts_not_read` only where
  its own record says its content could not be read (`part.unreadable is error`). A read refused otherwise (a strict
  load's skipped content) is not this warning, and reads as before, silently. A styles part so recorded gets no
  resolver (docx4j's constructor), so the reading is as without styles.
- **The numbering**, read by nothing typed in the text model (the emulator reads it with lxml), is checked by a
  parse that builds no tree (`_well_formed`, an lxml target that builds nothing, 0.1 ms on a 20 KB part); only where
  that fails is it read typed, for the engine's record. A typed read on every consult cost a first reading of
  `NumberingIndents.docx` 7 ms (18.5 ms before, 25.8 ms; best of fifteen); with the check it is 19.0 ms.
- **When the warning is raised** (`Extractor.consult`, docx4j's): the styles where `resolver` is asked (every
  paragraph's properties, a run's style, a paragraph mark, a run's Office JS span), so at the first paragraph; the
  numbering in `_numbering` for a paragraph whose effective `w:numPr` has a `w:numId` other than 0 (direct, or
  through a style read); the settings in `_fields_update_on_open` (at the start of every top-level reading and at the
  end of a part's or content's), `_is_forms_protected` (FORMTEXT under `fields: update`) and `_sections_with_headers`,
  there only where a section has an even header or footer reference (`_even_reference`). Counting checks only the
  record already made (`known_not_read`), never reads: a document with no numbered paragraph raises nothing.
- **One level down**, `fields_update_on_open` is False without consulting the settings (docx4j's `!nested &&`); the
  nested result's metadata was never used (`read_nested` takes blocks and warnings), and the first reading under
  `fields: update` is nested too. `read_nested` names the warning for the embedded package's part, as for
  `content.not_read`.
- **The warning**: kind `part.not_read`, no address, `part_name` the part's (the selector), kept once a part by
  `warn`'s key. `to_dict()` gives `{"kind", "address": None, "selector"}`; `str()` `part.not_read (/word/styles.xml)`.

### 10.4 The tests and the numbers

`tests/openpackaging/test_not_well_formed.py`, new, 145 tests (3.8 s):

- **section 2.2's twenty** (the base document made in code; each case's libxml2 message pinned by its start, as
  section 2.3 gives it): the load returning and every read refused, at both `strict` settings (32: `read_contents()`
  and `contents` the same exception, lxml's error the deepest cause, the message its `msg` with no `(<string>`, the
  line and column in it, nothing skipped, `source_tree()` None); a second read with `_parse` patched to fail (16); a
  change of `mc_load_mode` parsing once more (1); the three saves byte for byte for the broken part, a failed read
  among them (16); `set_contents` and `set_bytes` clearing the record, the replacement saved (5); `misplaced()`'s
  finding, for the package and the part, its `to_dict()`, and nothing changed (16); two parts not well formed, two
  findings (1); a cell finding in a replacement over a source not well formed refused by the save's check (1);
- **the load** (8): cases 17 to 20 at both settings, three `Docx4JException`s naming the relationships part and the
  content types' `InvalidFormatException`;
- **the `mc` passes** (6): `mce_preprocess=True`, `mc_load_mode="word"` and the default, with and without a
  block-level `mc:AlternateContent`, each reading the well-formed part and refusing the cut one alike;
- **a custom XML part** (1): `samples/invoice2013.docx`'s item cut in half, still registered by its id, `tree`
  refused naming it, the view's readers as before (Q10), saved as read;
- **docx4j's `UnreadablePartTest`, its five engine tests** (5), over its own breaks;
- **the sniffs** (18): over the corpus (`samples/` and every fixture's `.docx`, `.docm` and `.dotx`: 5,671 XML parts
  and 44 custom XML properties parts) as the recovering sniffs read them; each of the sixteen part cases read as far
  as its root (`{}` for case 13 alone); a properties part cut after its root still giving its id;
- **the census** (2): no `recover=` keyword but `False` or `RECOVER`, no `HTMLParser` or `lxml.html` in
  `docx4j_py/`; `RECOVER`, `parser_config().recover` and `parser().config.recover` False; the module-level parsers
  (`breaks`, `mce`, `mc_load_mode`, `default_xml_part`, `fragments`) refusing a cut document;
- **downstream** (17): the text model's refusal of cases 1 to 13 with lxml's reason; case 14's `content.not_read at
  header:rId5`; case 15's resolver and `body.text` raising the styles part's own exception, and case 16's
  `body.text` reading; the python-docx facade refusing case 1; the base document read whole (1).

`tests/text/test_part_not_read.py`, new, 52 tests (0.8 s):

- **the copy** (2) and **every row of `expected-warnings.tsv`** (24): the warnings exactly; the text exactly but for
  the case-mapping lines, each held to be the capitals of this package's line (Turkish under the settings'
  `themeFontLang`, where they can be read) and their number pinned per fixture (51 lines in all), so that CR-029
  turns them;
- **`PartNotReadTest`'s other four** (12): each not-well-formed file read as its absent twin in each view (9); the
  emulator counting direct numbering without readable styles; the unreadable part kept by an untouched save; one
  level down the embedded package's part named;
- **the routes** (8): one level down the settings not needed; a part's reading consulting them; a numbering part no
  paragraph needs raising nothing, with and without labels (2); nothing logged by the text model; the content API
  refusing an unreadable styles part while the text model reads on; a strict refusal not `part.not_read`; a
  well-formed numbering part never read typed;
- **section 4.6's measurement** (6): the document read whole; each part cut to two thirds reading as it removed,
  with its one warning (3); the font table never consulted; one warning a part over twenty numbered paragraphs, both
  headers, every view and markdown, in the order first needed (settings, styles, numbering).

`tests/test_parser_configs.py` gains `test_the_engine_s_config_never_recovers` (1); `tests/test_validator.py`
gains `test_the_part_not_read_fixture_holds_its_four_parts_not_well_formed` (1). 199 new tests in all.

Two mutations show the tests bite: `Extractor.consult` patched to answer True fails 26 of `test_part_not_read.py`'s
52; `recover=True` put back under `parser_config` fails 155 across it, `test_not_well_formed.py` and
`test_unreadable_parts.py`.

ruff (`check`, `format --check`, at the files' 120 columns): the two new files are formatted and clean. On the
changed files `check` finds what it found before (2 in `xml_part.py`, 1 in `load.py`, 2 in `opc_package.py`, 1 in
`extractor.py`, none elsewhere), and the formatter would change the same lines as before.

### 10.5 Departures, with why

1. **The case-mapping lines of docx4j's fixture are pending.** Its expected text maps case (`w:caps` by the run's
   language, the settings' `themeFontLang` tr-TR giving "İSTANBUL", the default "ISTANBUL"; `w:smallCaps` under
   `small_caps: upper`), and this package maps none yet: CR-029's, and there is no `small_caps` option here
   either. Phase A's order named the settings fixture's run under tr-TR; every case-mapping line differs, in seven
   of the eight files (three a file,
   one in the two without styles, where only the direct `w:caps` maps), 51 in all, and each is exactly the capitals
   of this package's line. Every other line and every warning is docx4j's. So the fixture is read with
   `list_labels: true` alone.
2. **No log line.** docx4j's `consult` logs WARN "Read as text without <part>, which could not be read" once a part;
   section 4.6 has none: the warning is the record.
3. **The numbering's consult checks well-formedness before a typed read** (10.3). docx4j's `readable` unmarshals the
   part, which costs it nothing extra since its emulator reads the same JAXB tree; here it would be a parse no reader
   made. A well-formed numbering part whose typed parse would fail otherwise (the model's binding, not the XML) is
   therefore read by the emulator, as before, without the warning.
4. **Counting asks the record, not the part.** docx4j's `count` calls `readable(numbering)` for every paragraph it
   counts; here `_counting_emulator` reads only the record `consult` made, and the emulator's own lxml reading of an
   unreadable part defines nothing anyway (section 4.4's "each keeps its handling"). The labels are the same.
5. **`source_tree()` gives None** for a source that cannot be read, as section 4.9 and the brief say; section 5's
   list has it raise with `contents`. A source whose parse fails for any reason held nothing readable, so it is None
   for any `_parse` failure, not only an `XMLSyntaxError`.
6. **The record is `_parse`'s failure only.** `unreadable` keeps a parse that failed (its XML, or the binding), not a
   strict load's refusal of skipped content, which re-parses as before; nor a part with no content yet (docx4j's
   "no content yet"), for which `contents` still raises "has no content". The record is keyed by `_ParseSettings`,
   which docx4j has no counterpart of.
7. **`misplaced()` reports only a part not well formed.** A part whose read fails otherwise (not seen: the lenient
   parse binds what it can) still raises, as before.
8. **docx4j's settings conveniences** (`removeEnforcement`, `attachTemplate`) have no counterpart here; the port of
   `settingsThatCannotBeReadAreNotReplaced` refuses `set_compat_setting` and `set_default_compat_settings`.
9. **The corpus of the sniffs' test** is `samples/` and `tests/fixtures/` one level deep (5,671 XML parts), not
   section 2.5's 7,126 with the nested packages and the rest of `tests/`.

### 10.6 Pre-existing tests whose expectation changed, and why

- `tests/content/test_integer_normal_form.py::test_an_empty_rsid_is_dropped_where_nothing_was_kept`: the fixture
  takes the first paragraph's own `w:rsidR` off before putting `w:rsidR=""` in, so the attribute is not given twice
  (section 2.6), and checks the edit is well formed. Its expectations are unchanged.
- `tests/text/test_unreadable_parts.py`: docx4j's breaks again (`</w:body>` to `</w:body` in the main part, the
  root's closing tag cut short elsewhere, `</w:hdr>` to `</w:hdr`), and the reason lxml's, `expected '>', line 2,
  column N` at the byte after the cut, the same text in the part's own exception. CR-027 section 10.8's departure 1
  is retired.
- `tests/openpackaging/test_mce_levels.py::test_a_corrupt_part_still_raises_in_word_mode` accepted lxml's
  `XMLSyntaxError` or `Docx4JException`; it now asks the engine's refusal, naming the part, lxml's error its cause.
- `tests/test_validator.py::test_the_corpus_holds_nothing_but_the_probes_values`: the copy's four parts not well
  formed are each a `part.not_well_formed`, exempted as the CR-066 probes' values are and held exactly by the new
  test.
- `tests/openpackaging/test_private_readers.py`: `PackageTrees._consult` is a `read_contents` caller, listed and
  reached; the settings moved to it from `_read`.

### 10.7 The gates

| gate | result |
|---|---|
| the full suite, slow included, the save check on (`tests/conftest.py`) | 9,495 passed, 8 skipped, 0 failed (6 min 47 s); the 13 warnings the fork's `ConverterWarning`s of other tests, as before. CR-030 Phase A's tree had 9,296 (its 9,295 and its 168th test); the 199 are this phase's |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50 parts canonically identical, 0 differences, 0 skipped, 34 benign respellings (`boolean-spelling`), 0 parts repaired |
| `scripts/compare_acceptance.py` | `out/acceptance/` built by `scripts/acceptance.py` before the change (at `1fc04d3`, `out/acceptance-before-cr031a/`) and after it (`out/acceptance-after-cr031a/`): 17 artefacts, 246 entries, 0 differ beyond the wall-clock dates |
| `codegen/generate.sh --check` | two generations byte identical (318 modules); no generated file changed |
| `scripts/release_check.sh`, the fork from PyPI at 26.2.5 | passed, at `57c0b94`: the wheel and the sdist of docx4j 0.2.0 built into `dist/`, `twine check --strict` passed; a clean CPython 3.12.13 environment with the wheel and docx4j-xsdata 26.2.5 from PyPI; the smoke script; the fast suite against the installed wheel, 6,831 passed and 53 skipped, every skip a sibling checkout (`../docx4j`, docx4j-core-ts), python-docx or ruff not found from the scratch directory, which was in the session's scratchpad rather than beside the checkouts (this phase's four files there: 215 passed); `out/acceptance/` from the wheel against the checkout's, 17 artefacts, 246 entries, 0 differ |
| the parse guard, `tests/test_performance.py` | Symbols' parse 199.8 ms against the limit of 211.2 ms (192.0 + 10%); the validator 34.1 ms, 17% of the parse. `recover=False` costs nothing measurable (section 2.7); the text model's first reading of a package with numbered paragraphs is within noise of before (10.3) |

No Word check (Q11): no save of a well-formed file changes, and a part not well formed is written as read.

### 10.8 Found, and left open

1. **The content API's list readers read an unreadable numbering part as absent.** Section 4.6's first bullet says a
   numbering part not well formed "raises in the readers that need the numbering". The writers do (they unmarshal
   the part), but the readers (`Paragraph.is_list_item`, `list_item`) go through the emulator, whose reader parses
   `get_xml()` with lxml and defines nothing for a part that will not parse; section 4.4 keeps each such reader's
   handling. So `is_list_item` is False there, silently. Measured on `part-not-read-numbering.docx`; not changed. With the
   styles part not well formed, the two headings numbered through their style read False too, while
   `paragraph.text` refuses (measured by hand). **Decided 2026-11-03** (the user: "make the list readers refuse"):
   section 4.6 as amended, Phase A2.
2. **The fork logs skipped items met before the fatal error.** The settings fixture cut inside
   `<w:compatSetting` logs "Skipped unknown element {w}compatSettin in CT_Compat" on `docx4j_xsdata`'s logger,
   from the parse that then fails: libxml2 reports the start of the cut element before the error. The part reports
   nothing (the parse raised), and the text model logs nothing. It is the fork's skipped-report log (`log=True`),
   CR-030 section 10.9 item 1's question again.
3. **A package with no styles part logs the resolver's warnings** ("No default paragraph style!!", "Style definition
   not found: Heading1") when the text model reads it (`part-not-read-styles-absent.docx`); one whose styles part
   cannot be read logs none, since no resolver is built over it. Before CR-031 too; docx4j's fix (b), a styles part
   added to a package without one, is not reproduced here (no part is added).
4. **WD 0.67 and 0.69's further parts** (`commentsExtended`, `commentsIds`, `commentsExtensible`, the docProps parts,
   the custom XML stores) await the user's ruling: none raises `part.not_read`. Such a part not well formed reads as
   before: `commentsExtended` through `_read` (none), the docProps through `metadata.py`'s and the field resolvers'
   own handling, a custom XML store through its view (Q10).
5. **The registry**, through the docx4j-portfolio session: `python/CR-031.A` implemented.

## 11. Phase A2 and Phase B (implementation notes, 2026-11-03)

Phase A2 and Phase B as built. The user ordered both on 2026-11-03 ("make the list readers refuse; start phase B"),
after section 4.6's amendment (`51f89dd`, section 10.8 item 1). Every question is as decided (section 8), section
4.6 as amended. No Word check (Q11): no save of a well-formed file changes, and a part not well formed is written as
read.

**Commits.** `f1cfa01` (Phase A2: the code, the tests and their map in `tests/README.md`), and this section's (Phase
B: the records).

### 11.1 What landed (Phase A2)

| piece | where | what |
|---|---|---|
| the refusal | `model/listnumbering/emulator.py`, `_root_of` | where lxml cannot read the part's bytes (`get_xml()`), or the part already holds the record of a read refused for its XML (`syntax_error_of(part.unreadable)`), it asks the part's own read (`read_contents()`), which raises the part's `Docx4JException` and keeps it as its `unreadable` (D8), or raises that record again without a parse. A well-formed part is read with lxml as before, its read never asked. A part that lxml cannot read here for another reason defines nothing, as before |
| the next read | `Emulator._ensure` | a read that raised is not marked done, so the next one raises again, from the part's record |
| a paragraph that names no list | `Emulator.names_a_list(p_pr)`; `resolve` split into `_named` (the style walk and `w:numId` 0, reading no definition) and the definitions half; `_arguments(p_pr)`, `number()`'s reading (CR-021's `decimal`), shared by both | where it is False, `number()` is None for the paragraph |
| the list readers | `model/content/lists.py`: `list_of`, `list_item_of`, `ListItem._label` (`_names_a_list`) | ask `names_a_list` before they count the paragraph's story (`labels_for` walks it whole), so a reader over a paragraph that names no list needs no definition; `ListItem.get_ancestor` and `get_descendants` read the label before they count. `is_list_item` already resolved first (`_resolve`) |
| the text model's emulator | `Emulator(package, read_styles=False)`; `model/text/trees.py`, `PackageTrees.__init__` | where the styles part is recorded not read (`known_not_read`), the numbering is read as for a package without styles: no style map, no resolver (docx4j's `resolverOf`, its fix (a)). `_ind_from_style_chain` asks `self.resolver` rather than the package's |
| the text model's resolver | `trees.without_numbering(resolver, error)`, `_WithoutNumbering`; `Extractor._without_numbering`, from `_properties`; `markdown._WmlToMarkdown._style_indent_left` | the paragraph merge asks the emulator for a numbered paragraph's level indent, so the shared resolver now raises the numbering part's refusal there. Where the error is that part's own record, the paragraph is resolved again over a resolver whose `numbering_indents` is None, which is what the merge gave a numbering that defines nothing, and the rest of the reading resolves over it. Found through the main part, as `PackageTrees` finds it, so that a trial works too. `part.not_read` stays `consult`'s |
| the census | `tests/openpackaging/test_private_readers.py` | `model/listnumbering/emulator.py::_root_of` listed, its docstring row, and reached over `REFUSED` (docx4j's numbering and styles fixtures not well formed) by `read_lists_refused` |

### 11.2 The list readers, fixture by fixture

Each reader from a fresh load of docx4j's `part-not-read` fixture (`tests/fixtures/part-not-read/`): ten paragraphs,
two `Heading1` numbered by their style (0, 1), two numbered directly (2, 3), six naming no list (4 to 9). "Raises"
is the part's own `Docx4JException`, its `unreadable`, with lxml's message and lxml's error the deepest cause:
"Problem with part /word/styles.xml: AttValue: ' expected, line 2, column 625" and "Problem with part
/word/numbering.xml: Specification mandates value for attribute w, line 2, column 395".

| reader | `-styles` | `-numbering`, 0 to 3 | `-numbering`, 4 to 9 | `-numbering-unneeded` | intact, `-absent` twins, `-settings` |
|---|---|---|---|---|---|
| `Paragraph.is_list_item` | raises | raises | False, as absent | False | as before |
| `Paragraph.list_item`, `Paragraph.list` | raises | raises | None, as absent | None | as before |
| `Paragraph.to_dict()` (its `list_item`) | raises (Phase A: its `alignment`) | raises | as absent | as absent | as before |
| a `ListItem`'s `list_string`, `level`, `sibling_index`, `get_ancestor`, `get_descendants`, `to_dict` | raises | raises | `""`, 0, 0, None, `[]`: as absent | as absent | as before |
| `Emulator.get_number`, `Emulator.peek` | raises | raises | None | None | as before |
| `Paragraph.to_api_script()` | raises at 2, 3 (its own `w:numPr`); 0, 1 and 4 to 9 as absent | raises at 2, 3; 0, 1 as absent | as absent | as absent | as before |
| `body.lists` | raises | raises | | `[]` | as before |
| a `List`'s `definition`, `level_types`, `paragraphs`, `level_exists`, `get_level_paragraphs`, `get_level_string`, `to_dict` | raises | raises | | | as before |
| the emulator's `instance_list_definitions`, `abstract_list_definitions` | raises | raises | | raises | as before |
| `body.to_api_script()` | raises | raises | | reads | as before |
| `attach_to_list`, `detach_from_list` | raises | raises | | | as before |

Readers that do not reach the emulator's refusal, each identical to its reading at `51f89dd` over all seven fixtures
(`HEAD`'s tree archived into the session's scratchpad): the text model (`extract` in the three views with
`list_labels`, its warnings, `markdown`, `to_markdown()`), `outline()` and `describe()` as dictionaries, and the
python-docx facade (`docx4j_py.docx.Document`, which has no list reader). The facade, `paragraph.text` and the
resolver's readers still raise the styles part's refusal over `-styles`, as in Phase A.

### 11.3 Readers that now raise beyond section 4.6's list

Section 4.6 names the list readers. The emulator's refusal reaches these too; each is recorded here, and none is
routed around:

1. **The property resolver's paragraph readers, over a paragraph that names a list**, the numbering part not well
   formed: `effective_p_pr`, `left_indent`, `first_line_indent`, `outline_level`, `alignment` (and every reader of
   `get_effective_ppr`), at the same four paragraphs. The paragraph merge asks the emulator for the level's
   `w:ind` (`StyleUtil`'s `NumberingIndents` seam, `Emulator.get_ind`), which needs a definition. Before, they read
   with no level indent, silently. Over the six other paragraphs they read as before, and `text` reads everywhere
   (run properties never ask the numbering).
2. **`Paragraph.to_dict()`**, through its `list_item`; **`to_api_script`**, for a paragraph with a `w:numPr` of its
   own (the list verbs it writes read the list), and so `body.to_api_script()`.
3. **The writers** `attach_to_list` and `detach_from_list` raise the part's refusal at the emulator, where they
   would have raised the same exception at `contents`.
4. **A dry run's text reading changed, for the better.** A trial cannot copy a part that cannot be read, so a
   trial's resolver met the part's refusal in its numbering seam before Phase A2 too, and the text model fell back
   to each numbered paragraph's own properties: over `-numbering`, `trial.to_markdown()` lost the two headings'
   `#`. It now resolves them as without the numbering, so the trial reads as the package does.

### 11.4 The other raw readers, unchanged

Section 2.1's last row, each over its part cut to two thirds (or as noted), measured from a fresh load; nothing here
changed in Phase A2. The probe is in this session's scratchpad (`cr031a2/raw_readers.py`), not in the repository.

| # | reader | part | returns or raises | its caller then |
|---|---|---|---|---|
| 1 | `content/styles.py`, `style_ids_of` | `styles.xml` (`part-not-read-styles.docx`) | `set()` | `ensure_style`, `insert_markdown` and `insert_comment` raise the styles part's refusal; an untouched save keeps its bytes |
| 2 | `content/reports.py`, `comments_in` | `comments.xml` (`comments-modern.docx`) | 0 | the outline's count, but `pkg.outline()` raises the part's refusal first (11.10 item 3); `body.get_comments()` raises it |
| 3 | `content/describe.py`, `_root` | `styles.xml`; `settings.xml`; `comments.xml` | None: `styles` is `[]` (6 intact), `compatibility_mode` 12 (15 intact), the comment authors `[]` (one intact); the part still in `parts` | a description; nothing acts on it |
| 4 | `customxml/template.py`, `_from_bytes` | `document.xml` (`invoice2013.docx`) | `()`: `describe_template` and `custom_xml_parts.describe()` give 0 bindings (20 intact) | `fill` and `apply_bindings` raise the main part's refusal |
| 5 | `customxml/parts.py`, `_properties_of` | `itemProps1.xml` cut after its root | the item id the load registered the data part by, which it lower-cases (`load.py`; `{5d7ba57f-...}` where intact gives `{5D7BA57F-...}`), and no schema refs (none intact) | `apply_bindings` applies 16 of 20, as intact. Cut inside its root's start tag, the load's sniff (D4) gives no id, the store is not registered, and `apply_bindings` applies 0 of 20, each skipped with "no custom XML part with store item id ...", the body canonically unchanged |
| 6 | `text/metadata.py`, `_core_dates` | `docProps/core.xml` (`invoice2013.docx`) | `{}` | the text model's `metadata` has no `core`, with no warning (section 10.8 item 4: WD 0.67 and 0.69's further parts await the user's ruling) |
| 7 | `text/bound_values.py`, `_parse` | `docProps/core.xml`, a plain-text control bound to `dc:title` | lxml's error, which `BoundValues._root` catches: no store | the text model reads the control's surface ("Old title", where intact it reads "Core title" with `sdt.binding_differs`), with no warning; `apply_bindings` applies 0 of 1 and says why |
| 8 | `revisions/removals.py`, `Removals._index` | a header with no revision markup, cut after a `w:commentReference` | an empty index: the header holds nothing | **acts on it**: `accept_all()` removes comment 7, which the unreadable header still references, and the save writes the comments part without it (the header byte for byte). Measured on a document made for it (`Headers.docx`, comment 7 referenced in a tracked deletion in the body and again in a header): with the header intact the comment is kept. The same reading counts a note's or a comment's reference in any story part nobody has read (11.10 item 1) |
| 9 | `chart/model.py`, `root_tag` | `charts/chart1.xml` (`charts/bar-chart.docx`) | the root's tag (libxml2 reports the root's start before the error); None where the cut is inside the root's start tag | the text model reads the chart (`ChartModel.read`, `read_contents`) and gives `content.not_read` either way |
| 10 | `parts/wml/comments.py`, `_root_for_read` | `commentsIds.xml`; `commentsExtensible.xml` (`comments-modern.docx`) | None: no `dateUtc` read | `Comment.creation_date` is read from `w:date` in the package's zone (08:21 UTC here, where intact `dateUtc` gives 23:21 UTC), with nothing said; `reply` raises the part's refusal (`tree`); `resolved` and `content` write other parts, the broken part saved as read |
| 11 | `openpackaging/load.py`, `_item_id`'s fallback | a properties part of another content type (`application/xml`), cut | None, and a `WARNING` "could not read ds:itemID" | the store is not registered: `apply_bindings` applies 0 of 20, each skipped with its reason |

**For the user:** row 8 is the one that reads silently as empty where its caller then acts: the accept and reject
engine removes a comment (measured) or a note (`_flush_notes` reads the same index; not measured) whose reference
stands in a story part it cannot read, and saves. Rows 7 and 10
read silently otherwise (a bound control's surface for its value; a comment's `w:date` in the package's zone for its
`dateUtc`), and nothing acts on either.

### 11.5 Tests and numbers

`tests/content/test_list_readers_refused.py`, new, 292 tests (1.1 s):

- **the styles part** (134): each of the twelve paragraph readers at each of the ten paragraphs, and each of the
  thirteen body and list readers, raising the part's own record, its message and lxml's error the cause; the
  emulator's first read of the styles refused, and refused again without a parse;
- **the numbering part** (137): each paragraph reader at the four paragraphs that name a list, raising; each at the
  six that name none reading as over the absent part, the part neither read nor unmarshalled; each body and list
  reader raising; `part-not-read-numbering-unneeded.docx` refusing nothing but the definitions themselves; the
  refusal raised again without a parse, at every reader; the record another read made raised as it is; a
  replacement read again;
- **`to_api_script`** (2): refused at the two paragraphs with a `w:numPr` of their own, as before elsewhere;
- **the resolver's readers** (6): each refusing the four paragraphs that name a list and reading the six others as
  over the absent part; `text` reading;
- **the twins** (7): the intact, `-styles-absent`, `-numbering-absent` and `-settings` fixtures reading their
  labels, lists and items as before; a well-formed numbering part neither parsed typed nor unmarshalled by the list
  readers (three fixtures);
- **the text model** (6): each broken fixture reading as its absent twin with its one `part.not_read`, before and
  after a list reader's refusal, the same record raised either way; the text model's emulator reading no styles
  part; `without_numbering` answering the numbering part's own record and nothing else; a dry run reading as the
  package; a numbered quote style's Markdown depth as without the numbering.

`tests/openpackaging/test_private_readers.py`: `_root_of` in `CALLERS`; `test_every_listed_caller_is_reached` reaches
it over `REFUSED`; `test_a_refused_list_reader_leaves_its_trees_as_parsed` (2). `tests/text/test_part_not_read.py`'s
`test_the_emulator_counts_direct_numbering_without_readable_styles` now asks the text model's emulator
(`read_styles=False`) for docx4j's oracle, and the package's for the refusal (11.6 item 1).

Mutations show the tests bite: Phase A's reading put back in `_root_of` fails 198 of the new file's 292, 1 of `test_part_not_read.py`'s 52
and 3 of `test_private_readers.py`'s fast 52; without the `names_a_list` guard, 54 of the new file's fail; the
text model's emulator reading the styles fails 8 of `test_part_not_read.py`'s 52 and 1 of the new file's; with no
resolver without numbering, 3 of the new file's fail (the dry run, the quote style, `without_numbering`).

**The cost**, best of fifteen, interleaved with `HEAD` (`51f89dd`) twice, CPython 3.14.6:

| | `HEAD` | Phase A2 |
|---|---|---|
| `extract(list_labels=True)`, `NumberingIndents.docx` | 18.78, 18.85 ms | 18.81, 18.75 ms |
| the same, `samples/Symbols.docx` | 248.4, 249.0 ms | 248.0, 246.6 ms |
| every paragraph's `is_list_item` and `list_item`, `NumberingIndents.docx` | 12.08, 12.07 ms | 12.22, 12.10 ms |
| the same, `part-not-read-intact.docx` | 2.76, 2.76 ms | 2.81, 2.79 ms |

The text model's first reading is unchanged; the list readers pay `names_a_list`'s style walk once a paragraph,
1 to 2% on a document whose every paragraph is asked. No well-formed part is parsed more often, and none is marked for
re-marshalling (held by the twins' test and by `test_private_readers.py`).

ruff (`check`, `format --check`, at the files' 120 columns): the new file is formatted and clean; on the changed
files `check` finds nothing new (the one `try`-`except`-`pass` added is marked `S110`, with its reason), and the
formatter would change no line this phase wrote.

### 11.6 Departures, with why

1. **docx4j's emulator reads on; this one refuses.** docx4j's `Emulator.resolverOf` (its fix (a), `3ad164bdf`) reads
   the numbering with no property resolver where the styles part cannot be read, for every caller. Here the content
   API's emulator raises the part's refusal (the user's ruling, section 4.6 as amended), and the text model's is
   built with `read_styles=False`, which is docx4j's reading. So docx4j's
   `theEmulatorCountsDirectNumberingWithoutReadableStyles` holds for the text model's emulator, and its port asks
   that one for the count and the package's for the refusal.
2. **The resolver's readers raise for a paragraph that names a list** (11.3 item 1). Section 4.6 named the list
   readers; the resolver's seam is the same emulator, and reading on there would read a numbered paragraph's indent
   without its level, silently, which is what the amendment refused for the list readers.
3. **The refusal comes from the part's own read**, asked only where lxml cannot read the bytes, so a well-formed
   part costs nothing and the exception is the one `contents` raises. A part lxml cannot read here for another
   reason (`get_xml()` decodes UTF-8, so a part in UTF-16 is one) is read typed once, privately, and defines nothing,
   as before; its tree goes unused.
4. **A dry run reads differently** (11.3 item 4): as the package does, where it lost its style-numbered headings.
5. **The text model's fallback resolver is made per reading**, only where the numbering part's refusal is met, and
   rebuilds its style caches; the shared resolver never caches what it computed without the numbering.

### 11.7 Phase B: the records

- **`CHANGELOG.md`**, `[Unreleased]`: under Changed, the refusal (section 4.13), with the text model's reading and the
  content API's refusals, the list readers and the resolver's readers over a paragraph that names a list among them;
  under Added, `TextWarning.PART_NOT_READ`, `XmlPart.unreadable` and `is_content_unreadable`, and
  `part.not_well_formed`; under Fixed, a relationships part refusing the load and the `mc` passes' exception; under
  Dependencies, the floor docx4j-xsdata 26.2.5 for stage 15's `recover`, 26.2.4's stage 14 with it.
- **`CLAUDE.md`**: the pin `docx4j-xsdata[lxml]>=26.2.5,<26.3` (since CR-031 Phase A, unreleased here); "Round-trip
  and fidelity rules"' "Nothing may be dropped silently" gaining the refusal; one clause in "The engine" (a read
  that fails kept as `unreadable`). No status bullet, as for CR-026 to CR-030: the index is the status record.
- **`tests/README.md`**: Phase A's rows checked against the tests (`test_not_well_formed.py`,
  `test_part_not_read.py`, `test_parser_configs.py`, `test_validator.py`, `test_unreadable_parts.py`,
  `test_integer_normal_form.py`, the `part-not-read/` fixture row): they hold, unchanged. Phase A2's rows (the new
  file's, `test_private_readers.py`'s sixteen callers, `test_part_not_read.py`'s changed oracle) landed with the code,
  as Phase A's did.
- **`docs/change-requests/README.md`**: CR-031's row, A2 and B done, complete.
- **The status line** of this CR.

### 11.8 The hand-offs

Both went through the docx4j-portfolio session on 2026-11-03, sent by this repository's session.

1. **Section 3.4's, to docx4j**: `Load3` reads a part's relationships that are not well formed as none, so an
   untouched save drops every part they named (case 17). The reproduction was staged at `out/cr031-handoff/`
   (git-ignored): `base.docx`, `document-rels-cut.docx`, `Cr031Probe.java`, a README and `SHA1SUMS`.
   **Received and confirmed** (docx4j's session, through the portfolio): on docx4j's own fixtures, docx4j-core
   17.3.6-SNAPSHOT (`Load3` and `RelationshipsPart` unchanged since `4fdd80063`), so the staged files were not
   needed. A cut `word/_rels/document.xml.rels` loads with no relationships for the main part, the parse error on
   stderr only, and an untouched save loses the relationships part and every part it named; a cut `_rels/.rels` is
   refused, but by a `NullPointerException`, not the parse error. The cause is `RelationshipsPart.unmarshal(InputStream)`,
   which catches every `Exception` and prints it. Both files are now docx4j's `broken-part-document-rels.docx` and
   `broken-part-package-rels.docx`, in its broken-part family, staged for the user's Word run (not yet started).
   **docx4j's fix awaits Word's reading of the two files**: if Word refuses both, docx4j refuses, naming the part and
   the parse error, as this package does; if Word repairs, keeping what it can, Q3 ("refuses the load") becomes a
   specification question, which comes back here through the portfolio. The fix is the user's to rule, as a docx4j
   CR-065 follow-on; no registry change yet.
2. **Q6's, `part.not_well_formed` to docx4j-core-ts**: **received, and deferred to the user**, with no counterpart
   report in core-ts yet. core-ts wrote no CR: it has no validator report, its nearest codes being its load-time part
   warnings, `PartUnmarshalException`, and `part.not_read` for the text model (proposed in its CR-012 section 28).
   Adopting the code is on the user's list under core-ts.

### 11.9 The gates

| gate | result |
|---|---|
| the full suite, slow included, the save check on (`tests/conftest.py`) | 9,789 passed, 8 skipped, 0 failed (6 min 53 s); the 13 warnings the fork's `ConverterWarning`s of other tests, as before. Phase A's tree had 9,495; the 294 are this phase's (the new file's 292, `test_private_readers.py`'s 2) |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50 parts canonically identical, 0 differences, 0 skipped, 34 benign respellings (`boolean-spelling`), 0 parts repaired |
| `scripts/compare_acceptance.py` | `out/acceptance/`'s artefacts built by `scripts/acceptance.py` from `HEAD`'s tree (`51f89dd`, archived into the session's scratchpad: `out/acceptance-before-cr031a2/`) and from the change (`out/acceptance-after-cr031a2/`): 17 artefacts, 246 entries, 0 differ beyond the wall-clock dates |
| `codegen/generate.sh --check` | two generations byte identical (318 modules); no generated file changed |
| the parse guard, `tests/test_performance.py` | Symbols' parse 198.5 ms against the limit of 211.2 ms (192.0 + 10%); the validator 35.0 ms, 18% of the parse |
| `scripts/release_check.sh` | not run: no packaging or dependency change since Phase A's run at `57c0b94` |

### 11.10 Found, and left open

1. **The accept and reject engine counts a story part it cannot read as holding nothing** (11.4 row 8). `Removals`
   scans a story part nobody has read in its bytes, with lxml, and reads one not well formed as empty, so a
   comment or a note whose reference stands in it can be removed with the reference the resolution took, and saved.
   A part with revision markup in its bytes is unmarshalled by the listing, so it refuses the resolution instead;
   the case is a part with no revision markup. Not changed (section 4.4: each raw reader keeps its handling); for the
   user.
2. **`is_list_item` reads a `w:numId` as written, `list_item` as Word reads it** (CR-021). `lists._resolve` passes
   the `w:numId`'s text, and `Emulator.number` its `decimal`: a paragraph with `w:numId w:val="1.0"` is no list item to
   `is_list_item` and has the label "1." through `list_item` (measured on `part-not-read-intact.docx`). Before Phase
   A2 too; `names_a_list` reads as `number` does, so the new guard agrees with `list_item`.
3. **`pkg.outline()` raises an unreadable comments part's refusal**, through `package_bodies` (`body_of` reads the
   part's `contents`), where the text model's reading would give `content.not_read`. Phase A's; not changed.
4. **A dry run over a package whose styles part is not well formed raises at its text reading**: the trial cannot
   copy the part. Phase A's; not changed.
5. **A custom XML properties part cut after its root gives its item id in lower case** (11.4 row 5): the fallback is
   the id the load registered the data part by, which it lower-cases. The bindings resolve as before.
6. **The registry**, through the docx4j-portfolio session: `python/CR-031.A2` and `python/CR-031.B` done, and
   CR-031 complete.
