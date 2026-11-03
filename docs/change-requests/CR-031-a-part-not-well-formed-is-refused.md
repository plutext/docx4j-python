# CR-031: A part not well formed is refused when read, not silently mended --- the fork's lxml handler without recovery, docx4j CR-065 D1's refusal

**Status:** Proposed 2026-11-03, at the user's request ("yes", to drafting CR-030 and CR-031). A part whose XML is not well formed raises `Docx4JException`, naming the part and the parser's reason, when its content is first read, as docx4j's does (CR-065 D1), where today the xsdata fork's lxml handler recovers, reads what it can and drops the rest with no report, even under `LoadOptions(strict=True)`; the load stays lazy, an untouched save still writes the part's bytes, and a relationships part not well formed refuses the load. Three phases (section 6): 0, the fork's stage 15 (`ParserConfig(recover=...)`), released as docx4j-xsdata 26.2.5; A, the engine; B, the records. 1¼ to 1¾ days here and ½ day in the fork; eleven questions, each with a recommendation (section 8); nothing built.

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
- **Counterparts:** docx4j as above. docx4j-core-ts and docx4j-generated-objects-ts are not measured here (section 7).

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
- **The model's own raw readers** (section 2.1's last row) are strict already, and each keeps its handling.

### 4.5 Relationships and content types (D5, Q3)

- **A relationships part that is not well formed refuses the load**, the package's or a part's:
  `Docx4JException("Problem with part /word/_rels/document.xml.rels: ...")` from `load()`. Without its relationships
  the load cannot know which parts exist. Reading the relationships before the error and dropping the rest is
  today's silent loss (case 17), and reading none is docx4j's (section 3.4). With recovery off this is what the
  engine already does (section 2.3), because `rels.list` reads the part at load. So D5 is a test and a sentence in
  `load()`'s docstring.
- **`[Content_Types].xml`** is refused as today (`InvalidFormatException`, "Bad [Content_Types].xml: ...").

### 4.6 Styles and numbering (D6, Q4)

- **The content API refuses.** `get_property_resolver()` raises the styles part's `Docx4JException` (docx4j: an
  NPE), and so does every content API reader that resolves styles (`body.text` and the rest). A numbering part not
  well formed raises in the readers that need the numbering, and nothing else.
- **The text model reads without them.** That is docx4j's behaviour and CR-027's `PackageTrees` code, unchanged. It
  logs one warning per part, naming the part and the reason:
  `docx4j_py.model.text: /word/styles.xml could not be read: hidden text and indents are not resolved through styles
  (...)`. Today nothing is logged (`except Exception: self.resolver = None`).
- **No `TextWarning`:** REQ-801 names the story parts. Whether the WD should name a styles or numbering part that
  could not be read is the portfolio's question for the WD, and docx4j's and core-ts's together (Q4).

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
  so read and gives `content.not_read` for another story part (CR-027), and reads without a styles or numbering part
  so read, with a warning logged. `misplaced()` reports such a part as `part.not_well_formed`. Needs docx4j-xsdata
  26.2.5.
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
  - `contents`, `read_contents()` and `source_tree()` raise the part's `Docx4JException` with lxml's message, its
    deepest cause an `XMLSyntaxError`;
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
- **The text model:** cases 14 to 16 give `content.not_read` for the header, read without styles and without
  numbering, and log one warning each for styles and numbering (`caplog`); CR-027's tests with docx4j's breaks.
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
| A | D1's engine half (`runtime.RECOVER`, the pin, `test_parser_configs.py`) once 26.2.5 is out; D2 to D10: `_parse`'s `try`, `DefaultXmlPart.tree`, the sniffs, `unreadable`, the validator's finding, `source_tree()`, the text model's log and reason, the census, the tests of section 5, the three fixed tests | 1 to 1½ days |
| B | The records: CHANGELOG (section 4.13), CLAUDE.md (the status, and "Round-trip and fidelity rules"' "Nothing may be dropped silently" gaining the refusal), `tests/README.md`'s map, the index row, the portfolio's entries through the docx4j-portfolio session, and the docx4j hand-off of section 3.4 | ¼ day |

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

## 8. Questions

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The rule: a part whose XML is not well formed is refused when its content is read, never mended, and the load stays lazy, with an untouched save writing its bytes (docx4j CR-065 D1)? | Yes. It is CLAUDE.md's "nothing dropped silently", XML 1.0's fatal error and docx4j's rule, and it costs nothing (section 2.7). |
| Q2 | Where: (a) the fork's stage 15, `ParserConfig(recover=...)`, which the engine sets False; (b) the engine alone, a strict lxml tree handed to the fork's `iterwalk` path; (c) the engine alone, a subclass of the fork's handler; (d) a strict pre-pass? | (a). It is free, keeps the parse streaming, follows CLAUDE.md's rule that the fork is where runtime behaviour changes, and is the patch to offer upstream. (b) is the fallback if the release should wait: no release, 5% faster, but about 9 MiB more peak memory per 790 KB part. |
| Q3 | A relationships part not well formed, the package's or a part's: refuse the load, naming it? And hand docx4j's reading of a part's relationships as none (section 3.4) to docx4j? | Yes to both. Without its relationships the load cannot know the parts, and both today's partial reading and docx4j's empty one lose parts on an untouched save. |
| Q4 | A styles or numbering part not well formed: the content API raises (the resolver cannot be built), and the text model reads without it, as docx4j does, logging one warning naming the part? A `TextWarning` for it goes to the WD as a question? | Yes, as written. The content API writes, so it must not resolve without styles. The text model is a reader, and docx4j's reading is the oracle. Whether REQ-801 should cover non-story parts is the portfolio's question for the WD. |
| Q5 | Record the failure on the part (`unreadable`, `is_content_unreadable`), raised again without a second parse until the content is replaced (docx4j's `unreadable` and `isContentUnreadable()`)? | Yes. It is docx4j's design, and it saves a parse per reader of a large broken part. |
| Q6 | The validator: `misplaced()` reports a part it cannot read as `part.not_well_formed` and goes on, and `source_tree()` gives None for a source that is not well formed? | Yes. `misplaced()` reports and never raises for content; the code goes to docx4j-core-ts through the portfolio, as CR-020's codes did. |
| Q7 | An opt-in to lxml's mended reading, `LoadOptions(recover=True)`, each part so read reported as recovered in `skipped`? | No, unless someone asks. docx4j has no such reading, and it needs more of the fork (lxml's error log in the report). |
| Q8 | The exception: `Docx4JException` with today's message form ("Problem with part /word/document.xml: <lxml's message>"), lxml's `XMLSyntaxError` the deepest cause, and no new class? | Yes. It is docx4j's hierarchy, a caller catching today's parse failures keeps working, and the cause carries the line and column. |
| Q9 | A generic XML part (custom XML, VML): refused when its `tree` is read, or at load, as docx4j's eager parse refuses it? | When read. The lazy load is CR-002's rule, and a custom XML part nobody reads should not refuse the file. |
| Q10 | The custom XML view's `document_element` (None) and `get_xml()` (`""`) for a part that is not well formed: change them here? | No: leave them, recorded (section 4.12). They never mended, and the view's contract is Office JS's, where such a part cannot arise. The engine's `tree` now raises naming the part, for a caller who must know. |
| Q11 | A Word check, or a Word probe of how Word treats such a part? | No. No saved output of a well-formed file changes, a part that is not well formed is written as read, and Word's repair isn't the oracle (section 3.5). |

## 9. Effort

1¼ to 1¾ days here (A 1 to 1½, B ¼), and ½ day in the fork for phase 0. No Word run.
