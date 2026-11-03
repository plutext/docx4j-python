# CR-030: A `w:br` outside a run, read as Word reads it (the port of docx4j CR-073)

**Status:** Proposed 2026-11-03, at the user's request ("yes", to drafting CR-030 and CR-031). It ports docx4j CR-073's three rules, so that a `w:br` found directly in a paragraph, a hyperlink or the body, or inside a `w:t`, `w:instrText`, `w:delText` or `w:delInstrText`, is put where Word 365 reads it before the part is bound, wherever the fork's skipped report names one, with no break or text lost and nothing measurable paid by a normal part (seven questions, section 8; two phases, 1½ days).

*The number:* CR-029 is the latest CR here. CR-031 is being drafted beside this one.

**Depends on:**
- **docx4j CR-073** (complete 2026-11-03; phase 1 `95c6b42d3` on `VERSION_17_3_6`, unreleased; registry
  `docx4j/CR-073.1`). Its rules in `mc-preprocessor.xslt` are the specification. Its probe and Word 365's saves of it
  (`docx4j-core-tests/src/test/resources/br-placement/`, rounds 1 and 2, `a2a971416` and `61147016e`) are the oracle.
- **The sibling, core-ts CR-021** (implemented 2026-11-03, uncommitted; registry `1a881d5`). It is the same three rules
  as a DOM pass before unmarshalling. The warning's code, its fields and its counts were agreed between the two ports
  through the portfolio session (section 3.2).
- **No fork change, no schema patch and no generated file.** The pass is a hand-written module under
  `openpackaging/`, which `codegen/clean.py`'s `KEEP` holds whole. `codegen/generate.sh --check` is a gate only.

**Settled before the questions** (these are not open):
- **The warning.** The code is `br.outside_run`, one per part, with the fields `part_name`, `breaks` and `text_elements`
  (core-ts: `partName`, `breaks`, `textElements`). It was agreed with core-ts through the portfolio, provisional on the
  user's word on each CR. Its message is docx4j's WARN verbatim, after the part name (section 4.3).
- **The counts are the stylesheet's** (`mc-preprocessor.xslt` lines 834 to 837 at `95c6b42d3`):
  - `breaks` counts every `w:br` whose parent is a run container or `w:body`. A `w:br` directly in a block-level
    `w:sdtContent` is counted but left where it is.
  - `text_elements` counts every text element holding an element, at any depth. So a `w:t` nested in a `w:t` that
    holds a break counts twice.
- **Rule 2's pieces take the `xml:space` of the nearest ancestor that has one, else none**, as the stylesheet writes
  them (lines 886 to 897). core-ts's section 3.1 had said "preserve", and it is corrected.
- **Saves are clean** (the user's rule): a save writes Word's normal form, here the repaired shape, and not the
  source's misplaced break.

## 1. Why

docx4j's load survey of the Common Crawl corpus found a `w:br` outside a `w:r` in 206,351 of 3,722,083 documents
(docx4j CR-073 section 2; portfolio `docs/docx_corpus_use.md` section 8.14).
- **The blank line.** 89% of them are one web-to-docx generator's blank line, `<w:p><w:br/></w:p>`.
- **Text with breaks in it.** PHPWord writes `<w:t>a<w:br/>b</w:t>`, and a template filler nests `w:t` in `w:t`.
  About 8,400 documents have one of these.
- **The rest.** A letterhead template puts breaks in a `w:instrText` outside any field, and ten documents have a
  `w:br` directly in `w:body`.

Word 365 opens all of them without a repair, and reads each as a break at that point, of its own type, keeping all the
text (docx4j's probe, section 3.1).

Here every such `w:br` is skipped by the parse and reported in `part.skipped`, so nothing is lost silently. But the
break is not in the tree, so every reader and every save goes without it:
- **The text model, `Paragraph.text`, search and Markdown** run the text either side together.
- **A save of the part, once unmarshalled, drops the break.**
- **Inside a `w:t`, the text after the break is lost too.** That is the reverse of docx4j 17.3.5, which keeps the text
  after the last break and loses the text before it.
- **A `w:t` nested in a `w:t` loses all its text.**
- **A break in a field code cuts the code**, so `fields: update` evaluates half a field.
- **`LoadOptions(strict=True)` refuses all thirteen of docx4j's probe files.**

## 2. What exists here (measured 2026-11-03 at `9d869f4`)

The probes are `head.py` (today's reading), `corpus.py` and `symbols.py` (the census and the costs), `prototype.py`
(docx4j's three rules over lxml, installed into `XmlPart._parse` in the probe's own process), `pairs.py` and `edges.py`.
They are in the session's scratchpad (`cr030/`) and run with `.venv-fork/bin/python`. Nothing in the repository was
changed to measure any of this.

### 2.1 The parse path

- **`XmlPart._parse(data, settings)`** (`openpackaging/parts/xml_part.py`) runs three steps:
  1. `splice_word` where `mc_load_mode="word"`, for CR-012 Phase A's block-, row- and cell-level
     `mc:AlternateContent`;
  2. then `mce_preprocess_bytes` where `LoadOptions(mce_preprocess=True)`;
  3. then one lenient parse by the fork, with the skipped report.
- **`_unmarshal`** keeps the report as `part.skipped` and `part.kept`, logs each skipped item, and raises under
  `strict=True`.
- **Every parse of a part goes through `_parse`:**
  - `contents`;
  - `read_contents()`, the private tree that `contents` later adopts (CR-002 section 16);
  - `source_tree()`, the fresh parse of the source that CR-020 Phase B's save check compares with;
  - a part whose bytes were set (`set_bytes`), on its next access.
- **`insert_ooxml`** takes one of two routes:
  - a flat OPC package goes through `load()`, and so through `_parse`;
  - a bare fragment goes through `wml.all(..., wrapper="body")` (`fragments._parse_one`), which is strict by default.
- **There is no repair pass here.** `mce.py` and `mc_load_mode.py` are markup compatibility only. The malformed-nesting
  repairs of docx4j's stylesheet are "a later CR" (CR-002 section 12.3 item 7). This would be the first repair of that
  family, and `splice_word` is the precedent for a rewrite of the bytes before binding.

### 2.2 docx4j's thirteen probe files, read today

Each probe is from `br-placement/` at `95c6b42d3`. "Word" is this package's reading of Word 365's save of the same file
(`*-word365.docx`). Each of the thirteen saves loads with nothing skipped.

| File | Skipped (the `w:br`'s parent) | What this package reads | Lost by a save after unmarshalling |
|---|---|---|---|
| `br-1-p-alone` | 1 (`w:p`) | an empty paragraph, where Word's save has a break | the break |
| `br-2-p-between-runs` | 1 (`w:p`) | `BR2 beforeBR2 after`, joined | the break |
| `br-3-t-mixed` | 1 (`w:t`) | `BR3 before` only; docx4j 17.3.5 keeps `BR3 after` only | the break and `BR3 after` |
| `br-4-t-nested` | 3 (two `w:t` and the `w:br`, in `w:t`) | an empty `w:t`: all the text lost | all of it |
| `br-5-hyperlink` | 1 (`w:hyperlink`) | joined | the break |
| `br-6-cell` | 1 (a cell's `w:p`) | joined | the break |
| `br-7-footer` | 1 (the footer's `w:p`) | joined, in the footer | the break |
| `br-8-p-page` | 1 (`w:p`, `w:type="page"`) | no page break | the page break |
| `br2-1-run-properties` | 9 (7 in `w:p`, 1 in `w:hyperlink`, 1 in `w:t`) | every break gone in R1 to R8, and R6's `R6 after` lost | 9 breaks and `R6 after` |
| `br2-2-instrtext-field` | 1 (the field's `w:instrText`) | the code cut to ` QUOTE "I1 code before`. `fields: update` gives `I1 start "I1 code before I1 end`, where Word's save gives `I1 start I1 code beforeI1 code after I1 end` | the break and `I1 code after" ` |
| `br2-3-instrtext-bare-header` | 2 (the header's bare `w:instrText`) | the header's `w:instrText` is `" "` alone: both lines lost | both lines and their breaks |
| `br2-4-body-br` | 1 (`w:body`) | no empty line before `B1 after` | the break |
| `br2-5-body-br-page` | 1 (`w:body`, `w:type="page"`) | no page break | the page break |

Over the thirteen:
- **24 skipped entries.** Each `w:br` is reported, as is each `w:t` nested in a `w:t`.
- **22 `w:br` in the sources' story parts, and none after an unmarshalled save.**
- **`strict=True` raises on all thirteen** (`Docx4JException`).
- **An untouched save writes every part byte for byte.** Only `[Content_Types].xml` and the relationship parts are
  rewritten, as for every document.
- **Compared as docx4j's `BrOutsideARunTest` compares them, all thirteen differ from Word's saves.** The comparison
  takes text, breaks and their types, field characters and paragraph ends, in order, each with its run's bold, italic
  and size. Plain text (headers and footers included) differs in twelve. `br2-2` agrees only because field codes are
  hidden.

### 2.3 The skipped report names every shape

Every misplaced element is in the report, so the report can be the trigger:
- a `w:br` with a run-container or `w:body` parent;
- an element inside a text element: the `w:br`, and a nested `w:t`.

The report's **`path` names the parent element** (`.../w:r/w:t/w:br`, `.../w:p/w:hyperlink/w:br`). Its
**`parent_qname` does not, for a choice class.** It is the namespace and the class's `Meta.name`:
- `{w}RT` for a `w:t`, `{w}PHyperlink` for a `w:hyperlink` and `{w}RInstrText` for a `w:instrText`;
- the element's name only where the class has one: `{w}p`, `{w}body`.

A trigger keyed on `parent_qname` missed five of the thirteen files. So the trigger reads the path's last step but one
(section 4.1).

### 2.4 The corpora

No document here has a misplaced `w:br`. This was checked by an exact lxml detection, docx4j's two `count()`
expressions, over every story part: the main part, headers, footers, notes, comments and the glossary.

| Corpus | Files | Story parts | Bytes | Parts with the shape |
|---|---|---|---|---|
| `samples/` and `tests/fixtures/` | 380 | 1,298 | 7.5 MB | 0 |
| `samples/` and all of `tests/` | 443 | 1,473 | 9.6 MB | 0 |
| `out/cr012-b-corpus/` (this repository, core-ts and docx4j fixtures, its `COMMITS.txt`) | 714 | 1,984 | 15.3 MB | 0 |

No test here pins a misplaced `w:br`: no inline XML of the shapes, and nothing that expects one skipped or refused.

### 2.5 The cost of each trigger

`samples/Symbols.docx`'s main part is 790,857 bytes and holds two legal `w:br`. Each figure is a median of nine runs
(99 for the microsecond ones).

| | Symbols' main part | `samples/` + `tests/fixtures/` | `samples/` + `tests/` | `cr012-b-corpus` |
|---|---|---|---|---|
| The fork's parse (`XmlPart._parse`) | 199.4 ms | 1,488 ms | 2,072 ms | 3,669 ms |
| **(a)** the skipped report scanned | 0.09 µs (empty) | 0 parts trigger | 0 | 0 |
| **(b)** a `<w:br` substring test, then the lxml detection where it hits | 84.7 µs + 12.2 ms (it hits) | 14 parts hit, 28 ms (1.9%) | 16 parts, 32 ms (1.6%) | 155 parts, 97 ms (2.6%) |
| **(c)** an lxml parse and the detection walk, always | 10.4 + 1.8 = 12.2 ms (6.1%) | 101 ms (6.8%) | 129 ms (6.2%) | 204 ms (5.6%) |

- **Two more numbers for (b).** A test that does not assume the `w:` prefix (a regular expression) takes 544 µs on
  Symbols. `b"br" in data` hits nearly any English text.
- **What an affected part pays under (a).** Its parse is thrown away, its bytes are rewritten, and it is parsed again:
  - Symbols' main part with one stray `w:br` added: 413.7 ms, 2.07× its parse. The rewrite alone is 17.5 ms.
  - The thirteen probes' 13 affected story parts: 3.0 ms to parse, 0.9 ms to rewrite, 3.2 ms to parse again (2.40×).
- **The parse guard won't catch it.** `tests/test_performance.py`'s guard (CR-020 section 13, 10%) times
  `runtime.parser()` directly, so a pass in `XmlPart._parse` is outside it. Section 5 adds a guard of its own.

### 2.6 docx4j's rules as a prototype

`prototype.py` has the three rules of `mc-preprocessor.xslt` (lines 833 to 951) over an lxml tree, decided on the tree as
read, then applied. They run under trigger (a) in `XmlPart._parse`. Over the thirteen:
- **All thirteen equal Word's saves**, by `BrOutsideARunTest`'s comparison over every story part. `br2-3` equals it with
  its `w:instrText` read as `w:t`, as docx4j's test reads it (D6, Q6).
  - D1's run properties are Word's in all eight of `br2-1`'s cases. R3 and R5 take the mark's `w:sz w:val="28"`. R1,
    R2, R4, R7 and R8 take none, and R6 stays in its bold run.
  - `br2-2`'s field code is split round its break. `fields: update` now gives Word's save's reading,
    `I1 start I1 code beforeI1 code after I1 end`.
- **Plain text** equals Word's in twelve. `br2-3`'s header differs by D6 alone: its bare `w:instrText`, now whole, is a
  field code and hidden.
- **The records.** Nothing is left in `part.skipped`, `strict=True` loads all thirteen, and the counts are docx4j's:
  `br2-1` has 8 breaks and 1 text element, `br-3`, `br-4` and `br2-2` have 0 and 1, and the rest have 1 and 0.
- **The save.** A save with `check=True`, read back, gives the same reading and no warning.
- **The diff.** `canon.compare` of each saved story part (15, 13 of them repaired) against its source run through the
  same pass is identical, normalising or not. Against the raw source, the differences are only what the pass did:
  - rule 1 shows as `child-order` and `element-renamed`;
  - rule 2 as `element-added`;
  - rule 3 as `element-dropped`.
- **docx4j's edge tests pass** (`BrOutsideARunTest`'s five methods, covering the six edges its section 8 lists):
  - a math run's `w:br` and a legal `w:r/w:br` are left alone;
  - whitespace is dropped without `xml:space` and kept with `preserve`;
  - the mark's `w:ins` is dropped and its `w:b` kept;
  - deleted text keeps its breaks, both in a `w:delText` and as a `w:br` in a `w:del`;
  - body breaks go before a paragraph with no `w:pPr`, and become a paragraph of their own before a table.
- **This package's own edges also pass:**
  - a `br` in no namespace inside a `w:t` (docx4j's survey sample `01af91808eac`) keeps the text whole, with no break;
  - a stretch of only U+00A0 is text, since the stylesheet's `normalize-space` treats only its four XML whitespace
    characters as space, unlike `str.strip`;
  - a `w:t` nested in a `w:t` that holds a break counts 2;
  - a nested piece takes the outer element's `xml:space="preserve"`;
  - in `preserve` mode, a typed paragraph in a kept block-level `mc:Choice` is repaired in place. Saved, it writes
    `<w:r><w:br/></w:r>` inside the Choice;
  - the same element reads the same under `mc_load_mode="word"` and under `mce_preprocess=True`;
  - a `w:br` directly in a block-level `w:sdtContent` counts 1 and 0 and stays in `part.skipped`, as docx4j leaves it;
  - a `w:tab` in a `w:t` keeps the text either side and loses the tab, as docx4j's rule does.
- **The save check needs `source_tree()` to rewrite too.** Take a cell whose `w:t` holds a break and whose last child is
  a table, so it does not end with a paragraph. `save(check=True)` carries the cell finding when `source_tree()`
  rewrites. When `source_tree()` parses without the pass, the save raises `SaveCheckError`: the cell's text differs
  (`a\nb` against `a`), so the carried finding looks introduced.
- **The fragment parser refuses the shapes today.** `wml('<w:p><w:br/></w:p>')` and `insert_ooxml` of a fragment
  holding one both raise `FragmentError`. `lenient=True` drops the break.

## 3. The counterpart

### 3.1 docx4j CR-073

The chain (its section 1): the first unmarshal fails, the stylesheet runs, and the second unmarshal drops what no
template matched. A `w:br` in a `w:t` restarts the text value, so docx4j keeps only the text after the last break.
Phase 1 (`95c6b42d3`, its section 8) adds three rules to `mc-preprocessor.xslt`:

1. **Rule 1** (lines 849 to 865). A `w:br` goes in a run of its own, its attributes kept, where:
   - its parent is `w:p`, `w:hyperlink`, `w:smartTag`, `w:customXml`, `w:fldSimple`, `w:ins`, `w:del`, `w:moveFrom`,
     `w:moveTo`, `w:sdtContent`, `w:dir` or `w:bdo`;
   - and its nearest paragraph-or-block ancestor is a `w:p`.

   The run takes the mark's properties, without the mark's own elements (lines 868 to 875), where no earlier sibling on
   the break's path from its paragraph holds a `w:r`. Otherwise it takes none (D1).
2. **Rule 2** (lines 880 to 908). A `w:t`, `w:instrText`, `w:delText` or `w:delInstrText` holding elements becomes its
   stretches of text, under its own name, and its breaks, in order, in the same run.
3. **Rule 3** (lines 914 to 951). A `w:br` directly in `w:body` goes into the next `w:p`, after its `w:pPr`, with that
   mark's properties. Where no paragraph follows, it is a paragraph of its own.

The other parts of the CR:
- **The WARN, once per part** (lines 833 to 844, D3): "w:br outside a w:r: n outside a run, m text elements holding
  elements; read as Word reads them (CR-073)".
- **D4:** the save writes the repaired shape.
- **D6:** a `w:instrText` with no field round it is left a field code.
- **`BrOutsideARunTest`** has 18 tests: twelve probes compared with Word's saves, `br2-3` under D6, and the edges. Its
  trigger is docx4j's retry after a failed unmarshal, so a clean part pays nothing.

### 3.2 core-ts CR-021

core-ts implemented the same rules on 2026-11-03 (uncommitted; registry `1a881d5`):
- **`src/opc/repair/breaks.mts`** holds the rules. They run as a DOM pass in `XmlPart`, ahead of the package's hook,
  and on `insertOoxml`'s fragment route.
- **The warning** is `br.outside_run` with `partName`, `breaks` and `textElements`, docx4j's counts, in the
  `PartWarning` union (`XmlPart.warnings`, `pkg.warnings`).
- **Results:** all thirteen probes equal Word's saves, and 1,688 tests pass.

Its trigger has to differ from this one. Its runtime refuses a text element holding an element
(`PartUnmarshalException`), so it can't parse first and look. Instead, a detection walk runs over every part's DOM, at
0.4% of parse and unmarshal. Here the fork's report already names every shape (section 2.3), so no walk is needed.

**One wording to align.** core-ts's message (`breaks.mts` line 140) ends at "read as Word reads them". docx4j's WARN
ends "read as Word reads them (CR-073)". This CR writes docx4j's verbatim. The portfolio session should settle which
one both ports write.

core-ts also holds its rules to check 72's 72e and 72j (its CR-016 section 7.1): a third Word run giving the same rule,
cited here and not copied.

## 4. Design

### 4.1 Where the pass runs, and what triggers it (Q1)

In `XmlPart._parse`, after `splice_word` and `mce_preprocess_bytes` and after the fork's parse:
1. **The test.** The parse's skipped report is scanned. A part triggers where an item is either:
   - a `w:br` whose parent, read from the item's `path`, is one of rule 1's containers or `w:body`;
   - any item whose parent, read the same way, is a `w:t`, `w:instrText`, `w:delText` or `w:delInstrText`.
2. **The pass.** Where an item triggers, the bytes the fork read go through the pass (`breaks.repair_bytes`), and the
   fork parses the result.
3. **The result.** The second parse's tree, skipped report and kept list are the part's. The first parse is discarded.

That is docx4j's own design, parse and then repair, keyed on the report rather than on a failure:
- **A clean part pays a scan of its skipped report**, which is empty for every story part of the three corpora:
  0.09 µs on Symbols.
- **An affected part pays a second parse:** 2.07× to 2.40× its parse (section 2.5). In docx4j an affected part pays
  the same.

Because the pass is in `_parse`, every route that parses a part gets it: `contents`, `read_contents()`,
`source_tree()`, and a part whose bytes were set. `_adopt` keeps what the private parse found.

**The module** is `docx4j_py/openpackaging/breaks.py`, docx4j's rules over lxml, beside `mc_load_mode.py`. It imports
no model class at module level (`tests/openpackaging/test_threads_and_import.py`'s rule). It holds:
- `triggers(skipped)`;
- `repair_bytes(data)`, which returns the bytes and the two counts;
- `repair(root)`, in place;
- the record, `BrOutsideRun`.

### 4.2 The rules, on an lxml tree

The rules are docx4j's, as section 3.1 gives them, with the stylesheet's own details:
- **Decided on the tree as read, then applied.** The stylesheet matches its input, so a run that rule 1 or rule 3
  creates is never a "run before" for another break. Two breaks first in a paragraph both take the mark's properties.
- **Rule 1** wraps the `w:br` in a `w:r` and keeps its attributes (`w:type`, `w:clear`) and the whitespace after it.
  The mark's properties are a copy of `w:pPr/w:rPr`'s children less `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` and
  `w:rPrChange`. What is left is in `CT_RPr`'s order, since `CT_ParaRPr` puts those five first and last.
- **Rule 2** walks the text element in document order:
  - each text node becomes an element of the outer element's name;
  - each `w:br` at any depth becomes a `w:br` with its attributes;
  - any other element gives its text, and its own tag is dropped;
  - comments and processing instructions give nothing.

  A stretch takes the `xml:space` of the nearest element at or above its parent that has one, and none where no
  element has one. A stretch of only whitespace (space, tab, CR, LF, as `normalize-space` defines it) is kept only
  under `preserve`. The pieces replace the element in its run, its tail kept.
- **Rule 3** takes, for each `w:br` child of `w:body`, the next sibling element that is not a `w:br`:
  - where that is a `w:p`, the breaks become its first runs, after its `w:pPr`, each with a copy of its mark's
    properties;
  - otherwise each becomes `<w:p><w:r><w:br/></w:r></w:p>` in its place.

  docx4j also treats a paragraph its nested-paragraph template splits (`w:hyperlink/w:p`, `w:r/w:p`,
  `w:hyperlink/w:r/w:p`, `w:r/w:r/w:p`) as no paragraph. That test is copied, though nothing here splits one, so that a
  later port of the nesting repairs (CR-002 section 12.3 item 7) gives docx4j's result.
- **The bytes** are written back as `splice_word` writes them: `etree.tostring(root, xml_declaration=True,
  encoding="UTF-8", standalone=True)`. The source's root declarations are read from the original bytes, as now.

The pass keeps:
- deleted text, as `w:delText` and `w:delInstrText`;
- a math run's `w:br`, since `m:r` is not a named parent;
- a legal `w:r/w:br`, untouched;
- the content of a kept `mc` branch, repaired in place.

**The counts** are the stylesheet's (settled; section 2.6 measures both rules).

### 4.3 The record (Q2)

`BrOutsideRun` is a frozen dataclass:
- `code: ClassVar[str] = "br.outside_run"`;
- `part_name: str`, `breaks: int` and `text_elements: int`;
- `message`, the part name and docx4j's WARN verbatim:
  `"/word/document.xml: w:br outside a w:r: 8 outside a run, 1 text elements holding elements; read as Word reads them (CR-073)"`;
- `to_dict()`, with those five keys, for the MCP server's JSON.

Where it is kept:
- **`part.warnings: list`** holds the part's record. It is empty for nearly every part, and at most one record per
  part. It is set by every parse, as `part.skipped` is, cleared by `set_bytes`, and kept by `_adopt`.
- **`package.warnings`** lists them over every part that has been parsed, as `package.skipped` does.
- **The log.** The record is logged once per part at `WARNING` on `docx4j_py.openpackaging`, as docx4j's WARN is (its
  D3).

What it means for the other reports:
- **The repaired elements are not in `part.skipped`.** What the pass leaves (a `w:br` in a block-level `w:sdtContent`)
  stays there, so nothing is dropped silently.
- **There is no list of each break's address.** Nothing is taken out, unlike `resolved_alternate_content`'s elements.
  The source bytes stay in the store for any comparison, and docx4j's WARN counts and no more.

### 4.4 Strict loading (Q3)

The pass runs before strict's test. A part whose only skipped items the pass repairs loads under `strict=True`. Word
reads these files as stated, with no repair prompt, so the repair is a reading of the file and not lenience. `strict`
keeps its meaning: content the model could not hold was dropped. Anything the pass leaves still raises.

### 4.5 What a save writes

- **A re-marshalled part** writes the repaired shape, which is Word's own save's (docx4j's D4; settled). That holds
  under `save(normalise=False)` too: the repair is the parse's, not the save's normal form.
- **An untouched part** is written byte for byte, as every untouched part is: it was never parsed.
- **A private read** (`read_contents()`, the resolver and the text model) sees the repaired tree, and leaves the part
  untouched.

### 4.6 The round trip

The rule is: "An unmarshalled part saved unchanged must be canonically identical" to its source. It gains a second
exception beside `mc_load_mode="word"`'s (CR-012 section 4.1 item 3):
- **The exception.** A part whose parse repaired a `w:br` outside a run (`part.warnings` holds `br.outside_run`) is
  saved in the repaired shape.
- **`canon.compare(..., repaired=True)`** runs the source through the same pass first, and counts what it did as
  `br-outside-run` differences, the warning's two counts. Every other category must still be zero. The prototype's
  saves are identical to their repaired sources in all 15 story parts (section 2.6).
- **`scripts/roundtrip.py`** (Q5) runs the pass before its own parse, as `XmlPart` does, prints each part's counts, and
  compares with `repaired=True`. So a corpus document with the shape is held to the same rule, rather than failing on
  skipped content.
- **`tests/openpackaging/test_round_trip.py`'s docstring** gains the exception. No sample triggers, so its tests are
  unchanged.

### 4.7 The MCE passes' order

The pass reads the bytes after `splice_word` and `mce_preprocess_bytes`, since its trigger is the fork's report on
those bytes. In each mode:
- **`preserve`:** a kept branch's typed paragraphs are repaired in place (measured);
- **`word`:** the promoted branch's content is repaired;
- **`mce_preprocess=True`:** the selected branch's content is repaired.

A `ResolvedAlternateContent` record's element is the source's, parsed on its own as today, with no report. It is not
repaired: Word does not read a discarded branch.

**One departure, for an unseen shape.** The stylesheet matches its input tree, so it does not repair a `w:br` that is
directly in a paragraph only once its branch is selected (`w:p/mc:AlternateContent/mc:Choice/w:br`). Here, under
`mce_preprocess=True`, the pass sees the resolved bytes and does repair it. Unmeasured.

### 4.8 `source_tree()`

`source_tree()` goes through `_parse`, so it rewrites as the part's own parse did, with no change of its own.
Section 2.6 measured why it must: a carried cell finding otherwise looks introduced, and the save raises. A test pins
it.

### 4.9 Fragments and `insert_ooxml` (Q4)

- **`wml()` stays as it is.** It is code's fragment builder, and CR-020 refuses what code puts where it does not belong.
  Its `FragmentError` names the skipped `w:br`.
- **`insert_ooxml`'s fragment route reads its OOXML as Word reads a file.** It runs the pass over the wrapped fragment
  before parsing, through a private keyword of `wml.all`. The warning goes in that call's `ChangeReport.warnings`.
  This matches core-ts's D1, which covers `insertOoxml`'s two routes.
- **`insert_ooxml`'s flat OPC route** has it already, through `load()`.

### 4.10 What the readers gain

Every reader works on the tree, so with no change of its own each sees the breaks and the whole text:
- `Paragraph.text`, `Range`, `search` and `replace_text`;
- the text model and Markdown;
- the numbering emulator;
- the revisions engine, which keeps a deleted line break in a `w:delText`;
- the python-docx facade.

A field's code is whole, so `fields: update` evaluates all of it (`br2-2`).

### 4.11 Not here

- **A `w:instrText` outside any field** (`br2-3`), which Word shows as text: Q6.
- **A `w:br` directly in another block container** (`w:tc`, `w:txbxContent`, a block-level `w:sdtContent`, a header's
  root). Rule 3 is `w:body`'s alone, and none was seen. One in a `w:sdtContent` is counted and stays skipped.
- **`w:tab`, `w:cr` or `w:noBreakHyphen` outside a run.** That is docx4j's catalogue query, not yet made.
- **The malformed-nesting repairs and Strict documents' content** (CR-002 section 12.3 item 7). The rules match the
  transitional namespace, as docx4j's do.
- **A break in an evaluated field code.** The QUOTE evaluator drops a `w:br` inside the code: on Word's save of
  `br2-2`, as on the repaired source, `fields: update` gives `I1 code beforeI1 code after`. What Word's update gives is
  unmeasured. It is a finding for CR-025's evaluators, beside core-ts CR-021 section 8's.
- **The fork's `SkippedNode.parent_qname`** gives a class's name for a choice class (section 2.3). The trigger reads
  `path`, which is right. The skipped report is the fork's own feature, so this is a finding for the fork, not for
  `docs/UPSTREAM.md`. Nothing in the fork changes here.

## 5. Tests and oracles

- **The oracle.** `tests/fixtures/br-placement/` holds:
  - docx4j's 26 files (the thirteen probes and Word 365's saves), copied unedited from `95c6b42d3` (Apache-2.0; git
    blob ids docx4j's);
  - its README, copied unedited as `docx4j-README.md`;
  - a `README.md` of this repository's, giving the provenance, as `tests/fixtures/field-probe-cr067/` does;
  - a `SHA1SUMS`.

  The two builders stay in docx4j, named in the README.
- **`tests/openpackaging/test_br_outside_run.py`:**
  - **Each probe against Word's save:**
    - `BrOutsideARunTest`'s comparison over every story part (text, breaks and their types, field characters and
      paragraph ends, each with its run's bold, italic and size);
    - plain text and the text model for twelve;
    - `br2-3` compared with its `w:instrText` read as `w:t`, as docx4j's test reads it, and its header's plain text
      pinned as D6's difference.
  - **The records:** each part's warning, its counts and its message; `package.warnings`; nothing in `part.skipped`;
    no warning on any of Word's saves.
  - **Strict, and the saves:**
    - `strict=True` loads all thirteen;
    - a save with `check=True`, read back, gives the same reading and no warning;
    - `canon.compare(..., repaired=True)` is identical, normalising and not;
    - an untouched save is byte for byte.
  - **docx4j's edges** (`BrOutsideARunTest`'s five methods), and section 2.6's own:
    - a `br` in no namespace;
    - U+00A0;
    - the double count;
    - `xml:space` inherited;
    - a kept branch in each MCE mode;
    - a `w:sdtContent`'s `w:br` counted, left and skipped;
    - a `w:tab` in a `w:t`.
  - **The routes:**
    - `source_tree()`'s carried cell finding (section 2.6);
    - `read_contents()` then `contents` adopting the repaired tree with its warning;
    - `set_bytes` clearing the warning;
    - `insert_ooxml`'s fragment route repairing and reporting;
    - `wml()` still refusing.
  - **The cost, held deterministically:** Symbols' main part through `XmlPart` calls the fork's parser once and the
    pass never. That catches a change of kind (a pass run on every part) without a timing.
- **The gates:**
  - the suite, with the save check on (`tests/conftest.py`);
  - `codegen/generate.sh --check`, with no generated change;
  - `scripts/roundtrip.py` over `samples/`, unchanged since nothing triggers;
  - `scripts/compare_acceptance.py` over a regenerated `out/acceptance/`, unchanged since its documents are built by
    code with legal breaks only.
- **The Word check** (Phase B), since this changes saved output:
  - `out/cr030-word-check/` holds each of the thirteen probes loaded, every story part unmarshalled, and saved by a
    default save, each read back with `strict=True`;
  - it is staged for the user at `vbShares/Office 2016/__tmp/python-cr030/`, beside Word's own saves;
  - the question: does each open without a repair prompt and show its breaks where Word's save of the source shows
    them (a line, a page, the header's two lines as field code under D6)?
  - its record goes in this CR and `tests/README.md`, and any hang in `../docx4j-portfolio/docs/word-hangs.md`.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | `openpackaging/breaks.py` (the three rules, the trigger, `BrOutsideRun`); the pass in `XmlPart._parse`; `part.warnings`, `package.warnings` and the log; `insert_ooxml`'s fragment route; `canon.compare(..., repaired=True)` and `scripts/roundtrip.py`; the oracle's copy; `test_br_outside_run.py` | 1 day |
| B | The records: CHANGELOG (Fixed: a `w:br` outside a run read as Word reads it; Added: `part.warnings`, `package.warnings`, `BrOutsideRun`); `CLAUDE.md`'s round-trip exception, its engine paragraph and its status; `tests/README.md`'s rows; the index row; the registry entries `python/CR-030` and `.A`/`.B` through the docx4j-portfolio session (`depends_on: [docx4j/CR-073.1]`, core-ts/CR-021 named as the sibling). The Word check's files, built for the coordinator to stage | ½ day, then the user's Word run |

## 7. Risks

- **The trigger rests on the fork's report.** A change to the report's `path` format would silence the pass. The
  thirteen probes pin it: each must trigger. `parent_qname` is not used (section 2.3).
- **The rules could match legal shapes.** The parents are named, so a math run's `w:br` and a `w:r/w:br` never
  match, and no story part of the three corpora (1,984 in the largest) has any of the shapes. A legal shape would have
  to be in the skipped report to trigger at all.
- **An affected part is parsed twice**, at 2.07× to 2.40× its parse. That is docx4j's own cost, and in the survey 5.5%
  of documents.
- **Output changes for affected files.** A document with the shape gains its breaks and its text in every reader and
  every save. The Word check covers the saves. The CHANGELOG says so, as docx4j's does for 17.3.6.
- **Children of a text element other than `w:br`** (a `w:tab`) keep their text and lose their element, as in docx4j.
  The count includes them, and nothing is dropped silently: the warning names the part.
- **Shapes neither engine has measured:** a body `w:br` before a table or at the body's end; a `w:br` first in a
  `w:hyperlink` with no run before it. Both are ported as docx4j builds them.
- **The `mce_preprocess=True` departure** (section 4.7), for an unseen shape.
- **The message's suffix** differs from core-ts's until the portfolio settles it (section 3.2).

## 8. Questions

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The trigger: (a) docx4j's way, parse and then repair and parse again where the skipped report names a misplaced `w:br` or an element in a text element; (b) a `<w:br` substring test, then an lxml detection walk on a hit; (c) the detection walk on every part. | **(a).** A clean part pays a scan of an empty list (0.09 µs). (b) pays the whole walk on any part with a legal break: 12.2 ms on Symbols, 1.6% to 2.6% over the corpora. It also assumes the `w:` prefix, or pays 544 µs for a regular expression. (c) pays 5.6% to 6.8% of the parse on every part, and on Symbols 12.2 ms (6.1%) of CR-020's 10% parse budget. (a)'s cost falls on affected parts only (2.07× to 2.40×), and the report names every shape (section 2.3). core-ts needs its walk only because its runtime cannot parse first. |
| Q2 | The record: (a) `part.warnings` and `package.warnings`, lists whose one kind is `BrOutsideRun` (the settled code and fields, `to_dict()`), with the counts only; (b) an attribute of its own, `part.br_outside_run`, beside `resolved_alternate_content`; (c) the log alone. | **(a).** It is core-ts's `XmlPart.warnings` and `pkg.warnings`, so the two ports report alike, and a later repair (the nesting family) joins the same list. It follows `package.skipped`'s shape. The counts are docx4j's WARN. No address list is needed, since nothing is taken out. |
| Q3 | Strict loading: (a) the pass runs, and a part whose only skipped items it repaired loads; (b) `strict=True` refuses the shapes, as it does today. | **(a).** It is Word's reading of the file, not lenience: Word opens all thirteen without a repair, and docx4j and core-ts read them so with no option. What the pass leaves still raises. |
| Q4 | Fragments: (a) `wml()` keeps refusing, and `insert_ooxml`'s fragment route reads as Word does, its warning in the call's `ChangeReport.warnings`; (b) both refuse; (c) both repair. | **(a).** `insert_ooxml` takes OOXML as Office JS's `insertOoxml` takes it, and core-ts's D1 repairs on both its routes. `wml()` is code's own builder, where CR-020 refuses what does not belong and the error names it. |
| Q5 | `scripts/roundtrip.py`: (a) it runs the pass as `XmlPart` does and compares with `canon.compare(..., repaired=True)`; (b) it stays a raw parse, and a document with the shape fails it on skipped content. | **(a).** The harness measures what the engine saves, and the exception is the engine's (section 4.6). No sample triggers either way. |
| Q6 | A `w:instrText` with no field round it (`br2-3`), which Word 365 saves as `w:t` and shows: (a) out of this CR, as docx4j's D6 and core-ts's D5, a later item following docx4j's field model when it decides; (b) here, read as text. | **(a).** It is not about `w:br`. Its reach is 29 documents of the survey, in one letterhead's template, and docx4j keeps it a field code. With this CR its text and breaks are whole, as a field code. The tests pin `br2-3` with the `w:instrText` read as `w:t` and its header's plain text as the difference. |
| Q7 | Release: (a) with the next release; (b) on its own. | **(a)**, as core-ts's D6. Phase A needs nothing unreleased: docx4j's rules are the specification and its probe the oracle. |

## 9. Effort

1½ days: A 1, B ½, then the user's Word run.
