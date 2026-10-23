# CR-015: Markdown output made safe --- the port of docx4j CR-056's Markdown half

**Status:** Proposed 2026-10-23, at the user's request ("Markdown-safety CR: draft it, covering both fixes above"), to port the Markdown half of docx4j CR-056, which docx4j decided and built the same day (`84ee9741c`, `30fb41024`), with its two fixes: a link whose target's scheme is not on docx4j's allowlist is written as its text, with no link; and under `markup_detail` `author` or `author_date` a CriticMarkup marker's author and date are written as text, escaped as any document text is, as REQ-703 of the text extraction WD 0.33 now requires (docx4j-portfolio `da07eab`). Both defects were measured here first (section 2): every link destination the Markdown writer writes is the document's target as it is, so a `javascript:` target is a live link, and a marker's detail goes out as raw HTML. The measuring found four more (section 2.5): character references such as `&#58;` carry a script scheme past docx4j's allowlist (Q3), and an address comment, `Outline.to_markdown()` and the readers' picture destination write document values unescaped (Q8 to Q10). docx4j-core-ts ported both fixes the same day (`57b2f98`, `86e2e9d`). Four phases: A, the links; B, the markers; C, the sinks found while measuring, as Q8 to Q10 decide; D, the records. Nothing is implemented; it waits on the user's answers to section 8 and a go.

**Depends on:**
- **docx4j**, cited by commit only (its working tree was not read):
  - CR-056, `docs/developer/change-requests/CR-056-html-output-script-links-and-css-identifiers.md` at `30fb41024`: section 3 (the rules), 4 (the proposal), 6 (D1 to D10), 10 (the phase 1 record) and 13 (Markdown's other sinks);
  - `org.docx4j.convert.out.common.writer.LinkTargets` (`docx4j-core/src/main/java/...`), new at `84ee9741c` and unchanged at `30fb41024`;
  - `WmlToMarkdown`'s two changes: the link at `84ee9741c`, the marker at `30fb41024`;
  - the tests: `docx4j-core-tests`' `LinkTargetsTest` (`84ee9741c`) and `HostileDocx` (`df014ec3d`, `84ee9741c`), and `docx4j-markdown`'s `MarkdownHostileLinksTest` (`df014ec3d`, `84ee9741c`) and `MarkdownHostileMarkupTest` (`30fb41024`);
  - docx4j's session's hand-off of 2026-10-23, relayed by the caller, which section 3.1 checks against the code and agrees with.
- **The WD** (`../docx4j-portfolio/docs/text_extraction_specification_wd.md` at `da07eab`, version 0.33):
  - REQ-703's sentence on a marker's author and date, and its EXAMPLE of docx4j's output;
  - REQ-703's links bullet, which says `[text](target)` and nothing about a target's scheme (Q11);
  - REQ-703's images bullet ("its external target is never written, since a renderer would fetch it");
  - section 11.
- **This repository:**
  - CR-011 Phase D, the writer (section 14: `markdown.py` and the `commonmark.py` port, byte for byte with docx4j's at `198d015a3`, section 14.5);
  - CR-011 section 14.3, the readers;
  - CR-003 section 3.5 and Phase K, the address comments;
  - CR-003's `Outline`.
- **No fork change and no schema change.** `codegen/generate.sh --check` is a gate only to prove it.

**Counterpart:**
- **docx4j CR-056's Markdown half.** Its other half doesn't apply: style IDs CSS-escaped in the HTML stylesheet (`CssIdentifiers`, phase 2), because this package writes no HTML and no stylesheet. WD 0.32's inspector kind `hyperlink.script` (docx4j-portfolio `55fab92`) doesn't apply either: there is no inspector here.
- **docx4j-core-ts's port of the same two fixes**, its CR-012 sections 21 and 21.1 (`57b2f98`, `86e2e9d`). Read here are those two commit messages, `src/model/text/LinkTargets.mts` at `57b2f98` and that commit's diff of `MarkdownExporter.mts`; its CR's text was not read.

## 1. Why

**A document's link target becomes a script link.** The Markdown writer writes `TextRun.link` as the link's destination. The extractor gives that field a `w:hyperlink`'s relationship target, or a HYPERLINK field's address, exactly as the document has it.
- A target `javascript:alert(1)` comes out as `click [here](<javascript:alert(1)>)` (measured, section 2.1).
- A Markdown renderer that doesn't check schemes makes that `<a href="javascript:alert(1)">`. Clicking it runs the script on the origin of the page that shows the Markdown. That is CWE-79, stored XSS: the defect apache/poi#1455 fixed in POI, and docx4j CR-056 fixed in docx4j's HTML, PDF and Markdown outputs.
- In Word, a `javascript:` hyperlink does nothing when clicked (CR-056 section 0). The output makes the document behave as Word never would.

**A change's author becomes markup.** In the `markup` view under `markup_detail: author` or `author_date`, the author and date go into the marker as one raw HTML-inline node, with the braces.
- An author of `<script>alert('A')</script>` reaches a renderer that passes HTML through (markdown-it-py's `commonmark` preset does) as a live `<script>` element (measured, section 2.2).
- WD 0.33's REQ-703 now says only the braces and operators are literal, and the author and date are escaped as document text is.

**Who renders this Markdown.** It is the agent surface's main read:
- `to_markdown(addresses=True)`;
- CR-006's `read` tool, which is `to_markdown(addresses=True)` under the default options and `markdown_blocks(extract(pkg, **options), pkg)` under others, `markup_detail` among them (CR-006 section 3.1);
- the README's "the cheapest thing to show a model".

A chat client, a notebook or a wiki that shows the answer as HTML is the page that runs the script.

**Parity.** This package's writer is a port of docx4j's. CR-011 section 14.5 found 1,385 of 1,385 exports byte-identical with docx4j at `198d015a3`.
- docx4j at `30fb41024` now refuses these targets and escapes these authors, so on hostile input the two disagree until this port lands.
- core-ts ported both fixes the same day.

**0.1.0 has the link defect in a worse form** (section 2.1): a target is written raw, so a `)` in it ends the destination and the rest is raw HTML.

## 2. What exists here (measured 2026-10-23 at `8d477e6`)

Every probe here is a read-only script in the scratch directory, run with `PYTHONPATH=$PWD .venv-fork/bin/python`, on documents built from XML for the purpose (a minimal zip after `tests/text/mini_docx.py`). None is committed.

### 2.1 Where a link destination is written

**One place.** The only `md.Link` the writer makes is in `_WmlToMarkdown._linked` (`docx4j_py/model/text/markdown.py`, line 894: `md.Link(target, None)`). Each run of consecutive runs with the same `TextRun.link` is one link.
- **What sets the field.** The extractor sets `TextRun.link` on `TEXT` runs only (`extractor.py` line 1455, `_link_of`), from one of two sources:
  - a `w:hyperlink`'s relationship target, then `#` and its `w:anchor` (`_hyperlink_target`, line 1532);
  - a HYPERLINK field's address and `\l` bookmark (`field_hyperlink_target`, line 2582), for complex and simple fields alike.
- **How the renderer writes it** (`commonmark.py`, `_link_like`, line 760). The destination goes out raw, or in angle brackets where it holds a space, a parenthesis, `<`, `>`, a line break or a backslash. Inside the brackets, `<`, `>` and `\` are backslash-escaped and a line break is written `&#10;`. `&` is never escaped.
- **Every block kind reaches `_linked`** through `_runs`: paragraphs, headings, list items, block quotes, table cells (`_cell_content`), note definitions and text boxes.
- **A code paragraph writes no link.** `_code_text` writes its runs' text alone.

**Every link path, today's output** (`markdown(extract(pkg), pkg)`):

| what the document holds | written |
|---|---|
| `w:hyperlink`, External relationship target `javascript:alert(1)` | `click [text H1](<javascript:alert(1)>)` |
| the same, target ` java` TAB `scri` LF `pt:alert(2)` (the tab and line feed as `&#9;` and `&#10;` in the relationships part) | `click [text H2](< java→scri&#10;pt:alert(2)>)`, with a raw tab at `→` |
| `vbscript:MsgBox(3)`; `data:text/html,<script>alert(4)</script>`; `JavaScript:alert(5)` | each a link: `(<vbscript:MsgBox(3)>)`, `(<data:text/html,\<script\>alert(4)\</script\>>)`, `(<JavaScript:alert(5)>)` |
| `onenote:///x`, and a `ms-word:ofe` target | each a link, unbracketed |
| an `r:id` to `javascript:alert(1)` with `w:anchor="frag"` | `[with anchor](<javascript:alert(1)#frag>)` |
| an anchor alone, `w:anchor="x) [y](javascript:z"` | `[hostile anchor](<#x) [y](javascript:z>)`: one destination, a fragment |
| a complex HYPERLINK field to `javascript:alert(6)`; a `w:fldSimple` to `javascript:alert(7)` | `[field F1](<javascript:alert(6)>)`; `[simple F2](<javascript:alert(7)>)` |
| a `\l` field, and an `https` field | `[field F3](#bm2)`; `[field F4](https://example.org/f)` |
| a picture whose `a:hlinkClick` (or `a:hlinkHover`) targets `javascript:alert(8)` | `![P1 clicked](data:image/png;base64,...)`: no link. The model never reads a picture's click target, and an `IMAGE` run has no `link` |
| a picture inside a `w:hyperlink` to `javascript:alert(1)` | `[see ](<javascript:alert(1)>)![P3 inside link](data:...)[ after](<javascript:alert(1)>)`: the link split around the image |
| a link in a table cell | `\|click [cell T1](<javascript:alert(1)>)\|` |
| benign targets: `https`, `mailto`, `file:///C:/...`, `sub/page.html`, a UNC path, `C:\docs\e.docx`, `tel:` | kept, the UNC and drive paths in angle brackets with their backslashes doubled |

**The readers write the same.**
- `pkg.to_markdown()`, `pkg.body.to_markdown()` and `paragraph.to_markdown()` gave the same links as `markdown(extract(pkg), pkg)`, pictures' destinations aside (the readers write `media/image1.png`).
- They are `markdown_blocks` over the text model's reading (`model/markdown/export.py`, `_written` and `_alone`).
- `Table.to_markdown()` reaches the same function through `table_markdown`. That is by reading the code; it was not probed alone.

**`Range.hyperlink`, measured on `create_package()`.**
- **The setter keeps the character.** `p.search("here")[0].hyperlink = " \x01java\tscri\npt:alert(2)"` keeps U+0001, and reading the hyperlink back gives the same string.
- **Every reader writes it raw.** Each wrote `click [here](< \x01java\tscri&#10;pt:alert(2)>)`, a raw U+0001 in the Markdown.
- **The save fails late.** It raises lxml's `ValueError` ("All strings must be XML compatible"), with no `code` (section 2.5 item 4).
- **A table cell's link** set to `vbscript:x` gave `|[cell](vbscript:x)|`.

**Not sinks:**
- **Plain text** writes no link target (REQ-701; measured).
- **The block model's JSON** (`TextResult.to_dict()`, REQ-706) carries `link` as the document wrote it. That is data, as docx4j's `TextRun` carries it ("the text model's run carries it and the outputs refuse it", CR-056 section 12). It stays.

**How one renderer reads today's output.** markdown-it-py 4.2.0, this package's importer dependency, in its `commonmark` preset:
- **The default refuses some.** Its `validateLink` refuses `[text H1](<javascript:alert(1)>)` as a link and keeps `onenote:///x`.
- **It percent-encodes the obfuscated target**, written `java%09scri%0Apt:alert(2)`. A browser doesn't strip encoded characters, so that is inert.
- **A renderer with no URL check writes the script link.** With `validateLink` replaced by one that accepts everything, as a renderer with no such check behaves, `javascript:alert(1)` became `<a href="javascript:alert(1)">text H1</a>`.
- No other renderer was measured.

**0.1.0, measured.** `git archive v0.1.0` (`d234949`) was extracted into the scratch directory and run over the editable fork (docx4j-xsdata 26.2.2, not 0.1.0's 26.2.1).
- **The script scheme.** `click [text H1](javascript:alert(1))`, as the caller measured in a `v0.1.0` worktree.
- **The breakout.** CR-003's exporter writes a target raw (`sink.node(f"[{inner}]({target})")`). So the target `https://example.org/a) <img src=x onerror=alert(9)> (b` ends the destination at its `)`: the output is `[text X1](https://example.org/a) <img src=x onerror=alert(9)> (b)`. markdown-it-py's `commonmark` preset renders that as `<a href="https://example.org/a">text X1</a> <img src=x onerror=alert(9)> (b)`, which runs script on load, with no click.
- **main has closed the breakout** through CR-011's writer: `[text X1](<https://example.org/a) \<img src=x onerror=alert(9)\> (b>)`, one inert destination.

### 2.2 Where marker detail is written

**Two places, both raw:**
- `_runs` (line 876) writes `md.HtmlInline(self._open(revision))`, where `_open` (line 964) is `{++` or `{--` followed by `_detail(author, date)`;
- `_comment` (line 947) writes `md.HtmlInline("{>>" + self._detail(...))`.

The renderer writes an `HtmlInline` as it is (`commonmark.py` line 597). `_detail` is `"author: "` or `"author, date: "`.

**Measured,** in the `markup` view with `comments: true`:

| what the document holds | `markup_detail: author` today |
|---|---|
| a `w:ins` by the author `[x](javascript:alert(1)) *e* {--z--}`, after document text holding the same string | `P1 doc text \[x\](javascript:alert(1)) \*e\* {--z--} {++[x](javascript:alert(1)) *e* {--z--}: NEW++}`: the text escaped, the author not |
| a `w:del` by `<script>alert('A')</script>` | `{--<script>alert('A')</script>: P2 GONE--}` |
| a `w:moveFrom` and a `w:moveTo`, each by such an author | `{--MF <script>alert('A')</script>: P3 MOVED FROM--}`, `{++MT <script>alert('A')</script>: P4 MOVED TO++}` |
| an insertion in a table cell | `\|{++TC <script>alert('A')</script>: T1 cell ins++}\|` |
| a comment, and a reply to it (`commentsExtended`), each by such an author | `{>>CA <script>alert('A')</script>: C1 text <<}{>>CR [x](javascript:alert(1)) *e* {--z--}: C2 reply<<}` |
| under `author_date`, a `w:date` of `<script>alert('A')</script>` | `{++Ann, <script>alert('A')</script>: NEW2++}`. The value is not an `xsd:dateTime`: the converter warns (`ConverterWarning`), keeps it as written and skips nothing, so a date is document text too |
| an inserted or deleted paragraph mark, a `w:rPrChange`, an inserted row | no marker at all, so no detail |
| `pkg.body.to_markdown(view="markup")` | `{++NEW++}`: the readers have no `markup_detail`, so their markers carry the type alone |
| `plain_text(extract(..., markup_detail="author"))` (REQ-702) | the detail as written: `{++[x](javascript:alert(1)) *e* {--z--}: NEW++}`. That is right, since plain text is not markup |

**Code blocks keep the string form, and should.**
- **How it is written.** `_code_text` (line 576) writes `_open`'s string into the fence's literal (line 587). The renderer makes the fence longer than any backtick run in the literal (`_max_run_length`).
- **Measured on a hostile author.** A `SourceCode` paragraph with an insertion whose author held three backticks and a line break was written in a four-backtick fence. The `<script>` stayed inside it, where nothing is markup.
- docx4j's commit says the same of its code blocks ("code blocks keep the string form").

**Everything else is escaped already** (measured, as CR-056 section 13 measured docx4j's): document text, a link's text, a picture's alt text and a comment's text (`\[x\]`, `\<script\>`, `\<\<}`).

**How the renderer reads it.** markdown-it-py's `commonmark` preset passes raw HTML through. `{--<script>alert('A')</script>: P2 GONE--}` rendered as `<p>{--<script>alert('A')</script>: P2 GONE--}</p>`, a live `<script>`. With the author as a text node, the same marker renders `{--&lt;script&gt;…`.

**The fix, simulated** as an `HtmlInline("{++")`, a `Text` and an `HtmlInline("++}")`, gives docx4j's form: `{++\[x\](javascript:alert(1)) \*e\* {--z--}: NEW++}`.

### 2.3 docx4j's tests, against this package today and against a sketch of the fix

The sketch is a scratch file, not committed:
- `LinkTargets` ported line for line;
- the writer's `_linked`, `_runs` and `_comment` monkeypatched to refuse a target and to write the detail as a text node, as docx4j's two commits do.

| docx4j test | today | under the sketch |
|---|---|---|
| `LinkTargetsTest`: 18 safe targets, 12 refused, 6 scheme readings, the property's added schemes and add-only rule, the excerpt | n/a (no `LinkTargets` here) | every assertion holds |
| `MarkdownHostileLinksTest.noScriptSchemeDestination` | 2 of 5 assertions hold. The obfuscated target's holds only vacuously: it looks for a literal line feed, and the writer writes `&#10;` | 5 of 5 |
| `MarkdownHostileLinksTest.benignLinksKept` | 3 of 3 | 3 of 3 |
| `MarkdownHostileMarkupTest.noRawMarkupFromAnyDocumentValue`, `type` / `author` / `author_date` | 5 of 5 / 5 of 7 / 5 of 7 (a raw `<script>`, and the two authors not escaped) | 5 of 5 / 7 of 7 / 7 of 7 |
| `MarkdownHostileMarkupTest.benignMarkersUnchanged` | 2 of 2 | 2 of 2 |

`MarkdownHostileLinksTest` builds its links in memory, since U+0001 cannot be saved, and names each paragraph with Java's `trim()` of its target. Java's `trim()` drops every character at or below U+0020. Python's `strip()` does not drop U+0001, and the content API refused the untrimmed text: `insert_paragraph` raised `text.not_writable` (CR-013 Q2). The port trims as Java does (section 6.1).

### 2.4 The corpus, before and after

**The corpus.** `out/cr012-b-corpus/` is the corpus CR-013 section 2.2 used (`COMMITS.txt`: this repository at `f1bb1d9`, core-ts at `df5b4af`, docx4j at `901a922a3`). It holds 422 distinct documents by content (`.docx`, `.docm`, `.dotx`, `.dotm`); every one loads and extracts.

**The links**, read with `include: headers, footers` and `comments: true`:
- 66 text runs in a link, in 32 documents, over 22 distinct targets;
- by scheme, in runs: `https` 38, `file` 12, no scheme 11, `http` 3, `mailto` 2;
- **none is refused**, and no target holds `&`.

**The markers.** No revision's or comment's author or date holds any of `` \ ` * _ [ ] < > ! & # ~ | { } ``.

**Under the sketch,** 0 of 422 exports in the accepted view change, and 0 of 422 in the `markup` view under `author_date` with comments.

**The conformance goldens.** `tests/fixtures/text-extraction/phase0/extraction-probe.md` links to `https://example.com/...` twice and to `file:///FLD49_FILE_MARK`, all kept.

**No test expectation in `tests/` writes a link outside the default schemes.** The only matches of a scheme after `](` are two pictures' `data:` URIs, which are images.

core-ts measured the same of its documents: "None of the 61 link targets in the 314 committed documents is refused" (`57b2f98`), and "The markup view's Markdown of the 315 committed documents under author and author_date (630 exports) is unchanged" (`86e2e9d`).

### 2.5 Found on the way

1. **A character reference carries a script scheme past docx4j's rule.** A CommonMark parser decodes entity and numeric references in a link destination; the writer doesn't escape `&`. Four targets pass the allowlist, as the sketch reads them:

   | target (as the relationship holds it) | `scheme()` | `is_safe` | after `html.unescape` | rendered by markdown-it-py with no URL check |
   |---|---|---|---|---|
   | `javascript&#58;alert(1)` | None (a `#` comes first) | true | `javascript:alert(1)`, refused | `<a href="javascript:alert(1)">` |
   | `&#106;avascript:alert(2)` | None | true | refused | `<a href="javascript:alert(2)">` |
   | `javascript&colon;alert(3)` | None (no `:` at all) | true | refused | `<a href="javascript:alert(3)">` |
   | `&#x6A;avascript:alert(4)` | None | true | refused | `<a href="javascript:alert(4)">` |

   - **Under the sketch, each is still written as a link**, for example `[E0](<javascript&#58;alert(1)>)`.
   - **markdown-it-py's default catches them**, because it validates after decoding.
   - **docx4j and core-ts refuse none of the four, by their code.** `LinkTargets.scheme` at `30fb41024` returns null at the `#`, or finds no `:`. commonmark-java 0.30.0 writes the destination's `&` raw, and this writer is its port. docx4j itself was not run. This is Q3.
2. **The address comment is raw HTML.**
   - `to_markdown(addresses=True)` writes `<!-- w14:<paraId> -->` (`model/markdown/export.py` line 123). `address_of` (`model/content/addresses.py` line 137) takes the `w14:paraId` as it is.
   - A paraId of `--><script>alert(1)</script><!--` loaded with nothing skipped. It gave `<!-- w14:--><script>alert(1)</script><!-- -->`, a `<script>` outside the comment.
   - A footnote whose `w:id` is the same string (kept as written by the converter) gave `<!-- footnote:--><script>alert(1)</script>/0 -->`.
   - 0.1.0 writes the same: `<!-- w14:--><script>alert(1)</script><!-- -->` (measured).
   - **Over the corpus,** all 5,982 paraIds of the body paragraphs are eight hex digits, as `ST_LongHexNumber` (`hexBinary`, four octets) requires. This is Q8.
3. **`Outline.to_markdown()` writes document text unescaped.**
   - `model/content/reports.py` line 288 builds its lines as strings: `f"{'#' * level} {entry.text}"` and `f"- {label}"`. The README and the PyPI readme call it "the cheapest thing to show a model".
   - Measured on `create_package()`, it gave `# <script>alert(1)</script> heading` and `- [x](javascript:alert(2)) and <img src=x onerror=alert(3)>`. markdown-it-py's `commonmark` preset rendered `<h1><script>alert(1)</script> heading</h1>` and a live `<img onerror>`. For the same body, `body.to_markdown()` gave `# \<script\>alert(1)\</script\> heading`.
   - The same code is at `v0.1.0`.
   - **Over the corpus** (418 documents; `outline()` raised on 4), 1,951 of 10,726 entries, in 125 documents, hold one of `` \ ` * _ [ ] < > `` and would change under CR-003's `escape()`. 45 hold `<` followed by a letter. This is Q9.
4. **The readers' picture destination is a part's name.**
   - `_word_relative` (`model/markdown/export.py` line 159) writes the part name relative to `/word/`. A part outside `/word/` keeps its whole name, its leading `/` dropped.
   - An image part `/javascript:alert(1).png`, related as `../javascript:alert(1).png`, loaded, and `body.to_markdown()` wrote `![alt](<javascript:alert(1).png>)`. That is an image source, not a link, and no browser was asked what it does with one.
   - **The corpus's part names were not measured.**
   - `markdown()` writes data URIs, which are its own. This is Q10.
5. **`Range.hyperlink` accepts a control character, and the save fails late.**
   - The save raises lxml's `ValueError`, with no `code`.
   - That is a writer's validation, CR-013 Q2's kind (`text.not_writable`). It is not this CR's.
6. **0.1.0's `markup` view writes a comment's text raw** (`{>><script>alert(2)</script> text<<}`, measured). main escapes it (`\<script\>`), so only 0.1.0 has this defect (Q6).
7. **CriticMarkup in text.** A `{--z--}` in document text, or after this CR in an author, reads as a deletion to a CriticMarkup processor. CommonMark would accept `\{`, but the writer, commonmark-java's port, doesn't escape braces, and REQ-703 asks for the detail "escaped as any document text is". It is the WD's to decide, not this CR's (Q11).
8. **The importer writes no script link.** `insert_markdown("a [x](javascript:alert(1)) b [y](onenote:///z)")` writes the first as text (markdown-it-py's `validateLink`) and the second as a link. Not this CR's.

## 3. What docx4j built

### 3.1 The link rule (CR-056 section 3 and D1 to D4; `LinkTargets` at `84ee9741c`)

- **An allowlist, as POI chose (D1).** A target is written as a link only if it:
  - has no scheme (a relative path, a fragment, a UNC path);
  - is a drive path (a single letter, `:`, then `\` or `/`);
  - or has a scheme in {`http`, `https`, `ftp`, `mailto`, `file`, `news`, `tel`}.
  A null or empty target is safe.
- **The scheme is read as a browser reads it** (WHATWG URL parsing):
  - leading characters at or below U+0020 are ignored;
  - tab, LF and CR are dropped anywhere;
  - the scheme is everything up to the first `:` that comes before any `/`, `?` or `#`, compared lower-cased.
  - One reading is stricter than a browser's, on purpose. A malformed scheme (`java script:x`, a control character inside it) is refused, not resolved as a relative URL, "since a PDF viewer or a Markdown renderer need not read it the browser's way" (section 10).
- **Added schemes (D3).** The property `docx4j.convert.out.linkSchemes` adds schemes.
  - It is a comma-separated list; each entry is trimmed and lower-cased, a trailing `:` is tolerated and empty entries are ignored.
  - It is read on each call, so a test can set and restore it.
  - It can only add: a default can never be removed.
- **A refused target (D4)** gives the link's text with no link. Markdown writes the text runs; the HTML output's picture keeps its `img` with no `a` around it.
- **Logging.** One WARN per document names the first refused target; the rest go to DEBUG.
  - The target is shown through `excerpt`: at most 80 characters and `...`, controls as spaces.
  - The message is `refusedMessage`: "Link target refused (scheme not in docx4j.convert.out.linkSchemes): the text is written without a link: ...".
  - Internal targets (`#bookmark`) are untouched.
- **`LinkTargets`' members:** `LINK_SCHEMES_PROPERTY`, `defaultSchemes()`, `schemes()`, `isSafe(target)`, `scheme(target)`, `excerpt(target)`, `refusedMessage(target)`, and the package-private `normalise`.
- **In `WmlToMarkdown.linked`** (`84ee9741c`): `if (target != null && !LinkTargets.isSafe(target))`, then the message at WARN the first time and DEBUG after, and `target = null`.

### 3.2 The marker detail (CR-056 section 13; `30fb41024`; WD 0.33)

- `open(revision, sink)` writes `{++` or `{--` raw, then `detail(author, date, sink)`: the same `"author: "` or `"author, date: "` string, through `sink.text`.
- `comment(...)` writes `{>>` raw, the detail as text, the comment's text, then `<<}` raw.
- The string `open(revision)` is kept "for a code block, where nothing is markup".
- The markers' `<br>` is a constant.
- Out of CR-056: the OMML-to-LaTeX text, and the Markdown importer.

### 3.3 What doesn't apply here

- **Phase 2's style IDs as CSS identifiers** (`CssIdentifiers`, the `\3c ` guard, D5, D11). No HTML, no stylesheet.
- **HTML's and FO's link writers** (`HyperlinkUtil`), the HTML picture's click target (`AbstractWordXmlPicture`), docx4j-tika's `href`, and `TagClass`. No such outputs here.
- **`HtmlChunkImport`'s `setLinkSchemesChecked(false)`, and the inspector's `hyperlink.script`** (D8, phase 4, deferred there, defined in WD 0.32). No altChunk import and no inspector here.
- **The JVM-wide property.** This package has no properties file (no `os.environ` read anywhere under `docx4j_py/`, measured by `grep`). Section 4.3 and Q2 give the added schemes their place.
- **D9 and D10** (no advisory, no backport) are docx4j's. This package's equivalents are Q5 and Q6.

### 3.4 core-ts's port (`57b2f98`, `86e2e9d`)

- **`LinkTargets.mts` is docx4j's**, with the extra schemes an argument: `isSafe(target, extra)` and `schemes(extra)`. They come in through `MarkdownExportOptions.linkSchemes`, "docx4j's JVM-wide property ... added to the defaults, never removing one".
- **`WmlToMarkdown.linked`** writes a refused target's runs as text and "logs the first per export through the settable logger (setLogger)".
- **The marker detail** is `openMarker` and `detailText`, and "a code block keeps the marker as one string".
- **The tests.** docx4j's three tests were ported into `test/markdown.test.mjs`.
- **The drive letter** is tested as `/\p{L}/u` on a scheme of length 1, where JavaScript counts UTF-16 units, as Java does.
- **The excerpt** counts UTF-16 units, "as docx4j's".

## 4. Design here

Follow docx4j rule for rule and test for test, in core-ts's shape where docx4j's is the JVM's. Section 4.8 lists each departure.

### 4.1 `docx4j_py/model/text/link_targets.py` (new)

The module sits in the text model's package, beside the writer, as core-ts's `src/model/text/LinkTargets.mts` does. This package has no `convert/out`. `model/` is in `codegen/clean.py`'s `KEEP` whole, so no `KEEP` change is needed; CLAUDE.md's description of `model/text/` gains the file (Phase D).

```python
"""Which link targets the Markdown writer writes as a link (CR-015). docx4j LinkTargets (84ee9741c)."""

#: docx4j's default schemes (POI's set); docx4j defaultSchemes().
DEFAULT_SCHEMES: tuple[str, ...] = ("http", "https", "ftp", "mailto", "file", "news", "tel")

def schemes(extra: str | Iterable[str] = ()) -> tuple[str, ...]:
    """The defaults, then each added scheme trimmed, lower-cased, a trailing ':' dropped, an empty one ignored."""

def scheme(target: str | None) -> str | None:
    """The scheme as a browser reads it, lower-cased; None for none (relative, fragment, UNC, empty)."""

def is_safe(target: str | None, extra: str | Iterable[str] = ()) -> bool:
    """docx4j isSafe: safe as written and with its character references decoded as CommonMark decodes them (Q3):
    no scheme, a drive path, or a scheme in schemes(extra). None and "" are safe."""

def decode_references(target: str) -> str:
    """Numeric references and the 44 one-character HTML5 names decoded once, as CommonMark does (Q3); not html.unescape."""

def excerpt(target: str | None) -> str:
    """For a log line: at most 80 characters and '...', controls shown as spaces; 'null' for None."""

def refused_message(target: str | None) -> str:
    """The writer's log line for a refused target."""

def _normalise(target: str) -> str:
    """Leading controls and spaces removed; tab, LF and CR removed anywhere (docx4j normalise)."""
```

- **Functions in a module, not a class of static methods**, as `traversal` is. `None` stands for Java's null.
- **`schemes(extra)`** takes an iterable of strings, or one string, which is docx4j's comma-separated property value split on `,`. A Python `str` is itself an iterable of one-letter strings, so `link_schemes="onenote"` would otherwise add `o`, `n` and `e` and not `onenote`.
- **The drive letter** is a scheme of one character that `str.isalpha()` accepts and that lies in the Basic Multilingual Plane. Java's `Character.isLetter(char)` and core-ts's test see a letter outside the BMP as two UTF-16 units, so it is no drive letter there either.
- **Lower-casing is `str.lower()`.** It agrees with Java's `toLowerCase(Locale.ROOT)` on every scheme the allowlist can hold, which is ASCII.

### 4.2 The writer: a refused target is its text

In `_WmlToMarkdown._linked`, as `WmlToMarkdown.linked` at `84ee9741c`:

```python
target = runs[i].link
...
if target is not None and not is_safe(target, self.link_schemes):
    self._refused(target)          # WARNING the first time in this export, DEBUG after (section 4.4)
    target = None                  # the runs as text, no link
```

- **What a refused link writes:** its runs, each written as it would be outside any link, so formatting and pictures are kept. Measured under the sketch:
  - `click text H1`;
  - the field `field F1`;
  - the picture inside a refused link, `see ![P3 inside link](data:...) after`, its image kept and no link around it;
  - the cell, `|click cell T1|`.
- **A refused target is the whole target.** A link with both a relationship and an anchor, `javascript:alert(1)#frag`, is refused as one (measured: `click with anchor`).
- **Unchanged:**
  - an anchor alone (`#bm1`) has no scheme and stays a link;
  - so does a hostile anchor, which the renderer already keeps in one destination;
  - a code paragraph writes no link at all;
  - a picture's `a:hlinkClick` is never read, so nothing changes for pictures.
- **The benign targets** stay byte for byte as they are: section 2.1's seven, the corpus's 22 and the goldens'.

### 4.3 Adding schemes (Q2)

**Recommended: the writer's options, as core-ts has them.**
- **`MarkdownExportOptions.link_schemes: tuple[str, ...] = ()`**, beside `image_destination`. That class's own docstring says "the options here are how it is written", and a link's refusal is how a link is written: the run keeps its target, as docx4j's does.
- **A `link_schemes=` keyword**, defaulting to `()`, on:
  - `markdown()` and `markdown_blocks()`;
  - the readers: `Body.to_markdown()` and `markdown_budget()`, `Paragraph.to_markdown()`, `Table.to_markdown()`, `pkg.to_markdown()` and `pkg.markdown_budget()`, and the trial's pass-throughs (`TrialPackage`);
  - the module functions behind them (`body_markdown`, `paragraph_markdown`, `table_markdown`, `package_markdown` and the two budgets, through `_written` and `_alone`).
- **Add-only, by construction.** The keyword is handed to `schemes(extra)`, which starts from `DEFAULT_SCHEMES`. Nothing can remove a default, as D3 has it.
- **Per call, not per process.** A host that wants `onenote:` links everywhere passes the same tuple to every call, or wraps them. docx4j's test sets and restores the property; here a call without the keyword is the restore.

**The alternative the brief named: a `TextOptions` field.**
- **How it would work:**
  - `TextOptions.link_schemes`, set by keyword or `set("link_schemes", ...)`, beyond section 10 as `image_describer` is;
  - not in `OPTION_NAMES` or `to_dict()`, and copied by `copy()`;
  - the writer reads `result.options.link_schemes`, so `extract(pkg, link_schemes=[...])` reaches `markdown(result)` with no new parameter;
  - the readers pass it to `reading_of(**options)`.
- **What it costs:**
  - **A tool's arguments could add a script scheme.** `TextOptions.set` is reachable by name: `TextOptions.of(mapping)` and `extract(pkg, **options)` take it. CR-006's server passes a tool's text options to `extract(pkg, **options)` and relies on the library to refuse an unknown name (CR-006 section 3.1). So an agent's arguments, possibly steered by the very document it reads, could add `javascript` unless every server filters them out.
  - **It is a writer's concern in the text model's options.** Plain text never reads it, and neither does the JSON.
  - **It departs from core-ts** for no gain.
- **Rejected: an environment variable**, the nearest thing to docx4j's property.
  - It is process-wide: two documents in one server can't differ, and a test must set and restore it.
  - It is invisible at the call.
  - Nothing in `docx4j_py` reads the environment today.

### 4.4 Logging

- **The logger:** `logging.getLogger("docx4j_py.model.text.markdown")`, the writer module's `__name__`, as docx4j's is `WmlToMarkdown`'s class logger. Python's logging is configured by name, so core-ts's `setLogger` has no counterpart.
- **How often:** one WARNING per export (per `_WmlToMarkdown`, so per `markdown`, `markdown_blocks`, `MarkdownExporter.export` or reader call) for the first refused target, and DEBUG for each one after. The flag lives on the writer instance, so nothing is shared between threads. One CR-006 tool call therefore logs at most one WARNING.
- **The message** is `refused_message(target)`, docx4j's wording with the option's name for the property's: "Link target refused (scheme not in link_schemes): the text is written without a link: " and the excerpt. The excerpt shows controls as spaces, so a target can't forge a log line.
- **No handler is configured**, as a library should not. With none, Python's last-resort handler writes a WARNING to stderr, which keeps an MCP stdio server's stdout clean.
- **No `TextWarning`** in the result: the writer runs after the extraction, and the WD has no code for it (Q4).

### 4.5 Marker detail as text (Phase B)

As `30fb41024` and core-ts's `86e2e9d`:

```python
# _runs
if revision is not None:
    sink.node(md.HtmlInline("{++" if revision.type == "INSERTION" else "{--"))
    self._detail_text(revision.author, revision.date, sink)
self._linked(runs[i:j], sink)
if revision is not None:
    sink.node(md.HtmlInline(_close(revision)))

def _comment(self, comment, sink):
    sink.node(md.HtmlInline("{>>"))
    self._detail_text(comment.author, comment.date, sink)
    sink.text(comment.text, False, False, False)
    sink.node(md.HtmlInline("<<}"))

def _detail_text(self, author, date, sink):      # docx4j detail(author, date, sink)
    detail = self._detail(author, date)
    if detail:
        sink.text(detail, False, False, False)
```

- **Only the braces and operators are raw.** The detail joins the sink's text, so the renderer escapes it as any document text: `{++\[x\](javascript:alert(1)) \*e\* {--z--}: NEW++}` (section 2.2).
- **Every marked kind takes it:** insertions, deletions and moves in any block (a cell's included), comments, and replies, which take `_comment` too.
- **`_open` stays as it is for `_code_text`**, whose fence holds the string form (section 2.2).
- **Unchanged:**
  - `_same_marker`;
  - the type-only markers, byte for byte (`type` is the default);
  - plain text (REQ-702);
  - the readers, which have no `markup_detail`. This CR adds none.

### 4.6 The readers

**They are covered through the same writer.** `to_markdown` and the budgets are `markdown_blocks` over a reading (`_written`, `_alone`); `Table.to_markdown` is `markdown()` over `extract()`. Phase A's refusal and Phase B's text reach all of them with no reader of its own changed, except for the `link_schemes=` keyword (Q2).

**The tests** (section 6) assert it reader by reader, as CR-011 section 14.5 compared 61 `to_markdown` calls with docx4j's export.

### 4.7 The sinks found on the way (Phase C, as Q8 to Q10 decide)

1. **Address comments (Q8).** Two rules:
   - **At the source.** `address_of` and `_paragraph_by_para_id` (`model/content/addresses.py`) take a `w14:paraId` only when it is eight hex digits (`ST_LongHexNumber`); otherwise the block has its ordinal. Addressing and the comment then agree: `element_at` of a malformed paraId's ordinal finds the block. Over the corpus no paraId moves (5,982 of 5,982 are eight hex digits).
   - **As the guard behind it.** `address_comment` writes any character outside `[0-9A-Za-z:/._-]` as `%XX` (UTF-8), and each `-` that follows a `-` as `%2D`, so no address can close the comment.
   - A text model address with a malformed note or comment id (read-only, `address.read_only`) is written so too. `element_at` of an address holding `%` answers `address.not_found`.
2. **`Outline.to_markdown()` (Q9).** Each entry's text and label is escaped with CR-003's `escape()` (`model/markdown/export.py` line 132: `` \ ` * _ [ ] < > ``), the function kept for exactly that.
   - Over the corpus, 1,951 entries in 125 documents change, by a backslash before each such character.
   - `tests/agent/test_outline.py`'s two `to_markdown` assertions are checked, and any line that moves is listed in the implementation notes.
3. **The readers' picture destination (Q10).** `_word_relative`'s answer goes through `is_safe`. A refused one is `""`, the `![alt]()` REQ-703 already writes for a picture the package doesn't hold. It is logged with the links, under the same once-per-export WARNING.

### 4.8 Departures from docx4j, each with why

1. **The added schemes are a call's argument, not a JVM-wide property** (Q2). This package has no properties file, and an argument can't leak between documents. It is core-ts's choice.
2. **References decoded (Q3)**: no longer a departure. docx4j took the decoded check into `LinkTargets.isSafe` itself (Q3's answer, uncommitted there), so this package's `is_safe` follows it, decoding as CommonMark does.
3. **`schemes` takes a string or an iterable** (section 4.1), as docx4j's property (a string) and core-ts's argument (an array) both are.
4. **`excerpt` counts code points**, not UTF-16 units. It matters only for a log line's length.
5. **`DEFAULT_SCHEMES` is a tuple constant**, for `defaultSchemes()`, and there is no `LINK_SCHEMES_PROPERTY`.
6. **The logger is named, with no setter** (section 4.4), where core-ts has `setLogger`.
7. **Phase C's three sinks are this package's own surfaces** (CR-003's address comments, `Outline`, and the readers' part-name destination). docx4j writes none of them in its Markdown.

### 4.9 Not in this CR

- **CR-056's style-ID half, its HTML, FO and tika outputs, and its inspector kind** (section 3.3).
- **`Range.hyperlink` writing a script target.** Word keeps such a target and makes it inert, and the outputs refuse it. Its acceptance of control characters is section 2.5 item 5's.
- **`insert_markdown`** (section 2.5 item 8).
- **`to_api_script`**, which sends a hyperlink to `insert_xml` (its docstring). Not measured.
- **The block model's `link`, and `TextRevision`'s and `TextComment`'s author and date as data.**
- **Plain text's detail** (REQ-702).
- **CriticMarkup braces in text** (section 2.5 item 7; Q11).
- **A `markup_detail` keyword on `to_markdown`.**
- **The host's `image_destination`**, which is the host's own.
- **The YAML front matter.** It quotes and escapes every value and every key that isn't an identifier (`_yaml_string`, `_yaml_key`). Checked by reading, not probed.

## 5. Phases

| phase | what | size |
|---|---|---|
| A | `link_targets.py`; `_linked`'s refusal with `is_safe` (Q3); the logging; `link_schemes` on `MarkdownExportOptions`, `markdown()`, `markdown_blocks()` and the readers (Q2); `LinkTargetsTest` and `MarkdownHostileLinksTest` ported, and this package's tests (section 6.2); the CHANGELOG's 0.2.0 entry; a paragraph in README's "Markdown, coarse and fine" on what is refused and `link_schemes`, as docx4j's Getting Started gained one | ½ day |
| B | `_runs` and `_comment` with the detail as text; `MarkdownHostileMarkupTest` ported, and this package's tests (moves, replies, a cell, the date, the code block); the CHANGELOG | ¼ day |
| C | as Q8 to Q10 decide: the address rule and guard, the outline's escape, the readers' picture destination; their tests; the CHANGELOG | ½ day (all three) |
| D | this CR's implementation notes (what landed, the numbers, the departures); `tests/README.md`'s map; CLAUDE.md's `model/text/` line; CR-006's line if Q7 is taken | ¼ day |

A and B are independent, and each is committed on its own. C waits for the user's answers. No phase changes a saved part: CLAUDE.md's Word-check rule names marshalling, the importer and the inserts, and none of them moves, so there is **no Word check**. Acceptance is regenerated only to prove it unchanged.

## 6. Tests and the gate

### 6.1 docx4j's tests, ported test for test, their strings kept

**`tests/text/test_link_targets.py`**, from `LinkTargetsTest` (`84ee9741c`), five tests:
- **`test_safe_targets`:** the 18 targets, from `None` and `""` to `"  http://example.org/leading-spaces"`, each `is_safe`.
- **`test_refused_targets`:** the 12, `"javascript:alert(1)"` to `"blob:x"`, `" \u0001java\tscri\npt:alert(2)"`, `"C:"` and `"java script:alert(1)"` among them.
- **`test_scheme_as_a_browser_reads_it`:** the six readings (`"javascript"`, `"http"`, `None` three times, `"c"`).
- **`test_added_schemes`,** for `propertyAddsSchemes`:
  - the extra `" OneNote:, ms-word ,,notes"`, given as docx4j's one string and again as a list;
  - the defaults all present, `onenote` added, `onenote:///x`, the `ms-word:ofe` target and `Notes://server/db.nsf` safe;
  - `javascript:alert(1)` still refused ("a default is never removed");
  - then `onenote:///x` refused with no extra, where docx4j's test asserts "restored".
- **`test_excerpt_for_the_log`:** `"java scri pt:alert(1)"`, and 83 characters for 200 `x`.

**`tests/text/test_markdown_hostile.py`**, from `MarkdownHostileLinksTest` (`84ee9741c`) and `MarkdownHostileMarkupTest` (`30fb41024`):
- **`test_no_script_scheme_destination`** and **`test_benign_links_kept`**, over docx4j's six targets.
  - **The document is built in memory**, with `create_package()` and `Range.hyperlink`, since U+0001 can't be saved. Each paragraph's text is `"text for "` plus the target trimmed as Java's `trim()` trims (every character at or below U+0020; section 2.3).
  - **The assertions are docx4j's**, plus one this package adds, as docx4j's HTML test was narrowed to `href` values: every destination in the output, its references decoded, is `is_safe`. docx4j's own check of the obfuscated target is vacuous here (section 2.3).
- **`test_no_raw_markup_from_any_document_value`** over the three `markup_detail` values, and **`test_benign_markers_unchanged`**.
  - **The document is built with `MiniDocx`** (`tests/text/mini_docx.py`) from XML:
    - T1 to T3 as text;
    - T4 as a link's text, to `https://example.org/ok`;
    - T5 as a picture's `wp:docPr` name and description, with docx4j's 1x1 PNG;
    - T6 as an insertion's author;
    - T7 and T8 as a comment's author and text.
  - **docx4j's strings and its `unescaped` helper** are kept.
- **From `HostileDocx`** (`df014ec3d`, `84ee9741c`), in Markdown:
  - its eleven `BENIGN` targets, each kept as a link;
  - its `CUSTOM_LINK` (`onenote:///C:/notes/section.one#page`), refused by default and linked under `link_schemes=["onenote"]`, which is docx4j's `propertyAddsAScheme` (an HTML test there).

### 6.2 This package's own

**In `tests/text/test_link_targets.py`:**
- **Q3's reference forms** are refused by `is_safe`, and docx4j's kept forms (`R&D/report.html`, `page.html?a=1&amp;b=2`, `https&#58;//example.org`, `&colon`, `page.html#&#58;`) are kept, as its `characterReferencesDecodedAsACommonMarkRendererDecodesThem` has them.
- **A string `link_schemes`** is split on commas, never iterated as characters.
- **A letter outside the BMP** before `:\` is not a drive letter.

**In `tests/text/test_markdown_hostile.py`:**
- **Section 2.1's link paths:**
  - a complex and a simple HYPERLINK field;
  - `javascript:...#frag`;
  - a picture inside a refused link (its image kept);
  - a cell, a footnote's definition and a text box's paragraph.
- **The logging,** by `caplog` on `docx4j_py.model.text.markdown`:
  - one WARNING per export, naming the first target's excerpt, and DEBUG for the rest;
  - a second export warns again.
- **Section 2.2's marker paths:** a deletion, both moves, a cell, a comment and a reply, the date under `author_date`, and the code block keeping the string form inside its fence.
- **What doesn't move:**
  - plain text, byte for byte;
  - the `type` detail's Markdown, byte for byte;
  - the conformance goldens.

**In `tests/content/test_markdown_safety.py`:**
- **The readers, one by one:** `pkg.to_markdown()`, `body.to_markdown()`, `body.markdown_budget()`, `paragraph.to_markdown()`, `table.to_markdown()`, and a `dry_run` trial's `to_markdown`. Each:
  - refuses a `javascript:` link set with `Range.hyperlink`;
  - keeps an `onenote:` link under `link_schemes=("onenote",)`;
  - writes `{++NEW++}` in `markup`.
- **Phase C's, if taken:**
  - the hostile paraId given its ordinal, both in the comment and in `element_at`;
  - the guard on a note's address;
  - the outline's escape;
  - the picture at `/javascript:alert(1).png` written `![alt]()`.

### 6.3 The gates

- **The new tests, and the suite**, fast and full (`.venv-fork/bin/python -m pytest`, and with `-m "not slow"`).
- **`tests/text/conformance.tsv` unchanged** (`test_conformance.py`), with its goldens 12 of 12.
- **`tests/test_readme.py`**, for the README paragraph.
- **The corpus census of section 2.4, re-run after A and B**: 0 of 422 exports move in either view, and the 22 targets are all kept. Phase C's moves are counted against section 2.5's figures (1,951 outline entries; 0 paraIds).
- **The byte check against docx4j at `30fb41024`,** CR-011 section 14.5's method: `git archive` into the scratch directory and an offline build.
  - The 277 documents' 1,385 exports must still be identical.
  - So must the hostile documents of section 6.1, saved (all but U+0001), Q3's forms included, once docx4j's fix is committed.
  - **Not run for this draft.**
- **The rest, each unchanged by construction:** `codegen/generate.sh --check`; `scripts/roundtrip.py`; and acceptance (`scripts/acceptance.py`, then `scripts/compare_acceptance.py` against a copy made first, 0 differ). The acceptance artefacts are `.docx` only.

## 7. Risks

1. **Links lost.** A document whose links use `onenote:`, `ms-word:`, `notes:`, `sms:` or `callto:` loses them in Markdown.
   - None of the corpus's 22 targets is refused, and core-ts's 61 aren't either. docx4j's own risk is the same (CR-056 section 7).
   - `link_schemes` restores a scheme, and the WARNING names the first target refused.
   - The default set stays POI's and docx4j's (Q1).
2. **A renderer reads a destination as the scheme reader doesn't.**
   - The reader is a browser's, and Q3 adds the parser's decoding of references.
   - Percent-encoding is inert in a scheme: a browser doesn't decode `%3A` there, and markdown-it-py encodes a tab as `%09`.
   - A renderer that percent-decodes a destination before linking would see more. None was measured.
3. **Parity with docx4j's bytes.** docx4j took Q3's decoded check into its own rule (uncommitted there), so parity holds on those forms too once it is committed (section 6.3).
4. **The added schemes reaching an agent.** Under Q2's recommendation they are never a text option, so no tool argument can add one unless a server wires it. Under the alternative, every server must strip them (section 4.3).
5. **Phase C moves outputs callers may pin.** The outline's Markdown changes in 125 of 418 corpus documents (1,951 entries). The address rule moves no corpus address. The picture rule's effect on the corpus is unmeasured. Each move goes into the CHANGELOG's 0.2.0 entry.
6. **A sink this reading missed.** The census of section 2:
   - the one `md.Link`;
   - the raw `HtmlInline`s (the markers and the constant `<br>`);
   - the two string-built Markdown writers (`Outline.to_markdown` and the address comments);
   - the image destinations.

   A new writer that builds Markdown as a string, or makes a `Link`, should go through the commonmark port and `is_safe`. The tests can't enforce that, as CR-056 says of its own.

## 8. Questions

1. **Q1. The rule as docx4j's?**
   - The seven schemes (`http`, `https`, `ftp`, `mailto`, `file`, `news`, `tel`), plus no scheme and a drive path.
   - The browser's reading.
   - A malformed scheme refused.
   - Added schemes add-only: trimmed, lower-cased, a trailing `:` tolerated.
   - A null or empty target safe.
   - Internal targets untouched.

   *Recommended: yes, verbatim (D1, D3),* with section 4.1's two Python readings: a string given as `link_schemes` is a comma-separated list, and a drive letter is a BMP letter. It keeps docx4j's and core-ts's answers on every target their tests name.
2. **Q2. Where do added schemes go?**
   - (a) A `TextOptions` field, as the brief suggested.
   - (b) `MarkdownExportOptions.link_schemes` with a `link_schemes=` keyword on `markdown()`, `markdown_blocks()` and the readers, core-ts's choice.
   - (c) An environment variable, docx4j's property's nearest form.

   *Recommended: (b).*
   - It is how a link is written, not what text is read: plain text and the JSON never use it.
   - It is core-ts's (`57b2f98`), beside `image_destination`, the writer's other host hook.
   - It is not reachable by name through `TextOptions.set` and `TextOptions.of`, which CR-006's server feeds from a tool's arguments, where an agent could add `javascript` (section 4.3).

   (a) works, at the cost of every server stripping it. (c) is process-wide and invisible at the call.
3. **Q3. Character references.** Should a target also be refused when it is refused once its references are decoded (`html.unescape`)? Today `javascript&#58;alert(1)`, `&#106;avascript:...`, `javascript&colon;...` and `&#x6A;avascript:...` pass docx4j's rule. They come out as `<a href="javascript:...">` from a CommonMark renderer with no URL check (section 2.5 item 1).

   **docx4j's answer (2026-10-23).** docx4j's session confirmed the finding and fixed it in docx4j. The fix is uncommitted, waits on the user's word there, and is recorded in docx4j CR-056 section 14; cite it by commit once it is committed. Its rule:
   - `LinkTargets.isSafe(target)` is true only if the target is safe as written **and** safe with its character references decoded as CommonMark decodes a link destination. It is **one rule for every output** (HTML, FO, Markdown, tika), not a Markdown-only check.
   - **What is decoded:** decimal and hex numeric references (U+FFFD for 0, for a value above 0x10FFFF, or for a surrogate), and the 44 HTML5 named references that stand for one ASCII character (`&colon;`, `&sol;`, `&num;`, `&quest;`, `&Tab;`, `&NewLine;`, `&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&lpar;`, `&rpar;` and the rest).
   - **What is left as written:** other named references, and a bare `&`. Decoding happens once, so `&amp;#58;` stays `&#58;`.
   - **Python's `html.unescape` won't do:** it decodes every HTML5 name and also names without a semicolon from the legacy set, while CommonMark needs the semicolon. The port decodes as CommonMark does (or is knowingly stricter).
   - **Refused as written**, since the scheme is everything up to the first `:` before any `/`, `?` or `#` and must be in the set: `&#106;avascript:x`, `&unknown;avascript:x`, `a&b:c/d`.
   - **Kept:** `R&D/report.html`, `page.html?a=1&amp;b=2`, `https&#58;//example.org`, `&colon`, `page.html#&#58;`.
   - **Oracle tests:** `docx4j-core-tests`' `LinkTargetsTest.characterReferencesDecodedAsACommonMarkRendererDecodesThem` and `docx4j-markdown`'s `MarkdownHostileLinksTest.noEntityEncodedScriptSchemeDestination`, ported once committed.

   *Recommended: yes, as docx4j's rule*: `link_targets.is_safe` takes the decoded check itself, one rule as docx4j's, rather than a Markdown-only `is_safe`. Section 4.8 item 2's departure goes, and parity with docx4j's bytes holds on these forms too. core-ts's session reported the same hole there (`57b2f98`) and is putting the decoded check to the user.

   Escaping `&` in destinations instead would move every link with a query string, away from commonmark-java's bytes.
4. **Q4. Logging as docx4j's?**
   - One WARNING per export, for the first refused target, and DEBUG after.
   - On the logger `docx4j_py.model.text.markdown`, with docx4j's message naming `link_schemes`.
   - No `TextWarning` in the result.

   *Recommended: yes.* A `TextWarning` would need a WD code that doesn't exist, and the writer runs after the extraction that owns the warnings.
5. **Q5. Does this go in 0.2.0?** *Recommended: yes, A and B (and C as decided), before 0.2.0 is released.*
   - It is a security fix.
   - 0.2.0 isn't released: its CHANGELOG entry still waits for CR-013's Phases B and C.
   - No corpus output moves under A and B.
6. **Q6. A 0.1.x fix release?** *That is the user's call.* Measured in 0.1.0's Markdown (section 2.1, 2.5):
   - a script-scheme link;
   - a target's `)` breaking out into raw HTML that runs on load;
   - a comment's text written raw in `markup`;
   - the address comment;
   - `Outline.to_markdown()` unescaped.

   The second and third are fixed on main by CR-011's writer; the rest are this CR's. 0.1.0's exporter is CR-003's (`model/markdown/export.py` at `d234949`), not this writer, so a 0.1.x fix would be its own patch, not a cherry-pick. It would need:
   - `is_safe` on its targets;
   - its destinations put in angle brackets with `<`, `>` and `\` escaped;
   - comment text escaped;
   - the address rule and the outline's escape.

   The `w:object` order fix (CR-012 section 12.15) is another 0.1.0 defect that could ride with it: its two schema patches apply at the tag over docx4j-xsdata 26.2.1 (CR-012 section 12.16). *Recommended: if a 0.1.x is made, take both. Otherwise 0.2.0 supersedes 0.1.0, and the CHANGELOG's 0.2.0 entry names what 0.1.0 writes.*
7. **Q7. CR-006's MCP server.** Should CR-006 (`../docx4j-mcp-py`, proposed; no such checkout exists yet) note that its Markdown inherits the fix? *Recommended: yes, one line in CR-006 section 3.1 when this CR is implemented:*
   - `read`'s Markdown, and `docx_to_markdown`'s, refuse script-scheme links and write marker detail as text;
   - `link_schemes` is the deployment's configuration, never a tool's argument.

   The line is Phase D's, not this draft's.
8. **Q8. The address comments (section 2.5 item 2).** Take Phase C's address rule (a paraId only when it is eight hex digits) and its guard (`%XX`)? *Recommended: yes.*
   - It is the same class of defect: a document value written as raw HTML.
   - It sits in `to_markdown(addresses=True)`, the agent surface's main read and CR-006's `read`.
   - No corpus address moves.
9. **Q9. `Outline.to_markdown()` (section 2.5 item 3).** Escape each entry's text with `escape()`? *Recommended: yes.*
   - It writes document text as Markdown with no escaping at all: a heading's `<script>` or a paragraph's `<img onerror>` runs on render.
   - The README offers it to models.
   - The cost is 1,951 entries in 125 of 418 corpus documents gaining backslashes, in 0.2.0, whose outline already moved under CR-011.
10. **Q10. The readers' picture destination (section 2.5 item 4).** Send `_word_relative`'s answer through `is_safe`, writing `![alt]()` when refused? *Recommended: yes, as the smallest of the three.*
    - An image source with a script scheme is inert in browsers as far as is known; that is unmeasured.
    - REQ-703 already keeps a document-supplied target out of an image's destination.
    - The check is a line.
11. **Q11. The WD.** REQ-703's links bullet says `[text](target)` and nothing of a target's scheme. docx4j and core-ts refuse with no WD item, where 0.33 did write the marker rule down. Should the portfolio's session be asked for a WD item? It would be REQ-703's links bullet and section 11: a link whose target's scheme isn't allowed is its text, the scheme read as a browser reads it, with Q3's decoding if taken. Section 2.5 item 7, braces in text, could go to it as a note. *Recommended: yes, sent by the user. This CR doesn't wait for it.* No conformance row moves: no fixture holds a refused link (section 2.4).

## 9. Effort

| phase | days |
|---|---|
| A, the links | ½ |
| B, the markers | ¼ |
| C, the three sinks found (Q8 to Q10) | ½ |
| D, the records | ¼ |
| **total** | **1** for the two fixes (A, B and D); **1½** with Phase C |

That agrees with the caller's guess of a day for the two fixes. The larger part is the tests, not the code:
- docx4j's four test classes ported string for string;
- the readers one by one;
- the census and the byte check against docx4j at `30fb41024`, which this draft didn't run.

The code is small: `link_targets.py`, about 100 lines; `_linked`, 10 lines; `_runs` and `_comment`, 15 lines; the keyword through eight call sites.
