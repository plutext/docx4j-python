# CR-029: docx4j CR-069 in the text model --- All Caps in capitals by the run's language, a drop cap joined to its word, `<sup>` and `<sub>` in Markdown, and `w:dstrike` struck

**Status:** Proposed 2026-11-02, at the user's word ("draft those 2 CRs"); **decided 2026-11-02**, Q1, Q2, Q4 and Q5 as recommended and Q3 as docx4j does (section 8); **amended 2026-11-02**: runs at a deleted-mark join merge as docx4j's do too (section 4.5a); **amended 2026-11-04** for WD 0.80, the user's ruling of REQ-607's chain: `w:themeFontLang` is struck from the language chain, which ends at a new option, `default_language` (section 4.3b). It ports the text-model half of docx4j CR-069
(WD REQ-607, REQ-521, REQ-710, and REQ-703's strikethrough) to `model/text/`, phase for phase:
- A, `w:caps` in capitals by the run's language, `w:smallCaps` as stored unless `small_caps: upper` (docx4j phase 1);
- B, a drop cap joined to its word, with `paragraph.drop_cap_joined` and `drop_caps` (phase 2);
- C, `<sup>` and `<sub>` in Markdown (phase 3);
- D, `strike` taking `w:dstrike` (phase 4; that phase's inspector kind, `text.formatting_join`, is named and left out);
- E, the records.

The questions are in section 8, each with a recommendation. **Decided 2026-11-02** (the user: "CR-029's questions Q1 to Q5 as recommended."; section 8): the case table at Unicode 15.0, generated and committed; a drop cap's letter merged into its word's first run as docx4j does (Q3, the user's ruling the same day); docx4j's member names; after CR-027 Phases A and D. Nothing built.

*The number:* CR-027 is the latest CR here. CR-028 is the number proposed for the port of docx4j CR-068 (manual redline
to tracked changes).

**Depends on:**
- **docx4j CR-069** (complete 2026-11-01, released in 17.3.5; tag `docx4j-17.3.5` at `64fae4cd0`, each commit below
  checked to be its ancestor):
  - phase 1, `75b9bdb54`, with `8ebd3d7de` (`CaseMapping`'s Unicode note); phase 2, `15c66769f`; phase 3, `5a2ce31ee`;
    phase 4, `7326296b1`;
  - phase 0, `993e385e7` (the `dpp/` suite), which CR-027 Phase A copies; phase 5, `c102d2ef6`, docs only.

  What is ported is the code and tests at the tag. Where the CR's prose is narrower, the code is followed (section 2.5).
- **The WD 0.56** (`../docx4j-portfolio/docs/text_extraction_specification_wd.md` at portfolio `1eb5791`), which matches
  docx4j 17.3.5: REQ-521, REQ-607, REQ-703, REQ-710, REQ-801, section 9.2's `text.formatting_join`, section 10's order
  and REQ-1002's preset. WD 0.57 to 0.60 change none of them.
- **CR-027** (decided 2026-11-02, Q1 to Q5 and Q7 as recommended, Q6 on 2026-11-01; nothing built):
  - Phase A copies docx4j's suite at the tag, `dpp/` included, with the harness's changes;
  - Phase D builds `profile: word-text` with the preset's 24 values that predate CR-069. This CR adds the other two.
- **This repository:**
  - CR-011's text model and its conformance harness (`tests/text/conformance.py`, `adapter.py`);
  - CR-018's readers in `model/values.py` (`is_on`, `vertical_alignment`, `in_enumeration`);
  - the `office_offset` invariant of CR-011 section 12.6 and CR-013 Q11, and the merge rule of CR-011 section 16.5 (Q3).
- **Counterparts:** docx4j-core-ts CR-012 section 24 (registry `core-ts/CR-012.24`, planned). Its D1 to D5 were decided
  2026-11-02: Unicode 15.0, docx4j's member names, landing with its suite copy and `profile`.

## 1. Why

Measured 2026-11-02 at `ed12ae7` on the `dpp/` fixtures at the tag (read from a scratch copy; nothing here changed):

- **Case.** The docx-tool-identifier's finding holds: on `dpp/bisect/C05.docx` this package gives `Roberts` where
  docx4j gives `ROBERTS`. The full lines are:
  - here: `C05_ Jackson, J., dissenting; Roberts, C.J.`;
  - docx4j 17.3.5: `Jackson` and `ROBERTS` (small caps stored, caps in capitals);
  - Word's Save As text: `JACKSON, J., dissenting; ROBERTS, C.J.`.

  Likewise `caps/caps-lang7d.docx` gives `ﬁne café ǆ` here and `ﬁNE CAFÉ Ǆ` in Word's save. `w:caps`, `w:smallCaps`,
  `w:vertAlign` and `w:lang` are read nowhere in `model/text/`.
- **Python cannot use its own `str.upper`.** It is the full mapping. Measured over every code point on 3.12, 3.13 and
  3.14, it expands 102 of them (`ß` to `SS`, `ﬁ` to `FI`). Taking its result only where it is one code point still
  differs from the simple mapping at 27: the Greek forms with ypogegrammeni, where U+1FB3's simple uppercase is U+1FBC
  and `str.upper` gives `ΑΙ`. It also ignores the language.
- **Drop cap.** `bisect/C02.docx` gives two blocks here, `T` at `body/1` and `he C02_ modern ...` at `body/2`, with no
  warning. docx4j gives one block, `The C02_ modern ...` at `body/2`, and warns at `body/1`.
- **Superscripts.** `bisect/C07.docx`'s Markdown here is `C07\_ Revenue 48,2103 in the period.`; docx4j's is
  `48,210<sup>3</sup>`. The block model here has one run where docx4j has three.
- **Double strikethrough.** `extractor.py` line 1268 reads `w:strike` alone, so a double-struck run has no `~~`.
- **Runs at a deleted-mark join** (measured 2026-11-02 at `80e5910`, after the user's ruling on Q3). The probe has
  three pairs of paragraphs. Here, where two paragraphs are one block because the view deletes the first's mark
  (CR-043's join), two runs of the same formatting stay two runs: `["J1_Alpha ", "beta J1_END"]` in the accepted view
  (a deleted mark), and `["J2_Gamma ", "delta J2_END"]` in the original view (an inserted mark). docx4j 17.3.5, run by
  its session on the released jar, gives one run each, `"J1_Alpha beta J1_END"` and `"J2_Gamma delta J2_END"`. Its
  `VERSION_17_3_6` gives the same. A bold run beside a plain one stays two runs in both engines. The block addresses
  and boundaries agree. Only the runs differ, because this package merges by `office_offset` contiguity within one
  `w:p` (CR-011 section 16.5) and docx4j's `addText` on kind, flags, revision and formatting alone.
- **The suite.** Once CR-027 Phase A copies `dpp/`, these rows pend on statements this CR builds:
  - 33 of the 34 token rows (REQ-607 or REQ-710);
  - 2 of the 8 findings rows (REQ-521; the other 6 are the inspector's);
  - both goldens;
  - 16 of the 22 `word-saves.tsv` rows (REQ-521, REQ-607 or REQ-710).

  The two hand-cut goldens differ from today's output by exactly this CR's three changes and nothing else: the
  drop cap's two lines, `Roberts`, and in Markdown C06's and C07's missing tags. C03's `~~1.316~~ 1.335`, C04's table and
  C08 to C10 already agree.

## 2. What docx4j built

### 2.1 Case mapping (phase 1, `75b9bdb54`; REQ-607)

- **`TextRun`** gains `isCaps()`, `isSmallCaps()`, `getVerticalAlignment()` (`BASELINE`, `SUPERSCRIPT`, `SUBSCRIPT`)
  and `getLanguage()`. All four are set in `Extractor.formatOf`, for text runs only:
  - caps: the run's own `w:caps`, else its styles' (the resolver's effective `w:rPr`, toggles XORed). Unlike bold,
    italic and strike, a paragraph style's caps is not cancelled.
  - small caps: the same, and false where caps is true (caps beats small caps).
  - vertical alignment: the run's own `w:vertAlign`, else its styles'. So a note reference, which is not a text run,
    is baseline whatever its style raises (REQ-710).
  - language, on a caps or small-caps run only: the run's own `w:lang/@w:val`, else its styles' together with the
    document defaults', else the settings' `w:themeFontLang/@w:val` (struck by WD 0.80: section 4.3b), else `en-US`
    (`CaseMapping.DEFAULT_LANGUAGE`). A
    blank value counts as none. A `w:lang` alone therefore never splits runs.
  - `sameFormat` and `copyFormat` take all four, so runs merge only where they agree.
- **`CaseMapping`**:
  - `upper(text, language)` maps each code point to `Character.toUpperCase(int)` (Unicode's simple mapping), and `i`
    to `İ` (U+0130) where the language's primary subtag (split at `-` or `_`, case-insensitive) is `tr` or `az`;
  - a code point whose uppercase would change its UTF-16 length is kept;
  - `isUpper(run, options)` is caps, or small caps under `small_caps: upper`;
  - `written(...)` gives a run's, a block's or a comment's text as the writers write it.
- **The writers.**
  - `PlainText` writes a text run and a comment's text through `CaseMapping.written`. Field codes and list labels stay
    as stored.
  - `WmlToMarkdown` does the same for a text run (a code span included), a code paragraph's text runs and a comment's
    text.
- **`TextOptions.small_caps`** (`asis` by default, `upper`) comes after `symbols` in `NAMES`. `WORD_TEXT_VALUES` sets
  `upper`.
- **The Unicode version is the running JDK's**, unpinned: 10.0 on Java 11, 15.0 on the OpenJDK 21.0.12 its tests ran on
  (CR-069 section 11.2 item 7). The tests and the `dpp/` rows map only characters whose simple uppercase has been
  stable since before 10.0.

  Measured here 2026-11-02, with OpenJDK 21.0.12's `Character.toUpperCase`, Perl 5's Unicode 16.0.0 tables, and Unicode
  10.0's `UnicodeData.txt` from an OpenJDK 11 source tree:
  - Unicode 15.0 (JDK 21) maps 1,450 code points.
  - 16.0 maps those and 27 more, and none less: U+019B and U+0264 (letters since Unicode 1.1, whose capitals 16.0
    added), U+1C8A, U+A7CD, U+A7DB, and Garay's U+10D70 to U+10D85. So U+1C8A is not the only divergence.
  - 10.0 maps 129 fewer than 15.0, Georgian Mkhedruli among them (Mtavruli capitals since 11.0).
  - No pair in 10.0, 15.0 or 16.0 changes a UTF-16 length.

  CI runs Python 3.12, 3.13 and 3.14, whose `unicodedata` is 15.0.0, 15.1.0 and 16.0.0.
- **The test** is `CaseMappingTest` (10).

### 2.2 The drop-cap join (phase 2, `15c66769f`; REQ-521)

- **`Extractor.blocks`** remembers a frame paragraph just read. The paragraph next takes it (`joinDropCap`), and
  `paragraph()` now returns its block, or null where it makes none.
- **A frame** (`dropCapFrame`) is a paragraph with all of these:
  - `drop_caps: join` is set;
  - it is one `w:p`, not a chain joined by a deleted mark;
  - its own `w:pPr` holds no `w:sectPr`;
  - its block's text holds no character for which `Character.isWhitespace` or `Character.isSpaceChar` is true (so
    U+00A0 counts). An empty frame, such as a picture's drop cap, passes;
  - its effective `w:framePr/@w:dropCap` is `drop` or `margin`. The effective `w:pPr` is the view's, through the
    resolver, so a style's frame counts.
- **What clears it.** The frame is cleared by any item that is not a paragraph or a marker. A marker is a range start or
  end, a bookmark, `w:permStart`, `w:permEnd` or `w:proofErr`. What the view removes is passed over.

  It is also cleared by a paragraph that goes into a table's first cell (CR-043's table join). A paragraph that makes
  no block takes the frame and leaves it as it is.
- **The join** (`joinDropCap`). The frame's block must still be in the sink. Then:
  - the frame's runs come first, then the block's;
  - the frame's last run and the block's first become one run where both are text, neither is math, and they agree in
    hidden, placeholder, revision and format;
  - the frame's fields go first, the revision counts are summed, and the block is hidden only if both were;
  - the block, with the second paragraph's address and properties, replaces the frame in the sink, so what the frame
    anchored (REQ-518) still follows it;
  - the inspector's record maps the frame's `w:p` to the joined block.
- **The warning** is `TextWarning.DROP_CAP_JOINED` (`paragraph.drop_cap_joined`) at the frame's address. Its selector is
  the name of the part holding the paragraphs (`/word/document.xml`).
- **`TextOptions.drop_caps`** (`join` by default, `asis`) comes after `toc`. The preset sets `asis`, Word's own text.
- **The test** is `DropCapJoinTest` (8). Its seventh case, `theInspectorFindsTheFrameInTheJoinedBlock`, is the
  inspector's.

### 2.3 `<sup>` and `<sub>` in Markdown (phase 3, `5a2ce31ee`; REQ-710)

- `WmlToMarkdown`'s `InlineSink` keeps the run's alignment beside its bold, italic and strike, and flushes when any
  changes. Adjacent runs alike share one pair of tags: `x<sup>ab</sup><sup>**c**</sup><sub>d</sub>`.
- `flush` writes the tags as raw inline HTML outside the emphasis: `48,210<sup>**3**</sup>`, `C<sub>~~*x*~~</sub>`.
  CommonMark would not open `**` after a digit and before `<` (WD 0.55's NOTE).
- A monospace run is a code span inside its own pair of tags: `` call <sup>`f()`</sup> here ``.
- A code paragraph gets no tags: its runs are written on the baseline (`x = 103`).
- A run's text is split at `\n` into lines, and each line is wrapped on its own, with a hard break between.
- A note reference is a reference run, never wrapped (`Body[^1]`).
- Document text cannot form the tags, because the renderer escapes its `<`. Plain text is unchanged (`48,2103`).
- Not ported: `Docx4J.toHTML`'s `sup` and `sub` elements (CR-069 D4). That is docx4j's converter, not the WD's HTML
  serialization, and this package writes no HTML.
- **The test** is `MarkdownVerticalAlignTest` (6), each Markdown also rendered through commonmark-java.

### 2.4 Strike and the formatting join (phase 4, `7326296b1`)

- **`TextRun.isStrike()`** is true where either `w:strike` or `w:dstrike` is on (the run's own, else its styles'), and
  no paragraph style has `w:strike` or `w:dstrike` on. WD 0.56's REQ-703 makes this the block model's one strike flag.
  Markdown writes `~~` for a double-struck run. The test is `FormattingJoinTest.aDoubleStruckRunIsStruck`.
- **`text.formatting_join`** (WD 9.2) is an inspector kind (`InspectWalk.formattingJoins`). Its other seven tests and
  `inspect_oracle.py` belong to the inspector.

### 2.5 Where CR-069's prose is narrower than its code

CR-069's own implementation sections (11.2, 12.2, 13.2, 14.2) record each of these. The port follows the code:
- Section 3.3 says "a body-level paragraph" and keeps the letter as "a run" of its own. The code covers every story and
  cell (WD 0.54) and merges the letter where the formatting agrees.
- Section 3.3's frame is "at most one word" and its "next sibling" a paragraph. The code uses the whitespace test above,
  leaves a frame that ends a section, passes over markers and what the view removes, and leaves the frame where the
  paragraph next makes no block.
- Section 3.2's language chain includes "the paragraph's". The code reads the run's own `w:lang`, then the effective
  style `w:rPr`. The paragraph mark's `w:rPr` plays no part (section 10.4).
- Section 3.1 folds the flags "as the other flags are"; the code gives case and alignment no paragraph-style cancelling.
- Section 3.2's "no tags inside a code span" is, in the code, tags around it. Section 4's Markdown capitals came in
  phase 1.

## 3. What exists here (measured 2026-11-02 at `ed12ae7`)

- **`extractor._run`** (line 1255) sets `bold`, `italic`, `strike` and `monospace` through `_on_of`. Here `styles` is
  `_style_rpr_of(r_style, p_style)`, the resolver's `get_effective_rpr` with the document defaults and the toggles
  XORed, and `paragraph` is the paragraph style's alone. That is docx4j's `formatOf` before CR-069.
- **The run merge** (`_add_text`, line 1459) merges on kind, hidden, placeholder, revision and `same_format`, as docx4j
  does. It also requires the two texts to be contiguous in one source paragraph's `Paragraph.text` (CR-011 section
  16.5, from core-ts CR-012 4.4). So runs from two `w:p` never merge while both carry an `office_offset`, where
  docx4j's merge ignores the source.
- **`TextRun`** (`result.py`, line 74) is a slotted dataclass. `same_format`, `copy_format` and `to_dict` list its flags
  by name. The JSON uses snake_case member names (`display_math`, `note_type`, `prog_id`).
- **`blocks()`** (line 847):
  - it builds the deleted-mark chains through `_next_paragraph`, which passes over `_is_marker` items and what the
    view removes;
  - `_paragraph` (line 947) returns None in every case;
  - nothing reads `w:framePr`.
- **`options.py`** has 26 `OPTION_NAMES` and its `_CHOICES` table. There is no `profile` yet (CR-027 Phase D).
- **The writers.**
  - `plain._Writer.run` (line 274) writes `TEXT`, `FIELD_CODE` and `LIST_LABEL` as stored; `comment` writes
    `comment.text`.
  - `markdown._text` (line 925) splits at `\n` and calls `_InlineSink.text(line, bold, italic, strike)`, or a `md.Code`
    for a monospace line. `_InlineSink.flush` nests strong, emphasis and strikethrough. Raw HTML is `md.HtmlInline`
    (`<br>`, CriticMarkup). `_code_text` writes a code paragraph.
  - `commonmark.py` escapes `<` in text (`_TEXT_ESCAPE`).
- **`TextWarning`** has six kinds and `part_name`. `_part_name(story.part)` names a story's part.
- **The harness.** `scenario_requirements` raises `ValueError("Unknown scenario")` on any option `_OPTION_STATEMENTS`
  lacks, and it lacks `small_caps`, `drop_caps` and `profile`. The adapter advertises `DOCX4J_IMPLEMENTED` at
  `198d015a3`, less `INSPECTOR` (REQ-901 to 907, REQ-1204), plus the HTML reader's absence. There is no inspector here
  (CR-011 section 3.7).
- **Documents that hold the properties.** Of the 406 `.docx` under `tests/` (the suite copy included):
  - in a story part: `w:caps` in 5, `w:smallCaps` in 3, `w:dstrike` in 2, a raised or lowered `w:vertAlign` in 12;
  - in `styles.xml` alone: 2, 27, 1 and 60 more;
  - a `w:dropCap` of `drop` or `margin` in none.

  Of the 13 under `samples/`, only styles hold any (1 caps, 2 small caps).
- **Kept files.** `resources/` and `model/` are in `codegen/clean.py`'s `KEEP`. `scripts/generate_fonts.py` is the
  precedent for a committed, generated resource. markdown-it-py is a dependency already (`insert_markdown`), so a test
  can render Markdown.

## 4. Design

### 4.1 The table (Phase A)

- **The script** is `scripts/generate_case_mapping.py`, in `generate_fonts.py`'s manner. It reads `UnicodeData.txt` at
  the pinned version (Q1):
  - from a path it is given, else downloaded from `https://www.unicode.org/Public/15.0.0/ucd/UnicodeData.txt`;
  - in either case checked against the SHA-256 the script pins.
- **What it writes.** It takes field 12 (Simple_Uppercase_Mapping) of every code point that has one, and refuses any
  pair that would change a UTF-16 length, so the three ports' offsets never part. It writes
  `docx4j_py/resources/case-mapping.json`:
  - the pairs, in code point order;
  - a `_source` object naming the version, the URL, the SHA-256 and the count (1,450 at 15.0).

  `--check` regenerates and compares. The output is never hand-edited, and nothing is computed from `unicodedata`
  (Q2).

### 4.2 `model/text/case_mapping.py` (Phase A; docx4j `CaseMapping`)

- `DEFAULT_LANGUAGE = "en-US"`.
- **`upper(text, language)`**: `text.translate(table)` over the committed table, read once on first use. Where
  `dotted_i(language)` holds (the primary subtag, split at `-` or `_` and lowercased, is `tr` or `az`), a second table
  adds `i` to U+0130. Each code point maps to one, so `len()` is kept; no runtime guard is needed.
- **`is_upper(run, options)`**: `run.caps or (run.small_caps and options.small_caps == "upper")`.
- **`written(item, options)`**, one function for docx4j's three overloads:
  - a run's text, in capitals where `is_upper`;
  - a block's text as `TextBlock.text` builds it, each text run written;
  - a comment's blocks a line apart.
- The docstring states the Unicode version, as WD 0.52's REQ-607 asks, and that a later case pair stays unmapped until
  the table moves.

### 4.3 The run's flags and language (Phase A)

- **`TextRun` gains four members**, under docx4j's names in this package's spelling (Q4):
  - `caps: bool`;
  - `small_caps: bool`;
  - `vertical_alignment: VerticalAlignment`, a `Literal["BASELINE", "SUPERSCRIPT", "SUBSCRIPT"]` beside `RunKind`,
    `"BASELINE"` by default;
  - `language: str | None`.

  `same_format` and `copy_format` take all four. `to_dict` writes `caps` and `small_caps` where true,
  `vertical_alignment` where it is not baseline, and `language` where it is set. The docstrings give the WD's `lang` and
  `vertAlign`.
- **`_run`** sets them as docx4j's `formatOf` does:
  - `caps` is `_on_of(r_pr, styles, "caps")`;
  - `small_caps` is `not caps and _on_of(r_pr, styles, "small_caps")`;
  - neither is cancelled by the paragraph style.
- **`vertical_alignment`** is `values.vertical_alignment` of the run's own `w:vertAlign/@w:val`, else of its styles'. A
  value outside the enumeration is unset, as CR-018 reads every enumeration, so the style's shows through. That is
  Word's reading. docx4j's `formatOf` takes the run's own element wherever it is present, so an invalid own value there
  likely gives baseline over a raised style (uncertain: read from the code, not run). A named departure, for an invalid
  value only.
- **`language`** is set where `caps or small_caps`:
  - the run's own `w:lang/@w:val` (stripped, a blank value none);
  - else `styles.lang`'s;
  - else the `default_language` option, `DEFAULT_LANGUAGE` (`en-US`) unless set (WD 0.80; section 4.3b). The settings
    part is not read for the language.
- The text stays the stored text, so the block model, addresses, excerpts and `office_offset` never move (REQ-607). The
  content API's Office JS readers, `text_of` and the python-docx facade keep the stored characters.

### 4.3b The language of last resort (amended 2026-11-04, WD 0.80)

WD 0.80 (portfolio `df66c75`) rules REQ-607's chain. The user struck `w:themeFontLang` from it: a caps run's language is
its own `w:lang`, else its styles', else `w:docDefaults`', else the option `default_language`, a BCP 47 tag, `en-US` by
default.

- **Measured.** docx4j's screen-caps probe, Word 365, 2026-11-04: a `w:caps` run with no language anywhere is cased by
  the installation's editing language, on screen, in the PDF and in the text save alike. A file with `w:themeFontLang`
  and one without behave identically, under an English and under a Turkish editing language.
- **The option** stands in for the editing language, which a library cannot see. It reaches only this mapping: caps,
  and small caps under `small_caps: upper`. `word-text` (CR-027 Phase D) leaves it at `en-US`.
- **Here**, `TextOptions.default_language`: a string, stripped. A blank or non-string value is refused with
  `TextOptionError` naming REQ-607. It is matched as a `w:lang` is (section 4.2's `dotted_i`), so `tr`, `tr-TR` and
  `az-Latn-AZ` give `İ`; no fuller BCP 47 check is made. It is not one of `_CHOICES`, since it's a free tag.
- **docx4j at the tag** reads `w:themeFontLang` (section 2.1). The ruling moves docx4j too, so the port follows the WD,
  not `75b9bdb54`.
- **The settings part** is no longer consulted for case mapping. CR-031's `part.not_read` for settings stays on its other
  three consults.
- **The oracles move.** docx4j's A.8 intact fixture, whose style-cased run stands under `w:themeFontLang` tr-TR, now
  reads `ISTANBUL`. Goldens taken from that fixture move with it.
- **Effort:** unchanged, one option for one consult fewer.

### 4.4 The writers (Phase A; the tags in Phase C)

- **`plain._Writer.run`** writes a `TEXT` run as `written(run, options)`, and a comment's text as
  `written(comment, options)`. `FIELD_CODE` and `LIST_LABEL` stay as stored (CR-069's unmeasured list).
- **`markdown`**:
  - `_text` takes `written(run, options)` before it splits into lines, so a code span gets the capitals too;
  - `_code_text` writes a code paragraph's `TEXT` runs written and its `FIELD_CODE` runs stored;
  - `_comment` writes `written(comment, options)`.

### 4.5 The drop-cap join (Phase B)

- **`blocks()`** keeps a local `frame: TextBlock | None`. It is cleared:
  - at any item that is not a `w:p`, an `_is_marker` item or one the view removes;
  - where a chain goes into a table (`joins_into_table`). A stranded table join, read after its table, takes none.

  A paragraph read while a frame is pending takes it. A paragraph read with none becomes the frame where
  `_drop_cap_frame` holds.
- **`_paragraph(..., frame=None)`** returns its block, or None in its two no-block returns (hidden under
  `hidden: false`; a table of contents' under `toc: false`). With a frame it calls `_join_drop_cap(frame, block, story,
  sink)` in place of `sink.append(block)`, and appends as usual where the frame is no longer in the sink.
- **`_drop_cap_frame(chain, block)`** holds where all of these do:
  - `drop_caps == "join"`;
  - `len(chain) == 1`;
  - the `w:p`'s own `p_pr.sect_pr` is None;
  - no character of `block.text` is one that `str.isspace()` takes, U+0085 excepted. Measured over every code point on
    3.12 and 3.14, `isspace` is Java's `isWhitespace or isSpaceChar` plus U+0085 (NEL) alone;
  - `in_enumeration(drop_cap, STDropCap)` is `DROP` or `MARGIN`, read from `resolver.get_effective_ppr(_p_pr_of(p))`.
- **`_join_drop_cap`** follows docx4j's `joinDropCap`:
  - the frame's runs first, then the block's;
  - the frame's fields first, the revision counts summed, `hidden` only where both;
  - `sink[at] = block` in the frame's place;
  - `TextWarning(TextWarning.DROP_CAP_JOINED, frame.address, part_name=_part_name(story.part))`.

  The letter's run merges into its word's first run as `joinDropCap` merges them (Q3): where the frame's last run and
  the block's first have the same kind, hidden flag, placeholder flag, revision and formatting (`same_format`),
  docx4j's merge rule, not this package's contiguity rule of CR-011 section 16.5 (section 4.5a, which also covers the word's later runs). Word's C02 is then one run,
  `The C02_ modern ...`, as in docx4j.
  - The merged run keeps the letter's `source`, as a merge keeps the first run's.
  - Its `office_offset` is None: its text is two paragraphs' (the frame's `Paragraph.text` and the next one's), which
    no one offset into one `Paragraph.text` describes. CR-011 section 12.6's invariant holds, since a run with no
    offset claims none. `TextRun.office_offset`'s docstring gains this case.
  - Where the two runs differ (a bold or cased letter), the letter stays its own run with its offset.
- **`TextWarning.DROP_CAP_JOINED`** (`"paragraph.drop_cap_joined"`) is the seventh kind. Its docstring says that the
  selector is the paragraphs' part. Under CR-027's selector order (the branch, else the part name, else the content
  type) the selector is the part name.

### 4.5a Runs across two paragraphs of one block (Phase B)

The user, 2026-11-02, after Q3's ruling: "check it against a document, but subjectto that, yes  deleted-mark joins changed to match docx4j as well". The probe (section 1) confirmed the difference. So the two
joins that make one block of two `w:p` merge runs as docx4j's `addText` does:
- **The joins:**
  - a joined drop cap (4.5);
  - CR-043's join, where the view deletes a paragraph's mark (a deleted mark in `accepted`, an inserted one in
    `original`, and either in a chain).
- **The rule.** Where the last run and the next come from two different `w:p` of the block, they merge on docx4j's
  rule alone: the same kind, hidden flag, placeholder flag and revision, and `same_format`. Contiguity is not
  required. Within one `w:p`, CR-011 section 16.5's contiguity rule stays as it is.
- **The merged run's provenance.**
  - It keeps the first run's `source`, as any merge does.
  - Its `office_offset` is None: its text is two paragraphs', and no one offset into one `Paragraph.text`
    describes it. CR-011 section 12.6's invariant holds, since a run with no offset claims none.
  - `TextRun.office_offset`'s docstring names the case.
- **Later runs.** A merged run takes later runs of the second `w:p` on the same rule, so a run split again inside
  the second paragraph (an `rsid` boundary, a bookmark) still joins it, as in docx4j. Without this, the merged run's
  missing offset would stop the next merge under the contiguity rule. Word's C02, `The C02_ modern ...`, is one run
  only with it.
- **What moves.** The block model's runs at every such join where the formatting agrees, and those runs'
  `office_offset`. Plain text, Markdown and every golden are unchanged, since a merge changes no text. Phase B
  records the moves over `samples/`, the text-extraction suite and the revisions fixtures, as CR-011 section 16.5
  recorded its own. Any caller that reads `office_offset` at a join, the content API's readers among them, is listed
  there.

### 4.6 The tags (Phase C)

- **`_InlineSink`** gains `vertical_alignment` beside `bold`, `italic` and `strike`. Its `text(s, b, i, st,
  va="BASELINE")` flushes when any of the four changes. `flush` puts `md.HtmlInline("<sup>")` and its closing tag
  (or `sub`) around the nested strong, emphasis and strikethrough node.
- **`_text`** passes the run's alignment. A monospace line is `md.HtmlInline` open, `md.Code(line)`, `md.HtmlInline`
  close. Lines split at `\n` keep their own tags, with `md.HardLineBreak` between.
- `_code_text` writes no tags. Note references never reach `_text`. `commonmark.py` already escapes a `<` the document
  holds. The table cells' pipe escape does not touch the tags.

### 4.7 Strike (Phase D), and the kind left out

- **`_run`'s `strike`** becomes `(_on_of(r_pr, styles, "strike") or _on_of(r_pr, styles, "dstrike"))`, and is false
  where the paragraph style has either on. That is REQ-703's one strike flag, and Markdown writes `~~` for a
  double-struck run.
- **`text.formatting_join`** is the inspector's (WD 9.2, REQ-905), and the inspector is not ported (CR-011 section 3.7;
  CR-027 section 4.5). The kind joins the adapter's `INSPECTOR`, so it is never advertised. Its four `dpp/` rows pend by
  capability, as the 214 do.

### 4.8 Options and the preset

- `OPTION_NAMES` gains `drop_caps` after `toc` (Phase B) and `small_caps` after `symbols` (Phase A), in WD section 10's
  order, with `profile` last (CR-027).
- `_CHOICES` gains `"drop_caps": (("join", "asis"), "REQ-521")` and `"small_caps": (("asis", "upper"), "REQ-607")`. The
  defaults are `join` and `asis`.
- `word-text` (CR-027 Phase D) gains `small_caps: upper` (Phase A) and `drop_caps: asis` (Phase B), giving docx4j's 26
  values at the tag. `TextOptions`' docstring lists both.
- `OPTION_NAMES` gains `default_language` (Phase A; WD 0.80, section 4.3b) in WD section 10's order. It's a free tag,
  not a choice; `word-text` leaves it at `en-US`.

### 4.9 The harness

- **The scenarios.** docx4j's `ExtractionConformance.requirements` has mapped three scenario names since `993e385e7`:
  - `small_caps` to REQ-607;
  - `drop_caps` to REQ-521;
  - `profile` to REQ-1002.

  Without them `scenario_requirements` raises on `dpp/`'s 8 `small_caps: upper` token rows and its 22
  `profile: word-text` rows. CR-027 section 4.4 takes them in its Phase A (since its review of 2026-11-02), so CR-029
  starts with them in place. If CR-027 Phase A lands without them, CR-029 Phase A adds them first.
- **What CR-029 adds on top:**
  - the adapter advertises REQ-607 from Phase A, REQ-521 from Phase B and REQ-710 from Phase C, each only once its
    phase is built, whatever `DOCX4J_IMPLEMENTED` records of docx4j's. That turns `dpp/`'s pending rows into checks
    (section 5);
  - a row still pending names the CR-029 phase that builds its statement, as docx4j's `pendingReason` names CR-069's;
  - `text.formatting_join` pends as the inspector's.

### 4.10 Not here

- The inspector and `text.formatting_join`.
- HTML: `Docx4J.toHTML`'s `sup` and `sub`, and REQ-710's HTML serialization, which neither docx4j nor this package
  writes.
- Capitals in a field code under `field_codes: true` and in a list label, which docx4j writes as stored and has not
  measured.
- Reading `<sup>` back in `insert_markdown`. docx4j's import drops inline HTML too (its section 13.2 item 6).

## 5. Tests and oracles

Each phase is held to docx4j's test at the tag and to `dpp/`'s rows. The counts assume CR-027 Phases A and D (Q5).

- **Phase A**, `tests/text/test_case_mapping.py` (`CaseMappingTest`'s ten):
  - Word's table (`ISTANBUL FILM`, `İSTANBUL FİLM`, `STRAßE`, `CAFÉ`, `ﬁNE CAFÉ Ǆ`) and the dotted-i rule;
  - the differences from `str.upper` pinned: `ß`, `ﬁ`, U+1FB3, the digraph `ǆ` to U+01C4 (not the titlecase U+01C5),
    and a supplementary letter;
  - the length kept, as a check over the whole table (one code point to one, none across the BMP's edge);
  - the flags and the language on Word's probe, a paragraph style's caps, a run under both, and runs merged by
    language;
  - the language of last resort: `default_language`, `en-US` unless set; `tr-TR` gives `İ`; a `w:themeFontLang`
    changes nothing (WD 0.80, section 4.3b);
  - a note reference baseline, and the option.

  Every character the tests map has had the same simple uppercase since Unicode 10.0. The table is held to its
  `_source` (15.0.0, 1,450 code points). In `dpp/`, the 25 REQ-607 token rows become checks, and 13 `word-saves.tsv`
  rows (C05 and the 12 caps-probe saves) are compared with Word's own text.
- **Phase B, the joins' runs (4.5a):** the probe built in the test, with `MiniDocx`, as section 1 describes it. J1
  accepted and J2 original are each one run with no `office_offset`; J3 is two runs; a second run inside the
  second paragraph joins the merged run. Then the moves over `samples/` and the fixtures are recorded.
- **Phase B**, `DropCapJoinTest`'s seven text-model cases:
  - Word's C02 joined in every view, at `body/0`, `body/2` and `body/3`, the warning at `body/1`, the selector
    `/word/document.xml`;
  - `asis` by keyword, by name and by the preset, and a caller's `join` over the preset;
  - the shapes left: two words, a table next, the story's end, `dropCap` none or absent, a frame ending a section, a
    control next, the cell's end;
  - the shapes joined: `margin`, a style's frame, a bookmark between or inside, a cell's next paragraph, an empty frame;
  - a frame whose mark the view deletes (CR-043's join, no warning), the letter and its word (Q3), and the option.

  In `dpp/`, the 2 `paragraph.drop_cap_joined` findings rows become checks, and C02's save under `drop_caps: asis`.
- **Phase C**, `MarkdownVerticalAlignTest`'s six, with docx4j's strings. Each Markdown is rendered by markdown-it-py's
  CommonMark preset with GFM strikethrough, as docx4j renders through commonmark-java. In `dpp/`, the 8 REQ-710 token
  rows become checks, with C06's and C07's saves and both goldens.
- **Phase D**, `aDoubleStruckRunIsStruck`, and a double-struck run's `~~` in Markdown. No `dpp/` row moves. The 6
  inspector findings rows stay pending.
- **After D**, every `dpp/` row is checked but the 6 inspector rows: 34 token rows, 2 findings, 2 goldens, 22 saves.
- **The gates:**
  - the suite;
  - `codegen/generate.sh --check`, with `model/text/case_mapping.py` and `resources/case-mapping.json` listed in `KEEP`
    by path, as CR-025's files are;
  - `generate_case_mapping.py --check`;
  - `scripts/roundtrip.py` and `compare_acceptance.py` unchanged, since nothing saved changes.
- **What moves.** Each phase records every test whose expected text moves (a document with `w:caps`, `w:dstrike` or a
  raised run; section 3's counts). A failure outside this CR is taken to the user, not fixed in passing.

No Word check: extraction output is not a saved file, and nothing this CR changes is written to one.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A (docx4j phase 1) | `generate_case_mapping.py` and `resources/case-mapping.json`; `model/text/case_mapping.py`; `TextRun.caps`, `small_caps`, `vertical_alignment` and `language`, read in `_run`; plain text and Markdown in capitals; `small_caps` and the preset's `upper`; REQ-607 advertised; the harness's three scenarios if CR-027 Phase A lacks them; `CaseMappingTest`'s ten | 1 day |
| B (phase 2) | The drop-cap join in `blocks()` and `_paragraph`; `TextWarning.DROP_CAP_JOINED`; `drop_caps` and the preset's `asis`; REQ-521 advertised; `DropCapJoinTest`'s seven; runs across two paragraphs of one block merged as docx4j's (4.5a), the deleted-mark join's included, and the moves recorded | ¾ day |
| C (phase 3) | `<sup>` and `<sub>` in `_InlineSink` and `_text`; REQ-710 advertised; `MarkdownVerticalAlignTest`'s six | ¼ day |
| D (phase 4) | `strike` taking `w:dstrike`; `text.formatting_join` in `INSPECTOR`; the dstrike test | ¼ day |
| E | The records: CHANGELOG (Added: the four run members, `small_caps`, `drop_caps`, `paragraph.drop_cap_joined`, `case_mapping`; Changed: capitals for `w:caps`, the drop cap joined, the tags, `w:dstrike` struck); CLAUDE.md's status and file list; `tests/README.md`'s rows and the index's row; the portfolio's `python/CR-029.*` entries through the docx4j-portfolio session | ¼ day |

## 7. Risks

- **Output changes for every caller.** Plain text and Markdown change for any document with `w:caps` text, Markdown for a
  double-struck or raised run, and the block model's runs split where case, alignment or language differ. docx4j's
  CHANGELOG for 17.3.5 names the same changes, and so will this one. Section 3's counts bound what moves here.
- **Word's own table is unmeasured.** Word uses Windows' casing at an unknown version. For example, 15.0 maps Georgian
  Mkhedruli under `w:caps` to Mtavruli capitals, and docx4j on Java 11 does not. Word is unmeasured for both. The pinned
  version is a stated choice, not a measurement (Q1).
- **Offsets at joins.** Runs merged across two paragraphs lose their `office_offset` (4.5a). A reader that maps a
  model run back to Office JS text has nothing for them. Phase B lists the readers that use the offset, and how many
  runs lose it.
- **Departures from docx4j.** An invalid own `w:vertAlign` (section 4.3) makes the block model differ from docx4j's.
  It changes neither plain text nor Markdown. The drop cap's letter follows docx4j (Q3), at the cost of the merged
  run's `office_offset`.
- **Drop-cap false joins.** Word writes `w:dropCap` for its drop-cap feature alone (the WD's NOTE). The one-word limit
  guards a hand-made frame, and `drop_caps: asis` gives Word's own text.
- **Unmeasured in Word** (CR-069 section 7): a style's caps, a `tr-TR` document default, a style's frame, a frame ending
  a section and a frame in a cell. Each is ported as docx4j builds it.
- **Tags in Markdown.** A consumer that strips HTML gets the baseline digits, which is what it gets today.
- **CR-027 first.** Without its Phase A, `dpp/`'s rows cannot be run here. Without its Phase D, the 16 saves cannot be
  compared. CR-029's own tests hold either way.

## 8. Questions (decided 2026-11-02)

The user, 2026-11-02: "CR-029's questions Q1 to Q5 as recommended." Then, the same day, on Q3: "you and core-ts should handle a drop-cap letter as docx4j does!" So Q1,
Q2, Q4 and Q5 are decided as recommended, and Q3 against its recommendation: the letter merges as docx4j's
`joinDropCap` merges it (section 4.5). The ruling covers core-ts too (its section 24.3 and D6). Later, the same day:
"check it against a document, but subjectto that, yes  deleted-mark joins changed to match docx4j as well". The probe confirmed it (section 1), so a deleted-mark join's runs merge as docx4j's do (4.5a).

| Q | Question | Recommendation |
|---|---|---|
| Q1 | The case table's Unicode version: 15.0 or 16.0? | **15.0.** It is the version of the JDK docx4j's tests ran on (1,450 code points), so this package and docx4j map every code point alike there. It is also core-ts's D4 (decided 2026-11-02), so the ports agree. 16.0, the version Python 3.14 ships, was the first plan, on the understanding that U+1C89/U+1C8A was the only difference. Measured, 16.0 maps 27 more code points (section 2.1), U+019B and U+0264 among them. With either version the table is committed and its version is in `case_mapping`'s docstring. The tests map only characters stable since 10.0, and a move is a deliberate regeneration. |
| Q2 | The table generated by `scripts/generate_case_mapping.py` from that version's `UnicodeData.txt` (field 12), its output committed as `resources/case-mapping.json`, the input fetched from unicode.org and checked against a pinned SHA-256 rather than committed; never computed from `unicodedata` at import, and `str.upper` never used? | Yes. `unicodedata` has no simple case mapping at all, and its version differs on each of CI's three Pythons (15.0.0, 15.1.0, 16.0.0). The pinned hash makes the fetch reproducible. Committing `UnicodeData.txt` (about 2 MB) would let the script run offline, at the cost of a large file nothing else reads. |
| Q3 | At a joined drop cap, does the letter's run merge into its word's first run, as docx4j's `joinDropCap` does (Word's C02 as one run, `The C02_ modern ...`)? Or does it stay its own run under this package's merge rule? | **Its own run.** The merge rule of CR-011 section 16.5 already keeps runs from two `w:p` apart wherever both carry an `office_offset`, in CR-043's chain join too (read from `_add_text`). Merging would leave the joined run's text in neither paragraph's `Paragraph.text`, so its `office_offset` and `source` would have to go, against CR-011 section 12.6's invariant. Plain text and Markdown are the same either way. The cost is a recorded departure: the block model has `["T", "he C02_ modern ..."]` where docx4j has one run. core-ts's section 24.3 plans docx4j's merge over a contiguity rule like this one, so the user may want to rule for both ports. **Decided 2026-11-02, against this recommendation: as docx4j does** (the user: "you and core-ts should handle a drop-cap letter as docx4j does!"). The letter's run merges into its word's first run where docx4j's rule merges them, and the merged run has no `office_offset` (4.5). |
| Q4 | The four run members named after docx4j's getters in this package's spelling (`caps`, `small_caps`, `vertical_alignment` with `"BASELINE"`, `"SUPERSCRIPT"`, `"SUBSCRIPT"`, and `language`), the same names in the JSON, the WD's `vertAlign` and `lang` given in the docstrings? | Yes. That matches core-ts's D3 (docx4j's names), the existing members (`display_math`, `note_type`, `prog_id`) and CR-027's JSON spelling. The JSON writes each only where it is set, as the other flags are written only where true. |
| Q5 | Order: CR-029 after CR-027's Phases A and D, its phases A to E in docx4j's order, each built and verified against `dpp/` before the next? | Yes. CR-027 Phase A brings the rows and Phase D the preset, which every `word-saves.tsv` row runs under, so each phase is measured against Word's own text as it lands. Phase A can begin after CR-027 Phase A alone, its 13 saves checked once Phase D lands. |

## 9. Effort

2½ days: A 1, B ¾, C ¼, D ¼, E ¼.
