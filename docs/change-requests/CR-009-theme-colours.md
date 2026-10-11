# CR-009: Theme colours --- the colour merge rules docx4j CR-034 changed, and its colour resolver

**Status:** Proposed 2026-10-12, on docx4j CR-034's hand-off (the docx4j session, 2026-10-12,
after its phase 3). **Decided 2026-10-12: all three recommendations of section 9 accepted**; Phase
A (the merge rules, small and parity-bound) in progress; Phase B (the WordprocessingML resolver) is
new scope; Phase C (DrawingML and SpreadsheetML) waits for a consumer.
**Depends on:** docx4j CR-034 (phases 0 to 3 done 2026-10-12: `3dbe4eee5` the resolver,
`2e381ef26` the outputs and `StyleUtil`, on `VERSION_17_3_3`, local and not yet pushed) and this
repository's CR-002 Phase B (implemented, section 15), whose property catalogue mirrors
`StyleUtil`. Phase C also depends on CR-001 Phase D (the PML and SML models) and has no consumer
before CR-004 and CR-005.
**Counterpart:** docx4j's `docs/developer/change-requests/CR-034-theme-colours-resolved.md`
(section 2 the rules, section 4 the design, section 5 the probes, section 11 Office's readings) and
docx4j-core-ts's `docs/change-requests/CR-008-theme-colours.md` (proposed the same day, the same
two phases; the two engines stay one design).

## 1. Why

docx4j CR-034 did two things that reach this package.

**The merge rules changed.** `StyleUtil`, which `docx4j_py/model/property_catalogue.py` mirrors
entry by entry (CR-002 section 15.1), now treats a colour as one thing. Word draws the theme
reference (`w:themeColor`, `w:themeFill`) and keeps the value beside it (`w:val`, `w:color`,
`w:fill`) as a cache, which can be stale (CR-034 probe P1: every one of its eight runs whose value
disagrees with its reference is drawn in the theme colour). Once the reference wins, the
attribute-by-attribute merge CR-015 introduced gives wrong answers: a run's own
`w:color w:val="auto"` under a style's `accent1` kept the style's reference and drew in the
accent, where Word draws the run black (corpus document 3229, Word's own PDF). So docx4j changed
the merge, and the parity goldens this package is held to (45 of them, zero differences since
CR-002 Phase B) are docx4j's answers: when docx4j-core-ts regenerates them at the CR-034 commit,
every golden whose effective properties carry a theme colour is rebuilt under the new rule.

**A resolver exists.** `org.docx4j.model.color` turns any colour reference Office writes into the
RGB value Word, PowerPoint or Excel draws, measured against 1,068 probe cases and the 36,919
theme attributes of four corpora. Nothing in this package resolves a theme colour today:
`Font.color` reports the effective `w:color`'s value, which for a theme colour is Word's cached
value or `auto`; the python-docx facade's `ColorFormat.theme_color` reports the reference's name
and nothing gives its value.

## 2. What this package does today (checked 2026-10-12)

`property_catalogue.py`:

- `is_empty_color` reads `w:val` alone (docx4j's `isEmpty(Color)` before 17.3.3), and
  `is_empty_u` reads `w:val` and `w:color` alone. A style's `w:color` stating only a theme colour
  is dropped as empty. `is_empty_shd` and `is_empty_border` already count the theme attributes,
  as docx4j's did.
- `merge_color`, `merge_u`, `merge_shd` and `merge_border` merge attribute by attribute: `_attr`
  for each value, tint and shade, `_theme_colour` for the enums (`none` does not override). This
  is exactly the behaviour CR-034 replaced.

**Measured against the goldens.** CR-034's rules swapped into the catalogue (a pytest plugin that
replaced the six functions' code in place, so the catalogue's entries saw them) and the three
golden tests run: **2 of the 45 goldens move**, both in `test_properties_goldens.py`, 152 of 154
tests still green, the numbering and fonts goldens untouched.

| Golden | What moves | Why, and which is right |
|---|---|---|
| `fonts-modesOfApplication`, `main[5]` run 3 | `Heading1` colours its text `accent1` shaded `BF` (`365F91`); the run's character style `SubtleReference` states `accent2` (`C0504D`) with no shade. Attribute by attribute the shade leaks through: docx4j recorded `accent2` **shaded `BF`** beside `C0504D`. As one, the run is `accent2` plain | The new answer. Word draws `SubtleReference` in `C0504D`, which is what the character style wrote beside its reference; a shade from an unrelated colour has no business there |
| `invoice2013`, table 0, `effectiveTableStyle` | `TableGrid`'s borders are `single 000000 text1`; the layer beneath or above states `none` with `color="auto"`. Attribute by attribute `themeColor="text1"` survives beside `auto`; as one, `auto` replaces the reference | The new answer: `auto` states a colour |

Both move towards Word and towards docx4j 17.3.3. Over `samples/` and `tests/fixtures/` (64
documents) the theme references sit almost wholly in styles --- 729 `w:color`, 294 border edges
and one `w:u` in `styles.xml`, against 28 attributes in document bodies --- and no run in them
states its own value under a style carrying a reference, so the content API's answers over this
corpus change in those two places alone.

## 3. Phase A: the merge rules

A mirror of `StyleUtil` at `2e381ef26`, in `property_catalogue.py`:

1. **Emptiness counts the theme reference.** `is_empty_color` is empty only when `w:val`,
   `w:themeColor`, `w:themeTint` and `w:themeShade` are all absent; `is_empty_u` likewise over
   `w:val`, `w:color` and the three theme attributes. (`areEqual` has no counterpart here: the
   catalogue compares nothing.) A style's `w:color w:themeColor="accent1"` with no `w:val` is now
   inherited.
2. **A colour inherits as one.** A new `_states_colour(value, theme_colour)`: true where the
   source states a value (`w:val`, `w:color`, `w:fill`, not empty) or a theme reference (not
   absent, not `none`). Where it is true, the source's value, reference, tint and shade replace
   the destination's **together**, `None` included; where false, tint and shade still inherit on
   their own through `_attr`. This applies to:
   - `merge_color`: `val` with `theme_color`, `theme_tint`, `theme_shade`;
   - `merge_u`: `color` with the three; `val` (the underline kind) stays separate;
   - `merge_shd`: the pattern-colour group (`color`, `theme_color`, `theme_tint`, `theme_shade`)
     and the fill group (`fill`, `theme_fill`, `theme_fill_tint`, `theme_fill_shade`), **each its
     own**, so CR-015's rule that a `themeFill` survives a direct `w:shd` stating no fill still
     holds; `val` (the pattern) stays separate;
   - `merge_border`: `color` with the three; `val`, `sz` and `space` stay separate. `merge_p_bdr`,
     the table and cell border sets inherit it edge by edge as they do now.
3. **Nothing else.** `unset` and `has_direct_formatting` go through the same `is_empty`
   functions and need no change. `ST_HexColor` is held as `bytes` in this model (`w:val` is
   `None | STHexColorAuto | bytes`), which `_empty_scalar` already handles; `_states_colour` must
   treat `auto` (the enum member) as a stated value, as the Java's `!isEmpty(String)` does.

**Tests.** The three `StyleUtilMergeRulesTest` cases ported into `tests/content/test_properties.py`
(a value replacing an inherited reference with its shade; a theme-only colour not empty, inherited,
replacing a value, a tint alone still inheriting; shading fill, border and underline colours as
one, the pattern colour its own). The golden tests then show the two moves of section 2 until the
goldens are regenerated at docx4j `2e381ef26` or later (docx4j-core-ts CR-008 phase 1 does that,
once docx4j pushes `VERSION_17_3_3`); until then the two are pinned as `xfail(strict=True)` with
the golden and the rule named, so that the regeneration turns them into unexpected passes and the
pins come out. Zero differences is the bar on either side of that.

CR-002 section 15.1 gains an item recording the change and its date, as 15.1 item 6 will for
CR-008.

## 4. Phase B: the WordprocessingML resolver

The port of `org.docx4j.model.color`'s Word half, in `docx4j_py/model/color/`, beside
`model/fonts/` (which already reads the theme part's `a:fontScheme`):

- `ResolvedColor` (`resolved.py`): a frozen dataclass of `red`, `green`, `blue`, `alpha` (OOXML's
  unit, 100,000 opaque, as docx4j's CR-034 D5 decided, not a colour library's); `of_hex`
  (`RRGGBB`, `auto` and `None` giving `None`), `hex`, `within(other, units)`, `BLACK`, `WHITE`.
- `ThemeColors` (`theme_colors.py`): the twelve slots (`dk1`, `lt1`, `dk2`, `lt2`, `accent1` to
  `accent6`, `hlink`, `folHlink`) read from a theme part's `a:clrScheme` (typed, `docx4j_py.dml`),
  with the colour map. `ThemeColors.of(package)`: the main document part's theme part, or the
  **default theme** the package is read as having when it has none (`pkg.default_theme`,
  CR-002 section 15.3; the bundled 2007 theme carries 2013's colour scheme, docx4j's D2, accepted
  and documented); then `w:settings/w:clrSchemeMapping` applied as `withWordMapping` does, mapping
  `bg1`, `tx1`, `bg2`, `tx2`, the accents and the two hyperlink slots. `word_theme_color(name)`:
  `text1`, `background1`, `text2`, `background2` through the map; `dark1` to `followedHyperlink`
  directly; `None` for `none`.
- `word_colors.py`: `resolve(val, theme_color, theme_tint, theme_shade, theme)` with docx4j's three
  rules: the reference wins over the value beside it; tint and shade in float HSL luminance
  (`L·t + 1 − t`, `L·t`, `t` the byte over 255), the tint applied where both are present, the
  channels **truncated**; and Word's written value returned where it is within one unit of the
  computation (a current cache), so a document Word itself wrote resolves to Word's own bytes.
  `None` where there is no reference and the value is `auto` or absent: the caller's default. Over
  it `of_color(Color)`, `of_border(CTBorder)`, `of_underline(U)`, `fill(CTShd)` and
  `pattern(CTShd)`, each accepting the model's `bytes`, the `auto` enum member or `None`.
- `PropertyResolver.theme_colors` (made once, re-read by `refresh()`), `resolve_color(obj)`,
  `resolve_fill(shd)` and `resolve_pattern_color(shd)` over effective properties, as docx4j's
  CR-034 phase 2 added them.
- The content API: `Font.color` reports the colour Word draws where the effective `w:color` carries
  a theme reference (question 2), and `Font.to_dict()` gains `theme_color`, `theme_tint` and
  `theme_shade` as written, so an agent sees both the reference and its value. `describe()` follows
  `to_dict()`. The python-docx facade's `ColorFormat` keeps python-docx's semantics (`rgb` is the
  written value, `theme_color` the reference); question 3 asks whether it gains an extension.

**Vectors.** Consumed from `../docx4j/docx4j-core-tests/src/test/resources/theme-colors/` the way
the goldens are consumed from docx4j-core-ts (CR-002 section 12.12: the sibling checkout is the
source, nothing copied, the test skips when it is absent), by a new `tests/content/test_colors.py`:

- `word-tint-shade.tsv`: the 117 distinct (theme value, operation, byte, Word's value) triples of
  four corpora; every computation within one unit of Word's value, and the count of exact matches
  reported (docx4j's resolver: 77 exact).
- `cases.tsv` and `readings.tsv`, the WordprocessingML cases: P1 (17: runs, paragraph and run
  shadings, top borders, an underline, with values that disagree with their references) and P2's
  six runs under a swapped `w:clrSchemeMapping`. Each case is found by its label in the probe
  document, resolved through a `PropertyResolver` over that document, and compared with what Word
  drew, exact or within one unit; P1-10 (a tint whose written value is unusable) is one unit off in
  docx4j too and is the one tolerated case. The probe documents are docx4j's; this repository
  does not build them.
- Units: a themeless package resolves against the default theme's scheme; `none` resolves to
  `None`; the cache rule (a value one unit off is returned as written, two units off is replaced).

No Word check: nothing in this phase changes what is written.

## 5. Phase C: DrawingML and SpreadsheetML, later

`DrawingMLColors` (the six colour kinds and their transforms in document order: tint, shade, inv,
gamma and the channel transforms in linear RGB; hue, sat and lum in HSL with sat unclamped; gray as
Rec. 709 luma; `sysClr` the machine's colour, not `lastClr`; a shape's style references through the
format scheme, with `StyleFill` for an entry that is not a solid fill) and `SpreadsheetColors`
(`rgb`, `theme` with Excel's swapped first four indexes, `indexed` against the legacy palette or
the workbook's own, `tint` in Windows' integer HLS with HLSMAX 240 and L truncated after adding
0.02), with the P3, P4 and P5 vectors (403 + 403 + 156 + 66 + 10 cases) and `ThemeColors.of` for a
slide (its master's theme and `p:clrMap`, then the layout's and the slide's `p:clrMapOvr`) and a
workbook. The DrawingML half can be built against `docx4j_py.dml` as generated today; its hosts and
the SpreadsheetML half wait for CR-001 Phase D and CR-002 Phase C, and its consumers are CR-004
(a shape's fill as PowerPoint draws it) and CR-005 (a cell's fill). Not before one of them asks.

## 6. Noted, not in scope: a hyperlinked TOC's entries

Coming in docx4j (in its working tree on 2026-10-12 as a CR-034 follow-up, not yet committed): Word
draws the entries of a hyperlinked TOC (runs in a `TOC \h` field result) without the `Hyperlink`
character style their runs carry --- colour, underline, weight and size alike. Measured on the
corpus goldens: about 1,780 such entries, none in the Hyperlink colour; corpus document 373 shows
the weight (its TOC1 entries bold from TOC1's own formatting, its TOC2 entries regular, all black).
docx4j's outputs drop the style from those runs in a preprocess step (`TocHyperlinkStyle`).

Here it touches an effective read only: `Font` over a run inside a `TOC \h` result reports the
Hyperlink style's colour and underline through the `PropertyResolver`, where Word draws neither;
`describe()` and `to_markdown` say the same. Nothing is proposed until a consumer describes TOC
formatting; when one does, the rule belongs in the resolver's caller (the content API's effective
reads), not in `StyleUtil`'s mirror, since docx4j keeps it out of `StyleUtil` too.

## 7. Phasing and effort

| Phase | What | Days |
|---|---|---|
| A | `is_empty_color`, `is_empty_u`, `_states_colour`, the four merge functions; the three ported tests; the two golden pins; CR-002 section 15.1's item | 0.5 |
| B | `model/color/` (`ResolvedColor`, `ThemeColors`, `word_colors`), the resolver's accessors, `Font.color` and `to_dict()`, `tests/content/test_colors.py` over the vectors, `CHANGELOG.md`, the readmes where they show a colour | 1.5 |
| C | `DrawingMLColors`, `SpreadsheetColors`, `StyleFill`, the slide and workbook `ThemeColors`, the P3 to P5 vectors | 2, later |

Phase A stands alone and is what the goldens will demand; Phase B needs it (a resolver over a merge
that leaks a shade resolves the wrong colour).

## 8. Risks

- **The goldens turn before or after.** Port Phase A first and two goldens are red until
  docx4j-core-ts regenerates them (the pins of section 3); port it after and every theme-coloured
  golden is red the morning they land. The pins are the cheaper side, and they name what they
  wait for.
- **docx4j `VERSION_17_3_3` is not pushed**, so the oracle is a local tree; the hashes above are
  re-stated if the branch is rebased, as CR-008 risk 1 was closed when 17.3.1 shipped.
- **Answers change under callers of 0.1.0.** Phase A in the two cases of section 2 and their
  kind; Phase B wherever `Font.color` meets a theme reference, from the cached value to the drawn
  one, which differ only where the cache is stale (20 attributes in 7 of docx4j's 561 corpus
  documents, all in styles). Both towards Word. The CHANGELOG entry says so.
- **Word's rounding is docx4j's reading, not Word's code.** Truncation after float HSL is exact
  for Word 365's three unusable-value cases and within one unit of every value older Words wrote;
  the vectors carry that tolerance, and so must this port's tests.
- **The vectors are docx4j's test resources, not a published artefact.** Like the goldens they
  are read from the sibling checkout and the tests skip without it; CI runs without them, as it
  does without the goldens.

## 9. Questions (decided 2026-10-12, every recommendation accepted)

1. **Port Phase A before the goldens are regenerated, with the two pins?** *Recommendation: yes.*
   **Decided 2026-10-12: yes.**
   The rule is measured (Word's PDF of 3229) and the Java is committed; the pins are strict, so
   the regeneration cannot pass silently. The alternative is to wait for docx4j-core-ts's
   regeneration and port the same day, which ties a half-day correction to two other
   repositories' schedules.
2. **Does `Font.color` report the drawn colour?** *Recommendation: yes*, in Phase B
   (**decided 2026-10-12: yes**): Office JS's
   `font.color` is what Word shows, the content API's reads have been effective since CR-002
   Phase B, and the written value stays reachable through the typed tree and `to_dict()`'s new
   keys. The alternative, a separate `Font.resolved_color`, keeps 0.1.0's answer for a stale cache
   that Word itself does not draw.
3. **Does the python-docx facade gain a resolved colour?** *Recommendation: no*, not in this CR
   (**decided 2026-10-12: no**):
   python-docx's `ColorFormat.rgb` is the written value and the facade promises python-docx's
   semantics (CR-003 section 20); a caller who wants Word's drawing has the content API. Revisit if
   `docx4j-mcp-py` asks for colours, which CR-006 does not.
