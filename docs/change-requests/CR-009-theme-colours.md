# CR-009: Theme colours --- the colour merge rules docx4j CR-034 changed, and its colour resolver

**Status:** Proposed 2026-10-12, on docx4j CR-034's hand-off (the docx4j session, 2026-10-12,
after its phase 3). **Decided 2026-10-12: all three recommendations of section 9 accepted**;
**Phase A (the merge rules) implemented 2026-10-12** (section 10): the catalogue mirrors
`StyleUtil` at docx4j `2e381ef26`, exactly the two comparisons of section 2 move, pinned
`xfail(strict=True)` until docx4j-core-ts regenerated the goldens at `2e381ef26` the same day
(its `1f78b94`), when the pins came out (10.10), and no content-API answer over
the 64 corpus documents changes; Phase B (the WordprocessingML resolver) is new scope; Phase C
(DrawingML and SpreadsheetML) waits for a consumer.
**Depends on:** docx4j CR-034 (phases 0 to 3 done 2026-10-12: `3dbe4eee5` the resolver,
`2e381ef26` the outputs and `StyleUtil`, on `VERSION_17_3_3`, pushed by 2026-10-13) and this
repository's CR-002 Phase B (implemented, section 15), whose property catalogue mirrors
`StyleUtil`. Phase C also depends on CR-001 Phase D (the PML and SML models) and has no consumer
before CR-004 and CR-005.
**Counterpart:** docx4j's `docs/developer/change-requests/CR-034-theme-colours-resolved.md`
(section 2 the rules, section 4 the design, section 5 the probes, section 11 Office's readings) and
docx4j-core-ts's `docs/change-requests/CR-008-theme-colours.md` (proposed the same day, the same
two phases; the two engines stay one design).
**Phase B scheduled 2026-10-24** (the user: "0.2.0 waits for CR-009 Phase B"): docx4j's colour resolver ported as section 11 records docx4j-core-ts did (`27be4a7`), after CR-018 Phase B. **Phase B implemented 2026-10-24** (section 12): `docx4j_py/model/color/` (docx4j's `org.docx4j.model.color` at `a9e0303ff`, its Word half), the resolver's `theme_colors` and `resolve_color`, `resolve_fill`, `resolve_pattern_color`, and `Font.color`'s effective view as Word draws it. docx4j's vectors are copied into `tests/fixtures/theme-colors/` with their SHA-1s, and every P1 and P2 case is at docx4j's own tallies through a `PropertyResolver`; the corpus triples are 76 exact and 41 within one. Over 728 corpus documents, 119 of 58,555 effective answers move, each a theme reference meeting a stale cache, and 0 direct answers; check 53's Heading 1 reads Word's `#0F4761`. Nothing written changes: no Word check. One question is open for the user (12.9: automatic as `#000000`). Phase C (DrawingML and SpreadsheetML) waits for a consumer.

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

**Vectors.** *(As built, section 12.3: copied into `tests/fixtures/theme-colors/` with their SHA-1s,
as section 11 says core-ts did and as `tests/fixtures/onoff-fun/` was, so CI runs them.)* Consumed from `../docx4j/docx4j-core-tests/src/test/resources/theme-colors/` the way
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

**Luminance rounding (docx4j `691620fb7`, CR-034 section 11 "Follow-up: luminance rounding",
pushed on `VERSION_17_3_3`; handed off 2026-10-13).** docx4j's CR-037 chart gate found Excel chart
styles whose text, `tx1` at `lumMod 50000 lumOff 50000`, Office draws as 7F7F7F where docx4j drew
808080. A colour whose last change is a luminance transform is now converted and rounded as Office
does:

- **Where it applies.** The last transform other than an alpha one decides. If it is `lum`,
  `lumOff` or `lumMod`, the rule applies. `alpha`, `alphaOff` and `alphaMod` don't count.
- **`lumMod`** is `l * val / 100000.0`, in that order, not `l *= val / 100000`.
- **Back from HSL**, `l` clamped first, by Foley and van Dam's form
  (`ColorMath.hslToRgbFoley`), not the CSS form:

  ```
  m2 = l * (1 + s) if l <= 0.5 else l + s - l * s
  m1 = 2 * l - m2
  red   = foley(m1, m2, h + 1/3)
  green = foley(m1, m2, h)
  blue  = foley(m1, m2, h - 1/3)

  foley(m1, m2, h):
      h -= floor(h)
      m1 + (m2 - m1) * 6 * h            if h < 1/6
      m2                                if h < 1/2
      m1 + (m2 - m1) * (2/3 - h) * 6    if h < 2/3
      m1                                otherwise
  ```

- **Each final channel rounded half down**: `math.ceil(clamp01(v) * 255 - 0.5)`
  (`ColorMath.channelHalfDown`). Never `round()`, which rounds half to even.

The exact form matters because luminance results often land on exactly half a unit, and Office
rounds those halves both ways. With this form and this order of operations, the direction follows
the floating-point error: exactly .5, or just below, goes down; just above goes up. That predicts
29 of 30 halves. Python's floats are IEEE doubles, so the same expressions in the same order give
Java's bits. **Don't simplify algebraically, and don't round early.** Comments at the code should
say so, as docx4j's do.

Unchanged, and still rounded half up through the CSS form: `hue`, `sat` and their `Off` and `Mod`
variants, `comp`, and every other transform. `WordColors` (Phase B's `word_colors`, Word's
`themeTint` and `themeShade`) and `SpreadsheetColors` are unchanged too.

The oracle is `readings.tsv`, unchanged; docx4j now matches more of it. In P3
(`tc-p3-transforms.docx` and `.pptx`), the luminance and colour-picker cases are exact in 97 of 98
per host, up from 85. Newly exact in both hosts: T041, T173 to T175, T240, T300, T335, T350, T365,
T369, T370 and T380. T239 (`tx2 lumMod 25000`) is still one unit off: docx4j's red is
3.5000000000000018, Office's 3. No case of any family got worse. docx4j's `ThemeColourProbesTest`
holds the luminance families exact but one (`Tally.luminanceOff() <= 1`), and Phase C's vector
test holds the same bar.

**Reproduced here on 2026-10-13** with a 60-line script in the scratchpad (nothing committed). It
ported the three pieces above, read each P3 probe's theme, and covered every `shape fill` case
written as a `schemeClr` with only `lum`, `lumOff` and `lumMod` children:

- **182 of 184** are exact against `readings.tsv`;
- the two misses are T239 in each host (ours 040A10, Office's 030A10);
- the same script rounding half up gets 164.

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
- ~~**docx4j `VERSION_17_3_3` is not pushed**, so the oracle is a local tree; the hashes above are
  re-stated if the branch is rebased, as CR-008 risk 1 was closed when 17.3.1 shipped.~~ Closed
  2026-10-13: `origin/VERSION_17_3_3` holds `2e381ef26`, `3dbe4eee5` and the luminance follow-up
  `691620fb7`, with the hashes unchanged.
- **Answers change under callers of 0.1.0.** Phase A in the two cases of section 2 and their
  kind; Phase B wherever `Font.color` meets a theme reference, from the cached value to the drawn
  one, which differ only where the cache is stale (20 attributes in 7 of docx4j's 561 corpus
  documents, all in styles). Both towards Word. The CHANGELOG entry says so.
- **Word's rounding is docx4j's reading, not Word's code.** Truncation after float HSL is exact
  for Word 365's three unusable-value cases and within one unit of every value older Words wrote;
  the vectors carry that tolerance, and so must this port's tests.
- ~~**The vectors are docx4j's test resources, not a published artefact.** Like the goldens they
  are read from the sibling checkout and the tests skip without it; CI runs without them, as it
  does without the goldens.~~ Closed 2026-10-24: Phase B copied them into
  `tests/fixtures/theme-colors/` with a `SHA1SUMS` (section 12.3), so CI runs them.

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

## 10. Phase A: implementation notes (2026-10-12)

Phase A is done. The property catalogue mirrors docx4j's `StyleUtil` at `2e381ef26` for the four
elements that carry a colour, the parity goldens move in exactly the two comparisons section 2
measured and nowhere else, and those two are pinned until docx4j-core-ts regenerates the goldens.
No content-API answer over the 64 corpus documents changes.

### 10.1 What landed

| file | lines | what changed |
|---|---:|---|
| `docx4j_py/model/property_catalogue.py` | 1,180 (+87, −33) | `is_empty_color` and `is_empty_u` count the theme attributes; `_states_colour` (docx4j's `statesColour`) and `_merge_colour`, the one rule the four merges share; `merge_color`, `merge_u`, `merge_shd` and `merge_border` rewritten over it; `_theme_colour` (the enum-by-enum rule) removed, nothing else using it; `_empty_scalar` counting `b""` (10.8 item 2) |
| `tests/content/test_properties.py` | 425 (+122, −2) | four tests: the three CR-034 cases of `StyleUtilMergeRulesTest`, and one for this model's spellings (10.3) |
| `tests/content/goldens.py` | 348 (+131, −2) | `Pin`, `PINS`, `PinnedDifference`, `goldens_for`, `each_golden` and `Comparisons`: the pins (10.4) |
| `tests/content/test_properties_goldens.py` | 203 (+33, −16) | its three tests parametrised through `each_golden` and comparing through `Comparisons` |

`merge_p_bdr`, `_merge_tbl_borders` and `_merge_tc_borders` call `merge_border` edge by edge and
needed nothing; `unset_catalogue` and `has_direct_formatting_of` go through the same `is_empty_*`
functions (section 3 item 3). Nothing in the engine, the content API or the facade changed, and
nothing that writes: no Word check, and the `slow` round trip was not run.

### 10.2 The rules as built

1. **Emptiness.** `is_empty_color` is true only when `w:val`, `w:themeColor`, `w:themeTint` and
   `w:themeShade` are all absent (the value, tint and shade also when empty); `is_empty_u` the
   same over `w:val` (the kind), `w:color` and the three. As in the Java, the reference counts when
   it is present at all (`getThemeColor() == null`, not `isEmpty(STThemeColor)`): a `w:color`
   stating only `w:themeColor="none"` is not empty, so its merge creates the destination element
   and changes nothing in it, as docx4j's does.
2. **`_states_colour(value, theme_colour)`** is `not _empty_scalar(value) or not
   _empty_theme_colour(theme_colour)`. `STHexColorAuto.AUTO` is not empty, so `auto` states a
   colour; `none` is not a stated reference.
3. **`_merge_colour(source, destination, names)`** takes the colour's four attribute names, value
   first: `_COLOR` (`w:color`'s `val`, `theme_color`, `theme_tint`, `theme_shade`),
   `_COLOR_ATTRIBUTE` (the same with `color`, for `w:u`, `w:shd`'s pattern colour and every border)
   or `_FILL` (`fill`, `theme_fill`, `theme_fill_tint`, `theme_fill_shade`). Where the source
   states the colour, all four are the source's, `None` included; otherwise the tint and the shade
   go through `_attr` on their own. `merge_shd` calls it once per group, so the two are
   independent and CR-015's rule (a `themeFill` survives a direct `w:shd` stating only a pattern)
   holds. `w:u`'s kind, `w:shd`'s pattern and a border's `val`, `sz` and `space` stay on `_attr`.

### 10.3 The tests

- `test_a_colour_stating_a_value_replaces_an_inherited_theme_reference`,
  `test_a_colour_stating_only_a_theme_reference` and
  `test_shading_fill_border_and_underline_colours_inherit_as_one`: docx4j's three cases, built
  from the typed classes with the hex as `bytes`. The Java's `areEqual` assertions are not ported:
  the catalogue compares nothing (section 3 item 1).
- `test_auto_states_a_colour_and_none_does_not`: corpus document 3229's case through a
  `PropertyResolver` (a paragraph style coloured `accent1` shaded `BF`, a run's own
  `w:color w:val="auto"`: the effective colour is `auto` with no reference and no shade), a
  `w:themeColor="none"` over an inherited accent leaving it whole, `none` alone not empty, and
  `b""` empty.

Each of the four fails on the catalogue before this phase and passes on it after.

### 10.4 The pins

`goldens.PINS` is the one place: a `Pin(test, golden, where, reason)` per expected difference,
`where` being the comparison's label exactly as the test hands it to `same()`. `each_golden`
parametrises a test over `goldens_for(its name)`, which marks a pinned (test, golden) pair
`xfail(strict=True, raises=PinnedDifference)`; `Comparisons.same` holds a pinned comparison's
difference and carries on comparing the rest of the golden, and `finish()` raises
`PinnedDifference` with what it held. So:

- the regeneration, or anything else that makes a pinned comparison agree, turns the pair into an
  unexpected pass, which `strict` makes a failure naming the pin to remove (checked by running the
  goldens with the catalogue of `3b0eb9c`: both pairs `[XPASS(strict)]`);
- any **other** difference in a pinned golden is a plain `AssertionError`, which `raises=` does
  not accept, so it fails exactly as it would without the pin.

That is finer than section 3 asked for (a mark per pair), for a reason found on the day: in the
goldens docx4j-core-ts committed during this session (10.6), `fonts-modesOfApplication`'s pinned
pair also differs in an unrelated comparison, which a mark on the pair would have reported as the
expected failure. The two pins:

| test | golden | comparison | ours | the golden's |
|---|---|---|---|---|
| `test_every_paragraph_and_run_resolves_as_docx4j_resolves_it` | `fonts-modesOfApplication` | `main[5] ... run 3: effectiveRPr` | `<w:color w:val="C0504D" w:themeColor="accent2"/>` | the same with `w:themeShade="BF"`, leaked from `Heading1`'s `accent1` |
| `test_every_table_resolves_as_docx4j_resolves_it` | `invoice2013` | `table 0: effectiveTableStyle` | six edges `w:val="none" w:color="auto" w:sz="0" w:space="0"` | each with `w:themeColor="text1"` beside `auto`, kept from `TableGrid`'s `single 000000 text1` under the table's own `w:tblBorders` |

Each `reason` names the golden, the rule ("a colour inherits as one"; "auto states a colour") and
what it waits for: docx4j-core-ts CR-008 phase 1 regenerating the goldens at docx4j `2e381ef26` or
later. Then both pins come out; the mechanism can stay for the next.

### 10.5 The numbers

Measured first against the goldens docx4j-core-ts had committed when the session began
(`48251956c`, recorded at docx4j `0349796f9`), taken out with `git archive` into a scratch
directory and `goldens.GOLDENS` pointed at it, because its working tree was being regenerated at
the time (10.6):

- `test_properties_goldens.py`: 135 tests over the 45 goldens with a fixture here (of 62
  committed; the other 17 are docx4j-core-ts's table-condition probes, CR-008's, and
  `invoice_Saxon_XPath2`): **133 passed, 2 xfailed**. Its 4,415 comparisons (1,647 `w:pPr`, 2,760
  `w:rPr`, 8 table styles): **4,413 equal, the 2 pinned differ**.
- `test_numbering_goldens.py`: 45 passed, 1 failed (`test_every_numbering_golden_has_its_fixture`,
  10.6, before this phase too). `test_fonts_goldens.py`: 90 passed.
- Over the three files, with a recorder in place of `same()` that does not stop at the first
  difference: **5,147 comparisons, 0 differing before** (the catalogue of `3b0eb9c`) **and 2
  after**, the two of section 2. Nothing else moves.

The fast suite (`-m "not slow"`) gains the 4 tests. Against those goldens: 1,949 passed, 8 skipped,
2 xfailed, 2 failed (10.6). Against the sibling as it stood by the end, with the goldens of
`3c01dbf`: 1,948 passed, 8 skipped, 1 xfailed, 4 failed, where `3b0eb9c` gives 1,945 passed, 8
skipped, 4 failed --- the same four tests, none of them this phase's (10.6).

### 10.6 The sibling checkout on the day

- **The goldens were regenerated during this session.** At 10:19 a docx4j-core-ts session rewrote
  62 of the 64 golden files at docx4j `856afc5da` (the `docx4j-17.3.2` tag) and committed them at
  10:31 as `3c01dbf` (its CR-001 section 22). 17.3.2 does **not** contain `2e381ef26`, so those
  goldens keep the attribute-by-attribute merge and the same two comparisons move against them,
  pinned as above. They also carry docx4j `fb2a764f7` (17.3.2: a paragraph mark's `w:rStyle`
  resolved as a level over the paragraph style), which docx4j-core-ts ported and this repository
  has not: `NumberingIndents/footer:rId10[1]` and `fonts-modesOfApplication/main[8]`,
  `paragraphMarkRPr`, differ before this phase and after it. That is a parity item of its own, not
  this CR's; the second sits in a pinned pair and fails as a plain difference (10.4).
- **Two failures predate this phase** and are the sibling's committed state:
  `test_every_numbering_golden_has_its_fixture` (docx4j-core-ts `9189ab6` added the golden
  `tables-numbered-in-cell`, a CR-008 fixture this repository has not copied) and
  `test_office_js_subset.py::test_the_script_would_write_what_is_committed` (docx4j-core-ts
  `d64ab67`, 2026-09-25, added `Range.hyperlink` after `tests/office_js_subset.json` was derived
  on 2026-09-20).
- docx4j's `2e381ef26` is still local: `origin/VERSION_17_3_3` is at `f6b87cd2f`, 17 commits
  behind (risk 2 stands).

### 10.7 The content API

A throwaway script (not committed) over the 64 documents of `samples/` and `tests/fixtures/`, run
with the catalogue before and after: **10,685 resolver answers** --- the effective `w:rPr` and
`w:pPr` of every style, `effectivePPr` and the paragraph mark of every paragraph and the
`effectiveRPr` of every run in all 113 stories (main, headers, footers, notes, comments; loaded
with `mce_preprocess=True`, so text boxes count), the effective style of every table --- and
**1,422 content-API answers**: `Font.to_dict()` of every paragraph of every body of
`pkg.bodies()`, `pkg.describe()` and `pkg.to_markdown()`. **3 resolver answers change and 0
content-API answers.** The three are section 2's two: `fonts-modesOfApplication`'s run, and
`invoice2013`'s table 0 twice, the document being in both directories byte for byte. The content
API sees neither: `Font.color` reports the effective `w:val`, `C0504D` before and after, and no
content-API read reports a table style's border reference. Section 2's claim holds.

### 10.8 Departures

1. **`areEqual` is not ported** (section 3 item 1), nor the `areEqual` assertions of docx4j's
   three tests.
2. **`_empty_scalar` counts `b""`.** Section 3 item 3 says `_empty_scalar` already handles the
   model's `bytes`. It did for every non-empty value, but `w:val=""` parses to `b""`, which it
   called non-empty where docx4j's `isEmpty("")` is true, so `_states_colour` would have called an
   empty value stated. Fixed in the one place, for every caller (`is_empty_shd`,
   `is_empty_border` and `is_empty_u` too), towards docx4j. No colour attribute in the 64
   documents is empty, so nothing measured moves.
3. **One `_merge_colour` over attribute-name tuples** in place of the Java's four hand-written
   blocks: an idiom, not a semantic; the four merges cannot drift apart.
4. **The pins are per comparison**, not per pair (10.4).

### 10.9 What Phase B needs

- Nothing more from the merge: a colour now reaches the resolver whole, so a resolver over the
  effective properties never sees a reference with a stranger's tint or shade.
- ~~The pins come out when docx4j-core-ts regenerates at `2e381ef26` or later.~~ Done the same
  day (10.10).
- 10.7's measurement is the baseline for `Font.color` reporting the drawn colour: Phase B should
  move content-API answers only where a theme reference meets a stale cache, and the same
  measurement says where.
- `CHANGELOG.md`'s Unreleased entry records Phase A; Phase B adds to it.

### 10.10 The goldens regenerated, the pins out (2026-10-12)

At this repository's request (Jason's), docx4j-core-ts ported the same rules (its CR-008 phase 1,
`9f4fab9`) and regenerated its 62 parity goldens and 62 XPath goldens at exactly docx4j
`2e381ef26` (`1f78b94`; `docx4j-core-17.3.3-SNAPSHOT.jar`, SHA-256 `f805ae4be89897e3`, every header
naming `2e381ef26`). It reports that exactly the two comparisons of section 2 moved, as measured
here, and nothing else; zero differences on its side, 811 of 811.

Here the two pinned pairs became `[XPASS(strict)]` failures, as 10.4 meant them to, and the two
`Pin`s came out: `goldens.PINS` is empty, the mechanism kept for the next time this package is
ahead of the goldens. `invoice2013` passes; `fonts-modesOfApplication` main[5] run 3 agrees, and
the pair still fails at main[8], the docx4j `fb2a764f7` difference of 10.6, now with nothing
pinned in front of it. The three golden files: properties 133 passed and 2 failed (`NumberingIndents`
and `fonts-modesOfApplication`, both `fb2a764f7`), numbering 45 passed and 1 failed (the fixture
`tables-numbered-in-cell`, 10.6), fonts 90 passed.

docx4j `2e381ef26` is still local; until docx4j pushes `VERSION_17_3_3`, docx4j-core-ts's weekly
parity run, which measures the pushed branch, will propose moving these two back, and that session
closes such a PR unmerged. Nothing here depends on it: this repository reads the sibling checkout.
*2026-10-13:* the branch is pushed with `2e381ef26` in it, so the weekly run now measures the same
rules.

## 11. docx4j-core-ts has ported the colour resolver: what Phase B now has to follow (2026-10-24)

docx4j-core-ts's session reported it at the user's word: CR-008 phase 2, `27be4a7`, pushed to its `main` at `7976c42` and not released. It covers the ground of this CR's Phase B. Under CR-003 section 23.3's parity, its behaviour is owed here, and Phase B now has a second reading beside docx4j's own.
- **`Font.color`'s effective view reads a theme colour as Word draws it.**
  - It takes the effective `w:color`'s `w:themeColor`, with its tint or shade (HSL luminance; tint wins), through the theme and `w:clrSchemeMapping`, over the `w:val` cache.
  - Word's own value is returned where it is within one unit (a current cache).
  - Check 53's Heading 1 reads `#0F4761` (it was `#365F91`), as Office JS does.
  - Automatic still reads `#000000`.
  - **The direct view** (`get_font(direct=True)` here) keeps the written `w:val`, core-ts's D2 as the user decided it there. That is this CR's question 2's recommendation in effect.
- **The port** is all of docx4j's `org.docx4j.model.color` as it stands at docx4j `1fffc20a3`: `ResolvedColor`, `ThemeColors` (docx, slides with their masters' maps and overrides, xlsx, charts), `WordColors`, `DrawingMLColors` (`691620fb7`'s luminance rounding), `SpreadsheetColors` and `StyleFill`.
  - `PropertyResolver` gains `getThemeColors()`, `resolveColor` and the rest.
  - Only the WordprocessingML half is this package's until CR-001 Phase D and CR-002 Phase C.
- **The vectors** are docx4j-core-tests' `theme-colors/`, copied into core-ts's `test/theme-colors/` with a SHA-1 manifest. Every probe's tally there equals docx4j's own `ThemeColourProbesTest`, family by family.
  - The corpus triples come out 76 exact and 41 within one, of 117.
  - docx4j's javadoc ("119 ... 77 exact") is stale.
- **A package with no theme** reads the default theme its `fonts.defaultTheme` names; the 2007 and 2013 themes both give `4472C4` (docx4j CR-034 D2).

**For Phase B here:** port docx4j's package as core-ts did, held to the same vectors copied from docx4j's tree with their SHA-1s. `Font.color`'s effective view then moves exactly where 10.7's measurement says a theme reference meets a stale cache. CR-013's check 53 assertion of `font.color` (`#000000` against `""`, listed "CR-009's" in CR-013 section 6.1) and CR-016's `Font` reading every run both take the new answer. When Phase B is scheduled is the user's call, including whether 0.2.0 waits for it.

**Scheduled 2026-10-24** (the user: "0.2.0 waits for CR-009 Phase B"). Phase B ports docx4j's `org.docx4j.model.color` as section 11 says, after CR-018 Phase B, whose colour reader and merge it sits on, and before or with CR-016 Phase B, whose `Font` reads every run. *Implemented the same day: section 12.*

## 12. Phase B: implementation notes (2026-10-24)

Phase B is done. docx4j's `org.docx4j.model.color` at `a9e0303ff` is ported, its WordprocessingML half, under `docx4j_py/model/color/`, as docx4j-core-ts ported the whole at `76fb45d` (`27be4a7`, with `76fb45d`'s `normalise` and `stated`). `PropertyResolver` gains the theme colours and the resolvers, and `Font.color`'s effective view reads a theme colour as Word draws it (question 2). Every probe case is held to Word's reading at docx4j's own tallies, family by family. Over 728 corpus documents, 119 effective `Font.color` answers move, each a theme reference meeting a stale cache or `auto`, and no direct answer moves. Nothing written changes, so there is no Word check (section 4).

### 12.1 What landed

| file | lines | what |
|---|---:|---|
| `docx4j_py/model/color/__init__.py` | 36 | the package: `ResolvedColor`, `ThemeColors`, `WordColors`, `OPAQUE` |
| `docx4j_py/model/color/resolved.py` | 104 | `ResolvedColor`: a frozen, slotted dataclass of `red`, `green`, `blue` and `alpha` (OOXML's unit, 100,000 opaque, docx4j's D5), clamped as docx4j's constructor clamps; `rgb`, `of_hex`, `with_alpha`, `is_opaque`, `hex`, `within`, `BLACK`, `WHITE`, `str` as `RRGGBB[/alpha]` |
| `docx4j_py/model/color/color_math.py` | 65 | `ColorMath`'s HSL half: `rgb_to_hsl`, `hsl_to_rgb` (the saturation unclamped), `clamp01`, expression for expression, Java's `%` as `math.fmod` with the same correction |
| `docx4j_py/model/color/tables.py` | 63 | `ColorTables.SYSTEM` (`a:sysClr`'s Windows defaults) alone (12.4 item 1) |
| `docx4j_py/model/color/theme_colors.py` | 319 | `ThemeColors`: `of(package)`, `of_theme(theme, default_theme)`, `with_word_mapping`, `slot`, `scheme_color`, `word_theme_color`, `mapping`; the slot reader `_slot_colour`, Phase C's seam |
| `docx4j_py/model/color/word_colors.py` | 217 | `WordColors`: `resolve`, `of` (a `w:color`, a border, a `w:u`), `fill`, `pattern`, `luminance`, and `normalise`/`stated` over `values.colour`/`states_colour` |
| `docx4j_py/model/property_resolver.py` | +60, −2 | `theme_colors` (and `get_theme_colors()`), made once and dropped by `refresh()`; `resolve_color`, `resolve_fill`, `resolve_pattern_color` |
| `docx4j_py/model/content/font.py` | +94, −14 | the effective read's drawn colour (`_drawn`), `to_dict()`'s `theme_color`, `theme_tint` and `theme_shade` (`_theme_attributes`, `THEME_KEYS`), in `Font` and `MixedFont` alike |
| `tests/content/test_colors.py` | 689 | 27 tests (12.3) |
| `tests/fixtures/theme-colors/` | 7 files | docx4j's vectors and `SHA1SUMS` (12.3) |
| `tests/content/test_ranges.py` | +20, −6 | check 53 compares `font.color` (12.6) |
| `tests/content/test_font.py` | +4 | the direct `to_dict()` with its three new keys |

`model/` is in `codegen/clean.py`'s `KEEP` whole, so `model/color/` needs no entry; `generate.sh --check` was run on the commit in a scratch worktree and left `model/color/` and the rest of the tree as committed (12.7). The python-docx facade's `ColorFormat` is untouched (question 3), and so are the engine, the writers and the merge.

### 12.2 The rules as built

1. **The reference wins.** `WordColors.resolve(val, theme_color, theme_tint, theme_shade, theme)`: the written value is `ResolvedColor.of_hex(val)`. Where there is no reference, the reference is `none`, or there is no theme, the written value is the answer (None for `auto`, absent or rejected: automatic, the caller's default). Otherwise the reference's colour is `theme.word_theme_color(theme_color)`: `text1`, `background1`, `text2` and `background2` go through the map, and `dark1` to `followedHyperlink` name their slots directly, accents and hyperlinks included, as docx4j's `wordThemeColor` does. A reference that names no colour also leaves the written value.
2. **Tint and shade** are float HSL luminance, `L·t + 1 − t` and `L·t`, with `t` the byte over 255. The tint is applied where both are present; one that is not a byte (`zz`) falls to the shade. The channels are truncated (`floor(clamp01(v) * 255 + 1e-9)`). A byte is read as Java's `Integer.parseInt(s, 16)` reads one, an optional sign and hex digits, where Python's `int(s, 16)` would also take `0x`, spaces and underscores.
3. **The cache rule.** Where the written value is within one unit per channel of the computation, the written value is returned (Word's own rounding); two units off, the computation.
4. **A value is read by the one reader.** `WordColors.normalise(val)` is `values.colour(val)` with its `AUTO` as None, and `stated` is `values.states_colour`. There is no second table: docx4j's names live in `WordColorValues.HTML` since `a9e0303ff`, not in `ColorTables`, and they equal `NETSCAPE_COLOURS` exactly, all 140 names and values (asserted from docx4j's source at `a9e0303ff`). `of`, `fill` and `pattern` normalise first, so the model's `bytes` (with their source spelling), the `auto` member and a kept `str` are all read as Word reads them.
5. **The theme.** `ThemeColors.of(package)` reads the main document part's theme part, or the bundled theme `package.default_theme` names where it has none (the process-wide setting where that is None; 2007 and 2013 both give `4472C4`, docx4j's D2). It then applies `w:settings/w:clrSchemeMapping` as `with_word_mapping`. Both parts are read with `read_contents()`, so neither is unmarshalled (asserted). A part that cannot be read is logged and taken as absent, as docx4j's is.
6. **Enumerations**, `w:themeColor` and `w:clrSchemeMapping`'s `ST_ColorSchemeIndex`, go through `values.in_enumeration`. A value outside the enumeration, such as `Accent1` kept as text by the exact parse, names nothing. That is docx4j's reading too (JAXB unmarshals it as null).
7. **The resolver.** `PropertyResolver.theme_colors` is made on first use and kept until `refresh()`. A resolver over a `w:styles` tree alone has the process default theme and the default map. `resolve_color(obj)` is docx4j's three `resolveColor` overloads, told apart by class. `resolve_fill` and `resolve_pattern_color` are its other two. The effective properties keep the attributes as written, as docx4j's do.
8. **`Font.color`**, in the effective view: where the effective `w:color` has a `w:themeColor` that is not `none`, the answer is `#` plus `resolve_color`'s hex, unless that is automatic. Otherwise it is read as before (CR-018 rule 3: `#RRGGBB`, `auto`, `""`). The direct view is unchanged, `w:val` as written (core-ts's D2). A range across paragraphs reads each run the same way (`MixedFont`). `to_dict()` gains `theme_color` (the member's value, or a kept string as written), `theme_tint` and `theme_shade`, each None where absent: the effective `w:color`'s in the effective view and the run's own in the direct view. In `MixedFont` each is the runs' common value, else None.

### 12.3 The vectors and the tallies

`tests/fixtures/theme-colors/` holds `README.md`, `cases.tsv`, `readings.tsv`, `word-tint-shade.tsv`, `tc-p1-precedence.docx` and `tc-p2-colour-map.docx`, copied by `git -C ../docx4j show a9e0303ff:docx4j-core-tests/src/test/resources/theme-colors/<file>`. Each is its git blob there, unchanged since `3dbe4eee5` (added) and the same at `1fffc20a3`, and has the SHA-1 core-ts's `manifest.json` gives at `76fb45d`. With them is a `SHA1SUMS`. **This replaces section 4's "consumed from the sibling checkout" with section 11's "copied with their SHA-1s"**, as `tests/fixtures/onoff-fun/` was copied for CR-018. So CI runs the vectors, and section 8's last risk is closed. The files are docx4j's, Apache-2.0. P3 to P5 are Phase C's and not copied.

| vector | here | docx4j | core-ts |
|---|---|---|---|
| `word-tint-shade.tsv`, 117 triples | 76 exact, 41 within one, 0 further | 76, 41, 0 (its javadoc's "119 ... 77" is stale) | 76, 41, 0 |
| P1, 17 cases, by family | paragraph shading 2/1/0, pattern blend 1/0/0, run shading 1/0/0, run text 8/0/0, top border 3/0/0, underline 1/0/0; the one off P1-10 (ours `F6C6AD`, Word `F6C5AC`) | the same | the same |
| P2, WordprocessingML (`run text`) | 21/0/0 | 21/0/0 (its shapes, 7/0/0, are Phase C's) | 21/0/0 |
| the cache rule's check (`2E74B5` under accent1 `5B9BD5` shaded `BF`) | `2E74B5` | `2E74B5` | `2E74B5` |

**Through a `PropertyResolver`.** docx4j's `p1Precedence` and `p2ColourMap` resolve each case's attributes as written. Here each case is found by its label, as docx4j finds it (every `w:p` of the main part in both branches of each kept `mc:AlternateContent`, its text every `w:t` beneath it). It is then resolved through the probe's resolver, over the effective `w:rPr` or `w:pPr`, by `resolve_color`, `resolve_fill` and `resolve_pattern_color`. A family is counted where the paragraph itself states it, as docx4j counts. The tallies are docx4j's to the case. P2's 21 `run text` readings are the six runs, the paragraph anchoring the shapes, and the seven shapes' labels in each of their two branches (core-ts CR-008 section 7.4 item 3). Through `Body` (core-ts's test), `Font.color` is what Word drew for P1's 17 paragraphs and P2's six, compared as `#RRGGBB` where the run states a colour. Where it states none, the reading is `""` (12.9).

**The ports, by name** (`tests/content/test_colors.py`, 27 tests):
- `ThemeColorsTest`: `newDocxHasTheOffice2023Theme` and `noThemePartTakesTheDefaultTheme` (2023 `156082`, 2013 and 2007 `4472C4`, with core-ts's `2F5496` through a refreshed resolver). Its sheet-index and chart cases are Phase C's.
- `ThemeColourOutputTest`: `resolverGivesTheColourWordDraws` (with core-ts's underline and automatic cases, made once, read again on refresh), `aDirectValueReplacesTheStylesThemeReference` and `aThemeColourWithNoValueIsInherited`. Its `FontColor` CSS and HTML tests have no counterpart here (12.4 item 6).
- `WordColorsNamesTest`: `namesAndHexSpellingsInEveryColour` and `aRejectedValueIsAutomatic`.
- `ThemeColourProbesTest`: `p1Precedence`, `p2ColourMap` (its WordprocessingML half) and `wordTintAndShadeAgainstTheCorpora`.

Then the units:
- `ResolvedColor`;
- `none` giving the written value or None, and a reference outside `ST_ThemeColor` likewise;
- the cache rule, one unit off returned as written and two units replaced;
- the tint winning over the shade, an unusable tint falling to the shade, and `0x99` not a byte;
- `w:clrSchemeMapping` swapping `text1`/`background1` and `text2`/`background2` (the accents named directly, a DrawingML `schemeClr` through the map, the original untouched);
- the seam (12.4 item 1);
- `normalise` and `stated` being `values.py`'s, with no table of their own;
- the tables being docx4j's (`WordColorValues.HTML` and `ColorTables.SYSTEM` at `a9e0303ff`, read with `git show`, skipped without `../docx4j`);
- `to_dict()`'s keys, effective, direct and across paragraphs;
- a colour written through `Font` having no reference;
- a detached run read direct;
- the theme and settings parts not unmarshalled;
- the vectors as copied, and each copied file docx4j's blob.

### 12.4 Departures, and why

1. **The WordprocessingML half only**, as scheduled. `DrawingMLColors`, `SpreadsheetColors`, `StyleFill`, `ColorTables.PRESET` and `INDEXED`, `ColorMath`'s sRGB, Foley and integer-HLS halves, and `ThemeColors`' slide, workbook and chart hosts (`with_mapping`, `with_override`, `with_scheme`, `for_chart`, `with_placeholder`, `sheet_theme_color`, the format scheme, `phClr`) are Phase C's. **The seam:** docx4j reads each scheme slot through `DrawingMLColors.resolve(slot, null)`. Here `theme_colors._slot_colour` reads an `a:srgbClr` or an `a:sysClr` with no transform. An `a:sysClr` reads as Windows' colour for its name, else its `lastClr`. Any other slot (the other four kinds, or any transform) is left unresolved and logged. A reference to it then reads as the value written beside it, as before this phase. Measured on 2026-10-24, the 785 theme parts of `samples/`, `tests/fixtures/` and `out/cr012-b-corpus/` (950 packages) hold 9,372 slots: 7,814 `a:srgbClr` and 1,558 `a:sysClr` (`windowText` and `window`), with no transform. So nothing in the corpus reaches the unported path. Phase C replaces `_slot_colour` with `DrawingMLColors.resolve`.
2. **Python names.** Getters are attributes, `toHex` is `hex()` and `ofHex` is `of_hex`. `of(Theme)` is `of_theme(theme, default_theme)`, since this package's default theme is per package (CR-002 section 15.3). docx4j's overloads by type, `WordColors.of(Color | CTBorder | U)` and `PropertyResolver.resolveColor`'s three, are one method each, told apart by class: Python can, where core-ts could not and split them into `of`/`ofBorder`/`ofUnderline` and `resolveColor`/`resolveBorderColor`/`resolveUnderlineColor`. `getThemeColors()` is the `theme_colors` property, with `get_theme_colors()` beside it. `luminance`, package-private in docx4j, is public, as core-ts made it, for the corpus test. `ThemeColors.mapping` (read-only) is new, for a test to read the map.
3. **A memo.** Each `ThemeColors` keeps its computations by (reference, tint, shade), up to 4,096, since the theme is immutable. docx4j recomputes every time; the answers are the same. It about halves what a theme-coloured read adds (12.8: `resolve_color` 5.8 µs without it, 3.2 with it).
4. **`of_hex` takes six ASCII hex digits**, core-ts's reading of Java's `parseInt` over a six-character string. Java's also takes a sign (`+FFFFF`) and non-ASCII digits, which no colour Office writes has.
5. **The P1 and P2 cases are resolved over the effective properties** (12.3), where docx4j resolves the attributes as written: the task's "through a `PropertyResolver`". The tallies are the same.
6. **docx4j's output tests** (`propertiesWriteTheDrawnColourOnceGivenTheTheme`, `htmlDrawsTheThemeColours`) test its CSS property classes and HTML pathways, which this package does not have.
7. **The vectors are copied** (12.3), where section 4 consumed them from the sibling.

### 12.5 The corpus measurement

`Font.color`, effective and direct, of every paragraph of every body (`package_bodies`: the main part, headers, footers, notes and comments) and of every run in them, each run through its own `Font`. It was measured over `samples/` and `out/cr012-b-corpus/` (728 packages, the latter holding docx4j's, core-ts's and this repository's fixtures at the commits its `COMMITS.txt` names), at `a061d6c` (a scratch worktree) and on this phase's code, loaded by default. That gives **58,555 readings each way, 0 load or read errors**.

- **Direct: 0 moved.**
- **Effective: 119 moved, 119 explained, 0 unexplained.** Each was checked mechanically. The effective `w:color` has a `w:themeColor`. The answer before was its `w:val` as `colour()` reads it. The answer after is `WordColors.resolve` over the document's `ThemeColors`. The `w:val` is more than one unit from that: a stale cache (117), or `auto` (2, P1-04's run, Word's `156082`).
- **Where**: 18 packages, all in `out/cr012-b-corpus/` and none in `samples/`. By reference: `accent1` shaded `BF` 52, `accent1` plain 30, `text2` shaded `BF` 14, `hyperlink` 11, `text1` 4, `accent1` tinted 2, `background1` 2, `text2` 2, `background2` 2. They come from 37 distinct (package, `w:color`) pairs:
  - **The Office 2010 caches of docx4j's default styles under an Office 2023 theme**, as a package docx4j's or core-ts's `createPackage` made has them: Heading 1's `365F91`/`accent1`/`BF` to `0F4761` (stale by 48 in a channel), Title's `17365D`/`text2`/`BF` to `0A1D30` (by 45), `4F81BD`/`accent1` to `156082` (by 59), and the Hyperlink style's `0000FF`/`hyperlink` to `467886` (by 121). They are in `revisions-base.docx` (26), core-ts's `paste-samples.docx` (18), the three copies of `review-for-word.docx` (30), `cr043-base.docx` (4), `37b-bound.docx` (3), the three `29-input.docx` (6), the two `33-input.docx` (4), the two `extraction-probe-html.docx` (4) and the two `hyperlink.docx` (2). These are Office 2010's caches, which section 10.7's measurement placed in the styles.
  - **The probes**, by design: `tc-p1-precedence.docx` 14 (P1-01 to P1-08 but P1-05, whose cache is current, each read through its paragraph and its run) and `tc-p2-colour-map.docx` 8 (P2-01 to P2-04, the four swapped references, likewise).
- So 10.9's expectation holds: only a theme reference meeting a stale cache moves.

The script, the two JSONs and the explanation are in the session's scratchpad, not committed, as 10.7's were.

### 12.6 Check 53

`test_ranges.py`'s check 53 now compares `font.color` with the rest, where it was excluded wholesale. Before this phase, 12 of the 13 `font.color` readings differed from Word's Simple Markup run; after it, 9 do.
- **Now Word's:** cases 3 to 5, Heading 1's `365F91` under `accent1` shaded `BF`, read `#0F4761`, as Word reads them and core-ts does (asserted).
- **Remaining:**
  - the eight automatic colours (cases 2, 7, 8, 9, 19, 20, 21 and 24): `""` here, `#000000` in Word and core-ts (12.9);
  - case 31: a paragraph's font reads its first run, which is red, where Word reads `""` for runs that differ. That is CR-016 Phase B's (core-ts CR-002 section 59), listed with its `font.highlightColor`.

Case 1, the colour across two paragraphs whose runs differ, was `""` in both before and still is. CR-016 Phase B, which will read `Font` over every run, is not built here and not blocked: the drawn colour is in `Font._read`, which `MixedFont`'s readings already go through.

### 12.7 The gates

| gate | before (`a061d6c`) | after |
|---|---|---|
| the full suite | 6,260 passed, 8 skipped, 0 failed (CR-016 section 10.8's at `cfde1db`; `a7c2da6` and `a061d6c` changed only records) | **6,287 passed, 8 skipped, 0 failed** (4 min 5 s) |
| the fast suite (`-m "not slow"`) | 4,953 passed, 8 skipped, 1,307 deselected (the same record) | **4,980 passed, 8 skipped, 1,307 deselected** (2 min 37 s) |
| the new tests | --- | 27 in `tests/content/test_colors.py`; `test_ranges.py` 37 as before, check 53 now comparing `font.color` (12.6); `test_font.py`'s direct `to_dict()` gaining its three keys, the one pre-existing expectation changed |
| `roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` over `samples/` | `--normalise` and `--no-normalise` (run in the `a061d6c` worktree): 50 of 50, 0 differences, 34 `boolean-spelling`, 0 kept, 0 normalised, 0 skipped | the same, both ways (`--out "$PWD/out/cr009b-roundtrip-normalise"` and `-no-normalise`) |
| `codegen/generate.sh --check` | --- | byte identical, 307 modules (301 before, and the six of `model/color/`), run on `c441b67` in a scratch worktree (`TMPDIR` in the scratchpad), whose tree it left as committed: `git status` clean, `model/color/` kept |
| acceptance (`compare_acceptance.py`) | a regeneration at `a061d6c` (in its worktree) | `out/cr009b-acceptance-after/` against a copy of `out/acceptance/` taken first, and against the `a061d6c` regeneration: 17 artefacts, 243 entries, 0 differing beyond the dates, both. No artefact is a `describe()` or `to_dict()` output, and none moves |
| the corpus `Font.color` measurement (12.5) | 58,555 readings | 119 effective moved, all explained; 0 direct moved |
| check 53's `font.color` (12.6) | 12 of 13 differing from Word | 9: the 8 automatic and case 31's first run |

Nothing written changes, so the goldens and the round trip cannot move: the resolver's effective properties are unchanged, and only `Font`'s reading of them is new. `out/acceptance/` was not regenerated, since this phase changes nothing CLAUDE.md lists for that. No Word check (section 4).

### 12.8 The cost

The resolver's theme colours are made once per resolver, and again after `refresh()`. The measure is `bench_colour.py`: every paragraph's and run's effective `Font.color`, read after a warm read, timed as the median of 7 runs, two rounds interleaved. The script is in the scratchpad.

| | before (`a061d6c`) | after |
|---|---|---|
| `Symbols.docx` (4,062 reads, 9 of them theme-coloured) | 168.1 and 164.8 ms (41.4 and 40.6 µs a read) | 165.8 and 172.0 ms (40.8 and 42.3 µs) |
| 2,000 Heading 1 paragraphs with a bold run each, every read theme-coloured (6,000 reads) | 309.2 and 304.6 ms (51.5 and 50.8 µs) | 350.6 and 363.2 ms (58.4 and 60.5 µs) |
| the first read after a `refresh()` | 0.87 ms | 0.93 ms |

A document with few theme colours reads within the noise. Where every read is theme-coloured, a read costs about 7 to 10 µs more (14 to 19%). In isolation the drawn colour is about 3 µs of that: `resolve_color` 3.2 µs with the memo (5.8 without it), and the theme attributes 0.2 µs. The resolver's `get_effective_rpr`, about 36 µs, is most of any read either way. Making the theme colours costs about 0.06 ms once per resolver.

### 12.9 The open question: automatic

**Automatic text reads `""` (none stated) or `"auto"` here, and `#000000` in Word and in core-ts.** core-ts reads it that way since its CR-002 section 59; Word 365 does in check 53's eight cases and in P1 and P2. This phase keeps CR-018 rule 3 as the task requires: the effective read of an automatic or unstated colour is unchanged. CR-018 section 4.9 item 1 says whatever CR-016 decides for `auto` becomes the answer for every automatic spelling.

**Question for the user** (not decided here): should the effective `Font.color` read automatic as `#000000`, as Office JS and core-ts do, with the direct view keeping `""` and `"auto"`? *If yes,* check 53's eight automatic differences close, P1's and P2's automatic paragraphs read Word's `#000000`, and `to_dict()`'s `color` changes for every paragraph whose colour is automatic, which is most of them. That belongs with CR-016 Phase B.

### 12.10 Found, and not done

1. **The merge reads a kept invalid `w:themeColor` as a stated reference.** Under the exact parse a value outside `ST_ThemeColor` (`Accent1`) is kept as its text. The property catalogue's `_empty_theme_colour` (CR-009 Phase A, CR-018's merge) counts it as stated, so a run's `w:color w:val="bogus" w:themeColor="Accent1"` replaces a style's `00FF00`, and `Font.color` reads `""` (checked 2026-10-24). docx4j, whose JAXB reads the value as null, takes the style's colour, as this resolver would if the merge had kept it. Word's reading is unmeasured. A one-line change towards docx4j (read the reference through `values.in_enumeration`), left to a probe or a decision.
   **Fixed 2026-10-24** (the coordinator, on verifying this phase): CLAUDE.md's rule already says every enumeration reader goes through `values.py`, where a value outside the enumeration is unset, and docx4j is the oracle, so it needed no new decision. `values.theme_colour` reads `ST_ThemeColor` through `in_enumeration`, and `_empty_theme_colour` uses it. So `w:val="bogus" w:themeColor="Accent1"` under Heading 1 reads the style's `#0F4761`. A valid `w:val` beside the invalid reference stands (`#00FF00`), and `to_dict()` keeps the reference as written. Held by `test_a_reference_outside_st_theme_color_states_nothing_in_the_merge_so_the_styles_colour_shows`. Over `samples/` and `out/cr012-b-corpus/` (723 packages) no `w:themeColor` or `w:themeFill` lies outside the enumeration, so no corpus answer moves. Word's own reading still needs a probe, which rounds 1 to 7 did not cover.
2. **`to_api_script` writes a theme-coloured run's cache** (CR-003 section 3.11). It reproduces `w:val`, so P1-03's script writes `font.color = "#FF0000"` where Word draws `0F4761`. `Font` cannot write a theme reference. Writing the drawn colour would make the script draw as Word does but drop the reference.
3. **`describe()` reads no font**, so section 4's "`describe()` follows `to_dict()`" has nothing to follow, and no `describe()` answer moves.
4. Java's `Character.digit` takes non-ASCII digits in `WordColorValues.normalise` and `hexByte`. `values.colour` and `_hex_byte` take ASCII only, which is CR-018's reading and core-ts's. No colour Office writes has one.
5. **The cost** (12.8) falls on CR-016 Phase B, which will read every run: about 3 µs per theme-coloured run.
