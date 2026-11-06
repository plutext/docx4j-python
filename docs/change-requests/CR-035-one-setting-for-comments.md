# CR-035: One setting for comments --- `comments: true` alone turns a comment story on, and `comments` leaves `include`

**Status:** Complete 2026-11-06: Q1 to Q4 decided as recommended (the user: "CR-035's Q1 to Q4 as recommended, implement when ready"), Phases A and B built the same day (section 8). Proposed 2026-11-05, at the user's word ("why do we have both `comments: true` and `include: comments`? could we have just a single setting?", then "yes" to the recommendation). It closes CR-027 section 13.9's open point 1: a caller's `comments=False` under `profile: word-text` does not hold. No Word check: nothing saved changes. The WD's change is committed (0.91, portfolio `9bb9299`; section 3.1).

**Depends on:**
- **The WD** (`../docx4j-portfolio/docs/text_extraction_specification_wd.md`): the change in section 3.1, sent to the docx4j-portfolio session on 2026-11-05 and committed as 0.91 at `9bb9299` on 2026-11-06, read against section 3.1.
- **CR-027** (complete): `profile: word-text` and `_explicit`.

**Counterpart:** docx4j `org.docx4j.model.text.TextOptions` and docx4j-core-ts `src/model/text/TextOptions.mts`, the same change in each, theirs to schedule. core-ts has built it, as its CR-012 section 29 (`80a4707`, local), with one `log.warn` a process where this CR has a `DeprecationWarning`. docx4j-mcp's `TextOptionsSupport` handles these options too.

## 1. Why

Comments can be switched on in two places, `comments: true` (REQ-504) and `include: comments` (REQ-501), and either one turns them on. `word-text` sets both. So a caller who overrides one finds the other still on, with no warning:

| caller writes | comments today | with one setting |
|---|---|---|
| `profile="word-text"` | in | in |
| `profile="word-text", comments=False` | **in** | out |
| `profile="word-text", include=["headers", "footers"]` | in | in: `include` no longer speaks for them |
| `profile="word-text", comments=False, include=["headers", "footers"]` | out | out |

Read literally, this breaks nothing: the caller's `comments` does override the preset's, and `to_dict()` shows it false. It is the preset's separate `include` that keeps comments on. But the call looks as if it turns comments off and doesn't, and REQ-1002's point is that a caller's option wins.

There is no recorded reason for the two. Both are in the WD from 0.1 (`548d41e`), in separate sections, and neither the WD's history nor docx4j's CR-046 says why. They have only ever meant the same thing.

## 2. What exists here (measured 2026-11-05 at `36ba4bd`)

- `TextOptions` (`model/text/options.py`) holds `_comments` and `_include`. `comments` reads `_comments or "comments" in _include`, `includes` adds `comments` where `_comments` is set, and `to_dict()` reports the two apart.
- `_WORD_TEXT` sets both. `include` is an option, a keyword, a mapping key and a fluent method, `TextOptions().include("headers", "comments")`.
- `extractor.comments_included()` reads `options.comments` alone, so the extractor needs no change.
- **Released:** both spellings are in 0.2.0, and in docx4j 17.3.5.
- **The conformance suite:** none of the copied `phase0/` and `dpp/` rows uses `include: comments`; eight use `comments: true` (`expected-inspection.tsv`, `expected-tokens.tsv`). `tests/text/conformance.py`'s `scenario_requirements` has a branch for the former, and `test_conformance.py` tests it.
- **Tests naming it:** `test_mce_levels.py` (three), `test_text_markdown.py`, `test_private_readers.py`, `test_text_options_profile.py` (the preset's table, `test_include_is_recorded`, and `test_comments_false_under_the_preset_is_docx4js_reading`, which pins today's reading).

## 3. Design

### 3.1 The WD

The change sent to the portfolio session:
- **Section 10's table:** `include` is a set of `headers`, `footers`, `glossary`.
- **REQ-501:** headers, footers and the glossary are available through `include`; comments through `comments` (REQ-504).
- **REQ-1002:** `word-text` is `include: [headers, footers]`, `comments: true`.
- **The open item at REQ-801's 0.67:** `include: comments` reads `comments: true`.
- **A compatibility note:** an implementation that has released `include: comments` may keep it as a deprecated spelling of `comments: true`.

### 3.2 One setting (D1)

`_comments` is the only state. `_include` never holds `comments`. `comments` reads `_comments`, and `includes` still adds `comments` where it is set, since it answers which stories are in scope.

### 3.3 The old spelling (D2, Q1, Q2)

`include` keeps accepting `comments` in every form it takes today, as a spelling of `comments: true`: it is taken out of the list, `_comments` is set, and a `DeprecationWarning` names the new form. It is a recorded caller's option, so it stands against a preset as `comments=True` does.

Where one constructor call or one mapping says `comments` false and `include` naming `comments`, it contradicts itself, and is refused with `TextOptionError` (reason `invalid`, REQ-504). Sequential calls are not a contradiction: the last one stands, as for any option.

This shape is also what a released `to_dict()` writes after `include=["comments"]` alone: `{"comments": false, "include": ["comments"]}`, which reads as comments on (measured 2026-11-06 here; core-ts found it in its own `toMap()`, its CR-012 section 29.2). So such a saved mapping is refused too, and the message names the fix (Q4).

### 3.4 The preset and the output (D3)

`_WORD_TEXT` becomes `include: (headers, footers)`, `comments: true`. A plain `word-text` call reads the same. `to_dict()` then never lists `comments` under `include`, and its `comments` is the one setting, so options still round-trip by name. A saved mapping with the old form loads through 3.3, but for the shape that section refuses.

### 3.5 Not here

- **`notes` and `text_boxes`** are booleans already, and stay.
- **Keeping `include: comments` and dropping the boolean** (a comment is a story) was considered and left: `notes` would stay a boolean anyway, and the WD says `comments: true` five times against `include: comments` once.
- **Removing the old spelling** is for a later release (Q2).

## 4. Tests and oracles

- The table of section 1, each row, as the contract.
- The old spelling in every form: a keyword, a mapping, `set`, `of` and the fluent method, each setting `comments`, none leaving `comments` in `to_dict()["include"]`, each warning.
- The contradiction refused, and sequential calls the last one standing, either order.
- A preset under each: `comments=False` before and after `profile`.
- `to_dict()` round-trips; a mapping in the old form loads; the mapping a released `to_dict()` writes after `include=["comments"]` alone is refused, the message naming the fix.
- `test_comments_false_under_the_preset_is_docx4js_reading` is replaced, and the preset's table follows the WD. Tests that use the old spelling move to `comments=True`, but those holding the old spelling itself.
- The suite, the agreement table and the goldens: no row should move, since no row names `include: comments` and the extractor reads the same value.
- **Until the WD and docx4j change**, the preset's oracle is the WD's text, not docx4j's `WORD_TEXT_VALUES`, so a test comparing them states it.

## 5. Phases

| Phase | What | Effort |
|---|---|---|
| A | 3.2 to 3.4 and the tests; `scenario_requirements`' branch stays while a re-pinned suite could still name it | ¼ day |
| B | The records: CHANGELOG under Changed and Deprecated, the `TextOptions` docstring, the PyPI readme if it names the old form, CR-027's open point closed, the differences document if a port lags | ¼ day |

## 6. Risks

1. **Callers of the old spelling** get a warning where they had none, and an error where a test suite treats `DeprecationWarning` as one, as this repository's does. It is the point of the warning.
2. **The ports differ for a while.** core-ts has built this, and docx4j and this package haven't, which is a difference to record while it lasts.
3. **The order of work.** Building before the WD changed would have put this package against REQ-1002's text, so Phase A waited on it (Q3). It has changed.

## 7. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | How is the old spelling flagged: `DeprecationWarning`, or a log line? | **`DeprecationWarning`.** It is the caller's code that has to change. `listnumbering/formats.py` logs instead because a document's content, not a caller, triggers it, and that reason doesn't apply. This would be the package's first `warnings.warn`. |
| Q2 | How long does the old spelling stay? | **Until 1.0, or a later CR.** It is released, and removing it buys little. |
| Q3 | Does Phase A wait for the WD? | **Yes.** It is one change, and the WD is the contract. It needn't wait for docx4j or core-ts. |
| Q4 | A self-contradicting call: refuse, or let `comments=False` win? | **Refuse, the message naming the fix.** Quietly picking one is the trap this CR removes. Amended 2026-11-06: it isn't only a mistake, since a released `to_dict()` writes this shape (3.3). Letting `comments: false` win would turn comments off in such a mapping with no word, and reading it as on would let a real contradiction through. core-ts refuses it too. |

## 8. Phases A and B: implementation notes (2026-11-06)

**Built as section 3.** `_comments` is the only state, `_WORD_TEXT` is the WD's, and `_refuse_contradiction` runs in the constructor and in `of`, over a mapping and its keywords together. The warning is the package's first `warnings.warn`, and looks past the text model to the caller's line (`skip_file_prefixes`).

**One thing section 3.3 didn't say.** `include=["comments"]` stays an `include` of the caller's, as it was, as well as a `comments`. So `TextOptions(include=["comments"], profile="word-text")` holds no headers, since the call never asked for them, and a preset doesn't add them.

**A second `include` doesn't turn comments off.** `set("include", ["headers"])` after the old spelling leaves comments on, where it used to clear them. That's section 1's third row: `include` no longer speaks for comments.

**Verified.** `tests/text/test_text_options_comments.py`, 44: section 1's table, every form of the old spelling, the contradiction in six forms, sequential calls, presets, and `to_dict()`. `test_comments_false_under_the_preset_is_docx4js_reading` is replaced, the preset's table follows the WD and says so, and the tests that used the old spelling moved to `comments=True`. 17 mutants: 16 caught, one of those by a test added for it, and the 17th equivalent, since `_include` can no longer hold `comments`. No row of the suite, the agreement table or the goldens moved.

The whole suite: 11,628 passed, 8 skipped.
