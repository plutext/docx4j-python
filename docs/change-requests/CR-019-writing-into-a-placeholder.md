# CR-019: Writing into a content control showing its placeholder, and placeholders written in Word's form --- the port of docx4j-core-ts CR-002 section 63 (`020c039`)

**Status:** Proposed 2026-10-25. On 2026-10-25 the user decided that 0.2.0 waits for core-ts to implement its CR-002 section 63 and for the port here ("Wait for core-ts"; CR-016 section 13.5). core-ts implemented it on 2026-10-25 at `020c039` (`main`, unpushed, unreleased), and this CR is that port. Five questions, **decided 2026-10-25** (the user, section 8): Q1 (c), the bare form's styling held for core-ts's check 67, with core-ts's D2 (a) the fallback if no answer comes by the end of Phase A; Q2 to Q5 as recommended. **Phase A implemented 2026-10-25** (section 10), in two commits: the rules check 64 measured (`b17bbab`: Word's form, every route, D4, D7) and those resting on check 63 alone (`dc027fd`: the bare form's tracked deletion and `insert_paragraph`, a control's `"Content"` delete, D3 with risk 2's and the equation refusals gone, D6 (b)), held to Word's runs of checks 63 and 64 but for the bare form's look, settled plain by core-ts's check 67 (Q1, so no D2 (a)), and the controls Phase B makes. **Then check 63 was run again with Design Mode off** (section 11): its first run had been made in Design Mode, and with it off the bare form writes as Word's form does. The user decided core-ts's revised rules there (D1, D2 (b), D3's tracked markup, D5 and a new D10); **Phase A2 implemented 2026-10-25** (section 11.4, `dbce995`): it ports them, correcting `dc027fd`, so this package's run of check 63 gives run 2's answer in all 52 writes and core-ts's check 67 Word's in all 72; the bare form's look stays Phase A's, which differs from core-ts's `d84f3da` only for a placeholder run carrying more than its style, unmeasured: **decided 2026-10-25** (the user: "Measure; keep ours meanwhile"), core-ts asked for a check 68 (section 11.5). **Phase B implemented 2026-10-25** (section 12), the ports of core-ts `aefc41f` (`dfb05fa`: `placeholder_text` read from the glossary, `show_placeholder`, the binding pass and made controls in Word's form, `package.glossary_entries` and the save hook; checks 65 and 66 Word's in every write but 65b's case 9, check 64 now in all 52), `c275a2a` (`0dc354a`: check 68, decided there, a write takes the control's run properties and an entry is shown as the glossary holds it; check 68 Word's in all 24 writes, 4 before) and `b12e0c5` (`a3a9571`: each kind's default entry, from the user's Word 365 file). **Phase C's records done 2026-10-25** (section 13): the CHANGELOG with CR-016 Phase B's placeholder lines rewritten, `tests/README.md`, CR-016 section 13.5, CR-014 section 4, CR-003 section 25 and the index; and the Word check's files (Q5) built in `out/cr019-word-check/` for the coordinator to stage. **The Word check is not yet run.** Four phases (sections 6 and 11): A, the writes; A2, run 2's answers; B, placeholders in Word's form; C, the records and the Word check. 0.2.0 follows C, with a new wheel, `release_check.sh` and the wheel's Word check.

**Depends on:**
- **docx4j-core-ts** `020c039`: CR-002 sections 63 to 63.12; its `src/model/content/placeholders.mts`, `ContentControl.mts`, `Paragraph.mts` and `Range.mts`, and `src/model/customxml/bindings.mts`; `test/placeholders.test.mjs` (25 tests); Word checks 63 to 66, with their files, Word's JSONs and Word's saves (`test/fixtures/check63/` to `check66/`, `test/check63/` to `check66/`). Checks 63 to 66 were written and run at `2070694`, `ae9215f` and `2d080d4`. Check 67 and check 63's run 2 (Design Mode off), and the revision to run 2, at `2bafbcb` and `d84f3da` (section 11).
- **This repository:** CR-016 Phase B (the placeholder reads and `ContentControl.insert_text`), CR-013 Phase C (the span edits, `mark_holder_deleted`, the tracked Replace's order), CR-003 Phase E (the control kinds, `apply_bindings`), and the typed `GlossaryDocumentPart` (CR-002).
- **docx4j** CR-063 (proposed; its D7 decided, D1 to D6 open) is docx4j's counterpart for its binder. It is not needed; see Q3.
- **No fork change.** `codegen/generate.sh --check` is a gate only.

**Counterpart:** docx4j-core-ts `docs/change-requests/CR-002-content-api.md` section 63 at `020c039`.

## 1. Why

CR-016 Phase B ported core-ts's rule that a control showing its placeholder reads as no text, and made `ContentControl.insert_text` write in the placeholder's place. It left three questions for a Word measurement, and sent them to core-ts (CR-016 sections 11.6 item 1, 11.11 items 1, 2 and 2a):
- what Word does with every other write into such a control;
- whether a Replace over a whole run-level control keeps it (here, `paragraph.text = ...` refuses);
- what a bound control's node gets.

core-ts's checks 63 to 66 answered them (Word 365 16.0.20326.20232, the user's runs, 2026-10-24 and 25). The answers depend on the placeholder's **form**:
- **The bare form:** no `w:placeholder`, no glossary entry. This is what docx4j's binding writes, and what this package writes today. Word takes the placeholder's runs for the control's content (check 63).
- **Word's form:** `w:placeholder` names a glossary entry. Word treats the placeholder as no content (check 64).

core-ts decided D1 to D9 (the user, there, 2026-10-24 and 25) and implemented them at `020c039`. Its tests replay checks 63, 64 and 66 with Word's answers in every write, and check 65 in all but one case. This package's design follows core-ts's (CLAUDE.md), so the port is owed here, and 0.2.0 waits for it.

## 2. What Word does (core-ts sections 63.2, 63.5, 63.8 and 63.11, in brief)

- **Every text write that reaches the control's content replaces the placeholder and clears `w:showingPlcHdr`,** in both forms and at every location. The routes are the control's `insertText`, its `Content` range's `insertText`, and a block control's paragraph's `insertText`.
  - **In Word's form** the new text is plain. Tracked, only the insertion is recorded.
  - **In the bare form** the text takes the placeholder run's properties. In a rich text control the Placeholder Text style stays, with `w:color w:val="auto"` added for some writes (an `'End'`, a paragraph's `'Start'` untracked). In a bound plain-text control the style goes. Tracked, Word writes the `w:ins` first, then the placeholder in a `w:del`, and records the property change in `w:rPrChange`.
  - **The control's `insertParagraph('End')`:** in the bare form the text goes in the placeholder's place and the flag stays on. In Word's form the placeholder goes, its paragraph is left empty, and the new paragraph follows.
- **A Replace or delete over a whole control showing its placeholder takes it:** a paragraph's `insertText(…, 'Replace')` or `Content` delete over a run-level control, a run-level control's `Whole` Replace, and a paragraph's `Content` delete over a block control's only paragraph. Tracked, the control is marked deleted between `w:customXmlDelRange` markers, with its placeholder in a `w:del` inside, and a Replace's insertion comes first.
- **A run-level control's `Content` range is inside the control**, so its writes land there.
- **A paragraph inserted beside a block control's only paragraph goes outside the control.**
- **A bound control's node:** in the bare form, Office JS's writes never reached it, and Word lost them on reopening. In Word's form every write reached it. Text typed by hand reached it in both forms (check 65's case 10, check 66).
- **What Word needs of Word's form:** the reference alone gives Word's form while the file is open. But Word's save drops a reference whose entry no glossary holds, and an entry whose glossary styles lack Placeholder Text loses the style (check 65). A glossary document part plus a styles part holding Placeholder Text and Default Paragraph Font is enough: Word keeps every reference and completes the glossary on save (check 66, D9 (a)).

## 3. What exists here (at `73580ec`)

CR-016 section 13.5 has the probe; in short:
- **Only `ContentControl.insert_text` clears the flag.** It writes in the placeholder's place at every location and drops the `PlaceholderText` style (`controls.py`, `_drop_placeholder_style`). That matches Word's save in check 58's cases 20 to 22 (pinned in `tests/content/test_readers.py`), and not check 63's (Q1).
- **A block control's paragraph and its `Content` range write inside the control and leave the flag on**, so the text reads as nothing.
- **A run-level control's `Content` range with its placeholder showing is empty and collapses to the run before.** Its `"Start"`, `"End"` and `"Replace"` all write outside the control. With text showing, `"Start"` writes outside and the others inside; core-ts writes all three inside (Q4).
- **`paragraph.text = ...` over a run-level placeholder is refused** (`range.holds_placeholder`; the user's decision of 2026-10-24, made pending a measurement). The span path (`insert_text(…, "Replace")`, a `Whole` or `Content` delete) keeps the control and empties it.
- **A tracked write into a block placeholder** stays one paragraph, but the flag stays on.
- **Only a run-level control's `insert_text` writes its node** (`write_through`). A block control's `insert_text` goes through its body, and writes no node.
- **`insert_paragraph` beside a block control's only paragraph goes inside the control.**
- **A tracked `Paragraph.insert_text(…, "Replace")` writes the insertion first already** (CR-013 Phase C). It refuses where the paragraph holds an equation (`tracking.equation`, CR-013 section 13.7 D2, core-ts's section 62.2), a refusal core-ts dropped at `020c039`.
- **`placeholder_text`** reads the control's runs while the flag is on, else `""`. Its setter writes docx4j's run with the given text and refuses over content (`binding.holds_content`). Nothing reads or writes a glossary.
- **`apply_bindings`** writes docx4j's run for an empty value: "Click here to enter text.", in `PlaceholderText`, in the bare form (CR-016 section 13.4 item 1). `w:placeholder` is kept.
- **`insert_content_control` on an empty paragraph or at a point** makes an empty control, where Word's shows its default placeholder.
- **A tracked `insert_paragraph("After")`** marks the mark before the new paragraph, as Word does. core-ts's issue plutext/docx4j-core-ts#4 (its own marks the new paragraph's mark) is not shared (probed 2026-10-25).

## 4. Design here: core-ts's rules, by this package's writers

Each rule is core-ts's at `020c039`. Its names are given where this package's differ.

### 4.1 The form

- **`ContentControl.placeholder_entry`** (an extension; core-ts `placeholderEntry`): the entry `w:placeholder/w:docPart/@w:val` names. It is `""` for a `w:placeholder` naming none, and None without one.
- With any entry, present in the glossary or not, the control is in Word's form; with None, the bare form. Every rule below goes by this alone.

### 4.2 One write for every route (D1, D2, D5)

A text write that reaches a block- or run-level control showing its placeholder goes through one internal method, `ContentControl._write_placeholder(text, *, colour=False, keep_flag=False, breaks=...)` (core-ts `writePlaceholder`). The routes:

| route | `colour` |
|---|---|
| `ContentControl.insert_text` | `location != "Replace"` (the `'Start'` unmeasured, taken as `'End'`) |
| its `"Content"` range's `insert_text` (4.4) | no |
| a block control's paragraph's `insert_text` | `"End"`, or `"Start"` untracked (check 63's case 34, tracked, has none) |
| a block control's paragraph's `text = ...` and `splice` | no |

What it does:
- The placeholder's runs come out (a block control's other paragraphs go), and the text takes their place in one paragraph. The flag is cleared, unless `keep_flag` is set.
- **Word's form:** the text is plain. Under tracking only the insertion is recorded.
- **The bare form:** the text takes the first placeholder run's properties. In a bound plain-text control the `PlaceholderText` style is dropped; `colour` adds `w:color w:val="auto"`.
  - Under tracking: the `w:ins` first, then the placeholder's runs back in a `w:del`, after it and before any inner `w:customXmlInsRangeStart`. A change to the runs' properties is recorded in `w:rPrChange` (`ChangeTracker.record_r_pr_change`), on the inserted run and, in a bound plain-text control, on the deleted ones (check 63's cases 28, 35, 40 and 47 to 51).
  - **Q1** asks whether the style rule is check 63's (core-ts's D2 (a)) or check 58's.
- A run-level control's refusals (a paragraph break, `range.splits_control`) come first, before anything changes, as `insert_text`'s do today.
- The bound node is written (4.7).

**`ContentControl.insert_paragraph("Start" | "End")`** on a block control showing its placeholder:
- bare form: `_write_placeholder(text, keep_flag=True)`, returning the control's first paragraph (check 63's cases 3 and 29);
- Word's form: the placeholder's runs go, the flag is cleared, and the new paragraph follows the emptied one, as today (check 64's cases 6 and 27).

CR-016's `_drop_placeholder_style` and its rule that `insert_text` drops the style go into `_write_placeholder`'s forms.

### 4.3 A Replace or delete over the whole control (D3)

**`ContentControl._take_whole(text="")`** (core-ts `takeWhole`) takes a control showing its placeholder:
- **Untracked:** the control goes. A block control leaves its first paragraph, emptied, in its place. `text` goes where a run-level control stood.
- **Tracked:** it is marked deleted in Word's form (`tracking.mark_holder_deleted`: `w:customXmlDelRange` markers around it and inside its content), its placeholder runs in a `w:del` inside. A Replace's insertion comes first: before the control, or before the paragraph's deleted text.

It is reached from:
- **a paragraph's span edit holding a run-level placeholder control whole**: `Paragraph.splice` untracked, the tracked delete (`_delete_text`), by the span's edges (core-ts `placeholderControlsIn`; a placeholder is no text, so the control has no width);
- **a run-level control's `"Whole"` range**: `insert_text(…, "Replace")` or `delete()` (the delete unmeasured, taken as the Replace);
- **a paragraph's `"Content"` range's `delete()`** over a block control's only paragraph.

**Consequences:**
- **Risk 2's refusal goes.** `paragraph.text = ...` over a run-level placeholder takes the control, as Word's Replace does. The error code `range.holds_placeholder` was never released (it came in CR-016 Phase B, for 0.2.0), so it goes with it.
- **A tracked `Paragraph.insert_text(…, "Replace")` no longer refuses an equation.** It already goes through its `"Content"` range's Replace here, which writes the equation deleted in Word's form, as a range's Replace does (core-ts section 62.2's refusal went at `020c039`).
- **A run-level control holding text** keeps today's behaviour under a paragraph's Replace (emptied, kept). This is core-ts's unmeasured choice.

### 4.4 A control's own ranges (D4, core-ts's F2)

`ContentControl.get_range("Content" | "Whole")` on a block- or run-level control returns a `Range` that knows its control (core-ts `Range.ofControl`, internal). Its `insert_text` and `delete()` ask the control first:
- **`"Content"`, the control showing its placeholder:** `_write_placeholder`.
- **`"Content"`, an empty run-level control:** the text goes inside it (check 64's cases 21 and 42).
- **`"Content"` `delete()`, the control showing its placeholder:** the control and its flag stay. In the bare form its runs go, marked deleted under tracking (check 63's cases 6, 17, 32 and 43). In Word's form nothing changes (unmeasured).
- **`"Content"` `delete()` of a Word's-form control holding text, untracked:** the control is emptied, then shows its entry again (check 66's case 3; 4.8).
- **`"Whole"`, a run-level control showing its placeholder:** `_take_whole` (4.3).

Anything else is the range's own write.

**With text showing, a run-level control's `"Content"` range at `"Start"` writes inside the control** as core-ts does, where it writes into the run before today (CR-016 section 13.5). Word hasn't measured it (Q4).

The knowledge is on that `Range` object only. A range derived from it (`expand_to`, `get_range("Start")`) is an ordinary range, as in core-ts.

### 4.5 A paragraph beside a block control's only paragraph (D6 (b))

`Paragraph.insert_paragraph(…, "Before" | "After")`, on the only paragraph of a block control showing its placeholder, puts the new paragraph outside the control, beside it. The control stays at block level. Word saved it run-level (D6 (a)), but no reader here needs that.

### 4.6 One paragraph under tracking (D5, core-ts's F1 and F3)

`_write_placeholder` writes one paragraph tracked as untracked. The flag is cleared, so the tracked write reads as text. The deleted placeholder reads as nothing in the accepted view, as any deletion does.

### 4.7 A bound control's node (D7)

`Paragraph.splice` writes the node of every bound control around the span, after the edit and outside the tracked branch, as `write_through` does (CR-003 sections 4 and 16.6). The controls are those holding the paragraph, and the run-level ones holding text in the span (core-ts `boundControlsAt`). Every route then writes it, not only a run-level control's `insert_text`. In the bare form this departs from Word, which left the node empty and lost the write on reopening. That was decided in core-ts (D7 (a)).

### 4.8 Placeholders written in Word's form (D8)

- **`placeholder_text`** reads the text of the entry `w:placeholder` names, a body's paragraphs `\r` apart. It is None in the bare form, or where the glossary lacks the entry, as Office JS reads it (check 65). Its type becomes `str | None` (Q2).
- **Setting it** names `DefaultPlaceholder_-1854013440` for Word's default text, `"Click or tap here to enter text."`, or `""`. Any other text gets a new entry named with 32 upper-case hex digits (a new UUID's). A control showing its placeholder, or holding nothing, shows the text. A control holding content keeps it, and only its reference changes. The refusal `binding.holds_content` goes from this setter.
- **`ContentControl.show_placeholder()`** (an extension; core-ts `showPlaceholder`) shows the named entry, else Word's default, which the control is made to name. It writes the entry's text in the `PlaceholderText` style over the control's run properties (`_placeholder_run`), with the flag on.
- **`apply_bindings`:** an empty value shows the placeholder for a text control (rich or plain text), or any control naming an entry. Another kind naming none keeps docx4j's run (Q3).
- **`Paragraph.insert_content_control`** on an empty paragraph, and **`Range.insert_content_control`** at a point, make a rich text or plain text control showing the default entry, as Word's (check 64).
- **`w:showingPlcHdr`** is written without `w:val`, as Word writes it.
- **A Word's-form control emptied** by its `"Content"` range's `delete()`, untracked, shows its entry again (4.4).
- **A control naming an entry the glossary lacks**, given its placeholder by the binding pass or `show_placeholder()`, has that entry written with the default text, so the reference survives Word's save.

`bindings.PLACEHOLDER_TEXT` ("Click here to enter text.", docx4j's) stays for the kinds that keep docx4j's run. A new `DEFAULT_PLACEHOLDER_TEXT` and `DEFAULT_PLACEHOLDER_ENTRY` go with the glossary code.

### 4.9 The glossary (D9 (a))

`package.glossary_entries` (an extension; core-ts `pkg.glossaryEntries`), in a new hand-written `model/content/placeholders.py` (added to `codegen/clean.py`'s `KEEP`):
- **Reading:** the glossary document part's entries, read privately (`read_contents`, CR-002 section 16), so an untouched glossary is saved byte for byte. Read again only where the part has changed, and from the live tree once the part is unmarshalled.
- **Writing:** entries named since are kept pending, and the save writes them. The save:
  - makes the glossary document part where the file has none (`GlossaryDocumentPart`, its relationship from the main document part, its content type);
  - adds each pending entry in Word's form (check 65): `w:docPartPr` with the name, category General / gallery `placeholder`, type `bbPlcHdr`, behaviour `content`, and a `w:guid` of its own; and a body of one paragraph, one run in `PlaceholderText`;
  - makes the glossary's styles part (`/word/glossary/styles.xml`), or completes one lacking them, with Default Paragraph Font and Placeholder Text as Word defines them. The definitions are core-ts's `test/fixtures/check66/build.mjs` `PLACEHOLDER_STYLES`, and `parts/wml/placeholderStyle.mts` for Placeholder Text. A styles part holding both stays untouched.
- **Where the save writes them:** a hook the save runs before writing anything. It is engine-neutral: `save_package` calls each callable a package has registered, and the engine imports no model class (`tests/openpackaging/test_threads_and_import.py`). core-ts also writes at save (`flush`); writing at once would put a trial's entries into the real package.
- **A trial** (`dry_run`) has its own pending entries, thrown away with it. `trial.save()` is refused already.
- **`last_save`** lists the parts a save added. The pending entries are cleared once written.

### 4.10 Unmeasured choices, as core-ts made them

- The control's `insert_text(…, "Start")` in the bare form adds the colour, as `"End"` does.
- A run-level control's `"Whole"` `delete()` takes it, as its Replace does.
- In Word's form, a `"Content"` `delete()` of a control showing its placeholder changes nothing.
- D3 is for controls showing their placeholder. A paragraph's Replace over a run-level control holding text keeps the control, emptied, as before.
- `placeholder_text = ""` names the default entry.
- `Body.insert_content_control` and a range across paragraphs make controls as before, with no placeholder.

### 4.11 Still differing from Word after the port

- check 65's 65b, case 9 (Word's glossary had lost the style);
- D6 (b) and D7's bare form, as decided in core-ts;
- Word merges runs where this package leaves two (check 63's case 18), and orders a made control's `w:sdtPr` children differently;
- whatever Q1 settles for the bare form's style.

## 5. Tests and oracles

- **Word's runs, ported.** As CR-013 and CR-016 did for checks 52 to 59 (`tests/office_js/checks.py`), port checks 63 to 66's Script Lab snippets step for step to the content API: `run_check_63` to `run_check_66`, in their report and run modes. Then hold this package's run to Word's JSON case by case.
  - The files are copied into `tests/office_js/` with their SHA-1s in `SHA1SUMS`: each check's `.docx`, Word's report and run JSONs, and Word's saves.
  - They are Apache-2.0 here, as checks 58's and 59's were (CR-016 Q2).
  - **The target is core-ts's:** Word's answer in every write of checks 63, 64 and 66, and in check 65 all but 65b's case 9 (and the bare form's style, by Q1). Each remaining difference is pinned, with its reason.
- **Word's saves.** The markup of the regions Word's saves hold (`63-run-word365.docx` and the rest) is compared with this package's save after the same run. Rsids, paraIds, `w:sdtEndPr` and `w:id` values are ignored, as CR-016's check 58 comparison does.
- **core-ts's 25 tests** (`test/placeholders.test.mjs`), ported where they are not Word-run replays: D3's markup accepted and rejected (`accept_all`, `reject_all`); D6's level; D7's nodes in both forms; D8's setter, binding and made controls; and D9's glossary made, added to, and left untouched, including a save of a file with a glossary whose bytes must not change.
- **Pre-existing expectations that move,** each listed in the implementation notes:
  - CR-016 Phase B's placeholder tests in `test_readers.py`: the style dropped (by Q1), `placeholder_text` from the runs, risk 2's refusal;
  - check 58's cases 20 to 22 markup pin (by Q1);
  - `test_controls_typed.py`'s `binding.holds_content`;
  - CR-013's tracked equation Replace refusal;
  - the binding tests that expect docx4j's run for an empty text control;
  - any acceptance artefact that makes a control on an empty paragraph.
- **Gates:**
  - the full suite;
  - `codegen/generate.sh --check` in a scratch worktree;
  - `roundtrip.py` both ways (`--normalise` and `--no-normalise`), 50 of 50;
  - `compare_acceptance.py` against a copy taken first, every moved artefact explained;
  - `scripts/compare_word_resolution.py --no-records`, 0 lines.

## 6. Phases

| phase | what | size |
|---|---|---|
| A | The writes: 4.1 to 4.7 and 4.10 (the form, `_write_placeholder` by every route, `insert_paragraph`, `_take_whole`, the control's ranges, D6, D7; risk 2's refusal and the equation refusal gone); checks 63 and 64 ported and held to Word's runs | 1½ to 2 days |
| B | Word's form written: 4.8 and 4.9 (`placeholder_text`, `show_placeholder`, the binding pass, `insert_content_control`, `package.glossary_entries` and the save hook); checks 65 and 66 ported and held to Word's runs | 1 to 1½ days |
| C | The records (this CR's notes, CR-016 section 13.5's pointer, CR-014 section 4, CR-003's member notes, `CHANGELOG.md` with CR-016 Phase B's placeholder lines rewritten, `CLAUDE.md`, `tests/README.md`), the Word check (Q5) | ½ day |

The phases go in order (A2 between A and B, section 11.3): B's emptied-control rule (4.4) needs A's ranges. Each phase is committed when its gates pass. After C, 0.2.0's own steps: the CHANGELOG date, `release_check.sh`, and the wheel's Word check.

## 7. Risks

1. **Two of Word's runs disagree on the bare form's style** (Q1). Check 58's Word save holds plain runs where check 63's keeps `PlaceholderText`, for the same write on the same markup, in the same Word build. core-ts implemented check 63's answer. Told to core-ts 2026-10-25, which confirmed it on its own copies: the style definition and `settings.xml` are byte-identical, and neither file tracks revisions. Four things differ, none of them measured: check 63 sets `changeTrackingMode = 'Off'` before each write; it finds the control through `getContentControls()` where check 58 uses `getByTag()`; check 63's file has a custom XML part and check 58's a numbering part; and check 58's fills ran after Accept and Reject in other regions, with All Markup showing. core-ts is proposing a check 67 to the user to separate them, and stays on D2 (a) until it runs.
2. **`placeholder_text` changes meaning** (Q2). A caller reading `""` for "no placeholder", or reading docx4j's prompt from a bound control, reads None. 0.1.0 shipped the old reading.
3. **A file gains a glossary.** A control made on an empty paragraph, a placeholder set, or an empty bound value now writes a glossary document part and a styles part on save. An untouched glossary is saved byte for byte, and a file that writes no placeholder is unchanged.
4. **The bare form's node write departs from Word** (D7, decided in core-ts). It keeps the write through Word's reopen, which Word's own write loses.
5. **The binding pass moves ahead of docx4j** (Q3) until docx4j's CR-063 lands.

## 8. Questions (decided 2026-10-25)

The user, 2026-10-25: Q1, "Hold for check 67" (option (c), as recommended); Q2 to Q5, "All as recommended".

- **Q1. The bare form's style.** For the bare form's text, take:
  - (a) check 63's answer, core-ts's D2 (a): the style kept in a rich text control, `w:color auto` where Word added it, check 58's markup pin moved and recorded as a Word disagreement;
  - (b) check 58's: the text plain, as CR-016 Phase B writes it, a departure from core-ts;
  - (c) hold the bare form's styling for core-ts's answer to the contradiction, and port the rest meanwhile.

  *Recommended: (c),* with (a) as the fallback if no answer comes by the end of Phase A. core-ts's check 67 (risk 1) would give that answer. The rest of the port doesn't depend on it, the styling is one method, and Word's form, which this package writes after B, is plain either way.
- **Q2. `placeholder_text` reads the glossary entry, None where there is none** (core-ts's D8, Word's reading), and its setter names an entry and no longer refuses over content. *Recommended: yes,* as core-ts, recorded in the CHANGELOG under "Changed".
- **Q3. The binding pass writes Word's form now:** an empty value in a text control shows the default entry, where docx4j's binder still writes the bare form (its CR-063 proposes the same change; its D1 to D6 are open). *Recommended: yes,* as core-ts. It answers CR-016 section 13.4 item 1 and CR-014 section 4's note, and the departure from docx4j is recorded until CR-063 lands.
- **Q4. A run-level control's `"Content"` range at `"Start"`, with text showing, writes inside the control,** as core-ts does (D4), where this package writes into the run before. Word hasn't measured it. *Recommended: yes.*
- **Q5. The Word check (Phase C).** Stage:
  - (1) this package's saves after its runs of checks 63 to 66, each to open without a repair prompt, with each region's text as Word's run left it;
  - (2) one file written here with no glossary: a placeholder set, an empty bound value, a control made on an empty paragraph, a bare-form bound control. The user opens it, saves it with nothing typed, reopens that save, types into the two bound controls, and saves again. Pass: no repair prompt; every reference kept and every placeholder grey through the saves; the typed values in their nodes (check 66's steps, on this writer's glossary).

  *Recommended: both.* The 0.2.0 wheel check follows as planned.

## 9. Effort

3 to 4 days in all: A 1½ to 2, B 1 to 1½, C ½. The Word check is about half an hour of the user's time.

## 10. Phase A implementation notes (2026-10-25)

Phase A as built, on CR-019 decided (`5696e9b`). The oracle is docx4j-core-ts `020c039` (its CR-002 sections 63 to 63.12; `ContentControl.mts`, `Paragraph.mts`, `Range.mts`, `placeholders.mts` and `customxml/bindings.mts`; `test/placeholders.test.mjs`; checks 63 and 64's files, snippets, Word's JSONs and saves), read with `git show 020c039:<path>` and never from core-ts's working tree; core-ts's `HEAD`, `4f3f993`, changes only its CR-002 Status line. Two things changed the brief while the phase ran (the coordinator, 2026-10-25):
- **Q1 is settled: the bare form's text is plain.** core-ts's check 67 (Word 365 16.0.20326.20232) gave all 72 bare-form writes a plain run, as check 58 has it: no Placeholder Text, no `w:color auto`, the flag cleared, across three files, both views, tracking untouched or set Off, and both lookups. So core-ts's D2 (a) was not written (no code for it exists), and check 63's style and colour cases are Word disagreements, "check 63 contradicted by checks 58 and 67", not held cases.
- **Check 63 itself is suspect.** Both of Word's saves of its file carry `<w:formsDesign/>` in `settings.xml`, which suggests Word ran it in Design Mode; core-ts proposes running it again. So the rules resting on check 63 alone are a commit of their own, after the rules check 64 measured, which can be held or backed out cleanly (10.4).

**Commits.**
- `b17bbab`: the rules check 64 measured (Word's form, every route, D4, D7), checks 63 and 64's files and port, `test_placeholder_writes.py`.
- `dc027fd`: the rules resting on check 63 alone (10.4), and the tests they move.
- the records, the commit after `dc027fd`: this section and the Status line.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the form | `controls.py`, `ContentControl.placeholder_entry` (an extension) | `w:placeholder/w:docPart/@w:val`; `""` for a `w:placeholder` naming none; None without one (section 4.1) |
| one write for every route | `controls.py`, `_write_placeholder`, `_take_placeholder_runs`, `_dress_as_placeholder`, `_insert_into_empty`, `_first_paragraph`, `_shown_text`, `shows_placeholder_content` | section 4.2; core-ts's `writePlaceholder`, `takePlaceholderRuns`, `dressAsPlaceholder`, `insertIntoEmpty`, `firstParagraph`, `showsPlaceholderContent` |
| the routes | `ContentControl.insert_text`; `Paragraph.insert_text`, `text = ...` and `splice` (`_placeholder_control`, core-ts's `placeholderControl`); `Range.insert_text` through a control's own range | section 4.2's table, less its `colour` column (10.5 item 1) |
| `insert_paragraph` | `ContentControl.insert_paragraph("Start" \| "End")` | Word's form in `b17bbab`, the bare form in `dc027fd` |
| a control's own ranges (D4) | `Range.of_control` (a slot, set by `ContentControl.get_range("Content" \| "Whole")`); `ContentControl._write_through`, `_delete_through` (`dc027fd`); `_run_place` | section 4.4; core-ts's `Range.ofControl`, `writeThrough`, `deleteThrough`, `runSpan`'s `empty` |
| D3 | `ContentControl._take_whole`, `_mark_whole_deleted`, `_put_back_deleted`; `Paragraph._placeholder_controls_in`, in `_splice_text` (untracked) and `_delete_text` (tracked); `Range.delete` | section 4.3; core-ts's `takeWhole`, `markWholeDeleted`, `putBackDeleted`, `placeholderControlsIn` (`dc027fd`) |
| D6 (b) | `Paragraph.insert_paragraph`, `controls.holds_one_paragraph` | section 4.5 (`dc027fd`) |
| D7 | `Paragraph.splice` and `_bound_controls_at` (core-ts's `boundControlsAt`), `_holds_bound_run_control`; `controls.control_view_of`; a block control's `insert_text` | section 4.7 |
| the refusals gone | `Paragraph._refuse_placeholder_controls` (`range.holds_placeholder`) removed; `Paragraph.insert_text`'s tracked equation refusal removed | section 4.3 (`dc027fd`) |
| checks 63 and 64 | `tests/office_js/` (ten files, `SHA1SUMS`), `tests/office_js/checks.py` (`run_check_63`, `run_check_64`, `parse_region_63`, `parse_region_64`, `node_of_63`) | 10.3 |
| the tests | `tests/content/test_placeholder_writes.py` (48, new); `test_readers.py`, `test_fields_equations.py`, `test_office_js_characters.py` | 10.3 and 10.6 |
| the subset | `tests/office_js_subset.json`, regenerated | core-ts's committed `HEAD` gained a note on `placeholderText` at `020c039` (`string \| null`, Q2, Phase B's), and `test_office_js_subset.py` has failed on it in this checkout since; only the note moved (`b17bbab`) |

No file under `docx4j_py/` is new, so `codegen/clean.py`'s `KEEP` is unchanged.

### 10.2 The rules as built

1. **The form** goes by `placeholder_entry` alone: an entry, named in the glossary or not, is Word's form; None the bare form.
2. **A text write into a block- or run-level control showing its placeholder** (`shows_placeholder_content`) goes through `_write_placeholder`, by every route: `ContentControl.insert_text` at any location; the control's `"Content"` range's `insert_text`; and a block control's paragraph's `insert_text` (any location), `text = ...` (one paragraph, `breaks="line"`) and `splice` (any text, as core-ts's `splice`). It:
   - refuses a paragraph break into a run-level control first (`range.splits_control`), before anything changes;
   - records the prompt as the report's `text_before` (CR-016 section 11.6 item 6);
   - takes the placeholder's runs out (a run-level control's content's runs; each block paragraph's direct runs, every paragraph after the first going), clears `w:showingPlcHdr`, and writes the text into the first paragraph (`Paragraph.splice(0, 0, ...)`) or the run-level control's content (`_insert_into_empty`: a new run, in a `w:ins` under tracking, before an inner `w:customXmlInsRangeStart`);
   - **Word's form:** plain text; the runs go untracked, so under tracking only the insertion is recorded (check 64);
   - **the bare form:** the text takes the first placeholder run's run properties, its `w:rPrChange` left out and `PlaceholderText` dropped (CR-016 Phase B's `_drop_placeholder_style`; checks 58 and 67's plain run). Under tracking the placeholder's runs go back marked deleted after the insertion (`dc027fd`; before it, the insertion alone, as in Word's form);
   - writes a bound control's node (`write_through`).
3. **`ContentControl.insert_paragraph("Start" | "End")`** on a block control showing its placeholder: in Word's form the runs go, the flag is cleared, and the paragraph goes after the emptied one; in the bare form (`dc027fd`) `_write_placeholder(text, keep_flag=True)`, returning the control's first paragraph.
4. **A control's own ranges** (`get_range("Content" | "Whole")` of a block or run control) carry `of_control`; `expand_to`, `get_range` and every other range made from one are ordinary. `Range.insert_text` asks `_write_through` first: the `"Content"` range of a control showing its placeholder is `_write_placeholder` (no other rule); of an empty run-level control, the text goes inside it at any location; of a run-level control holding text at `"Start"` or `"Before"`, inside it before its text (Q4: untracked into its first run, `_extend`; tracked a `w:ins` inside before it, `splice(..., prefer="after")`); a `"Whole"` Replace of a run-level control showing its placeholder is `_take_whole(text)` (`dc027fd`). `Range.delete` asks `_delete_through` (`dc027fd`): a `"Content"` delete of a control showing its placeholder keeps it and the flag, its runs going in the bare form (marked deleted tracked), nothing in Word's form; a `"Whole"` delete of a run-level one takes it.
5. **An empty run-level control stands where the walk passes it** (`_run_place`'s visit, core-ts's `runSpan`), after the text before it, where `_run_span` gave offset 0. So its `get_range()` is at its place, and writes through it land inside it.
6. **D3** (`dc027fd`): `_take_whole` takes a control showing its placeholder. Untracked it goes, a block control's first paragraph, emptied, taking its place; text goes where a run-level one stood. Tracked it is marked deleted (`mark_holder_deleted`: markers around it and inside its content), its placeholder runs in a `w:del` after the inner `w:customXmlDelRangeEnd`, the flag left on, a Replace's text inserted first. It is reached from a paragraph's span edit holding a run-level placeholder control whole (`_placeholder_controls_in`, by the span's edges: untracked in `_splice_text` before the edit, tracked in `_delete_text`), so `paragraph.text = ...`, `insert_text(..., "Replace")` and a `"Content"` or `"Whole"` range's delete or Replace take it; from a run-level control's `"Whole"` range; and from a paragraph's `"Content"` range's delete over the only paragraph of a block control showing its placeholder. A run-level control holding text keeps today's behaviour under a paragraph's Replace (section 4.10).
7. **D6 (b)** (`dc027fd`): `Paragraph.insert_paragraph("Before" | "After")` on the only paragraph of a block control showing its placeholder (`holds_one_paragraph`) inserts beside the control, through the control's body; the control stays block-level.
8. **D7**: `Paragraph.splice` collects the controls bound by `w:dataBinding` around the span before the write (those holding the paragraph, and, where the paragraph holds a bound run-level control, those holding a segment touching the span) and calls `write_through()` on each after it, outside the tracked branch. `write_control_to_node` skips a control still showing its placeholder and a node already holding the value. A block control's `insert_text` writes its node too, as core-ts's `insertText` does.

### 10.3 The oracles

**Checks 63 and 64**, run through the port (`run_check_63` and `run_check_64`, the snippets at `020c039` step for step) and held to Word's JSONs case by case through core-ts's `digest` and `view` (the region's control, flag, style, where `NEW<n>` landed, the paragraphs; the control's type, text and paragraphs; what the write returned; `placeholderText` and `showingPlaceholderText` left out, as core-ts leaves them):

| | before (`5696e9b`'s code) | `b17bbab` | `dc027fd` | core-ts at `020c039` |
|---|---|---|---|---|
| check 63's 52 writes agreeing | 8 (6, 21, 24, 32, 38, 45, 46, 52; core-ts's 8) | 16 | **23** | 52 |
| check 64's 52 writes agreeing | 20 | **45** | 45 | 52 |
| the report modes, every region | Word's (both) | Word's | Word's | Word's |
| check 63's 54 regions' saved markup Word's | 6 | 9 | **20** | --- |
| check 64's 52 regions' saved markup Word's | 13 | **38** | 38 | --- |
| check 64's ten nodes Word's (`NEW<n>`) | 2 (49, 51) | 10 | 10 | 10 |

The digest port is core-ts's: Word's JSONs give core-ts's `WORD_RUN` and `WORD_RUN_64` exactly. `63-placeholders.docx` and `64-placeholders-word.docx` are core-ts's git-ignored copies, written by `build.mjs` on 2026-10-24 before Word's run (21:16 and 21:57; the share is not reachable here); each reads paragraph for paragraph and control for control as Word's save of it as opened.

**Pinned at `dc027fd`, each with its reason:**
- **Check 63's run, 29 writes, all Q1's**, the digest's style letter alone (every other read Word's): 1 to 5, 7 to 9 and 13 to 16 (Word's `S`, the text in Placeholder Text; here plain, `-`); 27 to 31, 33 to 35 and 39 to 42 (tracked: here the deleted placeholder keeps its style, `s`); 47 to 51 (tracked, bound plain-text: Word took the style off the deleted runs too, recorded in a `w:rPrChange`, D2 (a)'s; here they keep it, `s`). "Check 63 contradicted by checks 58 and 67 (CR-019 Q1)".
- **Check 64's run, 7 writes, Phase B's** (D8): the controls `insert_content_control` makes (19 to 21, 36, 40 to 42) show no placeholder here. The made-run ones are also made in the wrong place, a pre-existing bug (10.8 item 1). The other 7 made cases (15 to 18, 37 to 39) agree already.
- **Check 63's saves, 34 regions:**
  - **Q1, 29:** the 29 writes above. Each is Word's markup once check 63's `PlaceholderText`, `w:color auto` and `w:rPrChange` come off the text written, and in 47 to 51 the deleted runs' `w:rPrChange` is read as the style it recorded. Case 42 also has Word's `w:lastRenderedPageBreak`.
  - **D6 (b), 4 (10, 11, 36, 37):** Word's save made the control run-level in its paragraph; here it stays block-level, as decided in core-ts. 10, 11 and 37 are otherwise Word's. In 36 (tracked `"After"`) the inserted mark is the new paragraph's, where Word's is the mark before it, its control's paragraph's.
  - **Run boundaries, 1 (18):** Word's save leaves `NEW18` a run of its own, where this package's text joined the run before; core-ts's own test pins this package's form (`PLAIN18A NEW18` and ` PLAIN18B`).
- **Check 64's saves, 14 regions, Phase B's:** the controls made (15 to 21, 36 to 42). Every region with a control from the file is Word's, Word's form tracked as an insertion alone included.
- **Check 63's nodes: D7 (a)'s departure.** 21 to 25 and 47 to 51 hold `NEW<n>`, where Word's write left them empty and lost on reopening. 26 and 52, whose control the write took, stay empty, as Word's.

**core-ts's tests that are no replay** (`placeholders.test.mjs` at `020c039`), their strings kept: D3's markup (18 to 20, 12, 44 to 46, 38, and the bare form's tracked 28 and 47, those two in this package's plain look), Accept All (45, 44) and Reject All (45 restored, showing its placeholder); D6's level (10, 11, 36, 37); D7's nodes in both forms. D8's and D9's tests are Phase B's.

**This package's own:** every route into a block placeholder, both forms, tracked or not, saved and read back (one paragraph, flag cleared, the look, the tracked markup); Word's form's `insert_paragraph("End")` (check 64 case 6); the bare form's `insert_paragraph` (case 3); a control's own `"Content"` delete in both forms; a run-level `"Whole"` delete tracked and not; an empty run-level control's `"Content"` write, and where it stands; Q4, tracked and not; a range made from a control's own range is ordinary; D7 by every route, block and run.

### 10.4 The rules resting on check 63 alone

All in `dc027fd`, which reverts cleanly to `b17bbab`'s rules:
1. the bare form's `insert_paragraph("Start" | "End")` writing in the placeholder's place and leaving the flag on (case 3, 29);
2. the bare form's tracked write: the `w:ins`, then the placeholder's runs in a `w:del` (cases 27 to 51);
3. a control's own `"Content"` range's delete: the control and flag kept, the bare form's runs going (deleted tracked), nothing in Word's form (cases 6, 17, 32, 43);
4. D3 (`_take_whole`): a Replace or delete over a whole control showing its placeholder takes it, tracked in Word's markers (cases 12, 18 to 20, 26, 38, 44 to 46, 52); with it, risk 2's refusal (`range.holds_placeholder`) and the tracked equation-Replace refusal (`tracking.equation`) go;
5. D6 (b): a paragraph beside a block control's only paragraph goes outside it (cases 10, 11, 36, 37).

Backed out, `b17bbab`'s rules apply: the bare form's tracked write records the insertion alone, as Word's form does (unmeasured for the bare form); its `insert_paragraph` adds the paragraph inside the control, the placeholder kept and the flag on, as 0.1.0 does; a control's `"Content"` delete does the range's own delete (a run-level one's collapsed range deletes nothing); a Replace over a run-level placeholder control keeps it, emptied, and `paragraph.text = ...` over one is refused, as CR-016 Phase B left them; a paragraph beside a block control's only paragraph goes inside it. The coordinator's list named items 1, 2, 4 and 5; item 3 rests on check 63 alone too, since check 64 has no delete.

### 10.5 Departures and choices, with why

1. **Section 4.2's `colour` column is not written**, nor core-ts's D2 (a) look in any form: Q1, settled by check 67 (above). `_write_placeholder` has no `colour` parameter. In a bound plain-text control the tracked deleted runs keep their style, where check 63's Word took it off them in a `w:rPrChange`: that is D2 (a) too.
2. **Without `dc027fd` the bare form's tracked write records the insertion alone** (10.4), the cleanest state to back out to.
3. **`holds_one_paragraph` counts blocks** (a paragraph, table, control, custom XML, `w:altChunk`, kept `mc:AlternateContent`), so a bookmark beside the paragraph doesn't stop D6 (b). core-ts's `holdsOneParagraph` counts every `BLOCK_LEVEL_TYPES` item, its markers among them. Unmeasured either way.
4. **Q4's write**, untracked, extends the control's first run (`_extend`, as `ContentControl.insert_text("Start")` writes); tracked, a `w:ins` inside the control before it. core-ts's goes through its `splice` with `'after'`.
5. **D6 (b) under tracking** marks the new paragraph's own mark inserted (`insert_element` beside the control). Word's save, its control run-level, marks the mark before the new paragraph. With the control at block level that mark is inside the control, and marking it would join across the control's edge on Reject. Recorded (case 36), not changed.
6. **An empty run-level control's offset** is now where it stands (10.2 item 5), where 0.1.0 and CR-016 gave 0: needed for D4, and core-ts's.
7. **A through-write returns its range without moving the control's own range object**, as core-ts's `writeThrough` returns early.
8. **The two checks' files** are core-ts's git-ignored copies, checked against Word's saves as opened (10.3).

### 10.6 Pre-existing tests whose expectation changed

- **`b17bbab`:** `tests/content/test_office_js_characters.py::test_every_fixture_is_the_manifests_and_the_files_word_read` counts 43 files in `SHA1SUMS` (33). `tests/office_js_subset.json` is regenerated (10.1), which `test_office_js_subset.py::test_the_script_would_write_what_is_committed` failed on.
- **`dc027fd`:**
  - `test_readers.py`: `test_risk_2_paragraph_text_written_over_a_run_level_placeholder_is_refused_tracked_or_not_nothing_changed` becomes `..._takes_the_control_as_words_replace_does`: untracked the control goes with the content, tracked it is marked deleted in Word's form. Its second half (a control written into first) is unchanged.
  - `test_fields_equations.py` (core-ts's section 62.2 D2 test, as core-ts moved it at `020c039`): the tracked `Paragraph.insert_text(..., "Replace")` over an equation goes ahead, its insertion first and the equation deleted in Word's form, where it was refused.
  - `test_office_js_characters.py::test_an_equation_reads_as_word_draws_it_...`: the same write over `equation_package()` goes ahead the same way.
- **Unmoved:** check 58's cases 20 to 22 markup pin (`test_readers.py`; now Word's twice over, checks 58 and 67), CR-016's placeholder tests in `test_readers.py` (every location, the refused write, the bound control, `apply_bindings`' placeholder, the text model), `test_controls_typed.py`'s `binding.holds_content` and the binding tests (Phase B's).

### 10.7 The gates

| gate | before (`5696e9b`) | `b17bbab` | `dc027fd` |
|---|---|---|---|
| the full suite | 6,335 passed, 8 skipped (the brief's); in this checkout `test_office_js_subset.py`'s check against core-ts's `HEAD` fails since core-ts's `020c039` (10.1) | 6,377 passed, 8 skipped (4 min 36 s) | **6,383 passed, 8 skipped** (4 min 32 s) |
| `roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata`, `--normalise` and `--no-normalise` | --- | 50 of 50 both ways, 0 differences, 34 `boolean-spelling`, 0 kept, 0 normalised, 0 skipped | the same |
| acceptance (`compare_acceptance.py` against `out/acceptance-before-cr019a/`, copied first) | a regeneration at `5696e9b` (in a scratch worktree) equal to the copy: 17 artefacts, 243 entries, 0 differing beyond the dates | 17 artefacts, 243 entries, 0 differing | the same: no artefact writes into a placeholder |
| `scripts/compare_word_resolution.py --no-records` | --- | 0 lines either side | 0 lines either side |
| `codegen/generate.sh --check` (a scratch worktree of the commit, `TMPDIR` in the scratchpad) | --- | "two generations of docx4j_py are byte identical", 307 modules; the tree left as committed | the same |
| checks 63 and 64 | 10.3 | 10.3 | 10.3 |

### 10.8 The cost

The median of 7 passes after a warm one, `5696e9b` (a scratch worktree) and `dc027fd` interleaved, three rounds (ms):

| | before | after |
|---|---|---|
| 600 `insert_text`s into 300 plain paragraphs | 41.1 to 41.9 | 43.2 to 44.0 |
| 280 reads of a run-level control's `text` (check 64's 14, twenty times) | 10.9 to 11.0 | 11.8 to 12.2 |
| 200 `splice`s in a bound run-level control | 15.0 to 15.2 | 37.8 to 38.4 |

About 3.5 µs a write where no control is near (`_placeholder_control`'s parent test and `_bound_controls_at`'s climb), 3.5 µs a run-level control's span (`_run_place`'s visit), and 115 µs a write into a bound control: D7's node write, an XPath to the node and a comparison, on every splice.

### 10.9 Found, and not done

1. **`Range.insert_content_control()` on a collapsed range wraps the run its point is in.** `insert_content_control_in_range` takes `span.runs`, and a collapsed range's runs are the run around the point. So check 64's made-run controls hold `PLAIN<n>A ` (19 to 21, 40 to 42). Phase B rewrites this path (section 4.8, a control made at a point showing the default entry); the fix belongs there: an empty control at the point, the run split there (`insert_items_at`).
2. **A tracked `ContentControl.insert_text("Start" | "End")` on a run-level control holding text writes the text untracked.** `_extend` sets the run's text directly. This is pre-existing, not a placeholder matter, and not changed; Q4's tracked route avoids it.
3. **Check 63's rerun** (Design Mode, above) may move 10.4's rules: they are `dc027fd`'s.
4. **`CLAUDE.md`'s CR-016 line** still says `paragraph.text = ...` over a run-level placeholder is refused, `range.holds_placeholder`, and `CHANGELOG.md`'s 0.2.0 entry says so too (line 298). Both are the coordinator's and Phase C's.
5. **Phase B's**, untouched: `placeholder_text`'s reading and setter, `show_placeholder`, the binding pass, `insert_content_control` showing a placeholder, `w:showingPlcHdr` without `w:val`, the glossary, and a Word's-form control emptied showing its entry again.

## 11. Check 63 run again, with Design Mode off (2026-10-25)

### 11.1 What core-ts found (its CR-002 sections 63.13 to 63.17)

- **Check 67** (core-ts `2bafbcb`; Word 365 16.0.20326.20232, the user's run) gave check 58's three bare-form writes in three
  files, four phases and both lookups. All 72 came out a plain run with the flag cleared, as check 58's did.
- **The cause was Design Mode.** Only check 63's two saves carry `<w:formsDesign/>`, and the user confirmed that Developer >
  Design Mode was on for the whole run, to see the controls. In Design Mode a control's placeholder is editable text. core-ts's
  `test/README.md` now says to keep Design Mode off while a check writes into controls.
- **Check 63's run 2, with Design Mode off** (core-ts section 63.16; `test/fixtures/check63/63-placeholders-word365-run2-*.json`,
  `test/check63/63-run2-*-word365.docx`): **the bare form gives Word's form's answers, case for case.**
  - A write into the control is plain, with the flag cleared, and no `w:color auto` (D2 (b)).
  - Tracked, only the insertion is recorded (D5).
  - The control's `insertParagraph` clears the flag; the placeholder's paragraph is left empty and the new one follows (D1).
  - Every bound node holds its write, and the controls show the writes on reopening (D7 is Word's behaviour, not a departure).
  - **The bare form's own:** a control's `Content` delete keeps the control and its flag over a run of five spaces, nothing marked
    tracked (D10; Word's filler for an emptied control, check 52's cases 13 and 26). A Replace or delete over the whole control
    takes it, as in run 1 (D3); tracked, it stands between Word's `w:customXmlDelRange` markers with nothing inside marked deleted
    and the flag gone. A paragraph beside a block control's only paragraph goes outside it, as in run 1 (D6).
  - **Design Mode refused to turn off** over check 63's control 53: a run-level control showing its placeholder over an empty
    `w:sdtContent`, which this package's `Content` delete leaves (section 11.2). D10's filler stops that.
- **Decided 2026-10-25** (the user, in core-ts: "section 63 decisions as recommended, including (b) for D2"): D1, D2 (b), D3's tracked
  markup, D5 and D10 (a) as Word; D7 needs no change; D4, D6 (b), D8 and D9 stand. core-ts implemented them as its section 63.17
  at `d84f3da` (`main`, unpushed), Phase A2's oracle.

### 11.2 What it means here

The user had decided, before the rerun, to hold the rules resting on check 63 alone until it came, and then apply its answers
("Hold them for the rerun"). The decision reached Phase A after `dc027fd` was committed. It is not reverted: run 2 confirms two of
its five rules, and Phase A2 corrects the other three.

| `dc027fd`'s rule | run 2 | Phase A2 |
|---|---|---|
| the bare form's tracked `w:ins` then `w:del` of the placeholder | only the insertion (D5) | the `w:del` goes: one rule for both forms |
| the bare form's `insert_paragraph("Start"/"End")` in the placeholder's place, the flag left on | as Word's form (D1) | `b17bbab`'s Word's-form rule for both forms |
| a control's `"Content"` delete: the bare form's runs out (deleted under tracking), the flag on, an empty `w:sdtContent` | five spaces under the flag, nothing marked (D10 (a)) | core-ts's `showFiller`, and `showEmptied` for a control holding text emptied untracked |
| D3, `_take_whole` | confirmed; tracked: nothing inside marked deleted, the flag cleared | the tracked markup changed |
| D6 (b) | confirmed | none |

Q1 is settled as Phase A writes it: plain, by checks 58, 67 and run 2. CR-016 section 13.5's "Start" question (Q4) and D4, D7, D8
and D9 are unchanged. Checked by hand at `b435a28`: the bare form's tracked write is the `w:ins` and the `w:del` (as `dc027fd`
wrote it), and a run-level control's `"Content"` delete leaves `<w:sdtContent/>` with the flag on, control 53's state.

### 11.3 Phase A2

- **The port** of core-ts's section 63.17 at `d84f3da`, onto `dc027fd`'s code: the table above. Pre-existing pins move, among them
  check 52's case 13 (the filler; Word's is bold, from `w:sdtPr/w:rPr` taken from the deleted text, which core-ts doesn't follow).
- **The oracle:** run 2's two JSONs and two saves, copied into `tests/office_js/` with their SHA-1s; this package's run of check 63
  held to run 2 in every write, as core-ts's is. Run 1's files stay, as the Design Mode record. Check 67's files are copied and held
  to as well.
- **The gates** as Phase A's.
- **Size:** ½ to 1 day.

### 11.4 Phase A2 implementation notes (2026-10-25)

Phase A2 as built, on `67db0d5` (the coordinator's `4616c68`, the commit before Phase A2's, touches nothing of placeholders, so every "before" figure is `67db0d5`'s code). The oracle is docx4j-core-ts `d84f3da` (its CR-002 sections 63.13 to 63.17; the diff to `ContentControl.mts`, `Paragraph.mts`, `Range.mts`, `placeholders.mts` and `customxml/bindings.mts`; `test/placeholders.test.mjs`, `ranges.test.mjs` and `readers.test.mjs`), and check 67 from `2bafbcb`, read from core-ts's commits and, for the git-ignored files, its working tree, which was left untouched.

**Commits.**
- `dbce995`: the port (one commit: `dc027fd`'s three corrected rules and D10 are one change of `_write_placeholder`'s callers and the control's own delete), check 63's run 2 and check 67 copied and held to, `run_check_67`, and the tests they move.
- the records, the commit after `dbce995`: this section and the Status line.

#### 11.4.1 What landed

| piece | where | what |
|---|---|---|
| one write for both forms (D1, D2 (b), D5) | `controls.py`, `_write_placeholder` (its `keep_flag` gone), `_dress_as_placeholder` (its tracked half gone), `_put_back_deleted` removed | the placeholder's runs go untracked in either form, so under tracking only the insertion is recorded; the bare form's look is Phase A's (11.4.4 item 1) |
| `insert_paragraph("Start" \| "End")` (D1) | `ContentControl.insert_paragraph` | `b17bbab`'s Word's-form rule for both forms: the placeholder's paragraph emptied, the flag cleared, the new paragraph after it |
| D3's tracked markup | `ContentControl._mark_whole_deleted` | the control between Word's `w:customXmlDelRange` markers, its placeholder's runs taken out unmarked, `w:showingPlcHdr` removed (core-ts `markWholeDeleted`) |
| D10 (a), Word's filler | `controls.py`, `FILLER`, `_show_filler`, `_show_emptied`; `_delete_through`; `range.py`, `Range.delete` and `_delete_own` | core-ts `showFiller` and `showEmptied`: a `"Content"` delete of a bare-form control showing its placeholder shows five spaces in one run in `w:sdtPr/w:rPr`'s properties under the flag, tracked or not, nothing marked; an untracked `"Content"` delete that leaves a bare-form control holding no text does the same; Word's form's emptied control is Phase B's seam (`_show_emptied`'s early return, where B calls `show_placeholder`) |
| D7's wording | `test_placeholder_writes.py` | the bare form's node write is Word's behaviour now (check 63's run 2), not a departure; the code's docstrings never called it one, and `_write_placeholder`'s now says so |
| comments | `insert_text`, `placeholder_entry`, `_take_whole`, `Range.delete`, `Paragraph._delete_text` | run 2's cases and `d84f3da`'s rules, where they cited run 1 |
| check 63's run 2 and check 67 | `tests/office_js/` (sixteen files, `SHA1SUMS`): `63-placeholders-word365-run2-report.json`, `-run2-run.json`, `63-run2-opened-word365.docx`, `63-run2-word365.docx`; `67a-plain.docx`, `67b-custom-xml.docx`, `67c-revisions.docx`, their three `-word365.json`, and Word's six saves `67{a,b,c}-{opened,run}-word365.docx` | Apache-2.0 here, as checks 58 to 64's; every SHA-1 checked against core-ts's commits (the JSONs and saves) and, for Word's saves, against its `test/check63/` and `test/check67/` manifests |
| check 67's snippet | `tests/office_js/checks.py`: `run_check_67`, `parse_region_67`, `PHASES_67`, `_control_of_67`, `_apply_67`, `_report_67` | `check67-script-lab.js` at `2bafbcb` (unchanged at `d84f3da`) step for step: its report and its four phases, each write alone, change tracking set off around each write in an `-off` phase and never touched in an `-untouched` one, the control found by `getByTag` or `getContentControls()` as the case's lookup says |
| the tests | `tests/content/test_placeholder_writes.py` | 11.4.3 and 11.4.5 |

No file under `docx4j_py/` is new, so `codegen/clean.py`'s `KEEP` is unchanged.

#### 11.4.2 The rules as built, against `d84f3da`'s

1. **`writePlaceholder`:** one rule, as core-ts's: the runs out untracked, the flag cleared, the text into the first paragraph (`Paragraph.splice(0, 0, ...)`) or the empty run-level control (`_insert_into_empty`), the node written. Only the bare form's look differs, in a case Word hasn't measured (11.4.4 item 1).
2. **`insertParagraph('Start' | 'End')`** on a block control showing its placeholder: core-ts's, in both forms.
3. **`markWholeDeleted`:** core-ts's: `mark_holder_deleted` with `_take_placeholder_runs` inside, then the flag off. Rejected, the control comes back empty without the flag, as core-ts's test has it (unmeasured).
4. **`showFiller`:** core-ts's: the content's runs out (`_take_placeholder_runs`, a block control's paragraphs after the first going), one run of `FILLER` with a deep copy of `w:sdtPr/w:rPr` where there is one, last in what held them but before an inner `w:customXmlInsRangeStart` or `DelRangeStart` (core-ts `pushBeforeMarker`; `_TRACKED_HOLDER_STARTS`), then `w:showingPlcHdr` on (`put_property`, in `SDT_PR_ORDER`'s place).
5. **`deleteThrough('Content')`:** core-ts's: a bare-form control showing its placeholder shows the filler, tracked or not; Word's form changes nothing. `'Whole'` unchanged (D3).
6. **`showEmptied`:** core-ts's condition (not showing its placeholder, a block or run control, its `text` `""`); the bare form's filler; Word's form left as the delete left it, for Phase B (core-ts calls `showPlaceholder`, `020c039`'s D8, which Phase A didn't port).
7. **`Range.delete` on a control's own `"Content"` range** the control didn't take: core-ts's `emptied` (untracked, read from the range's first paragraph before the delete), the range's own delete (`_delete_own`: `_delete_span` across paragraphs or with a mark, else the splice), then `_show_emptied`. The two are one `ChangeReport` (an outer `recording`, the inner ones adding to it).
8. **`Paragraph.insertText`:** core-ts dropped its `colour` option; this package never wrote it (Phase A's 10.5 item 1).
9. **`placeholders.mts` and `bindings.mts`:** comments only at `d84f3da`, on code that is Phase B's here.

Everything else of Phase A stands as section 10 has it: D4, Q4, D6 (b), risk 2's and the equation refusals gone, and every rule of `020c039` that `d84f3da` left alone.

#### 11.4.3 The oracles and the tallies

Run through the port and held to Word's JSONs through core-ts's `digest` (now `d84f3da`'s, which finds `FILL<n>` as well as `NEW<n>`) and `view`:

| | before (`67db0d5`'s code) | after |
|---|---|---|
| check 63's 52 writes agreeing with run 2 | 27 | **52** (core-ts: 52) |
| check 63's 52 writes agreeing with run 1, in Design Mode | 23 | 21, differing in exactly the 31 writes in which Word's two runs differ (core-ts's 31) |
| check 64's 52 writes | 45 | 45 (Phase B's seven, unchanged) |
| check 67's 72 writes (and its three report steps) | 72 (all) | **72** (all; core-ts's 72) |
| the report modes (check 63's two runs, check 64) | every region Word's | every region Word's |
| check 63's 54 regions' saved markup, Word's run-2 save | 22 | **49** |
| check 63's 54 regions, Word's run-1 save (the record) | 20 | 11 |
| check 64's 52 regions' saved markup | 38 | 38 |
| check 67's 72 controls' saved markup, Word's three run saves | 72 | 72 |
| check 63's twelve bound nodes, run 2's | 12 | 12 (D7 had written them since `b17bbab`) |

Phase A already wrote check 67's plain run, so check 67 moves nothing; it is held now, and it pins the look in all 72 writes.

**Pinned after Phase A2, each with its reason:**
- **check 63's writes:** none.
- **check 64's writes, 7:** the controls `insert_content_control` makes (19 to 21, 36, 40 to 42), Phase B's (D8), as section 10.3 has them.
- **check 63's saves against run 2's, 5:** D6 (b) (10, 11 and 37, otherwise Word's once Word's run-level control is put back at block level; 36, D6 (b) and the inserted mark the new paragraph's, section 10.5 item 5), and case 18's run boundary (Word's save leaves `NEW18` a run of its own). Every other region is Word's: D3's markers (12, 19, 20, 26, 38, 44 to 46, 52), D10's filler (6, 17, 32, 43), the emptied paragraph of `insert_paragraph` (3, 29) and every tracked insertion.
- **check 64's saves, 14:** the controls made, Phase B's, unchanged.
- **check 67:** none. Its title paragraph in 67c holds a tracked deletion, which Word's All Markup steps read into `paragraph.text` (`... Inserted. Deleted.`); that is Word's display view, which this package's text doesn't take, and the test compares the title as Simple Markup reads it.
- **check 52's case 13** (`test_ranges.py`): this package's filler is a plain run of five spaces, where Word's is bold: Word also gave that control a `w:sdtPr/w:rPr` taken from the text it deleted, which isn't followed, as core-ts doesn't (its `ranges.test.mjs`).

**core-ts's tests, ported** (`placeholders.test.mjs` at `d84f3da`): run 2's two JSONs of this file's cases and `w:formsDesign` in run 1's saves alone; Word's answers with Design Mode off (the plain writes, the insertion alone, `insert_paragraph`'s emptied paragraph, the filler, the taken controls without their flag, the nodes in the JSON and in Word's save); the port's run held to run 2 in every write, and its 31 differences from run 1; D3's markup (the markers with nothing inside, 38's empty paragraph, 28 and 47 the insertion alone, D10's 6, 17, 32 and 43), Accept All and Reject All (45 back empty, its flag off); check 67's files and snippet agreeing, Word's JSONs and saves as copied, Word's answers (every write a plain run, the flag cleared, in the JSONs and the saves), and the port's run held to Word's in every step.

**This package's own:** every route into a block placeholder, both forms, tracked or not (no `w:del` in either); `insert_paragraph("End")` in both forms, tracked or not; D10's filler for a block and a run control, tracked or not, read as no text by `text`, `Body.text` and `search`, read by the text model as the placeholder run it is (left out with `TextOptions(placeholders=False)`), and saved and read back; Word's form's `"Content"` delete changing nothing; a control holding text emptied by its `"Content"` range showing the filler untracked and nothing tracked, block and run, and Word's form's left empty for Phase B; a run-level `"Whole"` delete tracked with nothing inside marked; check 67's saved controls against Word's saves.

#### 11.4.4 Departures and choices, with why

1. **The bare form's look stays Phase A's, where `d84f3da` writes plain in both forms.** `_dress_as_placeholder` gives the text the first placeholder run's own properties with the `PlaceholderText` style dropped (CR-016 Phase B's rule, pinned by `test_readers.py`'s `test_every_location_writes_in_the_placeholders_place_the_flag_cleared_the_style_gone_the_runs_own_look_kept` and Word-checked in CR-016's file 2, region 7, bold); core-ts's `writePlaceholder` at `d84f3da` writes the text with none of the placeholder run's properties in either form. The two agree on every placeholder Word measured (checks 58, 63, 64 and 67: each placeholder run has the style alone, so both write a run with no properties) and differ only for a placeholder run carrying more, such as `w:b`, which no check has measured. The brief ("the plain look stays as it is") and section 11.2 ("Q1 is settled as Phase A writes it") keep it, so it is not changed here; it is the one rule of `d84f3da` not mapped as core-ts has it, and wants a decision: keep it as a recorded departure, or take core-ts's (delete `_dress_as_placeholder` and its call; that test's run-level cases move).
2. **Word's form's emptied control** is left as the delete left it (`_show_emptied`), where core-ts shows its entry (`showPlaceholder`, D8): Phase B's, as before Phase A2.
3. **The filler and the readers.** Every Office JS reader reads it as no text while the flag is on (CR-016 Phase B's rule, already in place for check 52's case 26 accepted); the text model reads it as it reads any placeholder's prompt (REQ-508's placeholder runs, on by default), five spaces, and leaves it out with `placeholders=False`. No reader changed.
4. **One `ChangeReport` for a delete and its filler** (11.4.2 item 7), where a nested write would otherwise report twice.
5. **Check 67's three files** are core-ts's git-ignored copies, written by `build.mjs` (unchanged since `2bafbcb`) at 08:00 on 2026-10-25, before Word's run at 08:16; each reads paragraph for paragraph and control for control as Word's save of it as opened, as checks 63's and 64's files do (section 10.5 item 8).
6. **Check 67's view** is core-ts's: the digest, the control's type and text (not its paragraphs) and what the write returned; and the title (above).
7. **One code commit**, where Phase A made two: Phase A2's rules are all run 2's and decided together, so there is nothing to hold apart.

#### 11.4.5 Pre-existing tests whose expectation changed

- **`test_placeholder_writes.py`** (Phase A's own):
  - `DIFFERS_63`'s 29 Q1 pins are gone: the run is held to run 2 in every write (`test_check_63s_run_is_word_365s_run_2_in_every_write`, which replaces `..._but_for_the_differences_pinned`); run 1 is kept as a record (`test_check_63s_run_differs_from_run_1_in_design_modes_31_writes`, `WORD_RUN_63`), with `WORD_RUN2_63` beside it.
  - `SAVES_63`, from 34 pins against run 1's save to 5 against run 2's (`test_check_63s_saved_markup_is_words_save_after_run_2_but_for_d6_b_and_a_run_boundary`); `_plain`, `_deleted_look` and `_no_layout`, the forms check 63's run 1 needed, are gone.
  - every route into a block placeholder: tracked, no `w:del` in the bare form (it held the placeholder's runs deleted after the insertion).
  - `test_the_bare_forms_insert_paragraph_at_end_writes_in_the_placeholders_place_and_leaves_the_flag_on_check_63_case_3` and `test_words_form_insert_paragraph_...` become one test over both forms.
  - `test_a_controls_own_content_range_deleted_keeps_it_and_its_flag_its_placeholder_going_in_the_bare_form` becomes the D10 tests: the filler, nothing marked tracked (it held the content empty, and tracked the placeholder's runs deleted).
  - `test_a_run_level_controls_whole_range_deleted_takes_it_as_its_replace_does`: tracked, nothing inside the markers and no flag (it held a `w:del` inside).
  - the D3 test: `_MARKED` without the flag or the `w:del`, region 38 an empty paragraph inside the markers, 28 and 47 the insertion alone, D10's four regions, and Reject All's control back without its flag (it came back showing its placeholder).
  - the D7 test: check 63's nodes are Word's run 2's, no departure.
- **`test_readers.py`:** `test_risk_2_paragraph_text_written_over_a_run_level_placeholder_takes_the_control_as_words_replace_does`, tracked: the control between the markers with nothing inside and no flag (it held the placeholder deleted inside).
- **`test_ranges.py`:** check 52's case 13 in the markup test: `P[ListParagraph/num|] "     "` (it was an empty paragraph), Word's still `{b}`; the comment on case 13 in the phase-2 run test (its `paragraphs` still differ: Word's filler reads bold).
- **`test_office_js_characters.py`:** `SHA1SUMS` counts 59 files (43).
- **Unmoved:** CR-016's look test above (11.4.4 item 1), check 58's cases 20 to 22 markup pin (now Word's in checks 58, 63's run 2 and 67), the accept-all test of check 52 (13 and 26 still Word's bold filler), and every binding test (Phase B's).

#### 11.4.6 The gates

| gate | before (`4616c68`) | `dbce995` |
|---|---|---|
| the full suite | 6,388 passed, 8 skipped (this commit's run less the 13 tests it adds; a clean worktree of `4616c68` skips 639, lacking the sibling checkouts) | **6,401 passed, 8 skipped** (4 min 16 s), twice; the 8 warnings are `test_effective.py`'s converter warnings, as before |
| `roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata`, `--normalise` and `--no-normalise` | --- | 50 of 50 both ways, 0 differences, 34 `boolean-spelling`, 0 kept, 0 normalised, 0 skipped |
| acceptance (`compare_acceptance.py out/acceptance-before-cr019a out/acceptance`) | Phase A left the copy equal to `HEAD` | 17 artefacts, 243 entries, 0 differing beyond the dates: no artefact writes into or deletes a placeholder |
| `scripts/compare_word_resolution.py --no-records` | --- | 0 lines either side (2 compared, 0 differ) |
| `codegen/generate.sh --check` (a scratch worktree of `dbce995`, `TMPDIR` in the scratchpad) | --- | "two generations of docx4j_py are byte identical", 307 modules; the tree left as committed |
| `ruff check`, and `ruff format --line-length 120` on the lines changed | --- | clean; `test_placeholder_writes.py` formatted whole (it was clean at 120); the other files' format differences are those `HEAD` had |

#### 11.4.7 Found, and not done

1. **The look question** (11.4.4 item 1), for the user.
2. **Word's bold filler** in check 52's case 13 (Word's `w:sdtPr/w:rPr` taken from the deleted text): not followed, as core-ts doesn't.
3. **Word's form's emptied control showing its entry:** Phase B's (section 4.8), at `_show_emptied`'s seam.
4. **Section 10.9's items 1 and 2** stand: a collapsed range's `insert_content_control` wraps the run around its point (Phase B's), and a tracked `ContentControl.insert_text("Start" | "End")` on a run control holding text writes untracked.
5. **`CLAUDE.md`'s and `CHANGELOG.md`'s placeholder lines** (section 10.9 item 4) are the coordinator's and Phase C's; the CHANGELOG's should now say that a `"Content"` delete of a bare-form control writes Word's filler.

### 11.5 After Phase A2: the bare form's look measured; Phase B's oracle (2026-10-25)

- **The bare form's look (section 11.4.4 item 1). Decided 2026-10-25** (the user: "Measure; keep ours meanwhile"). This package
  keeps the first placeholder run's own properties but `PlaceholderText`; core-ts's `d84f3da` writes the text plain. They differ only
  for a placeholder run carrying more than its style. That is what docx4j's binder writes under a control's `w:sdtPr/w:rPr`, such as
  bold. core-ts is asked for a check 68 (Design Mode off, untracked, `insertText` `'Replace'` and `'End'`):
  - (a) a bare-form block and run-level control whose placeholder run is `PlaceholderText` and `w:b`, with no `w:sdtPr/w:rPr`;
  - (b) the same with `w:sdtPr/w:rPr` `w:b`;
  - (c) a Word's-form control whose `w:sdtPr/w:rPr` has `w:b`.

  Both ports follow Word's answer before Phase C's Word check. Until then the rule stays as Phase A wrote it.
- **Phase B's oracle moves to core-ts's section 63.18**, committed at `aefc41f` ("A placeholder shows its entry's
  content; a missing entry gives way to Word's default"). It aligns core-ts with docx4j's CR-063 phases 1 and 2 (docx4j `5f5e53921`,
  `VERSION_17_3_5`, unpushed), which docx4j-25 reported the same day:
  - a control showing its placeholder holds a copy of its entry's content, the control's own `w:sdtPr/w:rPr` over each run's
    properties (a block control every paragraph, a bound one the first, a run-level one the first paragraph's runs);
  - a text control naming an entry the glossary lacks is made to name `DefaultPlaceholder_-1854013440`, nothing written under the
    missing name; another kind keeps its name.

  Section 4.8's "a control naming an entry the glossary lacks ... has that entry written with the default text" is replaced by the
  second. Q3's departure from docx4j shrinks to what core-ts records there: entries written on save (docx4j writes them during
  binding), a new `w:guid` (docx4j gives check 34's), and no configurable placeholder run. Phase B starts from `aefc41f`.
- **Check 68:** the user said yes (2026-10-25), with cases (a) to (c) and core-ts's (d): a Word's-form control whose `w:sdtPr/w:rPr` is bold, its content deleted, to see whether the entry Word brings back carries the control's bold (section 63.18's run-property precedence).

## 12. Phase B implementation notes (2026-10-25)

Phase B as built, on `caff8cf`. The oracle is docx4j-core-ts `aefc41f` (its CR-002 sections 63.7 to 63.12 and 63.18; `placeholders.mts`, `ContentControl.mts`, `Paragraph.mts`, `Range.mts`, `customxml/bindings.mts`, `WordprocessingMLPackage.mts`, `parts/wml/index.mts`; `test/placeholders.test.mjs`, `customxml.test.mjs`, `readers.test.mjs` and `ranges.test.mjs`; checks 65 and 66's snippets, Word's JSONs and saves), read with `git show <commit>:<path>`; the git-ignored files (each check's `.docx`) are core-ts's working copies, each written by its `build.mjs` before Word's run. docx4j's counterpart, CR-063 phases 1 and 2 (docx4j `5f5e53921`), was read for the binder's rules; where the two differ, core-ts is followed (section 11.5). Two more core-ts commits came while the phase ran, each the coordinator's instruction to port as a commit of its own:
- **`c275a2a`** (its sections 63.19 to 63.21): check 68, run by the user, and its two decisions ("decisions accepted"), D11 and section 63.18 revised (12.9). It supersedes section 11.5's "keep ours meanwhile": the bare form's look is check 68's now.
- **`b12e0c5`** (its section 63.22): Word's default placeholder entry for each kind of control, docx4j's CR-063 D8, decided there by the user ("D8 we go (a)") (12.10).

**Commits.**
- `dfb05fa`: the port of `aefc41f` (D8, D9, Q2, Q3), checks 65 and 66 copied and ported, the save hook, `test_placeholders_words_form.py`.
- `0dc354a`: the port of `c275a2a` (D11, section 63.18 revised), check 68 copied and ported; with it, the glossary's names and `w:guid`s drawn from the package's seeded generator (12.5 item 2).
- `a3a9571`: the port of `b12e0c5` (each kind's default entry), the user's `placeholders-word365.docx` copied.
- the records, the commit after `a3a9571`: this section and the Status line.

### 12.1 What landed

| piece | where | what |
|---|---|---|
| the glossary's entries | `docx4j_py/model/content/placeholders.py` (new, hand-written; `codegen/clean.py`'s `KEEP` names it beside `"model"`, which held it already) | `GlossaryEntries` (`load`, `loaded`, `text_of`, `body_of`, `has`, `require`, `waiting`, `pending`, `copy_to`, `flush`), `glossary_entries_of`, `load_glossary_entries`, `entries_of`, `entry_content_of`, `doc_part_of`, `default_paragraph_font`, `placeholder_style`, `has_placeholder_styles`, `ensure_glossary_styles`, `require_placeholder_style`, `placeholder_entry_of`, `name_placeholder_entry`, `new_entry_name`, `shows_placeholder_when_made`; since `a3a9571` `DefaultEntry`, `DEFAULT_ENTRIES`, `default_entry_for`. core-ts's `placeholders.mts` |
| `package.glossary_entries` | `model/content/__init__.py` (`_PACKAGE_PROPERTIES`); the slot `WordprocessingMLPackage._glossary_entries`; `TrialPackage.glossary_entries` and its slot | an extension, core-ts `pkg.glossaryEntries`; a trial's its own |
| the read when a body is first read | `model/content/body.py`, `body_of` | `load_glossary_entries`, core-ts `ensureGlossaryEntries` in `getBody`; again only where the glossary part is another |
| the save hook | `openpackaging/packages/opc_package.py` (`OpcPackage.save_hooks`), `openpackaging/save.py` (`run_save_hooks`, run by `save_package` before anything is written), `openpackaging/normalise.py` (`SaveReport.added`) | engine-neutral: a list of callables, each given the package and returning the names of the parts it added; the content API registers `GlossaryEntries.flush` there; the engine imports nothing from the model (`tests/openpackaging/test_threads_and_import.py` passes unchanged) |
| the control | `model/content/controls.py`: `placeholder_text` (getter and setter), `show_placeholder`, `shows_words_form_placeholder`, `_show_emptied`, `_glossary`, `_glossary_part_name`, `set_bound_content` (the main document's Placeholder Text, 12.5 item 3); `is_text_control` (`dfb05fa`), replaced by `default_placeholder_entry` (`a3a9571`); `_take_control_properties` (`0dc354a`), where `_dress_as_placeholder` and `_drop_placeholder_style` went | core-ts's `placeholderText`, `showPlaceholder`, `showsWordsFormPlaceholder`, `showEmptied`, `defaultPlaceholderEntry`, `writePlaceholder`'s D11 step |
| the binding pass | `model/customxml/bindings.py`, `apply_binding_to` | an empty value shows the placeholder where `shows_words_form_placeholder`; `PLACEHOLDER_TEXT` Word 365's since `a3a9571` |
| made controls | `model/customxml/insert.py`: `insert_content_control_in_paragraph` (an empty paragraph, `segments(objects=True)` empty), `insert_content_control_in_range` (a point), `_insert_at_point` | a control showing its kind's default entry; a collapsed range's control made at its point (section 10.9 item 1) |
| checks 65, 66 and 68 | `tests/office_js/` (31 files, `SHA1SUMS`, 90 entries; Apache-2.0 here, as checks 58 to 67's), `tests/office_js/checks.py` (`run_check_65`, `run_check_66`, `parse_region_65`, `run_check_68`, `parse_region_68`) | the snippets at `aefc41f` and `c275a2a` step for step; checks 65 and 66 differ only in their names |
| Word's default entries | `tests/office_js/placeholders-word365.docx` | the user's Word 365 file, one control of each kind (core-ts's `test/placeholder-defaults/`, SHA-1 `9cd6fb70…`) |
| the tests | `tests/content/test_placeholders_words_form.py` (new, 38); `test_placeholder_writes.py`, `test_readers.py`, `test_controls_typed.py`, `test_ranges.py`, `test_office_js_characters.py`; `tests/openpackaging/test_private_readers.py` (the two new private readers and a reader reaching them), `test_parts_and_relationships.py` (the save hook) | 12.3 and 12.6 |

### 12.2 The rules as built, against `aefc41f`'s

1. **`placeholder_text`, read:** `w:placeholder`'s entry's text from the glossary, its paragraphs `\r` apart (`text_of` of each); None with no `w:placeholder` or where neither the glossary nor this session holds the entry. Its type is `str | None` (Q2).
2. **`placeholder_text`, written:** the kind's default text (Word's default entry's for a text control and for a kind with none, `a3a9571`), or `""`, names the kind's default entry; any other text a new entry, 32 upper-case hex digits, drawn again while the glossary holds the name. The entry is required (`require`: pending where neither glossary nor session holds it). A control showing its placeholder, or whose text is `""`, shows one run of the text in `PlaceholderText` (the control's `w:sdtPr/w:rPr` under it at `dfb05fa`, not since `0dc354a`) with `w:showingPlcHdr`; one holding content keeps it. `binding.holds_content` is gone from the setter.
3. **`show_placeholder()`** (an extension): a control naming no entry, or one the glossary lacks, is made to name its kind's default (at `dfb05fa` a text control's, or any control's naming none; since `a3a9571` `default_placeholder_entry`), required; nothing is written under a missing name. With the entry's body, a copy of each paragraph's run content but `w:proofErr` and the bookmark ends (`entry_content_of`); a block control takes every paragraph, a bound one (`is_bound_block_control`) the first, the further paragraphs a copy of the control's first paragraph's `w:pPr`; a run-level control the first paragraph's runs. With no body, the kind's default text in one `PlaceholderText` run; a kind with no default, docx4j's run (`PLACEHOLDER_TEXT` in `PlaceholderText` over the control's `w:sdtPr/w:rPr`), with no reference (`a3a9571`). Then `set_bound_content(..., True)`.
4. **`w:showingPlcHdr`** is written bare: `el.showingPlcHdr()` already was; nothing writes `w:val` (D8's line needed no code).
5. **`apply_bindings`' empty value:** `show_placeholder()` where `shows_words_form_placeholder` (a kind with a default entry, or one naming an entry the glossary holds); any other kind docx4j's run, as before.
6. **Made controls:** `insert_content_control` around a paragraph holding nothing (`segments(objects=True)` empty), or on a collapsed range, makes a control showing its kind's default (rich text and plain text at `dfb05fa`; date picker, drop-down list and combo box too since `a3a9571`). The collapsed range's control is an empty `w:sdt` at its point, the run there split (`_insert_at_point`: `insert_items_at`'s refusals and placing, but after a run ending at the point, past whatever carries no text, so two controls made at one point keep their order). `Body.insert_content_control` and a range across paragraphs are unchanged (section 4.10).
7. **A Word's-form control emptied** by its `"Content"` range's untracked delete: `_show_emptied` shows its entry (`show_placeholder`) where it names one and its kind has a default or the glossary holds the entry; else Word's filler (Phase A2's).
8. **The glossary, read:** `load()` reads the glossary document part privately (`read_contents`), a part that will not parse holding no entry; again only where the main document part's `glossary_document_part` is another object; from the live tree once the part is unmarshalled. `body_of`, which makes the main part's body and the headers', footers' and notes', calls it, as core-ts's `getBody` does; and every query reads it first where it hasn't been read (12.5 item 1).
9. **The glossary, written** (`flush`, the save's hook): nothing while no entry waits. Else the glossary document part is made where the file has none (`/word/glossary/document.xml`, its relationship from the main document part and its content-type override, through `add_target_part`); each waiting entry the glossary lacks is appended in Word's form (`doc_part_of`: `w:name`, category General, gallery `placeholder`, type `bbPlcHdr`, behaviour `content`, a `w:guid` of its own; a body of one paragraph, one run of the text in `PlaceholderText`, `xml:space` where the text needs it); the waiting entries are cleared; then `ensure_glossary_styles`: a styles part made where the glossary has none (`/word/glossary/styles.xml`, related from it), or completed, with Default Paragraph Font and Placeholder Text as check 66's `PLACEHOLDER_STYLES` defines them (Placeholder Text's `w:color 666666`, `placeholderStyle.mts`'s); one already holding both is read privately and left byte for byte. `last_save.added` lists the parts made.
10. **A trial** has its own `GlossaryEntries`, made with the real package's waiting entries copied in (`copy_to`), never registered as a hook (a trial isn't saved), and dropped with it. There is no package copy or clone here to give `copy_to` to.

### 12.3 The oracles and the tallies

Run through the ports and held to Word's JSONs through core-ts's `digest` and `view`, and `placeholderText` beside them (Phase A's `view` leaves it out, as core-ts's does; Phase B reads it as Word does, so it is held too):

| | at `caff8cf` | after |
|---|---|---|
| check 63's 52 writes, run 2 | 52 | 52 |
| check 64's 52 writes | 45 (19 to 21, 36, 40 to 42 the made controls) | **52**; `placeholderText` Word's in every control |
| check 65a's 10 writes (report: 10 of 10 both) | 9 (9, the made control showing nothing); `placeholderText` differing in 9 controls of the report, 9 of the run | **10**; `placeholderText` Word's in every control |
| check 65b's 14 writes (report: 14 of 14 both) | 13 (9 showing nothing); `placeholderText` differing in 3 and 10 controls | 13 (9, below; core-ts's target); `placeholderText` Word's in every control |
| check 66a's and 66b's 6 writes each (report: 6 of 6) | 4 (3, the emptied control, and 4, the made one, showing nothing); `placeholderText` differing in 2 and 3 controls | **6** each; `placeholderText` Word's |
| check 67's 72 writes | 72 | 72 |
| check 68's 24 writes | 4 (9 to 12; at `dfb05fa` 4 too) | **24** (core-ts's 24) |
| check 63's 54 regions' saved markup, run 2's save | 49 | 49 |
| check 64's 52 regions' saved markup | 38 (14 made-control regions differing) | 38, and the 14 Word's but for Word's `w:id` and a run boundary |
| check 65a's saved controls, Word's run save | 0 of 10 | 2 of 10 (12.4) |
| check 65b's | 9 of 14 | 10 of 14 (12.4) |
| check 66a's and 66b's | 3 of 6 each | **6 of 6** each |
| check 67's 72 saved controls | 72 | 72 |
| check 68's 24 saved controls | 4 (at `dfb05fa` 4) | **24** |
| Word's default entry per kind, `placeholders-word365.docx` (10 controls) | --- | every kind's entry, text and this package's default Word's |

Checks 65, 66 and 68's files read as Word's report did: every region, control and `placeholderText` (since `dfb05fa`).

### 12.4 Pinned, each with its reason

- **Check 65b's case 9**, the run (`9 CPS - `, Word's `9 CP- - `): Word showed the control it made unstyled, its reading of 65b's glossary, whose styles lack Placeholder Text, having dropped the style; this package keeps a style a glossary doesn't define, and its own glossaries define it (core-ts's 63.18; core-ts pins the same).
- **Check 64's saves, 14 regions** (15 to 21, 36 to 42): Word gives a control it makes a random `w:id` where this package's is the next free one, and in 19 to 21 and 40 to 42 Word's save joins the run the point split off (`" "`) with the run after it (`"  PLAIN19B"`); each is Word's once both are taken out (`_no_control_id`, `_runs_joined`). The rest of each region, the default entry named, the flag cleared, the plain write and the tracked insertion, is Word's.
- **Check 65's saves:**
  - 65a's 1 to 5, 7 and 8: Word's session dropped every reference to an entry no glossary held when it saved the file as opened (65a has no glossary), before the run; this package's run writes the default entry (D9), so the references stay. 65b's 12 and 13 likewise, the missing entry's reference kept, with nothing written under the name (section 63.18);
  - 65a's and 65b's 10: typed into by hand after the run (`TYPED10`), which the snippet doesn't do;
  - 65b's 9: as the run's.
- **Check 63, 66, 67 and 68:** none beyond Phase A2's (section 11.4.3).

**core-ts's tests that are no replay, ported:** D8's setter, binding and made controls (`test_d8_…`, with the made-at-a-point control's place); D9's glossary made (parts, relationships, content types, each entry's markup and its own `w:guid`, the styles, the main document's style, reloaded, saved again nothing added twice), added to (65b's), and left untouched (66b's, byte for byte but `_rels`, and an entry added leaving its styles part); section 63.18's two (an entry of two paragraphs and two runs, as revised by 63.21 at `0dc354a`; a missing entry replaced, as revised by 63.22 at `a3a9571`); check 68's four; section 63.22's two. **This package's own:** a never-edited `samples/Normal.dotm` saved with its glossary byte for byte after every placeholder is read and the outline taken; a created document's placeholders saved, reopened and read back through `placeholder_text`, and its glossary parts unmarshalled and saved again canonically identical (`scripts/canon.py`, `normalise=True`, no category at all); a trial's entries its own; the same calls make the same glossary; the empty bound value of a date picker; `glossary_entries` read before any body; the save hook in the engine (`test_parts_and_relationships.py`).

### 12.5 Departures and choices, with why

1. **The glossary is read on first use as well as when a body is first read.** Reads here are synchronous, so `GlossaryEntries` reads the part whenever it is asked and hasn't (and `body_of` reads it, as core-ts's `getBody` does). core-ts's "until the glossary is read, a name is kept" exists because its read is asynchronous; here it never arises, and its test (an unread glossary keeping a missing name) has no counterpart: `test_the_glossary_is_read_on_first_use_where_no_body_has_read_it_yet` stands in its place.
2. **A new entry's name and every `w:guid` a save writes come from the package's seeded generator** (`pkg.id_generator`, as a custom XML part's id is; CR-003 section 3.4), where core-ts draws random ones. Found by the acceptance gate after `dfb05fa`: two runs of `scripts/acceptance.py` wrote different glossaries. The setter draws again while the glossary holds the name: a trial's generator starts where the real package's did, so its first name was the real package's waiting one (`test_a_trial_has_its_own_entries_thrown_away_with_it` found it). docx4j gives the default entry check 34's `w:guid` (core-ts section 63.18's last bullet); neither port does.
3. **A shown placeholder's style is defined in the main document's styles part** where it lacks it (`require_placeholder_style`, from `set_bound_content(..., True)`), with Word 365's definition (`placeholder_style()`, core-ts's `placeholderStyle.mts`, `w:color 666666`), not `KnownStyles.xml`'s Word 2010 one (`808080`). This is core-ts's `requireStyles([PLACEHOLDER_STYLE])` (its CR-002 section 40: Word 2010 drops a `w:rStyle` naming a style the part lacks), which this package had not ported; without it a document made here would hold placeholders naming an undefined style. core-ts splices it on save; here it is added at once, as a hyperlink's style is (CR-003 section 21.3 item 2), so the write's report names `/word/styles.xml`. A styles part defining it is only read (the property resolver's private reading).
4. **The save hook is a list, `OpcPackage.save_hooks`,** each hook returning the names of the parts it made, gathered into `SaveReport.added` (`package.last_save.added`, section 4.9's "last_save lists the parts a save added"). core-ts's `saveTo` calls `glossaryEntries.flush` by name; here the engine knows no glossary.
5. **A write's report names the glossary document part** where a call leaves an entry waiting for the save (core-ts names nothing), as this package's reports name every part a call will change.
6. **`default_placeholder_entry` replaces `is_text_control`**, as core-ts's `defaultPlaceholderEntry` replaced its `isTextControl` at `b12e0c5`; both are extensions, never released. They are public here, as `shows_placeholder_content` is, where core-ts marks them internal.
7. **`set_bound_content` doesn't hold a bound block control to one paragraph**, where core-ts's does (its section 61's `setBoundContent`); this predates CR-019. `show_placeholder` gives a bound block control the entry's first paragraph alone, as core-ts does; any other block it holds stays.
8. **The bare form's look** (section 11.5): `dfb05fa` kept Phase A's rule (the first placeholder run's properties less `PlaceholderText`), as section 11.5 decided pending check 68; `0dc354a` replaced it with check 68's (D11), as core-ts's `c275a2a` decided.
9. **`entry_run_properties`,** `dfb05fa`'s one helper combining a control's run properties with an entry run's (the coordinator's instruction, pending check 68), went at `0dc354a`, where there is no combination left: `entry_content_of` copies. The tests that pinned the combination were one test, and moved as one.

### 12.6 Pre-existing tests whose expectation changed

- **`dfb05fa`:**
  - `test_controls_typed.py`: `test_placeholder_text_is_written_and_refused_over_content` becomes `test_placeholder_text_names_a_glossary_entry_and_a_control_holding_content_keeps_it` (a made control reads Word's default text and entry; the setter's 32 hex digits; a control holding content keeps it, where `binding.holds_content` was raised; reloaded, both read their entries); `test_showing_placeholder_text_is_office_js_s_name_for_the_flag`: a made control shows its placeholder (`True`, it was `False`), and with the flag off `placeholder_text` still reads the entry (it read `""`).
  - `test_readers.py`: CR-016 section 18's test reads `[None, None]` (the runs' `"Click here."`); the look test and the refused write read `None` (`""` and `"Click here."`); `test_apply_bindings_writes_docx4js_placeholder_for_an_empty_value_which_reads_as_no_text` becomes `…_shows_words_default_placeholder_for_an_empty_value_in_a_text_control_…` (Word's default entry, not docx4j's `"Click here to enter text."`, reloaded too); check 58's case 11 compares `placeholderText` (null, Word's; it was left out and pinned to the runs).
  - `test_ranges.py`: check 53's 34 and 35 differ in `showingPlaceholderText` alone (`control.placeholderText` too, before).
  - `test_placeholder_writes.py`: check 64's run is Word's in all 52 writes (`DIFFERS_64` empty; the test renamed `test_check_64s_run_is_word_365s_in_every_write`, with `placeholderText`); its saves' 14 made-control regions pinned to Word's id and run boundary (they were pinned as differing, `MADE`); Phase A2's Word's-form emptied control shows its entry (a test of its own, block and run; it was left empty); check 63's run-2 test holds `placeholderText` null.
  - `test_office_js_characters.py`: `SHA1SUMS` counts 84 (59).
  - `test_private_readers.py`: the two new callers listed, and a `placeholders` reader reaching them.
  - `tests/agent/test_template_workflows.py`, unmoved: two controls made at one point kept their order once `_insert_at_point` went past what carries no text (12.2 item 6); a placeholder made at the point carries text to `text_of`, which the old placing read.
- **`0dc354a`:**
  - `test_readers.py`: `test_every_location_writes_in_the_placeholders_place_the_flag_cleared_the_style_gone_the_runs_own_look_kept` becomes `…_nothing_of_the_placeholder_runs_look_taken`: the run-level placeholder's bold is not kept (no `w:rPr` at all). **CR-016's Word check, file 2, region 7** (a run-level placeholder whose run was bold too, written at `"End"`) showed that write bold, CR-016 Phase B's rule; check 68's cases 1 to 4 (the placeholder italic, the control with no run properties) give Word's plain run, so a rerun of that region would differ from what was checked.
  - `test_placeholder_writes.py`: every route into a block placeholder: no bold in either form (the bare form's was bold), and, a new parameter, a control's own italic taken.
  - `test_placeholders_words_form.py` (`dfb05fa`'s): section 63.18's test: the entry's runs as held (`on`, `unset`, `unset`, no colour; they were `off` and `FF0000` throughout), the helper's assertions gone with it.
  - `test_office_js_characters.py`: 89.
- **`a3a9571`:** section 63.18's missing-entry test: the date control naming a missing entry is given the date picker's default (it kept the name); `test_office_js_characters.py`: 90. `bindings.PLACEHOLDER_TEXT`'s new words moved no test: none held `"Click here to enter text."`.
- **Unmoved:** checks 63's run 2, 64's writes, 67 and Phase A2's every other pin; the filler (check 52's case 13, D10 (a)); the D3, D6 (b) and D7 tests.

### 12.7 Acceptance

`out/acceptance-before-cr019b/` (a copy of `out/acceptance/` at `caff8cf`, equal to a regeneration at `caff8cf` in a scratch worktree: 17 artefacts, 243 entries, 0 differing beyond the dates) against `out/acceptance/`:
- **`9-template-filled.docx`, the one that moves:** its control made at the end of `Approved by: ` (`insert_content_control("PlainText")` at a point) now shows Word's default entry before `apply_bindings` fills it with `Grace Hopper`, so it keeps `w:placeholder` naming `DefaultPlaceholder_-1854013440`; the save makes the glossary (`word/glossary/document.xml`, `styles.xml` and its `.rels`, the relationship and two content-type overrides); and `word/styles.xml` gains Placeholder Text (invoice2013's styles lacked it; 12.5 item 3). Nothing else in it moves.
- At `0dc354a` only that glossary's `w:guid` moved against `dfb05fa`'s, seeded now (12.5 item 2), and two regenerations agree; at `a3a9571` nothing moved.
- Every other artefact is unchanged: none makes a control on an empty paragraph or writes into a placeholder.

### 12.8 The gates

| gate | `dfb05fa` | `0dc354a` | `a3a9571` |
|---|---|---|---|
| the full suite (6,401 passed, 8 skipped at `caff8cf`, this checkout) | 6,510 passed, 8 skipped (4 min 47 s) | 6,537 passed, 8 skipped | **6,540 passed, 8 skipped** (4 min 20 s) |
| `roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata`, `--normalise` and `--no-normalise` | 50 of 50 both ways, 0 differences, 34 `boolean-spelling`, 0 kept, 0 normalised, 0 skipped | the same | the same |
| acceptance (`compare_acceptance.py out/acceptance-before-cr019b out/acceptance`) | 17 artefacts, 243 entries, 1 differing (12.7) | the same | the same |
| `scripts/compare_word_resolution.py --no-records` | 0 lines (2 compared, 0 differ) | the same | the same |
| `codegen/generate.sh --check` (a scratch worktree of `HEAD` with the commit's files copied in before committing, `TMPDIR` in the scratchpad) | "two generations of docx4j_py are byte identical", 308 modules (307: `placeholders.py` the one more); the tree left as it was | the same | the same |
| `ruff check` on the files changed | no new finding | the same | the same |

The brief's figure of 6,401 "plus `4616c68`'s 5" double-counts: `4616c68` is before `dbce995`, whose 6,401 includes it, and `caff8cf` gives 6,401 here.

### 12.9 The port of `c275a2a`: check 68 (D11, section 63.18 revised)

core-ts's check 68 (its section 63.19; the user's run, Word 365 16.0.20326.20232, Design Mode off: neither save carries `w:formsDesign`): the placeholder run, or the entry, italic, and the control's `w:sdtPr/w:rPr` bold, so a result shows where each property came from; the bare form (cases 1 to 12, the run italic, plain or bold, the control with no properties or bold), Word's form (13 to 20), and Word's-form controls holding a bold value emptied (21 to 24, the default entry or an italic custom one).
- **Word's answer** (core-ts's section 63.20): a write takes the control's own run properties and nothing of the placeholder's, in both forms, at `"Replace"` and `"End"`, block and run (1 to 4 plain, 5 to 20 bold); an entry Word brings back is as the glossary holds it, the control's bold not over it (21 and 22 the default in Placeholder Text alone, 23 and 24 the custom entry in Placeholder Text and italic), the control's `w:sdtPr/w:rPr` kept. **This contradicted both the rule `dfb05fa` built and Phase A's look:** `aefc41f`'s section 63.18, docx4j's `createPlaceholder` order, put the control's properties over the entry's runs, and Phase A gave the bare form's text the placeholder run's own properties less the style. The coordinator's instruction, while `dfb05fa` was being built, was to keep the combination in one helper so that a follow-up could change it in one place; core-ts's decision ("decisions accepted") came before the follow-up was needed.
- **Built** (`0dc354a`): `_write_placeholder` gives each run the write leaves in the control a copy of `w:sdtPr/w:rPr`, and leaves the runs plain where the control has none (`_take_control_properties`), by every route; `entry_content_of` copies an entry as the glossary holds it; the setter and the no-entry default write the entry's run alone. Unchanged, as core-ts has them: the filler keeps the control's run properties (check 52's case 13); docx4j's run for a kind naming no entry; a write into an empty control not showing a placeholder (unmeasured).
- **The files** (`tests/office_js/`, `SHA1SUMS`): `68-placeholder-run-properties.docx` (core-ts's git-ignored copy, written by `build.mjs` at 10:25, Word's run at 10:32), Word's report and run JSONs, and Word's two saves, as `test/check68/manifest.json` has them. **The port**, `run_check_68`, `check68-script-lab.js` step for step (check 66's snippet with check 68's ops: `control-text-replace`, `control-text-end`, `content-delete`, tracking off).
- **Tallies:** check 68's run Word's in 4 of 24 writes at `caff8cf` and at `dfb05fa` (9 to 12, where the placeholder run and the control were both bold), **24** after, its saved controls 4 and **24**; checks 63's run 2, 64, 65, 66 and 67 unmoved (their controls have no run properties of their own), as core-ts found.

### 12.10 The port of `b12e0c5`: each kind's default entry (section 63.22)

The user made one control of each kind from Word 365's Developer tab and saved it (`placeholders.docx`, kept by core-ts as `test/placeholder-defaults/placeholders-word365.docx`, SHA-1 `9cd6fb70e5af3b6d52eab2e95bd13acfbcf3c2c8`, copied here as `tests/office_js/placeholders-word365.docx`; its settings carry `w:formsDesign`, which bears on edits into a placeholder, not on what Word makes). Each control names its kind's entry with the flag on; each entry is in the glossary with the default entry's `w:docPartPr` and a `w:guid` of its own, one run in Placeholder Text:

| kind | entry | text |
|---|---|---|
| rich text, plain text | `DefaultPlaceholder_-1854013440` | Click or tap here to enter text. |
| date picker | `DefaultPlaceholder_-1854013437` | Click or tap to enter a date. |
| drop-down list, combo box | `DefaultPlaceholder_-1854013438` | Choose an item. |
| building block gallery (`w:docPartList`) | `DefaultPlaceholder_-1854013436` | Choose a building block. |
| repeating section item | `DefaultPlaceholder_-1854013435` | Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table. |
| picture, check box, repeating section | none | none |

- **Built** (`a3a9571`): `DEFAULT_ENTRIES`, `DefaultEntry` and `default_entry_for` (`placeholders.py`), `ContentControl.default_placeholder_entry`. A control naming no entry, or one the glossary lacks, is given its kind's default wherever it shows its placeholder: the binding pass's empty value, an emptied Word's-form control, and `insert_content_control` on an empty range, which now makes a date picker, a drop-down list and a combo box show theirs too. The entry is written into the glossary on save in Word's form, its `w:guid` the package's (core-ts draws a new one; docx4j gives `placeholders.docx`'s). `placeholder_text` set to the kind's default text, or `""`, names that kind's entry. A kind with no default, naming none, shows docx4j's run with no reference; naming a missing entry, it keeps the name. **`bindings.PLACEHOLDER_TEXT`** is Word 365's "Click or tap here to enter text." (docx4j's D5), shown only for a kind with no default entry now.
- **Not done, as core-ts:** the lists' first `w:listItem` ("Choose an item.", which Word gives a list it makes) isn't emulated, and the building block and repeating section kinds aren't given a placeholder when made.
- **Tests:** core-ts's two, the table held to Word's file control by control (type, entry, `placeholderText` and this package's default; Word's entries' markup), and a date picker, the lists, a plain-text control and a check box made, the date picker emptied after a write, its placeholder set to its default text, the glossary on save, and a picture showing docx4j's run; this package's own, an empty bound date value. Checks 63 to 68 unmoved.

### 12.11 The cost

The median of 7 passes after a warm one, `caff8cf` (a scratch worktree) and `a3a9571` alternated, two rounds (ms):

| | before | after |
|---|---|---|
| load and the first `pkg.body`, `samples/Normal.dotm` (a glossary of 13 entries) | 1.0 to 1.1 | 3.7 to 3.8 |
| load and the first `pkg.body`, `samples/sample-docx.docx` (none) | 6.4 to 6.6 | 6.4 to 6.9 |
| 2,000 `pkg.body` reads | 8.4 to 9.1 | 8.8 |
| 100 controls made on empty paragraphs | 39.5 to 43.4 | 91.2 to 94.9 |
| 20 made and saved | 14.2 to 14.4 | 23.5 to 23.6 |

About 2.7 ms to read a glossary of 13 entries when a body is first read; nothing measurable for a document with none, or for a `pkg.body` read after the first. A made control costs about 0.5 ms more: `show_placeholder` itself about a tenth of a millisecond, and the rest `next_sdt_id`'s walk over the whole tree on each insert, longer now that each made control holds a run (that walk predates CR-019). Saving a created document (medians of 9): 20 controls holding text, no glossary, 8.0 ms; 20 placeholders naming the default entry, the glossary, its styles part and the main document's style made, 12.3 ms; 20 placeholders of their own text, 20 entries, 29.7 ms, about 0.9 ms an entry beyond the first (`doc_part_of` parses each entry's markup with `wml()`).

### 12.12 Found, and not done

1. **Phase C's Word check** (Q5) has more to look at: a created document's glossary, the main document's new Placeholder Text (12.5 item 3, unmeasured in Word 365 itself), the date picker's and the lists' made placeholders, D11's look, and artefact 9's glossary.
2. **CR-016's Word-checked file 2, region 7** (12.6, `0dc354a`) would now be written plain; its record stands as what was checked then.
3. **`set_bound_content`'s bound block control** keeps any block beyond its first paragraph (12.5 item 7), where core-ts's section 61 removes them; predates CR-019.
4. **A paragraph's `insert_content_control` under tracking** writes no `w:customXmlInsRange` markers here (CR-003's rule, a control is no revision), where core-ts marks it inserted; predates CR-019.
5. **`next_sdt_id` walks the whole tree** on every control made (12.11).
6. **Section 10.9 item 2** stands: a tracked `ContentControl.insert_text("Start" | "End")` on a run-level control holding text writes untracked.
7. **The records for Phase C** (the coordinator's): `CHANGELOG.md` under "Changed" (`placeholder_text`'s reading and setter, Q2; the binding pass's Word's form, Q3; made controls and their glossary; D11's look, which replaces CR-016 Phase B's "the run's own look"; `PLACEHOLDER_TEXT`'s words), `CLAUDE.md`'s CR-016 and CR-019 lines, `tests/README.md`, CR-016 section 13.5's pointer, CR-014 section 4's note (the binder's empty value), CR-003's member notes (`placeholder_text: str | None`, `show_placeholder`, `glossary_entries`, `default_placeholder_entry`, `shows_words_form_placeholder`).

## 13. Phase C notes (2026-10-25)

Phase C as done, on `551860e` (the coordinator's `CLAUDE.md` commit after `8bc2eeb`; no code has changed since `a3a9571`). No code changes: the records, in one commit (this section's), and the Word check's files (Q5), built from the checkout into `out/cr019-word-check/` (git-ignored) for the coordinator to stage. **The Word check is not yet run**; its record goes here and in `tests/README.md` once the user has run it.

### 13.1 The records

| record | what changed |
|---|---|
| `CHANGELOG.md`, `[0.2.0]` | **The summary:** CR-019's bullet (seven CRs); CR-016's no longer claims the placeholder writes. **Changed:** CR-016 Phase B's placeholder bullets rewritten, since CR-019 supersedes them before release: `ContentControl.insert_text`'s (its look, the run's own less the style, never released) is folded into CR-019's every-route bullet, the `range.holds_placeholder` refusal (never released) is gone, D3's bullet in its place, and the reading bullet points to CR-019 for `placeholder_text`. CR-019's ten bullets: every route and D11; D3, with the tracked equation Replace no longer refused; D4 and Q4; D6 (b); D7; D10; `placeholder_text` `str \| None` and its setter, `binding.holds_content` gone (0.1.0 shipped it); the binding pass in Word's form with each kind's default entry and `PLACEHOLDER_TEXT`'s words; made controls; the glossary on save and the main document's Placeholder Text. CR-013's D2 bullet no longer keeps the equation refusal. **Added:** the `ContentControl` extensions, `package.glossary_entries`, `OpcPackage.save_hooks` and `last_save.added`, the check fixtures. **Fixed:** CR-016's placeholder bullet widened to every route; a collapsed range's control wrapping its run; a run-level `"Content"` range's `"Start"` writing outside; the empty run-level placeholder Word wouldn't leave Design Mode over. Each "0.1.0 did" claim was probed on `v0.1.0`'s tree (`git archive` into the scratchpad): every one holds. The date is left alone |
| `tests/README.md` | **The map:** `content/test_placeholder_writes.py`, `content/test_placeholders_words_form.py` and `text/test_docproperty_case.py` (`4616c68`) mapped; `test_controls_typed.py`'s and `test_readers.py`'s rows say what moved (`placeholder_text`, risk 2, the binding pass). **The fixtures:** `docprop-case/`'s row; `tests/office_js/`'s paragraph, and twelve provenance rows for checks 63 (both runs) to 68 and `placeholders-word365.docx`, Apache-2.0, as checks 58's and 59's rows. No run record |
| CR-016 section 13.5 | a closing paragraph: CR-019 is the port, and what it settled for each item |
| CR-014 section 4 | three notes: the `syncBoundNode` difference gone (D7); the empty-node prompt decided (Q3, Word's form); checks 63 and 64 ported (D7, D8) |
| CR-003 section 25 | the member notes: `placeholder_text` (`str \| None`, its reading and its setter), `placeholder_entry`, `show_placeholder()`, `default_placeholder_entry`, `shows_words_form_placeholder`, `shows_placeholder_content` and `package.glossary_entries`, superseding section 3.2's `str` and section 4's refusal |
| `docs/change-requests/README.md` | CR-019's row: its summary (its "the bare form: the run's look kept" was out of date since Phase A2) and its status |
| this CR | this section and the Status line |

`CLAUDE.md`'s CR-019 line is the coordinator's (`551860e`).

### 13.2 The Word check's files (Q5)

Built by `out/cr019-word-check/build.py` (`PYTHONPATH=. .venv-fork/bin/python out/cr019-word-check/build.py`, about 6 s with the acceptance run), from the checkout. It runs the suite's own ports and comparisons (`tests/office_js/checks.py`, `test_placeholder_writes.py`'s and `test_placeholders_words_form.py`'s helpers), with the suite's pins and nothing more. Every save is a default save (Word's normal form), read back with `LoadOptions(strict=True)`: nothing skipped, every part well formed, no `w:formsDesign` or `w:trackRevisions` in the settings.

| file | what it holds | held by the build to |
|---|---|---|
| `cr019-63-run.docx` | check 63's run on `63-placeholders.docx`, tracking as each case says | Word's run 2 in 52 of 52 writes; Word's run-2 save in 49 of 54 regions, the five pinned as the suite pins them (D6 (b)'s 10, 11, 36 and 37; case 18's run boundary) |
| `cr019-64-run.docx` | check 64's run | Word's run in 52 of 52; Word's save in 38 of 52 regions, the 14 made controls Word's but for Word's `w:id` and a run boundary |
| `cr019-65a-run.docx`, `cr019-65b-run.docx` | check 65's runs, the `type` cases left as built, as the snippet leaves them | Word's run in 10 of 10 and 13 of 14 (65b's case 9 pinned), `placeholderText` Word's in every control; Word's saves in 2 of 10 and 10 of 14 controls, the rest pinned (references Word's session dropped, the cases typed by hand, 65b's 9) |
| `cr019-66a-run.docx`, `cr019-66b-run.docx` | check 66's runs, the `type` cases left as built | Word's run in 6 of 6 each; Word's saves in 6 of 6 controls each |
| `cr019-67a-run.docx`, `cr019-67b-run.docx`, `cr019-67c-run.docx` | check 67's report and its four phases, as Word's run | Word's in every step, 24 of 24 writes each; Word's saves in 24 of 24 controls each |
| `cr019-68-run.docx` | check 68's report and run | Word's in 24 of 24, each run's look; Word's save in 24 of 24 controls |
| `cr019-placeholders-written.docx` | regions A to G, from `create_package()` (below) | the assertions below |
| `cr019-acceptance-9-template-filled.docx` | artefact 9, from a fresh `scripts/acceptance.py` run into `out/acceptance/` | its glossary parts there, its control naming the default entry and holding "Grace Hopper"; `compare_acceptance.py out/acceptance-before-cr019b out/acceptance`: 17 artefacts, 243 entries, artefact 9 alone differing, as section 12.7 has it |
| `reference-word365/` | Word 365's own saves after the same runs, the ten `*-run-word365.docx` (check 63's run 2's) of `tests/office_js/` | for the user to compare with, nothing to check in them |
| `REGIONS.txt` | every region of every save, one line each, rendered from the save: controls, prompts, Placeholder Text, bold and italic, tracked insertions, deletions and marks, controls marked deleted | --- |
| `README.txt`, `BUILD-LOG.txt`, `SHA1SUMS` | the user's steps (Design Mode off throughout; All Markup for the tracked regions; each file's regions with every expected difference named; part 2's round trip; what to send back); the build's tallies and commit; the folder's SHA-1s | --- |

**`cr019-placeholders-written.docx`** (`create_package()`, `id_seed` 20261025, no glossary at the start), each region a plain-text label, `A:` to `G:`, before its control:
- **A:** a rich text control made on an empty paragraph: `DefaultPlaceholder_-1854013440`;
- **B:** a date picker and a drop-down list made at points within a paragraph's text, never at its end (13.4 item 1): `DefaultPlaceholder_-1854013437` and `-1854013438`;
- **C:** a plain-text control whose `placeholder_text` is set to "Enter the client's name.": an entry of its own, 32 hex digits;
- **D:** a block plain-text control holding "a stale value", bound to an empty node of a custom XML part, then `apply_bindings()`: Word's form, the default entry;
- **E:** a block plain-text control in the bare form (`w:showingPlcHdr`, one `PlaceholderText` run of `bindings.PLACEHOLDER_TEXT`, no `w:placeholder`), written with `insert_xml` after the binding pass and bound with `set_mapping` to another empty node;
- **F:** a control made on an empty paragraph, its `w:sdtPr/w:rPr` set bold (`put_property`), then `insert_text("Filled", location="End")`: "Filled" bold (D11);
- **G:** a control made on an empty paragraph, written into ("Temporary", `"Replace"`), then `get_range("Content").delete()`: its entry shown again (section 4.4's emptied Word's-form control, check 66 case 3). A control still showing its placeholder would be left as it is by that delete in Word's form (section 4.10), so the brief's (g) is read as the control emptied after a write.

The build asserts: the save added `/word/glossary/document.xml` and `/word/glossary/styles.xml` (`last_save.added`); the main document part's `glossaryDocument` relationship, the glossary's `styles` relationship and both content-type overrides; Default Paragraph Font and Placeholder Text in the glossary's styles, Placeholder Text in `word/styles.xml`; the glossary's four entries exactly the default, the date picker's, the list's and C's own; each control's `placeholder_entry`, `placeholder_text`, flag and text read back from the save (A, B, D, F and G their kind's default entry, C its own, E None and None; F not showing its placeholder and holding "Filled"); each prompt a `PlaceholderText` run of its text; F's run bold, nothing of the prompt left; D and E bound, E with no `w:placeholder`; nodes d and e empty.

### 13.3 Repair risks seen

None is expected to prompt a repair: each file opens strictly here, and checks 63 to 68's saves are Word's region for region but for the suite's pins. What Word hasn't seen before:
1. **A glossary this package made from nothing**, in `cr019-placeholders-written.docx`, 65a's save and artefact 9 (the document part and a styles part of two styles, their relationships and content types; entries in check 65's form, `w:guid`s from the seeded generator). Check 66a measured core-ts's glossary of the same two parts, not this package's; this is part 2's question.
2. **The main document's Placeholder Text**, Word 365's definition added where the styles part lacks it (section 12.5 item 3), unmeasured in Word itself.
3. **A drop-down list with no `w:listItem`** (region B). Word gives a list it makes a first item, "Choose an item."; this package doesn't (section 12.10). The schema allows it, and the README says the list is empty.
4. **65b's and 66b's glossaries are added to:** 65b's one-part glossary gains a styles part, and 66b's document part and its relationships part are re-serialised where case 5's entry is added, its styles part byte for byte (the suite's D9 tests hold both).
5. **Check 63's region 53**, left as built: the empty run-level placeholder Word refused to leave Design Mode over (section 11.1). The README says to keep Design Mode off, and what to do if it sticks.

### 13.4 Found, and not done

1. **A paragraph's `"End"` write beside a run-level control showing its placeholder at the paragraph's end lands inside the control, behind the prompt, with the flag on**, so every Office JS reader reads it as nothing. Probed over `X ` and a run-level control showing its placeholder, ending the paragraph, in either form: `Paragraph.insert_text("Y", location="End")`, `get_range("End").insert_text("Y")` and `Body.insert_text("Y", location="End")` each leave the paragraph reading `X `, the control `""` with its flag on (and saved so); `paragraph.text = paragraph.text + "Y"` takes the control (D3). With text showing, the same writes extend the control (`X VALY`). 0.1.0 also wrote inside the control, beside the prompt, where its readers saw it (`X Click here.Y`); since CR-016 Phase B's reading the text is hidden. None of section 4.2's routes covers it: they reach a run-level control through its own ranges. Word hasn't measured a paragraph's `"End"` beside a trailing control. Not changed here; the Word check's region B keeps text after each control. For the coordinator: a fix here (the write after the control, or through `_write_placeholder`), or a question to core-ts.
2. **The CHANGELOG's summary counts CR-019 among the CRs "new and complete, each Word-checked"**, which holds once its Word check passes.
3. Section 12.12's items 3 to 6 stand.

### 13.5 The gates

| gate | result |
|---|---|
| the full suite (6,540 passed, 8 skipped at `8bc2eeb`) | **6,540 passed, 8 skipped**: records only, nothing moved |
| the Word check's build | every assertion of 13.2 |

Records only, so the round trip, `generate.sh --check` and `compare_word_resolution.py` are Phase B's (section 12.8); the acceptance run is the build's (13.2).
