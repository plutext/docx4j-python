# CR-019: Writing into a content control showing its placeholder, and placeholders written in Word's form --- the port of docx4j-core-ts CR-002 section 63 (`020c039`)

**Status:** Proposed 2026-10-25. On 2026-10-25 the user decided that 0.2.0 waits for core-ts to implement its CR-002 section 63 and for the port here ("Wait for core-ts"; CR-016 section 13.5). core-ts implemented it on 2026-10-25 at `020c039` (`main`, unpushed, unreleased), and this CR is that port. Five questions, **decided 2026-10-25** (the user, section 8): Q1 (c), the bare form's styling held for core-ts's check 67, with core-ts's D2 (a) the fallback if no answer comes by the end of Phase A; Q2 to Q5 as recommended. **Phase A started** the same day. Three phases (section 6): A, the writes; B, placeholders in Word's form; C, the records and the Word check. They come to 3 to 4 days. 0.2.0 follows C, with a new wheel, `release_check.sh` and the wheel's Word check.

**Depends on:**
- **docx4j-core-ts** `020c039`: CR-002 sections 63 to 63.12; its `src/model/content/placeholders.mts`, `ContentControl.mts`, `Paragraph.mts` and `Range.mts`, and `src/model/customxml/bindings.mts`; `test/placeholders.test.mjs` (25 tests); Word checks 63 to 66, with their files, Word's JSONs and Word's saves (`test/fixtures/check63/` to `check66/`, `test/check63/` to `check66/`). Checks 63 to 66 were written and run at `2070694`, `ae9215f` and `2d080d4`.
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

The phases go in order: B's emptied-control rule (4.4) needs A's ranges. Each phase is committed when its gates pass. After C, 0.2.0's own steps: the CHANGELOG date, `release_check.sh`, and the wheel's Word check.

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
