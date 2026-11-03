# CR-034: An absent or unreadable numbering or properties part, read as Word reads it --- Word's default list for a numbered paragraph, and a property field over a missing part (WD 0.76)

**Status:** Proposed 2026-11-04, at the user's word ("draft CR-032 to CR-034"). It ports the two text-model rulings of
WD 0.76 (portfolio `7c5ee35`, recorded with this package's measurements in `4956d94`):
- **A numbered paragraph with no numbering part, or one that can't be read,** is counted as REQ-510's model counts
  Word's default definition (D1 to D4). A new warning, `part.absent`, marks the absent part; `part.not_read` marks the
  unreadable one.
- **A property field, or the metadata, over a core, application or custom properties part that is absent or can't be
  read** reads the part as absent: TITLE and the application properties empty, a custom DOCPROPERTY Word's "Error!
  Unknown document property name.". An unreadable part raises `part.not_read`, once a part (D5 to D8).

Three phases (section 6): A, the properties; B, the numbering; C, the records. About 1½ days (section 9). No Word
check: both rulings are held to Word's own saves of docx4j's broken-part family. Questions in section 8, each with a
recommendation.

*The number:* CR-032 and CR-033 are drafted beside this one the same day.

**Depends on:**
- **CR-031** (complete 2026-11-03): the refusal this reads around, `XmlPart.unreadable`, `PackageTrees.part_not_read`
  and `Extractor.consult`. Phase B replaces CR-031 section 4.6's numbering half in the text model (section 4.4).
- **docx4j's broken-part family** (`docx4j-core-tests/src/test/resources/text-extraction/broken-part/`, `e2af12b74`,
  touched since by `7af22e9dd`), with Word 365's text saves of each file. It is copied in Phase A (section 5.1).
- **docx4j CR-046's section 11 addendum** (WD 0.76; docx4j's defect (d) and the absent part, to come: CR-046 ends at
  section 10.1 at `7af22e9dd`). Not blocking. Its re-pin of `part-not-read/` moves our copy (section 5.3), and its
  phase 0 adds the absent-part files Word has not yet been run on (section 7).
- Nothing in another repository blocks it.

## 1. Why

Measured 2026-11-04 at `f23bbbb`, on docx4j's broken-part family (exported by `git archive e2af12b74` into the
scratchpad), against Word 365's text saves of the same files. Every row differs from Word, and none raises a warning
where the ruling asks for one.

**Numbering** (`list_labels: true`; Word's save writes one space after a label where ours writes the level's suffix,
a tab, which is the converter difference REQ-1002's NOTE records):

| file | here | Word | warnings here |
|---|---|---|---|
| `numbering` (part cut early) | `first heading`, `second heading`, `first item`, `second item` | `1. first heading`, `2. second heading`, `1. first item`, `2. second item` | `part.not_read /word/numbering.xml` |
| `numbering-absent` | the same, no labels | the same as above | none |
| `numbering-order` (cut; heading, item, item, heading, item) | no labels | `1.`, `1.`, `2.`, `3.`, `4.` | `part.not_read /word/numbering.xml` |
| `numbering-order-absent` | no labels | `1.`, `1.`, `2.`, `3.`, `4.` | none |

Markdown, with labels off, writes the direct list as bullets (`- first item`): `_numbering` builds a `TextNumbering`
with no format (`extractor.py:1180`), and `_create_list_block` makes a level with no format a bullet list
(`markdown.py:1120`). `TextNumbering.value` is None.

**Properties** (`fields: update`; each file's three fields are TITLE, DOCPROPERTY Company and DOCPROPERTY Custom1,
their cached results `*_FIELD_STALE`):

| file | here | Word |
|---|---|---|
| `core` (core.xml cut) | `TITLE_FIELD_STALE`, the other two updated | `""`, the other two updated |
| `app` | `COMPANY_FIELD_STALE` | `""` |
| `custom` | `CUSTOM_FIELD_STALE` | `Error! Unknown document property name.` |

No warning in any of them. Under `metadata: true` the unreadable part's section is left out of the metadata, also
with no warning. **An absent part reads the same way:** with each part's entry, relationship and override removed
from the intact twin, `fields: update` still keeps that field's cached result. This is docx4j's behaviour too. Its
evaluators are the ones CR-025 ported, and docx4j confirmed the absent case the same day (WD 0.76's NOTE).

## 2. What the WD says, and docx4j

- **REQ-510** (0.76): "Where the package has no numbering part, or it cannot be read (REQ-801), a numbered paragraph
  is counted as this model counts Word's default definition: one abstract definition of nine decimal levels, each
  level `w:start` 1 and `w:lvlText` `%{level+1}.`, one instance per `w:numId` other than 0 on that definition, the
  instance a style's `w:numPr` names plain and the instance direct `w:numPr`s name carrying a `w:startOverride` of 1
  at every level, each paragraph at its own `w:ilvl`." The reading raises `part.absent`, or `part.not_read` where the
  part is present but can't be read, once, at the first numbered paragraph. Unmeasured and read from the model until
  measured: a plain instance used after the restarting one has counted (predicted to continue the shared count), two
  direct `w:numId`s, and no style-numbered paragraph.
- **Word's definition, as its repaired save holds it** (the WD's 0.72 item, point 2):
  - one `w:abstractNum` 0, `multiLevelType` multilevel;
  - level n (0 to 8): `w:start` 1, `w:numFmt` decimal, `w:lvlText` `%{n+1}.`, `w:lvlJc` left;
  - at each level, a tab and a left indent at 720·(n+1) twips, hanging 720;
  - level 0 carrying `w:pStyle` Heading1;
  - `w:num` 1 the plain instance, `w:num` 2 the same definition with a `w:startOverride` of 1 on every level.
- **REQ-801** (0.76): `part.not_read` also names a document-properties part, core, application or custom, "where
  `fields: update` reads it for a property field or `metadata` is on". The reading proceeds as for a package lacking
  it, "a property field over a properties part that cannot be read evaluating as REQ-506 has it for the absent part".
  `part.absent` "is raised once per part where a part the reading needed is not in the package and the reading
  supplies what the part would have, today the numbering part for a numbered paragraph". Its selector is "the part's
  name the package would have held", and a part the document did not need raises nothing.
- **REQ-506** lists DOCPROPERTY and the core-property fields among those `fields: update` recomputes. It says nothing
  of a missing part. The 0.76 ruling supplies the reading: TITLE and the application properties empty, DOCPROPERTY of
  a custom property Word's error text.
- **docx4j** (17.3.6-SNAPSHOT at `7af22e9dd`) gives no label in all four numbering files and keeps the cached results
  over a cut or absent properties part. With the custom part cut it keeps all three fields stale, its defect (d). Its
  CR-046 section 11 addendum is to come, and will carry the same two rulings.

## 3. What exists here (measured 2026-11-04 at `f23bbbb`)

### 3.1 The numbering route

- `PackageTrees.__init__` (`trees.py:222-232`) builds `self.emulator` only where the main part has a numbering part:
  `emulator_of(package)`, or `Emulator(package, read_styles=False)` where the styles part couldn't be read (CR-031's
  fix (a)). With no numbering part, `emulator` is None.
- `Extractor._numbering_emulator` (`extractor.py:908`) consults the numbering part for a paragraph with a `w:numId`
  and returns None where it can't be read, raising `part.not_read` through `consult` (`extractor.py:938`).
  `_counting_emulator` (`extractor.py:919`) counts nothing where the part is known unreadable.
- `Emulator.number_of` returns None where `self.part is None` (`emulator.py:672`), so no absent-part paragraph is
  ever counted.
- The property resolver's paragraph merge asks the package's emulator for a numbered level's indent
  (`PropertyResolver.numbering_indents`, `property_resolver.py:327`; `Emulator.get_ind`, `emulator.py:748`). Over an
  unreadable part that raises the part's refusal. The text model then resolves again over `_WithoutNumbering`
  (`trees.py:117-149`; `Extractor._without_numbering`, `extractor.py:1162`; markdown's style indents,
  `markdown.py:606-615`), which gives no level indent.
- `_numbering` (`extractor.py:1180`) builds `TextNumbering(num_id, ilvl, fmt, text, start, value, abstract)`. Where
  the emulator is None, `fmt` and `text` are None and `start` is 1.

**A probe of the design** (scratchpad `wd076/default_def.py`) synthesises Word's definition as a `w:numbering` tree:
abstract 0, nine decimal levels with Word's indents, a plain `w:num` for each `w:numId` a style names, and one with
a `w:startOverride` of 1 at every level for each other `w:numId` a paragraph names. The tree is read by today's
`read_definitions` (`definitions.py:818`) into an `Emulator` subclass whose resolver asks it for indents. Today's
counting, with no change, gives Word's labels on all four files:
- `1. 2. 1. 2.` on `numbering` and `numbering-absent`;
- `1. 1. 2. 3. 4.` on the two order files;
- each paragraph at left 720, hanging 720.

The counters are per abstract definition and level, and a `w:startOverride` is spent on an instance's first use
(`NumberingState`, `definitions.py:226-279`). That is the WD's model as written, so no counting code changes.

### 3.2 The property route

- `DocPropertyResolver.__init__` (`model/fields/resolvers.py:91`) reads all three parts at once, each through
  `_contents`, which turns any failure into None (`resolvers.py:63`). So a cut part and an absent part are the same
  None, and neither is reported.
- `value` (`resolvers.py:97-123`) looks a key up in the custom properties first, then lets a built-in core or app
  name override it. A key with no value raises `FieldValueError`, and `evaluate` (`evaluators.py:115-145`) and
  `FieldValues.compute` (`model/text/field_values.py:96`) keep the cached result for it.
- `metadata.read` (`model/text/metadata.py:77-99`) reads each part through its own `_contents`, leaving a part that
  won't read out of the maps with no record. The maps are filled whatever the options; the text holds them under
  `metadata: true` (REQ-515).

### 3.3 Reach

- **No committed file is touched by the numbering change** apart from three fixtures that test the case: a scan of
  every `.docx` under `samples/` and `tests/` for one naming a `w:numId` other than 0 with no numbering part, or one
  not well formed, found only `tests/fixtures/part-not-read/part-not-read-numbering-absent.docx`, `-numbering.docx`
  and `-numbering-unneeded.docx`.
- **Phase B moves the part-not-read pins.** The last of the three numbers no paragraph and is unchanged.
  `tests/text/test_part_not_read.py` reads the other two with `list_labels=True` (line 123). So their six expected
  texts gain labels, and `expected-warnings.tsv` gains `part.absent /word/numbering.xml` on numbering-absent's three
  rows (section 5.3).
- No committed fixture has a property field over an absent or cut properties part apart from the broken-part family.

## 4. Design

### 4.1 Word's default definition (D1, Q5)

`model/listnumbering/default_definition.py`, new, hand written, added to `codegen/clean.py`'s `KEEP`:
- **`default_numbering(style_num_ids)`:** a `w:numbering` lxml tree holding:
  - one `w:abstractNum` 0, `multiLevelType` multilevel, nine levels;
  - at level n: `w:start` 1, `w:numFmt` decimal, `w:lvlText` `%{n+1}.`, `w:lvlJc` left, and `w:pPr` with a tab and
    `w:ind` left 720·(n+1), hanging 720, as Word's repaired save holds them;
  - a plain `w:num` on it for each `w:numId` the styles name.

  It is built in memory, read by `read_definitions`, never attached to the package and never saved.
- **`DefaultNumberingEmulator(Emulator)`:** an emulator over that tree.
  - **Its `part`** answers that there is a definition, so `number_of` counts.
  - **Its `_ensure`** reads the tree with the style map the package's styles give: `read_styles` honoured, so
    CR-031's fix (a) holds where the styles can't be read either.
  - **An instance per new `w:numId`:** a `w:numId` met that the styles don't name gets an instance the first time a
    paragraph names it. It carries a `w:startOverride` of 1 at every level, added to the instances through the same
    `read_definitions` reading of a one-`w:num` tree.
  - **A `w:numId` both a style and a paragraph name** is the style's, and plain.
  - **Its `resolver`** is the property resolver of section 4.4, so the merge's level indent is the default
    definition's.
- Level 0 carries no `w:pStyle`. Word's save writes Heading1 there, but REQ-510's owner rule already gives the level
  to the first paragraph style whose own `w:numPr` names it, which is what the family needs. Writing a `w:pStyle`
  would hand a level to a style the document may not have.

### 4.2 Where it is used (D2)

`PackageTrees` gains `default_numbering`, a `DefaultNumberingEmulator` made on first need:
- where the main part has no numbering part, or the part is found unreadable at its first consult;
- not where the document has no main part, and not one level down for a detached body.

`Extractor._numbering_emulator` returns it in both cases, after the warning of section 4.3. `_counting_emulator`
returns it where the part is known unreadable or absent, so `_count` (`extractor.py:1218`) counts. Nothing else in
`_numbering` and `_count` changes:
- `TextNumbering` gets format `decimal`, level text `%{n+1}.`, start 1, abstract id `0` and a counted value;
- under `list_labels: true` the label is `%1.`'s text plus the level's suffix, a tab, as `_list_label` writes any
  label (`extractor.py:2644`);
- Markdown's `_list_paragraph` writes an ordered list for it, as for a real part, and a heading's label leads the
  heading under `list_labels: true`.

### 4.3 The warnings (D3, Q6)

- **`TextWarning.PART_ABSENT = "part.absent"`** (`model/text/result.py`), in the kinds' docstring with
  `part.not_read`.
  - Raised once, with no address. Its selector is `/word/numbering.xml`, the name the package would have held.
  - Raised where today's `consult` would have been asked: the first paragraph naming a list, before the owner test of
    `style_linked_elsewhere`. That is REQ-801's "on first need" and docx4j's `consult`; a paragraph whose style
    names a level it doesn't own still needs the definitions to learn so.
  - One level down, it names the outer package's part, as `part.not_read` does (`Extractor.nested`'s rule).
  - Only the numbering part raises it (Q6): an absent styles or settings part supplies nothing, and the WD names the
    numbering part alone.
- **`part.not_read` for numbering** stays where it is (`consult`), once a part, now with the default definition's
  reading after it.
- A numbering part no paragraph needs raises nothing, as now (`numbering-unneeded`, and
  `test_a_numbering_part_no_paragraph_needs_raises_nothing`).

### 4.4 CR-031's route replaced (D4)

`without_numbering` and `_WithoutNumbering` (`trees.py:117-149`) become a resolver whose `numbering_indents` is the
default definition: `_DefaultNumberingResolver(PropertyResolver)`, its `numbering_indents` the
`DefaultNumberingEmulator`. It is used:
- where the numbering part is absent: today's resolver finds no emulator, so the merge gives no level indent, where
  Word's default definition gives one;
- where the numbering part's refusal is met: `Extractor._without_numbering` and markdown's style indents, as today.

A numbered paragraph's `indent_left` and `indent_hanging` are then Word's, 720·(n+1) and 720.

**`Emulator(package, read_styles=False)` stays.** It reads numbering without styles where the styles part can't be
read, which this CR doesn't change. The default emulator honours it.

**The content API is untouched.** CR-031 A2's list readers still refuse an unreadable numbering part over a
paragraph naming a list. An absent numbering part still gives no list item there. The WD is the text model's
specification, and the content API's lists are Office JS's (Q7).

### 4.5 A property part read as absent (D5, D6, Q1 to Q4)

`DocPropertyResolver` (`model/fields/resolvers.py`) reads each part lazily, the first time a key needs it, and
records, per part, whether it was absent, read, or unreadable. Unreadable means its read raised the part's own
`XmlPart.unreadable`, as `PackageTrees.part_not_read` tells it; any other failure is absent, as now. The resolver
keeps the parts it found unreadable: `unreadable_parts`, each part once, in the order met.

`value(key)`:
- **A built-in core name** (`_CORE`) needs only the core part, and **a built-in application name** (`_EXTENDED`) only
  the application part. The custom part is not read for a built-in name: it can't change the answer, since the
  built-in lookup overrides it (`resolvers.py:113-117`).
  - Where that part is absent or unreadable, the value is `""`: the core text properties (Title, Author, Comments,
    Keywords, LastSavedBy, Subject, so TITLE, AUTHOR, COMMENTS, KEYWORDS, LASTSAVEDBY and SUBJECT too) and every
    application property (Company, Template, Pages and the rest).
  - **Not dates:** CreateTime, LastSavedTime, LastPrinted and RevisionNumber, so CREATEDATE, SAVEDATE, PRINTDATE and
    REVNUM, keep their cached result, as now. Word over a missing core part is unmeasured for them (Q3).
- **Any other name** needs the custom part. Where it is absent or unreadable, the value is Word's `Error! Unknown
  document property name.`. It is returned as the field's result with no formatting switch applied (Q4).
- **A part that is present and read** is unchanged: a key with no value there still raises `FieldValueError`, and
  the field keeps its cached result (Q2).
- `Bytes` stays None.

`DocVariableResolver` and every other evaluator are unchanged.

### 4.6 The warnings for properties (D7)

- **Fields.** `FieldValues.compute` (`model/text/field_values.py`), after each evaluation, hands the resolver's newly
  unreadable parts to the extractor. The extractor warns `part.not_read` once a part, its selector the part's name
  (`/docProps/core.xml`, `/docProps/app.xml`, `/docProps/custom.xml`). So it is raised only where `fields: update`
  evaluated a property field over the part. A DOCPROPERTY Company over a cut custom part raises nothing, the custom
  part not being read.
- **Metadata.** `metadata.read` distinguishes the unreadable part from the absent one in the same way, and
  `Extractor.extract` warns `part.not_read` for each unreadable one where `metadata: true`. Under `metadata: false`
  the maps are still filled, the unreadable part still left out, and nothing is raised: the text holds no metadata,
  and REQ-801 asks for the warning where `metadata` is on.
- **No duplicates.** One part read by both a field and the metadata raises one warning. The extractor's `warn` keeps
  `part.not_read` once a part.
- **One level down,** an embedded package's properties are not read (`metadata` is turned off at
  `extractor.py:831`), and its fields name its own parts under the outer part's name, by `Extractor.nested`'s rule.

### 4.7 The CHANGELOG lines (Phase C)

- **Changed:**
  - the text model counts a numbered paragraph with no numbering part, or an unreadable one, with Word's default
    definition, labels and indents;
  - a property field over an absent or unreadable properties part evaluates as Word's does.
- **Added:** `TextWarning.PART_ABSENT`; `part.not_read` for the three properties parts.
- **Fixed:** a property field over a cut properties part kept its stale cached result with no warning.

### 4.8 Not here

- **The relationships prefix** (WD 0.76 point 3) is CR-033's option. Where it is built, a properties part only a lost
  package relationship named reads as absent through section 4.5. The family's `package-rels` Word save is then the
  oracle for both (section 5.2).
- **A styles part cut late kept as its prefix** is the "one rule" still open in the WD.
- **The caps language** is CR-029, amended 2026-11-04 for WD 0.80.
- **docx4j's fixes (b) and (c)** stay with CR-031 section 10.8.

## 5. Tests and oracles

### 5.1 The fixture (Phase A)

docx4j's whole broken-part family, 63 files and 732 KiB:
- **Copied unedited into `tests/fixtures/broken-part/`** by `git archive <commit> <dir> | tar -x` into a scratch
  directory, then into place, at docx4j's latest commit touching it when Phase A runs (`7af22e9dd` today, which
  relabelled the intact rows' REQ-607 tokens).
- Its README as `docx4j-README.md`, and a `README.md` of ours naming the commit.
- `SHA1SUMS` of the copied files, each docx4j's git blob, held by a copy test as `test_the_copy_is_docx4js` holds
  part-not-read's.
- The builder, `build_broken_part.py`, is copied with the rest but not run.

The whole family is copied, not the seven files this CR reads. CR-033's relationships prefix and CR-029's caps runs
read the same family, and the copy is cheap.

### 5.2 The tests (`tests/text/test_absent_or_unreadable_parts.py`, new)

Each file is held to the lines of Word's `-word365-utf8.txt` save that the ruling covers. A label is compared with the
separator normalised, Word's one space against the level's tab, the converter difference of REQ-1002's NOTE.
- **Phase A, properties** (`fields: update`):
  - `core`, `app` and `custom` against Word's saves, each with its one `part.not_read`;
  - the intact twin with each part removed (entry, relationship and override, built in the test as section 1's
    probe built it) reading the same, with no warning;
  - a DOCPROPERTY Company over the cut custom part: updated, no warning;
  - `metadata: true` over each cut part: its section absent, its `part.not_read`; `metadata: false`: nothing raised;
  - CREATEDATE and REVNUM over a missing core part keeping their cached results (Q3);
  - the error text taking no switch (Q4).
- **Phase B, numbering** (`list_labels: true`, and Markdown with labels off):
  - all four numbering files against Word's saves, `1. 2. 1. 2.` and `1. 1. 2. 3. 4.`;
  - `part.absent` on the two absent files, `part.not_read` on the two cut ones, each once;
  - `TextNumbering` format `decimal`, level text `%1.`, a counted value;
  - Markdown's ordered lists, and each paragraph's indents at 720 and 720;
  - `numbering-unneeded`: nothing raised;
  - styles and numbering both cut: the direct items counted, as fix (a) counts them;
  - one level down, the outer part's name.
- **The WD's unmeasured cases** are pinned as the model predicts them, so that a measurement which moves them is
  visible: a plain instance used after the restarting one, two direct `w:numId`s, and no style-numbered paragraph.
  They are built in the test.

### 5.3 The part-not-read pins (Phase B, Q8)

Under `list_labels=True` the numbering and numbering-absent files' six expected texts gain Word's labels, and
numbering-absent's three warning rows gain `part.absent /word/numbering.xml`. docx4j's addendum will regenerate the
same files.
- If it has landed when Phase B runs, the copy is re-taken at its commit by the same `git archive` route. That also
  takes `7af22e9dd`'s ISTANBUL lines, which stay pending until CR-029.
- If not, `test_part_not_read.py` keeps docx4j's files byte for byte and lists the moved lines and rows as CR-034's,
  as it lists CR-029's pending lines today (line 97). They are named with their reason and checked against this
  CR's reading, and the re-copy removes the list.

### 5.4 The suite

The full suite, the conformance harness and the goldens are run before and after each phase. No golden or harness row
is expected to move, by section 3.3's scan. One that does is recorded with its reason. CR-031's tests whose
expectation changes are named in the phase's notes: `test_a_part_that_cannot_be_read_reads_as_a_part_that_is_absent`
for numbering keeps its equality (cut and absent now both read Word's default definition), its warnings differing by
kind.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | The fixture copy (5.1); `DocPropertyResolver` read lazily per part, absent and unreadable told apart; the empty values and the custom error text (4.5); `part.not_read` from fields and metadata (4.6); the properties tests | ½ day |
| B | `default_definition.py` and the default emulator (4.1, 4.2); `part.absent` (4.3); CR-031's numbering route replaced (4.4); the numbering tests; the part-not-read pins (5.3) | ¾ day |
| C | The records: implementation notes, the Status line, the index, CHANGELOG [Unreleased], CLAUDE.md's text-model sentence in "Nothing may be dropped silently", `tests/README.md` | ¼ day |

Each phase commits on its own, and its registry entry changes in the same session, through the portfolio.

## 7. Risks

- **Word's reading of a package with no properties part is measured only once, through a repair.**
  - **The file.** `broken-part-package-rels.docx` has its `_rels/.rels` cut after the main part's relationship. So the
    three properties parts, themselves intact, are named by no relationship, and are not part of the package.
  - **Word's save of it** reads TITLE and Company empty, and Custom1 as the error text: the same as the three cut
    files.
  - **The gap.** That reading came after the relationships part's repair prompt. A package that simply lacks the parts
    hasn't been run in Word. docx4j's addendum phase 0 adds such files with a Word check.
  - **If Word reads one differently,** the absent case moves and the unreadable one stays.
- **The default definition's unmeasured cases** (section 2) are read from the WD's model. A later measurement may
  move them; the tests pin them so that it shows.
- **Markdown over a previously unlabelled list changes shape**, from bullets to an ordered list. That is the point,
  but a caller who diffed Markdown output over such a file sees it move. Only three committed files have the shape.
- **`DocPropertyResolver` read lazily** changes when each part is read, not what a present part answers. The CR-025
  field tests hold it.
- **A field evaluated outside the text model** (`model.fields.evaluate` with a `FieldContext` of the caller's) also
  gets the empty values and the error text, since the resolver is shared. That is the same reading, and no warning
  path is there: the caller can ask the resolver's `unreadable_parts`. Section 4.5's change is the evaluators', which
  CR-025 ported from docx4j, so docx4j's addendum should make the same change; the portfolio is told.

## 8. Questions

| # | Question | Recommendation |
|---|---|---|
| Q1 | Does an absent properties part read as the unreadable one does (empty, the error text)? | Yes. The ruling says "as absent", and Word reads `package-rels`, whose properties parts no relationship names, that way (section 7). Revisit if docx4j's phase-0 Word check of a package that simply lacks them differs. |
| Q2 | A property missing from a part that is present (no `dc:title`; a custom name the custom part lacks): empty and the error text too? | No, not here. It stays the cached result, as REQ-506's NOTE has it for Word's other error results. It's outside the ruling and unmeasured. The disagreement with the absent part's reading is named in the notes, and a probe is asked of the portfolio. |
| Q3 | CREATEDATE, SAVEDATE, PRINTDATE and REVNUM over a missing core part? | Keep the cached result until Word is measured. The ruling names TITLE and the application properties, and Word may show a default date rather than nothing. |
| Q4 | Does a formatting switch apply to the error text? | No: it is Word's error, written as is. Unmeasured; a `\* Upper` over it is pinned so that a measurement shows. |
| Q5 | The default definition's indents and tabs (720·(n+1), hanging 720): Word's, or none? | Word's, from its repaired save. The WD's definition names the format, text and start, and Word draws the indents. Markdown's follow-on paragraphs read them. |
| Q6 | `part.absent` for an absent styles or settings part too? | No. The WD raises it "where the reading supplies what the part would have", today numbering alone; an absent styles or settings part supplies nothing. |
| Q7 | The content API over an absent numbering part: Word's default list too? | No. The WD is the text model's. The content API's list readers keep A2's refusal of an unreadable part and read an absent one as no list. Recorded as a departure between the two readings. |
| Q8 | The part-not-read pins: wait for docx4j's addendum, or move them here? | Don't wait. Re-copy if the addendum has landed when Phase B runs; else keep docx4j's bytes and list the moved lines as CR-034's (5.3). |

## 9. Effort

About 1½ days: A ½, B ¾, C ¼. More than the 1¼ first estimated, for the fixture copy and the part-not-read pins.
The counting itself needs no change (section 3.1's probe).
