# CR-020: Content refused where it does not belong --- a check on every typed list as code adds to it, derived from the model, and a validator for what gets past

**Status:** Proposed 2026-10-26, at the user's request ("draft it as a CR"). The proposal is the user's, put through the portfolio session the same day: refuse misplaced content at runtime (a `Tbl` appended to `R.content`), generated from the annotations. docx4j-25's assessment of the Java side came with it. It is measured here (section 2) with a probe kept in the session's scratchpad; nothing in the repository was changed to measure it. **Revised the same day** at the user's request ("also add the out-of-range report to CR-020's Phase B list"): the validator also reports an integer measurement beyond Word's range, with CR-022's limits table (section 4.7). **Decided 2026-10-26** (the user: "CR-020 Q1 to Q9 answered as recommended"; section 8). **Nothing has started.** Three phases (section 6), 2 to 3 days.

*The number:* CR-019 is the latest CR here, and CR-017 stays reserved for the port of docx4j CR-057 (CR-018's header), so this one is CR-020.

**Depends on:** nothing for Phase A. Phase B's out-of-range report reads CR-022 Phase A's table (`resources/integer-limits.json`, done 2026-10-26; section 4.7).
- **docx4j** `ArrayListDocx4j.checkParent` (`docx4j-generated-objects`) is the counterpart, not a dependency (section 3). It refuses a few known-wrong children of `P`, `R` and `SdtContent`, and since `49637c6e4` (`VERSION_17_3_5`, local) of `P.Hyperlink`. docx4j-25 has proposed a wider set as a docx4j CR, which waits on the user.
- **No fork change, and no generated file changes.** The table is derived at run time from the `XmlContext` metadata the engine already builds (section 4.1). `codegen/generate.sh --check` is a gate only.

## 1. Why

Nothing here refuses misplaced content as code adds it. `ChildList` (`docx4j_py/child.py`) sets the parent of what goes in and nothing more. So `r.content.append(Tbl())` and `P(content=[Tbl()])` succeed, the save succeeds, and Word opens the file with "Word found unreadable content".

The generated types already say what each list may hold. `P.content` is a union of 35 element classes and `R.content` of 34, and neither includes `Tbl`. A type checker catches the mistake. Code that is only run never does:
- **LLM-written code is usually run, not type-checked:** sandboxes, notebooks and most agent loops. mypy skips the bodies of unannotated functions by default, and anything that passes through `Any` gets by.
- **A model can't open Word to see the damage.** It does read a traceback, and it fixes code well from a message that names where the element belongs.

The content API already refuses at its own level (`InvalidTargetError`, for example `model/content/body.py:1070`). The tree layer below it doesn't: the `el` factories, the `wml.builders` sugar, direct list mutation, and python-docx facade users reaching for `.element`.

docx4j has refused the commonest mistakes since `ArrayListDocx4j.checkParent` (section 3). docx4j-25's reading of the forum threads over the years, commonest first:
1. a run or bare text straight in the body or a cell;
2. a drawing straight in a paragraph, without a run;
3. a table or text in a paragraph;
4. a plain string added to a content list;
5. a property object (`sectPr`, `pPr`, `rPr`, `tblPr`, `tcPr`) added as content instead of set;
6. a table built by hand with a level skipped (`P` in `Tr`, `Tc` in `Tbl`).

## 2. What exists here (measured 2026-10-26 at `d35fbd9`)

The probe is a pytest plugin. It patches `ChildList` and `Child.__post_init__` at run time and records, without refusing, every item put in a list outside that (class, field)'s allowed set (section 4.2's first rule, by class only).

### 2.1 The lists, and the parse path

- **Construction.** `Child.__post_init__` hands each list field to its `ChildList`: `set_owner(self)` where the field already holds one (the generator's `default_factory=ChildList`), and `ChildList(value, owner=self)` where it holds a plain list, as in `P(content=[r])`. Both adopt every item (`_adopt_all`). `append`, `insert`, `extend`, `+=` and item and slice assignment adopt what they introduce (`_adopt`).
- **Parsing goes through the same path.** The fork's parser gathers a list field in a `PendingCollection` whose factory is the field's `default_factory` (`ChildList`), and evaluates it to an unowned `ChildList`. It then builds the object with `config.class_factory(cls, params)`, which is `cls(**params)`, so `__post_init__` calls `set_owner`. Mixed content arrives as a plain list (`bind_mixed_objects`) and is promoted by `__post_init__`. So **a check in the adoption path runs on every load** unless the parse switches it off.
- **Every parse in this package uses `runtime.parser_config()`** (`tests/test_parser_configs.py` enforces it). Its `on_bind` hook (`link_children`) wires the parent pointers as each object is built.
- **Parsed content conforms by construction.** A typed list without a wildcard never receives an element outside its choices: an unknown element is skipped and reported (`part.skipped`), and `LoadOptions(strict=True)` and `wml(...)` raise on it. A lax wildcard can hold a root-element class (`wildcard_class`, CR-012 Phase B), but wildcard lists are unchecked (section 2.2). Over about two million parsed items, section 2.4 found no misplacement.

### 2.2 What the model says each list may hold

- **Scale.** 1,969 generated `Child` dataclasses hold 537 (class, list field) pairs. 509 have element choices and can be checked. 28 are wildcards or have no element choices, and stay unchecked: WordprocessingML 5, DrawingML 2, `mce` 2, docProps 2, MathML 4, InkML 1, `oart` 12. Six are mixed content, all outside WordprocessingML, so no WordprocessingML list holds a string.
- **A compound field's own `XmlVar.types` is `(object,)`.** The allowed set has to come from `var.elements`, the element choices.
- **The annotations are strings** (the generator's `DeferredImports`). The metadata `XmlContext` builds is already resolved, so the table is read from it, not from `typing.get_type_hints`.

### 2.3 One element name, several classes

The generator gives an element a class per holder where the schema's types differ by holder. **71 element names** in compound fields have more than one class; `w:hyperlink` has seven (`PHyperlink`, `CtSdtContentRunHyperlink`, `CtSimpleFieldHyperlink` and four more).
- **For 60 of the 71, the classes have one shape:** the same fields holding the same element names, compared one level down. They're interchangeable as far as the XML goes.
- **For 11, the shapes differ:**
  - in WordprocessingML, `sdt` and `customXml` (four classes each, one per form), `ins` and `del` (three each) and `name` (two);
  - in DrawingML, `alphaMod`, `lum` and `tint` (three each), `xfrm` and `effect` (two each);
  - in chartex, `lvl` (two).

So **a check by class alone refuses legal XML**: the engine does exactly that once (section 2.4). **A check by element name alone accepts an `SdtBlock` in `P.content`**, a block-level control holding paragraphs, inside a paragraph.

### 2.4 The probe's results

- **The fast suite** (5,200 passed, 7 skipped; 1,983,407 items checked across 161 (class, field) pairs) tripped at three sites:
  - (a) **The engine:** `model/content/hyperlinks.py:310` (`_wrap`) builds a `PHyperlink` and, for a run inside a run-level control, puts it into `CTSdtContentRun.content`. That list's class for `w:hyperlink` is `CtSdtContentRunHyperlink`. The XML is legal, the save is right, and the two classes have one shape. The rule of section 4.2 accepts it.
  - (b) **A test:** `tests/test_builders_phase_a.py:546`, `test_the_form_can_be_overridden`, builds `sdt([p("a")], form="run")`, a paragraph inside a run-level control. The schema doesn't allow it, and the builder accepts it today.
  - (c) **A test:** `test_run_items_of_every_holder_is_its_own_live_list` (`:986`) appends an `R` to an `SdtContentBlock`. Its holder list counts a block control's content as a run holder, because `traversal.run_items_of` matches any `w:sdtContent` by element name and so returns a block-level list for a block control. No engine path the suite exercises inserts runs into one, but nothing stops a caller.
- **The corpus:** `tests/openpackaging/test_round_trip.py` with `tests/test_roundtrip.py` (131 passed, 36,625 items), and `scripts/roundtrip.py` run in-process (50 of 50 parts canonically identical, 49,016 items). No violations. The corpus is docx4j-25's gate (section 3), and it passes.

### 2.5 The cost

On `samples/Symbols.docx` (`document.xml` 790,857 bytes; 7,991 items checked per parse), with one exact-type set lookup per item and an `isinstance` only on a miss:
- the parse took 167 to 169 ms without the check and 173 to 176 ms with it (+3 to 4%);
- a full load took 207 to 213 ms without and 216 ms with (+2 to 4%).

That's an upper bound. The probe re-walks each list after adoption, where the check would sit inside the loop `_adopt_all` already runs. The parse doesn't need to skip the check for speed (section 4.4 skips it for another reason).

### 2.6 What goes around `ChildList`

- **Whole-list assignment** (`p.content = [...]`). Four hand-written sites do it: `model/content/comments.py:975` and `:1314`, `model/content/controls.py:1593` and `docx/table.py:119`. Each wraps `ChildList(..., owner=x)`, so they're easy to route through the check.
- **Single-valued fields** (`p.p_pr = R()`) are never adopted, so a list check can't see them.
- `deep_copy` copies a tree that already conforms.

### 2.7 docx4j-25's six mistakes, against the derived rule

Each was tested against the rule of section 4.2 on 2026-10-26:

| Mistake | Case tried | Verdict |
|---|---|---|
| 1. a run or bare text in the body or a cell | `R` in `Body` and in `Tc`; `RT` (`w:t`) in `Tc` | refused |
| 2. a drawing straight in a paragraph | `Drawing` in `P` | refused |
| 3. a table or text in a paragraph | `Tbl` in `P` and in `R`; `RT` in `P` | refused |
| 4. a plain string in a content list | `str` in `P.content`, `R.content` | refused (no WordprocessingML list is mixed) |
| 5. a property object as content | `PPr` in `P.content`; `SectPr` in `Body.content` | refused |
| 6. a level skipped | `P` in `Tr`; `Tc` in `Tbl` | refused |
| legal placements | `Tbl` in `Tc`; `R` in `P`; the hyperlink of 2.4 (a) | accepted |

## 3. The counterpart: docx4j

`ArrayListDocx4j.checkParent` runs on `add`, `add(int, E)` and `set`, and throws a `RuntimeException`:
- in a `P` or (since `49637c6e4`) a `P.Hyperlink`: a `Tbl` ("You can't add a table inside a paragraph; it should be a sibling"), a `P`, or bare `Text`;
- in an `R`: an `R`, a `P` or a `Tbl`;
- in an `SdtContent`: a bare `SdtContent`.

It's a refuse-list: it names known-wrong classes and lets everything else through. docx4j-25 recommends keeping it that way in Java, because JAXB loads through the same `add()`. There, a wrong refusal is a load failure on a legal document, the worst outcome for a library whose job is to open what Word wrote. It suggests the derived, allow-only model become a separate validator that walks a tree. Its gate for any new refusal is the test suite plus the fidelity corpus. A rule that trips on real documents is dropped or demoted to a warning.

**Why this package can allow only what the model allows (Q1):**
- The risk behind Java's choice is a load through `add()`. Here the parse can be separated from the check (section 4.4), so the risk is removed by construction, and section 2.4 measured no parsed content that a check would refuse.
- The allowed set is generated and exact, where a refuse-list would have to enumerate each mistake and would miss the rest.
- The one engine write that a class check would refuse is accepted by the shape rule (section 2.3).

Java's own cases are refused here too, with messages in Java's voice (section 4.6), so the two agree wherever Java has an opinion.

**docx4j CR-064 phase 1 shipped (2026-10-26, `befc57371`, local).** docx4j's content lists refuse a misplaced child at add time, with the fix in the message, as a refuse-list:
- a run, `Text`, `Drawing`, `Br`, cell or row at block level;
- a table, paragraph, `Text`, `Drawing` or `Br` in a run-level container;
- a cell, paragraph or run in a table, and a paragraph or run in a row;
- a property object added instead of set;
- a `String`.

A `JAXBElement` is unwrapped first. Its section 7 hands docx4j-python the choice of whether its list or builder layer refuses the same at add time, and this CR (decided 2026-10-26) does, by D2's allow-rule. Phase A checks D6's hand-written messages against CR-064 section 2's table of mistakes worth catching, so that the two ports name each mistake alike.

## 4. Design

### 4.1 The table (D1)

For each (class, list field), the allowed set is read once from the class's `XmlMeta` and cached, built lazily on the first check, so nothing is added at import:
- the classes of each element choice, and each choice's element name;
- **no table**, and so no check, for a field with a wildcard or no element choices (the 28 of section 2.2);
- strings allowed only where the var is mixed.

A shape table is memoised per class in the same way: the fields and the element names they hold, one level down, as section 2.3 measured.

### 4.2 The rule (D2)

An item goes into a checked list when:
1. its class is one of the list's element classes; or
2. its element name (`traversal.element_name`) is one of the list's choices, and its class has the shape of that choice's class; or
3. it is a `str` and the list is mixed.

A `DerivedElement` is judged by the value it wraps and by its qname. An `AnyElement` goes only where there's a wildcard. Anything else is refused.

Rule 2 is what lets the hyperlink of section 2.4 (a) through. It also keeps an `SdtBlock` out of `P.content`, and a `CTCustomXmlBlock` out of a run's holder.

### 4.3 Where it runs (D3)

- `ChildList` gains the field's name in a second slot. It's given by `set_owner(owner, field)`, by `ChildList(iterable, owner=..., field=...)`, by `__post_init__`, and by `link_parents` and `link_children`, which already iterate by field name.
- `_adopt` and `_adopt_all` check before they adopt. `append`, `insert`, `extend`, `+=` and item and slice assignment check through them.
- **A refused `extend` or slice assignment leaves the list unchanged:** every item is checked first.
- A `ChildList` with no field name stays unchecked, as an unowned one is unparented today.

### 4.4 The parse is not checked (D4)

`runtime.parser_config()` sets the fork's `class_factory` to one that builds with the check off, through a context variable. So every parse, and only a parse, is unchecked. That makes a load failure on a legal document impossible by construction, which answers docx4j-25's concern for this package.

The alternative is one context-variable set around each parse call: cheaper per object, but it needs every call site (`XmlPart` unmarshalling, `wml(...)`, `insert_xml`, `insert_ooxml`). Phase A measures both on `Symbols.docx` (Q3). A context variable is per thread, so `scripts/threads.py` stays a gate.

### 4.5 The error (D5)

`MisplacedContentError(InvalidTargetError, TypeError)`, in `model/content/errors.py`, with code `target.misplaced` (Q6). `InvalidTargetError` is already "a container was handed an element it cannot hold", and it carries CR-003's `code`, `message` and `hint` and `to_dict()`. `TypeError` is there so that a caller who catches the language's error for a wrong type catches this one. `except Docx4JException` catches it too.

`child.py` imports the error class at the moment it raises, so the object model's import cost doesn't change.

It's **an exception, not a warning** (Q2): warnings get lost in agent logs, and Java throws.

### 4.6 The messages (D6)

The message names the holder's field and the item, both the class and the element name: "`R.content` can't hold a `Tbl` (`w:tbl`)". The hint says where the item goes.
- **Hand-written** for docx4j-25's six and for Java's cases, in Java's voice. For example:
  - a table in a paragraph, run or hyperlink: "a table is the paragraph's sibling: put it in the body, a table cell or a block-level content control";
  - a run in the body or a cell: "a run goes in a paragraph: `p(r(...))`";
  - text outside a run: "text goes in a run: `r("...")`";
  - a drawing in a paragraph: "a drawing goes in a run: `el.r(content=[drawing])`";
  - a string: "a string is not an element: wrap it, `r("...")`";
  - a property object: "`pPr` is a property, not content: set `paragraph.p_pr`", where the field is found from the holder's metadata;
  - a level skipped: "a table holds rows: `el.tr(content=[el.tc(...)])`";
  - a form mismatch: "this is a block-level content control (`SdtBlock`); in a paragraph use the run form, `el.sdt_run(...)` or `sdt(..., form="run")`".
- **Derived** for everything else: "`w:x` goes in: `Body.content`, `Tc.content`, ..." (up to five), from a reverse index of the same table.

### 4.7 The validator (D7, Phase B)

`misplaced(root)` walks a tree, a part or a package and returns every finding: the holder, the field, the item, what the field takes, and the address where the content API has one. It applies D2 to every list and checks every single-valued field's class as well, which the add-time check can't see (section 2.6). It also carries the completeness rules, starting with CR-016's "a table cell ends with a paragraph" (`table_rules`). Ordering rules belong here too if any are added; the add-time check carries none.

**Out-of-range measurements** (added 2026-10-26, the user). An integer measurement beyond Word's range for its type is reported: an indent, spacing, a page size or a width that Word stores in 16 bits and wraps (CR-022 section 1: 7205001 in a signed twips field read as -3959). The limits are CR-022's table, from docx4j CR-065's probe round 5, and a type without a row isn't checked. **The finding** (agreed with docx4j-core-ts's CR-017, 2026-10-26, so the ports share one name) has the code `value.out_of_range`, the same as CR-022's `MeasureRangeError` for the write side, since it's the same condition, and these fields:
- `element`, a qualified name (`"w:ind"`);
- `attribute` (`"w:left"`);
- `value`, the attribute as written: an `int` where it binds as one, else the kept text;
- `word_reads`, what Word reads (-3959 for 7205001), or `None` where no reading is known;
- `storage`, the measured width: `"int16"`, `"uint16"`, `"8 bits"`, `"5 bits"` or `"int32"`, with a note for CR-022's special rules (first line and hanging one signed field; `w:trHeight` and exact line spacing with the sign as the rule; the tab's ±31,680 and the page width's 144 to 31,681 clamps);
- `range`, the faithful `[min, max]` in the attribute's own unit;
- the part and the path, or an address where the content API has one.

core-ts names the same fields in camelCase (`wordReads`), and its `value` is a number. **The value is never rewritten:** CR-021 Q7 recommends against emulating the wrap. This is the read side of CR-022. It also covers what CR-022's write-side check can't see: a value already in a file, or one assigned straight to a field. This part of Phase B waited for CR-022's table; the rest of Phase B didn't. **The table exists** (CR-022 Phase A, 2026-10-26; told here by its Phase B, 2026-10-27): `docx4j_py/resources/integer-limits.json`, 60 rows keyed by attribute path and context (`null` for the context attribute absent), read through `docx4j_py.model.values.integer_limit(attribute, context)`, `integer_limits()` and `unmeasured_attributes()` (CR-022 sections 10.3 and 10.4). What the report must add itself (CR-022 section 10.10 item 5): the table's `storage` is prose that begins with the token in most rows but not all (`w:firstLine`, `w:hanging`, `w:framePr/@w:h`), and it carries no model of the wrap, so `storage` as a token and `word_reads` are this phase's. What it must read: a row's `word_writes` (Word's save spells 65,535 as `"-1"` in `w:before`, `w:after` and the five unsigned page margins, whose rows stop at 65,534) and `sentinel` (`w:col/@w:space`, every negative saved as -1), without which a report over a Word-saved file takes Word's spelling for an out-of-range value; `test_integer_limits.py` shows the reading.

It would serve the inspector, the MCP server and tests. Optionally a save runs it: off by default, and turned on for the whole test suite by `conftest`, so every save the engine's own edits make is held to it (Q5).

### 4.8 The engine's own sites (D8)

- **The four reassignment sites** of section 2.6 pass the field name.
- **The hyperlink write** of section 2.4 (a) needs no change; D2 accepts it.
- **Test (b)** changes to the form override on a run, plus a `pytest.raises(MisplacedContentError)` for the paragraph. The `sdt` builder then refuses a form its content can't take, through the check.
- **Test (c)** takes `SdtContentBlock` out of the run holders. Phase A also looks at `run_items_of` for a block control's `w:sdtContent` and at its callers. If none needs the block-level list, it returns `None` there, as for any holder that keeps no runs.

### 4.9 An escape hatch (D9)

`with unchecked():`, the same context variable as D4, is for code that builds an invalid tree on purpose: a test reproducing a broken Word file, or a repair tool. It's documented, and it's never used by the engine outside a parse.

### 4.10 Scope (D10)

Every generated namespace: all 509 checkable pairs, not only WordprocessingML (Q7). The python-docx facade (`docx4j_py.docx`) gets it for free, because it builds through the content API and the tree layer.

### 4.11 What still gets past

- Single-field and whole-list assignment by the caller's own code: the validator catches these (Q4).
- `object.__setattr__`.
- The 28 unchecked fields, and what sits inside wildcards.

## 5. Tests and oracles

- **A test per case:** each of docx4j-25's six and each of Java's, checking the refusal, its code and its hint.
- **The legal placements, exhaustively:** for every checkable (class, field), an instance of each choice class goes in without a refusal. A test generated from the same table would check the table against itself, so this test builds the instances from `el`.
- **The shape rule:** the hyperlink case is accepted, `SdtBlock` in `P.content` is refused, and so is each of section 2.3's eleven differing names in the holder of another of its classes.
- **The parse path:** the corpus loads with no refusal and no check; the gate is section 2.4's run, repeated with the real check.
- **The cost:** `Symbols.docx`'s parse within 5% of today's. `tests/test_performance.py` is the place, and its thresholds are checkout-machine numbers.
- **Threads:** `scripts/threads.py`.
- **No Word check:** output is unchanged for valid code. The gates are `scripts/roundtrip.py` (50 of 50 parts), `scripts/compare_acceptance.py` (0 differ) and `codegen/generate.sh --check`.

## 6. Phases

| Phase | What | Effort |
|---|---|---|
| A | D1 to D6, D8 to D10: the table, the rule, the check in `ChildList`, the parse unchecked, the error and its messages, the engine's sites and the two tests, `unchecked()`, measured against the corpus, `Symbols.docx` and `scripts/threads.py` | 1 to 1½ days |
| B | D7: the validator, single fields and the cell rule, the out-of-range report (with CR-022's table, which exists since 2026-10-26: section 4.7), the save option, `conftest` turning it on | ½ to 1 day, and ¼ day for the report |
| C | The records: CHANGELOG (a "Changed" entry: a refusal of invalid input is a behaviour change), the README and `docs/pypi-readme.md`, CLAUDE.md, and the hand-off to docx4j through the portfolio (Q9) | ¼ day |

## 7. Risks

- **A legitimate write the suite doesn't exercise is refused.** Mitigations: the probe's run over the suite and the corpus, the exhaustive legal-placement test, and `unchecked()`.
- **The shape comparison is wrong at some depth.** Calling two classes one shape when they differ deeper can't produce a misplacement: the nested lists are checked against the item's own class. Calling interchangeable classes different is a refusal, and the suite and the corpus would show it.
- **Code that built invalid trees on purpose breaks.** `unchecked()` is the way out, and the CHANGELOG says so.
- **The first check of each (class, field) builds its table.** It's lazy, so a program pays only for the lists it touches.

## 8. Questions (decided 2026-10-26)

The user, 2026-10-26: "CR-020 Q1 to Q9 answered as recommended". So the check allows what the model allows (by class, or by element name and shape) and raises from the first release, the parse is unchecked through the `class_factory` (subject to Phase A's measurement), single-field and whole-list assignment are the validator's, the validator is `misplaced(root)` and `package.misplaced()` with `save(check=False)` by default and on in the test suite, the code is `target.misplaced`, every namespace is covered, the messages are hand-written for the common cases and derived for the rest, and the approach goes to docx4j through the portfolio after Phase A.

| Q | Question | Recommendation |
|---|---|---|
| Q1 | Allow only what the model allows (D2: by class, or by element name and shape), or refuse a list of known-wrong classes, as docx4j does? | Allow. Section 3 says why it's safe here, and it catches all six of docx4j-25's mistakes without listing them. |
| Q2 | An exception or a warning, and raising from the first release or after one release of warning? | An exception from the first release, with `unchecked()` as the way out: the misplacements it catches produce files Word can't open anyway. |
| Q3 | How is the parse unchecked: the `class_factory` (D4), a context variable around each parse call, or not at all? | The `class_factory`, unless Phase A's measurement shows it costs more than the check. |
| Q4 | Single-field and whole-list assignment: the validator only, or a `Child.__setattr__` that checks? | The validator only. A `__setattr__` runs on every field `__init__` sets, the parse included, and would need its own measurement. |
| Q5 | The validator's name and place, and the save option. | `misplaced(root)` in `traversal.py`, with `package.misplaced()`; `save(check=False)` by default, and on for the test suite. |
| Q6 | The error's code. | `target.misplaced`, beside `target.invalid`. |
| Q7 | Every namespace, or WordprocessingML only? | Every namespace: the table is derived, so it costs nothing more. |
| Q8 | The messages. | Hand-written for the six and Java's cases, derived for the rest (D6). |
| Q9 | Offer docx4j the derived table (from `@XmlElementRefs`) and section 2.3's one-shape finding after Phase A? | Yes, through the portfolio. JAXB has one class per type, so the per-holder classes are this model's; the derived validator is what docx4j-25 already suggests. |

## 9. Effort

2 to 3 days in all: A 1 to 1½, B ½ to 1, C ¼.
