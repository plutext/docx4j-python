# CR-020: Content refused where it does not belong --- a check on every typed list as code adds to it, derived from the model, and a validator for what gets past

**Status:** Proposed 2026-10-26, at the user's request ("draft it as a CR"). The proposal is the user's, put through the portfolio session the same day: refuse misplaced content at runtime (a `Tbl` appended to `R.content`), generated from the annotations. docx4j-25's assessment of the Java side came with it. It is measured here (section 2) with a probe kept in the session's scratchpad; nothing in the repository was changed to measure it. **Revised the same day** at the user's request ("also add the out-of-range report to CR-020's Phase B list"): the validator also reports an integer measurement beyond Word's range, with CR-022's limits table (section 4.7). **Decided 2026-10-26** (the user: "CR-020 Q1 to Q9 answered as recommended"; section 8). **Phase A implemented 2026-10-27** (section 10: `01fbe44`, `ebc3dcd`, `597097e`; Q3's measurement chose the context variable around each parse, 10.5); **its five open points decided as recommended 2026-10-27**, 10.14, nothing changed); **Phase B implemented 2026-10-28** (section 11: `740ea05`, `903a6d6`, `82800fd`, `5903d7d`; the validator `misplaced(root)` and `package.misplaced()`, `save(check=...)` with the suite saving checked, the out-of-range report held to Word 365's saves of docx4j CR-066 rounds 5 and 6; five points for the user in 11.9); **its five points decided 2026-10-28** (the user: "keep what's built"; 11.12, nothing changed); Phase C proposed. Three phases (section 6), 2 to 3 days.

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

**What Word does with the file (core-ts check 72, its CR-016 section 7.1, 2026-10-27; saves in core-ts `test/check72/`).** Of 18 files with one misplaced child each, by CR-064's rules, Word 365 refuses to open 11 ("Word experienced an error trying to open the file", no repair offered) and silently repairs 7, dropping the content of one (a string directly in a `w:r`). So section 1's "Word found unreadable content" is the better case; most of these files do not open at all, and one that does loses text without a word. Phase A refuses each of the 18 at add time (section 10.6).

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

**The other ports' reports, for alignment (2026-10-27).** Both counterparts landed their read side the day after this section was agreed, and Phase B holds its report to them:
- **docx4j-core-ts CR-017 phases P2 and P3** (the core-ts session, 2026-10-27): on load a warning on the part and `OpcPackage.warnings`, on save `lastSave.outOfRange` with the path from the part root (`/w:document/w:body[1]/w:p[3]/w:pPr[1]/w:ind[1]/@w:left`); fields `code`, `partName`, `line`/`column` or `path`, `element`, `attribute`, `value` (a number as held), `wordReads` (a number, or null for an unmeasured attribute such as `w:gridCol/@w:w` and for a negative `w:col/@w:w`), `storage` (objects-ts's strings: `int16`, `uint16, then 144..31681`, `int16, the sign the rule`, `int16, shared by firstLine (+) and hanging (-)`, `8 bits`, `5 bits`, `int32`, `15 bits, bit 15 clearing exact`, `unmeasured: changed outside -32768..65535`), `range` (null where none applies); two optional fields, `readAs` (`w:hanging` when a `w:firstLine` of 32768 is read in the shared field's other attribute) and `rule` (`exact`, `atLeast`, `auto`). Its oracle: docx4j CR-066 rounds 5 and 6 (660 cases), a warning for exactly the 432 cases Word's save changed, `wordReads` what Word wrote; CR-065 round 2's 64 decimals likewise.
- **docx4j CR-066 D6** (docx4j `338beec18`, its section 10): the inspector's finding per attribute occurrence, kind `value.out_of_range`, at the element carrying the attribute, in every WordprocessingML part, unselected `mc` branches and an embedded docx one level deep; its table `docx4j-core/src/main/resources/org/docx4j/model/text/word-integer-storage.tsv` (117 rows, seventeen storage words); the summary `attribute=w:tcW/@w:w; type=dxa; value=70000; storage=int16; range=-32768..32767; word_reads=4464`, the sibling token only where the key has one, `word_reads` a string in the attribute's own terms (`hanging 3959`, `atLeast 1`, `exact 32768`, `dropped`, `h=28809 hRule dropped`, a clamped number, `computed`, `-1`). Not checked: ids, counts, EMUs, `w:fitText`, `w:trHeight` under auto, a table width under auto or nil, `w:legacy`; Word's lexical reading (CR-021) comes first.
- **Read-side range semantics both ports share, where CR-022's write-side table narrows by choice:** 65,535 in `w:before`, `w:after` and the five unsigned page margins is in range (Word reads it as written and writes `-1`; a `-1` there reads as 65,535, not reported); `w:col/@w:space` `-1` is Word's sentinel, not reported, any other negative is; a row height under `hRule="auto"` is never reported. The report here reads those from the table's `word_writes` and `sentinel` and the unmeasured list, and does not take the seven rows' 65,534 cap as a reading limit.
- **The tables agree:** `integer-limits.json` was compared with `word-integer-storage.tsv` key by key on 2026-10-27; no measured row differs in range (the only differences are those write-side narrowings; docx4j's empty-context fallback rows are this table's explicit `auto` and `atLeast` rows; its type-level and unmeasured rows have no row here, by CR-022 Q3).

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

## 10. Phase A implementation notes (2026-10-27)

Phase A as built, D1 to D6 and D8 to D10 with section 5's tests, at the coordinator's brief after the decision of section 8. Phases B (the validator, section 4.7) and C (the records) are not begun: no CHANGELOG, README, `docs/pypi-readme.md`, CLAUDE.md or index change was made, and nothing in `../docx4j-portfolio/`. docx4j CR-064 (section 2's table, section 8's phase 1) and `ArrayListDocx4j.checkParent` (`docx4j-generated-objects`) were read for the mistakes and Java's voice, the fork (`ParserConfig`, `NodeParser.parse`, the serialiser's `convert_choice`) for behaviour; nothing outside this checkout changed. Section 2's probe code is gone, so every number below was re-derived from the model.

**Commits.**
- `01fbe44`: D1, D2, D5, D6: `MisplacedContentError`, and `docx4j_py/placement.py` (the table, the shape, the rule, the messages). Nothing calls it yet.
- `ebc3dcd`: D3, D4, D9: the check in `ChildList`, the parse unchecked, `unchecked()`; test (b), and test (c) with `run_items_of`.
- `597097e`: D8 and section 5: the engine's whole-list sites, `tests/test_misplaced.py`, the cost threshold in `tests/test_performance.py`, `tests/README.md`'s rows.
- the next commit: this section, and `test_every_namespace_is_checked` pinned at 39 modules.

### 10.1 What landed

| piece | where | what |
|---|---|---|
| the table (D1) | `docx4j_py/placement.py` (new, hand written, added to `codegen/clean.py`'s `KEEP`) | `entry(cls, field)`: an `Entry` per (class, list field), read from the class's `XmlMeta` through `runtime.context()` on first use and cached: the element classes, each element name's classes, whether a choice is of a simple type, whether the field is mixed, and a verdict cache per item class pre-filled with the exact classes; None (no table, no check) for a wildcard field, a field with no element choices, and a class xsdata cannot build. `shape_of(cls)`, memoised. `census()`, `differing_names()`, `model_classes()`, `holders_of(name)`, `property_holders_of(name)` |
| the rule (D2) | `placement.verdict`, `placement.written_as` | 10.3 |
| the check (D3) | `docx4j_py/child.py` | `ChildList.__slots__ = ("_field", "_owner")`, the `field` property, `ChildList(iterable, owner=..., field=...)`, `set_owner(owner, field=None)`; `Child.__post_init__`, `link_parents` and `link_children` pass the field name; `_check(owner, field, items)`, one dictionary lookup per item (`_verdict_tables`), the cold path `placement`'s; 10.4 |
| the parse (D4) | `docx4j_py/runtime.py` | `UncheckedXmlParser`, the fork's `XmlParser` whose `parse` (the one method `from_bytes`, `from_string` and `from_path` go through) runs under `unchecked()`; `parser()` builds it. 10.5 |
| the error (D5) | `docx4j_py/model/content/errors.py`, exported from `docx4j_py.model.content` | `MisplacedContentError(InvalidTargetError, TypeError)`, `default_code = "target.misplaced"`, with `holder`, `field`, `item` and `element`; `placement.refusal` imports it at the moment it raises |
| the messages (D6) | `placement.refusal` | 10.6 |
| the escape hatch (D9) | `docx4j_py.child.unchecked`, re-exported lazily from `docx4j_py.model.content` | a context manager over the context variable `child._CHECKING`; documented on `unchecked`, `ChildList`, `MisplacedContentError` and `UncheckedXmlParser` |
| the engine's sites (D8) | 16 sites in `model/content/`, `model/customxml/kinds.py`, `docx/`; `traversal.run_items_of` | 10.7 |
| scope (D10) | every generated namespace | 509 checked fields in 39 modules (10.2) |

### 10.2 The table as built

Over the model (`placement.census()`, held by `test_the_census_is_the_model_s`):

| | count |
|---|---|
| generated `Child` dataclasses | 1,969 |
| of them xsdata cannot build | 1: inkml's `MappingType`, whose annotations name `Math`, which does not resolve; its one list field is unchecked, and a parse of `inkml:mapping` fails the same way today |
| (class, list field) pairs of the rest | **537**, as section 2.2 |
| with a table, checked | **509**: 151 compound fields and 358 lists of one element (`TblGrid.grid_col`); 10 of them have a choice of a simple type (`c:dLbl`'s `c:separator`, `CTCopyrights.copyright`, VML's `x:ClientData`) |
| without, unchecked | **28**, as section 2.2: `oart` command 12, WordprocessingML 5, MathML 4, docProps 2, `mce` 2, DrawingML 2, InkML 1 |
| mixed list fields | 6, all wildcards, so rule 3's string never reaches a checked list in this model |
| element names in compound fields with more than one class | **71** |
| of one shape | **60** |
| of differing shapes | **11**, exactly section 2.3's: `w:sdt`, `w:customXml`, `w:ins`, `w:del`, `w:name`, `a:alphaMod`, `a:lum`, `a:tint`, `a:xfrm`, `a:effect`, chartex's `cx:lvl` |

**The shape** (`shape_of`) that gives section 2.3's eleven, and no other, is a class's attribute names and, for each field, the element names it holds and, one level further down, the attributes and fields-with-element-names of each of those elements' classes. Less gives other splits: the fields and element names alone, 65 and 6 (`a:lum`, `a:tint`, `w:name` and `cx:lvl`, whose classes differ only in their attributes, lost, and `w:sdt`); with the attributes but no second level, 61 and 10 (`w:sdt` lost: its four classes differ only in their `w:sdtContent`'s classes); two levels without the attributes, 64 and 7.

### 10.3 The rule, and what the serialiser writes

`verdict(cls, field, item)`, behind the cache:
1. **the item's class is one of the field's element classes, exactly** (pre-filled in the cache);
2. **its element name** (`traversal.element_name`) **is one of the field's choices, its class has the shape of one of that choice's classes, and the serialiser writes it under that name in this field** (`written_as`);
3. a `str` where the field is mixed, and **a value of a simple type (a `str` among them) where a choice is simple**.

A `DerivedElement` is judged by its qualified name and its value (that choice's class, or its shape with no type name of its own), each time; an `AnyElement` is refused, since a checked list has no wildcard; `None` is refused.

**Rule 2's third condition is this implementation's** (departure 2, 10.10). The fork's serialiser (`convert_choice`, then `convert_xsi_type`) writes an item whose class is none of a compound field's choices under the class's **own** name: its type's, else the `element_names` table's (`runtime._branch_element_names`, which knows the root elements' container classes, `PHyperlink` and `CtSdtContentRunHyperlink` among them), else its metadata's `qname`, which for an intermediate class is the class's name. And it writes a subclass of a choice class under the first choice it subclasses, which need not be its own element, with an `xsi:type` where the class has a type name of its own. So most same-shape classes are not interchangeable once written. Over every checked compound field, every class of a many-class element name that the field does not declare:

| | strangers | verdict |
|---|---|---|
| one shape, written under the field's element name with no `xsi:type` | 300 | accepted (the hyperlink of section 2.4 (a) among them) |
| one shape, written otherwise (`<MoveFrom1>` in `P.content`, `<Sat1>` in a colour, `MoveFromRangeStart1` as `w:bookmarkStart`) | 692 | refused |
| another shape | 448 | refused |

`written_as` follows the serialiser step for step, reading the shared context's metadata (where `XmlMeta.qname` is what the serialiser writes, the one place this package reads it); `test_an_accepted_stranger_is_saved_as_the_element_the_field_takes` renders every one of these 1,440 placements through `runtime.serializer()` and holds every accepted one to the element name with no `xsi:type`. **Rule 1 is exact** (departure 3): a subclass of a choice class is judged by rule 2, because the serialiser writes it under the first choice it subclasses: `MoveFromRangeStart1`, a `CTMoveBookmark` and so a `CTBookmark`, is written in `P.content` as `<w:bookmarkStart>`, its `w:author` and `w:date` with it, where `isinstance` would accept it as `w:moveFromRangeStart`.

### 10.4 Where it runs (D3)

- **A `ChildList` with an owner and a field name** checks, before it changes anything: `append`, `insert`, `extend`, `+=`, item and slice assignment, and what its constructor or `set_owner` is handed. Every item of an `extend`, a `+=` or a slice assignment is checked first, so a refused one leaves the list and every parent as they were; `test_a_refused_extend_or_slice_assignment_leaves_the_list_unchanged` holds it.
- **A list with no field name is not checked**: `ChildList(items)` and `ChildList(items, owner=x)`. Every list a model object holds gets its name from `__post_init__` (a `default_factory` list, or a plain one promoted), so `P(content=[Tbl()])`, `el.p(content=...)`, `p(tbl)` and `tc(r(...))` are refused at construction.
- **Binding a list to a new owner or field checks what it holds**: `set_owner(owner, field)` and so `link_parents`, which binds every `ChildList` it meets. So `run.content = ChildList([Tbl()])`, a whole-list assignment (the validator's, Q4), is refused at the next `link_parents`; a plain list assigned whole is not (`link_parents` parents its items and does not promote it). A list re-bound to its own owner and field is not checked again: it holds nothing new. `link_children`, the parser's hook, re-adopts such a list's items without `set_owner`'s comparisons.
- **A deep copy's relink is unchecked** (departure 5): the copy holds nothing its original did not, and a tree built under `unchecked()` copies as it is.
- **What the parser hands a constructor is a plain list**, not a `ChildList` as section 2.1 says: `XmlVar.factory` is `list`, and `__post_init__` promotes it, which is where the check would run.

### 10.5 The parse is unchecked (D4), and Q3's measurement

Q3 chose the `class_factory`, "unless Phase A's measurement shows it costs more than the check". `samples/Symbols.docx`'s `document.xml` (790,857 bytes, 24,758 objects, 4,065 of them holding a list), three ways in one process, interleaved, 20 rounds, three runs; and the code before the change in its own process after each (`out/wt/base`, at `eb8f2d5`):

| parse | best, ms (three runs) | median, ms | over the cheapest |
|---|---|---|---|
| a context variable set around each parse call (`UncheckedXmlParser`) | 194.7, 192.8, 193.6 | 196.1, 194.4, 194.4 | --- |
| the check on during the parse (the fork's `XmlParser` on `parser_config()`) | 196.0, 196.1, 196.7 | 197.7, 196.9, 197.5 | +1.3 to +3.3 ms best (+0.7 to +1.7%): **the check itself** |
| a `class_factory` that builds each object with the check off (only for a class with a list field, the variable read and set through bound locals) | 198.1, 197.4, 197.8 | 199.9, 198.6, 199.0 | +3.4 to +4.6 ms best (+1.8 to +2.4%) |
| before CR-020 (`eb8f2d5`) | 190.9, 191.4, 190.0 | 192.2, 192.7, 191.0 | |

Earlier runs, while the factory was being tuned, gave the same order every time: the factory +3.6 to +6.2 ms over the context variable, the check +2.7 to +4.1. **The `class_factory` costs more than the check**, so by Q3's own condition it is not the choice: its context-variable set and reset per object (and a function call per object) outweigh a dictionary lookup per list item. Of the two alternatives Q3 names, "not at all" would leave a parse checked, which D4 exists to prevent, so **the parse is unchecked by a context variable set around each parse call** (departure 1). Section 4.4 expected that to need every call site (`XmlPart` unmarshalling, `wml(...)`, `insert_xml`, `insert_ooxml`); it needs one: every parse in the package goes through `runtime.parser()`, which builds `UncheckedXmlParser`, whose `parse` is the method every `from_*` calls. `test_every_parser_in_the_package_is_runtime_parser_s` holds that no module of `docx4j_py` builds another parser. A parser a test or a script builds on `parser_config()` itself parses with the check on, which is correct (parsed content conforms, section 2.4) and only dearer; `scripts/threads.py` does that, on eight threads, unaffected.

The variable is per thread and per asyncio task, so `unchecked()` in one thread leaves another checked (`test_unchecked_is_per_thread`), and it is reset when the parse returns or raises (`test_the_check_is_on_again_after_a_parse_and_after_one_that_raises`). No check runs while every part of the 16 samples is unmarshalled (`test_every_sample_loads_with_no_check`).

### 10.6 The messages (D6), and docx4j CR-064's table

The message names the holder's field and the item, class and element: "`R.content` can't hold a `Tbl` (`w:tbl`)", "`P.content` can't hold an `RT` (`w:t`)", "`P.content` can't hold a `str`". The hint is chosen by what the list holds (runs, blocks, rows, cells, or a run's own content), so one rule covers every container CR-064 lists, and the hint says where the item goes, in Java's voice ("You can't X; do Y" becomes "X goes in Y"):

| CR-064 section 2 | Java's message (`checkParent`) | the hint here |
|---|---|---|
| 1. a run, or bare Text, in Body or a Tc (and the other block containers) | "add it to a paragraph (P) and add that"; Text: "add it to a run (R) in a paragraph (P)" | "a run goes in a paragraph: `p(r(...))`"; "text goes in a run, in a paragraph: `p(r("..."))`" |
| 2. a Drawing straight in a P | "add it to a run (R) and add that" | "a drawing goes in a run: `el.r(content=[drawing])`" (in a block container: "..., in a paragraph") |
| 3. a Tbl, or Text, in a P | "it should be a sibling of the paragraph"; "add it to a run (R)" | "a table is the paragraph's sibling: put it in the body, a table cell or a block-level content control"; "text goes in a run: `r("...")`" |
| 4. a String in a content list | "put it in a Text, in a run (R), and add the run" | "a string is not an element: wrap it, `r("...")`" (in a run `t("...")`, at block level `p("...")`, in a table or row "text goes in a cell, `tc("...")`") |
| 5. a property object added instead of set | "use setPPr" (and setRPr, setTblPr, setTrPr, setTcPr; Body.setSectPr "or PPr.setSectPr on the paragraph that ends the section") | "`pPr` is a property, not content: set `p.p_pr`", from the holder's own field, and so for `rPr`, `tblPr`, `trPr`, `tcPr`; "`sectPr` is ...: set `body.sect_pr`, or `p_pr.sect_pr` on the paragraph that ends the section"; held elsewhere, "set it as `R.r_pr`, ..." |
| 6. a P in a Tr, a Tc in a Tbl, a Tbl in a Tbl (and a P or R in a Tbl, a Tc or Tr at block level) | "add it to a cell (Tc)", "add it to a row (Tr)", "a nested table goes in a cell (Tc)", "a cell goes in a row (Tr), in a table (Tbl)" | "a paragraph goes in a cell: `el.tc(content=[p(...)])`", "a table holds rows: `el.tr(content=[el.tc(...)])`", "a nested table goes in a table cell (`Tc`), not straight in a table or a row", "a cell goes in a row, in a table", "a row goes in a table" |
| 7. a Br or Tab straight in a P | "add it to a run (R)" | "a break goes in a run: `r(br())`", "a tab goes in a run: `r(tab())`" |
| 8. the other run-level containers given a Tbl, P or Text | as P's | as P's: their lists hold runs |
| 9. a P in run-level sdt content, an R in block-level | as P's, as Body's | "a paragraph is a block, not inline content: ..."; "a run goes in a paragraph" |
| Java: a P in a P or an R | "You can't nest a paragraph inside ..." | "a paragraph is a block, not inline content: put it in the body, a table cell or a block-level content control" |
| Java: an R in an R | "You can't nest a run inside a run; add it to the paragraph" | "a run can't hold a run: put it in the paragraph, beside this one" |
| Java: SdtContent in SdtContent | "nest an Sdt (SdtBlock or SdtRun) that holds it" | "a `w:sdtContent` is a content control's own part, not content: nest the control that holds it (`sdt(...)` builds both)" |
| the form of a control (section 4.6) | --- | "this is a block-level content control (`SdtBlock`); in a paragraph use the run form, `el.sdt_run(...)` or `sdt(..., form="run")`", for all four forms, and the same for `w:customXml`'s four (`el.custom_xml_run(...)`) |

Java's `JAXBElement` unwrap (CR-064 D1) is the `DerivedElement` rule here. Java names the container in prose ("to the body"); the message here names the field (`Body.content`). **Derived** for everything else, from a reverse index over every class built on the first refusal that needs it (0.3 s, once per process): "`w:bookmarkStart` goes in: `Body.content`, `Tc.content`, `P.content`, `Tbl.content`, `Tr.content`, ...", well-known WordprocessingML holders first; an element of the field in another form, "`CTR.content` takes `w:ins` in another form, an `Ins1`; build that one, `el.ins_math(...)`" (a math run's `w:ins` is the one of the three `w:ins` shapes the other holders do not take); a property element (`*Pr`) held only by single-valued fields elsewhere, "set it as ..."; an item with no element name, "`P.content` takes `mc:AlternateContent`, `m:oMath`, `m:oMathPara`, `w:bdo`, `w:bookmarkEnd` and 30 more". `test_misplaced.py` holds each of section 2.7's cases, each of Java's and each of CR-064's rows with its hint.

### 10.7 The engine's own sites (D8)

- **Whole-list reassignment.** Section 2.6 named four; there are sixteen `ChildList(..., owner=x)` sites, and every one now names its field: `model/content/comments.py` (`_append`, `Comment.content`'s setter, `_ensure_para_id`), `controls.py` (`properties`), `lists.py` (four: `w:num`'s overrides, `w:numbering`'s lists, an override, a level), `styles.py`, `placeholders.py`, `body.py` (`Body.content`'s promotion), `model/customxml/kinds.py` (the list entries), and the facade's `docx/styles.py`, `docx/table.py` (`w:trPr`) and `docx/text.py` (`w:tabs`). The glossary's styles list (`placeholders.ensure_glossary_styles`), assigned unowned, is now owned. The constructor sites (`P(content=ChildList([...]))` and the like, in `comments.py`, `controls.py`, the builders, the importer and elsewhere) needed nothing: `__post_init__` binds them, and they are checked. `test_every_owned_child_list_the_engine_builds_names_its_field` holds that no `ChildList(..., owner=...)` in the package lacks `field=`, and `test_every_field_name_the_engine_gives_a_list_is_a_list_field` that each literal is a list field.
- **The hyperlink write** (`hyperlinks._wrap`, a `PHyperlink` put in `CTSdtContentRun.content`) needed no change: rule 2 accepts it, and `test_the_content_api_s_hyperlink_in_a_control_goes_through_the_check` makes it through the content API and reads `<w:hyperlink>` in the saved `w:sdtContent`.
- **Test (b)**, `test_builders_phase_a.py::test_the_form_can_be_overridden`, overrides the form of content that infers the block form and is run content too (a bookmark, then a run: refused in the block form, built in the run form); the paragraph in a run-level control is `test_a_form_its_content_cannot_take_is_refused`, a `MisplacedContentError` from the `sdt` builder through the check.
- **Test (c)** takes `SdtContentBlock` out of the run holders. **`run_items_of`** answers None for a block-level `w:sdtContent` and for a `w:sdt` holding one, **and for the row and cell forms too** (departure 4), by whether the content's list takes `w:r`. Its callers: `ContentControl.content`, `control_view_of` and `parent_control_of` take `block_children_of` first, which answers for every form; the run-level callers (the hyperlinks, comments, tracking, the text model, the split and the controls' run writes) pass run holders. None needed the other lists, and the suites pass. `test_run_items_of_a_control_reaches_through_to_its_content` now expects the list for the run form and None for the other three.

### 10.8 Tests and the numbers

- **`tests/test_misplaced.py`, 110 tests:** the error (a `TypeError`, an `InvalidTargetError`, a `ContentError`, a `Docx4JException`; `to_dict()`; its attributes); section 2.7's six mistakes as 13 cases and Java's as 11, each with its hint; CR-064's nine rows and phase 1's sets as 41 cases, and its accepted children (`mc:AlternateContent` in six holders, range markup, `w:altChunk`); the forms; the derived hints; the census and the eleven names; every namespace (a checked field in each of the 39 modules that have one refusing a `Tbl`); the table built lazily (a fresh process has none until a list is checked); **every legal placement**: for each of the 509 fields, an instance of each choice class, 3,480 built by an `el` factory and 155 (77 classes no factory builds: VML's `x:ClientData` children, inkml's types, chartex's intermediate classes) by the class, one by one, through the constructor and through `extend`, none refused; the shape rule (the hyperlink accepted and saved, `SdtBlock` in `P.content` refused, every differing name refused in another holder, every accepted stranger saved as its field's element, 10.3); the mutations and their atomicity; `unchecked()`; the parse; the engine's sites.
- **`tests/test_builders_phase_a.py`:** test (b) rewritten and one test added; test (c)'s `SdtContentBlock` case dropped; the control-content test changed (10.7).
- **`tests/test_performance.py`:** `test_the_list_check_leaves_the_parse_within_five_percent` (10.9).

### 10.9 The cost

Section 5: `Symbols.docx`'s parse within 5% of today's. Measured interleaved, `eb8f2d5` (`out/wt/base`) and the code after, each in its own process, the engine's parser (`runtime.parser()`, as `XmlPart._unmarshal` parses), 20 parses each: before 189.9 to 194.8 ms best over eleven runs, after 192.0 to 194.4 ms best over the last eight (at `597097e`'s code); inside pytest, the best of nine, before 191.9 to 196.6 ms and after 193.2 to 194.4 ms. **About 1% over**, the check's machinery in the lists the parse builds (a context-variable read per promoted list, the owner and field comparisons), all of it in the noise. The test holds the best of nine below 192.0 ms + 5% (201.6 ms), a checkout-machine number, in the `slow` module CI skips. The check itself, were it run during a parse, is a dictionary lookup per list item: +0.7 to +1.7% (10.5). A deep copy is not checked; `link_parents` over a parsed tree checks nothing (`test_link_parents_stays_within_its_budget`: 8.0% of the parse, 8.1% at `eb8f2d5`).

### 10.10 Departures, with why

1. **The parse is unchecked by a context variable around each parse call, not by the `class_factory`** (D4, Q3). Q3 made the `class_factory` conditional on its cost, and it costs more than the check (10.5). The alternative is implemented once, in `runtime.parser()`'s `UncheckedXmlParser`, not at each call site.
2. **Rule 2 has a third condition**: the serialiser must write the item under the element name it matched. Without it, 692 same-shape placements the serialiser writes under another name (`<MoveFrom1>`) would be accepted, and the save would be the "unreadable content" this CR exists to prevent (10.3). The hyperlink case passes it.
3. **Rule 1 is exact class membership**, not `isinstance` (section 2.5's probe took `isinstance` on a miss): a subclass instance is written under the first choice it subclasses, with an `xsi:type` where it has a type name (10.3). It goes by rule 2.
4. **`run_items_of` answers None for row and cell `w:sdtContent` as well as block**: they keep no runs either ("as for any holder that keeps no runs", section 4.8), and no caller wanted their lists from it.
5. **A deep copy's relink runs unchecked**, the one use of the context variable outside a parse, which D9 rules out for the engine. A copy holds nothing its original did not, so the check could only refuse a tree built under `unchecked()`, which would then not copy. For the user to confirm (10.12).
6. **Rule 3 is widened to values of a simple type** where a choice is simple: section 4.1's "strings only where the var is mixed" would refuse the `str` of `CTCopyrights.copyright` (a `list[str]` element field) and a kept text in an enumeration's list (`x:ClientData`); the 6 mixed fields are wildcards, so unchecked.
7. **`MisplacedContentError` carries `holder`, `field`, `item` and `element`**, beyond the code, message and hint section 4.5 names; `to_dict()` is unchanged.
8. **`unchecked` is exported from `docx4j_py.child` and `docx4j_py.model.content`**, not from `docx4j_py` or `docx4j_py.wml`, whose `__init__.py`s are generated.

### 10.11 Found, and not done

1. **The serialiser names a stranger by its class** (10.3): upstream xsdata's `convert_xsi_type` for a compound field writes `meta.target_qname`, which is None for the intermediate classes, then the class's `qname`. A fork change writing a same-shape class under the element its field matched would let rule 2 accept the 692; this CR changes nothing in the fork. Recorded by the fork session as `docs/UPSTREAM.md` report 28 on 2026-10-27 (the user: worth recording, not a blocker): to report upstream, no fork change.
2. **inkml's `MappingType`** cannot be built by xsdata (its annotations name `Math`; 10.2). Its list field is unchecked, and a parse of `inkml:mapping` fails today regardless.
3. **77 model classes have no `el` factory**: VML's `x:ClientData` children, inkml's types, chartex's intermediate classes (`test_every_legal_placement_goes_in` counts them).
4. **A plain list assigned whole, and every single-valued field, are still unchecked** (section 4.11, Q4): Phase B's validator.
5. **The portfolio's `python/CR-020.A` entry** wants `done` (with this section as its note), sent to the docx4j-portfolio session; the `tasks.yaml` rule (CLAUDE.md) is not this session's to apply.
6. **Section 2.1** says the parser hands a list field to the constructor as a `ChildList`; it hands a plain `list` (`XmlVar.factory`), which `__post_init__` promotes (10.4). Nothing depends on it.

### 10.12 For the user to decide

1. **Q3's outcome** (departure 1): the context variable around each parse, as Q3's condition gives; or the `class_factory` regardless, at about 1 to 2% more per parse.
2. **The deep copy's unchecked relink** (departure 5), or a copy checked like any other bind (one lookup per item; a tree built under `unchecked()` would then copy only under `unchecked()`).
3. **`run_items_of` for row and cell content** (departure 4).
4. **Rule 2's third condition** (departure 2): kept, or a fork change (10.11 item 1) so that the 692 same-shape strangers are written as their element and accepted.
5. Whether `unchecked` should also be reachable from `docx4j_py.wml` (a change to the generator's re-export footer).

### 10.13 The gates

| gate | before (`eb8f2d5`) | after (`597097e`) |
|---|---|---|
| the full suite (`.venv-fork/bin/python -m pytest`) | 7,654 passed, 8 skipped (4 min 45 s) | **7,765 passed, 8 skipped** (4 min 38 s) |
| the fast suite (`-m "not slow"`) | 6,273 passed, 8 skipped, 1,381 deselected (3 min 8 s) | **6,383 passed, 8 skipped**, 1,382 deselected (3 min 10 s) |
| the new tests | --- | 110 in `test_misplaced.py`, 1 in `test_performance.py` (`slow`); `test_builders_phase_a.py` one added, one case dropped |
| `test_parser_configs.py`, `openpackaging/test_threads_and_import.py` | pass | pass (in both suites) |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | 50 of 50 canonically identical, 0 differences, 0 skipped, 34 `boolean-spelling` | **the same** |
| acceptance (`scripts/compare_acceptance.py`, `out/acceptance/` at `eb8f2d5` in `out/wt/base` against the checkout's) | --- | **17 artefacts, 246 entries, 0 differ** beyond the dates |
| `codegen/generate.sh --check` (a worktree at `597097e`) | --- | two generations byte identical, 309 modules; nothing generated differs from the commit, and `placement.py` survives `clean.py` |
| `scripts/threads.py` | --- (passed at Phase A) | **passes** (eight threads, three modes, no document lost or altered) |
| `ebc3dcd` on its own (a worktree under `out/wt/`) | --- | the fast suite: 5,668 passed, 51 skipped (the worktree's skips: the sibling checkouts' goldens are not beside it), 0 failed |
| `Symbols.docx`'s parse, best | 189.9 to 194.8 ms | 192.0 to 194.4 ms (10.9) |

No Word check (section 5): output is unchanged for valid code, and the acceptance documents are identical.

### 10.14 The decisions (2026-10-27)

The user, on 10.12: "decisions as recommended (subject to what the fork says)". Each keeps what Phase A built, so no
code, test or generated file changed and no gate moved from 10.13's:

1. **Q3's outcome stands**: the parse is unchecked by the context variable around each parse (`UncheckedXmlParser`),
   not by the `class_factory`, which 10.5 measured at 1 to 2% more per parse for the same effect (departure 1).
2. **A deep copy's relink stays unchecked** (departure 5): a copy holds nothing its original did not, and a tree built
   under `unchecked()` must copy without it.
3. **`run_items_of` stays None for row and cell `w:sdtContent`** as for block (departure 4).
4. **Rule 2's third condition stays**: a same-shape stranger the serialiser would write under another name is refused
   (departure 2; the 692 of 10.3). This one is **subject to the fork session's answer**: 10.11 item 1 went to it on
   2026-10-27, and it is `docs/UPSTREAM.md` report 28 here, with no fork change. Should the fork later write a same-shape class under the element its field
   matched, `written_as` follows the serialiser (10.3), so the rule would accept those placements with no change here
   beyond the test's expected counts.
5. **`unchecked` is not re-exported from `docx4j_py.wml`**: it stays where it is, `docx4j_py.child` and
   `docx4j_py.model.content`, and the generator's footer is untouched.

## 11. Phase B implementation notes (2026-10-28)

Phase B as built, D7 (section 4.7) with Q5's save option and `conftest`, at the coordinator's brief after the user's
"start CR-020 Phase B" (2026-10-28). Phase C (the records) is not begun: no CHANGELOG, README, `docs/pypi-readme.md`,
CLAUDE.md or index change was made, and nothing in `../docx4j-portfolio/`. Read for behaviour, never edited: docx4j-core-ts
at `b7c2680` (CR-017 section 7, `src/model/wordIntegers.mts`, `XmlPart.valueOutOfRange`, `test/integer-attributes.test.mjs`),
docx4j-generated-objects-ts at `8ec65b3` (`src/helpers/wml.mts`: `WORD_INTEGER_STORAGE`, `wordIntegerReading`, `storageOf`;
`src/index.mts`'s `outOfRangeOf`) and docx4j at `fa3011a41` (`word-integer-storage.tsv`, CR-066 section 10).

**Commits.**
- `740ea05`: the validator, `misplaced(root)`, the findings, the wrap model, D2's `xsd:anyType` substitution (11.7
  departure 1), `tests/test_validator.py` without the save's check, the private reader, the cost guard.
- `903a6d6`: `save(check=...)`, `SaveCheckError`, `XmlPart.source_tree()`, `SaveReport.findings`, `tests/conftest.py`, the one
  opt-out, the save's tests.
- `82800fd`: departure 1 narrowed to an element declared with no type of its own (a named type with wildcard content is not
  one), and its test.
- `5903d7d`: the two save-check tests that assumed the suite's default hold with `DOCX4J_PY_SAVE_CHECK=0` too.
- the next commit: this section, the Status line, `tests/README.md`'s rows.

### 11.1 What landed

| piece | where | what |
|---|---|---|
| the validator | `docx4j_py/validator.py` (new, hand written, added to `codegen/clean.py`'s `KEEP`) | `misplaced(root, *, unmarshalled_only=False)` over a model object, an `XmlPart` or a package; `findings_in(root, *, part_name=None, root_name=None)` (a tree), `part_findings(part, tree=None)` (a part, with its name and its addresses); one pre-order walk (`_Walker`), the three checks of 11.2 on each node; `check_for_save(package)` (11.4) |
| the public names | `docx4j_py.traversal.misplaced(root, *, unmarshalled_only=False)`; `OpcPackage.misplaced(*, unmarshalled_only=False)` | Q5's names; both are the validator's `misplaced` |
| the findings | `docx4j_py.validator` | `Finding` (frozen, slotted: `message`, `part_name`, `path`, `address`, and `node`, the object the finding is about, neither in `to_dict()` nor compared; `code` a class constant) and its three kinds, `MisplacedContent`, `CellWithoutEndParagraph`, `ValueOutOfRange` (11.2); `to_dict()` leaves None out, as `model/content/reports.py`'s do, but for `ValueOutOfRange`'s `word_reads` and `range`, where None means something (core-ts's `null`) |
| the wrap model | `docx4j_py/model/word_integers.py` (new; inside `model`, listed in `KEEP` too) | `Storage` (the objects package's `WordIntegerStorage`, measured entries), `STORAGE` (one per attribute path of the limits table: 77 for its 82 rows by context), `Storage.token` (`storageOf`, character for character), `reading(storage, attribute, value, rule)` → `Reading(value, rule, attribute, kept, storage)` (`wordIntegerReading`, step for step), `kept_range` (`rangeOf`, core-ts's `keptRange`), `read_range(row)` (a row's range as Word reads, 11.3), `wrap` |
| the save's check | `docx4j_py/openpackaging/save.py`, `packages/opc_package.py`, `normalise.py` | `save(target, *, normalise=None, check=None)`, `save_to(sink, *, normalise=None, check=None)`, `save_package(package, sink, *, normalise=None, check=None)`; `checking(enabled=True)`, a context manager over a context variable (per thread, as `unchecked()` is), whose default is False; `saves_check(check)`; `check_before_save(package, check)`; `SaveReport.checked`, `SaveReport.findings`, `SaveReport.out_of_range` (core-ts's `lastSave.outOfRange`) |
| the error | `docx4j_py/model/content/errors.py`, exported from `docx4j_py.model.content` | `SaveCheckError(ContentError)`: `code` the first finding's (`target.misplaced`, `cell.no_end_paragraph`, `value.out_of_range`), its message and hint, `findings` every finding the session introduced; `SaveCheckError.of(findings)` |
| the source's tree | `docx4j_py/openpackaging/parts/xml_part.py` | `XmlPart.source_tree()`: a fresh tree of the part's source bytes under the parse settings a parse now uses, held by nothing, the part untouched; `XmlPart._parse(data, settings)`, the parse `_unmarshal` makes, factored out for it |
| D2 widened | `docx4j_py/placement.py` | `_any_type(cls)` and `_any_type_substitution`; `_derived_allowed` accepts a `DerivedElement` naming its value's own type where the choice is an element declared with no type (11.7 departure 1) |
| the suite | `tests/conftest.py` (new) | a session fixture, `checking(True)`, unless `DOCX4J_PY_SAVE_CHECK=0` |
| the tests | `tests/test_validator.py` (new, 346); `tests/test_performance.py` (one, `slow`); `tests/openpackaging/test_private_readers.py` (the validator a reader, its caller listed); `tests/content/test_table_rules.py` (the opt-out) | 11.8 |

### 11.2 The findings

| kind | code | fields beyond the four of every finding | message |
|---|---|---|---|
| `MisplacedContent` | `target.misplaced` | `holder` (the holder's class, `"R"`), `field` (`"content"`, `"p_pr"`), `index` (the item's position; None for a single-valued field), `item` (its class, `"Tbl"`, `"str"`), `element` (prefixed, `"w:tbl"`), `hint` | the add-time error's, from `placement.refusal`: "`R.content` can't hold a `Tbl` (`w:tbl`)", with its hint, so one search finds both |
| `CellWithoutEndParagraph` | `cell.no_end_paragraph`, proposed here (11.9 Q3): `table_rules` has no code, and neither counterpart reports the rule | `ends_with` (`"w:tbl"`, or None for a cell holding no block), `hint` (Word's repair, and `table_rules.end_with_paragraph(tc)`) | "a table cell must end with a paragraph, as Word reads it; this one ends with a table (`w:tbl`)" |
| `ValueOutOfRange` | `value.out_of_range`, CR-022's `MeasureRangeError`'s | `element`, `attribute`, `value`, `word_reads`, `storage`, `range`, `read_as`, `rule` (11.3) | core-ts's form: "`/word/document.xml: /w:document/w:body[1]/w:p[3]/w:pPr[1]/w:ind[1]/@w:left=7205001 is past what Word 365 stores there (int16, -32768 to 32767); Word reads it as -3959`"; `describeReading`'s `atLeast 1`, `32768 in w:hanging`, `a width Word computes` |

**Misplaced content.** Every item of every typed list and every single-valued element field's value is held to D2 by
`placement.verdict`, the add-time rule itself, so the two never disagree: one dictionary lookup per item
(`Entry.verdicts`), the cold path only on a miss. `placement.entry` already reads a single element field (`var.is_element`)
as a one-choice table, so a single field needed nothing new. That covers what the add-time check cannot see (section 4.11,
10.11 item 4): a plain list assigned whole (`run.content = [Tbl()]`, never promoted to a `ChildList`), a single field
assigned (`p.p_pr = R()`), `object.__setattr__`, a tree built under `unchecked()`. A misplaced item's step in the path is
named by its own element (`w:tbl`), as core-ts names it, not as the serialiser would write it: a `Tbl` in `R.content` is
written `<w:CT_Tbl/>`, its type's name (10.3, `docs/UPSTREAM.md` report 28). A string's path is its holder's, as core-ts
gives one. A misplaced item is walked into, so what is wrong inside it is reported too. The 28 unchecked fields and what
sits in a wildcard are not checked (section 4.11); the typed objects inside a wildcard (a kept `mc` branch's) are walked.

**The cell rule.** Each `w:tc` the walk meets, `mce="all"` (every branch is written back), held to
`table_rules.paragraph_end_of`: past range markup, through block-level controls and custom XML, through the branch of a kept
`mc:AlternateContent` Word reads (CR-016 Phase A).

**Paths and addresses.** The path is from the part's root, each step a prefixed name and its 1-based position among its
siblings of that name, an attribute's ending with it (`/w:document/w:body[1]/w:tbl[2]/w:tr[1]/w:tc[1]`,
`/w:settings/w:defaultTabStop[1]/@w:val`), core-ts's form; for a tree walked on its own, from that tree's root. The walk
counts a node's children in field order, which is the order the serialiser writes them. The address is the content API's:
the innermost block holding the finding (`address_of`'s rule: its `w14:paraId` where it has one, else its ordinal,
`body/2/0/0`), computed down the walk's own trail, so a tree whose parent pointers are stale still gets one; only in a part
that has a body (main document, header, footer, notes, comments), and with no `body_of`, which would unmarshal the part.

### 11.3 The out-of-range report

**What it reads.** CR-022's table (`integer_limit(attribute, context)`), by the attribute's path: the element the walk names
the node, the attribute's qualified name, the holder where the row names one (`w:rPr/w:sz`), the row's context from the
sibling attribute on the same node (an enumeration read as Word reads one, `values.in_enumeration`: an invalid value is
unset). The value is Word's lexical reading first (CR-021, `values.integer`): `"70000.5"` is 70000 and reported, `"720.5"`
is 720 and in range, `"abc"` is not a number Word reads and is CR-021's, not this report's. Each (path, context) is looked up
once and cached (11.6).

**Read-side semantics, section 4.7's.** The range a value is held to is the row's as Word **reads** (`read_range`): the
row's own, with the values its `word_writes` names put back, so 65,535 in `w:before`, `w:after` and the five unsigned page
margins is in range, the write side's 65,534 cap no read limit; a value equal to a `word_writes` spelling (`-1` there, which
Word reads as 65,535) is not reported, nor a row's `sentinel` (`w:col/@w:space`'s -1), while any other negative space is; an
attribute the table lists as unmeasured (`w:gridCol`, `w:wBefore`, `w:wAfter`, `w:legacy`'s two, `w:fitText`, a row height
under `hRule="auto"`) is not checked, nor one with no row under its context (a table width under `auto` or `nil`, or with
no `w:type`: CR-022's rows are `dxa` and `pct`).

**The wrap model** (`model/word_integers.py`), the table's missing half (CR-022 10.10 item 5). The table's `storage` is
prose and gives no reading, so each attribute path has a `Storage`: the objects package's entry for the class that path
binds (`CTBorder` for every border edge and `w:bdr`, `TblWidth` for the widths and margins, `HpsMeasure` for `w:sz`,
`w:szCs` and `w:kern`, `CTTwipsMeasure` for the two settings), and `reading` is `wordIntegerReading` in Python's spelling.
`test_the_wrap_model_keeps_exactly_the_row_s_range_as_read` holds the two halves to each other for every one of the 82
rows: `kept_range` (core-ts's `keptRange`) equals the row's range as read, and `reading().kept` is True exactly inside it and
at Word's own spellings, at up to 46 values each, the boundaries and one past them included. So the table decides what is
reported and the model only what Word reads, and a disagreement fails a test.

**Alignment with the other ports.**

| | docx4j-core-ts (CR-017 P2, P3) | here |
|---|---|---|
| code | `value.out_of_range` | the same |
| fields | `element`, `attribute`, `value`, `wordReads`, `storage`, `range`, `readAs`, `rule`, `partName`, `path` or `line`/`column` | the same in snake_case (`word_reads`, `read_as`, `part_name`), always `path`, never a line (the typed tree keeps none); `address` beside it |
| `value` | a number as held | an `int` where it binds as one, else the kept text (section 4.7) |
| `storage` | the objects package's strings | **the same strings** (`test_the_storage_strings_are_core_ts_s`): `int16`, `uint16`, `int32`, `8 bits`, `5 bits`, `int16, shared by firstLine (+) and hanging (-)`, `int16, the sign the rule`, `int16, then ±31680`, `uint16, then 144..31681`, `uint16, then at least 144`, `int16, a negative computed`, `int16, a negative written -1`, `15 bits, bit 15 clearing exact`; a line under `auto` `int16` |
| `range` | `keptRange` | the row's as read, equal to `kept_range` by test |
| `rule` | wherever the reading has one (a line under `auto` too) | the same |
| when | every load (a warning on the part) and every save (`lastSave.outOfRange`) | on demand (`misplaced`), and on a save with `check=True` (`last_save.findings`, `last_save.out_of_range`) (11.7 departure 3) |
| what is checked | the objects table, by class: also the unmeasured attributes outside the certain range -32768 to 65535 (`unmeasured: changed outside -32768..65535`, `wordReads` null), a `TblWidth` under any type (`w:wBefore`, `w:wAfter`, a width under `auto`), every `HpsMeasure` and `CTTwipsMeasure` (`w:hps`, `w:drawingGrid*`) | CR-022's table, by attribute path: nothing it lists as unmeasured or has no row for (11.7 departure 2) |

docx4j's inspector (CR-066 D6) spells `word_reads` as a string in the attribute's terms (`hanging 3959`, `atLeast 1`) and
checks a table width with no `w:type` as `int16`; this report follows core-ts's number with `read_as` and `rule`, as
section 4.7 agreed.

### 11.4 The save's check, and what it does with a finding

**The decision** (the brief's recommendation, kept: the evidence agrees with it):
- **Misplaced content raises, always.** A parse never produces it (section 2.1; 11.5's corpus holds none, once departure 1
  is made), so it is code's, and the file it would write is one Word refuses to open (core-ts check 72: 11 of 18) or
  repairs. It is found before anything is written.
- **A cell that does not end with a paragraph, or an integer measurement past what Word stores, raises where the session
  introduced it, and is carried where the source already held it.** Carried means saved as it is, never repaired
  (CLAUDE.md's "nothing is repaired on load or save"; CR-016 D4's "until an edit touches that cell"), and listed on
  `last_save.findings` (`last_save.out_of_range` the out-of-range ones, core-ts's `lastSave.outOfRange`). A file from
  elsewhere with a broken cell or a wrapped indent can be loaded, edited elsewhere and saved with the check on.
- **Introduced** is told from **carried** by a fresh parse of the part's source (`XmlPart.source_tree()`), made only for a
  part that has such a finding, so a save with nothing to report parses nothing more. A finding is matched to one of the
  source's by what it is, never by where, since an edit moves things: an out-of-range value by its element, attribute and
  value; a cell by what it ends with and its text. Matching is a multiset, so a copy of a carried value is the session's
  (`test_a_value_the_source_held_is_carried_and_one_the_session_set_raises`), and so is a cell the session empties.
- **The order.** The check runs first, before the save hooks (CR-019's glossary) and before the target is opened, in
  `OpcPackage.save` for a path or a stream, so a refusal leaves the target as it was (`ZipPartSink` truncates on
  construction) and writes nothing. `save_package` and `save_to` run it before anything is put in the sink.
- **What it covers.** The parts the save re-marshals: those unmarshalled. A part never unmarshalled is written byte for byte,
  as read; `package.misplaced()` reads it privately, which marks nothing for re-marshalling, so a `misplaced()` before a
  save does not widen the save.

**`SaveCheckError`** carries every introduced finding (`findings`), the first's code as its own, so a caller's
`except ContentError` and a tool's `to_dict()` read the condition. One class rather than the condition's own
(`MisplacedContentError`, `MeasureRangeError`): a save can introduce several kinds at once.

**The suite** (Q5): `tests/conftest.py` runs every save under `checking(True)`. One test opts out at its save:
`content/test_table_rules.py::test_d4_a_file_whose_cell_ends_with_a_table_is_left_as_it_is_until_an_edit_touches_that_cell`,
which breaks a cell behind the API's back to stand for a file from elsewhere (`save(check=False)`); its next save, of the
file reloaded, carries the cell. No other test of the suite introduced a finding: no engine edit the suite exercises
leaves misplaced content, a cell without its paragraph or an out-of-range value behind. `test_validator.py`'s save tests use
`check=` and `checking(False)` because they test the option.

### 11.5 The oracle and the corpus

**docx4j CR-066 rounds 5 and 6** (`tests/fixtures/integer-limits/`), as core-ts counts them: 660 integer cases (190 of
round 5 and 470 of round 6, less the controls, round 6's `E`, CR-021's forms, and its `N`, whose `w:legacy` Word 365 drops
whole). Each probe is loaded and `misplaced(package)` reads every part; each finding's path is evaluated as an XPath on the
source part and must reach the value, and is mapped back to its case by the body's child and its label. Word's save of each
is read as core-ts reads it (`_word_changed`, step for step: kept where the save writes the value back under the same name
and rule, or 65,535 where Word writes it -1; changed otherwise; left out, a 0).

| | cases |
|---|---|
| compared with Word's save | 614 (core-ts's), of them 8 row heights under `hRule="auto"`, which Word dropped and the table does not check |
| left out by Word (read 0) | 16, each reported with `word_reads` 0 |
| stated, not comparable (`w:gridCol` recomputed, `w:wBefore` and `w:wAfter` dropped) | 30, unmeasured here, none reported |
| **reported** | **413, exactly the set Word's save changed** (`test_the_report_flags_exactly_the_cases_word_s_save_changed`): one finding per case, `word_reads` what Word wrote, `read_as` the attribute it wrote the shared field under, `rule` the rule it wrote, and the value outside the finding's `range` |
| core-ts's count | **432 = 413 + 19**: the 19 are the stated cases core-ts reports from the objects table, which this table leaves unmeasured: 5 `w:gridCol` values outside the certain range -32768 to 65535 and 7 each of `w:wBefore` and `w:wAfter` outside int16. Not tuned: they are departure 2's |

Word's own saves of the probes hold nothing (`test_word_s_own_saves_hold_nothing`), nor round 6's `E` and `N`. **docx4j
CR-065 round 2's 64 decimals** (`tests/fixtures/integers/`, CR-021's) and CR-021's edges hold no out-of-range value in any
part, of the probes or of Word's saves: every value, read as Word reads it lexically, is in range.

**The corpus.** `samples/` and every fixture, 389 files, every part: **no misplaced content and no cell without its
paragraph; 413 findings, all `value.out_of_range` in the probe files' case parts** (the 413 above). Before departure 1 there
were 452 more, `dcterms:created` and `dcterms:modified` in 226 files' `docProps/core.xml`: the parse binds their
`xsi:type="dcterms:W3CDTF"` as a `DerivedElement` of a `W3CDTF`, which D2 refused; and the engine's own `create_package`
sets a bare `W3CDTF` there, which made the first save of the suite under the check raise.

### 11.6 The cost

- **The validator** over `samples/Symbols.docx`'s main part (790,857 bytes, 24,758 objects visited): **33
  ms** best, warm, against a parse of about 195 ms (17%); the first walk of a process 50 ms, the tables and plans built
  once. As first written it took 58 ms: the per-node path strings (now built only for a finding) and `integer_limit`'s scan
  of all 82 rows per value (now cached per path and context) were most of it; fetching the fields with one
  `operator.attrgetter` per class was slower (40 ms) and was not kept. Half of what is left is reading each single-valued
  field of each node (237,740 `getattr`s a walk), which the single-field check needs.
  `test_the_validator_costs_under_a_quarter_of_the_parse` (`slow`) holds it under 25% of the parse, best of nine.
- **The suite with the check on against off** (the fast suite at `82800fd`, one run after the other, nothing else
  running): **201.9 s on, 201.1 s off** (3 min 23.5 s and 3 min 22.8 s of wall time): +0.8 s, 0.4%, in the noise. The suite's
  documents are small, and a save checks only what it re-marshals. Off, the run failed the two tests of 11.8 that read
  the suite's default, which `5903d7d` makes read the switch. The fast suite before Phase B (`6cecdcf`) took 195.9 s for
  6,438 tests; the 350 added (346 in `test_validator.py`, 4 private-reader runs) take the rest.
- **The parse is unchanged.** `test_the_list_check_leaves_the_parse_within_five_percent` (Phase A's guard, `slow`, 201.6
  ms) failed once in the full suite at `82800fd`, at 201.8 ms; run alone it passes (193.4 to 198.4 ms), and outside pytest,
  interleaved, 20 parses each, `6cecdcf` gives 197.0 to 198.1 ms best and `82800fd` 195.3 to 197.4 ms. Phase B changes
  nothing on the parse's path (`XmlPart._parse` is not on it: the guard parses with `runtime.parser()`); the guard's margin
  is small under a five-minute run's heap. See 11.11 for the final run.

### 11.7 Departures, with why

1. **D2 accepts an `xsi:type` substitution of an element declared with no type of its own** (`placement._any_type`: a
   global element's class with no model base, no attributes and only wildcard content, the generator's reading of such a
   declaration as `xsd:anyType`: 48 of the model's, all Dublin Core's; a named type with wildcard content, `w:CT_PictureBase`
   or `w:CT_ShapeDefaults`, is not one, `82800fd`): a single field given a class with a type name of
   its own, which the serialiser writes with that `xsi:type`, and a `DerivedElement` naming its value's own type. Phase A's
   rule (10.3) refused every `xsi:type` write, since Word writes none; but OPC's core properties require
   `xsi:type="dcterms:W3CDTF"` on `dcterms:created` and `dcterms:modified`, every Word file has them, the parse binds them
   so, and `create_package` writes them so. Without this the validator reported two findings in every `.docx`, and the
   save's check refused every package whose core properties a session had unmarshalled. Every other `xsi:type` write is
   still refused, at add time and by the validator (`test_a_subclass_written_with_an_xsi_type_is_found`: a `PPr` in
   `w:pPrChange`). This changes Phase A's rule for the add-time check too; no Phase A test changed. For the user (11.9 Q2).
2. **Unmeasured attributes and attributes with no row are not reported**, where core-ts reports from the objects table
   (11.3): section 4.7 and CR-022 Q3 ("an attribute with no row is not checked"). It costs core-ts's 19 cases of 11.5 and
   the type-level attributes docx4j and the objects package range by class (`w:hps`, `w:drawingGrid*`, a table width with
   no `w:type`). For the user (11.9 Q4).
3. **No report on load, and none on a save without the check.** core-ts warns on every load and every save; here the
   validator is a walk of its own, run on demand and by `save(check=True)` (Q5's default, off), so a load and a default save
   cost nothing. For the user (11.9 Q5).
4. **One error class for the save's check**, `SaveCheckError`, not the condition's own (11.4), with the condition's code.
5. **A misplaced item's path names its element**, not what the serialiser writes for it (11.2).
6. **The checks are this phase's three.** No ordering rule exists to add (section 4.7: "ordering rules belong here too if
   any are added").

### 11.8 Tests and the numbers

- **`tests/test_validator.py`, 346 tests:** the findings' shape and `to_dict()`; Phase A's refused cases (`test_misplaced.py`'s
  65 and nine more, 74) found built under `unchecked()`, in a plain list assigned whole and through `object.__setattr__`,
  each finding's message the add-time error's (222 tests); a single field assigned the wrong class; a subclass written with
  an `xsi:type` found and its base not; only an element declared with no type taking any `xsi:type`; **section 10.3's 1,440 strangers**, each built under `unchecked()`: the 1,140 the
  check refuses found, the 300 it accepts not; a deep finding's path; a string's; a misplaced item walked into; the core
  properties' `xsi:type` legal; the cell rule, through the wrappers, and the engine's tables clean; the wrap model (one
  storage per row and row per storage, core-ts's strings, 82 rows held to the model, core-ts's stated readings); the
  out-of-range finding's fields, the lexical reading first, section 4.7's semantics, a column's sentinel and computed
  width, a row height under `auto`, an unmeasured attribute, a width with no row, a holder, a line's rule; **the oracle**
  (11.5); Word's saves; CR-065's decimals; the corpus; `misplaced()` marking no part; a tree, a part, a package; the address;
  the save's check (off by default and on in the suite, a refusal writing nothing, carried and introduced values and
  cells); `source_tree()` touching nothing.
- **`tests/test_performance.py`:** `test_the_validator_costs_under_a_quarter_of_the_parse` (`slow`).
- **`tests/openpackaging/test_private_readers.py`:** `validator.py::_tree_of` listed as a `read_contents` caller and
  `package.misplaced()` run as a reader (it writes nothing into the trees it reads).
- **`tests/content/test_table_rules.py`:** the opt-out (11.4).

### 11.9 For the user to decide

1. **The save's check** (11.4): misplaced content always raises; a cell or an out-of-range value raises only where the
   session introduced it, the source's carried and recorded. *Recommendation: keep.* The alternatives each fail on
   evidence: raising for every finding would refuse to save an edited file from elsewhere unless it was repaired, which the
   package never does on a save; raising only for misplaced content would let an engine edit that breaks a cell pass the
   suite.
2. **D2 widened for the `xsi:type` of an element declared with no type** (departure 1, the 48 Dublin Core elements).
   *Recommendation: keep*; it is the schema's, and without it every `.docx` holds two findings. docx4j's `checkParent` has no opinion on it.
3. **The cell rule's code, `cell.no_end_paragraph`.** *Recommendation: keep, and offer it to docx4j-core-ts and docx4j through
   the portfolio*, so that the ports name the finding alike when they report it (neither does yet).
4. **Unmeasured attributes and attributes with no row** (departure 2): not reported, as section 4.7 and CR-022 Q3 say; or
   reported as core-ts reports them (outside -32768 to 65535, `storage` `unmeasured: changed outside -32768..65535`,
   `word_reads` None; `w:wBefore` and `w:wAfter` as `int16`), which would bring the oracle to core-ts's 432. *Recommendation:
   keep*: a report of what Word does that rests on no measurement of it is the thing Q3 declined, and the 19 are all
   values Word's save dropped or recomputed anyway.
5. **Out-of-range on every save**, as core-ts's `lastSave.outOfRange` (departure 3), at 17% of a parse per re-marshalled part.
   *Recommendation: no; keep Q5's default*, and let the MCP server and the inspector call `misplaced()`.

### 11.10 Found, and not done

1. **The serialiser writes a misplaced `Tbl` in `R.content` as `<w:CT_Tbl/>`**, its type's name (the same mechanism as
   10.11 item 1, `docs/UPSTREAM.md` report 28). Nothing here depends on it; a path names the element instead.
2. **`docs/change-requests/README.md`'s CR-020 row** still says "Nothing implemented" (Phase A did not change it either):
   Phase C's.
3. **The portfolio's `python/CR-020.B` entry** wants `done`, with this section as its note, sent to the docx4j-portfolio
   session; the `tasks.yaml` rule (CLAUDE.md) is not this session's to apply. Section 11.9 Q3, if decided, is a note to
   docx4j-core-ts and docx4j through the same session.
4. **`tests/README.md`'s `test_private_readers.py` row** said twelve callers and nine readers, stale since CR-019 Phase B;
   it now says fifteen and eleven.

### 11.11 The gates

| gate | before (`6cecdcf`) | after (`5903d7d`) |
|---|---|---|
| the full suite (`.venv-fork/bin/python -m pytest`) | not run | **8,245 passed, 8 skipped**, 0 failed (5 min 12 s); at `82800fd` 8,244 passed and 1 failed, Phase A's parse guard at 201.8 ms against 201.6 (11.6), which passes here |
| the fast suite (`-m "not slow"`), the check on | 6,438 passed, 8 skipped, 1,382 deselected (195.9 s) | **6,788 passed, 8 skipped**, 1,457 deselected (221.5 s, the machine shared with a collection run; 201.9 s alone at `82800fd`) |
| the fast suite, the check off (`DOCX4J_PY_SAVE_CHECK=0`, at `82800fd`) | --- | 6,786 passed, 2 failed (the two `5903d7d` mends), 201.1 s |
| the new tests | --- | 346 in `test_validator.py`; 4 fast and 74 `slow` private-reader runs; 1 in `test_performance.py` (`slow`) |
| `scripts/roundtrip.py --models-module docx4j_py.wml --runtime docx4j_xsdata` | --- | **50 of 50 canonically identical**, 0 differences, 0 skipped, 34 `boolean-spelling` |
| acceptance (`scripts/compare_acceptance.py`) | `out/acceptance-cr024/`, CR-024's build, the code of `6cecdcf` | `out/acceptance-cr020b/`: **17 artefacts, 246 entries, 0 differ** beyond the dates; `out/acceptance-cr020b-checked/`, built under `checking(True)`: 0 differ; the same build counted (`out/acceptance-cr020b-counted/`): 18 saves checked, none refused, nothing carried |
| `codegen/generate.sh --check` (a worktree at `5903d7d`) | --- | two generations byte identical, 311 modules (Phase A's 309 and the two new files); nothing generated differs from the commit, and `validator.py` and `model/word_integers.py` survive `clean.py` |
| `scripts/threads.py` | --- (passed at Phase A) | **passes** (eight threads, three modes, no document lost or altered) |
| `ruff check` on the 16 files touched | --- | no new finding; the 8 there before (`opc_package.py` 2, `xml_part.py` 2, `save.py` 1, `test_performance.py` 3) left as they were |
| `Symbols.docx`'s parse, best, outside pytest | 197.0 to 198.1 ms | 195.3 to 197.4 ms |
| the validator over `Symbols.docx`'s main part, best, warm | --- | 33 to 35 ms (17 to 18% of the parse) |

No Word check (section 5): output is unchanged, and the acceptance documents are identical.

### 11.12 The decisions (2026-10-28)

The user, on 11.9: "keep what's built". So no code, test or generated file changed for them, and no gate moved from
11.11's. The coordinator re-ran the fast suite (6,788 passed, 8 skipped) and `roundtrip.py` (50 of 50) at `d0dade4` and
exercised the save's check by hand. One commit followed, `78e11f4`: `test_validator.py` filters the converter warnings
its probe files raise on purpose, as the other probe readers do (267 of the fast suite's 287).

1. **The save's check stays as built** (11.4): misplaced content always raises; a cell or an out-of-range value raises
   only where the session introduced it, and what the source held is saved as it is and listed on `last_save.findings`.
2. **D2 keeps the `xsi:type` of an element declared with no type** (departure 1, the 48 Dublin Core elements), at add
   time and in the validator.
3. **The cell rule's code is `cell.no_end_paragraph`**, offered to docx4j-core-ts and docx4j on 2026-10-28, so that a port
   reporting the finding names it alike.
   Both took it the same day. docx4j-core-ts also took `target.misplaced` for its CR-016 phase 2 warning on a misplaced
   child (its CR-016 section 4, D7), with these field names where its objects report has the data: `part_name`,
   `path`, `index`, `element`, and `holder` and `item` as docx4j's Java class names, as here. It has no `address`,
   `field` or `hint`; its own `rule`, `text` (a string's) and load `line` have no counterpart here; and its message
   is docx4j's wording. Its finding covers docx4j CR-064's refuse sets, of which this allow rule finds a superset.
   docx4j reports neither condition today (`TableRules` repairs a cell on load) and will use both codes if a
   validator there reports them.
4. **Unmeasured attributes and attributes with no row are not reported** (departure 2): the oracle stays 413, and
   core-ts's other 19 (`w:gridCol`, `w:wBefore`, `w:wAfter`) wait on a measurement of what Word does with them.
5. **No out-of-range report on every save** (departure 3): Q5's default, `check` off, stands; the MCP server and the
   inspector call `misplaced()`.
