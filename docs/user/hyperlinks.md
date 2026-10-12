# Hyperlinks

How to put a hyperlink into a paragraph, what the markup is, one thing still wrong at the end of
a link, and a proposal for what the content API still lacks (section 5). `Range.hyperlink`, Office
JS's member, is implemented since 2026-10-12 (CR-003 section 21); everything in sections 2 and 3
was run on that date and the output is what is shown.

## 1. What a hyperlink is in WordprocessingML

Word writes a link as a **`w:hyperlink` element wrapping runs**, inside the paragraph:

```xml
<w:p>
  <w:r><w:t xml:space="preserve">Read the news on </w:t></w:r>
  <w:hyperlink r:id="rId4" w:history="1">
    <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Slashdot</w:t></w:r>
  </w:hyperlink>
  <w:r><w:t xml:space="preserve"> every morning.</w:t></w:r>
</w:p>
```

Three things make it a link:

- **The destination is not in the element.** `r:id` names a relationship of the *part the
  paragraph lives in* (the main document part, a header, a footnote), of type
  `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`, with
  `TargetMode="External"` and the URL as its target. A link to a place in the same document uses
  `w:anchor="bookmarkName"` instead of `r:id`, and needs no relationship. `w:history="1"` is
  what Word and docx4j write on every link they make ("add the target to the viewed
  hyperlinks"); nothing reading a link depends on it.
- **The blue underline is a character style**, `Hyperlink`, on the runs inside. Nothing about
  `w:hyperlink` itself changes how the text looks; a link whose runs lack the style is a working
  link that looks like body text. `create_package()` writes the style, and `Range.hyperlink` and
  `insert_markdown` add it from `resources/KnownStyles.xml` when the document has none.
- `w:hyperlink` is a **run holder**, like a run-level `w:sdt` or a `w:ins`: it contains runs and
  sits between the paragraph and them. The content API's `Range` knows the holders: a comment or
  a content control over a span that crosses one is refused (CR-003 section 3.4, `SpanError`),
  and `Range.hyperlink` takes a link whose edge the span crosses into the new one, as Word does
  (section 2.1). A `w:ins` may sit **inside** a `w:hyperlink`, never the other way round
  (`CT_RunTrackChange` has no `w:hyperlink`).

There is an older form, the `HYPERLINK` **field** (`w:fldChar` begin, an instruction run holding
`HYPERLINK "https://..."`, the result runs, `w:fldChar` end), which Word still reads and some
generators still write. `w:hyperlink` is what Word writes since 2007 and is the only form this
document's examples make. Reading the field form is a separate matter (section 5.4).

## 2. Adding one

### 2.1 `Range.hyperlink`, Office JS's way

Write the text, find the span, set its `hyperlink`:

```python
from docx4j_py import load

pkg = load("in.docx")
p = pkg.body.insert_paragraph("Read the news on Slashdot every morning.")
p.search("Slashdot")[0].hyperlink = "https://slashdot.org/"
pkg.save("out.docx")
```

The value is Office JS's `address#location`: the address before the first `#` becomes an external
relationship of **the range's own part** (a header's paragraph relates from the header part) and
the link's `r:id`; what follows the `#` becomes its `w:anchor`. So `"#chapter1"` is a link to a
bookmark in this document, with no relationship, and `"https://example.com/doc#s2"` both. Reading
gives the link the range is in, in the same shape, or `""`:

```python
p.search("Slashdot")[0].hyperlink        # 'https://slashdot.org/'
p.search("every")[0].hyperlink           # ''
p.search("Slashdot")[0].hyperlink = ""   # unlinks: the runs stay, the style comes off
```

What the setter does, as Word does it (docx4j-core-ts measured each rule through Office JS):

- the runs are split at the span's boundaries and wrapped in one `w:hyperlink` carrying
  `w:history`; each takes the `Hyperlink` character style **in place of** any character style it
  had, and loses its direct colour, keeping the rest of its direct formatting;
- a span that crosses a link's edge takes that whole link into the new one; a span inside a link
  replaces it with a link over the span alone; `""` on part of a link removes all of it;
- an **empty** range (a caret) is refused, as Office JS refuses it; so is a span that would put
  the link across a content control, a field or a tracked change --- with a `SpanError`, before
  anything has changed;
- under change tracking the restyle is a formatting revision (`w:rPrChange`), the wrap itself is
  no revision, and a span inside a tracked insertion gets a link **holding** the insertion;
- the call records a `"hyperlink"` `ChangeReport` naming the paragraph, the `.rels` it wrote and
  `styles.xml` when it had to define the style, and a `dry_run` takes the relationship away again.

CR-003 section 21 is the whole of it, with every place it differs from docx4j-core-ts.

### 2.2 Through markdown

`insert_markdown` writes the element, the relationship on the right part and the `Hyperlink`
style, for a new paragraph:

```python
[p] = pkg.body.insert_markdown(
    "Read the news on [Slashdot](https://slashdot.org/) every morning."
)
```

and into an existing one, where a one-paragraph fragment has its runs **merged into the
paragraph** (CR-003 section 4's paste rule):

```python
p = pkg.body.paragraphs[3]
p.insert_markdown("[Slashdot](https://slashdot.org/)", location="End")
```

Both produce the markup of section 1, without `w:history`. This is the way for an agent, an MCP
tool or a script that already has the text as markdown; it is the way `docx4j-mcp-py`'s `insert`
tool will take (CR-006 section 3.2). Three things it gets wrong that 2.1 gets right (CR-003
section 21.6): its `ChangeReport` names the paragraph and `styles.xml` but **not** the `.rels` it
wrote; a `dry_run` of it leaves that relationship on the real package; and **under change
tracking** a link's runs in a new paragraph are not tracked, while a link merged into an existing
paragraph is written inside the `w:ins`, which is not valid markup --- with tracking on, use 2.1.

### 2.3 Through the typed tree, when you want control

The same three parts by hand --- the relationship first, so its id can go on the element:

```python
from docx4j_py.wml import el
from docx4j_py.wml.builders import r

HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

p = pkg.body.insert_paragraph("Read the news on ")
rel = pkg.main_document_part.relationships_part.add_external_relationship(
    HYPERLINK, "https://slashdot.org/"
)
p.p.content.append(el.hyperlink(id=rel.id, content=[r("Slashdot", style="Hyperlink")]))
```

`ChildList.append` sets the parent, so no `link_parents` call. For a paragraph in a header or a
footnote, hang the relationship on **that** part's `relationships_part`, not the main document's:
`r:id` is resolved against the part the element is in. For an internal link, `el.hyperlink(anchor="bookmark", ...)`
and no relationship. This is docx4j's own recipe (`HyperlinkTest.createHyperlink` in the Java
samples: add the relationship, unmarshal the element with its id) without the string template.
It records no `ChangeReport` and does not define the `Hyperlink` style.

### 2.4 Reading one back

- The content API: `range.hyperlink` on the link's text, or on a caret inside it or at its end.
  Every `w:hyperlink` in this repository's `samples/` and `tests/fixtures/` reads back its target
  or `#anchor` that way (a test holds it).
- The python-docx facade: `paragraph.hyperlinks` (`Hyperlink.text`, `.address`, `.url`,
  `.fragment`) and `paragraph.iter_inner_content()`, python-docx 1.2.0's read-only surface.
- `to_markdown()` renders every `w:hyperlink` as `[text](target)`, external or `#anchor`.

## 3. A defect in 0.1.0: appending after a link extends it

Section 2.1's way --- the whole sentence first, then the link --- is fine, and so are 2.2's. But
this sequence is not, whichever way the link was made:

```python
p = pkg.body.insert_paragraph("Read the news on ")
p.insert_text("Slashdot", location="End").hyperlink = "https://slashdot.org/"
p.insert_text(" every morning.", location="End")
```

gives

```xml
<w:hyperlink w:history="1" r:id="rId4">
  <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Slashdot every morning.</w:t></w:r>
</w:hyperlink>
```

--- the trailing text has become part of the link, and `p.search("every")[0].hyperlink` reads
`https://slashdot.org/`. `Paragraph.insert_text(location="End")` is `splice(length, length,
text)`, and the splice's rule (CR-003 section 11) is *extend the run whose end is at the offset,
else make a new run with the neighbour's formatting*. The run whose end is at the offset is the
link's, and the rule does not look at what holds it. Word closes a link when you type after it:
the hyperlink is a field-like unit, and text typed at its end is outside it, in the paragraph's
own formatting. python-docx's `add_run` after a hyperlink appends a `w:r` to the paragraph,
outside the `w:hyperlink`, for the same reason. The `insert_markdown` form of 2.2 followed by the
same `insert_text` does the same.

**The fix** is one condition in `splice`: a run inside a `w:hyperlink` (or any run holder) is not
extended at its holder's *last* offset; the new run goes after the holder, with the formatting
of the nearest run *outside* it, or none. The same applies at `"Start"` before a hyperlink that
opens the paragraph. `Range.insert_text("After")` on a range that ends at the link's end has the
same rule to follow. It is a bug fix, not a design change; section 5.5 keeps it in the proposal,
and section 5.8 question 4 asks when.

Until it lands, write the text first and link the span last (2.1), or put the whole sentence in
one `insert_markdown`.

## 4. How `Range.hyperlink` came to be in the subset

The content API's vocabulary is Office JS's, and the list of members it promises
(`tests/office_js_subset.json`) is **derived** from docx4j-core-ts's compile-time subset
declaration (`test/office-js-subset.ts`), so that the two engines cannot drift apart
(`scripts/office_js_subset.py`, CR-003 section 11.4). Until September that declaration stopped at
"the members an agent editing text needs", and `Range.hyperlink` was not among them --- not refused,
just not asked for, since the markdown importer covered "an agent adds a link" well enough.

On 2026-09-25 docx4j-core-ts added it (its `d64ab67`, its CR-002 section 20), at the request of
its editor, which needed a link it could set on existing text; over the next three days it
measured Word through Office JS and corrected the setter five times (its section 25: the style,
the edge of another link, the run merge, the caret). Re-deriving the list here on 2026-10-12 added
the one member, which CR-003 section 21 implements with core-ts's rules. That is the route
section 5.6 described, run in the other direction: core-ts first, then the re-derivation, then the
member here. `Range.getHyperlinkRanges()` (Office JS WordApi 1.3) is still not in core-ts's
declaration, and so not in the promise.

## 5. Proposal: what is left

A CR-003 amendment beyond section 21, in this order. Section 21 settled 5.1's first half;
the rest is unchanged in substance and its open questions are still open (5.8).

### 5.1 The vocabulary: `Range.get_hyperlink_ranges()`

`Range.hyperlink` is implemented (CR-003 section 21). Where it differs from what this section first
proposed, it follows docx4j-core-ts, which followed Word: a span that crosses a link's edge takes
the link in rather than being refused; the setter **replaces** a run's character style and removes
its direct colour, not only adds `Hyperlink`; an empty range is refused; and a span inside a
tracked insertion gets a link holding the `w:ins`, where this section said the link would go
inside it, which the schema does not allow. CR-003 section 21.4 lists them.

What remains of the vocabulary:

| Member | Office JS | Here |
|---|---|---|
| `Range.get_hyperlink_ranges()` | "the hyperlink child ranges within the range" | `list[Range]`, one per `w:hyperlink` inside the span, each a `Range` whose `hyperlink` reads the address and whose `text` and `font` edit the link text |
| `Paragraph.get_hyperlink_ranges()` | not in Office JS (`Paragraph` reaches it through `getRange()`) | **not added**; `paragraph.get_range().get_hyperlink_ranges()` is the spelling, one call longer and the same as Office JS |

Reading every link of a paragraph is then
`[(r.text, r.hyperlink) for r in p.get_range().get_hyperlink_ranges()]`. It goes into core-ts's
declaration first (5.6).

### 5.2 The sugar: `insert_hyperlink(text, address, location=)`

One call for the common case, on `Body`, `Paragraph` and `Range`, with the locations
`insert_text` takes on each; returns the link's `Range`:

```python
p.insert_hyperlink("Slashdot", "https://slashdot.org/", location="End")
```

It is exactly `insert_text(text, location=...)` followed by `.hyperlink = address` and is
implemented as that, so there is one code path --- which is why it waits on section 3's fix: until
then, text inserted after the link at `"End"` joins it. It is the **first extension member** of the
content API --- a name Office JS does not have --- and CR-003 section 3.2's rule for extensions
applies: it is marked `extension` in the subset, documented as sugar over the two vocabulary
calls, and `to_api_script` emits the two calls, not the sugar, so a script it writes is Office JS
shaped.

Whether the sugar is worth being the first extension is open question 1. The case for it: it is
the call every user asks for first ("how do I add a link"), and python-docx's most-requested
missing feature for a decade has been exactly `add_hyperlink` (its issue 74). The case against:
the two vocabulary calls of 2.1 are not many, and an extension is a promise core-ts then has to
match.

### 5.3 The python-docx facade

python-docx 1.2.0 has no writer for hyperlinks, so `tests/python_docx_subset.json` asks for
none and the facade's promise (305 of 312) is untouched. `Paragraph.add_hyperlink(text, url)`
is proposed as a facade extension **outside** the subset --- the shape of the community recipe
every python-docx user has copied into their project --- returning the facade's `Hyperlink`,
over `Range.hyperlink`. Open question 3.

### 5.4 Reading the `HYPERLINK` field form

`Range.hyperlink`'s getter reads `w:hyperlink` only, and `get_hyperlink_ranges()` would too. The
field form (section 1) is read by docx4j's field machinery (`org.docx4j.model.fields`), not by its
hyperlink traversal, and `to_markdown` does not render it today either. Reading it is a separate
small item once fields have a home in this API (they do not yet); the getter returns `""` for a
field-form link, which is at least not wrong about `w:hyperlink`.

### 5.5 The defect of section 3

`Paragraph.splice` and `Range.insert_text` at a holder's boundary put the new run **outside** the
holder (section 3). Tested by the exact sequence of section 3, and by the mirror at `"Start"`.
It also unblocks `to_api_script` emitting `Range.hyperlink` for a linked paragraph, which today
falls back to `insert_xml` (CR-003 section 21.8): the script inserts the link's text, sets the
link, and inserts the text after it, which the defect would pull into the link.

### 5.6 The route through docx4j-core-ts

The subset is derived, so the sequence is the one `Range.hyperlink` took (section 4): propose
`getHyperlinkRanges` on core-ts's `OfficeSubset.Range` (and `insertHyperlink` as an
`extension`-tagged member, if question 1 says yes) to the core-ts session; when its
`test/office-js-subset.ts` carries them, `scripts/office_js_subset.py` re-derives
`tests/office_js_subset.json` and the implemented test turns red here until the members land. If
core-ts declines or defers, the members are hand-carried here as Phase H's list interfaces were in
September, with the block that the script deletes once core-ts catches up.

### 5.7 Tests and effort

`tests/content/test_hyperlinks.py` has `Range.hyperlink`'s 29 tests. To it: `get_hyperlink_ranges()`
over the corpus's links (the same 23 links in 7 documents the getter's test reads); the sugar on a
new paragraph, an existing paragraph and a range, with the section 3 sequence and its mirror; the
facade's `add_hyperlink` read back by python-docx in the parity test; `to_api_script` emitting the
setter, with the executed round trip. A Word check: artefact 6 gained an external and an internal
link on 2026-10-12; the sugar adds a link followed by text.

About half a day, plus the core-ts round trip.

### 5.8 Open questions (recommendations for decision)

1. **The `insert_hyperlink` sugar.** Add it as the content API's first extension member
   (recommended: it is the question everyone asks, and the extension mechanism exists for
   exactly this) or stay with the two vocabulary calls.
2. ~~**Tracking the wrap.**~~ *Answered 2026-10-12 by docx4j-core-ts's setter, which this package
   follows (CR-003 section 21.2):* only the `w:rPrChange` on the runs, as recommended; the wrap is
   no revision.
3. **The facade.** Add `Paragraph.add_hyperlink(text, url)` as a documented extension outside the
   python-docx subset (recommended) or leave the facade at python-docx's read-only surface.
4. **Order.** Fix section 3's defect now as a patch, ahead of the sugar (recommended: it is a
   silent corruption of a link and a two-line change), or with the sugar.
