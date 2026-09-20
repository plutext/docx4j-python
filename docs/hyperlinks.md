# Hyperlinks

How to put a hyperlink into a paragraph with what `docx4j 0.1.0` has, what the markup is, one
thing 0.1.0 gets wrong at the end of a link, and a proposal for the verb the content API does
not have yet (section 5). Everything in sections 2 and 3 was run on 2026-09-21 and the output is
what is shown.

## 1. What a hyperlink is in WordprocessingML

Word writes a link as a **`w:hyperlink` element wrapping runs**, inside the paragraph:

```xml
<w:p>
  <w:r><w:t xml:space="preserve">Read the news on </w:t></w:r>
  <w:hyperlink r:id="rId4">
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
  `w:anchor="bookmarkName"` instead of `r:id`, and needs no relationship.
- **The blue underline is a character style**, `Hyperlink`, on the runs inside. Nothing about
  `w:hyperlink` itself changes how the text looks; a link whose runs lack the style is a working
  link that looks like body text. `create_package()` and `insert_markdown` write the style from
  `resources/KnownStyles.xml` when the document has none.
- `w:hyperlink` is a **run holder**, like a run-level `w:sdt` or a `w:ins`: it contains runs and
  sits between the paragraph and them. The content API's `Range` knows the holders and refuses a
  span that crosses one (CR-003 section 3.4, `SpanError`), which is why "the text of a link" and
  "the text after it" are never one `Range`.

There is an older form, the `HYPERLINK` **field** (`w:fldChar` begin, an instruction run holding
`HYPERLINK "https://..."`, the result runs, `w:fldChar` end), which Word still reads and some
generators still write. `w:hyperlink` is what Word writes since 2007 and is the only form this
document's examples make. Reading the field form is a separate matter (section 5.4).

## 2. Adding one today

### 2.1 Through markdown, the short way

`insert_markdown` writes the element, the relationship on the right part and the `Hyperlink`
style, for a new paragraph:

```python
from docx4j_py import load

pkg = load("in.docx")
[p] = pkg.body.insert_markdown(
    "Read the news on [Slashdot](https://slashdot.org/) every morning."
)
pkg.save("out.docx")
```

and into an existing one, where a one-paragraph fragment has its runs **merged into the
paragraph** (CR-003 section 4's paste rule):

```python
p = pkg.body.paragraphs[3]
p.insert_markdown("[Slashdot](https://slashdot.org/)", location="End")
```

Both produce the markup of section 1. The `ChangeReport` names the paragraph and the
relationships part it touched. This is the way to use from an agent, an MCP tool or a script that
already has the text as markdown; it is the way `docx4j-mcp-py`'s `insert` tool will take
(CR-006 section 3.2).

### 2.2 Through the typed tree, when you want control

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

### 2.3 Reading one back

- The python-docx facade: `paragraph.hyperlinks` (`Hyperlink.text`, `.address`, `.url`,
  `.fragment`) and `paragraph.iter_inner_content()`, python-docx 1.2.0's read-only surface.
- `to_markdown()` renders every `w:hyperlink` as `[text](target)`, external or `#anchor`.
- The content API: `paragraph.runs` descends into hyperlinks; a `Range` inside one edits its text
  and font like any other. There is no member that says "this range is a link" or gives its
  address (section 5).

## 3. A defect in 0.1.0: appending after a link extends it

Both ways above are fine. But this sequence is not:

```python
p = pkg.body.insert_paragraph("Read the news on ")
p.insert_markdown("[Slashdot](https://slashdot.org/)", location="End")
p.insert_text(" every morning.", location="End")
```

gives

```xml
<w:hyperlink r:id="rId4">
  <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Slashdot every morning.</w:t></w:r>
</w:hyperlink>
```

--- the trailing text has become part of the link. `Paragraph.insert_text(location="End")` is
`splice(length, length, text)`, and the splice's rule (CR-003 section 11) is *extend the run whose
end is at the offset, else make a new run with the neighbour's formatting*. The run whose end is
at the offset is the link's, and the rule does not look at what holds it. Word closes a link
when you type after it: the hyperlink is a field-like unit, and text typed at its end is outside
it, in the paragraph's own formatting. python-docx's `add_run` after a hyperlink appends a
`w:r` to the paragraph, outside the `w:hyperlink`, for the same reason.

**The fix** is one condition in `splice`: a run inside a `w:hyperlink` (or any run holder) is not
extended at its holder's *last* offset; the new run goes after the holder, with the formatting
of the nearest run *outside* it, or none. The same applies at `"Start"` before a hyperlink that
opens the paragraph. `Range.insert_text("After")` on a range that ends at the link's end has the
same rule to follow. It is a bug fix, not a design change; section 5.5 puts it in the proposal so
that it is tested with the verb.

Until it lands, order the calls so that the link goes in last, or put the whole sentence in one
`insert_markdown`.

## 4. Why there is no verb, and why `Range.hyperlink` is not in the subset

The content API's vocabulary is Office JS's, and the list of members it promises
(`tests/office_js_subset.json`) is **derived** from docx4j-core-ts's compile-time subset
declaration (`test/office-js-subset.ts`), so that the two engines cannot drift apart
(`scripts/office_js_subset.py`, CR-003 section 5). That declaration was chosen in core-ts's
CR-002 as "the members an agent editing text needs" --- `Word.Range` there is `text`, `font`,
`style`, `insertText`, `insertParagraph`, `search`, `delete`, `getRange`, comments and tracked
changes. `Range.hyperlink` and `Range.getHyperlinkRanges()` (Office JS WordApi 1.3) were not in
it, not because links were considered and refused but because neither engine's phases needed
them: the markdown importer of CR-003 Phase K covered "an agent adds a link" well enough that no
phase asked for more. Nothing in the subset is marked `extension` today; every member is
core-ts's.

So the honest answer to "why not `Range.hyperlink`" is that the subset stops where core-ts's
stopped, and the way to add the setter is to add it **there** first, then re-derive. That is
what section 5 proposes.

## 5. Proposal: hyperlinks in the content API

A CR-003 amendment (a Phase L, or a section 21 if the user prefers amendments to phases), about
a day, in this order.

### 5.1 The vocabulary: `Range.hyperlink` and `Range.get_hyperlink_ranges()`

Office JS's own, so that the Python and TypeScript engines and Office JS say the same thing:

| Member | Office JS | Here |
|---|---|---|
| `Range.hyperlink` (get) | "the first hyperlink in the range, or `''`" --- the address | `str`: the relationship's target for `r:id`, `#anchor` for `w:anchor`, `""` when the range holds no `w:hyperlink` |
| `Range.hyperlink` (set) | "sets a hyperlink on the range. Existing hyperlinks in the range are deleted" | wraps the range's runs in one `w:hyperlink`: a new external relationship on the range's part (or `w:anchor` for a value beginning `#`), the runs split at the span's boundaries as `Font` setters do, `Hyperlink` on each run's `w:rStyle`, any `w:hyperlink` the span already contained unwrapped first. Assigning `""` or `None` unwraps: the runs come out, the style comes off, the relationship stays (it is harmless and may be shared) |
| `Range.get_hyperlink_ranges()` | "the hyperlink child ranges within the range" | `list[Range]`, one per `w:hyperlink` inside the span, each a `Range` whose `hyperlink` reads the address and whose `text` and `font` edit the link text |
| `Paragraph.get_hyperlink_ranges()` | not in Office JS (`Paragraph` reaches it through `getRange()`) | **not added**; `paragraph.get_range().get_hyperlink_ranges()` is the spelling, one call longer and the same as Office JS |

Adding a link to a paragraph is then two calls in the vocabulary, the way an Office JS script
does it:

```python
p.insert_text("Slashdot", location="End").hyperlink = "https://slashdot.org/"
```

and reading them is `[(r.text, r.hyperlink) for r in p.get_range().get_hyperlink_ranges()]`.

**Rules.**

- A span that *crosses* a holder boundary is refused with the existing `SpanError` (section 1);
  a span *inside* a link replaces that link (Office JS's "existing hyperlinks in the range are
  deleted", read at the span's granularity); a span that *contains* whole links unwraps them and
  makes one.
- The `Hyperlink` style is written by id `Hyperlink` and created from `KnownStyles.xml` when the
  document lacks it, as `insert_markdown` does now; the setter does not put colour or underline
  on the runs directly.
- The setter records a `ChangeReport` like every mutating call: the paragraph's address, the
  relationships part when a relationship was added.
- Under change tracking (CR-003 Phase F), the runs' `w:rStyle` change is a `w:rPrChange`, as
  `Font` setters record theirs; the wrapping itself is **not** a tracked change --- Word records a
  hyperlink set on existing text the same way, as a formatting revision on the runs (a link
  inserted *with* its text, `insert_text` then the setter, is already a `w:ins`, and the
  `w:hyperlink` goes inside it). Open question 2.
- `dry_run` un-adds the relationship, as it un-adds parts.

### 5.2 The sugar: `insert_hyperlink(text, address, location=)`

One call for the common case, on `Body`, `Paragraph` and `Range`, with the locations
`insert_text` takes on each; returns the link's `Range`:

```python
p.insert_hyperlink("Slashdot", "https://slashdot.org/", location="End")
```

It is exactly `insert_text(text, location=...)` followed by `.hyperlink = address` and is
implemented as that, so there is one code path. It is the **first extension member** of the
content API --- a name Office JS does not have --- and CR-003 section 3.2's rule for extensions
applies: it is marked `extension` in the subset, documented as sugar over the two vocabulary
calls, and `to_api_script` emits the two calls, not the sugar, so a script it writes is Office JS
shaped.

Whether the sugar is worth being the first extension is open question 1. The case for it: it is
the call every user asks for first ("how do I add a link"), the answer in 0.1.0 is a markdown
detour or a relationships recipe, and python-docx's most-requested missing feature for a decade
has been exactly `add_hyperlink` (its issue 74). The case against: two calls in the vocabulary
are not many, and an extension is a promise core-ts then has to match.

### 5.3 The python-docx facade

python-docx 1.2.0 has no writer for hyperlinks, so `tests/python_docx_subset.json` asks for
none and the facade's promise (305 of 312) is untouched. `Paragraph.add_hyperlink(text, url)`
is proposed as a facade extension **outside** the subset --- the shape of the community recipe
every python-docx user has copied into their project --- returning the facade's `Hyperlink`,
over 5.1's setter. Open question 3.

### 5.4 Reading the `HYPERLINK` field form

`Range.hyperlink`'s getter and `get_hyperlink_ranges()` read `w:hyperlink` only in this
proposal. The field form (section 1) is read by docx4j's field machinery
(`org.docx4j.model.fields`), not by its hyperlink traversal, and `to_markdown` does not render it
today either. Reading it is a separate small item once fields have a home in this API (they do
not yet); the getter returns `""` for a field-form link, which is at least not wrong about
`w:hyperlink`.

### 5.5 The defect of section 3, fixed with the verb

`Paragraph.splice` and `Range.insert_text` at a holder's boundary put the new run **outside** the
holder (section 3). Tested by the exact sequence of section 3, and by the mirror at `"Start"`.

### 5.6 The route through docx4j-core-ts

The subset is derived, so the sequence is: propose `hyperlink` and `getHyperlinkRanges` on
core-ts's `OfficeSubset.Range` (and `insertHyperlink` as an `extension`-tagged member, if
question 1 says yes) to the core-ts session; when its `test/office-js-subset.ts` carries them,
`scripts/office_js_subset.py` re-derives `tests/office_js_subset.json` and the implemented test
turns red here until 5.1 lands. Both engines then implement against the same declaration, with
Java docx4j (`P.Hyperlink`, `HyperlinkTraverser`, the samples' `createHyperlink`) as the oracle
for the markup. If core-ts declines or defers, the members are hand-carried here as Phase H's
list interfaces were in September, with the block that the script deletes once core-ts catches
up.

### 5.7 Tests and effort

`tests/content/test_hyperlinks.py`: the getter over the corpus's links (every `w:hyperlink` in
`samples/` reads back its relationship target or anchor; `to_markdown` and the getter agree); the
setter on a new paragraph, an existing paragraph, a range inside a run, a range spanning runs, a
range containing a link, a header's paragraph (the relationship lands on the header part); the
unwrap; the anchor form; the style created when absent; the `ChangeReport`; `dry_run`; tracking
on; the section 3 sequence; `to_api_script` round trip for the sugar; the facade's
`add_hyperlink` read back by python-docx in the parity test. A Word check: one acceptance
artefact gains a paragraph with an external link, an internal link and a link followed by text.

About a day, plus the core-ts round trip.

### 5.8 Open questions (recommendations for decision)

1. **The `insert_hyperlink` sugar.** Add it as the content API's first extension member
   (recommended: it is the question everyone asks, and the extension mechanism exists for
   exactly this) or stay with the two vocabulary calls.
2. **Tracking the wrap.** Record only the `w:rPrChange` on the runs, as Word does for a link set
   on existing text (recommended), or also wrap in a `w:ins`/`w:del` pair as a replacement, which
   Word does not.
3. **The facade.** Add `Paragraph.add_hyperlink(text, url)` as a documented extension outside the
   python-docx subset (recommended) or leave the facade at python-docx's read-only surface.
4. **Order.** Fix section 3's defect now as a 0.1.1 patch, ahead of the verb (recommended:
   it is a silent corruption of a link and a two-line change), or with the verb.
