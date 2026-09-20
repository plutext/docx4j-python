# Find and replace: the main document part, headers and footers, notes, or everything

Every part that holds paragraphs has a **`Body`** --- the main document part, each header and
footer, the footnotes, the endnotes and the comments --- and `Body` has `search`, `find` and
`replace_text` (CR-003 Phase B, section 3.3). So "the main document part" is `pkg.body`, "a
header" is that part's body, and "everything" is `pkg.bodies()`, a list of all of them with the
main part first. There is no verb that spans parts by itself; the loop is one line.

Everything below was run on 2026-09-21 against `samples/Headers.docx` (a main part, six
headers, footnotes and endnotes) and the numbers are what came out.

## 1. The bodies of a document

```python
from docx4j_py import load

pkg = load("Headers.docx")
for body in pkg.bodies():
    print(body.prefix, type(body.part).__name__, len(body.paragraphs), repr(body.text[:30]))
```

```
body          MainDocumentPart 23 'Header and footer test\n\n\nPage'
header:rId8   HeaderPart        1 'First page header.'
header:rId13  HeaderPart        1 'ODD/even … BLACK.'
header:rId7   HeaderPart        1 'Normal header.'
header:rId12  HeaderPart        1 'Odd/EVEN.. white'
header:rId11  HeaderPart        1 'First page header.'
header:rId10  HeaderPart        1 'Just a plain header.'
footnotes     FootnotesPart     2 '\n'
endnotes      EndnotesPart      2 '\n'
```

`body.prefix` is the part's **address prefix** (CR-003 section 3.4): `body`, `header:<rId>`,
`footer:<rId>`, `footnotes`, `endnotes`, `comments`. Every address a `find` returns begins with
it, so a hit says which part it is in, and it is how to pick a subset: `prefix.startswith("header:")`.
A header or footer is named by the relationship id the main part reaches it through, since a
document can have six headers and none of them has a name of its own. (The two "footnotes" and
"endnotes" paragraphs in a document with no notes are Word's separator and continuation-separator
notes, which every document carries.)

The order is the main part, the headers, the footers, then footnotes, endnotes and comments.

## 2. Find

`search(text, **options)` returns `Range`s, live views over the runs, one per match; `find(text)`
returns hits with an **address** and a snippet, which is what to show a person or a model:

```python
pkg.body.search("Section Break")                 # [<Range 9:22 'Section Break'>]

for body in pkg.bodies():
    for hit in body.find("header", limit=3):
        print(hit.address, repr(hit.snippet))
```

```
body/0        'Header and footer test'
header:rId8/0 'First page header.'
header:rId7/0 'Normal header.'
header:rId11/0 'First page header.'
header:rId10/0 'Just a plain header.'
```

The options, on `search`, `find` and `replace_text` alike:

| Option | Default | Meaning |
|---|---|---|
| `match_case` | `False` | case-sensitive; the default is `re.IGNORECASE` |
| `match_whole_word` | `False` | Unicode word boundaries either side |
| `match_wildcards` | `False` | Word's wildcards: `?` one character, `*` any run, `[a-z]` a class, `[!x]` a negation, `<` and `>` word starts and ends, `@` and `{n,m}` repetition |
| `limit` | none | stop after this many |

A match is found on each paragraph's **whole text**, so it spans runs freely --- a word half
bold, a word split across a tracked insertion --- and only formatting or replacing splits them.
A match never spans two paragraphs, because a paragraph is a `Range`'s outer limit.

**`pkg.find(...)` is the main document part only.** It is the agent surface's shortcut for the
common case, and it does not look in the headers; use the loop above for every part.

## 3. Replace

`Body.replace_text(find, replace, **options)` replaces every match in that body, last match
first so the offsets stay valid, and returns the count. One `ChangeReport` covers the whole
body, naming the pair and every paragraph it touched:

```python
pkg = load("Headers.docx")
pkg.body.replace_text("Page 2", "Page two")       # 4
pkg.changes[-1].to_dict()
# {'operation': 'replace_text', 'addresses': ['body/2', 'body/13', 'body/18', 'body/21'],
#  'text_before': 'Page 2', 'text_after': 'Page two', 'parts_touched': ['/word/document.xml'], ...}
```

The replacement goes into the run the match starts in, so it takes that run's formatting; a
match that spanned two formattings becomes one run in the first's (Word's rule).

### 3.1 One part

```python
pkg.body.replace_text("Section Break", "section break")           # the main document part

header = next(b for b in pkg.bodies() if b.prefix == "header:rId7")
header.replace_text("Normal", "Ordinary")                          # one header, by its prefix
```

### 3.2 Headers and footers only

```python
for body in pkg.bodies():
    if body.prefix.startswith(("header:", "footer:")):
        body.replace_text("header", "HEADER", match_case=True)
```

```
header:rId8   1   'First page HEADER.'
header:rId13  0   'ODD/even … BLACK.'        (case did not match)
header:rId7   1   'Normal HEADER.'
header:rId12  0   'Odd/EVEN.. white'
header:rId11  1   'First page HEADER.'
header:rId10  1   'Just a plain HEADER.'
```

The same document can reach one header part from several sections (`rId8` and `rId11` here are
two parts with the same text, not one part twice), and `pkg.bodies()` yields each **part** once,
so a replacement is never made twice.

### 3.3 Notes

`footnotes` and `endnotes` are bodies like the others; a footnote's text is a paragraph of the
footnotes part, so `body.prefix == "footnotes"` is the filter. The reference marks in the main
text (`w:footnoteReference`) carry no text and are never matched.

### 3.4 Everything

```python
pkg = load("Headers.docx")
counts = {body.prefix: body.replace_text("header", "heading") for body in pkg.bodies()}
# {'body': 1, 'header:rId8': 1, 'header:rId13': 0, 'header:rId7': 1, 'header:rId12': 0,
#  'header:rId11': 1, 'header:rId10': 1, 'footnotes': 0, 'endnotes': 0}
pkg.save("out.docx")
```

Each part that had a match is re-marshalled on save and every other part is written back byte
for byte, headers included --- `replace_text` on a header that matches nothing still **reads**
its `contents`, which marks it for re-marshalling (CR-002's rule: no dirty tracking), so the
output of the loop above has every header re-serialised even where the count is 0. That is
canonically identical output, not a change; if byte-identity of untouched headers matters,
`search` first and replace only where there is a hit.

### 3.5 Tracked

With tracking on, every replacement is a `w:del` / `w:ins` pair under `pkg.author`, in every
part the loop reaches:

```python
pkg = load("Headers.docx")
pkg.change_tracking_mode = "TrackAll"
sum(body.replace_text("header", "heading") for body in pkg.bodies())      # 5

for body in pkg.bodies():
    if body.get_tracked_changes():
        print(body.prefix, [(c.type, c.text) for c in body.get_tracked_changes()])
```

```
body          [('Deleted', 'Header'), ('Added', 'heading')]
header:rId8   [('Deleted', 'header'), ('Added', 'heading')]
header:rId7   [('Deleted', 'header'), ('Added', 'heading')]
header:rId11  [('Deleted', 'header'), ('Added', 'heading')]
header:rId10  [('Deleted', 'header'), ('Added', 'heading')]
```

As with `find`, **`pkg.get_tracked_changes()` is the main part's**; each body has its own, and
`accept_all()` / `reject_all()` are per body too.

### 3.6 Trying it first

`dry_run` gives a trial whose `bodies()` are copies, with the same loop:

```python
with pkg.dry_run() as trial:
    n = sum(body.replace_text("header", "heading") for body in trial.bodies())    # 5
    trial.body.text                                                              # what it would look like
pkg.body.text                                                                    # unchanged
```

## 4. Where a match is not found, and what a replacement does not do

- **Text boxes.** A text box's paragraphs live inside a shape inside a run (`w:txbxContent`), and
  the content API does not descend into them (CR-003 decided question 8). A `find` over every
  body still does not see them; the typed tree does (`walk` from the part's contents with
  `text_of`), and so does `to_markdown` only where a shape is rendered.
- **Field instructions.** `w:instrText` --- the `HYPERLINK "..."`, `PAGE`, `TOC \o` inside a
  field --- is not paragraph text; the field's *result* is, and is matched and replaced like any
  text. Word recomputes the result at its next field update, so replacing inside a `PAGE` field's
  result is undone by Word.
- **Content controls and tables** are descended into: `body.paragraphs` includes the paragraphs
  of every `w:sdt` and every cell.
- **A match that crosses a hyperlink's boundary** ("on Slashdot", where only `Slashdot` is the
  link) is replaced, but the replacement is written as one run **outside** the link and the link
  is left empty (`<w:hyperlink r:id="rId4"/>`): the text is un-linked. A match *inside* the link (`Slashdot` →
  `slashdot.org`) keeps it. Section 3 of [hyperlinks.md](hyperlinks.md) has the related note on
  inserting after a link; if the empty element matters, search first and refuse a `Range` whose
  text crosses into a link (`range.paragraph.runs` shows the holders).
- **Comments** are a body (`comments`), so a replacement over "everything" reaches the comment
  text too; filter the prefix out if that is not wanted.
- **The python-docx facade** has no find or replace, because python-docx has none. From a facade
  `Document`, the package is `document._package`, and the loop above runs on it.

## 5. If a verb is wanted

A `pkg.replace_text(find, replace, parts="all" | "body" | "headers" | ...)` that runs the loop,
returns the per-prefix counts and records one `ChangeReport` naming every part is about an hour
over what is here, and is the shape CR-006's `edit` tool would call for a document-wide
replacement. It is not in a CR; the loop is the documented way until it is.
