# docx4j-python adapter for docx-platform-tests

docx4j-python's adapter for [open-agreements/docx-platform-tests](https://github.com/open-agreements/docx-platform-tests),
adapter protocol version 1: one operation on one document, as the suite's runner calls it. It is one script,
`run_adapter.py`, using the standard library and `docx4j_py` (PyPI distribution `docx4j`), so it runs under any
interpreter that has `docx4j` installed. It lives outside `docx4j_py/`, so it is not in the wheel or the sdist.

In the suite it sits at `adapters/docx4j-python/`. `install.sh` makes `adapters/docx4j-python/.venv` with
`docx4j==0.2.0` from PyPI (Python 3.12 or later). It is registered with `registry-fragment.json`:

```json
{
  "adapterName": "docx4j-python",
  "adapterCommand": ["adapters/docx4j-python/.venv/bin/python", "-I", "adapters/docx4j-python/run_adapter.py"],
  "adapterVersionCommand": ["adapters/docx4j-python/.venv/bin/python", "-I", "-c", "import importlib.metadata as m; print(m.version('docx4j') + ' (docx4j-xsdata ' + m.version('docx4j-xsdata') + ')')"]
}
```

Against a checkout instead of a release, put the checkout on the path:
`PYTHONPATH=<checkout> <its venv>/bin/python -P run_adapter.py --protocol-version 1 --operation ... --input ... --output ...`.

## What is library API and what is glue

docx4j-python is a typed tree over the package plus a content API modelled on Office JS (`Body`, `Paragraph`,
`Range`, `Table`, `Comment`, lists) and a python-docx facade (`docx4j_py.docx`). The adapter calls that API. Its own
code parses the arguments, finds the target, converts units, saves, and declines inputs the API cannot represent.

The glue every operation shares:
- "The first paragraph containing the text" is the first hit of the content API's `pkg.body.search(text,
  match_case=True, limit=1)`. The search runs over the main document's paragraphs in document order, inside
  tables and content controls too. A match is found on a paragraph's whole text, so it may cross runs.
- `^` in the text is passed as `^^`, Word's escape, so that the text is matched literally and not read as one of
  Word's Find codes.
- `tableIndex` is `pkg.body.tables[i]`: the body's own tables, as Office JS `Body.tables` lists them.
- `fontSizeHalfPoints` is halved, because the content API's `Font.size` is in points.
- Every output is written by the package's default `save()`. It re-marshals only the parts that were read and
  writes them in the normal form Word reads.

| Operation | docx4j-python |
|---|---|
| `acceptAllTrackedChanges`, `rejectAllTrackedChanges` | library: `docx4j_py.model.revisions.accept_all(pkg)` / `reject_all(pkg)`, Word 365's Accept All and Reject All over every story (CR-010) |
| `replaceFirstTextOccurrence` | library: the first hit's `insert_text(replaceText, location="Replace")`. A match across runs is replaced, in the first replaced character's formatting |
| `formatFirstTextOccurrence` | library: the first hit's `font.bold`, `font.italic`, `font.size`. The runs are split at the hit's edges |
| `addCommentOnFirstTextOccurrence` | library: `pkg.author = Author(name, initials=...)`, then the first hit's `insert_comment(text)`. The comment parts, their relationships and content types are created where the document has none |
| `removeAllComments` | library: for each body in `pkg.bodies()` (the main document, headers, footers, notes), `delete()` on each of `get_comments()`. That removes markers, reference runs and entries, replies included |
| `composeDocumentWithParagraphs` | library: `create_package()`, then `body.insert_paragraph(text)` and that paragraph's `font` |
| `composeDocumentWithTable` | library: `body.insert_table(rows, columns, values=...)`. A ragged matrix is declined: the table it makes is rectangular |
| `composeDocumentWithHyperlink` | library: `body.insert_paragraph(text)`, then its `get_range("Content").hyperlink = url` (Office JS `Range.hyperlink`: an external relationship from the part, a `w:hyperlink` naming it, and the `Hyperlink` character style on its runs) |
| `composeDocumentWithNumberedList` | library: `start_new_list()` on the first paragraph, `attach_to_list(id, level)` on each paragraph, and `List.set_level_numbering(level, ...)` for each level used. Glue: `numberFormat` is mapped to Office JS's `Word.ListNumbering` by inverting the library's own table (`docx4j_py.model.content.lists.NUMBERING`); `bullet` is `start_new_list(kind="Bullet")`. Any other format is declined |
| `composeDocumentWithCompatibilityMode` | library: `pkg.compatibility_mode = n`. The library writes 11, 12, 14 and 15; any other mode is declined |
| `composeDocumentWithHeaderText`, `setDefaultFooterText` | library: the facade's `Document`, the last section's `header` / `footer`, and `is_linked_to_previous = False`, which creates the part, its relationship, its content type and the `w:sectPr` reference. Then the part's `body.insert_text(text, location="Replace")` (an existing default header or footer has its content replaced). Glue: where `w:body` has no `w:sectPr`, an empty one is set on it in the typed tree first. That is what Word assumes such a body has |
| `appendParagraphWithText` | library: `body.insert_paragraph(text, location="End")` and its `font` |
| `insertParagraphAfterAnchorText` | library: the first hit's paragraph's `insert_paragraph(text, location="After")`. The new paragraph takes the anchor's paragraph properties, as Word's split does |
| `appendTableRow` | library: `table.add_rows(1, values=[cellTexts])`, using the table's columns and widths. More texts than the grid has columns is declined |
| `deleteTableRowAtIndex` | library: `table.delete_rows(rowIndex, 1)` |
| `setTableCellText` | library: `table.cell(rowIndex, columnIndex).value = text` |
| `applyParagraphStyleToAnchor` | library: the first hit's paragraph's `style = paragraphStyleId`. An id or a name the document defines is applied. A built-in Word style it lacks is defined first. Anything else is refused, which is an error |
| `applyNumberingToAnchorParagraph` | library: the first hit's paragraph's `attach_to_list(numberingInstanceId, indentationLevel)`. As Word does, this also gives *List Paragraph* to a paragraph that has no style of its own |
| `mergeTableCellsInRow` | **declined**: docx4j-python has no cell merge. The content API adds and removes rows and columns only, and the facade's `_Cell.merge` raises `NotImplementedError` (CR-003 section 20) |

Exit codes follow the protocol:
- 0: the output is written.
- 1: an error, with nothing written. A missing text, table or row is one line on stderr; any other failure is the traceback.
- 2: declined, with the reason as one line on stdout and nothing written.
- 3: another protocol version.

The library logs warnings on stderr, for example docx4j's `No default paragraph style!!` for a package with no
styles part. They are not part of the protocol.

## Honest gaps

- `mergeTableCellsInRow` is declined (above).
- In `removeAllComments`, a comment that no story anchors is not listed by `get_comments()`, so its entry stays in
  `comments.xml`.
- The parse does not keep an `mc:Ignorable` extension element inside `w:sdtPr`, or an extension attribute on `w:sdt`:
  both are reported in `part.skipped`. So once the main document part is edited and saved, they are gone (the
  suite's opaque content-control invariants). The namespace declaration and `mc:Ignorable` are kept. This is the
  library's behaviour, and the adapter does not work round it.

**Maintenance policy:** best-effort, by the docx4j-python project (Plutext). `install.sh` pins a released
`docx4j`. A new release is measured by changing that pin, and pip resolves `docx4j-xsdata` within the range that
release declares. The adapter calls only public `docx4j_py` API, so a break shows as `error` cells, not a silent
change.
