# docx4j CR-073's `br-placement` probe: a `w:br` outside a `w:r`, and Word 365's saves

CR-030's oracle (`tests/openpackaging/test_br_outside_run.py`).

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/br-placement/` at `95c6b42d3` (docx4j
CR-073 phase 1, on `VERSION_17_3_6`; Apache-2.0; the git blob ids are docx4j's), by
`git archive 95c6b42d3 <dir> | tar -x` (CR-030 Phase A, 2026-11-03). Rounds 1 and 2 of the probe are docx4j's
`a2a971416` and `61147016e`.

| File | What |
|---|---|
| `br-1-p-alone.docx` to `br-8-p-page.docx` | round 1: a `w:br` alone in a paragraph, between runs, in a `w:t`, between `w:t` nested in a `w:t`, in a hyperlink, in a cell, in a footer, and a page break directly in a paragraph |
| `br2-1-run-properties.docx` to `br2-5-body-br-page.docx` | round 2: the run properties Word gives the break (R1 to R8), a break in a field's `w:instrText`, in a bare `w:instrText` in a header, and directly in `w:body` (a line and a page break) |
| `*-word365.docx` | Word 365's save of each (Jason's runs, 3 November 2026), opened with no repair reported |
| `docx4j-README.md` | docx4j's README of the probe, unedited: each case, what Word's save and its PDF show, and what docx4j 17.3.5 read |

The two builders, `build_br_placement.py` and `build_br_placement_2.py`, stay in docx4j, and Word's PDFs on the share,
as docx4j's README says. `SHA1SUMS` lists the 27 files copied; this README is not among them.
