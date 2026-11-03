# docx4j CR-046 section 10's `part-not-read` fixture: a styles, numbering or settings part that cannot be read

CR-031's oracle for `part.not_read` (WD 0.66 REQ-801; `tests/text/test_part_not_read.py`).

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/text-extraction/part-not-read/` at
`3ad164bdf` (docx4j CR-046 section 10, on `VERSION_17_3_6`, local and unpushed on 2026-11-03; Apache-2.0; the git
blob ids are docx4j's), by `git archive 3ad164bdf <dir> | tar -x` (CR-031 Phase A, 2026-11-03). It is kept here, not
under `../text-extraction/`, whose copy is pinned at the `docx4j-17.3.5` tag (CR-027 Q6).

| File | What |
|---|---|
| `part-not-read-intact.docx` | ten paragraphs: two `Heading1` numbered by their style ("Article I"), two numbered directly ("1."), a run hidden by a character style, a run in capitals by a character style and one by direct `w:caps` (no `w:lang`, so the settings' `themeFontLang` tr-TR applies), a paragraph in small capitals by its style, a run hidden directly, an insertion and a deletion |
| `part-not-read-styles.docx`, `-numbering.docx`, `-settings.docx` | the same, that part not well formed: cut to two thirds of its characters |
| `part-not-read-styles-absent.docx`, `-numbering-absent.docx`, `-settings-absent.docx` | the same, that part absent (part, relationship and override removed) |
| `part-not-read-numbering-unneeded.docx` | two plain paragraphs, the numbering part not well formed: nothing needs it |
| `expected/<fixture>.<view>.txt` | docx4j's plain text in each view (`accepted`, `original`, `markup`), with `list_labels: true` and `small_caps: upper` |
| `expected-warnings.tsv` | the warnings in each view: kind and selector, `; ` between two |
| `docx4j-README.md` | docx4j's README of the fixture, unedited |

The builder, `build_part_not_read.py`, stays in docx4j, as `../br-placement/`'s do. `SHA1SUMS` lists the 34 files
copied; this README is not among them.

**Pending here:** the three case-mapping runs of each fixture whose expected text maps them (`w:caps` by a style
and directly, by the run's language, the settings' `themeFontLang` tr-TR giving "İSTANBUL", the default language
"ISTANBUL"; `w:smallCaps` under `small_caps: upper`) wait for CR-029, which maps case in the text model. This
package reads them as stored. Every other line, and every warning, is docx4j's.
