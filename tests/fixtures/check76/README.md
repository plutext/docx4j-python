# core-ts's Word check 76: an element Word doesn't know inside a `w:t`, and Word 365's saves

CR-030 section 11's oracle (`tests/openpackaging/test_br_outside_run.py`).

**Copied unedited** on 2026-11-05 from docx4j-core-ts's check 76 (its CR-026 section 4, CR-021 section 11; `33bae48`).
The sources are `test/fixtures/check76/build.mjs`'s, from the share's `__tmp/76/`; Word's saves are core-ts's
`test/check76/`, each matching its `manifest.json`.

| File | What is inside the middle paragraph's `w:t` | Word 365 reads, with no message |
|---|---|---|
| `76m-ignorable-in-t.docx` | an empty `ext:opaque`, `ext` in `mc:Ignorable` | `TARGET TEXT` |
| `76n-ignorable-with-text-in-t.docx` | `ext:opaque` holding `HIDDEN WORDS`, ignorable | `TARGET TEXT`, the words gone |
| `76u-madeup-with-text-in-t.docx` | `w:madeUp` holding `HIDDEN WORDS` | `TARGET TEXT`, the words gone |
| `76w-ignorable-holding-a-break-in-t.docx` | `ext:opaque` holding a `w:br`, ignorable | `TARGET TEXT`, no break |
| `76z-tab-in-t.docx` | a `w:tab` | `TARGETTEXT`, no tab |
| `*-word365.docx` | Word 365's save of each (Jason's runs, 2026-11-05) | |

Word could not repair `76v`, `76x` or `76y`, so they have no save here, and the tests hold those rows in inline XML.
`SHA1SUMS` lists the ten files; this README is not among them.
