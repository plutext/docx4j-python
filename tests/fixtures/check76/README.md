# core-ts's Word check 76: an element Word doesn't know inside a `w:t`, and Word 365's saves

CR-030 section 11's oracle (`tests/openpackaging/test_br_outside_run.py`), and CR-037's (`tests/openpackaging/test_read_through.py`).

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

## The `76t` files: an unknown `w:` element outside a `w:t`

**Copied unedited** on 2026-11-06 for CR-037, at docx4j-core-ts `1ffb120`. The 29 sources are from the share's `__tmp/76/`, each
entry of each the same as `node test/fixtures/check76/build.mjs`'s output. The 28 saves are core-ts's `test/check76/`, each
matching its `manifest.json`. What each file holds and what Word read is its CR-026 section 4.4.

| Files | `w:madeUp` | Word 365 |
|---|---|---|
| `76t`, `76t1` to `76t4`, `76t7` to `76t17`, `76t19` to `76t25` | holding what its host admits, in one of fifteen content hosts | reads what it held, the tags gone |
| `76t5`, `76t6` | holding a property, in a `w:pPr` or a `w:rPr` | drops it |
| `76t18`, `76t26` | none: controls, every host with the words standing bare | as the files above |
| `76t27` | holding a run, in a `w:pPr` | reads `HIDDEN WORDSTARGET TEXT`, one run |
| `76t28` | holding a run, in a `w:rPr` | refuses the file, so there's no save |

`76t27`'s save is a Word bridge run's (Word 16.0.20326.20232), so compare its body and not the whole file: the manifest says how a
bridge save differs from a hand save.

`SHA1SUMS` lists the 67 files; this README is not among them.
