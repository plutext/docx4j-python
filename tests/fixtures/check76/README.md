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

**Copied unedited** on 2026-11-06 for CR-037. `76t` to `76t28`, at docx4j-core-ts `1ffb120`: the 29 sources from the share's
`__tmp/76/`, each entry of each the same as `node test/fixtures/check76/build.mjs`'s output, and the 28 saves from core-ts's
`test/check76/`. `76t29` to `76t67`, rounds 7 to 10: the 39 sources from `build.mjs`, each the same as the input the Word bridge
ran, and the 30 saves from `test/check76/`. All 63 saves here match its `manifest.json`. What each file holds and what Word read
is its CR-026 sections 4.4 to 4.7.

| Files | `w:madeUp` | Word 365 |
|---|---|---|
| `76t`, `76t1` to `76t4`, `76t7` to `76t17`, `76t19` to `76t25`, `76t29` to `76t33`, `76t35`, `76t36`, `76t51` | holding what its host admits, in one of twenty content hosts, nested in some | reads what it held, the tags gone |
| `76t5`, `76t6` | holding a property, in a `w:pPr` or a `w:rPr` | drops it |
| `76t18`, `76t26`, `76t34`, `76t52`, `76t60` | none: controls, every host with the words standing bare | as the files above |
| `76t27`, `76t37` to `76t40`, `76t50` | holding a run, in a `w:pPr`, the `w:pPr` under a `w:pPrChange` or a paragraph's `w:sectPr` | reads it at the head of the paragraph; a `w:sectPr` goes |
| `76t41`, `76t42`, `76t61` to `76t67` | text, a `w:br`, in a host | the text gone, the break read, a run beside the text read |
| `76t43`, `76t56`, `76t58` | a `w:p`, last in a `w:p` with a run before it | it ends the host, which takes its properties; a further one a paragraph of its own; an empty `w:p` after |
| `76t44`, `76t45` | a run, directly in the `mc:Choice` or `mc:Fallback` taken | reads it |
| `76t28`, `76t46` to `76t49` | a run, in a `w:rPr`, a block `w:sdtPr`, `w:trPr`, `w:tcPr`, `w:tblPr` | refuses the file, so there's no save |
| `76t53` to `76t55`, `76t57`, `76t59` | a `w:p` in a `w:p`, any other shape | refuses the file, so there's no save |

Saves from round 6 on are Word bridge runs' (Word 16.0.20326.20232), so compare their stories and not the whole file: the manifest says
how a bridge save differs from a hand save.

`SHA1SUMS` lists the 136 files; this README is not among them.
