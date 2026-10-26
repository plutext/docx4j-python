# Integer attributes: the four unmeasured edges, as Word 365 reads them (CR-021 phase C)

CR-021 phase C's Word check, run by the user in Word 365 on 2026-10-26. Built by
`out/cr021-word-check/build.py` at docx4j-python `7fe0be8`, staged at
`vbShares/Office 2016/__tmp/python-cr021/`, and copied from there read-only. `SHA1SUMS` holds
each file's digest as copied.

| File | What it is |
|---|---|
| `cr021-edges-raw.docx` | the four edges together, as written (not saved by docx4j-python) |
| `cr021-edge-E<n>-raw.docx` | one edge each: E1 `w:spacing/@w:before="239.99999999999999"`, E2 `w:spacing/@w:before="."`, E3 `w:spacing/@w:before="+"`, E4 `w:ind/@w:left="-"` |
| `*-word365.docx` | Word 365's Save As of each |

What Word did, all five files opened, no repair:

- E1 saved as `w:before="239"`: Word truncates the string (a double would read 240), as
  `values.integer` and docx4j's `LenientInteger` do;
- E2, E3 and E4 opened, so they are not refused forms; Word's saves hold no `w:spacing` and no
  `w:ind` for them. Word leaves out direct formatting equal to the style's, and the style's
  value here is 0, so these files can't tell "read as 0" (this package) from "ignored" (docx4j
  CR-065 phase 2, which drops them). The two differ only under a style whose value is not 0.

The same check's `cr021-edges.docx` (not kept here: our normalised save, with `w:pgMar/@w:top="abc"`
dropped) opened with Word's Normal margins, a 2.54 cm top.
