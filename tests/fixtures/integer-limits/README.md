# Integer measurements at and beyond Word's range (CR-022)

docx4j CR-066's phase 0: rounds five and six of the shared Word probe, integer attributes
at and beyond 16 bits, and Word 365's saves of every file. Built by docx4j's session
docx4j-25 (round five at `1be8b6515`, local; its scratchpad generators `probe66/Probe66.java`
and `probe66/Probe66r6.java`) and run by the user in Word 365 on 2026-10-26: every group file
opened clean, with no repair offered, in both rounds. The analysis is docx4j's
`docs/developer/change-requests/CR-066-integer-attributes-within-words-range.md`, sections
3.1 (round five) and 3.4 (round six), and CR-022 sections 3.1 and 3.3 here. Copied read-only
from the share, `vbShares/Office 2016/__tmp/cr066/`, on 2026-10-26; `SHA1SUMS` holds each
file's digest as copied.

| File | What it is |
|---|---|
| `round5/cases.txt` | round five's 191 cases: id, group, attribute (a context in brackets: `w:spacing/@w:line[exact]`), value, and the value if Word wraps at 16 bits or clamps |
| `round5/README.txt` | the share's notes for round five: the groups, the steps, the result |
| `round5/cr066-<group>.docx` | a group's cases, `P` paragraph, `B` borders, `R` runs, `T` tables, `S` sections, `N` negatives in unsigned attributes |
| `round5/cr066-<group>-word365.docx` | Word 365's Save As of that group |
| `round6/cases.txt` | round six's 489 cases: id, group, attribute, place (`pPr`, `rPr`, `tbl`, `sectPr`, `numbering:<numId>`, `settings`), value, and the value if Word wraps signed or unsigned |
| `round6/README.txt` | the share's notes for round six |
| `round6/cr066-r6-<group>.docx` | a group's cases, `A` paragraph, `F` frames, `R` runs, `T` tables, `P` floating tables, `S` sections, `N` numbering, `E` the digitless forms (CR-021's) |
| `round6/cr066-r6-<group>-word365.docx` | Word 365's Save As of that group |
| `round6/cr066-r6-D1-<n>.docx`, `-D2-<n>.docx` | one value per file of `w:defaultTabStop` (`D1`) and `w:hyphenationZone` (`D2`), which are document-wide, with Word's save of each beside it (`-word365.docx`) |

Each case paragraph is labelled `<id> <attribute>="<value>"`. The label paragraph itself
carries the value for a paragraph property; in a run group the run after `then:` (character
style `ProbeRun`) carries it; in a table group the table after the label; in a section group
the section the paragraph after the label ends. Styles `ProbeBase` and `ProbeRun` give every
tested property a non-zero inherited value, so a value Word wraps to 0 is written as 0 in its
save, and one Word drops is absent.

Not copied: the share's `families/` and `singles/` fallbacks, which no round needed (Word
opened every group), except round six's `D1` and `D2` singles, which are the measurement;
`families.txt`, whose content `cases.txt` carries; and Word's `~$` lock files.

`docx4j_py/resources/integer-limits.json`, the limits table (CR-022 D1), cites these cases by
id, and `tests/content/test_integer_limits.py` holds every row to Word's saves: a value inside
the row's range is kept, one outside is changed.
