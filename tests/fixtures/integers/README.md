# Integer attributes as Word reads them (CR-021)

docx4j CR-065's phase 0 probe: decimals and odd forms in WordprocessingML integer
attributes, and Word 365's saves of the files it opened. Built by docx4j's session
(scratchpad `probe65/Probe65.java`) and run by the user in Word 365 on 2026-10-26, rounds
two to four; recorded in docx4j's
`docs/developer/change-requests/CR-065-integer-attributes-read-as-word-reads-them.md`
section 3, at docx4j commit `aacaa4c45` (round four, phase 0 done). Copied read-only from
the share, `vbShares/Office 2016/__tmp/cr065/` (`families/*.docx`, `cases.txt`,
`README.txt`), on 2026-10-26; `SHA1SUMS` holds each file's digest as copied.

| File | What it is |
|---|---|
| `cases.txt` | every case: family, id, attribute, value (65; the ids are the regenerated probe's, one higher from `B20` on than the first design, after `A19` was added) |
| `README.txt` | the share's own notes: the families, the steps of each round, and each round's result |
| `cr065-<family>.docx` | a family's cases, each labelled by its id, with the control paragraph `Z65` (`w:spacing/@w:before="120"`) |
| `cr065-<family>-word365.docx` | Word 365's Save As of that family, where Word opened it |

Families: `A` `w:spacing/@w:before` in every form (refused whole in round two, so split
in round three into `A1` the decimals, `A2` the six forms, `A3` the comma, `A4` hex, `A5`
the exponents, `A6` the empty string and `A7` letters, and `A2` again in round four into
`A2a` to `A2f`, one form each); `B` `w:ind/@w:left`, signed; `C` attributes on docx4j's
rounding list; `D` borders; `E` tabs; `F` run properties; `G` tables; `H` sections, with
SSRS 2012's empty attributes; `I` a style's indent, `w:defaultTabStop` and a header.

What Word did:

- **opened, and saved:** `A1`, `A2a` to `A2d`, `A6`, `B` to `I` (each `-word365.docx`);
- **refused the document** ("Word experienced an error trying to open the file", no repair
  offered), so no save: `A` and `A2` (as groups), `A2e` (`" 240.5 "`), `A2f` (`" 240 "`),
  `A3` (`"240,5"`), `A4` (`"0x10"`), `A5` (`"1.5e2"`, `"2.5e0"` and `"1e3"`, refused as a
  group) and `A7` (`"abc"`).

Word's saves of `B` to `I` are round two's, made before the probe was regenerated at
08:35: in `B` to `H` the first generation carried sentinel integers (7205001, 7205002,
1205003) in `styles.xml`, `settings.xml` and the header, which Word kept modulo 2^16
(docx4j CR-065 section 3's aside, CR-022). The sources here are the regenerated files,
whose parts outside the cases are plain integers; the cases' paragraphs are the same in
both, so a test compares only the cases.

Word writes direct formatting as a difference from the style, so a value equal to the
style's after truncation is not in Word's save: `A08` (`".5"`), `A17` (`""`) and `B25`
(`"-0.5"`) read 0, and `F41` (`w:sz="22.5"`) 22, the document default. Word recomputes a
table's grid on open, so `G50` and `G51` (`w:gridCol`) have no reading a save shows.
