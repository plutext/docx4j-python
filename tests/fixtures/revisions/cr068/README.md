# docx4j CR-068's oracle: a manual redline as tracked changes, and Word's dates

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/revisions/cr068/` at docx4j's release tag
`docx4j-17.3.5` (`64fae4cd0`; Apache-2.0; the git blob ids are docx4j's), once, by
`git archive docx4j-17.3.5 <dir> | tar -x` (CR-027 Phase B, 2026-11-03). This is the one copy of that directory
here: CR-027 Phase B reads `manual-redline-dates.docx`, CR-028 the rest (CR-028 Q11, decided 2026-11-02).

54 files, every one there but the four screenshots, which are left in docx4j (Q11): `cr068-comment-cards-pane.PNG`,
`cr068-comment-cards.PNG`, `cr068-default.PNG` and `cr068-optins.png`. `SHA1SUMS` lists the 54; this README is not
among them. The `.txt` files are Word's Save As text (CR LF lines, a bare CR between cells), and `.gitattributes`
keeps the whole directory as copied.

| File | What |
|---|---|
| `manual-redline-dates.docx` | CR-068 phase 0's date probe (docx4j's `ManualRedlineProbeBuilder`): four deletions by `manual redline`. D1 has `w:date="2026-10-31T00:00:00Z"` alone, D2 `w16du:dateUtc="2026-10-31T00:00:00Z"` alone, D3 both at `00:00:00Z`, D4 `w:date="2026-10-31T11:00:00Z"` with `dateUtc` at `00:00:00Z`, as the conversions write them. Word's Reviewing Pane showed D1 and D3 at midnight, D4 at 11:00 and D2 with no date (`cr068-log.txt`, `date 1899-11-30`: Word's empty date) |
| `manual-redline.docx`, `manual-redline-2.docx` | the two bases, a manual redline with no tracking (cases M01 to M17, N01 to N10) |
| `manual-redline-default.docx`, `-optins.docx`, `-d9.docx` | the conversions docx4j predicted before phase 1 |
| `manual-redline-phase1.docx`, `-phase2.docx`, `manual-redline-2-phase2.docx` | docx4j's own conversions at phases 1 and 2 |
| `manual-redline-word.docx` | Word's own tracked deletion of the base's struck text |
| `word-<key>-asis.txt` | Word's Save As text of each file as it is |
| `word-<key>-accept-all.docx`, `.txt` | Word 365's Accept All of each, as a document and as text |
| `word-<key>-reject-all.docx`, `.txt` | Word 365's Reject All of each, likewise |
| `comment-cards.docx` | the comment cards probe: five comments, C1 to C5, each holding one form of the same change |
| `cr068-probe.ps1`, `RUN.txt` | the script that drove Word (Word 16.0 build 16.0.20326) and the instructions for each run |
| `cr068-log.txt` | the script's log: Word's version, then every revision Word listed, story by story |

The keys are `base`, `default`, `optins`, `d9`, `own` (Word's own deletion), `phase1`, `phase2` and `phase2b` (the
second base).
