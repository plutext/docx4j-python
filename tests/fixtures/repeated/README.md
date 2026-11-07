# Word's reading of a repeated child and of a main part's root

CR-039's oracle (`tests/openpackaging/test_repeated.py`).

**Copied unedited** on 2026-11-07 from the Word bridge's share, `bridge/portfolio/`, the `-a` run of each job, on Word 16.0.20326.20232. Each input ran twice, and the two runs agreed (the portfolio's `text_extraction_behaviour_model.md` section 7.15; core-ts's CR-028 section 1.1). `or-14-comments-a` was killed at a dialog the bridge didn't yet know, so that file's save is `-a2`'s, its rerun. The `rc-*` and `rw-*` inputs are the objects session's, the `or-*` inputs core-ts's.

| Files | What they hold |
|---|---|
| `rc-*.docx` | A child given twice in `word/document.xml`, with three controls |
| `rw-*.docx` | A root that isn't `w:document`, `rw-2*` `mc:AlternateContent`, with three controls |
| `or-*.docx` | Another story's root, with a control |
| `rt-15` to `rt-17` | Three `w:body`: which section stands. The objects session's, added 2026-11-07 |
| `mr-*` | docx4j's `madeup-root` round: how far Word reads an unknown root |
| `nd-20` to `nd-24` | A `w:body` in a `w:body`. The objects session's |
| `sm-*` | The `sectpr-merge` round: what Word's rule for a repeated child is, one difference a file (CR-039 section 9.5). Saves only, no text saves |
| `ft-*` | The `footer-refs` round: a `w:footerReference` or `w:headerReference` on a later copy of a `w:sectPr`. Saves only |
| `*-word365.docx` | Word's save of each. A file Word refused has none |
| `*-word365.txt` | Word's text save. It leaves text boxes out, so a box's words are read from the save |

CR-039 sections 2.1, 2.2 and 2.4 say what Word read from each. `SHA1SUMS` holds every file here.
