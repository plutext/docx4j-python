# Word's reading of an `mc:AlternateContent` that selects nothing, and of what it understands

CR-040's oracle (`tests/openpackaging/test_nothing_selected.py`).

**Copied unedited** on 2026-11-07 from the Word bridge's share, `bridge/portfolio/`, the `-a` run of each job, on Word 16.0.20326.20232. Each input ran twice, and the two runs agreed (the portfolio's `text_extraction_behaviour_model.md` section 7.15).

| Files | What they hold |
|---|---|
| `rw-25-*` | A `Choice` in a namespace Word doesn't know, with no `Fallback`, in a body |
| `c305-*` | The `req305` check: three namespaces Word doesn't know, and namespaces it does, at block and run level |
| `uc-*` | The `understood-choice` round: namespaces the old list lacked |
| `*-word365.docx`, `*-word365.txt` | Word's save and text save of each |

`SHA1SUMS` holds every file here.
