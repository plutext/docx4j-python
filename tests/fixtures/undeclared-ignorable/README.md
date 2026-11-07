# Word's reading of a prefix `mc:Ignorable` names with no declaration

WD REQ-323's oracle (`tests/openpackaging/test_undeclared_ignorable.py`; CR-040 section 9).

**Copied unedited** on 2026-11-08 from the Word bridge's share, `bridge/portfolio/`, the `-a` run of each job of the
portfolio's `undeclared-ignorable` round, on Word 16.0.20326.20232. Each input ran twice, and the two runs agreed: the
same input, the same outcome, the same text save.

| Files | What they hold | Word |
|---|---|---|
| `ui-00-control` | Every prefix declared | silent |
| `ui-01` to `ui-04` | The main part's root lists a prefix it doesn't declare: unknown (`zz9`), known (`w14`), several (`w15 wp14`), and `w14` declared lower, on a paragraph using it | repaired, losing nothing |
| `ui-05` | `zz9` undeclared on the root, and an element declaring it for itself | refused |
| `ui-06` | `zz9` declared on the root, its element ignored | silent, the element and its text gone |
| `ui-07` | A leading space in the list | silent |
| `ui-10` to `ui-13` | `zz9` undeclared on a header's, footnotes', styles' or settings' root | repaired, losing nothing |
| `ui-14` | `w14` undeclared on all five | repaired, losing nothing |
| `*-word365.docx` | Word's save of each. A file Word refused has none | |
| `*-word365.txt` | Word's text save | |

`SHA1SUMS` holds every file here but this one.
