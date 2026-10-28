# docx4j CR-067 phase 1's field probe, and docx4j's extraction of it

CR-025 Phase B's oracle (`tests/text/test_field_probe.py`).

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/fields/cr067-phase1/` at
`d6ca8dbee` (Apache-2.0; the git blob ids are docx4j's):

| File | What |
|---|---|
| `field-probe-r1.docx` | docx4j's `FieldProbeBuilder` probe: 25 cases, every tier-1 field with the cached result `OLD` (`OLDUSER` for the user fields); core properties created `2024-02-03T04:05:06Z`, modified `2025-06-07T08:09:10Z`, revision 7; a bold bookmark `bm1`, a bookmark `bmNote` on a footnote reference, a section break before cases 24 and 25 |
| `field-probe-r1-word365.docx` | Word 365's save after Ctrl+A, F9 (Jason, 2026-10-28, version 2608 build 20326.20232, the machine in Sydney's zone) |
| `results.txt` | Jason's record of that run, the README's checklist with each case's result; `FieldProbeWordTest` parses its `Results:` section |

**docx4j's own output, generated here** on 2026-10-28 by `ProbeText.java` and `ProbeMd.java` (beside this file)
against docx4j's build of `c3e9920e5` (`docx4j-core`, `docx4j-JAXB-ReferenceImpl` and `docx4j-markdown`
`17.3.5-SNAPSHOT` jars from `target/`, the JVM in `Australia/Sydney`, `-Duser.timezone=Australia/Sydney`, logging off):

| File | docx4j call |
|---|---|
| `docx4j-update.txt` | `Texts.plainText(probe, fields: update)` |
| `docx4j-update-tokyo.txt` | the same with `pkg.setDateTimeZone(ZoneId.of("Asia/Tokyo"))` |
| `docx4j-cached.txt` | `Texts.plainText(probe, fields: cached)` |
| `docx4j-word365-cached.txt` | `Texts.plainText(Word's save, fields: cached)` |
| `docx4j-update.md` | `new MarkdownExporter().export(Texts.of(probe, fields: update), probe)` |

`SHA1SUMS` lists every file.
