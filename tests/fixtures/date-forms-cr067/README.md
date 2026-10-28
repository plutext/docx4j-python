# docx4j CR-067's date-forms probe

CR-025 Phase B2's oracle (`tests/text/test_date_forms_probe.py`).

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/fields/cr067-dates/` at
`c58e36e64` (Apache-2.0; the git blob ids are docx4j's):

| File | What |
|---|---|
| `date-forms-r1.docx` | docx4j's `DateFormsProbeBuilder` probe: three stored forms, a date alone (`2024-02-03`), a time with no offset (`2025-06-07T08:09:10`) and a UTC time (`2025-01-02T03:04:05Z`), each in a core property (created, modified, lastPrinted, read by `CREATEDATE`, `SAVEDATE` and `PRINTDATE`, D01 to D06, with `\@ "yyyy-MM-dd HH:mm:ss"` and without a switch, cached `OLD`), a revision's `w:date` (three `w:ins`, R01 to R03) and a date control's `w:fullDate` (B01 to B03, format `yyyy-MM-dd HH:mm`) |
| `date-forms-r1-word365.docx` | Word 365's save after Ctrl+A, F9 (Jason, 2026-10-28, version 2608 build 20326.20232, the machine in Sydney's zone) |
| `README.txt` | docx4j's checklist for the run, with its predictions made before it |
| `results.txt` | Jason's record of that run; `DateFormsProbeWordTest` parses its `Dnn_... = ` lines |

What Word showed (docx4j CR-067 section 12.2): a core property's date alone or time with no offset is UTC (created
`2024-02-03` showed 11:00:00 at UTC+11, modified `2025-06-07T08:09:10` 18:09:00 at UTC+10), to the minute; a
revision's `w:date` is the local clock in every form (a date alone local midnight, the `Z` ignored); a date control's
`w:fullDate` the nominal day in every form. Word's save wrote every `w:date` and `w:fullDate` as the nominal clock to
the minute labelled `Z`, `lastPrinted` to the minute, and `created` and `modified` as the save's moment.

`SHA1SUMS` lists the four copied files.
