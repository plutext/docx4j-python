# docx4j CR-067's memo, one file, many readings

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/fields/cr067-fields-fun/` at docx4j's release
tag `docx4j-17.3.5` (`64fae4cd0`; Apache-2.0; the git blob id is docx4j's), by `git show` (CR-027 Phase D, 2026-11-05).

One file, `one-file-many-readings.docx`: a one-page purchase approval memo whose ten fields tell each kind of reader a
different story. docx4j's `TextOptionsProfileTest.theMemoUnderThePreset` reads it under `profile: word-text`, and so
does `tests/text/test_text_options_profile.py`. CR-027 section 5 names this directory as CR-026's, or this CR's own
copy if CR-026 has not landed; it has not, so this is the memo alone. CR-026 section 4.5 copies the rest of the
directory here (Word's five saves, the README and the three tables).

`SHA1SUMS` lists the file; this README is not in it.
