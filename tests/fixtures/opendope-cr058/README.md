# docx4j CR-058's fixtures

CR-014 phases 1b and 1c's oracle (`tests/content/test_bound_dates.py`, `tests/content/test_store_item_repair.py`).

**Copied unedited** from `../docx4j/docx4j-core-tests/src/test/resources/OpenDoPE/` on branch `VERSION_17_3_6` (Apache-2.0;
docx4j's commits are the ones that last touched each file):

| File | docx4j commit | What |
|---|---|---|
| `typed-controls-dates.docx` | `24515bd85` | the portfolio's `bind-typed-formats` probe: thirteen date controls, DF1 to DF13, surfaces `DFn_STALE`, one per format and stored value; `BoundDateFormatTest.wordsFormatting` |
| `typed-controls-twelve.docx` | `9b9d053cb` | the `bind-typed-twelve` probe: six date controls, TW1 to TW6, the twelve-hour clock at 12 and stored offsets; `BoundDateFormatTest.twelveOClockAndOffsetsAsWordShowsThem` |
| `bind-plain.docx` | `24515bd85` | the `bind-plain` probe: B7's `w:storeItemID` `{...00FF}` names no part, its XPath fits the one part `{...0001}` |
| `bind-plain-two-parts.docx` | `24515bd85` | two parts fit: B7 has a node in both, B11 only in the second |
| `bind-plain-no-part.docx` | `24515bd85` | no part fits: B7's XPath in another namespace, B12's a root no part has |

Word 365 ran each on 22 and 23 October 2026 at UTC+11 (docx4j CR-058 section 8 has the tables; `StoreItemRepairTest` and
`BoundDateFormatTest` hold docx4j to them, and the tests here port theirs).

`SHA1SUMS` lists the five copied files.
