# docprop-case: DOCPROPERTY names in any case (docx4j #244)

Copied from docx4j `VERSION_17_3_5` at `ea5461ee9`, `docx4j-core-tests/src/test/resources/fields/docprop-case/`
(Apache-2.0), byte for byte; `SHA1SUMS` holds the files' SHA-1s, and their git blob ids are docx4j's.

- `docprop-case.docx`, built by docx4j's `build244.py`: the custom property `Client` = "Acme Ltd" and the title
  "The Title". Each marker's field result starts as `OLD`: K1 `DOCPROPERTY Client`, K2 `client`, K3 `CLIENT`,
  K4 `"cLiEnT"`, K5 `DOCPROPERTY Title`, K6 `title`, K7 `DOCPROPERTY Nope` (no such property).
- `docprop-case-word365.docx`: Word 365's save after select all and F9 (the user's run, 2026-10-25). Word found every
  name in any case, and wrote its error text for K7.

`tests/text/test_docproperty_case.py` holds the text model's `fields="update"` to Word's results, but for K7, where the
field keeps its cached result, as docx4j's `FieldUpdater` keeps the old one.
