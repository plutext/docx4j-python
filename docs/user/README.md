# User guides

How-to documents for `docx4j` users, each run against the corpus on the date it states, showing
the markup that results. The change requests under [`../change-requests/`](../change-requests/README.md)
are the design record; these are the "how do I" answers.

- [hyperlinks.md](hyperlinks.md) --- adding a link to a paragraph (`Range.hyperlink`, markdown,
  the typed tree) and reading one back, the markup, a defect at the end of a link in 0.1.0, and
  the proposal for what remains (`get_hyperlink_ranges`, an `insert_hyperlink` sugar, the
  facade's `add_hyperlink`).
- [properties.md](properties.md) --- adding or altering a property in `docProps/core.xml`,
  `app.xml` and `custom.xml`: the facade for core, the typed parts for all three, creating a
  missing part.
- [find-replace.md](find-replace.md) --- `search`, `find` and `replace_text` over the main
  document part, the headers and footers, the notes, or every body of the document; the options;
  tracked; a dry run; where a match is not found.
