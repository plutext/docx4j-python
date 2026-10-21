# CR-006: One MCP server over documents, decks and workbooks, with one vocabulary of tools

**Status:** Proposed 2026-09-19. Nothing implemented. Phase A is ready to start; B and C wait
for CR-004 and CR-005.
**Depends on:** CR-003 (implemented: every phase, sections 10 to 20) for Phase A; CR-004 Phase D
for Phase B; CR-005 Phase E for Phase C. The code lives in a separate product repository
(`../docx4j-mcp-py`), as CR-003 decided question 11 on 2026-09-16; this CR is the contract
between that product and this library, kept here because every tool is a few lines over the
library and the library is where the behaviour is.
**Counterpart:** `../docx4j-mcp` (Java; `CR-mcp-server.md`: the coarse tools, the path policy,
the inline cap, the execution model, stdio logging discipline), whose tool names this server
keeps for the tools it shares; the MCP Python SDK (`mcp`, the server's one dependency beyond
this library); CR-003 section 3.4 (the agent surface) and its decided question 11.

## 1. Why

CR-003 was designed for agents behind an MCP server first: paraId-first addresses, `outline()`
under a budget, `find()` with snippets, `describe()`, a `ChangeReport` on every call, `dry_run`,
`DocumentSession`, markdown in and out with addresses. Every one of those is in the library so
that every server gets the same behaviour, and none of them is a server. This CR is the server:
the thin product that turns the library into tools an agent can call over stdio.

The one thing it adds that the Java server does not have is **one vocabulary across three
formats**. The Java server is Word only. In Python, once CR-004 and CR-005 land, the same
`outline`, `find`, `read`, `edit`, `describe` and `fill` tools take a `.docx`, a `.pptx` or an
`.xlsx` and answer in the shape of that format: an outline of slides, a find over cells. Nobody
in the Python ecosystem offers that; python-docx, python-pptx and openpyxl are three projects
with three idioms, and an agent that has learned one has learned nothing about the others.

## 2. What the Java server settled, kept here

From `docx4j-mcp`'s CR, sections 3, 4 and 6, adopted as rules:

- **Paths, not blobs.** Every file parameter, inputs and `output_path`, is a local path validated
  against an allow-list of root directories given at server start. A path outside the roots is
  refused before anything is opened.
- **`overwrite: true` or refuse.** A writing tool never overwrites an existing file unless told
  to; the refusal names the file.
- **Inline caps.** Every tool that returns text inline caps it (`--max-inline-chars`, default
  200k); past the cap it writes to `output_path` when given, else truncates and says so in the
  result metadata (`truncated`, `chars`), which is what CR-003's budgets already produce.
- **Nothing is fetched.** No remote image, stylesheet or schema, ever (CR-003 section 4).
- **stdout is the protocol's.** All logging to stderr or a file; a stray print corrupts the
  session. Python's `logging` is configured accordingly at startup and `warnings` routed to it.
- **Warm-up at start**, before the first tool call: `docx4j_py.warm_up()` (the model's parse
  path) so the first call does not pay the 0.7 s import plus the first-parse cost.
- **The coarse tools keep the Java names**: `describe_template`, `fill_template`,
  `markdown_to_docx`, `docx_to_markdown`, `extract_text`, with the same parameters where the
  engines have the same capability, so a client configured for one server works with the other.

And one departure, decided by CR-003 question 5: **sessions, not stateless calls.** The Java
server loads, operates and saves on every call. Here a `DocumentSession` keeps a package open
across calls, keyed by its path, with the idle-close rule and the thread rule the library
already implements (CR-003 section 12.1), because an agent's edit is a conversation of a dozen
small calls over one document and the parse is the expensive part. A tool call on a path that
is not open opens it; `save` writes; `close` discards; the session closes itself when idle. The
stateless shape is still available (`edit` with `output_path` on a closed document opens, edits,
saves and closes in one call), so a client that expects the Java model gets it.

## 3. The tool surface

All tools take `path` (a document, deck or workbook, dispatched on the package kind the engine
sniffs) unless noted. Results are JSON from the library's `to_dict()`s.

### 3.1 Reading

| Tool | Input | Output |
|---|---|---|
| `outline` | `path`, `max_chars`, `budget`, `depth`, the text options (below) | `Outline.to_dict()`: blocks with addresses for a document (the text model's text, levels and counts; a text box's paragraphs after their anchor at `textbox:` addresses, flagged `textbox`), slides and shapes for a deck, sheets and used ranges for a workbook; `stats`; `truncated`; `warnings` |
| `find` | `path`, `text`, `limit`, `match_case`, `match_whole_word`, `context` | `SearchHit`s with addresses and snippets, over Office JS's text (below); `warnings` |
| `read` | `path`, `address` (optional; the whole when omitted), the text options (`view`: `accepted`, `original` or `markup`, and the rest), `format` (`markdown`, the default, `plain` or `blocks`), `max_chars`, `notes` (decks), `max_rows` (workbooks) | for a document, `markdown`: the Markdown with addresses, `to_markdown(addresses=True)` under the default options and `markdown_blocks(extract(pkg, **options), pkg)` with the address comments under others; `plain`: `plain_text`; `blocks`: `TextResult.to_dict()`, REQ-706's block model; the block, slide or range named, or the whole under the cap; `warnings` |
| `describe` | `path` | `describe()`: styles, bindings and repeats for a document; layouts, placeholders and theme for a deck; sheets, tables, names and validations for a workbook |
| `extract_text` | `path`, the text options | plain text, `plain_text(extract(pkg, **options))`: one line per paragraph, slide or row (the Java tool, kept, now docx4j's text model's as the Java tool's is since docx4j CR-046); `warnings` |
| `docx_to_markdown` | `path`, `output_path`, `tracked_changes` | the Java tool, kept; `read` without an address is the same thing |
| `get_comments`, `get_tracked_changes` | `path` | the threads / the changes, with addresses |
| `reveal` | `path`, `address` | `to_api_script` for the block, slide or range: the calls that would make it |

**The text the reading tools give** is the library's one text model,
[CR-011](CR-011-one-text-model.md) (`docx4j_py.model.text`, docx4j's CR-046, the portfolio's
specification at WD 0.28), the same as docx4j-mcp's since docx4j CR-046 phase 3:

- **The text options** are WD section 10's, by their names verbatim and taking its values
  (`view`, `include`, `text_boxes`, `notes`, `note_references`, `comments`, `hidden`,
  `field_codes`, `fields`, `toc`, `placeholders`, `bindings`, `list_labels`, `charts`, `smartart`,
  `embedded`, `altchunks`, `alt_text`, `metadata`, `symbols`, `math`, `tables`, `markup_detail`,
  `page_markers`, `layout`, `warnings`), passed through to `extract(pkg, **options)`
  (`TextOptions.of`); an unknown name or value is the library's `TextOptionError`, naming its
  statement (REQ-1001), returned as an error with that statement. `outline`, `read` and
  `extract_text` take them; `outline`'s entries are the text model's under the defaults and the
  options decide its `warnings`.
- **The blocks** are REQ-706's: `TextResult.to_dict()` (`blocks`, `comments`, `warnings`,
  `metadata`, `options`), the JSON docx4j-mcp's `blocks` output has.
- **The warnings** are REQ-801's, in every reading tool's result: each `TextWarning.to_dict()`
  (`kind`, `address`, `selector`: `mce.branches_differ`, `field.no_result`,
  `sdt.binding_differs`, `altchunk.not_expanded`, `content.not_read`, `image.not_described`), so
  an agent knows where what it read is not the whole document. `find` gives the default
  reading's.
- **`find` stays on Office JS's text** (CR-011 Q4): its matches, offsets and addresses are
  `Paragraph.text`'s, so a hit converts to a `Range` and an edit; a text box's paragraphs are not
  searched, and `outline` and `read` list them at their `textbox:` addresses, which are read only
  (Q9; the library's `address.read_only`).

### 3.2 Editing, always through the library's verbs

| Tool | Input | Output |
|---|---|---|
| `edit` | `path`, `address`, `markdown` (or `text`), `location` (`Replace` / `Before` / `After` / `Start` / `End`), `dry_run` | `ChangeReport.to_dict()`: parts touched, addresses created and moved, warnings; with `dry_run: true` the report of what would happen and nothing written |
| `insert` | `path`, `address`, `markdown`, `location` | `edit` with `After` as default; a slide from `# Title` + bullets in a deck; a GFM table into cells in a workbook |
| `delete` | `path`, `address` | the report |
| `set_style`, `set_format` | `path`, `address`, the setter's arguments in the content API's names | the report |
| `replace_text` | `path`, `find`, `replace`, the search options | the count and the report |
| `comment` | `path`, `address` (or `text` to anchor on), `text`, `reply_to` | the `Comment` |
| `track_changes` | `path`, `mode` (`Off` / `TrackAll`) | the report; every later `edit` on that session is tracked |
| `accept_all` / `reject_all` | `path` | the count |
| `describe_template` / `fill_template` | the Java tools, kept: `path`, `data` (JSON or XML), `output_path` | the description / the report |
| `markdown_to_docx` | `markdown`, `output_path`, `styles_template_path` | the Java tool, kept; a `.pptx` or `.xlsx` output path makes a deck (one slide per `#`) or a workbook (one sheet per GFM table) |
| `insert_ooxml` | `path`, `address`, `ooxml` | for the rare fragment markdown cannot carry |
| `save` | `path`, `output_path` (optional), `overwrite` | writes; the session stays open on the new path |
| `close` | `path`, `discard` | closes the session |

Every editing tool returns the `ChangeReport` the library records, so the audit trail CR-003
built for Word review is what the agent sees in the tool result.

### 3.3 What is deliberately not here

`convert_to_pdf`, `convert_to_html` and `html_to_docx`: there is no renderer and no XHTML
importer in this engine, and the Java server has them. A client that needs both runs both
servers; the shared tool names and path policy make that unremarkable. `anonymize`, `compare`,
`merge_documents`, `paginate`: the Java server's, with no Python counterpart planned.

## 4. Execution model

- **Transport:** stdio first (one process per client session), streamable HTTP later behind the
  same tool functions; the SDK's server object owns both.
- **Sessions:** `DocumentSession` per path, one thread per package as the library requires, an
  idle close (`--idle-seconds`, default 300), an explicit `close`. A crash loses unsaved edits,
  which is the trade CR-003 question 5 accepted; `save` is cheap and the agent is told so in
  each editing tool's description.
- **Determinism:** `--seed` sets `package.id_seed` for created documents so a test can compare
  bytes; the wall clock stays in the docProps dates and revision dates, as the library leaves it.
- **Security:** the path allow-list; no network; the engine's zip limits (entry count, size ratio)
  from `LoadOptions`; no code execution from a document (a `.docm`'s VBA is bytes).
- **Errors:** the library's `ContentError` hierarchy maps to a tool error with its `code` and
  `hint` (`"call outline() to list current addresses"` becomes `"call the outline tool"`), so an
  agent can recover without a human.
- **Budgets:** every tool that can be large takes the library's budget parameters and reports
  `truncated`; the server-wide inline cap is the last resort.

## 5. Where each piece lives

| Piece | Where |
|---|---|
| the server, its tool functions, the path policy, the session table, the CLI (`docx4j-mcp-py --root DIR ...`) | `../docx4j-mcp-py` (product repository; `pyproject.toml` depending on `docx4j-py` and `mcp`) |
| every behaviour a tool relies on | this library: `docx4j_py.model.content`, `.model.presentation` (CR-004), `.model.workbook` (CR-005), `.model.markdown`, `.model.sessions` |
| the tool-shaped scenario tests that already exist | `tests/agent/` here; the server's own tests call the tools over stdio with the SDK's client and assert the same things |
| this CR | here, as the contract; the product's README points at it |

## 6. Tests

- **In this library:** nothing new; `tests/agent/` is the tool shapes already (`outline`, `find`,
  edit by address, `ChangeReport`, `dry_run` then commit, sessions).
- **In the product:** every tool called over stdio through the SDK's client against the
  fixtures copied from here; the path policy (a path outside the roots refused, `overwrite`
  respected); the inline cap; the session lifecycle (open on first call, idle close, `close`
  with `discard`); a scripted agent session per format (read the outline, find, edit, dry run,
  comment, track, save, reopen, check) that asserts the saved file with the library; the
  Java-compatible tools called with the Java server's parameter shapes.
- **Manual:** the same acceptance checks as the library's (`tests/README.md` here) on the files
  the scripted sessions save; and one real client session (Claude Code with the server
  configured) recorded in the product's README.

## 7. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | The product repository; the server over CR-003: every tool of section 3 for `.docx`, sessions, path policy, caps, warm-up, logging; the Java-named coarse tools; tests over stdio; README with a Claude Code configuration | 5 days |
| B | Decks: the same tools dispatching on a `.pptx` once CR-004 Phase D lands; `markdown_to_docx` to a deck | 2 days |
| C | Workbooks: the same tools dispatching on an `.xlsx` once CR-005 Phase E lands; `markdown_to_docx` to a workbook; `read` with `max_rows` | 2 days |
| D | Streamable HTTP transport; resources (an opened package's inventory) and a `fill-template` prompt, as the Java CR's phase 4 | 2 days, later |

Phase A is ready now and is the MCP server's first release, as CR-003 section 8 said of its
phases B, D and K.

The tracked-change tools of section 3 --- `get_tracked_changes`, `accept_all` and `reject_all`
--- are backed by **`docx4j_py.model.revisions`** ([CR-010](CR-010-tracked-changes-as-word-resolves-them.md),
implemented 2026-10-18): the listing groups as Word's Review pane does and covers every story,
and accepting or rejecting resolves as Word's Accept All and Reject All do, held to Word's own
saves. Over a document, the tools call the package's verbs --- `pkg.get_tracked_changes()`,
`revisions.accept_all(pkg)` and `revisions.reject_all(pkg)`, every story --- rather than the
body's, which resolve the main story only (CR-010 section 3.3); each change's `to_dict()` is the
result row, with its `part`, and the count is what the verbs return. `TrackedChangeFilter` and
`revisions.report(pkg)` are there for a tool that takes an author or wants a summary, and the
`outline` tool's `stats.tracked_changes` counts the same listing over the stories an outline
shows.

A create-chart tool is not in this contract yet. If one is added, docx4j's CR-036 is its natural
backing in Java: `ChartCreator` and `ChartSettings`, handed off 2026-10-12 and recorded in CR-005
section 8 and CR-004 section 8. That CR was finished the same day, and since its phase 3 a chart
can be created in all three formats: `inSheet` for an xlsx, `inNewParagraph` for a docx, `inSlide`
for a pptx. A delete-chart tool would be `ChartCreator.delete`, for any host.

Here, the backing would be:

- CR-005 Phase G's port for workbooks;
- CR-004 Phase F's port for decks;
- for documents, a phase this library does not have yet (CR-004 section 8).

## 8. Open questions (recommendations for decision)

1. **Repository placement.** Recommendation: `../docx4j-mcp-py`, as decided; this CR stays here as
   the contract and the portfolio registry gets a repo key for the product when it exists.
2. **Tool naming: format prefixes or not.** Recommendation: no prefixes for the generic tools
   (`outline`, not `docx_outline`), because the point is that one name works for three formats;
   the Java names kept verbatim for the shared coarse tools even where `docx_` is in the name.
3. **Sessions keyed by path or by handle.** Recommendation: by path, with the session table
   invisible to the agent; a handle is one more thing for a model to lose.
4. **The MCP SDK.** Recommendation: the official `mcp` package (FastMCP-style decorators), pinned;
   no framework on top.
5. **Markdown to a deck and a workbook.** Recommendation: yes, as the one place a whole-document
   markdown becomes slides or sheets (CR-004 question 5, CR-005 section 3.4), because it is a
   loop in the server and not a library feature.
6. **PDF.** Recommendation: not in this server; document how to run the Java server beside it.
