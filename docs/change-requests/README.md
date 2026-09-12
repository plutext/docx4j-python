# Change requests

Modelled on docx4j-core-ts (`~/git/docx4j-core-ts/docs/change-requests/`). CR-002 follows that repository's CR-001 (engine) and CR-003 will follow its CR-002 (content API);
both record the departures.

| CR | Title | Status |
|---|---|---|
| [CR-001](CR-001-object-model.md) | The object model: a forked xsdata (generator flags for optional-everything, defaults on read, `ChildList`; renamed runtime), per-namespace packages with docx4j's names, `Child` parent pointers and `deep_copy`, a generated `el` object factory with `p`/`r`/`t` sugar and `wml(...)` fragments, `text_of`/`walk`/`find`, `warm_up`, schema patches, prefix table, skipped-content reporting | **Phases A, B and C implemented 2026-09-12** (fork at `~/git/docx4j-xsdata`, branch `docx4j`, local only; sections 12, 13 and 14). Round trip over `samples/` is 50/50 parts canonically identical with zero differences and nothing skipped; `el` covers 1,964 element names over 51 namespaces; a shared `XmlContext` is verified thread safe, a shared `ParserConfig` is not. Phase D proposed |
| [CR-002](CR-002-engine.md) | The engine: `PartStore` and `PartSink`, `PartName`, content types, relationships, typed parts with lazy synchronous `contents`, load and save with byte-identical untouched parts and a per-part skipped-content report, lossless MCE with a resolving traversal view (departure from docx4j), `mc:Ignorable` on write, `PropertyResolver`, numbering `Emulator`, fonts | **Phase A implemented 2026-09-12** (section 12), no fork commit needed. 16 documents, 223 parts: every untouched part byte identical, all 141 typed WordprocessingML parts canonically identical with nothing skipped, engine import 24 ms on top of the model, 8 threads over whole packages with 0 differences; Word acceptance passed 2026-09-12. Phases B (resolution utilities) and C (PML, SML, flat OPC) proposed |
| CR-003 | A content API in the shape of Office JS over the tree; addresses and `outline()` for agents | To be written |
