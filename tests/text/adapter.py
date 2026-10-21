"""docx4j's ``TextsAdapter``, over this package's text model. CR-011 sections 3.6 and 12.

Not a test: the :class:`~conformance.Engine` the conformance harness runs,
:meth:`TextsAdapter.implemented` and :meth:`TextsAdapter.extract`, with no
``inspect`` or ``counts`` (the inspector is out of scope, CR-011 section 3.7).
Since Phase B it reads the text model, as docx4j's adapter reads its own:

* plain text (``txt``): :func:`docx4j_py.model.text.plain_text` of
  :func:`~docx4j_py.model.text.extract`, the scenario's options set by their
  specification names (docx4j's ``TextsAdapter.options``: ``; `` apart, each
  ``name: value``; ``default``, ``markdown`` and the HTML reader scenarios set
  none);
* Markdown (``md``): ``pkg.to_markdown(view=)`` as before, which has
  ``accepted`` and ``markup`` (CriticMarkup) and refuses ``original``, until
  Phase D writes Markdown from the blocks;
* the findings: each warning by ``warning<TAB>kind<TAB>selector``, counted, and
  the metadata by ``metadata<TAB>name<TAB>`` (``fields_update_on_open``, and
  each core, application and custom property by its own name, as docx4j's
  adapter flattens them).

It advertises what docx4j's adapter advertises at ``198d015a3``, less the
inspector's statements, and ``html reader absent``: the target, not what the
package meets. So every check docx4j makes is made here, a gap is a ``FAIL``
row rather than a pending one, and a phase's progress shows as checks going
from ``FAIL`` to ``PASS`` in the committed report (CR-011 sections 3.6 and
11). Pending are docx4j's 19, which need an HTML reader, the inspector's 214,
and the 60 checks of the content CR-011 Q7 defers (embedded SpreadsheetML and
PresentationML packages and OLE 1.0 objects), each named in :data:`DEFERRED`.
"""

from __future__ import annotations

from collections.abc import Collection
from pathlib import Path
from typing import Any

from conformance import HTML_READER_ABSENT, Engine, Output

from docx4j_py import load
from docx4j_py.model.text import TextOptions, extract, plain_text

__all__ = ["DEFERRED", "DOCX4J_IMPLEMENTED", "INSPECTOR", "TextsAdapter"]

#: docx4j's ``TextsAdapter.IMPLEMENTED`` at ``198d015a3``: the statements its
#: CR-046 phases 1 to 3 implement, the chart description case 34 needs
#: (``CR-047``), and the inspector's (CR-050) with the suite's check of it.
DOCX4J_IMPLEMENTED = frozenset(
    {
        *(
            f"REQ-{n}"
            for n in (
                201, 202, 203, 204,
                301, 302, 303, 304, 305, 306, 307,
                401, 402, 403, 404, 405, 406,
                501, 502, 503, 504, 505, 506, 507, 508, 509, 510,
                511, 512, 513, 514, 515, 516, 517, 518, 520,
                601, 602, 603, 604, 605, 606,
                701, 702, 703, 704, 705, 706, 707, 708, 709,
                801, 802, 1001, 1202, 1203,
                901, 902, 903, 904, 905, 906, 907, 1204,
            )
        ),
        "CR-047",
    }
)  # fmt: skip

#: The inspector's statements (WD section 9) and REQ-1204, its check: out of
#: scope here (CR-011 section 3.7), so never advertised.
INSPECTOR = frozenset({*(f"REQ-{n}" for n in range(901, 908)), "REQ-1204"})

_WORKBOOK = "deferred (CR-011 Q7): an embedded SpreadsheetML package"
_PRESENTATION = "deferred (CR-011 Q7): an embedded PresentationML package"
_OLE10 = "deferred (CR-011 Q7): an OLE 1.0 package"
_EMBEDDED = "embedded: true"
_EMBEDDED_HIDDEN = "embedded: true; hidden: true"

#: The checks of the content CR-011 Q7 defers, by (fixture, case, item,
#: scenario): each a token the embedded package holds, under the scenario that
#: reads it (a hidden sheet or slide under ``hidden: true`` only, since
#: without it the token's absence is checked and holds), and case 53's order.
#: REQ-512 covers an embedded WordprocessingML package too, which is read, so
#: the statement stays advertised and these are named instead (60 checks).
DEFERRED: dict[tuple[str, str, str, str], str] = {
    ("extraction-probe.docx", "39", "EMB39_XLSX_CELL", _EMBEDDED): _WORKBOOK,
    ("extraction-probe.docx", "41", "OLE41_NATIVE_TEXT", _EMBEDDED): _OLE10,
    **{
        ("extraction-probe.docx", "53", token, _EMBEDDED): _PRESENTATION
        for token in (
            "EMB53_SLIDE_ONE", "EMB53_GROUP_FIRST", "EMB53_GROUP_NESTED", "EMB53_AFTER_GROUP", "EMB53_TABLE_A",
            "EMB53_TABLE_B", "EMB53_TABLE_C", "EMB53_TABLE_D", "EMB53_SLIDE_TWO", "order",
        )
    },
    ("extraction-probe.docx", "53", "EMB53_HIDDEN_SLIDE", _EMBEDDED_HIDDEN): _PRESENTATION,
    ("extraction-probe-structure.docx", "57", "SHT57_SHOWN", _EMBEDDED): _WORKBOOK,
    ("extraction-probe-structure.docx", "57", "SHT57_LAST", _EMBEDDED): _WORKBOOK,
    ("extraction-probe-structure.docx", "57", "SHT57_HIDDEN", _EMBEDDED_HIDDEN): _WORKBOOK,
    ("extraction-probe-structure.docx", "57", "SHT57_VERY_HIDDEN", _EMBEDDED_HIDDEN): _WORKBOOK,
    **{
        ("extraction-probe-inspect.docx", "66", token, _EMBEDDED): _WORKBOOK
        for token in ("ROW66_SHOWN", "COL66_HIDDEN_CELL", "ROW66_HIDDEN")
    },
}  # fmt: skip


def options(view: str, scenario: str) -> TextOptions:
    """A scenario as options, by the specification's names. docx4j ``TextsAdapter.options``."""
    out = TextOptions(view=view)
    if scenario in ("default", "markdown") or scenario.startswith("html reader"):
        return out
    for option in scenario.split(";"):
        name, colon, value = option.partition(":")
        if not colon:
            raise ValueError(f"Unknown scenario: {scenario}")
        out.set(name.strip(), value.strip())
    return out


class TextsAdapter(Engine):
    """The harness's engine over the text model, and ``to_markdown`` for Markdown.

    Each fixture is loaded once and read as often as the harness asks, as
    docx4j's adapter caches its packages.
    """

    def __init__(self) -> None:
        """An adapter with no package loaded yet."""
        self._packages: dict[Path, Any] = {}

    def implemented(self) -> Collection[str]:
        """docx4j's statements less the inspector's, and the class path's ``html reader absent``."""
        return (DOCX4J_IMPLEMENTED - INSPECTOR) | {HTML_READER_ABSENT}

    def deferred(self, fixture: str, case: str, item: str, scenario: str) -> str | None:
        """The checks of CR-011 Q7's deferred content (:data:`DEFERRED`), pending by name."""
        return DEFERRED.get((fixture, case, item, scenario))

    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """One fixture in one view under one scenario, as plain text or Markdown."""
        if format not in ("txt", "md"):
            raise NotImplementedError(f"format {format}")
        package = self._packages.get(fixture)
        if package is None:
            package = self._packages[fixture] = load(fixture)
        result = extract(package, options(view, scenario))
        findings: dict[str, str] = {}
        for warning in result.warnings:
            key = f"warning\t{warning.kind}\t{warning.selector or ''}"
            findings[key] = str(int(findings.get(key, "0")) + 1)
        # the properties (core, app, custom) by their own names, as the findings table names them
        for name, value in result.metadata.items():
            for key, item in value.items() if isinstance(value, dict) else ((name, value),):
                findings[f"metadata\t{key}\t"] = str(item).lower() if isinstance(item, bool) else str(item)
        if format == "md":
            return Output(package.to_markdown(view=view), findings)
        return Output(plain_text(result), findings)
