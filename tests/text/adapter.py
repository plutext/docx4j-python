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
  the metadata by ``metadata<TAB>name<TAB>`` (``fields_update_on_open``).

It advertises what docx4j's adapter advertises at ``198d015a3``, less the
inspector's statements, and ``html reader absent``: the target, not what the
package meets. So every check docx4j makes is made here, a gap is a ``FAIL``
row rather than a pending one, and a phase's progress shows as checks going
from ``FAIL`` to ``PASS`` in the committed report (CR-011 sections 3.6 and
11). Pending are docx4j's 19, which need an HTML reader, and the inspector's
214.
"""

from __future__ import annotations

from collections.abc import Collection
from pathlib import Path
from typing import Any

from conformance import HTML_READER_ABSENT, Engine, Output

from docx4j_py import load
from docx4j_py.model.text import TextOptions, extract, plain_text

__all__ = ["DOCX4J_IMPLEMENTED", "INSPECTOR", "TextsAdapter"]

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
        for name, value in result.metadata.items():
            findings[f"metadata\t{name}\t"] = str(value).lower() if isinstance(value, bool) else str(value)
        if format == "md":
            return Output(package.to_markdown(view=view), findings)
        return Output(plain_text(result), findings)
