"""docx4j's ``TextsAdapter``, over this package's text readers as they are before the text model.

CR-011 section 4. Not a test: the :class:`~conformance.Engine` the conformance
harness runs, :meth:`TextsAdapter.implemented` and
:meth:`TextsAdapter.extract`, with no ``inspect`` or ``counts`` (the inspector
is out of scope, CR-011 section 3.7). Phase A reads what the package ships
today, the baseline CR-011 section 1 measured:

* plain text (``txt``): ``pkg.body.get_text(view=)`` in ``accepted`` and
  ``original``. In ``markup``, the filter that ``get_text(view="markup")``
  applied until CR-011's Q10 made that call raise: each paragraph's runs
  outside every revision holder, joined as ``Body.get_text`` joins them. So
  the report records what the call returned, until Phase B gives it the
  model's markup view;
* Markdown (``md``): ``pkg.to_markdown(view=)``, which has ``accepted`` and
  ``markup`` (CriticMarkup) and refuses ``original``.

Every scenario's options are ignored, since today's readers take none, and
the findings are empty: nothing here raises a warning or reads the metadata.

It advertises what docx4j's adapter advertises at ``198d015a3``, less the
inspector's statements, and ``html reader absent``: the target, not what the
package meets. So every check docx4j makes is made here, today's gaps are
``FAIL`` rows rather than pending ones, and a phase's progress shows as checks
going from ``FAIL`` to ``PASS`` in the committed report (CR-011 sections 3.6
and 11). Pending are docx4j's 19, which need an HTML reader, and the
inspector's 214.
"""

from __future__ import annotations

from collections.abc import Collection
from pathlib import Path
from typing import Any

from conformance import HTML_READER_ABSENT, Engine, Output

from docx4j_py import load
from docx4j_py.model.content.text_model import text_of_view

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


class TextsAdapter(Engine):
    """The harness's engine over today's ``get_text`` and ``to_markdown``.

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
        """One fixture in one view as plain text or Markdown; the scenario is not read."""
        if format not in ("txt", "md"):
            raise NotImplementedError(f"format {format}")
        package = self._packages.get(fixture)
        if package is None:
            package = self._packages[fixture] = load(fixture)
        if format == "md":
            return Output(package.to_markdown(view=view), {})
        if view == "markup":
            paragraphs = package.body.iter_paragraphs()
            return Output("\n".join(text_of_view(p.element, view="markup") for p in paragraphs), {})
        return Output(package.body.get_text(view=view), {})
