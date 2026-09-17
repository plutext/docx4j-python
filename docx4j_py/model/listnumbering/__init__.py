"""List numbering: the label Word paints in front of a paragraph.

docx4j ``org.docx4j.model.listnumbering``, and CR-002 section 6.2's promise,
delivered early by CR-003 Phase H because ``Word.ListItem.listString`` and the
markdown exporter's list markers both need it (CR-002 section 12.11, CR-003
section 18).

    >>> from docx4j_py import load
    >>> from docx4j_py.model.listnumbering import labels_for
    >>> pkg = load("lists.docx")                        # doctest: +SKIP
    >>> labels_for(pkg.body)[id(pkg.body.paragraphs[2].element)].result.num_string
    '1.'

What is here is the **counting core**: the definitions (:class:`ListLevel`,
:class:`AbstractListNumberingDefinition`, :class:`ListNumberingDefinition`), the
counters (:class:`NumberingState`, one per story --- :class:`NumberingStates`),
the eight number formats of CR-002 section 6.2 (:mod:`.formats`) and the
:class:`Emulator` that puts them together. The twenty exotic formats docx4j
ships, and resolution through ``PropertyResolver``, are CR-002 Phase B's.

Nothing here unmarshals a part: the definitions and the paragraph styles are
read with lxml from the bytes the parts would be saved as.
"""

from __future__ import annotations

from docx4j_py.model.listnumbering.definitions import (
    LEVELS,
    AbstractListNumberingDefinition,
    Counter,
    Indent,
    ListLevel,
    ListNumberingDefinition,
    NumberingState,
    NumberingStates,
    StyleNumbering,
    read_definitions,
)
from docx4j_py.model.listnumbering.emulator import (
    Emulator,
    ItemLabel,
    NumberingResult,
    NumRef,
    ResultTriple,
    emulator_of,
    invalidate,
    invalidate_labels,
    invalidate_styles,
    labels_for,
    states_for,
)
from docx4j_py.model.listnumbering.formats import FORMATS, format_value, register

__all__ = [
    "FORMATS",
    "LEVELS",
    "AbstractListNumberingDefinition",
    "Counter",
    "Emulator",
    "Indent",
    "ItemLabel",
    "ListLevel",
    "ListNumberingDefinition",
    "NumRef",
    "NumberingResult",
    "NumberingState",
    "NumberingStates",
    "ResultTriple",
    "StyleNumbering",
    "emulator_of",
    "format_value",
    "invalidate",
    "invalidate_labels",
    "invalidate_styles",
    "labels_for",
    "read_definitions",
    "register",
    "states_for",
]
