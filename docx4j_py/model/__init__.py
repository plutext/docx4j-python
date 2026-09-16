"""The document model: views over the tree, in Office JS's vocabulary.

CR-003. :mod:`docx4j_py.model.content` holds ``Body``, ``Paragraph``, ``Range``
and ``Font`` (Phase B); the agent surface (addresses, ``Outline``,
``ChangeReport``), markdown, custom XML and the sessions registry join it in
later phases.

Nothing here is a proxy: a view holds an element and the container that holds
it, `paragraph.element` is the ``P``, and ``body.content`` is the live
``ChildList``. The tree stays reachable (CR-003 section 3.13).
"""

from __future__ import annotations

__all__: list[str] = []
