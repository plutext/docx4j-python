"""Markdown in and out, over the content API. CR-003 section 3.5, Phase K.

Markdown is what models read and write best, and docx4j-mcp made it the
exchange format for that reason. This is the Python engine's own, small and
honest, over the views of Phases B and D rather than beside them:

    >>> from docx4j_py import load
    >>> pkg = load("in.docx")                         # doctest: +SKIP
    >>> print(pkg.to_markdown(addresses=True))        # doctest: +SKIP
    <!-- w14:5A2B1C3D -->
    # Quarterly Report
    ...
    >>> pkg.paragraph_at("w14:5A2B1C3D").insert_paragraph("New")   # doctest: +SKIP

**Two workflows, and the README says which to use when.** The coarse one is
``pkg.to_markdown()`` out, an agent edits the markdown, ``create_package()``
plus ``insert_markdown`` back in: cheap, and it loses whatever the markdown
could not carry. The fine one is ``to_markdown(addresses=True)`` to read and
then in-place edits by address: nothing outside the blocks it touches changes,
and the untouched parts are still written back byte for byte.

Everything public is here; the two halves live in :mod:`.export` (which needs
no dependency at all) and :mod:`.importer` (which imports ``markdown-it-py``
inside the call). The verbs are reachable from the views themselves ---
``body.to_markdown()``, ``paragraph.to_markdown()``, ``body.insert_markdown()``,
``pkg.to_markdown()`` --- which is what a caller uses; these functions are what
those methods are.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.markdown.export import (
    ADDRESS_COMMENT,
    CODE_STYLE_IDS,
    MONOSPACE_FONTS,
    MarkdownView,
    MarkdownViewValue,
    address_comment,
    body_markdown,
    markdown_budget_of,
    paragraph_markdown,
    table_markdown,
)
from docx4j_py.model.markdown.importer import (
    CUSTOM_STYLE_XML,
    STYLE_IDS,
    blocks_for,
    ensure_style,
    style_ids_of,
)

__all__ = [
    "ADDRESS_COMMENT",
    "CODE_STYLE_IDS",
    "CUSTOM_STYLE_XML",
    "MONOSPACE_FONTS",
    "STYLE_IDS",
    "MarkdownView",
    "MarkdownViewValue",
    "address_comment",
    "body_markdown",
    "ensure_style",
    "insert_markdown_into",
    "markdown_budget_of",
    "paragraph_markdown",
    "style_ids_of",
    "table_markdown",
]


def insert_markdown_into(
    body: Any,
    markdown: str,
    *,
    location: str = "End",
    target: Any = None,
) -> list[Any]:
    """Insert markdown as blocks and return the views of what went in.

    What ``Body.insert_markdown`` and ``Paragraph.insert_markdown`` are. One
    :func:`~docx4j_py.model.content.reports.recording` is opened for the whole
    fragment, so twenty blocks are one ``ChangeReport`` listing twenty
    addresses (CR-003 section 12.7), and the nested verbs are no-ops inside it.

    Args:
        body: the :class:`~docx4j_py.model.content.body.Body` to insert into.
        markdown: CommonMark plus GFM tables and strikethrough.
        location: ``"Start"``, ``"End"``, ``"Before"`` or ``"After"``; the last
            two need `target`.
        target: the paragraph or block to insert relative to.

    Returns:
        The views of the inserted blocks, in document order.
    """
    from docx4j_py.model.content.reports import recording

    with recording(body, "insert_markdown") as change:
        elements, touched = blocks_for(body, markdown, change)
        if not elements:
            return []
        body.insert_element(elements, location=location, target=target)
        parts = getattr(change, "parts", None)
        if parts is not None:
            for name in sorted(touched):
                if name not in parts:
                    parts.append(name)
        return [body.view_for(element) for element in elements]
