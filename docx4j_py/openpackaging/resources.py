"""docx4j's own default parts, as package resources.

``styles.xml``, ``numbering.xml`` and ``fontTable.xml`` from
``org/docx4j/openpackaging/parts/WordprocessingML/`` in docx4j-core, copied under
``docx4j_py/resources/openpackaging/``. ``create_package`` uses the styles, as
docx4j's does; the numbering and the font table are loaded on demand by
``NumberingDefinitionsPart.unmarshal_default_numbering`` and
``FontTablePart.unmarshal_default_fonts``, which is also when docx4j loads them.

**One change on the way across**: docx4j ships ``styles.xml`` and
``numbering.xml`` as UTF-16 with a byte-order mark and an ``encoding="utf-16"``
declaration; both are transcoded to UTF-8 here and their declarations changed to
match. Nothing else is touched, and nothing downstream can tell: they are
unmarshalled and re-marshalled before they reach a document.
"""

from __future__ import annotations

from importlib import resources

__all__ = ["DEFAULT_PARTS", "default_part_bytes"]

#: The resources this module offers.
DEFAULT_PARTS: tuple[str, ...] = ("styles.xml", "numbering.xml", "fontTable.xml")


def default_part_bytes(name: str) -> bytes:
    """The bytes of one of docx4j's default parts.

    Args:
        name: one of :data:`DEFAULT_PARTS`.
    """
    if name not in DEFAULT_PARTS:
        raise KeyError(f"{name!r} is not one of {DEFAULT_PARTS}")
    return (resources.files("docx4j_py.resources.openpackaging") / name).read_bytes()
