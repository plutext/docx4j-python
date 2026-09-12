"""The two functions CR-002 section 7 puts on ``docx4j_py`` itself.

    >>> from docx4j_py import load, create_package  # doctest: +SKIP
    >>> pkg = load("in.docx")  # doctest: +SKIP
    >>> pkg.save("out.docx")  # doctest: +SKIP

They are thin: :func:`load` is ``OpcPackage.load`` and :func:`create_package` is
``WordprocessingMLPackage.create_package``. They exist so that the common case
is one import, and so that the name a reader of docx4j code looks for
(``WordprocessingMLPackage.load``) is still the one that does the work.
"""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.load import LoadOptions
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage
    from docx4j_py.openpackaging.packages.wordprocessingml_package import (
        PageSizePaper,
        WordprocessingMLPackage,
    )
    from docx4j_py.openpackaging.stores import PartStore

__all__ = ["create_package", "load"]


def load(
    source: str | os.PathLike[str] | bytes | IO[bytes] | PartStore,
    *,
    options: LoadOptions | None = None,
) -> OpcPackage:
    """Load a package from a path, a directory, bytes, a file object or a store.

    The class of the result follows the main part's content type, so a ``.docx``
    comes back as a
    :class:`~docx4j_py.openpackaging.packages.wordprocessingml_package.WordprocessingMLPackage`.
    """
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

    return OpcPackage.load(source, options=options)


def create_package(
    *,
    page_size: PageSizePaper = "A4",
    landscape: bool = False,
    margins: str = "NORMAL",
) -> WordprocessingMLPackage:
    """A new, empty ``.docx``. docx4j ``WordprocessingMLPackage.createPackage``."""
    from docx4j_py.openpackaging.packages.wordprocessingml_package import (
        WordprocessingMLPackage,
    )

    return WordprocessingMLPackage.create_package(
        page_size=page_size, landscape=landscape, margins=margins
    )
