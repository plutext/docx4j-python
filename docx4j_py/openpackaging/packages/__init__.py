"""The package classes. docx4j ``org.docx4j.openpackaging.packages``.

:class:`OpcPackage` is the generic one and loads anything; the WordprocessingML
subclass types the main part. PresentationML and SpreadsheetML are Phase C
(CR-002 section 10) and load through :class:`OpcPackage` in the meantime.
"""

from __future__ import annotations

from docx4j_py.openpackaging.packages.opc_package import (
    OpcPackage,
    register_package_class,
)
from docx4j_py.openpackaging.packages.wordprocessingml_package import (
    PAGE_MARGINS,
    PAGE_SIZES,
    PageSizePaper,
    WordprocessingMLPackage,
)

__all__ = [
    "PAGE_MARGINS",
    "PAGE_SIZES",
    "OpcPackage",
    "PageSizePaper",
    "WordprocessingMLPackage",
    "register_package_class",
]
