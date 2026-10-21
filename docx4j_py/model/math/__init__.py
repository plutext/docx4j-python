"""Office Math as LaTeX: docx4j's ``org.docx4j.model.math``. CR-011 section 6.

:class:`OmmlToLatex` converts an ``m:oMath`` or ``m:oMathPara`` to LaTeX for
docx4j's subset, and :class:`OmmlMathError` refuses anything outside it; the
text model writes an equation so (WD REQ-604), its runs' text where the
converter refuses it. :mod:`~docx4j_py.model.math.latex_symbols` holds the
tables docx4j's converters share.

Imported on first use: ``import docx4j_py`` does not import it.
"""

from __future__ import annotations

from docx4j_py.model.math.omml_to_latex import OmmlMathError, OmmlToLatex

__all__ = ["OmmlMathError", "OmmlToLatex"]
