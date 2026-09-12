"""Embedded parts, for :func:`docx4j_py.runtime.warm_up`.

``warmup_document.xml`` and ``warmup_styles.xml`` are hand written, not taken
from any document: between them they exercise paragraph and run properties, a
style reference, a hyperlink, a field, a bookmark, a tracked insertion and
deletion, a content control, a table with a grid and borders, two drawings, an
``mc:AlternateContent``, OMML, ``w14`` attributes, a symbol, a break and a tab
--- which is what makes them representative. Both parse with nothing skipped;
``tests/test_warm_up.py`` keeps them that way.
"""
