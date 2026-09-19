"""docx4j's font tables, and the three Office themes it supplies to a themeless package.

CR-002 section 15.3. ``substitutions.json``, ``families.json`` and
``registry.json`` are written by ``scripts/generate_fonts.py`` from docx4j's own
resources and carry a ``_source`` object naming the commit; ``theme-2007.xml``,
``theme-2013.xml`` and ``theme-2023.xml`` are docx4j's
``openpackaging/parts/WordprocessingML/theme-YYYY.xml``, copied byte for byte.
None of it is hand-edited: regenerate instead.
"""
