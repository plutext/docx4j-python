"""The fixed colour tables a WordprocessingML theme needs. docx4j ``org.docx4j.model.color.ColorTables``, in part.

CR-009 Phase B. A theme's scheme slots are ``a:srgbClr`` or ``a:sysClr``: in
the 785 theme parts of ``samples/``, ``tests/fixtures/`` and
``out/cr012-b-corpus/`` (2026-10-24) every one of the 9,372 slots is one of the
two, with no transform, and every ``a:sysClr`` is ``windowText`` or ``window``.
So only :data:`SYSTEM` is here, docx4j's table of Windows' default system
colours: Office draws an ``a:sysClr`` as the machine's colour for its name, not
its ``lastClr`` (docx4j CR-034 section 11).

docx4j's ``PRESET`` (``a:prstClr``) and ``INDEXED`` (SpreadsheetML's legacy
palette) are CR-009 Phase C's, with ``DrawingMLColors`` and
``SpreadsheetColors``. The HTML colour names Word reads in a WordprocessingML
colour are not a table here either: they are
:data:`docx4j_py.model.values.NETSCAPE_COLOURS`, the one reader's (CR-018),
as they are docx4j's ``WordColorValues.HTML`` and no longer ``ColorTables``'.

The values are docx4j's at ``a9e0303ff``, which are Apache POI's
``org.apache.poi.sl.usermodel.PresetColor`` (trunk ``ae62bb5116``, Apache-2.0),
keyed by the names docx4j's schema gives ``ST_SystemColorVal``.
"""

from __future__ import annotations

from types import MappingProxyType

__all__ = ["SYSTEM"]

#: ``a:sysClr`` names to Windows' default ``RRGGBB``. docx4j ``ColorTables.SYSTEM``.
SYSTEM = MappingProxyType(
    {
        "scrollBar": "C8C8C8",
        "background": "000000",
        "activeCaption": "99B4D1",
        "inactiveCaption": "BFCDDB",
        "menu": "F0F0F0",
        "window": "FFFFFF",
        "windowFrame": "646464",
        "menuText": "000000",
        "windowText": "000000",
        "captionText": "000000",
        "activeBorder": "B4B4B4",
        "inactiveBorder": "F4F7FC",
        "appWorkspace": "ABABAB",
        "highlight": "3399FF",
        "highlightText": "FFFFFF",
        "btnFace": "F0F0F0",
        "btnShadow": "696969",
        "grayText": "6D6D6D",
        "btnText": "000000",
        "inactiveCaptionText": "000000",
        "btnHighlight": "E3E3E3",
        "3dDkShadow": "000000",
        "3dLight": "E3E3E3",
        "infoText": "000000",
        "infoBk": "FFFFE1",
        "hotLight": "0066CC",
        "gradientActiveCaption": "B9D1EA",
        "gradientInactiveCaption": "D7E4F2",
        "menuHighlight": "3399FF",
        "menuBar": "F0F0F0",
    }
)
