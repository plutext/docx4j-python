"""``document.settings``: the one document-level setting python-docx offers.

CR-003 section 3.6, Phase J. python-docx's ``Settings`` has exactly one public
member, ``odd_and_even_pages_header_footer`` over ``w:evenAndOddHeaders``, and
it is cheap, so it is here rather than marked as an omission. Everything else
in ``/word/settings.xml`` is on the typed part
(``document.package.document_settings_part.contents``), and the one setting an
agent asks about --- the compatibility mode --- is ``pkg.compatibility_mode``
(CR-003 section 17.11).
"""

from __future__ import annotations

from typing import Any

from docx4j_py.wml import el

__all__ = ["Settings"]


class Settings:
    """``document.settings``: over ``/word/settings.xml``."""

    __slots__ = ("_package",)

    def __init__(self, package: Any) -> None:
        """Wrap the package whose settings part this is over."""
        self._package = package

    def __repr__(self) -> str:
        """``<docx4j_py.docx Settings>``."""
        return "<docx4j_py.docx Settings>"

    @property
    def _part(self) -> Any:
        return getattr(self._package, "document_settings_part", None)

    @property
    def element(self) -> Any:
        """The typed ``w:settings``; raises when the document has no settings part."""
        part = self._part
        if part is None:
            raise ValueError("this document has no /word/settings.xml")
        return part.contents

    @property
    def odd_and_even_pages_header_footer(self) -> bool:
        """``w:evenAndOddHeaders``: distinct odd- and even-page headers and footers."""
        part = self._part
        if part is None:
            return False
        toggle = getattr(part.contents, "even_and_odd_headers", None)
        if toggle is None:
            return False
        value = getattr(toggle, "val", None)
        return True if value is None else bool(value)

    @odd_and_even_pages_header_footer.setter
    def odd_and_even_pages_header_footer(self, value: bool) -> None:
        from docx4j_py.model.content.reports import recording_on

        part = self._part
        if part is None:
            raise ValueError("this document has no /word/settings.xml to write to")
        with recording_on(self._package, "docx.settings.odd_and_even", part):
            settings = part.contents
            if not value:
                settings.even_and_odd_headers = None
                return
            toggle = el.evenAndOddHeaders()
            settings.even_and_odd_headers = toggle
            toggle.parent = settings
