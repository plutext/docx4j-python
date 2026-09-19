"""``PhysicalFont``: a font a consumer has. docx4j ``org.docx4j.fonts.PhysicalFont``.

CR-002 section 15.3. docx4j's class wraps FOP's ``EmbedFontInfo`` and carries the
font's panose and its file URI, so that FOP can be told where to read the
outlines. This phase ports the **decision** half of ``org.docx4j.fonts`` and not
the rendering half (CR-002 section 6.3), so a physical font here is a name, the
family a CSS ``font-family`` would use, and the no-bold-face flag --- everything
a :class:`~docx4j_py.model.fonts.mapper.Mapper` answer means to a consumer that
draws with a browser, measures with a font library, or only reports. Reading
font files (docx4j's ``PhysicalFonts.discover``) is a later CR.
"""

from __future__ import annotations

import dataclasses

__all__ = ["NOBOLD_SUFFIX", "PhysicalFont", "strip_suffixes"]

#: The suffix of a no-bold-face alias; a registry lookup strips it, as docx4j's
#: ``PhysicalFonts.get`` does.
NOBOLD_SUFFIX = "+nobold"

#: The suffixes docx4j's FO layer stacks on a font family for a twin declared from
#: the same file: the kerned twin, the twin with no OpenType features, and the
#: no-bold-face alias. Nothing here produces the first two; they are stripped so
#: that a name written by docx4j's FO output resolves against a registry built here.
_SUFFIXES = ("+kerned", "+noliga", NOBOLD_SUFFIX)


def strip_suffixes(name: str) -> str:
    """The physical font name without the suffixes an FO layer adds.

    They stack, so all of them are stripped.
    """
    key = name
    stripped = True
    while stripped:
        stripped = False
        for suffix in _SUFFIXES:
            if key.endswith(suffix):
                key = key[: -len(suffix)]
                stripped = True
    return key


@dataclasses.dataclass(frozen=True, slots=True)
class PhysicalFont:
    """A font available to the consumer.

    Attributes:
        name: the face's full name, as a registry keys it. docx4j registers the
            first of FOP's triplets, the name table's id 4 --- ``Carlito
            Regular``, ``Arimo Bold``.
        family_name: the family a ``font-family`` names (``Carlito``); the name
            with its twin suffixes stripped where the registry does not say.
        no_bold_face: True for a :meth:`no_bold_face_alias`.
    """

    name: str
    family_name: str = ""
    no_bold_face: bool = False

    def __post_init__(self) -> None:
        """Default the family to the name with its suffixes stripped."""
        if not self.family_name:
            object.__setattr__(self, "family_name", strip_suffixes(self.name))

    def no_bold_face_alias(self) -> PhysicalFont:
        """This font again under ``name + '+nobold'``, reporting no bold face of its own.

        So that a consumer synthesises bold from this file at its own advances
        rather than taking the family's real bold --- which is what Word does for
        a family that has none (docx4j CR-016 probe ``fonts-light-bold``: Word
        draws a ``w:b`` run in Calibri Light with the Calibri Light font object,
        378.55 pt for a sentence its regular sets in 377.50, where Carlito Bold
        set it in 388.70). Held by the mapper, never put in a registry --- it
        stands for one document's font.
        """
        return PhysicalFont(self.name + NOBOLD_SUFFIX, self.family_name, True)

    def __str__(self) -> str:
        """The face's full name."""
        return self.name
