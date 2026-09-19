"""``w:themeFontLang`` to a theme's ``script``. docx4j ``LanguageTagToScriptMapping``.

CR-002 section 15.3. Maps ``w:themeFontLang`` in the settings part ---
``<w:themeFontLang w:val="en-US" w:eastAsia="ko-KR"/>`` --- to the ``script``
attribute of an ``a:font`` in the theme part's font scheme
(``<a:font script="Jpan" typeface="ＭＳ 明朝"/>``), so that
:func:`~docx4j_py.model.fonts.theme_fonts.theme_font_of` can pick the face for
the document's language.

**By exact subtag.** Until docx4j 17.1.1 this was a substring test over
comma-separated lists, so Estonian (``et``) matched inside ``eth`` and took the
Ethiopic face (Nyala), Mongolian (``mn``) inside ``mni`` the Bengali one, and
Wolof (``wo``) inside ``bwo``. docx4j's CR-016 probe ``fonts-theme-lang``
measured Word: with ``<w:themeFontLang w:val="et-EE"/>`` and ``w:docDefaults``
naming ``minorHAnsi``, Word sets the text in Calibri, the theme's Latin face, on
every line.

:func:`to_english` is docx4j's ``CJKToEnglish``: certain fonts have CJK names in
a docx where a Windows box in an English locale has the English name. Font
discovery collects the English name where there is one, so that the mapper is
populated with the name the machine would have.
"""

from __future__ import annotations

__all__ = ["get_script_for_language_tag", "to_english"]


def _put(script: str, *langs: str) -> dict[str, str]:
    return {lang: script for lang in langs}


#: The script a primary language subtag selects in a theme's font list.
_SCRIPT_BY_LANGUAGE: dict[str, str] = {
    **_put("Jpan", "ja"),
    **_put("Hang", "ko"),
    **_put("Arab", "ar"),
    **_put("Hebr", "he", "yi", "iw"),
    **_put("Thai", "th"),
    **_put("Ethi", "ti", "bwo", "eth", "kxh", "mdy"),
    **_put("Beng", "bn", "as", "mni"),
    **_put("Gujr", "gu"),
    **_put("Khmr", "km"),
    **_put("Knda", "kn"),
    **_put("Guru", "pa"),
    **_put("Cans", "iu"),
    **_put("Cher", "chr"),
    # Yiii (Microsoft Yi Baiti) omitted, as in docx4j
    **_put("Tibt", "bo"),
    **_put("Thaa", "dv"),
    **_put("Deva", "hi", "ks", "kok", "mr", "ne", "sa", "sd"),
    **_put("Telu", "te"),
    **_put("Taml", "ta"),
    **_put("Syrc", "syr"),
    **_put("Orya", "or"),
    **_put("Mlym", "ml"),
    **_put("Laoo", "lo"),
    **_put("Sinh", "si"),
    # Mong (Mongolian Baiti) and Uigh (Microsoft Uighur) omitted, as in docx4j
    **_put("Viet", "vi", "lha", "nut"),
    **_put("Geor", "ka"),
}


def get_script_for_language_tag(lang_tag: str | None) -> str | None:
    """The theme's ``script`` for a language tag, or None where the tag selects none.

    The theme's ``a:latin``, ``a:ea`` or ``a:cs`` answers instead. Chinese is
    decided on the region as well: mainland China and Singapore use simplified
    characters (``Hans``), everywhere else traditional (``Hant``).
    """
    if lang_tag is None:
        return None
    lang = lang_tag.strip()
    position = lang.find("-")
    if position < 0:
        position = lang.find("_")
    if position > -1:
        lang = lang[:position]
    lang = lang.lower()

    if lang == "zh":
        tag = lang_tag.lower()
        return "Hans" if tag in ("zh-cn", "zh-sg") else "Hant"
    return _SCRIPT_BY_LANGUAGE.get(lang)


#: docx4j ``CJKToEnglish``: the CJK font names Windows has under an English name.
_TO_ENGLISH: dict[str, str] = {
    "ＭＳ ゴシック": "MS Gothic",  # <a:font script="Jpan" typeface="ＭＳ ゴシック"/>
    "ＭＳ 明朝": "MS Mincho",
    "맑은 고딕": "Malgun Gothic",  # <a:font script="Hang" typeface="맑은 고딕"/>
    "宋体": "SimSun",  # <a:font script="Hans" typeface="宋体"/>
    "新細明體": "PMingLiU",
}


def to_english(font_name: str) -> str | None:
    """The English name of a CJK font name, or None where this is not one we know."""
    return _TO_ENGLISH.get(font_name)
