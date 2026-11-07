"""Which link targets the Markdown writer writes as a link (CR-015). docx4j ``LinkTargets`` at ``525d1ccb3``.

A document's ``w:hyperlink`` relationship target, or a HYPERLINK field's, reaches a link destination as the document has
it, so a crafted ``javascript:alert(1)`` target runs script on the page that shows the Markdown (CWE-79, the defect
apache/poi#1455 fixed in POI). The rule is an allowlist of schemes, as POI chose: a target with no scheme (a relative
path, a fragment, a UNC path), a Windows drive path, and :data:`DEFAULT_SCHEMES`, plus what a caller adds. A refused
target gives the link's text with no link.

The scheme is read as a browser reads one: leading controls and spaces are ignored and tabs and line breaks dropped
anywhere, so `` \\x01java\\tscri\\npt:`` is ``javascript:``, compared without case. Anything up to the first ``:`` before
any ``/``, ``?`` or ``#`` is the scheme, so a malformed one is refused, not resolved as a relative URL.

The target is also read with its character references decoded as a CommonMark renderer decodes a link destination
(:func:`decode_references`): ``javascript&#58;alert(1)`` passes the first reading and comes out of a renderer as
``href="javascript:alert(1)"`` (CR-015 section 2.5). A target is safe only if both readings are.
"""

from __future__ import annotations

import re
from collections.abc import Iterable

__all__ = ["DEFAULT_SCHEMES", "decode_references", "excerpt", "is_safe", "refused_message", "scheme", "schemes"]

#: docx4j's default schemes (POI's set); docx4j ``defaultSchemes()``.
DEFAULT_SCHEMES: tuple[str, ...] = ("http", "https", "ftp", "mailto", "file", "news", "tel")

#: The HTML5 named references that stand for one ASCII character. The rest decode to characters no scheme can hold, and
#: are left as written.
_ASCII_NAMED: dict[str, str] = {
    "Tab": "\t", "excl": "!", "quot": '"', "QUOT": '"', "num": "#", "dollar": "$", "percnt": "%", "amp": "&",
    "AMP": "&", "apos": "'", "lpar": "(", "rpar": ")", "ast": "*", "midast": "*", "plus": "+", "comma": ",",
    "period": ".", "sol": "/", "colon": ":", "semi": ";", "lt": "<", "LT": "<", "equals": "=", "gt": ">", "GT": ">",
    "quest": "?", "commat": "@", "lbrack": "[", "lsqb": "[", "bsol": "\\", "rbrack": "]", "rsqb": "]", "Hat": "^",
    "lowbar": "_", "UnderBar": "_", "DiacriticalGrave": "`", "grave": "`", "lbrace": "{", "lcub": "{", "verbar": "|",
    "vert": "|", "VerticalLine": "|", "rbrace": "}", "rcub": "}", "NewLine": "\n",
}  # fmt: skip
_DECIMAL = re.compile(r"[0-9]{1,7}")
_HEX = re.compile(r"[0-9A-Fa-f]{1,6}")


def schemes(extra: str | Iterable[str] = ()) -> tuple[str, ...]:
    """The defaults, then each added scheme trimmed and lower-cased, a trailing ``:`` dropped, an empty one ignored.

    Args:
        extra: an iterable of schemes, or one string split on commas, as docx4j's property is. A default is never removed.
    """
    entries = extra.split(",") if isinstance(extra, str) else extra
    out = list(DEFAULT_SCHEMES)
    for entry in entries:
        added = entry.strip().lower().removesuffix(":")
        if added and added not in out:
            out.append(added)
    return tuple(out)


def _normalise(target: str) -> str:
    """Leading controls and spaces removed; tabs and line breaks removed anywhere. docx4j ``normalise``."""
    kept: list[str] = []
    for character in target:
        if character in "\t\n\r" or (character <= " " and not kept):
            continue
        kept.append(character)
    return "".join(kept)


def scheme(target: str | None) -> str | None:
    """The scheme as a browser reads it, lower-cased; None for none (a relative path, a fragment, a UNC path, nothing)."""
    if target is None:
        return None
    read = _normalise(target)
    for index, character in enumerate(read):
        if character == ":":
            return read[:index].lower() if index else None
        if character in "/?#":
            return None
    return None


def _is_drive_letter(read: str) -> bool:
    """A letter in the Basic Multilingual Plane, what Java's ``Character.isLetter(char)`` and core-ts's test can see as one."""
    return len(read) == 1 and read.isalpha() and ord(read) <= 0xFFFF


def _safe_as_written(target: str, allowed: tuple[str, ...]) -> bool:
    """:func:`is_safe` for one reading of the target."""
    read = scheme(target)
    if read is None:
        return True
    if _is_drive_letter(read):
        rest = _normalise(target)
        return rest[2:3] in ("\\", "/")
    return read in allowed


def is_safe(target: str | None, extra: str | Iterable[str] = ()) -> bool:
    """Whether a link may go to `target`: safe as written, and with its character references decoded.

    Safe means no scheme, a drive path (``C:\\`` or ``C:/``), or a scheme in :func:`schemes`. None and ``""`` are safe,
    there being nothing to write.
    """
    if not target:
        return True
    allowed = schemes(extra)
    if not _safe_as_written(target, allowed):
        return False
    decoded = decode_references(target)
    return decoded == target or _safe_as_written(decoded, allowed)


def _reference(name: str) -> str | None:
    """What the reference between ``&`` and ``;`` stands for, or None to leave it as written."""
    if len(name) > 1 and name[0] == "#":
        hexadecimal = name[1] in "xX"
        digits = name[2:] if hexadecimal else name[1:]
        if (_HEX if hexadecimal else _DECIMAL).fullmatch(digits) is None:
            return None
        code_point = int(digits, 16 if hexadecimal else 10)
        if code_point == 0 or code_point > 0x10FFFF or 0xD800 <= code_point <= 0xDFFF:
            return "\ufffd"
        return chr(code_point)
    return _ASCII_NAMED.get(name)


def decode_references(target: str) -> str:
    """`target` with its character references decoded once, as CommonMark decodes a link destination.

    A decimal or hex reference gives its character, U+FFFD for zero or out of range, and a name standing for one ASCII
    character gives that character. Other names, and an ``&`` beginning no reference, stay as written. Not
    ``html.unescape``, which decodes every HTML5 name and some with no semicolon.
    """
    if "&" not in target:
        return target
    out: list[str] = []
    index = 0
    while index < len(target):
        character = target[index]
        end = target.find(";", index) if character == "&" else -1
        decoded = _reference(target[index + 1 : end]) if 0 <= end - index <= 40 else None
        if decoded is None:
            out.append(character)
            index += 1
        else:
            out.append(decoded)
            index = end + 1
    return "".join(out)


def excerpt(target: str | None) -> str:
    """`target` for a log line: at most 80 characters and ``...``, controls as spaces, so it can't forge a line."""
    if target is None:
        return "null"
    shown = f"{target[:80]}..." if len(target) > 80 else target
    return "".join(" " if character < " " else character for character in shown)


def refused_message(target: str | None) -> str:
    """The writer's log line for a refused target."""
    return f"Link target refused (scheme not in link_schemes): the text is written without a link: {excerpt(target)}"
