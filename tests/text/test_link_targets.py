"""CR-015 Phase A: the link target allowlist, docx4j's ``LinkTargetsTest`` at ``525d1ccb3``, its strings kept."""

from __future__ import annotations

import pytest

from docx4j_py.model.text.link_targets import (
    DEFAULT_SCHEMES,
    decode_references,
    excerpt,
    is_safe,
    refused_message,
    scheme,
    schemes,
)


@pytest.mark.parametrize(
    "target",
    [
        None,
        "",
        "http://example.org/a",
        "HTTPS://example.org/b",
        "mailto:a@example.org",
        "ftp://example.org/c",
        "file:///C:/docs/d.docx",
        "tel:+61000000",
        "news:comp.lang.java",
        "sub/dir/page.html",
        "#bookmark1",
        "C:\\docs\\e.docx",
        "c:/docs/e.docx",
        "\\\\server\\share\\f.docx",
        "page.html?x=javascript:1",
        "page.html#javascript:1",
        "/root/path:with:colons",
        "  http://example.org/leading-spaces",
    ],
)
def test_safe_targets(target: str | None) -> None:
    assert is_safe(target)


@pytest.mark.parametrize(
    "target",
    [
        "javascript:alert(1)",
        "JavaScript:alert(1)",
        " \u0001java\tscri\npt:alert(2)",
        "\r\njavascript:alert(3)",
        "vbscript:MsgBox(1)",
        "data:text/html,<script>alert(1)</script>",
        "onenote:///x",
        "ms-word:ofe|u|https://example.org/x.docx",
        "C:",
        "java script:alert(1)",
        "about:blank",
        "blob:x",
    ],
)
def test_refused_targets(target: str) -> None:
    assert not is_safe(target)


@pytest.mark.parametrize(
    "target",
    [
        "javascript&#58;alert(1)",
        "&#106;avascript:alert(2)",
        "javascript&colon;alert(3)",
        "&#x6A;avascript:alert(4)",
        "&#X6a;avascript:alert(4)",
        "java&Tab;script:alert(5)",
        "java&NewLine;script:alert(6)",
        "&#0000106;avascript:alert(7)",
        "vb&#115;cript:x",
        "d&#97;ta:text/html,x",
        # refused as written already: up to the first ':' is the scheme, and these are no scheme a browser accepts
        "&unknown;avascript:x",
        "a&b:c/d",
    ],
)
def test_character_references_refused_once_decoded(target: str) -> None:
    assert not is_safe(target)


@pytest.mark.parametrize(
    "target",
    ["R&D/report.html", "page.html?a=1&amp;b=2", "https&#58;//example.org", "&colon", "page.html#&#58;", "sub/dir?x=&#58;"],
)
def test_character_references_kept(target: str) -> None:
    assert is_safe(target)


@pytest.mark.parametrize(
    ("target", "decoded"),
    [
        ("javascript&#58;alert(1)", "javascript:alert(1)"),
        ("&#x6A;avascript&colon;x", "javascript:x"),
        ("&amp;#58;", "&#58;"),
        ("&#0;", "\ufffd"),
        ("&#xFFFFFF;", "\ufffd"),
        ("&#xD800;", "\ufffd"),
        ("&colon", "&colon"),
        ("&nbsp;x", "&nbsp;x"),
        ("R&D", "R&D"),
    ],
)
def test_references_decoded_as_commonmark_decodes_them(target: str, decoded: str) -> None:
    assert decode_references(target) == decoded


def test_a_digit_commonmark_doesn_t_read_is_left_as_written() -> None:
    """Java's ``Character.digit`` takes any Unicode digit, and CommonMark an ASCII one: the rule is the renderer's."""
    assert decode_references("&#\u0663;") == "&#\u0663;"


@pytest.mark.parametrize(
    ("target", "read"),
    [
        (" \u0001java\tscri\npt:alert(2)", "javascript"),
        ("HTTP://x", "http"),
        ("sub/dir:x", None),
        ("#frag:x", None),
        (":x", None),
        ("C:\\x", "c"),
    ],
)
def test_scheme_as_a_browser_reads_it(target: str, read: str | None) -> None:
    assert scheme(target) == read


@pytest.mark.parametrize("extra", [" OneNote:, ms-word ,,notes", [" OneNote:", "ms-word ", "", "notes"]])
def test_added_schemes(extra: str | list[str]) -> None:
    assert set(DEFAULT_SCHEMES) <= set(schemes(extra))
    assert "onenote" in schemes(extra)
    assert is_safe("onenote:///x", extra)
    assert is_safe("ms-word:ofe|u|https://example.org/x.docx", extra)
    assert is_safe("Notes://server/db.nsf", extra)
    assert not is_safe("javascript:alert(1)", extra), "a default is never removed"
    assert not is_safe("onenote:///x"), "nothing added by a call without it"


def test_a_string_is_split_on_commas_never_read_as_characters() -> None:
    assert schemes("onenote") == (*DEFAULT_SCHEMES, "onenote")


def test_a_letter_outside_the_bmp_is_no_drive_letter() -> None:
    assert not is_safe("\U0001d49c:\\x")


def test_excerpt_for_the_log() -> None:
    assert excerpt("java\tscri\npt:alert(1)") == "java scri pt:alert(1)"
    assert len(excerpt("x" * 200)) == 83
    assert excerpt(None) == "null"


def test_the_message_names_the_option_and_the_target() -> None:
    expected = "Link target refused (scheme not in link_schemes): the text is written without a link: javascript:x"
    assert refused_message("javascript:x") == expected
