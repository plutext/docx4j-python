"""CR-015 Phases A and B: a document's values in the Markdown writer.

docx4j's ``MarkdownHostileLinksTest`` and ``MarkdownHostileMarkupTest`` at ``525d1ccb3``, and ``HostileDocx``'s targets,
their strings kept, then the paths section 2 measured here. A link whose scheme isn't allowed is its text, and a marker's
author and date are text, escaped as any document text is (WD REQ-703).
"""

from __future__ import annotations

import base64
import html
import logging
import re
from typing import Any

import pytest

from docx4j_py import create_package
from docx4j_py.model.text import extract, markdown, markdown_blocks, plain_text
from docx4j_py.model.text.link_targets import is_safe
from docx4j_py.model.text.markdown import MarkdownExporter, MarkdownExportOptions
from tests.text.mini_docx import MiniDocx

HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")

TARGETS = [
    "javascript:alert(1)",
    " \u0001java\tscri\npt:alert(2)",
    "vbscript:MsgBox(3)",
    "https://example.org/ok",
    "#bookmark1",
    "sub/page.html",
    "javascript&#58;alert(4)",
    "&#106;avascript:alert(5)",
    "javascript&colon;alert(6)",
    "&#x6A;avascript:alert(7)",
    "page.html?a=1&amp;b=2",
    "R&D/report.html",
    "https&#58;//example.org/entity",
]
#: docx4j ``HostileDocx.BENIGN``.
BENIGN = [
    "http://example.org/a",
    "https://example.org/b",
    "mailto:a@example.org",
    "ftp://example.org/c",
    "file:///C:/docs/d.docx",
    "tel:+61000000",
    "news:comp.lang.java",
    "sub/dir/page.html",
    "#bookmark1",
    "C:\\docs\\e.docx",
    "\\\\server\\share\\f.docx",
]
CUSTOM_LINK = "onenote:///C:/notes/section.one#page"


def java_trim(text: str) -> str:
    """Java's ``trim()``: every character at or below U+0020 from each end."""
    return text.strip("".join(chr(code) for code in range(0x21)))


def linked(*targets: str) -> Any:
    """docx4j's ``docx()``: a paragraph a target, ``text for`` and the target its words, built in memory (U+0001 can't be saved)."""
    pkg = create_package()
    for target in targets:
        paragraph = pkg.body.insert_paragraph(f"text for {java_trim(target)}")
        paragraph.get_range("Content").hyperlink = target
    return pkg


def destinations(written: str) -> list[str]:
    """Every link destination in `written`, angle brackets off."""
    found = re.findall(r"\]\(<((?:[^>\\]|\\.)*)>\)|\]\(([^)<\s]*)\)", written)
    return [first or second for first, second in found]


def test_no_script_scheme_destination() -> None:
    written = markdown(extract(linked(*TARGETS)))
    assert "(javascript:" not in written and "(<javascript:" not in written
    assert "scri\npt:alert(2)>)" not in written and "scri&#10;pt:alert(2)>)" not in written
    assert "(vbscript:" not in written and "(<vbscript:" not in written
    assert "text for javascript:alert(1)" in written
    assert "[text for javascript:alert(1)]" not in written


def test_every_destination_is_safe_once_decoded() -> None:
    """This package's own: docx4j's check of the obfuscated target above looks for a line feed the writer doesn't write."""
    found = destinations(markdown(extract(linked(*TARGETS))))
    assert found
    assert all(is_safe(html.unescape(target)) for target in found), found


@pytest.mark.parametrize("target", ["javascript&#58;alert(4)", "&#106;avascript:alert(5)", "javascript&colon;alert(6)", "&#x6A;avascript:alert(7)"])
def test_no_entity_encoded_script_scheme_destination(target: str) -> None:
    written = markdown(extract(linked(*TARGETS)))
    assert f"]({target}" not in written and f"](<{target}" not in written
    as_text = "text for " + target.replace("&", "\\&")
    assert as_text in written, written
    assert f"[{as_text}]" not in written


def test_benign_links_kept() -> None:
    written = markdown(extract(linked(*TARGETS)))
    assert "(https://example.org/ok)" in written or "(<https://example.org/ok>)" in written
    for target in ["page.html?a=1&amp;b=2", "R&D/report.html", "https&#58;//example.org/entity"]:
        assert f"]({target})" in written or f"](<{target}>)" in written, target
    assert "(#bookmark1)" in written or "(<#bookmark1>)" in written
    assert "(sub/page.html)" in written or "(<sub/page.html>)" in written


@pytest.mark.parametrize("target", BENIGN)
def test_hostile_docx_s_benign_targets_stay_links(target: str) -> None:
    written = markdown(extract(linked(target)))
    assert "](" in written, written


def test_a_scheme_is_added_by_the_call() -> None:
    """docx4j's ``propertyAddsAScheme``: refused by default, linked with it, and a call without it is the restore."""
    pkg = linked(CUSTOM_LINK)
    assert "](" not in markdown(extract(pkg))
    assert f"]({CUSTOM_LINK})" in markdown(extract(pkg), link_schemes=["onenote"])
    assert f"]({CUSTOM_LINK})" in markdown(extract(pkg), link_schemes="onenote")
    assert f"]({CUSTOM_LINK})" in MarkdownExporter(MarkdownExportOptions(link_schemes=("onenote",))).export(pkg)
    assert f"]({CUSTOM_LINK})" in markdown_blocks(extract(pkg), link_schemes=("onenote",)).markdown
    assert "](" not in markdown(extract(pkg))


def test_a_text_option_can_t_add_one() -> None:
    """A tool's arguments reach ``extract`` by name, so an agent steered by a document can't add a scheme (Q2)."""
    with pytest.raises(Exception, match="link_schemes"):
        extract(linked(CUSTOM_LINK), link_schemes=["javascript"])


# --- every path to a destination -----------------------------------------------------


def hostile(body: str, *targets: tuple[str, str], **options: Any) -> str:
    made = MiniDocx.body(body)
    for rel_id, target in targets:
        made.external(rel_id, HYPERLINK, target)
    return markdown(extract(made.load(), **options))


def field(target: str, words: str) -> str:
    return (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> HYPERLINK "{target}" </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{words}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )


def link(rel_id: str, words: str, anchor: str = "") -> str:
    said = f' w:anchor="{anchor}"' if anchor else ""
    return f'<w:hyperlink r:id="{rel_id}"{said}><w:r><w:t>{words}</w:t></w:r></w:hyperlink>'


def test_a_complex_field() -> None:
    assert hostile(field("javascript:alert(6)", "field F1")).strip() == "field F1"


def test_a_simple_field() -> None:
    body = '<w:p><w:fldSimple w:instr=" HYPERLINK &quot;javascript:alert(7)&quot; "><w:r><w:t>simple F2</w:t></w:r></w:fldSimple></w:p>'
    assert hostile(body).strip() == "simple F2"


def test_a_target_and_an_anchor_are_refused_as_one() -> None:
    written = hostile(f"<w:p>{link('rId9', 'with anchor', 'frag')}</w:p>", ("rId9", "javascript:alert(1)"))
    assert written.strip() == "with anchor"


def test_an_anchor_alone_stays_a_link() -> None:
    assert hostile('<w:p><w:hyperlink w:anchor="bm1"><w:r><w:t>to it</w:t></w:r></w:hyperlink></w:p>').strip() == "[to it](#bm1)"


def test_a_cell() -> None:
    cell = f"<w:tc><w:p>{link('rId9', 'cell T1')}</w:p></w:tc>"
    written = hostile(f"<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>", ("rId9", "javascript:alert(1)"))
    assert "cell T1" in written and "](" not in written


def test_a_picture_in_a_refused_link_keeps_its_image() -> None:
    drawing = (
        '<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="n" descr="P3 inside link"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="1" name="n"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>'
        "<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    body = f'<w:p><w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">see </w:t></w:r>{drawing}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:hyperlink></w:p>'
    made = MiniDocx.body(body).part("media/image1.png", "image/png", IMAGE, "rId7", PNG).external("rId9", HYPERLINK, "javascript:alert(1)")
    written = markdown(extract(made.load()))
    assert written.startswith("see ![P3 inside link](data:image/png;base64,") and written.rstrip().endswith(") after")


def test_a_refused_link_keeps_its_formatting() -> None:
    body = '<w:p><w:hyperlink r:id="rId9"><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:hyperlink></w:p>'
    assert hostile(body, ("rId9", "javascript:alert(1)")).strip() == "**bold**"


def test_a_note_s_definition() -> None:
    made = MiniDocx.body('<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>')
    made.footnotes(f'<w:footnote w:id="2">{field("javascript:alert(1)", "note text")}</w:footnote>')
    written = markdown(extract(made.load()))
    assert "note text" in written and "javascript" not in written


def test_a_text_box_s_paragraph() -> None:
    box = (
        '<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p>'
        + link("rId9", "box words")
        + "</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"
    )
    written = hostile(box, ("rId9", "javascript:alert(1)"))
    assert "box words" in written and "](" not in written


# --- the logging ---------------------------------------------------------------------


def test_one_warning_an_export_then_debug(caplog: pytest.LogCaptureFixture) -> None:
    pkg = linked("javascript:one", "vbscript:two", "about:three")
    result = extract(pkg)
    with caplog.at_level(logging.DEBUG, logger="docx4j_py.model.text.markdown"):
        markdown(result)
    records = [record for record in caplog.records if record.name == "docx4j_py.model.text.markdown"]
    assert [record.levelno for record in records] == [logging.WARNING, logging.DEBUG, logging.DEBUG]
    assert records[0].getMessage().endswith("javascript:one")
    caplog.clear()
    with caplog.at_level(logging.DEBUG, logger="docx4j_py.model.text.markdown"):
        markdown(result)
    assert [record.levelno for record in caplog.records if record.name.endswith("text.markdown")][0] == logging.WARNING


def test_a_target_can_t_forge_a_log_line(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.model.text.markdown"):
        markdown(extract(linked("javascript:x\nFORGED LINE")))
    assert "\n" not in caplog.records[0].getMessage()


# --- a marker's author and date -------------------------------------------------------

AUTHOR = "<script>alert('A')</script>"
SAFE_AUTHOR = AUTHOR.replace("'", "&apos;").replace("<", "&lt;").replace(">", "&gt;")
ESCAPED = r"\<script\>alert('A')\</script\>"


def marked(tag: str, author: str, words: str, date: str = "") -> str:
    said = f' w:date="{date}"' if date else ""
    return f'<w:{tag} w:id="1" w:author="{author}"{said}><w:r><w:t>{words}</w:t></w:r></w:{tag}>'


def deleted(author: str, words: str) -> str:
    return f'<w:del w:id="2" w:author="{author}"><w:r><w:delText>{words}</w:delText></w:r></w:del>'


def in_markup(body: str, detail: str = "author", made: Any = None) -> str:
    made = made or MiniDocx.body(body)
    return markdown(extract(made.load(), view="markup", comments=True, markup_detail=detail))


def unescaped(markup: str, written: str) -> bool:
    """docx4j's helper: whether `markup` occurs other than with its ``<`` backslash-escaped."""
    return any(found.start() == 0 or written[found.start() - 1] != "\\" for found in re.finditer(re.escape(markup), written))


def hostile_document() -> MiniDocx:
    """docx4j's ``export``: T1 to T3 as text, T4 a link's, T5 a picture's, T6 an insertion's author, T7 and T8 a comment's."""
    def plain(words: str) -> str:
        return f'<w:p><w:r><w:t xml:space="preserve">{html.escape(words)}</w:t></w:r></w:p>'

    t6 = html.escape("T6 <script>alert('T6')</script>")
    t7 = html.escape("T7 <script>alert('T7')</script>")
    t8 = html.escape("T8 <script>alert('T8')</script> comment text")
    t5 = html.escape("T5 ](x) <b>alt</b> <script>alert('T5')</script>")
    t5_name = html.escape("T5\"><img src=x onerror=alert('T5')>")
    drawing = (
        f'<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="{t5_name}" descr="{t5}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="1" name="n"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>'
        "<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    body = (
        plain("T1 <script>alert('T1')</script> plain text")
        + plain("T2 [x](javascript:alert('T2')) as text")
        + plain("T3 <img src=x onerror=alert('T3')> as text")
        + f"<w:p>{link('rId9', html.escape(chr(84) + '4 a](javascript:alert(' + chr(39) + 'T4' + chr(39) + ')) [b'))}</w:p>"
        + drawing
        + f'<w:p>{marked("ins", t6, "inserted")}</w:p>'
        + '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="1"/>'
        + '<w:r><w:commentReference w:id="1"/></w:r></w:p>'
    )
    comments = (
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:comment w:id="1" w:author="{t7}"><w:p><w:r><w:t>{t8}</w:t></w:r></w:p></w:comment></w:comments>'
    )
    made = MiniDocx.body(body).external("rId9", HYPERLINK, "https://example.org/ok")
    return made.part("media/image1.png", "image/png", IMAGE, "rId7", PNG).part("comments.xml", COMMENTS_TYPE, COMMENTS, "rId8", comments)


@pytest.mark.parametrize("detail", ["type", "author", "author_date"])
def test_no_raw_markup_from_any_document_value(detail: str) -> None:
    written = in_markup("", detail, hostile_document())
    assert not unescaped("<script>", written), written
    assert not unescaped("<img ", written), written
    assert not unescaped("](javascript:", written), written
    assert r"T1 \<script\>alert('T1')\</script\> plain text" in written
    assert all(mark in written for mark in ("{++", "++}", "{>>", "<<}"))
    if detail != "type":
        assert r"{++T6 \<script\>alert('T6')\</script\>" in written
        assert r"{>>T7 \<script\>alert('T7')\</script\>" in written


def test_benign_markers_unchanged() -> None:
    written = in_markup("", "author", hostile_document())
    assert ": inserted++}" in written and "comment text<<}" in written


def test_a_deletion_and_a_cell() -> None:
    assert f"{{--{ESCAPED}: P2 GONE--}}" in in_markup(f"<w:p>{deleted(SAFE_AUTHOR, 'P2 GONE')}</w:p>")
    cell = f'<w:tc><w:p>{marked("ins", SAFE_AUTHOR, "T1 cell ins")}</w:p></w:tc>'
    assert f"{{++{ESCAPED}: T1 cell ins++}}" in in_markup(f"<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>")


def test_both_moves() -> None:
    gone = f'<w:moveFrom w:id="1" w:author="{SAFE_AUTHOR}"><w:r><w:t>P3 MOVED FROM</w:t></w:r></w:moveFrom>'
    came = f'<w:moveTo w:id="2" w:author="{SAFE_AUTHOR}"><w:r><w:t>P4 MOVED TO</w:t></w:r></w:moveTo>'
    written = in_markup(f"<w:p>{gone}</w:p><w:p>{came}</w:p>")
    assert ESCAPED in written and not unescaped("<script>", written)


def test_a_date_is_a_document_value_too() -> None:
    written = in_markup(f'<w:p>{marked("ins", "Ann", "NEW2", SAFE_AUTHOR)}</w:p>', "author_date")
    assert not unescaped("<script>", written), written


def test_an_author_can_t_write_markup() -> None:
    written = in_markup(f'<w:p>{marked("ins", "[x](javascript:alert(1)) *e* {--z--}", "NEW")}</w:p>')
    assert r"{++\[x\](javascript:alert(1)) \*e\* {--z--}: NEW++}" in written


def test_a_code_block_keeps_the_string_form() -> None:
    """In a fence nothing is markup, so a marker's detail stays as written, the fence longer than any run of backticks in it."""
    author = "``` <script>alert(1)</script>"
    styles = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="SourceCode"><w:name w:val="Source Code"/></w:style></w:styles>'
    )
    run = f'<w:ins w:id="1" w:author="{html.escape(author)}"><w:r><w:t>x = 1</w:t></w:r></w:ins>'
    body = f'<w:p><w:pPr><w:pStyle w:val="SourceCode"/></w:pPr>{run}</w:p>'
    made = MiniDocx.body(body).part(
        "styles.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "rId6",
        styles,
    )
    written = in_markup("", "author", made)
    assert written.startswith("````"), written
    assert f"{{++{author}: x = 1++}}" in written


def test_plain_text_writes_the_detail_as_it_is() -> None:
    """REQ-702: plain text isn't markup."""
    made = MiniDocx.body(f'<w:p>{marked("ins", SAFE_AUTHOR, "NEW")}</w:p>')
    assert AUTHOR in plain_text(extract(made.load(), view="markup", markup_detail="author"))


def test_the_type_detail_is_as_it_was() -> None:
    assert in_markup(f'<w:p>{marked("ins", SAFE_AUTHOR, "NEW")}</w:p>', "type").strip() == "{++NEW++}"
