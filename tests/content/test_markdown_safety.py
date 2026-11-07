"""CR-015 through the readers, reader by reader, and Phase C's sinks: an address comment, the outline, a picture's name."""

from __future__ import annotations

import base64
import logging
from collections.abc import Callable
from typing import Any

import pytest

from docx4j_py import create_package
from docx4j_py.model.content.errors import AddressError
from docx4j_py.model.markdown import ADDRESS_COMMENT
from docx4j_py.model.markdown.export import address_comment
from tests.text.mini_docx import MiniDocx

SCRIPT = "javascript:alert(1)"
NOTE = "onenote:///x"
IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
HEADER_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")
HOSTILE = "--><script>alert(1)</script><!--"


def document(target: str) -> Any:
    """A paragraph and a table cell, each linking to `target`."""
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("click here")
    paragraph.search("here")[0].hyperlink = target
    table = pkg.body.insert_table(1, 1, values=[["cell"]])
    table.cell(0, 0).body.paragraphs[0].get_range("Content").hyperlink = target
    return pkg


def trial_markdown(pkg: Any, **options: Any) -> str:
    with pkg.dry_run() as trial:
        return str(trial.to_markdown(**options))


READERS: dict[str, Callable[..., str]] = {
    "pkg.to_markdown": lambda pkg, **options: pkg.to_markdown(**options),
    "pkg.markdown_budget": lambda pkg, **options: pkg.markdown_budget(**options).text,
    "body.to_markdown": lambda pkg, **options: pkg.body.to_markdown(**options),
    "body.markdown_budget": lambda pkg, **options: pkg.body.markdown_budget(**options).text,
    "paragraph.to_markdown": lambda pkg, **options: pkg.body.paragraphs[0].to_markdown(**options),
    "table.to_markdown": lambda pkg, **options: pkg.body.tables[0].to_markdown(**options),
    "dry_run": trial_markdown,
}


@pytest.mark.parametrize("reader", READERS.values(), ids=READERS)
def test_a_reader_refuses_a_script_link(reader: Callable[..., str]) -> None:
    written = reader(document(SCRIPT))
    assert "javascript" not in written and "](" not in written


@pytest.mark.parametrize("reader", READERS.values(), ids=READERS)
def test_a_reader_keeps_an_added_scheme(reader: Callable[..., str]) -> None:
    pkg = document(NOTE)
    assert "onenote" not in reader(pkg)
    assert f"]({NOTE})" in reader(pkg, link_schemes=("onenote",))


@pytest.mark.parametrize("reader", READERS.values(), ids=READERS)
def test_a_reader_warns_once(reader: Callable[..., str], caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.model.text.markdown"):
        reader(document(SCRIPT))
    assert len([record for record in caplog.records if "Link target refused" in record.getMessage()]) == 1


def test_a_reader_marks_a_change_with_its_type_alone() -> None:
    pkg = MiniDocx.body('<w:p><w:ins w:id="1" w:author="&lt;b&gt;A"><w:r><w:t>NEW</w:t></w:r></w:ins></w:p>').load()
    assert pkg.body.to_markdown(view="markup").strip() == "{++NEW++}"


# --- the address comments (Q8) ---------------------------------------------------------


def with_para_id(para_id: str) -> Any:
    escaped = para_id.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return MiniDocx.body(
        f'<w:p w14:paraId="{escaped}"><w:r><w:t>Hostile</w:t></w:r></w:p><w:p w14:paraId="5A2B1C3D"><w:r><w:t>Fine</w:t></w:r></w:p>'
    ).load()


def test_a_malformed_para_id_is_no_address() -> None:
    pkg = with_para_id(HOSTILE)
    written = pkg.to_markdown(addresses=True)
    assert "<script>" not in written
    assert "<!-- body/0 -->" in written and "<!-- w14:5A2B1C3D -->" in written


def test_addressing_and_the_comment_agree() -> None:
    pkg = with_para_id(HOSTILE)
    [address, fine] = ADDRESS_COMMENT.findall(pkg.body.to_markdown(addresses=True))
    assert pkg.body.paragraphs[0].address == address == "body/0"
    assert pkg.body.element_at(address).text == "Hostile"
    assert pkg.body.element_at(fine).text == "Fine"
    with pytest.raises(AddressError) as caught:
        pkg.body.element_at(f"w14:{HOSTILE}")
    assert caught.value.code == "address.not_found"


@pytest.mark.parametrize("para_id", ["ABC", "5A2B1C3D0", "5A2B1C3G", "5A2B-C3D"])
def test_a_para_id_is_eight_hex_digits(para_id: str) -> None:
    assert with_para_id(para_id).body.paragraphs[0].address == "body/0"


def test_lower_case_hex_is_a_para_id() -> None:
    assert with_para_id("5a2b1c3d").body.paragraphs[0].address == "w14:5a2b1c3d"


@pytest.mark.parametrize(
    ("address", "written"),
    [
        ("w14:5A2B1C3D", "<!-- w14:5A2B1C3D -->"),
        ("body/3/0/1", "<!-- body/3/0/1 -->"),
        ("header:rId8/0", "<!-- header:rId8/0 -->"),
        ("body/3/textbox:0/1", "<!-- body/3/textbox:0/1 -->"),
        ("a-b_c.d", "<!-- a-b_c.d -->"),
        ("a--b", "<!-- a-%2Db -->"),
        ("a---b", "<!-- a-%2D%2Db -->"),
        ("-->", "<!-- -%2D%3E -->"),
        ("a b", "<!-- a%20b -->"),
        ("é", "<!-- %C3%A9 -->"),
    ],
)
def test_an_address_can_t_close_its_comment(address: str, written: str) -> None:
    assert address_comment(address) == written
    assert ADDRESS_COMMENT.fullmatch(written)


def test_a_note_s_malformed_id_stays_in_its_comment() -> None:
    made = MiniDocx.body('<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>')
    escaped = HOSTILE.replace("<", "&lt;").replace(">", "&gt;")
    made.footnotes(f'<w:footnote w:id="{escaped}"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote>')
    assert "<script>" not in made.load().to_markdown(addresses=True)


@pytest.mark.parametrize(("address", "code"), [("w14:%2D%3E", "address.not_found"), ("footnote:-%2D%3E/0", "address.read_only")])
def test_a_percent_address_finds_nothing(address: str, code: str) -> None:
    """A note's address is read only whatever it holds, so a guarded one stays so."""
    with pytest.raises(AddressError) as caught:
        with_para_id("5A2B1C3E").body.element_at(address)
    assert caught.value.code == code


# --- the outline (Q9) ------------------------------------------------------------------


def test_the_outline_escapes_its_entries() -> None:
    pkg = create_package()
    heading = pkg.body.insert_paragraph("<script>alert(1)</script> heading")
    heading.style_built_in = "Heading1"
    pkg.body.insert_paragraph("[x](javascript:alert(2)) and <img src=x onerror=alert(3)>")
    written = pkg.body.outline().to_markdown()
    assert r"# \<script\>alert(1)\</script\> heading" in written
    assert r"- \[x\](javascript:alert(2)) and \<img src=x onerror=alert(3)\>" in written


def test_a_table_s_and_a_control_s_label_are_escaped_and_the_size_isn_t() -> None:
    pkg = create_package()
    pkg.body.insert_table(1, 2, values=[["<b>", "two"]])
    written = pkg.body.outline().to_markdown()
    assert r"table 1x2: \<b\>" in written


def test_the_outline_s_section_comment_can_t_be_closed() -> None:
    """A header's prefix holds its relationship's id, a document value: found while building, not in the CR's census."""
    header = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:hdr>'
    rel = "a--&gt;&lt;script&gt;"
    section = f'<w:sectPr><w:headerReference w:type="default" r:id="{rel}"/></w:sectPr>'
    made = MiniDocx.body(f"<w:p><w:r><w:t>Body</w:t></w:r></w:p>{section}").part("header1.xml", HEADER_TYPE, HEADER, rel, header)
    written = made.load().outline().to_markdown()
    assert "<script>" not in written and "Top" in written


# --- a picture's destination (Q10) -----------------------------------------------------


def picture(rel_target: str) -> Any:
    drawing = (
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="n" descr="alt"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="1" name="n"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>'
        "<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    made = MiniDocx.body(drawing)
    made.entry(rel_target, "image/png", "")
    made.parts[rel_target] = PNG
    made.relationship("rId7", IMAGE, f"../{rel_target}")
    return made.load()


def test_a_picture_s_part_name_is_held_to_the_rule(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.model.text.markdown"):
        written = picture("javascript:alert(1).png").body.to_markdown()
    assert written.strip() == "![alt]()"
    assert [record for record in caplog.records if "Link target refused" in record.getMessage()]


def test_an_ordinary_picture_keeps_its_name() -> None:
    assert picture("word/media/image1.png").body.to_markdown().strip() == "![alt](media/image1.png)"


def test_a_host_s_own_destination_is_the_host_s() -> None:
    from docx4j_py.model.text import extract, markdown

    written = markdown(extract(picture("word/media/image1.png")), image_destination=lambda image: "s3:bucket/x.png")
    assert "![alt](s3:bucket/x.png)" in written
