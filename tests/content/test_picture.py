"""``InlinePicture`` and the two insert verbs. CR-003 Phase C, section 3.2.

The rules of section 4 that land here: ``/word/media/imageN.<ext>`` with N free
for that extension (docx4j's ``getNewPartName``), a relationship from **the part
the body belongs to**, the ``wp:inline`` of Phase A's ``inline_picture``, the
four header readers and no Pillow, and an image wider than the text area scaled
down as ``CxCy.scale`` does.
"""

from __future__ import annotations

import base64
import zipfile

import pytest
from conftest import SAMPLES, part_bytes, reloaded, sample

from docx4j_py import load
from docx4j_py.model.content import BuilderError, ContentError, InlinePicture


@pytest.fixture(scope="module")
def png() -> bytes:
    """The PNG out of ``samples/Images.docx``: 340 by 248, no declared density."""
    with zipfile.ZipFile(SAMPLES / "Images.docx") as archive:
        name = next(n for n in archive.namelist() if n.lower().endswith(".png"))
        return archive.read(name)


# ---------------------------------------------------------------------------
# inserting
# ---------------------------------------------------------------------------


def test_insert_inline_picture_adds_the_part_the_relationship_and_the_drawing(new_package, png):
    picture = new_package.body.insert_inline_picture(
        png, alt_text_description="A pangolin", alt_text_title="Pangolin"
    )

    assert isinstance(picture, InlinePicture)
    assert str(picture.image_part.part_name) == "/word/media/image1.png"
    assert picture.image_part.content_type == "image/png"
    assert picture.get_bytes() == png
    assert picture.get_base64() == base64.b64encode(png).decode("ascii")
    assert picture.image_format == "Png"
    assert picture.alt_text_description == "A pangolin"
    assert picture.alt_text_title == "Pangolin"

    # the relationship is on the part the body belongs to
    main = new_package.main_document_part
    assert main.relationships_part.get_relationship_by_id(picture.rel_id) is not None
    assert main.relationships_part.get_part(picture.rel_id) is picture.image_part

    # the report names both parts it wrote
    change = new_package.last_change
    assert change.operation == "insert_inline_picture"
    assert "/word/media/image1.png" in change.parts_touched
    assert "/word/_rels/document.xml.rels" in change.parts_touched

    back = reloaded(new_package)
    again = back.body.inline_pictures[0]
    assert again.get_bytes() == png
    assert again.alt_text_title == "Pangolin"
    assert again.image_format == "Png"


def test_the_part_name_is_the_first_free_n_for_that_extension(new_package, png):
    first = new_package.body.insert_inline_picture(png)
    second = new_package.body.insert_inline_picture(png)

    assert str(first.image_part.part_name) == "/word/media/image1.png"
    assert str(second.image_part.part_name) == "/word/media/image2.png"
    assert first.rel_id != second.rel_id
    assert len(new_package.body.inline_pictures) == 2


def test_the_size_comes_from_the_header_and_scales_to_the_text_column(new_package, png):
    from docx4j_py.wml import emu_for, image_size

    info = image_size(png)
    assert (info.width_px, info.height_px) == (340, 248)
    assert not info.has_density, "no pHYs: 96 dpi is assumed, as docx4j and Word do"

    picture = new_package.body.insert_inline_picture(png)
    natural = emu_for(info, max_width_emu=None)
    assert picture.inline.extent.cx == natural.cx
    assert picture.width == pytest.approx(natural.cx / 12700)

    # width alone keeps the aspect ratio; width and height are taken as given
    picture.width = 200
    assert picture.width == 200.0
    assert picture.height == pytest.approx(200 * natural.cy / natural.cx, rel=1e-3)
    picture.height = 50
    assert picture.height == 50.0

    sized = new_package.body.insert_inline_picture(png, width=72, height=36)
    assert (sized.width, sized.height) == (72.0, 36.0)
    assert sized.inline.extent.cx == 72 * 12700


def test_the_wp_inline_matches_what_word_writes(new_package, png):
    picture = new_package.body.insert_inline_picture(png, name="logo.png")
    inline = picture.inline

    assert inline.doc_pr.id == 1
    assert inline.doc_pr.name == "logo.png"
    assert inline.effect_extent.l == 0
    assert inline.c_nv_graphic_frame_pr.graphic_frame_locks.no_change_aspect is True
    assert inline.extent.cx == picture.inline.extent.cx
    # pic:cNvPr carries the file name, as Word writes it
    assert picture._pic().nv_pic_pr.c_nv_pr.name == "image1.png"

    # a second picture in the same part gets the next docPr id
    assert new_package.body.insert_inline_picture(png).inline.doc_pr.id == 2


def test_insert_into_a_paragraph_at_start_end_and_replace(new_package, png):
    paragraph = new_package.body.insert_paragraph("before after")
    paragraph.insert_inline_picture(png, location="Start", width=10)
    assert len(paragraph.inline_pictures) == 1
    assert paragraph.text == "before after", "a drawing carries no text"

    paragraph.insert_inline_picture(png, location="End", width=10)
    assert len(paragraph.inline_pictures) == 2

    paragraph.insert_inline_picture(png, location="Replace", width=10)
    assert len(paragraph.inline_pictures) == 1
    assert paragraph.text == ""

    with pytest.raises(ContentError) as raised:
        paragraph.insert_inline_picture(png, location="Before")
    assert raised.value.code == "location.invalid"

    assert len(reloaded(new_package).body.inline_pictures) == 1


def test_the_base64_twins(new_package, png):
    encoded = base64.b64encode(png).decode("ascii")
    one = new_package.body.insert_inline_picture_from_base64(encoded)
    paragraph = new_package.body.insert_paragraph("x")
    two = paragraph.insert_inline_picture_from_base64(encoded)

    assert one.get_bytes() == png
    assert two.get_bytes() == png


def test_an_image_no_reader_understands_names_the_four_formats(new_package):
    with pytest.raises(BuilderError) as raised:
        new_package.body.insert_inline_picture(b"not an image at all")
    assert raised.value.code == "image.unsupported_format"
    assert "PNG, JPEG, GIF or BMP" in raised.value.message
    assert "/word/media/image1.png" not in [str(name) for name in new_package.parts]


def test_a_body_with_no_part_refuses(png):
    from docx4j_py.model.content import Body
    from docx4j_py.wml import Body as BodyElement

    body = Body(None, BodyElement())
    with pytest.raises(ContentError) as raised:
        body.insert_inline_picture(png)
    assert raised.value.code == "picture.no_part"


# ---------------------------------------------------------------------------
# reading a document that has pictures
# ---------------------------------------------------------------------------


def test_reading_the_pictures_of_a_loaded_document():
    package = sample("Images.docx")
    pictures = package.body.inline_pictures

    assert len(pictures) == 2
    for picture in pictures:
        assert picture.image_format in ("Png", "Jpeg", "Gif", "Bmp")
        assert picture.width > 0
        assert len(picture.get_bytes()) > 100
        assert picture.image_part is not None
        assert picture.to_dict()["rel_id"] == picture.rel_id

    # Paragraph.inline_pictures and Body.inline_pictures agree
    from_paragraphs = [
        picture for paragraph in package.body.paragraphs for picture in paragraph.inline_pictures
    ]
    assert from_paragraphs == pictures


def test_delete_takes_the_run_with_it(new_package, png):
    paragraph = new_package.body.insert_paragraph("text")
    picture = paragraph.insert_inline_picture(png, width=10)
    runs = len(paragraph.runs)

    picture.delete()
    assert paragraph.inline_pictures == []
    assert len(paragraph.runs) == runs - 1
    assert paragraph.text == "text"
    assert reloaded(new_package).body.inline_pictures == []


def test_views_compare_equal_hash_and_repr(new_package, png):
    picture = new_package.body.insert_inline_picture(png, width=100, alt_text_description="alt")
    again = new_package.body.inline_pictures[0]

    assert picture == again
    assert hash(picture) == hash(again)
    assert repr(picture).startswith("<InlinePicture ")
    assert "Png" in repr(picture) and "'alt'" in repr(picture)
    assert picture.to_dict()["image_format"] == "Png"


def test_a_picture_in_a_header_is_related_from_the_header(png):
    package = sample("Headers.docx")
    header = package.header_parts()[0]
    picture = header.body.insert_inline_picture(png, width=20)

    assert header.relationships_part.get_part(picture.rel_id) is picture.image_part
    main = package.main_document_part
    assert main.relationships_part.get_relationship_by_id(picture.rel_id) is not picture.rel_id

    back = load(package.save())
    header_back = next(p for p in back.header_parts() if p.part_name == header.part_name)
    assert header_back.body.inline_pictures[0].get_bytes() == png


def test_only_the_touched_parts_are_rebuilt(png):
    package = sample("2010-sample1.docx")
    _ = package.body
    before = part_bytes(package.save())

    package.body.insert_inline_picture(png, width=20)
    after = part_bytes(package.save())

    assert "word/media/image1.png" in after
    # the image's content type is a new Override, so [Content_Types].xml moves
    assert b"/word/media/image1.png" in after["[Content_Types].xml"]
    for name, data in before.items():
        if name not in ("word/document.xml", "[Content_Types].xml"):
            assert after[name] == data, name


def test_a_dry_run_un_adds_the_part_it_added(png):
    package = sample("2010-sample1.docx")
    _ = package.body
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        picture = trial.body.insert_inline_picture(png, width=20)
        assert str(picture.image_part.part_name) == "/word/media/image1.png"
        assert package.get_part("/word/media/image1.png") is not None

    # CR-003 section 12.5, answered by the trial's undo log
    assert package.get_part("/word/media/image1.png") is None
    assert package.content_type_manager.get_override_content_type("/word/media/image1.png") is None
    assert part_bytes(package.save(), relationships=True) == before
