#!/usr/bin/env python
"""Write the twelve artefacts the Word acceptance checklist needs.

CR-002 section 8: Word acceptance is manual. This produces the artefacts and
prints what to look for; `tests/README.md` is the checklist.

    .venv-fork/bin/python scripts/acceptance.py [--out out/acceptance]

Twelve artefacts in seventeen files (8 has an accepted and a rejected copy beside
it, 12 a tracked copy with those two), each testing a different half of the save path:

``1-untouched-round-trip.docx``
    loaded and saved with nothing unmarshalled. Every part is the source's
    bytes; only ``[Content_Types].xml`` and the ``.rels`` parts are written by
    this engine. If Word repairs this, the zip writer or the content types are
    wrong.
``2-remarshalled-round-trip.docx``
    the same document with the main part, the styles and the settings
    unmarshalled and **edited** --- a paragraph added through ``el`` and one
    through the ``p`` builder. Every one of those parts is re-serialised, so
    this is the test of the prefix table, ``mc:Ignorable`` and ``xml:space``.
``3-created.docx``
    ``WordprocessingMLPackage.create_package()`` with a heading and a formatted
    run. Nothing came from a container at all.
``4-created-with-image.docx``
    as 3, plus an image part added through ``add_target_part`` and placed in the
    body with the ``w:drawing`` CR-003 Phase A's ``inline_picture`` builds,
    sized by ``image_size`` and ``emu_for`` from the image's own header, so the
    relationship, the content type and the stored (not deflated) media entry
    are all exercised.
``5-markdown-built.docx``
    a document built from one markdown string through CR-003 Phase K's
    ``insert_markdown``: headings, emphasis, inline code, a nested bullet list,
    an ordered list, a block quote, a fenced code block, a hyperlink and a GFM
    pipe table. It exercises the two parts the markdown importer writes besides
    the body --- ``styles.xml``, which gains the styles docx4j's
    ``KnownStyles.xml`` supplies and the two code styles Word has no equivalent
    of, and ``numbering.xml``, which is created from nothing --- plus an
    external relationship for the link.
``6-tables-and-pictures.docx``
    a **loaded** document given, through CR-003 Phase C's content API, a table
    from ``insert_table`` with ``values=`` and ``style="TableGrid"``, a row
    added with ``add_rows``, a picture at a paragraph through
    ``insert_inline_picture``, and a flat OPC ``pkg:package`` fragment through
    ``insert_ooxml``. It exercises the image part under a free
    ``/word/media/imageN.png``, its relationship and its content type, the
    ``pkg:package`` reader, and a table whose grid is sized from the loaded
    document's own ``w:sectPr``. Since CR-003 section 21 it ends with a
    paragraph whose two links were set through ``Range.hyperlink``: an external
    one (a relationship of the main part) and an internal one to a bookmark on
    the "Tables and pictures" heading, which adds the ``Hyperlink`` character
    style to a ``styles.xml`` that did not define it.
``7-comments.docx``
    a **loaded** document (``samples/2010-sample1.docx``, which has no comment
    parts at all) given, through CR-003 Phase G's content API, a comment on a
    range found by ``find()``, a reply to it, a second comment resolved through
    ``w15:done``, and a third on a whole paragraph. All four comment parts are
    **created** with their relationships and content types, and the two comment
    styles are added to ``styles.xml``, so this is the test of the part-creating
    half of section 3.9.
``8-tracked-changes.docx``
    a **loaded** document edited through CR-003 Phase F's content API with the
    mode on: a replacement, an inserted paragraph, a paragraph appended after a
    Heading 1 (the final-mark rule and check 29's ``w:pPrChange``), a deleted
    paragraph, a bold run, a table row added and one deleted, and a comment on
    the replacement. Beside it ``8-tracked-changes-accepted.docx`` and
    ``8-tracked-changes-rejected.docx``: the same document resolved by CR-010's
    ``revisions.accept_all`` and ``reject_all``, each listing nothing.
``9-template-filled.docx``
    ``samples/invoice2013.docx`` --- a Word-authored template with twenty
    bindings over ``/customXml/item1.xml`` --- **filled** through CR-003 Phase
    E's ``pkg.custom_xml_parts.fill()``: a new company and invoice number, the
    VAT checkbox cleared, a new date, and **three** line items passed as a list
    to the repeating section's XPath. The document keeps its **one**
    ``w15:repeatingSectionItem``, because Word expands a bound repeating section
    to the node set when it opens the file (CR-003 section 17.10), so this is
    also the check that three rows appear. A new content control is inserted and
    bound to a **new** custom XML part added through ``add()``, which is the
    test of ``addPropertiesPart``'s relationship from the main document part.
``10-lists.docx``
    a **loaded** document given, through CR-003 Phase H's content API, a
    numbered list of three items from ``start_new_list()``, a fourth item
    attached at level 1, a second list restarted over the same definition with
    ``like=``, a bullet list from ``kind="Bullet"``, one list restyled with
    ``set_level_numbering(0, "LowerLetter", "%1)")``, a paragraph detached, and
    --- with tracking on --- one paragraph attached, which Word shows as a
    *Formatted* paragraph-numbering revision. It exercises the numbering part
    created from nothing with its relationship and its content type, the
    definitions copied from docx4j's default numbering, and the ``w:pPrChange``
    a numbering change records.
``11-python-docx-facade.docx``
    a document built **entirely through CR-003 Phase J's python-docx facade**
    (``docx4j_py.docx``), by code that would run unchanged against python-docx:
    the section 3.6 snippet's constructs, a Title and Heading 1 to 3, a
    ``List Bullet`` and a ``List Number`` paragraph, a page break, a second
    section in landscape with a header and a footer of its own, a table with a
    bold header row, a picture two inches wide, and the core properties. It
    exercises the facade's write path end to end --- ``ensure_style`` for the
    headings ``create_package()`` does not define, the header and footer
    **parts** created by ``is_linked_to_previous = False``, the cloned
    ``w:sectPr``, the image part, and ``dcterms:created`` written by the facade
    rather than by ``create_package``.
``12-ranges-across-paragraphs.docx``
    a **loaded** document (``samples/2010-sample1.docx``, as 8) given sixteen
    labelled cases of CR-013 Phase C's writes of a ``Range`` across paragraphs
    (sections 6.4, 13.5 and 13.7), each a ``Case N:`` paragraph and then its
    blocks: a range from a heading's middle to a body paragraph's deleted, one
    with a whole table between, two list items joined, a paragraph's
    ``'Whole'`` range deleted, a bold across two paragraphs, a split
    (``insert_paragraph`` after a range across two), a comment across
    paragraphs and one across a table, a block control wrapped around two
    paragraphs, a whole control deleted, a range into a control deleted, a
    ``Replace`` within one paragraph, a comment whose reference a deletion
    takes, a paragraph break in a complex field's result and in a
    ``w:fldSimple``'s (check 57's D1), and a range taking an equation whole
    (D2). ``12-ranges-across-paragraphs-tracked.docx`` is the same writes with
    the mode on (the whole control between Word's ``w:customXmlDelRange*``
    markers, the equation in Word's deleted form, the replacement's ``w:ins``
    first), and beside it ``-tracked-accepted.docx`` and
    ``-tracked-rejected.docx``, resolved by CR-010's engine, each listing
    nothing.
"""

from __future__ import annotations

import argparse
import datetime
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.openpackaging import (
    ImagePart,
    WordprocessingMLPackage,
)
from docx4j_py.wml import el, emu_for, image_size, inline_picture, p, r

SOURCE = ROOT / "samples" / "2016_image_with_text_effects.docx"
IMAGE_SOURCE = ROOT / "samples" / "Images.docx"
#: Artefact 6 edits a document that was **loaded**, not created.
SAMPLE_SOURCE = ROOT / "samples" / "2010-sample1.docx"
#: Artefact 9's template: Word-authored, twenty bindings, a checkbox, a date, a
#: picture control and two ``w15:repeatingSection``\ s.
TEMPLATE_SOURCE = ROOT / "samples" / "invoice2013.docx"

#: The width of the text column on A4 with 2.54 cm margins, in EMU: what an
#: image wider than the page is scaled down to (docx4j's ``CxCy.scale``).
TEXT_WIDTH_EMU = 5731510


def untouched(out: Path) -> Path:
    """1. Load and save; nothing is unmarshalled."""
    target = out / "1-untouched-round-trip.docx"
    with WordprocessingMLPackage.load(SOURCE) as pkg:
        pkg.save(target)
    return target


def remarshalled(out: Path) -> Path:
    """2. Unmarshal and edit the three parts Word is fussiest about."""
    target = out / "2-remarshalled-round-trip.docx"
    with WordprocessingMLPackage.load(SOURCE) as pkg:
        body = pkg.main_document_part.contents.body
        body.content.append(p("Added by docx4j-python through the p() builder."))
        body.content.append(
            el.p(
                content=[
                    el.r(
                        content=[
                            el.t("  and through el, with xml:space preserved.  "),
                        ]
                    )
                ]
            )
        )
        # unmarshal the two other parts Word validates hardest
        pkg.style_definitions_part.contents
        pkg.document_settings_part.contents
        pkg.save(target)
    return target


def created(out: Path) -> Path:
    """3. A document from nothing."""
    target = out / "3-created.docx"
    pkg = WordprocessingMLPackage.create_package()
    body = pkg.main_document_part.contents.body
    body.content.append(p("Created by docx4j-python", style="Heading1"))
    body.content.append(p("Hello World"))
    body.content.append(
        el.p(
            content=[
                r("bold red 14pt", bold=True, color="#FF0000", size=14),
                r(" and plain."),
            ]
        )
    )
    pkg.save(target)
    return target


def created_with_image(out: Path) -> Path:
    """4. A created document with an image part and a ``w:drawing`` that uses it."""
    target = out / "4-created-with-image.docx"
    with zipfile.ZipFile(IMAGE_SOURCE) as zf:
        name = next(n for n in zf.namelist() if n.lower().endswith(".png"))
        png = zf.read(name)

    pkg = WordprocessingMLPackage.create_package()
    main = pkg.main_document_part
    body = main.contents.body
    body.content.append(p("An image added through add_target_part:", style="Heading1"))

    image = ImagePart("/word/media/image1.png")
    image.set_bytes(png)
    rel = main.add_target_part(image)

    # CR-003 Phase A: the size comes from the image's own header and the
    # drawing from the builder, where both used to be a hand-written fragment
    # with the numbers guessed.
    size = emu_for(image_size(png), max_width_emu=TEXT_WIDTH_EMU)
    body.content.append(
        el.p(
            content=[
                r(
                    inline_picture(
                        rel.id,
                        cx=size.cx,
                        cy=size.cy,
                        id=1,
                        name="image1.png",
                        descr="A pangolin, from samples/Images.docx",
                    )
                )
            ]
        )
    )
    pkg.save(target)
    return target


#: Artefact 5's source. One string, every construct CR-003 section 3.5 maps.
MARKDOWN = """# Built from markdown

This document was built by `pkg.body.insert_markdown(...)`, with **bold**,
*italic*, ~~struck~~ and `inline code` in this paragraph.

## A bullet list, with a level of nesting

- the first item
- the second item
  - a nested item
  - another nested item
- the third item

## An ordered list

1. read the outline
2. find the text
3. edit by address
4. check the report

> A block quote, which becomes the Quote style.

```
x = 1
y = x + 1
```

## A table

| Region | Quarter | Total |
| --- | --- | ---: |
| North | Q1 | 120 |
| South | Q2 | 240 |
| East | Q3 | 360 |

See [the docx4j site](https://www.docx4java.org/) for the Java original.
"""


def markdown_built(out: Path) -> Path:
    """5. A document built from one markdown string (CR-003 Phase K)."""
    target = out / "5-markdown-built.docx"
    pkg = WordprocessingMLPackage.create_package()
    pkg.id_seed = 20260917  # the same seed, the same bytes
    pkg.body.insert_markdown(MARKDOWN)
    pkg.save(target)
    return target


#: Artefact 6's pasted fragment: a flat OPC package, as Word's clipboard writes
#: one, built here from a package this engine saved so that nothing is hand-rolled.
def flat_opc(data: bytes) -> str:
    """A saved package as a ``pkg:package`` document (docx4j ``FlatOpcXmlExporter``)."""
    import base64
    import io

    from lxml import etree

    from docx4j_py.openpackaging.stores import FlatOpcStore

    pkg_ns = FlatOpcStore.NAMESPACE
    root = etree.Element(f"{{{pkg_ns}}}package", nsmap={"pkg": pkg_ns})
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        types = etree.fromstring(archive.read("[Content_Types].xml"))
        defaults = {
            element.get("Extension").lower(): element.get("ContentType")
            for element in types
            if element.tag.endswith("Default")
        }
        overrides = {
            element.get("PartName"): element.get("ContentType")
            for element in types
            if element.tag.endswith("Override")
        }
        for name in archive.namelist():
            if name == "[Content_Types].xml":
                continue
            part_name = "/" + name
            content_type = overrides.get(part_name) or defaults.get(
                name.rpartition(".")[2].lower(), "application/octet-stream"
            )
            part = etree.SubElement(root, f"{{{pkg_ns}}}part")
            part.set(f"{{{pkg_ns}}}name", part_name)
            part.set(f"{{{pkg_ns}}}contentType", content_type)
            payload = archive.read(name)
            if "xml" in content_type:
                etree.SubElement(part, f"{{{pkg_ns}}}xmlData").append(etree.fromstring(payload))
            else:
                binary = etree.SubElement(part, f"{{{pkg_ns}}}binaryData")
                binary.text = base64.b64encode(payload).decode("ascii")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode()


def tables_and_pictures(out: Path) -> Path:
    """6. A loaded document edited through the Phase C content API."""
    target = out / "6-tables-and-pictures.docx"
    with zipfile.ZipFile(IMAGE_SOURCE) as zf:
        png = zf.read(next(n for n in zf.namelist() if n.lower().endswith(".png")))

    # the fragment a paste brings: a heading, a picture and a table, in a
    # package of its own, written out as flat OPC
    source = WordprocessingMLPackage.create_package()
    source.id_seed = 20260917
    source.body.insert_paragraph("Pasted from another document", style="Heading 2")
    source.body.insert_inline_picture(png, width=120, alt_text_description="the same pangolin")
    source.body.insert_table(2, 2, values=[["pasted", "table"], ["row", "two"]], style="TableGrid")
    package_xml = flat_opc(source.save())

    pkg = WordprocessingMLPackage.load(SAMPLE_SOURCE)
    pkg.id_seed = 20260917
    body = pkg.body

    heading = body.insert_paragraph("Tables and pictures", style="Heading 1")
    # the target of the internal link at the end (CR-003 section 21)
    heading.element.content.insert(0, el.bookmarkStart(id=100, name="tables_and_pictures"))
    heading.element.content.append(el.bookmarkEnd(id=100))
    table = body.insert_table(
        3,
        3,
        values=[
            ["Region", "Quarter", "Total"],
            ["North", "Q1", "120"],
            ["South", "Q2", "240"],
        ],
        style="TableGrid",
    )
    table.header_row_count = 1
    table.add_rows(1, values=[["East", "Q3", "360"]])

    caption = body.insert_paragraph("A picture inserted at this paragraph:")
    caption.insert_inline_picture(
        png, width=180, alt_text_description="A pangolin", alt_text_title="Pangolin"
    )

    body.insert_ooxml(package_xml)

    # CR-003 section 21: two links set through Range.hyperlink, on a document
    # whose styles.xml has no Hyperlink style until the first one is set
    links = body.insert_paragraph(
        "Set through Range.hyperlink: the docx4j site, and back to the tables."
    )
    links.search("the docx4j site")[0].hyperlink = "https://www.docx4java.org/"
    links.search("back to the tables")[0].hyperlink = "#tables_and_pictures"
    pkg.save(target)
    return target


def comments(out: Path) -> Path:
    """7. A loaded document commented on through the Phase G content API."""
    from docx4j_py.model.content import Author

    target = out / "7-comments.docx"
    pkg = WordprocessingMLPackage.load(SAMPLE_SOURCE)
    pkg.id_seed = 20260917
    pkg.author = Author("Claude", initials="C", email="claude@example.com")
    body = pkg.body

    # a comment on a range the agent found, and a reply in the same thread
    hit = pkg.find("first")[0]
    thread = hit.range(body).insert_comment(
        "Changed 'first' because the source document says 'red'."
    )
    thread.reply("Checked against the source; agreed.")

    # a resolved thread: Word shows it greyed out and marked done
    done = body.search("document")[0].insert_comment("Fixed, and resolved.")
    done.reply("Thanks.")
    done.resolved = True

    # and a comment on a whole paragraph
    paragraph = body.paragraphs[0]
    paragraph.insert_comment("A comment on the whole first paragraph.")

    pkg.save(target)
    return target


def tracked_changes(out: Path) -> list[Path]:
    """8. A loaded document edited through the Phase F content API, tracking on.

    With two copies beside it, the same document resolved by CR-010's engine:
    ``8-tracked-changes-accepted.docx`` (``revisions.accept_all``) and
    ``8-tracked-changes-rejected.docx`` (``revisions.reject_all``), each saved and
    read back listing nothing.
    """
    from docx4j_py.model import revisions
    from docx4j_py.model.content import Author

    target = out / "8-tracked-changes.docx"
    pkg = WordprocessingMLPackage.load(SAMPLE_SOURCE)
    pkg.id_seed = 20260917
    pkg.author = Author("Claude", initials="C", email="claude@example.com")
    body = pkg.body

    # something to delete and something to change the formatting of, written
    # *before* the mode goes on, so that they are the document's own content
    table = body.insert_table(
        3, 3, values=[["Region", "Quarter", "Total"], ["West", "Q2", "240"], ["East", "Q2", "180"]],
        style="TableGrid",
    )
    table.header_row_count = 1
    body.insert_paragraph("This paragraph will be deleted, with its mark.")
    body.insert_paragraph("This paragraph will be made bold.")
    # the last block of the body is a paragraph, so that the paragraph appended
    # below has one before it to carry its mark (CR-003 section 16.10), and a
    # Heading 1, so that the appended paragraph records the heading it took the
    # old final mark from, as Word's Enter after a heading does (check 29,
    # CR-010 section 12): Reject All must give the heading back as a heading
    body.insert_paragraph("A heading the appended paragraph follows", style="Heading 1")

    pkg.change_tracking_mode = "TrackAll"

    # a replacement, explained in a comment on what it changed
    body.replace_text("first", "second")
    hit = pkg.find("second")[0]
    hit.range(body).insert_comment("Changed 'first' to 'second': the source says 2010 was the second.")

    # an insertion where it can be seen --- after the first paragraph, not at the
    # very end --- and one appended, which exercises the final-mark rule of
    # CR-003 section 16.10 (the appended paragraph's own mark is left alone and
    # the mark before it, the heading's, is marked instead, as Word does) and
    # check 29's w:pPrChange on the appended paragraph, recording Heading 1
    body.paragraphs[0].insert_paragraph(
        "Added by an agent, right after the first paragraph.", location="After"
    )
    body.insert_paragraph("And one appended at the very end.")
    body.paragraph_at(contains="will be deleted").delete()
    body.paragraph_at(contains="made bold").font.bold = True
    table.add_rows(1, values=[["North", "Q2", "300"]])
    table.delete_rows(2)

    pkg.save(target)

    # the same document resolved by the engine, every story, as Word's Accept
    # All and Reject All resolve it (CR-010); each must list nothing
    written = [target]
    for suffix, resolve in (("accepted", revisions.accept_all), ("rejected", revisions.reject_all)):
        copy = WordprocessingMLPackage.load(target)
        resolve(copy)
        resolved = out / f"8-tracked-changes-{suffix}.docx"
        copy.save(resolved)
        left = revisions.tracked_changes(WordprocessingMLPackage.load(resolved))
        assert left == [], f"{resolved.name} still lists {left}"
        written.append(resolved)
    return written


def template_filled(out: Path) -> Path:
    """9. A Word-authored template filled through the Phase E custom XML API."""
    target = out / "9-template-filled.docx"
    pkg = WordprocessingMLPackage.load(TEMPLATE_SOURCE)
    # the invoice is a Word 2010 document (compatibilityMode 14); a loaded document's
    # mode is never changed silently, so say so here, which is also the one place
    # the setter is exercised in a Word check (CR-003 section 17.11)
    pkg.compatibility_mode = 15
    pkg.id_seed = 20260917

    # what the template wants, before anything is written: the same call
    # docx4j-mcp's describe_template makes
    skeleton = pkg.custom_xml_parts.describe()
    assert len(skeleton) == 20, len(skeleton)

    # and the fill: a text binding, the checkbox, the date, and **three** line
    # items as a list under the repeating section's own XPath. The engine writes
    # three data nodes and leaves the document's one template item alone; Word
    # clones it to the node set when it opens the file (section 17.10).
    result = pkg.custom_xml_parts.fill(
        {
            "/invoice[1]/customer[1]/company[1]": "Acme Manufacturing Ltd",
            "/invoice[1]/customer[1]/contact[1]": "Ada Lovelace",
            "/invoice[1]/invoicenumber[1]": "INV-2026-0917",
            "/invoice[1]/VAT[1]/@applies": "false",
            "/invoice[1]/invoicedate[1]": "2026-09-17T00:00:00Z",
            "/invoice[1]/lines[1]/lineitem[1]": [
                {
                    "productcode": "ACME-9",
                    "description": "Anvil, large",
                    "quantity": "2",
                    "price": "199.00",
                },
                {
                    "productcode": "ACME-3",
                    "description": "Rope, 30 m",
                    "quantity": "10",
                    "price": "9.00",
                },
                {
                    "productcode": "ACME-1",
                    "description": "Dynamite, one stick",
                    "quantity": "1",
                    "price": "49.00",
                },
            ],
        }
    )
    assert not result.skipped, result.skipped
    assert len(result.created) == 1, result.created

    # a new part, a new control, and a binding between them: docx4j's
    # addPropertiesPart, with the relationship from the main document part
    part = pkg.custom_xml_parts.add(
        '<approval xmlns="http://example.com/approval">'
        "<by>Grace Hopper</by><status>Approved</status></approval>"
    )
    body = pkg.body
    paragraph = body.insert_paragraph("Approved by: ")
    control = paragraph.get_range("End").insert_content_control("PlainText")
    mappings = "xmlns:ns0='http://example.com/approval'"
    assert control.xml_mapping.set_mapping("/ns0:approval[1]/ns0:by[1]", mappings, part)
    pkg.custom_xml_parts.apply_bindings()

    pkg.save(target)
    return target


def lists(out: Path) -> Path:
    """10. A loaded document given lists through the Phase H content API."""
    from docx4j_py.model.content import Author

    target = out / "10-lists.docx"
    pkg = WordprocessingMLPackage.load(SAMPLE_SOURCE)
    # a Word 2010 document; a loaded document's mode is never changed silently,
    # so say so here (CR-003 section 17.11)
    pkg.compatibility_mode = 15
    pkg.id_seed = 20260917
    pkg.author = Author("Claude", initials="C", email="claude@example.com")
    body = pkg.body

    # a numbered list of three, the first item starting it: the numbering part,
    # its relationship and its content type are all created here
    first = body.insert_paragraph("Check the invoice total")
    numbered = first.start_new_list()
    body.insert_paragraph("Check the licence key").attach_to_list(numbered.id)
    body.insert_paragraph("Ship it").attach_to_list(numbered.id)
    # and a fourth, a level deeper: Word paints "a." in front of it
    body.insert_paragraph("Ship it by Friday").attach_to_list(numbered.id, 1)

    # Word's "Restart numbering at 1": a second w:num over the *same*
    # w:abstractNum with a w:startOverride of 1 at level 0
    restarted = body.insert_paragraph("A second list, starting again at 1").start_new_list(
        like=numbered
    )
    body.insert_paragraph("Its second item").attach_to_list(restarted.id)

    # a bullet list, from docx4j's own default bullet definition
    bullets = body.insert_paragraph("A bulleted point").start_new_list(kind="Bullet")
    body.insert_paragraph("And another").attach_to_list(bullets.id)

    # one list restyled: a) b) c), the definition it shares with the restarted
    # list copied first so that the change stays local to this one
    numbered.set_level_numbering(0, "LowerLetter", "%1)")

    # a paragraph taken out of a list again
    body.insert_paragraph("This one was in the list and is not any more").attach_to_list(
        bullets.id
    )
    body.paragraph_at(contains="is not any more").detach_from_list()

    # and, with tracking on, a numbering change a human reviews in Word. The
    # paragraph is written *before* the mode goes on, so that it is the
    # document's own: a paragraph this author inserted records no w:pPrChange
    # (CR-003 section 16.2 item 6), and there would be no revision to review
    body.insert_paragraph("This paragraph joins the bullets, with tracking on")
    pkg.change_tracking_mode = "TrackAll"
    body.paragraph_at(contains="joins the bullets").attach_to_list(bullets.id)

    pkg.save(target)
    return target


def python_docx_facade(out: Path) -> Path:
    """11. A document built entirely through the python-docx facade (Phase J)."""
    target = out / "11-python-docx-facade.docx"
    with zipfile.ZipFile(IMAGE_SOURCE) as zf:
        png = zf.read(next(n for n in zf.namelist() if n.lower().endswith(".png")))

    # Everything below this line is ordinary python-docx code; the only thing
    # that is not is the import, and `pkg.id_seed` and the two fixed dates,
    # which keep the artefact byte-for-byte reproducible.
    from docx4j_py.docx import (
        WD_ALIGN_PARAGRAPH,
        WD_ORIENT,
        WD_SECTION,
        Document,
        Inches,
        Pt,
        RGBColor,
    )

    doc = Document()
    doc.package.id_seed = 20260918
    fixed = datetime.datetime(2026, 9, 18, 10, 0, tzinfo=datetime.UTC)
    doc.core_properties.created = fixed
    doc.core_properties.modified = fixed
    doc.core_properties.title = "The python-docx facade"
    doc.core_properties.author = "Ada Lovelace"

    doc.add_heading("The python-docx facade", level=0)
    doc.add_heading("What this is", level=1)

    # the section 3.6 snippet's constructs
    paragraph = doc.add_paragraph("Plain, ")
    paragraph.add_run("bold").bold = True
    paragraph.add_run(", ")
    emphasis = paragraph.add_run("red italic 14 pt")
    emphasis.italic = True
    emphasis.font.size = Pt(14)
    emphasis.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    paragraph.add_run(" and ")
    paragraph.add_run("not bold").bold = False
    paragraph.add_run(".")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_heading("Headings", level=2)
    doc.add_heading("A third-level heading", level=3)

    doc.add_heading("Lists, as python-docx makes them", level=2)
    doc.add_paragraph("A bulleted point, by style only", style="List Bullet")
    doc.add_paragraph("A numbered step, by style only", style="List Number")

    doc.add_heading("A table", level=2)
    headers = ("Region", "Quarter", "Revenue")
    table = doc.add_table(rows=1, cols=3, style="Table Grid")
    for index, name in enumerate(headers):
        cell = table.cell(0, index)
        cell.text = name
        cell.paragraphs[0].runs[0].bold = True
    for row in (("North", "Q1", "120"), ("South", "Q1", "98")):
        cells = table.add_row().cells
        for index, value in enumerate(row):
            cells[index].text = value
    # `_Cell.merge` is not implemented (CR-003 section 20), so there is no
    # merged cell here; a horizontally spanned cell would be the one thing this
    # artefact cannot show.

    doc.add_page_break()

    doc.add_heading("A picture", level=1)
    doc.add_paragraph("Two inches wide, sized by the facade:")
    doc.add_picture(png, width=Inches(2))

    # a second section, landscape, with a header and a footer of its own
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Inches(11.69), Inches(8.27)
    section.header.is_linked_to_previous = False
    section.header.paragraphs[0].text = "docx4j-python"
    section.footer.is_linked_to_previous = False
    section.footer.paragraphs[0].text = "Built through docx4j_py.docx"
    doc.add_paragraph("This last section is landscape, with its own header and footer.")

    doc.save(str(target))
    return target


#: Artefact 12's two fields and its equation, as check 57's file holds them (docx4j-core-ts ``906d5e3``).
_COMPLEX_FIELD = (
    '<w:r><w:t xml:space="preserve">14A </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> QUOTE "14F field 14G" </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">14F field 14G</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> 14B</w:t></w:r>'
)
_SIMPLE_FIELD = (
    '<w:r><w:t xml:space="preserve">15A </w:t></w:r><w:fldSimple w:instr=" QUOTE &quot;15F field 15G&quot; ">'
    '<w:r><w:t xml:space="preserve">15F field 15G</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> 15B</w:t></w:r>'
)
_EQUATION = (
    '<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>E=m</m:t></m:r>'
    '<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr>'
    '</m:ctrlPr></m:sSupPr><m:e><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
    '<m:t>c</m:t></m:r></m:e><m:sup><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
    '<m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>'
)
_NAMESPACES = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)

#: Artefact 12's cases: (label, what is written). The label says what the case's blocks become.
RANGE_CASES = [
    "Case 1: from the heading's middle (1B) to the body paragraph's (1C), deleted: one paragraph, '1A heading  body 1D', in the body paragraph's style.",
    "Case 2: from 2B to 2C, the table between taken whole, deleted: one paragraph, '2A before  after 2D'.",
    "Case 3: from the first list item's 3B to the second's 3C, deleted: two list items, '3A item  item 3D' and '3E item 3F'.",
    "Case 4: the next paragraph's 'Whole' range deleted: the paragraph goes, '4C stays' stays.",
    "Case 5: from 5B to 5C made bold: '5B' and '5C' bold, the first paragraph's mark bold.",
    "Case 6: insert_paragraph after the range from 6B to 6C: a new paragraph '6X new' after '6C', ' two 6D' after it.",
    "Case 7: a comment from 7B to 7C, across the two paragraphs.",
    "Case 8: a comment from 8B to 8C, across the table between.",
    "Case 9: a block content control wrapped around the two paragraphs '9A first' and '9B second'.",
    "Case 10: from 10B to 10D, the content control between taken whole, deleted: one paragraph, '10A before  after 10E'.",
    "Case 11: from 11B into the content control, to 11C, deleted: '11A before ' and, in the control, ' inside 11D'.",
    "Case 12: 'replace' replaced with 'REPLACED' within one paragraph (tracked, the insertion first).",
    "Case 13: a comment on 13B, then the range from 13B to 13C deleted: the comment goes with its reference.",
    "Case 14: a paragraph after 14F, in a complex field's result: the field split, 14F ends one paragraph, NEWPARA14, then 'field 14G' and the field's end.",
    "Case 15: the same in a w:fldSimple's result: the field made a complex field first, then split.",
    "Case 16: from 16B to 16C, a display equation's paragraph between, and an inline equation in 16C's paragraph before 16C, deleted.",
]


def ranges_source() -> WordprocessingMLPackage:
    """Artefact 12's document before its writes: ``samples/2010-sample1.docx`` with the sixteen cases appended."""
    pkg = WordprocessingMLPackage.load(SAMPLE_SOURCE)
    pkg.id_seed = 20261023
    body = pkg.body
    body.insert_paragraph("Ranges across paragraphs (CR-013 Phase C)", style="Heading 1")
    labels = iter(RANGE_CASES)

    def label() -> None:
        body.insert_paragraph(next(labels))

    def xml(inner: str) -> None:
        body.insert_xml(f"<w:p {_NAMESPACES}>{inner}</w:p>")

    def control(paragraph_xml: str, tag: str) -> None:
        body.insert_xml(
            f'<w:sdt {_NAMESPACES}><w:sdtPr><w:alias w:val="{tag}"/><w:tag w:val="{tag}"/></w:sdtPr>'
            f"<w:sdtContent>{paragraph_xml}</w:sdtContent></w:sdt>"
        )

    label()
    body.insert_paragraph("1A heading 1B", style="Heading 2")
    body.insert_paragraph("1C body 1D")
    label()
    body.insert_paragraph("2A before 2B")
    body.insert_table(2, 2, values=[["2T1", "2T2"], ["2T3", "2T4"]], style="TableGrid")
    body.insert_paragraph("2C after 2D")
    label()
    first = body.insert_paragraph("3A item 3B")
    numbered = first.start_new_list()
    body.insert_paragraph("3C item 3D").attach_to_list(numbered.id)
    body.insert_paragraph("3E item 3F").attach_to_list(numbered.id)
    label()
    body.insert_paragraph("4A whole paragraph 4B")
    body.insert_paragraph("4C stays")
    label()
    body.insert_paragraph("5A bold 5B")
    body.insert_paragraph("5C bold 5D")
    label()
    body.insert_paragraph("6A one 6B")
    body.insert_paragraph("6C two 6D")
    label()
    body.insert_paragraph("7A comment 7B")
    body.insert_paragraph("7C comment 7D")
    label()
    body.insert_paragraph("8A before 8B")
    body.insert_table(2, 2, values=[["8T1", "8T2"], ["8T3", "8T4"]], style="TableGrid")
    body.insert_paragraph("8C after 8D")
    label()
    body.insert_paragraph("9A first")
    body.insert_paragraph("9B second")
    label()
    body.insert_paragraph("10A before 10B")
    control('<w:p><w:r><w:t>10C inside</w:t></w:r></w:p>', "case10")
    body.insert_paragraph("10D after 10E")
    label()
    body.insert_paragraph("11A before 11B")
    control('<w:p><w:r><w:t>11C inside 11D</w:t></w:r></w:p>', "case11")
    label()
    body.insert_paragraph("12A replace 12B")
    label()
    body.insert_paragraph("13A has 13B a comment")
    body.insert_paragraph("13C next 13D")
    label()
    xml(_COMPLEX_FIELD)
    label()
    xml(_SIMPLE_FIELD)
    label()
    body.insert_paragraph("16A before 16B")
    xml(f"<m:oMathPara>{_EQUATION}</m:oMathPara>")
    xml(f'<w:r><w:t xml:space="preserve">16X inline </w:t></w:r>{_EQUATION}<w:r><w:t xml:space="preserve"> 16C after 16D</w:t></w:r>')
    body.insert_paragraph("End of the cases.")
    return pkg


def ranges_writes(pkg: WordprocessingMLPackage) -> None:
    """Artefact 12's writes, case by case: the same calls untracked and tracked."""

    def word(token: str):
        # the labels name the tokens too: a case's own hit is the one outside its label
        hits = [
            hit
            for hit in pkg.body.search(token, match_case=True, match_whole_word=True)
            if not hit.paragraph.text.startswith("Case ")
        ]
        assert len(hits) == 1, f"{token}: {len(hits)} hits"
        return hits[0]

    def across(start: str, end: str):
        return word(start).expand_to(word(end))

    def paragraph(text: str):
        return next(p for p in pkg.body.paragraphs if p.text == text)

    across("1B", "1C").delete()
    across("2B", "2C").delete()
    across("3B", "3C").delete()
    paragraph("4A whole paragraph 4B").get_range("Whole").delete()
    across("5B", "5C").font.bold = True
    across("6B", "6C").insert_paragraph("6X new", location="After")
    across("7B", "7C").insert_comment("A comment across two paragraphs (case 7).")
    across("8B", "8C").insert_comment("A comment across a table (case 8).")
    paragraph("9A first").get_range("Start").expand_to(paragraph("9B second").get_range("Whole")).insert_content_control()
    across("10B", "10D").delete()
    across("11B", "11C").delete()
    word("replace").insert_text("REPLACED", location="Replace")
    word("13B").insert_comment("A comment whose reference the deletion takes (case 13).")
    across("13B", "13C").delete()
    word("14F").insert_paragraph("NEWPARA14", location="After")
    word("15F").insert_paragraph("NEWPARA15", location="After")
    across("16B", "16C").delete()


def ranges_across_paragraphs(out: Path) -> list[Path]:
    """12. A loaded document given CR-013 Phase C's writes across paragraphs, untracked and tracked.

    Beside the tracked copy, the same resolved by CR-010's engine, as for 8:
    ``-tracked-accepted.docx`` and ``-tracked-rejected.docx``, each saved and
    read back listing nothing.
    """
    from docx4j_py.model import revisions
    from docx4j_py.model.content import Author

    written = []
    for name, mode in (("12-ranges-across-paragraphs.docx", "Off"), ("12-ranges-across-paragraphs-tracked.docx", "TrackAll")):
        pkg = ranges_source()
        pkg.author = Author("Claude", initials="C", email="claude@example.com")
        pkg.tracked_change_date = datetime.datetime(2026, 10, 23, 12, 0, tzinfo=datetime.UTC)
        pkg.change_tracking_mode = mode
        ranges_writes(pkg)
        pkg.change_tracking_mode = "Off"
        target = out / name
        pkg.save(target)
        written.append(target)
    tracked = written[-1]
    for suffix, resolve in (("accepted", revisions.accept_all), ("rejected", revisions.reject_all)):
        copy = WordprocessingMLPackage.load(tracked)
        resolve(copy)
        resolved = out / f"12-ranges-across-paragraphs-tracked-{suffix}.docx"
        copy.save(resolved)
        left = revisions.tracked_changes(WordprocessingMLPackage.load(resolved))
        assert left == [], f"{resolved.name} still lists {left}"
        written.append(resolved)
    return written


def main(argv: list[str] | None = None) -> int:
    """Write the twelve artefacts and print a summary."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", default=str(ROOT / "out" / "acceptance"))
    args = parser.parse_args(argv)

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    for build in (
        untouched,
        remarshalled,
        created,
        created_with_image,
        markdown_built,
        tables_and_pictures,
        comments,
        tracked_changes,
        template_filled,
        lists,
        python_docx_facade,
        ranges_across_paragraphs,
    ):
        built = build(out)
        for target in built if isinstance(built, list) else [built]:
            with zipfile.ZipFile(target) as zf:
                bad = zf.testzip()
            size = target.stat().st_size
            print(f"{target.relative_to(ROOT)}  {size:>8,} bytes  {'BAD' if bad else 'ok'}")

    print()
    print("Open each in Word. The checklist is tests/README.md.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
