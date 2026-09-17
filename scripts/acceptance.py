#!/usr/bin/env python
"""Write the nine documents the Word acceptance checklist needs.

CR-002 section 8: Word acceptance is manual. This produces the artefacts and
prints what to look for; `tests/README.md` is the checklist.

    .venv-fork/bin/python scripts/acceptance.py [--out out/acceptance]

Nine files, each testing a different half of the save path:

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
    document's own ``w:sectPr``.
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
    mode on: a replacement, an inserted paragraph, a deleted paragraph, a bold
    run, a table row added and one deleted, and a comment on the replacement.
``9-template-filled.docx``
    ``samples/invoice2013.docx`` --- a Word-authored template with twenty
    bindings over ``/customXml/item1.xml`` --- **filled** through CR-003 Phase
    E's ``pkg.custom_xml_parts.fill()``: a new company and invoice number, the
    VAT checkbox cleared, a new date, and new data for the first line item of a
    repeating section. The repeat itself is **not** expanded: a repeating
    section is a container and is never bound (section 4), so the document shows
    the one item it already had, with the new values in it. A new content
    control is inserted and bound to a **new** custom XML part added through
    ``add()``, which is the test of ``addPropertiesPart``'s relationship from
    the main document part.
"""

from __future__ import annotations

import argparse
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

    body.insert_paragraph("Tables and pictures", style="Heading 1")
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


def tracked_changes(out: Path) -> Path:
    """8. A loaded document edited through the Phase F content API, tracking on."""
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
    # the last block of the body is a paragraph, so that the paragraph appended
    # below has one before it to carry its mark (CR-003 section 16.10)
    body.insert_paragraph("This paragraph will be deleted, with its mark.")
    body.insert_paragraph("This paragraph will be made bold.")

    pkg.change_tracking_mode = "TrackAll"

    # a replacement, explained in a comment on what it changed
    body.replace_text("first", "second")
    hit = pkg.find("second")[0]
    hit.range(body).insert_comment("Changed 'first' to 'second': the source says 2010 was the second.")

    # an insertion where it can be seen --- after the first paragraph, not at the
    # very end --- and one appended, which exercises the final-mark rule of
    # CR-003 section 16.10 (the appended paragraph's own mark is left alone and
    # the mark before it is marked instead, as Word does)
    body.paragraphs[0].insert_paragraph(
        "Added by an agent, right after the first paragraph.", location="After"
    )
    body.insert_paragraph("And one appended at the very end.")
    body.paragraph_at(contains="will be deleted").delete()
    body.paragraph_at(contains="made bold").font.bold = True
    table.add_rows(1, values=[["North", "Q2", "300"]])
    table.delete_rows(2)

    pkg.save(target)
    return target


def template_filled(out: Path) -> Path:
    """9. A Word-authored template filled through the Phase E custom XML API."""
    target = out / "9-template-filled.docx"
    pkg = WordprocessingMLPackage.load(TEMPLATE_SOURCE)
    pkg.id_seed = 20260917

    # what the template wants, before anything is written: the same call
    # docx4j-mcp's describe_template makes
    skeleton = pkg.custom_xml_parts.describe()
    assert len(skeleton) == 20, len(skeleton)

    # and the fill: a text binding, the checkbox, the date, and the data of the
    # repeating section's first line item. The repeat is not expanded --- a
    # container is never bound --- so the one item it has shows the new values.
    result = pkg.custom_xml_parts.fill(
        {
            "/invoice[1]/customer[1]/company[1]": "Acme Manufacturing Ltd",
            "/invoice[1]/customer[1]/contact[1]": "Ada Lovelace",
            "/invoice[1]/invoicenumber[1]": "INV-2026-0917",
            "/invoice[1]/VAT[1]/@applies": "false",
            "/invoice[1]/invoicedate[1]": "2026-09-17T00:00:00Z",
            "/invoice[1]/lines[1]/lineitem[1]/productcode[1]": "ACME-9",
            "/invoice[1]/lines[1]/lineitem[1]/description[1]": "Anvil, large",
            "/invoice[1]/lines[1]/lineitem[1]/quantity[1]": "2",
            "/invoice[1]/lines[1]/lineitem[1]/price[1]": "199.00",
        }
    )
    assert not result.skipped, result.skipped

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


def main(argv: list[str] | None = None) -> int:
    """Write the nine artefacts and print a summary."""
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
    ):
        target = build(out)
        with zipfile.ZipFile(target) as zf:
            bad = zf.testzip()
        size = target.stat().st_size
        print(f"{target.relative_to(ROOT)}  {size:>8,} bytes  {'BAD' if bad else 'ok'}")

    print()
    print("Open each in Word. The checklist is tests/README.md.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
