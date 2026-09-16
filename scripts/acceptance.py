#!/usr/bin/env python
"""Write the four documents the Word acceptance checklist needs.

CR-002 section 8: Word acceptance is manual. This produces the artefacts and
prints what to look for; `tests/README.md` is the checklist.

    .venv-fork/bin/python scripts/acceptance.py [--out out/acceptance]

Four files, each testing a different half of the save path:

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
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.openpackaging import (  # noqa: E402
    ImagePart,
    WordprocessingMLPackage,
)
from docx4j_py.wml import el, emu_for, image_size, inline_picture, p, r  # noqa: E402

SOURCE = ROOT / "samples" / "2016_image_with_text_effects.docx"
IMAGE_SOURCE = ROOT / "samples" / "Images.docx"

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


def main(argv: list[str] | None = None) -> int:
    """Write the four artefacts and print a summary."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", default=str(ROOT / "out" / "acceptance"))
    args = parser.parse_args(argv)

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    for build in (untouched, remarshalled, created, created_with_image):
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
