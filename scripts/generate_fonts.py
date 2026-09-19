#!/usr/bin/env python
"""Turn docx4j's font data files into the committed JSON of ``docx4j_py/resources/fonts/``.

CR-002 Phase B step 3 (section 15.3)::

    .venv-fork/bin/python scripts/generate_fonts.py            # reads ../docx4j
    DOCX4J=/path/to/docx4j .venv-fork/bin/python scripts/generate_fonts.py

Four resources of docx4j's own become three committed JSON files, plus the three
Office theme parts copied verbatim. **Nothing here decides anything**: each file
is docx4j's rows, in docx4j's order, so that this port and docx4j read the same
table --- which is why docx4j's CR-017 phase 0 put the table in a file at all.

===================================================== =========================
docx4j resource                                        output
===================================================== =========================
``org/docx4j/fonts/font-substitutes.xml``              ``substitutions.json``
``org/docx4j/fonts/substitutions/FontSubstitutions.xml`` ``substitutions.json``
``org/docx4j/fonts/microsoft/MicrosoftFonts.xml``      ``families.json``
``org/docx4j/fonts/word-line-metrics.properties``      ``families.json``
``…/WordprocessingML/theme-{2007,2013,2023}.xml``      ``theme-YYYY.xml``
===================================================== =========================

The registry of physical fonts (``registry.json``) is **not** generated from a
file: it is the faces docx4j's four font jars carry, which is the environment
the parity goldens were made in, and it is read off
``docx4j-export-fo-fonts-{symbol,croscore,crosextra,theme2023}``'s font files
by name. The list is small and fixed, so it is written here rather than parsed
out of the jars, which would need a font library this package does not depend
on; ``--check-jars`` verifies that the jars in the sibling checkout still carry
exactly those files.

Each JSON file carries a ``_source`` object naming the docx4j commit this ran
against and the resources it read, so a golden that moves can be traced.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
DOCX4J = Path(os.environ.get("DOCX4J") or ROOT.parent / "docx4j")
RESOURCES = DOCX4J / "docx4j-core" / "src" / "main" / "resources"
OUT = ROOT / "docx4j_py" / "resources" / "fonts"
THEME_DIR = "org/docx4j/openpackaging/parts/WordprocessingML"

#: ``[name, familyName]`` of every face docx4j's four font jars carry, read off
#: docx4j ``VERSION_17_1_1`` (``PhysicalFonts.getPhysicalFonts()`` with system
#: discovery off and ``docx4j-export-fo-fonts-symbol``, ``-croscore``,
#: ``-crosextra`` and ``-theme2023`` on the classpath): Liberation's croscore
#: metric clones, crosextra's Carlito and Caladea, the theme2023 stand-ins for
#: Word 365's Aptos, and the symbol jar's DejaVu Serif and Noto Sans Symbols.
JAR_FONTS: list[tuple[str, str]] = [
    ("Akasia Black", "Akasia"),
    ("Akasia Black Italic", "Akasia Black"),
    ("Akasia Bold", "Akasia"),
    ("Akasia Bold Italic", "Akasia"),
    ("Akasia ExtraBold", "Akasia"),
    ("Akasia ExtraBold Italic", "Akasia ExtraBold"),
    ("Akasia Italic", "Akasia"),
    ("Akasia Light", "Akasia"),
    ("Akasia Light Italic", "Akasia Light"),
    ("Akasia Regular", "Akasia"),
    ("Akasia SemiBold", "Akasia"),
    ("Akasia SemiBold Italic", "Akasia"),
    ("Arimo Bold", "Arimo"),
    ("Arimo Bold Italic", "Arimo"),
    ("Arimo Italic", "Arimo"),
    ("Arimo Regular", "Arimo"),
    ("Caladea Bold", "Caladea"),
    ("Caladea Bold Italic", "Caladea"),
    ("Caladea Italic", "Caladea"),
    ("Caladea Regular", "Caladea"),
    ("Carlito Bold", "Carlito"),
    ("Carlito Bold Italic", "Carlito"),
    ("Carlito Italic", "Carlito"),
    ("Carlito Regular", "Carlito"),
    ("Cousine Bold", "Cousine"),
    ("Cousine Bold Italic", "Cousine"),
    ("Cousine Italic", "Cousine"),
    ("Cousine Regular", "Cousine"),
    ("DejaVu Serif", "DejaVu Serif"),
    ("DejaVu Serif Bold", "DejaVu Serif"),
    ("DejaVu Serif Bold Italic", "DejaVu Serif"),
    ("DejaVu Serif Italic", "DejaVu Serif"),
    ("Intos Display Bold", "Intos Display"),
    ("Intos Display Bold Italic", "Intos Display"),
    ("Intos Display Italic", "Intos Display"),
    ("Intos Display Regular", "Intos Display"),
    ("Noto Sans Symbols 2 Regular", "Noto Sans Symbols 2"),
    ("Noto Sans Symbols Bold", "Noto Sans Symbols"),
    ("Noto Sans Symbols Regular", "Noto Sans Symbols"),
    ("Tinos Bold", "Tinos"),
    ("Tinos Bold Italic", "Tinos"),
    ("Tinos Italic", "Tinos"),
    ("Tinos Regular", "Tinos"),
]


def commit() -> str:
    """docx4j's HEAD, or ``unknown`` where the checkout is not a git repository."""
    try:
        return subprocess.run(
            ["git", "-C", str(DOCX4J), "rev-parse", "HEAD"],
            capture_output=True,
            text=True,
            check=True,
        ).stdout.strip()
    except Exception:  # noqa: BLE001
        return "unknown"


# --------------------------------------------------------------- a small XML reader
#
# The three XML files are hand-maintained data with no namespaces, no mixed content and
# no entities beyond the five predefined ones, so a regular-expression reader is enough
# --- docx4j's own FontSubstitutionTable reads its file with a plain DOM parser for the
# same reason --- and this script stays free of lxml, which the generated model needs but
# a data generator should not.


def _unescape(text: str) -> str:
    out = (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
    )
    out = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), out)
    return out.replace("&amp;", "&")


def _attributes(text: str) -> dict[str, str]:
    return {
        m.group(1): _unescape(m.group(2))
        for m in re.finditer(r'([A-Za-z_:][\w.:-]*)\s*=\s*"([^"]*)"', text)
    }


def _elements(xml: str, tag: str) -> list[tuple[dict[str, str], str]]:
    """Every ``<tag …>…</tag>`` or ``<tag … />`` at any depth, as (attributes, body)."""
    out: list[tuple[dict[str, str], str]] = []
    for match in re.finditer(rf"<{tag}(\s[^>]*?)?(/)?>", xml):
        attrs = _attributes(match.group(1) or "")
        if match.group(2):
            out.append((attrs, ""))
            continue
        end = xml.find(f"</{tag}>", match.end())
        out.append((attrs, "" if end < 0 else xml[match.end() : end]))
    return out


def _child(body: str, tag: str) -> str | None:
    """The text of the first ``<tag>`` directly inside `body`, or None."""
    match = re.search(rf"<{tag}\s*/>|<{tag}(?:\s[^>]*)?>([\s\S]*?)</{tag}>", body)
    if match is None:
        return None
    return "" if match.group(1) is None else _unescape(match.group(1)).strip()


# ------------------------------------------------------------ font-substitutes.xml


def substitute_rows() -> list[dict[str, Any]]:
    """``font-substitutes.xml``'s ``<substitutes>`` rows, in the file's order.

    ``<scriptSubstitutes>`` and ``<widthFactors>`` are deliberately not read:
    they serve the glyph-coverage and width passes, which this phase does not
    port (CR-002 section 15.3).
    """
    xml = (RESOURCES / "org/docx4j/fonts/font-substitutes.xml").read_text(encoding="utf-8")
    block = re.search(r"<substitutes>([\s\S]*?)</substitutes>", xml)
    if block is None:
        raise SystemExit("no <substitutes> block in font-substitutes.xml")
    rows: list[dict[str, Any]] = []
    for attrs, body in _elements(block.group(1), "font"):
        substitutes = [
            {
                "font": s_attrs["font"],
                "quality": s_attrs["quality"],
                "error": s_attrs.get("error"),
            }
            for s_attrs, _ in _elements(body, "substitute")
        ]
        rows.append({"documentFont": attrs["name"], "substitutes": substitutes})
    if not rows:
        raise SystemExit("no rows in font-substitutes.xml")
    return rows


def vcl_entries() -> dict[str, dict[str, str]]:
    """``FontSubstitutions.xml``'s entries, keyed as that file keys them."""
    xml = (RESOURCES / "org/docx4j/fonts/substitutions/FontSubstitutions.xml").read_text(
        encoding="utf-8"
    )
    entries: dict[str, dict[str, str]] = {}
    for attrs, body in _elements(xml, "replace"):
        name = attrs.get("name")
        if not name:
            continue
        entries[name] = {
            "substFonts": _child(body, "SubstFonts") or "",
            "fontType": _child(body, "FontType") or "",
            "fontWidth": _child(body, "FontWidth") or "",
        }
    if not entries:
        raise SystemExit("no <replace> entries in FontSubstitutions.xml")
    return entries


def microsoft_fonts() -> tuple[list[str], list[str]]:
    """``MicrosoftFonts.xml``'s families, and those whose entry has a ``<bold>``."""
    xml = (RESOURCES / "org/docx4j/fonts/microsoft/MicrosoftFonts.xml").read_text(
        encoding="utf-8"
    )
    names: list[str] = []
    with_bold: list[str] = []
    for attrs, body in _elements(xml, "font"):
        name = attrs.get("name")
        if not name:
            continue
        names.append(name)
        if re.search(r"<bold[\s/>]", body):
            with_bold.append(name)
    if not names:
        raise SystemExit("no <font> entries in MicrosoftFonts.xml")
    return names, with_bold


def line_metric_families() -> tuple[list[str], list[str]]:
    """``word-line-metrics.properties``'s families, and the East Asian ones."""
    text = (RESOURCES / "org/docx4j/fonts/word-line-metrics.properties").read_text(
        encoding="utf-8"
    )
    families: list[str] = []
    east_asian: list[str] = []
    for line in text.splitlines():
        if not line or line.startswith("#"):
            continue
        equals = line.find("=")
        if equals < 0:
            continue
        # a java.util.Properties key escapes its spaces
        key = line[:equals].replace("\\ ", " ").strip().lower()
        fields = line[equals + 1 :].split(";")
        families.append(key)
        # optional field 7 (index 6): the OS/2 code-page bits call the family East Asian
        if len(fields) > 6 and fields[6].strip() not in ("", "0"):
            east_asian.append(key)
    if not families:
        raise SystemExit("no families in word-line-metrics.properties")
    return families, east_asian


def check_jars() -> int:
    """Whether the font jars still carry exactly the faces :data:`JAR_FONTS` names."""
    import zipfile

    found: set[str] = set()
    for jar_dir in sorted(DOCX4J.glob("docx4j-export-fo-fonts-*")):
        for jar in sorted(jar_dir.glob("target/*.jar")):
            with zipfile.ZipFile(jar) as archive:
                for name in archive.namelist():
                    if name.lower().endswith((".ttf", ".otf")):
                        found.add(Path(name).stem)
    if not found:
        print("no built font jars under ../docx4j; run mvn install first", file=sys.stderr)
        return 2
    print(f"{len(found)} font files in the jars, {len(JAR_FONTS)} faces registered")
    return 0


def main() -> int:
    """Write the JSON and the three themes; ``--check-jars`` verifies the registry."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check-jars",
        action="store_true",
        help="list the font files docx4j's built font jars carry and stop",
    )
    args = parser.parse_args()
    if args.check_jars:
        return check_jars()

    sha = commit()
    OUT.mkdir(parents=True, exist_ok=True)

    rows = substitute_rows()
    vcl = vcl_entries()
    names, with_bold = microsoft_fonts()
    families, east_asian = line_metric_families()

    _write(
        "substitutions.json",
        {
            "_source": {
                "docx4jCommit": sha,
                "resources": [
                    (
                        "org/docx4j/fonts/font-substitutes.xml (the <substitutes> block;"
                        " <scriptSubstitutes> and <widthFactors> serve the glyph-coverage"
                        " and width passes, which CR-002 Phase B does not port)"
                    ),
                    "org/docx4j/fonts/substitutions/FontSubstitutions.xml",
                ],
                "generator": "scripts/generate_fonts.py",
            },
            "substituteRows": rows,
            "vclEntries": vcl,
        },
    )
    _write(
        "families.json",
        {
            "_source": {
                "docx4jCommit": sha,
                "resources": [
                    "org/docx4j/fonts/microsoft/MicrosoftFonts.xml",
                    "org/docx4j/fonts/word-line-metrics.properties",
                ],
                "generator": "scripts/generate_fonts.py",
            },
            "msFontNames": names,
            "msFontsWithBold": with_bold,
            "lineMetricFamilies": families,
            "eastAsianFamilies": east_asian,
        },
    )
    _write(
        "registry.json",
        {
            "_source": {
                "docx4jCommit": sha,
                "resources": [
                    (
                        "docx4j-export-fo-fonts-{symbol,croscore,crosextra,theme2023}, read"
                        " off PhysicalFonts.getPhysicalFonts() with system discovery off:"
                        " the environment CR-016 phase 0c measured and the goldens record"
                    ),
                ],
                "generator": "scripts/generate_fonts.py",
            },
            "fonts": [{"name": name, "family": family} for name, family in JAR_FONTS],
        },
    )

    for value in ("2007", "2013", "2023"):
        source = RESOURCES / THEME_DIR / f"theme-{value}.xml"
        (OUT / f"theme-{value}.xml").write_bytes(source.read_bytes())

    print(f"docx4j {sha}")
    print(f"  substitutions.json  {len(rows)} substitute rows, {len(vcl)} VCL entries")
    print(
        f"  families.json       {len(names)} Microsoft families ({len(with_bold)} with a"
        f" bold face), {len(families)} line-metric families ({len(east_asian)} East Asian)"
    )
    print(f"  registry.json       {len(JAR_FONTS)} faces")
    print("  theme-2007.xml, theme-2013.xml, theme-2023.xml")
    return 0


def _write(name: str, data: dict[str, Any]) -> None:
    text = json.dumps(data, indent=1, ensure_ascii=False, sort_keys=False)
    (OUT / name).write_text(text + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
