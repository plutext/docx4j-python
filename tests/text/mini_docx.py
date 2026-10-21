"""docx4j's ``MiniDocx``: a ``.docx`` made from raw XML, for the text model's tests. CR-011 section 13.

Not a test. A body's XML, parts related from the main document part or the
package, and zip entries as they are, written into a zip and loaded through
this package's own load path (so that each ``mc:Choice`` records its
bindings as it is parsed, as a real document's does)::

    pkg = MiniDocx.body("<w:p><w:r><w:t>One</w:t></w:r></w:p>").load()
"""

from __future__ import annotations

import io
import zipfile

from docx4j_py import load

__all__ = ["NS", "MiniDocx"]

#: The prefixes a body's XML may use undeclared (docx4j ``MiniDocx.NS``).
NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:v="urn:schemas-microsoft-com:vml"'
    ' xmlns:o="urn:schemas-microsoft-com:office:office"'
    ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)

_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


class MiniDocx:
    """A package being made: its parts, relationships and content types."""

    def __init__(self, body: str) -> None:
        """A document whose body is this XML."""
        self.parts: dict[str, bytes] = {
            "word/document.xml": f"{_DECLARATION}<w:document {NS}><w:body>{body}</w:body></w:document>".encode()
        }
        self.types: list[str] = []
        self.rels: list[str] = []
        self.package_rels: list[str] = []

    @classmethod
    def body(cls, body: str) -> MiniDocx:
        """A document whose body is this XML (the prefixes of :data:`NS` are declared)."""
        return cls(body)

    def part(self, name: str, content_type: str, rel_type: str, rel_id: str, data: str | bytes) -> MiniDocx:
        """A part related from the main document part, under ``word/``."""
        self.parts[f"word/{name}"] = data.encode() if isinstance(data, str) else data
        self.types.append(f'<Override PartName="/word/{name}" ContentType="{content_type}"/>')
        self.rels.append(f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{name}"/>')
        return self

    def package_part(self, path: str, content_type: str, rel_type: str, rel_id: str, xml: str) -> MiniDocx:
        """A part related from the package (document properties)."""
        self.parts[path] = xml.encode()
        self.types.append(f'<Override PartName="/{path}" ContentType="{content_type}"/>')
        self.package_rels.append(f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{path}"/>')
        return self

    def entry(self, path: str, content_type: str | None, xml: str) -> MiniDocx:
        """A zip entry as it is, with its content type where it has one."""
        self.parts[path] = xml.encode()
        if content_type is not None:
            self.types.append(f'<Override PartName="/{path}" ContentType="{content_type}"/>')
        return self

    def relationship(self, rel_id: str, rel_type: str, target: str) -> MiniDocx:
        """A relationship from the main document part to a target given relative to it."""
        self.rels.append(f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>')
        return self

    def footnotes(self, notes: str) -> MiniDocx:
        """A footnotes part holding these notes."""
        return self.part(
            "footnotes.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
            "rIdFn",
            f"{_DECLARATION}<w:footnotes {NS}>{notes}</w:footnotes>",
        )

    def save(self) -> bytes:
        """The package's bytes."""
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(
                "[Content_Types].xml",
                f"{_DECLARATION}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/word/document.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                f"{''.join(self.types)}</Types>",
            )
            z.writestr(
                "_rels/.rels",
                f"{_DECLARATION}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                f'Target="word/document.xml"/>{"".join(self.package_rels)}</Relationships>',
            )
            z.writestr(
                "word/_rels/document.xml.rels",
                f"{_DECLARATION}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                f"{''.join(self.rels)}</Relationships>",
            )
            for name, data in self.parts.items():
                z.writestr(name, data)
        return out.getvalue()

    def load(self):  # noqa: ANN201 - a WordprocessingMLPackage
        """The package, loaded."""
        return load(self.save())
