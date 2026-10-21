"""Content one level down: an ``altChunk``'s (WD REQ-513) and an embedded package's (REQ-512). CR-011 section 6.

docx4j's ``org.docx4j.model.text.AltChunks`` and ``Embedded`` at
``198d015a3``, read into blocks apart from the package (REQ-204):

* **an ``altChunk``**, as Word expands it on open, by its declared content
  type, then its part's extension, then its bytes: a plain-text chunk is its
  lines (decoded by its byte-order mark, else as UTF-8 where it is valid,
  else windows-1252); a WordprocessingML or Flat OPC chunk a package read
  one level down; an HTML, XHTML or MHT chunk is read only by an HTML reader,
  which this package does not have (CR-011 Q12, Q13), so it is not expanded;
  an RTF chunk, a Word 2003 XML chunk and anything else are not expanded.
* **an embedded package**, under ``embedded: true``: a WordprocessingML
  document through the text model, one level deep. A workbook, a
  presentation and an OLE 1.0 package are deferred (CR-011 Q7): they
  contribute nothing, and raise ``content.not_read`` with the part's name,
  where docx4j reads them through xlsx4j, pptx4j and POI's OLE reader.

A package read one level down has the reading's options, but none of its own
embedded packages or chunks expanded (its chunks raising
``altchunk.not_expanded``), and no metadata. Its addresses start with the
chunk's or the object's (``body/7/altchunk:0/body/0``,
``body/3/object:0/body/0``), its warnings join the reading's, and its main
story's blocks take the story the content is read into (an ``altChunk``'s
the story the chunk is in, as Word expands it there; an embedded package's
``EMBEDDED``).
"""

from __future__ import annotations

import re
from typing import Any

__all__ = ["chunk_blocks", "content_type", "decode", "package_blocks"]

_XML_PACKAGE = b"http://schemas.microsoft.com/office/2006/xmlPackage"
_HTML_TYPES = ("text/html", "application/xhtml+xml", "message/rfc822")


def content_type(part: Any) -> str:
    """A chunk's content type as the package declares it, without parameters: the warning's selector."""
    value = getattr(part, "content_type", None) or ""
    return value.split(";", 1)[0].strip()


def decode(data: bytes | None) -> str:
    """A text file's bytes: by their byte-order mark, else UTF-8 if they are, else windows-1252."""
    if not data:
        return ""
    if data[:3] == b"\xef\xbb\xbf":
        return data[3:].decode("utf-8", errors="replace")
    if data[:2] == b"\xff\xfe":
        return data[2:].decode("utf-16-le", errors="replace")
    if data[:2] == b"\xfe\xff":
        return data[2:].decode("utf-16-be", errors="replace")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def _extension(part: Any) -> str:
    name = str(getattr(part, "part_name", "") or "")
    leaf = name.rsplit("/", 1)[-1]
    return leaf.rsplit(".", 1)[-1].lower() if "." in leaf else ""


def _lines(text: str, story: str, prefix: str) -> list[Any]:
    """A plain-text chunk's lines, a paragraph each; the file's last line break makes no line."""
    from docx4j_py.model.text.extractor import paragraph_block

    lines = re.split(r"\r\n|\r|\n", text)
    if lines and not lines[-1]:
        lines.pop()
    return [paragraph_block(story, f"{prefix}/{i}", line) for i, line in enumerate(lines)]


def chunk_blocks(x: Any, part: Any, story: str, prefix: str) -> list[Any] | None:
    """An ``altChunk``'s blocks; None where it is not expanded."""
    try:
        data = part.data
    except Exception:  # noqa: BLE001 - a chunk whose bytes will not read is not expanded
        return None
    if data is None:
        return None
    ct = content_type(part).lower()
    ext = _extension(part)
    if ct == "text/plain" or (not ct and ext == "txt"):
        return _lines(decode(data), story, prefix)
    if ct in ("application/rtf", "text/rtf") or ext == "rtf":
        return None
    if ct in _HTML_TYPES or (not ct and ext in ("htm", "html", "xhtml")) or ext in ("mht", "mhtml"):
        return None  # read only by an HTML reader, which is absent (CR-011 Q12, Q13)
    if (len(data) > 3 and data[:2] == b"PK") or _XML_PACKAGE in data[:2048]:
        return _document(x, data, story, prefix)
    return None


def _load(data: bytes) -> Any:
    from docx4j_py import load

    if data[:2] == b"PK":
        return load(data)
    from docx4j_py.openpackaging.stores import FlatOpcStore

    return load(FlatOpcStore.parse(data))


def _document(x: Any, data: bytes, story: str, prefix: str) -> list[Any] | None:
    """A WordprocessingML or Flat OPC chunk, read one level down; None for a package of another kind."""
    try:
        package = _load(data)
    except Exception:  # noqa: BLE001 - a chunk that will not load is not expanded
        return None
    if getattr(package, "main_document_part", None) is None:
        return None
    return x.read_nested(package, prefix, story)


def package_blocks(x: Any, part: Any, prefix: str) -> list[Any]:
    """An embedded package's blocks (an ``EmbeddedPackagePart``'s or an OLE 1.0 package's bytes).

    Raises:
        Exception: where the package cannot be read, or is one of the kinds
            CR-011 Q7 defers: then it contributes nothing (``content.not_read``).
    """
    if type(part).__name__ != "EmbeddedPackagePart":
        raise NotImplementedError("an OLE 1.0 package (CR-011 Q7)")
    package = _load(part.data)
    if getattr(package, "main_document_part", None) is None:
        raise NotImplementedError(f"a {type(package).__name__} (CR-011 Q7)")
    return x.read_nested(package, prefix, "EMBEDDED")
