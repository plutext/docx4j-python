"""``pkg.describe()``: what an agent can use in this document.

CR-003 section 3.4. docx4j-mcp's ``describe_template`` widened to any document,
and the answer to "which style should I use for a heading here": the styles the
document defines and which of them it uses, the page setup, the parts present,
the authors of comments and revisions, whether tracking is on, and the
``skipped`` report.

**Nothing is unmarshalled.** CR-003 offered a choice --- read the styles part
through the model, which marks it for re-marshalling, or read the names out of
the XML with lxml, which keeps the part byte for byte --- and asked for the
second. So every part this reads is read as bytes (``part.xml``, which is the
source bytes for a part nobody has touched) and parsed with lxml, *except* the
main document part, whose tree is used when it is already unmarshalled, since a
caller with a ``body`` has unmarshalled it anyway and re-marshalling a
two-thousand-paragraph document to count its style references would be absurd.

``in_use`` is docx4j's ``stylesInUse``: the ``w:pStyle``, ``w:rStyle`` and
``w:tblStyle`` references in the main document part.
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.traversal import element_name, iter_nodes

__all__ = [
    "Description",
    "PageSetup",
    "PartInfo",
    "StyleInfo",
    "describe_package",
]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_DCTERMS = "http://purl.org/dc/terms/"
_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

#: Twips per point.
_TWIPS = 20.0

_STYLE_REFERENCES = frozenset(
    {f"{{{_W}}}pStyle", f"{{{_W}}}rStyle", f"{{{_W}}}tblStyle"}
)
_REVISION_ELEMENTS = frozenset(
    {
        f"{{{_W}}}ins",
        f"{{{_W}}}del",
        f"{{{_W}}}moveFrom",
        f"{{{_W}}}moveTo",
        f"{{{_W}}}rPrChange",
        f"{{{_W}}}pPrChange",
        f"{{{_W}}}tblPrChange",
        f"{{{_W}}}trPrChange",
        f"{{{_W}}}tcPrChange",
        f"{{{_W}}}sectPrChange",
    }
)


# ---------------------------------------------------------------------------
# the dataclasses
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class StyleInfo:
    """One style the document defines."""

    id: str
    """``w:styleId``: what goes in ``w:pStyle``."""
    name: str
    """``w:name/@w:val``, the display name Word shows."""
    kind: str
    """``paragraph``, ``character``, ``table`` or ``numbering``."""
    built_in: bool
    """False when ``w:customStyle="1"``: a style this document invented."""
    in_use: bool
    """True when the main document part references it (docx4j ``stylesInUse``)."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return dataclasses.asdict(self)


@dataclasses.dataclass(frozen=True, slots=True)
class PageSetup:
    """The page, from the last ``w:sectPr`` of the main document part."""

    width_pt: float = 0.0
    """The page width in points."""
    height_pt: float = 0.0
    """The page height in points."""
    orientation: str = "portrait"
    """``"portrait"`` or ``"landscape"``."""
    margins: dict[str, float] = dataclasses.field(default_factory=dict)
    """``top``, ``bottom``, ``left``, ``right``, ``header``, ``footer`` in points."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return dataclasses.asdict(self)


@dataclasses.dataclass(frozen=True, slots=True)
class PartInfo:
    """One part of the package."""

    name: str
    """``/word/document.xml``."""
    content_type: str
    """What ``[Content_Types].xml`` says it is."""
    kind: str
    """The part class: ``MainDocumentPart``, ``ImagePart``, ..."""
    unmarshalled: bool
    """True when something has read its contents, so it will be re-marshalled."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return dataclasses.asdict(self)


@dataclasses.dataclass(frozen=True, slots=True)
class Description:
    """What ``pkg.describe()`` returns."""

    styles: tuple[StyleInfo, ...] = ()
    """Every style the styles part defines, in its order."""
    page: PageSetup = dataclasses.field(default_factory=PageSetup)
    """The page setup."""
    parts: tuple[PartInfo, ...] = ()
    """Every part, by name."""
    headers: tuple[str, ...] = ()
    """The address prefixes of the header parts (``"header:rId8"``)."""
    footers: tuple[str, ...] = ()
    """The address prefixes of the footer parts."""
    custom_xml: tuple[tuple[str, str], ...] = ()
    """``(item id, namespace)`` per custom XML part: the template story."""
    authors: dict[str, list[str]] = dataclasses.field(default_factory=dict)
    """``{"comments": [...], "revisions": [...]}``, each sorted."""
    tracking_on: bool = False
    """``w:trackRevisions`` in the settings part."""
    compatibility_mode: int = 12
    """Which Word the document says it is for: 11, 12, 14 or 15 (CR-003 §3.4).

    12 --- Word 2007, what Word assumes --- when the document declares nothing.
    """
    skipped: int = 0
    """How many items lenient parsing dropped (CR-002 section 8)."""
    application: str | None = None
    """``docProps/app.xml``'s ``Application``."""
    created: str | None = None
    """``docProps/core.xml``'s ``dcterms:created``, as it is written."""
    modified: str | None = None
    """``dcterms:modified``."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a ``describe`` tool returns."""
        return {
            "styles": [style.to_dict() for style in self.styles],
            "page": self.page.to_dict(),
            "parts": [part.to_dict() for part in self.parts],
            "headers": list(self.headers),
            "footers": list(self.footers),
            "custom_xml": [list(pair) for pair in self.custom_xml],
            "authors": {key: list(value) for key, value in self.authors.items()},
            "tracking_on": self.tracking_on,
            "compatibility_mode": self.compatibility_mode,
            "skipped": self.skipped,
            "application": self.application,
            "created": self.created,
            "modified": self.modified,
        }

    def to_json(self, *, indent: int | None = None) -> str:
        """:meth:`to_dict` as JSON."""
        import json

        separators = (",", ":") if indent is None else None
        return json.dumps(self.to_dict(), indent=indent, separators=separators, ensure_ascii=False)

    def style_names(self, *, kind: str | None = None, in_use: bool | None = None) -> list[str]:
        """The display names, filtered. What "which style for a heading" asks. Extension."""
        return [
            style.name
            for style in self.styles
            if (kind is None or style.kind == kind) and (in_use is None or style.in_use == in_use)
        ]


# ---------------------------------------------------------------------------
# reading, without unmarshalling
# ---------------------------------------------------------------------------


def _root(part: Any) -> Any:
    """A part's root element as lxml, from the bytes it would be saved as.

    Reading this way leaves an untouched part untouched: no model class is
    built, nothing is marked for re-marshalling, and the bytes go out exactly
    as they came in.
    """
    if part is None:
        return None
    from lxml import etree

    try:
        return etree.fromstring(part.xml)
    except Exception:  # noqa: BLE001 - a part that will not parse describes nothing
        return None


def _styles(package: Any, in_use: set[str]) -> tuple[StyleInfo, ...]:
    root = _root(getattr(package, "style_definitions_part", None))
    if root is None:
        return ()
    from docx4j_py.model.content.styles import display_name_of

    out: list[StyleInfo] = []
    for style in root.findall(f"{{{_W}}}style"):
        style_id = style.get(f"{{{_W}}}styleId") or ""
        name_element = style.find(f"{{{_W}}}name")
        stored = name_element.get(f"{{{_W}}}val") if name_element is not None else None
        custom = style.get(f"{{{_W}}}customStyle")
        out.append(
            StyleInfo(
                id=style_id,
                # the display name, as ``paragraph.style`` reads and writes it:
                # Word stores ``heading 1`` and shows ``Heading 1`` (section 4)
                name=display_name_of(style_id, stored),
                kind=style.get(f"{{{_W}}}type") or "paragraph",
                built_in=custom not in ("1", "true"),
                in_use=style_id in in_use,
            )
        )
    return tuple(out)


def _styles_in_use(package: Any) -> set[str]:
    """docx4j ``stylesInUse``: every style the main document part references."""
    main = getattr(package, "main_document_part", None)
    if main is None:
        return set()
    out: set[str] = set()
    if getattr(main, "is_unmarshalled", False):
        for node in iter_nodes(main.contents, mce="all"):
            if element_name(node) in _STYLE_REFERENCES:
                value = getattr(node, "val", None)
                value = getattr(value, "value", value)
                if value:
                    out.add(str(value))
        return out
    root = _root(main)
    if root is None:
        return out
    for name in ("pStyle", "rStyle", "tblStyle"):
        for node in root.iter(f"{{{_W}}}{name}"):
            value = node.get(f"{{{_W}}}val")
            if value:
                out.add(value)
    return out


def _page(package: Any) -> PageSetup:
    main = getattr(package, "main_document_part", None)
    if main is None:
        return PageSetup()
    if getattr(main, "is_unmarshalled", False):
        body = getattr(main.contents, "body", None)
        sect_pr = getattr(body, "sect_pr", None) if body is not None else None
        if sect_pr is None:
            return PageSetup()
        size = getattr(sect_pr, "pg_sz", None)
        margin = getattr(sect_pr, "pg_mar", None)
        width = float(getattr(size, "w", 0) or 0)
        height = float(getattr(size, "h", 0) or 0)
        orient = getattr(size, "orient", None)
        orient = getattr(orient, "value", orient)
        margins = {
            key: float(getattr(margin, key, 0) or 0) / _TWIPS
            for key in ("top", "bottom", "left", "right", "header", "footer")
            if margin is not None
        }
        return PageSetup(
            width_pt=width / _TWIPS,
            height_pt=height / _TWIPS,
            orientation=str(orient) if orient else ("landscape" if width > height else "portrait"),
            margins=margins,
        )
    root = _root(main)
    if root is None:
        return PageSetup()
    sections = root.findall(f".//{{{_W}}}sectPr")
    if not sections:
        return PageSetup()
    sect_pr = sections[-1]
    size = sect_pr.find(f"{{{_W}}}pgSz")
    margin = sect_pr.find(f"{{{_W}}}pgMar")
    width = float(size.get(f"{{{_W}}}w") or 0) if size is not None else 0.0
    height = float(size.get(f"{{{_W}}}h") or 0) if size is not None else 0.0
    orient = size.get(f"{{{_W}}}orient") if size is not None else None
    margins = (
        {
            key: float(margin.get(f"{{{_W}}}{key}") or 0) / _TWIPS
            for key in ("top", "bottom", "left", "right", "header", "footer")
        }
        if margin is not None
        else {}
    )
    return PageSetup(
        width_pt=width / _TWIPS,
        height_pt=height / _TWIPS,
        orientation=orient or ("landscape" if width > height else "portrait"),
        margins=margins,
    )


def _parts(package: Any) -> tuple[PartInfo, ...]:
    out: list[PartInfo] = []
    for part in package.parts.parts():
        out.append(
            PartInfo(
                name=str(part.part_name),
                content_type=part.content_type,
                kind=type(part).__name__,
                unmarshalled=bool(getattr(part, "is_unmarshalled", False))
                or bool(getattr(part, "is_parsed", False)),
            )
        )
    out.sort(key=lambda info: info.name)
    return tuple(out)


def _comment_authors(package: Any) -> list[str]:
    main = getattr(package, "main_document_part", None)
    part = getattr(main, "comments_part", None) if main is not None else None
    root = _root(part)
    if root is None:
        return []
    found = {
        node.get(f"{{{_W}}}author")
        for node in root.findall(f"{{{_W}}}comment")
        if node.get(f"{{{_W}}}author")
    }
    return sorted(found)  # type: ignore[arg-type]


def _revision_authors(package: Any) -> list[str]:
    """``@w:author`` on every revision element, over every part that has a body.

    Through the tree for a part already unmarshalled and through lxml for one
    that is not, so that describing a document leaves its headers, footers and
    notes exactly as they came in.
    """
    from docx4j_py.model.content.addresses import package_parts

    found: set[str] = set()
    for part in package_parts(package):
        if getattr(part, "is_unmarshalled", False):
            for node in iter_nodes(part.contents, mce="all"):
                if element_name(node) in _REVISION_ELEMENTS:
                    author = getattr(node, "author", None)
                    if author:
                        found.add(str(author))
            continue
        root = _root(part)
        if root is None:
            continue
        for name in _REVISION_ELEMENTS:
            for node in root.iter(name):
                author = node.get(f"{{{_W}}}author")
                if author:
                    found.add(author)
    return sorted(found)


def _tracking_on(package: Any) -> bool:
    root = _root(getattr(package, "document_settings_part", None))
    if root is None:
        return False
    node = root.find(f"{{{_W}}}trackRevisions")
    if node is None:
        return False
    value = node.get(f"{{{_W}}}val")
    return value not in ("0", "false")


def _compatibility_mode(package: Any) -> int:
    """The Word version the document declares, read without unmarshalling."""
    from docx4j_py.model.content.compatibility import mode_of

    return mode_of(package)


def _custom_xml(package: Any) -> tuple[tuple[str, str], ...]:
    out: list[tuple[str, str]] = []
    for item_id, part in sorted(getattr(package, "custom_xml_data_storage_parts", {}).items()):
        namespace = ""
        if getattr(part, "is_parsed", False):
            tag = part.tree.tag
        else:
            root = _root(part)
            tag = root.tag if root is not None else ""
        if isinstance(tag, str) and tag.startswith("{"):
            namespace = tag[1:].partition("}")[0]
        out.append((str(item_id), namespace))
    return tuple(out)


def _properties(package: Any) -> tuple[str | None, str | None, str | None]:
    application = created = modified = None
    root = _root(getattr(package, "doc_props_extended_part", None))
    if root is not None:
        node = root.find(f"{{{_EXT_PROPS}}}Application")
        if node is None:
            node = root.find("Application")
        application = node.text if node is not None else None
    root = _root(getattr(package, "doc_props_core_part", None))
    if root is not None:
        for key, target in (("created", "created"), ("modified", "modified")):
            node = root.find(f"{{{_DCTERMS}}}{key}")
            if node is not None and node.text:
                if target == "created":
                    created = node.text
                else:
                    modified = node.text
    return application, created, modified


def describe_package(package: Any) -> Description:
    """``pkg.describe()``. See this module's docstring for what is read how."""
    from docx4j_py.model.content.addresses import package_parts, prefix_for_part

    in_use = _styles_in_use(package)
    headers: list[str] = []
    footers: list[str] = []
    for part in package_parts(package):
        # the prefix comes from the part and its relationship, not from a body,
        # so listing the headers does not unmarshal them
        prefix = prefix_for_part(part)
        if prefix.startswith("header"):
            headers.append(prefix)
        elif prefix.startswith("footer"):
            footers.append(prefix)
    application, created, modified = _properties(package)
    _ = _CP  # the core-properties namespace, for readers of this module
    return Description(
        styles=_styles(package, in_use),
        page=_page(package),
        parts=_parts(package),
        headers=tuple(headers),
        footers=tuple(footers),
        custom_xml=_custom_xml(package),
        authors={
            "comments": _comment_authors(package),
            "revisions": _revision_authors(package),
        },
        tracking_on=_tracking_on(package),
        compatibility_mode=_compatibility_mode(package),
        skipped=len(getattr(package, "skipped", ()) or ()),
        application=application,
        created=created,
        modified=modified,
    )
