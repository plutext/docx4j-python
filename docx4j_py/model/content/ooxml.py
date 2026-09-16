"""``insert_ooxml``: Word's flat OPC paste, and a bare fragment.

CR-003 section 3.2 and section 4, Phase C. Word's ``insertOoxml`` is given a
flat OPC ``pkg:package`` string --- what its own clipboard produces --- and so
is this one. The incoming package's body content is taken, every part that
content references through a relationship (an image, an embedded object, a
chart) is copied into the target package under a **free** name with a **fresh**
relationship id, the references in the content are rewritten, and the content is
inserted. A bare ``w:p`` / ``w:tbl`` fragment is accepted too, which is what
``insert_xml`` takes, and the two share this module.

The rules of CR-003 section 4, each of which cost the TypeScript engine a
correction:

* a copied part's **own** relationships are copied recursively and **keep their
  ids**, so nothing inside a copied part has to be rewritten;
* a reference is rewritten **by attribute name** and only when the incoming
  package really has a relationship of that id, so a numeric id
  (``w:bookmarkStart/@w:id``, ``wp:docPr/@id``) is never touched;
* the walk enters wildcard content (:func:`~docx4j_py.traversal.walk_all`),
  because a reference may sit in a DOM tree an ``xs:any`` holds --- a chart, a
  VML fallback;
* **styles and numbering are not merged.** That is docx4j's ``MergeDocx``, and
  a paste that silently redefined ``Heading 1`` would be worse than one that
  looks like the target document.

Reading the ``pkg:package`` is :class:`~docx4j_py.openpackaging.stores.FlatOpcStore`,
a read-only :class:`~docx4j_py.openpackaging.stores.PartStore` of eighty lines
that :func:`~docx4j_py.openpackaging.api.load` takes like any other; it is half
of CR-002 Phase C's flat OPC item, and writing one is still that phase's.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.model.content.enums import Location
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.picture import note_added_part
from docx4j_py.model.content.reports import current_recorder, recording
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.binary_part import BinaryPart, ImagePart
from docx4j_py.openpackaging.parts.relationships_part import (
    AddPartBehaviour,
    is_external,
)
from docx4j_py.openpackaging.stores import FlatOpcStore
from docx4j_py.traversal import walk_all
from docx4j_py.wml import wml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "REL_ATTRIBUTES",
    "content_of",
    "insert_ooxml_into_body",
    "insert_ooxml_into_paragraph",
    "insert_ooxml_into_range",
    "is_flat_opc",
    "rewrite_relationship_ids",
]

#: The field names that carry a relationship id. A value is rewritten only when
#: the incoming package really has a relationship of that id, so numeric ids are
#: never touched (CR-003 section 4).
REL_ATTRIBUTES: tuple[str, ...] = (
    "embed",
    "link",
    "id",
    "href",
    "pict",
    "dm",
    "lo",
    "qs",
    "cs",
    "top_left",
    "top_right",
    "bottom_left",
    "bottom_right",
)

#: The relationships namespace, for the same attributes on the lxml nodes a
#: wildcard field holds.
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: The file extension a copied image part is given, by content type.
_EXTENSION_BY_CONTENT_TYPE: dict[str, str] = {
    "image/png": "png",
    "image/x-png": "png",
    "image/jpeg": "jpeg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
    "image/svg+xml": "svg",
    "image/x-emf": "emf",
    "image/emf": "emf",
    "image/x-wmf": "wmf",
}


def is_flat_opc(ooxml: str) -> bool:
    """Whether a string is a flat OPC package rather than a bare fragment."""
    return FlatOpcStore.NAMESPACE in ooxml


# ---------------------------------------------------------------------------
# the content
# ---------------------------------------------------------------------------


def content_of(ooxml: str, *, target: Any = None, package: Any = None) -> list[Any]:
    """The block-level content of `ooxml`, ready to insert.

    A flat OPC package's body content with its parts copied and its references
    rewritten, or a fragment's elements.

    Args:
        ooxml: a ``pkg:package`` document, or a WordprocessingML fragment.
        target: the part new relationships are added to --- the main document
            part, or the header or footer the body belongs to. Required for a
            ``pkg:package``.
        package: the *caller's* package, which may be a
            :class:`~docx4j_py.model.content.trial.TrialPackage`; a part copied
            in is reported to it so that a dry run can un-add it.

    Raises:
        ContentError: a ``pkg:package`` with no target part, or with no main
            document part of its own.
    """
    if not is_flat_opc(ooxml):
        return list(wml.all(ooxml, wrapper="body"))
    if target is None:
        raise ContentError(
            "a pkg:package can only be inserted into a body that has a part",
            code="ooxml.no_part",
            hint="insert it into a body reached from a package, or pass a bare w:p fragment",
        )
    from docx4j_py.openpackaging.api import load
    from docx4j_py.openpackaging.exceptions import Docx4JException

    hint = (
        "Word's insertOoxml takes a pkg:package whose main part is a "
        "WordprocessingML document; a bare w:p / w:tbl fragment is accepted too"
    )
    try:
        source = load(FlatOpcStore.parse(ooxml))
        main = getattr(source, "main_document_part", None)
    except Docx4JException as error:
        raise ContentError(
            f"this pkg:package has no main document part: {error}",
            code="ooxml.no_main_part",
            hint=hint,
        ) from error
    if main is None:
        raise ContentError(
            "this pkg:package has no main document part, so there is nothing to insert",
            code="ooxml.no_main_part",
            hint=hint,
        )
    document = main.contents
    body = getattr(document, "body", None)
    content = list(getattr(body, "content", None) or ())
    if not content:
        return []
    rewrites = _copy_referenced_parts(content, main, target, package)
    if rewrites:
        rewrite_relationship_ids(content, rewrites)
    return content


def _copy_referenced_parts(
    content: list, source_part: Any, target: Any, package: Any = None
) -> dict[str, str]:
    """Copy every part the content references; returns old id -> new id."""
    rels = getattr(source_part, "relationships_part", None)
    out: dict[str, str] = {}
    if rels is None:
        return out
    referenced: list[str] = []
    _visit_references(content, lambda rel_id: _note(rels, rel_id, referenced))
    copied: dict[str, Any] = {}
    for rel_id in referenced:
        rel = rels.get_relationship_by_id(rel_id)
        if rel is None:
            continue
        owner_rels = target.get_relationships_part(True)
        if is_external(rel):
            new = owner_rels.add_external_relationship(rel.type_value, rel.target)
            out[rel_id] = new.id
            continue
        part = rels.get_part(rel)
        if part is None:
            continue
        new_id = _copy_part(part, target, copied, package)
        if new_id is not None:
            out[rel_id] = new_id
    return out


def _note(rels: Any, rel_id: str, out: list[str]) -> None:
    """Collect a reference the incoming package really has, once, in order."""
    if rels.get_relationship_by_id(rel_id) is not None and rel_id not in out:
        out.append(rel_id)


def _copy_part(
    source: Any, owner: Any, copied: dict[str, Any], notify: Any = None
) -> str | None:
    """Copy a part, and everything its own relationships target, under `owner`."""
    package = owner.package
    if package is None:  # pragma: no cover - the caller's part is in a package
        raise ContentError(
            "the target part is not in a package",
            code="ooxml.no_package",
            hint="load or create a package first",
        )
    rels = owner.get_relationships_part(True)
    already = copied.get(source.part_name.key)
    if already is not None:
        found = rels.get_rel(already.part_name)
        return found.id if found is not None else owner.add_target_part(already, "reuse").id
    copy = _copy_of(source, package)
    copied[source.part_name.key] = copy
    had_override = package.content_type_manager.get_override_content_type(copy.part_name)
    rel = owner.add_target_part(copy)
    note_added_part(
        notify if notify is not None else package,
        copy,
        rel,
        owner,
        added_content_type=had_override is None,
    )
    _record_part(notify if notify is not None else package, copy)
    _copy_child_relationships(source, copy, copied)
    return rel.id


def _record_part(package: Any, part: Any) -> None:
    """Name a copied part in the open report's ``parts_touched``."""
    parts = getattr(current_recorder(package), "parts", None)
    if parts is None:
        return
    name = str(part.part_name)
    if name not in parts:
        parts.append(name)


def _copy_child_relationships(source: Any, copy: Any, copied: dict[str, Any]) -> None:
    """Copy a part's own relationships, **keeping their ids** (CR-003 section 4)."""
    child_rels = getattr(source, "relationships_part", None)
    if child_rels is None or not len(child_rels):
        return
    copy_rels = copy.get_relationships_part(True)
    for child in list(child_rels):
        if is_external(child):
            copy_rels.add_external_relationship(child.type_value, child.target, child.id)
            continue
        child_part = child_rels.get_part(child)
        if child_part is None:
            continue
        already = copied.get(child_part.part_name.key)
        if already is not None:
            copy_rels.add_part(already, AddPartBehaviour.REUSE_EXISTING, child.id)
            continue
        nested = _copy_of(child_part, copy.package)
        copied[child_part.part_name.key] = nested
        copy_rels.add_part(nested, AddPartBehaviour.OVERWRITE_IF_NAME_EXISTS, child.id)
        _record_part(copy.package, nested)
        _copy_child_relationships(child_part, nested, copied)


def _copy_of(source: Any, package: Any) -> Any:
    """A binary copy of a part under a name free in the package."""
    name = _free_name(package, source)
    content_type = source.content_type or ""
    if content_type.startswith("image/"):
        copy: Any = ImagePart(name, content_type)
    else:
        copy = BinaryPart(name, content_type, source.relationship_type)
    copy.set_bytes(_bytes_of(source))
    return copy


def _bytes_of(part: Any) -> bytes:
    """A part's bytes, whatever kind of part it is."""
    data = getattr(part, "data", None)
    if isinstance(data, bytes):
        return data
    return bytes(part.bytes_for_save)


def _free_name(package: Any, source: Any) -> PartName:
    """``/word/media/imageN.<ext>`` for an image, else the source name with a suffix."""
    content_type = source.content_type or ""
    if content_type.startswith("image/"):
        extension = source.part_name.extension or _EXTENSION_BY_CONTENT_TYPE.get(
            content_type, "bin"
        )
        index = 1
        while True:
            candidate = PartName.of(f"/word/media/image{index}.{extension}")
            if package.get_part(candidate) is None:
                return candidate
            index += 1
    if package.get_part(source.part_name) is None:
        return source.part_name
    name = source.part_name.name
    dot = name.rfind(".")
    stem, suffix = (name[:dot], name[dot:]) if dot > name.rfind("/") else (name, "")
    index = 2
    while True:
        candidate = PartName.of(f"{stem}{index}{suffix}")
        if package.get_part(candidate) is None:
            return candidate
        index += 1


# ---------------------------------------------------------------------------
# the references
# ---------------------------------------------------------------------------


def rewrite_relationship_ids(value: Any, mapping: dict[str, str]) -> None:
    """Rewrite the relationship ids a tree refers to, in place."""
    _visit_references(value, lambda rel_id: mapping.get(rel_id))


def _visit_references(value: Any, rewrite: Any) -> None:
    """Offer every value that could be a relationship id to `rewrite`.

    Both halves of the tree are walked: the typed objects, which carry an id in
    one of :data:`REL_ATTRIBUTES`, and the lxml nodes a wildcard field holds,
    which carry one in an attribute of the relationships namespace.
    :func:`~docx4j_py.traversal.walk_all` is the walk that enters the second
    (CR-003 section 10.3 item 8: it defaults to ``mce="all"``, so a reference in
    the branch a consumer ignores is rewritten too --- it is written back on
    save).
    """

    def typed(node: Any, _parent: Any = None, _name: str | None = None) -> None:
        for field in REL_ATTRIBUTES:
            current = getattr(node, field, None)
            if not isinstance(current, str):
                continue
            replacement = rewrite(current)
            if replacement is not None:
                setattr(node, field, replacement)

    def wildcard(node: Any, _parent: Any = None, _name: str | None = None) -> None:
        attributes = getattr(node, "attributes", None)
        if not attributes:
            return
        for name, current in list(attributes.items()):
            if not isinstance(name, str) or not name.startswith(f"{{{_R_NS}}}"):
                continue
            if not isinstance(current, str):
                continue
            replacement = rewrite(current)
            if replacement is not None:
                attributes[name] = replacement

    for item in value if isinstance(value, list) else [value]:
        walk_all(item, typed, wildcard)


# ---------------------------------------------------------------------------
# the verbs
# ---------------------------------------------------------------------------


def insert_ooxml_into_body(
    body: Body,
    ooxml: str,
    *,
    location: Location = "End",
    target: Any = None,
) -> list[Any]:
    """What :meth:`Body.insert_ooxml` is. See its docstring."""
    with recording(body, "insert_ooxml") as change:
        part = getattr(body, "part", None)
        elements = content_of(
            ooxml, target=getattr(part, "_wrapped", part), package=body.package
        )
        if not elements:
            return []
        if location == "Replace":
            change.text(before=body.text)
            body.clear()
            body.insert_element(elements, location="End")
        else:
            body.insert_element(elements, location=location, target=target)
        views = [body.view_for(element) for element in elements]
        change.text(after="\n".join(view.text for view in views))
        return views


def insert_ooxml_into_paragraph(
    paragraph: Paragraph,
    ooxml: str,
    *,
    location: str = "After",
) -> list[Any]:
    """What :meth:`Paragraph.insert_ooxml` is. See its docstring."""
    from docx4j_py.wml import P

    body = paragraph.parent_body
    with recording(body, "insert_ooxml") as change:
        part = getattr(body, "part", None)
        elements = content_of(
            ooxml, target=getattr(part, "_wrapped", part), package=body.package
        )
        if not elements:
            return []
        only = elements[0] if len(elements) == 1 and isinstance(elements[0], P) else None
        if only is not None and location in ("Start", "End"):
            before = paragraph.text
            paragraph.insert_items_at(
                0 if location == "Start" else len(paragraph.text), list(only.content)
            )
            change.touched(paragraph)
            change.text(before=before, after=paragraph.text)
            return [paragraph]
        if location == "Replace":
            body.insert_element(elements, location="Before", target=paragraph)
            views = [body.view_for(element) for element in elements]
            paragraph.delete()
            return views
        where = "Before" if location == "Start" else "After" if location == "End" else location
        body.insert_element(elements, location=where, target=paragraph)
        return [body.view_for(element) for element in elements]


def insert_ooxml_into_range(
    span: Range,
    ooxml: str,
    *,
    location: str = "Replace",
) -> list[Any]:
    """What :meth:`Range.insert_ooxml` is. See its docstring."""
    from docx4j_py.wml import P

    paragraph = span.paragraph
    body = paragraph.parent_body
    with recording(body, "insert_ooxml") as change:
        part = getattr(body, "part", None)
        elements = content_of(
            ooxml, target=getattr(part, "_wrapped", part), package=body.package
        )
        if not elements:
            return []
        only = elements[0] if len(elements) == 1 and isinstance(elements[0], P) else None
        if only is not None:
            before = paragraph.text
            if location == "Replace":
                paragraph.splice(span.start, span.end, "")
                span.end = span.start
            offset = span.start if location in ("Before", "Replace") else span.end
            paragraph.insert_items_at(offset, list(only.content))
            change.touched(paragraph)
            change.text(before=before, after=paragraph.text)
            return [paragraph]
        where = "Before" if location in ("Before", "Replace") else "After"
        if location == "Replace":
            span.delete()
        body.insert_element(elements, location=where, target=paragraph)
        return [body.view_for(element) for element in elements]
