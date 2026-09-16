"""Content controls: the ``w:sdt`` family, in Office JS's vocabulary.

CR-003 Phase A (section 3.3). A ``w:sdt`` is one element name with four types
--- block, run, row and cell --- and its properties are one choice list rather
than named fields, which is docx4j's model too (``SdtPr.getRPrOrAliasOrLock``).
So the builders here are:

:func:`sdt`
    the control itself, in whichever of the four forms the content needs.
:func:`sdt_pr`
    its ``w:sdtPr``: the alias, the tag, the id, the lock, the placeholder and
    the one element that says what kind of control it is.
:func:`sdt_property`
    a ``w:sdtPr`` child by element name, in ``w:`` *and* ``w15:``, because Word
    writes the binding of a repeating section as ``w15:dataBinding``.
:func:`sdt_kind_of`
    the Office JS ``Word.ContentControlType`` the kind element stands for.
:func:`next_sdt_id`
    a ``w:id`` free in a tree, allocated from the tree's own state so that the
    same document and the same calls give the same bytes.

    >>> from docx4j_py.wml import p, sdt, sdt_kind_of, to_xml
    >>> control = sdt([p("Acme")], kind="RichText", tag="customer", id=42)
    >>> sdt_kind_of(control.sdt_pr)
    'RichText'
"""

from __future__ import annotations

from typing import Any, Literal

from docx4j_py.child import ChildList, link_parents
from docx4j_py.namespaces import WML_NS
from docx4j_py.w14 import el as w14_el
from docx4j_py.w15 import el as w15_el
from docx4j_py.wml import (
    CtBdo,
    CTCustomXmlRun,
    CtDir,
    CTSdtCell,
    CTSdtContentCell,
    CTSdtContentRow,
    CTSdtContentRun,
    CTSdtRow,
    CTSimpleField,
    CTSmartTagRun,
    MoveFrom2,
    MoveTo2,
    PHyperlink,
    R,
    RunDel,
    RunIns,
    SdtBlock,
    SdtContentBlock,
    SdtPr,
    SdtRun,
    Tc,
    Tr,
    el,
)
from docx4j_py.wml.builders import BuilderError

__all__ = [
    "ANY_NS",
    "KIND_BY_QNAME",
    "SDT_FORMS",
    "SDT_KINDS",
    "W14_NS",
    "W15_NS",
    "W_AND_W15",
    "W_NS",
    "SdtForm",
    "SdtKind",
    "next_sdt_id",
    "sdt",
    "sdt_kind_of",
    "sdt_pr",
    "sdt_property",
]

#: WordprocessingML, ``docx4j_py.namespaces.WML_NS`` under the name
#: ``sdt_property``'s ``namespace=`` argument reads best with.
W_NS = WML_NS
#: Word 2010's namespace: ``w14:checkbox``, ``w14:entityPicker``.
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
#: Word 2012's namespace: ``w15:repeatingSection``, ``w15:dataBinding``,
#: ``w15:appearance``, ``w15:color``.
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
#: :func:`sdt_property`'s default: the two namespaces Word writes ``w:sdtPr``
#: children in.
W_AND_W15 = (W_NS, W15_NS)
#: Pass as `namespace` to look in any namespace at all.
ANY_NS = "*"

#: Office JS ``Word.ContentControlType``, as Word reports a control.
SdtKind = Literal[
    "RichText",
    "PlainText",
    "Picture",
    "BuildingBlockGallery",
    "CheckBox",
    "ComboBox",
    "DropDownList",
    "DatePicker",
    "RepeatingSection",
    "RepeatingSectionItem",
    "Group",
    "Citation",
    "Bibliography",
    "Equation",
]

#: Which of the four ``w:sdt`` forms: the model types each differently.
SdtForm = Literal["block", "run", "row", "cell"]

#: Every kind :func:`sdt_pr` can write, for a ``in`` check or an enumeration.
SDT_KINDS: tuple[str, ...] = (
    "RichText",
    "PlainText",
    "Picture",
    "BuildingBlockGallery",
    "CheckBox",
    "ComboBox",
    "DropDownList",
    "DatePicker",
    "RepeatingSection",
    "RepeatingSectionItem",
    "Group",
    "Citation",
    "Bibliography",
    "Equation",
)

#: The four forms, and the (``w:sdt`` class, ``w:sdtContent`` class) each takes.
SDT_FORMS: dict[str, tuple[type, type]] = {
    "block": (SdtBlock, SdtContentBlock),
    "run": (SdtRun, CTSdtContentRun),
    "row": (CTSdtRow, CTSdtContentRow),
    "cell": (CTSdtCell, CTSdtContentCell),
}

#: The font Word writes a checkbox's two glyphs in (docx4j's ``SdtWriter``).
CHECKBOX_FONT = "MS Gothic"

#: The kind element's qualified name, back to the kind Word reports.
#: ``w:docPartList`` is a building block gallery too, and ``w:richText`` is
#: rich text spelled out where Word usually says nothing at all.
KIND_BY_QNAME: dict[str, str] = {
    f"{{{W_NS}}}text": "PlainText",
    f"{{{W_NS}}}richText": "RichText",
    f"{{{W_NS}}}picture": "Picture",
    f"{{{W_NS}}}docPartObj": "BuildingBlockGallery",
    f"{{{W_NS}}}docPartList": "BuildingBlockGallery",
    f"{{{W_NS}}}comboBox": "ComboBox",
    f"{{{W_NS}}}dropDownList": "DropDownList",
    f"{{{W_NS}}}date": "DatePicker",
    f"{{{W_NS}}}group": "Group",
    f"{{{W_NS}}}citation": "Citation",
    f"{{{W_NS}}}bibliography": "Bibliography",
    f"{{{W_NS}}}equation": "Equation",
    f"{{{W14_NS}}}checkbox": "CheckBox",
    f"{{{W15_NS}}}repeatingSection": "RepeatingSection",
    f"{{{W15_NS}}}repeatingSectionItem": "RepeatingSectionItem",
}

#: What makes a ``w:sdt`` a *run*-level one: the classes ``EG_PContent`` and
#: ``EG_ContentRunContent`` hold. Everything else that is not a ``w:tr`` or a
#: ``w:tc`` is block level.
_RUN_LEVEL: tuple[type, ...] = (
    R,
    PHyperlink,
    SdtRun,
    CTSmartTagRun,
    CTCustomXmlRun,
    CTSimpleField,
    CtDir,
    CtBdo,
    RunIns,
    RunDel,
    MoveFrom2,
    MoveTo2,
)

#: The largest ``w:id`` (``ST_DecimalNumber`` is a signed 32-bit integer).
_MAX_SDT_ID = 0x7FFFFFFF


def _kind_element(kind: str) -> Any | None:
    """The ``w:sdtPr`` child that types a control; None for rich text.

    Rich text is the *untyped* control, as Word writes it (CR-003 section 4:
    ``type`` is ``"RichText"`` when ``w:sdtPr`` names no kind), so this returns
    None for it rather than a ``w:richText``.
    """
    match kind:
        case "RichText":
            return None
        case "PlainText":
            return el.text()
        case "Picture":
            return el.picture()
        case "BuildingBlockGallery":
            return el.docPartObj()
        case "ComboBox":
            return el.comboBox()
        case "DropDownList":
            return el.dropDownList()
        case "Group":
            return el.group()
        case "Citation":
            return el.citation()
        case "Bibliography":
            return el.bibliography()
        case "Equation":
            return el.equation()
        case "DatePicker":
            return el.date(
                date_format=el.dateFormat(val="d/MM/yyyy"),
                store_mapped_data_as=el.storeMappedDataAs(val="dateTime"),
            )
        case "CheckBox":
            return w14_el.checkbox(
                checked=w14_el.checked(val="0"),
                checked_state=w14_el.checkedState(val="2612", font=CHECKBOX_FONT),
                unchecked_state=w14_el.uncheckedState(val="2610", font=CHECKBOX_FONT),
            )
        case "RepeatingSection":
            return w15_el.repeatingSection()
        case "RepeatingSectionItem":
            return w15_el.repeatingSectionItem()
    raise BuilderError(
        f"not a content-control kind: {kind!r}",
        code="sdt.unknown_kind",
        hint=f"one of {', '.join(SDT_KINDS)}",
    )


def sdt_pr(
    *,
    kind: SdtKind = "RichText",
    tag: str | None = None,
    title: str | None = None,
    id: int | None = None,  # noqa: A002 - w:id's name
    alias: str | None = None,
    lock: Literal["sdtLocked", "contentLocked", "unlocked", "sdtContentLocked"] | None = None,
    placeholder: str | None = None,
) -> SdtPr:
    """A ``w:sdtPr``: alias, tag, id, lock, placeholder, then the kind element.

    Args:
        kind: the Office JS ``Word.ContentControlType``. ``"RichText"`` writes
            no kind element at all, which is the untyped control Word writes
            and reads back as rich text.
        tag: ``w:tag``, the machine-readable name.
        title: ``w:alias``, which Word's user interface calls the title.
        id: ``w:id``. Absent by default; :func:`next_sdt_id` gives one free in
            a tree, and Word assigns one when it opens a control without.
        alias: an accepted spelling of `title`, since the element is
            ``w:alias`` and the Word dialog says "Title"; give one or the other.
        lock: ``w:lock/@w:val``, what Word's "Content control cannot be
            deleted / edited" checkboxes set.
        placeholder: the *name of the glossary document part* holding the
            placeholder text (``w:placeholder/w:docPart/@w:val``), which is
            what the element holds; the text itself lives in that part.

    Raises:
        BuilderError: for an unknown kind, or for `title` and `alias` given
            different values.
    """
    if alias is not None:
        if title is not None and title != alias:
            raise BuilderError(
                f"title={title!r} and alias={alias!r} are the same element, w:alias",
                code="sdt.title_and_alias",
                hint="give one of them",
            )
        title = alias

    items: list[Any] = []
    if title is not None:
        items.append(el.alias(val=title))
    if tag is not None:
        items.append(el.tag(val=tag))
    if id is not None:
        items.append(el.id_(val=id))
    if lock is not None:
        items.append(el.lock(val=lock))
    if placeholder is not None:
        items.append(el.placeholder(doc_part=el.placeholder_doc_part(val=placeholder)))
    kind_element = _kind_element(kind)
    if kind_element is not None:
        items.append(kind_element)
    return SdtPr(content=ChildList(items))


def _form_of(content: list[Any]) -> str:
    """The form the content asks for: a ``w:tr`` a row, a ``w:tc`` a cell, …"""
    for item in content:
        if isinstance(item, Tr):
            return "row"
        if isinstance(item, Tc):
            return "cell"
        if isinstance(item, _RUN_LEVEL):
            return "run"
        return "block"
    return "block"


def sdt(
    content: list[Any],
    *,
    kind: SdtKind = "RichText",
    tag: str | None = None,
    title: str | None = None,
    id: int | None = None,  # noqa: A002 - w:id's name
    form: SdtForm | None = None,
    alias: str | None = None,
    lock: Literal["sdtLocked", "contentLocked", "unlocked", "sdtContentLocked"] | None = None,
    placeholder: str | None = None,
) -> SdtBlock | SdtRun | CTSdtRow | CTSdtCell:
    """A content control wrapping `content`.

    ``w:sdt`` is one element name with four types, and which one this is
    follows the content unless `form` says otherwise: a list of ``w:tr`` is the
    row form, of ``w:tc`` the cell form, of runs and run-level content the run
    form, and anything else --- paragraphs, tables, block-level controls --- the
    block form. The options are :func:`sdt_pr`'s and go straight to it.

    Args:
        content: the children of ``w:sdtContent``, in document order.
        form: override the inferred form.

    Raises:
        BuilderError: for an unknown form, or for a repeating section asked for
            at run level, which Word does not have.
    """
    resolved = form or _form_of(list(content))
    if resolved not in SDT_FORMS:
        raise BuilderError(
            f"not a content-control form: {resolved!r}",
            code="sdt.unknown_form",
            hint=f"one of {', '.join(SDT_FORMS)}",
        )
    if kind == "RepeatingSection" and resolved == "run":
        raise BuilderError(
            "a repeating section is a block-level control, not a run-level one",
            code="sdt.form_mismatch",
            hint="wrap paragraphs, a table or rows, or pass form='block'",
        )
    sdt_class, content_class = SDT_FORMS[resolved]
    control = sdt_class(
        sdt_pr=sdt_pr(
            kind=kind,
            tag=tag,
            title=title,
            id=id,
            alias=alias,
            lock=lock,
            placeholder=placeholder,
        ),
        sdt_content=content_class(content=ChildList(content)),
    )
    link_parents(control)
    return control


def _properties(pr: Any) -> Any:
    """Yield ``(qname, child)`` for a ``w:sdtPr``, or nothing for None.

    A ``w:sdt`` is accepted as well as its ``w:sdtPr``, because a caller that
    has the control should not have to reach through it.
    """
    from docx4j_py.child import iter_children

    if pr is None:
        return ()
    if not isinstance(pr, SdtPr):
        pr = getattr(pr, "sdt_pr", None)
        if pr is None:
            return ()
    return iter_children(pr)


def sdt_property(
    sdt_pr: SdtPr | None,
    local_name: str,
    namespace: str | tuple[str, ...] | None = W_AND_W15,
) -> Any | None:
    """A ``w:sdtPr`` child by element name, or None.

    The model keeps every ``w:sdtPr`` child in one choice list, as docx4j does,
    so this is the accessor: ``sdt_property(pr, "tag")``,
    ``sdt_property(pr, "dataBinding")``, ``sdt_property(pr, "checkbox")``.

    Args:
        sdt_pr: the properties, or the ``w:sdt`` itself, or None.
        local_name: the element's local name, without a prefix.
        namespace: where to look --- by default :data:`W_AND_W15`, **both**
            WordprocessingML and Word 2012's, because Word writes the binding
            of a repeating section, or of a container-bound rich-text control,
            as ``w15:dataBinding`` rather than ``w:dataBinding`` (the same type
            and the same attributes; 3 of the 20 bindings in
            ``samples/invoice2013.docx``). Pass :data:`W_NS` for wml alone,
            :data:`W14_NS` for ``w14:checkbox``, a tuple for several, or
            :data:`ANY_NS` for any namespace at all.

    Returns:
        The first matching child, or None.
    """
    wanted: frozenset[str] | None
    if namespace is None or namespace == ANY_NS:
        wanted = None
    else:
        wanted = frozenset([namespace] if isinstance(namespace, str) else namespace)

    for qname, child in _properties(sdt_pr):
        if not isinstance(qname, str):
            continue
        uri, _, local = qname.rpartition("}")
        if local != local_name:
            continue
        if wanted is None or uri.lstrip("{") in wanted:
            return child
    return None


def sdt_kind_of(sdt_pr: SdtPr | None) -> str:
    """The Office JS ``Word.ContentControlType`` a ``w:sdtPr`` describes.

    The kind element, mapped through :data:`KIND_BY_QNAME`; ``"RichText"``
    when there is none, which is how Word writes and reports an untyped
    control (CR-003 section 4). A ``w:sdt`` is accepted as well as its
    ``w:sdtPr``, and so is None.
    """
    for qname, _child in _properties(sdt_pr):
        kind = KIND_BY_QNAME.get(qname) if isinstance(qname, str) else None
        if kind is not None:
            return kind
    return "RichText"


def next_sdt_id(root: Any) -> int:
    """A ``w:id`` free among the content controls under `root`.

    One above the highest in use, as Word's own allocation reads, and
    deterministic: the same tree gives the same id, which is CR-003 section
    3.4's rule that the same document and the same calls make the same bytes.
    Where the highest is already ``ST_DecimalNumber``'s maximum, the lowest
    free id is returned instead.
    """
    from docx4j_py.traversal import find

    used: set[int] = set()
    for pr in find(root, SdtPr, mce="all"):
        value = getattr(sdt_property(pr, "id", W_NS), "val", None)
        if isinstance(value, int):
            used.add(value)

    candidate = max(used) + 1 if used else 1
    if candidate <= _MAX_SDT_ID:
        return candidate
    candidate = 1
    while candidate in used:
        candidate += 1
    return candidate
