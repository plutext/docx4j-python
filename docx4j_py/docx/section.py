"""``Section``, ``Sections``, ``_Header`` and ``_Footer``.

CR-003 section 3.6, Phase J. A ``Section`` is a view over one ``w:sectPr``: the
last one is ``w:body/w:sectPr`` and every earlier one is a
``w:p/w:pPr/w:sectPr``, which is python-docx's ``sectPr_lst`` and the order
Word reads them in. The content API has no section view at all, so everything
here is written straight into the typed tree inside a
:class:`~docx4j_py.model.content.reports.recording` block.

``_Header`` and ``_Footer`` are the part-creating half. Setting
``is_linked_to_previous = False`` **creates** ``/word/headerN.xml`` with its
relationship from the main document part and its content-type override, and
puts a ``w:headerReference`` in this section's ``w:sectPr``; setting it back to
True removes the reference and drops the relationship, as python-docx does.
A dry run un-does both, because the part goes through
:func:`~docx4j_py.model.content.picture.note_added_part` like every other part
CR-003 creates (section 14.4).
"""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from typing import TYPE_CHECKING, Any, overload

from docx4j_py.docx._base import BlockItemContainer, Story, recording
from docx4j_py.docx.enums import (
    WD_HEADER_FOOTER,
    WD_ORIENTATION,
    WD_SECTION_START,
)
from docx4j_py.docx.shared import Length, Twips, twips_of
from docx4j_py.model.values import is_on
from docx4j_py.wml import CtSectPrType, P, SectPr, el

if TYPE_CHECKING:  # pragma: no cover - annotations only
    from docx4j_py.docx.document import Document

__all__ = ["Section", "Sections", "_BaseHeaderFooter", "_Footer", "_Header"]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def sect_pr_list(body: Any) -> list[SectPr]:
    """Every ``w:sectPr`` a document's body reaches, in document order.

    python-docx's ``./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr``: the
    paragraph-carried ones first, the sentinel last.
    """
    out: list[SectPr] = []
    for item in getattr(body, "content", None) or ():
        if isinstance(item, P):
            found = getattr(getattr(item, "p_pr", None), "sect_pr", None)
            if found is not None:
                out.append(found)
    if getattr(body, "sect_pr", None) is not None:
        out.append(body.sect_pr)
    return out


class Section:
    """One ``w:sectPr``: python-docx's ``Section``."""

    __slots__ = ("_document", "_sect_pr", "_story")

    def __init__(self, sect_pr: SectPr, story: Story, document: Document | None = None) -> None:
        """Wrap the ``w:sectPr``."""
        self._sect_pr = sect_pr
        self._story = story
        self._document = document

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:sectPr`` are equal."""
        return isinstance(other, Section) and other._sect_pr is self._sect_pr

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._sect_pr))

    def __repr__(self) -> str:
        """``<docx4j_py.docx Section NEW_PAGE PORTRAIT>``."""
        return (
            f"<docx4j_py.docx Section {self.start_type.name} {self.orientation.name}>"
        )

    # -- what is underneath ------------------------------------------------

    @property
    def element(self) -> SectPr:
        """The typed ``w:sectPr``."""
        return self._sect_pr

    @property
    def _sectPr(self) -> SectPr:
        """python-docx's private name for :attr:`element`."""
        return self._sect_pr

    @property
    def part(self) -> Any:
        """The main document part."""
        return self._story.part

    # -- the start type and the page ---------------------------------------

    @property
    def start_type(self) -> WD_SECTION_START:
        """``w:type/@val``; ``NEW_PAGE`` when there is none, as Word assumes."""
        value = getattr(getattr(self._sect_pr, "type_value", None), "val", None)
        value = getattr(value, "value", value)
        if value is None:
            return WD_SECTION_START.NEW_PAGE
        try:
            return WD_SECTION_START.from_xml(str(value))
        except ValueError:
            return WD_SECTION_START.NEW_PAGE

    @start_type.setter
    def start_type(self, value: WD_SECTION_START | None) -> None:
        with self._recording("docx.section.start_type"):
            if value is None or WD_SECTION_START(value) is WD_SECTION_START.NEW_PAGE:
                self._sect_pr.type_value = None
                return
            self._sect_pr.type_value = CtSectPrType(val=WD_SECTION_START(value).xml_value)
            self._sect_pr.type_value.parent = self._sect_pr

    def _pg_sz(self, *, create: bool) -> Any:
        pg_sz = getattr(self._sect_pr, "pg_sz", None)
        if pg_sz is None and create:
            pg_sz = el.pgSz()
            self._sect_pr.pg_sz = pg_sz
            pg_sz.parent = self._sect_pr
        return pg_sz

    def _pg_mar(self, *, create: bool) -> Any:
        pg_mar = getattr(self._sect_pr, "pg_mar", None)
        if pg_mar is None and create:
            pg_mar = el.pgMar()
            self._sect_pr.pg_mar = pg_mar
            pg_mar.parent = self._sect_pr
        return pg_mar

    @property
    def orientation(self) -> WD_ORIENTATION:
        """``w:pgSz/@orient``; ``PORTRAIT`` when absent."""
        value = getattr(self._pg_sz(create=False), "orient", None)
        value = getattr(value, "value", value)
        if value is None:
            return WD_ORIENTATION.PORTRAIT
        try:
            return WD_ORIENTATION.from_xml(str(value))
        except ValueError:
            return WD_ORIENTATION.PORTRAIT

    @orientation.setter
    def orientation(self, value: WD_ORIENTATION | None) -> None:
        with self._recording("docx.section.orientation"):
            member = WD_ORIENTATION.PORTRAIT if not value else WD_ORIENTATION(value)
            self._pg_sz(create=True).orient = member.xml_value

    @property
    def page_width(self) -> Length | None:
        """``w:pgSz/@w``, in EMU."""
        return _twips(getattr(self._pg_sz(create=False), "w", None))

    @page_width.setter
    def page_width(self, value: int | Length | None) -> None:
        with self._recording("docx.section.page_width"):
            self._pg_sz(create=True).w = twips_of(value)

    @property
    def page_height(self) -> Length | None:
        """``w:pgSz/@h``, in EMU."""
        return _twips(getattr(self._pg_sz(create=False), "h", None))

    @page_height.setter
    def page_height(self, value: int | Length | None) -> None:
        with self._recording("docx.section.page_height"):
            self._pg_sz(create=True).h = twips_of(value)

    def _margin(self, attribute: str) -> Length | None:
        return _twips(getattr(self._pg_mar(create=False), attribute, None))

    def _set_margin(self, attribute: str, value: int | Length | None) -> None:
        with self._recording(f"docx.section.{attribute}"):
            setattr(self._pg_mar(create=True), attribute, twips_of(value))

    @property
    def left_margin(self) -> Length | None:
        """``w:pgMar/@left``."""
        return self._margin("left")

    @left_margin.setter
    def left_margin(self, value: int | Length | None) -> None:
        self._set_margin("left", value)

    @property
    def right_margin(self) -> Length | None:
        """``w:pgMar/@right``."""
        return self._margin("right")

    @right_margin.setter
    def right_margin(self, value: int | Length | None) -> None:
        self._set_margin("right", value)

    @property
    def top_margin(self) -> Length | None:
        """``w:pgMar/@top``."""
        return self._margin("top")

    @top_margin.setter
    def top_margin(self, value: int | Length | None) -> None:
        self._set_margin("top", value)

    @property
    def bottom_margin(self) -> Length | None:
        """``w:pgMar/@bottom``."""
        return self._margin("bottom")

    @bottom_margin.setter
    def bottom_margin(self, value: int | Length | None) -> None:
        self._set_margin("bottom", value)

    @property
    def header_distance(self) -> Length | None:
        """``w:pgMar/@header``: the page edge to the top of the header."""
        return self._margin("header")

    @header_distance.setter
    def header_distance(self, value: int | Length | None) -> None:
        self._set_margin("header", value)

    @property
    def footer_distance(self) -> Length | None:
        """``w:pgMar/@footer``: the page edge to the bottom of the footer."""
        return self._margin("footer")

    @footer_distance.setter
    def footer_distance(self, value: int | Length | None) -> None:
        self._set_margin("footer", value)

    @property
    def gutter(self) -> Length | None:
        """``w:pgMar/@gutter``: the binding allowance."""
        return self._margin("gutter")

    @gutter.setter
    def gutter(self, value: int | Length | None) -> None:
        self._set_margin("gutter", value)

    # -- headers and footers ----------------------------------------------

    @property
    def different_first_page_header_footer(self) -> bool:
        """``w:titlePg``: this section has its own first-page header and footer."""
        return is_on(getattr(self._sect_pr, "title_pg", None))

    @different_first_page_header_footer.setter
    def different_first_page_header_footer(self, value: bool) -> None:
        with self._recording("docx.section.different_first_page_header_footer"):
            if not value:
                self._sect_pr.title_pg = None
                return
            self._sect_pr.title_pg = el.titlePg()
            self._sect_pr.title_pg.parent = self._sect_pr

    @property
    def header(self) -> _Header:
        """The section's default header (odd pages, or all of them)."""
        return _Header(self, WD_HEADER_FOOTER.PRIMARY)

    @property
    def footer(self) -> _Footer:
        """The section's default footer."""
        return _Footer(self, WD_HEADER_FOOTER.PRIMARY)

    @property
    def first_page_header(self) -> _Header:
        """The header for the first page of the section."""
        return _Header(self, WD_HEADER_FOOTER.FIRST_PAGE)

    @property
    def first_page_footer(self) -> _Footer:
        """The footer for the first page of the section."""
        return _Footer(self, WD_HEADER_FOOTER.FIRST_PAGE)

    @property
    def even_page_header(self) -> _Header:
        """The header for even pages of the section."""
        return _Header(self, WD_HEADER_FOOTER.EVEN_PAGE)

    @property
    def even_page_footer(self) -> _Footer:
        """The footer for even pages of the section."""
        return _Footer(self, WD_HEADER_FOOTER.EVEN_PAGE)

    # -- the content of the section ----------------------------------------

    def iter_inner_content(self) -> Iterator[Any]:
        """Each ``Paragraph`` or ``Table`` in this section, in document order.

        The blocks between the previous section's break and this one's. The
        paragraph that *carries* a ``w:sectPr`` belongs to its own section and
        is the last block of it, which is python-docx's rule.
        """
        from docx4j_py.docx.table import Table
        from docx4j_py.docx.text import Paragraph
        from docx4j_py.wml import Tbl

        body = self._story.body
        blocks = list(body.content)
        starts = [
            index
            for index, item in enumerate(blocks)
            if isinstance(item, P) and getattr(getattr(item, "p_pr", None), "sect_pr", None)
        ]
        own = next(
            (
                index
                for index in starts
                if blocks[index].p_pr.sect_pr is self._sect_pr
            ),
            None,
        )
        if own is None:
            begin = starts[-1] + 1 if starts else 0
            end = len(blocks)
        else:
            earlier = [index for index in starts if index < own]
            begin = earlier[-1] + 1 if earlier else 0
            end = own + 1
        for item in blocks[begin:end]:
            if isinstance(item, P):
                yield Paragraph(item, self._story)
            elif isinstance(item, Tbl):
                yield Table(item, self._story)

    # -- internals ---------------------------------------------------------

    @property
    def _preceding(self) -> SectPr | None:
        """The ``w:sectPr`` before this one, or None when this is the first."""
        found = sect_pr_list(self._story.body.container)
        for index, sect_pr in enumerate(found):
            if sect_pr is self._sect_pr:
                return found[index - 1] if index else None
        return None

    def _recording(self, operation: str) -> Any:
        return recording(self._story.body, operation)


def _twips(value: Any) -> Length | None:
    """A twips attribute as a :class:`Length`."""
    return None if value is None else Twips(int(value))


class Sections(Sequence):
    """``document.sections``: the document's sections, in order."""

    __slots__ = ("_document", "_story")

    def __init__(self, story: Story, document: Document | None = None) -> None:
        """Wrap the main document part's body."""
        self._story = story
        self._document = document

    def __repr__(self) -> str:
        """``<docx4j_py.docx Sections, 2 sections>``."""
        return f"<docx4j_py.docx Sections, {len(self)} sections>"

    @property
    def _list(self) -> list[SectPr]:
        return sect_pr_list(self._story.body.container)

    def __len__(self) -> int:
        """How many sections the document has."""
        return len(self._list)

    def __iter__(self) -> Iterator[Section]:
        """Each section, in document order."""
        return (Section(sect_pr, self._story, self._document) for sect_pr in self._list)

    @overload
    def __getitem__(self, key: int) -> Section: ...

    @overload
    def __getitem__(self, key: slice) -> list[Section]: ...

    def __getitem__(self, key: int | slice) -> Section | list[Section]:
        """Indexed and sliced access."""
        found = self._list
        if isinstance(key, slice):
            return [Section(sect_pr, self._story, self._document) for sect_pr in found[key]]
        return Section(found[key], self._story, self._document)


# ---------------------------------------------------------------------------
# headers and footers
# ---------------------------------------------------------------------------


class _BaseHeaderFooter(BlockItemContainer):
    """What a header and a footer share."""

    __slots__ = ("_index", "_section")

    #: ``"headerReference"`` or ``"footerReference"``.
    reference_name = ""
    #: The relationship type of the part.
    relationship_type = ""
    #: ``"header"`` or ``"footer"``, for the part name.
    kind = ""

    def __init__(self, section: Section, index: WD_HEADER_FOOTER) -> None:
        """Wrap one of a section's three header (or footer) definitions."""
        self._section = section
        self._index = index

    def __repr__(self) -> str:
        """``<docx4j_py.docx _Header PRIMARY, linked>``."""
        state = "linked" if self.is_linked_to_previous else "defined"
        return f"<docx4j_py.docx {type(self).__name__} {self._index.name}, {state}>"

    # -- the block container ----------------------------------------------

    @property
    def _story(self) -> Story:
        """The header's own part, as a story of its own."""
        from docx4j_py.model.content.body import body_of

        part = self._get_or_add_definition()
        return Story(body_of(part), getattr(self._section, "_document", None))

    @property
    def _container(self) -> Any:
        """The ``w:hdr`` or ``w:ftr``."""
        return self._story.body.container

    @property
    def part(self) -> Any:
        """The ``HeaderPart`` or ``FooterPart`` this definition lives in."""
        return self._get_or_add_definition()

    @property
    def element(self) -> Any:
        """The typed ``w:hdr`` or ``w:ftr``."""
        return self._container

    # -- linkage -----------------------------------------------------------

    @property
    def is_linked_to_previous(self) -> bool:
        """True when this section has no definition of its own."""
        return not self._has_definition

    @is_linked_to_previous.setter
    def is_linked_to_previous(self, value: bool) -> None:
        wanted = bool(value)
        if wanted == self.is_linked_to_previous:
            return
        if wanted:
            self._drop_definition()
        else:
            self._add_definition()

    @property
    def _has_definition(self) -> bool:
        return self._reference() is not None

    def _reference(self, sect_pr: SectPr | None = None) -> Any | None:
        """This section's ``w:headerReference`` of this index, or None."""
        from docx4j_py.traversal import element_name

        target = sect_pr if sect_pr is not None else self._section.element
        wanted = f"{_W}{self.reference_name}"
        for item in getattr(target, "header_reference_or_footer_reference", None) or ():
            if element_name(item) != wanted:
                continue
            value = getattr(item, "type_value", None)
            value = getattr(value, "value", value)
            kind = "default" if value is None else str(value)
            if kind == WD_HEADER_FOOTER(self._index).xml_value:
                return item
        return None

    def _add_definition(self) -> Any:
        """Create the part, relate it and reference it from this ``w:sectPr``."""
        from docx4j_py.model.content.picture import note_added_part
        from docx4j_py.model.content.reports import recording_on

        section = self._section
        main = section._story.part
        package = main.package
        with recording_on(package, f"docx.{self.kind}.create", main) as change:
            part = self._new_part(self._free_name(package))
            had_override = package.content_type_manager.get_override_content_type(part.part_name)
            rel = main.add_target_part(part)
            note_added_part(package, part, rel, main, added_content_type=had_override is None)
            reference = getattr(el, self.reference_name)(
                type_value=WD_HEADER_FOOTER(self._index).xml_value, id=rel.id
            )
            items = section.element.header_reference_or_footer_reference
            items.append(reference)
            change.part(part)
            return part

    def _drop_definition(self) -> None:
        """Remove the reference and the relationship, as python-docx does."""
        from docx4j_py.model.content.reports import recording_on

        section = self._section
        main = section._story.part
        package = main.package
        reference = self._reference()
        if reference is None:
            return
        with recording_on(package, f"docx.{self.kind}.unlink", main):
            items = section.element.header_reference_or_footer_reference
            for index, item in enumerate(list(items)):
                if item is reference:
                    del items[index]
                    break
            rels = main.relationships_part
            if rels is not None and getattr(reference, "id", None):
                rel = rels.get_relationship_by_id(reference.id)
                if rel is not None:
                    rels.remove_relationship(rel)

    def _definition(self) -> Any:
        """The part this section's reference points at."""
        reference = self._reference()
        main = self._section._story.part
        rels = getattr(main, "relationships_part", None)
        if reference is None or rels is None:
            raise ValueError(f"this section has no {self.kind} of its own")
        part = rels.get_part(reference.id)
        if part is None:
            raise ValueError(f"the {self.kind} relationship {reference.id} targets no part")
        return part

    def _get_or_add_definition(self) -> Any:
        """The part holding the content shown for this section.

        Inherited from the previous section when this one has no definition,
        recursively; the first section's is created when it has none, which is
        python-docx's rule and what makes ``section.header.paragraphs[0]`` work
        on a document with no header at all.
        """
        if self._has_definition:
            return self._definition()
        preceding = self._section._preceding
        if preceding is not None:
            prior = type(self)(
                Section(preceding, self._section._story, self._section._document), self._index
            )
            return prior._get_or_add_definition()
        return self._add_definition()

    def _free_name(self, package: Any) -> str:
        """``/word/header2.xml``, the first N free."""
        index = 1
        while True:
            candidate = f"/word/{self.kind}{index}.xml"
            if package.get_part(candidate) is None:
                return candidate
            index += 1

    def _new_part(self, part_name: str) -> Any:
        raise NotImplementedError


class _Header(_BaseHeaderFooter):
    """A ``w:headerReference`` and the ``/word/headerN.xml`` behind it."""

    __slots__ = ()
    reference_name = "headerReference"
    kind = "header"

    def _new_part(self, part_name: str) -> Any:
        """A new ``HeaderPart`` with the one empty paragraph Word requires."""
        from docx4j_py.openpackaging.parts.wml import HeaderPart
        from docx4j_py.wml import Hdr

        part = HeaderPart(part_name)
        part.set_contents(Hdr(content=[el.p()]))
        return part


class _Footer(_BaseHeaderFooter):
    """A ``w:footerReference`` and the ``/word/footerN.xml`` behind it."""

    __slots__ = ()
    reference_name = "footerReference"
    kind = "footer"

    def _new_part(self, part_name: str) -> Any:
        """A new ``FooterPart`` with the one empty paragraph Word requires."""
        from docx4j_py.openpackaging.parts.wml import FooterPart
        from docx4j_py.wml import Ftr

        part = FooterPart(part_name)
        part.set_contents(Ftr(content=[el.p()]))
        return part
