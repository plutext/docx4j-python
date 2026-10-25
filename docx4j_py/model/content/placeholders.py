"""Placeholders in Word's form: the glossary entries a content control names (CR-019 Phase B).

docx4j-core-ts's ``src/model/content/placeholders.mts`` at ``aefc41f`` (its
CR-002 sections 63.9 to 63.12 and 63.18; D8 and D9). Word keeps a
placeholder's text in the **glossary document part**, as a building block the
control names in ``w:placeholder/w:docPart/@w:val`` (core-ts check 64); the
control's runs show a copy of it while ``w:showingPlcHdr`` is on. A control
with no ``w:placeholder`` is in the **bare form**, what docx4j's binding
writes: Word writes into it as into Word's form (check 63's run 2), but has no
entry to show once it is emptied, nor for ``placeholder_text`` to read, and
fills an emptied one with five spaces (CR-019 section 11, D10). A control
naming any entry, present in the glossary or not, is in Word's form while the
file is open, though Word's save drops a reference whose entry no glossary
holds (check 65).

:class:`GlossaryEntries` is ``package.glossary_entries``:

* **reading**: the glossary document part's entries, read privately
  (``read_contents``, CR-002 section 16), so that an untouched glossary is
  saved byte for byte; read again only where the part is another, and from the
  live tree once the part is unmarshalled. The content API reads them when it
  first reads a body (``body_of``), as core-ts's ``getBody`` does, and here,
  where reads are synchronous, on first use too;
* **writing**: an entry a placeholder names that the glossary lacks waits for
  the save, which runs :meth:`GlossaryEntries.flush` through the package's
  ``save_hooks`` (the engine imports nothing from here). The save makes the
  glossary document part (with its relationship and content type) where the
  file has none, adds each entry in Word's form (check 65), and makes the
  glossary's styles part (``/word/glossary/styles.xml``), or completes one,
  with Default Paragraph Font and Placeholder Text as Word defines them (D9
  (a), check 66). A styles part holding both stays untouched;
* a trial (``dry_run``) has entries of its own, the real package's pending ones
  copied in, and thrown away with it.

What a control shows of an entry (section 63.18, as docx4j's CR-063 D2 has
it): a copy of its paragraphs' run content, the control's own
``w:sdtPr/w:rPr`` combined with each run's properties
(:func:`entry_run_properties`, the one place the combination is made).
"""

from __future__ import annotations

import dataclasses
import uuid
from typing import TYPE_CHECKING, Any
from xml.sax.saxutils import escape

from docx4j_py.child import ChildList, deep_copy, link_parents
from docx4j_py.traversal import element_name, text_of
from docx4j_py.wml.sdt import W_NS, sdt_property

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.controls import ContentControl

__all__ = [
    "DEFAULT_PLACEHOLDER_ENTRY",
    "DEFAULT_PLACEHOLDER_TEXT",
    "GLOSSARY_DOCUMENT_NAME",
    "GLOSSARY_STYLES_NAME",
    "PLACEHOLDER_STYLE",
    "GlossaryEntries",
    "GlossaryEntry",
    "default_paragraph_font",
    "doc_part_of",
    "ensure_glossary_styles",
    "entries_of",
    "entry_content_of",
    "entry_run_properties",
    "glossary_entries_of",
    "has_placeholder_styles",
    "load_glossary_entries",
    "name_placeholder_entry",
    "new_entry_name",
    "placeholder_entry_of",
    "placeholder_style",
    "require_placeholder_style",
    "shows_placeholder_when_made",
]

#: Word's default placeholder entry, which a control Word makes names (core-ts checks 64 to 66).
DEFAULT_PLACEHOLDER_ENTRY = "DefaultPlaceholder_-1854013440"
#: The default entry's text, as Word 365 writes it.
DEFAULT_PLACEHOLDER_TEXT = "Click or tap here to enter text."
#: The character style a placeholder's runs name.
PLACEHOLDER_STYLE = "PlaceholderText"
#: Where a save makes the glossary document part, and its styles part.
GLOSSARY_DOCUMENT_NAME = "/word/glossary/document.xml"
GLOSSARY_STYLES_NAME = "/word/glossary/styles.xml"

#: The paragraph items an entry's copy leaves out (core-ts ``entryContentOf``).
_LEFT_OUT = frozenset({f"{{{W_NS}}}proofErr", f"{{{W_NS}}}bookmarkStart", f"{{{W_NS}}}bookmarkEnd"})

#: The control kinds a made control shows Word's default placeholder for (core-ts ``showsPlaceholderWhenMade``).
_MADE_SHOWING = frozenset({"RichText", "PlainText"})


# ---------------------------------------------------------------------------
# the reference
# ---------------------------------------------------------------------------


def placeholder_entry_of(sdt_pr: Any) -> str | None:
    """The entry a ``w:sdtPr``'s ``w:placeholder`` names: ``""`` for one naming none, None for no ``w:placeholder``."""
    item = sdt_property(sdt_pr, "placeholder", W_NS)
    if item is None:
        return None
    value = getattr(getattr(item, "doc_part", None), "val", None)
    return "" if value is None else str(value)


def name_placeholder_entry(control: ContentControl, entry: str) -> None:
    """Name `entry` in the control's ``w:placeholder``, where Word writes it: after ``w:id``, before the flag and binding.

    core-ts ``namePlaceholderEntry`` (check 64's controls, as Word saved
    them); :meth:`~.controls.ContentControl.put_property` keeps Word's order.
    """
    from docx4j_py.wml import CtPlaceholderDocPart, el

    control.put_property(el.placeholder(doc_part=CtPlaceholderDocPart(val=entry)))


def new_entry_name() -> str:
    """A new entry's name, as Word names one: 32 upper-case hex digits, a new GUID's (core-ts check 65)."""
    return uuid.uuid4().hex.upper()


def shows_placeholder_when_made(kind: str | None) -> bool:
    """Whether a control of this kind made on an empty range shows Word's default placeholder, as Word's does.

    core-ts ``showsPlaceholderWhenMade`` (D8; check 64): a rich text or plain
    text control. Word's other kinds have entries of their own, unmeasured.
    """
    return kind is None or str(kind) in _MADE_SHOWING


# ---------------------------------------------------------------------------
# the entries
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class GlossaryEntry:
    """An entry: its text, its body's paragraphs ``\\r`` apart, and the body's blocks (not copies)."""

    text: str
    body: tuple[Any, ...]


class GlossaryEntries:
    """The placeholder entries of a package's glossary, and those waiting for the save. Extension.

    core-ts ``GlossaryEntries`` (``pkg.glossaryEntries``): those the glossary
    document part holds, read privately (the part stays untouched), and those
    written since, which go into the glossary when the package is saved
    (:meth:`flush`, run by the save through the package's ``save_hooks``).
    """

    __slots__ = ("_package", "_pending", "_read", "_source")

    def __init__(self, package: Any) -> None:
        """The entries of `package`'s glossary; nothing is read yet."""
        self._package = package
        #: Entries by name, as the glossary part was last read; None until :meth:`load`.
        self._read: dict[str, GlossaryEntry] | None = None
        #: The part they were read from.
        self._source: Any = None
        #: Entries to write on save: name -> text, in the order named.
        self._pending: dict[str, str] = {}

    def _part(self) -> Any:
        """The package's glossary document part, or None."""
        try:
            main = self._package.main_document_part
        except Exception:  # noqa: BLE001 - a package without a main document part has no glossary
            return None
        return getattr(main, "glossary_document_part", None) if main is not None else None

    def load(self) -> None:
        """Read the glossary's entries, privately; again only where the part is another.

        A glossary that will not parse reads as holding no entry, as core-ts's
        ``readContents().catch(...)`` has it: a placeholder is no reason for a
        read to fail.
        """
        part = self._part()
        if part is self._source and self._read is not None:
            return
        self._source = part
        tree = None
        if part is not None:
            try:
                tree = part.read_contents()
            except Exception:  # noqa: BLE001 - an unreadable glossary holds no entry
                tree = None
        self._read = entries_of(tree)

    @property
    def loaded(self) -> bool:
        """True once the glossary has been read (:meth:`load`), or its part is unmarshalled."""
        return self._read is not None or bool(getattr(self._source, "is_unmarshalled", False))

    def _entry(self, name: str) -> GlossaryEntry | None:
        pending = self._pending.get(name)
        if pending is not None:
            from docx4j_py.model.content.controls import _placeholder_run
            from docx4j_py.wml import P

            return GlossaryEntry(pending, (P(content=ChildList([_placeholder_run(pending, None)])),))
        if self._part() is not self._source or not self.loaded:
            # read here on first use too: reads are synchronous, so nothing waits for a body to be read
            self.load()
        part = self._source
        if part is not None and getattr(part, "is_unmarshalled", False):
            return entries_of(part.contents).get(name)
        return (self._read or {}).get(name)

    def text_of(self, name: str) -> str | None:
        """An entry's text, its paragraphs ``\\r`` apart; None where neither the glossary nor this session holds it.

        As Office JS reads ``placeholderText`` (core-ts check 65).
        """
        entry = self._entry(name)
        return None if entry is None else entry.text

    def body_of(self, name: str) -> tuple[Any, ...] | None:
        """An entry's blocks (not copies); None where neither the glossary nor this session holds it.

        An entry written this session is one paragraph, one run of its text in
        the Placeholder Text style.
        """
        entry = self._entry(name)
        return None if entry is None else entry.body

    def has(self, name: str) -> bool:
        """Whether the glossary holds an entry of this name, or one is waiting to be written."""
        return self._entry(name) is not None

    def require(self, name: str, text: str) -> bool:
        """An entry to write on save, where neither the glossary nor this session holds one of the name; whether it was new."""
        if self.has(name):
            return False
        self._pending[name] = text
        return True

    @property
    def waiting(self) -> bool:
        """True while entries wait for the save."""
        return bool(self._pending)

    @property
    def pending(self) -> dict[str, str]:
        """The entries waiting for the save, name -> text, a copy."""
        return dict(self._pending)

    def copy_to(self, target: GlossaryEntries) -> None:
        """Give `target` the entries waiting here, as a copy of the package takes them (core-ts ``copyTo``)."""
        for name, text in self._pending.items():
            target._pending.setdefault(name, text)

    def flush(self, package: Any = None) -> list[str]:
        """Write the waiting entries into the glossary; the names of the parts made. The save's hook.

        The glossary is made where the file has none (D9 (a): the glossary
        document part, related from the main document part, its content type
        an override), each entry is added in Word's form (:func:`doc_part_of`),
        and the glossary's styles part is made or completed
        (:func:`ensure_glossary_styles`): an entry whose style its glossary
        lacks loses it in Word's save (core-ts check 65's 65b). Nothing is
        done, and nothing read, while no entry waits.
        """
        if not self._pending:
            return []
        from docx4j_py.openpackaging.parts.wml import GlossaryDocumentPart
        from docx4j_py.wml import CTDocParts, GlossaryDocument

        package = self._package if package is None else package
        main = package.main_document_part
        added: list[str] = []
        glossary = main.glossary_document_part
        if glossary is None:
            glossary = GlossaryDocumentPart(GLOSSARY_DOCUMENT_NAME)
            glossary.set_contents(GlossaryDocument(doc_parts=CTDocParts(doc_part=ChildList([]))))
            main.add_target_part(glossary)
            added.append(str(glossary.part_name))
        document = glossary.contents
        if document.doc_parts is None:
            document.doc_parts = CTDocParts(doc_part=ChildList([]))
            link_parents(document)
        present = entries_of(document)
        for name, text in self._pending.items():
            if name in present:
                continue
            document.doc_parts.doc_part.append(doc_part_of(name, text))
        self._pending.clear()
        self._source = glossary
        added.extend(ensure_glossary_styles(glossary))
        return added


def glossary_entries_of(package: Any) -> GlossaryEntries | None:
    """A package's :class:`GlossaryEntries`, made on first use; None for an object that is no package.

    A real package's are written by its save (the entries' :meth:`~GlossaryEntries.flush`
    registered in ``package.save_hooks``); a trial's are its own, the real
    package's waiting entries copied in, and are thrown away with it, a trial
    never being saved.
    """
    if package is None:
        return None
    try:
        found = package._glossary_entries
    except AttributeError:
        return None
    if found is not None:
        return found
    from docx4j_py.model.content.trial import TrialPackage

    found = GlossaryEntries(package)
    package._glossary_entries = found
    if isinstance(package, TrialPackage):
        real = glossary_entries_of(package._package)
        if real is not None:
            real.copy_to(found)
    else:
        hooks = getattr(package, "save_hooks", None)
        if isinstance(hooks, list):
            hooks.append(found.flush)
    return found


def load_glossary_entries(package: Any) -> None:
    """Read a package's glossary entries, as the content API does when it first reads a body (core-ts ``getBody``)."""
    entries = glossary_entries_of(package)
    if entries is not None:
        entries.load()


def entries_of(document: Any) -> dict[str, GlossaryEntry]:
    """A ``w:glossaryDocument``'s entries by name: each named ``w:docPart``'s text and body."""
    from docx4j_py.wml import CTDocPartName, P

    out: dict[str, GlossaryEntry] = {}
    doc_parts = getattr(document, "doc_parts", None)
    for part in getattr(doc_parts, "doc_part", None) or ():
        pr = getattr(part, "doc_part_pr", None)
        name = next((item for item in (getattr(pr, "content", None) or ()) if isinstance(item, CTDocPartName)), None)
        value = getattr(name, "val", None)
        if not value:
            continue
        body = tuple(getattr(getattr(part, "doc_part_body", None), "content", None) or ())
        text = "\r".join(text_of(block) for block in body if isinstance(block, P))
        out.setdefault(str(value), GlossaryEntry(text, body))
    return out


# ---------------------------------------------------------------------------
# what a control shows of an entry (core-ts section 63.18)
# ---------------------------------------------------------------------------


def entry_run_properties(entry: Any, control: Any) -> Any:
    """An entry run's properties as a control shows them: the control's ``w:sdtPr/w:rPr`` over the run's own.

    core-ts ``entryContentOf``'s rule at ``aefc41f`` (its section 63.18, as
    docx4j's ``createPlaceholder`` applies them; docx4j CR-063 D2): each
    property the control's run properties state replaces the run's, the rest
    kept; with no control properties, the run's own. **The one place the
    combination is made**, for core-ts's check 68 (CR-019 section 12): Word
    brought an entry back as the glossary holds it, without the control's
    properties over it, and core-ts has proposed revising the rule.

    Returns:
        A new ``w:rPr``, or None where neither states anything.
    """
    from docx4j_py.wml import RPr

    if control is None:
        return deep_copy(entry) if entry is not None else None
    out = deep_copy(entry) if entry is not None else RPr()
    for field in dataclasses.fields(control):
        value = getattr(control, field.name)
        if value is None or (isinstance(value, (list, dict)) and not value):
            continue
        setattr(out, field.name, deep_copy(value))
    link_parents(out)
    return out


def entry_content_of(body: tuple[Any, ...] | list[Any], run_properties: Any) -> list[list[Any]]:
    """What a control shows of an entry: a copy of each of its paragraphs' run content, a list per paragraph.

    core-ts ``entryContentOf`` (section 63.18): every item as the glossary
    holds it but a proof error and a bookmark's ends, each run's properties
    :func:`entry_run_properties` with the control's `run_properties`. The
    paragraphs' properties stay in the glossary, whose styles they name.
    """
    from docx4j_py.wml import P, R

    out: list[list[Any]] = []
    for block in body:
        if not isinstance(block, P):
            continue
        copies = [deep_copy(item) for item in (block.content or ()) if element_name(item) not in _LEFT_OUT]
        for item in copies:
            if isinstance(item, R):
                item.r_pr = entry_run_properties(item.r_pr, run_properties)
                if item.r_pr is not None:
                    item.r_pr.parent = item
        out.append(copies)
    return out


# ---------------------------------------------------------------------------
# what a save writes (D9 (a))
# ---------------------------------------------------------------------------

_W_DECLARATION = f'xmlns:w="{W_NS}"'


def doc_part_of(name: str, text: str) -> Any:
    """A placeholder entry as Word writes one (core-ts check 65; ``docPartOf``).

    The default entry's ``w:docPartPr`` (category General, gallery
    placeholder, type ``bbPlcHdr``, behaviour content) with a ``w:guid`` of its
    own, and a body of one paragraph, one run of the text in the Placeholder
    Text style.
    """
    from docx4j_py.fragments import wml

    guid = "{" + str(uuid.uuid4()).upper() + "}"
    space = ' xml:space="preserve"' if text != text.strip(" ") else ""
    xml = (
        f'<w:docPart {_W_DECLARATION}><w:docPartPr><w:name w:val="{escape(name, {chr(34): "&quot;"})}"/>'
        '<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>'
        '<w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors>'
        f'<w:guid w:val="{guid}"/></w:docPartPr><w:docPartBody><w:p><w:r><w:rPr>'
        f'<w:rStyle w:val="{PLACEHOLDER_STYLE}"/></w:rPr><w:t{space}>{escape(text)}</w:t></w:r></w:p>'
        "</w:docPartBody></w:docPart>"
    )
    part = wml(xml)
    link_parents(part)
    return part


def default_paragraph_font() -> Any:
    """Default Paragraph Font as Word defines it, which Placeholder Text is based on (check 34's glossary)."""
    from docx4j_py.fragments import wml

    return wml(
        f'<w:style {_W_DECLARATION} w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
    )


def placeholder_style() -> Any:
    """Word 365's Placeholder Text (core-ts ``placeholderStyle``; check 66's ``PLACEHOLDER_STYLES``), a new copy each call."""
    from docx4j_py.fragments import wml

    return wml(
        f'<w:style {_W_DECLARATION} w:type="character" w:styleId="{PLACEHOLDER_STYLE}">'
        '<w:name w:val="Placeholder Text"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/>'
        '<w:unhideWhenUsed/><w:rPr><w:color w:val="666666"/></w:rPr></w:style>'
    )


def require_placeholder_style(package: Any) -> None:
    """The main document's styles define Placeholder Text, which a placeholder's runs name; added where they lack it.

    core-ts's ``requireStyles([PLACEHOLDER_STYLE])`` in ``setBoundContent`` (its
    CR-002 section 40, check 35: Word 2010 drops a ``w:rStyle`` naming a style
    the part lacks), with its definition, Word 365's (:func:`placeholder_style`).
    core-ts splices it on save; here it is added at once, as this package
    defines a hyperlink's style (CR-003 section 21.3 item 2), so a read before
    the save sees it and the write's report names ``/word/styles.xml``. A styles
    part that defines it already is only read, privately (the property
    resolver's reading). A package with no styles part is left as it is.
    """
    part = getattr(package, "style_definitions_part", None) if package is not None else None
    if part is None:
        return
    try:
        resolver = package.get_property_resolver()
    except Exception:  # noqa: BLE001 - a package that cannot resolve styles is asked the part instead
        resolver = None
    if resolver is not None and resolver.get_style(PLACEHOLDER_STYLE) is not None:
        return
    styles = part.contents
    if styles.style is None:
        styles.style = ChildList([], owner=styles)
    if any(getattr(style, "style_id", None) == PLACEHOLDER_STYLE for style in styles.style):
        return
    style = placeholder_style()
    styles.style.append(style)
    link_parents(style)
    style.parent = styles
    if resolver is not None:
        resolver.refresh()
    from docx4j_py.model.content.reports import current_recorder

    current_recorder(package).part(part)


def has_placeholder_styles(part: Any) -> bool:
    """Whether a styles part defines Default Paragraph Font and Placeholder Text, read privately."""
    try:
        styles = part.read_contents()
    except Exception:  # noqa: BLE001 - a styles part that will not parse is completed by its live tree, or fails the save
        return False
    ids = {getattr(style, "style_id", None) for style in (getattr(styles, "style", None) or ())}
    return "DefaultParagraphFont" in ids and PLACEHOLDER_STYLE in ids


def ensure_glossary_styles(glossary: Any) -> list[str]:
    """The glossary's styles part holds Placeholder Text and the style it is based on; the names of the parts made.

    core-ts ``ensureGlossaryStyles``: a styles part made where the glossary
    has none, related from it; one lacking either style completed; one holding
    both left untouched, byte for byte.
    """
    from docx4j_py.openpackaging.parts.wml import StyleDefinitionsPart
    from docx4j_py.wml import Styles

    part = getattr(glossary, "style_definitions_part", None)
    added: list[str] = []
    if part is None:
        part = StyleDefinitionsPart(GLOSSARY_STYLES_NAME)
        part.set_contents(Styles(style=ChildList([])))
        glossary.add_target_part(part)
        added.append(str(part.part_name))
    elif has_placeholder_styles(part):
        return added
    styles = part.contents
    if styles.style is None:
        styles.style = ChildList([])
    present = {getattr(style, "style_id", None) for style in styles.style}
    for style in (default_paragraph_font(), placeholder_style()):
        if style.style_id in present:
            continue
        styles.style.append(style)
    link_parents(styles)
    return added
