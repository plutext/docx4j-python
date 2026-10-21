"""One text model: the text of a document, as docx4j's ``org.docx4j.model.text`` reads it. CR-011.

The port of docx4j CR-046 (the portfolio's *WordprocessingML Text Extraction*,
WD 0.28), after core-ts CR-012. A text is
``serialize(scope(view(resolve(package))))`` (WD REQ-201): :func:`extract`
makes the block model (:class:`TextResult`), and :func:`plain_text` writes plain
text from it::

    from docx4j_py.model.text import extract, plain_text

    result = extract(pkg, view="markup", include="headers, footers", hidden=True)
    print(plain_text(result))
    for warning in result.warnings:
        print(warning)

* **resolve**: each ``mc:AlternateContent`` is read from the first
  ``mc:Choice`` whose ``Requires`` namespaces are all in docx4j's published
  understood list, matched by URI (REQ-302 to 306), else its ``mc:Fallback``,
  else its first Choice; ``mce.branches_differ`` where the branches' texts
  differ (REQ-307).
* **view**: ``accepted`` (the default; ``current`` is a synonym), ``original``
  or ``markup`` (REQ-401). The first two are the text of the document that
  CR-010's ``accept_all`` or ``reject_all`` leaves, worked out without
  resolving anything (:class:`~docx4j_py.model.text.view.ResolvedView`), and
  held equal to that copy by the agreement test (REQ-406). ``markup`` reads
  every run as stored, with its revision, written as CriticMarkup.
* **scope**: WD section 5 under section 10's options (:class:`TextOptions`,
  named verbatim): the main story; each text box once, after its anchoring
  paragraph (``body/3/textbox:0/1``); the notes in reference order, ``[n]``;
  headers and footers by section under ``include``; comments in ``markup``
  or under ``comments: true``; field codes under ``field_codes``; hidden text
  left out, through the styles; placeholders; list labels under
  ``list_labels``, with the level's owner (REQ-510); the table of contents
  only under ``toc: true`` (REQ-520); the document properties in the
  result's metadata, and in the text before the body under ``metadata: true``
  (REQ-515); a chart as a ``CHART`` block after its anchor, its description's
  note (docx4j CR-047), title and table (:mod:`docx4j_py.model.chart`), and
  SmartArt's lines in reading order (REQ-511); legacy WordArt as a
  one-paragraph text box (REQ-511); each
  picture an ``IMAGE`` run, described by the options' describer (the alt text
  by default; ``[image: ...]`` under ``alt_text: true``, REQ-514), and each
  object an ``OBJECT`` run with its ProgID, an embedded WordprocessingML
  package read one level down under ``embedded: true`` (REQ-512); a
  plain-text, WordprocessingML or Flat OPC ``altChunk`` expanded where it is
  (REQ-513); the glossary last under ``include: glossary`` (REQ-516);
  equations as LaTeX
  (REQ-604, docx4j's ``OmmlToLatex``, :mod:`docx4j_py.model.math`); REQ-605's
  table forms; ``[page n]`` under ``page_markers``; ``layout: indent``; a
  bound plain-text control's bound value, through the custom XML parts or the
  document properties (REQ-509, ``sdt.binding_differs``).

**The package is never changed** (REQ-204): every part is read with
``read_contents``, so a save afterwards is byte for byte where it was.

**Not the Office JS readers' text.** ``Paragraph.text``, ``Body.text``,
``get_text`` (in ``accepted`` and ``original``) and the other content API
readers keep Office JS's text (CR-011 Q4, core-ts CR-012's D2, WD REQ-202's
exemption for a reader of an emulated API). Each :class:`TextRun` records
where its text lies in its paragraph's ``get_text()`` (``office_offset``), so
the two map onto each other. ``get_text(view="markup")`` is this model's.

Under ``fields: update`` (REQ-506) the fields that need neither layout nor
fetching are computed (:mod:`docx4j_py.model.fields`), and a FORMTEXT is as
Word's update leaves it.

**Deferred** (CR-011 Q7): embedded SpreadsheetML and PresentationML packages
and OLE 1.0 packages, which under ``embedded: true`` raise
``content.not_read``; and the HTML reader (Q12, Q13), without which an HTML,
XHTML or MHT ``altChunk`` raises ``altchunk.not_expanded``.

Imported on first use: ``import docx4j_py`` does not import it.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.text.extractor import normalize
from docx4j_py.model.text.images import (
    ImageDescriber,
    ImageDescription,
    TextImage,
    alt_text_describer,
)
from docx4j_py.model.text.options import (
    NOT_YET,
    OPTION_NAMES,
    TextOptionError,
    TextOptions,
    TextView,
    view_of,
)
from docx4j_py.model.text.plain import plain_text
from docx4j_py.model.text.result import (
    TextBlock,
    TextCell,
    TextComment,
    TextField,
    TextNumbering,
    TextResult,
    TextRevision,
    TextRow,
    TextRun,
    TextRunSource,
    TextWarning,
)

__all__ = [
    "NOT_YET",
    "OPTION_NAMES",
    "ImageDescriber",
    "ImageDescription",
    "TextBlock",
    "TextCell",
    "TextComment",
    "TextField",
    "TextImage",
    "TextNumbering",
    "TextOptionError",
    "TextOptions",
    "TextResult",
    "TextRevision",
    "TextRow",
    "TextRun",
    "TextRunSource",
    "TextView",
    "TextWarning",
    "alt_text_describer",
    "extract",
    "extract_resolved",
    "normalize",
    "plain_text",
    "view_of",
]


def _options(options: TextOptions | dict[str, Any] | None, more: dict[str, Any]) -> TextOptions:
    out = TextOptions.of(options, **more)
    out.check_implemented()
    return out


def _package_of(target: Any) -> Any:
    return getattr(target, "package", None)


def extract(
    target: Any, options: TextOptions | dict[str, Any] | None = None, *, package: Any = None, **more: Any
) -> TextResult:
    """The text of a package, or of content in one, in the block model. docx4j ``Texts.of``.

    Args:
        target: a ``WordprocessingMLPackage`` (REQ-501's scope); a story
            part of one, read as docx4j's ``Texts.of(Part)`` reads it (the
            main document part, its story and text boxes addressed ``body``
            and no notes; a header or a footer, ``header:rId4``); a content API
            view (``Body``, ``Paragraph``, ``Table``); or a model object of a
            package's tree (a ``w:p``, a ``w:tbl``, a ``w:body``) with
            `package`. Content is read as the main story in the package's
            view, so a join or a move resolves as it does in the whole
            document, its addresses starting ``content``. Content with no
            package is read as it stands: no styles, and in ``accepted`` and
            ``original`` run-level revisions only.
        options: a :class:`TextOptions`, or a mapping of section 10's names.
        package: the package a model object belongs to.
        **more: options by name (``view="markup"``, ``hidden=True``).

    Raises:
        TextOptionError: an unknown option, an invalid value, or one refused
            (REQ-1001).
    """
    from docx4j_py.model.text.extractor import Extractor
    from docx4j_py.model.text.trees import PackageTrees

    opts = _options(options, more)
    if getattr(target, "main_document_part", None) is not None:
        return Extractor(PackageTrees(target), opts).extract()
    part = None
    content = target
    owner = package
    if hasattr(target, "container") and hasattr(target, "part") and hasattr(target, "package"):
        # a content API Body
        content, part, owner = target.container, target.part, owner or target.package
    elif hasattr(target, "element") and hasattr(target, "parent_body"):
        # a content API Paragraph or Table
        body = target.parent_body
        content, part, owner = target.element, body.part, owner or body.package
    elif hasattr(target, "contents") and hasattr(target, "part_name"):
        # a story part, docx4j's Texts.of(Part): the main document part (its story and its text
        # boxes, addressed body; no notes), a header or a footer (header:rId4); no first reading
        owner = owner or _package_of(target)
        if owner is None:
            raise TypeError(f"{target.part_name} belongs to no package; read its content with extract(content)")
        trees = PackageTrees(owner)
        if target is getattr(owner, "main_document_part", None):
            return Extractor(trees, opts).extract_content(trees.document, target, "body", "MAIN")
        found = trees.headers_footers.get(id(target))
        if found is None:
            raise TypeError(f"not a story part: {target.part_name}; use extract(package)")
        rel_id = trees.relationship_id(target)
        kind = "HEADER" if type(target).__name__ == "HeaderPart" else "FOOTER"
        prefix = f"{kind.lower()}:{rel_id}"
        return Extractor(trees, opts).extract_content(found[1], target, prefix, kind)
    # content with no package is read as it stands: no styles, and in the
    # accepted and original views run-level revisions only (no join, no move)
    trees = PackageTrees(owner)
    return Extractor(trees, opts).extract_content(content, part)


def extract_resolved(resolved: Any, options: TextOptions | dict[str, Any] | None = None, **more: Any) -> TextResult:
    """The text of a package ``accept_all`` or ``reject_all`` has resolved already, read with no view logic.

    The copy route of WD REQ-406, which :func:`extract`'s views are held equal
    to (``tests/text/test_agreement.py``). The options' view says which
    resolution it was (``accepted`` or ``original``).
    """
    from docx4j_py.model.text.extractor import Extractor
    from docx4j_py.model.text.trees import PackageTrees

    opts = _options(options, more)
    if opts.view == "markup":
        raise TextOptionError("A resolved package has no markup view", "view", "REQ-406", "invalid")
    return Extractor(PackageTrees(resolved), opts, resolved=True).extract()
