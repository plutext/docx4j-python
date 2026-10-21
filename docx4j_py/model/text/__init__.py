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
  (REQ-515); legacy WordArt as a one-paragraph text box (REQ-511); each
  picture an ``IMAGE`` run, described by the options' describer (the alt text
  by default; ``[image: ...]`` under ``alt_text: true``, REQ-514), and each
  object an ``OBJECT`` run with its ProgID (REQ-512).

**The package is never changed** (REQ-204): every part is read with
``read_contents``, so a save afterwards is byte for byte where it was.

**Not the Office JS readers' text.** ``Paragraph.text``, ``Body.text``,
``get_text`` (in ``accepted`` and ``original``) and the other content API
readers keep Office JS's text (CR-011 Q4, core-ts CR-012's D2, WD REQ-202's
exemption for a reader of an emulated API). Each :class:`TextRun` records
where its text lies in its paragraph's ``get_text()`` (``office_offset``), so
the two map onto each other. ``get_text(view="markup")`` is this model's.

**Not yet** (CR-011 Phase C, :data:`~docx4j_py.model.text.options.NOT_YET`):
fields computed under ``fields: update``, bound values, charts, SmartArt,
embedded packages, ``w:altChunk``, math as LaTeX (an equation is its ``m:t`` text inside ``$...$``), the
glossary.

Imported on first use: ``import docx4j_py`` does not import it.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.text.extractor import normalize
from docx4j_py.model.text.images import ImageDescriber, ImageDescription, TextImage, alt_text_describer
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
        target: a ``WordprocessingMLPackage`` (REQ-501's scope); a header or
            footer part, or the main document part, of one; a content API
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
        # a part: the main document part, a header or a footer
        owner = owner or _package_of(target)
        if owner is None:
            raise TypeError(f"{target.part_name} belongs to no package; read its content with extract(content)")
        if target is getattr(owner, "main_document_part", None):
            return Extractor(PackageTrees(owner), opts).extract()
        trees = PackageTrees(owner)
        found = trees.headers_footers.get(id(target))
        if found is None:
            raise TypeError(f"not a story part: {target.part_name}; use extract(package)")
        content, part = found[1], target
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
