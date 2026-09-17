"""One error hierarchy, and every message says what to do instead.

CR-003 section 3.1: the content API raises nothing bare. Every error carries a
stable :attr:`~ContentError.code` (``"style.not_found"``), the human message,
and a :attr:`~ContentError.hint`, one sentence an agent can act on. The root is
CR-002's :class:`~docx4j_py.openpackaging.exceptions.Docx4JException`, exported
here under CR-003's name :data:`Docx4JError` as well, so that
``except Docx4JException`` catches the content API and ``except ContentError``
catches only it.

::

    Docx4JException  (= Docx4JError, CR-002)
    └── ContentError            code, message, hint
        ├── AddressError        no block at that address
        ├── InvalidTargetError  this container cannot hold that element
        ├── StyleError          no such style, and the five closest names
        ├── SpanError           a span that cannot be formed or edited
        ├── TrackedChangeError  an edit a tracked document cannot carry
        ├── BindingError        a content control's XML mapping cannot be used
        └── BuilderError        a tree-layer builder was given the wrong thing

:class:`BuilderError` is CR-003 Phase A's, re-rooted here as Phase A's notes
(section 10.4) require, with its ``code`` strings unchanged. It keeps
``ValueError`` as a second base so that the call sites written against Phase A
(``pytest.raises(ValueError)``) go on working;
:mod:`docx4j_py.wml.builders` imports it from here.
"""

from __future__ import annotations

from docx4j_py.openpackaging.exceptions import Docx4JException

__all__ = [
    "AddressError",
    "BindingError",
    "BuilderError",
    "ContentError",
    "Docx4JError",
    "InvalidTargetError",
    "SpanError",
    "StyleError",
    "TrackedChangeError",
]

#: CR-003's name for CR-002's root. The same class, not a subclass: one
#: hierarchy means one root, and CR-002 got there first.
Docx4JError = Docx4JException


class ContentError(Docx4JError):
    """The content API refuses to do something, and says what to do instead.

    Attributes:
        code: a stable string such as ``"target.invalid"``, for a tool result.
        message: the human half, without the hint.
        hint: one sentence an agent can act on.
    """

    #: The code a subclass uses when the caller gives none.
    default_code = "content.error"
    #: The hint a subclass uses when the caller gives none.
    default_hint = "see the message"

    def __init__(self, message: str, *, code: str | None = None, hint: str | None = None) -> None:
        """Build the error from its message, its stable code and its hint."""
        code = code if code is not None else self.default_code
        hint = hint if hint is not None else self.default_hint
        super().__init__(f"{message} ({hint})")
        self.code = code
        self.hint = hint
        self.message = message

    def to_dict(self) -> dict[str, str]:
        """The JSON-ready view: ``code``, ``message``, ``hint``."""
        return {"code": self.code, "message": self.message, "hint": self.hint}


class AddressError(ContentError):
    """No block lives at that address any more (CR-003 section 3.4)."""

    default_code = "address.not_found"
    default_hint = "call outline() to list current addresses"


class InvalidTargetError(ContentError):
    """A container was handed an element it cannot hold.

    The message names what was passed and what the container takes, which is
    what makes an agent's next call right (CR-003 section 3.4).
    """

    default_code = "target.invalid"
    default_hint = "pass an element the container can hold"


class StyleError(ContentError):
    """No such style in this document; the message lists the closest names."""

    default_code = "style.not_found"
    default_hint = "use one of the names listed, or a built-in style id"


class SpanError(ContentError):
    """A span (a :class:`~docx4j_py.model.content.Range`) cannot be formed or edited.

    docx4j-core-ts raises ``RangeError`` for these; Python has no builtin of
    that name and ``Range`` here is the view, so the error is ``SpanError``.
    """

    default_code = "range.invalid"
    default_hint = "use get_range() or search() to obtain a span of this paragraph"


class TrackedChangeError(ContentError):
    """An edit a document with tracked changes cannot carry (CR-003 section 3.8).

    Editing text another author deleted, deleting a paragraph whose mark is
    already marked deleted, accepting or rejecting a change that is no longer in
    the tree. The message names the author where there is one, because that is
    what tells a caller whose revision is in the way.
    """

    default_code = "tracking.error"
    default_hint = "reject the revision first, or edit a span that is not deleted"


class BindingError(ContentError):
    """A content control's XML mapping cannot be used (CR-003 section 3.7).

    The custom XML part a ``w:dataBinding`` names is gone, an XPath selects
    nothing or more than one node, a typed view was asked of a control of
    another kind, a ``fill()`` key matches no binding, a repeating section was
    asked for at run level, or a span crosses a run holder. The message names
    the XPath or the store item id, because that is what tells a caller which
    binding is in the way.
    """

    default_code = "binding.error"
    default_hint = "call pkg.custom_xml_parts.describe() to list the bindings that resolve"


class BuilderError(ContentError, ValueError):
    """A tree-layer builder was given something it cannot make an element of.

    CR-003 Phase A raised this from :mod:`docx4j_py.wml.builders` as a plain
    ``ValueError`` while there was no hierarchy to root it in; Phase B moves it
    here, under :class:`ContentError`, and keeps every ``code`` string
    (``sdt.form_mismatch``, ``picture.empty_extent``,
    ``image.unsupported_format``, ``rpr.unknown_property``, ...) exactly as it
    was. ``ValueError`` stays as a second base for the code already written
    against it.
    """

    default_code = "builder.error"
    default_hint = "pass what the builder's signature asks for"

    def __init__(self, message: str, *, code: str, hint: str) -> None:
        """Build the error; Phase A's signature, where code and hint are required."""
        super().__init__(message, code=code, hint=hint)
