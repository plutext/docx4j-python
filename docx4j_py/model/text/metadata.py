"""The document properties of a result's metadata (WD REQ-515). CR-011 section 6.

docx4j's ``org.docx4j.model.text.Metadata``: the core, application and custom
properties, by their element or property names, read without changing the
package (each part with ``read_contents``, so nothing is unmarshalled and no
part is added).

* **core** (``docProps/core.xml``), in the schema's order: ``title``,
  ``subject``, ``creator``, ``keywords``, ``description``, ``category``,
  ``contentStatus``, ``language``, ``identifier``, ``version``, ``revision``,
  ``lastModifiedBy``, ``created``, ``modified``, ``lastPrinted``;
* **app** (``docProps/app.xml``): what describes the document --
  ``Application``, ``AppVersion``, ``Company``, ``Manager``, ``Template``,
  ``HyperlinkBase``, ``TotalTime``, ``DocSecurity``; the statistics (pages,
  words, characters, lines, paragraphs) are left out, being as of the last
  save and dependent on layout;
* **custom** (``docProps/custom.xml``): each property's typed value (a string,
  a number, a boolean or a W3C date as written), in the part's order; values of
  other types (a vector, a blob, a class id) left out.

An empty string is left out, as docx4j leaves it out. The maps are in the
result's metadata whatever the options (REQ-801); the text holds them only
under ``metadata: true``, as a line a kind before the body (REQ-515).
"""

from __future__ import annotations

from typing import Any

__all__ = ["APP", "CORE", "CUSTOM", "read"]

#: The metadata's keys (docx4j ``TextResult.CORE``, ``APP``, ``CUSTOM``).
CORE = "core"
APP = "app"
CUSTOM = "custom"

#: The core properties: (the field, the element's name), in the schema's order.
_CORE = (
    ("title", "title"),
    ("subject", "subject"),
    ("creator", "creator"),
    ("keywords", "keywords"),
    ("description", "description"),
    ("category", "category"),
    ("content_status", "contentStatus"),
    ("language", "language"),
    ("identifier", "identifier"),
    ("version", "version"),
    ("revision", "revision"),
    ("last_modified_by", "lastModifiedBy"),
    ("created", "created"),
    ("modified", "modified"),
    ("last_printed", "lastPrinted"),
)

#: The application properties that describe the document: (the field, the element's name).
_APP = (
    ("application", "Application"),
    ("app_version", "AppVersion"),
    ("company", "Company"),
    ("manager", "Manager"),
    ("template", "Template"),
    ("hyperlink_base", "HyperlinkBase"),
    ("total_time", "TotalTime"),
    ("doc_security", "DocSecurity"),
)

#: The custom property value types docx4j reads (``Metadata.value``), by class name.
_CUSTOM_TYPES = frozenset(
    {
        "Lpwstr", "Lpstr", "Bstr", "Bool", "I1", "I2", "I4", "I8", "Int", "Ui1", "Ui2", "Ui4", "Ui8", "Uint",
        "R4", "R8", "DecimalType", "Filetime", "Date", "Cy", "Error",
    }
)  # fmt: skip


def read(package: Any, metadata: dict[str, Any]) -> None:
    """The properties into `metadata`, under :data:`CORE`, :data:`APP` and :data:`CUSTOM`, where the parts exist."""
    if package is None:
        return
    core = _contents(getattr(package, "doc_props_core_part", None))
    if core is not None:
        metadata[CORE] = _core(core)
    app = _contents(getattr(package, "doc_props_extended_part", None))
    if app is not None:
        metadata[APP] = _app(app)
    custom = _contents(getattr(package, "doc_props_custom_part", None))
    if custom is not None:
        metadata[CUSTOM] = _custom(custom)


def _contents(part: Any) -> Any:
    if part is None:
        return None
    try:
        return part.read_contents()
    except Exception:  # noqa: BLE001 - a part that will not read has no properties here
        return None


def _put(out: dict[str, Any], name: str, value: Any) -> None:
    if value is None or (isinstance(value, str) and not value):
        return
    out[name] = value


def date_text(value: Any) -> str | None:
    """A date as the part wrote it (an ``XmlDateTime``'s lexical form)."""
    if value is None:
        return None
    to_string = getattr(value, "to_string", None)
    return to_string() if callable(to_string) else str(value)


def _literal(value: Any) -> Any:
    """A Dublin Core value: a ``SimpleLiteral``'s text, a ``W3CDTF``'s date, or the value itself."""
    if value is None:
        return None
    inner = getattr(value, "value", None)
    if type(value).__name__ == "DerivedElement":
        value = inner
        inner = getattr(value, "value", None)
    any_element = getattr(value, "any_element", None)
    if any_element is not None:
        return getattr(any_element, "text", None)
    if type(value).__name__ == "W3CDTF":
        content = getattr(value, "content", None)
        if content:
            return "".join(str(c) for c in content)
        return date_text(inner)
    if hasattr(value, "to_string"):
        return date_text(value)
    return value if isinstance(value, (str, int, float, bool)) else None


def _core(c: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for field, name in _CORE:
        _put(out, name, _literal(getattr(c, field, None)))
    return out


def _app(p: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for field, name in _APP:
        _put(out, name, getattr(p, field, None))
    return out


def _custom(p: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for prop in getattr(p, "property", None) or ():
        name = getattr(prop, "name", None)
        if name is None:
            continue
        _put(out, name, _value(getattr(prop, "content", None)))
    return out


def _value(content: Any) -> Any:
    """A custom property's typed value; None for a type docx4j does not read."""
    if content is None or type(content).__name__ not in _CUSTOM_TYPES:
        return None
    value = getattr(content, "value", None)
    if type(content).__name__ in ("Filetime", "Date"):
        return date_text(value)
    return value


def shown(value: Any) -> str:
    """A property's value as plain text writes it: Java's ``String.valueOf``, booleans in lower case."""
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)
