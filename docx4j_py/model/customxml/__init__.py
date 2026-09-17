"""Custom XML parts, XML mapping and the typed content controls.

CR-003 section 3.7, Phase E. Office JS's WordApiDesktop 1.3 surface ---
``CustomXmlPart``, ``CustomXmlNode``, ``XmlMapping``, the seven typed
content-control kinds --- over the ``/customXml/itemN.xml`` parts CR-002 holds
as lxml trees, plus the two template verbs docx4j-mcp's ``describe_template``
and ``fill_template`` are written over::

    part = pkg.custom_xml_parts.get_item("{5D7BA57F-1E52-4637-9F82-2D4025768D4F}")
    part.select_single_node("/invoice[1]/customer[1]/company[1]").text = "Acme Ltd"
    pkg.custom_xml_parts.apply_bindings()          # what Word does on open

    pkg.custom_xml_parts.describe()                # what the template wants
    pkg.custom_xml_parts.fill({"/invoice[1]/invoicenumber[1]": "INV-9"})

**XPath is lxml's** --- XPath 1.0, complete, with a namespace map --- so there
is no engine to build, nothing to warm and nothing asynchronous;
``select_nodes(xpath, namespace_mappings=None)`` takes Word's
``xmlns:ns0='urn:invoice'`` string or the part's own namespace manager.

``pkg.custom_xml_parts``, and ``insert_content_control`` on ``Body``,
``Paragraph`` and ``Range``, import this package on **first use**, so a caller
who never touches a content control pays nothing for it. CR-003 section 5's
import direction holds: this package imports the content API and the parts
layer, and neither imports it back.
"""

from __future__ import annotations

from docx4j_py.model.content.errors import BindingError
from docx4j_py.model.customxml.bindings import (
    CONTAINER_KINDS,
    PLACEHOLDER_STYLE,
    PLACEHOLDER_TEXT,
    BindingEntry,
    BindingResult,
    apply_bindings,
    runs_for_value,
    update_from_content_controls,
)
from docx4j_py.model.customxml.insert import (
    insert_content_control_in_body,
    insert_content_control_in_paragraph,
    insert_content_control_in_range,
)
from docx4j_py.model.customxml.kinds import (
    CHECKED_SYMBOL,
    UNCHECKED_SYMBOL,
    CheckboxContentControl,
    ComboBoxContentControl,
    ContentControlListItem,
    DatePickerContentControl,
    DropDownListContentControl,
    GroupContentControl,
    ListContentControl,
    PictureContentControl,
    RepeatingSectionContentControl,
    format_date,
)
from docx4j_py.model.customxml.mapping import XmlMapping
from docx4j_py.model.customxml.nodes import (
    CustomXmlNode,
    CustomXmlPrefixMapping,
    CustomXmlPrefixMappingCollection,
)
from docx4j_py.model.customxml.parts import (
    BUILT_IN_NAMESPACES,
    CustomXmlPart,
    CustomXmlPartCollection,
    custom_xml_parts_of,
    normalise_id,
)
from docx4j_py.model.customxml.template import (
    OPENDOPE_XPATHS_NS,
    BindingInfo,
    FillEntry,
    FillResult,
    RepeatInfo,
    Skeleton,
    TemplatePart,
)
from docx4j_py.model.customxml.xpath import (
    CanonicalXPath,
    canonical_xpath_of,
    format_prefix_mappings,
    parse_prefix_mappings,
)

__all__ = [
    "BUILT_IN_NAMESPACES",
    "CHECKED_SYMBOL",
    "CONTAINER_KINDS",
    "OPENDOPE_XPATHS_NS",
    "PLACEHOLDER_STYLE",
    "PLACEHOLDER_TEXT",
    "UNCHECKED_SYMBOL",
    "BindingEntry",
    "BindingError",
    "BindingInfo",
    "BindingResult",
    "CanonicalXPath",
    "CheckboxContentControl",
    "ComboBoxContentControl",
    "ContentControlListItem",
    "CustomXmlNode",
    "CustomXmlPart",
    "CustomXmlPartCollection",
    "CustomXmlPrefixMapping",
    "CustomXmlPrefixMappingCollection",
    "DatePickerContentControl",
    "DropDownListContentControl",
    "FillEntry",
    "FillResult",
    "GroupContentControl",
    "ListContentControl",
    "PictureContentControl",
    "RepeatInfo",
    "RepeatingSectionContentControl",
    "Skeleton",
    "TemplatePart",
    "XmlMapping",
    "apply_bindings",
    "canonical_xpath_of",
    "custom_xml_parts_of",
    "format_date",
    "format_prefix_mappings",
    "insert_content_control_in_body",
    "insert_content_control_in_paragraph",
    "insert_content_control_in_range",
    "normalise_id",
    "parse_prefix_mappings",
    "runs_for_value",
    "update_from_content_controls",
]


# The members this package adds to the views are installed by
# :func:`docx4j_py.model.content.register` --- ``pkg.custom_xml_parts`` as a
# property that imports this package on first use, and
# ``insert_content_control`` as a method of ``Body``, ``Paragraph`` and
# ``Range`` --- so that importing ``docx4j_py`` costs nothing for a caller who
# never touches a content control. CR-003 section 5's import direction is
# unchanged: this package imports the content API and the parts layer, and
# neither imports it back.
