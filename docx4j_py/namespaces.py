"""docx4j's namespace prefix table.

CR-001 section 7: "The prefix table is docx4j's ``NamespacePrefixMapper`` as a
Python dict ... used by ``wml(...)`` and by the engine."

The source of truth in docx4j is
``docx4j-core/src/main/java/org/docx4j/jaxb/NamespacePrefixMappings.java``
(``getPreferredPrefixStatic``), which itself follows the Open XML SDK's
``NamespaceIdMap.cs``. Every mapping below was extracted mechanically from that
method on 2026-09-12; the commented-out branches in the Java were skipped, and
the one branch that returns a prefix conditionally --- SpreadsheetML main, which
is ``""`` (the default namespace) unless a prefix is required, in which case it
is ``s`` --- is written here as ``s``, the only hand addition.

Three things are exported:

``PREFIXES``
    prefix -> namespace URI, the table itself.
``NAMESPACE_TO_PREFIX``
    the inverse, which is how docx4j stores it (URI in, prefix out).
``MC_IGNORABLE_PREFIXES``
    the subset a part root may legitimately name in ``mc:Ignorable``: the
    Office extension namespaces that markup-compatibility versioning was
    invented for. Word writes ``mc:Ignorable="w14 wp14"`` or
    ``mc:Ignorable="w14 w15 w16se w16cid wp14"`` on ``document.xml``; every
    prefix it names has to be declared on the same element or the file is
    corrupt, which is the reconciliation CR-001 section 13.5 point 7 leaves to
    the engine (CR-002). Here it is what :func:`declarations` guarantees to
    emit for a fragment.

and two helpers, :func:`declarations` and :func:`W_NAMESPACE_DECLARATION`, the
analogues of docx4j's ``Namespaces.W_NAMESPACE_DECLARATION``.
"""

from __future__ import annotations

__all__ = [
    "MC_IGNORABLE_PREFIXES",
    "NAMESPACE_TO_PREFIX",
    "PREFIXES",
    "UNDERSTOOD",
    "W_NAMESPACE_DECLARATION",
    "WML_NS",
    "XML_NS",
    "declarations",
    "ns_map",
    "qname",
    "split_qname",
]

#: Every prefix docx4j's ``NamespacePrefixMappings`` prefers, prefix -> URI.
PREFIXES: dict[str, str] = {
    # WordprocessingML and the package
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "pkg": "http://schemas.microsoft.com/office/2006/xmlPackage",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "prop": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "properties": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mdssi": "http://schemas.openxmlformats.org/package/2006/digital-signature",
    # DrawingML
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
    "b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "dsp": "http://schemas.microsoft.com/office/drawing/2008/diagram",
    "cdr14": "http://schemas.microsoft.com/office/drawing/2010/chartDrawing",
    "dgm14": "http://schemas.microsoft.com/office/drawing/2010/diagram",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "pic14": "http://schemas.microsoft.com/office/drawing/2010/picture",
    "c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
    "cs": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
    "a15": "http://schemas.microsoft.com/office/drawing/2012/main",
    "cdr": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "comp": "http://schemas.openxmlformats.org/drawingml/2006/compatibility",
    "lc": "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "xdr14": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
    # VML and the old Office namespaces
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "xvml": "urn:schemas-microsoft-com:office:excel",
    "pvml": "urn:schemas-microsoft-com:office:powerpoint",
    "mv": "urn:schemas-microsoft-com:mac:vml",
    "WX": "http://schemas.microsoft.com/office/word/2003/auxHint",
    "aml": "http://schemas.microsoft.com/aml/2001/core",
    # Word extensions
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wp15": "http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    # PowerPoint
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p13cmd": "http://schemas.microsoft.com/office/powerpoint/2013/main/command",
    "iact": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "p1510": "http://schemas.microsoft.com/office/powerpoint/2015/10/main",
    "p16": "http://schemas.microsoft.com/office/powerpoint/2015/main",
    "psez": "http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom",
    "pslz": "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom",
    "psuz": "http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom",
    "p166": "http://schemas.microsoft.com/office/powerpoint/2016/6/main",
    "p1710": "http://schemas.microsoft.com/office/powerpoint/2017/10/main",
    "p173": "http://schemas.microsoft.com/office/powerpoint/2017/3/main",
    "p184": "http://schemas.microsoft.com/office/powerpoint/2018/4/main",
    # SpreadsheetML extensions
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
    # newer DrawingML, several of the prefixes invented by docx4j
    "a13cmd": "http://schemas.microsoft.com/office/drawing/2013/main/command",
    "c16ac": "http://schemas.microsoft.com/office/drawing/2014/chart/ac",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "c16": "http://schemas.microsoft.com/office/drawing/2014/chart",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "dgm1611": "http://schemas.microsoft.com/office/drawing/2016/11/diagram",
    "a1611": "http://schemas.microsoft.com/office/drawing/2016/11/main",
    "dgm1612": "http://schemas.microsoft.com/office/drawing/2016/12/diagram",
    "ink16": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "a16svg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "c173": "http://schemas.microsoft.com/office/drawing/2017/03/chart",
    "adec": "http://schemas.microsoft.com/office/drawing/2017/decorative",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "an18": "http://schemas.microsoft.com/office/drawing/2018/animation",
    "anam3d": "http://schemas.microsoft.com/office/drawing/2018/animation/model3d",
    "a18hc": "http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor",
    # ink, maths, themes, web extensions
    "msink": "http://schemas.microsoft.com/ink/2010/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "thm15": "http://schemas.microsoft.com/office/thememl/2012/main",
    "wetp": "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11",
    "we": "http://schemas.microsoft.com/office/webextensions/webextension/2010/11",
    "ds": "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "cppr": "http://schemas.microsoft.com/office/2006/coverPageProps",
    "dssi": "http://schemas.microsoft.com/office/2006/digsig",
    # W3C and markup compatibility
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xs": "http://www.w3.org/2001/XMLSchema",
    "xf": "http://www.w3.org/2002/xforms",
    "xe": "http://www.w3.org/2001/xml-events",
    "xml": "http://www.w3.org/XML/1998/namespace",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "xd": "http://uri.etsi.org/01903/v1.3.2#",
    # Three namespaces this repository generates classes for that docx4j's
    # mapper has no entry for, so JAXB would invent `ns0` for them. The
    # prefixes are the conventional ones, and they are ours rather than
    # docx4j's; `tests/test_runtime.py` keeps the list from growing silently.
    "mml": "http://www.w3.org/1998/Math/MathML",
    "inkml": "http://www.w3.org/2003/InkML",
    "st": "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    # and one more, added with the docProps bindings of CR-002 section 9: the
    # DCMI type vocabulary, which `dcterms.xsd` imports and `docProps/core.xml`
    # can therefore name. docx4j's mapper has no entry for it either; `dcmitype`
    # is the prefix Dublin Core's own examples use.
    "dcmitype": "http://purl.org/dc/dcmitype/",
    # OpenDoPE
    "odx": "http://opendope.org/xpaths",
    "odc": "http://opendope.org/conditions",
    "odi": "http://opendope.org/components",
    "odq": "http://opendope.org/questions",
    "oda": "http://opendope.org/answers",
    "odgm": "http://opendope.org/SmartArt/DataHierarchy",
}

#: URI -> prefix; the direction ``getPreferredPrefixStatic`` answers in.
NAMESPACE_TO_PREFIX: dict[str, str] = {uri: prefix for prefix, uri in PREFIXES.items()}

#: The WordprocessingML namespace, docx4j's ``Namespaces.NS_WORD12``.
WML_NS = PREFIXES["w"]

#: The reserved XML namespace, which carries ``xml:space``.
XML_NS = PREFIXES["xml"]

#: docx4j's ``Namespaces.W_NAMESPACE_DECLARATION``.
W_NAMESPACE_DECLARATION = f'xmlns:w="{WML_NS}"'

#: The prefixes a part root may name in ``mc:Ignorable``. Everything that is an
#: Office *extension* to a base ECMA-376 namespace: a consumer that does not
#: know them must be able to ignore them, and a producer that names them must
#: declare them. Word's own ``document.xml`` writes a subset of these
#: (``w14 wp14`` through ``w14 w15 w16se w16cid wp14`` over the corpus in
#: ``samples/``).
MC_IGNORABLE_PREFIXES: frozenset[str] = frozenset(
    {
        "w14",
        "w15",
        "w16",
        "w16cex",
        "w16cid",
        "w16du",
        "w16sdtdh",
        "w16sdtfl",
        "w16se",
        "wp14",
        "wp15",
        "wpc",
        "wpg",
        "wpi",
        "wps",
        "wne",
        "a14",
        "a15",
        "a16",
        "a1611",
        "a16svg",
        "a18hc",
        "adec",
        "am3d",
        "an18",
        "anam3d",
        "c14",
        "c15",
        "c16",
        "c16ac",
        "c173",
        "cdr14",
        "cx",
        "cx1",
        "cx2",
        "cx3",
        "cx4",
        "cx5",
        "cx6",
        "cx7",
        "cx8",
        "dgm14",
        "dgm1611",
        "dgm1612",
        "dsp",
        "msink",
        "pic14",
        "x14",
        "x14ac",
        "x15",
        "x15ac",
        "xdr14",
    }
)


def __getattr__(name: str) -> object:
    """Resolve :data:`UNDERSTOOD`, which the generator writes, on first use.

    CR-002 section 5.6: the namespaces *this build of the model* understands,
    prefix -> URI, which is what ``mc:Choice``'s ``Requires`` is tested
    against. It is derived at generation time from the set of generated
    packages (``codegen/generate_el.py`` writes it into
    ``docx4j_py.el_index``) rather than hand kept here, so that adding a schema
    to ``codegen/generate.sh`` widens the MCE resolver by itself. Read it from
    here: ``docx4j_py.namespaces.UNDERSTOOD`` is the public name.

    It is *not* the same set as :data:`PREFIXES`, which is docx4j's whole
    preferred-prefix table and names plenty of namespaces this model has no
    classes for (``wpi``, the PML and SML extensions, VML): saying a prefix is
    understood when nothing can parse it is exactly the mistake
    ``mc:AlternateContent`` exists to prevent.
    """
    if name == "UNDERSTOOD":
        from docx4j_py.el_index import UNDERSTOOD

        globals()["UNDERSTOOD"] = UNDERSTOOD
        return UNDERSTOOD
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def ns_map(*, default: str | None = None) -> dict[str | None, str]:
    """Return an ``ns_map`` for the serialiser: prefix -> URI.

    ``xml`` is left out because it is bound by definition and a serialiser that
    declares it writes an invalid document.

    Args:
        default: a URI to bind to the default (prefix-less) namespace as well.
    """
    out: dict[str | None, str] = {p: u for p, u in PREFIXES.items() if p != "xml"}
    if default is not None:
        out[None] = default
    return out


def declarations(prefixes: object = None) -> str:
    """Return ``xmlns:...="..."`` declarations, as docx4j writes them inline.

    docx4j's ``W_NAMESPACE_DECLARATION`` is the one-prefix version of this; the
    fragment parser needs the whole table so that an author can write
    ``<w:p><w:r><w:t>...`` with ``w14:paraId``, ``r:id`` and
    ``mc:AlternateContent`` in it and declare nothing.

    Args:
        prefixes: an iterable of prefixes to restrict the output to; every
            prefix of the table except ``xml`` by default.
    """
    if prefixes is None:
        wanted = [p for p in PREFIXES if p != "xml"]
    else:
        wanted = [p for p in prefixes if p != "xml"]  # type: ignore[union-attr]
    return " ".join(f'xmlns:{p}="{PREFIXES[p]}"' for p in wanted)


def qname(prefix: str, local: str) -> str:
    """Return the James Clark qualified name ``{uri}local`` for a prefix."""
    return f"{{{PREFIXES[prefix]}}}{local}"


def split_qname(name: str) -> tuple[str | None, str]:
    """Split ``{uri}local`` into ``(uri, local)``; ``(None, name)`` if bare."""
    if name.startswith("{"):
        uri, _, local = name[1:].partition("}")
        return uri, local
    return None, name
