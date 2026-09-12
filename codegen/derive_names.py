#!/usr/bin/env python3
"""Derive the docx4j name tables (schema name -> docx4j Java class name).

The tables are consumed by the xsdata fork's ``<ClassNames>`` generator option so
that the Python classes we generate from the OOXML schemas carry the same names
as the classes docx4j generated from the very same schemas with XJC.

Source of truth is the JAXB-annotated Java that docx4j's build produces:

    <docx4j>/docx4j-generated-objects/target/generated-sources/xjc

We only look at text; there is no Java parsing here, just regexes anchored at
column 0 so that nested/inner classes (which are always indented) are ignored.

Usage::

    python codegen/derive_names.py [--java-root PATH] [--schemas-root PATH]
                                   [--out DIR] [--report PATH]

Output: ``codegen/names/<prefix>.json`` -- one file per namespace, with exactly
four top level keys in this order: ``namespace``, ``java_package``, ``types``,
``elements``.  ``--report`` additionally dumps the anomaly data (as JSON) that
``codegen/names/README.md`` is written from.

Stdlib only.  Re-runnable: running it twice produces byte-identical output.
"""

from __future__ import annotations

import argparse
import builtins
import json
import keyword
import os
import re
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)

DEFAULT_JAVA_ROOT = (
    "/home/jharrop/git/docx4j/docx4j-generated-objects/target/generated-sources/xjc"
)
DEFAULT_SCHEMAS_ROOT = os.path.join(REPO, "schemas")
DEFAULT_OUT = os.path.join(HERE, "names")

SKIP_FILES = {"ObjectFactory.java", "package-info.java"}

# ---------------------------------------------------------------------------
# regexes
# ---------------------------------------------------------------------------

# ``^`` + re.M => column 0 only => top level class annotations only.  Every
# nested/inner class in the generated sources is indented, so this is how we
# keep e.g. ``P.Hyperlink``'s own @XmlRootElement out of the top level answer.
RE_TOP_XMLTYPE = re.compile(r'^@XmlType\s*\(([^)]*)\)', re.MULTILINE)
RE_TOP_XMLROOT = re.compile(r'^@XmlRootElement\s*\(([^)]*)\)', re.MULTILINE)
RE_TOP_XMLENUM = re.compile(r'^@XmlEnum\b', re.MULTILINE)
RE_ATTR = lambda key: re.compile(r'\b%s\s*=\s*"([^"]*)"' % key)
RE_NAME_ATTR = RE_ATTR("name")
RE_NS_ATTR = RE_ATTR("namespace")

RE_PKG_NAMESPACE = re.compile(
    r'@(?:jakarta|javax)\.xml\.bind\.annotation\.XmlSchema\s*\('
    r'[^)]*?namespace\s*=\s*"([^"]*)"',
    re.DOTALL,
)

# @XmlElementDecl(...) followed by the factory method.  Two shapes occur:
#   public JAXBElement<Foo> createXxx(Foo value)          (the usual one)
#   public CTR.TMath createCTRTMath(CTText value)         (JAXBElement subclass)
RE_ELEMENT_DECL = re.compile(
    r'@XmlElementDecl\s*\(([^)]*)\)\s*'
    r'public\s+(?:JAXBElement\s*<\s*([\w.$\[\]]+)\s*>|([\w.$\[\]]+))\s+\w+\s*\(',
)
RE_SCOPE = re.compile(r'\bscope\s*=\s*([\w.$]+)\.class')
RE_IMPORT = re.compile(r'^import\s+static\s+.*?;|^import\s+([\w.$]+)\s*;', re.MULTILINE)

RE_TARGET_NS = re.compile(r'targetNamespace\s*=\s*"([^"]*)"')

PY_BUILTINS = set(dir(builtins))


# ---------------------------------------------------------------------------
# file name prefixes
# ---------------------------------------------------------------------------

# Namespaces whose mechanical prefix would be unhelpful.
PREFIX_OVERRIDES = {
    "http://www.w3.org/1998/Math/MathML": "mathml",
    "http://www.w3.org/2003/InkML": "inkml",
    "http://www.w3.org/XML/1998/namespace": "xml",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes": "shared_types",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "relationships",
}

# Leading java package segments that carry no information.
PACKAGE_STRIPS = (
    "org.docx4j.com.microsoft.schemas.office.",
    "org.docx4j.com.microsoft.schemas.",
    "org.xlsx4j.schemas.microsoft.com.office.",
    "org.xlsx4j.schemas.microsoft.com.",
    "org.pptx4j.com.microsoft.schemas.office.",
    "org.pptx4j.com.microsoft.schemas.",
    "org.docx4j.",
    "org.xlsx4j.",
    "org.pptx4j.",
)

RE_CAMEL = re.compile(r'(?<=[a-z0-9])(?=[A-Z])')


def prefix_for(namespace: str, java_package: str) -> str:
    """Short, stable file name for a namespace."""
    if namespace in PREFIX_OVERRIDES:
        return PREFIX_OVERRIDES[namespace]
    tail = java_package
    for strip in PACKAGE_STRIPS:
        if tail.startswith(strip):
            tail = tail[len(strip):]
            break
    parts = []
    for seg in tail.split("."):
        if re.fullmatch(r'x\d+', seg):       # x2010 -> 2010
            seg = seg[1:]
        seg = RE_CAMEL.sub("_", seg)          # wordprocessingDrawing -> ..._drawing
        parts.append(seg.lower())
    return "_".join(p for p in parts if p)


# ---------------------------------------------------------------------------
# scanning
# ---------------------------------------------------------------------------

def read(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        return fh.read()


def schema_closure(schemas_root: str) -> set:
    out = set()
    for root, _dirs, files in os.walk(schemas_root):
        for name in sorted(files):
            if name.endswith(".xsd"):
                out.update(RE_TARGET_NS.findall(read(os.path.join(root, name))))
    return out


def find_packages(java_root: str):
    """-> (packages, skipped) where packages maps java_package -> info dict."""
    packages, skipped = {}, []
    for root, _dirs, files in os.walk(java_root):
        javas = sorted(f for f in files if f.endswith(".java"))
        if not javas:
            continue
        pkg = os.path.relpath(root, java_root).replace(os.sep, ".")
        if "package-info.java" not in files:
            skipped.append({"java_package": pkg, "dir": root,
                            "reason": "no package-info.java",
                            "java_files": len(javas)})
            continue
        m = RE_PKG_NAMESPACE.search(read(os.path.join(root, "package-info.java")))
        ns = m.group(1) if m else ""
        if not ns:
            skipped.append({"java_package": pkg, "dir": root,
                            "reason": "empty or absent @XmlSchema namespace",
                            "java_files": len(javas)})
            continue
        packages[pkg] = {"namespace": ns, "dir": root, "files": javas}
    return packages, skipped


def simple_name(java_type: str) -> str:
    """Drop the package qualifier, keep any outer-class qualifier.

    ``org.docx4j.wml.P.Hyperlink`` -> ``P.Hyperlink``; ``CTR.TMath`` -> unchanged.
    Package segments are the leading lowercase-initial ones.
    """
    segs = java_type.split(".")
    while len(segs) > 1 and segs[0][:1].islower():
        segs.pop(0)
    return ".".join(segs)


def scan_package(pkg: str, info: dict):
    """Scan one java package; return the raw (pre-resolution) observations."""
    ns = info["namespace"]
    classes = {f[:-len(".java")] for f in info["files"]} - {"package-info", "ObjectFactory"}

    types = defaultdict(set)       # schema type name -> {class}
    roots = defaultdict(set)       # element name -> {class}
    anonymous = []                 # classes with @XmlType(name = "")
    no_xmltype = []                # classes with no top level @XmlType at all
    foreign_type_ns = []           # top level @XmlType(namespace=<other>)
    foreign_root_ns = []           # top level @XmlRootElement(namespace=<other>)

    for fname in info["files"]:
        if fname in SKIP_FILES:
            continue
        cls = fname[:-len(".java")]
        src = read(os.path.join(info["dir"], fname))

        mt = RE_TOP_XMLTYPE.search(src)
        if mt is None:
            if RE_TOP_XMLENUM.search(src):
                no_xmltype.append(cls)
            else:
                no_xmltype.append(cls)
        else:
            body = mt.group(1)
            mn = RE_NAME_ATTR.search(body)
            tname = mn.group(1) if mn else None
            mns = RE_NS_ATTR.search(body)
            if mns and mns.group(1) != ns:
                foreign_type_ns.append({"class": cls, "type": tname,
                                        "namespace": mns.group(1)})
            if tname is None:
                no_xmltype.append(cls)
            elif tname == "":
                anonymous.append(cls)
            else:
                types[tname].add(cls)

        mr = RE_TOP_XMLROOT.search(src)
        if mr is not None:
            body = mr.group(1)
            mn = RE_NAME_ATTR.search(body)
            if mn and mn.group(1):
                mns = RE_NS_ATTR.search(body)
                if mns and mns.group(1) != ns:
                    foreign_root_ns.append({"class": cls, "element": mn.group(1),
                                            "namespace": mns.group(1)})
                roots[mn.group(1)].add(cls)

    # ObjectFactory element declarations, possibly for other namespaces.
    of_decls = []                  # {namespace, element, class, scope, owner, ...}
    of_path = os.path.join(info["dir"], "ObjectFactory.java")
    if os.path.exists(of_path):
        src = read(of_path)
        imports = {}
        for fq in RE_IMPORT.findall(src):
            if fq:
                imports[fq.rsplit(".", 1)[-1]] = fq.rsplit(".", 1)[0]
        seen = 0
        for m in RE_ELEMENT_DECL.finditer(src):
            seen += 1
            ann = m.group(1)
            raw = m.group(2) or m.group(3)
            mn = RE_NAME_ATTR.search(ann)
            mns = RE_NS_ATTR.search(ann)
            if not mn or not mn.group(1):
                continue
            ms = RE_SCOPE.search(ann)
            # Where does the payload class live?  Either it is written fully
            # qualified, or it is imported, or it is in this very package.
            segs = raw.split(".")
            owner = None
            if len(segs) > 1 and segs[0][:1].islower():
                qual = []
                while len(segs) > 1 and segs[0][:1].islower():
                    qual.append(segs.pop(0))
                owner = ".".join(qual)
            name = ".".join(segs)
            top = segs[0]
            if owner is None:
                owner = imports.get(top, pkg if top in classes else None)
            of_decls.append({
                "namespace": mns.group(1) if mns else "",
                "element": mn.group(1),
                "class": name,
                "raw": raw,
                "owner_package": owner,
                "scope": simple_name(ms.group(1)) if ms else "GLOBAL",
                "from_package": pkg,
            })
        declared = src.count("@XmlElementDecl")
        if seen != declared:
            print("warning: %s: matched %d of %d @XmlElementDecl"
                  % (of_path, seen, declared), file=sys.stderr)

    return {
        "classes": classes,
        "java_package": pkg,
        "namespace": ns,
        "types": types,
        "roots": roots,
        "of_decls": of_decls,
        "anonymous": sorted(anonymous),
        "no_xmltype": sorted(no_xmltype),
        "foreign_type_ns": foreign_type_ns,
        "foreign_root_ns": foreign_root_ns,
        "class_count": len(classes),
    }


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def scoped_entries(name: str, decls, by_class: dict) -> dict:
    """Turn ``@XmlElementDecl(scope=...)`` declarations into scoped keys.

    JAXB resolves an element name that several classes claim by the class the
    element appears in.  The fork's name table expresses the same thing as an
    ``element@CT_EnclosingType`` key, so the scope -- a Java class name -- has
    to be translated back into the schema type name that produced it, which is
    what ``by_class`` (class name -> type name) is for.  A scope with no schema
    type behind it (an inner Java class such as ``Comments.Comment``) and a
    value that is an inner Java class are both skipped: neither can be written
    as a key or a python class name.
    """
    out = {}
    for cls, scope, _owner in sorted(decls):
        tname = by_class.get(scope)
        if tname is None or "." in cls:
            continue
        out["%s@%s" % (name, tname)] = cls
    return out


def build(java_root: str, schemas_root: str):
    closure = schema_closure(schemas_root)
    packages, skipped = find_packages(java_root)

    scans = {pkg: scan_package(pkg, info) for pkg, info in sorted(packages.items())}

    ns_to_pkgs = defaultdict(list)
    for pkg, info in sorted(packages.items()):
        ns_to_pkgs[info["namespace"]].append(pkg)

    emitted_ns = sorted(closure & set(ns_to_pkgs))
    emitted = set(emitted_ns)

    # Every docx4j class we know about, so an @XmlElementDecl payload can be
    # told apart from a plain Java type (String, BigInteger, byte[], ...).
    known = {pkg: scan["classes"] for pkg, scan in scans.items()}

    # Route ObjectFactory element declarations to their owning namespace.
    of_by_ns = defaultdict(lambda: defaultdict(set))   # ns -> element -> {(cls,scope,pkg)}
    dropped_foreign_ns = defaultdict(int)              # namespace -> count
    non_class = defaultdict(int)                       # payload -> count
    cross_pkg = defaultdict(set)                       # ns -> {(element, class, pkg)}
    for scan in scans.values():
        for rec in scan["of_decls"]:
            target = rec["namespace"] or scan["namespace"]
            owner, top = rec["owner_package"], rec["class"].split(".")[0]
            if owner is None or top not in known.get(owner, ()):
                # Not a docx4j class: java.lang.String, BigInteger, byte[], ...
                non_class[rec["raw"]] += 1
                continue
            if target not in emitted:
                dropped_foreign_ns[target] += 1
                continue
            of_by_ns[target][rec["element"]].add(
                (rec["class"], rec["scope"], owner))
            if owner != ns_to_pkgs[target][0]:
                cross_pkg[target].add((rec["element"], rec["class"], owner))

    results, anomalies = {}, {
        "element_collisions": [],
        "type_collisions": [],
        "non_class_element_decls": dict(sorted(non_class.items())),
        "class_from_other_java_package": [],
        "dropped_foreign_namespace_decls": dict(sorted(dropped_foreign_ns.items())),
        "foreign_type_namespace": [],
        "foreign_root_namespace": [],
        "skipped_packages": skipped,
        "closure_without_package": sorted(closure - set(ns_to_pkgs)),
        "package_outside_closure": sorted(
            (ns, ns_to_pkgs[ns]) for ns in set(ns_to_pkgs) - closure),
        "multi_package_namespaces": {ns: p for ns, p in sorted(ns_to_pkgs.items())
                                     if len(p) > 1},
        "nested_class_values": [],
        "anonymous_types": {},
        "no_xmltype": {},
    }

    for ns in emitted_ns:
        pkg = ns_to_pkgs[ns][0]
        scan = scans[pkg]
        pfx = prefix_for(ns, pkg)
        if pfx in results:
            raise SystemExit("prefix collision: %r claimed by %s and %s"
                             % (pfx, results[pfx]["namespace"], ns))

        # types -----------------------------------------------------------
        types = {}
        for tname, classes in sorted(scan["types"].items()):
            if len(classes) == 1:
                types[tname] = sorted(classes)[0]
            else:
                anomalies["type_collisions"].append(
                    {"prefix": pfx, "namespace": ns, "type": tname,
                     "candidates": sorted(classes)})

        # elements --------------------------------------------------------
        # A class that ``types`` already names needs no element entry: the
        # generator reaches it by its schema type name.  So when an element
        # name has several candidate classes, drop the ones ``types`` covers
        # first; JAXB's ambiguity is usually only about which of several
        # already-named classes an element wraps.  What is left is either one
        # class, which resolves the collision, or nothing, which means the
        # element needs no entry at all, or a genuine ambiguity, which is
        # written out as scoped keys, ``element@CT_EnclosingType``, the fork's
        # analogue of JAXB's ``@XmlElementDecl(scope=...)``.
        by_class = {}
        for tname, cname in types.items():
            by_class.setdefault(cname, tname)

        def narrow(cands, named=by_class):
            return sorted(c for c in cands if c not in named)

        elements = {}
        names = set(scan["roots"]) | set(of_by_ns.get(ns, {}))
        for name in sorted(names):
            root_classes = scan["roots"].get(name, set())
            if len(root_classes) == 1:
                elements[name] = sorted(root_classes)[0]
                continue
            if len(root_classes) > 1:
                left = narrow(root_classes)
                record = {"prefix": pfx, "namespace": ns, "element": name,
                          "source": "@XmlRootElement",
                          "candidates": sorted(root_classes),
                          "not_named_by_types": left}
                if len(left) == 1:
                    elements[name] = left[0]
                    record["resolution"] = left[0]
                elif not left:
                    record["resolution"] = "every candidate is named by types"
                else:
                    record["resolution"] = "unresolved"
                anomalies["element_collisions"].append(record)
                continue
            decls = of_by_ns.get(ns, {}).get(name, set())
            distinct = sorted({d[0] for d in decls})
            if len(distinct) == 1:
                elements[name] = distinct[0]
            elif len(distinct) > 1:
                left = narrow(distinct)
                record = {"prefix": pfx, "namespace": ns, "element": name,
                          "source": "ObjectFactory @XmlElementDecl",
                          "candidates": distinct,
                          "not_named_by_types": left,
                          "scopes": sorted("%s (scope %s, %s)" % d for d in decls)}
                if len(left) == 1:
                    elements[name] = left[0]
                    record["resolution"] = left[0]
                elif not left:
                    record["resolution"] = "every candidate is named by types"
                else:
                    scoped = scoped_entries(name, decls, by_class)
                    elements.update(scoped)
                    record["resolution"] = ("scoped: " + ", ".join(sorted(scoped))
                                            if scoped else "unresolved")
                anomalies["element_collisions"].append(record)

        for key, val in sorted(elements.items()):
            if "." in val:
                anomalies["nested_class_values"].append(
                    {"prefix": pfx, "element": key, "class": val})

        if scan["anonymous"]:
            anomalies["anonymous_types"][pfx] = scan["anonymous"]
        if scan["no_xmltype"]:
            anomalies["no_xmltype"][pfx] = scan["no_xmltype"]
        anomalies["foreign_type_namespace"].extend(
            dict(d, prefix=pfx) for d in scan["foreign_type_ns"])
        anomalies["foreign_root_namespace"].extend(
            dict(d, prefix=pfx) for d in scan["foreign_root_ns"])
        for elem, cls, owner in sorted(cross_pkg.get(ns, ())):
            if elements.get(elem) == cls:
                anomalies["class_from_other_java_package"].append(
                    {"prefix": pfx, "element": elem, "class": cls,
                     "java_package": owner})

        results[pfx] = {
            "namespace": ns,
            "java_package": pkg,
            "types": types,
            "elements": elements,
        }

    # cross-namespace / python hazards ------------------------------------
    by_class = defaultdict(set)
    for pfx, doc in results.items():
        for val in set(doc["types"].values()) | set(doc["elements"].values()):
            by_class[val].add(pfx)
    anomalies["class_name_across_namespaces"] = {
        c: sorted(p) for c, p in sorted(by_class.items()) if len(p) > 1}

    ident = re.compile(r'^[A-Za-z_][A-Za-z0-9_]*$')
    bad_ident, py_kw, dupes = [], [], []
    bad_type_ident, elem_kw = [], []
    for pfx, doc in sorted(results.items()):
        for name in doc["types"]:
            if not ident.match(name):
                bad_type_ident.append({"prefix": pfx, "type": name,
                                       "class": doc["types"][name]})
        for name in doc["elements"]:
            if not ident.match(name):
                bad_ident.append({"prefix": pfx, "element": name,
                                  "class": doc["elements"][name]})
            if keyword.iskeyword(name) or name in PY_BUILTINS:
                elem_kw.append({"prefix": pfx, "element": name,
                                "class": doc["elements"][name],
                                "kind": "keyword" if keyword.iskeyword(name) else "builtin"})
        used = defaultdict(list)
        for key, val in list(doc["types"].items()):
            used[val].append("type:" + key)
        for key, val in list(doc["elements"].items()):
            used[val].append("element:" + key)
        for val, keys in sorted(used.items()):
            if keyword.iskeyword(val) or keyword.iskeyword(val.lower()) or val in PY_BUILTINS:
                py_kw.append({"prefix": pfx, "class": val, "keys": keys})
            tkeys = [k for k in keys if k.startswith("type:")]
            ekeys = [k for k in keys if k.startswith("element:")]
            # one type + one element for the same class is the normal shape
            # (CT_PPr/pPr -> PPr); more than that is a genuine many-to-one.
            if len(tkeys) > 1 or len(ekeys) > 1:
                dupes.append({"prefix": pfx, "class": val,
                              "types": sorted(tkeys), "elements": sorted(ekeys)})
    anomalies["element_names_not_python_identifiers"] = bad_ident
    anomalies["type_names_not_python_identifiers"] = bad_type_ident
    anomalies["element_names_keyword_or_builtin"] = elem_kw
    anomalies["class_names_keyword_or_builtin"] = py_kw
    anomalies["class_name_reused_within_namespace"] = dupes

    return results, anomalies


def write_json(path: str, doc: dict) -> None:
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(doc, fh, indent=2, ensure_ascii=False, sort_keys=False)
        fh.write("\n")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--java-root", default=DEFAULT_JAVA_ROOT,
                    help="docx4j generated-sources/xjc root")
    ap.add_argument("--schemas-root", default=DEFAULT_SCHEMAS_ROOT,
                    help="schema tree that defines the namespace closure")
    ap.add_argument("--out", default=DEFAULT_OUT, help="output directory")
    ap.add_argument("--report", default=None,
                    help="also write the anomaly report as JSON to this path")
    args = ap.parse_args(argv)

    results, anomalies = build(args.java_root, args.schemas_root)

    os.makedirs(args.out, exist_ok=True)
    for pfx in sorted(results):
        doc = results[pfx]
        write_json(os.path.join(args.out, pfx + ".json"), {
            "namespace": doc["namespace"],
            "java_package": doc["java_package"],
            "types": dict(sorted(doc["types"].items())),
            "elements": dict(sorted(doc["elements"].items())),
        })

    if args.report:
        write_json(args.report, anomalies)

    print("wrote %d name tables to %s" % (len(results), args.out))
    for pfx in sorted(results):
        print("  %-34s types=%-5d elements=%-5d %s"
              % (pfx, len(results[pfx]["types"]), len(results[pfx]["elements"]),
                 results[pfx]["namespace"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
