#!/usr/bin/env python
"""Round-trip real .docx parts through the xsdata-generated WML bindings.

    python scripts/roundtrip.py                 # lenient (default)
    python scripts/roundtrip.py --strict        # fail_on_unknown_* = True
    python scripts/roundtrip.py --only tables.docx

Writes per-part artefacts to out/<docx>/<part>/{orig.xml,rt.xml,diff.txt}
and prints a markdown results table.
"""

from __future__ import annotations

import argparse
import inspect
import json
import os
import sys
import statistics
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon  # noqa: E402

# --------------------------------------------------------------------------
# namespace prefixes Word itself uses in document.xml
# --------------------------------------------------------------------------
NS_MAP = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "xml": "http://www.w3.org/XML/1998/namespace",
}


MISSING_REQUIRED: list[str] = []


def tolerant_class_factory(clazz, params):
    """Fill in required fields the source document omitted.

    xsdata generates ``minOccurs>=1`` elements / ``use="required"`` attributes
    as dataclass fields with no default, so a schema-invalid-but-real docx
    (e.g. a ``w:tbl`` without ``w:tblGrid``) raises TypeError instead of
    degrading.  ParserConfig has no switch for this; a class_factory is the
    only hook.
    """
    try:
        return clazz(**params)
    except TypeError as exc:
        import dataclasses
        added = []
        for f in dataclasses.fields(clazz):
            if f.name in params:
                continue
            if f.default is not dataclasses.MISSING or \
               f.default_factory is not dataclasses.MISSING:
                continue
            params[f.name] = None
            added.append(f.name)
        if not added:
            raise
        MISSING_REQUIRED.append(f"{clazz.__name__}: {', '.join(added)}")
        return clazz(**params)


# part -> (class name, module suffix under the bindings package).  The
# suffix only applies to the default --structure-style clusters layout; with
# --models-module every class comes out of that one module instead.
PART_CLASSES = {
    "word/document.xml": ("Document", "document"),
    "word/styles.xml": ("Styles", "styles"),
    "word/numbering.xml": ("Numbering", "numbering"),
    "word/fontTable.xml": ("Fonts", "fonts"),
    "word/comments.xml": ("Comments", "comments"),
    "__header__": ("Hdr", "hdr"),
    "__footer__": ("Ftr", "ftr"),
}


BASELINE = ROOT / "baseline"


def load_models(package: str = "docx4j_py.generated", models_module: str | None = None):
    """Import the part root classes from the generated bindings.

    Args:
        package: the bindings package, one module per cluster under it
        models_module: a single module holding every class instead, which is
            what --structure-style namespaces or single-package produces
    """
    import importlib

    # The pre-Phase-B baseline lives under baseline/docx4j_py/generated (the
    # generator owns every file in docx4j_py/), and its 1,625 modules import
    # each other as docx4j_py.generated.<module>. Putting baseline/ in front
    # of the repository root shadows the Phase B package with the overlay,
    # which is what the baseline needs and all it needs. Run it with .venv,
    # which has upstream `xsdata`; .venv-fork does not.
    if (models_module or package).startswith("docx4j_py.generated"):
        sys.path.insert(0, str(BASELINE))

    models = {}
    for part, (class_name, suffix) in PART_CLASSES.items():
        source = models_module or f"{package}.{suffix}"
        models[part] = getattr(importlib.import_module(source), class_name)
    return models


def load_runtime(name: str = "xsdata"):
    """Import the xsdata runtime the generated bindings were built against."""
    import importlib

    mods = {}
    for key in (
        "formats.dataclass.context",
        "formats.dataclass.parsers",
        "formats.dataclass.parsers.config",
        "formats.dataclass.parsers.handlers",
        "formats.dataclass.parsers.skipped",   # fork only
        "formats.dataclass.serializers",
        "formats.dataclass.serializers.config",
        "formats.dataclass.serializers.writers",
    ):
        try:
            mods[key] = importlib.import_module(f"{name}.{key}")
        except ModuleNotFoundError:
            pass
    return mods


def part_class(name: str, models) -> type | None:
    if name in models:
        return models[name]
    base = os.path.basename(name)
    if base.startswith("header") and base.endswith(".xml"):
        return models["__header__"]
    if base.startswith("footer") and base.endswith(".xml"):
        return models["__footer__"]
    return None


@dataclass
class PartResult:
    docx: str
    part: str
    bytes_in: int
    parse_ms: float = 0.0
    ser_ms: float = 0.0
    ok: bool = False
    error: str = ""
    identical: bool = False
    c14n_identical: bool = False
    counts: dict = field(default_factory=dict)
    samples: list = field(default_factory=list)
    strict_error: str = ""
    orig_prefixes: dict = field(default_factory=dict)
    rt_prefixes: dict = field(default_factory=dict)
    ignorable_orig: str = ""
    ignorable_rt: str = ""
    skipped: list = field(default_factory=list)


def run_part(docx: str, part: str, data: bytes, klass, out_dir: Path,
             parser, serializer, strict_parser, nsmap_from_source=False,
             repeat: int = 3) -> PartResult:
    res = PartResult(docx=docx, part=part, bytes_in=len(data))

    # --- strict pass (informational) ---
    if strict_parser is not None:
        try:
            strict_parser.from_bytes(data, klass)
        except Exception as exc:  # noqa: BLE001
            res.strict_error = f"{type(exc).__name__}: {exc}"[:400]

    # --- lenient parse ---
    try:
        obj = parser.from_bytes(data, klass)  # warm the XmlContext metadata
        samples = []
        for _ in range(repeat):
            t = time.perf_counter()
            obj = parser.from_bytes(data, klass)
            samples.append((time.perf_counter() - t) * 1000)
        res.parse_ms = statistics.median(samples)
    except Exception as exc:  # noqa: BLE001
        res.error = f"parse: {type(exc).__name__}: {exc}"[:500]
        (out_dir / "error.txt").write_text(traceback.format_exc())
        return res

    # --- skipped content (fork ParserConfig(skipped_report=True)) ---
    # parser.skipped is cleared at the start of every parse, so this is
    # the report of the last of the timed passes, one part's worth.
    report = getattr(parser, "skipped", None)
    if report is not None:
        res.skipped = [(n.kind, n.path, n.parent_class or "") for n in report]

    # --- serialize ---
    try:
        t = time.perf_counter()
        if nsmap_from_source:
            src = {k: v for k, v in canon.parse(data).getroot().nsmap.items()
                   if k}
            ns_map = src or dict(NS_MAP)
        else:
            ns_map = dict(NS_MAP)
        out = serializer.render(obj, ns_map=ns_map)
        res.ser_ms = (time.perf_counter() - t) * 1000
    except Exception as exc:  # noqa: BLE001
        res.error = f"serialize: {type(exc).__name__}: {exc}"[:500]
        (out_dir / "error.txt").write_text(traceback.format_exc())
        return res

    res.ok = True
    new = out.encode("utf-8")
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "orig.xml").write_bytes(data)
    (out_dir / "rt.xml").write_bytes(new)

    try:
        res.c14n_identical = canon.c14n(data) == canon.c14n(new)
    except Exception:  # noqa: BLE001
        res.c14n_identical = False

    try:
        res.orig_prefixes = canon.prefix_map(data)
        res.rt_prefixes = canon.prefix_map(new)
        from lxml import etree as _et
        MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
        res.ignorable_orig = canon.parse(data).getroot().get(MC) or ""
        res.ignorable_rt = canon.parse(new).getroot().get(MC) or ""
    except Exception:  # noqa: BLE001
        pass

    rep = canon.compare(data, new)
    res.identical = rep.identical
    res.counts = dict(rep.counts)
    res.samples = [(d.category, d.path[-110:], d.detail[:160]) for d in rep.diffs[:60]]

    with (out_dir / "diff.txt").open("w") as fh:
        fh.write(f"# {docx} :: {part}\n")
        fh.write(f"identical(normalised) = {rep.identical}\n")
        fh.write(f"identical(c14n)       = {res.c14n_identical}\n")
        fh.write(f"mc:Ignorable orig = {res.ignorable_orig!r}\n")
        fh.write(f"mc:Ignorable rt   = {res.ignorable_rt!r}\n")
        fh.write(f"prefixes orig = {sorted(res.orig_prefixes)}\n")
        fh.write(f"prefixes rt   = {sorted(res.rt_prefixes)}\n")
        fh.write(f"skipped       = {len(res.skipped)}\n")
        for kind, path, owner in res.skipped:
            fh.write(f"    SKIPPED {kind} {path} (in {owner})\n")
        fh.write("\n")
        for k, v in sorted(rep.counts.items(), key=lambda kv: -kv[1]):
            fh.write(f"{v:6d}  {k}\n")
        fh.write("\n--- first 400 diffs ---\n")
        for d in rep.diffs:
            fh.write(f"[{d.category}] {d.path}\n        {d.detail}\n")
    return res


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default=None)
    ap.add_argument("--strict", action="store_true",
                    help="also run a strict pass and record the first failure")
    ap.add_argument("--samples", default=str(ROOT / "samples"))
    ap.add_argument("--out", default=str(ROOT / "out"))
    ap.add_argument("--json", default=None)
    ap.add_argument("--repeat", type=int, default=3,
                    help="median of N timed parse/serialize passes (after a warm-up)")
    ap.add_argument("--ignore-defaults", action="store_true",
                    help="SerializerConfig(ignore_default_attributes=True)")
    ap.add_argument("--nsmap-from-source", action="store_true",
                    help="reuse the prefix bindings declared on the source root")
    ap.add_argument("--tolerant-factory", action="store_true",
                    help="class_factory that supplies None for missing required fields")
    ap.add_argument("--package", default="docx4j_py.generated",
                    help="generated bindings package (clusters layout)")
    ap.add_argument("--models-module", default=None,
                    help="single module holding every part root class, for the "
                         "namespaces/single-package layouts")
    ap.add_argument("--runtime", default="xsdata",
                    help="runtime package the bindings import, xsdata or "
                         "docx4j_xsdata")
    ap.add_argument("--bool-numeric", dest="bool_numeric",
                    action=argparse.BooleanOptionalAction, default=None,
                    help="SerializerConfig(bool_format=\"numeric\"), Word's "
                         "1/0 rather than true/false; the default is on for "
                         "the docx4j_xsdata runtime, which is the only one "
                         "that has the option")
    ap.add_argument("--allow-skipped", action="store_true",
                    help="do not fail when the parser skipped content")
    args = ap.parse_args()

    if args.bool_numeric is None:
        args.bool_numeric = args.runtime == "docx4j_xsdata"

    t0 = time.perf_counter()
    models = load_models(args.package, args.models_module)
    import_ms = (time.perf_counter() - t0) * 1000

    rt = load_runtime(args.runtime)
    XmlContext = rt["formats.dataclass.context"].XmlContext
    XmlParser = rt["formats.dataclass.parsers"].XmlParser
    ParserConfig = rt["formats.dataclass.parsers.config"].ParserConfig
    LxmlEventHandler = rt["formats.dataclass.parsers.handlers"].LxmlEventHandler
    XmlSerializer = rt["formats.dataclass.serializers"].XmlSerializer
    SerializerConfig = rt["formats.dataclass.serializers.config"].SerializerConfig
    LxmlEventWriter = rt["formats.dataclass.serializers.writers"].LxmlEventWriter

    ctx = XmlContext()
    lenient_kw = dict(fail_on_unknown_properties=False,
                      fail_on_unknown_attributes=False,
                      fail_on_converter_warnings=False)
    if args.tolerant_factory:
        lenient_kw["class_factory"] = tolerant_class_factory
    if "skipped_report" in inspect.signature(ParserConfig).parameters:
        # fork only: collect what lenient parsing dropped, CR-001 section 7.
        # log=False, the harness prints and fails on them itself.
        SkippedReport = rt["formats.dataclass.parsers.skipped"].SkippedReport
        lenient_kw["skipped_report"] = SkippedReport(log=False)
    lenient = ParserConfig(**lenient_kw)
    parser = XmlParser(config=lenient, context=ctx, handler=LxmlEventHandler)

    strict_parser = None
    if args.strict:
        strict_cfg = ParserConfig(fail_on_unknown_properties=True,
                                  fail_on_unknown_attributes=True,
                                  fail_on_converter_warnings=True)
        strict_parser = XmlParser(config=strict_cfg, context=ctx,
                                  handler=LxmlEventHandler)

    ser_kw = dict(xml_declaration=True, indent=None,
                  ignore_default_attributes=args.ignore_defaults)
    if args.bool_numeric:
        if "bool_format" not in inspect.signature(SerializerConfig).parameters:
            raise SystemExit(f"--bool-numeric needs the fork runtime, "
                             f"{args.runtime} has no SerializerConfig.bool_format")
        ser_kw["bool_format"] = "numeric"
    ser_cfg = SerializerConfig(**ser_kw)
    serializer = XmlSerializer(config=ser_cfg, context=ctx,
                               writer=LxmlEventWriter)

    out_root = Path(args.out)
    results: list[PartResult] = []

    docs = sorted(Path(args.samples).glob("*.docx"))
    if args.only:
        docs = [d for d in docs if args.only in d.name]

    for docx in docs:
        with zipfile.ZipFile(docx) as z:
            names = [n for n in z.namelist()
                     if n.startswith("word/") and n.endswith(".xml")]
            for name in sorted(names):
                klass = part_class(name, models)
                if klass is None:
                    continue
                data = z.read(name)
                od = out_root / docx.stem / name.replace("/", "_")[:-4]
                od.mkdir(parents=True, exist_ok=True)
                r = run_part(docx.name, name, data, klass, od, parser,
                             serializer, strict_parser,
                             args.nsmap_from_source, args.repeat)
                results.append(r)
                flag = "OK " if r.ok else "ERR"
                same = "SAME" if r.identical else "diff"
                print(f"{flag} {same} {docx.name:38s} {name:26s} "
                      f"{r.bytes_in/1024:7.1f}K parse={r.parse_ms:7.1f}ms "
                      f"ser={r.ser_ms:7.1f}ms {r.error}")

    print(f"\nmodel import: {import_ms:.0f} ms")
    if MISSING_REQUIRED:
        from collections import Counter
        print("\n### required fields absent from the source (filled with None)")
        for k, v in Counter(MISSING_REQUIRED).most_common():
            print(f"  {v:5d}x {k}")

    if args.json:
        Path(args.json).write_text(json.dumps(
            [r.__dict__ for r in results], indent=1, default=str))

    # ---- markdown table -------------------------------------------------
    print("\n| docx | part | KB | parse ms | ser ms | canon identical | top diff categories |")
    print("|---|---|---:|---:|---:|---|---|")
    for r in results:
        cats = ", ".join(f"{k}×{v}" for k, v in
                         sorted(r.counts.items(), key=lambda kv: -kv[1])[:4])
        verdict = "yes" if r.identical else ("PARSE/SER FAIL" if not r.ok else "no")
        print(f"| {r.docx} | {r.part} | {r.bytes_in/1024:.0f} | "
              f"{r.parse_ms:.0f} | {r.ser_ms:.0f} | {verdict} | "
              f"{cats or r.error} |")

    # ---- aggregate categories ------------------------------------------
    agg: dict[str, int] = {}
    for r in results:
        for k, v in r.counts.items():
            agg[k] = agg.get(k, 0) + v
    print("\n### aggregate diff categories")
    for k, v in sorted(agg.items(), key=lambda kv: -kv[1]):
        print(f"  {v:7d}  {k}")

    if args.strict:
        print("\n### strict-mode (fail_on_unknown_*=True) first failures")
        for r in results:
            if r.strict_error:
                print(f"  {r.docx} :: {r.part}\n      {r.strict_error}")

    # ---- verdict --------------------------------------------------------
    identical = sum(1 for r in results if r.identical)
    failed = [r for r in results if not r.ok]
    skipped_parts = [r for r in results if r.skipped]

    print("\n### skipped content (lenient parsing dropped it)")
    if "skipped_report" not in inspect.signature(ParserConfig).parameters:
        print("  not available: this runtime has no "
              "ParserConfig(skipped_report=...)")
    elif not skipped_parts:
        print(f"  none, over {len(results)} parts")
    else:
        for r in skipped_parts:
            print(f"  {r.docx} :: {r.part}  {len(r.skipped)} skipped")
            for kind, path, owner in r.skipped[:20]:
                print(f"      SKIPPED {kind} {path} (in {owner})")

    total_skipped = sum(len(r.skipped) for r in results)
    real = {k: v for k, v in agg.items() if k not in canon.DiffReport.BENIGN}
    benign = {k: v for k, v in agg.items() if k in canon.DiffReport.BENIGN}
    print(f"\n### verdict\n  parts                    {len(results)}"
          f"\n  canonically identical    {identical}"
          f"\n  parse/serialise failures {len(failed)}"
          f"\n  differences              {sum(real.values())}"
          f"\n  benign respellings       {sum(benign.values())}"
          f" {sorted(benign) if benign else ''}"
          f"\n  skipped content          {total_skipped}")

    problems = []
    if failed:
        problems.append(f"{len(failed)} part(s) failed to parse or serialise")
    if real:
        problems.append("%d difference(s): %s" % (
            sum(real.values()),
            ", ".join(f"{k}={v}" for k, v in sorted(real.items()))))
    if total_skipped and not args.allow_skipped:
        problems.append(f"{total_skipped} skipped node(s) in "
                        f"{len(skipped_parts)} part(s)")
    if problems:
        print("\nROUND TRIP NOT CLEAN: " + "; ".join(problems))
        return 1
    print("\nROUND TRIP CLEAN: "
          f"{identical}/{len(results)} parts canonically identical, "
          f"0 differences, 0 skipped"
          + (f" ({sum(benign.values())} benign respellings)" if benign else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
