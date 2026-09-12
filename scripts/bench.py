#!/usr/bin/env python
"""Import-time and throughput benchmark for the **baseline** bindings.

This is the REPORT.md experiment's benchmark: unmodified xsdata 26.2 and the
`clusters` layout, which now lives in `baseline/docx4j_py/generated/` and runs
with `.venv`. Putting `baseline/` in front of the repository root shadows the
Phase B package with the overlay, as `scripts/roundtrip.py` does.

    .venv/bin/python scripts/bench.py

The Phase B package's own figures are guarded by `tests/test_performance.py`.
"""

from __future__ import annotations

import statistics
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "baseline"))


def main() -> int:
    t = time.perf_counter()
    from docx4j_py.generated.document import Document  # noqa: F401
    imp = (time.perf_counter() - t) * 1000
    print(f"import docx4j_py.generated.document : {imp:.0f} ms "
          f"({len(sys.modules)} modules resident)")

    from lxml import etree
    from xsdata.formats.dataclass.context import XmlContext
    from xsdata.formats.dataclass.parsers import XmlParser
    from xsdata.formats.dataclass.parsers.config import ParserConfig
    from xsdata.formats.dataclass.parsers.handlers import (
        LxmlEventHandler,
        XmlEventHandler,
    )
    from xsdata.formats.dataclass.serializers import XmlSerializer
    from xsdata.formats.dataclass.serializers.config import SerializerConfig
    from xsdata.formats.dataclass.serializers.writers import (
        LxmlEventWriter,
        XmlEventWriter,
    )

    sys.path.insert(0, str(ROOT / "scripts"))
    from roundtrip import NS_MAP

    ctx = XmlContext()
    cfg = ParserConfig(fail_on_unknown_properties=False,
                       fail_on_unknown_attributes=False)
    data = zipfile.ZipFile(ROOT / "samples" / "Symbols.docx").read("word/document.xml")
    print(f"\nsample: Symbols.docx word/document.xml, {len(data)/1024:.0f} KiB")

    # baseline: raw lxml
    ts = []
    for _ in range(5):
        t = time.perf_counter()
        etree.fromstring(data)
        ts.append((time.perf_counter() - t) * 1000)
    print(f"  lxml etree.fromstring        {statistics.median(ts):8.1f} ms "
          f"({len(data)/1024/1024/(statistics.median(ts)/1000):6.1f} MiB/s)")

    for name, handler in (("LxmlEventHandler", LxmlEventHandler),
                          ("XmlEventHandler(std)", XmlEventHandler)):
        p = XmlParser(config=cfg, context=ctx, handler=handler)
        p.from_bytes(data, Document)  # warm
        ts = []
        for _ in range(5):
            t = time.perf_counter()
            obj = p.from_bytes(data, Document)
            ts.append((time.perf_counter() - t) * 1000)
        print(f"  xsdata parse {name:18s}{statistics.median(ts):8.1f} ms "
              f"({len(data)/1024/1024/(statistics.median(ts)/1000):6.1f} MiB/s)")

    p = XmlParser(config=cfg, context=ctx, handler=LxmlEventHandler)
    obj = p.from_bytes(data, Document)
    for name, writer in (("LxmlEventWriter", LxmlEventWriter),
                         ("XmlEventWriter(std)", XmlEventWriter)):
        s = XmlSerializer(config=SerializerConfig(), context=ctx, writer=writer)
        s.render(obj, ns_map=dict(NS_MAP))  # warm
        ts = []
        for _ in range(5):
            t = time.perf_counter()
            out = s.render(obj, ns_map=dict(NS_MAP))
            ts.append((time.perf_counter() - t) * 1000)
        print(f"  xsdata ser   {name:18s}{statistics.median(ts):8.1f} ms "
              f"({len(out)/1024/1024/(statistics.median(ts)/1000):6.1f} MiB/s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
