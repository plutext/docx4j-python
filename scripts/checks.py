#!/usr/bin/env python
"""Targeted fidelity checks (a)-(e) over the artefacts produced by roundtrip.py.

Run roundtrip.py first, then:

    python scripts/checks.py out/tuned
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import canon  # noqa: E402
from lxml import etree  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def signature(root, tags) -> list[str]:
    """Ordered list of child tags for every element in `tags`."""
    out = []
    for el in root.iter():
        if el.tag in tags:
            out.append("|".join(
                c.tag for c in el if isinstance(c.tag, str)))
    return out


def ns_used(root, prefix_uri: str) -> int:
    n = 0
    for el in root.iter():
        if isinstance(el.tag, str) and el.tag.startswith(prefix_uri):
            n += 1
        for a in el.attrib:
            if a.startswith(prefix_uri):
                n += 1
    return n


def texts(root) -> list[str]:
    return [el.text or "" for el in root.iter(W + "t")]


def space_attrs(root) -> list[str]:
    return [el.get(XMLSPACE) or "" for el in root.iter(W + "t")]


def main(out_root: str) -> int:
    rows = []
    for diff in sorted(Path(out_root).rglob("diff.txt")):
        d = diff.parent
        o = (d / "orig.xml").read_bytes()
        n = (d / "rt.xml").read_bytes()
        a = canon.normalize_tree(o)
        b = canon.normalize_tree(n)
        raw_a = canon.parse(o).getroot()
        raw_b = canon.parse(n).getroot()

        # (a) run order in w:p / w:body
        sig_a = signature(a, {W + "p", W + "body", W + "hdr", W + "ftr"})
        sig_b = signature(b, {W + "p", W + "body", W + "hdr", W + "ftr"})
        order_ok = sig_a == sig_b

        # (b) w14 / w15
        w14 = (ns_used(a, W14), ns_used(b, W14))
        w15 = (ns_used(a, W15), ns_used(b, W15))

        # (c) mc:AlternateContent
        ac_a = len(list(a.iter(MC + "AlternateContent")))
        ac_b = len(list(b.iter(MC + "AlternateContent")))
        fb_a = len(list(a.iter(MC + "Fallback")))
        fb_b = len(list(b.iter(MC + "Fallback")))

        # (d) xml:space + text
        sp_a, sp_b = space_attrs(raw_a), space_attrs(raw_b)
        tx_a, tx_b = texts(raw_a), texts(raw_b)

        # (e) mc:Ignorable vs declared prefixes
        ig_a = (raw_a.get(MC + "Ignorable") or "").split()
        ig_b = (raw_b.get(MC + "Ignorable") or "").split()
        pre_a = set(canon.prefix_map(o)) - {""}
        pre_b = set(canon.prefix_map(n)) - {""}

        rows.append(dict(
            part=str(d.relative_to(out_root)),
            order_ok=order_ok,
            w14=w14, w15=w15,
            ac=(ac_a, ac_b), fb=(fb_a, fb_b),
            space_ok=sp_a == sp_b, space_n=sum(1 for x in sp_a if x),
            text_ok=tx_a == tx_b, text_n=len(tx_a),
            ig_a=ig_a, ig_b=ig_b,
            pre_extra=sorted(pre_b - pre_a), pre_missing=sorted(pre_a - pre_b),
        ))

    def bad(r, k):
        return not r[k]

    print("## (a) element order inside w:body / w:p / w:hdr / w:ftr")
    bad_order = [r for r in rows if not r["order_ok"]]
    print(f"   {len(rows) - len(bad_order)}/{len(rows)} parts preserve child order exactly")
    for r in bad_order:
        print("   FAIL", r["part"])

    print("\n## (b) w14 / w15 node+attribute counts (orig -> round-trip)")
    for r in rows:
        if r["w14"][0] or r["w15"][0]:
            flag = "OK " if r["w14"][0] == r["w14"][1] and r["w15"][0] == r["w15"][1] else "LOSS"
            print(f"   {flag} {r['part']:52s} w14 {r['w14'][0]}->{r['w14'][1]}  w15 {r['w15'][0]}->{r['w15'][1]}")

    print("\n## (c) mc:AlternateContent / mc:Fallback counts")
    any_ac = False
    for r in rows:
        if r["ac"][0]:
            any_ac = True
            flag = "OK " if r["ac"][0] == r["ac"][1] and r["fb"][0] == r["fb"][1] else "LOSS"
            print(f"   {flag} {r['part']:52s} AC {r['ac'][0]}->{r['ac'][1]}  Fallback {r['fb'][0]}->{r['fb'][1]}")
    if not any_ac:
        print("   (no mc:AlternateContent in the sampled parts)")

    print("\n## (d) w:t text + xml:space")
    bad_t = [r for r in rows if not r["text_ok"]]
    bad_s = [r for r in rows if not r["space_ok"]]
    tot_sp = sum(r["space_n"] for r in rows)
    tot_t = sum(r["text_n"] for r in rows)
    print(f"   {tot_t} w:t nodes compared; text mismatches in {len(bad_t)} parts")
    print(f"   {tot_sp} xml:space attributes present; mismatches in {len(bad_s)} parts")
    for r in bad_t + bad_s:
        print("   FAIL", r["part"])

    print("\n## (e) mc:Ignorable and prefix declarations")
    for r in rows:
        if r["ig_a"] or r["ig_b"] or r["pre_extra"] or r["pre_missing"]:
            print(f"   {r['part']}")
            print(f"      Ignorable  {r['ig_a']} -> {r['ig_b']}")
            if r["pre_extra"]:
                print(f"      prefixes added by serializer: {r['pre_extra']}")
            if r["pre_missing"]:
                print(f"      prefixes lost:                {r['pre_missing']}")
            undeclared = [p for p in r["ig_b"] if p not in r["pre_extra"] and p not in
                          (set(canon.prefix_map((Path(out_root) / r['part'] / 'rt.xml').read_bytes())))]
            if undeclared:
                print(f"      !! Ignorable names undeclared prefixes: {undeclared}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1 else "out/tuned"))
