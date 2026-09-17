# Can xsdata bind WordprocessingML? — a feasibility experiment

**Question.** docx4j binds the ECMA-376 schemas to typed Java objects with JAXB.
Could a Python analogue use [xsdata](https://xsdata.readthedocs.io) the same way?
This repo generates Python dataclass bindings from docx4j's own `wml.xsd` tree and
round-trips 12 real Word documents (50 XML parts) through them, comparing the
re-serialized XML against the original.

**Short answer.** Yes, mostly — with two caveats that are *not* the ones usually
predicted. Content fidelity is much better than expected: **not one element was
dropped, reordered or corrupted in 50 parts**, mc:AlternateContent survives,
w14/w15 survive, `xml:space` and text survive. What actually bites is
(1) a ~0.9 s unavoidable import cost and ~25x slower parsing than raw lxml, and
(2) a class of *attribute*-level fidelity problems (schema defaults, `mc:Ignorable`
on types that don't declare it, boolean `1` -> `true`) plus a hard `TypeError` on
schema-invalid-but-real documents.

---

## 1. Environment and versions

| | |
|---|---|
| Python | **3.14.6** (system CPython, `/usr/bin/python3.14`) |
| xsdata | **26.2** |
| lxml | **6.1.3** |
| ruff | 0.16.7 (xsdata shells out to it to format the output — it must be on `PATH`) |
| uv | 0.12.3 |
| OS | Linux 6.6.151 x86_64, Manjaro |

xsdata 26.2 installs and runs fine on Python 3.14; no downgrade was needed.

```bash
uv venv --python 3.14 .venv
uv pip install --python .venv/bin/python 'xsdata[cli,lxml]==26.2'
export PATH="$PWD/.venv/bin:$PATH"   # xsdata invokes `ruff` by bare name
```

---

## 2. Choosing the entry schema

docx4j's JAXB build uses the `*__ROOT.xsd` aggregator files, but for WML that turns
out to be unnecessary:

* `xsd/wml/wml__ROOT.xsd` only imports `wml.xsd` + `w14_word_2010_wordml.xsd`, and
  `wml.xsd` *already* imports w14. Its own comment says *"no real need for this,
  but keep it for experimental purposes"*. It adds nothing.
* `xsd/wml/package.xsd` is the `pkg:` flat-OPC schema, plus the `jaxb:package`
  binding declarations (`org.docx4j.wml`) — irrelevant to xsdata, which derives
  package names from namespace URIs or from `--structure-style`.
* `xsd/bindings.xjb` contains exactly one customization —
  `<jaxb:globalBindings collectionType="org.docx4j.list.ArrayListDocx4j"/>` —
  i.e. docx4j swaps the list implementation so it can maintain parent pointers.
  **There is no xsdata equivalent** (see 7.4).

So the entry point used here is **`schemas/wml/wml.xsd`**. Its transitive closure is
**91 xsd files** (wml, dml, odrawxml, shared, mce, customXml, offmacro, inkml,
mathml). Those 91 files — and only those — were copied into `schemas/`, so the
experiment is self-contained and does not read `/home/jharrop/git/docx4j`.

### 2.1 The one schema change I had to make

Generation failed with:

```
Error: Missing inner class
parent: Class(qname='{...wordprocessingml/2006/main}EG_PContent', tag='Group', ...)
qname: {...wordprocessingml/2006/main}dir
```

Cause: `w:dir` and `w:bdo` are declared inside the **recursive** group
`EG_ContentRunContent` with *anonymous* `<xsd:complexType>`s that themselves
reference `EG_PContent`. xsdata 26.2's group-flattening loses the inner class.
This is an xsdata bug, not a schema error (XJC accepts it).

Minimal workaround, applied in `schemas/wml/wml.xsd` and marked with a
`docx4j-python patch:` comment — promote the two anonymous types to named ones:

```xml
<xsd:element name="dir" type="CT_Dir"/>
<xsd:element name="bdo" type="CT_Bdo"/>
...
<xsd:complexType name="CT_Dir">
  <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
  <xsd:attribute name="val" type="ST_Direction" use="optional"/>
</xsd:complexType>
<!-- same for CT_Bdo -->
```

Semantics are unchanged. That was the **only** schema edit required.

---

## 3. Generation

### 3.1 `--structure-style namespaces` does not work for OOXML

The suggested command uses `--structure-style namespaces`. It produces output, but
the output **cannot be imported**:

```
Error: Circular Dependencies Found
help: Try a different structure style and/or enable unnest classes.
...
ImportError: cannot import name 'CtEmpty' from partially initialized module
  'docx4j_py.generated.org.openxmlformats.schemas.wordprocessingml.pkg_2006.main'
  (most likely due to a circular import)
```

WML <-> DML <-> w14 <-> OMML are mutually recursive namespaces, so one module per
namespace is by construction an import cycle.

`--structure-style namespace-clusters` fails even harder, refusing to emit anything:

```
Error: Found strongly connected types from different namespaces
namespaces: {'...wordprocessingml/2006/main', '...officeDocument/2006/math'}
```

(w:r contains m:oMath, m:r contains w:rPr — a genuine cycle in ECMA-376.)

Two styles do work: **`clusters`** (one module per strongly-connected component,
1 625 modules) and **`single-package`** (one 64 109-line module). I kept `clusters`.

Also note: **`--kw-only` does not exist as a CLI flag in xsdata 26.2** — `kw_only=True`
is emitted unconditionally (`filters.py:156`). `--slots` does exist and is used.
`UnnestClasses` had to stay off (it triggered the same "Missing inner class" path).

### 3.2 Final config and result

`.xsdata.xml`:

```xml
<Output maxLineLength="99">
  <Package>docx4j_py.generated</Package>
  <Format slots="true">dataclasses</Format>
  <Structure>clusters</Structure>
  <CompoundFields defaultName="content" useSubstitutionGroups="true">true</CompoundFields>
  <UnnestClasses>false</UnnestClasses>
</Output>
```

```bash
.venv/bin/xsdata generate -c .xsdata.xml schemas/wml/wml.xsd
```

| metric | value |
|---|---|
| generation wall time | **4.3 s** (4.32 s / 4.43 s over two runs) |
| xsd files parsed | 91 |
| classes into the analyzer | 1 983 main + 164 inner |
| modules emitted | **1 625** |
| `@dataclass` classes | **1 373** |
| `Enum` classes | **319** |
| lines of Python | **78 878** |
| on-disk size | 15 MiB (`.py`), 3.6 MiB for the `single-package` variant |
| warnings | none (beyond the fatal errors above) |

Comparison: `--structure-style single-package` generates in 4.0 s -> one 64 109-line
module, 3.6 MiB.

Sanity check of the generated shape — `CT_Body` gets a proper compound field, so
document order is representable:

```python
@dataclass(slots=True, kw_only=True)
class CtBody:
    content: list[CtCustomXmlBlock | Sdt | P | CtTbl | ProofErr | ... ] = field(
        default_factory=list, metadata={"type": "Elements", "choices": (...)})
```

---

## 4. Import cost

`python scripts/bench.py`, importing `docx4j_py.generated.document` (which
transitively pulls the whole WML+DML graph — 1 755 modules resident):

| | clusters (1 625 modules) | single-package (1 module) |
|---|---:|---:|
| cold (no `__pycache__`) | **1 329 ms** | 855-916 ms |
| warm (bytecode cached) | **857-905 ms** | 693-730 ms |

Cold **time-to-first-parsed-document** (import + lazy `XmlContext` metadata build +
parse of `tables.docx` `document.xml`, 51 KiB) = **958 ms**, of which
877 ms import, 50 ms first parse, 21 ms for each subsequent parse of the same part.

Resident memory: **~46 MiB** after import (35 MiB attributable to the models),
~62 MiB after parsing a 772 KiB `document.xml`.

**Verdict: the import-time concern is real.** ~0.9 s and ~46 MiB before you touch a
document is fine for a server process and painful for a CLI. Roughly 25% can be
recovered by switching to `single-package`. The cost is dataclass *construction*
(1 373 `@dataclass` decorators running at import), not file I/O — so lazy
per-module importing would not help much either, because `document.py` reaches
essentially the whole graph.

---

## 5. Throughput

`Symbols.docx` `word/document.xml`, 772 KiB, median of 5 after warm-up:

| | time | throughput |
|---|---:|---:|
| `lxml.etree.fromstring` (baseline, no binding) | 10.9 ms | 69.5 MiB/s |
| xsdata parse, `LxmlEventHandler` | 250.7 ms | 3.0 MiB/s |
| xsdata parse, `XmlEventHandler` (stdlib) | **195.0 ms** | 3.9 MiB/s |
| xsdata serialize, `LxmlEventWriter` | 371.9 ms | 2.0 MiB/s |
| xsdata serialize, `XmlEventWriter` (stdlib) | 354.7 ms | 2.1 MiB/s |

**Verdict: the speed concern is real.** ~23x slower than raw lxml to parse, ~34x
to serialize. Interestingly the **lxml handler is slower than the stdlib one** for
parsing (250 ms vs 195 ms) — the lxml path builds an intermediate tree. Serializing
is *slower than parsing*, which is unusual and is the bigger practical problem for a
read-modify-write library.

For typical documents (a 51 KiB `document.xml`) it is 21 ms parse + 26 ms
serialize — perfectly usable. It is the 772 KiB outlier that costs 0.6 s.

---

## 6. Round-trip results

`scripts/roundtrip.py` unzips each sample, parses `word/document.xml`,
`styles.xml`, `numbering.xml`, `fontTable.xml`, `comments.xml` and every
`header*.xml` / `footer*.xml` with `XmlParser(handler=LxmlEventHandler)`, then
re-serializes with `XmlSerializer(writer=LxmlEventWriter)` and an explicit
`ns_map`. `scripts/canon.py` compares the two: attributes sorted,
prefixes collapsed to `{uri}local` (so a prefix rename is *not* counted as a
difference), whitespace-only text dropped **except** under `xml:space="preserve"`.
lxml exclusive C14N is also computed but is stricter than useful (it is
prefix-sensitive).

Two configurations were run:

* **A — as generated.** `fail_on_unknown_*=False`, serializer defaults, a fixed
  25-prefix `ns_map` of what Word emits.
* **B — tuned.** `+ ignore_default_attributes=True`, `ns_map` taken from the
  source document's root element.

Both use a tolerant `class_factory` (see 7.3), without which `invoice2013.docx`
crashes outright.

### 6.1 Per-file results

Times are the median of 3 timed passes after a warm-up, so they exclude the
one-off metadata build.

| docx | part | KiB | parse ms | serialize ms | identical (A) | identical (B) | remaining diffs (B) |
|---|---|---:|---:|---:|:--:|:--:|---|
| 2010-glow-then-AlternateContent.docx | document.xml | 9 | 0.8 | 1.7 | no | **yes** | - |
| 2010-glow-then-AlternateContent.docx | styles.xml | 12 | 4.2 | 4.5 | no | no | attr-dropped x10 |
| 2010-mcAlternateContent-in-header.docx | document.xml | 1 | 0.2 | 0.3 | no | **yes** | - |
| 2010-mcAlternateContent-in-header.docx | fontTable.xml | 2 | 0.4 | 0.5 | **yes** | **yes** | - |
| 2010-mcAlternateContent-in-header.docx | footer1.xml | 1 | 0.3 | 0.7 | no | **yes** | - |
| 2010-mcAlternateContent-in-header.docx | header1.xml | 7 | 0.4 | 0.7 | no | **yes** | - |
| 2010-mcAlternateContent-in-header.docx | styles.xml | 18 | 3.9 | 5.3 | no | no | attr-value x260, attr-dropped x10 |
| 2010-sample1.docx | document.xml | 4 | 1.2 | 2.0 | no | **yes** | - |
| 2010-sample1.docx | fontTable.xml | 2 | 0.4 | 0.6 | no | no | attr-dropped x1 |
| 2010-sample1.docx | styles.xml | 15 | 3.1 | 4.3 | no | no | attr-value x260, attr-dropped x4 |
| 2016_image_with_text_effects.docx | document.xml | 4 | 0.8 | 1.2 | no | no | attr-value x2 |
| 2016_image_with_text_effects.docx | fontTable.xml | 1 | 0.3 | 0.4 | no | no | attr-dropped x1 |
| 2016_image_with_text_effects.docx | styles.xml | 28 | 8.2 | 8.1 | no | no | attr-dropped x313, attr-value x32 |
| DrawingML_GraphicData_wps.docx | document.xml | 14 | 1.1 | 1.9 | no | **yes** | - |
| DrawingML_GraphicData_wps.docx | fontTable.xml | 1 | 0.3 | 0.4 | no | no | attr-dropped x1 |
| DrawingML_GraphicData_wps.docx | styles.xml | 43 | 10.3 | 12.4 | no | no | attr-dropped x320, attr-value x35 |
| Headers.docx | document.xml | 6 | 1.5 | 2.4 | no | no | attr-value x1 |
| Headers.docx | fontTable.xml | 2 | 0.4 | 0.6 | no | no | attr-dropped x1 |
| Headers.docx | header1.xml | 1 | 0.1 | 0.2 | **yes** | **yes** | - |
| Headers.docx | header2.xml | 1 | 0.1 | 0.2 | **yes** | **yes** | - |
| Headers.docx | header3.xml | 1 | 0.1 | 0.2 | **yes** | **yes** | - |
| Headers.docx | header4.xml | 1 | 0.1 | 0.2 | **yes** | **yes** | - |
| Headers.docx | header5.xml | 1 | 0.1 | 0.2 | **yes** | **yes** | - |
| Headers.docx | header6.xml | 1 | 0.1 | 0.2 | **yes** | **yes** | - |
| Headers.docx | styles.xml | 18 | 3.3 | 4.8 | no | no | attr-value x260, attr-dropped x7 |
| Images.docx | document.xml | 3 | 0.8 | 1.3 | no | no | attr-value x4 |
| Images.docx | fontTable.xml | 1 | 0.3 | 0.4 | **yes** | **yes** | - |
| Images.docx | styles.xml | 15 | 3.0 | 4.2 | no | no | attr-value x260, attr-dropped x4 |
| Symbols.docx | document.xml | 772 | 260.9 | 377.4 | no | no | attr-value x30 |
| Symbols.docx | fontTable.xml | 3 | 0.8 | 1.0 | no | no | attr-dropped x1 |
| Symbols.docx | styles.xml | 21 | 4.5 | 6.6 | no | no | attr-value x278, attr-dropped x14 |
| invoice2013.docx | document.xml | 25 | 5.2 | 7.0 | no | no | attr-value x4 |
| invoice2013.docx | fontTable.xml | 2 | 0.5 | 0.7 | no | no | attr-dropped x1 |
| invoice2013.docx | styles.xml | 12 | 4.0 | 4.8 | no | no | attr-value x25, attr-dropped x12 |
| sample-docx.docx | document.xml | 17 | 6.0 | 8.5 | no | no | attr-value x9, attr-dropped x1 |
| sample-docx.docx | fontTable.xml | 3 | 0.7 | 0.9 | no | no | attr-dropped x1 |
| sample-docx.docx | header1.xml | 1 | 0.1 | 0.3 | **yes** | **yes** | - |
| sample-docx.docx | numbering.xml | 6 | 2.0 | 2.5 | no | no | attr-dropped x7 |
| sample-docx.docx | styles.xml | 20 | 4.3 | 5.8 | no | no | attr-value x260, attr-dropped x8 |
| tables.docx | document.xml | 51 | 20.6 | 25.8 | no | no | attr-value x3 |
| tables.docx | fontTable.xml | 2 | 0.5 | 0.7 | no | no | attr-dropped x1 |
| tables.docx | styles.xml | 24 | 5.7 | 7.7 | no | no | attr-value x260, attr-dropped x11 |
| toc.docx | document.xml | 51 | 10.4 | 14.5 | no | no | attr-dropped x11 |
| toc.docx | fontTable.xml | 2 | 0.5 | 0.7 | no | no | attr-dropped x1 |
| toc.docx | footer1.xml | 2 | 0.3 | 0.6 | no | **yes** | - |
| toc.docx | footer2.xml | 2 | 0.4 | 0.7 | no | **yes** | - |
| toc.docx | styles.xml | 15 | 3.5 | 4.8 | no | no | attr-value x38, attr-dropped x28 |
| w14_texteffects.docx | document.xml | 6 | 1.2 | 1.8 | no | **yes** | - |
| w14_texteffects.docx | fontTable.xml | 1 | 0.3 | 0.5 | no | no | attr-dropped x1 |
| w14_texteffects.docx | styles.xml | 15 | 2.3 | 3.6 | no | no | attr-value x260, attr-dropped x4 |

**Bit-identical (canonical) parts: 9/50 in A, 18/50 in B.**

### 6.2 Aggregate difference categories

| category | A (as generated) | B (tuned) |
|---|---:|---:|
| `element-dropped` | **0** | **0** |
| `element-added` | **0** | **0** |
| `child-order` (lost order in a choice group) | **0** | **0** |
| `text` / `tail-text` | **0** | **0** |
| `attribute-added` (schema default materialised) | 5 516 | 0 |
| `attribute-value` (boolean `1`/`0` -> `true`/`false`) | 3 034 | 2 281 |
| `attribute-dropped` | 10 | 773 |

Every single difference across 50 parts is an **attribute** difference. Nothing
structural was lost.

---

## 7. The concrete fidelity problems

### 7.1 Schema defaults: damned either way

xsdata materialises a field default from `<xsd:attribute default="...">`, and the
serializer writes every non-`None` field. `word/styles.xml` gets ~200-800 spurious
attributes per file:

```
[attribute-added] /w:styles/w:latentStyles/w:lsdException   w:locked='true'
[attribute-added] /w:styles/w:docDefaults/.../w:spacing     w:afterAutospacing='false'
[attribute-added] /w:document/w:body/w:sectPr/w:cols        w:equalWidth='true'
```

(1 400x `w:lsdException/@locked`, 304x `@semiHidden`, 290x `@unhideWhenUsed`,
183x each of `w:spacing/@before|afterAutospacing`, ...)

`SerializerConfig(ignore_default_attributes=True)` removes all 5 516 of them —
but then it also removes attributes that *were explicitly present* and happen to
equal the default, because the dataclass cannot distinguish "absent" from
"present and equal to default":

```
[attribute-dropped] /w:latentStyles      w:defSemiHidden='1'       (parsed True == default True)
[attribute-dropped] /w:style             w:customStyle='1'
[attribute-dropped] /w:lvl               w:tentative='1'
[attribute-dropped] /w:hyperlink         w:history='1'
```

773 such drops in config B (323x `lsdException/@semiHidden`, 317x
`@unhideWhenUsed`, 69x `w:style/@customStyle`, ...).

**There is no config that gets both right.** The fix would be `Optional[bool]` with
`default=None` and the schema default applied only on *read* — i.e. either strip
the `default="..."` attributes from the schema copy (making all of them
`Optional`, as docx4j's Java code effectively does by using `Boolean`), or
post-process the generated dataclasses.

### 7.2 `mc:Ignorable` is silently dropped on types that don't declare it

docx4j's `wml.xsd` explicitly declares `mc:Ignorable` on `CT_Document`, `CT_HdrFtr`,
`CT_Styles`, `CT_Numbering` — but **not on `CT_FontsList`**. Result, on every
single sample:

```
[attribute-dropped] /w:fonts    mc:Ignorable='w14'
```

Strict mode confirms it is the *only* unknown-attribute class in the corpus:

```
ParserError: Unknown attribute {...main}fonts:{...markup-compatibility/2006}Ignorable
```
(10 occurrences — one per sample's `fontTable.xml`; no other part reported any
unknown property or attribute.)

**Verified remedy:** add `<xsd:anyAttribute namespace="##other"
processContents="lax"/>` to the affected complexTypes. xsdata maps it to an
`other_attributes: dict` and round-trips it exactly:

```python
Fonts(font=['a'], other_attributes={'{...markup-compatibility/2006}Ignorable': 'w14 w15'})
# -> <t:fonts ... mc:Ignorable="w14 w15">...
```

### 7.2b `mc:Ignorable` + `ns_map` can produce XML Word will reject

`mc:Ignorable` is round-tripped as an opaque string. If your `ns_map` does not
declare a prefix that the string names, you emit MCE-invalid XML. With the fixed
25-prefix `ns_map` (config A) this happened on 4 parts:

```
INVALID 2016_image_with_text_effects/word_document:
  Ignorable=['w14','w15','w16se','w16cid','wp14']  undeclared=['w16se','w16cid']
INVALID DrawingML_GraphicData_wps/word_styles:  undeclared=['w16se','w16cid']
```

`w16se`/`w16cid` are 2016 namespaces that aren't in docx4j's `wml.xsd` set at all.
Config B (ns_map taken from the source root) has zero such cases. **Any real
library must cross-check `mc:Ignorable` against the emitted `ns_map`** — xsdata
will not do it for you.

### 7.3 Required fields crash on schema-invalid-but-real documents

`invoice2013.docx` has a `<w:tbl>` with no `<w:tblGrid>`. That is schema-invalid;
Word and JAXB both accept it. xsdata generates `tbl_grid` as a *required*
keyword-only field, so:

```
TypeError: CtTbl.__init__() missing 1 required keyword-only argument: 'tbl_grid'
```

`fail_on_unknown_properties=False` does **not** help — this is missing *known*
content, not unknown content, and `ParserConfig` has no switch for it. The only
hook is `class_factory`; `scripts/roundtrip.py --tolerant-factory` supplies one
that fills missing required fields with `None` (4 hits across the corpus, all
`CtTbl.tbl_grid`). A real library would have to ship that, or generate everything
as `Optional`.

### 7.4 No unmarshal callbacks, and `--slots` blocks the usual workaround

There is no JAXB `afterUnmarshal` equivalent. `ParserConfig` exposes exactly one
hook — `class_factory(clazz, params)` — which sees the class and its parameters
but **not the parent node**, so docx4j's parent-pointer pattern cannot be built
during parsing; it needs a separate post-order walk.

Worse, `--slots` makes the objects non-extensible *and* non-weak-referenceable:

```python
>>> p = P(); p.parent = None
AttributeError: 'P' object has no attribute 'parent' and no __dict__ for setting new attributes
>>> weakref.ref(p)
TypeError: cannot create weak reference to 'P' object
```

So a parent pointer can't even be retrofitted at runtime. You must either drop
`--slots` (paying memory) or generate a `parent` field. docx4j's
`bindings.xjb` swaps in `ArrayListDocx4j` precisely to keep parents wired; there
is no xsdata analogue.

### 7.5 Lenient mode is completely silent

With `fail_on_unknown_properties=False` the parser returns a `SkipNode()`
(`parsers/nodes/element.py:468`) — **no warning, no log record**. I captured
`warnings` across all 50 parts: *zero* warnings emitted, even though
`mc:Ignorable` was being discarded 10 times. The only way to learn what is being
dropped is a strict run, which aborts on the *first* problem. In practice you need
the diff harness in this repo (or a custom handler) to audit losses.

### 7.6 Boolean formatting — not xsdata's fault

All 2 281-3 034 `attribute-value` differences are the same thing:
`w:val="1"` -> `w:val="true"`, `"0"` -> `"false"`. This is because docx4j's
`wml.xsd` deliberately uses `xsd:boolean` for these
(`BooleanDefaultTrue`/`BooleanDefaultFalse`, with a long comment explaining the
choice) rather than `ST_OnOff`. JAXB does exactly the same thing. Where the schema
*does* use `ST_OnOff`, xsdata generates a 6-member `Enum`
(`TRUE/FALSE/ON/OFF/VALUE_0/VALUE_1`) which round-trips the literal byte-for-byte.
No sample in the corpus used `on`/`off`, so that path is untested here.

### 7.7 Minor

* `standalone="yes"` in the XML declaration is not preserved
  (`SerializerConfig` has no `standalone` option). Harmless for OPC.
* Namespace declarations that the source declares but never uses (`ve`, `odc`,
  `odq`, `dsp`, ...) are dropped, and the serializer emits *all* `ns_map` entries on
  the root whether used or not. Both are semantically irrelevant but make byte
  diffs noisy.
* With `--nsmap-from-source`, namespaces Word declares *locally* (on `<w:drawing>`
  descendants: `a`, `pic`, `a14`) are missing from the root map, so xsdata invents
  `ns15`, `ns29`... prefixes. Supplying a complete static map (config A) avoids this.
* 205 of 1 625 generated modules are numbered duplicates (`st_on_off_1.py` ...
  `st_on_off_4.py`, `ct_page_margins_1/2.py`): the same ECMA type reached through
  different schema documents is emitted several times. Cosmetic, but it means
  `isinstance` checks can surprise you.

---

## 8. The specific checks (a)-(e)

`scripts/checks.py out/tuned`:

```
## (a) element order inside w:body / w:p / w:hdr / w:ftr
   50/50 parts preserve child order exactly

## (b) w14 / w15 node+attribute counts (orig -> round-trip)
   OK  2010-glow-then-AlternateContent/word_document   w14 8->8     w15 0->0
   OK  2010-sample1/word_document                      w14 8->8     w15 0->0
   OK  2016_image_with_text_effects/word_document      w14 2->2     w15 0->0
   OK  DrawingML_GraphicData_wps/word_document         w14 99->99   w15 0->0
   OK  invoice2013/word_document                       w14 9->9     w15 7->7
   OK  w14_texteffects/word_document                   w14 144->144 w15 0->0

## (c) mc:AlternateContent / mc:Fallback counts
   OK  2010-glow-then-AlternateContent/word_document   AC 1->1  Fallback 1->1
   OK  2010-mcAlternateContent-in-header/word_header1  AC 1->1  Fallback 1->1
   OK  DrawingML_GraphicData_wps/word_document         AC 1->1  Fallback 1->1

## (d) w:t text + xml:space
   3709 w:t nodes compared; text mismatches in 0 parts
   114 xml:space attributes present; mismatches in 0 parts
```

* **(a) Run order survives.** Compound fields work. All 50 parts reproduce the exact
  child sequence of `w:body`, `w:p`, `w:hdr`, `w:ftr` — including interleaved
  `w:commentRangeStart` / `w:r` / `w:bookmarkStart` / `w:ins` / `m:oMathPara`.
* **(b) w14/w15 survive**, both elements and attributes (`w14:paraId`,
  `w14:textId`, `w14:glow`, `w14:textOutline`, `w15:collapsed`, ...). 270 w14 nodes
  and 7 w15 nodes, all preserved, because `wml.xsd` imports both.
* **(c) `mc:AlternateContent` survives**, including `mc:Choice`/`mc:Fallback` and
  their `Requires` attributes and the `wps:`/`v:` payload inside. This is the
  result I least expected; docx4j's `markup-compatibility-2006-MINIMAL.xsd` models
  `AlternateContent` with lax wildcards and xsdata honours them.
* **(d) `xml:space="preserve"` and text survive.** 3 709 `w:t` nodes compared
  character-for-character, 114 `xml:space` attributes — zero mismatches.
  Note: `w:t` with leading/trailing spaces but *no* `xml:space` in the source stays
  that way (xsdata does not add one), so no new whitespace bugs are introduced.
* **(e) `ns_map` vs `mc:Ignorable` — this is the one that fails.** See 7.2/7.2b:
  the attribute itself is dropped on `w:fonts`, and if your `ns_map` misses a
  prefix named in `Ignorable` you produce MCE-invalid XML.

---

## 9. Scorecard against the original concerns

| concern | verdict | evidence |
|---|---|---|
| **Import time** | **REAL** | 857-905 ms warm, 1 329 ms cold, ~46 MiB RSS, 1 755 modules. `single-package` saves ~25%. |
| **Unknown content silently dropped** | **REAL but narrow** | Zero elements dropped in 50 parts. Only `mc:Ignorable` on `w:fonts` (10x). Fixable with `xsd:anyAttribute` (verified). But lenient mode is *totally silent* — no warning, no log. |
| **No unmarshal callbacks** | **REAL** | Only `class_factory(clazz, params)`, with no parent context. `--slots` additionally blocks runtime attribute injection *and* weakrefs, so docx4j's parent-pointer design needs a codegen change. |
| **Prefix control** | **NON-ISSUE (with care)** | `serializer.render(obj, ns_map=...)` gives full control; prefixes come out exactly as asked. Two gotchas: xsdata declares everything on the root and invents `nsN` for anything missing, and nothing cross-checks `mc:Ignorable`. |
| **Speed** | **REAL** | 3.0 MiB/s parse, 2.0 MiB/s serialize vs 69.5 MiB/s for bare lxml. Serialization is *slower than parsing*. Fine at 50 KiB (21+26 ms), 0.6 s for a 772 KiB part. |
| **Element order in choice groups** | **NON-ISSUE** | `--compound-fields` works; 50/50 parts preserve order exactly. |
| **w14/w15 extension markup** | **NON-ISSUE** | fully preserved. |
| **mc:AlternateContent** | **NON-ISSUE** | fully preserved, both branches. |
| **whitespace / `xml:space`** | **NON-ISSUE** | 3 709 text nodes, 114 `xml:space` attrs, zero mismatches. |

### New problems this experiment found

1. **`--structure-style namespaces` produces an unimportable package** for OOXML
   (circular imports); `namespace-clusters` refuses to generate at all. Only
   `clusters` / `single-package` / `filenames` are viable.
2. **xsdata 26.2 cannot generate anonymous complexTypes inside recursive groups**
   (`Error: Missing inner class` on `w:dir`/`w:bdo`). Needs a schema patch.
3. **Required (`minOccurs>=1`) elements become mandatory constructor arguments**, so
   a real-world schema-invalid document raises `TypeError` with no config to
   relax it. This is the only *hard* failure in the corpus and it hit 1 of 12
   documents.
4. **The schema-default dilemma** (7.1): both serializer settings are wrong, in
   opposite directions, on thousands of attributes per `styles.xml`.
5. **`--kw-only` does not exist** in xsdata 26.2 (it is unconditional), and
   **`xsdata` shells out to `ruff` by bare name**, so generation fails with
   `FileNotFoundError: 'ruff'` unless the venv's `bin` is on `PATH`.

### Overall assessment

xsdata is a **viable** foundation for a docx4j-analogue, and the structural
fidelity is genuinely good — better than the prior analysis assumed. But
"generate and ship" is not enough. A real library would need, at minimum:

* a patched schema copy (promote the anonymous `dir`/`bdo` types; add
  `xsd:anyAttribute` to the part-root complexTypes; strip `default="..."` so
  optional booleans stay `None`),
* a post-generation pass or `class_factory` to make required fields tolerant,
* its own `ns_map`/`mc:Ignorable` reconciliation on write,
* a decision on `--slots` vs parent pointers,
* and acceptance of ~0.9 s import and ~3 MiB/s throughput.

None of those are blockers; all of them are work that JAXB did for docx4j for free.

---

## 10. Reproducing

```bash
cd /home/jharrop/git/docx4j-python
./scripts/reproduce.sh
```

or step by step:

```bash
uv venv --python 3.14 .venv
uv pip install --python .venv/bin/python 'xsdata[cli,lxml]==26.2'
export PATH="$PWD/.venv/bin:$PATH"

# the baseline lives in baseline/ (git-ignored); xsdata writes relative to the working directory
(cd baseline && ../.venv/bin/xsdata generate -c ../.xsdata.xml ../schemas/wml/wml.xsd)

.venv/bin/python scripts/bench.py

.venv/bin/python scripts/roundtrip.py --strict --tolerant-factory --out out/default
.venv/bin/python scripts/roundtrip.py --ignore-defaults --nsmap-from-source \
        --tolerant-factory --out out/tuned

.venv/bin/python scripts/checks.py out/tuned
.venv/bin/python scripts/checks.py out/default
```

### Layout

```
.xsdata.xml            generator config
pyproject.toml
schemas/               91 xsd files, the transitive closure of wml/wml.xsd
                       (wml.xsd carries one marked patch, 2.1)
samples/               12 real .docx from docx4j's sample-docs
baseline/docx4j_py/generated/   1 625 modules, 1 373 dataclasses, 319 enums (git-ignored; see 10)
scripts/roundtrip.py   parse -> serialize -> diff harness
scripts/canon.py       canonicalisation + categorised structural diff
scripts/checks.py      targeted checks (a)-(e)
scripts/bench.py       import time + throughput
scripts/reproduce.sh
out/default/           per-part orig.xml / rt.xml / diff.txt, config A
out/tuned/             ditto, config B
```

### Samples used

| file | why |
|---|---|
| `2010-sample1.docx` | simple; w14 attributes, OMML, `xml:space` |
| `2010-glow-then-AlternateContent.docx` | `mc:AlternateContent` + w14 in `document.xml` |
| `2010-mcAlternateContent-in-header.docx` | `mc:AlternateContent` in `header1.xml` |
| `w14_texteffects.docx` | heaviest w14 usage (144 nodes) |
| `DrawingML_GraphicData_wps.docx` | `AlternateContent` wrapping `wps:` shapes |
| `2016_image_with_text_effects.docx` | 2016 markup, `w16se`/`w16cid` in `mc:Ignorable` |
| `invoice2013.docx` | w15 elements; contains the schema-invalid `w:tbl` |
| `Images.docx` | `w:drawing` / `wp:` / `a:` / `pic:` |
| `Headers.docx` | six headers |
| `tables.docx` | 51 KiB of nested tables |
| `toc.docx` | fields, footers, `xml:space` |
| `Symbols.docx` | 772 KiB `document.xml` — throughput stress |

Sources: `/home/jharrop/git/docx4j/docx4j-samples-docx4j/sample-docs/{,2010/,2016/,databinding/}`.
Nothing under `/home/jharrop/git/docx4j` was modified.

---

## Phase B result (2026-09-12)

This report is the feasibility experiment and is left as it was written. What it proposed has
since been built: see **`docs/change-requests/CR-001-object-model.md`**, whose section 12 records
Phase A (the fork) and section 13 Phase B (the schema patches, the docx4j name tables, the
per-namespace regeneration, parent pointers, the numeric-boolean serialiser and the
skipped-content report).

The headline numbers in this report have all moved. Where section 7 measured 18 of 50 parts
canonically identical, 5,516 invented attributes, 773 dropped ones and 2,281 value differences,
the round trip today is **50 of 50 parts canonically identical, zero differences in every
category, and zero skipped elements or attributes**, with `invoice2013.docx` parsing unaided and
no tolerant class factory. The only thing the corpus and the output still disagree about is the
spelling of 34 `xsd:boolean` attributes in one `styles.xml` (`true` against `1`), which
`scripts/canon.py` now counts as its own `boolean-spelling` category because the two are the same
value.

Two pointers changed with the layout. The unmodified-xsdata output this report measured is now
**`baseline/docx4j_py/generated/`** rather than `docx4j_py/generated/` — the generator owns every
file under `docx4j_py/` — and it is still run with `.venv` and `scripts/roundtrip.py`'s defaults,
which put `baseline/` in front of the repository root on the path. The Phase B package is
`docx4j_py.wml` and the rest, generated by `codegen/generate.sh` against the fork; see
`codegen/README.md`.
