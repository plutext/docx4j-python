CR-065 phase 0: decimals and odd forms in integer attributes - what does Word do? (docx4j-25, 2026-10-26; Jason runs)

Why: a decimal in an integer attribute (w:spacing w:before="120.5", w:ind w:left="720.5", a border w:sz="7.5", ...)
makes docx4j's unmarshal fail unless the attribute is on mc-preprocessor.xslt's fixed rounding list; a failed main part
is then silently replaced by an empty one on the first edit (CR-065 phase 1 fixes that). The three ports want ONE
measured rule for how Word reads such values, before any of them fixes a rounding rule. Today's rounding (half-even) is
Xalan's format-number, not Word's.

Files:
  cr065-decimals-probe.docx   all 65 cases in one file (built by docx4j, then the attribute values edited in).
  cr065-structure-control.docx  the SAME document with every tested value made a plain integer (added after Word
                              refused the probe outright on 2026-10-26). If this opens clean, the refusal was the values.
  cases.txt                   the case list: id, attribute, value.
  families/cr065-A..I.docx    the same cases by family, each with the control paragraph Z65. Use these ONLY if Word
                              refuses the whole probe, to find which family it refuses.

Families: A w:spacing/@w:before (twips, not on docx4j's list) with the decimal set and the odd value forms
("240.", ".5", "+240", " 240.5 ", "240,5", "0x10", "1.5e2", "", "abc" ...); B w:ind/@w:left (signed twips) with
positive and negative halves; C attributes already on docx4j's list (does Word agree with docx4j's repair?);
D borders (@w:sz, @w:space); E tab stops; F run properties (w:sz, w:w, w:spacing, w:position); G tables
(tblW, tblInd, gridCol, trHeight, tcW pct); H section properties (pgSz, pgMar, SSRS's empty attributes), each a
continuous section; I the same decimals in styles.xml, settings.xml and the header.

RESULT 2026-10-26, round 1: cr065-decimals-probe.docx -> "Word experienced an error trying to open the file" (no repair offered).
Round 2 steps (Word 365):
 0. Open cr065-structure-control.docx. Does it open clean? (It should; if not, the fault is mine, not the values.)
    Then open each families/cr065-A.docx .. cr065-I.docx and say, per file: opens clean / offers repair / refuses.
    Save As each one that opens (repaired or not) as cr065-X-word365.docx in families/.
    For B (if it opens): Paragraph dialog on B20, B23, B24 -> Indentation Left. For A: A01, A03, A04 -> Spacing Before.
    Skip steps 1-4 below; they were round 1.

Round 1 steps (Word 365):
 1. Open cr065-decimals-probe.docx. Say whether Word reports a problem or offers a repair, and whether anything
    looks odd on the page (a case missing, a huge gap, a table collapsed).
 2. File > Save As cr065-decimals-probe-word365.docx, in this folder. (Word's save is the oracle for every attribute
    at once: it writes its own reading of each value, or drops it.)
 3. For these six paragraphs, click in the paragraph, open the Paragraph dialog (Home > Paragraph launcher) and
    note the value Word shows, with its unit:
      A01 (before="120.5")  -> Spacing Before
      A03 (before="121.5")  -> Spacing Before
      A04 (before="122.5")  -> Spacing Before
      B20 (left="720.5")    -> Indentation Left
      B23 (left="-120.5")   -> Indentation Left
      B24 (left="-121.5")   -> Indentation Left
    (6 pt = 120 twips; 0.5 in = 720 twips; the dialog may show cm or pt.)
 4. Close. If step 1 refused the file outright (no repair offered, or the repair failed), open each
    families/cr065-X.docx instead, say which open, which repair, which refuse, and Save As each one that opens
    as cr065-X-word365.docx in families/.

Then tell docx4j-25; it reads Word's saves against cases.txt, records the rule in CR-065 §2, and sends the result to
docx4j-python-cd, docx4j-generated-objects-ts-82 and docx4j-core-ts-5b.

RESULT 2026-10-26, round 2: the structural control opens fine; families B..I all opened fine (saved as *-word365.docx);
family A refused ("no good").
Round 3 steps: family A split into groups, in families/: cr065-A1-decimals.docx (120.5, 120.4, 121.5, 122.5, 240.0,
141.99999999999994), cr065-A2-forms.docx ("240.", ".5", "+240", "+240.5", " 240.5 ", " 240 "), cr065-A3-comma.docx
("240,5"), cr065-A4-hex.docx ("0x10"), cr065-A5-exponents.docx ("1.5e2", "2.5e0", "1e3"), cr065-A6-empty.docx (""),
cr065-A7-abc.docx ("abc"). Open each; say opens clean / repairs / refuses; Save As the ones that open as
cr065-A<n>-...-word365.docx in families/. Each has the control paragraph Z65.

RESULT 2026-10-26, round 3: A1 (decimals) and A6 ("") open clean and are saved; A2, A3, A4, A5, A7 refused.
Round 4 (OPTIONAL, only if cheap for you): A2's six forms one per file, in families/: cr065-A2a-trailing-dot.docx
("240."), cr065-A2b-leading-dot.docx (".5"), cr065-A2c-plus.docx ("+240"), cr065-A2d-plus-decimal.docx ("+240.5"),
cr065-A2e-spaced-decimal.docx (" 240.5 "), cr065-A2f-spaced-integer.docx (" 240 "). Say which open; Save As any that do.

ROUND 5 (2026-10-26) is CR-066's phase 0: integer attributes at and beyond 16 bits (Word wraps 7205001 to -3959, round 2's aside).
Its files and steps are in ../cr066/README.txt.
