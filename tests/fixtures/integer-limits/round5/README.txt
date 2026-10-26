CR-066 phase 0 = round 5 of the CR-065 probe: integer attributes at and beyond 16 bits - what does Word do?
(docx4j-25 builds the files, 2026-10-26; Jason runs Word 365)

Why: in CR-065's round 2, three values the generator left in by mistake showed that Word stores twips in 16 bits and
WRAPS an out-of-range value: 7205001 in a style's w:ind/@w:left became -3959 (signed), 7205002 in w:defaultTabStop
became 61578 and 1205003 in a header's w:spacing/@w:before became 25355 (unsigned); no error, no repair. docx4j (and
the ports) write whatever they are given, and docx4j's own helpers compute twips from units with no range check. Before
docx4j reports or guards such values (the new write-side CR-066), the three ports want ONE measurement of the range
per attribute type: does Word wrap, clamp (at its storage limit or at its UI limit), keep, or refuse?

Every value here is a plain integer (Word's lexical form, CR-065 round 4), so nothing is refused for its FORM.
Expect huge gaps and page breaks: 32767 twips of space before is 22.7 inches; that is the point.

Files:
  groups/cr066-P.docx   paragraph properties: w:spacing/@w:before (unsigned), w:ind/@w:left, w:spacing/@w:line under
                        lineRule auto and exact, w:tab/@w:pos (signed)                                    (42 cases)
  groups/cr066-B.docx   borders: w:pBdr/w:top/@w:sz (eighth-points), @w:space (points)                    (18 cases)
  groups/cr066-R.docx   run properties: w:sz (half-points), w:position, w:spacing, w:w (text scale)       (33 cases)
  groups/cr066-T.docx   tables: w:tblW, w:tcW, w:tblInd with type=dxa; w:tcW, w:tblW with type=pct;
                        w:trHeight                                                                        (49 cases)
  groups/cr066-S.docx   sections (each continuous): w:pgSz/@w:w, @w:h, w:pgMar/@w:top, w:pgNumType/@w:start
                        (the footer's PAGE field shows the start)                                         (32 cases)
  groups/cr066-N.docx   a NEGATIVE (-1, -32769) in attributes whose type is unsigned (before, border sz and
                        space, w:sz, text scale, trHeight, pgMar top, pgSz w), kept apart in case Word refuses
                        the sign where the type is unsigned                                               (16 cases)
  families/cr066-<F>.docx   the same cases by family (one attribute each; P1..P5, B1, B2, R1..R4, T1..T6, S1..S4, N1),
                        each with the control paragraph Z-1. Use these ONLY for a group Word refuses.
  singles/cr066-<id>.docx   one value each, for the S families and N1 only. Use these ONLY for a family Word refuses.
  cases.txt             id, group, attribute, value, the value Word would keep if it WRAPS at 16 bits, and if it CLAMPS.
  families.txt          the family list with the type and Word's UI limit where known.

Values per family: the 16-bit boundaries 32767, 32768, 65535, 65536, 65537; for signed types also -1, -32768, -32769;
the two wrap controls 7205001 (signed; -3959 if it wraps) and 1205003 (unsigned; 25355); and each type's UI or schema
limit and the value just past it (pgSz 31680/31681, w:sz 3276/3277, position 3168/3169, character spacing
31680/31681, border sz 48/49, border space 31/32, text scale 600/601, pct 5000/5001, pgNumType start 2147483647/8 and
4294967295/6).

Each case paragraph is labelled "<id> <attribute>="<value>"". For the P and B groups the label paragraph itself
carries the value; in R the run after "then:" carries it; in T the table after the label; in S the section that the
paragraph after the label ends. The paragraph style ProbeBase (before 100, left 100, line 300 auto) and character
style ProbeRun (sz 18, position 2, spacing 2, w 90) give every tested property a non-zero inherited value, so that a
value Word wraps to 0 is WRITTEN as 0 in its save and a value Word drops is ABSENT.

Steps (Word 365):
 1. Open each groups/cr066-X.docx (P, B, R, T, S, N). Per file say: opens clean / offers repair / refuses.
    N and S are the likely refusals; the others should open (round 2's wrapped values opened clean).
 2. File > Save As each one that opens (repaired or not) as cr066-X-word365.docx in groups/.
    Word's save is the oracle: it writes its own reading of each value (kept, wrapped, clamped) or drops it.
 3. For S only, a glance: the footer page number on the pages of the S4 sections (start=32767, 32768, 65535, 65536,
    65537, 2147483647, 2147483648, 4294967295, 4294967296, -1) - say what you see if it is quick; the save answers too.
 4. For a group that refuses: open its families/cr066-<F>.docx files, say which open, Save As those that do
    (cr066-<F>-word365.docx in families/). For a refused S or N family, the singles/ files the same way.

Then tell docx4j-25: it reads the saves against cases.txt (read66.py: kept / wrap / clamp / absent per case), records
the rule per type in CR-066 §3, and sends it to docx4j-python-cd, docx4j-generated-objects-ts-82 and docx4j-core-ts-5b.

RESULT 2026-10-26 (Jason): all six group files opened clean, no repair; saves in groups/*-word365.docx; families/ and singles/
not needed. Read by docx4j-25: no value refused or repaired; the storage width is per attribute (int16 for most signed
measures and for pgMar top; uint16 for before, w:sz, text scale, pgSz; tab pos clamped to +/-31680 and pgSz w to 144..31681
after the wrap; trHeight and exact line spacing carry the rule in the sign; border sz 8 bits, border space 5 bits;
pgNumType start int32). Full table in CR-066 section 3.1. Round 5 is complete.

ROUND 6 (2026-10-26, objects-ts's request, Jason's yes): the attributes round 5 left unmeasured or inferred; files and steps in round6/README.txt.
