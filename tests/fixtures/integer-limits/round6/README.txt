CR-066 round 6: the integer attributes round 5 left unmeasured or inferred (docx4j-25 builds, 2026-10-26, at
objects-ts's request with Jason's yes; Jason runs Word 365). Same method as round 5: Word opens a file and saves it,
and the save is the oracle. Every value is a plain integer (Word's lexical form), so nothing is refused for its form.

Files:
  groups/cr066-r6-A.docx   paragraph: w:ind firstLine, hanging, right, start, end; w:spacing after; w:spacing line
                           under lineRule=atLeast                                                   (68 cases)
  groups/cr066-r6-F.docx   frames: w:framePr w, h, hSpace, vSpace, x, y                             (60 cases)
  groups/cr066-r6-R.docx   runs: w:fitText, w:szCs, w:kern, w:bdr sz and space                     (44 cases)
  groups/cr066-r6-T.docx   tables: trHeight exact and auto; gridCol under tblLayout fixed (cells without tcW,
                           tblW auto); tblCellSpacing; tblCellMar left; tcMar top; wBefore, wAfter;
                           tcBorders top sz/space; tblBorders top sz/space                          (104 cases)
  groups/cr066-r6-P.docx   floating tables: w:tblpPr leftFromText, rightFromText, topFromText, bottomFromText,
                           tblpX, tblpY                                                             (60 cases)
  groups/cr066-r6-S.docx   sections (each continuous): w:pgMar left, right, header, footer, gutter, bottom;
                           w:cols space; w:cols/w:col w and space (equalWidth=0); w:lnNumType distance;
                           w:pgBorders top sz/space                                                 (114 cases)
  groups/cr066-r6-N.docx   numbering.xml: ten lists; list i's level 0 has w:legacy legacySpace and legacyIndent
                           both at value i (32767, 32768, 65535, 65536, 65537, -1, -32768, -32769, 7205001,
                           1205003); each case paragraph is numbered with that list                  (10 cases)
  groups/cr066-r6-E.docx   the no-digit forms ".", "+", "-", "+." in w:spacing/@w:before and w:ind/@w:left under a
                           style whose before and left are 100 (CR-065: Word opened python's E2-E4; 0 or ignored?) (8 cases)
  singles/cr066-r6-D1-<n>.docx   settings.xml w:defaultTabStop at the ten values, one per file (document-wide)
  singles/cr066-r6-D2-<n>.docx   settings.xml w:hyphenationZone at the ten values, one per file
  families/cr066-r6-<F>.docx     one attribute each (46), with the control Z-1; ONLY for a group Word refuses
  singles/cr066-r6-<id>.docx     one value each for the F, P, S and N families; ONLY for a refused family
  cases.txt                 id, group, attribute, place (pPr, rPr, tbl, sectPr, numbering:<numId>, settings), value,
                            the value if Word wraps signed, if it wraps unsigned
  families.txt              the family list with the type and what round 5 inferred

Not covered: w:object dxaOrig/dyaOrig (needs an embedded OLE object).

Values: 32767, 32768, 65535, 65536, 65537, -1, -32768, -32769, 7205001 (-3959 if it wraps signed) and 1205003 (25355);
page-sized attributes (frame w/h, col w) also 31680/31681; the rule rows (line atLeast, trHeight exact/auto) the eight
values either side of 32768; borders 48/49/255/256/32767/65535/65536/-1 for sz and 31/32/255/256/65536/-1 for space.
The styles ProbeBase (before 100, after 100, line 300 auto, left 100, right 100, firstLine 100) and ProbeRun
(kern 20, sz 18, szCs 18) give the tested properties a non-zero inherited value, so a value wrapped to 0 is written
and a dropped one is absent. Expect page breaks and odd layouts (a 22-inch frame, a 45-inch column); that is the point.

Steps (Word 365):
 1. Open each groups/cr066-r6-X.docx (A, F, R, T, P, S, N, E). Per file: opens clean / offers repair / refuses.
 2. File > Save As each one that opens as cr066-r6-X-word365.docx in groups/.
 3. Open and Save As each singles/cr066-r6-D1-<n>.docx and D2-<n>.docx (20 files) as ...-word365.docx in singles/.
    (One value per file because the setting is document-wide; the TAB line in D1 shows the default tab stops.)
 4. For a refused group: its families/ files, then singles/ for a refused family, Save As likewise.

Then tell docx4j-25: it reads the saves against cases.txt, records the width per attribute in CR-066 §3, and sends it
to docx4j-generated-objects-ts-82 (objects CR-016), docx4j-python-cd (CR-022) and docx4j-core-ts-5b.
