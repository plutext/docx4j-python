On/off probes for w:b, built 24 October 2026 by the docx4j-portfolio session (build script: onoff/build.py in its
scratch area), at Jason's request, after docx4j, core-ts and docx4j-python each read <w:b/> as bold on and the python
session found its readers disagree on a w:val they cannot parse.

Each probe paragraph is "label: MARKER (end)", and only the MARKER run carries the w:b being tested. The question for
each marker is: does Word show it bold?

onoff-valid.docx   every value ST_OnOff allows in transitional OOXML, plus the bare and absent forms:
                   V1 <w:b/>             V1_BARE
                   V2 w:val="true"       V2_TRUE
                   V3 w:val="false"      V3_FALSE
                   V4 w:val="on"         V4_ON
                   V5 w:val="off"        V5_OFF
                   V6 w:val="1"          V6_ONE
                   V7 w:val="0"          V7_ZERO
                   V8 no w:b             V8_ABSENT

Each of the next four holds one value ST_OnOff does not allow, so that a file Word refuses tells us which value it
refused. Each also holds two controls: C1_BARE (<w:b/>, bold) and C2_ABSENT (no w:b, not bold).
onoff-empty.docx   E1 w:val=""           E1_EMPTY
onoff-space.docx   S1 w:val=" "          S1_SPACE
onoff-TRUE.docx    U1 w:val="TRUE"       U1_UPPER_TRUE
onoff-yes.docx     Y1 w:val="yes"        Y1_YES

Please, for each file:
1. Does Word open it, and does it offer a repair? If it repairs, what does it say?
2. A screenshot (filename.PNG), so each marker's bold or not is visible.
3. A save, as filename-word365.docx, so we can see how Word writes the value back.

LibreOffice 26.2 opened all five (PDF export, 24 October); its reading of each marker was not recorded.


Second round, built 24 October 2026 (build script: onoff/build2.py in the same scratch area), after the first round
showed Word reads w:b as bold unless its w:val is false, off or 0, every other value (even "" and "yes") bold.
The same layout: "label: MARKER (end)", only the MARKER run carrying the property tested.

props-b-case.docx      w:b's off words in other cases, and with spaces: K1 "FALSE", K2 "OFF", K3 "False", K4 "Off",
                       K5 " false", K6 "false ", K7 "00", K8 "no"; controls C1 bare (bold), C2 absent
props-b-styles.docx    off values against bold styles (ParaBold, CharBold) and underline against CharUnderline:
                       S1 CharBold alone (bold); S2/S3/S4 CharBold with direct "false"/"off"/"0";
                       S5 ParaBold alone (bold); S6 ParaBold with direct "false";
                       S7 ParaBold + CharBold (the toggle: two bold styles cancel, per ECMA-376 17.7.3);
                       S8 the same plus direct <w:b/>; S9 CharUnderline alone; S10 with direct w:u "none";
                       S11 with direct <w:u/>
props-toggles-1.docx   i, caps, smallCaps, strike, dstrike: each bare, "off", "0", "" and "FALSE" (T1 to T5)
props-toggles-2.docx   outline, shadow, emboss, imprint, vanish: the same (T6 to T10); vanish is hidden text, so
                       please screenshot with formatting marks shown, as before
                       (marker text is lower case, so that caps and small caps show)
props-u-valid.docx     w:u, whose w:val is ST_Underline, not on/off: U1 <w:u/> with no w:val, U2 single, U3 double,
                       U4 words, U5 wave, U6 none, U7 a colour and no w:val, U8 single in red
props-u-case.docx      w:u values in another case, and empty: X1 "", X2 "SINGLE", X3 "Single", X4 "NONE";
                       controls C1 single, C2 absent
props-u-other.docx     w:u values ST_Underline lacks, on/off words among them: Y1 "bogus", Y2 "true", Y3 "1", Y4 "on",
                       Y5 "off", Y6 "0", Y7 "false"; the same controls
props-enum-missing.docx  properties whose required w:val is missing: M1 <w:highlight/>, M2 <w:vertAlign/>, M3 <w:em/>,
                       M4 <w:color/>; controls C1 highlight yellow, C2 superscript, C3 emphasis dot, C4 red
props-enum-other.docx  the same properties with values in another case or not in the enumeration: N1 highlight
                       "YELLOW", N2 highlight "bogus", N3 vertAlign "Superscript", N4 vertAlign "bogus", N5 em "DOT",
                       N6 color "red", N7 color "ff0000" (lower-case hex); the same controls

As before, for each file: does Word open it (any repair prompt, and its words); a screenshot (filename.PNG); a save as
filename-word365.docx. LibreOffice 26.2 opened all nine (PDF export, 24 October).


Third round: colour names, built 24 October 2026 (build script: onoff/build3.py in the same scratch area), after round
two showed w:color w:val="red" drawn red and saved as FF0000, though ST_HexColor allows only "auto" and six hex digits.
The same layout; only the MARKER run carries a w:color. Each file has controls C1 (FF0000, red) and C2 (no w:color).

colour-names-highlight.docx  the sixteen names ST_HighlightColor uses, as w:color values: H1 black, H2 blue, H3 cyan,
                             H4 green, H5 magenta, H6 red, H7 yellow, H8 white, H9 darkBlue, H10 darkCyan,
                             H11 darkGreen, H12 darkMagenta, H13 darkRed, H14 darkYellow, H15 darkGray, H16 lightGray
colour-names-css.docx        CSS names ST_HighlightColor lacks: N1 orange, N2 purple, N3 navy, N4 teal, N5 maroon,
                             N6 silver, N7 gray, N8 grey, N9 lime, N10 olive, N11 aqua, N12 fuchsia, N13 pink,
                             N14 brown, N15 gold
colour-case.docx             auto and names in other cases: K1 auto, K2 Auto, K3 AUTO, K4 Red, K5 RED, K6 darkred,
                             K7 DarkRed, K8 DARKRED, K9 Blue
colour-hex.docx              hex spelt other ways: X1 ff0000, X2 Ff0000, X3 F00, X4 #FF0000, X5 0xFF0000,
                             X6 FF00000 (seven digits), X7 FF000 (five), X8 GG0000, X9 "FF0000 ", X10 " FF0000", X11 "",
                             X12 0000FF (blue, a second valid control)

White (H8) is invisible on the page, and black is hard to tell from automatic; the saves will show what Word made of
them. As before, for each file: does Word open it, and does it repair (its words); a PDF; a save as
filename-word365.docx. LibreOffice 26.2 opened all four (PDF export, 24 October).


Fourth round: the colour table, built 24 October 2026 (build script: onoff/build4.py in the same scratch area), after
round three showed Word maps colour names case-insensitively but drops darkYellow, lightGray and grey. That fits the
classic HTML (Netscape) colour table, which spells most grays with an a but has lightgrey, and has no grey; it does not
fit .NET's KnownColor (which has LightGray) or CSS 3 (which has both spellings). The same layout and controls.

colour-table-grey.docx   G1 lightgrey, G2 darkgrey, G3 dimgray, G4 dimgrey, G5 slategray, G6 slategrey,
                         G7 lightslategray, G8 lightslategrey, G9 darkslategray, G10 darkslategrey, G11 gainsboro
colour-table-other.docx  CSS 4 and later: T1 rebeccapurple, T2 transparent, T3 currentcolor; X11 names no HTML list
                         has: T4 lightgoldenrod, T5 violetred, T6 navyblue; a space: T7 "dark red"; system colours:
                         T8 window, T9 buttonface, T10 windowtext; the long HTML list: T11 aliceblue, T12 papayawhip,
                         T13 lightgoldenrodyellow, T14 mediumvioletred, T15 cornflowerblue, T16 darkorange,
                         T17 yellowgreen

As before: does Word open it (and repair?); a PDF; a save as filename-word365.docx. LibreOffice 26.2 opened both.


Fifth round: values over styles, built 24 October 2026 by docx4j-25 (build script: build5.py, beside these files),
after rounds 1 to 4 left one question each for the save rules of docx4j's CR-060, core-ts's CR-015 (its D5) and
docx4j-python's CR-018: does a value Word doesn't accept override a style that states the same property, or does the
style's value show, as if the direct property were absent? The same layout: "label: MARKER (end)", only the MARKER
run carrying the direct property tested; where a paragraph style is the one tested, the whole paragraph has it.

props-u-styles.docx      w:u over styles that underline DOUBLE, so that the style's underline (double), no underline,
                         and a single one can be told apart: A1 to A12 over the character style CharUnderlineDouble,
                         B1 to B12 the same over the paragraph style ParaUnderlineDouble. Each: the style alone (the
                         control); direct <w:u/>; w:val "bogus", "Single", "SINGLE", "", "true", "NONE"; "none" (the
                         control: removes); "single" (the control: single); w:color="FF0000" with no w:val; "bogus"
                         with w:color="FF0000". Controls C1 (no style, no w:u) and C2 (no style, single).
                         For each marker: double, single or no underline, and its colour?
colour-styles.docx       w:color over styles that colour the text: R1 to R11 over the character style CharRed (red),
                         G1 to G11 over the paragraph style ParaGreen (green). Each: the style alone; direct "F00",
                         "bogus", "", bare <w:color/>, "grey", "0xFF0000", "FF00000" (seven digits); "auto" (the
                         control); "0000FF" and "blue" (the controls: blue). Control C1 (no style, no w:color).
                         For each marker: red (or green: the style's), black (automatic), or blue?
props-enum-styles.docx   highlight, vertAlign and em over styles that state all three (yellow highlight, superscript,
                         dot emphasis): M1 to M12 over the character style CharMarked, and P1, P2, P4, P7, P10 over
                         the paragraph style ParaMarked. Each marker has one direct value: highlight "YELLOW", "bogus",
                         bare, "none" (the control: removes), "green" (the control); vertAlign "Superscript", bare,
                         "baseline" (the control: removes); em "DOT", bare, "none" (the control: removes).
                         Word is expected to offer a repair (rounds 2 and 3 found these values repaired away); please
                         let it repair, and note its words. For each marker after the repair: highlighted yellow,
                         raised, dotted (the style's), or not?
colour-elsewhere.docx    colour names in the colour attributes other than w:color, no styles: U1 to U4 a single
                         underline whose w:color is "red", "Navy", "F00", "FF0000" (the control); S1 to S4 run shading
                         whose w:fill is "red", "Aqua", "F00", "FF0000" (the control); B1 to B3 a run border whose
                         w:color is "red", "F00", "FF0000" (the control).
                         For each marker: the underline's, the shading's or the border's colour (red, navy, aqua,
                         black/automatic, or none)?

As before, for each file: does Word open it (any repair prompt, and its words); a PDF; a save as
filename-word365.docx. LibreOffice 26.2 opened all four (PDF export, 24 October).


Sixth round: the questions round 5 left, built 24 October 2026 by docx4j-25 (build script: build6.py, which uses
build5.py, both beside these files). Round 5 found an invalid value over a style reads as absent (the style's value
shows), except w:u w:val="true", which removed the style's underline; that Word drew no highlight from a character
style (in a file it had repaired); and that Word reads colour names in a shading fill, a border colour and an
underline colour. The same layout as before.

props-u-onoff-styles.docx  the on/off words as a w:u value, over styles that underline DOUBLE: A1 to A15 over the
                           character style CharUnderlineDouble, B1 to B15 over the paragraph style
                           ParaUnderlineDouble. Each: the style alone (the control); "true", "True", "TRUE", "1", "on",
                           "yes", "false", "0", "off", "no"; "none", "single", "words" (the controls); "true" with
                           w:color="FF0000". C1 "true" and C2 "on" with no style.
                           For each marker: double (the style's), single, words only, or none, and its colour?
colour-auto-styles.docx    auto in other spellings over styles that colour the text: R1 to R8 over CharRed (red), G1 to
                           G8 over the paragraph style ParaGreen (green): the style alone; "auto" (round 5: black);
                           "Auto", "AUTO", " auto", "auto ", "#auto"; "0000FF" (the control: blue).
                           For each marker: red (or green: the style's), black (automatic), or blue?
highlight-char-styles.docx highlight from a character style, with nothing for Word to repair: H1 CharHighlight (yellow
                           highlight) alone; H2 CharHighlightBold (bold and yellow highlight); H3 CharBasedOnHighlight
                           (italic, based on CharHighlight); H4 CharShading (yellow shading, for comparison); H5 the
                           paragraph style ParaHighlight; H6 direct yellow (the control); H7 CharHighlight with direct
                           green; H8 CharHighlight with direct none; H9 ParaHighlight paragraph and CharHighlight run.
                           Please say if Word offers a repair (it shouldn't). For each marker: highlighted or not?
colour-elsewhere-styles.docx  over character styles that make a shading, a border and an underline red: S1 to S6 the
                           shading style alone, then a direct fill of "F00", "bogus", "Aqua", "auto", "00FFFF" (aqua,
                           the control); B1 to B4 the border style alone, then a direct border colour "F00", "Navy",
                           "000080" (navy, the control); U1 to U5 the underline style (single, red) alone, then a direct
                           single underline coloured "F00", "Navy", "auto", and a direct double with no colour.
                           For each marker: the shading's, the border's or the underline's colour (the style's red,
                           aqua, navy, black/automatic, or none)?

As before, for each file: does Word open it (any repair prompt, and its words); a PDF; a save as
filename-word365.docx. LibreOffice 26.2 opened all four (PDF export, 24 October).


Seventh round, and docx4j's saves: CR-061's Word check, built 24 October 2026 by docx4j-25 (build script build7.py,
which uses build5.py; the -docx4j files are docx4j's own saves of the earlier probes, made by docx4j at commit
aaadd93f9). docx4j now writes every colour plainly on save: a name or any hex spelling as upper-case RRGGBB, a value
Word rejects left out of a text or underline colour and written "auto" in a fill or border colour, and the highlight,
vertAlign and em Word repairs away left out. The check is that Word reads docx4j's saves exactly as it read the
probes, and never offers a repair.

<name>-docx4j.docx         docx4j's saves of colour-case, colour-hex, colour-names-css, colour-names-highlight,
                           colour-table-grey, colour-table-other, colour-styles, colour-auto-styles, colour-elsewhere,
                           colour-elsewhere-styles, props-enum-missing, props-enum-other, props-enum-styles,
                           props-u-styles and props-u-onoff-styles (15 files). For each: does Word offer a repair
                           (it shouldn't, even for props-enum-*, which Word repaired as probes); and a PDF, which
                           should look the same as Word's PDF of the probe. (The macro also saves a -word365 copy;
                           that's fine.)
colour-pattern.docx        a 50% shading pattern's colour (w:shd w:val="pct50" w:color=..., white fill): P1 "red",
                           P2 "Navy", P3 "F00", P4 "FF0000" and P5 "auto" (the controls); P6 to P8 over the character
                           style CharPatternRed (a red pattern): alone, with a direct "F00", with a direct "Navy".
                           For each marker: the pattern's colour (red, navy, black/automatic, or the style's red)?
colour-background-red.docx, -navy.docx, -f00.docx, -hex.docx
                           the page colour, w:background w:color = "red", "Navy", "F00", "FF0000" (the control), one
                           per file (settings turn the page colour on). The page colour may not be in a PDF (Word
                           prints it only with "Print background colors and images" on); the save shows how Word read
                           it. If you can, note the page colour on screen (red, navy, white/none).

As before: does Word open each (any repair prompt, and its words); a PDF; a save as filename-word365.docx.
LibreOffice 26.2 opened all twenty (PDF export, 24 October).
