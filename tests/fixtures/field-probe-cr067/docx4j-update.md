CR-067 phase 1 field probe

Select all, press F9, then save as described in README.txt. Each case: its token, then the field. A field Word leaves alone still says OLD.

A bookmark: **Bookmarked text**. A footnote reference: [^1].

## Case 1: P01\_CREATEDATE { CREATEDATE }

P01\_CREATEDATE = 2/3/2024 3:05:00 PM

## Case 2: P02\_CREATEDATE\_SW { CREATEDATE \\@ "yyyy-MM-dd HH:mm" }

P02\_CREATEDATE\_SW = 2024-02-03 15:05

## Case 3: P03\_SAVEDATE { SAVEDATE }

P03\_SAVEDATE = 6/7/2025 6:09:00 PM

## Case 4: P04\_SAVEDATE\_SW { SAVEDATE \\@ "d MMMM yyyy" }

P04\_SAVEDATE\_SW = 7 June 2025

## Case 5: P05\_REVNUM { REVNUM }

P05\_REVNUM = 7

## Case 6: P06\_INFO\_TITLE { INFO Title }

P06\_INFO\_TITLE = Probe Title

## Case 7: P07\_INFO\_AUTHOR { INFO Author }

P07\_INFO\_AUTHOR = Probe Author

## Case 8: P08\_INFO\_NOARG { INFO }

P08\_INFO\_NOARG = OLD

## Case 9: P09\_INFO\_NOPE { INFO Nope }

P09\_INFO\_NOPE = OLD

## Case 10: P10\_SECTION { SECTION }

P10\_SECTION = 1

## Case 11: P11\_QUOTE { QUOTE "hello world" }

P11\_QUOTE = hello world

## Case 12: P12\_QUOTE\_NOARG { QUOTE }

P12\_QUOTE\_NOARG = 

## Case 13: P13\_REF { REF bm1 }

P13\_REF = Bookmarked text

## Case 14: P14\_REF\_UPPER { REF bm1 \\\* Upper }

P14\_REF\_UPPER = BOOKMARKED TEXT

## Case 15: P15\_REF\_NOPE { REF nope }

P15\_REF\_NOPE = OLD

## Case 16: P16\_REF\_H { REF bm1 \\h }

P16\_REF\_H = Bookmarked text

## Case 17: P17\_NOTEREF { NOTEREF bmNote }

P17\_NOTEREF = 1

## Case 18: P18\_NOTEREF\_F { NOTEREF bmNote \\f }

P18\_NOTEREF\_F = 1

## Case 19: P19\_NOTEREF\_NOPE { NOTEREF nope }

P19\_NOTEREF\_NOPE = OLD

## Case 20: P20\_USERNAME { USERNAME }

P20\_USERNAME = OLDUSER

## Case 21: P21\_USERINITIALS { USERINITIALS }

P21\_USERINITIALS = OLDUSER

## Case 22: P22\_USERADDRESS { USERADDRESS }

P22\_USERADDRESS = OLDUSER

## Case 23: P23\_KEYWORDS { KEYWORDS }

P23\_KEYWORDS = alpha, beta

## Case 24: P24\_SECTION2 { SECTION }

P24\_SECTION2 = 2

## Case 25: P25\_CREATEDATE\_SIMPLE { CREATEDATE \\@ "yyyy-MM-dd" }

P25\_CREATEDATE\_SIMPLE = 2024-02-03

[^1]: The probe's footnote.
