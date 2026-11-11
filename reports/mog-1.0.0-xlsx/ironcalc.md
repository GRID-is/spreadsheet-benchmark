# ironcalc vs mog@1.0.0 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 211. Fails where both answered an error, but a different one: 212. Fails where the case's own cell matches but its spill does not: 22.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 698 | 9 | 1 | 0 | 0 | 98.6% |
| DATABASE | 566 | 479 | 87 | 0 | 0 | 0 | 84.6% |
| DATE_AND_TIME | 2623 | 2482 | 138 | 0 | 0 | 3 | 94.7% |
| ENGINEERING | 2976 | 2215 | 761 | 0 | 0 | 0 | 74.4% |
| FINANCIAL | 3086 | 2527 | 559 | 0 | 0 | 0 | 81.9% |
| INFORMATION | 409 | 386 | 22 | 1 | 0 | 0 | 94.4% |
| LOGICAL | 533 | 512 | 15 | 0 | 0 | 6 | 97.2% |
| LOOKUP_AND_REFERENCE | 1047 | 957 | 90 | 0 | 0 | 0 | 91.4% |
| MATH_AND_TRIGONOMETRY | 56100 | 54378 | 1644 | 76 | 0 | 2 | 96.9% |
| STATISTICAL | 3184 | 2664 | 520 | 0 | 0 | 0 | 83.7% |
| TEXT | 1457 | 1339 | 118 | 0 | 0 | 0 | 91.9% |
| **total** | 72689 | 68637 | 3963 | 78 | 0 | 11 | 94.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 68637 | 3963 | 78 | 0 | 11 | 94.4% |
| **total** | 72689 | 68637 | 3963 | 78 | 0 | 11 | 94.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ACCRINT | 373 | fail 373 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, Mog 253 |
| IMREAL | 256 | fail 256 | `ENGINEERING/COMPLEXs.xlsx!AL26` =IMREAL(IMSQRT(D26)) → Excel 6.1257422745431e-17, Mog "#VALUE!" |
| IMAGINARY | 256 | fail 256 | `ENGINEERING/COMPLEXs.xlsx!AM26` =IMAGINARY(IMSQRT(D26)) → Excel 1, Mog "#VALUE!" |
| COUNTIF | 256 | fail 256 | `MATH_AND_TRIGONOMETRY/SUMIF_AVERAGE_IF.xlsx!G9` =COUNTIF($B$2:$B$10,D9) → Excel 0, Mog 1 |
| ROMAN | 226 | fail 226 | `MATH_AND_TRIGONOMETRY/ARABIC_ROMAN.xlsx!F3789` =ROMAN($A3789,4) → Excel "MMMDCCIL", Mog "MMMDCCVLIV" |
| FLOOR.MATH | 194 | fail 193, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI127` =FLOOR.MATH($A127) → Excel "#VALUE!", Mog 45301 |
| CEILING.MATH | 168 | fail 167, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI139` =CEILING.MATH($A139,$47:$47) → Excel 8, Mog 0 |
| FLOOR.PRECISE | 144 | fail 143, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI83` =FLOOR.PRECISE(AM:AM) → Excel 3, Mog 0 |
| ISO.CEILING | 141 | fail 140, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG214` =ISO.CEILING($A214,$47:$47) → Excel -744, Mog 0 |
| CEILING.PRECISE | 141 | fail 140, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG87` =CEILING.PRECISE($A87,$47:$47) → Excel 710, Mog 0 |
| CONVERT | 129 | fail 129 | `ENGINEERING/CONVERT.xlsx!G14` =CONVERT(0.05,F14,$B$2) → Excel "#N/A", Mog 50802345440.00001 |
| CEILING | 128 | fail 122, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI44` =CEILING($A44,$47:$47) → Excel -1000, Mog 0 |
| FLOOR | 127 | fail 121, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG45` =FLOOR($A45,$AD$47:$AH$47) → Excel 4, Mog 7 |
| COUNTIFS | 114 | fail 114 | `STATISTICAL/COUNTIFS.xlsx!M15` =COUNTIFS($I:$I,J15,$I:$I,K15) → Excel 0, Mog 8 |
| LOG10 | 60 | fail 1, unsupported 59 | `MATH_AND_TRIGONOMETRY/LOG_LOG10_LN.xlsx!D68` =LOG10(A68) → Excel -0.3010299956639812, Mog "#NAME?" |
| TEXT | 55 | fail 55 | `TEXT/TEXT.xlsx!K12` =TEXT($A12,K$2) → Excel "1234560000000000000000.00%", Mog "1.23E21" |
| WEEKNUM | 38 | fail 38 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!U13` =WEEKNUM($A13) → Excel "#VALUE!", Mog 1 |
| XLOOKUP | 33 | fail 33 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!I3` =XLOOKUP(E2,A:A,3) → Excel "#VALUE!", Mog 0 |
| SUMIF | 25 | fail 25 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!E12` =SUMIF($B$2:$B$20, C12,$A$2:$A$20) → Excel 77, Mog 0 |
| BASE | 24 | fail 24 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!E112` =BASE(5,2,I109:I112) → Excel "000101", Mog "0101" |
| MAXIFS | 24 | fail 24 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!H32` =MAXIFS($A$25:$A$40,$B$25:$B$40,C32,$B$25:$B$40,D32) → Excel 2, Mog 0 |
| MINIFS | 24 | fail 24 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!G32` =MINIFS($A$25:$A$40,$B$25:$B$40,C32,$B$25:$B$40,D32) → Excel 2, Mog 0 |
| SUMIFS | 24 | fail 24 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!E32` =SUMIFS($A$25:$A$40,$B$25:$B$40,C32,$B$25:$B$40,D32) → Excel 2, Mog 0 |
| LCM | 23 | fail 23 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!K44` =LCM({4,TRUE,12}) → Excel "#VALUE!", Mog 12 |
| GCD | 21 | fail 21 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!L44` =GCD({4,TRUE,12}) → Excel "#VALUE!", Mog 4 |
