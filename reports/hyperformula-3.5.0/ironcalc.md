# ironcalc vs hyperformula@3.5.0 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 87. Fails where both answered an error, but a different one: 391. Fails where the case's own cell matches but its spill does not: 19.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 482 | 133 | 93 | 0 | 0 | 68.1% |
| DATABASE | 566 | 441 | 122 | 3 | 0 | 0 | 77.9% |
| DATE_AND_TIME | 2623 | 2334 | 282 | 4 | 0 | 3 | 89.1% |
| ENGINEERING | 2976 | 1871 | 680 | 425 | 0 | 0 | 62.9% |
| FINANCIAL | 3086 | 2027 | 145 | 914 | 0 | 0 | 65.7% |
| INFORMATION | 409 | 281 | 75 | 53 | 0 | 0 | 68.7% |
| LOGICAL | 533 | 265 | 76 | 186 | 0 | 6 | 50.3% |
| LOOKUP_AND_REFERENCE | 1047 | 483 | 259 | 201 | 104 | 0 | 46.1% |
| MATH_AND_TRIGONOMETRY | 56100 | 54041 | 1656 | 401 | 0 | 2 | 96.3% |
| STATISTICAL | 3184 | 2171 | 260 | 753 | 0 | 0 | 68.2% |
| TEXT | 1457 | 814 | 356 | 270 | 17 | 0 | 55.9% |
| **total** | 72689 | 65210 | 4044 | 3303 | 121 | 11 | 89.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 65210 | 4044 | 3303 | 121 | 11 | 89.7% |
| **total** | 72689 | 65210 | 4044 | 3303 | 121 | 11 | 89.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ACCRINT | 587 | unsupported 587 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, HyperFormula "#NAME?" |
| COUNTIF | 507 | fail 120, unsupported 387 | `MATH_AND_TRIGONOMETRY/SUMIF_AVERAGE_IF.xlsx!G9` =COUNTIF($B$2:$B$10,D9) → Excel 0, HyperFormula "#VALUE!" |
| CONVERT | 357 | unsupported 357 | `ENGINEERING/CONVERT.xlsx!G14` =CONVERT(0.05,F14,$B$2) → Excel "#N/A", HyperFormula "#NAME?" |
| CEILING.MATH | 290 | fail 163, unsupported 127 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AF138` =CEILING.MATH($A138,AF$2) → Excel 0, HyperFormula "#VALUE!" |
| FLOOR.MATH | 288 | fail 161, unsupported 127 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI126` =FLOOR.MATH($A126) → Excel 45301, HyperFormula "#VALUE!" |
| TEXT | 282 | fail 281, unsupported 1 | `INFORMATION/error_broadcast.xlsx!A1` =TEXT("x", NA()) → Excel "#N/A", HyperFormula "#VALUE!" |
| IMREAL | 270 | fail 270 | `ENGINEERING/COMPLEXs.xlsx!AL26` =IMREAL(IMSQRT(D26)) → Excel 6.1257422745431e-17, HyperFormula "#VALUE!" |
| IMAGINARY | 270 | fail 270 | `ENGINEERING/COMPLEXs.xlsx!AM26` =IMAGINARY(IMSQRT(D26)) → Excel 1, HyperFormula "#VALUE!" |
| CEILING | 227 | fail 226, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AF45` =CEILING($A45,AF$2) → Excel 45301, HyperFormula "#VALUE!" |
| COUNTIFS | 206 | fail 56, unsupported 150 | `STATISTICAL/COUNTIFS.xlsx!M15` =COUNTIFS($I:$I,J15,$I:$I,K15) → Excel 0, HyperFormula "#VALUE!" |
| XLOOKUP | 129 | fail 126, unsupported 3 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!I6` =XLOOKUP(E2,A:A,B:B,,,TRUE) → Excel "+99", HyperFormula "#NAME?" |
| FLOOR.PRECISE | 120 | fail 119, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI79` =FLOOR.PRECISE($A79) → Excel 45301, HyperFormula "#VALUE!" |
| ISO.CEILING | 120 | fail 119, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI204` =ISO.CEILING($A204) → Excel 45301, HyperFormula "#VALUE!" |
| CEILING.PRECISE | 120 | fail 119, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI79` =CEILING.PRECISE($A79) → Excel 45301, HyperFormula "#VALUE!" |
| FLOOR | 118 | fail 117, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AH32` =FLOOR($A32,AH$2) → Excel 45298.62447211122, HyperFormula "#VALUE!" |
| ROMAN | 90 | fail 90 | `MATH_AND_TRIGONOMETRY/ARABIC_ROMAN.xlsx!F4040` =ROMAN($A4040,4) → Excel "#VALUE!", HyperFormula "#NUM!" |
| SUM | 77 | fail 33, unsupported 31, error 13 | `INFORMATION/ISREF.xlsx!D14` =SUM(FILTER(L1:L2, M1:M2, L1:L2)) → Excel 30, HyperFormula "#N/A" |
| TEXTAFTER | 71 | unsupported 71 | `TEXT/TEXTBEFORE_TEXTAFTER.xlsx!G40` =TEXTAFTER(A40,B40,C40,D40,E40,F40) → Excel "#VALUE!", HyperFormula "#NAME?" |
| TEXTBEFORE | 70 | unsupported 70 | `TEXT/TEXTBEFORE_TEXTAFTER.xlsx!J39` =TEXTBEFORE(A39,B39) → Excel "#N/A", HyperFormula "#NAME?" |
| AND | 68 | fail 18, unsupported 50 | `LOGICAL/AND_OR_XOR.xlsx!E108` =AND(A10:B10) → Excel "#VALUE!", HyperFormula false |
| XOR | 68 | fail 17, unsupported 51 | `LOGICAL/AND_OR_XOR.xlsx!D103` =XOR(A8) → Excel "#VALUE!", HyperFormula false |
| OR | 65 | fail 16, unsupported 49 | `LOGICAL/AND_OR_XOR.xlsx!C106` =OR(A9:B9) → Excel "#VALUE!", HyperFormula false |
| LCM | 65 | fail 64, unsupported 1 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!K31` =LCM({4,"2024-01-10",12}) → Excel 543612, HyperFormula 12 |
| IFERROR | 63 | fail 63 | `INFORMATION/error_broadcast.xlsx!E67` =IFERROR(A1:C1-AVERAGE(K1:K3), 0) → Excel 0, HyperFormula 0 |
| AVERAGEIFS | 50 | unsupported 50 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!F15` =AVERAGEIFS($A$2:$A$20,$B$2:$B$20,C15) → Excel "#DIV/0!", HyperFormula "#NAME?" |

## Errors

- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: Sheet size limit exceeded
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
- `TEXT/REPT.xlsx`: Invalid string length
