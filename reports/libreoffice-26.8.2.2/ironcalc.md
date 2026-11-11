# ironcalc vs libreoffice@26.8.2.2 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 866. Fails where both answered an error, but a different one: 2780. Fails where the case's own cell matches but its spill does not: 27.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 685 | 16 | 7 | 0 | 0 | 96.8% |
| DATABASE | 566 | 401 | 165 | 0 | 0 | 0 | 70.8% |
| DATE_AND_TIME | 2623 | 2176 | 429 | 15 | 0 | 3 | 83.1% |
| ENGINEERING | 2976 | 1624 | 936 | 416 | 0 | 0 | 54.6% |
| FINANCIAL | 3086 | 2103 | 983 | 0 | 0 | 0 | 68.1% |
| INFORMATION | 409 | 335 | 73 | 1 | 0 | 0 | 81.9% |
| LOGICAL | 533 | 464 | 53 | 10 | 0 | 6 | 88.0% |
| LOOKUP_AND_REFERENCE | 1047 | 908 | 134 | 5 | 0 | 0 | 86.7% |
| MATH_AND_TRIGONOMETRY | 56100 | 53980 | 2043 | 75 | 0 | 2 | 96.2% |
| STATISTICAL | 3184 | 2151 | 975 | 58 | 0 | 0 | 67.6% |
| TEXT | 1457 | 1069 | 364 | 24 | 0 | 0 | 73.4% |
| **total** | 72689 | 65896 | 6171 | 611 | 0 | 11 | 90.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 65896 | 6171 | 611 | 0 | 11 | 90.7% |
| **total** | 72689 | 65896 | 6171 | 611 | 0 | 11 | 90.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ACCRINT | 519 | fail 519 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, LibreOffice 253 |
| FLOOR | 370 | fail 360, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG33` =FLOOR($A33,$AD$47:$AH$47) → Excel "#VALUE!", LibreOffice 45300 |
| IMREAL | 334 | fail 126, unsupported 208 | `ENGINEERING/COMPLEXs.xlsx!AL26` =IMREAL(IMSQRT(D26)) → Excel 6.1257422745431e-17, LibreOffice "#VALUE!" |
| IMAGINARY | 334 | fail 126, unsupported 208 | `ENGINEERING/COMPLEXs.xlsx!AM26` =IMAGINARY(IMSQRT(D26)) → Excel 1, LibreOffice "#VALUE!" |
| COUNTIF | 297 | fail 297 | `MATH_AND_TRIGONOMETRY/SUMIF_AVERAGE_IF.xlsx!G6` =COUNTIF($B$2:$B$10,D6) → Excel 1, LibreOffice "#N/A" |
| CEILING | 284 | fail 274, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG43` =CEILING($A43,AG$2) → Excel "#VALUE!", LibreOffice 0 |
| CEILING.MATH | 172 | fail 162, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG136` =CEILING.MATH($A136,AG$2) → Excel "#VALUE!", LibreOffice 0 |
| FLOOR.MATH | 168 | fail 158, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI127` =FLOOR.MATH($A127) → Excel "#VALUE!", LibreOffice 45301 |
| FLOOR.PRECISE | 168 | fail 158, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI80` =FLOOR.PRECISE($A80) → Excel "#VALUE!", LibreOffice 45301 |
| ISO.CEILING | 168 | fail 158, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI205` =ISO.CEILING($A205) → Excel "#VALUE!", LibreOffice 45301 |
| CEILING.PRECISE | 168 | fail 158, unsupported 10 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI80` =CEILING.PRECISE($A80) → Excel "#VALUE!", LibreOffice 45301 |
| CONVERT | 160 | fail 160 | `ENGINEERING/CONVERT.xlsx!G14` =CONVERT(0.05,F14,$B$2) → Excel "#N/A", LibreOffice "#VALUE!" |
| COUNTIFS | 149 | fail 149 | `STATISTICAL/COUNTIFS.xlsx!M15` =COUNTIFS($I:$I,J15,$I:$I,K15) → Excel 0, LibreOffice 1 |
| WEEKNUM | 142 | fail 142 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!U8` =WEEKNUM($A8) → Excel 53, LibreOffice 1 |
| CHAR | 128 | fail 128 | `TEXT/CHAR.xlsx!B177` =CHAR(A177) → Excel "±", LibreOffice "�" |
| TEXT | 125 | fail 125 | `INFORMATION/error_broadcast.xlsx!E54` =TEXT(A6:C6, NA()) → Excel "#N/A", LibreOffice "#N/A" |
| LCM | 87 | fail 86, unsupported 1 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!K31` =LCM({4,"2024-01-10",12}) → Excel 543612, LibreOffice "#VALUE!" |
| XLOOKUP | 83 | fail 83 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!V14` =XLOOKUP(T14,R:R,S:S,"Not fount :)",,-1) → Excel 0, LibreOffice "Not fount :)" |
| BASE | 69 | fail 69 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!E84` =BASE(A84,B84,C84) → Excel "#NUM!", LibreOffice "#VALUE!" |
| WEEKDAY | 64 | fail 64 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!R12` =WEEKDAY($A12,R$1) → Excel "#NUM!", LibreOffice 6 |
| BESSELK | 56 | fail 56 | `ENGINEERING/BESSEL.xlsx!F18` =BESSELK(A18, 10) → Excel "#NUM!", LibreOffice "#VALUE!" |
| BESSELY | 56 | fail 56 | `ENGINEERING/BESSEL.xlsx!F18` =BESSELY(A18, 10) → Excel "#NUM!", LibreOffice "#VALUE!" |
| GCD | 49 | fail 48, unsupported 1 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!L44` =GCD({4,TRUE,12}) → Excel "#VALUE!", LibreOffice 1 |
| BINOM.DIST.RANGE | 48 | unsupported 48 | `STATISTICAL/BINOM.xlsx!E43` =BINOM.DIST.RANGE(B43,C43,A43,D43) → Excel "#VALUE!", LibreOffice "#NAME?" |
| BIN2OCT | 43 | fail 43 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!G26` =BIN2OCT(A26) → Excel "#NUM!", LibreOffice "#VALUE!" |
