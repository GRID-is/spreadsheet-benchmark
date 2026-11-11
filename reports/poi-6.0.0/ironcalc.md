# ironcalc vs poi@6.0.0 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 600. Fails where both answered an error, but a different one: 922. Fails where the case's own cell matches but its spill does not: 16.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 679 | 9 | 20 | 0 | 0 | 95.9% |
| DATABASE | 566 | 416 | 146 | 4 | 0 | 0 | 73.5% |
| DATE_AND_TIME | 2623 | 2151 | 303 | 166 | 0 | 3 | 82.1% |
| ENGINEERING | 2976 | 538 | 105 | 2333 | 0 | 0 | 18.1% |
| FINANCIAL | 3086 | 1553 | 121 | 1412 | 0 | 0 | 50.3% |
| INFORMATION | 409 | 261 | 52 | 96 | 0 | 0 | 63.8% |
| LOGICAL | 533 | 386 | 28 | 113 | 0 | 6 | 73.2% |
| LOOKUP_AND_REFERENCE | 1047 | 677 | 295 | 75 | 0 | 0 | 64.7% |
| MATH_AND_TRIGONOMETRY | 56100 | 28539 | 5283 | 22276 | 0 | 2 | 50.9% |
| STATISTICAL | 3184 | 1303 | 406 | 1475 | 0 | 0 | 40.9% |
| TEXT | 1457 | 989 | 201 | 267 | 0 | 0 | 67.9% |
| **total** | 72689 | 37492 | 6949 | 28237 | 0 | 11 | 51.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 37492 | 6949 | 28237 | 0 | 11 | 51.6% |
| **total** | 72689 | 37492 | 6949 | 28237 | 0 | 11 | 51.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ARABIC | 20040 | unsupported 20040 | `MATH_AND_TRIGONOMETRY/ARABIC_ROMAN.xlsx!F3835` =ARABIC(ROMAN!F3830) → Excel 3790, Apache POI "#NAME?" |
| ROMAN | 4079 | fail 4079 | `MATH_AND_TRIGONOMETRY/ARABIC_ROMAN.xlsx!O33` =ROMAN($A33,O$1) → Excel "#N/A", Apache POI "#VALUE!" |
| ISO.CEILING | 1396 | unsupported 1396 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI217` =ISO.CEILING($A217) → Excel 7, Apache POI "#NAME?" |
| ACCRINT | 587 | unsupported 587 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, Apache POI "#NAME?" |
| IMREAL | 416 | fail 23, unsupported 393 | `ENGINEERING/COMPLEXs.xlsx!AL26` =IMREAL(IMSQRT(D26)) → Excel 6.1257422745431e-17, Apache POI "#NAME?" |
| IMAGINARY | 416 | fail 23, unsupported 393 | `ENGINEERING/COMPLEXs.xlsx!AM27` =IMAGINARY(IMSQRT(D27)) → Excel 0.707106781186547, Apache POI "#NAME?" |
| CONVERT | 357 | unsupported 357 | `ENGINEERING/CONVERT.xlsx!G14` =CONVERT(0.05,F14,$B$2) → Excel "#N/A", Apache POI "#NAME?" |
| BASE | 250 | unsupported 250 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!E112` =BASE(5,2,I109:I112) → Excel "000101", Apache POI "#NAME?" |
| XLOOKUP | 206 | fail 206 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!I6` =XLOOKUP(E2,A:A,B:B,,,TRUE) → Excel "+99", Apache POI "#VALUE!" |
| CEILING.MATH | 190 | fail 190 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AF138` =CEILING.MATH($A138,AF$2) → Excel 0, Apache POI "#VALUE!" |
| FLOOR.MATH | 186 | fail 186 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI126` =FLOOR.MATH($A126) → Excel 45301, Apache POI "#VALUE!" |
| TEXT | 135 | fail 134, unsupported 1 | `INFORMATION/error_broadcast.xlsx!E53` =TEXT(A1:C1, NA()) → Excel "#N/A", Apache POI "#VALUE!" |
| CEILING | 135 | fail 135 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AF45` =CEILING($A45,AF$2) → Excel 45301, Apache POI "#VALUE!" |
| FLOOR.PRECISE | 124 | fail 124 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI79` =FLOOR.PRECISE($A79) → Excel 45301, Apache POI "#VALUE!" |
| CEILING.PRECISE | 124 | fail 124 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI79` =CEILING.PRECISE($A79) → Excel 45301, Apache POI "#VALUE!" |
| FLOOR | 122 | fail 122 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AH32` =FLOOR($A32,AH$2) → Excel 45298.62447211122, Apache POI "#VALUE!" |
| BESSELK | 119 | unsupported 119 | `ENGINEERING/BESSEL.xlsx!F18` =BESSELK(A18, 10) → Excel "#NUM!", Apache POI "#NAME?" |
| BESSELY | 119 | unsupported 119 | `ENGINEERING/BESSEL.xlsx!F18` =BESSELY(A18, 10) → Excel "#NUM!", Apache POI "#NAME?" |
| BESSELI | 118 | unsupported 118 | `ENGINEERING/BESSEL.xlsx!F18` =BESSELI(A18, 10) → Excel 0.0000019435535412095194, Apache POI "#NAME?" |
| XOR | 116 | unsupported 116 | `LOGICAL/AND_OR_XOR.xlsx!D103` =XOR(A8) → Excel "#VALUE!", Apache POI "#NAME?" |
| COUNTIF | 112 | fail 112 | `STATISTICAL/COUNTIF.xlsx!AK12` =COUNTIF(AG:AG,AH12) → Excel 1, Apache POI 0 |
| WEEKNUM | 108 | fail 108 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!U12` =WEEKNUM($A12) → Excel "#N/A", Apache POI "#VALUE!" |
| NETWORKDAYS.INTL | 104 | unsupported 104 | `DATE_AND_TIME/NETWORKDAYS_NETWORKDAYS.INTL.xlsx!I36` =NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,12,31),J36) → Excel 105, Apache POI "#NAME?" |
| LCM | 101 | fail 101 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!K31` =LCM({4,"2024-01-10",12}) → Excel 543612, Apache POI "#VALUE!" |
| DECIMAL | 100 | unsupported 100 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!C86` =DECIMAL(G83:G86,H83:H86) → Excel 35, Apache POI "#NAME?" |
