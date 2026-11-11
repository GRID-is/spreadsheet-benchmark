# ironcalc vs spread-sheets@19.2.5 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 61. Fails where both answered an error, but a different one: 394. Fails where the case's own cell matches but its spill does not: 46.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 663 | 43 | 2 | 0 | 0 | 93.6% |
| DATABASE | 566 | 477 | 89 | 0 | 0 | 0 | 84.3% |
| DATE_AND_TIME | 2623 | 2213 | 194 | 0 | 213 | 3 | 84.5% |
| ENGINEERING | 2976 | 2267 | 709 | 0 | 0 | 0 | 76.2% |
| FINANCIAL | 3086 | 2533 | 553 | 0 | 0 | 0 | 82.1% |
| INFORMATION | 409 | 335 | 48 | 26 | 0 | 0 | 81.9% |
| LOGICAL | 533 | 511 | 16 | 0 | 0 | 6 | 97.0% |
| LOOKUP_AND_REFERENCE | 1047 | 815 | 128 | 0 | 104 | 0 | 77.8% |
| MATH_AND_TRIGONOMETRY | 56100 | 54797 | 1258 | 43 | 0 | 2 | 97.7% |
| STATISTICAL | 3184 | 2524 | 597 | 0 | 63 | 0 | 79.3% |
| TEXT | 1457 | 1311 | 146 | 0 | 0 | 0 | 90.0% |
| **total** | 72689 | 68446 | 3781 | 71 | 380 | 11 | 94.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 68446 | 3781 | 71 | 380 | 11 | 94.2% |
| **total** | 72689 | 68446 | 3781 | 71 | 380 | 11 | 94.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ACCRINT | 335 | fail 335 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, SpreadJS 248.852459016393 |
| IMREAL | 256 | fail 256 | `ENGINEERING/COMPLEXs.xlsx!AL26` =IMREAL(IMSQRT(D26)) → Excel 6.1257422745431e-17, SpreadJS "#VALUE!" |
| IMAGINARY | 256 | fail 256 | `ENGINEERING/COMPLEXs.xlsx!AM26` =IMAGINARY(IMSQRT(D26)) → Excel 1, SpreadJS "#VALUE!" |
| COUNTIF | 241 | fail 241 | `STATISTICAL/COUNTIF.xlsx!AK12` =COUNTIF(AG:AG,AH12) → Excel 1, SpreadJS 3 |
| COUNTIFS | 134 | fail 134 | `STATISTICAL/COUNTIFS.xlsx!M15` =COUNTIFS($I:$I,J15,$I:$I,K15) → Excel 0, SpreadJS 1 |
| CEILING | 126 | fail 120, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI44` =CEILING($A44,$47:$47) → Excel -1000, SpreadJS "#SPILL!" |
| FLOOR.PRECISE | 115 | fail 109, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AH83` =FLOOR.PRECISE($AM:$AM,AH$2) → Excel 0, SpreadJS "#SPILL!" |
| ISO.CEILING | 115 | fail 109, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI208` =ISO.CEILING($AM:$AM) → Excel 0, SpreadJS "#SPILL!" |
| CEILING.PRECISE | 115 | fail 109, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI83` =CEILING.PRECISE($A83) → Excel 0, SpreadJS "#SPILL!" |
| FLOOR | 114 | fail 108, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG43` =FLOOR($A43,$AD$47:$AH$47) → Excel 0, SpreadJS "#SPILL!" |
| CEILING.MATH | 108 | fail 102, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI139` =CEILING.MATH($A139,$47:$47) → Excel 8, SpreadJS "#SPILL!" |
| BASE | 98 | fail 98 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!E112` =BASE(5,2,I109:I112) → Excel "000101", SpreadJS "#SPILL!" |
| WORKDAY.INTL | 93 | error 93 | `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx!I44` =WORKDAY.INTL(DATE(2025,1,10),J44,17,K44) → Excel 45672, SpreadJS null |
| TEXT | 90 | fail 90 | `INFORMATION/error_broadcast.xlsx!E53` =TEXT(A1:C1, NA()) → Excel "#N/A", SpreadJS "#N/A" |
| FLOOR.MATH | 90 | fail 84, unsupported 6 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI130` =FLOOR.MATH($AM:$AM) → Excel 3, SpreadJS "#SPILL!" |
| LCM | 73 | fail 73 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!K31` =LCM({4,"2024-01-10",12}) → Excel 543612, SpreadJS 12 |
| XLOOKUP | 72 | fail 72 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!H6` =XLOOKUP(E2,A:A,B:B,,12) → Excel "#VALUE!", SpreadJS "+99" |
| WORKDAY | 67 | error 67 | `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx!H46` =WORKDAY(DATE(2025,1,10),J46,K46) → Excel 45673, SpreadJS null |
| POISSON.DIST | 60 | error 60 | `STATISTICAL/POISSON.xlsx!D54` =POISSON.DIST(A54,B54,C54) → Excel "#VALUE!", SpreadJS null |
| DATE | 55 | fail 10, error 45 | `DATE_AND_TIME/DATE.xlsx!D45` =DATE(A45,B45,C45) → Excel "#NUM!", SpreadJS 367 |
| WEEKNUM | 43 | fail 43 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!U13` =WEEKNUM($A13) → Excel "#VALUE!", SpreadJS 53 |
| INDIRECT | 36 | fail 1, error 35 | `LOOKUP_AND_REFERENCE/INDIRECT.xlsx!E3` =INDIRECT("B2", 1) → Excel 5, SpreadJS null |
| NETWORKDAYS.INTL | 34 | fail 34 | `DATE_AND_TIME/NETWORKDAYS_NETWORKDAYS.INTL.xlsx!I36` =NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,12,31),J36) → Excel 105, SpreadJS "#VALUE!" |
| SUM | 28 | fail 14, unsupported 1, error 13 | `LOOKUP_AND_REFERENCE/INDIRECT.xlsx!E4` =SUM(INDIRECT("A1:C2", TRUE)) → Excel 21, SpreadJS null |
| CHAR | 27 | fail 27 | `TEXT/CHAR.xlsx!B149` =CHAR(A149) → Excel "•", SpreadJS "" |

## Errors

- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `DATE_AND_TIME/WORKDAY_WORKDAY.INTL.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
- `STATISTICAL/POISSON.xlsx`: timeout after 60000 ms
