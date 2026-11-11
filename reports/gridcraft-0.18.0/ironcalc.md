# ironcalc vs gridcraft@0.18.0 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 68. Fails where both answered an error, but a different one: 190. Fails where the case's own cell matches but its spill does not: 16.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 672 | 35 | 1 | 0 | 0 | 94.9% |
| DATABASE | 566 | 508 | 58 | 0 | 0 | 0 | 89.8% |
| DATE_AND_TIME | 2623 | 2532 | 88 | 0 | 0 | 3 | 96.6% |
| ENGINEERING | 2976 | 2952 | 24 | 0 | 0 | 0 | 99.2% |
| FINANCIAL | 3086 | 2757 | 329 | 0 | 0 | 0 | 89.3% |
| INFORMATION | 409 | 401 | 7 | 1 | 0 | 0 | 98.0% |
| LOGICAL | 533 | 518 | 9 | 0 | 0 | 6 | 98.3% |
| LOOKUP_AND_REFERENCE | 1047 | 932 | 114 | 1 | 0 | 0 | 89.0% |
| MATH_AND_TRIGONOMETRY | 56100 | 54972 | 1119 | 7 | 0 | 2 | 98.0% |
| STATISTICAL | 3184 | 2997 | 187 | 0 | 0 | 0 | 94.1% |
| TEXT | 1457 | 1403 | 54 | 0 | 0 | 0 | 96.3% |
| **total** | 72689 | 70644 | 2024 | 10 | 0 | 11 | 97.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 70644 | 2024 | 10 | 0 | 11 | 97.2% |
| **total** | 72689 | 70644 | 2024 | 10 | 0 | 11 | 97.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ACCRINT | 253 | fail 253 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, GridCraft 252.99999999999997 |
| CEILING.MATH | 126 | fail 125, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI139` =CEILING.MATH($A139,$47:$47) → Excel 8, GridCraft "#VALUE!" |
| CEILING | 124 | fail 123, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI44` =CEILING($A44,$47:$47) → Excel -1000, GridCraft "#VALUE!" |
| FLOOR.MATH | 122 | fail 121, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI127` =FLOOR.MATH($A127) → Excel "#VALUE!", GridCraft 45301 |
| FLOOR.PRECISE | 122 | fail 121, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AH83` =FLOOR.PRECISE($AM:$AM,AH$2) → Excel 0, GridCraft "#VALUE!" |
| ISO.CEILING | 122 | fail 121, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI208` =ISO.CEILING($AM:$AM) → Excel 0, GridCraft "#VALUE!" |
| CEILING.PRECISE | 122 | fail 121, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI83` =CEILING.PRECISE($A83) → Excel 0, GridCraft "#VALUE!" |
| FLOOR | 120 | fail 119, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG43` =FLOOR($A43,$AD$47:$AH$47) → Excel 0, GridCraft "#VALUE!" |
| COUNTIF | 84 | fail 84 | `STATISTICAL/COUNTIF.xlsx!AL10` =COUNTIF(AG:AG,CONCAT("<>", AH10)) → Excel 1048573, GridCraft 13 |
| LCM | 54 | fail 54 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!K31` =LCM({4,"2024-01-10",12}) → Excel 543612, GridCraft 12 |
| GCD | 45 | fail 45 | `MATH_AND_TRIGONOMETRY/GCD_LCM.xlsx!L44` =GCD({4,TRUE,12}) → Excel "#VALUE!", GridCraft 4 |
| XLOOKUP | 40 | fail 40 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!N1` =XLOOKUP(,L:L,M:M,,1,2) → Excel 6, GridCraft 1 |
| WEEKNUM | 31 | fail 31 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!U13` =WEEKNUM($A13) → Excel "#VALUE!", GridCraft 1 |
| COUNTIFS | 27 | fail 27 | `STATISTICAL/COUNTIFS.xlsx!O12` =COUNTIFS($J$2:$J$15, CONCAT(">=",K12)) → Excel 0, GridCraft 2 |
| TEXT | 25 | fail 25 | `TEXT/TEXT.xlsx!H20` =TEXT($A20,H$2) → Excel "-0.00000001", GridCraft "-1E-08" |
| BASE | 15 | fail 15 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!E112` =BASE(5,2,I109:I112) → Excel "000101", GridCraft "#VALUE!" |
| ODDFPRICE | 12 | fail 12 | `FINANCIAL/PRICE_YIELD.xlsx!J44` =ODDFPRICE(A44,B44,C44,D44,E44,F44,G44,H44,I44) → Excel 115.99277260461348, GridCraft 116.09059198938289 |
| ROW | 11 | fail 11 | `LOOKUP_AND_REFERENCE/ROW_COLUM.xlsx!K23` =ROW(INDIRECT(J23)) → Excel 77, GridCraft "#VALUE!" |
| MROUND | 11 | fail 11 | `MATH_AND_TRIGONOMETRY/MROUND_TRUNC_INT.xlsx!C34` =MROUND(,B34) → Excel "#N/A", GridCraft 0 |
| SUM | 10 | fail 9, unsupported 1 | `LOOKUP_AND_REFERENCE/INDIRECT.xlsx!H3` =SUM(INDIRECT("R[-4]", FALSE)) → Excel 41, GridCraft "#REF!" |
| COLUMN | 10 | fail 10 | `LOOKUP_AND_REFERENCE/ROW_COLUM.xlsx!L23` =COLUMN(INDIRECT(J23)) → Excel 513, GridCraft 513 |
| CUMPRINC | 9 | fail 9 | `FINANCIAL/CUMPRINC_CUMIPMT.xlsx!H27` =CUMPRINC(A27,B27,C27,D27,E27,F27) → Excel "#VALUE!", GridCraft -1942.984346059738 |
| XNPV | 9 | fail 9 | `FINANCIAL/XIRR.xlsx!Q20` =XNPV(O20,A20:F20,H20:M20) → Excel "#NUM!", GridCraft "#VALUE!" |
| BETA.INV | 9 | fail 9 | `STATISTICAL/BETADIST_BETAINV.xlsx!H69` =BETA.INV(A69,B69,C69,E69,F69) → Excel "#NUM!", GridCraft 1 |
| NEGBINOM.DIST | 9 | fail 9 | `STATISTICAL/BINOM.xlsx!E73` =NEGBINOM.DIST(A73,B73,C73,D73) → Excel "#NUM!", GridCraft 0 |
