# ironcalc vs gnumeric@1.12.62 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 60. Fails where both answered an error, but a different one: 337. Fails where the case's own cell matches but its spill does not: 9.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 695 | 6 | 7 | 0 | 0 | 98.2% |
| DATABASE | 566 | 382 | 184 | 0 | 0 | 0 | 67.5% |
| DATE_AND_TIME | 2623 | 2239 | 371 | 10 | 0 | 3 | 85.5% |
| ENGINEERING | 2976 | 1896 | 370 | 0 | 710 | 0 | 63.7% |
| FINANCIAL | 3086 | 2554 | 515 | 17 | 0 | 0 | 82.8% |
| INFORMATION | 409 | 348 | 48 | 13 | 0 | 0 | 85.1% |
| LOGICAL | 533 | 470 | 52 | 5 | 0 | 6 | 89.2% |
| LOOKUP_AND_REFERENCE | 1047 | 806 | 175 | 66 | 0 | 0 | 77.0% |
| MATH_AND_TRIGONOMETRY | 56100 | 47357 | 378 | 7973 | 390 | 2 | 84.4% |
| STATISTICAL | 3184 | 2580 | 384 | 220 | 0 | 0 | 81.0% |
| TEXT | 1457 | 1270 | 146 | 24 | 17 | 0 | 87.2% |
| **total** | 72689 | 60597 | 2629 | 8335 | 1117 | 11 | 83.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 60597 | 2629 | 8335 | 1117 | 11 | 83.4% |
| **total** | 72689 | 60597 | 2629 | 8335 | 1117 | 11 | 83.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| FLOOR.MATH | 1900 | unsupported 1900 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI134` =FLOOR.MATH($A134) → Excel 709, Gnumeric "#NAME?" |
| CEILING.MATH | 1858 | unsupported 1858 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AK138` =CEILING.MATH($A138,AK$2) → Excel 0, Gnumeric "#NAME?" |
| FLOOR.PRECISE | 1396 | unsupported 1396 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI89` =FLOOR.PRECISE($A89) → Excel -745, Gnumeric "#NAME?" |
| ISO.CEILING | 1396 | unsupported 1396 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI217` =ISO.CEILING($A217) → Excel 7, Gnumeric "#NAME?" |
| CEILING.PRECISE | 1396 | unsupported 1396 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI92` =CEILING.PRECISE($A92) → Excel 7, Gnumeric "#NAME?" |
| ACCRINT | 288 | fail 288 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, Gnumeric 253 |
| BASE | 250 | error 250 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!E112` =BASE(5,2,I109:I112) → Excel "000101", Gnumeric null |
| COUNTIF | 197 | fail 197 | `MATH_AND_TRIGONOMETRY/SUMIF_AVERAGE_IF.xlsx!G6` =COUNTIF($B$2:$B$10,D6) → Excel 1, Gnumeric "#N/A" |
| CONVERT | 163 | fail 163 | `ENGINEERING/CONVERT.xlsx!G14` =CONVERT(0.05,F14,$B$2) → Excel "#N/A", Gnumeric 50802345440 |
| XLOOKUP | 140 | fail 140 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!H6` =XLOOKUP(E2,A:A,B:B,,12) → Excel "#VALUE!", Gnumeric "+99" |
| DECIMAL | 100 | error 100 | `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx!C86` =DECIMAL(G83:G86,H83:H86) → Excel 35, Gnumeric null |
| WEEKNUM | 96 | fail 96 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!U13` =WEEKNUM($A13) → Excel "#VALUE!", Gnumeric 1 |
| BIN2OCT | 91 | error 91 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!G25` =BIN2OCT(A25) → Excel "140", Gnumeric null |
| CEILING | 89 | fail 88, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG43` =CEILING($A43,AG$2) → Excel "#VALUE!", Gnumeric 0 |
| FLOOR | 87 | fail 86, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG33` =FLOOR($A33,$AD$47:$AH$47) → Excel "#VALUE!", Gnumeric 45300 |
| HEX2DEC | 65 | error 65 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!J35` =HEX2DEC(E35) → Excel 16, Gnumeric null |
| DEC2BIN | 64 | error 64 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!C31` =DEC2BIN(A31) → Excel "#NUM!", Gnumeric null |
| BIN2DEC | 63 | error 63 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!C28` =BIN2DEC(A28) → Excel 14, Gnumeric null |
| DEC2OCT | 62 | error 62 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!G28` =DEC2OCT(A28) → Excel "7777777777", Gnumeric null |
| HEX2OCT | 62 | error 62 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!G27` =HEX2OCT(A27) → Excel "#NUM!", Gnumeric null |
| HEX2BIN | 62 | error 62 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!D31` =HEX2BIN(A31) → Excel "#NUM!", Gnumeric null |
| BIN2HEX | 60 | error 60 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!E20` =BIN2HEX(A20) → Excel "FFFFFFFF1B", Gnumeric null |
| OCT2BIN | 55 | error 55 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!L16` =OCT2BIN(776) → Excel "111111110", Gnumeric null |
| WEEKDAY | 54 | fail 54 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!R13` =WEEKDAY($A13,R$1) → Excel 1, Gnumeric 2 |
| OCT2HEX | 54 | error 54 | `ENGINEERING/NUMBER_SYSTEMS.xlsx!G21` =OCT2HEX(A21) → Excel "18000000", Gnumeric null |

## Errors

- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `ENGINEERING/NUMBER_SYSTEMS.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11208): CRITICAL **: 03:22:55.764: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `MATH_AND_TRIGONOMETRY/BASE_DECIMAL.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:11448): CRITICAL **: 03:23:26.233: go_conf_add_monitor: assertion 'node || key' fai
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
- `TEXT/REPT.xlsx`: Cannot create a string longer than 0x1fffffe8 characters
