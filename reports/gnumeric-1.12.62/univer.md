# univer vs gnumeric@1.12.62 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 254. Fails where the case's own cell matches but its spill does not: 95.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 156 | 16 | 3 | 0 | 81 | 89.1% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 98 | 9 | 0 | 0 | 0 | 91.6% |
| database | 182 | 85 | 97 | 0 | 0 | 0 | 46.7% |
| date | 378 | 192 | 151 | 0 | 0 | 35 | 56.0% |
| engineering | 572 | 345 | 155 | 0 | 42 | 30 | 63.7% |
| financial | 800 | 544 | 245 | 11 | 0 | 0 | 68.0% |
| information | 156 | 100 | 15 | 0 | 0 | 41 | 87.0% |
| logical | 107 | 80 | 16 | 10 | 0 | 1 | 75.5% |
| lookup | 625 | 369 | 107 | 142 | 0 | 7 | 59.7% |
| math | 788 | 533 | 104 | 97 | 11 | 43 | 71.5% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1043 | 189 | 30 | 0 | 41 | 82.6% |
| text | 549 | 386 | 115 | 19 | 0 | 29 | 74.2% |
| web | 8 | 4 | 2 | 2 | 0 | 0 | 50.0% |
| **total** | 5856 | 3935 | 1221 | 314 | 53 | 333 | 71.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 3935 | 1221 | 314 | 53 | 333 | 71.2% |
| **total** | 5856 | 3935 | 1221 | 314 | 53 | 333 | 71.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| AGGREGATE | 54 | unsupported 54 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", Gnumeric "#NAME?" |
| INDEX | 40 | fail 40 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =INDEX(A1,1) → Excel 0, Gnumeric "#REF!" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", Gnumeric "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", Gnumeric "#NAME?" |
| XLOOKUP | 29 | fail 29 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, Gnumeric "#VALUE!" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, Gnumeric 46335 |
| DATEDIF | 28 | fail 28 | `date/datedif.xlsx!Z1` =DATEDIF("2011/1/29","2021/3/31","Y") → Excel 10, Gnumeric "#VALUE!" |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", Gnumeric "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", Gnumeric "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", Gnumeric "#NAME?" |
| EXPAND | 17 | unsupported 17 | `lookup/expand.xlsx!Z1` =EXPAND(A1:B2,3,3) → Excel 2, Gnumeric "#NAME?" |
| GROWTH | 17 | fail 15, unsupported 2 | `statistical/growth.xlsx!Z1` =GROWTH(A1:F1,A3:F3,A5:B5,TRUE) → Excel 136608.33775488482, Gnumeric 136608.3377548852 |
| TAKE | 16 | unsupported 16 | `lookup/take.xlsx!Z1` =TAKE(A1:B2,1,1) → Excel 2, Gnumeric "#NAME?" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", Gnumeric "#NAME?" |
| SORT | 15 | fail 1, unsupported 14 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =SORT(A1:B5,2,-1) → Excel "Red", Gnumeric "#NAME?" |
| XIRR | 15 | fail 12, unsupported 3 | `financial/xirr.xlsx!Z1` =XIRR(A1:E1,A3:E3,0.1) → Excel "#NUM!", Gnumeric "#VALUE!" |
| DROP | 15 | unsupported 15 | `lookup/drop.xlsx!Z1` =DROP(A1:B2,1,1) → Excel 2, Gnumeric "#NAME?" |
| DEC2BIN | 14 | error 14 | `engineering/dec2bin.xlsx!Z1` =DEC2BIN(9,4) → Excel "1001", Gnumeric null |
| DEC2HEX | 14 | error 14 | `engineering/dec2hex.xlsx!Z1` =DEC2HEX(100,4) → Excel "0064", Gnumeric null |
| DEC2OCT | 14 | error 14 | `engineering/dec2oct.xlsx!Z1` =DEC2OCT(58,3) → Excel "072", Gnumeric null |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", Gnumeric "#NAME?" |
| SUBTOTAL | 14 | fail 14 | `math/subtotal.xlsx!Z1` =SUBTOTAL(6,A1) → Excel 0, Gnumeric 1 |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", Gnumeric "#NAME?" |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, Gnumeric "#NAME?" |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", Gnumeric "#NAME?" |

## Errors

- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14386): CRITICAL **: 03:28:43.869: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14389): CRITICAL **: 03:28:44.136: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `engineering/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14392): CRITICAL **: 03:28:44.406: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
- `math/base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:14988): CRITICAL **: 03:29:50.381: go_conf_add_monitor: assertion 'node || key' fai
