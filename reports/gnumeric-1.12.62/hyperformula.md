# hyperformula vs gnumeric@1.12.62 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 113. Fails where the case's own cell matches but its spill does not: 4.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3260 | 485 | 106 | 74 | 241 | 83.1% |
| other | 576 | 550 | 16 | 0 | 0 | 10 | 97.2% |
| **total** | 4742 | 3810 | 501 | 106 | 74 | 251 | 84.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 3810 | 501 | 106 | 74 | 251 | 84.8% |
| **total** | 4742 | 3810 | 501 | 106 | 74 | 251 | 84.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 44 | fail 44 | `functions/function-xlookup.xlsx!A6` =XLOOKUP(1, B1:E1, B2:C3) → Excel "#VALUE!", Gnumeric "#N/A" |
| CEILING.MATH | 24 | unsupported 24 | `functions/function-ceiling.math.xlsx!A1` =CEILING.MATH("foo") → Excel "#VALUE!", Gnumeric "#NAME?" |
| FLOOR.MATH | 24 | unsupported 24 | `functions/function-floor.math.xlsx!A1` =FLOOR.MATH("foo") → Excel "#VALUE!", Gnumeric "#NAME?" |
| DEC2BIN | 20 | error 20 | `functions/function-dec2bin.xlsx!A1` =DEC2BIN("foo") → Excel "#VALUE!", Gnumeric null |
| BASE | 18 | error 18 | `functions/function-base.xlsx!A1` =BASE("foo", 2, 3) → Excel "#VALUE!", Gnumeric null |
| DEC2HEX | 18 | error 18 | `functions/function-dec2hex.xlsx!A1` =DEC2HEX("foo") → Excel "#VALUE!", Gnumeric null |
| DEC2OCT | 18 | error 18 | `functions/function-dec2oct.xlsx!A1` =DEC2OCT("foo") → Excel "#VALUE!", Gnumeric null |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| VLOOKUP | 17 | fail 17 | `functions/function-vlookup.xlsx!A1` =VLOOKUP(D1:E1, A2:B3, 2, TRUE()) → Excel "#VALUE!", Gnumeric "#N/A" |
| CEILING.PRECISE | 14 | unsupported 14 | `functions/function-ceiling.precise.xlsx!A1` =CEILING.PRECISE(1, "bar") → Excel "#VALUE!", Gnumeric "#NAME?" |
| FLOOR.PRECISE | 14 | unsupported 14 | `functions/function-floor.precise.xlsx!A1` =FLOOR.PRECISE(1, "bar") → Excel "#VALUE!", Gnumeric "#NAME?" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| NETWORKDAYS.INTL | 12 | fail 12 | `functions/function-networkdays.intl.xlsx!A3` =NETWORKDAYS.INTL(0, 1, "1010102") → Excel "#VALUE!", Gnumeric "#NUM!" |
| SUBTOTAL | 12 | fail 12 | `functions/function-subtotal.xlsx!B1` =SUBTOTAL(101, A2:A4, A5) → Excel 3.5, Gnumeric "#NUM!" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", Gnumeric "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", Gnumeric "#NAME?" |

## Errors

- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-base.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10060): CRITICAL **: 03:20:56.297: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10246): CRITICAL **: 03:21:16.442: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10246): CRITICAL **: 03:21:16.443: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10249): CRITICAL **: 03:21:16.708: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
- `functions/function-dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:10252): CRITICAL **: 03:21:16.974: go_conf_add_monitor: assertion 'node || key' fai
