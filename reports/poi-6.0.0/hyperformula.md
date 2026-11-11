# hyperformula vs poi@6.0.0 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 4. Fails where both answered an error, but a different one: 64. Fails where the case's own cell matches but its spill does not: 1.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 2390 | 305 | 1230 | 0 | 241 | 60.9% |
| other | 576 | 539 | 26 | 1 | 0 | 10 | 95.2% |
| **total** | 4742 | 2929 | 331 | 1231 | 0 | 251 | 65.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 2929 | 331 | 1231 | 0 | 251 | 65.2% |
| **total** | 4742 | 2929 | 331 | 1231 | 0 | 251 | 65.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| DATEDIF | 100 | unsupported 100 | `functions/function-datedif.xlsx!A1` =DATEDIF("foo", 1, "Y") → Excel "#VALUE!", Apache POI "#NAME?" |
| XLOOKUP | 29 | fail 29 | `functions/function-xlookup.xlsx!A6` =XLOOKUP(1, B1:E1, B2:C3) → Excel "#VALUE!", Apache POI "#N/A" |
| NETWORKDAYS.INTL | 28 | unsupported 28 | `functions/function-networkdays.intl.xlsx!A3` =NETWORKDAYS.INTL(0, 1, "1010102") → Excel "#VALUE!", Apache POI "#NAME?" |
| HEX2OCT | 27 | unsupported 27 | `functions/function-hex2oct.xlsx!A1` =HEX2OCT("foo") → Excel "#NUM!", Apache POI "#NAME?" |
| ADDRESS | 25 | fail 25 | `functions/function-address.xlsx!A3` =ADDRESS(1,1,3,FALSE()) → Excel "R[1]C1", Apache POI "$A1" |
| HYPGEOM.DIST | 25 | unsupported 25 | `functions/function-hypgeom.dist.xlsx!A3` =HYPGEOM.DIST(1, 2, "baz", 4, TRUE()) → Excel "#VALUE!", Apache POI "#NAME?" |
| OCT2BIN | 24 | unsupported 24 | `functions/function-oct2bin.xlsx!A1` =OCT2BIN("foo") → Excel "#NUM!", Apache POI "#NAME?" |
| OCT2HEX | 24 | unsupported 24 | `functions/function-oct2hex.xlsx!A1` =OCT2HEX("foo") → Excel "#NUM!", Apache POI "#NAME?" |
| ISOWEEKNUM | 23 | unsupported 23 | `functions/function-isoweeknum.xlsx!A1` =ISOWEEKNUM("foo") → Excel "#VALUE!", Apache POI "#NAME?" |
| HEX2BIN | 22 | unsupported 22 | `functions/function-hex2bin.xlsx!A1` =HEX2BIN("foo") → Excel "#NUM!", Apache POI "#NAME?" |
| XOR | 21 | unsupported 21 | `functions/function-xor.xlsx!A3` =XOR(FALSE(), TRUE()) → Excel true, Apache POI "#NAME?" |
| BIN2HEX | 20 | unsupported 20 | `functions/function-bin2hex.xlsx!A1` =BIN2HEX("foo") → Excel "#NUM!", Apache POI "#NAME?" |
| BIN2OCT | 20 | unsupported 20 | `functions/function-bin2oct.xlsx!A1` =BIN2OCT("foo") → Excel "#NUM!", Apache POI "#NAME?" |
| TEXT | 20 | fail 20 | `functions/function-text.xlsx!A1` =TEXT(2, 42) → Excel "42", Apache POI "2.0" |
| BITLSHIFT | 19 | unsupported 19 | `functions/function-bitlshift.xlsx!A1` =BITLSHIFT(1, "foo") → Excel "#VALUE!", Apache POI "#NAME?" |
| BITRSHIFT | 19 | unsupported 19 | `functions/function-bitrshift.xlsx!A1` =BITRSHIFT(1, "foo") → Excel "#VALUE!", Apache POI "#NAME?" |
| BASE | 18 | unsupported 18 | `functions/function-base.xlsx!A1` =BASE("foo", 2, 3) → Excel "#VALUE!", Apache POI "#NAME?" |
| DEC2OCT | 18 | unsupported 18 | `functions/function-dec2oct.xlsx!A1` =DEC2OCT("foo") → Excel "#VALUE!", Apache POI "#NAME?" |
| BETA.DIST | 17 | unsupported 17 | `functions/function-beta.dist.xlsx!A4` =BETA.DIST(1, 2, 3, "abcd") → Excel "#VALUE!", Apache POI "#NAME?" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", Apache POI "#VALUE!" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", Apache POI "#VALUE!" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", Apache POI "#VALUE!" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", Apache POI "#VALUE!" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", Apache POI "#VALUE!" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", Apache POI "#VALUE!" |
