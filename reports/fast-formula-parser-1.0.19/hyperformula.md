# hyperformula vs fast-formula-parser@1.0.19 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 34. Fails where both answered an error, but a different one: 135. Fails where the case's own cell matches but its spill does not: 2.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 2492 | 531 | 864 | 38 | 241 | 63.5% |
| other | 576 | 442 | 121 | 3 | 0 | 10 | 78.1% |
| **total** | 4742 | 2934 | 652 | 867 | 38 | 251 | 65.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 2934 | 652 | 867 | 38 | 251 | 65.3% |
| **total** | 4742 | 2934 | 652 | 867 | 38 | 251 | 65.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 139 | unsupported 139 | `functions/function-xlookup.xlsx!A6` =XLOOKUP(1, B1:E1, B2:C3) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| MATCH | 66 | unsupported 66 | `functions/function-match.xlsx!A1` =MATCH(C2:C3, B1) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| TEXT | 63 | fail 63 | `other/date-and-time-arithmetic.xlsx!C1` =TEXT(A1-B1, "hh:mm") → Excel "01:23", fast-formula-parser "#VALUE!" |
| SUM | 39 | fail 1, error 38 | `functions/function-arrayformula.xlsx!A1` =SUM(A2:C2+A2:C2) → Excel 12, fast-formula-parser "#VALUE!" |
| SUBSTITUTE | 38 | unsupported 38 | `functions/function-substitute.xlsx!A1` =SUBSTITUTE("foo", "f", "bb") → Excel "bboo", fast-formula-parser "#NAME?" |
| DATEDIF | 37 | fail 36, unsupported 1 | `functions/function-datedif.xlsx!A1` =DATEDIF(1, 2, "abcd") → Excel "#NUM!", fast-formula-parser "#NAME?" |
| VLOOKUP | 28 | fail 28 | `functions/function-vlookup.xlsx!A1` =VLOOKUP(D1:E1, A2:B3, 2, TRUE()) → Excel "#VALUE!", fast-formula-parser "#N/A" |
| LARGE | 25 | unsupported 25 | `functions/function-large.xlsx!A1` =LARGE(1, "baz") → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| SMALL | 25 | unsupported 25 | `functions/function-small.xlsx!A1` =SMALL(1, "baz") → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| MAXIFS | 24 | unsupported 24 | `functions/function-maxifs.xlsx!A1` =MAXIFS(B1:B2, C1:C2, "><foo") → Excel 0, fast-formula-parser "#NAME?" |
| MINIFS | 24 | unsupported 24 | `functions/function-minifs.xlsx!A1` =MINIFS(B1:B2, C1:C2, "><foo") → Excel 0, fast-formula-parser "#NAME?" |
| SUBTOTAL | 24 | unsupported 24 | `functions/function-subtotal.xlsx!B1` =SUBTOTAL(101, A2:A4, A5) → Excel 3.5, fast-formula-parser "#NAME?" |
| WEEKDAY | 22 | fail 5, unsupported 17 | `functions/function-weekday.xlsx!A1` =WEEKDAY("foo", 1) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| NETWORKDAYS.INTL | 21 | fail 21 | `functions/function-networkdays.intl.xlsx!A4` =NETWORKDAYS.INTL(0, 1, -1) → Excel "#NUM!", fast-formula-parser "#ERROR!" |
| SUMIF | 18 | fail 18 | `functions/function-sumif.xlsx!A3` =SUMIF(B1:B2, {"a","b"}, A1:A2) → Excel 2, fast-formula-parser 2 |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| WORKDAY.INTL | 17 | fail 17 | `functions/function-workday.intl.xlsx!A1` =WORKDAY.INTL(0, 1, TRUE()) → Excel 2, fast-formula-parser "#VALUE!" |
| DOLLARDE | 16 | unsupported 16 | `functions/function-dollarde.xlsx!A1` =DOLLARDE(1,0) → Excel "#DIV/0!", fast-formula-parser "#NAME?" |
| DOLLARFR | 16 | unsupported 16 | `functions/function-dollarfr.xlsx!A1` =DOLLARFR(1,0) → Excel "#DIV/0!", fast-formula-parser "#NAME?" |
| TDIST | 16 | unsupported 16 | `functions/function-tdist.xlsx!A1` =TDIST("foo", 2, 3) → Excel "#VALUE!", fast-formula-parser "#NAME?" |

## Errors

- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
- `functions/function-sum.xlsx`: timeout after 60000 ms
